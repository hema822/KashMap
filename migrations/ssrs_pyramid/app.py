"""
KashMap SSRS to Pyramid Analytics Migration Module
==================================================
Developer-focused BI migration tool for Microsoft SQL Server Reporting Services (SSRS)
and Power BI Paginated Report definition files (.rdl, .xml, .zip) into Pyramid Analytics:
  - Modeler Data Flow (ETL, source tables, multi-CTE decomposition, joins, calculated columns)
  - Modeler Semantic Model (DAX semantic imports, tables, attributes, hierarchies, PQL measures)
  - Discover (grids, matrix tables, charts, scoped calculations, conditional formatting)
  - Present (dashboards, publications, page-break bursting, slicers, compound vouchers)

Built and validated against real production SSRS RDL datasets:
  1. Costco Cigars.rdl (Single view paginated extract with totals)
  2. Costco Rogue.rdl (Single view paginated extract with totals)
  3. Inventory Movement Report.rdl (8-CTE complex in-query ETL pipeline)
  4. NS vs DSI Inventory.rdl (Power BI DAX dataset with 18 design variables & DAX measures)
  5. International Sales 2.rdl (Multi-dataset parameter query model with group subtotals & margin PQL)
  6. In & Out Report.rdl (Compound purchase voucher with HTML text runs, rectangles, alerts, & hourly prompts)
  7. In Out Report.rdl (Email version with alert triggers and col-span banner cells)
"""

import os
import re
import io
import json
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple, Set, Literal, Union
from collections import defaultdict

import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pydantic import BaseModel, ConfigDict, Field, AliasChoices, field_validator, ValidationError

# ==============================================================================
# 1. XML & NAMESPACE-AGNOSTIC HELPER FUNCTIONS
# ==============================================================================

def _local(tag: str) -> str:
    """Strip namespace URI from ElementTree tag, e.g. '{http://...}Report' -> 'Report'."""
    return tag.split("}", 1)[-1] if "}" in tag else tag


def _iter_all(root: ET.Element, tagname: str) -> List[ET.Element]:
    """Find all descendant elements matching a given local tag name regardless of namespace."""
    return [el for el in root.iter() if _local(el.tag) == tagname]


def _find_child(el: Optional[ET.Element], tagname: str) -> Optional[ET.Element]:
    """Find immediate child element matching tagname, ignoring namespace."""
    if el is None:
        return None
    for child in el:
        if _local(child.tag) == tagname:
            return child
    return None


def _find_children(el: Optional[ET.Element], tagname: str) -> List[ET.Element]:
    """Find all immediate children matching tagname, ignoring namespace."""
    if el is None:
        return []
    return [c for c in el if _local(c.tag) == tagname]


def _text_of(el: Optional[ET.Element], default: str = "") -> str:
    """Safely return stripped text content of element."""
    return (el.text or "").strip() if el is not None and el.text else default


# ==============================================================================
# 2. QUERY DECOMPOSITION (SQL, MULTI-CTE & POWER BI DAX)
# ==============================================================================

_SQL_CLEAN_RE = re.compile(r"--[^\n]*|\/\*[\s\S]*?\*\/", re.MULTILINE)
_TABLE_REF_RE = re.compile(
    r"\b(?:FROM|JOIN)\s+(?:\[?([\w@#]+)\]?\.)?(?:\[?([\w@#]+)\]?\.)?\[?([\w@#]+)\]?(?:\s+(?:AS\s+)?([A-Za-z_][A-Za-z0-9_]*))?",
    re.IGNORECASE
)
_JOIN_FULL_RE = re.compile(
    r"\b(LEFT\s+(?:OUTER\s+)?JOIN|RIGHT\s+(?:OUTER\s+)?JOIN|FULL\s+(?:OUTER\s+)?JOIN|INNER\s+JOIN|CROSS\s+JOIN|JOIN)\s+(?:\[?([\w@#]+)\]?\.)?(?:\[?([\w@#]+)\]?\.)?\[?([\w@#]+)\]?(?:\s+(?:AS\s+)?([A-Za-z_][A-Za-z0-9_]*))?\s+ON\s+([\s\S]+?)(?=\b(?:LEFT|RIGHT|FULL|INNER|CROSS|JOIN|WHERE|GROUP|HAVING|ORDER|UNION|\))\b|$)",
    re.IGNORECASE
)
_PARAM_REF_RE = re.compile(r"@([A-Za-z0-9_]+)")
_WHERE_CLAUSE_RE = re.compile(r"\bWHERE\s+([\s\S]+?)(?=\b(?:GROUP\s+BY|HAVING|ORDER\s+BY|UNION|\))\b|$)", re.IGNORECASE)


def extract_all_ctes(clean_sql: str) -> List[str]:
    """Extract all chained CTE names from a WITH clause (e.g. in Inventory Movement Report)."""
    ctes = []
    with_match = re.search(r"\bWITH\s+([A-Za-z0-9_]+)\s+AS\s*\(", clean_sql, re.IGNORECASE)
    if with_match:
        ctes.append(with_match.group(1).strip())
        rest = clean_sql[with_match.end():]
        sub_ctes = re.findall(r"\)\s*,\s*([A-Za-z0-9_]+)\s+AS\s*\(", rest, re.IGNORECASE)
        for sc in sub_ctes:
            if sc.strip() not in ctes:
                ctes.append(sc.strip())
    return ctes


def decompose_dax_query(dax_text: str) -> Dict[str, Any]:
    """
    Decompose a Power BI Dataset (DAX) query (e.g. from NS vs DSI Inventory.rdl)
    into tables, columns, DAX measures, and model relationships.
    """
    tbl_cols = re.findall(r"'([^']+)'\[([^\]]+)\]", dax_text)
    t_map = defaultdict(set)
    for tbl, col in tbl_cols:
        t_map[tbl].add(col)

    tables = []
    for tbl_name, col_set in t_map.items():
        tables.append({
            "table_name": tbl_name,
            "alias": "",
            "is_cte": False,
            "columns": sorted(list(col_set))
        })

    # Extract DAX Measures: "MeasureName", CALCULATE(SUM('Table'[Col]))
    dax_measures = []
    meas_matches = re.findall(r'"([A-Za-z0-9_]+)"\s*,\s*(CALCULATE\s*\([\s\S]+?\)|SUM\s*\([\s\S]+?\))', dax_text, re.IGNORECASE)
    for mname, mexpr in meas_matches:
        dax_measures.append({
            "measure_name": mname,
            "dax_expression": mexpr.strip()
        })

    # Relationships between extracted tables
    joins = []
    t_names = list(t_map.keys())
    if len(t_names) > 1:
        for i in range(len(t_names) - 1):
            joins.append({
                "join_type": "DAX Semantic Model Relationship",
                "target_table": t_names[i+1],
                "alias": "",
                "on_condition": f"Power BI Model Relationship ({t_names[i]} <-> {t_names[i+1]})",
            })

    return {
        "tables": tables,
        "joins": joins,
        "ctes": [],
        "where_clause": "Evaluated in Power BI Tabular Engine",
        "query_parameters": [],
        "is_stored_proc": False,
        "is_dax": True,
        "dax_measures": dax_measures
    }


def decompose_query(sql_text: str, data_provider: str = "") -> Dict[str, Any]:
    """
    Decompose an SSRS CommandText SQL query or DAX query into structured components:
    CTEs, source tables, joins, where clauses, DAX measures, and query parameters.
    """
    if not sql_text or not sql_text.strip():
        return {
            "tables": [],
            "joins": [],
            "ctes": [],
            "where_clause": "",
            "query_parameters": [],
            "is_stored_proc": False,
            "is_dax": False,
            "dax_measures": []
        }

    clean_text = _SQL_CLEAN_RE.sub(" ", sql_text).strip()

    # Check for DAX query (PBIDATASET provider or DAX keywords)
    if (data_provider and "PBIDATASET" in data_provider.upper()) or \
       clean_text.upper().startswith(("EVALUATE", "DEFINE VAR")) or \
       "SUMMARIZECOLUMNS" in clean_text.upper():
        return decompose_dax_query(clean_text)

    # Check if this is a bare Stored Procedure call
    if not re.search(r"\b(SELECT|WITH)\b", clean_text, re.IGNORECASE) and re.match(r"^[A-Za-z0-9_\[\].]+", clean_text):
        return {
            "tables": [{"table_name": clean_text, "alias": "", "is_cte": False}],
            "joins": [],
            "ctes": [],
            "where_clause": "",
            "query_parameters": list(set(_PARAM_REF_RE.findall(clean_text))),
            "is_stored_proc": True,
            "is_dax": False,
            "dax_measures": []
        }

    # Extract all CTEs
    ctes = extract_all_ctes(clean_text)

    # Extract Source Tables & Views
    tables = []
    seen_tables = set()
    for match in _TABLE_REF_RE.finditer(clean_text):
        s1, s2, tbl, alias = match.groups()
        parts = [p for p in (s1, s2, tbl) if p]
        full_table = ".".join(parts) if parts else tbl
        alias_clean = alias.strip() if alias and alias.upper() not in {"ON", "WHERE", "JOIN", "LEFT", "RIGHT", "INNER", "CROSS", "GROUP", "ORDER"} else ""
        
        if full_table.upper() in {"SELECT", "FROM", "WHERE", "GROUP", "ORDER", "JOIN", "ON", "AS", "BY"}:
            continue
        if full_table and full_table not in seen_tables:
            seen_tables.add(full_table)
            tables.append({
                "table_name": full_table,
                "alias": alias_clean,
                "is_cte": full_table in ctes
            })

    # Extract Joins
    joins = []
    for match in _JOIN_FULL_RE.finditer(clean_text):
        j_type, s1, s2, tbl, alias, on_condition = match.groups()
        parts = [p for p in (s1, s2, tbl) if p]
        target_table = ".".join(parts) if parts else tbl
        j_type_clean = re.sub(r"\s+", " ", j_type.strip().upper())
        cond_clean = re.sub(r"\s+", " ", (on_condition or "").strip())
        joins.append({
            "join_type": j_type_clean,
            "target_table": target_table,
            "alias": alias or "",
            "on_condition": cond_clean
        })

    # Extract Where Clause
    where_match = _WHERE_CLAUSE_RE.search(clean_text)
    where_clause = re.sub(r"\s+", " ", where_match.group(1).strip()) if where_match else ""

    # Extract Query Parameters (@Param)
    params = sorted(list(set(_PARAM_REF_RE.findall(clean_text))))

    return {
        "tables": tables,
        "joins": joins,
        "ctes": ctes,
        "where_clause": where_clause,
        "query_parameters": params,
        "is_stored_proc": False,
        "is_dax": False,
        "dax_measures": []
    }


# ==============================================================================
# 3. SSRS EXPRESSION PARSER & PYRAMID PQL TRANSLATOR
# ==============================================================================

_FIELD_REF_RE = re.compile(r"Fields!([A-Za-z0-9_]+)(?:\.Value)?", re.IGNORECASE)
_PARAM_EXPR_RE = re.compile(r"Parameters!([A-Za-z0-9_]+)(?:\.Value)?", re.IGNORECASE)
_VAR_EXPR_RE = re.compile(r"Variables!([A-Za-z0-9_]+)(?:\.Value)?", re.IGNORECASE)
_SCOPED_AGG_RE = re.compile(r"\b(Sum|Avg|Count|CountDistinct|Min|Max|First|Last)\s*\(\s*Fields!([A-Za-z0-9_]+)(?:\.Value)?\s*,\s*\"?([A-Za-z0-9_]+)\"?\s*\)", re.IGNORECASE)
_AGG_FUNC_RE = re.compile(r"\b(Sum|Avg|Count|CountDistinct|Min|Max|First|Last|StDev|Var)\s*\(", re.IGNORECASE)
_IIF_RE = re.compile(r"\bIIF\s*\(", re.IGNORECASE)
_SWITCH_RE = re.compile(r"\bSwitch\s*\(", re.IGNORECASE)


def parse_and_translate_expression(expr_text: str, dataset_name: str = "") -> Dict[str, Any]:
    """
    Parse an SSRS expression, classify its business intent, and translate it into Pyramid PQL.
    Handles scoped aggregates (e.g. Sum(..., "Posting_Period")), HTML strings, date conversions,
    report variables, alternating rows, and conditional alert triggers.
    """
    raw_expr = (expr_text or "").strip()
    clean_expr = raw_expr[1:].strip() if raw_expr.startswith("=") else raw_expr

    field_refs = _FIELD_REF_RE.findall(clean_expr)
    param_refs = _PARAM_EXPR_RE.findall(clean_expr)
    var_refs = _VAR_EXPR_RE.findall(clean_expr)

    if not raw_expr.startswith("="):
        return {
            "raw_expression": raw_expr,
            "clean_expression": clean_expr,
            "expression_type": "Static Text",
            "field_references": [],
            "parameter_references": [],
            "variable_references": [],
            "pql_translation": f"'{raw_expr}'",
            "proposed_layer": "Pyramid Present Label",
            "triage_level": "automatic",
            "confidence": 1.0,
            "review_notes": "Static text literal."
        }

    # Alternating row background shading
    if "RowNumber" in clean_expr and "Mod 2" in clean_expr:
        return {
            "raw_expression": raw_expr,
            "clean_expression": clean_expr,
            "expression_type": "Grid Alternating Row Shading",
            "field_references": [],
            "parameter_references": [],
            "variable_references": var_refs,
            "pql_translation": "/* Configure in Discover Grid Formatting: Alternating Rows */",
            "proposed_layer": "Pyramid Discover Grid Formatting",
            "triage_level": "automatic",
            "confidence": 1.0,
            "review_notes": "Standard grid row shading; configured via Discover Theme / Row Striping."
        }

    # Report Variables reference (Theme/Style)
    if var_refs and not field_refs and not param_refs:
        return {
            "raw_expression": raw_expr,
            "clean_expression": clean_expr,
            "expression_type": "Design System Variable",
            "field_references": [],
            "parameter_references": [],
            "variable_references": var_refs,
            "pql_translation": f"[variables].[{var_refs[0]}]",
            "proposed_layer": "Pyramid Present Theme / Design Palette",
            "triage_level": "automatic",
            "confidence": 0.95,
            "review_notes": f"Maps to corporate Theme Variable '{var_refs[0]}'."
        }

    # Conditional Business Alert (e.g. HEADER_ERROR = 'FAILED TO ENTER NETSUITE')
    if "HEADER_ERROR" in clean_expr or "FAILED TO ENTER" in clean_expr:
        return {
            "raw_expression": raw_expr,
            "clean_expression": clean_expr,
            "expression_type": "Critical Business Rule Alert",
            "field_references": field_refs,
            "parameter_references": param_refs,
            "variable_references": var_refs,
            "pql_translation": "Iif([HEADER_ERROR] = 'FAILED TO ENTER NETSUITE', 1, 0)",
            "proposed_layer": "Pyramid Discover Conditional Formatting Alert",
            "triage_level": "recommended",
            "confidence": 0.90,
            "review_notes": "Configure as Discover Conditional Formatting rule or KPI alert banner."
        }

    # Scoped Aggregations: Sum(Fields!Dollars.Value, "Posting_Period")
    scoped_match = _SCOPED_AGG_RE.search(clean_expr)
    if scoped_match:
        func_name, fld, scope = scoped_match.groups()
        pql_trans = clean_expr
        pql_trans = _SCOPED_AGG_RE.sub(rf"{func_name.capitalize()}([measures].[{fld}], [dimension].[{scope}])", pql_trans)
        for f in set(field_refs):
            pql_trans = re.sub(rf"Fields!{re.escape(f)}(\.Value)?", f"[measures].[{f}]", pql_trans, flags=re.IGNORECASE)
        pql_trans = re.sub(r"\bIIF\b", "Iif", pql_trans, flags=re.IGNORECASE)
        pql_trans = re.sub(r"\bNothing\b", "Null", pql_trans, flags=re.IGNORECASE)
        return {
            "raw_expression": raw_expr,
            "clean_expression": clean_expr,
            "expression_type": "Scoped Aggregation (Group Subtotal)",
            "field_references": field_refs,
            "parameter_references": param_refs,
            "variable_references": var_refs,
            "pql_translation": pql_trans,
            "proposed_layer": "Pyramid Semantic Model Scoped Measure",
            "triage_level": "recommended",
            "confidence": 0.90,
            "review_notes": f"Aggregated at the '{scope}' hierarchy level."
        }

    # Standard Aggregate Expression
    if _AGG_FUNC_RE.search(clean_expr):
        expr_type = "Aggregation / Measure"
        proposed_layer = "Pyramid Semantic Model Measure"
        translated = clean_expr
        for f in set(field_refs):
            translated = re.sub(rf"Fields!{re.escape(f)}(\.Value)?", f"[measures].[{f}]", translated, flags=re.IGNORECASE)
        translated = re.sub(r"\bCountDistinct\b", "DistinctCount", translated, flags=re.IGNORECASE)
        translated = re.sub(r"\bFirst\b", "First", translated, flags=re.IGNORECASE)
        translated = re.sub(r"\bIIF\b", "Iif", translated, flags=re.IGNORECASE)
        translated = re.sub(r"\bNothing\b", "Null", translated, flags=re.IGNORECASE)
        return {
            "raw_expression": raw_expr,
            "clean_expression": clean_expr,
            "expression_type": expr_type,
            "field_references": field_refs,
            "parameter_references": param_refs,
            "variable_references": var_refs,
            "pql_translation": translated,
            "proposed_layer": proposed_layer,
            "triage_level": "automatic",
            "confidence": 0.95,
            "review_notes": "Standard aggregation translated to Pyramid PQL Measure."
        }

    # Conditional Logic (IIF / Switch)
    if _IIF_RE.search(clean_expr) or _SWITCH_RE.search(clean_expr):
        translated = re.sub(r"\bIIF\b", "Iif", clean_expr, flags=re.IGNORECASE)
        translated = re.sub(r"\bNothing\b", "Null", translated, flags=re.IGNORECASE)
        for f in set(field_refs):
            translated = re.sub(rf"Fields!{re.escape(f)}(\.Value)?", f"[{f}]", translated, flags=re.IGNORECASE)
        for p in set(param_refs):
            translated = re.sub(rf"Parameters!{re.escape(p)}(\.Value)?", f"[parameters].[{p}]", translated, flags=re.IGNORECASE)
        return {
            "raw_expression": raw_expr,
            "clean_expression": clean_expr,
            "expression_type": "Conditional Logic",
            "field_references": field_refs,
            "parameter_references": param_refs,
            "variable_references": var_refs,
            "pql_translation": translated,
            "proposed_layer": "Pyramid Modeler Calculated Column or PQL Measure",
            "triage_level": "recommended",
            "confidence": 0.88,
            "review_notes": "Translated to Pyramid Iif() syntax."
        }

    # HTML / Rich Text formatting
    if "<b>" in clean_expr or "&lt;b&gt;" in clean_expr or "<br" in clean_expr:
        clean_html = clean_expr.replace("&lt;", "<").replace("&gt;", ">")
        for f in set(field_refs):
            clean_html = re.sub(rf"Fields!{re.escape(f)}(\.Value)?", f"[{f}]", clean_html, flags=re.IGNORECASE)
        return {
            "raw_expression": raw_expr,
            "clean_expression": clean_expr,
            "expression_type": "HTML Rich Text Voucher",
            "field_references": field_refs,
            "parameter_references": param_refs,
            "variable_references": var_refs,
            "pql_translation": clean_html.replace("&", "+"),
            "proposed_layer": "Pyramid Present Rich Text Block",
            "triage_level": "recommended",
            "confidence": 0.90,
            "review_notes": "Multi-line HTML text voucher formatted for Present card widget."
        }

    # Parameter reference only
    if param_refs and not field_refs:
        return {
            "raw_expression": raw_expr,
            "clean_expression": clean_expr,
            "expression_type": "Parameter Dynamic Expression",
            "field_references": [],
            "parameter_references": param_refs,
            "variable_references": var_refs,
            "pql_translation": f"[parameters].[{param_refs[0]}]",
            "proposed_layer": "Pyramid Present Slicer / Discover Parameter",
            "triage_level": "automatic",
            "confidence": 0.95,
            "review_notes": f"Bound to parameter [{param_refs[0]}]."
        }

    # Date conversion / DateValue / Month
    if re.search(r"\b(DateValue|Month|Year|DateAdd|CDate)\b", clean_expr, re.IGNORECASE):
        translated = clean_expr
        for f in set(field_refs):
            translated = re.sub(rf"Fields!{re.escape(f)}(\.Value)?", f"[{f}]", translated, flags=re.IGNORECASE)
        return {
            "raw_expression": raw_expr,
            "clean_expression": clean_expr,
            "expression_type": "Date Hierarchy Expression",
            "field_references": field_refs,
            "parameter_references": param_refs,
            "variable_references": var_refs,
            "pql_translation": translated,
            "proposed_layer": "Pyramid Modeler Date Dimension / Attribute",
            "triage_level": "recommended",
            "confidence": 0.88,
            "review_notes": "Maps to Pyramid Time Intelligence / Date Dimension hierarchy."
        }

    # Direct Field reference
    if field_refs and len(clean_expr) < 50 and not re.search(r"[\+\-\*\/\&]", clean_expr):
        return {
            "raw_expression": raw_expr,
            "clean_expression": clean_expr,
            "expression_type": "Direct Attribute",
            "field_references": field_refs,
            "parameter_references": param_refs,
            "variable_references": var_refs,
            "pql_translation": f"[{field_refs[0]}]",
            "proposed_layer": "Pyramid Discover Grid Attribute",
            "triage_level": "automatic",
            "confidence": 1.0,
            "review_notes": "Direct attribute projection."
        }

    # General calculation
    translated = clean_expr
    for f in set(field_refs):
        translated = re.sub(rf"Fields!{re.escape(f)}(\.Value)?", f"[{f}]", translated, flags=re.IGNORECASE)
    translated = translated.replace("&", "+")

    return {
        "raw_expression": raw_expr,
        "clean_expression": clean_expr,
        "expression_type": "Calculated Expression",
        "field_references": field_refs,
        "parameter_references": param_refs,
        "variable_references": var_refs,
        "pql_translation": translated,
        "proposed_layer": "Pyramid Modeler Calculated Column",
        "triage_level": "recommended",
        "confidence": 0.85,
        "review_notes": "Custom calculated expression."
    }


# ==============================================================================
# 4. CORE SSRS RDL XML PARSER (AST EXTRACTION)
# ==============================================================================

def parse_ssrs_rdl(xml_content: str, filename: str = "Report.rdl") -> Dict[str, Any]:
    """
    Parse an SSRS RDL XML string into an Abstract Syntax Tree (AST) data dictionary.
    Namespace-agnostic: supports RDL 2005, 2008, 2010, 2016, and Power BI Paginated schemas.
    """
    try:
        if xml_content.startswith("\ufeff"):
            xml_content = xml_content[1:]
        root = ET.fromstring(xml_content.encode("utf-8"))
    except Exception as e:
        return {
            "error": f"Failed to parse XML: {str(e)}",
            "filename": filename,
            "report_name": Path(filename).stem,
            "data_sources": [],
            "data_sets": [],
            "parameters": [],
            "variables": [],
            "tablixes": [],
            "charts": [],
            "textboxes": [],
            "subreports": [],
            "drillthroughs": [],
            "conditional_styles": [],
            "expressions": [],
            "metadata": {}
        }

    report_name = Path(filename).stem

    # 4.1 Metadata
    metadata = {
        "report_id": _text_of(_find_child(root, "ReportID")),
        "default_font": _text_of(_find_child(root, "DefaultFontFamily"), "Segoe UI"),
        "unit_type": _text_of(_find_child(root, "ReportUnitType"), "Inch"),
        "auto_refresh": _text_of(_find_child(root, "AutoRefresh"), "0"),
        "created_by": "",
        "updated_by": "",
    }
    auth_meta = _find_child(root, "AuthoringMetadata")
    if auth_meta is not None:
        metadata["created_by"] = _text_of(_find_child(auth_meta, "CreatedBy"))
        metadata["updated_by"] = _text_of(_find_child(auth_meta, "UpdatedBy"))

    for page_el in _iter_all(root, "Page"):
        metadata["page_width"] = _text_of(_find_child(page_el, "PageWidth"))
        metadata["page_height"] = _text_of(_find_child(page_el, "PageHeight"))
        metadata["top_margin"] = _text_of(_find_child(page_el, "TopMargin"))
        metadata["bottom_margin"] = _text_of(_find_child(page_el, "BottomMargin"))
        metadata["left_margin"] = _text_of(_find_child(page_el, "LeftMargin"))
        metadata["right_margin"] = _text_of(_find_child(page_el, "RightMargin"))
        break

    # 4.2 Data Sources
    data_sources = []
    for ds_el in _iter_all(root, "DataSource"):
        ds_name = ds_el.attrib.get("Name", "DataSource1")
        conn_props = _find_child(ds_el, "ConnectionProperties")
        data_provider = "SQL Server / Relational"
        connect_string = ""
        security_type = ""
        if conn_props is not None:
            data_provider = _text_of(_find_child(conn_props, "DataProvider"), "SQL")
            connect_string = _text_of(_find_child(conn_props, "ConnectString"))
        sec_el = _find_child(ds_el, "SecurityType")
        if sec_el is not None:
            security_type = _text_of(sec_el)
        ds_ref = _text_of(_find_child(ds_el, "DataSourceReference"))

        data_sources.append({
            "name": ds_name,
            "data_provider": data_provider,
            "connect_string": connect_string,
            "security_type": security_type,
            "reference": ds_ref,
        })

    # 4.3 Report Variables (Corporate Design Palette)
    variables = []
    for var_el in _iter_all(root, "Variable"):
        v_name = var_el.attrib.get("Name", "")
        v_val = _text_of(_find_child(var_el, "Value"))
        variables.append({"name": v_name, "value": v_val})

    # 4.4 Report Parameters
    parameters = []
    for param_el in _iter_all(root, "ReportParameter"):
        p_name = param_el.attrib.get("Name", "Parameter")
        p_type = _text_of(_find_child(param_el, "DataType"), "String")
        p_prompt = _text_of(_find_child(param_el, "Prompt"), p_name)
        p_nullable = _text_of(_find_child(param_el, "Nullable"), "false").lower() == "true"
        p_allow_blank = _text_of(_find_child(param_el, "AllowBlank"), "false").lower() == "true"
        p_multi = _text_of(_find_child(param_el, "MultiValue"), "false").lower() == "true"

        default_vals = []
        def_el = _find_child(param_el, "DefaultValue")
        if def_el is not None:
            vals_el = _find_child(def_el, "Values")
            if vals_el is not None:
                for v in _find_children(vals_el, "Value"):
                    default_vals.append(_text_of(v))
            ds_ref = _find_child(def_el, "DataSetReference")
            if ds_ref is not None:
                default_vals.append(f"QueryDriven:{_text_of(_find_child(ds_ref, 'DataSetName'))}.{_text_of(_find_child(ds_ref, 'ValueField'))}")

        valid_type = "Static / None"
        valid_el = _find_child(param_el, "ValidValues")
        if valid_el is not None:
            if _find_child(valid_el, "DataSetReference") is not None:
                valid_type = "Query-Driven (DataSet)"
            elif _find_child(valid_el, "ParameterValues") is not None:
                valid_type = "Static List"

        parameters.append({
            "name": p_name,
            "data_type": p_type,
            "prompt": p_prompt,
            "nullable": p_nullable,
            "allow_blank": p_allow_blank,
            "multi_value": p_multi,
            "default_values": default_vals,
            "valid_values_type": valid_type,
        })

    # 4.5 DataSets & Query Decomposition
    data_sets = []
    all_raw_expressions = set()

    for ds_el in _iter_all(root, "DataSet"):
        ds_name = ds_el.attrib.get("Name", "DataSet1")
        query_el = _find_child(ds_el, "Query")
        ds_source_name = ""
        cmd_type = "Text"
        cmd_text = ""
        query_params = []

        if query_el is not None:
            ds_source_name = _text_of(_find_child(query_el, "DataSourceName"))
            cmd_type = _text_of(_find_child(query_el, "CommandType"), "Text")
            cmd_text = _text_of(_find_child(query_el, "CommandText"))
            
            qp_parent = _find_child(query_el, "QueryParameters")
            if qp_parent is not None:
                for qp in _find_children(qp_parent, "QueryParameter"):
                    qp_name = qp.attrib.get("Name", "")
                    qp_val = _text_of(_find_child(qp, "Value"))
                    query_params.append({"name": qp_name, "value": qp_val})
                    if qp_val.startswith("="):
                        all_raw_expressions.add(qp_val)

        # Match data provider for DAX vs SQL
        prov = next((ds["data_provider"] for ds in data_sources if ds["name"] == ds_source_name), "")
        query_decomp = decompose_query(cmd_text, prov)

        # Physical & Calculated Fields
        physical_fields = []
        calc_fields = []
        fields_parent = _find_child(ds_el, "Fields")
        if fields_parent is not None:
            for f_el in _find_children(fields_parent, "Field"):
                f_name = f_el.attrib.get("Name", "")
                df_el = _find_child(f_el, "DataField")
                val_el = _find_child(f_el, "Value")
                type_el = _find_child(f_el, "TypeName")
                data_type = _text_of(type_el, "System.String")
                if df_el is not None and _text_of(df_el):
                    physical_fields.append({
                        "name": f_name,
                        "data_field": _text_of(df_el),
                        "data_type": data_type
                    })
                elif val_el is not None and _text_of(val_el):
                    calc_expr = _text_of(val_el)
                    calc_fields.append({
                        "name": f_name,
                        "expression": calc_expr,
                        "data_type": data_type
                    })
                    all_raw_expressions.add(calc_expr)
                else:
                    physical_fields.append({
                        "name": f_name,
                        "data_field": f_name,
                        "data_type": data_type
                    })

        # DataSet Filters
        ds_filters = []
        filters_el = _find_child(ds_el, "Filters")
        if filters_el is not None:
            for f_item in _find_children(filters_el, "Filter"):
                f_expr = _text_of(_find_child(f_item, "FilterExpression"))
                f_op = _text_of(_find_child(f_item, "Operator"), "Equal")
                f_vals = []
                f_v_parent = _find_child(f_item, "FilterValues")
                if f_v_parent is not None:
                    for val_el in _find_children(f_v_parent, "FilterValue"):
                        f_vals.append(_text_of(val_el))
                ds_filters.append({
                    "expression": f_expr,
                    "operator": f_op,
                    "values": f_vals
                })
                if f_expr.startswith("="):
                    all_raw_expressions.add(f_expr)

        data_sets.append({
            "name": ds_name,
            "data_source_name": ds_source_name,
            "command_type": cmd_type,
            "command_text": cmd_text,
            "query_parameters": query_params,
            "physical_fields": physical_fields,
            "calculated_fields": calc_fields,
            "filters": ds_filters,
            "sql_decomposition": query_decomp,
        })

    # 4.6 Tablixes, Matrices, Rectangles & Page Breaks
    tablixes = []
    conditional_styles = []

    for tab_el in _iter_all(root, "Tablix"):
        t_name = tab_el.attrib.get("Name", "Tablix")
        t_dataset = _text_of(_find_child(tab_el, "DataSetName"))

        # Row Groups & Page Breaks
        row_groups = []
        has_page_break = False
        for r_hier in _iter_all(tab_el, "TablixRowHierarchy"):
            for member in _iter_all(r_hier, "TablixMember"):
                grp = _find_child(member, "Group")
                if grp is not None:
                    g_name = grp.attrib.get("Name", "")
                    g_exprs = [_text_of(ge) for ge in _iter_all(grp, "GroupExpression")]
                    for ge in g_exprs:
                        if ge.startswith("="):
                            all_raw_expressions.add(ge)
                    sort_exprs = []
                    sort_el = _find_child(member, "SortExpressions")
                    if sort_el is not None:
                        for se in _iter_all(sort_el, "SortExpression"):
                            s_text = _text_of(se)
                            sort_exprs.append(s_text)
                            if s_text.startswith("="):
                                all_raw_expressions.add(s_text)
                    
                    # Detect page break by group
                    pb_el = _find_child(grp, "PageBreak")
                    if pb_el is not None:
                        has_page_break = True

                    row_groups.append({
                        "group_name": g_name,
                        "expressions": g_exprs,
                        "sort_expressions": sort_exprs,
                        "page_break": pb_el is not None
                    })

        # Column Groups (Matrix detection)
        col_groups = []
        for c_hier in _iter_all(tab_el, "TablixColumnHierarchy"):
            for member in _iter_all(c_hier, "TablixMember"):
                grp = _find_child(member, "Group")
                if grp is not None:
                    cg_name = grp.attrib.get("Name", "")
                    cg_exprs = [_text_of(ge) for ge in _iter_all(grp, "GroupExpression")]
                    col_groups.append({
                        "group_name": cg_name,
                        "expressions": cg_exprs
                    })

        # Detect Nested Rectangles inside Tablix Cells (Compound Cards)
        rectangles = []
        for rect in _iter_all(tab_el, "Rectangle"):
            r_name = rect.attrib.get("Name", "Rectangle")
            r_tb = [_text_of(_find_child(t, "Value")) for t in _iter_all(rect, "Textbox")]
            rectangles.append({"name": r_name, "textboxes_count": len(r_tb)})

        # Cells, Expressions & Conditional Formatting
        cell_exprs = []
        for cell in _iter_all(tab_el, "TablixCell"):
            for v_el in _iter_all(cell, "Value"):
                v_text = _text_of(v_el)
                if v_text:
                    cell_exprs.append(v_text)
                    if v_text.startswith("="):
                        all_raw_expressions.add(v_text)

            # Check dynamic background & font color
            for bg_el in _iter_all(cell, "BackgroundColor"):
                bg_t = _text_of(bg_el)
                if bg_t.startswith("="):
                    all_raw_expressions.add(bg_t)
                    conditional_styles.append({"property": "BackgroundColor", "expression": bg_t})
            for col_el in _iter_all(cell, "Color"):
                col_t = _text_of(col_el)
                if col_t.startswith("="):
                    all_raw_expressions.add(col_t)
                    conditional_styles.append({"property": "Color", "expression": col_t})

        tablixes.append({
            "name": t_name,
            "dataset_name": t_dataset,
            "is_matrix": len(col_groups) > 0,
            "has_page_break": has_page_break,
            "row_groups": row_groups,
            "col_groups": col_groups,
            "nested_rectangles": rectangles,
            "cell_expressions_count": len(cell_exprs),
            "sample_expressions": list(set([e for e in cell_exprs if e.startswith("=")]))[:10]
        })

    # 4.7 Standalone Textboxes & HTML Runs
    textboxes = []
    for tb_el in _iter_all(root, "Textbox"):
        tb_name = tb_el.attrib.get("Name", "Textbox")
        tb_vals = []
        for v in _iter_all(tb_el, "Value"):
            v_t = _text_of(v)
            if v_t:
                tb_vals.append(v_t)
                if v_t.startswith("="):
                    all_raw_expressions.add(v_t)
        if tb_vals:
            textboxes.append({"name": tb_name, "value": " ".join(tb_vals)})

    # Global scan: harvest any expression from any Value, Color, BackgroundColor, Hidden tag
    for val_el in _iter_all(root, "Value"):
        v_str = _text_of(val_el)
        if v_str.startswith("="):
            all_raw_expressions.add(v_str)

    for hid_el in _iter_all(root, "Hidden"):
        h_str = _text_of(hid_el)
        if h_str.startswith("="):
            all_raw_expressions.add(h_str)
            conditional_styles.append({"property": "Visibility/Hidden", "expression": h_str})

    for col_el in _iter_all(root, "Color"):
        c_str = _text_of(col_el)
        if c_str.startswith("="):
            all_raw_expressions.add(c_str)

    for bg_el in _iter_all(root, "BackgroundColor"):
        b_str = _text_of(bg_el)
        if b_str.startswith("="):
            all_raw_expressions.add(b_str)

    # 4.8 Charts, Subreports & Drillthrough
    charts = []
    for chart_el in _iter_all(root, "Chart"):
        c_name = chart_el.attrib.get("Name", "Chart")
        c_dataset = _text_of(_find_child(chart_el, "DataSetName"))
        chart_type = "Column"
        for series in _iter_all(chart_el, "ChartSeries"):
            type_el = _find_child(series, "Type")
            if type_el is not None:
                chart_type = _text_of(type_el)
                break

        categories = [grp.attrib.get("Name", "") for cat_el in _iter_all(chart_el, "ChartCategoryHierarchy") for grp in _iter_all(cat_el, "Group")]
        series_groups = [grp.attrib.get("Name", "") for s_el in _iter_all(chart_el, "ChartSeriesHierarchy") for grp in _iter_all(s_el, "Group")]
        measures = [_text_of(v) for val_el in _iter_all(chart_el, "ChartDataPoint") for v in _iter_all(val_el, "Value") if _text_of(v)]

        charts.append({
            "name": c_name,
            "dataset_name": c_dataset,
            "chart_type": chart_type,
            "categories": categories,
            "series_groups": series_groups,
            "measures": list(set(measures))
        })

    subreports = []
    for sub_el in _iter_all(root, "Subreport"):
        s_name = sub_el.attrib.get("Name", "Subreport")
        report_ref = _text_of(_find_child(sub_el, "ReportName"))
        params_passed = [{"name": p.attrib.get("Name", ""), "value": _text_of(_find_child(p, "Value"))} for p in _iter_all(sub_el, "Parameter")]
        subreports.append({
            "name": s_name,
            "report_reference": report_ref,
            "parameters": params_passed
        })

    drillthroughs = []
    for act_el in _iter_all(root, "Action"):
        dt_el = _find_child(act_el, "Drillthrough")
        if dt_el is not None:
            tgt_rep = _text_of(_find_child(dt_el, "ReportName"))
            dt_params = [{"name": p.attrib.get("Name", ""), "value": _text_of(_find_child(p, "Value"))} for p in _iter_all(dt_el, "Parameter")]
            drillthroughs.append({
                "action_type": "Drillthrough Report",
                "target_report": tgt_rep,
                "parameters": dt_params
            })
        hl_el = _find_child(act_el, "Hyperlink")
        if hl_el is not None:
            hl_val = _text_of(hl_el)
            drillthroughs.append({
                "action_type": "External Hyperlink",
                "target_report": hl_val,
                "parameters": []
            })

    # 4.9 Parse and Translate All Expressions
    translated_expressions = [parse_and_translate_expression(e) for e in sorted(list(all_raw_expressions))]

    return {
        "filename": filename,
        "report_name": report_name,
        "metadata": metadata,
        "data_sources": data_sources,
        "variables": variables,
        "data_sets": data_sets,
        "parameters": parameters,
        "tablixes": tablixes,
        "charts": charts,
        "textboxes": textboxes,
        "subreports": subreports,
        "drillthroughs": drillthroughs,
        "conditional_styles": conditional_styles,
        "expressions": translated_expressions,
    }


# ==============================================================================
# 5. GENERIC MIGRATION REPRESENTATION BUILDER
# ==============================================================================

def build_generic_migration_representation(parsed: Dict[str, Any]) -> Dict[str, Any]:
    """
    Transform SSRS AST into KashMap's canonical generic migration representation.
    """
    report_name = parsed.get("report_name", "Report")
    
    # Sources
    sources = []
    for ds in parsed.get("data_sources", []):
        sources.append({
            "source_id": ds["name"],
            "source_type": ds["data_provider"],
            "connection_string": ds["connect_string"],
            "security": ds["security_type"],
            "role": "Power BI Semantic Gateway" if "PBIDATASET" in ds["data_provider"].upper() else "Primary Database Source"
        })

    # Physical Tables, Joins & Filters
    physical_tables = []
    joins = []
    filters = []

    for dset in parsed.get("data_sets", []):
        ds_name = dset["name"]
        decomp = dset.get("sql_decomposition", {})
        
        for t in decomp.get("tables", []):
            if isinstance(t, dict):
                t_name = t["table_name"]
                t_alias = t["alias"]
                is_cte = t["is_cte"]
                cols = t.get("columns", [])
            else:
                t_name = str(t)
                t_alias = ""
                is_cte = False
                cols = []

            physical_tables.append({
                "dataset_name": ds_name,
                "table_name": t_name,
                "alias": t_alias,
                "is_cte": is_cte,
                "columns": cols
            })

        for j in decomp.get("joins", []):
            joins.append({
                "dataset_name": ds_name,
                "target_table": j["target_table"],
                "join_type": j["join_type"],
                "alias": j.get("alias", ""),
                "condition": j.get("on_condition", ""),
                "pyramid_action": "Modeler Join Node / Relationship",
                "confidence": 0.92,
                "review_reason": "Standard SQL Join" if "OUTER" not in j["join_type"] else "Verify Outer Join cardinality in Pyramid Model"
            })

        if decomp.get("where_clause"):
            filters.append({
                "dataset_name": ds_name,
                "scope": "SQL / DAX Where Filter",
                "expression": decomp["where_clause"],
                "pyramid_target": "Modeler SQL Node / Data Flow Filter",
                "triage": "automatic"
            })

        for dsf in dset.get("filters", []):
            filters.append({
                "dataset_name": ds_name,
                "scope": "DataSet Filter",
                "expression": f"{dsf['expression']} {dsf['operator']} {', '.join(dsf['values'])}",
                "pyramid_target": "Pyramid Modeler Filter Node or Slicer",
                "triage": "recommended"
            })

    # Parameters
    parameters = []
    for p in parsed.get("parameters", []):
        parameters.append({
            "name": p["name"],
            "data_type": p["data_type"],
            "prompt": p["prompt"],
            "multi_value": p["multi_value"],
            "default_value": ", ".join(p["default_values"]) if p["default_values"] else "None",
            "valid_values_type": p["valid_values_type"],
            "pyramid_slicer": "Present / Discover Dropdown Slicer" if not p["multi_value"] else "Present Multi-Select Slicer",
            "triage": "automatic"
        })

    # Calculations
    calculations = []
    for expr in parsed.get("expressions", []):
        calculations.append({
            "expression": expr["raw_expression"],
            "type": expr["expression_type"],
            "pql_translation": expr["pql_translation"],
            "proposed_layer": expr["proposed_layer"],
            "triage_level": expr["triage_level"],
            "confidence": expr["confidence"],
            "review_notes": expr["review_notes"]
        })

    # Visuals
    visuals = []
    for t in parsed.get("tablixes", []):
        visuals.append({
            "name": t["name"],
            "type": "Matrix Grid" if t["is_matrix"] else "Tabular Grid",
            "dataset_name": t["dataset_name"],
            "row_groups_count": len(t["row_groups"]),
            "col_groups_count": len(t["col_groups"]),
            "has_page_break": t.get("has_page_break", False),
            "pyramid_discover_type": "Matrix Grid" if t["is_matrix"] else "Tabular Grid",
            "triage": "recommended" if t["is_matrix"] else "automatic"
        })

    for c in parsed.get("charts", []):
        visuals.append({
            "name": c["name"],
            "type": f"Chart ({c['chart_type']})",
            "dataset_name": c["dataset_name"],
            "row_groups_count": len(c["categories"]),
            "col_groups_count": len(c["series_groups"]),
            "has_page_break": False,
            "pyramid_discover_type": f"Discover {c['chart_type']} Chart",
            "triage": "automatic"
        })

    # Semantic Model (Pyramid Model Pro Representation)
    sem_tables = []
    seen_sem_tables = set()
    for src in sources:
        sem_tables.append({
            "table_name": f"Conn_{src['source_id']}",
            "role": "Connection / Gateway",
            "source_type": src["source_type"],
            "grain": "Server / Connection",
            "description": f"Pyramid Data Source connection ({src['source_type']})"
        })

    for tbl in physical_tables:
        t_name = tbl["table_name"]
        if t_name not in seen_sem_tables:
            seen_sem_tables.add(t_name)
            is_cte = tbl.get("is_cte", False)
            sem_tables.append({
                "table_name": t_name,
                "role": "Data Flow CTE Prep Node" if is_cte else "Model Table",
                "source_type": "CTE / Subquery" if is_cte else "Physical Table / View",
                "grain": "Calculated CTE Grain" if is_cte else "Transactional Grain",
                "description": f"Pyramid Data Model table representing {t_name}"
            })

    # Columns
    sem_columns = []
    seen_sem_cols = set()
    for dset in parsed.get("data_sets", []):
        ds_name = dset["name"]
        all_flds = dset.get("physical_fields", []) + dset.get("calculated_fields", []) + dset.get("fields", [])
        for fld in all_flds:
            f_name = fld.get("name", "")
            if not f_name:
                continue
            key = (ds_name, f_name)
            if key not in seen_sem_cols:
                seen_sem_cols.add(key)
                sem_columns.append({
                    "table_name": ds_name,
                    "column_name": f_name,
                    "data_type": fld.get("data_type", "") or ("Text" if any(w in f_name.lower() for w in ["name", "desc", "type", "code", "status"]) else ("Numeric" if any(w in f_name.lower() for w in ["amt", "qty", "total", "cost", "price", "val"]) else "Database Column")),
                    "role": "Calculated Column" if fld.get("is_calculated") or "calculated_fields" in fld else "Attribute / Dimension",
                    "origin": "DataSet Field (RDL)",
                    "formula": fld.get("value", "") or fld.get("data_field", "")
                })

    # Measures
    sem_measures = []
    seen_meas = set()
    for calc in calculations:
        if calc["type"] in {"Aggregation / Measure", "Scoped Aggregation (Group Subtotal)"}:
            m_name = re.sub(r"[^\w]", "_", calc["expression"].replace("=", "").replace("Fields!", "").replace(".Value", "")).strip("_")
            if m_name and m_name not in seen_meas:
                seen_meas.add(m_name)
                sem_measures.append({
                    "measure_name": m_name,
                    "source_field": calc["expression"],
                    "table_name": parsed.get("data_sets", [{}])[0].get("name", "Model") if parsed.get("data_sets") else "Model",
                    "aggregation": "Custom Formulate PQL" if "scoped" in calc["type"].lower() else "Sum / Aggregate",
                    "type": "Calculated Measure (PQL)",
                    "pql_formula": calc["pql_translation"],
                    "format_string": "$#,##0.00" if any(w in calc["expression"].lower() for w in ["dollar", "cost", "price", "amount", "sales"]) else "#,##0"
                })

    # Also extract DAX measures if present
    for dset in parsed.get("data_sets", []):
        dax_decomp = dset.get("sql_decomposition", {})
        for dax_m in dax_decomp.get("dax_measures", []):
            m_name = dax_m["measure_name"]
            if m_name not in seen_meas:
                seen_meas.add(m_name)
                sem_measures.append({
                    "measure_name": m_name,
                    "source_field": m_name,
                    "table_name": dset["name"],
                    "aggregation": "DAX Semantic Measure",
                    "type": "Power BI / DAX Measure",
                    "pql_formula": dax_m.get("expression", f"[{m_name}]"),
                    "format_string": "#,##0.00"
                })

    # Relationships
    sem_relationships = []
    for j in joins:
        sem_relationships.append({
            "from_table": j.get("dataset_name", "Source"),
            "to_table": j.get("target_table", "Target"),
            "from_column": "ON Condition Left Key" if j.get("condition") else "Manual Review / Source metadata required",
            "to_column": "ON Condition Right Key" if j.get("condition") else "Manual Review / Source metadata required",
            "condition": j.get("condition") or "Manual Review / Source metadata required",
            "implementation": "Pyramid Modeler Join Node",
            "alternative": "Semantic Model Relationship",
            "confidence": 92 if j.get("condition") else 60,
            "review_level": "automatic" if j.get("condition") and "OUTER" not in j.get("join_type", "") else "recommended",
            "reason": f"{j.get('join_type', 'INNER')} join detected in DataSet query" if j.get("condition") else "Source metadata required to resolve foreign keys"
        })

    # If multiple datasets or tables exist without explicit joins in RDL:
    if len(physical_tables) > 1 and not joins:
        sem_relationships.append({
            "from_table": physical_tables[0]["table_name"],
            "to_table": physical_tables[1]["table_name"],
            "from_column": "Manual Review / Source metadata required",
            "to_column": "Manual Review / Source metadata required",
            "condition": "Manual Review / Source metadata required",
            "implementation": "Pyramid Modeler Join Node",
            "alternative": "Semantic Model Relationship",
            "confidence": 50,
            "review_level": "manual_review",
            "reason": "Multiple source tables detected, but explicit join keys are not present in RDL query text. Source database metadata required."
        })

    # Hierarchies
    sem_hierarchies = []
    for t in parsed.get("tablixes", []):
        r_groups = [g.get("group_name") or g.get("name") for g in t.get("row_groups", []) if (g.get("group_name") or g.get("name"))]
        if len(r_groups) >= 2:
            sem_hierarchies.append({
                "name": f"{t['name']} Row Group Hierarchy",
                "table_name": t.get("dataset_name", "Model"),
                "levels": r_groups,
                "type": "Attribute Hierarchy",
                "description": f"Drill-down path: {' -> '.join(r_groups)}"
            })

    # Calendar / Time hierarchy if date or period columns present
    date_cols = [c["column_name"] for c in sem_columns if any(w in c["column_name"].lower() for w in ["date", "year", "month", "day", "period", "posting"])]
    if len(date_cols) >= 2:
        tbl_target = sem_columns[0]["table_name"] if sem_columns else "Model"
        sem_hierarchies.append({
            "name": "Date / Period Hierarchy",
            "table_name": tbl_target,
            "levels": date_cols[:4],
            "type": "Time Hierarchy",
            "description": f"Drill-down path: {' -> '.join(date_cols[:4])}"
        })

    semantic_model = {
        "tables": sem_tables,
        "columns": sem_columns,
        "measures": sem_measures,
        "relationships": sem_relationships,
        "hierarchies": sem_hierarchies,
    }

    return {
        "report_name": report_name,
        "sources": sources,
        "variables": parsed.get("variables", []),
        "physical_tables": physical_tables,
        "joins": joins,
        "filters": filters,
        "parameters": parameters,
        "calculations": calculations,
        "visuals": visuals,
        "subreports": parsed.get("subreports", []),
        "drillthroughs": parsed.get("drillthroughs", []),
        "conditional_styles": parsed.get("conditional_styles", []),
        "semantic_model": semantic_model
    }


# ==============================================================================
# 6. PYRAMID TARGET MAPPING & TRIAGE CLASSIFICATION
# ==============================================================================

def map_generic_to_pyramid(generic_logic: Dict[str, Any]) -> Dict[str, Any]:
    """
    Map generic logic to concrete Pyramid Analytics objects:
      - 8-Stage Structured Modeler Flow (Ingestion, Extraction, Prep, Integration, Semantic Model, Discover, Present, Validation)
      - Modeler Semantic Model (Tables, Columns, Relationships, Measures with PQL)
      - Discover (Visualizations, Matrix Grids, Charts, Conditional Formatting)
      - Present (Publications, Slicers, Bursting, Page Breaks, Vouchers)
    """
    report_name = generic_logic.get("report_name", "Report")
    modeler_steps = []
    modeler_plan_rows = []
    step_num = 1

    # Stage 1: Ingestion
    for src in generic_logic.get("sources", []):
        is_pbi = "PBIDATASET" in src["source_type"].upper()
        op_name = "Power BI Semantic Gateway" if is_pbi else "Relational Database Connection"
        modeler_steps.append({
            "step_number": step_num,
            "stage": "1. Ingestion",
            "step_name": f"Connect_{src['source_id']}",
            "pyramid_layer": "Modeler Data Flow",
            "source_node": src["source_type"],
            "operation": op_name,
            "target_node": f"Source_{src['source_id']}",
            "triage": "automatic",
            "confidence": 0.96,
            "guidelines": f"Configure connection string in Pyramid Modeler: {src['connection_string'] or 'Use shared Pyramid Server Connection'}."
        })
        modeler_plan_rows.append({
            "stage": "1. Ingestion",
            "category": "Source",
            "name": src["source_id"],
            "ssrs_logic": f"DataSource: {src['source_type']}",
            "pyramid_target": "Pyramid Data Source / Connection",
            "pyramid_action": f"Configure data source connection in Pyramid Modeler ({src['source_type']})",
            "review_level": "automatic",
            "confidence": 96,
            "reason": "Direct operational or semantic database connection",
        })
        step_num += 1

    # Stage 2: Table / Query Extraction
    seen_tables = set()
    for tbl in generic_logic.get("physical_tables", []):
        t_name = tbl["table_name"]
        if t_name not in seen_tables:
            seen_tables.add(t_name)
            is_cte = tbl.get("is_cte", False)
            modeler_steps.append({
                "step_number": step_num,
                "stage": "2. Table Extraction",
                "step_name": f"Extract_{t_name.replace('.', '_')}",
                "pyramid_layer": "Modeler Data Flow",
                "source_node": tbl.get("dataset_name", "DataSet"),
                "operation": "Modeler CTE Stage Node" if is_cte else "Table Select Node",
                "target_node": f"Prep_{t_name.replace('.', '_')}",
                "triage": "recommended" if is_cte else "automatic",
                "confidence": 0.92 if is_cte else 0.98,
                "guidelines": f"Import table {t_name} into Pyramid Modeler canvas."
            })
            modeler_plan_rows.append({
                "stage": "2. Table Extraction",
                "category": "Table Extraction",
                "name": t_name,
                "ssrs_logic": f"DataSet: {tbl.get('dataset_name', '')} (Table: {t_name})",
                "pyramid_target": "Modeler CTE Stage Node" if is_cte else "Table Select Node",
                "pyramid_action": f"Extract table '{t_name}' into Pyramid Modeler canvas",
                "review_level": "recommended" if is_cte else "automatic",
                "confidence": 92 if is_cte else 98,
                "reason": "Common Table Expression (CTE)" if is_cte else "Direct operational or warehouse table source",
            })
            step_num += 1

    # Stage 3: Data Preparation (CTEs, SQL Filters, Dataset Filters)
    has_ctes = any(tbl.get("is_cte") for tbl in generic_logic.get("physical_tables", []))
    filters = generic_logic.get("filters", [])
    if has_ctes or filters:
        for tbl in generic_logic.get("physical_tables", []):
            if tbl.get("is_cte"):
                modeler_steps.append({
                    "step_number": step_num,
                    "stage": "3. Data Preparation",
                    "step_name": f"Prep_CTE_{tbl['table_name'].replace('.', '_')}",
                    "pyramid_layer": "Modeler Data Flow",
                    "source_node": tbl.get("dataset_name", "DataSet"),
                    "operation": "Modeler CTE Stage Transform",
                    "target_node": f"Staged_{tbl['table_name'].replace('.', '_')}",
                    "triage": "recommended",
                    "confidence": 0.92,
                    "guidelines": f"Configure CTE transform stage node for {tbl['table_name']}."
                })
                modeler_plan_rows.append({
                    "stage": "3. Data Preparation",
                    "category": "Data Preparation",
                    "name": tbl["table_name"],
                    "ssrs_logic": f"WITH {tbl['table_name']} AS (...) CTE Expression",
                    "pyramid_target": "Pyramid Data Flow CTE / Prep Step",
                    "pyramid_action": f"Implement staging CTE transform node for '{tbl['table_name']}'",
                    "review_level": "recommended",
                    "confidence": 92,
                    "reason": "Multi-CTE query requires sequential Data Flow preparation stage",
                })
                step_num += 1

        for idx, filt in enumerate(filters[:6], start=1):
            modeler_steps.append({
                "step_number": step_num,
                "stage": "3. Data Preparation",
                "step_name": f"Filter_{idx}",
                "pyramid_layer": "Modeler Data Flow",
                "source_node": filt.get("dataset_name", "DataSet"),
                "operation": "Data Flow Row Filter",
                "target_node": "Filtered_Data",
                "triage": filt.get("triage", "automatic"),
                "confidence": 0.90,
                "guidelines": f"Apply filter: {filt.get('expression', '')[:80]}"
            })
            modeler_plan_rows.append({
                "stage": "3. Data Preparation",
                "category": "Data Preparation",
                "name": f"Filter_{idx}",
                "ssrs_logic": filt.get("expression", "")[:120],
                "pyramid_target": filt.get("pyramid_target", "Modeler Data Flow Filter"),
                "pyramid_action": "Apply row filter in Data Flow node",
                "review_level": filt.get("triage", "automatic"),
                "confidence": 90,
                "reason": filt.get("scope", "DataSet Filter"),
            })
            step_num += 1

    # Stage 4: Data Integration (Joins & Lookups)
    for j in generic_logic.get("joins", []):
        is_outer = "OUTER" in j.get("join_type", "").upper()
        modeler_steps.append({
            "step_number": step_num,
            "stage": "4. Data Integration",
            "step_name": f"Join_{j['target_table'].replace('.', '_')}",
            "pyramid_layer": "Modeler Data Flow / Model",
            "source_node": j["dataset_name"],
            "operation": f"Join Node ({j['join_type']})",
            "target_node": "Unified_Model_Table",
            "triage": "recommended" if is_outer else "automatic",
            "confidence": 0.92,
            "guidelines": f"Apply join condition in Modeler Join Node: {j['condition']}"
        })
        modeler_plan_rows.append({
            "stage": "4. Data Integration",
            "category": "Relationship",
            "name": f"{j['dataset_name']} -> {j['target_table']}",
            "ssrs_logic": j.get("condition") or f"{j.get('join_type', 'INNER')} JOIN",
            "pyramid_target": "Pyramid Modeler Join Node or Semantic Relationship",
            "pyramid_action": f"Configure join between tables on {j.get('condition') or 'Key'}",
            "review_level": "recommended" if is_outer else "automatic",
            "confidence": 92,
            "reason": f"{j.get('join_type', 'INNER')} join condition extracted from query",
        })
        step_num += 1

    # Stage 5: Semantic Model Target & Measures
    modeler_steps.append({
        "step_number": step_num,
        "stage": "5. Semantic Model",
        "step_name": f"Target_Model_{report_name}",
        "pyramid_layer": "Modeler Semantic Model",
        "source_node": "Unified_Model_Table",
        "operation": "Model Target Node (In-Memory or Relational Direct)",
        "target_node": f"Model_{report_name}",
        "triage": "automatic",
        "confidence": 0.95,
        "guidelines": "Process and deploy Semantic Model to Pyramid IMDB or direct query database."
    })
    modeler_plan_rows.append({
        "stage": "5. Semantic Model",
        "category": "Semantic Model",
        "name": f"Model_{report_name}",
        "ssrs_logic": "Semantic Model Target",
        "pyramid_target": "Pyramid Semantic Model (Model Pro)",
        "pyramid_action": f"Deploy Semantic Model '{report_name}' to Pyramid IMDB or direct query",
        "review_level": "automatic",
        "confidence": 95,
        "reason": "Target semantic reporting model",
    })
    step_num += 1

    # Semantic Model Elements (Measures & Attributes)
    semantic_measures = []
    semantic_attributes = []

    for calc in generic_logic.get("calculations", []):
        if calc["type"] in {"Aggregation / Measure", "Scoped Aggregation (Group Subtotal)"}:
            m_name = re.sub(r"[^\w]", "_", calc["expression"].replace("=", "").replace("Fields!", "").replace(".Value", "")).strip("_")
            semantic_measures.append({
                "measure_name": m_name,
                "expression": calc["expression"],
                "pql_formula": calc["pql_translation"],
                "format_string": "$#,##0.00" if any(w in calc["expression"].lower() for w in ["dollar", "cost", "price", "amount", "sales"]) else "#,##0",
                "triage": calc["triage_level"],
                "confidence": calc["confidence"]
            })
            modeler_plan_rows.append({
                "stage": "5. Semantic Model",
                "category": "Calculation / Measure",
                "name": m_name,
                "ssrs_logic": calc["expression"][:100],
                "pyramid_target": "Formulate Measure (PQL)",
                "pyramid_action": f"Create PQL Formulate Measure: {calc['pql_translation'][:80]}",
                "review_level": calc["triage_level"],
                "confidence": int(calc["confidence"] * 100) if calc["confidence"] <= 1 else calc["confidence"],
                "reason": calc["type"],
            })
        elif calc["type"] not in {"Static Text", "Direct Attribute", "Grid Alternating Row Shading"}:
            semantic_attributes.append({
                "column_name": re.sub(r"[^\w]", "_", calc["expression"][:35]).strip("_"),
                "expression": calc["expression"],
                "pql_formula": calc["pql_translation"],
                "layer": calc["proposed_layer"],
                "triage": calc["triage_level"],
                "confidence": calc["confidence"]
            })

    # Stage 6: Discover Visual Mapping
    discover_visuals = []
    for vis in generic_logic.get("visuals", []):
        discover_visuals.append({
            "name": vis["name"],
            "source_type": vis["type"],
            "pyramid_discover_app": "Pyramid Discover",
            "pyramid_visual": vis["pyramid_discover_type"],
            "pyramid_discover_type": vis["pyramid_discover_type"],
            "triage": vis["triage"],
            "recommendations": "Directly configure in Discover canvas using the generated Semantic Model."
        })
        modeler_steps.append({
            "step_number": step_num,
            "stage": "6. Discover",
            "step_name": f"Discover_{vis['name'].replace('.', '_')}",
            "pyramid_layer": "Pyramid Discover",
            "source_node": f"Model_{report_name}",
            "operation": f"Discover {vis['pyramid_discover_type']}",
            "target_node": f"Vis_{vis['name'].replace('.', '_')}",
            "triage": vis["triage"],
            "confidence": 0.95,
            "guidelines": f"Build Discover {vis['pyramid_discover_type']} report canvas."
        })
        modeler_plan_rows.append({
            "stage": "6. Discover",
            "category": "Visual",
            "name": vis["name"],
            "ssrs_logic": f"{vis['type']} ({vis['dataset_name']})",
            "pyramid_target": f"Pyramid Discover ({vis['pyramid_discover_type']})",
            "pyramid_action": f"Build Discover {vis['pyramid_discover_type']} report canvas",
            "review_level": vis["triage"],
            "confidence": 95,
            "reason": f"Direct translation of SSRS {vis['type']} to Pyramid Discover visual",
        })
        step_num += 1

    # Stage 7: Present Dashboard & Publication Items
    present_items = []
    for p in generic_logic.get("parameters", []):
        present_items.append({
            "component": f"Slicer: {p['name']}",
            "source_type": "ReportParameter",
            "pyramid_present_type": p["pyramid_slicer"],
            "prompt": p["prompt"],
            "triage": p["triage"],
            "notes": f"Bind slicer to attribute [{p['name']}] across all Discover visual queries on the Present dashboard."
        })
        modeler_steps.append({
            "step_number": step_num,
            "stage": "7. Present",
            "step_name": f"Slicer_{p['name']}",
            "pyramid_layer": "Pyramid Present",
            "source_node": p["name"],
            "operation": "Present Parameter Slicer",
            "target_node": "Dashboard_Slicers",
            "triage": p["triage"],
            "confidence": 0.95,
            "guidelines": f"Configure interactive {p['pyramid_slicer']} on Present dashboard."
        })
        modeler_plan_rows.append({
            "stage": "7. Present",
            "category": "Parameter / Slicer",
            "name": p["name"],
            "ssrs_logic": f"ReportParameter: {p['name']} ({p.get('data_type', 'Text')})",
            "pyramid_target": p["pyramid_slicer"],
            "pyramid_action": f"Create interactive slicer bound to parameter '{p['name']}'",
            "review_level": p["triage"],
            "confidence": 95,
            "reason": "User prompt input parameter converted to Pyramid Present Slicer",
        })
        step_num += 1

    for v in generic_logic.get("variables", []):
        present_items.append({
            "component": f"Design Variable: {v['name']}",
            "source_type": "ReportVariable",
            "pyramid_present_type": "Theme Color / Geometry Asset",
            "prompt": v["value"],
            "triage": "automatic",
            "notes": "Define under Pyramid Present Theme / Template Design Palette."
        })

    for vis in generic_logic.get("visuals", []):
        if vis.get("has_page_break"):
            present_items.append({
                "component": f"Publication Bursting: {vis['name']}",
                "source_type": "Group PageBreak",
                "pyramid_present_type": "Pyramid Present Publication (Page-by-Slice)",
                "prompt": "Burst by Group",
                "triage": "recommended",
                "notes": "Configure Publication Schedule bursting by the row hierarchy group attribute."
            })
            modeler_steps.append({
                "step_number": step_num,
                "stage": "7. Present",
                "step_name": f"Bursting_{vis['name'].replace('.', '_')}",
                "pyramid_layer": "Pyramid Publish",
                "source_node": vis["name"],
                "operation": "Publication Page-by-Slice Bursting",
                "target_node": "Scheduled_Publication",
                "triage": "recommended",
                "confidence": 0.90,
                "guidelines": "Configure Publication Schedule bursting by group."
            })
            modeler_plan_rows.append({
                "stage": "7. Present",
                "category": "Publication",
                "name": vis["name"],
                "ssrs_logic": "Group PageBreak",
                "pyramid_target": "Pyramid Present Publication (Page-by-Slice Bursting)",
                "pyramid_action": "Configure scheduled publication bursting by group",
                "review_level": "recommended",
                "confidence": 90,
                "reason": "SSRS page break mapped to Pyramid publication bursting",
            })
            step_num += 1

    for sub in generic_logic.get("subreports", []):
        present_items.append({
            "component": f"Subreport: {sub['name']}",
            "source_type": "SSRS Subreport",
            "pyramid_present_type": "Compound Present Slide / Publication",
            "prompt": f"Target: {sub['report_reference']}",
            "triage": "manual_review",
            "notes": "Convert child subreport into a distinct Discover report and combine on a composite Present slide."
        })
        modeler_steps.append({
            "step_number": step_num,
            "stage": "7. Present",
            "step_name": f"Subreport_{sub['name']}",
            "pyramid_layer": "Pyramid Present",
            "source_node": sub["name"],
            "operation": "Composite Slide Integration",
            "target_node": "Composite_Slide",
            "triage": "manual_review",
            "confidence": 0.75,
            "guidelines": f"Review subreport dataset parameters and combine into composite Present slide."
        })
        modeler_plan_rows.append({
            "stage": "7. Present",
            "category": "Subreport",
            "name": sub["name"],
            "ssrs_logic": f"Subreport reference: {sub.get('report_reference', '')}",
            "pyramid_target": "Composite Present Slide / Storyboard",
            "pyramid_action": "Combine child Discover visual into Present dashboard card",
            "review_level": "manual_review",
            "confidence": 75,
            "reason": "Subreport requires manual review of parameters and dataset linkage",
        })
        step_num += 1

    for dt in generic_logic.get("drillthroughs", []):
        present_items.append({
            "component": f"Action: {dt['action_type']}",
            "source_type": "SSRS Action",
            "pyramid_present_type": "Pyramid Jump-to-Content Interaction",
            "prompt": f"Target: {dt['target_report']}",
            "triage": "recommended",
            "notes": "Configure Jump-to-Slide or Jump-to-Report interaction passing context filters."
        })
        modeler_plan_rows.append({
            "stage": "7. Present",
            "category": "Interaction",
            "name": dt["action_type"],
            "ssrs_logic": f"Action to {dt.get('target_report', '')}",
            "pyramid_target": "Pyramid Jump-to-Content Action",
            "pyramid_action": f"Configure jump interaction to target report",
            "review_level": "recommended",
            "confidence": 88,
            "reason": "SSRS report action mapped to Pyramid Jump-to-Content",
        })

    # Stage 8: Validation (Always present)
    modeler_steps.append({
        "step_number": step_num,
        "stage": "8. Validation",
        "step_name": "Verification_Checklist",
        "pyramid_layer": "Validation",
        "source_node": "Migrated Report Stack",
        "operation": "Migration QA Verification",
        "target_node": "Production Promotion",
        "triage": "manual_review",
        "confidence": 1.0,
        "guidelines": "Execute row count reconciliation, measure verification, and slicer functional test."
    })
    modeler_plan_rows.append({
        "stage": "8. Validation",
        "category": "Validation",
        "name": "Migration Verification QA",
        "ssrs_logic": "Verification Checklist",
        "pyramid_target": "Pyramid Migration Validation Protocol",
        "pyramid_action": "Verify row counts, measure totals, slicers, and layout fidelity against SSRS source",
        "review_level": "manual_review",
        "confidence": 100,
        "reason": "Mandatory QA validation step before production deployment",
    })

    # Modeler Plan DataFrame
    modeler_plan_df = pd.DataFrame(modeler_plan_rows)

    # Triage Summary Stats
    all_triage = (
        [s["triage"] for s in modeler_steps] +
        [c["triage_level"] for c in generic_logic.get("calculations", [])] +
        [v["triage"] for v in discover_visuals] +
        [p["triage"] for p in present_items]
    )
    triage_counts = {
        "automatic": all_triage.count("automatic"),
        "recommended": all_triage.count("recommended"),
        "manual_review": all_triage.count("manual_review"),
    }
    total_items = len(all_triage) or 1
    complexity_score = min(100, int((triage_counts["manual_review"] * 5 + triage_counts["recommended"] * 2 + triage_counts["automatic"] * 0.5) / total_items * 25))

    return {
        "modeler_steps": modeler_steps,
        "modeler_plan_df": modeler_plan_df,
        "semantic_measures": semantic_measures,
        "semantic_attributes": semantic_attributes,
        "discover_visuals": discover_visuals,
        "present_items": present_items,
        "triage_counts": triage_counts,
        "complexity_score": complexity_score,
    }


# ==============================================================================
# 7. MULTI-SHEET EXCEL EVIDENCE GENERATOR (OPENPYXL)
# ==============================================================================

def _fill(hex_code: str) -> PatternFill:
    return PatternFill(start_color=hex_code, end_color=hex_code, fill_type="solid")


def _font(hex_code: str = "000000", bold: bool = False, size: int = 10) -> Font:
    return Font(name="Segoe UI", size=size, bold=bold, color=hex_code)


def _border_thin() -> Border:
    thin = Side(border_style="thin", color="D3D3D3")
    return Border(left=thin, right=thin, top=thin, bottom=thin)


def _write_cell(ws, row: int, col: int, val: Any, bg: str = "FFFFFF", fg: str = "000000", bold: bool = False, align: str = "left"):
    cell = ws.cell(row=row, column=col, value=str(val) if val is not None else "")
    cell.fill = _fill(bg)
    cell.font = _font(fg, bold=bold)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    cell.border = _border_thin()
    return cell


def setup_sheet(ws, headers: List[str], widths: List[int], title: str = ""):
    ws.views.sheetView[0].showGridLines = True
    start_row = 1
    if title:
        cell = ws.cell(row=1, column=1, value=title)
        cell.font = Font(name="Segoe UI", size=13, bold=True, color="1F497D")
        ws.row_dimensions[1].height = 24
        start_row = 3

    ws.row_dimensions[start_row].height = 24
    for col_idx, (header, width) in enumerate(zip(headers, widths), start=1):
        _write_cell(ws, start_row, col_idx, header, bg="1F497D", fg="FFFFFF", bold=True, align="center")
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width

    return start_row + 1


def build_output_workbook(parsed_results: List[Dict[str, Any]]) -> bytes:
    """Generate the complete 11-sheet Excel evidence workbook for SSRS -> Pyramid Analytics migration."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    all_generic = []
    all_pyramid_maps = []

    # 1. Overview Sheet
    ws_ov = wb.create_sheet("Overview")
    headers_ov = ["Report Name", "File", "DataSources", "DataSets", "Tables", "Parameters", "Visuals", "Expressions", "Complexity", "Automatic", "Recommended", "Manual Review"]
    widths_ov = [25, 25, 12, 12, 12, 12, 12, 14, 14, 12, 14, 14]
    r_ov = setup_sheet(ws_ov, headers_ov, widths_ov, "SSRS to Pyramid Analytics Migration — Portfolio Overview")

    for item in parsed_results:
        gen = build_generic_migration_representation(item)
        pyr = map_generic_to_pyramid(gen)
        all_generic.append(gen)
        all_pyramid_maps.append(pyr)

        bg = "F9FAFB" if (r_ov % 2 == 0) else "FFFFFF"
        _write_cell(ws_ov, r_ov, 1, item["report_name"], bg=bg, bold=True)
        _write_cell(ws_ov, r_ov, 2, item["filename"], bg=bg)
        _write_cell(ws_ov, r_ov, 3, len(item["data_sources"]), bg=bg, align="center")
        _write_cell(ws_ov, r_ov, 4, len(item["data_sets"]), bg=bg, align="center")
        _write_cell(ws_ov, r_ov, 5, len(gen["physical_tables"]), bg=bg, align="center")
        _write_cell(ws_ov, r_ov, 6, len(item["parameters"]), bg=bg, align="center")
        _write_cell(ws_ov, r_ov, 7, len(gen["visuals"]), bg=bg, align="center")
        _write_cell(ws_ov, r_ov, 8, len(item["expressions"]), bg=bg, align="center")
        _write_cell(ws_ov, r_ov, 9, f"{pyr['complexity_score']}/100", bg="E8F5E9" if pyr['complexity_score'] < 40 else "FFF3E0", bold=True, align="center")
        _write_cell(ws_ov, r_ov, 10, pyr["triage_counts"]["automatic"], bg=bg, align="center")
        _write_cell(ws_ov, r_ov, 11, pyr["triage_counts"]["recommended"], bg=bg, align="center")
        _write_cell(ws_ov, r_ov, 12, pyr["triage_counts"]["manual_review"], bg=bg, align="center")
        ws_ov.row_dimensions[r_ov].height = 20
        r_ov += 1

    # 2. Data Sources & DataSets Sheet
    ws_ds = wb.create_sheet("Data Sources & DataSets")
    headers_ds = ["Report Name", "DataSet Name", "Data Source", "Provider", "Command Type", "Physical Fields", "Calc Fields", "Parameters Count", "Filters Count"]
    widths_ds = [22, 22, 20, 18, 16, 15, 14, 16, 14]
    r_ds = setup_sheet(ws_ds, headers_ds, widths_ds, "SSRS DataSources and DataSets Inventory")
    for item in parsed_results:
        for dset in item.get("data_sets", []):
            bg = "F9FAFB" if (r_ds % 2 == 0) else "FFFFFF"
            _write_cell(ws_ds, r_ds, 1, item["report_name"], bg=bg)
            _write_cell(ws_ds, r_ds, 2, dset["name"], bg=bg, bold=True)
            _write_cell(ws_ds, r_ds, 3, dset["data_source_name"], bg=bg)
            provider = next((ds["data_provider"] for ds in item["data_sources"] if ds["name"] == dset["data_source_name"]), "Relational")
            _write_cell(ws_ds, r_ds, 4, provider, bg=bg)
            _write_cell(ws_ds, r_ds, 5, dset["command_type"], bg=bg, align="center")
            _write_cell(ws_ds, r_ds, 6, len(dset["physical_fields"]), bg=bg, align="center")
            _write_cell(ws_ds, r_ds, 7, len(dset["calculated_fields"]), bg=bg, align="center")
            _write_cell(ws_ds, r_ds, 8, len(dset["query_parameters"]), bg=bg, align="center")
            _write_cell(ws_ds, r_ds, 9, len(dset["filters"]), bg=bg, align="center")
            ws_ds.row_dimensions[r_ds].height = 20
            r_ds += 1

    # 3. Queries & SQL Sheet
    ws_sql = wb.create_sheet("Queries & SQL")
    headers_sql = ["Report Name", "DataSet Name", "Is DAX / Proc", "Extracted Tables", "CTEs", "Where Filter Clause", "Full CommandText"]
    widths_sql = [20, 20, 14, 25, 20, 35, 55]
    r_sql = setup_sheet(ws_sql, headers_sql, widths_sql, "Decomposed SQL & DAX Queries")
    for item in parsed_results:
        for dset in item.get("data_sets", []):
            decomp = dset.get("sql_decomposition", {})
            tbls = [t["table_name"] if isinstance(t, dict) else str(t) for t in decomp.get("tables", [])]
            query_nature = "DAX" if decomp.get("is_dax") else ("StoredProc" if decomp.get("is_stored_proc") else "SQL")
            bg = "F9FAFB" if (r_sql % 2 == 0) else "FFFFFF"
            _write_cell(ws_sql, r_sql, 1, item["report_name"], bg=bg)
            _write_cell(ws_sql, r_sql, 2, dset["name"], bg=bg, bold=True)
            _write_cell(ws_sql, r_sql, 3, query_nature, bg=bg, align="center")
            _write_cell(ws_sql, r_sql, 4, ", ".join(tbls), bg=bg)
            _write_cell(ws_sql, r_sql, 5, ", ".join(decomp.get("ctes", [])), bg=bg)
            _write_cell(ws_sql, r_sql, 6, decomp.get("where_clause", ""), bg=bg)
            _write_cell(ws_sql, r_sql, 7, dset.get("command_text", ""), bg=bg)
            ws_sql.row_dimensions[r_sql].height = 24
            r_sql += 1

    # 4. Joins & Relationships Sheet
    ws_jn = wb.create_sheet("Joins & Relationships")
    headers_jn = ["Report Name", "DataSet Name", "Target Table", "Join Type", "Alias", "Join Condition", "Pyramid Action", "Review Notes"]
    widths_jn = [20, 20, 25, 18, 12, 35, 25, 30]
    r_jn = setup_sheet(ws_jn, headers_jn, widths_jn, "SSRS Joins to Pyramid Modeler Relationships")
    for item in parsed_results:
        gen = build_generic_migration_representation(item)
        for j in gen.get("joins", []):
            bg = "F9FAFB" if (r_jn % 2 == 0) else "FFFFFF"
            _write_cell(ws_jn, r_jn, 1, item["report_name"], bg=bg)
            _write_cell(ws_jn, r_jn, 2, j["dataset_name"], bg=bg)
            _write_cell(ws_jn, r_jn, 3, j["target_table"], bg=bg, bold=True)
            _write_cell(ws_jn, r_jn, 4, j["join_type"], bg=bg, align="center")
            _write_cell(ws_jn, r_jn, 5, j["alias"], bg=bg, align="center")
            _write_cell(ws_jn, r_jn, 6, j["condition"], bg=bg)
            _write_cell(ws_jn, r_jn, 7, j["pyramid_action"], bg=bg)
            _write_cell(ws_jn, r_jn, 8, j["review_reason"], bg=bg)
            ws_jn.row_dimensions[r_jn].height = 20
            r_jn += 1

    # 5. Filters & Parameters Sheet
    ws_fp = wb.create_sheet("Filters & Parameters")
    headers_fp = ["Report Name", "Parameter / Filter Name", "Category", "Data Type", "Prompt", "Multi-Value", "Default Value", "Pyramid Slicer Target", "Triage"]
    widths_fp = [20, 25, 18, 14, 25, 14, 18, 28, 15]
    r_fp = setup_sheet(ws_fp, headers_fp, widths_fp, "SSRS Parameters and Filters Mapping")
    for item in parsed_results:
        gen = build_generic_migration_representation(item)
        for p in gen.get("parameters", []):
            bg = "F9FAFB" if (r_fp % 2 == 0) else "FFFFFF"
            _write_cell(ws_fp, r_fp, 1, item["report_name"], bg=bg)
            _write_cell(ws_fp, r_fp, 2, p["name"], bg=bg, bold=True)
            _write_cell(ws_fp, r_fp, 3, "Report Parameter", bg=bg)
            _write_cell(ws_fp, r_fp, 4, p["data_type"], bg=bg, align="center")
            _write_cell(ws_fp, r_fp, 5, p["prompt"], bg=bg)
            _write_cell(ws_fp, r_fp, 6, "Yes" if p["multi_value"] else "No", bg=bg, align="center")
            _write_cell(ws_fp, r_fp, 7, p["default_value"], bg=bg)
            _write_cell(ws_fp, r_fp, 8, p["pyramid_slicer"], bg=bg)
            _write_cell(ws_fp, r_fp, 9, p["triage"], bg="E8F5E9", align="center")
            ws_fp.row_dimensions[r_fp].height = 20
            r_fp += 1

        for f in gen.get("filters", []):
            bg = "F9FAFB" if (r_fp % 2 == 0) else "FFFFFF"
            _write_cell(ws_fp, r_fp, 1, item["report_name"], bg=bg)
            _write_cell(ws_fp, r_fp, 2, f["scope"], bg=bg, bold=True)
            _write_cell(ws_fp, r_fp, 3, "Query/DataSet Filter", bg=bg)
            _write_cell(ws_fp, r_fp, 4, "Expression", bg=bg, align="center")
            _write_cell(ws_fp, r_fp, 5, f["expression"], bg=bg)
            _write_cell(ws_fp, r_fp, 6, "-", bg=bg, align="center")
            _write_cell(ws_fp, r_fp, 7, "-", bg=bg)
            _write_cell(ws_fp, r_fp, 8, f["pyramid_target"], bg=bg)
            _write_cell(ws_fp, r_fp, 9, f["triage"], bg="FFF3E0", align="center")
            ws_fp.row_dimensions[r_fp].height = 20
            r_fp += 1

    # 6. Calculations & Expressions Sheet
    ws_calc = wb.create_sheet("Calculations & Expressions")
    headers_calc = ["Report Name", "Original SSRS Expression", "Expression Type", "Translated Pyramid PQL", "Proposed Pyramid Layer", "Triage Level", "Confidence", "Review Notes"]
    widths_calc = [20, 35, 20, 35, 28, 15, 14, 30]
    r_calc = setup_sheet(ws_calc, headers_calc, widths_calc, "SSRS Expressions & Pyramid PQL Translation")
    for item in parsed_results:
        gen = build_generic_migration_representation(item)
        for c in gen.get("calculations", []):
            bg = "F9FAFB" if (r_calc % 2 == 0) else "FFFFFF"
            triage_bg = "E8F5E9" if c["triage_level"] == "automatic" else "FFF3E0" if c["triage_level"] == "recommended" else "FFEBEE"
            _write_cell(ws_calc, r_calc, 1, item["report_name"], bg=bg)
            _write_cell(ws_calc, r_calc, 2, c["expression"], bg=bg, bold=True)
            _write_cell(ws_calc, r_calc, 3, c["type"], bg=bg)
            _write_cell(ws_calc, r_calc, 4, c["pql_translation"], bg=bg)
            _write_cell(ws_calc, r_calc, 5, c["proposed_layer"], bg=bg)
            _write_cell(ws_calc, r_calc, 6, c["triage_level"], bg=triage_bg, align="center", bold=True)
            _write_cell(ws_calc, r_calc, 7, f"{int(c['confidence']*100)}%", bg=bg, align="center")
            _write_cell(ws_calc, r_calc, 8, c["review_notes"], bg=bg)
            ws_calc.row_dimensions[r_calc].height = 20
            r_calc += 1

    # 7. Tablix & Visuals Sheet
    ws_vis = wb.create_sheet("Tablix & Visuals")
    headers_vis = ["Report Name", "Visual Name", "SSRS Visual Type", "DataSet Name", "Row Groups", "Column Groups", "Pyramid Discover Visual", "Triage"]
    widths_vis = [20, 22, 20, 20, 14, 15, 25, 15]
    r_vis = setup_sheet(ws_vis, headers_vis, widths_vis, "Tablix, Matrices, and Charts Inventory")
    for item in parsed_results:
        gen = build_generic_migration_representation(item)
        for v in gen.get("visuals", []):
            bg = "F9FAFB" if (r_vis % 2 == 0) else "FFFFFF"
            _write_cell(ws_vis, r_vis, 1, item["report_name"], bg=bg)
            _write_cell(ws_vis, r_vis, 2, v["name"], bg=bg, bold=True)
            _write_cell(ws_vis, r_vis, 3, v["type"], bg=bg)
            _write_cell(ws_vis, r_vis, 4, v["dataset_name"], bg=bg)
            _write_cell(ws_vis, r_vis, 5, v["row_groups_count"], bg=bg, align="center")
            _write_cell(ws_vis, r_vis, 6, v["col_groups_count"], bg=bg, align="center")
            _write_cell(ws_vis, r_vis, 7, v["pyramid_discover_type"], bg=bg)
            _write_cell(ws_vis, r_vis, 8, v["triage"], bg="E8F5E9" if v["triage"] == "automatic" else "FFF3E0", align="center")
            ws_vis.row_dimensions[r_vis].height = 20
            r_vis += 1

    # 8. Subreports & Drillthrough Sheet
    ws_sub = wb.create_sheet("Subreports & Drillthrough")
    headers_sub = ["Report Name", "Item Name", "Item Type", "Target Reference", "Parameters Passed", "Pyramid Present Target", "Triage"]
    widths_sub = [20, 22, 20, 25, 30, 28, 15]
    r_sub = setup_sheet(ws_sub, headers_sub, widths_sub, "Subreports, Drillthrough, and Navigation Links")
    for item in parsed_results:
        for sub in item.get("subreports", []):
            bg = "F9FAFB" if (r_sub % 2 == 0) else "FFFFFF"
            p_str = ", ".join([f"{p['name']}={p['value']}" for p in sub["parameters"]])
            _write_cell(ws_sub, r_sub, 1, item["report_name"], bg=bg)
            _write_cell(ws_sub, r_sub, 2, sub["name"], bg=bg, bold=True)
            _write_cell(ws_sub, r_sub, 3, "Subreport", bg=bg)
            _write_cell(ws_sub, r_sub, 4, sub["report_reference"], bg=bg)
            _write_cell(ws_sub, r_sub, 5, p_str, bg=bg)
            _write_cell(ws_sub, r_sub, 6, "Compound Present Slide / Publication", bg=bg)
            _write_cell(ws_sub, r_sub, 7, "manual_review", bg="FFEBEE", align="center", bold=True)
            ws_sub.row_dimensions[r_sub].height = 20
            r_sub += 1

        for dt in item.get("drillthroughs", []):
            bg = "F9FAFB" if (r_sub % 2 == 0) else "FFFFFF"
            p_str = ", ".join([f"{p['name']}={p['value']}" for p in dt["parameters"]])
            _write_cell(ws_sub, r_sub, 1, item["report_name"], bg=bg)
            _write_cell(ws_sub, r_sub, 2, dt["action_type"], bg=bg, bold=True)
            _write_cell(ws_sub, r_sub, 3, dt["action_type"], bg=bg)
            _write_cell(ws_sub, r_sub, 4, dt["target_report"], bg=bg)
            _write_cell(ws_sub, r_sub, 5, p_str, bg=bg)
            _write_cell(ws_sub, r_sub, 6, "Pyramid Jump-to-Slide Interaction", bg=bg)
            _write_cell(ws_sub, r_sub, 7, "recommended", bg="FFF3E0", align="center")
            ws_sub.row_dimensions[r_sub].height = 20
            r_sub += 1

    # 9. Pyramid Modeler Flow Sheet
    ws_mod = wb.create_sheet("Pyramid Modeler Flow")
    headers_mod = ["Report Name", "Step #", "Stage", "Step Name", "Source Node", "Operation", "Target Node", "Triage", "Guidelines"]
    widths_mod = [20, 8, 18, 25, 20, 25, 22, 14, 40]
    r_mod = setup_sheet(ws_mod, headers_mod, widths_mod, "Pyramid Modeler Data Flow Implementation Steps")
    for item, pyr in zip(parsed_results, all_pyramid_maps):
        for step in pyr["modeler_steps"]:
            bg = "F9FAFB" if (r_mod % 2 == 0) else "FFFFFF"
            _write_cell(ws_mod, r_mod, 1, item["report_name"], bg=bg)
            _write_cell(ws_mod, r_mod, 2, step["step_number"], bg=bg, align="center")
            _write_cell(ws_mod, r_mod, 3, step["stage"], bg=bg)
            _write_cell(ws_mod, r_mod, 4, step["step_name"], bg=bg, bold=True)
            _write_cell(ws_mod, r_mod, 5, step["source_node"], bg=bg)
            _write_cell(ws_mod, r_mod, 6, step["operation"], bg=bg)
            _write_cell(ws_mod, r_mod, 7, step["target_node"], bg=bg)
            _write_cell(ws_mod, r_mod, 8, step["triage"], bg="E8F5E9" if step["triage"] == "automatic" else "FFF3E0", align="center")
            _write_cell(ws_mod, r_mod, 9, step["guidelines"], bg=bg)
            ws_mod.row_dimensions[r_mod].height = 20
            r_mod += 1

    # 10. Pyramid Semantic Model Sheet
    ws_sem = wb.create_sheet("Pyramid Semantic Model")
    headers_sem = ["Report Name", "Element Name", "Element Type", "PQL Formula / Definition", "Format / Layer", "Triage", "Confidence"]
    widths_sem = [20, 25, 18, 40, 25, 14, 14]
    r_sem = setup_sheet(ws_sem, headers_sem, widths_sem, "Pyramid Semantic Model Measures & Attributes")
    for item, pyr in zip(parsed_results, all_pyramid_maps):
        for m in pyr["semantic_measures"]:
            bg = "F9FAFB" if (r_sem % 2 == 0) else "FFFFFF"
            _write_cell(ws_sem, r_sem, 1, item["report_name"], bg=bg)
            _write_cell(ws_sem, r_sem, 2, m["measure_name"], bg=bg, bold=True)
            _write_cell(ws_sem, r_sem, 3, "Measure", bg=bg)
            _write_cell(ws_sem, r_sem, 4, m["pql_formula"], bg=bg)
            _write_cell(ws_sem, r_sem, 5, m["format_string"], bg=bg)
            _write_cell(ws_sem, r_sem, 6, m["triage"], bg="E8F5E9", align="center")
            _write_cell(ws_sem, r_sem, 7, f"{int(m['confidence']*100)}%", bg=bg, align="center")
            ws_sem.row_dimensions[r_sem].height = 20
            r_sem += 1

        for a in pyr["semantic_attributes"]:
            bg = "F9FAFB" if (r_sem % 2 == 0) else "FFFFFF"
            _write_cell(ws_sem, r_sem, 1, item["report_name"], bg=bg)
            _write_cell(ws_sem, r_sem, 2, a["column_name"], bg=bg, bold=True)
            _write_cell(ws_sem, r_sem, 3, "Calculated Column", bg=bg)
            _write_cell(ws_sem, r_sem, 4, a["pql_formula"], bg=bg)
            _write_cell(ws_sem, r_sem, 5, a["layer"], bg=bg)
            _write_cell(ws_sem, r_sem, 6, a["triage"], bg="FFF3E0", align="center")
            _write_cell(ws_sem, r_sem, 7, f"{int(a['confidence']*100)}%", bg=bg, align="center")
            ws_sem.row_dimensions[r_sem].height = 20
            r_sem += 1

    # 11. Migration Triage Sheet
    ws_tr = wb.create_sheet("Migration Triage")
    headers_tr = ["Report Name", "Component", "Type", "Triage Category", "Pyramid Layer", "Confidence", "Developer Action Required"]
    widths_tr = [20, 25, 20, 16, 25, 14, 45]
    r_tr = setup_sheet(ws_tr, headers_tr, widths_tr, "Triage Classification & Developer Review Guide")
    for item, pyr, gen in zip(parsed_results, all_pyramid_maps, all_generic):
        for c in gen.get("calculations", []):
            if c["triage_level"] in {"recommended", "manual_review"}:
                bg = "FFF3E0" if c["triage_level"] == "recommended" else "FFEBEE"
                _write_cell(ws_tr, r_tr, 1, item["report_name"], bg=bg)
                _write_cell(ws_tr, r_tr, 2, c["expression"], bg=bg, bold=True)
                _write_cell(ws_tr, r_tr, 3, c["type"], bg=bg)
                _write_cell(ws_tr, r_tr, 4, c["triage_level"], bg=bg, align="center", bold=True)
                _write_cell(ws_tr, r_tr, 5, c["proposed_layer"], bg=bg)
                _write_cell(ws_tr, r_tr, 6, f"{int(c['confidence']*100)}%", bg=bg, align="center")
                _write_cell(ws_tr, r_tr, 7, c["review_notes"], bg=bg)
                ws_tr.row_dimensions[r_tr].height = 20
                r_tr += 1

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# ==============================================================================
# 8. STREAMLIT UI COMPONENTS & TECHNICAL ARCHITECTURE FLOW
# ==============================================================================

def render_technical_architecture_flow():
    """Renders the ASCII / visual Technical Migration Architecture Flow diagram."""
    st.markdown("""
```
SSRS / Power BI Paginated Report (.rdl, .xml, .zip)
                      │
                      ▼
┌────────────────────────────────────────────────────────┐
│ AST Extraction & SQL/DAX Decomposition Engine          │
│                                                        │
│ • Report & Page Properties (Margins, Units, Fonts)     │
│ • DataSources & Connection Strings                     │
│ • DataSets & Embedded CommandText SQL / DAX Queries    │
│ • Decomposition (FROM, JOIN, Multi-CTEs, WHERE, DAX)   │
│ • Report Parameters (Prompts, Defaults, ValidValues)   │
│ • Report Variables & Design System Theme Palettes      │
│ • Calculations & Aggregates (=Sum, =Avg, =IIF, etc.)   │
│ • Tablix Structures (Row Groups, Column Groups, Cells) │
│ • Rectangles & ColSpan Multi-Column Voucher Containers │
│ • Conditional Formatting (Color, Background, Alerts)   │
│ • Page Breaks by Group for Publication Bursting        │
└───────────────────────────┬────────────────────────────┘
                            │
                            ▼
┌────────────────────────────────────────────────────────┐
│ Generic Migration Representation                       │
│                                                        │
│ • Sources         • Joins           • Filters          │
│ • Parameters      • Calculations    • Visuals          │
│ • Drillthroughs   • Design Palettes • Outputs          │
└───────────────────────────┬────────────────────────────┘
                            │
                            ▼
┌────────────────────────────────────────────────────────┐
│ Pyramid Analytics Target Architecture Mapping          │
├────────────────────────────────────────────────────────┤
│  Modeler Data Flow  │ Source DB ➔ CTE Stages ➔ Joins   │
├─────────────────────┼──────────────────────────────────┤
│  Semantic Model     │ Tables ➔ Columns ➔ PQL Measures  │
├─────────────────────┼──────────────────────────────────┤
│  Discover           │ Grids, Matrix Tables, Alerts     │
├─────────────────────┼──────────────────────────────────┤
│  Present            │ Dashboards, Bursting, Slicers    │
└────────────────────────────────────────────────────────┘
```
""")


def render_kpi_cards(parsed_results: List[Dict[str, Any]]):
    """Renders developer-focused summary KPIs across the loaded SSRS portfolio."""
    total_reports = len(parsed_results)
    total_datasets = sum(len(r.get("data_sets", [])) for r in parsed_results)
    total_tables = sum(sum(len(d.get("sql_decomposition", {}).get("tables", [])) for d in r.get("data_sets", [])) for r in parsed_results)
    total_visuals = sum(len(r.get("tablixes", [])) + len(r.get("charts", [])) for r in parsed_results)
    total_expressions = sum(len(r.get("expressions", [])) for r in parsed_results)
    
    auto_count = 0
    rec_count = 0
    man_count = 0
    for r in parsed_results:
        gen = build_generic_migration_representation(r)
        pyr = map_generic_to_pyramid(gen)
        auto_count += pyr["triage_counts"]["automatic"]
        rec_count += pyr["triage_counts"]["recommended"]
        man_count += pyr["triage_counts"]["manual_review"]

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1:
        st.metric("Reports", total_reports)
    with c2:
        st.metric("DataSets", total_datasets)
    with c3:
        st.metric("Source Tables", total_tables)
    with c4:
        st.metric("Visuals & Grids", total_visuals)
    with c5:
        st.metric("Expressions", total_expressions)
    with c6:
        st.metric("Triage Auto / Rec / Man", f"{auto_count} / {rec_count} / {man_count}")


# ==============================================================================
# 9. PERSISTENT TAB RENDERERS (DEVELOPER VIEW)
# ==============================================================================

def render_start_here_tab(report: Dict[str, Any], generic: Dict[str, Any], pyramid: Dict[str, Any]):
    st.subheader("Migration Blueprint & Portfolio Summary")
    
    col_a, col_b = st.columns([1, 1])
    with col_a:
        st.markdown(f"""
**Report Name:** `{report['report_name']}`  
**Source File:** `{report['filename']}`  
**Report ID:** `{report['metadata'].get('report_id', 'N/A')}`  
**Default Font:** `{report['metadata'].get('default_font', 'Segoe UI')}`  
**Page Dimensions:** `{report['metadata'].get('page_width', 'N/A')} x {report['metadata'].get('page_height', 'N/A')}`  
**Estimated Complexity Score:** `{pyramid['complexity_score']} / 100`  
""")
    with col_b:
        st.markdown(f"""
**Architecture Mapping Highlights:**
- **DataSources:** {len(report['data_sources'])} Relational/DAX connection(s)
- **DataSets:** {len(report['data_sets'])} SQL / DAX Queries
- **Extracted Tables:** {len(generic['physical_tables'])} Database views/tables
- **Report Parameters:** {len(report['parameters'])} Filter prompts
- **Visual Items:** {len(report['tablixes'])} Tablixes + {len(report['charts'])} Charts
- **Calculated Measures:** {len(pyramid['semantic_measures'])} PQL Measures
- **Design Variables:** {len(report.get('variables', []))} Design System variables
""")

    st.markdown("---")
    st.markdown("#### Technical Migration Architecture Flow")
    render_technical_architecture_flow()

    st.markdown("---")
    st.markdown("#### Migration Checklist & Next Steps")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.success(f"**Automatic Steps ({pyramid['triage_counts']['automatic']})**\nDirect translation of relational sources, table queries, standard aggregations, and single-select prompts.")
    with c2:
        st.warning(f"**Recommended Steps ({pyramid['triage_counts']['recommended']})**\nVerify SQL CTEs in Pyramid Modeler, validate multi-value slicers, and review matrix row/col hierarchies.")
    with c3:
        st.info(f"**Manual Review Steps ({pyramid['triage_counts']['manual_review']})**\nReview custom VB code, stored procedures, and subreport dependencies in Present.")


def render_datasources_datasets_tab(report: Dict[str, Any]):
    st.subheader("Data Sources")
    if report["data_sources"]:
        ds_data = []
        for ds in report["data_sources"]:
            ds_data.append({
                "Data Source Name": ds["name"],
                "Data Provider": ds["data_provider"],
                "Connection String": ds["connect_string"] or "(Embedded in Server)",
                "Security Type": ds["security_type"] or "Integrated",
                "Reference": ds["reference"] or "None"
            })
        st.dataframe(pd.DataFrame(ds_data), use_container_width=True)
    else:
        st.info("No explicit DataSources declared in this report definition.")

    if report.get("variables"):
        st.markdown("---")
        st.subheader("Report Design System Variables")
        v_df = pd.DataFrame(report["variables"])
        st.dataframe(v_df, use_container_width=True)

    st.markdown("---")
    st.subheader("DataSets & Query Definitions")
    if report["data_sets"]:
        for dset in report["data_sets"]:
            with st.expander(f"DataSet: {dset['name']} (Source: {dset['data_source_name']}, Type: {dset['command_type']})", expanded=True):
                c1, c2 = st.columns([1, 1])
                with c1:
                    st.markdown(f"**Physical Fields ({len(dset['physical_fields'])}):**")
                    if dset["physical_fields"]:
                        st.dataframe(pd.DataFrame(dset["physical_fields"]), height=200, use_container_width=True)
                    else:
                        st.write("No physical fields declared.")
                with c2:
                    st.markdown(f"**Calculated Fields ({len(dset['calculated_fields'])}):**")
                    if dset["calculated_fields"]:
                        st.dataframe(pd.DataFrame(dset["calculated_fields"]), height=200, use_container_width=True)
                    else:
                        st.write("No calculated fields declared.")

                if dset["query_parameters"]:
                    st.markdown(f"**Query Parameters ({len(dset['query_parameters'])}):**")
                    st.dataframe(pd.DataFrame(dset["query_parameters"]), use_container_width=True)


def render_queries_sql_tab(report: Dict[str, Any]):
    st.subheader("Decomposed SQL & DAX Queries")
    for dset in report["data_sets"]:
        decomp = dset.get("sql_decomposition", {})
        query_type = "Power BI DAX" if decomp.get("is_dax") else ("Stored Procedure" if decomp.get("is_stored_proc") else "SQL Query")
        st.markdown(f"### DataSet: `{dset['name']}` ({query_type})")
        
        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("Source Tables Extracted", len(decomp.get("tables", [])))
        with c2:
            st.metric("CTEs Extracted", len(decomp.get("ctes", [])))
        with c3:
            st.metric("Joins Extracted", len(decomp.get("joins", [])))

        if decomp.get("ctes"):
            st.markdown("**Common Table Expressions (CTEs):**")
            st.write(", ".join([f"`{cte}`" for cte in decomp["ctes"]]))

        if decomp.get("tables"):
            st.markdown("**Extracted Tables & Views:**")
            t_rows = []
            for t in decomp["tables"]:
                t_rows.append({
                    "Table / View Name": t["table_name"] if isinstance(t, dict) else str(t),
                    "Alias": t.get("alias", "") if isinstance(t, dict) else "",
                    "Is CTE": "Yes" if (isinstance(t, dict) and t.get("is_cte")) else "No",
                })
            st.dataframe(pd.DataFrame(t_rows), use_container_width=True)

        if decomp.get("where_clause"):
            st.markdown("**WHERE Filter Predicates:**")
            st.code(decomp["where_clause"], language="sql")

        st.markdown("**Raw CommandText:**")
        st.code(dset.get("command_text", ""), language="sql" if not decomp.get("is_dax") else "text")
        st.markdown("---")


def render_joins_relationships_tab(report: Dict[str, Any], generic: Dict[str, Any]):
    st.subheader("Joins & Relationship Graph")
    if generic["joins"]:
        j_df = pd.DataFrame(generic["joins"])
        st.dataframe(j_df, use_container_width=True)
        st.markdown("#### Modeler Relationship Recommendations")
        for j in generic["joins"]:
            st.info(f"**Join on `{j['target_table']}`:** {j['join_type']} condition `{j['condition']}` ➔ **Pyramid Action:** {j['pyramid_action']} ({j['review_reason']})")
    else:
        st.info("No multi-table SQL joins were extracted. Queries operate as single-table/view extracts.")


def render_filters_parameters_tab(report: Dict[str, Any], generic: Dict[str, Any]):
    st.subheader("Report Parameters")
    if generic["parameters"]:
        p_df = pd.DataFrame(generic["parameters"])
        st.dataframe(p_df, use_container_width=True)
    else:
        st.info("No report-level parameters defined in this RDL.")

    st.markdown("---")
    st.subheader("DataSet & Visual Filters")
    if generic["filters"]:
        f_df = pd.DataFrame(generic["filters"])
        st.dataframe(f_df, use_container_width=True)
    else:
        st.info("No DataSet-level filters found.")


def render_calculations_expressions_tab(report: Dict[str, Any], generic: Dict[str, Any]):
    st.subheader("SSRS Calculations & Translated Pyramid PQL")
    if generic["calculations"]:
        calc_df = pd.DataFrame(generic["calculations"])
        
        f_type = st.multiselect("Filter by Expression Type", sorted(list(calc_df["type"].unique())), default=list(calc_df["type"].unique()))
        filtered_df = calc_df[calc_df["type"].isin(f_type)]
        st.dataframe(filtered_df, use_container_width=True)

        st.markdown("#### Sample PQL Translations")
        sample_calcs = filtered_df[filtered_df["type"].isin(["Aggregation / Measure", "Conditional Logic", "Scoped Aggregation (Group Subtotal)"])].head(5)
        for _, r in sample_calcs.iterrows():
            with st.expander(f"{r['type']}: {r['expression']}"):
                st.markdown(f"**SSRS Expression:** `{r['expression']}`")
                st.markdown(f"**Pyramid PQL:** `{r['pql_translation']}`")
                st.markdown(f"**Target Layer:** `{r['proposed_layer']}` | **Triage:** `{r['triage_level']}` ({int(r['confidence']*100)}% confidence)")
    else:
        st.info("No dynamic expressions extracted.")


def render_grouping_sorting_tab(report: Dict[str, Any]):
    st.subheader("Tablix Grouping, Sorting & Page Breaks")
    has_groups = False
    for tab in report["tablixes"]:
        if tab["row_groups"] or tab["col_groups"]:
            has_groups = True
            st.markdown(f"### Tablix: `{tab['name']}` (DataSet: `{tab['dataset_name']}`)")
            if tab["row_groups"]:
                st.markdown("**Row Groups:**")
                rg_data = []
                for rg in tab["row_groups"]:
                    rg_data.append({
                        "Group Name": rg["group_name"],
                        "Expressions": ", ".join(rg["expressions"]),
                        "Sort Expressions": ", ".join(rg["sort_expressions"]),
                        "Page Break Between Groups": "Yes" if rg.get("page_break") else "No"
                    })
                st.dataframe(pd.DataFrame(rg_data), use_container_width=True)

            if tab["col_groups"]:
                st.markdown("**Column Groups (Matrix Pivots):**")
                cg_data = []
                for cg in tab["col_groups"]:
                    cg_data.append({
                        "Group Name": cg["group_name"],
                        "Expressions": ", ".join(cg["expressions"])
                    })
                st.dataframe(pd.DataFrame(cg_data), use_container_width=True)
    if not has_groups:
        st.info("No custom row/column groups defined in Tablix structures.")


def render_tables_tablix_tab(report: Dict[str, Any]):
    st.subheader("Tablix Structures, Nested Rectangles & ColSpans")
    if report["tablixes"]:
        t_data = []
        for t in report["tablixes"]:
            t_data.append({
                "Tablix Name": t["name"],
                "DataSet Name": t["dataset_name"],
                "Is Matrix": "Yes" if t["is_matrix"] else "No (Tabular)",
                "Row Groups Count": len(t["row_groups"]),
                "Column Groups Count": len(t["col_groups"]),
                "Nested Rectangles": len(t.get("nested_rectangles", [])),
                "Cell Expressions Count": t["cell_expressions_count"]
            })
        st.dataframe(pd.DataFrame(t_data), use_container_width=True)

        for t in report["tablixes"]:
            if t.get("nested_rectangles"):
                st.info(f"**Nested Layout Alert:** `{t['name']}` contains {len(t['nested_rectangles'])} nested Rectangle container(s) across multi-column ColSpan cells. Maps to Pyramid Present Composite Cards.")

            with st.expander(f"Tablix: {t['name']} - Cell Expressions Sample", expanded=False):
                if t["sample_expressions"]:
                    for expr in t["sample_expressions"]:
                        st.code(expr, language="vb")
                else:
                    st.write("No dynamic cell expressions in this Tablix.")
    else:
        st.info("No Tablixes found in report.")


def render_charts_visuals_tab(report: Dict[str, Any]):
    st.subheader("Charts & Visualizations")
    if report["charts"]:
        c_data = []
        for c in report["charts"]:
            c_data.append({
                "Chart Name": c["name"],
                "DataSet Name": c["dataset_name"],
                "Chart Type": c["chart_type"],
                "Categories": ", ".join(c["categories"]),
                "Series Groups": ", ".join(c["series_groups"]),
                "Measures Count": len(c["measures"])
            })
        st.dataframe(pd.DataFrame(c_data), use_container_width=True)
    else:
        st.info("No Chart elements found in report.")


def render_subreports_drillthrough_tab(report: Dict[str, Any]):
    st.subheader("Subreports, Drillthrough Links & Conditional Alerts")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### Subreports")
        if report["subreports"]:
            s_data = [{"Subreport Name": s["name"], "Target Report": s["report_reference"], "Parameters Passed": len(s["parameters"])} for s in report["subreports"]]
            st.dataframe(pd.DataFrame(s_data), use_container_width=True)
        else:
            st.info("No subreports referenced.")

    with c2:
        st.markdown("#### Drillthrough Actions & Links")
        if report["drillthroughs"]:
            dt_data = [{"Action Type": dt["action_type"], "Target Destination": dt["target_report"], "Parameters Passed": len(dt["parameters"])} for dt in report["drillthroughs"]]
            st.dataframe(pd.DataFrame(dt_data), use_container_width=True)
        else:
            st.info("No drillthrough actions configured.")

    if report.get("conditional_styles"):
        st.markdown("---")
        st.subheader("Conditional Formatting & Visibility Alerts")
        st.dataframe(pd.DataFrame(report["conditional_styles"]), use_container_width=True)


# ==============================================================================
# 9. PYRAMID ANALYTICS BUILD PLAN & TARGET ARCHITECTURE INSPECTORS
# ==============================================================================

VALID_PYRAMID_LAYERS = [
    "Data Source / Model Table",
    "Data Flow Step",
    "Semantic Relationship",
    "Calculated Column",
    "Formulate Measure (PQL)",
    "Discover Filter / Slicer",
    "Parameter / Global List",
    "Discover Visual (Grid/Chart)",
    "Present Dashboard",
    "Publish Publication",
    "Manual Review",
]

PYRAMID_SYNONYMS = {
    "data source": "Data Source / Model Table",
    "model table": "Data Source / Model Table",
    "data model": "Data Source / Model Table",
    "data flow": "Data Flow Step",
    "data flow step": "Data Flow Step",
    "data prep": "Data Flow Step",
    "data prep step": "Data Flow Step",
    "relationship": "Semantic Relationship",
    "join": "Semantic Relationship",
    "semantic relationship": "Semantic Relationship",
    "calculated column": "Calculated Column",
    "calculated field": "Calculated Column",
    "measure": "Formulate Measure (PQL)",
    "pql measure": "Formulate Measure (PQL)",
    "formulate measure": "Formulate Measure (PQL)",
    "formulate measure (pql)": "Formulate Measure (PQL)",
    "filter": "Discover Filter / Slicer",
    "slicer": "Discover Filter / Slicer",
    "discover filter": "Discover Filter / Slicer",
    "discover filter / slicer": "Discover Filter / Slicer",
    "parameter": "Parameter / Global List",
    "global list": "Parameter / Global List",
    "parameter / global list": "Parameter / Global List",
    "visual": "Discover Visual (Grid/Chart)",
    "grid": "Discover Visual (Grid/Chart)",
    "matrix": "Discover Visual (Grid/Chart)",
    "chart": "Discover Visual (Grid/Chart)",
    "discover visual": "Discover Visual (Grid/Chart)",
    "discover visual (grid/chart)": "Discover Visual (Grid/Chart)",
    "dashboard": "Present Dashboard",
    "present dashboard": "Present Dashboard",
    "present": "Present Dashboard",
    "publication": "Publish Publication",
    "publish publication": "Publish Publication",
    "burst": "Publish Publication",
    "bursting": "Publish Publication",
    "manual review": "Manual Review",
}


def _coerce_pyramid_layer(value: Any) -> str:
    if not value:
        return "Manual Review"
    val_clean = str(value).strip()
    if val_clean in VALID_PYRAMID_LAYERS:
        return val_clean
    val_lower = val_clean.lower()
    if val_lower in PYRAMID_SYNONYMS:
        return PYRAMID_SYNONYMS[val_lower]
    for k, v in PYRAMID_SYNONYMS.items():
        if k in val_lower:
            return v
    return "Manual Review"


Classification = Literal["confirmed", "rule_based", "inferred", "suggested", "manual_review"]


class PlanStep(BaseModel):
    model_config = ConfigDict(populate_by_name=True)

    step_number: int = Field(default=1, ge=1)
    title: str = Field(default="Step")
    action: str = Field(default="")
    pyramid_layer: str = Field(
        default="Manual Review",
        validation_alias=AliasChoices('pyramid_layer', 'layer', 'target_layer')
    )
    classification: Classification = Field(default="manual_review")
    reason: str = Field(default="")
    referenced_sources: List[str] = Field(default_factory=list)
    referenced_fields: List[str] = Field(default_factory=list)
    confidence: int = Field(default=75, ge=0, le=100)
    manual_validation: Optional[str] = None

    @field_validator('pyramid_layer', mode='before')
    @classmethod
    def coerce_layer(cls, v):
        return _coerce_pyramid_layer(v)

    @field_validator('classification', mode='before')
    @classmethod
    def coerce_classification(cls, v):
        valid = {"confirmed", "rule_based", "inferred", "suggested", "manual_review"}
        v_str = str(v or '').strip().lower().replace(' ', '_').replace('-', '_')
        if v_str in valid:
            return v_str
        return "manual_review"


class ReportSummary(BaseModel):
    report_name: str
    likely_purpose: str
    report_type: str
    complexity_level: Literal["Very Small", "Small", "Medium", "Large", "Very Complex"]
    metadata_level: Literal["rdl_only"] = "rdl_only"
    overall_confidence: int = Field(ge=0, le=100)


class BuildPlan(BaseModel):
    report_summary: ReportSummary
    implementation_plan: List[PlanStep]
    data_preparation: List[PlanStep] = Field(default_factory=list)
    calculation_plan: List[PlanStep] = Field(default_factory=list)
    filter_parameter_plan: List[PlanStep] = Field(default_factory=list)
    page_recommendations: List[PlanStep] = Field(default_factory=list)
    validation_checks: List[str] = Field(default_factory=list)
    manual_review_items: List[str] = Field(default_factory=list)
    limitations: List[str] = Field(default_factory=list)


def badge(text: str, level: str = 'default') -> str:
    styles = {
        'default': 'background: #eef2f6; color: #334155; border: 1px solid #cbd5e1;',
        'ok': 'background: #ecfdf5; color: #065f46; border: 1px solid #a7f3d0;',
        'medium': 'background: #fffbeb; color: #92400e; border: 1px solid #fde68a;',
        'high': 'background: #fef2f2; color: #991b1b; border: 1px solid #fecaca;',
    }
    st_style = styles.get(level, styles['default'])
    return f'<span style="display: inline-block; padding: 2px 8px; border-radius: 9999px; font-size: 0.75rem; font-weight: 600; {st_style}">{text}</span>'


def _render_plan_step(step: PlanStep, source_df: Optional[pd.DataFrame] = None):
    level_map = {'confirmed': 'ok', 'rule_based': 'ok', 'inferred': 'medium', 'suggested': 'medium', 'manual_review': 'high'}
    classification = getattr(step, 'classification', 'manual_review')
    level = level_map.get(classification, 'medium')
    raw_layer = getattr(step, 'pyramid_layer', 'Manual Review')
    layer = _coerce_pyramid_layer(raw_layer)
    confidence = getattr(step, 'confidence', 75)
    title = getattr(step, 'title', 'Step')
    action = getattr(step, 'action', '')
    reason = getattr(step, 'reason', '')
    sources = getattr(step, 'referenced_sources', []) or []
    fields = getattr(step, 'referenced_fields', []) or []
    manual_validation = getattr(step, 'manual_validation', None)

    with st.container(border=True):
        st.markdown(f"**{title}** — {layer}")
        st.markdown(
            f"{badge(str(classification).replace('_', ' ').title(), level)} "
            f"{badge(f'Confidence: {confidence}%')}",
            unsafe_allow_html=True,
        )
        st.write(action)
        if reason:
            st.caption(f"Why: {reason}")
        if sources or fields:
            refs = ', '.join(sources + fields)
            st.caption(f"References: {refs}")
        if manual_validation:
            st.warning(manual_validation)


def get_openai_client():
    key = None
    if 'openai_api_key' in st.session_state and st.session_state['openai_api_key']:
        key = st.session_state['openai_api_key']
    elif hasattr(st, 'secrets') and 'OPENAI_API_KEY' in st.secrets:
        key = st.secrets['OPENAI_API_KEY']
    else:
        key = os.environ.get('OPENAI_API_KEY')

    if not key:
        return None

    try:
        from openai import OpenAI
        return OpenAI(api_key=key)
    except Exception:
        return None


def parse_llm_json(text: str) -> Optional[Dict[str, Any]]:
    if not text:
        return None
    clean = re.sub(r'^```(?:json)?\s*', '', text.strip(), flags=re.MULTILINE)
    clean = re.sub(r'\s*```$', '', clean.strip(), flags=re.MULTILINE)
    try:
        return json.loads(clean)
    except Exception:
        m = re.search(r'(\{[\s\S]*\})', text)
        if m:
            try:
                return json.loads(m.group(1))
            except Exception:
                pass
    return None


AI_BUILD_PLAN_SYSTEM_PROMPT = """You are a senior BI migration engineer converting one parsed SSRS (SQL Server Reporting Services) RDL report to Pyramid Analytics.
The input contains RDL-derived report logic only (DataSources, DataSets, Query SQL/CTEs, joins, parameters, expressions, tablix, charts, grouping, subreports).
rule_findings, if present in the input, are deterministic and authoritative - never re-derive or contradict them.

Pyramid Analytics target concepts to use:
- Data Source / Model Table: initial data source, database connection, or physical model table in Pyramid Data Model.
- Data Flow Step: Pyramid Data Flow (Prep) ETL step for staging, joins, CTE processing, column cleanup, or query lineage before model load.
- Semantic Relationship: join relationship defined between tables in the Pyramid Data Model (cardinality, cross-filter).
- Calculated Column: row-level calculated or virtual column in the Pyramid Data Model.
- Formulate Measure (PQL): custom metric, KPI, or aggregate formula created in Pyramid Formulate using PQL (Pyramid Query Language).
- Discover Filter / Slicer: slice, background filter, or interactive slicer configured in Pyramid Discover.
- Parameter / Global List: global parameter or dynamic custom list in Pyramid Formulate (equivalent to SSRS ReportParameters).
- Discover Visual (Grid/Chart): discovery visual in Pyramid Discover (Tablix/Grid, Matrix, Chart).
- Present Dashboard: interactive multi-visual dashboard, card/rectangle composition, or canvas in Pyramid Present.
- Publish Publication: scheduled or burst distribution publication in Pyramid Publish (e.g. for SSRS Page Breaks, page-by-slice bursting).
- Manual Review: items requiring human engineer review or verification.

Rules:
1. Use only tables, sources, fields, and calculations present in the input. Never invent a table, field, measure, join condition, or business rule.
2. Do not claim primary keys, cardinality, fact/dimension roles, or table grain unless explicitly supplied. If unresolvable from RDL, cite "Manual Review / Source metadata required".
3. Label every implementation step as one of: confirmed, rule_based, inferred, suggested, manual_review.
4. Every step must include referenced_sources and referenced_fields drawn only from the input - this is used for automated validation, so accuracy here is critical.
5. Any uncertainty must go into manual_review_items, not be silently guessed.
6. Confidence (0-100) must reflect how directly the input supports the recommendation, not general optimism.
7. Return ONLY strict JSON matching the required schema - no markdown fences, no commentary outside the JSON.

Return JSON with exactly these keys:
{
  "report_summary": {"report_name": "", "likely_purpose": "", "report_type": "",
                      "complexity_level": "Very Small|Small|Medium|Large|Very Complex",
                      "metadata_level": "rdl_only", "overall_confidence": 0},
  "implementation_plan": [{"step_number": 1, "title": "", "action": "",
                            "pyramid_layer": "Data Source / Model Table|Data Flow Step|Semantic Relationship|Calculated Column|Formulate Measure (PQL)|Discover Filter / Slicer|Parameter / Global List|Discover Visual (Grid/Chart)|Present Dashboard|Publish Publication|Manual Review",
                            "classification": "confirmed|rule_based|inferred|suggested|manual_review",
                            "reason": "", "referenced_sources": [], "referenced_fields": [],
                            "confidence": 0, "manual_validation": null}],
  "data_preparation": [... same PlanStep shape ...],
  "calculation_plan": [... same PlanStep shape ..., must cover every item in unresolved_formulas],
  "filter_parameter_plan": [... same PlanStep shape ..., must cover parameters and joins],
  "page_recommendations": [... same PlanStep shape ...],
  "validation_checks": ["..."],
  "manual_review_items": ["..."],
  "limitations": ["..."]
}
"""


def build_ai_build_plan_payload(report: Dict[str, Any], generic_logic: Dict[str, Any], pyramid_map: Dict[str, Any]) -> Dict[str, Any]:
    sources = [s["source_id"] for s in generic_logic.get("sources", [])]
    tables = [t["table_name"] for t in generic_logic.get("physical_tables", [])]

    cte_steps = []
    for tbl in generic_logic.get("physical_tables", []):
        if tbl.get("is_cte"):
            cte_steps.append({
                "cte_name": tbl["table_name"],
                "dataset": tbl.get("dataset_name", ""),
            })

    joins = []
    high_risk_joins = []
    for j in generic_logic.get("joins", []):
        j_entry = {
            "dataset": j.get("dataset_name", ""),
            "target_table": j.get("target_table", ""),
            "join_type": j.get("join_type", "INNER"),
            "condition": j.get("condition", ""),
        }
        joins.append(j_entry)
        if "OUTER" in j.get("join_type", "").upper() or not j.get("condition"):
            high_risk_joins.append(j_entry)

    filters = [f["expression"][:80] for f in generic_logic.get("filters", [])][:10]
    parameters = [f"{p['name']} ({p.get('data_type', 'Text')})" for p in generic_logic.get("parameters", [])]

    unresolved_formulas = []
    for c in generic_logic.get("calculations", []):
        if c.get("triage_level") in {"recommended", "manual_review"} or "scoped" in c.get("type", "").lower() or "iif" in c.get("expression", "").lower():
            unresolved_formulas.append({
                "expression": c["expression"][:150],
                "type": c["type"],
                "pql_translation": c.get("pql_translation", "")[:150],
                "review_notes": c.get("review_notes", "")[:100],
            })
    unresolved_formulas = unresolved_formulas[:15]

    visuals = []
    for v in generic_logic.get("visuals", []):
        visuals.append({
            "name": v["name"],
            "type": v["type"],
            "pyramid_discover_type": v["pyramid_discover_type"],
            "row_groups": v.get("row_groups_count", 0),
            "col_groups": v.get("col_groups_count", 0),
            "has_page_break": v.get("has_page_break", False),
        })

    subreports = [s["name"] for s in generic_logic.get("subreports", [])]
    drillthroughs = [d["action_type"] for d in generic_logic.get("drillthroughs", [])]

    score = pyramid_map.get("complexity_score", 30)
    if score >= 75:
        complexity_level = "Very Complex"
    elif score >= 55:
        complexity_level = "Large"
    elif score >= 35:
        complexity_level = "Medium"
    elif score >= 20:
        complexity_level = "Small"
    else:
        complexity_level = "Very Small"

    return {
        "report_name": report.get("report_name", "Report"),
        "complexity_level": complexity_level,
        "sources": sources,
        "physical_tables": tables,
        "cte_prep_steps": cte_steps,
        "joins": joins,
        "high_risk_joins": high_risk_joins,
        "filters_sample": filters,
        "parameters": parameters,
        "unresolved_formulas": unresolved_formulas,
        "visuals_summary": visuals,
        "subreports": subreports,
        "drillthroughs": drillthroughs,
        "constraints": {
            "database_metadata_available": False,
            "may_not_state_primary_keys": True,
            "may_not_state_cardinality": True,
            "may_not_state_confirmed_grain": True,
            "may_not_invent_tables_or_fields": True,
        }
    }


def generate_ai_build_plan(client, model: str, payload: Dict[str, Any]):
    user_prompt = json.dumps(payload, ensure_ascii=False)
    complexity = payload.get('complexity_level', 'Medium')
    base_budget = {
        'Very Small': 1500, 'Small': 2200, 'Medium': 3500,
        'Large': 5500, 'Very Complex': 7000,
    }.get(complexity, 3500)
    max_budget_cap = 16000
    output_budget = base_budget

    while True:
        try:
            if hasattr(client, 'chat') and hasattr(client.chat, 'completions'):
                res = client.chat.completions.create(
                    model=model,
                    messages=[
                        {"role": "system", "content": AI_BUILD_PLAN_SYSTEM_PROMPT},
                        {"role": "user", "content": user_prompt},
                    ],
                    temperature=0.1,
                    max_tokens=output_budget,
                )
                last_raw_text = res.choices[0].message.content or ''
                was_truncated = getattr(res.choices[0], 'finish_reason', None) == 'length'
            else:
                return None, "OpenAI client does not support chat completions."
        except Exception as e:
            return None, f"LLM error: {e}"

        raw = parse_llm_json(last_raw_text)
        if not raw and was_truncated and output_budget < max_budget_cap:
            output_budget = min(output_budget * 2, max_budget_cap)
            continue
        if not raw:
            return None, f"Model did not return parseable JSON. Raw: {last_raw_text[:400]}"

        for section in ['implementation_plan', 'data_preparation', 'calculation_plan', 'filter_parameter_plan', 'page_recommendations']:
            for s in raw.get(section, []):
                if isinstance(s, dict):
                    s['pyramid_layer'] = _coerce_pyramid_layer(s.get('pyramid_layer'))

        try:
            plan = BuildPlan.model_validate(raw)
            return plan, None
        except ValidationError as ve:
            return None, f"AI response failed schema validation: {ve}"


def build_pyramid_prep_build_order(generic_logic: Dict[str, Any], plan: Optional[BuildPlan]):
    ordered_stages = []
    sources = generic_logic.get("sources", [])
    if sources:
        ordered_stages.append({
            "stage_number": 1,
            "output_table": "Data Sources Ingestion",
            "what_happens": f"Connect to {', '.join(s['source_id'] for s in sources)} ({sources[0].get('source_type', 'DB')})",
            "keys": "",
            "matched_steps": [],
        })

    stage_idx = 2
    for tbl in generic_logic.get("physical_tables", []):
        if tbl.get("is_cte"):
            ordered_stages.append({
                "stage_number": stage_idx,
                "output_table": f"CTE Node: {tbl['table_name']}",
                "what_happens": f"Process Common Table Expression '{tbl['table_name']}' in Modeler Data Flow",
                "keys": tbl.get("alias", ""),
                "matched_steps": [],
            })
            stage_idx += 1

    ordered_stages.append({
        "stage_number": stage_idx,
        "output_table": f"Model Target: {generic_logic.get('report_name', 'Model')}",
        "what_happens": "Apply column transformations, rename fields, and load into Pyramid Semantic Model",
        "keys": "",
        "matched_steps": [],
    })

    if plan is not None:
        all_steps = (
            getattr(plan, 'data_preparation', []) or []
        ) + (
            getattr(plan, 'calculation_plan', []) or []
        ) + (
            getattr(plan, 'filter_parameter_plan', []) or []
        )
        for step in all_steps:
            matched = False
            sources = getattr(step, 'referenced_sources', []) or []
            fields = getattr(step, 'referenced_fields', []) or []
            refs = set(sources) | set(fields)
            for stage in ordered_stages:
                if stage['output_table'] in refs or any(r in stage['what_happens'] for r in refs):
                    stage['matched_steps'].append(step)
                    matched = True
                    break
            if not matched and ordered_stages:
                ordered_stages[-1]['matched_steps'].append(step)

    return ordered_stages


def render_pyramid_prep_build_order(generic_logic: Dict[str, Any], plan: Optional[BuildPlan]):
    stages = build_pyramid_prep_build_order(generic_logic, plan)
    if not stages:
        st.info("No staging steps detected for this report.")
        return

    st.markdown(
        "Build these as **Pyramid Data Flow flow steps (or Data Source tab steps), in this exact order**. "
        "Each stage below corresponds to one SSRS query stage / CTE, in the sequence it was originally executed."
    )

    for stage in stages:
        with st.container(border=True):
            st.markdown(f"**Stage {stage['stage_number']}: {stage['output_table']}**")
            st.caption(stage['what_happens'])
            if stage['keys']:
                st.caption(f"Alias / Scope: {stage['keys']}")
            if not stage['matched_steps']:
                st.caption("No specific AI transformation steps returned for this stage — review deterministically.")
            for step in stage['matched_steps']:
                st.markdown(f"- **{getattr(step, 'title', 'Step')}**: {getattr(step, 'action', '')}")
                manual_val = getattr(step, 'manual_validation', None)
                if manual_val:
                    st.warning(manual_val)


def render_structured_modeler_flow(modeler_plan_df: Optional[pd.DataFrame], generic_logic: Optional[Dict[str, Any]] = None):
    st.markdown("### 🗺️ Structured Pyramid Modeler Flow")
    st.caption("End-to-end migration pipeline mapping SSRS report logic directly to Pyramid Analytics architecture layers.")

    if modeler_plan_df is None or modeler_plan_df.empty:
        st.info("No structured modeler flow items were generated for this report.")
        return

    total_items = len(modeler_plan_df)
    rev_counts = modeler_plan_df['review_level'].value_counts()
    auto_count = rev_counts.get('automatic', 0)
    rec_count = rev_counts.get('recommended', 0)
    manual_count = rev_counts.get('manual_review', 0)

    auto_pct = (auto_count / total_items * 100) if total_items else 0
    rec_pct = (rec_count / total_items * 100) if total_items else 0
    manual_pct = (manual_count / total_items * 100) if total_items else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        with st.container(border=True):
            st.metric("Total Elements", total_items)
    with c2:
        with st.container(border=True):
            st.metric("Automatic", f"{auto_count} ({auto_pct:.0f}%)")
    with c3:
        with st.container(border=True):
            st.metric("Recommended", f"{rec_count} ({rec_pct:.0f}%)")
    with c4:
        with st.container(border=True):
            st.metric("Manual Review", f"{manual_count} ({manual_pct:.0f}%)")

    fc1, fc2 = st.columns(2)
    categories = ['All'] + sorted(modeler_plan_df['category'].dropna().unique().tolist())
    review_levels = ['All'] + sorted(modeler_plan_df['review_level'].dropna().unique().tolist())

    file_tag = re.sub(r'[^A-Za-z0-9_]+', '_', str(generic_logic.get('report_name', 'report') if generic_logic else 'report'))
    with fc1:
        sel_cat = st.selectbox("Filter by Category", categories, key=f"mod_cat_{file_tag}")
    with fc2:
        sel_rev = st.selectbox("Filter by Review Level", review_levels, key=f"mod_rev_{file_tag}")

    filtered_df = modeler_plan_df
    if sel_cat != 'All':
        filtered_df = filtered_df[filtered_df['category'] == sel_cat]
    if sel_rev != 'All':
        filtered_df = filtered_df[filtered_df['review_level'] == sel_rev]

    display_cols = ['stage', 'category', 'name', 'ssrs_logic', 'pyramid_target', 'pyramid_action', 'review_level', 'confidence', 'reason']
    avail_cols = [c for c in display_cols if c in filtered_df.columns]
    st.dataframe(filtered_df[avail_cols], use_container_width=True, hide_index=True)

    if generic_logic:
        with st.expander("🔍 View Generic Migration Representation (JSON)", expanded=False):
            st.caption("Language-neutral intermediate representation extracted from SSRS RDL before Pyramid targeting.")
            st.json(generic_logic)


def render_semantic_model_inspector(semantic_model: Optional[Dict[str, Any]], generic_logic: Optional[Dict[str, Any]] = None):
    st.markdown("### 🏛️ Pyramid Semantic Model (Model Pro)")
    st.caption("Pyramid Analytics semantic layer representation: Model Tables, Attributes, Measures, Relationships, and Hierarchies.")

    if not semantic_model or not any(semantic_model.get(k) for k in ('tables', 'columns', 'measures', 'relationships', 'hierarchies')):
        st.info("No semantic model elements were generated for this report.")
        return

    tables = semantic_model.get('tables', [])
    columns = semantic_model.get('columns', [])
    measures = semantic_model.get('measures', [])
    relationships = semantic_model.get('relationships', [])
    hierarchies = semantic_model.get('hierarchies', [])

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        with st.container(border=True):
            st.metric("Model Tables", len(tables))
    with c2:
        with st.container(border=True):
            st.metric("Attributes / Cols", len(columns))
    with c3:
        with st.container(border=True):
            st.metric("Measures", len(measures))
    with c4:
        with st.container(border=True):
            st.metric("Relationships", len(relationships))
    with c5:
        with st.container(border=True):
            st.metric("Hierarchies", len(hierarchies))

    sem_tabs = st.tabs([
        "📋 Model Tables",
        "🏷️ Attributes & Columns",
        "📊 Measures & PQL Formulas",
        "🔗 Relationships",
        "🌲 Hierarchies",
    ])

    with sem_tabs[0]:
        st.markdown("#### Model Tables & Data Flow Staging")
        if tables:
            st.dataframe(pd.DataFrame(tables), use_container_width=True, hide_index=True)
        else:
            st.info("No tables detected.")

    with sem_tabs[1]:
        st.markdown("#### Column Attributes & Calculated Columns")
        if columns:
            st.dataframe(pd.DataFrame(columns), use_container_width=True, hide_index=True)
        else:
            st.info("No columns detected.")

    with sem_tabs[2]:
        st.markdown("#### Pyramid Measures & PQL Expressions")
        if measures:
            st.dataframe(pd.DataFrame(measures), use_container_width=True, hide_index=True)
        else:
            st.info("No measures detected.")

    with sem_tabs[3]:
        st.markdown("#### Relationships & Join Implementation Strategy")
        if relationships:
            st.dataframe(pd.DataFrame(relationships), use_container_width=True, hide_index=True)
        else:
            st.info("No relationships detected.")

    with sem_tabs[4]:
        st.markdown("#### Drill-down & Attribute Hierarchies")
        if hierarchies:
            disp_h = []
            for h in hierarchies:
                hd = dict(h)
                if isinstance(hd.get('levels'), list):
                    hd['levels'] = ' -> '.join(str(x) for x in hd['levels'])
                disp_h.append(hd)
            st.dataframe(pd.DataFrame(disp_h), use_container_width=True, hide_index=True)
        else:
            st.info("No hierarchies detected.")

    with st.expander("🔍 View Full Semantic Model (JSON)", expanded=False):
        st.json(semantic_model)


def render_ai_build_plan(report: Dict[str, Any], generic_logic: Dict[str, Any], pyramid_map: Dict[str, Any]):
    st.markdown(
        "Generates a **hybrid rule-based + AI** build plan: deterministic RDL idioms, CTE pipelines, and structural "
        "parameters are translated deterministically (no AI, no cost), and one explicit AI call organizes the "
        "remaining judgment calls — join risk, novel expressions, and worksheet/dashboard suggestions for **Pyramid**."
    )

    st.markdown("#### Rule-Based Translations (no AI, zero cost)")
    col_rb1, col_rb2 = st.columns(2)
    with col_rb1:
        st.markdown(f"**DataSources & DataSets:** {len(report.get('data_sources', []))} source(s), {len(report.get('data_sets', []))} dataset(s) mapped deterministically.")
        ctes = [t['table_name'] for t in generic_logic.get('physical_tables', []) if t.get('is_cte')]
        if ctes:
            st.markdown(f"**CTEs Extracted:** `{', '.join(ctes)}` → mapped to sequential Data Flow Prep nodes.")
        else:
            st.markdown("**CTEs:** None detected (direct query execution).")
        st.markdown(f"**Report Parameters:** {len(report.get('parameters', []))} parameters mapped to Present Slicers.")
    with col_rb2:
        meas_count = len(pyramid_map.get('semantic_measures', []))
        st.markdown(f"**PQL Aggregate Measures:** {meas_count} measure(s) generated via deterministic syntax rules.")
        vis_count = len(report.get('tablixes', [])) + len(report.get('charts', []))
        st.markdown(f"**Discover Visuals:** {vis_count} Tablix / Chart visual(s) mapped.")

    st.markdown("---")

    client = get_openai_client()
    if client is None:
        with st.expander("🔑 OpenAI API Key Required", expanded=True):
            st.warning(
                "OpenAI is not configured. Add `OPENAI_API_KEY` to `.streamlit/secrets.toml` "
                "or enter your API key below for this session:"
            )
            key_input = st.text_input(
                "OpenAI API Key",
                type="password",
                key="ssrs_openai_key_input",
                placeholder="sk-proj-...",
            )
            if key_input:
                st.session_state['openai_api_key'] = key_input.strip()
                st.success("API key saved for this session!")
                st.rerun()
        return

    report_key = re.sub(r'[^A-Za-z0-9_]+', '_', report.get('report_name', 'report'))
    cache = st.session_state.setdefault('ssrs_pyramid_ai_build_plan_cache', {})

    generate_clicked = st.button(
        "Generate AI Build Plan" if report_key not in cache else "Regenerate AI Build Plan",
        key=f"ssrs_ai_build_plan_btn_{report_key}",
    )

    if generate_clicked:
        payload = build_ai_build_plan_payload(report, generic_logic, pyramid_map)
        with st.spinner("Calling OpenAI (one request for this report)..."):
            plan, error = generate_ai_build_plan(client, "gpt-4o-mini", payload)
        cache[report_key] = (plan, error)

    cached = cache.get(report_key)
    if cached is None:
        st.info("Click the button above to generate the AI-assisted portion of the build plan.")
        return

    plan, error = cached
    if error:
        st.error(error)
        return

    if plan is None:
        st.info("Click the button above to generate the AI-assisted portion of the build plan.")
        return

    summary = getattr(plan, 'report_summary', None)
    st.markdown("#### AI-Assisted Plan")
    if summary:
        c1, c2, c3 = st.columns(3)
        c1.metric("Complexity", getattr(summary, 'complexity_level', 'Medium'))
        c2.metric("Overall confidence", f"{getattr(summary, 'overall_confidence', 0)}%")
        manual_items = getattr(plan, 'manual_review_items', []) or []
        c3.metric("Manual review items", len(manual_items))
        st.write(f"**Likely purpose:** {getattr(summary, 'likely_purpose', '')}")
        st.write(f"**Report type:** {getattr(summary, 'report_type', '')}")

    tabs = st.tabs(["Build Order", "Implementation Steps", "Data Prep", "Calculations", "Filters/Joins", "Pages", "Validation", "Manual Review"])
    with tabs[0]:
        render_pyramid_prep_build_order(generic_logic, plan)

    with tabs[1]:
        imp_steps = getattr(plan, 'implementation_plan', []) or []
        if not imp_steps:
            st.info("No implementation steps returned.")
        for step in sorted(imp_steps, key=lambda s: getattr(s, 'step_number', 1)):
            _render_plan_step(step)

    with tabs[2]:
        data_steps = getattr(plan, 'data_preparation', []) or []
        if not data_steps:
            st.info("No data preparation steps returned.")
        for step in data_steps:
            _render_plan_step(step)

    with tabs[3]:
        calc_steps = getattr(plan, 'calculation_plan', []) or []
        if not calc_steps:
            st.info("No calculation translations returned.")
        for step in calc_steps:
            _render_plan_step(step)

    with tabs[4]:
        filt_steps = getattr(plan, 'filter_parameter_plan', []) or []
        if not filt_steps:
            st.info("No filter/join items returned.")
        for step in filt_steps:
            _render_plan_step(step)

    with tabs[5]:
        page_steps = getattr(plan, 'page_recommendations', []) or []
        if not page_steps:
            st.info("No page/visual recommendations returned.")
        for step in page_steps:
            _render_plan_step(step)

    with tabs[6]:
        val_checks = getattr(plan, 'validation_checks', []) or []
        if not val_checks:
            st.info("No validation checks returned.")
        for check in val_checks:
            st.markdown(f"- {check}")

    with tabs[7]:
        man_items = getattr(plan, 'manual_review_items', []) or []
        if not man_items:
            st.info("No manual review items flagged.")
        for item in man_items:
            st.warning(item)


def render_sql_view_recommendation_inspector(report: Dict[str, Any], generic_logic: Dict[str, Any]):
    st.markdown("### 🗄️ SQL View / Query View Recommendation")

    dsets = report.get("data_sets", [])
    total_ctes = []
    total_tables = []
    total_joins = []
    raw_query = ""
    for ds in dsets:
        decomp = ds.get("sql_decomposition", {})
        total_ctes.extend(decomp.get("ctes", []))
        total_tables.extend(decomp.get("tables", []))
        total_joins.extend(decomp.get("joins", []))
        if ds.get("command_text") and not raw_query:
            raw_query = ds.get("command_text")

    is_recommended = len(total_ctes) >= 2 or len(total_tables) >= 3
    is_optional = len(total_ctes) == 1 or len(total_tables) == 2 or len(total_joins) >= 1
    need_str = "Recommended" if is_recommended else ("Optional" if is_optional else "Not Required")

    clean_name = re.sub(r'[^A-Za-z0-9_]+', '_', report.get('report_name', 'report')).lower().strip('_')
    view_name = f"vw_pyramid_{clean_name}"

    st.markdown(f"**SQL View Recommended:** `{need_str}`")

    reasons = []
    if total_ctes:
        reasons.append(f"Contains {len(total_ctes)} Common Table Expression(s) (CTEs): `{', '.join(total_ctes[:5])}`.")
    if len(total_tables) >= 3:
        reasons.append(f"Involves {len(total_tables)} physical tables with {len(total_joins)} relational join(s).")
    if is_recommended:
        reasons.append("Encapsulating multi-CTE and multi-join logic into a database SQL View significantly streamlines Pyramid Modeler Data Flow.")
        reasons.append("Allows database engine optimizer to materialize and execute CTE execution plans efficiently.")
    elif is_optional:
        reasons.append("Single join or lightweight query; can be managed either via SQL View or direct Pyramid Modeler Join nodes.")
    else:
        reasons.append("Simple direct table query; direct Pyramid Data Model table connection is optimal.")

    for r in reasons:
        st.write(f"- {r}")

    st.write(f"**Recommended SQL View Name:** `{view_name}`")
    st.write("**Preferred Implementation:** Create this view in the source database catalog, then point Pyramid's data source directly to the view.")
    st.write("**Alternative Implementation:** Use Pyramid's Custom SQL Query node in Modeler Data Flow if database schema write permissions are restricted.")

    st.markdown("#### SQL CTE Outline & Query Draft")
    if total_ctes:
        outline_lines = ["-- Recommended SQL View Draft Definition", f"CREATE OR ALTER VIEW dbo.{view_name} AS"]
        outline_lines.append("WITH " + ",\n".join(f"  {c} AS (\n    -- [Stage: {c}]\n    SELECT ...\n  )" for c in total_ctes))
        outline_lines.append(f"SELECT *\nFROM {total_ctes[-1]};")
        st.code("\n".join(outline_lines), language="sql")
    elif raw_query:
        st.code(raw_query[:1500] + ("\n... [truncated]" if len(raw_query) > 1500 else ""), language="sql")
    else:
        st.code(f"CREATE OR ALTER VIEW dbo.{view_name} AS\nSELECT *\nFROM {total_tables[0]['table_name'] if total_tables else 'SourceTable'};", language="sql")

    st.markdown(
        """
        <div style="background: rgba(127,127,127,0.06); padding: 12px 16px; border-radius: 8px; border-left: 4px solid #2563EB;">
          <strong>SQL Draft Status:</strong> Generated by deterministic AST parser.<br>
          <strong>Ready to run:</strong> Review required — verify schema qualifiers, WHERE clause parameters, and performance indexes.
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_data_prep_inspector(report: Dict[str, Any], generic_logic: Dict[str, Any]):
    st.markdown("### 🧹 Pyramid Data Flow / Data Model (Data Prep)")
    dsets = report.get("data_sets", [])
    total_ctes = [t for t in generic_logic.get("physical_tables", []) if t.get("is_cte")]
    clean_name = re.sub(r'[^A-Za-z0-9_]+', '_', report.get('report_name', 'report')).lower().strip('_')
    view_name = f"vw_pyramid_{clean_name}"

    if len(total_ctes) >= 2 or len(generic_logic.get("physical_tables", [])) >= 3:
        st.write("Pyramid Data Flow / Model should be **light cleanup only**.")
        st.info(f"Because a SQL View is recommended, Pyramid Data Flow should not rebuild the multi-stage logic from scratch. Use it after `{view_name}` is created and validated.")
        steps = [
            f"Connect Pyramid data source to database view `{view_name}`",
            "Confirm preview data loads in Modeler canvas",
            "Verify auto-detected data types (dates, decimals, integers)",
            "Rename technical database fields to business-friendly display names",
            "Remove unnecessary surrogate keys or internal ID helper columns",
            "Verify row count consistency with source SSRS output",
            "Publish and deploy data source to Pyramid IMDB or direct query",
        ]
    else:
        st.write("Pyramid Data Flow should **recreate the SSRS staging flow directly**.")
        steps = [
            f"Load source table(s) from connection `{generic_logic.get('sources', [{}])[0].get('source_id', 'Source') if generic_logic.get('sources') else 'Source'}`",
            "Create staging Data Flow nodes on canvas",
            "Apply dataset row filters matching SSRS WHERE conditions",
            "Configure row-level calculated columns in Data Flow",
            "Join related tables using Modeler Join nodes",
            "Create final output Semantic Model target",
        ]

    rep_tag = re.sub(r'[^A-Za-z0-9_]+', '_', report.get('report_name', 'report'))
    for idx, step in enumerate(steps, start=1):
        st.checkbox(step, value=False, key=f"ssrs_prep_chk_{rep_tag}_{idx}")


def render_data_model_and_calc_inspector(report: Dict[str, Any], generic_logic: Dict[str, Any], pyramid_map: Dict[str, Any]):
    st.markdown("### 📊 Data Model & Calculated Fields")

    st.markdown("#### Recommended Data Model Architecture")
    dsets = report.get("data_sets", [])
    total_ctes = [t for t in generic_logic.get("physical_tables", []) if t.get("is_cte")]
    clean_name = re.sub(r'[^A-Za-z0-9_]+', '_', report.get('report_name', 'report')).lower().strip('_')
    view_name = f"vw_pyramid_{clean_name}"

    if len(total_ctes) >= 2:
        st.write("**Model Architecture:** Simple Flat Model from View.")
        st.info(f"Use main model table `{view_name}`. Pre-aggregating multi-pass CTE logic in the view eliminates cross-table join fanout.")
    elif len(generic_logic.get("physical_tables", [])) > 1:
        st.write("**Model Architecture:** Star / Snowflake Schema Model.")
        st.info("Multiple tables detected. Configure dimension-to-fact relationships in Pyramid Model Pro canvas with verified foreign keys.")
    else:
        st.write("**Model Architecture:** Single Table Direct Model.")
        st.info(f"Direct connection to table `{generic_logic.get('physical_tables', [{}])[0].get('table_name', 'Table') if generic_logic.get('physical_tables') else 'Table'}`.")

    st.markdown("#### Calculated Fields & Expressions")
    calcs = generic_logic.get("calculations", [])
    if calcs:
        c_df = pd.DataFrame(calcs)[["expression", "type", "pql_translation", "proposed_layer", "triage_level", "confidence"]]
        c_df["confidence"] = c_df["confidence"].apply(lambda v: f"{int(v*100) if v <= 1 else int(v)}%")
        st.dataframe(c_df, use_container_width=True, hide_index=True)
    else:
        st.info("No calculated expressions detected on this report.")


def render_visuals_inspector(report: Dict[str, Any], generic_logic: Dict[str, Any]):
    st.markdown("### 📈 Visuals & Dashboard Inspector")
    st.markdown(f"**Recommended Pyramid Present Dashboard:** `Dashboard: {report.get('report_name', 'Report')}`")

    visuals = generic_logic.get("visuals", [])
    if not visuals:
        st.info("No visual components (Tablix or Chart) detected.")
        return

    for idx, vis in enumerate(visuals, start=1):
        with st.container(border=True):
            st.markdown(f"**Visual {idx}: {vis['name']}** — `{vis['pyramid_discover_type']}`")
            v_type = vis.get("type", "Grid")
            if "Matrix" in v_type:
                st.write("**Layout Structure:** Matrix / Crosstab")
                st.write(f"- Row Group Count: {vis.get('row_groups_count', 0)}")
                st.write(f"- Column Group Count: {vis.get('col_groups_count', 0)}")
            elif "Chart" in v_type:
                st.write(f"**Chart Visualization:** {v_type}")
            else:
                st.write("**Layout Structure:** Text / Tabular Grid")

            if vis.get("has_page_break"):
                st.info("⚡ Contains PageBreak: Configure as Pyramid Present Publication with Page-by-Slice Bursting.")


def render_validation_checklist(report: Dict[str, Any], generic_logic: Dict[str, Any]):
    st.markdown("### ✅ Migration Validation Checklist")
    st.caption("Complete these verification steps before promoting the migrated Pyramid report to production.")

    checks = [
        ("Data Source Connection", "Verify database connection credentials and gateway connectivity in Pyramid Modeler."),
        ("Row Count Reconciliation", "Execute SELECT COUNT(*) on source DataSet query and confirm exact match in Pyramid Data Model."),
        ("Measure Calculations", "Reconcile aggregate measure totals (Sum, Count, Avg) against SSRS report output."),
        ("Matrix & Grouping Subtotals", "Validate row and column hierarchy groupings and subtotal calculations."),
        ("Parameter Slicers", "Test all dynamic parameters and multi-select dropdown slicers in Pyramid Present."),
        ("Conditional Formatting", "Verify alerting colors and background threshold shading rules in Discover."),
    ]

    rep_tag = re.sub(r'[^A-Za-z0-9_]+', '_', report.get('report_name', 'report'))
    chk_keys = [f"ssrs_val_chk_{rep_tag}_{i}" for i in range(len(checks))]
    completed = sum(1 for k in chk_keys if st.session_state.get(k, False))
    total = len(checks)

    status_text = "Complete" if completed == total else ("In progress" if completed > 0 else "Not started")
    status_color = "#059669" if completed == total else ("#D97706" if completed > 0 else "#6B7280")

    st.markdown(
        f"""
        <div style="background: rgba(127,127,127,0.08); padding: 14px 18px; border-radius: 8px; margin-bottom: 16px; display: flex; justify-content: space-between; align-items: center;">
            <div style="font-weight: 700; font-size: 1.05rem;">{completed} / {total} checks completed</div>
            <div style="font-weight: 700; color: {status_color}; font-size: 0.95rem;">Status: {status_text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    for i, (title, desc) in enumerate(checks):
        st.checkbox(f"**{title}**: {desc}", value=False, key=chk_keys[i])


def render_table_db_mapping(parsed_results: List[Dict[str, Any]], selected_report: Dict[str, Any]):
    st.markdown("### 🗄️ Table → Database Mapping")
    st.caption("Complete lineage mapping of physical source tables, datasets, and target Pyramid Modeler nodes.")

    rows = []
    seen = set()
    for item in parsed_results:
        gen = build_generic_migration_representation(item)
        src_map = {s["source_id"]: s for s in gen.get("sources", [])}
        for tbl in gen.get("physical_tables", []):
            t_name = tbl["table_name"]
            ds_name = tbl.get("dataset_name", "")
            key = (item["report_name"], ds_name, t_name)
            if key not in seen:
                seen.add(key)
                src_info = src_map.get(ds_name, {})
                if not src_info and gen.get("sources"):
                    src_info = gen["sources"][0]
                rows.append({
                    "Table / Entity Name": t_name,
                    "DataSet Name": ds_name,
                    "Report Name": item["report_name"],
                    "Data Provider": src_info.get("source_type", "Relational SQL"),
                    "Connection String": src_info.get("connection_string", "Shared Connection") or "Shared Connection",
                    "Target Modeler Node": f"Prep_{t_name.replace('.', '_')}",
                    "Status": "Direct Mapped" if src_info.get("source_type") else "Needs Review",
                    "Recommendations": "Modeler CTE Stage Node" if tbl.get("is_cte") else "Model Table Select Node",
                })

    if not rows:
        st.info("No physical tables detected across reports.")
        return

    df = pd.DataFrame(rows)
    m1, m2, m3 = st.columns(3)
    with m1:
        st.metric("Total Unique Tables", len(df["Table / Entity Name"].unique()))
    with m2:
        mapped_count = len(df[df["Status"] == "Direct Mapped"])
        st.metric("Direct Mapped", mapped_count)
    with m3:
        needs_rev_count = len(df[df["Status"] == "Needs Review"])
        st.metric("Needs Review", needs_rev_count)

    st.dataframe(df, use_container_width=True, hide_index=True)


def render_pyramid_build_plan_tab(report: Dict[str, Any], generic: Dict[str, Any], pyramid: Dict[str, Any], excel_bytes: bytes):
    st.subheader(f"Pyramid Analytics Build Plan & Target Architecture: {report['report_name']}")

    # Informational message placed IMMEDIATELY BEFORE the download button
    st.info("The full evidence workbook is available from the Download Full Excel Report button.")

    # Excel Download Button
    st.download_button(
        label="Download This Report (Excel)",
        data=excel_bytes,
        file_name=f"Pyramid_Migration_Report_{report['report_name']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"dl_excel_{report['report_name']}",
        type="primary",
        use_container_width=True,
    )

    st.markdown("---")

    pyramid_subtab_labels = [
        "Structured Modeler Flow",
        "Pyramid Semantic Model",
        "AI Build Plan",
        "SQL View / Query View",
        "Pyramid Data Flow / Data Model",
        "Data Model & Calculated Fields",
        "Visuals",
        "Validation",
    ]

    report_tag = re.sub(r'[^A-Za-z0-9_]+', '_', report.get('report_name', 'report'))
    active_subtab_key = f"active_pyramid_subtab_{report_tag}"
    if active_subtab_key not in st.session_state:
        st.session_state[active_subtab_key] = pyramid_subtab_labels[0]

    subtabs = st.tabs(pyramid_subtab_labels)

    with subtabs[0]:
        render_structured_modeler_flow(pyramid.get("modeler_plan_df"), generic)

    with subtabs[1]:
        render_semantic_model_inspector(generic.get("semantic_model"), generic)

    with subtabs[2]:
        render_ai_build_plan(report, generic, pyramid)

    with subtabs[3]:
        render_sql_view_recommendation_inspector(report, generic)

    with subtabs[4]:
        render_data_prep_inspector(report, generic)

    with subtabs[5]:
        render_data_model_and_calc_inspector(report, generic, pyramid)

    with subtabs[6]:
        render_visuals_inspector(report, generic)

    with subtabs[7]:
        render_validation_checklist(report, generic)


# ==============================================================================
# 10. FILE UPLOADER & ARCHIVE HANDLING
# ==============================================================================

def load_uploaded_files(uploaded_files, return_errors: bool = False):
    """
    Ingest uploaded files: supports individual .rdl, .xml, and .zip archives.
    Extracts nested files and returns parsed AST dictionaries (and optional errors).
    """
    parsed_items = []
    errors = []
    
    for f in uploaded_files:
        if f is None:
            continue
        fname = getattr(f, "name", "uploaded_file")
        try:
            fbytes = f.read()
        except Exception as e:
            errors.append(f"{fname}: Failed to read file data ({e})")
            continue

        if fname.lower().endswith(".zip"):
            try:
                with zipfile.ZipFile(io.BytesIO(fbytes), "r") as z:
                    found_any = False
                    for member in z.namelist():
                        if member.startswith("__MACOSX") or member.endswith("/"):
                            continue
                        if member.lower().endswith((".rdl", ".xml", ".rsd")):
                            found_any = True
                            try:
                                content = z.read(member).decode("utf-8", errors="ignore")
                                parsed = parse_ssrs_rdl(content, filename=Path(member).name)
                                if "error" in parsed:
                                    errors.append(f"{member}: {parsed['error']}")
                                elif not parsed.get("data_sources") and not parsed.get("data_sets") and not parsed.get("tablixes"):
                                    errors.append(f"{member}: Does not appear to be a valid SSRS report definition.")
                                else:
                                    parsed_items.append(parsed)
                            except Exception as ex:
                                errors.append(f"{member}: {str(ex)}")
                    if not found_any:
                        errors.append(f"{fname}: No .rdl or .xml report files found in ZIP archive.")
            except Exception as e:
                errors.append(f"{fname}: Failed to open ZIP archive ({str(e)})")

        elif fname.lower().endswith((".rdl", ".xml")):
            try:
                content = fbytes.decode("utf-8", errors="ignore")
                parsed = parse_ssrs_rdl(content, filename=fname)
                if "error" in parsed:
                    errors.append(f"{fname}: {parsed['error']}")
                elif not parsed.get("data_sources") and not parsed.get("data_sets") and not parsed.get("tablixes"):
                    errors.append(f"{fname}: Does not appear to be a valid SSRS report definition.")
                else:
                    parsed_items.append(parsed)
            except Exception as e:
                errors.append(f"{fname}: {str(e)}")
        else:
            errors.append(f"{fname}: Unsupported file type. Please upload .rdl, .xml, or .zip.")

    if return_errors:
        return parsed_items, errors
    return parsed_items


# ==============================================================================
# 11. STREAMLIT APP MAIN
# ==============================================================================

def main():
    # Inject KashMap Theme & Card Styles
    st.markdown(
        """
        <style>
        .kash-hero {
            padding: 28px 32px;
            background: linear-gradient(135deg, rgba(37, 99, 235, 0.08) 0%, rgba(255, 90, 31, 0.08) 100%);
            border: 1px solid rgba(127, 127, 127, 0.2);
            border-radius: 12px;
            margin-bottom: 24px;
        }
        .kash-title {
            font-size: 2.2rem;
            font-weight: 900;
            letter-spacing: -0.02em;
            margin-bottom: 6px;
        }
        .kash-title-kash { color: #1F2933; }
        .kash-title-map { color: #FF5A1F; }
        .kash-subtitle {
            font-size: 1.05rem;
            color: #4B5563;
            font-weight: 500;
        }
        .kash-subtitle-accent {
            color: #2563EB;
            font-weight: 700;
        }
        .section-label {
            font-size: 0.82rem;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            color: #6B7280;
            margin: 18px 0 10px 0;
        }
        .card-title { font-size: 1rem; font-weight: 800; margin-bottom: 4px; }
        .card-subtitle { font-size: 0.86rem; opacity: 0.75; margin-bottom: 12px; }
        .badge-required {
            width: fit-content;
            margin-left: auto;
            font-size: 0.72rem;
            font-weight: 700;
            color: #B42318;
            background: #FFF1F0;
            border: 1px solid #FFE1DF;
            border-radius: 999px;
            padding: 4px 10px;
            text-align: center;
        }
        .file-count-box {
            background: rgba(127, 127, 127, 0.08);
            border: 1px solid rgba(127, 127, 127, 0.2);
            border-radius: 8px;
            padding: 8px 12px;
            margin-top: 10px;
            font-size: 0.88rem;
            font-weight: 600;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="kash-hero">
          <div class="kash-title"><span class="kash-title-kash">Kash</span><span class="kash-title-map">Map</span></div>
          <div class="kash-subtitle">Break Down <span class="kash-subtitle-accent">SSRS</span> Complexity for Faster <span class="kash-subtitle-accent">Pyramid Analytics</span> Migration</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.expander("How to use KashMap", expanded=True):
        st.markdown(
            """
            1. Upload one or more SSRS `.rdl` files or supported XML/ZIP input.
            2. Click **Analyze Reports**.
            3. Review the overview first, then open the guided Pyramid build plan and Structured Modeler Flow.
            4. Use the detailed tabs to inspect DataSources, DataSets, Queries & SQL, Joins, Filters, Parameters, and Calculations.
            """
        )

    st.markdown('<div class="section-label">Upload Files</div>', unsafe_allow_html=True)
    upload_col, side_col = st.columns([1, 1], gap="large")

    with upload_col:
        upload_card = st.container(border=True)
        with upload_card:
            title_col, badge_col = st.columns([0.82, 0.18])
            with title_col:
                st.markdown(
                    """
                    <div class="card-title">Upload SSRS Definitions</div>
                    <div class="card-subtitle">Upload multiple .rdl/.xml files or one ZIP folder containing report definitions.</div>
                    """,
                    unsafe_allow_html=True,
                )
            with badge_col:
                st.markdown('<div class="badge-required">Required</div>', unsafe_allow_html=True)

            mode = st.radio("Report input type", ["Individual Report Files", "ZIP File"], horizontal=True, key="ssrs_input_type")

            if mode == "Individual Report Files":
                uploaded_rdl_files = st.file_uploader(
                    "Upload one or more .rdl or .xml files",
                    type=["rdl", "xml"],
                    accept_multiple_files=True,
                    key="rdl_files_uploader",
                )
                uploaded_zip = None
                if uploaded_rdl_files:
                    total_size = sum(getattr(f, "size", 0) for f in uploaded_rdl_files)
                    st.markdown(
                        f'<div class="file-count-box">{len(uploaded_rdl_files)} report file(s) uploaded · {total_size / 1024:.1f} KB total</div>',
                        unsafe_allow_html=True,
                    )
            else:
                uploaded_zip = st.file_uploader(
                    "Upload ZIP file containing .rdl/.xml reports",
                    type=["zip"],
                    key="ssrs_zip_uploader"
                )
                uploaded_rdl_files = []
                if uploaded_zip:
                    st.markdown(
                        f'<div class="file-count-box">ZIP uploaded · {uploaded_zip.size / 1024:.1f} KB</div>',
                        unsafe_allow_html=True,
                    )

    with side_col:
        output_card = st.container(border=True)
        with output_card:
            st.markdown(
                """
                <div class="card-title">Output Settings</div>
                <div class="card-subtitle">Name of the generated Excel migration workbook.</div>
                """,
                unsafe_allow_html=True,
            )
            output_name = st.text_input("Output file name", value="SSRS_to_Pyramid_Migration_Output.xlsx", key="output_file_name")

    # Developer regression testing tools (sidebar only, collapsed, never runs automatically)
    with st.sidebar.expander("🛠️ Developer / Testing Tools", expanded=False):
        st.caption("Explicit developer testing only (never runs automatically):")
        dev_submit = st.button("Run Bundled Test Dataset (7 Reports)", key="btn_dev_test")

    submit = st.button("Analyze Reports", type="primary", use_container_width=True)

    # Manage upload state and session invalidation
    current_files = uploaded_rdl_files if mode == "Individual Report Files" else ([uploaded_zip] if uploaded_zip else [])
    current_upload_keys = [getattr(f, "name", "") for f in current_files if f is not None]
    stored_upload_keys = st.session_state.get("ssrs_pyramid_upload_keys", [])

    # Invalidate session state if no files uploaded and dev button not clicked
    if not current_upload_keys and not dev_submit:
        st.session_state.pop("ssrs_pyramid_analysis", None)
        st.session_state.pop("ssrs_pyramid_upload_keys", None)
    elif current_upload_keys != stored_upload_keys and not dev_submit and not submit:
        # User changed files in uploader without submitting yet
        st.session_state.pop("ssrs_pyramid_analysis", None)
        st.session_state.pop("ssrs_pyramid_upload_keys", None)

    if submit:
        if not current_upload_keys:
            st.warning("Please upload at least one SSRS .rdl or .zip file before analyzing.")
        else:
            with st.spinner("Analyzing SSRS report definitions..."):
                parsed_items, errors = load_uploaded_files(current_files, return_errors=True)
            if not parsed_items:
                st.error("Unable to parse the uploaded SSRS definition. Please upload a valid .rdl, .xml, or supported ZIP file.")
                if errors:
                    with st.expander("Parsing Details", expanded=False):
                        for err in errors:
                            st.write(err)
            else:
                if errors:
                    for err in errors:
                        st.warning(err)
                st.session_state["ssrs_pyramid_upload_keys"] = current_upload_keys
                st.session_state["ssrs_pyramid_analysis"] = {
                    "parsed_results": parsed_items,
                    "file_names": current_upload_keys,
                    "output_name": output_name
                }

    elif dev_submit:
        local_dataset_dir = Path(__file__).resolve().parent.parent.parent / "test_reports" / "ssrs"
        local_rdl_files = sorted(list(local_dataset_dir.glob("*.rdl"))) if local_dataset_dir.exists() else []
        if local_rdl_files:
            with st.spinner("Running bundled test dataset (7 reports)..."):
                parsed_items = []
                for rdl_f in local_rdl_files:
                    with open(rdl_f, "r", encoding="utf-8", errors="ignore") as f:
                        c = f.read()
                    parsed_items.append(parse_ssrs_rdl(c, filename=rdl_f.name))
            st.session_state["ssrs_pyramid_upload_keys"] = [f.name for f in local_rdl_files]
            st.session_state["ssrs_pyramid_analysis"] = {
                "parsed_results": parsed_items,
                "file_names": [f.name for f in local_rdl_files],
                "output_name": "Bundled_Test_SSRS_Pyramid_Output.xlsx"
            }
        else:
            st.error("Bundled test dataset directory not found at test_reports/ssrs.")

    analysis_result = st.session_state.get("ssrs_pyramid_analysis")

    # If no analysis result, render initial state message and stop
    if not analysis_result:
        st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)
        st.info("No report loaded yet. Please upload an SSRS RDL file to begin.")
        return

    # Render results only when analysis is available
    parsed_results = analysis_result["parsed_results"]
    output_fn = analysis_result.get("output_name", "SSRS_to_Pyramid_Migration_Output.xlsx")
    if not output_fn.lower().endswith(".xlsx"):
        output_fn = f"{output_fn}.xlsx"

    st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)
    with st.expander("Uploaded files", expanded=True):
        for fn in analysis_result.get("file_names", []):
            st.markdown(f"- `{fn}`")

    st.markdown("---")

    # Render Portfolio KPIs
    render_kpi_cards(parsed_results)
    st.markdown("---")

    # Global Portfolio Excel Export Button
    excel_all_bytes = build_output_workbook(parsed_results)
    col_dl1, col_dl2 = st.columns([3, 1])
    with col_dl1:
        st.write(f"**Loaded {len(parsed_results)} report definition(s).** Select a report below to inspect detailed logic.")
    with col_dl2:
        st.download_button(
            label="Download Full Portfolio (Excel)",
            data=excel_all_bytes,
            file_name=output_fn,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_portfolio_excel"
        )

    # Report Selector
    report_names = [r["report_name"] for r in parsed_results]
    selected_name = st.selectbox("Choose an SSRS report to inspect", report_names, index=0)
    selected_report = next(r for r in parsed_results if r["report_name"] == selected_name)

    # Build generic and Pyramid mappings for selected report
    generic_logic = build_generic_migration_representation(selected_report)
    pyramid_map = map_generic_to_pyramid(generic_logic)
    single_report_excel = build_output_workbook([selected_report])

    # 12 Persistent Report Tabs
    tabs_list = [
        "Start Here",
        "Data Sources & DataSets",
        "Queries & SQL / CTEs",
        "Joins & Relationships",
        "Filters & Parameters",
        "Calculations & Expressions",
        "Grouping & Sorting",
        "Tables / Tablix",
        "Charts & Visuals",
        "Subreports & Drillthrough",
        "Pyramid Build Plan",
        "Table → Database Mapping",
    ]

    tabs = st.tabs(tabs_list)

    with tabs[0]:
        render_start_here_tab(selected_report, generic_logic, pyramid_map)

    with tabs[1]:
        render_datasources_datasets_tab(selected_report)

    with tabs[2]:
        render_queries_sql_tab(selected_report)

    with tabs[3]:
        render_joins_relationships_tab(selected_report, generic_logic)

    with tabs[4]:
        render_filters_parameters_tab(selected_report, generic_logic)

    with tabs[5]:
        render_calculations_expressions_tab(selected_report, generic_logic)

    with tabs[6]:
        render_grouping_sorting_tab(selected_report)

    with tabs[7]:
        render_tables_tablix_tab(selected_report)

    with tabs[8]:
        render_charts_visuals_tab(selected_report)

    with tabs[9]:
        render_subreports_drillthrough_tab(selected_report)

    with tabs[10]:
        render_pyramid_build_plan_tab(selected_report, generic_logic, pyramid_map, single_report_excel)

    with tabs[11]:
        render_table_db_mapping(parsed_results, selected_report)


if __name__ == "__main__":
    main()
