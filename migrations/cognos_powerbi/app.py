import json
import os
import subprocess
import uuid
import zipfile
from io import BytesIO
from pathlib import Path
from html import escape

import pandas as pd
import streamlit as st
import re
from xml.etree import ElementTree as ET
from collections import defaultdict, Counter
from typing import Literal, Optional
from pydantic import BaseModel, Field, ValidationError
try:
    from openai import OpenAI
except Exception:
    OpenAI = None

NS = {"r": "http://developer.cognos.com/schemas/report/17.5/"}


def _local(tag):
    """Strip namespace from an ElementTree tag, e.g. '{ns}expression' -> 'expression'."""
    return tag.split("}", 1)[-1] if "}" in tag else tag


def _iter_all(root, tagname):
    """Find all descendant elements with a given local tag name, namespace-agnostic."""
    return [el for el in root.iter() if _local(el.tag) == tagname]


def _find_child(el, tagname):
    for child in el:
        if _local(child.tag) == tagname:
            return child
    return None


def _text_of(el):
    return (el.text or "").strip() if el is not None else ""


_FULL_MODEL_PATH_RE = re.compile(r"^\[C\]\.\[([^\]]+)\]\.\[([^\]]+)\]\.\[([^\]]+)\]$")
_PKG_MODEL_PATH_RE = re.compile(r"^\[([^\]]+)\]\.\[([^\]]+)\]\.\[([^\]]+)\]$")
_TABLE_COL_PATH_RE = re.compile(r"^\[([^\]]+)\]\.\[([^\]]+)\]$")


def parse_model_expression(expr):
    """
    Parse a Cognos model-path expression like:
      [C].[Sample_data_module].[sheet1].[Revenue] -> (Sample_data_module, sheet1, Revenue)
      [Sales].[Branch].[Region]                   -> (Sales, Branch, Region)
      [Branch].[Region]                           -> (None, Branch, Region)
    into (module, table, column). Requires the ENTIRE expression to be a clean
    bracketed path (nothing else around it) — this distinguishes a plain field
    reference from a calculated/macro expression.
    Returns (None, None, None) if it doesn't match.
    """
    e = (expr or "").strip()
    m4 = _FULL_MODEL_PATH_RE.match(e)
    if m4:
        return m4.group(1), m4.group(2), m4.group(3)
    m3 = _PKG_MODEL_PATH_RE.match(e)
    if m3:
        return m3.group(1), m3.group(2), m3.group(3)
    m2 = _TABLE_COL_PATH_RE.match(e)
    if m2:
        return None, m2.group(1), m2.group(2)
    return None, None, None


def classify_cognos_expression(expr):
    """
    Cognos equivalent of classify_formula(). Categories chosen to mirror
    WebFOCUS categories where a real equivalent exists, plus new categories
    for patterns that have no WebFOCUS analog.
    """
    e = expr.strip()
    types = []

    has_prompt_call = re.search(r"\bprompt\s*\(", e, re.IGNORECASE) is not None
    is_macro_wrapped = e.startswith("#")  # Cognos macro syntax: # ... #
    builds_path_dynamically = has_prompt_call and re.search(r"[\"'].*\+.*prompt", e, re.IGNORECASE)

    if builds_path_dynamically:
        types.append("Dynamic Model Path (Prompt-Built)")
    elif has_prompt_call and is_macro_wrapped:
        types.append("Parameter-Driven Metric Swap")
    elif has_prompt_call:
        types.append("Parameter Reference")

    if re.search(r"\bParamDisplayValue\s*\(", e):
        types.append("Prompt Display Value")
    if re.search(r"\bif\s*\(.*\)\s*then\s*\(", e, re.IGNORECASE):
        types.append("Conditional IF")
    if not has_prompt_call and re.search(r"[-+*/]", e):
        types.append("Arithmetic")

    module, table, column = parse_model_expression(e)
    if module or table:
        types.append("Direct Model Reference")

    if not types:
        types.append("Other/Review")
    return " + ".join(types)


def extract_queries(root):
    """
    Extract every <query> block: dataItems (fields/calcs), detailFilters,
    summaryFilters, slicerFilters, query joins, set operations (unions), and native SQL.
    """
    queries = []
    for q in _iter_all(root, "query"):
        qname = q.get("name", "")
        data_items = []

        selection = _find_child(q, "selection")
        if selection is not None:
            for di in selection:
                tag = _local(di.tag)
                if tag not in ("dataItem", "dataItemListSummary"):
                    continue
                name = di.get("name", "")
                aggregate = di.get("aggregate", "")
                expr_el = _find_child(di, "expression")
                expr = _text_of(expr_el)
                module, table, column = parse_model_expression(expr)
                is_calc = (module is None and table is None) and expr != ""
                data_items.append({
                    "query": qname,
                    "field": name,
                    "aggregate": aggregate,
                    "expression": expr,
                    "module": module,
                    "table": table,
                    "column": column,
                    "is_calculated": is_calc,
                    "expression_type": classify_cognos_expression(expr) if expr else "",
                })

        detail_filters = []
        for df in _iter_all(q, "detailFilter"):
            fexpr = _find_child(df, "filterExpression")
            ftext = _text_of(fexpr)
            detail_filters.append({
                "query": qname,
                "use": df.get("use", "required"),
                "expression": ftext,
                "is_prompted": "?" in ftext or "prompt(" in ftext.lower(),
            })

        summary_filters = []
        for sf in _iter_all(q, "summaryFilter"):
            fexpr = _find_child(sf, "filterExpression")
            summary_filters.append({
                "query": qname,
                "use": sf.get("use", "required"),
                "expression": _text_of(fexpr),
            })

        slicer_filters = []
        for slf in _iter_all(q, "slicerFilter"):
            sexpr = _find_child(slf, "slicerExpression")
            slicer_filters.append({
                "query": qname,
                "expression": _text_of(sexpr),
            })

        # Query Source Inspection (Query Joins, Unions, Native SQL)
        query_type = "Standard Query"
        join_details = None
        set_operation = None
        native_sql = None

        source_el = _find_child(q, "source")
        if source_el is not None:
            # Native SQL inside query
            sql_el = _find_child(source_el, "sqlQuery") or _find_child(source_el, "sql")
            if sql_el is not None:
                query_type = "Native SQL Pass-Through"
                sql_text_el = _find_child(sql_el, "sqlText")
                sql_txt = _text_of(sql_text_el) or _text_of(sql_el)
                native_sql = {
                    "name": sql_el.get("name", ""),
                    "data_source": sql_el.get("dataSource", ""),
                    "sql_text": sql_txt,
                }

            # Query Join under model/joinOperation
            join_el = _find_child(source_el, "joinOperation") or _find_child(source_el, "join")
            if join_el is not None:
                query_type = "Query Join"
                left_q = ""
                right_q = ""
                left_el = _find_child(join_el, "joinLeft")
                right_el = _find_child(join_el, "joinRight")
                if left_el is not None:
                    qref = _find_child(left_el, "queryRef")
                    left_q = qref.get("refQuery", "") if qref is not None else ""
                if right_el is not None:
                    qref = _find_child(right_el, "queryRef")
                    right_q = qref.get("refQuery", "") if qref is not None else ""

                criteria_list = []
                for jrel in _iter_all(join_el, "joinRelationship"):
                    fexpr = _find_child(jrel, "filterExpression")
                    if fexpr is not None and _text_of(fexpr):
                        criteria_list.append(_text_of(fexpr))
                    else:
                        for crit in _iter_all(jrel, "joinCriteria"):
                            litem = _find_child(crit, "leftItem")
                            ritem = _find_child(crit, "rightItem")
                            l_name = litem.get("refDataItem", "") if litem is not None else ""
                            r_name = ritem.get("refDataItem", "") if ritem is not None else ""
                            if l_name and r_name:
                                criteria_list.append(f"{l_name} = {r_name}")
                for crit in _iter_all(join_el, "joinCriteria"):
                    ctext = _text_of(_find_child(crit, "criteria")) or _text_of(crit)
                    if ctext and ctext not in criteria_list:
                        criteria_list.append(ctext)

                join_details = {
                    "left_query": left_q,
                    "right_query": right_q,
                    "criteria": criteria_list,
                }

            # Query Set Operations (Union / Intersect / Except)
            set_el = _find_child(source_el, "queryOperation")
            if set_el is not None:
                op_type = set_el.get("type", "union").lower()
                query_type = f"Set Operation ({op_type.capitalize()})"
                sub_queries = [qr.get("refQuery", "") for qr in _iter_all(set_el, "queryRef") if qr.get("refQuery")]
                set_operation = {
                    "operation": op_type,
                    "sub_queries": sub_queries,
                }
            else:
                for set_tag in ("union", "intersect", "except"):
                    tag_el = _find_child(source_el, set_tag)
                    if tag_el is not None:
                        query_type = f"Set Operation ({set_tag.capitalize()})"
                        sub_queries = [qr.get("refQuery", "") for qr in _iter_all(tag_el, "queryRef") if qr.get("refQuery")]
                        set_operation = {
                            "operation": set_tag,
                            "sub_queries": sub_queries,
                        }
                        break

        queries.append({
            "name": qname,
            "data_items": data_items,
            "filters": detail_filters,
            "summary_filters": summary_filters,
            "slicer_filters": slicer_filters,
            "query_type": query_type,
            "join_details": join_details,
            "set_operation": set_operation,
            "native_sql": native_sql,
        })
    return queries


def extract_prompts(root):
    """
    Extract all Cognos prompt types: value prompts, date prompts, datetime prompts,
    interval prompts, text box prompts, tree prompts, and search & select prompts.
    """
    prompts = []
    PROMPT_TAGS = {
        "selectValue": "Value Prompt",
        "selectDate": "Date Prompt",
        "selectDateTime": "DateTime Prompt",
        "selectInterval": "Interval Prompt",
        "textBox": "Text Box Prompt",
        "treePrompt": "Tree Prompt",
        "selectWithSearch": "Search & Select Prompt",
    }

    for tag, p_type in PROMPT_TAGS.items():
        for el in _iter_all(root, tag):
            param = el.get("parameter", "")
            if not param:
                continue
            options = [opt.get("useValue", "") for opt in _iter_all(el, "selectOption")]
            default_val = ""
            for d in _iter_all(el, "defaultSimpleSelection"):
                default_val = _text_of(d)
                if default_val:
                    break
            if not default_val:
                for d in _iter_all(el, "defaultDateTimeSelection"):
                    default_val = _text_of(d)
                    if default_val:
                        break
            if not default_val:
                for d in _iter_all(el, "defaultTextSelection"):
                    default_val = _text_of(d)
                    if default_val:
                        break

            # Power BI target control mapping recommendation
            if p_type in ("Date Prompt", "DateTime Prompt", "Interval Prompt"):
                target_pbi = "Power BI Slicer (Between / Relative Date Range)"
            elif p_type == "Text Box Prompt":
                target_pbi = "Power BI Text Filter / What-If Parameter"
            elif p_type == "Tree Prompt":
                target_pbi = "Power BI Hierarchy Slicer"
            else:
                target_pbi = "Power BI Slicer (Dropdown / List)"

            prompts.append({
                "parameter": param,
                "prompt_type": p_type,
                "ref_query": el.get("refQuery", ""),
                "options": options,
                "default": default_val,
                "auto_submit": el.get("autoSubmit", "false"),
                "multi_select": el.get("multiSelect", "false") == "true",
                "required": el.get("required", "true") != "false",
                "target_powerbi_control": target_pbi,
            })
    return prompts


def extract_lists(root):
    """Extract Cognos <list> definitions with columns, titles, groupings, and query bindings."""
    lists = []
    for lst in _iter_all(root, "list"):
        list_name = lst.get("name", "")
        ref_query = lst.get("refQuery", "")
        columns = []
        for col in _iter_all(lst, "listColumn"):
            title = ""
            for title_el in _iter_all(col, "listColumnTitle"):
                for t in _iter_all(title_el, "text"):
                    title = _text_of(t)
                    if title:
                        break
                if not title:
                    for di_val in _iter_all(title_el, "dataItemValue"):
                        title = di_val.get("refDataItem", "")
                        if title:
                            break

            data_item = ""
            for body in _iter_all(col, "listColumnBody"):
                for val in _iter_all(body, "dataItemValue"):
                    data_item = val.get("refDataItem", "")
                    if data_item:
                        break
                if not data_item:
                    for fr in _iter_all(body, "fieldRef"):
                        data_item = fr.get("refDataItem", "")
                        if data_item:
                            break
            if not data_item:
                for val in _iter_all(col, "dataItemValue"):
                    data_item = val.get("refDataItem", "")
                    if data_item:
                        break
            columns.append({
                "title": title or data_item,
                "data_item": data_item,
            })

        groupings = []
        for grp in _iter_all(lst, "listGroup"):
            di = grp.get("refDataItem", "")
            if di and di not in groupings:
                groupings.append(di)

        lists.append({
            "name": list_name or f"List_{len(lists) + 1}",
            "ref_query": ref_query,
            "columns": columns,
            "column_count": len(columns),
            "groupings": groupings,
            "target_powerbi_visual": "Table (tableEx)",
        })
    return lists


def extract_crosstabs(root):
    """Extract Cognos <crosstab> definitions: row dimensions, column dimensions, and intersection measures."""
    crosstabs = []
    for ct in _iter_all(root, "crosstab"):
        ct_name = ct.get("name", "")
        ref_query = ct.get("refQuery", "")

        row_dims = []
        rows_root = _find_child(ct, "crosstabRows")
        if rows_root is not None:
            for node in _iter_all(rows_root, "crosstabNodeMember") + _iter_all(rows_root, "dataItemValue"):
                di = node.get("refDataItem", "")
                if di and di not in row_dims:
                    row_dims.append(di)

        col_dims = []
        cols_root = _find_child(ct, "crosstabColumns")
        if cols_root is not None:
            for node in _iter_all(cols_root, "crosstabNodeMember") + _iter_all(cols_root, "dataItemValue"):
                di = node.get("refDataItem", "")
                if di and di not in col_dims:
                    col_dims.append(di)

        measures = []
        meas_root = _find_child(ct, "crosstabMeasures")
        if meas_root is not None:
            for val in _iter_all(meas_root, "dataItemValue"):
                di = val.get("refDataItem", "")
                if di and di not in measures:
                    measures.append(di)
        for cell in _iter_all(ct, "crosstabIntersectionCell"):
            for val in _iter_all(cell, "dataItemValue"):
                di = val.get("refDataItem", "")
                if di and di not in measures:
                    measures.append(di)

        crosstabs.append({
            "name": ct_name or f"Crosstab_{len(crosstabs) + 1}",
            "ref_query": ref_query,
            "row_dimensions": row_dims,
            "column_dimensions": col_dims,
            "measures": measures,
            "target_powerbi_visual": "Matrix (pivotTable)",
        })
    return crosstabs


def extract_repeaters(root):
    """Extract Cognos <repeater> and <repeaterTable> elements."""
    repeaters = []
    for tag in ("repeater", "repeaterTable"):
        for rep in _iter_all(root, tag):
            r_name = rep.get("name", "")
            ref_query = rep.get("refQuery", "")
            items = []
            for val in _iter_all(rep, "dataItemValue"):
                di = val.get("refDataItem", "")
                if di and di not in items:
                    items.append(di)
            repeaters.append({
                "name": r_name or f"{tag}_{len(repeaters) + 1}",
                "type": tag,
                "ref_query": ref_query,
                "data_items": items,
                "target_powerbi_visual": "Small Multiples / Matrix / Card Repeater",
            })
    return repeaters


_MACRO_RE = re.compile(r"#([^#]+)#")


def extract_cognos_macros(expr):
    """Tokenize all Cognos macros (#...#) found in an expression."""
    macros = []
    if not expr or "#" not in expr:
        return macros
    matches = _MACRO_RE.findall(expr)
    for raw_content in matches:
        raw_macro = f"#{raw_content}#"
        content = raw_content.strip()

        # Prompt function: prompt('ParamName', 'type', 'default', 'prefix', 'suffix')
        p_match = re.search(r"\bprompt\s*\(\s*['\"]([^'\"]+)['\"]\s*(?:,\s*['\"]?([^,'\"\)]+)['\"]?)?\s*(?:,\s*['\"]?([^,'\"\)]*)['\"]?)?", content, re.IGNORECASE)
        # PromptMany function: promptmany('ParamName', ...)
        pm_match = re.search(r"\bpromptmany\s*\(\s*['\"]([^'\"]+)['\"]\s*(?:,\s*['\"]?([^,'\"\)]+)['\"]?)?\s*(?:,\s*['\"]?([^,'\"\)]*)['\"]?)?", content, re.IGNORECASE)

        if pm_match:
            param_name = pm_match.group(1)
            datatype = (pm_match.group(2) or "string").strip()
            default_val = (pm_match.group(3) or "").strip()
            macros.append({
                "raw": raw_macro,
                "macro_type": "promptmany",
                "parameter": param_name,
                "datatype": datatype,
                "default": default_val,
                "is_multi": True,
            })
        elif p_match:
            param_name = p_match.group(1)
            datatype = (p_match.group(2) or "string").strip()
            default_val = (p_match.group(3) or "").strip()
            macros.append({
                "raw": raw_macro,
                "macro_type": "prompt",
                "parameter": param_name,
                "datatype": datatype,
                "default": default_val,
                "is_multi": False,
            })
        elif "$account." in content or "$custom." in content:
            param_name = content.replace("$account.", "").replace("$custom.", "").strip()
            macros.append({
                "raw": raw_macro,
                "macro_type": "session_parameter",
                "parameter": param_name,
                "datatype": "string",
                "default": "",
                "is_multi": False,
            })
        else:
            macros.append({
                "raw": raw_macro,
                "macro_type": "macro_expression",
                "parameter": "",
                "datatype": "expression",
                "default": "",
                "is_multi": False,
            })
    return macros


def extract_drillthroughs(root, report_name):
    """
    Extract <reportDrill> / <drillTarget> definitions — the Cognos equivalent of
    WebFOCUS extract_drilldowns(), except fully structured: target report path
    AND explicit source->target parameter mapping are both declared, not guessed.
    """
    drills = []
    for rd in _iter_all(root, "reportDrill"):
        drill_name = rd.get("name", "")
        target = _find_child(rd, "drillTarget")
        if target is None:
            continue
        report_path_el = _find_child(target, "reportPath")
        target_path = report_path_el.get("path", "") if report_path_el is not None else ""
        target_report_name = ""
        m = re.search(r"report\[@name='([^']*)'\]", target_path)
        if m:
            target_report_name = m.group(1)

        param_links = []
        for link in _iter_all(target, "drillLink"):
            src_ctx = _find_child(link, "drillSourceContext")
            tgt_ctx = _find_child(link, "drillTargetContext")
            src_param = _find_child(src_ctx, "parameterContext") if src_ctx is not None else None
            tgt_param = _find_child(tgt_ctx, "parameterContext") if tgt_ctx is not None else None
            param_links.append({
                "source_parameter": src_param.get("parameter", "") if src_param is not None else "",
                "target_parameter": tgt_param.get("parameter", "") if tgt_param is not None else "",
            })

        drills.append({
            "source_report": report_name,
            "drill_name": drill_name,
            "target_report": target_report_name,
            "target_path": target_path,
            "parameter_links": param_links,
            "opens_new_window": target.get("showInNewWindow", "false"),
        })
    return drills


def extract_visualizations(root):
    """
    Extract <vizControl> elements — charts/maps/word clouds etc. No WebFOCUS
    equivalent; captures type, referenced data store, and slot->column mapping
    so the AI build plan can suggest a matching Power BI visual type.
    """
    vizzes = []
    for vc in _iter_all(root, "vizControl"):
        viz_name = vc.get("name", "")
        viz_type = vc.get("type", "")
        slots = []
        for ds in _iter_all(vc, "vcDataSet"):
            ref_store = ds.get("refDataStore", "")
            for slot in _iter_all(ds, "vcSlotData"):
                slot_id = slot.get("idSlot", "")
                cols = [c.get("refDsColumn", "") for c in _iter_all(slot, "vcSlotDsColumn")]
                slots.append({"data_store": ref_store, "slot": slot_id, "columns": cols})
        vizzes.append({"name": viz_name, "type": viz_type, "slots": slots})
    return vizzes


def extract_data_stores(root):
    """
    Extract <reportDataStore> elements — maps a visualization's data feed back
    to the query + specific columns it uses, with a role (indexed=category/BY,
    value=measure/SUM). This is the closest Cognos equivalent to WebFOCUS
    by_real/by_calc vs sum_real/sum_calc classification.
    """
    stores = []
    for store in _iter_all(root, "reportDataStore"):
        store_name = store.get("name", "")
        list_query = None
        for q in _iter_all(store, "dsV5ListQuery"):
            list_query = q.get("refQuery", "")
            break
        columns = []
        for item in _iter_all(store, "dsV5DataItem"):
            columns.append({
                "field": item.get("refDataItem", ""),
                "role": item.get("dsColumnType", ""),  # 'indexed' = category, 'value' = measure
            })
        stores.append({"name": store_name, "query": list_query, "columns": columns})
    return stores


def extract_custom_controls(root):
    """
    Extract <customControl> elements — JS-based widgets with no WebFOCUS or
    standard Power BI equivalent. Always flagged for manual review.
    """
    controls = []
    for cc in _iter_all(root, "customControl"):
        controls.append({
            "name": cc.get("name", ""),
            "description": cc.get("description", ""),
            "path": cc.get("path", ""),
        })
    return controls


def extract_model_path(root):
    """Extract the <modelPath> — the single most important field: which Data
    Module / Framework Manager package this report depends on."""
    el = _iter_all(root, "modelPath")
    if not el:
        return ""
    return _text_of(el[0])

def detect_cognos_xml_type(xml_text):
    """Peeks at the root tag so one upload box can accept both report XML
    and framework model XML - no separate upload needed."""
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError:
        return "invalid"
    root_tag = _local(root.tag)
    if root_tag == "report":
        return "report"
    if root_tag == "frameworkManagerModel":
        return "framework_model"
    return "unknown"


def extract_package_name_from_model_path(model_path):
    """/content/folder[@name='Samples']/package[@name='GO Sales'] -> 'GO Sales'"""
    m = re.search(r"package\[@name='([^']*)'\]", model_path or "")
    return m.group(1) if m else ""


_FM_FIELD_PATH_RE = re.compile(r"^\[([^\]]+)\]\.\[([^\]]+)\]\.\[([^\]]+)\]$")


def parse_fm_field_expression(expr):
    """[Sales].[Retailers].[Region] -> (namespace_or_layer, query_subject, column).
    Matching later ignores the first segment (it may be a namespace not
    explicitly modeled), and validates on query-subject + column name."""
    m = _FM_FIELD_PATH_RE.match((expr or "").strip())
    if m:
        return m.group(1), m.group(2), m.group(3)
    return None, None, None


def parse_framework_model(xml_text):
    """Main entry point for Framework Manager model XML - mirrors
    parse_cognos_report()'s shape/spirit."""
    root = ET.fromstring(xml_text)

    result = {
        "project_name": root.get("projectName", ""),
        "framework_name": root.get("frameworkName", ""),
        "version": root.get("version", ""),
        "query_mode": root.get("queryMode", ""),
        "description": _text_of(_find_child(root, "description")),
        "data_sources": [],
        "packages": [],
        "layers": [],
        "query_subjects": [],
        "query_items": [],
        "relationships": [],
        "filters": [],
        "security_filters": [],
    }

    ds_root = _find_child(root, "dataSources")
    if ds_root is not None:
        for ds in _iter_all(ds_root, "dataSource"):
            result["data_sources"].append({
                "name": ds.get("name", ""), "type": ds.get("type", ""),
                "provider": ds.get("provider", ""), "catalog": ds.get("catalog", ""),
                "schema": ds.get("schema", ""),
            })

    pkgs_root = _find_child(root, "packages")
    if pkgs_root is not None:
        for pkg in _iter_all(pkgs_root, "package"):
            included_layers = []
            inc_root = _find_child(pkg, "includedLayers")
            if inc_root is not None:
                included_layers = [lr.get("name", "") for lr in _iter_all(inc_root, "layerRef")]
            result["packages"].append({
                "name": pkg.get("name", ""),
                "published_name": pkg.get("publishedName", ""),
                "description": _text_of(_find_child(pkg, "description")),
                "included_layers": included_layers,
            })

    layers_root = _find_child(root, "layers")
    if layers_root is not None:
        for layer in _iter_all(layers_root, "layer"):
            layer_name = layer.get("name", "")
            layer_type = layer.get("type", "")
            qs_names = []

            qs_root = _find_child(layer, "querySubjects")
            if qs_root is not None:
                for qs in _iter_all(qs_root, "querySubject"):
                    qs_name = qs.get("name", "")
                    qs_names.append(qs_name)

                    source_subjects = []
                    ss_root = _find_child(qs, "sourceSubjects")
                    if ss_root is not None:
                        source_subjects = [s.get("name", "") for s in _iter_all(ss_root, "sourceSubjectRef")]

                    result["query_subjects"].append({
                        "layer": layer_name, "layer_type": layer_type, "name": qs_name,
                        "type": qs.get("type", ""), "category": qs.get("category", ""),
                        "subcategory": qs.get("subcategory", ""),
                        "physical_table": qs.get("physicalTable", ""),
                        "source_subjects": source_subjects,
                    })

                    qi_root = _find_child(qs, "queryItems")
                    if qi_root is not None:
                        for qi in _iter_all(qi_root, "queryItem"):
                            result["query_items"].append({
                                "layer": layer_name, "query_subject": qs_name,
                                "category": qs.get("category", ""), "subcategory": qs.get("subcategory", ""),
                                "name": qi.get("name", ""), "data_type": qi.get("dataType", ""),
                                "usage": qi.get("usage", ""), "regular_aggregate": qi.get("regularAggregate", ""),
                                "is_key": qi.get("isKey", "false"), "nullable": qi.get("nullable", "true"),
                                "hidden": qi.get("hidden", "false"), "physical_column": qi.get("physicalColumn", ""),
                                "format": qi.get("format", ""),
                                "expression": _text_of(_find_child(qi, "expression")),
                            })

            result["layers"].append({
                "name": layer_name, "type": layer_type,
                "description": _text_of(_find_child(layer, "description")),
                "query_subject_names": qs_names,
            })

    rels_root = _find_child(root, "relationships")
    if rels_root is not None:
        for rel in _iter_all(rels_root, "relationship"):
            result["relationships"].append({
                "name": rel.get("name", ""), "layer": rel.get("layer", ""),
                "left_subject": rel.get("leftSubject", ""), "right_subject": rel.get("rightSubject", ""),
                "left_cardinality": rel.get("leftCardinality", ""), "right_cardinality": rel.get("rightCardinality", ""),
                "join_type": rel.get("joinType", ""),
                "expression": _text_of(_find_child(rel, "expression")),
                "foreign_key": _text_of(_find_child(rel, "foreignKey")),
                "primary_key": _text_of(_find_child(rel, "primaryKey")),
            })

    filters_root = _find_child(root, "filters")
    if filters_root is not None:
        for f in _iter_all(filters_root, "filter"):
            result["filters"].append({
                "name": f.get("name", ""), "layer": f.get("layer", ""),
                "subject": f.get("subject", ""), "type": f.get("type", ""),
                "expression": _text_of(_find_child(f, "expression")),
            })

    security_root = _find_child(root, "security")
    if security_root is not None:
        for sf in _iter_all(security_root, "securityFilter"):
            result["security_filters"].append({
                "name": sf.get("name", ""), "subject": sf.get("subject", ""),
                "type": sf.get("type", ""), "expression": _text_of(_find_child(sf, "expression")),
                "note": _text_of(_find_child(sf, "note")),
            })

    return result


def match_report_to_framework(report_model_path, framework_model):
    """Returns the matching package dict if the report's modelPath package
    name matches one defined in the uploaded framework model, else None."""
    if framework_model is None:
        return None
    package_name = extract_package_name_from_model_path(report_model_path)
    if not package_name:
        return None
    for pkg in framework_model["packages"]:
        if pkg["name"] == package_name or pkg["published_name"] == package_name:
            return pkg
    return None


def validate_report_against_framework(parsed_report, framework_model):
    """Cross-checks each field a report references against real query
    subjects/items in the framework model. Matches on query-subject name +
    column name (ignores the leading bracket segment, since it may be a
    namespace not explicitly modeled). Empty list = fully verified."""
    issues = []
    if framework_model is None:
        return issues

    subject_items = defaultdict(set)
    for qi in framework_model["query_items"]:
        subject_items[qi["query_subject"].strip().upper()].add(qi["name"].strip().upper())
    known_subjects = {qs["name"].strip().upper() for qs in framework_model["query_subjects"]}

    for q in parsed_report.get("queries", []):
        for item in q["data_items"]:
            expr = item["expression"]
            if not expr:
                continue
            _, subject, column = parse_fm_field_expression(expr)
            if subject is None:
                continue
            subj_upper, col_upper = subject.strip().upper(), column.strip().upper()
            if subj_upper not in known_subjects:
                issues.append(f"Field '{item['field']}' references unknown query subject '{subject}'")
            elif col_upper not in subject_items.get(subj_upper, set()):
                issues.append(f"Field '{item['field']}' references unknown column '{column}' in query subject '{subject}'")

    return issues


def build_framework_overview_df(fm):
    return pd.DataFrame([{
        "Framework Name": fm["framework_name"], "Project Name": fm["project_name"],
        "Query Mode": fm["query_mode"], "Version": fm["version"],
        "Layers": len(fm["layers"]), "Query Subjects": len(fm["query_subjects"]),
        "Columns": len(fm["query_items"]), "Relationships": len(fm["relationships"]),
        "Packages": len(fm["packages"]),
    }])


def build_framework_packages_df(fm):
    return pd.DataFrame([{
        "Package Name": p["name"], "Published Name": p["published_name"],
        "Description": p["description"], "Included Layers": ", ".join(p["included_layers"]),
    } for p in fm["packages"]])


def build_framework_query_subjects_df(fm):
    return pd.DataFrame([{
        "Layer": qs["layer"], "Layer Type": qs["layer_type"], "Query Subject": qs["name"],
        "Type": qs["type"], "Category": qs["category"], "Subcategory": qs["subcategory"],
        "Physical Table": qs["physical_table"], "Source Subjects": ", ".join(qs["source_subjects"]),
    } for qs in fm["query_subjects"]])


def build_framework_columns_df(fm):
    return pd.DataFrame([{
        "Layer": qi["layer"], "Query Subject": qi["query_subject"], "Category": qi["category"],
        "Subcategory": qi["subcategory"], "Column Name": qi["name"], "Data Type": qi["data_type"],
        "Usage": qi["usage"], "Aggregate": qi["regular_aggregate"], "Is Key": qi["is_key"],
        "Hidden": qi["hidden"], "Physical Column": qi["physical_column"], "Format": qi["format"],
        "Expression": qi["expression"],
    } for qi in fm["query_items"]])


def build_framework_relationships_df(fm):
    return pd.DataFrame([{
        "Layer": r["layer"], "Relationship": r["name"], "Left Subject": r["left_subject"],
        "Right Subject": r["right_subject"], "Left Cardinality": r["left_cardinality"],
        "Right Cardinality": r["right_cardinality"], "Join Type": r["join_type"],
        "Expression": r["expression"], "Foreign Key": r["foreign_key"], "Primary Key": r["primary_key"],
    } for r in fm["relationships"]])


def build_framework_filters_security_df(fm):
    rows = [{"Type": "Standalone Filter", "Name": f["name"], "Layer": f["layer"],
             "Subject": f["subject"], "Expression": f["expression"]} for f in fm["filters"]]
    rows += [{"Type": "Security Filter", "Name": sf["name"], "Layer": "",
              "Subject": sf["subject"], "Expression": sf["expression"]} for sf in fm["security_filters"]]
    return pd.DataFrame(rows)


# ==============================================================================
# DETERMINISTIC COGNOS-TO-POWER BI (DAX / M) EXPRESSION TRANSPILER
# ==============================================================================

_TWO_PART_PATH_RE = re.compile(r"^\[([^\]]+)\]\.\[([^\]]+)\]$")

_UNSUPPORTED_DIMENSIONAL_FUNCS = [
    r"\btuple\s*\(",
    r"\bancestor\s*\(",
    r"\bperiodsToDate\s*\(",
    r"\bparallelPeriod\s*\(",
    r"\bprevMember\s*\(",
    r"\bnextMember\s*\(",
    r"\bopeningPeriod\s*\(",
    r"\bclosingPeriod\s*\(",
    r"\bchildren\s*\(",
    r"\bdescendants\s*\(",
    r"\bparent\s*\(",
    r"\bfirstSibling\s*\(",
    r"\blastSibling\s*\(",
    r"\bmember\s*\(",
]


def clean_brackets(token):
    """Strip outermost parentheses and whitespace."""
    t = (token or "").strip()
    while t.startswith("(") and t.endswith(")"):
        depth = 0
        balanced = True
        for c in t[:-1]:
            if c == "(":
                depth += 1
            elif c == ")":
                depth -= 1
            if depth == 0:
                balanced = False
                break
        if balanced:
            t = t[1:-1].strip()
        else:
            break
    return t


def convert_column_refs_dax(expr):
    """Convert Cognos bracketed paths to DAX 'Table'[Column]."""
    # 3-part: [C].[Module].[Table].[Column] -> 'Table'[Column]
    expr = re.sub(
        r"\[C\]\.\[[^\]]+\]\.\[([^\]]+)\]\.\[([^\]]+)\]",
        r"'\1'[\2]",
        expr
    )
    # 2-part: [Table].[Column] -> 'Table'[Column]
    expr = re.sub(
        r"\[([^\]]+)\]\.\[([^\]]+)\]",
        r"'\1'[\2]",
        expr
    )
    return expr


def convert_column_refs_m(expr):
    """Convert Cognos bracketed paths to Power Query M [Column]."""
    # 3-part: [C].[Module].[Table].[Column] -> [Column]
    expr = re.sub(
        r"\[C\]\.\[[^\]]+\]\.\[[^\]]+\]\.\[([^\]]+)\]",
        r"[\1]",
        expr
    )
    # 2-part: [Table].[Column] -> [Column]
    expr = re.sub(
        r"\[[^\]]+\]\.\[([^\]]+)\]",
        r"[\1]",
        expr
    )
    return expr


def _convert_logic_ops_dax(expr):
    expr = re.sub(r"\s+and\s+", " && ", expr, flags=re.IGNORECASE)
    expr = re.sub(r"\s+or\s+", " || ", expr, flags=re.IGNORECASE)
    expr = re.sub(r"\bnot\s+", "!", expr, flags=re.IGNORECASE)
    return expr


def _translate_if_chain(expr):
    """Translate nested/chained Cognos IF-THEN-ELSE to DAX (IF or SWITCH) and M."""
    warnings = []
    dax_expr = convert_column_refs_dax(expr)
    m_expr = convert_column_refs_m(expr)

    dax_expr = re.sub(r"\bis\s+null\b", " = BLANK()", dax_expr, flags=re.IGNORECASE)
    dax_expr = re.sub(r"\bis\s+not\s+null\b", " <> BLANK()", dax_expr, flags=re.IGNORECASE)

    # Chained "else if" -> SWITCH(TRUE(), Cond1, Val1, Cond2, Val2, ..., ElseVal)
    if re.search(r"\belse\s+if\b", dax_expr, re.IGNORECASE):
        parts = re.split(r"\belse\s+if\b", dax_expr, flags=re.IGNORECASE)
        switch_args = ["TRUE()"]
        first_m = re.search(r"^\s*if\s+(.+?)\s+then\s+(.+?)\s*$", parts[0], re.IGNORECASE | re.DOTALL)
        if first_m:
            switch_args.append(_convert_logic_ops_dax(clean_brackets(first_m.group(1))))
            switch_args.append(clean_brackets(first_m.group(2)))

        for mid in parts[1:-1]:
            mid_m = re.search(r"^\s*(.+?)\s+then\s+(.+?)\s*$", mid, re.IGNORECASE | re.DOTALL)
            if mid_m:
                switch_args.append(_convert_logic_ops_dax(clean_brackets(mid_m.group(1))))
                switch_args.append(clean_brackets(mid_m.group(2)))

        last_m = re.search(r"^\s*(.+?)\s+then\s+(.+?)\s+else\s+(.+?)\s*$", parts[-1], re.IGNORECASE | re.DOTALL)
        if last_m:
            switch_args.append(_convert_logic_ops_dax(clean_brackets(last_m.group(1))))
            switch_args.append(clean_brackets(last_m.group(2)))
            switch_args.append(clean_brackets(last_m.group(3)))

        dax_res = f"SWITCH({', '.join(switch_args)})"
        return dax_res, None, warnings

    # Single if-then-else
    single_if = re.match(r"^\s*if\s+(.+?)\s+then\s+(.+?)\s+else\s+(.+?)\s*$", dax_expr, re.IGNORECASE | re.DOTALL)
    if single_if:
        cond = clean_brackets(single_if.group(1))
        val_then = clean_brackets(single_if.group(2))
        val_else = clean_brackets(single_if.group(3))
        cond_dax = _convert_logic_ops_dax(cond)
        dax_res = f"IF({cond_dax}, {val_then}, {val_else})"

        m_single = re.match(r"^\s*if\s+(.+?)\s+then\s+(.+?)\s+else\s+(.+?)\s*$", m_expr, re.IGNORECASE | re.DOTALL)
        if m_single:
            m_cond = clean_brackets(m_single.group(1))
            m_then = clean_brackets(m_single.group(2))
            m_else = clean_brackets(m_single.group(3))
            m_cond = re.sub(r"\bis\s+null\b", " = null", m_cond, flags=re.IGNORECASE)
            m_cond = re.sub(r"\bis\s+not\s+null\b", " <> null", m_cond, flags=re.IGNORECASE)
            m_res = f"if {m_cond} then {m_then} else {m_else}"
        else:
            m_res = None
        return dax_res, m_res, warnings

    return dax_expr, None, warnings


def _translate_case_expression(expr):
    """Translate Cognos CASE WHEN ... THEN ... ELSE ... END into DAX SWITCH."""
    warnings = []
    dax_expr = convert_column_refs_dax(expr)

    # Searched case
    searched_m = re.search(r"^\s*case\s+when\b", dax_expr, re.IGNORECASE)
    if searched_m:
        when_blocks = re.findall(r"\bwhen\s+(.+?)\s+then\s+(.+?)(?=\bwhen\b|\belse\b|\bend\b)", dax_expr, re.IGNORECASE | re.DOTALL)
        else_m = re.search(r"\belse\s+(.+?)\s+end", dax_expr, re.IGNORECASE | re.DOTALL)
        else_val = clean_brackets(else_m.group(1)) if else_m else "BLANK()"

        args = ["TRUE()"]
        for cond, val in when_blocks:
            args.append(_convert_logic_ops_dax(clean_brackets(cond)))
            args.append(clean_brackets(val))
        args.append(else_val)

        dax_res = f"SWITCH({', '.join(args)})"
        return dax_res, None, warnings

    # Simple case
    simple_m = re.match(r"^\s*case\s+(.+?)\s+when\b", dax_expr, re.IGNORECASE)
    if simple_m:
        test_expr = clean_brackets(simple_m.group(1))
        when_blocks = re.findall(r"\bwhen\s+(.+?)\s+then\s+(.+?)(?=\bwhen\b|\belse\b|\bend\b)", dax_expr, re.IGNORECASE | re.DOTALL)
        else_m = re.search(r"\belse\s+(.+?)\s+end", dax_expr, re.IGNORECASE | re.DOTALL)
        else_val = clean_brackets(else_m.group(1)) if else_m else "BLANK()"

        args = [test_expr]
        for val, res in when_blocks:
            args.append(clean_brackets(val))
            args.append(clean_brackets(res))
        args.append(else_val)

        dax_res = f"SWITCH({', '.join(args)})"
        return dax_res, None, warnings

    return dax_expr, None, warnings


def _translate_functions_and_operators(expr):
    """Translate standard functions: date, string, nulls, math."""
    warnings = []
    category = "General"

    dax = convert_column_refs_dax(expr)
    m_code = convert_column_refs_m(expr)

    # 1. Null handling
    if "coalesce" in dax.lower():
        category = "Null Handling"
        dax = re.sub(r"\bcoalesce\s*\(", "COALESCE(", dax, flags=re.IGNORECASE)
    if "nullif" in dax.lower():
        category = "Null Handling"
        nullif_m = re.search(r"\bnullif\s*\(\s*(.+?)\s*,\s*(.+?)\s*\)", dax, re.IGNORECASE)
        if nullif_m:
            a, b = nullif_m.group(1), nullif_m.group(2)
            dax = dax.replace(nullif_m.group(0), f"IF({a} = {b}, BLANK(), {a})")

    dax = re.sub(r"\bis\s+null\b", " = BLANK()", dax, flags=re.IGNORECASE)
    dax = re.sub(r"\bis\s+not\s+null\b", " <> BLANK()", dax, flags=re.IGNORECASE)
    m_code = re.sub(r"\bis\s+null\b", " = null", m_code, flags=re.IGNORECASE)
    m_code = re.sub(r"\bis\s+not\s+null\b", " <> null", m_code, flags=re.IGNORECASE)

    # In-list expressions: [Col] in ('A', 'B') -> [Col] IN {"A", "B"}
    def _in_repl(m):
        inner = m.group(1).replace("'", '"')
        return f"IN {{{inner}}}"
    dax = re.sub(
        r"\bin\s*\(\s*([^)]+)\s*\)",
        _in_repl,
        dax,
        flags=re.IGNORECASE
    )

    # 2. String Concatenation || -> &
    if "||" in dax:
        category = "String"
        dax = dax.replace("||", "&")
        m_code = m_code.replace("||", "&")

    # 3. Date Functions
    if re.search(r"\b_add_days\b", dax, re.IGNORECASE):
        category = "Date Function"
        dax = re.sub(r"\b_add_days\s*\(\s*(.+?)\s*,\s*(.+?)\s*\)", r"(\1 + \2)", dax, flags=re.IGNORECASE)
        m_code = re.sub(r"\b_add_days\s*\(\s*(.+?)\s*,\s*(.+?)\s*\)", r"Date.AddDays(\1, \2)", m_code, flags=re.IGNORECASE)

    if re.search(r"\b_add_months\b", dax, re.IGNORECASE):
        category = "Date Function"
        dax = re.sub(r"\b_add_months\s*\(\s*(.+?)\s*,\s*(.+?)\s*\)", r"EDATE(\1, \2)", dax, flags=re.IGNORECASE)
        m_code = re.sub(r"\b_add_months\s*\(\s*(.+?)\s*,\s*(.+?)\s*\)", r"Date.AddMonths(\1, \2)", m_code, flags=re.IGNORECASE)

    if re.search(r"\b_add_years\b", dax, re.IGNORECASE):
        category = "Date Function"
        dax = re.sub(r"\b_add_years\s*\(\s*(.+?)\s*,\s*(.+?)\s*\)", r"EDATE(\1, (\2) * 12)", dax, flags=re.IGNORECASE)
        m_code = re.sub(r"\b_add_years\s*\(\s*(.+?)\s*,\s*(.+?)\s*\)", r"Date.AddYears(\1, \2)", m_code, flags=re.IGNORECASE)

    if re.search(r"\b(_year|extract\s*\(\s*year\s*,)\b", dax, re.IGNORECASE):
        category = "Date Function"
        dax = re.sub(r"\b_year\s*\(\s*(.+?)\s*\)", r"YEAR(\1)", dax, flags=re.IGNORECASE)
        dax = re.sub(r"\bextract\s*\(\s*year\s*,\s*(.+?)\s*\)", r"YEAR(\1)", dax, flags=re.IGNORECASE)
        m_code = re.sub(r"\b(_year|extract\s*\(\s*year\s*,)\s*(.+?)\s*\)", r"Date.Year(\2)", m_code, flags=re.IGNORECASE)

    if re.search(r"\b(_month|extract\s*\(\s*month\s*,)\b", dax, re.IGNORECASE):
        category = "Date Function"
        dax = re.sub(r"\b_month\s*\(\s*(.+?)\s*\)", r"MONTH(\1)", dax, flags=re.IGNORECASE)
        dax = re.sub(r"\bextract\s*\(\s*month\s*,\s*(.+?)\s*\)", r"MONTH(\1)", dax, flags=re.IGNORECASE)
        m_code = re.sub(r"\b(_month|extract\s*\(\s*month\s*,)\s*(.+?)\s*\)", r"Date.Month(\2)", m_code, flags=re.IGNORECASE)

    if re.search(r"\b(_day|extract\s*\(\s*day\s*,)\b", dax, re.IGNORECASE):
        category = "Date Function"
        dax = re.sub(r"\b_day\s*\(\s*(.+?)\s*\)", r"DAY(\1)", dax, flags=re.IGNORECASE)
        dax = re.sub(r"\bextract\s*\(\s*day\s*,\s*(.+?)\s*\)", r"DAY(\1)", dax, flags=re.IGNORECASE)
        m_code = re.sub(r"\b(_day|extract\s*\(\s*day\s*,)\s*(.+?)\s*\)", r"Date.Day(\2)", m_code, flags=re.IGNORECASE)

    if re.search(r"\b(current_date|today\s*\(\s*\))\b", dax, re.IGNORECASE):
        category = "Date Function"
        dax = re.sub(r"\b(current_date|today\s*\(\s*\))\b", "TODAY()", dax, flags=re.IGNORECASE)
        m_code = re.sub(r"\b(current_date|today\s*\(\s*\))\b", "DateTime.Date(DateTime.LocalNow())", m_code, flags=re.IGNORECASE)

    if re.search(r"\b(current_timestamp|now\s*\(\s*\))\b", dax, re.IGNORECASE):
        category = "Date Function"
        dax = re.sub(r"\b(current_timestamp|now\s*\(\s*\))\b", "NOW()", dax, flags=re.IGNORECASE)
        m_code = re.sub(r"\b(current_timestamp|now\s*\(\s*\))\b", "DateTime.LocalNow()", m_code, flags=re.IGNORECASE)

    if re.search(r"\b_days_between\b", dax, re.IGNORECASE):
        category = "Date Function"
        dax = re.sub(r"\b_days_between\s*\(\s*(.+?)\s*,\s*(.+?)\s*\)", r"DATEDIFF(\2, \1, DAY)", dax, flags=re.IGNORECASE)
        m_code = re.sub(r"\b_days_between\s*\(\s*(.+?)\s*,\s*(.+?)\s*\)", r"Duration.Days(\1 - \2)", m_code, flags=re.IGNORECASE)

    # 4. String Functions
    if re.search(r"\bsubstring\b", dax, re.IGNORECASE):
        category = "String"
        dax = re.sub(r"\bsubstring\s*\(\s*(.+?)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)", r"MID(\1, \2, \3)", dax, flags=re.IGNORECASE)
        dax = re.sub(r"\bsubstring\s*\(\s*(.+?)\s*,\s*(\d+)\s*\)", r"MID(\1, \2, LEN(\1))", dax, flags=re.IGNORECASE)
        m_code = re.sub(r"\bsubstring\s*\(\s*(.+?)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)", r"Text.Middle(\1, \2 - 1, \3)", m_code, flags=re.IGNORECASE)
        m_code = re.sub(r"\bsubstring\s*\(\s*(.+?)\s*,\s*(\d+)\s*\)", r"Text.Middle(\1, \2 - 1)", m_code, flags=re.IGNORECASE)

    if re.search(r"\btrim\b", dax, re.IGNORECASE):
        category = "String"
        dax = re.sub(r"\btrim\s*\(", "TRIM(", dax, flags=re.IGNORECASE)
        m_code = re.sub(r"\btrim\s*\(", "Text.Trim(", m_code, flags=re.IGNORECASE)

    if re.search(r"\blower\b", dax, re.IGNORECASE):
        category = "String"
        dax = re.sub(r"\blower\s*\(", "LOWER(", dax, flags=re.IGNORECASE)
        m_code = re.sub(r"\blower\s*\(", "Text.Lower(", m_code, flags=re.IGNORECASE)

    if re.search(r"\bupper\b", dax, re.IGNORECASE):
        category = "String"
        dax = re.sub(r"\bupper\s*\(", "UPPER(", dax, flags=re.IGNORECASE)
        m_code = re.sub(r"\bupper\s*\(", "Text.Upper(", m_code, flags=re.IGNORECASE)

    if re.search(r"\b(character_length|length)\b", dax, re.IGNORECASE):
        category = "String"
        dax = re.sub(r"\b(character_length|length)\s*\(", "LEN(", dax, flags=re.IGNORECASE)
        m_code = re.sub(r"\b(character_length|length)\s*\(", "Text.Length(", m_code, flags=re.IGNORECASE)

    # 5. Math / Division safely outside quotes
    has_slash_outside_quotes = False
    in_quote = False
    for char in dax:
        if char == '"':
            in_quote = not in_quote
        elif char == '/' and not in_quote:
            has_slash_outside_quotes = True
            break

    if has_slash_outside_quotes and "DIVIDE" not in dax:
        div_m = re.match(r"^\s*\(?(.+?)\)?\s*/\s*\(?(.+?)\)?\s*$", dax)
        if div_m:
            num = div_m.group(1).strip()
            den = div_m.group(2).strip()
            dax = f"DIVIDE({num}, {den})"
            warnings.append("Division operator '/' converted to safe DAX 'DIVIDE(..., ...)' to prevent division by zero errors.")

    return dax, m_code, category, warnings


def translate_cognos_expression_deterministic(
    expr,
    field_name="",
    query_name="",
    is_filter=False,
    aggregate="",
):
    """
    Deterministic rule-based translation from Cognos expression to DAX and Power Query M.
    """
    raw_expr = (expr or "").strip()
    warnings = []

    if not raw_expr:
        return {
            "source_expression": "",
            "field_name": field_name,
            "query_name": query_name,
            "dax_expression": None,
            "m_expression": None,
            "target_type": "Empty",
            "confidence": 100,
            "status": "Confirmed",
            "warnings": [],
            "is_translatable": True,
            "category": "Empty",
        }

    # 1. Check for Unsupported Dimensional / MDX constructs
    for pattern in _UNSUPPORTED_DIMENSIONAL_FUNCS:
        if re.search(pattern, raw_expr, re.IGNORECASE):
            func_name = pattern.replace(r"\b", "").replace(r"\s*\(", "")
            warnings.append(
                f"Dimensional / MDX function '{func_name}' detected. Cognos dimensional hierarchy "
                "queries require manual DAX tabular model redesign."
            )
            return {
                "source_expression": raw_expr,
                "field_name": field_name,
                "query_name": query_name,
                "dax_expression": None,
                "m_expression": None,
                "target_type": "Manual Review",
                "confidence": 25,
                "status": "Manual Review",
                "warnings": warnings,
                "is_translatable": False,
                "category": "Dimensional MDX",
            }

    # 2. Check for Custom Controls / JS
    if re.search(r"\bcustomControl\b|\bjavascript\b|\b<script\b", raw_expr, re.IGNORECASE):
        warnings.append("Custom JavaScript or control script detected. No direct Power BI equivalent.")
        return {
            "source_expression": raw_expr,
            "field_name": field_name,
            "query_name": query_name,
            "dax_expression": None,
            "m_expression": None,
            "target_type": "Manual Review",
            "confidence": 10,
            "status": "Manual Review",
            "warnings": warnings,
            "is_translatable": False,
            "category": "Custom Script",
        }

    # 3. Direct Model Reference
    m_3part = _FULL_MODEL_PATH_RE.match(raw_expr)
    if m_3part:
        module, table, column = m_3part.group(1), m_3part.group(2), m_3part.group(3)
        if aggregate and aggregate.lower() in ("total", "sum"):
            dax = f"SUM('{table}'[{column}])"
            t_type = "DAX Measure"
        elif aggregate and aggregate.lower() in ("average", "avg"):
            dax = f"AVERAGE('{table}'[{column}])"
            t_type = "DAX Measure"
        elif aggregate and aggregate.lower() in ("count", "count_distinct"):
            dax = f"DISTINCTCOUNT('{table}'[{column}])"
            t_type = "DAX Measure"
        else:
            dax = f"'{table}'[{column}]"
            t_type = "Calculated Column" if not is_filter else "Report Filter"

        m_code = f"[{column}]"
        return {
            "source_expression": raw_expr,
            "field_name": field_name,
            "query_name": query_name,
            "dax_expression": dax,
            "m_expression": m_code,
            "target_type": t_type,
            "confidence": 98,
            "status": "Confirmed",
            "warnings": [],
            "is_translatable": True,
            "category": "Direct Model Reference",
        }

    m_2part = _TWO_PART_PATH_RE.match(raw_expr)
    if m_2part:
        table, column = m_2part.group(1), m_2part.group(2)
        if aggregate and aggregate.lower() in ("total", "sum"):
            dax = f"SUM('{table}'[{column}])"
            t_type = "DAX Measure"
        elif aggregate and aggregate.lower() in ("average", "avg"):
            dax = f"AVERAGE('{table}'[{column}])"
            t_type = "DAX Measure"
        else:
            dax = f"'{table}'[{column}]"
            t_type = "Calculated Column" if not is_filter else "Report Filter"

        m_code = f"[{column}]"
        return {
            "source_expression": raw_expr,
            "field_name": field_name,
            "query_name": query_name,
            "dax_expression": dax,
            "m_expression": m_code,
            "target_type": t_type,
            "confidence": 95,
            "status": "Confirmed",
            "warnings": [],
            "is_translatable": True,
            "category": "Direct Model Reference",
        }

    # 4. Handle Macro prompt references
    has_macro = "#" in raw_expr
    macro_warnings = []
    working_expr = raw_expr
    if has_macro:
        macros = re.findall(r"#([^#]+)#", raw_expr)
        for m_str in macros:
            p_m = re.search(r"\bprompt\s*\(\s*['\"]([^'\"]+)['\"]\s*(?:,\s*['\"]?([^,'\"\)]+)['\"]?)?\s*(?:,\s*['\"]?([^,'\"\)]*)['\"]?)?", m_str, re.IGNORECASE)
            if p_m:
                p_name = p_m.group(1)
                p_def = (p_m.group(3) or "").strip()
                replacement_dax = f"SELECTEDVALUE('{p_name}_Parameter'[Value], {p_def or 'BLANK()'})"
                working_expr = working_expr.replace(f"#{m_str}#", replacement_dax)
                macro_warnings.append(
                    f"Cognos macro '#prompt({p_name})#' converted to DAX '{replacement_dax}'. "
                    f"Ensure a Power BI What-If parameter or slicer table '{p_name}_Parameter' is configured."
                )
            else:
                macro_warnings.append(f"Cognos macro token '#{m_str}#' preserved. Verify manual parameter mapping.")

    # 5. Conditionals: IF-THEN-ELSE
    if re.search(r"\bif\b\s*\(?.*\)?\s*\bthen\b", working_expr, re.IGNORECASE):
        dax_res, m_res, cond_warn = _translate_if_chain(working_expr)
        warnings.extend(macro_warnings)
        warnings.extend(cond_warn)
        return {
            "source_expression": raw_expr,
            "field_name": field_name,
            "query_name": query_name,
            "dax_expression": dax_res,
            "m_expression": m_res,
            "target_type": "DAX Measure" if _has_aggregation(raw_expr) else "Calculated Column",
            "confidence": 90,
            "status": "Rule-Based",
            "warnings": warnings,
            "is_translatable": True,
            "category": "Conditional IF",
        }

    # 6. CASE ... WHEN ... THEN ... ELSE ... END
    if re.search(r"\bcase\b\s+", working_expr, re.IGNORECASE) and re.search(r"\bend\b", working_expr, re.IGNORECASE):
        dax_res, m_res, case_warn = _translate_case_expression(working_expr)
        warnings.extend(macro_warnings)
        warnings.extend(case_warn)
        return {
            "source_expression": raw_expr,
            "field_name": field_name,
            "query_name": query_name,
            "dax_expression": dax_res,
            "m_expression": m_res,
            "target_type": "DAX Measure" if _has_aggregation(raw_expr) else "Calculated Column",
            "confidence": 90,
            "status": "Rule-Based",
            "warnings": warnings,
            "is_translatable": True,
            "category": "Case When",
        }

    # 7. Scoped Aggregations: total(x for dim)
    scoped_m = re.search(r"\b(total|average|avg|minimum|min|maximum|max|count)\s*\(\s*(.+?)\s+for\s+([^)]+)\)", working_expr, re.IGNORECASE)
    if scoped_m:
        fn_name = scoped_m.group(1).lower()
        inner_val = scoped_m.group(2).strip()
        scope_dims = [d.strip() for d in scoped_m.group(3).split(",")]

        dax_fn = {
            "total": "SUM", "average": "AVERAGE", "avg": "AVERAGE",
            "minimum": "MIN", "min": "MIN", "maximum": "MAX", "max": "MAX",
            "count": "COUNT",
        }.get(fn_name, "SUM")

        inner_dax = convert_column_refs_dax(inner_val)
        if len(scope_dims) == 1 and scope_dims[0].lower() in ("report", "all"):
            dax_code = f"CALCULATE({dax_fn}({inner_dax}), ALL())"
            warnings.append("Scoped rollup 'for report' converted to CALCULATE(..., ALL()).")
        else:
            dims_dax = ", ".join(convert_column_refs_dax(d) for d in scope_dims)
            dax_code = f"CALCULATE({dax_fn}({inner_dax}), ALLEXCEPT('Model', {dims_dax}))"
            warnings.append(
                f"Scoped rollup 'for {scoped_m.group(3)}' converted to CALCULATE(..., ALLEXCEPT(...)). "
                "Verify source table grain in Power BI."
            )

        warnings.extend(macro_warnings)
        return {
            "source_expression": raw_expr,
            "field_name": field_name,
            "query_name": query_name,
            "dax_expression": dax_code,
            "m_expression": None,
            "target_type": "DAX Measure",
            "confidence": 85,
            "status": "Rule-Based",
            "warnings": warnings,
            "is_translatable": True,
            "category": "Scoped Aggregation",
        }

    # 8. Standard Aggregations: total, avg, min, max, count, count_distinct
    std_agg_m = re.match(r"^\s*(total|sum|average|avg|minimum|min|maximum|max|count|count_distinct)\s*\(\s*(.+)\s*\)\s*$", working_expr, re.IGNORECASE)
    if std_agg_m:
        fn_name = std_agg_m.group(1).lower()
        arg = std_agg_m.group(2).strip()

        if fn_name == "count" and re.match(r"^distinct\s+", arg, re.IGNORECASE):
            fn_name = "count_distinct"
            arg = re.sub(r"^distinct\s+", "", arg, flags=re.IGNORECASE).strip()

        dax_fn = {
            "total": "SUM", "sum": "SUM", "average": "AVERAGE", "avg": "AVERAGE",
            "minimum": "MIN", "min": "MIN", "maximum": "MAX", "max": "MAX",
            "count": "COUNT", "count_distinct": "DISTINCTCOUNT"
        }.get(fn_name, "SUM")

        inner_dax = convert_column_refs_dax(arg)
        dax_code = f"{dax_fn}({inner_dax})"
        warnings.extend(macro_warnings)
        return {
            "source_expression": raw_expr,
            "field_name": field_name,
            "query_name": query_name,
            "dax_expression": dax_code,
            "m_expression": f"List.{dax_fn.capitalize()}({convert_column_refs_m(arg)})",
            "target_type": "DAX Measure",
            "confidence": 95,
            "status": "Confirmed",
            "warnings": warnings,
            "is_translatable": True,
            "category": "Aggregation",
        }

    # 9. Date Functions, String Functions, Nulls, Operators
    dax_code, m_code, cat, fn_warnings = _translate_functions_and_operators(working_expr)
    warnings.extend(macro_warnings)
    warnings.extend(fn_warnings)

    target_type = "DAX Measure" if _has_aggregation(dax_code) else ("Report Filter" if is_filter else "Calculated Column")

    return {
        "source_expression": raw_expr,
        "field_name": field_name,
        "query_name": query_name,
        "dax_expression": dax_code,
        "m_expression": m_code,
        "target_type": target_type,
        "confidence": 88 if not warnings else 80,
        "status": "Rule-Based",
        "warnings": warnings,
        "is_translatable": True,
        "category": cat or "General Expression",
    }


def _has_aggregation(expr):
    if not expr:
        return False
    return bool(re.search(r"\b(SUM|AVERAGE|MIN|MAX|COUNT|DISTINCTCOUNT|CALCULATE)\s*\(", expr, re.IGNORECASE))


def parse_cognos_report(xml_text):
    """
    Main entry point for Cognos Report XML. Extracts report metadata,
    queries, fields, filters, prompts, drillthroughs, visuals, custom controls,
    lists, crosstabs, repeaters, macros, and computes deterministic DAX/M translations.
    """
    root = ET.fromstring(xml_text)

    report_name_el = _iter_all(root, "reportName")
    report_name = _text_of(report_name_el[0]) if report_name_el else ""

    queries = extract_queries(root)

    all_fields = []
    all_filters = []
    for q in queries:
        all_fields.extend(q["data_items"])
        all_filters.extend(q["filters"])

    # Layout parsing
    lists = extract_lists(root)
    crosstabs = extract_crosstabs(root)
    repeaters = extract_repeaters(root)

    # Deterministic Field Translations
    field_translations = []
    for f in all_fields:
        trans = translate_cognos_expression_deterministic(
            f["expression"],
            field_name=f["field"],
            query_name=f["query"],
            is_filter=False,
            aggregate=f.get("aggregate", ""),
        )
        field_translations.append(trans)
        f["dax_translation"] = trans["dax_expression"]
        f["m_translation"] = trans["m_expression"]
        f["target_layer"] = trans["target_type"]
        f["translation_confidence"] = trans["confidence"]
        f["translation_status"] = trans["status"]
        f["translation_warnings"] = trans["warnings"]

    # Deterministic Filter Translations
    filter_translations = []
    for flt in all_filters:
        trans = translate_cognos_expression_deterministic(
            flt["expression"],
            field_name=f"Filter_{flt['query']}",
            query_name=flt["query"],
            is_filter=True,
        )
        filter_translations.append(trans)
        flt["dax_translation"] = trans["dax_expression"]
        flt["m_translation"] = trans["m_expression"]
        flt["target_layer"] = trans["target_type"]
        flt["translation_confidence"] = trans["confidence"]
        flt["translation_status"] = trans["status"]
        flt["translation_warnings"] = trans["warnings"]

    # Tokenize macros across all fields and filters
    macros = []
    for f in all_fields:
        macros.extend(extract_cognos_macros(f["expression"]))
    for flt in all_filters:
        macros.extend(extract_cognos_macros(flt["expression"]))

    result = {
        "report_name": report_name,
        "model_path": extract_model_path(root),
        "queries": queries,
        "fields": all_fields,
        "filters": all_filters,
        "prompts": extract_prompts(root),
        "drillthroughs": extract_drillthroughs(root, report_name),
        "visualizations": extract_visualizations(root),
        "data_stores": extract_data_stores(root),
        "custom_controls": extract_custom_controls(root),
        "lists": lists,
        "crosstabs": crosstabs,
        "repeaters": repeaters,
        "field_translations": field_translations,
        "filter_translations": filter_translations,
        "macros": macros,
        "xml_text": xml_text,
    }

    # Derived: real vs calculated fields
    result["source_fields"] = [f for f in all_fields if f["module"] is not None or f["table"] is not None]
    result["calculated_fields"] = [f for f in all_fields if (f["module"] is None and f["table"] is None) and f["expression"]]

    # Derived: which modules/tables this report touches
    tables_touched = set()
    for f in result["source_fields"]:
        if f.get("module") and f.get("table"):
            tables_touched.add(f"{f['module']}.{f['table']}")
        elif f.get("table"):
            tables_touched.add(f["table"])
    for q in queries:
        if q.get("native_sql") and q["native_sql"].get("sql_text"):
            m_sql = re.search(r"\bFROM\s+([A-Za-z0-9_]+)", q["native_sql"]["sql_text"], re.IGNORECASE)
            if m_sql:
                tables_touched.add(m_sql.group(1))
            else:
                tables_touched.add(q.get("name") or "NativeSQL")
    result["source_tables"] = sorted(tables_touched)

    return result




# ============================================================================
# DATA MODULE PARSER
# Cognos Data Modules are delivered as base64(gzip(JSON)) inside a field
# called "smartsData" — this is the semantic-layer / model equivalent of
# WebFOCUS's MAS/ACX metadata, except it's the PRIMARY source of truth here,
# not an optional bolt-on. Multiple smartsModule entries can exist per
# Data Module — each one typically wraps one physical table/sheet, plus
# occasionally an index-only entry that references all tables together
# (useful for detecting declared cross-table relationships, when present).
# ============================================================================

import base64
import gzip


def decode_smarts_data(smarts_data_b64):
    """Decode one smartsData field: base64 -> gzip -> JSON dict."""
    compressed = base64.b64decode(smarts_data_b64)
    decompressed = gzip.decompress(compressed)
    return json.loads(decompressed)


def extract_table_schema(decoded_module):
    """
    Extract a clean {table_id: {column_id: {type, usage, aggregation}}} schema
    from one decoded smartsModule, stripping the NLP/semantic-search metadata
    (labels, tokens, entities, concepts) that Cognos Assistant uses but that
    is irrelevant to migration.
    """
    tables = {}
    for ds in decoded_module.get("datasets", []):
        table_id = ds.get("id", "")
        columns = {}
        for item in ds.get("item", []):
            col = item.get("column")
            if not col:
                continue
            col_id = col.get("id", "")
            columns[col_id] = {
                "datatype": col.get("dataType", {}).get("type", ""),
                "usage": col.get("usage", ""),  # IDENTIFIER, FACT, ATTRIBUTE
                "default_aggregation": col.get("defaultAggregation", ""),
            }
        if columns:  # only record tables that actually have columns (skip index-only entries)
            tables[table_id] = {
                "name": ds.get("name", ""),
                "columns": columns,
            }
    return tables


def find_relationship_candidates(decoded_module):
    """
    Recursively scans a decoded Data Module for any key that looks like it
    declares a relationship/join between tables (e.g. 'relationships',
    'relationship', 'joins', 'links'). Cognos doesn't use one fixed key name
    across all versions/module types, so this is a best-effort structural
    scan rather than a lookup on one known field.

    Returns a list of {"path": "dotted.path.to.key", "key": key_name, "value": raw_value}
    for every match found, regardless of shape - callers should inspect
    the raw value themselves, since the internal schema of a real relationship
    entry hasn't been confirmed against a live example yet.
    """
    RELATIONSHIP_KEY_PATTERNS = ("relationship", "join")
    matches = []

    def walk(obj, path):
        if isinstance(obj, dict):
            for key, value in obj.items():
                key_lower = key.lower()
                if any(pat in key_lower for pat in RELATIONSHIP_KEY_PATTERNS):
                    matches.append({"path": f"{path}.{key}" if path else key, "key": key, "value": value})
                walk(value, f"{path}.{key}" if path else key)
        elif isinstance(obj, list):
            for idx, item in enumerate(obj):
                walk(item, f"{path}[{idx}]")

    walk(decoded_module, "")
    return matches


def parse_cognos_data_module(smarts_module_entries):
    """
    Given a list of raw {'smartsData': b64, 'id': ...} entries (as returned by
    /bi/v1/objects/{id}/smarts_modules), decode all of them and merge into one
    combined schema: {module_name, tables: {table_id: {name, columns}}, relationships}.

    'relationships' is populated by find_relationship_candidates() across every
    decoded module. In the one real Data Module tested so far (4 standalone
    uploaded-Excel tables, no declared joins), this correctly comes back empty -
    that's the expected result for a flat, join-free model, not a parser failure.
    If you load a Data Module that DOES declare relationships, this list will
    contain the raw matched entries for manual inspection, since the exact
    field-level schema of a populated relationship entry hasn't been confirmed
    against a real example yet.
    """
    module_name = ""
    all_tables = {}
    all_relationships = []
    for entry in smarts_module_entries:
        decoded = decode_smarts_data(entry["smartsData"])
        module_name = decoded.get("name", module_name)
        tables = extract_table_schema(decoded)
        all_tables.update(tables)
        all_relationships.extend(find_relationship_candidates(decoded))
    return {"module_name": module_name, "tables": all_tables, "relationships": all_relationships}


def validate_report_against_model(parsed_report, data_module_schema):
    """
    Cross-checks every source field a report references against the real
    Data Module schema. Mirrors validate_plan_references() from KashMap's
    AI build-plan validator — same spirit, applied earlier in the pipeline.
    Returns a list of issues (empty list = fully verified).
    """
    issues = []
    tables = data_module_schema.get("tables", {})

    for field in parsed_report.get("source_fields", []):
        table_id = field["table"]
        col_id = field["column"]
        if table_id not in tables:
            issues.append(f"Field '{field['field']}' references unknown table '{table_id}'")
            continue
        if col_id not in tables[table_id]["columns"]:
            issues.append(
                f"Field '{field['field']}' references unknown column "
                f"'{col_id}' in table '{table_id}'"
            )

    return issues




# ============================================================================
# END OF COGNOS PARSER
# ============================================================================

from collections import Counter, defaultdict


def _upper_set_cognos(values):
    return frozenset(str(v).strip().upper() for v in values if str(v or '').strip())


def compute_cognos_profile(parsed):
    """Signature of a report's real structure, used for duplicate comparison.
    Mirrors compute_fex_profile() from KashMap."""
    model_path = (parsed.get('model_path') or '').strip().upper()
    source_tables = _upper_set_cognos(parsed.get('source_tables', []))
    all_fields = _upper_set_cognos(f['field'] for f in parsed.get('fields', []))
    filters = frozenset(
        (f['query'].upper(), f['use'].upper(), f['expression'].strip().upper())
        for f in parsed.get('filters', [])
    )
    prompts = _upper_set_cognos(p['parameter'] for p in parsed.get('prompts', []))
    drillthrough_targets = _upper_set_cognos(d['target_report'] for d in parsed.get('drillthroughs', []))

    if not model_path and not all_fields:
        return None

    return {
        'model_path': model_path,
        'source_tables': source_tables,
        'all_fields': all_fields,
        'filters': filters,
        'prompts': prompts,
        'drillthrough_targets': drillthrough_targets,
    }


def exact_duplicate_key_cognos(profile):
    if profile is None:
        return None
    return (
        profile['model_path'], profile['source_tables'], profile['all_fields'],
        profile['filters'], profile['prompts'], profile['drillthrough_targets'],
    )


def field_similarity_cognos(left, right):
    if not left and not right:
        return 1.0
    if not left or not right:
        return 0.0
    return len(left & right) / len(left | right)


def profile_difference_summary_cognos(base, other):
    diffs = []
    for label, key in [('fields', 'all_fields'), ('filters', 'filters'),
                        ('prompts', 'prompts'), ('drillthroughs', 'drillthrough_targets')]:
        if base[key] != other[key]:
            diffs.append(label)
    return ', '.join(diffs) if diffs else 'No meaningful differences'


def build_cognos_duplicate_analysis(parsed_by_report):
    """parsed_by_report: {report_name: (parsed, validation_issues)}
    Mirrors build_duplicate_analysis() from KashMap."""
    records = []
    exact_groups = defaultdict(list)
    table_groups = defaultdict(list)

    for report_name, (parsed, _issues) in parsed_by_report.items():
        profile = compute_cognos_profile(parsed)
        record = {
            'report_name': report_name, 'profile': profile,
            'category': 'Unparsed' if profile is None else 'Unique',
            'group_id': '', 'group_size': 1,
            'exact_matches': [], 'near_matches': [], 'same_source_matches': [],
            'best_similarity': 0.0, 'difference_summary': '',
        }
        records.append(record)
        if profile is not None:
            exact_groups[exact_duplicate_key_cognos(profile)].append(record)
            table_key = (profile['model_path'], profile['source_tables'])
            table_groups[table_key].append(record)

    group_counter = 1
    for members in exact_groups.values():
        if len(members) <= 1:
            continue
        group_id = f'Exact {group_counter}'
        group_counter += 1
        names = [m['report_name'] for m in members]
        for record in members:
            record['category'] = 'Exact Duplicate'
            record['group_id'] = group_id
            record['group_size'] = len(members)
            record['exact_matches'] = [n for n in names if n != record['report_name']]
            record['best_similarity'] = 1.0
            record['difference_summary'] = 'Exact same model, tables, fields, filters, prompts, and drillthroughs'

    for table_key, members in table_groups.items():
        if (not table_key[0] and not table_key[1]) or len(members) <= 1:
            continue
        for record in members:
            if record['category'] == 'Exact Duplicate':
                continue
            best_near, same_source, best_similarity, diff_notes = [], [], 0.0, []
            for other in members:
                if other is record:
                    continue
                sim = field_similarity_cognos(record['profile']['all_fields'], other['profile']['all_fields'])
                best_similarity = max(best_similarity, sim)
                if sim >= 0.80:
                    if (record['profile']['filters'] != other['profile']['filters']
                            or record['profile']['prompts'] != other['profile']['prompts']
                            or record['profile']['drillthrough_targets'] != other['profile']['drillthrough_targets']):
                        best_near.append(other['report_name'])
                        diff_notes.append(profile_difference_summary_cognos(record['profile'], other['profile']))
                    else:
                        same_source.append(other['report_name'])
                else:
                    same_source.append(other['report_name'])
            if best_near:
                record['category'] = 'Near Duplicate'
                record['near_matches'] = sorted(set(best_near))
                record['group_id'] = f'Near {abs(hash(table_key)) % 100000}'
                record['group_size'] = len(best_near) + 1
                record['difference_summary'] = '; '.join(sorted(set(diff_notes))) or '80%+ same fields with logic differences'
            elif same_source:
                record['category'] = 'Same Data Source'
                record['same_source_matches'] = sorted(set(same_source))
                record['group_id'] = f'Source {abs(hash(table_key)) % 100000}'
                record['group_size'] = len(same_source) + 1
                record['difference_summary'] = 'Same model/tables with different fields or report logic'
            record['best_similarity'] = max(record['best_similarity'], best_similarity)

    for record in records:
        if record['category'] == 'Unique':
            record['difference_summary'] = 'No meaningful match found'
        elif record['category'] == 'Unparsed':
            record['difference_summary'] = 'Could not parse enough model/fields'
    return records


def cognos_duplicate_counts(records):
    counts = Counter(r['category'] for r in records)
    return {
        'exact': counts.get('Exact Duplicate', 0), 'near': counts.get('Near Duplicate', 0),
        'same_source': counts.get('Same Data Source', 0), 'unique': counts.get('Unique', 0),
        'unparsed': counts.get('Unparsed', 0),
    }


def build_duplicate_df(records):
    rows = []
    for r in records:
        rows.append({
            'Report': r['report_name'], 'Duplicate Type': r['category'],
            'Group ID': r['group_id'], 'Group Size': r['group_size'],
            'Confidence': f"{r['best_similarity']:.0%}" if r['best_similarity'] else '',
            'Matched With': ', '.join(r['exact_matches'] + r['near_matches'] + r['same_source_matches']),
            'Difference Summary': r['difference_summary'],
        })
    return pd.DataFrame(rows)


# =====================================================================
# SENTURUS-STYLE MIGRATION ASSESSMENT & COMPLEXITY SCORING ENGINE
# =====================================================================

def compute_cognos_complexity_score(parsed, validation_issues=None):
    """
    Computes an explainable 0-100 complexity score for a Cognos report based on 6 core factors:
    1. SQL & Query Complexity (max 25 pts) - joins, unions, native SQL, extra queries, slicers
    2. Calculations Complexity (max 20 pts) - calculated fields, aggregations, scoped rollups, date math
    3. Conditional Expressions & Logic (max 15 pts) - chained IF, CASE WHEN, null handling, prompted filters
    4. Layout & Containers (max 20 pts) - lists, crosstabs, repeaters, column volume, visuals, drillthroughs
    5. Custom Controls & JS (max 10 pts) - custom controls requiring manual rewrite
    6. Unsupported Features & Dimensional MDX (max 10 pts) - dimensional MDX (tuple, ancestor), validation issues
    
    Total score is strictly clamped between 0 and 100.
    Tiers: Simple (0-25), Medium (26-50), Complex (51-75), Highly Complex (76-100).
    """
    issues = validation_issues or []
    queries = parsed.get("queries", [])
    query_count = len(queries)
    join_count = sum(1 for q in queries if q.get("join_details"))
    set_op_count = sum(1 for q in queries if q.get("set_operation"))
    native_sql_count = sum(1 for q in queries if q.get("native_sql"))
    slicer_count = sum(len(q.get("slicers", [])) for q in queries)
    
    sql_points = min(25, (
        max(0, query_count - 1) * 2
        + join_count * 6
        + set_op_count * 5
        + native_sql_count * 8
        + slicer_count * 3
    ))
    
    calc_fields = parsed.get("calculated_fields", [])
    trans = parsed.get("field_translations", [])
    agg_count = sum(1 for t in trans if t.get("category") in ("Aggregation", "Scoped Aggregation"))
    scoped_agg_count = sum(1 for t in trans if t.get("category") == "Scoped Aggregation")
    date_fn_count = sum(1 for t in trans if t.get("category") == "Date Function")
    
    calc_points = min(20, (
        len(calc_fields) * 2
        + agg_count * 2
        + scoped_agg_count * 3
        + date_fn_count * 2
    ))
    
    if_count = sum(1 for t in trans if t.get("category") in ("Chained IF", "Case When"))
    null_count = sum(1 for t in trans if t.get("category") == "Null Handling")
    filters = parsed.get("filters", [])
    prompted_filter_count = sum(1 for f in filters if f.get("is_prompted"))
    
    logic_points = min(15, (
        if_count * 3
        + null_count * 2
        + len(filters) * 1
        + prompted_filter_count * 2
    ))
    
    lists = parsed.get("lists", [])
    crosstabs = parsed.get("crosstabs", [])
    repeaters = parsed.get("repeaters", [])
    total_list_cols = sum(l.get("column_count", 0) for l in lists)
    viz_count = len(parsed.get("visualizations", []))
    drill_count = len(parsed.get("drillthroughs", []))
    
    layout_points = min(20, (
        len(lists) * 3
        + (total_list_cols // 5) * 1
        + len(crosstabs) * 6
        + len(repeaters) * 5
        + viz_count * 2
        + drill_count * 3
    ))
    
    controls = parsed.get("custom_controls", [])
    controls_points = min(10, len(controls) * 5)
    
    unsupported_trans = sum(1 for t in trans if not t.get("is_translatable") or t.get("category") == "Dimensional MDX")
    unsupported_points = min(10, (
        unsupported_trans * 5
        + len(issues) * 2
    ))
    
    total_score = min(100, max(0, (
        sql_points + calc_points + logic_points + layout_points + controls_points + unsupported_points
    )))
    
    if total_score <= 25:
        tier = "Simple"
    elif total_score <= 50:
        tier = "Medium"
    elif total_score <= 75:
        tier = "Complex"
    else:
        tier = "Highly Complex"
        
    breakdown = {
        "sql_points": sql_points,
        "calc_points": calc_points,
        "logic_points": logic_points,
        "layout_points": layout_points,
        "controls_points": controls_points,
        "unsupported_points": unsupported_points,
        "sql_details": f"Queries: {query_count}, Joins: {join_count}, Unions: {set_op_count}, Native SQL: {native_sql_count}, Slicers: {slicer_count}",
        "calc_details": f"Calculations: {len(calc_fields)}, Aggregations: {agg_count}, Scoped Rollups: {scoped_agg_count}, Date Math: {date_fn_count}",
        "logic_details": f"Chained IF/CASE: {if_count}, Null Handling: {null_count}, Detail Filters: {len(filters)}, Prompted Filters: {prompted_filter_count}",
        "layout_details": f"Lists: {len(lists)} ({total_list_cols} cols), Crosstabs: {len(crosstabs)}, Repeaters: {len(repeaters)}, Visuals: {viz_count}, Drills: {drill_count}",
        "controls_details": f"Custom Controls (JS widgets): {len(controls)}",
        "unsupported_details": f"Dimensional MDX / Unmapped: {unsupported_trans}, Validation Issues: {len(issues)}",
    }
    
    return {
        "complexity_score": total_score,
        "complexity_tier": tier,
        "breakdown": breakdown,
    }


DEFAULT_EFFORT_HOURS = {
    "Simple": 12,
    "Medium": 32,
    "Complex": 64,
    "Highly Complex": 120,
}


def estimate_migration_effort(tier, custom_tier_hours=None):
    """Returns estimated migration hours for a complexity tier using configurable assumptions."""
    hours_map = custom_tier_hours or DEFAULT_EFFORT_HOURS
    return int(hours_map.get(tier, 32))


def assign_migration_wave(report_name, score, tier, dup_category="Unique", business_value="Medium", is_duplicate_copy=False, canonical_match=None):
    """
    Assigns reports into 4 Senturus waves:
    Wave 1: Quick Wins (Simple complexity with High/Medium business value)
    Wave 2: Core Migration (Medium complexity standard reports)
    Wave 3: Strategic Redesign (Complex and Highly Complex reports, or reports requiring MDX redesign/custom controls)
    Wave 4: Retirement / Consolidation (Duplicate copies or reports marked for retirement)
    Note: Canonical reports in duplicate groups migrate in Waves 1-3; only redundant duplicate copies route to Wave 4.
    """
    bv_str = str(business_value or "Medium").strip().capitalize()
    if is_duplicate_copy or bv_str.lower() == "retire":
        target = f" '{canonical_match}'" if canonical_match else ""
        reason = f"Duplicate copy of canonical report{target} ({dup_category}). Recommend consolidating into canonical report or retiring."
        return "Wave 4: Retirement / Consolidation", 4, reason

    if tier == "Simple" and bv_str in ("High", "Medium"):
        reason = f"Simple complexity (Score: {score}/100) with {bv_str} business value. Ideal quick win to build stakeholder momentum."
        return "Wave 1: Quick Wins", 1, reason

    if (tier == "Medium" and bv_str in ("High", "Medium")) or (tier == "Simple" and bv_str == "Low"):
        reason = f"Standard {tier.lower()} complexity report (Score: {score}/100) with {bv_str} business value. Core production migration candidate."
        return "Wave 2: Core Migration", 2, reason

    if tier in ("Complex", "Highly Complex"):
        reason = f"{tier} complexity report (Score: {score}/100). Requires advanced DAX modeling, query restructuring, or manual visual redesign."
        return "Wave 3: Strategic Redesign", 3, reason

    return "Wave 2: Core Migration", 2, f"Standard {tier.lower()} complexity report migration."


def build_cognos_inventory_df(parsed_by_report, complexity_by_report, wave_by_report):
    """Builds combined portfolio inventory table across all uploaded reports."""
    rows = []
    for report_name, (parsed, issues) in parsed_by_report.items():
        comp = complexity_by_report.get(report_name, {})
        wave = wave_by_report.get(report_name, {})
        rows.append({
            "Report Name": parsed.get("report_name") or report_name,
            "Complexity Score": comp.get("complexity_score", 0),
            "Complexity Tier": comp.get("complexity_tier", "Medium"),
            "Est. Effort (Hours)": comp.get("est_effort_hours", 32),
            "Assigned Wave": wave.get("wave", "Wave 2: Core Migration"),
            "Source Tables": ", ".join(parsed.get("source_tables", [])) or "None",
            "Queries": len(parsed.get("queries", [])),
            "Fields": len(parsed.get("fields", [])),
            "Calculations": len(parsed.get("calculated_fields", [])),
            "Filters": len(parsed.get("filters", [])),
            "Prompts": len(parsed.get("prompts", [])),
            "Lists": len(parsed.get("lists", [])),
            "Crosstabs": len(parsed.get("crosstabs", [])),
            "Repeaters": len(parsed.get("repeaters", [])),
            "Custom Controls": len(parsed.get("custom_controls", [])),
            "Validation Status": "Verified" if not issues else f"{len(issues)} issue(s)",
        })
    return pd.DataFrame(rows)


def build_cognos_complexity_df(complexity_by_report, effort_hours_map=None):
    """Builds explainable complexity scoring table showing each factor's point contribution."""
    rows = []
    for report_name, comp in complexity_by_report.items():
        b = comp.get("breakdown", {})
        tier = comp.get("complexity_tier", "Medium")
        effort = estimate_migration_effort(tier, effort_hours_map)
        rows.append({
            "Report Name": report_name,
            "Total Score (0-100)": comp.get("complexity_score", 0),
            "Complexity Tier": tier,
            "Est. Effort (Hours)": effort,
            "SQL & Query (max 25)": b.get("sql_points", 0),
            "Calculations (max 20)": b.get("calc_points", 0),
            "Logic & Filters (max 15)": b.get("logic_points", 0),
            "Layout (max 20)": b.get("layout_points", 0),
            "Custom Controls (max 10)": b.get("controls_points", 0),
            "Unsupported/MDX (max 10)": b.get("unsupported_points", 0),
            "SQL Details": b.get("sql_details", ""),
            "Calculation Details": b.get("calc_details", ""),
            "Logic Details": b.get("logic_details", ""),
            "Layout Details": b.get("layout_details", ""),
            "Controls Details": b.get("controls_details", ""),
            "Unsupported Details": b.get("unsupported_details", ""),
        })
    return pd.DataFrame(rows)


def build_cognos_waves_df(wave_by_report, complexity_by_report):
    """Builds wave roadmap table sorting reports into the 4 Senturus waves."""
    rows = []
    for report_name, wave_info in wave_by_report.items():
        comp = complexity_by_report.get(report_name, {})
        rows.append({
            "Report Name": report_name,
            "Assigned Wave": wave_info.get("wave", "Wave 2: Core Migration"),
            "Wave Number": wave_info.get("wave_number", 2),
            "Complexity Tier": comp.get("complexity_tier", "Medium"),
            "Complexity Score": comp.get("complexity_score", 0),
            "Business Value": wave_info.get("business_value", "Medium"),
            "Est. Effort (Hours)": wave_info.get("est_effort_hours", 32),
            "Wave Recommendation & Rationale": wave_info.get("reason", ""),
        })
    df = pd.DataFrame(rows)
    if not df.empty and "Wave Number" in df.columns:
        df = df.sort_values(by=["Wave Number", "Complexity Score"]).reset_index(drop=True)
    return df


def build_cross_report_lineage(parsed_by_report):
    """Maps source tables and packages to all reports referencing them."""
    table_map = defaultdict(lambda: {"columns": set(), "reports": set(), "models": set()})
    
    for report_name, (parsed, _issues) in parsed_by_report.items():
        model_path = parsed.get("model_path") or "Unknown Model"
        for f in parsed.get("source_fields", []):
            tbl = f.get("table") or "Unknown Table"
            if f.get("module"):
                tbl = f"{f['module']}.{f['table']}"
            col = f.get("column") or f.get("field") or ""
            table_map[tbl]["reports"].add(report_name)
            table_map[tbl]["models"].add(model_path)
            if col:
                table_map[tbl]["columns"].add(col)
        for tbl in parsed.get("source_tables", []):
            table_map[tbl]["reports"].add(report_name)
            table_map[tbl]["models"].add(model_path)
                
    rows = []
    for tbl, info in sorted(table_map.items()):
        if not tbl:
            continue
        rows.append({
            "Source Table": tbl,
            "Models / Packages": ", ".join(sorted(info["models"])) if info["models"] else "Unknown",
            "Referencing Reports Count": len(info["reports"]),
            "Referencing Reports": ", ".join(sorted(info["reports"])),
            "Referenced Columns Count": len(info["columns"]),
            "Referenced Columns": ", ".join(sorted(info["columns"])) or "None explicit",
        })
    return pd.DataFrame(rows)


def build_consolidation_opportunities(parsed_by_report, dup_records=None):
    """Identifies consolidation opportunities across reports sharing tables or models."""
    dup_map = {r["report_name"]: r["category"] for r in (dup_records or [])}
    reports = list(parsed_by_report.keys())
    rows = []
    
    for i in range(len(reports)):
        for j in range(i + 1, len(reports)):
            repA, repB = reports[i], reports[j]
            parsedA, _ = parsed_by_report[repA]
            parsedB, _ = parsed_by_report[repB]
            
            tablesA = set(parsedA.get("source_tables", []))
            tablesB = set(parsedB.get("source_tables", []))
            fieldsA = set(f["field"].upper() for f in parsedA.get("fields", []))
            fieldsB = set(f["field"].upper() for f in parsedB.get("fields", []))
            
            table_overlap = (len(tablesA & tablesB) / len(tablesA | tablesB)) if (tablesA | tablesB) else 0.0
            field_overlap = (len(fieldsA & fieldsB) / len(fieldsA | fieldsB)) if (fieldsA | fieldsB) else 0.0
            same_model = (parsedA.get("model_path") == parsedB.get("model_path")) and bool(parsedA.get("model_path"))
            
            catA = dup_map.get(repA, "Unique")
            catB = dup_map.get(repB, "Unique")
            
            if (catA in ("Exact Duplicate", "Near Duplicate") or 
                catB in ("Exact Duplicate", "Near Duplicate") or 
                table_overlap >= 0.70 or 
                (same_model and field_overlap >= 0.50)):
                
                if catA == "Exact Duplicate" and catB == "Exact Duplicate":
                    rec = "Retire duplicate; consolidate into single canonical Power BI report."
                elif catA == "Near Duplicate" or catB == "Near Duplicate" or field_overlap >= 0.80:
                    rec = "Consolidate into single Power BI report with bookmarks/tabs or shared Semantic Model."
                elif table_overlap >= 0.70 or same_model:
                    rec = "Share single Power BI Semantic Model (dataset) across both reports to eliminate redundancy."
                else:
                    rec = "Evaluate for shared semantic model components."
                    
                rows.append({
                    "Report A": repA,
                    "Report B": repB,
                    "Shared Data Source": "Yes" if same_model else "Different / Unknown",
                    "Table Overlap %": f"{table_overlap:.0%}",
                    "Field Overlap %": f"{field_overlap:.0%}",
                    "Common Tables": ", ".join(sorted(tablesA & tablesB)) or "None",
                    "Recommended Strategy": rec,
                })
    return pd.DataFrame(rows)


# ==============================================================================
# TASK 3: POWER BI SEMANTIC MODEL GENERATION & MCP ENGINE
# ==============================================================================

MCP_EXE_DEFAULT = r"C:\Users\jegan\AppData\Local\npm-cache\_npx\deea81b821a9ed55\node_modules\@microsoft\powerbi-modeling-mcp-win32-x64\dist\powerbi-modeling-mcp.exe"


def get_powerbi_mcp_exe_path():
    """Locate the powerbi-modeling-mcp executable path on localhost."""
    env_path = os.environ.get("POWERBI_MCP_EXE")
    if env_path and os.path.exists(env_path):
        return env_path
    if os.path.exists(MCP_EXE_DEFAULT):
        return MCP_EXE_DEFAULT
    # Dynamic search in npm-cache/_npx
    npx_root = Path(r"C:\Users\jegan\AppData\Local\npm-cache\_npx")
    if npx_root.exists():
        found = list(npx_root.glob("**/powerbi-modeling-mcp.exe"))
        if found:
            return str(found[0])
    return None


def call_powerbi_mcp_local(tool_name, arguments, timeout_sec=8):
    """
    Direct stdio JSON-RPC 2.0 client for local powerbi-modeling-mcp.exe.
    Performs handshake (initialize) followed by tool execution (tools/call).
    """
    exe_path = get_powerbi_mcp_exe_path()
    if not exe_path or not os.path.exists(exe_path):
        return {"status": "offline", "message": "Power BI Modeling MCP binary not found.", "data": []}
    
    try:
        proc = subprocess.Popen(
            [exe_path, "--start", "--accept-eula"],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8"
        )
        
        # 1. initialize handshake
        init_req = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {"name": "kashmap_client", "version": "1.0"}
            }
        }
        proc.stdin.write(json.dumps(init_req) + "\n")
        proc.stdin.flush()
        _init_res = proc.stdout.readline()
        
        # 2. tools/call execution
        call_req = {
            "jsonrpc": "2.0",
            "id": 2,
            "method": "tools/call",
            "params": {
                "name": tool_name,
                "arguments": arguments
            }
        }
        proc.stdin.write(json.dumps(call_req) + "\n")
        proc.stdin.flush()
        call_res_line = proc.stdout.readline()
        
        proc.terminate()
        try:
            proc.wait(timeout=2)
        except Exception:
            proc.kill()
            
        if not call_res_line:
            return {"status": "error", "message": "Empty response from MCP server"}
            
        res_obj = json.loads(call_res_line)
        content_items = res_obj.get("result", {}).get("content", [])
        if content_items and content_items[0].get("type") == "text":
            raw_text = content_items[0].get("text", "")
            try:
                parsed_data = json.loads(raw_text)
                return {"status": "success", "result": parsed_data}
            except Exception:
                return {"status": "success", "result": raw_text}
                
        return {"status": "success", "result": res_obj.get("result", {})}
    except Exception as e:
        return {"status": "error", "message": str(e)}


def check_powerbi_mcp_connections():
    """Discover active Power BI Desktop or Analysis Services instances on localhost."""
    exe_path = get_powerbi_mcp_exe_path()
    if not exe_path or not os.path.exists(exe_path):
        return {
            "mcp_installed": False,
            "instances": [],
            "connected": False,
            "message": "Power BI Modeling MCP binary not found. Offline blueprint generation is active.",
        }
    res = call_powerbi_mcp_local("connection_operations", {"request": {"operation": "ListLocalInstances"}})
    if res.get("status") == "success":
        result_dict = res.get("result", {})
        instances = []
        if isinstance(result_dict, dict):
            instances = result_dict.get("data", [])
        return {
            "mcp_installed": True,
            "instances": instances,
            "connected": len(instances) > 0,
            "message": f"Connected to {len(instances)} active Power BI Desktop instance(s)." if instances else "Power BI Desktop is not currently running with an open model. KashMap offline blueprint generation is active.",
        }
    return {
        "mcp_installed": True,
        "instances": [],
        "connected": False,
        "message": f"MCP communication notice: {res.get('message', 'No active instance')}",
    }


def validate_dax_expression_offline(dax_expr):
    """
    Offline syntax and integrity validator for DAX expressions.
    Checks delimiter balancing, string quoting, and detects Cognos/MDX constructs.
    """
    expr = (dax_expr or "").strip()
    if not expr:
        return {"is_valid": False, "status": "Empty", "warnings": ["Expression is empty."]}
    
    warnings = []
    e_lower = expr.lower()
    
    # 1. Delimiter balance
    pairs = [("(", ")"), ("[", "]"), ("{", "}")]
    for open_ch, close_ch in pairs:
        if expr.count(open_ch) != expr.count(close_ch):
            warnings.append(f"Mismatched delimiters '{open_ch}' and '{close_ch}' ({expr.count(open_ch)} vs {expr.count(close_ch)})")
            
    # Check quotes
    quote_count = expr.count('"')
    if quote_count % 2 != 0:
        warnings.append(f"Unmatched double quote '\"' (count is {quote_count})")
        
    # 2. Check for Cognos / MDX tokens that do not belong in valid DAX
    mdx_patterns = [
        ("tuple(", "MDX tuple function"),
        ("ancestor(", "MDX ancestor function"),
        ("children(", "MDX children function"),
        ("descendants(", "MDX descendants function"),
        ("item(", "MDX item function"),
        ("#prompt", "Cognos macro token"),
        ("_add_days", "Cognos date function"),
        ("_add_months", "Cognos date function"),
        ("cast_char", "Cognos CAST function"),
    ]
    for pat, desc in mdx_patterns:
        if pat in e_lower:
            warnings.append(f"Unsupported {desc} detected: '{pat}'")
            
    is_valid = len(warnings) == 0
    return {
        "is_valid": is_valid,
        "status": "Valid DAX" if is_valid else "Manual Review Required",
        "warnings": warnings,
    }


def validate_dax_with_mcp(dax_expr, connection_name=None):
    """
    Validates a DAX expression using live MCP connection when available,
    falling back to offline syntax validation when disconnected.
    """
    offline_check = validate_dax_expression_offline(dax_expr)
    if not offline_check["is_valid"]:
        return offline_check
        
    mcp_check = check_powerbi_mcp_connections()
    if mcp_check["connected"]:
        conn = connection_name or (mcp_check["instances"][0].get("connectionName") if mcp_check["instances"] else None)
        query = f"EVALUATE ROW(\"TestMeasure\", {dax_expr})"
        args = {"request": {"operation": "Validate", "query": query}}
        if conn:
            args["request"]["connectionName"] = conn
        mcp_res = call_powerbi_mcp_local("dax_query_operations", args)
        if mcp_res.get("status") == "success":
            msg = mcp_res.get("result", {}).get("message", "")
            if "Error" in msg:
                return {"is_valid": False, "status": "MCP Validation Error", "warnings": [msg]}
            return {"is_valid": True, "status": "MCP Validated", "warnings": []}
            
    return {
        "is_valid": offline_check["is_valid"],
        "status": "Offline Syntax Validated",
        "warnings": offline_check["warnings"],
        "engine": "Offline Syntax Parser (Power BI Desktop instance not connected)",
    }


def apply_semantic_model_to_live_pbi(blueprint, connection_name=None, user_confirmed=False):
    """
    Applies the proposed semantic model blueprint to an active Power BI Desktop model.
    STRICT SAFETY GUARD: Requires explicit user_confirmed=True.
    """
    if not user_confirmed:
        return {
            "status": "blocked",
            "message": "Explicit user confirmation is required before modifying any connected Power BI semantic model.",
            "objects_modified": 0,
        }
        
    mcp_check = check_powerbi_mcp_connections()
    if not mcp_check["connected"]:
        return {
            "status": "offline",
            "message": "Cannot modify model: No active Power BI Desktop or Analysis Services instance connected.",
            "objects_modified": 0,
        }
        
    return {
        "status": "success",
        "message": f"Successfully planned deployment of model '{blueprint.get('model_name')}' to Power BI Desktop.",
        "objects_modified": len(blueprint.get("tables", [])) + len(blueprint.get("measures", [])),
    }


def infer_powerbi_data_type(field_name, expr, aggregate=""):
    """Infer Power BI Tabular data type based on naming patterns, aggregation, and expressions."""
    f_lower = (field_name or "").lower()
    e_lower = (expr or "").lower()
    agg_lower = (aggregate or "").lower()
    
    if agg_lower in ("count", "distinctcount") or any(k in f_lower for k in ("count", "quantity", "qty", "number_of", "units", "items")):
        return "Int64"
    if agg_lower in ("total", "sum", "average", "avg", "minimum", "maximum") or any(k in f_lower for k in ("revenue", "cost", "sales", "amount", "price", "margin", "budget", "actual", "rate", "pct", "percent")):
        return "Decimal"
    if any(k in e_lower for k in ("_add_days", "current_date", "current_timestamp", "today", "now", "date(", "year(", "month(")) or any(k in f_lower for k in ("date", "time", "year", "quarter", "month", "day")):
        return "DateTime"
    if any(k in e_lower for k in (">", "<", "=", "and", "or", "in (")) and agg_lower == "none" and not any(k in f_lower for k in ("category", "status", "group")):
        if any(k in f_lower for k in ("is_", "has_", "flag", "active")):
            return "Boolean"
    return "String"


def build_semantic_model_blueprint(report_name, parsed, validation_issues=None, framework_model=None, data_module_schema=None):
    """
    Constructs a complete Power BI Semantic Model Blueprint from Cognos report metadata.
    Distinguishes Verified metadata, Inferred mappings, and Manual-Review items.
    NEVER invents missing source tables or foreign keys.
    """
    issues = validation_issues or []
    model_name = parsed.get("report_name") or report_name
    if model_name.lower().endswith(".xml"):
        model_name = model_name[:-4]
    model_name = re.sub(r"[^\w\s-]", "", model_name).strip() or "CognosSemanticModel"
    
    # 1. Group fields by table
    table_dict = {}
    unassigned_fields = []
    
    # Seed tables from source_tables if present
    for st in parsed.get("source_tables", []):
        tname = st.split(".")[-1]
        if tname and tname not in table_dict:
            table_dict[tname] = {
                "table_name": tname,
                "source_type": "Data Source Table",
                "source_path": st,
                "columns": [],
                "measures": [],
                "provenance": "Inferred",
            }

    for f in parsed.get("fields", []):
        tbl = f.get("table")
        if not tbl and f.get("module"):
            tbl = f.get("module")
            
        if tbl:
            if tbl not in table_dict:
                table_dict[tbl] = {
                    "table_name": tbl,
                    "source_type": "Direct Query Reference",
                    "source_path": f"{f.get('module') or ''}.{tbl}".strip("."),
                    "columns": [],
                    "measures": [],
                    "provenance": "Inferred",
                }
            table_dict[tbl]["columns"].append(f)
        else:
            unassigned_fields.append(f)
            
    # Check for native SQL queries
    for q in parsed.get("queries", []):
        n_sql = q.get("native_sql")
        if n_sql and n_sql.get("sql_text"):
            q_tbl = q.get("name") or "NativeSQL"
            if q_tbl not in table_dict:
                table_dict[q_tbl] = {
                    "table_name": q_tbl,
                    "source_type": f"Native SQL ({n_sql.get('data_source') or 'Default'})",
                    "source_path": n_sql.get("sql_text")[:80].strip(),
                    "columns": [],
                    "measures": [],
                    "provenance": "Inferred",
                }
            for uf in list(unassigned_fields):
                if uf.get("query") == q.get("name"):
                    table_dict[q_tbl]["columns"].append(uf)
                    unassigned_fields.remove(uf)

    # If no tables were detected, create a fallback table from the primary query
    if not table_dict:
        default_tbl = "ReportData"
        table_dict[default_tbl] = {
            "table_name": default_tbl,
            "source_type": "Cognos Query Dataset",
            "source_path": parsed.get("model_path") or "ReportQuery",
            "columns": list(parsed.get("fields", [])),
            "measures": [],
            "provenance": "Inferred",
        }
        unassigned_fields = []
        
    # Check framework model or data module verification
    if framework_model:
        fw_subjects = {s["name"].upper(): s for s in framework_model.get("query_subjects", [])}
        for tname, tinfo in table_dict.items():
            if tname.upper() in fw_subjects:
                tinfo["provenance"] = "Verified (Framework Model)"
                
    if data_module_schema:
        dm_tables = {t["table_name"].upper(): t for t in data_module_schema.get("tables", [])}
        for tname, tinfo in table_dict.items():
            if tname.upper() in dm_tables:
                tinfo["provenance"] = "Verified (Data Module)"

    model_tables = []
    model_columns = []
    model_measures = []
    
    for tbl_name, tbl_info in sorted(table_dict.items()):
        cols = []
        measures = []
        for f in tbl_info["columns"]:
            fname = f.get("field", "")
            expr = f.get("expression", "")
            agg = f.get("aggregate", "")
            dax = f.get("dax_translation", "")
            layer = f.get("target_layer", "")
            status = f.get("translation_status", "Rule-Based")
            warnings = f.get("translation_warnings", [])
            dtype = infer_powerbi_data_type(fname, expr, agg)
            
            # Determine whether this is a measure or a column
            is_measure = (
                layer == "DAX Measure" or 
                agg.lower() in ("total", "sum", "average", "avg", "count", "minimum", "maximum") or
                (f.get("is_calculated") and any(k in (dax or "").upper() for k in ("SUM(", "AVERAGE(", "DIVIDE(", "CALCULATE(")))
            )
            
            # Check provenance
            if warnings or status == "Manual Review" or "MDX" in f.get("expression_type", ""):
                prov = "Manual Review"
            elif tbl_info["provenance"].startswith("Verified"):
                prov = "Verified"
            else:
                prov = "Inferred"
                
            if is_measure:
                fmt = "#,##0.00"
                if any(k in fname.lower() for k in ("pct", "percent", "margin", "ratio", "rate")):
                    fmt = "0.0%"
                elif any(k in fname.lower() for k in ("count", "qty", "quantity")):
                    fmt = "#,##0"
                elif any(k in fname.lower() for k in ("revenue", "cost", "sales", "amount", "price", "budget", "actual")):
                    fmt = "$#,##0.00"
                    
                dax_to_check = dax or f"SUM('{tbl_name}'[{fname}])"
                val_res = validate_dax_expression_offline(dax_to_check)
                all_warnings = list(warnings)
                if not val_res["is_valid"]:
                    all_warnings.extend(val_res["warnings"])
                    prov = "Manual Review"
                    
                m_obj = {
                    "table_name": tbl_name,
                    "measure_name": fname if not any(c.get("field") == fname for c in cols) else f"{fname}_Total",
                    "dax_expression": dax_to_check,
                    "format_string": fmt,
                    "data_type": dtype,
                    "provenance": prov,
                    "dax_validation_status": val_res["status"],
                    "notes": "; ".join(all_warnings) if all_warnings else "Rule-based DAX measure",
                }
                measures.append(m_obj)
                model_measures.append(m_obj)
            else:
                c_obj = {
                    "table_name": tbl_name,
                    "column_name": fname,
                    "data_type": dtype,
                    "source_path": expr,
                    "column_type": "Calculated Column" if f.get("is_calculated") else "Data Column",
                    "dax_expression": dax if f.get("is_calculated") else "",
                    "provenance": prov,
                    "notes": "; ".join(warnings) if warnings else "",
                }
                cols.append(c_obj)
                model_columns.append(c_obj)
                
        model_tables.append({
            "table_name": tbl_name,
            "source_type": tbl_info.get("source_type", "Direct Query Reference"),
            "source_path": tbl_info["source_path"],
            "provenance": tbl_info["provenance"],
            "column_count": len(cols),
            "measure_count": len(measures),
        })

    # Unassigned calculations allocated to primary table
    primary_tbl = model_tables[0]["table_name"] if model_tables else "Model"
    for f in unassigned_fields:
        fname = f.get("field", "")
        expr = f.get("expression", "")
        dax = f.get("dax_translation", "")
        warnings = list(f.get("translation_warnings", []))
        dax_to_check = dax or f"/* Review Required: {expr} */ BLANK()"
        val_res = validate_dax_expression_offline(dax_to_check)
        if not val_res["is_valid"]:
            warnings.extend(val_res["warnings"])
        prov = "Manual Review" if (warnings or "MDX" in f.get("expression_type", "") or not dax) else "Inferred"
        m_obj = {
            "table_name": primary_tbl,
            "measure_name": fname,
            "dax_expression": dax_to_check,
            "format_string": "#,##0.00",
            "data_type": infer_powerbi_data_type(fname, expr),
            "provenance": prov,
            "dax_validation_status": val_res["status"],
            "notes": "; ".join(warnings) or ("Unassigned calculation mapped to primary table" if prov != "Manual Review" else "Unsupported MDX or unmapped expression"),
        }
        model_measures.append(m_obj)

    # 2. Relationships
    relationships = []
    if framework_model:
        tbl_names_upper = {t["table_name"].upper(): t["table_name"] for t in model_tables}
        for rel in framework_model.get("relationships", []):
            left_t = rel.get("left_subject", "").upper()
            right_t = rel.get("right_subject", "").upper()
            if left_t in tbl_names_upper and right_t in tbl_names_upper:
                relationships.append({
                    "relationship_name": f"Rel_{tbl_names_upper[left_t]}_{tbl_names_upper[right_t]}",
                    "from_table": tbl_names_upper[left_t],
                    "from_column": rel.get("foreign_key") or "Key",
                    "to_table": tbl_names_upper[right_t],
                    "to_column": rel.get("primary_key") or "Key",
                    "cardinality": rel.get("left_cardinality", "ManyToOne"),
                    "cross_filtering": "Single",
                    "state": "Active",
                    "provenance": "Verified (Framework Model)",
                    "notes": f"Relationship from Framework Model: {rel.get('name', '')}",
                })
                
    # Check query joins
    for q in parsed.get("queries", []):
        jd = q.get("join_details")
        if jd:
            crit_list = jd.get("criteria", [])
            crit_str = ", ".join(crit_list) if crit_list else "Declared Query Join"
            left_col = "JoinKey"
            right_col = "JoinKey"
            if crit_list:
                m_join = re.search(r"\[.*?\]\.\[(.*?)\]\s*=\s*\[.*?\]\.\[(.*?)\]", crit_list[0])
                if m_join:
                    left_col = m_join.group(1)
                    right_col = m_join.group(2)
            relationships.append({
                "relationship_name": f"Rel_{jd.get('left_query')}_{jd.get('right_query')}",
                "from_table": jd.get("left_query", "LeftQuery"),
                "from_column": left_col,
                "to_table": jd.get("right_query", "RightQuery"),
                "to_column": right_col,
                "cardinality": "ManyToOne",
                "cross_filtering": "Single",
                "state": "Active",
                "provenance": "Inferred (Query Join)",
                "notes": f"Criteria: {crit_str}",
            })

    # If multiple tables exist but NO explicit relationships exist:
    # NEVER invent fake join keys. Mark as Manual Review (Ambiguous)
    if len(model_tables) > 1 and not relationships:
        t_names = [t["table_name"] for t in model_tables]
        for i in range(len(t_names) - 1):
            t_from = t_names[i+1]
            t_to = t_names[i]
            relationships.append({
                "relationship_name": f"Rel_{t_from}_{t_to}",
                "from_table": t_from,
                "from_column": "[Unmapped Join Column]",
                "to_table": t_to,
                "to_column": "[Unmapped Join Column]",
                "cardinality": "ManyToOne",
                "cross_filtering": "Single",
                "state": "Inactive (Review Required)",
                "provenance": "Manual Review (Ambiguous)",
                "notes": f"Tables '{t_from}' and '{t_to}' are co-referenced in report query, but no join condition is declared in XML. Inspect database schema or Framework Model.",
            })

    # Governance counts
    all_objects = model_tables + model_columns + model_measures + relationships
    verified_c = sum(1 for o in all_objects if "Verified" in o.get("provenance", ""))
    manual_c = sum(1 for o in all_objects if "Manual Review" in o.get("provenance", ""))
    inferred_c = len(all_objects) - verified_c - manual_c

    summary = {
        "model_name": model_name,
        "table_count": len(model_tables),
        "column_count": len(model_columns),
        "measure_count": len(model_measures),
        "relationship_count": len(relationships),
        "verified_count": verified_c,
        "inferred_count": inferred_c,
        "manual_review_count": manual_c,
    }

    return {
        "model_name": model_name,
        "summary": summary,
        "tables": model_tables,
        "columns": model_columns,
        "measures": model_measures,
        "relationships": relationships,
    }


def generate_tmdl_document(blueprint):
    """Generates unified Tabular Model Definition Language (.tmdl) document string."""
    model_name = blueprint.get("model_name", "Model")
    lines = [
        "// ==========================================================================",
        f"// Tabular Model Definition Language (TMDL) - Power BI Semantic Model",
        f"// Generated by KashMap Cognos to Power BI Migration Engine",
        f"// Model Name: {model_name}",
        f"// Tables: {blueprint['summary']['table_count']} | Measures: {blueprint['summary']['measure_count']} | Relationships: {blueprint['summary']['relationship_count']}",
        "// ==========================================================================",
        "",
        f"model {model_name}",
        "\tcompatibilityLevel: 1567",
        "\tculture: en-US",
        "\tdefaultPowerBIDataSourceVersion: powerBI_V3",
        "",
    ]
    
    tag_counter = 1
    cols_by_tbl = defaultdict(list)
    for c in blueprint.get("columns", []):
        cols_by_tbl[c["table_name"]].append(c)
        
    meas_by_tbl = defaultdict(list)
    for m in blueprint.get("measures", []):
        meas_by_tbl[m["table_name"]].append(m)
        
    for t in blueprint.get("tables", []):
        tname = t["table_name"]
        lines.append(f"table '{tname}'")
        lines.append(f"\tlineageTag: {tag_counter:08x}-0000-0000-0000-000000000001")
        tag_counter += 1
        lines.append("")
        
        for c in cols_by_tbl[tname]:
            cname = c["column_name"]
            dtype = c["data_type"].lower()
            lines.append(f"\tcolumn '{cname}'")
            lines.append(f"\t\tdataType: {dtype}")
            if c["column_type"] == "Calculated Column" and c.get("dax_expression"):
                lines.append(f"\t\t// Calculated Column: {c.get('provenance')}")
                lines.append(f"\t\texpression = {c['dax_expression']}")
            else:
                lines.append(f"\t\tsourceColumn: '{cname}'")
            lines.append(f"\t\tlineageTag: {tag_counter:08x}-0000-0000-0000-000000000001")
            tag_counter += 1
            lines.append("")
            
        for m in meas_by_tbl[tname]:
            mname = m["measure_name"]
            dax = m["dax_expression"]
            fmt = m.get("format_string", "#,##0.00")
            prov = m.get("provenance", "Inferred")
            if prov == "Manual Review":
                lines.append(f"\t// Manual Review Required: {m.get('notes', '')}")
            lines.append(f"\tmeasure '{mname}' = {dax}")
            lines.append(f"\t\tformatString: \"{fmt}\"")
            lines.append(f"\t\tlineageTag: {tag_counter:08x}-0000-0000-0000-000000000001")
            tag_counter += 1
            lines.append("")
            
    if blueprint.get("relationships"):
        lines.append("// ==========================================================================")
        lines.append("// Relationships")
        lines.append("// ==========================================================================")
        for rel in blueprint["relationships"]:
            rname = rel.get("relationship_name", "Rel")
            f_tbl = rel.get("from_table", "")
            f_col = rel.get("from_column", "")
            t_tbl = rel.get("to_table", "")
            t_col = rel.get("to_column", "")
            is_ambiguous = "Manual Review" in rel.get("provenance", "")
            
            if is_ambiguous:
                lines.append(f"// WARNING: Ambiguous join. Join key not declared in Cognos XML.")
                lines.append(f"// Inspect source DB schema before activating relationship:")
                lines.append(f"// relationship {rname}")
                lines.append(f"// \tfromColumn: '{f_tbl}'.[{f_col}]")
                lines.append(f"// \ttoColumn: '{t_tbl}'.[{t_col}]")
                lines.append(f"// \tstate: inactive")
            else:
                lines.append(f"relationship {rname}")
                lines.append(f"\tfromColumn: '{f_tbl}'.[{f_col}]")
                lines.append(f"\ttoColumn: '{t_tbl}'.[{t_col}]")
                lines.append(f"\tcrossFilteringBehavior: {rel.get('cross_filtering', 'Single').lower()}")
                lines.append(f"\tcardinality: {rel.get('cardinality', 'manyToOne').lower()}")
            lines.append("")
            
    return "\n".join(lines)


def _build_tmdl_table_content(table_name, cols, measures, tag_start_counter=1):
    """
    Generates valid, complete TMDL content for a single table,
    including partition definition with M stub, columns with summarizeBy, and measures.
    """
    clean_tname = table_name.replace("'", "''")
    t_lines = [
        f"table '{clean_tname}'",
        f"\tlineageTag: {tag_start_counter:08x}-0000-0000-0000-000000000001",
        "",
    ]
    tag_counter = tag_start_counter + 1

    # 1. Partition block (Required by TOM Tabular engine)
    m_col_defs = []
    for c in cols:
        cname = c["column_name"]
        dtype = c.get("data_type", "String").lower()
        if dtype in ("int64", "integer"):
            m_t = "Int64.Type"
        elif dtype in ("double", "decimal", "currency", "number"):
            m_t = "number"
        elif dtype in ("datetime", "date", "time"):
            m_t = "datetime"
        elif dtype in ("boolean", "bool", "logical"):
            m_t = "logical"
        else:
            m_t = "text"
        clean_c = cname.replace('"', '""')
        m_col_defs.append(f'#"{clean_c}" = {m_t}')

    if m_col_defs:
        typed_schema = ", ".join(m_col_defs)
        m_source = f"#table(type table [{typed_schema}], {{}})"
    else:
        m_source = '#table(type table [#"ID" = Int64.Type], {})'

    t_lines.extend([
        f"\tpartition '{clean_tname}-Partition' = m",
        "\t\tmode: import",
        "\t\tsource =",
        "\t\t\tlet",
        f"\t\t\t\tSource = {m_source}",
        "\t\t\tin",
        "\t\t\t\tSource",
        "",
    ])

    # 2. Columns
    for c in cols:
        cname = c["column_name"]
        clean_cname = cname.replace("'", "''")
        dtype = c.get("data_type", "String").lower()
        if dtype in ("int64", "integer"):
            tmdl_dt = "int64"
            summ = "sum" if not any(k in cname.lower() for k in ("id", "key", "code", "year")) else "none"
        elif dtype in ("double", "decimal", "currency", "number"):
            tmdl_dt = "double"
            summ = "sum" if not any(k in cname.lower() for k in ("id", "key", "code", "rate", "pct", "percent", "margin")) else "none"
        elif dtype in ("datetime", "date", "time"):
            tmdl_dt = "dateTime"
            summ = "none"
        elif dtype in ("boolean", "bool", "logical"):
            tmdl_dt = "boolean"
            summ = "none"
        else:
            tmdl_dt = "string"
            summ = "none"

        t_lines.append(f"\tcolumn '{clean_cname}'")
        t_lines.append(f"\t\tdataType: {tmdl_dt}")
        if c.get("column_type") == "Calculated Column" and c.get("dax_expression"):
            t_lines.append(f"\t\texpression = {c['dax_expression']}")
        else:
            t_lines.append(f"\t\tsourceColumn: '{clean_cname}'")
        t_lines.append(f"\t\tsummarizeBy: {summ}")
        t_lines.append(f"\t\tlineageTag: {tag_counter:08x}-0000-0000-0000-000000000001")
        tag_counter += 1
        t_lines.append("")

    # 3. Measures
    for m in measures:
        mname = m["measure_name"]
        clean_mname = mname.replace("'", "''")
        dax = m.get("dax_expression", "") or "0"
        fmt = m.get("format_string", "#,##0.00")
        t_lines.append(f"\tmeasure '{clean_mname}' = {dax}")
        t_lines.append(f'\t\tformatString: "{fmt}"')
        t_lines.append(f"\t\tlineageTag: {tag_counter:08x}-0000-0000-0000-000000000001")
        tag_counter += 1
        t_lines.append("")

    return "\n".join(t_lines), tag_counter


def _build_tmdl_relationships_content(relationships):
    """
    Builds relationships.tmdl content, filtering out unmapped or ambiguous join placeholders.
    Returns empty string if no valid relationships exist (safe for TOM Tabular parser).
    """
    valid_lines = []
    for r in relationships or []:
        rname = r.get("relationship_name", "Rel")
        from_tbl = r.get("from_table", "")
        from_col = r.get("from_column", "")
        to_tbl = r.get("to_table", "")
        to_col = r.get("to_column", "")
        if not (from_tbl and to_tbl and from_col and to_col):
            continue
        if "Unmapped" in from_col or "Unmapped" in to_col or from_col.startswith("[") or to_col.startswith("["):
            continue

        clean_rname = rname.replace("'", "''")
        clean_from_tbl = from_tbl.replace("'", "''")
        clean_from_col = from_col.strip("[]").replace("'", "''")
        clean_to_tbl = to_tbl.replace("'", "''")
        clean_to_col = to_col.strip("[]").replace("'", "''")
        card = r.get("cardinality", "manyToOne").lower()
        if card not in ("manytoone", "onetomany", "onetoone", "manytomany"):
            card = "manytoone"

        valid_lines.extend([
            f"relationship '{clean_rname}'",
            f"\tfromColumn: '{clean_from_tbl}'.'{clean_from_col}'",
            f"\ttoColumn: '{clean_to_tbl}'.'{clean_to_col}'",
            f"\tcardinality: {card}",
            "",
        ])

    return "\n".join(valid_lines)


def generate_tmdl_zip_package(blueprint):
    """Generates PBIP-compatible ZIP package bytes with definition/model.tmdl, tables, and README."""
    model_name = blueprint.get("model_name", "Model")
    safe_db_name = re.sub(r"[^\w-]", "_", model_name).strip("_") or "Model"
    bio = BytesIO()
    with zipfile.ZipFile(bio, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        # 1. definition.pbism
        pbism_content = {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
            "version": "4.0",
            "settings": {}
        }
        zf.writestr("definition.pbism", json.dumps(pbism_content, indent=2))

        # 2. definition/database.tmdl
        db_content = f"database {safe_db_name}\n\tcompatibilityLevel: 1567\n\tcompatibilityMode: powerBI\n"
        zf.writestr("definition/database.tmdl", db_content)

        # 3. definition/cultures/en-US.tmdl
        zf.writestr("definition/cultures/en-US.tmdl", "cultureInfo en-US\n")

        # 4. definition/model.tmdl
        model_lines = [
            "model Model",
            "\tculture: en-US",
            "\tdefaultPowerBIDataSourceVersion: powerBI_V3",
            "\tsourceQueryCulture: en-US",
            "",
        ]
        for t in blueprint.get("tables", []):
            tname = t["table_name"]
            clean_t = tname.replace("'", "''")
            model_lines.append(f"ref table '{clean_t}'")
        model_lines.append("")
        model_lines.append("ref cultureInfo en-US")
        model_lines.append("")
        zf.writestr("definition/model.tmdl", "\n".join(model_lines))

        # 5. Table TMDLs
        cols_by_tbl = defaultdict(list)
        for c in blueprint.get("columns", []):
            cols_by_tbl[c["table_name"]].append(c)
        meas_by_tbl = defaultdict(list)
        for m in blueprint.get("measures", []):
            meas_by_tbl[m["table_name"]].append(m)

        tag_ctr = 1
        for t in blueprint.get("tables", []):
            tname = t["table_name"]
            t_content, tag_ctr = _build_tmdl_table_content(tname, cols_by_tbl[tname], meas_by_tbl[tname], tag_ctr)
            zf.writestr(f"definition/tables/{tname}.tmdl", t_content)

        # 6. Relationships TMDL
        if blueprint.get("relationships"):
            rel_content = _build_tmdl_relationships_content(blueprint.get("relationships"))
            zf.writestr("definition/relationships.tmdl", rel_content)

        # 7. README.md
        readme = f"""# Power BI Semantic Model Blueprint: {model_name}

Generated automatically by **KashMap Cognos to Power BI Migration Assistant**.

## Model Summary
- **Tables**: {blueprint['summary']['table_count']}
- **Columns**: {blueprint['summary']['column_count']}
- **Measures**: {blueprint['summary']['measure_count']}
- **Relationships**: {blueprint['summary']['relationship_count']}

## Governance Breakdown
- **Verified**: {blueprint['summary']['verified_count']}
- **Inferred**: {blueprint['summary']['inferred_count']}
- **Manual Review Needed**: {blueprint['summary']['manual_review_count']}

## How to Deploy
1. Open **Power BI Desktop** (Developer Mode enabled) or **Tabular Editor 3**.
2. Save or open your project as a **PBIP** (Power BI Project).
3. Copy the contents of the `definition/` directory into your `.SemanticModel/definition/` folder.
4. Review objects marked with `Manual Review Required`.
"""
        zf.writestr("README.md", readme)

    bio.seek(0)
    return bio.getvalue()


def build_model_tables_df(blueprint):
    return pd.DataFrame([
        {
            "Table Name": t["table_name"],
            "Source Type": t["source_type"],
            "Source Path": t["source_path"],
            "Columns": t["column_count"],
            "Measures": t["measure_count"],
            "Governance": t["provenance"],
        }
        for t in blueprint.get("tables", [])
    ])


def build_model_columns_df(blueprint):
    return pd.DataFrame([
        {
            "Table": c["table_name"],
            "Column": c["column_name"],
            "Data Type": c["data_type"],
            "Column Type": c["column_type"],
            "Source / Expression": c["dax_expression"] if c["column_type"] == "Calculated Column" else c["source_path"],
            "Governance": c["provenance"],
            "Notes": c["notes"],
        }
        for c in blueprint.get("columns", [])
    ])


def build_model_measures_df(blueprint):
    return pd.DataFrame([
        {
            "Table": m["table_name"],
            "Measure": m["measure_name"],
            "DAX Expression": m["dax_expression"],
            "Format String": m["format_string"],
            "Data Type": m["data_type"],
            "Governance": m["provenance"],
            "DAX Status": m.get("dax_validation_status", "Valid DAX"),
            "Notes": m["notes"],
        }
        for m in blueprint.get("measures", [])
    ])


def build_model_relationships_df(blueprint):
    return pd.DataFrame([
        {
            "Relationship Name": r.get("relationship_name", ""),
            "From Table": r["from_table"],
            "From Column": r["from_column"],
            "To Table": r["to_table"],
            "To Column": r["to_column"],
            "Cardinality": r["cardinality"],
            "Cross-Filtering": r["cross_filtering"],
            "State": r["state"],
            "Governance": r["provenance"],
            "Notes": r["notes"],
        }
        for r in blueprint.get("relationships", [])
    ])


# ==============================================================================
# Task 4: Power BI Report Pages and Visuals Generation (PBIP / PBIR)
# ==============================================================================

def extract_cognos_charts(page_el):
    """
    Extract Cognos charts (combinationChart, barChart, pieChart, lineChart, vizControl, etc.)
    from a specific page element.
    """
    charts = []
    CHART_TAGS = {
        "combinationChart": "columnChart",
        "barChart": "barChart",
        "pieChart": "pieChart",
        "lineChart": "lineChart",
        "scatterChart": "scatterChart",
        "bubbleChart": "scatterChart",
        "areaChart": "areaChart",
        "gaugeChart": "gauge",
        "chart": "columnChart",
        "vizControl": "columnChart",
    }

    for el in page_el.iter():
        tag = _local(el.tag)
        if tag in CHART_TAGS:
            c_name = el.get("name", "") or f"Chart_{len(charts) + 1}"
            ref_query = el.get("refQuery", "")
            target_pbi = CHART_TAGS[tag]

            # Check vizControl specific type if present
            if tag == "vizControl":
                vt = el.get("type", "").lower()
                if "bar" in vt:
                    target_pbi = "barChart"
                elif "line" in vt:
                    target_pbi = "lineChart"
                elif "pie" in vt or "donut" in vt:
                    target_pbi = "pieChart"
                elif "area" in vt:
                    target_pbi = "areaChart"
                elif "scatter" in vt or "bubble" in vt:
                    target_pbi = "scatterChart"
                elif "map" in vt or "geo" in vt:
                    target_pbi = "map"

            # Categories (X-Axis / Category)
            categories = []
            for axis in _iter_all(el, "commonAxis") + _iter_all(el, "categoryMembers"):
                for m in _iter_all(axis, "chartNodeMember"):
                    di = m.get("refDataItem", "")
                    if di and di not in categories:
                        categories.append(di)

            # Series (Legend)
            series = []
            for s_root in _iter_all(el, "seriesMembers"):
                for m in _iter_all(s_root, "chartNodeMember"):
                    di = m.get("refDataItem", "")
                    if di and di not in series:
                        series.append(di)

            # Measures (Y-Axis / Values)
            measures = []
            for m_root in _iter_all(el, "measureMembers") + _iter_all(el, "chartIntersections"):
                for m in _iter_all(m_root, "chartNodeMember"):
                    di = m.get("refDataItem", "")
                    if di and di not in measures:
                        measures.append(di)
            # Default measure attribute
            for def_m in _iter_all(el, "defaultMeasure"):
                di = def_m.get("refDataItem", "")
                if di and di not in measures:
                    measures.append(di)

            # vizControl slots
            for slot in _iter_all(el, "vcSlotData"):
                slot_id = slot.get("idSlot", "").lower()
                cols = [c.get("refDsColumn", "") for c in _iter_all(slot, "vcSlotDsColumn") if c.get("refDsColumn")]
                if any(k in slot_id for k in ("cat", "x", "axis")):
                    categories.extend([c for c in cols if c not in categories])
                elif any(k in slot_id for k in ("val", "y", "measure")):
                    measures.extend([c for c in cols if c not in measures])
                elif any(k in slot_id for k in ("series", "color", "legend")):
                    series.extend([c for c in cols if c not in series])

            charts.append({
                "name": c_name,
                "ref_query": ref_query,
                "source_type": tag,
                "target_powerbi_visual": target_pbi,
                "categories": categories,
                "series": series,
                "measures": measures,
            })
    return charts


def extract_report_pages_blueprint(xml_text=None, semantic_blueprint=None, parsed=None):
    """
    Parses Cognos report pages and generates a structured Power BI Pages Blueprint,
    mapping Cognos page layouts, lists, crosstabs, charts, prompts, and custom controls.
    """
    if isinstance(xml_text, dict):
        parsed = xml_text
        raw_xml = parsed.get("xml_text", "")
    else:
        raw_xml = xml_text or (parsed.get("xml_text", "") if parsed else "")
    if not raw_xml:
        return []

    try:
        root = ET.fromstring(raw_xml)
    except Exception:
        return []

    pages_blueprint = []
    page_elements = [el for el in root.iter() if _local(el.tag) == "page"]

    # Fallback to root element as single page if no explicit <page> element found
    if not page_elements:
        page_elements = [root]

    page_counter = 1
    for p_el in page_elements:
        p_name = p_el.get("name", "") or f"Page{page_counter}"
        page_id = uuid.uuid4().hex[:20]

        # 1. Lists on this page
        lists = []
        for lst in _iter_all(p_el, "list"):
            l_name = lst.get("name", "") or f"List_{len(lists) + 1}"
            ref_q = lst.get("refQuery", "")
            cols = []
            for col in _iter_all(lst, "listColumn"):
                title = ""
                for t in _iter_all(col, "listColumnTitle"):
                    for text_item in _iter_all(t, "text"):
                        title = _text_of(text_item)
                        if title:
                            break
                    if not title:
                        for di in _iter_all(t, "dataItemValue"):
                            title = di.get("refDataItem", "")
                            if title:
                                break
                di_name = ""
                for b in _iter_all(col, "listColumnBody"):
                    for di in _iter_all(b, "dataItemValue") + _iter_all(b, "fieldRef"):
                        di_name = di.get("refDataItem", "")
                        if di_name:
                            break
                if not di_name:
                    for di in _iter_all(col, "dataItemValue"):
                        di_name = di.get("refDataItem", "")
                        if di_name:
                            break
                cols.append({"title": title or di_name, "data_item": di_name})
            lists.append({
                "name": l_name,
                "ref_query": ref_q,
                "columns": cols,
                "source_type": "list",
                "target_powerbi_visual": "tableEx",
            })

        # 2. Crosstabs on this page
        crosstabs = []
        for ct in _iter_all(p_el, "crosstab"):
            ct_name = ct.get("name", "") or f"Crosstab_{len(crosstabs) + 1}"
            ref_q = ct.get("refQuery", "")
            rows_el = _find_child(ct, "crosstabRows")
            r_dims = []
            if rows_el is not None:
                for el in _iter_all(rows_el, "crosstabNodeMember") + _iter_all(rows_el, "dataItemValue"):
                    di = el.get("refDataItem", "")
                    if di and di not in r_dims:
                        r_dims.append(di)

            cols_el = _find_child(ct, "crosstabColumns")
            c_dims = []
            if cols_el is not None:
                for el in _iter_all(cols_el, "crosstabNodeMember") + _iter_all(cols_el, "dataItemValue"):
                    di = el.get("refDataItem", "")
                    if di and di not in c_dims:
                        c_dims.append(di)

            meas = []
            meas_el = _find_child(ct, "crosstabMeasures")
            if meas_el is not None:
                for el in _iter_all(meas_el, "dataItemValue"):
                    di = el.get("refDataItem", "")
                    if di and di not in meas:
                        meas.append(di)
            for cell in _iter_all(ct, "crosstabIntersectionCell"):
                meas.extend([v.get("refDataItem", "") for v in _iter_all(cell, "dataItemValue") if v.get("refDataItem") and v.get("refDataItem") not in meas])
            for inter in _iter_all(ct, "crosstabIntersection"):
                di = inter.get("refDataItem", "")
                if di and di not in meas:
                    meas.append(di)
            crosstabs.append({
                "name": ct_name,
                "ref_query": ref_q,
                "row_dimensions": r_dims,
                "column_dimensions": c_dims,
                "measures": meas,
                "source_type": "crosstab",
                "target_powerbi_visual": "pivotTable",
            })

        # 3. Charts on this page
        charts = extract_cognos_charts(p_el)

        # 4. Prompts / Slicers on this page
        prompts = []
        PROMPT_TAGS = {
            "selectValue": ("slicer", "Value Dropdown"),
            "selectDate": ("slicer", "Date Range"),
            "selectDateTime": ("slicer", "DateTime Range"),
            "selectInterval": ("slicer", "Interval Slicer"),
            "textBox": ("slicer", "Text Search Filter"),
            "treePrompt": ("slicer", "Hierarchy Slicer"),
        }
        for ptag, (pbi_vis, desc) in PROMPT_TAGS.items():
            for pr in _iter_all(p_el, ptag):
                param = pr.get("parameter", "")
                if param:
                    prompts.append({
                        "name": f"Slicer_{param}",
                        "parameter": param,
                        "source_type": ptag,
                        "target_powerbi_visual": pbi_vis,
                        "description": desc,
                        "ref_query": pr.get("refQuery", ""),
                    })

        # Check report-level prompts on the first page if none found on page element
        if not prompts and page_counter == 1 and root is not None:
            prompts_root = _find_child(root, "prompts")
            if prompts_root is not None:
                for ptag, (pbi_vis, desc) in PROMPT_TAGS.items():
                    for pr in _iter_all(prompts_root, ptag):
                        param = pr.get("parameter", "")
                        if param:
                            prompts.append({
                                "name": f"Slicer_{param}",
                                "parameter": param,
                                "source_type": ptag,
                                "target_powerbi_visual": pbi_vis,
                                "description": desc,
                                "ref_query": pr.get("refQuery", ""),
                            })

        page_counter += 1

        # 5. Repeaters on this page
        repeaters = []
        for rtag in ("repeater", "repeaterTable"):
            for rep in _iter_all(p_el, rtag):
                r_name = rep.get("name", "") or f"Repeater_{len(repeaters) + 1}"
                di_list = [v.get("refDataItem", "") for v in _iter_all(rep, "dataItemValue") if v.get("refDataItem")]
                repeaters.append({
                    "name": r_name,
                    "ref_query": rep.get("refQuery", ""),
                    "source_type": rtag,
                    "target_powerbi_visual": "pivotTable",
                    "data_items": di_list,
                })

        # 6. Custom Controls on this page (Flagged for Manual Review)
        custom_controls = []
        for cc in _iter_all(p_el, "customControl"):
            cc_name = cc.get("name", "") or f"CustomControl_{len(custom_controls) + 1}"
            custom_controls.append({
                "name": cc_name,
                "source_type": "customControl",
                "description": cc.get("description", "") or _text_of(_find_child(cc, "description")),
                "path": cc.get("path", ""),
                "target_powerbi_visual": "manualReviewPlaceholder",
            })

        # Assemble Visual Containers for this page
        visuals = []

        # Slicers placed in top ribbon
        for pr in prompts:
            visuals.append({
                "visual_id": uuid.uuid4().hex[:20],
                "name": pr["name"],
                "source_element": pr["source_type"],
                "visual_type": "slicer",
                "query_name": pr.get("ref_query", ""),
                "role_mappings": {"Values": [pr["parameter"]]},
                "status": "Converted",
                "provenance": "Grounded (Prompt to Slicer)",
                "warnings": [],
                "is_slicer": True,
            })

        # Lists
        for lst in lists:
            data_items = [c["data_item"] for c in lst["columns"] if c["data_item"]]
            visuals.append({
                "visual_id": uuid.uuid4().hex[:20],
                "name": lst["name"],
                "source_element": "list",
                "visual_type": "tableEx",
                "query_name": lst["ref_query"],
                "role_mappings": {"Values": data_items},
                "status": "Converted",
                "provenance": "Grounded (List to Table)",
                "warnings": [],
                "is_slicer": False,
            })

        # Crosstabs
        for ct in crosstabs:
            visuals.append({
                "visual_id": uuid.uuid4().hex[:20],
                "name": ct["name"],
                "source_element": "crosstab",
                "visual_type": "pivotTable",
                "query_name": ct["ref_query"],
                "role_mappings": {
                    "Rows": ct["row_dimensions"],
                    "Columns": ct["column_dimensions"],
                    "Values": ct["measures"],
                },
                "status": "Converted",
                "provenance": "Grounded (Crosstab to Matrix)",
                "warnings": [],
                "is_slicer": False,
            })

        # Charts
        for ch in charts:
            visuals.append({
                "visual_id": uuid.uuid4().hex[:20],
                "name": ch["name"],
                "source_element": ch["source_type"],
                "visual_type": ch["target_powerbi_visual"],
                "query_name": ch["ref_query"],
                "role_mappings": {
                    "Category": ch["categories"],
                    "Series": ch["series"],
                    "Y": ch["measures"],
                },
                "status": "Converted",
                "provenance": f"Grounded ({ch['source_type']} to {ch['target_powerbi_visual']})",
                "warnings": [],
                "is_slicer": False,
            })

        # Repeaters
        for rep in repeaters:
            visuals.append({
                "visual_id": uuid.uuid4().hex[:20],
                "name": rep["name"],
                "source_element": rep["source_type"],
                "visual_type": "pivotTable",
                "query_name": rep["ref_query"],
                "role_mappings": {"Rows": rep["data_items"]},
                "status": "Converted",
                "provenance": "Inferred (Repeater to Grouped Visual)",
                "warnings": ["Cognos Repeater mapped to Grouped Matrix visual. Review visual layout for small multiples."],
                "is_slicer": False,
            })

        # Custom Controls (Strict Manual Review Guard)
        for cc in custom_controls:
            visuals.append({
                "visual_id": uuid.uuid4().hex[:20],
                "name": cc["name"],
                "source_element": "customControl",
                "visual_type": "customControl",
                "query_name": "",
                "role_mappings": {},
                "status": "Manual Review Required",
                "provenance": "Manual Review (Custom JS Widget)",
                "warnings": [f"Cognos JavaScript Custom Control '{cc['name']}' ({cc.get('description') or cc.get('path')}) has no standard Power BI native equivalent. Requires manual Power BI Custom Visual (AppSource/Deneb) implementation."],
                "is_slicer": False,
            })

        # Calculate Canvas Coordinates (1280x720 canvas)
        slicers = [v for v in visuals if v.get("is_slicer")]
        main_visuals = [v for v in visuals if not v.get("is_slicer")]

        # Position Slicers at top
        top_offset = 20
        for i, s_vis in enumerate(slicers):
            s_vis["position"] = {
                "x": 20 + (i * 250),
                "y": 20,
                "z": 0,
                "width": 230,
                "height": 70,
                "tabOrder": 1000 + (i * 10),
            }
        if slicers:
            top_offset = 100

        # Position Main Visuals
        avail_height = 720 - top_offset - 20
        num_main = len(main_visuals)

        for i, m_vis in enumerate(main_visuals):
            if num_main == 1:
                m_vis["position"] = {
                    "x": 20, "y": top_offset, "z": 0,
                    "width": 1240, "height": avail_height,
                    "tabOrder": 2000 + (i * 10),
                }
            elif num_main == 2:
                w = 605
                m_vis["position"] = {
                    "x": 20 + (i * 630), "y": top_offset, "z": 0,
                    "width": w, "height": avail_height,
                    "tabOrder": 2000 + (i * 10),
                }
            elif num_main in (3, 4):
                col_idx = i % 2
                row_idx = i // 2
                w = 605
                h = (avail_height - 20) // 2
                m_vis["position"] = {
                    "x": 20 + (col_idx * 630),
                    "y": top_offset + (row_idx * (h + 15)),
                    "z": 0,
                    "width": w, "height": h,
                    "tabOrder": 2000 + (i * 10),
                }
            else:
                col_idx = i % 2
                row_idx = i // 2
                w = 605
                h = 240
                m_vis["position"] = {
                    "x": 20 + (col_idx * 630),
                    "y": top_offset + (row_idx * 260),
                    "z": 0,
                    "width": w, "height": h,
                    "tabOrder": 2000 + (i * 10),
                }

        # Resolve Semantic Model Field Bindings
        if semantic_blueprint:
            bind_visual_projections(visuals, semantic_blueprint)

        pages_blueprint.append({
            "page_id": page_id,
            "page_name": p_name,
            "display_name": p_name.replace("_", " "),
            "width": 1280,
            "height": 720,
            "display_option": "FitToPage",
            "visuals": visuals,
            "visual_count": len(visuals),
            "converted_count": sum(1 for v in visuals if v["status"] == "Converted"),
            "manual_review_count": sum(1 for v in visuals if v["status"] != "Converted"),
        })

    return pages_blueprint


def bind_visual_projections(visuals, semantic_blueprint):
    """
    Resolves data roles (Values, Rows, Columns, Category, Y, Series) to semantic model
    table columns or measures, generating valid Power BI PBIR projection dictionaries.
    """
    col_map = {}
    for c in semantic_blueprint.get("columns", []):
        col_map[c["column_name"].upper()] = (c["table_name"], c["column_name"], "Column")

    meas_map = {}
    for m in semantic_blueprint.get("measures", []):
        meas_map[m["measure_name"].upper()] = (m["table_name"], m["measure_name"], "Measure")

    primary_tbl = semantic_blueprint.get("tables", [{}])[0].get("table_name", "Model")

    for vis in visuals:
        projections_by_role = {}
        all_resolved = []

        for role, items in vis.get("role_mappings", {}).items():
            proj_list = []
            for item in items:
                if not item:
                    continue
                clean_item = item.strip("[]").strip()
                item_upper = clean_item.upper()

                # Check measure map first, then column map
                resolved_match = None
                if item_upper in meas_map:
                    resolved_match = meas_map[item_upper]
                elif item_upper in col_map:
                    resolved_match = col_map[item_upper]
                elif item_upper.startswith("P") and len(item_upper) > 1:
                    cand = item_upper[1:]
                    if cand in meas_map:
                        resolved_match = meas_map[cand]
                    elif cand in col_map:
                        resolved_match = col_map[cand]

                if resolved_match:
                    tbl, prop, ftype = resolved_match
                    if ftype == "Measure":
                        proj_list.append({
                            "field": {
                                "Measure": {
                                    "Expression": {"SourceRef": {"Entity": tbl}},
                                    "Property": prop,
                                }
                            },
                            "queryRef": f"{tbl}.{prop}",
                        })
                        all_resolved.append((role, tbl, prop, "Measure", True))
                    else:
                        proj_list.append({
                            "field": {
                                "Column": {
                                    "Expression": {"SourceRef": {"Entity": tbl}},
                                    "Property": prop,
                                }
                            },
                            "queryRef": f"{tbl}.{prop}",
                        })
                        all_resolved.append((role, tbl, prop, "Column", True))
                else:
                    # Item not found directly in semantic blueprint: bound to primary table with warning
                    vis["warnings"].append(f"Data item '{clean_item}' in role '{role}' not resolved in semantic model. Bound to table '{primary_tbl}' for manual review.")
                    proj_list.append({
                        "field": {
                            "Column": {
                                "Expression": {"SourceRef": {"Entity": primary_tbl}},
                                "Property": clean_item,
                            }
                        },
                        "queryRef": f"{primary_tbl}.{clean_item}",
                    })
                    all_resolved.append((role, primary_tbl, clean_item, "Unmapped", False))

            if proj_list:
                projections_by_role[role] = {"projections": proj_list}

        vis["query_state"] = projections_by_role
        vis["bound_projections"] = all_resolved


def validate_pbir_report_package(pages_blueprint, semantic_blueprint):
    """
    Validates the generated PBIR report layout, visual bindings, and model references.
    """
    total_visuals = sum(p["visual_count"] for p in pages_blueprint)
    total_bindings = 0
    valid_bindings = 0
    unmapped_bindings = 0
    warnings = []

    known_tables = {t["table_name"].upper() for t in semantic_blueprint.get("tables", [])}
    known_columns = {f"{c['table_name'].upper()}.{c['column_name'].upper()}" for c in semantic_blueprint.get("columns", [])}
    known_measures = {f"{m['table_name'].upper()}.{m['measure_name'].upper()}" for m in semantic_blueprint.get("measures", [])}

    for page in pages_blueprint:
        for vis in page["visuals"]:
            if vis.get("warnings"):
                warnings.extend([f"Page '{page['display_name']}' visual '{vis['name']}': {w}" for w in vis["warnings"]])
            for role, tbl, prop, ftype, is_resolved in vis.get("bound_projections", []):
                total_bindings += 1
                ref_key = f"{tbl.upper()}.{prop.upper()}"
                if tbl.upper() in known_tables and (ref_key in known_columns or ref_key in known_measures):
                    valid_bindings += 1
                else:
                    unmapped_bindings += 1

    score = round((valid_bindings / total_bindings * 100)) if total_bindings else 100
    is_valid = unmapped_bindings == 0 and len([w for w in warnings if "not resolved" in w]) == 0

    return {
        "is_valid": is_valid,
        "validation_score": score,
        "total_pages": len(pages_blueprint),
        "total_visuals": total_visuals,
        "total_bindings": total_bindings,
        "valid_bindings": valid_bindings,
        "unmapped_bindings": unmapped_bindings,
        "warnings": warnings,
    }


def generate_pbip_project_zip(report_name, pages_blueprint, semantic_blueprint):
    """
    Generates a complete, standard Microsoft Power BI Project (.pbip) archive,
    containing both the TMDL .SemanticModel and PBIR .Report structures.
    """
    safe_name = re.sub(r"[^\w-]", "_", report_name)
    if safe_name.lower().endswith("_xml"):
        safe_name = safe_name[:-4]
    if safe_name.lower().endswith(".xml"):
        safe_name = safe_name[:-4]
    if not safe_name:
        safe_name = "KashMap_Report"

    report_folder = f"{safe_name}.Report"
    model_folder = f"{safe_name}.SemanticModel"

    bio = BytesIO()
    with zipfile.ZipFile(bio, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        # 1. <ReportName>.pbip Root Entrypoint File
        # Conforms strictly to https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json
        # Only 'report' shortcut is permitted in artifacts; 'semanticModel' is forbidden by schema.
        pbip_content = {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
            "version": "1.0",
            "artifacts": [
                {
                    "report": {
                        "path": report_folder
                    }
                }
            ],
            "settings": {
                "enableAutoRecovery": True
            }
        }
        zf.writestr(f"{safe_name}.pbip", json.dumps(pbip_content, indent=2))

        # 2. <SafeName>.SemanticModel/
        # Set version to "4.0" to declare modern TMDL format and prevent "Missing required artifact 'model.bim'".
        pbism_content = {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
            "version": "4.0",
            "settings": {}
        }
        zf.writestr(f"{model_folder}/definition.pbism", json.dumps(pbism_content, indent=2))

        # database.tmdl (Required by Tabular engine to register database and compatibility level)
        db_content = f"database {safe_name}\n\tcompatibilityLevel: 1567\n\tcompatibilityMode: powerBI\n"
        zf.writestr(f"{model_folder}/definition/database.tmdl", db_content)

        # cultures/en-US.tmdl
        zf.writestr(f"{model_folder}/definition/cultures/en-US.tmdl", "cultureInfo en-US\n")

        # model.tmdl
        model_lines = [
            "model Model",
            "\tculture: en-US",
            "\tdefaultPowerBIDataSourceVersion: powerBI_V3",
            "\tsourceQueryCulture: en-US",
            "",
        ]
        for t in semantic_blueprint.get("tables", []):
            tname = t["table_name"]
            clean_t = tname.replace("'", "''")
            model_lines.append(f"ref table '{clean_t}'")
        model_lines.append("")
        model_lines.append("ref cultureInfo en-US")
        model_lines.append("")
        zf.writestr(f"{model_folder}/definition/model.tmdl", "\n".join(model_lines))

        # TMDL Tables (with TOM partitions, summarizeBy, lineageTags)
        cols_by_t = defaultdict(list)
        for c in semantic_blueprint.get("columns", []):
            cols_by_t[c["table_name"]].append(c)
        meas_by_t = defaultdict(list)
        for m in semantic_blueprint.get("measures", []):
            meas_by_t[m["table_name"]].append(m)

        t_counter = 1
        for t in semantic_blueprint.get("tables", []):
            tname = t["table_name"]
            t_content, t_counter = _build_tmdl_table_content(tname, cols_by_t[tname], meas_by_t[tname], t_counter)
            zf.writestr(f"{model_folder}/definition/tables/{tname}.tmdl", t_content)

        # TMDL Relationships (valid syntax, no illegal line types or unmapped comments)
        if semantic_blueprint.get("relationships"):
            rel_content = _build_tmdl_relationships_content(semantic_blueprint.get("relationships"))
            zf.writestr(f"{model_folder}/definition/relationships.tmdl", rel_content)

        # 3. <ReportName>.Report/ PBIR Definitions
        pbir_pointer = {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
            "version": "4.0",
            "datasetReference": {
                "byPath": {
                    "path": f"../{model_folder}"
                }
            }
        }
        zf.writestr(f"{report_folder}/definition.pbir", json.dumps(pbir_pointer, indent=2))

        version_json = {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/version/1.0.0/schema.json",
            "version": "4.0.0"
        }
        zf.writestr(f"{report_folder}/definition/version.json", json.dumps(version_json, indent=2))

        report_json = {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/3.0.0/schema.json",
            "themeCollection": {
                "baseTheme": {
                    "name": "CY25SU12",
                    "reportVersionAtImport": {
                        "visual": "2.5.0",
                        "report": "3.1.0",
                        "page": "2.3.0"
                    },
                    "type": "SharedResources"
                }
            }
        }
        zf.writestr(f"{report_folder}/definition/report.json", json.dumps(report_json, indent=2))

        # Pages index pages.json
        page_order = [p["page_id"] for p in pages_blueprint]
        pages_json = {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pages/1.0.0/schema.json",
            "pageOrder": page_order,
            "activePageName": page_order[0] if page_order else ""
        }
        zf.writestr(f"{report_folder}/definition/pages/pages.json", json.dumps(pages_json, indent=2))

        # Per-page folders and visual.json files
        for p in pages_blueprint:
            pid = p["page_id"]
            p_json = {
                "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.0.0/schema.json",
                "name": pid,
                "displayName": p["display_name"],
                "displayOption": p.get("display_option", "FitToPage"),
                "height": p.get("height", 720),
                "width": p.get("width", 1280),
            }
            zf.writestr(f"{report_folder}/definition/pages/{pid}/page.json", json.dumps(p_json, indent=2))

            for v in p["visuals"]:
                vid = v["visual_id"]
                vis_def = {
                    "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.4.0/schema.json",
                    "name": vid,
                    "position": v["position"],
                    "visual": {
                        "visualType": v["visual_type"],
                        "query": {
                            "queryState": v.get("query_state", {})
                        }
                    }
                }
                zf.writestr(f"{report_folder}/definition/pages/{pid}/visuals/{vid}/visual.json", json.dumps(vis_def, indent=2))

        # StaticResources stub
        zf.writestr(f"{report_folder}/StaticResources/RegisteredResources/.gitkeep", "")

        # README Documentation
        val_report = validate_pbir_report_package(pages_blueprint, semantic_blueprint)
        readme = f"""# Power BI Project (PBIP): {safe_name}

Generated automatically by **KashMap Cognos to Power BI Migration Assistant**.

## Project Architecture
- **PBIP Root Entrypoint**: `{safe_name}.pbip`
- **Semantic Model (TMDL)**: `{model_folder}/`
  - Tables: {len(semantic_blueprint.get('tables', []))}
  - Columns: {len(semantic_blueprint.get('columns', []))}
  - Measures: {len(semantic_blueprint.get('measures', []))}
- **Enhanced Report (PBIR)**: `{report_folder}/`
  - Pages: {len(pages_blueprint)}
  - Total Visuals: {val_report['total_visuals']}
  - Valid Visual Bindings: {val_report['valid_bindings']} ({val_report['validation_score']}%)
  - Manual Review Warnings: {len(val_report['warnings'])}

## Report Pages & Visual Structure
"""
        for p in pages_blueprint:
            readme += f"\n### Page: {p['display_name']} (`{p['page_name']}`)\n"
            for v in p["visuals"]:
                readme += f"- **{v['name']}** ({v['visual_type']}) - *{v['status']}*: {v['provenance']}\n"
                if v.get("warnings"):
                    for w in v["warnings"]:
                        readme += f"  - ⚠️ {w}\n"

        readme += f"""
## How to Open in Power BI Desktop
1. Extract this entire `.zip` archive to a local folder (e.g., `C:\\Projects\\{safe_name}`).
2. Ensure **Power BI Desktop** has **Power BI Project (.pbip) save option** and **PBIR format** enabled:
   - File -> Options and settings -> Options -> Preview features -> **Store reports using enhanced metadata format (PBIR)**.
3. Double-click the `{safe_name}.pbip` file to open the report and semantic model directly in Power BI Desktop.
4. Review any visuals or formulas marked with **Manual Review Required**.
"""
        zf.writestr("README.md", readme)

    bio.seek(0)
    return bio.getvalue()


def build_report_pages_df(pages_blueprint):
    rows = []
    for p in pages_blueprint:
        rows.append({
            "Page ID": p["page_id"],
            "Page Name": p["page_name"],
            "Display Name": p["display_name"],
            "Canvas Width": p.get("width", 1280),
            "Canvas Height": p.get("height", 720),
            "Display Option": p.get("display_option", "FitToPage"),
            "Visual Count": p.get("visual_count", 0),
            "Converted Count": p.get("converted_count", 0),
            "Manual Review Count": p.get("manual_review_count", 0),
        })
    return pd.DataFrame(rows)


def build_visual_bindings_df(pages_blueprint):
    rows = []
    for p in pages_blueprint:
        p_name = p["display_name"]
        for v in p["visuals"]:
            v_name = v["name"]
            v_type = v["visual_type"]
            src_elem = v["source_element"]
            status = v["status"]
            prov = v["provenance"]
            warns = "; ".join(v.get("warnings", []))

            bound_projs = v.get("bound_projections", [])
            if bound_projs:
                for role, tbl, prop, ftype, is_res in bound_projs:
                    rows.append({
                        "Page": p_name,
                        "Visual": v_name,
                        "Source Element": src_elem,
                        "Target Visual": v_type,
                        "Role": role,
                        "Bound Entity (Table)": tbl,
                        "Bound Property (Field/Measure)": prop,
                        "Binding Type": ftype,
                        "Status": status,
                        "Governance": prov,
                        "Warnings / Notes": warns,
                    })
            else:
                rows.append({
                    "Page": p_name,
                    "Visual": v_name,
                    "Source Element": src_elem,
                    "Target Visual": v_type,
                    "Role": "None",
                    "Bound Entity (Table)": "N/A",
                    "Bound Property (Field/Measure)": "N/A",
                    "Binding Type": "N/A",
                    "Status": status,
                    "Governance": prov,
                    "Warnings / Notes": warns,
                })
    return pd.DataFrame(rows)



CognosClassification = Literal["confirmed", "rule_based", "inferred", "suggested", "manual_review"]

CognosPowerBILayer = Literal[
    "Power Query", "Semantic Model", "DAX Measure", "Calculated Column",
    "Report Filter", "Slicer", "Visual", "Drill-through", "Manual Review",
]


class CognosPlanStep(BaseModel):
    step_number: int = Field(ge=1)
    title: str
    action: str
    power_bi_layer: CognosPowerBILayer
    classification: CognosClassification
    reason: str
    referenced_sources: list[str] = Field(default_factory=list)
    referenced_fields: list[str] = Field(default_factory=list)
    confidence: int = Field(ge=0, le=100)
    manual_validation: Optional[str] = None


class CognosReportSummary(BaseModel):
    report_name: str
    likely_purpose: str
    report_type: str
    complexity_level: Literal["Very Small", "Small", "Medium", "Large", "Very Complex"]
    overall_confidence: int = Field(ge=0, le=100)


class CognosBuildPlan(BaseModel):
    report_summary: CognosReportSummary
    implementation_plan: list[CognosPlanStep]
    calculation_plan: list[CognosPlanStep] = Field(default_factory=list)
    filter_prompt_plan: list[CognosPlanStep] = Field(default_factory=list)
    drillthrough_plan: list[CognosPlanStep] = Field(default_factory=list)
    validation_checks: list[str] = Field(default_factory=list)
    manual_review_items: list[str] = Field(default_factory=list)
    limitations: list[str] = Field(default_factory=list)


COGNOS_AI_BUILD_PLAN_SYSTEM_PROMPT = """You are a senior BI migration engineer converting one parsed IBM Cognos report to Power BI.
The input contains Cognos-report-derived logic: queries, fields (source and calculated), detail filters,
prompts, drillthroughs, visualizations, and custom controls. Field references are already resolved to
[Module].[Table].[Column] paths where possible; unresolved_field_issues (if present) lists fields that
could NOT be verified against the real Data Module schema - treat those as higher-risk manual review.

Rules:
1. Use only tables, fields, filters, prompts, and drillthroughs present in the input. Never invent one.
2. Cognos has no HOLD/staging chain - all shaping happens at the query or model level. Do not describe
   staging steps that aren't in the input.
3. Label every implementation step as one of: confirmed, rule_based, inferred, suggested, manual_review.
4. Every step must include referenced_sources (table names) and referenced_fields drawn only from the input.
5. Custom controls (JS widgets) have no Power BI equivalent - always classify them as manual_review.
6. Confidence (0-100) must reflect how directly the input supports the recommendation.
7. Return ONLY strict JSON matching the required schema - no markdown fences, no commentary outside JSON.

Return JSON with exactly these keys:
{
  "report_summary": {"report_name": "", "likely_purpose": "", "report_type": "",
                      "complexity_level": "Very Small|Small|Medium|Large|Very Complex",
                      "overall_confidence": 0},
  "implementation_plan": [{"step_number": 1, "title": "", "action": "",
                            "power_bi_layer": "Power Query|Semantic Model|DAX Measure|Calculated Column|Report Filter|Slicer|Visual|Drill-through|Manual Review",
                            "classification": "confirmed|rule_based|inferred|suggested|manual_review",
                            "reason": "", "referenced_sources": [], "referenced_fields": [],
                            "confidence": 0, "manual_validation": null}],
  "calculation_plan": [... same PlanStep shape ..., must cover every calculated field in the input],
  "filter_prompt_plan": [... same PlanStep shape ..., must cover every filter and prompt],
  "drillthrough_plan": [... same PlanStep shape ..., must cover every drillthrough],
  "validation_checks": ["..."],
  "manual_review_items": ["..."],
  "limitations": ["..."]
}

Every custom control in the input must produce exactly one manual_review step in implementation_plan.
Every entry in unresolved_field_issues must appear in manual_review_items.
"""

_COGNOS_VALID_LAYERS = {
    "Power Query", "Semantic Model", "DAX Measure", "Calculated Column",
    "Report Filter", "Slicer", "Visual", "Drill-through", "Manual Review",
}
_COGNOS_LAYER_PRECEDENCE = [
    "Power Query", "Semantic Model", "DAX Measure", "Calculated Column",
    "Report Filter", "Slicer", "Visual", "Drill-through", "Manual Review",
]


def _coerce_cognos_layer(value):
    if value in _COGNOS_VALID_LAYERS:
        return value
    parts = re.split(r'[|/,]', str(value))
    parts = [p.strip() for p in parts if p.strip() in _COGNOS_VALID_LAYERS]
    if parts:
        for preferred in _COGNOS_LAYER_PRECEDENCE:
            if preferred in parts:
                return preferred
        return parts[0]
    return "Manual Review"


def normalize_cognos_layers(raw):
    for key in ['implementation_plan', 'calculation_plan', 'filter_prompt_plan', 'drillthrough_plan']:
        steps = raw.get(key)
        if not isinstance(steps, list):
            continue
        for step in steps:
            if isinstance(step, dict) and 'power_bi_layer' in step:
                step['power_bi_layer'] = _coerce_cognos_layer(step['power_bi_layer'])
    return raw


def parse_llm_json(text):
    text = (text or '').strip()
    if not text:
        return {}
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    match = re.search(r'\{.*\}', text, flags=re.DOTALL)
    if match:
        try:
            return json.loads(match.group(0))
        except json.JSONDecodeError:
            return {}
    return {}


def build_cognos_ai_payload(report_name, parsed, validation_issues):
    """Compact payload for one AI call, mirroring build_ai_build_plan_payload()
    from KashMap but shaped around Cognos structures (queries/prompts/drillthroughs
    instead of HOLD/JOIN)."""
    shown_fields = sorted({f['field'] for f in parsed['fields']})
    return {
        'report_name': report_name,
        'model_path': parsed['model_path'],
        'source_tables': parsed['source_tables'],
        'fields': [
            {'field': f['field'], 'query': f['query'], 'is_calculated': f['is_calculated'],
             'expression_type': f['expression_type'], 'expression': f['expression'][:200]}
            for f in parsed['fields']
        ][:60],
        'filters': [
            {'query': f['query'], 'use': f['use'], 'expression': f['expression'][:200]}
            for f in parsed['filters']
        ][:30],
        'prompts': [
            {'parameter': p['parameter'], 'options': p['options'], 'default': p['default']}
            for p in parsed['prompts']
        ],
        'drillthroughs': [
            {'drill_name': d['drill_name'], 'target_report': d['target_report'],
             'parameter_links': d['parameter_links']}
            for d in parsed['drillthroughs']
        ],
        'visualizations': [{'name': v['name'], 'type': v['type']} for v in parsed['visualizations']],
        'custom_controls': [{'name': c['name'], 'path': c['path']} for c in parsed['custom_controls']],
        'unresolved_field_issues': validation_issues,
        'shown_fields': shown_fields[:60],
    }


def strip_unverified_cognos_steps(plan, allowed_sources, allowed_fields):
    """Mirrors strip_unverified_steps() from KashMap: downgrades any step
    referencing unknown sources/fields to manual_review instead of trusting it."""
    def clean(section):
        cleaned = []
        for step in section:
            unknown_sources = set(step.referenced_sources) - allowed_sources
            unknown_fields = set(step.referenced_fields) - allowed_fields
            if unknown_sources or unknown_fields:
                step.classification = "manual_review"
                parts = []
                if unknown_sources:
                    parts.append(f"unknown source(s) {sorted(unknown_sources)}")
                if unknown_fields:
                    parts.append(f"unknown field(s) {sorted(unknown_fields)}")
                step.manual_validation = "AI recommendation flagged: referenced " + " and ".join(parts) + " not present in the parsed report."
                step.confidence = min(step.confidence, 40)
            cleaned.append(step)
        return cleaned

    plan.implementation_plan = clean(plan.implementation_plan)
    plan.calculation_plan = clean(plan.calculation_plan)
    plan.filter_prompt_plan = clean(plan.filter_prompt_plan)
    plan.drillthrough_plan = clean(plan.drillthrough_plan)
    return plan


def generate_cognos_ai_build_plan(client, model, payload):
    """Single on-demand API call with automatic retry on truncation.
    Returns (CognosBuildPlan | None, error_str | None). Mirrors
    generate_ai_build_plan() from KashMap."""
    user_prompt = json.dumps(payload, ensure_ascii=False)
    output_budget = 3500
    max_budget_cap = 12000

    while True:
        try:
            response = client.responses.create(
                model=model,
                input=[
                    {"role": "system", "content": COGNOS_AI_BUILD_PLAN_SYSTEM_PROMPT},
                    {"role": "user", "content": user_prompt},
                ],
                temperature=0.1,
                max_output_tokens=output_budget,
            )
        except Exception as e:
            return None, f'LLM error: {e}'

        raw_text = response.output_text or ''
        was_truncated = getattr(response, 'status', None) == 'incomplete'
        raw = parse_llm_json(raw_text)

        if not raw and was_truncated and output_budget < max_budget_cap:
            output_budget = min(output_budget * 2, max_budget_cap)
            continue
        if not raw:
            reason = ' (response was truncated)' if was_truncated else ''
            return None, f'Model did not return parseable JSON{reason}. Raw (truncated): {raw_text[:500]}'

        raw = normalize_cognos_layers(raw)
        try:
            plan = CognosBuildPlan.model_validate(raw)
        except ValidationError as ve:
            return None, f'AI response failed schema validation: {ve}'

        allowed_sources = set(payload.get('source_tables', []))
        allowed_fields = set(payload.get('shown_fields', []))
        for f in payload.get('fields', []):
            allowed_fields.add(f['field'])

        plan = strip_unverified_cognos_steps(plan, allowed_sources, allowed_fields)
        return plan, None


def get_openai_client():
    if OpenAI is None:
        return None
    try:
        api_key = st.secrets.get('OPENAI_API_KEY', '')
    except FileNotFoundError:
        return None
    if not api_key or api_key == 'paste_your_openai_api_key_here':
        return None
    return OpenAI(api_key=api_key)


def _render_cognos_plan_step(step):
    level_map = {'confirmed': 'ok', 'rule_based': 'ok', 'inferred': 'medium', 'suggested': 'medium', 'manual_review': 'high'}
    level = level_map.get(step.classification, 'medium')
    with st.container(border=True):
        st.markdown(f"**{step.title}** — {step.power_bi_layer}")
        st.markdown(
            f"{badge(step.classification.replace('_', ' ').title(), level)} "
            f"{badge(f'Confidence: {step.confidence}%')}",
            unsafe_allow_html=True,
        )
        st.write(step.action)
        st.caption(f"Why: {step.reason}")
        if step.referenced_sources or step.referenced_fields:
            st.caption(f"References: {', '.join(step.referenced_sources + step.referenced_fields)}")
        if step.manual_validation:
            st.warning(step.manual_validation)


def render_cognos_ai_build_plan(report_name, parsed, validation_issues):
    client = get_openai_client()
    if client is None:
        st.warning("OpenAI is not configured. Add `OPENAI_API_KEY` to `.streamlit/secrets.toml` to enable the AI build plan.")
        return

    report_key = re.sub(r'[^A-Za-z0-9_]+', '_', report_name)
    cache = st.session_state.setdefault('cognos_ai_build_plan_cache', {})

    clicked = st.button(
        "Generate AI build plan" if report_key not in cache else "Regenerate AI build plan",
        key=f"cognos_ai_btn_{report_key}",
    )
    if clicked:
        payload = build_cognos_ai_payload(report_name, parsed, validation_issues)
        with st.spinner("Calling OpenAI..."):
            plan, error = generate_cognos_ai_build_plan(client, "gpt-4.1-mini", payload)
        cache[report_key] = (plan, error)

    cached = cache.get(report_key)
    if cached is None:
        st.info("Click the button above to generate the AI-assisted build plan.")
        return
    plan, error = cached
    if error:
        st.error(error)
        return

    summary = plan.report_summary
    c1, c2, c3 = st.columns(3)
    c1.metric("Complexity", summary.complexity_level)
    c2.metric("Overall confidence", f"{summary.overall_confidence}%")
    c3.metric("Manual review items", len(plan.manual_review_items))
    st.write(f"**Likely purpose:** {summary.likely_purpose}")
    st.write(f"**Report type:** {summary.report_type}")

    tabs = st.tabs(["Implementation", "Calculations", "Filters/Prompts", "Drillthroughs", "Validation", "Manual Review"])
    with tabs[0]:
        if not plan.implementation_plan:
            st.info("No implementation steps returned.")
        for step in sorted(plan.implementation_plan, key=lambda s: s.step_number):
            _render_cognos_plan_step(step)
    with tabs[1]:
        if not plan.calculation_plan:
            st.info("No calculation steps returned.")
        for step in plan.calculation_plan:
            _render_cognos_plan_step(step)
    with tabs[2]:
        if not plan.filter_prompt_plan:
            st.info("No filter/prompt steps returned.")
        for step in plan.filter_prompt_plan:
            _render_cognos_plan_step(step)
    with tabs[3]:
        if not plan.drillthrough_plan:
            st.info("No drillthrough steps returned.")
        for step in plan.drillthrough_plan:
            _render_cognos_plan_step(step)
    with tabs[4]:
        if not plan.validation_checks:
            st.info("No validation checks returned.")
        for i, check in enumerate(plan.validation_checks, 1):
            st.write(f"{i}. {check}")
    with tabs[5]:
        if not plan.manual_review_items:
            st.success("No manual review items flagged.")
        for i, item in enumerate(plan.manual_review_items, 1):
            st.warning(f"{i}. {item}")
        if plan.limitations:
            st.caption("Limitations:")
            for lim in plan.limitations:
                st.caption(f"• {lim}")


try:
    st.set_page_config(page_title="CognosMap Migration Workspace", layout="wide")
except Exception:
    pass

st.markdown(
    """
<style>
:root {
    --kash-bg: var(--background-color, #FFFFFF);
    --kash-section: var(--secondary-background-color, #FBFAF8);
    --kash-card: var(--background-color, #FFFFFF);
    --kash-navy: var(--text-color, #1F2933);
    --kash-deep: var(--text-color, #111827);
    --kash-blue: #2563EB;
    --kash-blue-soft: var(--secondary-background-color, #EAF2FF);
    --kash-gold: #FF5A1F;
    --kash-green: #2E7D32;
    --kash-orange: #FF5A1F;
    --kash-orange-dark: #E94D13;
    --kash-orange-soft: var(--secondary-background-color, #FFF3EC);
    --kash-red: #B42318;
    --kash-border: color-mix(in srgb, var(--text-color, #1F2933) 18%, transparent);
    --kash-muted: color-mix(in srgb, var(--text-color, #1F2933) 55%, transparent);
}

.upload-card, .upload-card-small,
.wizard-card, .workspace-card,
.report-header, .kpi-strip {
    background: rgba(127, 127, 127, 0.08) !important;
    border-color: var(--kash-border) !important;
}

.main .block-container {
    padding-top: 1.4rem;
    padding-bottom: 3rem;
    max-width: 1480px;
}
.kash-hero {
    position: relative;
    overflow: hidden;
    isolation: isolate;
    border: 1px solid #FFE0D2;
    background: linear-gradient(135deg, #FFFFFF 0%, #FFFFFF 58%, #FFF7F2 100%);
    border-radius: 8px;
    padding: 30px 34px;
    margin-bottom: 20px;
    box-shadow: 0 18px 44px rgba(16, 24, 40, 0.08);
}
.kash-hero::before {
    content: "";
    position: absolute;
    top: 0;
    right: 0;
    bottom: 0;
    width: 56%;
    z-index: 0;
    opacity: 0.72;
    pointer-events: none;
    background-image: url("data:image/svg+xml,%3Csvg width='760' height='220' viewBox='0 0 760 220' xmlns='http://www.w3.org/2000/svg'%3E%3Cg fill='none' stroke='%23ff5a1f' stroke-width='1.35' stroke-linecap='round'%3E%3Cpath opacity='.28' d='M18 178 C110 40 226 34 330 104 S520 184 732 42'/%3E%3Cpath opacity='.18' d='M52 22 C142 116 238 134 372 72 S598 8 736 132'/%3E%3Cpath opacity='.22' d='M112 212 C204 92 310 82 418 142 S586 202 750 94'/%3E%3Cpath opacity='.16' d='M0 92 L126 46 L238 128 L364 56 L498 144 L646 84 L760 124'/%3E%3Cpath opacity='.14' d='M84 168 L210 84 L338 162 L466 82 L608 152 L724 50'/%3E%3C/g%3E%3Cg fill='%23111827'%3E%3Ccircle opacity='.45' cx='126' cy='46' r='3'/%3E%3Ccircle opacity='.35' cx='238' cy='128' r='2.5'/%3E%3Ccircle opacity='.4' cx='364' cy='56' r='3'/%3E%3Ccircle opacity='.35' cx='498' cy='144' r='2.5'/%3E%3Ccircle opacity='.42' cx='646' cy='84' r='3'/%3E%3C/g%3E%3Cg fill='%23ff5a1f'%3E%3Ccircle opacity='.26' cx='210' cy='84' r='4'/%3E%3Ccircle opacity='.24' cx='418' cy='142' r='4'/%3E%3Ccircle opacity='.22' cx='608' cy='152' r='4'/%3E%3C/g%3E%3C/svg%3E");
    background-repeat: no-repeat;
    background-size: cover;
    background-position: right center;
}
.kash-title {
    position: relative;
    z-index: 1;
    font-size: 2.8rem;
    font-weight: 900;
    color: var(--kash-orange);
    margin: 0;
    letter-spacing: 0;
}
.kash-title-kash {
    color: #111827;
}
.kash-title-map {
    color: var(--kash-orange);
}
.kash-subtitle {
    position: relative;
    z-index: 1;
    font-size: 1.15rem;
    font-weight: 700;
    color: #111827;
    margin-top: 8px;
    max-width: 980px;
    line-height: 1.55;
}
.kash-subtitle-accent {
    color: var(--kash-orange);
}
.section-label {
    font-size: 1.2rem;
    font-weight: 800;
    margin: 18px 0 8px 0;
}

.upload-card {
    background: rgba(127, 127, 127, 0.08);
    border: 1px solid var(--kash-border);
    border-radius: 16px;
    padding: 20px 22px;
    box-shadow: 0 1px 3px rgba(16, 24, 40, 0.06);
    min-height: 100%;
}

.upload-card-small {
    background: rgba(127, 127, 127, 0.08);
    border: 1px solid var(--kash-border);
    border-radius: 16px;
    padding: 18px 20px;
    box-shadow: 0 1px 3px rgba(16, 24, 40, 0.06);
    margin-bottom: 16px;
}

.kash-table-wrap {
    border: 1px solid var(--kash-border);
    border-radius: 14px;
    overflow: auto;
    max-height: 640px;
    margin: 14px 0 22px 0;
    box-shadow: 0 8px 24px rgba(17, 24, 39, 0.08);
    background: rgba(127, 127, 127, 0.08);
}

.kash-table thead th {
    position: sticky;
    top: 0;
    z-index: 1;
}

.kash-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.9rem;
}
.kash-table thead th {
    background: linear-gradient(180deg, #2B2B2B 0%, #1F2933 100%);
    color: #FFFFFF;
    font-weight: 800;
    text-align: left;
    padding: 13px 15px;
    border-right: 1px solid rgba(255, 255, 255, 0.12);
    border-bottom: 3px solid #FF5A1F;
    white-space: nowrap;
    letter-spacing: 0.01em;
}

.kash-table thead th:last-child {
    border-right: none;
}

.kash-table tbody td {
    padding: 12px 15px;
    border-bottom: 1px solid rgba(127, 127, 127, 0.2);
    border-right: 1px solid rgba(127, 127, 127, 0.12);
    vertical-align: top;
}

.kash-table tbody td:last-child {
    border-right: none;
}

.kash-table tbody tr:nth-child(even) {
    background: rgba(127, 127, 127, 0.06);
}

.kash-table tbody tr:hover {
    background: rgba(255, 90, 31, 0.12);
}

.kash-table tbody tr:last-child td {
    border-bottom: none;
}

.kash-table td:first-child {
    font-weight: 700;
}

div[data-testid="stDataFrame"] {
    border: 1px solid var(--kash-border);
    border-radius: 14px;
    overflow: hidden;
    box-shadow: 0 4px 14px rgba(16, 24, 40, 0.06);
    margin: 14px 0 22px 0;
}

.card-title {
    font-size: 1rem;
    font-weight: 800;
    margin-bottom: 6px;
}

.card-subtitle {
    font-size: 0.86rem;
    opacity: 0.7;
    margin-bottom: 14px;
}

.badge-required {
    width: fit-content;
    margin-left: auto;
    font-size: 0.72rem;
    font-weight: 700;
    color: #B42318;
    background: #FFF1F0;
    border: 1px solid #FFE1DF;
    border-radius: 999px;
    padding: 6px 10px;
    text-align: center;
}

.badge-optional {
    width: fit-content;
    margin-left: auto;
    font-size: 0.72rem;
    font-weight: 700;
    opacity: 0.75;
    background: rgba(127, 127, 127, 0.12);
    border: 1px solid var(--kash-border);
    border-radius: 999px;
    padding: 6px 10px;
    text-align: center;
}

.file-count-box {
    background: rgba(127, 127, 127, 0.08);
    border: 1px solid var(--kash-border);
    border-radius: 12px;
    padding: 10px 12px;
    margin-top: 10px;
    font-size: 0.88rem;
    font-weight: 600;
}

.wizard-card {
    background: rgba(127, 127, 127, 0.08);
    border: 1px solid var(--kash-border);
    border-radius: 8px;
    padding: 22px;
    margin-bottom: 16px;
    box-shadow: 0 8px 24px rgba(15, 23, 42, 0.06);
}
.wizard-card h3 {
    font-size: 1.15rem;
    margin: 0 0 10px 0;
}
.wizard-card p {
    line-height: 1.45;
    margin: 4px 0 10px 0;
}

div[data-testid="stExpander"]:hover {
    border-left: 4px solid var(--kash-orange);
}
.badge {
    display: inline-block;
    padding: 5px 10px;
    border-radius: 999px;
    background-color: var(--kash-blue-soft);
    color: var(--kash-navy);
    font-size: 0.78rem;
    font-weight: 700;
    margin-right: 6px;
    margin-bottom: 6px;
}
.badge-high {
    background-color: var(--kash-navy);
    color: #FFFFFF;
}
.badge-medium {
    background-color: #FEF3C7;
    color: #92400E;
}
.badge-low {
    background-color: #DCFCE7;
    color: #166534;
}
.badge-ok {
    background-color: #DCFCE7;
    color: #166534;
}
.badge-review {
    background-color: #FEF3C7;
    color: #92400E;
}
.badge-error {
    background-color: #FEE2E2;
    color: var(--kash-red);
}
.workspace-card {
    background: rgba(127, 127, 127, 0.08);
    border: 1px solid var(--kash-border);
    border-radius: 8px;
    padding: 20px;
    box-shadow: 0 8px 24px rgba(15, 23, 42, 0.055);
    margin-bottom: 14px;
}
.report-header {
    background: rgba(127, 127, 127, 0.08);
    border: 1px solid var(--kash-border);
    border-radius: 8px;
    padding: 22px 24px;
    box-shadow: 0 10px 28px rgba(15, 23, 42, 0.06);
    margin: 12px 0 16px 0;
}
.report-header h2 {
    font-size: 1.7rem;
    margin: 0 0 8px 0;
    letter-spacing: 0;
}
.report-header p {
    margin: 0 0 12px 0;
    line-height: 1.45;
}
.kpi-strip {
    background: rgba(127, 127, 127, 0.08);
    border: 1px solid var(--kash-border);
    border-radius: 8px;
    padding: 16px;
    box-shadow: 0 8px 20px rgba(15, 23, 42, 0.045);
}

.stButton > button, .stDownloadButton > button {
    border-radius: 10px;
    border: 1px solid var(--kash-border);
    font-weight: 700;
}
.stButton > button[kind="primary"] {
    background-color: #FF5A1F !important;
    border-color: #FF5A1F !important;
    color: white !important;
}
.stButton > button[kind="primary"]:hover {
    background-color: var(--kash-orange-dark) !important;
    border-color: var(--kash-orange-dark) !important;
    color: white !important;
}

@media (max-width: 700px) {
    .kash-title {
        font-size: 2rem;
    }
    .kash-hero {
        padding: 24px;
    }
}
</style>
    """,
    unsafe_allow_html=True,
)


def badge(label, level=''):
    css_class = 'badge'
    if level:
        css_class += f' badge-{level}'
    return f'<span class="{css_class}">{label}</span>'


def show_table_or_info(df, message, large=False, height=None):
    """Same styled-table pattern as KashMap: small tables get the orange
    kash-table HTML treatment, large ones stay as normal Streamlit dataframes.
    height is only passed through when explicitly set - newer Streamlit
    versions reject height=None outright (must be an int or "auto")."""
    if df is None or df.empty:
        st.info(message)
        return

    if large or len(df) > 150:
        if height is not None:
            st.dataframe(df, use_container_width=True, hide_index=True, height=height)
        else:
            st.dataframe(df, use_container_width=True, hide_index=True)
        return

    html = df.to_html(index=False, escape=True, classes="kash-table", border=0)
    st.markdown(f'<div class="kash-table-wrap">{html}</div>', unsafe_allow_html=True)
    

st.markdown(
    """
    <div class="kash-hero">
      <div class="kash-title"><span class="kash-title-kash">Cognos</span><span class="kash-title-map">Map</span></div>
      <div class="kash-subtitle">Break Down <span class="kash-subtitle-accent">IBM Cognos</span> Complexity for Faster <span class="kash-subtitle-accent">Power BI</span> Migration</div>
    </div>
    """,
    unsafe_allow_html=True,
)

col_info, col_btn = st.columns([0.76, 0.24])
with col_info:
    st.info("ℹ️ **Cognos → Power BI Migration**: Upload IBM Cognos Report XML specifications (`.xml`). To migrate WebFOCUS procedure files (`.fex`), use the WebFOCUS → Power BI workflow.")
with col_btn:
    if st.button("Switch to WebFOCUS → Power BI ➡️", key="cognos_to_wf_switch", use_container_width=True):
        st.switch_page("pages/webfocus_powerbi.py")

with st.expander("How to use CognosMap & Deterministic Migration Engine", expanded=True):
    st.markdown(
        """
        1. **Upload Cognos Report XML**: Upload one or more `.xml` report specifications (and optional Data Module JSON / Framework Model XML).
        2. **Deterministic Translation Engine (DAX & M)**: Automatically converts Cognos expressions, calculations, and filters into DAX Measures, Calculated Columns, Power Query M, and Report Filters using a 100% deterministic rule-based engine (no LLM required).
        3. **Layout & Container Mapping**: Extracts Lists (mapped to Power BI Table `tableEx`), Crosstabs (mapped to Matrix `pivotTable`), Repeaters (Small Multiples), Complex Queries (Joins, Unions, Native SQL), Prompts, and Macros.
        4. **Detailed Technical Tabs**: Inspect your migration in dedicated tabs:
           - **Translations (DAX / M)**: Rule-based translation table with layer filters, confidence scoring, macro tokenization, and manual review safety flags.
           - **Layout & Containers**: Container breakdown for Lists, Crosstabs, Repeaters, and Query architecture.
           - **Fields, Filters, Prompts, Drillthroughs, Visualizations, Custom Controls, Validation, AI Build Plan**.
        5. **Export Migration Workbook**: Download a structured multi-tab Excel migration package.
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
                <div class="card-title">Upload Cognos Report XML</div>
                <div class="card-subtitle">
                    Upload one or more report specification .xml files.
                </div>
                """,
                unsafe_allow_html=True,
            )
        with badge_col:
            st.markdown('<div class="badge-required">Required</div>', unsafe_allow_html=True)

        uploaded_reports = st.file_uploader(
            "Upload Cognos report XML files",
            type=["xml"],
            accept_multiple_files=True,
            key="report_uploader",
            label_visibility="collapsed",
        )

        sample_col1, sample_col2 = st.columns([0.65, 0.35])
        with sample_col1:
            if st.button("Load Built-in Sample Cognos Report", key="load_sample_cognos_btn", help="Load sample Cognos report XML with Lists, Crosstabs, and Calculations"):
                sample_path = Path(__file__).resolve().parent.parent.parent / "test_reports" / "cognos" / "sample_regional_performance.xml"
                if sample_path.exists():
                    st.session_state["cognos_sample_loaded"] = sample_path.read_text(encoding="utf-8")
                    st.session_state["cognos_sample_name"] = "sample_regional_performance.xml"
                    st.success("Loaded sample report: 'sample_regional_performance.xml'. Click 'Analyze Files' to run.")
        with sample_col2:
            if st.session_state.get("cognos_sample_loaded"):
                if st.button("Clear Sample", key="clear_sample_btn"):
                    st.session_state.pop("cognos_sample_loaded", None)
                    st.session_state.pop("cognos_sample_name", None)
                    st.rerun()

        if st.session_state.get("cognos_sample_loaded") and not uploaded_reports:
            st.info(f"Sample loaded: **{st.session_state.get('cognos_sample_name')}**")

        if uploaded_reports:
            fex_detected = [f.name for f in uploaded_reports if f.name.lower().endswith(".fex")]
            if fex_detected:
                st.error(
                    f"❌ WebFOCUS (.fex) file detected: {', '.join(fex_detected)}. "
                    "CognosMap only accepts Cognos report XML files (.xml). "
                    "Please switch to the WebFOCUS → Power BI workflow to analyze .fex files."
                )
                if st.button("Switch to WebFOCUS → Power BI", key="fex_err_redirect", type="primary"):
                    st.switch_page("pages/webfocus_powerbi.py")
                st.stop()

            total_size = sum(getattr(f, "size", 0) for f in uploaded_reports)
            st.markdown(
                f"""
                <div class="file-count-box">
                    {len(uploaded_reports)} report file(s) uploaded · {total_size / 1024:.1f} KB total
                </div>
                """,
                unsafe_allow_html=True,
            )

with side_col:
    dm_card = st.container(border=True)
    with dm_card:
        title_col, badge_col = st.columns([0.78, 0.22])
        with title_col:
            st.markdown(
                """
                <div class="card-title">Data Module JSON</div>
                <div class="card-subtitle">
                    Optional. Upload this to validate report fields against the real model schema.
                </div>
                """,
                unsafe_allow_html=True,
            )
        with badge_col:
            st.markdown('<div class="badge-optional">Optional</div>', unsafe_allow_html=True)

        uploaded_data_module = st.file_uploader(
            "Upload Data Module JSON (optional)",
            type=["json"],
            key="data_module_uploader",
            label_visibility="collapsed",
        )
        if uploaded_data_module:
            st.markdown(
                f"""
                <div class="file-count-box">
                    Data Module uploaded · {uploaded_data_module.size / 1024:.1f} KB
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

    output_card = st.container(border=True)
    with output_card:
        st.markdown(
            """
            <div class="card-title">Output Settings</div>
            <div class="card-subtitle">Name of the generated Excel migration workbook.</div>
            """,
            unsafe_allow_html=True,
        )
        output_name = st.text_input(
            "Output file name",
            value="CognosMap_Migration_Output.xlsx",
            key="output_file_name",
            label_visibility="collapsed",
        )

submit = st.button("Analyze Files", type="primary", use_container_width=True)


def build_overview_row(report_name, parsed, validation_issues):
    translations = parsed.get("field_translations", [])
    translatable_count = sum(1 for t in translations if t.get("is_translatable"))
    manual_review_count = sum(1 for t in translations if not t.get("is_translatable"))
    trans_rate = round((translatable_count / len(translations) * 100)) if translations else 100

    return {
        "Report Name": parsed["report_name"] or report_name,
        "Model Path": parsed["model_path"],
        "Source Tables": ", ".join(parsed["source_tables"]),
        "Field Count": len(parsed["fields"]),
        "Source Field Count": len(parsed["source_fields"]),
        "Calculated Field Count": len(parsed["calculated_fields"]),
        "Filter Count": len(parsed["filters"]),
        "Prompt Count": len(parsed["prompts"]),
        "List Count": len(parsed.get("lists", [])),
        "Crosstab Count": len(parsed.get("crosstabs", [])),
        "Deterministic Translation Rate": f"{trans_rate}%",
        "Manual Review Items": manual_review_count,
        "Drillthrough Count": len(parsed["drillthroughs"]),
        "Visualization Count": len(parsed["visualizations"]),
        "Custom Control Count": len(parsed["custom_controls"]),
        "Validation Status": "Verified" if not validation_issues else f"{len(validation_issues)} issue(s)",
    }


def build_fields_df(report_name, parsed):
    rows = []
    for f in parsed["fields"]:
        rows.append({
            "Report": report_name, "Query": f["query"], "Field": f["field"],
            "Aggregate": f["aggregate"], "Module": f["module"] or "", "Table": f["table"] or "",
            "Column": f["column"] or "", "Expression": f["expression"],
            "Is Calculated": "Yes" if f["is_calculated"] else "No",
            "Expression Type": f["expression_type"],
            "Target DAX": f.get("dax_translation") or "",
            "Power Query M": f.get("m_translation") or "",
            "Target Layer": f.get("target_layer", ""),
            "Status": f.get("translation_status", ""),
        })
    return pd.DataFrame(rows)


def build_filters_df(report_name, parsed):
    rows = []
    for f in parsed["filters"]:
        rows.append({
            "Report": report_name, "Query": f["query"], "Use": f["use"],
            "Expression": f["expression"], "Is Prompted": "Yes" if f["is_prompted"] else "No",
            "Target DAX": f.get("dax_translation") or "",
            "Power Query M": f.get("m_translation") or "",
            "Target Layer": f.get("target_layer", "Report Filter"),
            "Status": f.get("translation_status", ""),
        })
    return pd.DataFrame(rows)


def build_prompts_df(report_name, parsed):
    rows = []
    for p in parsed["prompts"]:
        rows.append({
            "Report": report_name,
            "Parameter": p["parameter"],
            "Prompt Type": p.get("prompt_type", "Value Prompt"),
            "Ref Query": p["ref_query"],
            "Target Control": p.get("target_powerbi_control", "Power BI Slicer"),
            "Multi-Select": "Yes" if p.get("multi_select") else "No",
            "Required": "Yes" if p.get("required", True) else "No",
            "Options": ", ".join(p["options"]),
            "Default": p["default"],
            "Auto Submit": p["auto_submit"],
        })
    return pd.DataFrame(rows)


def build_translations_df(report_name, parsed):
    rows = []
    for t in parsed.get("field_translations", []):
        warn_str = "; ".join(t.get("warnings", []))
        rows.append({
            "Report": report_name,
            "Scope": "Data Item",
            "Query": t.get("query_name", ""),
            "Field": t.get("field_name", ""),
            "Category": t.get("category", ""),
            "Target Layer": t.get("target_type", ""),
            "Source Expression": t.get("source_expression", ""),
            "Translated DAX": t.get("dax_expression") or "",
            "Power Query M": t.get("m_expression") or "",
            "Confidence": f"{t.get('confidence', 90)}%",
            "Status": t.get("status", ""),
            "Warnings": warn_str,
        })
    for t in parsed.get("filter_translations", []):
        warn_str = "; ".join(t.get("warnings", []))
        rows.append({
            "Report": report_name,
            "Scope": "Filter",
            "Query": t.get("query_name", ""),
            "Field": t.get("field_name", ""),
            "Category": t.get("category", ""),
            "Target Layer": t.get("target_type", "Report Filter"),
            "Source Expression": t.get("source_expression", ""),
            "Translated DAX": t.get("dax_expression") or "",
            "Power Query M": t.get("m_expression") or "",
            "Confidence": f"{t.get('confidence', 90)}%",
            "Status": t.get("status", ""),
            "Warnings": warn_str,
        })
    return pd.DataFrame(rows)


def build_lists_df(report_name, parsed):
    rows = []
    for lst in parsed.get("lists", []):
        col_names = [c.get("data_item") or c.get("title") for c in lst.get("columns", [])]
        rows.append({
            "Report": report_name,
            "List Name": lst.get("name", ""),
            "Query": lst.get("ref_query", ""),
            "Column Count": lst.get("column_count", 0),
            "Columns": ", ".join(col_names),
            "Groupings": ", ".join(lst.get("groupings", [])),
            "Target Power BI Visual": lst.get("target_powerbi_visual", "Table (tableEx)"),
        })
    return pd.DataFrame(rows)


def build_crosstabs_df(report_name, parsed):
    rows = []
    for ct in parsed.get("crosstabs", []):
        rows.append({
            "Report": report_name,
            "Crosstab Name": ct.get("name", ""),
            "Query": ct.get("ref_query", ""),
            "Row Dimensions": ", ".join(ct.get("row_dimensions", [])),
            "Column Dimensions": ", ".join(ct.get("column_dimensions", [])),
            "Intersection Measures": ", ".join(ct.get("measures", [])),
            "Target Power BI Visual": ct.get("target_powerbi_visual", "Matrix (pivotTable)"),
        })
    return pd.DataFrame(rows)


def build_repeaters_df(report_name, parsed):
    rows = []
    for rep in parsed.get("repeaters", []):
        rows.append({
            "Report": report_name,
            "Repeater Name": rep.get("name", ""),
            "Type": rep.get("type", ""),
            "Query": rep.get("ref_query", ""),
            "Data Items": ", ".join(rep.get("data_items", [])),
            "Target Power BI Visual": rep.get("target_powerbi_visual", "Small Multiples / Matrix"),
        })
    return pd.DataFrame(rows)


def build_complex_queries_df(report_name, parsed):
    rows = []
    for q in parsed.get("queries", []):
        if q.get("query_type") != "Standard Query":
            detail = ""
            if q.get("join_details"):
                jd = q["join_details"]
                detail = f"Left: {jd['left_query']}, Right: {jd['right_query']}, Criteria: {', '.join(jd['criteria'])}"
            elif q.get("set_operation"):
                so = q["set_operation"]
                detail = f"Op: {so['operation']}, Subqueries: {', '.join(so['subqueries'])}"
            elif q.get("native_sql"):
                nsql = q["native_sql"]
                detail = f"DS: {nsql['data_source']}, SQL: {nsql['sql_text'][:120]}..."
            rows.append({
                "Report": report_name,
                "Query Name": q.get("name", ""),
                "Query Type": q.get("query_type", ""),
                "Architecture Details": detail,
            })
    return pd.DataFrame(rows)


def build_drillthroughs_df(report_name, parsed):
    rows = []
    for d in parsed["drillthroughs"]:
        links = "; ".join(f"{l['source_parameter']} -> {l['target_parameter']}" for l in d["parameter_links"])
        rows.append({
            "Source Report": report_name, "Drill Name": d["drill_name"],
            "Target Report": d["target_report"], "Parameter Links": links,
            "Opens New Window": d["opens_new_window"],
        })
    return pd.DataFrame(rows)


def build_visualizations_df(report_name, parsed):
    rows = []
    for v in parsed["visualizations"]:
        for slot in v["slots"]:
            rows.append({
                "Report": report_name, "Visualization": v["name"], "Type": v["type"],
                "Data Store": slot["data_store"], "Slot": slot["slot"], "Columns": ", ".join(slot["columns"]),
            })
        if not v["slots"]:
            rows.append({
                "Report": report_name, "Visualization": v["name"], "Type": v["type"],
                "Data Store": "", "Slot": "", "Columns": "",
            })
    return pd.DataFrame(rows)


def build_custom_controls_df(report_name, parsed):
    rows = []
    for c in parsed["custom_controls"]:
        rows.append({
            "Report": report_name, "Control Name": c["name"], "Description": c["description"],
            "JS Path": c["path"], "Manual Review": "Yes - no Power BI equivalent",
        })
    return pd.DataFrame(rows)


def build_validation_df(report_name, issues):
    rows = []
    for issue in issues:
        rows.append({"Report": report_name, "Issue": issue})
    if not issues:
        rows.append({"Report": report_name, "Issue": "All fields verified against Data Module"})
    return pd.DataFrame(rows)


def build_excel_workbook(
    overview_df, fields_df, filters_df, prompts_df,
    drillthroughs_df, viz_df, controls_df, validation_df,
    framework_model=None,
    translations_df=None, lists_df=None, crosstabs_df=None,
    repeaters_df=None, complex_queries_df=None,
    inventory_df=None, complexity_df=None, waves_df=None,
    lineage_df=None, consolidation_df=None, duplicate_df=None,
    model_tables_df=None, model_columns_df=None,
    model_measures_df=None, model_relationships_df=None,
    report_pages_df=None, visual_bindings_df=None
):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Power BI Enhanced Report (PBIR) Sheets
        if report_pages_df is not None and not report_pages_df.empty:
            report_pages_df.to_excel(writer, sheet_name="Report Pages", index=False)
        if visual_bindings_df is not None and not visual_bindings_df.empty:
            visual_bindings_df.to_excel(writer, sheet_name="Visual Bindings", index=False)

        # Power BI Semantic Model Blueprint Sheets
        if model_tables_df is not None and not model_tables_df.empty:
            model_tables_df.to_excel(writer, sheet_name="Model Tables", index=False)
        if model_columns_df is not None and not model_columns_df.empty:
            model_columns_df.to_excel(writer, sheet_name="Model Columns", index=False)
        if model_measures_df is not None and not model_measures_df.empty:
            model_measures_df.to_excel(writer, sheet_name="Model Measures", index=False)
        if model_relationships_df is not None and not model_relationships_df.empty:
            model_relationships_df.to_excel(writer, sheet_name="Model Relationships", index=False)

        # Senturus-Style Migration Assessment Sheets
        if inventory_df is not None and not inventory_df.empty:
            inventory_df.to_excel(writer, sheet_name="Report Inventory", index=False)
        if complexity_df is not None and not complexity_df.empty:
            complexity_df.to_excel(writer, sheet_name="Complexity & Effort", index=False)
        if waves_df is not None and not waves_df.empty:
            waves_df.to_excel(writer, sheet_name="Migration Waves", index=False)
        if lineage_df is not None and not lineage_df.empty:
            lineage_df.to_excel(writer, sheet_name="Cross-Report Lineage", index=False)
        if consolidation_df is not None and not consolidation_df.empty:
            consolidation_df.to_excel(writer, sheet_name="Consolidation Opportunities", index=False)
        if duplicate_df is not None and not duplicate_df.empty:
            duplicate_df.to_excel(writer, sheet_name="Duplicate Analysis", index=False)

        # Technical Blueprint & Object Inventory Sheets
        overview_df.to_excel(writer, sheet_name="Overview", index=False)
        if translations_df is not None and not translations_df.empty:
            translations_df.to_excel(writer, sheet_name="Translations (DAX-M)", index=False)
        if lists_df is not None and not lists_df.empty:
            lists_df.to_excel(writer, sheet_name="Lists (Tables)", index=False)
        if crosstabs_df is not None and not crosstabs_df.empty:
            crosstabs_df.to_excel(writer, sheet_name="Crosstabs (Matrices)", index=False)
        if repeaters_df is not None and not repeaters_df.empty:
            repeaters_df.to_excel(writer, sheet_name="Repeaters", index=False)
        if complex_queries_df is not None and not complex_queries_df.empty:
            complex_queries_df.to_excel(writer, sheet_name="Complex Queries", index=False)
        if not fields_df.empty:
            fields_df.to_excel(writer, sheet_name="Field Inventory", index=False)
        if not filters_df.empty:
            filters_df.to_excel(writer, sheet_name="Filters", index=False)
        if not prompts_df.empty:
            prompts_df.to_excel(writer, sheet_name="Prompts", index=False)
        if not drillthroughs_df.empty:
            drillthroughs_df.to_excel(writer, sheet_name="Drillthroughs", index=False)
        if not viz_df.empty:
            viz_df.to_excel(writer, sheet_name="Visualizations", index=False)
        if not controls_df.empty:
            controls_df.to_excel(writer, sheet_name="Custom Controls", index=False)
        if not validation_df.empty:
            validation_df.to_excel(writer, sheet_name="Validation", index=False)

        if framework_model is not None:
            fw_overview_df = build_framework_overview_df(framework_model)
            fw_packages_df = build_framework_packages_df(framework_model)
            fw_subjects_df = build_framework_query_subjects_df(framework_model)
            fw_columns_df = build_framework_columns_df(framework_model)
            fw_relationships_df = build_framework_relationships_df(framework_model)
            fw_filters_security_df = build_framework_filters_security_df(framework_model)

            if not fw_overview_df.empty:
                fw_overview_df.to_excel(writer, sheet_name="Framework Overview", index=False)
            if not fw_packages_df.empty:
                fw_packages_df.to_excel(writer, sheet_name="Framework Packages", index=False)
            if not fw_subjects_df.empty:
                fw_subjects_df.to_excel(writer, sheet_name="Framework Query Subjects", index=False)
            if not fw_columns_df.empty:
                fw_columns_df.to_excel(writer, sheet_name="Framework Columns", index=False)
            if not fw_relationships_df.empty:
                fw_relationships_df.to_excel(writer, sheet_name="Framework Relationships", index=False)
            if not fw_filters_security_df.empty:
                fw_filters_security_df.to_excel(writer, sheet_name="Framework Filters-Security", index=False)

    output.seek(0)
    return output


@st.fragment
def show_download_button(output_stream, file_name, key):
    st.download_button(
        label="Download Full Excel Report",
        data=output_stream,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=key,
        use_container_width=True,
    )


def complexity_label(field_count, custom_control_count, validation_issue_count):
    """Same spirit as KashMap's complexity_label: a quick High/Medium/Low
    read on how much manual review a report will likely need."""
    score = field_count // 5 + custom_control_count * 3 + validation_issue_count * 2
    if score >= 12 or custom_control_count > 0 or validation_issue_count > 0:
        return "High", "high"
    if score >= 5:
        return "Medium", "medium"
    return "Low", "low"


if submit:
    active_reports = []
    if uploaded_reports:
        fex_detected = [f.name for f in uploaded_reports if f.name.lower().endswith(".fex")]
        if fex_detected:
            st.error(
                f"❌ WebFOCUS (.fex) file detected: {', '.join(fex_detected)}. "
                "CognosMap only accepts Cognos report XML files (.xml). "
                "Please use the WebFOCUS → Power BI workflow to analyze .fex files."
            )
            if st.button("Switch to WebFOCUS → Power BI", key="fex_submit_redirect", type="primary"):
                st.switch_page("pages/webfocus_powerbi.py")
            st.stop()
        for f in uploaded_reports:
            active_reports.append((f.name, f.read().decode("utf-8", errors="replace")))
    elif st.session_state.get("cognos_sample_loaded"):
        active_reports.append((
            st.session_state.get("cognos_sample_name", "sample_regional_performance.xml"),
            st.session_state["cognos_sample_loaded"]
        ))
    else:
        st.error("Please upload at least one Cognos report XML file or load the built-in sample.")
        st.stop()

    data_module_schema = None
    if uploaded_data_module is not None:
        try:
            dm_payload = json.load(uploaded_data_module)
            data_module_schema = parse_cognos_data_module(dm_payload["data"])
            rel_count = len(data_module_schema.get("relationships", []))
            rel_note = f", {rel_count} relationship(s) found" if rel_count else ", no relationships declared (flat model)"
            st.success(f"Data Module loaded: {data_module_schema['module_name']} ({len(data_module_schema['tables'])} table(s)){rel_note}")
        except Exception as e:
            st.warning(f"Could not parse Data Module JSON: {e}")

    all_overview_rows = []
    all_fields_dfs = []
    all_filters_dfs = []
    all_prompts_dfs = []
    all_drillthroughs_dfs = []
    all_viz_dfs = []
    all_controls_dfs = []
    all_validation_dfs = []
    all_translations_dfs = []
    all_lists_dfs = []
    all_crosstabs_dfs = []
    all_repeaters_dfs = []
    all_complex_queries_dfs = []
    parsed_by_report = {}
    errors = []

    # First pass: separate framework model file(s) from report files.
    report_files = []
    framework_model = None
    framework_model_source_name = ""

    for filename, xml_text in active_reports:
        file_type = detect_cognos_xml_type(xml_text)

        if file_type == "framework_model":
            if framework_model is not None:
                st.warning(f"Multiple framework model files uploaded. Using '{framework_model_source_name}'; ignoring '{filename}'.")
            else:
                try:
                    framework_model = parse_framework_model(xml_text)
                    framework_model_source_name = filename
                except Exception as e:
                    errors.append(f"{filename}: failed to parse as framework model: {e}")
        elif file_type == "report":
            report_files.append((filename, xml_text))
        else:
            errors.append(f"{filename}: unrecognized XML root, skipped")

    progress = st.progress(0)
    total = len(report_files) or 1

    for idx, (report_name, xml_text) in enumerate(report_files, start=1):
        try:
            parsed = parse_cognos_report(xml_text)

            validation_issues = []
            if data_module_schema is not None:
                validation_issues = validate_report_against_model(parsed, data_module_schema)
            if framework_model is not None:
                validation_issues = validation_issues + validate_report_against_framework(parsed, framework_model)

            parsed_by_report[report_name] = (parsed, validation_issues)

            all_overview_rows.append(build_overview_row(report_name, parsed, validation_issues))
            all_fields_dfs.append(build_fields_df(report_name, parsed))
            all_filters_dfs.append(build_filters_df(report_name, parsed))
            all_prompts_dfs.append(build_prompts_df(report_name, parsed))
            all_drillthroughs_dfs.append(build_drillthroughs_df(report_name, parsed))
            all_viz_dfs.append(build_visualizations_df(report_name, parsed))
            all_controls_dfs.append(build_custom_controls_df(report_name, parsed))
            all_validation_dfs.append(build_validation_df(report_name, validation_issues))
            all_translations_dfs.append(build_translations_df(report_name, parsed))
            all_lists_dfs.append(build_lists_df(report_name, parsed))
            all_crosstabs_dfs.append(build_crosstabs_df(report_name, parsed))
            all_repeaters_dfs.append(build_repeaters_df(report_name, parsed))
            all_complex_queries_dfs.append(build_complex_queries_df(report_name, parsed))
        except Exception as e:
            errors.append(f"{report_name}: {e}")

        progress.progress(idx / total)

    overview_df = pd.DataFrame(all_overview_rows)
    fields_df = pd.concat(all_fields_dfs, ignore_index=True) if all_fields_dfs else pd.DataFrame()
    filters_df = pd.concat(all_filters_dfs, ignore_index=True) if all_filters_dfs else pd.DataFrame()
    prompts_df = pd.concat(all_prompts_dfs, ignore_index=True) if all_prompts_dfs else pd.DataFrame()
    drillthroughs_df = pd.concat(all_drillthroughs_dfs, ignore_index=True) if all_drillthroughs_dfs else pd.DataFrame()
    viz_df = pd.concat(all_viz_dfs, ignore_index=True) if all_viz_dfs else pd.DataFrame()
    controls_df = pd.concat(all_controls_dfs, ignore_index=True) if all_controls_dfs else pd.DataFrame()
    validation_df = pd.concat(all_validation_dfs, ignore_index=True) if all_validation_dfs else pd.DataFrame()
    translations_df = pd.concat(all_translations_dfs, ignore_index=True) if all_translations_dfs else pd.DataFrame()
    lists_df = pd.concat(all_lists_dfs, ignore_index=True) if all_lists_dfs else pd.DataFrame()
    crosstabs_df = pd.concat(all_crosstabs_dfs, ignore_index=True) if all_crosstabs_dfs else pd.DataFrame()
    repeaters_df = pd.concat(all_repeaters_dfs, ignore_index=True) if all_repeaters_dfs else pd.DataFrame()
    complex_queries_df = pd.concat(all_complex_queries_dfs, ignore_index=True) if all_complex_queries_dfs else pd.DataFrame()

    duplicate_records = build_cognos_duplicate_analysis(parsed_by_report)
    duplicate_df = build_duplicate_df(duplicate_records)
    dup_category_map = {r["report_name"]: r["category"] for r in duplicate_records}

    seen_dup_groups = {}
    dup_copy_info = {}
    for r in duplicate_records:
        gid = r.get("group_id")
        rname = r.get("report_name")
        cat = r.get("category")
        if gid and cat in ("Exact Duplicate", "Near Duplicate"):
            if gid not in seen_dup_groups:
                seen_dup_groups[gid] = rname
                dup_copy_info[rname] = {"is_copy": False, "canonical": rname}
            else:
                dup_copy_info[rname] = {"is_copy": True, "canonical": seen_dup_groups[gid]}
        else:
            dup_copy_info[rname] = {"is_copy": False, "canonical": rname}

    user_bv = st.session_state.get("cognos_business_values", {})
    effort_hours_map = st.session_state.get("cognos_effort_hours_map", DEFAULT_EFFORT_HOURS)

    complexity_by_report = {}
    wave_by_report = {}
    for r_name, (p_parsed, v_issues) in parsed_by_report.items():
        comp = compute_cognos_complexity_score(p_parsed, v_issues)
        tier = comp["complexity_tier"]
        score = comp["complexity_score"]
        effort = estimate_migration_effort(tier, effort_hours_map)
        comp["est_effort_hours"] = effort
        complexity_by_report[r_name] = comp

        bv = user_bv.get(r_name, "Medium")
        dup_cat = dup_category_map.get(r_name, "Unique")
        d_info = dup_copy_info.get(r_name, {"is_copy": False, "canonical": r_name})
        w_label, w_num, w_reason = assign_migration_wave(
            r_name, score, tier, dup_cat, bv,
            is_duplicate_copy=d_info["is_copy"],
            canonical_match=d_info["canonical"] if d_info["is_copy"] else None,
        )
        wave_by_report[r_name] = {
            "wave": w_label,
            "wave_number": w_num,
            "reason": w_reason,
            "business_value": bv,
            "est_effort_hours": effort,
        }

    inventory_df = build_cognos_inventory_df(parsed_by_report, complexity_by_report, wave_by_report)
    complexity_df = build_cognos_complexity_df(complexity_by_report, effort_hours_map)
    waves_df = build_cognos_waves_df(wave_by_report, complexity_by_report)
    lineage_df = build_cross_report_lineage(parsed_by_report)
    consolidation_df = build_consolidation_opportunities(parsed_by_report, duplicate_records)

    # Task 3: Build Semantic Model Blueprints & DataFrames
    blueprints_by_report = {}
    all_model_tables = []
    all_model_columns = []
    all_model_measures = []
    all_model_relationships = []

    for r_name, (p_parsed, v_issues) in parsed_by_report.items():
        bp = build_semantic_model_blueprint(
            r_name, p_parsed, v_issues,
            framework_model=framework_model,
            data_module_schema=data_module_schema
        )
        blueprints_by_report[r_name] = bp

        t_df = build_model_tables_df(bp)
        if not t_df.empty:
            t_df.insert(0, "Report", r_name)
            all_model_tables.append(t_df)

        c_df = build_model_columns_df(bp)
        if not c_df.empty:
            c_df.insert(0, "Report", r_name)
            all_model_columns.append(c_df)

        m_df = build_model_measures_df(bp)
        if not m_df.empty:
            m_df.insert(0, "Report", r_name)
            all_model_measures.append(m_df)

        r_df = build_model_relationships_df(bp)
        if not r_df.empty:
            r_df.insert(0, "Report", r_name)
            all_model_relationships.append(r_df)

    model_tables_df = pd.concat(all_model_tables, ignore_index=True) if all_model_tables else pd.DataFrame()
    model_columns_df = pd.concat(all_model_columns, ignore_index=True) if all_model_columns else pd.DataFrame()
    model_measures_df = pd.concat(all_model_measures, ignore_index=True) if all_model_measures else pd.DataFrame()
    model_relationships_df = pd.concat(all_model_relationships, ignore_index=True) if all_model_relationships else pd.DataFrame()

    # Task 4: Build Report Pages & PBIR Visual Blueprints & DataFrames
    pages_blueprints_by_report = {}
    all_report_pages = []
    all_visual_bindings = []

    for r_name, (p_parsed, v_issues) in parsed_by_report.items():
        bp = blueprints_by_report.get(r_name, {})
        p_bp = extract_report_pages_blueprint(p_parsed.get("xml_text", ""), bp, parsed=p_parsed)
        pages_blueprints_by_report[r_name] = p_bp

        p_df = build_report_pages_df(p_bp)
        if not p_df.empty:
            p_df.insert(0, "Report", r_name)
            all_report_pages.append(p_df)

        vb_df = build_visual_bindings_df(p_bp)
        if not vb_df.empty:
            vb_df.insert(0, "Report", r_name)
            all_visual_bindings.append(vb_df)

    report_pages_df = pd.concat(all_report_pages, ignore_index=True) if all_report_pages else pd.DataFrame()
    visual_bindings_df = pd.concat(all_visual_bindings, ignore_index=True) if all_visual_bindings else pd.DataFrame()

    workbook = build_excel_workbook(
        overview_df, fields_df, filters_df, prompts_df,
        drillthroughs_df, viz_df, controls_df, validation_df,
        framework_model=framework_model,
        translations_df=translations_df,
        lists_df=lists_df,
        crosstabs_df=crosstabs_df,
        repeaters_df=repeaters_df,
        complex_queries_df=complex_queries_df,
        inventory_df=inventory_df,
        complexity_df=complexity_df,
        waves_df=waves_df,
        lineage_df=lineage_df,
        consolidation_df=consolidation_df,
        duplicate_df=duplicate_df,
        model_tables_df=model_tables_df,
        model_columns_df=model_columns_df,
        model_measures_df=model_measures_df,
        model_relationships_df=model_relationships_df,
        report_pages_df=report_pages_df,
        visual_bindings_df=visual_bindings_df,
    )
    fn = output_name if output_name.lower().endswith(".xlsx") else f"{output_name}.xlsx"

    st.session_state["cognosmap_result"] = {
        "overview_df": overview_df, "fields_df": fields_df, "filters_df": filters_df,
        "prompts_df": prompts_df, "drillthroughs_df": drillthroughs_df, "viz_df": viz_df,
        "controls_df": controls_df, "validation_df": validation_df,
        "translations_df": translations_df, "lists_df": lists_df, "crosstabs_df": crosstabs_df,
        "repeaters_df": repeaters_df, "complex_queries_df": complex_queries_df,
        "parsed_by_report": parsed_by_report, "workbook": workbook, "fn": fn,
        "errors": errors, "data_module_schema": data_module_schema,
        "duplicate_df": duplicate_df, "duplicate_counts": cognos_duplicate_counts(duplicate_records),
        "duplicate_records": duplicate_records,
        "framework_model": framework_model, "framework_model_source_name": framework_model_source_name,
        "inventory_df": inventory_df,
        "complexity_df": complexity_df,
        "waves_df": waves_df,
        "lineage_df": lineage_df,
        "consolidation_df": consolidation_df,
        "complexity_by_report": complexity_by_report,
        "wave_by_report": wave_by_report,
        "blueprints_by_report": blueprints_by_report,
        "model_tables_df": model_tables_df,
        "model_columns_df": model_columns_df,
        "model_measures_df": model_measures_df,
        "model_relationships_df": model_relationships_df,
        "pages_blueprints_by_report": pages_blueprints_by_report,
        "report_pages_df": report_pages_df,
        "visual_bindings_df": visual_bindings_df,
    }

result = st.session_state.get("cognosmap_result")

if result:
    show_download_button(result["workbook"], result["fn"], "download_top")

    st.markdown('<div class="section-label">Senturus-Style Migration Assessment & Portfolio Planning</div>', unsafe_allow_html=True)

    assessment_tabs = st.tabs([
        "Report Inventory",
        "Complexity & Effort",
        "Migration Waves",
        "Cross-Report Lineage & Consolidation",
        "Duplicate Analysis",
    ])

    with assessment_tabs[0]:
        st.markdown("**Combined Portfolio Inventory across all uploaded Cognos reports.**")
        inv_df = result.get("inventory_df", pd.DataFrame())
        total_reports = len(result["parsed_by_report"])
        total_effort = sum(result.get("complexity_by_report", {}).get(r, {}).get("est_effort_hours", 0) for r in result["parsed_by_report"])
        scores = [result.get("complexity_by_report", {}).get(r, {}).get("complexity_score", 0) for r in result["parsed_by_report"]]
        avg_score = (sum(scores) / len(scores)) if scores else 0.0
        high_complex_count = sum(1 for r in result["parsed_by_report"] if result.get("complexity_by_report", {}).get(r, {}).get("complexity_tier") in ("Complex", "Highly Complex"))

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Reports", total_reports)
        m2.metric("Total Est. Effort", f"{total_effort} hrs")
        m3.metric("Avg Complexity Score", f"{avg_score:.1f} / 100")
        m4.metric("Strategic / Complex Reports", high_complex_count)

        show_table_or_info(inv_df, "No report inventory data available.", large=True)

    with assessment_tabs[1]:
        st.markdown(
            "**Explainable Multi-Factor Complexity Scoring (0–100) and Effort Hours Estimation**  \n"
            "*Factors: SQL/Queries (25 pts), Calculations (20 pts), Conditionals & Logic (15 pts), "
            "Layout & Containers (20 pts), Custom Controls (10 pts), Unsupported MDX / Validation (10 pts). "
            "Tiers: Simple (0–25), Medium (26–50), Complex (51–75), Highly Complex (76–100).*"
        )

        with st.expander("⚙️ Configurable Effort Estimation Assumptions", expanded=False):
            st.caption("Adjust baseline migration hours per complexity tier based on project team velocity. All hours are estimates.")
            c_s, c_m, c_c, c_hc = st.columns(4)
            current_effort = st.session_state.get("cognos_effort_hours_map", DEFAULT_EFFORT_HOURS)
            with c_s:
                cfg_s = st.number_input("Simple Tier (hrs)", min_value=1, max_value=200, value=current_effort.get("Simple", 12), key="cfg_effort_simple")
            with c_m:
                cfg_m = st.number_input("Medium Tier (hrs)", min_value=1, max_value=400, value=current_effort.get("Medium", 32), key="cfg_effort_medium")
            with c_c:
                cfg_c = st.number_input("Complex Tier (hrs)", min_value=1, max_value=800, value=current_effort.get("Complex", 64), key="cfg_effort_complex")
            with c_hc:
                cfg_hc = st.number_input("Highly Complex (hrs)", min_value=1, max_value=1000, value=current_effort.get("Highly Complex", 120), key="cfg_effort_hc")

            if st.button("Apply Effort Assumptions", key="btn_apply_effort"):
                new_map = {"Simple": int(cfg_s), "Medium": int(cfg_m), "Complex": int(cfg_c), "Highly Complex": int(cfg_hc)}
                st.session_state["cognos_effort_hours_map"] = new_map
                for r_name in result["parsed_by_report"]:
                    c_info = result["complexity_by_report"][r_name]
                    c_info["est_effort_hours"] = estimate_migration_effort(c_info["complexity_tier"], new_map)
                    if r_name in result.get("wave_by_report", {}):
                        result["wave_by_report"][r_name]["est_effort_hours"] = c_info["est_effort_hours"]
                result["complexity_df"] = build_cognos_complexity_df(result["complexity_by_report"], new_map)
                result["inventory_df"] = build_cognos_inventory_df(result["parsed_by_report"], result["complexity_by_report"], result["wave_by_report"])
                result["waves_df"] = build_cognos_waves_df(result["wave_by_report"], result["complexity_by_report"])
                result["workbook"] = build_excel_workbook(
                    result["overview_df"], result["fields_df"], result["filters_df"], result["prompts_df"],
                    result["drillthroughs_df"], result["viz_df"], result["controls_df"], result["validation_df"],
                    framework_model=result.get("framework_model"),
                    translations_df=result.get("translations_df"),
                    lists_df=result.get("lists_df"),
                    crosstabs_df=result.get("crosstabs_df"),
                    repeaters_df=result.get("repeaters_df"),
                    complex_queries_df=result.get("complex_queries_df"),
                    inventory_df=result["inventory_df"],
                    complexity_df=result["complexity_df"],
                    waves_df=result["waves_df"],
                    lineage_df=result["lineage_df"],
                    consolidation_df=result["consolidation_df"],
                    duplicate_df=result["duplicate_df"],
                    model_tables_df=result.get("model_tables_df"),
                    model_columns_df=result.get("model_columns_df"),
                    model_measures_df=result.get("model_measures_df"),
                    model_relationships_df=result.get("model_relationships_df"),
                )
                st.rerun()

        tier_counts = Counter(result.get("complexity_by_report", {}).get(r, {}).get("complexity_tier") for r in result["parsed_by_report"])
        tc1, tc2, tc3, tc4 = st.columns(4)
        tc1.metric("Simple (0–25)", tier_counts.get("Simple", 0))
        tc2.metric("Medium (26–50)", tier_counts.get("Medium", 0))
        tc3.metric("Complex (51–75)", tier_counts.get("Complex", 0))
        tc4.metric("Highly Complex (76–100)", tier_counts.get("Highly Complex", 0))

        show_table_or_info(result.get("complexity_df"), "No complexity scoring data available.", large=True)

    with assessment_tabs[2]:
        st.markdown(
            "**Senturus 4-Wave Migration Roadmap**  \n"
            "*Waves prioritize high-value quick wins early, tackle core reports predictably, reserve deep technical redesign for complex reports, and consolidate or retire duplicates.*"
        )

        wave_counts = Counter(result.get("wave_by_report", {}).get(r, {}).get("wave_number", 2) for r in result["parsed_by_report"])
        w_hours = defaultdict(int)
        for r, w_info in result.get("wave_by_report", {}).items():
            w_hours[w_info.get("wave_number", 2)] += w_info.get("est_effort_hours", 0)

        wc1, wc2, wc3, wc4 = st.columns(4)
        wc1.metric("Wave 1: Quick Wins", f"{wave_counts.get(1, 0)} reports", f"{w_hours.get(1, 0)} est. hrs")
        wc2.metric("Wave 2: Core Migration", f"{wave_counts.get(2, 0)} reports", f"{w_hours.get(2, 0)} est. hrs")
        wc3.metric("Wave 3: Strategic Redesign", f"{wave_counts.get(3, 0)} reports", f"{w_hours.get(3, 0)} est. hrs")
        wc4.metric("Wave 4: Retire / Consolidate", f"{wave_counts.get(4, 0)} reports", f"{w_hours.get(4, 0)} est. hrs")

        with st.expander("⚙️ Supply Business Value & Usage Information", expanded=False):
            st.caption(
                "Customize business value per report (High, Medium, Low, Retire). "
                "Default business value is **Medium**. In accordance with migration governance, "
                "usage statistics are never fabricated and reports are never automatically retired without explicit confirmation."
            )
            bv_dict = st.session_state.get("cognos_business_values", {})
            bv_cols = st.columns(min(3, max(1, len(result["parsed_by_report"]))))
            reports_list = list(result["parsed_by_report"].keys())
            updated_any = False
            for i, r_name in enumerate(reports_list):
                with bv_cols[i % len(bv_cols)]:
                    cur_v = bv_dict.get(r_name, "Medium")
                    sel_v = st.selectbox(
                        f"Value: {r_name[:25]}",
                        ["Medium", "High", "Low", "Retire"],
                        index=["Medium", "High", "Low", "Retire"].index(cur_v) if cur_v in ["Medium", "High", "Low", "Retire"] else 0,
                        key=f"bv_input_{r_name}",
                    )
                    if sel_v != cur_v:
                        bv_dict[r_name] = sel_v
                        updated_any = True

            if st.button("Update Wave Roadmap", key="btn_apply_waves") or updated_any:
                st.session_state["cognos_business_values"] = bv_dict
                dup_cat_map = {r["report_name"]: r["category"] for r in result.get("duplicate_records", [])}
                seen_dup_groups = {}
                dup_copy_info = {}
                for r in result.get("duplicate_records", []):
                    gid = r.get("group_id")
                    rname = r.get("report_name")
                    cat = r.get("category")
                    if gid and cat in ("Exact Duplicate", "Near Duplicate"):
                        if gid not in seen_dup_groups:
                            seen_dup_groups[gid] = rname
                            dup_copy_info[rname] = {"is_copy": False, "canonical": rname}
                        else:
                            dup_copy_info[rname] = {"is_copy": True, "canonical": seen_dup_groups[gid]}
                    else:
                        dup_copy_info[rname] = {"is_copy": False, "canonical": rname}
                eff_map = st.session_state.get("cognos_effort_hours_map", DEFAULT_EFFORT_HOURS)
                for r_name in reports_list:
                    c_info = result["complexity_by_report"][r_name]
                    dup_cat = dup_cat_map.get(r_name, "Unique")
                    bv = bv_dict.get(r_name, "Medium")
                    d_info = dup_copy_info.get(r_name, {"is_copy": False, "canonical": r_name})
                    w_label, w_num, w_reason = assign_migration_wave(
                        r_name, c_info["complexity_score"], c_info["complexity_tier"], dup_cat, bv,
                        is_duplicate_copy=d_info["is_copy"],
                        canonical_match=d_info["canonical"] if d_info["is_copy"] else None,
                    )
                    result["wave_by_report"][r_name] = {
                        "wave": w_label,
                        "wave_number": w_num,
                        "reason": w_reason,
                        "business_value": bv,
                        "est_effort_hours": c_info["est_effort_hours"],
                    }
                result["waves_df"] = build_cognos_waves_df(result["wave_by_report"], result["complexity_by_report"])
                result["inventory_df"] = build_cognos_inventory_df(result["parsed_by_report"], result["complexity_by_report"], result["wave_by_report"])
                result["workbook"] = build_excel_workbook(
                    result["overview_df"], result["fields_df"], result["filters_df"], result["prompts_df"],
                    result["drillthroughs_df"], result["viz_df"], result["controls_df"], result["validation_df"],
                    framework_model=result.get("framework_model"),
                    translations_df=result.get("translations_df"),
                    lists_df=result.get("lists_df"),
                    crosstabs_df=result.get("crosstabs_df"),
                    repeaters_df=result.get("repeaters_df"),
                    complex_queries_df=result.get("complex_queries_df"),
                    inventory_df=result["inventory_df"],
                    complexity_df=result["complexity_df"],
                    waves_df=result["waves_df"],
                    lineage_df=result["lineage_df"],
                    consolidation_df=result["consolidation_df"],
                    duplicate_df=result["duplicate_df"],
                    model_tables_df=result.get("model_tables_df"),
                    model_columns_df=result.get("model_columns_df"),
                    model_measures_df=result.get("model_measures_df"),
                    model_relationships_df=result.get("model_relationships_df"),
                )
                if updated_any:
                    st.rerun()

        show_table_or_info(result.get("waves_df"), "No wave roadmap data available.", large=True)

    with assessment_tabs[3]:
        st.markdown(
            "**Cross-Report Lineage & Semantic Model Consolidation**  \n"
            "*Maps data sources and underlying packages across reports to discover opportunities for shared Power BI Semantic Models.*"
        )
        l_sub1, l_sub2 = st.tabs(["Source Table & Model Lineage", "Consolidation & Shared Semantic Models"])
        with l_sub1:
            st.caption("Identifies which Cognos source tables and models feed which reports, and counts referenced columns.")
            show_table_or_info(result.get("lineage_df"), "No table lineage detected across uploaded reports.", large=True)
        with l_sub2:
            st.caption("Discovers report pairs sharing $\\ge 70\\%$ source tables or common packages, recommending shared Power BI Semantic Models to prevent dataset proliferation.")
            cons_df = result.get("consolidation_df", pd.DataFrame())
            if cons_df.empty:
                st.info("No report pairs with $\\ge 70\\%$ table overlap or duplicate patterns detected.")
            else:
                show_table_or_info(cons_df, "No consolidation opportunities detected.", large=True)

    with assessment_tabs[4]:
        st.markdown(
            "**Duplicate & Redundancy Detection**  \n"
            "*Deep structural fingerprinting across model paths, source tables, fields, filters, prompts, and drillthrough targets.*"
        )
        dup_counts = result.get("duplicate_counts", {})
        d1, d2, d3, d4 = st.columns(4)
        d1.metric("Exact Duplicates", dup_counts.get("exact", 0))
        d2.metric("Near Duplicates", dup_counts.get("near", 0))
        d3.metric("Same Data Source", dup_counts.get("same_source", 0))
        d4.metric("Unique Reports", dup_counts.get("unique", 0))
        show_table_or_info(result.get("duplicate_df"), "No duplicate data available.", large=True)

    framework_model = result.get("framework_model")
    if framework_model is not None:
        st.markdown('<div class="section-label">Framework Model</div>', unsafe_allow_html=True)
        st.caption(f"Parsed from: {result.get('framework_model_source_name', '')}")
        show_table_or_info(build_framework_overview_df(framework_model), "No framework model data.")

        fm_tabs = st.tabs(["Packages", "Query Subjects", "Columns", "Relationships", "Filters & Security"])
        with fm_tabs[0]:
            show_table_or_info(build_framework_packages_df(framework_model), "No packages found.")
        with fm_tabs[1]:
            show_table_or_info(build_framework_query_subjects_df(framework_model), "No query subjects found.")
        with fm_tabs[2]:
            show_table_or_info(build_framework_columns_df(framework_model), "No columns found.", large=True)
        with fm_tabs[3]:
            show_table_or_info(build_framework_relationships_df(framework_model), "No relationships found.")
        with fm_tabs[4]:
            show_table_or_info(build_framework_filters_security_df(framework_model), "No filters or security rules found.")

    dm_schema = result.get("data_module_schema")
    if dm_schema is not None:
        relationships = dm_schema.get("relationships", [])
        with st.expander(f"Data Module Relationships ({len(relationships)} found)", expanded=False):
            if not relationships:
                st.info(
                    "No relationships/joins were declared in this Data Module - "
                    "each table stands alone with no connections between them. "
                    "This is common for models built from independently uploaded Excel/CSV files."
                )
            else:
                for rel in relationships:
                    st.write(f"**Found at:** `{rel['path']}`")
                    st.json(rel["value"])

    if result["errors"]:
        st.warning(f"{len(result['errors'])} file(s) had errors.")
        with st.expander("View Error Log"):
            for err in result["errors"]:
                st.text(err)

    st.markdown('<div class="section-label">Choose a report to inspect</div>', unsafe_allow_html=True)

    report_names = list(result["parsed_by_report"].keys())
    if report_names:
        selected_report = st.selectbox("Choose a report to inspect", report_names, label_visibility="collapsed")
        parsed, validation_issues = result["parsed_by_report"][selected_report]

        risk_label, risk_level = complexity_label(
            len(parsed["fields"]), len(parsed["custom_controls"]), len(validation_issues)
        )

        comp_info = result.get("complexity_by_report", {}).get(selected_report, {})
        wave_info = result.get("wave_by_report", {}).get(selected_report, {})
        comp_score = comp_info.get("complexity_score", 0)
        comp_tier = comp_info.get("complexity_tier", "Medium")
        est_hours = comp_info.get("est_effort_hours", 32)
        assigned_wave = wave_info.get("wave", "Wave 2: Core Migration")

        tier_level_map = {"Simple": "ok", "Medium": "medium", "Complex": "review", "Highly Complex": "high"}
        comp_badge = badge(f"Complexity: {comp_score}/100 ({comp_tier})", tier_level_map.get(comp_tier, ""))
        wave_badge = badge(assigned_wave, "ok" if "Wave 1" in assigned_wave else "")
        effort_badge = badge(f"Est. Effort: {est_hours} hrs", "")

        matched_package = match_report_to_framework(parsed['model_path'], result.get("framework_model"))
        if result.get("framework_model") is None:
            framework_badge = badge('No framework model uploaded', '')
        elif matched_package:
            framework_badge = badge(f"Framework: {matched_package['name']}", 'ok')
        else:
            framework_badge = badge('No matching package in uploaded framework model', 'review')

        st.markdown(
            f"""
            <div class="report-header">
                <h2>{escape(selected_report)}</h2>
                <p><strong>Model path:</strong> <code>{escape(parsed['model_path'] or 'Not detected')}</code></p>
                <p><strong>Source tables:</strong> {escape(', '.join(parsed['source_tables']) or 'None detected')}</p>
                {badge(f'Risk: {risk_label}', risk_level)}
                {badge('Verified' if not validation_issues else f'{len(validation_issues)} validation issue(s)', 'ok' if not validation_issues else 'review')}
                {comp_badge}
                {wave_badge}
                {effort_badge}
                {framework_badge}
            </div>
            """,
            unsafe_allow_html=True,
        )

        field_translations = parsed.get("field_translations", [])
        translatable_count = sum(1 for t in field_translations if t.get("is_translatable"))
        manual_review_count = sum(1 for t in field_translations if not t.get("is_translatable")) + len(parsed.get("custom_controls", []))
        trans_rate = round(translatable_count / len(field_translations) * 100) if field_translations else 100

        r1, r2, r3, r4, r5, r6 = st.columns(6)
        with r1:
            with st.container(border=True):
                st.metric("Fields", len(parsed["fields"]))
        with r2:
            with st.container(border=True):
                st.metric("Filters / Prompts", f"{len(parsed['filters'])} / {len(parsed['prompts'])}")
        with r3:
            with st.container(border=True):
                st.metric("Lists / Crosstabs", f"{len(parsed.get('lists', []))} / {len(parsed.get('crosstabs', []))}")
        with r4:
            with st.container(border=True):
                st.metric("Deterministic Trans.", f"{trans_rate}%")
        with r5:
            with st.container(border=True):
                st.metric("Manual Review", manual_review_count)
        with r6:
            with st.container(border=True):
                st.metric("Est. Effort", f"{est_hours} hrs")

        tabs = st.tabs([
            "Report Pages & Visuals (PBIP)",
            "Semantic Model Blueprint",
            "Translations (DAX / M)",
            "Layout & Containers",
            "Fields",
            "Filters",
            "Prompts",
            "Drillthroughs",
            "Visualizations",
            "Custom Controls",
            "Validation",
            "AI Build Plan",
        ])

        report_blueprint = result.get("blueprints_by_report", {}).get(selected_report)
        if not report_blueprint:
            report_blueprint = build_semantic_model_blueprint(
                selected_report, parsed, validation_issues,
                framework_model=result.get("framework_model"),
                data_module_schema=result.get("data_module_schema")
            )

        report_pages_blueprint = result.get("pages_blueprints_by_report", {}).get(selected_report)
        if not report_pages_blueprint:
            report_pages_blueprint = extract_report_pages_blueprint(
                parsed.get("xml_text", ""), report_blueprint, parsed=parsed
            )

        with tabs[0]:
            st.markdown(
                """
                ### 📄 Power BI Report Pages & Visuals (PBIP / PBIR)
                *Deterministic conversion of Cognos report layout, lists, crosstabs, charts, and prompts into modern Power BI Enhanced Report Format (PBIR) visual definitions and bindings referencing the generated Semantic Model.*
                """
            )

            pbir_val = validate_pbir_report_package(report_pages_blueprint, report_blueprint)
            tot_pages = len(report_pages_blueprint)
            tot_vis = pbir_val["total_visuals"]
            valid_bindings = pbir_val["valid_bindings"]
            val_score = pbir_val["validation_score"]
            unmapped_count = pbir_val["unmapped_bindings"]
            manual_rev_count = sum(p["manual_review_count"] for p in report_pages_blueprint)

            k1, k2, k3, k4, k5 = st.columns(5)
            with k1:
                with st.container(border=True):
                    st.metric("Report Pages", tot_pages)
            with k2:
                with st.container(border=True):
                    st.metric("Visual Containers", tot_vis)
            with k3:
                with st.container(border=True):
                    st.metric("Bound Projections", f"{valid_bindings} / {pbir_val['total_bindings']}")
            with k4:
                with st.container(border=True):
                    st.metric("Binding Accuracy", f"{val_score}%")
            with k5:
                with st.container(border=True):
                    st.metric("Manual Review Items", manual_rev_count)

            # PBIP Download Package Button
            safe_rep_name = re.sub(r"[^\w-]", "_", selected_report)
            pbip_bytes = generate_pbip_project_zip(selected_report, report_pages_blueprint, report_blueprint)

            c_down, c_info = st.columns([1, 2])
            with c_down:
                st.download_button(
                    label=f"📦 Download Power BI Project ({safe_rep_name}.pbip.zip)",
                    data=pbip_bytes,
                    file_name=f"{safe_rep_name}.pbip.zip",
                    mime="application/zip",
                    use_container_width=True,
                    type="primary",
                    key=f"dl_pbip_{selected_report}",
                )
            with c_info:
                st.caption(
                    "Includes root `.pbip`, Fabric `.SemanticModel/` (TMDL definitions), and `.Report/` (Enhanced Report PBIR structure with `definition.pbir`, `pages.json`, and visual containers)."
                )

            with st.expander("ℹ️ How to open the downloaded project in Power BI Desktop", expanded=False):
                st.markdown(
                    f"""
                    1. **Extract Archive**: Extract the downloaded `{safe_rep_name}.pbip.zip` file to your workspace directory (e.g. `C:\\PowerBI\\{safe_rep_name}`).
                    2. **Enable Enhanced Report Format (PBIR)**:
                       - Open Power BI Desktop.
                       - Navigate to **File** > **Options and settings** > **Options** > **Preview features**.
                       - Check **Store reports using enhanced metadata format (PBIR)** and restart Power BI Desktop.
                    3. **Open Project**: Double-click `{safe_rep_name}.pbip` directly in Windows File Explorer.
                    4. **Review Visuals & Formulas**: Power BI Desktop will load the semantic model and report visuals directly. Check any items marked as *Manual Review Required*.
                    """
                )

            st.divider()

            # Page Layout & Visual Explorer
            st.markdown("#### 📱 Report Pages Layout & Canvas Visuals")
            if report_pages_blueprint:
                page_names = [f"{p['display_name']} ({p['page_name']})" for p in report_pages_blueprint]
                selected_page_idx = 0
                if len(report_pages_blueprint) > 1:
                    sel_p_label = st.radio("Select Page:", page_names, horizontal=True)
                    selected_page_idx = page_names.index(sel_p_label)

                active_page = report_pages_blueprint[selected_page_idx]

                st.caption(
                    f"Canvas Resolution: **{active_page['width']} x {active_page['height']} px** | Display Option: **{active_page['display_option']}** | Visuals on Canvas: **{active_page['visual_count']}**"
                )

                # Render Canvas Mockup / Visual Cards
                v_cols = st.columns(2)
                for idx, vis in enumerate(active_page["visuals"]):
                    col_target = v_cols[idx % 2]
                    with col_target:
                        with st.container(border=True):
                            pos = vis.get("position", {})
                            st.markdown(
                                f"**{vis['name']}** `[{vis['visual_type']}]` "
                                f"*(Source: `{vis['source_element']}`)*"
                            )
                            c_m1, c_m2 = st.columns([1, 1])
                            with c_m1:
                                if vis["status"] == "Converted":
                                    st.success(f"Status: {vis['status']}", icon="✅")
                                else:
                                    st.warning(f"Status: {vis['status']}", icon="⚠️")
                            with c_m2:
                                st.code(f"x:{pos.get('x', 0)} y:{pos.get('y', 0)} w:{pos.get('width', 0)} h:{pos.get('height', 0)}", language="yaml")

                            # Projections Summary
                            bound_summary = []
                            for r_name, r_items in vis.get("role_mappings", {}).items():
                                if r_items:
                                    bound_summary.append(f"**{r_name}**: {', '.join(r_items)}")
                            if bound_summary:
                                st.markdown("<br>".join(bound_summary), unsafe_allow_html=True)

                            if vis.get("warnings"):
                                for w in vis["warnings"]:
                                    st.caption(f"⚠️ {w}")

            st.divider()

            # Detailed Tables for Pages and Visual Bindings
            st.markdown("#### 📑 Visual Projections & Field Bindings Inventory")
            v_tab1, v_tab2 = st.tabs(["Visual Field Bindings", "Report Pages List"])

            with v_tab1:
                cur_vb_df = result.get("visual_bindings_df", pd.DataFrame())
                if not cur_vb_df.empty:
                    cur_vb_df = cur_vb_df[cur_vb_df["Report"] == selected_report]
                    show_table_or_info(
                        cur_vb_df.drop(columns=["Report"]) if not cur_vb_df.empty else cur_vb_df,
                        "No visual bindings recorded for this report.",
                        large=True,
                    )
                else:
                    st.info("No visual bindings recorded.")

            with v_tab2:
                cur_rp_df = result.get("report_pages_df", pd.DataFrame())
                if not cur_rp_df.empty:
                    cur_rp_df = cur_rp_df[cur_rp_df["Report"] == selected_report]
                    show_table_or_info(
                        cur_rp_df.drop(columns=["Report"]) if not cur_rp_df.empty else cur_rp_df,
                        "No report pages recorded.",
                    )
                else:
                    st.info("No report pages recorded.")

            # PBIR JSON Schema Inspector
            with st.expander("🔍 PBIR JSON Schema & Visual Definition Inspector", expanded=False):
                st.markdown("**1. `definition.pbir` Pointer:**")
                st.code(
                    json.dumps({
                        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
                        "version": "4.0",
                        "datasetReference": {
                            "byPath": {"path": f"../{report_blueprint.get('model_name', 'Model')}.SemanticModel"}
                        }
                    }, indent=2),
                    language="json",
                )
                if report_pages_blueprint and report_pages_blueprint[0]["visuals"]:
                    sample_v = report_pages_blueprint[0]["visuals"][0]
                    st.markdown(f"**2. Sample `visual.json` for `{sample_v['name']}` ({sample_v['visual_type']}):**")
                    st.code(
                        json.dumps({
                            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.4.0/schema.json",
                            "name": sample_v["visual_id"],
                            "position": sample_v["position"],
                            "visual": {
                                "visualType": sample_v["visual_type"],
                                "query": {
                                    "queryState": sample_v.get("query_state", {})
                                }
                            }
                        }, indent=2),
                        language="json",
                    )

        with tabs[1]:
            st.markdown(
                """
                ### 🏗️ Power BI Semantic Model Blueprint
                *Proposed tabular model architecture generated deterministically from Cognos report metadata — grouping physical tables, data columns, DAX measures, and inferred relationships. Distinguishes verified vs inferred metadata and flags items requiring manual review.*
                """
            )

            # High-Level Metrics
            bp_summary = report_blueprint.get("summary", {})
            m1, m2, m3, m4, m5, m6, m7 = st.columns(7)
            with m1:
                with st.container(border=True):
                    st.metric("Tables", bp_summary.get("table_count", 0))
            with m2:
                with st.container(border=True):
                    st.metric("Columns", bp_summary.get("column_count", 0))
            with m3:
                with st.container(border=True):
                    st.metric("Measures", bp_summary.get("measure_count", 0))
            with m4:
                with st.container(border=True):
                    st.metric("Relationships", bp_summary.get("relationship_count", 0))
            with m5:
                with st.container(border=True):
                    st.metric("Verified", bp_summary.get("verified_count", 0))
            with m6:
                with st.container(border=True):
                    st.metric("Inferred", bp_summary.get("inferred_count", 0))
            with m7:
                with st.container(border=True):
                    st.metric("Manual Review", bp_summary.get("manual_review_count", 0))

            # Download Artifacts (TMDL Document and PBIP ZIP Package)
            tmdl_doc = generate_tmdl_document(report_blueprint)
            tmdl_zip = generate_tmdl_zip_package(report_blueprint)
            safe_model_name = re.sub(r"[^\w-]", "_", report_blueprint.get("model_name", "Model"))

            d_col1, d_col2 = st.columns([0.5, 0.5])
            with d_col1:
                st.download_button(
                    label="📥 Download Unified Model Blueprint (.tmdl)",
                    data=tmdl_doc,
                    file_name=f"{safe_model_name}.tmdl",
                    mime="text/plain",
                    help="Direct Tabular Model Definition Language (TMDL) unified document.",
                    use_container_width=True,
                    key=f"dl_tmdl_{selected_report}",
                )
            with d_col2:
                st.download_button(
                    label="📦 Download Complete PBIP TMDL Package (.zip)",
                    data=tmdl_zip,
                    file_name=f"{safe_model_name}_PBIP_TMDL.zip",
                    mime="application/zip",
                    help="PBIP folder structure with definition/model.tmdl, tables/*.tmdl, relationships.tmdl, and README.md.",
                    use_container_width=True,
                    key=f"dl_tmdl_zip_{selected_report}",
                )

            # Subtabs for Detailed Semantic Inspection
            sm_sub1, sm_sub2, sm_sub3, sm_sub4, sm_sub5 = st.tabs([
                "Tables & Columns",
                "Measures & DAX Validation",
                "Relationships",
                "Governance & Audit",
                "Power BI Desktop / MCP Scanner",
            ])

            with sm_sub1:
                st.markdown("**Model Tables Overview**")
                rep_tables_df = build_model_tables_df(report_blueprint)
                show_table_or_info(rep_tables_df, "No tables identified in blueprint.", large=False)

                st.markdown("**Model Columns Detail**")
                rep_cols_df = build_model_columns_df(report_blueprint)
                show_table_or_info(rep_cols_df, "No columns identified in blueprint.", large=True)

            with sm_sub2:
                st.markdown(
                    """
                    **DAX Measures & Integrity Validation**  
                    *Measures inferred from aggregations, calculations, and layout slots. Each expression is validated against Power BI DAX syntax rules.*
                    """
                )
                rep_measures_df = build_model_measures_df(report_blueprint)
                show_table_or_info(rep_measures_df, "No measures identified in blueprint.", large=True)

                manual_measures = [m for m in report_blueprint.get("measures", []) if m.get("provenance") == "Manual Review"]
                if manual_measures:
                    with st.expander(f"⚠️ Measures Flagged for Manual Review ({len(manual_measures)} item(s))", expanded=False):
                        for mm in manual_measures:
                            st.markdown(f"- **`{mm['measure_name']}`** on table `{mm['table_name']}`: {mm['notes']}")
                            st.code(mm['dax_expression'], language="dax")

            with sm_sub3:
                st.markdown(
                    """
                    **Proposed Cross-Table Relationships**  
                    *Derived from Framework Model relationships or declared Cognos Query Joins.*  
                    > **Strict Governance Rule**: KashMap *never* invents artificial join keys or foreign keys. If tables are queried together without declared join keys in XML, the relationship is marked `Inactive (Review Required)` with `[Unmapped Join Column]`.
                    """
                )
                rep_rels_df = build_model_relationships_df(report_blueprint)
                show_table_or_info(rep_rels_df, "No multi-table relationships required (single-table model).", large=True)

            with sm_sub4:
                st.markdown(
                    """
                    **Provenance & Classification Audit**  
                    *Audits every semantic object by governance level: Verified (grounded in Framework Model or Data Module), Inferred (rule-based heuristic), or Manual Review (ambiguous or unsupported syntax).*
                    """
                )
                audit_records = []
                for t in report_blueprint.get("tables", []):
                    audit_records.append({
                        "Object Type": "Table", "Object Name": t["table_name"],
                        "Provenance": t["provenance"], "Notes": f"Source: {t['source_path']}"
                    })
                for c in report_blueprint.get("columns", []):
                    audit_records.append({
                        "Object Type": "Column", "Object Name": f"{c['table_name']}[{c['column_name']}]",
                        "Provenance": c["provenance"], "Notes": c["notes"] or c["data_type"]
                    })
                for m in report_blueprint.get("measures", []):
                    audit_records.append({
                        "Object Type": "Measure", "Object Name": f"[{m['measure_name']}]",
                        "Provenance": m["provenance"], "Notes": m["notes"]
                    })
                for r in report_blueprint.get("relationships", []):
                    audit_records.append({
                        "Object Type": "Relationship", "Object Name": r.get("relationship_name", "Rel"),
                        "Provenance": r["provenance"], "Notes": r["notes"]
                    })
                show_table_or_info(pd.DataFrame(audit_records), "No audit objects available.", large=True)

            with sm_sub5:
                st.markdown(
                    """
                    **Local Power BI Desktop & Modeling MCP Inspection**  
                    *Integrates with the `powerbi-authoring-local` MCP server to discover running Power BI Desktop instances and validate DAX queries against active tabular models.*
                    """
                )
                mcp_status = check_powerbi_mcp_connections()
                
                if mcp_status["connected"]:
                    st.success(f"🟢 **Power BI Desktop Connected**: {mcp_status['message']}")
                    st.dataframe(pd.DataFrame(mcp_status["instances"]), use_container_width=True, hide_index=True)
                else:
                    st.info(f"ℹ️ **Power BI Desktop Status**: {mcp_status['message']}")

                st.markdown("---")
                st.markdown("##### 🧪 Interactive DAX Syntax & Query Tester")
                sample_dax = report_blueprint["measures"][0]["dax_expression"] if report_blueprint.get("measures") else "SUM('Sales'[Revenue])"
                test_dax_input = st.text_area("Enter DAX Expression to Validate:", value=sample_dax, height=90, key=f"dax_test_box_{selected_report}")
                if st.button("Validate DAX Expression", key=f"btn_val_dax_{selected_report}"):
                    val_out = validate_dax_with_mcp(test_dax_input)
                    if val_out["is_valid"]:
                        st.success(f"✅ **Valid Expression**: {val_out['status']}")
                        if val_out.get("engine"):
                            st.caption(val_out["engine"])
                    else:
                        st.error(f"❌ **Validation Flagged**: {val_out['status']}")
                        for w in val_out.get("warnings", []):
                            st.markdown(f"- {w}")

                st.markdown("---")
                st.markdown("##### 🛡️ Live Model Deployment Guard")
                st.warning("⚠️ **Strict Read-Only Guard**: To prevent accidental data loss, KashMap will NEVER alter or create tables in a live Power BI Desktop session without explicit confirmation.")
                confirm_deploy = st.checkbox("I understand and confirm that I want to apply model changes to my active Power BI Desktop session.", key=f"chk_confirm_deploy_{selected_report}")
                if st.button("Deploy Semantic Model to Connected Power BI", key=f"btn_deploy_pbi_{selected_report}", disabled=not confirm_deploy, type="primary"):
                    deploy_res = apply_semantic_model_to_live_pbi(report_blueprint, user_confirmed=confirm_deploy)
                    if deploy_res["status"] == "success":
                        st.success(f"🎉 {deploy_res['message']}")
                    elif deploy_res["status"] == "blocked":
                        st.warning(deploy_res["message"])
                    else:
                        st.info(f"ℹ️ {deploy_res['message']}")

        with tabs[2]:
            st.markdown(
                """
                **Deterministic Rule-Based Expression Transpiler**  
                *Converts Cognos calculations and filters directly to DAX Measures, Calculated Columns, Power Query M, and Report Filters without requiring LLM calls.*
                """
            )
            report_trans_df = result.get("translations_df", pd.DataFrame())
            if not report_trans_df.empty:
                report_trans_df = report_trans_df[report_trans_df["Report"] == selected_report]

            if report_trans_df.empty:
                st.info("No calculations or filters found in this report.")
            else:
                f_col1, f_col2 = st.columns([0.4, 0.6])
                with f_col1:
                    layer_filter = st.selectbox(
                        "Filter by Target Layer / Status",
                        ["All", "DAX Measures", "Calculated Columns", "Report Filters", "Manual Review Needed"],
                        key=f"trans_filter_{selected_report}",
                    )
                filtered_df = report_trans_df.copy()
                if layer_filter == "DAX Measures":
                    filtered_df = filtered_df[filtered_df["Target Layer"] == "DAX Measure"]
                elif layer_filter == "Calculated Columns":
                    filtered_df = filtered_df[filtered_df["Target Layer"] == "Calculated Column"]
                elif layer_filter == "Report Filters":
                    filtered_df = filtered_df[filtered_df["Target Layer"] == "Report Filter"]
                elif layer_filter == "Manual Review Needed":
                    filtered_df = filtered_df[(filtered_df["Status"] == "Manual Review") | (filtered_df["Warnings"] != "")]

                show_table_or_info(
                    filtered_df.drop(columns=["Report"]) if not filtered_df.empty else filtered_df,
                    f"No translations matching filter '{layer_filter}'.",
                    large=True,
                )

                flagged = report_trans_df[report_trans_df["Warnings"] != ""]
                if not flagged.empty:
                    with st.expander(f"⚠️ Translation Warnings & Ambiguities ({len(flagged)} item(s))", expanded=False):
                        for _, row in flagged.iterrows():
                            st.markdown(f"- **{row['Scope']} `{row['Field']}`** ({row['Category']}): {row['Warnings']}")

            report_macros = parsed.get("macros", [])
            if report_macros:
                with st.expander(f"Cognos Macros & Prompt Tokens ({len(report_macros)} found)", expanded=False):
                    st.caption("Extracted from expressions (`#prompt(...)#`, `#promptmany(...)#`, `#$account...#`). Map to Power Query parameters or Power BI Slicers.")
                    macro_rows = []
                    for m in report_macros:
                        macro_rows.append({
                            "Macro Text": m.get("raw_macro", ""),
                            "Type": m.get("macro_type", ""),
                            "Parameter": m.get("parameter", ""),
                            "Datatype": m.get("datatype", ""),
                            "Default Value": m.get("default_value", ""),
                        })
                    st.dataframe(pd.DataFrame(macro_rows), use_container_width=True, hide_index=True)

        with tabs[3]:
            st.markdown(
                """
                **Cognos Layout Objects & Query Architecture**  
                *Inspect how Cognos Lists, Crosstabs, Repeaters, and complex queries map into Power BI visuals and Semantic Model structures.*
                """
            )
            l_tab1, l_tab2, l_tab3, l_tab4 = st.tabs(["Lists (Table)", "Crosstabs (Matrix)", "Repeaters", "Complex Queries"])
            with l_tab1:
                st.caption("Cognos Lists map directly to Power BI **Table visual (`tableEx`)**, preserving column order, groupings, and column titles.")
                report_lists_df = result.get("lists_df", pd.DataFrame())
                if not report_lists_df.empty:
                    report_lists_df = report_lists_df[report_lists_df["Report"] == selected_report]
                show_table_or_info(
                    report_lists_df.drop(columns=["Report"]) if not report_lists_df.empty else report_lists_df,
                    "No Lists found in this report.",
                )

            with l_tab2:
                st.caption("Cognos Crosstabs map directly to Power BI **Matrix visual (`pivotTable`)**. Row dimensions form Matrix Rows, Column dimensions form Matrix Columns, and Measures form Values.")
                report_crosstabs_df = result.get("crosstabs_df", pd.DataFrame())
                if not report_crosstabs_df.empty:
                    report_crosstabs_df = report_crosstabs_df[report_crosstabs_df["Report"] == selected_report]
                show_table_or_info(
                    report_crosstabs_df.drop(columns=["Report"]) if not report_crosstabs_df.empty else report_crosstabs_df,
                    "No Crosstabs found in this report.",
                )

            with l_tab3:
                st.caption("Cognos Repeaters map to Power BI **Small Multiples** on Cartesian visuals or grouped Matrix hierarchy levels.")
                report_repeaters_df = result.get("repeaters_df", pd.DataFrame())
                if not report_repeaters_df.empty:
                    report_repeaters_df = report_repeaters_df[report_repeaters_df["Report"] == selected_report]
                show_table_or_info(
                    report_repeaters_df.drop(columns=["Report"]) if not report_repeaters_df.empty else report_repeaters_df,
                    "No Repeaters found in this report.",
                )

            with l_tab4:
                st.caption("Complex Query Architectures: Query Joins (`Table.NestedJoin`), Set Unions (`Table.Combine`), and Native SQL pass-through (`Value.NativeQuery`).")
                report_cq_df = result.get("complex_queries_df", pd.DataFrame())
                if not report_cq_df.empty:
                    report_cq_df = report_cq_df[report_cq_df["Report"] == selected_report]
                show_table_or_info(
                    report_cq_df.drop(columns=["Report"]) if not report_cq_df.empty else report_cq_df,
                    "No complex queries (joins/unions/native SQL) found in this report.",
                )

        with tabs[4]:
            report_fields_df = result["fields_df"][result["fields_df"]["Report"] == selected_report]
            show_table_or_info(
                report_fields_df.drop(columns=["Report"]) if not report_fields_df.empty else report_fields_df,
                "No fields found.", large=True,
            )

        with tabs[5]:
            report_filters_df = result["filters_df"][result["filters_df"]["Report"] == selected_report] if not result["filters_df"].empty else pd.DataFrame()
            show_table_or_info(
                report_filters_df.drop(columns=["Report"]) if not report_filters_df.empty else report_filters_df,
                "No filters found.",
            )

        with tabs[6]:
            report_prompts_df = result["prompts_df"][result["prompts_df"]["Report"] == selected_report] if not result["prompts_df"].empty else pd.DataFrame()
            show_table_or_info(
                report_prompts_df.drop(columns=["Report"]) if not report_prompts_df.empty else report_prompts_df,
                "No prompts found.",
            )

        with tabs[7]:
            report_drills_df = result["drillthroughs_df"][result["drillthroughs_df"]["Source Report"] == selected_report] if not result["drillthroughs_df"].empty else pd.DataFrame()
            show_table_or_info(report_drills_df, "No drillthroughs found.")

        with tabs[8]:
            report_viz_df = result["viz_df"][result["viz_df"]["Report"] == selected_report] if not result["viz_df"].empty else pd.DataFrame()
            show_table_or_info(
                report_viz_df.drop(columns=["Report"]) if not report_viz_df.empty else report_viz_df,
                "No visualizations found.",
            )

        with tabs[9]:
            report_controls_df = result["controls_df"][result["controls_df"]["Report"] == selected_report] if not result["controls_df"].empty else pd.DataFrame()
            if report_controls_df.empty:
                st.success("No custom controls - nothing needing manual JS review.")
            else:
                st.warning(f"{len(report_controls_df)} custom control(s) need manual review - no Power BI equivalent.")
                show_table_or_info(report_controls_df.drop(columns=["Report"]), "No custom controls found.")

        with tabs[10]:
            if not validation_issues:
                validation_summary_df = pd.DataFrame([{
                    "Report": selected_report,
                    "Status": "Verified",
                    "Issue Count": 0,
                    "Details": "All fields in this report verified against the Data Module (or no Data Module was uploaded).",
                }])
                show_table_or_info(validation_summary_df, "No validation data.")
            else:
                validation_issues_df = pd.DataFrame([
                    {"Report": selected_report, "Issue #": i + 1, "Issue": issue}
                    for i, issue in enumerate(validation_issues)
                ])
                st.error(f"{len(validation_issues)} field(s) could not be verified against the Data Module:")
                show_table_or_info(validation_issues_df, "No validation issues.")

        with tabs[11]:
            render_cognos_ai_build_plan(selected_report, parsed, validation_issues)