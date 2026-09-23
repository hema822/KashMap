import re
import json
import zipfile
from io import BytesIO
from html import escape
from pathlib import Path
from collections import Counter, defaultdict
from typing import Literal, Optional
from pydantic import BaseModel, Field, ValidationError

import difflib
import pandas as pd
import streamlit as st

try:
    from openai import OpenAI
except Exception:
    OpenAI = None
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


WF_KEYWORDS = {
    'IF', 'THEN', 'ELSE', 'AND', 'OR', 'NOT', 'EQ', 'NE', 'GT', 'LT', 'GE', 'LE',
    'CONTAINS', 'MISSING', 'LIKE', 'IN', 'IS', 'END', 'DEFINE', 'FILE', 'TABLE',
    'SUM', 'BY', 'WHERE', 'ON', 'COMPUTE', 'PRINT', 'LIST', 'JOIN', 'TO', 'UNIQUE',
    'AS', 'SET', 'DEFAULT', 'INCLUDE', 'FORMAT', 'NOPRINT', 'SUMMARIZE',
    'COLUMN', 'TOTAL', 'PAGE', 'NUM', 'OFF', 'STYLE', 'ENDSTYLE', 'PCHOLD',
    'HTMLCSS', 'UNITS', 'PAGESIZE', 'LEFTMARGIN', 'RIGHTMARGIN', 'TOPMARGIN',
    'BOTTOMMARGIN', 'SQUEEZE', 'ORIENTATION', 'PORTRAIT', 'LANDSCAPE', 'FONT',
    'SIZE', 'BOLD', 'ITALIC', 'NORMAL', 'COLOR', 'BACKCOLOR', 'BORDER', 'JUSTIFY',
    'CENTER', 'LEFT', 'RIGHT', 'WIDTH', 'LINE', 'OBJECT', 'TEXT', 'FIELD', 'ITEM',
    'TYPE', 'REPORT', 'TITLE', 'HEADING', 'FOOTING', 'SUBHEAD', 'SUBFOOT',
    'SUBTOTAL', 'GRANDTOTAL', 'ACROSSVALUE', 'ACROSSTITLE', 'TABHEADING',
    'TABFOOTING', 'SILVER', 'RGB', 'LIGHT',
    'MATCH', 'RUN', 'AFTER', 'MORE', 'OLD', 'NEW', 'HOLD', 'PCHOLD',
    'EDIT', 'LAST', 'DECODE', 'SUBSTR', 'TRIM', 'UPCASE', 'LOWCASE', 'LJUST',
    'RJUST', 'CONCAT', 'DATE', 'DMY', 'MDY', 'YMD', 'HDATE', 'TODAY',
    'DATEDIF', 'DATEADD', 'MIN', 'MAX', 'AVE', 'AVG', 'CNT', 'DST', 'FST', 'LST',
}

THIN = Side(style='thin', color='CCCCCC')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP_TOP = Alignment(wrap_text=True, vertical='top')

COLORS = {
    'header': ('1F4E79', 'FFFFFF'),
    'source': ('DDEBF7', '000000'),
    'define': ('E2EFDA', '000000'),
    'compute': ('FFF2CC', '000000'),
    'by_real': ('EBF3FB', '000000'),
    'by_calc': ('F4ECFA', '000000'),
    'grp_a': ('FFE699', '7B5B00'),
    'grp_b': ('C6EFCE', '276221'),
    'grp_c': ('DAEEF3', '17375E'),
    'grp_d': ('F2DCDB', '833C00'),
    'unique': ('F2F2F2', '7F7F7F'),
    'unparsed': ('EDEDED', 'AAAAAA'),
    's2_yes': ('FFE699', '7B5B00'),
    's2_no': ('F2F2F2', '7F7F7F'),
}

GROUP_COLOR_CYCLE = ['grp_a', 'grp_b', 'grp_c', 'grp_d']
FAST_DATA_FORMATTING = True
EXCEL_CELL_LIMIT = 32767
CLOUD_SAFE_WORKBOOK = True

FIELD_TYPE_COLORS = {
    'Source Field (DB Column)': COLORS['source'],
    'Calculated - DEFINE': COLORS['define'],
    'Calculated - COMPUTE': COLORS['compute'],
    'BY Field (Real)': COLORS['by_real'],
    'BY Field (Calculated)': COLORS['by_calc'],
}

FIELD_INVENTORY_HEADERS = [
    'File Path', 'File Name', 'Source Table', 'File Def', 'Fields Used',
    'Field Origin', 'Field Role', 'Formula Used', 'Raw Fields Used',
    'Result Datatype', 'Formula Source', 'Is Field Shown In Report',
    'Output Type',
]
FIELD_INVENTORY_WIDTHS = [42, 30, 28, 55, 32, 18, 24, 70, 40, 18, 18, 24, 38]

DUPLICATE_ANALYSIS_HEADERS = [
    'Duplicate Category', 'Group ID', 'Group Size', 'File Path', 'File Name',
    'Source Tables', 'Field Count', 'Exact Match Files', 'Near Match Files',
    'Same Data Source Files', 'Best Field Similarity', 'Difference Summary',
]
DUPLICATE_ANALYSIS_WIDTHS = [22, 16, 14, 42, 30, 60, 14, 55, 55, 55, 20, 70]

RESOURCE_ANALYZER_HEADERS = ['Resource Analyzer Program', 'Matched FEX File']
RESOURCE_ANALYZER_WIDTHS = [55, 55]

FINAL_SUMMARY_HEADERS = ['Metric', 'Count']
FINAL_SUMMARY_WIDTHS = [48, 38]

FILTER_PARAM_HEADERS = [
    'File Path', 'File Name', 'Filter Type', 'Expression', 'Parameter',
    'Fields Used',
]

FILTER_PARAM_WIDTHS = [42, 30, 16, 80, 35, 40]

FORMULA_MAPPING_HEADERS = [
    'File Path', 'File Name', 'Formula Source', 'Field Name',
    'Result Datatype', 'Formula Type', 'Raw WebFOCUS Formula',
    'Raw Columns Used', 'Parameters',
]

FORMULA_MAPPING_WIDTHS = [42, 30, 18, 28, 18, 22, 80, 40, 35]

JOIN_MATCH_HEADERS = [
    'File Path', 'File Name', 'Mapping Type', 'From Table', 'From Key',
    'To Table', 'To Key', 'WebFOCUS Join Mode',
    'Output HOLD', 'Raw WebFOCUS Statement',
]

JOIN_MATCH_WIDTHS = [42, 30, 18, 28, 32, 28, 32, 24, 24, 90]

VALIDATION_CHECKLIST_HEADERS = [
    'File Name', 'Check Category', 'Validation Check', 'Trigger Rule',
]
VALIDATION_CHECKLIST_WIDTHS = [30, 26, 90, 36]

TABLE_DB_MAPPING_HEADERS = [
    'Unique Table from KashMap', 'Actual Database Table', 'Connection',
    'MAS File', 'ACX File', 'Evidence / Notes',
]
TABLE_DB_MAPPING_WIDTHS = [40, 40, 20, 55, 55, 80]


# AI Build Plan 

Classification = Literal[
    "confirmed",
    "rule_based",
    "inferred",
    "suggested",
    "manual_review",
]

CognosLayer = Literal[
    "Data Source Query",
    "Framework Manager Query Subject",
    "Data Module",
    "Merged Query",
    "Calculation",
    "Report Filter",
    "Prompt",
    "List/Crosstab/Chart",
    "Drill-through",
    "Burst Report",
    "Manual Review",
]


class PlanStep(BaseModel):#one migration step
    step_number: int = Field(ge=1)
    title: str
    action: str
    cognos_layer: CognosLayer          
    classification: Classification
    reason: str
    referenced_sources: list[str] = Field(default_factory=list)
    referenced_fields: list[str] = Field(default_factory=list)
    confidence: int = Field(ge=0, le=100)
    manual_validation: Optional[str] = None


class ReportSummary(BaseModel):#overall information about the report
    report_name: str
    likely_purpose: str
    report_type: str
    complexity_level: Literal["Very Small", "Small", "Medium", "Large", "Very Complex"]
    metadata_level: Literal["fex_only"] = "fex_only"
    overall_confidence: int = Field(ge=0, le=100)


class BuildPlan(BaseModel):#structured AI output, and the two validation functions protect that output from unsupported AI-generated references.
    report_summary: ReportSummary
    implementation_plan: list[PlanStep]
    data_preparation: list[PlanStep] = Field(default_factory=list)
    calculation_plan: list[PlanStep] = Field(default_factory=list)
    filter_parameter_plan: list[PlanStep] = Field(default_factory=list)
    page_recommendations: list[PlanStep] = Field(default_factory=list)
    validation_checks: list[str] = Field(default_factory=list)
    manual_review_items: list[str] = Field(default_factory=list)
    limitations: list[str] = Field(default_factory=list)


def validate_plan_references(plan: BuildPlan, allowed_sources: set, allowed_fields: set) -> list[str]:
    """Cross-checks every table/field the AI referenced against the actual
    parsed facts for this report. Anything not present gets flagged so it
    can be stripped/downgraded rather than silently trusted."""
    errors: list[str] = []
    sections = [
        plan.implementation_plan,
        plan.data_preparation,
        plan.calculation_plan,
        plan.filter_parameter_plan,
        plan.page_recommendations,
    ]
    for section in sections:
        for step in section:
            unknown_sources = set(step.referenced_sources) - allowed_sources
            unknown_fields = set(step.referenced_fields) - allowed_fields
            if unknown_sources:
                errors.append(f"Step {step.step_number} references unknown sources: {sorted(unknown_sources)}")
            if unknown_fields:
                errors.append(f"Step {step.step_number} references unknown fields: {sorted(unknown_fields)}")
    return errors


def strip_unverified_steps(plan: BuildPlan, allowed_sources: set, allowed_fields: set) -> BuildPlan:
    """Downgrades any step referencing unknown sources/fields to manual_review
    classification and appends an explanatory note, instead of deleting it
    outright (keeps the plan complete while flagging what needs a human)."""
    def clean_section(section: list[PlanStep]) -> list[PlanStep]:
        cleaned = []
        for step in section:
            unknown_sources = set(step.referenced_sources) - allowed_sources
            unknown_fields = set(step.referenced_fields) - allowed_fields
            if unknown_sources or unknown_fields:
                step.classification = "manual_review"
                note = "AI recommendation flagged: referenced "
                parts = []
                if unknown_sources:
                    parts.append(f"unknown source(s) {sorted(unknown_sources)}")
                if unknown_fields:
                    parts.append(f"unknown field(s) {sorted(unknown_fields)}")
                step.manual_validation = note + " and ".join(parts) + " not present in the parsed report."
                step.confidence = min(step.confidence, 40)
            cleaned.append(step)
        return cleaned

    plan.implementation_plan = clean_section(plan.implementation_plan)
    plan.data_preparation = clean_section(plan.data_preparation)
    plan.calculation_plan = clean_section(plan.calculation_plan)
    plan.filter_parameter_plan = clean_section(plan.filter_parameter_plan)
    plan.page_recommendations = clean_section(plan.page_recommendations)
    return plan

#  Rule-based one


def generate_decode_case_expression(source_field, derived_field, pairs, final_table):
    lines = [f"-- Cognos calculation: {derived_field}", "CASE"]
    for source_value, mapped_value in pairs:
        safe_source = source_value.replace("'", "''")
        safe_mapped = mapped_value.replace("'", "''")
        lines.append(f"    WHEN [{final_table}].[{source_field}] = '{safe_source}' THEN '{safe_mapped}'")
    lines.append("    ELSE 'Unmapped'")
    lines.append("END")
    return '\n'.join(lines)


def generate_decode_lookup_table_csv(source_field, derived_field, pairs):
    lines = [f"{source_field},{derived_field}"]
    for source_value, mapped_value in pairs:
        safe_source = source_value.replace(',', ';').replace('"', "'")
        safe_mapped = mapped_value.replace(',', ';').replace('"', "'")
        lines.append(f"{safe_source},{safe_mapped}")
    return '\n'.join(lines)


def generate_decode_fm_merge_step(source_field, derived_field, lookup_table_name):
    return (
        f"-- Framework Manager / Data Module setup for {lookup_table_name}\n"
        f"1. Import {lookup_table_name} as a query subject / data module table.\n"
        f"2. Create a relationship: main table.{source_field} = {lookup_table_name}.{source_field}\n"
        f"3. Expose {lookup_table_name}.{derived_field} as a query item for reports to consume.\n"
        f"   (Alternative: build it as a Merged Query in Cognos Query Studio/Report Studio "
        f"if you cannot modify the Framework Manager model.)"
    )

def build_decode_migration_pack(mapping_df, final_table_name, case_threshold=60):
    packs = []
    if mapping_df is None or mapping_df.empty:
        return packs

    for table_name, group in mapping_df.groupby('Mapping Table Name'):
        first = group.iloc[0]
        source_field = first['Source Field']
        derived_field = first['Derived Field']
        pairs = [(r['Source Value'], r['Mapped Value']) for _, r in group.iterrows()]

        if len(pairs) <= case_threshold:
            packs.append({
                'mapping_table_name': table_name,
                'source_field': source_field,
                'derived_field': derived_field,
                'pair_count': len(pairs),
                'recommended_approach': 'Cognos CASE calculation (inline)',
                'code': generate_decode_case_expression(source_field, derived_field, pairs, final_table_name),
            })
        else:
            lookup_name = f"lkp_{sql_identifier(derived_field, 'lookup')}"
            packs.append({
                'mapping_table_name': table_name,
                'source_field': source_field,
                'derived_field': derived_field,
                'pair_count': len(pairs),
                'recommended_approach': f'Framework Manager / Data Module lookup table ({lookup_name})',
                'lookup_csv': generate_decode_lookup_table_csv(source_field, derived_field, pairs),
                'merge_code': generate_decode_fm_merge_step(source_field, derived_field, lookup_name),
            })

    return packs


def detect_positional_date_idiom(define_fields):
    """Structural detector (not name-based): finds fields derived via EDIT()
    with a positional digit mask (made of only $ and 9), followed by an
    IF-chain reassembly. Works regardless of what the fields are named."""
    edit_mask_fields = []
    for f in define_fields:
        formula = f['formula'].upper()
        if re.search(r"EDIT\s*\([A-Za-z_]\w*\s*,\s*'[\$9]+'\s*\)", formula):
            edit_mask_fields.append(f)

    if len(edit_mask_fields) < 4:
        return None

    reassembly_candidates = [
        f for f in define_fields
        if re.search(r'\bIF\b.+\bEQ\b.+\bTHEN\b', f['formula'].upper())
        and any(ef['field'].upper() in f['formula'].upper() for ef in edit_mask_fields)
    ]
    if not reassembly_candidates:
        return None

    return {
        'pattern': 'positional_date_deconstruction',
        'contributing_fields': [f['field'] for f in edit_mask_fields],
        'reassembly_field': reassembly_candidates[0]['field'],
        'confidence': 'Medium' if len(edit_mask_fields) >= 6 else 'Low',
    }


def detect_manual_if_chain_lookup(define_fields, min_branches=5):
    """Structural detector: finds a long IF/ELSE-IF chain testing the same
    variable against different literal values - functionally a lookup
    table written as code instead of DECODE."""
    results = []
    for f in define_fields:
        formula = f['formula']
        matches = re.findall(r"IF\s+(\w+)\s+EQ\s+'[^']*'\s+THEN", formula, re.IGNORECASE)
        if not matches:
            continue
        var_counts = Counter(matches)
        for var, count in var_counts.items():
            if count >= min_branches:
                results.append({
                    'pattern': 'manual_if_chain_lookup',
                    'field': f['field'],
                    'tested_variable': var,
                    'branch_count': count,
                    'suggestion': (
                        'This is functionally a lookup table written as IF/ELSE. '
                        'Consider converting to a DECODE (auto-translated by the rule engine) '
                        'or directly to a Framework Manager/Data Module lookup table.'
                    ),
                })
    return results


def build_rule_findings(parsed, mapping_df, final_table_name):
    """Produces the deterministic rule_findings block that gets marked
    authoritative in the AI prompt, so the model never re-derives what the
    rule engine already resolved with certainty."""
    findings = []

    decode_pack = build_decode_migration_pack(mapping_df, final_table_name)
    for item in decode_pack:
        findings.append({
            'type': 'large_decode',
            'field': item['derived_field'],
            'mapping_count': item['pair_count'],
            'recommendation': f"Use {item['recommended_approach']}.",
            'confidence': 100,
        })

    date_idiom = detect_positional_date_idiom(parsed.get('define_fields', []))
    if date_idiom:
        findings.append({
            'type': 'positional_date',
            'contributing_fields': date_idiom['contributing_fields'],
            'reassembly_field': date_idiom['reassembly_field'],
            'recommendation': 'Convert positional date-string parsing into a typed Date query item in Framework Manager / Data Module.',
            'confidence': 90 if date_idiom['confidence'] == 'Medium' else 70,
        })

    if_chain_findings = detect_manual_if_chain_lookup(parsed.get('define_fields', []))
    for item in if_chain_findings:
        findings.append({
            'type': 'manual_if_chain_lookup',
            'field': item['field'],
            'tested_variable': item['tested_variable'],
            'branch_count': item['branch_count'],
            'recommendation': item['suggestion'],
            'confidence': 85,
        })

    return findings, decode_pack

def _fill(bg):
    return PatternFill('solid', start_color=bg)


def _font(fg, bold=False):
    return Font(name='Arial', size=9, color=fg, bold=bold)


def _excel_safe_value(val):
    if isinstance(val, str) and len(val) > EXCEL_CELL_LIMIT:
        return val[:EXCEL_CELL_LIMIT - 20] + ' ... [truncated]'

    return val


def _write_cell(ws, row, col, val, bg='FFFFFF', fg='000000', bold=False):
    c = ws.cell(row=row, column=col, value=_excel_safe_value(val))

    if FAST_DATA_FORMATTING and row > 1:
        if bold:
            c.font = _font(fg, bold)
        c.alignment = WRAP_TOP
        return

    c.fill = _fill(bg)
    c.font = _font(fg, bold)
    c.alignment = WRAP_TOP
    c.border = BORDER


def setup_sheet(ws, headers, widths):
    hbg, hfg = COLORS['header']

    for col, (h, w) in enumerate(zip(headers, widths), 1):
        _write_cell(ws, 1, col, h, hbg, hfg, bold=True)
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"


def blank_llm_translation(status='Not requested'):
    return {
        'suggested_cognos_type': '',
        'suggested_cognos_expression_action': '',
        'llm_notes': '',
        'needs_manual_review': 'Yes' if status != 'Not requested' else '',
        'review_reason': status,
        'confidence': '',
    }


def get_openai_client():
    if OpenAI is None:
        return None

    api_key = (
        st.secrets.get('OPENAI_API_KEY', '')
        or os.environ.get('OPENAI_API_KEY', '')
        or st.session_state.get('openai_api_key', '')
    )
    if not api_key or api_key == 'paste_your_openai_api_key_here':
        return None

    try:
        return OpenAI(api_key=api_key)
    except Exception:
        return None


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


def normalize_llm_text(value):
    if isinstance(value, list):
        return '\n\n'.join(str(item).strip() for item in value if str(item).strip())
    if isinstance(value, dict):
        return json.dumps(value, indent=2)
    return str(value or '').strip()


def translate_expression_with_llm(client, model, expression_type, expression, context):
    if not expression:
        return blank_llm_translation('No raw expression to translate')

    cache_key = (model, expression_type, expression, context)
    cache = st.session_state.setdefault('llm_translation_cache', {})
    if cache_key in cache:
        return cache[cache_key]


    prompt = f"""
You translate WebFOCUS migration logic into IBM Cognos Analytics guidance.

Return only valid JSON with these exact keys:
suggested_cognos_type
suggested_cognos_expression_action
llm_notes
needs_manual_review
review_reason
confidence

Rules:
- The value for suggested_cognos_expression_action must be one plain multiline string.
- Do not return suggested_cognos_expression_action as a JSON array, Python list, bullet list, or numbered list.
- If there are multiple calculated fields, format them in one multiline string like:
  PROFIT =
  [SALES].[SALES_AMOUNT] - [SALES].[COST_AMOUNT]

  SALES_LABEL =
  [SALES].[REGION] || ' - ' || [SALES].[PRODUCT_CATEGORY]
- Prefer a Framework Manager calculated query item for reusable, row-level DEFINE/COMPUTE logic.
- Prefer a report-level Cognos calculation only when the logic is specific to one report/page.
- Prefer Framework Manager/Data Module relationships or a Cognos Merged Query for source shaping, joins, merges, and cleanup.
- For JOIN or MATCH, explain the likely Framework Manager relationship or Cognos Merged Query approach.
- For filters and joins, the suggested output may be an action, not a formula.
- Keep query subject/query item names as placeholders when the exact Cognos model name is unknown.
- Be honest about uncertainty. The result is a migration suggestion, not final production content.
- Set needs_manual_review to "Yes".
- Set confidence to High, Medium, or Low.

Expression type: {expression_type}
Context: {context}
Raw WebFOCUS expression:
{expression}
""".strip()

    try:
        response = client.responses.create(
            model=model,
            input=prompt,
            temperature=0.1,
        )
        data = parse_llm_json(response.output_text)
        result = {
            'suggested_cognos_type': normalize_llm_text(data.get('suggested_cognos_type', '')),
            'suggested_cognos_expression_action': normalize_llm_text(data.get('suggested_cognos_expression_action', '')),
            'llm_notes': normalize_llm_text(data.get('llm_notes', '')),
            'needs_manual_review': normalize_llm_text(data.get('needs_manual_review', 'Yes')) or 'Yes',
            'review_reason': normalize_llm_text(data.get('review_reason', '')) or 'Review before using in Cognos.',
            'confidence': normalize_llm_text(data.get('confidence', '')),
        }
    except Exception as e:
        result = blank_llm_translation(f'LLM error: {e}')

    cache[cache_key] = result
    return result


def add_llm_fields(item, translation):
    item.update(translation)


def apply_llm_translations(parsed_results, llm_options):
    if not llm_options or not llm_options.get('enabled'):
        return

    client = get_openai_client()
    if client is None:
        st.warning("LLM translation is enabled, but OpenAI is not configured. Check .streamlit/secrets.toml and requirements.txt.")
        return

    model = llm_options.get('model') or 'gpt-4o-mini'
    max_rows = int(llm_options.get('max_rows') or 200)
    translated_count = 0

    progress = st.progress(0)
    status = st.empty()
    status.text("Translating raw expressions with LLM...")

    candidates = []
    for folder, fex_name, parsed in parsed_results:
        if parsed is None:
            continue

        if llm_options.get('formulas'):
            for item in parsed['formula_mapping']:
                candidates.append((
                    item,
                    item.get('formula', ''),
                    item.get('formula_source', 'FORMULA'),
                    f"File={fex_name}; Field={item.get('field', '')}; Source={item.get('formula_source', '')}; Datatype={item.get('format', '')}",
                ))

        if llm_options.get('filters'):
            for item in parsed['filter_parameter_mapping']:
                candidates.append((
                    item,
                    item.get('expression', ''),
                    item.get('type', 'FILTER'),
                    f"File={fex_name}; Parameters={', '.join(item.get('parameters', []))}; Fields={', '.join(item.get('fields', []))}",
                ))

    total = min(len(candidates), max_rows)
    for idx, (item, expression, expression_type, context) in enumerate(candidates[:max_rows], start=1):
        translation = translate_expression_with_llm(client, model, expression_type, expression, context)
        add_llm_fields(item, translation)
        translated_count += 1
        progress.progress(idx / total if total else 1.0)

    if len(candidates) > max_rows:
        skipped = len(candidates) - max_rows
        skipped_message = f"Skipped by LLM row limit. Increase the limit to translate the remaining {skipped} rows."
        for item, _, _, _ in candidates[max_rows:]:
            add_llm_fields(item, blank_llm_translation(skipped_message))

    progress.progress(1.0)
    status.text(f"LLM translation complete for {translated_count} row(s).")


def _numeric_measure_candidates(field_df):
    """Reuses the same value-field heuristic as the rule-based calculation
    inspector, so the AI payload and the rule-based fallback stay consistent."""
    candidates = []
    if field_df is None or field_df.empty:
        return candidates
    value_words = r'ORDER_VALUE|DOLLARS?_ORDERED|DOLLARS?_INVOICED|POUNDS?_ORDERED|POUNDS?_INVOICED|QTY|QUANTITY|RUN_HRS|DELAY_HRS|AMOUNT|COST|SALES|PRICE|COUNT|TOTAL'
    helper_words = r'DATE|DAY|OFFSET|RANK|SEQ|SEQUENCE|KEY|CODE|STATUS|FLAG|YEAR|MONTH|WEEK'
    field_col = 'Fields Used'
    if field_col not in field_df.columns:
        return candidates
    for field in field_df[field_col].dropna().astype(str).unique():
        if re.search(value_words, field, re.IGNORECASE) and not re.search(helper_words, field, re.IGNORECASE):
            candidates.append(field)
    return list(dict.fromkeys(candidates))[:15]


def build_ai_build_plan_payload(report_row, source_df, lineage_df, join_df, filter_df, formula_df, mapping_df, final_df, field_df):
    """Compact FEX-only payload for the single AI call. Rule-resolved items
    (DECODE, structural idioms) are summarized as authoritative rule_findings
    rather than re-sent as raw content, so the AI never re-derives what's
    already certain and never receives duplicated large mapping content."""

    sources = []
    if source_df is not None and not source_df.empty and 'Source Name' in source_df.columns:
        real = source_df[source_df.get('Notes', '') == 'Original report source'] if 'Notes' in source_df.columns else source_df
        sources = unique_text_values(real['Source Name'].dropna().astype(str).tolist())[:20]

    hold_steps = []
    if lineage_df is not None and not lineage_df.empty:
        for _, row in lineage_df.sort_values('Step #').iterrows():
            output_table = str(row.get('Output Table', '') or '').strip()
            if not output_table:
                continue
            hold_steps.append({
                'table': output_table,
                'does': str(row.get('What Happens In This Step', ''))[:80],
                'keys': str(row.get('Key Fields / BY / ACROSS', ''))[:60],
            })
    hold_steps = hold_steps[:30]

    joins = []
    high_risk_joins = []
    if join_df is not None and not join_df.empty:
        for _, row in join_df.iterrows():
            mapping_type = str(row.get('Mapping Type', ''))
            entry = {
                'type': mapping_type,
                'left': f"{row.get('From Table', '')}[{row.get('From Key', '')}]",
                'right': f"{row.get('To Table', '')}[{row.get('To Key', '')}]",
            }
            joins.append(entry)
            join_mode = str(row.get('WebFOCUS Join Mode', '')).upper()
            if mapping_type == 'MATCH FILE' or join_mode == 'JOIN TO ALL':
                high_risk_joins.append({**entry, 'raw': str(row.get('Raw WebFOCUS Statement', ''))[:200]})

    flagged_filters = []
    if filter_df is not None and not filter_df.empty:
        for _, row in filter_df.iterrows():
            expr = str(row.get('Expression', ''))
            if expr.count("' AND '") >= 4 or expr.count("' OR '") >= 4:
                flagged_filters.append({'expr': expr[:300], 'value_count': expr.count("'") // 2})

    filter_summary = []
    if filter_df is not None and not filter_df.empty:
        seen = set()
        for _, row in filter_df.iterrows():
            expr = str(row.get('Expression', ''))
            key = expr[:40]
            if key in seen:
                continue
            seen.add(key)
            filter_summary.append({'type': str(row.get('Filter Type', '')), 'expr': expr[:80]})
        filter_summary = filter_summary[:15]

    final_shape = ''
    final_output = ''
    if final_df is not None and not final_df.empty:
        final_shape = str(final_df.iloc[0].get('Dataset Shape', ''))
        final_output = str(final_df.iloc[0].get('Recommended Final Dataset', ''))

    shown_fields = []
    if field_df is not None and not field_df.empty and 'Is Field Shown In Report' in field_df.columns:
        shown = field_df[field_df['Is Field Shown In Report'] == 'Yes']
        shown_fields = unique_text_values(shown['Fields Used'].dropna().astype(str).tolist())[:40]

    # Formulas NOT already resolved by DECODE rule findings (avoid re-sending
    # the same content the rule engine already handled deterministically).
    decode_fields = set()
    if mapping_df is not None and not mapping_df.empty and 'Derived Field' in mapping_df.columns:
        decode_fields = set(mapping_df['Derived Field'].dropna().astype(str).unique())

    unresolved_formulas = []
    if formula_df is not None and not formula_df.empty:
        risky_mask = formula_df['Formula Type'].astype(str).str.contains(
            'Conditional IF|Parameter-based|Arithmetic|Concatenation/String|Date/Time',
            case=False, na=False
        )
        risky_df = formula_df[risky_mask]
        for _, row in risky_df.iterrows():
            field = str(row.get('Field Name', ''))
            if field in decode_fields:
                continue
            unresolved_formulas.append({
                'field': field,
                'formula_type': str(row.get('Formula Type', '')),
                'raw_formula': str(row.get('Raw WebFOCUS Formula', ''))[:300],
                'raw_columns_used': str(row.get('Raw Columns Used', '')),
            })
        unresolved_formulas = unresolved_formulas[:15]

    return {
        'report_name': str(report_row.get('FEX Name', '')),
        'complexity_level': str(report_row.get('Complexity', 'Medium')),
        'final_output_table': final_output or str(report_row.get('Recommended Final Dataset', '')),
        'final_dataset_shape': final_shape,
        'output_type': str(report_row.get('Output Format', '')),
        'sources': sources,
        'shown_fields': shown_fields,
        'hold_staging_steps': hold_steps,
        'joins': joins,
        'high_risk_joins': high_risk_joins,
        'filters_sample': filter_summary,
        'flagged_large_filters': flagged_filters,
        'unresolved_formulas': unresolved_formulas,
        'numeric_measure_candidates': _numeric_measure_candidates(field_df),
        'constraints': {
            'database_metadata_available': False,
            'may_not_state_primary_keys': True,
            'may_not_state_cardinality': True,
            'may_not_state_confirmed_grain': True,
            'may_not_invent_tables_or_fields': True,
        },
    }


AI_BUILD_PLAN_SYSTEM_PROMPT = """You are a senior BI migration engineer converting one parsed WebFOCUS report to IBM Cognos Analytics.
The input contains FEX-derived report logic only (no verified database primary keys, cardinality, table grain,
or complete relationships). rule_findings, if present in the input, are deterministic and authoritative - never
re-derive or contradict them.

Cognos target concepts to use:
- Framework Manager / Data Module: where source tables, relationships, and reusable query subjects/items live.
- Query Subject / Query Item: Framework Manager modeled tables/columns.
- Merged Query: combining two data sources at the report level when no Framework Manager relationship exists.
- Calculation: a report-level or model-level derived value (equivalent to WebFOCUS DEFINE/COMPUTE).
- Prompt: equivalent to a WebFOCUS parameter (&VARIABLE); used for filters the user provides at run time.
- Report Filter (detail or summary filter): equivalent to WebFOCUS WHERE/IF.
- List/Crosstab/Chart: the visual object in a Cognos report.
- Drill-through: equivalent to a WebFOCUS drilldown/URL link between reports.
- Burst Report: equivalent to distributing/segmenting a WebFOCUS scheduled output by a key field.

Rules:
1. Use only tables, sources, fields, and calculations present in the input. Never invent a table, field,
   measure, join condition, or business rule.
2. Do not claim primary keys, cardinality, fact/dimension roles, or table grain unless explicitly supplied.
3. Label every implementation step as one of: confirmed, rule_based, inferred, suggested, manual_review.
4. Every step must include referenced_sources and referenced_fields drawn only from the input - this is used
   for automated validation, so accuracy here is critical.
5. Any uncertainty must go into manual_review_items, not be silently guessed.
6. Confidence (0-100) must reflect how directly the input supports the recommendation, not general optimism.
7. Return ONLY strict JSON matching the required schema - no markdown fences, no commentary outside the JSON.

Return JSON with exactly these keys:
{
  "report_summary": {"report_name": "", "likely_purpose": "", "report_type": "",
                      "complexity_level": "Very Small|Small|Medium|Large|Very Complex",
                      "metadata_level": "fex_only", "overall_confidence": 0},
  "implementation_plan": [{"step_number": 1, "title": "", "action": "",
                            "cognos_layer": "Data Source Query|Framework Manager Query Subject|Data Module|Merged Query|Calculation|Report Filter|Prompt|List/Crosstab/Chart|Drill-through|Burst Report|Manual Review",
                            "classification": "confirmed|rule_based|inferred|suggested|manual_review",
                            "reason": "", "referenced_sources": [], "referenced_fields": [],
                            "confidence": 0, "manual_validation": null}],
  "data_preparation": [... same PlanStep shape ...],
  "calculation_plan": [... same PlanStep shape ..., must cover every item in unresolved_formulas],
  "filter_parameter_plan": [... same PlanStep shape ..., must cover flagged_large_filters and high_risk_joins],
  "page_recommendations": [... same PlanStep shape ...],
  "validation_checks": ["..."],
  "manual_review_items": ["..."],
  "limitations": ["..."]
}

Every entry in high_risk_joins must appear in filter_parameter_plan or implementation_plan with classification
"manual_review" and a reason explaining the specific risk (MATCH FILE OLD/NEW/AFTER/MORE behavior, or JOIN TO ALL
cardinality). Every entry in unresolved_formulas must get exactly one calculation_plan step.
"""

VALID_COGNOS_LAYERS = {
    "Data Source Query", "Framework Manager Query Subject", "Data Module",
    "Merged Query", "Calculation", "Report Filter", "Prompt",
    "List/Crosstab/Chart", "Drill-through", "Burst Report", "Manual Review",
}

_LAYER_PRECEDENCE = [
    "Merged Query", "Data Source Query", "Framework Manager Query Subject", "Data Module",
    "Calculation", "Report Filter", "Prompt",
    "List/Crosstab/Chart", "Drill-through", "Burst Report", "Manual Review",
]

def _coerce_cognos_layer(value):
    if value in VALID_COGNOS_LAYERS:
        return value
    parts = re.split(r'[|/,]', str(value))
    parts = [p.strip() for p in parts if p.strip() in VALID_COGNOS_LAYERS]
    if parts:
        for preferred in _LAYER_PRECEDENCE:
            if preferred in parts:
                return preferred
        return parts[0]
    return "Manual Review"

def normalize_cognos_layers(raw):
    section_keys = [
        'implementation_plan', 'data_preparation', 'calculation_plan',
        'filter_parameter_plan', 'page_recommendations',
    ]
    for key in section_keys:
        steps = raw.get(key)
        if not isinstance(steps, list):
            continue
        for step in steps:
            if isinstance(step, dict) and 'cognos_layer' in step:
                step['cognos_layer'] = _coerce_cognos_layer(step['cognos_layer'])
    return raw


_COMPLEXITY_ALIASES = {'High': 'Very Complex', 'Low': 'Small'}

def generate_ai_build_plan(client, model, payload):
    """Single on-demand API call, with automatic retry on truncation.
    Returns (BuildPlan | None, error_str | None)."""
    user_prompt = json.dumps(payload, ensure_ascii=False)

    complexity = payload.get('complexity_level', 'Medium')
    complexity = _COMPLEXITY_ALIASES.get(complexity, complexity)

    base_budget = {
        'Very Small': 1500, 'Small': 2200, 'Medium': 3500,
        'Large': 5500, 'Very Complex': 7000,
    }.get(complexity, 3500)

    max_budget_cap = 16000
    output_budget = base_budget
    last_raw_text = ''

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
                response = client.responses.create(
                    model=model,
                    input=[
                        {"role": "system", "content": AI_BUILD_PLAN_SYSTEM_PROMPT},
                        {"role": "user", "content": user_prompt},
                    ],
                    temperature=0.1,
                    max_output_tokens=output_budget,
                )
                last_raw_text = getattr(response, 'output_text', '') or ''
                was_truncated = getattr(response, 'status', None) == 'incomplete'
        except Exception as e:
            return None, f'LLM error: {e}'

        raw = parse_llm_json(last_raw_text)

        if not raw and was_truncated and output_budget < max_budget_cap:
            output_budget = min(output_budget * 2, max_budget_cap)
            continue

        if not raw:
            reason = (
                ' (response was truncated before completing JSON - '
                'consider shrinking the payload or raising max_output_tokens further)'
                if was_truncated else ''
            )
            return None, f'Model did not return parseable JSON{reason}. Raw response (truncated): {last_raw_text[:500]}'

        raw = normalize_cognos_layers(raw)

        try:
            plan = BuildPlan.model_validate(raw)
        except ValidationError as ve:
            return None, f'AI response failed schema validation: {ve}'

        allowed_sources = set(payload.get('sources', []))
        # HOLD/staging tables are legitimate pipeline elements even though they
        # aren't "Original report source" - include them so the AI isn't
        # penalized for correctly citing intermediate HOLD tables.
        allowed_sources |= {
            step.get('table', '') for step in payload.get('hold_staging_steps', [])
            if step.get('table')
        }
        for j in payload.get('joins', []) + payload.get('high_risk_joins', []):
            for side_key in ('left', 'right'):
                table_part = j.get(side_key, '').split('[')[0].strip()
                if table_part:
                    allowed_sources.add(table_part)

        allowed_fields = set(payload.get('shown_fields', [])) | {
            item['field'] for item in payload.get('unresolved_formulas', [])
        } | {c for c in payload.get('numeric_measure_candidates', [])}

        for item in payload.get('unresolved_formulas', []):
            raw_cols = item.get('raw_columns_used', '')
            if raw_cols:
                allowed_fields.update(c.strip() for c in raw_cols.split(',') if c.strip())

        for f in payload.get('filters_sample', []) + payload.get('flagged_large_filters', []):
            expr = f.get('expr', '')
            allowed_fields.update(re.findall(r'\b[A-Z][A-Z0-9_]{2,}\b', expr))

        # Key/grouping fields mentioned in HOLD staging steps ("BY WO_NBR, BY OPNO2")
        # are real fields used throughout the pipeline, not just output-shown fields.
        for step in payload.get('hold_staging_steps', []):
            allowed_fields.update(re.findall(r'\b[A-Z][A-Z0-9_]{2,}\b', step.get('keys', '').upper()))

        # Fields embedded in join descriptions (format "TABLE[FIELD]")
        for j in payload.get('joins', []) + payload.get('high_risk_joins', []):
            for side_key in ('left', 'right'):
                m = re.search(r'\[([^\]]+)\]', j.get(side_key, ''))
                if m:
                    allowed_fields.add(m.group(1).strip().upper())

        plan = strip_unverified_steps(plan, allowed_sources, allowed_fields)
        return plan, None
    





AI_SQL_VIEW_SYSTEM_PROMPT = """You are a senior BI migration engineer converting a parsed WebFOCUS report's data
preparation logic into ONE production-oriented SQL view for IBM Cognos (Framework Manager / Data Module) to consume.

You will receive JSON describing:
- real_sources: real source tables
- table_resolution: maps each real_source to {resolved_name, confirmed}
- hold_staging_steps: HOLD/staging steps in execution order, with input/output tables and BY/ACROSS keys
- joins: JOIN and MATCH FILE operations with from/to tables and keys
- filters: WHERE/IF filters and parameters
- formulas: DEFINE/COMPUTE formulas, with the raw WebFOCUS formula and raw columns used
- shown_fields: fields actually shown/used in the report output

HARD RULES:
1. Use ONLY tables, fields, and keys explicitly present in the input JSON. Never invent a table or column.
2. If something needed (data type, exact join cardinality, key uniqueness, a parameter's default) is not present
   in the input, do NOT guess silently. Add a `-- REVIEW:` comment on that exact line explaining what must be confirmed.
3. Structure the SQL as `CREATE VIEW <view_name> AS` built from CTEs: one CTE per real source table, then one CTE
   per HOLD/staging step in original order, then the JOIN/MATCH logic, then the final SELECT.
3b. table_resolution maps each real_source to {resolved_name, confirmed}. When confirmed=true, use resolved_name
    directly with no review comment - it is a verified physical database table. When confirmed=false, use
    resolved_name (the WebFOCUS synonym) but add `-- REVIEW: WebFOCUS synonym name, confirm actual DB table`
    on that CTE's FROM line.
3c. At the very top of the SQL, before the CREATE VIEW statement, add a comment block based on
    constraints.database_metadata_available:
    - If true: add "-- METADATA STATUS: MAS/ACX metadata was provided. Table names below marked
      as confirmed are verified real database tables."
    - If false: add "-- ⚠️ METADATA STATUS: NO MAS/ACX FILE WAS UPLOADED. ⚠️" followed by
      "-- ALL TABLE NAMES BELOW ARE WEBFOCUS ALIAS/SYNONYM NAMES, NOT CONFIRMED REAL DATABASE TABLES."
      followed by "-- This SQL WILL NOT RUN until every table name is manually verified/replaced."
4. For MATCH FILE steps, do not assume a plain INNER JOIN. Emit a LEFT JOIN by default with
   `-- REVIEW: MATCH FILE OLD/NEW/AFTER/MORE behavior must be confirmed before trusting this LEFT JOIN`.
5. Reproduce filters as WHERE clauses in the right CTE. For parameter-driven filters, use a placeholder bind
   variable and add `-- REVIEW: WebFOCUS parameter (&PARAM) - decide if this becomes a Cognos prompt or fixed value`.
6. Reproduce DEFINE/COMPUTE formulas as computed columns in ANSI SQL (CASE for IF/THEN/ELSE, || for concatenation,
   normal arithmetic). Do not change the logic.
7. SELECT only fields known to be used/shown (shown_fields, and fields referenced in formulas/keys). Never SELECT *.
8. Return ONLY the SQL code. No markdown fences, no commentary before or after.
9. This is still a review-required draft, not guaranteed-correct production SQL - every line must trace back to the input JSON.
"""


def build_ai_sql_view_payload(report_row, source_df, lineage_df, join_df, filter_df, formula_df, final_df, field_df=None, table_db_mapping_df=None):
    view_name = sql_view_name(report_row)

    sources = []
    if source_df is not None and not source_df.empty and 'Source Name' in source_df.columns:
        real = source_df[source_df.get('Notes', '') == 'Original report source'] if 'Notes' in source_df.columns else source_df
        sources = unique_text_values(real['Source Name'].dropna().astype(str).tolist())

    # Resolve each WebFOCUS synonym to a real DB table name if MAS/ACX metadata is available
    table_resolution = {}
    mapping_lookup = {}
    if table_db_mapping_df is not None and not table_db_mapping_df.empty:
        for _, row in table_db_mapping_df.iterrows():
            key = normalize_name(row.get('Unique Table from KashMap', ''))
            actual = str(row.get('Actual Database Table', '') or '').strip()
            if key and actual:
                mapping_lookup[key] = actual

    for source in sources:
        key = normalize_name(source)
        if key in mapping_lookup:
            table_resolution[source] = {'resolved_name': mapping_lookup[key], 'confirmed': True}
        else:
            table_resolution[source] = {'resolved_name': source, 'confirmed': False}

    hold_steps = []
    if lineage_df is not None and not lineage_df.empty:
        for _, row in lineage_df.sort_values('Step #').iterrows():
            hold_steps.append({
                'step': int(row.get('Step #', 0)),
                'input_tables': str(row.get('Input Table(s)', '')),
                'output_table': str(row.get('Output Table', '')),
                'what_happens': str(row.get('What Happens In This Step', '')),
                'keys': str(row.get('Key Fields / BY / ACROSS', '')),
                'formulas_created': str(row.get('Formulas Created', ''))[:400],
                'filters_applied': str(row.get('Filters Applied', ''))[:400],
            })

    joins = []
    if join_df is not None and not join_df.empty:
        for _, row in join_df.iterrows():
            joins.append({
                'type': str(row.get('Mapping Type', '')),
                'from_table': str(row.get('From Table', '')),
                'from_key': str(row.get('From Key', '')),
                'to_table': str(row.get('To Table', '')),
                'to_key': str(row.get('To Key', '')),
                'webfocus_join_mode': str(row.get('WebFOCUS Join Mode', '')),
                'output_hold': str(row.get('Output HOLD', '')),
                'raw_statement': str(row.get('Raw WebFOCUS Statement', ''))[:300],
            })

    filters = []
    if filter_df is not None and not filter_df.empty:
        for _, row in filter_df.iterrows():
            filters.append({
                'type': str(row.get('Filter Type', '')),
                'expression': str(row.get('Expression', ''))[:300],
                'parameter': str(row.get('Parameter', '')),
                'fields_used': str(row.get('Fields Used', '')),
            })

    formulas = []
    if formula_df is not None and not formula_df.empty:
        for _, row in formula_df.iterrows():
            formulas.append({
                'source': str(row.get('Formula Source', '')),
                'field': str(row.get('Field Name', '')),
                'result_datatype': str(row.get('Result Datatype', '')),
                'formula_type': str(row.get('Formula Type', '')),
                'raw_formula': str(row.get('Raw WebFOCUS Formula', ''))[:300],
                'raw_columns_used': str(row.get('Raw Columns Used', '')),
            })

    shown_fields = []
    if field_df is not None and not field_df.empty and 'Is Field Shown In Report' in field_df.columns:
        shown = field_df[field_df['Is Field Shown In Report'] == 'Yes']
        shown_fields = unique_text_values(shown['Fields Used'].dropna().astype(str).tolist())

    final_output = ''
    if final_df is not None and not final_df.empty:
        final_output = str(final_df.iloc[0].get('Recommended Final Dataset', ''))

    return {
        'view_name': view_name,
        'real_sources': sources,
        'table_resolution': table_resolution,
        'hold_staging_steps': hold_steps,
        'joins': joins,
        'filters': filters,
        'formulas': formulas,
        'shown_fields': shown_fields,
        'final_output_table': final_output or view_name,
        'constraints': {
            'database_metadata_available': bool(mapping_lookup),
            'may_not_invent_tables_or_fields': True,
            'may_not_state_primary_keys_or_cardinality_unless_given': True,
        },
    }


def generate_ai_sql_view(client, model, payload):
    """Single on-demand API call that returns real production-oriented SQL text
    (not JSON) for the SQL view, using only facts already parsed from the FEX."""
    user_prompt = json.dumps(payload, ensure_ascii=False)

    try:
        response = client.responses.create(
            model=model,
            input=[
                {"role": "system", "content": AI_SQL_VIEW_SYSTEM_PROMPT},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.1,
            max_output_tokens=4000,
        )
    except Exception as e:
        return None, f'LLM error: {e}'

    sql_text = (response.output_text or '').strip()
    sql_text = re.sub(r'^```sql\s*', '', sql_text, flags=re.IGNORECASE)
    sql_text = re.sub(r'^```\s*', '', sql_text)
    sql_text = re.sub(r'```\s*$', '', sql_text)
    sql_text = sql_text.strip()

    if not sql_text:
        return None, 'Model returned an empty response.'

    return sql_text, None
















    
def sort_overview_df(df, sort_by):
    if df is None or df.empty:
        return df

    if sort_by == "Name (Z-A)":
        return df.sort_values(['FEX Name', 'File Path'], ascending=[False, True])
    if sort_by == "Complexity (High to Low)":
        order = {'High': 0, 'Medium': 1, 'Low': 2}
        tmp = df.copy()
        tmp['_complexity_rank'] = tmp['Complexity'].map(order).fillna(3)
        return tmp.sort_values(['_complexity_rank', 'FEX Name'])
    if sort_by == "Duplicate Type":
        return df.sort_values(['Duplicate Type', 'FEX Name'])
    if sort_by == "Join/Match Count (High to Low)":
        return df.sort_values(['Join / Match Count', 'FEX Name'], ascending=[False, True])
    if sort_by == "Formula Count (High to Low)":
        return df.sort_values(['Formula Count', 'FEX Name'], ascending=[False, True])

    return df.sort_values(['FEX Name', 'File Path'])    


def _render_plan_step(step, source_df):
    level_map = {'confirmed': 'ok', 'rule_based': 'ok', 'inferred': 'medium', 'suggested': 'medium', 'manual_review': 'high'}
    level = level_map.get(step.classification, 'medium')
    with st.container(border=True):
        st.markdown(f"**{step.title}** — {step.cognos_layer}")
        st.markdown(
            f"{badge(step.classification.replace('_', ' ').title(), level)} "
            f"{badge(f'Confidence: {step.confidence}%')}",
            unsafe_allow_html=True,
        )
        st.write(step.action)
        st.caption(f"Why: {step.reason}")
        if step.referenced_sources or step.referenced_fields:
            refs = ', '.join(step.referenced_sources + step.referenced_fields)
            st.caption(f"References: {refs}")
        if step.manual_validation:
            st.warning(step.manual_validation)


def build_cognos_prep_build_order(lineage_df, plan):
    """Combines the HOLD chain's real execution order (from lineage_df) with
    the AI plan's Framework Manager / Data Module steps, so the report maker
    gets one numbered sequence matching the order they'd build query subjects
    or Data Module steps - instead of hunting through separate Data Prep/
    Calculation cards."""
    ordered_stages = []

    if lineage_df is not None and not lineage_df.empty:
        for _, row in lineage_df.sort_values('Step #').iterrows():
            output_table = str(row.get('Output Table', '') or '').strip()
            if not output_table:
                continue
            ordered_stages.append({
                'stage_number': int(row.get('Step #', len(ordered_stages) + 1)),
                'output_table': output_table,
                'what_happens': str(row.get('What Happens In This Step', '')),
                'keys': str(row.get('Key Fields / BY / ACROSS', '')),
                'matched_steps': [],
            })

    if plan is not None:
        prep_steps = [
            step for step in (plan.data_preparation + plan.calculation_plan + plan.filter_parameter_plan)
            if step.cognos_layer in {"Merged Query", "Framework Manager Query Subject", "Data Module"}
        ]
        for step in prep_steps:
            matched = False
            for stage in ordered_stages:
                refs = set(step.referenced_sources) | set(step.referenced_fields)
                if stage['output_table'] in refs or any(r in stage['what_happens'] for r in refs):
                    stage['matched_steps'].append(step)
                    matched = True
                    break
            if not matched and ordered_stages:
                ordered_stages[-1]['matched_steps'].append(step)
            elif not matched:
                ordered_stages.append({
                    'stage_number': len(ordered_stages) + 1,
                    'output_table': step.title,
                    'what_happens': step.action,
                    'keys': '',
                    'matched_steps': [step],
                })

    return ordered_stages


def render_cognos_prep_build_order(lineage_df, plan):
    stages = build_cognos_prep_build_order(lineage_df, plan)
    if not stages:
        st.info("No staging steps detected for this report - it likely reads directly from source with no HOLD chain.")
        return

    st.markdown(
        "Build these as **Framework Manager query subjects / Data Module steps, in this exact order**. "
        "Each stage below corresponds to one WebFOCUS HOLD table, in the sequence it was originally computed."
    )

    for stage in stages:
        with st.container(border=True):
            st.markdown(f"**Stage {stage['stage_number']}: {stage['output_table']}**")
            st.caption(stage['what_happens'])
            if stage['keys']:
                st.caption(f"Key fields / grouping: {stage['keys']}")
            if not stage['matched_steps']:
                st.caption("No specific transformation steps returned for this stage - review manually.")
            for step in stage['matched_steps']:
                st.markdown(f"- **{step.title}**: {step.action}")
                if step.manual_validation:
                    st.warning(step.manual_validation)


def render_ai_build_plan(report_row, source_df, lineage_df, join_df, filter_df, formula_df, mapping_df, final_df, field_df):
    st.markdown(
        "Generates a **hybrid rule-based + AI** build plan: DECODE tables and structural idioms are "
        "translated deterministically (no AI, no cost), and one AI call organizes the remaining judgment "
        "calls - merge/relationship risk, novel formulas, and page/visual suggestions for **Cognos Analytics**."
    )

    final_table_name = final_df.iloc[0]['Recommended Final Dataset'] if final_df is not None and not final_df.empty else 'final_table'
    decode_pack = build_decode_migration_pack(mapping_df, final_table_name)

    st.markdown("#### Rule-Based Translations (no AI, zero cost)")
    if not decode_pack:
        st.info("No DECODE mapping tables were detected on this report.")
    for item in decode_pack:
        with st.container(border=True):
            st.markdown(f"**{item['mapping_table_name']}** ({item['pair_count']} pairs) → {item['recommended_approach']}")
            if 'code' in item:
                st.code(item['code'], language='sql')
            if 'lookup_csv' in item:
                st.caption("Lookup table CSV:")
                st.code(item['lookup_csv'], language='text')
                st.caption("Framework Manager / Data Module setup:")
                st.code(item['merge_code'], language='text')

    st.markdown("---")

    client = get_openai_client()
    if client is None:
        st.warning(
            "OpenAI is not configured. Add `OPENAI_API_KEY` to `.streamlit/secrets.toml` "
            "(e.g. `OPENAI_API_KEY = \"sk-...\"`) to enable the AI-assisted portion."
        )
        return

    report_key = re.sub(r'[^A-Za-z0-9_]+', '_', str(report_row.get('FEX Name', '')) + '_' + str(report_row.get('File Path', '')))
    cache = st.session_state.setdefault('ai_build_plan_cache', {})

    generate_clicked = st.button(
        "Generate AI build plan" if report_key not in cache else "Regenerate AI build plan",
        key=f"ai_build_plan_btn_{report_key}",
    )

    if generate_clicked:
        payload = build_ai_build_plan_payload(
            report_row, source_df, lineage_df, join_df, filter_df, formula_df, mapping_df, final_df, field_df
        )
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

    summary = plan.report_summary
    st.markdown("#### AI-Assisted Plan")
    c1, c2, c3 = st.columns(3)
    c1.metric("Complexity", summary.complexity_level)
    c2.metric("Overall confidence", f"{summary.overall_confidence}%")
    c3.metric("Manual review items", len(plan.manual_review_items))
    st.write(f"**Likely purpose:** {summary.likely_purpose}")
    st.write(f"**Report type:** {summary.report_type}")

    tabs = st.tabs(["Build Order", "Implementation Steps", "Data Prep", "Calculations", "Filters/Joins", "Pages", "Validation", "Manual Review"])
    with tabs[0]:
        render_cognos_prep_build_order(lineage_df, plan)

    with tabs[1]:
        if not plan.implementation_plan:
            st.info("No implementation steps returned.")
        for step in sorted(plan.implementation_plan, key=lambda s: s.step_number):
            _render_plan_step(step, source_df)

    with tabs[2]:
        if not plan.data_preparation:
            st.info("No data preparation steps returned.")
        for step in plan.data_preparation:
            _render_plan_step(step, source_df)

    with tabs[3]:
        if not plan.calculation_plan:
            st.info("No calculation translations returned.")
        for step in plan.calculation_plan:
            _render_plan_step(step, source_df)

    with tabs[4]:
        if not plan.filter_parameter_plan:
            st.info("No filter/join items returned.")
        for step in plan.filter_parameter_plan:
            _render_plan_step(step, source_df)

    with tabs[5]:
        if not plan.page_recommendations:
            st.info("No page/visual recommendations returned.")
        for step in plan.page_recommendations:
            _render_plan_step(step, source_df)

    with tabs[6]:
        if not plan.validation_checks:
            st.info("No validation checks returned.")
        for idx, check in enumerate(plan.validation_checks, start=1):
            st.write(f"{idx}. {check}")

    with tabs[7]:
        if not plan.manual_review_items:
            st.success("No manual review items flagged.")
        for idx, item in enumerate(plan.manual_review_items, start=1):
            st.warning(f"{idx}. {item}")
        if plan.limitations:
            st.caption("Limitations:")
            for lim in plan.limitations:
                st.caption(f"• {lim}")


def normalize_program_name(value):
    if value is None:
        return ''

    text = str(value).strip()

    if not text or text.lower() == 'nan':
        return ''

    text = text.replace('\\', '/')
    text = text.split('/')[-1]
    text = text.strip()

    if text.lower().endswith('.fex'):
        text = text[:-4]

    text = re.sub(r'[^A-Za-z0-9_.$#-]+', '', text)

    return text.upper()


def extract_program_tokens_from_text(value):
    if value is None:
        return set()

    text = str(value).strip()

    if not text or text.lower() == 'nan':
        return set()

    tokens = set()

    fex_matches = re.findall(
        r'([A-Za-z0-9_.$#/-]+\.fex)',
        text,
        flags=re.IGNORECASE
    )

    for item in fex_matches:
        normalized = normalize_program_name(item)
        if normalized:
            tokens.add(normalized)

    if not tokens:
        possible_words = re.findall(r'\b[A-Za-z0-9_.$#-]{3,}\b', text)

        for word in possible_words:
            cleaned = normalize_program_name(word)

            if cleaned and cleaned not in WF_KEYWORDS:
                tokens.add(cleaned)

    return tokens


def read_resource_analyzer_file(uploaded_ra_file):
    file_name = uploaded_ra_file.name.lower()
    program_names = set()
    raw_values = []

    uploaded_ra_file.seek(0)

    try:
        if file_name.endswith('.csv'):
            df_map = {
                'ResourceAnalyzer': pd.read_csv(uploaded_ra_file, dtype=str, header=None)
            }
        else:
            df_map = pd.read_excel(
                uploaded_ra_file,
                sheet_name=None,
                dtype=str,
                header=None
            )
    except Exception as e:
        st.warning(f"Resource Analyzer file could not be read properly: {e}")
        return program_names, raw_values

    for sheet_name, df in df_map.items():
        if df is None or df.empty:
            continue

        df = df.fillna('')

        for row_index in range(df.shape[0]):
            row_values = df.iloc[row_index].astype(str).tolist()

            for value in row_values:
                value = str(value).strip()

                if not value or value.lower() == 'nan':
                    continue

                extracted = extract_program_tokens_from_text(value)

                for item in extracted:
                    if item:
                        program_names.add(item)
                        raw_values.append((value, item))

    return program_names, raw_values


def filter_fex_items_by_resource_analyzer(fex_items, allowed_program_names):
    matched_items = []
    matched_pairs = []

    allowed_clean = set()

    for item in allowed_program_names:
        item_clean = normalize_program_name(item)
        allowed_clean.add(item_clean)
        allowed_clean.add(item_clean.replace('.FEX', ''))

    for folder, fex_name, content in fex_items:
        normalized_fex = normalize_program_name(fex_name)
        normalized_without_ext = normalized_fex.replace('.FEX', '')

        if normalized_fex in allowed_clean or normalized_without_ext in allowed_clean:
            matched_items.append((folder, fex_name, content))
            matched_pairs.append((normalized_without_ext, fex_name))

    return matched_items, matched_pairs


def raw_db_fields(formula, defined_names):
    formula_text = str(formula or '')

    # Remove quoted literal values before extracting field names.
    # Example:
    # DECODE(REGION, 'CHN', 'Chennai', 'BLR', 'Bangalore')
    # should identify REGION as the source field,
    # not CHN, Chennai, BLR, or Bangalore.
    formula_text = re.sub(
        r"'[^']*'|\"[^\"]*\"",
        " ",
        formula_text
    )

    tokens = re.findall(
        r'\b([A-Z][A-Z0-9_]{2,}|[0-9][A-Z][A-Z0-9_]{1,})\b',
        formula_text.upper()
    )

    return sorted({
        t for t in tokens
        if t not in WF_KEYWORDS
        and t not in defined_names
        and not t.isdigit()
    })

def strip_comments(text):
    lines = [
        line for line in text.splitlines()
        if not line.strip().startswith('-*')
        and not line.strip().startswith('-!')
    ]

    return '\n'.join(lines)


def normalize_expression(value):
    return ' '.join(str(value or '').replace('\r', ' ').split())


def extract_parameters(value):
    params = re.findall(r'&&?[A-Za-z_][A-Za-z0-9_]*', str(value or ''))
    return sorted(dict.fromkeys(p.upper() for p in params))


def extract_candidate_fields(value, defined_names=None):
    defined_names = defined_names or set()
    text = re.sub(r"'[^']*'|\"[^\"]*\"", ' ', str(value or ''))
    text = re.sub(r'&&?[A-Za-z_][A-Za-z0-9_]*', ' ', text)
    tokens = re.findall(r'\b[A-Za-z_][A-Za-z0-9_]*(?:\.[A-Za-z_][A-Za-z0-9_]*)?\b', text)
    fields = []

    skip_next = False
    for token in tokens:
        token_upper = token.upper()

        if skip_next:
            skip_next = False
            continue

        if token_upper == 'AS':
            skip_next = True
            continue

        if token_upper in WF_KEYWORDS:
            continue

        if token_upper in {
            'BY', 'WHERE', 'IF', 'ON', 'TABLE', 'END', 'COMPUTE', 'DEFINE',
            'FORMAT', 'HOLD', 'PCHOLD', 'MATCH', 'RUN', 'AFTER', 'MORE',
            'OVER', 'RECAP', 'HEADING', 'FOOTING', 'SUBHEAD', 'SUBFOOT',
            'NOPRINT', 'MISSING', 'TOTAL', 'RECOMPUTE', 'CNT', 'DST',
            'AVE', 'AVG', 'MIN', 'MAX', 'FST', 'LST',
        }:
            continue

        if re.fullmatch(r'[A-Z]\d+', token_upper):
            continue

        fields.append(token)

    return list(dict.fromkeys(fields))


def extract_print_sum_fields(section_text, defined_names):
    cleaned = re.sub(
        r'\bAS\b\s+(?:\'[^\']*\'|"[^"]*")',
        ' ',
        section_text,
        flags=re.IGNORECASE
    )
    cleaned = re.sub(r"'[^']*'|\"[^\"]*\"", ' ', cleaned)
    cleaned = re.sub(r'\b(?:NOPRINT|OVER|WITHIN|TOTAL|RECOMPUTE)\b', ' ', cleaned, flags=re.IGNORECASE)
    return extract_candidate_fields(cleaned, defined_names)


def classify_formula(formula):
    fc = normalize_expression(formula)
    upper = fc.upper()
    formula_types = []

    if extract_parameters(fc):
        formula_types.append('Parameter-based')
    if re.search(r'\bIF\b.+\bTHEN\b', upper):
        formula_types.append('Conditional IF')
    if re.search(r'\bDECODE\s*\(', upper) or re.search(r'\bDECODE\b', upper):
        formula_types.append('Decode/Mapping')
    if re.search(r'\b(?:EDIT|SUBSTR|TRIM|UPCASE|LOWCASE|LJUST|RJUST|CONCAT)\s*\(', upper) or '|' in fc:
        formula_types.append('Concatenation/String')
    if re.search(r'(?<![<>=])[-+*/](?![<>=])', fc):
        formula_types.append('Arithmetic')
    if re.search(r'\b(?:DATE|DMY|MDY|YMD|HDATE|TODAY|DATEDIF|DATEADD)\b', upper):
        formula_types.append('Date/Time')

    return ' + '.join(formula_types) if formula_types else 'Direct/Other'


def suggest_formula_rebuild(formula_type):
    if 'Decode/Mapping' in formula_type:
        return 'Use a Framework Manager/Data Module lookup (merged query subject) or a Cognos CASE calculation.'
    if 'Conditional IF' in formula_type:
        return 'Use a Cognos report or Framework Manager CASE/IF calculation.'
    if 'Concatenation/String' in formula_type:
        return 'Use Cognos string functions (concatenation operator, substring(), trim()) in a calculation.'
    if 'Arithmetic' in formula_type:
        return 'Use a Framework Manager calculated query item or a report-level Cognos calculation.'
    if 'Parameter-based' in formula_type:
        return 'Use a Cognos prompt (parameter) with a prompt page, or a macro.'
    if 'Date/Time' in formula_type:
        return 'Use Cognos date/time functions (e.g. _add_days, extract, _days_between) in a calculation.'
    return 'Review expression and rebuild as a Framework Manager or report-level Cognos calculation.'


def suggest_filter_rebuild(expr):
    if extract_parameters(expr):
        return 'Use a Cognos prompt (parameter), prompt page, or macro (#prompt()#).'
    return 'Use a Cognos detail/summary filter, or a Framework Manager/Data Module filter if it should always apply.'

def suggest_output_format(fmt):
    fmt = str(fmt or '').upper()
    if fmt in {'HTML', 'HTMTABLE', 'HTMLCSS'}:
        return 'Cognos report viewed in HTML (Cognos Viewer) or a portal tab.'
    if fmt in {'PDF'}:
        return 'Cognos report run or scheduled with PDF output.'
    if fmt in {'XLSX', 'EXL2K', 'EXCEL', 'CSV', 'ALPHA', 'DFIX'}:
        return 'Cognos report exported to Excel/CSV, or delivered as a Cognos Data Set.'
    if fmt in {'FOCUS', 'HOLD'}:
        return 'Framework Manager/Data Module staging query subject or table.'
    return 'Map to a Cognos report, export, or data module based on usage.'



def suggest_join_cognos_action(from_table, from_key, join_mode, to_table, to_key):
    from_ref = f"{from_table}.{from_key}" if from_table and from_key else "FROM_TABLE.FROM_KEY"
    to_ref = f"{to_table}.{to_key}" if to_table and to_key else "TO_TABLE.TO_KEY"

    action = (
        f"Create Framework Manager / Data Module relationship:\n{from_ref} -> {to_ref}\n\n"
        f"Alternative: use a Cognos Merged Query joining {from_table or 'FROM_TABLE'} with "
        f"{to_table or 'TO_TABLE'} on {from_key or 'FROM_KEY'} = {to_key or 'TO_KEY'} if you "
        f"cannot modify the Framework Manager/Data Module model."
    )

    if str(join_mode or '').upper() == 'JOIN TO ALL':
        action += (
            "\n\nWebFOCUS TO ALL suggests multi-row matching may exist. Review key uniqueness, "
            "cardinality, and query behavior before implementing in Framework Manager or the Merged Query."
        )

    return action


def suggest_join_review_reason(join_mode):
    if str(join_mode or '').upper() == 'JOIN TO ALL':
        return 'JOIN TO ALL can imply one-to-many or many-to-one behavior. Confirm key uniqueness, cardinality, and query behavior in Framework Manager/Cognos.'
    return 'Confirm whether this should be a Framework Manager relationship or a Cognos Merged Query, and verify key uniqueness/cardinality.'


def suggest_match_cognos_action(first_table, to_tables, keys, hold):
    table_text = ', '.join([t for t in [first_table, to_tables] if t])
    key_text = keys or 'matched key fields'
    hold_text = f" Output HOLD table: {hold}." if hold else ''

    return (
        f"MATCH FILE creates a combined/intermediate result from {table_text} using {key_text}."
        f"{hold_text}\n\n"
        "Rebuild as a Framework Manager staging query subject or a Cognos Merged Query. Review "
        "OLD/NEW/AFTER/MORE behavior manually before choosing merge vs. append (Union Query)."
    )


def classify_field_role(field_name, source_name_set, calculated_name_set):
    in_source = field_name in source_name_set
    in_calc = field_name in calculated_name_set

    if in_source and in_calc:
        return 'Both DB Source and Calculated'
    if in_source:
        return 'DB Source Only'
    if in_calc:
        return 'Calculated Only'

    return ''


def extract_hold_names(text):
    hold_names = set()

    for name in re.findall(
        r'ON\s+TABLE\s+HOLD\s+AS\s+([A-Za-z0-9_]\w*)',
        text,
        re.IGNORECASE
    ):
        hold_names.add(name.upper())

    for name in re.findall(
        r'ON\s+MATCH\s+HOLD\s+AS\s+([A-Za-z0-9_]\w*)',
        text,
        re.IGNORECASE
    ):
        hold_names.add(name.upper())

    hold_names.add('HOLD')

    return hold_names


def extract_file_definitions(text):
    file_defs = {}

    for match in re.finditer(r'^\s*FILEDEF\s+(\S+)\s+(.+?)\s*$', text, re.IGNORECASE | re.MULTILINE):
        name = match.group(1).strip().upper()
        statement = normalize_expression(match.group(0))
        file_defs[name] = statement

    return file_defs


def is_hold_like_table(table_name, explicit_hold_names=None):
    if not table_name:
        return False

    t = str(table_name).strip().upper()

    if not t:
        return False

    explicit_hold_names = explicit_hold_names or set()

    if t.startswith('&'):
        return True
    if t in explicit_hold_names:
        return True
    if t.startswith('HOLD'):
        return True
    if t.startswith('HLD'):
        return True
    if 'HOLD' in t:
        return True

    return False


def extract_filter_parameter_mapping(text, defined_names):
    rows = []
    filter_pattern = re.compile(
        r'^\s*(WHERE|IF)\s+(.+?)(?:;)?\s*$',
        re.IGNORECASE | re.MULTILINE
    )

    for match in filter_pattern.finditer(text):
        filter_type = match.group(1).upper()
        expr = normalize_expression(match.group(2))

        if not expr or expr.upper().startswith(('TOTAL ', 'RECORDLIMIT ')):
            continue
        if filter_type == 'IF' and re.search(r'\bTHEN\b|\bELSE\b', expr, re.IGNORECASE):
            continue

        rows.append({
            'type': filter_type,
            'expression': expr,
            'parameters': extract_parameters(expr),
            'fields': extract_candidate_fields(expr, defined_names),
            'cognos_guidance': suggest_filter_rebuild(expr),
        })

    parameter_rows = []
    seen_params = set()

    for param in extract_parameters(text):
        if param in seen_params:
            continue
        seen_params.add(param)

        parameter_rows.append({
            'type': 'PARAMETER',
            'expression': param,
            'parameters': [param],
            'fields': [],
            'cognos_guidance': 'Create a Cognos prompt (parameter) or macro.',
        })

    return rows + parameter_rows


def extract_output_formats(text):
    rows = []

    patterns = [
        (
            'PCHOLD',
            re.compile(
                r'\bON\s+TABLE\s+PCHOLD(?:\s+AS\s+([A-Za-z0-9_][\w-]*))?\s+FORMAT\s+([A-Za-z0-9_]+)',
                re.IGNORECASE
            )
        ),
        (
            'HOLD',
            re.compile(
                r'\bON\s+TABLE\s+HOLD(?:\s+AS\s+([A-Za-z0-9_][\w-]*))?(?:\s+FORMAT\s+([A-Za-z0-9_]+))?',
                re.IGNORECASE
            )
        ),
        (
            'SET ONLINE-FMT',
            re.compile(
                r'\bSET\s+ONLINE-FMT\s*=\s*([A-Za-z0-9_]+)',
                re.IGNORECASE
            )
        ),
    ]

    for command, pattern in patterns:
        for match in pattern.finditer(text):
            if command == 'SET ONLINE-FMT':
                output_name = ''
                fmt = match.group(1)
            else:
                output_name = match.group(1) or ''
                fmt = match.group(2) or ('FOCUS' if command == 'HOLD' else '')

            fmt = fmt.upper()
            rows.append({
                'command': command,
                'output_name': output_name,
                'format': fmt,
                'cognos_guidance': suggest_output_format(fmt),
            })

    for fmt_match in re.finditer(r'\bFORMAT\s+([A-Za-z0-9_]+)', text, re.IGNORECASE):
        fmt_context = text[max(0, fmt_match.start() - 40):fmt_match.start()].upper()

        if 'PCHOLD' in fmt_context or ' HOLD' in fmt_context or '\nHOLD' in fmt_context:
            continue

        fmt = fmt_match.group(1)
        fmt = fmt.upper()
        if fmt in {'HTML', 'HTMLCSS', 'PDF', 'XLSX', 'EXL2K', 'EXCEL', 'FOCUS', 'ALPHA', 'CSV'}:
            rows.append({
                'command': 'FORMAT',
                'output_name': '',
                'format': fmt,
                'cognos_guidance': suggest_output_format(fmt),
            })

    deduped = []
    seen = set()

    for row in rows:
        key = (row['command'], row['output_name'], row['format'])
        if key not in seen:
            seen.add(key)
            deduped.append(row)

    return deduped


def extract_join_match_mapping(text):
    rows = []

    join_pattern = re.compile(
        r'^\s*JOIN\s+(.+?)\s+IN\s+(\S+)\s+TO\s+(ALL\s+)?(.+?)\s+IN\s+(\S+)(.*)$',
        re.IGNORECASE | re.MULTILINE
    )

    for match in join_pattern.finditer(text):
        from_key = normalize_expression(match.group(1))
        from_table = match.group(2)
        join_type = 'JOIN TO ALL' if match.group(3) else 'JOIN'
        to_key = normalize_expression(match.group(4))
        to_table = match.group(5)
        raw_stmt = normalize_expression(match.group(0))

        rows.append({
            'mapping_type': 'JOIN',
            'from_table': from_table,
            'from_key': from_key,
            'join_type': join_type,
            'to_table': to_table,
            'to_key': to_key,
            'output_hold': '',
            'raw_statement': raw_stmt,
            'cognos_guidance': 'Use a Framework Manager relationship or Cognos Merged Query; TO ALL usually maps to one-to-many/many-to-one relationship behavior.',
            'suggested_cognos_type': 'Framework Manager relationship / Cognos Merged Query',
            'suggested_cognos_expression_action': suggest_join_cognos_action(
                from_table,
                from_key,
                join_type,
                to_table,
                to_key,
            ),
            'needs_manual_review': 'Yes',
            'review_reason': suggest_join_review_reason(join_type),
            'confidence': 'High',
        })

    match_block_pattern = re.compile(
        r'\bMATCH\s+FILE\s+(\S+)(.*?)(?=\nEND\b|\nMATCH\s+FILE\b|\Z)',
        re.IGNORECASE | re.DOTALL
    )

    for match in match_block_pattern.finditer(text):
        first_table = match.group(1)
        block = match.group(2)
        all_tables = [first_table] + re.findall(r'\bFILE\s+(\S+)', block, re.IGNORECASE)
        keys = re.findall(r'\bBY\s+([A-Za-z_][\w.]*)', block, re.IGNORECASE)
        hold = ''
        hold_match = re.search(r'\b(?:HOLD|PCHOLD)\s+AS\s+([A-Za-z0-9_][\w-]*)', block, re.IGNORECASE)

        if hold_match:
            hold = hold_match.group(1)

        rows.append({
            'mapping_type': 'MATCH FILE',
            'from_table': first_table,
            'from_key': ', '.join(keys),
            'join_type': 'MATCH',
            'to_table': ', '.join(dict.fromkeys(all_tables[1:])),
            'to_key': ', '.join(keys),
            'output_hold': hold,
            'raw_statement': normalize_expression(match.group(0)),
            'cognos_guidance': 'Use a Cognos Merged Query/Union or Framework Manager relationships depending on MATCH logic and AFTER/MORE usage.',
            'suggested_cognos_type': 'Framework Manager staging query subject / Cognos Merged Query',
            'suggested_cognos_expression_action': suggest_match_cognos_action(
                first_table,
                ', '.join(dict.fromkeys(all_tables[1:])),
                ', '.join(keys),
                hold,
            ),
            'needs_manual_review': 'Yes',
            'review_reason': 'MATCH FILE behavior depends on full OLD/NEW/AFTER/MORE logic and should be manually rebuilt as merge, append, or staging logic.',
            'confidence': 'Medium',
        })

    return rows




def extract_drilldowns(text):
    rows = []
    seen = set()
    patterns = [
        re.compile(r'\bURL\s*=\s*([^\s,;]+)', re.IGNORECASE),
        re.compile(r'\bIBIF_ex\s*=\s*([A-Za-z0-9_.$#/-]+)', re.IGNORECASE),
        re.compile(r'\bFOCEXURL\s*=?\s*([^\s,;]+)?', re.IGNORECASE),
    ]

    for pattern in patterns:
        for match in pattern.finditer(text):
            raw = normalize_expression(match.group(0))
            target = ''
            target_match = re.search(r'\bIBIF_ex\s*=\s*([A-Za-z0-9_.$#/-]+)', raw, re.IGNORECASE)
            if target_match:
                target = target_match.group(1)
            elif match.lastindex:
                value = match.group(1) or ''
                if value.lower().endswith('.fex'):
                    target = value

            context = text[max(0, match.start() - 160): min(len(text), match.end() + 220)]
            params = extract_parameters(context)
            fields = extract_candidate_fields(context, set())
            key = (raw, target, tuple(params), tuple(fields))
            if key in seen:
                continue
            seen.add(key)

            rows.append({
                'raw_drilldown': raw,
                'target_report': target,
                'parameter': ', '.join(params),
                'fields': ', '.join(fields[:8]),
                'cognos_replacement': 'Cognos drill-through definition or report link',
                'review': 'Confirm target report, passed parameter fields, and expected navigation behavior.',
            })

    return rows


def source_type_from_reference(source_name, raw_reference):
    combined = f"{source_name} {raw_reference}".upper()

    if '.CSV' in combined:
        return 'CSV'
    if any(ext in combined for ext in ['.XLS', '.XLSX']):
        return 'Excel'
    if any(ext in combined for ext in ['.TXT', '.DAT', '.FTM']):
        return 'Flat file'
    if 'DISK' in combined:
        return 'File'
    if source_name and '.' in source_name:
        return 'Database / synonym'
    return 'WebFOCUS synonym / table'


def suggested_stage_name(prefix, source_name):
    clean = normalize_program_name(source_name).title().replace('_', '')
    return f"{prefix}_{clean}" if clean else f"{prefix}_Table"


def extract_source_inventory(parsed):
    rows = []
    file_definitions = parsed.get('file_definitions', {})
    sources = list(dict.fromkeys(parsed.get('real_sources', []) + parsed.get('sources', [])))

    for source in sources:
        raw_reference = file_definitions.get(str(source).upper(), '')
        source_type = source_type_from_reference(source, raw_reference)
        layer = suggested_stage_name('stg', source)

        if source_type in {'CSV', 'Excel', 'Flat file', 'File'}:
            action = f"Create a Framework Manager/Data Module source query subject {layer} from the FILEDEF path."
        elif source_type == 'Database / synonym':
            action = f"Create a Framework Manager/Data Module or SQL source query subject {layer} from the database/synonym reference."
        else:
            action = f"Create a Framework Manager staging query subject {layer}; confirm the synonym's physical source."

        if not raw_reference:
            action += " FILEDEF was not found, so validate the physical source manually."

        rows.append({
            'Source Name': source,
            'Detected Source Type': source_type,
            'FILEDEF / Raw Source Reference': raw_reference,
            'Suggested Cognos Layer': layer,
            'Suggested Cognos Action': action,
            'Notes': 'Intermediate HOLD table' if is_hold_like_table(source, parsed.get('hold_names')) else 'Original report source',
        })

    return rows


def extract_table_blocks(text):
    pattern = re.compile(
        r'\b(TABLE|GRAPH)\s+FILE\s+(\S+)\s*(.*?)(?=\nEND\b|\nTABLE\s+FILE\b|\nGRAPH\s+FILE\b|\Z)',
        re.IGNORECASE | re.DOTALL
    )
    return list(pattern.finditer(text))


def summarize_table_block(block):
    parts = []
    upper = block.upper()

    if re.search(r'\bSUM\b', upper):
        parts.append('aggregates measures')
    if re.search(r'\bPRINT\b', upper):
        parts.append('selects/detail rows')
    if re.search(r'\bCOMPUTE\b', upper):
        parts.append('creates computed fields')
    if re.search(r'\bWHERE\b|\bIF\b', upper):
        parts.append('filters rows')
    if re.search(r'\bACROSS\b', upper):
        parts.append('pivots categories with ACROSS')
    if re.search(r'\bHOLD\b|\bPCHOLD\b', upper):
        parts.append('writes an output/HOLD table')

    return '; '.join(parts) if parts else 'reads source table'


def extract_key_fields(block):
    keys = []
    for label, pattern in [
        ('BY', r'\bBY\s+([A-Za-z_][\w.]*)'),
        ('ACROSS', r'\bACROSS\s+([A-Za-z_][\w.]*)'),
    ]:
        for value in re.findall(pattern, block, re.IGNORECASE):
            if value.upper() not in {'TABLE', 'ON', 'END'}:
                keys.append(f"{label} {value}")
    return ', '.join(dict.fromkeys(keys))


def extract_block_formulas(block):
    formulas = []
    for field, formula in re.findall(
        r'\bCOMPUTE\s+([A-Za-z_]\w*)\s*/\s*[A-Za-z0-9%.]+\s*=\s*(.*?);',
        block,
        re.IGNORECASE | re.DOTALL
    ):
        formulas.append(f"{field} = {normalize_expression(formula)}")
    return '\n'.join(formulas)


def extract_block_filters(block):
    filters = []
    for kind, expr in re.findall(r'^\s*(WHERE|IF)\s+(.+?)(?:;)?\s*$', block, re.IGNORECASE | re.MULTILINE):
        expr = normalize_expression(expr)
        if expr and not expr.upper().startswith(('TOTAL ', 'RECORDLIMIT ')):
            filters.append(f"{kind.upper()} {expr}")
    return '\n'.join(filters)


def extract_block_measures(block, defined_names):
    ss = re.search(
        r'(?:SUM|PRINT)\b(.*?)(?=\nBY\b|\nWHERE\b|\nON\s+TABLE\b|\Z)',
        block,
        re.IGNORECASE | re.DOTALL
    )
    if not ss:
        return ''
    return ', '.join(extract_print_sum_fields(ss.group(1), defined_names))


def extract_hold_lineage(text, parsed):
    rows = []
    clean_text = strip_comments(text)
    defined_names = {f['field'].upper() for f in parsed.get('define_fields', [])}
    step_no = 1

    for join in parsed.get('join_match_mapping', []):
        if join['mapping_type'] != 'JOIN':
            continue

        rows.append({
            'Step #': step_no,
            'Step Type': 'JOIN',
            'Input Table(s)': f"{join['from_table']}, {join['to_table']}",
            'What Happens In This Step': f"Joins {join['from_table']}[{join['from_key']}] to {join['to_table']}[{join['to_key']}].",
            'Output Table': join.get('output_hold', ''),
            'Output Role': 'Join context',
            'Key Fields / BY / ACROSS': f"{join['from_key']} = {join['to_key']}",
            'Formulas Created': '',
            'Filters Applied': '',
            'Cognos Table Suggestion': '',
            'Cognos Action': join.get('suggested_cognos_expression_action', ''),
            'Needs Manual Review': join.get('needs_manual_review', 'Yes'),
        })
        step_no += 1

    for match in extract_table_blocks(clean_text):
        command = match.group(1).upper()
        source = match.group(2)
        block = match.group(3)
        hold_match = re.search(
            r'\bON\s+TABLE\s+(?:PCHOLD|HOLD)(?:\s+AS\s+([A-Za-z0-9_][\w-]*))?(?:\s+FORMAT\s+([A-Za-z0-9_]+))?',
            block,
            re.IGNORECASE
        )
        output = ''
        output_role = 'Final Report Dataset'

        if hold_match:
            output = hold_match.group(1) or 'HOLD'
            output_role = 'Intermediate HOLD' if hold_match.group(1) else 'Default HOLD'

        measures = extract_block_measures(block, defined_names)
        keys = extract_key_fields(block)
        formulas = extract_block_formulas(block)
        filters = extract_block_filters(block)
        table_prefix = 'final' if output_role == 'Final Report Dataset' else 'stg'
        cognos_table = suggested_stage_name(table_prefix, output or source)
        action_parts = []

        if output_role in {'Intermediate HOLD', 'Default HOLD'}:
            action_parts.append('Create this as a Framework Manager staging query subject / Data Module step.')
        else:
            action_parts.append('Use this as the final query subject or report dataset.')
        if measures:
            action_parts.append(f"Rebuild SUM/PRINT fields: {measures}.")
        if keys:
            action_parts.append(f"Keep grouping/category fields: {keys}.")
        if filters:
            action_parts.append('Apply row filters before load or as report filters/prompts.')
        if 'ACROSS' in keys.upper():
            action_parts.append('For stacked visuals, unpivot ACROSS/category columns into a long table.')

        rows.append({
            'Step #': step_no,
            'Step Type': command,
            'Input Table(s)': source,
            'What Happens In This Step': summarize_table_block(block),
            'Output Table': output,
            'Output Role': output_role,
            'Key Fields / BY / ACROSS': keys,
            'Formulas Created': formulas,
            'Filters Applied': filters,
            'Cognos Table Suggestion': cognos_table,
            'Cognos Action': ' '.join(action_parts),
            'Needs Manual Review': 'Yes',
        })
        step_no += 1

    for match_row in parsed.get('join_match_mapping', []):
        if match_row['mapping_type'] != 'MATCH FILE':
            continue

        rows.append({
            'Step #': step_no,
            'Step Type': 'MATCH FILE',
            'Input Table(s)': ', '.join(v for v in [match_row['from_table'], match_row['to_table']] if v),
            'What Happens In This Step': 'MATCH FILE creates an intermediate combined result.',
            'Output Table': match_row.get('output_hold', ''),
            'Output Role': 'Intermediate HOLD' if match_row.get('output_hold') else 'Combined result',
            'Key Fields / BY / ACROSS': match_row.get('from_key', ''),
            'Formulas Created': '',
            'Filters Applied': '',
            'Cognos Table Suggestion': suggested_stage_name('stg', match_row.get('output_hold') or match_row.get('from_table')),
            'Cognos Action': match_row.get('suggested_cognos_expression_action', ''),
            'Needs Manual Review': match_row.get('needs_manual_review', 'Yes'),
        })
        step_no += 1

    return rows


def decode_pairs_from_formula(formula):
    text_value = normalize_expression(formula)
    match = re.search(r'\bDECODE\s+([A-Za-z_][\w.]*)\s*\((.*?)\)', text_value, re.IGNORECASE)
    if not match:
        match = re.search(r'\bDECODE\s+([A-Za-z_][\w.]*)\s+(.*)', text_value, re.IGNORECASE)
    if not match:
        return '', []

    source_field = match.group(1)
    tokens = re.findall(r"'[^']*'|\"[^\"]*\"|\S+", match.group(2).strip())
    cleaned = [t.strip().strip("'\"") for t in tokens if t.strip()]
    pairs = []

    for idx in range(0, len(cleaned) - 1, 2):
        pairs.append((cleaned[idx], cleaned[idx + 1]))

    return source_field, pairs


def extract_mapping_table_preview(parsed):
    rows = []

    for item in parsed.get('formula_mapping', []):
        if 'Decode/Mapping' not in item.get('formula_type', ''):
            continue

        source_field, pairs = decode_pairs_from_formula(item.get('formula', ''))
        if not source_field or not pairs:
            continue

        table_name = f"map_{source_field}_{item.get('field', 'Value')}"
        for source_value, mapped_value in pairs:
            rows.append({
                'Mapping Table Name': table_name,
                'Source Field': source_field,
                'Source Value': source_value,
                'Mapped Value': mapped_value,
                'Derived Field': item.get('field', ''),
                'Suggested Cognos Action': f"Create {table_name} as a Framework Manager/Data Module lookup table and relate it on {source_field}.",
            })

    return rows


def recommend_final_dataset(parsed):
    final_steps = [
        row for row in parsed.get('hold_lineage', [])
        if row.get('Output Role') == 'Final Report Dataset'
    ]
    last_step = final_steps[-1] if final_steps else (parsed.get('hold_lineage') or [{}])[-1]
    measures = ', '.join(dict.fromkeys(
        f['field'] for f in parsed.get('sum_real', []) + parsed.get('sum_calc', [])
    ))
    categories = ', '.join(dict.fromkeys(
        f['field'] for f in parsed.get('by_real', []) + parsed.get('by_calc', [])
    ))
    has_across = any('ACROSS' in row.get('Key Fields / BY / ACROSS', '').upper() for row in parsed.get('hold_lineage', []))
    has_graph = any(row.get('Step Type') == 'GRAPH' for row in parsed.get('hold_lineage', []))

    if has_across or len([m for m in measures.split(', ') if m]) > 1:
        shape = 'Long / unpivoted'
        reason = 'WebFOCUS uses multiple measures or ACROSS-style categories. Cognos crosstabs/charts usually work best when category/type is stored as rows.'
    else:
        shape = 'Wide or standard fact table'
        reason = 'Detected output can likely be rebuilt as a normal fact/dimension model table without forced unpivoting.'

    flow = ' -> '.join(
        row.get('Output Table') or row.get('Input Table(s)', '')
        for row in parsed.get('hold_lineage', [])
        if row.get('Input Table(s)') or row.get('Output Table')
    )

    return [{
        'Recommended Final Dataset': last_step.get('Cognos Table Suggestion') or 'final_ReportDataset',
        'Dataset Shape': shape,
        'Reason': reason,
        'Suggested Cognos Visual / Output': 'Cognos chart (Graph) visual' if has_graph else output_type_text(parsed) or 'Cognos list/crosstab report',
        'Source Tables / HOLD Flow': flow,
        'Measures / Value Columns': measures,
        'Legend / Category Fields': categories,
        'Needs Manual Review': 'Yes',
    }]


def analyze_sql_view_need(parsed):
    source_count = len(parsed.get('real_sources', []))
    join_rows = parsed.get('join_match_mapping', [])
    lineage_rows = parsed.get('hold_lineage', [])
    formula_rows = parsed.get('formula_mapping', [])
    mapping_rows = parsed.get('mapping_table_preview', [])
    output_formats = output_type_text(parsed).upper()

    join_count = len(join_rows)
    hold_count = sum(1 for row in lineage_rows if 'HOLD' in str(row.get('Output Role', '')).upper())
    match_count = sum(1 for row in join_rows if row.get('mapping_type') == 'MATCH FILE')
    join_to_all_count = sum(1 for row in join_rows if str(row.get('join_type', '')).upper() == 'JOIN TO ALL')
    complex_formula_count = sum(
        1 for row in formula_rows
        if any(label in row.get('formula_type', '') for label in ['Decode/Mapping', 'Date/Time', 'Parameter-based'])
    )

    score = (
        join_count * 2
        + hold_count * 2
        + match_count * 3
        + join_to_all_count * 3
        + len(mapping_rows)
        + complex_formula_count
    )

    if source_count >= 3 or match_count or join_to_all_count or join_count >= 2 or hold_count >= 3 or score >= 8:
        need = 'Recommended'
        reason = 'Complex JOIN/HOLD/MATCH logic is easier to validate and reuse as a SQL view before modeling it in Cognos.'
        path = 'SQL View -> Framework Manager / Data Module -> Calculations -> Report'
        action = 'Create a SQL view for joins/HOLD logic, then model it in Framework Manager/Data Module and finish shaping/reporting in Cognos.'
    elif join_count == 1 or hold_count in {1, 2} or complex_formula_count >= 2:
        need = 'Optional'
        reason = 'The report has moderate shaping logic. Framework Manager/Data Module can handle it, but a SQL view may be cleaner if reused.'
        path = 'Framework Manager / Data Module -> Calculations -> Report'
        action = 'Model directly in Framework Manager/Data Module first; consider a SQL view if performance, reuse, or governance is important.'
    else:
        need = 'Not Needed'
        reason = 'The report appears straightforward enough for Framework Manager/Data Module and the Cognos report layer.'
        path = 'Framework Manager / Data Module -> Report'
        action = 'Load source, apply filters/calculations in Framework Manager/Data Module, then build the report.'

    if 'GRAPH' in output_formats and 'Report' not in path:
        path = f"{path} -> Report"

    return {
        'need': need,
        'reason': reason,
        'build_path': path,
        'action': action,
        'score': score,
    }


def infer_report_type(parsed):
    output_text = output_type_text(parsed).upper()
    has_graph = any(row.get('Step Type') == 'GRAPH' for row in parsed.get('hold_lineage', []))
    has_across = any('ACROSS' in row.get('Key Fields / BY / ACROSS', '').upper() for row in parsed.get('hold_lineage', []))
    join_count = len(parsed.get('join_match_mapping', []))
    hold_count = sum(1 for row in parsed.get('hold_lineage', []) if 'HOLD' in str(row.get('Output Role', '')).upper())

    if has_graph or 'GRAPH' in output_text:
        return 'Dashboard / chart'
    if has_across:
        return 'Matrix / pivot report'
    if join_count >= 2 or hold_count >= 2:
        return 'Complex multi-step report'
    if parsed.get('sum_real') or parsed.get('sum_calc'):
        return 'Summary report'
    return 'Detail table'


def infer_report_purpose(fex_name, parsed):
    base = Path(str(fex_name)).stem.replace('_', ' ').replace('-', ' ').strip()
    report_type = infer_report_type(parsed).lower()
    sources = ', '.join(parsed.get('real_sources', [])[:3])

    if sources:
        return f"{base} {report_type} using {sources}"
    return f"{base} {report_type}"


def build_cognos_plan(parsed):
    rows = []
    step = 1
    sql_view = analyze_sql_view_need(parsed)

    if sql_view['need'] in {'Recommended', 'Optional'}:
        rows.append({
            'Plan Step': step,
            'Cognos Layer': 'SQL View',
            'WebFOCUS Object': 'JOIN/HOLD/MATCH logic',
            'What To Build': sql_view['action'],
            'Why': sql_view['reason'],
            'Manual Review': 'Yes',
        })
        step += 1

    for source in parsed.get('source_inventory', []):
        rows.append({
            'Plan Step': step,
            'Cognos Layer': source['Suggested Cognos Layer'],
            'WebFOCUS Object': source['Source Name'],
            'What To Build': source['Suggested Cognos Action'],
            'Why': 'Start with source/staging query subjects before rebuilding HOLD outputs.',
            'Manual Review': 'Yes' if not source['FILEDEF / Raw Source Reference'] else 'No',
        })
        step += 1

    for lineage in parsed.get('hold_lineage', []):
        rows.append({
            'Plan Step': step,
            'Cognos Layer': lineage['Cognos Table Suggestion'],
            'WebFOCUS Object': lineage['Output Table'] or lineage['Input Table(s)'],
            'What To Build': lineage['Cognos Action'],
            'Why': lineage['What Happens In This Step'],
            'Manual Review': lineage['Needs Manual Review'],
        })
        step += 1

    for mapping in parsed.get('mapping_table_preview', []):
        rows.append({
            'Plan Step': step,
            'Cognos Layer': mapping['Mapping Table Name'],
            'WebFOCUS Object': mapping['Derived Field'],
            'What To Build': mapping['Suggested Cognos Action'],
            'Why': 'DECODE logic is easier to maintain as a Framework Manager lookup/mapping table.',
            'Manual Review': 'Yes',
        })
        step += 1

    for final_dataset in parsed.get('final_dataset_plan', []):
        rows.append({
            'Plan Step': step,
            'Cognos Layer': final_dataset['Recommended Final Dataset'],
            'WebFOCUS Object': 'Final output',
            'What To Build': f"Build {final_dataset['Dataset Shape']} dataset for {final_dataset['Suggested Cognos Visual / Output']}.",
            'Why': final_dataset['Reason'],
            'Manual Review': final_dataset['Needs Manual Review'],
        })

    return rows

def parse_fex(fex_text):
    text = strip_comments(fex_text)
    explicit_hold_names = extract_hold_names(text)

    result = {
        'sources': [],
        'file_definitions': extract_file_definitions(text),
        'define_fields': [],
        'compute_fields': [],
        'sum_real': [],
        'sum_calc': [],
        'by_real': [],
        'by_calc': [],
        'source_fields': [],
        'calculated_counts': {},
        'source_name_set': set(),
        'calculated_name_set': set(),
        'real_sources': [],
        'hold_names': explicit_hold_names,
        'filter_parameter_mapping': [],
        'formula_mapping': [],
        'output_formats': [],
        'join_match_mapping': [],
        'source_inventory': [],
        'hold_lineage': [],
        'mapping_table_preview': [],
        'final_dataset_plan': [],
        'cognos_build_plan': [],
        'drilldowns': [],
    }

    table_src = re.findall(r'TABLE\s+FILE\s+(\S+)', text, re.IGNORECASE)
    define_src = re.findall(r'DEFINE\s+FILE\s+(\S+)', text, re.IGNORECASE)

    result['sources'] = list(dict.fromkeys(table_src + define_src))

    result['real_sources'] = [
        s for s in result['sources']
        if not is_hold_like_table(s, explicit_hold_names)
    ]

    for def_src, block in re.findall(
        r'DEFINE\s+FILE\s+(\S+)\s*(.*?)END',
        text,
        re.IGNORECASE | re.DOTALL
    ):
        for fname, fmt, formula in re.findall(
            r'([A-Za-z0-9_]\w*)\s*/\s*([A-Za-z0-9%.]+)\s*=\s*(.*?);',
            block,
            re.DOTALL
        ):
            result['define_fields'].append({
                'field': fname,
                'format': fmt,
                'formula': ' '.join(formula.split()),
                'source': def_src,
            })

    defined_names = {f['field'].upper() for f in result['define_fields']}
    raw_set = set()

    for f in result['define_fields']:
        f['raw_fields'] = raw_db_fields(f['formula'], defined_names)
        formula_type = classify_formula(f['formula'])

        result['formula_mapping'].append({
            'formula_source': 'DEFINE',
            'field': f['field'],
            'format': f['format'],
            'formula_type': formula_type,
            'formula': f['formula'],
            'raw_fields': f['raw_fields'],
            'parameters': extract_parameters(f['formula']),
            'cognos_guidance': suggest_formula_rebuild(formula_type),
        })

        for r in f['raw_fields']:
            raw_set.add((r, f['source']))

    for tbl_src, block in re.findall(
        r'TABLE\s+FILE\s+(\S+)\s*(.*?)(?=\nEND\b|\Z)',
        text,
        re.IGNORECASE | re.DOTALL
    ):
        for fname, fmt, formula, alias in re.findall(
            r'COMPUTE\s+([A-Za-z0-9_]\w*)\s*/\s*([A-Za-z0-9%.]+)\s*=\s*(.*?);'
            r'(?:\s*AS\s*[\'"]([^\'"]*)[\'"])?',
            block,
            re.IGNORECASE | re.DOTALL
        ):
            fc = ' '.join(formula.split())
            raws = raw_db_fields(fc, defined_names)

            result['compute_fields'].append({
                'field': fname,
                'format': fmt,
                'formula': fc,
                'alias': alias or '',
                'source': tbl_src,
                'raw_fields': raws,
            })

            formula_type = classify_formula(fc)

            result['formula_mapping'].append({
                'formula_source': 'COMPUTE',
                'field': fname,
                'format': fmt,
                'formula_type': formula_type,
                'formula': fc,
                'raw_fields': raws,
                'parameters': extract_parameters(fc),
                'cognos_guidance': suggest_formula_rebuild(formula_type),
            })

            for r in raws:
                raw_set.add((r, tbl_src))

        ss = re.search(
            r'(?:SUM|PRINT)\b(.*?)(?=\nBY\b|\nWHERE\b|\nON\s+TABLE\b|\Z)',
            block,
            re.IGNORECASE | re.DOTALL
        )

        if ss:
            for fn in extract_print_sum_fields(ss.group(1), defined_names):
                if fn.upper() in defined_names:
                    result['sum_calc'].append({'field': fn, 'source': tbl_src})
                else:
                    result['sum_real'].append({'field': fn, 'source': tbl_src})
                    raw_set.add((fn, tbl_src))

        for m in re.finditer(r'\bBY\s+([A-Za-z_]\w*)', block, re.IGNORECASE):
            fn = m.group(1)

            if fn.upper() in ('TABLE', 'ON', 'END'):
                continue

            if fn.upper() in defined_names:
                result['by_calc'].append({'field': fn, 'source': tbl_src})
            else:
                result['by_real'].append({'field': fn, 'source': tbl_src})
                raw_set.add((fn, tbl_src))

    seen = set()

    for fn, src in sorted(raw_set):
        if (fn, src) not in seen:
            seen.add((fn, src))
            result['source_fields'].append({'field': fn, 'source': src})

    calc_names = (
        [f['field'] for f in result['define_fields']]
        + [f['field'] for f in result['compute_fields']]
    )

    result['calculated_counts'] = dict(Counter(calc_names))
    result['calculated_name_set'] = set(calc_names)
    result['source_name_set'] = {f['field'] for f in result['source_fields']}
    all_defined_names = {name.upper() for name in calc_names}
    result['filter_parameter_mapping'] = extract_filter_parameter_mapping(text, all_defined_names)
    result['output_formats'] = extract_output_formats(text)
    result['join_match_mapping'] = extract_join_match_mapping(text)
    result['source_inventory'] = extract_source_inventory(result)
    result['hold_lineage'] = extract_hold_lineage(text, result)
    result['mapping_table_preview'] = extract_mapping_table_preview(result)
    result['final_dataset_plan'] = recommend_final_dataset(result)
    result['cognos_build_plan'] = build_cognos_plan(result)
    result['drilldowns'] = extract_drilldowns(text)

    return result


def _upper_set(values):
    return frozenset(
        normalize_expression(v).upper()
        for v in values
        if normalize_expression(v)
    )


def _formula_signature(parsed):
    return frozenset(
        (
            item['formula_source'].upper(),
            item['field'].upper(),
            normalize_expression(item['formula']).upper(),
            item.get('format', '').upper(),
        )
        for item in parsed['formula_mapping']
    )


def _filter_signature(parsed):
    return frozenset(
        (
            item['type'].upper(),
            normalize_expression(item['expression']).upper(),
        )
        for item in parsed['filter_parameter_mapping']
        if item['type'].upper() in {'WHERE', 'IF'}
    )


def _join_signature(parsed):
    return frozenset(
        (
            item['mapping_type'].upper(),
            item['from_table'].upper(),
            normalize_expression(item['from_key']).upper(),
            item['join_type'].upper(),
            item['to_table'].upper(),
            normalize_expression(item['to_key']).upper(),
            item.get('output_hold', '').upper(),
        )
        for item in parsed['join_match_mapping']
    )


def _output_signature(parsed):
    return frozenset(
        (
            item['command'].upper(),
            item['output_name'].upper(),
            item['format'].upper(),
        )
        for item in parsed['output_formats']
    )


def compute_fex_profile(parsed):
    real_tables = _upper_set(parsed['real_sources'])
    source_fields = _upper_set(f['field'] for f in parsed['source_fields'])
    shown_fields = _upper_set(
        [f['field'] for f in parsed['sum_real']]
        + [f['field'] for f in parsed['sum_calc']]
        + [f['field'] for f in parsed['by_real']]
        + [f['field'] for f in parsed['by_calc']]
    )
    formula_fields = _upper_set(
        [f['field'] for f in parsed['define_fields']]
        + [f['field'] for f in parsed['compute_fields']]
    )
    all_fields = source_fields | shown_fields | formula_fields

    if not real_tables and not all_fields:
        return None

    return {
        'real_tables': real_tables,
        'source_fields': source_fields,
        'shown_fields': shown_fields,
        'formula_fields': formula_fields,
        'all_fields': all_fields,
        'formulas': _formula_signature(parsed),
        'filters': _filter_signature(parsed),
        'joins': _join_signature(parsed),
        'outputs': _output_signature(parsed),
    }


def exact_duplicate_key(profile):
    if profile is None:
        return None

    return (
        profile['real_tables'],
        profile['source_fields'],
        profile['shown_fields'],
        profile['formulas'],
        profile['filters'],
        profile['joins'],
        profile['outputs'],
    )


def field_similarity(left_fields, right_fields):
    if not left_fields and not right_fields:
        return 1.0
    if not left_fields or not right_fields:
        return 0.0

    return len(left_fields & right_fields) / len(left_fields | right_fields)


def profile_difference_summary(base, other):
    differences = []

    for label, key in [
        ('fields', 'all_fields'),
        ('filters', 'filters'),
        ('formulas', 'formulas'),
        ('joins', 'joins'),
        ('outputs', 'outputs'),
    ]:
        if base[key] != other[key]:
            differences.append(label)

    return ', '.join(differences) if differences else 'No meaningful differences'


def build_duplicate_analysis(parsed_results):
    records = []
    exact_groups = defaultdict(list)
    table_groups = defaultdict(list)

    for folder, fex_name, parsed in parsed_results:
        profile = compute_fex_profile(parsed) if parsed is not None else None
        record = {
            'folder': folder,
            'fex_name': fex_name,
            'parsed': parsed,
            'profile': profile,
            'category': 'Unparsed' if profile is None else 'Unique',
            'group_id': '',
            'group_size': 1,
            'exact_matches': [],
            'near_matches': [],
            'same_source_matches': [],
            'best_similarity': 0.0,
            'difference_summary': '',
        }
        records.append(record)

        if profile is not None:
            exact_groups[exact_duplicate_key(profile)].append(record)
            table_groups[profile['real_tables']].append(record)

    group_counter = 1

    for members in exact_groups.values():
        if len(members) <= 1:
            continue

        group_id = f'Exact {group_counter}'
        group_counter += 1

        names = [m['fex_name'] for m in members]
        for record in members:
            record['category'] = 'Exact Duplicate'
            record['group_id'] = group_id
            record['group_size'] = len(members)
            record['exact_matches'] = [name for name in names if name != record['fex_name']]
            record['best_similarity'] = 1.0
            record['difference_summary'] = 'Exact same source tables, fields, formulas, filters, joins, and outputs'

    for table_key, members in table_groups.items():
        if not table_key or len(members) <= 1:
            continue

        for record in members:
            if record['category'] == 'Exact Duplicate':
                continue

            best_near = []
            same_source = []
            best_similarity = 0.0
            difference_notes = []

            for other in members:
                if other is record:
                    continue

                sim = field_similarity(record['profile']['all_fields'], other['profile']['all_fields'])
                best_similarity = max(best_similarity, sim)

                if sim >= 0.80:
                    if (
                        record['profile']['filters'] != other['profile']['filters']
                        or record['profile']['formulas'] != other['profile']['formulas']
                        or record['profile']['joins'] != other['profile']['joins']
                        or record['profile']['outputs'] != other['profile']['outputs']
                    ):
                        best_near.append(other['fex_name'])
                        difference_notes.append(profile_difference_summary(record['profile'], other['profile']))
                    else:
                        same_source.append(other['fex_name'])
                else:
                    same_source.append(other['fex_name'])

            if best_near:
                record['category'] = 'Near Duplicate'
                record['near_matches'] = sorted(set(best_near))
                record['group_id'] = f'Near Source {abs(hash(table_key)) % 100000}'
                record['group_size'] = len(best_near) + 1
                record['difference_summary'] = '; '.join(sorted(set(difference_notes))) or '80%+ same fields with logic differences'
            elif same_source:
                record['category'] = 'Same Data Source'
                record['same_source_matches'] = sorted(set(same_source))
                record['group_id'] = f'Source {abs(hash(table_key)) % 100000}'
                record['group_size'] = len(same_source) + 1
                record['difference_summary'] = 'Same real source tables with different fields or report logic'

            record['best_similarity'] = max(record['best_similarity'], best_similarity)

    for record in records:
        if record['category'] == 'Unique':
            record['difference_summary'] = 'No meaningful match found'
        elif record['category'] == 'Unparsed':
            record['difference_summary'] = 'Could not parse enough source tables or fields'

    return records


def duplicate_counts(duplicate_records):
    counts = Counter(record['category'] for record in duplicate_records)
    return {
        'exact': counts.get('Exact Duplicate', 0),
        'near': counts.get('Near Duplicate', 0),
        'same_source': counts.get('Same Data Source', 0),
        'unique': counts.get('Unique', 0),
        'unparsed': counts.get('Unparsed', 0),
    }



def build_validation_rows_raw(
    file_name,
    sql_view_need='',
    has_join=False,
    has_match=False,
    dataset_shape='',
    has_drilldown=False,
):
    rows = []

    def add_check(category, check, rule):
        rows.append({
            'File Name': file_name,
            'Check Category': category,
            'Validation Check': check,
            'Trigger Rule': rule,
        })

    if sql_view_need == 'Recommended':
        add_check(
            'SQL / Data Validation',
            'Validate SQL view output before modeling it in Framework Manager/Data Module for Cognos.',
            'SQL View Need = Recommended',
        )

    if has_join:
        add_check(
            'SQL / Data Validation',
            'Check join key formatting, nulls, duplicates, and cardinality.',
            'JOIN/MATCH detected',
        )

    add_check(
        'SQL / Data Validation',
        'Compare final row count between WebFOCUS output and the Cognos dataset.',
        'Always',
    )
    add_check(
        'Business Total Validation',
        'Validate totals for the main measure/value columns.',
        'Always',
    )

    if has_match:
        add_check(
            'Business Total Validation',
            'MATCH FILE output validated against WebFOCUS row counts and business totals.',
            'MATCH FILE detected',
        )

    add_check(
        'Report Logic Validation',
        'Check filters and parameters match the WebFOCUS report.',
        'Always',
    )
    add_check(
        'Report Logic Validation',
        'Confirm final visual type and grouping fields.',
        'Always',
    )

    if str(dataset_shape or '').startswith('Long'):
        add_check(
            'Report Logic Validation',
            'Confirm unpivoted delay/category rows match the original stacked/matrix output.',
            'Long/unpivoted dataset detected',
        )

    if has_drilldown:
        add_check(
            'Interaction Validation',
            'Drillthrough opens the correct detail page and passes the expected field/parameter.',
            'Drilldown detected',
        )

    return rows


def build_validation_rows(report_row, join_df, final_df, drilldown_df=None):
    file_name = report_row.get('FEX Name', '')
    has_join = join_df is not None and not join_df.empty
    has_match = False
    if has_join and 'Mapping Type' in join_df.columns:
        has_match = (join_df['Mapping Type'].astype(str) == 'MATCH FILE').any()

    dataset_shape = ''
    if final_df is not None and not final_df.empty:
        dataset_shape = str(final_df.iloc[0].get('Dataset Shape', ''))

    has_drilldown = drilldown_df is not None and not drilldown_df.empty

    return build_validation_rows_raw(
        file_name=file_name,
        sql_view_need=report_row.get('SQL View Need', ''),
        has_join=has_join,
        has_match=has_match,
        dataset_shape=dataset_shape,
        has_drilldown=has_drilldown,
    )


def build_display_tables(parsed_results, duplicate_records, allowed_program_names, matched_pairs):
    duplicate_lookup = {
        (record['folder'], record['fex_name']): record
        for record in duplicate_records
    }

    overview_rows = []
    field_rows = []
    filter_rows = []
    formula_rows = []
    join_rows = []
    duplicate_rows = []
    source_rows = []
    lineage_rows = []
    mapping_rows = []
    final_dataset_rows = []
    build_plan_rows = []
    drilldown_rows = []
    validation_rows = []

    for folder, fex_name, parsed in parsed_results:
        dup = duplicate_lookup.get((folder, fex_name), {})
        profile = dup.get('profile') or {}
        output_text = output_type_text(parsed) if parsed is not None else ''
        sql_view = analyze_sql_view_need(parsed) if parsed is not None else {
            'need': 'Manual Review',
            'build_path': 'Needs Review',
            'action': 'Review parser output before choosing SQL View, Framework Manager/Data Module, or the report layer.',
            'score': 0,
        }
        complexity_score = (
            len(parsed['formula_mapping'])
            + len(parsed['filter_parameter_mapping'])
            + len(parsed['join_match_mapping']) * 2
            + len(parsed['hold_lineage'])
        ) if parsed is not None else 999
        if parsed is None:
            complexity = 'Manual Review'
        elif complexity_score >= 18 or sql_view['need'] == 'Recommended':
            complexity = 'High'
        elif complexity_score >= 9 or sql_view['need'] == 'Optional':
            complexity = 'Medium'
        else:
            complexity = 'Low'
        main_source = ''
        if parsed is not None and parsed.get('real_sources'):
            main_source = parsed['real_sources'][0]

        overview_rows.append({
            'File Path': folder,
            'FEX Name': fex_name,
            'Report': fex_name,
            'Purpose': infer_report_purpose(fex_name, parsed) if parsed is not None else 'Needs parser review',
            'Type': infer_report_type(parsed) if parsed is not None else 'Manual review',
            'Complexity': complexity,
            'Duplicate Group': dup.get('group_id') or dup.get('category', 'Unique'),
            'Main Source': main_source,
            'SQL View Need': sql_view['need'],
            'Build Path': sql_view['build_path'],
            'Cognos Action': sql_view['action'],
            'Status': 'Needs Review' if complexity in {'High', 'Manual Review'} or sql_view['need'] == 'Recommended' else 'Ready',
            'Source Tables': ', '.join(sorted(profile.get('real_tables', []))),
            'Field Count': len(profile.get('all_fields', [])),
            'Formula Count': len(parsed['formula_mapping']) if parsed is not None else 0,
            'Filter Count': len(parsed['filter_parameter_mapping']) if parsed is not None else 0,
            'Join / Match Count': len(parsed['join_match_mapping']) if parsed is not None else 0,
            'Drilldown Count': len(parsed['drilldowns']) if parsed is not None else 0,
            'HOLD Lineage Steps': len(parsed['hold_lineage']) if parsed is not None else 0,
            'Build Plan Steps': len(parsed['cognos_build_plan']) if parsed is not None else 0,
            'Mapping Table Rows': len(parsed['mapping_table_preview']) if parsed is not None else 0,
            'Recommended Final Dataset': parsed['final_dataset_plan'][0]['Recommended Final Dataset'] if parsed is not None and parsed['final_dataset_plan'] else '',
            'Output Format': output_text,
            'Duplicate Type': dup.get('category', 'Unparsed'),
            'Parse Status': 'Parsed' if parsed is not None else 'Unparsed',
        })

        if parsed is None:
            validation_rows.extend(build_validation_rows_raw(file_name=fex_name))
        else:
            validation_rows.extend(build_validation_rows_raw(
                file_name=fex_name,
                sql_view_need=sql_view['need'],
                has_join=bool(parsed.get('join_match_mapping')),
                has_match=any(row.get('mapping_type') == 'MATCH FILE' for row in parsed.get('join_match_mapping', [])),
                dataset_shape=parsed['final_dataset_plan'][0].get('Dataset Shape', '') if parsed.get('final_dataset_plan') else '',
                has_drilldown=bool(parsed.get('drilldowns')),
            ))

        duplicate_rows.append({
            'File Path': folder,
            'FEX Name': fex_name,
            'Duplicate Type': dup.get('category', 'Unparsed'),
            'Duplicate Group': dup.get('group_id', ''),
            'Duplicate Confidence': f"{dup.get('best_similarity', 0):.0%}" if dup.get('best_similarity') else '',
            'Matched With': ', '.join(
                dup.get('exact_matches', [])
                + dup.get('near_matches', [])
                + dup.get('same_source_matches', [])
            ),
            'Why Duplicate': dup.get('difference_summary', ''),
            'Differences': dup.get('difference_summary', ''),
        })

        if parsed is None:
            continue

        for row in field_inventory_value_rows(parsed, folder, fex_name):
            field_rows.append(dict(zip(FIELD_INVENTORY_HEADERS, row)))

        for item in parsed['filter_parameter_mapping']:
            filter_rows.append({
                'File Path': folder,
                'FEX Name': fex_name,
                'Filter Type': item['type'],
                'Expression': item['expression'],
                'Parameter': ', '.join(item['parameters']),
                'Fields Used': ', '.join(item['fields']),
                'Cognos Guidance': item['cognos_guidance'],
                'Suggested Cognos Type': item.get('suggested_cognos_type', ''),
                'Suggested Cognos Expression / Action': item.get('suggested_cognos_expression_action', ''),
                'LLM Notes': item.get('llm_notes', ''),
                'Needs Manual Review': item.get('needs_manual_review', ''),
                'Review Reason': item.get('review_reason', ''),
                'Confidence': item.get('confidence', ''),
            })

        for item in parsed['formula_mapping']:
            formula_rows.append({
                'File Path': folder,
                'FEX Name': fex_name,
                'Formula Source': item['formula_source'],
                'Field Name': item['field'],
                'Result Datatype': item['format'],
                'Formula Type': item['formula_type'],
                'Raw WebFOCUS Formula': item['formula'],
                'Raw Columns Used': ', '.join(item['raw_fields']),
                'Parameters': ', '.join(item['parameters']),
                'Suggested Cognos Type': item.get('suggested_cognos_type', ''),
                'Suggested Cognos Expression / Action': item.get('suggested_cognos_expression_action', ''),
                'Needs Manual Review': item.get('needs_manual_review', ''),
                'Review Reason': item.get('review_reason', ''),
                'Confidence': item.get('confidence', ''),
            })

        for item in parsed['join_match_mapping']:
            join_rows.append({
                'File Path': folder,
                'File Name': fex_name,
                'Mapping Type': item['mapping_type'],
                'From Table': item['from_table'],
                'From Key': item['from_key'],
                'To Table': item['to_table'],
                'To Key': item['to_key'],
                'WebFOCUS Join Mode': item['join_type'],
                'Output HOLD': item['output_hold'],
                'Raw WebFOCUS Statement': item['raw_statement'],
                'Suggested Cognos Action': item.get('suggested_cognos_expression_action', ''),
                'Needs Manual Review': item.get('needs_manual_review', ''),
                'Review Reason': item.get('review_reason', ''),
                'Confidence': item.get('confidence', ''),
            })

        for item in parsed['source_inventory']:
            row = {'File Path': folder, 'File Name': fex_name}
            row.update(item)
            source_rows.append(row)

        for item in parsed['hold_lineage']:
            row = {'File Path': folder, 'File Name': fex_name}
            row.update(item)
            lineage_rows.append(row)

        for item in parsed['mapping_table_preview']:
            row = {'File Path': folder, 'File Name': fex_name}
            row.update(item)
            mapping_rows.append(row)

        for item in parsed['final_dataset_plan']:
            row = {'File Path': folder, 'File Name': fex_name}
            row.update(item)
            final_dataset_rows.append(row)

        for item in parsed['cognos_build_plan']:
            row = {'File Path': folder, 'File Name': fex_name}
            row.update(item)
            build_plan_rows.append(row)

        for item in parsed['drilldowns']:
            row = {'File Path': folder, 'File Name': fex_name}
            row.update(item)
            drilldown_rows.append(row)

    matched_lookup = defaultdict(list)
    for ra_program, fex_name in matched_pairs:
        matched_lookup[ra_program].append(fex_name)

    ra_rows = []
    for program_name in sorted(allowed_program_names):
        matched_files = sorted(set(matched_lookup.get(program_name, [])))
        ra_rows.append({
            'Resource Analyzer Program': program_name,
            'Matched FEX File': ', '.join(matched_files) if matched_files else 'Not Found in Uploaded FEX Folder/ZIP',
            'Match Status': 'Matched' if matched_files else 'Not Found',
        })

    return {
        'overview': pd.DataFrame(overview_rows),
        'fields': pd.DataFrame(field_rows),
        'filters': pd.DataFrame(filter_rows),
        'formulas': pd.DataFrame(formula_rows),
        'joins': pd.DataFrame(join_rows),
        'sources': pd.DataFrame(source_rows),
        'lineage': pd.DataFrame(lineage_rows),
        'mapping_tables': pd.DataFrame(mapping_rows),
        'final_dataset': pd.DataFrame(final_dataset_rows),
        'build_plan': pd.DataFrame(build_plan_rows),
        'drilldowns': pd.DataFrame(drilldown_rows),
        'duplicates': pd.DataFrame(duplicate_rows),
        'resource_analyzer': pd.DataFrame(ra_rows),
        'validation': pd.DataFrame(validation_rows),
    }



def write_filter_parameter_sheet(ws, parsed_results):
    setup_sheet(ws, FILTER_PARAM_HEADERS, FILTER_PARAM_WIDTHS)
    row = 2

    for folder, fex_name, parsed in parsed_results:
        if parsed is None:
            continue

        for item in parsed['filter_parameter_mapping']:
            vals = [
    folder,
    fex_name,
    item['type'],
    item['expression'],
    ', '.join(item['parameters']),
    ', '.join(item['fields']),
]

            for col, val in enumerate(vals, 1):
                _write_cell(ws, row, col, val, 'FFFFFF', '000000')

            row += 1


def write_formula_mapping_sheet(ws, parsed_results):
    setup_sheet(ws, FORMULA_MAPPING_HEADERS, FORMULA_MAPPING_WIDTHS)
    row = 2

    for folder, fex_name, parsed in parsed_results:
        if parsed is None:
            continue

        for item in parsed['formula_mapping']:
            vals = [
    folder,
    fex_name,
    item['formula_source'],
    item['field'],
    item['format'],
    item['formula_type'],
    item['formula'],
    ', '.join(item['raw_fields']),
    ', '.join(item['parameters']),
]

            for col, val in enumerate(vals, 1):
                _write_cell(ws, row, col, val, 'FFFFFF', '000000')

            row += 1




def write_join_match_sheet(ws, parsed_results):
    setup_sheet(ws, JOIN_MATCH_HEADERS, JOIN_MATCH_WIDTHS)
    row = 2

    for folder, fex_name, parsed in parsed_results:
        if parsed is None:
            continue

        for item in parsed['join_match_mapping']:
            vals = [
                folder,
                fex_name,
                item['mapping_type'],
                item['from_table'],
                item['from_key'],
                item['to_table'],
                item['to_key'],
                item['join_type'],
                item['output_hold'],
                item['raw_statement'],
            ]

            for col, val in enumerate(vals, 1):
                _write_cell(ws, row, col, val, 'FFFFFF', '000000')

            row += 1
def clean_value(value) -> str:
    """Clean a value pulled from Excel or metadata text."""
    if value is None:
        return ""

    text = str(value).strip()
    text = text.strip('"').strip("'")
    text = text.rstrip(",;")
    text = text.rstrip("$")
    return text.strip()


def normalize_name(value) -> str:
    """Normalize table/file names for MAS/ACX matching."""
    text = clean_value(value).upper()
    text = text.replace(".MAS", "").replace(".ACX", "")
    return text


def collect_metadata_from_zip(uploaded_zip):
    """Reads .mas and .acx files from an uploaded ZIP, keyed by normalized
    stem name. Returns two dicts: {normalized_name: (entry_name, text_content)}."""
    mas_files = {}
    acx_files = {}

    uploaded_zip.seek(0)

    with zipfile.ZipFile(uploaded_zip, 'r') as zf:
        for name in zf.namelist():
            if name.endswith('/'):
                continue

            if name.startswith('__MACOSX/') or '/__MACOSX/' in name:
                continue

            if Path(name).name.startswith('._'):
                continue

            suffix = Path(name).suffix.lower()
            if suffix not in ('.mas', '.acx'):
                continue

            stem = normalize_name(Path(name).stem)

            try:
                content = zf.read(name).decode('utf-8', errors='replace')
            except Exception:
                content = ''

            if suffix == '.mas':
                if stem not in mas_files:
                    mas_files[stem] = (name, content)
            else:
                if stem not in acx_files:
                    acx_files[stem] = (name, content)

    return mas_files, acx_files


def find_best_metadata_match(unique_table, mas_files, acx_files):
    """Find matching MAS/ACX entries for a KashMap unique table name."""
    table_upper = normalize_name(unique_table)
    table_short = re.split(r"[./]", table_upper)[-1]

    mas_entry = None
    acx_entry = None
    match_parts = []

    if table_upper in mas_files:
        mas_entry = mas_files[table_upper]
        match_parts.append("Exact MAS filename match")

    if table_upper in acx_files:
        acx_entry = acx_files[table_upper]
        match_parts.append("Exact ACX filename match")

    if not mas_entry and table_short in mas_files:
        mas_entry = mas_files[table_short]
        match_parts.append("Short MAS filename match")

    if not acx_entry and table_short in acx_files:
        acx_entry = acx_files[table_short]
        match_parts.append("Short ACX filename match")

    if not mas_entry:
        for name, entry in mas_files.items():
            if table_short == name or table_short in name or name in table_short:
                mas_entry = entry
                match_parts.append("Partial MAS filename match")
                break

    if not acx_entry:
        for name, entry in acx_files.items():
            if table_short == name or table_short in name or name in table_short:
                acx_entry = entry
                match_parts.append("Partial ACX filename match")
                break

    match_type = " + ".join(match_parts) if match_parts else "No metadata filename match"

    return mas_entry, acx_entry, match_type


def extract_acx_details_from_text(acx_text):
    """Extract actual DB table and connection from ACX file text."""
    details = {
        "actual_db_tables": [],
        "connection": "",
        "segname": "",
        "evidence": [],
    }

    if not acx_text:
        return details

    text_clean = acx_text.replace("\n", " ").replace("\r", " ")
    actual_tables = []
    evidence = []

    seg_match = re.search(r"\bSEGNAME\s*=\s*([^,\s$]+)", text_clean, flags=re.IGNORECASE)
    if seg_match:
        segname = clean_value(seg_match.group(1))
        details["segname"] = segname
        evidence.append(f"SEGNAME={segname}")

    table_matches = re.findall(r"\bTABLENAME\s*=\s*([^,\s$]+)", text_clean, flags=re.IGNORECASE)
    for value in table_matches:
        value = clean_value(value)
        if value and value not in actual_tables:
            actual_tables.append(value)
            evidence.append(f"TABLENAME={value}")

    if not actual_tables:
        fallback_keys = [
            "TABLE_NAME", "SQLTABLENAME", "SQL_TABLE", "DBMS_TABLE",
            "OBJECT", "LOCATION", "DATASET", "DATASETNAME", "PHYSICAL_NAME",
        ]
        for key in fallback_keys:
            pattern = rf"\b{re.escape(key)}\s*=\s*([^,\s$]+)"
            matches = re.findall(pattern, text_clean, flags=re.IGNORECASE)
            for value in matches:
                value = clean_value(value)
                if value and value.upper() not in {
                    "FOCUS", "XFOCUS", "SQLMSS", "SQLORA", "DB2", "MYSQL", "POSTGRES",
                }:
                    if value not in actual_tables:
                        actual_tables.append(value)
                        evidence.append(f"{key}={value}")

    conn_match = re.search(r"\bCONNECTION\s*=\s*([^,\s$]+)", text_clean, flags=re.IGNORECASE)
    if conn_match:
        connection = clean_value(conn_match.group(1))
        details["connection"] = connection
        evidence.append(f"CONNECTION={connection}")

    details["actual_db_tables"] = actual_tables
    details["evidence"] = evidence

    return details


def build_table_db_mapping(unique_tables, mas_files, acx_files):
    """Match each KashMap unique table to actual DB tables via MAS/ACX metadata."""
    rows = []
    mapped_count = 0
    needs_review_count = 0

    for unique_table in sorted(unique_tables):
        mas_entry, acx_entry, match_type = find_best_metadata_match(unique_table, mas_files, acx_files)

        mas_name = mas_entry[0] if mas_entry else ''
        acx_name = acx_entry[0] if acx_entry else ''
        acx_text = acx_entry[1] if acx_entry else ''

        acx_details = extract_acx_details_from_text(acx_text)
        actual_db_tables = ', '.join(acx_details.get('actual_db_tables', []))
        connection = acx_details.get('connection', '')
        evidence = ' | '.join(acx_details.get('evidence', []))

        if actual_db_tables:
            notes = evidence if evidence else 'Found using TABLENAME in ACX'
            mapped_count += 1
        elif mas_entry or acx_entry:
            notes = f"Metadata file found but actual database table not parsed. Match type: {match_type}"
            needs_review_count += 1
        else:
            notes = 'No matching MAS/ACX metadata file found. Needs review.'
            needs_review_count += 1

        rows.append({
            'Unique Table from KashMap': unique_table,
            'Actual Database Table': actual_db_tables,
            'Connection': connection,
            'MAS File': mas_name,
            'ACX File': acx_name,
            'Evidence / Notes': notes,
        })

    return rows, mapped_count, needs_review_count

def read_uploaded_fex(uploaded_file):
    return uploaded_file.read().decode('utf-8', errors='replace')

def _is_real_fex_entry(name: str) -> bool:
    if not name.lower().endswith('.fex'):
        return False
    if name.startswith('__MACOSX/') or '/__MACOSX/' in name:
        return False
    if Path(name).name.startswith('._'):
        return False
    return True

def count_fex_in_zip(uploaded_zip):
    uploaded_zip.seek(0)
    try:
        with zipfile.ZipFile(uploaded_zip, 'r') as zf:
            count = sum(
                1 for name in zf.namelist()
                if _is_real_fex_entry(name) and not name.endswith('/')
            )
    except Exception:
        count = None
    finally:
        uploaded_zip.seek(0)
    return count


def collect_fex_from_zip(uploaded_zip):
    files = []

    uploaded_zip.seek(0)

    with zipfile.ZipFile(uploaded_zip, 'r') as zf:
        for name in zf.namelist():
            if _is_real_fex_entry(name):
                try:
                    content = zf.read(name).decode('utf-8', errors='replace')
                    files.append((str(Path(name).parent), Path(name).name, content))
                except Exception as e:
                    raise ValueError(f"Failed reading {name}: {e}")

    return files



def _append_row_values(ws, values, bg='FFFFFF', fg='000000', bold=False):
    cells = []

    for value in values:
        cell = WriteOnlyCell(ws, value=_excel_safe_value(value))
        cell.font = _font(fg, bold)
        cell.alignment = WRAP_TOP
        cell.fill = _fill(bg)
        cell.border = BORDER
        cells.append(cell)

    ws.append(cells)


def _setup_write_only_sheet(wb, title, headers, widths):
    ws = wb.create_sheet(title)
    hbg, hfg = COLORS['header']
    _append_row_values(ws, headers, hbg, hfg, bold=True)

    for col, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = width

    return ws


def output_type_text(parsed):
    if parsed is None or not parsed['output_formats']:
        return ''

    values = []
    for item in parsed['output_formats']:
        label = item['command']
        if item['output_name']:
            label += f" AS {item['output_name']}"
        if item['format']:
            label += f" FORMAT {item['format']}"
        values.append(label)

    return '; '.join(dict.fromkeys(values))


def field_inventory_value_rows(parsed, folder, fex_name):
    rows = []
    source_name_set = parsed['source_name_set']
    calculated_name_set = parsed['calculated_name_set']
    file_definitions = parsed.get('file_definitions', {})
    shown_fields = {
        f['field'].upper()
        for f in (
            parsed['sum_real']
            + parsed['sum_calc']
            + parsed['by_real']
            + parsed['by_calc']
        )
    }
    output_type = output_type_text(parsed)

    def add_row(
        source_table,
        field_name,
        field_origin,
        field_role,
        formula='',
        raw_fields='',
        result_datatype='',
        formula_source='',
        is_shown=None,
    ):
        field_upper = field_name.upper()
        shown = 'Yes' if (is_shown if is_shown is not None else field_upper in shown_fields) else 'No'
        file_def = file_definitions.get(str(source_table).upper(), '')

        rows.append([
            folder,
            fex_name,
            source_table,
            file_def,
            field_name,
            field_origin,
            field_role,
            formula,
            raw_fields,
            result_datatype,
            formula_source,
            shown,
            output_type,
        ])

    source_names_only = {f['field'] for f in parsed['source_fields']}

    for f in parsed['source_fields']:
        add_row(
            f['source'],
            f['field'],
            'DB',
            classify_field_role(f['field'], source_name_set, calculated_name_set),
        )

    for f in parsed['define_fields']:
        add_row(
            f['source'],
            f['field'],
            'Both' if f['field'] in source_name_set else 'Calc',
            classify_field_role(f['field'], source_name_set, calculated_name_set),
            f['formula'],
            ', '.join(f['raw_fields']),
            f.get('format', ''),
            'DEFINE',
        )

    for f in parsed['compute_fields']:
        add_row(
            f['source'],
            f['field'],
            'Both' if f['field'] in source_name_set else 'Calc',
            classify_field_role(f['field'], source_name_set, calculated_name_set),
            f['formula'],
            ', '.join(f['raw_fields']),
            f.get('format', ''),
            'COMPUTE',
        )

    for f in parsed['by_real']:
        if f['field'] in source_names_only:
            continue
        add_row(
            f['source'],
            f['field'],
            'DB',
            classify_field_role(f['field'], source_name_set, calculated_name_set),
            is_shown=True,
        )

    for f in parsed['by_calc']:
        add_row(
            f['source'],
            f['field'],
            'Calc',
            classify_field_role(f['field'], source_name_set, calculated_name_set),
            is_shown=True,
        )

    return rows


def detail_value_rows(parsed, folder, fex_name):
    return field_inventory_value_rows(parsed, folder, fex_name)


def build_output_workbook_cloud_safe(fex_items, allowed_program_names, matched_pairs, llm_options=None, mas_files=None, acx_files=None):
    errors = []
    total = len(fex_items)

    progress = st.progress(0)
    status = st.empty()

    status.text("Pass 1 of 2: Parsing selected FEX files...")

    parsed_results = []

    for idx, (folder, fex_name, content) in enumerate(fex_items, start=1):
        try:
            parsed = parse_fex(content)
            parsed_results.append((folder, fex_name, parsed))
        except Exception as e:
            errors.append(f"{fex_name}: {e}")
            parsed_results.append((folder, fex_name, None))

        progress.progress(idx / total * 0.50 if total else 1.0)

    duplicate_records = build_duplicate_analysis(parsed_results)
    apply_llm_translations(parsed_results, llm_options)
    counts = duplicate_counts(duplicate_records)
    display_tables = build_display_tables(
        parsed_results,
        duplicate_records,
        allowed_program_names,
        matched_pairs,
    )
    dup_group_count = counts['exact'] + counts['near'] + counts['same_source']
    unique_count = counts['unique']
    unparsed_count = counts['unparsed']
    migration_count = counts['exact'] + counts['near'] + counts['same_source'] + counts['unique'] + counts['unparsed']

    unique_tables = {
        table.upper()
        for _, _, parsed in parsed_results
        if parsed is not None
        for table in parsed['real_sources']
    }
    total_tables = len(unique_tables)

    table_db_mapping_rows = []
    table_db_mapped_count = 0
    table_db_needs_review_count = 0
    if mas_files or acx_files:
        status.text("Matching unique tables to MAS/ACX metadata...")
        table_db_mapping_rows, table_db_mapped_count, table_db_needs_review_count = build_table_db_mapping(
            unique_tables, mas_files or {}, acx_files or {}
        )

    status.text("Pass 2 of 2: Creating low-memory Excel workbook...")
    wb = Workbook(write_only=True)

    ws_detail = _setup_write_only_sheet(wb, 'Field Inventory', FIELD_INVENTORY_HEADERS, FIELD_INVENTORY_WIDTHS)
    ws_duplicates = _setup_write_only_sheet(wb, 'Duplicate Analysis', DUPLICATE_ANALYSIS_HEADERS, DUPLICATE_ANALYSIS_WIDTHS)
    ws_ra = _setup_write_only_sheet(wb, 'Resource Analyzer Matches', RESOURCE_ANALYZER_HEADERS, RESOURCE_ANALYZER_WIDTHS)
    ws_filters = _setup_write_only_sheet(wb, 'Filters and Parameters', FILTER_PARAM_HEADERS, FILTER_PARAM_WIDTHS)
    ws_formulas = _setup_write_only_sheet(wb, 'Formula Mapping', FORMULA_MAPPING_HEADERS, FORMULA_MAPPING_WIDTHS)
    ws_join = _setup_write_only_sheet(wb, 'Join Match', JOIN_MATCH_HEADERS, JOIN_MATCH_WIDTHS)
    ws_summary = _setup_write_only_sheet(wb, 'Final Summary', FINAL_SUMMARY_HEADERS, FINAL_SUMMARY_WIDTHS)
    ws_validation = _setup_write_only_sheet(wb, 'Validation Checklist', VALIDATION_CHECKLIST_HEADERS, VALIDATION_CHECKLIST_WIDTHS)
    ws_table_db_mapping = _setup_write_only_sheet(wb, 'Table DB Mapping', TABLE_DB_MAPPING_HEADERS, TABLE_DB_MAPPING_WIDTHS)

    status.text("Writing Field Inventory sheet...")
    for idx, (folder, fex_name, parsed) in enumerate(parsed_results, start=1):
        if parsed is None:
            continue
        try:
            for row_values in field_inventory_value_rows(parsed, folder, fex_name):
                origin = row_values[5]
                if origin == 'DB':
                    bg, fg = COLORS['source']
                elif row_values[10] == 'DEFINE':
                    bg, fg = COLORS['define']
                elif row_values[10] == 'COMPUTE':
                    bg, fg = COLORS['compute']
                else:
                    bg, fg = 'FFFFFF', '000000'
                _append_row_values(ws_detail, row_values, bg, fg)
        except Exception as e:
            errors.append(f"{fex_name} write error: {e}")

        progress.progress(0.50 + idx / total * 0.20 if total else 0.70)

    status.text("Writing Duplicate Analysis sheet...")
    category_colors = {
        'Exact Duplicate': COLORS['grp_a'],
        'Near Duplicate': COLORS['grp_b'],
        'Same Data Source': COLORS['grp_c'],
        'Unique': COLORS['unique'],
        'Unparsed': COLORS['unparsed'],
    }
    for record in sorted(duplicate_records, key=lambda r: (r['category'], r['folder'].lower(), r['fex_name'].lower())):
        profile = record['profile'] or {}
        row_values = [
            record['category'],
            record['group_id'],
            record['group_size'],
            record['folder'],
            record['fex_name'],
            ', '.join(sorted(profile.get('real_tables', []))),
            len(profile.get('all_fields', [])),
            ', '.join(record['exact_matches']),
            ', '.join(record['near_matches']),
            ', '.join(record['same_source_matches']),
            f"{record['best_similarity']:.0%}" if record['best_similarity'] else '',
            record['difference_summary'],
        ]
        bg, fg = category_colors.get(record['category'], ('FFFFFF', '000000'))
        _append_row_values(ws_duplicates, row_values, bg, fg, bold=record['category'] in {'Exact Duplicate', 'Near Duplicate'})
    progress.progress(0.76)

    status.text("Writing Resource Analyzer match sheet...")
    matched_lookup = defaultdict(list)
    for ra_program, fex_name in matched_pairs:
        matched_lookup[ra_program].append(fex_name)
    for program_name in sorted(allowed_program_names):
        matched_files = ', '.join(sorted(set(matched_lookup.get(program_name, []))))
        _append_row_values(ws_ra, [
            program_name,
            matched_files if matched_files else 'Not Found in Uploaded FEX Folder/ZIP',
        ])
    progress.progress(0.82)

    status.text("Writing Filters and Parameters sheet...")
    for folder, fex_name, parsed in parsed_results:
        if parsed is None:
            continue
        for item in parsed['filter_parameter_mapping']:
            _append_row_values(ws_filters, [
                folder, fex_name, item['type'], item['expression'],
                ', '.join(item['parameters']), ', '.join(item['fields']),
            ])
    progress.progress(0.88)

    status.text("Writing Formula Mapping sheet...")
    for folder, fex_name, parsed in parsed_results:
        if parsed is None:
            continue
        for item in parsed['formula_mapping']:
            _append_row_values(ws_formulas, [
                folder, fex_name, item['formula_source'], item['field'],
                item['format'], item['formula_type'], item['formula'],
                ', '.join(item['raw_fields']), ', '.join(item['parameters']),
            ])
    progress.progress(0.92)

    status.text("Writing Join Match sheet...")
    for folder, fex_name, parsed in parsed_results:
        if parsed is None:
            continue
        for item in parsed['join_match_mapping']:
            _append_row_values(ws_join, [
                folder, fex_name, item['mapping_type'], item['from_table'], item['from_key'],
                item['to_table'], item['to_key'], item['join_type'], item['output_hold'],
                item['raw_statement'],

            ])
    progress.progress(0.95)

    status.text("Writing Table DB Mapping sheet...")
    for row in table_db_mapping_rows:
        _append_row_values(ws_table_db_mapping, [
            row['Unique Table from KashMap'],
            row['Actual Database Table'],
            row['Connection'],
            row['MAS File'],
            row['ACX File'],
            row['Evidence / Notes'],
        ])

    status.text("Writing Validation Checklist sheet...")
    validation_df = display_tables.get('validation', pd.DataFrame())
    if validation_df is not None and not validation_df.empty:
        for _, item in validation_df.iterrows():
            _append_row_values(ws_validation, [
                item.get('File Name', ''),
                item.get('Check Category', ''),
                item.get('Validation Check', ''),
                item.get('Trigger Rule', ''),
            ])
    progress.progress(0.965)

    status.text("Writing Final Summary sheet...")
    summary_rows = [
        ('Total Reports Listed in Resource Analyzer', len(allowed_program_names)),
        ('Total FEX Files Processed', len(fex_items)),
        ('Exact Duplicate Files', counts['exact']),
        ('Near Duplicate Files', counts['near']),
        ('Same Data Source Files', counts['same_source']),
        ('Unique Files', counts['unique']),
        ('Unparsed Files', counts['unparsed']),
        ('Total Unique Real Source Tables Used', total_tables),
        ('Field Inventory Rows', sum(1 for _, _, p in parsed_results if p is not None for _ in field_inventory_value_rows(p, '', ''))),
        ('Filter / Parameter Rows', sum(len(p['filter_parameter_mapping']) for _, _, p in parsed_results if p is not None)),
        ('Join / Match Rows', sum(len(p['join_match_mapping']) for _, _, p in parsed_results if p is not None)),
        ('Output Format Rows Included In Field Inventory', sum(len(p['output_formats']) for _, _, p in parsed_results if p is not None)),
        ('', ''),
        ('All Unique Tables Used', 'Table Name'),
    ]
    for row_values in summary_rows:
      if row_values == ('', ''):
          _append_row_values(ws_summary, row_values)
      elif row_values == ('All Unique Tables Used', 'Table Name'):
          _append_row_values(ws_summary, row_values, bg='FFE7D6', fg='7A2E0E', bold=True)
      else:
        _append_row_values(ws_summary, row_values)

    for idx, table_name in enumerate(sorted(unique_tables), start=1):
      _append_row_values(ws_summary, [idx, table_name])
    progress.progress(0.98)

    status.text("Saving Excel workbook...")
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    progress.progress(1.0)
    status.text("Excel workbook ready.")

    table_db_mapping_df = pd.DataFrame(table_db_mapping_rows)
    if not table_db_mapping_df.empty:
        table_db_mapping_df = table_db_mapping_df.astype(str).replace('nan', '')
    display_tables['table_db_mapping'] = table_db_mapping_df
    
    return (
            output,
            errors,
            dup_group_count,
            unique_count,
            unparsed_count,
            migration_count,
            total_tables,
            counts,
            display_tables,
            table_db_mapped_count,
            table_db_needs_review_count,
        )


def build_output_workbook(fex_items, allowed_program_names, matched_pairs, llm_options=None, mas_files=None, acx_files=None):
    return build_output_workbook_cloud_safe(
        fex_items,
        allowed_program_names,
        matched_pairs,
        llm_options,
        mas_files,
        acx_files,
    )


try:
    st.set_page_config(page_title="KashMap Migration Workspace", layout="wide")
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

.diff-side-by-side {
    width: 100%;
    border-collapse: collapse;
    font-family: 'Consolas', 'Menlo', monospace;
    font-size: 0.82rem;
    table-layout: fixed;
}
.diff-side-by-side th {
    background: #1F2933;
    color: #FFFFFF;
    text-align: left;
    padding: 8px 10px;
    font-family: Arial, sans-serif;
    font-size: 0.82rem;
    position: sticky;
    top: 0;
}
.diff-side-by-side th.diff-linenum-header {
    width: 48px;
}
.diff-side-by-side td {
    padding: 3px 10px;
    vertical-align: top;
    white-space: pre-wrap;
    word-break: break-word;
    border-bottom: 1px solid #EEF1F5;
}
.diff-side-by-side td.diff-linenum {
    width: 48px;
    text-align: right;
    color: #94A3B8;
    background: #FAFBFC;
    user-select: none;
    padding-right: 8px;
    border-right: 1px solid #EEF1F5;
    white-space: nowrap;
}
.diff-cell-removed { background: #FFEBEE; color: #B42318; }
.diff-cell-added { background: #E8F5E9; color: #1B5E20; }
.diff-cell-equal { background: #FFFFFF; color: #444B56; }
.diff-cell-empty { background: #FAFAFA; }
.diff-cell-collapsed { background: #F8FAFC; color: #94A3B8; font-style: italic; text-align: center; }

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

.card-title-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
    margin-bottom: 4px;
}

.file-list-preview {
    margin-top: 10px;
    display: grid;
    gap: 8px;
}

.file-row-preview {
    background: rgba(127, 127, 127, 0.08);
    border: 1px solid var(--kash-border);
    border-radius: 10px;
    padding: 8px 10px;
    font-size: 0.84rem;
    display: flex;
    justify-content: space-between;
    align-items: center;
}

.file-size-muted {
    opacity: 0.65;
    font-size: 0.78rem;
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
.badge-sql {
    background-color: var(--kash-navy);
    color: #FFFFFF;
}
.badge-pq {
    background-color: var(--kash-blue-soft);
    color: #0B4DB3;
}
.badge-calc {
    background-color: #F3E8FF;
    color: #6B21A8;
}
.badge-visual {
    background-color: #FFF7D6;
    color: #7A5C00;
}
.badge-review {
    background-color: #FEF3C7;
    color: #92400E;
}
.badge-error {
    background-color: #FEE2E2;
    color: var(--kash-red);
}
.step-title {
    font-weight: 800;
    margin-bottom: 4px;
}
.workspace-card {
    background: rgba(127, 127, 127, 0.08);
    border: 1px solid var(--kash-border);
    border-radius: 8px;
    padding: 20px;
    box-shadow: 0 8px 24px rgba(15, 23, 42, 0.055);
    margin-bottom: 14px;
}
.report-card {
    min-height: 168px;
    border-top: 4px solid var(--kash-gold);
}
.report-title {
    font-size: 1.05rem;
    font-weight: 800;
    margin-bottom: 8px;
}
.report-meta {
    opacity: 0.65;
    font-size: 0.9rem;
    line-height: 1.45;
    margin-bottom: 12px;
}
.status-pill {
    display: inline-block;
    border-radius: 999px;
    padding: 5px 10px;
    font-size: 0.76rem;
    font-weight: 800;
    margin-right: 6px;
    margin-bottom: 6px;
}
.pill-sql { background: var(--kash-navy); color: #FFFFFF; }
.pill-pq { background: var(--kash-blue-soft); color: var(--kash-blue); }
.pill-review { background: #FEF3C7; color: #92400E; }
.pill-ok { background: #DCFCE7; color: #166534; }
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


st.markdown(
    """
    <div class="kash-hero">
      <div class="kash-title"><span class="kash-title-kash">Kash</span><span class="kash-title-map">Map</span></div>
      <div class="kash-subtitle">Break Down <span class="kash-subtitle-accent">WebFOCUS</span> Complexity for Faster <span class="kash-subtitle-accent">Cognos</span> Migration</div>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.expander("How to use KashMap", expanded=True):
    st.markdown(
        """
        1. Upload one or more `.fex` files, or upload a ZIP folder containing FEX files.
        2. Optional: upload a Resource Analyzer Excel/CSV file to match or filter active reports.
        3. Click **Analyze Files**.
        4. Review report cards first, then open the guided Cognos build plan.
        5. Use the raw tables only when you need technical detail.
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
                <div class="card-title">Upload FEX Reports</div>
                <div class="card-subtitle">
                    Upload multiple .fex files or one ZIP folder containing FEX reports.
                </div>
                """,
                unsafe_allow_html=True,
            )

        with badge_col:
            st.markdown(
                '<div class="badge-required">Required</div>',
                unsafe_allow_html=True,
            )

        mode = st.radio(
            "FEX input type",
            ["Multiple FEX Files", "ZIP File"],
            horizontal=True,
            key="fex_input_type",
        )

        if mode == "Multiple FEX Files":
            uploaded_fex_files = st.file_uploader(
                "Upload one or more .fex files",
                type=["fex"],
                accept_multiple_files=True,
                key="fex_files_uploader",
            )
            uploaded_zip = None

            if uploaded_fex_files:
                total_size = sum(getattr(file, "size", 0) for file in uploaded_fex_files)

                st.markdown(
                    f"""
                    <div class="file-count-box">
                        {len(uploaded_fex_files)} FEX file(s) uploaded · {total_size / 1024:.1f} KB total
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

        else:
            uploaded_zip = st.file_uploader(
                "Upload ZIP file containing FEX files",
                type=["zip"],
                key="zip_file_uploader",
            )
            uploaded_fex_files = []

            if uploaded_zip:
                fex_count = count_fex_in_zip(uploaded_zip)

                if fex_count is None:
                    count_text = "Could not read ZIP contents · "
                elif fex_count == 0:
                    count_text = "0 .fex files detected · "
                else:
                    count_text = f"{fex_count} .fex file(s) detected · "

                st.markdown(
                    f"""
                    <div class="file-count-box">
                        {count_text}ZIP uploaded · {uploaded_zip.size / 1024:.1f} KB
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                if fex_count == 0:
                    st.warning("This ZIP doesn't appear to contain any .fex files.")


with side_col:
    ra_card = st.container(border=True)

    with ra_card:
        title_col, badge_col = st.columns([0.78, 0.22])

        with title_col:
            st.markdown(
                """
                <div class="card-title">Resource Analyzer File</div>
                <div class="card-subtitle">
                    Upload this only if you want to match or filter active reports.
                </div>
                """,
                unsafe_allow_html=True,
            )

        with badge_col:
            st.markdown(
                '<div class="badge-optional">Optional</div>',
                unsafe_allow_html=True,
            )

        resource_analyzer_file = st.file_uploader(
            "Upload Resource Analyzer file",
            type=["xlsx", "xls", "csv"],
            help="Optional. Upload this only if you want to match/filter reports using Resource Analyzer data.",
            key="resource_analyzer_uploader",
        )

    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

    metadata_card = st.container(border=True)

    with metadata_card:
        title_col, badge_col = st.columns([0.78, 0.22])

        with title_col:
            st.markdown(
                """
                <div class="card-title">Metadata ZIP (.mas / .acx)</div>
                <div class="card-subtitle">
                    Optional. Upload a ZIP of your WebFOCUS metadata folders to map unique tables to actual database tables.
                </div>
                """,
                unsafe_allow_html=True,
            )

        with badge_col:
            st.markdown(
                '<div class="badge-optional">Optional</div>',
                unsafe_allow_html=True,
            )

        metadata_zip_file = st.file_uploader(
            "Upload metadata ZIP",
            type=["zip"],
            help="ZIP containing your extracted .mas and .acx metadata folders.",
            key="metadata_zip_uploader",
        )

        if metadata_zip_file:
            st.markdown(
                f"""
                <div class="file-count-box">
                    Metadata ZIP uploaded · {metadata_zip_file.size / 1024:.1f} KB
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
            <div class="card-subtitle">
                Name of the generated Excel migration workbook.
            </div>
            """,
            unsafe_allow_html=True,
        )

        output_name = st.text_input(
            "Output file name",
            value="KashMap_WebFOCUS_Migration_Output.xlsx",
            key="output_file_name",
        )

    ai_usage_mode = "Minimal"
    use_llm_translation = False
    llm_translate_formulas = False
    llm_translate_filters = False
    llm_max_rows = 200
    llm_model = "gpt-4o-mini"
llm_options = {
    "enabled": use_llm_translation,
    "formulas": llm_translate_formulas,
    "filters": llm_translate_filters,
    "max_rows": llm_max_rows,
    "model": llm_model,
    "mode": ai_usage_mode,
}

st.markdown('<div class="section-label">Manual Expression Translator</div>', unsafe_allow_html=True)
with st.expander("Translate one WebFOCUS expression or query", expanded=False):
    manual_type = st.selectbox(
        "Expression type",
        ["Formula / DEFINE / COMPUTE", "Filter / WHERE / IF", "JOIN / MATCH", "Other"],
        key="manual_llm_type",
    )
    manual_context = st.text_input(
        "Optional context",
        placeholder="Example: field name, source table, report name, expected output",
        key="manual_llm_context",
    )
    manual_expression = st.text_area(
        "Paste WebFOCUS expression/query",
        height=150,
        key="manual_llm_expression",
    )
    if st.button("Translate to Cognos", key="manual_llm_translate"):
        client = get_openai_client()
        if client is None:
            st.error("OpenAI is not configured. Add OPENAI_API_KEY to .streamlit/secrets.toml and make sure openai is in requirements.txt.")
        elif not manual_expression.strip():
            st.warning("Paste a WebFOCUS expression or query first.")
        else:
            with st.spinner("Translating with LLM..."):
                result = translate_expression_with_llm(
                    client,
                    llm_model,
                    manual_type,
                    manual_expression,
                    manual_context,
                )
            st.subheader("Cognos Translation")
            st.write("**Suggested Cognos Type**")
            st.write(result.get("suggested_cognos_type", "") or "Not returned")
            st.write("**Suggested Cognos Expression / Action**")
            st.code(result.get("suggested_cognos_expression_action", "") or "Not returned", language="sql")
            st.write("**LLM Notes**")
            st.write(result.get("llm_notes", "") or "Not returned")
            st.write("**Needs Manual Review**")
            st.write(result.get("needs_manual_review", "") or "Yes")
            st.write("**Review Reason**")
            st.write(result.get("review_reason", "") or "Review before using in Cognos.")
            st.write("**Confidence**")
            st.write(result.get("confidence", "") or "Not returned")

submit = st.button("Analyze Files", type="primary", use_container_width=True)


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

def show_table_or_info(df, message, large=False, height="content"):
    if df is None or df.empty:
        st.info(message)
        return

    # Raw / very large tables stay as normal Streamlit tables.
    # Only tables explicitly marked large=True (e.g. huge mapping tables
    # with hundreds/thousands of rows) or genuinely oversized tables skip
    # the styled KashMap look — mid-size tables (lineage steps, formulas,
    # etc. on complex reports) still get the styled, scrollable version.
    if large or len(df) > 150:
        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True,
            height=height,
        )
        return

    # Small summary tables become styled HTML tables
    html = df.to_html(
        index=False,
        escape=True,
        classes="kash-table",
        border=0,
    )

    st.markdown(
        f"""
        <div class="kash-table-wrap">
            {html}
        </div>
        """,
        unsafe_allow_html=True,
    )



def find_duplicate_match_options(duplicate_df, overview_df, selected_name, selected_path):
    """Returns a list of (display_label, match_name, match_path) tuples for
    files that were flagged as exact/near duplicates of the selected report."""
    if duplicate_df is None or duplicate_df.empty:
        return []

    row_matches = duplicate_df[
        (duplicate_df['FEX Name'] == selected_name)
        & (duplicate_df['File Path'] == selected_path)
    ]
    if row_matches.empty:
        return []

    matched_names_raw = str(row_matches.iloc[0].get('Matched With', '') or '')
    matched_names = [n.strip() for n in matched_names_raw.split(',') if n.strip()]

    options = []
    for matched_name in matched_names:
        candidates = overview_df[overview_df['FEX Name'] == matched_name]
        for _, cand_row in candidates.iterrows():
            cand_path = cand_row['File Path']
            if matched_name == selected_name and cand_path == selected_path:
                continue
            label = f"{matched_name} | {cand_path}"
            options.append((label, matched_name, cand_path))

    seen = set()
    deduped = []
    for label, name, path in options:
        key = (name, path)
        if key not in seen:
            seen.add(key)
            deduped.append((label, name, path))

    return deduped


def get_raw_fex_text(raw_fex_content, folder, fex_name):
    if not raw_fex_content:
        return ''
    return raw_fex_content.get((folder, fex_name), '')


def build_fex_diff_text(text_a, name_a, text_b, name_b):
    lines_a = text_a.splitlines(keepends=True)
    lines_b = text_b.splitlines(keepends=True)

    diff = difflib.unified_diff(lines_a, lines_b, fromfile=name_a, tofile=name_b, lineterm='')
    return '\n'.join(diff)


def compute_diff_blocks(text_a, text_b, context=2):
    lines_a = text_a.splitlines()
    lines_b = text_b.splitlines()
    sm = difflib.SequenceMatcher(None, lines_a, lines_b, autojunk=False)
    blocks = []

    def numbered(lines, start_idx):
        # start_idx is 0-based index into the source list; line numbers shown are 1-based.
        return [(start_idx + offset + 1, line) for offset, line in enumerate(lines)]

    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        blocks.append((tag, numbered(lines_a[i1:i2], i1), numbered(lines_b[j1:j2], j1)))

    return blocks


def diff_summary_counts(blocks):
    removed = 0
    added = 0
    changed = 0

    for tag, a, b in blocks:
        if tag == 'delete':
            removed += len(a)
        elif tag == 'insert':
            added += len(b)
        elif tag == 'replace':
            changed += max(len(a), len(b))

    return removed, added, changed


def render_diff_blocks_html(blocks, name_a, name_b):
    rows_html = []

    for tag, a, b in blocks:
        if tag == 'equal':
            for (num_a, line_a), (num_b, line_b) in zip(a, b):
                rows_html.append(
                    f'<tr>'
                    f'<td class="diff-linenum">{num_a}</td>'
                    f'<td class="diff-cell-equal">{escape(line_a)}</td>'
                    f'<td class="diff-linenum">{num_b}</td>'
                    f'<td class="diff-cell-equal">{escape(line_b)}</td>'
                    f'</tr>'
                )
        elif tag == 'collapsed':
            count = a
            rows_html.append(
                f'<tr><td class="diff-cell-collapsed" colspan="4">... {count} unchanged line(s) ...</td></tr>'
            )
        elif tag == 'delete':
            for num_a, line_a in a:
                rows_html.append(
                    f'<tr>'
                    f'<td class="diff-linenum">{num_a}</td>'
                    f'<td class="diff-cell-removed">{escape(line_a)}</td>'
                    f'<td class="diff-linenum"></td>'
                    f'<td class="diff-cell-empty"></td>'
                    f'</tr>'
                )
        elif tag == 'insert':
            for num_b, line_b in b:
                rows_html.append(
                    f'<tr>'
                    f'<td class="diff-linenum"></td>'
                    f'<td class="diff-cell-empty"></td>'
                    f'<td class="diff-linenum">{num_b}</td>'
                    f'<td class="diff-cell-added">{escape(line_b)}</td>'
                    f'</tr>'
                )
        elif tag == 'replace':
            max_len = max(len(a), len(b))
            for idx in range(max_len):
                if idx < len(a):
                    num_a, line_a = a[idx]
                    left = f'<td class="diff-linenum">{num_a}</td><td class="diff-cell-removed">{escape(line_a)}</td>'
                else:
                    left = '<td class="diff-linenum"></td><td class="diff-cell-empty"></td>'

                if idx < len(b):
                    num_b, line_b = b[idx]
                    right = f'<td class="diff-linenum">{num_b}</td><td class="diff-cell-added">{escape(line_b)}</td>'
                else:
                    right = '<td class="diff-linenum"></td><td class="diff-cell-empty"></td>'

                rows_html.append(f'<tr>{left}{right}</tr>')

    return f"""
    <div class="kash-table-wrap" style="max-height: 520px;">
        <table class="diff-side-by-side">
            <thead>
                <tr>
                    <th class="diff-linenum-header">#</th>
                    <th>{escape(name_a)}</th>
                    <th class="diff-linenum-header">#</th>
                    <th>{escape(name_b)}</th>
                </tr>
            </thead>
            <tbody>
                {''.join(rows_html)}
            </tbody>
        </table>
    </div>
    """

def render_duplicate_code_diff(report_row, duplicate_df, overview_df, raw_fex_content, selected_name, selected_path):
    dup_type = str(report_row.get('Duplicate Type', '') or '')

    if dup_type not in {'Exact Duplicate', 'Near Duplicate'}:
        return

    st.markdown("---")
    st.subheader("Compare Code With Matched Duplicate")

    match_options = find_duplicate_match_options(duplicate_df, overview_df, selected_name, selected_path)

    if not match_options:
        st.info("This report is flagged as a duplicate, but no comparable matched file could be located.")
        return

    labels = [label for label, _, _ in match_options]
    chosen_label = st.selectbox(
        "Compare against",
        labels,
        key=f"dup_compare_choice_{re.sub(r'[^A-Za-z0-9_]+', '_', selected_name + selected_path)}",
    )

    chosen = next((m for m in match_options if m[0] == chosen_label), None)
    if not chosen:
        return

    _, match_name, match_path = chosen

    text_a = get_raw_fex_text(raw_fex_content, selected_path, selected_name)
    text_b = get_raw_fex_text(raw_fex_content, match_path, match_name)

    if not text_a or not text_b:
        st.warning("Raw FEX text for one or both files is not available for comparison. Re-run analysis if this file was uploaded in a prior session.")
        return

    if text_a == text_b:
        st.success(f"'{selected_name}' and '{match_name}' are byte-for-byte identical.")
        return

    blocks = compute_diff_blocks(text_a, text_b, context=2)
    removed, added, changed = diff_summary_counts(blocks)

    if removed == 0 and added == 0 and changed == 0:
        st.success(f"'{selected_name}' and '{match_name}' have no line-level differences.")
        return

    c1, c2, c3 = st.columns(3)
    c1.metric(f"Only in {selected_name}", removed)
    c2.metric(f"Only in {match_name}", added)
    c3.metric("Changed line pairs", changed)

    st.caption("🔴 Red = only in the currently selected report. 🟢 Green = only in the compared file. White = identical in both.")

    diff_html = render_diff_blocks_html(blocks, selected_name, match_name)
    st.markdown(diff_html, unsafe_allow_html=True)
    
def build_single_report_excel(report_row, field_df, formula_df, filter_df, join_df, source_df, lineage_df, mapping_df, final_df, drilldown_df, build_plan_df):
    """Builds a small multi-sheet Excel workbook containing only the selected
    report's data, so a user can download one report at a time instead of
    the full combined workbook."""
    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        overview_df = pd.DataFrame([report_row]).astype(str)
        overview_df.to_excel(writer, sheet_name='Overview', index=False)

        sheets = [
            ('Field Inventory', field_df),
            ('Sources', source_df),
            ('HOLD Lineage', lineage_df),
            ('Formulas', formula_df),
            ('Filters', filter_df),
            ('Joins', join_df),
            ('Mapping Tables', mapping_df),
            ('Final Dataset', final_df),
            ('Drilldowns', drilldown_df),
            ('Build Plan', build_plan_df),
        ]

        for sheet_name, df in sheets:
            if df is not None and not df.empty:
                safe_df = df.astype(str)
                safe_df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

    output.seek(0)
    return output    

def filter_report_df(df, fex_name, file_path):
    if df is None or df.empty:
        return pd.DataFrame()

    name_col = 'File Name' if 'File Name' in df.columns else 'FEX Name'
    if name_col not in df.columns or 'File Path' not in df.columns:
        return pd.DataFrame()

    return df[(df[name_col] == fex_name) & (df['File Path'] == file_path)]


def filter_validation_df(df, fex_name):
    if df is None or df.empty:
        return pd.DataFrame()

    name_col = 'File Name' if 'File Name' in df.columns else 'FEX Name'
    if name_col not in df.columns:
        return pd.DataFrame()

    return df[df[name_col] == fex_name]


def badge(label, level=''):
    css_class = 'badge'
    if level:
        css_class += f' badge-{level}'
    return f'<span class="{css_class}">{label}</span>'


def report_complexity_score(report_row):
    score = 0
    score += int(report_row.get('Join / Match Count', 0) or 0) * 2
    score += int(report_row.get('Formula Count', 0) or 0)
    score += int(report_row.get('Filter Count', 0) or 0)
    score += int(report_row.get('HOLD Lineage Steps', 0) or 0)
    if report_row.get('Duplicate Type') in {'Near Duplicate', 'Same Data Source'}:
        score += 2
    if str(report_row.get('Output Format', '')).upper().find('GRAPH') >= 0:
        score += 2
    return score


def complexity_label(score):
    if score >= 18:
        return 'High', 'high'
    if score >= 9:
        return 'Medium', 'medium'
    return 'Low', 'low'


def estimate_effort_saving(duplicate_counts_map, total_reports):
    if not total_reports:
        return "0%"

    duplicate_saving = (
        duplicate_counts_map.get('exact', 0) * 8
        + duplicate_counts_map.get('near', 0) * 5
        + duplicate_counts_map.get('same_source', 0) * 3
    )
    base_saving = min(22, total_reports * 2)
    return f"{min(65, base_saving + duplicate_saving)}%"


def render_overview_story(report_row, source_df, final_df):
    source_names = ', '.join(source_df['Source Name'].dropna().astype(str).tolist()) if not source_df.empty else report_row.get('Source Tables', '')
    final_text = final_df.iloc[0].to_dict() if not final_df.empty else {}
    dataset_name = final_text.get('Recommended Final Dataset', report_row.get('Recommended Final Dataset', 'final Cognos dataset'))
    visual = final_text.get('Suggested Cognos Visual / Output', report_row.get('Output Format', 'Cognos report/list'))
    shape = final_text.get('Dataset Shape', 'Cognos-ready dataset')
    score = report_complexity_score(report_row)
    risk, risk_level = complexity_label(score)

    with st.container(border=True):
        st.markdown("#### What this WebFOCUS report does")
        st.markdown(
            f"This report starts from **{source_names or 'detected WebFOCUS sources'}**, "
            f"applies WebFOCUS logic such as filters, formulas, joins, HOLD tables, and grouping, "
            f"then produces **{visual}**."
        )
        st.markdown(
            f"**Cognos version:** build **{dataset_name}** as a **{shape}**, "
            f"then use it for the final report page or visual."
        )
        st.markdown(
            f"{badge('Build type: Cognos report')} "
            f"{badge(f'Risk: {risk}', risk_level)} "
            f"{badge(f'Complexity score: {score}')}",
            unsafe_allow_html=True,
        )


def source_role_text(source_name, sql_need):
    if sql_need == 'Recommended':
        return 'Include in SQL view'
    if sql_need == 'Optional':
        return 'Framework Manager/Data Module or SQL view'
    return 'Load in Framework Manager/Data Module'


def recommend_formula_layer(formula_type):
    text = str(formula_type or '')
    if 'Decode/Mapping' in text:
        return 'Mapping table, then SQL View / Framework Manager merge'
    if 'Date/Time' in text or 'Concatenation/String' in text or 'Conditional IF' in text:
        return 'SQL View or Framework Manager/Data Module calculated query item'
    if 'Arithmetic' in text:
        return 'SQL View / Framework Manager/Data Module calculated query item'
    if 'Parameter-based' in text:
        return 'Cognos prompt plus SQL/Framework Manager filter'
    return 'Review; likely SQL View or Framework Manager/Data Module'


def maybe_int(value):
    try:
        if value is None or str(value).strip() == '':
            return None
        return int(float(value))
    except (TypeError, ValueError):
        return None


def unique_text_values(values):
    seen = set()
    result = []
    for value in values or []:
        text = str(value or '').strip()
        if not text or text.lower() in {'nan', 'none'}:
            continue
        key = text.upper()
        if key not in seen:
            seen.add(key)
            result.append(text)
    return result


def split_source_text(value):
    return unique_text_values(str(value or '').replace(';', ',').split(','))


def join_natural(items):
    items = [str(item).strip() for item in items if str(item).strip()]
    if not items:
        return ''
    if len(items) == 1:
        return items[0]
    if len(items) == 2:
        return f"{items[0]} and {items[1]}"
    return f"{', '.join(items[:-1])}, and {items[-1]}"


def source_business_phrase(sources):
    upper_sources = [source.upper() for source in sources]
    pieces = []

    if any('LEDGR' in source or 'LEDGER' in source for source in upper_sources):
        pieces.append('ledger')
    if any('BUD' in source for source in upper_sources):
        pieces.append('budget')
    if any('OPENPO' in source or 'PO' in source for source in upper_sources):
        pieces.append('open PO')

    if pieces:
        prefix_parts = []
        if any('MONTH' in source for source in upper_sources):
            prefix_parts.append('monthly')
        if any(source.startswith('AX') for source in upper_sources):
            prefix_parts.append('AX')
        prefix = f"{' '.join(prefix_parts)} " if prefix_parts else ''
        return f"{prefix}{join_natural(pieces)} data"

    if sources:
        return f"data from {', '.join(sources[:6])}"

    return 'the detected report data'


def build_start_here_facts(report_row, source_df, lineage_df, join_df, final_df):
    report = str(report_row.get('FEX Name', report_row.get('Report', 'Selected report')) or 'Selected report')

    if source_df is not None and not source_df.empty and 'Source Name' in source_df:
        real_source_df = (
            source_df[source_df['Notes'] == 'Original report source']
            if 'Notes' in source_df.columns
            else source_df
        )
        # Fall back to the full source list only if every detected source
        # turned out to be an intermediate HOLD table (edge case safeguard).
        if real_source_df.empty:
            real_source_df = source_df
        sources = unique_text_values(real_source_df['Source Name'].dropna().astype(str).tolist())
    else:
        sources = split_source_text(report_row.get('Source Tables', ''))

    hold_count = maybe_int(report_row.get('HOLD Lineage Steps'))
    if hold_count is None:
        hold_count = (
            int((lineage_df['Output Role'].astype(str).str.contains('HOLD', case=False, na=False)).sum())
            if lineage_df is not None and not lineage_df.empty and 'Output Role' in lineage_df
            else 0
        )

    join_count = maybe_int(report_row.get('Join / Match Count'))
    if join_count is None:
        join_count = len(join_df) if join_df is not None else 0

    final_output = ''
    if final_df is not None and not final_df.empty and 'Recommended Final Dataset' in final_df:
        final_output = str(final_df.iloc[0]['Recommended Final Dataset'] or '').strip()
    if not final_output:
        final_output = str(report_row.get('Recommended Final Dataset', '') or '').strip()
    if not final_output:
        final_output = 'final report-ready dataset/table'

    sql_need = str(report_row.get('SQL View Need', 'Needs Review') or 'Needs Review')
    recommended_layer = 'SQL View' if sql_need == 'Recommended' else 'Framework Manager/Data Module'
    if sql_need == 'Optional':
        recommended_layer = 'SQL View or Framework Manager/Data Module'

    return {
        'report': report,
        'sources': sources,
        'final_output': final_output,
        'hold_count': hold_count,
        'join_count': join_count,
        'formula_count': maybe_int(report_row.get('Formula Count')) or 0,
        'filter_count': maybe_int(report_row.get('Filter Count')) or 0,
        'duplicate_status': str(report_row.get('Duplicate Type', 'Unique') or 'Unique'),
        'recommended_layer': recommended_layer,
        'sql_need': sql_need,
        'complexity': complexity_label(report_complexity_score(report_row))[0],
    }


def fallback_report_summary(facts):
    source_count = len(facts['sources'])
    source_word = 'source' if source_count == 1 else 'sources'
    hold_word = 'step' if facts['hold_count'] == 1 else 'steps'
    formula_word = 'formula' if facts['formula_count'] == 1 else 'formulas'
    filter_word = 'filter' if facts['filter_count'] == 1 else 'filters'
    join_word = 'operation' if facts['join_count'] == 1 else 'operations'
    duplicate_status = facts['duplicate_status']
    source_text = ', '.join(facts['sources']) if facts['sources'] else 'the detected WebFOCUS sources'

    if duplicate_status in {'Exact Duplicate', 'Near Duplicate', 'Same Data Source'}:
        duplicate_text = (
            f"{duplicate_status} - rebuild once where possible and reuse the same "
            "Cognos approach for this group."
        )
    else:
        duplicate_text = "Unique - no duplicate group was detected for this report."

    if facts['recommended_layer'] == 'SQL View':
        direction = (
            "Build the preparation logic in a SQL view first, then model it in Framework "
            "Manager/Data Module for Cognos and validate row counts/totals."
        )
    elif facts['recommended_layer'] == 'SQL View or Framework Manager/Data Module':
        direction = (
            "Use a SQL view or Framework Manager/Data Module for the preparation logic, "
            "then model it in Cognos and validate row counts/totals."
        )
    else:
        direction = (
            "Build the preparation logic in Framework Manager/Data Module, then create the "
            "Cognos report and validate row counts/totals."
        )

    return {
        'purpose': (
            f"Reads {source_text}, applies WebFOCUS preparation logic, "
            f"and produces {facts['final_output']}."
        ),
        'complexity': (
            f"{facts['complexity']} - {source_count} {source_word}, "
            f"{facts['hold_count']} HOLD/intermediate {hold_word}, "
            f"{facts['formula_count']} {formula_word}, "
            f"{facts['filter_count']} {filter_word}, and "
            f"{facts['join_count']} JOIN/MATCH {join_word}."
        ),
        'duplicate_status': duplicate_text,
        'recommended_direction': direction,
    }


def render_start_here(report_row, source_df, lineage_df, join_df, final_df):
    facts = build_start_here_facts(report_row, source_df, lineage_df, join_df, final_df)
    summary = fallback_report_summary(facts)
    source_text = ', '.join(facts['sources']) if facts['sources'] else 'No source tables detected'

    with st.container(border=True):
        st.markdown(f"**Report:** {escape(facts['report'])}")
        st.markdown(f"**Purpose:** {escape(summary['purpose'])}")
        st.markdown(f"**Main sources:** {escape(source_text)}")
        st.markdown(f"**Final output:** {escape(facts['final_output'])}")
        st.markdown(f"**Migration complexity:** {escape(summary['complexity'])}")
        st.markdown(f"**Duplicate status:** {escape(summary['duplicate_status'])}")
        st.markdown(f"**Recommended direction:** {escape(summary['recommended_direction'])}")


def render_source_files_tables(source_df, report_row):
    if source_df is None or source_df.empty:
        st.info("No source files/tables were detected.")
        return

    df = source_df.copy()
    if 'Notes' in df.columns:
        real_df = df[df['Notes'] == 'Original report source']
        # Safeguard: if everything got flagged as HOLD-like, fall back to
        # showing the full list rather than an empty table.
        if not real_df.empty:
            df = real_df
    df['Source'] = df['Source Name']
    df['How detected'] = df['FILEDEF / Raw Source Reference'].apply(lambda x: 'FILEDEF' if str(x).strip() else 'Detected from TABLE FILE / DEFINE FILE')
    df['Type'] = df['Detected Source Type']
    df['Role in Report'] = df['Notes'].replace({'Original report source': 'Source table/file'})
    df['Cognos Action'] = df['Source Name'].apply(lambda x: source_role_text(x, report_row.get('SQL View Need')))
    show_table_or_info(df[['Source', 'How detected', 'Type', 'Role in Report', 'Cognos Action']], "No source files/tables were detected.")
    st.caption("Intermediate HOLD tables created by this report are listed separately below, not in this source list.")


def render_hold_lineage_inspector(lineage_df, report_row):
    st.write("WebFOCUS HOLD files are temporary datasets. In Cognos migration, these usually become SQL CTEs, SQL view steps, or Framework Manager/Data Module staging query subjects.")
    if report_row.get('SQL View Need') == 'Recommended':
        st.info("Recommended: convert these HOLD steps into SQL CTEs inside one SQL view.")

    if lineage_df is None or lineage_df.empty:
        st.info("No HOLD lineage was detected.")
        return

    df = lineage_df.copy()
    hold_df = df[df['Output Table'].astype(str).str.len() > 0].copy()
    if hold_df.empty:
        hold_df = df.copy()

    hold_df['Order'] = hold_df['Step #']
    hold_df['Temporary Dataset'] = hold_df['Output Table'].replace('', 'Final output step')
    hold_df['What happens here'] = hold_df['What Happens In This Step']
    hold_df['Migration equivalent'] = hold_df['Output Role'].apply(
        lambda role: 'SQL CTE inside SQL view' if report_row.get('SQL View Need') == 'Recommended' else 'Framework Manager/Data Module staging query subject'
    )
    show_table_or_info(hold_df[['Order', 'Temporary Dataset', 'What happens here', 'Migration equivalent']], "No HOLD lineage was detected.")


def render_joins_matches_inspector(join_df, report_row):
    if join_df is None or join_df.empty:
        st.success("No JOIN/MATCH statements were detected.")
        return

    st.write("JOIN rows can duplicate data if keys are not unique. MATCH FILE rows are higher risk because WebFOCUS MATCH behavior may not equal a normal SQL join.")
    st.write(f"This report has {len(join_df)} JOIN/MATCH operation(s). Because joins can duplicate rows, validate row counts after each join.")
    df = join_df.copy()
    df['Risk'] = df['Mapping Type'].apply(lambda x: 'High' if str(x).upper() == 'MATCH FILE' else 'Medium')
    df['Join / Match'] = df['Mapping Type']
    df['Left Side'] = df['From Table'].astype(str) + '[' + df['From Key'].astype(str) + ']'
    df['Right Side'] = df['To Table'].astype(str) + '[' + df['To Key'].astype(str) + ']'
    if report_row.get('SQL View Need') == 'Recommended':
        df['Recommended Cognos/SQL Action'] = 'Implement inside SQL view first. Use a Framework Manager relationship only if the joined table remains separate.'
    else:
        df['Recommended Cognos/SQL Action'] = df['Suggested Cognos Action']
    df['Review'] = df['Review Reason']
    show_table_or_info(df[['Risk', 'Join / Match', 'Left Side', 'Right Side', 'Recommended Cognos/SQL Action', 'Review']], "No JOIN/MATCH statements were detected.")


def render_filters_parameters_inspector(filter_df):
    if filter_df is None or filter_df.empty:
        st.success("No filters or parameters were detected.")
        return

    df = filter_df.copy()
    df['Filter / Parameter'] = df['Filter Type'].astype(str) + ': ' + df['Expression'].astype(str)
    df['Meaning'] = df['Fields Used'].apply(lambda x: f"Limits rows using {x}" if str(x).strip() else 'Runtime/user input or report condition')
    df['Cognos Treatment'] = df.apply(short_filter_treatment, axis=1)
    field_col = 'Fields Used' if 'Fields Used' in df.columns else 'Fields'
    show_table_or_info(df[['Filter / Parameter', field_col, 'Meaning', 'Cognos Treatment']], "No filters or parameters were detected.")


def short_filter_treatment(row):
    filter_type = str(row.get('Filter Type', '')).upper()
    expression = str(row.get('Expression', ''))
    parameter = str(row.get('Parameter', ''))

    if filter_type == 'PARAMETER' or parameter.strip():
        return 'Cognos prompt'
    if filter_type == 'WHERE':
        return 'Fixed SQL filter' if 'SQL' in str(row.get('Cognos Guidance', '')).upper() else 'Framework Manager/Data Module row filter'
    if filter_type == 'IF':
        return 'SQL WHERE / Framework Manager/Data Module row filter'
    if '&' in expression:
        return 'Cognos prompt'
    return 'Review filter placement'


def render_calculations_mappings_inspector(formula_df, mapping_df):
    calc_tab, mapping_tab = st.tabs(["Calculations", "Mappings"])
    with calc_tab:
        if formula_df is None or formula_df.empty:
            st.info("No DEFINE/COMPUTE calculations were detected.")
        else:
            df = formula_df.copy()
            st.subheader("Calculation Groups")
            group_counts = {
                'Date parsing': int(df['Formula Type'].astype(str).str.contains('Date/Time', case=False, na=False).sum()),
                'String / concatenation': int(df['Formula Type'].astype(str).str.contains('Concatenation/String', case=False, na=False).sum()),
                'Conditional IF logic': int(df['Formula Type'].astype(str).str.contains('Conditional IF', case=False, na=False).sum()),
                'Numeric / arithmetic calculations': int(df['Formula Type'].astype(str).str.contains('Arithmetic', case=False, na=False).sum()),
                'Mappings / DECODE': int(df['Formula Type'].astype(str).str.contains('Decode/Mapping', case=False, na=False).sum()),
                'Unknown / review': int(df['Formula Type'].astype(str).str.contains('Direct/Other', case=False, na=False).sum()),
            }
            calculation_group_df = pd.DataFrame(
               [{'Calculation Group': k, 'Count': v} for k, v in group_counts.items()]
            )

            show_table_or_info(
                 calculation_group_df,
                 "No calculation groups were detected."
            )
            df['WebFOCUS Calculation'] = df['Formula Source'].astype(str) + ' ' + df['Field Name'].astype(str)
            df['Purpose'] = df['Formula Type']
            df['Recommended layer'] = df['Formula Type'].apply(recommend_formula_layer)
            for group_name, pattern in [
                ('Date parsing', 'Date/Time'),
                ('String / concatenation', 'Concatenation/String'),
                ('Conditional IF logic', 'Conditional IF'),
                ('Numeric / arithmetic calculations', 'Arithmetic'),
                ('Mappings / DECODE', 'Decode/Mapping'),
                ('Unknown / review', 'Direct/Other'),
            ]:
                group_df = df[df['Formula Type'].astype(str).str.contains(pattern, case=False, na=False)]
                if not group_df.empty:
                    with st.expander(f"{group_name} ({len(group_df)})", expanded=False):
                        group_df['Review Needed'] = group_df['Needs Manual Review'].replace('', 'Yes') if 'Needs Manual Review' in group_df.columns else 'Yes'
                        cols = [
                            'WebFOCUS Calculation',
                            'Purpose',
                            'Raw WebFOCUS Formula',
                            'Raw Columns Used',
                            'Recommended layer',
                            'Review Needed',
                        ]
                        show_table_or_info(group_df[[col for col in cols if col in group_df.columns]], "No calculations were detected.")
    with mapping_tab:
        render_mapping_cards(mapping_df)

    with st.expander("Show raw DEFINE / COMPUTE formula evidence"):
        cols = [
            col for col in [
                'Formula Source',
                'Field Name',
                'Formula Type',
                'Raw WebFOCUS Formula',
                'Raw Columns Used',
                'Needs Manual Review',
            ]
            if formula_df is not None and not formula_df.empty and col in formula_df.columns
        ]
        show_table_or_info(
    formula_df[cols] if cols else formula_df,
    "No formula evidence found.",
    large=True,
)


def sql_view_reasons(report_row, source_df, lineage_df, join_df):
    reasons = []
    source_count = len(source_df) if source_df is not None else 0
    hold_count = int((lineage_df['Output Role'].astype(str).str.contains('HOLD', case=False, na=False)).sum()) if lineage_df is not None and not lineage_df.empty and 'Output Role' in lineage_df else 0
    join_count = len(join_df) if join_df is not None else 0
    match_count = int((join_df['Mapping Type'].astype(str).eq('MATCH FILE')).sum()) if join_df is not None and not join_df.empty and 'Mapping Type' in join_df else 0

    if source_count >= 3:
        reasons.append(f"Multiple source tables detected: {source_count}")
    if hold_count >= 1:
        reasons.append(f"HOLD/intermediate steps detected: {hold_count}")
    if join_count >= 1:
        reasons.append(f"JOIN/MATCH logic detected: {join_count}")
    if match_count >= 1:
        reasons.append(f"MATCH FILE logic detected: {match_count}")
    if report_row.get('SQL View Need') == 'Not Needed':
        reasons.append("Simple enough for Framework Manager/Data Module based on current parser rules")
    return reasons


def sql_identifier(value, fallback='cte_step'):
    text = str(value or '').strip()
    text = re.sub(r'[^A-Za-z0-9_]+', '_', text)
    text = text.strip('_').lower()
    if not text:
        text = fallback
    if re.match(r'^\d', text):
        text = f"t_{text}"
    return text


def sql_source_cte_name(source_name):
    return f"src_{sql_identifier(source_name, 'source')}"


def sql_view_name(report_row):
    return f"vw_{sql_identifier(Path(str(report_row.get('FEX Name', 'report'))).stem, 'report')}_base"


def build_sql_cte_outline(source_df, lineage_df, join_df, report_row):
    lines = []
    sources = source_df['Source Name'].dropna().astype(str).tolist() if source_df is not None and not source_df.empty else []
    hold_outputs = []
    if lineage_df is not None and not lineage_df.empty and 'Output Table' in lineage_df:
        hold_outputs = [x for x in lineage_df['Output Table'].dropna().astype(str).tolist() if x]

    lines.extend(sources or ['Detected source tables'])
    if hold_outputs:
        lines.append('   ↓')
        lines.extend(hold_outputs)
    if join_df is not None and not join_df.empty:
        lines.append('   ↓')
        lines.append('JOIN / MATCH logic')
    lines.append('   ↓')
    lines.append(sql_view_name(report_row))
    return '\n'.join(lines)


def render_sql_view_recommendation_inspector(report_row, source_df, lineage_df, join_df, filter_df=None, formula_df=None, field_df=None):
    need = report_row.get('SQL View Need', 'Needs Review')
    view_name = sql_view_name(report_row)
    st.markdown(f"### SQL View Recommended: {'Yes' if need == 'Recommended' else 'Optional' if need == 'Optional' else 'No'}")
    for reason in sql_view_reasons(report_row, source_df, lineage_df, join_df):
        st.write(f"- {reason}")
    st.write(f"**Recommended SQL View Name:** `{view_name}`")
    st.write("**Preferred:** create this as a database view, then model it in Framework Manager or a Cognos Data Module.")
    st.write("**Alternative:** paste the final reviewed SELECT query into a Cognos Data Module's custom SQL, or a Framework Manager data source query, if database permissions do not allow creating views.")

    st.subheader("SQL CTE Outline")
    st.code(build_sql_cte_outline(source_df, lineage_df, join_df, report_row), language='text')

    st.markdown("---")
    st.subheader("AI-Powered SQL View")
    st.caption(
        "Uses the parsed facts above (real sources, HOLD steps, joins, filters, formulas) to write "
        "actual column-level SQL instead of SELECT *. Still requires human review before production use."
    )

    table_db_mapping_df = st.session_state.get("analysis_result", {}).get("display_tables", {}).get("table_db_mapping")
    has_mas_acx = table_db_mapping_df is not None and not table_db_mapping_df.empty

    if has_mas_acx:
        st.success("✅ **Metadata ZIP (.mas/.acx) was uploaded.** Table names resolved from it are treated as confirmed real database tables.")
        st.caption("Note: some tables in this report may still fall back to alias names if they weren't found in the uploaded MAS/ACX metadata — check for `-- REVIEW` comments in the SQL below.")
    else:
        st.markdown(
            "<div style='background:#FFF3CD;border:1px solid #FFECB5;border-radius:8px;padding:12px 16px;'>"
            "<strong>⚠️ No Metadata ZIP (.mas/.acx) was uploaded.</strong><br>"
            "This SQL uses <strong>WebFOCUS alias/synonym names</strong> (e.g. SASDATA, AXOPER), "
            "<strong>not confirmed real database table names</strong>. "
            "This SQL will <strong>not run as-is</strong> until every table reference is manually verified "
            "and replaced with the real physical table name. Upload the Metadata ZIP and regenerate for "
            "verified names."
            "</div>",
            unsafe_allow_html=True,
        )

    client = get_openai_client()
    if client is None:
        st.warning(
            "OpenAI is not configured. Add `OPENAI_API_KEY` to `.streamlit/secrets.toml` to enable SQL view generation."
        )
    else:
        sql_report_key = re.sub(
            r'[^A-Za-z0-9_]+', '_',
            str(report_row.get('FEX Name', '')) + '_' + str(report_row.get('File Path', ''))
        )
        ai_sql_cache = st.session_state.setdefault('ai_sql_view_cache', {})

        gen_clicked = st.button(
            "Generate AI-powered SQL view" if sql_report_key not in ai_sql_cache else "Regenerate AI-powered SQL view",
            key=f"ai_sql_view_btn_{sql_report_key}",
        )

        if gen_clicked:
            sql_payload = build_ai_sql_view_payload(
                report_row, source_df, lineage_df, join_df, filter_df, formula_df, None, field_df,
                table_db_mapping_df=table_db_mapping_df,
            )
            with st.spinner("Calling OpenAI to draft the SQL view..."):
                ai_sql, ai_sql_error = generate_ai_sql_view(client, "gpt-4o-mini", sql_payload)
            ai_sql_cache[sql_report_key] = (ai_sql, ai_sql_error)

        cached_sql = ai_sql_cache.get(sql_report_key)
        if cached_sql:
            ai_sql, ai_sql_error = cached_sql
            if ai_sql_error:
                st.error(ai_sql_error)
            else:
                st.code(ai_sql, language='sql')
                st.download_button(
                    "Download SQL view (.sql)",
                    data=ai_sql.encode('utf-8'),
                    file_name=f"{view_name}.sql",
                    mime="text/sql",
                    key=f"download_ai_sql_{sql_report_key}",
                )
        else:
            st.info("Click the button above to generate the SQL view.")

    with st.expander("Show source trace"):
        trace_df = pd.DataFrame({
            'Trace Section': ['Sources', 'HOLD / lineage steps', 'JOIN / MATCH rows', 'Filters / parameters', 'Formulas / calculations'],
            'Detected Count': [
                len(source_df) if source_df is not None else 0,
                len(lineage_df) if lineage_df is not None else 0,
                len(join_df) if join_df is not None else 0,
                len(filter_df) if filter_df is not None else 0,
                len(formula_df) if formula_df is not None else 0,
            ],
        })
        st.dataframe(trace_df, use_container_width=True, hide_index=True)


        
def render_cognos_approach_inspector(report_row, build_plan_df):
    need = report_row.get('SQL View Need', 'Needs Review')
    view_name = f"vw_{Path(str(report_row.get('FEX Name', 'report'))).stem.lower()}_base"
    if need == 'Recommended':
        st.write("**Recommended approach:** SQL View -> Framework Manager / Data Module -> Cognos Report")
        steps = [
            f"Create SQL View: ask the DBA/migration developer to create `{view_name}`.",
            f"Model `{view_name}` in Framework Manager or a Data Module.",
            "Clean up in Framework Manager/Data Module: set data types, rename query items, remove unused columns, keep report-needed fields.",
            "Build the report page: create the final list, crosstab, or chart.",
            "Validate: compare row counts, totals, filters, and join results against WebFOCUS output.",
        ]
    else:
        st.write("**Recommended approach:** Framework Manager / Data Module -> Cognos Report")
        steps = [
            "Connect Framework Manager/Data Module to the source file/table.",
            "Use Framework Manager/Data Module for filters, type cleanup, and report-specific transformations.",
            "Create Cognos calculations only if final totals/KPIs need summarized values.",
            "Build the final list, crosstab, or chart.",
            "Validate row counts, totals, filters, and output shape against WebFOCUS.",
        ]

    for idx, step in enumerate(steps, start=1):
        st.markdown(f"**Step {idx}**")
        st.write(step)

    if build_plan_df is not None and not build_plan_df.empty:
        with st.expander("Rule-based build steps"):
            show_table_or_info(build_plan_df, "No build steps found.")

def render_drilldown_output_inspector(report_row, final_df, drilldown_df, lineage_df=None):
    st.write("This section describes the final output detected from the WebFOCUS report.")
    final_dataset = final_df.iloc[0]['Recommended Final Dataset'] if final_df is not None and not final_df.empty else report_row.get('Recommended Final Dataset', 'Review final dataset')
    output_recommendation = 'Matrix visual / table visual'
    if str(report_row.get('Output Format', '')).upper().find('GRAPH') >= 0:
        output_recommendation = 'Chart visual'

    st.markdown(f"**Final Dataset Detected:** `{final_dataset}`")
    st.markdown(f"**Cognos Output Recommendation:** {output_recommendation}")
    hold_outputs = []
    if lineage_df is not None and not lineage_df.empty and 'Output Table' in lineage_df.columns:
        hold_outputs = [x for x in lineage_df['Output Table'].dropna().astype(str).tolist() if x]
    if hold_outputs:
        st.markdown("**Intermediate HOLD outputs:**")
        st.write(', '.join(dict.fromkeys(hold_outputs)))

    if final_df is not None and not final_df.empty:
        with st.expander("View final dataset details"):
            render_visual_plan(final_df)
    else:
        st.info(report_row.get('Output Format', 'No final output format was detected.'))

    st.subheader("Drilldown / URL Links")
    if drilldown_df is None or drilldown_df.empty:
        st.info("No drilldown detected.")
    else:
        df = drilldown_df.copy()
        df['WebFOCUS drilldown'] = df['raw_drilldown']
        df['Target report'] = df['target_report'].replace('', 'Review target')
        df['Parameter'] = df.apply(
            lambda row: row['parameter'] if str(row.get('parameter', '')).strip() else row.get('fields', ''),
            axis=1,
        )
        df['Cognos replacement'] = df['cognos_replacement']
        show_table_or_info(df[['WebFOCUS drilldown', 'Target report', 'Parameter', 'Cognos replacement']], "No drilldown detected.")

def render_data_module_inspector(report_row):
    sql_needed = report_row.get('SQL View Need') == 'Recommended'
    if sql_needed:
        st.write("Framework Manager/Data Module modeling should be light cleanup only.")
        st.info(f"Because a SQL View is recommended, Framework Manager/Data Module should not rebuild the full WebFOCUS logic. Model it only after `{sql_view_name(report_row)}` is created and validated.")
        steps = [
            f"Connect Data Module / Framework Manager data source to {sql_view_name(report_row)}",
            "Confirm preview loads",
            "Set data types",
            "Rename business-facing query items",
            "Remove technical/helper columns",
            "Keep final report fields",
            "Publish/expose to Cognos reports",
        ]
    else:
        st.write("Framework Manager/Data Module should recreate the WebFOCUS staging flow.")
        steps = [
            "Load source files/tables",
            "Create staging query subjects",
            "Apply row filters",
            "Create calculated query items",
            "Set up relationships between source queries",
            "Create final query subject/report dataset",
        ]

    for idx, step in enumerate(steps, start=1):
        st.checkbox(step, value=False, key=f"dm_inspector_{report_row['FEX Name']}_{idx}")



def render_data_model_inspector(report_row, source_df, mapping_df):
    if report_row.get('SQL View Need') == 'Recommended' and (mapping_df is None or mapping_df.empty):
        st.write("Recommended data model: Simple flat model from SQL view.")
        st.info(
            f"Use this model: main table `{sql_view_name(report_row)}`. "
            "Recommended for first build: start with one flat table from the SQL view. "
            "Add lookup tables only after validation."
        )
        optional_rows = [
            {'Table': sql_view_name(report_row), 'Role': 'Main report table'},
            {'Table': 'Date', 'Role': 'Optional date dimension if date filtering is needed'},
            {'Table': 'Product / Customer lookup', 'Role': 'Optional slicer/lookup tables if users need reusable dimensions'},
        ]
        st.dataframe(pd.DataFrame(optional_rows), use_container_width=True, hide_index=True)
        return

    st.write("Recommended data model: Fact/report table plus lookup tables where useful.")
    rows = [{
        'Table': report_row.get('Recommended Final Dataset') or 'final_report',
        'Role': 'Fact/report table',
    }]
    if mapping_df is not None and not mapping_df.empty:
        for table_name in sorted(mapping_df['Mapping Table Name'].dropna().unique()):
            rows.append({'Table': table_name, 'Role': 'Lookup/mapping table'})
    if source_df is not None and not source_df.empty and len(source_df) > 2:
        rows.append({'Table': 'Date', 'Role': 'Optional date dimension if date slicing is needed'})
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


def infer_calc_field_name(field_name):
    clean = str(field_name or '').replace('_', ' ').title().strip()
    if clean:
        return f"Total {clean}"
    return "Total Value"


def render_calc_inspector(report_row, field_df, formula_df):
    st.info("Do not move all WebFOCUS DEFINE formulas into report-level calculations. Most row-level formulas should be handled in the SQL View or Framework Manager/Data Module.")
    candidates = []
    if field_df is not None and not field_df.empty:
        value_words = r'ORDER_VALUE|DOLLARS?_ORDERED|DOLLARS?_INVOICED|POUNDS?_ORDERED|POUNDS?_INVOICED|QTY|QUANTITY|RUN_HRS|DELAY_HRS|AMOUNT|COST|SALES|PRICE|COUNT|TOTAL'
        helper_words = r'DATE|DAY|OFFSET|RANK|SEQ|SEQUENCE|KEY|CODE|STATUS|FLAG|YEAR|MONTH|WEEK'
        field_col = 'Fields Used'
        if field_col in field_df.columns:
            for field in field_df[field_col].dropna().astype(str).unique():
                if re.search(value_words, field, re.IGNORECASE) and not re.search(helper_words, field, re.IGNORECASE):
                    candidates.append(field)

    if not candidates:
        st.info("No obvious summary calculations detected. Add a Cognos calculation only for totals, KPIs, or interactive aggregations.")
        return

    final_table = sql_identifier(report_row.get('Recommended Final Dataset') or 'final_report', 'final_report')
    for field in list(dict.fromkeys(candidates))[:8]:
        measure_name = infer_calc_field_name(field)
        st.code(f"-- Cognos calculation: {measure_name}\ntotal([{final_table}].[{field}])", language='sql')


def render_data_model_and_calc_inspector(report_row, source_df, mapping_df, field_df, formula_df):
    st.subheader("Data Model")
    render_data_model_inspector(report_row, source_df, mapping_df)
    st.subheader("Calculations")
    render_calc_inspector(report_row, field_df, formula_df)


def render_visuals_inspector(report_row, final_df, field_df):
    output_text = str(report_row.get('Output Format', '')).upper()
    if 'GRAPH' in output_text:
        visual_type = 'Cognos chart (Graph)'
    elif 'VBRSTK' in output_text:
        visual_type = 'Stacked column chart'
    elif 'PIE' in output_text:
        visual_type = 'Pie or donut chart'
    elif report_row.get('Type') == 'Matrix / pivot report':
        visual_type = 'Crosstab'
    else:
        visual_type = 'List report'

    st.markdown(f"**Recommended Cognos Report Page:** {Path(str(report_row.get('FEX Name', 'report'))).stem.replace('_', ' ').title()}")
    st.markdown(f"**Visual 1:** {visual_type}")
    rows = []
    values = []
    if field_df is not None and not field_df.empty:
        shown_df = field_df[field_df.get('Is Field Shown In Report', '') == 'Yes'] if 'Is Field Shown In Report' in field_df else field_df
        for field in shown_df.get('Fields Used', pd.Series(dtype=str)).dropna().astype(str):
            if re.search(r'DOLLAR|AMT|AMOUNT|QTY|POUND|HOUR|TOTAL|COUNT|PRICE', field, re.IGNORECASE):
                values.append(field)
            else:
                rows.append(field)
    rows = list(dict.fromkeys(rows))
    values = list(dict.fromkeys(values))
    st.write("**Place these fields:**")
    st.write(f"Rows: {', '.join(rows[:8]) or 'Confirm report grouping fields'}")
    st.write("Columns: ACROSS/RANK field if the WebFOCUS report used ACROSS logic")
    st.write(f"Values: {', '.join(values[:8]) or 'Confirm numeric totals from SUM fields'}")
    st.write("Prompts: Date fields, customer fields, product fields, order/status fields, and WebFOCUS parameters.")
    st.warning("Review: confirm which numeric fields should be summarized before creating final calculations.")


def render_detailed_report_summary(report_row, source_df, lineage_df, join_df, filter_df, formula_df):
    report = report_row.get('FEX Name', 'This report')
    sources = source_df['Source Name'].dropna().astype(str).tolist() if source_df is not None and not source_df.empty else []
    source_sentence = ', '.join(sources) if sources else 'the detected source tables/files'
    hold_count = int((lineage_df['Output Role'].astype(str).str.contains('HOLD', case=False, na=False)).sum()) if lineage_df is not None and not lineage_df.empty and 'Output Role' in lineage_df else 0
    join_count = len(join_df) if join_df is not None else 0
    filter_count = len(filter_df) if filter_df is not None else 0
    formula_count = len(formula_df) if formula_df is not None else 0
    sql_need = report_row.get('SQL View Need', 'Needs Review')

    summary = (
        f"{report} uses {source_sentence}. The report creates {hold_count} detected HOLD/intermediate step(s), "
        f"uses {join_count} JOIN/MATCH operation(s), applies {filter_count} filter/parameter rule(s), "
        f"and contains {formula_count} calculation/mapping row(s). "
    )
    if sql_need == 'Recommended':
        summary += (
            "Because the report contains multi-step source preparation, the recommended migration approach is to create "
            "a SQL view that reproduces the WebFOCUS data preparation logic. Cognos, via Framework Manager/Data Module, "
            "should connect to the SQL view for light cleanup such as data types, naming, and final field selection. "
        )
    else:
        summary += (
            "Based on the detected logic, this can likely be migrated with Framework Manager/Data Module source shaping and a Cognos report. "
        )
    summary += "The main validation focus should be row counts, totals, filter behavior, join cardinality, and final output shape."

    st.write(summary)


def render_walkthrough(lineage_df):
    if lineage_df.empty:
        st.info("No lineage steps were detected for this report.")
        return

    for _, row in lineage_df.sort_values('Step #').iterrows():
        step_title = f"Step {row['Step #']} - {row['Step Type']}: {row['Input Table(s)']}"
        with st.expander(step_title, expanded=int(row['Step #']) <= 3):
            st.markdown(f"<div class='step-title'>{row['What Happens In This Step']}</div>", unsafe_allow_html=True)
            st.write(f"Cognos action: {row['Cognos Action'] or 'Review this step and rebuild it in Framework Manager/Data Module.'}")
            if row.get('Output Table'):
                st.write(f"Output table: {row['Output Table']} ({row['Output Role']})")
            if row.get('Cognos Table Suggestion'):
                st.write(f"Suggested Cognos table/query: {row['Cognos Table Suggestion']}")
            if row.get('Key Fields / BY / ACROSS'):
                st.write(f"Keys / grouping: {row['Key Fields / BY / ACROSS']}")
            if row.get('Formulas Created'):
                st.code(row['Formulas Created'], language='text')
            if row.get('Filters Applied'):
                st.code(row['Filters Applied'], language='text')


def render_pipeline(lineage_df, final_df):
    if lineage_df.empty:
        st.info("No visual pipeline could be generated.")
        return

    lines = []
    for _, row in lineage_df.sort_values('Step #').iterrows():
        input_value = str(row.get('Input Table(s)', '') or '').strip()
        output_value = str(row.get('Output Table', '') or row.get('Cognos Table Suggestion', '') or '').strip()
        action = str(row.get('What Happens In This Step', '') or '').strip()
        if input_value:
            lines.append(input_value)
        if action:
            lines.append(f"   -> {action}")
        if output_value:
            lines.append(output_value)
        lines.append("")

    if not final_df.empty:
        final_row = final_df.iloc[0]
        lines.append(str(final_row.get('Recommended Final Dataset', 'final dataset')))
        lines.append(f"   -> {final_row.get('Suggested Cognos Visual / Output', 'Cognos visual')}")

    st.code('\n'.join(lines).strip(), language='text')


def render_cognos_prep_checklist(build_plan_df):
    if build_plan_df.empty:
        st.info("No Framework Manager/Data Module checklist was generated.")
        return

    dm_rows = build_plan_df[
        build_plan_df['What To Build'].astype(str).str.contains('Framework Manager|Data Module|staging|merge|source query|custom column|unpivot', case=False, na=False)
    ]
    if dm_rows.empty:
        dm_rows = build_plan_df

    for _, row in dm_rows.sort_values('Plan Step').iterrows():
        label = f"Step {row['Plan Step']}: {row['Cognos Layer'] or row['WebFOCUS Object']}"
        with st.expander(label, expanded=int(row['Plan Step']) <= 4):
            st.checkbox(row['What To Build'], value=False, key=f"dm_{row['File Name']}_{row['File Path']}_{row['Plan Step']}")
            st.write(row['Why'])
            if row.get('Manual Review') == 'Yes':
                st.warning("Manual review recommended for this step.")


def render_mapping_cards(mapping_df):
    if mapping_df.empty:
        st.info("No DECODE mapping table rows were detected.")
        return

    for table_name, group in mapping_df.groupby('Mapping Table Name'):
        with st.expander(f"{table_name} ({len(group)} rows detected)", expanded=False):
            first = group.iloc[0]
            st.write(f"Purpose: map `{first['Source Field']}` values into `{first['Derived Field']}`.")
            st.write(first['Suggested Cognos Action'])
            cols = [
                'Mapping Table Name',
                'Source Field',
                'Source Value',
                'Mapped Value',
                'Derived Field',
                'Suggested Cognos Action',
            ]
            st.dataframe(group[[col for col in cols if col in group.columns]].head(25), use_container_width=True, hide_index=True)
            csv_data = group.to_csv(index=False).encode('utf-8')
            st.download_button(
                "Download mapping CSV",
                data=csv_data,
                file_name=f"{table_name}.csv",
                mime="text/csv",
                key=f"download_{table_name}_{first['File Name']}_{first['File Path']}",
            )


def render_visual_plan(final_df):
    if final_df.empty:
        st.info("No final visual recommendation was generated.")
        return

    for _, row in final_df.iterrows():
        with st.container(border=True):
            st.markdown(f"#### {row['Recommended Final Dataset']}")
            st.markdown(f"**Dataset shape:** {row['Dataset Shape']}")
            st.markdown(f"**Visual/output:** {row['Suggested Cognos Visual / Output']}") 
            st.markdown(f"**Reason:** {row['Reason']}")
            st.markdown(badge('Review before build', 'medium'), unsafe_allow_html=True)

        fields = [
            f"Values: {row['Measures / Value Columns'] or 'Confirm value fields'}",
            f"Axis/legend/category: {row['Legend / Category Fields'] or 'Confirm category fields'}",
            f"Source flow: {row['Source Tables / HOLD Flow'] or 'Review source flow'}",
        ]
        st.code('\n'.join(fields), language='text')

def collect_review_items(build_plan_df, joins_df, filters_df, formulas_df):
    items = []

    if not build_plan_df.empty:
        for _, row in build_plan_df[build_plan_df['Manual Review'].astype(str).eq('Yes')].iterrows():
            items.append(('Medium', row['Cognos Layer'] or row['WebFOCUS Object'], row['Why'], row['What To Build']))

    if not joins_df.empty:
        for _, row in joins_df.iterrows():
            priority = 'High' if str(row.get('WebFOCUS Join Mode', '')).upper() == 'JOIN TO ALL' else 'Medium'
            items.append((priority, 'Join key check', row.get('Review Reason', 'Confirm join behavior.'), row.get('Suggested Cognos Action', 'Review join.')))

    if not filters_df.empty:
        risky_filters = filters_df[filters_df['Parameter'].astype(str).str.len() > 0]
        for _, row in risky_filters.iterrows():
            items.append(('Medium', 'Parameter/filter check', row['Expression'], row['Cognos Guidance']))

    if not formulas_df.empty:
        risky_formulas = formulas_df[
            formulas_df['Formula Type'].astype(str).str.contains('Decode/Mapping|Date/Time|Parameter', case=False, na=False)
        ]
        for _, row in risky_formulas.iterrows():
            priority = 'High' if 'Decode/Mapping' in str(row['Formula Type']) else 'Medium'
            items.append((priority, f"Formula: {row['Field Name']}", row['Raw WebFOCUS Formula'], row['Suggested Cognos Expression / Action'] or 'Review formula rebuild.'))

    return items


def render_manual_review(build_plan_df, joins_df, filters_df, formulas_df):
    items = collect_review_items(build_plan_df, joins_df, filters_df, formulas_df)

    if not items:
        st.success("No high-priority manual review items were detected.")
        return

    priority_order = {'High': 0, 'Medium': 1, 'Low': 2}
    for idx, (priority, title, reason, action) in enumerate(sorted(items, key=lambda x: priority_order.get(x[0], 9)), start=1):
        level = priority.lower()
        with st.expander(f"{idx}. {priority} - {title}", expanded=idx <= 5):
            st.markdown(badge(priority, level if level in {'high', 'medium', 'low'} else ''), unsafe_allow_html=True)
            st.write(f"Why review: {reason}")
            st.write(f"Action: {action}")



def render_sql_view_plan(build_plan_df, report_row):
    sql_rows = pd.DataFrame()
    if build_plan_df is not None and not build_plan_df.empty:
        sql_rows = build_plan_df[
            build_plan_df['Cognos Layer'].astype(str).str.contains('SQL View', case=False, na=False)
        ]

    if sql_rows.empty and report_row.get('SQL View Need') == 'Not Needed':
        st.success("SQL View Need: Not Needed")
        st.write("This report appears suitable for Framework Manager/Data Module and the Cognos report layer.")
        return

    st.markdown(
        f"""
        <div class="wizard-card">
          <h3>SQL View Need: {report_row.get('SQL View Need', 'Needs Review')}</h3>
          <p><strong>Build Path:</strong> {report_row.get('Build Path', 'Needs Review')}</p>
          <p><strong>Action:</strong> {report_row.get('Cognos Action', 'Review SQL View need.')}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if sql_rows.empty:
        st.info("No explicit SQL View build step was generated. Review joins, HOLD tables, and performance needs.")
    else:
        for _, row in sql_rows.iterrows():
            with st.expander(f"Step {row['Plan Step']}: {row['Cognos Layer']}", expanded=True):
                st.write(row['What To Build'])
                st.write(row['Why'])
                st.warning("Validate row counts and join cardinality against WebFOCUS output.")



def render_calc_plan(formula_df, final_df):
    if formula_df is None or formula_df.empty:
        st.info("No calculation-specific formulas were detected. This may be mostly Framework Manager/Data Module and visual setup.")
        return

    calc_candidates = formula_df[
        formula_df['Formula Type'].astype(str).str.contains('Arithmetic|Conditional|Date/Time|Direct/Other', case=False, na=False)
    ]
    if calc_candidates.empty:
        st.info("No clear calculation candidates were detected. Use Framework Manager/Data Module for source shaping and mapping logic.")
        return

    st.write("Review these formulas and decide whether they belong in Framework Manager/Data Module custom query items or Cognos report calculations.")
    show_table_or_info(
        calc_candidates[[
            'Field Name',
            'Formula Type',
            'Raw WebFOCUS Formula',
            'Suggested Cognos Expression / Action',
            'Needs Manual Review',
        ]],
        "No calculation candidates found.",
    )


def render_fields_needed(field_df):
    if field_df is None or field_df.empty:
        st.info("No field list was detected.")
        return

    cols = [
        col for col in [
            'Source Table',
            'Fields Used',
            'Field Origin',
            'Field Role',
            'Is Field Shown In Report',
            'Formula Used',
        ]
        if col in field_df.columns
    ]
    show_table_or_info(field_df[cols], "No field list was detected.")


def render_filters_calcs(filter_df, formula_df):
    f1, f2 = st.tabs(["Filters", "Calculations"])
    with f1:
        show_table_or_info(filter_df, "No filters or parameters were detected.")
    with f2:
        show_table_or_info(formula_df, "No formulas were detected.")


def render_validation_checklist(report_row, join_df, final_df, drilldown_df=None):
    rows = build_validation_rows(report_row, join_df, final_df, drilldown_df)

    st.subheader("Migration Validation Checklist")
    st.caption("Use this before marking the report as rebuilt in Cognos. Work through each group, tick completed checks, and add tester notes if needed.")

    if not rows:
        st.info("No validation checks were generated for this report.")
        return

    report_key = re.sub(r'[^A-Za-z0-9_]+', '_', str(report_row.get('FEX Name', '')))
    checkbox_keys = [f"validation_{report_key}_{idx}" for idx in range(1, len(rows) + 1)]

    completed = sum(1 for key in checkbox_keys if st.session_state.get(key, False))
    total = len(rows)
    progress_value = completed / total if total else 0

    if completed == total:
        status_text = "Complete"
        status_class = "validation-complete"
    elif completed == 0:
        status_text = "Not started"
        status_class = "validation-not-started"
    else:
        status_text = "In progress"
        status_class = "validation-progress"

    st.markdown(
        f"""
        <div class="validation-summary-card">
            <div>
                <div class="validation-summary-title">{completed} / {total} checks completed</div>
                <div class="validation-summary-subtitle">Status: <span class="{status_class}">{status_text}</span></div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.progress(progress_value)

    grouped = defaultdict(list)
    for idx, row in enumerate(rows, start=1):
        grouped[row['Check Category']].append((idx, row))

    group_order = [
        'SQL / Data Validation',
        'Business Total Validation',
        'Report Logic Validation',
        'Interaction Validation',
    ]

    for category in group_order:
        items = grouped.get(category, [])
        if not items:
            continue
        with st.container(border=True):
            st.markdown(f"**{category}**")
            for idx, row in items:
                st.checkbox(
                    row['Validation Check'],
                    key=checkbox_keys[idx - 1],
                )

    st.text_area(
        "Validation notes",
        key=f"validation_notes_{report_key}",
        placeholder="Example: WebFOCUS row count = 12,450; Cognos row count = 12,450. Totals matched.",
        height=110,
    )

    reset_col_1, reset_col_2 = st.columns([0.78, 0.22])
    with reset_col_2:
        if st.button("Reset checklist", key=f"reset_validation_{report_key}", use_container_width=True):
            for key in checkbox_keys:
                if key in st.session_state:
                    del st.session_state[key]

            notes_key = f"validation_notes_{report_key}"
            if notes_key in st.session_state:
                del st.session_state[notes_key]

            st.rerun()


def render_sources_hold_combined(source_df, lineage_df, report_row):
    st.subheader("Sources")
    st.markdown('<div class="tab-intro">Confirm which files and lookup tables feed the report before rebuilding it in Cognos.</div>', unsafe_allow_html=True)
    render_source_files_tables(source_df, report_row)
    with st.expander("Show raw FILEDEF/source evidence"):
        cols = [col for col in ['Source Name', 'Detected Source Type', 'FILEDEF / Raw Source Reference', 'Notes'] if col in source_df.columns]
        show_table_or_info(
    source_df[cols] if source_df is not None and not source_df.empty else source_df,
    "No source evidence found.",
    large=True,
)
    st.subheader("HOLD / Temporary Datasets")
    render_hold_lineage_inspector(lineage_df, report_row)
    with st.expander("Show lineage evidence"):
        show_table_or_info(
            lineage_df,
            "No HOLD lineage steps found for this report.",
            large=True,
        )


def render_joins_filters_combined(join_df, filter_df, report_row):
    st.subheader("Joins / Matches")
    render_joins_matches_inspector(join_df, report_row)
    with st.expander("Show raw JOIN/MATCH statements"):
        cols = [col for col in ['Mapping Type', 'From Table', 'From Key', 'To Table', 'To Key', 'Raw WebFOCUS Statement', 'Review Reason'] if col in join_df.columns]
        show_table_or_info(
    join_df[cols] if join_df is not None and not join_df.empty else join_df,
    "No JOIN/MATCH statements found.",
    large=True,
)
    st.subheader("Filters / Parameters")
    render_filters_parameters_inspector(filter_df)
    with st.expander("Show raw filter expressions"):
        cols = [col for col in ['Filter Type', 'Expression', 'Parameter', 'Fields Used', 'Cognos Guidance'] if col in filter_df.columns]
        show_table_or_info(
    filter_df[cols] if filter_df is not None and not filter_df.empty else filter_df,
    "No filter expressions found.",
    large=True,
)


def render_joins_tab(join_df, report_row):
    st.subheader("Joins")
    st.markdown('<div class="tab-intro">Review JOIN and MATCH logic, keys, and whether the relationship belongs in SQL, Framework Manager, or the Cognos Data Module.</div>', unsafe_allow_html=True)
    render_joins_matches_inspector(join_df, report_row)
    with st.expander("Show raw JOIN/MATCH statements"):
        cols = [col for col in ['Mapping Type', 'From Table', 'From Key', 'To Table', 'To Key', 'Raw WebFOCUS Statement', 'Review Reason'] if col in join_df.columns]
        show_table_or_info(join_df[cols] if join_df is not None and not join_df.empty else join_df, "No JOIN/MATCH statements found.")


def render_filters_tab(filter_df):
    st.subheader("Filters")
    st.markdown('<div class="tab-intro">Translate WebFOCUS WHERE, IF, and parameter logic into Framework Manager/Data Module filters, report filters, or calculations.</div>', unsafe_allow_html=True)
    render_filters_parameters_inspector(filter_df)
    with st.expander("Show raw filter expressions"):
        cols = [col for col in ['Filter Type', 'Expression', 'Parameter', 'Fields Used', 'Cognos Guidance'] if col in filter_df.columns]


def render_cognos_build_plan(report_row, source_df, lineage_df, join_df, filter_df, formula_df, mapping_df, final_df, field_df, drilldown_df=None):
    st.info("The full evidence workbook is available from the Download Full Excel Report button.")

    ai_tab, sql_tab, dm_tab, model_calc_tab, visual_tab, validation_tab = st.tabs([
        "AI Build Plan",
        "SQL View",
        "Framework Manager / Data Module",
        "Data Model & Calculations",
        "Visuals",
        "Validation",
    ])

    with ai_tab:
        render_ai_build_plan(report_row, source_df, lineage_df, join_df, filter_df, formula_df, mapping_df, final_df, field_df)

    with sql_tab:
        render_sql_view_recommendation_inspector(report_row, source_df, lineage_df, join_df, filter_df, formula_df, field_df)

    with dm_tab:
        render_data_module_inspector(report_row)

    with model_calc_tab:
        render_data_model_and_calc_inspector(report_row, source_df, mapping_df, field_df, formula_df)

    with visual_tab:
        render_visuals_inspector(report_row, final_df, field_df)

    with validation_tab:
        render_validation_checklist(report_row, join_df, final_df, drilldown_df)


def render_output_drilldowns_combined(report_row, final_df, drilldown_df, lineage_df):
    render_drilldown_output_inspector(report_row, final_df, drilldown_df, lineage_df)


if submit:
    st.session_state["analysis_has_run"] = True
    st.session_state.pop("report_explorer_choice", None)
    try:
        if mode == "Multiple FEX Files":
            if not uploaded_fex_files:
                st.error("Please upload at least one .fex file.")
                st.stop()

            all_fex_items = [
                ("uploaded_files", f.name, read_uploaded_fex(f))
                for f in uploaded_fex_files
            ]
        else:
            if not uploaded_zip:
                st.error("Please upload a ZIP file.")
                st.stop()

            all_fex_items = collect_fex_from_zip(uploaded_zip)

            if not all_fex_items:
                st.error("No .fex files found inside the ZIP.")
                st.stop()

        ra_uploaded = resource_analyzer_file is not None

        if ra_uploaded:
            allowed_program_names, raw_ra_values = read_resource_analyzer_file(resource_analyzer_file)
        else:
            allowed_program_names, raw_ra_values = set(), []

        if allowed_program_names:
            selected_fex_items, matched_pairs = filter_fex_items_by_resource_analyzer(
                all_fex_items,
                allowed_program_names,
            )

            if not selected_fex_items:
                st.warning(
                    "Resource Analyzer names were detected, but none matched the uploaded FEX files. "
                    "KashMap will process all uploaded FEX files instead."
                )
                selected_fex_items = all_fex_items
        else:
            if ra_uploaded:
                st.warning(
                    "No program/report/FEX names could be detected from the Resource Analyzer file. "
                    "KashMap will process all uploaded FEX files."
                )
            else:
                st.info("No Resource Analyzer file uploaded. This step is optional, so KashMap will process all uploaded FEX files.")

            selected_fex_items = all_fex_items
            allowed_program_names = {
                normalize_program_name(fex_name)
                for _, fex_name, _ in all_fex_items
            }
            matched_pairs = [
                (normalize_program_name(fex_name), fex_name)
                for _, fex_name, _ in all_fex_items
            ]

        raw_fex_content = {
            (folder, fex_name): content
            for folder, fex_name, content in selected_fex_items
        }

        mas_files_dict, acx_files_dict = {}, {}
        if metadata_zip_file is not None:

            try:
                mas_files_dict, acx_files_dict = collect_metadata_from_zip(metadata_zip_file)
            except Exception as e:
                st.warning(f"Could not read Metadata ZIP: {e}")

        (
            output_stream,
            errors,
            dup_group_count,
            unique_count,
            unparsed_count,
            migration_count,
            total_tables,
            duplicate_counts_map,
            display_tables,
            table_db_mapped_count,
            table_db_needs_review_count,
        ) = build_output_workbook(
            selected_fex_items,
            allowed_program_names,
            matched_pairs,
            llm_options,
            mas_files_dict,
            acx_files_dict,
        )

        fn = output_name if output_name.lower().endswith(".xlsx") else f"{output_name}.xlsx"

        st.session_state["analysis_result"] = {
            "output_stream": output_stream,
            "errors": errors,
            "dup_group_count": dup_group_count,
            "unique_count": unique_count,
            "unparsed_count": unparsed_count,
            "migration_count": migration_count,
            "total_tables": total_tables,
            "duplicate_counts_map": duplicate_counts_map,
            "display_tables": display_tables,
            "selected_fex_count": len(selected_fex_items),
            "ra_uploaded": ra_uploaded,
            "fn": fn,
            "llm_options": llm_options,
            "table_db_mapped_count": table_db_mapped_count,
            "table_db_needs_review_count": table_db_needs_review_count,
            "raw_fex_content": raw_fex_content,
        }

    except Exception as e:
        st.error(str(e))

analysis_result = st.session_state.get("analysis_result")

if analysis_result:
    try:
        output_stream = analysis_result["output_stream"]
        errors = analysis_result["errors"]
        dup_group_count = analysis_result["dup_group_count"]
        unique_count = analysis_result["unique_count"]
        unparsed_count = analysis_result["unparsed_count"]
        migration_count = analysis_result["migration_count"]
        total_tables = analysis_result["total_tables"]
        duplicate_counts_map = analysis_result["duplicate_counts_map"]
        display_tables = analysis_result["display_tables"]
        selected_fex_count = analysis_result["selected_fex_count"]
        ra_uploaded = analysis_result["ra_uploaded"]
        fn = analysis_result["fn"]
        overview_df = display_tables['overview']

        show_download_button(output_stream, fn, "download_top")

        st.markdown('<div class="section-label">Choose a WebFOCUS report to inspect</div>', unsafe_allow_html=True)
        if overview_df.empty:
            st.info("No report rows were available to inspect.")
        else:
            total_reports = len(overview_df)
            search_term = ""
            sort_choice = "Name (A-Z)"

            if total_reports >= 20:
                search_col, sort_col = st.columns([0.6, 0.4])
                with search_col:
                    search_term = st.text_input(
                        "Search reports by name",
                        placeholder="Type part of a report name...",
                        key="report_search_term",
                    )
                with sort_col:
                    sort_choice = st.selectbox(
                        "Sort by",
                        [
                            "Name (A-Z)",
                            "Name (Z-A)",
                            "Complexity (High to Low)",
                            "Duplicate Type",
                            "Join/Match Count (High to Low)",
                            "Formula Count (High to Low)",
                        ],
                        key="report_sort_choice",
                    )

            filtered_df = overview_df
            if search_term.strip():
                filtered_df = filtered_df[
                    filtered_df['FEX Name'].astype(str).str.contains(search_term.strip(), case=False, na=False)
                ]

            if filtered_df.empty:
                st.warning(f"No reports matched '{search_term}'. Showing all {total_reports} reports instead.")
                filtered_df = overview_df

            filtered_df = sort_overview_df(filtered_df, sort_choice)

            report_options = [
                f"{row['FEX Name']} | {row['File Path']} | {row['Duplicate Type']}"
                for _, row in filtered_df.iterrows()
            ]

            if total_reports >= 20:
                st.caption(f"Showing {len(report_options)} of {total_reports} reports")

            selected_option = st.selectbox(
                "Choose a report to inspect",
                report_options,
                key="report_explorer_choice",
                label_visibility="collapsed",
            )
            selected_name, selected_path, _ = [part.strip() for part in selected_option.split('|', 2)]

            selected_overview = overview_df[
                (overview_df['FEX Name'] == selected_name)
                & (overview_df['File Path'] == selected_path)
            ]

            if not selected_overview.empty:
                report_row = selected_overview.iloc[0]
                source_df = filter_report_df(display_tables['sources'], selected_name, selected_path)
                lineage_df = filter_report_df(display_tables['lineage'], selected_name, selected_path)
                build_plan_df = filter_report_df(display_tables['build_plan'], selected_name, selected_path)
                mapping_df = filter_report_df(display_tables['mapping_tables'], selected_name, selected_path)
                final_df = filter_report_df(display_tables['final_dataset'], selected_name, selected_path)
                field_df = filter_report_df(display_tables['fields'], selected_name, selected_path)
                formula_df = filter_report_df(display_tables['formulas'], selected_name, selected_path)
                filter_df = filter_report_df(display_tables['filters'], selected_name, selected_path)
                join_df = filter_report_df(display_tables['joins'], selected_name, selected_path)
                duplicate_df = filter_report_df(display_tables['duplicates'], selected_name, selected_path)
                drilldown_df = filter_report_df(display_tables['drilldowns'], selected_name, selected_path)

                st.subheader(f"Report Inspector: {selected_name}")

                r1, r2, r3, r4, r5 = st.columns(5)

                with r1:
                    with st.container(border=True):
                        st.metric("Fields", int(report_row['Field Count']))
                with r2:
                    with st.container(border=True):
                        st.metric("Formulas", int(report_row['Formula Count']))
                with r3:
                    with st.container(border=True):
                        st.metric("Filters", int(report_row['Filter Count']))
                with r4:
                    with st.container(border=True):
                        st.metric("JOIN/MATCH", int(report_row['Join / Match Count']))
                with r5:
                    with st.container(border=True):
                        st.metric("Drilldowns", int(report_row.get('Drilldown Count', 0)))

                single_report_excel = build_single_report_excel(
                    report_row, field_df, formula_df, filter_df, join_df,
                    source_df, lineage_df, mapping_df, final_df, drilldown_df, build_plan_df,
                )
                report_download_key = re.sub(r'[^A-Za-z0-9_]+', '_', f"{selected_name}_{selected_path}")
                report_file_name = re.sub(r'[^A-Za-z0-9_.-]+', '_', Path(str(selected_name)).stem) + "_report.xlsx"
                st.download_button(
                    "Download This Report (Excel)",
                    data=single_report_excel,
                    file_name=report_file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_single_report_{report_download_key}",
                    use_container_width=True,
                )

                report_tabs = st.tabs([
                    "Start Here",
                    "Sources & HOLD",
                    "Joins, Filters & Parameters",
                    "Calculations / Mappings",
                    "Output & Drilldowns",
                    "Cognos Build Plan",
                    "Table → Database Mapping",
                ])

                with report_tabs[0]:
                    render_start_here(report_row, source_df, lineage_df, join_df, final_df)
                    render_duplicate_code_diff(
                        report_row,
                        display_tables['duplicates'],
                        overview_df,
                        analysis_result.get('raw_fex_content', {}),
                        selected_name,
                        selected_path,
                    )

                with report_tabs[1]:
                    render_sources_hold_combined(source_df, lineage_df, report_row)

                with report_tabs[2]:
                    render_joins_filters_combined(join_df, filter_df, report_row)

                with report_tabs[3]:
                    render_calculations_mappings_inspector(formula_df, mapping_df)

                with report_tabs[4]:
                    render_output_drilldowns_combined(report_row, final_df, drilldown_df, lineage_df)

                with report_tabs[5]:
                    render_cognos_build_plan(report_row, source_df, lineage_df, join_df, filter_df, formula_df, mapping_df, final_df, field_df, drilldown_df)

                with report_tabs[6]:
                    table_db_mapping_df = display_tables.get('table_db_mapping')
                    if table_db_mapping_df is not None and not table_db_mapping_df.empty:
                        st.caption("This table applies to all reports — it's the same regardless of which report is selected above.")

                        m1, m2, m3 = st.columns(3)
                        m1.metric("Total Unique Tables", len(table_db_mapping_df))
                        m2.metric("Mapped", analysis_result.get("table_db_mapped_count", 0))
                        m3.metric("Needs Review", analysis_result.get("table_db_needs_review_count", 0))

                        show_table_or_info(table_db_mapping_df, "No table mapping rows found.", large=True)

                        mapping_csv = table_db_mapping_df.to_csv(index=False).encode('utf-8')
                        st.download_button(
                            "Download Table DB Mapping CSV",
                            data=mapping_csv,
                            file_name="table_db_mapping.csv",
                            mime="text/csv",
                            key="download_table_db_mapping",
                        )
                    else:
                        st.info("No Metadata ZIP (.mas/.acx) was uploaded, so no table-to-database mapping is available. Upload a Metadata ZIP and re-analyze to see this data.")    


        if errors:
            st.warning(f"{len(errors)} file(s) had errors.")

            with st.expander("View Error Log"):
                for err in errors:
                    st.text(err)

    except Exception as e:
        st.error(str(e))