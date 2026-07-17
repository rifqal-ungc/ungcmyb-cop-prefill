import os, re, io
from urllib.parse import unquote
from flask import Flask, request, jsonify, Response
import openpyxl
from pypdf import PdfReader, PdfWriter

app = Flask(__name__)

BASE      = os.path.dirname(os.path.abspath(__file__))
XLSX_PATH = os.path.join(BASE, 'data', 'submissions.xlsx')
PDF_PATH  = os.path.join(BASE, 'api', 'template.pdf')

# ---------------------------------------------------------------------------
# Excel data loading (cached at first request)
# ---------------------------------------------------------------------------
_cache = None   # (sorted_company_names, {name: [rows]})

def _load():
    global _cache
    if _cache is not None:
        return _cache
    wb = openpyxl.load_workbook(XLSX_PATH, read_only=True, data_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    hdrs = [str(c or '') for c in next(rows)]
    col = {h: i for i, h in enumerate(hdrs)}

    def _get(row, key, default=''):
        idx = col.get(key)
        if idx is None or idx >= len(row):
            return default
        return str(row[idx] or '').strip()

    data, names = {}, set()
    for row in rows:
        if not row:
            continue
        name    = _get(row, 'NAME')
        country = _get(row, 'COUNTRY')
        if not name or (country and country not in ('Malaysia', 'Brunei')):
            continue
        names.add(name)
        data.setdefault(name, []).append({
            'section':     _get(row, 'SECTION'),
            'question_id': _get(row, 'QUESTION_ID'),
            'subquestion': _get(row, 'SUBQUESTION'),
            'choice':      _get(row, 'CHOICE'),
            'response':    _get(row, 'RESPONSE'),
        })
    wb.close()
    _cache = (sorted(names), data)
    return _cache

# ---------------------------------------------------------------------------
# Text normalisation helpers
# ---------------------------------------------------------------------------
def _norm(s):
    s = str(s or '').replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')
    s = re.sub(r'\s*/\s*', '/', s)
    s = re.sub(r'[\s:\-]+$', '', s)
    return re.sub(r'\s+', ' ', s.lower().strip())

def _match(a, b):
    a, b = _norm(a), _norm(b)
    return bool(a and b and (a == b or b.startswith(a) or a.startswith(b)))

# ---------------------------------------------------------------------------
# 2025 → 2026 question-ID shift
# ---------------------------------------------------------------------------
ID_REMAP = {
    'G12': 'G13', 'G13': 'G14',
    'E5': 'E7', 'E7': 'E8', 'E10': 'E11',
    # HR/L section: Excel may store without the "HR/" prefix
    'L2': 'HR/L2', 'L4.1': 'HR/L4.1', 'L5': 'HR/L5',
    'L6': 'HR/L6', 'L7': 'HR/L7',
}

# ---------------------------------------------------------------------------
# Choice-text wording fixes (2025 phrasing → 2026 phrasing)
# ---------------------------------------------------------------------------
_CREMAP = {
    'yes, focused on our own operations and the value chain (e.g., suppliers, consumers, communities, other business relationships)':
        'yes, focused on employees and the value chain (e.g., suppliers, consumers, communities, other business relationships)',
    'yes, focused on our own operations and the value chain':
        'yes, focused on employees and the value chain',
    'yes, with direct influence of some outcomes':
        'yes, with direct influence on some outcomes',
    'no, but we plan to within two years':
        'no, but we plan to within the next two years',
    'no, but we plan to in the next two years':
        'no, but we plan to within the next two years',
    'yes, adverse impact identified, and remedy provided/enabled':
        'yes, adverse impact(s) identified, and remedy provided/enabled',
    'yes, adverse impact identified, but no remedy provided/enabled':
        'yes, adverse impact(s) identified, but no remedy provided/enabled',
    'non-discrimination in respect of employment and occupation':
        'non-discrimination and equality (in respect of employment and occupation)',
    'to discuss potential ways to prevent or mitigate the risks/impacts in question':
        'to discuss potential ways to prevent/mitigate the risks/impacts in question',
    'choose to not disclose': 'choose not to disclose',
    'every year': 'annually',
    'every 2 years': 'every two years',
    # E6: 2025 Scope 3 measurement answer "Known" maps to 2026 "yes"
    'known': 'yes',
}

def _remap(text):
    k = _norm(text)
    return _CREMAP.get(k, k)

# Subquestion label remap: normalises 2025 row labels to match our row mappings.
# G4/G5/G6 use "X risks" while our gov_rows labels use "human rights", "anti-corruption" etc.
_SUBQ_REMAP = {
    'human rights risks':    'human rights',
    'labour rights risks':   'labour rights/decent work',
    'environmental risks':   'environment',
    'corruption risks':      'anti-corruption',
    # HR/L non-discrimination: 2025 omits "and equality" that 2026 added
    'non-discrimination in respect of employment and occupation':
        'non-discrimination and equality (in respect of employment and occupation)',
    # Env topics: 2025 used verbose "Waste (e.g., ...)" label; 2026 shortened to "Waste management"
    'waste (e.g., chemical spills, solid waste, hazardous, plastic, etc.)': 'waste management',
    # Shorter 2025 aliases for topics that kept the same meaning
    'waste':              'waste management',
    'energy & resource use': 'energy & resource use',   # explicit passthrough (no change needed)
}

def _remap_subq(subq):
    n = _norm(subq)
    return _SUBQ_REMAP.get(n, n)

def _best_option(choice, options):
    """Return index of best-matching option, using longest-prefix match to avoid ambiguity.
    'yes, focused on employees and the value chain' correctly beats 'yes, focused on employees'."""
    choice_n = _norm(choice)
    if not choice_n:
        return -1
    # Pass 1: exact match
    for i, opt in enumerate(options):
        if choice_n == _norm(opt):
            return i
    # Pass 2: bidirectional prefix match — take the longest overlapping match.
    # Handles both: choice is longer than option (choice has detail opt lacks)
    #           and: option is longer than choice (option has detail choice lacks).
    # "yes, focused on employees and value chain" must beat "yes, focused on employees".
    best_i, best_len = -1, -1
    for i, opt in enumerate(options):
        opt_n = _norm(opt)
        if not opt_n:
            continue
        if choice_n.startswith(opt_n):          # choice >= option in length
            match_len = len(opt_n)
        elif opt_n.startswith(choice_n):        # option >= choice in length
            match_len = len(choice_n)
        else:
            continue
        if match_len > best_len:
            best_i, best_len = i, match_len
    return best_i

# ---------------------------------------------------------------------------
# PDF form field mapping
# ---------------------------------------------------------------------------
# Radio button state values:  "Choice1", "Choice2", ... (1-indexed, left→right)
# Checkbox values:            "/Yes" (checked) | "/Off" (unchecked)
# Text field values:          plain string
#
# MATRIX_RADIO  – questions with rows × columns (radio per row)
#   { qid: { 'options': [...], 'rows': [(norm_label, field_name), ...],
#             'text_field': str|None } }
#
# SINGLE_RADIO  – one radio group for the whole question
#   { qid: { 'options': [...], 'field': str, 'text_field': str|None } }
#
# CHECKBOX_SEQ  – select-all checkboxes, sequential from start_n
#   { qid: { 'options': [...], 'prefix': str, 'start': int,
#             'text_field': str|None } }
#
# CHECKBOX_MAP  – checkboxes with explicit {norm_option: field_name} dict
#   { qid: { 'fields': {norm_opt: fname}, 'text_field': str|None } }
#
# TEXT_FIELDS   – free-text response only
#   { qid: str }
# ---------------------------------------------------------------------------

# Shared row labels for Governance matrix questions (G2-G7)
_GOV_ROWS = [
    ('human rights',              '{q} Radio Button 1'),
    ('labour rights/decent work', '{q} Radio Button 2'),
    ('environment',               '{q} Radio Button 3'),
    ('anti-corruption',           '{q} Radio Button 4'),
]

def _gov_rows(q):
    return [(lab, fld.replace('{q}', q)) for lab, fld in _GOV_ROWS]

# Shared 5-level progress scale (used by E3.1 and HR/L4.1)
_PROGRESS_OPTS = [
    'no action taken or planned',
    'early progress – commitments or initial actions taken',
    'some progress – partially implemented or piloted',
    'good progress – largely implemented across operations',
    'fully implemented across the company',
]

# Shared 9-row environmental topic list (used by E1, E3.1, E4)
_ENV_ROWS_9 = [
    ('climate change',               '{q} Radio Button 1'),
    ('water',                        '{q} Radio Button 2'),
    ('oceans',                       '{q} Radio Button 3'),
    ('nature and biodiversity',      '{q} Radio Button 4'),
    ('air pollution',                '{q} Radio Button 5'),
    ('waste management',             '{q} Radio Button 6'),
    ('circularity',                  '{q} Radio Button 7'),
    ('energy & resource use',        '{q} Radio Button 8'),
    ('other environmental topic(s)', '{q} Radio Button 9'),
]

def _env_rows(q):
    return [(lab, fld.replace('{q}', q)) for lab, fld in _ENV_ROWS_9]

# Shared 8-row HR/L topic list (used by HR/L2, HR/L4.1, HR/L5).
# Order matches the 2026 PDF template (confirmed from PDF screenshot):
# freedom of assoc → child → forced → non-discrimination → safe/healthy →
# wages → gender equality → [HR/L1.1 topics / other]
_HRL_ROWS_8 = [
    ('freedom of association and the right to collective bargaining', '{p} Radio Button 1'),
    ('child labour',                                                   '{p} Radio Button 2'),
    ('forced labour',                                                  '{p} Radio Button 3'),
    ('non-discrimination and equality (in respect of employment and occupation)', '{p} Radio Button 4'),
    ('safe and healthy working environment',                            '{p} Radio Button 5'),
    ('wages',                                                          '{p} Radio Button 6'),
    ("gender equality and women's rights",                             '{p} Radio Button 7'),
    ('other topic(s)',                                                  '{p} Radio Button 8'),
]

def _hrl_rows(prefix):
    return [(lab, fld.replace('{p}', prefix)) for lab, fld in _HRL_ROWS_8]

MATRIX_RADIO = {
    # ── Governance ──────────────────────────────────────────────────────────
    # Options listed in PDF left-to-right column order (kid 0 = leftmost)
    'G2': {
        'options': [
            'no, this is not a current priority',
            'no, but we plan to within the next two years',
            'yes, focused on employees',
            'yes, focused on employees and suppliers',
            'yes, focused on employees and the value chain (e.g., suppliers, consumers, communities, other business relationships)',
        ],
        'rows': _gov_rows('G2'),
        'text_field': 'G2 Text Field 4',
    },
    'G3': {
        'options': [
            'no one is specifically responsible for this topic',
            'yes, with limited influence on outcomes',
            'yes, with moderate influence on outcomes',
            'yes, with direct influence on some outcomes',
            'yes, with direct influence at the highest levels of the company',
        ],
        'rows': _gov_rows('G3'),
        'text_field': 'G3 Text Field 5',
    },
    # G3.1 reuses G3's PDF fields (confirmed: no G3.1 fields in radio debug).
    # Processing G3.1 last lets it overwrite G3 with the governance structure answer.
    'G3.1': {
        'options': [
            'no formal structure',
            'yes, with limited influence on outcomes',
            'yes, with moderate influence on outcomes',
            'yes, with direct influence on some outcomes',
            'yes, with direct influence at the highest level of the company',
        ],
        'rows': _gov_rows('G3'),   # same field names as G3
        'text_field': 'G3 Text Field 5',
    },
    'G4': {
        # PDF confirmed 5 kids (not 6) — removed "and external stakeholders" option
        'options': [
            'no, this is not a current priority',
            'no, but we plan to within the next two years',
            'yes, conducted by a designated individual or group',
            'yes, engaging employees across the company',
            'yes, engaging employees and business partners',
        ],
        'rows': _gov_rows('G4'),
        'text_field': 'G4 Text Field 6',
    },
    'G5': {
        # PDF confirmed 6 kids
        'options': [
            'no, this is not a current priority',
            'no, but we plan to within the next two years',
            'yes, related to our own operations',
            'yes, related to our own operations and suppliers',
            'yes, related to our own operations and the value chain',
            'choose not to disclose',
        ],
        'rows': _gov_rows('G5'),
        'text_field': 'G5 Text Field 7',
    },
    'G6': {
        # PDF confirmed 5 kids (was 4) — added extended formal process option
        'options': [
            'no, this is not a current priority',
            'no, but we plan to within the next two years',
            'yes, we have an informal process (e.g., through supervisors, others)',
            'yes, we have a formal process',
            'yes, we have a formal process accessible to employees and external stakeholders or the value chain',
        ],
        'rows': _gov_rows('G6'),
        'text_field': 'G6 Text Field 8',
    },
    # G5.1: conditional on G5 — which topics are included in due diligence (Yes/No per topic)
    # PDF column order: Yes (kid 0) | No (kid 1)
    'G5.1': {
        'options': ['yes', 'no'],
        'rows': _gov_rows('G5.1'),
        'text_field': 'G5.1 Text Field 7',
    },
    # G6.1: conditional on G6 — grievance mechanism attributes (Yes/No per attribute).
    # Row labels use the 2025 subquestion prefix so _match() hits on startswith.
    # /api/fields confirms only 4 radio buttons (1-4) — no Button 5.
    # PDF column order: Yes (kid 0) | No (kid 1)
    'G6.1': {
        'options': ['yes', 'no'],
        'rows': [
            ('is the process communicated to all employees',             'G6.1 Radio Button 1'),
            ('is the process available to non-employees',                'G6.1 Radio Button 2'),
            ('is the process confidential',                              'G6.1 Radio Button 3'),
            ('are there processes in place to avoid retaliation',        'G6.1 Radio Button 4'),
        ],
        'text_field': 'G6.1 Text Field 8',
    },
    # G7: tracking effectiveness per topic (4 rows × 4 options)
    'G7': {
        'options': [
            'we do not track the effectiveness of our actions on this topic',
            'we track this, but informally or through indirect measures',
            'we track this formally against qualitative goals or milestones',
            'we track this formally against quantitative targets',
        ],
        'rows': _gov_rows('G7'),
        'text_field': None,
    },
    # G7.1: conditional on G7 — public reporting on tracking (Yes/No per topic)
    # PDF column order: Yes (kid 0) | No (kid 1)
    'G7.1': {
        'options': ['yes', 'no'],
        'rows': [
            ('human rights',              'G7.1 Radio Button 1'),
            ('labour rights/decent work', 'G7.1 Radio Button 2'),
            ('environment',               'G7.1 Radio Button 3'),
            ('anti-corruption',           'G7.1 Radio Button 4'),
            ('gender equality',           'G7.1 Radio Button 5'),
            ('supply chain sustainability', 'G7.1 Radio Button 6'),
        ],
        'text_field': 'G7.1 Text Field 9',
    },
    # ── Environment ─────────────────────────────────────────────────────────
    'E1': {
        'options': [
            'no, and we have no plans to develop a policy',
            'no, but we plan to within the next two years',
            'yes, included within a broader policy or as a stand-alone policy',
            'not applicable (please provide additional information)',
        ],
        'rows': _env_rows('E1'),
        'text_field': 'E1 Text Field 1',
    },
    # E3.1: conditional on E3 — progress on environmental prevention per topic (9 rows × 5 options)
    'E3.1': {
        'options': _PROGRESS_OPTS,
        'rows': _env_rows('E3.1'),
        'text_field': None,
    },
    'E4': {
        'options': [
            'no adverse impact identified or caused',
            'yes, adverse impact(s) identified, but no remedy provided/enabled',
            'yes, adverse impact(s) identified, and remedy provided/enabled',
            'choose not to disclose (please provide additional information)',
        ],
        'rows': _env_rows('E4'),
        'text_field': 'E4 Text Field 1',
    },
    # ── Human Rights & Labour ────────────────────────────────────────────────
    'HR/L2': {
        'options': [
            # 2025: "...a policy"  2026: "...any policy/recommendation"
            # Use the shared prefix so both match via startswith.
            'no, and we have no plans to develop',
            'no, but we plan to within the next two years',
            'yes, included within a broader policy or as a stand-alone policy',
            'not applicable (please provide additional information)',
        ],
        'rows': _hrl_rows('L2'),
        'text_field': 'L2 Text Field 13',
    },
    # HR/L4.1: conditional on HR/L4 — progress on HR/L prevention per topic (8 rows × 5 options)
    'HR/L4.1': {
        'options': _PROGRESS_OPTS,
        'rows': _hrl_rows('L4.1'),
        'text_field': None,
    },
    'HR/L5': {
        'options': [
            'no adverse impact identified or caused',
            'yes, adverse impact(s) identified, but no remedy provided/enabled',
            'yes, adverse impact(s) identified, and remedy provided/enabled',
            'choose not to disclose (please provide additional information)',
        ],
        'rows': _hrl_rows('L5'),
        'text_field': None,
    },
    # ── Anti-Corruption ──────────────────────────────────────────────────────
    # AC4.1: conditional on AC4 — training frequency per group (3 rows × 4 options)
    'AC4.1': {
        'options': [
            'annually',
            'every two years',
            'less frequently than every two years',
            'varies by employee group or topic',
        ],
        'rows': [
            ('all employees',                                              'AC4.1 Radio Button 1'),
            ('selected employees (please provide additional information)', 'AC4.1 Radio Button 2'),
            ('third-party suppliers, contractors and/or consultants',      'AC4.1 Radio Button 3'),
        ],
        'text_field': None,
    },
}

SINGLE_RADIO = {
    # ── CoP Introduction ────────────────────────────────────────────────────
    'R1': {
        'options': [
            'complete the digital questionnaire with the option to also add a sustainability report (recommended)',
            'only upload a sustainability report',
        ],
        'field': 'Radio Button R1',
        'text_field': None,
    },
    # ── Governance ──────────────────────────────────────────────────────────
    'G12': {
        'options': [
            'no, this is not a current priority',
            'no, but we plan to within the next two years',
            'choose not to disclose',
            'yes, we consider sustainability in our financial planning and decision-making, but not through a structured approach',
            'yes, we take a structured approach to considering sustainability through a sustainability-informed investment or financing strategy, but this does not include specific targets tied to sustainability impact',
            'yes, we take a structured approach to considering sustainability in financing and investment through sdg-aligned investment or sdg-linked financing strategies, including specific targets tied to sustainability impact',
        ],
        'field': 'G12 V2 Radio Button ',
        'text_field': None,
    },
    'G14': {
        'options': [
            'yes (please provide additional information)',
            'no',
        ],
        'field': 'G14 Radio Button ',
        'text_field': 'G14 Text Field 12',
    },
    # ── Environment ─────────────────────────────────────────────────────────
    'E6': {
        # PDF confirmed 2 kids only — "partially measured" and "measured total" both map to index 1
        'options': [
            'we did not measure scope 3 emissions',
            'yes',
        ],
        'field': 'E6 Radio Button 1',
        'text_field': 'E6 Text Field 1',
    },
    'E14': {
        'options': [
            'no plan yet',
            'plan development is in progress (please provide additional information)',
            'yes, plan is developed but not yet implemented (please provide additional information)',
            'yes, it is implemented for selected priority locations/products/commodities only (please provide additional information)',
            'yes, it is implemented across the company',
        ],
        'field': 'E14 V2 Radio Button 1',
        'text_field': 'E14 Text Field 1',
    },
    # ── HR/L ────────────────────────────────────────────────────────────────
    'HR/L6': {
        'options': [
            'percentage of women (%)',
            'unknown',
        ],
        'field': 'L6 Radio Button 1',
        'text_field': 'L6 Text Field 1',
    },
    'HR/L7': {
        'options': [
            'rate of work-related accidents',
            'unknown',
        ],
        'field': 'L7 Radio Button 1',
        'text_field': 'L7 Text Field ',
    },
    # ── Anti-Corruption ──────────────────────────────────────────────────────
    'AC1': {
        'options': [
            'no, this is not a current priority',
            'no, but we plan to within the next two years',
            'yes',
        ],
        'field': 'AC1 Radio Button 1',
        'text_field': 'AC1 Text Field 15',
    },
    'AC2': {
        'options': [
            'no, and we have no plans to develop any policy/recommendation',
            'no, but we plan to within the next two years',
            'yes, included within a broader policy or as a stand-alone policy',
        ],
        'field': 'AC2 Radio Button 1',
        'text_field': 'AC2 Text Field 18',
    },
    'AC3': {
        'options': [
            'no, this is not a current priority',
            'no, but we plan to within the next two years',
            'yes (please provide additional information)',
        ],
        'field': 'AC3 Radio Button 1',
        'text_field': 'AC3 Text Field 19',
    },
}

CHECKBOX_SEQ = {
    # ── Success Stories ──────────────────────────────────────────────────────
    'S1': {
        'options': ['governance', 'human rights', 'labour', 'environment', 'anti-corruption'],
        'prefix': 'S1 Check Box', 'start': 1,
        'text_field': 'S1 Text Field 3',
    },
    'S2': {
        'options': ['governance', 'human rights', 'labour', 'environment', 'anti-corruption',
                    'none (please provide additional information)'],
        'prefix': 'S2 Check Box', 'start': 1,
        'text_field': 'S2 Text Field 4',
    },
    # ── Governance ──────────────────────────────────────────────────────────
    'G1': {
        'options': [
            'issue an annual statement about the relevance of sustainable development to the company',
            'issue an annual statement that addresses impacts on both people and the environment',
            'issue an annual statement highlighting a zero tolerance for corruption',
            'sign off on organizational sustainability targets',
            'supervise environmental, social, and governance reporting',
            'regularly review potential risks related to the business model',
            'none of the above',
        ],
        'prefix': 'G1 Check Box', 'start': 7,
        'text_field': 'G1 Text Field 5',
    },
    # ── Environment ─────────────────────────────────────────────────────────
    'E8': {
        'options': [
            'yes, and it includes physical risk assessments',
            'yes, and it includes transition risk assessments',
            'yes, and it includes a physical climate risk scenario analysis',
            'yes, and it includes actions to increase adaptation and resilience in the communities in which we operate',
            'no, but we plan to within the next two years',
            'no (please provide additional information)',
        ],
        'prefix': 'E8', 'start': 1,   # field names are "E8 v2 Text Field" + "E8 Check Box N"
        'text_field': 'E8 Text Field 1',
    },
    'E9': {
        'options': [
            'yes, we have set targets to phase out fossil fuel-based materials',
            'yes, we have set targets for investment in non-fossil fuel emitting activities',
            'yes, we have set targets for renewable energy procurement',
            'yes, we have set targets to end the exploration of fossil fuels, the expansion of existing fossil fuel reserves, the extraction of fossil fuels',
            'yes, we have set other targets to phase out fossil fuel usage',
            'no, but we plan to within the next two years',
            'no (please provide additional information)',
        ],
        'prefix': 'E9 Check Box', 'start': 1,
        'text_field': 'E9 Text Field 1',
    },
    # E11 removed from CHECKBOX_SEQ — it uses radio buttons, handled in BINARY_SELECT below
    'E16': {
        'options': [
            'has a formal circular economy policy or commitment',
            'has dedicated resources (budget and/or staff) to circular economy initiatives',
            'has integrated circularity considerations into product or service design',
            'promotes circular business models',
            'applies circular economy practices to the company\'s own operations',
            'engages and collaborates with suppliers and value chain partners to implement circular economy practices',
            'applies circular economy practices to waste management',
            'tracks and/or monitors circularity outcomes (please provide additional information)',
            'has other circular economy practices not listed above (please specify)',
            'none – the company has not implemented any circular economy practices in the reporting period',
        ],
        'prefix': 'E15 Check Box', 'start': 1,   # PDF field prefix for E16
        'text_field': 'E15 Text Field 14',
    },
    # ── Anti-Corruption ──────────────────────────────────────────────────────
    'AC4': {
        'options': [
            'all employees',
            'selected employees (please provide additional information)',
            'third-party suppliers, contractors and/or consultants',
            'no training provided',
        ],
        'prefix': 'AC4 Check Box', 'start': 1,
        'text_field': 'AC4 Text Field 20',
    },
    'AC5': {
        'options': [
            'internal reporting channels (e.g., confidential or anonymous speak-up mechanisms)',
            'whistleblower protection',
            'internal audits, compliance reviews, or other forms of controls',
            'third-party due diligence and ongoing monitoring',
            'external reporting channels (e.g., grievance or complaint mechanisms)',
            'external audits or independent reviews',
            'other (please provide additional information)',
            'no mechanisms in place to detect incidents of corruption',
        ],
        'prefix': 'AC5 v2 Check Box', 'start': 1,
        'text_field': 'AC5 Text Field 20',
    },
    'AC6': {
        'options': [
            'internal measures (e.g. internal investigation, review by board of directors, review by ethics committee)',
            'external measures (e.g., audit, review, report to and collaborate with authorities)',
            'other (please provide additional information)',
            'no actions were taken to address suspected incidents of corruption',
            'no incidents of corruption suspected',
        ],
        'prefix': 'AC6 Check Box', 'start': 1,
        'text_field': 'AC6 Text Field 20',
    },
    # AC1.2: monitoring mechanisms for anti-corruption programme (6 checkboxes)
    'AC1.2': {
        'options': [
            'yes, through review on an ad hoc basis',
            'yes, through internal employee self-evaluations',
            'yes, through automated controls monitoring',
            'yes, through external independent monitoring',
            'yes, through other mechanisms',
            'no, we do not monitor the anti-corruption compliance programme',
        ],
        'prefix': 'AC1.2 Check Box', 'start': 1,
        'text_field': 'AC1.2 Text Field 17',
    },
}

# Explicit checkbox maps where numbering is non-sequential or known
CHECKBOX_MAP = {
    'G13': {
        'fields': {
            'national/local regulation on sustainability':             'G13 Check Box 14',
            'security exchange regulations':                           'G13 Check Box 15',
            'corporate sustainability reporting directive (csrd) (formerly known as non-financial reporting directive of the european union (nfrd))': 'G13 Check Box 16',
            'voluntary sustainability reporting standards for non-listed smes (vsme)': 'G13 Check Box 17',
            'global reporting initiative (gri)':                       'G13 Check Box 18',
            'sustainability accounting standards board (sasb, now consolidated into the ifrs foundation)': 'G13 Check Box 19',
            'international integrated reporting council (iirc, now consolidated into the ifrs foundation)': 'G13 Check Box 20',
            'climate disclosure standards board (cdsb, now consolidated into the ifrs foundation)': 'G13 Check Box 21',
            'ifrs sustainability disclosure standards (ifrs s1 and s2) (incorporating the task force on climate-related financial disclosures (tcfd))': 'G13 Check Box 22',
            'taskforce on nature-related financial disclosures (tnfd)': 'G13 Check Box 23',
            'taskforce on inequality and social-related financial disclosures (tisfd)': 'G13 Check Box 24',
            'cdp (formerly known as carbon disclosure project)':        'G13 Check Box 25',
            'science based targets initiative (sbti)':                  'G13 Check Box 26',
            'other voluntary frameworks':                               'G13 Check Box 27',
            'no sustainability reporting according to any frameworks nor regulations outside of this communication on progress': 'G13 Check Box 28',
        },
        'text_field': 'G13 Text Field 12',
    },
    # HR/L2.2: single checkbox — whether policy references right to form/join a trade union
    'HR/L2.2': {
        'fields': {
            'reference the respect for the right': 'L2.2 Check Box 1',
        },
        'text_field': 'L2.2 Text Field 15',
    },
    'L2.2': {
        'fields': {
            'reference the respect for the right': 'L2.2 Check Box 1',
        },
        'text_field': 'L2.2 Text Field 15',
    },
    # HR/L2.3: single checkbox — no collective bargaining agreement in place
    'HR/L2.3': {
        'fields': {
            'there is no existing collective bargaining': 'L2.3 Check Box 1',
        },
        'text_field': 'L2.3 Text Field 15',
    },
    'L2.3': {
        'fields': {
            'there is no existing collective bargaining': 'L2.3 Check Box 1',
        },
        'text_field': 'L2.3 Text Field 15',
    },
    # E6.1: Scope 3 categories (17 sequential checkboxes in PDF order)
    'E6.1': {
        'fields': {
            'purchased goods and services':                    'E6.1 Check Box 1',
            'processing of sold products':                     'E6.1 Check Box 2',
            'capital goods':                                   'E6.1 Check Box 3',
            'use of sold products':                            'E6.1 Check Box 4',
            'fuel- and energy-related activities':             'E6.1 Check Box 5',
            'end-of-life treatment of sold products':          'E6.1 Check Box 6',
            'upstream transportation and distribution':        'E6.1 Check Box 7',
            'downstream leased assets':                        'E6.1 Check Box 8',
            'waste generated in operations':                   'E6.1 Check Box 9',
            'franchises':                                      'E6.1 Check Box 10',
            'business travel':                                 'E6.1 Check Box 11',
            'investments':                                     'E6.1 Check Box 12',
            'employee commuting':                              'E6.1 Check Box 13',
            'other - upstream':                                'E6.1 Check Box 14',
            'upstream leased assets':                          'E6.1 Check Box 15',
            'other - downstream':                              'E6.1 Check Box 16',
            'downstream transportation and distribution':      'E6.1 Check Box 17',
        },
        'text_field': 'E6.1 Text Field 1',
    },
    'AC1.1': {
        # /api/fields confirms AC1.1 Check Box 1-6 exist.
        # Note: "applied to own operations" is a prefix of "applied to own operations and suppliers"
        # so both the shorter and longer choices will set Box 3 (bidirectional prefix match).
        # This is acceptable — Box 3 checked as a superset doesn't harm the PDF read.
        'fields': {
            'publicly available':                                               'AC1.1 Check Box 1',
            'approved at most senior level of the company':                     'AC1.1 Check Box 2',
            "applied to the company's own operations":                          'AC1.1 Check Box 3',
            "applied to the company's own operations and suppliers":            'AC1.1 Check Box 4',
            "applied to the other stakeholders within the company's value chain": 'AC1.1 Check Box 5',
            'other (please provide additional information)':                    'AC1.1 Check Box 6',
        },
        'text_field': 'AC1.1 Text Field 16',
    },
}

# Binary Yes/No radio matrices where the CHOICE (or SUBQUESTION) identifies the row.
# For each matching row: radio_values[field] = 1 (Yes = kid index 1).
# Used for E11 where the 2025 data has one row per selected topic with choice = topic name.
CONCAT_TEXT_FIELDS = {
    # HR/L1 and HR/L1.1 are BOTH backed by the same PDF text field "HR/L1.1 Text Field 001".
    # The AcroForm for HR/L1 is a group with one /Tx child (T="1 Text Field 001"), giving the
    # full path "HR/L1.1 Text Field 001". The visual checkboxes are decorative/JS-rendered —
    # the actual form data goes here as text. Accumulate all selected topic choices.
    'HR/L1':   'HR/L1.1 Text Field 001',
    'HR/L1.1': 'HR/L1.1 Text Field 001',
}

BINARY_SELECT = {
    'E11': {
        'rows': [
            ('climate change',          'E11 Radio Button 1'),
            ('water',                   'E11 Radio Button 2'),
            ('nature and biodiversity', 'E11 Radio Button 3'),
            ('air pollution',           'E11 Radio Button 4'),
        ],
    },
}

# ---------------------------------------------------------------------------
# Checkbox matrix questions — select-all per row, row×col layout.
# box_number = start + row_idx * n_cols + col_idx
# ---------------------------------------------------------------------------
# 2026 env topic order (9 rows): climate, water, oceans, nature, air, waste, circularity, energy, other
_ENV_TOPICS_9 = [
    'climate change',
    'water',
    'oceans',
    'nature and biodiversity',
    'air pollution',
    'waste management',
    'circularity',
    'energy & resource use',
    'other environmental topic(s)',
]

# 7 fixed HR/L topics (same order in HR/L2.1, HR/L3, HR/L4)
_HRL_TOPICS_7 = [
    'freedom of association and the right to collective bargaining',
    'child labour',
    'forced labour',
    'non-discrimination and equality (in respect of employment and occupation)',
    'safe and healthy working environment',
    'wages',
    "gender equality and women's rights",
]

MATRIX_CHECKBOX = {
    # E1.1: policy attributes per env topic (9 topics × 8 options = 72 boxes)
    'E1.1': {
        'rows': _ENV_TOPICS_9,
        'options': [
            'aligned with international environmental standards',
            'publicly available',
            'approved at most senior level of the company',
            "applied to the company's own operations",
            "applied to the company's suppliers",
            "applied to the other stakeholders within the company's value chain",
            'developed involving environmental expertise from inside and/or outside the company',
            'other',
        ],
        'prefix': 'E1.1 Check Box', 'start': 1, 'n_cols': 8,
        'text_field': 'E1.1 Text Field 1',
    },
    # E2: stakeholder engagement per env topic (9 topics × 7 options = 63 boxes)
    'E2': {
        'rows': _ENV_TOPICS_9,
        'options': [
            'no engagement on this topic',
            'to better understand the risks/impacts in question',
            'to discuss potential ways to prevent/mitigate the risks/impacts in question',
            'to agree on a way to prevent/mitigate the risks/impacts in question',
            'to assess progress in preventing/mitigating the risks/impacts in question',
            'to collaborate in the prevention/mitigation of the risks/impacts in question',
            'other',
        ],
        'prefix': 'E2 Check Box', 'start': 1, 'n_cols': 7,
        'text_field': 'E2 Text Field 1',
    },
    # E3: prevention actions per env topic (9 topics × 7 options = 63 boxes)
    'E3': {
        'rows': _ENV_TOPICS_9,
        'options': [
            'no action within reporting period',
            'provided internal training/capacity building for the direct workforce',
            'built capacity among relevant business relationships',
            'conducted an audit process and/or corrective action plan',
            'collective action with peers or other stakeholders to address the issue',
            'collaborated with governmental or regulatory bodies',
            'other',
        ],
        'prefix': 'E3 Check Box', 'start': 1, 'n_cols': 7,
        'text_field': 'E3 Text Field 1',
    },
    # E3.1.2: progress tracking per env topic (9 topics × 5 options = 45 boxes)
    # Note: PDF field name has trailing dot: "E3.1.2. Check Box N"
    'E3.1.2': {
        'rows': _ENV_TOPICS_9,
        'options': [
            'progress is not tracked',
            'progress is reviewed against goals annually or more frequently',
            'progress is reported internally to the most senior level',
            'progress is reported externally',
            'other',
        ],
        'prefix': 'E3.1.2. Check Box', 'start': 1, 'n_cols': 5,
        'text_field': 'E3.1.2 Text Field 10',
    },
    # HR/L2.1: policy attributes per HR/L topic (8 rows × 8 options = 64 boxes)
    # Row 8 = first HR/L1.1-selected topic (handled by topic lookup below)
    'HR/L2.1': {
        'rows': _HRL_TOPICS_7,
        'options': [
            'aligned with international human rights/labour standards',
            'publicly available',
            'approved at most senior level of the company',
            "applied to the company's own operations",
            "applied to the company's suppliers",
            "applied to the other stakeholders within the company's value chain",
            'developed involving human rights/labour expertise from inside and/or outside the company',
            'other',
        ],
        'prefix': 'L2.1 Check Box', 'start': 1, 'n_cols': 8,
        'text_field': 'L2.1 Text Field 14',
    },
    # HR/L3: stakeholder engagement per HR/L topic (8 rows × 6 options = 48 boxes)
    'HR/L3': {
        'rows': _HRL_TOPICS_7,
        'options': [
            'no engagement on this topic',
            'to better understand the risks/impacts in question',
            'to discuss potential ways to prevent or mitigate the risks/impacts in question',
            'to agree on a way to prevent/mitigate the risks/impacts in question',
            'to assess progress in preventing/mitigating the risks/impacts in question',
            'to collaborate in the prevention/mitigation of the risks/impacts in question',
        ],
        'prefix': 'L3 Check Box', 'start': 1, 'n_cols': 6,
        'text_field': 'L3 Text Field ',
    },
    # HR/L4: prevention actions per HR/L topic (8 rows × 7 options = 56 boxes)
    'HR/L4': {
        'rows': _HRL_TOPICS_7,
        'options': [
            'no action within reporting period',
            'provided internal training/capacity building for the direct workforce',
            'built capacity among relevant business relationships',
            'conducted an audit process and/or corrective action plan',
            'collective action with peers or other stakeholders to address the issue',
            'collaborated with governmental or regulatory bodies',
            'other',
        ],
        'prefix': 'L4 Check Box', 'start': 1, 'n_cols': 7,
        'text_field': 'L4 Text Field',
    },
}

# ---------------------------------------------------------------------------
# Board composition (G9) — maps (qid, subq) → PDF text field
# RESP is the numeric value (write only when CHOICE == "Known")
# ---------------------------------------------------------------------------
BOARD_FIELDS = {
    'G9':  {'total number of board members (#)': 'G 10 Text Field 7'},
    'G9B': {
        'male (%)':   'G 10 Text Field 8',
        'female (%)': 'G 10 Text Field 9',
        'other (%)':  'G 10 Text Field 10',
    },
    'G9C': {
        'under 30 years old (%)':  'G 10 Text Field 11',
        '30-50 years old (%)':     'G 10 Text Field 12',
        'above 50 years old (%)':  'G 10 Text Field 13',
    },
    'G9D': {
        'from minority or vulnerable groups (%)': 'G 10 Text Field 14',
    },
    'G9E': {
        'executive (%)':   'G 10 Text Field 15',
        'independent (%)': 'G 10 Text Field 17',
    },
}

# ---------------------------------------------------------------------------
# GHG emissions (E5.1) — maps subq → (text_field, no_measure_radio)
# Write RESP to text_field when CHOICE == "Known"; set radio when "did not measure"
# ---------------------------------------------------------------------------
GHG_SCOPE_FIELDS = {
    'scope 1 emissions': {
        'text_field': 'E5 Text Field 11',
        'no_measure_radio': 'E5.1 Radio Button 1',
    },
    'scope 2 (market-based) emissions': {
        'text_field': 'E5 Text Field 12',
        'no_measure_radio': 'E5.1 Radio Button 2',
    },
    'scope 2 (location-based) emissions': {
        'text_field': 'E5 Text Field 17',
        'no_measure_radio': 'E5.1 Radio Button 3',
    },
}

TEXT_FIELDS = {
    'R2': 'Text Field R2',
    'R3': 'Text Field R3',
    'G8': 'G8 Text Field 9',
    'G10': 'G10 Text Field 10',
    'G11': 'G11 Text Field 11',
    'E5': 'E5 Text Field 7',
    'E8': 'E8 Text Field 1',
    'E13': 'E13 Text Field 5',
    'AC7': 'AC7 Text Field 20',
    # "A" suffix variants — same text field, different question_id in Excel
    'G2A':  'G2 Text Field 4',
    'G3A':  'G3 Text Field 5',
    'G3.1A': 'G3 Text Field 5',   # G3.1 reuses G3 fields; G3.1A overwrites G3A (G3.1 is more specific)
    'G7A':  'G7 Text Field 9',
    'G4A':  'G4 Text Field 6',
    'G5A':  'G5 Text Field 7',
    'G6A':  'G6 Text Field 8',
    'G8A':  'G8 Text Field 9',
    'G12A': 'G12 V2 Radio Button ',   # unlikely but safe
    'G13A': 'G13 Text Field 12',
    'G14A': 'G14 Text Field 12',
    'E1A':  'E1 Text Field 1',
    'E4A':  'E4 Text Field 1',
    'E5A':  'E5 Text Field 7',
    'E6A':  'E6 Text Field 1',
    'E8A':  'E8 Text Field 1',
    'E9A':  'E9 Text Field 1',
    'E13A': 'E13 Text Field 5',
    'E14A': 'E14 Text Field 1',
    'E16A': 'E15 Text Field 14',
    'L2A':  'L2 Text Field 13',
    'HR/L2A': 'L2 Text Field 13',
    'AC1A': 'AC1 Text Field 15',
    'AC2A': 'AC2 Text Field 18',
    'AC3A': 'AC3 Text Field 19',
    'AC4A': 'AC4 Text Field 20',
    'AC5A': 'AC5 Text Field 20',
    'AC6A': 'AC6 Text Field 20',
    'AC7A': 'AC7 Text Field 20',
    'S1A':  'S1 Text Field 3',
    'S2A':  'S2 Text Field 4',
    'G1A':  'G1 Text Field 5',
    # Confirmed from /api/fields:
    'G5.1A': 'G5.1 Text Field 7',
    'G6.1A': 'G6.1 Text Field 8',
    'G7.1A': 'G7.1 Text Field 9',
    # HR/L8: text description of living wage commitment
    'HR/L8': 'L8 Text Field ',
    'L8':    'L8 Text Field ',
    # AC1.1A: additional info for anti-corruption policy scope
    'AC1.1A': 'AC1.1 Text Field 16',
    # E2 and E3 "AA" text variants seen in some submissions
    'E2AA': 'E2 Text Field 1',
    'E3AA': 'E3 Text Field 1',
    # L6 second text field (denominator / total headcount)
    'HR/L6A': 'L6 Text Field 2',
    'L6A':    'L6 Text Field 2',
    # HR/L5 additional info text (note: period in field name confirmed by field-tree)
    'HR/L5A': 'L5. Text Field 9',
    'L5A':    'L5. Text Field 9',
    # HR/L3, HR/L4 additional info text fields (trailing space confirmed by field-tree)
    'HR/L3A': 'L3 Text Field ',
    'L3A':    'L3 Text Field ',
    'HR/L4A': 'L4 Text Field ',
    'L4A':    'L4 Text Field ',
    # HR/L4.1 additional info
    'HR/L4.1A': 'L4.1 Text Field ',
    'L4.1A':    'L4.1 Text Field ',
    # HR/L2.1 additional info
    'HR/L2.1A': 'L2.1 Text Field 14',
    'L2.1A':    'L2.1 Text Field 14',
    # E3.1 additional info text
    'E3.1A':  'E3.1 Text Field 1',
    # E3.1.1 target descriptions — concatenate all into first text field
    'E3.1.1': 'E3.1.1 Text Field 1',
    'E3.1.1A': 'E3.1.1 Text Field 1',
    # E5 (stays as E5 in 2026; ID_REMAP shifted old E5→E7, E7→E8)
    # Our existing TEXT_FIELDS['E5'] = 'E5 Text Field 7' is unreachable after ID_REMAP.
    # 'E7A': 'E5 Text Field 7' would handle the remapped path but E7 is now a checkbox Q.
    # E11 additional text fields (GHG measurements per topic)
    'E11A': 'E11 Text Field 1',
    # Additional IDs seen in some submissions
    'E1AAA':   'E1 Text Field 1',       # extra additional-info variant for E1
    'AC3AA':   'AC3 Text Field 19',     # second additional-info for AC3
    'G6.1AA':  'G6.1 Text Field 8',    # additional additional-info for G6.1
    'HR/L1A':  'HR/L1.1 Text Field 001',  # HR/L1 additional topic info (same text field)
    'E5.1A':   'E5 Text Field 7',       # E5.1 additional info → E5 general text field
    'E5.1AA':  'E5 Text Field 7',       # E5.1 extra additional info (same field)
}

# ---------------------------------------------------------------------------
# Radio button filler
# ---------------------------------------------------------------------------
# pypdf's update_page_form_field_values sets the parent field /V but does NOT
# update each child widget's /AS (appearance state), so radio buttons remain
# visually blank even though the value is stored.  This function walks the
# AcroForm field tree directly and sets both parent /V and all kids' /AS.
#
# radio_values: {field_name: choice_index_0based}
# ---------------------------------------------------------------------------
def _set_radio_fields(writer, radio_values):
    from pypdf.generic import NameObject
    if not radio_values:
        return
    try:
        acroform = writer._root_object['/AcroForm'].get_object()
    except Exception:
        return
    _walk_radio(acroform.get('/Fields', []), radio_values, '')


def _walk_radio(fields, radio_values, path):
    from pypdf.generic import NameObject
    for field_ref in fields:
        try:
            field = field_ref.get_object()
        except Exception:
            continue

        ft   = str(field.get('/FT', ''))
        t    = str(field.get('/T',  ''))
        ff   = int(field.get('/Ff', 0))
        # Build full qualified name (matches debug_radio endpoint logic)
        name = (path + '.' + t if path and t else (path or t)).strip('.')
        is_radio = (ft == '/Btn') and bool(ff & (1 << 15))

        if is_radio and name in radio_values:
            target_idx = radio_values[name]
            kids = field.get('/Kids', [])

            # Get the on-state from the TARGET kid specifically.
            # Each kid's on-state is its own index-based name (/0, /1, /2, …).
            target_on_state = None
            if 0 <= target_idx < len(kids):
                try:
                    kid_obj = kids[target_idx].get_object()
                    ap = kid_obj.get('/AP', {})
                    if hasattr(ap, 'get_object'):
                        ap = ap.get_object()
                    n_dict = ap.get('/N', {})
                    if hasattr(n_dict, 'get_object'):
                        n_dict = n_dict.get_object()
                    on_states = [k for k in n_dict.keys() if k != '/Off']
                    if on_states:
                        target_on_state = on_states[0]
                except Exception:
                    pass
            if target_on_state is None:
                target_on_state = f'/{target_idx}'  # fallback: /0, /1, /2 …
            if not target_on_state.startswith('/'):
                target_on_state = f'/{target_on_state}'

            # Set parent /V to the target kid's on-state name
            field[NameObject('/V')] = NameObject(target_on_state)

            # Set each kid /AS: the selected kid gets its own on-state, rest get /Off
            for i, kid_ref in enumerate(kids):
                try:
                    kid_obj = kid_ref.get_object()
                    if i == target_idx:
                        kid_obj[NameObject('/AS')] = NameObject(target_on_state)
                    else:
                        kid_obj[NameObject('/AS')] = NameObject('/Off')
                except Exception:
                    pass

        # Recurse into field groups (non-radio, non-leaf /Kids)
        elif '/Kids' in field and ft not in ('/Btn', '/Tx', '/Ch'):
            _walk_radio(field['/Kids'], radio_values, name)


# ---------------------------------------------------------------------------
# Core fill function
# ---------------------------------------------------------------------------
def _fill_pdf(subs):
    reader = PdfReader(PDF_PATH)
    writer = PdfWriter()
    writer.append(reader)

    # Deduplicate and remap IDs
    seen, clean = set(), []
    for s in subs:
        key = (s['question_id'], s['subquestion'], s['choice'], s['response'])
        if key in seen:
            continue
        seen.add(key)
        qid = ID_REMAP.get(s['question_id'], s['question_id'])
        clean.append({**s, 'question_id': qid})

    field_values = {}   # text + checkbox fields
    radio_values = {}   # radio field name → choice index (0-based)
    filled_qids  = set()

    for s in clean:
        qid      = s['question_id']
        choice   = _remap(s['choice'])
        subq     = _remap_subq(s['subquestion'])
        response = s['response'].strip()

        # ── MATRIX RADIO ──────────────────────────────────────────────────
        if qid in MATRIX_RADIO:
            q = MATRIX_RADIO[qid]
            field = None
            for label, fname in q['rows']:
                if _match(subq, label):
                    field = fname
                    break
            if field:
                best_i = _best_option(choice, q['options'])
                if best_i >= 0:
                    radio_values[field] = best_i
                    filled_qids.add(qid)
            if response and q.get('text_field'):
                field_values[q['text_field']] = response

        # ── SINGLE RADIO ──────────────────────────────────────────────────
        elif qid in SINGLE_RADIO:
            q = SINGLE_RADIO[qid]
            best_i = _best_option(choice, q['options'])
            if best_i >= 0:
                radio_values[q['field']] = best_i
                filled_qids.add(qid)
            if response and q.get('text_field'):
                field_values[q['text_field']] = response

        # ── SEQUENTIAL CHECKBOXES ─────────────────────────────────────────
        elif qid in CHECKBOX_SEQ:
            q = CHECKBOX_SEQ[qid]
            for i, opt in enumerate(q['options']):
                if _match(choice, opt):
                    n = q['start'] + i
                    field_values[f"{q['prefix']} {n}"] = '/Yes'
                    filled_qids.add(qid)
                    break
            if response and q.get('text_field'):
                field_values[q['text_field']] = response

        # ── EXPLICIT CHECKBOX MAP ─────────────────────────────────────────
        elif qid in CHECKBOX_MAP:
            q = CHECKBOX_MAP[qid]
            for opt_norm, fname in q['fields'].items():
                if _match(choice, opt_norm):
                    field_values[fname] = '/Yes'
                    filled_qids.add(qid)
                    break
            if response and q.get('text_field'):
                field_values[q['text_field']] = response

        # ── MATRIX CHECKBOX (select-all per row, row×col layout) ─────────────
        elif qid in MATRIX_CHECKBOX:
            q = MATRIX_CHECKBOX[qid]
            row_idx = next(
                (i for i, r in enumerate(q['rows']) if _match(subq, r)), -1
            )
            col_idx = _best_option(choice, q['options'])
            if row_idx >= 0 and col_idx >= 0:
                box_num = q['start'] + row_idx * q['n_cols'] + col_idx
                field_values[f"{q['prefix']} {box_num}"] = '/Yes'
                filled_qids.add(qid)
            if response and q.get('text_field'):
                field_values[q['text_field']] = response

        # ── BINARY SELECT (Yes/No radio where choice IS the row topic) ────────
        elif qid in BINARY_SELECT:
            q = BINARY_SELECT[qid]
            for label, fname in q['rows']:
                if _match(subq, label) or _match(choice, label):
                    radio_values[fname] = 1   # Yes = kid index 1
                    filled_qids.add(qid)

        # ── CONCAT TEXT (multi-row choices accumulated into one text field) ──
        elif qid in CONCAT_TEXT_FIELDS:
            fname = CONCAT_TEXT_FIELDS[qid]
            val = (response or choice).strip()
            if val:
                existing = field_values.get(fname, '')
                field_values[fname] = (existing + '; ' + val) if existing else val
                filled_qids.add(qid)

        # ── BOARD COMPOSITION (G9/G9B/G9C/G9D/G9E → G 10 Text Field N) ─────
        elif qid in BOARD_FIELDS:
            subq_fields = BOARD_FIELDS[qid]
            subq_norm = _norm(s['subquestion'])
            fname = subq_fields.get(subq_norm)
            if fname and _match(choice, 'known') and response:
                field_values[fname] = str(response)
                filled_qids.add(qid)

        # ── GHG SCOPE EMISSIONS (E5.1) ────────────────────────────────────
        elif qid == 'E5.1':
            scope_norm = _norm(s['subquestion'])
            sf = GHG_SCOPE_FIELDS.get(scope_norm)
            if sf:
                if _match(choice, 'we did not measure'):
                    radio_values[sf['no_measure_radio']] = 0
                    filled_qids.add(qid)
                elif _match(choice, 'known') and response:
                    field_values[sf['text_field']] = str(response)
                    filled_qids.add(qid)

        # ── TEXT FIELDS ───────────────────────────────────────────────────
        elif qid in TEXT_FIELDS:
            val = response or choice
            if val:
                field_values[TEXT_FIELDS[qid]] = val
                filled_qids.add(qid)

    # Fill text + checkboxes via standard pypdf API.
    # auto_regenerate=True rebuilds each field's appearance stream from its /V value,
    # ensuring text fields display their content even when the template has a stale /AP.
    for page in writer.pages:
        writer.update_page_form_field_values(page, field_values, auto_regenerate=True)

    # Fill radio buttons via direct AcroForm tree walk
    _set_radio_fields(writer, radio_values)

    buf = io.BytesIO()
    writer.write(buf)
    buf.seek(0)
    return buf, filled_qids

# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------
@app.route('/')
def index():
    html_path = os.path.join(BASE, 'index.html')
    with open(html_path, encoding='utf-8') as f:
        return f.read(), 200, {'Content-Type': 'text/html; charset=utf-8'}


@app.after_request
def cors(resp):
    resp.headers['Access-Control-Allow-Origin']  = '*'
    resp.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    resp.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    return resp


@app.route('/api/companies', methods=['GET', 'OPTIONS'])
def companies():
    if request.method == 'OPTIONS':
        return '', 200
    try:
        names, _ = _load()
        return jsonify(names)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/generate', methods=['POST', 'OPTIONS'])
def generate():
    if request.method == 'OPTIONS':
        return '', 200
    try:
        body = request.get_json() or {}
        company_name = str(body.get('company_name', '')).strip()
        if not company_name:
            return jsonify({'error': 'company_name is required'}), 400

        _, data = _load()
        subs = data.get(company_name)
        if not subs:
            return jsonify({'error': f'No submissions found for "{company_name}"'}), 404

        pdf_buf, filled_qids = _fill_pdf(subs)
        safe_name = re.sub(r'[^\w\s-]', '', company_name).strip().replace(' ', '_')

        return Response(
            pdf_buf.read(),
            mimetype='application/pdf',
            headers={
                'Content-Disposition': f'attachment; filename="CoP_2026_{safe_name}.pdf"',
                'X-Filled-Count': str(len(filled_qids)),
                'X-Filled-QIDs': ','.join(sorted(filled_qids)),
            }
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/test-fill', methods=['GET'])
def test_fill():
    """Debug: fill every field in the PDF — checkboxes checked, radios set to kid 0,
    text fields filled with 'Answer N' (sequential). Download the result to verify
    that every field type is writable."""
    try:
        from pypdf.generic import NameObject
        reader = PdfReader(PDF_PATH)
        writer = PdfWriter()
        writer.append(reader)

        field_values = {}   # text + checkbox
        radio_values = {}   # radio field name (full qualified) → kid index
        text_counter = [0]

        def walk(fields, path=''):
            for fref in fields:
                try:
                    f = fref.get_object()
                except Exception:
                    continue
                ft   = str(f.get('/FT', ''))
                t    = str(f.get('/T',  ''))
                ff   = int(f.get('/Ff', 0))
                name = (path + '.' + t if path and t else (path or t)).strip('.')
                is_radio    = (ft == '/Btn') and bool(ff & (1 << 15))
                is_checkbox = (ft == '/Btn') and not is_radio
                if is_radio and name:
                    radio_values[name] = 0   # select first kid
                elif is_checkbox and t:
                    field_values[t] = '/Yes'
                elif ft == '/Tx' and t:
                    text_counter[0] += 1
                    field_values[t] = f'Answer {text_counter[0]}'
                if '/Kids' in f:
                    walk(f['/Kids'], name)

        acroform = writer._root_object['/AcroForm'].get_object()
        walk(acroform.get('/Fields', []), '')

        for page in writer.pages:
            writer.update_page_form_field_values(page, field_values, auto_regenerate=True)
        _set_radio_fields(writer, radio_values)

        buf = io.BytesIO()
        writer.write(buf)
        buf.seek(0)
        return Response(
            buf.read(),
            mimetype='application/pdf',
            headers={
                'Content-Disposition': 'attachment; filename="test_fill_all_fields.pdf"',
                'X-Text-Fields': str(text_counter[0]),
                'X-Radio-Fields': str(len(radio_values)),
                'X-Checkbox-Fields': str(sum(1 for v in field_values.values() if v == '/Yes')),
            }
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/fields', methods=['GET'])
def dump_fields():
    """Debug: dump all PDF form field names and their current values."""
    try:
        reader = PdfReader(PDF_PATH)
        fields = reader.get_fields()
        out = {k: str(v.get('/V', '')) for k, v in (fields or {}).items()}
        return jsonify(out)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/debug/company/<path:company_name>', methods=['GET'])
def debug_company(company_name):
    """Debug: dump raw rows for a company to inspect question IDs and responses."""
    try:
        company_name = unquote(company_name)
        _, data = _load()
        # Exact match first; fall back to case-insensitive match
        rows = data.get(company_name)
        if rows is None:
            lc = company_name.lower()
            matched_key = next((k for k in data if k.lower() == lc), None)
            if matched_key:
                rows = data.get(matched_key, [])
                return jsonify({'count': len(rows), 'matched_key': matched_key, 'rows': rows})
            # Diagnostics: show received bytes to catch hidden characters
            return jsonify({
                'count': 0, 'rows': [],
                '_diag': {
                    'received': company_name,
                    'bytes': list(company_name.encode('utf-8')),
                    'sample_keys': list(data.keys())[:10],
                }
            })
        return jsonify({'count': len(rows), 'rows': rows})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/debug/company-search/<path:query>', methods=['GET'])
def debug_company_search(query):
    """Debug: search company names by substring (case-insensitive)."""
    try:
        query = unquote(query)
        names, data = _load()
        lq = query.lower()
        matches = [n for n in names if lq in n.lower()]
        return jsonify({'query': query, 'count': len(matches), 'matches': matches[:50]})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/debug/text', methods=['GET'])
def debug_pdf_text():
    """Debug: extract raw text from template PDF pages."""
    try:
        reader = PdfReader(PDF_PATH)
        pages = {}
        for i, page in enumerate(reader.pages):
            pages[f'page_{i+1}'] = page.extract_text() or ''
        return jsonify(pages)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/debug/radio', methods=['GET'])
def debug_radio():
    """Debug: dump radio button kid structure from PDF template."""
    from pypdf.generic import NameObject
    try:
        reader = PdfReader(PDF_PATH)
        writer = PdfWriter()
        writer.append(reader)
        acroform = writer._root_object['/AcroForm'].get_object()
        result = {}
        def walk(fields, path=''):
            for fref in fields:
                try:
                    f = fref.get_object()
                except Exception:
                    continue
                ft  = str(f.get('/FT', ''))
                t   = str(f.get('/T', ''))
                ff  = int(f.get('/Ff', 0))
                name = (path + '.' + t if path else t).strip('.')
                is_radio = (ft == '/Btn') and bool(ff & (1 << 15))
                if is_radio:
                    kids = f.get('/Kids', [])
                    kid_states = []
                    for kid_ref in kids:
                        try:
                            kid = kid_ref.get_object()
                            ap = kid.get('/AP', {})
                            if hasattr(ap, 'get_object'): ap = ap.get_object()
                            n = ap.get('/N', {})
                            if hasattr(n, 'get_object'): n = n.get_object()
                            states = list(n.keys())
                        except Exception:
                            states = []
                        kid_states.append(states)
                    result[name] = kid_states
                elif '/Kids' in f and ft not in ('/Btn', '/Tx', '/Ch'):
                    walk(f['/Kids'], name)
        walk(acroform.get('/Fields', []))
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/debug/field-kids/<path:field_name>', methods=['GET'])
def debug_field_kids(field_name):
    """Debug: dump the child widget annotations of a named AcroForm group field.
    Use to find kid ordering for orphan-annotation checkboxes like HR/L1."""
    field_name = unquote(field_name)
    try:
        reader = PdfReader(PDF_PATH)
        acroform = reader.trailer['/Root']['/AcroForm'].get_object()

        def find_field(fields, target):
            for fref in fields:
                try:
                    f = fref.get_object()
                except Exception:
                    continue
                t = str(f.get('/T', ''))
                if t == target:
                    return f
                if '/Kids' in f:
                    found = find_field(f['/Kids'], target)
                    if found is not None:
                        return found
            return None

        field = find_field(acroform.get('/Fields', []), field_name)
        if field is None:
            return jsonify({'error': f'field "{field_name}" not found'}), 404

        kids = field.get('/Kids', [])
        result = []
        for i, kid_ref in enumerate(kids):
            try:
                kid = kid_ref.get_object()
                ft  = str(kid.get('/FT', ''))
                t   = str(kid.get('/T', ''))
                v   = str(kid.get('/V', ''))
                as_ = str(kid.get('/AS', ''))
                tu  = str(kid.get('/TU', ''))    # tooltip / label
                contents = str(kid.get('/Contents', ''))
                ap = kid.get('/AP', {})
                if hasattr(ap, 'get_object'): ap = ap.get_object()
                n_dict = ap.get('/N', {})
                if hasattr(n_dict, 'get_object'): n_dict = n_dict.get_object()
                on_states = [k for k in (n_dict.keys() if hasattr(n_dict, 'keys') else []) if k != '/Off']
                rect = [str(x) for x in kid.get('/Rect', [])]
                result.append({
                    'index': i, 'T': t, 'FT': ft, 'V': v, 'AS': as_,
                    'TU': tu, 'Contents': contents,
                    'on_states': on_states, 'Rect': rect,
                })
            except Exception as e:
                result.append({'index': i, 'error': str(e)})

        return jsonify({'field': field_name, 'kid_count': len(kids), 'kids': result})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/debug/field-tree', methods=['GET'])
def debug_field_tree():
    """Debug: dump EVERY field in the AcroForm tree (including nested checkbox children).
    Use this to find the actual field names for HR/L1 checkboxes and similar nested groups."""
    try:
        reader = PdfReader(PDF_PATH)
        acroform = reader.trailer['/Root']['/AcroForm'].get_object()
        result = {}

        def walk(fields, path=''):
            for fref in fields:
                try:
                    f = fref.get_object()
                except Exception:
                    continue
                ft   = str(f.get('/FT', ''))
                t    = str(f.get('/T',  ''))
                ff   = int(f.get('/Ff', 0))
                name = (path + '.' + t if path else t).strip('.')
                is_radio    = (ft == '/Btn') and bool(ff & (1 << 15))
                is_checkbox = (ft == '/Btn') and not is_radio
                if is_radio:
                    kids = f.get('/Kids', [])
                    result[name] = {'type': 'radio', 'kids': len(kids)}
                elif is_checkbox:
                    result[name] = {'type': 'checkbox', 'v': str(f.get('/V', ''))}
                elif ft == '/Tx':
                    result[name] = {'type': 'text', 'v': str(f.get('/V', ''))}
                elif '/Kids' in f:
                    result[name] = {'type': 'group'}
                    walk(f['/Kids'], name)
                    continue
                if '/Kids' in f and is_radio:
                    for i, kid_ref in enumerate(f['/Kids']):
                        try:
                            kid = kid_ref.get_object()
                            kid_t = str(kid.get('/T', ''))
                            kid_name = f'{name}.kid{i}' if not kid_t else f'{name}.{kid_t}'
                            kid_ft = str(kid.get('/FT', ''))
                            kid_ff = int(kid.get('/Ff', 0))
                            kid_is_checkbox = (kid_ft == '/Btn') and not ((kid_ft == '/Btn') and bool(kid_ff & (1 << 15)))
                            if kid_is_checkbox:
                                result[kid_name] = {'type': 'checkbox', 'v': str(kid.get('/V', ''))}
                        except Exception:
                            pass

        walk(acroform.get('/Fields', []))
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, port=5001)
