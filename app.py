import os, json, re, io, base64, urllib.request, urllib.parse
from flask import Flask, request, jsonify, Response
from openpyxl import load_workbook

app = Flask(__name__)

SUPABASE_URL = os.environ.get('SUPABASE_URL', 'https://qlvlpvgyjoeprvghmoyn.supabase.co').lstrip('﻿').strip()
SUPABASE_KEY = os.environ.get('SUPABASE_SERVICE_ROLE_KEY', '').lstrip('﻿').strip()
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'api', 'template.xlsx')

# Supabase section name → Excel sheet name
SHEET_SECTION_MAP = {
    'Governance':               ' Governance',
    'Human Rights and Labour':  'Human Rights & Labour',
    'Environment':              'Environment',
    'Anti-Corruption':          ' Anti-Corruption',
    'CEO Statement':            'CEO Statement',
    'Sustainability Report':    None,  # no fillable sheet in template
}

# CEO Statement section has S1/S2 (Success Stories) and C1/C2 (statement form).
# S-prefixed questions go to the Success Stories sheet; C-prefixed to CEO Statement.
SUCCESS_STORIES_SHEET = 'Success Stories & Future Priori'
SUCCESS_STORIES_Q_PREFIX = re.compile(r'^S\d+')

Q_PATTERN = re.compile(
    r'^(?:\(Optional\)\s*)?'
    r'(G\d+(?:\.\d+)?|HR/L\d+(?:\.\d+)?|E\d+(?:\.\d+)?|AC\d+(?:\.\d+)?'
    r'|C\d+|S\d+|R\d+)[.\s]'
)
SUFFIX_PATTERN = re.compile(r'^(.*?\d+)([A-Z]+)$')

# Maps 2025 question IDs → 2026 question IDs where numbering shifted between years.
QUESTION_ID_MAP_2025_TO_2026 = {
    # Governance: new G12 (financing/investment) inserted, pushing G12/G13 up by 1
    'G12': 'G13',   # "Do you produce sustainability reporting according to" → 2026 G13
    'G13': 'G14',   # "Is information assured by a third-party" → 2026 G14
    # Environment: Scope 1/2 measurement added as E5; climate/nature section reorganised
    'E5':  'E7',    # "Does company have GHG target validated by third-party" → 2026 E7
    'E7':  'E8',    # "Does company have a climate adaptation plan" → 2026 E8
    'E10': 'E11',   # "Material environmental topics" → 2026 E11
}

# Maps 2025 choice/subquestion texts → 2026 equivalents where option wording changed.
CHOICE_TEXT_MAP_2025_TO_2026 = {
    # G2/G3/G4/G5/G6/G7/G8 — broadest scope option renamed between 2025 and 2026
    'yes, focused on our own operations and the value chain (e.g., suppliers, consumers, communities, other business relationships)':
        'yes, focused on employees and the value chain (e.g., suppliers, consumers, communities, other business relationships)',
    'yes, focused on our own operations and the value chain':
        'yes, focused on employees and the value chain',
}

# Subquestion label normalization — slashes with/without spaces, abbreviations
SUBQ_TEXT_MAP_2025_TO_2026 = {
    'labour rights/decent work': 'labour rights / decent work',
    'human rights and labour':   'human rights & labour',
}

CHECKBOX = '❑'
CHECKED  = '☑'
RADIO    = '🔾'
SELECTED = '🔘'


def supabase_get(path):
    req = urllib.request.Request(
        f"{SUPABASE_URL}/rest/v1/{path}",
        headers={'apikey': SUPABASE_KEY, 'Authorization': f'Bearer {SUPABASE_KEY}'}
    )
    with urllib.request.urlopen(req, timeout=20) as resp:
        return json.loads(resp.read())


def normalize(s):
    return re.sub(r'\s+', ' ', str(s or '').lower().strip()
                  .replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' '))


def remap_choice(text):
    """Apply 2025→2026 choice text remapping."""
    key = normalize(text)
    return CHOICE_TEXT_MAP_2025_TO_2026.get(key, key)


def remap_subq(text):
    """Apply 2025→2026 subquestion label remapping."""
    key = normalize(text)
    return SUBQ_TEXT_MAP_2025_TO_2026.get(key, key)


def texts_match(a, b):
    """Match two option/label texts after normalization."""
    a, b = normalize(a), normalize(b)
    if not a or not b:
        return False
    # Exact match or one is a prefix of the other (handles truncated Supabase strings)
    return a == b or b.startswith(a) or a.startswith(b)


def parse_question_locs(ws):
    locs = []
    for i, row in enumerate(ws.iter_rows()):
        val = row[1].value if len(row) > 1 else None
        if val:
            m = Q_PATTERN.match(str(val).strip())
            if m:
                locs.append((m.group(1), i))
    return locs


def fill_sheet(ws, subs):
    all_rows = list(ws.iter_rows())
    n = len(all_rows)
    q_locs = parse_question_locs(ws)
    q_dict = {qid: idx for qid, idx in q_locs}

    filled_qids = []

    for sub in subs:
        qid      = sub['question_id']
        choice   = str(sub.get('choice', '')      or '').strip()
        subq     = str(sub.get('subquestion', '') or '').strip()
        response = str(sub.get('response', '')    or '').strip()

        # Apply choice/subquestion text remapping
        choice_norm = remap_choice(choice)
        subq_norm   = remap_subq(subq)

        q_row = q_dict.get(qid)
        parent_qid = qid
        if q_row is None:
            sm = SUFFIX_PATTERN.match(qid)
            if sm:
                parent_qid = sm.group(1)
                q_row = q_dict.get(parent_qid)

        if q_row is None:
            continue

        end_row = n
        for _, other_row in q_locs:
            if other_row > q_row and other_row < end_row:
                end_row = other_row

        success = False

        if response and not choice:
            # Text response → write in "Please provide additional information" row
            for i in range(q_row, end_row):
                cell = all_rows[i][1] if len(all_rows[i]) > 1 else None
                if cell and cell.value and 'Please provide additional information' in str(cell.value):
                    cell.value = 'Response: ' + response
                    success = True
                    break

        elif choice and not subq:
            # Standalone checkbox (❑) or radio (🔾) in column B
            for i in range(q_row, end_row):
                cell = all_rows[i][1] if len(all_rows[i]) > 1 else None
                if not cell or not cell.value:
                    continue
                val = str(cell.value)
                if val.startswith(CHECKBOX):
                    option_text = val[len(CHECKBOX):].strip()
                    if texts_match(choice_norm, normalize(option_text)):
                        cell.value = CHECKED + ' ' + option_text
                        success = True
                elif val.startswith(RADIO):
                    option_text = val[len(RADIO):].strip()
                    if texts_match(choice_norm, normalize(option_text)):
                        cell.value = SELECTED + ' ' + option_text
                        success = True
                        break  # Radio — only one selection

        elif choice and subq:
            # Matrix question — find header row, then data row by subquestion
            header_row = None
            header_cols = {}
            for i in range(q_row + 1, end_row):
                row = all_rows[i]
                b_val = row[1].value if len(row) > 1 else None
                c_val = row[2].value if len(row) > 2 else None
                if (not b_val or str(b_val).strip() == '') and c_val:
                    header_row = i
                    for j, cell in enumerate(row):
                        if j >= 2 and cell.value:
                            header_cols[j] = normalize(str(cell.value))
                    break

            if header_row is not None:
                # Find the data row matching the subquestion label
                target_data_row = None
                for i in range(header_row + 1, end_row):
                    label = all_rows[i][1].value if len(all_rows[i]) > 1 else None
                    if label and texts_match(subq_norm, normalize(str(label))):
                        target_data_row = i
                        break

                if target_data_row is not None:
                    data_row = all_rows[target_data_row]

                    # Match choice to a column header (use remapped choice_norm)
                    target_col = None
                    for col_idx, header_text in header_cols.items():
                        if texts_match(choice_norm, header_text):
                            target_col = col_idx
                            break

                    if target_col is not None:
                        if len(data_row) > target_col:
                            cell = data_row[target_col]
                            cell_val = str(cell.value or '').strip()
                            if cell_val == RADIO:
                                cell.value = SELECTED
                                success = True
                            elif cell_val == CHECKBOX:
                                cell.value = CHECKED
                                success = True

                    elif choice.lower() == 'known' and response:
                        for j in range(2, len(data_row)):
                            cell = data_row[j]
                            if '_' in str(cell.value or ''):
                                cell.value = response
                                success = True
                                break

                    elif choice.lower() in ('unknown', 'not applicable'):
                        for col_idx, header_text in header_cols.items():
                            if normalize(header_text) in ('unknown', 'not applicable'):
                                if len(data_row) > col_idx:
                                    cell = data_row[col_idx]
                                    cell_val = str(cell.value or '').strip()
                                    if cell_val == RADIO:
                                        cell.value = SELECTED
                                        success = True
                                    elif cell_val == CHECKBOX:
                                        cell.value = CHECKED
                                        success = True
                                break

        if success:
            filled_qids.append(parent_qid if parent_qid != qid else qid)

    return filled_qids


def get_all_excel_qids(wb):
    results = []
    sheets_to_scan = list(SHEET_SECTION_MAP.items()) + [('Success Stories', SUCCESS_STORIES_SHEET)]
    seen_sheets = set()
    for section, sheet_name in sheets_to_scan:
        if not sheet_name or sheet_name in seen_sheets:
            continue
        seen_sheets.add(sheet_name)
        try:
            ws = wb[sheet_name]
        except KeyError:
            continue
        for qid, row_idx in parse_question_locs(ws):
            all_rows = list(ws.iter_rows())
            text = str(all_rows[row_idx][1].value or '').strip() if row_idx < len(all_rows) else ''
            results.append({'qid': qid, 'section': section, 'text': text})
    return results


@app.route('/')
def index():
    html_path = os.path.join(os.path.dirname(__file__), 'index.html')
    with open(html_path, encoding='utf-8') as f:
        return f.read(), 200, {'Content-Type': 'text/html; charset=utf-8'}


@app.after_request
def cors(response):
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    return response


@app.route('/api/companies', methods=['GET', 'OPTIONS'])
def companies():
    if request.method == 'OPTIONS':
        return '', 200
    try:
        rows = supabase_get('companies?network=eq.MY&select=company_name&order=company_name')
        names = [r['company_name'] for r in rows if r.get('company_name')]
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

        encoded = urllib.parse.quote(company_name)
        subs = supabase_get(
            f'company_submissions?network=eq.MY'
            f'&company_name=eq.{encoded}'
            f'&select=question_id,section,subquestion,choice,response'
            f'&order=section,question_id'
        )

        # Deduplicate
        seen, unique_subs = set(), []
        for s in subs:
            key = (s['question_id'], s.get('subquestion', ''),
                   s.get('choice', ''), s.get('response', ''))
            if key not in seen:
                seen.add(key)
                unique_subs.append(s)

        # Remap 2025 question IDs that shifted in 2026 questionnaire
        for s in unique_subs:
            mapped = QUESTION_ID_MAP_2025_TO_2026.get(s['question_id'])
            if mapped:
                s['question_id'] = mapped

        # Group by section, but split CEO Statement → Success Stories (S*) vs CEO sheet (C*)
        by_section = {}
        for s in unique_subs:
            section = s.get('section', '')
            qid = s['question_id']
            if section == 'CEO Statement' and SUCCESS_STORIES_Q_PREFIX.match(qid):
                section = '_SuccessStories'
            by_section.setdefault(section, []).append(s)

        # Fill template
        wb = load_workbook(TEMPLATE_PATH)
        all_filled = []

        for section, section_subs in by_section.items():
            if section == '_SuccessStories':
                sheet_name = SUCCESS_STORIES_SHEET
            else:
                sheet_name = SHEET_SECTION_MAP.get(section)
            if not sheet_name:
                continue
            try:
                ws = wb[sheet_name]
            except KeyError:
                continue
            filled = fill_sheet(ws, section_subs)
            all_filled.extend(filled)

        all_excel_qids = get_all_excel_qids(wb)
        filled_set = set(all_filled)
        pending = [q for q in all_excel_qids if q['qid'] not in filled_set]

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        excel_b64 = base64.b64encode(buf.read()).decode()

        safe_name = re.sub(r'[^\w\s-]', '', company_name).strip().replace(' ', '_')

        return jsonify({
            'excel_base64': excel_b64,
            'filename': f'CoP_2026_{safe_name}.xlsx',
            'total_filled': len(filled_set),
            'total_questions': len(all_excel_qids),
            'pending': pending,
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, port=5001)
