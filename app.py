import os, json, re, io, base64, urllib.request, urllib.parse
from flask import Flask, request, jsonify, Response
from openpyxl import load_workbook

app = Flask(__name__)

SUPABASE_URL = os.environ.get('SUPABASE_URL', 'https://qlvlpvgyjoeprvghmoyn.supabase.co')
SUPABASE_KEY = os.environ.get('SUPABASE_SERVICE_ROLE_KEY', '')
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'api', 'template.xlsx')

SHEET_SECTION_MAP = {
    'Governance':               ' Governance',
    'Human Rights and Labour':  'Human Rights & Labour',
    'Environment':              'Environment',
    'Anti-Corruption':          ' Anti-Corruption',
    'CEO Statement':            'CEO Statement',
}

Q_PATTERN = re.compile(
    r'^(G\d+(?:\.\d+)?|HR/L\d+(?:\.\d+)?|E\d+(?:\.\d+)?|AC\d+(?:\.\d+)?'
    r'|C\d+|S\d+|R\d+)[.\s]'
)
SUFFIX_PATTERN = re.compile(r'^(.*?\d+)([A-Z]+)$')

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


def texts_match(a, b):
    a, b = normalize(a), normalize(b)
    if not a or not b:
        return False
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
        qid     = sub['question_id']
        choice  = str(sub.get('choice', '')      or '').strip()
        subq    = str(sub.get('subquestion', '') or '').strip()
        response = str(sub.get('response', '')   or '').strip()

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
            # Checkbox question
            for i in range(q_row, end_row):
                cell = all_rows[i][1] if len(all_rows[i]) > 1 else None
                if cell and cell.value and str(cell.value).startswith(CHECKBOX):
                    option_text = str(cell.value)[1:].strip()
                    if texts_match(choice, option_text):
                        cell.value = CHECKED + ' ' + option_text
                        success = True

        elif choice and subq:
            # Matrix question
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
                            header_cols[j] = str(cell.value)
                    break

            if header_row is not None:
                target_col = None
                for col_idx, header_text in header_cols.items():
                    if texts_match(choice, header_text):
                        target_col = col_idx
                        break

                if target_col is not None:
                    for i in range(header_row + 1, end_row):
                        row = all_rows[i]
                        label = row[1].value if len(row) > 1 else None
                        if label and texts_match(subq, str(label)):
                            if len(row) > target_col:
                                cell = row[target_col]
                                if str(cell.value or '').strip() == RADIO:
                                    cell.value = SELECTED
                                    success = True
                            break

        if success:
            filled_qids.append(parent_qid if parent_qid != qid else qid)

    return filled_qids


def get_all_excel_qids(wb):
    results = []
    for section, sheet_name in SHEET_SECTION_MAP.items():
        try:
            ws = wb[sheet_name]
        except KeyError:
            continue
        for qid, row_idx in parse_question_locs(ws):
            all_rows = list(ws.iter_rows())
            text = str(all_rows[row_idx][1].value or '').strip() if row_idx < len(all_rows) else ''
            results.append({'qid': qid, 'section': section, 'text': text})
    return results


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

        # Group by section
        by_section = {}
        for s in unique_subs:
            by_section.setdefault(s.get('section', ''), []).append(s)

        # Fill template
        wb = load_workbook(TEMPLATE_PATH)
        all_filled = []
        for section, section_subs in by_section.items():
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
