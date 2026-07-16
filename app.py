import os, csv, re, io, base64
from flask import Flask, request, jsonify
from openpyxl import load_workbook

app = Flask(__name__)

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
API_DIR      = os.path.join(os.path.dirname(__file__), 'api')
TEMPLATE_PATH = os.path.join(API_DIR, 'template.xlsx')
CSV_PATH      = os.path.join(API_DIR, 'cop_data_malaysia.csv')

# ---------------------------------------------------------------------------
# Sheet / section config
# ---------------------------------------------------------------------------
SHEET_SECTION_MAP = {
    'Governance':               ' Governance',
    'Human Rights and Labour':  'Human Rights & Labour',
    'Environment':              'Environment',
    'Anti-Corruption':          ' Anti-Corruption',
}

SUCCESS_STORIES_SHEET = 'Success Stories & Future Priori'

Q_PATTERN = re.compile(
    r'^(?:\(Optional\)\s*)?'
    r'(G\d+(?:\.\d+)?|HR/L\d+(?:\.\d+)?|E\d+(?:\.\d+)?|AC\d+(?:\.\d+)?'
    r'|C\d+|S\d+|R\d+)[.\s]'
)
SUFFIX_PATTERN = re.compile(r'^(.*?\d+)([A-Z]+)$')

# ---------------------------------------------------------------------------
# 2025 → 2026 remappings
# ---------------------------------------------------------------------------
QUESTION_ID_MAP_2025_TO_2026 = {
    'G12': 'G13',
    'G13': 'G14',
    'E5':  'E7',
    'E7':  'E8',
    'E10': 'E11',
}

CHECKBOX = '❑'
CHECKED  = '☑'
RADIO    = '🔾'
SELECTED = '🔘'

# ---------------------------------------------------------------------------
# CSV data loading
# ---------------------------------------------------------------------------

def _load_csv():
    """Load cop_data_malaysia.csv → dict: company_name → list of answer rows."""
    data = {}
    if not os.path.exists(CSV_PATH):
        return data
    with open(CSV_PATH, encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            name = row.get('NAME', '').strip()
            if not name:
                continue
            data.setdefault(name, []).append(row)
    return data


# Load once at startup; Vercel re-uses the instance across warm requests.
_CSV_DATA = _load_csv()


def get_companies():
    return sorted(_CSV_DATA.keys())


def get_submissions(company_name):
    """Return list of dicts with keys: SECTION, QUESTION_ID, SUBQUESTION, CHOICE, RESPONSE."""
    rows = _CSV_DATA.get(company_name, [])
    seen, unique = set(), []
    for r in rows:
        key = (r.get('QUESTION_ID',''), r.get('SUBQUESTION',''),
               r.get('CHOICE',''), r.get('RESPONSE',''))
        if key not in seen:
            seen.add(key)
            unique.append(r)
    return unique

# ---------------------------------------------------------------------------
# Text matching helpers
# ---------------------------------------------------------------------------

def normalize(s):
    return re.sub(r'\s+', ' ', str(s or '').lower().strip()
                  .replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' '))


def texts_match(a, b):
    a, b = normalize(a), normalize(b)
    if not a or not b:
        return False
    return a == b or b.startswith(a) or a.startswith(b)

# ---------------------------------------------------------------------------
# Excel helpers
# ---------------------------------------------------------------------------

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
        qid      = sub['QUESTION_ID'].strip()
        choice   = str(sub.get('CHOICE', '')    or '').strip()
        subq     = str(sub.get('SUBQUESTION','')or '').strip()
        response = str(sub.get('RESPONSE', '')  or '').strip()

        # Apply 2025→2026 question ID remap
        qid = QUESTION_ID_MAP_2025_TO_2026.get(qid, qid)

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
            for i in range(q_row, end_row):
                cell = all_rows[i][1] if len(all_rows[i]) > 1 else None
                if cell and cell.value and 'Please provide additional information' in str(cell.value):
                    cell.value = 'Response: ' + response
                    success = True
                    break

        elif choice and not subq:
            for i in range(q_row, end_row):
                cell = all_rows[i][1] if len(all_rows[i]) > 1 else None
                if not cell or not cell.value:
                    continue
                val = str(cell.value)
                if val.startswith(CHECKBOX):
                    option_text = val[len(CHECKBOX):].strip()
                    if texts_match(choice, option_text):
                        cell.value = CHECKED + ' ' + option_text
                        success = True
                elif val.startswith(RADIO):
                    option_text = val[len(RADIO):].strip()
                    if texts_match(choice, option_text):
                        cell.value = SELECTED + ' ' + option_text
                        success = True
                        break

        elif choice and subq:
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
                target_data_row = None
                for i in range(header_row + 1, end_row):
                    label = all_rows[i][1].value if len(all_rows[i]) > 1 else None
                    if label and texts_match(subq, str(label)):
                        target_data_row = i
                        break

                if target_data_row is not None:
                    data_row = all_rows[target_data_row]
                    choice_n = normalize(choice)

                    target_col = None
                    for col_idx, header_text in header_cols.items():
                        if texts_match(choice_n, header_text):
                            target_col = col_idx
                            break

                    if target_col is not None:
                        if len(data_row) > target_col:
                            cell = data_row[target_col]
                            cell_val = str(cell.value or '').strip()
                            if cell_val == RADIO:
                                cell.value = SELECTED; success = True
                            elif cell_val == CHECKBOX:
                                cell.value = CHECKED; success = True

                    elif choice.lower() == 'known' and response:
                        for j in range(2, len(data_row)):
                            cell = data_row[j]
                            if '_' in str(cell.value or ''):
                                cell.value = response; success = True; break

                    elif choice.lower() in ('unknown', 'not applicable'):
                        for col_idx, header_text in header_cols.items():
                            if normalize(header_text) in ('unknown', 'not applicable'):
                                if len(data_row) > col_idx:
                                    cell = data_row[col_idx]
                                    cell_val = str(cell.value or '').strip()
                                    if cell_val == RADIO:
                                        cell.value = SELECTED; success = True
                                    elif cell_val == CHECKBOX:
                                        cell.value = CHECKED; success = True
                                break

        if success:
            filled_qids.append(parent_qid if parent_qid != qid else qid)

    return filled_qids


def get_all_excel_qids(wb):
    results = []
    sheets = list(SHEET_SECTION_MAP.items()) + [('Success Stories', SUCCESS_STORIES_SHEET)]
    seen_sheets = set()
    for section, sheet_name in sheets:
        if sheet_name in seen_sheets:
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

# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

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
    return jsonify(get_companies())


@app.route('/api/generate', methods=['POST', 'OPTIONS'])
def generate():
    if request.method == 'OPTIONS':
        return '', 200
    try:
        body = request.get_json() or {}
        company_name = str(body.get('company_name', '')).strip()
        if not company_name:
            return jsonify({'error': 'company_name is required'}), 400

        subs = get_submissions(company_name)
        if not subs:
            return jsonify({'error': f'No 2025 CoP data found for "{company_name}". They may not have submitted a 2025 CoP yet.'}), 404

        # Group by section — S* questions go to Success Stories sheet
        by_section: dict = {}
        for s in subs:
            section = str(s.get('SECTION', '')).strip()
            qid     = str(s.get('QUESTION_ID', '')).strip()
            if re.match(r'^S\d+', qid):
                section = '_SuccessStories'
            by_section.setdefault(section, []).append(s)

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
            'filename':       f'CoP_2026_{safe_name}.xlsx',
            'total_filled':   len(filled_set),
            'total_questions': len(all_excel_qids),
            'pending':        pending,
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, port=5001)
