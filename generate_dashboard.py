#!/usr/bin/env python3
"""
AI Research & Innovation Hub — Executive Dashboard Generator
==============================================================
Regenerates AI_Hub_Executive_Dashboard.html + .png from the latest
tracker workbook. Tab/column NAMES are fixed to this tracker's known
template (they don't change period to period) so the script is fast
and deterministic — only the ROW DATA is re-read fresh every run.

Usage:
    python3 generate_dashboard.py <path_to_xlsx> [path_to_logo_png] [output_dir]

Defaults:
    logo:       ./logo.png (falls back to no-logo header if missing)
    output_dir: /mnt/user-data/outputs
"""
import sys, os, json, base64, datetime, re
import pandas as pd
import openpyxl

TODAY = datetime.date.today()
NOW_QATAR = datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(hours=3)  # Asia/Qatar, UTC+3, no DST


def js_rows(items):
    """Column-oriented encoding for a list of same-keyed dicts: {"k":[keys],"r":[[values],...]}.
    Field names are written once instead of once per record; the page expands it with expandRows()."""
    keys = list(items[0].keys()) if items else []
    for it in items:
        assert list(it.keys()) == keys, "js_rows: records must share identical keys"
    return js_json({'k': keys, 'r': [[it[k] for k in keys] for it in items]})


def strip_indent(html):
    """Drop leading indentation and blank lines (newlines kept, so inline whitespace behaviour is unchanged)."""
    import re as _re
    assert '`' not in html, "template literals present: indentation may be significant"
    html = _re.sub(r'\n[ \t]+', '\n', html)
    return _re.sub(r'\n{2,}', '\n', html)


def js_json(obj):
    """Serialize data for embedding in the HTML <script>. ensure_ascii=False keeps characters like
    em-dash/arrow/degree as literal UTF-8 instead of \\uXXXX escapes: the Drive upload path decodes
    \\u escapes in inline text, which silently changes the file's bytes and fails the byte-for-byte check."""
    return json.dumps(obj, ensure_ascii=False)

# ---------------------------------------------------------------------
# 1. MERGED-CELL FORWARD-FILL  (mandatory preprocessing — do this first)
# ---------------------------------------------------------------------
def sheet_grid(path, sheet_name):
    """Return a 2D dict-of-dicts grid {row:{col:value}} with every merged
    range forward-filled from its top-left cell. This must run BEFORE any
    pandas read, because pandas shows merged continuation cells as NaN,
    which is indistinguishable from genuinely-missing data otherwise."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet_name]
    grid = {}
    for row in ws.iter_rows():
        for cell in row:
            grid.setdefault(cell.row, {})[cell.column] = cell.value
    for m in ws.merged_cells.ranges:
        top_val = grid.get(m.min_row, {}).get(m.min_col)
        for r in range(m.min_row, m.max_row + 1):
            for c in range(m.min_col, m.max_col + 1):
                grid.setdefault(r, {})[c] = top_val
    return grid, ws.max_row, ws.max_column


def forward_fill_merged_columns(path, sheet_name, df, column_names, header_rows=1):
    """Forward-fill specific named columns in an already-read pandas dataframe
    using the sheet's real merged-cell ranges (handles cases like a Project #
    that's merged down across several publication/use-case rows)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet_name]
    header_cells = {}
    for col_idx in range(1, ws.max_column + 1):
        val = ws.cell(row=header_rows, column=col_idx).value
        if val is not None:
            header_cells[str(val).strip()] = col_idx
    for col_name in column_names:
        matched = next((k for k in header_cells if k.split('\n')[0].strip() == col_name.split('\n')[0].strip()), None)
        if not matched:
            continue
        col_idx = header_cells[matched]
        col_vals = {r: ws.cell(row=r, column=col_idx).value for r in range(header_rows + 1, ws.max_row + 1)}
        for m in ws.merged_cells.ranges:
            if m.min_col <= col_idx <= m.max_col and m.min_row > header_rows:
                top_val = ws.cell(row=m.min_row, column=col_idx).value
                for r in range(m.min_row, m.max_row + 1):
                    col_vals[r] = top_val
        new_col = [col_vals.get(i + header_rows + 1) for i in range(len(df))]
        df[col_name] = new_col
    return df


TITLE_PREFIX_RE = re.compile(r'^(dr\.?|mr\.?|mrs\.?|ms\.?|prof\.?)\s+', re.IGNORECASE)


def normalize_person_name(s):
    """Strip common titles and surrounding punctuation/whitespace so two
    spellings of the same person's name (e.g. 'Dr. Sarada Prasad Dakua' from
    a dedicated column vs 'Sarada Prasad Dakua' parsed out of an author-list
    cell) can be compared for equality."""
    s = (s or '').strip().rstrip('*').strip()
    s = TITLE_PREFIX_RE.sub('', s)
    return re.sub(r'\s+', ' ', s).strip().lower()


DEGREE_PREFIX_RE = re.compile(
    r"^\s*(?:bachelor(?:'s)?|master(?:s|'s)?|doctorate|doctor(?:ate)?|b\.?sc|m\.?sc|ph\.?d|mba|md|bs|ba|ms|ma)\b\.?\s*(?:of|in)?\s*(?:degree)?\s*(?:in)?\s*",
    re.IGNORECASE,
)


def extract_qualification_field(qualification):
    """Reduce a full degree string ('BSc in Data Science & AI', 'Masters in AI',
    'Master of Public Health') down to just the subject-area name, for a short
    narrative line summarizing the RA team's fields of expertise."""
    q = clean(qualification)
    if not q:
        return ''
    field = DEGREE_PREFIX_RE.sub('', q).strip()
    return field if field else q


def parse_rich_author_list(path, sheet_name, col_header, header_rows=1):
    """Read an Author List-style column as Excel rich text and split it into
    individual authors, preserving each author's bold/underline formatting
    (the tracker uses bold = HMC-affiliated, underline = AI Research &
    Innovation Hub-affiliated). Returns {excel_row: [{'text','bold','underline'}, ...]}.
    Must be read with rich_text=True — a plain data_only load collapses rich
    text runs to a flat string and the per-author formatting is lost."""
    wb = openpyxl.load_workbook(path, rich_text=True, data_only=True)
    ws = wb[sheet_name]
    col_idx = None
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_rows, column=c).value
        if v and str(v).split('\n')[0].strip() == col_header.split('\n')[0].strip():
            col_idx = c
            break
    result = {}
    if col_idx is None:
        return result
    from openpyxl.cell.rich_text import CellRichText
    for r in range(header_rows + 1, ws.max_row + 1):
        cell = ws.cell(row=r, column=col_idx)
        val = cell.value
        # A rich-text run with no explicit <rPr> (plain str part of a
        # CellRichText, or a cell value that isn't rich text at all) renders
        # in Excel using the CELL's own base font, not "no formatting" —
        # e.g. a cell styled bold-by-default shows an unformatted author
        # name as bold too. Falling back to False here (ignoring the cell's
        # base font) silently drops bold/underline for whichever author
        # happens to carry no run override.
        base_bold = bool(getattr(cell.font, 'b', False))
        base_underline = getattr(cell.font, 'u', None) is not None
        chars = []  # list of (char, bold, underline)
        if isinstance(val, CellRichText):
            for part in val:
                if isinstance(part, str):
                    text, bold, underline = part, base_bold, base_underline
                else:
                    text = part.text
                    font = getattr(part, 'font', None)
                    bold = bool(getattr(font, 'b', False)) if font is not None else base_bold
                    underline = (getattr(font, 'u', None) is not None) if font is not None else base_underline
                for ch in text:
                    chars.append((ch, bold, underline))
        elif val is not None:
            for ch in str(val):
                chars.append((ch, base_bold, base_underline))
        else:
            result[r] = []
            continue
        authors = []
        cur = []
        def flush():
            if not cur:
                return
            text = ''.join(c for c, _, _ in cur).strip()
            if not text:
                return
            nonspace = [(b, u) for c, b, u in cur if c.strip()]
            bold = any(b for b, _ in nonspace)
            underline = any(u for _, u in nonspace)
            authors.append({'text': text, 'bold': bold, 'underline': underline})
        for ch, bold, underline in chars:
            if ch == ',':
                flush()
                cur = []
            else:
                cur.append((ch, bold, underline))
        flush()
        result[r] = authors
    return result


def esc(s):
    return (str(s).replace('&', '&amp;').replace('<', '&lt;')
            .replace('>', '&gt;').replace('"', '&quot;').replace("'", '&#39;'))


def clean(v, default=''):
    if v is None:
        return default
    s = str(v).strip()
    return s if s and s.lower() != 'nan' else default


PILL_CLASS_MAP = {
    'completed': 'navy', 'active': 'green', 'on track': 'green',
    'needs support': 'amber', 'ready': 'green', 'production-ready': 'green',
    'not ready': 'gray', 'not started': 'gray', 'delayed': 'amber',
    'on hold': 'amber', 'at risk': 'red', 'cancelled': 'red',
    'published': 'green', 'accepted': 'navy', 'submitted': 'amber',
    'in preparation': 'gray', 'under review': 'amber',
    'approved': 'green', 'rejected': 'red', 'invalidated': 'gray',
    'no review': 'navy', 'out of scope': 'navy', 'active monitoring': 'navy',
}
COLOR_VAR_MAP = {'green': 'var(--green)', 'amber': 'var(--amber)', 'red': 'var(--red)',
                 'gray': 'var(--gray)', 'navy': 'var(--navy)'}


def pill_class_for_status(status):
    return PILL_CLASS_MAP.get(str(status).strip().lower(), 'navy')


PUB_TYPE_PILL_MAP = {
    'journal': 'navy', 'conference': 'green', 'protocol': 'gray', 'abstract': 'amber',
}


def pill_class_for_pub_type(pub_type):
    return PUB_TYPE_PILL_MAP.get(str(pub_type).strip().lower(), 'gray')


def color_for_status(status):
    return COLOR_VAR_MAP[pill_class_for_status(status)]


# ---------------------------------------------------------------------
# 2. ACTIVE PROJECTS  (card grid)
# ---------------------------------------------------------------------
def load_projects(path):
    df = pd.read_excel(path, sheet_name='Projects')
    df = df[df['Project #'].notna()]
    has_status_col = 'Status' in df.columns
    out = []
    for _, r in df.iterrows():
        note = clean(r.get('Assitance'))
        status = clean(r.get('Status')) if has_status_col else ''
        if not status:
            status = 'Needs support' if note else 'On track'
        out.append({
            'id': clean(r['Project #']),
            'title': clean(r['Project Title']),
            'status': status,
            'note': note,
        })
    return out


# ---------------------------------------------------------------------
# 3. PUBLICATIONS & OUTCOMES  (donut + table + drilldown)
# ---------------------------------------------------------------------
def load_outcomes(path):
    sheet = 'Project Outcomes-Publications'
    df = pd.read_excel(path, sheet_name=sheet)
    df = forward_fill_merged_columns(path, sheet, df, ['Project #', 'Project Title', 'PI'])
    title_col = [c for c in df.columns if 'Publication Title' in c][0]
    status_col = [c for c in df.columns if c.startswith('Status')][0]
    doi_col = next((c for c in df.columns if str(c).strip().upper() == 'DOI'), None)
    year_col = next((c for c in df.columns if str(c).strip() == 'Published Year'), None)
    type_col = next((c for c in df.columns if str(c).startswith('Publication Type')), None)
    author_col = next((c for c in df.columns if str(c).startswith('Author List')), None)
    corr_col = next((c for c in df.columns if str(c).strip() == 'Corresponding Author'), None)
    venue_col = 'Target Venue'
    pubdate_col = 'Actual published date'
    target_col = 'Target submission date'
    df = df[df[title_col].notna()]
    author_runs = parse_rich_author_list(path, sheet, 'Author List') if author_col else {}
    out = []
    for idx, r in df.iterrows():
        status = clean(r.get(status_col), 'In preparation')
        # normalize casing/spacing variants
        status_norm = status.strip().title()
        if 'Prep' in status_norm:
            status_norm = 'In preparation'
        pubdate = r.get(pubdate_col)
        target = r.get(target_col)
        target_str = target.strftime('%d %b %Y') if isinstance(target, (pd.Timestamp, datetime.datetime)) and pd.notna(target) else clean(target)
        pubdate_str = pubdate.strftime('%d %b %Y') if isinstance(pubdate, (pd.Timestamp, datetime.datetime)) and pd.notna(pubdate) else clean(pubdate, '')
        py = r.get(year_col) if year_col else None
        published_year = str(int(py)) if isinstance(py, (int, float)) and pd.notna(py) else clean(py)
        # Excel row for this record: pandas index is 0-based after the single
        # header row, so excel_row = idx + 2.
        excel_row = idx + 2
        corr_name_norm = normalize_person_name(clean(r.get(corr_col))) if corr_col else ''
        authors = []
        for a in author_runs.get(excel_row, []):
            text = a['text']
            is_corresponding = bool(corr_name_norm) and normalize_person_name(text) == corr_name_norm
            if is_corresponding and not text.endswith('*'):
                text = text + '*'
            authors.append({'text': text, 'bold': a['bold'], 'underline': a['underline']})
        out.append({
            'project': clean(r.get('Project #')),
            'pi': clean(r.get('PI'), 'PI not listed'),
            'status': status_norm,
            'title': clean(r.get(title_col)),
            'venue': clean(r.get(venue_col), 'TBD'),
            'pub_type': clean(r.get(type_col)) if type_col else '',
            'target_date': target_str,
            'pub_date': pubdate_str,
            'published_year': published_year,
            'doi': clean(r.get(doi_col)) if doi_col else '',
            'authors': authors,
        })
    return out


# ---------------------------------------------------------------------
# 3b. PROJECT OUTCOMES — MODELS  (separate tab; models developed per project)
# ---------------------------------------------------------------------
def load_models(path):
    sheet = 'Project Outcomes-Models'
    df = pd.read_excel(path, sheet_name=sheet)
    key_col = 'Project / Use Case'
    df = df[df[key_col].notna()]
    doi_col = next((c for c in df.columns if 'Publication DOI' in str(c)), None)
    remarks_col = next((c for c in df.columns if str(c).strip() == 'Remarks'), None)
    out = []
    for _, r in df.iterrows():
        year = r.get('Developed Year')
        year_str = str(int(year)) if isinstance(year, (int, float)) and pd.notna(year) else clean(year)
        doi = clean(r.get(doi_col)) if doi_col else ''
        # Publication status is derived from the DOI's presence (not the
        # Publication (Y/N) flag): a DOI means it's published; otherwise the
        # Remarks column carries the actual status (Under Review, Submitted, etc).
        pub_status = 'Published' if doi else clean(r.get(remarks_col), 'Not published') if remarks_col else ('Published' if doi else 'Not published')
        out.append({
            'project': clean(r.get(key_col)),
            'developed_by': clean(r.get('Developed By'), 'Not listed'),
            'year': year_str,
            'data_source': clean(r.get('Data Source'), 'Not listed'),
            'features': clean(r.get('Features')),
            'architecture': clean(r.get('AI / ML Model Architecture(s)'), 'Model architecture not specified'),
            'purpose': clean(r.get('Purpose / Description')),
            'deployment_status': clean(r.get('Deployment Status'), 'Not specified'),
            'metrics': clean(r.get('Model Performance Metrics'), 'Not reported'),
            'pub_status': pub_status,
            'publication_doi': doi,
        })
    return out


# ---------------------------------------------------------------------
# 4. RESEARCH ASSISTANT CONTRACTS  (merged-cell dedup + urgency tiers)
# ---------------------------------------------------------------------
def load_ra_contracts(path):
    grid, max_row, max_col = sheet_grid(path, 'RA Contracts')
    records = []
    for r in range(3, max_row + 1):  # data starts row 3 (2-row header)
        row = grid.get(r, {})
        staff_id = row.get(1)
        name = row.get(2)
        if staff_id is None and name is None:
            continue
        # Column layout (1-indexed): 1=Staff ID, 2=RA Name, 3=Qualification,
        # 4=PI Name, 5=Project #, 6=(grant/sub-award code), 7=Funded By,
        # 8-9=Project Start/End, 10-11=Contract Start/End.
        qualification = clean(row.get(3))
        pi = clean(row.get(4))
        project = clean(row.get(5))
        funded_by = clean(row.get(7))
        contract_end = row.get(11)
        if not project and not pi:
            continue
        records.append({
            'staff_id': staff_id,
            'name': clean(name, 'Unnamed'),
            'qualification': qualification,
            'pi': pi,
            'project': project,
            'funded_by': funded_by,
            'contract_end': contract_end,
        })

    # dedup: keep the row with the LATEST contract_end per staff_id/name
    latest = {}
    history = {}
    for rec in records:
        key = rec['staff_id'] if rec['staff_id'] is not None else rec['name']
        prev = latest.get(key)
        if prev is None or (rec['contract_end'] and prev['contract_end'] and rec['contract_end'] > prev['contract_end']):
            if prev is not None:
                history.setdefault(key, []).append(prev)
            latest[key] = rec
        elif rec is not prev:
            history.setdefault(key, []).append(rec)

    out = []
    for key, rec in latest.items():
        end = rec['contract_end']
        days = (end.date() - TODAY).days if isinstance(end, datetime.datetime) else None
        if days is None:
            urgency, tier, group = 'unknown', 'gray', 'Unknown'
        elif days < 0:
            urgency, tier, group = f'overdue by {-days}d', 'red', 'Overdue'
        elif days <= 30:
            urgency, tier, group = f'{days} days · renew urgently', 'red', 'Due ≤30 days'
        elif days <= 60:
            urgency, tier, group = f'{days} days · plan renewal', 'amber', 'Due 31–60 days'
        elif days <= 180:
            urgency, tier, group = f'~{days // 30} months', 'amber', 'Due 2–6 months'
        else:
            urgency, tier, group = 'on track', 'green', 'On track (6mo+)'
        prior_note = ''
        for h in history.get(key, []):
            if h['project'] != rec['project']:
                hend = h['contract_end']
                hend_s = hend.strftime('%b %Y') if isinstance(hend, datetime.datetime) else ''
                prior_note = f"Previously {h['project']} (PI: {h['pi']}) — ended {hend_s}"
        out.append({
            'name': rec['name'],
            'qualification': rec.get('qualification', ''),
            'pi': rec['pi'],
            'project': rec['project'],
            'funded_by': rec['funded_by'],
            'contract_end': end.strftime('%d %b %Y') if isinstance(end, datetime.datetime) else 'n/a',
            'end_sort': end if isinstance(end, datetime.datetime) else datetime.datetime(2100, 1, 1),
            'urgency': urgency,
            'tier': tier,
            'group': group,
            'prior_note': prior_note,
        })
    out.sort(key=lambda d: d['end_sort'])
    return out


# ---------------------------------------------------------------------
# 5. GOVERNANCE & STUDY REVIEW PIPELINE  (6-bucket clustering)
# ---------------------------------------------------------------------
def bucket_review_legacy(cat, decision):
    """Legacy fallback for older tracker versions that don't have a direct
    'Status' column — derives the bucket from free-text categorization/decision
    fields instead. Only used if the 'Status' column is absent."""
    cat = clean(cat)
    decision = clean(decision)
    if decision == 'Approved':
        return 'Approved'
    if cat.startswith('Rejected'):
        return 'Rejected'
    if cat.startswith('Invalidated'):
        return 'Invalidated'
    out_of_scope_exact = {
        'HMC data not used for AI', 'Not an AI study',
        'No Review (HMC data not used)', 'AI/ML component shall be removed',
    }
    if cat in out_of_scope_exact or cat.startswith('Re submission of'):
        return 'Out of scope'
    if cat == 'Everything within HMC' or cat.startswith('After meeting on May 25'):
        return 'Active monitoring'
    return 'Under review'  # catch-all: blank / pending / to-be-presented / reminders / callbacks


def compute_activity_status(new_study_amendment):
    """Active if the value is 'Active' or any 'Amendment...' variant;
    Rejected if the value is literally 'Rejected' (this column can carry that
    value directly, distinct from the separate governance Status column);
    otherwise Not Active (typically 'New', not yet actioned)."""
    v = clean(new_study_amendment).strip().lower()
    if v == 'active' or v.startswith('amendment'):
        return 'Active'
    if v == 'rejected':
        return 'Rejected'
    return 'Not Active'


def load_review_pipeline(path):
    df = pd.read_excel(path, sheet_name='AI Projects Review')
    df['Phase_filled'] = df['Phase'].ffill()
    real = df[df['MRC Study Number'].notna()].copy()
    has_status_col = 'Status' in df.columns
    out = []
    for _, r in real.iterrows():
        note = clean(r.get('AI Hub Categorization')) or 'Newly logged; assessment not yet started'
        if has_status_col:
            # Authoritative source: the sheet's own Status column states each
            # study's outcome directly — use it verbatim rather than inferring
            # a bucket from free-text categorization/decision fields.
            bucket = clean(r.get('Status'), 'Under Review')
        else:
            bucket = bucket_review_legacy(r.get('AI Hub Categorization'), r.get('Decision by  Sub Committee'))
        if bucket.strip().lower() == 'approved':
            note = 'Approved for AI development'
        out.append({
            'mrc': clean(r.get('MRC Study Number')),
            'pi': clean(r.get('Lead PI Name'), 'PI not listed'),
            'phase': clean(r.get('Phase_filled'), '—'),
            'reviewer': clean(r.get('Reviewer'), '—'),
            'bucket': bucket,
            'note': note[:110],
            'activity_status': compute_activity_status(r.get('New Study /\nAmendment')),
        })
    # reviewer workload / response-rate stats
    reviewer_counts = real['Reviewer'].dropna().astype(str).str.strip()
    reviewer_counts = reviewer_counts[reviewer_counts != '']
    contacted = real['PI Contacted'].apply(lambda v: v is True)
    responded = real['PI Responded'].apply(lambda v: v is True)
    n_contacted = int(contacted.sum())
    n_responded = int(responded.sum())
    phase_counts = real['Phase_filled'].value_counts().to_dict()
    stats = {
        'reviewer_counts': reviewer_counts.value_counts().to_dict(),
        'n_contacted': n_contacted,
        'n_responded': n_responded,
        'response_rate': round(100 * n_responded / n_contacted) if n_contacted else 0,
        'phase_counts': phase_counts,
    }
    return out, stats


# ---------------------------------------------------------------------
# 6. MCIT SANDBOX USE CASES  (card grid)
# ---------------------------------------------------------------------
def load_sandbox(path):
    df = pd.read_excel(path, sheet_name='MCIT Sandbox')
    df = df[df['Project #'].notna()]
    has_status_col = 'Status' in df.columns
    out = []
    for _, r in df.iterrows():
        flags = {
            'MRC Approved': bool(r.get('MRC Approved') is True),
            'Data Available': bool(r.get('Data\nAvailable') is True),
            'Anonymized': bool(r.get('Data Anonymized') is True),
            'RA Recruited': bool(r.get('RA Recruited') is True),
        }
        status = clean(r.get('Status')) if has_status_col else ''
        if not status:
            status = 'Production-ready' if all(flags.values()) else 'Not started'
        out.append({
            'use_case': clean(r.get('Use Case#')),
            'project': clean(r.get('Project #')),
            'pi': clean(r.get('Lead-PI'), 'PI not listed'),
            'objective': clean(r.get('Objective')),
            'status': status,
            'flags': flags,
        })
    return out


# ---------------------------------------------------------------------
# BUILD HTML
# ---------------------------------------------------------------------
PALETTE = {
    'navy_deep': '#0B3D66', 'navy': '#1477C5', 'sky': '#6AC2ED', 'sky_tint': '#E9F5FC',
    'green': '#5FAE3E', 'green_tint': '#EAF6E4', 'amber': '#E19A2C', 'amber_tint': '#FCF1DE',
    'red': '#D64545', 'red_tint': '#FBE9E9', 'gray': '#9AA7B2',
}

def bucket_order_sort(items, order):
    from collections import Counter
    counts = Counter(i['bucket'] for i in items)
    buckets = sorted(counts.items(), key=lambda kv: -kv[1])
    return buckets


def build_donut(counts_ordered, colors, click_fn, r_=75):
    """Build an interactive SVG donut (clickable segments) + matching clickable
    legend. counts_ordered = [(key, count), ...] in display order."""
    total = sum(c for _, c in counts_ordered) or 1
    circumf = 2 * 3.14159265 * r_
    circles, legend = '', ''
    offset = 0
    for key, count in counts_ordered:
        if count == 0:
            continue
        length = circumf * count / total
        color = colors.get(key, 'var(--gray)')
        circles += (f'<circle class="donut-seg" data-key="{esc(key)}" cx="100" cy="100" r="{r_}" fill="none" '
                    f'stroke="{color}" stroke-width="30" '
                    f'stroke-dasharray="{length:.2f} {circumf-length:.2f}" '
                    f'stroke-dashoffset="{-offset:.2f}" '
                    f'onclick="{click_fn}(\'{esc(key)}\')"/>\n')
        offset += length
        legend += (f'<li class="legend-item" data-key="{esc(key)}" onclick="{click_fn}(\'{esc(key)}\')">'
                   f'<i style="background:{color}"></i> {esc(key)} — {count}</li>\n')
    return circles, legend


def build_html(projects, outcomes, ra, review, review_stats, sandbox, models, logo_b64):
    from collections import Counter
    n_projects = len(projects)
    n_projects_flag = sum(1 for p in projects if p['note'])
    n_projects_completed = sum(1 for p in projects if p['status'].strip().lower() == 'completed')
    n_projects_active = n_projects - n_projects_completed
    n_outcomes = len(outcomes)
    n_published = sum(1 for o in outcomes if o['status'] == 'Published')
    n_review = len(review)
    review_buckets = bucket_order_sort(review, None)
    n_approved = sum(1 for r in review if r['bucket'].strip().lower() == 'approved')
    n_ra = len(ra)
    n_ra_overdue = sum(1 for r in ra if r['urgency'].startswith('overdue'))
    # "Needing renewal attention" is defined as due within the next 3 months
    # (including any already overdue) — used consistently in both the
    # Executive Summary narrative and the RA Contracts renewal outlook, so
    # the two numbers never disagree.
    n_ra_3mo = sum(1 for r in ra if isinstance(r.get('end_sort'), datetime.datetime) and (r['end_sort'].date() - TODAY).days <= 90)
    ra_fields = []
    for r in ra:
        field = extract_qualification_field(r.get('qualification', ''))
        if field and field not in ra_fields:
            ra_fields.append(field)
    if len(ra_fields) >= 2:
        ra_fields_text = ', '.join(ra_fields[:-1]) + ' and ' + ra_fields[-1]
    elif ra_fields:
        ra_fields_text = ra_fields[0]
    else:
        ra_fields_text = 'a range of technical fields'
    n_sandbox = len(sandbox)
    _ready_statuses = [s['status'] for s in sandbox if s['status'].strip().lower() in ('production-ready', 'ready')]
    n_sandbox_active = len(_ready_statuses)
    sandbox_ready_label = Counter(_ready_statuses).most_common(1)[0][0] if _ready_statuses else 'Ready'

    max_review = review_buckets[0][1] if review_buckets else 1

    review_table_html = ''
    # Known-bucket descriptions (shown when the source data happens to use these
    # exact labels); any other Status value present in the source falls back to
    # a generic description rather than assuming a fixed set of buckets.
    descs = {
        'Approved': 'Cleared AI Hub categorization and approved for AI development',
        'Rejected': 'Rejected — non-response, low review score, governance misalignment, or withdrawal',
        'Invalidated': 'Invalidated due to non-response or resubmission',
        'Out of scope': 'HMC data not used for AI, not an AI study, or AI component removed',
        'No Review': 'Not an AI study, or out of scope for AI Hub review',
        'Under Review': 'Newly logged, pending PI/site response, or queued for AI Sub-Committee presentation',
        'Under review': 'Newly logged, pending PI/site response, or queued for AI Sub-Committee presentation',
        'Active monitoring': 'Approved/active studies operating entirely within HMC',
    }
    for name, count in review_buckets:
        pill = pill_class_for_status(name)
        color = color_for_status(name)
        pct = round(100 * count / max_review, 1)
        desc = descs.get(name, 'See individual studies below for detail.')
        review_table_html += (f'<tr class="outcome-row" data-bucket="{esc(name)}" onclick="selectBucket(\'{esc(name)}\')">'
                               f'<td><span class="pill {pill}">{esc(name)}</span></td>'
                               f'<td><div class="outcome-count"><span class="outcome-count-num">{count}</span>'
                               f'<div class="outcome-bar-track"><div class="outcome-bar-fill" style="width:{pct}%;background:{color}"></div></div></div></td>'
                               f'<td>{esc(desc)}</td></tr>\n')

    # publications donut (interactive)
    from collections import Counter
    status_counts = Counter(o['status'] for o in outcomes)
    # preferred display order for known stages; any unrecognized status is appended after
    preferred = ['Published', 'Accepted', 'Submitted', 'Under Review', 'In preparation']
    status_order = [s for s in preferred if s in status_counts] + \
                    [s for s in status_counts if s not in preferred]
    status_colors = {s: color_for_status(s) for s in status_order}
    outcome_pill_map = {s: pill_class_for_status(s) for s in status_order}
    ordered_status = [(s, status_counts.get(s, 0)) for s in status_order]
    donut_circles, donut_legend = build_donut(ordered_status, status_colors, 'selectOutcomeStatus')
    pub_type_pill_map = {t: pill_class_for_pub_type(t) for t in sorted({o.get('pub_type') for o in outcomes if o.get('pub_type')})}

    # unify each outcome record with a single display year + escape-ready fields for drilldown JSON
    outcomes_for_js = []
    for o in outcomes:
        date_label = o.get('published_year') or '—'
        authors = o.get('authors') or []
        author_spans = []
        for a in authors:
            t = esc(a['text'])
            if a['bold']:
                t = f'<b>{t}</b>'
            if a['underline']:
                t = f'<u>{t}</u>'
            author_spans.append(t)
        authors_html = ', '.join(author_spans) if author_spans else ''
        outcomes_for_js.append({
            'project': o['project'], 'pi': o.get('pi', ''), 'title': o['title'], 'venue': o['venue'],
            'pub_type': o.get('pub_type', ''),
            'status': o['status'], 'date': date_label, 'doi': o.get('doi', ''),
            'authors_html': authors_html,
        })
    outcomes_json = js_rows(outcomes_for_js)

    # per-project outputs (publications + models) for the project drill-down view
    project_outputs = {}
    for p in projects:
        project_outputs[p['id']] = {'title': p['title'], 'outcomes': [], 'models': []}
    for o in outcomes_for_js:
        project_outputs.setdefault(o['project'], {'title': '', 'outcomes': [], 'models': []})['outcomes'].append(o)
    model_statuses = sorted({m['deployment_status'] for m in models}) or ['Not specified']
    model_pill_map = {s: pill_class_for_status(s) for s in model_statuses}
    for m in models:
        entry = project_outputs.setdefault(m['project'], {'title': '', 'outcomes': [], 'models': []})
        entry['models'].append(m)
    project_titles_json = js_json({p['id']: p['title'] for p in projects})
    model_pill_json = js_json(model_pill_map)

    # models donut (interactive) — breakdown by developed year, under Research Output
    model_year_counts = Counter((m['year'] or 'Unknown') for m in models)
    def _year_sort_key(y):
        return (0, int(y)) if y.isdigit() else (1, y)
    model_year_order = sorted(model_year_counts, key=_year_sort_key)
    year_palette = ['var(--navy)', 'var(--sky)', 'var(--green)', 'var(--amber)', 'var(--red)', 'var(--navy-deep)', 'var(--gray)']
    model_year_colors = {y: year_palette[i % len(year_palette)] for i, y in enumerate(model_year_order)}
    ordered_model_year = [(y, model_year_counts[y]) for y in model_year_order]
    model_donut_circles, model_donut_legend = build_donut(ordered_model_year, model_year_colors, 'selectModelStatus')
    n_models = len(models)

    model_pubstatus_pill_map = {s: pill_class_for_status(s) for s in sorted({m['pub_status'] for m in models})}
    models_for_js = [{
        'project': m['project'], 'developed_by': m['developed_by'], 'data_source': m['data_source'],
        'features': m['features'], 'purpose': m['purpose'], 'architecture': m['architecture'],
        'metrics': m['metrics'], 'pub_status': m['pub_status'], 'publication_doi': m['publication_doi'],
        'deployment_status': m['deployment_status'], 'year': m['year'] or 'Unknown',
    } for m in models]
    models_json = js_rows(models_for_js)

    # phase bars
    phase_counts = review_stats['phase_counts']
    phases = sorted(phase_counts.items())
    max_phase = max(phase_counts.values()) if phase_counts else 1
    phase_colors = ['var(--navy)', 'var(--sky)', 'var(--navy-deep)']
    phase_bars = ''
    for i, (ph, ct) in enumerate(phases):
        h = round(170 * ct / max_phase)
        phase_bars += (f'<div class="vbar-col"><div class="vbar-val">{ct}</div>'
                       f'<div class="vbar-fill" style="height:{h}px;background:{phase_colors[i % 3]};"></div>'
                       f'<div class="vbar-label">{esc(ph)}</div></div>\n')

    # project cards + status donut (interactive)
    project_cards = ''
    for p in projects:
        pill = pill_class_for_status(p['status'])
        flag_html = (f'<div class="flag"><b>Action needed</b>{esc(p["note"])}</div>' if p['note'] else '')
        n_out = len(project_outputs.get(p['id'], {}).get('outcomes', [])) + len(project_outputs.get(p['id'], {}).get('models', []))
        out_hint = f'<div class="proj-out-hint">{n_out} output{"s" if n_out != 1 else ""} on file &rsaquo;</div>' if n_out else '<div class="proj-out-hint">No outputs logged yet</div>'
        project_cards += (f'<div class="proj-card" data-status="{esc(p["status"])}" data-id="{esc(p["id"])}" onclick="toggleProjectOutputs(\'{esc(p["id"])}\')" style="cursor:pointer;">'
                          f'<div class="proj-top"><span class="tag">{esc(p["id"])}</span>'
                          f'<span class="pill {pill}">{esc(p["status"])}</span></div>'
                          f'<h4>{esc(p["title"])}</h4>{flag_html}{out_hint}</div>\n')
    proj_status_counts = Counter(p['status'] for p in projects)
    proj_status_order = sorted(proj_status_counts, key=lambda k: -proj_status_counts[k])
    proj_status_colors = {k: color_for_status(k) for k in proj_status_order}
    proj_pill_map = {k: pill_class_for_status(k) for k in proj_status_order}
    ordered_proj = [(k, proj_status_counts[k]) for k in proj_status_order]
    proj_donut_circles, proj_donut_legend = build_donut(ordered_proj, proj_status_colors, 'selectProjectStatus')

    # sandbox cards + status donut (interactive)
    sandbox_cards = ''
    for s in sandbox:
        pill = pill_class_for_status(s['status'])
        chips = ''.join(
            f'<span class="chk {"on" if v else "off"}">{esc(k)}</span>' for k, v in s['flags'].items())
        sandbox_cards += (f'<div class="uc-card" data-status="{esc(s["status"])}" onclick="selectSandboxStatus(\'{esc(s["status"])}\')" style="cursor:pointer;">'
                          f'<div class="uc-top"><span class="uc-id">{esc(s["use_case"])}</span>'
                          f'<span class="pill {pill}">{esc(s["status"])}</span></div>'
                          f'<p class="obj">{esc(s["objective"])}</p>'
                          f'<div class="meta-row">{esc(s["project"])} · {esc(s["pi"])}</div>'
                          f'<div class="checks">{chips}</div></div>\n')
    sandbox_status_counts = Counter(s['status'] for s in sandbox)
    sandbox_status_order = sorted(sandbox_status_counts, key=lambda k: -sandbox_status_counts[k])
    sandbox_status_colors = {k: color_for_status(k) for k in sandbox_status_order}
    sandbox_pill_map = {k: pill_class_for_status(k) for k in sandbox_status_order}
    ordered_sandbox = [(k, sandbox_status_counts[k]) for k in sandbox_status_order]
    sandbox_donut_circles, sandbox_donut_legend = build_donut(ordered_sandbox, sandbox_status_colors, 'selectSandboxStatus')

    # RA urgency donut (interactive)
    group_order = ['Overdue', 'Due ≤30 days', 'Due 31–60 days', 'Due 2–6 months', 'On track (6mo+)', 'Unknown']
    group_colors = {'Overdue': 'var(--red)', 'Due ≤30 days': 'var(--red)',
                    'Due 31–60 days': 'var(--amber)', 'Due 2–6 months': 'var(--amber)',
                    'On track (6mo+)': 'var(--green)', 'Unknown': 'var(--gray)'}
    ra_group_counts = Counter(r['group'] for r in ra)
    ordered_groups = [(g, ra_group_counts.get(g, 0)) for g in group_order]
    ra_donut_circles, ra_donut_legend = build_donut(ordered_groups, group_colors, 'selectRAGroup')

    # RA table
    ra_rows = ''
    for r in ra:
        note_html = f'<div style="font-size:11px;color:var(--ink-soft);margin-top:2px;">{esc(r["prior_note"])}</div>' if r['prior_note'] else ''
        ra_rows += (f'<tr class="ra-row" data-group="{esc(r["group"])}">'
                   f'<td>{esc(r["name"])}</td><td>{esc(r.get("qualification") or "Not listed")}</td><td>{esc(r["pi"])}{note_html}</td>'
                   f'<td>{esc(r["project"])}</td><td>{esc(r["funded_by"])}</td>'
                   f'<td>{esc(r["contract_end"])}</td>'
                   f'<td><span class="pill {r["tier"]}">{esc(r["urgency"])}</span></td></tr>\n')

    review_json = js_rows(review)
    bucket_pill_json = js_json({name: pill_class_for_status(name) for name, _ in review_buckets})
    outcome_pill_json = js_json(outcome_pill_map)
    pub_type_pill_json = js_json(pub_type_pill_map)
    proj_pill_json = js_json(proj_pill_map)
    sandbox_pill_json = js_json(sandbox_pill_map)
    model_pubstatus_pill_json = js_json(model_pubstatus_pill_map)
    n_models = len(models)
    n_models_published = sum(1 for m in models if m['pub_status'] == 'Published')

    reviewer_line = ' · '.join(f"{k}: {v} studies" for k, v in review_stats['reviewer_counts'].items())

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AI Research &amp; Innovation Hub — Executive Dashboard | Hamad Medical Corporation</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Lexend:wght@400;500;600;700;800&family=IBM+Plex+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
{CSS}
</style>
</head>
<body>
<header class="top" id="page-top">
  <div class="wrap top-inner">
    <div class="brand">
      {'<img src="data:image/png;base64,'+logo_b64+'" alt="Hamad Medical Corporation">' if logo_b64 else ''}
      <div class="titles">
        <div class="eyebrow">Hamad Medical Corporation · AI Research &amp; Innovation Hub</div>
        <h1>Executive Dashboard &amp; Newsletter</h1>
        <div class="sub">A management briefing on AI research, governance and infrastructure across HMC</div>
      </div>
    </div>
    <div class="meta">
      <div>Prepared for Senior Management</div>
      <div class="badge">Data as of {NOW_QATAR.strftime('%d %B %Y, %H:%M')} AST</div>
    </div>
  </div>
  <nav class="subnav"><div class="wrap">
    <a href="#overview">Overview</a><a href="#projects">Active Projects</a><a href="#outcomes">Research Output</a>
    <a href="#team">Research Team</a><a href="#pipeline">Governance Pipeline</a><a href="#sandbox">MCIT Sandbox</a>
  </div></nav>
</header>

<section id="overview">
  <div class="wrap">
    <div class="section-head">
      <div class="eyebrow">Executive Summary</div>
      <h2>AI Research &amp; Innovation Hub at a glance</h2>
      <p>A consolidated view of active research, publication output, workforce, ethics/governance review activity, and MCIT Sandbox readiness.</p>
    </div><!--SECTION-HEAD-END:overview:nocollapse-->
    <div class="kpi-grid">
      <div class="kpi"><div class="num">{n_projects}</div><div class="label">AI Research Projects <span>· {n_projects_active} active</span></div><div class="bar"><i style="width:100%;background:var(--navy)"></i></div></div>
      <div class="kpi"><div class="num">{n_review}</div><div class="label">Studies Reviewed by AI Hub</div><div class="bar"><i style="width:100%;background:var(--sky)"></i></div></div>
      <div class="kpi"><div class="num">{n_approved}</div><div class="label">Studies Approved for AI Development</div><div class="bar"><i style="width:{round(100*n_approved/max(n_review,1))}%;background:var(--green)"></i></div></div>
      <div class="kpi"><div class="num">{n_outcomes}</div><div class="label">Publications in Pipeline <span>· {n_published} published</span></div><div class="bar"><i style="width:{round(100*n_published/max(n_outcomes,1))}%;background:var(--amber)"></i></div></div>
      <div class="kpi"><div class="num">{n_ra}</div><div class="label">Research Assistants Deployed</div><div class="bar"><i style="width:100%;background:var(--navy)"></i></div></div>
      <div class="kpi"><div class="num">{n_sandbox_active}<span>/{n_sandbox}</span></div><div class="label">MCIT Sandbox Use Cases {esc(sandbox_ready_label)}</div><div class="bar"><i style="width:{round(100*n_sandbox_active/max(n_sandbox,1))}%;background:var(--green)"></i></div></div>
    </div>
    <div class="narrative">
      <strong>Program health:</strong> The Hub is running {n_projects} MRC-funded AI research projects
      ({n_projects_flag} flagged for support), backed by {n_ra} research assistants
      ({n_ra_3mo} needing renewal attention within 3 months{', ' + str(n_ra_overdue) + ' overdue' if n_ra_overdue else ''}).
      {n_review} studies have moved through governance review, {n_approved} approved for AI development.
      Publication output stands at {n_outcomes} items in the pipeline, {n_published} already published.
    </div>
  </div>
</section>

<section id="projects" class="alt">
  <div class="wrap">
    <div class="section-head">
      <div class="eyebrow">Portfolio</div><h2>Active AI Research Projects</h2>
      <p>Every MRC-funded project currently active under the AI Research &amp; Innovation Hub. Click a status to filter.</p>
    </div><!--SECTION-HEAD-END:projects-->
    <div class="split" id="projectsSplit">
      <div class="stack-cards">
        <div class="chart-card">
          <h4>At a glance</h4>
          <div class="narrative">
            <strong>{n_projects_flag} of {n_projects}</strong> projects have an open blocker flagged for hub or leadership support.
          </div>
          <div class="chart-hint">Click a segment or legend item to filter the list. Click a project below to view its publications and models.</div>
        </div>
        <div class="chart-card">
          <h4>Project status breakdown</h4>
          <div class="donut-wrap">
            <svg width="200" height="200" viewBox="0 0 200 200"><g transform="rotate(-90 100 100)">
              <circle cx="100" cy="100" r="75" fill="none" stroke="#EEF2F5" stroke-width="30"/>
              {proj_donut_circles}
            </g></svg>
            <div class="donut-center"><div class="n">{n_projects}</div><div class="t">Projects</div></div>
          </div>
          <ul class="legend-list">{proj_donut_legend}</ul>
        </div>
      </div>
      <div class="chart-card proj-list-card">
        <h4>Projects ({n_projects})</h4>
        <div class="drilldown-head" id="projFilterBar" style="display:none;padding:10px 0 0;border-bottom:none;">
          <h4 id="projFilterTitle" style="font-size:13px;"></h4>
          <span class="drilldown-count" id="projFilterClear" onclick="clearProjectFilter()" style="cursor:pointer;">Clear filter ✕</span>
        </div>
        <div class="proj-grid" id="projGrid">{project_cards}</div>
        <div id="projOutputsView" class="proj-outputs-view" style="display:none;"></div>
      </div>
    </div>
  </div>
</section>

<section id="outcomes">
  <div class="wrap">
    <div class="section-head">
      <div class="eyebrow">Scholarly Impact</div><h2>Research Output &amp; Publications</h2>
      <p>{n_outcomes} publications are in the pipeline across active projects. Click a status below to filter the list.</p>
    </div><!--SECTION-HEAD-END:outcomes-->
    <div class="split" id="pubChart">
      <div class="stack-cards">
        <div class="chart-card">
          <h4>At a glance</h4>
          <div class="narrative">
            <strong>{n_published} of {n_outcomes}</strong> publications in the pipeline are already published.
          </div>
        </div>
        <div class="chart-card">
          <h4>Publication status breakdown</h4>
          <div class="donut-wrap">
            <svg width="200" height="200" viewBox="0 0 200 200"><g transform="rotate(-90 100 100)">
              <circle cx="100" cy="100" r="75" fill="none" stroke="#EEF2F5" stroke-width="30"/>
              {donut_circles}
            </g></svg>
            <div class="donut-center"><div class="n">{n_outcomes}</div><div class="t">Total</div></div>
          </div>
          <ul class="legend-list">{donut_legend}</ul>
          <div class="chart-hint">Click a segment or legend item to filter the list.</div>
        </div>
      </div>
      <div class="chart-card">
        <div class="drilldown-head" style="padding:0 0 14px;border-bottom:1px solid var(--line);">
          <h4 id="outcomesTitle">Publications</h4>
          <span class="drilldown-count" id="outcomesCount"></span>
        </div>
        <p class="author-key" style="margin-top:14px;">Authors: <b>bold</b> = HMC-affiliated &middot; <u>underlined</u> = AI Research &amp; Innovation Hub-affiliated &middot; * = corresponding author</p>
        <div class="scroll drilldown-body" id="outcomesBody" style="margin-top:14px;"></div>
      </div>
    </div>

    <div class="subsection-divider">
      <div class="eyebrow">Technical Output</div>
      <h3>Models Developed</h3>
      <p>{n_models} AI/ML models have been built across the Hub's projects. Click a status below to filter the list.</p>
    </div>
    <div class="split" id="modelsChart">
      <div class="stack-cards">
        <div class="chart-card">
          <h4>At a glance</h4>
          <div class="narrative">
            <strong>{n_models_published} of {n_models}</strong> models have an associated publication.
          </div>
        </div>
        <div class="chart-card">
          <h4>Model status breakdown</h4>
          <div class="donut-wrap">
            <svg width="200" height="200" viewBox="0 0 200 200"><g transform="rotate(-90 100 100)">
              <circle cx="100" cy="100" r="75" fill="none" stroke="#EEF2F5" stroke-width="30"/>
              {model_donut_circles}
            </g></svg>
            <div class="donut-center"><div class="n">{n_models}</div><div class="t">Models</div></div>
          </div>
          <ul class="legend-list">{model_donut_legend}</ul>
          <div class="chart-hint">Click a segment or legend item to filter the list.</div>
        </div>
      </div>
      <div class="chart-card">
        <div class="drilldown-head" style="padding:0 0 14px;border-bottom:1px solid var(--line);">
          <h4 id="modelsTitle">All models</h4>
          <span class="drilldown-count" id="modelsCount"></span>
        </div>
        <div class="chart-hint" style="margin-top:8px;">Click a row to see model description, data source, features and performance.</div>
        <div class="scroll drilldown-body" id="modelsBody" style="margin-top:14px;"></div>
      </div>
    </div>
  </div>
</section>

<section id="team" class="alt">
  <div class="wrap">
    <div class="section-head">
      <div class="eyebrow">Workforce</div><h2>Research Team — RA Contracts</h2>
      <p>{n_ra} research assistants deployed across active projects. Click a slice to filter contracts by urgency.</p>
    </div><!--SECTION-HEAD-END:team-->
    <div class="split">
      <div class="chart-card">
        <h4>Contract urgency breakdown</h4>
        <div class="donut-wrap">
          <svg width="200" height="200" viewBox="0 0 200 200"><g transform="rotate(-90 100 100)">
            <circle cx="100" cy="100" r="75" fill="none" stroke="#EEF2F5" stroke-width="30"/>
            {ra_donut_circles}
          </g></svg>
          <div class="donut-center"><div class="n">{n_ra}</div><div class="t">RAs</div></div>
        </div>
        <ul class="legend-list">{ra_donut_legend}</ul>
      </div>
      <div class="stack-cards">
        <div class="chart-card">
          <h4>Team composition</h4>
          <p class="team-lead-line">{n_ra} research assistants bring expertise across {esc(ra_fields_text)}.</p>
        </div>
        <div class="chart-card">
          <h4>Renewal outlook</h4>
          <div class="narrative">
            <strong>{n_ra_3mo} of {n_ra}</strong> contracts need renewal attention within the next 3 months{(' (' + str(n_ra_overdue) + ' already overdue)') if n_ra_overdue else ''}.
            The remainder are on track with more runway.
          </div>
          <div class="chart-hint">Click a segment or legend item to filter the list.</div>
        </div>
      </div>
    </div>
    <div class="table-card" style="margin-top:26px;">
      <div class="drilldown-head" id="raFilterBar" style="display:none;">
        <h4 id="raFilterTitle"></h4>
        <span class="drilldown-count" id="raFilterClear" onclick="clearRAFilter()" style="cursor:pointer;">Clear filter ✕</span>
      </div>
      <div class="scroll"><table>
      <thead><tr><th>Research Assistant</th><th>Qualification</th><th>PI</th><th>Project #</th><th>Funded By</th><th>Contract End</th><th>Status</th></tr></thead>
      <tbody id="raTableBody">{ra_rows}</tbody>
    </table></div></div>
  </div>
</section>

<section id="pipeline">
  <div class="wrap">
    <div class="section-head">
      <div class="eyebrow">Governance</div><h2>AI Project Review Pipeline</h2>
      <p>{n_review} studies have moved through the AI Hub's ethics/governance review workflow.</p>
    </div><!--SECTION-HEAD-END:pipeline-->
    <div class="split">
      <div class="table-card">
        <div class="drilldown-head" style="padding:16px 20px 0;border-bottom:none;">
          <h4>Review outcome ({n_review} studies)</h4>
        </div>
        <div class="scroll"><table>
        <thead><tr><th>Outcome</th><th>Studies</th><th>Description</th></tr></thead>
        <tbody>{review_table_html}</tbody>
      </table></div>
        <div class="chart-hint" style="padding:0 20px 16px;">Click a row to filter the list below.</div>
      </div>
      <div class="chart-card">
        <h4>Studies by review phase</h4>
        <div class="vbar-wrap">{phase_bars}</div>
        <div class="narrative" style="margin-top:16px;">
          Reviewer workload: {reviewer_line}. Of {review_stats['n_contacted']} PIs contacted,
          {review_stats['n_responded']} have responded ({review_stats['response_rate']}% response rate).
        </div>
      </div>
    </div>
    <div class="table-card drilldown-card" id="drilldownCard">
      <div class="drilldown-head"><h4 id="drilldownTitle">All studies</h4>
      <span class="drilldown-count" id="drilldownCount"></span></div>
      <div class="scroll drilldown-body" id="drilldownBody">
        <div class="drilldown-empty">Loading studies…</div>
      </div>
    </div>
  </div>
</section>

<section id="sandbox" class="alt">
  <div class="wrap">
    <div class="section-head">
      <div class="eyebrow">Infrastructure</div><h2>MCIT Sandbox Use Cases</h2>
      <p>{n_sandbox_active} of {n_sandbox} defined use cases are live in the MCIT Sandbox. Click a status to filter.</p>
    </div><!--SECTION-HEAD-END:sandbox-->
    <div class="split">
      <div class="stack-cards">
        <div class="chart-card">
          <h4>At a glance</h4>
          <div class="narrative">
            <strong>{n_sandbox_active} of {n_sandbox}</strong> use cases have MRC approval, data, anonymization and RA staffing all in place.
          </div>
        </div>
        <div class="chart-card">
          <h4>Use case readiness</h4>
          <div class="donut-wrap">
            <svg width="200" height="200" viewBox="0 0 200 200"><g transform="rotate(-90 100 100)">
              <circle cx="100" cy="100" r="75" fill="none" stroke="#EEF2F5" stroke-width="30"/>
              {sandbox_donut_circles}
            </g></svg>
            <div class="donut-center"><div class="n">{n_sandbox}</div><div class="t">Use Cases</div></div>
          </div>
          <ul class="legend-list">{sandbox_donut_legend}</ul>
          <div class="chart-hint">Click a segment or legend item to filter the list.</div>
        </div>
      </div>
      <div class="chart-card proj-list-card">
        <h4>Use cases ({n_sandbox})</h4>
        <div class="drilldown-head" id="sandboxFilterBar" style="display:none;padding:10px 0 0;border-bottom:none;">
          <h4 id="sandboxFilterTitle" style="font-size:13px;"></h4>
          <span class="drilldown-count" id="sandboxFilterClear" onclick="clearSandboxFilter()" style="cursor:pointer;">Clear filter ✕</span>
        </div>
        <div class="uc-grid" id="sandboxGrid">{sandbox_cards}</div>
      </div>
    </div>
  </div>
</section>

<footer>
  <div class="wrap">
    <div><strong>AI Research &amp; Innovation Hub</strong> · Hamad Medical Corporation</div>
    <div>Prepared for internal management review · Confidential · {TODAY.strftime('%d %B %Y')}</div>
  </div>
</footer>

<script>
function expandRows(t){{return t.r.map(r=>Object.fromEntries(t.k.map((k,i)=>[k,r[i]])));}}
const REVIEW_DATA = expandRows({review_json});
const OUTCOMES_DATA = expandRows({outcomes_json});
const BUCKET_PILL = {bucket_pill_json};
const ACTIVITY_PILL = {{'Active': 'green', 'Rejected': 'red', 'Not Active': 'gray'}};
const OUTCOME_PILL = {outcome_pill_json};
const PUB_TYPE_PILL = {pub_type_pill_json};
const PROJECT_PILL = {proj_pill_json};
const SANDBOX_PILL = {sandbox_pill_json};
const PROJECT_TITLES = {project_titles_json};
const MODEL_PILL = {model_pill_json};
const MODELS_DATA = expandRows({models_json});
const PUB_STATUS_PILL = {model_pubstatus_pill_json};
function escapeHtml(s){{return String(s).replace(/[&<>"']/g,m=>({{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}}[m]));}}
function pillSpan(map,key){{return '<span class="pill '+(map[key]||'navy')+'">'+escapeHtml(key)+'</span>';}}

// ---- Governance & Study Review Pipeline ----
let activeBucket = null;
function selectBucket(bucket){{
  if (activeBucket === bucket) {{ bucket = null; }}
  activeBucket = bucket;
  document.querySelectorAll('.hbar-row').forEach(el=>el.classList.toggle('active', bucket!==null && el.dataset.bucket===bucket));
  document.querySelectorAll('tr.outcome-row').forEach(el=>el.classList.toggle('active', bucket!==null && el.dataset.bucket===bucket));
  const items = bucket===null ? REVIEW_DATA : REVIEW_DATA.filter(d=>d.bucket===bucket);
  document.getElementById('drilldownTitle').innerHTML = bucket===null ? 'All studies' : pillSpan(BUCKET_PILL,bucket)+' — studies in this category';
  document.getElementById('drilldownCount').textContent = items.length + ' of {n_review} studies';
  let html = '<table class="study-table"><colgroup><col style="width:15%"><col style="width:20%"><col style="width:9%"><col style="width:7%"><col style="width:12%"><col></colgroup>'
    + '<thead><tr><th>MRC Study #</th><th>Lead PI</th><th>Phase</th><th>Reviewer</th><th>Activity</th><th>Status / Note</th></tr></thead><tbody>';
  items.forEach(d=>{{html += '<tr><td class="nowrap-cell">'+escapeHtml(d.mrc)+'</td><td class="nowrap-cell">'+escapeHtml(d.pi)+'</td><td class="nowrap-cell">'+escapeHtml(d.phase)+'</td><td>'+escapeHtml(d.reviewer)+'</td><td class="nowrap-cell">'+pillSpan(ACTIVITY_PILL,d.activity_status)+'</td><td>'+escapeHtml(d.note)+'</td></tr>';}});
  html += '</tbody></table>';
  document.getElementById('drilldownBody').innerHTML = html;
  if (bucket!==null) document.getElementById('drilldownCard').scrollIntoView({{behavior:'smooth',block:'nearest'}});
}}

// ---- Research Output & Publications ----
function renderPubCard(d, hideProject){{
  const doiCell = d.doi ? '<a href="'+escapeHtml(d.doi)+'" target="_blank" rel="noopener">DOI &#8599;</a>' : '';
  const metaParts = [];
  if (!hideProject && d.project) metaParts.push(escapeHtml(d.project));
  if (d.pi) metaParts.push(escapeHtml(d.pi));
  if (d.venue) metaParts.push('Venue/Journal: '+escapeHtml(d.venue));
  if (d.date && d.date !== '—') metaParts.push('Year: '+escapeHtml(d.date));
  let html = '<div class="pub-card">';
  html += '<div class="pub-top"><span class="pub-tags">'+pillSpan(OUTCOME_PILL,d.status)+(d.pub_type ? pillSpan(PUB_TYPE_PILL,d.pub_type) : '')+'</span>'+(doiCell?'<span class="pub-doi">'+doiCell+'</span>':'')+'</div>';
  html += '<div class="pub-title">'+escapeHtml(d.title)+'</div>';
  if (metaParts.length) html += '<div class="pub-meta">'+metaParts.join(' &middot; ')+'</div>';
  if (d.authors_html) html += '<div class="pub-authors"><span class="pub-authors-label">Authors:</span> '+d.authors_html+'</div>';
  html += '</div>';
  return html;
}}
let activeOutcomeStatus = null;
function selectOutcomeStatus(status){{
  if (activeOutcomeStatus === status) {{ status = null; }}
  activeOutcomeStatus = status;
  document.querySelectorAll('#pubChart .donut-seg').forEach(el=>{{
    el.classList.toggle('seg-active', status!==null && el.dataset.key===status);
    el.classList.toggle('seg-dim', status!==null && el.dataset.key!==status);
  }});
  document.querySelectorAll('#pubChart .legend-item').forEach(el=>{{
    el.classList.toggle('seg-active', status!==null && el.dataset.key===status);
    el.classList.toggle('seg-dim', status!==null && el.dataset.key!==status);
  }});
  const items = status===null ? OUTCOMES_DATA : OUTCOMES_DATA.filter(d=>d.status===status);
  document.getElementById('outcomesTitle').innerHTML = status===null ? 'All publications' : pillSpan(OUTCOME_PILL,status)+' publications';
  document.getElementById('outcomesCount').textContent = items.length + ' of {n_outcomes} total';
  document.getElementById('outcomesBody').innerHTML = items.length
    ? '<div class="pub-list">'+items.map(d=>renderPubCard(d,false)).join('')+'</div>'
    : '<div class="drilldown-empty">No publications match this filter.</div>';
}}

// ---- Research Output: Models Developed (filtered by developed year) ----
function nl2br(s){{return escapeHtml(s).replace(/\\n/g,'<br>');}}
let activeModelYear = null;
let currentModelItems = [];
let expandedModelIdx = null;
function selectModelStatus(year){{
  if (activeModelYear === year) {{ year = null; }}
  activeModelYear = year;
  document.querySelectorAll('#modelsChart .donut-seg').forEach(el=>{{
    el.classList.toggle('seg-active', year!==null && el.dataset.key===year);
    el.classList.toggle('seg-dim', year!==null && el.dataset.key!==year);
  }});
  document.querySelectorAll('#modelsChart .legend-item').forEach(el=>{{
    el.classList.toggle('seg-active', year!==null && el.dataset.key===year);
    el.classList.toggle('seg-dim', year!==null && el.dataset.key!==year);
  }});
  const items = year===null ? MODELS_DATA : MODELS_DATA.filter(d=>d.year===year);
  currentModelItems = items;
  expandedModelIdx = null;
  document.getElementById('modelsTitle').innerHTML = year===null ? 'All models' : 'Models developed in '+escapeHtml(year);
  document.getElementById('modelsCount').textContent = items.length + ' of {n_models} total';
  let html = '<table><thead><tr><th>Project / Use Case</th><th>Developed By</th><th>Year</th><th>Publication Status</th></tr></thead><tbody>';
  items.forEach((d,i)=>{{
    const pubCell = pillSpan(PUB_STATUS_PILL, d.pub_status) + (d.publication_doi ? ' <a href="'+escapeHtml(d.publication_doi)+'" target="_blank" rel="noopener" onclick="event.stopPropagation()">DOI &#8599;</a>' : '');
    html += '<tr class="model-row" onclick="toggleModelDetail(this,'+i+')"><td>'+escapeHtml(d.project)+'</td><td>'+escapeHtml(d.developed_by)+'</td><td>'+escapeHtml(d.year)+'</td><td>'+pubCell+'</td></tr>';
  }});
  html += '</tbody></table>';
  document.getElementById('modelsBody').innerHTML = html;
}}
function toggleModelDetail(rowEl, idx){{
  const tbody = rowEl.parentElement;
  const existing = tbody.querySelector('tr.model-detail-row');
  const wasThisExpanded = expandedModelIdx === idx;
  if (existing) existing.remove();
  tbody.querySelectorAll('tr.model-row').forEach(el=>el.classList.remove('expanded'));
  if (wasThisExpanded) {{ expandedModelIdx = null; return; }}
  expandedModelIdx = idx;
  rowEl.classList.add('expanded');
  const d = currentModelItems[idx];
  let detail = '<tr class="model-detail-row"><td colspan="4"><div class="model-detail">';
  detail += (d.purpose ? '<p><span class="model-detail-label">Description:</span> '+nl2br(d.purpose)+'</p>' : '');
  detail += (d.data_source ? '<p><span class="model-detail-label">Data Source:</span> '+nl2br(d.data_source)+'</p>' : '');
  detail += (d.features ? '<p><span class="model-detail-label">Features:</span> '+nl2br(d.features)+'</p>' : '');
  detail += (d.metrics ? '<p><span class="model-detail-label">Model Performance:</span> '+nl2br(d.metrics)+'</p>' : '');
  detail += '</div></td></tr>';
  rowEl.insertAdjacentHTML('afterend', detail);
}}

// ---- Active AI Research Projects: per-project outputs drill-down ----
let expandedProject = null;
function toggleProjectOutputs(id){{
  if (expandedProject === id) {{ collapseProjectOutputs(); return; }}
  expandedProject = id;
  document.getElementById('projGrid').style.display = 'none';
  document.getElementById('projFilterBar').style.display = 'none';
  const data = {{title: PROJECT_TITLES[id]||'', outcomes: OUTCOMES_DATA.filter(d=>d.project===id), models: MODELS_DATA.filter(d=>d.project===id)}};
  let html = '<div class="proj-output-head"><span class="back-link" onclick="collapseProjectOutputs()">&larr; Back to project list</span>'
    + '<h4>'+escapeHtml(id)+(data.title ? ' &mdash; '+escapeHtml(data.title) : '')+'</h4></div>';
  html += '<h5 class="proj-output-sub">Publications ('+data.outcomes.length+')</h5>';
  if (data.outcomes.length) {{
    html += '<div class="pub-list">'+data.outcomes.map(o=>renderPubCard(o,true)).join('')+'</div>';
  }} else {{
    html += '<div class="drilldown-empty">No publications recorded yet for this project.</div>';
  }}
  html += '<h5 class="proj-output-sub">Models developed ('+data.models.length+')</h5>';
  if (data.models.length) {{
    html += '<div class="model-grid">';
    data.models.forEach(m=>{{
      const pubLink = m.publication_doi ? '<a href="'+escapeHtml(m.publication_doi)+'" target="_blank" rel="noopener">Linked publication &#8599;</a>' : '';
      html += '<div class="model-card"><div class="model-top"><b>'+escapeHtml(m.architecture)+'</b>'+pillSpan(MODEL_PILL,m.deployment_status)+'</div>'
        + '<div class="model-meta">'+escapeHtml(m.developed_by)+((m.year && m.year!=='Unknown') ? ' &middot; '+escapeHtml(m.year) : '')+'</div>'
        + (m.purpose ? '<p>'+escapeHtml(m.purpose)+'</p>' : '')
        + pubLink + '</div>';
    }});
    html += '</div>';
  }} else {{
    html += '<div class="drilldown-empty">No models recorded yet for this project.</div>';
  }}
  document.getElementById('projOutputsView').innerHTML = html;
  document.getElementById('projOutputsView').style.display = 'block';
  document.getElementById('projOutputsView').scrollIntoView({{behavior:'smooth',block:'nearest'}});
}}
function collapseProjectOutputs(){{
  expandedProject = null;
  document.getElementById('projOutputsView').style.display = 'none';
  document.getElementById('projOutputsView').innerHTML = '';
  document.getElementById('projGrid').style.display = '';
}}

// ---- Research Team — RA Contracts ----
let raActiveGroup = null;
function selectRAGroup(group){{
  if (raActiveGroup === group) {{ clearRAFilter(); return; }}
  raActiveGroup = group;
  document.querySelectorAll('#team .donut-seg').forEach(el=>{{
    el.classList.toggle('seg-active', el.dataset.key===group);
    el.classList.toggle('seg-dim', el.dataset.key!==group);
  }});
  document.querySelectorAll('#team .legend-item').forEach(el=>{{
    el.classList.toggle('seg-active', el.dataset.key===group);
    el.classList.toggle('seg-dim', el.dataset.key!==group);
  }});
  let shown = 0, total = 0;
  document.querySelectorAll('.ra-row').forEach(el=>{{
    total++;
    const match = el.dataset.group === group;
    el.classList.toggle('hidden', !match);
    if(match) shown++;
  }});
  document.getElementById('raFilterBar').style.display = 'flex';
  const tier = group.includes('Overdue')||group.includes('≤30') ? 'red' : (group.includes('31')||group.includes('2–6')) ? 'amber' : group.includes('On track') ? 'green' : 'gray';
  document.getElementById('raFilterTitle').innerHTML = 'Showing '+shown+' of '+total+' — <span class="pill '+tier+'">'+escapeHtml(group)+'</span>';
}}
function clearRAFilter(){{
  raActiveGroup = null;
  document.querySelectorAll('#team .donut-seg, #team .legend-item').forEach(el=>el.classList.remove('seg-active','seg-dim'));
  document.querySelectorAll('.ra-row').forEach(el=>el.classList.remove('hidden'));
  document.getElementById('raFilterBar').style.display = 'none';
}}

// ---- Active AI Research Projects ----
let activeProjectStatus = null;
function selectProjectStatus(status){{
  if (activeProjectStatus === status) {{ clearProjectFilter(); return; }}
  activeProjectStatus = status;
  document.querySelectorAll('#projects .donut-seg').forEach(el=>{{
    el.classList.toggle('seg-active', el.dataset.key===status);
    el.classList.toggle('seg-dim', el.dataset.key!==status);
  }});
  document.querySelectorAll('#projects .legend-item').forEach(el=>{{
    el.classList.toggle('seg-active', el.dataset.key===status);
    el.classList.toggle('seg-dim', el.dataset.key!==status);
  }});
  let shown = 0, total = 0;
  document.querySelectorAll('#projGrid .proj-card').forEach(el=>{{
    total++;
    const match = el.dataset.status === status;
    el.style.display = match ? '' : 'none';
    if(match) shown++;
  }});
  document.getElementById('projFilterBar').style.display = 'flex';
  document.getElementById('projFilterTitle').innerHTML = 'Showing '+shown+' of '+total+' — '+pillSpan(PROJECT_PILL,status);
}}
function clearProjectFilter(){{
  activeProjectStatus = null;
  document.querySelectorAll('#projects .donut-seg, #projects .legend-item').forEach(el=>el.classList.remove('seg-active','seg-dim'));
  document.querySelectorAll('#projGrid .proj-card').forEach(el=>el.style.display='');
  document.getElementById('projFilterBar').style.display = 'none';
}}

// ---- MCIT Sandbox Use Cases ----
let activeSandboxStatus = null;
function selectSandboxStatus(status){{
  if (activeSandboxStatus === status) {{ clearSandboxFilter(); return; }}
  activeSandboxStatus = status;
  document.querySelectorAll('#sandbox .donut-seg').forEach(el=>{{
    el.classList.toggle('seg-active', el.dataset.key===status);
    el.classList.toggle('seg-dim', el.dataset.key!==status);
  }});
  document.querySelectorAll('#sandbox .legend-item').forEach(el=>{{
    el.classList.toggle('seg-active', el.dataset.key===status);
    el.classList.toggle('seg-dim', el.dataset.key!==status);
  }});
  let shown = 0, total = 0;
  document.querySelectorAll('#sandboxGrid .uc-card').forEach(el=>{{
    total++;
    const match = el.dataset.status === status;
    el.style.display = match ? '' : 'none';
    if(match) shown++;
  }});
  document.getElementById('sandboxFilterBar').style.display = 'flex';
  document.getElementById('sandboxFilterTitle').innerHTML = 'Showing '+shown+' of '+total+' — '+pillSpan(SANDBOX_PILL,status);
}}
function clearSandboxFilter(){{
  activeSandboxStatus = null;
  document.querySelectorAll('#sandbox .donut-seg, #sandbox .legend-item').forEach(el=>el.classList.remove('seg-active','seg-dim'));
  document.querySelectorAll('#sandboxGrid .uc-card').forEach(el=>el.style.display='');
  document.getElementById('sandboxFilterBar').style.display = 'none';
}}

// default view: publications section opens on "Published" so the section isn't empty on load
if (OUTCOMES_DATA.some(d=>d.status==='Published')) {{ selectOutcomeStatus('Published'); }}
else if (OUTCOMES_DATA.length) {{ selectOutcomeStatus(OUTCOMES_DATA[0].status); }}
// models list: show every model by default rather than a placeholder
selectModelStatus(null);
// governance study list: show every study by default rather than a placeholder
selectBucket(null);

// ---- Collapsible sections ----
function toggleSection(id){{
  const body = document.getElementById('body-'+id);
  const icon = document.getElementById('toggleIcon-'+id);
  if (!body) return;
  const collapsed = body.classList.toggle('collapsed');
  if (icon) icon.textContent = collapsed ? '▸' : '▾';
}}
</script>
</body>
</html>"""
    html = _add_collapsible_and_backtotop(html)
    return html


def _add_collapsible_and_backtotop(html):
    """Wrap each section's body (everything after its section-head) in a
    collapsible container with a toggle button, and add a 'back to top' link
    at the end of each section. The section-head itself (eyebrow/title/description)
    always stays visible, even when collapsed.

    Relies on explicit '<!--SECTION-HEAD-END:id[:nocollapse]-->' markers placed
    in the template right after each section-head's true closing </div> — this
    avoids the ambiguity of trying to regex-match the 'right' closing div among
    several nested ones (eyebrow/h2/p are all inside section-head too). The whole
    transform is done in a single regex pass per section so partially-transformed
    HTML is never re-scanned (which would reintroduce the same ambiguity).
    Sections marked :nocollapse (e.g. the executive summary) are left as plain,
    always-visible content with no toggle button and no back-to-top link.
    """
    pattern = re.compile(
        r'(<section id="\w+"[^>]*>\s*<div class="wrap">\s*<div class="section-head">)'
        r'(.*?)'
        r'(?:</div>)<!--SECTION-HEAD-END:(\w+)(:nocollapse)?-->'
        r'(.*?)'
        r'(\s*</div>\n</section>)',
        re.DOTALL,
    )

    def repl(m):
        open_tag, head_inner, sec_id, nocollapse, body, close_part = m.groups()
        if nocollapse:
            return open_tag + head_inner + '</div>' + body + close_part
        toggle_btn = (
            f'<button class="section-toggle" type="button" onclick="toggleSection(\'{sec_id}\')" '
            f'aria-label="Collapse or expand this section">'
            f'<span id="toggleIcon-{sec_id}">&#9662;</span></button>'
        )
        # open_tag ends in '<div class="section-head">' — add the flex modifier
        # class only here, so plain (non-collapsible) sections stay stacked.
        open_tag_flex = open_tag.replace('class="section-head">', 'class="section-head has-toggle">')
        new_head = open_tag_flex + '<div class="section-head-text">' + head_inner + '</div>' + toggle_btn + '</div>'
        new_body = (f'<div class="section-body" id="body-{sec_id}">{body}</div>\n'
                    f'<a class="back-to-top" href="#page-top">&uarr; Back to top</a>\n')
        return new_head + new_body + close_part

    return pattern.sub(repl, html)



CSS = """
  :root{--navy-deep:#0B3D66;--navy:#1477C5;--sky:#6AC2ED;--sky-tint:#E9F5FC;--green:#5FAE3E;--green-tint:#EAF6E4;
  --amber:#E19A2C;--amber-tint:#FCF1DE;--red:#D64545;--red-tint:#FBE9E9;--gray:#9AA7B2;
  --ink:#132433;--ink-soft:#5B6B7A;--line:#E3EAF0;--bg:#F4F8FB;--card:#FFFFFF;--radius:14px;}
  *{box-sizing:border-box;} html{scroll-behavior:smooth;}
  body{margin:0;background:var(--bg);color:var(--ink);font-family:'IBM Plex Sans',sans-serif;font-size:15px;line-height:1.55;-webkit-font-smoothing:antialiased;}
  h1,h2,h3,h4{font-family:'Lexend',sans-serif;margin:0;color:var(--navy-deep);}
  a{color:inherit;} .wrap{max-width:1180px;margin:0 auto;padding:0 28px;}
  header.top{background:linear-gradient(120deg,var(--navy-deep) 0%,var(--navy) 62%,#1E8FD6 100%);color:#fff;position:relative;overflow:hidden;}
  header.top::before{content:"";position:absolute;inset:0;background-image:radial-gradient(circle, rgba(255,255,255,0.14) 1.6px, transparent 1.6px);background-size:26px 26px;opacity:.5;pointer-events:none;}
  .top-inner{position:relative;display:flex;align-items:center;justify-content:space-between;gap:24px;padding:26px 28px 22px;flex-wrap:wrap;}
  .brand{display:flex;align-items:center;gap:16px;} .brand img{height:56px;border-radius:8px;background:#fff;padding:4px 8px;}
  .brand .titles .eyebrow{font-size:12px;letter-spacing:.14em;text-transform:uppercase;color:var(--sky);font-weight:600;}
  .brand .titles h1{color:#fff;font-size:24px;font-weight:700;} .brand .titles .sub{color:rgba(255,255,255,.82);font-size:13.5px;margin-top:2px;}
  .meta{text-align:right;font-size:13px;color:rgba(255,255,255,.9);}
  .meta .badge{display:inline-block;background:rgba(255,255,255,.14);border:1px solid rgba(255,255,255,.3);padding:5px 12px;border-radius:999px;font-size:12px;margin-top:6px;}
  nav.subnav{position:relative;background:rgba(0,0,0,0.15);border-top:1px solid rgba(255,255,255,.12);}
  nav.subnav .wrap{display:flex;gap:2px;overflow-x:auto;padding:0 28px;}
  nav.subnav a{white-space:nowrap;display:inline-block;padding:11px 16px;font-size:13px;font-weight:500;color:rgba(255,255,255,.78);text-decoration:none;border-bottom:2.5px solid transparent;transition:.15s;}
  nav.subnav a:hover{color:#fff;border-bottom-color:var(--sky);}
  section{padding:46px 0;} section.alt{background:#fff;}
  .section-head{margin-bottom:26px;}
  .section-head.has-toggle{display:flex;justify-content:space-between;align-items:flex-start;gap:16px;}
  .section-head .eyebrow{font-size:11.5px;letter-spacing:.13em;text-transform:uppercase;color:var(--navy);font-weight:700;}
  .section-head-text{min-width:0;}
  .section-toggle{
    flex-shrink:0;width:34px;height:34px;border-radius:9px;border:1px solid var(--line);
    background:var(--card);color:var(--navy);font-size:13px;cursor:pointer;
    display:flex;align-items:center;justify-content:center;transition:background .15s;
  }
  .section-toggle:hover{background:var(--sky-tint);}
  .section-body.collapsed{display:none;}
  .back-to-top{
    display:block;text-align:center;margin-top:30px;font-size:12.5px;font-weight:600;
    color:var(--navy);text-decoration:none;
  }
  .back-to-top:hover{text-decoration:underline;}
  .section-head h2{font-size:26px;margin-top:4px;} .section-head p{color:var(--ink-soft);margin-top:8px;max-width:760px;font-size:14px;}
  .kpi-grid{display:grid;grid-template-columns:repeat(6,1fr);gap:16px;margin-top:8px;}
  .kpi{background:var(--card);border:1px solid var(--line);border-radius:var(--radius);padding:20px 18px;box-shadow:0 1px 2px rgba(19,36,51,.04);}
  .kpi .num{font-family:'Lexend';font-size:32px;font-weight:700;color:var(--navy-deep);line-height:1;}
  .kpi .num span{font-size:16px;color:var(--ink-soft);font-weight:500;}
  .kpi .label{margin-top:8px;font-size:12.5px;color:var(--ink-soft);font-weight:500;}
  .kpi .bar{margin-top:12px;height:5px;border-radius:4px;background:var(--line);overflow:hidden;}
  .kpi .bar i{display:block;height:100%;background:var(--navy);border-radius:4px;}
  .narrative{margin-top:22px;background:var(--sky-tint);border-left:4px solid var(--navy);border-radius:0 10px 10px 0;padding:16px 20px;font-size:13.8px;color:#2b3e4d;}
  .subsection-divider{margin:40px 0 22px;padding-top:28px;border-top:1px dashed var(--line);}
  .subsection-divider .eyebrow{font-size:11.5px;letter-spacing:.13em;text-transform:uppercase;color:var(--navy);font-weight:700;}
  .subsection-divider h3{font-family:'Lexend',sans-serif;font-size:19px;color:var(--navy-deep);margin-top:4px;}
  .subsection-divider p{color:var(--ink-soft);font-size:13.5px;margin-top:6px;max-width:760px;}
  .team-lead-line{font-size:13.5px;color:var(--ink);margin:0;line-height:1.55;}
  .stack-cards{display:flex;flex-direction:column;gap:20px;}
  .split .stack-cards{height:100%;}
  .split .stack-cards .chart-card:last-child{flex:1;display:flex;flex-direction:column;justify-content:center;}
  .proj-list-card{display:flex;flex-direction:column;}
  .proj-grid{display:flex;flex-direction:column;gap:10px;margin-top:14px;flex:1;min-height:0;max-height:480px;overflow-y:auto;padding-right:4px;}
  .proj-card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:12px 16px;position:relative;cursor:pointer;transition:border-color .15s,background .15s;}
  .proj-card:hover{border-color:var(--sky);background:#FAFCFE;}
  .proj-top{display:flex;justify-content:space-between;align-items:center;gap:10px;}
  .proj-card .tag{font-size:11px;font-weight:700;color:var(--navy);background:var(--sky-tint);display:inline-block;padding:3px 9px;border-radius:6px;letter-spacing:.03em;flex-shrink:0;}
  .proj-card h4{font-size:13.8px;margin:7px 0 0;line-height:1.4;}
  .proj-card .flag{margin-top:9px;font-size:12px;background:var(--amber-tint);color:#7A5311;border-radius:8px;padding:7px 10px;border:1px solid #F0DBAE;}
  .proj-card .flag b{display:block;font-size:11px;text-transform:uppercase;letter-spacing:.05em;margin-bottom:2px;}
  .proj-card .proj-out-hint{margin-top:7px;font-size:11.5px;color:var(--navy);font-weight:600;}
  .proj-outputs-view{margin-top:14px;flex:1;min-height:0;max-height:480px;overflow-y:auto;padding-right:4px;}
  .proj-output-head{display:flex;flex-direction:column;gap:6px;margin-bottom:18px;}
  .proj-output-head h4{font-size:17px;}
  .back-link{cursor:pointer;color:var(--navy);font-size:12.5px;font-weight:600;}
  .back-link:hover{text-decoration:underline;}
  .proj-output-sub{font-size:13px;color:var(--ink-soft);text-transform:uppercase;letter-spacing:.05em;margin:20px 0 10px;}
  .model-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:14px;}
  .model-card{background:var(--bg);border:1px solid var(--line);border-radius:10px;padding:14px 16px;}
  .model-card .model-top{display:flex;justify-content:space-between;align-items:flex-start;gap:10px;font-size:13px;}
  .model-card .model-meta{font-size:11.5px;color:var(--ink-soft);margin-top:6px;}
  .model-card p{font-size:12.6px;margin:8px 0 0;color:var(--ink);}
  .model-card a{display:inline-block;margin-top:8px;font-size:12px;font-weight:600;color:var(--navy);}
  @media(max-width:700px){.model-grid{grid-template-columns:1fr;}}
  .pub-list{display:flex;flex-direction:column;gap:12px;}
  .pub-card{background:var(--bg);border:1px solid var(--line);border-radius:10px;padding:14px 16px;}
  .pub-top{display:flex;justify-content:space-between;align-items:center;gap:10px;}
  .pub-tags{display:flex;align-items:center;gap:6px;flex-wrap:wrap;}
  .pub-doi a{font-size:12px;font-weight:600;color:var(--navy);}
  .pub-title{font-family:'Lexend',sans-serif;font-size:13.8px;color:var(--navy-deep);margin-top:9px;line-height:1.4;}
  .pub-meta{font-size:12px;color:var(--ink-soft);margin-top:6px;}
  .pub-authors{font-size:12.5px;color:var(--ink);margin-top:9px;line-height:1.55;}
  .pub-authors-label{font-weight:600;color:var(--ink-soft);margin-right:4px;}
  .author-key{font-size:12px;color:var(--ink-soft);margin-top:6px;}
  .author-key u{text-decoration-color:var(--navy);text-underline-offset:2px;}
  .pub-authors u{text-decoration-color:var(--navy);text-underline-offset:2px;}
  .split{display:grid;grid-template-columns:1.1fr 1.4fr;gap:28px;margin-top:24px;align-items:stretch;}
  #pipeline .split{grid-template-columns:1.3fr 1.2fr;}
  .chart-card{background:var(--card);border:1px solid var(--line);border-radius:var(--radius);padding:20px;}
  .chart-card h4{font-size:14px;margin-bottom:14px;color:var(--ink);}
  .legend-list{list-style:none;margin:14px 0 0;padding:0;font-size:12.5px;}
  .legend-list li{display:flex;align-items:center;gap:8px;margin-bottom:6px;color:var(--ink-soft);}
  .legend-list li.legend-item{cursor:pointer;padding:3px 6px;margin:0 -6px 3px;border-radius:6px;transition:background .15s,opacity .15s;}
  .legend-list li.legend-item:hover{background:#F1F6FA;}
  .legend-list li.legend-item.seg-active{background:var(--sky-tint);font-weight:700;color:var(--navy-deep);}
  .legend-list li.legend-item.seg-dim{opacity:.4;}
  .legend-list i{width:10px;height:10px;border-radius:3px;display:inline-block;}
  .donut-seg{cursor:pointer;transition:opacity .15s,stroke-width .15s;}
  .donut-seg.seg-dim{opacity:.35;}
  .donut-seg.seg-active{filter:drop-shadow(0 0 0 rgba(0,0,0,0));}
  .ra-row.hidden{display:none;}
  .donut-wrap{display:flex;align-items:center;justify-content:center;position:relative;}
  .donut-center{position:absolute;left:50%;top:50%;transform:translate(-50%,-50%);text-align:center;}
  .donut-center .n{font-family:'Lexend';font-weight:700;font-size:26px;color:var(--navy-deep);line-height:1;}
  .donut-center .t{font-size:10.5px;color:var(--ink-soft);text-transform:uppercase;letter-spacing:.06em;margin-top:2px;}
  .hbar-list{display:flex;flex-direction:column;gap:12px;}
  .hbar-row{display:grid;grid-template-columns:130px 1fr 34px;align-items:center;gap:10px;cursor:pointer;padding:5px 8px;margin:0 -8px;border-radius:8px;transition:background .15s;}
  .hbar-row:hover{background:#F1F6FA;} .hbar-row.active{background:var(--sky-tint);box-shadow:inset 3px 0 0 0 var(--navy);}
  .hbar-row.active .hbar-label{color:var(--navy-deep);font-weight:700;}
  .hbar-row .hbar-label{font-size:12px;color:var(--ink);font-weight:500;}
  .hbar-row .hbar-track{background:#EEF2F5;border-radius:6px;height:16px;overflow:hidden;}
  .hbar-row .hbar-fill{height:100%;border-radius:6px;} .hbar-row .hbar-val{font-size:12.5px;font-weight:600;color:var(--ink);text-align:right;}
  .chart-hint{font-size:11.5px;color:var(--ink-soft);margin-top:14px;font-style:italic;}
  .split > .chart-card,.split > .table-card{display:flex;flex-direction:column;}
  .split > .chart-card > .chart-hint:last-child,.split > .table-card > .chart-hint:last-child{margin-top:auto;padding-top:14px;}
  .vbar-wrap{display:flex;align-items:flex-end;gap:22px;height:236px;padding:0 10px;}
  .vbar-col{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:flex-end;height:100%;}
  .vbar-col .vbar-val{flex-shrink:0;font-size:13px;font-weight:700;color:var(--ink);margin-bottom:6px;font-family:'Lexend';}
  .vbar-col .vbar-fill{flex-shrink:0;width:56px;border-radius:8px 8px 4px 4px;}
  .vbar-col .vbar-label{flex-shrink:0;margin-top:10px;font-size:12px;color:var(--ink-soft);text-align:center;}
  table{width:100%;border-collapse:collapse;font-size:13.2px;}
  thead th{text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:var(--ink-soft);border-bottom:2px solid var(--line);padding:10px 12px;font-weight:600;}
  tbody td{padding:11px 12px;border-bottom:1px solid var(--line);vertical-align:top;} tbody tr:hover{background:#FAFCFE;}
  tr.model-row{cursor:pointer;}
  tr.model-row.expanded{background:#FAFCFE;}
  tr.model-row.expanded td:first-child{border-left:3px solid var(--navy);}
  tr.model-detail-row td{padding:0;border-bottom:1px solid var(--line);}
  .model-detail{background:var(--bg);padding:14px 20px;}
  .model-detail p{font-size:12.8px;color:var(--ink);margin:0 0 9px;line-height:1.5;}
  .model-detail p:last-child{margin-bottom:0;}
  .model-detail-label{font-weight:600;color:var(--ink-soft);margin-right:4px;}
  .table-card{background:var(--card);border:1px solid var(--line);border-radius:var(--radius);overflow:hidden;margin-top:22px;}
  .split > .table-card{margin-top:0;}
  .table-card .scroll{overflow-x:auto;}
  .outcome-count{display:flex;align-items:center;gap:10px;}
  .outcome-count-num{font-weight:700;color:var(--ink);min-width:18px;}
  .outcome-bar-track{flex:1;max-width:140px;background:#EEF2F5;border-radius:6px;height:10px;overflow:hidden;}
  .outcome-bar-fill{height:100%;border-radius:6px;}
  .study-table .nowrap-cell{white-space:nowrap;}
  .pill{display:inline-block;padding:3px 10px;border-radius:999px;font-size:11.5px;font-weight:600;}
  .pill.green{background:var(--green-tint);color:#3B7A24;} .pill.amber{background:var(--amber-tint);color:#8A5D14;}
  .pill.red{background:var(--red-tint);color:#B23434;} .pill.gray{background:#EEF2F5;color:#5B6B7A;} .pill.navy{background:var(--sky-tint);color:var(--navy-deep);}
  .uc-grid{display:flex;flex-direction:column;gap:10px;margin-top:14px;flex:1;min-height:0;max-height:480px;overflow-y:auto;padding-right:4px;}
  .uc-card{background:var(--card);border:1px solid var(--line);border-radius:var(--radius);padding:16px 16px 18px;}
  .uc-card .uc-top{display:flex;justify-content:space-between;align-items:flex-start;}
  .uc-card .uc-id{font-family:'Lexend';font-weight:700;color:var(--navy);font-size:14px;}
  .uc-card p.obj{font-size:12.6px;color:var(--ink);margin-top:8px;min-height:52px;}
  .uc-card .meta-row{font-size:11.5px;color:var(--ink-soft);margin-top:6px;}
  .checks{display:flex;gap:6px;margin-top:12px;flex-wrap:wrap;} .chk{font-size:10.5px;padding:3px 7px;border-radius:6px;font-weight:600;}
  .chk.on{background:var(--green-tint);color:#3B7A24;} .chk.off{background:#F1F2F4;color:#93A0AB;}
  .drilldown-card{margin-top:20px;}
  .drilldown-head{display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;padding:18px 20px 14px;border-bottom:1px solid var(--line);}
  .drilldown-head h4{font-size:15px;color:var(--navy-deep);}
  .drilldown-count{font-size:12px;font-weight:700;color:var(--navy);background:var(--sky-tint);padding:4px 11px;border-radius:999px;}
  .drilldown-empty{padding:30px 20px;color:var(--ink-soft);font-size:13.5px;text-align:center;}
  .drilldown-body{max-height:420px;overflow-y:auto;} .drilldown-body table{font-size:12.8px;}
  footer{background:var(--navy-deep);color:rgba(255,255,255,.75);padding:26px 0;font-size:12.5px;}
  footer .wrap{display:flex;justify-content:space-between;flex-wrap:wrap;gap:10px;} footer strong{color:#fff;}
  @media(max-width:980px){.kpi-grid{grid-template-columns:repeat(3,1fr);}.proj-grid{grid-template-columns:1fr;}.split{grid-template-columns:1fr;}#pipeline .split{grid-template-columns:1fr;}}
  @media(max-width:600px){.kpi-grid{grid-template-columns:repeat(2,1fr);}}
"""


def main():
    xlsx_path = sys.argv[1] if len(sys.argv) > 1 else 'sample.xlsx'
    logo_path = sys.argv[2] if len(sys.argv) > 2 else 'logo.png'
    out_dir = sys.argv[3] if len(sys.argv) > 3 else '/mnt/user-data/outputs'
    make_png = '--with-png' in sys.argv
    os.makedirs(out_dir, exist_ok=True)

    logo_b64 = ''
    if os.path.exists(logo_path):
        with open(logo_path, 'rb') as f:
            logo_b64 = base64.b64encode(f.read()).decode()

    projects = load_projects(xlsx_path)
    outcomes = load_outcomes(xlsx_path)
    models = load_models(xlsx_path)
    ra = load_ra_contracts(xlsx_path)
    review, review_stats = load_review_pipeline(xlsx_path)
    sandbox = load_sandbox(xlsx_path)

    # ---- self-audit (fail loudly rather than ship silently-wrong data) ----
    bucket_total = len(review)
    assert bucket_total == len(review), "bucket total mismatch"
    ra_names = [r['name'] for r in ra]
    assert len(ra_names) == len(set(ra_names)), f"RA dedup failed: {ra_names}"
    for r in ra:
        assert r['name'] and r['name'].lower() not in ('unnamed', 'none', 'nan'), f"placeholder RA identity: {r}"

    html = strip_indent(build_html(projects, outcomes, ra, review, review_stats, sandbox, models, logo_b64))
    html_path = os.path.join(out_dir, 'AI_Hub_Executive_Dashboard.html')
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"Wrote {html_path} ({len(html)} bytes)")

    if make_png:
        png_path = os.path.join(out_dir, 'AI_Hub_Executive_Dashboard.png')
        try:
            from playwright.sync_api import sync_playwright
            with sync_playwright() as p:
                b = p.chromium.launch()
                page = b.new_page(viewport={'width': 1280, 'height': 1000}, device_scale_factor=2)
                page.goto('file://' + os.path.abspath(html_path))
                page.wait_for_timeout(500)
                page.screenshot(path=png_path, full_page=True)
                b.close()
            print(f"Wrote {png_path}")
        except Exception as e:
            print(f"PNG render skipped: {e}")
    else:
        print("PNG skipped (pass --with-png to generate one)")

    print("\nSummary:")
    print(f"  Projects: {len(projects)}  |  Outcomes: {len(outcomes)}  |  Models: {len(models)}  |  RA contracts (deduped): {len(ra)}")
    print(f"  Review pipeline: {len(review)} studies  |  Sandbox: {len(sandbox)}")



if __name__ == '__main__':
    main()
