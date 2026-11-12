"""
Post Hoc Analysis System  (Calculations, Report, All 3/4 Column Questions, NFD)

Local chalane ka tareeqa:
    python -m pip install -r requirements.txt
    python -m streamlit run post_hoc_app.py      (ya editor ka Run button)

Excel file ki shakal (sheets ki tarteeb koi bhi ho sakti hai):
    Responses sheet : row 1 = Q1, Q2, Q3 ...; har row ek student, A/B/C/D/E
    Keys sheet      : ya to QNO / KEY columns, ya row 1 = Q1, Q2 ... aur row 2 = sahi jawab
"""
from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

NFD_LIMIT = 5.0  # koi option is percent se kam students ne chuna to woh Non-Functional hai
DEFAULT_HEADING = ["Ziauddin University", "Examination Department"]
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="D9E1F2")
LOW_FILL = PatternFill("solid", fgColor="F8CBAD")


# ---------------------------------------------------------------- loading
def _is_q(x):
    """Kya ye 'Q12' jaisa label hai?"""
    x = str(x).strip().upper()
    return x.startswith("Q") and x[1:].isdigit()


def load_excel(file):
    """Sheets ki tarteeb se farq nahi parta: app khud pehchanti hai ke kaun si sheet kya hai."""
    xls = pd.ExcelFile(file)
    resp, key = None, {}

    for name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=name, header=None, dtype=str).fillna("")
        df = df.apply(lambda c: c.str.strip())
        if df.empty:
            continue

        # 1) Keys neeche ki taraf: QNO / KEY header (pehli 5 rows mein kahin bhi)
        found = False
        for i in range(min(5, len(df))):
            row = [c.upper() for c in df.iloc[i]]
            if "QNO" in row and "KEY" in row:
                q_pos, k_pos = row.index("QNO"), row.index("KEY")
                for _, r in df.iloc[i + 1:].iterrows():
                    if r.iloc[q_pos] == "":
                        continue
                    num = str(int(float(r.iloc[q_pos])))
                    key["Q" + num] = r.iloc[k_pos].upper()
                found = True
                break
        if found:
            continue

        # 2) Pehli row mein Q1, Q2 ... hon to ya keys (side mein) hain ya responses
        header = list(df.iloc[0])
        q_pos = [j for j, h in enumerate(header) if _is_q(h)]
        if not q_pos:
            continue
        labels = [header[j].upper() for j in q_pos]

        if "KEY" in name.upper() or len(df) == 2:
            for j, lab in zip(q_pos, labels):
                key[lab] = df.iloc[1, j].upper()
        elif resp is None:
            data = df.iloc[1:, q_pos].copy()
            data.columns = labels
            data = data.apply(lambda c: c.str.upper())
            resp = data[(data != "").any(axis=1)].reset_index(drop=True)  # absent students nikaalo

    if resp is None:
        raise ValueError("Responses wali sheet nahi mili (pehli row mein Q1, Q2, Q3 ... hona chahiye).")
    if not key:
        raise ValueError("Keys wali sheet nahi mili.")

    missing = [q for q in resp.columns if q not in key]
    if missing:
        raise ValueError(f"In questions ki key nahi mili: {missing}")
    return resp, pd.Series(key)[resp.columns]


# ----------------------------------------------------------- calculations
def calculations(resp, key):
    """0/1 matrix, Total, Total Sum of Answers, Group, Diff. Index, Disc. Index"""
    qcols = list(resp.columns)

    marks = resp.eq(key).astype(int)                  # sahi = 1, ghalat/khali = 0
    marks["Total"] = marks[qcols].sum(axis=1)         # row-wise total
    marks = marks.sort_values("Total", ascending=False, kind="stable").reset_index(drop=True)
    marks.index = marks.index + 1                     # student no. 1 se shuru

    n = len(marks)
    group = max(1, int(n / 100 * 27 + 0.5))           # Group = Total Students/100*27 (round half up)

    total_sum = marks[qcols].sum()                    # Total Sum of Answers
    upper = marks[qcols].head(group).sum()
    lower = marks[qcols].tail(group).sum()

    summary = pd.DataFrame(
        {
            "Total Sum of Answers": total_sum,
            "Upper Group Sum": upper,
            "Lower Group Sum": lower,
            "Diff. Index": total_sum / n * 100,
            "Disc. Index": (upper - lower) / group,
        }
    ).T
    return marks, summary, n, group


# ----------------------------------------------------------------- report
def _alpha(X):
    """Cronbach's Alpha (X = students x items, 1/0)"""
    k = X.shape[1]
    total_var = X.sum(axis=1).var(ddof=1)
    if k < 2 or not total_var or np.isnan(total_var):
        return np.nan
    return k / (k - 1) * (1 - X.var(ddof=1).sum() / total_var)


REPORT_COLS = [
    "Questions", "Difficulty Index", "Disc. Index", "Cronbach's Alpha if Item Deleted",
    "Non Functional Distractor", "Flagged In", "No. of Columns",
]


def report(resp, key):
    """Report, NFD percent table, Cronbach's Alpha aur All 3 / All 4 Column lists."""
    _, summary, n, group = calculations(resp, key)
    qcols = list(resp.columns)

    # --- NFD: har option (A, B, C ...) kitne percent students ne chuna
    letters = sorted({c for c in resp.stack() if c} | set(key))
    options = [chr(o) for o in range(ord("A"), ord(max(letters)) + 1)]
    pct = pd.DataFrame(
        {o: [(resp[q] == o).sum() / n * 100 for q in qcols] for o in options}, index=qcols
    )
    nfd = (pct < NFD_LIMIT).sum(axis=1)   # key wala option bhi gina jata hai

    # --- Reliability (SPSS ki tarah sirf un students se jinhon ne koi jawab khali nahi chhora)
    marks01 = resp.eq(key).astype(int)
    full = marks01[(resp != "").all(axis=1)]
    rel_students = len(full)
    if rel_students >= 3:
        alpha_del = {q: _alpha(full.drop(columns=q)) for q in qcols}
        alpha = _alpha(full)
    else:
        alpha_del = {q: np.nan for q in qcols}
        alpha = np.nan
    alpha_ref = round(alpha, 3) if not np.isnan(alpha) else np.nan

    # --- Flags (4 columns): Diff, Disc, Alpha-if-deleted, NFD
    rows = []
    for q in qcols:
        d, s = summary.loc["Diff. Index", q], summary.loc["Disc. Index", q]
        flags = {
            "Difficulty": d < 30 or d > 70,
            "Disc.": s < 0.2,
            "Alpha": bool(alpha_del[q] > alpha_ref),
            "NFD": nfd[q] > 1,
        }
        rows.append(
            {
                "Questions": q,
                "Difficulty Index": d,
                "Disc. Index": s,
                "Cronbach's Alpha if Item Deleted": alpha_del[q],
                "Non Functional Distractor": int(nfd[q]),
                "Flagged In": ", ".join(k for k, v in flags.items() if v),
                "No. of Columns": sum(flags.values()),
            }
        )
    table = pd.DataFrame(rows)[REPORT_COLS]
    three = table.loc[table["No. of Columns"] == 3, "Questions"].tolist()
    four = table.loc[table["No. of Columns"] == 4, "Questions"].tolist()

    return {
        "table": table, "alpha": alpha, "items": len(qcols), "students": n,
        "rel_students": rel_students, "three": three, "four": four,
        "pct": pct, "nfd": nfd, "options": options, "key": key,
    }


def _list_df(questions):
    return pd.DataFrame({"Questions": questions, "Remarks": [""] * len(questions)})


def _shown(table):
    """Web page par dikhane ke liye numbers gol karo."""
    t = table.copy()
    t["Difficulty Index"] = t["Difficulty Index"].round(1)
    t["Disc. Index"] = t["Disc. Index"].round(2)
    t["Cronbach's Alpha if Item Deleted"] = t["Cronbach's Alpha if Item Deleted"].round(4)
    return t.set_index("Questions")


def _nfd_shown(rep):
    t = rep["pct"].round(1)
    t["Non Functional Distractor"] = rep["nfd"]
    t.index.name = "Questions"
    return t


# ----------------------------------------------------------------- excel
def calculations_excel(marks, summary):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.concat([marks, summary]).to_excel(w, sheet_name="Calculations")
    return buf.getvalue()


def _heading_size(i):
    """Default 2 lines (Ziauddin University / Examination Department) = 24, user ki lines = 16."""
    return 24 if i <= len(DEFAULT_HEADING) else 16


def _heading(ws, lines, ncols):
    """Heading ki lines: har line poori table ki width par Merge & Center. Table ki pehli row wapis karta hai."""
    for i, text in enumerate(lines, start=1):
        if ncols > 1:
            ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=ncols)
        c = ws.cell(i, 1, text)
        c.font = Font(bold=True, size=_heading_size(i))
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    return len(lines) + 2  # heading ke baad ek khali row


def _fit_heading(ws, lines, ncols):
    """Columns ki width dekh kar heading ki rows ki height set karo (taake bade font ka text kate nahi)."""
    width = sum(ws.column_dimensions[get_column_letter(j)].width or 8.43 for j in range(1, ncols + 1))
    for i, text in enumerate(lines, start=1):
        size = _heading_size(i)
        per_line = max(1, int(width / (size / 11 * 1.15)))   # ek line mein kitne akshar aate hain
        n_lines = -(-len(text) // per_line)
        ws.row_dimensions[i].height = n_lines * size * 1.4


def _clean(v):
    if isinstance(v, np.integer):
        return int(v)
    if isinstance(v, (float, np.floating)):
        return None if np.isnan(v) else float(v)
    return v


def _table(ws, df, top, formats=None):
    """DataFrame ko sheet mein likho (header, borders, widths). Aakhri row ka number wapis karta hai."""
    formats = formats or {}
    for j, name in enumerate(df.columns, start=1):
        c = ws.cell(top, j, name)
        c.font = Font(bold=True)
        c.fill = HEAD_FILL
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, rec in enumerate(df.itertuples(index=False), start=top + 1):
        for j, v in enumerate(rec, start=1):
            c = ws.cell(i, j, _clean(v))
            c.border = BORDER
            c.alignment = Alignment(horizontal="center")
            if df.columns[j - 1] in formats:
                c.number_format = formats[df.columns[j - 1]]
    for j, name in enumerate(df.columns, start=1):
        longest = max([len(str(name))] + [len(str(x)) for x in df.iloc[:, j - 1]])
        ws.column_dimensions[get_column_letter(j)].width = min(max(longest + 2, 12), 36)
    return top + len(df)


def report_excel(rep, lines):
    wb = Workbook()
    ws = wb.active
    ws.title = "Report"
    ncols = len(REPORT_COLS)
    top = _heading(ws, lines, ncols)
    last = _table(
        ws, rep["table"], top,
        {"Difficulty Index": "0.0", "Disc. Index": "0.00", "Cronbach's Alpha if Item Deleted": "0.0000"},
    )
    # Table ke neeche: Cronbach's Alpha aur items ki ginti
    a = "-" if np.isnan(rep["alpha"]) else f"{rep['alpha']:.3f}"
    r = last + 2
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
    c = ws.cell(r, 1, f"Cronbach's Alpha: {a}  |  No. of Items: {rep['items']}")
    c.font = Font(bold=True)
    c.alignment = Alignment(horizontal="center", vertical="center")
    _fit_heading(ws, lines, ncols)
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def columns_excel(rep, lines):
    """All 3 Column aur All 4 Column lists, alag Excel file mein (har list ki apni sheet)."""
    wb = Workbook()
    first = True
    for title, qs in (("All 3 Column Questions", rep["three"]), ("All 4 Column Questions", rep["four"])):
        ws = wb.active if first else wb.create_sheet()
        first = False
        ws.title = title
        top = _heading(ws, lines, 2)
        _table(ws, _list_df(qs), top)
        ws.column_dimensions["A"].width = 20
        ws.column_dimensions["B"].width = 44
        _fit_heading(ws, lines, 2)
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def nfd_excel(rep, lines):
    """Har question ke samne A, B, C ... options kitne percent students ne chune."""
    pct, options, key = rep["pct"], rep["options"], rep["key"]
    qcols = list(pct.index)
    df = (pct / 100).copy()
    df.insert(0, "Questions", qcols)
    df["Non Functional Distractor"] = rep["nfd"].values
    ncols = len(df.columns)

    wb = Workbook()
    ws = wb.active
    ws.title = "NFD"
    top = _heading(ws, lines, ncols)
    last = _table(ws, df, top, {o: "0.0%" for o in options})

    for i, q in enumerate(qcols):
        for j, o in enumerate(options, start=2):
            cell = ws.cell(top + 1 + i, j)
            if pct.loc[q, o] < NFD_LIMIT:
                cell.fill = LOW_FILL          # 5% se kam
            if key[q] == o:
                cell.font = Font(bold=True)   # key wala option

    r = last + 2
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
    note = ws.cell(r, 1, f"Shaded cells = options chosen by less than {NFD_LIMIT:g}% of students (Non Functional).  "
                         "Bold = key (correct) option.")
    note.font = Font(italic=True, size=10)
    note.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 30

    _fit_heading(ws, lines, ncols)
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ------------------------------------------------------------------- UI
def download_button(label, data, file_name):
    """Download button jo click par app ko dobara run nahi karta (isse 'File wasn't available' error aata tha)."""
    try:
        version = tuple(int(x) for x in st.__version__.split(".")[:2])
    except Exception:
        version = (0, 0)
    if version >= (1, 43):
        st.download_button(label, data, file_name=file_name, mime=XLSX_MIME, on_click="ignore")
    else:
        st.download_button(label, data, file_name=file_name, mime=XLSX_MIME)


def main():
    st.set_page_config(page_title="Post Hoc Analysis", layout="wide")
    st.title("Post Hoc Analysis System")

    file = st.file_uploader("Excel file upload karein (Responses sheet + Keys sheet)", type=["xlsx"])
    if file is None:
        st.stop()

    try:
        resp, key = load_excel(file)
    except Exception as e:
        st.error(f"File parh nahi saka: {e}")
        st.stop()

    st.success(f"{len(resp)} students aur {len(key)} questions mile.")

    # --- Excel files ki heading: 2 default lines + user ki 3 lines tak
    st.subheader("Excel files ki heading")
    st.caption("Har Excel file ke upar 'Ziauddin University' aur 'Examination Department' khud aayega. "
               "Uske neeche ki lines yahan likhein (zyada se zyada 3). Likh kar Enter dabayein.")
    h1 = st.text_input("Line 1", placeholder="Bachelors of Dental Surgery", max_chars=80)
    h2 = st.text_input("Line 2", placeholder="Batch-15 3rd Prof", max_chars=80)
    h3 = st.text_input("Line 3", placeholder="Oral Medicine", max_chars=80)
    lines = DEFAULT_HEADING + [x.strip() for x in (h1, h2, h3) if x.strip()]

    st.divider()
    c1, c2, c3, c4, c5 = st.columns(5)
    if c1.button("Calculations", use_container_width=True):
        st.session_state.view = "calc"
    if c2.button("Report", use_container_width=True):
        st.session_state.view = "report"
    if c3.button("All 3 Column Questions", use_container_width=True):
        st.session_state.view = "three"
    if c4.button("All 4 Column Questions", use_container_width=True):
        st.session_state.view = "four"
    if c5.button("NFD", use_container_width=True):
        st.session_state.view = "nfd"

    view = st.session_state.get("view")

    if view == "calc":
        marks, summary, n, group = calculations(resp, key)
        st.write(f"**Total Students:** {n}   |   **Group:** {group}")
        st.subheader("Calculations")
        st.dataframe(marks, use_container_width=True)
        st.subheader("Total Sum of Answers, Diff. Index, Disc. Index")
        st.dataframe(summary, use_container_width=True)
        download_button("Calculations Excel mein download karein",
                        calculations_excel(marks, summary), "Calculations.xlsx")

    elif view == "report":
        rep = report(resp, key)
        st.subheader("Report")
        st.dataframe(_shown(rep["table"]), use_container_width=True, height=600)
        a = "-" if np.isnan(rep["alpha"]) else f"{rep['alpha']:.3f}"
        st.write(f"**Cronbach's Alpha:** {a}   |   **No of Items:** {rep['items']}")
        st.caption(
            f"Alpha un {rep['rel_students']} students se nikla hai jinhon ne koi jawab khali nahi chhora (SPSS ki tarah). "
            "Columns: Difficulty (30 se kam / 70 se zyada), Disc. Index (0.2 se kam), "
            "Alpha if Item Deleted (Cronbach's Alpha se zyada), Non Functional Distractor (1 se zyada)."
        )
        download_button("Report Excel mein download karein", report_excel(rep, lines), "Report.xlsx")

    elif view in ("three", "four"):
        rep = report(resp, key)
        title, qs = ("All 3 Column Questions", rep["three"]) if view == "three" else ("All 4 Column Questions", rep["four"])
        st.subheader(f"{title} ({len(qs)})")
        st.dataframe(_list_df(qs), use_container_width=True, hide_index=True)
        st.caption("Download wali Excel file mein 3 Column aur 4 Column dono lists (alag alag sheet mein) aati hain.")
        download_button("3 aur 4 Column lists Excel mein download karein",
                        columns_excel(rep, lines), "All_3_and_4_Column_Questions.xlsx")

    elif view == "nfd":
        rep = report(resp, key)
        st.subheader("Non Functional Distractor")
        st.caption(f"Har option kitne percent students ne chuna. {NFD_LIMIT:g}% se kam wale option Non Functional ginte hain "
                   "(key wala option bhi).")
        st.dataframe(_nfd_shown(rep), use_container_width=True, height=600)
        download_button("NFD Excel mein download karein", nfd_excel(rep, lines), "NFD.xlsx")


if __name__ == "__main__":
    from streamlit.runtime.scriptrunner import get_script_run_ctx

    if get_script_run_ctx() is None:
        # Editor ke Run button se chalaya gaya hai: khud streamlit se dobara chalao
        import subprocess
        import sys

        subprocess.run([sys.executable, "-m", "streamlit", "run", __file__])
    else:
        main()
