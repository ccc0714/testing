"""
Pigeon PRN Converter — Streamlit Web App
Run with: streamlit run app.py

Requires: pip install streamlit openpyxl
Place this file in the same folder as prn_to_excel.py (your main script).
"""

import streamlit as st
import tempfile
import os
import io
import openpyxl
from prn_to_excel import prn_to_excel, parse_prn

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Pigeon PRN Converter",
    page_icon="🐦",
    layout="centered",
)

# ── Styling ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');

    html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }
    .main { background-color: #f7f7f5; }
    h1, h2, h3 { font-family: 'IBM Plex Mono', monospace; }

    .stButton > button {
        background-color: #1a1a1a;
        color: #f7f7f5;
        border: none;
        border-radius: 4px;
        padding: 0.6rem 2rem;
        font-family: 'IBM Plex Mono', monospace;
        font-size: 0.9rem;
        letter-spacing: 0.05em;
        transition: background 0.2s;
    }
    .stButton > button:hover { background-color: #333; }

    .file-card {
        background: white;
        border: 1px solid #e0e0e0;
        border-radius: 6px;
        padding: 0.75rem 1rem;
        margin-bottom: 0.5rem;
        display: flex;
        justify-content: space-between;
        align-items: center;
        font-family: 'IBM Plex Mono', monospace;
        font-size: 0.82rem;
    }
    .file-card .status-ok  { color: #2d7a2d; font-weight: 600; }
    .file-card .status-err { color: #c0392b; font-weight: 600; }

    .summary-box {
        background: #1a1a1a;
        color: #f7f7f5;
        border-radius: 6px;
        padding: 1rem 1.25rem;
        font-family: 'IBM Plex Mono', monospace;
        font-size: 0.8rem;
        line-height: 1.7;
        margin-top: 1rem;
    }
    .summary-box .accent { color: #a8d8a8; }
    .summary-box .dim    { color: #888; }

    .pill {
        display: inline-block;
        background: #e8f4e8;
        color: #2d7a2d;
        border-radius: 12px;
        padding: 0.15rem 0.6rem;
        font-size: 0.75rem;
        font-family: 'IBM Plex Mono', monospace;
        margin-left: 0.4rem;
    }
    .pill-warn { background: #fff4e0; color: #b36b00; }
    .pill-err  { background: #fde8e8; color: #c0392b; }
    hr { border: none; border-top: 1px solid #e0e0e0; margin: 1.5rem 0; }
</style>
""", unsafe_allow_html=True)


# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("# 🐦 Pigeon PRN Converter")
st.markdown(
    "Upload one or more MED-PC `.PRN` files. "
    "All files will be combined into a single Excel workbook, one sheet per file."
)
st.markdown("<hr>", unsafe_allow_html=True)


# ── Mode toggle ───────────────────────────────────────────────────────────────
mode = st.radio(
    "What would you like to do?",
    options=["Create a new workbook", "Append to an existing workbook"],
    horizontal=True,
    help="Choose 'Append' to add new sheets into an Excel file you already have.",
)
append_mode = (mode == "Append to an existing workbook")
st.markdown("<br>", unsafe_allow_html=True)


# ── Existing workbook uploader (append mode only) ─────────────────────────────
existing_xlsx_file = None
existing_sheet_names = []

if append_mode:
    st.markdown("**Step 1 — Upload your existing workbook**")
    existing_xlsx_file = st.file_uploader(
        "Upload the Excel workbook to append to",
        type=["xlsx"],
        key="existing_wb",
        help="New sheets will be added to this file. Your existing sheets are kept.",
    )

    if existing_xlsx_file:
        try:
            wb_check = openpyxl.load_workbook(
                io.BytesIO(existing_xlsx_file.read()), read_only=True
            )
            existing_sheet_names = wb_check.sheetnames
            existing_xlsx_file.seek(0)

            st.markdown(
                f'<div class="file-card">'
                f'<span>📊 <b>{existing_xlsx_file.name}</b>'
                f'<span class="pill">{len(existing_sheet_names)} existing sheet(s)</span>'
                f'</span>'
                f'<span class="status-ok">✓ loaded</span>'
                f'</div>',
                unsafe_allow_html=True,
            )
            with st.expander("View existing sheets"):
                for s in existing_sheet_names:
                    st.markdown(f"- `{s}`")

        except Exception as e:
            st.error(f"Could not read workbook: {e}")
            existing_xlsx_file = None

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("**Step 2 — Upload PRN files to add**")
else:
    st.markdown("**Upload your PRN files**")


# ── PRN file uploader ─────────────────────────────────────────────────────────
uploaded_files = st.file_uploader(
    "Drop PRN files here or click to browse",
    type=["PRN", "prn"],
    accept_multiple_files=True,
    key="prn_files",
    help="You can select multiple files at once using Shift+Click or Ctrl+Click.",
)


# ── Options ───────────────────────────────────────────────────────────────────
with st.expander("⚙️ Options"):
    if append_mode and existing_xlsx_file:
        output_filename = existing_xlsx_file.name
        st.markdown(f"Output filename: **{output_filename}** *(matches uploaded workbook)*")
    else:
        output_filename = st.text_input(
            "Output filename",
            value="output.xlsx",
            help="Name of the Excel file you will download.",
        )
        if not output_filename.endswith(".xlsx"):
            output_filename += ".xlsx"

    use_subject_as_sheet = st.checkbox(
        "Use Subject ID as sheet name",
        value=True,
        help="When checked, each sheet is named after the Subject field in the PRN header. "
             "Uncheck to use the filename instead.",
    )

st.markdown("<hr>", unsafe_allow_html=True)


# ── Preview uploaded PRN files ────────────────────────────────────────────────
if uploaded_files:
    st.markdown(f"### {len(uploaded_files)} PRN file(s) ready")

    previews = []
    for uf in uploaded_files:
        try:
            content = uf.read().decode("utf-8", errors="replace")
            uf.seek(0)

            lines = content.splitlines()
            meta = {"subject": "?", "date": "?", "box": "?"}
            for line in lines[:20]:
                s = line.strip()
                if s.startswith("Subject:"):
                    meta["subject"] = s[8:].strip()
                elif s.startswith("Start Date:"):
                    meta["date"] = s[11:].strip()
                elif s.startswith("Box:"):
                    meta["box"] = s[4:].strip()

            sheet_would_be = (
                meta["subject"] if use_subject_as_sheet
                else os.path.splitext(uf.name)[0]
            )
            already_exists = sheet_would_be in existing_sheet_names
            previews.append((uf.name, meta, True, "", already_exists, sheet_would_be))

        except Exception as e:
            previews.append((uf.name, {}, False, str(e), False, ""))

    for fname, meta, ok, err, already_exists, sheet_would_be in previews:
        if ok:
            warn = (
                f'<span class="pill pill-warn">⚠ "{sheet_would_be}" exists — will rename</span>'
                if already_exists else ""
            )
            st.markdown(
                f'<div class="file-card">'
                f'<span>📄 <b>{fname}</b>'
                f'<span class="pill">Subject {meta["subject"]}</span>'
                f'<span class="pill">Box {meta["box"]}</span>'
                f'<span class="pill">{meta["date"]}</span>'
                f'{warn}</span>'
                f'<span class="status-ok">✓ ready</span>'
                f'</div>',
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                f'<div class="file-card">'
                f'<span>📄 <b>{fname}</b></span>'
                f'<span class="status-err">✗ {err}</span>'
                f'</div>',
                unsafe_allow_html=True,
            )

    st.markdown("<br>", unsafe_allow_html=True)

    # Guard: append mode requires an existing workbook first
    if append_mode and not existing_xlsx_file:
        st.warning("Please upload an existing workbook above before converting.")
    else:
        if st.button("Convert to Excel"):
            results = []

            with tempfile.TemporaryDirectory() as tmp:
                output_filename_local = (
                    existing_xlsx_file.name if (append_mode and existing_xlsx_file)
                    else output_filename
                )
                xlsx_path = os.path.join(tmp, output_filename_local)

                # Seed the temp file with the existing workbook if appending
                if append_mode and existing_xlsx_file:
                    existing_xlsx_file.seek(0)
                    with open(xlsx_path, "wb") as f:
                        f.write(existing_xlsx_file.read())

                progress = st.progress(0, text="Converting...")

                for i, uf in enumerate(uploaded_files):
                    uf.seek(0)
                    prn_path = os.path.join(tmp, uf.name)
                    with open(prn_path, "wb") as f:
                        f.write(uf.read())

                    try:
                        if use_subject_as_sheet:
                            parsed = parse_prn(prn_path)
                            sheet_name = parsed.get("subject") or os.path.splitext(uf.name)[0]
                        else:
                            sheet_name = os.path.splitext(uf.name)[0]

                        prn_to_excel(prn_path, xlsx_path, sheet_name)
                        results.append((uf.name, sheet_name, True, ""))

                    except Exception as e:
                        results.append((uf.name, "", False, str(e)))

                    progress.progress(
                        (i + 1) / len(uploaded_files),
                        text=f"Converting {uf.name}…",
                    )

                progress.empty()

                succeeded = [r for r in results if r[2]]
                failed    = [r for r in results if not r[2]]

                if succeeded:
                    lines_html = ""
                    for fname, sheet, ok, err in results:
                        if ok:
                            action = "appended as" if append_mode else "added as"
                            lines_html += (
                                f'<span class="accent">✓</span> {fname} '
                                f'<span class="dim">→ {action} sheet "{sheet}"</span><br>'
                            )
                        else:
                            lines_html += (
                                f'<span style="color:#e87070">✗</span> {fname} '
                                f'<span class="dim">— {err}</span><br>'
                            )

                    total_sheets = (
                        len(existing_sheet_names) + len(succeeded)
                        if append_mode else len(succeeded)
                    )
                    extra = (
                        f' <span class="dim">· workbook now has {total_sheets} sheet(s)</span>'
                        if append_mode else ""
                    )
                    st.markdown(
                        f'<div class="summary-box">'
                        f'<b>{len(succeeded)}/{len(results)} files converted</b>{extra}'
                        f'<br><br>{lines_html}'
                        f'</div>',
                        unsafe_allow_html=True,
                    )

                    with open(xlsx_path, "rb") as f:
                        xlsx_bytes = f.read()

                    st.download_button(
                        label=f"⬇️  Download {output_filename_local}",
                        data=xlsx_bytes,
                        file_name=output_filename_local,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                if failed and not succeeded:
                    st.error("All files failed to convert. Check that they are valid PRN files.")

else:
    st.markdown(
        '<div style="text-align:center; color:#aaa; padding: 2rem 0; '
        'font-family: IBM Plex Mono, monospace; font-size: 0.85rem;">'
        'No PRN files uploaded yet.'
        '</div>',
        unsafe_allow_html=True,
    )
