from __future__ import annotations

import io
import re
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Proteasome Replicate QC", page_icon="🔬", layout="wide")
st.title("🔬 Proteasome Technical-Replicate QC")
st.caption("Compare two analysis workbooks and append only Recurring candidates and Non-recurring candidates to a copy of replicate 1.")

QC_OPTIONS = [
    "Not reviewed",
    "Keep / confident",
    "Interesting / investigate",
    "Low intensity",
    "Noise",
    "Ambiguous",
    "Exclude",
]


def parse_peak(value):
    m = re.search(r"Peak\s*(\d+)", str(value), re.I)
    return f"Peak {int(m.group(1))}" if m else ""


def infer_meta(filename, report=""):
    text = f"{filename} {report}".lower()
    age = ""
    if any(x in text for x in ["old", "adult", "aged"]):
        age = "Old/Adult"
    elif any(x in text for x in ["infant", "young", "juvenile", "pup"]):
        age = "Infant/Young"

    organ = ""
    for key, label in {
        "liver": "Liver", "brain": "Brain", "lung": "Lung",
        "kidney": "Kidney", "kideny": "Kidney", "heart": "Heart", "spleen": "Spleen",
    }.items():
        if key in text:
            organ = label
            break
    return age, organ


def parse_combined(uploaded):
    raw = uploaded.getvalue()
    xls = pd.ExcelFile(io.BytesIO(raw), engine="openpyxl")
    if "Combined results" not in xls.sheet_names:
        raise ValueError('Workbook has no "Combined results" sheet.')

    grid = pd.read_excel(io.BytesIO(raw), sheet_name="Combined results", header=None, engine="openpyxl")
    rows = []
    report = ""
    datafile = ""
    peak = ""
    headers = None

    for i in range(len(grid)):
        vals = grid.iloc[i].tolist()
        first = str(vals[0]).strip() if pd.notna(vals[0]) else ""

        if "|  Peak" in first or re.search(r"\.docx.*Peak\s*\d+", first, re.I):
            report = first.split("|")[0].strip()
            peak = parse_peak(first)
            headers = None
            continue
        if first == "Report file":
            report = str(vals[1]) if len(vals) > 1 and pd.notna(vals[1]) else report
            peak = peak or parse_peak(report)
            continue
        if first == "Data file":
            datafile = str(vals[1]) if len(vals) > 1 and pd.notna(vals[1]) else ""
            continue
        if first == "Subunit" and "Detected_Mass_Da" in vals:
            headers = [str(v) if pd.notna(v) else "" for v in vals]
            continue

        if headers and "Detected_Mass_Da" in headers:
            d = dict(zip(headers, vals))
            mass = d.get("Detected_Mass_Da")
            if pd.notna(mass) and isinstance(mass, (int, float, np.number)):
                age, organ = infer_meta(uploaded.name, report)
                d.update({
                    "Source_Workbook": uploaded.name,
                    "Report_File": report,
                    "Data_File_Name": datafile,
                    "Peak": peak or parse_peak(report),
                    "Age_Group": age,
                    "Organ": organ,
                })
                rows.append(d)

    if not rows:
        raise ValueError("No detected masses were found in Combined results.")

    out = pd.DataFrame(rows)
    for col in [
        "Detected_Mass_Da", "Calculated_Mass_Da", "Difference_Da", "Difference_ppm",
        "Charge_State_Count", "Sum_Intensity", "Relative_Intensity_pct",
    ]:
        if col in out.columns:
            out[col] = pd.to_numeric(out[col], errors="coerce")
    return out


def best_assignment(row):
    sub = row.get("Subunit", "")
    mod = row.get("Modification", "")
    cat = row.get("Category", "")
    expl = row.get("Possible_Explanation", "")
    parts = [str(x).strip() for x in [sub, mod] if pd.notna(x) and str(x).strip()]
    assignment = " — ".join(parts)
    if not assignment and pd.notna(expl) and str(expl).strip():
        assignment = str(expl).strip()
    return assignment, cat


def greedy_match(a, b, tolerance, same_peak=False):
    """One-to-one nearest matching within tolerance."""
    A = a.reset_index(drop=True)
    B = b.reset_index(drop=True)
    candidates = []
    for ia, ra in A.iterrows():
        for ib, rb in B.iterrows():
            if same_peak and str(ra.get("Peak", "")) != str(rb.get("Peak", "")):
                continue
            diff = abs(float(ra["Detected_Mass_Da"]) - float(rb["Detected_Mass_Da"]))
            if diff <= tolerance:
                candidates.append((diff, ia, ib))
    candidates.sort(key=lambda x: x[0])

    used_a, used_b, pairs = set(), set(), []
    for diff, ia, ib in candidates:
        if ia not in used_a and ib not in used_b:
            used_a.add(ia)
            used_b.add(ib)
            pairs.append((ia, ib, diff))

    recurring = []
    for ia, ib, diff in pairs:
        ra, rb = A.loc[ia], B.loc[ib]
        a_assign, a_cat = best_assignment(ra)
        b_assign, b_cat = best_assignment(rb)
        recurring.append({
            "Consensus_Mass_Da": np.mean([ra["Detected_Mass_Da"], rb["Detected_Mass_Da"]]),
            "Rep1_Mass_Da": ra["Detected_Mass_Da"],
            "Rep2_Mass_Da": rb["Detected_Mass_Da"],
            "Replicate_Difference_Da": diff,
            "Rep1_App_Category": a_cat,
            "Rep2_App_Category": b_cat,
            "Rep1_Assignment": a_assign,
            "Rep2_Assignment": b_assign,
            "Rep1_Peak": ra.get("Peak", ""),
            "Rep2_Peak": rb.get("Peak", ""),
            "Rep1_Sum_Intensity": ra.get("Sum_Intensity", np.nan),
            "Rep2_Sum_Intensity": rb.get("Sum_Intensity", np.nan),
            "Rep1_Relative_Intensity_pct": ra.get("Relative_Intensity_pct", np.nan),
            "Rep2_Relative_Intensity_pct": rb.get("Relative_Intensity_pct", np.nan),
            "Rep1_Charge_State_Range": ra.get("Charge_State_Range", ""),
            "Rep2_Charge_State_Range": rb.get("Charge_State_Range", ""),
            "Rep1_Charge_State_Count": ra.get("Charge_State_Count", np.nan),
            "Rep2_Charge_State_Count": rb.get("Charge_State_Count", np.nan),
            "Age_Group": ra.get("Age_Group", "") or rb.get("Age_Group", ""),
            "Organ": ra.get("Organ", "") or rb.get("Organ", ""),
            "Rep1_Report_File": ra.get("Report_File", ""),
            "Rep2_Report_File": rb.get("Report_File", ""),
            "Manual_QC": "Not reviewed",
            "Reviewer_Notes": "",
        })

    nonrecurring = []
    for label, frame, used in [("Replicate 1 only", A, used_a), ("Replicate 2 only", B, used_b)]:
        for idx, row in frame.iterrows():
            if idx in used:
                continue
            assignment, category = best_assignment(row)
            nonrecurring.append({
                "Detected_Mass_Da": row["Detected_Mass_Da"],
                "Found_In": label,
                "App_Category": category,
                "Assignment": assignment,
                "Peak": row.get("Peak", ""),
                "Sum_Intensity": row.get("Sum_Intensity", np.nan),
                "Relative_Intensity_pct": row.get("Relative_Intensity_pct", np.nan),
                "Charge_State_Range": row.get("Charge_State_Range", ""),
                "Charge_State_Count": row.get("Charge_State_Count", np.nan),
                "Age_Group": row.get("Age_Group", ""),
                "Organ": row.get("Organ", ""),
                "Report_File": row.get("Report_File", ""),
                "Possible_Explanation": row.get("Possible_Explanation", ""),
                "Manual_QC": "Not reviewed",
                "Reviewer_Notes": "",
            })

    return pd.DataFrame(recurring), pd.DataFrame(nonrecurring)


def style_sheet(ws):
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    for col in range(1, ws.max_column + 1):
        letter = get_column_letter(col)
        vals = [
            len(str(ws.cell(r, col).value))
            for r in range(1, min(ws.max_row, 400) + 1)
            if ws.cell(r, col).value is not None
        ]
        ws.column_dimensions[letter].width = min(max(vals + [10]) + 2, 38)


def append_dataframe_sheet(wb, name, df):
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    for c_idx, col in enumerate(df.columns, 1):
        ws.cell(1, c_idx, col)
    for r_idx, values in enumerate(df.itertuples(index=False, name=None), 2):
        for c_idx, value in enumerate(values, 1):
            if pd.isna(value):
                value = None
            ws.cell(r_idx, c_idx, value)
    style_sheet(ws)


def export_into_rep1(rep1_bytes, recurring, nonrecurring):
    """Preserve replicate-1 workbook and append/replace only two QC sheets."""
    wb = load_workbook(io.BytesIO(rep1_bytes))
    append_dataframe_sheet(wb, "Recurring candidates", recurring)
    append_dataframe_sheet(wb, "Non-recurring candidates", nonrecurring)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


st.sidebar.header("Comparison settings")
tolerance = st.sidebar.number_input("Mass tolerance (± Da)", min_value=0.05, value=1.0, step=0.05)
same_peak = st.sidebar.checkbox("Require same LC peak", value=False)
st.sidebar.caption("Default: masses within ±1 Da are treated as recurring. Replicate-only does not automatically mean noise.")

st.subheader("1. Upload the two technical-replicate analysis workbooks")
rep1_file = st.file_uploader("Technical replicate 1", type=["xlsx"], key="rep1")
rep2_file = st.file_uploader("Technical replicate 2", type=["xlsx"], key="rep2")
if rep1_file is None or rep2_file is None:
    st.info("Upload both original analysis workbooks to begin.")
    st.stop()

try:
    rep1 = parse_combined(rep1_file)
    rep2 = parse_combined(rep2_file)
except Exception as exc:
    st.error(str(exc))
    st.stop()

recurring, nonrecurring = greedy_match(rep1, rep2, tolerance, same_peak)

st.subheader("2. Recurring candidates")
st.caption("Detected in both technical replicates within the selected mass tolerance. Review the RAW data and set Manual QC if needed.")
if recurring.empty:
    recurring_edit = recurring.copy()
    st.info("No recurring masses were found with the current tolerance.")
else:
    recurring_edit = st.data_editor(
        recurring,
        hide_index=True,
        use_container_width=True,
        num_rows="fixed",
        key="recurring_editor",
        column_config={
            "Manual_QC": st.column_config.SelectboxColumn("Manual QC", options=QC_OPTIONS, required=True),
            "Reviewer_Notes": st.column_config.TextColumn("Reviewer notes", width="large"),
        },
    )

st.subheader("3. Non-recurring candidates")
st.caption("Detected in only one technical replicate after applying the selected tolerance. These are review candidates, not automatically noise.")
if nonrecurring.empty:
    nonrecurring_edit = nonrecurring.copy()
    st.info("Every detected mass found a partner in the other replicate.")
else:
    nonrecurring_edit = st.data_editor(
        nonrecurring,
        hide_index=True,
        use_container_width=True,
        num_rows="fixed",
        key="nonrecurring_editor",
        column_config={
            "Manual_QC": st.column_config.SelectboxColumn("Manual QC", options=QC_OPTIONS, required=True),
            "Reviewer_Notes": st.column_config.TextColumn("Reviewer notes", width="large"),
        },
    )

st.subheader("4. Summary")
a, b, c = st.columns(3)
a.metric("Replicate 1 masses", len(rep1))
b.metric("Replicate 2 masses", len(rep2))
c.metric("Recurring pairs", len(recurring_edit))
st.caption(f"Non-recurring observations: {len(nonrecurring_edit)}")

st.subheader("5. Download")
st.write("The download is a copy of **technical replicate 1** with only two QC sheets appended: **Recurring candidates** and **Non-recurring candidates**. Its original analysis sheets are preserved.")
output = export_into_rep1(rep1_file.getvalue(), recurring_edit, nonrecurring_edit)
base = re.sub(r"\.xlsx$", "", rep1_file.name, flags=re.I)
st.download_button(
    "Download replicate 1 + QC sheets",
    data=output,
    file_name=f"{base}_with_replicate_QC.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
