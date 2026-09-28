    nonrecurring = all_df[all_df["Detected_In_n"] < min_reps].copy()
    return recurring, nonrecurring


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
        vals = [len(str(ws.cell(r, col).value)) for r in range(1, min(ws.max_row, 400) + 1) if ws.cell(r, col).value is not None]
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
    wb = load_workbook(io.BytesIO(rep1_bytes))
    append_dataframe_sheet(wb, "Recurring candidates", recurring)
    append_dataframe_sheet(wb, "Non-recurring candidates", nonrecurring)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


st.sidebar.header("Comparison settings")
tolerance = st.sidebar.number_input("Mass tolerance (± Da)", min_value=0.05, value=1.0, step=0.05)
same_peak = st.sidebar.checkbox("Require same LC peak", value=False)
st.sidebar.caption("Masses are grouped across replicates using the selected tolerance. Replicate-only does not automatically mean noise.")

st.subheader("1. Upload 2 or 3 technical-replicate analysis workbooks")
files = st.file_uploader(
    "Select replicate Excel files in order (Rep1, Rep2, Rep3)",
    type=["xlsx"], accept_multiple_files=True,
)
if len(files) < 2:
    st.info("Upload at least 2 technical replicates. You can upload up to 3.")
    st.stop()
if len(files) > 3:
    st.error("Please upload a maximum of 3 technical replicates in this version.")
    st.stop()

min_default = 2 if len(files) >= 2 else 1
min_reps = st.sidebar.number_input(
    "Minimum replicates required to call recurring",
    min_value=2, max_value=len(files), value=min_default, step=1,
)

try:
    frames = [parse_combined(f, i + 1) for i, f in enumerate(files)]
except Exception as exc:
    st.error(str(exc))
    st.stop()

recurring, nonrecurring = build_tables(frames, tolerance, same_peak, int(min_reps))

st.subheader("2. Recurring candidates")
st.caption(f"Detected in at least {int(min_reps)}/{len(files)} technical replicates within ±{tolerance:g} Da. Mass SD is calculated only from your replicate measurements.")
if recurring.empty:
    recurring_edit = recurring.copy()
    st.info("No recurring masses were found with the current settings.")
else:
    recurring_edit = st.data_editor(
        recurring, hide_index=True, use_container_width=True, num_rows="fixed", key="recurring_editor",
        column_config={
            "Manual_QC": st.column_config.SelectboxColumn("Manual QC", options=QC_OPTIONS, required=True),
            "Reviewer_Notes": st.column_config.TextColumn("Reviewer notes", width="large"),
        },
    )

st.subheader("3. Non-recurring candidates")
st.caption(f"Detected in fewer than {int(min_reps)} of {len(files)} replicates. These are review candidates, not automatically noise.")
if nonrecurring.empty:
    nonrecurring_edit = nonrecurring.copy()
    st.info("All mass clusters satisfy the current recurrence requirement.")
else:
    nonrecurring_edit = st.data_editor(
        nonrecurring, hide_index=True, use_container_width=True, num_rows="fixed", key="nonrecurring_editor",
        column_config={
            "Manual_QC": st.column_config.SelectboxColumn("Manual QC", options=QC_OPTIONS, required=True),
            "Reviewer_Notes": st.column_config.TextColumn("Reviewer notes", width="large"),
        },
    )

st.subheader("4. Summary")
cols = st.columns(len(frames) + 2)
for i, frame in enumerate(frames):
    cols[i].metric(f"Replicate {i+1} masses", len(frame))
cols[-2].metric("Recurring clusters", len(recurring_edit))
cols[-1].metric("Non-recurring clusters", len(nonrecurring_edit))

st.subheader("5. Download")
st.write("The download is a copy of **replicate 1** with only two QC sheets appended/replaced: **Recurring candidates** and **Non-recurring candidates**. All original analysis sheets are preserved.")
output = export_into_rep1(files[0].getvalue(), recurring_edit, nonrecurring_edit)
base = re.sub(r"\.xlsx$", "", files[0].name, flags=re.I)
st.download_button(
    "Download replicate 1 + QC sheets", data=output,
    file_name=f"{base}_with_replicate_QC.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
