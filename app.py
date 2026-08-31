
import io
from pathlib import Path

import pandas as pd
import streamlit as st

from mor_parser import (
    combine_same_named_datasets,
    detect_file,
    unpack_upload,
)


APP_TITLE = "MOR Column Extractor V2"


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="mm/dd/yyyy") as writer:
        df.to_excel(writer, sheet_name="Extracted Data", index=False)

        workbook = writer.book
        worksheet = writer.sheets["Extracted Data"]

        header = workbook.add_format(
            {
                "bold": True,
                "font_color": "white",
                "bg_color": "#1F4E78",
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
                "border": 1,
            }
        )
        date_fmt = workbook.add_format({"num_format": "mm/dd/yyyy"})

        for col_idx, col in enumerate(df.columns):
            worksheet.write(0, col_idx, col, header)

            sample = [str(x) for x in df[col].head(80).fillna("").tolist()]
            width = min(max([len(str(col))] + [len(x) for x in sample]) + 2, 38)
            worksheet.set_column(col_idx, col_idx, max(width, 12))

            if pd.api.types.is_datetime64_any_dtype(df[col]):
                worksheet.set_column(col_idx, col_idx, max(width, 12), date_fmt)

        worksheet.set_row(0, 38)
        worksheet.freeze_panes(1, 0)

        if len(df) and len(df.columns):
            worksheet.add_table(
                0,
                0,
                len(df),
                len(df.columns) - 1,
                {
                    "name": "MORExtractedData",
                    "columns": [{"header": c} for c in df.columns],
                    "style": "Table Style Medium 2",
                },
            )

    return output.getvalue()


def sample_values(series, n=8):
    vals = []
    for v in series.tolist():
        if pd.isna(v) or str(v).strip() == "":
            continue
        if isinstance(v, pd.Timestamp):
            vals.append(v.strftime("%m/%d/%Y"))
        else:
            vals.append(str(v))
        if len(vals) >= n:
            break
    return vals


st.set_page_config(page_title=APP_TITLE, page_icon="📊", layout="wide")

st.title(APP_TITLE)
st.caption(
    "Upload MOR PDFs, Excel workbooks, or a ZIP of monthly reports. "
    "The app detects the data fields, shows sample values for verification, "
    "and exports only the columns you choose."
)

with st.expander("What changed in V2", expanded=False):
    st.markdown(
        """
        - **No fixed character slicing for the Fourth Creek MOR.** The app uses actual PDF x-coordinates.
        - **Multi-row MOR headers are presented as clean field names** for the saved Fourth Creek/KUB layout.
        - **Page 1 DO and Page 2 Secondary System DO are labeled separately.**
        - **Sample values are shown before export** so you can verify you picked the correct column.
        - Unknown text-based MOR layouts use a positional fallback and are clearly marked **Review required**.
        """
    )

uploads = st.file_uploader(
    "Upload MOR file(s)",
    type=["pdf","xls","xlsx", "xlsm", "zip"],
    accept_multiple_files=True,
    help="You can upload one month, several months, or a ZIP containing monthly MOR PDFs.",
)

if uploads:
    signature = tuple((u.name, len(u.getvalue())) for u in uploads)

    if st.session_state.get("upload_signature") != signature:
        st.session_state["upload_signature"] = signature
        st.session_state.pop("datasets", None)
        st.session_state.pop("errors", None)

    if st.button("Detect MOR Fields", type="primary"):
        datasets = []
        errors = []

        progress = st.progress(0)
        items = []

        for upload in uploads:
            try:
                unpacked = unpack_upload(upload.name, upload.getvalue())
                items.extend(unpacked)
            except Exception as exc:
                errors.append(f"{upload.name}: {exc}")

        total = max(len(items), 1)

        for idx, (name, data) in enumerate(items, start=1):
            try:
                detected = detect_file(name, data)
                if not detected:
                    errors.append(f"{name}: no usable daily table was detected.")
                datasets.extend(detected)
            except Exception as exc:
                errors.append(f"{name}: {exc}")

            progress.progress(idx / total)

        progress.empty()

        st.session_state["datasets"] = combine_same_named_datasets(datasets)
        st.session_state["errors"] = errors


if "datasets" in st.session_state:
    datasets = st.session_state["datasets"]
    errors = st.session_state.get("errors", [])

    if errors:
        with st.expander(f"Processing notes ({len(errors)})", expanded=False):
            for error in errors:
                st.write("• " + error)

    if not datasets:
        st.error(
            "No usable daily MOR table was found. If the PDF is a scan/image, "
            "this version does not OCR it. A text-based PDF or Excel workbook is required."
        )
        st.stop()

    st.success(f"Detected {len(datasets)} data set(s).")

    labels = []
    for i, ds in enumerate(datasets):
        labels.append(
            f"{ds.name} • {len(ds.dataframe):,} rows • {len(ds.dataframe.columns)} fields • {ds.confidence}"
        )

    selected_idx = st.selectbox(
        "Choose the table you want to extract from",
        options=range(len(datasets)),
        format_func=lambda i: labels[i],
    )

    ds = datasets[selected_idx]
    df = ds.dataframe.copy()

    col1, col2, col3 = st.columns(3)
    col1.metric("Rows detected", f"{len(df):,}")
    col2.metric("Fields detected", len(df.columns))
    col3.metric("Confidence", ds.confidence)

    for note in ds.notes:
        st.caption(note)

    st.subheader("1. Choose fields")

    default = ["Date"] if "Date" in df.columns else []
    selected = st.multiselect(
        "Detected fields",
        options=list(df.columns),
        default=default,
        placeholder="Choose one or more fields",
    )

    st.subheader("2. Verify the values")

    if not selected:
        st.info("Choose at least one field above.")
        st.stop()

    # Verification cards are critical: this prevents exporting the wrong DO,
    # grease, or neighboring field without the user seeing sample values first.
    non_date = [c for c in selected if c != "Date"]

    if non_date:
        for field in non_date:
            values = sample_values(df[field], 8)
            shown = " • ".join(values) if values else "No nonblank sample values detected"
            st.markdown(f"**{field}**")
            st.code(shown, language=None)

    preview = df[selected].copy()

    st.subheader("3. Preview")
    st.dataframe(preview, use_container_width=True, height=500)

    # QA information.
    if "Date" in preview.columns:
        valid_dates = pd.to_datetime(preview["Date"], errors="coerce").dropna()
        if len(valid_dates):
            st.caption(
                f"Date coverage: {valid_dates.min().strftime('%m/%d/%Y')} "
                f"through {valid_dates.max().strftime('%m/%d/%Y')}"
            )

    blank_cells = int(preview.isna().sum().sum())
    st.caption(f"Blank cells in selected output: {blank_cells:,}")

    st.subheader("4. Export")

    safe_stem = "MOR_extracted"
    if len(uploads) == 1:
        safe_stem = Path(uploads[0].name).stem + "_extracted"

    xlsx = to_excel_bytes(preview)
    st.download_button(
        "Download Excel",
        data=xlsx,
        file_name=f"{safe_stem}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    st.download_button(
        "Download CSV",
        data=preview.to_csv(index=False).encode("utf-8"),
        file_name=f"{safe_stem}.csv",
        mime="text/csv",
    )
