
import io
from pathlib import Path

import pandas as pd
import streamlit as st

from mor_parser import (
    combine_same_named_datasets,
    detect_file,
    get_excel_sheet_names,
    unpack_upload,
)


APP_TITLE = "MOR Column Extractor"


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


def prepare_items(uploads):
    items = []
    errors = []

    for upload in uploads:
        try:
            unpacked = unpack_upload(upload.name, upload.getvalue())
            items.extend(unpacked)
        except Exception as exc:
            errors.append(f"{upload.name}: {exc}")

    return items, errors


st.set_page_config(page_title=APP_TITLE, page_icon="📊", layout="wide")

st.title(APP_TITLE)
st.caption(
    "Upload MOR PDFs, Excel workbooks, or a ZIP of monthly reports. "
    "Choose the worksheet when an Excel file has multiple tabs, select the fields you need, "
    "verify the values, and export the result."
)

uploads = st.file_uploader(
    "Upload MOR file(s)",
    type=["pdf", "xls", "xlsx", "xlsm", "zip"],
    accept_multiple_files=True,
    help="You can upload one month, several months, or a ZIP containing monthly MOR files.",
)

if uploads:
    signature = tuple((u.name, len(u.getvalue())) for u in uploads)

    if st.session_state.get("upload_signature") != signature:
        st.session_state["upload_signature"] = signature
        st.session_state.pop("datasets", None)
        st.session_state.pop("errors", None)
        st.session_state.pop("prepared_items", None)

    if "prepared_items" not in st.session_state:
        items, prep_errors = prepare_items(uploads)
        st.session_state["prepared_items"] = items
        st.session_state["prep_errors"] = prep_errors

    items = st.session_state["prepared_items"]
    prep_errors = st.session_state.get("prep_errors", [])

    # ---------------------------------------------------------------
    # Worksheet selection
    # ---------------------------------------------------------------
    excel_items = [
        (name, data)
        for name, data in items
        if Path(name).suffix.lower() in {".xls", ".xlsx", ".xlsm"}
    ]
    selected_sheets_by_file = {}

    if excel_items:
        st.subheader("Choose worksheet(s)")

        sheet_lists = {}
        sheet_errors = []

        for name, data in excel_items:
            try:
                sheet_lists[name] = get_excel_sheet_names(name, data)
            except Exception as exc:
                sheet_lists[name] = []
                sheet_errors.append(f"{name}: {exc}")

        for err in sheet_errors:
            st.error(f"Could not read worksheet names from {err}")

        # Build a case-insensitive map of worksheet names shared by every workbook.
        common_keys = None
        display_name_by_key = {}

        for name, sheet_names in sheet_lists.items():
            current = {}
            for s in sheet_names:
                key = s.strip().lower()
                current[key] = s
                display_name_by_key.setdefault(key, s)

            keys = set(current)
            common_keys = keys if common_keys is None else common_keys & keys

        common_keys = common_keys or set()
        common_options = [display_name_by_key[k] for k in sorted(common_keys)]

        if len(excel_items) > 1 and common_options:
            preferred = next(
                (
                    s for s in common_options
                    if s.strip().lower() == "state"
                ),
                next(
                    (
                        s for s in common_options
                        if any(k in s.lower() for k in ("dmr", "mor", "state"))
                    ),
                    common_options[0],
                ),
            )

            apply_same = st.checkbox(
                "Use the same worksheet for all uploaded Excel files",
                value=True,
                help=(
                    "Choose one worksheet once and apply it to every uploaded workbook. "
                    "Turn this off if a particular file needs a different tab."
                ),
            )

            if apply_same:
                chosen_common = st.selectbox(
                    "Worksheet to use for all files",
                    options=common_options,
                    index=common_options.index(preferred),
                    help="Only worksheets found in every uploaded workbook are listed here.",
                )

                chosen_key = chosen_common.strip().lower()

                for name, sheet_names in sheet_lists.items():
                    actual = next(
                        (s for s in sheet_names if s.strip().lower() == chosen_key),
                        None,
                    )
                    selected_sheets_by_file[name] = [actual] if actual else []

                st.caption(
                    f'Using "{chosen_common}" for all {len(excel_items)} Excel file(s).'
                )

            else:
                st.caption("Choose worksheet(s) separately for each workbook.")

                for idx, (name, data) in enumerate(excel_items):
                    sheet_names = sheet_lists.get(name, [])
                    if not sheet_names:
                        continue

                    default_sheet = sheet_names[0]
                    preferred_individual = next(
                        (
                            s for s in sheet_names
                            if s.strip().lower() == "state"
                        ),
                        next(
                            (
                                s for s in sheet_names
                                if any(k in s.lower() for k in ("dmr", "mor", "state"))
                            ),
                            default_sheet,
                        ),
                    )

                    selection = st.multiselect(
                        f"{name}",
                        options=sheet_names,
                        default=[preferred_individual],
                        key=f"sheets_{idx}_{name}",
                        help="Select one or more tabs for this workbook.",
                    )
                    selected_sheets_by_file[name] = selection

        else:
            # Single Excel file, or multiple workbooks with no shared tab names.
            if len(excel_items) > 1 and not common_options:
                st.info(
                    "These workbooks do not share a common worksheet name, "
                    "so choose the worksheet separately for each file."
                )

            for idx, (name, data) in enumerate(excel_items):
                sheet_names = sheet_lists.get(name, [])
                if not sheet_names:
                    continue

                default_sheet = sheet_names[0]
                preferred_individual = next(
                    (
                        s for s in sheet_names
                        if s.strip().lower() == "state"
                    ),
                    next(
                        (
                            s for s in sheet_names
                            if any(k in s.lower() for k in ("dmr", "mor", "state"))
                        ),
                        default_sheet,
                    ),
                )

                selection = st.multiselect(
                    f"{name}",
                    options=sheet_names,
                    default=[preferred_individual],
                    key=f"sheets_{idx}_{name}",
                    help="Select one or more tabs for the parser to inspect.",
                )
                selected_sheets_by_file[name] = selection

        st.caption(
            "Only the worksheet tabs selected above will be parsed. "
            "PDF files do not need a worksheet selection."
        )

    if st.button("Detect MOR Fields", type="primary"):
        datasets = []
        errors = list(prep_errors)

        progress = st.progress(0)
        total = max(len(items), 1)

        for idx, (name, data) in enumerate(items, start=1):
            try:
                selected_sheets = selected_sheets_by_file.get(name)

                # If an Excel workbook is present but the user selected no tabs,
                # skip it instead of silently parsing every worksheet.
                if Path(name).suffix.lower() in {".xls", ".xlsx", ".xlsm"} and selected_sheets == []:
                    errors.append(f"{name}: no worksheet was selected.")
                    progress.progress(idx / total)
                    continue

                detected = detect_file(name, data, selected_sheets=selected_sheets)
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

    labels = [
        f"{ds.name} • {len(ds.dataframe):,} rows • {len(ds.dataframe.columns)} fields • {ds.confidence}"
        for ds in datasets
    ]

    selected_idx = st.selectbox(
        "Choose the detected table you want to extract from",
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
