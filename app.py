
import io
import hashlib
from pathlib import Path

import pandas as pd
import streamlit as st
from streamlit_sortables import sort_items

from mor_parser import (
    combine_same_named_datasets,
    detect_file,
    get_excel_sheet_names,
    get_pdf_page_info,
    unpack_upload,
)


APP_TITLE = "MORganizer 3000"


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="mm/dd/yyyy") as writer:
        df.to_excel(writer, sheet_name="Extracted Data", index=False, na_rep="")

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
st.caption("Turn messy Monthly Operating Reports into clean, usable data.")
st.caption(
    "Upload MOR PDFs, scanned PDFs, Excel workbooks, or a ZIP of monthly reports. "
    "Choose the worksheet or scanned-PDF page you need, select the fields, "
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

    # ---------------------------------------------------------------
    # Scanned PDF page selection
    # ---------------------------------------------------------------
    pdf_items = [
        (name, data)
        for name, data in items
        if Path(name).suffix.lower() == ".pdf"
    ]
    selected_pdf_pages_by_file = {}

    scanned_pdf_count = 0
    if pdf_items:
        pdf_info_by_file = {}

        for name, data in pdf_items:
            try:
                pdf_info_by_file[name] = get_pdf_page_info(data)
            except Exception as exc:
                prep_errors.append(f"{name}: could not inspect PDF pages ({exc})")

        scanned_items = [
            (name, data, pdf_info_by_file[name])
            for name, data in pdf_items
            if name in pdf_info_by_file and pdf_info_by_file[name].get("is_scanned")
        ]

        scanned_pdf_count = len(scanned_items)

        if scanned_items:
            st.subheader("Choose scanned PDF page(s)")
            st.info(
                "Scanned PDF detected. MORganizer 3000 will OCR only the pages you select, "
                "so supporting lab reports and attachments do not have to be processed."
            )

            for idx, (name, data, info) in enumerate(scanned_items):
                page_count = int(info.get("page_count", 0))
                options = list(range(1, page_count + 1))
                default_pages = [
                    p for p in info.get("suggested_pages", [1])
                    if p in options
                ] or ([1] if options else [])

                selected_pages = st.multiselect(
                    f"{name} — {page_count} page(s)",
                    options=options,
                    default=default_pages,
                    key=f"pdf_pages_{idx}_{name}",
                    help=(
                        "Select only the page(s) containing the monthly operating table. "
                        "Page 1 is suggested by default because the primary MOR is commonly first."
                    ),
                )
                selected_pdf_pages_by_file[name] = selected_pages

            st.caption(
                "OCR preserves word positions so values can be assigned back to table columns. "
                "Because scans can be faint or skewed, always verify the sample values before export."
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

                selected_pdf_pages = selected_pdf_pages_by_file.get(name)

                if Path(name).suffix.lower() == ".pdf" and selected_pdf_pages == []:
                    errors.append(f"{name}: no PDF page was selected.")
                    progress.progress(idx / total)
                    continue

                detected = detect_file(
                    name,
                    data,
                    selected_sheets=selected_sheets,
                    selected_pdf_pages=selected_pdf_pages,
                )
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
            "No usable daily MOR table was found. For a scanned PDF, try selecting the page "
            "that contains the daily operating table and make sure the day/date column is visible."
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

    st.subheader("1. Choose and arrange fields")

    default = ["Date"] if "Date" in df.columns else []
    selected = st.multiselect(
        "Detected fields",
        options=list(df.columns),
        default=default,
        placeholder="Choose one or more fields",
        key=f"field_picker_{selected_idx}",
    )

    if not selected:
        st.info("Choose at least one field above.")
        st.stop()

    # Preserve the user's previous drag order for fields that remain selected,
    # then append newly selected fields at the end.
    order_state_key = f"field_order_{selected_idx}"
    previous_order = st.session_state.get(order_state_key, [])

    ordered_selected = [c for c in previous_order if c in selected]
    ordered_selected.extend(c for c in selected if c not in ordered_selected)

    st.markdown("**Drag fields into the order you want in the exported file:**")
    st.caption("Grab any field and move it up or down. The preview, Excel, and CSV will follow this order.")

    # Changing the selected field set intentionally creates a fresh sortable
    # component, which avoids stale component state when fields are added/removed.
    selection_signature = hashlib.md5(
        "||".join(sorted(selected)).encode("utf-8")
    ).hexdigest()[:12]

    sortable_style = """
    .sortable-component {
        padding: 0.15rem 0;
        counter-reset: field-order;
    }
    .sortable-item {
        padding: 0.55rem 0.8rem;
        margin-bottom: 0.35rem;
        border-radius: 0.45rem;
        cursor: grab;
    }
    .sortable-item::before {
        content: counter(field-order) ". ";
        counter-increment: field-order;
        font-weight: 700;
    }
    """

    selected = sort_items(
        ordered_selected,
        header=None,
        direction="vertical",
        custom_style=sortable_style,
        key=f"field_sort_{selected_idx}_{selection_signature}",
    )

    # Defensive cleanup in case a browser/component refresh returns stale data.
    selected = [c for c in selected if c in df.columns and c in set(ordered_selected)]
    for c in ordered_selected:
        if c not in selected:
            selected.append(c)

    st.session_state[order_state_key] = selected

    st.subheader("2. Verify the values")

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
        data=preview.to_csv(index=False, na_rep="").encode("utf-8"),
        file_name=f"{safe_stem}.csv",
        mime="text/csv",
    )
