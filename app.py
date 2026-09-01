
import io
import hashlib
import re
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


class StoredUpload:
    """Small UploadedFile-compatible wrapper for files kept across upload batches."""

    def __init__(self, name: str, data: bytes):
        self.name = name
        self._data = data

    def getvalue(self) -> bytes:
        return self._data



def is_percentage_column(column_name: str) -> bool:
    """Recognize headers that explicitly identify the field as a percentage."""
    name = str(column_name).strip().lower()
    return (
        "%" in name
        or "percent" in name
        or "percentage" in name
    )


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()

    # Excel stores 75.6% as 0.756. MOR source tables commonly store that same
    # value as 75.6. Convert only the Excel-export copy so the workbook displays
    # 75.6%, not 7560%, while leaving the app's extracted data unchanged.
    export_df = df.copy()

    percentage_columns = [
        col for col in export_df.columns
        if is_percentage_column(col)
    ]

    for col in percentage_columns:
        numeric = pd.to_numeric(export_df[col], errors="coerce")
        nonblank = export_df[col].notna() & export_df[col].astype(str).str.strip().ne("")
        numeric_mask = nonblank & numeric.notna()

        # Preserve blanks and any nonnumeric OCR text exactly as-is.
        export_df.loc[numeric_mask, col] = numeric.loc[numeric_mask] / 100.0

    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="mm/dd/yyyy") as writer:
        export_df.to_excel(writer, sheet_name="Extracted Data", index=False, na_rep="")

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
        percent_fmt = workbook.add_format({"num_format": "0.0%"})

        for col_idx, col in enumerate(export_df.columns):
            worksheet.write(0, col_idx, col, header)
            sample = [str(x) for x in df[col].head(80).fillna("").tolist()]
            width = min(max([len(str(col))] + [len(x) for x in sample]) + 2, 38)
            worksheet.set_column(col_idx, col_idx, max(width, 12))

            if pd.api.types.is_datetime64_any_dtype(export_df[col]):
                worksheet.set_column(col_idx, col_idx, max(width, 12), date_fmt)
            elif is_percentage_column(col):
                worksheet.set_column(col_idx, col_idx, max(width, 12), percent_fmt)

        worksheet.set_row(0, 38)
        worksheet.freeze_panes(1, 0)

        if len(export_df) and len(export_df.columns):
            worksheet.add_table(
                0,
                0,
                len(export_df),
                len(export_df.columns) - 1,
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


st.set_page_config(page_title=APP_TITLE, page_icon="morganizer_icon.jpg", layout="wide")

st.title(APP_TITLE)
st.caption("Turn messy Monthly Operating Reports into clean, usable data.")
st.caption(
    "Upload MOR PDFs, scanned PDFs, Excel workbooks, or a ZIP of monthly reports. "
    "Choose the worksheet or scanned-PDF page you need, select the fields, "
    "verify the values, and export the result."
)

# -------------------------------------------------------------------
# Additive upload queue
# -------------------------------------------------------------------
# Streamlit can accept many files at once, but this queue also lets the user
# add another batch later without losing files that were already chosen.
if "stored_uploads" not in st.session_state:
    st.session_state["stored_uploads"] = []

if "uploader_generation" not in st.session_state:
    st.session_state["uploader_generation"] = 0

new_uploads = st.file_uploader(
    "Choose MOR file(s) to add",
    type=["pdf", "xls", "xlsx", "xlsm", "zip"],
    accept_multiple_files=True,
    help=(
        "Choose one or more files. After they are added, you can browse again "
        "and add another batch without losing the files already in MORganizer."
    ),
    key=f"mor_upload_batch_{st.session_state['uploader_generation']}",
)

if new_uploads:
    if st.button("Add selected files", type="primary", key="add_selected_uploads"):
        existing_fingerprints = {
            (item["name"], item["size"], item["sha256"])
            for item in st.session_state["stored_uploads"]
        }

        added = 0
        skipped = 0

        for upload in new_uploads:
            data = upload.getvalue()
            fingerprint = (
                upload.name,
                len(data),
                hashlib.sha256(data).hexdigest(),
            )

            if fingerprint in existing_fingerprints:
                skipped += 1
                continue

            st.session_state["stored_uploads"].append(
                {
                    "name": upload.name,
                    "data": data,
                    "size": len(data),
                    "sha256": fingerprint[2],
                }
            )
            existing_fingerprints.add(fingerprint)
            added += 1

        # Reset the native picker so another batch can be chosen immediately.
        st.session_state["uploader_generation"] += 1

        if added:
            st.success(f"Added {added} file{'s' if added != 1 else ''}.")
        if skipped:
            st.info(
                f"Skipped {skipped} duplicate file{'s' if skipped != 1 else ''} "
                "that were already added."
            )

        st.rerun()

stored = st.session_state["stored_uploads"]

if stored:
    total_mb = sum(item["size"] for item in stored) / (1024 * 1024)
    st.caption(
        f"**{len(stored)} file{'s' if len(stored) != 1 else ''} ready** "
        f"({total_mb:,.1f} MB total). You can add another batch above at any time."
    )

    with st.expander("Files currently added", expanded=False):
        for item in stored:
            st.write(f"• {item['name']}  —  {item['size'] / (1024 * 1024):,.1f} MB")

        if st.button("Clear all added files", key="clear_all_stored_uploads"):
            st.session_state["stored_uploads"] = []
            st.session_state["uploader_generation"] += 1

            for key in [
                "upload_signature",
                "datasets",
                "errors",
                "prepared_items",
                "prep_errors",
            ]:
                st.session_state.pop(key, None)

            for key in list(st.session_state.keys()):
                if (
                    key.startswith("selected_fields_dataset_")
                    or key.startswith("picker_widget_dataset_")
                    or key.startswith("global_field_sort_")
                    or key in {"global_field_order", "active_dataset_for_field_selection"}
                ):
                    st.session_state.pop(key, None)

            st.rerun()

uploads = [StoredUpload(item["name"], item["data"]) for item in stored]

if uploads:
    signature = tuple(
        (item["name"], item["size"], item["sha256"])
        for item in stored
    )

    if st.session_state.get("upload_signature") != signature:
        st.session_state["upload_signature"] = signature
        st.session_state.pop("datasets", None)
        st.session_state.pop("errors", None)
        st.session_state.pop("prepared_items", None)

        for key in list(st.session_state.keys()):
            if (
                key.startswith("selected_fields_dataset_")
                or key.startswith("picker_widget_dataset_")
                or key.startswith("global_field_sort_")
                or key in {"global_field_order", "active_dataset_for_field_selection"}
            ):
                st.session_state.pop(key, None)

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



def sanitize_download_name(name: str, fallback: str = "MOR_extracted") -> str:
    """
    Make a user-entered download filename safe across common operating systems.
    Extensions are removed here because the export buttons add .xlsx / .csv.
    """
    name = str(name or "").strip()

    # Remove an extension if the user typed one.
    name = re.sub(r"\.(xlsx|xls|csv)$", "", name, flags=re.IGNORECASE)

    # Replace characters that are invalid or awkward in Windows/macOS filenames.
    name = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", name)

    # Collapse repeated whitespace and trim trailing spaces/dots.
    name = re.sub(r"\s+", " ", name).strip(" .")

    return name or fallback


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

    # ------------------------------------------------------------------
    # Multi-page / multi-sheet field selection
    # ------------------------------------------------------------------
    st.subheader("1. Choose fields from any page or worksheet")
    st.caption(
        "Pick a page or worksheet, choose the fields you want, then switch to another one. "
        "Your earlier selections stay saved."
    )

    dataset_labels = [
        f"{ds.name} • {len(ds.dataframe):,} rows • {len(ds.dataframe.columns)} fields • {ds.confidence}"
        for ds in datasets
    ]

    active_idx = st.selectbox(
        "Page / worksheet to choose fields from",
        options=range(len(datasets)),
        format_func=lambda i: dataset_labels[i],
        key="active_dataset_for_field_selection",
    )

    active_ds = datasets[active_idx]
    active_df = active_ds.dataframe.copy()

    col1, col2, col3 = st.columns(3)
    col1.metric("Rows detected", f"{len(active_df):,}")
    col2.metric("Fields detected", len(active_df.columns))
    col3.metric("Confidence", active_ds.confidence)

    for note in active_ds.notes:
        st.caption(note)

    selection_key = f"selected_fields_dataset_{active_idx}"
    if selection_key not in st.session_state:
        if "Date" in active_df.columns:
            st.session_state[selection_key] = ["Date"]
        elif "Day" in active_df.columns:
            st.session_state[selection_key] = []
        else:
            st.session_state[selection_key] = []

    picked_here = st.multiselect(
        f"Fields from {active_ds.name}",
        options=list(active_df.columns),
        default=st.session_state[selection_key],
        placeholder="Choose one or more fields",
        key=f"picker_widget_dataset_{active_idx}",
    )
    st.session_state[selection_key] = picked_here

    # Collect selections from every page/sheet.
    selected_specs = []
    raw_name_counts = {}

    for ds_idx, ds in enumerate(datasets):
        chosen = st.session_state.get(f"selected_fields_dataset_{ds_idx}", [])
        for field in chosen:
            if field not in ds.dataframe.columns:
                continue
            uid = f"{ds_idx}::{field}"
            selected_specs.append(
                {
                    "uid": uid,
                    "dataset_idx": ds_idx,
                    "field": field,
                    "source": ds.name,
                }
            )
            raw_name_counts[field] = raw_name_counts.get(field, 0) + 1

    if not selected_specs:
        st.info("Choose at least one field from one of the detected pages or worksheets.")
        st.stop()

    key_fields = {"Date", "Day"}
    display_for_uid = {}
    uid_for_display = {}

    for spec in selected_specs:
        field = spec["field"]
        source = spec["source"]

        if field in key_fields:
            display = field
        elif raw_name_counts.get(field, 0) > 1:
            display = f"{field}  ·  {source}"
        else:
            display = field

        base_display = display
        counter = 2
        while display in uid_for_display and uid_for_display[display] != spec["uid"]:
            display = f"{base_display} ({counter})"
            counter += 1

        display_for_uid[spec["uid"]] = display
        uid_for_display[display] = spec["uid"]

    # Keep Date/Day only once in the global reorder list even if selected on
    # several pages/sheets.
    sortable_uids = []
    seen_global_keys = set()
    for spec in selected_specs:
        if spec["field"] in key_fields:
            global_key = spec["field"]
            if global_key in seen_global_keys:
                continue
            seen_global_keys.add(global_key)
        sortable_uids.append(spec["uid"])

    selected_uid_set = set(sortable_uids)
    previous_global_order = st.session_state.get("global_field_order", [])
    ordered_uids = [uid for uid in previous_global_order if uid in selected_uid_set]
    ordered_uids.extend(uid for uid in sortable_uids if uid not in ordered_uids)

    ordered_display = [display_for_uid[uid] for uid in ordered_uids]

    st.markdown("**Selected across all pages / worksheets**")
    st.caption(
        "Drag the fields below into the exact order you want. "
        "Fields from different pages or worksheets can be mixed freely."
    )

    selection_signature = hashlib.md5(
        "||".join(sorted(selected_uid_set)).encode("utf-8")
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

    sorted_display = sort_items(
        ordered_display,
        header=None,
        direction="vertical",
        custom_style=sortable_style,
        key=f"global_field_sort_{selection_signature}",
    )

    sorted_uids = []
    for label in sorted_display:
        uid = uid_for_display.get(label)
        if uid and uid in selected_uid_set and uid not in sorted_uids:
            sorted_uids.append(uid)

    for uid in ordered_uids:
        if uid not in sorted_uids:
            sorted_uids.append(uid)

    st.session_state["global_field_order"] = sorted_uids

    # ------------------------------------------------------------------
    # Combine fields from multiple pages/sheets into one daily table
    # ------------------------------------------------------------------
    selected_by_dataset = {}
    for uid in sorted_uids:
        ds_idx_str, field = uid.split("::", 1)
        ds_idx = int(ds_idx_str)
        selected_by_dataset.setdefault(ds_idx, [])
        if field not in selected_by_dataset[ds_idx]:
            selected_by_dataset[ds_idx].append(field)

    participating = sorted(selected_by_dataset)

    def choose_join_key(indices):
        if indices and all("Date" in datasets[i].dataframe.columns for i in indices):
            return "Date"
        if indices and all("Day" in datasets[i].dataframe.columns for i in indices):
            return "Day"
        return None

    join_key = choose_join_key(participating)
    merged = None
    column_for_uid = {}

    for ds_idx in participating:
        ds = datasets[ds_idx]
        source_df = ds.dataframe.copy()
        chosen_fields = selected_by_dataset[ds_idx]

        work_cols = list(chosen_fields)
        if join_key and join_key in source_df.columns and join_key not in work_cols:
            work_cols = [join_key] + work_cols

        part = source_df[work_cols].copy()

        rename_map = {}
        for field in chosen_fields:
            uid = f"{ds_idx}::{field}"

            if field == join_key:
                column_for_uid[uid] = join_key
                continue

            output_name = display_for_uid[uid]
            rename_map[field] = output_name
            column_for_uid[uid] = output_name

        part = part.rename(columns=rename_map)

        if merged is None:
            merged = part
        elif join_key:
            merged = merged.merge(part, on=join_key, how="outer", sort=False)
        else:
            merged = pd.concat(
                [merged.reset_index(drop=True), part.reset_index(drop=True)],
                axis=1,
            )

    if merged is None:
        st.info("Choose at least one field.")
        st.stop()

    final_columns = []
    for uid in sorted_uids:
        _, field = uid.split("::", 1)

        if field in key_fields:
            if field in merged.columns and field not in final_columns:
                final_columns.append(field)
            continue

        col_name = column_for_uid.get(uid)
        if col_name in merged.columns and col_name not in final_columns:
            final_columns.append(col_name)

    preview = merged[final_columns].copy()

    if join_key == "Date" and "Date" in preview.columns:
        preview["Date"] = pd.to_datetime(preview["Date"], errors="coerce")
        preview = preview.sort_values("Date", kind="stable").reset_index(drop=True)
    elif join_key == "Day" and "Day" in preview.columns:
        preview = preview.sort_values("Day", kind="stable").reset_index(drop=True)

    # ------------------------------------------------------------------
    # Verification
    # ------------------------------------------------------------------
    st.subheader("2. Verify the values")

    for col in preview.columns:
        if col in {"Date", "Day"}:
            continue
        values = sample_values(preview[col], 8)
        shown = " • ".join(values) if values else "No nonblank sample values detected"
        st.markdown(f"**{col}**")
        st.code(shown, language=None)

    if len(participating) > 1:
        if join_key:
            st.caption(
                f"Fields from {len(participating)} detected pages/worksheets were combined using "
                f"{join_key} so values stay on the correct daily row."
            )
        else:
            st.warning(
                "These selected tables do not share a Date or Day field, so MORganizer aligned "
                "them by row position. Verify the preview carefully before export."
            )

    # ------------------------------------------------------------------
    # Preview + export
    # ------------------------------------------------------------------
    st.subheader("3. Preview")
    display_preview = preview.astype(object).where(pd.notna(preview), "")
    st.dataframe(display_preview, use_container_width=True, height=500)

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

    default_stem = "MOR_extracted"
    if len(uploads) == 1:
        default_stem = Path(uploads[0].name).stem + "_extracted"

    requested_name = st.text_input(
        "Name your extracted file",
        value=default_stem,
        help="You do not need to type .xlsx or .csv. MORganizer will add the correct extension.",
        key="custom_export_filename",
    )

    safe_stem = sanitize_download_name(requested_name, fallback=default_stem)

    if requested_name.strip() and safe_stem != requested_name.strip():
        st.caption(f'Your download filename will be saved as: **{safe_stem}**')

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
