
import io
import re

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Excel AI Analyzer",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Excel AI Analyzer")
st.caption(
    "Upload multiple Excel/CSV files, compare worksheet headers, "
    "analyse data and download results."
)

st.info(
    "Free version: Python-based analysis, no paid AI API or API key required. "
    "Commands are rule-based, not free-form generative AI."
)


# --------------------------------------------------
# FILE LOADING
# --------------------------------------------------

uploads = st.file_uploader(
    "Upload one or more Excel/CSV files",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True,
)


def load_file(file):
    """Load every worksheet from an Excel file, or one CSV dataset."""
    file.seek(0)
    filename = file.name.lower()

    if filename.endswith(".csv"):
        try:
            df = pd.read_csv(file, encoding="utf-8-sig")
        except UnicodeDecodeError:
            file.seek(0)
            df = pd.read_csv(file, encoding="latin-1")

        return {"CSV": df}

    workbook = pd.ExcelFile(file)
    result = {}

    for sheet_name in workbook.sheet_names:
        file.seek(0)
        result[sheet_name] = pd.read_excel(
            file,
            sheet_name=sheet_name,
        )

    return result


def normalise_name(value):
    """Normalise a worksheet name for case-insensitive matching."""
    return re.sub(r"\s+", "", str(value)).casefold()


def get_sheet(sheets, requested_name):
    """Find a worksheet by its actual name."""
    for actual_name in sheets:
        if normalise_name(actual_name) == normalise_name(requested_name):
            return actual_name

    return None


# --------------------------------------------------
# HEADER COMPARISON
# --------------------------------------------------

def compare_headers(df1, df2, file1, sheet1, file2, sheet2):
    """
    Compare column headers by exact text.
    Header order is preserved from File 1, followed by headers
    found only in File 2.
    """
    headers1 = [str(col) for col in df1.columns]
    headers2 = [str(col) for col in df2.columns]

    set1 = set(headers1)
    set2 = set(headers2)

    common = [h for h in headers1 if h in set2]
    only1 = [h for h in headers1 if h not in set2]
    only2 = [h for h in headers2 if h not in set1]

    rows = []

    for header in headers1:
        rows.append({
            "header": header,
            "file_1": file1,
            "sheet_1": sheet1,
            "file_2": file2,
            "sheet_2": sheet2,
            "status": (
                "Same"
                if header in set2
                else "Only in File 1"
            ),
        })

    for header in headers2:
        if header not in set1:
            rows.append({
                "header": header,
                "file_1": file1,
                "sheet_1": sheet1,
                "file_2": file2,
                "sheet_2": sheet2,
                "status": "Only in File 2",
            })

    result = pd.DataFrame(rows)

    total1 = len(headers1)
    total2 = len(headers2)
    matched = len(common)

    summary = {
        "file1": file1,
        "sheet1": sheet1,
        "file2": file2,
        "sheet2": sheet2,
        "total1": total1,
        "total2": total2,
        "same": matched,
        "only1": len(only1),
        "only2": len(only2),
        "same_headers": common,
        "only1_headers": only1,
        "only2_headers": only2,
    }

    return result, summary


# --------------------------------------------------
# GENERAL DATA ANALYSIS
# --------------------------------------------------

def find_column(request, columns):
    """Find a column mentioned in the user's command."""
    request_lower = request.lower()

    # Longest names first to avoid partial matches.
    for column in sorted(
        columns,
        key=lambda col: len(str(col)),
        reverse=True,
    ):
        if str(column).lower() in request_lower:
            return column

    return None


def analyse_data(df, request):
    text = request.strip().lower()
    columns = list(df.columns)

    # Unique rows or unique combinations
    if any(word in text for word in [
        "unique",
        "distinct",
        "kombinasi unik",
        "nilai unik",
    ]):
        selected = [
            col for col in columns
            if str(col).lower() in text
        ]

        if selected:
            result = (
                df[selected]
                .drop_duplicates()
                .reset_index(drop=True)
            )
            return result, (
                "Unique combinations for: "
                + ", ".join(map(str, selected))
            )

        return (
            df.drop_duplicates().reset_index(drop=True),
            "Unique rows based on all columns.",
        )

    # Missing values
    if any(word in text for word in [
        "missing",
        "null",
        "kosong",
        "nilai hilang",
    ]):
        result = pd.DataFrame({
            "column": columns,
            "missing_count": [
                int(df[col].isna().sum())
                for col in columns
            ],
            "missing_percent": [
                round(float(df[col].isna().mean() * 100), 2)
                for col in columns
            ],
        })

        result = result.sort_values(
            "missing_count",
            ascending=False,
        ).reset_index(drop=True)

        return result, "Missing values by column."

    # Duplicate rows
    if any(word in text for word in [
        "duplicate",
        "duplicates",
        "pengulangan",
        "pendua",
    ]):
        col = find_column(request, columns)

        if col and any(word in text for word in [
            "based on",
            "by",
            "berdasarkan",
            "mengikut",
        ]):
            result = df[
                df.duplicated(subset=[col], keep=False)
            ].sort_values(col).reset_index(drop=True)

            return result, f"Duplicate rows based on {col}."

        result = df[
            df.duplicated(keep=False)
        ].reset_index(drop=True)

        return result, "Duplicate rows based on all columns."

    # Group and count
    if any(word in text for word in [
        "group",
        "grouped",
        "count by",
        "count per",
        "kira mengikut",
        "bilangan mengikut",
    ]):
        col = find_column(request, columns)

        if col is None:
            return None, (
                "Please specify a column. "
                "Example: Count records grouped by node_name"
            )

        result = (
            df.groupby(col, dropna=False)
            .size()
            .reset_index(name="count")
            .sort_values("count", ascending=False)
            .reset_index(drop=True)
        )

        return result, f"Record count grouped by {col}."

    # Filter text using contains
    if any(word in text for word in [
        "contains",
        "contain",
        "mengandungi",
    ]):
        col = find_column(request, columns)

        if col is None:
            return None, (
                "Please specify the column. "
                "Example: Show rows where evpn_vrf_name contains L3"
            )

        match = re.search(
            r"\bcontains?\b\s+['\"]?(.+?)['\"]?\s*$",
            request,
            flags=re.IGNORECASE,
        )

        if not match:
            match = re.search(
                r"mengandungi\s+['\"]?(.+?)['\"]?\s*$",
                request,
                flags=re.IGNORECASE,
            )

        if not match:
            return None, "Please specify the text to search for."

        value = match.group(1).strip().strip("'\"")

        result = df[
            df[col].astype(str).str.contains(
                value,
                case=False,
                na=False,
                regex=False,
            )
        ].reset_index(drop=True)

        return result, f"Rows where {col} contains '{value}'."

    # Column names and data types
    if any(word in text for word in [
        "column",
        "columns",
        "kolum",
        "data type",
        "jenis data",
    ]):
        result = pd.DataFrame({
            "column": columns,
            "data_type": [
                str(df[col].dtype)
                for col in columns
            ],
            "missing_count": [
                int(df[col].isna().sum())
                for col in columns
            ],
        })

        return result, "Column names and data types."

    # Overall counts
    if any(word in text for word in [
        "row count",
        "total rows",
        "how many rows",
        "jumlah baris",
        "jumlah rekod",
        "berapa rekod",
    ]):
        result = pd.DataFrame([
            {"metric": "total_rows", "value": len(df)},
            {"metric": "total_columns", "value": len(columns)},
            {
                "metric": "duplicate_rows",
                "value": int(df.duplicated().sum()),
            },
        ])

        return result, "Data summary."

    # Sort
    if "sort" in text or "susun" in text:
        col = find_column(request, columns)

        if col is None:
            return None, "Please specify a column to sort."

        descending = any(word in text for word in [
            "descending",
            "largest first",
            "terbesar dahulu",
        ])

        result = df.sort_values(
            col,
            ascending=not descending,
            na_position="last",
        ).reset_index(drop=True)

        return result, f"Data sorted by {col}."

    return None, (
        "Command not recognised. Please use one of the examples below."
    )


# --------------------------------------------------
# MAIN APPLICATION
# --------------------------------------------------

if uploads:
    loaded_files = []
    errors = []

    # File 1, File 2, etc. follow the uploader's displayed order.
    for file in uploads:
        try:
            sheets = load_file(file)
            loaded_files.append({
                "filename": file.name,
                "sheets": sheets,
            })
        except Exception as error:
            errors.append(f"{file.name}: {error}")

    for error in errors:
        st.error(f"Could not read file: {error}")

    if loaded_files:
        st.subheader("Uploaded files")

        for number, item in enumerate(loaded_files, start=1):
            details = []

            for sheet_name, df in item["sheets"].items():
                details.append(
                    f"{sheet_name} "
                    f"({len(df):,} rows, {len(df.columns)} columns)"
                )

            st.markdown(
                f"**File {number}: {item['filename']}**  \n"
                + " | ".join(details)
            )

        st.divider()
        st.subheader("Analysis command")

        request = st.text_area(
            "Type your command",
            placeholder=(
                "Compare headers from file 1 - Sheet1 "
                "vs file 2 - Sheet1"
            ),
            height=100,
        )

        with st.expander("Supported command examples", expanded=True):
            st.markdown("""
            **Compare headers across files**
            - `Compare headers from file 1 - Sheet1 vs file 2 - Sheet1`
            - `Compare headers from file 1 - Sheet2 vs file 3 - Sheet1`

            **Analyse one worksheet**
            - `Show unique combinations of agg_node_name, evpn_vrf_name, ip_gateway, and node_name`
            - `Count records grouped by node_name`
            - `Show rows where evpn_vrf_name contains L3`
            - `Find duplicate rows`
            - `Show missing values`
            - `List all columns and data types`
            - `Show total rows`
            - `Sort by node_name`
            """)

        if st.button("🔎 Analyse", type="primary"):
            if not request.strip():
                st.warning("Please enter a command.")
            else:
                # Header comparison command parser
                compare_pattern = re.compile(
                    r"file\s*(\d+)\s*[-,:]?\s*"
                    r"(?:sheet\s*)?(.+?)\s*"
                    r"(?:vs\.?|versus|and)\s*"
                    r"file\s*(\d+)\s*[-,:]?\s*"
                    r"(?:sheet\s*)?(.+?)\s*$",
                    flags=re.IGNORECASE,
                )

                is_comparison = (
                    any(word in request.lower() for word in [
                        "compare",
                        "comparison",
                        "bandingkan",
                    ])
                    and any(word in request.lower() for word in [
                        "header",
                        "headers",
                        "column",
                        "columns",
                        "kolum",
                    ])
                )

                match = compare_pattern.search(request)

                if is_comparison:
                    if not match:
                        st.error(
                            "Could not read the comparison command. "
                            "Use: Compare headers from file 1 - Sheet1 "
                            "vs file 2 - Sheet1"
                        )
                    else:
                        file1_num = int(match.group(1))
                        sheet1_text = match.group(2).strip(" -,:")
                        file2_num = int(match.group(3))
                        sheet2_text = match.group(4).strip(" -,:")
                        
                        if (
                            file1_num < 1
                            or file1_num > len(loaded_files)
                            or file2_num < 1
                            or file2_num > len(loaded_files)
                        ):
                            st.error(
                                f"Invalid file number. "
                                f"You uploaded {len(loaded_files)} file(s)."
                            )
                        else:
                            item1 = loaded_files[file1_num - 1]
                            item2 = loaded_files[file2_num - 1]

                            sheet1 = get_sheet(
                                item1["sheets"], sheet1_text
                            )
                            sheet2 = get_sheet(
                                item2["sheets"], sheet2_text
                            )

                            if sheet1 is None:
                                st.error(
                                    f"Worksheet '{sheet1_text}' not found "
                                    f"in File {file1_num} "
                                    f"({item1['filename']})."
                                )
                            elif sheet2 is None:
                                st.error(
                                    f"Worksheet '{sheet2_text}' not found "
                                    f"in File {file2_num} "
                                    f"({item2['filename']})."
                                )
                            else:
                                result, summary = compare_headers(
                                    item1["sheets"][sheet1],
                                    item2["sheets"][sheet2],
                                    item1["filename"],
                                    sheet1,
                                    item2["filename"],
                                    sheet2,
                                )

                                st.subheader("Header comparison result")

                                st.markdown(
                                    f"**File {file1_num}:** "
                                    f"{summary['file1']} — {summary['sheet1']}  \n"
                                    f"**File {file2_num}:** "
                                    f"{summary['file2']} — {summary['sheet2']}"
                                )

                                c1, c2, c3 = st.columns(3)

                                c1.metric(
                                    "Same headers",
                                    summary["same"],
                                )
                                c2.metric(
                                    f"Only in File {file1_num}",
                                    summary["only1"],
                                )
                                c3.metric(
                                    f"Only in File {file2_num}",
                                    summary["only2"],
                                )

                                st.dataframe(
                                    result,
                                    use_container_width=True,
                                )

                                with st.expander("Header lists"):
                                    st.write(
                                        "**Headers in both files**",
                                        summary["same_headers"],
                                    )
                                    st.write(
                                        f"**Only in File {file1_num}**",
                                        summary["only1_headers"],
                                    )
                                    st.write(
                                        f"**Only in File {file2_num}**",
                                        summary["only2_headers"],
                                    )

                                csv_bytes = result.to_csv(
                                    index=False
                                ).encode("utf-8-sig")

                                excel_buffer = io.BytesIO()

                                with pd.ExcelWriter(
                                    excel_buffer,
                                    engine="openpyxl",
                                ) as writer:
                                    result.to_excel(
                                        writer,
                                        index=False,
                                        sheet_name="Header Comparison",
                                    )

                                col_a, col_b = st.columns(2)

                                with col_a:
                                    st.download_button(
                                        "⬇️ Download CSV",
                                        data=csv_bytes,
                                        file_name="header_comparison.csv",
                                        mime="text/csv",
                                    )

                                with col_b:
                                    st.download_button(
                                        "⬇️ Download Excel",
                                        data=excel_buffer.getvalue(),
                                        file_name="header_comparison.xlsx",
                                        mime=(
                                            "application/vnd.openxmlformats-officedocument."
                                            "spreadsheetml.sheet"
                                        ),
                                    )

                else:
                    # Select a dataset for normal analysis.
                    choices = []

                    for number, item in enumerate(
                        loaded_files,
                        start=1,
                    ):
                        for sheet_name in item["sheets"]:
                            choices.append(
                                f"File {number}: "
                                f"{item['filename']} — {sheet_name}"
                            )

                    selected = st.selectbox(
                        "Dataset for analysis",
                        choices,
                    )

                    selected_index = choices.index(selected)
                    index = 0
                    selected_df = None

                    for item in loaded_files:
                        for df in item["sheets"].values():
                            if index == selected_index:
                                selected_df = df
                                break
                            index += 1
                        if selected_df is not None:
                            break

                    result, description = analyse_data(
                        selected_df,
                        request,
                    )

                    st.write(description)

                    if result is not None:
                        st.caption(
                            f"{len(result):,} rows × "
                            f"{len(result.columns)} columns"
                        )

                        st.dataframe(
                            result,
                            use_container_width=True,
                        )

                        csv_bytes = result.to_csv(
                            index=False
                        ).encode("utf-8-sig")

                        excel_buffer = io.BytesIO()

                        with pd.ExcelWriter(
                            excel_buffer,
                            engine="openpyxl",
                        ) as writer:
                            result.to_excel(
                                writer,
                                index=False,
                                sheet_name="Result",
                            )

                        col_a, col_b = st.columns(2)

                        with col_a:
                            st.download_button(
                                "⬇️ Download CSV",
                                data=csv_bytes,
                                file_name="analysis_result.csv",
                                mime="text/csv",
                            )

                        with col_b:
                            st.download_button(
                                "⬇️ Download Excel",
                                data=excel_buffer.getvalue(),
                                file_name="analysis_result.xlsx",
                                mime=(
                                    "application/vnd.openxmlformats-officedocument."
                                    "spreadsheetml.sheet"
                                ),
                            )

else:
    st.info("Upload one or more Excel/CSV files to begin.")
