
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
st.write(
    "Upload Excel atau CSV, taip arahan analisis, "
    "dan download hasilnya. Tiada API berbayar diperlukan."
)

st.info(
    "Versi percuma ini menggunakan pemprosesan Python dan "
    "pengenalan arahan. Ia bukan model AI perbualan bebas."
)

uploaded = st.file_uploader(
    "Pilih fail Excel atau CSV",
    type=["xlsx", "xls", "csv"],
)


def read_data(file, sheet=None):
    file.seek(0)
    name = file.name.lower()

    if name.endswith(".csv"):
        try:
            return pd.read_csv(file, encoding="utf-8-sig")
        except UnicodeDecodeError:
            file.seek(0)
            return pd.read_csv(file, encoding="latin-1")

    return pd.read_excel(file, sheet_name=sheet)


def find_column(text, columns):
    """Cari nama kolum sebenar yang disebut dalam arahan."""
    text_lower = text.lower()

    # Utamakan nama kolum terpanjang untuk mengelakkan padanan separa.
    for column in sorted(columns, key=lambda c: len(str(c)), reverse=True):
        name = str(column)
        if name.lower() in text_lower:
            return column

    return None


def run_analysis(df, request):
    text = request.strip().lower()
    columns = list(df.columns)

    # 1. Unique combinations / distinct rows
    if any(word in text for word in [
        "unique", "distinct", "kombinasi unik", "nilai unik"
    ]):
        mentioned = [
            col for col in columns
            if str(col).lower() in text
        ]

        if mentioned:
            result = df[mentioned].drop_duplicates().reset_index(drop=True)
            description = (
                "Kombinasi unik untuk kolum: "
                + ", ".join(map(str, mentioned))
            )
        else:
            result = df.drop_duplicates().reset_index(drop=True)
            description = "Baris unik berdasarkan semua kolum."

        return result, description

    # 2. Missing / empty values
    if any(word in text for word in [
        "missing", "null", "kosong", "nilai hilang"
    ]):
        result = pd.DataFrame({
            "column": columns,
            "missing_count": [
                int(df[c].isna().sum()) for c in columns
            ],
            "missing_percent": [
                round(float(df[c].isna().mean() * 100), 2)
                for c in columns
            ],
        })
        result = result.sort_values(
            "missing_count", ascending=False
        ).reset_index(drop=True)
        return result, "Bilangan dan peratus nilai kosong."

    # 3. Duplicate rows
    if any(word in text for word in [
        "duplicate", "duplicates", "pengulangan", "pendua"
    ]):
        col = find_column(text, columns)

        if col is not None and any(
            word in text for word in ["by", "based on", "mengikut", "berdasarkan"]
        ):
            result = df[
                df.duplicated(subset=[col], keep=False)
            ].sort_values(col).reset_index(drop=True)
            description = f"Baris yang berulang berdasarkan {col}."
        else:
            result = df[
                df.duplicated(keep=False)
            ].reset_index(drop=True)
            description = "Baris yang berulang berdasarkan semua kolum."

        return result, description

    # 4. Group and count
    if any(word in text for word in [
        "group", "grouped", "count by", "count per",
        "kira mengikut", "bilangan mengikut"
    ]):
        col = find_column(text, columns)

        if col is None:
            return None, (
                "Saya tidak dapat mengenal pasti kolum untuk kiraan. "
                "Contoh: Count records grouped by node_name"
            )

        result = (
            df.groupby(col, dropna=False)
            .size()
            .reset_index(name="count")
            .sort_values("count", ascending=False)
            .reset_index(drop=True)
        )
        return result, f"Bilangan rekod mengikut {col}."

    # 5. Filter: contains
    if any(word in text for word in [
        "contains", "contain", "mengandungi"
    ]):
        col = find_column(text, columns)

        if col is None:
            return None, (
                "Nyatakan nama kolum. Contoh: "
                "Show rows where evpn_vrf_name contains L3"
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
            return None, "Nyatakan teks yang hendak dicari."

        value = match.group(1).strip().strip("'\"")
        # Elakkan nama kolum daripada termasuk dalam nilai carian.
        result = df[
            df[col].astype(str).str.contains(
                value, case=False, na=False, regex=False
            )
        ].reset_index(drop=True)

        return result, f"Rekod {col} yang mengandungi '{value}'."

    # 6. Column names and data types
    if any(word in text for word in [
        "column", "columns", "kolum", "data type", "jenis data"
    ]):
        result = pd.DataFrame({
            "column": columns,
            "data_type": [str(df[c].dtype) for c in columns],
            "missing_count": [
                int(df[c].isna().sum()) for c in columns
            ],
        })
        return result, "Senarai kolum dan jenis data."

    # 7. Overall counts
    if any(word in text for word in [
        "row count", "total rows", "how many rows",
        "jumlah baris", "jumlah rekod", "berapa rekod"
    ]):
        result = pd.DataFrame([
            {"metric": "total_rows", "value": len(df)},
            {"metric": "total_columns", "value": len(columns)},
            {"metric": "duplicate_rows", "value": int(df.duplicated().sum())},
        ])
        return result, "Ringkasan jumlah data."

    # 8. Sort by column
    if any(word in text for word in [
        "sort", "sort by", "susun"
    ]):
        col = find_column(text, columns)

        if col is None:
            return None, "Nyatakan kolum untuk disusun."

        descending = any(word in text for word in [
            "descending", "largest first", "terbesar dahulu"
        ])
        result = df.sort_values(
            col, ascending=not descending, na_position="last"
        ).reset_index(drop=True)

        return result, f"Data disusun berdasarkan {col}."

    return None, (
        "Arahan belum dikenali. Cuba salah satu contoh di bawah."
    )


if uploaded:
    sheet = None

    if uploaded.name.lower().endswith((".xlsx", ".xls")):
        try:
            workbook = pd.ExcelFile(uploaded)
            sheet = st.selectbox("Pilih worksheet", workbook.sheet_names)
        except Exception as error:
            st.error(f"Tidak dapat membaca fail Excel: {error}")
            st.stop()

    try:
        df = read_data(uploaded, sheet)
    except Exception as error:
        st.error(f"Tidak dapat membaca fail: {error}")
        st.stop()

    st.success(
        f"Fail berjaya dibaca: {len(df):,} baris, "
        f"{len(df.columns)} kolum."
    )

    with st.expander("Lihat preview data", expanded=True):
        st.dataframe(df.head(100), use_container_width=True)

    st.write("**Kolum tersedia:**", ", ".join(map(str, df.columns)))

    request = st.text_area(
        "Taip arahan analisis",
        placeholder=(
            "Contoh: Show unique combinations of "
            "agg_node_name, evpn_vrf_name, ip_gateway, and node_name"
        ),
        height=100,
    )

    with st.expander("Contoh arahan"):
        st.markdown("""
        - `Show unique combinations of agg_node_name, evpn_vrf_name, ip_gateway, and node_name`
        - `Count records grouped by node_name`
        - `Show rows where evpn_vrf_name contains L3`
        - `Find duplicate rows`
        - `Show missing values`
        - `List all columns and data types`
        - `Show total rows`
        - `Sort by node_name`
        """)

    if st.button("🔎 Analisis", type="primary"):
        if not request.strip():
            st.warning("Sila taip arahan terlebih dahulu.")
        else:
            result, description = run_analysis(df, request)
            st.subheader("Hasil analisis")
            st.write(description)

            if result is None:
                st.warning(description)
            else:
                st.caption(
                    f"{len(result):,} baris × {len(result.columns)} kolum"
                )
                st.dataframe(result, use_container_width=True)

                csv_bytes = result.to_csv(
                    index=False
                ).encode("utf-8-sig")

                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(
                    excel_buffer, engine="openpyxl"
                ) as writer:
                    result.to_excel(
                        writer, index=False, sheet_name="Result"
                    )

                left, right = st.columns(2)

                with left:
                    st.download_button(
                        "⬇️ Download CSV",
                        data=csv_bytes,
                        file_name="analysis_result.csv",
                        mime="text/csv",
                    )

                with right:
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
    st.info("Muat naik fail Excel atau CSV untuk bermula.")
