from core.excel_writer import append_rows_to_sheet
from core.mapper import classify_table

import streamlit as st
import tempfile
from core.docx_extractor import extract_all_tables

st.set_page_config(page_title="Mapping Agent", layout="wide")
st.title("Requirements to Mapping Agent (DOCX → Excel)")

uploaded = st.file_uploader("Upload Requirements DOCX", type=["docx"])

if uploaded:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(uploaded.read())
        docx_path = tmp.name

    st.success("DOCX uploaded successfully.")

    tables = extract_all_tables(docx_path)
    st.write(f"Found {len(tables)} tables in the document.")

    with st.expander("Preview extracted tables"):
        for i, t in enumerate(tables[:3]):
            st.write(f"Table {i+1}")
            st.dataframe(t)

mapping_df = None

for df in tables:
    if classify_table(df) == "table_mapping":
        mapping_df = df
        break

if mapping_df is None:
    st.error("Could not find Table Mapping section")
    st.stop()

rows = []

for _, r in mapping_df.iterrows():
    row = {
        "data_entity": None,  # optional for MVP
        "source_field": r.get("Source Field") or r.get("Source Column"),
        "raw_table": r.get("Raw Table"),
        "raw_column": r.get("Raw Column"),
        "bronze_table": r.get("Bronze Table"),
        "bronze_column": r.get("Bronze Column"),
    }
    rows.append(row)

template_path = "templates/Master_Mapping_Template.xlsx"
output_path = "Generated_Mapping.xlsx"

append_rows_to_sheet(
    template_path=template_path,
    output_path=output_path,
    sheet_name="Source to Raw to Bronze",
    rows=rows
)
with open(output_path, "rb") as f:
    st.download_button(
        label="Download Excel Mapping",
        data=f,
        file_name="Generated_Mapping.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

