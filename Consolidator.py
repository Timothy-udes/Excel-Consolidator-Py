import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Excel Multi-Sheet Consolidator")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:

    excel_file = pd.ExcelFile(uploaded_file)
    sheet_names = excel_file.sheet_names

    st.write("Sheets found:", sheet_names)

    sheets_to_merge = [s for s in sheet_names if s.lower() != "consolidated"]

    dataframes = []

    for sheet in sheets_to_merge:
        df = pd.read_excel(uploaded_file, sheet_name=sheet)

        # Add column to track source sheet
        df["SourceSheet"] = sheet

        dataframes.append(df)

    consolidated_df = pd.concat(dataframes, ignore_index=True)

    st.subheader("Preview of Consolidated Data")
    st.dataframe(consolidated_df)

    # Create downloadable Excel file
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        consolidated_df.to_excel(writer, sheet_name="Consolidated", index=False)

    st.download_button(
        label="Download Consolidated Excel",
        data=output.getvalue(),
        file_name="consolidated_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )