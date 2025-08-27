import streamlit as st
import pandas as pd
import os
from datetime import datetime

st.title("Excel to CSV Converter")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])
if uploaded_file:
    filename = uploaded_file.name
    if filename.endswith(".xlsx"):
        df = pd.read_excel(uploaded_file, engine="openpyxl")
    elif filename.endswith(".xls"):
        df = pd.read_excel(uploaded_file, engine="xlrd")
    else:
        st.error("Unsupported file type.")
        df = None

    if df is not None:
        st.write("Preview:", df.head())
        csv = df.to_csv(index=False).encode('utf-8')

        # Set file name to converted_{MM-DD-YYYY}.csv
        current_date = datetime.now().strftime("%m-%d-%Y")
        output_filename = f"converted_{current_date}.csv"
        output_path = os.path.join(
            "C:\\Users\\haruk\\OneDrive\\Desktop\\Projects\\hoa_automation", output_filename
        )

        # Save the file to the specified directory
        with open(output_path, "wb") as f:
            f.write(csv)

        st.success(f"File saved to {output_path}")

        st.download_button(
            label="Download as CSV",
            data=csv,
            file_name=output_filename,
            mime="text/csv"
        )

#run this: python -m streamlit run screen.py