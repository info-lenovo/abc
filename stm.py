
import streamlit as st
import pdfplumber
import pandas as pd
import os
import re
from datetime import datetime

def clean_dataframe(df):
    def clean_cell(x):
        if pd.isnull(x):
            return x
        s = str(x).replace('\n', '').strip()

        # Remove leading 'F ', 'L ', 'A ' only when followed by a space or a number
        if re.match(r'^[FLA]\s+', s):
            s = re.sub(r'^[FLA]\s+', '', s)
        elif re.match(r'^[FLA](\d+(\.\d+)?|\.\d+)$', s):  # A123, A0.00, A.50 etc.
            s = re.sub(r'^[FLA]', '', s)

        return s

    df = df.applymap(clean_cell)

    # Drop mostly empty columns
    threshold = 0.9
    df = df.dropna(axis=1, thresh=int((1 - threshold) * len(df)))

    # Drop trailing empty columns (like N to Q if blank)
    last_valid_col = None
    for col in reversed(df.columns):
        if df[col].replace('', pd.NA).dropna().shape[0] > 0:
            last_valid_col = col
            break
    if last_valid_col:
        df = df.loc[:, :last_valid_col]

    return df

def extract_table_from_pdf(pdf_path):
    all_rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                all_rows.extend(table)

    if len(all_rows) < 8:
        raise ValueError("PDF table too short")

    # Main table: from row 8 onwards
    header = all_rows[7]
    df = pd.DataFrame(all_rows[8:], columns=header)

    # Metadata from rows 1–7 (2nd col as key, 8th col as value)
    for row in all_rows[:7]:
        if len(row) > 7 and row[1] and row[7]:
            key = row[1].strip()
            val = row[7].strip()
            if val.lower() not in ("", "none", "0", "0.00"):
                df[key] = val

    # Extract ARN and DATE from first 6 rows
    arn = ""
    date = ""
    for row in all_rows[:6]:
        lowered = [str(cell).lower() if cell else "" for cell in row]
        if 'arn' in lowered:
            i = lowered.index('arn')
            arn = row[i + 1] if len(row) > i + 1 else ""
        if 'date' in lowered:
            i = lowered.index('date')
            date = row[i + 1] if len(row) > i + 1 else ""
    df["ARN"] = arn
    df["DATE"] = date

    return clean_dataframe(df)


# Streamlit UI
st.title("📄 PDF to Excel Converter")

uploaded_files = st.file_uploader("Upload one or more PDF files", accept_multiple_files=True, type="pdf")

if uploaded_files:
    all_data = []
    for file in uploaded_files:
        df, name = process_pdf(file)
        if df is not None:
            all_data.append(df)
            st.success(f"✅ Processed {name}")
        else:
            st.warning(f"⚠️ Skipped {name} — Not enough rows")

    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"combined_output_{timestamp}.xlsx"
        combined_df.to_excel(file_name, index=False)

        with open(file_name, 'rb') as f:
            st.download_button("📥 Download Excel", f, file_name)
