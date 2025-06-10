import streamlit as st 
import pdfplumber
import pandas as pd
import os
import re
from datetime import datetime

# 📌 Clean the DataFrame (remove prefixes like "F ", "L ", "A123", etc.)
def clean_dataframe(df):
    def clean_cell(x):
        if pd.isnull(x):
            return x
        s = str(x).replace('\n', '').strip()
        if re.match(r'^[FLA]\s+', s):
            s = re.sub(r'^[FLA]\s+', '', s)
        elif re.match(r'^[FLA](\d+(\.\d+)?|\.\d+)$', s):
            s = re.sub(r'^[FLA]', '', s)
        return s

    df = df.applymap(clean_cell)

    # Drop columns that are mostly empty
    threshold = 0.9
    df = df.dropna(axis=1, thresh=int((1 - threshold) * len(df)))

    # Remove trailing empty columns
    last_valid_col = None
    for col in reversed(df.columns):
        if df[col].replace('', pd.NA).dropna().shape[0] > 0:
            last_valid_col = col
            break
    if last_valid_col:
        df = df.loc[:, :last_valid_col]

    return df

# 📌 Extract and process data from a single PDF
def extract_table_from_pdf(pdf_file):
    all_rows = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                all_rows.extend(table)

    if len(all_rows) < 8:
        raise ValueError("PDF table too short")

    header = all_rows[7]
    df = pd.DataFrame(all_rows[8:], columns=header)

    for row in all_rows[:7]:
        if len(row) > 7 and row[1] and row[7]:
            key = row[1].strip()
            val = row[7].strip()
            if val.lower() not in ("", "none", "0", "0.00"):
                df[key] = val

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

# 📌 Wrap it in a callable for Streamlit to use
def process_pdf(uploaded_file):
    try:
        df = extract_table_from_pdf(uploaded_file)
        return df, uploaded_file.name
    except Exception as e:
        st.error(f"❌ Error processing {uploaded_file.name}: {e}")
        return None, uploaded_file.name

# 📌 Streamlit UI
st.title("📄 PDF to Excel Converter")

uploaded_files = st.file_uploader("Upload one or more PDF files", accept_multiple_files=True, type="pdf")

if uploaded_files:
    all_data = []
    for file in uploaded_files:
        df, name = process_pdf(file)
        if df is not None:
            all_data.append(df)
            st.success(f"✅ Processed: {name}")
        else:
            st.warning(f"⚠️ Skipped: {name}")

    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"combined_output_{timestamp}.xlsx"
        combined_df.to_excel(output_filename, index=False)

        with open(output_filename, 'rb') as f:
            st.download_button("📥 Download Excel File", f, output_filename)
