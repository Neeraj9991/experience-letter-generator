import streamlit as st
import pandas as pd
import os
from docx import Document
from io import BytesIO

# Function to replace placeholders in Word (case-insensitive)
def replace_placeholders(template, row):
    # Handle paragraphs
    for p in template.paragraphs:
        text = p.text
        for key, val in row.items():
            placeholder = f"{{{key}}}"
            text = text.replace(placeholder, str(val))
        # Clear existing runs and replace with new text
        if text != p.text:
            for r in p.runs:
                r.text = ""
            p.add_run(text)

    # Handle tables
    for table in template.tables:
        for row_cells in table.rows:
            for cell in row_cells.cells:
                for p in cell.paragraphs:
                    text = p.text
                    for key, val in row.items():
                        placeholder = f"{{{key}}}"
                        text = text.replace(placeholder, str(val))
                    if text != p.text:
                        for r in p.runs:
                            r.text = ""
                        p.add_run(text)
    return template

st.title("📄 Experience Letter Generator")

# Upload Word template
template_file = st.file_uploader("Upload Word Template (.docx)", type=["docx"])

# Upload Excel file
excel_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])

if template_file and excel_file:
    # Read Excel and clean headers
    df = pd.read_excel(excel_file)
    df.columns = df.columns.str.strip().str.lower()  # Normalize headers
    st.write("### Detected Columns:", df.columns.tolist())

    st.write("### Preview of Data")
    st.dataframe(df.head())

    output_folder = "Generated_Letters"
    os.makedirs(output_folder, exist_ok=True)

    if st.button("Generate Letters"):
        template_bytes = template_file.read()
        count = 0

        for idx, row in df.iterrows():
            doc = Document(BytesIO(template_bytes))

            # Build data dict (safe .get to avoid KeyError)
            data = {
                "name": row.get("name", ""),
                "empcode": row.get("empcode", ""),
                "designation": row.get("designation", ""),
                "department": row.get("department", ""),
                "doj": row.get("doj").strftime("%d-%m-%Y") if pd.notna(row.get("doj")) else "",
                "lwd": row.get("lwd").strftime("%d-%m-%Y") if pd.notna(row.get("lwd")) else ""
            }

            # Debug values
            st.write(f"Processing: {data}")

            # Replace placeholders
            doc = replace_placeholders(doc, data)

            # Save output file
            filename = f"{data['empcode']}_{data['name']}.docx".replace(" ", "_")
            filepath = os.path.join(output_folder, filename)
            doc.save(filepath)
            count += 1

        st.success(f"✅ {count} letters generated in folder: {output_folder}")
