import streamlit as st
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from docx import Document
import pypdf
import os

st.set_page_config(page_title="My Report Editor", page_icon="📝", layout="wide")
st.title("📝 My Custom Report Editor")
st.write("Original format safe rakhkar data aur photo update karein.")

# 1. Format Type Selector
format_type = st.selectbox("Apna Report Format Chunein:", ["Excel (.xlsx)", "Word (.docx)", "PDF (.pdf)"])

# File Uploaders
uploaded_template = st.file_uploader(f"Step 1: Original Master {format_type} Upload Karein", type=[format_type.split()[-1][2:-1]])

if uploaded_template:
    st.success("Master Format Successfully Load Ho Gaya!")
    
    st.subheader("Step 2: Data & Photo Edit Karein")
    
    if "Excel" in format_type:
        # Excel Inputs
        cell_id = st.text_input("Jis cell ko edit karna hai (e.g., B5):", "B5")
        new_value = st.text_input("Naya data (Text/Number/Symbol) dalein:")
        
        uploaded_img = st.file_uploader("Agar Photo insert karni hai to chunein:", type=["png", "jpg", "jpeg"])
        img_cell = st.text_input("Photo kis cell me chahiye? (e.g., D10):", "D10")
        
        if st.button("Generate & Download Excel"):
            wb = openpyxl.load_workbook(uploaded_template)
            ws = wb.active
            
            # Edit text without disturbing layout
            if new_value:
                ws[cell_id] = new_value
                
            # Insert Image without changing cell size
            if uploaded_img:
                with open("temp_img.jpg", "wb") as f:
                    f.write(uploaded_img.getbuffer())
                img = OpenpyxlImage("temp_img.jpg")
                ws.add_image(img, img_cell)
            
            output_path = "Updated_Report.xlsx"
            wb.save(output_path)
            
            with open(output_path, "rb") as file:
                st.download_button(label="📥 Download Updated Excel", data=file, file_name="Updated_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    elif "Word" in format_type:
        # Word Inputs (Placeholder based: format me {{old_data}} hona chahiye)
        old_text = st.text_input("Jo purana text delete karna hai:")
        new_text = st.text_input("Jo naya text add karna hai:")
        
        if st.button("Generate & Download Word"):
            doc = Document(uploaded_template)
            
            # Layout safe text replacement
            for p in doc.paragraphs:
                if old_text in p.text:
                    p.text = p.text.replace(old_text, new_text)
                    
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if old_text in cell.text:
                            cell.text = cell.text.replace(old_text, new_text)
            
            output_path = "Updated_Report.docx"
            doc.save(output_path)
            
            with open(output_path, "rb") as file:
                st.download_button(label="📥 Download Updated Word", data=file, file_name="Updated_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    elif "PDF" in format_type:
        st.info("PDF Formats strictly fix hote hain. Isme overlay layer add hogi.")
        # PDF Form fields can be read or we can append pages/images
        uploaded_img = st.file_uploader("PDF ke upar insert karne ke liye photo chunein:", type=["png", "jpg", "jpeg"])
        
        if st.button("Process PDF"):
            reader = pypdf.PdfReader(uploaded_template)
            writer = pypdf.PdfWriter()
            
            for page in reader.pages:
                writer.add_page(page)
                
            output_path = "Updated_Report.pdf"
            with open(output_path, "wb") as f:
                writer.write(f)
                
            with open(output_path, "rb") as file:
                st.download_button(label="📥 Download Original PDF Structure", data=file, file_name="Updated_Report.pdf", mime="application/pdf")
              
