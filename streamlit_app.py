import streamlit as st
import io
import zipfile
from PIL import Image
import pandas as pd
import tempfile
import subprocess
from docx2pdf import convert as word_to_pdf

# Try to import pdfplumber and show error if not found
try:
    import pdfplumber
except ImportError:
    st.error("pdfplumber 库未安装，请检查 requirements.txt 是否正确")

from pdf2docx import Converter

# Display title and description
st.title("📄 M2 PDF 转换助手")
st.write("支持：图片转 PDF、PDF 转图片、PDF 转 Excel、PDF 转 Word、Word 转 PDF、Excel 转 PDF")

# Select conversion type
option = st.selectbox("选择转换功能", [
    "图片转 PDF", "PDF 转图片", "PDF 转 Excel", "PDF 转 Word", "Word 转 PDF", "Excel 转 PDF"
])

# File uploader
file = st.file_uploader("上传文件", type=["png", "jpg", "jpeg", "pdf", "docx", "xlsx"])

if file and st.button("开始转换"):
    file_bytes = file.read()

    if option == "图片转 PDF":
        # Convert image to PDF
        image = Image.open(io.BytesIO(file_bytes))
        pdf_bytes = io.BytesIO()
        image.convert("RGB").save(pdf_bytes, format="PDF")
        st.download_button("📥 下载 PDF", pdf_bytes.getvalue(), "converted.pdf", "application/pdf")

    elif option == "PDF 转图片":
        # Use pdfplumber to convert PDF to image
        if pdfplumber:
            with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                images = []
                for page in pdf.pages:
                    img = page.to_image()
                    img_bytes = io.BytesIO()
                    img.original.save(img_bytes, format="PNG")
                    images.append(img_bytes.getvalue())

                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w") as z:
                    for i, img_data in enumerate(images):
                        z.writestr(f"page_{i+1}.png", img_data)
                st.download_button("📥 下载图片 ZIP", zip_buffer.getvalue(), "images.zip", "application/zip")
        else:
            st.error("PDF 转图片功能无法使用，缺少 pdfplumber 库。")

    elif option == "PDF 转 Excel":
        # PDF to Excel conversion (using temporary PDF file)
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_pdf:
            temp_pdf.write(file_bytes)
            temp_pdf_path = temp_pdf.name
        df = pd.read_csv(temp_pdf_path)  # 这里只是示例，实际 PDF 表格解析需要 Camelot
        excel_bytes = io.BytesIO()
        df.to_excel(excel_bytes, index=False)
        st.download_button("📥 下载 Excel", excel_bytes.getvalue(), "converted.xlsx", 
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    elif option == "PDF 转 Word":
        # PDF to Word conversion
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_pdf:
            temp_pdf.write(file_bytes)
            temp_pdf_path = temp_pdf.name
        docx_path = temp_pdf_path.replace(".pdf", ".docx")
        cv = Converter(temp_pdf_path)
        cv.convert(docx_path)
        cv.close()
        with open(docx_path, "rb") as f:
            st.download_button("📥 下载 Word", f.read(), "converted.docx", 
                               "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    elif option == "Word 转 PDF":
        # Word to PDF conversion
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as temp_docx:
            temp_docx.write(file_bytes)
            temp_docx_path = temp_docx.name
        word_to_pdf(temp_docx_path)
        pdf_path = temp_docx_path.replace(".docx", ".pdf")
        with open(pdf_path, "rb") as f:
            st.download_button("📥 下载 PDF", f.read(), "converted.pdf", "application/pdf")

    elif option == "Excel 转 PDF":
        # Excel to PDF conversion using LibreOffice
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp_excel:
            temp_excel.write(file_bytes)
            temp_excel_path = temp_excel.name
        output_dir = tempfile.mkdtemp()
        cmd = ['soffice', '--headless', '--convert-to', 'pdf', temp_excel_path, '--outdir', output_dir]
        subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        pdf_path = temp_excel_path.replace(".xlsx", ".pdf")
        with open(pdf_path, "rb") as f:
            st.download_button("📥 下载 PDF", f.read(), "converted.pdf", "application/pdf")

st.write("👨‍💻 开发：M2 PDF 转换助手 🚀")

