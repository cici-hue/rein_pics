
# -*- coding: utf-8 -*-
import re
from io import BytesIO
from typing import Tuple
import streamlit as st
import pandas as pd
import pdfplumber
from PIL import Image
import pytesseract

st.set_page_config(page_title="Expense Report OCR", page_icon="📄", layout="centered")
st.title("📄 Expense Report OCR（只识别三项：报销单号 / QC name / 金额）")

def clean_name_to_english(name: str) -> str:
    name = re.sub(r"（.*?）", "", name)
    name = re.sub(r"[^A-Za-z ]+", "", name).strip()
    name = re.sub(r"\s+", " ", name)
    return name

def read_text_from_bytes(file_bytes: bytes, suffix: str) -> str:
    suffix = suffix.lower()
    if suffix == ".pdf":
        text_chunks = []
        with pdfplumber.open(BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                text_chunks.append(page.extract_text() or "")
        return "\n".join(text_chunks)
    elif suffix in (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"):
        img = Image.open(BytesIO(file_bytes))
        ocr_text = pytesseract.image_to_string(img, lang="eng+chi_sim", config="--psm 6")
        return ocr_text
    else:
        return ""

def extract_fields(text: str) -> Tuple[str, str, str]:
    m_no = re.search(r"(?:Expense Report(?: Number)?[:\s]+)(SHPC-[A-Z0-9]+)", text, re.IGNORECASE)
    report_no = m_no.group(1) if m_no else ""

    
m_name = re.search(r"Expense Report[:\s]+SHPC-[A-Z0-9]+,\s*([A-Za-z ]+)", text)
    if m_name:
        words = re.findall(r"[A-Za-z]+", m_name.group(1))
        qc_name = " ".join(words[:2])
    else:
        m_owner = re.search(r"(?:Report Owner|QC Name?)[:\s]+([A-Za-z ]+)", text)
        words = re.findall(r"[A-Za-z]+", m_owner.group(1)) if m_owner else []
        qc_name = " ".join(words[:2]) if words else ""


    m_amt = re.search(r"for\s*￥?\s*([0-9,]+\.[0-9]{2})", text, re.IGNORECASE)
    amount = m_amt.group(1) if m_amt else ""
    if not amount:
        m_amt2 = re.search(r"(?:Reimbursement|Total Amount)[:\s]+￥?\s*([0-9,]+\.[0-9]{2})", text, re.IGNORECASE)
        amount = m_amt2.group(1) if m_amt2 else ""

    return report_no, qc_name, amount

uploaded = st.file_uploader("上传 Expense Report 文件（PDF/图片）", type=["pdf","png","jpg","jpeg","bmp","tif","tiff"])
if uploaded is not None:
    suffix = "." + uploaded.name.split(".")[-1]
    text = read_text_from_bytes(uploaded.getvalue(), suffix)
    report_no, qc_name, amount = extract_fields(text)

    st.subheader("识别结果")
    st.text(f"Expense Report Number: {report_no}")
    st.text(f"QC name: {qc_name}")
    st.text(f"Amount: {amount}")

    # 下载：单行 Excel
    df = pd.DataFrame([{
        "Expense Report Number": report_no,
        "QC name": qc_name,
        "Amount": amount,
    }])
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    bio.seek(0)
    suggest_name = re.sub(r"\.(pdf|png|jpg|jpeg|bmp|tif|tiff)$", ".xlsx", uploaded.name, flags=re.IGNORECASE)
    st.download_button("⬇️ 下载识别结果（Excel）", bio.read(), file_name=suggest_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with st.expander("调试：原始文本预览"):
        st.code(text[:4000] + ("\n...\n" if len(text) > 4000 else ""), language="text")
else:
    st.info("请上传 PDF 或图片文件。")
