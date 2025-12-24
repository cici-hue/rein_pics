
# -*- coding: utf-8 -*-
"""
Expense Report OCR（批量版，手机可用）
- 上传 PDF/图片（可多选），识别3字段：
  1) Expense Report Number（报销单号）: 形如 SHPC-E253024
  2) QC name（仅保留两个英文单词）: 例如 George Zhang
  3) Amount（金额）: 例如 3,847.08
- 结果汇总为一张表并提供 Excel 下载
"""

import re
from io import BytesIO
from typing import Dict, List, Tuple
import streamlit as st
import pandas as pd
from PIL import Image, ImageOps, ImageFilter
import pytesseract
import pdfplumber

# ---------------- 页面配置（手机友好） ----------------
st.set_page_config(page_title="Expense OCR (Batch)", page_icon="📄", layout="centered")
st.markdown(
    """
    <style>
    .stButton>button {font-size: 16px; padding: 0.6rem 1rem;}
    .stDownloadButton>button {font-size: 16px; padding: 0.6rem 1rem;}
    </style>
    """,
    unsafe_allow_html=True
)
st.title("📄 Expense Report OCR（批量识别）")
st.caption("上传 PDF 或图片（可多选），自动识别：报销单号 / QC name（两个英文单词）/ 金额，并导出合并 Excel。")

# ---------------- 工具函数 ----------------
def _clean_name_english(name: str) -> str:
    """只保留英文名两个单词，去掉中文括注等"""
    name = re.sub(r"（.*?）", "", name)                  # 去中文括注
    words = re.findall(r"[A-Za-z]+", name)              # 只保留英文字母
    return " ".join(words[:2]).strip()                  # 最多两个单词

def _ocr_image(img: Image.Image) -> str:
    """图片 OCR（轻量预处理，提高手机拍照识别稳定性）"""
    # 修正 EXIF 方向、转灰度、增强对比度、轻微锐化
    img = ImageOps.exif_transpose(img)
    img = ImageOps.grayscale(img)
    img = ImageOps.autocontrast(img)
    img = img.filter(ImageFilter.SHARPEN)
    # 分辨率过小则放大到宽至少 1200px
    if img.width < 1200:
        ratio = 1200 / img.width
        img = img.resize((int(img.width * ratio), int(img.height * ratio)), Image.LANCZOS)
    # Tesseract OCR（中英文，psm 6 适合块状文本）
    return pytesseract.image_to_string(img, lang="eng+chi_sim", config="--psm 6")

def _read_text_from_bytes(file_bytes: bytes, suffix: str) -> str:
    """同时支持 PDF 与图片，返回整份文本"""
    suffix = suffix.lower()
    if suffix == ".pdf":
        chunks = []
        with pdfplumber.open(BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                chunks.append(page.extract_text() or "")
        return "\n".join(chunks)
    elif suffix in (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"):
        img = Image.open(BytesIO(file_bytes))
        return _ocr_image(img)
    else:
        return ""

def _extract_fields(text: str) -> Tuple[str, str, str]:
    """返回 (report_no, qc_name, amount) —— 含兜底规则"""
    # 报销单号：标题行或“Expense Report Number:”
    m_no = re.search(r"(?:Expense Report(?: Number)?[:\s]+)(SHPC-[A-Za-z0-9]+)", text, re.IGNORECASE)
    if not m_no:
        m_no = re.search(r"\b(SHPC-[A-Za-z0-9]+)\b", text)   # 再兜底
    report_no = m_no.group(1) if m_no else ""

    # QC name：标题行 "... SHPC-XXXXXX, NAME, on ..."
    m_name = re.search(r"Expense Report[:\s]+SHPC-[A-Za-z0-9]+,\s*(.+?)\s*,?\s*on\b", text, re.IGNORECASE)
    qc_name = _clean_name_english(m_name.group(1)) if m_name else ""
    if not qc_name:
        # 兜底：Report Owner / QC Name 标签
        m_owner = re.search(r"(?:Report Owner|QC Name?)[:\s]+(.+)", text, re.IGNORECASE)
        if m_owner:
            qc_name = _clean_name_english(m_owner.group(1))

    # 金额：标题行 "for ￥3,847.08"；兜底 Reimbursement/Total Amount
    m_amt = re.search(r"for\s*￥?\s*([0-9,]+\.[0-9]{2})", text, re.IGNORECASE)
    amount = m_amt.group(1) if m_amt else ""
    if not amount:
        m_amt2 = re.search(r"(?:Reimbursement|Total Amount)[:\s]+(?:CNY|￥)?\s*([0-9,]+\.[0-9]{2})", text, re.IGNORECASE)
        amount = m_amt2.group(1) if m_amt2 else ""

    return report_no, qc_name, amount

# ---------------- 上传与批量处理 ----------------
uploads = st.file_uploader(
    "上传 Expense Report（PDF/图片，可多选；手机可直接拍照或选相册）",
    type=["pdf", "png", "jpg", "jpeg", "bmp", "tif", "tiff"],
    accept_multiple_files=True
)

if uploads:
    rows: List[Dict[str, str]] = []
    with st.status("正在识别…", expanded=False) as status:
        for f in uploads:
            suffix = "." + f.name.split(".")[-1]
            text = _read_text_from_bytes(f.getvalue(), suffix)
            report_no, qc_name, amount = _extract_fields(text)
            rows.append({
                "Expense Report Number": report_no,
                "QC name": qc_name,
                "Amount": amount,
            })
        status.update(label="识别完成", state="complete")

    # 预览
    st.subheader("识别结果预览（合并表）")
    df = pd.DataFrame(rows, columns=["Expense Report Number", "QC name", "Amount"])
    st.dataframe(df, use_container_width=True)

    # 下载合并 Excel
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    bio.seek(0)
    st.download_button(
        "⬇️ 下载合并 Excel",
        data=bio.read(),
        file_name="Expense_OCR_All.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # 调试：原始文本片段（可折叠）
    with st.expander("调试：原始文本片段（每份文件取前 2,000 字）"):
        for i, f in enumerate(uploads, start=1):
            suffix = "." + f.name.split(".")[-1]
            text = _read_text_from_bytes(f.getvalue(), suffix)
            st.markdown(f"**文件 {i}: {f.name}**")
            st.code(text[:2000] + ("\n...\n" if len(text) > 2000 else ""), language="text")

else:
    st.info("请选择一份或多份文件进行上传。")
