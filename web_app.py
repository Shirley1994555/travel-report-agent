import streamlit as st
import tempfile
from pathlib import Path
from excel_processor import process_uploaded_file

st.title("差旅报表自动转换 Agent")
uploaded = st.file_uploader("上传原始Excel报表", type="xlsx")

if uploaded:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded.read())
        tmp_path = Path(tmp.name)
    st.info("正在自动处理，请稍候...")
    ticket_path, hotel_path = process_uploaded_file(tmp_path)
    st.success("处理完成！请分别下载：")
    with open(ticket_path, "rb") as f:
        st.download_button("📥 下载 国内机票_转换后.xlsx", f, file_name=ticket_path.name)
    with open(hotel_path, "rb") as f:
        st.download_button("📥 下载 酒店_转换后.xlsx", f, file_name=hotel_path.name)