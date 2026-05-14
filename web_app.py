import streamlit as st
from excel_processor import process_uploaded_file

st.title("差旅报表自动转换 Agent")

uploaded_file = st.file_uploader("上传 Excel 文件", type="xlsx")

if uploaded_file is not None:
    with st.spinner("处理中..."):
        ticket_path, hotel_path = process_uploaded_file(uploaded_file)

    st.success("处理完成！")

    with open(ticket_path, "rb") as f:
        st.download_button("下载 国内机票", f, "国内机票_转换后.xlsx")
    with open(hotel_path, "rb") as f:
        st.download_button("下载 酒店", f, "酒店_转换后.xlsx")
