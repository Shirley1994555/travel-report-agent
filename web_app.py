import streamlit as st
from excel_processor import process_uploaded_file

st.title("差旅报表自动转换 Agent")
st.write("上传原始Excel报表")

uploaded_file = st.file_uploader("选择文件", type="xlsx")

if uploaded_file is not None:
    with st.spinner("正在自动处理，请稍候..."):
        # 直接把 uploaded_file 对象传进去，而不是路径
        ticket_path, hotel_path = process_uploaded_file(uploaded_file)

    st.success("处理完成！")

    with open(ticket_path, "rb") as f:
        st.download_button("下载国内机票报表", f, "国内机票_转换后.xlsx")
    with open(hotel_path, "rb") as f:
        st.download_button("下载酒店报表", f, "酒店_转换后.xlsx")
