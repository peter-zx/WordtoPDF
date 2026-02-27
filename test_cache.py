import streamlit as st

st.title("🧪 测试页面")
st.write("如果你看到这个页面，说明 Streamlit 正常工作")

st.info("💡 **重要提示**: 这是测试提示信息")

st.markdown("""
<div style='background-color: #fff3cd; padding: 15px; border-radius: 5px; margin: 10px 0; border-left: 4px solid #ffc107;'>
<strong>🔴 必须上传 ZIP 格式压缩包!</strong><br>
请将包含 Word 文件的整个文件夹压缩为 ZIP 格式后上传。<br>
系统会保持原始文件夹结构显示所有文件。
</div>
""", unsafe_allow_html=True)

st.success("✅ 如果看到黄色警告框，说明 HTML 渲染正常")
