import streamlit as st
import os
from dotenv import load_dotenv
from agent import CampusDocAgent

load_dotenv()

st.set_page_config(page_title="校园文书智能处理", layout="wide")
st.title("📄 校园行政文书智能处理 Agent")

agent = CampusDocAgent()

with st.form("input_form"):
    user_input = st.text_area(
        "请描述您的文书需求",
        placeholder="例如：帮我写一份实习报告，我在阿里巴巴实习，岗位是算法工程师...",
        height=150
    )
    col1, col2 = st.columns(2)
    with col1:
        name = st.text_input("姓名（可选）")
        student_id = st.text_input("学号（可选）")
    with col2:
        department = st.text_input("院系（可选）")
        other_info = st.text_area("其他信息（JSON格式，可选）", placeholder='{"unit":"腾讯", "position":"产品经理"}')

    submitted = st.form_submit_button("✨ 生成文书")

if submitted:
    if not user_input.strip():
        st.error("请输入需求描述")
    else:
        user_info = {}
        if name: user_info["name"] = name
        if student_id: user_info["student_id"] = student_id
        if department: user_info["department"] = department
        if other_info.strip():
            try:
                extra = eval(other_info)  # 简单解析，生产环境用json.loads
                user_info.update(extra)
            except:
                st.warning("额外信息格式错误，已忽略")

        with st.spinner("AI 正在处理您的文书，请稍候..."):
            result = agent.process(user_input, user_info)

        if result["status"] == "error":
            st.error(f"处理失败：{result['final_suggestions']}")
        else:
            st.success("文书生成成功！")
            tab1, tab2, tab3 = st.tabs(["📝 文书内容", "📋 合规性报告", "💾 下载文件"])

            with tab1:
                st.subheader(result["doc_type"])
                st.markdown(result["document"])

            with tab2:
                st.json(result["compliance_report"])
                if result["final_suggestions"]:
                    st.info("修改建议：\n" + "\n".join(result["final_suggestions"]))

            with tab3:
                if result.get("word_file"):
                    with open(result["word_file"], "rb") as f:
                        st.download_button("下载 Word 文件", f, file_name=os.path.basename(result["word_file"]))
                if result.get("markdown_file"):
                    with open(result["markdown_file"], "rb") as f:
                        st.download_button("下载 Markdown 文件", f, file_name=os.path.basename(result["markdown_file"]))

            st.caption(f"本次预估 Token 消耗：{result['token_used']}")