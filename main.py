import streamlit as st
import os
import tempfile
from planner.content_planner import ContentPlanner
from formatter.content_formatter import ContentFormatter
from writer.ppt_writer import PPTWriter
from llm.llm_api import LLMApi

def main():
    st.title("SmartPPT")
    
    # 侧边栏配置
    with st.sidebar:
        st.header("LLM配置")
        api_key = st.text_input("API密钥", type="password")
        base_url = st.text_input("API地址", value="https://api.siliconflow.cn/v1")
        model = st.selectbox("模型", ["Qwen/Qwen2.5-7B-Instruct"])


        if st.button("测试连接"):
            if api_key:
                llm_api = LLMApi(api_key, base_url, model)
                if llm_api.test_connection():
                    st.success("连接成功！")
                else:
                    st.error("连接失败，请检查配置")
            else:
                st.warning("请输入API密钥")
    
    # 主界面
    col1, col2 = st.columns(2)
    
    with col1:
        topic = st.text_input("请输入PPT主题：")
        num_pages = st.number_input("请输入页数：", min_value=1, max_value=20, value=5)
    
    with col2:
        style = st.selectbox("请选择PPT样式：", ["默认", "简约", "商务"])
        # use_template = st.checkbox("使用自定义模板")
        use_template = False
    
    # 模板上传
    template_file = None
    if use_template:
        st.subheader("上传PPT模板")
        uploaded_file = st.file_uploader(
            "选择PPT模板文件 (.pptx)", 
            type=['pptx'],
            help="上传一个PPT模板文件，系统将基于此模板生成最终PPT"
        )
        
        if uploaded_file is not None:
            # 保存上传的文件到临时目录
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                template_file = tmp_file.name
            
            st.success(f"模板上传成功: {uploaded_file.name}")
            
            # 显示模板信息
            try:
                writer = PPTWriter()
                template_info = writer.get_template_info(template_file)
                if "error" not in template_info:
                    st.info(f"模板包含 {template_info['total_layouts']} 种版式")
                    with st.expander("查看版式详情"):
                        for layout in template_info['layouts']:
                            st.write(f"版式 {layout['index']}: {layout['name']} (占位符: {layout['placeholders']})")
                else:
                    st.error(f"无法读取模板文件: {template_info['error']}")
            except Exception as e:
                st.error(f"无法读取模板文件: {e}")
    
    generate_btn = st.button("生成PPT")

    if generate_btn and topic:
        if not api_key:
            st.error("请先在侧边栏配置API密钥")
            return
            
        llm_api = LLMApi(api_key, base_url, model)
        planner = ContentPlanner(llm_api)
        formatter = ContentFormatter()
        writer = PPTWriter()

        with st.spinner("正在规划内容..."):
            planned_content = planner.plan_content(topic, num_pages)
            
            # 显示调试信息
            st.info(f"生成的内容页数: {len(planned_content)}")
            
        # 显示生成的内容预览
        st.subheader("内容预览")
        for i, page in enumerate(planned_content):
            with st.expander(f"第{i+1}页: {page['title']} (包含{len(page.get('points', []))}个论点)"):
                # 显示总结
                if "summary" in page:
                    st.markdown(f"**📋 总结:** {page['summary']}")
                    st.markdown("---")
                
                # 显示要点
                if "points" in page:
                    st.info(f"本页包含 {len(page['points'])} 个论点")
                    for j, point in enumerate(page["points"], 1):
                        if isinstance(point, dict) and "main_point" in point:
                            # 新格式：显示主要论点和支持事实
                            st.markdown(f"**{j}. {point['main_point']}**")
                            if "supporting_facts" in point:
                                for fact in point["supporting_facts"]:
                                    if isinstance(fact, dict) and "fact" in fact and "explanation" in fact:
                                        # 新格式：显示事实和说明，用冒号连接
                                        st.markdown(f"   • {fact['fact']}: {fact['explanation']}")
                                    else:
                                        # 旧格式：简单事实
                                        st.markdown(f"   • {fact}")
                        else:
                            # 旧格式：简单要点
                            st.write(f"{j}. {point}")
                        st.write("")  # 空行
        
        with st.spinner("正在格式化内容..."):
            formatted_content = formatter.format_content(planned_content)
        with st.spinner("正在生成PPT..."):
            output_path = f"{topic}_pptx.pptx"
            
            # 根据是否使用模板调用不同的生成方法
            if use_template and template_file:
                try:
                    writer.write_ppt_with_template(formatted_content, template_file, output_path, style)
                    st.success("基于模板生成PPT成功！")
                except Exception as e:
                    st.error(f"使用模板生成PPT失败: {e}")
                    st.info("正在使用默认样式重新生成...")
                    writer.write_ppt(formatted_content, output_path, style)
            else:
                writer.write_ppt(formatted_content, output_path, style)
                
        with open(output_path, "rb") as f:
            st.success("PPT生成成功！")
            st.download_button("下载PPT", f, file_name=output_path)
        
        # 清理临时文件
        if template_file and os.path.exists(template_file):
            os.unlink(template_file)
        if os.path.exists(output_path):
            os.remove(output_path)

if __name__ == "__main__":
    main() 