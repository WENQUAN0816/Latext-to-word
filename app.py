import streamlit as st
import subprocess
import tempfile
import os

st.title("LaTeX 转 Word (Pandoc) 在线工具")

uploaded_file = st.file_uploader("选择 LaTeX 文件（.tex）", type=['tex'])
reference_file = st.file_uploader("选择 Word 模板 reference.docx（可选）", type=['docx'])

output_format = st.selectbox("输出格式", ['docx', 'odt'])
number_sections = st.checkbox("标题自动编号", value=True)
fix_syntax = st.checkbox("自动修复语法（简单实现，仅修复部分错误）", value=True)
extra_options = st.text_input("附加 Pandoc 参数", "")

if st.button("开始转换"):
    if not uploaded_file:
        st.error("请先上传 LaTeX 文件")
        st.stop()

    # 临时文件保存输入
    with tempfile.NamedTemporaryFile(suffix=".tex", delete=False) as tf:
        tf.write(uploaded_file.getvalue())
        tex_path = tf.name

    # 自动修复语法（例子，仅补齐 \end{document}）
    if fix_syntax:
        with open(tex_path, 'r+', encoding='utf-8', errors='replace') as f:
            content = f.read()
            if '\\end{document}' not in content:
                content += '\n\\end{document}\n'
                f.seek(0)
                f.write(content)
                f.truncate()
        st.info("已自动补齐 \\end{document}")

    # 生成输出文件路径
    output_path = tex_path.replace('.tex', '.' + output_format)

    # 组装 Pandoc 命令
    cmd = ['pandoc', tex_path, '-o', output_path]
    if reference_file:
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as ref:
            ref.write(reference_file.getvalue())
            ref_path = ref.name
        cmd.extend(['--reference-doc', ref_path])
    if number_sections:
        cmd.append('--number-sections')
    if extra_options.strip():
        cmd.extend(extra_options.strip().split())

    # 运行 Pandoc
    st.write("执行命令: ", " ".join(cmd))
    result = subprocess.run(cmd, capture_output=True, text=True)
    st.text_area("Pandoc 输出日志", value=result.stdout + "\n" + result.stderr, height=200)

    # 成功则可下载
    if os.path.exists(output_path):
        with open(output_path, 'rb') as f:
            st.download_button('下载 Word 文件', f, file_name='converted.' + output_format)
    else:
        st.error("转换失败，请检查 LaTeX 源文件格式或 Pandoc 日志")

