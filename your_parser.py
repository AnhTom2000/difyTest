# your_parser.py

import io
import zipfile
import pandas as pd

def _parse_txt_content(file_content_str: str, source_name: str, all_commands: list):
    """
    一个内部辅助函数，用于解析TXT文件内容。
    """
    commands = []
    in_command_section = False
    # 使用StringIO来逐行读取字符串，更稳定
    for line in io.StringIO(file_content_str).readlines():
        line_stripped = line.strip()
        if '操作指令:' in line_stripped:
            in_command_section = True
            continue
        # 只有在'操作指令:'之后并且非空的行才会被添加
        if in_command_section and line_stripped:
            commands.append(line_stripped)
    if commands:
        all_commands.append({
            "source_file": source_name,
            "attachment_type": "text/plain",
            "commands": commands
        })

def _parse_excel_content(df: pd.DataFrame, source_name: str, all_commands: list):
    """
    一个内部辅助函数，用于解析DataFrame（来自Excel文件）。
    它能同时处理两种格式的Excel模板。
    """
    # 检查是否为字符命令类模板
    if '操作指令' in df.columns:
        for _, row in df.iterrows():
            hostname = row.get('hostname', 'N/A')
            # 确保指令内容是字符串，并处理空值
            cmds_raw = str(row['操作指令']) if pd.notna(row['操作指令']) else ""
            cmds = cmds_raw.strip().split('\n')
            # 清理掉空的命令
            cmds_cleaned = [cmd.strip() for cmd in cmds if cmd.strip()]
            if cmds_cleaned:
                all_commands.append({
                    "source_file": source_name,
                    "attachment_type": "excel_char_based",
                    "context": f"主机: {hostname}",
                    "commands": cmds_cleaned
                })
    # 检查是否为图形网管类模板
    elif '操作指令编码' in df.columns:
        for _, row in df.iterrows():
            platform = row.get('网管平台名称', 'N/A')
            op_code = str(row['操作指令编码']) if pd.notna(row['操作指令编码']) else ""
            if op_code and op_code.strip():
                all_commands.append({
                    "source_file": source_name,
                    "attachment_type": "excel_gui_based",
                    "context": f"平台: {platform}",
                    "commands": [op_code.strip()] # 编码也视为一种命令
                })

def analyze_docx_attachments(docx_binary_data: bytes) -> dict:
    """
    接收DOCX文件的二进制数据，提取所有内嵌的附件并解析命令。
    返回一个包含命令和分析结果的结构化字典。
    """
    all_commands = []
    try:
        # 将二进制数据加载到内存中的文件流
        docx_stream = io.BytesIO(docx_binary_data)
        
        # .docx文件本质是一个ZIP压缩包，我们用zipfile库来读取它
        with zipfile.ZipFile(docx_stream, 'r') as docx_zip:
            # 遍历压缩包中的所有成员
            for member in docx_zip.infolist():
                # 内嵌的附件对象通常存储在 'word/embeddings/' 目录下
                if member.filename.startswith('word/embeddings/'):
                    attachment_name = member.filename.split('/')[-1]
                    
                    # 读取附件的二进制数据
                    attachment_data = docx_zip.read(member.filename)
                    attachment_stream = io.BytesIO(attachment_data)
                    
                    # --- 开始尝试解析附件 ---
                    try:
                        # 1. 首先尝试作为Excel文件进行解析
                        # 使用一个流的副本来解析，以免影响后续操作
                        stream_copy_for_excel = io.BytesIO(attachment_stream.getvalue())
                        # pandas能自动识别.xls和.xlsx
                        df = pd.read_excel(stream_copy_for_excel, engine=None) 
                        _parse_excel_content(df, attachment_name, all_commands)
                        # 如果作为Excel解析成功，则跳过后续的TXT解析，继续处理下一个附件
                        continue
                    except Exception:
                        # 2. 如果作为Excel解析失败，则尝试作为TXT文件进行解析
                        try:
                            # 使用原始流的数据进行解码
                            txt_content = attachment_stream.getvalue().decode('utf-8', errors='ignore')
                            # 通过关键词简单判断是否是我们需要的TXT文件
                            if "hostname:" in txt_content or "操作指令:" in txt_content:
                                 _parse_txt_content(txt_content, attachment_name, all_commands)
                        except Exception:
                            # 如果两种方式都解析失败，说明它可能是图片或其他不支持的对象，我们选择静默忽略。
                            pass

    except Exception as e:
        # 如果在处理DOCX文件本身时就出错，则返回错误信息
        return {"status": "error", "message": f"解析DOCX文件结构失败: {str(e)}"}

    # 在这里可以增加知识库比对等更复杂的业务逻辑。
    # 例如，遍历 all_commands 列表，将每个命令与您的知识库进行比对。

    # 返回最终的结构化结果
    return {
        "status": "success",
        "analysis_results": all_commands
    }