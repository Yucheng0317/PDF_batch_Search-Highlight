import fitz  # PyMuPDF
import os
import logging
import msvcrt
from tkinter import Tk, filedialog

def upload_file(file_type, file_extension):
    """显示文件上传对话框"""
    root = Tk()
    root.withdraw()  # 隐藏主窗口
    file_path = filedialog.askopenfilename(
        title=f"请选择 {file_type} 文件",
        filetypes=[(file_type, f"*{file_extension}")]
    )
    root.destroy()
    return file_path

def save_file_as():
    """显示文件保存对话框"""
    root = Tk()
    root.withdraw()  # 隐藏主窗口
    file_path = filedialog.asksaveasfilename(
        title="保存高亮后的PDF",
        defaultextension=".pdf",
        filetypes=[("PDF 文件", "*.pdf")]
    )
    root.destroy()
    return file_path

def highlight_phrases_in_pdf(pdf_path, phrases, output_path, highlight_all=False):
    """在PDF中高亮短语"""
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        print(f"无法打开PDF文件: {e}")
        return None, None

    not_found_phrases = {}  # 记录未找到的短语
    found_count = {}  # 记录每个短语找到的次数

    print("开始在PDF中查找短语并高亮...")
    
    for phrase in phrases:
        found_count[phrase] = 0
        for page_num in range(doc.page_count):
            page = doc.load_page(page_num)
            text_instances = page.search_for(phrase)
            if text_instances:
                found_count[phrase] += len(text_instances)
                if highlight_all:
                    for inst in text_instances:
                        highlight = page.add_highlight_annot(inst)
                        highlight.update()
                else:
                    highlight = page.add_highlight_annot(text_instances[0])
                    highlight.update()
                    break
        if found_count[phrase] == 0:
            not_found_phrases[phrase] = 0

    try:
        # 保存高亮后的PDF
        print(f"正在保存高亮后的PDF到: {output_path}")
        doc.save(output_path)
        doc.close()
    except Exception as e:
        print(f"保存PDF文件时出错: {e}")
        return None, None

    print("短语高亮完成并保存PDF。")
    
    return found_count, not_found_phrases

def main():
    print('''\033[1;32m
=============
功能说明：
在PDF文档中批量查找文档里的短语并高亮
短语库支持.txt格式（每行一个，精确匹配，不分大小写）
=============\033[0m
\033[1;33m
高亮配置：
按数字1键--只高亮首次找到的文本
按数字2键--高亮每一个找到的文本\033[0m''')
    choice = input("请选择功能（1或2）：")
    if choice not in ['1', '2']:
        print("无效的选择。退出程序。")
        return

    highlight_all = choice == '2'

    print("上传 PDF 文件...")
    # 上传 PDF 文件
    pdf_path = upload_file("PDF", ".pdf")
    if not pdf_path:
        print("未选择 PDF 文件。")
        return

    print("上传短语库...")
    # 上传短语库
    phrases_path = upload_file("文本", ".txt")
    if not phrases_path:
        print("未选择短语库。")
        return

    print("读取短语库...")
    # 读取短语
    try:
        with open(phrases_path, 'r', encoding='utf-8') as f:
            phrases = [line.strip() for line in f if line.strip()]
    except Exception as e:
        print(f"读取短语库时出错: {e}")
        return

    if not phrases:
        print("短语库为空或未能读取到短语。")
        return

    print("选择保存路径...")
    # 选择保存路径
    output_path = save_file_as()
    if not output_path:
        print("未选择保存路径。")
        return

    # 设置日志文件路径
    log_file = output_path.replace(".pdf", "_log.txt")
    logging.basicConfig(filename=log_file, level=logging.INFO, format='%(message)s')

    print("开始查找并高亮短语...")
    # 查找并高亮短语
    found_count, not_found_phrases = highlight_phrases_in_pdf(pdf_path, phrases, output_path, highlight_all)

    if found_count is None or not_found_phrases is None:
        print("查找或保存过程中发生错误。")
        return

    print("生成日志文件...")
    # 记录查找情况
    try:
        with open(log_file, 'w', encoding='utf-8') as log:
            for phrase, count in found_count.items():
                log.write(f"查找短语：'{phrase}'，找到 {count} 处\n")
            if not_found_phrases:
                log.write("\n未找到的短语：\n")
                for phrase in not_found_phrases:
                    log.write(f"'{phrase}'\n")
        print(f"查找结果记录在: {log_file}")
    except Exception as e:
        print(f"生成日志文件时出错: {e}")
        return

    print(f"高亮处理后的PDF保存为: {output_path}")

    # 按任意键退出
    print("\033[1;33m已完成任务，按任意键退出...\033[0m")
    msvcrt.getch()

if __name__ == "__main__":
    main()
