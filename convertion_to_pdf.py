import os
from pdf2docx import Converter
from docx import Document
import win32com.client
from pathlib import Path

def pdf_to_word(pdf_path, word_path):
    """
    将 PDF 转换为 Word 文件，支持中文字符。
    
    参数:
        pdf_path (str): 输入 PDF 文件路径
        word_path (str): 输出 Word 文件路径
    """
    try:
        # 创建 PDF 转换器
        cv = Converter(pdf_path)
        # 转换为 Word
        cv.convert(word_path, start=0, end=None)
        cv.close()
        print(f"成功将 {pdf_path} 转换为 {word_path}")
    except Exception as e:
        print(f"PDF 转 Word 失败: {str(e)}")

def word_to_pdf(word_path, pdf_path):
    """
    将 Word 转换为 PDF 文件，支持中文字符。
    
    参数:
        word_path (str): 输入 Word 文件路径
        pdf_path (str): 输出 PDF 文件路径
    """
    try:
        # 初始化 Word COM 对象
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # 隐藏 Word 窗口
        # 打开 Word 文档
        doc = word.Documents.Open(str(Path(word_path).absolute()))
        # 保存为 PDF
        doc.SaveAs(str(Path(pdf_path).absolute()), FileFormat=17)  # 17 表示 PDF 格式
        doc.Close()
        word.Quit()
        print(f"成功将 {word_path} 转换为 {pdf_path}")
    except Exception as e:
        print(f"Word 转 PDF 失败: {str(e)}")
        if 'word' in locals():
            word.Quit()

def main():
    # 用户输入文件路径
    input_file = input("请输入文件路径（PDF 或 Word 文件）: ").strip()
    output_dir = input("请输入输出文件夹路径（默认 'output'）: ").strip() or "output"

    # 确保输出目录存在
    output_dir = str(Path(output_dir).absolute())
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 检查输入文件是否存在
    if not os.path.exists(input_file):
        print(f"文件 {input_file} 不存在！")
        return

    # 获取文件扩展名（小写）
    file_ext = os.path.splitext(input_file)[1].lower()
    file_name = os.path.splitext(os.path.basename(input_file))[0]

    # 根据文件类型执行转换
    if file_ext == '.pdf':
        # PDF 转 Word
        word_path = os.path.join(output_dir, f"{file_name}.docx")
        pdf_to_word(input_file, word_path)
    elif file_ext in ['.docx', '.doc']:
        # Word 转 PDF
        pdf_path = os.path.join(output_dir, f"{file_name}.pdf")
        word_to_pdf(input_file, pdf_path)
    else:
        print(f"不支持的文件格式: {file_ext}，请提供 PDF 或 Word 文件！")

if __name__ == "__main__":
    main()