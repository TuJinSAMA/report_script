import os
import glob
import win32com.client
from docx import Document

report_dir = r'E:\Temp\个体报告批量导出'


def change_file_extension(file_path):
    # 检查文件是否存在
    if not os.path.isfile(file_path):
        print(f"文件不存在: {file_path}")
        return None

    # 获取文件的目录路径和文件名
    directory, filename = os.path.split(file_path)

    # 分离文件名和后缀名
    name, extension = os.path.splitext(filename)

    # 检查文件后缀是否为 .doc
    if extension.lower() != ".doc":
        print(f"文件后缀不是 .doc: {file_path}")
        return None

    # 构建新的文件路径
    new_filename = name + ".docx"
    new_file_path = os.path.join(directory, new_filename)

    return new_file_path


def convert_doc_to_docx(doc_path, docx_path):
    # 创建 Word 应用程序对象
    word = win32com.client.Dispatch("Word.Application")

    try:
        # 打开 .doc 文件
        doc = word.Documents.Open(doc_path)

        # 将文件另存为 .docx 格式
        doc.SaveAs(docx_path, FileFormat=16)  # FileFormat=16 表示 .docx 格式

        print(f"转换完成: {doc_path} -> {docx_path}")
    except Exception as e:
        print(f"转换失败: {doc_path}")
        print(f"错误信息: {str(e)}")
    finally:
        # 关闭文档和 Word 应用程序
        doc.Close()
        word.Quit()


def read_docx(path):
    # 打开 Word 文档
    doc = Document(path)

    # 指定要匹配的表头
    target_header = '纬度一：总评'

    # 遍历文档中的所有表格
    for table in doc.tables:
        print(table)
        # 检查表格的第一行是否包含目标表头
        if target_header in table.rows[0].cells[0].text:
            # 找到目标表格，可以进行操作
            print(f"找到目标表格：{target_header}")

            # 示例操作：打印表格的行数和列数
            print(f"表格行数：{len(table.rows)}")
            print(f"表格列数：{len(table.columns)}")

            # 示例操作：遍历表格的每个单元格并打印其内容
            for row in table.rows:
                for cell in row.cells:
                    print(cell.text)

            # 示例操作：修改表格的内容
            # table.rows[1].cells[0].text = "新的内容"

            break  # 找到目标表格后退出循环

    # 保存修改后的文档
    # doc.save('example_modified.docx')


def exec_script():
    # convert_doc_to_docx(report_dir, report_dir)
    # 使用 glob 模块获取当前文件夹下所有 .doc 文件
    # doc_files = glob.glob(os.path.join(report_dir, "*.doc"))
    # print(doc_files)
    # print("当前文件夹下的 .doc 文件:")
    # for file in doc_files:
    #     print(file)
    #     out_path = change_file_extension(file)
    #     convert_doc_to_docx(file, out_path)

    docx_files = glob.glob(os.path.join(report_dir, "*.docx"))
    for file in docx_files:
        print(file)
        read_docx(file)

    return 1


if __name__ == '__main__':
    exec_script()
