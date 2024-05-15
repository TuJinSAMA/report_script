import os
import glob
import win32com.client

from main import report_dir


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


def convert_doc_to_docx(word, doc_path, docx_path):


    try:
        # 打开 .doc 文件
        doc = word.Documents.Open(doc_path)

        # 将文件另存为 .docx 格式
        doc.SaveAs(docx_path, FileFormat=12)  # FileFormat=16 表示 .docx 格式

        print(f"转换完成: {doc_path} -> {docx_path}")
    except Exception as e:
        print(f"转换失败: {doc_path}")
        print(f"错误信息: {str(e)}")
    finally:
        # 关闭文档和 Word 应用程序
        doc.Close()



def exec_script(path):
    # convert_doc_to_docx(report_dir, report_dir)
    # 使用 glob 模块获取当前文件夹下所有 .doc 文件
    doc_files = glob.glob(os.path.join(path, "*.doc"))
    print(doc_files)
    print("当前文件夹下的 .doc 文件:")
    # 创建 Word 应用程序对象
    word = win32com.client.Dispatch("Word.Application")
    for file in doc_files:
        print(file)
        out_path = change_file_extension(file)
        convert_doc_to_docx(word, file, out_path)
    word.Quit()
    return 1


if __name__ == '__main__':
    # exec_script()
    for dir_name in os.listdir(report_dir):
        # 获取完整的文件夹路径
        dir_path = os.path.join(report_dir, dir_name)
        # 检查是否为文件夹
        if os.path.isdir(dir_path):
            print(f"找到文件夹: {dir_name}")
            exec_script(dir_path)