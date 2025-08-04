import os
import shutil
from main import report_dir
# from main_2025 import report_dir

def main(directory):
    # 遍历目录下的所有文件
    for filename in os.listdir(directory):
        # 检查文件是否为 .doc 文件
        if filename.endswith(".doc"):
            # 提取人名
            name = filename.split("_")[0].split("-")[-1]

            folder_path = os.path.join(directory, name)
            # 创建以人名命名的文件夹(如果不存在)
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)

            # 将文件移动到对应的文件夹中
            source = os.path.join(directory, filename)
            destination = os.path.join(directory, name, filename)
            shutil.move(source, destination)

if __name__ == '__main__':
    main(report_dir)
