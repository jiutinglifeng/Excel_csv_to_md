import os
import glob
import pandas as pd
import csv

def excel_to_markdown(excel_file_path):
    """根据指定的规则将 Excel 文件转换为 Markdown 格式。"""
    try:
        df = pd.read_excel(excel_file_path)
    except FileNotFoundError:
        return f"错误：在 {excel_file_path} 未找到 Excel 文件"
    except Exception as e:
        return f"读取 Excel 文件 {excel_file_path} 时出错：{e}"

    markdown_content = ""
    headers = list(df.columns)
    first_column_header = headers[0] if headers else None # 处理空数据帧的情况

    for index, row in df.iterrows():
        first_column_value = str(row[first_column_header]) if first_column_header else "未命名项" # 处理第一列没有表头的情况
        markdown_content += f"# {first_column_value}\n\n"

        for i in range(1, len(headers)):
            header = str(headers[i])
            cell_value = str(row[header])
            markdown_content += f"### {header}\n\n"
            markdown_content += f"{cell_value}\n\n"

    return markdown_content

def csv_to_markdown(csv_file_path):
    """根据指定的规则将 CSV 文件转换为 Markdown 格式。"""
    try:
        with open(csv_file_path, 'r', encoding='utf-8') as csvfile: # 添加 encoding 以支持更广泛的字符集
            csv_reader = csv.reader(csvfile)
            headers = next(csv_reader, None) # 获取表头行，处理空 csv 文件的情况
            if headers is None:
                return "错误：CSV 文件为空或没有表头行。"

            markdown_content = ""
            first_column_header = headers[0] if headers else None

            for row in csv_reader:
                if not row: # 跳过空行
                    continue
                first_column_value = row[0] if row else "未命名项" # 处理行比预期短的情况
                markdown_content += f"# {first_column_value}\n\n"

                for i in range(1, len(headers)):
                    if i < len(row): # 检查行是否有足够的列
                        header = str(headers[i])
                        cell_value = str(row[i])
                        markdown_content += f"### {header}\n\n"
                        markdown_content += f"{cell_value}\n\n"
                    else:
                        # 如果需要，处理 CSV 行中缺少的值，目前跳过
                        pass

            return markdown_content

    except FileNotFoundError:
        return f"错误：在 {csv_file_path} 未找到 CSV 文件"
    except Exception as e:
        return f"读取 CSV 文件 {csv_file_path} 时出错：{e}"


def process_files_in_directory(directory_path):
    """处理给定目录中的所有 Excel 和 CSV 文件。"""
    excel_files = glob.glob(os.path.join(directory_path, "*.xlsx"))
    csv_files = glob.glob(os.path.join(directory_path, "*.csv"))
    all_files = excel_files + csv_files

    for file_path in all_files:
        file_name_without_ext = os.path.splitext(os.path.basename(file_path))[0]
        output_md_file_path = os.path.join(directory_path, file_name_without_ext + ".md")

        if file_path.endswith(".xlsx"):
            markdown_content = excel_to_markdown(file_path)
        elif file_path.endswith(".csv"):
            markdown_content = csv_to_markdown(file_path)
        else:
            continue # 应该不会到达这里，只是以防万一

        if markdown_content.startswith("错误："): # 检查转换时是否发生错误
            print(f"处理 {file_path} 时出错：{markdown_content[3:]}") # 打印错误消息，不带 "错误： " 前缀
            continue # 跳到下一个文件

        try:
            with open(output_md_file_path, 'w', encoding='utf-8') as md_file: # 添加 encoding 以支持更广泛的字符集
                md_file.write(markdown_content)
            print(f"成功将 '{file_path}' 转换为 '{output_md_file_path}'")
        except Exception as e:
            print(f"为 '{file_path}' 写入 Markdown 文件时出错：{e}")


if __name__ == "__main__":
    target_directory = "./"  # 替换为包含您的 Excel/CSV 文件的目录，或保留 "./" 表示当前目录
    process_files_in_directory(target_directory)
    print("批量转换过程完成。")