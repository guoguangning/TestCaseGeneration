import json
import re
import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


def repair_json(content):
    """
    增强版JSON修复函数
    修复以下常见问题：
    1. 未转义的引号
    2. 中文引号
    3. 未加引号的键名
    4. 数组中的裸字符串
    5. JavaScript语法（如.repeat()）
    6. 缺失的分隔符
    """
    # 基础修复
    content = content.replace('\\"', '"')
    content = content.replace('“', '"').replace('”', '"')

    # 修复未加引号的键名（包括中文键名）
    content = re.sub(
        r'([{,]\s*)([a-zA-Z\u4e00-\u9fa5][^:\s]*)(\s*:)',
        r'\1"\2"\3',
        content
    )

    # 修复数组中的裸字符串（如 "值" 变成 ["值"]）
    content = re.sub(
        r'("\s*:\s*)([^"\[\]\s][^,\]]*)(\s*[,\]])',
        lambda m: f'{m.group(1)}"{m.group(2).strip()}"{m.group(3)}',
        content
    )

    # 修复JavaScript的.repeat()语法
    content = re.sub(
        r'"([^"]+)"\s*\.repeat\s*\(\s*(\d+)\s*\)',
        lambda m: '"' + m.group(1) * int(m.group(2)) + '"',
        content
    )

    # 修复缺失的分隔符（尝试自动补全）
    content = re.sub(
        r'("\s*:\s*)([^"{}\[\]\s][^,\n]*)(\s*\n)',
        lambda m: f'{m.group(1)}"{m.group(2).strip()}"{m.group(3)}',
        content
    )

    # 修复未闭合的引号（简单场景）
    content = re.sub(
        r'("\s*:\s*)([^"]+)(\s*[,}\]])',
        lambda m: f'{m.group(1)}"{m.group(2).strip()}"{m.group(3)}',
        content
    )

    return content


def extract_json_from_md(md_file_path):
    """从Markdown中提取并修复JSON数据（增强版）"""
    encodings = ['utf-8', 'gbk', 'gb2312', 'utf-16']

    for encoding in encodings:
        try:
            with open(md_file_path, 'r', encoding=encoding) as f:
                content = f.read()

            # 更灵活的JSON部分匹配
            json_match = re.search(r'(?s)\[\s*{.*}\s*\]', content) or \
                         re.search(r'(?s){\s*".*"\s*}', content)

            if not json_match:
                print(f"⚠️ 未在文件 {os.path.basename(md_file_path)} 中找到有效的JSON结构")
                return None

            raw_json = json_match.group(0)
            print(f"✅ 提取到原始JSON内容（长度: {len(raw_json)} 字符）")

            # 多阶段修复
            repaired_json = raw_json
            for attempt in range(3):  # 最多尝试修复3次
                try:
                    return json.loads(repaired_json)
                except json.JSONDecodeError as e:
                    if attempt == 2:  # 最后一次尝试失败后报错
                        error_context = repaired_json[max(0, e.pos - 50):e.pos + 50]
                        print(f"🔴 最终修复失败 {os.path.basename(md_file_path)}:")
                        print(f"错误类型: {e.msg}")
                        print(f"位置: 第{e.lineno}行第{e.colno}列 (字符{e.pos})")
                        print("错误上下文:")
                        print("..." + error_context.replace('\n', '\\n') + "...")
                        print("建议手动检查该部分JSON语法")
                    repaired_json = repair_json(repaired_json)

        except UnicodeDecodeError:
            continue

    raise ValueError(f"无法解码文件 {md_file_path}，尝试了所有支持的编码")


def extract_tags(title):
    """从标题中提取标签（[]中的内容）"""
    if not title or not isinstance(title, str):
        return ""
    tags = re.findall(r'\[(.*?)\]', title)
    return ", ".join(tags) if tags else ""


def transform_data(json_data, md_file_path):
    """转换数据为指定格式（增强容错性）"""
    transformed = []
    md_filename = os.path.basename(md_file_path).replace('.md', '').replace('-', '|')

    if not isinstance(json_data, list):
        json_data = [json_data]

    for item in json_data:
        if not isinstance(item, dict):
            continue

        # 处理预期结果（兼容字符串/列表/字典）
        expected_result = item.get("预期结果", "")
        if isinstance(expected_result, list):
            # 如果是列表，转换为字符串并去掉方括号
            expected_result = ", ".join(str(x) for x in expected_result)
        elif isinstance(expected_result, dict):
            # 如果是字典，转换为JSON字符串
            expected_result = json.dumps(expected_result, ensure_ascii=False)
        elif not isinstance(expected_result, str):
            # 其他类型转为字符串
            expected_result = str(expected_result)

        # 处理操作步骤（确保是列表）
        steps = item.get("操作步骤", [])
        if not isinstance(steps, list):
            steps = [steps] if steps else []

        transformed.append({
            "标题": item.get("用例标题", ""),
            "目录": md_filename,
            "负责人": item.get("负责人", "郭光宁"),
            "前置条件": item.get("前置条件", ""),
            "步骤描述": "\n".join([str(step) for step in steps]),
            "预期结果": expected_result,
            "关联需求": item.get("关联需求", ""),
            "优先级": item.get("优先级", ""),
            "类型": item.get("类型", "功能测试"),
            "标签": extract_tags(item.get("用例标题", ""))
        })
    return transformed


def save_to_excel(data, excel_file_path):
    """保存数据到Excel（增强格式处理）"""
    if not data:
        print(f"⚠️ 无有效数据，跳过生成 {excel_file_path}")
        return False

    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "测试用例"

        # 定义表头顺序
        headers = [
            "标题", "目录", "负责人", "前置条件",
            "步骤描述", "预期结果", "关联需求",
            "优先级", "类型", "标签"
        ]
        ws.append(headers)

        # 写入数据
        for item in data:
            row = [item.get(header, "") for header in headers]
            ws.append(row)

        # 自动调整列宽（限制最大50字符）
        for col in ws.columns:
            max_len = max(
                (len(str(cell.value)) for cell in col),
                default=0
            )
            adjusted_width = min(max_len + 2, 50)
            ws.column_dimensions[col[0].column_letter].width = adjusted_width

        # 保存文件
        wb.save(excel_file_path)
        print(f"✅ Excel文件已保存: {excel_file_path}")
        return True
    except Exception as e:
        print(f"🔴 保存Excel失败: {str(e)}")
        return False


def process_md_folder(input_folder, output_folder):
    """
    处理文件夹中的所有Markdown文件
    :param input_folder: 输入文件夹路径
    :param output_folder: 输出文件夹路径
    """
    if not os.path.exists(input_folder):
        raise ValueError(f"输入文件夹不存在: {input_folder}")

    os.makedirs(output_folder, exist_ok=True)
    processed_files = 0
    failed_files = 0

    print(f"\n🔧 开始处理文件夹: {input_folder}")
    for filename in sorted(os.listdir(input_folder)):
        if not filename.lower().endswith('.md'):
            continue

        md_file = os.path.join(input_folder, filename)
        excel_file = os.path.join(output_folder, f"{os.path.splitext(filename)[0]}.xlsx")

        print(f"\n📄 正在处理: {filename}")
        try:
            json_data = extract_json_from_md(md_file)
            if json_data:
                transformed_data = transform_data(json_data, md_file)
                if save_to_excel(transformed_data, excel_file):
                    processed_files += 1
                else:
                    failed_files += 1
            else:
                failed_files += 1
                print(f"⚠️ 跳过文件（无有效JSON数据）")
        except Exception as e:
            failed_files += 1
            print(f"🔴 处理失败: {str(e)}")
            continue

    print(f"\n🎉 处理完成！成功: {processed_files} 个, 失败: {failed_files} 个")
    return processed_files


if __name__ == "__main__":
    # 配置输入输出路径
    INPUT_FOLDER = r"C:\Users\Administrator\Desktop\监理\测试用例"
    OUTPUT_FOLDER = r"C:\Users\Administrator\Desktop\监理\测试用例\Excel输出"

    print("=" * 50)
    print("Markdown测试用例转换工具".center(40))
    print("=" * 50)

    try:
        success_count = process_md_folder(INPUT_FOLDER, OUTPUT_FOLDER)
        if success_count == 0:
            print("\n⚠️ 警告：没有成功转换任何文件！")
            print("可能原因：")
            print("1. 输入文件夹中没有.md文件")
            print("2. 所有文件都包含格式错误的JSON")
            print("3. 文件编码不被支持")
    except Exception as e:
        print(f"\n🔴 程序运行出错: {str(e)}")
    finally:
        input("\n按Enter键退出...")

