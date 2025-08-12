import json
import re
import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


def repair_json(content):
    """
    增强版JSON修复函数 - 特别加强嵌套对象处理
    """
    # 基础修复
    content = content.replace('\\"', '"')
    content = content.replace('“', '"').replace('”', '"')

    # 新增：修复双重引号包裹的嵌套JSON对象（加强版）
    content = re.sub(
        r'("\s*:\s*)"({[^{}]*})"',
        lambda m: f'{m.group(1)}{m.group(2)}',
        content,
        flags=re.DOTALL
    )

    # 新增：修复双重引号包裹的嵌套JSON数组
    content = re.sub(
        r'("\s*:\s*)"(\[[^\[\]]*\])"',
        lambda m: f'{m.group(1)}{m.group(2)}',
        content,
        flags=re.DOTALL
    )

    # 修复未加引号的键名
    content = re.sub(
        r'([{,]\s*)([a-zA-Z\u4e00-\u9fa5][^:\s]*)(\s*:)',
        r'\1"\2"\3',
        content
    )

    # 修复JavaScript语法
    content = re.sub(
        r'"([^"]+)"\s*\.repeat\s*\(\s*(\d+)\s*\)',
        lambda m: '"' + m.group(1) * int(m.group(2)) + '"',
        content
    )

    return content


def extract_json_from_md(md_file_path):
    """从Markdown中提取并修复JSON数据"""
    encodings = ['utf-8', 'gbk', 'gb2312', 'utf-16']

    for encoding in encodings:
        try:
            with open(md_file_path, 'r', encoding=encoding) as f:
                content = f.read()
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError("无法解码文件，尝试了所有支持的编码")

    json_match = re.search(r'(?s)\[\s*{.*}\s*\]', content)
    if not json_match:
        print("未找到有效的JSON数组")
        return None

    raw_json = json_match.group(0)
    repaired_json = repair_json(raw_json)

    try:
        return json.loads(repaired_json)
    except json.JSONDecodeError as e:
        print(f"JSON解析错误: {e}")
        print(f"错误位置: {e.pos}")
        print("错误上下文:")
        print(repaired_json[max(0, e.pos - 50):e.pos + 50])
        return None


def extract_tags(title):
    """从标题中提取标签（[]中的内容）"""
    tags = re.findall(r'\[(.*?)\]', title)
    return ", ".join(tags) if tags else ""


def transform_data(json_data, md_file_path):
    """转换数据为指定格式"""
    transformed = []
    md_filename = os.path.basename(md_file_path).replace('.md', '')

    for item in json_data:
        transformed.append({
            "标题": item.get("用例标题", ""),
            "目录": md_filename,
            "负责人": "郭光宁",  # 默认空
            "前置条件": item.get("前置条件", ""),
            "步骤描述": "\n".join(item.get("操作步骤", [])),
            "预期结果": "\n".join(
                item.get("预期结果", []) if isinstance(item.get("预期结果"), list) else [item.get("预期结果", "")]),
            "关联需求": "",  # 默认空
            "优先级": item.get("优先级", ""),
            "类型": "功能测试",  # 固定值
            "标签": extract_tags(item.get("用例标题", ""))
        })
    return transformed


def save_to_excel(data, excel_file_path):
    """保存数据到Excel"""
    df = pd.DataFrame(data)

    # 创建Excel文件
    wb = Workbook()
    ws = wb.active

    # 写入表头
    headers = ["标题", "目录", "负责人", "前置条件", "步骤描述", "预期结果", "关联需求", "优先级", "类型", "标签"]
    ws.append(headers)

    # 写入数据
    for item in data:
        row = [
            item["标题"],
            item["目录"],
            item["负责人"],
            item["前置条件"],
            item["步骤描述"],
            item["预期结果"],
            item["关联需求"],
            item["优先级"],
            item["类型"],
            item["标签"]
        ]
        ws.append(row)

    # 调整列宽
    for col in ws.columns:
        max_length = max(len(str(cell.value)) for cell in col)
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[col[0].column_letter].width = adjusted_width

    # 保存文件
    wb.save(excel_file_path)
    print(f"Excel文件已成功保存到: {excel_file_path}")


if __name__ == "__main__":
    # 输入输出文件路径
    md_file = r"C:\Users\Administrator\Desktop\滨州消防\测试用例\市本级主管部门增加审批节点.md"
    excel_file = r"C:\Users\Administrator\Desktop\滨州消防\测试用例\市本级主管部门增加审批节点.xlsx"

    print("开始处理文件...")
    json_data = extract_json_from_md(md_file)

    if json_data:
        print(f"成功解析 {len(json_data)} 条测试用例，正在转换Excel...")
        transformed_data = transform_data(json_data, md_file)
        save_to_excel(transformed_data, excel_file)
    else:
        print("转换失败，请检查以下问题：")
        print("1. 确保JSON部分完整且格式正确")
        print("2. 检查所有引号是否成对出现")
        print("3. 验证所有逗号和括号是否匹配")
        print("4. 可使用在线JSON验证工具检查格式")
        print("5. 可尝试手动修复后重新运行脚本")