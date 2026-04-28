import json
import os
import shutil
import sys

import pandas as pd
from openpyxl import load_workbook
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

SPRAY_EXPECTED_COLUMNS = [
    'A', 'B', 'C', 'D', 'E',
    'F', 'G', 'H', 'I', 'J',
    'K', 'L', 'M', 'N', 'O',
]

COLUMN_MAPPING = {
    'directurl': 'Direct URL',
}

URL_COLUMN_CANDIDATES = ['e', 'url', 'direct url', 'directurl', '网址', '链接']
STATUS_COLUMN_CANDIDATES = ['j', 'status', 'status code', 'status_code', 'code', '状态码', '响应码', 'http code']


def extract_names(data):
    """从嵌套字典中提取所有'name'字段并拼接"""
    names = []
    try:
        if isinstance(data, str):
            data = json.loads(data)
        for key, value in data.items():
            if isinstance(value, dict) and 'name' in value:
                names.append(value['name'])
    except (json.JSONDecodeError, AttributeError, TypeError):
        return ""
    return ' | '.join(names)


def find_semantic_column(columns, candidates):
    normalized = {str(column).strip().lower(): column for column in columns}
    for candidate in candidates:
        match = normalized.get(candidate.lower())
        if match is not None:
            return match
    return None


def normalize_status_column(series):
    return pd.to_numeric(series, errors='coerce')


def normalize_url_values(values):
    urls = []
    seen = set()
    for value in values:
        if pd.isna(value):
            continue
        url = str(value).strip()
        if not url.startswith(('http://', 'https://')):
            continue
        if url in seen:
            continue
        seen.add(url)
        urls.append(url)
    return urls


def detect_spray_columns(df):
    url_col = df.columns[4] if len(df.columns) > 4 else None
    status_col = df.columns[9] if len(df.columns) > 9 else None
    semantic_url_col = find_semantic_column(df.columns, URL_COLUMN_CANDIDATES)
    semantic_status_col = find_semantic_column(df.columns, STATUS_COLUMN_CANDIDATES)
    return url_col, status_col, semantic_url_col, semantic_status_col


def filter_valid_urls(df, status_code_col='J', url_col='E'):
    """筛选状态码为200的URL记录"""
    if status_code_col not in df.columns or url_col not in df.columns:
        print(f"警告: 未找到列 {status_code_col} 或 {url_col}，跳过筛选")
        return df

    status_series = normalize_status_column(df[status_code_col])
    return df[(status_series == 200) & (df[url_col].notna())]


def beautify_spray_excel(file_path):
    """美化spray生成的Excel表格"""
    try:
        wb = load_workbook(file_path)
        ws = wb.active

        header_font = Font(bold=True, color="FFFFFF", size=12)
        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        data_font = Font(color="000000", size=11)
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.font = data_font
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='left', vertical='center')

        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    max_length = max(max_length, len(str(cell.value)))
                except Exception:
                    pass
            ws.column_dimensions[column].width = min(max_length + 2, 50)

        if 'J' in [get_column_letter(cell.column) for cell in ws[1]]:
            col_idx = [get_column_letter(cell.column) for cell in ws[1]].index('J') + 1
            status_range = f"{get_column_letter(col_idx)}2:{get_column_letter(col_idx)}{ws.max_row}"
            ws.conditional_formatting.add(
                status_range,
                ColorScaleRule(
                    start_type='min', start_color='FFC7CE',
                    mid_type='percentile', mid_value=50, mid_color='FFFFCC',
                    end_type='max', end_color='C6EFCE'
                )
            )

        wb.save(file_path)
        print(f"Spray Excel表格美化完成: {file_path}")
    except Exception as e:
        print(f"美化Spray Excel失败: {e}")


def beautify_ehole_excel(file_path):
    """深度美化ehole生成的Excel表格"""
    try:
        wb = load_workbook(file_path)
        ws = wb.active

        if ws.max_row <= 1:
            print(f"警告: ehole结果表格为空: {file_path}")
            return

        header_font = Font(bold=True, color="FFFFFF", size=14)
        header_fill = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
        data_font = Font(color="000000", size=12)
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.font = data_font
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    max_length = max(max_length, len(str(cell.value)))
                except Exception:
                    pass
            ws.column_dimensions[column].width = min(max_length + 2, 60)

        for col_name in ['URL', 'url', '网址']:
            if col_name in [cell.value for cell in ws[1]]:
                col_idx = [cell.value for cell in ws[1]].index(col_name) + 1
                for row in range(2, ws.max_row + 1):
                    cell = ws.cell(row=row, column=col_idx)
                    if cell.value and isinstance(cell.value, str) and cell.value.startswith(('http', 'https')):
                        cell.hyperlink = cell.value
                        cell.font = Font(color="0563C1", underline="single")

        for col_name in ['Risk', '风险等级', '危险程度']:
            if col_name in [cell.value for cell in ws[1]]:
                col_idx = [cell.value for cell in ws[1]].index(col_name) + 1
                risk_range = f"{get_column_letter(col_idx)}2:{get_column_letter(col_idx)}{ws.max_row}"
                ws.conditional_formatting.add(
                    risk_range,
                    CellIsRule(operator='containsText', formula=['"高"'], fill=PatternFill(bgColor='FFC7CE'), font=Font(color='9C0006'))
                )
                ws.conditional_formatting.add(
                    risk_range,
                    CellIsRule(operator='containsText', formula=['"中"'], fill=PatternFill(bgColor='FFEB9C'), font=Font(color='9C5700'))
                )
                ws.conditional_formatting.add(
                    risk_range,
                    CellIsRule(operator='containsText', formula=['"低"'], fill=PatternFill(bgColor='C6EFCE'), font=Font(color='006100'))
                )
                break

        if ws.max_row > 10:
            try:
                data_range = f"'{ws.title}'!$A$1:${get_column_letter(ws.max_column)}${ws.max_row}"
                pivot_ws = wb.create_sheet(title="数据透视表")
                pivot_ws['A1'] = "指纹识别结果统计"
                pivot_ws['A1'].font = Font(bold=True, size=16)

                from openpyxl.pivot.fields import PageField
                from openpyxl.pivot.table import PivotField, PivotTable

                pt = PivotTable(srcRange=data_range, dest=f"'{pivot_ws.title}'!$A$3", name="指纹统计")
                pt.addRow('A')
                if ws.max_column >= 2:
                    pt.addColumn('B')
                pt.addData('A', function='count')
                pivot_ws.add_pivot(pt)

                for row in pivot_ws.iter_rows(min_row=3, max_row=3):
                    for cell in row:
                        cell.font = Font(bold=True, color="FFFFFF")
                        cell.fill = PatternFill(start_color="4A86E8", end_color="4A86E8", fill_type="solid")

                if ws.max_column >= 3:
                    filter_values = list({ws.cell(row=i, column=3).value for i in range(2, ws.max_row + 1)})
                    dv = DataValidation(type="list", formula1='"{}"'.format(','.join([str(v) for v in filter_values if v])))
                    dv.add(pivot_ws['D1'])
                    pivot_ws.add_data_validation(dv)
                    pivot_ws['C1'] = "筛选:"
                    pivot_ws['C1'].font = Font(bold=True)

                print("已为ehole结果添加数据透视表")
            except Exception as e:
                print(f"创建数据透视表失败: {e}")

        summary_ws = wb.create_sheet(title="汇总信息")
        summary_ws['A1'] = "指纹识别结果汇总"
        summary_ws['A1'].font = Font(bold=True, size=16)
        summary_ws['A3'] = "总记录数:"
        summary_ws['B3'] = ws.max_row - 1

        wb.save(file_path)
        print(f"Ehole Excel表格深度美化完成: {file_path}")
    except Exception as e:
        print(f"美化Ehole Excel失败: {e}")
        raise


def process_data(input_file, output_file):
    """处理JSON输入文件，生成Excel和TXT输出"""
    try:
        file_ext = os.path.splitext(input_file)[1].lower()

        if file_ext == '.json':
            print(f"开始处理spray结果: {input_file}")
            data_list = []
            with open(input_file, 'r', encoding='utf-8') as f:
                for line in f:
                    try:
                        data = json.loads(line.strip())
                        data_list.append(data)
                    except json.JSONDecodeError:
                        print(f"警告: 无法解析JSON行: {line[:50]}...")

            if not data_list:
                print(f"错误: 文件 {input_file} 中没有有效JSON数据")
                return 1

            df = pd.DataFrame(data_list)
            df = df.rename(columns=COLUMN_MAPPING)

            if 'redirect_url' in df.columns:
                print("检测到'redirect_url'列，已删除")
                df = df.drop(columns=['redirect_url'])
            else:
                print("未检测到'redirect_url'列")

            if 'O' in df.columns:
                df['O'] = df['O'].apply(extract_names)

            existing_columns = df.columns.tolist()
            ordered_columns = []
            for col_letter in SPRAY_EXPECTED_COLUMNS:
                col_index = ord(col_letter.upper()) - 65
                if col_index < len(existing_columns):
                    ordered_columns.append(existing_columns[col_index])
                else:
                    ordered_columns.append(None)
            ordered_columns = [col for col in ordered_columns if col is not None]
            for col in existing_columns:
                if col not in ordered_columns:
                    ordered_columns.append(col)
            df = df[ordered_columns]

            url_col, status_col, semantic_url_col, semantic_status_col = detect_spray_columns(df)
            print(f"固定列识别 - URL列: {url_col}, 状态码列: {status_col}")
            if semantic_url_col or semantic_status_col:
                print(f"语义列识别 - URL列: {semantic_url_col}, 状态码列: {semantic_status_col}")

            valid_df = df
            if status_col is not None and url_col is not None:
                filtered_df = df[(normalize_status_column(df[status_col]) == 200) & (df[url_col].notna())]
                if len(filtered_df) > 0:
                    valid_df = filtered_df
                    print(f"按固定列筛选后保留 {len(valid_df)}/{len(df)} 条记录")
                elif semantic_status_col is not None and semantic_url_col is not None:
                    fallback_df = df[(normalize_status_column(df[semantic_status_col]) == 200) & (df[semantic_url_col].notna())]
                    if len(fallback_df) > 0:
                        valid_df = fallback_df
                        print(f"按语义列筛选后保留 {len(valid_df)}/{len(df)} 条记录")
                    else:
                        print(f"未筛出状态码200记录，保留所有 {len(df)} 条记录")
                else:
                    print(f"未筛出状态码200记录，保留所有 {len(df)} 条记录")
            else:
                print(f"未找到稳定的状态码列或URL列，保留所有 {len(df)} 条记录")

            valid_df.to_excel(output_file, index=False)
            print(f"Excel文件已保存: {output_file}")

            urls = []
            if len(valid_df.columns) >= 5:
                urls = normalize_url_values(valid_df.iloc[:, 4].tolist())
                print(f"优先使用第五列提取URL，得到 {len(urls)} 条")
            if not urls:
                fallback_url_col = semantic_url_col or find_semantic_column(valid_df.columns, URL_COLUMN_CANDIDATES)
                if fallback_url_col is not None:
                    urls = normalize_url_values(valid_df[fallback_url_col].tolist())
                    print(f"回退使用列 '{fallback_url_col}' 提取URL，得到 {len(urls)} 条")

            if urls:
                txt_output = os.path.splitext(output_file)[0] + ".txt"
                with open(txt_output, 'w', encoding='utf-8') as f:
                    f.write('\n'.join(urls))
                print(f"已提取 {len(urls)} 个URL保存到: {txt_output}")
            else:
                print("警告: 未找到有效URL")

            beautify_spray_excel(output_file)
            return 0

        if file_ext in ['.xlsx', '.xls']:
            print(f"开始美化ehole结果: {input_file}")
            if input_file != output_file:
                shutil.copy2(input_file, output_file)
            beautify_ehole_excel(output_file)
            return 0

        print(f"错误: 不支持的文件类型: {file_ext}")
        return 1
    except Exception as e:
        print(f"处理文件时出错: {e}")
        return 1


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("用法: python process_data.py <输入JSON/Excel文件> <输出Excel文件>")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]
    sys.exit(process_data(input_file, output_file))
