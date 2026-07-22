import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
import re

st.set_page_config(page_title="GD单 → 公务飞行计划信息备案表", layout="wide")
st.title("🛫 GD单 → 公务飞行计划信息备案表")
st.markdown("上传 GD单（General Declaration）Excel，自动提取机组和乘客信息，填入备案表模板。")

# ---------- 国籍映射 ----------
NATION_MAP = {
    "CHN": "中国", "HKG": "香港", "DEU": "德国", "USA": "美国", "GBR": "英国",
    "FRA": "法国", "RUS": "俄罗斯", "JPN": "日本", "KOR": "韩国", "SGP": "新加坡",
    "MYS": "马来西亚", "THA": "泰国", "VNM": "越南", "PHL": "菲律宾", "IDN": "印度尼西亚",
    "IND": "印度", "AUS": "澳大利亚", "CAN": "加拿大", "BRA": "巴西", "MEX": "墨西哥",
    "ZAF": "南非", "EGY": "埃及", "NGA": "尼日利亚", "KEN": "肯尼亚", "TZA": "坦桑尼亚",
    "ZWE": "津巴布韦", "NLD": "荷兰", "ITA": "意大利", "ESP": "西班牙", "PRT": "葡萄牙",
    "GRC": "希腊", "TUR": "土耳其", "SAU": "沙特阿拉伯", "ARE": "阿联酋", "ISR": "以色列",
    "IRN": "伊朗", "PAK": "巴基斯坦", "BGD": "孟加拉", "NPL": "尼泊尔", "LKA": "斯里兰卡",
    "MMR": "缅甸", "KHM": "柬埔寨", "LAO": "老挝", "MNG": "蒙古", "PRK": "朝鲜",
    "TWN": "台湾地区", "MAC": "澳门"
}

def get_nation_name(code):
    code = code.strip().upper()
    return NATION_MAP.get(code, code)

def extract_chinese_name(full_name):
    """优先提取中文姓名，若无则返回全名"""
    if not full_name:
        return ""
    parts = full_name.split()
    chinese_parts = [p for p in parts if re.search(r'[\u4e00-\u9fff]', p)]
    if chinese_parts:
        return " ".join(chinese_parts)
    else:
        return full_name

def parse_document_type(passport_no, doc_type):
    """根据证件号和证件类型判断：身份证 or 护照"""
    doc_type_str = str(doc_type).strip() if pd.notna(doc_type) else ""
    if "身份证" in doc_type_str or "居民身份证" in doc_type_str:
        return "身份证"
    if "通行证" in doc_type_str or "护照" in doc_type_str or "passport" in doc_type_str.lower():
        return "护照"
    pn = str(passport_no).strip() if pd.notna(passport_no) else ""
    pn = re.sub(r'\s+', '', pn)
    if re.match(r'^[0-9]{15}$', pn) or re.match(r'^[0-9]{17}[0-9Xx]$', pn):
        return "身份证"
    else:
        return "护照"

def safe_set_cell_value(ws, row, col, value):
    """
    安全设置单元格值，如果目标单元格属于合并区域，则设置合并区域的左上角。
    """
    # 检查是否在合并区域内
    for merged_range in ws.merged_cells.ranges:
        if merged_range.min_row <= row <= merged_range.max_row and \
           merged_range.min_col <= col <= merged_range.max_col:
            # 写入到合并区域的左上角
            ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = value
            return
    # 未合并，直接写入
    ws.cell(row=row, column=col).value = value

# ---------- 解析GD单 ----------
def parse_general_declaration(file_bytes):
    wb = load_workbook(file_bytes)
    ws = wb.active
    data = {}
    # 提取关键信息
    for row in ws.iter_rows(min_row=1, max_row=20):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                val = cell.value.strip()
                if "OPERATOR:" in val:
                    target = ws.cell(row=cell.row, column=cell.column+1)
                    data["operator"] = target.value.strip() if target.value else ""
                elif "REG NO./FLT NO.:" in val:
                    target = ws.cell(row=cell.row, column=cell.column+1)
                    reg_val = target.value.strip() if target.value else ""
                    parts = reg_val.split()
                    data["reg"] = parts[0] if parts else ""
                    data["flt"] = parts[1] if len(parts) > 1 else ""
                elif "AC TYPE:" in val:
                    target = ws.cell(row=cell.row, column=cell.column+1)
                    data["ac_type"] = target.value.strip() if target.value else ""
                elif "FROM:" in val:
                    target = ws.cell(row=cell.row, column=cell.column+1)
                    data["from"] = target.value.strip() if target.value else ""
                elif "TO:" in val:
                    target = ws.cell(row=cell.row, column=cell.column+1)
                    data["to"] = target.value.strip() if target.value else ""
                elif "DATE/TIME:" in val:
                    target = ws.cell(row=cell.row, column=cell.column+1)
                    data["date_time"] = target.value.strip() if target.value else ""

    crew_data = []
    passenger_data = []
    section = None
    for row in ws.iter_rows(min_row=1):
        # 检查是否进入某个表格
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                val = cell.value.strip()
                if "CREW MANIFEST" in val:
                    section = 'crew'
                    break
                elif "PASSENGER MANIFEST" in val:
                    section = 'passenger'
                    break
                elif "CARGO MANIFEST" in val or "DECLARATION OF HEALTH" in val:
                    section = None
                    break
        else:
            # 如果没找到关键词，按当前section处理
            pass
        if section == 'crew':
            first_cell = row[0]
            if first_cell.value and isinstance(first_cell.value, (int, float)):
                if len(row) >= 7:
                    name_cell = row[1]
                    if name_cell.value and isinstance(name_cell.value, str):
                        crew_data.append({
                            "name": name_cell.value.strip(),
                            "dob": row[2].value if row[2].value else "",
                            "gender": row[3].value if row[3].value else "",
                            "nationality": row[4].value if row[4].value else "",
                            "doc_type": row[5].value if row[5].value else "",
                            "passport_no": row[6].value if row[6].value else "",
                        })
        elif section == 'passenger':
            first_cell = row[0]
            if first_cell.value and isinstance(first_cell.value, (int, float)):
                if len(row) >= 7:
                    name_cell = row[1]
                    if name_cell.value and isinstance(name_cell.value, str):
                        passenger_data.append({
                            "name": name_cell.value.strip(),
                            "dob": row[2].value if row[2].value else "",
                            "gender": row[3].value if row[3].value else "",
                            "nationality": row[4].value if row[4].value else "",
                            "doc_type": row[5].value if row[5].value else "",
                            "passport_no": row[6].value if row[6].value else "",
                        })
    return data, crew_data, passenger_data

# ---------- 填充模板 ----------
def fill_template(template_bytes, data, crew_list, passenger_list):
    try:
        wb = load_workbook(template_bytes)
    except Exception as e:
        if "Bad magic number" in str(e) or "BadZipFile" in str(e):
            st.error("❌ 模板文件格式不正确。请确保模板为 **.xlsx** 格式（非 .xls）。")
            st.info("💡 解决方法：用 Excel 打开该模板，选择“另存为”，将文件类型选为 **Excel工作簿（.xlsx）**，然后重新上传。")
        raise e

    ws = wb.active

    # 1. 基础信息：通过关键词定位并写入
    for row in ws.iter_rows(min_row=1, max_row=30):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                val = cell.value.strip()
                if "机型" in val:
                    safe_set_cell_value(ws, cell.row, cell.column+1, data.get("ac_type", ""))
                elif "注册号" in val:
                    safe_set_cell_value(ws, cell.row, cell.column+1, data.get("reg", ""))
                elif "航班号" in val:
                    safe_set_cell_value(ws, cell.row, cell.column+1, data.get("flt", ""))
                elif "航班行程" in val:
                    from_ = data.get("from", "")
                    to_ = data.get("to", "")
                    safe_set_cell_value(ws, cell.row, cell.column+1, f"{from_}-{to_}" if from_ and to_ else "")

    # 2. 机组信息
    crew_positions = ["机长", "副驾驶", "乘务", "机务"]
    for idx, position in enumerate(crew_positions):
        if idx >= len(crew_list):
            break
        crew = crew_list[idx]
        for row in ws.iter_rows(min_row=1, max_row=50):
            for cell in row:
                if cell.value and isinstance(cell.value, str) and position in cell.value:
                    row_num = cell.row
                    safe_set_cell_value(ws, row_num, 2, extract_chinese_name(crew["name"]))
                    safe_set_cell_value(ws, row_num, 3, crew.get("gender", ""))
                    safe_set_cell_value(ws, row_num, 4, crew.get("dob", ""))
                    # 证件号码、执照号码、联系方式不填
                    break

    # 3. 乘客信息
    passenger_start_row = None
    for row in ws.iter_rows(min_row=1, max_row=100):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                val = cell.value.strip()
                if "姓名" in val and "性别" in val and "出生日期" in val:
                    passenger_start_row = cell.row + 1
                    break
        if passenger_start_row:
            break

    if passenger_start_row is None:
        for row in ws.iter_rows(min_row=1, max_row=100):
            for cell in row:
                if cell.value and isinstance(cell.value, str) and "乘客信息" in cell.value:
                    passenger_start_row = cell.row + 2
                    break
            if passenger_start_row:
                break

    if passenger_start_row:
        # 清空旧数据（保留表头，清空数据行）
        for i in range(20):  # 最多清空20行
            row_num = passenger_start_row + i
            for col in range(1, 7):
                safe_set_cell_value(ws, row_num, col, None)
        # 写入新数据
        for i, pax in enumerate(passenger_list):
            if i >= 20:
                break
            row_num = passenger_start_row + i
            safe_set_cell_value(ws, row_num, 1, extract_chinese_name(pax["name"]))
            safe_set_cell_value(ws, row_num, 2, pax.get("gender", ""))
            safe_set_cell_value(ws, row_num, 3, pax.get("dob", ""))
            safe_set_cell_value(ws, row_num, 4, get_nation_name(pax.get("nationality", "")))
            doc_type = parse_document_type(pax.get("passport_no", ""), pax.get("doc_type", ""))
            safe_set_cell_value(ws, row_num, 5, doc_type)
            safe_set_cell_value(ws, row_num, 6, pax.get("passport_no", ""))

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ---------- Streamlit UI ----------
st.subheader("📂 上传文件")
st.info("⚠️ 注意：模板文件必须是 **.xlsx** 格式（不是 .xls），否则无法打开。如果您的模板是 .xls，请用 Excel 另存为 .xlsx 后再上传。")

data_file = st.file_uploader("上传 GD单（General Declaration）Excel（.xlsx）", type=["xlsx"], key="data")
template_file = st.file_uploader("上传备案表模板 Excel（必须是 .xlsx）", type=["xlsx"], key="template")

if data_file and template_file:
    try:
        data, crew_list, passenger_list = parse_general_declaration(data_file)
        st.success(f"✅ 解析成功：机组 {len(crew_list)} 人，乘客 {len(passenger_list)} 人")
        if crew_list:
            st.write("提取的机组信息：", pd.DataFrame(crew_list))
        if passenger_list:
            st.write("提取的乘客信息（前5行）：", pd.DataFrame(passenger_list).head(5))

        result_bytes = fill_template(template_file, data, crew_list, passenger_list)

        st.download_button(
            label="⬇️ 下载填充后的备案表",
            data=result_bytes,
            file_name="公务飞行计划信息备案表_生成.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"❌ 处理失败：{e}")
        st.exception(e)
else:
    st.info("👆 请同时上传两个文件。")
