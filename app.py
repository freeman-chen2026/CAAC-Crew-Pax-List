import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import re
from datetime import datetime, timedelta

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
    if not full_name:
        return ""
    parts = full_name.split()
    chinese_parts = [p for p in parts if re.search(r'[\u4e00-\u9fff]', p)]
    if chinese_parts:
        return " ".join(chinese_parts)
    else:
        return full_name

def parse_document_type(passport_no, doc_type):
    doc_type_str = str(doc_type).strip() if pd.notna(doc_type) else ""
    if doc_type_str:
        return doc_type_str
    pn = str(passport_no).strip() if pd.notna(passport_no) else ""
    pn = re.sub(r'\s+', '', pn)
    if re.match(r'^[0-9]{15}$', pn) or re.match(r'^[0-9]{17}[0-9Xx]$', pn):
        return "身份证"
    else:
        return "护照"

def safe_set_cell_value(ws, row, col, value):
    for merged_range in ws.merged_cells.ranges:
        if merged_range.min_row <= row <= merged_range.max_row and \
           merged_range.min_col <= col <= merged_range.max_col:
            ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = value
            return
    ws.cell(row=row, column=col).value = value

def get_value_right(ws, row, start_col):
    for col in range(start_col, start_col + 10):
        cell = ws.cell(row=row, column=col)
        if cell.value and str(cell.value).strip():
            return str(cell.value).strip()
    return ""

def parse_utc_to_beijing(utc_str, date_str):
    try:
        time_part = utc_str.replace('Z', '').strip()
        if len(time_part) == 4:
            hour = int(time_part[:2])
            minute = int(time_part[2:])
        elif len(time_part) == 3:
            hour = int(time_part[:1])
            minute = int(time_part[1:])
        else:
            return "0000"
        day = int(re.search(r'\d+', date_str).group()) if re.search(r'\d+', date_str) else 1
        month_map = {"Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
                     "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12}
        month_str = re.search(r'[A-Za-z]{3}', date_str).group() if re.search(r'[A-Za-z]{3}', date_str) else "Jan"
        month = month_map.get(month_str[:3], 1)
        year = 2026
        dt = datetime(year, month, day, hour, minute)
        dt_beijing = dt + timedelta(hours=8)
        return dt_beijing.strftime("%H%M")
    except:
        return "0000"

def parse_date_display(date_str):
    try:
        day = re.search(r'\d+', date_str).group()
        month_str = re.search(r'[A-Za-z]{3}', date_str).group()
        month_map = {"Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
                     "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12}
        month = month_map.get(month_str[:3], 1)
        return f"{month}月{int(day)}日"
    except:
        return date_str

# ---------- 解析GD单 ----------
def parse_general_declaration(file_bytes):
    wb = load_workbook(file_bytes)
    ws = wb.active
    data = {}
    for row in ws.iter_rows(min_row=1, max_row=20):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                val = cell.value.strip()
                if "OPERATOR:" in val:
                    data["operator"] = get_value_right(ws, cell.row, cell.column+1)
                elif "REG NO./FLT NO.:" in val:
                    reg_val = get_value_right(ws, cell.row, cell.column+1)
                    parts = reg_val.split()
                    data["reg"] = parts[0] if parts else reg_val
                    data["flt"] = parts[0] if parts else reg_val
                    if len(parts) > 1:
                        data["flt"] = parts[1]
                elif "AC TYPE:" in val:
                    data["ac_type"] = get_value_right(ws, cell.row, cell.column+1)
                elif "FROM:" in val:
                    data["from"] = get_value_right(ws, cell.row, cell.column+1)
                elif "TO:" in val:
                    data["to"] = get_value_right(ws, cell.row, cell.column+1)
                elif "DATE/TIME:" in val:
                    date_time = get_value_right(ws, cell.row, cell.column+1)
                    data["date_time"] = date_time
                    if date_time:
                        parts = date_time.split()
                        data["utc_time"] = parts[0] if len(parts) > 0 else ""
                        data["date_str"] = parts[1] if len(parts) > 1 else ""

    crew_data = []
    passenger_data = []
    section = None
    for row in ws.iter_rows(min_row=1):
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

# ---------- 辅助函数：复制行样式 ----------
def copy_row_style(ws, source_row, target_row, max_col=7):
    """复制source_row的样式到target_row"""
    for col in range(1, max_col+1):
        source_cell = ws.cell(row=source_row, column=col)
        target_cell = ws.cell(row=target_row, column=col)
        if source_cell.has_style:
            target_cell.font = source_cell.font.copy()
            target_cell.border = source_cell.border.copy()
            target_cell.fill = source_cell.fill.copy()
            target_cell.number_format = source_cell.number_format
            target_cell.protection = source_cell.protection.copy()
            target_cell.alignment = source_cell.alignment.copy()

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

    # ----- 0. 飞行目的 -----
    if not passenger_list:
        for row in ws.iter_rows(min_row=1, max_row=10):
            for cell in row:
                if cell.value and isinstance(cell.value, str) and "飞行目的" in cell.value:
                    target_row = cell.row + 1
                    safe_set_cell_value(ws, target_row, 2, "调机")
                    break
            else:
                continue
            break

    # ----- 1. 基础信息 -----
    info_row = None
    for row in ws.iter_rows(min_row=1, max_row=20):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                val = cell.value.strip()
                if val in ["机型", "注册号", "航班号", "航班行程"]:
                    info_row = cell.row
                    break
        if info_row:
            break

    if info_row:
        data_row = info_row + 1
        safe_set_cell_value(ws, data_row, 2, data.get("ac_type", ""))
        safe_set_cell_value(ws, data_row, 3, data.get("reg", ""))
        safe_set_cell_value(ws, data_row, 4, data.get("flt", ""))

        from_code = data.get("from", "")
        to_code = data.get("to", "")
        date_str = data.get("date_str", "")
        utc_time = data.get("utc_time", "")

        if date_str and from_code and to_code:
            date_display = parse_date_display(date_str)
            if utc_time:
                bj_time = parse_utc_to_beijing(utc_time, date_str)
            else:
                bj_time = "0000"
            route_display = f"{date_display} {from_code} {bj_time} XXXX {to_code}"
        else:
            route_display = f"{from_code}-{to_code}" if from_code and to_code else ""

        safe_set_cell_value(ws, data_row, 5, route_display)

    # ----- 2. 机组信息（支持自动插入乘务行） -----
    # 先找到各职位行
    captain_row = None
    copilot_row = None
    cabin_row = None   # 乘务行
    mechanic_row = None # 机务行

    for row in ws.iter_rows(min_row=1, max_row=50):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                val = cell.value.strip()
                if "机长" in val:
                    captain_row = cell.row
                elif "副驾驶" in val:
                    copilot_row = cell.row
                elif "乘务" in val and "信息" not in val:  # 避免匹配到“乘客信息”
                    cabin_row = cell.row
                elif "机务" in val:
                    mechanic_row = cell.row
        if captain_row and copilot_row and cabin_row and mechanic_row:
            break

    # 如果没有找到全部，停止处理机组
    if not (captain_row and copilot_row and cabin_row and mechanic_row):
        st.warning("未在模板中找到完整的机组行，请检查模板。")
    else:
        # 写入机长
        if len(crew_list) >= 1:
            crew = crew_list[0]
            safe_set_cell_value(ws, captain_row, 2, extract_chinese_name(crew["name"]))
            safe_set_cell_value(ws, captain_row, 3, crew.get("gender", ""))
            safe_set_cell_value(ws, captain_row, 4, crew.get("dob", ""))
            safe_set_cell_value(ws, captain_row, 5, crew.get("passport_no", ""))

        # 副驾驶
        if len(crew_list) >= 2:
            crew = crew_list[1]
            safe_set_cell_value(ws, copilot_row, 2, extract_chinese_name(crew["name"]))
            safe_set_cell_value(ws, copilot_row, 3, crew.get("gender", ""))
            safe_set_cell_value(ws, copilot_row, 4, crew.get("dob", ""))
            safe_set_cell_value(ws, copilot_row, 5, crew.get("passport_no", ""))

        # 乘务行（第3个机组）
        if len(crew_list) >= 3:
            crew = crew_list[2]
            safe_set_cell_value(ws, cabin_row, 1, "乘务")   # 职务列写“乘务”
            safe_set_cell_value(ws, cabin_row, 2, extract_chinese_name(crew["name"]))
            safe_set_cell_value(ws, cabin_row, 3, crew.get("gender", ""))
            safe_set_cell_value(ws, cabin_row, 4, crew.get("dob", ""))
            safe_set_cell_value(ws, cabin_row, 5, crew.get("passport_no", ""))

            # 检查是否有第4个机组（额外乘务或机务）
            if len(crew_list) >= 4:
                # 第4个机组
                extra_crew = crew_list[3]
                # 判断第4个是否为女性，若是则作为乘务插入新行，否则作为机务候选（但机务由后面逻辑处理）
                gender = str(extra_crew.get("gender", "")).strip()
                if gender in ["女", "Female", "F"]:
                    # 在乘务行下方插入新行
                    new_row = cabin_row + 1
                    ws.insert_rows(new_row)
                    # 复制乘务行的样式到新行
                    copy_row_style(ws, cabin_row, new_row, max_col=7)
                    # 填入数据
                    safe_set_cell_value(ws, new_row, 1, "乘务")
                    safe_set_cell_value(ws, new_row, 2, extract_chinese_name(extra_crew["name"]))
                    safe_set_cell_value(ws, new_row, 3, extra_crew.get("gender", ""))
                    safe_set_cell_value(ws, new_row, 4, extra_crew.get("dob", ""))
                    safe_set_cell_value(ws, new_row, 5, extra_crew.get("passport_no", ""))
                    # 注意：插入行后，机务行行号会+1，但我们之后会重新查找，所以暂时不处理
                # 如果第4个是男性，则可能作为机务（但后面会从第4个及以后找第一个男性）

        # 机务：从第4个机组开始（如果第4个已被插入，则从第5个开始）找第一个男性
        mechanic = None
        start_idx = 3  # 第4个机组索引（从0开始）
        # 如果第4个是女性且已被插入，则从第5个开始
        if len(crew_list) >= 4:
            gender4 = str(crew_list[3].get("gender", "")).strip()
            if gender4 in ["女", "Female", "F"]:
                start_idx = 4
        for i in range(start_idx, len(crew_list)):
            gender = str(crew_list[i].get("gender", "")).strip()
            if gender in ["男", "Male", "M"]:
                mechanic = crew_list[i]
                break
        if mechanic:
            # 重新查找机务行（因为可能插入了新行，行号变化）
            mechanic_row_new = None
            for row in ws.iter_rows(min_row=1, max_row=50):
                for cell in row:
                    if cell.value and isinstance(cell.value, str) and "机务" in cell.value:
                        mechanic_row_new = cell.row
                        break
                if mechanic_row_new:
                    break
            if mechanic_row_new:
                safe_set_cell_value(ws, mechanic_row_new, 2, extract_chinese_name(mechanic["name"]))
                safe_set_cell_value(ws, mechanic_row_new, 3, mechanic.get("gender", ""))
                safe_set_cell_value(ws, mechanic_row_new, 4, mechanic.get("dob", ""))
                safe_set_cell_value(ws, mechanic_row_new, 5, mechanic.get("passport_no", ""))

    # ----- 3. 乘客信息 -----
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
        for i, pax in enumerate(passenger_list):
            row_num = passenger_start_row + i
            for col in range(1, 7):
                safe_set_cell_value(ws, row_num, col, None)
            safe_set_cell_value(ws, row_num, 1, extract_chinese_name(pax["name"]))
            safe_set_cell_value(ws, row_num, 2, pax.get("gender", ""))
            safe_set_cell_value(ws, row_num, 3, pax.get("dob", ""))
            safe_set_cell_value(ws, row_num, 4, get_nation_name(pax.get("nationality", "")))
            doc_type = pax.get("doc_type", "")
            if pd.notna(doc_type) and str(doc_type).strip():
                safe_set_cell_value(ws, row_num, 5, str(doc_type).strip())
            else:
                safe_set_cell_value(ws, row_num, 5, parse_document_type(pax.get("passport_no", ""), ""))
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

        st.subheader("📋 提取的航班信息")
        from_code = data.get("from", "")
        to_code = data.get("to", "")
        date_str = data.get("date_str", "")
        utc_time = data.get("utc_time", "")
        if date_str and from_code and to_code:
            date_display = parse_date_display(date_str)
            if utc_time:
                bj_time = parse_utc_to_beijing(utc_time, date_str)
            else:
                bj_time = "0000"
            route_display = f"{date_display} {from_code} {bj_time} XXXX {to_code}"
            st.info(f"✈️ 航班行程将显示为：{route_display}")

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
