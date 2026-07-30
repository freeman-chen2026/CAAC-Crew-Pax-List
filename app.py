import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
import re
from datetime import datetime, timedelta

st.set_page_config(page_title="GD单 → 公务飞行计划信息备案表", layout="wide")
st.title("🛫 GD单 → 公务飞行计划信息备案表")
st.markdown("上传 GD单 和模板，自动生成备案表（联系方式及执照号码已内置）。")

# ---------- 内置联系方式映射 ----------
BUILTIN_CONTACT_MAP = {
    "庚凡": "139 2463 9747",
    "张永一": "139 0125 9544",
    "梅峰": "135 0967 8127",
    "王斌": "139 2527 2867",
    "王少雄": "186 8387 9841",
    "苗旺旺": "138 1871 5251",
    "赵岩松": "186 1161 8385",
    "Bruce Roderick, WAINES": "186 6532 9796",
    "Oliver Viktor, RACZ": "186 1197 3165",
    "Yiftah RAUCH": "186 1045 0563",
    "尤欣": "139 1608 5072",
    "李亚民": "133 6632 0878",
    "赵镭": "138 0883 9660",
    "彭罡": "186 1263 1888",
    "Wan Leung WU": "132 6695 8816",
    "Kwan Leung WU": "132 6695 8816",
    "吴鹏": "136 1110 5901",
    "刘汇川": "188 5827 2791",
    "于龙飞": "156 5288 0812",
    "林帅": "138 1156 6711",
    "蔡国俊": "157 1220 8304",
    "李庆宏": "135 0909 0503",
    "宋炜": "136 3256 5565",
    "昝昭君": "182 9570 0579",
    "孙浩": "136 7012 1990",
    "朱正宇": "189 8335 3697",
    "金尚明": "136 7113 8047",
    "张哲": "139 0247 5026",
    "王莹": "159 1009 9069",
    "赵婷婷": "138 2883 3162",
    "王凯珮": "852 68588410",
    "范蕾蕾": "182 1000 6866",
    "危慧": "152 1349 1328",
    "廉卓群": "133 5632 3949",
    "张佳妮": "136 6169 9966",
    "樊婉程": "186 2017 4817",
    "姚艳阁": "156 0127 9399",
    "李潇恩": "158 0599 1600",
    "赖小燕": "60 1239 05520",
    "Siau Mui LAI": "60 1239 05520",
    "花佩": "186 2631 0634",
    "丁燕栒": "135 6035 3829",
    "何静文": "852 6421 0994",
    "蔡雨桐": "852 6426 7445",
    "茅邂文": "152 5181 7375",
    "张欢乐": "186 1652 1529",
    "詹佩佩": "137 1440 5925",
    "周丽欢": "152 5710 6140",
    "翁英": "130 6785 2000",
    "李卉妍": "138 5142 0321",
    "卢江": "158 0045 6521",
    "AYA, MUGURUMA": "81 8071140700",
    "梁广煜": "137 9428 7177",
    "李园": "139 1178 3914",
    "孔铮": "139 1085 3981",
    "高峰": "135 8186 9017",
    "李海": "136 8131 8388",
    "李阳": "133 6603 6567",
    "程佳俊": "134 8010 3029",
    "陈居瑜": "158 8962 6660",
    "林生": "159 2162 9406",
    "谢依椿": "159 8942 4501",
    "廖关荣": "181 0755 9103",
    "林峰": "135 2260 7955",
    "丘东": "136 0044 6505",
    "王庆辉": "134 8013 9352",
    "姜磊": "135 1006 5318",
    "黄彦杰": "132 1157 2184",
    "王珍": "198 6662 9312",
    "赵国庆": "155 8849 2975",
    "俞凯": "130 0579 0326",
    "王军": "853 62666900",
    "张德桃": "130 2883 6410",
    "江焰辉": "136 3148 0927",
    "孙龙": "156 9558 0691",
    "梁平": "138 2750 6225",
    "冯仁毫": "185 2028 6463",
    "焦石军": "139 2388 3525",
    "万子辰": "177 7005 7193",
    "卓辉": "157 7070 8632",
    "万虹波": "133 1297 9906",
    "王晟磊": "150 2689 7493",
    "杨杰": "186 1694 8903",
    "林帅": "138 1156 6711",
    "苏志斌": "159 0150 7150",
    "孙辉": "139 1626 9572",
    "孟周聪": "135 6434 5029",
    "熊立凌": "135 3821 6276",
    "赵康": "191 6764 6172",
    "翟征宇": "134 1448 9793",
}

# ---------- 内置执照号码映射 ----------
BUILTIN_LICENSE_MAP = {
    "吴鹏": "130103197602102115",
    "刘汇川": "330103199003191618",
    "于龙飞": "ZN00915",
    "林帅": "110227198601130015",
}

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

def normalize_name(name):
    """标准化姓名：去除中文、去除逗号、压缩多余空格为单个空格，转小写"""
    if not name:
        return ""
    name = re.sub(r'[\u4e00-\u9fff]+', '', name)
    name = re.sub(r'[,\s]+', ' ', name).strip()
    return name.lower()

def find_contact(crew_name):
    """根据机组姓名从内置映射中查找联系方式"""
    if not crew_name or not BUILTIN_CONTACT_MAP:
        return ""
    chinese = extract_chinese_name(crew_name)
    if chinese and chinese in BUILTIN_CONTACT_MAP:
        return BUILTIN_CONTACT_MAP[chinese]
    norm = normalize_name(crew_name)
    if not norm:
        return ""
    for key, val in BUILTIN_CONTACT_MAP.items():
        if normalize_name(key) == norm:
            return val
    return ""

def find_license(crew_name):
    """根据机组姓名从内置映射中查找执照号码"""
    if not crew_name or not BUILTIN_LICENSE_MAP:
        return ""
    chinese = extract_chinese_name(crew_name)
    if chinese and chinese in BUILTIN_LICENSE_MAP:
        return BUILTIN_LICENSE_MAP[chinese]
    norm = normalize_name(crew_name)
    if not norm:
        return ""
    for key, val in BUILTIN_LICENSE_MAP.items():
        if normalize_name(key) == norm:
            return val
    return ""

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
        return f"{month:02d}月{int(day):02d}日"
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

# ---------- 填充模板 ----------
def fill_template(template_bytes, data, crew_list, passenger_list, route_display):
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
        safe_set_cell_value(ws, data_row, 5, route_display if route_display else "")

    # ----- 2. 机组信息 -----
    # 机长
    if len(crew_list) >= 1:
        crew = crew_list[0]
        for row in ws.iter_rows(min_row=1, max_row=50):
            for cell in row:
                if cell.value and isinstance(cell.value, str) and "机长" in cell.value:
                    row_num = cell.row
                    safe_set_cell_value(ws, row_num, 2, extract_chinese_name(crew["name"]))
                    safe_set_cell_value(ws, row_num, 3, crew.get("gender", ""))
                    safe_set_cell_value(ws, row_num, 4, crew.get("dob", ""))
                    safe_set_cell_value(ws, row_num, 5, crew.get("passport_no", ""))
                    # 执照号码
                    license_num = find_license(crew["name"])
                    safe_set_cell_value(ws, row_num, 6, license_num)
                    # 联系方式
                    contact = find_contact(crew["name"])
                    safe_set_cell_value(ws, row_num, 7, contact)
                    break
            else:
                continue
            break

    # 副驾驶
    if len(crew_list) >= 2:
        crew = crew_list[1]
        for row in ws.iter_rows(min_row=1, max_row=50):
            for cell in row:
                if cell.value and isinstance(cell.value, str) and "副驾驶" in cell.value:
                    row_num = cell.row
                    safe_set_cell_value(ws, row_num, 2, extract_chinese_name(crew["name"]))
                    safe_set_cell_value(ws, row_num, 3, crew.get("gender", ""))
                    safe_set_cell_value(ws, row_num, 4, crew.get("dob", ""))
                    safe_set_cell_value(ws, row_num, 5, crew.get("passport_no", ""))
                    license_num = find_license(crew["name"])
                    safe_set_cell_value(ws, row_num, 6, license_num)
                    contact = find_contact(crew["name"])
                    safe_set_cell_value(ws, row_num, 7, contact)
                    break
            else:
                continue
            break

    # 乘务、机务：从第3位开始遍历，分别取第一个女性和第一个男性
    cabin_crew = None
    mechanic = None
    for i in range(2, len(crew_list)):
        crew = crew_list[i]
        gender = str(crew.get("gender", "")).strip()
        if gender in ["女", "Female", "F"] and cabin_crew is None:
            cabin_crew = crew
        elif gender in ["男", "Male", "M"] and mechanic is None:
            mechanic = crew
        if cabin_crew and mechanic:
            break

    # 写入乘务行
    for row in ws.iter_rows(min_row=1, max_row=50):
        for cell in row:
            if cell.value and isinstance(cell.value, str) and "乘务" in cell.value:
                row_num = cell.row
                if cabin_crew:
                    safe_set_cell_value(ws, row_num, 2, extract_chinese_name(cabin_crew["name"]))
                    safe_set_cell_value(ws, row_num, 3, cabin_crew.get("gender", ""))
                    safe_set_cell_value(ws, row_num, 4, cabin_crew.get("dob", ""))
                    safe_set_cell_value(ws, row_num, 5, cabin_crew.get("passport_no", ""))
                    license_num = find_license(cabin_crew["name"])
                    safe_set_cell_value(ws, row_num, 6, license_num)
                    contact = find_contact(cabin_crew["name"])
                    safe_set_cell_value(ws, row_num, 7, contact)
                else:
                    safe_set_cell_value(ws, row_num, 2, "无")
                    safe_set_cell_value(ws, row_num, 3, "")
                    safe_set_cell_value(ws, row_num, 4, "")
                    safe_set_cell_value(ws, row_num, 5, "")
                    safe_set_cell_value(ws, row_num, 6, "")
                    safe_set_cell_value(ws, row_num, 7, "")
                break
        else:
            continue
        break

    # 写入机务行
    for row in ws.iter_rows(min_row=1, max_row=50):
        for cell in row:
            if cell.value and isinstance(cell.value, str) and "机务" in cell.value:
                row_num = cell.row
                if mechanic:
                    safe_set_cell_value(ws, row_num, 2, extract_chinese_name(mechanic["name"]))
                    safe_set_cell_value(ws, row_num, 3, mechanic.get("gender", ""))
                    safe_set_cell_value(ws, row_num, 4, mechanic.get("dob", ""))
                    safe_set_cell_value(ws, row_num, 5, mechanic.get("passport_no", ""))
                    license_num = find_license(mechanic["name"])
                    safe_set_cell_value(ws, row_num, 6, license_num)
                    contact = find_contact(mechanic["name"])
                    safe_set_cell_value(ws, row_num, 7, contact)
                else:
                    safe_set_cell_value(ws, row_num, 2, "无")
                    safe_set_cell_value(ws, row_num, 3, "")
                    safe_set_cell_value(ws, row_num, 4, "")
                    safe_set_cell_value(ws, row_num, 5, "")
                    safe_set_cell_value(ws, row_num, 6, "")
                    safe_set_cell_value(ws, row_num, 7, "")
                break
        else:
            continue
        break

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
st.info("⚠️ 注意：模板文件必须是 **.xlsx** 格式（非 .xls）。联系方式及部分执照号码已内置，无需额外上传。")

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

        # 生成默认行程显示文本
        from_code = data.get("from", "")
        to_code = data.get("to", "")
        date_str = data.get("date_str", "")
        utc_time = data.get("utc_time", "")
        default_route = ""
        if date_str and from_code and to_code:
            date_display = parse_date_display(date_str)
            if utc_time:
                bj_time = parse_utc_to_beijing(utc_time, date_str)
            else:
                bj_time = "0000"
            default_route = f"{date_display} {from_code} {bj_time} XXXX {to_code}"
        else:
            default_route = f"{from_code}-{to_code}" if from_code and to_code else ""

        st.subheader("📋 提取的航班信息")
        st.info(f"✈️ 默认航班行程：{default_route}")

        file_name_input = st.text_input("📝 自定义航班行程 / 下载文件名", value=default_route)

        route_display = file_name_input.strip()
        if not route_display:
            route_display = default_route

        result_bytes = fill_template(template_file, data, crew_list, passenger_list, route_display)

        safe_file_name = re.sub(r'[\\/*?:"<>|]', "_", route_display).strip()
        if not safe_file_name:
            safe_file_name = "备案表"
        download_file_name = f"{safe_file_name}.xlsx"

        st.download_button(
            label="⬇️ 下载填充后的备案表",
            data=result_bytes,
            file_name=download_file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"❌ 处理失败：{e}")
        st.exception(e)
else:
    st.info("👆 请同时上传 GD单 和 模板文件。")
