import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import re

st.set_page_config(page_title="通用声明 → 备案表生成器", layout="wide")
st.title("🛫 通用声明 → 公务飞行计划信息备案表")
st.markdown("上传通用声明 Excel，自动提取机组和乘客信息，填入备案表模板。")

# ---------- 国籍映射 ----------
NATION_MAP = {
    "CHN": "中国",
    "HKG": "香港",
    "DEU": "德国",
    "USA": "美国",
    "GBR": "英国",
    "FRA": "法国",
    "RUS": "俄罗斯",
    "JPN": "日本",
    "KOR": "韩国",
    "SGP": "新加坡",
    "MYS": "马来西亚",
    "THA": "泰国",
    "VNM": "越南",
    "PHL": "菲律宾",
    "IDN": "印度尼西亚",
    "IND": "印度",
    "AUS": "澳大利亚",
    "CAN": "加拿大",
    "BRA": "巴西",
    "MEX": "墨西哥",
    "ZAF": "南非",
    "EGY": "埃及",
    "NGA": "尼日利亚",
    "KEN": "肯尼亚",
    "TZA": "坦桑尼亚",
    "ZWE": "津巴布韦",
    "NLD": "荷兰",
    "ITA": "意大利",
    "ESP": "西班牙",
    "PRT": "葡萄牙",
    "GRC": "希腊",
    "TUR": "土耳其",
    "SAU": "沙特阿拉伯",
    "ARE": "阿联酋",
    "ISR": "以色列",
    "IRN": "伊朗",
    "PAK": "巴基斯坦",
    "BGD": "孟加拉",
    "NPL": "尼泊尔",
    "LKA": "斯里兰卡",
    "MMR": "缅甸",
    "KHM": "柬埔寨",
    "LAO": "老挝",
    "MNG": "蒙古",
    "PRK": "朝鲜",
    "TWN": "台湾地区",
    "MAC": "澳门",
    "HKG": "香港",
}
# 补充常见
NATION_MAP["CHN"] = "中国"
NATION_MAP["HKG"] = "香港"
NATION_MAP["DEU"] = "德国"

def get_nation_name(code):
    code = code.strip().upper()
    return NATION_MAP.get(code, code)  # 若未映射则保留原代码

def extract_chinese_name(full_name):
    """从 '梅峰 MEI Feng' 中提取中文名（优先），若无则返回原字符串"""
    # 如果包含中文字符，提取中文部分（以中文开头或包含中文）
    # 简单方法：按空格分割，取包含中文的部分
    parts = full_name.split()
    chinese_parts = [p for p in parts if re.search(r'[\u4e00-\u9fff]', p)]
    if chinese_parts:
        return " ".join(chinese_parts)  # 如 "梅峰"
    else:
        return full_name  # 全英文

def parse_document_type(passport_no, doc_type):
    """
    根据证件号和证件类型判断：身份证 or 护照
    优先依据doc_type，如果doc_type包含"身份证"或"居民身份证"则为身份证；
    如果包含"通行证"或"护照"则为护照；
    若无法判断，则根据passport_no格式：纯数字15/18位 -> 身份证，否则护照。
    """
    doc_type_str = str(doc_type).strip() if pd.notna(doc_type) else ""
    if "身份证" in doc_type_str or "居民身份证" in doc_type_str:
        return "身份证"
    if "通行证" in doc_type_str or "护照" in doc_type_str or "passport" in doc_type_str.lower():
        return "护照"
    # 如果doc_type为"机组证"或"空勤登机证"，我们仍然看passport_no格式
    pn = str(passport_no).strip() if pd.notna(passport_no) else ""
    # 移除空格
    pn = re.sub(r'\s+', '', pn)
    if re.match(r'^[0-9]{15}$', pn) or re.match(r'^[0-9]{17}[0-9Xx]$', pn):
        return "身份证"
    else:
        return "护照"

# ---------- 解析通用声明 ----------
def parse_general_declaration(file_bytes):
    wb = load_workbook(file_bytes)
    ws = wb.active
    data = {}
    # 1. 读取基础信息：OPERATOR, REG NO./FLT NO., AC TYPE, FROM, TO, DATE/TIME
    # 利用openpyxl查找关键词
    for row in ws.iter_rows(min_row=1, max_row=20):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                val = cell.value.strip()
                if "OPERATOR:" in val:
                    # 下一列
                    op_cell = ws.cell(row=cell.row, column=cell.column+1)
                    data["operator"] = op_cell.value.strip() if op_cell.value else ""
                elif "REG NO./FLT NO.:" in val:
                    # 下一列
                    reg_cell = ws.cell(row=cell.row, column=cell.column+1)
                    reg_val = reg_cell.value.strip() if reg_cell.value else ""
                    # 可能包含空格，分成注册号和航班号
                    parts = reg_val.split()
                    if len(parts) >= 2:
                        data["reg"] = parts[0]
                        data["flt"] = parts[1]
                    else:
                        data["reg"] = reg_val
                        data["flt"] = ""
                elif "AC TYPE:" in val:
                    type_cell = ws.cell(row=cell.row, column=cell.column+1)
                    data["ac_type"] = type_cell.value.strip() if type_cell.value else ""
                elif "FROM:" in val:
                    from_cell = ws.cell(row=cell.row, column=cell.column+1)
                    data["from"] = from_cell.value.strip() if from_cell.value else ""
                elif "TO:" in val:
                    to_cell = ws.cell(row=cell.row, column=cell.column+1)
                    data["to"] = to_cell.value.strip() if to_cell.value else ""
                elif "DATE/TIME:" in val:
                    date_cell = ws.cell(row=cell.row, column=cell.column+1)
                    data["date_time"] = date_cell.value.strip() if date_cell.value else ""

    # 2. 寻找机组和乘客表格
    # 使用状态机：找到 "CREW MANIFEST" 和 "PASSENGER MANIFEST"
    crew_data = []
    passenger_data = []
    section = None  # 'crew' or 'passenger'

    for row in ws.iter_rows(min_row=1):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                val = cell.value.strip()
                if "CREW MANIFEST" in val:
                    section = 'crew'
                    continue
                elif "PASSENGER MANIFEST" in val:
                    section = 'passenger'
                    continue
                elif "CARGO MANIFEST" in val or "DECLARATION OF HEALTH" in val:
                    section = None  # 结束
                    continue
        if section == 'crew':
            # 读取每行，检查是否有S/NO.列（第一列通常为序号）
            # 我们判断如果第一列为数字，且第二列有姓名，则认为是数据行
            first_cell = row[0]
            if first_cell.value and isinstance(first_cell.value, (int, float)):
                # 可能是一行数据
                if len(row) >= 7:
                    name_cell = row[1]
                    if name_cell.value and isinstance(name_cell.value, str):
                        # 提取
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
    wb = load_workbook(template_bytes)
    ws = wb.active

    # 1. 基础信息
    # 我们直接定位到特定单元格，手动写入
    # 假设模板固定结构，我们按位置写入
    # 机型（如Global 6000） -> 填到"机型"旁边的单元格，大约在B8? 我们通过文本搜索定位
    # 更稳健：搜索包含"机型"的单元格，然后写入右侧
    for row in ws.iter_rows(min_row=1, max_row=30):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                val = cell.value.strip()
                if "机型" in val:
                    target = ws.cell(row=cell.row, column=cell.column+1)
                    target.value = data.get("ac_type", "")
                elif "注册号" in val:
                    target = ws.cell(row=cell.row, column=cell.column+1)
                    target.value = data.get("reg", "")
                elif "航班号" in val:
                    target = ws.cell(row=cell.row, column=cell.column+1)
                    target.value = data.get("flt", "")
                elif "航班行程" in val:
                    target = ws.cell(row=cell.row, column=cell.column+1)
                    from_ = data.get("from", "")
                    to_ = data.get("to", "")
                    target.value = f"{from_}-{to_}" if from_ and to_ else ""
                elif "飞行目的" in val:
                    # 下面一行有"商务"等，我们找到"商务"单元格并保留，但用户固定为商务，所以可以不修改
                    pass
                elif "填报人" in val:
                    # 不修改，保留模板原有
                    pass

    # 2. 机组信息（最多4人）
    # 模板中机长、副驾驶、乘务、机务分别在特定行，我们通过查找"机长"、"副驾驶"等
    crew_positions = ["机长", "副驾驶", "乘务", "机务"]
    for idx, position in enumerate(crew_positions):
        if idx >= len(crew_list):
            break
        crew = crew_list[idx]
        # 寻找包含"机长"的单元格，然后写入该行
        for row in ws.iter_rows(min_row=1, max_row=50):
            for cell in row:
                if cell.value and isinstance(cell.value, str) and position in cell.value:
                    # 找到该行
                    row_num = cell.row
                    # 姓名列（第2列）
                    name_cell = ws.cell(row=row_num, column=2)
                    name_cell.value = extract_chinese_name(crew["name"])
                    # 性别列（第3列）
                    gender_cell = ws.cell(row=row_num, column=3)
                    gender_cell.value = crew.get("gender", "")
                    # 出生日期列（第4列）
                    dob_cell = ws.cell(row=row_num, column=4)
                    dob_cell.value = crew.get("dob", "")
                    # 证件号码、执照号码、联系方式不填
                    break

    # 3. 乘客信息（从第?行开始，需要找到"乘客信息"下面的表格）
    # 找到"乘客信息"标题，然后下面一行是表头（姓名、性别、出生日期、国籍、证件种类、证件号码）
    # 我们定位到"姓名"单元格，然后往下一行开始填写
    passenger_start_row = None
    for row in ws.iter_rows(min_row=1, max_row=100):
        for cell in row:
            if cell.value and isinstance(cell.value, str) and "姓名" in cell.value and "性别" in cell.value:
                # 表头行
                passenger_start_row = cell.row + 1
                break
        if passenger_start_row:
            break

    if passenger_start_row is None:
        # 尝试找"乘客信息"标题
        for row in ws.iter_rows(min_row=1, max_row=100):
            for cell in row:
                if cell.value and isinstance(cell.value, str) and "乘客信息" in cell.value:
                    passenger_start_row = cell.row + 2  # 可能表头在下一行
                    break
            if passenger_start_row:
                break

    if passenger_start_row:
        # 清空原有乘客数据（从起始行开始，删除旧数据，但保留格式？我们直接覆盖）
        # 我们直接写入，但不会删除多余行，如果数据少于原模板保留的空行，我们只覆盖前N行，后面的清空或保留空白
        # 我们先清空这些行内容，然后写入新数据
        for i in range(20):  # 最多20个乘客
            row_num = passenger_start_row + i
            for col in [1,2,3,4,5,6]:  # 6列
                ws.cell(row=row_num, column=col).value = None
        # 写入新数据
        for i, pax in enumerate(passenger_list):
            if i >= 20:
                break
            row_num = passenger_start_row + i
            ws.cell(row=row_num, column=1).value = extract_chinese_name(pax["name"])
            ws.cell(row=row_num, column=2).value = pax.get("gender", "")
            ws.cell(row=row_num, column=3).value = pax.get("dob", "")
            # 国籍：转换
            nationality = pax.get("nationality", "")
            ws.cell(row=row_num, column=4).value = get_nation_name(nationality)
            # 证件种类
            doc_type = parse_document_type(pax.get("passport_no", ""), pax.get("doc_type", ""))
            ws.cell(row=row_num, column=5).value = doc_type
            # 证件号码
            ws.cell(row=row_num, column=6).value = pax.get("passport_no", "")

    # 保存
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ---------- Streamlit UI ----------
st.subheader("📂 上传文件")
data_file = st.file_uploader("上传通用声明 Excel（如 MLLINGen DecZGSZ-ZLXY 23Jul.xlsx）", type=["xlsx"], key="data")
template_file = st.file_uploader("上传备案表模板 Excel（如 07月23日 MLLIN 深圳-西安.xls）", type=["xls", "xlsx"], key="template")

if data_file and template_file:
    try:
        # 解析通用声明
        data, crew_list, passenger_list = parse_general_declaration(data_file)
        st.success(f"✅ 解析成功：机组 {len(crew_list)} 人，乘客 {len(passenger_list)} 人")
        st.write("提取的机组信息：", pd.DataFrame(crew_list))
        st.write("提取的乘客信息：", pd.DataFrame(passenger_list))

        # 填充模板
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
