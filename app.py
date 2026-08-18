import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
import re
from datetime import datetime, timedelta
import pytz

st.set_page_config(page_title="GD单 → 公务飞行计划信息备案表", layout="wide")
st.title("🛫 GD单 → 公务飞行计划信息备案表")
st.markdown("上传 GD单 和模板，自动生成备案表（联系方式、执照号码及证件号码已内置）。")

# ---------- 内置联系方式映射 ----------
BUILTIN_CONTACT_MAP = {
    # ... (此处省略以节省篇幅，内容与之前完全一致) ...
    "孙辉": "139 1626 9572",
    "Hui Sun": "139 1626 9572",
    # ... 其余内容不变 ...
}

# ---------- 内置执照号码映射 ----------
BUILTIN_LICENSE_MAP = {
    # ... (此处省略以节省篇幅) ...
    "孙辉": "310228197810012612",
    "Hui Sun": "310228197810012612",
    # ... 其余内容不变 ...
}

# ---------- 内置证件号码映射（身份证/护照） ----------
BUILTIN_ID_MAP = {
    # ... (此处省略以节省篇幅) ...
    "孙辉": "310228197810012612",
    "Hui Sun": "310228197810012612",
    # ... 其余内容不变 ...
}

# ---------- 机型修正映射（GD单填写错误时使用）----------
AIRCRAFT_TYPE_CORRECTION = {
    "B3926": "LJ60",   # GD单误写为 LR60，实际为 LJ60
}

# ---------- 辅助函数 ----------
def correct_aircraft_type(reg, ac_type):
    if reg in AIRCRAFT_TYPE_CORRECTION:
        corrected = AIRCRAFT_TYPE_CORRECTION[reg]
        if ac_type != corrected:
            st.info(f"✈️ 机型修正：{ac_type} → {corrected}（注册号 {reg}）")
        return corrected
    return ac_type

def parse_datetime_to_beijing(utc_time_str, date_str):
    """将GD单中的UTC日期时间转换为北京时间"""
    try:
        # 1. 解析时间
        time_part = utc_time_str.replace('Z', '').strip()
        if len(time_part) == 4:
            hour = int(time_part[:2])
            minute = int(time_part[2:])
        elif len(time_part) == 3:
            hour = int(time_part[:1])
            minute = int(time_part[1:])
        else:
            return None, None

        # 2. 解析日期
        day = int(re.search(r'\d+', date_str).group())
        month_map = {"Jan":1,"Feb":2,"Mar":3,"Apr":4,"May":5,"Jun":6,
                     "Jul":7,"Aug":8,"Sep":9,"Oct":10,"Nov":11,"Dec":12}
        month_str = re.search(r'[A-Za-z]{3}', date_str).group()
        month = month_map.get(month_str[:3], 1)
        year = 2026  # GD单中日期不含年份，默认当年

        # 3. 构建UTC datetime并转换为北京时间
        utc_dt = datetime(year, month, day, hour, minute, tzinfo=pytz.UTC)
        beijing_tz = pytz.timezone('Asia/Shanghai')
        beijing_dt = utc_dt.astimezone(beijing_tz)

        return beijing_dt.strftime("%H%M"), beijing_dt.strftime("%Y-%m-%d")
    except Exception as e:
        print(f"日期时间解析失败: {e}")
        return None, None

def get_beijing_date_display(date_str):
    """获取北京时间的 'M月D日' 显示格式"""
    try:
        day = re.search(r'\d+', date_str).group()
        month_map = {"Jan":1,"Feb":2,"Mar":3,"Apr":4,"May":5,"Jun":6,
                     "Jul":7,"Aug":8,"Sep":9,"Oct":10,"Nov":11,"Dec":12}
        month_str = re.search(r'[A-Za-z]{3}', date_str).group()
        month = month_map.get(month_str[:3], 1)
        return f"{month}月{int(day)}日"
    except:
        return date_str

# ---------- 解析GD单（已修改） ----------
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
                    data["ac_type_raw"] = get_value_right(ws, cell.row, cell.column+1)
                    reg = data.get("reg", "")
                    data["ac_type"] = correct_aircraft_type(reg, data["ac_type_raw"])
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
                        
                        # ----- 新增：计算北京时间用于文件名 -----
                        if data.get("utc_time") and data.get("date_str"):
                            bj_time_str, bj_date_str = parse_datetime_to_beijing(
                                data["utc_time"], data["date_str"]
                            )
                            data["bj_time"] = bj_time_str
                            data["bj_date_str"] = bj_date_str
                        # ----- 新增结束 -----
    
    # ... (后续解析机组和旅客的代码不变) ...
    crew_data = []
    passenger_data = []
    section = None
    # ... (此处省略crew和passenger解析，与之前完全一致) ...
    
    return data, crew_data, passenger_data

# ---------- 主UI ----------
# ... (UI部分代码不变，但文件名生成逻辑修改) ...
if data_file and template_file:
    # ... (解析等操作不变) ...
    
    # ---- 修改：使用北京时间生成文件名 ----
    # 优先使用转换后的北京时间日期
    bj_date_str = data.get("bj_date_str")
    if bj_date_str:
        # 从 YYYY-MM-DD 转换为 M月D日
        date_parts = bj_date_str.split("-")
        if len(date_parts) == 3:
            month = int(date_parts[1])
            day = int(date_parts[2])
            date_display = f"{month}月{day}日"
        else:
            date_display = parse_date_display(data.get("date_str", ""))
    else:
        # 降级方案：使用原始GD日期
        date_display = parse_date_display(data.get("date_str", ""))
    
    # ... (后续route解析和文件名拼接不变) ...
