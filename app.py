# -*- coding: utf-8 -*-
"""
备案表 / 世界时行程 / 航路处理 / 批复核对
"""

import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import io
from io import BytesIO
from openpyxl import load_workbook
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import re
import csv
import copy
import datetime as _pcdt
from datetime import datetime, timedelta
import traceback
import json
import os

# ---------- 页面设置 ----------
st.set_page_config(page_title="备案表&世界时行程&航路处理&批复核对", layout="wide")
st.title("🛫 备案表 / 世界时行程 / 航路处理 / 批复核对")

# ---------- 创建选项卡 ----------
tab1, tab2, tab3, tab4 = st.tabs([
    "📋 功能1：备案表生成",
    "🌐 功能2：世界时行程",
    "✈️ 功能3：航路处理工具",
    "🔍 功能4：批复核对",
])

# ================================================================
# 功能1：备案表生成
# ================================================================
with tab1:
    st.markdown("上传 GD单 和模板，自动生成备案表（联系方式、执照号码及证件号码已内置）。")

    from copy import copy as _copy_style
    import json as _json

    BUILTIN_CREW_DATA = [
        ("庚凡", "139 2463 9747", "430104197901184015", "430104197901184015"),
        ("张永一 / Yongyi ZHANG", "139 0125 9544", "110102196605202336", "110102196605202336"),
        ("梅峰 / Feng MEI", "135 0967 8127", "17205/1 FCL", "510107197911242636"),
        ("王斌 / Bin WANG", "139 2527 2867", "3340398", "610104197911058331"),
        ("王少雄 / Shaoxiong WANG", "186 8387 9841", "510105198609042555", "510105198609042555"),
        ("苗旺旺 / Wangwang MIAO", "138 1871 5251", "410781198608019797", "410781198608019797"),
        ("赵岩松 / Yansong ZHAO", "186 1161 8385", "410103197004017014", "410103197004017014"),
        ("Bruce Roderick, WAINES", "186 6532 9796", "3448726", "000336198206158001"),
        ("Oliver Viktor, RACZ", "186 1197 3165", "000336198206158001", "000336198206158001"),
        ("Yiftah RAUCH / Yiftah, RAUCH", "186 1045 0563", "000972198112152001", "000972198112152001"),
        ("尤欣 / Xin YOU", "139 1608 5072", "620102197604293015", "620102197604293015"),
        ("李亚民 / Yamin Li", "133 6632 0878", "350104197107184915", "350104197107184915"),
        ("赵镭 / Lei ZHAO", "138 0883 9660", "440301198204157271", "440301198204157271"),
        ("彭罡", "186 1263 1888", "3240393", "440111198403244812"),
        ("胡君量 / Wan Leung WU / Kwan Leung WU", "132 6695 8816", "3025478", "124133200"),
        ("吴鹏", "136 1110 5901", "130103197602102115", "130103197602102115"),
        ("刘汇川", "188 5827 2791", "330103199003191618", "330103199003191618"),
        ("于龙飞", "156 5288 0812", "ZN00915", "210103198808123928"),
        ("林帅", "138 1156 6711", "110227198601130015", "110227198601130015"),
        ("张佳妮 / Jiani ZHANG", "136 6169 9966", "10183", "31010619840115002X"),
        ("张欢乐 / Huanle ZHANG", "186 1652 1529", "ZN00883", "330381198705292523"),
        ("昝昭君 / Zhaojun ZAN", "182 9570 0579", "14272419950203313X", "14272419950203313X"),
        ("孙赫 / He SUN", "186 0102 1216", "4070599", "410102197702243012"),
        ("李晓龙 / Xiaolong LI", "138 2378 1747", "3781955", "420104197906150015"),
        ("李卉妍 / Huiyan LI", "138 5142 0321", "10299", ""),
        ("熊立凌 / Liling Xiong", "135 3821 6276", "421003199203222631", "421003199203222631"),
        ("HEALY, Darran William / Darran William HEALY", "852 6891 2350", "3141079", ""),
        ("ROEDER, SIMONE ELKE / SIMONE ELKE ROEDER", "852 6263 4569", "4213276", ""),
        ("王凯珮 / HOI PUI, WONG", "852 6858 8410", "10068", "Z411582(2)"),
        ("马坚 / Jian MA", "189 1770 2918", "320103196607089512", "320103196607089512"),
        ("卢江 / Jiang LU", "158 0045 6521", "10302", "420521198809280021"),
        ("孟周聪 / Zhoucong Meng", "135 6434 5029", "10303", "310113198307243215"),
        ("张帆 / Fan ZHANG", "138 0135 1294", "2552356", "211002197306050057"),
        ("魏思远 / Siyuan WEI", "133 2110 4588", "230102198911054315", "230102198911054315"),
        ("王晟磊 / Shenglei WANG", "150 2689 7493", "310109198409264012", "310109198409264012"),
        ("Nathon Andrew G, NORBERG / Nathon Andrew G NORBERG", "186 0019 4610", "3050802", "A04700868"),
        ("Keith Robert, SHERREN / Keith Robert SHERREN", "137 3540 9744", "2755204", "P024165FF"),
        ("Rodolfo, BONETTI / Rodolfo BONETTI", "132 6284 1083", "12660", ""),
        ("危慧 / Hui WEI", "152 1349 1328", "10137", "43072119941115468X"),
        ("李辛欣 / Xinxin LI", "135 5008 8666", "4209424", "510105198605023015"),
        ("Herve Daniel, STAMM / Herve Daniel STAMM", "183 1709 0300", "2666622", "M578620(2)"),
        ("樊婉程 / Wancheng FAN", "186 2017 4817", "ZN00434", "440106199608180326"),
        ("刘爽 / Shuang LIU", "138 0125 8789", "4101498", "110108196308296450"),
        ("刘凯 / Kai LIU", "135 2157 9157", "2833670", "230103198103275511"),
        ("詹佩佩 / Peipei ZHAN", "137 1440 5925", "10283", "440301198809275123"),
        ("花佩 / Pei HUA", "186 2631 0634", "ZN00495", "320281198911117761"),
        ("翁英 / Ying WENG", "130 6785 2000", "ZN00905", "330106198801182024"),
        ("王莹 / Ying WANG", "159 1009 9069", "370125199004215621", "370125199004215621"),
        ("程佳俊 / Jiajun CHENG", "134 8010 3029", "511202198208161358", "511202198208161358"),
        ("蔡雨桐 / Yu Tong CHOI", "852 6426 7445", "10269", "V146532(5)"),
        ("俞凯 / Kai YU", "130 0579 0326", "120110196912180351", "120110196912180351"),
        ("杨杰 / Jie YANG", "186 1694 8903", "310104198411304413", "310104198411304413"),
        ("徐卓 / Zhuo XU", "139 1028 5510", "110105198906307113", "110105198906307113"),
        ("丁燕栒 / Yanxun DING", "135 6035 3829", "ZN00499", "440510199107260826"),
        ("万虹波 / Hongbo WAN", "133 1297 9906", "36050219860709003X", "36050219860709003X"),
        ("王国勤 / Guoqin WANG", "138 1815 8715", "2589322", "340822197610240215"),
        ("李潇恩 / Xiaoen LI", "158 0599 1600", "ZN00468", "510802199512100046"),
        ("孙辉 / Hui Sun", "139 1626 9572", "310228197810012612", "310228197810012612"),
        ("范蕾蕾 / Leilei FAN", "182 1000 6866", "ZN00347", "37010319861027002X"),
        ("张贺新 / Hexin ZHANG", "136 3773 1210", "420106197101020435", "420106197101020435"),
        ("李庆宏 / Qinghong LI", "135 0909 0503", "17260/1 FCL", "441402199107140256"),
        ("Eduard Pascal, Roski / Eduard Pascal Roski", "49 170 1534666", "3863535", "C9TR22VZM"),
        ("Peter Robert, JACKSON / Peter Robert JACKSON", "186 8245 1935", "000044197906253001", "000044197906253001"),
        ("廉卓群 / Zhuoqun LIAN", "133 5632 3949", "ZN00430", "370402199702018029"),
        ("孙浩 / Hao SUN", "136 7012 1990", "650103199102160037", "650103199102160037"),
        ("姚艳阁 / Yange YAO", "156 0127 9399", "10204", "130922199601151224"),
        ("茅邂文 / Xiewen MAO", "152 5181 7375", "ZN00903", "320683199911098627"),
        ("宋炜 / Wei SONG", "136 3256 5565", "37060219820621211X", "37060219820621211X"),
        ("BEEBE, Thaddeus John / Thaddeus John BEEBE", "852 6930 1609", "2743899", "R909452(A)"),
        ("张哲 / Zhe ZHANG", "139 0247 5026", "650104196604163310", "650104196604163310"),
        ("李海 / Hai LI", "136 8131 8388", "110105197201106130", "110105197201106130"),
        ("蔡国俊 / Kuo-Chun, TSAI", "86 157 1220 8304", "17203/1 FCL", "360019647"),
        ("杨涛 / Tao, YANG", "86 186 0122 5737", "140103197302010034", "140103197302010034"),
        ("朱正宇", "189 8335 3697", "350111197207152412", "350111197207152412"),
        ("金尚明 / Shangming JIN", "136 7113 8047", "210381197511034612", "210381197511034612"),
        ("赵婷婷", "138 2883 3162", "372524198212240023", "372524198212240023"),
        ("赖小燕 / Siau Mui LAI", "60 1239 05520", "A71366020", "A71366020"),
        ("何静文 / Ching Man, HO", "852 6421 0994", "Z630284(0)", "Z630284(0)"),
        ("周丽欢", "152 5710 6140", "330411199101195440", "330411199101195440"),
        ("AYA, MUGURUMA", "81 8071140700", "TT5868294", ""),
        ("梁广煜", "137 9428 7177", "EK9629672", ""),
        ("李园", "139 1178 3914", "110105198309149639", "110105198309149639"),
        ("孔铮", "139 1085 3981", "11010419801116125X", "11010419801116125X"),
        ("高峰", "135 8186 9017", "130826198001175332", "130826198001175332"),
        ("李阳", "133 6603 6567", "11010319850705091X", "11010319850705091X"),
        ("林生", "159 2162 9406", "350627198512262530", "350627198512262530"),
        ("谢依椿", "159 8942 4501", "441427198601221716", "441427198601221716"),
        ("廖关荣", "181 0755 9103", "360732199009095838", "360732199009095838"),
        ("林峰", "135 2260 7955", "150402198501122713", "150402198501122713"),
        ("丘东", "136 0044 6505", "11010119650406453x", "11010119650406453x"),
        ("王庆辉", "134 8013 9352", "350627198212052013", "350627198212052013"),
        ("姜磊", "135 1006 5318", "360502198312071333", "360502198312071333"),
        ("黄彦杰", "132 1157 2184", "450881199810196233", "450881199810196233"),
        ("王珍", "198 6662 9312", "622627199605283017", "622627199605283017"),
        ("赵国庆", "155 8849 2975", "370403200003133433", "370403200003133433"),
        ("王军", "853 62666900", "1458107(8)", "1458107(8)"),
        ("张德桃", "130 2883 6410", "360311199603192012", "360311199603192012"),
        ("江焰辉", "136 3148 0927", "440182199307270615", "440182199307270615"),
        ("孙龙", "156 9558 0691", "341623200010015613", "341623200010015613"),
        ("梁平", "138 2750 6225", "441223198510246217", "441223198510246217"),
        ("冯仁毫", "185 2028 6463", "440582199101153650", "440582199101153650"),
        ("焦石军", "139 2388 3525", "421224198310151013", "421224198310151013"),
        ("万子辰", "177 7005 7193", "", ""),
        ("卓辉", "157 7070 8632", "36073219981213009X", "36073219981213009X"),
        ("苏志斌", "159 0150 7150", "350782198308221539", "350782198308221539"),
        ("赵康", "191 6764 6172", "430321200303160170", "430321200303160170"),
        ("翟征宇", "134 1448 9793", "140211198612050031", "140211198612050031"),
        ("陈居瑜", "158 8962 6660", "460004197905030814", "460004197905030814"),
        ("林毅", "136 8642 0153", "350102197903213219", "350102197903213219"),
        ("Daniel, RICHTER", "374 9521 3832", "2556011", "C4W9RFH14"),
        ("郭春旭 / Guo Chunxu", "138 0136 1720", "110107197305150016", "110107197305150016"),
        ("黄海东", "138 0179 9315", "310105197506021215", "310105197506021215"),
    ]

    CREW_COLUMNS = ["姓名", "联系方式", "执照号码", "证件号码"]
    CREW_FILL_COLUMNS = ["职务", "姓名", "性别", "出生日期", "证件号码", "执照号码", "联系方式"]

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
        "TWN": "中国台湾", "MAC": "澳门"
    }

    AIRCRAFT_TYPE_MAP = {
        "B3926": "LJ60", "B652R": "GLF4", "B8105": "GLEX", "B8160": "GLF5",
        "B8262": "GLF4", "B8292": "GLF5", "B8309": "GLF5", "MLLIN": "GLEX",
        "N2QE": "GL5T", "N328LM": "GL7T", "N550DR": "GLF5", "N577QT": "F900",
        "N7777U": "GLEX", "N777ZH": "GLF5", "N88AY": "GLF5", "T7178HT": "GL7T",
        "T7CJK": "GLEX", "VPCSZ": "GL7T", "VPCVA": "GLF6", "B652Q": "GLF4",
        "B652S": "GLF4", "B65AP": "GLF4",
    }

    def get_aircraft_type(reg, gd_type=""):
        reg_clean = str(reg).strip().upper() if reg else ""
        gd_type_str = str(gd_type).strip() if gd_type else ""
        if reg_clean and reg_clean in AIRCRAFT_TYPE_MAP:
            mapped = AIRCRAFT_TYPE_MAP[reg_clean]
            if gd_type_str and gd_type_str.upper() != mapped.upper():
                st.info(f"✈️ 机型按注册号确定：{gd_type_str} → {mapped}（注册号 {reg_clean}）")
            return mapped
        return gd_type_str

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
        if not name:
            return ""
        name = re.sub(r'[\u4e00-\u9fff]+', '', name)
        name = re.sub(r'[,\s]+', ' ', name).strip().lower()
        return ' '.join(sorted(name.split()))

    def _split_crew_name(name_val):
        return [p.strip() for p in re.split(r'\s*[/|、]\s*', name_val) if p.strip()]

    def _find_crew_field(crew_name, field):
        if not crew_name:
            return ""
        target = str(crew_name).strip()
        target_cn = extract_chinese_name(target)
        target_norm = normalize_name(target)
        field_idx = CREW_COLUMNS.index(field)
        result = ""
        for row in BUILTIN_CREW_DATA:
            name_val = str(row[0] or "").strip()
            if not name_val:
                continue
            parts = _split_crew_name(name_val)
            matched = False
            if name_val == target:
                matched = True
            if not matched:
                for p in parts:
                    if p == target:
                        matched = True; break
                    if target_cn and p == target_cn:
                        matched = True; break
                    if target_norm and normalize_name(p) == target_norm:
                        matched = True; break
            if matched:
                val = str(row[field_idx] or "").strip()
                if val:
                    result = val
        return result

    def find_contact(crew_name):
        return _find_crew_field(crew_name, "联系方式")

    def find_license(crew_name):
        return _find_crew_field(crew_name, "执照号码")

    def find_id(crew_name):
        return _find_crew_field(crew_name, "证件号码")

    def is_18digit_id_card(val):
        if not val:
            return False
        s = re.sub(r'\s+', '', str(val))
        return bool(re.match(r'^[0-9]{17}[0-9Xx]$', s))

    def resolve_crew_id_and_license(crew_name, fallback_passport=""):
        id_val = find_id(crew_name)
        license_val = find_license(crew_name)
        if is_18digit_id_card(id_val):
            return id_val, id_val
        return (id_val if id_val else fallback_passport), license_val

    def parse_document_type(passport_no, doc_type):
        doc_type_str = str(doc_type).strip() if pd.notna(doc_type) else ""
        if doc_type_str:
            if "中华人民共和国居民身份证" in doc_type_str:
                return "身份证"
            if "港澳居民来往内地通行证" in doc_type_str:
                return "回乡证"
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
                hour = int(time_part[:2]); minute = int(time_part[2:])
            elif len(time_part) == 3:
                hour = int(time_part[:1]); minute = int(time_part[1:])
            else:
                return "0000"
            day = int(re.search(r'\d+', date_str).group()) if re.search(r'\d+', date_str) else 1
            month_map = {"Jan":1,"Feb":2,"Mar":3,"Apr":4,"May":5,"Jun":6,
                         "Jul":7,"Aug":8,"Sep":9,"Oct":10,"Nov":11,"Dec":12}
            month_str = re.search(r'[A-Za-z]{3}', date_str).group() if re.search(r'[A-Za-z]{3}', date_str) else "Jan"
            month = month_map.get(month_str[:3], 1)
            dt_beijing = datetime(2026, month, day, hour, minute) + timedelta(hours=8)
            return dt_beijing.strftime("%H%M")
        except:
            return "0000"

    def parse_date_display(date_str):
        try:
            day = re.search(r'\d+', date_str).group()
            month_str = re.search(r'[A-Za-z]{3}', date_str).group()
            month_map = {"Jan":1,"Feb":2,"Mar":3,"Apr":4,"May":5,"Jun":6,
                         "Jul":7,"Aug":8,"Sep":9,"Oct":10,"Nov":11,"Dec":12}
            month = month_map.get(month_str[:3], 1)
            return f"{month}月{int(day)}日"
        except:
            return date_str

    def get_beijing_date_display(utc_time_str, date_str):
        if not utc_time_str or not date_str:
            return parse_date_display(date_str)
        try:
            time_part = utc_time_str.replace('Z', '').strip()
            if len(time_part) == 4:
                hour = int(time_part[:2]); minute = int(time_part[2:])
            elif len(time_part) == 3:
                hour = int(time_part[:1]); minute = int(time_part[1:])
            else:
                return parse_date_display(date_str)
            day = int(re.search(r'\d+', date_str).group())
            month_str = re.search(r'[A-Za-z]{3}', date_str).group()
            month_map = {"Jan":1,"Feb":2,"Mar":3,"Apr":4,"May":5,"Jun":6,
                         "Jul":7,"Aug":8,"Sep":9,"Oct":10,"Nov":11,"Dec":12}
            month = month_map.get(month_str[:3], 1)
            dt_beijing = datetime(2026, month, day, hour, minute) + timedelta(hours=8)
            return f"{dt_beijing.month}月{dt_beijing.day}日"
        except:
            return parse_date_display(date_str)

    def strip_single_letter_prefix(text):
        if text and re.match(r'^[A-Za-z]\s+', text):
            return re.sub(r'^[A-Za-z]\s+', '', text)
        return text

    def parse_no_time_route(input_text, date_display):
        input_text = strip_single_letter_prefix(input_text)
        if not input_text or not input_text.strip():
            return None
        text = input_text.strip()
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        if len(lines) >= 2:
            first_line = lines[0]; second_line = lines[1]
            flight_number = first_line.split()[0] if first_line.split() else None
            if not flight_number:
                return None
            parts = re.split(r'\s*[-—–]\s*', second_line)
            if len(parts) >= 2:
                dep_airport = parts[0].strip(); arr_airport = parts[1].strip()
                if dep_airport and arr_airport:
                    return f"{date_display} {flight_number} {dep_airport}-{arr_airport}"
        else:
            flight_number = text.split()[0] if text.split() else None
            if not flight_number:
                return None
            time_pattern = (
                r'(?<![A-Za-z0-9])(\d{1,2}:\d{2}|\d{4})(?![A-Za-z0-9])'
                r'\s*[-—–]\s*'
                r'(\d{1,2}:\d{2}|\d{4})(?![A-Za-z0-9])'
            )
            remaining = re.sub(time_pattern, '', text).strip()
            remaining = re.sub(r'\s*\+\s*\d+\s*', '', remaining).strip()
            airport_parts = re.split(r'\s*[-—–]\s*', remaining)
            if len(airport_parts) >= 2:
                dep_airport = airport_parts[-2].strip(); arr_airport = airport_parts[-1].strip()
                ch_dep = re.findall(r'[\u4e00-\u9fff]+', dep_airport)
                ch_arr = re.findall(r'[\u4e00-\u9fff]+', arr_airport)
                if ch_dep and ch_arr:
                    dep_airport = ''.join(ch_dep); arr_airport = ''.join(ch_arr)
                if dep_airport and arr_airport:
                    return f"{date_display} {flight_number} {dep_airport}-{arr_airport}"
        return None

    def parse_with_time_route(input_text, date_display):
        input_text = strip_single_letter_prefix(input_text)
        if not input_text or not input_text.strip():
            return input_text
        text = input_text.strip()
        time_pattern = (
            r'(?<![A-Za-z0-9])(\d{1,2}:\d{2}|\d{4})(?![A-Za-z0-9])'
            r'\s*[-—–]\s*'
            r'(\d{1,2}:\d{2}|\d{4})(?![A-Za-z0-9])'
        )
        time_match = re.search(time_pattern, text)
        if time_match:
            dep_time = time_match.group(1).replace(':', ''); arr_time = time_match.group(2).replace(':', '')
            remaining = re.sub(time_pattern, '', text).strip()
        else:
            time_pattern2 = (
                r'(?<![A-Za-z0-9])(\d{1,2}:\d{2}|\d{4})(?![A-Za-z0-9])'
                r'\s+'
                r'(\d{1,2}:\d{2}|\d{4})(?![A-Za-z0-9])'
            )
            time_match2 = re.search(time_pattern2, text)
            if time_match2:
                dep_time = time_match2.group(1).replace(':', ''); arr_time = time_match2.group(2).replace(':', '')
                remaining = re.sub(time_pattern2, '', text).strip()
            else:
                return input_text
        remaining = re.sub(r'\s*\+\s*\d+\s*', '', remaining).strip()
        airport_pattern = r'(.+?)\s*[-—–]\s*(.+)'
        airport_match = re.search(airport_pattern, remaining)
        if airport_match:
            dep_airport = airport_match.group(1).strip(); arr_airport = airport_match.group(2).strip()
            def extract_chinese(text):
                chinese = re.findall(r'[\u4e00-\u9fff]+', text)
                return ''.join(chinese) if chinese else text
            dep_airport = extract_chinese(dep_airport); arr_airport = extract_chinese(arr_airport)
            if dep_airport and arr_airport:
                return f"{date_display} {dep_airport}{dep_time}-{arr_time}{arr_airport}"
        return input_text

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
        data["ac_type"] = get_aircraft_type(
            data.get("reg", ""), data.get("ac_type_raw", "")
        )

        crew_data = []; passenger_data = []; section = None
        for row in ws.iter_rows(min_row=1):
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    val = cell.value.strip()
                    if "CREW MANIFEST" in val:
                        section = 'crew'; break
                    elif "PASSENGER MANIFEST" in val:
                        section = 'passenger'; break
                    elif "CARGO MANIFEST" in val or "DECLARATION OF HEALTH" in val:
                        section = None; break
            if section == 'crew':
                first_cell = row[0]
                if first_cell.value and isinstance(first_cell.value, (int, float)) and len(row) >= 7:
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
                if first_cell.value and isinstance(first_cell.value, (int, float)) and len(row) >= 7:
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

    def _parse_hhmm(val):
        if val is None:
            return None
        try:
            if isinstance(val, float) and pd.isna(val):
                return None
        except Exception:
            pass
        if hasattr(val, 'strftime'):
            try:
                return val.strftime('%H:%M')
            except Exception:
                pass
        s = str(val).strip()
        m = re.match(r'^(\d{1,2}):(\d{2})', s)
        if m:
            return f"{int(m.group(1)):02d}:{m.group(2)}"
        return None

    def parse_route_plan(file_bytes):
        try:
            df = pd.read_excel(file_bytes, skiprows=1)
            df.columns = [str(c).strip() for c in df.columns]
            if '飞机注册号' not in df.columns:
                df = pd.read_excel(file_bytes, header=1)
                df.columns = [str(c).strip() for c in df.columns]
            return df
        except Exception as e:
            st.warning(f"⚠️ 航段数据解析失败：{e}")
            return None

    def _parse_gd_date(date_str):
        if not date_str:
            return None
        m = re.match(r'(\d{1,2})\s*([A-Za-z]{3})', str(date_str).strip())
        if not m:
            return None
        try:
            day = int(m.group(1))
        except Exception:
            return None
        month_map = {"Jan":1,"Feb":2,"Mar":3,"Apr":4,"May":5,"Jun":6,
                     "Jul":7,"Aug":8,"Sep":9,"Oct":10,"Nov":11,"Dec":12}
        month = month_map.get(m.group(2).capitalize()[:3])
        if not month:
            return None
        return (month, day)

    def _extract_row_date(row_date):
        if row_date is None:
            return None
        try:
            if isinstance(row_date, float) and pd.isna(row_date):
                return None
        except Exception:
            pass
        if hasattr(row_date, 'year'):
            try:
                return datetime(row_date.year, row_date.month, row_date.day)
            except Exception:
                pass
        s = str(row_date).strip()
        m = re.match(r'(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})', s)
        if m:
            try:
                return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except Exception:
                return None
        m = re.match(r'(\d{4})(\d{2})(\d{2})', s)
        if m:
            try:
                return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except Exception:
                return None
        return None

    def find_matching_flight(df, reg, from_airport, to_airport, gd_date=None):
        if df is None or df.empty:
            return None
        required = ['飞机注册号', '出发地', '到达地', '计划出发', '预计到达', '出发城市', '到达城市']
        if not all(c in df.columns for c in required):
            return None

        mask = (
            (df['飞机注册号'].astype(str).str.strip() == str(reg).strip()) &
            (df['出发地'].astype(str).str.strip().str.upper() == str(from_airport).strip().upper()) &
            (df['到达地'].astype(str).str.strip().str.upper() == str(to_airport).strip().upper())
        )
        matched = df[mask]
        if matched.empty:
            return None

        if gd_date is None or '出发日期' not in matched.columns:
            return matched.iloc[0]

        gd_month, gd_day = gd_date

        def _same_month_day(row_date):
            dt = _extract_row_date(row_date)
            return dt is not None and dt.month == gd_month and dt.day == gd_day

        exact = matched[matched['出发日期'].apply(_same_month_day)]
        if not exact.empty:
            return exact.iloc[0]

        def _nearby_month_day(row_date):
            dt = _extract_row_date(row_date)
            if dt is None:
                return False
            try:
                target = datetime(dt.year, gd_month, gd_day)
            except Exception:
                return False
            return abs((dt - target).days) <= 1

        nearby = matched[matched['出发日期'].apply(_nearby_month_day)]
        if not nearby.empty:
            return nearby.iloc[0]

        return None

    def build_route_from_flight(flight_row):
        reg = str(flight_row.get('飞机注册号', '') or '').strip()
        dep_time = _parse_hhmm(flight_row.get('计划出发'))
        arr_time = _parse_hhmm(flight_row.get('预计到达'))
        dep_city = str(flight_row.get('出发城市', '') or '').strip()
        arr_city = str(flight_row.get('到达城市', '') or '').strip()
        if not (reg and dep_time and arr_time and dep_city and arr_city):
            return None
        dep_time_clean = dep_time.replace(':', '')
        arr_time_clean = arr_time.replace(':', '')
        return f"F {reg} {dep_time_clean} - {arr_time_clean}  {dep_city} - {arr_city}"

    def _style_name_cell(ws, row, col, text):
        target_row, target_col = row, col
        for merged_range in ws.merged_cells.ranges:
            if merged_range.min_row <= row <= merged_range.max_row and \
               merged_range.min_col <= col <= merged_range.max_col:
                target_row = merged_range.min_row
                target_col = merged_range.min_col
                break

        cell = ws.cell(row=target_row, column=target_col)
        text_str = str(text) if text is not None else ""

        new_align = _copy_style(cell.alignment)
        new_align.wrap_text = True
        cell.alignment = new_align

        n = len(text_str)
        if n <= 12:
            return

        if n <= 17:
            new_size = 10
        elif n <= 25:
            new_size = 9
        elif n <= 35:
            new_size = 8
        else:
            new_size = 7

        new_font = _copy_style(cell.font)
        new_font.size = new_size
        cell.font = new_font

        min_h = new_size * 2 + 4
        rd = ws.row_dimensions[target_row]
        current_h = rd.height
        if current_h is None or current_h < min_h:
            rd.height = min_h

    def _fill_crew_row_by_data(ws, label_keyword, crew_row):
        for row in ws.iter_rows(min_row=1, max_row=50):
            for cell in row:
                if cell.value and isinstance(cell.value, str) and label_keyword in cell.value:
                    row_num = cell.row
                    name_text = crew_row.get("姓名", "")
                    safe_set_cell_value(ws, row_num, 2, name_text)
                    safe_set_cell_value(ws, row_num, 3, crew_row.get("性别", ""))
                    safe_set_cell_value(ws, row_num, 4, crew_row.get("出生日期", ""))
                    safe_set_cell_value(ws, row_num, 5, crew_row.get("证件号码", ""))
                    safe_set_cell_value(ws, row_num, 6, crew_row.get("执照号码", ""))
                    safe_set_cell_value(ws, row_num, 7, crew_row.get("联系方式", ""))
                    _style_name_cell(ws, row_num, 2, name_text)
                    return True
        return False

    def _fill_empty_crew_row(ws, label_keyword):
        for row in ws.iter_rows(min_row=1, max_row=50):
            for cell in row:
                if cell.value and isinstance(cell.value, str) and label_keyword in cell.value:
                    row_num = cell.row
                    for c in range(2, 8):
                        safe_set_cell_value(ws, row_num, c, "无" if c == 2 else "")
                    return True
        return False

    # ---------- ★ 新增：乘客区自动扩展 ----------
    def _style_snapshot(cell):
        return {
            'font': _copy_style(cell.font),
            'border': _copy_style(cell.border),
            'fill': _copy_style(cell.fill),
            'number_format': cell.number_format,
            'protection': _copy_style(cell.protection),
            'alignment': _copy_style(cell.alignment),
        }

    def _style_apply(cell, snap):
        cell.font = _copy_style(snap['font'])
        cell.border = _copy_style(snap['border'])
        cell.fill = _copy_style(snap['fill'])
        cell.number_format = snap['number_format']
        cell.protection = _copy_style(snap['protection'])
        cell.alignment = _copy_style(snap['alignment'])

    def _find_first_content_row(ws, start_row, max_scan=300):
        """从 start_row 向下找第一个有内容的行（返回行号，找不到返回 None）。"""
        max_col = ws.max_column or 10
        for r in range(start_row, start_row + max_scan):
            for c in range(1, max_col + 1):
                v = ws.cell(r, c).value
                if v is not None and str(v).strip():
                    return r
        return None

    def _shift_images_down(ws, end_row, extra):
        """insert_rows 之后调整图片/形状 anchor，让它们跟着内容整体下移。

        ★ 只要图片的 from.row 或 to.row 有一个 ≥ 插入点，整张图片一起 +extra。
        openpyxl 内部的行号是 0-based。
        """
        if extra <= 0:
            return
        end_row_0 = end_row - 1
        for img in getattr(ws, '_images', []) or []:
            anchor = getattr(img, 'anchor', None)
            if anchor is None:
                continue

            frm = getattr(anchor, '_from', None)
            to = getattr(anchor, 'to', None)

            need_shift = False
            if frm is not None:
                r = getattr(frm, 'row', None)
                if r is not None and r >= end_row_0:
                    need_shift = True
            if not need_shift and to is not None:
                r = getattr(to, 'row', None)
                if r is not None and r >= end_row_0:
                    need_shift = True

            if not need_shift:
                continue

            if frm is not None:
                try:
                    frm.row = frm.row + extra
                except Exception:
                    pass
            if to is not None:
                try:
                    to.row = to.row + extra
                except Exception:
                    pass

    def _ensure_passenger_rows(ws, data_start_row, needed_count):
        """确保乘客数据区至少有 needed_count 行；不够则在承诺行之前插入。

        - 复制某一行（优先选有合并结构的行）的样式 + 行高 + 合并到新行
        - 保存并恢复 end_row 及以下所有行的行高
        - 只处理受影响的合并单元格，其余不动
        - 图片整体下移，保持宽高比
        """
        if needed_count <= 0:
            return data_start_row

        from openpyxl.utils import get_column_letter

        end_row = _find_first_content_row(ws, data_start_row)
        if end_row is None:
            return data_start_row

        existing_rows = end_row - data_start_row
        if needed_count <= existing_rows:
            return data_start_row

        extra = needed_count - existing_rows
        template_row = end_row - 1

        max_col = ws.max_column or 10

        merge_template_row = template_row
        for r in range(template_row, data_start_row - 1, -1):
            hit = False
            for mr in ws.merged_cells.ranges:
                if mr.min_row == r and mr.max_row == r:
                    hit = True
                    break
            if hit:
                merge_template_row = r
                break

        template_styles = {
            c: _style_snapshot(ws.cell(template_row, c))
            for c in range(1, max_col + 1)
        }
        template_height = ws.row_dimensions[template_row].height

        template_merges = []
        for mr in list(ws.merged_cells.ranges):
            if mr.min_row == merge_template_row and mr.max_row == merge_template_row:
                template_merges.append((mr.min_col, mr.max_col))

        max_row = ws.max_row
        saved_heights = {}
        for r in range(end_row, max_row + 1):
            h = ws.row_dimensions[r].height
            if h is not None:
                saved_heights[r] = h

        affected = []
        for mr in list(ws.merged_cells.ranges):
            if mr.min_row >= end_row or mr.min_row < end_row <= mr.max_row:
                affected.append((mr.min_row, mr.min_col, mr.max_row, mr.max_col))
                try:
                    ws.unmerge_cells(str(mr))
                except Exception:
                    pass

        ws.insert_rows(end_row, extra)

        _shift_images_down(ws, end_row, extra)

        for min_r, min_c, max_r, max_c in affected:
            if min_r >= end_row:
                new_min_r, new_max_r = min_r + extra, max_r + extra
            else:
                new_min_r, new_max_r = min_r, max_r + extra
            rng = (f"{get_column_letter(min_c)}{new_min_r}:"
                   f"{get_column_letter(max_c)}{new_max_r}")
            try:
                ws.merge_cells(rng)
            except Exception:
                pass

        for r in list(ws.row_dimensions.keys()):
            if r >= end_row:
                try:
                    del ws.row_dimensions[r]
                except Exception:
                    pass
        for old_r, h in saved_heights.items():
            ws.row_dimensions[old_r + extra].height = h

        for i in range(extra):
            new_row = end_row + i
            for c, snap in template_styles.items():
                _style_apply(ws.cell(new_row, c), snap)
            if template_height:
                ws.row_dimensions[new_row].height = template_height
            for min_c, max_c in template_merges:
                rng = (f"{get_column_letter(min_c)}{new_row}:"
                       f"{get_column_letter(max_c)}{new_row}")
                try:
                    ws.merge_cells(rng)
                except Exception:
                    pass

        return data_start_row
    # ---------- ★ 新增结束 ----------

    def fill_template(template_bytes, data, crew_rows, passenger_list, route_display):
        try:
            wb = load_workbook(template_bytes)
        except Exception as e:
            if "Bad magic number" in str(e) or "BadZipFile" in str(e):
                st.error("❌ 模板文件格式不正确。请确保模板为 **.xlsx** 格式（非 .xls）。")
                st.info("💡 解决方法：用 Excel 打开该模板，选择“另存为”，将文件类型选为 **Excel工作簿（.xlsx）**，然后重新上传。")
            raise e

        ws = wb.active

        if not passenger_list:
            for row in ws.iter_rows(min_row=1, max_row=10):
                for cell in row:
                    if cell.value and isinstance(cell.value, str) and "飞行目的" in cell.value:
                        safe_set_cell_value(ws, cell.row + 1, 2, "调机")
                        break
                else:
                    continue
                break

        info_row = None
        for row in ws.iter_rows(min_row=1, max_row=20):
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    if cell.value.strip() in ["机型", "注册号", "航班号", "航班行程"]:
                        info_row = cell.row; break
            if info_row: break
        if info_row:
            data_row = info_row + 1
            safe_set_cell_value(ws, data_row, 2, data.get("ac_type", ""))
            safe_set_cell_value(ws, data_row, 3, data.get("reg", ""))
            safe_set_cell_value(ws, data_row, 4, data.get("flt", ""))
            safe_set_cell_value(ws, data_row, 5, route_display if route_display else "")

        role_map = {}
        for cr in crew_rows:
            role = str(cr.get("职务", "")).strip()
            name = str(cr.get("姓名", "")).strip()
            if role and name and role not in role_map:
                role_map[role] = cr

        def _is_valid(row):
            if row is None:
                return False
            name = str(row.get("姓名", "")).strip()
            return name != "" and name != "无"

        if _is_valid(role_map.get("机长")):
            _fill_crew_row_by_data(ws, "机长", role_map["机长"])
        else:
            _fill_empty_crew_row(ws, "机长")

        if _is_valid(role_map.get("副驾驶")):
            _fill_crew_row_by_data(ws, "副驾驶", role_map["副驾驶"])
        else:
            _fill_empty_crew_row(ws, "副驾驶")

        if _is_valid(role_map.get("乘务")):
            _fill_crew_row_by_data(ws, "乘务", role_map["乘务"])
        else:
            _fill_empty_crew_row(ws, "乘务")

        if _is_valid(role_map.get("机务")):
            _fill_crew_row_by_data(ws, "机务", role_map["机务"])
        else:
            _fill_empty_crew_row(ws, "机务")

        passenger_start_row = None
        for row in ws.iter_rows(min_row=1, max_row=100):
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    val = cell.value.strip()
                    if "姓名" in val and "性别" in val and "出生日期" in val:
                        passenger_start_row = cell.row + 1; break
            if passenger_start_row: break
        if passenger_start_row is None:
            for row in ws.iter_rows(min_row=1, max_row=100):
                for cell in row:
                    if cell.value and isinstance(cell.value, str) and "乘客信息" in cell.value:
                        passenger_start_row = cell.row + 2; break
                if passenger_start_row: break

        if passenger_start_row and passenger_list:
            passenger_start_row = _ensure_passenger_rows(
                ws, passenger_start_row, len(passenger_list)
            )

        if passenger_start_row:
            for i, pax in enumerate(passenger_list):
                row_num = passenger_start_row + i
                for col in range(1, 7):
                    safe_set_cell_value(ws, row_num, col, None)
                pax_name = extract_chinese_name(pax["name"])
                safe_set_cell_value(ws, row_num, 1, pax_name)
                _style_name_cell(ws, row_num, 1, pax_name)
                safe_set_cell_value(ws, row_num, 2, pax.get("gender", ""))
                safe_set_cell_value(ws, row_num, 3, pax.get("dob", ""))
                safe_set_cell_value(ws, row_num, 4, get_nation_name(pax.get("nationality", "")))
                doc_type = pax.get("doc_type", "")
                doc_type_clean = parse_document_type("", doc_type) if (pd.notna(doc_type) and str(doc_type).strip()) else parse_document_type(pax.get("passport_no", ""), "")
                safe_set_cell_value(ws, row_num, 5, doc_type_clean)
                safe_set_cell_value(ws, row_num, 6, pax.get("passport_no", ""))

        output = BytesIO()
        wb.save(output); output.seek(0)
        return output

    # ---------- ★ 缓存：避免每次 rerun 都重跑一遍生成 ----------
    @st.cache_data(show_spinner=False, max_entries=16)
    def _cached_fill_template(template_bytes, data_json, crew_json, pax_json, route_display):
        """把 fill_template 结果缓存。只有内容/输入变化才会重新生成。"""
        _data = _json.loads(data_json)
        _crew = _json.loads(crew_json)
        _pax = _json.loads(pax_json)
        buf = fill_template(template_bytes, _data, _crew, _pax, route_display)
        return buf.getvalue()
    # ---------- ★ 缓存结束 ----------

    st.subheader("📂 上传文件")
    st.info("⚠️ 注意：模板文件必须是 **.xlsx** 格式（非 .xls）。联系方式、执照号码及证件号码已内置，机型已内置，无需额外上传。")

    data_file = st.file_uploader(
        "上传 GD单（General Declaration）Excel（.xlsx）",
        type=["xlsx"], key="data"
    )
    template_file = st.file_uploader(
        "上传总调模板：Jetops申请一览-",
        type=["xlsx"], key="template"
    )
    flight_plan_file = st.file_uploader(
        "（可选）上传飞行计划（航段数据导出）（北京时间），用于预填文件名",
        type=["xlsx"], key="flight_plan"
    )

    if data_file and template_file:
        try:
            data, crew_list, passenger_list = parse_general_declaration(data_file)
            st.success(f"✅ 解析成功：机组 {len(crew_list)} 人，乘客 {len(passenger_list)} 人")

            if data.get("ac_type") or data.get("reg"):
                st.caption(
                    f"✈️ 注册号：**{data.get('reg', '')}** ｜ 机型：**{data.get('ac_type', '')}**"
                    + (f"（GD单原始机型：{data.get('ac_type_raw', '')}）"
                       if data.get('ac_type_raw') and data.get('ac_type_raw') != data.get('ac_type') else "")
                )

            st.subheader("📋 本次机组信息（可编辑）")
            st.caption(
                "系统已从内置名单匹配出**将要写入的证件号码 / 执照号码 / 联系方式**，"
                "如有出入可直接在表格里修改（也可改姓名、性别、出生日期或调整职务）。"
                "改完下方下载按钮生成的就是最新数据。"
            )

            role_order = ["机长", "副驾驶", "乘务", "机务"]
            role_rows = {r: None for r in role_order}
            overflow_rows = []

            for idx, crew in enumerate(crew_list):
                name_cn = extract_chinese_name(crew["name"])
                id_fill, license_fill = resolve_crew_id_and_license(
                    crew["name"], crew.get("passport_no", "")
                )
                contact = find_contact(crew["name"])

                if idx == 0:
                    preferred_role = "机长"
                elif idx == 1:
                    preferred_role = "副驾驶"
                else:
                    gender = str(crew.get("gender", "")).strip()
                    preferred_role = "乘务" if gender in ["女", "Female", "F"] else "机务"

                row = {
                    "职务": preferred_role,
                    "姓名": name_cn,
                    "性别": crew.get("gender", ""),
                    "出生日期": crew.get("dob", ""),
                    "证件号码": id_fill,
                    "执照号码": license_fill,
                    "联系方式": contact,
                }

                if role_rows.get(preferred_role) is None:
                    role_rows[preferred_role] = row
                else:
                    overflow_rows.append(row)

            initial_rows = []
            for role in role_order:
                if role_rows[role] is not None:
                    initial_rows.append(role_rows[role])
                else:
                    initial_rows.append({
                        "职务": role,
                        "姓名": "无",
                        "性别": "",
                        "出生日期": "",
                        "证件号码": "",
                        "执照号码": "",
                        "联系方式": "",
                    })
            initial_rows.extend(overflow_rows)

            if overflow_rows:
                overflow_desc = "、".join(
                    f"{r['姓名']}（当前标为「{r['职务']}」）" for r in overflow_rows
                )
                st.warning(
                    f"⚠️ **本次机组共 {len(crew_list)} 人**，而模板机组区只有 **4 行**"
                    f"（机长 / 副驾驶 / 乘务 / 机务）。以下人员**不会自动写入模板**：\n\n"
                    f"**{overflow_desc}**\n\n"
                    f"如需写入模板，请下载后在 Excel 中手动插入行。"
                )

            crew_signature = "|".join([c.get("name", "") for c in crew_list])
            editor_key = f"crew_fill_editor_{abs(hash(crew_signature)) % (10**8)}"

            df_fill = pd.DataFrame(initial_rows, columns=CREW_FILL_COLUMNS)
            edited_crew_df = st.data_editor(
                df_fill,
                num_rows="dynamic",
                use_container_width=True,
                height=280,
                key=editor_key,
                column_config={
                    "职务": st.column_config.SelectboxColumn(
                        "职务",
                        options=["机长", "副驾驶", "乘务", "机务"],
                        width="small",
                    ),
                    "姓名": st.column_config.TextColumn("姓名", width="medium"),
                    "性别": st.column_config.TextColumn("性别", width="small"),
                    "出生日期": st.column_config.TextColumn("出生日期", width="medium"),
                    "证件号码": st.column_config.TextColumn("证件号码", width="medium"),
                    "执照号码": st.column_config.TextColumn("执照号码", width="medium"),
                    "联系方式": st.column_config.TextColumn("联系方式", width="medium"),
                },
            )

            crew_for_template = (
                edited_crew_df.fillna("").astype(str).to_dict("records")
                if not edited_crew_df.empty else []
            )

            if passenger_list:
                st.caption(f"👥 本次乘客共 **{len(passenger_list)}** 人，模板行数不足时会自动插入行。")

            st.markdown("---")

            from_code = data.get("from", ""); to_code = data.get("to", "")
            date_str = data.get("date_str", ""); utc_time = data.get("utc_time", "")
            reg = data.get("reg", "")
            date_display = get_beijing_date_display(utc_time, date_str) if date_str else ""

            default_route = ""
            matched_note = ""
            gd_date_parsed = _parse_gd_date(date_str)
            if flight_plan_file is not None and reg and from_code and to_code:
                route_df = parse_route_plan(flight_plan_file)
                if route_df is not None:
                    matched_flight = find_matching_flight(
                        route_df, reg, from_code, to_code, gd_date=gd_date_parsed
                    )
                    if matched_flight is not None:
                        built = build_route_from_flight(matched_flight)
                        if built:
                            default_route = built
                            flight_no = str(matched_flight.get('航班号', '') or '').strip()
                            row_date = _extract_row_date(matched_flight.get('出发日期'))
                            date_hint = row_date.strftime("%m-%d") if row_date else "?"
                            matched_note = (
                                f"✅ 已从航段数据自动匹配到飞行计划（航班号 {flight_no}，"
                                f"{from_code} → {to_code}，{date_hint}），已预填到下方输入框"
                            )

            if not default_route:
                if date_str and from_code and to_code:
                    bj_time = parse_utc_to_beijing(utc_time, date_str) if utc_time else "0000"
                    default_route = f"{date_display} {from_code} {bj_time} XXXX {to_code}"
                else:
                    default_route = f"{from_code}-{to_code}" if from_code and to_code else ""

            if matched_note:
                st.success(matched_note)
            elif flight_plan_file is not None and reg and from_code and to_code:
                date_hint = f"{gd_date_parsed[0]}月{gd_date_parsed[1]}日" if gd_date_parsed else "未知日期"
                st.info(
                    f"ℹ️ 航段数据中未找到 {reg} {from_code}→{to_code}（{date_hint}）的匹配记录，"
                    f"已使用 GD单 信息生成默认值。"
                )

            raw_route = st.text_input(
                "从Jetops复制航班信息并适当调整起落时间 比如： F B652S 08:00 - 14:00  柬埔寨金边 德崇 - 日本东京 羽田",
                value=default_route
            ).strip()

            with_time = parse_with_time_route(raw_route, date_display)
            if with_time != raw_route and with_time is not None:
                route_display = with_time
            else:
                route_display = raw_route if raw_route else default_route

            no_time = parse_no_time_route(raw_route, date_display)
            if no_time is not None:
                file_name_base = no_time
            else:
                reg_fallback = data.get("reg", "")
                if reg_fallback and from_code and to_code:
                    file_name_base = f"{date_display} {reg_fallback} {from_code}-{to_code}"
                else:
                    file_name_base = route_display

            safe_file_name = re.sub(r'[\\/*?:"<>|]', "_", file_name_base).strip()
            if not safe_file_name:
                safe_file_name = "备案表"
            download_file_name = f"{safe_file_name}.xlsx"

            # ★ 用缓存版本；内容不变时秒出，不再重复跑 fill_template
            try:
                result_bytes = _cached_fill_template(
                    template_file.getvalue(),
                    _json.dumps(data, ensure_ascii=False, default=str),
                    _json.dumps(crew_for_template, ensure_ascii=False, default=str),
                    _json.dumps(passenger_list, ensure_ascii=False, default=str),
                    route_display,
                )
            except Exception:
                # 缓存失败 fallback 到直调
                result_bytes = fill_template(
                    template_file, data, crew_for_template, passenger_list, route_display
                ).getvalue()

            st.download_button(
                label="⬇️ 下载填充后的备案表",
                data=result_bytes,
                file_name=download_file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.caption("⏳ 输入文件名之后稍等3秒钟后再点击下载")

        except Exception as e:
            st.error(f"❌ 处理失败：{e}")
            st.exception(e)
    else:
        st.info("👆 请同时上传 GD单 和 模板文件。")

# ================================================================
# 功能2：世界时行程
# ================================================================
with tab2:
    F_HTML = r"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8"><title>世界时行程转换</title>
<script src="https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js"></script>
<style>
body{font-family:-apple-system,"Segoe UI","Microsoft YaHei",sans-serif;margin:12px;color:#333;font-size:16px;}
.upload-hint{font-size:15px;color:#555;margin:4px 0 6px 0;}
input[type=file]{padding:6px;font-size:15px;}
button{padding:9px 16px;font-size:15px;border-radius:6px;border:1px solid #ddd;background:#fff;cursor:pointer;margin-right:6px;margin-top:6px;}
button:hover{background:#f5f5f5;}
.reg-block{margin-bottom:20px;}
.reg-title{font-weight:bold;font-size:19px;margin-bottom:8px;}
.new-flag{color:#d32f2f;font-size:1rem;margin-left:8px;font-weight:normal;}
.seg-list{background:#f7f7f7;border:1px solid #ddd;border-radius:6px;padding:10px 14px;font-family:Consolas,"Courier New",monospace;font-size:16px;line-height:1.9;color:#222;}
.seg-line{padding:3px 0;}
.status{color:#555;font-size:15px;margin-left:8px;}
.error{color:#d32f2f;background:#ffebee;padding:8px;border-radius:4px;margin:6px 0;}
.success{color:#2e7d32;background:#e8f5e9;padding:8px;border-radius:4px;margin:6px 0;}
.info{color:#1976d2;background:#e3f2fd;padding:8px;border-radius:4px;margin:6px 0;}
details{margin:10px 0;padding:8px;border:1px solid #eee;border-radius:4px;background:#fafafa;}
summary{cursor:pointer;font-weight:bold;padding:4px 0;font-size:16px;}
ol{margin:6px 0 6px 20px;padding:0;} li{margin:2px 0;font-size:15px;}
.full-text-box{background:#f5f5f5;padding:10px;border-radius:4px;font-family:Consolas,"Courier New",monospace;font-size:16px;white-space:pre;overflow-x:auto;border:1px solid #e0e0e0;max-height:400px;overflow-y:auto;}
</style>
</head>
<body>
<div class="upload-hint">📤 上传未来航段（北京时间）：</div>
<input type="file" id="fileInput" accept=".xlsx,.xls">
<div id="status"></div>
<div id="result" style="display:none;">
<div id="plans"></div>
<details><summary>📦 全部计划合并（点击展开）</summary><div class="full-text-box" id="fullTextBox"></div></details>
<details><summary>📜 历史记录</summary><div id="historyList"></div><button id="clearHistoryBtn" style="margin-top:8px;">🗑️ 清除所有历史</button></details>
</div>
<script>
const MONTHS=['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
const PRIORITY=['B652Q','B65AP','B652S','MLLIN','N88AY','B652R'];
const HISTORY_KEY='worldtime_history_v2',LAST_PLANS_KEY='worldtime_last_plans_v2',LAST_FILE_KEY='worldtime_last_file_v2';
function loadHistory(){try{const r=localStorage.getItem(HISTORY_KEY);if(!r)return{records:[]};const o=JSON.parse(r);if(!o.records)o.records=[];return o;}catch(e){return{records:[]};}}
function saveHistory(h){try{localStorage.setItem(HISTORY_KEY,JSON.stringify(h));}catch(e){}}
function loadJSON(k){try{const r=localStorage.getItem(k);return r?JSON.parse(r):null;}catch(e){return null;}}
function saveJSON(k,v){try{localStorage.setItem(k,JSON.stringify(v));}catch(e){}}
function pad2(n){return String(n).padStart(2,'0');}
function escapeHtml(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
function parseDate(v){if(v==null||v==='')return null;if(v instanceof Date)return new Date(v.getFullYear(),v.getMonth(),v.getDate());if(typeof v==='number'){const ms=Math.round((v-25569)*86400*1000);const d=new Date(ms);return new Date(d.getUTCFullYear(),d.getUTCMonth(),d.getUTCDate());}const s=String(v).trim();if(!s)return null;const m=s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);if(m)return new Date(parseInt(m[1]),parseInt(m[2])-1,parseInt(m[3]));const d=new Date(s);if(!isNaN(d.getTime()))return new Date(d.getFullYear(),d.getMonth(),d.getDate());return null;}
function parseTime(v){if(v==null||v==='')return null;if(v instanceof Date)return pad2(v.getHours())+':'+pad2(v.getMinutes());if(typeof v==='number'){let f=v;if(v>1)f=v-Math.floor(v);const t=Math.round(f*24*60);return pad2(Math.floor(t/60)%24)+':'+pad2(t%60);}const s=String(v).trim();const m=s.match(/(\d{1,2}):(\d{2})/);if(m)return pad2(parseInt(m[1]))+':'+m[2];return null;}
function toUTCLabel(dateVal,timeStr){const d=parseDate(dateVal);if(!d||!timeStr)return null;const p=timeStr.split(':');const h=parseInt(p[0]),m=parseInt(p[1]);let t=h*60+m-8*60;let off=0;while(t<0){t+=24*60;off--;}while(t>=24*60){t-=24*60;off++;}const uh=Math.floor(t/60),um=t%60;const dd=new Date(d.getFullYear(),d.getMonth(),d.getDate());dd.setDate(dd.getDate()+off);return{day:dd.getDate(),month:dd.getMonth()+1,hours:uh,minutes:um,sortKey:dd.getFullYear()*100000000+(dd.getMonth()+1)*1000000+dd.getDate()*10000+uh*100+um};}
function formatLabel(u){if(!u)return'';return pad2(u.day)+MONTHS[u.month-1]+' '+pad2(u.hours)+pad2(u.minutes)+'Z';}
function processRows(rows){let h=-1;for(let i=0;i<Math.min(rows.length,10);i++){const v=rows[i].map(x=>String(x).trim());if(v.includes('飞机注册号')&&v.includes('出发地')&&v.includes('到达地')&&v.includes('计划出发')){h=i;break;}}if(h===-1)return{error:'未找到表头行'};const hs=rows[h].map(x=>String(x).trim());function fc(c){for(const n of c){const i=hs.indexOf(n);if(i!==-1)return i;}for(const n of c){for(let i=0;i<hs.length;i++){if(hs[i].includes(n))return i;}}return -1;}const colReg=fc(['飞机注册号','注册号','机号']),colDep=fc(['出发地']),colArr=fc(['到达地']),colDepDate=fc(['出发日期']),colDepTime=fc(['计划出发']),colArrDate=fc(['到达日期']),colArrTime=fc(['预计到达']),colPurpose=fc(['用途']);const req={'飞机注册号':colReg,'出发地':colDep,'到达地':colArr,'出发日期':colDepDate,'计划出发':colDepTime,'到达日期':colArrDate,'预计到达':colArrTime,'用途':colPurpose};for(const n in req){if(req[n]===-1)return{error:'缺少列：'+n};}const plans={};for(let i=h+1;i<rows.length;i++){const r=rows[i];if(!r||r.length===0)continue;const g=i=>i>=0&&i<r.length?r[i]:'';const dep=g(colDep),arr=g(colArr),depDate=g(colDepDate),depTimeRaw=g(colDepTime),arrDate=g(colArrDate),arrTimeRaw=g(colArrTime);if(dep===''||arr===''||depDate===''||depTimeRaw==='')continue;const dts=parseTime(depTimeRaw),ats=parseTime(arrTimeRaw);if(!dts||!ats)continue;const du=toUTCLabel(depDate,dts),au=toUTCLabel(arrDate,ats);if(!du||!au)continue;let reg=g(colReg);if(reg===''||reg==null)reg='N/A';else reg=String(reg).trim();const use=String(g(colPurpose)||'');const ft=use.indexOf('调机')!==-1?'FERRY':'PAX';const line='ETD '+String(dep).trim()+' '+formatLabel(du)+' // ETA '+String(arr).trim()+' '+formatLabel(au)+'  '+ft;if(!plans[reg])plans[reg]=[];plans[reg].push({sortKey:du.sortKey,line:line});}const res={};for(const reg in plans){plans[reg].sort((a,b)=>a.sortKey-b.sortKey);const ls=[reg];plans[reg].forEach(it=>ls.push(it.line));res[reg]=ls.join('\n');}return{plans:res};}
function sortPlans(p){const ks=Object.keys(p);const pk=PRIORITY.filter(k=>ks.indexOf(k)!==-1);const rk=ks.filter(k=>PRIORITY.indexOf(k)===-1&&k!=='N/A').sort();const nk=ks.filter(k=>k==='N/A');const so=pk.concat(rk,nk);const r={};so.forEach(k=>r[k]=p[k]);return r;}
function diffPlans(o,n){const c={};const all=new Set(Object.keys(o||{}).concat(Object.keys(n||{})));all.forEach(reg=>{const ol=new Set(((o&&o[reg])||'').split('\n'));const nl=new Set(((n&&n[reg])||'').split('\n'));ol.delete(reg);nl.delete(reg);nl.forEach(l=>{if(!ol.has(l)){c[reg+'\u0001'+l]='added';}});});return c;}
function renderPlans(plans,changes,restored){const con=document.getElementById('plans');con.innerHTML='';let ft='';for(const reg in plans){const text=plans[reg];const lines=text.split('\n');const routes=lines.filter(l=>l!==reg);const hc=!restored&&routes.some(l=>changes[reg+'\u0001'+l]);const b=document.createElement('div');b.className='reg-block';const t=document.createElement('div');t.className='reg-title';t.innerHTML='✈️ '+escapeHtml(reg)+(hc?'<span class="new-flag">🔴 有新增或变更</span>':'');b.appendChild(t);const sl=document.createElement('div');sl.className='seg-list';routes.forEach(l=>{const d=document.createElement('div');d.className='seg-line';d.textContent=l;sl.appendChild(d);});b.appendChild(sl);const cb=document.createElement('button');cb.textContent='📋 复制该飞机';cb.onclick=()=>{const ct=reg+'\n'+routes.join('\n');navigator.clipboard.writeText(ct).then(()=>{cb.textContent='✅ 已复制';setTimeout(()=>{cb.textContent='📋 复制该飞机';},1500);}).catch(()=>{fb(ct);cb.textContent='✅ 已复制';setTimeout(()=>{cb.textContent='📋 复制该飞机';},1500);});};b.appendChild(cb);con.appendChild(b);ft+=reg+'\n'+routes.join('\n')+'\n\n';}document.getElementById('fullTextBox').textContent=ft.trim();}
function renderHistory(h){const c=document.getElementById('historyList');if(!h.records||h.records.length===0){c.innerHTML='<div class="info">暂无历史记录</div>';return;}let html='<ol>';for(const r of h.records){html+='<li>'+escapeHtml(r.timestamp)+' - '+escapeHtml(r.filename)+'</li>';}html+='</ol>';c.innerHTML=html;}
function fb(t){const ta=document.createElement('textarea');ta.value=t;ta.style.position='fixed';ta.style.left='-9999px';document.body.appendChild(ta);ta.select();try{document.execCommand('copy');}catch(e){}document.body.removeChild(ta);}
function handleFile(file){const s=document.getElementById('status');s.innerHTML='<div class="info">⏳ 正在读取文件...</div>';const rd=new FileReader();rd.onload=(ev)=>{try{const data=new Uint8Array(ev.target.result);const wb=XLSX.read(data,{type:'array',cellDates:true});const ws=wb.Sheets[wb.SheetNames[0]];const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:true});const r=processRows(rows);if(r.error){s.innerHTML='<div class="error">❌ '+escapeHtml(r.error)+'</div>';return;}const np=r.plans;const sp=sortPlans(np);const h=loadHistory();let op={};if(h.records.length>0){op=h.records[h.records.length-1].data||{};}const ch=diffPlans(op,np);const now=new Date();const ts=now.getFullYear()+'-'+pad2(now.getMonth()+1)+'-'+pad2(now.getDate())+' '+pad2(now.getHours())+':'+pad2(now.getMinutes())+':'+pad2(now.getSeconds());h.records.push({timestamp:ts,filename:file.name,data:np});if(h.records.length>20){h.records=h.records.slice(-20);}saveHistory(h);saveJSON(LAST_PLANS_KEY,sp);saveJSON(LAST_FILE_KEY,{name:file.name,timestamp:ts});document.getElementById('result').style.display='block';s.innerHTML='<div class="success">✅ 文件读取成功：'+escapeHtml(file.name)+'（'+ts+'，历史累计 '+h.records.length+' 条）</div>';renderPlans(sp,ch,false);renderHistory(h);}catch(err){s.innerHTML='<div class="error">❌ 处理失败：'+escapeHtml(err.message)+'</div>';console.error(err);}};rd.readAsArrayBuffer(file);}
window.addEventListener('DOMContentLoaded',()=>{const lp=loadJSON(LAST_PLANS_KEY);const lf=loadJSON(LAST_FILE_KEY);const h=loadHistory();if(lp&&Object.keys(lp).length>0){document.getElementById('result').style.display='block';const s=document.getElementById('status');const i=lf?'（上次加载：'+escapeHtml(lf.name)+'，'+escapeHtml(lf.timestamp)+'）':'';s.innerHTML='<div class="info">💾 已恢复上次解析结果 '+i+'</div>';renderPlans(lp,{},true);renderHistory(h);}});
document.getElementById('fileInput').addEventListener('change',(e)=>{const f=e.target.files[0];if(f)handleFile(f);});
document.getElementById('clearHistoryBtn').addEventListener('click',()=>{if(!confirm('确定清除所有历史记录吗？'))return;saveHistory({records:[]});try{localStorage.removeItem(LAST_PLANS_KEY);localStorage.removeItem(LAST_FILE_KEY);}catch(e){}location.reload();});
</script>
</body>
</html>
"""
    components.html(F_HTML, height=1000, scrolling=True)

# ================================================================
# 功能3：航路处理工具
# ================================================================
with tab3:
    G_HTML = r"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8"><title>航路处理工具</title>
<style>
body{font-family:-apple-system,"Segoe UI","Microsoft YaHei",sans-serif;margin:12px;color:#333;font-size:16px;}
textarea{width:100%;height:300px;box-sizing:border-box;font-family:Consolas,"Courier New",monospace;font-size:15px;padding:10px;border:1px solid #ccc;border-radius:6px;line-height:1.6;}
button{padding:10px 20px;font-size:16px;border-radius:6px;border:1px solid #ddd;background:#fff;cursor:pointer;margin-right:8px;}
button.primary{background:#ff4b4b;color:#fff;border-color:#ff4b4b;font-weight:bold;}
button.primary:hover{background:#e63939;}
button:hover{background:#f5f5f5;}
.toolbar{margin:12px 0;}
.result-box{background:#f5f5f5;border:1px solid #e0e0e0;border-radius:6px;padding:14px;font-family:Consolas,"Courier New",monospace;font-size:16px;line-height:1.7;white-space:pre-wrap;word-break:break-all;max-height:500px;overflow-y:auto;}
.success{color:#2e7d32;background:#e8f5e9;padding:8px;border-radius:4px;margin:8px 0;font-size:15px;}
.error{color:#d32f2f;background:#ffebee;padding:8px;border-radius:4px;margin:8px 0;font-size:15px;}
.info{color:#1976d2;background:#e3f2fd;padding:8px;border-radius:4px;margin:8px 0;font-size:15px;}
.label{color:#555;font-size:15px;margin:6px 0;}
</style>
</head>
<body>
<div class="label">📋 请输入待处理的航路文本</div>
<textarea id="inputText" placeholder="粘贴民航航线数据，支持多行表格格式/纯中文描述格式..."></textarea>
<div class="toolbar">
<button class="primary" id="processBtn">⚙️ 处理</button>
<button id="clearBtn">🗑️ 清空</button>
<button id="copyBtn">📋 复制结果</button>
<span id="copyStatus" style="margin-left:8px;color:#2e7d32;font-size:15px;"></span>
</div>
<div id="status"></div>
<div id="resultSection" style="display:none;">
<div class="label">📊 处理结果</div>
<div class="result-box" id="resultBox"></div>
</div>
<script>
const INPUT_CACHE_KEY='route_input_cache_v1';
function saveInput(t){try{localStorage.setItem(INPUT_CACHE_KEY,t);}catch(e){}}
function loadInput(){try{return localStorage.getItem(INPUT_CACHE_KEY)||'';}catch(e){return'';}}
function escapeHtml(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
function parseCoord(c){const l=c[0],n=c.slice(1);if(l==='N'){let d=parseInt(n.slice(0,2)),m=parseInt(n.slice(2,4)),sp=n.slice(4),si;if(sp.indexOf('.')!==-1)si=Math.round(parseFloat(sp));else si=parseInt(sp);if(si>=60){si-=60;m+=1;if(m>=60){m-=60;d+=1;}}return String(d).padStart(2,'0')+String(m).padStart(2,'0')+String(si).padStart(2,'0');}else if(l==='E'){let d=parseInt(n.slice(0,3)),m=parseInt(n.slice(3,5)),sp=n.slice(5),si;if(sp.indexOf('.')!==-1)si=Math.round(parseFloat(sp));else si=parseInt(sp);if(si>=60){si-=60;m+=1;if(m>=60){m-=60;d+=1;}}return String(d).padStart(3,'0')+String(m).padStart(2,'0')+String(si).padStart(2,'0');}throw new Error('未知的坐标前缀: '+l);}
function baseName(s){return s.split('@')[0];}
function isOpenPoint(s){const b=baseName(s);if(/^[A-Z]{2,5}$/.test(b))return true;if(/^P[A-Z]+$/.test(b))return true;return false;}
function isPPoint(s){const b=baseName(s);return /^P\d+$/.test(b);}
function cleanRoute(r){if(r.startsWith('#'))return r.slice(1);return r;}
function isOpenRoute(rt){return rt&&(rt[0]!=='H'&&rt[0]!=='J'&&rt[0]!=='V');}
function isClosedRoute(rt){return rt.startsWith('H')||rt.startsWith('J')||rt.startsWith('V');}
function extractTable(text){let t=text.trim().split(/\s+/);let si=0;for(let i=0;i<t.length;i++){if(/^\d+$/.test(t[i])&&parseInt(t[i])>=1&&parseInt(t[i])<=40){si=i;break;}}t=t.slice(si);const lines=[];let i=0;while(i<t.length){if(/^\d+$/.test(t[i])){const l=[t[i]];i++;while(i<t.length&&!/^\d+$/.test(t[i])){l.push(t[i]);i++;}lines.push(l);}else{i++;}}const pts=[],rts=[];for(const l of lines){let la=-1;for(let k=0;k<l.length;k++){if(l[k].startsWith('N')&&/^\d+(\.\d+)?$/.test(l[k].slice(1))){la=k;break;}}if(la===-1)continue;const lo=la+1;if(lo>=l.length||!l[lo].startsWith('E'))continue;const lat=l[la],lon=l[lo];let rt=null;if(lo+1<l.length){const nt=l[lo+1];if(/^[A-Z][A-Z0-9]*$/.test(nt)&&!/^\d/.test(nt[0]))rt=nt;}let pn=null;for(let j=la-1;j>0;j--){if(isOpenPoint(l[j])||isPPoint(l[j])){pn=l[j];break;}}if(pn===null)continue;let pd;if(isPPoint(pn)){const li=parseCoord(lat),loi=parseCoord(lon);pd=pn+'@'+li+'N'+loi+'E';}else{pd=pn;}pts.push(pd);if(rt!==null)rts.push(rt);}const seq=[];for(let i=0;i<pts.length;i++){seq.push(pts[i]);if(i<rts.length)seq.push(rts[i]);}return seq;}
function extractChinese(text){text=text.replace(/[\u4e00-\u9fa5，、。；：""''（）【】]/g,' ');const ws=text.split(/\s+/).filter(w=>w);const seq=[];for(const w of ws){if(w.indexOf('(')!==-1&&w.indexOf(')')!==-1){const m=w.match(/\(([A-Z]+)\)/);if(m){const p=m[1];const pr=w.slice(0,w.indexOf('('));const mr=pr.match(/([A-Z]\d+)$/);if(mr)seq.push(mr[1]);seq.push(p);}}else if(/^[A-Z]\d+[A-Z]{2,5}$/.test(w)||/^[A-Z]\d+P\d+$/.test(w)){const m=w.match(/^([A-Z]\d+)([A-Z]{2,5}|P\d+)$/);if(m){seq.push(m[1]);seq.push(m[2]);}}else if(/^[A-Z]\d+$/.test(w)){seq.push(w);}else if(isOpenPoint(w)||isPPoint(w)){seq.push(w);}}return seq;}
function step1Extract(t){if(/N\d{5,6}(\.\d+)?\s+E\d{6,7}(\.\d+)?/.test(t)){return{seq:extractTable(t),fmt:'table'};}else{return{seq:extractChinese(t),fmt:'chinese'};}}
function step2Reduce(s){let L=s.slice();let ch=true;while(ch){ch=false;const n=L.length;const c=[];for(let i=0;i<n;i+=2){if(!isOpenPoint(L[i]))continue;if(i+1>=n)continue;const fr=cleanRoute(L[i+1]);if(!isOpenRoute(fr))continue;for(let j=i+2;j<n;j+=2){let all=true;for(let k=i+1;k<j;k+=2){const rt=cleanRoute(L[k]);if(rt!==fr||!isOpenRoute(rt)){all=false;break;}}if(!all)break;if(isOpenPoint(L[j])){const len=Math.floor((j-i)/2);if(len>=2)c.push([i,j,len]);}}}if(c.length===0)break;c.sort((a,b)=>b[2]-a[2]);const[bi,bj]=c[0];const ns=[L[bi],L[bi+1],L[bj]];L=L.slice(0,bi).concat(ns).concat(L.slice(bj+1));ch=true;}return L;}
function step3AddHash(s){const p=s.filter((_,i)=>i%2===0),r=s.filter((_,i)=>i%2===1);const res=[p[0]];for(let i=0;i<r.length;i++){const rt=r[i],lf=p[i],rg=p[i+1];let nh=false;if(isClosedRoute(rt))nh=true;else if(isPPoint(lf)||isPPoint(rg))nh=true;res.push(nh?'#'+rt:rt);res.push(rg);}return res;}
function process(){const s=document.getElementById('status');const it=document.getElementById('inputText').value;if(!it.trim()){s.innerHTML='<div class="error">请输入待处理的航路文本</div>';document.getElementById('resultSection').style.display='none';return;}try{const{seq:s1,fmt}=step1Extract(it);let seq=s1;if(fmt==='table'){seq=step2Reduce(seq);seq=step3AddHash(seq);}const res=seq.length>0?seq.join(' '):'⚠️ 未提取到有效航路数据';document.getElementById('resultBox').textContent=res;document.getElementById('resultSection').style.display='block';s.innerHTML='<div class="success">✅ 处理完成</div>';}catch(e){s.innerHTML='<div class="error">❌ 处理失败：'+escapeHtml(e.message)+'</div>';console.error(e);}}
const ie=document.getElementById('inputText');let st=null;ie.addEventListener('input',()=>{clearTimeout(st);st=setTimeout(()=>saveInput(ie.value),500);});
document.getElementById('processBtn').addEventListener('click',process);
document.getElementById('clearBtn').addEventListener('click',()=>{ie.value='';saveInput('');document.getElementById('resultSection').style.display='none';document.getElementById('status').innerHTML='';document.getElementById('copyStatus').textContent='';});
document.getElementById('copyBtn').addEventListener('click',async()=>{const t=document.getElementById('resultBox').textContent;if(!t)return;const se=document.getElementById('copyStatus');try{await navigator.clipboard.writeText(t);se.textContent='✅ 已复制';setTimeout(()=>{se.textContent='';},1500);}catch(e){const ta=document.createElement('textarea');ta.value=t;ta.style.position='fixed';ta.style.left='-9999px';document.body.appendChild(ta);ta.select();try{document.execCommand('copy');se.textContent='✅ 已复制';setTimeout(()=>{se.textContent='';},1500);}catch(e2){se.textContent='❌ 复制失败';se.style.color='#d32f2f';}document.body.removeChild(ta);}});
window.addEventListener('DOMContentLoaded',()=>{const s=loadInput();if(s)ie.value=s;});
</script>
</body>
</html>
"""
    components.html(G_HTML, height=900, scrolling=True)

# ================================================================
# 功能4：批复核对
# ================================================================
with tab4:
    st.markdown("上传批复汇总表 + 航段数据，粘贴文本航班信息，即可自动核对差异。")

    PC_MONTHS = {
        "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4,
        "MAY": 5, "JUN": 6, "JUL": 7, "AUG": 8,
        "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12,
    }

    PC_RED = "FF0000"
    PC_GREEN = "00B050"
    PC_HIGHLIGHT_ADDED = "blue"

    PC_MAX_CROSS_DAY_GAP_MIN = 600
    PC_EARLY_GREEN_THRESHOLD_MIN = 600

    PC_PILOT_RAW = """P001,庚凡,gengfan@amber-aviation.com
P002,张永一,zhangyongyi@amber-aviation.com
P003,梅峰,fmei@amber-aviation.com
P004,王斌,wangbin@amber-aviation.com
P019,"HEALY, Darran William",darranhealy@amber-aviation.com
P020,"BEEBE, Thaddeus John",thaddeusbeebe@amber-aviation.com
P032,林毅,ericlin@amber-aviation.com
P035,"Peter Robert, JACKSON",prjackson@amber-aviation.com
P036,王少雄,warrenwang@amber-aviation.com
P038,苗旺旺,johnmiao@amber-aviation.com
P039,"Yiftah, RAUCH",yiftahrauch@amber-aviation.com
P044,李辛欣,rockli@amber-aviation.com
P046,赵岩松,zhyszhao@amber-aviation.com
P051,彭罡,eugene.peng@humbleholding.com
P052,胡君量,brian.wu@humbleholding.com
P053,"Bruce Roderick, WAINES",brwaines@amber-aviation.com
P054,"Rodolfo, BONETTI",rbonetti@amber-aviation.com
P056,"Keith Robert, SHERREN",krsherren@amber-aviation.com
P057,"Oliver Viktor, RACZ",ovracz@amber-aviation.com
P059,蔡国俊,kctsai@amber-aviation.com
P061,李庆宏,qhli@amber-aviation.com
P065,宋炜,wsong@amber-aviation.com
P068,昝昭君,zjzan@amber-aviation.com
P069,"ROEDER, SIMONE ELKE",simoneroeder@amber-aviation.com
P070,"Herve Daniel, STAMM",hdstamm@amber-aviation.com
P071,孙浩,jasonsun@amber-aviation.com
P072,朱正宇,zyzhu@amber-aviation.com
P074,金尚明,smjin@amber-aviation.com
P075,"Eduard Pascal, Roski",eduardroski@amber-aviation.com
P077,刘凯,andyliu@amber-aviation.com
P078,张帆,fzhang@amber-aviation.com
P079,魏思远,wesleywei@amber-aviation.com
P080,刘爽,sliu@amber-aviation.com
P081,吴鹏,richardwu@amber-aviation.com
P082,刘汇川,frankliu@amber-aviation.com
P083,尤欣,xyou@amber-aviation.com
P084,李亚民,ymli@amber-aviation.com
P085,赵镭,lzhao@amber-aviation.com
P086,张贺新,hxzhang@amber-aviation.com
P087,孙赫,hesun@amber-aviation.com
P088,马坚,harryma@amber-aviation.com
P089,李晓龙,xlli@amber-aviation.com
P090,黄海东,hdhuang@amber-aviation.com
P091,马洪双,mikema@amber-aviation.com
PJZ001,张哲,zzhang@amber-aviation.com
PJZ002,郭春旭,charlesguo@amber-aviation.com
PJZ004,王国勤,leowang@amber-aviation.com
PJZ005,王莹,evawang@amber-aviation.com
PJZ007,徐卓,frankxu@amber-aviation.com
PJZ008,杨华,ariayang@amber-aviation.com
W070,王彦海,wang_yanhai@163.com
W213,沈志伟,cshum@tagaviation.com
W267,"Nathon Andrew G, NORBERG",naten7@hotmail.com
W268,"Daniel, RICHTER",pilotlocalizer@gmail.com
W270,杨涛,yang_tao2005@aliyun.com
W272,"Andrew Nigel, KING",Andrew.king@aero.bombardier.com
"""

    PC_AIRCRAFT_TYPE_MAP = {
        "B3926": "LJ60", "B652R": "GLF4", "B8105": "GLEX", "B8160": "GLF5",
        "B8262": "GLF4", "B8292": "GLF5", "B8309": "GLF5", "MLLIN": "GLEX",
        "N2QE": "GL5T", "N328LM": "GL7T", "N550DR": "GLF5", "N577QT": "F900",
        "N7777U": "GLEX", "N777ZH": "GLF5", "N88AY": "GLF5", "T7178HT": "GL7T",
        "T7CJK": "GLEX", "VPCSZ": "GL7T", "VPCVA": "GLF6", "B652Q": "GLF4",
        "B652S": "GLF4", "B65AP": "GLF4",
    }

    PC_FERRY_KEYWORDS = ("调机", "维修")

    # ---------- 工具 ----------
    def pc_parse_date_token(token):
        token = token.strip().upper()
        day = int(token[:2])
        mon = PC_MONTHS[token[2:5]]
        year_str = token[5:]
        year = 2000 + int(year_str) if len(year_str) == 2 else int(year_str)
        return _pcdt.date(year, mon, day)

    def pc_parse_hhmm(token):
        token = str(token).strip().zfill(4)
        return _pcdt.time(int(token[:2]), int(token[2:]))

    def pc_parse_hhmm_str(s):
        if not s:
            return None
        m = re.match(r"^(\d{1,2}):(\d{2})$", str(s).strip())
        if not m:
            return None
        return _pcdt.time(int(m.group(1)), int(m.group(2)))

    def pc_is_b_reg(reg):
        return str(reg).strip().upper().startswith("B")

    def pc_to_beijing_datetime(reg, date_obj, hhmm_token):
        dt = _pcdt.datetime.combine(date_obj, pc_parse_hhmm(hhmm_token))
        if not pc_is_b_reg(reg):
            dt += _pcdt.timedelta(hours=8)
        return dt

    def pc_fmt_time(value):
        if value is None:
            return ""
        if isinstance(value, _pcdt.datetime):
            return value.strftime("%H:%M")
        if isinstance(value, _pcdt.time):
            return value.strftime("%H:%M")
        s = str(value).strip()
        m = re.match(r"(\d{1,2}):(\d{2})", s)
        return f"{int(m.group(1)):02d}:{m.group(2)}" if m else s

    def pc_fmt_date(value):
        if value is None:
            return None
        if isinstance(value, _pcdt.datetime):
            return value.date()
        if isinstance(value, _pcdt.date):
            return value
        s = str(value).strip()
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
            try:
                return _pcdt.datetime.strptime(s, fmt).date()
            except ValueError:
                pass
        return None

    def pc_has_chinese(text):
        return bool(re.search(r"[\u4e00-\u9fff]", str(text)))

    def pc_time_diff_minutes(t1, t2):
        def to_min(t):
            h, m = t.split(":")
            return int(h) * 60 + int(m)
        try:
            return abs(to_min(t1) - to_min(t2))
        except Exception:
            return 9999

    def pc_hhmm_to_minutes(t):
        h, m = t.split(":")
        return int(h) * 60 + int(m)

    def pc_flight_duration_minutes(dep_t, arr_t):
        try:
            d = pc_hhmm_to_minutes(dep_t)
            a = pc_hhmm_to_minutes(arr_t)
        except Exception:
            return None
        if a < d:
            a += 1440
        return a - d

    def pc_fmt_duration(mins):
        if mins is None:
            return ""
        sign = "-" if mins < 0 else ""
        m = abs(mins)
        return f"{sign}{m // 60}:{m % 60:02d}"

    def pc_compute_real_minute_diff(ap_date, ap_time_str, xl_date, xl_time_str):
        if ap_date is None or xl_date is None:
            return None
        try:
            ap_min = pc_hhmm_to_minutes(ap_time_str)
            xl_min = pc_hhmm_to_minutes(xl_time_str)
        except Exception:
            return None
        ap_abs = ap_date.toordinal() * 1440 + ap_min
        xl_abs = xl_date.toordinal() * 1440 + xl_min
        return ap_abs - xl_abs

    @st.cache_data
    def pc_load_pilots():
        pilots = {}
        for line in PC_PILOT_RAW.strip().splitlines():
            line = line.strip()
            if not line:
                continue
            try:
                parts = next(csv.reader([line]))
            except Exception:
                parts = line.split(",")
            if len(parts) >= 2:
                pilots[parts[0].strip()] = parts[1].strip().strip('"')
        return pilots

    def pc_crew_all_chinese(crew_codes, pilots):
        pilot_codes = [c.strip() for c in crew_codes if c.strip().startswith(("P", "W"))]
        if not pilot_codes:
            return None
        for code in pilot_codes:
            if code not in pilots:
                return None
            if not pc_has_chinese(pilots[code]):
                return False
        return True

    def pc_actual_nation_label(crew_codes, pilots):
        all_cn = pc_crew_all_chinese(crew_codes, pilots)
        if all_cn is None:
            return "机组未定"
        return "中国籍" if all_cn else "外籍"

    def pc_is_ferry_use(use_text):
        return any(k in use_text for k in PC_FERRY_KEYWORDS)

    # ---------- 解析批复 ----------
    PC_APPROVAL_RE = re.compile(
        r"^(?P<reg>[A-Z0-9\-]+)\s+"
        r"(?P<second>[A-Z0-9]+)\s+"
        r"(?P<dep>[A-Z]{4})\s*(?P<dep_time>\d{4})\s+"
        r"(?P<arr_time>\d{4})\s*(?P<arr>[A-Z]{4})\s+"
        r"ON\s+(?P<date>\d{2}[A-Z]{3}\d{2,4})\s*"
        r"(?P<rest>.*)$",
        re.IGNORECASE,
    )

    def pc_parse_approval_line(text):
        m = PC_APPROVAL_RE.match(text.strip())
        if not m:
            return None

        reg = m.group("reg").upper().replace("-", "")
        second = m.group("second").upper()
        dep = m.group("dep").upper()
        arr = m.group("arr").upper()
        dep_raw = m.group("dep_time")
        arr_raw = m.group("arr_time")
        date_raw = m.group("date").upper()
        rest = m.group("rest").strip()

        if pc_is_b_reg(reg):
            ac_type = second
            flight_no = ""
        else:
            if second == reg:
                flight_no = second
                ac_type = ""
            else:
                ac_type = second
                flight_no = ""

        service, remark = "", ""
        sm = re.match(r"^(U/H|N/M)\s*(.*)$", rest, re.IGNORECASE)
        if sm:
            service = sm.group(1).upper()
            remark = sm.group(2).strip()
        else:
            upper = rest.upper()
            if upper.startswith("FERRY"):
                service = "N/M"
                remark = rest[5:].strip(" -–—\t")
            elif upper.startswith("BUSINESS"):
                service = "U/H"
                remark = rest[8:].strip(" -–—\t")
            else:
                service = ""
                remark = rest.strip()

        date_obj = pc_parse_date_token(date_raw)
        return {
            "raw": text.strip(),
            "reg": reg,
            "type": ac_type,
            "flight_no": flight_no,
            "is_domestic": pc_is_b_reg(reg),
            "dep": dep,
            "dep_time_raw": dep_raw,
            "arr_time_raw": arr_raw,
            "arr": arr,
            "date_raw": date_raw,
            "dep_dt_bj": pc_to_beijing_datetime(reg, date_obj, dep_raw),
            "arr_dt_bj": pc_to_beijing_datetime(reg, date_obj, arr_raw),
            "service": service,
            "remark": remark,
        }

    # ---------- 待申请行签名 ----------
    def pc_parse_pending_signature(text):
        text = str(text).strip()
        if not text:
            return None

        m = re.match(
            r'^[A-Z0-9\-]+\s+'
            r'(?P<dep>[A-Z]{4})-(?P<arr>[A-Z]{4})\s+'
            r'(?P<day>\d{1,2})(?P<mon>[A-Za-z]{3})',
            text
        )
        if m:
            try:
                mon = PC_MONTHS[m.group("mon").upper()[:3]]
                return (m.group("dep").upper(), m.group("arr").upper(),
                        mon, int(m.group("day")))
            except Exception:
                pass

        m = re.match(
            r'^(?P<month>\d{1,2})[.\-/](?P<day>\d{1,2})\s+'
            r'(?P<dep>[A-Z]{4})-(?P<arr>[A-Z]{4})',
            text
        )
        if m:
            try:
                return (m.group("dep").upper(), m.group("arr").upper(),
                        int(m.group("month")), int(m.group("day")))
            except Exception:
                pass

        return None

    def pc_parse_pending_reg(text):
        text = str(text).strip()
        if not text:
            return ""
        m = re.match(r'^([A-Z0-9\-]+)\s+', text)
        if not m:
            return ""
        return m.group(1).upper().replace("-", "")

    def pc_iter_doc_paragraphs(doc):
        for p in doc.paragraphs:
            yield p
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        yield p
                    for nested in cell.tables:
                        for nrow in nested.rows:
                            for ncell in nrow.cells:
                                for np in ncell.paragraphs:
                                    yield np

    def pc_collect_all_paragraphs(doc):
        result = []
        for p in doc.paragraphs:
            result.append(p)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        result.append(p)
                    for nested in cell.tables:
                        for nrow in nested.rows:
                            for ncell in nrow.cells:
                                for np in ncell.paragraphs:
                                    result.append(np)
        return result

    def pc_split_paragraph_by_br(p_elem):
        parent = p_elem.getparent()
        if parent is None:
            return
        idx_in_parent = list(parent).index(p_elem)
        pPr = p_elem.find(qn('w:pPr'))

        groups = [[]]
        for child in list(p_elem):
            if child.tag == qn('w:pPr'):
                continue
            if child.tag == qn('w:r'):
                brs = child.findall(qn('w:br'))
                if not brs:
                    groups[-1].append(copy.deepcopy(child))
                else:
                    cur = []
                    for rc in list(child):
                        if rc.tag == qn('w:br'):
                            if cur:
                                nr = OxmlElement('w:r')
                                for x in cur:
                                    nr.append(copy.deepcopy(x))
                                groups[-1].append(nr)
                                cur = []
                            groups.append([])
                        else:
                            cur.append(rc)
                    if cur:
                        nr = OxmlElement('w:r')
                        for x in cur:
                            nr.append(copy.deepcopy(x))
                        groups[-1].append(nr)
            else:
                groups[-1].append(copy.deepcopy(child))

        if len(groups) <= 1:
            return

        parent.remove(p_elem)
        for i, group in enumerate(groups):
            new_p = OxmlElement('w:p')
            if pPr is not None:
                new_p.append(copy.deepcopy(pPr))
            for child in group:
                new_p.append(child)
            parent.insert(idx_in_parent + i, new_p)

    # ★ 合并软换行拆散的批复行，并保证 ON 和日期之间有空格
    def pc_merge_split_approvals(doc):
        def _process_parent(parent_elem):
            changed = True
            while changed:
                changed = False
                ps = parent_elem.findall(qn('w:p'))
                i = 0
                while i < len(ps) - 1:
                    p1 = ps[i]
                    p2 = ps[i + 1]
                    t1 = "".join(t.text or "" for t in p1.iter(qn('w:t'))).strip()
                    t2 = "".join(t.text or "" for t in p2.iter(qn('w:t'))).strip()
                    if re.search(r'\bON\s*$', t1, re.IGNORECASE) and \
                       re.match(r'^\d{2}[A-Za-z]{3}\d{2,4}', t2):
                        # ★ 给 p1 最后一个 w:t 结尾加空格（若没有）
                        if not t1.endswith(' '):
                            last_t = None
                            for t_elem in p1.iter(qn('w:t')):
                                last_t = t_elem
                            if last_t is not None:
                                txt = last_t.text or ''
                                if not txt.endswith(' '):
                                    last_t.text = txt + ' '
                                    last_t.set(qn('xml:space'), 'preserve')

                        for child in list(p2):
                            if child.tag == qn('w:pPr'):
                                continue
                            p1.append(copy.deepcopy(child))
                        parent_elem.remove(p2)
                        changed = True
                        break
                    i += 1

        body = doc.element.body
        _process_parent(body)
        for tc in body.iter(qn('w:tc')):
            _process_parent(tc)

    def pc_normalize_soft_breaks(doc):
        for p in pc_collect_all_paragraphs(doc):
            pc_split_paragraph_by_br(p._element)
        pc_merge_split_approvals(doc)

    # ---------- Excel ----------
    def pc_load_excel_rows_from_bytes(data: bytes):
        wb = load_workbook(io.BytesIO(data), data_only=True)
        ws = wb["航段(北京时)"] if "航段(北京时)" in wb.sheetnames else wb.active

        rows = []
        for r in range(3, ws.max_row + 1):
            reg = ws.cell(r, 3).value
            if not reg:
                continue
            rows.append({
                "_idx": len(rows),
                "reg": str(reg).strip().upper(),
                "use": str(ws.cell(r, 4).value or "").strip(),
                "dep_date": pc_fmt_date(ws.cell(r, 7).value),
                "dep_time": pc_fmt_time(ws.cell(r, 8).value),
                "dep": str(ws.cell(r, 11).value or "").strip().upper(),
                "dep_city": str(ws.cell(r, 12).value or "").strip(),
                "arr": str(ws.cell(r, 13).value or "").strip().upper(),
                "arr_city": str(ws.cell(r, 14).value or "").strip(),
                "arr_date": pc_fmt_date(ws.cell(r, 15).value),
                "arr_time": pc_fmt_time(ws.cell(r, 16).value),
            })
        return rows

    # ---------- 文本航班 ----------
    PC_FLIGHT_HEADER_RE = re.compile(
        r"^([A-Z0-9]+)\s+(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})(?:\s*\+1)?$"
    )

    def pc_load_text_flights(text: str):
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        flights = []
        i = 0
        pending_f = False

        while i < len(lines):
            line = lines[i]

            if line.upper() == "F":
                pending_f = True
                i += 1
                continue

            if line.upper() in ("TBA",):
                i += 1
                continue

            m = PC_FLIGHT_HEADER_RE.match(line)
            if m:
                reg = m.group(1).upper().replace("-", "")
                dep_time, arr_time = m.group(2), m.group(3)
                if i + 1 < len(lines):
                    cm = re.match(r"^(.+?)\s+-\s+(.+)$", lines[i + 1])
                    if cm:
                        crew = []
                        step = 2
                        if i + 2 < len(lines):
                            cl = lines[i + 2].replace(" ", "")
                            if re.match(r"^[A-Z0-9,]+$", cl):
                                crew = [x for x in cl.split(",") if x]
                                step = 3

                        flights.append({
                            "_idx": len(flights),
                            "reg": reg,
                            "dep_time": dep_time,
                            "arr_time": arr_time,
                            "dep_city": cm.group(1).strip(),
                            "arr_city": cm.group(2).strip(),
                            "crew": crew,
                            "is_ferry": pending_f,
                        })
                        pending_f = False
                        i += step
                        continue
            i += 1
        return flights

    # ---------- 覆盖率 ----------
    def pc_check_text_coverage(excel_rows, text_flights, city_to_icao):
        domestic_rows = [
            r for r in excel_rows
            if (r["dep"].startswith("Z") or r["arr"].startswith("Z"))
            and r["dep"]
            and r["arr"]
        ]

        text_by_reg = {}
        for tf in text_flights:
            text_by_reg.setdefault(tf["reg"], []).append(tf)

        covered = 0
        missing = []
        for r in domestic_rows:
            candidates = text_by_reg.get(r["reg"], [])
            found = False
            for tf in candidates:
                dep_icao = city_to_icao.get(tf["dep_city"])
                arr_icao = city_to_icao.get(tf["arr_city"])
                if dep_icao == r["dep"] and arr_icao == r["arr"]:
                    found = True
                    break
            if found:
                covered += 1
            else:
                missing.append(r)

        return len(domestic_rows), covered, missing

    # ---------- 匹配 ----------
    def pc_find_excel_match(approval, excel_rows, used_excel):
        base_candidates = [
            r for r in excel_rows
            if r["_idx"] not in used_excel
            and r["reg"] == approval["reg"]
            and r["dep"] == approval["dep"]
            and r["arr"] == approval["arr"]
        ]
        if not base_candidates:
            return None

        approval_date = approval["dep_dt_bj"].date()
        approval_dep_time = approval["dep_dt_bj"].strftime("%H:%M")

        same_date = [r for r in base_candidates if r["dep_date"] == approval_date]
        if same_date:
            same_date.sort(key=lambda r: pc_time_diff_minutes(r["dep_time"], approval_dep_time))
            matched = same_date[0]
            used_excel.add(matched["_idx"])
            return matched

        close_date = []
        for r in base_candidates:
            if r["dep_date"] is None:
                continue
            delta_days = (approval_date - r["dep_date"]).days
            if abs(delta_days) != 1:
                continue
            real_diff = pc_compute_real_minute_diff(
                approval_date, approval_dep_time,
                r["dep_date"], r["dep_time"]
            )
            if real_diff is not None and abs(real_diff) <= PC_MAX_CROSS_DAY_GAP_MIN:
                close_date.append((abs(real_diff), r))

        if close_date:
            close_date.sort(key=lambda x: x[0])
            matched = close_date[0][1]
            used_excel.add(matched["_idx"])
            return matched

        return None

    def pc_find_text_match(approval, text_flights, city_to_icao, used_text):
        candidates = []
        for i, tf in enumerate(text_flights):
            if i in used_text:
                continue
            if tf["reg"] != approval["reg"]:
                continue
            if (city_to_icao.get(tf["dep_city"]) == approval["dep"]
                    and city_to_icao.get(tf["arr_city"]) == approval["arr"]):
                candidates.append((i, tf))

        if not candidates:
            return None

        target = approval["dep_dt_bj"].strftime("%H:%M")
        candidates.sort(key=lambda x: pc_time_diff_minutes(x[1]["dep_time"], target))
        idx, matched = candidates[0]
        used_text.add(idx)
        return matched

    # ---------- docx 样式 ----------
    PC_W_R = qn('w:r')
    PC_W_RPR = qn('w:rPr')
    PC_W_COLOR = qn('w:color')
    PC_W_T = qn('w:t')
    PC_W_HIGHLIGHT = qn('w:highlight')
    PC_W_STRIKE = qn('w:strike')
    PC_W_DSTRIKE = qn('w:dstrike')

    def pc_set_run_color(run_element, color_hex):
        rPr = run_element.find(PC_W_RPR)
        if rPr is None:
            rPr = run_element.makeelement(PC_W_RPR, {})
            run_element.insert(0, rPr)
        for c in rPr.findall(PC_W_COLOR):
            rPr.remove(c)
        color = rPr.makeelement(PC_W_COLOR, {qn('w:val'): color_hex})
        rPr.append(color)

    def pc_set_run_highlight(run_element, color_name=PC_HIGHLIGHT_ADDED):
        rPr = run_element.find(PC_W_RPR)
        if rPr is None:
            rPr = run_element.makeelement(PC_W_RPR, {})
            run_element.insert(0, rPr)
        for h in rPr.findall(PC_W_HIGHLIGHT):
            rPr.remove(h)
        hl = rPr.makeelement(PC_W_HIGHLIGHT, {qn('w:val'): color_name})
        rPr.append(hl)

    def pc_set_run_strike(run_element):
        rPr = run_element.find(PC_W_RPR)
        if rPr is None:
            rPr = run_element.makeelement(PC_W_RPR, {})
            run_element.insert(0, rPr)
        for tag in (PC_W_STRIKE, PC_W_DSTRIKE):
            for e in rPr.findall(tag):
                rPr.remove(e)
        strike = rPr.makeelement(PC_W_STRIKE, {})
        rPr.append(strike)

    def pc_clear_run_strike(run_element):
        rPr = run_element.find(PC_W_RPR)
        if rPr is None:
            return
        for tag in (PC_W_STRIKE, PC_W_DSTRIKE):
            for e in rPr.findall(tag):
                rPr.remove(e)

    def pc_make_run_like(src_run_elem, text, color_hex=None, highlight=None):
        new_r = copy.deepcopy(src_run_elem)
        for t in new_r.findall(PC_W_T):
            new_r.remove(t)
        t = new_r.makeelement(PC_W_T, {})
        t.text = text
        t.set(qn('xml:space'), 'preserve')
        new_r.append(t)
        if color_hex:
            pc_set_run_color(new_r, color_hex)
        if highlight:
            pc_set_run_highlight(new_r, highlight)
        return new_r

    def pc_set_paragraph_runs(paragraph, text, color_overrides, highlight_overrides=None):
        if not color_overrides and not highlight_overrides:
            return
        runs = list(paragraph.runs)
        if not runs:
            return
        full_text = "".join(r.text for r in runs)
        if not full_text:
            return

        char_color = [None] * len(full_text)
        char_high = [None] * len(full_text)

        for color_hex, parts in (color_overrides or []):
            for part in parts:
                if not part:
                    continue
                start = 0
                while True:
                    idx = full_text.find(part, start)
                    if idx == -1:
                        break
                    for i in range(idx, idx + len(part)):
                        char_color[i] = color_hex
                    start = idx + len(part)

        for hl_color, parts in (highlight_overrides or []):
            for part in parts:
                if not part:
                    continue
                start = 0
                while True:
                    idx = full_text.find(part, start)
                    if idx == -1:
                        break
                    for i in range(idx, idx + len(part)):
                        char_high[i] = hl_color
                    start = idx + len(part)

        if not any(c is not None for c in char_color) and not any(h is not None for h in char_high):
            return

        pos = 0
        for run in runs:
            r_text = run.text
            if not r_text:
                continue
            r_start = pos
            r_len = len(r_text)
            run_colors = [char_color[r_start + i] for i in range(r_len)]
            run_highs = [char_high[r_start + i] for i in range(r_len)]

            unique = set(zip(run_colors, run_highs))
            if len(unique) == 1:
                c, h = run_colors[0], run_highs[0]
                if c is None and h is None:
                    pos += r_len
                    continue
                if c is not None:
                    pc_set_run_color(run._element, c)
                if h is not None:
                    pc_set_run_highlight(run._element, h)
                pos += r_len
                continue

            run_elem = run._element
            parent = run_elem.getparent()
            idx_in_parent = list(parent).index(run_elem)

            pieces = []
            i = 0
            while i < r_len:
                c = run_colors[i]
                h = run_highs[i]
                j = i + 1
                while j < r_len and run_colors[j] == c and run_highs[j] == h:
                    j += 1
                pieces.append((r_text[i:j], c, h))
                i = j

            parent.remove(run_elem)
            for k, (seg, c, h) in enumerate(pieces):
                new_r = pc_make_run_like(run_elem, seg, color_hex=c, highlight=h)
                parent.insert(idx_in_parent + k, new_r)

            pos += r_len

    def pc_highlight_text(paragraph, target_text, color_hex=PC_RED, highlight=PC_HIGHLIGHT_ADDED):
        if not target_text:
            return
        runs = list(paragraph.runs)
        if not runs:
            return
        full_text = "".join(r.text for r in runs)
        if not full_text:
            return

        positions = []
        start = 0
        while True:
            idx = full_text.find(target_text, start)
            if idx == -1:
                break
            positions.append((idx, idx + len(target_text)))
            start = idx + len(target_text)

        if not positions:
            return

        char_color = [None] * len(full_text)
        char_high = [None] * len(full_text)
        for s, e in positions:
            for i in range(s, e):
                char_color[i] = color_hex
                char_high[i] = highlight

        pos = 0
        for run in runs:
            r_text = run.text
            if not r_text:
                continue
            r_start = pos
            r_len = len(r_text)
            colors = [char_color[r_start + i] for i in range(r_len)]
            highs = [char_high[r_start + i] for i in range(r_len)]

            unique = set(zip(colors, highs))
            if len(unique) == 1:
                c, h = colors[0], highs[0]
                if c is not None:
                    pc_set_run_color(run._element, c)
                if h is not None:
                    pc_set_run_highlight(run._element, h)
                pos += r_len
                continue

            run_elem = run._element
            parent = run_elem.getparent()
            idx_in_parent = list(parent).index(run_elem)

            pieces = []
            i = 0
            while i < r_len:
                c = colors[i]
                h = highs[i]
                j = i + 1
                while j < r_len and colors[j] == c and highs[j] == h:
                    j += 1
                pieces.append((r_text[i:j], c, h))
                i = j

            parent.remove(run_elem)
            for k, (seg, c, h) in enumerate(pieces):
                new_r = pc_make_run_like(run_elem, seg, color_hex=c, highlight=h)
                parent.insert(idx_in_parent + k, new_r)

            pos += r_len

    # ★ 新增：在段落里 "ON <date>" 之后插入一段带样式的文字
    def pc_insert_after_on_date(paragraph, date_raw, insert_text,
                                color_hex=PC_RED, highlight=PC_HIGHLIGHT_ADDED):
        """在段落的 'ON <date>' 之后插入带样式的文字。返回 True 表示成功。"""
        runs = list(paragraph.runs)
        if not runs:
            return False

        full_text = "".join(r.text for r in runs)
        if not full_text:
            return False

        # 找 "ON <date>"
        m = re.search(r'ON\s+' + re.escape(date_raw), full_text, re.IGNORECASE)
        if not m:
            return False
        insert_pos = m.end()

        # 判断插入点后是否紧跟空格
        next_char = full_text[insert_pos] if insert_pos < len(full_text) else ""
        if next_char == " ":
            actual_insert = f"{insert_text} "
        else:
            actual_insert = f" {insert_text} "

        # 找 insert_pos 属于哪个 run
        pos = 0
        for run in runs:
            r_text = run.text
            r_len = len(r_text)
            r_start = pos
            r_end = pos + r_len

            if insert_pos <= r_end and insert_pos > r_start:
                local = insert_pos - r_start
                before = r_text[:local]
                after = r_text[local:]

                run.text = before

                parent = run._element.getparent()
                idx_in_parent = list(parent).index(run._element)

                # 插入带样式的新 run
                new_r = pc_make_run_like(run._element, actual_insert,
                                          color_hex=color_hex, highlight=highlight)
                parent.insert(idx_in_parent + 1, new_r)

                # 保留后面的文本
                if after:
                    after_r = pc_make_run_like(run._element, after)
                    pc_clear_run_strike(after_r)
                    parent.insert(idx_in_parent + 2, after_r)

                return True

            pos += r_len

        return False

    def _pc_append_styled_text(paragraph, text, color_hex):
        runs = list(paragraph.runs)
        p_elem = paragraph._element
        if runs:
            src = runs[-1]._element
            new_r = pc_make_run_like(src, text, color_hex=color_hex, highlight=PC_HIGHLIGHT_ADDED)
            pc_clear_run_strike(new_r)
        else:
            new_r = p_elem.makeelement(PC_W_R, {})
            rPr = new_r.makeelement(PC_W_RPR, {})
            new_r.insert(0, rPr)
            color = rPr.makeelement(PC_W_COLOR, {qn('w:val'): color_hex})
            rPr.append(color)
            hl = rPr.makeelement(PC_W_HIGHLIGHT, {qn('w:val'): PC_HIGHLIGHT_ADDED})
            rPr.append(hl)
            t = new_r.makeelement(PC_W_T, {})
            t.text = text
            t.set(qn('xml:space'), 'preserve')
            new_r.append(t)
        p_elem.append(new_r)

    def pc_append_red_text(paragraph, text):
        _pc_append_styled_text(paragraph, text, PC_RED)

    def pc_append_warn_text(paragraph, text):
        _pc_append_styled_text(paragraph, text, PC_RED)

    def pc_find_target_cell(doc, reg):
        reg_norm = reg.upper().replace("-", "")
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        ap = pc_parse_approval_line(p.text.strip())
                        if ap and ap["reg"] == reg_norm:
                            return cell
        for table in doc.tables:
            for row_idx, row in enumerate(table.rows):
                for cell in row.cells:
                    text = cell.text.strip().strip("*").strip().replace("-", "").upper()
                    if text == reg_norm:
                        if row_idx + 1 < len(table.rows):
                            next_row = table.rows[row_idx + 1]
                            return next_row.cells[0]
                        return cell
        return None

    def pc_find_global_template_paragraph(doc):
        for p in pc_iter_doc_paragraphs(doc):
            if pc_parse_approval_line(p.text.strip()):
                return p
        return None

    def pc_make_red_paragraph_element(text, template_p):
        p_elem = OxmlElement('w:p')
        if template_p is not None:
            template_pPr = template_p._element.find(qn('w:pPr'))
            if template_pPr is not None:
                p_elem.append(copy.deepcopy(template_pPr))

        r_elem = OxmlElement('w:r')
        rPr = None
        if template_p is not None:
            for r in template_p.runs:
                rPr_src = r._element.find(PC_W_RPR)
                if rPr_src is not None:
                    rPr = copy.deepcopy(rPr_src)
                    break
        if rPr is None:
            rPr = OxmlElement('w:rPr')

        for c in rPr.findall(PC_W_COLOR):
            rPr.remove(c)
        for h in rPr.findall(PC_W_HIGHLIGHT):
            rPr.remove(h)
        for tag in (PC_W_STRIKE, PC_W_DSTRIKE):
            for e in rPr.findall(tag):
                rPr.remove(e)

        color = OxmlElement('w:color')
        color.set(qn('w:val'), PC_RED)
        rPr.append(color)

        hl = OxmlElement('w:highlight')
        hl.set(qn('w:val'), PC_HIGHLIGHT_ADDED)
        rPr.append(hl)

        r_elem.append(rPr)

        t_elem = OxmlElement('w:t')
        t_elem.text = text
        t_elem.set(qn('xml:space'), 'preserve')
        r_elem.append(t_elem)

        p_elem.append(r_elem)
        return p_elem

    def pc_reorder_cell_with_pending(cell, pending_items, global_template_p):
        existing_sigs = set()
        existing_paras = []
        for p in cell.paragraphs:
            raw = p.text.strip()
            if not raw:
                continue

            dt = None
            ap = pc_parse_approval_line(raw)
            if ap:
                dt = ap["dep_dt_bj"].date()

            sig = pc_parse_pending_signature(raw)
            if sig:
                dep, arr, month, day = sig
                try:
                    dt = _pcdt.date(2026, month, day)
                except Exception:
                    dt = None
                existing_sigs.add(sig)

            existing_paras.append((p, dt))

        if not existing_paras:
            return

        to_insert = []
        for item in pending_items:
            row = item["row"]
            if row["dep_date"] is None:
                continue
            sig = (row["dep"], row["arr"],
                   row["dep_date"].month, row["dep_date"].day)
            if sig in existing_sigs:
                continue
            existing_sigs.add(sig)
            to_insert.append({
                "date": row["dep_date"],
                "dt": item["dt"],
                "text": item["text"],
            })

        if not to_insert:
            return

        template_p = None
        for p, _ in existing_paras:
            if pc_parse_approval_line(p.text.strip()):
                template_p = p
                break
        if template_p is None:
            template_p = global_template_p
        if template_p is None:
            return

        to_insert.sort(key=lambda x: x["dt"])

        pending_by_date = {}
        for item in to_insert:
            pending_by_date.setdefault(item["date"], []).append(item)

        result_seq = []
        if existing_paras:
            known_dates = [d for _, d in existing_paras if d is not None]
            if known_dates:
                min_date = min(known_dates)
                for d in sorted(k for k in pending_by_date if k < min_date):
                    for item in pending_by_date[d]:
                        result_seq.append(("pending", item))
                    del pending_by_date[d]

        for i, (p, date) in enumerate(existing_paras):
            result_seq.append(("original", p._element))

            is_last_of_date = True
            if date is not None:
                for j in range(i + 1, len(existing_paras)):
                    if existing_paras[j][1] == date:
                        is_last_of_date = False
                        break

            if is_last_of_date and date is not None and date in pending_by_date:
                for item in pending_by_date[date]:
                    result_seq.append(("pending", item))
                del pending_by_date[date]

        for d in sorted(pending_by_date.keys()):
            for item in pending_by_date[d]:
                result_seq.append(("pending", item))

        tc = cell._tc
        for p_elem in list(tc.findall(qn('w:p'))):
            tc.remove(p_elem)

        for kind, item in result_seq:
            if kind == "original":
                tc.append(item)
            else:
                new_p = pc_make_red_paragraph_element(item["text"], template_p)
                tc.append(new_p)

    def pc_build_approval_text(excel_row, is_domestic, note_kind=""):
        reg = excel_row["reg"]
        dep_date = excel_row["dep_date"]
        if not dep_date:
            return None

        date_str = dep_date.strftime("%d%b").upper()
        parts = [reg, f"{excel_row['dep']}-{excel_row['arr']}", date_str]

        if is_domestic:
            if note_kind == "中国籍":
                parts.append("中国籍")
            elif note_kind == "外籍":
                parts.append("外籍")
            elif note_kind == "机组未定":
                parts.append("机组未定")

        parts.append("待申请")
        return " ".join(parts)

    # ---------- 主核对 ----------
    def pc_run_check(docx_bytes, excel_bytes, text_content, pilots):
        excel_rows = pc_load_excel_rows_from_bytes(excel_bytes)
        text_flights = pc_load_text_flights(text_content)

        city_to_icao = {}
        for row in excel_rows:
            if row["dep_city"]:
                city_to_icao[row["dep_city"]] = row["dep"]
            if row["arr_city"]:
                city_to_icao[row["arr_city"]] = row["arr"]

        doc = Document(io.BytesIO(docx_bytes))
        pc_normalize_soft_breaks(doc)

        result_rows = []
        approval_red_map = {}
        approval_green_map = {}
        approval_highlight_map = {}
        cancel_paragraphs = set()
        change_paragraphs = set()
        nationality_pending_paragraphs = set()
        # ★ 新增：记录"需要补 U/H 或 N/M"的段落 [(paragraph, date_raw, service)]
        service_to_insert = []

        used_excel = set()
        used_text = set()
        unapproved_keys = set()

        for p in pc_iter_doc_paragraphs(doc):
            raw_text = p.text.strip()
            if not raw_text:
                continue

            approval = pc_parse_approval_line(raw_text)
            if not approval:
                continue

            if not approval["is_domestic"] and approval["flight_no"]:
                unapproved_keys.add(
                    (approval["reg"], approval["dep"], approval["arr"])
                )
                result_rows.append({
                    "批复": raw_text,
                    "飞机号": approval["reg"],
                    "航班号": approval["flight_no"],
                    "机型": PC_AIRCRAFT_TYPE_MAP.get(approval["reg"], ""),
                    "内/外机": "外机",
                    "批复起飞(北京时)": approval["dep_dt_bj"].strftime("%H:%M"),
                    "批复落地(北京时)": approval["arr_dt_bj"].strftime("%H:%M"),
                    "批复日期": approval["dep_dt_bj"].date().isoformat(),
                    "Excel 用途": "",
                    "文本标记": "",
                    "Excel 计划起飞": "",
                    "Excel 计划落地": "",
                    "Excel 出发地": "",
                    "Excel 到达地": "",
                    "差异": "外机未批（已申请）",
                    "是否一致": "待确认",
                    "备注": "未批",
                })
                continue

            red_parts = []
            green_parts = []
            highlight_parts = []
            diffs = []
            note_parts = []
            info_note = ""

            excel_row = pc_find_excel_match(approval, excel_rows, used_excel)

            # ★ 新增：service 为空 → 记录需要补的内容
            if approval["service"] == "" and excel_row is not None:
                _ferry = pc_is_ferry_use(excel_row["use"])
                _expected_svc = "N/M" if _ferry else "U/H"
                service_to_insert.append((p, approval["date_raw"], _expected_svc))

            if excel_row is None:
                note_parts.append("待取消")
                red_parts.append(approval["dep"])
                red_parts.append(approval["arr"])
                red_parts.append(approval["date_raw"])

                result_rows.append({
                    "批复": raw_text,
                    "飞机号": approval["reg"],
                    "航班号": approval["flight_no"] if not approval["is_domestic"] else "",
                    "机型": approval["type"] if approval["is_domestic"] else "",
                    "内/外机": "内机" if approval["is_domestic"] else "外机",
                    "批复起飞(北京时)": approval["dep_dt_bj"].strftime("%H:%M"),
                    "批复落地(北京时)": approval["arr_dt_bj"].strftime("%H:%M"),
                    "批复日期": approval["dep_dt_bj"].date().isoformat(),
                    "Excel 用途": "",
                    "文本标记": "",
                    "Excel 计划起飞": "",
                    "Excel 计划落地": "",
                    "Excel 出发地": "",
                    "Excel 到达地": "",
                    "差异": f"Excel 中无 {approval['dep']}→{approval['arr']}（±1 天）匹配",
                    "是否一致": "否",
                    "备注": "待取消",
                })
                approval_red_map[raw_text] = red_parts
                cancel_paragraphs.add(raw_text)
                continue

            text_flight = pc_find_text_match(
                approval, text_flights, city_to_icao, used_text
            )

            approval_dep_time = approval["dep_dt_bj"].strftime("%H:%M")
            approval_arr_time = approval["arr_dt_bj"].strftime("%H:%M")

            if approval["is_domestic"]:
                expected_type = PC_AIRCRAFT_TYPE_MAP.get(approval["reg"])
                if expected_type is None:
                    diffs.append(f"机型：注册号 {approval['reg']} 不在机型对照表中")
                    red_parts.append(approval["type"])
                    highlight_parts.append(approval["type"])
                elif approval["type"] != expected_type:
                    diffs.append(f"机型：批复 {approval['type']} vs 对照表 {expected_type}")
                    red_parts.append(approval["type"])
                    highlight_parts.append(approval["type"])

            ap_dur = pc_flight_duration_minutes(approval_dep_time, approval_arr_time)
            xl_dur = pc_flight_duration_minutes(excel_row["dep_time"], excel_row["arr_time"])
            dur_diff = None
            if ap_dur is not None and xl_dur is not None:
                dur_diff = ap_dur - xl_dur
                if abs(dur_diff) > 30:
                    diffs.append(
                        f"飞行时长：批复 {pc_fmt_duration(ap_dur)} vs 计划 {pc_fmt_duration(xl_dur)}"
                        f"（差 {dur_diff:+d} 分钟）"
                    )
                    red_parts.append(approval["arr_time_raw"])
                    highlight_parts.append(approval["arr_time_raw"])

            real_dep_diff = pc_compute_real_minute_diff(
                approval["dep_dt_bj"].date(), approval_dep_time,
                excel_row["dep_date"], excel_row["dep_time"]
            )
            dep_ok = False
            if real_dep_diff is None:
                dep_ok = False
            elif real_dep_diff == 0:
                dep_ok = True
            elif -PC_EARLY_GREEN_THRESHOLD_MIN <= real_dep_diff < 0:
                dep_ok = True
                green_parts.append(approval["dep_time_raw"])
            elif real_dep_diff > 0:
                dep_ok = False
                diffs.append(
                    f"起飞时间：批复 {approval_dep_time} 晚于计划 {excel_row['dep_time']}，需重新申请"
                )
                red_parts.append(approval["dep_time_raw"])
                highlight_parts.append(approval["dep_time_raw"])
                change_paragraphs.add(raw_text)
            else:
                dep_ok = False
                diffs.append(
                    f"起飞时间：批复 {approval_dep_time} vs 计划 {excel_row['dep_time']}"
                    f"（相差 {real_dep_diff} 分钟）"
                )
                red_parts.append(approval["dep_time_raw"])
                highlight_parts.append(approval["dep_time_raw"])

            xl_arr_date = excel_row["arr_date"] or excel_row["dep_date"]
            real_arr_diff = pc_compute_real_minute_diff(
                approval["arr_dt_bj"].date(), approval_arr_time,
                xl_arr_date, excel_row["arr_time"]
            )
            if real_arr_diff is not None and real_arr_diff != 0:
                if -PC_EARLY_GREEN_THRESHOLD_MIN <= real_arr_diff < 0 and dep_ok:
                    green_parts.append(approval["arr_time_raw"])

            excel_ferry = pc_is_ferry_use(excel_row["use"])
            expected_service = "N/M" if excel_ferry else "U/H"

            if approval["service"] and approval["service"] != expected_service:
                diffs.append(
                    f"用途：批复 {approval['service']} vs 计划 {excel_row['use']}"
                    f"（应为 {expected_service}）"
                )
                red_parts.append(approval["service"])
                highlight_parts.append(approval["service"])

            if text_flight is not None:
                text_ferry = text_flight.get("is_ferry", False)
                if excel_ferry and not text_ferry:
                    diffs.append(
                        f"⚠ 文本漏 F 标记（Excel 为调机：{excel_row['use']}）"
                    )
                    if approval["service"] and approval["service"] not in red_parts:
                        red_parts.append(approval["service"])
                        highlight_parts.append(approval["service"])
                elif not excel_ferry and text_ferry:
                    diffs.append(
                        f"⚠ 文本多标 F 标记（Excel 为 {excel_row['use']}，非调机）"
                    )
                    if approval["service"] and approval["service"] not in red_parts:
                        red_parts.append(approval["service"])
                        highlight_parts.append(approval["service"])

            if approval["is_domestic"]:
                if text_flight:
                    all_cn = pc_crew_all_chinese(text_flight["crew"], pilots)
                    if all_cn is None:
                        info_note = "机组国籍待确认"
                    else:
                        has_cn = "中国籍" in approval["remark"]
                        has_foreign = "外籍" in approval["remark"]
                        if all_cn:
                            if has_foreign:
                                diffs.append("国籍标注错误（应为中国籍）")
                                red_parts.append("外籍")
                                highlight_parts.append("外籍")
                            elif not has_cn:
                                diffs.append("国籍未标注（应为中国籍）")
                        else:
                            if has_cn:
                                diffs.append("国籍标注错误（应为外籍）")
                                red_parts.append("中国籍")
                                highlight_parts.append("中国籍")
                            elif not has_foreign:
                                diffs.append("国籍未标注（应为外籍）")
                else:
                    info_note = "机组国籍待确认"

            if raw_text in change_paragraphs:
                note_parts.append("待变更")

            final_notes = list(note_parts)
            if info_note:
                final_notes.append(info_note)

            if diffs or note_parts:
                consistency = "否"
            elif info_note:
                consistency = "待确认"
            else:
                consistency = "是"

            text_ferry_label = ""
            if text_flight is not None:
                text_ferry_label = "调机(F)" if text_flight.get("is_ferry", False) else "载客(无F)"

            result_rows.append({
                "批复": raw_text,
                "飞机号": approval["reg"],
                "航班号": approval["flight_no"] if not approval["is_domestic"] else "",
                "机型": approval["type"] if approval["is_domestic"] else "",
                "内/外机": "内机" if approval["is_domestic"] else "外机",
                "批复起飞(北京时)": approval_dep_time,
                "批复落地(北京时)": approval_arr_time,
                "批复日期": approval["dep_dt_bj"].date().isoformat(),
                "Excel 用途": excel_row["use"] if excel_row else "",
                "文本标记": text_ferry_label,
                "Excel 计划起飞": excel_row["dep_time"] if excel_row else "",
                "Excel 计划落地": excel_row["arr_time"] if excel_row else "",
                "Excel 出发地": excel_row["dep"] if excel_row else "",
                "Excel 到达地": excel_row["arr"] if excel_row else "",
                "差异": "；".join(diffs) if diffs else "无",
                "是否一致": consistency,
                "备注": "；".join(final_notes),
            })

            if red_parts:
                approval_red_map[raw_text] = red_parts
            if green_parts:
                approval_green_map[raw_text] = green_parts
            if highlight_parts:
                approval_highlight_map[raw_text] = highlight_parts
            if info_note:
                nationality_pending_paragraphs.add(raw_text)

        # ★ 先补 U/H 或 N/M（在应用其它样式之前）
        for _para, _date_raw, _svc in service_to_insert:
            pc_insert_after_on_date(_para, _date_raw, _svc,
                                     color_hex=PC_RED,
                                     highlight=PC_HIGHLIGHT_ADDED)

        for p in pc_iter_doc_paragraphs(doc):
            raw_text = p.text.strip()
            red = approval_red_map.get(raw_text, [])
            green = approval_green_map.get(raw_text, [])
            highlight = approval_highlight_map.get(raw_text, [])
            if red or green or highlight:
                pc_set_paragraph_runs(
                    p, raw_text,
                    color_overrides=[(PC_GREEN, green), (PC_RED, red)],
                    highlight_overrides=[(PC_HIGHLIGHT_ADDED, highlight)] if highlight else None,
                )

        for p in pc_iter_doc_paragraphs(doc):
            raw_text = p.text.strip()
            has_cancel = "待取消" in raw_text
            has_change = "待变更" in raw_text
            has_nation = "机组国籍待确认" in raw_text

            if raw_text in cancel_paragraphs:
                for run in p.runs:
                    pc_set_run_strike(run._element)
                if not (has_cancel or has_change):
                    pc_append_red_text(p, "  待取消")
            elif raw_text in change_paragraphs:
                if not (has_cancel or has_change):
                    pc_append_red_text(p, "  待变更")

            if raw_text in nationality_pending_paragraphs:
                if not has_nation:
                    pc_append_warn_text(p, "  机组国籍待确认")

        # 手写待申请行：只对"Excel 里没有该航段"的情况加删除线（=已取消）
        excel_route_keys = set()
        for r in excel_rows:
            if r["dep_date"] and r["dep"] and r["arr"]:
                excel_route_keys.add((r["dep"], r["arr"],
                                      r["dep_date"].month, r["dep_date"].day))

        tf_by_reg_route = {}
        for tf in text_flights:
            dep_icao = city_to_icao.get(tf["dep_city"])
            arr_icao = city_to_icao.get(tf["arr_city"])
            if dep_icao and arr_icao:
                tf_by_reg_route.setdefault((tf["reg"], dep_icao, arr_icao), []).append(tf)

        for p in pc_iter_doc_paragraphs(doc):
            raw_text = p.text.strip()
            if not raw_text:
                continue
            sig = pc_parse_pending_signature(raw_text)
            if not sig:
                continue
            dep, arr, month, day = sig

            if (dep, arr, month, day) not in excel_route_keys:
                for run in p.runs:
                    pc_set_run_strike(run._element)
                continue

            reg_written = pc_parse_pending_reg(raw_text)
            if not pc_is_b_reg(reg_written):
                continue

            written_nation = ""
            if "中国籍" in raw_text:
                written_nation = "中国籍"
            elif "外籍" in raw_text:
                written_nation = "外籍"
            elif "机组未定" in raw_text:
                written_nation = "机组未定"

            if not written_nation:
                continue

            candidates = tf_by_reg_route.get((reg_written, dep, arr), [])
            if not candidates:
                continue

            matched_tf = candidates[0]
            actual_nation = pc_actual_nation_label(matched_tf["crew"], pilots)

            if written_nation != actual_nation:
                pc_highlight_text(p, written_nation,
                                  color_hex=PC_RED,
                                  highlight=PC_HIGHLIGHT_ADDED)

        pending_by_reg = {}
        for row in excel_rows:
            if not row["reg"]:
                continue
            if not (row["dep"].startswith("Z") or row["arr"].startswith("Z")):
                continue
            if row["_idx"] in used_excel:
                continue
            if row["dep_date"] is None:
                continue
            if (row["reg"], row["dep"], row["arr"]) in unapproved_keys:
                continue
            pending_by_reg.setdefault(row["reg"], []).append(row)

        pending_rows = []
        global_template_p = pc_find_global_template_paragraph(doc)

        for reg, rows in pending_by_reg.items():
            target_cell = pc_find_target_cell(doc, reg)
            if target_cell is None:
                continue

            is_domestic = pc_is_b_reg(reg)
            items = []
            for row in rows:
                dep_t = pc_parse_hhmm_str(row["dep_time"])
                if not row["dep_date"] or not dep_t:
                    continue
                dep_dt_bj = _pcdt.datetime.combine(row["dep_date"], dep_t)

                note_kind = ""
                if is_domestic:
                    fake_ap = {
                        "reg": reg,
                        "dep": row["dep"],
                        "arr": row["arr"],
                        "dep_dt_bj": dep_dt_bj,
                    }
                    tf = pc_find_text_match(fake_ap, text_flights, city_to_icao, used_text)
                    if tf is None or not tf["crew"]:
                        note_kind = "机组未定"
                    else:
                        all_cn = pc_crew_all_chinese(tf["crew"], pilots)
                        if all_cn is True:
                            note_kind = "中国籍"
                        elif all_cn is False:
                            note_kind = "外籍"
                        else:
                            note_kind = "机组未定"

                text = pc_build_approval_text(row, is_domestic, note_kind)
                if not text:
                    continue

                items.append({
                    "dt": dep_dt_bj,
                    "text": text,
                    "row": row,
                    "note_kind": note_kind,
                })

            if not items:
                continue

            pc_reorder_cell_with_pending(target_cell, items, global_template_p)

            for item in items:
                row = item["row"]
                if is_domestic:
                    bj_dep = row["dep_time"]
                    bj_arr = row["arr_time"]
                else:
                    dep_t2 = pc_parse_hhmm_str(row["dep_time"])
                    arr_t2 = pc_parse_hhmm_str(row["arr_time"])
                    if row["dep_date"] and dep_t2:
                        bj_dep = (_pcdt.datetime.combine(row["dep_date"], dep_t2)
                                  + _pcdt.timedelta(hours=8)).strftime("%H:%M")
                    else:
                        bj_dep = ""
                    if row["arr_date"] and arr_t2:
                        bj_arr = (_pcdt.datetime.combine(row["arr_date"], arr_t2)
                                  + _pcdt.timedelta(hours=8)).strftime("%H:%M")
                    else:
                        bj_arr = ""

                if item["note_kind"] == "机组未定":
                    remark = "待申请（机组未定，需确认）"
                    diff_text = "Excel 有计划，批复汇总表缺失；文本未提供该航段，机组需确认"
                else:
                    remark = "待申请"
                    diff_text = "Excel 有计划，批复汇总表缺失"

                pending_rows.append({
                    "批复": item["text"],
                    "飞机号": reg,
                    "航班号": reg if not is_domestic else "",
                    "机型": PC_AIRCRAFT_TYPE_MAP.get(reg, "") if is_domestic else "",
                    "内/外机": "内机" if is_domestic else "外机",
                    "批复起飞(北京时)": bj_dep,
                    "批复落地(北京时)": bj_arr,
                    "批复日期": row["dep_date"].isoformat() if row["dep_date"] else "",
                    "Excel 用途": row["use"],
                    "文本标记": "",
                    "Excel 计划起飞": row["dep_time"],
                    "Excel 计划落地": row["arr_time"],
                    "Excel 出发地": row["dep"],
                    "Excel 到达地": row["arr"],
                    "差异": diff_text,
                    "是否一致": "否",
                    "备注": remark,
                })

        result_rows.extend(pending_rows)

        out_buf = io.BytesIO()
        doc.save(out_buf)
        out_buf.seek(0)
        return result_rows, out_buf

    # ---------- 功能4 UI ----------
    pc_col1, pc_col2 = st.columns([1, 1])
    with pc_col1:
        pc_docx_file = st.file_uploader(
            "① 国内批复信息汇总表 (.docx)",
            type=["docx"],
            key="pc_docx",
        )
    with pc_col2:
        pc_excel_file = st.file_uploader(
            "② 航段数据导出 (.xlsx)",
            type=["xlsx"],
            key="pc_excel",
        )

    pc_text_input = st.text_area(
        "③ 粘贴文本版航班信息",
        height=320,
        key="pc_textarea",
        placeholder=(
            "例如：\n"
            "B65AP 16:30 - 17:45\n"
            "香港 - 泉州晋江\n"
            "P057,P039,C046,M035"
        ),
    )

    pc_text_content = pc_text_input.strip()

    pc_coverage_ok = False
    if pc_docx_file and pc_excel_file and pc_text_content:
        preview_excel_rows = pc_load_excel_rows_from_bytes(pc_excel_file.getvalue())
        preview_text_flights = pc_load_text_flights(pc_text_content)

        preview_city_to_icao = {}
        for row in preview_excel_rows:
            if row["dep_city"]:
                preview_city_to_icao[row["dep_city"]] = row["dep"]
            if row["arr_city"]:
                preview_city_to_icao[row["arr_city"]] = row["arr"]

        total, covered, missing = pc_check_text_coverage(
            preview_excel_rows, preview_text_flights, preview_city_to_icao
        )

        st.subheader("🔍 文本覆盖率检查")

        if total == 0:
            st.warning("Excel 里没有国内航段（Z 开头机场），无需核对。")
            pc_coverage_ok = True
        else:
            coverage = covered / total

            if coverage >= 1.0:
                st.success(
                    f"✅ 文本计划覆盖率：{covered}/{total}（100.0%）—— 可以开始核对"
                )
                pc_coverage_ok = True
            else:
                st.error(
                    f"❌ 文本计划覆盖率不足：{covered}/{total}（{coverage*100:.1f}%），"
                    f"缺 {total - covered} 条。**必须 100% 才能开始核对**，请先补全文本。"
                )
                pc_coverage_ok = False

            if missing:
                with st.expander(
                    f"📋 缺失航段明细（{len(missing)} 条）—— 点开查看 / 复制",
                    expanded=True,
                ):
                    missing_sorted = sorted(
                        missing,
                        key=lambda r: (
                            r["dep_date"] or _pcdt.date.min,
                            r["reg"],
                            r["dep_time"],
                        ),
                    )
                    lines = []
                    for r in missing_sorted:
                        date_str = r["dep_date"].strftime("%m-%d") if r["dep_date"] else "??-??"
                        lines.append(
                            f"{r['reg']}  {date_str}  {r['dep']}-{r['arr']}  "
                            f"{r['dep_time']}-{r['arr_time']}  ({r['dep_city']} → {r['arr_city']})"
                        )
                    st.code("\n".join(lines), language=None)

    pc_can_run = bool(pc_docx_file and pc_excel_file and pc_text_content) and pc_coverage_ok

    if st.button("🚀 开始核对", type="primary", disabled=not pc_can_run, key="pc_run_btn"):
        with st.spinner("正在核对..."):
            try:
                rows, out_buf = pc_run_check(
                    pc_docx_file.getvalue(),
                    pc_excel_file.getvalue(),
                    pc_text_content,
                    pc_load_pilots(),
                )
            except Exception as e:
                st.exception(e)
                st.stop()

        df = pd.DataFrame(rows)
        total_rows = len(df)
        diff_count = (df["是否一致"] == "否").sum() if total_rows else 0
        pending_count = (df["是否一致"] == "待确认").sum() if total_rows else 0

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("批复条数", total_rows)
        c2.metric("一致", total_rows - diff_count - pending_count)
        c3.metric("有差异", diff_count)
        c4.metric("待确认", pending_count)

        st.subheader("📋 核对结果")
        if total_rows == 0:
            st.warning("未在 docx 中识别到任何批复行。")
        else:
            def highlight(row):
                if row["是否一致"] == "待确认":
                    return ["background-color: #ffe082; font-weight: bold"] * len(row)
                if row["是否一致"] == "否":
                    if "待变更" in str(row["备注"]):
                        return ["background-color: #e5f0ff"] * len(row)
                    return ["background-color: #ffe5e5"] * len(row)
                return [""] * len(row)

            st.dataframe(
                df.style.apply(highlight, axis=1),
                use_container_width=True,
                hide_index=True,
            )

            st.subheader("🚨 差异明细")
            diffs_df = df[df["是否一致"] != "是"][["批复", "差异", "备注", "是否一致"]]
            if diffs_df.empty:
                st.success("✅ 所有批复与计划一致，未发现差异。")
            else:
                for _, r in diffs_df.iterrows():
                    if r["是否一致"] == "待确认":
                        note_html = (
                            " <span style='color:#cc0000;font-weight:bold;"
                            "background-color:#ffe082;padding:1px 6px;border-radius:3px'>"
                            f"【{r['备注']}】</span>"
                        )
                    elif "待变更" in str(r["备注"]):
                        note_html = (
                            " <span style='color:#0066cc;font-weight:bold'>"
                            f"【{r['备注']}】</span>"
                        )
                    elif r["备注"]:
                        note_html = (
                            f" <span style='color:red;font-weight:bold'>【{r['备注']}】</span>"
                        )
                    else:
                        note_html = ""
                    st.markdown(f"**`{r['批复']}`**{note_html}", unsafe_allow_html=True)
                    if r["差异"] != "无":
                        for line in r["差异"].split("；"):
                            st.markdown(f"- {line}")

        st.subheader("📥 下载标红后的批复汇总表")
        st.download_button(
            label=f"下载 {pc_docx_file.name}",
            data=out_buf.getvalue(),
            file_name=pc_docx_file.name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="pc_download_btn",
        )
