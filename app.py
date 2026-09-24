import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
import re
from datetime import datetime, timedelta
import traceback
import json
import os

# ---------- 页面设置 ----------
st.set_page_config(page_title="备案表&世界时行程&航路处理", layout="wide")
st.title("🛫 备案表 / 世界时行程 / 航路处理")

# ---------- 创建选项卡 ----------
tab1, tab2, tab3 = st.tabs(["📋 功能1：备案表生成", "🌐 功能2：世界时行程", "✈️ 功能3：航路处理工具"])

# ================================================================
# 功能1：备案表生成
# ================================================================
with tab1:
    st.markdown("上传 GD单 和模板，自动生成备案表（联系方式、执照号码及证件号码已内置）。")

    # ---------- 内置机组信息（唯一真源） ----------
    BUILTIN_CREW_DATA = [
        # (姓名, 联系方式, 执照号码, 证件号码)
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
        ("Keith Robert, SHERREN / Keith Robert SHERREN", "137 3540 9744", "12262", ""),
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
        ("郭春旭 / Guo Chunxu", "138 0136 1720", "110107197305150016", "110107197305150016"),
        ("黄海东", "138 0179 9315", "310105197506021215", "310105197506021215"),
    ]

    CREW_COLUMNS = ["姓名", "联系方式", "执照号码", "证件号码"]
    CREW_FILL_COLUMNS = ["职务", "姓名", "性别", "出生日期", "证件号码", "执照号码", "联系方式"]

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
        "TWN": "中国台湾", "MAC": "澳门"
    }

    AIRCRAFT_TYPE_CORRECTION = {"B3926": "LJ60"}

    def correct_aircraft_type(reg, ac_type):
        if reg in AIRCRAFT_TYPE_CORRECTION:
            corrected = AIRCRAFT_TYPE_CORRECTION[reg]
            if ac_type != corrected:
                st.info(f"✈️ 机型修正：{ac_type} → {corrected}（注册号 {reg}）")
            return corrected
        return ac_type

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
                        data["ac_type"] = correct_aircraft_type(data.get("reg", ""), data["ac_type_raw"])
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

    # ---------- 航段数据解析与匹配 ----------
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

    # ---------- 姓名单元格样式优化 ----------
    from copy import copy as _copy_style

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

    # ---------- 功能1 UI ----------
    st.subheader("📂 上传文件")
    st.info("⚠️ 注意：模板文件必须是 **.xlsx** 格式（非 .xls）。联系方式、执照号码及证件号码已内置，无需额外上传。")

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

            # ---------- 本次机组信息（可编辑） ----------
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

            # ---------- 乘客数量警告 ----------
            MAX_PAX_ROWS = 14
            if len(passenger_list) > MAX_PAX_ROWS:
                extra = len(passenger_list) - MAX_PAX_ROWS
                st.warning(
                    f"⚠️ **本次乘客共 {len(passenger_list)} 人**，而模板乘客区只有 **{MAX_PAX_ROWS} 行**。\n\n"
                    f"以下 **{extra} 位乘客不会自动写入模板**，请在下载后手动插入行并复制下方内容："
                )

                lines = ["姓名\t性别\t出生日期\t国籍\t证件种类\t证件号码"]
                for pax in passenger_list[MAX_PAX_ROWS:]:
                    pax_name = extract_chinese_name(pax["name"])
                    doc_type = pax.get("doc_type", "")
                    doc_type_clean = parse_document_type("", doc_type) if (pd.notna(doc_type) and str(doc_type).strip()) else parse_document_type(pax.get("passport_no", ""), "")
                    lines.append(
                        "\t".join([
                            pax_name,
                            str(pax.get("gender", "") or ""),
                            str(pax.get("dob", "") or ""),
                            get_nation_name(pax.get("nationality", "")),
                            doc_type_clean,
                            str(pax.get("passport_no", "") or ""),
                        ])
                    )
                copy_text = "\n".join(lines)

                st.markdown("**📋 超出部分乘客信息（点击右上角复制按钮，直接粘贴到 Excel）**")
                st.code(copy_text, language="text")

            st.markdown("---")

            # ---------- 航班信息 ----------
            from_code = data.get("from", ""); to_code = data.get("to", "")
            date_str = data.get("date_str", ""); utc_time = data.get("utc_time", "")
            reg = data.get("reg", "")
            date_display = get_beijing_date_display(utc_time, date_str) if date_str else ""

            # 优先从航段数据匹配
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

            # 未匹配到则用 GD 单信息生成默认值
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

            result_bytes = fill_template(template_file, data, crew_for_template, passenger_list, route_display)

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
# 功能2：世界时行程（HTML/JS 沙箱版 + localStorage 持久化）
# ================================================================
with tab2:
    F_HTML = r"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>世界时行程转换</title>
    <script src="https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js"></script>
    <style>
        body { font-family: -apple-system, "Segoe UI", "Microsoft YaHei", sans-serif; margin: 12px; color:#333; font-size:15px; }
        h2 { margin: 6px 0 10px 0; font-size: 20px; }
        h3 { margin: 16px 0 8px 0; font-size: 17px; }
        input[type=file] { padding: 6px; }
        button { padding: 8px 14px; font-size: 14px; border-radius: 6px; border: 1px solid #ddd; background:#fff; cursor: pointer; margin-right: 6px; }
        button.primary { background:#ff4b4b; color:#fff; border-color:#ff4b4b; font-weight: bold; }
        button.primary:hover { background:#e63939; }
        button:hover { background:#f5f5f5; }
        .reg-block { margin-bottom: 18px; }
        .reg-title { font-weight: bold; font-size: 16px; margin-bottom: 6px; }
        .new-flag { color:#d32f2f; font-size: 0.9rem; margin-left:8px; font-weight: normal; }
        .seg-list {
            background:#f7f7f7; border: 1px solid #ddd; border-radius: 6px;
            padding: 8px 12px; font-family: Consolas, "Courier New", monospace;
            font-size: 14px; line-height: 1.8; color:#222;
        }
        .seg-line { padding: 2px 0; }
        .status { color:#555; font-size: 14px; margin-left: 8px; }
        .error { color:#d32f2f; background:#ffebee; padding:8px; border-radius:4px; margin:6px 0; }
        .success { color:#2e7d32; background:#e8f5e9; padding:8px; border-radius:4px; margin:6px 0; }
        .info { color:#1976d2; background:#e3f2fd; padding:8px; border-radius:4px; margin:6px 0; }
        .toolbar { margin: 8px 0 12px 0; }
        details { margin: 10px 0; padding: 8px; border: 1px solid #eee; border-radius: 4px; background:#fafafa; }
        summary { cursor: pointer; font-weight: bold; padding: 4px 0; }
        ol { margin: 6px 0 6px 20px; padding: 0; }
        li { margin: 2px 0; }
        .full-text-box {
            background:#f5f5f5; padding:10px; border-radius:4px;
            font-family: Consolas, "Courier New", monospace; font-size: 14px;
            white-space: pre; overflow-x: auto; border:1px solid #e0e0e0;
            max-height: 400px; overflow-y: auto;
        }
    </style>
</head>
<body>
    <h2>🌐 世界时行程</h2>
    <p>从 Jetops 系统导出的北京时间行程 Excel 转换为世界时（UTC），便于复制粘贴。</p>
    <p>💡 上传结果和历史记录会自动保存在浏览器，关闭后再打开仍保留。</p>

    <input type="file" id="fileInput" accept=".xlsx,.xls">
    <div id="status"></div>

    <div id="result" style="display:none;">
        <div class="toolbar">
            <button class="primary" id="copyAllBtn">📋 复制全部计划</button>
            <span class="status" id="copyAllStatus"></span>
        </div>

        <h3>📋 生成的飞行计划（红色为新增/变更）</h3>
        <div id="plans"></div>

        <details>
            <summary>📦 全部计划合并（点击展开）</summary>
            <div class="full-text-box" id="fullTextBox"></div>
        </details>

        <details>
            <summary>📜 历史记录</summary>
            <div id="historyList"></div>
            <button id="clearHistoryBtn" style="margin-top:8px;">🗑️ 清除所有历史</button>
        </details>
    </div>

    <script>
        const MONTHS = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
        const PRIORITY = ['B652Q', 'B65AP', 'B652S', 'MLLIN', 'N88AY', 'B652R'];
        const HISTORY_KEY = 'worldtime_history_v2';
        const LAST_PLANS_KEY = 'worldtime_last_plans_v2';
        const LAST_FILE_KEY = 'worldtime_last_file_v2';

        function loadHistory() {
            try {
                const raw = localStorage.getItem(HISTORY_KEY);
                if (!raw) return {records: []};
                const obj = JSON.parse(raw);
                if (!obj.records) obj.records = [];
                return obj;
            } catch (e) { return {records: []}; }
        }
        function saveHistory(h) {
            try { localStorage.setItem(HISTORY_KEY, JSON.stringify(h)); } catch (e) {}
        }
        function loadJSON(key) {
            try {
                const raw = localStorage.getItem(key);
                return raw ? JSON.parse(raw) : null;
            } catch (e) { return null; }
        }
        function saveJSON(key, val) {
            try { localStorage.setItem(key, JSON.stringify(val)); } catch (e) {}
        }

        function pad2(n) { return String(n).padStart(2, '0'); }
        function escapeHtml(s) {
            return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        }

        function parseDate(v) {
            if (v == null || v === '') return null;
            if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
            if (typeof v === 'number') {
                const ms = Math.round((v - 25569) * 86400 * 1000);
                const d = new Date(ms);
                return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
            }
            const s = String(v).trim();
            if (!s) return null;
            const m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
            if (m) return new Date(parseInt(m[1]), parseInt(m[2])-1, parseInt(m[3]));
            const d = new Date(s);
            if (!isNaN(d.getTime())) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
            return null;
        }

        function parseTime(v) {
            if (v == null || v === '') return null;
            if (v instanceof Date) return pad2(v.getHours()) + ':' + pad2(v.getMinutes());
            if (typeof v === 'number') {
                let frac = v;
                if (v > 1) frac = v - Math.floor(v);
                const totalMin = Math.round(frac * 24 * 60);
                return pad2(Math.floor(totalMin / 60) % 24) + ':' + pad2(totalMin % 60);
            }
            const s = String(v).trim();
            const m = s.match(/(\d{1,2}):(\d{2})/);
            if (m) return pad2(parseInt(m[1])) + ':' + m[2];
            return null;
        }

        function toUTCLabel(dateVal, timeStr) {
            const d = parseDate(dateVal);
            if (!d || !timeStr) return null;
            const parts = timeStr.split(':');
            const h = parseInt(parts[0]);
            const m = parseInt(parts[1]);
            let total = h * 60 + m - 8 * 60;
            let dayOff = 0;
            while (total < 0) { total += 24 * 60; dayOff--; }
            while (total >= 24 * 60) { total -= 24 * 60; dayOff++; }
            const uh = Math.floor(total / 60);
            const um = total % 60;
            const dd = new Date(d.getFullYear(), d.getMonth(), d.getDate());
            dd.setDate(dd.getDate() + dayOff);
            return {
                day: dd.getDate(),
                month: dd.getMonth() + 1,
                hours: uh,
                minutes: um,
                sortKey: dd.getFullYear() * 100000000
                       + (dd.getMonth() + 1) * 1000000
                       + dd.getDate() * 10000
                       + uh * 100 + um
            };
        }

        function formatLabel(u) {
            if (!u) return '';
            return pad2(u.day) + MONTHS[u.month - 1] + ' ' + pad2(u.hours) + pad2(u.minutes) + 'Z';
        }

        function processRows(rows) {
            let headerIdx = -1;
            for (let i = 0; i < Math.min(rows.length, 10); i++) {
                const vals = rows[i].map(v => String(v).trim());
                if (vals.includes('飞机注册号') && vals.includes('出发地') && vals.includes('到达地') && vals.includes('计划出发')) {
                    headerIdx = i;
                    break;
                }
            }
            if (headerIdx === -1) {
                return {error: '未找到表头行（需包含：飞机注册号、出发地、到达地、计划出发）'};
            }
            const headers = rows[headerIdx].map(h => String(h).trim());
            function findCol(cands) {
                for (const c of cands) {
                    const idx = headers.indexOf(c);
                    if (idx !== -1) return idx;
                }
                for (const c of cands) {
                    for (let i = 0; i < headers.length; i++) {
                        if (headers[i].includes(c)) return i;
                    }
                }
                return -1;
            }
            const colReg      = findCol(['飞机注册号', '注册号', '机号']);
            const colDep      = findCol(['出发地']);
            const colArr      = findCol(['到达地']);
            const colDepDate  = findCol(['出发日期']);
            const colDepTime  = findCol(['计划出发']);
            const colArrDate  = findCol(['到达日期']);
            const colArrTime  = findCol(['预计到达']);
            const colPurpose  = findCol(['用途']);
            const required = {
                '飞机注册号': colReg, '出发地': colDep, '到达地': colArr,
                '出发日期': colDepDate, '计划出发': colDepTime,
                '到达日期': colArrDate, '预计到达': colArrTime, '用途': colPurpose
            };
            for (const name in required) {
                if (required[name] === -1) return {error: '缺少列：' + name};
            }
            const plans = {};
            for (let i = headerIdx + 1; i < rows.length; i++) {
                const r = rows[i];
                if (!r || r.length === 0) continue;
                const get = idx => (idx >= 0 && idx < r.length) ? r[idx] : '';
                const dep       = get(colDep);
                const arr       = get(colArr);
                const depDate   = get(colDepDate);
                const depTimeRaw = get(colDepTime);
                const arrDate   = get(colArrDate);
                const arrTimeRaw = get(colArrTime);
                if (dep === '' || arr === '' || depDate === '' || depTimeRaw === '') continue;
                const depTimeStr = parseTime(depTimeRaw);
                const arrTimeStr = parseTime(arrTimeRaw);
                if (!depTimeStr || !arrTimeStr) continue;
                const depUtc = toUTCLabel(depDate, depTimeStr);
                const arrUtc = toUTCLabel(arrDate, arrTimeStr);
                if (!depUtc || !arrUtc) continue;
                let reg = get(colReg);
                if (reg === '' || reg == null) reg = 'N/A';
                else reg = String(reg).trim();
                const use = String(get(colPurpose) || '');
                const flightType = use.indexOf('调机') !== -1 ? 'FERRY' : 'PAX';
                const line = 'ETD ' + String(dep).trim() + ' ' + formatLabel(depUtc) +
                             ' // ETA ' + String(arr).trim() + ' ' + formatLabel(arrUtc) +
                             '  ' + flightType;
                if (!plans[reg]) plans[reg] = [];
                plans[reg].push({sortKey: depUtc.sortKey, line: line});
            }
            const result = {};
            for (const reg in plans) {
                plans[reg].sort((a, b) => a.sortKey - b.sortKey);
                const lines = [reg];
                plans[reg].forEach(it => lines.push(it.line));
                result[reg] = lines.join('\n');
            }
            return {plans: result};
        }

        function sortPlans(plans) {
            const keys = Object.keys(plans);
            const priorityKeys = PRIORITY.filter(k => keys.indexOf(k) !== -1);
            const remainingKeys = keys.filter(k => PRIORITY.indexOf(k) === -1 && k !== 'N/A').sort();
            const naKeys = keys.filter(k => k === 'N/A');
            const sorted = priorityKeys.concat(remainingKeys, naKeys);
            const result = {};
            sorted.forEach(k => result[k] = plans[k]);
            return result;
        }

        function diffPlans(oldPlans, newPlans) {
            const changes = {};
            const allRegs = new Set(Object.keys(oldPlans || {}).concat(Object.keys(newPlans || {})));
            allRegs.forEach(reg => {
                const oldLines = new Set(((oldPlans && oldPlans[reg]) || '').split('\n'));
                const newLines = new Set(((newPlans && newPlans[reg]) || '').split('\n'));
                oldLines.delete(reg);
                newLines.delete(reg);
                newLines.forEach(line => {
                    if (!oldLines.has(line)) {
                        changes[reg + '\u0001' + line] = 'added';
                    }
                });
            });
            return changes;
        }

        function renderPlans(plans, changes, isRestored) {
            const container = document.getElementById('plans');
            container.innerHTML = '';
            let fullText = '';
            for (const reg in plans) {
                const text = plans[reg];
                const lines = text.split('\n');
                const routes = lines.filter(l => l !== reg);
                const hasChanges = !isRestored && routes.some(line => changes[reg + '\u0001' + line]);
                const block = document.createElement('div');
                block.className = 'reg-block';
                const titleDiv = document.createElement('div');
                titleDiv.className = 'reg-title';
                titleDiv.innerHTML = '✈️ ' + escapeHtml(reg) +
                    (hasChanges ? '<span class="new-flag">🔴 有新增或变更</span>' : '');
                block.appendChild(titleDiv);
                const segList = document.createElement('div');
                segList.className = 'seg-list';
                routes.forEach(line => {
                    const lineDiv = document.createElement('div');
                    lineDiv.className = 'seg-line';
                    lineDiv.textContent = line;
                    segList.appendChild(lineDiv);
                });
                block.appendChild(segList);
                const copyBtn = document.createElement('button');
                copyBtn.textContent = '📋 复制该飞机';
                copyBtn.style.marginTop = '6px';
                copyBtn.style.fontSize = '12px';
                copyBtn.onclick = () => {
                    const copyText = reg + '\n' + routes.join('\n');
                    navigator.clipboard.writeText(copyText).then(() => {
                        copyBtn.textContent = '✅ 已复制';
                        setTimeout(() => { copyBtn.textContent = '📋 复制该飞机'; }, 1500);
                    }).catch(() => {
                        fallbackCopy(copyText);
                        copyBtn.textContent = '✅ 已复制';
                        setTimeout(() => { copyBtn.textContent = '📋 复制该飞机'; }, 1500);
                    });
                };
                block.appendChild(copyBtn);
                container.appendChild(block);
                fullText += reg + '\n' + routes.join('\n') + '\n\n';
            }
            document.getElementById('fullTextBox').textContent = fullText.trim();
        }

        function renderHistory(history) {
            const container = document.getElementById('historyList');
            if (!history.records || history.records.length === 0) {
                container.innerHTML = '<div class="info">暂无历史记录</div>';
                return;
            }
            let html = '<ol>';
            for (const rec of history.records) {
                html += '<li>' + escapeHtml(rec.timestamp) + ' - ' + escapeHtml(rec.filename) + '</li>';
            }
            html += '</ol>';
            container.innerHTML = html;
        }

        function fallbackCopy(text) {
            const ta = document.createElement('textarea');
            ta.value = text;
            ta.style.position = 'fixed';
            ta.style.left = '-9999px';
            document.body.appendChild(ta);
            ta.select();
            try { document.execCommand('copy'); } catch (e) {}
            document.body.removeChild(ta);
        }

        function handleFile(file) {
            const status = document.getElementById('status');
            status.innerHTML = '<div class="info">⏳ 正在读取文件...</div>';
            const reader = new FileReader();
            reader.onload = (ev) => {
                try {
                    const data = new Uint8Array(ev.target.result);
                    const wb = XLSX.read(data, { type: 'array', cellDates: true });
                    const ws = wb.Sheets[wb.SheetNames[0]];
                    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
                    const result = processRows(rows);
                    if (result.error) {
                        status.innerHTML = '<div class="error">❌ ' + escapeHtml(result.error) + '</div>';
                        return;
                    }
                    const newPlans = result.plans;
                    const sortedNewPlans = sortPlans(newPlans);
                    const history = loadHistory();
                    let oldPlans = {};
                    if (history.records.length > 0) {
                        oldPlans = history.records[history.records.length - 1].data || {};
                    }
                    const changes = diffPlans(oldPlans, newPlans);
                    const now = new Date();
                    const timestamp = now.getFullYear() + '-' + pad2(now.getMonth()+1) + '-' + pad2(now.getDate()) +
                                      ' ' + pad2(now.getHours()) + ':' + pad2(now.getMinutes()) + ':' + pad2(now.getSeconds());
                    history.records.push({
                        timestamp: timestamp,
                        filename: file.name,
                        data: newPlans
                    });
                    if (history.records.length > 20) {
                        history.records = history.records.slice(-20);
                    }
                    saveHistory(history);
                    saveJSON(LAST_PLANS_KEY, sortedNewPlans);
                    saveJSON(LAST_FILE_KEY, {name: file.name, timestamp: timestamp});
                    document.getElementById('result').style.display = 'block';
                    status.innerHTML = '<div class="success">✅ 文件读取成功：' +
                        escapeHtml(file.name) + '（历史累计 ' + history.records.length + ' 条）</div>';
                    renderPlans(sortedNewPlans, changes, false);
                    renderHistory(history);
                } catch (err) {
                    status.innerHTML = '<div class="error">❌ 处理失败：' + escapeHtml(err.message) + '</div>';
                    console.error(err);
                }
            };
            reader.readAsArrayBuffer(file);
        }

        window.addEventListener('DOMContentLoaded', () => {
            const lastPlans = loadJSON(LAST_PLANS_KEY);
            const lastFile = loadJSON(LAST_FILE_KEY);
            const history = loadHistory();
            if (lastPlans && Object.keys(lastPlans).length > 0) {
                document.getElementById('result').style.display = 'block';
                const status = document.getElementById('status');
                const info = lastFile
                    ? '（上次加载：' + escapeHtml(lastFile.name) + '，' + escapeHtml(lastFile.timestamp) + '）'
                    : '';
                status.innerHTML = '<div class="info">💾 已恢复上次解析结果 ' + info + '</div>';
                renderPlans(lastPlans, {}, true);
                renderHistory(history);
            }
        });

        document.getElementById('fileInput').addEventListener('change', (e) => {
            const f = e.target.files[0];
            if (f) handleFile(f);
        });

        document.getElementById('clearHistoryBtn').addEventListener('click', () => {
            if (!confirm('确定清除所有历史记录吗？')) return;
            saveHistory({records: []});
            try {
                localStorage.removeItem(LAST_PLANS_KEY);
                localStorage.removeItem(LAST_FILE_KEY);
            } catch (e) {}
            location.reload();
        });

        document.getElementById('copyAllBtn').addEventListener('click', async () => {
            const status = document.getElementById('copyAllStatus');
            const text = document.getElementById('fullTextBox').textContent;
            try {
                await navigator.clipboard.writeText(text);
                status.textContent = '✅ 已复制';
                status.style.color = '#2e7d32';
            } catch (e) {
                fallbackCopy(text);
                status.textContent = '✅ 已复制（降级模式）';
                status.style.color = '#2e7d32';
            }
        });
    </script>
</body>
</html>
"""
    components.html(F_HTML, height=1300, scrolling=True)


# ================================================================
# 功能3：航路处理工具
# ================================================================
with tab3:
    st.markdown("支持表格格式（带N/E坐标）/中文描述格式，自动精简航路+添加#前缀，兼容不规整数据")

    if "last_processed_input_route" not in st.session_state:
        st.session_state.last_processed_input_route = ""
    if "result_text_route" not in st.session_state:
        st.session_state.result_text_route = ""

    def parse_coord(coord_str):
        letter = coord_str[0]
        num_part = coord_str[1:]
        if letter == 'N':
            deg = int(num_part[0:2])
            minute = int(num_part[2:4])
            sec_part = num_part[4:]
            if '.' in sec_part:
                sec_float = float(sec_part)
                sec_int = int(round(sec_float))
            else:
                sec_int = int(sec_part)
            if sec_int >= 60:
                sec_int -= 60
                minute += 1
                if minute >= 60:
                    minute -= 60
                    deg += 1
            return f"{deg:02d}{minute:02d}{sec_int:02d}"
        elif letter == 'E':
            deg = int(num_part[0:3])
            minute = int(num_part[3:5])
            sec_part = num_part[5:]
            if '.' in sec_part:
                sec_float = float(sec_part)
                sec_int = int(round(sec_float))
            else:
                sec_int = int(sec_part)
            if sec_int >= 60:
                sec_int -= 60
                minute += 1
                if minute >= 60:
                    minute -= 60
                    deg += 1
            return f"{deg:03d}{minute:02d}{sec_int:02d}"
        else:
            raise ValueError(f"未知的坐标前缀: {letter}")

    def base_name(s):
        return s.split('@')[0]

    def is_open_point(s):
        base = base_name(s)
        if re.match(r'^[A-Z]{2,5}$', base):
            return True
        if re.match(r'^P[A-Z]+$', base):
            return True
        return False

    def is_p_point(s):
        base = base_name(s)
        return re.match(r'^P\d+$', base) is not None

    def clean_route(r):
        if r.startswith('#'):
            return r[1:]
        return r

    def is_open_route(rt):
        return rt and rt[0] not in ('H', 'J', 'V')

    def extract_table(text):
        tokens = text.strip().split()
        start_idx = 0
        for i, tok in enumerate(tokens):
            if tok.isdigit() and 1 <= int(tok) <= 40:
                start_idx = i
                break
        tokens = tokens[start_idx:]

        lines = []
        i = 0
        while i < len(tokens):
            if tokens[i].isdigit():
                line = [tokens[i]]
                i += 1
                while i < len(tokens) and not tokens[i].isdigit():
                    line.append(tokens[i])
                    i += 1
                lines.append(line)

        points = []
        routes = []
        for line in lines:
            lat_idx = None
            for idx, tok in enumerate(line):
                if tok.startswith('N') and tok[1:].replace('.', '', 1).isdigit():
                    lat_idx = idx
                    break
            if lat_idx is None:
                continue
            lon_idx = lat_idx + 1
            if lon_idx >= len(line) or not line[lon_idx].startswith('E'):
                continue
            lat_str = line[lat_idx]
            lon_str = line[lon_idx]

            route = None
            if lon_idx + 1 < len(line):
                next_tok = line[lon_idx + 1]
                if re.match(r'^[A-Z][A-Z0-9]*$', next_tok) and not next_tok[0].isdigit():
                    route = next_tok

            point_name = None
            for j in range(lat_idx - 1, 0, -1):
                tok = line[j]
                if is_open_point(tok) or is_p_point(tok):
                    point_name = tok
                    break
            if point_name is None:
                continue

            if is_p_point(point_name):
                lat_int = parse_coord(lat_str)
                lon_int = parse_coord(lon_str)
                point_display = f"{point_name}@{lat_int}N{lon_int}E"
            else:
                point_display = point_name

            points.append(point_display)
            if route is not None:
                routes.append(route)

        seq = []
        for i in range(len(points)):
            seq.append(points[i])
            if i < len(routes):
                seq.append(routes[i])
        return seq

    def extract_chinese(text):
        text = re.sub(r'[\u4e00-\u9fa5，、。；：""''（）【】]', ' ', text)
        words = text.split()
        seq = []
        for w in words:
            if '(' in w and ')' in w:
                m = re.search(r'\(([A-Z]+)\)', w)
                if m:
                    point = m.group(1)
                    prefix = w[:w.find('(')]
                    m_route = re.search(r'([A-Z]\d+)$', prefix)
                    if m_route:
                        seq.append(m_route.group(1))
                    seq.append(point)
            elif re.match(r'^[A-Z]\d+[A-Z]{2,5}$', w) or re.match(r'^[A-Z]\d+P\d+$', w):
                m = re.match(r'^([A-Z]\d+)([A-Z]{2,5}|P\d+)$', w)
                if m:
                    seq.append(m.group(1))
                    seq.append(m.group(2))
            elif re.match(r'^[A-Z]\d+$', w):
                seq.append(w)
            elif is_open_point(w) or is_p_point(w):
                seq.append(w)
        return seq

    def step1_extract(text):
        if re.search(r'N\d{5,6}(?:\.\d+)?\s+E\d{6,7}(?:\.\d+)?', text):
            return extract_table(text), 'table'
        else:
            return extract_chinese(text), 'chinese'

    def step2_reduce(seq):
        L = seq[:]
        changed = True
        while changed:
            changed = False
            n = len(L)
            candidates = []
            for i in range(0, n, 2):
                if not is_open_point(L[i]):
                    continue
                if i + 1 >= n:
                    continue
                first_route = clean_route(L[i+1])
                if not is_open_route(first_route):
                    continue
                for j in range(i+2, n, 2):
                    all_same = True
                    for k in range(i+1, j, 2):
                        rt = clean_route(L[k])
                        if rt != first_route or not is_open_route(rt):
                            all_same = False
                            break
                    if not all_same:
                        break
                    if is_open_point(L[j]):
                        length = (j - i) // 2
                        if length >= 2:
                            candidates.append((i, j, length))
            if not candidates:
                break
            candidates.sort(key=lambda x: -x[2])
            best_i, best_j, _ = candidates[0]
            new_segment = [L[best_i], L[best_i+1], L[best_j]]
            L = L[:best_i] + new_segment + L[best_j+1:]
            changed = True
        return L

    def step3_add_hash(seq):
        pts = seq[0::2]
        rts = seq[1::2]
        m = len(rts)

        def is_closed_route(rt):
            return rt.startswith(('H', 'J', 'V'))

        def is_p(pt):
            base = base_name(pt)
            return re.match(r'^P\d+$', base) is not None

        res = [pts[0]]
        for i, rt in enumerate(rts):
            left = pts[i]
            right = pts[i+1]
            need_hash = False
            if is_closed_route(rt):
                need_hash = True
            elif is_p(left) or is_p(right):
                need_hash = True
            res.append('#' + rt if need_hash else rt)
            res.append(right)
        return res

    # ---------- 功能3 UI ----------
    st.markdown("""
        <style>
        .stButton>button {border-radius: 8px; height: 2.5rem; font-size: 1rem;}
        .stProgress>div>div {background-color: #1890ff;}
        .stCaption {color: #666666; font-size: 0.9rem;}
        </style>
    """, unsafe_allow_html=True)

    input_text_route = st.text_area(
        "📋 请输入待处理的航路文本",
        key="input_text_route",
        height=300,
        placeholder="粘贴民航航线数据，支持多行表格格式/纯中文描述格式..."
    )

    btn_col1, btn_col2, btn_col3 = st.columns([2, 2, 8])
    with btn_col1:
        process_btn = st.button("⚙️ 处理", type="primary", use_container_width=True, key="process_route")
    with btn_col2:
        clear_btn = st.button("🗑️ 清空", use_container_width=True, key="clear_route")

    if clear_btn:
        st.session_state.input_text_route = ""
        st.session_state.last_processed_input_route = ""
        st.session_state.result_text_route = ""
        st.rerun()

    if process_btn and st.session_state.get("input_text_route", "").strip():
        progress_bar = st.progress(0)
        status_text = st.empty()
        total_steps = 4
        current_step = 0

        try:
            current_step += 1
            progress_bar.progress(current_step / total_steps)
            status_text.text(f"处理中：第{current_step}步/共{total_steps}步（识别输入类型）")
            seq, fmt = step1_extract(st.session_state.input_text_route)

            current_step += 1
            progress_bar.progress(current_step / total_steps)
            status_text.text(f"处理中：第{current_step}步/共{total_steps}步（精简相同开放航路）")
            if fmt == 'table':
                seq = step2_reduce(seq)

            current_step += 1
            progress_bar.progress(current_step / total_steps)
            status_text.text(f"处理中：第{current_step}步/共{total_steps}步（添加航路#前缀）")
            if fmt == 'table':
                seq = step3_add_hash(seq)

            current_step += 1
            progress_bar.progress(current_step / total_steps)
            status_text.text(f"处理中：第{current_step}步/共{total_steps}步（生成最终结果）")
            result = ' '.join(seq) if seq else "⚠️ 未提取到有效航路数据"

            st.session_state.result_text_route = result
            st.session_state.last_processed_input_route = st.session_state.input_text_route

            progress_bar.empty()
            status_text.empty()
            st.success("✅ 处理完成！结果如下：")

        except Exception as e:
            progress_bar.empty()
            status_text.empty()
            st.error(f"❌ 处理失败：{str(e)}")
            with st.expander("🔍 查看详细错误信息", expanded=False):
                st.code(traceback.format_exc(), language="text")

    if st.session_state.get("result_text_route", ""):
        current_input = st.session_state.get("input_text_route", "")
        last_input = st.session_state.last_processed_input_route

        st.subheader("📊 处理结果", divider="blue")

        if current_input != last_input:
            st.warning("⚠️ 输入已更改，当前显示的是上一次处理的结果，如需更新请点击「处理」按钮。")

        st.code(st.session_state.result_text_route, language="text")

    if not st.session_state.get("result_text_route", "") and not st.session_state.get("input_text_route", "").strip():
        st.info("💡 提示：粘贴航路数据后，点击「处理」即可，支持30+行不规整表格数据")

    st.markdown("---")
    st.caption("✈️ 支持表格格式（带N/E坐标）/中文描述格式，自动精简航路+添加#前缀")
