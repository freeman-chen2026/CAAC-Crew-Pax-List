import streamlit as st
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
# 功能1：备案表生成（原有完整代码，此处省略，实际部署时保留）
# ================================================================
with tab1:
    st.markdown("上传 GD单 和模板，自动生成备案表（联系方式、执照号码及证件号码已内置）。")
    # ...（此处省略，使用之前完整代码）...

# ================================================================
# 功能2：世界时行程（带记忆对比功能）
# ================================================================
with tab2:
    st.markdown("从Jetops系统导出的北京时间的行程 Excel 文件转换为世界时的行程，便于复制粘贴。")
    st.info("💡 每次上传将自动与上一次记录对比，新增或变更的航段会标红。")

    # ---------- 历史数据管理 ----------
    HISTORY_FILE = "flight_history.json"

    def load_history():
        """加载历史记录"""
        if os.path.exists(HISTORY_FILE):
            try:
                with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except:
                return {"records": []}
        else:
            return {"records": []}

    def save_history(history):
        """保存历史记录"""
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(history, f, ensure_ascii=False, indent=2)

    def parse_time_column(val):
        if pd.isna(val):
            return None
        if isinstance(val, (pd.Timestamp, datetime)):
            return val.strftime('%H:%M')
        if hasattr(val, 'strftime'):
            return val.strftime('%H:%M')
        s = str(val).strip()
        if ':' in s:
            return s[:5]
        return s

    def convert_to_utc(date_val, time_str):
        if pd.isna(date_val) or time_str is None:
            return None
        if isinstance(date_val, (pd.Timestamp, datetime)):
            date_str = date_val.strftime('%Y-%m-%d')
        else:
            date_str = str(date_val).split()[0]
        dt_str = f"{date_str} {time_str}"
        try:
            dt_local = pd.to_datetime(dt_str)
            dt_utc = dt_local - timedelta(hours=8)
            return dt_utc
        except:
            return None

    def format_utc(dt):
        if dt is None:
            return ""
        months = ['JAN','FEB','MAR','APR','MAY','JUN',
                  'JUL','AUG','SEP','OCT','NOV','DEC']
        day = dt.day
        month = months[dt.month - 1]
        hour = dt.hour
        minute = dt.minute
        return f"{day:02d}{month} {hour:02d}{minute:02d}Z"

    def generate_plans(df):
        required = ['飞机注册号', '出发地', '到达地', '出发日期', '计划出发', '到达日期', '预计到达', '用途']
        for col in required:
            if col not in df.columns:
                st.error(f"❌ 缺少列：{col}")
                return None

        df = df.dropna(subset=['出发地', '到达地', '出发日期', '计划出发'])
        if df.empty:
            st.warning("没有有效的航段数据")
            return None

        plans = {}
        for idx, row in df.iterrows():
            reg = row['飞机注册号']
            if pd.isna(reg) or str(reg).strip() == '':
                reg = "N/A"
            else:
                reg = str(reg).strip()

            dep_time = parse_time_column(row['计划出发'])
            arr_time = parse_time_column(row['预计到达'])
            if dep_time is None or arr_time is None:
                continue

            dep_utc = convert_to_utc(row['出发日期'], dep_time)
            arr_utc = convert_to_utc(row['到达日期'], arr_time)
            if dep_utc is None or arr_utc is None:
                continue

            use = str(row['用途']) if not pd.isna(row['用途']) else ''
            flight_type = 'FERRY' if '调机' in use else 'PAX'

            line = (f"ETD {row['出发地']} {format_utc(dep_utc)} // "
                    f"ETA {row['到达地']} {format_utc(arr_utc)}  {flight_type}")

            if reg not in plans:
                plans[reg] = []
            plans[reg].append((dep_utc, line))

        result = {}
        for reg, items in plans.items():
            items.sort(key=lambda x: x[0])
            lines = [reg]
            lines.extend([item[1] for item in items])
            result[reg] = "\n".join(lines)

        return result

    def sort_plans(plans_dict):
        priority_order = ['B652Q', 'B65AP', 'B652S', 'MLLIN', 'N88AY', 'B652R']
        all_keys = list(plans_dict.keys())
        priority_keys = [k for k in priority_order if k in all_keys]
        remaining_keys = [k for k in all_keys if k not in priority_order and k != "N/A"]
        remaining_keys.sort()
        na_keys = [k for k in all_keys if k == "N/A"]
        sorted_keys = priority_keys + remaining_keys + na_keys
        return {k: plans_dict[k] for k in sorted_keys}

    def diff_plans(old_plans, new_plans):
        """对比新旧计划，返回变更集合: (reg, line) -> 'added' or 'modified'"""
        changes = {}
        # 比较每个注册号下的航段
        all_regs = set(old_plans.keys()) | set(new_plans.keys())
        for reg in all_regs:
            old_lines = set(old_plans.get(reg, "").split('\n')) if old_plans.get(reg) else set()
            new_lines = set(new_plans.get(reg, "").split('\n')) if new_plans.get(reg) else set()
            # 去除第一行（注册号本身）
            old_lines.discard(reg)
            new_lines.discard(reg)
            # 新增
            added = new_lines - old_lines
            for line in added:
                changes[(reg, line)] = 'added'
            # 删除（我们不标记删除，只标记新增）
            # 修改：如果同一行内容不同，但这里以整行为单位，所以如果行内容变化则视为删除+新增，我们标记为modified
            # 简单方式：比较行集合，如果新行不在旧行中，则视为新增
            # 由于我们只关心新增，不关心删除，所以只标记added
        return changes

    # ---------- UI ----------
    uploaded_file_2 = st.file_uploader("📤 上传航段数据导出（北京时间）", type=["xlsx"], key="worldtime")

    if uploaded_file_2 is not None:
        try:
            df = pd.read_excel(uploaded_file_2, skiprows=1)
            st.success("✅ 文件读取成功")

            new_plans = generate_plans(df)
            if new_plans is None:
                st.stop()

            sorted_new_plans = sort_plans(new_plans)

            # 加载历史
            history = load_history()
            # 上一次的计划（取最后一条记录的data，如果有）
            old_plans = {}
            if history["records"]:
                last_record = history["records"][-1]
                old_plans = last_record.get("data", {})

            # 对比变更
            changes = diff_plans(old_plans, new_plans)

            # 保存新记录到历史
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            filename = uploaded_file_2.name
            new_record = {
                "timestamp": timestamp,
                "filename": filename,
                "data": new_plans
            }
            history["records"].append(new_record)
            # 限制历史记录数量（可选）
            if len(history["records"]) > 20:
                history["records"] = history["records"][-20:]
            save_history(history)

            # ---------- 显示结果（带标红） ----------
            st.subheader("📋 生成的飞行计划（红色为新增/变更）")

            # 构建带颜色的显示文本
            display_text = ""
            for reg, text in sorted_new_plans.items():
                lines = text.split('\n')
                display_text += f"**{reg}**\n\n"
                for line in lines:
                    if line == reg:
                        continue
                    if (reg, line) in changes:
                        # 标红
                        display_text += f'<span style="color:red">{line}</span>\n'
                    else:
                        display_text += f'{line}\n'
                display_text += "\n"

            st.markdown(display_text, unsafe_allow_html=True)
            st.caption("🔴 红色航段表示本次新增或内容有变更")

            # 同时提供纯文本复制
            full_text = ""
            for reg, text in sorted_new_plans.items():
                full_text += f"{text}\n\n"
            st.text_area("📦 纯文本版本（可直接复制）", full_text, height=250)

            # ---------- 历史记录展开 ----------
            with st.expander("📜 查看历史上传记录"):
                if history["records"]:
                    for i, rec in enumerate(history["records"]):
                        st.write(f"{i+1}. {rec['timestamp']} - {rec['filename']}")
                else:
                    st.write("暂无历史记录")

            # 清除历史按钮
            if st.button("🗑️ 清除所有历史记录", key="clear_history"):
                save_history({"records": []})
                st.success("历史已清除，请刷新页面")
                st.rerun()

        except Exception as e:
            st.error(f"❌ 处理出错：{e}")
            st.stop()
    else:
        st.info("请上传一个符合格式的 Excel 文件。")

    st.markdown("---")
    st.caption("🛠️ 工具说明：日期/时间按北京时间（UTC+8）自动转换为世界时（Z）。对比功能基于上一次上传的记录。")


# ================================================================
# 功能3：航路处理工具
# ================================================================
with tab3:
    st.markdown("支持表格格式（带N/E坐标）/中文描述格式，自动精简航路+添加#前缀，兼容不规整数据")

    # ...（此处省略，使用之前完整代码）...
