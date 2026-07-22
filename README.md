# 通用声明 → 公务飞行计划信息备案表 生成器

上传通用声明 Excel 和备案表模板，自动提取机组/乘客信息，填充到模板中。

## 功能
- 解析通用声明中的 **OPERATOR, REG, FLT, AC TYPE, FROM/TO, DATE/TIME**
- 提取 **机组人员**（顺序：机长、副驾驶、乘务、机务），自动提取中文姓名，若无则保留英文
- 提取 **乘客列表**，自动转换国籍（CHN→中国等），根据证件号自动判断身份证/护照
- 填充到备案表模板，保留模板格式
- 下载填充后的 Excel 文件

## 使用方法
1. 上传通用声明 Excel（如 `MLLINGen DecZGSZ-ZLXY 23Jul.xlsx`）
2. 上传备案表模板 Excel（如 `07月23日 MLLIN 深圳-西安.xls`）
3. 点击下载生成的文件

## 部署
适配 [Streamlit Cloud](https://streamlit.io/cloud)，直接连接 GitHub 仓库即可部署。
