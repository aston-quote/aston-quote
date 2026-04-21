import json
import os
import openpyxl
import re

# 配置文件路径
excel_path = '/Users/kris/Desktop/AKI/ASTON报价软件数据库设计.xlsx'
html_path = '/Users/kris/Desktop/AKI/aston-quote-web/index.html'

def clean_val(v):
    if v == "#REF!" or v is None: return 0
    if isinstance(v, (int, float)): return float(v)
    try:
        match = re.search(r"(\d+\.?\d*)", str(v))
        if match: return float(match.group(1))
    except: pass
    return 0

try:
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    db = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = list(ws.rows)
        if not rows: continue
        headers = [str(cell.value).strip() if cell.value else f"Col_{i}" for i, cell in enumerate(rows[0])]
        sheet_data = []
        for row in rows[1:]:
            row_dict = {headers[i]: row[i].value for i in range(min(len(headers), len(row)))}
            sheet_data.append(row_dict)
        db[sheet_name] = sheet_data

    # 1. 提取款式与图片
    js_styles = []
    for s in db.get("尺码用料矩阵", []):
        name = s.get("基础版型名称")
        if not name: continue
        img = str(s.get("款式图片链接", "")).strip()
        # 兜底图逻辑
        if not img or img == "None" or not img.startswith("http"):
            img = "https://via.placeholder.com/300?text=" + name.replace(" ", "+")
        
        js_styles.append({
            "name": str(name), "img": img,
            "baseLabor": clean_val(s.get("基础工费", 25.0)),
            "sizes": {k: clean_val(v) for k, v in s.items() if k not in ["基础版型名称", "款式图片链接", "描述", "基础工费"]}
        })

    # 2. 提取工艺
    js_processes = []
    for p in db.get("工艺费率", []):
        name = p.get("工艺/辅料名称")
        if not name: continue
        targets = str(p.get("适用版型", "通用")).replace("，", ",").split(",")
        js_processes.append({
            "name": str(name), "price": clean_val(p.get("基础单价")), 
            "targets": [t.strip() for t in targets]
        })

    # 3. 提取面料
    js_fabrics = []
    for f in db.get("面料库", []):
        name = f.get("面料名称")
        if not name: continue
        js_fabrics.append({
            "id": str(f.get("ID", hash(name))), "name": str(name),
            "price_loose": clean_val(f.get("平方面料单价(散剪)", 0)),
            "rollQty": clean_val(f.get("起订量（KG或米）", 20.0))
        })

    js_modifiers = [{"name": str(m["修正名称"]), "factor": clean_val(m.get("系数", 1.0)), "target": str(m.get("适用版型", "通用"))} for m in db.get("版型修正", []) if m.get("修正名称")]
    
    # 读取现有的 HTML 模板并替换数据部分
    with open(html_path, 'r') as f:
        content = f.read()

    # 构造新的 DB JS 对象
    db_js = f"const DB = {{ styles: {json.dumps(js_styles, ensure_ascii=False)}, processes: {json.dumps(js_processes, ensure_ascii=False)}, modifiers: {json.dumps(js_modifiers, ensure_ascii=False)}, fabrics: {json.dumps(js_fabrics, ensure_ascii=False)} }};"
    
    # 使用正则替换 HTML 里的 const DB 部分
    new_content = re.sub(r"const DB = \{.*?\};", db_js, content, flags=re.DOTALL)
    
    with open(html_path, 'w') as f:
        f.write(new_content)
    
    print("✅ 数据已成功从 Excel 注入到 index.html")

except Exception as e:
    print(f"❌ 运行失败: {e}")
