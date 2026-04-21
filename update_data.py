import json
import os
import openpyxl
import re

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
    db = {sn: [] for sn in wb.sheetnames}
    for sn in wb.sheetnames:
        ws = wb[sn]
        rows = list(ws.rows)
        if not rows: continue
        headers = [str(cell.value).strip() if cell.value else f"Col_{i}" for i, cell in enumerate(rows[0])]
        for row in rows[1:]:
            db[sn].append({headers[i]: row[i].value for i in range(min(len(headers), len(row)))})

    # 1. 款式
    js_styles = []
    for s in db.get("尺码用料矩阵", []):
        name = s.get("基础版型名称")
        if not name: continue
        img = str(s.get("款式图片链接", "")).strip()
        # 强制检测：如果链接不是以http开头，自动尝试补全或使用占位图
        if not img or img.lower() == "none" or not img.startswith("http"):
            img = f"https://via.placeholder.com/300?text={str(name).replace(' ', '+')}"
        
        js_styles.append({
            "name": str(name), "img": img,
            "baseLabor": clean_val(s.get("基础工费", 25.0)),
            "sizes": {k: clean_val(v) for k, v in s.items() if k not in ["基础版型名称", "款式图片链接", "描述", "基础工费"]}
        })

    # 2. 其他数据
    js_proc = [{"name": str(p["工艺/辅料名称"]), "price": clean_val(p.get("基础单价")), "targets": str(p.get("适用版型", "通用")).replace("，", ",").split(",")} for p in db.get("工艺费率", []) if p.get("工艺/辅料名称")]
    js_fabrics = [{"id": str(f.get("ID", hash(f.get("面料名称")))), "name": str(f.get("面料名称")), "price_loose": clean_val(f.get("平方面料单价(散剪)", 0))} for f in db.get("面料库", []) if f.get("面料名称")]
    js_mods = [{"name": str(m["修正名称"]), "factor": clean_val(m.get("系数", 1.0)), "target": str(m.get("适用版型", "通用"))} for m in db.get("版型修正", []) if m.get("修正名称")]
    
    # 3. 业务员与国家 (找回丢失项)
    js_sales = ["Alicia 李丽", "Caitlin 赵萌萌", "Elsa吕晓娈", "Bella白惠文", "Linn刘逍", "Allie王安冉", "Nancy王钰静", "Cecily郑晨希", "Casey裴枭君", "Cloud刘佩旺", "Tessa柴瑞蝶", "Taylor李倩", "Aria刘思彤", "Rowan张家乐", "秦宜萍", "张茹冰"]
    # 尝试从Excel提取国家，如果没有则使用默认常用国家
    js_countries = sorted(list(set([str(c.get("国家名称", "")).strip() for c in db.get("国家库", []) if c.get("国家名称")])))
    if not js_countries: js_countries = ["USA", "UK", "Australia", "Canada", "Germany", "France"]

    # 注入数据
    db_full = {
        "styles": js_styles, "processes": js_proc, "modifiers": js_mods, 
        "fabrics": js_fabrics, "sales": js_sales, "countries": js_countries
    }
    
    with open(html_path, 'r') as f: content = f.read()
    new_db_js = f"const DB = {json.dumps(db_full, ensure_ascii=False)};"
    new_content = re.sub(r"const DB = \{.*?\};", new_db_js, content, flags=re.DOTALL)
    
    with open(html_path, 'w') as f: f.write(new_content)
    print("✅ V7.1 数据同步完成（含国家库与图片增强）")

except Exception as e:
    print(f"❌ 运行失败: {e}")
