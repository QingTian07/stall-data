#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""同步内核：解析销售文字 + 更新《销售、库存、经营分析.xlsx》全部联动。

被 Streamlit 网页（app.py）调用，也可独立测试。
任何校验失败抛 SyncError，不修改文件。
"""
import os, re
from copy import copy
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

class SyncError(Exception):
    pass

# 兼职手写/输入常见简称 → 标准名（持续补充）
ALIASES = {
    "亚索剑": "大剑", "疾风亚索剑": "大剑", "亚索": "大剑", "疾风亚索": "大剑",
    "机甲剑": "机甲伸缩剑", "机架剑": "机甲伸缩剑", "机甲": "机甲伸缩剑", "机械剑": "机甲伸缩剑",
    "弯月剑": "中剑", "弯月": "中剑", "机甲剑": "中剑", "机甲": "中剑", "机械剑": "中剑",
    "中弯月": "中剑", "小弯月": "小剑", "小弯月剑": "小剑",
    "竹筒剑": "竹筒伸缩剑", "猫爪剑": "猫爪伸缩剑", "猫爪": "猫爪伸缩剑",
    "链子剑": "雷公鞭子剑", "雷公鞭": "雷公鞭子剑", "鞭子剑": "雷公鞭子剑",
    "三头龙": "三头龙基多拉", "基多拉": "三头龙基多拉",
    "雪糕萝卜塔": "冰淇淋萝卜塔", "冰淇凌塔": "冰淇淋萝卜塔", "冰淇淋塔": "冰淇淋萝卜塔",
    "雪糕塔": "冰淇淋萝卜塔", "雪糕": "冰淇淋萝卜塔", "冰淇淋": "冰淇淋萝卜塔",
    "小桌面萝卜塔": "薄桌面萝卜塔", "薄桌面": "薄桌面萝卜塔", "小桌面": "薄桌面萝卜塔",
    "长柄萝卜": "长柄萝卜塔", "花瓣萝卜": "花瓣萝卜塔", "萝卜塔": "长柄萝卜塔",
    "螺旋锥": "螺旋锥组", "葫芦": "葫芦螺旋锥组", "冰山": "冰山螺旋锥组", "方形螺旋锥": "方形螺旋锥组",
    "魔方": "无限魔方", "浑天仪": "浑天仪小球", "机械手": "机械手臂",
    "恐龙夹": "恐龙夹子", "夹子": "恐龙夹子", "按压球": "解压按压球", "解压球": "解压按压球",
    "指间陀螺": "指尖陀螺", "陀螺": "指尖陀螺",
    "中恐龙": "中龙", "小恐龙蛇": "迷你蛇", "小蛇": "迷你蛇", "小龙蛇": "小龙",
}
_CN_NUM = {"一": 1, "二": 2, "两": 2, "三": 3, "四": 4, "五": 5,
           "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}

def resolve(name, known):
    """名称 → 标准名；匹配不到或不唯一返回 None"""
    name = name.strip()
    if not name:
        return None
    if name in known:
        return name
    if name in ALIASES and ALIASES[name] in known:
        return ALIASES[name]
    hits = [k for k in known if name in k or k in name]
    if len(hits) == 1:
        return hits[0]
    return None

def parse_sales_text(text):
    """自由文字 → [{name, qty, amt, dmg}]；返回 (items, problems)
    支持示例：
      寄居蟹 1 10
      小蛇 8
      两个长柄萝卜塔 50
      损:猫爪剑 1
    """
    items, problems = {}, []
    for raw in re.split(r"[\n；;，,、]+", text or ""):
        line = raw.strip()
        if not line:
            continue
        dmg = bool(re.match(r"^损(坏)?[:：\s]", line))
        if dmg:
            line = re.sub(r"^损(坏)?[:：\s]*", "", line)
        qty_cn = None
        m = re.match(r"^([一二两三四五六七八九十]+)个\s*(.+)$", line)
        if m:
            qty_cn = _CN_NUM.get(m.group(1)[-1], None) if len(m.group(1)) == 1 else \
                (_CN_NUM.get(m.group(1)[0], 0) * 10 if m.group(1).startswith("十")
                 else _CN_NUM.get(m.group(1)[0], 0) + 10 if len(m.group(1)) == 2 and m.group(1)[1] == "十"
                 else None)
            if qty_cn is None:
                qty_cn = None
            line = m.group(2)
        nums = [float(x) for x in re.findall(r"\d+(?:\.\d+)?", line)]
        name = re.sub(r"\d+(?:\.\d+)?", "", line)
        name = re.sub(r"[元块个件支\s\.]+", "", name)
        if not name:
            problems.append(f"这一行没认出商品名：「{raw.strip()}」")
            continue
        if dmg:
            qty = int(nums[0]) if nums else 1
            amt = None
        elif qty_cn:
            qty, amt = qty_cn, (nums[0] if nums else None)
        elif len(nums) == 0:
            qty, amt = 1, None
        elif len(nums) == 1:
            qty, amt = 1, nums[0]
        else:
            qty, amt = int(nums[0]), nums[1]
        key = (name, dmg)
        t = items.setdefault(key, {"name": name, "qty": 0, "amt": 0.0, "dmg": dmg})
        t["qty"] += qty
        t["amt"] = (t["amt"] or 0) + (amt or 0)
        t["amt"] = t["amt"] or None
    return list(items.values()), problems

# ---------------- 同步引擎（适配自 process_sales.py，脱离录入表直写） ----------------

def _stock_map(stock):
    names, rows = {}, {}
    for r in range(5, 80):
        if isinstance(stock.cell(row=r, column=2).value, int):
            nm = stock.cell(row=r, column=3).value
            if isinstance(nm, str):
                names[nm] = stock.cell(row=r, column=5).value or 0  # E列=库存（件）
                rows[nm] = r
    return names, rows

def _cost_map(price_ws):
    return {price_ws.cell(row=r, column=3).value: price_ws.cell(row=r, column=4).value
            for r in range(5, 60) if price_ws.cell(row=r, column=3).value}

def apply_sync(inv_path, price_path, date_str, staff, items, wage=None):
    """items: [{name(标准名), qty, amt, dmg}]；wage: 当天实付工资（元，可空）
    成功返回报告 dict；失败抛 SyncError（文件未被修改）。"""
    wb = load_workbook(inv_path)
    stock, ledger, daily, entry = wb["当前库存"], wb["销售明细"], wb["销售记录"], wb["销售录入"]
    cost_map = _cost_map(load_workbook(price_path, read_only=True)["成本利润表"])

    entry_names = set()
    for r in range(6, 80):
        nm = entry.cell(row=r, column=2).value
        if isinstance(nm, str) and isinstance(entry.cell(row=r, column=3).value, (int, float)):
            entry_names.add(nm)
    stock_names, stock_row = _stock_map(stock)

    staff = (staff or "").strip() or "—"
    records = []
    for it in items:
        nm = it["name"]
        if nm not in stock_names:
            raise SyncError(f"商品「{nm}」不在当前库存名单中")
        qty, amt, dmg = int(it.get("qty") or 0), it.get("amt"), bool(it.get("dmg"))
        if qty <= 0:
            raise SyncError(f"{nm} 数量应为正整数")
        records.append((nm, qty, float(amt) if amt else None, dmg))

    need = {}
    for nm, q, _, dmg in records:
        need[nm] = need.get(nm, 0) + q
    short = {n: (stock_names[n], q) for n, q in need.items() if stock_names[n] < q}
    if short:
        raise SyncError("库存不足：" + "；".join(
            f"{n} 现存{h}<需扣{q}" for n, (h, q) in short.items()))

    # ---- 写销售明细（预留空白行优先，不够在合计行上方插行）----
    total_row = next(r for r in range(6, ledger.max_row + 1)
                     if str(ledger.cell(row=r, column=2).value or "").startswith("合计"))
    last_data = max(r for r in range(6, total_row) if ledger.cell(row=r, column=3).value)
    blanks = list(range(last_data + 1, total_row))
    deficit = len(records) - len(blanks)
    if deficit > 0:
        unmerged = []
        for m in list(ledger.merged_cells.ranges):
            if m.min_row >= total_row:
                unmerged.append((m.min_row, m.max_col))
                ledger.unmerge_cells(str(m))
        ledger.insert_rows(total_row, deficit)
        for i in range(deficit):
            r_new = total_row + i
            for cc in range(2, 9):
                ledger.cell(row=r_new, column=cc)._style = copy(
                    ledger.cell(row=last_data, column=cc)._style)
            ledger.row_dimensions[r_new].height = 22
        for min_row, max_col in unmerged:
            ledger.merge_cells(start_row=min_row + deficit, start_column=2,
                               end_row=min_row + deficit, end_column=max_col)
        blanks = list(range(last_data + 1, total_row + deficit))
    for i, (nm, q, amt, dmg) in enumerate(records):
        r = blanks[i]
        vals = (date_str, nm, 0 if dmg else q, q if dmg else 0, None if dmg else amt, staff, None)
        for cc, v in zip(range(2, 9), vals):
            ledger.cell(row=r, column=cc, value=v)

    # ---- 扣库存 ----
    for nm, q in need.items():
        stock.cell(row=stock_row[nm], column=5, value=stock_names[nm] - q)

    # ---- 销售记录（按日）加行 ----
    r_total = next(r for r in range(6, daily.max_row + 1)
                   if str(daily.cell(row=r, column=2).value or "").startswith("合计"))
    last_d = max(r for r in range(6, r_total) if daily.cell(row=r, column=2).value not in (None, ""))
    daily.insert_rows(r_total, 1)
    for cc in range(2, 13):
        daily.cell(row=r_total, column=cc)._style = copy(daily.cell(row=last_d, column=cc)._style)
    daily.row_dimensions[r_total].height = 22
    amt_sum = round(sum(a for _, _, a, _ in records if a), 2)
    cost_sum = round(sum((cost_map.get(nm) or 0) * q for nm, q, _, dmg in records if not dmg), 2)
    f_scheme = ('=IF(F{r}="","",CEILING($O$5+MIN(F{r},$O$6)*$O$7'
                '+MAX(0,MIN(F{r},$O$8)-$O$6)*$O$9'
                '+MAX(0,F{r}-$O$8)*$O$10,1))').format(r=r_total)
    sold_cnt = sum(q for _, q, _, d in records if not d)
    dmg_cnt = sum(q for _, q, _, d in records if d)
    for cc, v in ((2, date_str), (3, staff), (4, sold_cnt), (5, dmg_cnt),
                  (6, amt_sum if amt_sum else None), (7, wage),
                  (8, f_scheme),
                  (9, f'=IF(OR(F{r_total}="",G{r_total}=""),"",ROUND(G{r_total}-H{r_total},2))'),
                  (10, cost_sum if cost_sum else None),
                  (11, f'=IF(OR(F{r_total}="",G{r_total}=""),"",ROUND(F{r_total}-G{r_total}-J{r_total},2))'),
                  (12, None)):
        daily.cell(row=r_total, column=cc, value=v)
    for cc in range(4, 12):
        letter = get_column_letter(cc)
        daily.cell(row=r_total + 1, column=cc,
                   value=f"=SUM({letter}6:{letter}{r_total})")

    # ---- 刷新产品成本率详情（全量重算+按销量重排；未定价款不在表内）----
    a = wb["产品成本率详情"]
    info = {}
    for r in range(6, 50):
        nm = a.cell(row=r, column=3).value
        if isinstance(nm, str) and nm != "兼职工资累计":
            info[nm] = (a.cell(row=r, column=4).value, a.cell(row=r, column=5).value,
                        a.cell(row=r, column=6).value)
    stat = {}
    for r in range(6, ledger.max_row + 1):
        if str(ledger.cell(row=r, column=2).value or "").startswith("合计"):
            break
        nm = ledger.cell(row=r, column=3).value
        if not isinstance(nm, str):
            continue
        t = stat.setdefault(nm, [0, 0, 0.0])
        t[0] += ledger.cell(row=r, column=4).value or 0
        t[1] += ledger.cell(row=r, column=5).value or 0
        t[2] += ledger.cell(row=r, column=6).value or 0
    ordered = sorted(info, key=lambda n: (-stat.get(n, [0, 0, 0.0])[0],
                                          -stat.get(n, [0, 0, 0.0])[2]))
    for i, nm in enumerate(ordered):
        r = 6 + i
        cat, ask, floor_ = info[nm]
        t = stat.get(nm, [0, 0, 0.0])
        sold, dmg = t[0], t[1]
        amt = round(t[2], 2)
        avg = round(amt / sold, 2) if sold else None
        vals = ((2, i + 1), (3, nm), (4, cat), (5, ask), (6, floor_),
                (7, cost_map.get(nm)),
                (8, f'=IF(OR(F{r}=0,G{r}=0,G{r}=""),"",ROUND(G{r}/F{r},3))'),
                (9, sold), (10, dmg),
                (11, f'=IF(I{r}+J{r}=0,"",ROUND(J{r}/(I{r}+J{r}),3))'),
                (12, amt if amt else None), (13, avg),
                (14, f'=IF(OR(I{r}=0,M{r}=0,G{r}=0,G{r}=""),"",ROUND(0.5-G{r}/M{r},3))'),
                (15, f'=IF(OR(I{r}=0,M{r}=0,G{r}=0,G{r}=""),"",ROUND(0.5*M{r}-G{r},2))'),
                (16, i + 1 if sold > 0 else None))
        for cc, v in vals:
            a.cell(row=r, column=cc, value=v)
    # 兼职工资累计 = 销售记录 G 列合计
    wage_row = next((r for r in range(6 + len(ordered), 60)
                     if a.cell(row=r, column=2).value == "兼职工资累计"), None)
    if wage_row:
        paid = sum(daily.cell(row=r, column=7).value or 0
                   for r in range(6, 60)
                   if daily.cell(row=r, column=2).value not in (None, "")
                   and not str(daily.cell(row=r, column=2).value).startswith(("合计", "注")))
        a.cell(row=wage_row, column=12, value=round(paid, 2))

    wb.save(inv_path)
    unpriced = sorted({nm for nm, _, _, _ in records} - set(info))
    return {"date": date_str, "staff": staff, "sold": sold_cnt, "dmg": dmg_cnt,
            "amount": amt_sum, "cost": cost_sum, "wage": wage,
            "stock": {nm: stock_names[nm] - q for nm, q in need.items()},
            "unpriced": unpriced}


# ---------------- 补货/库存调整 ----------------

def parse_stock_text(text):
    """补货文本 → [{name, qty}]；每行『商品 数量』，正数补货、负数调减。
    返回 (items, problems)"""
    items, problems = {}, []
    for raw in re.split(r"[\n；;，,、]+", text or ""):
        line = raw.strip()
        if not line:
            continue
        m = re.search(r"([+-]?\d+(?:\.\d+)?)\s*(?:个|件)?\s*$", line)
        if not m:
            problems.append(f"「{line}」没认出数量（格式：商品 数量，调减用负数）")
            continue
        name = re.sub(r"[+-]?\d+(?:\.\d+)?", "", line)
        name = re.sub(r"[个件\s\+]+", "", name)
        if not name:
            problems.append(f"「{line}」没认出商品名")
            continue
        qty = int(float(m.group(1)))
        t = items.setdefault(name, {"name": name, "qty": 0})
        t["qty"] += qty
    return list(items.values()), problems


def apply_stock_adjust(inv_path, date_str, entries, note="", kind="进货"):
    """补货/调整库存：entries=[{name,qty}]（正补负减）。
    更新当前库存数量 + 在「库存记录」留痕。失败抛 SyncError 不改文件。"""
    wb = load_workbook(inv_path)
    stock, hist = wb["当前库存"], wb["库存记录"]
    names, rows = _stock_map(stock)
    recs = []
    for it in entries:
        nm = it["name"]
        if nm not in names:
            raise SyncError(f"商品「{nm}」不在当前库存名单中")
        q = int(it.get("qty") or 0)
        if q != 0:
            recs.append((nm, q))
    if not recs:
        raise SyncError("没有可执行的调整（数量都是 0 或为空）")
    for nm, q in recs:
        new = names[nm] + q
        if new < 0:
            raise SyncError(f"{nm} 调整后为负（{names[nm]}{q:+d}），请核对数量")

    # 扣/加当前库存
    for nm, q in recs:
        stock.cell(row=rows[nm], column=5, value=names[nm] + q)

    # 库存记录留痕：插在注行之前（注行有合并，先拆后复位）
    last = max((r for r in range(6, hist.max_row + 1)
                if isinstance(hist.cell(row=r, column=3).value, str)), default=5)
    note_row = next((r for r in range(last + 1, hist.max_row + 1)
                     if str(hist.cell(row=r, column=2).value or "").startswith("注：")), None)
    insert_at = note_row if note_row else last + 1
    unmerged = []
    if note_row:
        for m in list(hist.merged_cells.ranges):
            if m.min_row >= note_row:
                unmerged.append((m.min_row, m.max_col))
                hist.unmerge_cells(str(m))
    hist.insert_rows(insert_at, len(recs))
    for i, (nm, q) in enumerate(recs):
        r = insert_at + i
        for cc in range(2, 9):
            hist.cell(row=r, column=cc)._style = copy(
                hist.cell(row=insert_at + len(recs), column=cc)._style
                if False else hist.cell(row=max(6, insert_at - 1) if insert_at > 6 else 6, column=cc)._style)
        hist.row_dimensions[r].height = 22
        rm = f"{kind}" + (f"：{note}" if note else "")
        for cc, v in zip(range(2, 9), (date_str, nm, q, None, None, "—", rm)):
            hist.cell(row=r, column=cc, value=v)
    for min_row, max_col in unmerged:
        hist.merge_cells(start_row=min_row + len(recs), start_column=2,
                         end_row=min_row + len(recs), end_column=max_col)

    wb.save(inv_path)
    return {"stock": {nm: names[nm] + q for nm, q in recs}}
