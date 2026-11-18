# -*- coding: utf-8 -*-
"""3D打印摊位记账 · 手机网页版
数据存放在 GitHub 私有仓库的两份 Excel 中；录入即自动完成全部联动更新。
"""
import io, re, tempfile, os
from datetime import date, datetime
import streamlit as st
import pandas as pd
import requests
from openpyxl import load_workbook
from sync_core import parse_sales_text, resolve, apply_sync, SyncError, parse_stock_text, apply_stock_adjust, undo_last_sync, undo_last_adjust

INV = "销售、库存、经营分析.xlsx"
PRICE = "兼职销售登记表、成本利润登记表.xlsx"
API = "https://api.github.com/repos"

st.set_page_config(page_title="3D打印摊位记账", page_icon="🎪",
                   layout="wide", initial_sidebar_state="expanded")

# ---------------- GitHub 读写 ----------------
def gh_cfg():
    tok = st.secrets.get("GH_TOKEN", "")
    repo = st.secrets.get("REPO", "")
    if not tok or not repo:
        st.error("未配置 GitHub 密钥（GH_TOKEN / REPO），按部署指南在 Streamlit Secrets 里填好。")
        st.stop()
    return repo, {"Authorization": f"Bearer {tok}",
                  "Accept": "application/vnd.github+json"}

@st.cache_data(ttl=60, show_spinner=False)
def gh_get(path):
    repo, hdr = gh_cfg()
    r = requests.get(f"{API}/{repo}/contents/{path}", headers=hdr, timeout=30)
    if r.status_code != 200:
        raise RuntimeError(f"读取 {path} 失败：{r.json().get('message')}")
    import base64
    return base64.b64decode(r.json()["content"]), r.json()["sha"]

def gh_put(path, content_bytes, sha, msg):
    repo, hdr = gh_cfg()
    import base64
    r = requests.put(f"{API}/{repo}/contents/{path}", headers=hdr, timeout=30,
                     json={"message": msg, "content": base64.b64encode(content_bytes).decode(), "sha": sha})
    if r.status_code not in (200, 201):
        raise RuntimeError(f"写回 {path} 失败：{r.json().get('message')}")

@st.cache_data(ttl=60, show_spinner=False)
def load_wb_bytes(path):
    b, sha = gh_get(path)
    return b

@st.cache_data(ttl=60, show_spinner=False)
def product_names():
    """销售录入表里的定价款（标准名列表）"""
    b = load_wb_bytes(INV)
    wb = load_workbook(io.BytesIO(b), read_only=False)
    e = wb["销售录入"]
    names = [e.cell(row=r, column=2).value for r in range(6, 80)
             if isinstance(e.cell(row=r, column=2).value, str)
             and isinstance(e.cell(row=r, column=3).value, (int, float))]
    wb.close()
    return names

@st.cache_data(ttl=60, show_spinner=False)
def all_stock_names():
    """当前库存全部款名（含未定价款，补货下拉用）"""
    b = load_wb_bytes(INV)
    wb = load_workbook(io.BytesIO(b), read_only=False)
    s = wb["当前库存"]
    names = [s.cell(row=r, column=3).value for r in range(5, 100)
             if isinstance(s.cell(row=r, column=3).value, str)]
    wb.close()
    return names

@st.cache_data(ttl=60, show_spinner=False)
def recorded_dates():
    b = load_wb_bytes(INV)
    wb = load_workbook(io.BytesIO(b), read_only=False)
    d = wb["销售记录"]
    out = [d.cell(row=r, column=2).value for r in range(6, 40)
           if isinstance(d.cell(row=r, column=2).value, str)
           and not d.cell(row=r, column=2).value.startswith(("合计", "注"))]
    wb.close()
    return out

# ---------------- 页面 ----------------
st.sidebar.title("🎪 3D打印摊位记账")
page = st.sidebar.radio("功能", ["📝 今日录入", "📥 补货/调整", "📦 当前库存", "🧾 销售记录", "📊 产品分析", "⬇️ 下载Excel"])
st.sidebar.caption("数据自动同步保存\n手机可添加到主屏幕当 App 用")

if page == "📝 今日录入":
    st.title("📝 今日录入")
    st.caption("填写日期、兼职、销售内容（每行一款：`商品 数量 金额`，只写金额默认 1 件；损坏写 `损:商品 数量`）。简称自动识别（如：亚索剑、雪糕、小蛇）。")

    c1, c2, c3 = st.columns(3)
    with c1:
        d = st.date_input("日期", value=date.today())
    with c2:
        # 常用兼职名单（默认小孔；要加人往这里添即可）
        _staff_sel = st.selectbox("值班兼职", ["小孔", "亚特", "其他（手填）"], index=0)
        staff = _staff_sel if _staff_sel != "其他（手填）" else st.text_input("兼职姓名", placeholder="输入名字")
    with c3:
        wage = st.number_input("实付工资（元，没收摊可留空）", min_value=0.0, step=5.0, value=None)
    text = st.text_area("销售内容", height=180)
    date_str = f"{d.month}月{d.day}日"

    if st.button("🔍 解析预览", use_container_width=True, type="primary"):
        names = product_names()
        parsed, problems = parse_sales_text(text)
        rows, probs2 = [], []
        for it in parsed:
            std = resolve(it["name"], names)
            if std is None:
                probs2.append(f"「{it['name']}」没匹配到商品，请在下表手动选择")
                rows.append({"商品": it["name"], "数量": it["qty"], "金额": it["amt"] or 0,
                             "损坏": it["dmg"]})
            else:
                if std != it["name"]:
                    st.info(f"名称映射：{it['name']} → {std}")
                rows.append({"商品": std, "数量": it["qty"], "金额": it["amt"] or 0,
                             "损坏": it["dmg"]})
        for p in problems + probs2:
            st.warning(p)
        st.session_state["rows"] = rows

    if "rows" in st.session_state and st.session_state["rows"]:
        st.subheader("核对无误后提交（可下拉改商品名）")
        df = pd.DataFrame(st.session_state["rows"])
        edited = st.data_editor(df, use_container_width=True, hide_index=True,
                                column_config={"商品": st.column_config.SelectboxColumn(
                                    "商品", options=product_names(), required=True)})
        if date_str in recorded_dates():
            st.warning(f"⚠️ {date_str} 已有录入记录，本次将追加记账，请确认没有重复。")
        if edited is not None and not edited.empty and st.button("✅ 确认同步全部更新", type="primary",
                                                                 use_container_width=True):
            inv_bytes, inv_sha = gh_get(INV)
            price_bytes, _ = gh_get(PRICE)
            with tempfile.TemporaryDirectory() as td:
                ip, pp = os.path.join(td, "inv.xlsx"), os.path.join(td, "price.xlsx")
                open(ip, "wb").write(inv_bytes)
                open(pp, "wb").write(price_bytes)
                items = [{"name": r["商品"], "qty": int(r["数量"]),
                          "amt": float(r["金额"]) or None, "dmg": bool(r["损坏"])}
                         for _, r in edited.iterrows()]
                try:
                    rep = apply_sync(ip, pp, date_str, staff, items,
                                     wage=float(wage) if wage else None)
                except SyncError as ex:
                    st.error(f"同步失败（未做任何修改）：{ex}")
                    st.stop()
                new_bytes = open(ip, "rb").read()
            try:
                gh_put(INV, new_bytes, inv_sha, f"录入 {date_str} {rep['sold']}件/{rep['amount']}元")
            except Exception as ex:
                st.error(f"⚠️ 本地已更新但写回云端失败：{ex}\n请截图此页发给摊主处理，不要重复提交。")
                st.stop()
            st.success(f"✅ 同步完成：{date_str} {rep['staff']} · 售出 {rep['sold']} 件 "
                       f"损坏 {rep['dmg']} 件 · 营业额 {rep['amount']:.0f} 元 · 成本 {rep['cost']:.1f} 元"
                       + (f" · 工资 {rep['wage']:.0f} 元" if rep['wage'] else ""))
            st.cache_data.clear()
            low = {n: q for n, q in rep["stock"].items() if q <= 5}
            if low:
                st.markdown("**扣减后库存：** " + "、".join(f"{n} 剩{q}" for n, q in sorted(low.items(), key=lambda x: x[1])) + " ⚠️ 建议补货")
            if rep["unpriced"]:
                st.caption("未定价不计入统计：" + "、".join(rep["unpriced"]))
            st.session_state.pop("rows")

    with st.expander("↩️ 撤销最近一次录入（误录补救）"):
        if st.button("查看最近一次录入"):
            b3 = load_wb_bytes(INV)
            wb3 = load_workbook(io.BytesIO(b3), read_only=False)
            d3 = wb3["销售记录"]
            r_t = next(r for r in range(6, d3.max_row + 1)
                       if str(d3.cell(row=r, column=2).value or "").startswith("合计"))
            last_r = max(r for r in range(6, r_t) if d3.cell(row=r, column=2).value)
            dt = d3.cell(row=last_r, column=2).value
            l3 = wb3["销售明细"]
            rows3 = [(l3.cell(row=r, column=3).value, l3.cell(row=r, column=4).value,
                      l3.cell(row=r, column=5).value) for r in range(6, 60)
                     if l3.cell(row=r, column=2).value == dt]
            wb3.close()
            st.session_state["undo_prev"] = (dt, rows3, d3.cell(row=last_r, column=6).value,
                                             d3.cell(row=last_r, column=7).value)
        if "undo_prev" in st.session_state:
            dt, rows3, amt3, wage3 = st.session_state["undo_prev"]
            st.warning(f"将撤销 **{dt}** 的全部录入：营业额 {amt3 or 0} 元、工资 {wage3 or 0} 元，"
                       f"共 {sum(q or 0 for _, q, _ in rows3)} 件（库存会自动加回，统计自动重算）。"
                       "只能撤销最近一天，更早的找摊主处理。")
            if st.button(f"✅ 确认撤销 {dt}", type="primary"):
                inv_b3, inv_sha3 = gh_get(INV)
                price_b3, _ = gh_get(PRICE)
                with tempfile.TemporaryDirectory() as td3:
                    ip3, pp3 = os.path.join(td3, "i3.xlsx"), os.path.join(td3, "p3.xlsx")
                    open(ip3, "wb").write(inv_b3)
                    open(pp3, "wb").write(price_b3)
                    try:
                        rep3 = undo_last_sync(ip3, pp3)
                    except SyncError as ex:
                        st.error(f"撤销失败：{ex}")
                        st.stop()
                    nb3 = open(ip3, "rb").read()
                try:
                    gh_put(INV, nb3, inv_sha3, f"撤销 {rep3['date']} 录入")
                except Exception as ex:
                    st.error(f"本地已撤销但写回云端失败：{ex}，请截图给摊主，不要重复操作。")
                    st.stop()
                st.success(f"✅ 已撤销 {rep3['date']}：" + "、".join(
                    f"{n}+{q}" for n, q in rep3["restored"].items()))
                st.cache_data.clear()
                st.session_state.pop("undo_prev")

elif page == "📥 补货/调整":
    st.title("📥 补货 / 库存调整")
    st.caption("每行一款：`商品 数量`——正数=补货入库，负数=盘点调减（例：`迷你蛇 20`、`龙蛋 -2`）。简称自动识别。每笔自动记入「库存记录」留痕，不影响销售统计。")
    c1, c2 = st.columns(2)
    with c1:
        d2 = st.date_input("日期", value=date.today(), key="adj_date")
    with c2:
        kind = st.selectbox("类型", ["进货", "盘点调整"])
    note = st.text_input("备注（可选）", placeholder="如：第二批到货")
    text2 = st.text_area("补货/调整内容", height=160, key="adj_text")

    if st.button("🔍 解析预览", use_container_width=True, type="primary"):
        names_all = all_stock_names()
        parsed, problems = parse_stock_text(text2)
        rows2, probs2 = [], []
        for it in parsed:
            std = resolve(it["name"], names_all)
            if std is None:
                probs2.append(f"「{it['name']}」没匹配到商品，请在下表手动选择")
                rows2.append({"商品": it["name"], "数量": it["qty"]})
            else:
                if std != it["name"]:
                    st.info(f"名称映射：{it['name']} → {std}")
                rows2.append({"商品": std, "数量": it["qty"]})
        for pmsg in problems + probs2:
            st.warning(pmsg)
        st.session_state["adj_rows"] = rows2

    if "adj_rows" in st.session_state and st.session_state["adj_rows"]:
        st.subheader("核对无误后提交（可下拉改商品名，数量正=补货 负=调减）")
        df2 = pd.DataFrame(st.session_state["adj_rows"])
        edited2 = st.data_editor(df2, use_container_width=True, hide_index=True,
                                 column_config={"商品": st.column_config.SelectboxColumn(
                                     "商品", options=all_stock_names(), required=True)})
        if edited2 is not None and not edited2.empty and st.button(
                "✅ 确认更新库存", type="primary", use_container_width=True):
            inv_bytes2, inv_sha2 = gh_get(INV)
            with tempfile.TemporaryDirectory() as td:
                ip2 = os.path.join(td, "inv2.xlsx")
                open(ip2, "wb").write(inv_bytes2)
                items2 = [{"name": r["商品"], "qty": int(r["数量"])}
                          for _, r in edited2.iterrows()]
                try:
                    rep2 = apply_stock_adjust(ip2, f"{d2.month}月{d2.day}日", items2,
                                              note=note, kind=kind)
                except SyncError as ex:
                    st.error(f"失败（未做任何修改）：{ex}")
                    st.stop()
                new_bytes2 = open(ip2, "rb").read()
            try:
                gh_put(INV, new_bytes2, inv_sha2, f"{kind} {d2} " + "、".join(
                    f"{n}{q:+d}" for n, q in rep2["stock"].items()))
            except Exception as ex:
                st.error(f"⚠️ 本地已更新但写回云端失败：{ex}\n请截图发给摊主，不要重复提交。")
                st.stop()
            st.success(f"✅ 已更新：" + "、".join(f"{n} → {q}" for n, q in rep2["stock"].items()))
            st.cache_data.clear()
            low = {n: q for n, q in rep2["stock"].items() if q <= 5}
            if low:
                st.markdown("⚠️ 仍低于推荐值的：" + "、".join(f"{n}({q})" for n, q in sorted(low.items(), key=lambda x: x[1])))
            st.session_state.pop("adj_rows")

elif page == "📦 当前库存":
    st.title("📦 当前库存")
    b = load_wb_bytes(INV)
    wb = load_workbook(io.BytesIO(b), read_only=False)
    s = wb["当前库存"]
    rows = []
    for r in range(5, 80):
        if isinstance(s.cell(row=r, column=2).value, int):
            nm = s.cell(row=r, column=3).value
            if isinstance(nm, str):
                rec = s.cell(row=r, column=4).value
                qty = s.cell(row=r, column=5).value or 0
                rows.append({"商品名称": nm, "推荐库存": rec, "库存（件）": qty,
                             "备注": s.cell(row=r, column=6).value or ""})
    wb.close()
    df = pd.DataFrame(rows)
    # 推荐库存/库存列转可空整数，去掉 .0 小数显示
    for _col in ("推荐库存", "库存（件）"):
        df[_col] = pd.to_numeric(df[_col], errors="coerce").astype("Int64")
    def hl(row):
        rec, qty = row["推荐库存"], row["库存（件）"]
        return ["background-color:#FFC7CE;color:#9C0006" if isinstance(rec, (int, float))
                and qty < rec else "" for _ in row]
    st.dataframe(df.style.apply(hl, axis=1), use_container_width=True, hide_index=True, height=560)
    st.caption("红色 = 库存低于推荐值，建议补货")

    # ---- 快捷加减库存：在表里填「调整±」列（+补货 / -调减），确认后生效 ----
    st.subheader("⚡ 快捷调整库存")
    st.caption("在下方表格的「调整±」列填数字（正=补货入库，负=调减），点「检查调整」→ 确认后生效。每笔自动记入库存记录。")
    df_adj = df.copy()
    df_adj["调整±"] = 0
    edited3 = st.data_editor(df_adj, use_container_width=True, hide_index=True, key="stock_adj_editor",
                             column_config={"商品名称": st.column_config.TextColumn(disabled=True),
                                            "推荐库存": st.column_config.NumberColumn(disabled=True),
                                            "库存（件）": st.column_config.NumberColumn(disabled=True),
                                            "备注": st.column_config.TextColumn(disabled=True)})
    if st.button("🔍 检查调整", use_container_width=True):
        changes = edited3[edited3["调整±"].fillna(0) != 0]
        if changes.empty:
            st.info("没有填任何调整量（「调整±」列全是 0）")
        else:
            st.session_state["stock_changes"] = [
                {"name": r["商品名称"], "qty": int(r["调整±"])} for _, r in changes.iterrows()]
    if "stock_changes" in st.session_state and st.session_state["stock_changes"]:
        chs = st.session_state["stock_changes"]
        st.warning("将调整：" + "、".join(f"{c['name']} {c['qty']:+d}" for c in chs)
                   + "（正=补货 负=调减，自动记入库存记录）")
        if st.button("✅ 确认调整", type="primary", use_container_width=True):
            inv_b4, inv_sha4 = gh_get(INV)
            with tempfile.TemporaryDirectory() as td4:
                ip4 = os.path.join(td4, "i4.xlsx")
                open(ip4, "wb").write(inv_b4)
                try:
                    rep4 = apply_stock_adjust(ip4, f"{date.today().month}月{date.today().day}日",
                                              chs, note="库存页快捷调整", kind="盘点调整")
                except SyncError as ex:
                    st.error(f"失败（未做修改）：{ex}")
                    st.stop()
                nb4 = open(ip4, "rb").read()
            try:
                gh_put(INV, nb4, inv_sha4, "库存页快捷调整 " + "、".join(
                    f"{n}{q:+d}" for n, q in rep4["stock"].items()))
            except Exception as ex:
                st.error(f"本地已更新但写回失败：{ex}，请截图给摊主，不要重复提交。")
                st.stop()
            st.success("✅ 已调整：" + "、".join(f"{n} → {q}" for n, q in rep4["stock"].items()))
            st.cache_data.clear()
            st.session_state.pop("stock_changes", None)
            st.rerun()

    with st.expander("↩️ 撤销最近一次库存调整"):
        if st.button("查看最近一次调整"):
            b5 = load_wb_bytes(INV)
            wb5 = load_workbook(io.BytesIO(b5), read_only=False)
            h5 = wb5["库存记录"]
            rows5 = [r for r in range(6, h5.max_row + 1)
                     if isinstance(h5.cell(row=r, column=3).value, str)]
            if rows5:
                last5 = rows5[-1]
                d5 = h5.cell(row=last5, column=2).value
                batch5 = [(h5.cell(row=r, column=3).value, h5.cell(row=r, column=4).value,
                           h5.cell(row=r, column=8).value) for r in rows5
                          if h5.cell(row=r, column=2).value == d5]
                wb5.close()
                st.session_state["adj_undo_prev"] = (d5, batch5)
            else:
                wb5.close()
                st.info("库存记录为空")
        if "adj_undo_prev" in st.session_state:
            d5, batch5 = st.session_state["adj_undo_prev"]
            rm_ok = all(str(rm or "").startswith(("进货", "盘点调整")) for _, _, rm in batch5)
            if not rm_ok:
                st.info(f"最近一批（{d5}）不是进货/盘点调整（可能是历史记录），不能自动撤销。")
            else:
                st.warning(f"将撤销 **{d5}** 的调整：" + "、".join(
                    f"{n} {q:+d}" for n, q, _ in batch5) + "（数量反向加回）")
                if st.button(f"✅ 确认撤销 {d5} 的调整", type="primary"):
                    inv_b6, inv_sha6 = gh_get(INV)
                    with tempfile.TemporaryDirectory() as td6:
                        ip6 = os.path.join(td6, "i6.xlsx")
                        open(ip6, "wb").write(inv_b6)
                        try:
                            rep6 = undo_last_adjust(ip6)
                        except SyncError as ex:
                            st.error(f"撤销失败：{ex}")
                            st.stop()
                        nb6 = open(ip6, "rb").read()
                    try:
                        gh_put(INV, nb6, inv_sha6, f"撤销库存调整 {rep6['date']}")
                    except Exception as ex:
                        st.error(f"写回失败：{ex}，请截图给摊主。")
                        st.stop()
                    st.success("✅ 已撤销：" + "、".join(f"{n}{q:+d}" for n, q in rep6["reverted"].items()))
                    st.cache_data.clear()
                    st.session_state.pop("adj_undo_prev", None)
                    st.rerun()

elif page == "🧾 销售记录":
    st.title("🧾 销售记录")
    b = load_wb_bytes(INV)
    wb = load_workbook(io.BytesIO(b), read_only=False)
    d = wb["销售记录"]
    # 方案工资参数（表右侧 O5..O10：保底/一档上限/一档比例/二档上限/二档比例/三档比例）
    _raw = [d.cell(row=5 + i, column=15).value for i in range(6)]
    params = [v if v is not None else d0 for v, d0 in
              zip(_raw, [120, 300, 0.2, 600, 0.25, 0.2])]  # 缺项用默认，避免报错
    import math
    def _scheme(rev):
        if rev in (None, ""):
            return None
        p0, u1, r1, u2, r2, r3 = params
        return math.ceil(p0 + min(rev, u1) * r1 + max(0, min(rev, u2) - u1) * r2
                         + max(0, rev - u2) * r3)
    rows = []
    for r in range(6, 40):
        date_v = d.cell(row=r, column=2).value
        if not date_v or str(date_v).startswith(("合计", "注")):
            continue
        sold = d.cell(row=r, column=4).value or 0
        dmg = d.cell(row=r, column=5).value or 0
        rev = d.cell(row=r, column=6).value
        paid = d.cell(row=r, column=7).value
        cost = round(rev * 0.2, 2) if rev is not None else None   # 货品成本=营业额×20%（摊主口径）
        sch = _scheme(rev)
        diff = round((paid or 0) - sch, 2) if (paid is not None and sch is not None) else None
        profit = round(rev - (paid or 0) - (cost or 0), 2) if (rev is not None and paid is not None) else None
        theo = round(profit - 80, 2) if profit is not None else None   # 理论利润=再扣80摊位费
        rows.append([date_v, d.cell(row=r, column=3).value or "—", sold, dmg, rev, paid,
                     diff, cost, profit, theo, d.cell(row=r, column=12).value or ""])
    wb.close()
    df = pd.DataFrame(rows, columns=["日期", "兼职", "卖出件数", "损坏件数", "营业额", "实付工资",
                                     "差额", "货品成本", "利润", "理论利润", "备注"])
    if not df.empty:
        _r2 = lambda s: round(float(s.fillna(0).sum()), 2)
        total = {"日期": "合计", "兼职": "", "卖出件数": df["卖出件数"].sum(),
                 "损坏件数": df["损坏件数"].sum(), "营业额": _r2(df["营业额"]),
                 "实付工资": _r2(df["实付工资"]),
                 "差额": _r2(df["差额"]), "货品成本": _r2(df["货品成本"]),
                 "利润": _r2(df["利润"]), "理论利润": _r2(df["理论利润"]), "备注": ""}
        df = pd.concat([df, pd.DataFrame([total])], ignore_index=True)
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.caption("货品成本=营业额×20% · 利润=营业额−实付工资−货品成本 · 理论利润=再扣80元/天摊位费")

elif page == "📊 产品分析":
    st.title("📊 产品成本率详情（按销量排序）")
    b = load_wb_bytes(INV)
    wb = load_workbook(io.BytesIO(b), read_only=False)
    a = wb["产品成本率详情"]
    hdr = [a.cell(row=5, column=c).value for c in range(2, 17)]
    rows = []
    for r in range(6, 46):
        nm = a.cell(row=r, column=3).value
        if not nm:
            continue
        floor_p = a.cell(row=r, column=6).value   # 底价
        cost = a.cell(row=r, column=7).value      # 成本价
        sold = a.cell(row=r, column=9).value or 0
        dmg = a.cell(row=r, column=10).value or 0
        avg = a.cell(row=r, column=13).value      # 均价
        rate_cost = round(cost / floor_p, 3) if (cost and floor_p) else None
        rate_dmg = round(dmg / (sold + dmg), 3) if (sold + dmg) else None
        margin = round(0.5 - cost / avg, 3) if (cost and avg) else None
        unit_profit = round(avg * 0.5 - cost, 2) if (cost and avg) else None
        rows.append([a.cell(row=r, column=2).value, nm, a.cell(row=r, column=4).value,
                     a.cell(row=r, column=5).value, floor_p, cost, rate_cost,
                     sold, dmg, rate_dmg, a.cell(row=r, column=12).value, avg,
                     margin, unit_profit, a.cell(row=r, column=16).value])
    wb.close()
    st.dataframe(pd.DataFrame(rows, columns=hdr), use_container_width=True, hide_index=True)
    st.caption("单件利润=均价×50%−成本（已扣35%提成+15%摊位费）· 未录成本的款利润为空")

elif page == "⬇️ 下载Excel":
    st.title("⬇️ 下载 Excel")
    c1, c2 = st.columns(2)
    with c1:
        st.download_button("销售、库存、经营分析.xlsx", load_wb_bytes(INV),
                           file_name="销售、库存、经营分析.xlsx")
    with c2:
        st.download_button("兼职销售登记表、成本利润登记表.xlsx", load_wb_bytes(PRICE),
                           file_name="兼职销售登记表、成本利润登记表.xlsx")
    st.caption("下载的是当前云端最新版；在 WPS 里改过本地旧文件的话以云端为准，别再改本地版。")
