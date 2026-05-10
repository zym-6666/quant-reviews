#!/usr/bin/env python3
"""
短线复盘日报 v4 — 正确连板算法
Sheet1: 最高标梯队（近30天）
Sheet2: 今日复盘
Sheet3: 昨日涨停今日表现
Sheet4: 板块分布
"""
import sqlite3, sys, os
from collections import defaultdict
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DB = "/home/zym/.hermes/quant/quant.db"
OUT_DIR = "/home/zym/.hermes/cron/output"
os.makedirs(OUT_DIR, exist_ok=True)

thin = Side(style='thin', color="BDBDBD")
B = Border(left=thin, right=thin, top=thin, bottom=thin)
C = Alignment(horizontal="center", vertical="center", wrap_text=True)
L = Alignment(horizontal="left", vertical="center", wrap_text=True)

def wr(ws, r, c, v, bold=False, fg=None, fc="000000", sz=10, al=C):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = Font(bold=bold, color=fc, size=sz)
    cell.alignment = al
    if fg:
        cell.fill = PatternFill("solid", fgColor=fg)
    cell.border = B
    return cell

conn = sqlite3.connect(DB, timeout=60)
cur = conn.cursor()

TARGET = sys.argv[1] if len(sys.argv) > 1 else "20260508"

# ── 近30个交易日（ASC序）────────────────────────────
cur.execute("SELECT DISTINCT trade_date FROM price_daily ORDER BY trade_date ASC")
all_dates = [r[0] for r in cur.fetchall()]
recent = all_dates[-30:]
idx_start = all_dates.index(recent[0])
all_dates_asc = recent  # 已是ASC

# ── 一次性拉近30天涨停数据 ─────────────────────────
cur.execute(f"""
    SELECT trade_date, ts_code, pct_chg, close FROM price_daily
    WHERE trade_date IN ({','.join('?'*len(all_dates_asc))}) AND pct_chg>=9.9
    ORDER BY trade_date ASC
""", all_dates_asc)
rows = cur.fetchall()

zt = defaultdict(dict)  # {td: {code: (pct, close)}}
for td, code, pct, close in rows:
    zt[td][code] = (pct, close)

# ── 名称映射 ───────────────────────────────────────
all_codes = set(code for td_dict in zt.values() for code in td_dict)
if all_codes:
    ph = ','.join('?' * len(all_codes))
    cur.execute(f"SELECT code, name FROM stocks WHERE code IN ({ph})", list(all_codes))
    name_map = dict(cur.fetchall())
else:
    name_map = {}

# ── ASC序计算连板天数 ───────────────────────────────
dates_in_zt = [d for d in all_dates_asc if d in zt]
streak = {}  # (td, code) -> days

for i, td in enumerate(dates_in_zt):
    if i == 0:
        for code in zt[td]:
            streak[(td, code)] = 1
    else:
        prev_td = dates_in_zt[i-1]
        for code in zt[td]:
            if code in zt[prev_td]:
                streak[(td, code)] = streak.get((prev_td, code), 0) + 1
            else:
                streak[(td, code)] = 1

# ── 每日最高标 ─────────────────────────────────────
print("每日最高标:")
daily_top = {}
for td in all_dates_asc:
    codes = zt.get(td, {})
    if not codes:
        daily_top[td] = None
        continue
    best = None
    best_days = 0
    for code in codes:
        s = streak.get((td, code), 0)
        if s > best_days:
            best_days = s
            best = (code, name_map.get(code, code), codes[code][0], s)
    daily_top[td] = best
    tag = f"{best[1]}{best[3]}板" if best else "无"
    print(f"  {td}: {tag}")

# ── 今日数据 ──────────────────────────────────────
today = TARGET
today_zt = zt.get(today, {})
yd = cur.execute(
    "SELECT trade_date FROM price_daily WHERE trade_date<? ORDER BY trade_date DESC LIMIT 1", (today,)
).fetchone()
yd = yd[0] if yd else None
yd_zt = zt.get(yd, {}) if yd else {}

today_dt = cur.execute(
    "SELECT COUNT(*) FROM price_daily WHERE trade_date=? AND pct_chg<=-9.9", (today,)
).fetchone()[0]
today_pct = {r[0]: r[1] for r in cur.execute(
    "SELECT ts_code, pct_chg FROM price_daily WHERE trade_date=?", (today,)
).fetchall()}

lianban = set(today_zt.keys()) & set(yd_zt.keys())
max_lian = 0
for code in lianban:
    s = streak.get((today, code), 0)
    if s > max_lian:
        max_lian = s

print(f"\n今日 {today}: 涨停{len(today_zt)} 跌停{today_dt} 连板{len(lianban)} 最高{max_lian}板")

# ── Excel ──────────────────────────────────────────
wb = openpyxl.Workbook()

bc = {1:"E8F5E9",2:"FFF9C4",3:"FFE0B2",4:"FFCDD2",5:"F8BBD0",
      6:"E1BEE7",7:"CE93D8",8:"D1C4E9",9:"C5CAE9",10:"BBDEFB"}
fc_map = {1:"2E7D32",2:"F57F17",3:"E65100",4:"C62828",5:"AD1457",
          6:"6A1B9A",7:"4A148C",8:"4A148C",9:"4A148C",10:"4A148C"}

# ══ Sheet1: 最高标梯队 ══════════════════════════════
ws1 = wb.active
ws1.title = "最高标梯队"

hdrs = ["日期"] + [f"{i}板" for i in range(1, 11)]
for ci, h in enumerate(hdrs, 1):
    wr(ws1, 1, ci, h, bold=True, fg="1F4E79", fc="FFFFFF", sz=11)
ws1.row_dimensions[1].height = 22

all_dates_fmt = [f"{d[:4]}-{d[4:6]}-{d[6:]}" for d in all_dates_asc]
for ri, (d, ds) in enumerate(zip(all_dates_asc, all_dates_fmt), 2):
    wr(ws1, ri, 1, ds, bold=True, fg="D6E4F0")
    info = daily_top.get(d)
    if info:
        code, name, pct, days = info
        bg = bc.get(days, "EEEEEE")
        fc = fc_map.get(days, "000000")
        wr(ws1, ri, days+1, f"{name}({days}板)", bold=True, fg=bg, fc=fc, sz=9)
    for ci in range(2, 12):
        cell = ws1.cell(row=ri, column=ci)
        if cell.value is None:
            cell.fill = PatternFill("solid", fgColor="FAFAFA")
            cell.border = B
    ws1.row_dimensions[ri].height = 20

ws1.column_dimensions['A'].width = 13
for i in range(1, 11):
    ws1.column_dimensions[get_column_letter(i+1)].width = 17

# ══ Sheet2: 今日复盘 ═════════════════════════════════
ws2 = wb.create_sheet("今日复盘")
ws2.column_dimensions['A'].width = 22
ws2.column_dimensions['B'].width = 14
ws2.column_dimensions['C'].width = 35

tc = ws2.cell(row=1, column=1, value=f"今日复盘 {f'{today[:4]}-{today[4:6]}-{today[6:]}'}")
tc.font = Font(color="FFFFFF", bold=True, size=14)
tc.fill = PatternFill("solid", fgColor="1F4E79")
tc.alignment = C
ws2.merge_cells("A1:C1")
ws2.row_dimensions[1].height = 30

r = 3
ws2.cell(row=r, column=1, value="📊 情绪指标").font = Font(bold=True, color="FFFFFF", size=12)
ws2.cell(row=r, column=1).fill = PatternFill("solid", fgColor="1F4E79")
ws2.merge_cells("A3:C3"); r=4

mdata = [
    ("今日涨停", len(today_zt), "FFCCCC", "C62828"),
    ("今日跌停", today_dt, "CCFFCC", "2E7D32"),
    ("连板成功", len(lianban), "FFE0B2", "E65100"),
    ("最高标", f"{max_lian}板", "F8BBD0", "AD1457"),
]
for lb, val, bg, fc in mdata:
    wr(ws2, r, 1, lb, fg="D6E4F0")
    c2 = ws2.cell(row=r, column=2, value=val)
    c2.font = Font(bold=True, size=14, color=fc)
    c2.fill = PatternFill("solid", fgColor=bg)
    c2.border = B; c2.alignment = C
    ws2.cell(row=r, column=3).fill = PatternFill("solid", fgColor=bg)
    r += 1

r += 1
ws2.cell(row=r, column=1, value="📈 今日涨停股").font = Font(bold=True, color="FFFFFF", size=12)
ws2.cell(row=r, column=1).fill = PatternFill("solid", fgColor="1F4E79")
ws2.merge_cells(f"A{r}:C{r}"); r+=1

# 按streak天数排序
sorted_zt = sorted(today_zt.items(), key=lambda x: streak.get((today, x[0]), 0), reverse=True)
for code, (pct, close) in sorted_zt[:80]:
    name = name_map.get(code, code)
    days = streak.get((today, code), 0)
    is_lb = code in lianban
    bg = "FFCCCC" if is_lb else "FFF9C4"
    fc_days = fc_map.get(days, "666666") if days >= 2 else "000000"
    wr(ws2, r, 1, name, bold=is_lb, fg=bg)
    wr(ws2, r, 2, code, fg=bg)
    day_str = f"{days}板" if days >= 2 else ""
    wr(ws2, r, 3, f"{pct:.2f}% {day_str}", bold=True, fg=bg, fc="C62828")
    r += 1

# ══ Sheet3: 昨日涨停今日表现 ══════════════════════════
ws3 = wb.create_sheet("昨日涨停今日表现")
hdrs3 = ["股票名称","代码","昨日涨幅","今日涨幅","今收","状态"]
for ci, h in enumerate(hdrs3, 1):
    wr(ws3, 1, ci, h, bold=True, fg="1F4E79", fc="FFFFFF", sz=11)

yd_zt_sorted = sorted(yd_zt.items(), key=lambda x: today_pct.get(x[0], 0), reverse=True)
for ri, (code, (yd_pct, yd_close)) in enumerate(yd_zt_sorted, 2):
    name = name_map.get(code, code)
    td_pct = today_pct.get(code, 0)
    td_close = today_zt.get(code, (None,None))[0] if code in today_zt else None

    if code in today_zt:
        st, bg = "涨停✓", "FFCCCC"
    elif td_pct > 0:
        st, bg = f"+{td_pct:.1f}%", "FFE0B2"
    elif td_pct < 0:
        st, bg = f"{td_pct:.1f}%", "CCFFCC"
    else:
        st, bg = "平", "F5F5F5"

    wr(ws3, ri, 1, name, fg=bg)
    wr(ws3, ri, 2, code, fg=bg)
    wr(ws3, ri, 3, f"{yd_pct:.2f}%", fg=bg, fc="C62828")
    wr(ws3, ri, 4, f"{td_pct:+.2f}%", bold=True, fg=bg,
       fc="C62828" if td_pct>0 else ("2E7D32" if td_pct<0 else "000000"))
    wr(ws3, ri, 5, f"{td_close:.2f}" if td_close else "-", fg=bg)
    wr(ws3, ri, 6, st, bold=True, fg=bg)

for ci, w in enumerate([14,13,12,12,10,12], 1):
    ws3.column_dimensions[get_column_letter(ci)].width = w

# ══ Sheet4: 板块分布 ═════════════════════════════════
ws4 = wb.create_sheet("板块分布")
hdrs4 = ["概念板块","涨停家数","占比","代表股票"]
for ci, h in enumerate(hdrs4, 1):
    wr(ws4, 1, ci, h, bold=True, fg="1F4E79", fc="FFFFFF", sz=11)

codes_today_list = list(today_zt.keys())
if codes_today_list:
    ph = ','.join('?' * len(codes_today_list))
    cur.execute(f"SELECT code, industry FROM stocks WHERE code IN ({ph})", codes_today_list)
    ind_map = dict(cur.fetchall())
else:
    ind_map = {}

industry_zt = {}
for code, (pct, close) in today_zt.items():
    ind = (ind_map.get(code) or "其他")[:8]
    name = name_map.get(code, code)
    industry_zt.setdefault(ind, []).append(name)

ind_sorted = sorted(industry_zt.items(), key=lambda x: len(x[1]), reverse=True)
total = len(today_zt) or 1
for ri, (ind, stocks) in enumerate(ind_sorted, 2):
    cnt = len(stocks)
    pct = cnt/total*100
    bg = "FFE0B2" if cnt>=5 else ("FFF9C4" if cnt>=3 else "FAFAFA")
    wr(ws4, ri, 1, ind, bold=(cnt>=3), fg=bg)
    wr(ws4, ri, 2, cnt, fg=bg)
    wr(ws4, ri, 3, f"{pct:.1f}%", fg=bg)
    wr(ws4, ri, 4, "、".join(stocks[:5]), al=L, fg=bg)

for ci, w in enumerate([20,10,8,55], 1):
    ws4.column_dimensions[get_column_letter(ci)].width = w

conn.close()

out = f"{OUT_DIR}/短线复盘_{today}.xlsx"
wb.save(out)
print(f"\n✅ {out} ({os.path.getsize(out)/1024:.0f}KB)")
