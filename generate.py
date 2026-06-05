# -*- coding: utf-8 -*-
"""
健康台灣深耕計畫 — 課程表產生器
================================
讀取課程時段 Excel，依「上課地點 → 行政區（北投／士林／中山）」與「處方類型」
分類、統計，並把資料填入既有的前端版型（日系大字版 / 手機版），
同時輸出一份結構化資料庫 courses.json 與入口頁 index.html。

原則：
  * 狀態為「取消」的時段一律排除，不計入統計、不顯示。
  * 一筆 Excel 列 = 一堂課（忠實呈現，不做人工合併）。
  * 版面 CSS/JS 完全沿用 templates/ 內的前端設計，本程式只抽換「資料區塊」。

重跑方式：  python generate.py
"""
import os
import json
import html
import urllib.request
import datetime as dt
from pathlib import Path
from collections import defaultdict, Counter

import openpyxl

# ---------------------------------------------------------------- 路徑設定
ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
TPL = ROOT / "templates"


def _latest_xlsx():
    """自動挑 data/ 內最新（修改時間最晚）的 .xlsx；找不到就報錯。"""
    files = [p for p in DATA.glob("*.xlsx") if not p.name.startswith("~$")]
    if not files:
        raise SystemExit(f"❌ 在 {DATA} 找不到任何 .xlsx，請先把課表 Excel 丟進去。")
    return max(files, key=lambda p: p.stat().st_mtime)


def resolve_xlsx():
    """資料來源優先序：
       1. 環境變數 COURSE_XLSX_URL（例如 Google Sheets 的 xlsx 匯出網址）
          → 下載到 data/_from_sheet.xlsx 使用（雲端自動更新走這條）。
       2. 否則用 data/ 內最新的本機 .xlsx。
    """
    url = os.environ.get("COURSE_XLSX_URL", "").strip()
    if url:
        DATA.mkdir(exist_ok=True)
        dest = DATA / "_from_sheet.xlsx"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req) as r, open(dest, "wb") as f:
            f.write(r.read())
        print(f"已自共用試算表下載資料 → {dest.name}")
        return dest
    return _latest_xlsx()


XLSX = resolve_xlsx()
OUT_BIG = ROOT / "大字版"
OUT_MOBILE = ROOT / "手機版"

# ---------------------------------------------------------------- 分類設定
# 處方類型 → (data-key, 標籤, 邊框色, 背景色)
TYPE_MAP = {
    "營養處方":   ("n", "營養", "#445230", "#EAF0E0"),
    "運動處方":   ("s", "運動", "#72491F", "#F4EADD"),
    "社會處方":   ("o", "社會", "#385268", "#E4EDF4"),
    "情緒調適處方": ("e", "情緒", "#6E4858", "#F4E8EE"),
}

WD_ZH = ["日", "一", "二", "三", "四", "五", "六"]
WD_EN = ["SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT"]

DISTRICTS = ["北投", "士林", "中山"]

# 全區綜合輸出檔名
ALL_BIG = "全區綜合_2026年6月課程表_日系大字版.html"
ALL_MOB = "全區綜合_2026年6月課程表_手機版.html"


def classify_district(loc: str):
    """依上課地點文字判斷行政區；回傳 北投/士林/中山 或 None。"""
    s = str(loc).replace(" ", "")
    # 北投關鍵字
    if any(k in s for k in ("北投", "石牌", "榮陽", "蔡秉勳", "翰譽耳鼻喉",
                            "永安", "瑜伽之光", "臻心", "ULifeFitnes",
                            "ULifeFitness")):
        return "北投"
    # 士林關鍵字
    if any(k in s for k in ("士林", "天母", "ZenYoga", "小船254", "後街21巷",
                            "福港街", "福華路")):
        return "士林"
    # 中山關鍵字
    if any(k in s for k in ("中山", "圓山", "夢想館", "晨昕")):
        return "中山"
    return None


# ---------------------------------------------------------------- 讀取資料
def load_courses():
    wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    # 欄位：課程名稱 處方類型 日期 時段 上課地點 已報名 人數上限 是否額滿 狀態
    courses = []
    skipped_cancel = 0
    unknown_loc = Counter()
    for r in rows[1:]:
        if not r or r[0] is None:
            continue
        name, ptype, date, slot, loc, enrolled, cap, full, status = r[:9]
        if str(status).strip() == "取消":      # 取消不放
            skipped_cancel += 1
            continue
        district = classify_district(loc)
        if district is None:
            unknown_loc[str(loc)] += 1
            continue
        d = _parse_date(date)
        if d is None:
            continue
        start, end = _parse_slot(slot)
        courses.append({
            "name": str(name).strip(),
            "type": str(ptype).strip(),
            "date": d.isoformat(),
            "start": start,
            "end": end,
            "venue": str(loc).strip(),
            "district": district,
            "enrolled": _int(enrolled),
            "capacity": _int(cap),
            "full": str(full).strip() == "是",
            "status": str(status).strip(),
        })
    return courses, header, skipped_cancel, unknown_loc


def _parse_date(v):
    """容錯解析日期：支援 datetime、'YYYY-MM-DD'、'YYYY/M/D'。"""
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = str(v).strip()[:10]
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def _parse_slot(slot):
    s = str(slot).replace("－", "-").replace("–", "-")
    parts = [p.strip() for p in s.split("-")]
    if len(parts) == 2:
        return parts[0], parts[1]
    return s.strip(), ""


def _int(v):
    try:
        return int(v)
    except (TypeError, ValueError):
        return None


# ---------------------------------------------------------------- 小工具
def esc(t):
    return html.escape(str(t), quote=False)


def tm_str(c):
    return f'{c["start"]}–{c["end"]}' if c["end"] else c["start"]


def sort_key(c):
    return (c["date"], c["start"], c["end"], c["name"])


def sunday_of(d: dt.date):
    dow = (d.weekday() + 1) % 7      # 週日=0 … 週六=6
    return d - dt.timedelta(days=dow)


# ---------------------------------------------------------------- 區塊抽換
def splice(template: str, head_marker: str, tail_marker: str, body: str):
    i = template.index(head_marker)
    j = template.index(tail_marker)
    return template[:i] + body + template[j:]


# ---------------------------------------------------------------- 大字版（週表格）
def render_big(courses, show_district=False):
    by_week = defaultdict(list)
    for c in courses:
        by_week[sunday_of(dt.date.fromisoformat(c["date"]))].append(c)
    weeks_html = []
    for wi, sun in enumerate(sorted(by_week), start=1):
        wk_courses = by_week[sun]
        days = [sun + dt.timedelta(days=i) for i in range(7)]
        sat = days[6]
        rg = f"{sun.month:02d}.{sun.day:02d} – {sat.month:02d}.{sat.day:02d}"
        # 表頭
        ths = ['<th class="tcol">時間</th>']
        for i, day in enumerate(days):
            ths.append(f'<th><span class="wd">{WD_ZH[i]}</span>'
                       f'<span class="en">{WD_EN[i]}</span>{day.month}/{day.day}</th>')
        # 哪些「整點列」有課
        hours = sorted({int(c["start"][:2]) for c in wk_courses})
        # (日期 → 整點 → 課程清單)
        cell_map = defaultdict(lambda: defaultdict(list))
        for c in wk_courses:
            cell_map[c["date"]][int(c["start"][:2])].append(c)
        rows_html = []
        for h in hours:
            tds = [f'<td class="time">{h:02d}:00</td>']
            for day in days:
                cs = sorted(cell_map.get(day.isoformat(), {}).get(h, []), key=sort_key)
                tds.append("<td>" + "".join(_course_span(c, show_district) for c in cs) + "</td>")
            rows_html.append("<tr>" + "".join(tds) + "</tr>")
        weeks_html.append(
            f'<div class="week"><div class="week-label">'
            f'<span class="no">Week {wi}</span><span class="rg">{rg}</span></div>'
            f'<table><thead><tr>{"".join(ths)}</tr></thead>'
            f'<tbody>{"".join(rows_html)}</tbody></table></div>')
    return "\n".join(weeks_html)


def _course_span(c, show_district=False):
    k, tag, b, bg = TYPE_MAP[c["type"]]
    vn = (f'［{c["district"]}］' if show_district else "") + esc(c["venue"])
    return (f'<span class="course" style="border-color:{b};background:{bg}">'
            f'<span class="tag" style="background:{b};color:#fff">{tag}</span>'
            f'<span class="tm">{tm_str(c)}</span>'
            f'<span class="nm">{esc(c["name"])}</span>'
            f'<span class="vn">{vn}</span></span>')


# ---------------------------------------------------------------- 手機版（日卡片）
def render_mobile(courses, show_district=False):
    by_day = defaultdict(list)
    for c in courses:
        by_day[c["date"]].append(c)
    days_html = []
    for day in sorted(by_day):
        cs = sorted(by_day[day], key=sort_key)
        d = dt.date.fromisoformat(day)
        wd = WD_ZH[(d.weekday() + 1) % 7]
        cards = "".join(_card(c, show_district) for c in cs)
        days_html.append(
            f'<div class="day"><div class="day-h">'
            f'<span class="dn">{d.month}/{d.day}</span>'
            f'<span class="wd">（{wd}）</span>'
            f'<span class="cnt">{len(cs)} 堂</span></div>{cards}</div>')
    return "\n".join(days_html)


def _card(c, show_district=False):
    k, tag, b, bg = TYPE_MAP[c["type"]]
    vn = (f'［{c["district"]}］' if show_district else "") + esc(c["venue"])
    return (f'<div class="card" data-k="{k}" style="border-color:{b};background:{bg}">'
            f'<div class="top"><span class="tag" style="background:{b}">{tag}</span>'
            f'<span class="tm">{tm_str(c)}</span></div>'
            f'<span class="nm">{esc(c["name"])}</span>'
            f'<span class="vn">{vn}</span></div>')


# ---------------------------------------------------------------- 返回主選單按鈕
BACK_BTN = (
    '<a class="homefab" href="../index.html">← 主選單</a>'
    '<style>.homefab{position:fixed;left:16px;bottom:18px;z-index:65;'
    'background:#3a362d;color:#fff;text-decoration:none;font-weight:700;'
    'font-size:14px;padding:10px 16px;border-radius:999px;'
    'box-shadow:0 3px 10px rgba(0,0,0,.25);'
    'font-family:"Noto Sans TC","Microsoft JhengHei",sans-serif;}</style>'
)


def inject_back(page: str):
    return page.replace("</body>", BACK_BTN + "</body>", 1)


# ---------------------------------------------------------------- 統計
def build_stats(courses):
    stats = {"total": len(courses), "by_district": {}}
    for dist in DISTRICTS:
        dc = [c for c in courses if c["district"] == dist]
        by_type = Counter(c["type"] for c in dc)
        stats["by_district"][dist] = {
            "total": len(dc),
            "by_type": {t: by_type.get(t, 0) for t in TYPE_MAP},
            "days": len({c["date"] for c in dc}),
        }
    return stats


# ---------------------------------------------------------------- 主流程
def main():
    print(f"來源 Excel：{XLSX.name}")
    courses, header, skipped, unknown = load_courses()
    print(f"讀入 {len(courses)} 筆（已排除取消 {skipped} 筆）")
    if unknown:
        print("⚠ 無法歸類的地點：")
        for k, v in unknown.most_common():
            print(f"   {k}  ×{v}")

    stats = build_stats(courses)

    tpl_name = {
        ("北投", "big"):   "健康台灣深耕計畫_北投區_2026年6月課程表_日系大字版.html",
        ("北投", "mob"):   "健康台灣深耕計畫_北投區_2026年6月課程表_手機版.html",
        ("士林", "big"):   "健康台灣深耕計畫_士林區_2026年6月課程表_日系大字版.html",
        ("士林", "mob"):   "健康台灣深耕計畫_士林區_2026年6月課程表_手機版.html",
        ("中山", "big"):   "健康台灣深耕計畫_中山區_2026年6月課程表_日系大字版.html",
        ("中山", "mob"):   "健康台灣深耕計畫_中山區_2026年6月課程表_手機版.html",
    }

    for dist in DISTRICTS:
        dc = [c for c in courses if c["district"] == dist]
        # 大字版
        tpl = (TPL / tpl_name[(dist, "big")]).read_text(encoding="utf-8")
        out = splice(tpl, '<div class="week">', '<div class="foot">', render_big(dc))
        (OUT_BIG / tpl_name[(dist, "big")]).write_text(inject_back(out), encoding="utf-8")
        # 手機版
        tpl = (TPL / tpl_name[(dist, "mob")]).read_text(encoding="utf-8")
        out = splice(tpl, '<div class="day">', '<div class="empty">', render_mobile(dc))
        (OUT_MOBILE / tpl_name[(dist, "mob")]).write_text(inject_back(out), encoding="utf-8")
        s = stats["by_district"][dist]
        print(f"  {dist}區：{s['total']} 堂 / {s['days']} 天  "
              + " ".join(f"{TYPE_MAP[t][1]}{n}" for t, n in s["by_type"].items()))

    # 全區綜合（三區合併，課程標出所屬區）
    base_big = (TPL / tpl_name[("北投", "big")]).read_text(encoding="utf-8")
    base_big = (base_big.replace("<title>北投區課程表</title>", "<title>全區綜合課程表</title>")
                        .replace("<h1>北投區　六月課程表</h1>", "<h1>全區綜合　六月課程表</h1>"))
    out = splice(base_big, '<div class="week">', '<div class="foot">',
                 render_big(courses, show_district=True))
    (OUT_BIG / ALL_BIG).write_text(inject_back(out), encoding="utf-8")

    base_mob = (TPL / tpl_name[("北投", "mob")]).read_text(encoding="utf-8")
    base_mob = (base_mob.replace("<title>北投區課程表</title>", "<title>全區綜合課程表</title>")
                        .replace("<h1>北投區　六月課程表</h1>", "<h1>全區綜合　六月課程表</h1>"))
    out = splice(base_mob, '<div class="day">', '<div class="empty">',
                 render_mobile(courses, show_district=True))
    (OUT_MOBILE / ALL_MOB).write_text(inject_back(out), encoding="utf-8")
    print(f"  全區綜合：{len(courses)} 堂")

    # 資料庫 JSON
    db = {
        "generated": "2026-06-05",
        "month": "2026-06",
        "source": XLSX.name,
        "stats": stats,
        "courses": sorted(courses, key=lambda c: (c["district"],) + sort_key(c)),
    }
    (ROOT / "courses.json").write_text(
        json.dumps(db, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"已輸出 courses.json（{len(courses)} 筆）")

    write_index(stats)
    print("已輸出 index.html")


def write_index(stats):
    cards = []
    files = {
        "北投": ("大字版/健康台灣深耕計畫_北投區_2026年6月課程表_日系大字版.html",
                "手機版/健康台灣深耕計畫_北投區_2026年6月課程表_手機版.html"),
        "士林": ("大字版/健康台灣深耕計畫_士林區_2026年6月課程表_日系大字版.html",
                "手機版/健康台灣深耕計畫_士林區_2026年6月課程表_手機版.html"),
        "中山": ("大字版/健康台灣深耕計畫_中山區_2026年6月課程表_日系大字版.html",
                "手機版/健康台灣深耕計畫_中山區_2026年6月課程表_手機版.html"),
    }
    total = stats["total"]
    # 全區綜合（置頂）
    all_types = Counter()
    for dist in DISTRICTS:
        for t, n in stats["by_district"][dist]["by_type"].items():
            all_types[t] += n
    all_days = sum(stats["by_district"][d]["days"] for d in DISTRICTS)  # 顯示用近似
    chips = "".join(
        f'<span class="t t-{TYPE_MAP[t][0]}">{TYPE_MAP[t][1]} {all_types[t]}</span>'
        for t in TYPE_MAP)
    cards.append(f"""    <section class="card card-all">
      <h2>全區綜合 <small>北投・士林・中山　共 {total} 堂</small></h2>
      <div class="types">{chips}</div>
      <div class="links">
        <a href="{esc('大字版/' + ALL_BIG)}">電腦／大字版</a>
        <a href="{esc('手機版/' + ALL_MOB)}">手機版</a>
      </div>
    </section>""")

    for dist in DISTRICTS:
        s = stats["by_district"][dist]
        big, mob = files[dist]
        chips = "".join(
            f'<span class="t t-{TYPE_MAP[t][0]}">{TYPE_MAP[t][1]} {n}</span>'
            for t, n in s["by_type"].items())
        cards.append(f"""    <section class="card">
      <h2>{dist}區 <small>{s['total']} 堂 · {s['days']} 天</small></h2>
      <div class="types">{chips}</div>
      <div class="links">
        <a href="{esc(big)}">電腦／大字版</a>
        <a href="{esc(mob)}">手機版</a>
      </div>
    </section>""")
    page = f"""<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>健康台灣深耕計畫 · 2026年6月課程表</title>
<style>
*{{box-sizing:border-box;}}
body{{margin:0;padding:56px 20px 80px;background:#FBFAF6;color:#262420;
 font-family:"Noto Sans TC","Microsoft JhengHei",sans-serif;line-height:1.6;}}
.wrap{{max-width:760px;margin:0 auto;}}
.head{{text-align:center;margin-bottom:40px;}}
.head .ey{{font-size:12px;letter-spacing:.3em;color:#7a715f;text-transform:uppercase;}}
.head h1{{font-size:28px;font-weight:700;margin:10px 0 6px;color:#23211c;}}
.head .sub{{color:#4f4838;}}
.card{{background:#fff;border:1px solid #e6dfcf;border-radius:14px;padding:22px 24px;
 margin-bottom:18px;box-shadow:0 1px 3px rgba(0,0,0,.05);}}
.card-all{{border:2px solid #b7ad99;background:#fbf7ee;}}
.card h2{{margin:0 0 12px;font-size:22px;color:#23211c;}}
.card h2 small{{font-size:14px;color:#8a8170;font-weight:500;margin-left:8px;}}
.types{{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:16px;}}
.types .t{{font-size:14px;font-weight:700;padding:4px 12px;border-radius:999px;color:#fff;}}
.t-n{{background:#445230;}}.t-s{{background:#72491F;}}.t-o{{background:#385268;}}.t-e{{background:#6E4858;}}
.links{{display:flex;gap:12px;}}
.links a{{flex:1;text-align:center;padding:12px;border-radius:10px;text-decoration:none;
 font-weight:700;background:#3a362d;color:#fff;}}
.links a:last-child{{background:#fff;color:#3a362d;border:1.5px solid #d9cfba;}}
.foot{{text-align:center;margin-top:40px;font-size:13px;letter-spacing:.16em;color:#8a8170;}}
</style></head><body><div class="wrap">
<div class="head"><div class="ey">Healthy Taiwan Program</div>
<h1>2026 年 6 月課程表</h1>
<div class="sub">士林 · 北投 · 中山　四大處方　共 {total} 堂</div></div>
{chr(10).join(cards)}
<div class="foot">台北市醫師公會 ‧ 健康台灣深耕計畫</div>
</div></body></html>"""
    (ROOT / "index.html").write_text(page, encoding="utf-8")


if __name__ == "__main__":
    main()
