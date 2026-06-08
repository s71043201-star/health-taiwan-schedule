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
import shutil
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


# ---------------------------------------------------------------- 健康處方系統資料來源
# 課程時段（slots）的兩條自動來源，皆免經 Excel / Google Sheet：
#   (S) Supabase：n8n 已把 course_slots 同步到 Supabase prescription_data(id='main')，
#       本程式直接讀那份（公開唯讀 anon key，repo 零帳密）。← 雲端自動更新走這條。
#   (A) 健康處方系統 API：直接登入 API 逐月抓 slots/summary（需帳密）。
# 公開 repo 不可寫死帳密，一律由環境變數 / Actions secrets 帶入。
API_BACKEND = os.environ.get("COURSE_API_BACKEND",
                             "https://healthcheck-backend.delixir.cc").rstrip("/")
API_ORIGIN = os.environ.get("COURSE_API_ORIGIN",
                            "https://healthcheck-rx.delixir.cc")

# Supabase 唯讀來源（anon key 已是公開金鑰、RLS 保護，與 tpma-statistics 共用同一份真相）
SUPABASE_URL = os.environ.get(
    "COURSE_SUPABASE_URL", "https://ilcnqpywxaseeyasiwws.supabase.co").rstrip("/")
SUPABASE_ANON_KEY = os.environ.get(
    "COURSE_SUPABASE_KEY",
    "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImlsY25xcHl3"
    "eGFzZWV5YXNpd3dzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzM4MzIzODgsImV4cCI6MjA4OTQwODM4OH0"
    ".TrjIr4IMdvpstkN8tNBQtAEvvNDTLg3XcXXIpptJIs0")
SUPABASE_ROW_ID = os.environ.get("COURSE_SUPABASE_ROW", "main")

# 處方類型代碼 → TYPE_MAP 的中文 key（API 與 Supabase 同樣用英文代碼）
API_TYPE_MAP = {
    "nutrition": "營養處方",
    "exercise":  "運動處方",
    "social":    "社會處方",
    "mental":    "情緒調適處方",
}


def _slot_to_course(s):
    """把一筆 slot（API 或 Supabase 格式）轉成 course dict。
       回傳 (course, None) 成功；或 (None, reason)，reason ∈
       {'cancel', ('loc', 地點), ('type', 類型), 'baddate'}。
       兩種來源欄位名略有差異，這裡統一吃。"""
    status = str(s.get("status") or "").strip()
    if status == "cancelled":                       # 取消不放（對應 Excel 的「取消」）
        return None, "cancel"
    loc = s.get("course_location") or s.get("slot_location") or ""
    district = classify_district(loc)
    if district is None:
        return None, ("loc", str(loc))
    type_code = str(s.get("course_prescription_type") or s.get("course_type") or "").strip().lower()
    ptype = API_TYPE_MAP.get(type_code)
    if ptype is None:
        return None, ("type", type_code)
    d = _parse_date(s.get("slot_date"))
    if d is None:
        return None, "baddate"
    return {
        "name": str(s.get("course_name") or "").strip(),
        "type": ptype,
        "date": d.isoformat(),
        "start": _hhmm(s.get("start_time")),
        "end": _hhmm(s.get("end_time")),
        "venue": str(loc).strip(),
        "district": district,
        "enrolled": _int(s.get("booked_count")),
        "capacity": _int(s.get("capacity_effective") or s.get("capacity")),
        "full": bool(s.get("is_full")),
        "status": status,
    }, None


def _slots_to_courses(slots):
    """共用：把 slots 清單轉成 courses，並彙整略過統計。"""
    courses, skipped_cancel = [], 0
    unknown_loc, unknown_type = Counter(), Counter()
    for s in slots:
        course, reason = _slot_to_course(s)
        if course is not None:
            courses.append(course)
        elif reason == "cancel":
            skipped_cancel += 1
        elif isinstance(reason, tuple) and reason[0] == "loc":
            unknown_loc[reason[1]] += 1
        elif isinstance(reason, tuple) and reason[0] == "type":
            unknown_type[reason[1]] += 1
    return courses, skipped_cancel, unknown_loc, unknown_type


def load_courses_supabase():
    """讀 Supabase prescription_data(id='main') 的 course_slots（n8n 已同步好的那份）。"""
    url = (SUPABASE_URL + "/rest/v1/prescription_data"
           f"?id=eq.{SUPABASE_ROW_ID}&select=course_slots")
    req = urllib.request.Request(url, headers={
        "apikey": SUPABASE_ANON_KEY, "Authorization": "Bearer " + SUPABASE_ANON_KEY})
    with urllib.request.urlopen(req) as r:
        rows = json.loads(r.read().decode("utf-8"))
    if not rows:
        raise SystemExit(f"❌ Supabase 找不到 id='{SUPABASE_ROW_ID}' 的列。")
    slots = rows[0].get("course_slots") or []
    return _slots_to_courses(slots)


def _api_login(account, password):
    body = json.dumps({"account": account, "password": password}).encode("utf-8")
    req = urllib.request.Request(
        API_BACKEND + "/api/v1/auth/login", data=body, method="POST",
        headers={"Content-Type": "application/json", "Origin": API_ORIGIN})
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read().decode("utf-8"))["access_token"]


def _api_get(path, qs, token):
    import urllib.parse
    url = API_BACKEND + path + "?" + urllib.parse.urlencode(qs)
    req = urllib.request.Request(
        url, headers={"Authorization": "Bearer " + token, "Origin": API_ORIGIN})
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read().decode("utf-8"))


def _month_ranges(start: dt.date, end: dt.date):
    """逐月產生 (該月首日, 該月末日)，涵蓋 start~end。"""
    cur = dt.date(start.year, start.month, 1)
    while cur <= end:
        nxt = dt.date(cur.year + cur.month // 12, cur.month % 12 + 1, 1)
        yield cur, nxt - dt.timedelta(days=1)
        cur = nxt


def _hhmm(t):
    """'09:00:00' → '09:00'；容錯空值。"""
    s = str(t or "").strip()
    return s[:5] if len(s) >= 5 else s


def load_courses_api(account, password, start_date, end_date):
    """從健康處方系統 API 逐月抓 slots/summary，組成 courses（與 Supabase 來源共用轉換）。"""
    import urllib.error
    token = _api_login(account, password)
    slots = []
    for mstart, mend in _month_ranges(start_date, end_date):
        qs = {"start_date": mstart.isoformat(), "end_date": mend.isoformat(),
              "recent_page": 1, "recent_page_size": 1}
        try:
            data = _api_get("/api/v1/admin/slots/summary", qs, token)
        except urllib.error.HTTPError as e:
            if e.code in (401, 403):     # token 過期就重登一次再試
                token = _api_login(account, password)
                data = _api_get("/api/v1/admin/slots/summary", qs, token)
            else:
                raise
        slots.extend(data.get("slots", []))
    return _slots_to_courses(slots)


def resolve_source():
    """決定資料來源，回傳 (kind, payload)。優先序：
       1. 'supabase'：COURSE_SOURCE=supabase（或設了 COURSE_SUPABASE_KEY）→ 讀 n8n 同步好的 Supabase。
       2. 'api'：設了 COURSE_API_ACCOUNT + COURSE_API_PASSWORD → 直接登入 API 抓。
       3. 'xlsx'：fallback 到 Google Sheet（COURSE_XLSX_URL）或 data/ 內的 Excel。"""
    src = os.environ.get("COURSE_SOURCE", "").strip().lower()
    acc = os.environ.get("COURSE_API_ACCOUNT", "").strip()
    pwd = os.environ.get("COURSE_API_PASSWORD", "").strip()
    if src == "supabase" or (not src and os.environ.get("COURSE_SUPABASE_KEY", "").strip()):
        return ("supabase", None)
    if src == "api" or (not src and acc and pwd):
        start = (_parse_date(os.environ.get("COURSE_START_DATE", "").strip())
                 or dt.date(2025, 12, 20))
        end_env = os.environ.get("COURSE_END_DATE", "").strip()
        end = _parse_date(end_env) if end_env else (dt.date.today() + dt.timedelta(days=183))
        return ("api", (acc, pwd, start, end))
    return ("xlsx", resolve_xlsx())


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

# 月份中文（索引 = 月份數字）
CN_MONTH = ["", "一", "二", "三", "四", "五", "六",
            "七", "八", "九", "十", "十一", "十二"]

# 各月資料夾內的輸出檔名（月份資訊由資料夾承載）
OUT_NAME = {
    ("北投", "big"): "北投區_大字版.html",
    ("北投", "mob"): "北投區_手機版.html",
    ("士林", "big"): "士林區_大字版.html",
    ("士林", "mob"): "士林區_手機版.html",
    ("中山", "big"): "中山區_大字版.html",
    ("中山", "mob"): "中山區_手機版.html",
    ("全區", "big"): "全區綜合_大字版.html",
    ("全區", "mob"): "全區綜合_手機版.html",
}

# 來源版型（皆以「北投區 / 2026年6月」為基準，產生時再依實際年月重新標記）
TPL_NAME = {
    ("北投", "big"): "健康台灣深耕計畫_北投區_2026年6月課程表_日系大字版.html",
    ("北投", "mob"): "健康台灣深耕計畫_北投區_2026年6月課程表_手機版.html",
    ("士林", "big"): "健康台灣深耕計畫_士林區_2026年6月課程表_日系大字版.html",
    ("士林", "mob"): "健康台灣深耕計畫_士林區_2026年6月課程表_手機版.html",
    ("中山", "big"): "健康台灣深耕計畫_中山區_2026年6月課程表_日系大字版.html",
    ("中山", "mob"): "健康台灣深耕計畫_中山區_2026年6月課程表_手機版.html",
}


def restamp(tpl: str, year: int, month: int):
    """把版型基準的「2026 / 六月 / 6 月 / .06 / M=6」改成實際年月。"""
    return (tpl
            .replace("2026", str(year))
            .replace("六月", CN_MONTH[month] + "月")
            .replace(" . 06", f" . {month:02d}")
            .replace("年 6 月", f"年 {month} 月")
            .replace(",M=6;", f",M={month};"))


def classify_district(loc: str):
    """依上課地點文字判斷行政區；回傳 北投/士林/中山 或 None。"""
    s = str(loc).replace(" ", "")
    # 北投關鍵字
    if any(k in s for k in ("北投", "石牌", "榮陽", "蔡秉勳", "翰譽耳鼻喉",
                            "永安", "瑜伽之光", "臻心", "ULifeFitnes",
                            "ULifeFitness", "泉源", "王永良")):
        return "北投"
    # 士林關鍵字
    if any(k in s for k in ("士林", "天母", "ZenYoga", "小船254", "後街21巷",
                            "福港街", "福華路")):
        return "士林"
    # 中山關鍵字
    if any(k in s for k in ("中山", "圓山", "夢想館", "晨昕", "劍南")):
        return "中山"
    return None


# ---------------------------------------------------------------- 讀取資料
def load_courses(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
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
def render_mobile(courses, show_district=False, hide_past=False):
    by_day = defaultdict(list)
    for c in courses:
        by_day[c["date"]].append(c)
    today = dt.date.today()

    def day_block(day):
        cs = sorted(by_day[day], key=sort_key)
        d = dt.date.fromisoformat(day)
        wd = WD_ZH[(d.weekday() + 1) % 7]
        cards = "".join(_card(c, show_district) for c in cs)
        return (f'<div class="day"><div class="day-h">'
                f'<span class="dn">{d.month}/{d.day}</span>'
                f'<span class="wd">（{wd}）</span>'
                f'<span class="cnt">{len(cs)} 堂</span></div>{cards}</div>')

    upcoming, past = [], []
    for day in sorted(by_day):
        if hide_past and dt.date.fromisoformat(day) < today:
            past.append(day_block(day))
        else:
            upcoming.append(day_block(day))
    out = "\n".join(upcoming)
    if past:                          # 當月已過去的日子收進底部隱藏式選單
        out += (f'\n<details class="pastdays"><summary>'
                f'<span>已過去的日期（{len(past)} 天）</span>'
                f'<span class="arr">›</span></summary>\n{chr(10).join(past)}</details>')
    return out


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


# ---------------------------------------------------------------- 課表頁頂部備註
SYSNOTE = (
    '<div class="sysnote">📌 本頁僅供查詢課表；要<b>預約課程</b>請至 '
    'LINE 圖文選單的「預約課程」。</div>'
    '<style>.sysnote{margin:0 auto 16px;max-width:1120px;background:#FBEFE4;'
    'border:1.5px solid #E0A86B;border-radius:10px;padding:10px 14px;color:#5c4326;'
    'font-size:15px;font-weight:600;line-height:1.5;}.sysnote b{color:#9a5a1f;}</style>'
)


def inject_sysnote(page: str):
    return page.replace('<div class="wrap">', '<div class="wrap">' + SYSNOTE, 1)


# ---------------------------------------------------------------- 課程查詢（下拉選單 + 時段 modal）
COURSE_FILTER_CSS = """<style>
.cfstep{max-width:1120px;margin:0 auto 8px;font-weight:700;color:#4f4838;font-size:15px;
 display:flex;align-items:center;gap:8px;}
.cfstep.s1{margin-top:6px;}
.cfbar{max-width:1120px;margin:0 auto 22px;display:flex;align-items:center;gap:12px;
 background:#fff;border:1.5px solid #e0d9c8;border-radius:12px;padding:12px 16px;
 box-shadow:0 1px 3px rgba(0,0,0,.05);}
.cfbar .cfl{font-weight:700;color:#4f4838;font-size:15px;white-space:nowrap;}
.cfsel{flex:1;min-width:0;font-size:16px;padding:10px 12px;border:1.5px solid #d9cfba;
 border-radius:8px;background:#fbfaf6;color:#23211c;font-family:inherit;cursor:pointer;}
.cfmask{position:fixed;inset:0;background:rgba(35,33,28,.55);z-index:90;display:none;
 align-items:flex-start;justify-content:center;padding:40px 16px;overflow:auto;}
.cfmask.show{display:flex;}
.cfmodal{background:#FBFAF6;max-width:520px;width:100%;border-radius:16px;
 box-shadow:0 12px 40px rgba(0,0,0,.3);overflow:hidden;}
.cfm-h{display:flex;align-items:flex-start;justify-content:space-between;gap:12px;
 padding:18px 20px 8px;border-left:6px solid;}
.cfm-h #cfmTitle{font-size:19px;font-weight:700;color:#23211c;line-height:1.4;}
.cfm-x{border:0;background:#ece5d6;color:#5b5446;width:32px;height:32px;border-radius:50%;
 font-size:15px;cursor:pointer;flex:none;}
.cfm-sub{padding:0 20px 14px;color:#8a8170;font-size:14px;font-weight:600;}
.cfm-body{padding:0 16px 18px;}
.cfslot{display:flex;align-items:center;gap:10px;padding:11px 13px;margin:7px 0;background:#fff;
 border-left:5px solid;border-radius:6px;box-shadow:0 1px 2px rgba(0,0,0,.05);font-size:15px;}
.cfslot .dt{font-weight:700;color:#23211c;white-space:nowrap;}
.cfslot .tm{color:#3f3a30;font-weight:600;white-space:nowrap;}
.cfslot .vn{color:#5b5446;font-size:13px;margin-left:auto;text-align:right;line-height:1.4;}
.pastdays{margin:18px 0 8px;border:1.5px dashed #d9cfba;border-radius:12px;background:#fdfcf9;}
.pastdays>summary{list-style:none;cursor:pointer;padding:14px 16px;font-weight:700;color:#8a8170;
 display:flex;align-items:center;justify-content:space-between;}
.pastdays>summary::-webkit-details-marker{display:none;}
.pastdays>summary .arr{color:#b7ad99;transition:transform .2s;font-size:18px;}
.pastdays[open]>summary .arr{transform:rotate(90deg);}
.pastdays .day{margin:0 12px 12px;}
@media(max-width:760px){.cfbar{flex-direction:column;align-items:stretch;gap:8px;}
 .cfstep{padding-left:14px;}
 .cfslot{flex-wrap:wrap;}
 .cfslot .vn{flex-basis:100%;margin-left:0;text-align:left;margin-top:3px;}}
</style>"""


def inject_course_filter(page: str, courses, show_district=False):
    """注入「課程下拉選單 → 點選跳出該課全部時段」的查詢介面。
       不爬 DOM，直接把該頁課程依名稱分組後以 JSON 嵌入，modal 讀它呈現。"""
    if not courses:
        return page
    order = list(TYPE_MAP)
    by_name = defaultdict(list)
    for c in sorted(courses, key=sort_key):
        by_name[c["name"]].append(c)
    names = sorted(by_name, key=lambda n: (order.index(by_name[n][0]["type"])
                   if by_name[n][0]["type"] in order else 99, n))

    data = []
    for n in names:
        lst = by_name[n]
        _, tag, b, _bg = TYPE_MAP[lst[0]["type"]]
        slots = []
        for c in lst:
            d = dt.date.fromisoformat(c["date"])
            _, _tag, sb, _ = TYPE_MAP[c["type"]]
            venue = (f'［{c["district"]}］' if show_district else "") + c["venue"]
            slots.append({"date": f"{d.month}/{d.day}", "wd": WD_ZH[(d.weekday() + 1) % 7],
                          "tm": tm_str(c), "venue": esc(venue), "b": sb})
        data.append({"name": n, "dn": esc(n), "tag": tag, "b": b, "slots": slots})
    payload = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")

    # 下拉選項由 JS 依目前選的處方類型動態產生（見 populate），這裡只放空殼
    bar = (COURSE_FILTER_CSS
           + '<div class="cfbar"><span class="cfl">課程查詢</span>'
           + '<select id="cfSel" class="cfsel"></select></div>')
    # 步驟①標籤：注入到處方類型篩選（大字版 legend / 手機版 bar）上方
    step1 = '<div class="cfstep s1">先選處方類型（想看全部就點「全部」）</div>'
    modal = (
        '<div class="cfmask" id="cfMask"><div class="cfmodal">'
        '<div class="cfm-h" id="cfmHead"><span id="cfmTitle"></span>'
        '<button class="cfm-x" id="cfmX" aria-label="關閉">✕</button></div>'
        '<div class="cfm-sub" id="cfmSub"></div><div class="cfm-body" id="cfmBody"></div>'
        '</div></div>'
        '<script>(function(){var D=' + payload + ';'
        'var PH=\'<option value="">— 選擇課程，查看所有時段 —</option>\';'
        'var sel=document.getElementById("cfSel"),mask=document.getElementById("cfMask"),'
        'head=document.getElementById("cfmHead"),ttl=document.getElementById("cfmTitle"),'
        'sub=document.getElementById("cfmSub"),body=document.getElementById("cfmBody");'
        # 依處方類型 tag 過濾下拉內容（tag 為空=全部）；value 仍是 D 的原始索引
        'function populate(tag){var h=PH,i;for(i=0;i<D.length;i++){if(!tag||D[i].tag===tag){'
        'h+=\'<option value="\'+i+\'">\'+D[i].dn+"（"+D[i].slots.length+" 場）</option>";}}'
        'sel.innerHTML=h;sel.value="";}'
        'function open(i){var c=D[i];if(!c)return;ttl.textContent=c.name;'
        'head.style.borderLeftColor=c.b;sub.textContent=c.tag+"　共 "+c.slots.length+" 場時段";'
        'body.innerHTML=c.slots.map(function(s){return \'<div class="cfslot" style="border-color:\''
        '+s.b+\'"><span class="dt">\'+s.date+"（"+s.wd+"）</span>"'
        '+\'<span class="tm">\'+s.tm+"</span>"+\'<span class="vn">\'+s.venue+"</span></div>";'
        '}).join("");mask.classList.add("show");}'
        'function close(){mask.classList.remove("show");sel.value="";}'
        'sel.onchange=function(){if(sel.value!=="")open(+sel.value);};'
        'document.getElementById("cfmX").onclick=close;'
        'mask.onclick=function(e){if(e.target===mask)close();};'
        'document.addEventListener("keydown",function(e){if(e.key==="Escape")close();});'
        # 連動處方類型篩選：大字版 .legend .it（文字）/ 手機版 .bar .chip（data-k 代碼）
        'var K2T={n:"營養",s:"運動",o:"社會",e:"情緒",all:null};'
        'function tagOf(el){if(el.getAttribute&&el.getAttribute("data-k")!=null){'
        'var k=el.getAttribute("data-k");return K2T.hasOwnProperty(k)?K2T[k]:null;}'
        'if(el.classList.contains("allk")||el.textContent.trim()==="全部")return null;'
        'return el.textContent.trim().slice(0,2);}'
        'var chips=document.querySelectorAll(".legend .it, .bar .chip");'
        '[].slice.call(chips).forEach(function(el){'
        'el.addEventListener("click",function(){populate(tagOf(el));});});'
        'populate(null);'
        '})();</script>')
    # ① 標籤放在處方類型篩選上方（大字版 legend / 手機版 bar）
    filt = ('<div class="legend">' if '<div class="legend">' in page
            else ('<div class="bar">' if '<div class="bar">' in page else None))
    if filt:
        page = page.replace(filt, step1 + filt, 1)
    # ② 課程查詢放在篩選下面：注入到課表內容最前面（取最早出現的錨點，
    #    含「已過去日期」收合區，避免被注入到收合選單內部）
    cands = [m for m in ('<details class="pastdays">', '<div class="week">', '<div class="day">')
             if m in page]
    if cands:
        marker = min(cands, key=page.index)
        page = page.replace(marker, bar + marker, 1)
    else:
        page = page.replace('<div class="wrap">', '<div class="wrap">' + bar, 1)
    page = page.replace('</body>', modal + '</body>', 1)
    return page


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


# ---------------------------------------------------------------- 入口頁共用樣式
INDEX_CSS = """
*{box-sizing:border-box;}
body{margin:0;padding:56px 20px 80px;background:#FBFAF6;color:#262420;
 font-family:"Noto Sans TC","Microsoft JhengHei",sans-serif;line-height:1.6;}
.wrap{max-width:760px;margin:0 auto;}
.back{position:fixed;left:16px;bottom:18px;z-index:65;background:#3a362d;color:#fff;
 text-decoration:none;font-weight:700;font-size:14px;padding:10px 16px;border-radius:999px;
 box-shadow:0 3px 10px rgba(0,0,0,.25);}
.head{text-align:center;margin-bottom:40px;}
.head .ey{font-size:12px;letter-spacing:.3em;color:#7a715f;text-transform:uppercase;}
.head h1{font-size:28px;font-weight:700;margin:10px 0 6px;color:#23211c;}
.head .sub{color:#4f4838;}
.card{background:#fff;border:1px solid #e6dfcf;border-radius:14px;padding:22px 24px;
 margin-bottom:18px;box-shadow:0 1px 3px rgba(0,0,0,.05);}
.card-all{border:2px solid #b7ad99;background:#fbf7ee;}
.card h2{margin:0 0 12px;font-size:22px;color:#23211c;}
.card h2 small{font-size:14px;color:#8a8170;font-weight:500;margin-left:8px;}
.types{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:16px;}
.types .t{font-size:14px;font-weight:700;padding:4px 12px;border-radius:999px;color:#fff;}
.t-n{background:#445230;}.t-s{background:#72491F;}.t-o{background:#385268;}.t-e{background:#6E4858;}
.links{display:flex;gap:12px;}
.links a{flex:1;text-align:center;padding:12px;border-radius:10px;text-decoration:none;
 font-weight:700;background:#3a362d;color:#fff;}
.links a:last-child{background:#fff;color:#3a362d;border:1.5px solid #d9cfba;}
.month{background:#fff;border:1.5px solid #d9cfba;border-radius:14px;padding:22px 24px;
 margin-bottom:16px;text-decoration:none;color:#23211c;box-shadow:0 1px 3px rgba(0,0,0,.05);
 display:flex;align-items:center;justify-content:space-between;}
.month b{font-size:22px;}.month span{color:#8a8170;font-weight:600;}
.past{margin:0 0 18px;border:1.5px dashed #d9cfba;border-radius:14px;background:#fdfcf9;}
.past>summary{list-style:none;cursor:pointer;padding:16px 20px;font-weight:700;
 color:#6b6354;display:flex;align-items:center;justify-content:space-between;}
.past>summary::-webkit-details-marker{display:none;}
.past>summary .arr{color:#b7ad99;transition:transform .2s;font-size:18px;}
.past[open]>summary .arr{transform:rotate(90deg);}
.past .month{margin:0 14px 12px;padding:14px 18px;}
.past .month b{font-size:18px;}
.notice{background:#FBEFE4;border:1.5px solid #E0A86B;border-radius:12px;
 padding:16px 18px;margin-bottom:26px;color:#5c4326;}
.notice .nt{font-weight:800;font-size:16px;margin:0 0 4px;color:#9a5a1f;}
.notice p{margin:0 0 6px;font-size:15px;line-height:1.6;}
.notice b{color:#9a5a1f;}
.notice img{display:block;width:100%;max-width:420px;margin:10px auto 0;
 border-radius:10px;border:1px solid #e6dfcf;}
.foot{text-align:center;margin-top:40px;font-size:13px;letter-spacing:.16em;color:#8a8170;}
"""


def notice_html(img_prefix):
    """『僅供查詢』提示區塊；img_prefix 為相對於該頁到網站根目錄的前綴。"""
    return f"""<div class="notice">
<p class="nt">⚠️ 本系統僅供「課表查詢」</p>
<p>要<b>預約課程</b>，請回到 LINE 的圖文選單，點選左上角的「<b>預約課程</b>」（如下圖紅框處）。</p>
<img src="{img_prefix}assets/booking-location.jpg" alt="預約課程位於圖文選單左上角">
</div>"""


# ---------------------------------------------------------------- 主流程
def main():
    kind, src = resolve_source()
    unknown_type = None
    if kind == "supabase":
        print(f"來源：Supabase course_slots（{SUPABASE_URL}，id='{SUPABASE_ROW_ID}'，由 n8n 同步）")
        courses, skipped, unknown, unknown_type = load_courses_supabase()
        source_name = "supabase:n8n-sync"
    elif kind == "api":
        account, password, start, end = src
        print(f"來源：健康處方系統 API（帳號 {account}，{start} ~ {end}）")
        courses, skipped, unknown, unknown_type = load_courses_api(account, password, start, end)
        source_name = "healthcheck-api"
    else:
        print(f"來源 Excel：{src.name}")
        courses, _header, skipped, unknown = load_courses(src)
        source_name = src.name
    if unknown_type:
        print("⚠ 未知處方類型（已略過）：")
        for k, v in unknown_type.most_common():
            print(f"   {k}  ×{v}")
    print(f"讀入 {len(courses)} 筆（已排除取消 {skipped} 筆）")
    if unknown:
        print("⚠ 無法歸類的地點：")
        for k, v in unknown.most_common():
            print(f"   {k}  ×{v}")

    # 依年月分組
    by_month = defaultdict(list)
    for c in courses:
        d = dt.date.fromisoformat(c["date"])
        by_month[(d.year, d.month)].append(c)
    month_keys = sorted(by_month)

    # 清掉舊的月份輸出資料夾，避免殘留已不存在的月份
    for old in ROOT.glob("[0-9][0-9][0-9][0-9]-[0-9][0-9]"):
        if old.is_dir():
            shutil.rmtree(old)

    month_infos = []
    for (year, month) in month_keys:
        mc = by_month[(year, month)]
        mdir = ROOT / f"{year}-{month:02d}"
        big_dir, mob_dir = mdir / "大字版", mdir / "手機版"
        big_dir.mkdir(parents=True, exist_ok=True)
        mob_dir.mkdir(parents=True, exist_ok=True)
        stats = build_stats(mc)
        today = dt.date.today()
        is_current = (year, month) == (today.year, today.month)   # 只有當月才收合已過去的日期

        for dist in DISTRICTS:
            dc = [c for c in mc if c["district"] == dist]
            tpl = restamp((TPL / TPL_NAME[(dist, "big")]).read_text(encoding="utf-8"), year, month)
            out = splice(tpl, '<div class="week">', '<div class="foot">', render_big(dc))
            (big_dir / OUT_NAME[(dist, "big")]).write_text(
                inject_back(inject_sysnote(inject_course_filter(out, dc))), encoding="utf-8")
            tpl = restamp((TPL / TPL_NAME[(dist, "mob")]).read_text(encoding="utf-8"), year, month)
            out = splice(tpl, '<div class="day">', '<div class="empty">', render_mobile(dc, hide_past=is_current))
            (mob_dir / OUT_NAME[(dist, "mob")]).write_text(
                inject_back(inject_sysnote(inject_course_filter(out, dc))), encoding="utf-8")

        # 全區綜合（三區合併，課程標出所屬區）
        big = restamp((TPL / TPL_NAME[("北投", "big")]).read_text(encoding="utf-8"), year, month)
        big = (big.replace("<title>北投區課程表</title>", "<title>全區綜合課程表</title>")
                  .replace("<h1>北投區　", "<h1>全區綜合　"))
        out = splice(big, '<div class="week">', '<div class="foot">',
                     render_big(mc, show_district=True))
        (big_dir / OUT_NAME[("全區", "big")]).write_text(
            inject_back(inject_sysnote(inject_course_filter(out, mc, show_district=True))), encoding="utf-8")

        mob = restamp((TPL / TPL_NAME[("北投", "mob")]).read_text(encoding="utf-8"), year, month)
        mob = (mob.replace("<title>北投區課程表</title>", "<title>全區綜合課程表</title>")
                  .replace("<h1>北投區　", "<h1>全區綜合　"))
        out = splice(mob, '<div class="day">', '<div class="empty">',
                     render_mobile(mc, show_district=True, hide_past=is_current))
        (mob_dir / OUT_NAME[("全區", "mob")]).write_text(
            inject_back(inject_sysnote(inject_course_filter(out, mc, show_district=True))), encoding="utf-8")

        write_month_index(mdir, year, month, stats)
        month_infos.append((year, month, len(mc)))
        by_type = " ".join(f"{TYPE_MAP[t][1]}{sum(1 for c in mc if c['type'] == t)}" for t in TYPE_MAP)
        print(f"  {year}-{month:02d}：{len(mc)} 堂  {by_type}")

    # 資料庫 JSON（所有月份）
    db = {
        "months": [f"{y}-{m:02d}" for (y, m) in month_keys],
        "source": source_name,
        "total": len(courses),
        "courses": sorted(courses, key=lambda c: (c["district"],) + sort_key(c)),
    }
    (ROOT / "courses.json").write_text(
        json.dumps(db, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"已輸出 courses.json（{len(courses)} 筆）")

    write_root_index(month_infos)
    print(f"已輸出 index.html（{len(month_infos)} 個月份）")


def _region_cards(stats):
    """月份內各區卡片（連結相對於該月資料夾）。"""
    cards = []
    all_types = Counter()
    for dist in DISTRICTS:
        for t, n in stats["by_district"][dist]["by_type"].items():
            all_types[t] += n
    chips = "".join(f'<span class="t t-{TYPE_MAP[t][0]}">{TYPE_MAP[t][1]} {all_types[t]}</span>'
                    for t in TYPE_MAP)
    cards.append(f"""    <section class="card card-all">
      <h2>全區綜合 <small>北投・士林・中山　共 {stats['total']} 堂</small></h2>
      <div class="types">{chips}</div>
      <div class="links">
        <a href="{esc('手機版/' + OUT_NAME[('全區', 'mob')])}" style="background:#3a362d;color:#fff;border:0;">查看課表 ›</a>
      </div>
    </section>""")
    for dist in DISTRICTS:
        s = stats["by_district"][dist]
        chips = "".join(f'<span class="t t-{TYPE_MAP[t][0]}">{TYPE_MAP[t][1]} {n}</span>'
                        for t, n in s["by_type"].items())
        cards.append(f"""    <section class="card">
      <h2>{dist}區 <small>{s['total']} 堂 · {s['days']} 天</small></h2>
      <div class="types">{chips}</div>
      <div class="links">
        <a href="{esc('手機版/' + OUT_NAME[(dist, 'mob')])}" style="background:#3a362d;color:#fff;border:0;">查看課表 ›</a>
      </div>
    </section>""")
    return "\n".join(cards)


def write_month_index(mdir, year, month, stats):
    page = f"""<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>健康台灣深耕計畫 · {year}年{month}月課程表</title>
<style>{INDEX_CSS}</style></head><body><div class="wrap">
<a class="back" href="../index.html">← 選擇月份</a>
<div class="head"><div class="ey">Healthy Taiwan Program</div>
<h1>{year} 年 {month} 月課程表</h1>
<div class="sub">士林 · 北投 · 中山　四大處方　共 {stats['total']} 堂</div></div>
{notice_html("../")}
{_region_cards(stats)}
<div class="foot">台北市醫師公會 ‧ 健康台灣深耕計畫</div>
</div></body></html>"""
    (mdir / "index.html").write_text(page, encoding="utf-8")


def _month_btn(year, month, total):
    href = esc(f"{year}-{month:02d}/index.html")
    return (f'<a class="month" href="{href}">'
            f'<b>{year} 年 {month} 月</b><span>{total} 堂 ›</span></a>')


def write_root_index(month_infos):
    """當月與未來月份正常列出；過去月份收進可點開的『過往月份』區塊。"""
    today = dt.date.today()
    cur = (today.year, today.month)
    upcoming = [mi for mi in month_infos if (mi[0], mi[1]) >= cur]
    past = [mi for mi in month_infos if (mi[0], mi[1]) < cur]

    parts = []
    if past:                                                # 過往月份收合，放最上面
        past_btns = "\n".join(_month_btn(*mi) for mi in past)  # 舊到新（2→5 月）
        parts.append(
            '<details class="past"><summary>'
            f'<span>過往月份（{len(past)}）</span><span class="arr">›</span></summary>'
            f'{past_btns}</details>')
    if upcoming:
        parts.append("\n".join(_month_btn(*mi) for mi in upcoming))
    elif past:
        parts.append('<p style="text-align:center;color:#8a8170">本月與未來尚無課程，'
                     '可展開上方查詢過往月份。</p>')
    body = ("\n".join(parts) if parts
            else '<p style="text-align:center;color:#8a8170">目前沒有課程資料</p>')
    page = f"""<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>健康台灣深耕計畫 · 課程表</title>
<style>{INDEX_CSS}</style></head><body><div class="wrap">
<div class="head"><div class="ey">Healthy Taiwan Program</div>
<h1>課程表</h1>
<div class="sub">請選擇月份</div></div>
{notice_html("")}
{body}
<div class="foot">台北市醫師公會 ‧ 健康台灣深耕計畫</div>
</div></body></html>"""
    (ROOT / "index.html").write_text(page, encoding="utf-8")


if __name__ == "__main__":
    main()
