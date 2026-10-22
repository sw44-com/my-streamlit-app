import io
import json
import os
import re
from datetime import date, datetime, timedelta

import numpy as np
import pandas as pd
import streamlit as st
from streamlit_gsheets import GSheetsConnection

BASE = os.path.dirname(os.path.abspath(__file__))
SAVE_DIR = os.path.join(BASE, "saved_plans")
DC_PATH = os.path.join(BASE, "discontinued_list.xlsx")
DC_META = os.path.join(BASE, "discontinued_meta.json")
DC_MARKS = {11: "단종", 22: "단종예정"}
OV_PATH = os.path.join(BASE, "mart_overrides.json")

LABEL = dict(code="상품코드", name="상품명", on_stock="온라인창고재고", mart_stock="마트창고재고", exp="유통기한(소비기한)",
             remain="남은유통기한", in_date="최근입고날짜", in_qty="최근입고수량", all4="전채널 판매수량(직전4주)",
             stockout="품절일수", sell_days="실판매일수", mt4="마트/도매 직전4주", mt_ly4="마트/도매 작년직전4주",
             mt_lyn4="마트/도매 작년이후4주", on4="비앤씨마켓 직전4주", on_ly4="비앤씨마켓 작년직전4주",
             on_lyn4="비앤씨마켓 작년이후4주", mgr="상품관리담당자", box="박스당수량", date="날짜")
LAYOUT_REQ = list(LABEL)
METH = {"S": "계절보정", "A": "단순평균"}
NUM = ["on_stock", "mart_stock", "in_qty", "all4", "stockout", "sell_days", "mt4", "mt_ly4",
       "mt_lyn4", "on4", "on_ly4", "on_lyn4", "box", "oqty"]


# 구글 시트 연결 객체 생성
def get_gsheets_connection():
    try:
        return st.connection("gsheets", type=GSheetsConnection)
    except Exception:
        return None


@st.cache_data(show_spinner=False)
def parse_dc(data: bytes):
    from openpyxl import load_workbook
    ws = load_workbook(io.BytesIO(data), data_only=True).active
    hdr = next((r for r in range(1, 11) if str(ws.cell(r, 2).value or "").strip() == "상품코드"), None)
    if hdr is None:
        return None, ["B열 머리글 '상품코드'를 찾지 못했습니다 (1~10행 확인)."]
    marks, warns, unknown, dup = {}, [], {}, []
    for r in range(hdr + 1, ws.max_row + 1):
        code = ws.cell(r, 2).value
        if code is None:
            continue
        code = str(int(code)) if isinstance(code, float) and code.is_integer() else str(code).strip()
        raw = ws.cell(r, 1).value
        try:
            mk = int(float(raw)) if raw not in (None, "") else None
        except (TypeError, ValueError):
            mk = -1
        if mk is not None and mk not in DC_MARKS:
            unknown[raw] = unknown.get(raw, 0) + 1
            mk = None
        if code in marks:
            dup.append(code)
            if marks[code] is None:
                marks[code] = mk
            continue
        marks[code] = mk
    marks = {k: v for k, v in marks.items() if v is not None}
    if unknown:
        warns.append("11/22가 아닌 표시는 무시했습니다: " + ", ".join(f"{k!r}×{n}" for k, n in unknown.items()))
    if dup:
        warns.append(f"중복 상품코드 {len(dup)}개 (먼저 나온 표시 사용): " + ", ".join(dup[:5]))
    return marks, warns


def load_dc(up_dc):
    name, saved_at = None, None
    if up_dc is not None:
        data, name = up_dc.getvalue(), up_dc.name
        marks, warns = parse_dc(data)
        if marks is None:
            st.sidebar.error(warns[0])
            return None
        try:
            same = os.path.exists(DC_PATH) and open(DC_PATH, "rb").read() == data
            if not same:
                with open(DC_PATH, "wb") as f:
                    f.write(data)
                with open(DC_META, "w", encoding="utf-8") as f:
                    json.dump(dict(file_name=name, saved_at=datetime.now().strftime("%Y-%m-%d %H:%M")), f, ensure_ascii=False)
        except Exception as e:
            st.sidebar.warning(f"단종 목록 저장 실패: {e}")
    elif os.path.exists(DC_PATH):
        with open(DC_PATH, "rb") as f:
            marks, warns = parse_dc(f.read())
        if marks is None:
            return None
    else:
        return None
    try:
        with open(DC_META, encoding="utf-8") as f:
            meta = json.load(f)
        name, saved_at = meta.get("file_name", name), meta.get("saved_at")
    except Exception:
        pass
    vals = list(marks.values())
    return dict(marks=marks, warns=warns, file_name=name, saved_at=saved_at, n11=vals.count(11), n22=vals.count(22))


def num(s):
    return pd.to_numeric(s, errors="coerce").fillna(0)


@st.cache_data(show_spinner=False)
def sheet_names(data: bytes):
    return pd.ExcelFile(io.BytesIO(data)).sheet_names


def detect_layout(raw):
    nz = lambda v: re.sub(r"\s+", "", str(v)) if pd.notna(v) else ""
    grp, g = [], ""
    for j in range(raw.shape[1]):
        if nz(raw.iloc[0, j]):
            g = nz(raw.iloc[0, j])
        grp.append(g)
    nm = [nz(raw.iloc[1, j]) or nz(raw.iloc[0, j]) for j in range(raw.shape[1])]

    def find(pred):
        return next((j for j in range(len(nm)) if pred(grp[j], nm[j])), None)

    L = dict(
        code=find(lambda g, n: n == "상품코드"), name=find(lambda g, n: n == "상품명"),
        on_stock=find(lambda g, n: "온라인창고재고" in n), mart_stock=find(lambda g, n: "마트창고재고" in n),
        exp=find(lambda g, n: n in ("유통기한", "소비기한")),
        remain=find(lambda g, n: n.startswith(("남은유통기한", "남은소비기한"))),
        in_date=find(lambda g, n: n == "최근입고날짜"), in_qty=find(lambda g, n: n == "최근입고수량"),
        stockout=find(lambda g, n: n == "품절일수"), sell_days=find(lambda g, n: n == "실판매일수"),
        mgr=find(lambda g, n: n == "상품관리담당자"), box=find(lambda g, n: n == "박스당수량"),
        date=find(lambda g, n: n == "날짜"),
        odate=find(lambda g, n: n.startswith("최근발주일")), oqty=find(lambda g, n: n == "최근발주수량"),
        eta=find(lambda g, n: n == "입고예정일"))
    for i in (1, 2, 3):
        L[f"m{i}"] = find(lambda g, n, i=i: n == f"{i}개월전")
    sufs = {"4": "(직전4주)", "_ly4": "(작년직전4주)", "_lyn4": "(작년이후4주)"}
    for key, gname in (("all", "전채널"), ("mt", "마트/도매"), ("on", "비앤씨마켓")):
        for k, sf in sufs.items():
            L[f"{key}{k}"] = find(lambda g, n, gname=gname, sf=sf: gname in g and n.endswith(sf))
    return L


@st.cache_data(show_spinner="엑셀 읽는 중...")
def load_base(data: bytes, sheet: str):
    raw = pd.read_excel(io.BytesIO(data), sheet_name=sheet, header=None)
    L = detect_layout(raw)
    bad = [LABEL[k] for k in LAYOUT_REQ if L.get(k) is None]
    if bad:
        return None, bad
    d = raw.iloc[3:]
    d = d[d[L["code"]].notna()].reset_index(drop=True)
    keys = ["code", "name", "on_stock", "mart_stock", "exp", "remain", "in_date", "in_qty", "all4", "stockout",
            "sell_days", "mt4", "mt_ly4", "mt_lyn4", "on4", "on_ly4", "on_lyn4", "mgr", "box", "odate", "oqty", "eta"]
    g = pd.DataFrame({k: (d[L[k]] if L.get(k) is not None else np.nan) for k in keys})
    for k in NUM:
        g[k] = num(g[k])
    g["code"] = g["code"].astype(str).str.strip()
    has_m = all(L.get(f"m{i}") is not None for i in (1, 2, 3))
    g["m3ref"] = (sum(num(d[L[f"m{i}"]]) for i in (1, 2, 3)) / 3 * 28 / 30) if has_m else np.nan
    g["exp_str"] = g["exp"].astype(str).str[:10]
    g["remain"] = pd.to_numeric(g["remain"], errors="coerce")
    g["lot"] = g["on_stock"] + g["mart_stock"]

    rows = []
    for code, x in g.groupby("code", sort=False):
        f = x.iloc[0]
        lots = x[x["lot"] > 0].assign(_k=lambda t: t["remain"].fillna(1e9)).sort_values("_k")
        fmt = lambda s: "무기한" if s.startswith("9999") else s
        if len(lots):
            e = lots.iloc[0]
            exp_min, remain_min = fmt(e["exp_str"]), e["remain"]
            detail = " / ".join(f"{fmt(r.exp_str)}({r.lot:g})" for r in lots.itertuples())
        else:
            exp_min, remain_min, detail = "", np.nan, ""
        rows.append(dict(
            code=code, name=str(f["name"]), mgr=f["mgr"],
            in_date=str(f["in_date"])[:10] if pd.notna(f["in_date"]) else "", in_qty=f["in_qty"],
            box=f["box"], on_stock=x["on_stock"].sum(), mart_stock=x["mart_stock"].sum(),
            exp_min=exp_min, remain_min=remain_min, detail=detail,
            all4=f["all4"], sell_days=f["sell_days"], m3ref=f["m3ref"],
            odate=str(f["odate"])[:10] if pd.notna(f["odate"]) else "", oqty=f["oqty"],
            eta=str(f["eta"]) if pd.notna(f["eta"]) else "",
            mt4=f["mt4"], mt_ly4=f["mt_ly4"], mt_lyn4=f["mt_lyn4"],
            on4=f["on4"], on_ly4=f["on_ly4"], on_lyn4=f["on_lyn4"]))
    b = pd.DataFrame(rows)
    b["auto"] = (b["mt_ly4"] > 0) | (b["mt_lyn4"] > 0)
    return b, bad


def apply_incoming(b, bdate, days, on):
    inc, inc_d, stat = np.zeros(len(b)), [""] * len(b), dict(valid=0, expired=0, total=0)
    for i, r in enumerate(b.itertuples()):
        if not (r.oqty and r.oqty > 0 and r.eta):
            continue
        stat["total"] += 1
        try:
            od = pd.to_datetime(r.odate).date()
            m = re.search(r"(\d{1,2})/(\d{1,2})", r.eta)
            eta = date(od.year, int(m.group(1)), int(m.group(2)))
            if eta < od:
                eta = date(od.year + 1, eta.month, eta.day)
        except Exception:
            continue
        rec = pd.to_datetime(str(r.in_date).replace("/", "-"), errors="coerce")
        if pd.notna(rec) and rec.date() >= od:
            continue
        if eta < bdate:
            stat["expired"] += 1
            continue
        if on and od <= bdate and eta <= bdate + timedelta(days=days):
            inc[i], inc_d[i] = r.oqty, f"{eta.month}/{eta.day}"
            stat["valid"] += 1
    b = b.copy()
    b["inc"], b["inc_date"] = inc, inc_d
    return b, stat


def season(this4, ly4, lyn4, cap):
    ly4, lyn4, this4 = ly4.to_numpy(float), lyn4.to_numpy(float), this4.to_numpy(float)
    f = np.where(ly4 > 0, lyn4 / np.where(ly4 > 0, ly4, 1), 1.0)
    if cap > 0:
        f = np.minimum(f, cap)
    return this4 * f


def estimates(b, flags, cap, min_base=0):
    b = b.copy()
    on_s = season(b["on4"], b["on_ly4"], b["on_lyn4"], cap)
    mt_s = season(b["mt4"], b["mt_ly4"], b["mt_lyn4"], cap)
    on_a, mt_a = b["on4"].to_numpy(float), b["mt4"].to_numpy(float)
    if min_base > 0:
        on_s = np.where(b["on_ly4"].to_numpy() >= min_base, on_s, on_a)
        mt_s = np.where(b["mt_ly4"].to_numpy() >= min_base, mt_s, mt_a)
    on = b["code"].map(flags).fillna(False).to_numpy(bool)
    b["on_S"], b["mt_S"] = on_s, np.where(on, mt_s, 0.0)
    b["on_A"], b["mt_A"] = on_a, np.where(on, mt_a, 0.0)
    b["mart_on"] = np.where(on, "반영", "제외")
    return b


def rec_qty(short, lot):
    short = np.maximum(short, 0)
    lot_safe = np.where(lot > 0, lot, 1)
    return np.where(short <= 0, 0, np.where(lot > 0, np.ceil(short / lot_safe) * lot, np.ceil(short)))


def build(b, view, meths, thr, qty=None):
    tot = b["on_stock"] + b["mart_stock"]
    inc = b["inc"] if "inc" in b else pd.Series(0.0, index=b.index)
    on_eff = b["on_stock"] + inc
    stock = {"전체": tot + inc, "온라인": on_eff, "마트": b["mart_stock"]}[view]
    o = pd.DataFrame({"상품코드": b["code"], "상품명": b["name"],
                      "직전입고일": b["in_date"], "직전입고량": b["in_qty"],
                      "소비기한(최단)": b["exp_min"], "남은일수": b["remain_min"],
                      "소비기한(재고별)": b["detail"]})
    if b["mark"].notna().any():
        o.insert(2, "구분", b["mark"].map({float(k): v for k, v in DC_MARKS.items()}).fillna(""))
    o["합계재고"], o["온라인재고"], o["마트재고"] = tot, b["on_stock"], b["mart_stock"]
    o["마트보정"] = b["mart_on"]
    if inc.any():
        o["입고예정"], o["입고예정일"] = inc, b["inc_date"]
    for nm, v, cs in (("온라인", "온라인", ("on4", "on_ly4", "on_lyn4")), ("마트·도매", "마트", ("mt4", "mt_ly4", "mt_lyn4"))):
        if view in ("전체", v):
            for lab, c in zip(("직전4주", "작년직전4주", "작년이후4주"), cs):
                o[f"{nm} {lab}판매"] = b[c]
    needs, days_l, short_l, on_short_l = [], [], [], []
    for m in meths:
        p = f"[{METH[m]}] " if len(meths) > 1 else ""
        on_e, mt_e = b[f"on_{m}"], b[f"mt_{m}"]
        est = {"전체": on_e + mt_e, "온라인": on_e, "마트": mt_e}[view]
        if view == "전체":
            o[p + "온라인예상"], o[p + "마트예상(보정)"] = on_e.round(1), mt_e.round(1)
            on_sh = (on_e - on_eff).clip(lower=0)
            on_short_l.append(on_sh.to_numpy())
        pos = est > 0
        ratio = np.where(pos, stock / est.where(pos, 1), np.nan)
        need = pos & ((stock < est * thr) if thr else True)
        if qty:
            need = need & (est > qty)
        short = (est - stock).clip(lower=0)
        o[p + "4주예상판매"] = est.round(1)
        o[p + "소진예상일"] = np.round(ratio * 28, 0)
        o[p + "추천생산량"] = rec_qty(short.to_numpy(float), b["in_qty"].to_numpy(float))
        needs.append(need.to_numpy())
        days_l.append(np.where(pos, ratio * 28, np.inf))
        short_l.append(short.to_numpy())
    if "m3ref" in b and b["m3ref"].notna().any():
        o["3개월평균(참고)"] = b["m3ref"].round(1)
    need = np.any(needs, axis=0)
    dmin = np.min(days_l, axis=0)
    o["_need"], o["_dmin"], o["_short"] = need, dmin, np.max(short_l, axis=0)
    o["_key"] = np.max(on_short_l, axis=0) if view == "전체" else o["_short"]
    st_ = np.select([stock.to_numpy() <= 0, dmin <= 7, dmin <= 14], ["🔴 품절", "🔴 긴급(≤7일)", "🟠 주의(≤14일)"], "🟡 관찰")
    o.insert(0, "상태", np.where(need, st_, ""))
    return o


def pick(o):
    s = o[o["_need"]]
    s = s.sort_values(["_key", "_short"], ascending=[False, False], kind="stable")
    return s.drop(columns=[c for c in s.columns if c.startswith("_")]).reset_index(drop=True)


def _dw(v):
    return sum(2 if ord(c) > 0x2E80 else 1 for c in str(v))


HEAD_BREAK = [("입고예정일", "입고\n예정일"), ("입고예정", "입고\n예정"), ("3개월평균(참고)", "3개월평균\n(참고)"),
              ("작년직전4주판매", "작년직전\n4주판매"), ("작년이후4주판매", "작년이후\n4주판매"), ("직전4주판매", "직전\n4주판매"),
              ("4주예상판매", "4주예상\n판매"), ("추천생산량", "추천\n생산량"), ("소진예상일", "소진\n예상일"),
              ("소비기한(최단)", "소비기한\n(최단)"), ("소비기한(재고별)", "소비기한\n(재고별)"), ("직전입고일", "직전\n입고일"),
              ("직전입고량", "직전\n입고량"), ("마트예상(보정)", "마트예상\n(보정)"), ("온라인예상", "온라인\n예상"),
              ("합계재고", "합계\n재고"), ("온라인재고", "온라인\n재고"), ("마트재고", "마트\n재고"), ("마트보정", "마트\n보정"),
              ("남은일수", "남은\n일수"), ("] ", "]\n"), ("마트·도매 ", "마트·도매\n"), ("온라인 ", "온라인\n")]


def _wrap_head(h):
    h = str(h)
    for a, b in HEAD_BREAK:
        h = h.replace(a, b)
    return h


def warn_style(df, ratio):
    if not ratio or df.empty:
        return df
    sty = pd.DataFrame("", index=df.index, columns=df.columns)
    for c in df.columns:
        if str(c).endswith("추천생산량"):
            e = str(c)[:-len("추천생산량")] + "4주예상판매"
            if e in df.columns:
                m = (df[c] > df[e] * ratio) & (df[c] > 0)
                sty.loc[m, c] = "background-color: #FFFF00; color: #000000"
    return df.style.apply(lambda _: sty, axis=None)


def to_excel(sheets, params, warn=None):
    import math
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.properties import PageSetupProperties

    def clean(df):
        df = df.copy()
        if "상태" in df:
            df["상태"] = df["상태"].str.replace(r"[🔴🟠🟡]\s*", "", regex=True)
        if "소비기한(재고별)" in df:
            df["소비기한(재고별)"] = df["소비기한(재고별)"].astype(str).str.replace(" / ", "\n", regex=False)
        return df

    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as w:
        for name, df in sheets.items():
            clean(df).to_excel(w, sheet_name=name, index=False)
        pd.DataFrame(params, columns=["항목", "값"]).to_excel(w, sheet_name="산출기준", index=False)
        fills = {"품절": "F8CBAD", "긴급": "F8CBAD", "주의": "FFE699", "관찰": "FFF2CC"}
        for ws in w.book.worksheets:
            orig = [str(c.value) for c in ws[1]]
            for c in ws[1]:
                c.value = _wrap_head(c.value)
                c.font = Font(bold=True, color="FFFFFF", size=10)
                c.fill = PatternFill("solid", fgColor="1F4E78")
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.row_dimensions[1].height = 15 * max((str(c.value).count("\n") + 1 for c in ws[1]), default=1) + 3
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for i, col in enumerate(ws.columns, 1):
                head = str(col[0].value)
                data_w = max((max(_dw(t) for t in str(c.value).split("\n")) for c in col[1:] if c.value is not None), default=0)
                head_w = max(_dw(t) for t in head.replace("\n", " ").split())
                width = max(data_w, head_w) + 2
                if head == "상품명":
                    width = min(width, 30)
                ws.column_dimensions[get_column_letter(i)].width = max(width, 6)
            for row in ws.iter_rows(min_row=2):
                lines = 1
                for c in row:
                    c.font = Font(size=10)
                    wrap = isinstance(c.value, str) and ("\n" in c.value or (ws.cell(1, c.column).value == "상품명" and _dw(c.value) > ws.column_dimensions[get_column_letter(c.column)].width - 2))
                    c.alignment = Alignment(vertical="center", wrap_text=bool(wrap))
                    if isinstance(c.value, str):
                        lines = max(lines, c.value.count("\n") + 1)
                        if wrap and "\n" not in c.value:
                            lines = max(lines, math.ceil(_dw(c.value) / (ws.column_dimensions[get_column_letter(c.column)].width - 2)))
                if lines > 1:
                    ws.row_dimensions[row[0].row].height = 14 * lines
            if ws.title.startswith("생산필요") and ws.max_row > 1:
                for row in ws.iter_rows(min_row=2):
                    key = str(row[0].value or "")[:2]
                    if key in fills:
                        row[0].fill = PatternFill("solid", fgColor=fills[key])
            for i, h in enumerate(orig, 1):
                tint = "DCEEFB" if h == "합계재고" else ("FCE4EC" if h.endswith("4주예상판매") else None)
                if tint:
                    for r in range(1, ws.max_row + 1):
                        ws.cell(r, i).fill = PatternFill("solid", fgColor=tint)
                    ws.cell(1, i).font = Font(bold=True, color="1F2937", size=10)
            if warn:
                ic = {h: i for i, h in enumerate(orig, 1)}
                for h, i in ic.items():
                    e = ic.get(h[:-len("추천생산량")] + "4주예상판매") if h.endswith("추천생산량") else None
                    if e:
                        for r in range(2, ws.max_row + 1):
                            rv, ev = ws.cell(r, i).value, ws.cell(r, e).value
                            if isinstance(rv, (int, float)) and isinstance(ev, (int, float)) and rv > 0 and rv > ev * warn:
                                ws.cell(r, i).fill = PatternFill("solid", fgColor="FFFF00")
            ws.page_setup.orientation = "landscape"
            ws.page_setup.paperSize = 9
            ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            ws.print_title_rows = "1:1"
            ws.page_margins.left = ws.page_margins.right = 0.4
            ws.page_margins.top = ws.page_margins.bottom = 0.5
            ws.print_options.horizontalCentered = True
    return bio.getvalue()


def load_ov():
    try:
        with open(OV_PATH, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_ov(ov):
    try:
        with open(OV_PATH, "w", encoding="utf-8") as f:
            json.dump(ov, f, ensure_ascii=False)
    except Exception as e:
        st.warning(f"보정값 저장 실패: {e}")


VIEWS = {"전체": "전체", "온라인(비앤씨)": "온라인", "마트·도매": "마트"}
MODES = {"계절보정": ["S"], "단순평균": ["A"], "둘 다 비교": ["S", "A"]}
drop_h = lambda df: df.drop(columns=[c for c in df.columns if c.startswith("_")])


@st.cache_data(show_spinner=False)
def base_date_of(data: bytes, sheet: str):
    try:
        raw = pd.read_excel(io.BytesIO(data), sheet_name=sheet, header=None)
        v = raw.iloc[3, detect_layout(raw)["date"]]
        return pd.to_datetime(str(v).split("~")[-1].strip()).date()
    except Exception:
        return None


def week_key(d):
    iso = d.isocalendar()
    return f"{iso[0]}-W{iso[1]:02d}"


# ★ 구글 시트 및 로컬 저장 병행 로직
def save_week(data, fname, sheet, bdate, be, cfg, ov, dc=None):
    key = week_key(bdate)
    
    # 1. 로컬 저장 (서버 가동용)
    d = os.path.join(SAVE_DIR, key)
    os.makedirs(d, exist_ok=True)
    with open(os.path.join(d, "source.xlsx"), "wb") as f:
        f.write(data)
    be.to_csv(os.path.join(d, "estimates.csv"), index=False, encoding="utf-8-sig")
    act = be[be["mark"].isna()]
    summ = {m: {v: int(len(pick(build(act, v, ms, cfg["thr"], cfg.get("qty"))))) for v in ("전체", "온라인", "마트")}
            for m, ms in MODES.items()}
    if dc and os.path.exists(DC_PATH):
        with open(DC_PATH, "rb") as f, open(os.path.join(d, "discontinued_list.xlsx"), "wb") as g:
            g.write(f.read())
    meta = dict(week=key, base_date=str(bdate), saved_at=datetime.now().strftime("%Y-%m-%d %H:%M"),
                file_name=fname, sheet=sheet, **cfg, mart_overrides=ov, summary=summ,
                dc_file=dc["file_name"] if dc else None, dc_saved_at=dc["saved_at"] if dc else None,
                dc_counts=dict(단종=dc["n11"], 단종예정=dc["n22"]) if dc else None)
    with open(os.path.join(d, "settings.json"), "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=1)

    # 2. 구글 시트로 직접 저장 (DB 연동)
    conn = get_gsheets_connection()
    if conn is not None:
        try:
            # 설정을 시트에 업데이트 (settings 탭)
            settings_df = pd.DataFrame([{
                "week": key, "base_date": str(bdate), "saved_at": meta["saved_at"],
                "file_name": fname, "sheet": sheet, "cap": cfg.get("cap"),
                "min_base": cfg.get("min_base"), "thr": cfg.get("thr"), "qty": cfg.get("qty")
            }])
            conn.update(worksheet="settings", data=settings_df)

            # 계산된 예측 데이터(be)를 구글 시트 estimates 탭에 저장
            be_save = be.copy()
            be_save["week"] = key
            conn.update(worksheet="estimates", data=be_save.fillna(""))
            st.toast("구글 시트 저장 성공! 📊")
        except Exception as e:
            st.warning(f"구글 시트 연동 저장 중 일부 문제가 발생했습니다: {e}")

    return key


def read_meta(key):
    try:
        with open(os.path.join(SAVE_DIR, key, "settings.json"), encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def list_saved():
    if not os.path.isdir(SAVE_DIR):
        return []
    return [m for m in (read_meta(k) for k in sorted(os.listdir(SAVE_DIR), reverse=True)) if m]


def load_saved(key):
    be = pd.read_csv(os.path.join(SAVE_DIR, key, "estimates.csv"),
                     dtype={"code": str, "in_date": str, "exp_min": str, "detail": str, "mgr": str,
                            "inc_date": str, "odate": str, "eta": str})
    for c in ("in_date", "exp_min", "detail", "inc_date", "odate", "eta"):
        if c in be:
            be[c] = be[c].fillna("")
    if "mark" not in be:
        be["mark"] = np.nan
    return be


def compare(sbe, cur, view, meth, thr, qty=None):
    act = cur[["code", "on4", "mt4"]].rename(columns={"on4": "on_act", "mt4": "mt_act"})
    m = sbe.merge(act, on="code", how="inner")
    stock = {"전체": m["on_stock"] + m["mart_stock"], "온라인": m["on_stock"], "마트": m["mart_stock"]}[view]
    actual = {"전체": m["on_act"] + m["mt_act"], "온라인": m["on_act"], "마트": m["mt_act"]}[view]
    est = {k: {"전체": m[f"on_{k}"] + m[f"mt_{k}"], "온라인": m[f"on_{k}"], "마트": m[f"mt_{k}"]}[view] for k in "SA"}
    r = pd.DataFrame({"상품코드": m["code"], "상품명": m["name"],
                      "저장시 재고": stock, "저장시 직전입고량": m["in_qty"], "실제 판매(직전4주)": actual})
    for k in "SA":
        e = est[k] - actual
        r[f"예상({METH[k]})"] = est[k].round(1)
        r[f"오차({METH[k]})"] = e.round(1)
        r[f"오차율({METH[k]})"] = np.where(actual > 0, (e / actual.where(actual > 0, 1) * 100).round(0), np.nan)
    need = (est[meth] > 0) & ((stock < est[meth] * thr) if thr else True)
    if qty:
        need = need & (est[meth] > qty)
    short = actual > stock
    r["판정"] = np.select([need & short, need & ~short, ~need & short], ["✅ 적중", "⚠️ 과다추천", "❌ 누락"], "정상")
    r["_abs"] = (est[meth] - actual).abs()
    stat = {}
    for k in "SA":
        tot = float(actual.sum())
        stat[k] = dict(wape=float((est[k] - actual).abs().sum() / tot) if tot else np.nan,
                       bias=float((est[k].sum() - tot) / tot) if tot else np.nan)
    return r, stat


def plan_tab(up, dc, mode, cap, min_base, thr, qty, apply_dc, inc_on, inc_days, warn):
    if not up:
        st.info("왼쪽에서 엑셀을 업로드하세요. (제조도우미_테스트 시트 기준)")
        return None
    data = up.getvalue()
    names = sheet_names(data)
    default = next((i for i, n in enumerate(names) if n == "(뉴)제조도우미"),
                   next((i for i, n in enumerate(names) if n == "제조도우미_테스트"), 0))
    sheet = st.selectbox("시트", names, index=default)
    b, bad = load_base(data, sheet)
    if bad:
        st.error("이 시트에서 필요한 열을 찾지 못했습니다: " + ", ".join(bad) + " — 시트를 확인하세요. (신규 양식은 '(뉴)제조도우미' 시트)")
        return None

    bdate0 = base_date_of(data, sheet) or datetime.now().date()
    b, inc_stat = apply_incoming(b, bdate0, inc_days, inc_on)
    if inc_on:
        st.caption(f"입고예정 반영(기준일 {bdate0} +{inc_days}일 이내): **{inc_stat['valid']}건** 반영 · 입고예정일 경과/미입고 확인 불가 {inc_stat['expired']}건 제외 · 발주정보 {inc_stat['total']}건")
    dc_on = dc if (dc and apply_dc) else None
    b["mark"] = pd.to_numeric(b["code"].map(dc_on["marks"]), errors="coerce") if dc_on else np.nan
    if not dc:
        st.warning("단종·단종예정 목록이 없습니다. 왼쪽에서 목록 엑셀을 올려야 단종 품목이 제외됩니다.")
    elif not apply_dc:
        st.warning("단종·단종예정 제외가 꺼져 있습니다 (사이드바 체크박스).")

    fid = (up.name, up.size, sheet)
    if st.session_state.get("fid") != fid:
        st.session_state["fid"], st.session_state["ov0"] = fid, load_ov()
        st.session_state.pop("mart_editor", None)
    ov0 = st.session_state["ov0"]
    with st.expander("🛒 마트·도매 1회성 출고 보정 (체크 = 4주 예상에 반영)", expanded=False):
        st.caption("기본값: 작년 동기(직전4주/이후4주)에 마트·도매 판매가 있던 품목만 반영. 직접 바꾸면 코드별로 저장되어 다음 파일에도 유지됩니다.")
        mt = b[(b["mt4"] > 0) & b["mark"].isna()].sort_values("mt4", ascending=False)
        ed_in = pd.DataFrame({"반영": [bool(ov0.get(c, a)) for c, a in zip(mt["code"], mt["auto"])],
                              "상품코드": mt["code"], "상품명": mt["name"], "마트·도매 직전4주": mt["mt4"],
                              "작년 직전4주": mt["mt_ly4"], "작년 이후4주": mt["mt_lyn4"],
                              "자동판정": np.where(mt["auto"], "반영", "제외")})
        ed = st.data_editor(ed_in, hide_index=True, key="mart_editor", height=320, use_container_width=True,
                            disabled=[c for c in ed_in.columns if c != "반영"])
    flags = dict(zip(b["code"], b["auto"]))
    flags.update({k: bool(v) for k, v in ov0.items()})
    flags.update({c: bool(v) for c, v in zip(ed["상품코드"], ed["반영"])})
    ov = dict(ov0)
    auto = dict(zip(b["code"], b["auto"]))
    for c, v in zip(ed["상품코드"], ed["반영"]):
        if bool(v) == bool(auto[c]):
            ov.pop(c, None)
        else:
            ov[c] = bool(v)
    if ov != st.session_state.get("ov_saved", ov0):
        save_ov(ov)
    st.session_state["ov_saved"] = ov

    be = estimates(b, flags, cap, min_base)
    meths = MODES[mode]
    params = [("엑셀 시트", sheet), ("판매량 기준", mode), ("단순평균", "직전 4주 실적"),
              ("입고예정 반영", f"입고예정일 기준일+{inc_days}일 이내·미입고 발주수량을 온라인·합계 재고에 가산 (마트 재고 미반영)" if inc_on else "미적용"),
              ("추천생산량 주의(노란색)", f"추천생산량 > 4주예상판매 × {warn}배 — 직전입고량 단위 올림 때문에 과생산될 수 있음(마트 일회성 출고로 직전입고가 큰 경우 등). 실무자가 확인 후 생산량 결정" if warn else "미적용"),
              ("3개월평균(참고)", "월별 1~3개월 전 판매(전채널) 평균을 4주로 환산한 참고값 — 판정에 미사용"),
              ("계절 배율 상한", cap or "없음"),
              ("계절보정 최소 기준량", f"작년 직전4주 < {min_base}이면 단순평균 사용" if min_base else "미적용"),
              ("소진 기준", f"재고 ÷ 4주예상 < {thr}" if thr else "미적용 (4주 예상이 있는 품목 모두)"),
              ("4주 예상판매량 필터", f"선택 보기의 4주 예상판매 > {qty}개" if qty else "미적용"),
              ("생성일시", datetime.now().strftime("%Y-%m-%d %H:%M")),
              ("계절보정", "4주예상 = 직전4주 × (작년 이후4주 ÷ 작년 직전4주), 작년 직전4주=0이면 보정 없음 (최소 기준량 적용 시 단순평균)"),
              ("온라인", "비앤씨마켓"), ("마트·도매", "전채널 − 비앤씨마켓 (마트 보정 반영/제외 적용)"),
              ("단종 제외", f"업로드한 단종 목록 기준 11(단종)·22(단종예정) 제외 — {dc['file_name']} (단종 {dc['n11']}, 단종예정 {dc['n22']})" if dc_on else "미적용"),
              ("재고", "같은 상품코드는 소비기한별 행 합산 (소비기한 임박 재고 포함)"),
              ("정렬", "4주 예상 부족분 많은 순 (전체 보기: 온라인(비앤씨) 부족분 우선 → 전체 부족분)"),
              ("추천생산량", "부족분(4주예상−재고)을 직전입고량 단위로 올림(최소 1회분), 입고이력 없으면 부족분 그대로")]

    bdate = base_date_of(data, sheet) or datetime.now().date()
    key = week_key(bdate)
    old = read_meta(key)
    s1, s2 = st.columns([3, 1])
    s1.caption(f"저장 주차: **{key}** (파일 기준일 {bdate})" + (f" — ⚠️ 이미 저장됨({old['saved_at']}), 저장하면 덮어씁니다." if old else ""))
    if s2.button("💾 이번 주 저장", use_container_width=True):
        save_week(data, up.name, sheet, bdate, be, dict(avg_mode="직전 4주 실적", cap=cap, min_base=min_base, thr=thr, qty=qty, inc_on=inc_on, inc_days=inc_days), ov, dc_on)
        st.success(f"{key} 저장 완료 — 구글 시트에 업데이트되었습니다.")

    label = st.radio("보기 (채널 선택 — 어느 채널이든 전체·온라인·마트 재고를 함께 표시)", list(VIEWS), horizontal=True)
    view = VIEWS[label]
    be_ex, be_a = be[be["mark"].notna()], be[be["mark"].isna()]
    res = {v: build(be_a, v, meths, thr, qty) for v in VIEWS.values()}
    res_ex = {v: pick(build(be_ex, v, meths, thr, qty)) for v in VIEWS.values()} if len(be_ex) else {v: pd.DataFrame() for v in VIEWS.values()}
    o = res[view]
    show = pick(o)
    c1, c2, c3 = st.columns(3)
    c1.metric("생산 필요 품목", len(show))
    c2.metric("🔴 품절·긴급", int(show["상태"].str.startswith("🔴").sum()) if len(show) else 0)
    c3.metric("단종·단종예정 제외", len(be_ex), help=f"제외 품목 중 이 보기에서 생산필요였던 품목: {len(res_ex[view])}개 (아래 목록에서 확인)")
    st.caption("정렬: 4주 예상 부족분이 많은 순" + (" (전체 보기는 온라인(비앤씨) 부족분 우선 → 전체 부족분 순)" if view == "전체" else ""))
    if warn and len(show):
        nw = sum(int(((show[c] > show[c[:-5] + "4주예상판매"] * warn) & (show[c] > 0)).sum())
                 for c in show.columns if c.endswith("추천생산량") and c[:-5] + "4주예상판매" in show.columns)
        st.caption(f"🟨 노란색 = 추천생산량이 4주예상판매의 {warn}배 초과 ({nw}건) — 직전입고량이 마트 일회성 출고로 커진 경우 등 과생산 가능성이 있어 실무자 확인이 필요해요.")
    st.dataframe(warn_style(show, warn), hide_index=True, use_container_width=True, height=620)
    with st.expander(f"🚫 제외된 단종·단종예정 품목 중 생산필요였던 품목 ({len(res_ex[view])}개 — 생산계획에는 포함되지 않음)"):
        st.dataframe(warn_style(res_ex[view], warn), hide_index=True, use_container_width=True)

    d1, d2 = st.columns(2)
    d1.download_button(f"📥 {label} 생산필요 엑셀", to_excel(
        {f"생산필요_{label[:2]}": show, "전체품목": drop_h(o), "제외_단종": res_ex[view]}, params + [("보기", label)], warn),
        file_name=f"생산필요_{label[:2]}_{datetime.now():%Y%m%d}.xlsx", key=f"dl_{view}",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    d2.download_button("📥 채널 전체 엑셀 (전체·온라인·마트 시트)", to_excel(
        {f"생산필요_{k[:2]}": pick(res[v]) for k, v in VIEWS.items()}, params, warn),
        file_name=f"생산필요_채널별_{datetime.now():%Y%m%d}.xlsx", key="dl_all",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    return dict(b=b, bdate=bdate)


def history_tab(ctx, warn=None):
    saved = list_saved()
    if not saved:
        st.info("저장된 주차가 없습니다. '생산계획' 탭에서 [이번 주 저장]을 눌러 저장하세요.")
        return
    st.caption(f"저장 위치: {SAVE_DIR}")
    hist = pd.DataFrame([{"주차": m["week"], "기준일": m["base_date"], "저장일시": m["saved_at"], "파일명": m["file_name"],
                          "계절보정 필요(전체/온라인/마트)": "/".join(str(m["summary"]["계절보정"][v]) for v in ("전체", "온라인", "마트")),
                          "단순평균 필요(전체/온라인/마트)": "/".join(str(m["summary"]["단순평균"][v]) for v in ("전체", "온라인", "마트")),
                          "상한": m["cap"], "최소기준량": m["min_base"]} for m in saved])
    st.dataframe(hist, hide_index=True, use_container_width=True)

    key = st.selectbox("주차 선택", [m["week"] for m in saved])
    meta = read_meta(key)
    sbe = load_saved(key)
    n_ex = int(sbe["mark"].notna().sum())
    sbe = sbe[sbe["mark"].isna()].reset_index(drop=True)
    with st.expander("저장 시 설정 / 원본 파일"):
        st.json({k: v for k, v in meta.items() if k not in ("summary", "mart_overrides")})
        with open(os.path.join(SAVE_DIR, key, "source.xlsx"), "rb") as f:
            st.download_button("📥 저장된 구매도우미 원본 엑셀", f.read(), file_name=f"{key}_{meta['file_name']}",
                               key="dl_src", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.subheader(f"{key} 생산추천 조회")
    f1, f2, f3, f4 = st.columns([2, 2, 2, 3])
    mode = f1.radio("판매량 기준", list(MODES), key="h_mode")
    label = f2.radio("보기", list(VIEWS), key="h_view")
    scope = f3.radio("범위", ["생산필요만", "전체 품목"], key="h_scope")
    q = f4.text_input("품목 검색 (상품명/코드)", key="h_q", help="검색어를 입력하면 생산필요 여부와 관계없이 전체 품목에서 찾습니다.")
    view = VIEWS[label]
    o = build(sbe, view, MODES[mode], meta["thr"], meta.get("qty"))
    if q.strip():
        t = drop_h(o.sort_values("_key", ascending=False))
        t = t[t["상품명"].str.contains(q.strip(), case=False, na=False, regex=False) | t["상품코드"].str.contains(q.strip(), case=False, na=False, regex=False)]
    elif scope == "전체 품목":
        t = drop_h(o.sort_values("_key", ascending=False))
    else:
        t = pick(o)
    st.caption(f"{len(t)}개 품목 · 저장 시 단종·단종예정 {n_ex}개 제외 (목록: {meta.get('dc_file') or '없음'})")
    st.dataframe(warn_style(t.reset_index(drop=True), warn), hide_index=True, use_container_width=True, height=480)
    params = [("주차", key), ("기준일", meta["base_date"]), ("저장일시", meta["saved_at"]), ("판매량 기준", mode),
              ("단순평균", meta.get("avg_mode", "직전 4주 실적")), ("계절 배율 상한", meta["cap"] or "없음"),
              ("계절보정 최소 기준량", meta["min_base"]), ("소진 기준", f"재고 ÷ 4주예상 < {meta['thr']}" if meta["thr"] else "미적용"),
              ("4주 예상판매량 필터", f"> {meta['qty']}개" if meta.get("qty") else "미적용")]
    st.download_button(f"📥 {key} {label} 엑셀", to_excel({f"생산필요_{label[:2]}": t.reset_index(drop=True)}, params, warn),
                       file_name=f"{key}_생산필요_{label[:2]}.xlsx", key="dl_hist",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.subheader("🔍 예상 검증 (저장된 4주 예상 vs 현재 업로드 파일의 실제 직전4주 판매)")
    if not ctx:
        st.info("검증하려면 왼쪽에서 이후 주차의 엑셀(저장 기준일로부터 약 4주 뒤)을 업로드하세요.")
        return
    gap = (ctx["bdate"] - pd.to_datetime(meta["base_date"]).date()).days
    if gap <= 0:
        st.warning(f"현재 파일 기준일({ctx['bdate']})이 저장 기준일({meta['base_date']})보다 앞서거나 같아 검증할 수 없습니다.")
        return
    msg = f"저장 기준일 {meta['base_date']} → 현재 파일 기준일 {ctx['bdate']} ({gap}일 후)."
    if 21 <= gap <= 35:
        st.caption(msg + " 현재 파일의 직전 4주가 저장 예상 구간과 거의 겹칩니다.")
    else:
        st.warning(msg + " 4주 후(약 28일)가 아니라서 직전 4주 구간이 예상 구간과 어긋납니다. 참고용으로만 보세요.")
    v1, v2, v3 = st.columns(3)
    vlabel = v1.radio("검증 보기", list(VIEWS), key="v_view", horizontal=True)
    vmeth = v2.radio("판정 기준", ["계절보정", "단순평균"], key="v_meth", horizontal=True)
    vf = v3.selectbox("표시", ["정상 제외", "✅ 적중", "⚠️ 과다추천", "❌ 누락", "전체"], key="v_filter")
    r, stat = compare(sbe, ctx["b"], VIEWS[vlabel], "S" if vmeth == "계절보정" else "A", meta["thr"], meta.get("qty"))
    cnt = r["판정"].value_counts()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("✅ 적중 (필요·실제 부족)", int(cnt.get("✅ 적중", 0)))
    m2.metric("⚠️ 과다추천 (필요·실제 충분)", int(cnt.get("⚠️ 과다추천", 0)))
    m3.metric("❌ 누락 (불필요·실제 부족)", int(cnt.get("❌ 누락", 0)))
    m4.metric("비교 품목", len(r))
    pct = lambda x: "-" if pd.isna(x) else f"{x*100:.0f}%"
    st.caption(f"전체 오차(WAPE, 낮을수록 좋음) — 계절보정 {pct(stat['S']['wape'])} / 단순평균 {pct(stat['A']['wape'])}  |  "
               f"편향(+는 과대예측) — 계절보정 {pct(stat['S']['bias'])} / 단순평균 {pct(stat['A']['bias'])}  |  "
               "※ 실제 부족은 '실제 판매 > 저장 당시 재고'이며 이후 입고·생산분은 반영하지 않았습니다.")
    t = r if vf == "전체" else (r[r["판정"] != "정상"] if vf == "정상 제외" else r[r["판정"] == vf])
    t = t.sort_values("_abs", ascending=False)
    st.dataframe(drop_h(t).reset_index(drop=True), hide_index=True, use_container_width=True, height=480)
    st.download_button("📥 검증 결과 엑셀", to_excel({"검증": drop_h(t).reset_index(drop=True)},
                       [("저장 주차", key), ("현재 파일 기준일", str(ctx["bdate"])), ("판정 기준", vmeth), ("보기", vlabel)]),
                       file_name=f"{key}_검증_{vlabel[:2]}.xlsx", key="dl_val",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def main():
    st.set_page_config(page_title="BGF 생산계획 도우미", layout="wide")
    st.title("🏭 BGF 생산계획 도우미")
    with st.sidebar:
        up = st.file_uploader("구매도우미 12주판매기준 엑셀", type=["xlsx"])
        mode = st.radio("판매량 기준", list(MODES))
        st.markdown("**적용 기준 (체크한 것만 사용)**")
        use_cap = st.checkbox("계절 배율 상한", True)
        cap_v = st.number_input("배율 상한 (배)", 0.5, 50.0, 3.0, 0.5, disabled=not use_cap,
                                help="작년 이후4주/작년 직전4주 배율의 상한")
        use_min = st.checkbox("계절보정 최소 기준량", True)
        min_v = st.number_input("작년 직전4주 최소 판매량 (개)", 1, 10000, 30, 5, disabled=not use_min,
                                help="작년 직전4주 판매가 이 값 미만이면 계절보정 대신 단순평균 사용 (온라인·마트 각각 적용)")
        use_thr = st.checkbox("소진 기준 (재고 ÷ 4주예상 <)", True,
                              help="해제하면 소진 판정 없이 4주 예상판매가 있는 품목을 모두 표시")
        thr_v = st.number_input("소진 기준값", 0.1, 5.0, 1.0, 0.1, disabled=not use_thr)
        use_qty = st.checkbox("4주 예상판매량 필터", False,
                              help="선택한 보기의 4주 예상판매가 아래 수량을 초과하는 품목만 표시")
        qty_v = st.number_input("예상판매량 초과 수량 (개)", 0, 100000, 30, 10, disabled=not use_qty)
        apply_dc = st.checkbox("단종·단종예정 제외", True)
        use_inc = st.checkbox("입고예정 물량 재고에 반영", True,
                              help="미입고 발주 중 입고예정일이 기준일 이후 N일 이내인 수량을 온라인·합계 재고에 더해 계산 (마트 재고에는 미반영)")
        inc_days = st.number_input("입고예정 반영 기간 (기준일 +일)", 1, 120, 28, 1, disabled=not use_inc)
        use_warn = st.checkbox("추천생산량 주의 표시 (노란색)", True,
                               help="추천생산량이 4주예상판매의 N배를 넘으면 노란색으로 표시 (화면·엑셀). 직전입고량이 마트 일회성 출고로 커진 경우 등 확인용")
        warn_v = st.number_input("4주예상판매 대비 N배 초과", 1.0, 100.0, 2.0, 0.5, disabled=not use_warn)
        warn = warn_v if use_warn else None
        cap, min_base = (cap_v if use_cap else 0.0), (min_v if use_min else 0)
        thr, qty = (thr_v if use_thr else None), (qty_v if use_qty else None)
        st.divider()
        up_dc = st.file_uploader("단종·단종예정 목록 엑셀 (한 번 올리면 저장됨)", type=["xlsx"], key="dc")
    dc = load_dc(up_dc)
    with st.sidebar:
        if dc:
            st.caption(f"적용 중: {dc['file_name']} (등록 {dc['saved_at'] or '-'}) · 단종 {dc['n11']} · 단종예정 {dc['n22']}")
            for w in dc["warns"]:
                st.warning(w)
        else:
            st.caption("⚠️ 단종 목록 없음 — 제외 미적용")
    t1, t2 = st.tabs(["📊 생산계획", "🗂️ 저장 이력·검증"])
    with t1:
        ctx = plan_tab(up, dc, mode, cap, min_base, thr, qty, apply_dc, use_inc, int(inc_days), warn)
    with t2:
        history_tab(ctx, warn)


if __name__ == "__main__":
    main()
