# -*- coding: utf-8 -*-
"""
report_hwpx.py — 공사기간 산정 검토 보고서(한글 hwpx)

templates/report_template.hwpx(사용자 샘플 보고서에서 만든 템플릿, tools/make_report_template.py)의
스타일·고정 문구를 그대로 쓰고, 사업마다 바뀌는 값({{...}})을 채우고 표를 다시 만든다.
  산정 근거(고시 발췌) → 1. 공사기간 → 2. 준비기간 → 3. 비작업일수(법정공휴일·기상조건·월별 산정표)
  → 4. 작업일수(분야별 표) → 5. 정리기간 → 6. 적정성 검토(유사사업 표는 직접 작성)
숫자는 부록 엑셀과 같은 데이터(app.report_data)에서 온다 — 보고서와 부록이 서로 다를 수 없다.
"""

from __future__ import annotations

import io
import math
import os
import random
import zipfile
from typing import Any, Dict, List, Sequence
from xml.sax.saxutils import escape

TEMPLATE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates", "report_template.hwpx")

# 표 셀 스타일(템플릿 header.xml의 ID): 머리행·본문·계 행 테두리, 가운데 정렬 문단, 글자(굵게·보통)
BF_HEAD, BF_BODY, BF_SUM = 21, 36, 52
PP_CENTER = 20
CP_HEAD, CP_BODY = 216, 232
CP_SMALL, CP_SMALL_BOLD = 257, 258     # 템플릿에 추가한 9pt 표 글자(tools/make_report_template.py)
ROW_H = 1500                    # 행 높이(HWPUNIT). 글이 넘치면 한글이 늘린다
TABLE_W = 47320                 # 표 너비(본문 폭)

DISC_LABEL = {"건축": "건축", "건축기계설비": "건축기계설비", "기계": "기계",
              "전기": "전기 및 계측제어", "조경": "조경"}


def _esc(v) -> str:
    return escape("" if v is None else str(v))


def _num(v, nd=2) -> str:
    """3,203 · 4.73 · 0.5 처럼(정수면 소수점 없이)."""
    if v is None or v == "":
        return ""
    f = float(v)
    if abs(f - round(f)) < 1e-9:
        return f"{int(round(f)):,}"
    s = f"{f:,.{nd}f}".rstrip("0").rstrip(".")
    return s


def _round_half_up(x: float) -> int:
    return int(math.floor(x + 0.5))


# ── 표 ─────────────────────────────────────────────────────────────────
class Table:
    """병합(colspan·rowspan)을 지원하는 hwpx 표. add()는 (행, 열)에 셀을 놓는다."""

    def __init__(self, widths: Sequence[int], header_rows: int = 1, small: bool = True):
        self.w = list(widths)
        self.small = small
        self.cells: List[Dict[str, Any]] = []
        self.n_rows = 0
        self.header_rows = header_rows

    def add(self, row, col, text, colspan=1, rowspan=1, kind="body"):
        self.cells.append({"r": row, "c": col, "t": text, "cs": colspan, "rs": rowspan, "k": kind})
        self.n_rows = max(self.n_rows, row + rowspan)

    def _tc(self, c) -> str:
        bf = {"head": BF_HEAD, "body": BF_BODY, "sum": BF_SUM}[c["k"]]
        if self.small:
            cp = CP_SMALL if c["k"] == "body" else CP_SMALL_BOLD
        else:
            cp = CP_BODY if c["k"] == "body" else CP_HEAD
        w = sum(self.w[c["c"]:c["c"] + c["cs"]])
        h = ROW_H * c["rs"]
        hdr = 1 if c["r"] < self.header_rows else 0
        return (f'<hp:tc name="" header="{hdr}" hasMargin="0" protect="0" editable="0" dirty="0" '
                f'borderFillIDRef="{bf}"><hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" '
                'vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" '
                'hasTextRef="0" hasNumRef="0">'
                f'<hp:p id="2147483648" paraPrIDRef="{PP_CENTER}" styleIDRef="0" pageBreak="0" '
                f'columnBreak="0" merged="0"><hp:run charPrIDRef="{cp}"><hp:t>{_esc(c["t"])}</hp:t></hp:run>'
                f'</hp:p></hp:subList><hp:cellAddr colAddr="{c["c"]}" rowAddr="{c["r"]}"/>'
                f'<hp:cellSpan colSpan="{c["cs"]}" rowSpan="{c["rs"]}"/><hp:cellSz width="{w}" height="{h}"/>'
                '<hp:cellMargin left="4294967295" right="4294967295" top="4294967295" bottom="4294967295"/>'
                '</hp:tc>')

    def xml(self, z: int) -> str:
        rows = []
        for r in range(self.n_rows):
            cs = sorted((c for c in self.cells if c["r"] == r), key=lambda c: c["c"])
            rows.append("<hp:tr>" + "".join(self._tc(c) for c in cs) + "</hp:tr>")
        tid = random.randint(1_000_000_000, 2_000_000_000)
        return (f'<hp:tbl id="{tid}" zOrder="{z}" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" '
                'textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="1" '
                f'rowCnt="{self.n_rows}" colCnt="{len(self.w)}" cellSpacing="0" borderFillIDRef="6" noAdjust="0">'
                f'<hp:sz width="{sum(self.w)}" widthRelTo="ABSOLUTE" height="{ROW_H * self.n_rows}" '
                'heightRelTo="ABSOLUTE" protect="0"/><hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1" '
                'allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" '
                'horzAlign="LEFT" vertOffset="0" horzOffset="0"/><hp:outMargin left="0" right="0" top="0" '
                'bottom="0"/><hp:inMargin left="141" right="141" top="85" bottom="85"/>'
                + "".join(rows) + "</hp:tbl>")

    def validate(self):
        """모든 칸이 정확히 한 셀로 덮이는지(한글은 어긋난 표를 열지 못한다)."""
        grid = {}
        for c in self.cells:
            for r in range(c["r"], c["r"] + c["rs"]):
                for k in range(c["c"], c["c"] + c["cs"]):
                    assert (r, k) not in grid, f"겹친 셀 ({r},{k})"
                    grid[(r, k)] = 1
        for r in range(self.n_rows):
            for k in range(len(self.w)):
                assert (r, k) in grid, f"빈 칸 ({r},{k})"


def _wrap_table(tbl: Table, z: int) -> str:
    tbl.validate()
    return ('<hp:p id="2147483648" paraPrIDRef="216" styleIDRef="125" pageBreak="0" columnBreak="0" '
            f'merged="0"><hp:run charPrIDRef="13">{tbl.xml(z)}<hp:t/></hp:run></hp:p>')


def _para(text, cp=51, pp=62, page_break=False) -> str:
    return (f'<hp:p id="2147483648" paraPrIDRef="{pp}" styleIDRef="0" pageBreak="{1 if page_break else 0}" '
            f'columnBreak="0" merged="0"><hp:run charPrIDRef="{cp}"><hp:t>{_esc(text)}</hp:t></hp:run></hp:p>')


def _scale(widths: Sequence[int]) -> List[int]:
    tot = sum(widths)
    out = [int(w * TABLE_W / tot) for w in widths]
    out[-1] += TABLE_W - sum(out)
    return out


MONTHS = [f"{m}월" for m in range(1, 13)]


def holidays_table(years: Dict[int, List[int]]) -> Table:
    t = Table(_scale([5320] + [3200] * 12 + [3600]), header_rows=2)
    t.add(0, 0, "구분", rowspan=2, kind="head")
    t.add(0, 1, "월간 법정공휴일", colspan=12, kind="head")
    t.add(0, 13, "소계", rowspan=2, kind="head")
    for m in range(12):
        t.add(1, m + 1, MONTHS[m], kind="head")
    for i, y in enumerate(sorted(years)):
        r = 2 + i
        t.add(r, 0, f"{y}년")
        for m, v in enumerate(years[y]):
            t.add(r, m + 1, _num(v))
        t.add(r, 13, _num(sum(years[y])))
    return t


def conditions_table(station: str, conds) -> Table:
    t = Table(_scale([9320] + [2800] * 12 + [4400]), header_rows=2)
    t.add(0, 0, f"구분({station})" if station else "구분", rowspan=2, kind="head")
    t.add(0, 1, "월평균 기상데이터(일)", colspan=12, kind="head")
    t.add(0, 13, "소계", rowspan=2, kind="head")
    for m in range(12):
        t.add(1, m + 1, MONTHS[m], kind="head")
    r = 2
    tot = [0.0] * 12
    for label, vals in conds:
        t.add(r, 0, label)
        for m, v in enumerate(vals):
            t.add(r, m + 1, _num(v, 1))
            tot[m] += float(v)
        t.add(r, 13, _num(sum(vals), 1))
        r += 1
    t.add(r, 0, "계", kind="sum")
    for m in range(12):
        t.add(r, m + 1, _num(tot[m], 1), kind="sum")
    t.add(r, 13, _num(sum(tot), 1), kind="sum")
    return t


# 기상조건 이름 → 보고서 표의 네 갈래(혹서기·동절기·강우·바람)
_GROUPS = (("혹서기", ("혹서", "폭염", "33", "35")), ("동절기", ("동절", "한랭", "0℃", "적설", "영하")),
           ("강우", ("강우", "강수", "비")), ("바람", ("풍속", "바람", "강풍")))


def _group_of(label: str) -> str:
    for g, kws in _GROUPS:
        if any(k in label for k in kws):
            return g
    return ""


def nonwork_table(rows) -> Table:
    t = Table(_scale([3400, 3400, 3200, 4200, 3800, 3600, 3600, 3600, 3600, 4000, 5200, 5720]), header_rows=2)
    t.add(0, 0, "년/월", colspan=2, rowspan=2, kind="head")
    t.add(0, 2, "일수", rowspan=2, kind="head")
    t.add(0, 3, "법정공휴일(A)", rowspan=2, kind="head")
    t.add(0, 4, "기후여건으로 인한 작업불가일수", colspan=5, kind="head")
    t.add(0, 9, "중복일수(C)", rowspan=2, kind="head")
    t.add(0, 10, "비작업일수(A+B−C)", rowspan=2, kind="head")
    t.add(0, 11, "적용 비작업일수", rowspan=2, kind="head")
    for i, lab in enumerate(("계(B)", "혹서기", "동절기", "강우", "바람")):
        t.add(1, 4 + i, lab, kind="head")
    by_year: Dict[str, List[dict]] = {}
    for row in rows:
        by_year.setdefault(str(row["월"])[:4], []).append(row)
    r = 2
    sums = {k: 0.0 for k in ("일수", "B", "A", "혹서기", "동절기", "강우", "바람", "C", "비작업", "적용")}
    for y in sorted(by_year):
        ms = by_year[y]
        t.add(r, 0, f"{y[2:]}년", rowspan=len(ms))
        for row in ms:
            g = {k: 0.0 for k in ("혹서기", "동절기", "강우", "바람")}
            for lab, v in (row.get("cond") or {}).items():
                k = _group_of(lab)
                if k:
                    g[k] += float(v or 0)
            t.add(r, 1, f"{int(str(row['월'])[5:7])}월")
            vals = [row["대상일수"], row["B"], row["A"], g["혹서기"], g["동절기"], g["강우"], g["바람"],
                    row["C"], row["비작업"], row["적용"]]
            for k, v in zip(sums, vals):
                sums[k] += float(v)
            for i, v in enumerate(vals):
                t.add(r, 2 + i, _num(v, 1))
            r += 1
    t.add(r, 0, "계", colspan=2, kind="sum")
    for i, k in enumerate(sums):
        t.add(r, 2 + i, _num(sums[k], 1), kind="sum")
    return t


def civil_work_table(civil: Dict[str, Any]) -> Table:
    t = Table(_scale([3300, 5200, 16000, 5200, 3600, 4800, 3200, 6020]), header_rows=1)
    for c, lab in ((0, "공 종"), (2, "구 분"), (3, "수량"), (4, "단위"), (5, "1일 작업량"), (6, "조"), (7, "적용(일)")):
        t.add(0, c, lab, colspan=2 if c == 0 else 1, kind="head")
    r = 1
    cat_days = []
    for cat in civil.get("cats", []):
        lines = [ln for ln in cat["lines"] if ln["items"]]
        n_items = sum(len(ln["items"]) for ln in lines)
        multi = len(lines) > 1
        t.add(r, 0, cat["name"], colspan=1 if multi else 2, rowspan=n_items + 1)
        line_days = []
        for ln in lines:
            if multi:
                t.add(r, 1, ln["name"], rowspan=len(ln["items"]))
            for it in ln["items"]:
                t.add(r, 2, f"{it['name']} {it.get('spec', '')}".strip())
                t.add(r, 3, _num(it["qty"], 3))
                t.add(r, 4, it.get("unit", ""))
                t.add(r, 5, _num(it["daily"], 2))
                t.add(r, 6, _num(it["crew"]))
                t.add(r, 7, _num(it["days"]))
                r += 1
            line_days.append(sum(i["days"] for i in ln["items"]))
        cd = max(line_days) if line_days else 0
        cat_days.append(cd)
        if multi:
            t.add(r, 1, f"소계(구간 {len(lines)}개 중 최장)", colspan=2, kind="sum")
        else:
            t.add(r, 2, "소계", kind="sum")
        for c in range(3, 7):
            t.add(r, c, "", kind="sum")
        t.add(r, 7, _num(cd), kind="sum")
        r += 1
    total = sum(cat_days) if civil.get("combine") == "sum" else max(cat_days, default=0)
    t.add(r, 0, "토목공사 작업일수(" + ("대공종 합" if civil.get("combine") == "sum" else "대공종 중 최장") + ")",
          colspan=7, kind="sum")
    t.add(r, 7, _num(total), kind="sum")
    return t


def disc_work_table(disc: Dict[str, Any]) -> Table:
    t = Table(_scale([9000, 12500, 9500, 7000, 3500, 5820]), header_rows=1)
    for c, lab in enumerate(("동(설비)", "공 종", "주 직종(병목)", "작업량(인·일)", "조", "적용(일)")):
        t.add(0, c, lab, kind="head")
    r = 1
    unit_days = []
    for u in disc.get("units", []):
        pk = u["packages"]
        t.add(r, 0, u["name"], rowspan=len(pk) + 1)
        for p in pk:
            load = sum(i["qty"] * i["rate"] for i in p["items"])
            t.add(r, 1, p["group"])
            t.add(r, 2, "노무비 역산" if p["lead"] == "노무비역산" else p["lead"])
            t.add(r, 3, _num(load, 1))
            t.add(r, 4, _num(p["crew"]))
            t.add(r, 5, _num(p["days"]))
            r += 1
        ud = sum(p["days"] for p in pk)
        unit_days.append(ud)
        t.add(r, 1, "소계(공종 순차)", colspan=4, kind="sum")
        t.add(r, 5, _num(ud), kind="sum")
        r += 1
    par = disc.get("mode") != "seq"
    total = max(unit_days, default=0) if par else sum(unit_days)
    lab = DISC_LABEL.get(disc["name"], disc["name"])
    t.add(r, 0, f"{lab}공사 작업일수(" + ("동 중 최장" if par else "동 합") + ")", colspan=5, kind="sum")
    t.add(r, 5, _num(total), kind="sum")
    return t


def similar_table() -> Table:
    t = Table(_scale([5000, 12000, 9000, 8000, 4500, 4820, 4000]), header_rows=1, small=False)
    for c, lab in enumerate(("구분", "사업명", "시설규모", "공사기간", "공사개월", "공사비(백만원)", "비 고")):
        t.add(0, c, lab, kind="head")
    for r in range(1, 4):
        for c in range(7):
            t.add(r, c, "")
    return t


# ── 본체 ────────────────────────────────────────────────────────────────
def build_report_hwpx(data: Dict[str, Any], template: str = TEMPLATE) -> bytes:
    """앱 산정 결과(data — 부록 엑셀과 같은 구조 + summary)로 보고서 hwpx를 만든다(바이트)."""
    zin = zipfile.ZipFile(template)
    sec = zin.read("Contents/section0.xml").decode("utf-8")
    per, sm = data["periods"], data["summary"]
    total = int(sm["total"])
    months = _round_half_up(total / 365 * 12)
    discs = data.get("discs", [])
    used = ["토목공사"] + [DISC_LABEL.get(d["name"], d["name"]) + "공사" for d in discs if d.get("use", True)]
    others = [DISC_LABEL.get(d["name"], d["name"]) + "공사" for d in discs if not d.get("use", True)]
    formula = (f"{per['prep']:,}일 (준비기간) + {int(sm['non_work']):,}일 (비작업일수) + "
               f"{int(sm['work']):,}일 (작업일수)"
               + (f" + {per['commission']:,}일 (시운전)" if per.get("commission") else "")
               + f" + {per['wrapup']:,}일 (정리기간)")
    station = data["nonwork"].get("station") or "-"
    texts = {
        "TITLE": data.get("title") or "공사",
        "FORMULA": formula,
        "TOTAL_DAYS": f"{total:,}",
        "TOTAL_MONTHS": f"{months}",
        "PREP_LABEL": per.get("prep_label") or "-",
        "PREP_DAYS": f"{per['prep']}",
        "PREP_NOTE": "가이드라인 공사 유형별 준비기간",
        "DISC_CP": ", ".join(used),
        "DISC_ALL": ", ".join(used + others),
        "DISC_OTHER": (f", {'·'.join(used)} 기간 내 {', '.join(others)}를 완료하는 것으로 계획." if others else "."),
        "STATION": station,
        "WRAPUP": f"{per['wrapup']}",
    }
    z = [100]

    def tbl(t: Table) -> str:
        z[0] += 1
        return _wrap_table(t, z[0])

    blocks = {
        "T_HOLIDAYS": tbl(holidays_table(data.get("holiday_years", {}))),
        "T_CONDITIONS": tbl(conditions_table(station, data["nonwork"].get("conditions", []))),
        "T_NONWORK": tbl(nonwork_table(data["nonwork"].get("rows", []))),
        "T_SIMILAR": tbl(similar_table()),
    }
    work = [_para("  토목공사 작업일수", cp=97, pp=201, page_break=True), tbl(civil_work_table(data["civil"])),
            _para(" ※ 1일 작업량 근거와 항목별 산출은 부록(작업일수 산정근거) 참조. 구간·구조물 안은 순차, "
                  "구간끼리는 병행")]
    for d in discs:
        if not d.get("units"):
            continue
        lab = DISC_LABEL.get(d["name"], d["name"])
        work += [_para(f"- {lab}공사 작업일수" + ("" if d.get("use", True) else " (공기 미반영·참고)"),
                       cp=97, page_break=True),
                 tbl(disc_work_table(d)),
                 _para(" ※ 공종 작업일수 = 가장 많이 드는 직종(병목)의 작업량 ÷ 조(1조 = 직종별 1인). "
                       "항목별 산출은 부록 참조")]
    blocks["WORK_TABLES"] = "".join(work)
    for k, v in blocks.items():
        assert "{{" + k + "}}" in sec, k
        sec = sec.replace("{{" + k + "}}", v)
    for k, v in texts.items():
        sec = sec.replace("{{" + k + "}}", _esc(v))
    assert "{{" not in sec, "채우지 못한 자리표시가 있습니다"

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as zout:
        for info in zin.infolist():
            n = info.filename
            d = zin.read(n)
            if n == "Contents/section0.xml":
                d = sec.encode("utf-8")
            elif n in ("Contents/content.hpf", "Preview/PrvText.txt"):
                d = d.decode("utf-8").replace("{{TITLE}}", _esc(texts["TITLE"])).encode("utf-8")
            zout.writestr(n, d, compress_type=info.compress_type)
    return out.getvalue()
