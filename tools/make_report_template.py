# -*- coding: utf-8 -*-
"""사용자 샘플 보고서(hwpx)에서 공사기간 산정 보고서 템플릿을 만든다(한 번만 실행).

    python tools/make_report_template.py "<샘플.hwpx>"  →  templates/report_template.hwpx

샘플의 스타일(header.xml)과 고정 문구(고시 발췌·주석 등)는 그대로 두고,
  - 사업마다 바뀌는 값은 {{TITLE}} 같은 자리표시로 바꾸고
  - 다시 만들 표(공휴일·기상조건·월별 비작업일수·분야별 작업일수·유사사업)는 표시 문단 하나로 바꾼다.
샘플의 사업 데이터(사업명·표 내용)는 템플릿에 남기지 않는다. 줄 배치 캐시(linesegarray)는 내용이
바뀌면 맞지 않으므로 모두 지운다(한글이 열 때 다시 계산한다).
"""
import io
import os
import re
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "templates", "report_template.hwpx")


def split_paragraphs(body: str):
    """최상위 <hp:p> 목록(표 안 문단은 바깥 문단에 포함)."""
    out, depth, start = [], 0, None
    for m in re.finditer(r"<(/?)hp:p\b[^>]*?(/?)>", body):
        closing, selfclose = m.group(1) == "/", m.group(2) == "/"
        if not closing and not selfclose:
            if depth == 0:
                start = m.start()
            depth += 1
        elif closing:
            depth -= 1
            if depth == 0:
                out.append(body[start:m.end()])
    return out


def sub1(text, old, new):
    assert old in text, f"원문에 없음: {old[:40]}"
    return text.replace(old, new, 1)


def runs_para(open_tag: str, runs):
    return open_tag + "".join(f'<hp:run charPrIDRef="{cp}"><hp:t>{t}</hp:t></hp:run>' for cp, t in runs) + "</hp:p>"


def open_tag(p: str) -> str:
    return re.match(r"<hp:p\b[^>]*>", p).group(0)


def main(src):
    zin = zipfile.ZipFile(src)
    sec = zin.read("Contents/section0.xml").decode("utf-8")
    sec = re.sub(r"<hp:linesegarray>.*?</hp:linesegarray>", "", sec, flags=re.S)
    root_end = sec.index(">", sec.index("<hs:sec")) + 1
    prefix, body = sec[:root_end], sec[root_end:]
    suffix = body[body.rindex("</hs:sec>"):]
    P = split_paragraphs(body[:body.rindex("</hs:sec>")])
    assert len(P) == 85, f"문단 수가 샘플과 다릅니다: {len(P)}"

    P[0] = re.sub(r"<hp:t>\s*[^<]*설치사업\s*</hp:t>", "<hp:t>{{TITLE}}</hp:t>", P[0], count=1)
    assert "{{TITLE}}" in P[0]
    P[10] = re.sub(r'(<hp:run charPrIDRef="12"><hp:t>).*?(</hp:t></hp:run>)', r"\1{{FORMULA}}\2", P[10],
                   count=1, flags=re.S)
    P[11] = re.sub(r"<hp:t>\s*= [^<]*</hp:t>", "<hp:t>    = {{TOTAL_DAYS}}일 ≒ {{TOTAL_MONTHS}}개월</hp:t>", P[11])
    P[19] = sub1(P[19], "상수도공사의 준비기간을 적용", "{{PREP_LABEL}}의 준비기간을 적용")
    P[20] = sub1(P[20], "<hp:t>상수도공사</hp:t>", "<hp:t>{{PREP_LABEL}}</hp:t>")
    P[20] = sub1(P[20], "<hp:t>60일</hp:t>", "<hp:t>{{PREP_DAYS}}일</hp:t>")
    P[20] = sub1(P[20], "<hp:t>전체 공종의 평균과 동일</hp:t>", "<hp:t>{{PREP_NOTE}}</hp:t>")
    P[29] = sub1(P[29], "금회사업은 공공하수처리장 증설로 ", "금회 사업은 ")
    P[29] = sub1(P[29], "토목공사, 건축공사, 기계공사, 전기 및 계측공사, 조경공사", "{{DISC_CP}}")
    P[30] = runs_para(open_tag(P[30]), [(30, "    - 공사의 품질 확보 및 근로자 안전을 고려하여 기상조건으로 인한 "
                                            "비작업일수를 산정하였으며, 가이드라인 부록의 최근 10년(2015~2024) "
                                            "월평균 기상자료({{STATION}})를 적용")])
    P[36] = "{{T_HOLIDAYS}}"
    P[38] = sub1(P[38], "의성군 적용", "{{STATION}} 적용")
    P[39] = "{{T_CONDITIONS}}"
    P[40] = sub1(P[40], "의성군,", "{{STATION}},")
    P[45] = "{{T_NONWORK}}"
    P[46] = sub1(P[46], "가이드라인 의성군 기상정보", "가이드라인 {{STATION}} 기상정보")
    P[58] = runs_para(open_tag(P[58]), [(62, "  본 사업은 {{DISC_ALL}} 등으로 구분되며 "),
                                        (235, "주요공정(Critical path)은 {{DISC_CP}}를 적용하였으며"),
                                        (61, "{{DISC_OTHER}}")])
    P[62] = "{{WORK_TABLES}}"
    P[78] = sub1(P[78], "<hp:t>30일</hp:t>", "<hp:t>{{WRAPUP}}일</hp:t>")
    P[82] = runs_para(open_tag(P[82]), [(223, "    - "),
                                        (244, "최근 발주·준공한 유사사업의 시설규모와 공사기간을 조사하여 비교함"),
                                        (243, "(아래 표는 직접 작성)")])
    P[83] = "{{T_SIMILAR}}"
    P[84] = runs_para(open_tag(P[84]), [(41, " "), (224, " "), (247, "산정 공사기간 "),
                                        (249, "{{TOTAL_DAYS}}일(≒{{TOTAL_MONTHS}}개월)"),
                                        (207, "은 유사사업의 사업규모·공사기간과 비교하여 적정성을 검토함")])
    # 뺄 문단: 미세먼지 기준표(앱이 반영하지 않음), '적용 기후조건' 중복 표, 샘플 고유 문장, 표 63~74(다시 만듦)
    drop = {34, 41, 42, 43, 61, *range(63, 75)}
    keep = [p for i, p in enumerate(P) if i not in drop]
    new_sec = prefix + "".join(keep) + suffix
    for leftover in ("의성", "영덕", "영해", "가평", "연무", "강하", "내촌"):
        assert leftover not in new_sec, f"샘플 사업 정보가 남음: {leftover}"

    # 표 전용 9pt 글자(보통·굵게) — 빽빽한 월별 표에서 '10월'이 세로로 꺾이지 않게.
    # 샘플 표 글자(232 보통, 216 굵게, 11pt)를 복제해 크기만 줄이고 id를 이어 붙인다.
    head = zin.read("Contents/header.xml").decode("utf-8")
    m = re.search(r'<hh:charProperties itemCnt="(\d+)">', head)
    n0 = int(m.group(1))
    assert n0 == 257, f"글자 스타일 수가 샘플과 다릅니다: {n0}"
    clones = []
    for new_id, src_id in ((n0, 232), (n0 + 1, 216)):
        c = re.search(rf'<hh:charPr id="{src_id}" height="\d+".*?</hh:charPr>', head, flags=re.S).group(0)
        c = re.sub(r'id="\d+" height="\d+"', f'id="{new_id}" height="900"', c, count=1)
        clones.append(c)
    head = head.replace(m.group(0), f'<hh:charProperties itemCnt="{n0 + 2}">', 1)
    head = head.replace("</hh:charProperties>", "".join(clones) + "</hh:charProperties>", 1)

    hpf = zin.read("Contents/content.hpf").decode("utf-8")
    hpf = re.sub(r"<opf:title>.*?</opf:title>", "<opf:title>{{TITLE}}</opf:title>", hpf)
    hpf = re.sub(r'(<opf:meta name="(creator|lastsaveby)" content="text">)[^<]*(</opf:meta>)', r"\1\3", hpf)

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with zipfile.ZipFile(OUT, "w") as zout:
        for info in zin.infolist():
            n = info.filename
            if n == "Preview/PrvImage.png":
                continue                                    # 샘플 첫 쪽 그림 — 넣지 않는다
            data = zin.read(n)
            if n == "Contents/section0.xml":
                data = new_sec.encode("utf-8")
            elif n == "Contents/header.xml":
                data = head.encode("utf-8")
            elif n == "Contents/content.hpf":
                data = hpf.encode("utf-8")
            elif n == "Preview/PrvText.txt":
                data = "{{TITLE}} 공기 적정성 검토".encode("utf-8")
            ctype = zipfile.ZIP_STORED if n in ("mimetype", "version.xml") else zipfile.ZIP_DEFLATED
            zout.writestr(n, data, compress_type=ctype)
    print("템플릿:", os.path.abspath(OUT), f"({len(keep)}개 문단)")


if __name__ == "__main__":
    main(sys.argv[1])
