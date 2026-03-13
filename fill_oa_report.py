"""
fill_oa_report.py
=================
의견제출통지서(PDF) + OA 검토보고서 템플릿(DOCX) →
서지사항 표를 자동으로 채운 DOCX를 출력합니다.

사용법:
    python fill_oa_report.py <통지서.pdf> <템플릿.docx> [출력.docx]

예시:
    python fill_oa_report.py 의견제출통지서.pdf 1차_OA_검토보고서.docx 결과.docx

필요 라이브러리:
    pip install python-docx pdfplumber
"""

import sys, re, shutil, calendar
from pathlib import Path
from datetime import date
import pdfplumber
from docx import Document

# ── 상수 ─────────────────────────────────────
FONT_NAME = "SpoqaHanSans-Light"
FONT_SIZE_PT = 10          # 포인트 단위 (OOXML sz = half-point, 즉 ×2)

REJECTION_LABELS = [
    "신규성", "신규사항추가", "기재불비",
    "미완성 발명", "진보성", "기타(비법정 발명)"
]

# 특허법 조항 → 거절이유 매핑 (띄어쓰기 허용 패턴)
LAW_TO_REASON = {
    r"제\s*29\s*조\s*제\s*1\s*항": "신규성",
    r"제\s*29\s*조\s*제\s*2\s*항": "진보성",
    r"제\s*47\s*조\s*제\s*2\s*항": "신규사항추가",
    r"제\s*42\s*조":               "기재불비",
    r"미\s*완\s*성":               "미완성 발명",
    r"비\s*법\s*정\s*발\s*명":     "기타(비법정 발명)",
}

# 서지사항 표(표2) 레이아웃: {행: {고유셀_열: 필드키}}
# 고유셀 기준 열 인덱스 (레이블=짝수, 값=홀수)
TABLE_LAYOUT = {
    0: {1: "출원종류",       3: "출원국가"},
    1: {1: "출원번호",       3: "출원일"},
    2: {1: "당소관리번호",    3: "출원인"},
    3: {1: "발명자"},
    4: {1: "발명의 명칭"},
    5: {1: "통지서 발행일",   3: "OA 종류"},
    6: {1: "의견서 마감일",   3: "발명자 검토회신 요청일"},
    7: {1: "거절이유"},
}

LABEL_TEXTS = {
    "출원종류","출원국가","출원번호","출원일","당소관리번호","출원인",
    "발명자","발명의 명칭","통지서 발행일","OA 종류","의견서 마감일",
    "발명자 검토회신 요청일","거절이유",
}


# ── 1. PDF 파싱 ──────────────────────────────

def add_months(d: date, months: int) -> date:
    """날짜에 월 수를 더합니다 (월말 처리 포함)."""
    month = d.month - 1 + months
    year  = d.year + month // 12
    month = month % 12 + 1
    day   = min(d.day, calendar.monthrange(year, month)[1])
    return date(year, month, day)


def parse_rejection_from_table(pdf_path: str) -> set:
    """
    거절이유 표(1페이지)의 '관련 법조항' 컬럼에서 직접 추출.
    텍스트 추출보다 훨씬 안정적.
    """
    found = set()
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # 1페이지 표에서 거절이유 추출
            tables = pdf.pages[0].extract_tables()
            for tbl in tables:
                if not tbl or len(tbl[0]) < 2:
                    continue
                # 헤더 확인: '관련 법조항' 포함 여부
                header = [str(c or "") for c in tbl[0]]
                if not any("법조항" in h for h in header):
                    continue
                law_col = next(i for i, h in enumerate(header) if "법조항" in h)
                for row in tbl[1:]:
                    law_text = str(row[law_col] or "")
                    for pattern, reason in LAW_TO_REASON.items():
                        if re.search(pattern, law_text):
                            found.add(reason)
    except Exception:
        pass
    return found


def parse_oa_pdf(pdf_path: str, template_path: str = "") -> dict:
    with pdfplumber.open(pdf_path) as pdf:
        full_text = "\n".join(p.extract_text() or "" for p in pdf.pages)

    def find(pattern, flags=0):
        m = re.search(pattern, full_text, flags)
        return m.group(1).strip() if m else ""

    # ── 기본 항목 (패턴 강화: 띄어쓰기·구분자 유연하게)
    info = {}

    # 출원번호
    info["출원번호"] = find(r"출\s*원\s*번\s*호\s+([\d\-]+)")

    # 출원일 (YYYY.MM.DD 또는 YYYY-MM-DD)
    raw = find(r"출\s*원\s*일\s*자?\s*[:\s]\s*(\d{4}[\.\-]\d{1,2}[\.\-]\d{1,2}\.?)")
    info["출원일"] = raw.rstrip(".").replace("-", ".")

    # 출원인 (스마트쿼트 제거)
    raw = find(r"출\s*원\s*인\s*성\s*명\s+(.+?)(?:\s*\(특허고객번호|$)", re.DOTALL)
    info["출원인"] = raw.split("\n")[0].strip().strip("\u2018\u2019\u201c\u201d'\"")

    # 발명자 (여러 명 모두 추출, 쉼표로 연결)
    names = re.findall(r"발\s*명\s*자\s*성\s*명\s+(\S+)", full_text)
    info["발명자"] = ", ".join(names)

    # 발명의 명칭
    raw = find(r"발\s*명\s*의\s*명\s*칭\s+(.+?)(?:\n발송번호|\n출원번호|\n1\.|$)", re.DOTALL)
    info["발명의 명칭"] = " ".join(raw.split())

    # 통지서 발행일 (발송일자)
    info["통지서 발행일"] = find(r"발송일자[:\s]+(\d{4}\.\d{2}\.\d{2}\.?)").rstrip(".")

    # 의견서 마감일 (제출기일)
    raw = find(r"제출기일[:\s]+(\d{4}\.\d{2}\.\d{2}\.?)")
    info["의견서 마감일"] = raw.rstrip(".") + "." if raw else ""

    # ── 당소관리번호: PDF 파일명에서 추출
    # 형식: [DP|OP|IP][A|M|E|C] + 숫자
    filename = Path(pdf_path).stem
    m = re.search(r'([DIO]P[AMEC]\d+)', filename, re.IGNORECASE)
    info["당소관리번호"] = m.group(1).upper() if m else ""

    # ── 발명자 검토회신 요청일: 통지서 발행일 + 2개월
    try:
        parts = info["통지서 발행일"].split(".")
        y, mo, d = int(parts[0]), int(parts[1]), int(parts[2])
        info["발명자 검토회신 요청일"] = add_months(date(y, mo, d), 2).strftime("%Y.%m.%d.")
    except Exception:
        info["발명자 검토회신 요청일"] = ""

    # ── OA 종류: 문서 제목으로 판별
    if re.search(r"거\s*절\s*결\s*정\s*서", full_text):
        info["OA 종류"] = "거절결정"
    else:
        info["OA 종류"] = "1차 OA"

    # ── 거절이유: 거절이유 표에서 직접 추출 (가장 정확)
    info["거절이유_set"] = parse_rejection_from_table(pdf_path)

    # 표 추출 실패 시 본문 텍스트로 fallback
    if not info["거절이유_set"]:
        for pattern, reason in LAW_TO_REASON.items():
            if re.search(pattern, full_text):
                info["거절이유_set"].add(reason)

    return info


# ── 2. 셀 조작 헬퍼 ──────────────────────────

def unique_cell(row, col_index: int):
    """병합 고려: col_index번째 고유 셀 반환"""
    seen, col = set(), 0
    for cell in row.cells:
        cid = id(cell._tc)
        if cid not in seen:
            seen.add(cid)
            if col == col_index:
                return cell
            col += 1
    return None


def apply_font(run, bold=False, underline=False, color_rgb=None):
    """run에 표준 폰트를 적용합니다. 추가 서식 옵션 지원."""
    from docx.oxml.ns import qn
    from lxml import etree
    rpr = run._r.find(qn('w:rPr'))
    if rpr is None:
        rpr = etree.SubElement(run._r, qn('w:rPr'))
        run._r.insert(0, rpr)
    # 폰트
    fonts = rpr.find(qn('w:rFonts'))
    if fonts is None:
        fonts = etree.SubElement(rpr, qn('w:rFonts'))
    fonts.set(qn('w:ascii'),    FONT_NAME)
    fonts.set(qn('w:eastAsia'), FONT_NAME)
    fonts.set(qn('w:hAnsi'),    FONT_NAME)
    fonts.set(qn('w:cs'),       FONT_NAME)
    # 크기 (half-point 단위)
    half_pt = str(FONT_SIZE_PT * 2)
    for tag in (qn('w:sz'), qn('w:szCs')):
        el = rpr.find(tag)
        if el is None:
            el = etree.SubElement(rpr, tag)
        el.set(qn('w:val'), half_pt)
    # 볼드
    b_el = rpr.find(qn('w:b'))
    if bold:
        if b_el is None:
            b_el = etree.SubElement(rpr, qn('w:b'))
    else:
        if b_el is not None:
            rpr.remove(b_el)
    # 밑줄
    u_el = rpr.find(qn('w:u'))
    if underline:
        if u_el is None:
            u_el = etree.SubElement(rpr, qn('w:u'))
        u_el.set(qn('w:val'), 'single')
    else:
        if u_el is not None:
            rpr.remove(u_el)
    # 글자색
    color_el = rpr.find(qn('w:color'))
    if color_rgb:
        if color_el is None:
            color_el = etree.SubElement(rpr, qn('w:color'))
        color_el.set(qn('w:val'), color_rgb)
    else:
        if color_el is not None:
            rpr.remove(color_el)


def set_cell_text(cell, text: str, bold=False, underline=False, color_rgb=None):
    """첫 번째 단락·run에 텍스트 설정, 나머지 run 제거, 서식 보존"""
    if not cell.paragraphs:
        return
    for para in cell.paragraphs[1:]:
        para._p.getparent().remove(para._p)
    para = cell.paragraphs[0]
    if not para.runs:
        run = para.add_run(text)
        apply_font(run, bold=bold, underline=underline, color_rgb=color_rgb)
        return
    para.runs[0].text = text
    apply_font(para.runs[0], bold=bold, underline=underline, color_rgb=color_rgb)
    for run in para.runs[1:]:
        run._r.getparent().remove(run._r)


def fill_rejection(cell, found_reasons: set):
    """
    각 단락의 run들을 하나로 합친 뒤, ◆/◇ 기호를 교체하고
    단락 전체를 단일 run으로 재작성합니다.
    공백·들여쓰기는 보존합니다.
    """
    for para in cell.paragraphs:
        if not para.runs:
            continue

        # 1) 단락 전체 텍스트 합치기
        full_text = "".join(run.text for run in para.runs)

        # 2) ◆/◇ 교체
        for label in REJECTION_LABELS:
            mark = "◆" if label in found_reasons else "◇"
            full_text = full_text.replace(f"◆{label}", f"{mark}{label}")
            full_text = full_text.replace(f"◇{label}", f"{mark}{label}")

        # 3) 첫 번째 run에 전체 텍스트 쓰고, 나머지 run 제거
        first_run = para.runs[0]
        apply_font(first_run)
        first_run.text = full_text
        for run in para.runs[1:]:
            run._r.getparent().remove(run._r)


# ── 3. DOCX 채우기 ───────────────────────────

def fill_docx(template_path: str, info: dict, output_path: str):
    shutil.copy(template_path, output_path)
    doc = Document(output_path)

    if len(doc.tables) < 3:
        raise ValueError(f"서지사항 표를 찾을 수 없습니다. (표 개수: {len(doc.tables)})")

    # ── 표지 채우기 (표1, 0-indexed)
    cover = doc.tables[1]

    def set_cover_cell(row_idx, text):
        """표지 값 셀: 기존 run들을 합쳐 첫 run에 쓰고 서식 보존."""
        seen=set(); cells=[]
        for cell in cover.rows[row_idx].cells:
            cid=id(cell._tc)
            if cid not in seen:
                seen.add(cid); cells.append(cell)
        val_cell = cells[1]
        para = val_cell.paragraphs[0]
        if not para.runs:
            return
        para.runs[0].text = text
        for run in para.runs[1:]:
            run._r.getparent().remove(run._r)

    def set_cover_title(text):
        """표지 제목 셀: 기존 run들을 합쳐 첫 run에 쓰고 서식 보존."""
        para = cover.rows[0].cells[0].paragraphs[0]
        if not para.runs:
            return
        para.runs[0].text = text
        for run in para.runs[1:]:
            run._r.getparent().remove(run._r)

    # 제목
    oa_type = info.get("OA 종류", "")
    if oa_type == "거절결정":
        cover_title = "거절결정 검토 보고서"
    else:
        cover_title = "1차 OA 검토 보고서"
    set_cover_title(cover_title)

    # 의뢰인 → 출원인
    set_cover_cell(1, info.get("출원인", ""))

    # 작성일 → 오늘
    from datetime import date as _date
    today = _date.today().strftime("%Y. %m. %d.").replace(" 0", " ")
    set_cover_cell(2, today)

    # 담당자 → 빈칸
    set_cover_cell(3, "")

    tbl = doc.tables[2]
    fill_map = {
        "출원종류":             "특허",
        "출원국가":             "한국",
        "출원번호":             info.get("출원번호", ""),
        "출원일":              info.get("출원일", ""),
        "당소관리번호":          info.get("당소관리번호", ""),
        "출원인":              info.get("출원인", ""),
        "발명자":              info.get("발명자", ""),
        "발명의 명칭":          info.get("발명의 명칭", ""),
        "통지서 발행일":         info.get("통지서 발행일", ""),
        "OA 종류":             info.get("OA 종류", ""),
        "의견서 마감일":         info.get("의견서 마감일", ""),
        "발명자 검토회신 요청일": info.get("발명자 검토회신 요청일", ""),
    }
    found_reasons = info.get("거절이유_set", set())

    for row_idx, col_map in TABLE_LAYOUT.items():
        if row_idx >= len(tbl.rows):
            continue
        row = tbl.rows[row_idx]
        for col_idx, field_key in col_map.items():
            cell = unique_cell(row, col_idx)
            if cell is None or cell.text.strip() in LABEL_TEXTS:
                continue
            if field_key == "거절이유":
                fill_rejection(cell, found_reasons)
            elif field_key == "의견서 마감일":
                set_cell_text(cell, fill_map.get(field_key, ""),
                              bold=True, underline=True)
            elif field_key == "발명자 검토회신 요청일":
                set_cell_text(cell, fill_map.get(field_key, ""),
                              bold=True, underline=True, color_rgb="FF0000")
            else:
                set_cell_text(cell, fill_map.get(field_key, ""))

    doc.save(output_path)


# ── 4. 진입점 ────────────────────────────────

def make_output_filename(info: dict) -> str:
    """OA 종류와 관리번호로 출력 파일명 자동 생성."""
    mgmt_no = info.get("당소관리번호", "")
    oa_type = info.get("OA 종류", "")

    if oa_type == "거절결정":
        suffix = "거절결정검토보고서"
    else:
        suffix = "1차OA검토보고서"

    return f"[파이특허][{mgmt_no}]{suffix}.docx"


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    pdf_path      = sys.argv[1]
    template_path = sys.argv[2]

    print(f"[1/3] PDF 파싱: {pdf_path}")
    info = parse_oa_pdf(pdf_path, template_path)
    print("      결과:")
    for k, v in info.items():
        print(f"        {'거절이유 항목' if k == '거절이유_set' else k}: {v}")

    # 출력 파일명: 인수로 지정하면 그걸 쓰고, 없으면 자동 생성
    output_path = sys.argv[3] if len(sys.argv) >= 4 else make_output_filename(info)

    print(f"\n[2/3] 템플릿: {template_path}")
    print(f"[3/3] 저장:   {output_path}")
    fill_docx(template_path, info, output_path)
    print(f"\n✓ 완료: {output_path}")


if __name__ == "__main__":
    main()
