"""
app.py — OA 검토보고서 자동완성 Streamlit 웹앱
실행: streamlit run app.py
"""

import tempfile
from pathlib import Path
from datetime import datetime

import streamlit as st

from fill_oa_report import (
    parse_oa_pdf, fill_docx, make_output_filename,
    FEASIBILITY_OPTIONS, REASON_GROUPS, DEFAULT_FEASIBILITY_LABEL,
    REASON_TO_LAW,
)

# ── 템플릿 경로 설정 ──────────────────────────
TEMPLATE_PATH = Path(__file__).parent / "template.docx"

# ── 페이지 설정 ───────────────────────────────
st.set_page_config(
    page_title="OA 검토보고서 자동완성",
    page_icon="📄",
    layout="centered",
)

st.title("📄 OA 검토보고서 자동완성")

if not TEMPLATE_PATH.exists():
    st.error(
        f"❌ 템플릿 파일을 찾을 수 없습니다: `{TEMPLATE_PATH.name}`\n\n"
        "`app.py`와 같은 폴더에 템플릿 DOCX를 `template.docx`로 저장해주세요."
    )
    st.stop()

st.divider()

# ── 1단계: 파일 업로드 ────────────────────────
col1, col2 = st.columns(2)

with col1:
    oa_pdf_file = st.file_uploader(
        "의견제출통지서 / 거절결정서",
        type=["pdf"],
        help="파일명에 관리번호(예: DPA240078)가 포함되어 있어야 합니다.",
        key="oa_pdf",
    )

with col2:
    app_file = st.file_uploader(
        "출원서",
        type=["rtf", "docx"],
        help="RTF 또는 DOCX 파일. 청구항 1 및 전체 청구범위 자동 채우기에 사용됩니다.",
        key="app_file",
    )

# ── 2단계: 파싱 및 서지사항 표시 ──────────────
if oa_pdf_file:
    # 파일이 바뀌었을 때만 재파싱 (session_state 캐싱)
    file_key = f"{oa_pdf_file.name}_{oa_pdf_file.size}"
    if st.session_state.get("parsed_key") != file_key:
        with st.spinner("파싱 중..."):
            try:
                with tempfile.TemporaryDirectory() as tmpdir:
                    oa_path = Path(tmpdir) / oa_pdf_file.name
                    oa_path.write_bytes(oa_pdf_file.read())
                    info = parse_oa_pdf(str(oa_path), str(TEMPLATE_PATH))
                st.session_state["info"]       = info
                st.session_state["parsed_key"] = file_key
            except Exception as e:
                st.error(f"❌ 파싱 오류: {e}")
                st.stop()

    info = st.session_state.get("info", {})

    st.divider()

    # ── 서지사항 표시 ─────────────────────────
    st.subheader("📋 파싱된 서지사항")

    fields = [
        ("당소관리번호",   "당소관리번호"),
        ("출원번호",      "출원번호"),
        ("출원일",        "출원일"),
        ("출원인",        "출원인"),
        ("발명자",        "발명자"),
        ("발명의 명칭",    "발명의 명칭"),
        ("통지서 발행일",  "통지서 발행일"),
        ("OA 종류",       "OA 종류"),
    ]

    col_a, col_b = st.columns(2)
    for i, (key, label) in enumerate(fields):
        val = info.get(key, "") or "–"
        target_col = col_a if i % 2 == 0 else col_b
        target_col.markdown(f"**{label}**  \n{val}")

    # 거절이유 + 청구항 + 법조항
    reasons   = info.get("거절이유_set", set())
    claim_map = info.get("거절이유_청구항", {})

    st.markdown("**거절이유 / 청구항 / 관련 법조항**")
    if reasons:
        # REASON_GROUPS 순서로 정렬 후 표시
        from fill_oa_report import REASON_GROUPS as _RG
        remaining = set(reasons)
        ordered = []
        for reason_set, label in _RG:
            if reason_set <= remaining:
                ordered.append(label)
                remaining -= reason_set
        for r in sorted(remaining):
            ordered.append(r)

        rows = []
        for label in ordered:
            law   = REASON_TO_LAW.get(label, "–")
            # 단일 거절이유인 경우 원래 키로 청구항 조회
            claims = claim_map.get(label, "")
            if not claims:
                # 신규성/진보성 합쳐진 경우 개별 키로 시도
                parts = [claim_map.get(r, "") for r in (label.split("/") if "/" in label else [label]) if claim_map.get(r)]
                claims = " / ".join(parts)
            rows.append((label, claims or "–", law))

        import pandas as pd
        df = pd.DataFrame(rows, columns=["거절이유", "해당 청구항", "관련 법조항"])
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.markdown("–")

    st.divider()

    # ── 의견서 마감일 표시 + 발명자 회신일 입력 ──
    deadline    = info.get("의견서 마감일", "")
    auto_review = info.get("발명자 검토회신 요청일", "")

    col_d, col_r = st.columns(2)

    with col_d:
        st.markdown("**📅 의견서 마감일**")
        st.markdown(
            f"<span style='font-size:1.1em; font-weight:bold; color:#d32f2f;'>{deadline or '–'}</span>",
            unsafe_allow_html=True,
        )

    with col_r:
        review_input = st.text_input(
            "✏️ 발명자 검토회신 요청일",
            value=auto_review,
            help=f"입력하지 않으면 의견서 마감일 2개월 전({auto_review})으로 자동 입력됩니다. 형식: YYYY.MM.DD.",
            placeholder=f"예: {auto_review}",
        )

    # 입력값 또는 자동값 적용
    final_review = review_input.strip() if review_input.strip() else auto_review
    if final_review != info.get("발명자 검토회신 요청일", ""):
        info = dict(info)
        info["발명자 검토회신 요청일"] = final_review

    st.divider()

    # ── 3단계: 극복가능성 선택 ───────────────────
    st.subheader("📊 극복가능성 선택")

    # 거절이유 레이블 목록 계산 (REASON_GROUPS 순서)
    remaining = set(reasons)
    reason_labels = []
    for reason_set, label in REASON_GROUPS:
        if reason_set <= remaining:
            reason_labels.append(label)
            remaining -= reason_set
    for r in sorted(remaining):
        reason_labels.append(r)

    # 극복가능성 옵션 및 배경색
    f_options = list(FEASIBILITY_OPTIONS.keys())          # ["매우 높음", "다소 높음", "중간"]
    f_colors  = {"매우 높음": "#5AFFAF", "다소 높음": "#91FFCA", "중간": "#FFD899"}

    override_feasibility = {}
    if reason_labels:
        cols = st.columns(len(reason_labels))
        for col, label in zip(cols, reason_labels):
            with col:
                key = f"feasibility_{label}"
                # session_state에 없으면 기본값으로 초기화
                if key not in st.session_state:
                    st.session_state[key] = DEFAULT_FEASIBILITY_LABEL
                selected = col.selectbox(
                    label,
                    options=f_options,
                    key=key,
                )
                override_feasibility[label] = selected
                # 선택된 배경색으로 색상 표시
                bg = f_colors[selected]
                col.markdown(
                    f"<div style='background:{bg}; padding:4px 10px; border-radius:4px; "
                    f"text-align:center; font-size:0.85em;'>{FEASIBILITY_OPTIONS[selected][0].replace(chr(10), ' ')}</div>",
                    unsafe_allow_html=True,
                )
    else:
        st.info("거절이유가 파싱되지 않았습니다.")

    st.divider()

    # ── 4단계: 문서 생성 버튼 ─────────────────
    run_btn = st.button(
        "🚀 검토보고서 생성",
        type="primary",
        use_container_width=True,
    )

    if run_btn:
        with st.spinner("문서 생성 중..."):
            try:
                with tempfile.TemporaryDirectory() as tmpdir:
                    tmpdir = Path(tmpdir)

                    app_path_str = ""
                    if app_file:
                        app_path = tmpdir / app_file.name
                        app_path.write_bytes(app_file.read())
                        app_path_str = str(app_path)

                    out_path = tmpdir / "output.docx"
                    fill_docx(str(TEMPLATE_PATH), info, str(out_path),
                              app_pdf_path=app_path_str,
                              override_feasibility=override_feasibility)

                    output_filename = make_output_filename(info)
                    result_bytes    = out_path.read_bytes()

                st.success("✅ 완료!")

                if app_file:
                    st.info("📎 출원서 청구항이 '현재 청구항' 및 '[첨부 1] 당소 보정안'에 채워졌습니다.")

                st.download_button(
                    label=f"⬇️ {output_filename} 다운로드",
                    data=result_bytes,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    type="primary",
                )

            except Exception as e:
                st.error(f"❌ 오류 발생: {e}")
                st.exception(e)

else:
    st.info("👆 의견제출통지서 PDF를 업로드하면 서지사항이 표시됩니다.")

# ── 사용 안내 ─────────────────────────────────
with st.expander("💡 사용 방법"):
    st.markdown("""
1. **의견제출통지서** 업로드 (필수)
   - 파일명에 관리번호(예: `DPA240078`)가 포함되어 있어야 합니다
   - 업로드 즉시 서지사항이 자동으로 파싱됩니다
2. **출원서** 업로드 (선택, RTF 또는 DOCX)
3. **발명자 검토회신 요청일** 확인 또는 수정
   - 비워두면 의견서 마감일 기준 2개월 전으로 자동 설정됩니다
4. **검토보고서 생성** 버튼 클릭
5. 결과 DOCX **다운로드**

---
**자동으로 채워지는 항목**

*의견제출통지서 기반*
- 출원번호, 출원일, 출원인, 발명자, 발명의 명칭
- 통지서 발행일, 의견서 마감일, OA 종류, 거절이유
- 당소관리번호 (파일명에서 추출)
- 발명자 검토회신 요청일 (마감일 기준 2개월 전)
- 거절이유별 대응방안 요약 표, OA 내용 분석 표

*출원서 기반 (출원서 업로드 시)*
- 현재 청구항 (독립항 제 1항) 본문
- [첨부 1] 당소 보정안 전체 청구범위
""")
