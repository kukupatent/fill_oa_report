"""
app.py — OA 검토보고서 자동완성 Streamlit 웹앱
실행: streamlit run app.py
"""

import tempfile
from pathlib import Path

import streamlit as st

from fill_oa_report import parse_oa_pdf, fill_docx, make_output_filename

# ── 템플릿 경로 설정 ──────────────────────────
TEMPLATE_PATH = Path(__file__).parent / "template.docx"

# ── 페이지 설정 ───────────────────────────────
st.set_page_config(
    page_title="OA 검토보고서 자동완성",
    page_icon="📄",
    layout="centered",
)

st.title("📄 OA 검토보고서 자동완성")

# 템플릿 파일 존재 확인
if not TEMPLATE_PATH.exists():
    st.error(
        f"❌ 템플릿 파일을 찾을 수 없습니다: `{TEMPLATE_PATH.name}`\n\n"
        "`app.py`와 같은 폴더에 템플릿 DOCX를 `template.docx`로 저장해주세요."
    )
    st.stop()

st.divider()

# ── 파일 업로드: 두 칸을 나란히 ─────────────────
col1, col2 = st.columns(2)

with col1:
    oa_pdf_file = st.file_uploader(
        "의견제출통지서 / 거절결정서",
        type=["pdf"],
        help="파일명에 관리번호(예: DPA240078)가 포함되어 있어야 합니다.",
        key="oa_pdf",
    )

with col2:
    app_pdf_file = st.file_uploader(
        "출원서",
        type=["rtf", "docx"],
        help="RTF 또는 DOCX 파일. 청구항 1 및 전체 청구범위 자동 채우기에 사용됩니다.",
        key="app_pdf",
    )

st.divider()

# ── 실행 버튼 ─────────────────────────────────
run_btn = st.button(
    "🚀 자동완성 실행",
    type="primary",
    use_container_width=True,
    disabled=not oa_pdf_file,
)

if run_btn and oa_pdf_file:
    with st.spinner("처리 중..."):
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir = Path(tmpdir)

                # 의견제출통지서 저장
                oa_path = tmpdir / oa_pdf_file.name
                oa_path.write_bytes(oa_pdf_file.read())

                # 출원서 저장 (선택)
                app_pdf_path = ""
                if app_pdf_file:
                    app_path = tmpdir / app_pdf_file.name
                    app_path.write_bytes(app_pdf_file.read())
                    app_pdf_path = str(app_path)

                out_path = tmpdir / "output.docx"

                # 파싱 & 채우기
                info = parse_oa_pdf(str(oa_path), str(TEMPLATE_PATH))
                fill_docx(str(TEMPLATE_PATH), info, str(out_path),
                          app_pdf_path=app_pdf_path)

                output_filename = make_output_filename(info)
                result_bytes    = out_path.read_bytes()

            # ── 결과 표시 ──────────────────────
            st.success("✅ 완료!")

            with st.expander("📋 파싱된 서지사항 확인", expanded=True):
                labels = {
                    "출원번호":           "출원번호",
                    "출원일":            "출원일",
                    "당소관리번호":        "당소관리번호",
                    "출원인":            "출원인",
                    "발명자":            "발명자",
                    "발명의 명칭":        "발명의 명칭",
                    "통지서 발행일":       "통지서 발행일",
                    "의견서 마감일":       "의견서 마감일",
                    "OA 종류":           "OA 종류",
                    "발명자 검토회신 요청일": "발명자 검토회신 요청일",
                }
                for key, label in labels.items():
                    val = info.get(key, "")
                    st.markdown(f"**{label}**: {val if val else '–'}")

                reasons = info.get("거절이유_set", set())
                st.markdown(f"**거절이유**: {', '.join(sorted(reasons)) if reasons else '–'}")

                if app_pdf_file:
                    st.info("📎 출원서 청구항이 '현재 청구항' 및 '[첨부 1] 당소 보정안'에 채워졌습니다.")

            # ── 다운로드 버튼 ──────────────────
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

# ── 사용 안내 ─────────────────────────────────
with st.expander("💡 사용 방법"):
    st.markdown("""
1. **의견제출통지서** 업로드 (필수)
   - 파일명에 관리번호(예: `DPA240078`)가 포함되어 있어야 합니다
2. **출원서** 업로드 (선택, RTF 또는 DOCX)
   - 업로드 시 청구항 자동 채우기 기능이 활성화됩니다
3. **자동완성 실행** 버튼 클릭
4. 결과 DOCX **다운로드**

---
**자동으로 채워지는 항목**

*의견제출통지서 기반*
- 출원번호, 출원일, 출원인, 발명자, 발명의 명칭
- 통지서 발행일, 의견서 마감일, OA 종류, 거절이유
- 당소관리번호 (파일명에서 추출)
- 발명자 검토회신 요청일 (통지서 발행일 +2개월)
- 거절이유별 대응방안 요약 표, OA 내용 분석 표

*출원서 기반 (출원서 업로드 시)*
- 현재 청구항 (독립항 제 1항) 본문
- [첨부 1] 당소 보정안 전체 청구범위
""")
