"""
app.py — OA 검토보고서 자동완성 Streamlit 웹앱
실행: streamlit run app.py
"""

import io
import tempfile
from pathlib import Path

import streamlit as st

from fill_oa_report import parse_oa_pdf, fill_docx, make_output_filename

# ── 템플릿 경로 설정 ──────────────────────────
# app.py와 같은 폴더에 템플릿 파일을 넣어두세요
TEMPLATE_PATH = Path(__file__).parent / "template.docx"

# ── 페이지 설정 ───────────────────────────────
st.set_page_config(
    page_title="OA 검토보고서 자동완성",
    page_icon="📄",
    layout="centered",
)

st.title("📄 OA 검토보고서 자동완성")
st.caption("의견제출통지서 PDF → 서지사항 자동 채우기")

# 템플릿 파일 존재 확인
if not TEMPLATE_PATH.exists():
    st.error(
        f"❌ 템플릿 파일을 찾을 수 없습니다: `{TEMPLATE_PATH.name}`\n\n"
        "`app.py`와 같은 폴더에 템플릿 DOCX를 `template.docx`로 저장해주세요."
    )
    st.stop()

st.divider()

# ── 파일 업로드 ───────────────────────────────
pdf_file = st.file_uploader(
    "의견제출통지서 / 거절결정서 (PDF)",
    type=["pdf"],
    help="파일명에 관리번호(예: DPA240078)가 포함되어 있어야 합니다.",
)

st.divider()

# ── 실행 버튼 ─────────────────────────────────
run_btn = st.button("🚀 자동완성 실행", type="primary", use_container_width=True,
                    disabled=not pdf_file)

if run_btn and pdf_file:
    with st.spinner("처리 중..."):
        try:
            # 임시 파일로 저장
            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir = Path(tmpdir)

                pdf_path  = tmpdir / pdf_file.name
                docx_path = TEMPLATE_PATH          # 고정 템플릿 사용
                out_path  = tmpdir / "output.docx"

                pdf_path.write_bytes(pdf_file.read())

                # 파싱
                info = parse_oa_pdf(str(pdf_path), str(docx_path))

                # 채우기
                fill_docx(str(docx_path), info, str(out_path))

                # 출력 파일명
                output_filename = make_output_filename(info)

                # 결과 바이트 읽기
                result_bytes = out_path.read_bytes()

            # ── 파싱 결과 표시 ─────────────────
            st.success("✅ 완료!")

            with st.expander("📋 파싱된 서지사항 확인", expanded=True):
                labels = {
                    "출원번호": "출원번호",
                    "출원일": "출원일",
                    "당소관리번호": "당소관리번호",
                    "출원인": "출원인",
                    "발명자": "발명자",
                    "발명의 명칭": "발명의 명칭",
                    "통지서 발행일": "통지서 발행일",
                    "의견서 마감일": "의견서 마감일",
                    "OA 종류": "OA 종류",
                    "발명자 검토회신 요청일": "발명자 검토회신 요청일",
                }
                for key, label in labels.items():
                    val = info.get(key, "")
                    st.markdown(f"**{label}**: {val if val else '–'}")

                reasons = info.get("거절이유_set", set())
                st.markdown(f"**거절이유**: {', '.join(sorted(reasons)) if reasons else '–'}")

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
1. **의견제출통지서 PDF** 업로드
   - 파일명에 관리번호(예: `DPA240078`)가 포함되어 있어야 합니다
   - 예: `_파이특허__DPA240078_의견제출통지서.pdf`
2. **자동완성 실행** 버튼 클릭
3. 결과 DOCX **다운로드**

---
**자동으로 채워지는 항목**
- 출원번호, 출원일, 출원인, 발명자, 발명의 명칭
- 통지서 발행일, 의견서 마감일, OA 종류
- 당소관리번호 (파일명에서 추출)
- 발명자 검토회신 요청일 (통지서 발행일 +2개월)
- 거절이유 ◆/◇ 표시
- 표지 제목, 의뢰인, 작성일 자동 설정
""")
