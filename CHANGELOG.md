# Changelog

## v1.2.3

- PPTX + PDF + PNG 동시 출력 시 PDF/PNG가 실패하던 저장 순서 문제 수정
- `Presentation.SaveAs()` 연쇄 호출 제거
- PDF 출력: `Presentation.ExportAsFixedFormat()` 사용
- PNG 출력: 각 슬라이드의 `Slide.Export()` 사용
- PPTX 출력: 마지막 단계에서 `SaveCopyAs()` 사용
- 출력 형식별 예외를 독립 처리하여 한 형식 실패가 다른 형식을 막지 않도록 변경
- PowerPoint 백그라운드 모드를 `WithWindow=True` + `ppWindowMinimized` 방식으로 변경
- COM 오류 발생 시 HRESULT/source/description을 가능한 범위에서 로그에 표시

## v1.2.1

- XLSX 읽기를 `openpyxl` 대신 Excel Win32 COM으로 변경
- PPTX 슬라이드 비율 확인을 ZIP/XML 파싱 대신 PowerPoint Win32 COM으로 변경
- 런타임의 `zipfile`, `openpyxl`, `load_workbook` 의존 제거
- EXE 내부 샘플 리소스명을 `template.pptx`, `data.xlsx`로 ASCII 고정
- 현재 처리 중 이름, 처리 건수, 퍼센트 진행률 표시 추가

## 초기 기능

- `{컬럼명}` 기반 PPTX 텍스트/파일명 치환
- Excel 첫 시트 1행을 헤더로 사용
- PPTX / PDF / PNG 선택 출력
- PNG 너비/높이 슬라이더 및 비율 고정
- 세로형 공식 인증서 샘플 제공
