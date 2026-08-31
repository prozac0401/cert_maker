# CertificateMaker

Excel 데이터와 PowerPoint 템플릿을 이용해 인증서를 일괄 생성하는 **Windows / Win32 COM 기반 GUI 도구**입니다.

현재 보관 기준 버전은 **v1.2.3**입니다.

## 동작 개요

1. `template.pptx`의 텍스트 상자 등에 `{이름}`, `{과정명}`처럼 변수를 작성합니다.
2. 데이터 XLSX의 첫 번째 시트 1행에 `이름`, `과정명` 등 동일한 헤더를 작성합니다.
3. 프로그램이 각 행의 값을 PPTX 텍스트와 출력 파일명에 치환합니다.
4. PPTX / PDF / PNG 중 선택한 형식으로 일괄 생성합니다.

## 주요 기능

- Excel 첫 번째 시트의 첫 행을 헤더로 사용
- PPTX 텍스트 상자, 일반 도형, 표 셀, 그룹 도형의 `{컬럼명}` 치환
- 출력 파일명 패턴의 `{컬럼명}` 치환
- PPTX / PDF / PNG 동시 또는 선택 출력
- PNG 너비/높이 400~5000px 슬라이더 조절
- PNG 비율 고정 옵션
- PowerPoint 창 표시 옵션 기본 해제
- 현재 처리 중인 이름, `처리건수 / 전체건수`, 퍼센트 및 Progress Bar 표시
- 같은 파일명이 있으면 `_2`, `_3`을 붙여 충돌 회피
- 예제 PPTX/XLSX 제공

## v1.2.3의 핵심 안정화

실사용 테스트에서 **PPTX 출력을 체크하면 PDF/PNG가 실패하고, PPTX 출력을 끄면 PDF/PNG가 정상 생성되는 현상**이 확인되었습니다.

v1.2.3에서는 출력 경로를 다음처럼 분리했습니다.

- PDF → `Presentation.ExportAsFixedFormat()`
- PNG → `Slide.Export()`
- PPTX → 마지막 단계에서 `Presentation.SaveCopyAs()`

`Presentation.SaveAs()` 연쇄 호출은 제거했습니다. 또한 한 형식에서 오류가 발생해도 나머지 형식은 계속 생성하도록 독립 처리합니다.

PowerPoint 백그라운드 실행 역시 `WithWindow=False`를 사용하지 않고, 실제 `DocumentWindow`를 `WithWindow=True`로 만든 뒤 `ppWindowMinimized`로 최소화합니다. 일부 Office/DRM 환경에서 창이 없는 프레젠테이션의 Save/Export가 실패하던 문제를 회피하기 위한 구조입니다.

## 예제 파일

- `samples/template.pptx`
  - A4에 가까운 세로형 인증서
  - 변수: `{증서번호}`, `{이름}`, `{소속}`, `{과정명}`, `{수료일}`
  - 발급기관 표기: **첨단기술아카데미**
- `samples/data.xlsx`
  - 헤더: `이름`, `과정명`, `소속`, `수료일`, `증서번호`
  - 저장소에는 예시용 가상 데이터만 포함합니다.

## 실행 환경

- Windows 10/11 x64 권장
- Microsoft Excel Desktop 설치 필요
- Microsoft PowerPoint Desktop 설치 필요
- 소스 실행 / EXE 빌드: Python 3.10 이상 권장

웹 버전 Office만으로는 Win32 COM 자동화가 동작하지 않습니다.

## 소스 실행

```bat
build_exe.bat
run_from_source.bat
```

`build_exe.bat`는 `.venv` 생성과 의존성 설치까지 수행합니다.

## Windows EXE 빌드

```bat
build_exe.bat
```

생성 결과:

```text
dist\CertificateMaker.exe
```

또는 GitHub Actions의 **Build Windows EXE** workflow를 실행하면 Windows x64 실행파일을 artifact로 받을 수 있습니다.

> 저장소에는 빌드 산출물(EXE, ZIP)을 직접 커밋하지 않습니다. 소스, 샘플, 빌드 방법을 영구 보관하고 실행파일은 버전별 Release 또는 CI artifact로 관리하는 방식을 권장합니다.

## 인코딩 / 바이너리 관리

- Python/Markdown/YAML: UTF-8
- BAT 실행 시 `chcp 65001` 및 `PYTHONUTF8=1` 사용
- PPTX/XLSX는 `.gitattributes`에서 binary로 고정
- 런타임에서 PPTX/XLSX를 ZIP 파서로 직접 열지 않음
  - XLSX: Excel COM
  - PPTX: PowerPoint COM
- EXE 내부 샘플 리소스명은 `template.pptx`, `data.xlsx`로 ASCII 고정

## 현재 검증 범위

- Windows x64 환경에서 PyInstaller one-file GUI EXE 빌드 성공
- Python 문법 검사 및 출력 메서드 정적 검증
- 실제 Excel/PowerPoint COM 자동화는 Microsoft Office Desktop이 설치된 PC에서 최종 검증 필요

자세한 변경 이력은 [CHANGELOG.md](CHANGELOG.md)를 참고하세요.
