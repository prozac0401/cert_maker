# -*- coding: utf-8 -*-
"""
Win32 COM 기반 인증서 자동 제작 도구
- template .pptx 내부의 {컬럼명} 변수를 엑셀 첫 번째 시트의 헤더/값으로 치환
- 템플릿 파일명 또는 UI의 출력 파일명 패턴에 있는 {컬럼명}도 치환
- PPTX, PDF, PNG 출력 지원

실행 환경: Windows + Microsoft PowerPoint + Excel 설치 필요
"""
from __future__ import annotations

import os
import re
import sys
import shutil
import threading
import queue
import traceback
from dataclasses import dataclass
from datetime import datetime, date
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

APP_TITLE = "인증서 자동 제작 도구"
APP_VERSION = "1.2.3"

# PowerPoint COM constants
PP_SAVE_AS_OPENXML_PRESENTATION = 24  # ppSaveAsOpenXMLPresentation (.pptx)
PP_FIXED_FORMAT_TYPE_PDF = 2          # ppFixedFormatTypePDF
PP_FIXED_FORMAT_INTENT_PRINT = 2      # ppFixedFormatIntentPrint
PP_PRINT_HANDOUT_HORIZONTAL_FIRST = 2 # ppPrintHandoutHorizontalFirst
PP_PRINT_OUTPUT_SLIDES = 1            # ppPrintOutputSlides
PP_PRINT_ALL = 1                      # ppPrintAll
PP_WINDOW_NORMAL = 1                 # ppWindowNormal
PP_WINDOW_MINIMIZED = 2              # ppWindowMinimized
PP_ALERTS_NONE = 1
MSO_FALSE = 0
MSO_TRUE = -1
XL_UP = -4162
XL_TO_LEFT = -4159

TOKEN_RE = re.compile(r"\{([^{}]+)\}")
INVALID_FILENAME_RE = re.compile(r'[\\/:*?"<>|\r\n\t]+')


@dataclass
class JobSettings:
    template_path: Path
    data_path: Path
    output_dir: Path
    filename_pattern: str
    export_pptx: bool
    export_pdf: bool
    export_png: bool
    png_width: Optional[int]
    png_height: Optional[int]
    keep_powerpoint_visible: bool


def resource_path(relative_path: str) -> Path:
    """PyInstaller one-file EXE와 소스 실행 모두에서 리소스 경로를 찾습니다."""
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return Path(base) / relative_path
    return Path(__file__).resolve().parent / relative_path


def minimize_powerpoint_window(ppt_app, presentation=None) -> None:
    """PowerPoint에는 실제 DocumentWindow를 유지하되 화면에서는 최소화합니다.

    일부 PowerPoint/보안 환경은 WithWindow=False로 연 프레젠테이션에서 Save/Export
    계열 COM 호출이 실패합니다. 따라서 백그라운드 모드도 WithWindow=True로 연 뒤
    공식 WindowState=ppWindowMinimized를 사용합니다.
    """
    try:
        ppt_app.Visible = MSO_TRUE
    except Exception:
        pass

    if presentation is not None:
        try:
            if int(presentation.Windows.Count) > 0:
                presentation.Windows(1).WindowState = PP_WINDOW_MINIMIZED
        except Exception:
            pass
    try:
        ppt_app.WindowState = PP_WINDOW_MINIMIZED
    except Exception:
        try:
            ppt_app.ActiveWindow.WindowState = PP_WINDOW_MINIMIZED
        except Exception:
            pass


def get_pptx_aspect_ratio_com(pptx_path: Path) -> Optional[float]:
    """PowerPoint COM으로 템플릿을 열어 슬라이드 비율을 읽습니다.

    PPTX를 ZIP/XML로 직접 파싱하지 않습니다. 안정성을 위해 실제 DocumentWindow를
    생성(WithWindow=True)한 뒤 즉시 최소화합니다.
    """
    if os.name != "nt":
        return None

    ppt_app = None
    presentation = None
    pythoncom = None
    initialized = False
    try:
        import pythoncom as _pythoncom
        import win32com.client
        pythoncom = _pythoncom
        pythoncom.CoInitialize()
        initialized = True
        ppt_app = win32com.client.DispatchEx("PowerPoint.Application")
        ppt_app.Visible = MSO_TRUE
        presentation = ppt_app.Presentations.Open(str(pptx_path.resolve()), MSO_TRUE, MSO_FALSE, MSO_TRUE)
        minimize_powerpoint_window(ppt_app, presentation)
        slide_w = float(presentation.PageSetup.SlideWidth)
        slide_h = float(presentation.PageSetup.SlideHeight)
        if slide_w > 0 and slide_h > 0:
            return slide_w / slide_h
    except Exception:
        return None
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        if ppt_app is not None:
            try:
                ppt_app.Quit()
            except Exception:
                pass
        if initialized and pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
    return None


def format_cell_value(value) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return str(value)
    return str(value)


def sanitize_filename(name: str, fallback: str = "certificate") -> str:
    name = INVALID_FILENAME_RE.sub("_", name).strip()
    name = re.sub(r"\s+", " ", name).strip(" .")
    return name or fallback


def replace_tokens(text: str, row_map: Dict[str, str], leave_unknown: bool = True) -> str:
    def repl(match: re.Match) -> str:
        key = match.group(1).strip()
        if key in row_map:
            return row_map[key]
        return match.group(0) if leave_unknown else ""
    return TOKEN_RE.sub(repl, text)


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem, suffix = path.stem, path.suffix
    parent = path.parent
    idx = 2
    while True:
        candidate = parent / f"{stem}_{idx}{suffix}"
        if not candidate.exists():
            return candidate
        idx += 1


def _normalize_range_values(values, rows: int, cols: int):
    """Excel COM Range.Value 반환값을 항상 2차원 tuple 형태로 정규화합니다."""
    if rows == 1 and cols == 1:
        return ((values,),)
    if rows == 1:
        # 보통 ((a,b,c),) 이지만 환경 차이를 방어합니다.
        if isinstance(values, tuple) and values and isinstance(values[0], tuple):
            return values
        return (tuple(values),)
    if cols == 1:
        if isinstance(values, tuple) and values and isinstance(values[0], tuple):
            return values
        return tuple((v,) for v in values)
    return values


def load_excel_rows_com(xlsx_path: Path, excel_app, log=None) -> Tuple[List[str], List[Dict[str, str]]]:
    """Excel COM으로 첫 번째 시트의 1행을 헤더로 읽습니다.

    openpyxl/ZIP 파서를 사용하지 않으므로 Office/DRM 환경에서 Excel 자체가 열 수 있는
    파일이라면 동일한 Office COM 경로를 사용합니다.
    """
    workbook = None
    try:
        workbook = excel_app.Workbooks.Open(str(xlsx_path.resolve()), 0, True)
        worksheet = workbook.Worksheets(1)

        # UsedRange는 과거 서식만 남아 있어도 지나치게 커질 수 있으므로 사용하지 않습니다.
        # 실제 1행의 마지막 헤더 열을 찾고, 해당 헤더 열들 중 마지막 데이터 행을 계산합니다.
        last_col = int(worksheet.Cells(1, worksheet.Columns.Count).End(XL_TO_LEFT).Column)
        first_header = worksheet.Cells(1, 1).Value
        if last_col == 1 and (first_header is None or format_cell_value(first_header).strip() == ""):
            raise ValueError("엑셀 첫 번째 시트의 첫 줄에 헤더가 없습니다.")

        last_row = 1
        for col_idx in range(1, last_col + 1):
            col_last_row = int(worksheet.Cells(worksheet.Rows.Count, col_idx).End(XL_UP).Row)
            last_row = max(last_row, col_last_row)

        # 헤더는 반드시 실제 1행을 기준으로 합니다.
        header_range = worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(1, last_col))
        header_values = _normalize_range_values(header_range.Value, 1, last_col)[0]

        headers: List[str] = []
        seen = set()
        for idx, raw in enumerate(header_values, start=1):
            header = format_cell_value(raw).strip()
            if not header:
                header = f"Column{idx}"
            if header in seen:
                raise ValueError(f"중복 헤더가 있습니다: {header}")
            seen.add(header)
            headers.append(header)

        if last_row < 2:
            raise ValueError("엑셀 데이터 행이 없습니다. 2행부터 데이터를 입력하세요.")

        data_range = worksheet.Range(worksheet.Cells(2, 1), worksheet.Cells(last_row, last_col))
        raw_rows = _normalize_range_values(data_range.Value, last_row - 1, last_col)

        rows: List[Dict[str, str]] = []
        for excel_row in raw_rows:
            if excel_row is None:
                continue
            values = list(excel_row) if isinstance(excel_row, tuple) else [excel_row]
            if all(v is None for v in values):
                continue
            row_map: Dict[str, str] = {}
            for idx, header in enumerate(headers):
                value = values[idx] if idx < len(values) else None
                row_map[header] = format_cell_value(value)
            rows.append(row_map)

        if not rows:
            raise ValueError("엑셀 데이터 행이 없습니다. 2행부터 데이터를 입력하세요.")
        if log:
            log(f"Excel COM 읽기 완료: 첫 번째 시트 '{worksheet.Name}', {len(rows)}건")
        return headers, rows
    except Exception as e:
        raise RuntimeError(f"Excel COM으로 데이터 파일을 여는 중 오류가 발생했습니다: {e}") from e
    finally:
        if workbook is not None:
            try:
                workbook.Close(False)
            except Exception:
                pass


def replace_text_range(text_range, row_map: Dict[str, str]) -> None:
    # TextRange.Replace는 서식을 최대한 보존합니다. 실패 시 전체 텍스트 치환으로 fallback합니다.
    for key, value in row_map.items():
        token = "{" + key + "}"
        try:
            text_range.Replace(FindWhat=token, ReplaceWhat=value, MatchCase=False, WholeWords=False)
        except Exception:
            try:
                current = text_range.Text
                if token in current:
                    text_range.Text = current.replace(token, value)
            except Exception:
                pass


def replace_shape(shape, row_map: Dict[str, str]) -> None:
    # 일반 텍스트 상자 / 도형 텍스트
    try:
        if getattr(shape, "HasTextFrame", 0):
            text_frame = shape.TextFrame
            if getattr(text_frame, "HasText", 0):
                replace_text_range(text_frame.TextRange, row_map)
    except Exception:
        pass

    # 표 내부 셀
    try:
        if getattr(shape, "HasTable", 0):
            table = shape.Table
            for r in range(1, table.Rows.Count + 1):
                for c in range(1, table.Columns.Count + 1):
                    try:
                        cell_shape = table.Cell(r, c).Shape
                        if getattr(cell_shape, "HasTextFrame", 0) and getattr(cell_shape.TextFrame, "HasText", 0):
                            replace_text_range(cell_shape.TextFrame.TextRange, row_map)
                    except Exception:
                        pass
    except Exception:
        pass

    # 그룹 도형
    try:
        group_items = shape.GroupItems
        for idx in range(1, group_items.Count + 1):
            replace_shape(group_items.Item(idx), row_map)
    except Exception:
        pass


def replace_presentation(presentation, row_map: Dict[str, str]) -> None:
    for slide_idx in range(1, presentation.Slides.Count + 1):
        slide = presentation.Slides(slide_idx)
        for shape_idx in range(1, slide.Shapes.Count + 1):
            replace_shape(slide.Shapes(shape_idx), row_map)


def calc_png_size(presentation, width: Optional[int], height: Optional[int]) -> Tuple[int, int]:
    slide_w = float(presentation.PageSetup.SlideWidth)
    slide_h = float(presentation.PageSetup.SlideHeight)
    if not width and not height:
        width = 1920
        height = int(round(width * slide_h / slide_w))
    elif width and not height:
        height = int(round(width * slide_h / slide_w))
    elif height and not width:
        width = int(round(height * slide_w / slide_h))
    return int(width), int(height)


def parse_dimension(raw: str, field_name: str) -> Optional[int]:
    raw = raw.strip()
    if not raw:
        return None
    try:
        value = int(raw)
    except ValueError:
        raise ValueError(f"{field_name}은 숫자로 입력하세요.")
    if value < 100 or value > 10000:
        raise ValueError(f"{field_name}은 100~10000 사이로 입력하세요.")
    return value


def describe_com_error(exc: Exception) -> str:
    """pywintypes.com_error를 포함한 COM 예외를 로그 친화적으로 표시합니다."""
    parts = [str(exc)]
    hresult = getattr(exc, "hresult", None)
    if hresult is not None:
        try:
            parts.append(f"HRESULT=0x{(int(hresult) & 0xFFFFFFFF):08X}")
        except Exception:
            parts.append(f"HRESULT={hresult}")
    excepinfo = getattr(exc, "excepinfo", None)
    if excepinfo:
        try:
            source = excepinfo[1] if len(excepinfo) > 1 else None
            description = excepinfo[2] if len(excepinfo) > 2 else None
            scode = excepinfo[5] if len(excepinfo) > 5 else None
            if source:
                parts.append(f"source={source}")
            if description:
                parts.append(f"description={description}")
            if scode is not None:
                parts.append(f"scode={scode}")
        except Exception:
            pass
    return " / ".join(dict.fromkeys(parts))


def ensure_export_created(path: Path, label: str) -> None:
    """COM 호출이 예외 없이 끝났지만 파일이 생성되지 않은 경우도 실패로 판정합니다."""
    if not path.exists():
        raise RuntimeError(f"{label} COM 호출은 완료됐지만 출력 파일이 생성되지 않았습니다: {path}")
    try:
        if path.stat().st_size <= 0:
            raise RuntimeError(f"{label} 출력 파일의 크기가 0 byte입니다: {path}")
    except OSError as e:
        raise RuntimeError(f"{label} 출력 파일을 확인하지 못했습니다: {path} / {e}") from e


def export_pdf_com(presentation, pdf_path: Path) -> None:
    """PowerPoint의 고정 형식 내보내기 API로 PDF를 생성합니다.

    SaveAs(ppSaveAsPDF)를 사용하지 않습니다. 일부 Office/보안 환경에서 SaveAs PDF가
    실패하면서 이후 PNG 출력까지 중단되던 문제를 피하기 위함입니다.
    """
    # ExportAsFixedFormat의 PrintRange는 Office 버전에 따라 생략 시 Type mismatch가
    # 발생할 수 있어 Nothing(None)을 포함한 전체 주요 인수를 명시합니다.
    presentation.ExportAsFixedFormat(
        str(pdf_path.resolve()),
        PP_FIXED_FORMAT_TYPE_PDF,
        PP_FIXED_FORMAT_INTENT_PRINT,
        MSO_FALSE,
        PP_PRINT_HANDOUT_HORIZONTAL_FIRST,
        PP_PRINT_OUTPUT_SLIDES,
        MSO_FALSE,
        None,
        PP_PRINT_ALL,
        "",
        False,
        False,
        True,
        True,
        False,
    )
    ensure_export_created(pdf_path, "PDF")


def export_png_com(presentation, output_dir: Path, base_name: str, width: Optional[int], height: Optional[int], log) -> None:
    """각 슬라이드를 Slide.Export로 직접 PNG 파일로 내보냅니다."""
    png_w, png_h = calc_png_size(presentation, width, height)
    slide_count = int(presentation.Slides.Count)
    for slide_idx in range(1, slide_count + 1):
        if slide_count == 1:
            png_name = f"{base_name}.png"
        else:
            png_name = f"{base_name}_slide{slide_idx:02d}.png"
        png_path = unique_path(output_dir / png_name)
        presentation.Slides(slide_idx).Export(str(png_path.resolve()), "PNG", png_w, png_h)
        ensure_export_created(png_path, f"PNG slide {slide_idx}")
        log(f"  PNG 저장: {png_path.name} ({png_w}x{png_h}) [Slide.Export]")


def export_pptx_copy_com(presentation, pptx_path: Path) -> None:
    """열려 있는 프레젠테이션의 상태/경로를 바꾸지 않고 PPTX 복사본을 저장합니다."""
    presentation.SaveCopyAs(str(pptx_path.resolve()), PP_SAVE_AS_OPENXML_PRESENTATION, MSO_FALSE)
    ensure_export_created(pptx_path, "PPTX")


def export_for_row(ppt_app, settings: JobSettings, row_map: Dict[str, str], row_number: int, log) -> None:
    presentation = None
    format_errors: List[str] = []
    try:
        # 안정성 우선: 표시 여부와 관계없이 실제 DocumentWindow를 생성합니다.
        # WithWindow=False는 일부 Office/DRM 환경에서 Save/Export 예외를 재현하므로 사용하지 않습니다.
        try:
            presentation = ppt_app.Presentations.Open(
                str(settings.template_path.resolve()), MSO_TRUE, MSO_FALSE, MSO_TRUE
            )
            if settings.keep_powerpoint_visible:
                try:
                    ppt_app.WindowState = PP_WINDOW_NORMAL
                except Exception:
                    pass
            else:
                minimize_powerpoint_window(ppt_app, presentation)
                log("  PowerPoint 백그라운드 최소화 모드: WithWindow=True + ppWindowMinimized")
        except Exception as e:
            raise RuntimeError(f"PowerPoint COM으로 템플릿을 열지 못했습니다: {describe_com_error(e)}") from e
        replace_presentation(presentation, row_map)

        raw_base = replace_tokens(settings.filename_pattern, row_map, leave_unknown=False)
        base_name = sanitize_filename(raw_base, fallback=f"certificate_{row_number:04d}")
        output_dir = settings.output_dir.resolve()

        # 각 출력 형식을 완전히 독립시킵니다. PDF가 실패해도 PNG/PPTX는 계속 시도합니다.
        # 또한 SaveAs로 프레젠테이션의 현재 파일 상태를 바꾸지 않도록 PPTX는 SaveCopyAs를 사용합니다.
        if settings.export_pdf:
            pdf_path = unique_path(output_dir / f"{base_name}.pdf")
            try:
                log(f"  PDF 출력 시작: {pdf_path.name} [ExportAsFixedFormat]")
                export_pdf_com(presentation, pdf_path)
                log(f"  PDF 저장: {pdf_path.name} [ExportAsFixedFormat]")
            except Exception as e:
                detail = describe_com_error(e)
                format_errors.append(f"PDF: {detail}")
                log(f"  PDF 오류: {detail}")

        if settings.export_png:
            try:
                log("  PNG 출력 시작 [Slide.Export]")
                export_png_com(
                    presentation, output_dir, base_name,
                    settings.png_width, settings.png_height, log
                )
            except Exception as e:
                detail = describe_com_error(e)
                format_errors.append(f"PNG: {detail}")
                log(f"  PNG 오류: {detail}")

        if settings.export_pptx:
            pptx_path = unique_path(output_dir / f"{base_name}.pptx")
            try:
                log(f"  PPTX 출력 시작: {pptx_path.name} [SaveCopyAs]")
                export_pptx_copy_com(presentation, pptx_path)
                log(f"  PPTX 저장: {pptx_path.name} [SaveCopyAs]")
            except Exception as e:
                detail = describe_com_error(e)
                format_errors.append(f"PPTX: {detail}")
                log(f"  PPTX 오류: {detail}")

        if format_errors:
            raise RuntimeError("일부 출력 형식 생성 실패 | " + " | ".join(format_errors))
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass


def run_job(settings: JobSettings, progress_cb, log_cb, done_cb) -> None:
    pythoncom = None
    initialized = False
    excel_app = None
    ppt_app = None
    try:
        try:
            import pythoncom as _pythoncom
            import win32com.client
        except Exception as e:
            raise RuntimeError(
                "pywin32/win32com을 사용할 수 없습니다. Windows용 EXE 빌드 또는 pywin32 설치 상태를 확인하세요."
            ) from e

        pythoncom = _pythoncom
        pythoncom.CoInitialize()
        initialized = True

        log_cb(f"실행 환경: Windows / frozen={bool(getattr(sys, 'frozen', False))}")
        settings.output_dir.mkdir(parents=True, exist_ok=True)

        # Excel도 COM으로 읽습니다. openpyxl/ZIP 파서는 사용하지 않습니다.
        log_cb("Excel COM 시작 중...")
        try:
            excel_app = win32com.client.DispatchEx("Excel.Application")
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            try:
                log_cb(f"Excel COM 연결 완료: version={excel_app.Version}")
            except Exception:
                log_cb("Excel COM 연결 완료")
            headers, rows = load_excel_rows_com(settings.data_path, excel_app, log_cb)
            log_cb(f"엑셀 로드 완료: {len(rows)}건 / 헤더: {', '.join(headers)}")
        except Exception as e:
            raise RuntimeError(
                "Excel COM 처리에 실패했습니다. Microsoft Excel 설치 여부와 데이터 파일이 Excel에서 정상적으로 열리는지 확인하세요. "
                f"원인: {e}"
            ) from e
        finally:
            if excel_app is not None:
                try:
                    excel_app.Quit()
                except Exception:
                    pass
                excel_app = None

        log_cb("PowerPoint COM 시작 중...")
        try:
            ppt_app = win32com.client.DispatchEx("PowerPoint.Application")
            # PowerPoint 내보내기 안정성을 위해 항상 실제 Application/DocumentWindow를 사용합니다.
            # 체크 해제 시에는 숨기는 대신 공식 ppWindowMinimized 상태로 즉시 최소화합니다.
            ppt_app.Visible = MSO_TRUE
            try:
                ppt_app.DisplayAlerts = PP_ALERTS_NONE
            except Exception:
                pass
            try:
                ppt_version = str(ppt_app.Version)
            except Exception:
                ppt_version = "unknown"
            log_cb(
                f"PowerPoint COM 연결 완료: version={ppt_version} / "
                + ("창 표시 모드" if settings.keep_powerpoint_visible else "백그라운드 최소화 모드(WithWindow=True + ppWindowMinimized)")
            )
        except Exception as e:
            raise RuntimeError(
                "PowerPoint COM 시작에 실패했습니다. Microsoft PowerPoint 설치/등록 상태를 확인하세요. "
                f"원인: {e}"
            ) from e

        success_count = 0
        error_count = 0
        total = len(rows)
        progress_cb(0, total, "준비 완료", "대기")

        for idx, row_map in enumerate(rows, start=1):
            name_hint = row_map.get("이름", f"{idx}행") or f"{idx}행"
            progress_cb(idx - 1, total, name_hint, "처리 중")
            log_cb(f"[{idx}/{total}] 처리 시작: {name_hint}")
            try:
                export_for_row(ppt_app, settings, row_map, idx, log_cb)
                success_count += 1
                log_cb(f"[{idx}/{total}] 완료")
                progress_cb(idx, total, name_hint, "완료")
            except Exception as row_error:
                error_count += 1
                log_cb(f"[{idx}/{total}] 오류: {row_error}")
                progress_cb(idx, total, name_hint, "오류")

        if error_count:
            done_cb(False, f"완료: 성공 {success_count}건, 오류 {error_count}건")
        else:
            done_cb(True, f"완료: {success_count}건 생성")
    except Exception as e:
        log_cb(traceback.format_exc())
        done_cb(False, str(e))
    finally:
        if excel_app is not None:
            try:
                excel_app.Quit()
            except Exception:
                pass
        if ppt_app is not None:
            try:
                ppt_app.Quit()
            except Exception:
                pass
        if initialized and pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


class CertificateMakerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE} v{APP_VERSION}")
        self.geometry("900x760")
        self.minsize(840, 700)

        self.ui_queue: queue.Queue = queue.Queue()
        self.worker: Optional[threading.Thread] = None

        self.template_var = tk.StringVar()
        self.data_var = tk.StringVar()
        self.output_var = tk.StringVar(value=str(Path.cwd() / "output"))
        self.filename_var = tk.StringVar(value="인증서_{이름}_{과정명}")
        self.export_pptx_var = tk.BooleanVar(value=True)
        self.export_pdf_var = tk.BooleanVar(value=True)
        self.export_png_var = tk.BooleanVar(value=False)
        # 세로형(A4 계열) 기본값. 슬라이더는 400~5000px 범위입니다.
        self.png_width_var = tk.DoubleVar(value=1240)
        self.png_height_var = tk.DoubleVar(value=1754)
        self.png_lock_ratio_var = tk.BooleanVar(value=True)
        self.png_ratio = 1240 / 1754
        self._syncing_png_size = False
        # PowerPoint 창 표시는 기본 해제. 안정성을 위해 실제 창은 생성하고 즉시 최소화합니다.
        self.visible_var = tk.BooleanVar(value=False)

        self._build_ui()
        self.after(100, self._poll_queue)

    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 6}
        root = ttk.Frame(self)
        root.pack(fill="both", expand=True)

        header = ttk.Frame(root)
        header.pack(fill="x", padx=14, pady=(12, 8))
        ttk.Label(header, text="인증서 자동 제작", font=("맑은 고딕", 16, "bold")).pack(anchor="w")
        ttk.Label(
            header,
            text="PPTX 템플릿의 {헤더명} 변수를 엑셀 첫 번째 시트 데이터로 치환하여 PPTX/PDF/PNG 파일을 생성합니다.",
            foreground="#555555",
        ).pack(anchor="w", pady=(4, 0))

        notice = ttk.LabelFrame(root, text="작업 전 유의사항")
        notice.pack(fill="x", padx=14, pady=(0, 8))
        ttk.Label(
            notice,
            text="Windows + Microsoft PowerPoint + Excel이 필요합니다. 실행 중에는 템플릿/출력 파일을 열어두지 말고, PowerPoint 창 표시는 기본 해제되어 있으며 백그라운드 모드에서는 실제 프레젠테이션 창을 연 뒤 즉시 최소화합니다. 창 표시를 켠 경우 작업 중 PowerPoint 조작은 피해주세요. 파일명 중복은 자동으로 _2, _3 형식으로 회피합니다.",
            wraplength=820,
            foreground="#7A3E00",
        ).pack(anchor="w", padx=10, pady=7)

        form = ttk.LabelFrame(root, text="입력 파일")
        form.pack(fill="x", padx=14, pady=8)

        self._file_row(form, "템플릿 PPTX", self.template_var, self._browse_template, row=0)
        self._file_row(form, "데이터 XLSX", self.data_var, self._browse_data, row=1)
        self._file_row(form, "출력 폴더", self.output_var, self._browse_output, row=2)

        ttk.Label(form, text="출력 파일명 패턴").grid(row=3, column=0, sticky="w", **pad)
        ttk.Entry(form, textvariable=self.filename_var).grid(row=3, column=1, sticky="ew", **pad)
        ttk.Label(form, text="예: 수료증_{이름}_{과정명}", foreground="#666666").grid(row=3, column=2, sticky="w", **pad)
        form.columnconfigure(1, weight=1)

        options = ttk.LabelFrame(root, text="출력 형식")
        options.pack(fill="x", padx=14, pady=8)
        ttk.Checkbutton(options, text="PPTX", variable=self.export_pptx_var).grid(row=0, column=0, sticky="w", **pad)
        ttk.Checkbutton(options, text="PDF", variable=self.export_pdf_var).grid(row=0, column=1, sticky="w", **pad)
        ttk.Checkbutton(options, text="PNG", variable=self.export_png_var).grid(row=0, column=2, sticky="w", **pad)
        ttk.Checkbutton(
            options, text="PowerPoint 창 표시", variable=self.visible_var
        ).grid(row=0, column=3, sticky="w", padx=14, pady=6)
        ttk.Checkbutton(
            options, text="PNG 비율 고정", variable=self.png_lock_ratio_var, command=self._toggle_png_ratio_lock
        ).grid(row=0, column=4, sticky="w", padx=14, pady=6)

        ttk.Label(options, text="PNG 너비").grid(row=1, column=0, sticky="w", padx=(10, 4), pady=(6, 2))
        self.png_width_scale = ttk.Scale(
            options, from_=400, to=5000, variable=self.png_width_var, command=self._on_png_width_changed
        )
        self.png_width_scale.grid(row=1, column=1, columnspan=3, sticky="ew", padx=4, pady=(6, 2))
        self.png_width_value = ttk.Label(options, text="1240 px", width=10, anchor="e")
        self.png_width_value.grid(row=1, column=4, sticky="e", padx=10, pady=(6, 2))

        ttk.Label(options, text="PNG 높이").grid(row=2, column=0, sticky="w", padx=(10, 4), pady=(2, 6))
        self.png_height_scale = ttk.Scale(
            options, from_=400, to=5000, variable=self.png_height_var, command=self._on_png_height_changed
        )
        self.png_height_scale.grid(row=2, column=1, columnspan=3, sticky="ew", padx=4, pady=(2, 6))
        self.png_height_value = ttk.Label(options, text="1754 px", width=10, anchor="e")
        self.png_height_value.grid(row=2, column=4, sticky="e", padx=10, pady=(2, 6))

        self.png_ratio_label = ttk.Label(
            options, text="비율 고정: 템플릿 슬라이드 비율", foreground="#666666"
        )
        self.png_ratio_label.grid(row=3, column=0, columnspan=5, sticky="w", padx=10, pady=(0, 8))
        for c in range(1, 4):
            options.columnconfigure(c, weight=1)

        actions = ttk.Frame(root)
        actions.pack(fill="x", padx=14, pady=8)
        ttk.Button(actions, text="예제 파일 만들기", command=self._copy_samples).pack(side="left")
        ttk.Button(actions, text="실행 시작", command=self._start_job).pack(side="right")

        progress_frame = ttk.LabelFrame(root, text="진행 상황")
        progress_frame.pack(fill="both", expand=True, padx=14, pady=(8, 14))

        self.current_item_var = tk.StringVar(value="현재 처리 중: 대기 중")
        self.progress_count_var = tk.StringVar(value="0 / 0건 (0%)")
        progress_header = ttk.Frame(progress_frame)
        progress_header.pack(fill="x", padx=10, pady=(10, 4))
        ttk.Label(progress_header, textvariable=self.current_item_var, font=("맑은 고딕", 10, "bold")).pack(side="left")
        ttk.Label(progress_header, textvariable=self.progress_count_var).pack(side="right")

        self.progress = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=(0, 6))
        self.log_text = ScrolledText(progress_frame, height=14, font=("Consolas", 10))
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self._log("준비 완료. 예제 파일 만들기로 세로형 인증서 템플릿과 데이터 샘플을 바로 테스트할 수 있습니다.")

    def _update_png_labels(self) -> None:
        self.png_width_value.configure(text=f"{int(round(self.png_width_var.get()))} px")
        self.png_height_value.configure(text=f"{int(round(self.png_height_var.get()))} px")

    def _on_png_width_changed(self, _value=None) -> None:
        if self._syncing_png_size:
            return
        self._syncing_png_size = True
        try:
            width = max(400, min(5000, int(round(self.png_width_var.get()))))
            if self.png_lock_ratio_var.get() and self.png_ratio > 0:
                height = int(round(width / self.png_ratio))
                height = max(400, min(5000, height))
                self.png_height_var.set(height)
            self._update_png_labels()
        finally:
            self._syncing_png_size = False

    def _on_png_height_changed(self, _value=None) -> None:
        if self._syncing_png_size:
            return
        self._syncing_png_size = True
        try:
            height = max(400, min(5000, int(round(self.png_height_var.get()))))
            if self.png_lock_ratio_var.get() and self.png_ratio > 0:
                width = int(round(height * self.png_ratio))
                width = max(400, min(5000, width))
                self.png_width_var.set(width)
            self._update_png_labels()
        finally:
            self._syncing_png_size = False

    def _toggle_png_ratio_lock(self) -> None:
        if self.png_lock_ratio_var.get():
            # 고정 활성화 시 현재 너비를 기준으로 템플릿 비율에 맞춰 높이를 즉시 동기화합니다.
            self._on_png_width_changed()
        self._update_png_labels()

    def _apply_template_ratio(self, pptx_path: Path) -> None:
        ratio = get_pptx_aspect_ratio_com(pptx_path)
        if not ratio or ratio <= 0:
            self.png_ratio_label.configure(text="비율 고정: PowerPoint COM에서 비율을 읽지 못했습니다. 현재 비율을 유지합니다.")
            return
        self.png_ratio = ratio
        orientation = "세로" if ratio < 1 else ("가로" if ratio > 1 else "정사각")
        self.png_ratio_label.configure(text=f"비율 고정: 템플릿 슬라이드 비율 ({orientation}, {ratio:.4f}:1)")
        if self.png_lock_ratio_var.get():
            self._on_png_width_changed()

    def _file_row(self, parent, label: str, var: tk.StringVar, command, row: int) -> None:
        pad = {"padx": 10, "pady": 6}
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", **pad)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", **pad)
        ttk.Button(parent, text="찾기", command=command).grid(row=row, column=2, sticky="e", **pad)

    def _browse_template(self) -> None:
        path = filedialog.askopenfilename(title="템플릿 PPTX 선택", filetypes=[("PowerPoint", "*.pptx"), ("All files", "*.*")])
        if path:
            self.template_var.set(path)
            stem = Path(path).stem
            if stem:
                self.filename_var.set(stem)
            self._apply_template_ratio(Path(path))

    def _browse_data(self) -> None:
        path = filedialog.askopenfilename(title="데이터 XLSX 선택", filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")])
        if path:
            self.data_var.set(path)

    def _browse_output(self) -> None:
        path = filedialog.askdirectory(title="출력 폴더 선택")
        if path:
            self.output_var.set(path)

    def _copy_samples(self) -> None:
        dest = self.output_var.get().strip()
        if not dest:
            dest = filedialog.askdirectory(title="예제 파일을 복사할 폴더 선택")
            if not dest:
                return
            self.output_var.set(dest)
        dest_path = Path(dest)
        dest_path.mkdir(parents=True, exist_ok=True)

        sample_dir = resource_path("samples")
        # PyInstaller/압축 도구에서 한글 리소스 파일명이 깨지는 환경을 피하기 위해
        # EXE 내부 샘플 리소스명은 ASCII로 고정합니다.
        sample_files = [
            sample_dir / "template.pptx",
            sample_dir / "data.xlsx",
        ]
        copied = []
        for src in sample_files:
            if not src.exists():
                messagebox.showerror("예제 파일 없음", f"예제 파일을 찾을 수 없습니다:\n{src}")
                return
            target = unique_path(dest_path / src.name)
            shutil.copy2(src, target)
            copied.append(target)

        self.template_var.set(str(copied[0]))
        self.data_var.set(str(copied[1]))
        self.filename_var.set(copied[0].stem)
        self._apply_template_ratio(copied[0])
        self._log(f"예제 파일 생성 완료: {dest_path}")
        messagebox.showinfo("완료", "세로형 예제 PPTX와 XLSX를 출력 폴더에 만들었습니다.")

    def _validate_settings(self) -> JobSettings:
        template_path = Path(self.template_var.get().strip())
        data_path = Path(self.data_var.get().strip())
        output_dir = Path(self.output_var.get().strip())
        filename_pattern = self.filename_var.get().strip()

        if not template_path.exists() or template_path.suffix.lower() != ".pptx":
            raise ValueError("템플릿 PPTX 파일을 선택하세요.")
        if not data_path.exists() or data_path.suffix.lower() != ".xlsx":
            raise ValueError("데이터 XLSX 파일을 선택하세요.")
        if not filename_pattern:
            filename_pattern = template_path.stem
        if not (self.export_pptx_var.get() or self.export_pdf_var.get() or self.export_png_var.get()):
            raise ValueError("PPTX, PDF, PNG 중 하나 이상을 선택하세요.")

        png_width = int(round(self.png_width_var.get())) if self.export_png_var.get() else None
        png_height = int(round(self.png_height_var.get())) if self.export_png_var.get() else None
        if self.export_png_var.get():
            if not 400 <= png_width <= 5000 or not 400 <= png_height <= 5000:
                raise ValueError("PNG 너비/높이는 400~5000px 범위로 설정하세요.")

        return JobSettings(
            template_path=template_path,
            data_path=data_path,
            output_dir=output_dir,
            filename_pattern=filename_pattern,
            export_pptx=self.export_pptx_var.get(),
            export_pdf=self.export_pdf_var.get(),
            export_png=self.export_png_var.get(),
            png_width=png_width,
            png_height=png_height,
            keep_powerpoint_visible=self.visible_var.get(),
        )

    def _start_job(self) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showwarning("실행 중", "이미 작업이 실행 중입니다.")
            return
        try:
            settings = self._validate_settings()
        except Exception as e:
            messagebox.showerror("입력 확인", str(e))
            return

        self.progress["value"] = 0
        self.progress["maximum"] = 100
        self.current_item_var.set("현재 처리 중: 준비 중")
        self.progress_count_var.set("0 / 0건 (0%)")
        self._log("=== 작업 시작 ===")
        self._log(f"템플릿: {settings.template_path}")
        self._log(f"데이터: {settings.data_path}")
        self._log(f"출력 폴더: {settings.output_dir}")

        self.worker = threading.Thread(
            target=run_job,
            args=(settings, self._thread_progress, self._thread_log, self._thread_done),
            daemon=True,
        )
        self.worker.start()

    def _thread_log(self, text: str) -> None:
        self.ui_queue.put(("log", text))

    def _thread_progress(self, current: int, total: int, name: str, state: str) -> None:
        self.ui_queue.put(("progress", current, total, name, state))

    def _thread_done(self, ok: bool, message: str) -> None:
        self.ui_queue.put(("done", ok, message))

    def _poll_queue(self) -> None:
        try:
            while True:
                item = self.ui_queue.get_nowait()
                kind = item[0]
                if kind == "log":
                    self._log(item[1])
                elif kind == "progress":
                    _, current, total, name, state = item
                    self.progress["maximum"] = max(total, 1)
                    self.progress["value"] = current
                    percent = int(round((current / total) * 100)) if total else 0
                    self.progress_count_var.set(f"{current} / {total}건 ({percent}%)")
                    self.current_item_var.set(f"현재 처리 중: {name} [{state}]")
                elif kind == "done":
                    _, ok, msg = item
                    if ok:
                        self.current_item_var.set("현재 처리 중: 작업 완료")
                    else:
                        self.current_item_var.set("현재 처리 중: 확인 필요")
                    self._log(msg)
                    self._log("=== 작업 종료 ===")
                    if ok:
                        messagebox.showinfo("완료", msg)
                    else:
                        messagebox.showwarning("확인 필요", msg)
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _log(self, text: str) -> None:
        now = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{now}] {text}\n")
        self.log_text.see("end")


def main() -> None:
    app = CertificateMakerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
