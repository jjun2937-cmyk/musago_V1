"""
build_report_v6.py로 만들고 recalc로 재계산까지 끝낸 결합 파일을
"데이터 파일"(원본데이터+매핑정보+최대전환매핑+파생컬럼, 수식 유지)과
"보고 파일"(세부보고/합산보고 x 부문별/상품별, 값만)로 분리.

주의(v4에서 발견된 버그, v5/v6에서도 동일하게 적용): 데이터 파일을 만들 때
openpyxl로 결합 파일을 data_only=False로 열어 레이아웃 시트만 지우고 다시
저장하면, openpyxl은 수식 셀의 캐시된 계산값을 보존하지 못하고 수식 문자열만
남긴 채 값을 전부 비워버린다. 그래서 데이터 파일을 저장한 뒤 반드시 한 번 더
recalc_util로 재계산해서 캐시값을 복원해야 한다.

사용법: python split_report_v6.py <재계산된_결합.xlsx> <데이터출력.xlsx> <보고출력.xlsx> [--timeout 초]
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import openpyxl
from openpyxl.styles import Font
import recalc_util

LAYOUT_SHEET_NAMES = [
    '세부보고_부문별', '합산보고_부문별',
    '세부보고_상품별', '합산보고_상품별',
]


FONT_NAME = '맑은 고딕'


def copy_values_only(value_ws, style_ws, dst_wb, sheet_name):
    """value_ws(재계산된 값)에서 값을, style_ws(재계산 전 원본, 서식 보존)에서
    서식/병합/열너비/행높이를 가져와 합친다.

    LibreOffice가 대용량 파일(11만 행 규모)을 재계산/재저장하는 과정에서 셀
    서식(배경색·테두리·정렬)을 시트 전체에 걸쳐 초기화해버리는 현상이
    관측되어(작은 샘플에서는 재현되지 않음, 값/병합/열너비는 멀쩡함), 재계산
    자체는 값 캐시를 얻는 용도로만 신뢰하고 서식은 재계산을 타지 않은 원본
    파일에서 가져오도록 분리했다."""
    dst_ws = dst_wb.create_sheet(sheet_name)
    for row in style_ws.iter_rows():
        for style_cell in row:
            value_cell = value_ws.cell(row=style_cell.row, column=style_cell.column)
            # 값이 없어도(예: C열 스파인 배경색만 있는 빈 칸) 서식이 있으면
            # 여전히 옮겨 담아야 한다 - 값 기준으로만 건너뛰면 배경색이 없는
            # 셀처럼 사라져버린다.
            if value_cell.value is None and not style_cell.has_style:
                continue
            new_cell = dst_ws.cell(row=style_cell.row, column=style_cell.column,
                                    value=value_cell.value)
            if style_cell.has_style:
                # LibreOffice가 재계산/저장하면서 '맑은 고딕'이 시스템에 없으면 자체
                # 대체 폰트(예: WenQuanYi Zen Hei)로 폰트명을 바꿔 써버리는 문제가
                # 있어서, 크기/굵기/색상은 그대로 두고 이름만 강제로 복원한다.
                f = style_cell.font
                new_cell.font = Font(name=FONT_NAME, size=f.size, bold=f.bold,
                                      italic=f.italic, color=f.color)
                new_cell.fill = style_cell.fill.copy()
                new_cell.border = style_cell.border.copy()
                new_cell.alignment = style_cell.alignment.copy()
                new_cell.number_format = style_cell.number_format
    for coord, dim in style_ws.column_dimensions.items():
        if dim.width:
            dst_ws.column_dimensions[coord].width = dim.width
    for r, dim in style_ws.row_dimensions.items():
        if dim.height:
            dst_ws.row_dimensions[r].height = dim.height
    for merged_range in style_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(merged_range))
    dst_ws.sheet_view.showGridLines = style_ws.sheet_view.showGridLines
    return dst_ws


def split(combined_path, data_out_path, report_out_path, timeout=600, style_source_path=None):
    # 1) 보고 파일: 값은 재계산된 결합 파일(data_only=True)에서, 서식은
    # 재계산을 타지 않은 원본(style_source_path, 없으면 결합 파일 자체)에서
    # 가져와 새 워크북에 합쳐 담는다.
    wb_values = openpyxl.load_workbook(combined_path, data_only=True)
    wb_styles = (openpyxl.load_workbook(style_source_path, data_only=False)
                 if style_source_path else openpyxl.load_workbook(combined_path, data_only=False))
    report_wb = openpyxl.Workbook()
    report_wb.remove(report_wb.active)
    for name in LAYOUT_SHEET_NAMES:
        copy_values_only(wb_values[name], wb_styles[name], report_wb, name)
    report_wb.save(report_out_path)
    wb_values.close()
    wb_styles.close()
    print(f"보고 파일(값만, 레이아웃 {len(LAYOUT_SHEET_NAMES)}종) 저장 완료 -> {report_out_path}")

    # 2) 데이터 파일: 수식 유지된 원본에서 레이아웃 시트만 제거하고 저장
    wb_formulas = openpyxl.load_workbook(combined_path, data_only=False)
    for name in LAYOUT_SHEET_NAMES:
        if name in wb_formulas.sheetnames:
            del wb_formulas[name]
    wb_formulas.save(data_out_path)
    wb_formulas.close()

    # 2-1) *** 중요 *** 위 저장 과정에서 남아있던 수식(파생컬럼, 매핑정보 키값)의
    # 캐시된 계산값이 openpyxl에 의해 지워졌으므로, 다시 한 번 재계산해서 복원한다.
    print("데이터 파일 재계산 중 (분리 과정에서 지워진 수식 캐시값 복원)...")
    result = recalc_util.recalc(data_out_path, timeout=timeout)
    if 'error' in result:
        print(f"  ! 데이터 파일 재계산 실패: {result['error']}", file=sys.stderr)
        sys.exit(1)
    print(f"  - 상태: {result['status']}, 수식 오류: {result['total_errors']}건")
    print(f"데이터 파일(원본데이터+매핑정보+최대전환매핑, 수식 유지) 저장 완료 -> {data_out_path}")
    return result['total_errors'] == 0


if __name__ == '__main__':
    if len(sys.argv) < 4:
        print("Usage: python split_report_v6.py <재계산된_결합.xlsx> <데이터출력.xlsx> <보고출력.xlsx> [timeout]")
        sys.exit(1)
    timeout = int(sys.argv[4]) if len(sys.argv) > 4 else 600
    ok = split(sys.argv[1], sys.argv[2], sys.argv[3], timeout=timeout)
    sys.exit(0 if ok else 2)
