"""
무사고전환 데이터/매핑정보/최대전환매핑 입력 파일 하나로, 아래 3단계를 한 번에
실행하는 통합 실행기(v6).

  1) build_report_v6   : 데이터 시트에 경과전환판매플랜코드(매핑 실패 시 최대전환
                         매핑으로 폴백)·전환여부·최대전환플랜코드·년수비교 등을
                         채우고, 세부보고/합산보고 x 부문별/상품별 레이아웃을
                         수식으로 생성
  2) recalc_util       : LibreOffice로 전체 수식 재계산 (openpyxl은 계산된 값을
                         남기지 않으므로 필수 단계)
  3) split_report_v6   : 데이터 파일(수식 유지)과 보고 파일(레이아웃 4종 값만)로
                         분리 (분리 과정에서 데이터 파일 쪽 수식 캐시값이 openpyxl에
                         의해 지워지므로, 분리 후 데이터 파일을 한 번 더 재계산)

요구사항: LibreOffice(soffice)가 설치되어 있어야 함
  - Ubuntu/Debian: sudo apt-get install libreoffice-calc
  - macOS: brew install --cask libreoffice

사용법:
  python scripts/run_pipeline_v6.py <입력.xlsx> <데이터출력.xlsx> <보고출력.xlsx> [--timeout 초] [--keep-combined]

입력 파일은 시트 3개를 담고 있어야 한다:
  - 데이터 시트: '수금부문명' 헤더 보유
  - 매핑정보 시트: '*무사고기간' 헤더 보유
  - 최대전환매핑 시트: '최대전환년수' 헤더 보유
"""
import argparse
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import build_report_v6 as builder
import recalc_util
from split_report_v6 import split


def run(input_path, data_out_path, report_out_path, timeout=900, keep_combined=False):
    combined_fd, combined_path = tempfile.mkstemp(suffix='.xlsx', prefix='musago_combined_v6_')
    os.close(combined_fd)

    try:
        builder.build(input_path, combined_path)

        print(f"[재계산] LibreOffice로 전체 수식 재계산 중 (최대 {timeout}초)...")
        result = recalc_util.recalc(combined_path, timeout=timeout)
        if 'error' in result:
            print(f"  ! 재계산 실패: {result['error']}", file=sys.stderr)
            sys.exit(1)
        print(f"  - 상태: {result['status']}, 수식 오류: {result['total_errors']}건")
        if result['total_errors']:
            print("  ! 오류 셀 목록(최대 50개):")
            for c in result['error_cells']:
                print(f"      {c}")

        print("[분리] 데이터 파일 / 보고 파일 분리 중...")
        data_ok = split(combined_path, data_out_path, report_out_path, timeout=timeout)

        print(f"  - 데이터 파일 -> {data_out_path}")
        print(f"  - 보고 파일   -> {report_out_path}")
        print("완료.")
        return result['total_errors'] == 0 and data_ok
    finally:
        if keep_combined:
            print(f"(--keep-combined) 결합 중간 파일 보존됨: {combined_path}")
        else:
            try:
                os.remove(combined_path)
            except OSError:
                pass


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument('input', help='입력 파일 (데이터 시트 + 매핑정보 시트 + 최대전환매핑 시트)')
    p.add_argument('data_out', help='데이터 파일 출력 경로')
    p.add_argument('report_out', help='보고 파일 출력 경로')
    p.add_argument('--timeout', type=int, default=900, help='LibreOffice 재계산 제한시간(초), 기본 900')
    p.add_argument('--keep-combined', action='store_true', help='중간 결합 파일을 지우지 않고 보존')
    args = p.parse_args()

    ok = run(args.input, args.data_out, args.report_out,
             timeout=args.timeout, keep_combined=args.keep_combined)
    sys.exit(0 if ok else 2)


if __name__ == '__main__':
    main()
