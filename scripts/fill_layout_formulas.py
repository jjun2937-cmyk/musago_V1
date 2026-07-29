"""
결과레이아웃 시트 (26.4월 블록, 6~9행: 개인/전략/신사업/법인) 수식 채우기

레이아웃 시트 정의 (업무정의 시트 기준, 개념상 A~I를 실제 시트열 B,C,D,E,F,G,H,I,K에 매핑):
  B 전체계약건수
  C 사고有            : V열(사고구분코드)이 '01' 또는 '02'
  D = B - C           (기존 수식, 손대지 않음)
  E 기존전환완료       : Z열이 '기전환'
  F = D - E           (기존 수식, 손대지 않음)
  G 기존미전환         : V not in(01,02), Z<>기전환, Y=FALSE
  H = F - G           (기존 수식, 손대지 않음)
  I 전환완료           : V not in(01,02), Z<>기전환, Y=TRUE, X=TRUE
  K = H - I           (기존 수식, 손대지 않음)
  O~U 현장활동확인(001~007): I의 여집합인 "전환미완료"
      (V not in(01,02), Z<>기전환, Y=TRUE, X=FALSE) 대상 중 전환처리구분코드별 건수

주의 (엑셀 COUNTIFS 조건문자열 버그):
  V열 값 '01'/'02', U열 값 '001'~'007'은 텍스트지만 숫자로 보이는 문자열입니다.
  COUNTIFS(range, "<>01") 처럼 조건문자열로 비교하면 실제 엑셀은 "01"을 숫자 1로
  해석해버려서, 텍스트 셀과는 절대 같을 수 없으므로 "<>1" 조건이 모든 행에서
  TRUE가 되어(=필터가 무력화되어) 잘못된 결과가 나옵니다. (LibreOffice는 텍스트로
  올바르게 비교하기 때문에 이 문제가 재현되지 않아 처음에 놓쳤던 버그입니다.)
  그래서 V/U열과 비교하는 C, G, I, O~U는 COUNTIFS 대신 SUMPRODUCT 배열비교로
  작성합니다. SUMPRODUCT 안의 "=" / "<>"는 조건문자열 파서를 거치지 않고 셀 값을
  있는 그대로(텍스트면 텍스트로) 비교하므로 이 문제가 없습니다.
  B(수금부문명 비교), E(Z="기전환" 비교)는 숫자로 오인될 값이 아니라 COUNTIFS 그대로 둡니다.

사용법: python fill_layout_formulas.py <입력.xlsx> <출력.xlsx>
"""
import sys
import openpyxl

CODES = {
    'O': '001',  # 연락두절 등
    'P': '002',  # 사고고지
    'Q': '003',  # 전환의사無
    'R': '004',  # 압류계약
    'S': '005',  # 전환예정
    'T': '006',  # 기타
    'U': '007',  # ARS 거부
}

GROUP_ROWS = (6, 7, 8, 9)  # 개인, 전략, 신사업, 법인


def fill(input_path, output_path):
    wb = openpyxl.load_workbook(input_path, data_only=False)
    ws = wb['결과레이아웃']
    data_last = wb['데이터'].max_row

    def rng(col):
        return f"데이터!${col}$2:${col}${data_last}"

    for r in GROUP_ROWS:
        grp = f"$A{r}"
        J, V, Z, Y, X, U = rng("J"), rng("V"), rng("Z"), rng("Y"), rng("X"), rng("U")

        ws[f'B{r}'] = f'=COUNTIFS({J},{grp})'

        ws[f'C{r}'] = f'=SUMPRODUCT(({J}={grp})*(({V}="01")+({V}="02")))'

        ws[f'E{r}'] = f'=COUNTIFS({J},{grp},{Z},"기전환")'

        ws[f'G{r}'] = (
            f'=SUMPRODUCT(({J}={grp})*({V}<>"01")*({V}<>"02")*({Z}<>"기전환")*({Y}=FALSE))'
        )

        ws[f'I{r}'] = (
            f'=SUMPRODUCT(({J}={grp})*({V}<>"01")*({V}<>"02")*({Z}<>"기전환")*'
            f'({Y}=TRUE)*({X}=TRUE))'
        )

        for col, code in CODES.items():
            ws[f'{col}{r}'] = (
                f'=SUMPRODUCT(({J}={grp})*({V}<>"01")*({V}<>"02")*({Z}<>"기전환")*'
                f'({Y}=TRUE)*({X}=FALSE)*({U}="{code}"))'
            )

    wb.save(output_path)
    print(f"결과레이아웃 26.4월 블록({GROUP_ROWS[0]}~{GROUP_ROWS[-1]}행) 완료 -> {output_path}")


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print("Usage: python fill_layout_formulas.py <입력.xlsx> <출력.xlsx>")
        sys.exit(1)
    fill(sys.argv[1], sys.argv[2])
