# 샘플 Python 스크립트입니다.

# Shift+F10을(를) 눌러 실행하거나 내 코드로 바꿉니다.
# 클래스, 파일, 도구 창, 액션 및 설정을 어디서나 검색하려면 Shift 두 번을(를) 누릅니다.

# 스크립트를 실행하려면 여백의 녹색 버튼을 누릅니다.
import pandas as pd

def load_data(e, lvl, aLoc):
    import openpyxl
    max_price = int(0)
    min_price = int(99999999999)
    datas = []
    # 엑셀 파일 열기
    workbook = openpyxl.load_workbook('CyMIA_DOO.xlsx', data_only=True)

    # 시트 선택
    sheet = workbook['CyMIA_CounterMeasure']

    # 각 컬럼 출력
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if int(row[11]) > max_price:
            max_price = int(row[11])
        if int(row[11]) < min_price:
            min_price = int(row[11])

        print(row[0], row[11], row[13], row[14], row[16])
        print(min_price, max_price)
        datas.append({"No": row[0], "price": row[11], "dLvl1": row[13], "dLvl2": row[14], "time": row[16]})

if __name__ == '__main__':
    load_data(1, 1, 1)