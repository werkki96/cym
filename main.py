# 샘플 Python 스크립트입니다.

# Shift+F10을(를) 눌러 실행하거나 내 코드로 바꿉니다.
# 클래스, 파일, 도구 창, 액션 및 설정을 어디서나 검색하려면 Shift 두 번을(를) 누릅니다.

# 스크립트를 실행하려면 여백의 녹색 버튼을 누릅니다.
import pandas as pd
def modeHandler():


def load_data(e, lvl, aLoc):
    # 랜덤난수로 효과도, 난이도, 공격 경로 상 위치 임시 설정
    import random
    import openpyxl
    max_price = int(0)
    min_price = int(999999999)
    # 효과도(공격자의 의도) 랜덤 난수 생성 (1-5)
    # 저장된 DB를 기반으로 공격자의 의도에 따라 조견표에 기반하여 점수 산출 하도록 수정해야함
    e = random.random(1, 5)
    # 난이도(lvl)는 자산에 따라 다르므로 우선적으로 0-10 사이로 난수 생성
    # 방책 평가 지표의 난이도 식은 (자산의 cvss 총합) / (자산의 cve 개수) 이므로 저장된 DB에 기반하여 수정 필요
    lvl = random.random(0, 10)


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
        # CounterMeasure No, 효과도, 비용, 난이도, 방어단계 1, 방어단계 2, 공격 경로 상 위치, 적용 가능시간

        datas.append({"No": row[0], "effect": e, "price": row[11], "lvl": lvl, "dLvl1": row[13], "dLvl2": row[14], "time": row[16]})
    return datas, max_price, min_price

if __name__ == '__main__':
    load_data(1, 1, 1)