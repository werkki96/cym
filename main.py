# 0.5반올림 문제점
def roundUp(num):
    if (num - int(num)) >= 0.5:
        return int(num) + 1
    else:
        return float(num)


# basic, advanced1, advanced2 구분하여 전달
# basic & advanced1 calculator function
def fBasicAd1Calc(data, dLvlType, mode, formula="n"):
    data["eL"] = (data["effect"] * data["lvl"]) / data["price"]
    data["edLvl1"] = (data["effect"] * data["dLvl1"]) / data["price"]
    data["edLvl2"] = (data["effect"] * data["dLvl2"]) / data["price"]
    data["eTime"] = (data["effect"] * data["time"]) / data["price"]
    data["eaLoc"] = (data["effect"] * data["aLoc"]) / data["price"]

    if mode == "basic":
        if dLvlType == True:
            data["score"] = data["eL"] + data["edLvl2"] + data["eTime"] + data["eaLoc"]
        else:
            data["score"] = data["eL"] + data["edLvl1"] + data["eTime"] + data["eaLoc"]
    else:
        # advanced 1
        e = data["효과도"]
        lvl = data["난이도"]
        dLvl1 = data["방어 단계1"]
        dLvl2 = data["방어 단계2"]
        aLoc = data["공격 경로 상 위치"]
        time = data["적용 가능 시간"]
        if dLvlType == True:
            data["score"] = calculate_score(e, lvl, dLvl2, aLoc, time, formula)
        else:
            data["score"] = calculate_score(e, lvl, dLvl2, aLoc, time, formula)

    return data

def calculate_score(e, lvl, dLvl, aLoc, time, formula):
    eL = eval(formula['eL'])
    #if dLvlType == True:
        #formula["eDLvl"] = formula["eDLvl"].replace("dLvl", "dLvl2")
    eDLvl = eval(formula['eDLvl'])
    #else:
        #formula["eDLvl"] = formula["eDLvl"].replace("dLvl", "dLvl1")
        #eDLvl = eval(formula['eDLvl2'])
    eTime = eval(formula['eTime'])
    eALoc = eval(formula['eALoc'])
    score = eval(formula['score'])
    return score


# 핸들러
def modeHandler(mode, formula="n", filepath='CyMIA_DOO.xlsx'):
    # 랜덤난수로 효과도, 난이도, 공격 경로 상 위치 임시 설정
    import random
    # 효과도(공격자의 의도) 랜덤 난수 생성 (1-5)
    # 저장된 DB를 기반으로 공격자의 의도에 따라 조견표에 기반하여 점수 산출 하도록 수정해야함
    e = random.randrange(1, 5)
    # 난이도(lvl)는 자산에 따라 다르므로 우선적으로 0-10 사이로 난수 생성
    # 방책 평가 지표의 난이도 식은 (자산의 cvss 총합) / (자산의 cve 개수) 이므로 저장된 DB에 기반하여 수정 필요
    lvl = random.randrange(0, 10)
    # 공격 경로상 위치(aLoc)도 마찬가지로 추후 수정
    # 1 - (노드위치 / 어택그래프상 전체 길이)
    aLoc = random.randrange(0, 10) / 10
    #print(e, lvl, aLoc)
    if mode == 'basic':
        data = load_data(e, lvl, aLoc, filepath, fBasicAd1Calc, formula, mode)
        data = sorted(data, key=lambda e: (e['score']))
    elif mode == "advanced1":
        data = load_data(e, lvl, aLoc, filepath, fBasicAd1Calc, formula, mode)
        data = sorted(data, key=lambda e: (e['score']))
    elif mode == 'advanced2':
        data = advanced2(e, lvl, aLoc, filepath, formula)


# 정규화
def standardize(var, rang):
    if rang == '0-1':
        # 0과 1사이값을 *10 함으로써 0-10범위로 변경
        var = round(var, 1)
        var *= 10

    # 0-10, 1-10범위는 모두 리턴식 하나만으로 1-5로 가능
    # 단, 0의경우 추후 계산식에 문제가 생길 수 있으므로 0은 1로 치환(효과가 아에 없는 경우는 이미 사용하는 방어 절차인 경우 외에는 없음)
    if var == 0:
        var = 1
    var = roundUp(var / 2)

    return round(var)


def dLvlStandardize(dLvl):
    score = 0
    if "초기대응" in dLvl:
        score = 5
    elif "탐지" in dLvl:
        score = 4
    elif "복구대응" in dLvl:
        score = 3
    elif "조사분석" in dLvl:
        score = 1

    return score


# 데이터 로드
def load_data(e, lvl, aLoc, excelPath, method, formula, mode):
    import openpyxl

    datas = []
    # 엑셀 파일 열기
    # /content/CyMIA_DOO.xlsx
    workbook = openpyxl.load_workbook(excelPath, data_only=True)

    # 시트 선택
    sheet = workbook['CyMIA_CounterMeasure']
    max_price = sheet['T1'].value
    min_price = sheet['U1'].value

    # 하나의 자산에대한 고정값 정규화
    # e = standardize(e, '1-5')
    lvl = standardize(lvl, '0-10')
    aLoc = standardize(aLoc, '0-1')
    # 각 컬럼 출력
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[0] == None:
            break
            # 비용 최소최대 정규화 후 방책에 적용하기위한 1-5 정규화
        try:
            price = (row[11] - min_price) / (max_price - min_price)
        except ZeroDivisionError:
            price = 0
        price = standardize(price, "0-1")

        # 방어단계 1-5 정규화. 기존 0.5(조사분석), 0.7(복구대응), 0.8(탐지), 1(초기대응)의 형태로 되어 있었으나
        # 방책 정규화 적용 시 탐지와 초기대응이 동일한 점수를 가지게 되므로 초기대응 = 5 탐지 = 4 복구대응 = 3 조사분석 = 1로 변경
        dLvl1 = dLvlStandardize(row[13])
        dLvl2 = dLvlStandardize(row[14])
        # 적용가능 시간 정규화
        time = standardize(row[16], "0-1")

        # CounterMeasure No, 효과도, 비용, 난이도, 방어단계 1, 방어단계 2, 공격 경로 상 위치, 적용 가능시간
        dataJson = {"No": row[0], "효과도": e, "비용": price, "난이도": lvl, "방어 단계1": dLvl1, "방어 단계2": dLvl2,
                    "공격 경로 상 위치": aLoc, "적용 가능시간": time}
        # 방어단계2 가있는경우
        if dLvl2 != 0:
            datas.append(method(dataJson, True, mode, formula))

        datas.append(method(dataJson, False, mode, formula))
    print(datas)
    return datas


# advanced 2(PCA) 효과도, 난이도, 공격 경로상 위치, 방책DB 파일 위치, 선택된 효과지표
def advanced2(e, lvl, aLoc, excelPath, sel_cols):
    import pandas as pd
    import numpy as np
    from sklearn.decomposition import PCA
    from sklearn.preprocessing import StandardScaler

    # csv 파일 읽기
    df = pd.read_excel(excelPath, sheet_name='CyMIA_CounterMeasure')

    # 사용자가 선택한 컬럼 추출
    selected_cols = sel_cols
    X = df[selected_cols].values
    # print(X)
    target_index = []
    if "효과도" in selected_cols:
        # target_index.append(selected_cols.index('효과도'))
        X[:, selected_cols.index('효과도')] = e
    if "난이도" in selected_cols:
        target_index.append(selected_cols.index('난이도'))
        X[:, selected_cols.index('난이도')] = lvl
    if "공격 경로 상 위치" in selected_cols:
        X[:, selected_cols.index('공격 경로 상 위치')] = aLoc
    if "방어 단계1" in selected_cols:
        X = [[x[i] if i != selected_cols.index('방어 단계1') or type(x[i]) == type(1) else 5 if '초기대응' in x[i] else x[i] for
              i in range(len(x))] for x in X]
        X = [[x[i] if i != selected_cols.index('방어 단계2') or type(x[i]) == type(1) else 5 if '초기대응' in x[i] else x[i] for
              i in range(len(x))] for x in X]
        X = [[x[i] if i != selected_cols.index('방어 단계1') or type(x[i]) == type(1) else 3 if '복구대응' in x[i] else x[i] for
              i in range(len(x))] for x in X]
        X = [[x[i] if i != selected_cols.index('방어 단계2') or type(x[i]) == type(1) else 3 if '복구대응' in x[i] else x[i] for
              i in range(len(x))] for x in X]
        X = [[x[i] if i != selected_cols.index('방어 단계1') or type(x[i]) == type(1) else 1 if '조사분석' in x[i] else x[i] for
              i in range(len(x))] for x in X]
        X = [[x[i] if i != selected_cols.index('방어 단계2') or type(x[i]) == type(1) else 1 if '조사분석' in x[i] else x[i] for
              i in range(len(x))] for x in X]
        X = [[x[i] if i != selected_cols.index('방어 단계1') or type(x[i]) == type(1) else 4 if '탐지' in x[i] else x[i] for i
              in range(len(x))] for x in X]
        X = [[x[i] if i != selected_cols.index('방어 단계2') or type(x[i]) == type(1) else 4 if '탐지' in x[i] else x[i] for i
              in range(len(x))] for x in X]
        X = [[x[i] if i != selected_cols.index('방어 단계2') or type(x[i]) == type(1) else 0 if '-' in x[i] else x[i] for i
              in range(len(x))] for x in X]


    # 방어단계 2가 있는 경우 row 하나 추가(score 때문)
    separate_X = []

    #print(X)
    if "방어 단계1" in selected_cols:
        for d in X:

            if d[selected_cols.index('방어 단계2')] >= 1:
                #print(d[selected_cols.index('방어 단계2')] >= 1)
                #print(d[selected_cols.index('방어 단계2')])
                tmp = d[:]
                tmp[selected_cols.index('방어 단계1')] = tmp[selected_cols.index('방어 단계2')]
                del tmp[selected_cols.index('방어 단계2')]
                #tmp["방어 단계"] = tmp.pop("방어 단계1")
                separate_X.append(tmp)
            tmp = d[:]
            del tmp[selected_cols.index('방어 단계2')]
            #tmp["방어 단계"] = tmp.pop("방어 단계1")
            separate_X.append(tmp)
            #print(separate_X)
        del selected_cols[selected_cols.index('방어 단계2')]
        selected_cols[selected_cols.index('방어 단계1')] = "방어 단계"
    else:
        separate_X = X
    #print(separate_X)

    # Z-score normalization
    scaler = StandardScaler()
    X_norm = scaler.fit_transform(X)

    # PCA 수행
    pca = PCA()
    pca.fit(X_norm)

    # 가장 신뢰성 있는 차원 찾기
    max_ratio = 0
    max_idx = 0
    for i in range(len(pca.explained_variance_ratio_)):
        if pca.explained_variance_ratio_[i] > max_ratio:
            max_ratio = pca.explained_variance_ratio_[i]
            max_idx = i

    # 가장 신뢰성 있는 차원으로 데이터 압축
    X_compressed = pca.transform(X_norm)[:, max_idx]
    # 점수 매기기
    scores = []
    datas = []
    for i in range(len(X_compressed)):
        score = X_compressed[i]
        scores.append(score)
        d = {"No": df.loc[i, 'No.']}

        for idx, t in enumerate(selected_cols):
            d[t] = X[i][idx]
        d["score"] = round(score, 2)
        datas.append(d)

    # 우선순위 매기기
    sorted_idx = sorted(range(len(scores)), key=lambda k: scores[k], reverse=True)
    #print(sorted(datas, key=lambda e: (e['score']), reverse=True))
    for i in sorted_idx:
        print(f"Row {i + 1}: Score = {datas[i]}")
    datas = sorted(datas, key=lambda e: (e['score']), reverse=True)

    return datas


if __name__ == '__main__':
    # datas = modeHandler('basic')
    datas = modeHandler('advanced1', {"eL":  "e * lvl * price", "eDLvl": "e * dLvl * price", "eTime": "e * time * price", "eALoc": "e * aLoc * price", "score": "(eL + eDLvl + eTime + eALoc) / 4"})
    print(datas)
    #datas = modeHandler('advanced2', ["비용", "난이도", "방어 단계1", "방어 단계2", "효과도"])