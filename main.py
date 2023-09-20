
import pandas as pd

# 0.5반올림 문제점
def roundUp(num):
    chk_num = 0
    if num < 0:
        num = num * -1
        chk_num = 1
    if (num - int(num)) >= 0.5:
        num = int(num) + 1
        if chk_num == 1:
            num = num * -1
        return num
    else:
        num = float(num)
        if chk_num == 1:
            num = num * -1
        return float(num)

# 검색
def search_data(cur, sql):
    cur.execute(sql)
    rows = cur.fetchall()
    return rows

# 1 row 검색
def search_data_one(cur, sql):
    cur.execute(sql)
    rows = cur.fetchone()
    return rows

# db 연결
def connect_db():
    import pymysql
    con = pymysql.connect(host='localhost', user='root', password='1234', db='cymia_test', charset='utf8')
    cur = con.cursor(pymysql.cursors.DictCursor)
    return cur

# basic, advanced1, advanced2 구분하여 전달
# basic & advanced1 calculator function
def fBasicAd1Calc(data, dLvlType, mode, formula="n"):
    # data["eL"] = (data["효과도"] * data["난이도"]) / data["비용"]
    # data["edLvl1"] = (data["효과도"] * data["방어 단계1"]) / data["비용"]
    # data["edLvl2"] = (data["효과도"] * data["방어 단계2"]) / data["비용"]
    # data["eTime"] = (data["효과도"] * data["적용 가능시간"]) / data["비용"]
    # data["eaLoc"] = (data["효과도"] * data["공격 경로 상 위치"]) / data["비용"]

    if mode == "basic":
        if dLvlType == True:
            #data["score"] = data["eL"] + data["edLvl2"] + data["eTime"] + data["eaLoc"]
            data["score"] = (data["효과도"] + data["방어 단계2"] + data["공격 경로 상 위치"] + data["적용 가능시간"] + data["난이도"]) / data["비용"]
        else:
            #data["score"] = data["eL"] + data["edLvl1"] + data["eTime"] + data["eaLoc"]
            data["score"] = (data["효과도"] + data["방어 단계1"] + data["공격 경로 상 위치"] + data["적용 가능시간"] + data["난이도"]) / data["비용"]
    else:
        # advanced 1
        e = data["효과도"]
        lvl = data["난이도"]
        dLvl1 = data["방어 단계1"]
        dLvl2 = data["방어 단계2"]
        aLoc = data["공격 경로 상 위치"]
        #time = data["적용 가능시간"]
        #price = data["비용"]
        if dLvlType == True:
            data["score"] = calculate_score(data["효과도"], data["비용"], data["난이도"], data["방어 단계2"], data["공격 경로 상 위치"], data["적용 가능시간"], formula)
        else:
            data["score"] = calculate_score(data["효과도"], data["비용"], data["난이도"], data["방어 단계1"], data["공격 경로 상 위치"], data["적용 가능시간"], formula)

    return data

def calculate_score(e, lvl, price, dLvl, aLoc, time, formula):
    eL = eval(formula['eL'])
    eDLvl = eval(formula['eDLvl'])
    eTime = eval(formula['eTime'])
    eALoc = eval(formula['eALoc'])
    score = eval(formula['score'])
    return score


# 핸들러
def modeHandler(mode, asset_infos, formula="n", filepath='/content/drive/MyDrive/Colab Notebooks/CyMIA_DOO.xlsx'):
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
    #ratio = []
    if mode == 'basic':
        data = load_data(asset_infos, lvl, aLoc, filepath, fBasicAd1Calc, formula, mode)
    elif mode == "advanced1":
        data = load_data(asset_infos, lvl, aLoc, filepath, fBasicAd1Calc, formula, mode)
    elif mode == 'advanced2':
        data, ratios = advanced2(e, lvl, aLoc, filepath, formula, asset_infos)

    data = sorted(data, key=lambda e: (e['score']), reverse=True)
    return data, ratios

# 정규화
def normalization(var, rang):
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


def dLvlNormalization(dLvl):
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
def load_data(asset_infos, lvl, aLoc, excelPath, method, formula, mode):
    import openpyxl
    cur = connect_db()

    datas = []
    # 엑셀 파일 열기
    # /content/CyMIA_DOO.xlsx
    #/content/drive/MyDrive/Colab Notebooks/CyMIA_DOO.xlsx
    workbook = openpyxl.load_workbook(excelPath, data_only=True)
    sql = f"""
        SELECT * FROM cymia_cm
    """
    defence_data = search_data(cur, sql)
    # 시트 선택
    sheet = workbook['CyMIA_CounterMeasure']
    max_price = 10000#sheet['T1'].value
    min_price = 13000000#sheet['U1'].value

    # 하나의 자산에대한 고정값 정규화
    sql = f"""
        SELECT ROUND(SUM(base_score)/COUNT(cve_id), 2) AS lvl FROM cve_test
        WHERE cve_id IN (SELECT cve_id FROM asset_cve_mapping WHERE asset_id='{asset_infos[0]['selected_asset']}')
    """
    lvl = search_data_one(cur, sql)
    lvl = lvl['lvl']
    lvl = normalization(lvl, '0-10')
    # 각 컬럼 출력
    for rows in defence_data:
        row = list(rows.values())
#    for row in sheet.iter_rows(min_row=2, values_only=True):
        ### 09월 수정 1차 ( 공격 경로 상 존재하는 테크닉과 연관되지 않은 방어 테크닉은 패스 )
        if not chk_attack_relationship(asset_infos, row[1]):
            continue
        ### 9월 수정 1차 End ###
        if row[0] == None:
            break
            # 비용 최소최대 정규화 후 방책에 적용하기위한 1-5 정규화
        try:
            price = (row[11] - min_price) / (max_price - min_price)
        except ZeroDivisionError:
            price = 0
        price = normalization(price, "0-1")

        # 효과도 시작
        ###########
        base_cvss = 0.0
        def_cvss = 0.0
        # 자산의 cvss 총합
        sub_sql_cve = "("
        cnt = 0
        leng = len(asset_infos[asset_infos[0]["selected_asset"]]["cve_list"])
        for cves in asset_infos[asset_infos[0]["selected_asset"]]["cve_list"]:
            base_cvss += float(cves["base_score"])
            if leng > 1:
                if leng == cnt+1:
                    sub_sql_cve += f"cve_id = '{cves['cve_id']}')"
                else:
                    sub_sql_cve += f"cve_id = '{cves['cve_id']}' or "
            elif leng == 1:
                sub_sql_cve += f"cve_id = '{cves['cve_id']}')"
            cnt += 1
        # 방책이 적용되는 cve의 cvss합구하기
        tmp = []
        for tech in asset_infos[asset_infos[0]["selected_asset"]]["tech"]:
            sql = f"SELECT distinct `technique_id` from attack_defend_mapping where defense_id='{row[1]}' and `technique_id`='{tech.strip()}'"
            te = search_data(cur, sql)
            if len(te) > 0:
                sql = f"""select cve_id, base_score from cve_mapping cm, cve_test ct where (`Primary Impact` like '%{tech.strip()}%' or `Exploitation Technique` like '%{tech.strip()}%') 
                                                               and (`Primary Impact` not like '%{tech.strip()}.%' and `Exploitation Technique` not like '%{tech.strip()}.%') and cm.`CVE ID` =ct.cve_id
                                                               and {sub_sql_cve}"""
                tmps = search_data(cur, sql)
                tmp += tmps
                tmp = list(map(dict, set(tuple(sorted(d.items())) for d in tmp)))
                def_cvss = sum(float(item["base_score"]) for item in tmp)

        e = round(1-((base_cvss-def_cvss)/base_cvss), 2)
        e = normalization(e, '0-1')
        # 효과도 종료 #####

        # 공격 경로 상 위치 시작
        cnt = 0
        for asset in asset_infos:
            if cnt != 0:
                tech_list = ', '.join([f"'{tech}'" for tech in asset['tech']])
                sql = f"select distinct 'ok' as chk from attack_defend_mapping adfm where `Technique_ID` in ({tech_list}) and defense_id='{row[1]}'"
                chk = search_data(cur, sql)
                if len(chk) > 0:
                    break
            cnt += 1
        aLoc = 1-round(cnt/len(asset_infos), 2)
        aLoc = normalization(aLoc, '0-1')
        # 공격 경로 상 위치 종료


        # 방어단계 1-5 정규화. 기존 0.5(조사분석), 0.7(복구대응), 0.8(탐지), 1(초기대응)의 형태로 되어 있었으나
        # 방책 정규화 적용 시 탐지와 초기대응이 동일한 점수를 가지게 되므로 초기대응 = 5 탐지 = 4 복구대응 = 3 조사분석 = 1로 변경
        dLvl1 = dLvlNormalization(row[13])
        dLvl2 = dLvlNormalization(row[14])
        # 적용가능 시간 정규화
        time = normalization(row[16], "0-1")

        # CounterMeasure No, 효과도, 비용, 난이도, 방어단계 1, 방어단계 2, 공격 경로 상 위치, 적용 가능시간
        dataJson = {"No": row[0], "효과도": e, "비용": price, "난이도": lvl, "방어 단계1": dLvl1, "방어 단계2": dLvl2,
                    "공격 경로 상 위치": aLoc, "적용 가능시간": time}
        # 방어단계2 가있는경우
        if dLvl2 != 0:
            datas.append(method(dataJson, True, mode, formula))

        datas.append(method(dataJson, False, mode, formula))
    #print(datas)
    return datas


# advanced 2(PCA) 효과도, 난이도, 공격 경로상 위치, 방책DB 파일 위치, 선택된 효과지표
def advanced2(e, lvl, aLoc, excelPath, sel_cols, asset_infos):
    import pandas as pd
    import numpy as np
    from sklearn.decomposition import PCA
    from sklearn.preprocessing import StandardScaler
    import matplotlib.pyplot as plt
    import copy
    np.set_printoptions(threshold=np.inf, linewidth=np.inf)
    # csv 파일 읽기
    cur = connect_db()

    df = pd.read_excel(excelPath, sheet_name='CyMIA_CounterMeasure')


    # 사용자가 선택한 컬럼 추출
    selected_cols = sel_cols
    ### 09월 수정 1차 ( 공격 경로 상 존재하는 테크닉과 연관되지 않은 방어 테크닉은 삭제 )
    for idx, data in df.iterrows():
        if not chk_attack_relationship(asset_infos, data["Mitigation ID"]):
            df.drop(idx, axis=0, inplace=True)
    df.reset_index(drop=True, inplace=True)
    ### 9월 수정 2차 End ###
    X = df[selected_cols].values
    target_index = []
    if "효과도" in selected_cols:
        e = []
        for i in range(len(X)):
            base_cvss = 0.0
            def_cvss = 0.0
            # 자산의 cvss 총합
            sub_sql_cve = "("
            cnt = 0
            leng = len(asset_infos[asset_infos[0]["selected_asset"]]["cve_list"])
            for cves in asset_infos[asset_infos[0]["selected_asset"]]["cve_list"]:
                base_cvss += float(cves["base_score"])
                if leng > 1:
                    if leng == cnt + 1:
                        sub_sql_cve += f"cve_id = '{cves['cve_id']}')"
                    else:
                        sub_sql_cve += f"cve_id = '{cves['cve_id']}' or "
                elif leng == 1:
                    sub_sql_cve += f"cve_id = '{cves['cve_id']}')"
                cnt += 1

            # 방책이 적용되는 cve의 cvss합구하기
            tmp = []
            for tech in asset_infos[asset_infos[0]["selected_asset"]]["tech"]:
                sql = f"SELECT distinct `technique_id` from attack_defend_mapping where defense_id='{df['Mitigation ID'][i]}' and `technique_id`='{tech.strip()}'"
                te = search_data(cur, sql)

                if len(te) > 0:
                    sql = f"""select cve_id, base_score from cve_mapping cm, cve_test ct where (`Primary Impact` like '%{tech.strip()}%' or `Secondary Impact` like '%{tech.strip()}%' or `Exploitation Technique` like '%{tech.strip()}%') 
                                                                          and (`Primary Impact` not like '%{tech.strip()}.%' and `Secondary Impact` not like '%{tech.strip()}.%' and `Exploitation Technique` not like '%{tech.strip()}.%') and cm.`CVE ID` =ct.cve_id
                                                                          and {sub_sql_cve}"""
                    tmps = search_data(cur, sql)
                    tmp += tmps
                    tmp = list(map(dict, set(tuple(sorted(d.items())) for d in tmp)))
                    def_cvss = sum(float(item["base_score"]) for item in tmp)
            X[i][selected_cols.index('효과도')] = round(1 - ((base_cvss - def_cvss) / base_cvss), 2)
    if "난이도" in selected_cols:
        target_index.append(selected_cols.index('난이도'))
        sql = f"""
                SELECT ROUND(SUM(base_score)/COUNT(cve_id), 2) AS lvl FROM cve_test
                WHERE cve_id IN (SELECT cve_id FROM asset_cve_mapping WHERE asset_id='{asset_infos[0]['selected_asset']}')
            """
        lvl = search_data_one(cur, sql)
        lvl = lvl['lvl']
        X[:, selected_cols.index('난이도')] = lvl
    if "공격 경로 상 위치" in selected_cols:
        for i in range(len(X)):
            cnt = 0
            for asset in asset_infos:
                if cnt != 0:
                    tech_list = ', '.join([f"'{tech}'" for tech in asset['tech']])
                    sql = f"select distinct 'ok' as chk from attack_defend_mapping adfm where `Technique_ID` in ({tech_list}) and defense_id='{df['Mitigation ID'][i]}'"
                    chk = search_data(cur, sql)
                    if len(chk) > 0:
                        break
                cnt += 1
            aLoc = 1 - round(cnt / len(asset_infos), 2)
            X[i][selected_cols.index('공격 경로 상 위치')] = aLoc
    if "비용" in selected_cols:
        # 비용은 낮을수록 좋으므로 음수처리하여 다른 요소들과 동일하게 높을수록 좋게 함
        X[:, selected_cols.index('비용')] = -1 * X[:, selected_cols.index('비용')]
    if "방어 단계1" in selected_cols:
        X = [[x[i] if i != selected_cols.index('방어 단계1') or type(x[i]) == type(1) else 1 if '초기대응' in x[i] else x[i] for
              i in range(len(x))] for x in X]
        X = [[x[i] if i != selected_cols.index('방어 단계2') or type(x[i]) == type(1) else 1 if '초기대응' in x[i] else x[i] for
              i in range(len(x))] for x in X]
        X = [[x[i] if i != selected_cols.index('방어 단계1') or type(x[i]) == type(1) or type(x[i]) == type(
            0.1) else 0.7 if '복구대응' in x[i] else x[i] for
              i in range(len(x))] for x in X]
        X = [[x[i] if i != selected_cols.index('방어 단계2') or type(x[i]) == type(1) or type(x[i]) == type(
            0.1) else 0.7 if '복구대응' in x[i] else x[i] for
              i in range(len(x))] for x in X]
        X = [[x[i] if i != selected_cols.index('방어 단계1') or type(x[i]) == type(1) or type(x[i]) == type(
            0.1) else 0.5 if '조사분석' in x[i] else x[i] for
              i in range(len(x))] for x in X]
        X = [[x[i] if i != selected_cols.index('방어 단계2') or type(x[i]) == type(1) or type(x[i]) == type(
            0.1) else 0.5 if '조사분석' in x[i] else x[i] for
              i in range(len(x))] for x in X]
        X = [[x[i] if i != selected_cols.index('방어 단계1') or type(x[i]) == type(1) or type(x[i]) == type(
            0.1) else 0.8 if '탐지' in x[i] else x[i] for i
              in range(len(x))] for x in X]
        X = [[x[i] if i != selected_cols.index('방어 단계2') or type(x[i]) == type(1) or type(x[i]) == type(
            0.1) else 0.8 if '탐지' in x[i] else x[i] for i
              in range(len(x))] for x in X]
        X = [[x[i] if i != selected_cols.index('방어 단계2') or type(x[i]) == type(1) or type(x[i]) == type(
            0.1) else 0 if '-' in x[i] else x[i] for i
              in range(len(x))] for x in X]

    # 방어단계 2가 있는 경우 row 하나 추가(score 때문)
    separate_X = []
    flg = 0
    if "방어 단계1" in selected_cols:
        for idx, d in enumerate(X):
            if d[selected_cols.index('방어 단계2')] > 0:
                flg = 1
                tmp = copy.deepcopy(d)
                tmp[selected_cols.index('방어 단계1')] = tmp[selected_cols.index('방어 단계2')]
                del tmp[selected_cols.index('방어 단계2')]
                tmp.append(flg)
                tmp.insert(0, df.loc[idx, "No."])
                separate_X.append(tmp)
            tmp = copy.deepcopy(d)
            del tmp[selected_cols.index('방어 단계2')]
            tmp.append(flg)
            tmp.insert(0, df.loc[idx, "No."])
            flg = 0
            separate_X.append(tmp)
        del selected_cols[selected_cols.index('방어 단계2')]
        selected_cols[selected_cols.index('방어 단계1')] = "방어 단계"
        selected_cols.append("flag")
    else:
        separate_X = X

    # Z-score normalization
    # 데이터를 평균이 0이고 편차가 1인 정규 분포로 변환
    # print(separate_X)
    separate_X = np.array(separate_X)
    components_X = copy.deepcopy(separate_X[:, 1:])
    components_X = components_X.astype(float)
    #print(separate_X)
    scaler = StandardScaler()
    X_norm = scaler.fit_transform(components_X)
    #print(X_norm)
    # X_norm = min_max_normalize(separate_X)
    # PCA 수행
    pca = PCA()
    pca.fit(X_norm)
    # row마다 정규화. 각 row에서 가장 멀리있는 요소를 찾기 위한 사전작업
    Y_norm = components_X / np.max(components_X, axis=1, keepdims=True)
    # 70% 이상의 설명력으로 차원 축소
    ratio70 = 0
    ratio_components = 0
    #print(pca.explained_variance_ratio_)
    for i in range(len(pca.explained_variance_ratio_)):
        ratio70 += pca.explained_variance_ratio_[i]
        ratio_components = i + 1
        if ratio70 > 0.70:
            break
    # 70%이상의 설명력을 가진 요소들로 축소한 차원
    transformed_data = pca.transform(X_norm)[:, :ratio_components]
    transformed_Y = Y_norm[:, :ratio_components]
    datas = []
    distances = 0
    for i in range(len(transformed_data)):
        # 원점으로부터 거리 계산
        for r in range(ratio_components):
            distances = transformed_data[i][r] ** 2 + distances
        distances = np.sqrt(distances)

        # 원점과의 거리에 따른 위치값을 방어행위 평가 지수로 사용
        d = {"No.": separate_X[i][0]}

        for idx in range(len(X_norm[i])):
            if selected_cols[idx] == "flag":
                continue
            d[selected_cols[idx]] = round(components_X[i][idx], 2)
            d[selected_cols[idx] + "norm"] = round(X_norm[i][idx], 2)
        d["score"] = round(distances, 2)
        datas.append(d)

    return datas, pca.explained_variance_ratio_

### 9월 수정 1차 (방책 후보목록 추천) ###
def ranking_defend(matrix):
    import itertools

    # 주어진 2차원 배열 테스트용
    matrixs = [
        [18.0, 18.0, 17.0, 17.0, 17.0],
        [20.0, 20.0, 17.0, 17.0, 17.0],
        [22.0, 22.0, 20.0, 20.0, 20.0],
        [16.0, 16.0, 16.0, 16.0, 16.0],
        [20.0, 19.0, 19.0, 19.0, 17.0],
        [21.0, 21.0, 21.0, 20.0, 19.0],
        [17.0, 17.0, 16.0, 16.0, 16.0]
    ]


    # 가능한 모든 행의 중복순열 생성
    num_rows = len(matrix)
    num_cols = len(matrix[0])
    row_indices = list(range(num_rows))
    row_permutations = list(itertools.product(row_indices, repeat=num_rows))

    # 결과를 저장할 리스트 초기화
    results = []

    # 모든 중복순열에 대한 합을 계산하고 결과 리스트에 저장
    for perm in row_permutations:
        current_sum = sum(matrix[i][min(perm[i], num_cols-1)] for i in range(num_rows))
        results.append((perm, current_sum))

    # 합을 기준으로 내림차순으로 정렬하여 순위를 계산
    ranked_results = sorted(results, key=lambda x: x[1], reverse=True)

    # 결과 출력
    for rank, (perm, total_sum) in enumerate(ranked_results, start=1):
        print(f"{rank}순위 합: {total_sum}, 순열: {perm}")
        if rank > 100:
            break

    print(f"\n1순위 합: {ranked_results[0][1]}, 1순위 순열: {ranked_results[0][0]}")
    return ranked_results

### 9월 수정 1차 (방책 관련 테크닉 확인) ###
def chk_attack_relationship(asset_info, d3fend_technique):
    tech = []

    # 'tech' 값을 1차원 배열에 추가
    for info in asset_info:
        if 'tech' in info:
            tech.extend(info['tech'])
    tech = set(tech)
    tech = ', '.join([f"'{item}'" for item in tech])
    cur = connect_db()
    sql = f"""SELECT technique_ID
                      FROM attack_defend_mapping
                      WHERE defense_ID=\"{d3fend_technique}\" and technique_ID in ({tech})"""
    res = search_data(cur, sql)
    #print(res)
    if len(res) >= 1:
        return 1
    return 0
### 9월 수정 1차 End ###

if __name__ == '__main__':
    import json
    filePath = "CyMIA_DOO.xlsx"
    cur = connect_db()
    # 어택그래프의 3번째 자산이 선택되었다고 가정.
    ag = search_data(cur, "select * from attack_graph_test")
    ag = json.loads(ag[0]['attack_path'])
    #자산 정보 저장 첫 번째 요소는 무조건 선택된 자산
    asset_infos = [{"selected_asset" : 3}]
    for i in range(len(ag['paths']) + 1):
        asset_info = {}

        if i < len(ag['paths']):
            asset_info["asset_id"] = ag['paths'][i]['source']
        else:
            asset_info["asset_id"] = ag['paths'][i - 1]['target']
        asset_info["cve_list"] = search_data(cur
                                             , f"""
                                             SELECT ct.cve_id AS cve_id
                                                    , (CASE
                                                        WHEN ct.base_score  = 0
                                                        -- 만약 cvss가 0인경우(없는경우) 어택그래프에 존재하는 자산들의 전체 cvss를 평균으로 하여 계산
                                                        THEN (SELECT ROUND(SUM(base_score)/COUNT(cve_id), 1) AS score
                                                              FROM cve_test
                                                              WHERE cve_id IN (SELECT cve_id FROM asset_cve_mapping)
                                                              )
                                                        ELSE ct.base_score
                                                      END) AS base_score
                                             FROM asset_cve_mapping acm, cve_test ct
                                             WHERE asset_id = {asset_info['asset_id']}
                                             AND ct.cve_id=acm.cve_id
                                             """)
        asset_info["tech"] = []
        for cve in asset_info['cve_list']:
            a_tech = search_data(cur, f"""
                    select concat_ws(';', `primary impact`, `secondary impact`, `exploitation technique`) as tech
                    from cve_mapping
                    where `cve id` = '{cve['cve_id']}'
                    """)
            if len(a_tech) > 0:
                asset_info["tech"] = asset_info["tech"]+a_tech[0]['tech'].split(';')
        # 어택 테크닉 중복, 공백 제거
        asset_info['tech'] = list(set(filter(None, asset_info['tech'])))
        asset_infos.append(asset_info)

    #filePath = "/content/drive/MyDrive/Colab Notebooks/CyMIA_DOO.xlsx"

    basic_list = list()
    for i in range(len(ag['paths']) + 1):
        print(i + 1, "번 자산")
        # test.append()
        asset_infos[0]["selected_asset"] = i + 1
        # basic, ratio = modeHandler('basic', asset_infos,
        #                            ["비용", "난이도", "방어 단계1", "방어 단계2", "효과도", "공격 경로 상 위치", "적용 가능 시간"], filePath)
        pcaMode, ratio = modeHandler('advanced2', asset_infos, ["난이도", "방어 단계1", "방어 단계2", "효과도", "공격 경로 상 위치", "적용 가능 시간"], filePath)
        # pcaMode, ratio = modeHandler('advanced2', asset_infos, ["난이도", "방어 단계1", "방어 단계2", "효과도", "공격 경로 상 위치", "적용 가능 시간"], filePath)
        # print("----------------basic------------------")
        # print("\n".join([str(d) for d in basic]))
        # tmp = [item['score'] for item in basic]
        # basic_list.append(tmp)

        # print("--------------advanced1------------------")
        # print("\n".join([str(d) for d in advanced1]))
        print("---------------pca------------------")
        print("\n".join([str(d) for d in pcaMode]))
        tmp = [item['score'] for item in pcaMode]
        basic_list.append(tmp)
        print("---------------설명력------------------")
        print("\n".join([str(round(d, 2)) for d in ratio]))
    ranking_defend(basic_list)

