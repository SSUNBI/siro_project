import openpyxl

def add_calculate(wb, year, month):
    #야놀자
    ws = wb["yanolza"]

    #여기어때 시트 정산금액 텍스트 형식 숫자형식으로 W열에 저장
    sheet_yogi_A = wb["yogi_A"]
    for i in range(4,105):
        sheet_yogi_A.cell(row = i, column = 23, value = '=I'+str(i)+'*1')

    sheet_yogi_B = wb["yogi_B"]
    for i in range(4,105):
        sheet_yogi_B.cell(row = i, column = 23, value = '=I'+str(i)+'*1')

    sheet_yogi_C = wb["yogi_C"]
    for i in range(4,105):
        sheet_yogi_C.cell(row = i, column = 23, value = '=I'+str(i)+'*1')

    #수식 적는 곳 시작
    sheet = wb["calculate"]

    sheet['A1'] = '매출종합본'
    sheet['B1'] = year
    sheet['C1'] = month

    sheet['A3'] = '회사명'
    sheet['A4'] = '떠나요'
    sheet['A5'] = '여기어때A동,M동'
    sheet['A6'] = '여기어때B동'
    sheet['A7'] = '여기어때C동'
    sheet['A8'] = '야놀자'
    sheet['A9'] = '총합'

    sheet['B3'] = '월매출총합'
    sheet['B4'] = '=SUM(ddnayo!J:J)'
    sheet['B5'] = '=SUM(yogi_A!W4:W105)'
    sheet['B6'] = '=SUM(yogi_B!W4:W105)'
    sheet['B7'] = '=SUM(yogi_C!W4:W105)'
    sheet['B8'] = '=SUM(yanolza!L:L)'
    sheet['B9'] = '=SUM(B4:B8)'