import openpyxl
import calendar

def week_calculate(wb, name, week_end_day):
    cs = wb["Sheet1"]
    week_list = list()
    week_sum = 0
    now_date = cs["A2"].value
    now_year = now_date // 10000 + 2000
    now_month = (now_date % 10000) // 100
    now_day = now_date % 100
    weekday = calendar.weekday(now_year, now_month, now_day)
    month_range = calendar.monthrange(now_year, now_month)[1]
    cnt = 2
    temp_date = now_date
    temp_year = temp_date // 10000 + 2000
    temp_month = (temp_date % 10000) // 100
    temp_day = now_date % 100
    start_date = temp_date
    while True:
        if cnt > cs.max_row:
            if weekday != week_end_day: #week_end_day ddnayo : 6
                week_list.append([start_date, temp_date, week_sum])
            break
        
        now_date = cs["A" + str(cnt)].value
        if temp_date == now_date:
            now_year = now_date // 10000 + 2000
            now_month = (now_date % 10000) // 100
            now_day = now_date % 100
            month_range = calendar.monthrange(now_year, now_month)[1]
            if cs["D" + str(cnt)].value == name:
                week_sum += cs["C" + str(cnt)].value
            cnt += 1
        else:
            while temp_date != now_date:
                end_date = temp_date
                temp_day += 1
                if temp_day > month_range:
                    temp_day = 1
                    temp_month += 1
                    if temp_month > 12:
                        temp_month = 1
                        temp_year += 1
                temp_date = (temp_year - 2000) * 10000 + temp_month * 100 + temp_day
            weekday = calendar.weekday(temp_year, temp_month, temp_day)
            if weekday == (week_end_day + 1) % 7: #ddnayo
                week_list.append([start_date, end_date, week_sum])
                start_date = temp_date
                week_sum = 0
    print(week_list)
    return week_list

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

    #주간 정산금액 불러오기
    #A_bank (야놀자 모든정산금액 + 떠나요 M+A동 + 여기어때 M+A동) - C,E,H열 데이터 뽑아오기
    #B_bank (떠나요B동 + 여기어때B동)
    #C_bank (떠나요C동 + 여기어때C동)
    #야놀자는 '야놀자펜션0117-'라는 이름으로 들어옴
    #떠나요는 '주식회사떠나요'라는 이름으로 들어옴
    #여기어때는 '호텔타임'라는 이름으로 들어옴
    #떠나요 계산기 일~토 화요일정산
    #여기어때 월~일 수요일정산
    #야놀자 월~일 목요일정산
    #mon : 0 tue = 1 wed = 2 ...

    ddn_list = week_calculate(wb, "ddnayo", 5)
    yanolza_list = week_calculate(wb, "yanolza", 6)
    yogi_A_list = week_calculate(wb, "yogi_A", 6)
    yogi_B_list = week_calculate(wb, "yogi_B", 6)
    yogi_C_list = week_calculate(wb, "yogi_C", 6)

    sheet_A_bank = wb["A_bank"]
    yanolza_bank_list = []
    ddn_bank_list = []
    yogi_M_A_bank_list = []
    print(sheet_A_bank.max_column)
    print(sheet_A_bank.max_row)
    for i in range(9,int(sheet_A_bank.max_row)):
        if sheet_A_bank['H'+str(i)] == '주식회사떠나요':
            ddn_bank_list.append([sheet_A_bank['C'+str(i)],sheet_A_bank['E'+str(i)],sheet_A_bank['H'+str(i)]])
            print(ddn_bank_list) 
    else:
        print()


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

    sheet['A10'] = '떠나요'
    sheet['B10'] = '떠나요금액'
    sheet['C10'] = '야놀자'
    sheet['D10'] = '야놀자금액'
    sheet['E10'] = '여기어때A'
    sheet['F10'] = '여기어때A금액'
    sheet['G10'] = '여기어때B'
    sheet['H10'] = '여기어때B금액'
    sheet['I10'] = '여기어때C'
    sheet['J10'] = '여기어때C금액'


    start = 11
    for i in ddn_list:
        sheet['A' + str(start)] = str(i[0]) + ' ~ ' + str(i[1])
        sheet['B' + str(start)] = i[2]
        start += 1
    start = 11
    for i in yanolza_list:
        sheet['C' + str(start)] = str(i[0]) + ' ~ ' + str(i[1])
        sheet['D' + str(start)] = i[2]
        start += 1
    start = 11
    for i in yogi_A_list:
        sheet['E' + str(start)] = str(i[0]) + ' ~ ' + str(i[1])
        sheet['F' + str(start)] = i[2]
        start += 1
    start = 11
    for i in yogi_B_list:
        sheet['G' + str(start)] = str(i[0]) + ' ~ ' + str(i[1])
        sheet['H' + str(start)] = i[2]
        start += 1
    start = 11
    for i in yogi_C_list:
        sheet['I' + str(start)] = str(i[0]) + ' ~ ' + str(i[1])
        sheet['J' + str(start)] = i[2]
        start += 1