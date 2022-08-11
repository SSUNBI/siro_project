from unittest import skip
import openpyxl
import win32com.client
from openpyxl.worksheet.table import Table, TableStyleInfo

excel = win32com.client.Dispatch("Excel.Application")
def copy_sheet(wb_ct, directory_cf, sheet_num, sheet_name): #wb_ct : workbook copy to / directory_cf : directory copy from / sheet_num : 복사해올 시트 번호
    wb = excel.Workbooks.Open(directory_cf)
    wb.Worksheets(wb.Sheets[sheet_num].Name).Copy(Before=wb_ct.Worksheets("Sheet1"))
    ws_t = wb_ct.Worksheets(wb.Sheets[sheet_num].Name)
    ws_t.Name = sheet_name

def merge_sheet(filelist, sd, wb_new):
    for i in filelist:
        if i[0] == '2':
            if 'A' in i:
                A_bank = i
            elif 'B' in i:
                B_bank = i
            elif 'C' in i:
                C_bank = i
            elif 'e' in i:
                cafe_bank = i
            elif 'M' in i:
                M_bank = i
            else:
                yanolza = i
        elif i[0] == 'A' and i[1] == 'd':
            ddnayo = i
        else:
            if 'A' in i:
                yogi_A = i
            elif 'B' in i:
                yogi_B = i
            elif 'C' in i:
                yogi_C = i
    
    copy_sheet(wb_new, sd + '\\' + yanolza, 0, "yanolza")
    copy_sheet(wb_new, sd + '\\' + ddnayo, 0, "ddnayo")
    copy_sheet(wb_new, sd + '\\' + yogi_A, 0, "yogi_A")
    copy_sheet(wb_new, sd + '\\' + yogi_B, 0, "yogi_B")
    copy_sheet(wb_new, sd + '\\' + yogi_C, 0, "yogi_C")
    copy_sheet(wb_new, sd + '\\' + cafe_bank, 0, "cafe_bank")
    copy_sheet(wb_new, sd + '\\' + M_bank, 0, "M_bank")
    copy_sheet(wb_new, sd + '\\' + A_bank, 0, "A_bank")
    copy_sheet(wb_new, sd + '\\' + B_bank, 0, "B_bank")
    copy_sheet(wb_new, sd + '\\' + C_bank, 0, "C_bank")

def merge_data(wb):
    ws = wb["Sheet1"]

    #야놀자
    yanolza_s = wb["yanolza"]
    cnt = 0
    temp = list()
    for row in yanolza_s.rows:
        if cnt < 2:
            cnt += 1
            continue
        room = row[1].value
        year = row[2].value[:2]
        month = row[2].value[3:5]
        day = row[2].value[6:8]
        date = year + month + day
        money = int(row[11].value)
        funnel = "yanolza" #유입경로
        temp.append([date, room, money, funnel])
    
    ddnayo_s = wb["ddnayo"]
    cnt = 0
    for row in ddnayo_s.rows:
        if cnt < 1:
            cnt += 1
            continue
        room = row[1].value[:4]
        year = row[6].value[2:4]
        month = row[6].value[5:7]
        day = row[6].value[8:10]
        date = year + month + day
        money = int(row[9].value)
        funnel = "ddnayo" #유입경로
        temp.append([date, room, money, funnel])
    
    yogi_A_s = wb["yogi_A"]
    cnt = 0
    for row in yogi_A_s.rows:
        if cnt < 3:
            cnt += 1
            continue
        room = row[5].value[-4:]
        year = row[0].value[2:4]
        month = row[0].value[5:7]
        day = row[0].value[8:10]
        date = year + month + day
        money = int(row[8].value.replace(",", ""))
        funnel = "yogi_A" #유입경로
        temp.append([date, room, money, funnel])

    yogi_B_s = wb["yogi_B"]
    cnt = 0
    for row in yogi_B_s.rows:
        if cnt < 3:
            cnt += 1
            continue
        room = row[5].value[-4:]
        year = row[0].value[2:4]
        month = row[0].value[5:7]
        day = row[0].value[8:10]
        date = year + month + day
        money = int(row[8].value.replace(",", ""))
        funnel = "yogi_B" #유입경로
        temp.append([date, room, money, funnel])

    yogi_C_s = wb["yogi_C"]
    cnt = 0
    for row in yogi_C_s.rows:
        if cnt < 3:
            cnt += 1
            continue
        room = "C" + row[5].value[-3:]
        year = row[0].value[2:4]
        month = row[0].value[5:7]
        day = row[0].value[8:10]
        date = year + month + day
        money = int(row[8].value.replace(",", ""))
        funnel = "yogi_C" #유입경로
        temp.append([date, room, money, funnel])
    
    ws.append(["이용날짜", "이용객실", "결제금액", "유입경로"])
    for row in temp:
        ws.append(row)
    reference = "A1:D" + str(len(temp) + 1)
    tab = Table(displayName="Table1", ref=reference)
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                       showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tab.tableStyleInfo = style
    ws.add_table(tab)
