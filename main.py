import os
import openpyxl
import win32com.client
#파일 중 다른 년도, 다른 월 들어오면 애러표시 아직 구현 안함


def copy_sheet(wb_ct, directory_cf, sheet_num, sheet_name): #wb_ct : workbook copy to / directory_cf : directory copy from / sheet_num : 복사해올 시트 번호
    wb = excel.Workbooks.Open(directory_cf)
    wb.Worksheets(wb.Sheets[sheet_num].Name).Copy(Before=wb_ct.Worksheets("Sheet1"))
    ws_t = wb_ct.Worksheets(wb.Sheets[sheet_num].Name)
    ws_t.Name = sheet_name


wd = os.getcwd() #working directory
dd = wd + "\\" + "data_files" #data files directory
filelist = os.listdir(dd)


#파일 병합 시작
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
        

year = int(yanolza[:4])
month = int(yanolza[5:7])
rd = wd + "\\" + "result" + "\\" + yanolza[:4] + "_" + yanolza[5:7]

if not os.path.exists(rd):
    os.makedirs(rd)

#파일 옮기기
print("""
1. 파일 옮겨짐 (실제 사용 목적)
2. 파일 안옮겨짐(테스트 목적)
""")
a = int(input())

sd = rd + "\\" + "data_files" #source 파일 저장할 디렉토리 생성
if not os.path.exists(sd):
    os.makedirs(sd)
if a == 1:
    for i in filelist:
        os.replace(dd + "\\" + i, sd +"\\" + i)
elif a == 2:
    sd = dd

excel = win32com.client.Dispatch("Excel.Application")
wb_new = excel.Workbooks.Add()

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


ws1 = wb_new.Worksheets.Add()
ws1.Name = "calculate"

filename = rd + "\\" + "main.xlsx"
wb_new.SaveAs(filename)
excel.Quit()
#파일 병합 끝

#데이터 한 시트에 묶기
wb = openpyxl.load_workbook(filename)
wb.move_sheet("calculate", -9)

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

wb.save(filename)
#수식 적는 곳 끝
