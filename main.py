import os
import openpyxl
import win32com.client
import sheet_merge
import excel_calculate
#파일 중 다른 년도, 다른 월 들어오면 애러표시 아직 구현 안함

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

sheet_merge.merge_sheet(filelist, sd, wb_new)

ws1 = wb_new.Worksheets.Add()
ws1.Name = "calculate"

filename = rd + "\\" + "main.xlsx"
wb_new.SaveAs(filename)
excel.Quit()
#파일 병합 끝

#데이터 한 시트에 묶기
wb = openpyxl.load_workbook(filename)
wb.move_sheet("calculate", -9)
sheet_merge.merge_data(wb)
excel_calculate.add_calculate(wb, year, month)

wb.save(filename)
#수식 적는 곳 끝