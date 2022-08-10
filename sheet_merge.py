import os
import openpyxl
import win32com.client

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