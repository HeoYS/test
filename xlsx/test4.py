
from openpyxl import load_workbook
#엑셀에서 행, 열삽입, 삭제, 이동

wb = load_workbook('test3.xlsx')
ws = wb.active

#행(Row)추가 ws.insert_rows(idx, amount), **열(column)도 같다
ws.insert_rows(8)
ws.insert_rows(8, 5) #8번째 row 위치에서 아래로 5row 추가

wb.save('test3.xlsx')

