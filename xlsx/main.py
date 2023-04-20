# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

from openpyxl import Workbook #openpyxl 패키지 불러오기

wb = Workbook() #새 엑셀파일 생성
ws = wb.active #가장 처음 만들어진 sheet 사용
ws_new = wb.create_sheet() #신규 시트 생성

ws.title = "First Sheet" #시트 이름 바꾸기

ws_new.title = "New Sheet" #새 시트 이름 바꾸

wb.save("test.xlsx") #엑셀파일 저장
wb.close()