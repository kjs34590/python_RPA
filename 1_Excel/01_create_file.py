from openpyxl import Workbook
wb = Workbook() # 새 워크북 생성
ws = wb.active # 현재 활성화된 sheet 가져오기
ws.title = "My Sheet" # sheet 이름 변경
wb.save("./PYTHON_WORKS/python_RPA/1_Excel/DATAsample.xlsx")
wb.close # 파일 닫기

