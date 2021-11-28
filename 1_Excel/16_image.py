from openpyxl import Workbook
from openpyxl.drawing.image import Image

wb = Workbook()
ws = wb.active

img = Image("./PYTHON_WORKS/python_RPA/1_Excel/excel_image.png")

# C3 위치에 이미지 삽입
ws.add_image(img, "C3")

wb.save("./PYTHON_WORKS/python_RPA/1_Excel/sample_image.xlsx")

# ImpoertError : You must install Pillow to fetch image....
# 위 오류 발생 시 : pip install Pillow
