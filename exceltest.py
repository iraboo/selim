import openpyxl
from openpyxl import Workbook
from openpyxl.drawing.image import Image

current_folder = 'C:/Users/iraboo/Documents/my_project/selim/'
filename = current_folder + 'AutoTest.xlsx'
wb = openpyxl.load_workbook(filename)
sheet = wb['전기수도요금']
ws_list = wb.sheetnames  #해당 Workbook의 시트 목록을 리스트로 저장
print(sheet)
print(ws_list) #리스트 출력


i = sheet._images[:]
sheet._images.remove(i[0])
sheet._images.remove(i[1])
sheet._images.remove(i[2])

#sheet['A41'].value = '0'
#sheet['E41'].value = '0'

wb.save(filename)