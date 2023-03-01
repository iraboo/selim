import pytesseract
from PIL import Image
import cv2

current_folder = 'C:/Users/iraboo/Documents/my_project/selim/'
filename = current_folder + '검침.jpg'

img = Image.open(filename)
result = pytesseract.image_to_string(img,lang='kor')

#img = cv2.imread(filename, cv2.IMREAD_COLOR)
#result = pytesseract.image_to_string(img)
print(result)