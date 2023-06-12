from skimage import io
import pytesseract


# Read name of Team
img = io.imread('./scan_ahickers_2023-06-09-15-52-40_1.jpeg')
img = img[240:300, 610:2000]

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
team = pytesseract.image_to_string(img)
print("TEAM:", team)



# Read names of Presenters
img = io.imread('./scan_ahickers_2023-06-09-15-52-40_2.jpeg')
name1 = img[330:400,   700:1000]
name2 = img[1280:1340, 700:1000]
name3 = img[2220:2290, 700:1000]

print(pytesseract.image_to_string(name1))
print(pytesseract.image_to_string(name2))
print(pytesseract.image_to_string(name3))



