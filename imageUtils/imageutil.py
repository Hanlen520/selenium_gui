# -*-encoding:utf-8-*-
import pytesseract
from PIL import Image

# class GetImageDate(object):
#     def m(self):
#         pytesseract.pytesseract.tesseract_cmd = 'D:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe'
#         image = Image.open("C:\\Users\\Administrator\\Desktop\\p.png")
#         text = pytesseract.image_to_string(image)
#         return text
#
#     def SaveResultToDocument(self):
#         text = self.m()
#         f = open("Verification.txt", "w")
#         print
#         text
#         f.write(str(text))
#         f.close()
# if __name__ == '__main__':
#     g = GetImageDate()
#     g.SaveResultToDocument()

pytesseract.pytesseract.tesseract_cmd = 'D:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe'
image = Image.open('c.png')
vcode = pytesseract.image_to_string(image)
print(vcode)