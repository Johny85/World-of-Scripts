from pdf2image import convert_from_path
from pytesseract import pytesseract
from PIL import Image
from pytesseract import image_to_string
from pyzbar.pyzbar import decode

tesseract_path = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
pytesseract.tesseract_cmd = tesseract_path

def convert_pdf_to_img(pdf_file):
    return convert_from_path(pdf_file, poppler_path=r'C:\\poppler-0.68.0\\bin')

def convert_image_to_text(file):
    text = image_to_string(file)
    return text

def get_text_from_any_pdf(pdf_file):
    images = convert_pdf_to_img(pdf_file)
    final_text = ""
    for pg, img in enumerate(images):
        final_text += convert_image_to_text(img)
    
    return final_text

path_to_pdf = 'Pan4.pdf'

print(get_text_from_any_pdf(path_to_pdf))

image = convert_from_path(path_to_pdf, poppler_path=r'C:\\poppler-0.68.0\\bin')

for page in image:
    page.save('out.jpg', 'JPEG')

img = Image.open('out.jpg')
result = decode(img)
for i in result:
    print(i.data.decode("utf-8"))