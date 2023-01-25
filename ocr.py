import cv2
import pytesseract

# lendo imagem com opencv
img = cv2.imread("meunome.png")
# apontar onde está o executável do tesseract

caminho = r"C:\Users\Priscila\AppData\Local\Programs\Tesseract-OCR"
pytesseract.pytesseract.tesseract_cmd = caminho + r"\tesseract.exe"

resultado = pytesseract.image_to_string(img, lang="por")

print(resultado)
