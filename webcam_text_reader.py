## webcam_text_reader.py from Vincent Fischer

## Inspired from https://vimeo.com/643514582 


## pip install pytesseract
## pip install opencv-python
## pip install openpyxl
from openpyxl import Workbook
import cv2
import pytesseract

## Path to modify -> https://www.youtube.com/watch?v=PY_N1XdFp4w -> install from https://github.com/UB-Mannheim/tesseract/wiki 
pytesseract.pytesseract.tesseract_cmd = r'E:\Code\Tesseract (OCR)\\tesseract.exe'
wb = Workbook()

def readWebcamPic():
    text = pytesseract.image_to_string(frame)
    if len(text) > 3:
        return text
        


names = ['courses1.jpg','courses2.jpg','courses3.jpg']
# names = ['c1.jpg','c2.jpg']
# names = ['prixc2.jpg']

cap = cv2.VideoCapture('1bis.jpg')

if not cap.isOpened():
    raise IOError("Cannot Open Webcam")

pic_index = 0 
row_index = 0
make_excel = True          ## Si True on fait l'excel, sinon False

while pic_index < len(names):
    cap = cv2.VideoCapture(names[pic_index])
    ret, frame = cap.read()
    frame = cv2.resize(frame, None, fx=0.5, fy = 0.5, interpolation=cv2.INTER_AREA)

    try:
        text = readWebcamPic()
        with open('out.txt', 'w') as f:
            print(text, file=f)

        if make_excel :   
            sentences = text.split("\n")                    ## https://stackoverflow.com/questions/30484252/pytesser-next-line-of-text-in-image 
            ws = wb.active
            for index in range(len(sentences)):
                ws['A'+str(1+row_index)] = sentences[index]
                
                if sentences[index][-2:].isdigit() and sentences[index][-4].isdigit() :
                    if sentences[index][-5:].isdigit() :
                        ws['B'+str(1+row_index)] = sentences[index][:-11]
                        ws['C'+str(1+row_index)] = float(sentences[index][-5:-3]) + float(sentences[index][-2:])/100
                    else :
                        ws['B'+str(1+row_index)] = sentences[index][:-11]
                        ws['C'+str(1+row_index)] = float(sentences[index][-4]) + float(sentences[index][-2:])/100
                else :
                    ws['B'+str(1+row_index)] = sentences[index]
                    ws['C'+str(1+row_index)] = 'No price'
                row_index += 1

            if pic_index < len(names)-1 :
                row_index += 1
                ws['A'+str(1+row_index)] = 'Photo numero ' + str(pic_index+2)
                row_index += 3
                pic_index += 1
            else :
                pic_index += 1 
        else :
            pic_index += 1 
    except:
        print("Something didn't work...")
        pic_index += 1
        pass

    ## https://stackoverflow.com/questions/14494101/using-other-keys-for-the-waitkey-function-of-opencv 
    # if cv2.waitKey(1) == ord('a'):
    #     print("Exiting Webcam")
    #     break

if make_excel:
    wb.save('sample.xlsx')
cap.release()
cv2.destroyAllWindows