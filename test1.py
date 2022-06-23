import os
import numpy as np
from PIL import Image
import openpyxl

wb = openpyxl.load_workbook('test.xlsx')
ws = ['Sheet']

path = './mnist_mini_1000'
files = os.listdir(path)
for file in files:
    if 'num4' in file:
        imgPath = os.path.join(path, file)
        img = Image.open(imagPath)
        imgArray = np.array(img)

        for i in range(9):
            for j in range(9):
                ws.cell(row=1+i, column=1+j+9*imgDone).value = imagArray[i][j]
        imgDone = imgDone + 1
wb.save(test.xlsx')