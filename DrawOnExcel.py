from PIL import Image
from xlutils.copy import copy
import time
import sys
import xlwt
import xlrd
import math


def getPixelsFromImage(image):
    pixels = []
    for row in range(0, image.size[1]):
        pixelsPerRow = []
        for column in range(0, image.size[0]):
            pixelsPerRow.append(image.getpixel((column, row)))
        pixels.append(pixelsPerRow)

    return pixels

def writePixelsIntoExcel(pixels):
    workbook = xlwt.Workbook(style_compression=2)  
    sheet = workbook.add_sheet('PixelImage')
    workbook.save(sys.argv[1] + ".xls")

    for row in range(0, len(pixels)):
        for column in range(0, len(pixels[0])):
            (r, g, b) = pixels[row][column]
            print("draw on cell:(" + str(row) + "," + str(column) + ") RGB:(" + str(r) + "," + str(g) + "," + str(b) + ")")
            (colorName, color) = matchSimilarColor((r, g, b))
            xlwt.add_palette_colour(str(colorName), colorName) 
            workbook.set_colour_RGB(colorName, color[0], color[1], color[2])
            style = xlwt.easyxf('pattern: pattern solid, fore_colour ' + str(colorName))
            sheet.write(row, column, "", style)
    workbook.save(sys.argv[1] + ".xls")


#xlwt only support 56 colors
def matchSimilarColor(rgb):
    (r, g, b) = rgb
    colorNum = 0
    color = [] 
    for i in range(0, len(rgb)):
        color.append(rgb[i] / 86 * 85 + 43)
        colorNum = int(colorNum + rgb[i] / 86 * math.pow(3, i) + 8)
    return (colorNum, color)


def main():
    pixels = getPixelsFromImage(Image.open(sys.argv[1]))
    writePixelsIntoExcel(pixels)


if __name__ == "__main__": 
    startTime = time.time()
    main()
    endTime = time.time()
    print("--------------Finished At:" + time.asctime(time.localtime(endTime))+ "-------------")
    print("--------------Time Cost:" + str(endTime - startTime) + "-------------")
