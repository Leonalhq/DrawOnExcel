from PIL import Image
from xlutils.copy import copy
import time
import sys
import xlwt
import xlrd
import math

imageName = sys.argv[1]
mergeSize = int(sys.argv[2])         #to merge a size of pixel into one color

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
    workbook.save("./Output" + imageName + ".xls")

    for row in range(0, int(math.ceil(len(pixels) / mergeSize))):
        for column in range(0, int(math.ceil(len(pixels[0]) / mergeSize))):
            (r, g, b) = pixels[row * mergeSize][column * mergeSize]
            print("draw on cell:(" + str(row * mergeSize) + "," + str(column * mergeSize) + ") RGB:(" + str(r) + "," + str(g) + "," + str(b) + ")")
            (colorNum, color) = matchSimilarColor((r, g, b))
            xlwt.add_palette_colour(str(colorNum), colorNum) 
            workbook.set_colour_RGB(colorNum, color[0], color[1], color[2])
            style = xlwt.easyxf('pattern: pattern solid, fore_colour ' + str(colorNum))
            
            for mergedRow in range(row * mergeSize, row * mergeSize + mergeSize):
                for mergedColumn in range(column * mergeSize, column * mergeSize + mergeSize):
                    sheet.write(mergedRow, mergedColumn, "", style)
    workbook.save("Output/" + imageName + "." + str(mergeSize) + ".xls")


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
    pixels = getPixelsFromImage(Image.open("./Input/" + imageName))
    writePixelsIntoExcel(pixels)


if __name__ == "__main__": 
    startTime = time.time()
    main()
    endTime = time.time()
    print("--------------Finished At:" + time.asctime(time.localtime(endTime))+ "-------------")
    print("--------------Time Cost:" + str(endTime - startTime) + "-------------")
