from PIL import Image
from xlutils.copy import copy
import time
import sys
import xlwt
import xlrd


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
    workbook.save("PixelImage.xls")

    for row in range(0, len(pixels)):
        for column in range(0, len(pixels)):
            (r, g, b) = pixels[row][column]
            print("draw on cell:(" + str(row) + "," + str(column) + ") RGB:(" + str(r) + "," + str(g) + "," + str(b) + ")")
#            workbook = xlrd.open_workbook("PixelImage.xls")
#            workbook = copy(workbook)
#            sheet = workbook.get_sheet(0)
            cellName = str(row) + "*" + str(column)
            print(cellName)
            xlwt.add_palette_colour(cellName, (row * 2 + column) % 52 + 8) 
            workbook.set_colour_RGB((row * 2 + column) % 52 + 8, r, g, b)
            style = xlwt.easyxf('pattern: pattern solid, fore_colour ' + cellName)
            sheet.write(row, column, "", style)
    workbook.save("PixelImage.xls")
    


def main():
    pixels = getPixelsFromImage(Image.open(sys.argv[1]))
    writePixelsIntoExcel(pixels)


if __name__ == "__main__": 
    startTime = time.time()
    main()
    endTime = time.time()
    print("--------------Finished At:" + time.asctime(time.localtime(endTime))+ "-------------")
    print("--------------Time Cost:" + str(endTime - startTime) + "-------------")
