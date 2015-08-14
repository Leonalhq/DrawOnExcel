from PIL import Image
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

    for row in range(0, len(pixels)):
        for column in range(0, len(pixels)):
            (r, g, b) = pixels[row][column]
            print("draw on cell:(" + str(row) + "," + str(column) + ") RGB:(" + str(r) + "," + str(g) + "," + str(b) + ")")
            cellName = str(row) + "*" + str(column)
            xlwt.add_palette_colour(cellName, 0x21)
            workbook.set_colour_RGB(0x21, r, g, b)
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
