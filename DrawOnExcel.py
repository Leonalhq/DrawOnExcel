from PIL import Image
import time
import sys


def getPixelsFromImage(image):
    pixels = []
    for row in range(0, image.size[1]):
        pixelsPerRow = []
        for column in range(0, image.size[0]):
            pixelsPerRow.append(image.getpixel((column, row)))
        pixels.append(pixelsPerRow)

    return pixels
 

def main():
    pixels = getPixelsFromImage(Image.open(sys.argv[1]))
    print(pixels)


if __name__ == "__main__": 
    startTime = time.time()
    main()
    endTime = time.time()
    print("\n--------------Finished At:" + time.asctime(time.localtime(endTime))+ "-------------")
    print("--------------Time Cost:" + str(endTime - startTime) + "-------------")
