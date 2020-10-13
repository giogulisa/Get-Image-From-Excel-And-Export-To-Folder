from PIL import ImageGrab
import re
import win32com.client as win32
import os

excel = win32.gencache.EnsureDispatch('Excel.Application')
workbook = excel.Workbooks.Open(r'C:\Users\Gio\Desktop\t\t1.xls')

list = []

parent_dir = "C:/Users/Gio/Desktop/Gulisa"

for sheet in workbook.Worksheets:
    for i, shape in enumerate(sheet.Shapes):
        if shape.Name.startswith('Picture'):
            #shape.Copy()

            #print(shape.TopLeftCell.Address) დასაწყისი
            #print(shape.BottomRightCell.Address) დასასრული

            # დასაწყისი
            cellStart = shape.TopLeftCell.Address
            # დასასრული
            cellEnd = shape.BottomRightCell.Address
            # დასაწყისი ინტში
            cellStartInt = cellStart[3:]
            #print(cellStartInt)
            # დასასრული ინტში
            cellEndInt = cellEnd[3:]
            #print(cellEndInt)

            directory = ""

            if int(cellStartInt) > 0:
                if cellStartInt < cellEndInt:
                    for i in range(int(cellStartInt), int(cellEndInt)):
                        directory = sheet.Cells(i, 'A')
                        dir1 = "gio#" + str(directory)
                        path = os.path.join(str(parent_dir), str(dir1))
                        if (str(directory) not in list):
                            if shape.Name.startswith('Picture'):
                                shape.Copy()
                                list.append(str(directory))
                                os.mkdir(path)
                                image = ImageGrab.grabclipboard()
                                img = path + "/img.JPEG"
                                image.save(img, 'png')

                else:
                    directory = sheet.Cells(cellStartInt, 'A')
                    dir1 = "gio#" + str(directory)
                    path = os.path.join(str(parent_dir), str(dir1))
                    if(str(directory) not in list):
                        if shape.Name.startswith('Picture'):
                            shape.Copy()
                            list.append(str(directory))
                            os.mkdir(path)
                            image = ImageGrab.grabclipboard()
                            img = path + "/img.JPEG"
                            image.save(img, 'png')
