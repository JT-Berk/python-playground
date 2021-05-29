from PIL import ImageGrab
import win32com.client as win32
import os

cheating = True

excel = win32.gencache.EnsureDispatch('Excel.Application')
workbook = excel.Workbooks.Open(r'your_path_here')

# for loop to work through all the worksheets
for sheet in workbook.Worksheets:
    # finds all embedded objects in xlsx file
    for i, shape in enumerate(sheet.Shapes):
        if shape.Name.startswith('Picture'):

            # assuming that each item has three images associated with it
            if i % 3 == 0:
                # makes new folder for piece of inventory
                myfolder = "your_path_here" + str(int(i/3))
                os.mkdir(myfolder)

            # grabs images by order of when they were put on the excel file by the user
            # might be good to find where this information comes from and how to manipulate
            shape.Copy()
            image = ImageGrab.grabclipboard()

            # accounting for the one instance of an item having only two images
            if cheating & i == 131:
                myfolder = "your_path_here" + str(int(i/3))
                os.mkdir(myfolder)

            # save to location
            image.save(myfolder + "/" + str(i) + ".png")

            # Pillow has issue converting to .jpg - something with RGB?
            # image.save('{}.jpg'.format(i+1), 'jpeg')
