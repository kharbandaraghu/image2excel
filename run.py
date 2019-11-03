# Necessary imports
import cv2
import xlsxwriter

# This function is supposed to convert RGB values to hex
# On more info how we created this visit: http://psychocodes.in/без-рубрики/rgb-to-hex-conversion-and-hex-to-rgb-conversion-in-python/
def rgb2hex(r,g,b):
    return "#{:02x}{:02x}{:02x}".format(r,g,b)

# Now we will read the image
# make sure the image is in the same directory as your python file
# ENTER THE NAME OF YOUR FILE HERE
image = cv2.imread("photo.jpg")   # when reading the image the image original size is 150x150
print("The image size is: " + str(image.shape))

# Usually images are of larger size wich will slow down the program
# We will resize the image to your preferred size
inp = input("Enter 's' to input size or anything else to use default size: ")

# creating default size
default_size = (75,150)

if inp == 's':
    h = input('Enter Height: ')
    w = input('Enter Width: ')
    default_size = (int(w),int(h))

scaled_image = cv2.resize(image, default_size)  # when scaling we scale original image to 24x24 
print("The new size is: " + str(scaled_image.shape))

# Now creating workbook
workbook = xlsxwriter.Workbook('output.xlsx')
worksheet = workbook.add_worksheet()

# Looping through rows and columns
for row in range(scaled_image.shape[0]):
    # we will print also whats being processed
    print("Processing row "+str(row))
    for col in range(scaled_image.shape[1]):

        cell_format = workbook.add_format()
        r = scaled_image[row,col,0]
        g = scaled_image[row,col,1]
        b = scaled_image[row,col,2]
        # Getting hex code
        color = rgb2hex(r,g,b)
        cell_format.set_bg_color(color)
        # you can have the cells say whatever you want
        worksheet.write(row, col, 'hello', cell_format)

# closing the spreadsheet
workbook.close()