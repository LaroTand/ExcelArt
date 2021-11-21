import PIL.Image
from openpyxl import Workbook
import openpyxl

# Resize Image
def resize_image(image, new_width=100):
    width, height = image.size
    ratio = height / width
    new_height = int(new_width * ratio)
    resized_image = image.resize((new_width, new_height), PIL.Image.NEAREST)
    return resized_image

def pixel_to_Excelpoint(pixel):
    return pixel / 7

def pixel_to_excel(image):
    workbook = Workbook()
    sheet = workbook.active

    WIDTH = pixel_to_Excelpoint(20)
    X_LENGTH_PIXEL = image.size[0];

    # # Format every column with width 20px
    for col in range(1, sheet.max_column + X_LENGTH_PIXEL):
        sheet.column_dimensions[openpyxl.utils.get_column_letter(col)].width = WIDTH

    # Get the hexadecimal value of each pixel and fill the cells with it
    for i in range(1, image.size[0]):
        for j in range(1, image.size[1]):
            pixel = image.getpixel((i, j))
            # Convert RGB into hexadecimal
            hex_color = '%02x%02x%02x' % (pixel[0], pixel[1], pixel[2])
            # Fill the cell with the hexadecimal value
            sheet.cell(row=j, column=i).fill = openpyxl.styles.PatternFill(patternType="solid", fgColor=hex_color)

    workbook.save("draw.xlsx")


def main():
    image = PIL.Image.open("image.png")
    image = resize_image(image, 100)
    pixel_to_excel(image)


if __name__ == '__main__':
    main()