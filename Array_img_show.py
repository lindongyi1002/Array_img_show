import openpyxl
import numpy as np
import matplotlib.pyplot as pyplot
import os

# only run grayscale division

if __name__ == "__main__":

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    File_Dir = r"03-修正.xlsx"
    File_Dir = r"230318红外成像测试/灰度划分.xlsx"

    # read 64*64 excel sheet
    filepath = os.path.join(BASE_DIR, File_Dir)
    sheetname = 'Sheet1'
    workbook = openpyxl.load_workbook(filepath)
    sheet = workbook[sheetname]

    # row number & col number
    num_rows = 63
    num_cols = 63

    # generate image matrix
    def gener_img_mat(num_rows, num_cols):
        img_mat = np.zeros((num_rows, num_cols), dtype=np.float64)
        for rown in range(num_rows):
            for coln in range(num_cols):
                img_mat[rown][coln] = sheet.cell(row=rown + 1, column=coln + 1).value

        # find the max data and print
        max_value = 0
        row_coord = None
        col_coord = None
        for i, row in enumerate(img_mat):
            for j, value in enumerate(row):
                if value >= max_value:
                    max_value = value
                    row_coord = i + 1  # only for understand easily
                    col_coord = j + 1  # only for understand easily
                else:
                    max_value = max_value

        print(f'Max Value: {max_value}')
        print(f'The Coord of Max Value: ({row_coord}, {col_coord})\n')
        return img_mat

    # grayscale division
    def gray_div(img_mat, gray_scl, lsb, msb):
        img_divide_done = np.zeros((num_rows, num_cols), dtype=np.int32)
        for rown in range(num_rows):
            for coln in range(num_cols):
                if img_mat[rown][coln] < lsb:
                    img_divide_done[rown][coln] = 0
                elif img_mat[rown][coln] > lsb + ((msb - lsb) / gray_scl) * gray_scl:
                    img_divide_done[rown][coln] = gray_scl
                else:
                    for i in range(1, gray_scl, 1):
                        if lsb + ((msb - lsb) / gray_scl) * (i - 1) <= img_mat[rown][coln] <= lsb + (
                                (msb - lsb) / gray_scl) * i:
                            img_divide_done[rown][coln] = 1 + (i - 1)
                            break
                        else:
                            img_divide_done[rown][coln] = img_divide_done[rown][coln]
        return img_divide_done

    # show a gray image
    def out_img(fig1, fig2, gray1=16363, gray2=255):
        # fig1
        # pyplot.subplot2grid((1, 2), (0, 0))
        # pyplot.title('Origin Image')
        # pyplot.imshow(fig1, vmin=0, vmax=gray1)
        # pyplot.gray()
        # pyplot.colorbar()
        # fig2
        # pyplot.subplot2grid((1, 2), (0, 1))
        pyplot.title('Gray Divided Image')
        pyplot.imshow(fig2, vmin=0, vmax=gray2)
        pyplot.gray()
        pyplot.colorbar()
        # pyplot.savefig(os.path.join(BASE_DIR, r'save_data\Image.png'))
        pyplot.show()


    img_mat = gener_img_mat(num_rows, num_cols)

    fig1_gray = 16383
    fig2_gray = 255
    img_divide_done = gray_div(img_mat, fig2_gray, 4900, 7500)

    out_img(img_mat, img_divide_done, gray1=fig1_gray, gray2=fig2_gray)
