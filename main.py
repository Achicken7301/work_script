import array
from asyncio.windows_events import NULL
import sys
import xlwings as xw

TEMPLATE_SHEET = "template"
INPUT_SHEET = "Input"
CELL_HEIGH = 120

START_ADD_ROW = 13
PHENO_VI = "Hiện Trạng"
PHENO_EN = "Phenomenon"
JUD_VI = "Đánh giá"
JUD_EN = "Judgement"
M_VI = "Yêu cầu sửa chữa"
M_EN = "Repair required"
R_VI = "Khuyến nghị sửa chữa"
R_EN = "Repair recommended"
SEPERATOR = " / "

DATA_PART_POS = 0
DATA_VI = 1
DATA_EN = 2
DATA_M_R = 3


def insert_row_xlwings(wb, sheetname, start_row, data: array):
    # app = xw.App(visible=False)
    # wb = app.books.open(filepath)
    # wb = xw.apps.active.books[filepath]
    sheet = wb.sheets[sheetname]
    # print(
    #     f"Receive data: {data[DATA_PART_POS]},{data[DATA_VI]},{data[DATA_EN]},{data[DATA_M_R]}"
    # )

    # 1. Insert 4 new rows tại start_row
    for _ in range(4):
        sheet.range(f"A{start_row}").api.EntireRow.Insert()

    # 2. Set row height của row thứ hai (start_row + 1) bằng 120 pixel
    sheet.range(f"{start_row + 1}:{start_row + 1}").row_height = CELL_HEIGH

    # 3. Merge cells and pass input
    sheet.range(f"A{start_row}:B{start_row+3}").merge()
    cell_a0 = sheet.range(f"A{start_row}:B{start_row+3}")
    a0_value = data[DATA_PART_POS]
    a0_value_index = a0_value.find(chr(10))
    cell_a0.value = a0_value
    sheet.range(f"A{start_row}").characters[0:a0_value_index].font.bold = False
    sheet.range(f"A{start_row}").characters[
        a0_value_index : len(a0_value)
    ].font.italic = True
    sheet.range(f"A{start_row}").characters[
        a0_value_index : len(a0_value)
    ].font.bold = False

    sheet.range(f"C{start_row+1}:G{start_row+1}").merge()
    cell_c2 = sheet.range(f"C{start_row+1}:G{start_row+1}")
    c2_value = data[DATA_VI] + chr(10) + data[DATA_EN]
    cell_c2.value = c2_value
    c2_value_index = c2_value.find(chr(10))
    sheet.range(f"C{start_row+1}").characters[0:c2_value_index].font.bold = False
    sheet.range(f"C{start_row+1}").characters[
        c2_value_index : len(c2_value)
    ].font.italic = True
    sheet.range(f"C{start_row+1}").characters[
        c2_value_index : len(c2_value)
    ].font.bold = False

    sheet.range(f"H{start_row}:K{start_row+3}").merge()

    # 4. C[start_row]: text "ABC / DEF", bold cho phần ABC, italic phần DEF, căn trái
    cell_c0 = sheet.range(f"C{start_row}")
    text0 = PHENO_VI + SEPERATOR + PHENO_EN
    cell_c0.value = text0
    index = text0.find(SEPERATOR)
    sheet.range(f"C{start_row}").characters[0:index].font.bold = True
    sheet.range(f"C{start_row}").characters[index : len(text0)].font.italic = True
    sheet.range(f"C{start_row}").characters[index : len(text0)].font.bold = False
    sheet.range(f"C{start_row}").api.HorizontalAlignment = (
        xw.constants.HAlign.xlHAlignLeft
    )

    # 5. C[start_row+2]: tương tự version trên
    cell_c2 = sheet.range(f"C{start_row+2}")
    text0 = JUD_VI + SEPERATOR + JUD_EN
    cell_c2.value = text0
    index = text0.find(SEPERATOR)
    sheet.range(f"C{start_row+2}").characters[0:index].font.bold = True
    sheet.range(f"C{start_row+2}").characters[index : len(text0)].font.italic = True
    sheet.range(f"C{start_row+2}").characters[index : len(text0)].font.bold = False
    sheet.range(f"C{start_row+2}").api.HorizontalAlignment = (
        xw.constants.HAlign.xlHAlignLeft
    )

    # 6. G[start_row+3]: text trắng, căn phải
    cell_g3 = sheet.range(f"G{start_row+3}")
    cell_g3.value = M_VI + SEPERATOR + M_EN
    cell_g3.font.color = (255, 255, 255)  # WHITE
    # cell_g3.font.color = (0,0,0) # Black
    sheet.range(f"G{start_row+3}").characters[0 : len(cell_g3.value)].font.bold = False
    cell_g3.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignRight

    # 7. Màu background cell từ C[start_row+3] tới G[start_row+3]
    rng = sheet.range(f"C{start_row+3}:G{start_row+3}")
    rng.color = (0, 112, 192)  # màu xanh (blue)
    # rng.color = (255, 255, 0)  # màu vàng (yellow)


def ExcelProcess(file: str):
    book = xw.Book(file)
    sheet = book.sheets[INPUT_SHEET]
    all_data = sheet.used_range.value
    print(all_data)
    for d_row in all_data:
        if d_row != [None, None, None, None]:
            insert_row_xlwings(book, TEMPLATE_SHEET, START_ADD_ROW, d_row)


filepath = "Book1_copy.xlsm"


def main():
    # filepath = sys.argv[1]
    print("Processing:", filepath)
    ExcelProcess(file=filepath)


if __name__ == "__main__":
    main()
