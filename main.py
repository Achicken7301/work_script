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

INFO_ROW = 1
JOB_ROW = 5
ENDUSER_ROW = 6
MODEL_ROW = 7
SN_ROW = 8

JOB_INFO_RANGE = f"F1:I3"


def insert_infos():
    pass


def insert_conclusion():
    pass


def format_R_cells(sheet, start_row):
    cell_g3 = sheet.range(f"G{start_row+3}")
    cell_g3.value = M_VI + SEPERATOR + M_EN
    # cell_g3.font.color = (255, 255, 255)  # WHITE
    cell_g3.font.color = (0, 0, 0)  # Black
    sheet.range(f"G{start_row+3}").characters[0 : len(cell_g3.value)].font.bold = False
    cell_g3.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignRight

    # 7. Màu background cell từ C[start_row+3] tới G[start_row+3]
    rng = sheet.range(f"C{start_row+3}:G{start_row+3}")
    # rng.color = (0, 112, 192)  # màu xanh (blue)
    rng.color = (255, 255, 0)  # màu vàng (yellow)


def format_M_cells(sheet, start_row):
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


def insert_row_xlwings(wb, sheetname, start_row, data: list):
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
    cell_a0.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
    sheet.range(f"A{start_row}").characters[0:a0_value_index].font.bold = True
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
    cell_c2.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
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
    if data[DATA_M_R] == 1:
        format_M_cells(sheet=sheet, start_row=start_row)
    else:
        format_R_cells(sheet=sheet, start_row=start_row)

    # 8. Create border
    cell_C0_G4 = sheet.range(f"C{start_row}:G{start_row+3}").api
    for border_id in (
        xw.constants.BordersIndex.xlEdgeTop,
        xw.constants.BordersIndex.xlEdgeBottom,
        xw.constants.BordersIndex.xlEdgeRight,
        xw.constants.BordersIndex.xlEdgeLeft,
        xw.constants.BordersIndex.xlInsideHorizontal,
    ):
        cell_C0_G4.Borders(border_id).LineStyle = 1  # 1 for xlContinuous
        cell_C0_G4.Borders(border_id).Weight = 2  # 2 for xlThin (adjust as needed)

    cell_C0_G4 = sheet.range(f"A{start_row}:K{start_row+3}").api
    for border_id in (
        xw.constants.BordersIndex.xlEdgeTop,
        xw.constants.BordersIndex.xlEdgeBottom,
        xw.constants.BordersIndex.xlEdgeRight,
        xw.constants.BordersIndex.xlEdgeLeft,
    ):
        cell_C0_G4.Borders(border_id).LineStyle = 1  # 1 for xlContinuous
        cell_C0_G4.Borders(border_id).Weight = (
            xw.constants.BorderWeight.xlMedium
        )  # 2 for xlThin (adjust as needed)

def insert_job_info(book, job_infos):
    # hard code but no choice
    internal_job = job_infos[1][0]
    user_en = job_infos[1][1]
    user_vi = job_infos[2][1]
    model = job_infos[1][2]
    sn = job_infos[1][3]

    temp_sheet = book.sheets[TEMPLATE_SHEET]
    temp_sheet.range("B3").value = internal_job

    # Input User_vi/User_en
    user_cell = temp_sheet.range(f"D5")
    text0 = user_vi + SEPERATOR + user_en
    user_cell.value = text0
    index = text0.find(SEPERATOR)
    temp_sheet.range(f"D5").characters[0:index].font.bold = False
    temp_sheet.range(f"D5").characters[index : len(text0)].font.italic = True
    temp_sheet.range(f"D5").characters[index : len(text0)].font.bold = False
    temp_sheet.range(f"D5").api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft

    temp_sheet.range("D6").value = model
    temp_sheet.range("I6").value = sn

    # Insert suitable image


def ExcelProcess(file: str):
    book = xw.Book(file)
    sheet = book.sheets[INPUT_SHEET]
    job_infos = sheet.range(JOB_INFO_RANGE).value
    insert_job_info(book, job_infos)
    # insert_job_conclusion(book, job_infos)

    all_data = sheet.used_range.value
    none_arr = [None] * len(all_data[0])
    for d_row in all_data[::-1]:
        if d_row != none_arr:
            insert_row_xlwings(book, TEMPLATE_SHEET, START_ADD_ROW, d_row)


# filepath = "Book1_copy.xlsm"




def main():
    filepath = sys.argv[1]
    print("Processing:", filepath)
    ExcelProcess(file=filepath)


if __name__ == "__main__":
    main()
