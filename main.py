import sys
import os
from tempfile import template
import xlwings as xw

OUTPUT_SHEET = "Sheet1"
INPUT_SHEET = "Input"
TEMPLATE_SHEET = "template"
TEMPLATE_SHEET_RANGE = f"A2:K5"
JOB_INFO_RANGE = f"F2:K2"
JOB_PROBS_RANGE = f"A2:D21"
PATH_DIR_CELL = f"G6"
NAME_DIR_CELL = f"F3"
IS_EXPORT_DIR_CELL = f"G5"

USER_CELL = "D5"
CELL_HEIGH = 120

START_ADD_ROW = 11
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



def format_R_cells(sheet, start_row):
    cell_g3 = sheet.range(f"G{start_row+3}")
    cell_g3.value = R_VI + SEPERATOR + R_EN
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
    [internal_job, model, sn, user_en, user_vi, issue_date] = job_infos
    temp_sheet = book.sheets[OUTPUT_SHEET]
    temp_sheet.range("B3").value = internal_job

    # Input User_vi/User_en
    user_cell = temp_sheet.range(USER_CELL)
    text0 = user_vi + SEPERATOR + user_en
    user_cell.value = text0
    index = text0.find(SEPERATOR)
    temp_sheet.range(USER_CELL).characters[0:index].font.bold = True
    temp_sheet.range(USER_CELL).characters[index : len(text0)].font.italic = True
    temp_sheet.range(USER_CELL).characters[index : len(text0)].font.bold = False
    temp_sheet.range(USER_CELL).api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft

    temp_sheet.range("D6").value = model
    temp_sheet.range("I6").value = sn
    temp_sheet.range("J3").value = issue_date

    # Insert suitable image


def copy_insert_row_xlwings(wb: xw.Book, start_row: int, d_row: list):
    """copy from template rows then paste, more efficency than create the whole thing from scratch

    Args:
        wb (xw.Book): The current book
        start_row (int): The row which we want to insert new data into
        d_row (list): list (array) of data get from workbook
    """
    [d_row_part, d_row_probs_vi, d_row_probs_en, d_row_m_or_r] = d_row
    output_sheet = wb.sheets[OUTPUT_SHEET]
    template_sheet = wb.sheets[TEMPLATE_SHEET]


    src_range = template_sheet.range(TEMPLATE_SHEET_RANGE)

    for _ in range(4):
        output_sheet.range(f"A{start_row}").api.EntireRow.Insert()

    # 2. Set row height của row thứ hai (start_row + 1) bằng 120 pixel
    output_sheet.range(f"{start_row + 1}:{start_row + 1}").row_height = CELL_HEIGH

    # 3 Paste into new 4 rows
    dest_range = output_sheet.range(f"A{start_row}:K{start_row+3}")
    src_range.api.Copy(dest_range.api)  # This keeps values + formatting

    # 4 Set row height for A[start_row+1]:K[start_row+1]
    output_sheet.range(f"A{start_row+1}:K{start_row+1}").row_height = 120
    
    # 5. pass input
    cell_a0 = output_sheet.range(f"A{start_row}:B{start_row+3}")
    a0_value = d_row_part
    a0_value_index = a0_value.find(chr(10))
    cell_a0.value = a0_value
    cell_a0.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
    output_sheet.range(f"A{start_row}").characters[0:a0_value_index].font.bold = True
    output_sheet.range(f"A{start_row}").characters[
        a0_value_index : len(a0_value)
    ].font.italic = True
    output_sheet.range(f"A{start_row}").characters[
        a0_value_index : len(a0_value)
    ].font.bold = False

    cell_c2 = output_sheet.range(f"C{start_row+1}:G{start_row+1}")
    c2_value = d_row_probs_vi + chr(10) + d_row_probs_en
    cell_c2.value = c2_value
    c2_value_index = c2_value.find(chr(10))
    cell_c2.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
    output_sheet.range(f"C{start_row+1}").characters[0:c2_value_index].font.bold = False
    output_sheet.range(f"C{start_row+1}").characters[
        c2_value_index : len(c2_value)
    ].font.italic = True
    output_sheet.range(f"C{start_row+1}").characters[
        c2_value_index : len(c2_value)
    ].font.bold = False

    # 6. G[start_row+3]: text trắng, căn phải
    if d_row_m_or_r == 1:
        format_M_cells(sheet=output_sheet, start_row=start_row)
    else:
        format_R_cells(sheet=output_sheet, start_row=start_row)



def insert_conclusion(wb:xw.Book, data):
    pass


def export2Dir(wb:xw.Book, path:str, name:str):
    # Copy Sheet1
    print(f"Export path: {path}\nName: {name}")
    copy_wb = wb.sheets[OUTPUT_SHEET]

    # Create dir 
    dir_path =path+"\\"+name 
    if not os.path.exists(dir_path):
        os.mkdir(dir_path)

    output_wb = xw.Book()
    # Paste a copy to path dir
    copy_wb.api.Copy(Before=output_wb.sheets[0].api)
    output_wb.save(dir_path+"\\"+"TR "+name+".xlsx")
    output_wb.sheets["Sheet1"].delete()


def ExcelProcess(file: str):
    book = xw.Book(file)
    sheet = book.sheets[INPUT_SHEET]
    job_infos = sheet.range(JOB_INFO_RANGE).value
    all_data = sheet.range(JOB_PROBS_RANGE).value

    insert_job_info(book, job_infos)
    # insert_job_conclusion(book, job_infos)

    none_arr = [None] * len(all_data[0])
    for d_row in all_data[::-1]:
        if d_row != none_arr:
            copy_insert_row_xlwings(book, START_ADD_ROW, d_row)
            # insert_row_xlwings(book, TEMPLATE_SHEET, START_ADD_ROW)
    # copy to new file
    is_export = sheet.range(IS_EXPORT_DIR_CELL).value
    if is_export:
        export_path = sheet.range(PATH_DIR_CELL).value
        export_name = sheet.range(NAME_DIR_CELL).value
        export2Dir(wb=book, path=export_path, name=export_name)


# filepath = "Book1.xlsm"

def main():
    filepath = sys.argv[1]
    # print("Processing:", filepath)
    ExcelProcess(file=filepath)


if __name__ == "__main__":
    main()
