import sys
import os

from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell
from openpyxl.styles import Alignment
from openpyxl.styles import Font as CellFont
from openpyxl.styles import Border, Side

class Survey() :
    def __init__(self, fname:str="설문지") :

        idx = 0
        postfix = ""
        while True :
            if not os.path.exists(f"./{fname}{postfix}.xlsx") :
                break

            idx += 1
            postfix = f"({idx})"
        
        self.fileDir = f"./{fname}{postfix}.xlsx"

        self.workbook = Workbook()
        self.worksheet = self.workbook.active
        self.worksheet.title = f"{fname}"

        self.worksheet.row_dimensions[2].height = 30

        titleCell:Cell = self.worksheet["B2"]
        title_font = CellFont(name='맑은 고딕', size=16, bold=True)
        title_alignment = Alignment(horizontal='center', vertical='center', shrink_to_fit=True)
        title_border = Border(left=Side(border_style='thin'),
                              top=Side(border_style='thin'),
                              right=Side(border_style='thin'),
                              bottom=Side(border_style='thin'))
        
        cells = self.worksheet["B2:F2"]
        for row in cells :
            for cell in row :
                cell:Cell = cell
                cell.font = title_font
                cell.alignment = title_alignment
                cell.border = title_border
        
        titleCell.value = "제목"
        self.worksheet.merge_cells("B2:F2")

        catCell:Cell = self.worksheet["B3"]
        cat_alignment = Alignment(horizontal='center', vertical='center')
        cat_border = Border(top=Side(border_style='thin'),
                            bottom=Side(border_style='thin'))
        for i in range(7) :
            catCell.value = "질문" + str(i)
            catCell.alignment = cat_alignment
            catCell.border = cat_border
            self.worksheet.column_dimensions[catCell.column_letter].width = 23.5
            catCell = catCell.offset(0, 1)

        return
    
    def create(self) :
        self.workbook.save(self.fileDir)
        self.workbook.close()
        return