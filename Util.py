from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell
from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt

def isNumber(x:str) -> bool :
    if x.isdecimal() :
        return True
    
    try :
        y = float(x)
    except ValueError :
        return False
    
    return True

def isSameInteger(x) -> bool :
    if not isNumber(str(x)) :
        return False
    
    y = float(x)

    if y == float(int(y)) :
        return True
    else :
        return False

def resizeCellWidth(work_sheet:Worksheet, coordinate:str) :
    isMerged = False
    if coordinate.find(":") >= 0 :
        work_sheet.unmerge_cells(coordinate)
        matrix:tuple[tuple[Cell]] = work_sheet[coordinate]
        isMerged = True
    else :
        matrix:tuple[tuple[Cell]] = ((work_sheet[coordinate],),)
    
    
    for row in matrix :
        if len(row) <= 0 :
            continue

        cell_value:str = str(row[0].value)
        if len(cell_value) <= 4 :
            continue

        target_width = 13.0 + (1.25 * (len(cell_value) - 4))
        column_letter_list:list[tuple[str, float]] = list()

        total_width = 0.0
        for cell in row :
            column_letter:str = cell.column_letter
            cell_width:float = work_sheet.column_dimensions[column_letter].width
            column_letter_list.append((column_letter, cell_width))
            total_width += cell_width
        
        if total_width >= target_width :
            continue
        
        column_letter_list = sorted(column_letter_list, key=lambda x:(x[1], x[0]))

        for idx, (letter, width) in enumerate(column_letter_list) :
            augmented_width = width + (target_width - total_width) * (column_letter_list[-(idx+1)][1] / total_width)
            work_sheet.column_dimensions[letter].width = augmented_width
    
    if isMerged :
        work_sheet.merge_cells(coordinate)

    return

def clearLayout(layout:QLayout) :
    if layout is not None :
        while layout.count() :
            item = layout.takeAt(0)
            widget = item.widget()
            if widget is not None :
                widget.setParent(None)
            else :
                clearLayout(item.layout())
    return