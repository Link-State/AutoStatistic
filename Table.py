from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Alignment
from openpyxl.styles import Font as CellFont
from openpyxl.styles import Border, Side
from openpyxl.cell.cell import Cell
from Exception import StatisticException
from Stat import Stat
from Category import Category
from Util import *

class Table() :
    def __init__(self, stat:Stat, cat:Category, orderNum:int=1) :
        from Chart import Chart

        self.stat: Stat = stat
        self.targetCategory = cat
        self.title = ""
        self.width = 2
        self.height = 2
        self.area = None
        self.orderNumber = orderNum
        self.matrix = dict()
        self.connectedSheet: Worksheet = None
        self.chart: Chart = None
        return
    
    def fitCellSize(self, cell:Cell) :

        if self.connectedSheet == None :
            return False
        
        cellData:str = str(cell.value)
        if len(cellData) <= 4 :
            return False
        
        size = 13.0 + (1.25 * (len(cellData) - 4))
        
        width:float = self.connectedSheet.column_dimensions[cell.column_letter].width
        height:float = self.connectedSheet.row_dimensions[cell.row].height

        if size <= width :
            return False
        
        self.connectedSheet.column_dimensions[cell.column_letter].width = size

        return True
    
    def fitMergedCellSize(self, range:str) :

        if self.connectedSheet == None :
            return False
        
        self.connectedSheet.unmerge_cells(range)
        matrix = self.connectedSheet[range]

        for row in matrix :
            if len(row) <= 0 :
                continue

            cellData:str = str(row[0].value)
            if len(cellData) <= 4 :
                continue

            target_width = 13.0 + (1.25 * (len(cellData) - 4))
            column_letter_list = list()
            size_list = list()

            for cell in row :
                column_letter:str = cell.column_letter
                cellWidth:float = self.connectedSheet.column_dimensions[column_letter].width
                size_list.append(cellWidth)
                column_letter_list.append((column_letter, cellWidth))
            
            total_width = sum(size_list)
            if total_width >= target_width :
                continue
            
            idx = 0
            size_list = sorted(size_list, reverse=True)
            column_letter_list = sorted(column_letter_list, key=lambda x:(x[1], x[0]))

            for letter, width in column_letter_list :
                augmented_width = width + (target_width - total_width) * (size_list[idx] / total_width)
                self.connectedSheet.column_dimensions[letter].width = augmented_width
                idx += 1
        
        self.connectedSheet.merge_cells(range)

        return True

class FundamentalTable(Table) :
    def __init__(self, stat:Stat, cat:Category, orderNum:int) :
        super().__init__(stat=stat, cat=cat, orderNum=orderNum)
        self.title = "전반적인 " + cat.name
        self.matrix = self.targetCategory.legend.legend
        self.summation_coordinate = ""
        return
    
    def inflate(self, sheet:Worksheet, coor:str) :
        self.connectedSheet = sheet
        startCell:Cell = self.connectedSheet[coor]
        startCell.value = self.title

        catCell = startCell.offset(1, 0)
        valueCell = catCell.offset(0, 1)
        cats = sorted(self.matrix.keys())

        data_font = CellFont(name='맑은 고딕')
        axis_alignment = Alignment(horizontal='center', vertical='center', shrink_to_fit=True)
        data_border = Border(left=Side(border_style='thin'),
                                     top=Side(border_style='thin'),
                                     right=Side(border_style='thin'),
                                     bottom=Side(border_style='thin'))
        total_summation = 0
        for cat in cats :
            value = self.matrix[cat]

            # 각 셀에 데이터 입력
            catCell.value = (str(cat) + self.targetCategory.unit)
            valueCell.value = value
            total_summation += value
            
            # 각 셀에 스타일 적용
            catCell.font = data_font
            catCell.alignment = axis_alignment
            catCell.border = data_border
            self.fitCellSize(catCell)
            
            valueCell.font = data_font
            valueCell.border = data_border

            # 셀 이동
            catCell = catCell.offset(1, 0)
            valueCell = valueCell.offset(1, 0)
        
        # 합계 셀 입력
        catCell.value = "합계"
        valueCell.value = total_summation
        self.summation_coordinate = valueCell.coordinate
        
        # 합계 셀 스타일 적용
        summation_font:CellFont = data_font.__copy__()
        summation_font.bold = True

        catCell.font = summation_font
        catCell.alignment = axis_alignment
        catCell.border = data_border
        self.fitCellSize(catCell)
        
        valueCell.font = summation_font
        valueCell.border = data_border
        
        # 표 범위좌표 저장
        self.height = len(cats) + 2
        LeftTop = coor
        RightBottom = valueCell.offset(-1, 0).coordinate
        self.area = (LeftTop, RightBottom)
        
        return


class CorrelationTable(Table) :
    def __init__(self, stat:Stat, targetCat:Category, comparisonCat:Category, orderNum:int) :
        super().__init__(stat=stat, cat=targetCat, orderNum=orderNum)
        self.width = len(targetCat.legend.legend.keys()) + 1
        self.title = comparisonCat.name + "별 " + targetCat.name
        self.comparisonCategory = comparisonCat
        targetAxis = sorted(list(self.targetCategory.legend.legend.keys()))
        comparisonAxis = sorted(list(self.comparisonCategory.legend.legend.keys()))
        
        targetData = self.targetCategory.legend.data.readLine()
        comparisonData = self.comparisonCategory.legend.data.readLine()

        if len(targetData) <= 0 :
            raise StatisticException(4, self.targetCategory.name)
        if len(comparisonData) <= 0 :
            raise StatisticException(4, self.comparisonCategory.name)

        for target in targetAxis :
            for comparison in comparisonAxis :
                key = (comparison, target)
                
                if len(targetData) != len(comparisonData) :
                    detail = self.targetCategory.name + ", " + self.comparisonCategory.name
                    raise StatisticException(2, detail)
                
                count = 0
                for i in range(len(targetData)) :
                    if targetData[i] == None :
                        continue
                    
                    if target in targetData[i] and comparison in comparisonData[i] :
                        count += 1
                
                self.matrix[key] = count
        return
    
    def inflate(self, sheet:Worksheet, coor:str) :
        self.connectedSheet = sheet
        startCell:Cell = self.connectedSheet[coor]
        startCell.value = self.title

        cats = sorted(self.matrix.keys(), key=lambda x:(x[0], x[1]))
        rowCell = startCell.offset(2, 0)
        colCell = startCell.offset(1, 1)
        diagonalCell = startCell.offset(1, 0)
        rowSet = dict()
        colSet = dict()

        data_font = CellFont(name='맑은 고딕')
        axis_alignment = Alignment(horizontal='center', vertical='center', shrink_to_fit=True)
        data_border = Border(left=Side(border_style='thin'),
                                     top=Side(border_style='thin'),
                                     right=Side(border_style='thin'),
                                     bottom=Side(border_style='thin'))
        diagonal_border = Border(left=Side(border_style='thin'),
                                     top=Side(border_style='thin'),
                                     right=Side(border_style='thin'),
                                     bottom=Side(border_style='thin'),
                                     diagonalDown=True,
                                     diagonal=Side(border_style='thin'))
        
        diagonalCell.border = diagonal_border

        createAverage = isNumber(cats[0][1])
        summation_of_col:dict = dict()
        weighted_summation_of_row:dict = dict()

        for cat in cats :
            if str(rowCell.row) not in weighted_summation_of_row :
                weighted_summation_of_row[str(rowCell.row)] = 0.0
            if colCell.column_letter not in summation_of_col :
                summation_of_col[colCell.column_letter] = 0.0
            if cat[0] not in rowSet :
                rowSet[cat[0]] = str(rowCell.row)
                rowCell.value = (str(cat[0]) + self.comparisonCategory.unit)
                rowCell.font = data_font
                rowCell.alignment = axis_alignment
                rowCell.border = data_border
                self.fitCellSize(rowCell)
                rowCell = rowCell.offset(1, 0)
            if cat[1] not in colSet :
                colSet[cat[1]] = colCell.column_letter
                colCell.value = (str(cat[1]) + self.targetCategory.unit)
                colCell.font = data_font
                colCell.alignment = axis_alignment
                colCell.border = data_border
                self.fitCellSize(colCell)
                colCell = colCell.offset(0, 1)
            
            summation_of_col[colCell.column_letter] += self.matrix[cat]
            if createAverage :
                weighted_summation_of_row[str(rowCell.row)] += float(cat[1]) * self.matrix[cat]
            
            rowIdx = rowSet[cat[0]]
            colIdx = colSet[cat[1]]

            valueCell:Cell = self.connectedSheet[colIdx + rowIdx]
            valueCell.value = self.matrix[cat]
            valueCell.font = data_font
            valueCell.border = data_border

        self.height = len(rowSet) + 2
        LeftTop = coor
        RightBottom = valueCell.coordinate
        self.area = (LeftTop, RightBottom)

        # 합계
        summationCell = valueCell.offset(1, -len(colSet))
        summationCell.value = "합계"
        summationCell.font = data_font
        summationCell.alignment = axis_alignment
        summationCell.border = data_border

        # 평균 점수 (colSet이 숫자형일때만 가능)
        if not createAverage :
            return
        
        averageCell = valueCell.offset(-len(rowSet), 1)
        averageCell.value = "평균 점수"
        averageCell.font = data_font
        averageCell.alignment = axis_alignment
        averageCell.border = data_border


        return
