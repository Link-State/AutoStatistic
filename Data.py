
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell
from Stat import Stat
from Exception import StatisticException
from Util import *

class Data() :
    def __init__(self, stat:Stat, cell:Cell=None, moveDir:tuple[int, int]=None) :
        self.stat = stat
        self.startCell = cell
        self.moveDir = moveDir
        self.data = list()

        if self.startCell == None :
            self.startCell = self.stat.dataStartPoint

        if self.moveDir == None :
            self.moveDir = self.stat.moveDir
        
        return
    
    def getCell(self, coor:tuple[int, int]=None, name:str=None, cell:Cell=None) -> Cell | None :
        inquirySheetName:str = self.stat.referenceExcel.sheetnames[0]
        inquirySheet:Worksheet = self.stat.referenceExcel[inquirySheetName]

        if coor != None :
            newCell:Cell = inquirySheet.cell(coor[0], coor[1])
            return newCell
        
        if name != None :
            newCell:Cell = inquirySheet[name]
            return newCell
        
        if cell != None :
            newCell:Cell = inquirySheet[cell.coordinate]
            return newCell

        return None
    
    def readLine(self, startCell:Cell=None, moveDir:tuple[int, int]=None, container=True) -> list :
        if startCell == None :
            startCell = self.getCell(cell=self.startCell)
        if moveDir == None :
            moveDir = self.moveDir
        multiSelectSplitSymbol: str = self.stat.multiSelectSplitSymbol

        # -1 = default, 0 = int|float, 1 = string
        originType = -1
        result = list()
        while True :
            if type(startCell.value) == type(None) :
                break
            
            data = str(startCell.value).strip()
            if data == "" :
                break

            isNull = False
            
            currentType = -1
            if isNumber(data) :
                if data.isdecimal() :
                    data = [int(data)]
                else :
                    data = [float(data)]
                currentType = 0
            else :
                data = data.split(multiSelectSplitSymbol)
                for d in data :
                    if d == self.stat.nullSymbol :
                        isNull = True
                currentType = 1
            
            if originType != -1 and isNull == False and originType != currentType :
                raise StatisticException(1, startCell.coordinate)
            elif originType == -1 and isNull == False :
                originType = currentType

            startCell = startCell.offset(moveDir[0], moveDir[1])

            if container :
                if isNull :
                    result.append(None)
                else :
                    result.append(data)
            else :
                if isNull :
                    continue
                result.extend(data)
        return result
