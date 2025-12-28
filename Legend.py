
from openpyxl.cell.cell import Cell
from Stat import Stat
from Category import Category
from Data import Data

class Legend() :
    def __init__(self, stat:Stat, cat:Category) :
        self.stat: Stat = stat
        self.category: Category = cat
        self.legend: dict = dict()
        self.legendType:type = None

        row, col = self.stat.moveDir
        dataCell: Cell = self.category.cell
        dataCell = dataCell.offset(col, row)

        self.data: Data = Data(stat=self.stat, cell=dataCell, moveDir=(col, row))
        datas: list = self.data.readLine(container=False)
        self.legendType = type(datas[0])

        keys = set(datas)
        legend = dict()
        for key in keys :
            legend[key] = datas.count(key)
        self.legend = legend

        return