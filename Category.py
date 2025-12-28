from openpyxl.cell.cell import Cell
from Stat import Stat

class Category() :
    def __init__(self, stat:Stat, cell:Cell) :
        from Legend import Legend

        self.stat: Stat = stat
        self.cell: Cell = cell
        self.name: str = str(cell.value)
        self.address: tuple[int, int] = (cell.row, cell.column)
        self.isComparison: bool = False
        self.legend: Legend = Legend(stat=stat, cat=self)
        self.unit:str = ""
        
        return
