from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell
from Stat import Stat
from Category import Category
from Table import FundamentalTable, CorrelationTable

class Sheet() :
    def __init__(self, stat:Stat, name:str, comparisonCats:list) :
        self.stat: Stat = stat
        self.name: str = name
        self.chartStartPoint: str = "K2"
        self.tableStartPoint: str = "B2"
        self.table_gap: int = stat.table_gap
        self.chart_gap: int = stat.chart_gap
        self.table: list = list()
        self.chart: list = list()
        self.sheet: Worksheet = None
        self.comparisonCategories = comparisonCats
        return
    
    def inflate(self) :
        self.sheet: Worksheet = self.stat.statExcel.create_sheet(title=self.name)
        return

class FundamentalSheet(Sheet) :
    def __init__(self, stat:Stat, comparisonCats:list) :
        super().__init__(stat=stat, name="기본통계", comparisonCats=comparisonCats)

        for idx, cat in enumerate(self.comparisonCategories) :
            table: FundamentalTable = FundamentalTable(stat=stat, cat=cat, orderNum=idx)
            self.table.append(table)
        return
    
    def inflate(self):
        from Chart import FundamentalChart
        super().inflate()

        # 기본통계 시트에 테이블 생성
        maxCol = -1
        originCell:Cell = self.sheet[self.tableStartPoint]
        for table in self.table :
            table:FundamentalTable = table
            table.inflate(sheet=self.sheet, coor=originCell.coordinate)
            originCell = originCell.offset(table.height+self.table_gap, 0)

            if maxCol < table.width :
                maxCol = table.width

        # 기본통계 시트에 차트 생성
        originCell:Cell = self.sheet[self.chartStartPoint]
        for table in self.table :
            chart:FundamentalChart = FundamentalChart(stat=self.stat, table=table, location=originCell.coordinate)
            table.chart = chart
            self.chart.append(chart)

            chart.inflate()

            chart_gap = round(chart.chart.height / 0.6096) + self.chart_gap
            originCell = originCell.offset(chart_gap, 0)

        return


class CorrelationSheet(Sheet) :
    def __init__(self, stat:Stat, targetCat:Category, comparisonCats:list) :
        super().__init__(stat=stat, name=targetCat.name, comparisonCats=comparisonCats)
        self.targetCategory: Category = targetCat
        self.table_fundamental = None
        self.chart_fundamental = None
        
        table: FundamentalTable = FundamentalTable(stat=stat, cat=self.targetCategory, orderNum=0)
        self.table_fundamental = table

        for idx, comparisonCat in enumerate(self.comparisonCategories) :
            table: CorrelationTable = CorrelationTable(stat=stat, targetCat=self.targetCategory, comparisonCat=comparisonCat, orderNum=idx+1)
            self.table.append(table)
        return
    
    def inflate(self):
        from Chart import FundamentalChart
        from Chart import CorrelationChart
        super().inflate()
        
        maxCol = -1
        originCell:Cell = self.sheet[self.tableStartPoint]

        # 각 항목 시트에 기본 테이블 생성
        self.table_fundamental.inflate(sheet=self.sheet, coor=originCell.coordinate)
        originCell = originCell.offset(self.table_fundamental.height+self.table_gap, 0)
        maxCol = self.table_fundamental.width

        # 각 항목 시트에 테이블 생성
        for table in self.table :
            table:CorrelationTable = table
            table.inflate(sheet=self.sheet, coor=originCell.coordinate)
            originCell = originCell.offset(table.height+self.table_gap, 0)
            
            if maxCol < table.width :
                maxCol = table.width
        
        
        # 각 항목 시트에 기본 차트 생성
        originCell:Cell = self.sheet[self.chartStartPoint]
        chart_fundamental:FundamentalChart = FundamentalChart(stat=self.stat, table=self.table_fundamental, location=originCell.coordinate)
        self.table_fundamental.chart = chart_fundamental
        self.chart_fundamental = chart_fundamental

        chart_fundamental.inflate()
        
        chart_gap = round(chart_fundamental.chart.height / 0.6096) + self.chart_gap
        originCell = originCell.offset(chart_gap, 0)

        # 각 항목 시트에 차트 생성
        for table in self.table :
            chart:CorrelationChart = CorrelationChart(stat=self.stat, table=table, location=originCell.coordinate)
            table.chart = chart
            self.chart.append(chart)

            chart.inflate()

            chart_gap = round(chart.chart.height / 0.6096) + self.chart_gap
            originCell = originCell.offset(chart_gap, 0)

        return


class ChartOrganizeSheet(Sheet) :
    def __init__(self, stat:Stat, sheetNames:list) :
        self.sheetNames = sheetNames
        super().__init__(stat=stat, name="차트총괄", comparisonCats=[])
        return
    
    def inflate(self):
        from Chart import FundamentalChart, CorrelationChart
        super().inflate()

        heightCell:Cell = self.sheet[self.tableStartPoint]
        for name in self.sheetNames :
            sheet:Sheet = self.stat.sheets[name]
            widthCell:Cell = heightCell.offset(0, 0)

            if isinstance(sheet, FundamentalSheet) :
                # 기본통계 시트에 차트 생성
                chart_height_gap = 0
                for table in sheet.table :
                    fch:FundamentalChart = FundamentalChart(stat=self.stat, table=table, location=widthCell.coordinate)
                    fch.inflate(sh=self.sheet)
                    chart_width_gap = round(fch.chart.width / 1.8361) + self.chart_gap
                    chart_height_gap = round(fch.chart.height / 0.6096) + self.chart_gap
                    widthCell = widthCell.offset(0, chart_width_gap)
                heightCell = heightCell.offset(chart_height_gap, 0)
            elif isinstance(sheet, CorrelationSheet) :
                # 각 항목 시트에 기본 차트 생성
                chart_fundamental:FundamentalChart = FundamentalChart(stat=self.stat, table=sheet.table_fundamental, location=widthCell.coordinate)
                chart_fundamental.inflate(sh=self.sheet)
                chart_width_gap = round(chart_fundamental.chart.width / 1.8361) + self.chart_gap
                chart_height_gap = round(chart_fundamental.chart.height / 0.6096) + self.chart_gap
                widthCell = widthCell.offset(0, chart_width_gap)

                # 각 항목 시트에 차트 생성
                for table in sheet.table :
                    cch:CorrelationChart = CorrelationChart(stat=self.stat, table=table, location=widthCell.coordinate)
                    cch.inflate(sh=self.sheet)
                    chart_width_gap = round(cch.chart.width / 1.8361) + self.chart_gap
                    chart_height_gap = max(chart_height_gap, round(chart_fundamental.chart.height / 0.6096) + self.chart_gap)
                    widthCell = widthCell.offset(0, chart_width_gap)
                heightCell = heightCell.offset(chart_height_gap, 0)
            
        return