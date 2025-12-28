import sys
import os
import openpyxl as Excel
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell
from Exception import StatisticException
import pandas as pd

class Stat() :
    def __init__(self, refer:str, name="무제1", tableGap=2, chartGap=2, table_chartGap=3, nullSymbol="-", catStartPoint="B3", multiSelectSplitSymbol=";") :
        self.referenceFile = refer
        self.statFileName = name
        self.table_gap = tableGap
        self.chart_gap = chartGap
        self.table_chart_gap = table_chartGap
        self.nullSymbol = nullSymbol
        self.catStartPoint = catStartPoint
        self.multiSelectSplitSymbol = multiSelectSplitSymbol
        self.referenceExcel:Workbook = None
        self.statFilePath: str = None
        self.statExcel: Workbook = None
        self.dataframe: pd.DataFrame = None
        self.fundamental_category:list[str] = list()
        self.correlation_category:list[str] = list()
        self.unit: dict[str] = dict()

        # 설문지 엑셀 파일 불러오
        if not os.path.exists(self.referenceFile) :
            detail = self.referenceFile
            self.referenceFile = None
            raise StatisticException(3, detail)
        
        # 설문지 파일 분석
        self.referenceExcel: Workbook = Excel.load_workbook(refer)
        surveySheetName: str = self.referenceExcel.sheetnames[0]
        surveySheet: Worksheet = self.referenceExcel[surveySheetName]
        categoryCell: Cell = surveySheet[catStartPoint]

        # 설문지 파일 -> pandas
        df = pd.read_excel(
            self.referenceFile,
            sheet_name=0,
            header=categoryCell.row-1,
            na_values=self.nullSymbol
        )
        self.dataframe = df.iloc[:, (categoryCell.column-1):].fillna("미응답")
        return
    
    def inflate(self) :
        from FundamentalSheet import FundamentalSheet
        from CorrelationSheet import CorrelationSheet

        ws = self.statExcel.active

        # 기본통계 시트 생성
        fundamental_sheet = FundamentalSheet(stat=self)

        # 상관분석 시트 생성
        for cat in self.correlation_category :
            correlation_sheet = CorrelationSheet(stat=self, category=cat)
        return
    
    def show(self) :
        self.statExcel = Workbook()

        # 기본 시트 삭제
        for sh in self.statExcel :
            self.statExcel.remove(sh)
        
        # 통계 분석
        self.inflate()
        
        # 통계 결과 저장
        try :
            self.statExcel.save(self.statFilePath)
        except PermissionError :
            raise StatisticException(0, self.statFilePath)
        finally :
            self.statExcel.close()
        
        # 생성 완료!
        return
    
    def clear(self) :
        self.fundamental_category:list[str] = list()
        self.correlation_category:list[str] = list()
        self.unit: dict[str] = dict()
        return