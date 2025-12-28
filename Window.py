import sys
import os

from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QIcon
from Action import *
from MouseEvent import *


def resource_path(relpath) :
    try :
        abspath = sys._MEIPASS
    except Exception :
        abspath = os.path.abspath(".")
    return os.path.join(abspath, relpath)

class MainWindow(QMainWindow) :
    def __init__(self) :
        super().__init__()
        self.initUI()
        self.statistic = None

    def initUI(self) :

        widget = QWidget()
        body = QVBoxLayout()

        # 설문지 파일 경로
        targetPathLayout = QHBoxLayout()
        targetPathLayout.setObjectName("targetPathLayout")

        qlb = QLabel("설문지 파일 위치 : ")
        targetPathLayout.addWidget(qlb)

        qle = ClickEditText(action=findExcelFile, this=True, self_=self)
        qle.setObjectName("targetPathEditText")
        targetPathLayout.addWidget(qle)

        body.addLayout(targetPathLayout)


        # 비교 대상 항목
        categoryLayout = QVBoxLayout()
        categoryLayout.addStretch(1)
        categoryLayout.setObjectName("categoryLayout")

        categoryTitleLayout = QHBoxLayout()
        title = QLabel("비교 대상 항목 선택")
        font = QFont("맑은고딕", weight=75)
        title.setFont(font)
        categoryTitleLayout.addWidget(title)

        categoryCheckboxLayout = QGridLayout()
        categoryCheckboxLayout.setObjectName("categoryCheckboxLayout")

        categoryLayout.addLayout(categoryTitleLayout)
        categoryLayout.addLayout(categoryCheckboxLayout)
        categoryLayout.addStretch(1)

        body.addLayout(categoryLayout)


        # 고급옵션
        advancedOption = QVBoxLayout()
        
        optionGroup = QGroupBox("고급 옵션")
        optionContainer = QGridLayout()

        # 고급옵션 - 표 좌우 간격
        gapGroup = QGroupBox("간격 설정")
        gapContainer = QVBoxLayout()

        tableGapContainer = QHBoxLayout()
        tableGapLabel = QLabel("테이블 간 상하 간격 : ")
        tableGap = QLineEdit("2")
        tableGap.setObjectName("tableGap")
        tableGapContainer.addWidget(tableGapLabel)
        tableGapContainer.addWidget(tableGap)
        
        chartGapContainer = QHBoxLayout()
        chartGapLabel = QLabel("차트 간 상하 간격 : ")
        chartGap = QLineEdit("2")
        chartGap.setObjectName("chartGap")
        chartGapContainer.addWidget(chartGapLabel)
        chartGapContainer.addWidget(chartGap)
        
        table_chartGapContainer = QHBoxLayout()
        table_chartGapLabel = QLabel("테이블-차트 간 간격 : ")
        table_chartGap = QLineEdit("3")
        table_chartGap.setObjectName("table_chartGap")
        table_chartGapContainer.addWidget(table_chartGapLabel)
        table_chartGapContainer.addWidget(table_chartGap)

        gapContainer.addLayout(tableGapContainer)
        gapContainer.addLayout(chartGapContainer)
        gapContainer.addLayout(table_chartGapContainer)
        gapGroup.setLayout(gapContainer)

        optionContainer.addWidget(gapGroup, 0, 0)

        # 고급옵션 - 시작 위치 설정
        startPointGroup = QGroupBox("설문지 검토 좌표 설정")
        startPointContainer = QVBoxLayout()

        ## 항목 시작 위치
        catStartPointContainer = QHBoxLayout()
        catStartPointlb = QLabel("항목 셀 검토 시작 좌표 : ")
        catStartPointEdtText = QLineEdit("B3")
        catStartPointEdtText.setObjectName("catStartPointEdtText")
        catStartPointContainer.addWidget(catStartPointlb)
        catStartPointContainer.addWidget(catStartPointEdtText)
        startPointContainer.addLayout(catStartPointContainer)

        startPointGroup.setLayout(startPointContainer)
        optionContainer.addWidget(startPointGroup, 0, 1)

        # 기호 설정
        symbolGroup = QGroupBox("기호 설정")
        symbolContainer = QVBoxLayout()

        ## 복수응답 구분자 기호 설정
        multiSelectedSplitSymbolContainer = QHBoxLayout()
        multiSelectedSplitSymbollb = QLabel("복수응답 구분 기호 : ")
        multiSelectedSplitSymbolEdtText = QLineEdit(";")
        multiSelectedSplitSymbolEdtText.setObjectName("multiSelectedSplitSymbolEdtText")
        multiSelectedSplitSymbolContainer.addWidget(multiSelectedSplitSymbollb)
        multiSelectedSplitSymbolContainer.addWidget(multiSelectedSplitSymbolEdtText)

        ## 결측치 기호 설정
        nullSymbolContainer = QHBoxLayout()
        nullSymbolLabel = QLabel("결측치 처리 기호 : ")
        nullSymbol = QLineEdit("-")
        nullSymbol.setObjectName("nullSymbol")
        nullSymbolContainer.addWidget(nullSymbolLabel)
        nullSymbolContainer.addWidget(nullSymbol)

        symbolContainer.addLayout(multiSelectedSplitSymbolContainer)
        symbolContainer.addLayout(nullSymbolContainer)
        symbolGroup.setLayout(symbolContainer)
        optionContainer.addWidget(symbolGroup, 0, 2)


        optionGroup.setLayout(optionContainer)
        advancedOption.addWidget(optionGroup)

        body.addLayout(advancedOption)
        

        # 통계 파일 저장 경로
        savePathLayout = QHBoxLayout()
        savePathLayout.setObjectName("savePathLayout")

        qlb = QLabel("저장 위치 : ")
        savePathLayout.addWidget(qlb)

        qle = ClickEditText(action=findSavePath, this=True, self_=self)
        qle.setObjectName("savePathEditText")
        qle.setText(os.getcwd() + "\\무제1_결과.xlsx")
        savePathLayout.addWidget(qle)

        body.addLayout(savePathLayout)


        # 통계 생성 버튼
        generateBtnLayout = QHBoxLayout()

        generateBtnLayout.addStretch(1)

        surveyBtn = ClickButton(text="설문지 양식", action=createSurveyForm, this=True, self_=self)
        # helpBtn = ClickButton(text="도움말", action=helpWindow, this=True, self_=self)
        qpb = ClickButton(text="통계 산출", action=generateStatistic, this=True, self_=self)
        generateBtnLayout.addWidget(surveyBtn)
        # generateBtnLayout.addWidget(helpBtn)
        generateBtnLayout.addWidget(qpb)

        body.addLayout(generateBtnLayout)

        # 메인 윈도우 설정
        widget.setLayout(body)
        self.setCentralWidget(widget)

        # 타이틀바 설정
        self.setWindowTitle("통계 산출 프로그램 v2.0.0")
        # 아이콘 설정
        self.setWindowIcon(QIcon(resource_path('./statistic.ico')))
        # 초기 위치 설정
        self.move(50, 50)
        # 창 초기 크기 설정
        self.resize(800, 600)

        
        self.show()


def run() :
    app = QApplication(sys.argv)
    ex = MainWindow()
    sys.exit(app.exec_())
