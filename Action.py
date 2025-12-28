import sys
import os
import traceback
from Exception import StatisticException
from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt
from Util import *


# 엑셀파일 찾으려고 할 때
def findExcelFile(kwargs) :
    from Window import MainWindow
    from MouseEvent import ClickEditText
    from Stat import Stat

    self_: MainWindow = kwargs["self_"]
    this: ClickEditText = kwargs["this"]

    try :
        # 파일 열기
        refDir = QFileDialog.getOpenFileName(self_, '파일 열기', '', 'Excel File(*.xlsx);; Macro Excel File(*.xlsm);; 모든 파일(*.*)')
        refDir = refDir[0]
        if refDir == "" :
            return
        
        refDir = refDir.replace("/", "\\")
        this.setText(refDir)

        # 검토 좌표 설정 불러오기
        catStartPointEdtText: QLineEdit = self_.findChild(QLineEdit, name="catStartPointEdtText")
        catStartPoint = catStartPointEdtText.text().strip()

        # 항목 객체 불러오기
        refName = os.path.basename(refDir).strip()
        refName = os.path.splitext(refName)
        statistic: Stat = Stat(refer=refDir, name=refName[0]+"_결과", catStartPoint=catStartPoint)

        if len(statistic.dataframe.keys()) <= 0 :
            QMessageBox.critical(self_, '통계 산출 프로그램 알림', f'설문지에서 항목을 찾을 수 없습니다.')
            return
        
        # 기존 항목 제거
        categoryItem: QGridLayout = self_.findChild(QGridLayout, name="categoryCheckboxLayout").layout()
        clearLayout(categoryItem)
        
        # 항목 객체로 체크박스 만들기
        for i, key in enumerate(statistic.dataframe.keys()) :
            
            categoryBox = QHBoxLayout()

            checkBox = QCheckBox(str(key))

            unitEdtText = QLineEdit()
            unitEdtText.setFixedWidth(100)
            unitEdtText.setObjectName("unitEdtText")
            unitEdtText.setPlaceholderText("단위")

            categoryBox.addWidget(checkBox)
            categoryBox.addWidget(unitEdtText)
            categoryBox.addStretch(1)
            categoryItem.addLayout(categoryBox, i//2, i%2)

        savePathEdtText: ClickEditText = self_.findChild(ClickEditText, name="savePathEditText")

        dirname = os.path.dirname(refDir).replace("/", "\\")
        savePathEdtText.setText(dirname + "\\" + refName[0] + "_결과.xlsx")

        # 통계 객체 저장
        self_.statistic = statistic
    
    except Exception as e :
        exc_type, exc_obj, exc_tb = sys.exc_info()
        err_fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        err_line = exc_tb.tb_lineno
        if type(e) == type(StatisticException(-1)) :
            QMessageBox.critical(self_, '통계 산출 프로그램 알림', f"[{err_fname} : {err_line}]\n{e}")
        else :
            QMessageBox.critical(self_, '통계 산출 프로그램 알림', f"[{err_fname} : {err_line}]\n{traceback.format_exc()}")
    return


# 저장 위치 선택
def findSavePath(kwargs) :
    from Window import MainWindow
    from MouseEvent import ClickEditText
    from Stat import Stat

    self_: MainWindow = kwargs["self_"]
    this: ClickEditText = kwargs["this"]
    statistic: Stat = self_.statistic
    
    try :
        # 저장 폴더 선택
        folderPath: str = QFileDialog.getExistingDirectory(self_, '저장할 폴더 선택', '')
        if folderPath == "" :
            return
        
        # 파일 이름 결정
        fName = "무제1_결과"
        if statistic != None :
            fName = statistic.statFileName
        
        savePath = folderPath + "\\" + fName + ".xlsx"
        savePath = savePath.replace("/", "\\")
        this.setText(savePath)
    
    except Exception as e :
        exc_type, exc_obj, exc_tb = sys.exc_info()
        err_fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        err_line = exc_tb.tb_lineno
        if type(e) == type(StatisticException(-1)) :
            QMessageBox.critical(self_, '통계 산출 프로그램 알림', f"[{err_fname} : {err_line}]\n{e}")
        else :
            QMessageBox.critical(self_, '통계 산출 프로그램 알림', f"[{err_fname} : {err_line}]\n{traceback.format_exc()}")

    return


# 통계산출 버튼 클릭
def generateStatistic(kwargs) :
    from Window import MainWindow
    from MouseEvent import ClickEditText
    from Stat import Stat

    self_: MainWindow = kwargs["self_"]
    this: ClickEditText = kwargs["this"]
    statistic: Stat = self_.statistic

    statistic.clear()

    try :
        # 통계 객체 생성 체크
        if statistic == None :
            QMessageBox.critical(self_, '통계 산출 프로그램 알림', f'설문지 엑셀 파일이 지정되지 않았습니다.')
            return

        # 저장 경로 가져오기
        savePathEdtText: ClickEditText = self_.findChild(ClickEditText, name="savePathEditText")
        savePath: str = savePathEdtText.text().strip()
        if savePath == "" :
            QMessageBox.critical(self_, '통계 산출 프로그램 알림', f'통계 결과 파일 저장 경로가 지정되지 않았습니다.')
            return
        
        # 저장 경로 존재 체크
        statistic.statFilePath = savePath
        
        # 저장 파일 이름
        fileName = os.path.basename(savePath).strip()
        if fileName == "" :
            QMessageBox.critical(self_, '통계 산출 프로그램 알림', f'통계 결과 파일 이름이 지정되지 않았습니다.')
            return
        statistic.statFileName = fileName
        
        # 동일 폴더 내 저장 파일 이름 중복 체크
        if os.path.exists(savePath) :
            reply = QMessageBox.question(self_, '통계 산출 프로그램 알림', f"'{fileName}' 동일한 이름의 파일이 존재합니다. 덮어 씌우시겠습니까?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            
            if reply == QMessageBox.No :
                return

        # 선택한 비교 대상 항목 처리
        categoryItem: QGridLayout = self_.findChild(QGridLayout, name="categoryCheckboxLayout").layout()
        for i in range(categoryItem.count()) :
            boxContainer: QHBoxLayout = categoryItem.itemAt(i).layout()
            item:QCheckBox = boxContainer.itemAt(0).widget()
            unit:QLineEdit = boxContainer.itemAt(1).widget()

            if item.text() in statistic.dataframe.keys() :
                statistic.unit[item.text()] = unit.text()
                if item.isChecked() :
                    statistic.fundamental_category.append(item.text())
                else :
                    statistic.correlation_category.append(item.text())
        
        # 고급 옵션 불러와서 적용하기
        tableGapEdtText: QLineEdit = self_.findChild(QLineEdit, name="tableGap")
        chartGapEdtText: QLineEdit = self_.findChild(QLineEdit, name="chartGap")
        table_chartGapEdtText: QLineEdit = self_.findChild(QLineEdit, name="table_chartGap")
        nullSymbolEdtText: QLineEdit = self_.findChild(QLineEdit, name="nullSymbol")
        catStartPointEdtText: QLineEdit = self_.findChild(QLineEdit, name="catStartPointEdtText")
        multiSelectedSplitSymbolEdtText: QLineEdit = self_.findChild(QLineEdit, name="multiSelectedSplitSymbolEdtText")

        tableGap = tableGapEdtText.text().strip()
        chartGap = chartGapEdtText.text().strip()
        table_chartGap = table_chartGapEdtText.text().strip()
        nullSymbol = nullSymbolEdtText.text().strip()
        catStartPoint = catStartPointEdtText.text().strip()
        multiSelectedSplitSymbol = multiSelectedSplitSymbolEdtText.text().strip()

        if tableGap.isdigit() :
            statistic.table_gap = int(tableGap)
        if chartGap.isdigit() :
            statistic.chart_gap = int(chartGap)
        if table_chartGap.isdigit() :
            statistic.table_chart_gap = int(table_chartGap)
        statistic.nullSymbol = nullSymbol
        statistic.catStartPoint = catStartPoint
        statistic.multiSelectSplitSymbol = multiSelectedSplitSymbol
        
        # 시트 생성
        statistic.show()

        # 성공 창 띄우기
        QMessageBox.about(self_, '통계 산출 프로그램 알림', f'통계 결과가 {statistic.statFilePath}에 저장되었습니다.')
    
    except Exception as e :
        # https://wikidocs.net/21935
        # https://mr-doosun.tistory.com/29

        exc_type, exc_obj, exc_tb = sys.exc_info()
        err_fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        err_line = exc_tb.tb_lineno
        if type(e) == type(StatisticException(-1)) :
            QMessageBox.critical(self_, '통계 산출 프로그램 알림', f"[{err_fname} : {err_line}]\n{e}")
        else :
            QMessageBox.critical(self_, '통계 산출 프로그램 알림', f"[{err_fname} : {err_line}]\n{traceback.format_exc()}")

    return

def helpWindow(kwargs) :
    from Window import MainWindow
    from MouseEvent import ClickEditText

    self_: MainWindow = kwargs["self_"]
    this: ClickEditText = kwargs["this"]

    try :
        pass
    except Exception as e :
        exc_type, exc_obj, exc_tb = sys.exc_info()
        err_fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        err_line = exc_tb.tb_lineno
        QMessageBox.critical(self_, '통계 산출 프로그램 알림', f"[{err_fname} : {err_line}]\n{e}")

    return

def createSurveyForm(kwargs) : 
    from Window import MainWindow
    from MouseEvent import ClickEditText
    from Survey import Survey

    self_: MainWindow = kwargs["self_"]
    this: ClickEditText = kwargs["this"]

    try :
        newSurvey = Survey()
        newSurvey.create()
        QMessageBox.about(self_, '통계 산출 프로그램 알림', f"'{os.path.abspath(newSurvey.fileDir)}'에 설문지 파일이 저장되었습니다.")
    except Exception as e :
        exc_type, exc_obj, exc_tb = sys.exc_info()
        err_fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        err_line = exc_tb.tb_lineno
        if type(e) == type(StatisticException(-1)) :
            QMessageBox.critical(self_, '통계 산출 프로그램 알림', f"[{err_fname} : {err_line}]\n{e}")
        else :
            QMessageBox.critical(self_, '통계 산출 프로그램 알림', f"[{err_fname} : {err_line}]\n{traceback.format_exc()}")
    
    return
