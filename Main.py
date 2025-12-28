import sys
import os
import Window


# ---- versions ----
# Windows 11
# Python 3.12.6
# pip 24.2

# PyQt5 5.15.11
# PyQt5-Qt5 5.15.2
# PyQt5_sip 12.15.0
# openpyxl 3.1.5
# Pyinstaller 6.11.1


# ---- build command ----
# pyinstaller -w -F -n="통계 산출 프로그램" --icon=statistic.ico Main.py --add-data=statistic.ico;./


# ---- developer ----
# https://github.com/Link-State


class Main() :
    def __init__(self) :
        return
    
    def main() :
        # pyqt5로 gui 만들기
        Window.run()
        return

def main() :
    Main.main()
    return

if __name__ == "__main__" :
    main()
