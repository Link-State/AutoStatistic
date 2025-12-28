
from PyQt5.QtWidgets import *

class ClickEditText(QLineEdit) :
    def __init__(self, action=None, **actionParam) :
        super().__init__()
        self.init()
        self.action = action
        self.actionParam = actionParam
    
    def init(self) :
        QLineEdit()
        return
    
    def mouseReleaseEvent(self, event) :
        # action
        if self.action != None :
            if "this" in self.actionParam :
                self.actionParam["this"] = self
            if "event" in self.actionParam :
                self.actionParam["event"] = event
            
            if len(self.actionParam) > 0 :
                return self.action(self.actionParam)
            else :
                return self.action()
        return


class ClickButton(QPushButton) :
    def __init__(self, text="", action=None, **actionParam) :
        super().__init__(text=text)
        self.init()
        self.action = action
        self.actionParam = actionParam
    
    def init(self) :
        QPushButton()
        return
    
    def mouseReleaseEvent(self, event) :
        self.clearFocus()

        # action
        if self.action != None :
            if "this" in self.actionParam :
                self.actionParam["this"] = self
            if "event" in self.actionParam :
                self.actionParam["event"] = event
            
            if len(self.actionParam) > 0 :
                return self.action(self.actionParam)
            else :
                return self.action()
        return
