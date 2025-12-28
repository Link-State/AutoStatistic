
class Code() :
    FILE_PERMISSION = 0
    TYPE_CONFLICT = 1
    DIFFERENT_DATA_LENGTH = 2
    NOT_EXIST_REFERENCE_FILE = 3
    NO_DATA = 4

class StatisticException(Exception) :
    def __init__(self, errcode:int, detail="") :

        err_msg = ""
        if errcode == Code.FILE_PERMISSION :
            err_msg = f"'{detail}' 파일이 열려있습니다."
        elif errcode == Code.TYPE_CONFLICT :
            err_msg = f"셀 {detail}에서 데이터 타입이 변경됩니다."
        elif errcode == Code.DIFFERENT_DATA_LENGTH :
            err_msg = f"({detail}) 두 열이 갖는 데이터의 갯수가 다릅니다."
        elif errcode == Code.NOT_EXIST_REFERENCE_FILE :
            err_msg = f"'{detail}' 파일을 찾을 수 없습니다."
        elif errcode == Code.NO_DATA :
            err_msg = f"'{detail}' 열에 데이터가 존재하지 않습니다."
        else :
            err_msg = "알 수 없음"

        super().__init__(err_msg)

        return
