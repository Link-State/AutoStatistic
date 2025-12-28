
def isNumber(x:str) -> bool :
    if x.isdecimal() :
        return True
    
    try :
        y = float(x)
    except ValueError :
        return False
    
    return True