IS_WORKING = False
Cvar1 = 2
_window = None

def set_Cv(var: int):
    global Cvar1
    Cvar1 = var
    
def get_Cv():
    print(f"[get_Cv] 가져온 값: {Cvar1}")  # 디버깅용
    return Cvar1

def set_window(w):
    global _window
    _window = w
    
def get_window():
    if _window is None:
        raise RuntimeError("window가 설정되지 않음")
    return _window