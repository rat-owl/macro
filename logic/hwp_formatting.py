# logic/hwp_formatting.py (1/3)

from pyhwpx import Hwp
from core.com_utils import hwp_opened
from ui.ui_main import MainWindow
import time
from core.global_state import get_window
from core.global_state import get_Cv


def 표문단설정():
    hwp=Hwp()
    
    cv = get_Cv()
    print(cv)
    ctrl = hwp.HeadCtrl
    
    n=0

    while ctrl:
        if ctrl.UserDesc == "표":
            if ctrl.Next is None:
                if cv == 2:
                    hwp.get_into_nth_table(n)
                    hwp.TableCellBlock()
                    hwp.TableLowerCell()
                    hwp.TableCellBlockExtend()
                    hwp.TableRightCell()
                    hwp.TableColPageDown()
                    hwp.TableUpperCell()
                    hwp.TableUpperCell()
                    hwp.TableUpperCell()
                    hwp.TableUpperCell()
                    hwp.TableUpperCell()
                
                    pset=hwp.HParameterSet.HParaShape
            
                    hwp.HAction.GetDefault("ParagraphShape", pset.HSet)
                    pset.LeftMargin = hwp.PointToHwpUnit(10.0)
                    hwp.HAction.Execute("ParagraphShape", pset.HSet)
                    n +=1
                else:
                    print("→ 체크 안됨: cv != 2 처리 실행")
                    hwp.get_into_nth_table(n)
                    hwp.TableCellBlock()
                    hwp.TableLowerCell()
                    hwp.TableRightCell()
                    hwp.TableCellBlockExtend()
                    hwp.TableColPageDown()
                    hwp.TableUpperCell()
                    
                    pset=hwp.HParameterSet.HParaShape
                    
                    hwp.HAction.GetDefault("ParagraphShape", pset.HSet)
                    pset.LeftMargin = hwp.PointToHwpUnit(30.0)
                    hwp.HAction.Execute("ParagraphShape", pset.HSet)
                    
                    hwp.get_into_nth_table(n)
                    hwp.TableCellBlock()
                    hwp.TableRightCell()
                    hwp.TableRightCell()
                    hwp.TableLowerCell()
                    hwp.TableCellBlockExtend()
                    hwp.TableColPageDown()
                    hwp.TableUpperCell()
                    
                    hwp.HAction.GetDefault("ParagraphShape", pset.HSet)
                    pset.LeftMargin = hwp.PointToHwpUnit(20.0)
                    hwp.HAction.Execute("ParagraphShape", pset.HSet)
                    
                    n +=1
            else:
                hwp.get_into_nth_table(n)
                hwp.TableCellBlock()
                hwp.TableLowerCell()
                hwp.TableRightCell()
                hwp.TableCellBlockExtend()
                hwp.TableColPageDown()
                hwp.TableUpperCell()
                
                pset=hwp.HParameterSet.HParaShape
                
                hwp.HAction.GetDefault("ParagraphShape", pset.HSet)
                pset.LeftMargin = hwp.PointToHwpUnit(30.0)
                hwp.HAction.Execute("ParagraphShape", pset.HSet)
                
                hwp.get_into_nth_table(n)
                hwp.TableCellBlock()
                hwp.TableRightCell()
                hwp.TableRightCell()
                hwp.TableLowerCell()
                hwp.TableCellBlockExtend()
                hwp.TableColPageDown()
                hwp.TableUpperCell()
                
                hwp.HAction.GetDefault("ParagraphShape", pset.HSet)
                pset.LeftMargin = hwp.PointToHwpUnit(20.0)
                hwp.HAction.Execute("ParagraphShape", pset.HSet)
                
                n +=1
                  
        ctrl = ctrl.Next
    
    pset = hwp.HParameterSet.HFindReplace
    hwp.HAction.GetDefault("RepeatFind", pset.HSet)
    pset.FindString = "출현종수"
    pset.Direction = 1
    pset.IgnoreMessage = 1
    hwp.HAction.Execute("RepeatFind", pset.HSet)
    
    start = hwp.GetPos()
    
    while True:
        pset = hwp.HParameterSet.HFindReplace
        hwp.HAction.GetDefault("RepeatFind", pset.HSet)
        pset.FindString = "출현종수"
        pset.Direction = 1
        pset.IgnoreMessage = 1
        hwp.HAction.Execute("RepeatFind", pset.HSet)
        
        now = hwp.GetPos()
        
        hwp.TableCellBlock()
        pset=hwp.HParameterSet.HParaShape
        hwp.HAction.GetDefault("ParagraphShape", pset.HSet)
        pset.LeftMargin = hwp.PointToHwpUnit(0.0)
        hwp.HAction.Execute("ParagraphShape", pset.HSet)
        
        if now == start:
            break

# 한글에서 표를 선택하여 기본 틀을 잡을 것(표 속성)
def 표기본():
    cv = get_Cv()  # CVar1의 값을 cv에 할당
    # cv를 사용하여 작업 수행

    hwp=hwp_opened()
    hwp=Hwp()
    
    # 페이지 설정 조정
    pset = hwp.HParameterSet.HSecDef
    hwp.HAction.GetDefault("PageSetup", pset.HSet)
    # 페이지 여백 및 방향 설정
    pset.PageDef.LeftMargin = hwp.MiliToHwpUnit(5.0)  # 왼쪽 여백 5mm
    pset.PageDef.RightMargin = hwp.MiliToHwpUnit(5.0)  # 오른쪽 여백 5mm
    pset.PageDef.TopMargin = hwp.MiliToHwpUnit(5.0)  # 위쪽 여백 5mm
    pset.PageDef.BottomMargin = hwp.MiliToHwpUnit(5.0)  # 아래쪽 여백 5mm
    pset.PageDef.HeaderLen = hwp.MiliToHwpUnit(5.0)  # 머리글 길이 5mm
    pset.PageDef.FooterLen = hwp.MiliToHwpUnit(5.0)  # 바닥글 길이 5mm
    pset.PageDef.Landscape = 1  # 가로 방향 설정
    # 적용 범위 설정
    pset.HSet.SetItem("ApplyClass", 24)
    pset.HSet.SetItem("ApplyTo", 3)
    # 페이지 설정 실행
    hwp.HAction.Execute("PageSetup", pset.HSet)
    
    #클래스 윗테두리
    for _ in range(2):
        pset = hwp.HParameterSet.HFindReplace
        hwp.HAction.GetDefault("RepeatFind", pset.HSet)
        pset.FindString = "Class"  # 찾을 문자열
        pset.Direction = 0  # 찾기 방향: 앞으로
        pset.IgnoreMessage = 1
        hwp.HAction.Execute("RepeatFind", pset.HSet)
        
        hwp.TableCellBlock()
        hwp.TableCellBlockExtend()
        hwp.TableColEnd()
        
        pset = hwp.HParameterSet.HCellBorderFill
        hwp.HAction.GetDefault("CellBorderFill", pset.HSet)
        pset.BorderWidthTop = hwp.HwpLineWidth("0.12mm")
        pset.BorderTypeTop = hwp.HwpLineType("Solid")
        hwp.HAction.Execute("CellBorderFill", pset.HSet)
        
        

    ctrl = hwp.HeadCtrl
    cv = get_Cv()
    n=0

    while ctrl:
        if ctrl.UserDesc == "표":
            hwp.get_into_nth_table(n)
            hwp.TableCellBlock()
            hwp.TableRightCell()
            
            if ctrl.Next is None:
                if cv == 2:
                    hwp.TableLeftCell()
                else:
                    print(f"▶ 마지막 표 + 체크안됨(cv={cv}): 스킵")
            
            hwp.TableRightCell()
            hwp.TableRightCell()
            hwp.TableCellBlockExtend()
            hwp.TableColEnd()
            hwp.TableColPageDown()
            
            pset = hwp.HParameterSet.HShapeObject
            hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
            pset.HSet.SetItem("ShapeType", 3)  # 표 모양 유형 설정
            pset.HSet.SetItem("ShapeCellSize", 1)  # 셀 크기 설정
            pset.ShapeTableCell.Width = hwp.MiliToHwpUnit(10.0)
            hwp.HAction.Execute("TablePropertyDialog", pset.HSet)

            hwp.get_into_nth_table(n)
            hwp.TableCellBlock()
            hwp.TableCellBlockExtend()
            hwp.TableColEnd()
            
            # 줄간격 설정
            pset = hwp.HParameterSet.HParaShape
            hwp.HAction.GetDefault("ParagraphShape", pset.HSet)
            line_spacing = get_window().line.value()
            pset.LineSpacing = line_spacing
            hwp.HAction.Execute("ParagraphShape", pset.HSet)
            
            hwp.TableColPageDown()
            
            # 문자 스타일 변경 - UI에서 폰트와 크기 가져오기
            pset = hwp.HParameterSet.HCharShape
            hwp.HAction.GetDefault("CharShape", pset.HSet)
            # UI에서 선택된 폰트와 크기 가져오기
            selected_font = get_window().get_selected_font()
            font_size = get_window().get_font_size()
            font_size = get_window().font_size.value()
            
            # 폰트 설정
            pset.FaceNameHangul = selected_font
            pset.FaceNameLatin = selected_font
            pset.FaceNameHanja = selected_font
            pset.FaceNameJapanese = selected_font
            pset.FaceNameOther = selected_font
            pset.FaceNameSymbol = selected_font
            pset.FaceNameUser = selected_font
            
            pset.FontTypeHangul = 1
            pset.FontTypeLatin = 1
            pset.FontTypeHanja = 1
            pset.FontTypeJapanese = 1
            pset.FontTypeOther = 1
            pset.FontTypeSymbol = 1
            pset.FontTypeUser = 1
            
            # 크기 설정
            pset.Height = hwp.PointToHwpUnit(font_size)
            
            hwp.HAction.Execute("CharShape", pset.HSet)
            
            #표속성 변경
            pset = hwp.HParameterSet.HShapeObject
            hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
            pset.CellMarginBottom = hwp.MiliToHwpUnit(0.5)
            pset.CellMarginTop = hwp.MiliToHwpUnit(0.5)
            pset.CellMarginRight = hwp.MiliToHwpUnit(0.5)
            pset.CellMarginLeft = hwp.MiliToHwpUnit(0.5)
            pset.OutsideMarginBottom = hwp.MiliToHwpUnit(0.5)
            pset.OutsideMarginTop = hwp.MiliToHwpUnit(0.5)
            pset.OutsideMarginRight = hwp.MiliToHwpUnit(0.5)
            pset.OutsideMarginLeft = hwp.MiliToHwpUnit(0.5)

            pset.ShapeTableCell.VertAlign = hwp.VAlign("Center")
            pset.ShapeTableCell.LineWrap = 1
            
            # 표의 셀 높이 조절
            pset.HSet.SetItem("ShapeType", 3)  # 표 모양 유형 설정
            pset.HSet.SetItem("ShapeCellSize", 1)  # 셀 크기 설정
            pset.ShapeTableCell.Height = hwp.MiliToHwpUnit(4.0)  # 높이를 4mm로 설정
            
            hwp.HAction.Execute("TablePropertyDialog", pset.HSet)


            
            if ctrl.Next is None:
                if cv == 2:
                    pset = hwp.HParameterSet.HCellBorderFill
                
                    #바깥 테두리
                    hwp.HAction.GetDefault("CellBorderFill", pset.HSet)
                    pset.BorderWidthBottom = hwp.HwpLineWidth("0.4mm")
                    pset.BorderTypeBottom = hwp.HwpLineType("Solid")
                    pset.BorderWidthTop = hwp.HwpLineWidth("0.4mm")
                    pset.BorderTypeTop = hwp.HwpLineType("Solid")
                    pset.BorderWidthRight = hwp.HwpLineWidth("0.4mm")
                    pset.BorderTypeRight = hwp.HwpLineType("Solid")
                    pset.BorderWidthLeft = hwp.HwpLineWidth("0.4mm")
                    pset.BorderTypeLeft = hwp.HwpLineType("Solid")
                    pset.TypeVert = hwp.HwpLineType("Solid")
                    pset.WidthVert = hwp.HwpLineWidth("0.12mm")
                    hwp.HAction.Execute("CellBorderFill", pset.HSet)
                
                    #아래쪽 분석값 실선
                    for _ in range(5):
                        hwp.TableUpperCell()
                        hwp.HAction.GetDefault("CellBorderFill", pset.HSet)
                        pset.BorderWidthBottom = hwp.HwpLineWidth("0.12mm")
                        pset.BorderTypeBottom = hwp.HwpLineType("Solid")
                        hwp.HAction.Execute("CellBorderFill", pset.HSet)
                
                    #수평선 점선
                    hwp.HAction.GetDefault("CellBorderFill", pset.HSet)
                    pset.TypeHorz = hwp.HwpLineType("Dot")
                    pset.WidthHorz = hwp.HwpLineWidth("0.12mm")
                    hwp.HAction.Execute("CellBorderFill", pset.HSet)
                
                    #헤드 실선
                    hwp.TableColPageUp()
                    hwp.HAction.GetDefault("CellBorderFill", pset.HSet)
                    pset.BorderWidthBottom = hwp.HwpLineWidth("0.5mm")
                    pset.BorderTypeBottom = hwp.HwpLineType("DoubleSlim")
                    hwp.HAction.Execute("CellBorderFill", pset.HSet)
                else:
                    print(f"▶ 마지막 표 + 체크안됨(cv={cv}): 스킵")
                
            pset = hwp.HParameterSet.HCellBorderFill
            hwp.HAction.GetDefault("CellBorderFill", pset.HSet)
            pset.BorderWidthBottom = hwp.HwpLineWidth("0.4mm")
            pset.BorderTypeBottom = hwp.HwpLineType("Solid")
            pset.BorderWidthTop = hwp.HwpLineWidth("0.4mm")
            pset.BorderTypeTop = hwp.HwpLineType("Solid")
            pset.BorderWidthRight = hwp.HwpLineWidth("0.4mm")
            pset.BorderTypeRight = hwp.HwpLineType("Solid")
            pset.BorderWidthLeft = hwp.HwpLineWidth("0.4mm")
            pset.BorderTypeLeft = hwp.HwpLineType("Solid")
            pset.TypeVert = hwp.HwpLineType("Solid")
            pset.WidthVert = hwp.HwpLineWidth("0.12mm")
            hwp.HAction.Execute("CellBorderFill", pset.HSet)

            #헤드 실선
            hwp.TableColPageUp()
            hwp.HAction.GetDefault("CellBorderFill", pset.HSet)
            pset.BorderWidthBottom = hwp.HwpLineWidth("0.5mm")
            pset.BorderTypeBottom = hwp.HwpLineType("DoubleSlim")
            hwp.HAction.Execute("CellBorderFill", pset.HSet)

            n += 1
        ctrl = ctrl.Next

#학명스타일
def 학명스타일(find_name, style_name):
    hwp=Hwp()
    pset = hwp.HParameterSet.HFindReplace

    hwp.HAction.GetDefault("AllReplace", pset.HSet)
    pset.FindString = find_name  # 찾을 문자열
    pset.ReplaceString = find_name
    pset.ReplaceMode = 1
    pset.Direction = 0  # 찾기 방향: 앞으로
    pset.IgnoreMessage = 1
    pset.ReplaceStyle = style_name
    #pset.ReplaceParaShape.LeftMargin = hwp.PointToHwpUnit(0.0)
    hwp.HAction.Execute("AllReplace", pset.HSet)
    
    hwp.HAction.GetDefault("AllReplace", pset.HSet)
    pset.FindString = find_name  # 찾을 문자열
    pset.ReplaceString = find_name
    pset.ReplaceMode = 1
    pset.Direction = 0  # 찾기 방향: 앞으로
    pset.IgnoreMessage = 1
    pset.ReplaceStyle = style_name
    #pset.ReplaceParaShape.LeftMargin = hwp.PointToHwpUnit(0.0)
    return hwp.HAction.Execute("AllReplace", pset.HSet)

#국명스타일
def 국명():
    hwp=Hwp()
    ctrl = hwp.HeadCtrl
    n=0

    cv = get_Cv()
    
    while ctrl:
        if ctrl.UserDesc == "표":
            if ctrl.Next is None:
                if cv == 2:
                    break
                else:
                    hwp.get_into_nth_table(n)
                    hwp.TableCellBlock()
                    hwp.TableColEnd()
                    hwp.TableLowerCell()
                    hwp.TableCellBlockExtend()
                    hwp.TableColPageDown()
                    hwp.TableUpperCell()

                    pset=hwp.HParameterSet.HSort
                    hwp.HAction.GetDefault("Sort", pset.HSet)
                    pset.CreateItemArray("KeyOption", 2)
                    pset.KeyOption.SetItem(0,1)
                    pset.KeyOption.SetItem(1,2)
                    hwp.HAction.Execute("Sort", pset.HSet)

                    def 스타일뿌셔(hwp, style_name):
                            hwp.get_into_nth_table(n)
                            pset = hwp.HParameterSet.HFindReplace
                            hwp.HAction.GetDefault("RepeatFind", pset.HSet)
                            pset.FindString = "family"  # 찾을 문자열
                            pset.Direction = 0  # 찾기 방향: 앞으로
                            pset.IgnoreMessage = 1
                            hwp.HAction.Execute("RepeatFind", pset.HSet)
        
                            if type(style_name) != int:
                                print(f"[DEBUG] style_name 입력값: {style_name}")
                                style_dict = hwp.get_style_dict(as_="dict")
                                print(f"[DEBUG] 반환된 style_dict: {style_dict}")

                                found = False  # 찾았는지 여부 체크

                                for key, value in style_dict.items():
                                    print(f"[DEBUG] key={key}, value={value}")

                                    style_label = value.get('Name')
                                    print(f"[DEBUG] 비교 중: '{style_label}' vs '{style_name}'")

                                    if style_label == style_name:
                                        print(f"[INFO] 스타일 '{style_name}' 적용 준비 완료 (key={key})")
                                        style_name = key
                                        found = True
                                        break

                                if not found:
                                    print(f"[ERROR] 스타일 '{style_name}'을 찾을 수 없습니다.")
                                    raise KeyError(f"해당하는 스타일이 없습니다: {style_name}")
        
                            hwp.TableCellBlock()
                            hwp.TableRightCell()
                            hwp.TableCellBlockExtend()
                            hwp.TableColPageDown()
                            hwp.TableUpperCell()
        
                            pset = hwp.HParameterSet.HStyle
                            hwp.HAction.GetDefault("StyleEx", pset.HSet)
                            pset.Apply = style_name
                            hwp.HAction.Execute("StyleEx", pset.HSet)

            else:
                hwp.get_into_nth_table(n)
                hwp.TableCellBlock()
                hwp.TableColEnd()
                hwp.TableLowerCell()
                hwp.TableCellBlockExtend()
                hwp.TableColPageDown()
                hwp.TableUpperCell()

                pset=hwp.HParameterSet.HSort
                hwp.HAction.GetDefault("Sort", pset.HSet)
                pset.CreateItemArray("KeyOption", 2)
                pset.KeyOption.SetItem(0,1)
                pset.KeyOption.SetItem(1,2)
                hwp.HAction.Execute("Sort", pset.HSet)

                def 스타일뿌셔(hwp, style_name):
                        hwp.get_into_nth_table(n)
                        pset = hwp.HParameterSet.HFindReplace
                        hwp.HAction.GetDefault("RepeatFind", pset.HSet)
                        pset.FindString = "family"  # 찾을 문자열
                        pset.Direction = 0  # 찾기 방향: 앞으로
                        pset.IgnoreMessage = 1
                        hwp.HAction.Execute("RepeatFind", pset.HSet)
    
                        if type(style_name) != int:
                            print(f"[DEBUG] style_name 입력값: {style_name}")
                            style_dict = hwp.get_style_dict(as_="dict")
                            print(f"[DEBUG] 반환된 style_dict: {style_dict}")

                            found = False  # 찾았는지 여부 체크

                            for key, value in style_dict.items():
                                print(f"[DEBUG] key={key}, value={value}")

                                style_label = value.get('Name')
                                print(f"[DEBUG] 비교 중: '{style_label}' vs '{style_name}'")

                                if style_label == style_name:
                                    print(f"[INFO] 스타일 '{style_name}' 적용 준비 완료 (key={key})")
                                    style_name = key
                                    found = True
                                    break

                            if not found:
                                print(f"[ERROR] 스타일 '{style_name}'을 찾을 수 없습니다.")
                                raise KeyError(f"해당하는 스타일이 없습니다: {style_name}")
    
                        hwp.TableCellBlock()
                        hwp.TableRightCell()
                        hwp.TableCellBlockExtend()
                        hwp.TableColPageDown()
                        hwp.TableUpperCell()
    
                        pset = hwp.HParameterSet.HStyle
                        hwp.HAction.GetDefault("StyleEx", pset.HSet)
                        pset.Apply = style_name
                        hwp.HAction.Execute("StyleEx", pset.HSet)

            스타일뿌셔(hwp, "동물상,Korean Name")
        
            hwp.get_into_nth_table(n)
            hwp.TableCellBlock()
            hwp.TableLowerCell()
            hwp.TableCellBlockExtend()
            hwp.TableColPageDown()
            hwp.TableUpperCell()

            pset=hwp.HParameterSet.HSort
            hwp.HAction.GetDefault("Sort", pset.HSet)
            pset.CreateItemArray("KeyOption", 2)
            pset.KeyOption.SetItem(0,1)
            pset.KeyOption.SetItem(1,2)
            hwp.HAction.Execute("Sort", pset.HSet)
        
            n+=1
        ctrl = ctrl.Next

def 필요없는것():
    hwp=Hwp()
    cv = get_Cv()
       
    ctrl = hwp.HeadCtrl
    n=0

    while ctrl:
        if ctrl.UserDesc == "표":
            if ctrl.Next is None:
                if cv == 2:
                    hwp.get_into_nth_table(n)
                    hwp.TableCellBlock()
                    hwp.TableColEnd()
                    hwp.HAction.GetDefault("TableDeleteRowColumn", pset.HSet)
                    hwp.HAction.Execute("TableDeleteRowColumn", pset.HSet)
                    
                    hwp.get_into_nth_table(n)
                    hwp.TableColPageDown()
                    
                    for _ in range(5):
                        hwp.TableCellBlock()
                        hwp.TableCellBlockExtend()
                        hwp.TableRightCell()
                        hwp.HAction.Run("TableMergeCell")
                        hwp.TableUpperCell()
                    n+=1
                else:
                    hwp.get_into_nth_table(n)
                    hwp.TableCellBlock()

                    pset=hwp.HParameterSet.HTableDeleteLine

                    hwp.HAction.GetDefault("TableDeleteRowColumn", pset.HSet)
                    hwp.HAction.Execute("TableDeleteRowColumn", pset.HSet)

                    hwp.get_into_nth_table(n)
                    hwp.TableCellBlock()
                    hwp.TableColEnd()
                    hwp.HAction.GetDefault("TableDeleteRowColumn", pset.HSet)
                    hwp.HAction.Execute("TableDeleteRowColumn", pset.HSet)
                    
                    hwp.get_into_nth_table(n)
                    hwp.TableCellBlock()
                    hwp.TableColPageDown()
                    hwp.TableCellBlockExtend()
                    hwp.TableRightCell()
                    hwp.HAction.Run("TableMergeCell")
                    hwp.TableUpperCell()
                    n+=1
            else:
                hwp.get_into_nth_table(n)
                hwp.TableCellBlock()

                pset=hwp.HParameterSet.HTableDeleteLine

                hwp.HAction.GetDefault("TableDeleteRowColumn", pset.HSet)
                hwp.HAction.Execute("TableDeleteRowColumn", pset.HSet)

                hwp.get_into_nth_table(n)
                hwp.TableCellBlock()
                hwp.TableColEnd()
                hwp.HAction.GetDefault("TableDeleteRowColumn", pset.HSet)
                hwp.HAction.Execute("TableDeleteRowColumn", pset.HSet)
                
                hwp.get_into_nth_table(n)
                hwp.TableCellBlock()
                hwp.TableColPageDown()
                hwp.TableCellBlockExtend()
                hwp.TableRightCell()
                hwp.HAction.Run("TableMergeCell")
                hwp.TableUpperCell()
                n+=1
        ctrl = ctrl.Next

def 빈공간지우기():
    hwp=Hwp()
    
    pset = hwp.HParameterSet.HFindReplace
    hwp.CreateAction("AllReplace")
    hwp.HAction.GetDefault("AllReplace", pset.HSet)
    pset.Direction = 0  # 찾기 방향: 앞으로
    pset.FindString = "　"  # 찾을 문자열
    pset.ReplaceString = ""
    pset.ReplaceMode = 1
    pset.ReplaceStyle = ""
    pset.IgnoreMessage = 1
    hwp.HAction.Execute("AllReplace", pset.HSet)

def 표정리():
    try:
        hwp=Hwp()
        hwp.ShapeObjTableSelCell()
        hwp.TableCellBlock()
        hwp.TableCellBlockExtend()
        hwp.TableCellBlockExtend()
        hwp.TableCellBlockExtend()

        #표속성 변경
        pset = hwp.HParameterSet.HShapeObject
        hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
        pset.CellMarginBottom = hwp.MiliToHwpUnit(0.5)
        pset.CellMarginTop = hwp.MiliToHwpUnit(0.5)
        pset.CellMarginRight = hwp.MiliToHwpUnit(0.5)
        pset.CellMarginLeft = hwp.MiliToHwpUnit(0.5)
        pset.OutsideMarginBottom = hwp.MiliToHwpUnit(0.5)
        pset.OutsideMarginTop = hwp.MiliToHwpUnit(0.5)
        pset.OutsideMarginRight = hwp.MiliToHwpUnit(0.5)
        pset.OutsideMarginLeft = hwp.MiliToHwpUnit(0.5)

        pset.ShapeTableCell.VertAlign = hwp.VAlign("Center")
        pset.ShapeTableCell.LineWrap = 1
        hwp.HAction.Execute("TablePropertyDialog", pset.HSet)
            
        #문자스타일 변경
        pset = hwp.HParameterSet.HCharShape
        hwp.HAction.GetDefault("CharShape", pset.HSet)
        hwp.HAction.Execute("CharShape", pset.HSet)
        
        #선긋기기
        pset = hwp.HParameterSet.HCellBorderFill
        hwp.HAction.GetDefault("CellBorderFill", pset.HSet)
        pset.BorderWidthBottom = hwp.HwpLineWidth("0.4mm")
        pset.BorderTypeBottom = hwp.HwpLineType("Solid")
        pset.BorderWidthTop = hwp.HwpLineWidth("0.4mm")
        pset.BorderTypeTop = hwp.HwpLineType("Solid")
        pset.BorderWidthRight = hwp.HwpLineWidth("0.4mm")
        pset.BorderTypeRight = hwp.HwpLineType("Solid")
        pset.BorderWidthLeft = hwp.HwpLineWidth("0.4mm")
        pset.BorderTypeLeft = hwp.HwpLineType("Solid")
        pset.TypeVert = hwp.HwpLineType("Solid")
        pset.WidthVert = hwp.HwpLineWidth("0.12mm")
        hwp.HAction.Execute("CellBorderFill", pset.HSet)

        #헤드 실선
        hwp.TableColPageUp()
        hwp.HAction.GetDefault("CellBorderFill", pset.HSet)
        pset.BorderWidthBottom = hwp.HwpLineWidth("0.5mm")
        pset.BorderTypeBottom = hwp.HwpLineType("DoubleSlim")
        hwp.HAction.Execute("CellBorderFill", pset.HSet)

        hwp.Cancel()

        hwp.TableColBegin()
        hwp.TableColPageUp()
        hwp.TableCellBlock()
        hwp.TableLowerCell()
        hwp.TableRightCell()
        hwp.TableCellBlockExtend()
        hwp.TableColPageDown()
        hwp.TableUpperCell()
            
        pset=hwp.HParameterSet.HParaShape
            
        hwp.HAction.GetDefault("ParagraphShape", pset.HSet)
        pset.LeftMargin = hwp.PointToHwpUnit(30.0)
        hwp.HAction.Execute("ParagraphShape", pset.HSet)

        hwp.Cancel()
        
        hwp.TableColBegin()
        hwp.TableColPageUp()
        hwp.TableCellBlock()
        hwp.TableRightCell()
        hwp.TableRightCell()
        hwp.TableLowerCell()
        hwp.TableCellBlockExtend()
        hwp.TableColPageDown()
        hwp.TableUpperCell()
            
        hwp.HAction.GetDefault("ParagraphShape", pset.HSet)
        pset.LeftMargin = hwp.PointToHwpUnit(20.0)
        hwp.HAction.Execute("ParagraphShape", pset.HSet)
        
        hwp.Cancel()
        
        hwp.TableColBegin()
        hwp.TableColPageUp()
        hwp.TableCellBlock()
        hwp.TableColEnd()
        hwp.TableLowerCell()
        hwp.TableCellBlockExtend()
        hwp.TableColPageDown()
        hwp.TableUpperCell()

        pset=hwp.HParameterSet.HSort
        hwp.HAction.GetDefault("Sort", pset.HSet)
        pset.CreateItemArray("KeyOption", 2)
        pset.KeyOption.SetItem(0,1)
        pset.KeyOption.SetItem(1,2)
        hwp.HAction.Execute("Sort", pset.HSet)
        
        hwp.Cancel()

        hwp.TableColBegin()
        hwp.TableColPageUp()

        def 스타일뿌셔(hwp, style_name):
            pset = hwp.HParameterSet.HFindReplace
            hwp.HAction.GetDefault("RepeatFind", pset.HSet)
            pset.FindString = "family"  # 찾을 문자열
            pset.Direction = 0  # 찾기 방향: 앞으로
            pset.IgnoreMessage = 1
            hwp.HAction.Execute("RepeatFind", pset.HSet)
        
            if type(style_name) != int:
                style_dict = hwp.get_style_dict(as_="dict")
                found = False
                for key, value in style_dict.items():
                    if value.get('Name') == style_name:
                        style_name = key
                        found = True
                        break
                if not found:
                    raise KeyError(f"해당하는 스타일이 없습니다: {style_name}")
            
            hwp.TableCellBlock()
            hwp.TableRightCell()
            hwp.TableCellBlockExtend()
            hwp.TableColPageDown()
            hwp.TableUpperCell()
        
            pset = hwp.HParameterSet.HStyle
            hwp.HAction.GetDefault("StyleEx", pset.HSet)
            pset.Apply = style_name
            hwp.HAction.Execute("StyleEx", pset.HSet)

        스타일뿌셔(hwp, "동물상,Korean Name")

        hwp.TableColBegin()
        hwp.TableColPageUp()
        hwp.TableCellBlock()
        hwp.TableLowerCell()
        hwp.TableCellBlockExtend()
        hwp.TableColPageDown()
        hwp.TableUpperCell()

        pset=hwp.HParameterSet.HSort
        hwp.HAction.GetDefault("Sort", pset.HSet)
        pset.CreateItemArray("KeyOption", 2)
        pset.KeyOption.SetItem(0,1)
        pset.KeyOption.SetItem(1,2)
        hwp.HAction.Execute("Sort", pset.HSet)

        hwp.Cancel()
        time.sleep(2)
        
        #마지막에 넣어야 해결될듯?
        pset = hwp.HParameterSet.HFindReplace
        hwp.HAction.GetDefault("AllReplace", pset.HSet)
        pset.FindString = "출현종수"  # 찾을 문자열
        pset.ReplaceString = "출현종수"
        pset.ReplaceMode = 1
        pset.Direction = 0  # 찾기 방향: 앞으로
        pset.IgnoreMessage = 1
        pset.ReplaceStyle = ""
        pset.ReplaceParaShape.LeftMargin = hwp.PointToHwpUnit(0.0)
        hwp.HAction.Execute("AllReplace", pset.HSet)

        학명스타일("Family","동물상,Family")
        학명스타일("Order","동물상,Order")
        학명스타일("Class","동물상,Class")
        
        hwp.TableColBegin()
        hwp.TableColPageUp()
        hwp.TableCellBlock()

        pset=hwp.HParameterSet.HTableDeleteLine

        hwp.HAction.GetDefault("TableDeleteRowColumn", pset.HSet)
        hwp.HAction.Execute("TableDeleteRowColumn", pset.HSet)

        window = get_window()
        window.show_message_signal.emit("종 개별작업 완료")
    
    except Exception as e:
        print(f"오류 발생: {str(e)}")

def 조류군집표():
    try:
        hwp=hwp_opened()
        hwp=Hwp()
        
        hwp.ShapeObjTableSelCell()
        hwp.TableCellBlock()
        hwp.TableLowerCell()
        hwp.TableCellBlockExtend()
        hwp.TableRightCell()
        hwp.TableColPageDown()
        hwp.TableUpperCell()
        hwp.TableUpperCell()
        hwp.TableUpperCell()
        hwp.TableUpperCell()
        hwp.TableUpperCell()
        
        pset=hwp.HParameterSet.HParaShape
            
        hwp.HAction.GetDefault("ParagraphShape", pset.HSet)
        pset.LeftMargin = hwp.PointToHwpUnit(10.0)
        hwp.HAction.Execute("ParagraphShape", pset.HSet)

        hwp.TableColBegin()
        hwp.TableColPageUp()
        hwp.TableCellBlock()
        hwp.TableCellBlockExtend()
        hwp.TableCellBlockExtend()

        #문자스타일 변경
        pset = hwp.HParameterSet.HCharShape
        hwp.HAction.GetDefault("CharShape", pset.HSet)
        hwp.HAction.Execute("CharShape", pset.HSet)
            
        #표속성 변경
        pset = hwp.HParameterSet.HShapeObject
        hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
        pset.CellMarginBottom = hwp.MiliToHwpUnit(0.5)
        pset.CellMarginTop = hwp.MiliToHwpUnit(0.5)
        pset.CellMarginRight = hwp.MiliToHwpUnit(0.5)
        pset.CellMarginLeft = hwp.MiliToHwpUnit(0.5)
        pset.OutsideMarginBottom = hwp.MiliToHwpUnit(0.5)
        pset.OutsideMarginTop = hwp.MiliToHwpUnit(0.5)
        pset.OutsideMarginRight = hwp.MiliToHwpUnit(0.5)
        pset.OutsideMarginLeft = hwp.MiliToHwpUnit(0.5)
        pset.ShapeTableCell.VertAlign = hwp.VAlign("Center")
        pset.ShapeTableCell.LineWrap = 1
        hwp.HAction.Execute("TablePropertyDialog", pset.HSet)

        pset = hwp.HParameterSet.HCellBorderFill
                
        #바깥 테두리
        hwp.HAction.GetDefault("CellBorderFill", pset.HSet)
        pset.BorderWidthBottom = hwp.HwpLineWidth("0.4mm")
        pset.BorderTypeBottom = hwp.HwpLineType("Solid")
        pset.BorderWidthTop = hwp.HwpLineWidth("0.4mm")
        pset.BorderTypeTop = hwp.HwpLineType("Solid")
        pset.BorderWidthRight = hwp.HwpLineWidth("0.4mm")
        pset.BorderTypeRight = hwp.HwpLineType("Solid")
        pset.BorderWidthLeft = hwp.HwpLineWidth("0.4mm")
        pset.BorderTypeLeft = hwp.HwpLineType("Solid")
        pset.TypeVert = hwp.HwpLineType("Solid")
        pset.WidthVert = hwp.HwpLineWidth("0.12mm")
        hwp.HAction.Execute("CellBorderFill", pset.HSet)
                
        #아래쪽 분석값 실선
        for _ in range(5):
            hwp.TableUpperCell()
            hwp.HAction.GetDefault("CellBorderFill", pset.HSet)
            pset.BorderWidthBottom = hwp.HwpLineWidth("0.12mm")
            pset.BorderTypeBottom = hwp.HwpLineType("Solid")
            hwp.HAction.Execute("CellBorderFill", pset.HSet)
                
        #수평선 점선
        hwp.HAction.GetDefault("CellBorderFill", pset.HSet)
        pset.TypeHorz = hwp.HwpLineType("Dot")
        pset.WidthHorz = hwp.HwpLineWidth("0.12mm")
        hwp.HAction.Execute("CellBorderFill", pset.HSet)
                
        #헤드 실선
        hwp.TableColPageUp()
        hwp.HAction.GetDefault("CellBorderFill", pset.HSet)
        pset.BorderWidthBottom = hwp.HwpLineWidth("0.5mm")
        pset.BorderTypeBottom = hwp.HwpLineType("DoubleSlim")
        hwp.HAction.Execute("CellBorderFill", pset.HSet)
        hwp.Cancel()

        hwp.TableColBegin()
        hwp.TableColPageUp()
        hwp.TableCellBlock()
        hwp.TableColEnd()
        hwp.HAction.GetDefault("TableDeleteRowColumn", pset.HSet)
        hwp.HAction.Execute("TableDeleteRowColumn", pset.HSet)

        pset = hwp.HParameterSet.HFindReplace
        hwp.CreateAction("AllReplace")
        hwp.HAction.GetDefault("AllReplace", pset.HSet)
        pset.Direction = 0  # 찾기 방향: 앞으로
        pset.FindString = "　"  # 찾을 문자열
        pset.ReplaceString = ""
        pset.ReplaceMode = 1
        pset.ReplaceStyle = ""
        pset.IgnoreMessage = 1
        hwp.HAction.Execute("AllReplace", pset.HSet)
        
        hwp.TableColBegin()
        hwp.TableColPageDown()
                    
        for _ in range(5):
            hwp.TableCellBlock()
            hwp.TableCellBlockExtend()
            hwp.TableRightCell()
            hwp.HAction.Run("TableMergeCell")
            hwp.TableUpperCell()
            
        window = get_window()
        window.show_message_signal.emit("조류군집표 작업 완료")
    except Exception as e:
        print(f"오류 발생: {str(e)}")

def 곤충문단():
    try:
        hwp=hwp_opened()
        hwp=Hwp()
        
        ctrl = hwp.HeadCtrl
        n=0

        while ctrl:
            if ctrl.UserDesc == "표":
                if n == 0:
                    hwp.get_into_nth_table(n)
                    hwp.TableCellBlock()
                    hwp.TableLowerCell()
                    hwp.TableRightCell()
                    hwp.TableCellBlockExtend()
                    hwp.TableColPageDown()

                elif ctrl.Next is None:
                    hwp.get_into_nth_table(n)
                    hwp.TableCellBlock()
                    hwp.TableRightCell()
                    hwp.TableCellBlockExtend()
                    hwp.TableColPageDown()
                    hwp.TableUpperCell()
                else:
                    hwp.get_into_nth_table(n)
                    hwp.TableCellBlock()
                    hwp.TableRightCell()
                    hwp.TableCellBlockExtend()
                    hwp.TableColPageDown()
          
                pset=hwp.HParameterSet.HParaShape
                hwp.HAction.GetDefault("ParagraphShape", pset.HSet)
                pset.LeftMargin = hwp.PointToHwpUnit(30.0)
                hwp.HAction.Execute("ParagraphShape", pset.HSet)
                
                if n == 0:
                    hwp.get_into_nth_table(n)
                    hwp.TableCellBlock()
                    hwp.TableLowerCell()
                    hwp.TableRightCell()
                    hwp.TableRightCell()
                    hwp.TableCellBlockExtend()
                    hwp.TableColPageDown()

                elif ctrl.Next is None:
                    hwp.get_into_nth_table(n)
                    hwp.TableCellBlock()
                    hwp.TableRightCell()
                    hwp.TableRightCell()
                    hwp.TableCellBlockExtend()
                    hwp.TableColPageDown()
                    hwp.TableUpperCell()
                else:
                    hwp.get_into_nth_table(n)
                    hwp.TableCellBlock()
                    hwp.TableRightCell()
                    hwp.TableRightCell()
                    hwp.TableCellBlockExtend()
                    hwp.TableColPageDown()
                    
                hwp.HAction.GetDefault("ParagraphShape", pset.HSet)
                pset.LeftMargin = hwp.PointToHwpUnit(20.0)
                hwp.HAction.Execute("ParagraphShape", pset.HSet)
                n +=1
                    
            ctrl = ctrl.Next
        
    except Exception as e:
        print(f"오류 발생: {str(e)}")

def 곤충기본():
    try:
        hwp=hwp_opened()
        hwp=Hwp()
        
        # 페이지 설정 조정
        pset = hwp.HParameterSet.HSecDef
        hwp.HAction.GetDefault("PageSetup", pset.HSet)
        # 페이지 여백 및 방향 설정
        pset.PageDef.LeftMargin = hwp.MiliToHwpUnit(5.0)  # 왼쪽 여백 5mm
        pset.PageDef.RightMargin = hwp.MiliToHwpUnit(5.0)  # 오른쪽 여백 5mm
        pset.PageDef.TopMargin = hwp.MiliToHwpUnit(5.0)  # 위쪽 여백 5mm
        pset.PageDef.BottomMargin = hwp.MiliToHwpUnit(5.0)  # 아래쪽 여백 5mm
        pset.PageDef.HeaderLen = hwp.MiliToHwpUnit(5.0)  # 머리글 길이 5mm
        pset.PageDef.FooterLen = hwp.MiliToHwpUnit(5.0)  # 바닥글 길이 5mm
        pset.PageDef.Landscape = 1  # 가로 방향 설정
        # 적용 범위 설정
        pset.HSet.SetItem("ApplyClass", 24)
        pset.HSet.SetItem("ApplyTo", 3)
        # 페이지 설정 실행
        hwp.HAction.Execute("PageSetup", pset.HSet)

        ctrl = hwp.HeadCtrl
        cv = get_Cv()
        n=0

        while ctrl:
            if ctrl.UserDesc == "표":
                hwp.get_into_nth_table(n)
                hwp.TableCellBlock()
                hwp.TableRightCell()
                hwp.TableRightCell()
                hwp.TableRightCell()
                hwp.TableCellBlockExtend()
                hwp.TableColEnd()
                hwp.TableColPageDown()
                
                pset = hwp.HParameterSet.HShapeObject
                hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
                pset.HSet.SetItem("ShapeType", 3)  # 표 모양 유형 설정
                pset.HSet.SetItem("ShapeCellSize", 1)  # 셀 크기 설정
                pset.ShapeTableCell.Width = hwp.MiliToHwpUnit(10.0)
                hwp.HAction.Execute("TablePropertyDialog", pset.HSet)
                
                hwp.get_into_nth_table(n)
                hwp.TableCellBlock()
                hwp.TableCellBlockExtend()
                hwp.TableColEnd()
                
                if n == 0:
                    # 줄간격 설정
                    pset = hwp.HParameterSet.HParaShape
                    hwp.HAction.GetDefault("ParagraphShape", pset.HSet)
                    line_spacing = get_window().line.value()
                    pset.LineSpacing = line_spacing
                    hwp.HAction.Execute("ParagraphShape", pset.HSet)
                
                hwp.TableColPageDown()
                
                # 문자 스타일 변경 - UI에서 폰트와 크기 가져오기
                pset = hwp.HParameterSet.HCharShape
                hwp.HAction.GetDefault("CharShape", pset.HSet)
                # UI에서 선택된 폰트와 크기 가져오기
                selected_font = get_window().get_selected_font()
                font_size = get_window().get_font_size()
                font_size = get_window().font_size.value()
                
                # 폰트 설정
                pset.FaceNameHangul = selected_font
                pset.FaceNameLatin = selected_font
                pset.FaceNameHanja = selected_font
                pset.FaceNameJapanese = selected_font
                pset.FaceNameOther = selected_font
                pset.FaceNameSymbol = selected_font
                pset.FaceNameUser = selected_font
                
                pset.FontTypeHangul = 1
                pset.FontTypeLatin = 1
                pset.FontTypeHanja = 1
                pset.FontTypeJapanese = 1
                pset.FontTypeOther = 1
                pset.FontTypeSymbol = 1
                pset.FontTypeUser = 1
                
                # 크기 설정
                pset.Height = hwp.PointToHwpUnit(font_size)
                
                hwp.HAction.Execute("CharShape", pset.HSet)
                
                #표속성 변경
                pset = hwp.HParameterSet.HShapeObject
                hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
                pset.CellMarginBottom = hwp.MiliToHwpUnit(0.5)
                pset.CellMarginTop = hwp.MiliToHwpUnit(0.5)
                pset.CellMarginRight = hwp.MiliToHwpUnit(0.5)
                pset.CellMarginLeft = hwp.MiliToHwpUnit(0.5)
                pset.OutsideMarginBottom = hwp.MiliToHwpUnit(0.5)
                pset.OutsideMarginTop = hwp.MiliToHwpUnit(0.5)
                pset.OutsideMarginRight = hwp.MiliToHwpUnit(0.5)
                pset.OutsideMarginLeft = hwp.MiliToHwpUnit(0.5)

                pset.ShapeTableCell.VertAlign = hwp.VAlign("Center")
                pset.ShapeTableCell.LineWrap = 1
                
                # 표의 셀 높이 조절
                pset.HSet.SetItem("ShapeType", 3)  # 표 모양 유형 설정
                pset.HSet.SetItem("ShapeCellSize", 1)  # 셀 크기 설정
                pset.ShapeTableCell.Height = hwp.MiliToHwpUnit(4.0)  # 높이를 4mm로 설정
                
                hwp.HAction.Execute("TablePropertyDialog", pset.HSet)
                
                pset = hwp.HParameterSet.HCellBorderFill
                pset.BorderWidthRight = hwp.HwpLineWidth("0.4mm")
                pset.BorderTypeRight = hwp.HwpLineType("Solid")
                pset.BorderWidthLeft = hwp.HwpLineWidth("0.4mm")
                pset.BorderTypeLeft = hwp.HwpLineType("Solid")
                pset.TypeVert = hwp.HwpLineType("Solid")
                pset.WidthVert = hwp.HwpLineWidth("0.12mm")
                hwp.HAction.Execute("CellBorderFill", pset.HSet)
                
                if ctrl.Next is None:
                        pset = hwp.HParameterSet.HCellBorderFill
                    
                        #바깥 테두리
                        hwp.HAction.GetDefault("CellBorderFill", pset.HSet)
                        pset.BorderWidthBottom = hwp.HwpLineWidth("0.4mm")
                        pset.BorderTypeBottom = hwp.HwpLineType("Solid")
                        
                if n == 0:
                    #헤드 실선
                    hwp.TableColPageUp()
                    hwp.HAction.GetDefault("CellBorderFill", pset.HSet)
                    pset.BorderWidthBottom = hwp.HwpLineWidth("0.5mm")
                    pset.BorderTypeBottom = hwp.HwpLineType("DoubleSlim")
                    hwp.HAction.Execute("CellBorderFill", pset.HSet)

                n += 1
            ctrl = ctrl.Next
    except Exception as e:
        print(f"오류 발생: {str(e)}")

def 곤충국명():
    try:
        hwp=hwp_opened()
        hwp=Hwp()
        
        ctrl = hwp.HeadCtrl
        n=0

        while ctrl:
            if ctrl.UserDesc == "표":
                if ctrl.Next is None:
                    hwp.get_into_nth_table(n)
                    hwp.TableCellBlock()
                    hwp.TableColEnd()
                    hwp.TableCellBlockExtend()
                    hwp.TableColPageDown()
                    hwp.TableUpperCell()
                elif n == 0:
                    hwp.get_into_nth_table(n)
                    hwp.TableCellBlock()
                    hwp.TableColEnd()
                    hwp.TableLowerCell()
                    hwp.TableCellBlockExtend()
                    hwp.TableColPageDown()
                else:
                    hwp.get_into_nth_table(n)
                    hwp.TableCellBlock()
                    hwp.TableColEnd()
                    hwp.TableCellBlockExtend()
                    hwp.TableColPageDown()

                pset=hwp.HParameterSet.HSort
                hwp.HAction.GetDefault("Sort", pset.HSet)
                pset.CreateItemArray("KeyOption", 2)
                pset.KeyOption.SetItem(0,1)
                pset.KeyOption.SetItem(1,2)
                hwp.HAction.Execute("Sort", pset.HSet)

                def 곤충뿌셔(hwp, style_name):
                    hwp.get_into_nth_table(n)
                    pset = hwp.HParameterSet.HFindReplace
                    hwp.HAction.GetDefault("RepeatFind", pset.HSet)
                    pset.FindString = "family"  # 찾을 문자열
                    pset.Direction = 0  # 찾기 방향: 앞으로
                    pset.IgnoreMessage = 1
                    hwp.HAction.Execute("RepeatFind", pset.HSet)

                    if type(style_name) != int:
                        style_dict = hwp.get_style_dict()
                        for key, value in style_dict.items():
                            if value.get('Name') == style_name:
                                style_name = key
                                break
                        else:
                            raise KeyError("해당하는 스타일이 없습니다.")

                    hwp.TableCellBlock()
                    hwp.TableRightCell()
                    hwp.TableCellBlockExtend()
                    hwp.TableColPageDown()
                    
                    if ctrl.Next is None:
                        hwp.TableUpperCell()

                    pset = hwp.HParameterSet.HStyle
                    hwp.HAction.GetDefault("StyleEx", pset.HSet)
                    pset.Apply = style_name
                    hwp.HAction.Execute("StyleEx", pset.HSet)

                곤충뿌셔(hwp, "동물상,Korean Name")
            
                if ctrl.Next is None:
                    hwp.get_into_nth_table(n)
                    hwp.TableCellBlock()
                    hwp.TableCellBlockExtend()
                    hwp.TableColPageDown()
                    hwp.TableUpperCell()
                elif n == 0:
                    hwp.get_into_nth_table(n)
                    hwp.TableCellBlock()
                    hwp.TableLowerCell()
                    hwp.TableCellBlockExtend()
                    hwp.TableColPageDown()
                else:
                    hwp.get_into_nth_table(n)
                    hwp.TableCellBlock()
                    hwp.TableCellBlockExtend()
                    hwp.TableColPageDown()

                pset=hwp.HParameterSet.HSort
                hwp.HAction.GetDefault("Sort", pset.HSet)
                pset.CreateItemArray("KeyOption", 2)
                pset.KeyOption.SetItem(0,1)
                pset.KeyOption.SetItem(1,2)
                hwp.HAction.Execute("Sort", pset.HSet)

                hwp.get_into_nth_table(n)
                hwp.TableCellBlock()

                pset=hwp.HParameterSet.HTableDeleteLine

                hwp.HAction.GetDefault("TableDeleteRowColumn", pset.HSet)
                hwp.HAction.Execute("TableDeleteRowColumn", pset.HSet)

                hwp.get_into_nth_table(n)
                hwp.TableCellBlock()
                hwp.TableColEnd()
                
                hwp.HAction.GetDefault("TableDeleteRowColumn", pset.HSet)
                hwp.HAction.Execute("TableDeleteRowColumn", pset.HSet)
                
                n+=1
            ctrl = ctrl.Next
           
    except Exception as e:
        print(f"오류 발생: {str(e)}")