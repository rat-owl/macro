# logic/styles.py

from pyhwpx import Hwp
from core.com_utils import hwp_opened
from core.global_state import get_Cv


def remove_superscript():
    hwp = hwp_opened()
    hwp = Hwp()
    find_strings = ["※", "☆", "★", "◈", "▣"]

    for find_string in find_strings:
        while True:
            hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
            hwp.HParameterSet.HFindReplace.FindString = find_string
            hwp.HParameterSet.HFindReplace.Direction = 0
            hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
            found = hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
            if not found:
                break
            hwp.HAction.Run("CharShapeSuperscript")
    print("윗첨자가 제거되었습니다.")


def 학명스타일(find_name, style_name):
    hwp = Hwp()
    pset = hwp.HParameterSet.HFindReplace

    hwp.HAction.GetDefault("AllReplace", pset.HSet)
    pset.FindString = find_name
    pset.ReplaceString = find_name
    pset.ReplaceMode = 1
    pset.Direction = 0
    pset.IgnoreMessage = 1
    pset.ReplaceStyle = style_name
    hwp.HAction.Execute("AllReplace", pset.HSet)

    hwp.HAction.GetDefault("AllReplace", pset.HSet)
    pset.FindString = find_name
    pset.ReplaceString = find_name
    pset.ReplaceMode = 1
    pset.Direction = 0
    pset.IgnoreMessage = 1
    pset.ReplaceStyle = style_name
    return hwp.HAction.Execute("AllReplace", pset.HSet)


def 국명():
    hwp = Hwp()
    ctrl = hwp.HeadCtrl
    n = 0
    cv = get_Cv()

    while ctrl:
        if ctrl.UserDesc == "표":
            if ctrl.Next is None and cv == 2:
                break
            else:
                hwp.get_into_nth_table(n)
                hwp.TableCellBlock()
                hwp.TableColEnd()
                hwp.TableLowerCell()
                hwp.TableCellBlockExtend()
                hwp.TableColPageDown()
                hwp.TableUpperCell()

                pset = hwp.HParameterSet.HSort
                hwp.HAction.GetDefault("Sort", pset.HSet)
                pset.CreateItemArray("KeyOption", 2)
                pset.KeyOption.SetItem(0, 1)
                pset.KeyOption.SetItem(1, 2)
                hwp.HAction.Execute("Sort", pset.HSet)

                def 스타일뿌셔(hwp, style_name):
                    hwp.get_into_nth_table(n)
                    pset = hwp.HParameterSet.HFindReplace
                    hwp.HAction.GetDefault("RepeatFind", pset.HSet)
                    pset.FindString = "family"
                    pset.Direction = 0
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

                hwp.get_into_nth_table(n)
                hwp.TableCellBlock()
                hwp.TableLowerCell()
                hwp.TableCellBlockExtend()
                hwp.TableColPageDown()
                hwp.TableUpperCell()

                pset = hwp.HParameterSet.HSort
                hwp.HAction.GetDefault("Sort", pset.HSet)
                pset.CreateItemArray("KeyOption", 2)
                pset.KeyOption.SetItem(0, 1)
                pset.KeyOption.SetItem(1, 2)
                hwp.HAction.Execute("Sort", pset.HSet)
                n += 1
        ctrl = ctrl.Next
