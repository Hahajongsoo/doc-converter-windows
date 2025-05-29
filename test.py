import win32com.client as win32
import os   
import time
import pythoncom

class HwpManager:
    def __init__(self):
        self.hwp = None
        
    def __enter__(self):
        pythoncom.CoInitialize()
        self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        return self.hwp
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            if self.hwp:
                self.hwp.Quit()
        except:
            pass
        finally:
            pythoncom.CoUninitialize()

    
def extract_red_text(hwp):
    table_pos = []
    ctrl = hwp.HeadCtrl
    while ctrl != None:
        if ctrl.UserDesc == "표":
            hwp.SetPosBySet(ctrl.GetAnchorPos(0))
            hwp.Run("MoveDown")
            table_pos.append(hwp.GetPos()[0])
        ctrl = ctrl.Next
    table_pos = table_pos[3:]
    all_results = [[] for _ in range(len(table_pos))]
    find_replace_pset = hwp.HParameterSet.HFindReplace
    hwp.HAction.GetDefault("RepeatFind", find_replace_pset.HSet)
    find_replace_pset.FindCharShape.TextColor = 0x000000FF
    find_replace_pset.IgnoreMessage = 1
    find_replace_pset.FindType = 1   

    i = 0
    row_check = 0
    table_results = []
    while hwp.HAction.Execute("RepeatFind", find_replace_pset.HSet):
        # 표의 pos.list와 현재 pos.list가 같으면 새로운 리스트에 담을 준비
        # 표를 한 번 처리했다면 table_result를 all_results에 담기
        try:
            i = table_pos.index(hwp.GetPos()[0])
            if table_pos[i] == hwp.GetPos()[0]:
                if table_results:  # 빈 리스트가 아닐 때만 추가
                    all_results[i-1] = table_results
                table_results = []
        except ValueError:  # 구체적인 예외 처리
            pass

        # 2번 마다 새로운 리스트에 담을 준비
        if row_check % 2 == 0:
            row_results = []
        
        text = hwp.GetTextFile('UNICODE', 'saveblock')
        if text.strip():  # 빈 문자열이 아닐 때만 추가
            # 쉼표로 분리하고 각 항목의 앞뒤 공백 제거, 빈 문자열 제외
            words = [word.strip() for word in text.split(',') if word.strip()]
            row_results.extend(words)  # extend를 사용하여 리스트에 개별 항목 추가
        
        row_check += 1
        if row_check % 2 == 0 and row_results:  # row_results가 비어있지 않을 때만 추가
            table_results.append(row_results)

        hwp.HAction.GetDefault("RepeatFind", find_replace_pset.HSet)
        find_replace_pset.FindCharShape.TextColor = 0x000000FF
        find_replace_pset.IgnoreMessage = 1
        find_replace_pset.FindType = 1
    i = all_results.index([])
    all_results[i] = table_results
    return all_results

def create_synonym_questions_from_red_text(hwp, all_results):
    tables = []
    ctrl = hwp.HeadCtrl
    while ctrl != None:
        if ctrl.UserDesc == "표":
            tables.append(ctrl)
        ctrl = ctrl.Next
    tables = tables[3:]
    
    for i,table in enumerate(tables):
        first_words = {}
        rest_map = {}
        for group in all_results[i]:
            if not group:
                continue
            first, *rest = group
            first_words[first] = -1
            if rest:
                if isinstance(rest[0], str) and ',' in rest[0]:
                    rest = [word.strip() for word in rest[0].split(',')]
                rest_map[first] = rest
        if i == 0:
            goto_pset = hwp.HParameterSet.HGotoE
            hwp.HAction.GetDefault("Goto", goto_pset.HSet)
            goto_pset.HSet.SetItem("DialogResult", 2)
            goto_pset.SetSelectionIndex = 1
            hwp.HAction.Execute("Goto", goto_pset.HSet)
            spara = hwp.GetPos()[1]
            spos = hwp.GetPos()[2]
            hwp.SetPosBySet(tables[0].GetAnchorPos(0))
            hwp.Run("MoveLeft")
            epara = hwp.GetPos()[1]
            epos = hwp.GetPos()[2]
        else:
            hwp.HAction.Run("MoveRight")
            spara = hwp.GetPos()[1]
            spos = hwp.GetPos()[2]
            hwp.SetPosBySet(tables[1].GetAnchorPos(0))
            hwp.Run("MoveLeft")
            epara = hwp.GetPos()[1]
            epos = hwp.GetPos()[2]

        hwp.SelectText(spara, spos, epara, epos)
        hwp.Run("Copy")
        hwp.Run("FileNewTab")
        hwp.Run("Paste")

        find_replace_pset = hwp.HParameterSet.HFindReplace
        word_positions = []
        for f_word in first_words:
            hwp.HAction.GetDefault("RepeatFind", find_replace_pset.HSet)
            find_replace_pset.FindString = f_word
            find_replace_pset.WholeWordOnly = 1
            find_replace_pset.IgnoreMessage = 1
            find_replace_pset.FindCharShape.TextColor = 0
            hwp.HAction.Execute("RepeatFind", find_replace_pset.HSet)
            word_positions.append((f_word, hwp.GetPos()))
            hwp.HAction.Run("CharShapeBold")
            hwp.HAction.Run("CharShapeItalic")
            hwp.HAction.Run("CharShapeUnderline")
        sorted_words = [word for word, _ in sorted(word_positions, key=lambda x: x[1])]
        hwp.MovePos(3)
        hwp.HAction.Run("BreakPara")
        act = hwp.CreateAction("ParagraphShape") 
        pset = act.CreateSet() 
        act.GetDefault(pset)  
        pset.SetItem("LineSpacing", 200) 
        act.Execute(pset)
        time.sleep(10)
        for s_word in sorted_words:
            hwp.HAction.Run("BreakPara")
            new_spara, new_spos, new_epara, new_epos = hwp.GetPos()[1], hwp.GetPos()[2], hwp.GetPos()[1], hwp.GetPos()[2] + len(s_word)
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
            hwp.HParameterSet.HInsertText.Text = f"{s_word} 의 동의어를 쓰시오."
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
            hwp.SelectText(new_spara, new_spos, new_epara, new_epos)
            hwp.HAction.Run("CharShapeBold")
            hwp.HAction.Run("CharShapeItalic")
            hwp.HAction.Run("CharShapeUnderline")
            hwp.MovePos(7)
            if s_word in rest_map:
                synonyms = ", ".join(rest_map[s_word])
            hwp.HAction.Run("InsertEndnote")
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
            hwp.HParameterSet.HInsertText.Text = synonyms
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
            hwp.HAction.Run("CloseEx")
            hwp.HAction.Run("BreakPara")
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
            hwp.HParameterSet.HInsertText.Text = " → "
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
            for j, hint in enumerate(rest_map[s_word]):
                if " " not in hint:
                    new_spara, new_spos, new_epara, new_epos = hwp.GetPos()[1], hwp.GetPos()[2] + 1, hwp.GetPos()[1], hwp.GetPos()[2] + 1 + len(hint)
                    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
                    hwp.HParameterSet.HInsertText.Text = hint[0].lower() + " " * len(hint)
                    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
                    hwp.SelectText(new_spara, new_spos, new_epara, new_epos)
                    hwp.HAction.Run("CharShapeUnderline")
                    hwp.MovePos(7)
                    hwp.HAction.Run("CharShapeUnderline")
                else:
                    for h_word in hint.split():
                        new_spara, new_spos, new_epara, new_epos = hwp.GetPos()[1], hwp.GetPos()[2] + 1, hwp.GetPos()[1], hwp.GetPos()[2] + 1 + len(h_word)
                        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
                        hwp.HParameterSet.HInsertText.Text = h_word[0].lower() + " " * len(h_word) + " "
                        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
                        hwp.SelectText(new_spara, new_spos, new_epara, new_epos)
                        hwp.HAction.Run("CharShapeUnderline")
                        hwp.MovePos(7)
                    hwp.HAction.Run("CharShapeUnderline")

                if j != len(rest_map[s_word]) - 1:
                    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
                    hwp.HParameterSet.HInsertText.Text = ", "
                    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HAction.Run("SelectAll")
        hwp.HAction.Run("Copy")
        hwp.HAction.Run("WindowNextTab")
        hwp.SelectText(spara, spos, epara, epos)
        hwp.HAction.Run("Delete")
        hwp.HAction.Run("Paste")
        hwp.DeleteCtrl(table)
        hwp.XHwpDocuments.Item(1).SetActive_XHwpDocument()
        hwp.XHwpDocuments.Item(1).Close(0)
        
if __name__ == "__main__":
    with HwpManager() as hwp:
        hwp.Open(os.path.join(os.getcwd(), "test.hwp"))
        window = hwp.XHwpWindows.Item(0)
        window.Visible = True
        all_results = extract_red_text(hwp)
        create_synonym_questions_from_red_text(hwp, all_results)
        hwp.SaveAs(os.path.join(os.getcwd(), "test_result.hwp"))
        
        hwp.Clear(1)

