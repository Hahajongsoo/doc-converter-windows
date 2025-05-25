import win32com.client as win32
import os   
import time
import pythoncom
import win32clipboard

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
    all_results = []
    table_pos = []

    ctrl = hwp.HeadCtrl
    while ctrl != None:
        if ctrl.UserDesc == "표":
            hwp.SetPosBySet(ctrl.GetAnchorPos(0))
            hwp.Run("MoveDown")
            table_pos.append(hwp.GetPos()[0])
        ctrl = ctrl.Next
    
    find_replace_pset = hwp.HParameterSet.HFindReplace
    hwp.HAction.GetDefault("RepeatFind", find_replace_pset.HSet)
    find_replace_pset.FindCharShape.TextColor = 0x000000FF
    find_replace_pset.IgnoreMessage = 1
    find_replace_pset.FindType = 1   

    i = 0
    cell_end_pos = -1
    row_check = 0
    table_results = []
    while hwp.HAction.Execute("RepeatFind", find_replace_pset.HSet):
        # 표의 pos.list와 현재 pos.list가 같으면 새로운 리스트에 담을 준비
        # 표를 한 번 처리했다면 table_result를 all_results에 담기
        try:
            i = table_pos.index(hwp.GetPos()[0])
            if table_pos[i] == hwp.GetPos()[0]:
                if table_results:  # 빈 리스트가 아닐 때만 추가
                    all_results.append(table_results)
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
    all_results.append(table_results)
    return all_results

if __name__ == "__main__":
    with HwpManager() as hwp:
        hwp.Open(os.path.join(os.getcwd(), "test.hwp"))
        # hwp.XHwpWindows.Item(0).Visible = True
        all_results = extract_red_text(hwp)
        first_words = {}
        rest_map = {}
        print(all_results[0])
        for group in all_results[0]:
            if not group:
                continue
            first, *rest = group
            first_words[first] = -1
            if rest:
                if isinstance(rest[0], str) and ',' in rest[0]:
                    rest = [word.strip() for word in rest[0].split(',')]
                rest_map[first] = rest
        print(first_words)
        print(rest_map)
        hwp.Clear(1)

