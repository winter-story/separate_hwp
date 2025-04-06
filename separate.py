from pyhwpx import Hwp
import win32clipboard
# import tkinter as tk
# from tkinter import filedialog

# # Tkinter 기본 창 숨기기
# root = tk.Tk()
# root.withdraw()

# # 파일 선택 창 열기
# file_path = filedialog.askopenfilename(
#     title="파일을 선택하세요",
#     filetypes=[("모든 파일", "*.*")]  # 원하는 파일 형식으로 바꿀 수 있어요
# )

hwp = Hwp()
hwp.Open("C:\\mjwork\\hwp_files\\250224 감정보완 2차 (판교밸리 호반써밋)_VER1.hwp")

# 커서 전체 선택
hwp.MovePos(3)  # 문서 시작점으로 이동
hwp.HAction.Run("SelectAll")  # 전체 선택

# 선택된 내용 복사
hwp.HAction.Run("Copy")

# 클립보드에서 텍스트 가져오기
win32clipboard.OpenClipboard()
data = win32clipboard.GetClipboardData()
win32clipboard.CloseClipboard()

# 출력
print(data)
hwp.Quit()