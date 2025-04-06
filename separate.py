from pyhwpx import Hwp
import pyhwpx
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
hwp.MovePos(1000)
# hwp.HAction.Run("SelectAll")  # 전체 선택

# # 선택된 내용 복사
# hwp.HAction.Run("Copy")
#
# # 클립보드에서 텍스트 가져오기
# win32clipboard.OpenClipboard()
# data = win32clipboard.GetClipboardData()
# win32clipboard.CloseClipboard()
#
# # 출력
# print(data)
# hwp.Quit()

i = 0
while hwp.set_pos(0, i, 0):  # 본문의 para만 순회하면서
    if hwp.GetHeadingString():  # 문단번호가 매겨져 있으면
        prop = hwp.ParaShape
        prop.SetItem("PrevSpacing", hwp.PointToHwpUnit(500))
        hwp.ParaShape = prop
    i += 1

# # 개요를 찾기 위한 함수 (예시)
# def get_outline(document):
#     outline = []
#     return document.ParaShape.Item("AlignType")
#     # # 문서의 각 항목들을 검색
#     # for section in document.sections():
#     #     for paragraph in section.paragraphs():
#     #         # 예시: 제목 스타일이 있는 문단을 개요로 간주
#     #         if paragraph.style == 'Heading':
#     #             outline.append(paragraph.text)
#     #
#     # return outline
#
#
# # 개요 출력
# parashape = get_outline(hwp)
# print("문서 개요:", parashape)