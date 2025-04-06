import win32com.client

# 한글 열기
hwp = win32com.client.Dispatch("HWPFrame.HwpObject")

# 한글 문서를 새로 열거나 기존 파일 열기
hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")  # 보안 모듈 등록

# 파일 경로 지정
file_path = r"C:\mjwork\hwp_files\250224 감정보완 2차 (판교밸리 호반써밋)_VER1.hwp"

# 파일 열기
hwp.Open(file_path)

# 내용 확인 (예: 문서의 전체 텍스트 추출)
text = hwp.HAction.GetDefault("AllText", hwp.HParameterSet.HFindReplace.HSet)
hwp.HAction.Execute("AllText", hwp.HParameterSet.HFindReplace.HSet)
all_text = hwp.HParameterSet.HFindReplace.String
print(all_text)

# 작업 후 닫기 (필요시 저장 여부 선택)
# hwp.Quit()  # 완전히 종료할 때 사용