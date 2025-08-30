"""
파일 선택 UI 모듈
사용자가 파일과 폴더를 직접 선택할 수 있게 해주는 간단한 GUI

왜 이 기능이 필요한가?
- 하드코딩된 경로 대신 사용자가 원하는 파일을 선택할 수 있게 하기 위해
- 다양한 환경에서 유연하게 프로그램을 사용할 수 있게 하기 위해
- 사용자 친화적인 인터페이스를 제공하기 위해
"""

# 왜 이 라이브러리들을 import하는가?
try:
    import tkinter as tk                    # GUI 기본 창 만들기
    from tkinter import filedialog         # 파일 선택 대화상자
    from tkinter import messagebox         # 메시지 박스 (알림창)
    from tkinter import ttk                # 더 예쁜 GUI 컴포넌트들
    TKINTER_AVAILABLE = True
    
    # 새로운 승인 UI import (선택적)
    try:
        from .user_confirmation import show_worksheet_confirmation
        NEW_UI_AVAILABLE = True
    except ImportError:
        try:
            from user_confirmation import show_worksheet_confirmation
            NEW_UI_AVAILABLE = True
        except ImportError:
            NEW_UI_AVAILABLE = False
            print("⚠️ 새로운 승인 UI를 사용할 수 없습니다. 기본 UI를 사용합니다.")
except ImportError:
    # 왜 try-except로 감싸는가?
    # tkinter가 설치되지 않은 환경에서도 프로그램이 멈추지 않게 하기 위해
    print("⚠️ tkinter를 사용할 수 없습니다. 콘솔 모드로 실행됩니다.")
    TKINTER_AVAILABLE = False
    NEW_UI_AVAILABLE = False

import os
import openpyxl


def select_previous_file():
    """
    전기 조서 파일 선택
    
    이 함수가 하는 일:
    1. 파일 선택 대화상자를 열기
    2. 사용자가 Excel 파일을 선택하게 하기  
    3. 선택된 파일 경로를 반환하기
    
    Returns:
        str: 선택된 파일의 경로, 취소하면 None
    """
    
    # tkinter가 없으면 콘솔에서 직접 입력받기
    if not TKINTER_AVAILABLE:
        print("\n📄 전기 조서 파일을 선택해주세요:")
        file_path = input("파일 경로를 입력하세요 (또는 Enter로 기본값 사용): ").strip()
        
        # 입력하지 않으면 None 반환 (사용자가 직접 선택하게 함)
        if not file_path:
            print("⚠️ 전기 조서 파일을 직접 선택해주세요.")
            return None
        
        # 파일 존재 확인
        if os.path.exists(file_path):
            return file_path
        else:
            print(f"⚠️ 파일을 찾을 수 없습니다: {file_path}")
            return None
    
    # GUI 모드: tkinter 파일 선택 대화상자
    # 왜 Tk()를 만들고 withdraw()하는가?
    # tkinter는 메인 창이 있어야 대화상자를 열 수 있음
    # 하지만 메인 창은 보이지 않게 숨김
    root = tk.Tk()
    root.withdraw()  # 메인 창 숨기기
    
    # 왜 이런 옵션들을 설정하는가?
    file_path = filedialog.askopenfilename(
        title="전기 조서 파일을 선택하세요",           # 대화상자 제목
        filetypes=[                                   # 선택 가능한 파일 형식
            ("Excel 파일", "*.xlsx *.xls"),          # Excel 파일만 보이게
            ("모든 파일", "*.*")                     # 필요시 모든 파일도 볼 수 있게
        ],
        initialdir="."                               # 현재 폴더에서 시작
    )
    
    root.destroy()  # 숨겨진 메인 창 완전히 제거 (메모리 정리)
    
    # 사용자가 취소를 눌렀으면 빈 문자열이 반환됨
    if not file_path:
        return None
    
    return file_path

def select_current_folder():
    """
    당기 PBC 폴더 선택
    
    이 함수가 하는 일:
    1. 폴더 선택 대화상자를 열기
    2. 사용자가 당기 워크페이퍼들이 들어있는 폴더를 선택하게 하기
    3. 선택된 폴더 경로를 반환하기
    
    Returns:
        str: 선택된 폴더의 경로, 취소하면 None
    """
    
    # tkinter가 없으면 콘솔에서 직접 입력받기
    if not TKINTER_AVAILABLE:
        print("\n📁 당기 PBC 폴더를 선택해주세요:")
        folder_path = input("폴더 경로를 입력하세요 (또는 Enter로 기본값 사용): ").strip()
        
        # 입력하지 않으면 None 반환 (사용자가 직접 선택하게 함)
        if not folder_path:
            print("⚠️ 당기 PBC 폴더를 직접 선택해주세요.")
            return None
        
        # 폴더 존재 확인
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            # 왜 끝에 /를 추가하는가?
            # 나중에 파일 경로를 만들 때 일관성을 위해
            if not folder_path.endswith('/'):
                folder_path += '/'
            return folder_path
        else:
            print(f"⚠️ 폴더를 찾을 수 없습니다: {folder_path}")
            return None
    
    # GUI 모드: tkinter 폴더 선택 대화상자
    root = tk.Tk()
    root.withdraw()  # 메인 창 숨기기
    
    folder_path = filedialog.askdirectory(
        title="당기 PBC 폴더를 선택하세요",
        initialdir="."
    )
    
    root.destroy()  # 메인 창 제거
    
    if not folder_path:
        return None
    
    # 왜 끝에 /를 추가하는가?
    # 나중에 folder_path + file_name으로 경로를 만들 때
    # 일관된 형태를 유지하기 위해
    if not folder_path.endswith('/'):
        folder_path += '/'
    
    return folder_path

def get_excel_files_in_folder(folder_path):
    """
    폴더에서 Excel 파일들 찾기
    
    이 함수가 하는 일:
    1. 지정된 폴더를 스캔하기
    2. .xlsx, .xls 확장자를 가진 파일들만 찾기
    3. 찾은 파일들의 리스트 반환하기
    
    Parameters:
        folder_path (str): 스캔할 폴더 경로
        
    Returns:
        list: Excel 파일들의 전체 경로 리스트
    """
    
    # 폴더 존재 확인
    if not os.path.exists(folder_path):
        print(f"⚠️ 폴더를 찾을 수 없습니다: {folder_path}")
        return []
    
    excel_files = []  # 찾은 Excel 파일들을 저장할 리스트
    
    try:
        # 왜 os.listdir을 사용하는가?
        # 폴더 내의 모든 파일과 하위폴더 목록을 가져오기 위해
        for item in os.listdir(folder_path):
            # 왜 lower()를 사용하는가?
            # 파일 확장자가 .XLSX, .XLS 등 대문자일 수도 있기 때문
            # 소문자로 변환해서 비교하면 더 안전함
            
            # 왜 ~$로 시작하는 파일을 제외하는가?
            # Excel을 열면 ~$로 시작하는 임시 파일이 자동 생성됨 (예: ~$workbook1.xlsx)
            # 이런 임시 파일들은 실제 데이터 파일이 아니므로 처리에서 제외해야 함
            # "File is not a zip file" 오류의 주요 원인이기도 함
            if item.lower().endswith(('.xlsx', '.xls')) and not item.startswith('~$'):
                # 전체 경로로 저장 (파일명만이 아님)
                full_path = os.path.join(folder_path, item)
                excel_files.append(full_path)
        
        # 왜 정렬하는가?
        # 파일 목록을 알파벳 순으로 정리해서 사용자가 찾기 쉽게 하기 위해
        excel_files.sort()
        
    except Exception as e:
        print(f"❌ 폴더 스캔 실패: {e}")
        return []
    
    return excel_files

def show_selection_summary(previous_file, current_folder, excel_files):
    """
    사용자가 선택한 파일들의 요약 정보 보여주기
    
    왜 이 함수가 필요한가?
    - 사용자가 올바른 파일을 선택했는지 확인할 수 있게 하기 위해
    - 실제 작업을 시작하기 전에 마지막으로 점검할 기회를 제공하기 위해
    """
    
    print("\n" + "="*60)
    print("📋 선택 사항 요약")
    print("="*60)
    print(f"📄 전기 조서 파일: {previous_file}")
    print(f"📁 당기 PBC 폴더: {current_folder}")
    print(f"📊 처리할 Excel 파일 수: {len(excel_files)}개")
    
    if excel_files:
        print("📝 처리할 파일 목록:")
        for i, file in enumerate(excel_files, 1):
            print(f"   {i}. {file}")
    else:
        print("⚠️ 당기 폴더에 Excel 파일이 없습니다!")
    
    print("="*60)

def confirm_selection():
    """
    사용자에게 계속 진행할지 확인하기
    
    Returns:
        bool: 계속 진행하면 True, 취소하면 False
    """
    
    if not TKINTER_AVAILABLE:
        # 콘솔 모드: 직접 입력받기
        while True:
            answer = input("\n계속 진행하시겠습니까? (y/n): ").strip().lower()
            if answer in ['y', 'yes', '예', 'ㅇ']:
                return True
            elif answer in ['n', 'no', '아니오', 'ㄴ']:
                return False
            else:
                print("y 또는 n으로 답해주세요.")
    
    # GUI 모드: 메시지박스 사용
    root = tk.Tk()
    root.withdraw()
    
    # askquestion은 'yes' 또는 'no' 문자열을 반환
    result = messagebox.askquestion(
        "확인", 
        "선택한 파일들로 롤포워딩을 시작하시겠습니까?",
        icon='question'
    )
    
    root.destroy()
    
    return result == 'yes'

def test_file_selector():
    """
    파일 선택기 테스트 함수
    
    왜 테스트 함수를 만드는가?
    - 파일 선택 기능이 제대로 작동하는지 확인하기 위해
    - GUI와 콘솔 모드 둘 다 정상적으로 동작하는지 검증하기 위해
    """
    
    print("🧪 파일 선택기 테스트 시작...")
    print(f"GUI 모드 사용 가능: {TKINTER_AVAILABLE}")
    
    # 1단계: 전기 조서 파일 선택 테스트
    print("\n1. 전기 조서 파일 선택 테스트")
    previous_file = select_previous_file()
    if previous_file:
        print(f"✅ 선택된 파일: {previous_file}")
    else:
        print("❌ 파일 선택이 취소되었습니다.")
        return
    
    # 2단계: 당기 폴더 선택 테스트
    print("\n2. 당기 폴더 선택 테스트")
    current_folder = select_current_folder()
    if current_folder:
        print(f"✅ 선택된 폴더: {current_folder}")
    else:
        print("❌ 폴더 선택이 취소되었습니다.")
        return
    
    # 3단계: 폴더 내 Excel 파일 찾기 테스트
    print("\n3. Excel 파일 검색 테스트")
    excel_files = get_excel_files_in_folder(current_folder)
    print(f"✅ 찾은 파일 수: {len(excel_files)}개")
    
    # 4단계: 요약 정보 표시 테스트
    show_selection_summary(previous_file, current_folder, excel_files)
    
    # 5단계: 확인 대화상자 테스트
    if confirm_selection():
        print("✅ 사용자가 계속 진행을 선택했습니다.")
    else:
        print("❌ 사용자가 취소를 선택했습니다.")
    
    print("🎉 파일 선택기 테스트 완료!")

def get_worksheet_names(file_path):
    """
    Excel 파일에서 모든 워크시트 이름 추출
    
    이 함수가 하는 일:
    1. Excel 파일을 열어서 모든 워크시트 이름을 가져오기
    2. 사용자가 본 조서를 선택할 수 있도록 목록 제공
    
    Parameters:
        file_path (str): Excel 파일 경로
        
    Returns:
        list: 워크시트 이름 리스트
    """
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True)
        worksheet_names = wb.sheetnames
        wb.close()
        return worksheet_names
    except Exception as e:
        print(f"[file_selector.get_worksheet_names] ❌ 워크시트 목록 추출 실패: {e}")
        return []

def select_main_worksheets(file_path):
    """
    본 조서에 해당하는 워크시트들을 다중 선택 (지능형 감지 시스템)
    
    이 함수가 하는 일:
    1. 지능형 워크시트 자동 감지 실행
    2. 직관적이고 접근성 좋은 확인 UI 표시
    3. 사용자 확인 및 수정 가능
    4. 화이트리스트 저장 옵션 제공
    
    Parameters:
        file_path (str): 전기 조서 파일 경로
        
    Returns:
        tuple: (선택된 본 조서 워크시트 리스트, 백데이터 워크시트 리스트)
    """
    
    if not file_path or not os.path.exists(file_path):
        print(f"[file_selector.select_main_worksheets] ❌ 파일이 존재하지 않습니다: {file_path}")
        return [], []
    
    try:
        # 새로운 지능형 승인 인터페이스 사용
        from pathlib import Path
        
        print(f"\n🔍 지능형 워크시트 자동 감지 시작...")
        print(f"📋 파일: {Path(file_path).name}")
        
        if TKINTER_AVAILABLE and NEW_UI_AVAILABLE:
            # 새로운 직관적 승인 인터페이스 사용
            main_worksheets, back_data_worksheets = show_worksheet_confirmation(Path(file_path))
            
            print(f"\n✅ 워크시트 분류 완료:")
            print(f"   📄 본 조서: {len(main_worksheets)}개 - {main_worksheets}")  
            print(f"   🔴 백데이터: {len(back_data_worksheets)}개 - {back_data_worksheets}")
            
            return main_worksheets, back_data_worksheets
        else:
            # tkinter가 없으면 기존 콘솔 방식 사용
            print("⚠️ GUI를 사용할 수 없어 콘솔 모드로 실행합니다.")
            worksheet_names = get_worksheet_names(file_path)
            return _select_worksheets_console(worksheet_names)
            
    except Exception as e:
        print(f"[file_selector.select_main_worksheets] ⚠️ 새로운 인터페이스 실행 실패: {e}")
        print("기존 방식으로 fallback합니다...")
        
        # Fallback: 기존 방식
        worksheet_names = get_worksheet_names(file_path)
        if not worksheet_names:
            print(f"[file_selector.select_main_worksheets] ❌ 워크시트를 찾을 수 없습니다")
            return [], []
        
        if TKINTER_AVAILABLE:
            return _select_worksheets_gui(worksheet_names)
        else:
            return _select_worksheets_console(worksheet_names)

def _select_worksheets_gui(worksheet_names):
    """GUI 모드로 워크시트 선택"""
    
    root = tk.Tk()
    root.title("본 조서 워크시트 선택")
    root.geometry("500x400")
    
    selected_worksheets = []
    
    # 안내 문구
    label = tk.Label(root, text="본 조서에 해당하는 워크시트를 선택하세요.\n(선택되지 않은 워크시트는 백데이터로 분류됩니다)", 
                    font=("맑은 고딕", 10), justify="left")
    label.pack(pady=10)
    
    # 체크박스 프레임
    checkbox_frame = tk.Frame(root)
    checkbox_frame.pack(pady=10, fill="both", expand=True)
    
    # 스크롤바 추가
    canvas = tk.Canvas(checkbox_frame)
    scrollbar = ttk.Scrollbar(checkbox_frame, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)
    
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    
    # 체크박스 변수들
    checkbox_vars = []
    
    for name in worksheet_names:
        var = tk.BooleanVar()
        checkbox_vars.append(var)
        
        checkbox = tk.Checkbutton(scrollable_frame, text=name, variable=var, 
                                font=("맑은 고딕", 9))
        checkbox.pack(anchor="w", padx=10, pady=2)
    
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    
    # 버튼 프레임
    button_frame = tk.Frame(root)
    button_frame.pack(pady=10)
    
    def on_confirm():
        nonlocal selected_worksheets
        selected_worksheets = [name for name, var in zip(worksheet_names, checkbox_vars) if var.get()]
        root.quit()
    
    def on_cancel():
        nonlocal selected_worksheets
        selected_worksheets = None
        root.quit()
    
    confirm_btn = tk.Button(button_frame, text="확인", command=on_confirm, 
                          font=("맑은 고딕", 10), bg="#4CAF50", fg="white")
    confirm_btn.pack(side="left", padx=5)
    
    cancel_btn = tk.Button(button_frame, text="취소", command=on_cancel, 
                         font=("맑은 고딕", 10), bg="#f44336", fg="white")
    cancel_btn.pack(side="left", padx=5)
    
    # 창을 중앙에 배치
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()
    root.destroy()
    
    if selected_worksheets is None:  # 취소된 경우
        return [], []
    
    # 본 조서와 백데이터 분류
    main_worksheets = selected_worksheets
    back_data_worksheets = [name for name in worksheet_names if name not in main_worksheets]
    
    print(f"\n✅ 본 조서 워크시트 ({len(main_worksheets)}개):")
    for name in main_worksheets:
        print(f"   📄 {name}")
    
    print(f"\n📊 백데이터 워크시트 ({len(back_data_worksheets)}개):")
    for name in back_data_worksheets:
        print(f"   📊 {name}")
    
    return main_worksheets, back_data_worksheets

def _select_worksheets_console(worksheet_names):
    """콘솔 모드로 워크시트 선택"""
    
    print(f"\n본 조서에 해당하는 워크시트 번호를 선택하세요 (쉼표로 구분, 예: 1,3,5):")
    print("선택되지 않은 워크시트는 백데이터로 분류됩니다.")
    
    while True:
        try:
            user_input = input("선택할 워크시트 번호들: ").strip()
            if not user_input:
                print("⚠️ 번호를 입력해주세요.")
                continue
            
            selected_indices = [int(x.strip()) - 1 for x in user_input.split(',')]
            
            # 유효한 범위 확인
            if all(0 <= i < len(worksheet_names) for i in selected_indices):
                main_worksheets = [worksheet_names[i] for i in selected_indices]
                back_data_worksheets = [name for i, name in enumerate(worksheet_names) if i not in selected_indices]
                
                print(f"\n✅ 본 조서 워크시트 ({len(main_worksheets)}개):")
                for name in main_worksheets:
                    print(f"   📄 {name}")
                
                print(f"\n📊 백데이터 워크시트 ({len(back_data_worksheets)}개):")
                for name in back_data_worksheets:
                    print(f"   📊 {name}")
                
                return main_worksheets, back_data_worksheets
            else:
                print(f"❌ 잘못된 번호입니다. 1-{len(worksheet_names)} 범위에서 선택하세요.")
                
        except ValueError:
            print("❌ 숫자만 입력해주세요 (예: 1,3,5)")
        except Exception as e:
            print(f"❌ 입력 오류: {e}")

if __name__ == "__main__":
    # 이 파일을 직접 실행했을 때만 테스트 실행
    test_file_selector()