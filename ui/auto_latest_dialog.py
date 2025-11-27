"""
자동 최신 파일 선택 다이얼로그
"""
import tkinter as tk
from tkinter import filedialog, messagebox
from utils.ui_helpers import set_dialog_icon, center_window_on_parent
import os
import glob
import re


class AutoLatestFileDialog(tk.Toplevel):
    """자동 최신 파일 선택 설정 다이얼로그"""

    def __init__(self, parent, default_directory="", source=None):
        super().__init__(parent)
        self.title("자동 최신 파일 선택")
        self.geometry("550x780")
        self.resizable(False, False)

        self.source = source  # 엑셀 소스 전체 정보
        self.auto_latest = False
        self.auto_directory = default_directory
        self.auto_prefix = ""
        self.auto_mode = "modified"  # 'modified' (수정일) or 'number' (순번)
        self.remove_empty_rows = True
        self.remove_duplicates = False
        self.cancelled = False

        # 모달 설정
        self.transient(parent)
        self.grab_set()
        self.attributes('-topmost', True)
        set_dialog_icon(self)

        self.setup_ui()
        center_window_on_parent(self, parent)

        self.lift()
        self.focus_force()

    def setup_ui(self):
        """UI 구성"""
        main_frame = tk.Frame(self, padx=20, pady=20)
        main_frame.pack(fill='both', expand=True)

        # 제목 및 설명
        tk.Label(
            main_frame,
            text="자동 최신 파일 선택",
            font=("맑은 고딕", 14, "bold")
        ).pack(anchor='w', pady=(0, 10))

        tk.Label(
            main_frame,
            text="매크로 실행 시마다 지정된 디렉토리에서 조건에 맞는 최신 파일을 자동으로 로드합니다.",
            font=("맑은 고딕", 9),
            fg='#7f8c8d',
            justify='left'
        ).pack(anchor='w', pady=(0, 15))

        # 자동 선택 활성화 체크박스
        self.auto_var = tk.BooleanVar(value=False)
        check_frame = tk.Frame(main_frame, bg='#e8f6e8', padx=10, pady=8)
        check_frame.pack(fill='x', pady=(0, 15))

        self.auto_check = tk.Checkbutton(
            check_frame,
            text="✅ 자동 최신 파일 선택 활성화",
            variable=self.auto_var,
            font=("맑은 고딕", 11, "bold"),
            bg='#e8f6e8',
            command=self.toggle_inputs
        )
        self.auto_check.pack(anchor='w')

        # 디렉토리 입력
        dir_frame = tk.Frame(main_frame)
        dir_frame.pack(fill='x', pady=(0, 10))

        tk.Label(
            dir_frame,
            text="📁 검색 디렉토리:",
            font=("맑은 고딕", 10, "bold")
        ).pack(anchor='w', pady=(0, 5))

        dir_input_frame = tk.Frame(dir_frame)
        dir_input_frame.pack(fill='x')

        self.dir_entry = tk.Entry(
            dir_input_frame,
            font=("맑은 고딕", 10),
            state='disabled'
        )
        self.dir_entry.pack(side='left', fill='x', expand=True, padx=(0, 5))
        self.dir_entry.insert(0, self.auto_directory)

        self.browse_btn = tk.Button(
            dir_input_frame,
            text="찾아보기",
            font=("맑은 고딕", 9),
            padx=10,
            pady=5,
            command=self.browse_directory,
            state='disabled'
        )
        self.browse_btn.pack(side='right')

        # Prefix 입력
        prefix_frame = tk.Frame(main_frame)
        prefix_frame.pack(fill='x', pady=(0, 15))

        tk.Label(
            prefix_frame,
            text="📝 파일명 시작 문자 (Prefix):",
            font=("맑은 고딕", 10, "bold")
        ).pack(anchor='w', pady=(0, 3))

        tk.Label(
            prefix_frame,
            text="비워두면 모든 엑셀 파일 대상, 입력하면 해당 문자로 시작하는 파일만 검색",
            font=("맑은 고딕", 8),
            fg='#7f8c8d'
        ).pack(anchor='w', pady=(0, 5))

        self.prefix_entry = tk.Entry(
            prefix_frame,
            font=("맑은 고딕", 10),
            state='disabled'
        )
        self.prefix_entry.pack(fill='x')

        # 로드 기준 선택
        mode_frame = tk.LabelFrame(
            main_frame,
            text="🎯 파일 선택 기준",
            font=("맑은 고딕", 10, "bold"),
            padx=15,
            pady=10
        )
        mode_frame.pack(fill='x', pady=(0, 15))

        self.mode_var = tk.StringVar(value="modified")

        # 수정일 기준
        modified_frame = tk.Frame(mode_frame)
        modified_frame.pack(fill='x', pady=3)

        self.modified_radio = tk.Radiobutton(
            modified_frame,
            text="📅 파일 수정일 기준 (가장 최근 수정된 파일)",
            variable=self.mode_var,
            value="modified",
            font=("맑은 고딕", 10),
            state='disabled'
        )
        self.modified_radio.pack(anchor='w')

        tk.Label(
            modified_frame,
            text="      → 파일명과 관계없이 가장 최근에 수정된 파일을 선택",
            font=("맑은 고딕", 8),
            fg='#7f8c8d'
        ).pack(anchor='w')

        # 순번 기준
        number_frame = tk.Frame(mode_frame)
        number_frame.pack(fill='x', pady=(8, 3))

        self.number_radio = tk.Radiobutton(
            number_frame,
            text="🔢 파일명 숫자 기준 (가장 큰 숫자)",
            variable=self.mode_var,
            value="number",
            font=("맑은 고딕", 10),
            state='disabled'
        )
        self.number_radio.pack(anchor='w')

        tk.Label(
            number_frame,
            text="      → 파일명에서 숫자를 추출하여 가장 큰 값을 가진 파일 선택\n"
                 "      → 예: list1.xlsx, list2.xlsx, list3.xlsx → list3.xlsx 선택\n"
                 "      → 날짜형식도 자동 인식: 20250120, 250120, (20250120) 등",
            font=("맑은 고딕", 8),
            fg='#7f8c8d',
            justify='left'
        ).pack(anchor='w')

        # 미리보기 버튼
        preview_frame = tk.Frame(main_frame)
        preview_frame.pack(fill='x', pady=(0, 10))

        self.preview_btn = tk.Button(
            preview_frame,
            text="🔍 미리보기",
            font=("맑은 고딕", 9),
            padx=15,
            pady=5,
            command=self.preview_files,
            state='disabled'
        )
        self.preview_btn.pack(side='left')

        self.preview_label = tk.Label(
            preview_frame,
            text="",
            font=("맑은 고딕", 9),
            fg='#2980b9'
        )
        self.preview_label.pack(side='left', padx=10)

        # 데이터 처리 옵션
        data_frame = tk.LabelFrame(
            main_frame,
            text="📊 데이터 처리 옵션",
            font=("맑은 고딕", 10, "bold"),
            padx=15,
            pady=10
        )
        data_frame.pack(fill='x', pady=(0, 15))

        # 공백 행 제거 옵션
        self.remove_empty_var = tk.BooleanVar(value=self.source.get('remove_empty_rows', True) if self.source else True)
        tk.Checkbutton(
            data_frame,
            text="🧹 상위 공백 행 제거",
            variable=self.remove_empty_var,
            font=("맑은 고딕", 9)
        ).pack(anchor='w', pady=2)

        tk.Label(
            data_frame,
            text="      → 데이터 시작 전 빈 행들을 자동으로 제거",
            font=("맑은 고딕", 8),
            fg='#7f8c8d'
        ).pack(anchor='w')

        # 중복 행 제거 옵션
        self.remove_dup_var = tk.BooleanVar(value=self.source.get('remove_duplicates', True) if self.source else True)
        tk.Checkbutton(
            data_frame,
            text="🔄 중복 행 제거",
            variable=self.remove_dup_var,
            font=("맑은 고딕", 9)
        ).pack(anchor='w', pady=2)

        tk.Label(
            data_frame,
            text="      → 선택한 컬럼 기준으로 중복된 행 제거",
            font=("맑은 고딕", 8),
            fg='#7f8c8d'
        ).pack(anchor='w')

        # 버튼
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(fill='x', pady=(10, 0))

        tk.Button(
            btn_frame,
            text="확인",
            font=("맑은 고딕", 10),
            bg='#27ae60',
            fg='white',
            padx=25,
            pady=8,
            command=self.on_ok
        ).pack(side='right', padx=5)

        tk.Button(
            btn_frame,
            text="취소",
            font=("맑은 고딕", 10),
            bg='#95a5a6',
            fg='white',
            padx=25,
            pady=8,
            command=self.on_cancel
        ).pack(side='right', padx=5)

    def toggle_inputs(self):
        """입력 필드 활성화/비활성화"""
        if self.auto_var.get():
            self.dir_entry.config(state='normal')
            self.browse_btn.config(state='normal')
            self.prefix_entry.config(state='normal')
            self.modified_radio.config(state='normal')
            self.number_radio.config(state='normal')
            self.preview_btn.config(state='normal')
        else:
            self.dir_entry.config(state='disabled')
            self.browse_btn.config(state='disabled')
            self.prefix_entry.config(state='disabled')
            self.modified_radio.config(state='disabled')
            self.number_radio.config(state='disabled')
            self.preview_btn.config(state='disabled')
            self.preview_label.config(text="")

    def browse_directory(self):
        """디렉토리 찾아보기"""
        directory = filedialog.askdirectory(
            title="검색할 디렉토리 선택",
            initialdir=self.dir_entry.get() or "C:/"
        )
        if directory:
            self.dir_entry.delete(0, tk.END)
            self.dir_entry.insert(0, directory)

    def find_latest_file(self, directory, prefix, mode):
        """최신 파일 찾기"""
        if not os.path.exists(directory):
            return None, "디렉토리가 존재하지 않습니다"

        # 엑셀 파일 패턴
        if prefix:
            pattern = os.path.join(directory, f"{prefix}*.xlsx")
        else:
            pattern = os.path.join(directory, "*.xlsx")

        files = glob.glob(pattern)

        # xls 파일도 포함
        if prefix:
            pattern_xls = os.path.join(directory, f"{prefix}*.xls")
        else:
            pattern_xls = os.path.join(directory, "*.xls")

        files.extend([f for f in glob.glob(pattern_xls) if not f.endswith('.xlsx')])

        if not files:
            return None, "조건에 맞는 파일이 없습니다"

        if mode == "modified":
            # 수정일 기준
            latest_file = max(files, key=os.path.getmtime)
            return latest_file, None
        else:
            # 순번/날짜 기준 - 파일명에서 숫자 추출
            def extract_number(filepath):
                filename = os.path.basename(filepath)
                # 괄호 안의 숫자, 일반 숫자 등 모든 숫자 추출
                # 패턴: 20250120, 250120, (20250120), [250120], 20250120132234 등
                numbers = re.findall(r'\d+', filename)
                if numbers:
                    # 가장 긴 숫자를 선택 (날짜+시간이 보통 가장 김)
                    return max(int(n) for n in numbers)
                return 0

            try:
                latest_file = max(files, key=extract_number)
                return latest_file, None
            except:
                return None, "숫자를 추출할 수 없습니다"

    def preview_files(self):
        """미리보기"""
        directory = self.dir_entry.get().strip()
        prefix = self.prefix_entry.get().strip()
        mode = self.mode_var.get()

        if not directory:
            self.preview_label.config(text="❌ 디렉토리를 입력하세요", fg='#e74c3c')
            return

        latest_file, error = self.find_latest_file(directory, prefix, mode)

        if error:
            self.preview_label.config(text=f"❌ {error}", fg='#e74c3c')
        else:
            filename = os.path.basename(latest_file)
            self.preview_label.config(text=f"✅ 선택될 파일: {filename}", fg='#27ae60')

    def on_ok(self):
        """확인 버튼"""
        self.auto_latest = self.auto_var.get()

        if self.auto_latest:
            # 디렉토리 검증
            directory = self.dir_entry.get().strip()
            if not directory:
                messagebox.showwarning("경고", "디렉토리를 입력하세요.", parent=self)
                return

            if not os.path.exists(directory):
                messagebox.showerror("오류", f"디렉토리가 존재하지 않습니다:\n{directory}", parent=self)
                return

            # 파일 존재 여부 확인
            prefix = self.prefix_entry.get().strip()
            mode = self.mode_var.get()

            latest_file, error = self.find_latest_file(directory, prefix, mode)
            if error:
                messagebox.showerror("오류", f"조건에 맞는 파일을 찾을 수 없습니다:\n{error}", parent=self)
                return

            self.auto_directory = directory
            self.auto_prefix = prefix
            self.auto_mode = mode

        # 데이터 처리 옵션 저장
        self.remove_empty_rows = self.remove_empty_var.get()
        self.remove_duplicates = self.remove_dup_var.get()

        # source 객체가 있으면 직접 업데이트
        if self.source:
            self.source['remove_empty_rows'] = self.remove_empty_rows
            self.source['remove_duplicates'] = self.remove_duplicates

        self.cancelled = False
        self.destroy()

    def on_cancel(self):
        """취소 버튼"""
        self.cancelled = True
        self.destroy()
