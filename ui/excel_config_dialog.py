"""
엑셀 소스 설정 다이얼로그
"""
import tkinter as tk
from tkinter import filedialog, messagebox
from utils.ui_helpers import set_dialog_icon, center_window_on_parent
from core.project_manager import ProjectManager


class ExcelConfigDialog(tk.Toplevel):
    """엑셀 소스 자동 최신 파일 설정 다이얼로그"""

    def __init__(self, parent, excel_source, project_data, project_filepath):
        super().__init__(parent)
        self.parent = parent
        self.excel_source = excel_source
        self.project_data = project_data
        self.project_filepath = project_filepath
        self.modified = False

        self.title("📊 엑셀 설정")
        self.geometry("500x500")
        self.resizable(False, False)

        # 모달 설정
        self.transient(parent)
        self.grab_set()
        self.attributes('-topmost', True)
        set_dialog_icon(self)

        self.setup_ui()
        center_window_on_parent(self, parent)

    def setup_ui(self):
        """UI 구성"""
        # 제목
        title_frame = tk.Frame(self, bg='#2c3e50', height=60)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)

        tk.Label(
            title_frame,
            text="📊 엑셀 자동 최신 파일 설정",
            font=("맑은 고딕", 14, "bold"),
            bg='#2c3e50',
            fg='white'
        ).pack(pady=15)

        # 메인 컨텐츠
        content = tk.Frame(self, bg='#F0F0F0', padx=30, pady=20)
        content.pack(fill='both', expand=True)

        # 자동 최신 파일 선택 체크박스
        self.auto_latest_var = tk.BooleanVar(value=self.excel_source.get('auto_latest', False))
        auto_check = tk.Checkbutton(
            content,
            text="자동으로 최신 파일 선택",
            font=("맑은 고딕", 11, "bold"),
            bg='#F0F0F0',
            variable=self.auto_latest_var,
            command=self.toggle_auto_mode
        )
        auto_check.pack(anchor='w', pady=(0, 15))

        # 자동 설정 프레임
        self.auto_frame = tk.Frame(content, bg='#F0F0F0')
        self.auto_frame.pack(fill='x', pady=(0, 20))

        # 디렉토리
        tk.Label(
            self.auto_frame,
            text="검색 디렉토리:",
            font=("맑은 고딕", 10),
            bg='#F0F0F0'
        ).grid(row=0, column=0, sticky='w', pady=5)

        dir_frame = tk.Frame(self.auto_frame, bg='#F0F0F0')
        dir_frame.grid(row=0, column=1, sticky='ew', pady=5, padx=(10, 0))

        self.directory_var = tk.StringVar(value=self.excel_source.get('auto_directory', ''))
        dir_entry = tk.Entry(
            dir_frame,
            textvariable=self.directory_var,
            font=("맑은 고딕", 10),
            relief='solid',
            borderwidth=1
        )
        dir_entry.pack(side='left', fill='x', expand=True)

        tk.Button(
            dir_frame,
            text="찾기",
            font=("맑은 고딕", 9),
            bg='#3498db',
            fg='white',
            padx=10,
            command=self.browse_directory
        ).pack(side='left', padx=(5, 0))

        # 파일명 prefix
        tk.Label(
            self.auto_frame,
            text="파일명 prefix:",
            font=("맑은 고딕", 10),
            bg='#F0F0F0'
        ).grid(row=1, column=0, sticky='w', pady=5)

        self.prefix_var = tk.StringVar(value=self.excel_source.get('auto_prefix', 'list'))
        prefix_entry = tk.Entry(
            self.auto_frame,
            textvariable=self.prefix_var,
            font=("맑은 고딕", 10),
            width=20,
            relief='solid',
            borderwidth=1
        )
        prefix_entry.grid(row=1, column=1, sticky='w', pady=5, padx=(10, 0))

        tk.Label(
            self.auto_frame,
            text="(예: list → list20250120.xlsx)",
            font=("맑은 고딕", 8),
            fg='#7f8c8d',
            bg='#F0F0F0'
        ).grid(row=2, column=1, sticky='w', padx=(10, 0))

        self.auto_frame.grid_columnconfigure(1, weight=1)

        # 버튼
        btn_frame = tk.Frame(content, bg='#F0F0F0')
        btn_frame.pack(pady=(10, 0))

        tk.Button(
            btn_frame,
            text="💾 저장",
            font=("맑은 고딕", 11, "bold"),
            bg='#27ae60',
            fg='white',
            padx=30,
            pady=10,
            command=self.save_config
        ).pack(side='left', padx=5)

        tk.Button(
            btn_frame,
            text="취소",
            font=("맑은 고딕", 11),
            bg='#95a5a6',
            fg='white',
            padx=30,
            pady=10,
            command=self.destroy
        ).pack(side='left', padx=5)

        # 초기 상태 설정
        self.toggle_auto_mode()

    def toggle_auto_mode(self):
        """자동 모드 토글"""
        if self.auto_latest_var.get():
            # 자동 모드 활성화
            for child in self.auto_frame.winfo_children():
                if isinstance(child, (tk.Entry, tk.Button)):
                    child.config(state='normal')
                elif isinstance(child, tk.Frame):
                    for subchild in child.winfo_children():
                        if isinstance(subchild, (tk.Entry, tk.Button)):
                            subchild.config(state='normal')
        else:
            # 자동 모드 비활성화
            for child in self.auto_frame.winfo_children():
                if isinstance(child, (tk.Entry, tk.Button)):
                    child.config(state='disabled')
                elif isinstance(child, tk.Frame):
                    for subchild in child.winfo_children():
                        if isinstance(subchild, (tk.Entry, tk.Button)):
                            subchild.config(state='disabled')

    def browse_directory(self):
        """디렉토리 선택"""
        directory = filedialog.askdirectory(
            parent=self,
            title="검색할 디렉토리 선택",
            initialdir=self.directory_var.get() or '.'
        )
        if directory:
            self.directory_var.set(directory)

    def save_config(self):
        """설정 저장"""
        auto_latest = self.auto_latest_var.get()

        # 자동 모드가 활성화되어 있으면 검증
        if auto_latest:
            directory = self.directory_var.get().strip()
            prefix = self.prefix_var.get().strip()

            if not directory:
                messagebox.showwarning("입력 오류", "검색 디렉토리를 입력해주세요.", parent=self)
                return

            if not prefix:
                messagebox.showwarning("입력 오류", "파일명 prefix를 입력해주세요.", parent=self)
                return

        # 엑셀 소스 업데이트
        excel_sources = self.project_data.get('excel_sources', [])
        if excel_sources:
            excel_sources[0]['auto_latest'] = auto_latest
            excel_sources[0]['auto_directory'] = self.directory_var.get().strip() if auto_latest else None
            excel_sources[0]['auto_prefix'] = self.prefix_var.get().strip() if auto_latest else 'list'

        # 프로젝트 파일 저장
        if ProjectManager.save_project(self.project_filepath, self.project_data):
            messagebox.showinfo("완료", "엑셀 설정이 저장되었습니다!", parent=self)
            self.modified = True
            self.destroy()
        else:
            messagebox.showerror("오류", "설정 저장에 실패했습니다.", parent=self)
