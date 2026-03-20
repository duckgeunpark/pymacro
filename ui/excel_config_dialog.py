"""
엑셀 소스 설정 다이얼로그
"""
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox
from utils.ui_helpers import set_dialog_icon, center_window_on_parent
from core.project_manager import ProjectManager
from ui.theme import Colors, Fonts, Sizes


class ExcelConfigDialog(ctk.CTkToplevel):
    """엑셀 소스 자동 최신 파일 설정 다이얼로그"""

    def __init__(self, parent, excel_source, project_data, project_filepath):
        super().__init__(parent)
        self.parent = parent
        self.excel_source = excel_source
        self.project_data = project_data
        self.project_filepath = project_filepath
        self.modified = False

        self.title("엑셀 설정")
        self.geometry("500x500")
        self.resizable(False, False)

        # 모달 설정
        self.withdraw()
        self.transient(parent)
        self.grab_set()
        self.attributes('-topmost', True)
        set_dialog_icon(self)

        self.setup_ui()

        def _show():
            center_window_on_parent(self, parent)
            self.deiconify()
            self.lift()
            self.focus_force()
        self.after(50, _show)

    def setup_ui(self):
        """UI 구성"""
        # 제목
        title_frame = ctk.CTkFrame(self, fg_color=Colors.BG_CARD, height=Sizes.HEADER_HEIGHT, corner_radius=0)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)

        ctk.CTkLabel(
            title_frame,
            text="엑셀 자동 최신 파일 설정",
            font=Fonts.HEADING,
            text_color=Colors.TEXT_PRIMARY
        ).pack(pady=Sizes.PAD_MD)

        # 메인 컨텐츠
        content = ctk.CTkFrame(self, fg_color="transparent")
        content.pack(fill='both', expand=True, padx=Sizes.PAD_XL, pady=Sizes.PAD_LG)

        # 자동 최신 파일 선택 체크박스
        self.auto_latest_var = tk.BooleanVar(value=self.excel_source.get('auto_latest', False))
        auto_check = ctk.CTkCheckBox(
            content,
            text="자동으로 최신 파일 선택",
            font=Fonts.BODY_BOLD,
            variable=self.auto_latest_var,
            command=self.toggle_auto_mode
        )
        auto_check.pack(anchor='w', pady=(0, Sizes.PAD_MD))

        # 자동 설정 프레임
        self.auto_frame = ctk.CTkFrame(content, fg_color="transparent")
        self.auto_frame.pack(fill='x', pady=(0, Sizes.PAD_LG))

        # 디렉토리
        ctk.CTkLabel(
            self.auto_frame,
            text="검색 디렉토리:",
            font=Fonts.SMALL
        ).grid(row=0, column=0, sticky='w', pady=Sizes.PAD_XS)

        dir_frame = ctk.CTkFrame(self.auto_frame, fg_color="transparent")
        dir_frame.grid(row=0, column=1, sticky='ew', pady=Sizes.PAD_XS, padx=(Sizes.PAD_SM, 0))

        self.directory_var = tk.StringVar(value=self.excel_source.get('auto_directory', ''))
        self.dir_entry = ctk.CTkEntry(
            dir_frame,
            textvariable=self.directory_var,
            font=Fonts.SMALL
        )
        self.dir_entry.pack(side='left', fill='x', expand=True)

        self.dir_browse_btn = ctk.CTkButton(
            dir_frame,
            text="찾기",
            font=Fonts.CAPTION,
            fg_color=Colors.PRIMARY,
            hover_color=Colors.PRIMARY_HOVER,
            text_color=Colors.TEXT_INVERSE,
            width=60,
            command=self.browse_directory
        )
        self.dir_browse_btn.pack(side='left', padx=(Sizes.PAD_XS, 0))

        # 파일명 prefix
        ctk.CTkLabel(
            self.auto_frame,
            text="파일명 prefix:",
            font=Fonts.SMALL
        ).grid(row=1, column=0, sticky='w', pady=Sizes.PAD_XS)

        self.prefix_var = tk.StringVar(value=self.excel_source.get('auto_prefix', 'list'))
        self.prefix_entry = ctk.CTkEntry(
            self.auto_frame,
            textvariable=self.prefix_var,
            font=Fonts.SMALL,
            width=180
        )
        self.prefix_entry.grid(row=1, column=1, sticky='w', pady=Sizes.PAD_XS, padx=(Sizes.PAD_SM, 0))

        ctk.CTkLabel(
            self.auto_frame,
            text="(예: list -> list20250120.xlsx)",
            font=Fonts.TINY,
            text_color=Colors.TEXT_MUTED
        ).grid(row=2, column=1, sticky='w', padx=(Sizes.PAD_SM, 0))

        self.auto_frame.grid_columnconfigure(1, weight=1)

        # 버튼
        btn_frame = ctk.CTkFrame(content, fg_color="transparent")
        btn_frame.pack(pady=(Sizes.PAD_SM, 0))

        ctk.CTkButton(
            btn_frame,
            text="저장",
            font=Fonts.BODY_BOLD,
            fg_color=Colors.SUCCESS,
            hover_color=Colors.SUCCESS_HOVER,
            text_color=Colors.TEXT_INVERSE,
            width=120,
            command=self.save_config
        ).pack(side='left', padx=Sizes.PAD_XS)

        ctk.CTkButton(
            btn_frame,
            text="취소",
            font=Fonts.BODY,
            fg_color=Colors.BG_ELEVATED,
            hover_color=Colors.GRAY_600,
            text_color=Colors.TEXT_PRIMARY,
            width=120,
            command=self.destroy
        ).pack(side='left', padx=Sizes.PAD_XS)

        # 초기 상태 설정
        self.toggle_auto_mode()

    def toggle_auto_mode(self):
        """자동 모드 토글"""
        if self.auto_latest_var.get():
            # 자동 모드 활성화
            self.dir_entry.configure(state='normal')
            self.dir_browse_btn.configure(state='normal')
            self.prefix_entry.configure(state='normal')
        else:
            # 자동 모드 비활성화
            self.dir_entry.configure(state='disabled')
            self.dir_browse_btn.configure(state='disabled')
            self.prefix_entry.configure(state='disabled')

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
