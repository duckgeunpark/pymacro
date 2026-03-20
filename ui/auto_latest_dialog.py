"""
자동 최신 파일 선택 다이얼로그
"""
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox
from utils.ui_helpers import set_dialog_icon, center_window_on_parent
from ui.theme import Colors, Fonts, Sizes, Styles
import os
import glob
import re


class AutoLatestFileDialog(ctk.CTkToplevel):
    """자동 최신 파일 선택 설정 다이얼로그"""

    def __init__(self, parent, default_directory="", source=None):
        super().__init__(parent)
        self.title("자동 최신 파일 선택")
        self.geometry("560x800")
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
        outer = ctk.CTkFrame(self, fg_color=Colors.BG_SECONDARY)
        outer.pack(fill='both', expand=True)

        # ── 헤더 카드 ──
        header_card = Styles.card(outer, corner_radius=0, border_width=0)
        header_card.pack(fill='x')

        header_inner = ctk.CTkFrame(header_card, fg_color="transparent")
        header_inner.pack(fill='x', padx=Sizes.PAD_XL, pady=Sizes.PAD_LG)

        ctk.CTkLabel(
            header_inner,
            text="자동 최신 파일 선택",
            font=Fonts.HEADING,
            text_color=Colors.TEXT_PRIMARY,
        ).pack(anchor='w')

        ctk.CTkLabel(
            header_inner,
            text="매크로 실행 시마다 지정된 디렉토리에서 조건에 맞는 최신 파일을 자동으로 로드합니다.",
            font=Fonts.CAPTION,
            text_color=Colors.TEXT_MUTED,
            justify='left',
        ).pack(anchor='w', pady=(Sizes.PAD_XS, 0))

        # ── 스크롤 가능 본문 ──
        body = ctk.CTkFrame(outer, fg_color="transparent")
        body.pack(fill='both', expand=True, padx=Sizes.PAD_LG, pady=(Sizes.PAD_SM, 0))

        # ── 자동 선택 토글 ──
        self.auto_var = tk.BooleanVar(value=False)
        toggle_frame = ctk.CTkFrame(
            body,
            fg_color=Colors.AUTOSTART_BG,
            corner_radius=Sizes.RADIUS_MD,
            border_width=1,
            border_color=Colors.AUTOSTART_BORDER,
        )
        toggle_frame.pack(fill='x', pady=(0, Sizes.PAD_MD))

        self.auto_check = ctk.CTkCheckBox(
            toggle_frame,
            text="자동 최신 파일 선택 활성화",
            variable=self.auto_var,
            font=Fonts.BODY_BOLD,
            text_color=Colors.AUTOSTART_FG,
            command=self.toggle_inputs,
        )
        self.auto_check.pack(anchor='w', padx=Sizes.PAD_MD, pady=Sizes.PAD_MD)

        # ── 설정 카드 1: 디렉토리 & Prefix ──
        Styles.section_title(body, text="파일 검색 설정").pack(anchor='w', pady=(0, Sizes.PAD_XS))

        settings_card = Styles.card(body)
        settings_card.pack(fill='x', pady=(0, Sizes.PAD_MD))

        settings_inner = ctk.CTkFrame(settings_card, fg_color="transparent")
        settings_inner.pack(fill='x', padx=Sizes.PAD_MD, pady=Sizes.PAD_MD)

        # 디렉토리 입력
        ctk.CTkLabel(
            settings_inner,
            text="검색 디렉토리:",
            font=Fonts.CAPTION_BOLD,
            text_color=Colors.TEXT_MUTED,
        ).pack(anchor='w', pady=(0, Sizes.PAD_XS))

        dir_input_frame = ctk.CTkFrame(settings_inner, fg_color="transparent")
        dir_input_frame.pack(fill='x', pady=(0, Sizes.PAD_MD))

        self.dir_entry = Styles.input_field(dir_input_frame, state='disabled')
        self.dir_entry.pack(side='left', fill='x', expand=True, padx=(0, Sizes.PAD_XS))
        if self.auto_directory:
            self.dir_entry.configure(state='normal')
            self.dir_entry.insert(0, self.auto_directory)
            self.dir_entry.configure(state='disabled')

        self.browse_btn = Styles.secondary_button(
            dir_input_frame,
            text="찾아보기",
            command=self.browse_directory,
        )
        self.browse_btn.configure(width=80, state='disabled')
        self.browse_btn.pack(side='right')

        # Prefix 입력
        ctk.CTkLabel(
            settings_inner,
            text="파일명 시작 문자 (Prefix):",
            font=Fonts.CAPTION_BOLD,
            text_color=Colors.TEXT_MUTED,
        ).pack(anchor='w', pady=(0, 3))

        ctk.CTkLabel(
            settings_inner,
            text="비워두면 모든 엑셀 파일 대상, 입력하면 해당 문자로 시작하는 파일만 검색",
            font=Fonts.TINY,
            text_color=Colors.TEXT_MUTED,
        ).pack(anchor='w', pady=(0, Sizes.PAD_XS))

        self.prefix_entry = Styles.input_field(settings_inner, state='disabled')
        self.prefix_entry.pack(fill='x')

        # ── 설정 카드 2: 파일 선택 기준 ──
        Styles.section_title(body, text="파일 선택 기준").pack(anchor='w', pady=(0, Sizes.PAD_XS))

        mode_card = Styles.card(body)
        mode_card.pack(fill='x', pady=(0, Sizes.PAD_MD))

        self.mode_var = tk.StringVar(value="modified")

        # 수정일 기준
        modified_frame = ctk.CTkFrame(mode_card, fg_color="transparent")
        modified_frame.pack(fill='x', padx=Sizes.PAD_MD, pady=(Sizes.PAD_MD, Sizes.PAD_XS))

        self.modified_radio = ctk.CTkRadioButton(
            modified_frame,
            text="파일 수정일 기준 (가장 최근 수정된 파일)",
            variable=self.mode_var,
            value="modified",
            font=Fonts.SMALL,
            state='disabled',
        )
        self.modified_radio.pack(anchor='w')

        ctk.CTkLabel(
            modified_frame,
            text="      -> 파일명과 관계없이 가장 최근에 수정된 파일을 선택",
            font=Fonts.TINY,
            text_color=Colors.TEXT_MUTED,
        ).pack(anchor='w')

        # 순번 기준
        number_frame = ctk.CTkFrame(mode_card, fg_color="transparent")
        number_frame.pack(fill='x', padx=Sizes.PAD_MD, pady=(Sizes.PAD_SM, Sizes.PAD_MD))

        self.number_radio = ctk.CTkRadioButton(
            number_frame,
            text="파일명 숫자 기준 (가장 큰 숫자)",
            variable=self.mode_var,
            value="number",
            font=Fonts.SMALL,
            state='disabled',
        )
        self.number_radio.pack(anchor='w')

        ctk.CTkLabel(
            number_frame,
            text="      -> 파일명에서 숫자를 추출하여 가장 큰 값을 가진 파일 선택\n"
                 "      -> 예: list1.xlsx, list2.xlsx, list3.xlsx -> list3.xlsx 선택\n"
                 "      -> 날짜형식도 자동 인식: 20250120, 250120, (20250120) 등",
            font=Fonts.TINY,
            text_color=Colors.TEXT_MUTED,
            justify='left',
        ).pack(anchor='w')

        # ── 미리보기 ──
        preview_frame = ctk.CTkFrame(body, fg_color="transparent")
        preview_frame.pack(fill='x', pady=(0, Sizes.PAD_SM))

        self.preview_btn = Styles.secondary_button(
            preview_frame,
            text="미리보기",
            command=self.preview_files,
        )
        self.preview_btn.configure(width=100, state='disabled')
        self.preview_btn.pack(side='left')

        self.preview_label = ctk.CTkLabel(
            preview_frame,
            text="",
            font=Fonts.CAPTION,
            text_color=Colors.PRIMARY_LIGHT,
        )
        self.preview_label.pack(side='left', padx=Sizes.PAD_SM)

        # ── 설정 카드 3: 데이터 처리 옵션 ──
        Styles.section_title(body, text="데이터 처리 옵션").pack(anchor='w', pady=(0, Sizes.PAD_XS))

        data_card = Styles.card(body)
        data_card.pack(fill='x', pady=(0, Sizes.PAD_MD))

        data_inner = ctk.CTkFrame(data_card, fg_color="transparent")
        data_inner.pack(fill='x', padx=Sizes.PAD_MD, pady=Sizes.PAD_MD)

        # 공백 행 제거 옵션
        self.remove_empty_var = tk.BooleanVar(value=self.source.get('remove_empty_rows', True) if self.source else True)
        ctk.CTkCheckBox(
            data_inner,
            text="상위 공백 행 제거",
            variable=self.remove_empty_var,
            font=Fonts.SMALL,
        ).pack(anchor='w', pady=(0, 2))

        ctk.CTkLabel(
            data_inner,
            text="      -> 데이터 시작 전 빈 행들을 자동으로 제거",
            font=Fonts.TINY,
            text_color=Colors.TEXT_MUTED,
        ).pack(anchor='w', pady=(0, Sizes.PAD_SM))

        # 중복 행 제거 옵션
        self.remove_dup_var = tk.BooleanVar(value=self.source.get('remove_duplicates', True) if self.source else True)
        ctk.CTkCheckBox(
            data_inner,
            text="중복 행 제거",
            variable=self.remove_dup_var,
            font=Fonts.SMALL,
        ).pack(anchor='w', pady=(0, 2))

        ctk.CTkLabel(
            data_inner,
            text="      -> 선택한 컬럼 기준으로 중복된 행 제거",
            font=Fonts.TINY,
            text_color=Colors.TEXT_MUTED,
        ).pack(anchor='w')

        # ── 하단 버튼 ──
        btn_frame = ctk.CTkFrame(outer, fg_color="transparent")
        btn_frame.pack(fill='x', padx=Sizes.PAD_LG, pady=(0, Sizes.PAD_LG))

        Styles.secondary_button(
            btn_frame,
            text="취소",
            command=self.on_cancel,
        ).configure(width=100)
        btn_frame.winfo_children()[-1].pack(side='right', padx=(Sizes.PAD_XS, 0))

        ctk.CTkButton(
            btn_frame,
            text="확인",
            font=Fonts.BODY_BOLD,
            fg_color=Colors.SUCCESS,
            hover_color=Colors.SUCCESS_HOVER,
            text_color=Colors.TEXT_INVERSE,
            height=Sizes.BTN_HEIGHT_LG,
            corner_radius=Sizes.RADIUS_MD,
            width=100,
            command=self.on_ok,
        ).pack(side='right', padx=(0, Sizes.PAD_XS))

    def toggle_inputs(self):
        """입력 필드 활성화/비활성화"""
        if self.auto_var.get():
            self.dir_entry.configure(state='normal')
            self.browse_btn.configure(state='normal')
            self.prefix_entry.configure(state='normal')
            self.modified_radio.configure(state='normal')
            self.number_radio.configure(state='normal')
            self.preview_btn.configure(state='normal')
        else:
            self.dir_entry.configure(state='disabled')
            self.browse_btn.configure(state='disabled')
            self.prefix_entry.configure(state='disabled')
            self.modified_radio.configure(state='disabled')
            self.number_radio.configure(state='disabled')
            self.preview_btn.configure(state='disabled')
            self.preview_label.configure(text="")

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
            self.preview_label.configure(text="디렉토리를 입력하세요", text_color=Colors.DANGER)
            return

        latest_file, error = self.find_latest_file(directory, prefix, mode)

        if error:
            self.preview_label.configure(text=f"{error}", text_color=Colors.DANGER)
        else:
            filename = os.path.basename(latest_file)
            self.preview_label.configure(text=f"선택될 파일: {filename}", text_color=Colors.SUCCESS)

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
