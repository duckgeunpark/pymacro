"""
각종 다이얼로그 (CustomTkinter)
"""

import tkinter as tk
import customtkinter as ctk
from tkinter import messagebox
from datetime import datetime
from ui.theme import Colors, Fonts, Sizes, Styles
from utils.ui_helpers import (
    set_dialog_icon,
    center_window_on_parent,
    center_window_on_screen,
    setup_modal,
)


class NewProjectDialog(ctk.CTkToplevel):
    """새 프로젝트 생성 다이얼로그"""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("새 프로젝트")
        self.geometry("360x200")
        self.resizable(False, False)
        self.configure(fg_color=Colors.BG_SECONDARY)

        self.result = None

        # 모달 설정 - withdraw 먼저 해서 깜빡임 방지
        self.withdraw()
        self.transient(parent)
        self.grab_set()
        self.attributes("-topmost", True)
        set_dialog_icon(self)
        self.setup_ui()

        # 창 중앙 배치
        def _show():
            center_window_on_parent(self, parent)
            self.deiconify()
            self.lift()
            self.focus_force()

        self.after(50, _show)

    def setup_ui(self):
        """UI 구성"""
        # 메인 프레임
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=Sizes.PAD_LG, pady=Sizes.PAD_LG)

        # 섹션 타이틀
        Styles.section_title(main_frame, text="프로젝트 이름").pack(
            anchor="w", pady=(0, Sizes.PAD_SM)
        )

        # 카드 영역
        card = Styles.card(main_frame)
        card.pack(fill="x", pady=(0, Sizes.PAD_MD))

        card_inner = ctk.CTkFrame(card, fg_color="transparent")
        card_inner.pack(fill="x", padx=Sizes.PAD_MD, pady=Sizes.PAD_MD)

        self.name_entry = Styles.input_field(
            card_inner, placeholder_text="프로젝트 이름을 입력하세요"
        )
        self.name_entry.pack(fill="x")
        self.name_entry.focus()

        # 글자 수 제한 (14자)
        self.name_entry.bind("<KeyRelease>", self.on_name_change)

        # 확인 버튼
        Styles.primary_button(
            main_frame,
            text="확인",
            command=self.on_create,
        ).pack(fill="x", pady=(Sizes.PAD_SM, 0))

        # Enter 키 바인딩
        self.name_entry.bind("<Return>", lambda e: self.on_create())

    def on_name_change(self, event=None):
        """이름 입력 시 길이 제한 (14자)"""
        MAX_NAME_LENGTH = 14
        current_text = self.name_entry.get()

        if len(current_text) > MAX_NAME_LENGTH:
            # 14자까지만 유지
            self.name_entry.delete(MAX_NAME_LENGTH, tk.END)

    def on_create(self):
        """생성 버튼 클릭"""
        name = self.name_entry.get().strip()

        if not name:
            messagebox.showerror("오류", "프로젝트 이름을 입력하세요.")
            return

        # 파일명으로 사용할 수 없는 문자 체크
        invalid_chars = ["\\", "/", ":", "*", "?", '"', "<", ">", "|"]
        if any(char in name for char in invalid_chars):
            messagebox.showerror(
                "오류",
                f"프로젝트 이름에 다음 문자를 사용할 수 없습니다:\n{' '.join(invalid_chars)}",
            )
            return

        description = ""

        self.result = {"name": name, "description": description}

        self.destroy()

    def on_cancel(self):
        """취소 버튼 클릭"""
        self.result = None
        self.destroy()


class NameInputDialog(ctk.CTkToplevel):
    """이름 입력 다이얼로그"""

    def __init__(
        self, parent, title="이름 입력", message="이름을 입력하세요:", initial_value=""
    ):
        super().__init__(parent)
        self.title(title)
        self.geometry("360x220")
        self.resizable(False, False)
        self.configure(fg_color=Colors.BG_SECONDARY)

        self.result = None

        # 모달 설정
        self.withdraw()
        self.transient(parent)
        self.grab_set()
        self.attributes("-topmost", True)
        set_dialog_icon(self)

        self.setup_ui(message, initial_value)

        def _show():
            center_window_on_parent(self, self.master)
            self.deiconify()
            self.lift()
            self.focus_force()

        self.after(50, _show)

    def setup_ui(self, message, initial_value):
        """UI 구성"""
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=Sizes.PAD_LG, pady=Sizes.PAD_LG)

        # 메시지
        ctk.CTkLabel(
            main_frame,
            text=message,
            font=Fonts.BODY_BOLD,
            text_color=Colors.TEXT_PRIMARY,
        ).pack(anchor="w", pady=(0, Sizes.PAD_SM))

        # 카드 영역
        card = Styles.card(main_frame)
        card.pack(fill="x", pady=(0, Sizes.PAD_MD))

        card_inner = ctk.CTkFrame(card, fg_color="transparent")
        card_inner.pack(fill="x", padx=Sizes.PAD_MD, pady=Sizes.PAD_MD)

        # 입력 필드
        self.entry = Styles.input_field(
            card_inner, placeholder_text="이름을 입력하세요"
        )
        self.entry.pack(fill="x")

        # 글자 수 제한 (10자)
        self.entry.bind("<KeyRelease>", self.on_name_change)

        if initial_value:
            self.entry.insert(0, initial_value)
            self.entry.select_range(0, tk.END)

        self.entry.focus_set()

        # 확인 버튼
        Styles.primary_button(
            main_frame,
            text="확인",
            command=self.on_ok,
        ).pack(fill="x", pady=(Sizes.PAD_SM, 0))

        # Enter 키로 확인
        self.entry.bind("<Return>", lambda e: self.on_ok())
        self.entry.bind("<Escape>", lambda e: self.on_cancel())

    def on_name_change(self, event=None):
        """이름 입력 시 길이 제한 (한글 6자, 숫자/영문 12자)"""
        current_text = self.entry.get()

        # 한글 글자 수 계산
        korean_count = sum(
            1
            for char in current_text
            if ord("가") <= ord(char) <= ord("힣")
            or ord("ㄱ") <= ord(char) <= ord("ㅣ")
        )
        # 나머지 문자 수 (숫자, 영문 등)
        other_count = len(current_text) - korean_count

        # 한글은 2 비중, 나머지는 1 비중으로 계산
        total_weight = korean_count * 2 + other_count

        # 총 비중이 12를 초과하면 (한글 6자 또는 숫자/영문 12자)
        if total_weight > 12:
            # 마지막 입력 제거
            self.entry.delete(len(current_text) - 1, tk.END)

    def on_ok(self):
        """확인 버튼"""
        value = self.entry.get().strip()
        if not value:
            messagebox.showwarning("경고", "이름을 입력하세요.", parent=self)
            return

        self.result = value
        self.destroy()

    def on_cancel(self):
        """취소 버튼"""
        self.result = None
        self.destroy()


class ImageNameDialog(ctk.CTkToplevel):
    """이미지 이름 및 정확도 입력 다이얼로그"""

    def __init__(
        self, parent, title="이미지 이름", initial_name="", initial_confidence=80
    ):
        super().__init__(parent)
        self.title(title)
        self.geometry("340x230")
        self.resizable(False, False)
        self.configure(fg_color=Colors.BG_SECONDARY)

        self.result = None  # {'name': str, 'confidence': float}

        # 모달 설정
        self.withdraw()
        self.transient(parent)
        self.grab_set()
        self.attributes("-topmost", True)
        set_dialog_icon(self)

        self.setup_ui(initial_name, initial_confidence)

        def _show():
            center_window_on_parent(self, self.master)
            self.deiconify()
            self.lift()
            self.focus_force()

        self.after(50, _show)

    def setup_ui(self, initial_name, initial_confidence):
        """UI 구성"""
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=Sizes.PAD_LG, pady=Sizes.PAD_LG)

        # 카드 영역
        card = Styles.card(main_frame)
        card.pack(fill="x", pady=(0, Sizes.PAD_MD))

        card_inner = ctk.CTkFrame(card, fg_color="transparent")
        card_inner.pack(fill="x", padx=Sizes.PAD_MD, pady=Sizes.PAD_MD)

        # 이름 + 정확도 나란히
        row_frame = ctk.CTkFrame(card_inner, fg_color="transparent")
        row_frame.pack(fill="x", pady=(0, Sizes.PAD_SM))

        # 이름 입력
        name_col = ctk.CTkFrame(row_frame, fg_color="transparent")
        name_col.pack(side="left", fill="x", expand=True, padx=(0, Sizes.PAD_SM))

        Styles.section_title(name_col, text="이름").pack(
            anchor="w", pady=(0, Sizes.PAD_XS)
        )
        self.name_entry = Styles.input_field(name_col, placeholder_text="이미지 이름")
        self.name_entry.pack(fill="x")

        if initial_name:
            self.name_entry.insert(0, initial_name)

        self.name_entry.focus_set()

        # 정확도 입력
        conf_col = ctk.CTkFrame(row_frame, fg_color="transparent")
        conf_col.pack(side="left")

        Styles.section_title(conf_col, text="정확도").pack(
            anchor="w", pady=(0, Sizes.PAD_XS)
        )

        conf_row = ctk.CTkFrame(conf_col, fg_color="transparent")
        conf_row.pack(fill="x")

        self.confidence_var = tk.IntVar(value=initial_confidence)

        self.confidence_entry = Styles.input_field(
            conf_row,
            width=65,
            textvariable=self.confidence_var,
        )
        self.confidence_entry.pack(side="left")

        ctk.CTkLabel(
            conf_row,
            text="%",
            font=Fonts.SMALL,
            text_color=Colors.TEXT_SECONDARY,
        ).pack(side="left", padx=(Sizes.PAD_XS, 0))

        # 힌트
        ctk.CTkLabel(
            card_inner,
            text="정확도 범위: 50 ~ 100",
            font=Fonts.TINY,
            text_color=Colors.TEXT_MUTED,
            anchor="w",
        ).pack(anchor="w")

        # 버튼
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x")

        Styles.primary_button(
            btn_frame,
            text="확인",
            command=self.on_ok,
        ).pack(side="left", expand=True, fill="x", padx=(0, Sizes.PAD_XS))

        Styles.secondary_button(
            btn_frame,
            text="취소",
            command=self.on_cancel,
        ).pack(side="left", expand=True, fill="x", padx=(Sizes.PAD_XS, 0))

        # Enter 키로 확인
        self.name_entry.bind("<Return>", lambda e: self.on_ok())
        self.bind("<Escape>", lambda e: self.on_cancel())

    def on_ok(self):
        """확인 버튼"""
        name = self.name_entry.get().strip()
        if not name:
            messagebox.showwarning("경고", "이미지 이름을 입력하세요.", parent=self)
            return

        confidence = self.confidence_var.get()

        self.result = {
            "name": name,
            "confidence": confidence / 100.0,  # 0.5 ~ 1.0으로 변환
        }
        self.destroy()

    def on_cancel(self):
        """취소 버튼"""
        self.result = None
        self.destroy()


class KeyInputDialog(ctk.CTkToplevel):
    """키 입력 감지 다이얼로그"""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("키 입력")
        self.geometry("340x260")
        self.resizable(False, False)
        self.configure(fg_color=Colors.BG_SECONDARY)

        self.result = None
        self.captured_key = None
        self.is_first_key = True  # 첫 키 입력 플래그

        # 모달 설정
        self.withdraw()
        self.transient(parent)
        self.grab_set()
        self.attributes("-topmost", True)
        set_dialog_icon(self)

        self.setup_ui()

        def _show():
            self.center_window()
            self.deiconify()
            self.lift()
            self.focus_force()

        self.after(50, _show)

        # Alt 키 상태 초기화 (추가)
        self.after(100, self.reset_key_state)

    def reset_key_state(self):
        """키 상태 초기화"""
        try:
            # 포커스 재설정으로 Alt 상태 초기화
            self.focus_set()
        except:
            pass

    def center_window(self):
        """창을 부모 중앙에 배치"""
        self.update_idletasks()

        parent_x = self.master.winfo_x()
        parent_y = self.master.winfo_y()
        parent_width = self.master.winfo_width()
        parent_height = self.master.winfo_height()

        dialog_width = self.winfo_width()
        dialog_height = self.winfo_height()

        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2

        self.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")

    def setup_ui(self):
        """UI 구성"""
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=Sizes.PAD_LG, pady=Sizes.PAD_LG)

        # 설명
        ctk.CTkLabel(
            main_frame,
            text="입력할 키를 누르세요",
            font=Fonts.SUBHEADING,
            text_color=Colors.TEXT_PRIMARY,
        ).pack(pady=(0, Sizes.PAD_MD))

        # 키 표시 카드
        key_card = Styles.card(main_frame)
        key_card.pack(fill="x", pady=(0, Sizes.PAD_SM))

        self.key_display = ctk.CTkLabel(
            key_card,
            text="(키를 누르세요...)",
            font=Fonts.HEADING,
            text_color=Colors.TEXT_MUTED,
            fg_color=Colors.BG_INPUT,
            corner_radius=Sizes.RADIUS_SM,
            height=70,
        )
        self.key_display.pack(fill="x", padx=Sizes.PAD_SM, pady=Sizes.PAD_SM)

        # 도움말
        ctk.CTkLabel(
            main_frame,
            text="예: enter, tab, esc, space, f1, f11 등",
            font=Fonts.CAPTION,
            text_color=Colors.TEXT_MUTED,
        ).pack(anchor="w", pady=(0, Sizes.PAD_MD))

        # 확인 버튼
        self.confirm_btn = Styles.primary_button(
            main_frame,
            text="확인",
            command=self.on_ok,
        )
        self.confirm_btn.configure(state="disabled")
        self.confirm_btn.pack(fill="x")

        # 키 이벤트 바인드
        self.bind("<KeyPress>", self.on_key_press)
        self.bind("<KeyRelease>", self.on_key_release)  # 추가
        self.focus_force()

    def on_key_release(self, event):
        """키 해제 감지 (Alt 상태 초기화용)"""
        # Alt 키만 눌렸을 때는 무시
        if event.keysym in ["alt_l", "alt_r"]:
            return

    def on_key_press(self, event):
        """키 입력 감지"""
        # 키 이름 매핑
        key_mapping = {
            "Return": "enter",
            "Tab": "tab",
            "Escape": "esc",
            "space": "space",
            "BackSpace": "backspace",
            "Delete": "delete",
            "Up": "up",
            "Down": "down",
            "Left": "left",
            "Right": "right",
            "Home": "home",
            "End": "end",
            "Prior": "pageup",
            "Next": "pagedown",
            "Insert": "insert",
            "Pause": "pause",
            "Print": "print",
        }

        # 키 이름 가져오기
        key_name = event.keysym.lower()

        # 매핑된 키 확인
        if event.keysym in key_mapping:
            key_name = key_mapping[event.keysym]

        # F1~F12 키 처리
        if event.keysym.startswith("F") and event.keysym[1:].isdigit():
            key_name = event.keysym.lower()

        # 특수 조합 무시 (Shift, Control, Alt만 눌렀을 때) - 수정
        if key_name in [
            "shift_l",
            "shift_r",
            "control_l",
            "control_r",
            "alt_l",
            "alt_r",
            "meta_l",
            "meta_r",
        ]:
            return

        # 수정자 키 (Ctrl, Alt, Shift) 확인 - 수정
        modifiers = []

        # Alt 키 상태만 확인 (Grab으로 인한 이상 제거)
        if event.state & 0x0004:  # Ctrl
            modifiers.append("ctrl")
        # Alt 무시 (Alt 메뉴 접근 방지)
        # if event.state & 0x0008:  # Alt
        #     modifiers.append('alt')
        if event.state & 0x0001:  # Shift
            modifiers.append("shift")

        # 최종 키 조합 생성
        if modifiers:
            self.captured_key = "+".join(modifiers + [key_name])
        else:
            self.captured_key = key_name

        # 디스플레이 업데이트
        self.key_display.configure(text=self.captured_key, text_color=Colors.SUCCESS)

        # 확인 버튼 활성화
        self.confirm_btn.configure(state="normal")

        # 첫 키 입력 후 플래그 해제
        self.is_first_key = False

    def on_ok(self):
        """확인 버튼"""
        if self.captured_key:
            self.result = self.captured_key
            self.destroy()

    def on_cancel(self):
        """취소 버튼"""
        self.result = None
        self.destroy()


class ActionSelectDialog(ctk.CTkToplevel):
    """액션 선택 다이얼로그"""

    def _show_error_dialog(self, message):
        dialog = ctk.CTkToplevel(self)
        dialog.title("오류")
        dialog.geometry("340x170")
        dialog.resizable(False, False)
        dialog.configure(fg_color=Colors.BG_SECONDARY)
        dialog.withdraw()
        dialog.transient(self)
        dialog.grab_set()
        dialog.attributes("-topmost", True)
        set_dialog_icon(dialog)

        inner = ctk.CTkFrame(dialog, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=Sizes.PAD_LG, pady=Sizes.PAD_LG)

        ctk.CTkLabel(
            inner,
            text=message,
            font=Fonts.BODY,
            text_color=Colors.TEXT_PRIMARY,
            wraplength=280,
        ).pack(pady=(Sizes.PAD_MD, Sizes.PAD_LG))

        Styles.primary_button(
            inner,
            text="확인",
            command=lambda: dialog.destroy(),
        ).pack(fill="x")

        def _show():
            self._center_sub_dialog(dialog)
            dialog.deiconify()
            dialog.lift()
            dialog.focus_force()

        dialog.after(50, _show)

        self.wait_window(dialog)

    def _center_sub_dialog(self, dialog):
        """서브 다이얼로그를 부모 중앙에 배치"""
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        parent_x = self.winfo_x()
        parent_y = self.winfo_y()
        parent_width = self.winfo_width()
        parent_height = self.winfo_height()
        x = parent_x + (parent_width - width) // 2
        y = parent_y + (parent_height - height) // 2
        dialog.geometry(f"{width}x{height}+{x}+{y}")

    def __init__(self, parent, coord_mgr, excel_mgr, image_mgr):
        super().__init__(parent)
        self.title("액션 추가")
        self.geometry("440x560")
        self.resizable(False, False)
        self.configure(fg_color=Colors.BG_SECONDARY)

        self.coord_mgr = coord_mgr
        self.excel_mgr = excel_mgr
        self.image_mgr = image_mgr

        self.result = None

        self.withdraw()
        self.transient(parent)
        self.grab_set()
        self.attributes("-topmost", True)
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
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=Sizes.PAD_LG, pady=Sizes.PAD_LG)

        ctk.CTkLabel(
            main_frame,
            text="액션 유형 선택",
            font=Fonts.SUBHEADING,
            text_color=Colors.TEXT_PRIMARY,
        ).pack(pady=(0, Sizes.PAD_LG))

        # 액션 버튼들
        self.create_action_buttons(main_frame)

    def _create_action_btn(self, parent, text, color, hover_color, command, row, col):
        """개별 액션 버튼 생성 헬퍼"""
        btn = ctk.CTkButton(
            parent,
            text=text,
            font=Fonts.SMALL_BOLD,
            width=120,
            height=Sizes.BTN_HEIGHT,
            fg_color=color,
            hover_color=hover_color,
            text_color=Colors.TEXT_INVERSE,
            corner_radius=Sizes.RADIUS_SM,
            command=command,
        )
        btn.grid(row=row, column=col, padx=Sizes.PAD_XS, pady=Sizes.PAD_XS)
        return btn

    def create_action_buttons(self, parent):
        """액션 버튼 생성"""
        # --- 마우스 동작 ---
        Styles.section_title(parent, text="마우스 동작").pack(
            fill="x", pady=(Sizes.PAD_SM, Sizes.PAD_XS)
        )

        mouse_card = Styles.card(parent)
        mouse_card.pack(fill="x", pady=(0, Sizes.PAD_SM))

        btn_frame = ctk.CTkFrame(mouse_card, fg_color="transparent")
        btn_frame.pack(padx=Sizes.PAD_MD, pady=Sizes.PAD_MD)

        self._create_action_btn(
            btn_frame,
            "좌표 클릭",
            Colors.ACTION_MOUSE,
            "#1d4ed8",
            lambda: self.select_action("click_coord"),
            0,
            0,
        )
        self._create_action_btn(
            btn_frame,
            "이미지 클릭",
            Colors.ACTION_MOUSE,
            "#1d4ed8",
            lambda: self.select_action("click_image"),
            0,
            1,
        )
        self._create_action_btn(
            btn_frame,
            "마우스 스크롤",
            Colors.ACTION_MOUSE,
            "#1d4ed8",
            lambda: self.select_action("mouse_scroll"),
            0,
            2,
        )

        # --- 키보드 동작 ---
        Styles.section_title(parent, text="키보드 동작").pack(
            fill="x", pady=(Sizes.PAD_SM, Sizes.PAD_XS)
        )

        keyboard_card = Styles.card(parent)
        keyboard_card.pack(fill="x", pady=(0, Sizes.PAD_SM))

        btn_frame2 = ctk.CTkFrame(keyboard_card, fg_color="transparent")
        btn_frame2.pack(padx=Sizes.PAD_MD, pady=Sizes.PAD_MD)

        self._create_action_btn(
            btn_frame2,
            "텍스트 타이핑",
            Colors.ACTION_KEYBOARD,
            "#059669",
            lambda: self.select_action("type_text"),
            0,
            0,
        )
        self._create_action_btn(
            btn_frame2,
            "변수 타이핑",
            Colors.ACTION_KEYBOARD,
            "#059669",
            lambda: self.select_action("type_variable"),
            0,
            1,
        )
        self._create_action_btn(
            btn_frame2,
            "키 입력",
            Colors.ACTION_KEYBOARD,
            "#059669",
            lambda: self.select_action("key_press"),
            0,
            2,
        )

        # --- 제어 동작 ---
        Styles.section_title(parent, text="제어 동작").pack(
            fill="x", pady=(Sizes.PAD_SM, Sizes.PAD_XS)
        )

        control_card = Styles.card(parent)
        control_card.pack(fill="x", pady=(0, Sizes.PAD_SM))

        btn_frame3 = ctk.CTkFrame(control_card, fg_color="transparent")
        btn_frame3.pack(padx=Sizes.PAD_MD, pady=Sizes.PAD_MD)

        self._create_action_btn(
            btn_frame3,
            "딜레이",
            Colors.ACTION_CONTROL,
            "#c2410c",
            lambda: self.select_action("delay"),
            0,
            0,
        )
        self._create_action_btn(
            btn_frame3,
            "이미지 대기",
            Colors.ACTION_CONTROL,
            "#c2410c",
            lambda: self.select_action("wait_image"),
            0,
            1,
        )

        # --- 기타 ---
        Styles.section_title(parent, text="기타").pack(
            fill="x", pady=(Sizes.PAD_SM, Sizes.PAD_XS)
        )

        other_card = Styles.card(parent)
        other_card.pack(fill="x", pady=(0, Sizes.PAD_SM))

        btn_frame4 = ctk.CTkFrame(other_card, fg_color="transparent")
        btn_frame4.pack(padx=Sizes.PAD_MD, pady=Sizes.PAD_MD)

        self._create_action_btn(
            btn_frame4,
            "스크린샷",
            Colors.ACTION_OTHER,
            Colors.PURPLE_HOVER,
            lambda: self.select_action("screenshot"),
            0,
            0,
        )

    def select_action(self, action_type):
        """액션 선택"""
        params = None

        if action_type == "click_coord":
            params = self.config_click_coord()
        elif action_type == "click_image":
            params = self.config_click_image()
        elif action_type == "mouse_scroll":
            params = self.config_mouse_scroll()
        elif action_type == "type_text":
            params = self.config_type_text()
        elif action_type == "type_variable":
            params = self.config_type_variable()
        elif action_type == "key_press":
            params = self.config_key_press()
        elif action_type == "delay":
            params = self.config_delay()
        elif action_type == "wait_image":
            params = self.config_wait_image()
        elif action_type == "screenshot":
            params = self.config_screenshot()
        if params is not None:
            self.result = {"type": action_type, "params": params}
            self.destroy()

    def _setup_sub_dialog(self, title, geometry):
        """서브 다이얼로그 공통 설정 헬퍼"""
        dialog = ctk.CTkToplevel(self)
        dialog.title(title)
        dialog.geometry(geometry)
        dialog.configure(fg_color=Colors.BG_SECONDARY)
        dialog.withdraw()
        dialog.transient(self)
        dialog.grab_set()
        dialog.attributes("-topmost", True)
        set_dialog_icon(dialog)
        return dialog

    def _finalize_sub_dialog(self, dialog):
        """서브 다이얼로그 공통 마무리 (center + focus)"""

        def _show():
            self._center_sub_dialog(dialog)
            dialog.deiconify()
            dialog.lift()
            dialog.focus_force()

        dialog.after(50, _show)

    def config_click_coord(self):
        """좌표 클릭 설정"""
        if not self.coord_mgr.coordinates:
            self._show_error_dialog("등록된 좌표가 없습니다.")
            return None

        dialog = self._setup_sub_dialog("좌표 클릭 설정", "360x270")

        result = [None]

        main = ctk.CTkFrame(dialog, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=Sizes.PAD_LG, pady=Sizes.PAD_LG)

        # 좌표 선택
        Styles.section_title(main, text="좌표 선택").pack(
            anchor="w", pady=(0, Sizes.PAD_XS)
        )

        coord_values = [f"{c['id']}. {c['name']}" for c in self.coord_mgr.coordinates]
        coord_var = tk.StringVar(value=coord_values[0])
        coord_combo = Styles.combobox(main, variable=coord_var, values=coord_values)
        coord_combo.set(coord_values[0])
        coord_combo.pack(fill="x", pady=(0, Sizes.PAD_MD))

        # 클릭 유형
        click_frame = ctk.CTkFrame(main, fg_color="transparent")
        click_frame.pack(fill="x", pady=(0, Sizes.PAD_SM))

        Styles.section_title(click_frame, text="클릭 유형").pack(
            side="left", padx=(0, Sizes.PAD_MD)
        )

        click_var = tk.StringVar(value="left")
        ctk.CTkRadioButton(
            click_frame,
            text="좌클릭",
            variable=click_var,
            value="left",
            font=Fonts.SMALL,
            text_color=Colors.TEXT_PRIMARY,
            fg_color=Colors.PRIMARY,
            hover_color=Colors.PRIMARY_HOVER,
            border_color=Colors.BORDER_SUBTLE,
        ).pack(side="left", padx=Sizes.PAD_XS)
        ctk.CTkRadioButton(
            click_frame,
            text="우클릭",
            variable=click_var,
            value="right",
            font=Fonts.SMALL,
            text_color=Colors.TEXT_PRIMARY,
            fg_color=Colors.PRIMARY,
            hover_color=Colors.PRIMARY_HOVER,
            border_color=Colors.BORDER_SUBTLE,
        ).pack(side="left", padx=Sizes.PAD_XS)

        # 클릭 횟수
        count_frame = ctk.CTkFrame(main, fg_color="transparent")
        count_frame.pack(fill="x", pady=(0, Sizes.PAD_MD))

        Styles.section_title(count_frame, text="클릭 횟수").pack(
            side="left", padx=(0, Sizes.PAD_MD)
        )

        count_var = tk.IntVar(value=1)
        count_entry = Styles.input_field(
            count_frame,
            width=80,
            textvariable=count_var,
        )
        count_entry.pack(side="left")

        def on_ok():
            selected = coord_var.get()
            selected_idx = (
                coord_values.index(selected) if selected in coord_values else 0
            )
            coord_id = self.coord_mgr.coordinates[selected_idx]["id"]

            click_type = click_var.get()
            click_count = count_var.get()

            result[0] = {
                "coord_id": coord_id,
                "click_type": click_type,
                "click_count": click_count,
                "pre_delay": 0.2,
                "post_delay": 0.2,
            }
            dialog.destroy()

        Styles.primary_button(main, text="확인", command=on_ok).pack(fill="x")

        self._finalize_sub_dialog(dialog)

        self.wait_window(dialog)
        return result[0]

    def config_mouse_scroll(self):
        """마우스 스크롤 설정"""
        dialog = self._setup_sub_dialog("마우스 스크롤 설정", "320x200")

        result = [None]

        main = ctk.CTkFrame(dialog, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=Sizes.PAD_LG, pady=Sizes.PAD_LG)

        # 카드 영역
        card = Styles.card(main)
        card.pack(fill="x", pady=(0, Sizes.PAD_MD))
        card_inner = ctk.CTkFrame(card, fg_color="transparent")
        card_inner.pack(fill="x", padx=Sizes.PAD_MD, pady=Sizes.PAD_MD)

        # 스크롤 방향
        direction_frame = ctk.CTkFrame(card_inner, fg_color="transparent")
        direction_frame.pack(fill="x", pady=(0, Sizes.PAD_SM))

        Styles.section_title(direction_frame, text="스크롤 방향").pack(
            side="left", padx=(0, Sizes.PAD_MD)
        )

        direction_var = tk.StringVar(value="down")
        ctk.CTkRadioButton(
            direction_frame,
            text="위로",
            variable=direction_var,
            value="up",
            font=Fonts.SMALL,
            text_color=Colors.TEXT_PRIMARY,
            fg_color=Colors.PRIMARY,
            hover_color=Colors.PRIMARY_HOVER,
            border_color=Colors.BORDER_SUBTLE,
        ).pack(side="left", padx=Sizes.PAD_XS)
        ctk.CTkRadioButton(
            direction_frame,
            text="아래로",
            variable=direction_var,
            value="down",
            font=Fonts.SMALL,
            text_color=Colors.TEXT_PRIMARY,
            fg_color=Colors.PRIMARY,
            hover_color=Colors.PRIMARY_HOVER,
            border_color=Colors.BORDER_SUBTLE,
        ).pack(side="left", padx=Sizes.PAD_XS)

        # 스크롤 양
        amount_frame = ctk.CTkFrame(card_inner, fg_color="transparent")
        amount_frame.pack(fill="x")

        Styles.section_title(amount_frame, text="스크롤 양").pack(
            side="left", padx=(0, Sizes.PAD_MD)
        )

        amount_var = tk.IntVar(value=3)
        amount_entry = Styles.input_field(
            amount_frame,
            width=80,
            textvariable=amount_var,
        )
        amount_entry.pack(side="left")

        def on_ok():
            direction = direction_var.get()
            amount = amount_var.get()

            result[0] = {"direction": direction, "amount": amount}
            dialog.destroy()

        Styles.primary_button(main, text="확인", command=on_ok).pack(fill="x")

        self._finalize_sub_dialog(dialog)

        self.wait_window(dialog)
        return result[0]

    def config_click_image(self):
        """이미지 클릭 설정"""
        if not self.image_mgr.images:
            self._show_error_dialog("등록된 이미지가 없습니다.")
            return None

        dialog = self._setup_sub_dialog("이미지 클릭 설정", "340x220")

        result = [None]

        main = ctk.CTkFrame(dialog, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=Sizes.PAD_LG, pady=Sizes.PAD_LG)

        Styles.section_title(main, text="이미지 선택").pack(
            anchor="w", pady=(0, Sizes.PAD_XS)
        )

        image_values = [f"{img['id']}. {img['name']}" for img in self.image_mgr.images]
        image_var = tk.StringVar(value=image_values[0])
        image_combo = Styles.combobox(main, variable=image_var, values=image_values)
        image_combo.set(image_values[0])
        image_combo.pack(fill="x", pady=(0, Sizes.PAD_LG))

        def on_ok():
            selected = image_var.get()
            selected_idx = (
                image_values.index(selected) if selected in image_values else 0
            )
            image_id = self.image_mgr.images[selected_idx]["id"]

            result[0] = {"image_id": image_id}
            dialog.destroy()

        Styles.primary_button(main, text="확인", command=on_ok).pack(fill="x")

        self._finalize_sub_dialog(dialog)

        self.wait_window(dialog)
        return result[0]

    def config_type_text(self):
        """텍스트 타이핑 설정"""
        dialog = self._setup_sub_dialog("텍스트 입력", "340x200")
        dialog.resizable(False, False)

        result = [None]

        main = ctk.CTkFrame(dialog, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=Sizes.PAD_LG, pady=Sizes.PAD_LG)

        ctk.CTkLabel(
            main,
            text="입력할 텍스트",
            font=Fonts.BODY_BOLD,
            text_color=Colors.TEXT_PRIMARY,
        ).pack(anchor="w", pady=(0, Sizes.PAD_SM))

        # 카드 영역
        card = Styles.card(main)
        card.pack(fill="x", pady=(0, Sizes.PAD_MD))
        card_inner = ctk.CTkFrame(card, fg_color="transparent")
        card_inner.pack(fill="x", padx=Sizes.PAD_MD, pady=Sizes.PAD_MD)

        text_entry = Styles.input_field(
            card_inner, placeholder_text="텍스트를 입력하세요"
        )
        text_entry.pack(fill="x")
        text_entry.focus()

        def on_ok():
            text = text_entry.get().strip()
            if not text:
                messagebox.showwarning("경고", "텍스트를 입력하세요.", parent=dialog)
                return
            result[0] = {"text": text, "interval": 0.05}
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        Styles.primary_button(main, text="확인", command=on_ok).pack(fill="x")

        text_entry.bind("<Return>", lambda e: on_ok())
        text_entry.bind("<Escape>", lambda e: on_cancel())

        self._finalize_sub_dialog(dialog)

        self.wait_window(dialog)
        return result[0]

    def config_type_variable(self):
        """변수 타이핑 설정"""
        if not self.excel_mgr.excel_sources:
            self._show_error_dialog("등록된 엑셀 데이터가 없습니다.")
            return None

        dialog = self._setup_sub_dialog("변수 타이핑 설정", "420x520")

        result = [None]

        main = ctk.CTkFrame(dialog, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=Sizes.PAD_LG, pady=Sizes.PAD_LG)

        # 변수 유형
        Styles.section_title(main, text="변수 유형").pack(
            anchor="w", pady=(0, Sizes.PAD_XS)
        )

        var_type = tk.StringVar(value="excel")
        radio_frame = ctk.CTkFrame(main, fg_color="transparent")
        radio_frame.pack(anchor="w", padx=Sizes.PAD_SM, pady=(0, Sizes.PAD_MD))

        ctk.CTkRadioButton(
            radio_frame,
            text="엑셀 데이터",
            variable=var_type,
            value="excel",
            font=Fonts.SMALL,
            text_color=Colors.TEXT_PRIMARY,
            fg_color=Colors.PRIMARY,
            hover_color=Colors.PRIMARY_HOVER,
            border_color=Colors.BORDER_SUBTLE,
        ).pack(side="left", padx=Sizes.PAD_XS)
        ctk.CTkRadioButton(
            radio_frame,
            text="현재 행 번호",
            variable=var_type,
            value="counter",
            font=Fonts.SMALL,
            text_color=Colors.TEXT_PRIMARY,
            fg_color=Colors.PRIMARY,
            hover_color=Colors.PRIMARY_HOVER,
            border_color=Colors.BORDER_SUBTLE,
        ).pack(side="left", padx=Sizes.PAD_XS)
        ctk.CTkRadioButton(
            radio_frame,
            text="타임스탬프",
            variable=var_type,
            value="timestamp",
            font=Fonts.SMALL,
            text_color=Colors.TEXT_PRIMARY,
            fg_color=Colors.PRIMARY,
            hover_color=Colors.PRIMARY_HOVER,
            border_color=Colors.BORDER_SUBTLE,
        ).pack(side="left", padx=Sizes.PAD_XS)

        # 엑셀 소스 선택
        Styles.section_title(main, text="엑셀 소스 선택").pack(
            anchor="w", pady=(0, Sizes.PAD_XS)
        )

        excel_values = [
            f"{idx+1}. {src['name']}"
            for idx, src in enumerate(self.excel_mgr.excel_sources)
        ]
        excel_var = tk.StringVar(value=excel_values[0])
        excel_combo = Styles.combobox(main, variable=excel_var, values=excel_values)
        excel_combo.set(excel_values[0])
        excel_combo.pack(fill="x", pady=(0, Sizes.PAD_MD))

        # 칼럼 선택
        Styles.section_title(main, text="칼럼 선택 (엑셀 데이터인 경우)").pack(
            anchor="w", pady=(0, Sizes.PAD_XS)
        )

        column_var = tk.StringVar()
        column_combo = Styles.combobox(main, variable=column_var, values=[""])
        column_combo.pack(fill="x", pady=(0, Sizes.PAD_MD))

        # 데이터 처리 옵션
        Styles.section_title(main, text="데이터 처리 옵션").pack(
            anchor="w", pady=(0, Sizes.PAD_XS)
        )

        options_card = Styles.card(main)
        options_card.pack(fill="x", pady=(0, Sizes.PAD_MD))

        options_inner = ctk.CTkFrame(options_card, fg_color="transparent")
        options_inner.pack(fill="x", padx=Sizes.PAD_MD, pady=Sizes.PAD_SM)

        # 공백 행 제거 옵션
        remove_empty_var = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            options_inner,
            text="상위 공백 행 제거",
            variable=remove_empty_var,
            font=Fonts.CAPTION,
            text_color=Colors.TEXT_PRIMARY,
            fg_color=Colors.PRIMARY,
            hover_color=Colors.PRIMARY_HOVER,
            border_color=Colors.BORDER_SUBTLE,
            checkmark_color=Colors.TEXT_PRIMARY,
        ).pack(anchor="w", pady=(0, Sizes.PAD_XS))

        # 중복 행 제거 옵션
        remove_dup_var = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            options_inner,
            text="중복 행 제거",
            variable=remove_dup_var,
            font=Fonts.CAPTION,
            text_color=Colors.TEXT_PRIMARY,
            fg_color=Colors.PRIMARY,
            hover_color=Colors.PRIMARY_HOVER,
            border_color=Colors.BORDER_SUBTLE,
            checkmark_color=Colors.TEXT_PRIMARY,
        ).pack(anchor="w", pady=(0, Sizes.PAD_XS))

        ctk.CTkLabel(
            options_inner,
            text="※ 옵션 변경 시 선택한 엑셀 소스에 적용됩니다",
            font=Fonts.TINY,
            text_color=Colors.TEXT_MUTED,
        ).pack(anchor="w", pady=(Sizes.PAD_XS, 0))

        # 엑셀 소스 변경 시 칼럼 목록 및 옵션 업데이트
        def on_excel_change(choice=None):
            selected = excel_var.get()
            selected_idx = (
                excel_values.index(selected) if selected in excel_values else 0
            )
            if selected_idx >= 0:
                excel_source = self.excel_mgr.excel_sources[selected_idx]
                columns = excel_source["columns"]
                column_combo.configure(values=columns if columns else [""])
                if columns:
                    column_combo.set(columns[0])

                # 현재 엑셀 소스의 옵션값 로드
                remove_empty_var.set(excel_source.get("remove_empty_rows", True))
                remove_dup_var.set(excel_source.get("remove_duplicates", False))

        excel_combo.configure(command=on_excel_change)

        # 초기 칼럼 목록 및 옵션 설정
        on_excel_change()

        def on_ok():
            vtype = var_type.get()
            selected = excel_var.get()
            selected_excel_idx = (
                excel_values.index(selected) if selected in excel_values else 0
            )
            excel_source = self.excel_mgr.excel_sources[selected_excel_idx]
            excel_id = excel_source.get("id", selected_excel_idx + 1)
            excel_name = excel_source.get("name")
            var_name = column_combo.get() if vtype == "excel" else ""

            # 엑셀 소스의 데이터 처리 옵션 업데이트
            excel_source["remove_empty_rows"] = remove_empty_var.get()
            excel_source["remove_duplicates"] = remove_dup_var.get()

            result[0] = {
                "var_type": vtype,
                "var_name": var_name,
                "excel_id": excel_id,  # 엑셀 ID 추가
                "excel_name": excel_name,
                "interval": 0.05,
            }
            dialog.destroy()

        Styles.primary_button(main, text="확인", command=on_ok).pack(fill="x")

        self._finalize_sub_dialog(dialog)

        self.wait_window(dialog)
        return result[0]

    def config_key_press(self):
        """키 입력 설정 (자동 감지)"""
        dialog = KeyInputDialog(self)
        self.wait_window(dialog)

        if dialog.result:
            return {"key": dialog.result}
        return None

    def config_delay(self):
        """딜레이 설정"""
        dialog = self._setup_sub_dialog("딜레이 설정", "340x240")

        result = [None]

        main = ctk.CTkFrame(dialog, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=Sizes.PAD_LG, pady=Sizes.PAD_LG)

        ctk.CTkLabel(
            main,
            text="대기 시간 설정",
            font=Fonts.SUBHEADING,
            text_color=Colors.TEXT_PRIMARY,
        ).pack(anchor="w", pady=(0, Sizes.PAD_XS))

        ctk.CTkLabel(
            main,
            text="대기시간 입력 : 0.1 ~ 3600(초)",
            font=Fonts.SMALL,
            text_color=Colors.TEXT_SECONDARY,
        ).pack(anchor="w", pady=(0, Sizes.PAD_MD))

        # 카드 영역
        card = Styles.card(main)
        card.pack(fill="x", pady=(0, Sizes.PAD_MD))
        card_inner = ctk.CTkFrame(card, fg_color="transparent")
        card_inner.pack(fill="x", padx=Sizes.PAD_MD, pady=Sizes.PAD_MD)

        entry = Styles.input_field(card_inner, width=150)
        entry.insert(0, "1.0")
        entry.pack(anchor="w")
        entry.focus()
        entry.select_range(0, tk.END)

        def on_ok():
            try:
                seconds = float(entry.get())
                if 0.1 <= seconds <= 3600:
                    result[0] = {"seconds": seconds}
                    dialog.destroy()
                else:
                    self._show_error_dialog("0.1초 ~ 3600초 사이의 값을 입력하세요.")
            except ValueError:
                self._show_error_dialog("올바른 숫자를 입력하세요.")

        Styles.primary_button(main, text="확인", command=on_ok).pack(fill="x")

        entry.bind("<Return>", lambda e: on_ok())

        self._finalize_sub_dialog(dialog)

        self.wait_window(dialog)
        return result[0]

    def config_wait_image(self):
        """이미지 대기 설정"""
        if not self.image_mgr.images:
            self._show_error_dialog("등록된 이미지가 없습니다.")
            return None

        dialog = self._setup_sub_dialog("이미지 대기 설정", "340x280")

        result = [None]

        main = ctk.CTkFrame(dialog, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=Sizes.PAD_LG, pady=Sizes.PAD_LG)

        # 이미지 선택
        Styles.section_title(main, text="이미지 선택").pack(
            anchor="w", pady=(0, Sizes.PAD_XS)
        )

        image_values = [f"{img['id']}. {img['name']}" for img in self.image_mgr.images]
        image_var = tk.StringVar(value=image_values[0])
        image_combo = Styles.combobox(main, variable=image_var, values=image_values)
        image_combo.set(image_values[0])
        image_combo.pack(fill="x", pady=(0, Sizes.PAD_MD))

        # 최대 대기 시간
        Styles.section_title(main, text="최대 대기 시간 (초)").pack(
            anchor="w", pady=(0, Sizes.PAD_XS)
        )

        timeout_entry = Styles.input_field(main)
        timeout_entry.insert(0, "10")
        timeout_entry.pack(fill="x", pady=(0, Sizes.PAD_LG))

        def on_ok():
            selected = image_var.get()
            selected_idx = (
                image_values.index(selected) if selected in image_values else 0
            )
            image_id = self.image_mgr.images[selected_idx]["id"]
            timeout = float(timeout_entry.get())

            result[0] = {"image_id": image_id, "timeout": timeout}
            dialog.destroy()

        Styles.primary_button(main, text="확인", command=on_ok).pack(fill="x")

        self._finalize_sub_dialog(dialog)

        self.wait_window(dialog)
        return result[0]

    def config_screenshot(self):
        """스크린샷 설정"""
        # Step 1: 영역 설정 모드 선택 (전체 or 선택)
        region_mode_dialog = ScreenshotRegionModeDialog(self)
        self.wait_window(region_mode_dialog)

        mode = region_mode_dialog.result
        if not mode:  # 취소
            return None

        region_data = None

        # Step 2: 선택 모드인 경우 영역 선택
        if mode == "region":
            region_selector = ScreenRegionSelectorDialog(self)
            self.wait_window(region_selector)

            region_data = region_selector.result
            if not region_data:  # 취소
                return None

        # Step 3: 파일명 입력
        dialog = self._setup_sub_dialog("스크린샷", "340x230")
        dialog.resizable(False, False)

        result = [None]

        main = ctk.CTkFrame(dialog, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=Sizes.PAD_LG, pady=Sizes.PAD_LG)

        ctk.CTkLabel(
            main,
            text="파일명 (자동으로 타임스탬프 추가)",
            font=Fonts.BODY_BOLD,
            text_color=Colors.TEXT_PRIMARY,
        ).pack(anchor="w", pady=(0, Sizes.PAD_SM))

        # 카드 영역
        card = Styles.card(main)
        card.pack(fill="x", pady=(0, Sizes.PAD_MD))
        card_inner = ctk.CTkFrame(card, fg_color="transparent")
        card_inner.pack(fill="x", padx=Sizes.PAD_MD, pady=Sizes.PAD_MD)

        filename_entry = Styles.input_field(
            card_inner, placeholder_text="파일명을 입력하세요"
        )
        filename_entry.insert(0, "screenshot")
        filename_entry.pack(fill="x")
        filename_entry.select_range(0, tk.END)
        filename_entry.focus()

        def on_ok():
            filename = filename_entry.get().strip()
            if not filename:
                messagebox.showwarning("경고", "파일명을 입력하세요.", parent=dialog)
                return

            # 파일명과 영역 데이터를 함께 반환
            result[0] = {
                "filename": filename,
                "mode": mode,
                "region": region_data,  # None (전체) 또는 {'x', 'y', 'width', 'height'}
            }
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        Styles.primary_button(main, text="확인", command=on_ok).pack(fill="x")

        filename_entry.bind("<Return>", lambda e: on_ok())
        filename_entry.bind("<Escape>", lambda e: on_cancel())

        self._finalize_sub_dialog(dialog)

        self.wait_window(dialog)
        return result[0]


class ScreenshotRegionModeDialog(ctk.CTkToplevel):
    """스크린샷 영역 모드 선택 다이얼로그"""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("스크린샷 영역 설정")
        self.geometry("360x180")
        self.resizable(False, False)
        self.configure(fg_color=Colors.BG_SECONDARY)

        self.result = None  # 'full' or 'region'

        # 모달 설정
        self.withdraw()
        self.transient(parent)
        self.grab_set()
        self.attributes("-topmost", True)
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
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=Sizes.PAD_LG, pady=Sizes.PAD_LG)

        # 제목
        ctk.CTkLabel(
            main_frame,
            text="스크린샷 영역 설정",
            font=Fonts.SUBHEADING,
            text_color=Colors.TEXT_PRIMARY,
        ).pack(pady=(0, Sizes.PAD_SM))

        # 설명
        ctk.CTkLabel(
            main_frame,
            text="스크린샷 영역을 선택하세요.",
            font=Fonts.CAPTION,
            text_color=Colors.TEXT_MUTED,
        ).pack(pady=(0, Sizes.PAD_LG))

        # 버튼 프레임
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x")

        # 전체 화면 버튼
        Styles.primary_button(
            btn_frame,
            text="전체 화면",
            command=self.select_full,
        ).pack(side="left", expand=True, fill="x", padx=(0, Sizes.PAD_XS))

        # 영역 선택 버튼
        ctk.CTkButton(
            btn_frame,
            text="영역 선택",
            font=Fonts.BODY_BOLD,
            fg_color=Colors.SUCCESS,
            hover_color=Colors.SUCCESS_HOVER,
            text_color=Colors.TEXT_INVERSE,
            height=Sizes.BTN_HEIGHT_LG,
            corner_radius=Sizes.RADIUS_MD,
            command=self.select_region,
        ).pack(side="left", expand=True, fill="x", padx=(Sizes.PAD_XS, 0))

    def select_full(self):
        """전체 화면 선택"""
        self.result = "full"
        self.destroy()

    def select_region(self):
        """영역 선택"""
        self.result = "region"
        self.destroy()


class ScreenRegionSelectorDialog(tk.Toplevel):
    """화면 영역 선택 오버레이 - tk.Toplevel 유지 (raw Canvas 필요)"""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("영역 선택")

        self.result = None  # {'x': int, 'y': int, 'width': int, 'height': int}

        # 부모 윈도우와 메인 윈도우 찾기
        self.parent_window = parent

        # 실제 root 윈도우(Tk 인스턴스) 찾기
        current = parent
        while current.master:
            current = current.master
        self.main_window = current

        # 윈도우를 숨긴 상태로 시작
        self.withdraw()

        # 윈도우 초기화 후 최소화 및 UI 구성
        self.after(100, self.minimize_and_setup)

    def minimize_and_setup(self):
        """윈도우 최소화 후 UI 구성"""
        try:
            # 부모 윈도우들 업데이트
            self.update_idletasks()
            self.parent_window.update_idletasks()
            self.main_window.update_idletasks()

            # 모든 윈도우 최소화 (액션 다이얼로그 + 메인 윈도우)
            self.parent_window.withdraw()  # 액션 다이얼로그 숨기기
            self.main_window.iconify()  # 메인 윈도우 최소화

            # 최소화가 적용될 시간 대기 후 UI 구성 (200ms)
            self.after(200, self.setup_ui)
        except Exception as e:
            print(f"[WARNING] 윈도우 최소화 오류: {e}")
            # 오류 발생 시에도 UI는 구성
            self.after(200, self.setup_ui)

    def setup_ui(self):
        """UI 구성 (최소화 후 실행)"""
        # 전체 화면 크기
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # 전체 화면 크기로 설정
        self.geometry(f"{screen_width}x{screen_height}+0+0")
        self.attributes("-topmost", True)
        self.attributes("-alpha", 0.3)  # 반투명
        self.configure(bg=Colors.GRAY_800)
        self.overrideredirect(True)  # 타이틀바 제거

        # 드래그 변수
        self.start_x = None
        self.start_y = None
        self.rect = None

        # 캔버스 생성
        self.canvas = tk.Canvas(
            self,
            width=screen_width,
            height=screen_height,
            bg=Colors.GRAY_800,
            highlightthickness=0,
            cursor="crosshair",
        )
        self.canvas.pack(fill="both", expand=True)

        # 안내 텍스트
        self.canvas.create_text(
            screen_width // 2,
            30,
            text="드래그하여 영역을 선택하세요 (ESC: 취소)",
            font=Fonts.HEADING,
            fill=Colors.TEXT_PRIMARY,
        )

        # 이벤트 바인딩
        self.canvas.bind("<Button-1>", self.on_mouse_down)
        self.canvas.bind("<B1-Motion>", self.on_mouse_move)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_up)
        self.bind("<Escape>", lambda e: self.cancel())

        # 윈도우 표시 및 포커스
        self.deiconify()
        self.lift()
        self.focus_force()

    def on_mouse_down(self, event):
        """마우스 다운"""
        self.start_x = event.x
        self.start_y = event.y

        # 기존 사각형 제거
        if self.rect:
            self.canvas.delete(self.rect)

    def on_mouse_move(self, event):
        """마우스 드래그"""
        if self.start_x is None or self.start_y is None:
            return

        # 기존 사각형 제거
        if self.rect:
            self.canvas.delete(self.rect)

        # 새 사각형 그리기
        self.rect = self.canvas.create_rectangle(
            self.start_x,
            self.start_y,
            event.x,
            event.y,
            outline=Colors.PRIMARY_LIGHT,
            width=2,
            fill="",
        )

    def on_mouse_up(self, event):
        """마우스 업"""
        if self.start_x is None or self.start_y is None:
            return

        end_x = event.x
        end_y = event.y

        # 좌표 정규화 (왼쪽 위가 시작점이 되도록)
        x1 = min(self.start_x, end_x)
        y1 = min(self.start_y, end_y)
        x2 = max(self.start_x, end_x)
        y2 = max(self.start_y, end_y)

        width = x2 - x1
        height = y2 - y1

        # 최소 크기 체크 (10x10 이상)
        if width < 10 or height < 10:
            messagebox.showwarning("경고", "영역이 너무 작습니다. 다시 선택해주세요.")
            self.start_x = None
            self.start_y = None
            if self.rect:
                self.canvas.delete(self.rect)
            return

        # 결과 저장
        self.result = {"x": x1, "y": y1, "width": width, "height": height}

        # 메인 윈도우 복원 후 닫기
        self.restore_main_window()
        self.destroy()

    def cancel(self):
        """취소"""
        self.result = None
        # 메인 윈도우 복원 후 닫기
        self.restore_main_window()
        self.destroy()

    def restore_main_window(self):
        """메인 윈도우 및 부모 윈도우 복원"""
        try:
            # 메인 윈도우 복원
            self.main_window.deiconify()
            self.main_window.lift()

            # 액션 다이얼로그 복원
            self.parent_window.deiconify()
            self.parent_window.lift()
            self.parent_window.focus_force()
        except Exception as e:
            print(f"[WARNING] 윈도우 복원 오류: {e}")


class AutostartSelectDialog(ctk.CTkToplevel):
    """자동시작 프로젝트 선택 다이얼로그"""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("자동시작 설정")
        self.geometry("340x240")
        self.resizable(False, False)
        self.configure(fg_color=Colors.BG_SECONDARY)

        self.result = None

        # 모달 설정
        self.withdraw()
        self.transient(parent)
        self.grab_set()
        self.attributes("-topmost", True)
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
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=Sizes.PAD_LG, pady=Sizes.PAD_LG)

        # 제목
        ctk.CTkLabel(
            main_frame,
            text="Auto",
            font=Fonts.SUBHEADING,
            text_color=Colors.TEXT_PRIMARY,
        ).pack(pady=(0, Sizes.PAD_XS))

        # 설명
        ctk.CTkLabel(
            main_frame,
            text="자동실행할 프로젝트를 선택하세요.",
            font=Fonts.CAPTION,
            text_color=Colors.TEXT_MUTED,
        ).pack(pady=(0, Sizes.PAD_MD))

        # 콤보박스
        Styles.section_title(main_frame, text="프로젝트").pack(
            anchor="w", pady=(0, Sizes.PAD_XS)
        )

        self.project_var = tk.StringVar()
        self.project_combo = Styles.combobox(
            main_frame, variable=self.project_var, values=[""]
        )
        self.project_combo.pack(fill="x", pady=(0, Sizes.PAD_LG))

        # 프로젝트 목록 로드
        self.load_projects()

        # 버튼 프레임
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x")

        # 확인 버튼
        Styles.primary_button(
            btn_frame,
            text="확인",
            command=self.on_ok,
        ).pack(side="left", padx=(0, Sizes.PAD_XS), expand=True, fill="x")

        # 취소 버튼
        Styles.secondary_button(
            btn_frame,
            text="취소",
            command=self.on_cancel,
        ).pack(side="left", padx=(Sizes.PAD_XS, 0), expand=True, fill="x")

    def load_projects(self):
        """프로젝트 및 체인 목록 로드"""
        from core.project_manager import ProjectManager
        import json

        projects = ProjectManager.get_project_list()
        self.projects = projects

        # 현재 자동시작 설정 읽기
        current_autostart = None
        try:
            with open("settings.json", "r", encoding="utf-8") as f:
                settings = json.load(f)
                current_autostart = settings.get("autostart")
        except:
            pass

        # 콤보박스 항목 리스트
        combo_items = ["없음"]
        selected_index = 0

        if projects:
            for idx, project in enumerate(projects):
                # 프로젝트와 체인 모두 표시
                item_type = project.get("type", "project")
                display_text = f"{project['name']}"

                combo_items.append(display_text)

                # 현재 자동시작 프로젝트인지 확인
                if current_autostart and project["filepath"] == current_autostart:
                    selected_index = idx + 1  # "없음" 다음 인덱스

        # 콤보박스 설정
        self.combo_items = combo_items
        self.project_combo.configure(values=combo_items)
        self.project_combo.set(combo_items[selected_index])

    def on_ok(self):
        """확인 버튼"""
        selected = self.project_var.get()
        selected_index = (
            self.combo_items.index(selected) if selected in self.combo_items else 0
        )

        # "없음" 선택
        if selected_index == 0:
            self.result = {"filepath": None, "name": None}
            self.destroy()
            return

        # 프로젝트/체인 선택
        if not self.projects or selected_index - 1 >= len(self.projects):
            messagebox.showwarning("경고", "올바른 항목을 선택하세요.", parent=self)
            return

        # "없음"을 제외한 인덱스
        project_index = selected_index - 1
        selected_item = self.projects[project_index]

        self.result = {
            "filepath": selected_item["filepath"],
            "name": selected_item["name"],
            "type": selected_item.get("type", "project"),
        }
        self.destroy()

    def on_cancel(self):
        """취소 버튼"""
        self.result = None
        self.destroy()
