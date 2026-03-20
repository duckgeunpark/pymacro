"""
전역 단축키 설정 다이얼로그
"""

import tkinter as tk
import customtkinter as ctk
from tkinter import messagebox
from utils.ui_helpers import set_dialog_icon, center_window_on_parent
from core.settings_manager import SettingsManager
from ui.theme import Colors, Fonts, Sizes, Styles


class HotkeySettingsDialog(ctk.CTkToplevel):
    """전역 단축키 설정 다이얼로그"""

    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("단축키 설정")
        self.geometry("300x380")
        self.resizable(False, False)
        self.withdraw()
        self.transient(parent)
        self.grab_set()
        self.attributes("-topmost", True)
        set_dialog_icon(self)

        self.hotkey_entries = {}
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
        outer.pack(fill="both", expand=True)

        # ── 헤더 ──
        header = Styles.card(
            outer, corner_radius=0, border_width=0, height=Sizes.HEADER_HEIGHT
        )
        header.pack(fill="x")
        header.pack_propagate(False)

        ctk.CTkLabel(
            header,
            text="전역 단축키 설정",
            font=Fonts.HEADING,
            text_color=Colors.TEXT_PRIMARY,
        ).pack(expand=True)

        # ── 메인 컨텐츠 카드 ──
        content_card = Styles.card(outer)
        content_card.pack(
            fill="both", expand=True, padx=Sizes.PAD_LG, pady=Sizes.PAD_LG
        )

        # 센터 정렬용 wrapper
        center_wrapper = ctk.CTkFrame(content_card, fg_color="transparent")
        center_wrapper.place(relx=0.5, rely=0.5, anchor="center")

        # 현재 설정 로드
        hotkeys = SettingsManager.get_hotkeys()

        Styles.section_title(center_wrapper, text="키 바인딩").pack(
            pady=(0, Sizes.PAD_SM)
        )

        # 단축키 입력 필드
        for idx, (action, label) in enumerate(
            [
                ("start", "시작"),
                ("pause", "일시정지"),
                ("stop", "중지"),
                ("focus", "맨 앞"),
            ]
        ):
            row_frame = ctk.CTkFrame(center_wrapper, fg_color="transparent")
            row_frame.pack(pady=Sizes.PAD_XS)

            ctk.CTkLabel(
                row_frame,
                text=label,
                font=Fonts.BODY,
                text_color=Colors.TEXT_SECONDARY,
                width=70,
                anchor="e",
            ).pack(side="left")

            entry = Styles.input_field(row_frame, width=120)
            value = hotkeys.get(action, "")
            entry.insert(0, value.upper())
            entry.pack(side="left", padx=(Sizes.PAD_SM, 0))

            self.hotkey_entries[action] = entry
            entry.bind(
                "<Button-1>",
                lambda event, a=action, e=entry: self.open_key_input_dialog(a, e),
            )

        # ── 하단 버튼 ──
        btn_frame = ctk.CTkFrame(center_wrapper, fg_color="transparent")
        btn_frame.pack(pady=(Sizes.PAD_LG, 0))

        ctk.CTkButton(
            btn_frame,
            text="저장",
            font=Fonts.BODY_BOLD,
            fg_color=Colors.SUCCESS,
            hover_color=Colors.SUCCESS_HOVER,
            text_color=Colors.TEXT_INVERSE,
            height=Sizes.BTN_HEIGHT_LG,
            corner_radius=Sizes.RADIUS_MD,
            width=100,
            command=self.save_hotkeys,
        ).pack(side="left", padx=(0, Sizes.PAD_XS))

        Styles.secondary_button(
            btn_frame,
            text="취소",
            command=self.destroy,
            width=100,
        ).pack(side="left", padx=(Sizes.PAD_XS, 0))

    def open_key_input_dialog(self, action, entry):
        """키 입력 다이얼로그 열기"""
        key_dialog = ctk.CTkToplevel(self)
        key_dialog.title("키 입력")
        key_dialog.geometry("360x280")
        key_dialog.resizable(False, False)
        key_dialog.transient(self)
        key_dialog.grab_set()
        key_dialog.attributes("-topmost", True)

        set_dialog_icon(key_dialog)

        dlg_outer = ctk.CTkFrame(key_dialog, fg_color=Colors.BG_SECONDARY)
        dlg_outer.pack(fill="both", expand=True)

        main_frame = ctk.CTkFrame(dlg_outer, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=Sizes.PAD_XL, pady=Sizes.PAD_XL)

        ctk.CTkLabel(
            main_frame,
            text="입력할 키를 누르세요",
            font=Fonts.SUBHEADING,
            text_color=Colors.TEXT_PRIMARY,
        ).pack(pady=(0, Sizes.PAD_MD))

        # 큰 키 표시 영역
        key_display = ctk.CTkLabel(
            main_frame,
            text="(키를 누르세요...)",
            font=Fonts.HEADING,
            fg_color=Colors.BG_INPUT,
            text_color=Colors.PRIMARY_LIGHT,
            corner_radius=Sizes.RADIUS_MD,
            height=70,
        )
        key_display.pack(fill="x", pady=(0, Sizes.PAD_MD))

        ctk.CTkLabel(
            main_frame,
            text="예: F1, F8, F11, Enter, Esc 등",
            font=Fonts.CAPTION,
            text_color=Colors.TEXT_MUTED,
        ).pack(pady=(0, Sizes.PAD_LG))

        captured_key = [None]

        def on_key_press(event):
            """키 입력 감지"""
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

            key_name = event.keysym.lower()

            if event.keysym in key_mapping:
                key_name = key_mapping[event.keysym]

            if event.keysym.startswith("F") and event.keysym[1:].isdigit():
                key_name = event.keysym.lower()

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

            modifiers = []
            if event.state & 0x0004:
                modifiers.append("ctrl")
            if event.state & 0x0001:
                modifiers.append("shift")

            if modifiers:
                captured_key[0] = "+".join(modifiers + [key_name])
            else:
                captured_key[0] = key_name

            key_display.configure(
                text=captured_key[0].upper(), text_color=Colors.SUCCESS
            )
            confirm_btn.configure(state="normal")

        def on_ok():
            if captured_key[0]:
                entry.delete(0, tk.END)
                entry.insert(0, captured_key[0].upper())
                key_dialog.destroy()

        def on_cancel():
            key_dialog.destroy()

        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x")

        confirm_btn = ctk.CTkButton(
            btn_frame,
            text="확인",
            font=Fonts.BODY_BOLD,
            fg_color=Colors.SUCCESS,
            hover_color=Colors.SUCCESS_HOVER,
            text_color=Colors.TEXT_INVERSE,
            height=Sizes.BTN_HEIGHT,
            corner_radius=Sizes.RADIUS_SM,
            command=on_ok,
            state="disabled",
        )
        confirm_btn.pack(side="left", expand=True, padx=(0, Sizes.PAD_XS))

        Styles.secondary_button(
            btn_frame,
            text="취소",
            command=on_cancel,
        ).pack(side="left", expand=True, padx=(Sizes.PAD_XS, 0))

        key_dialog.bind("<KeyPress>", on_key_press)
        key_dialog.focus_force()

        center_window_on_parent(key_dialog, self)

    def save_hotkeys(self):
        """단축키 저장"""
        hotkeys = {}
        for action, entry in self.hotkey_entries.items():
            hotkeys[action] = entry.get().strip().lower()

        if SettingsManager.save_hotkeys(hotkeys):
            messagebox.showinfo(
                "완료",
                "단축키가 저장되었습니다!\n\n모든 매크로에 적용됩니다.",
                parent=self,
            )
            self.destroy()
        else:
            messagebox.showerror("오류", "설정 저장에 실패했습니다.", parent=self)
