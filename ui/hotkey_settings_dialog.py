"""
전역 단축키 설정 다이얼로그
"""
import tkinter as tk
from tkinter import messagebox
from utils.ui_helpers import set_dialog_icon, center_window_on_parent
from core.settings_manager import SettingsManager


class HotkeySettingsDialog(tk.Toplevel):
    """전역 단축키 설정 다이얼로그"""

    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("⌨️ 단축키 설정")
        self.geometry("300x330")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.attributes('-topmost', True)

        set_dialog_icon(self)

        self.hotkey_entries = {}
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
            text="⌨️ 전역 단축키 설정",
            font=("맑은 고딕", 14, "bold"),
            bg='#2c3e50',
            fg='white'
        ).pack(pady=15)

        # 메인 컨텐츠
        content = tk.Frame(self, bg='#F0F0F0', padx=30, pady=20)
        content.pack(fill='both', expand=True)

        # tk.Label(
        #     content,
        #     text="모든 매크로에 적용되는 전역 단축키입니다.",
        #     font=("맑은 고딕", 10),
        #     bg='#F0F0F0',
        #     fg='#7f8c8d'
        # ).pack(pady=(0, 20))

        # 현재 설정 로드
        hotkeys = SettingsManager.get_hotkeys()

        # 단축키 입력 필드
        for idx, (action, label) in enumerate([
            ('start', '시작'),
            ('pause', '일시정지'),
            ('stop', '중지'),
            ('focus', '맨 앞')
        ]):
            row_frame = tk.Frame(content, bg='#F0F0F0')
            row_frame.pack(fill='x', pady=8)

            tk.Label(
                row_frame,
                text=label,
                font=("맑은 고딕", 11),
                bg='#F0F0F0',
                width=10,
                anchor='w'
            ).pack(side='left')

            entry = tk.Entry(
                row_frame,
                font=("맑은 고딕", 11),
                width=8,
                relief='solid',
                borderwidth=1
            )
            value = hotkeys.get(action, '')
            entry.insert(0, value.upper())
            entry.pack(side='left', padx=10)

            self.hotkey_entries[action] = entry
            entry.bind(
                "<Button-1>",
                lambda event, a=action, e=entry: self.open_key_input_dialog(a, e)
            )

        # # 힌트
        # tk.Label(
        #     content,
        #     text="※ 예: f8, f9, f10, f12, enter, ctrl+s 등",
        #     font=("맑은 고딕", 9),
        #     fg='#95a5a6',
        #     bg='#F0F0F0'
        # ).pack(pady=(20, 10))

        # 버튼
        btn_frame = tk.Frame(content, bg='#F0F0F0')
        btn_frame.pack(pady=15)

        tk.Button(
            btn_frame,
            text="저장",
            font=("맑은 고딕", 11, "bold"),
            bg='#27ae60',
            fg='white',
            padx=30,
            pady=5,
            command=self.save_hotkeys
        ).pack(side='left', padx=5)

        tk.Button(
            btn_frame,
            text="취소",
            font=("맑은 고딕", 11),
            bg='#95a5a6',
            fg='white',
            padx=30,
            pady=5,
            command=self.destroy
        ).pack(side='left', padx=5)

    def open_key_input_dialog(self, action, entry):
        """키 입력 다이얼로그 열기"""
        key_dialog = tk.Toplevel(self)
        key_dialog.title("키 입력")
        key_dialog.geometry("350x250")
        key_dialog.resizable(False, False)
        key_dialog.transient(self)
        key_dialog.grab_set()
        key_dialog.attributes('-topmost', True)

        set_dialog_icon(key_dialog)

        main_frame = tk.Frame(key_dialog, padx=20, pady=20, bg='white')
        main_frame.pack(fill='both', expand=True)

        tk.Label(
            main_frame,
            text="입력할 키를 누르세요",
            font=("맑은 고딕", 12, "bold"),
            bg='white'
        ).pack(pady=(0, 15))

        key_display = tk.Label(
            main_frame,
            text="(키를 누르세요...)",
            font=("맑은 고딕", 14, "bold"),
            bg='#ecf0f1',
            fg='#3498db',
            padx=20,
            pady=15,
            relief='sunken',
            borderwidth=2
        )
        key_display.pack(fill='x', pady=(0, 15))

        tk.Label(
            main_frame,
            text="예: F1, F8, F11, Enter, Esc 등",
            font=("맑은 고딕", 9),
            fg='#7f8c8d',
            bg='white'
        ).pack(pady=(0, 15))

        captured_key = [None]

        def on_key_press(event):
            """키 입력 감지"""
            key_mapping = {
                'Return': 'enter',
                'Tab': 'tab',
                'Escape': 'esc',
                'space': 'space',
                'BackSpace': 'backspace',
                'Delete': 'delete',
                'Up': 'up',
                'Down': 'down',
                'Left': 'left',
                'Right': 'right',
                'Home': 'home',
                'End': 'end',
                'Prior': 'pageup',
                'Next': 'pagedown',
                'Insert': 'insert',
                'Pause': 'pause',
                'Print': 'print',
            }

            key_name = event.keysym.lower()

            if event.keysym in key_mapping:
                key_name = key_mapping[event.keysym]

            if event.keysym.startswith('F') and event.keysym[1:].isdigit():
                key_name = event.keysym.lower()

            if key_name in ['shift_l', 'shift_r', 'control_l', 'control_r', 'alt_l', 'alt_r', 'meta_l', 'meta_r']:
                return

            modifiers = []
            if event.state & 0x0004:
                modifiers.append('ctrl')
            if event.state & 0x0001:
                modifiers.append('shift')

            if modifiers:
                captured_key[0] = '+'.join(modifiers + [key_name])
            else:
                captured_key[0] = key_name

            key_display.config(text=captured_key[0].upper(), fg='#27ae60')
            confirm_btn.config(state='normal')

        def on_ok():
            if captured_key[0]:
                entry.delete(0, tk.END)
                entry.insert(0, captured_key[0].upper())
                key_dialog.destroy()

        def on_cancel():
            key_dialog.destroy()

        btn_frame = tk.Frame(main_frame, bg='white')
        btn_frame.pack(fill='x')

        confirm_btn = tk.Button(
            btn_frame,
            text="확인",
            font=("맑은 고딕", 10),
            bg='#27ae60',
            fg='white',
            padx=20,
            pady=8,
            command=on_ok,
            state='disabled'
        )
        confirm_btn.pack(side='left', expand=True, padx=(0, 5))

        tk.Button(
            btn_frame,
            text="취소",
            font=("맑은 고딕", 10),
            bg='#95a5a6',
            fg='white',
            padx=20,
            pady=8,
            command=on_cancel
        ).pack(side='left', expand=True, padx=(5, 0))

        key_dialog.bind('<KeyPress>', on_key_press)
        key_dialog.focus_force()

        center_window_on_parent(key_dialog, self)

    def save_hotkeys(self):
        """단축키 저장"""
        hotkeys = {}
        for action, entry in self.hotkey_entries.items():
            hotkeys[action] = entry.get().strip().lower()

        if SettingsManager.save_hotkeys(hotkeys):
            messagebox.showinfo("완료", "단축키가 저장되었습니다!\n\n모든 매크로에 적용됩니다.", parent=self)
            self.destroy()
        else:
            messagebox.showerror("오류", "설정 저장에 실패했습니다.", parent=self)
