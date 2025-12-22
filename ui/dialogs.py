"""
각종 다이얼로그
"""
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from datetime import datetime
from utils.ui_helpers import set_dialog_icon, center_window_on_parent, center_window_on_screen

class NewProjectDialog(tk.Toplevel):
    """새 프로젝트 생성 다이얼로그"""
    
    def __init__(self, parent):
        super().__init__(parent)
        self.title("새 프로젝트")
        self.geometry("300x130")
        self.resizable(False, False)
        
        self.result = None
        
        # 모달 설정
        self.transient(parent)
        self.grab_set()
        self.attributes('-topmost', True)
        set_dialog_icon(self)
        self.setup_ui()

        # 창 중앙 배치
        center_window_on_screen(self)

        self.lift()
        self.focus_force()
    
    def setup_ui(self):
        """UI 구성"""
        # 메인 프레임
        main_frame = tk.Frame(self, padx=20, pady=10)
        main_frame.pack(fill='both', expand=True)
        
        # 프로젝트 이름
        tk.Label(
            main_frame,
            text="프로젝트 이름:",
            font=("맑은 고딕", 10)
        ).pack(anchor='w', pady=(0, 5))
        
        self.name_entry = tk.Entry(
            main_frame,
            font=("맑은 고딕", 10),
            width=45
        )
        self.name_entry.pack(fill='x', pady=(0, 5))
        self.name_entry.focus()

        # 글자 수 제한 (14자)
        self.name_entry.bind('<KeyRelease>', self.on_name_change)

        # # 설명
        # tk.Label(
        #     main_frame,
        #     text="설명 (선택사항):",
        #     font=("맑은 고딕", 10)
        # ).pack(anchor='w', pady=(0, 5))
        
        # self.desc_text = tk.Text(
        #     main_frame,
        #     font=("맑은 고딕", 10),
        #     height=4,
        #     width=40
        # )
        # self.desc_text.pack(fill='x', pady=(0, 20))
        
        # 버튼
        tk.Button(
            main_frame,
            text="확인",
            font=("맑은 고딕", 10),
            bg='#3498db',
            fg='white',
            padx=20,
            pady=5,
            command=self.on_create
        ).pack(pady=(10, 5))
        
        # Enter 키 바인딩
        self.name_entry.bind('<Return>', lambda e: self.on_create())

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
        invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
        if any(char in name for char in invalid_chars):
            messagebox.showerror(
                "오류",
                f"프로젝트 이름에 다음 문자를 사용할 수 없습니다:\n{' '.join(invalid_chars)}"
            )
            return
        
        description = ""
        
        self.result = {
            'name': name,
            'description': description
        }
        
        self.destroy()
    
    def on_cancel(self):
        """취소 버튼 클릭"""
        self.result = None
        self.destroy()

        
class NameInputDialog(tk.Toplevel):
    """이름 입력 다이얼로그"""
    
    def __init__(self, parent, title="이름 입력", message="이름을 입력하세요:", initial_value=""):
        super().__init__(parent)
        self.title(title)
        self.geometry("300x180")
        self.resizable(False, False)
        
        self.result = None
        
        # 모달 설정
        self.transient(parent)
        self.grab_set()
        
        # 항상 위에 표시
        self.attributes('-topmost', True)
        
        set_dialog_icon(self)


        self.setup_ui(message, initial_value)
        center_window_on_parent(self, self.master)

        # 포커스
        self.lift()
        self.focus_force()
    
    def setup_ui(self, message, initial_value):
        """UI 구성"""
        # 메시지
        tk.Label(
            self,
            text=message,
            font=("맑은 고딕", 11, "bold")
        ).pack(pady=20)

        # 입력 필드
        self.entry = tk.Entry(
            self,
            font=("맑은 고딕", 11),
            width=35
        )
        self.entry.pack(padx=30, pady=10)

        # 글자 수 제한 (10자)
        self.entry.bind('<KeyRelease>', self.on_name_change)

        if initial_value:
            self.entry.insert(0, initial_value)
            self.entry.select_range(0, tk.END)

        self.entry.focus_set()

        # 버튼
        tk.Button(
            self,
            text="확인",
            font=("맑은 고딕", 10),
            bg='#3498db',
            fg='white',
            padx=25,
            pady=8,
            command=self.on_ok
        ).pack(pady=20)

        # Enter 키로 확인
        self.entry.bind('<Return>', lambda e: self.on_ok())
        self.entry.bind('<Escape>', lambda e: self.on_cancel())

    def on_name_change(self, event=None):
        """이름 입력 시 길이 제한 (한글 6자, 숫자/영문 12자)"""
        current_text = self.entry.get()

        # 한글 글자 수 계산
        korean_count = sum(1 for char in current_text if ord('가') <= ord(char) <= ord('힣') or ord('ㄱ') <= ord(char) <= ord('ㅣ'))
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

class ImageNameDialog(tk.Toplevel):
    """이미지 이름 및 정확도 입력 다이얼로그"""

    def __init__(self, parent, title="이미지 이름", initial_name="", initial_confidence=80):
        super().__init__(parent)
        self.title(title)
        self.geometry("350x220")
        self.resizable(False, False)

        self.result = None  # {'name': str, 'confidence': float}

        # 모달 설정
        self.transient(parent)
        self.grab_set()
        self.attributes('-topmost', True)

        set_dialog_icon(self)

        self.setup_ui(initial_name, initial_confidence)
        center_window_on_parent(self, self.master)

        # 포커스
        self.lift()
        self.focus_force()

    def setup_ui(self, initial_name, initial_confidence):
        """UI 구성"""
        main_frame = tk.Frame(self, padx=20, pady=15)
        main_frame.pack(fill='both', expand=True)

        # 이름 입력
        tk.Label(
            main_frame,
            text="이미지 이름:",
            font=("맑은 고딕", 10, "bold")
        ).pack(anchor='w', pady=(5, 5))

        self.name_entry = tk.Entry(
            main_frame,
            font=("맑은 고딕", 10),
            width=30
        )
        self.name_entry.pack(fill='x', pady=(0, 15))

        if initial_name:
            self.name_entry.insert(0, initial_name)

        self.name_entry.focus_set()

        # 정확도 입력
        confidence_frame = tk.Frame(main_frame)
        confidence_frame.pack(fill='x', pady=(0, 20))

        tk.Label(
            confidence_frame,
            text="정확도:",
            font=("맑은 고딕", 10, "bold"),
            width=8,
            anchor='w'
        ).pack(side='left')

        self.confidence_var = tk.IntVar(value=initial_confidence)

        confidence_spinbox = tk.Spinbox(
            confidence_frame,
            from_=50,
            to=100,
            textvariable=self.confidence_var,
            font=("맑은 고딕", 10),
            width=10
        )
        confidence_spinbox.pack(side='left', padx=5)

        tk.Label(
            confidence_frame,
            text="%",
            font=("맑은 고딕", 10)
        ).pack(side='left')

        tk.Label(
            confidence_frame,
            text="(50-100)",
            font=("맑은 고딕", 8),
            fg='gray'
        ).pack(side='left', padx=5)

        # 버튼
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack()

        tk.Button(
            btn_frame,
            text="확인",
            font=("맑은 고딕", 10),
            bg='#3498db',
            fg='white',
            padx=25,
            pady=8,
            command=self.on_ok
        ).pack(side='left', padx=5)

        tk.Button(
            btn_frame,
            text="취소",
            font=("맑은 고딕", 10),
            bg='#95a5a6',
            fg='white',
            padx=25,
            pady=8,
            command=self.on_cancel
        ).pack(side='left', padx=5)

        # Enter 키로 확인
        self.name_entry.bind('<Return>', lambda e: self.on_ok())
        self.bind('<Escape>', lambda e: self.on_cancel())

    def on_ok(self):
        """확인 버튼"""
        name = self.name_entry.get().strip()
        if not name:
            messagebox.showwarning("경고", "이미지 이름을 입력하세요.", parent=self)
            return

        confidence = self.confidence_var.get()

        self.result = {
            'name': name,
            'confidence': confidence / 100.0  # 0.5 ~ 1.0으로 변환
        }
        self.destroy()

    def on_cancel(self):
        """취소 버튼"""
        self.result = None
        self.destroy()


class KeyInputDialog(tk.Toplevel):
    """키 입력 감지 다이얼로그"""
    
    def __init__(self, parent):
        super().__init__(parent)
        self.title("키 입력")
        self.geometry("300x200")
        self.resizable(False, False)
        
        self.result = None
        self.captured_key = None
        self.is_first_key = True  # 첫 키 입력 플래그
        
        # 모달 설정
        self.transient(parent)
        self.grab_set()
        self.attributes('-topmost', True)
        
        # 아이콘 설정
        set_dialog_icon(self)
        
        self.setup_ui()
        self.center_window()
        
        # 포커스
        self.lift()
        self.focus_force()
        
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
        
        self.geometry(f'{dialog_width}x{dialog_height}+{x}+{y}')
    
    def setup_ui(self):
        """UI 구성"""
        main_frame = tk.Frame(self, padx=10, pady=10)
        main_frame.pack(fill='both', expand=True)
        
        # 설명
        tk.Label(
            main_frame,
            text="입력할 키를 누르세요",
            font=("맑은 고딕", 12, "bold")
        ).pack(pady=(0, 10))
        
        # 입력한 키 표시
        # tk.Label(
        #     main_frame,
        #     text="입력된 키:",
        #     font=("맑은 고딕", 11)
        # ).pack(anchor='w', pady=(0, 5))
        
        self.key_display = tk.Label(
            main_frame,
            text="(키를 누르세요...)",
            font=("맑은 고딕", 14, "bold"),
            bg='#ecf0f1',
            fg='#3498db',
            padx=15,
            pady=10,
            relief='sunken',
            borderwidth=2
        )
        self.key_display.pack(fill='x', pady=(0, 10))
        
        # 도움말
        tk.Label(
            main_frame,
            text="예: enter, tab, esc, space, f1, f11 등",
            font=("맑은 고딕", 9),
            fg='#7f8c8d'
        ).pack(anchor='w', pady=(0, 10))
        
        # 버튼
        self.confirm_btn = tk.Button(
            main_frame,
            text="확인",
            font=("맑은 고딕", 10),
            bg='#3498db',
            fg='white',
            padx=20,
            pady=5,
            command=self.on_ok,
            state='disabled'
        )
        self.confirm_btn.pack(pady=10)
        
        # 키 이벤트 바인드
        self.bind('<KeyPress>', self.on_key_press)
        self.bind('<KeyRelease>', self.on_key_release)  # 추가
        self.focus_force()
    
    def on_key_release(self, event):
        """키 해제 감지 (Alt 상태 초기화용)"""
        # Alt 키만 눌렸을 때는 무시
        if event.keysym in ['alt_l', 'alt_r']:
            return
    
    def on_key_press(self, event):
        """키 입력 감지"""
        # 키 이름 매핑
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
        
        # 키 이름 가져오기
        key_name = event.keysym.lower()
        
        # 매핑된 키 확인
        if event.keysym in key_mapping:
            key_name = key_mapping[event.keysym]
        
        # F1~F12 키 처리
        if event.keysym.startswith('F') and event.keysym[1:].isdigit():
            key_name = event.keysym.lower()
        
        # 특수 조합 무시 (Shift, Control, Alt만 눌렀을 때) - 수정
        if key_name in ['shift_l', 'shift_r', 'control_l', 'control_r', 'alt_l', 'alt_r', 'meta_l', 'meta_r']:
            return
        
        # 수정자 키 (Ctrl, Alt, Shift) 확인 - 수정
        modifiers = []
        
        # Alt 키 상태만 확인 (Grab으로 인한 이상 제거)
        if event.state & 0x0004:  # Ctrl
            modifiers.append('ctrl')
        # Alt 무시 (Alt 메뉴 접근 방지)
        # if event.state & 0x0008:  # Alt
        #     modifiers.append('alt')
        if event.state & 0x0001:  # Shift
            modifiers.append('shift')
        
        # 최종 키 조합 생성
        if modifiers:
            self.captured_key = '+'.join(modifiers + [key_name])
        else:
            self.captured_key = key_name
        
        # 디스플레이 업데이트
        self.key_display.config(text=self.captured_key, fg='#27ae60')
        
        # 확인 버튼 활성화
        self.confirm_btn.config(state='normal')
        
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



class ActionSelectDialog(tk.Toplevel):
    """액션 선택 다이얼로그"""
    def _show_error_dialog(self, message):
        dialog = tk.Toplevel(self)
        dialog.title("오류")
        dialog.geometry("320x150")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        dialog.attributes('-topmost', True)
        set_dialog_icon(dialog)
        
        tk.Label(
            dialog,
            text=message,
            font=("맑은 고딕", 11),
            wraplength=280
        ).pack(pady=30)
        
        tk.Button(
            dialog,
            text="확인",
            font=("맑은 고딕", 10),
            bg='#3498db',
            fg='white',
            padx=20,
            command=lambda: dialog.destroy()
        ).pack(pady=10)
        
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        parent_x = self.winfo_x()
        parent_y = self.winfo_y()
        parent_width = self.winfo_width()
        parent_height = self.winfo_height()
        x = parent_x + (parent_width - width) // 2
        y = parent_y + (parent_height - height) // 2
        dialog.geometry(f'{width}x{height}+{x}+{y}')
        
        dialog.lift()
        dialog.focus_force()
        
        self.wait_window(dialog)


    def __init__(self, parent, coord_mgr, excel_mgr, image_mgr):
        super().__init__(parent)
        self.title("액션 추가")
        self.geometry("400x440")
        self.resizable(False, False)
        
        self.coord_mgr = coord_mgr
        self.excel_mgr = excel_mgr
        self.image_mgr = image_mgr
        
        self.result = None
        
        self.transient(parent)
        self.grab_set()
        self.attributes('-topmost', True)
        
        # 아이콘 설정 (추가)
        set_dialog_icon(self)

        self.setup_ui()
        center_window_on_parent(self, parent)

        # 포커스
        self.lift()
        self.focus_force()
    
    def setup_ui(self):
        """UI 구성"""
        main_frame = tk.Frame(self, padx=20, pady=20)
        main_frame.pack(fill='both', expand=True)
        
        tk.Label(
            main_frame,
            text="액션 유형 선택",
            font=("맑은 고딕", 13, "bold")
        ).pack(pady=(0, 15))
        
        # 액션 버튼들
        self.create_action_buttons(main_frame)


    def create_action_buttons(self, parent):
        """액션 버튼 생성"""
        # 마우스 동작
        section = tk.LabelFrame(
            parent,
            text="🖱️ 마우스 동작",
            font=("맑은 고딕", 11, "bold"),
            padx=10,
            pady=10
        )
        section.pack(fill='x', pady=5)

        btn_frame = tk.Frame(section)
        btn_frame.pack()

        tk.Button(
            btn_frame,
            text="좌표 클릭",
            font=("맑은 고딕", 9),
            width=12,
            command=lambda: self.select_action('click_coord')
        ).grid(row=0, column=0, padx=5, pady=5)

        tk.Button(
            btn_frame,
            text="이미지 클릭",
            font=("맑은 고딕", 9),
            width=12,
            command=lambda: self.select_action('click_image')
        ).grid(row=0, column=1, padx=5, pady=5)

        tk.Button(
            btn_frame,
            text="마우스 스크롤",
            font=("맑은 고딕", 9),
            width=12,
            command=lambda: self.select_action('mouse_scroll')
        ).grid(row=0, column=2, padx=5, pady=5)

        # 키보드 동작
        section = tk.LabelFrame(
            parent,
            text="⌨️ 키보드 동작",
            font=("맑은 고딕", 11, "bold"),
            padx=10,
            pady=10
        )
        section.pack(fill='x', pady=5)

        btn_frame = tk.Frame(section)
        btn_frame.pack()

        tk.Button(
            btn_frame,
            text="텍스트 타이핑",
            font=("맑은 고딕", 9),
            width=12,
            command=lambda: self.select_action('type_text')
        ).grid(row=0, column=0, padx=5, pady=5)

        tk.Button(
            btn_frame,
            text="변수 타이핑",
            font=("맑은 고딕", 9),
            width=12,
            command=lambda: self.select_action('type_variable')
        ).grid(row=0, column=1, padx=5, pady=5)

        tk.Button(
            btn_frame,
            text="키 입력",
            font=("맑은 고딕", 9),
            width=12,
            command=lambda: self.select_action('key_press')
        ).grid(row=0, column=2, padx=5, pady=5)

        # 제어 동작
        section = tk.LabelFrame(
            parent,
            text="⏱️ 제어 동작",
            font=("맑은 고딕", 11, "bold"),
            padx=10,
            pady=10
        )
        section.pack(fill='x', pady=5)

        btn_frame = tk.Frame(section)
        btn_frame.pack()

        tk.Button(
            btn_frame,
            text="딜레이",
            font=("맑은 고딕", 9),
            width=12,
            command=lambda: self.select_action('delay')
        ).grid(row=0, column=0, padx=5, pady=5)

        tk.Button(
            btn_frame,
            text="이미지 대기",
            font=("맑은 고딕", 9),
            width=12,
            command=lambda: self.select_action('wait_image')
        ).grid(row=0, column=1, padx=5, pady=5)


        # 기타
        section = tk.LabelFrame(
            parent,
            text="💾 기타",
            font=("맑은 고딕", 11, "bold"),
            padx=10,
            pady=10
        )
        section.pack(fill='x', pady=5)
        
        btn_frame = tk.Frame(section)
        btn_frame.pack()
        
        tk.Button(
            btn_frame,
            text="스크린샷",
            font=("맑은 고딕", 9),
            width=12,
            command=lambda: self.select_action('screenshot')
        ).grid(row=0, column=0, padx=5, pady=5)
        
    
    def select_action(self, action_type):
        """액션 선택"""
        params = None

        if action_type == 'click_coord':
            params = self.config_click_coord()
        elif action_type == 'click_image':
            params = self.config_click_image()
        elif action_type == 'mouse_scroll':
            params = self.config_mouse_scroll()
        elif action_type == 'type_text':
            params = self.config_type_text()
        elif action_type == 'type_variable':
            params = self.config_type_variable()
        elif action_type == 'key_press':
            params = self.config_key_press()
        elif action_type == 'delay':
            params = self.config_delay()
        elif action_type == 'wait_image':
            params = self.config_wait_image()
        elif action_type == 'screenshot':
            params = self.config_screenshot()
        if params is not None:
            self.result = {
                'type': action_type,
                'params': params
            }
            self.destroy()
    
    def config_click_coord(self):
        """좌표 클릭 설정"""
        if not self.coord_mgr.coordinates:
            self._show_error_dialog("등록된 좌표가 없습니다.")
            return None
    
        dialog = tk.Toplevel(self)
        dialog.title("좌표 클릭 설정")
        dialog.geometry("300x220")
        dialog.transient(self)
        dialog.grab_set()
        dialog.attributes('-topmost', True)
        set_dialog_icon(dialog)
        
        result = [None]
        
        tk.Label(
            dialog,
            text="좌표 선택:",
            font=("맑은 고딕", 10, "bold")
        ).pack(anchor='w', padx=20, pady=(10, 5))
        
        coord_var = tk.StringVar()
        coord_combo = ttk.Combobox(
            dialog,
            textvariable=coord_var,
            font=("맑은 고딕", 10),
            state='readonly'
        )
        coord_values = [f"{c['id']}. {c['name']}" for c in self.coord_mgr.coordinates]
        coord_combo['values'] = coord_values
        coord_combo.current(0)
        coord_combo.pack(fill='x', padx=20, pady=(0, 10))
        

        click_frame = tk.Frame(dialog)
        click_frame.pack(fill='x', padx=20, pady=(10, 5))

        tk.Label(
            click_frame,
            text="클릭 유형:",
            font=("맑은 고딕", 10, "bold"),
            width=10,
            anchor='w'
        ).pack(side='left')

        click_var = tk.StringVar(value='left')
        tk.Radiobutton(click_frame, text="좌클릭", variable=click_var, value='left').pack(side='left', padx=5)
        tk.Radiobutton(click_frame, text="우클릭", variable=click_var, value='right').pack(side='left', padx=5)

        count_frame = tk.Frame(dialog)
        count_frame.pack(fill='x', padx=20, pady=(10, 5))

        tk.Label(
            count_frame,
            text="클릭 횟수:",
            font=("맑은 고딕", 10, "bold"),
            width=10,
            anchor='w'
        ).pack(side='left')

        count_var = tk.IntVar(value=1)
        count_spinbox = tk.Spinbox(
            count_frame,
            from_=1,
            to=10,
            textvariable=count_var,
            font=("맑은 고딕", 10),
            width=10
        )
        count_spinbox.pack(side='left')

        def on_ok():
            selected_idx = coord_combo.current()
            coord_id = self.coord_mgr.coordinates[selected_idx]['id']

            click_type = click_var.get()
            click_count = count_var.get()

            result[0] = {
                'coord_id': coord_id,
                'click_type': click_type,
                'click_count': click_count,
                'pre_delay': 0.2,
                'post_delay': 0.2
            }
            dialog.destroy()
        
        tk.Button(
            dialog,
            text="확인",
            command=on_ok,
            bg='#3498db',
            fg='white',
            padx=20,
            pady=5
        ).pack(pady=20)
        
        # 중앙 배치 및 포커스
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        parent_x = self.winfo_x()
        parent_y = self.winfo_y()
        parent_width = self.winfo_width()
        parent_height = self.winfo_height()
        x = parent_x + (parent_width - width) // 2
        y = parent_y + (parent_height - height) // 2
        dialog.geometry(f'{width}x{height}+{x}+{y}')
        
        dialog.lift()
        dialog.focus_force()

        self.wait_window(dialog)
        return result[0]

    def config_mouse_scroll(self):
        """마우스 스크롤 설정"""
        dialog = tk.Toplevel(self)
        dialog.title("마우스 스크롤 설정")
        dialog.geometry("250x140")
        dialog.transient(self)
        dialog.grab_set()
        dialog.attributes('-topmost', True)
        set_dialog_icon(dialog)

        result = [None]

        # 스크롤 방향 (한 줄)
        direction_frame = tk.Frame(dialog)
        direction_frame.pack(fill='x', padx=20, pady=(10, 10))

        tk.Label(
            direction_frame,
            text="스크롤 방향:",
            font=("맑은 고딕", 10, "bold"),
            width=10,
            anchor='w'
        ).pack(side='left')

        direction_var = tk.StringVar(value='down')
        tk.Radiobutton(direction_frame, text="위로", variable=direction_var, value='up').pack(side='left', padx=5)
        tk.Radiobutton(direction_frame, text="아래로", variable=direction_var, value='down').pack(side='left', padx=5)

        # 스크롤 양 (한 줄)
        amount_frame = tk.Frame(dialog)
        amount_frame.pack(fill='x', padx=20, pady=(0, 10))

        tk.Label(
            amount_frame,
            text="스크롤 양:",
            font=("맑은 고딕", 10, "bold"),
            width=10,
            anchor='w'
        ).pack(side='left')

        amount_var = tk.IntVar(value=3)
        tk.Spinbox(
            amount_frame,
            from_=1,
            to=20,
            textvariable=amount_var,
            font=("맑은 고딕", 10),
            width=10
        ).pack(side='left')

        def on_ok():
            direction = direction_var.get()
            amount = amount_var.get()

            result[0] = {
                'direction': direction,
                'amount': amount
            }
            dialog.destroy()

        tk.Button(
            dialog,
            text="확인",
            command=on_ok,
            bg='#3498db',
            fg='white',
            padx=20,
            pady=5
        ).pack(pady=(10))

        # 중앙 배치 및 포커스
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        parent_x = self.winfo_x()
        parent_y = self.winfo_y()
        parent_width = self.winfo_width()
        parent_height = self.winfo_height()
        x = parent_x + (parent_width - width) // 2
        y = parent_y + (parent_height - height) // 2
        dialog.geometry(f'{width}x{height}+{x}+{y}')

        dialog.lift()
        dialog.focus_force()

        self.wait_window(dialog)
        return result[0]


    def config_click_image(self):
        """이미지 클릭 설정"""
        if not self.image_mgr.images:
            self._show_error_dialog("등록된 이미지가 없습니다.")
            return None
        
        dialog = tk.Toplevel(self)
        dialog.title("이미지 클릭 설정")
        dialog.geometry("300x150")
        dialog.transient(self)
        dialog.grab_set()
        dialog.attributes('-topmost', True)
        set_dialog_icon(dialog)
        
        result = [None]
        
        tk.Label(
            dialog,
            text="이미지 선택:",
            font=("맑은 고딕", 10, "bold")
        ).pack(anchor='w', padx=20, pady=(20, 5))
        
        image_var = tk.StringVar()
        image_combo = ttk.Combobox(
            dialog,
            textvariable=image_var,
            font=("맑은 고딕", 10),
            state='readonly'
        )
        image_values = [f"{img['id']}. {img['name']}" for img in self.image_mgr.images]
        image_combo['values'] = image_values
        image_combo.current(0)
        image_combo.pack(fill='x', padx=20, pady=(0, 15))
        
        def on_ok():
            selected_idx = image_combo.current()
            image_id = self.image_mgr.images[selected_idx]['id']
            
            result[0] = {
                'image_id': image_id
            }
            dialog.destroy()
        
        tk.Button(
            dialog,
            text="확인",
            command=on_ok,
            bg='#3498db',
            fg='white',
            padx=20,
            pady=5
        ).pack(pady=20)
        
        # 중앙 배치 및 포커스
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        parent_x = self.winfo_x()
        parent_y = self.winfo_y()
        parent_width = self.winfo_width()
        parent_height = self.winfo_height()
        x = parent_x + (parent_width - width) // 2
        y = parent_y + (parent_height - height) // 2
        dialog.geometry(f'{width}x{height}+{x}+{y}')
        
        dialog.lift()
        dialog.focus_force()
        
        self.wait_window(dialog)
        return result[0]


        
    def config_type_text(self):
        """텍스트 타이핑 설정"""
        dialog = tk.Toplevel(self)
        dialog.title("텍스트 입력")
        dialog.geometry("300x180")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        dialog.attributes('-topmost', True)
        set_dialog_icon(dialog)
        
        result = [None]
        
        tk.Label(
            dialog,
            text="입력할 텍스트",
            font=("맑은 고딕", 11, "bold")
        ).pack(pady=15)
        
        text_entry = tk.Entry(dialog, font=("맑은 고딕", 11), width=40)
        text_entry.pack(padx=20, pady=10, fill='x')
        text_entry.focus()
        
        def on_ok():
            text = text_entry.get().strip()
            if not text:
                messagebox.showwarning("경고", "텍스트를 입력하세요.", parent=dialog)
                return
            result[0] = {'text': text, 'interval': 0.05}
            dialog.destroy()
        
        def on_cancel():
            dialog.destroy()

        tk.Button(
            dialog,
            text="확인",
            font=("맑은 고딕", 10),
            bg='#3498db',
            fg='white',
            padx=20,
            pady=8,
            command=on_ok
        ).pack(pady=15)

        text_entry.bind('<Return>', lambda e: on_ok())
        text_entry.bind('<Escape>', lambda e: on_cancel())
        
        # 중앙 배치 및 포커스
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        parent_x = self.winfo_x()
        parent_y = self.winfo_y()
        parent_width = self.winfo_width()
        parent_height = self.winfo_height()
        x = parent_x + (parent_width - width) // 2
        y = parent_y + (parent_height - height) // 2
        dialog.geometry(f'{width}x{height}+{x}+{y}')
        
        dialog.lift()
        dialog.focus_force()
        
        self.wait_window(dialog)
        return result[0]

        
    def config_type_variable(self):
        """변수 타이핑 설정"""
        if not self.excel_mgr.excel_sources:
            self._show_error_dialog("등록된 엑셀 데이터가 없습니다.")
            return None

        dialog = tk.Toplevel(self)
        dialog.title("변수 타이핑 설정")
        dialog.geometry("400x430")
        dialog.transient(self)
        dialog.grab_set()
        dialog.attributes('-topmost', True)
        set_dialog_icon(dialog)

        result = [None]

        tk.Label(
            dialog,
            text="변수 유형:",
            font=("맑은 고딕", 10, "bold")
        ).pack(anchor='w', padx=20, pady=(20, 5))

        var_type = tk.StringVar(value='excel')
        radio_frame = tk.Frame(dialog)
        radio_frame.pack(anchor='w', padx=40, pady=5)

        tk.Radiobutton(radio_frame, text="엑셀 데이터", variable=var_type, value='excel').pack(side='left', padx=5)
        tk.Radiobutton(radio_frame, text="현재 행 번호", variable=var_type, value='counter').pack(side='left', padx=5)
        tk.Radiobutton(radio_frame, text="타임스탬프", variable=var_type, value='timestamp').pack(side='left', padx=5)

        # 엑셀 소스 선택 (다중 엑셀 지원)
        tk.Label(
            dialog,
            text="엑셀 소스 선택:",
            font=("맑은 고딕", 10, "bold")
        ).pack(anchor='w', padx=20, pady=(15, 5))

        excel_var = tk.StringVar()
        excel_combo = ttk.Combobox(
            dialog,
            textvariable=excel_var,
            font=("맑은 고딕", 10),
            state='readonly'
        )
        excel_values = [f"{idx+1}. {src['name']}" for idx, src in enumerate(self.excel_mgr.excel_sources)]
        excel_combo['values'] = excel_values
        excel_combo.current(0)
        excel_combo.pack(fill='x', padx=20, pady=(0, 15))

        tk.Label(
            dialog,
            text="칼럼 선택 (엑셀 데이터인 경우):",
            font=("맑은 고딕", 10, "bold")
        ).pack(anchor='w', padx=20, pady=(0, 5))

        # 칼럼 선택
        column_var = tk.StringVar()
        column_combo = ttk.Combobox(
            dialog,
            textvariable=column_var,
            font=("맑은 고딕", 10),
            state='readonly'
        )
        column_combo.pack(fill='x', padx=20, pady=(0, 15))

        # 데이터 처리 옵션
        options_frame = tk.LabelFrame(
            dialog,
            text="데이터 처리 옵션",
            font=("맑은 고딕", 10, "bold"),
            padx=15,
            pady=10
        )
        options_frame.pack(fill='x', padx=20, pady=(0, 15))

        # 공백 행 제거 옵션
        remove_empty_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            options_frame,
            text="🧹 상위 공백 행 제거",
            variable=remove_empty_var,
            font=("맑은 고딕", 9)
        ).pack(anchor='w', pady=2)

        # 중복 행 제거 옵션
        remove_dup_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            options_frame,
            text="🔄 중복 행 제거",
            variable=remove_dup_var,
            font=("맑은 고딕", 9)
        ).pack(anchor='w', pady=2)

        tk.Label(
            options_frame,
            text="※ 옵션 변경 시 선택한 엑셀 소스에 적용됩니다",
            font=("맑은 고딕", 8),
            fg='#7f8c8d'
        ).pack(anchor='w', pady=(5, 0))

        # 엑셀 소스 변경 시 칼럼 목록 및 옵션 업데이트
        def on_excel_change(event=None):
            selected_idx = excel_combo.current()
            if selected_idx >= 0:
                excel_source = self.excel_mgr.excel_sources[selected_idx]
                column_combo['values'] = excel_source['columns']
                if excel_source['columns']:
                    column_combo.current(0)

                # 현재 엑셀 소스의 옵션값 로드
                remove_empty_var.set(excel_source.get('remove_empty_rows', True))
                remove_dup_var.set(excel_source.get('remove_duplicates', False))

        excel_combo.bind('<<ComboboxSelected>>', on_excel_change)

        # 초기 칼럼 목록 및 옵션 설정
        on_excel_change()

        def on_ok():
            vtype = var_type.get()
            selected_excel_idx = excel_combo.current()
            excel_source = self.excel_mgr.excel_sources[selected_excel_idx]
            excel_id = excel_source.get('id', selected_excel_idx + 1)
            excel_name = excel_source.get('name')
            var_name = column_combo.get() if vtype == 'excel' else ''

            # 엑셀 소스의 데이터 처리 옵션 업데이트
            excel_source['remove_empty_rows'] = remove_empty_var.get()
            excel_source['remove_duplicates'] = remove_dup_var.get()

            result[0] = {
                'var_type': vtype,
                'var_name': var_name,
                'excel_id': excel_id,  # 엑셀 ID 추가
                'excel_name': excel_name,
                'interval': 0.05
            }
            dialog.destroy()

        tk.Button(
            dialog,
            text="확인",
            command=on_ok,
            bg='#3498db',
            fg='white',
            padx=20,
            pady=5
        ).pack(pady=10)

        # 중앙 배치 및 포커스
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        parent_x = self.winfo_x()
        parent_y = self.winfo_y()
        parent_width = self.winfo_width()
        parent_height = self.winfo_height()
        x = parent_x + (parent_width - width) // 2
        y = parent_y + (parent_height - height) // 2
        dialog.geometry(f'{width}x{height}+{x}+{y}')

        dialog.lift()
        dialog.focus_force()

        self.wait_window(dialog)
        return result[0]


    
    def config_key_press(self):
        """키 입력 설정 (자동 감지)"""
        dialog = KeyInputDialog(self)
        self.wait_window(dialog)
        
        if dialog.result:
            return {'key': dialog.result}
        return None

    def config_delay(self):
        """딜레이 설정"""
        dialog = tk.Toplevel(self)
        dialog.title("딜레이 설정")
        dialog.geometry("300x180")
        dialog.transient(self)
        dialog.grab_set()
        dialog.attributes('-topmost', True)
        set_dialog_icon(dialog)
        
        result = [None]
        
        tk.Label(
            dialog,
            text="대기 시간 설정",
            font=("맑은 고딕", 12, "bold")
        ).pack(pady=10)
        
        tk.Label(
            dialog,
            text="대기시간 입력 : 0.1 ~ 3600(초)",
            font=("맑은 고딕", 10)
        ).pack(pady=5)
        
        entry = tk.Entry(dialog, font=("맑은 고딕", 11), width=15)
        entry.insert(0, "1.0")
        entry.pack(pady=10)
        entry.focus()
        entry.select_range(0, tk.END)
        
        def on_ok():
            try:
                seconds = float(entry.get())
                if 0.1 <= seconds <= 3600:
                    result[0] = {'seconds': seconds}
                    dialog.destroy()
                else:
                    self._show_error_dialog("0.1초 ~ 3600초 사이의 값을 입력하세요.")
            except ValueError:
                self._show_error_dialog("올바른 숫자를 입력하세요.")
        
        tk.Button(
            dialog,
            text="확인",
            command=on_ok,
            bg='#3498db',
            fg='white',
            padx=20,
            pady=5
        ).pack(pady=10)
        
        entry.bind('<Return>', lambda e: on_ok())
        
        # 중앙 배치 및 포커스
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        parent_x = self.winfo_x()
        parent_y = self.winfo_y()
        parent_width = self.winfo_width()
        parent_height = self.winfo_height()
        x = parent_x + (parent_width - width) // 2
        y = parent_y + (parent_height - height) // 2
        dialog.geometry(f'{width}x{height}+{x}+{y}')
        
        dialog.lift()
        dialog.focus_force()
        
        self.wait_window(dialog)
        return result[0]


    def config_wait_image(self):
        """이미지 대기 설정"""
        if not self.image_mgr.images:
            self._show_error_dialog("등록된 이미지가 없습니다.")
            return None
        
        dialog = tk.Toplevel(self)
        dialog.title("이미지 대기 설정")
        dialog.geometry("300x230")
        dialog.transient(self)
        dialog.grab_set()
        dialog.attributes('-topmost', True)
        set_dialog_icon(dialog)
        
        result = [None]
        
        tk.Label(
            dialog,
            text="이미지 선택:",
            font=("맑은 고딕", 10, "bold")
        ).pack(anchor='w', padx=20, pady=(20, 5))
        
        image_var = tk.StringVar()
        image_combo = ttk.Combobox(
            dialog,
            textvariable=image_var,
            font=("맑은 고딕", 10),
            state='readonly'
        )
        image_values = [f"{img['id']}. {img['name']}" for img in self.image_mgr.images]
        image_combo['values'] = image_values
        image_combo.current(0)
        image_combo.pack(fill='x', padx=20, pady=(0, 15))
        
        tk.Label(
            dialog,
            text="최대 대기 시간 (초):",
            font=("맑은 고딕", 10, "bold")
        ).pack(anchor='w', padx=20, pady=(10, 5))
        
        timeout_entry = tk.Entry(dialog, font=("맑은 고딕", 10))
        timeout_entry.insert(0, "10")
        timeout_entry.pack(fill='x', padx=20, pady=(0, 15))
        
        def on_ok():
            selected_idx = image_combo.current()
            image_id = self.image_mgr.images[selected_idx]['id']
            timeout = float(timeout_entry.get())
            
            result[0] = {
                'image_id': image_id,
                'timeout': timeout
            }
            dialog.destroy()
        
        tk.Button(
            dialog,
            text="확인",
            command=on_ok,
            bg='#3498db',
            fg='white',
            padx=20,
            pady=5
        ).pack(pady=20)
        
        # 중앙 배치 및 포커스
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        parent_x = self.winfo_x()
        parent_y = self.winfo_y()
        parent_width = self.winfo_width()
        parent_height = self.winfo_height()
        x = parent_x + (parent_width - width) // 2
        y = parent_y + (parent_height - height) // 2
        dialog.geometry(f'{width}x{height}+{x}+{y}')
        
        dialog.lift()
        dialog.focus_force()
        
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
        if mode == 'region':
            region_selector = ScreenRegionSelectorDialog(self)
            self.wait_window(region_selector)

            region_data = region_selector.result
            if not region_data:  # 취소
                return None

        # Step 3: 파일명 입력
        dialog = tk.Toplevel(self)
        dialog.title("스크린샷")
        dialog.geometry("300x200")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        dialog.attributes('-topmost', True)
        set_dialog_icon(dialog)

        result = [None]

        tk.Label(
            dialog,
            text="파일명 (자동으로 타임스탬프 추가):",
            font=("맑은 고딕", 11, "bold")
        ).pack(pady=20)

        filename_entry = tk.Entry(dialog, font=("맑은 고딕", 11), width=40)
        filename_entry.insert(0, "screenshot")
        filename_entry.pack(padx=20, pady=10, fill='x')
        filename_entry.select_range(0, tk.END)
        filename_entry.focus()

        def on_ok():
            filename = filename_entry.get().strip()
            if not filename:
                messagebox.showwarning("경고", "파일명을 입력하세요.", parent=dialog)
                return

            # 파일명과 영역 데이터를 함께 반환
            result[0] = {
                'filename': filename,
                'mode': mode,
                'region': region_data  # None (전체) 또는 {'x', 'y', 'width', 'height'}
            }
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        tk.Button(
            dialog,
            text="확인",
            font=("맑은 고딕", 10),
            bg='#3498db',
            fg='white',
            padx=20,
            pady=8,
            command=on_ok
        ).pack(pady=15)

        filename_entry.bind('<Return>', lambda e: on_ok())
        filename_entry.bind('<Escape>', lambda e: on_cancel())

        # 중앙 배치 및 포커스
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        parent_x = self.winfo_x()
        parent_y = self.winfo_y()
        parent_width = self.winfo_width()
        parent_height = self.winfo_height()
        x = parent_x + (parent_width - width) // 2
        y = parent_y + (parent_height - height) // 2
        dialog.geometry(f'{width}x{height}+{x}+{y}')

        dialog.lift()
        dialog.focus_force()

        self.wait_window(dialog)
        return result[0]


class ScreenshotRegionModeDialog(tk.Toplevel):
    """스크린샷 영역 모드 선택 다이얼로그"""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("스크린샷 영역 설정")
        self.geometry("300x180")
        self.resizable(False, False)

        self.result = None  # 'full' or 'region'

        # 모달 설정
        self.transient(parent)
        self.grab_set()
        self.attributes('-topmost', True)
        set_dialog_icon(self)

        self.setup_ui()
        center_window_on_parent(self, parent)

        # 포커스
        self.lift()
        self.focus_force()

    def setup_ui(self):
        """UI 구성"""
        main_frame = tk.Frame(self, padx=20, pady=20)
        main_frame.pack(fill='both', expand=True)

        # 제목
        tk.Label(
            main_frame,
            text="📸 스크린샷 영역 설정",
            font=("맑은 고딕", 12, "bold")
        ).pack(pady=(0, 15))

        # 설명
        tk.Label(
            main_frame,
            text="스크린샷 영역을 선택하세요.",
            font=("맑은 고딕", 9),
            fg='#7f8c8d'
        ).pack(pady=(0, 15))

        # 버튼 프레임
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(fill='x')

        # 전체 화면 버튼
        tk.Button(
            btn_frame,
            text="🖥️ 전체 화면",
            font=("맑은 고딕", 11),
            bg='#3498db',
            fg='white',
            padx=20,
            pady=10,
            command=self.select_full
        ).pack(side='left', expand=True, fill='x', padx=(0, 5))

        # 영역 선택 버튼
        tk.Button(
            btn_frame,
            text="🎯 영역 선택",
            font=("맑은 고딕", 11),
            bg='#27ae60',
            fg='white',
            padx=20,
            pady=10,
            command=self.select_region
        ).pack(side='left', expand=True, fill='x', padx=(5, 0))

    def select_full(self):
        """전체 화면 선택"""
        self.result = 'full'
        self.destroy()

    def select_region(self):
        """영역 선택"""
        self.result = 'region'
        self.destroy()


class ScreenRegionSelectorDialog(tk.Toplevel):
    """화면 영역 선택 오버레이"""

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
        self.attributes('-topmost', True)
        self.attributes('-alpha', 0.3)  # 반투명
        self.configure(bg='gray')
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
            bg='gray',
            highlightthickness=0,
            cursor='crosshair'
        )
        self.canvas.pack(fill='both', expand=True)

        # 안내 텍스트
        self.canvas.create_text(
            screen_width // 2,
            30,
            text="드래그하여 영역을 선택하세요 (ESC: 취소)",
            font=("맑은 고딕", 14, "bold"),
            fill='white'
        )

        # 이벤트 바인딩
        self.canvas.bind('<Button-1>', self.on_mouse_down)
        self.canvas.bind('<B1-Motion>', self.on_mouse_move)
        self.canvas.bind('<ButtonRelease-1>', self.on_mouse_up)
        self.bind('<Escape>', lambda e: self.cancel())

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
            self.start_x, self.start_y, event.x, event.y,
            outline='red',
            width=2,
            fill=''
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
        self.result = {
            'x': x1,
            'y': y1,
            'width': width,
            'height': height
        }

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


class AutostartSelectDialog(tk.Toplevel):
    """자동시작 프로젝트 선택 다이얼로그"""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("자동시작 설정")
        self.geometry("250x180")
        self.resizable(False, False)

        self.result = None

        # 모달 설정
        self.transient(parent)
        self.grab_set()
        self.attributes('-topmost', True)
        set_dialog_icon(self)

        self.setup_ui()
        center_window_on_parent(self, parent)

        # 포커스
        self.lift()
        self.focus_force()

    def setup_ui(self):
        """UI 구성"""
        main_frame = tk.Frame(self, padx=20, pady=20)
        main_frame.pack(fill='both', expand=True)

        # 제목
        tk.Label(
            main_frame,
            text="🚀 자동시작 프로젝트 설정",
            font=("맑은 고딕", 13, "bold")
        ).pack(pady=(0, 10))

        # 설명
        tk.Label(
            main_frame,
            text="자동실행할 프로젝트를 선택하세요.",
            font=("맑은 고딕", 9),
            fg='#7f8c8d'
        ).pack(pady=(0, 15))

        # 콤보박스 프레임
        combo_frame = tk.Frame(main_frame)
        combo_frame.pack(fill='x', pady=(0, 15))

        self.project_var = tk.StringVar()
        self.project_combo = ttk.Combobox(
            combo_frame,
            textvariable=self.project_var,
            font=("맑은 고딕", 10),
            state='readonly'
        )
        self.project_combo.pack(side='left', fill='x', expand=True)

        # 프로젝트 목록 로드
        self.load_projects()

        # 버튼 프레임
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(fill='x')

        # 확인 버튼
        tk.Button(
            btn_frame,
            text="확인",
            font=("맑은 고딕", 10),
            bg='#3498db',
            fg='white',
            padx=20,
            pady=8,
            command=self.on_ok
        ).pack(side='left', padx=(0, 5), expand=True, fill='x')

        # 취소 버튼
        tk.Button(
            btn_frame,
            text="취소",
            font=("맑은 고딕", 10),
            bg='#95a5a6',
            fg='white',
            padx=20,
            pady=8,
            command=self.on_cancel
        ).pack(side='left', padx=(5, 0), expand=True, fill='x')

    def load_projects(self):
        """프로젝트 및 체인 목록 로드"""
        from core.project_manager import ProjectManager
        import json

        projects = ProjectManager.get_project_list()
        self.projects = projects

        # 현재 자동시작 설정 읽기
        current_autostart = None
        try:
            with open('settings.json', 'r', encoding='utf-8') as f:
                settings = json.load(f)
                current_autostart = settings.get('autostart')
        except:
            pass

        # 콤보박스 항목 리스트
        combo_items = ["없음"]
        selected_index = 0

        if projects:
            for idx, project in enumerate(projects):
                # 프로젝트와 체인 모두 표시
                item_type = project.get('type', 'project')
                icon = "🔗" if item_type == 'chain' else "📁"
                display_text = f"{icon} {project['name']}"

                combo_items.append(display_text)

                # 현재 자동시작 프로젝트인지 확인
                if current_autostart and project['filepath'] == current_autostart:
                    selected_index = idx + 1  # "없음" 다음 인덱스

        # 콤보박스 설정
        self.project_combo['values'] = combo_items
        self.project_combo.current(selected_index)

    def on_ok(self):
        """확인 버튼"""
        selected_index = self.project_combo.current()

        # "없음" 선택
        if selected_index == 0:
            self.result = {'filepath': None, 'name': None}
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
            'filepath': selected_item['filepath'],
            'name': selected_item['name'],
            'type': selected_item.get('type', 'project')
        }
        self.destroy()

    def on_cancel(self):
        """취소 버튼"""
        self.result = None
        self.destroy()
