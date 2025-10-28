"""
각종 다이얼로그
"""
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog 
from datetime import datetime


class NewProjectDialog(tk.Toplevel):
    """새 프로젝트 생성 다이얼로그"""
    
    def __init__(self, parent):
        super().__init__(parent)
        self.title("새 프로젝트")
        self.geometry("400x250")
        self.resizable(False, False)
        
        self.result = None
        
        # 모달 설정
        self.transient(parent)
        self.grab_set()
        
        self.setup_ui()
        
        # 창 중앙 배치
        self.center_window()
    
    def center_window(self):
        """창을 화면 중앙에 배치"""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')
    
    def setup_ui(self):
        """UI 구성"""
        # 메인 프레임
        main_frame = tk.Frame(self, padx=20, pady=20)
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
            width=40
        )
        self.name_entry.pack(fill='x', pady=(0, 15))
        self.name_entry.focus()
        
        # 설명
        tk.Label(
            main_frame,
            text="설명 (선택사항):",
            font=("맑은 고딕", 10)
        ).pack(anchor='w', pady=(0, 5))
        
        self.desc_text = tk.Text(
            main_frame,
            font=("맑은 고딕", 10),
            height=4,
            width=40
        )
        self.desc_text.pack(fill='x', pady=(0, 20))
        
        # 버튼
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack()
        
        tk.Button(
            btn_frame,
            text="생성",
            font=("맑은 고딕", 10),
            bg='#3498db',
            fg='white',
            padx=20,
            pady=5,
            command=self.on_create
        ).pack(side='left', padx=5)
        
        tk.Button(
            btn_frame,
            text="취소",
            font=("맑은 고딕", 10),
            bg='#95a5a6',
            fg='white',
            padx=20,
            pady=5,
            command=self.on_cancel
        ).pack(side='left', padx=5)
        
        # Enter 키 바인딩
        self.name_entry.bind('<Return>', lambda e: self.on_create())
    
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
        
        description = self.desc_text.get('1.0', 'end-1c').strip()
        
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
        self.geometry("400x180")
        self.resizable(False, False)
        
        self.result = None
        
        # 모달 설정
        self.transient(parent)
        self.grab_set()
        
        # 항상 위에 표시
        self.attributes('-topmost', True)
        
        self.setup_ui(message, initial_value)
        self.center_window()
        
        # 포커스
        self.lift()
        self.focus_force()
    
    def center_window(self):
        """창을 부모 중앙에 배치"""
        self.update_idletasks()
        
        # 부모 창의 위치와 크기
        parent_x = self.master.winfo_x()
        parent_y = self.master.winfo_y()
        parent_width = self.master.winfo_width()
        parent_height = self.master.winfo_height()
        
        # 다이얼로그 크기
        dialog_width = self.winfo_width()
        dialog_height = self.winfo_height()
        
        # 중앙 위치 계산
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        
        self.geometry(f'{dialog_width}x{dialog_height}+{x}+{y}')
    
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
        
        if initial_value:
            self.entry.insert(0, initial_value)
            self.entry.select_range(0, tk.END)
        
        self.entry.focus_set()
        
        # 버튼
        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=20)
        
        tk.Button(
            btn_frame,
            text="확인",
            font=("맑은 고딕", 10),
            bg='#27ae60',
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
        self.entry.bind('<Return>', lambda e: self.on_ok())
        self.entry.bind('<Escape>', lambda e: self.on_cancel())
    
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

class ActionSelectDialog(tk.Toplevel):
    """액션 선택 다이얼로그"""
    
    def __init__(self, parent, coord_mgr, excel_mgr, image_mgr):
        super().__init__(parent)
        self.title("액션 추가")
        self.geometry("600x500")
        self.resizable(False, False)
        
        self.coord_mgr = coord_mgr
        self.excel_mgr = excel_mgr
        self.image_mgr = image_mgr
        
        self.result = None
        
        self.transient(parent)
        self.grab_set()
        
        self.setup_ui()
        self.center_window()
    
    def center_window(self):
        """창 중앙 배치"""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')
    
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
        ).grid(row=1, column=0, padx=5, pady=5)
        
        tk.Button(
            btn_frame,
            text="단축키",
            font=("맑은 고딕", 9),
            width=12,
            command=lambda: self.select_action('hotkey')
        ).grid(row=1, column=1, padx=5, pady=5)
        
        tk.Button(
            btn_frame,
            text="붙여넣기",
            font=("맑은 고딕", 9),
            width=12,
            command=lambda: self.select_action('paste')
        ).grid(row=2, column=0, padx=5, pady=5)
        
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
        
        tk.Button(
            btn_frame,
            text="메모",
            font=("맑은 고딕", 9),
            width=12,
            command=lambda: self.select_action('memo')
        ).grid(row=0, column=1, padx=5, pady=5)
    
    def select_action(self, action_type):
        """액션 선택"""
        params = None
        
        if action_type == 'click_coord':
            params = self.config_click_coord()
        elif action_type == 'click_image':
            params = self.config_click_image()
        elif action_type == 'type_text':
            params = self.config_type_text()
        elif action_type == 'type_variable':
            params = self.config_type_variable()
        elif action_type == 'key_press':
            params = self.config_key_press()
        elif action_type == 'hotkey':
            params = self.config_hotkey()
        elif action_type == 'paste':
            params = {}
        elif action_type == 'delay':
            params = self.config_delay()
        elif action_type == 'wait_image':
            params = self.config_wait_image()
        elif action_type == 'screenshot':
            params = self.config_screenshot()
        elif action_type == 'memo':
            params = self.config_memo()
        
        if params is not None:
            self.result = {
                'type': action_type,
                'params': params
            }
            self.destroy()
    
    def config_click_coord(self):
        """좌표 클릭 설정"""
        if not self.coord_mgr.coordinates:
            messagebox.showerror("오류", "등록된 좌표가 없습니다.")
            return None
        
        dialog = tk.Toplevel(self)
        dialog.title("좌표 클릭 설정")
        dialog.geometry("400x350")
        dialog.transient(self)
        dialog.grab_set()
        
        result = [None]
        
        tk.Label(
            dialog,
            text="좌표 선택:",
            font=("맑은 고딕", 10, "bold")
        ).pack(anchor='w', padx=20, pady=(20, 5))
        
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
        coord_combo.pack(fill='x', padx=20, pady=(0, 15))
        
        tk.Label(
            dialog,
            text="클릭 유형:",
            font=("맑은 고딕", 10, "bold")
        ).pack(anchor='w', padx=20, pady=(10, 5))
        
        click_var = tk.StringVar(value='left')
        tk.Radiobutton(dialog, text="좌클릭", variable=click_var, value='left').pack(anchor='w', padx=40)
        tk.Radiobutton(dialog, text="우클릭", variable=click_var, value='right').pack(anchor='w', padx=40)
        tk.Radiobutton(dialog, text="더블클릭", variable=click_var, value='double').pack(anchor='w', padx=40)
        
        def on_ok():
            selected_idx = coord_combo.current()
            coord_id = self.coord_mgr.coordinates[selected_idx]['id']
            
            click_type = click_var.get()
            click_count = 2 if click_type == 'double' else 1
            
            result[0] = {
                'coord_id': coord_id,
                'click_type': 'left' if click_type == 'double' else click_type,
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
        
        dialog.wait_window()
        return result[0]
    
    def config_click_image(self):
        """이미지 클릭 설정"""
        if not self.image_mgr.images:
            messagebox.showerror("오류", "등록된 이미지가 없습니다.")
            return None
        
        dialog = tk.Toplevel(self)
        dialog.title("이미지 클릭 설정")
        dialog.geometry("400x200")
        dialog.transient(self)
        dialog.grab_set()
        
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
        
        dialog.wait_window()
        return result[0]
    
    def config_type_text(self):
        """텍스트 타이핑 설정"""
        text = simpledialog.askstring("텍스트 입력", "입력할 텍스트:")
        if text:
            return {'text': text, 'interval': 0.05}
        return None
    
    def config_type_variable(self):
        """변수 타이핑 설정"""
        if not self.excel_mgr.excel_sources:
            messagebox.showerror("오류", "등록된 엑셀 데이터가 없습니다.")
            return None
        
        dialog = tk.Toplevel(self)
        dialog.title("변수 타이핑 설정")
        dialog.geometry("400x300")
        dialog.transient(self)
        dialog.grab_set()
        
        result = [None]
        
        tk.Label(
            dialog,
            text="변수 유형:",
            font=("맑은 고딕", 10, "bold")
        ).pack(anchor='w', padx=20, pady=(20, 5))
        
        var_type = tk.StringVar(value='excel')
        tk.Radiobutton(dialog, text="엑셀 데이터", variable=var_type, value='excel').pack(anchor='w', padx=40)
        tk.Radiobutton(dialog, text="현재 행 번호", variable=var_type, value='counter').pack(anchor='w', padx=40)
        tk.Radiobutton(dialog, text="타임스탬프", variable=var_type, value='timestamp').pack(anchor='w', padx=40)
        
        tk.Label(
            dialog,
            text="칼럼 선택 (엑셀 데이터인 경우):",
            font=("맑은 고딕", 10, "bold")
        ).pack(anchor='w', padx=20, pady=(15, 5))
        
        # 칼럼 선택
        excel_source = self.excel_mgr.excel_sources[0]
        column_var = tk.StringVar()
        column_combo = ttk.Combobox(
            dialog,
            textvariable=column_var,
            font=("맑은 고딕", 10),
            state='readonly'
        )
        column_combo['values'] = excel_source['columns']
        if excel_source['columns']:
            column_combo.current(0)
        column_combo.pack(fill='x', padx=20, pady=(0, 15))
        
        def on_ok():
            vtype = var_type.get()
            var_name = column_combo.get() if vtype == 'excel' else ''
            
            result[0] = {
                'var_type': vtype,
                'var_name': var_name,
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
        ).pack(pady=20)
        
        dialog.wait_window()
        return result[0]
    
    def config_key_press(self):
        """키 입력 설정"""
        keys = ['enter', 'tab', 'esc', 'space', 'backspace', 'delete', 
                'up', 'down', 'left', 'right', 'home', 'end', 'pageup', 'pagedown']
        
        key = simpledialog.askstring(
            "키 입력",
            f"입력할 키:\n(예: {', '.join(keys[:5])}...)"
        )
        if key:
            return {'key': key.lower()}
        return None
    
    def config_hotkey(self):
        """단축키 설정"""
        hotkey = simpledialog.askstring(
            "단축키",
            "단축키 입력 (예: ctrl+c, ctrl+v, alt+f4):"
        )
        if hotkey:
            keys = [k.strip() for k in hotkey.split('+')]
            return {'keys': keys}
        return None
    
    def config_delay(self):
        """딜레이 설정"""
        dialog = tk.Toplevel(self)
        dialog.title("딜레이 설정")
        dialog.geometry("300x180")
        dialog.transient(self)
        dialog.grab_set()
        
        result = [None]
        
        tk.Label(
            dialog,
            text="대기 시간 설정",
            font=("맑은 고딕", 12, "bold")
        ).pack(pady=15)
        
        tk.Label(
            dialog,
            text="대기 시간 (초):",
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
                if 0.1 <= seconds <= 300:
                    result[0] = {'seconds': seconds}
                    dialog.destroy()
                else:
                    messagebox.showerror("오류", "0.1초 ~ 300초 사이의 값을 입력하세요.", parent=dialog)
            except ValueError:
                messagebox.showerror("오류", "올바른 숫자를 입력하세요.", parent=dialog)
        
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
        
        dialog.wait_window()
        return result[0]
    
    def config_wait_image(self):
        """이미지 대기 설정"""
        if not self.image_mgr.images:
            messagebox.showerror("오류", "등록된 이미지가 없습니다.")
            return None
        
        dialog = tk.Toplevel(self)
        dialog.title("이미지 대기 설정")
        dialog.geometry("400x250")
        dialog.transient(self)
        dialog.grab_set()
        
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
        
        dialog.wait_window()
        return result[0]
    
    def config_screenshot(self):
        """스크린샷 설정"""
        filename = simpledialog.askstring(
            "스크린샷",
            "파일명 (자동으로 타임스탬프 추가):",
            initialvalue="screenshot"
        )
        if filename:
            return {'filename': f"{filename}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"}
        return None
    
    def config_memo(self):
        """메모 설정"""
        text = simpledialog.askstring("메모", "메모 내용:")
        if text:
            return {'text': text}
        return None
