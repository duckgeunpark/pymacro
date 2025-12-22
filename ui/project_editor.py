"""
프로젝트 편집 화면 - 좌표/엑셀/이미지/플로우 관리
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from datetime import datetime
import os
from utils.ui_helpers import set_dialog_icon, center_window_on_parent

from core.project_manager import ProjectManager
from core.coordinate_manager import CoordinateManager
from core.excel_manager import ExcelManager
from core.image_manager import ImageManager
from core.flow_manager import FlowManager
from ui.dialogs import ActionSelectDialog, NameInputDialog



class ProjectEditor(tk.Frame):
    def __init__(self, parent, app, project_data, filepath):
        super().__init__(parent)
        self.app = app
        self.parent = parent
        self.project_data = project_data
        self.filepath = filepath
        
        # 관리자 초기화
        self.coord_mgr = CoordinateManager()
        self.coord_mgr.load_from_list(project_data.get('coordinates', []))
        
        self.excel_mgr = ExcelManager()
        self.excel_mgr.load_from_list(project_data.get('excel_sources', []))
        
        self.image_mgr = ImageManager()
        self.image_mgr.load_from_list(project_data.get('images', []))
        
        self.flow_mgr = FlowManager()
        self.flow_mgr.load_from_list(project_data.get('flow_sequence', []))

        # 드래그 앤 드롭 관련 변수
        self.drag_data = {"item": None, "index": None, "y": 0}

        self.setup_ui()
    
    def setup_ui(self):
        """UI 구성"""
        # 상단 헤더
        header = tk.Frame(self, bg='#34495e', height=60)
        header.pack(fill='x')
        header.pack_propagate(False)
        
        # 프로젝트 이름
        tk.Label(
            header,
            text=f"📝 {self.project_data['name']}",
            font=("맑은 고딕", 14, "bold"),
            bg='#34495e',
            fg='white'
        ).pack(side='left', padx=10, pady=15)
        
        # 헤더 버튼들
        btn_frame = tk.Frame(header, bg='#34495e')
        btn_frame.pack(side='right', padx=10)
        
        tk.Button(
            btn_frame,
            text="✅ 완료",
            font=("맑은 고딕", 10),
            bg='#3498db',
            fg='white',
            padx=20,
            pady=8,
            command=self.finish_editing
        ).pack(side='left', padx=5)
        
        # 메인 컨텐츠 (좌우 분할) - 비율 조정
        main_paned = tk.PanedWindow(self, orient='horizontal', bg='#bdc3c7', sashwidth=0)
        main_paned.pack(fill='both', expand=True)

        # 좌측: 리소스 관리 (고정)
        left_frame = tk.Frame(main_paned, width=175, bg='#ecf0f1')
        main_paned.add(left_frame, minsize=175, width=175, stretch='never')

        self.setup_resource_panel(left_frame)

        # 우측: 플로우 에디터 (넓게)
        right_frame = tk.Frame(main_paned, width=460, bg='white')
        main_paned.add(right_frame, minsize=460)

        self.setup_flow_panel(right_frame)

    
    def setup_resource_panel(self, parent):
        """리소스 패널 구성"""
        # 스크롤 가능한 캔버스
        canvas = tk.Canvas(parent, bg='#ecf0f1', highlightthickness=0)
        scrollbar = tk.Scrollbar(parent, orient='vertical', command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg='#ecf0f1')
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # 좌표 섹션
        self.setup_coordinate_section(scrollable_frame)
        
        # 엑셀 섹션
        self.setup_excel_section(scrollable_frame)
        
        # 이미지 섹션
        self.setup_image_section(scrollable_frame)
    
    def setup_coordinate_section(self, parent):
        """좌표 섹션"""
        section = tk.LabelFrame(
            parent,
            text="좌표 목록",
            font=("맑은 고딕", 11, "bold"),
            bg='#ecf0f1',
            padx=5,
            pady=10
        )
        section.pack(fill='x', padx=10, pady=10)
        
        # 좌표 리스트
        self.coord_list_frame = tk.Frame(section, bg='#ecf0f1')
        self.coord_list_frame.pack(fill='x')
        
        # 추가 버튼
        tk.Button(
            section,
            text="+ 새 좌표 추가",
            font=("맑은 고딕", 9),
            bg='#3498db',
            fg='white',
            command=self.add_coordinate_dialog
        ).pack(fill='x', pady=(10, 0))
        
        self.refresh_coordinate_list()
    
    def setup_excel_section(self, parent):
        """엑셀 섹션"""
        section = tk.LabelFrame(
            parent,
            text="엑셀 데이터",
            font=("맑은 고딕", 11, "bold"),
            bg='#ecf0f1',
            padx=5,
            pady=10
        )
        section.pack(fill='x', padx=10, pady=10)
        
        # 엑셀 리스트
        self.excel_list_frame = tk.Frame(section, bg='#ecf0f1')
        self.excel_list_frame.pack(fill='x')
        
        # 추가 버튼
        tk.Button(
            section,
            text="+ 새 엑셀 추가",
            font=("맑은 고딕", 9),
            bg='#2ecc71',
            fg='white',
            command=self.add_excel_dialog
        ).pack(fill='x', pady=(10, 0))
        
        self.refresh_excel_list()
    
    def setup_image_section(self, parent):
        """이미지 섹션"""
        section = tk.LabelFrame(
            parent,
            text="이미지 템플릿",
            font=("맑은 고딕", 11, "bold"),
            bg='#ecf0f1',
            padx=5,
            pady=10
        )
        section.pack(fill='x', padx=10, pady=10)
        
        # 이미지 리스트
        self.image_list_frame = tk.Frame(section, bg='#ecf0f1')
        self.image_list_frame.pack(fill='x')
        
        # 추가 버튼 - 직접 add_image_from_file 호출
        tk.Button(
            section,
            text="+ 새 이미지 추가",
            font=("맑은 고딕", 9),
            bg='#9b59b6',
            fg='white',
            command=self.add_image_from_file  # 수정
        ).pack(fill='x', pady=(10, 0))
        
        self.refresh_image_list()
    
    def setup_flow_panel(self, parent):
        """플로우 패널 구성"""
        # 제목
        title_frame = tk.Frame(parent, bg='white')
        title_frame.pack(fill='x', padx=0, pady=10)
        
        tk.Label(
            title_frame,
            text="⚙️ 플로우 시퀀스",
            font=("맑은 고딕", 13, "bold"),
            bg='white'
        ).pack(side='left')
        
        # 플로우 리스트 (스크롤 가능)
        list_frame = tk.Frame(parent, bg='white')
        list_frame.pack(fill='both', expand=True, padx=0, pady=10)

        canvas = tk.Canvas(list_frame, bg='white', highlightthickness=0)
        scrollbar = tk.Scrollbar(list_frame, orient='vertical', command=canvas.yview)
        self.flow_list_frame = tk.Frame(canvas, bg='white')

        canvas_window = canvas.create_window((0, 0), window=self.flow_list_frame, anchor='nw')

        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def on_canvas_configure(event):
            # 캔버스 너비에 맞춰 프레임 너비 설정
            canvas_width = event.width
            canvas.itemconfig(canvas_window, width=canvas_width)

        self.flow_list_frame.bind("<Configure>", on_frame_configure)
        canvas.bind("<Configure>", on_canvas_configure)
        canvas.configure(yscrollcommand=scrollbar.set)

        # 마우스 휠 스크롤 지원
        def _on_mousewheel(event):
            try:
                if canvas.winfo_exists():
                    canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except tk.TclError:
                pass

        canvas.bind("<MouseWheel>", _on_mousewheel)
        self.flow_list_frame.bind("<MouseWheel>", _on_mousewheel)

        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # 액션 추가 버튼
        btn_frame = tk.Frame(parent, bg='white')
        btn_frame.pack(fill='x', padx=20, pady=15)
        
        tk.Button(
            btn_frame,
            text="➕ 액션 추가",
            font=("맑은 고딕", 10, "bold"),
            bg='#3498db',
            fg='white',
            padx=20,
            pady=10,
            command=self.add_action_menu
        ).pack()
        
        self.refresh_flow_list()
    
    def refresh_coordinate_list(self):
        """좌표 목록 새로고침"""
        for widget in self.coord_list_frame.winfo_children():
            widget.destroy()
        
        if not self.coord_mgr.coordinates:
            tk.Label(
                self.coord_list_frame,
                text="좌표가 없습니다",
                font=("맑은 고딕", 9),
                fg='gray',
                bg='#ecf0f1'
            ).pack(fill='x', pady=5)
            return
        
        for coord in self.coord_mgr.coordinates:
            self.create_coordinate_item(coord)
    
    def create_coordinate_item(self, coord):
        """좌표 아이템 생성"""
        item = tk.Frame(self.coord_list_frame, bg='white', relief='ridge', borderwidth=1)
        item.pack(fill='x', pady=2)

        info_frame = tk.Frame(item, bg='white')
        info_frame.pack(side='left', fill='both', expand=True, padx=5, pady=5)

        # 제목 길이 제한 (최대 12자)
        MAX_TITLE_LENGTH = 10
        display_name = coord['name'] if len(coord['name']) <= MAX_TITLE_LENGTH else coord['name'][:MAX_TITLE_LENGTH] + '...'

        tk.Label(
            info_frame,
            text=f"{coord['id']}. {display_name}",
            font=("맑은 고딕", 9, "bold"),
            bg='white',
            anchor='w'
        ).pack(anchor='w', fill='x')
        
        tk.Label(
            info_frame,
            text=f"({coord['x']}, {coord['y']})",
            font=("맑은 고딕", 8),
            fg='gray',
            bg='white',
            anchor='w'
        ).pack(anchor='w')
        
        btn_frame = tk.Frame(item, bg='white')
        btn_frame.pack(side='right', padx=3)
        
        tk.Button(
            btn_frame,
            text="❌",
            font=("맑은 고딕", 7),
            width=3,
            command=lambda: self.delete_coordinate(coord['id'])
        ).pack(padx=(0,4))

    
    def refresh_excel_list(self):
        """엑셀 목록 새로고침"""
        for widget in self.excel_list_frame.winfo_children():
            widget.destroy()
        
        if not self.excel_mgr.excel_sources:
            tk.Label(
                self.excel_list_frame,
                text="엑셀 데이터가 없습니다",
                font=("맑은 고딕", 9),
                fg='gray',
                bg='#ecf0f1'
            ).pack(fill='x', pady=5)
            return
        
        for source in self.excel_mgr.excel_sources:
            self.create_excel_item(source)
    
    def create_excel_item(self, source):
        """엑셀 아이템 생성"""
        item = tk.Frame(self.excel_list_frame, bg='white', relief='ridge', borderwidth=1)
        item.pack(fill='x', pady=2)

        info_frame = tk.Frame(item, bg='white')
        info_frame.pack(side='left', fill='both', expand=True, padx=5, pady=5)

        # 이름만 표시 (제목 길이 제한)
        MAX_TITLE_LENGTH = 12
        display_name = source['name'] if len(source['name']) <= MAX_TITLE_LENGTH else source['name'][:MAX_TITLE_LENGTH] + '...'

        tk.Label(
            info_frame,
            text=display_name,
            font=("맑은 고딕", 9, "bold"),
            bg='white',
            anchor='w'
        ).pack(anchor='w', fill='x')

        tk.Label(
            info_frame,
            text=f"{source['row_count']} rows, {len(source['columns'])} cols",
            font=("맑은 고딕", 8),
            fg='gray',
            bg='white',
            anchor='w'
        ).pack(anchor='w')

        btn_frame = tk.Frame(item, bg='white')
        btn_frame.pack(side='right', padx=3)

        tk.Button(
            btn_frame,
            text="❌",
            font=("맑은 고딕", 7),
            width=3,
            command=lambda: self.delete_excel(source['id'])
        ).pack(padx=(0,4))
    
    def refresh_image_list(self):
        """이미지 목록 새로고침"""
        for widget in self.image_list_frame.winfo_children():
            widget.destroy()
        
        if not self.image_mgr.images:
            tk.Label(
                self.image_list_frame,
                text="이미지가 없습니다",
                font=("맑은 고딕", 9),
                fg='gray',
                bg='#ecf0f1'
            ).pack(fill='x', pady=5)
            return
        
        for image in self.image_mgr.images:
            self.create_image_item(image)
    
    def create_image_item(self, image):
        """이미지 아이템 생성"""
        item = tk.Frame(self.image_list_frame, bg='white', relief='ridge', borderwidth=1)
        item.pack(fill='x', pady=2)

        info_frame = tk.Frame(item, bg='white')
        info_frame.pack(side='left', fill='both', expand=True, padx=5, pady=5)

        # 제목 길이 제한 (최대 12자)
        MAX_TITLE_LENGTH = 10
        display_name = image['name'] if len(image['name']) <= MAX_TITLE_LENGTH else image['name'][:MAX_TITLE_LENGTH] + '...'

        tk.Label(
            info_frame,
            text=f"{image['id']}. {display_name}",
            font=("맑은 고딕", 9, "bold"),
            bg='white',
            anchor='w'
        ).pack(anchor='w', fill='x')
        
        tk.Label(
            info_frame,
            text=f"정확도: {int(image['confidence']*100)}%",
            font=("맑은 고딕", 8),
            fg='gray',
            bg='white',
            anchor='w'
        ).pack(anchor='w')
        
        btn_frame = tk.Frame(item, bg='white')
        btn_frame.pack(side='right', padx=3)
        
        tk.Button(
            btn_frame,
            text="❌",
            font=("맑은 고딕", 7),
            width=3,
            command=lambda: self.delete_image(image['id'])
        ).pack(padx=(0,4))
    
    def refresh_flow_list(self):
        """플로우 목록 새로고침"""
        for widget in self.flow_list_frame.winfo_children():
            widget.destroy()
        
        if not self.flow_mgr.flow_sequence:
            tk.Label(
                self.flow_list_frame,
                text="액션을 추가하세요",
                font=("맑은 고딕", 10),
                fg='gray',
                bg='white'
            ).pack(fill='x', padx=5 ,pady=20)
            return
        
        for idx, action in enumerate(self.flow_mgr.flow_sequence):
            self.create_flow_item(idx, action)
    
    def create_flow_item(self, idx, action):
        """플로우 아이템 생성 (넓게)"""
        item = tk.Frame(self.flow_list_frame, bg='#ecf0f1', relief='raised', borderwidth=1)
        item.pack(fill='x', pady=3, expand=True, padx=10)

        # 드래그 앤 드롭 이벤트 바인딩
        item.bind("<Button-1>", lambda e: self.on_drag_start(e, item, idx))
        item.bind("<B1-Motion>", lambda e: self.on_drag_motion(e, item))
        item.bind("<ButtonRelease-1>", lambda e: self.on_drag_release(e, item, idx))

        # 액션 타입별 색상 가져오기
        action_color = self.get_action_color(action.get('type', ''))

        # 번호
        num_label = tk.Label(
            item,
            text=f"{idx+1}",
            font=("맑은 고딕", 11, "bold"),
            bg=action_color,
            fg='white',
            width=3,
            height=1
        )
        num_label.pack(side='left', padx=(8,5), pady=8)
        num_label.bind("<Button-1>", lambda e: self.on_drag_start(e, item, idx))
        num_label.bind("<B1-Motion>", lambda e: self.on_drag_motion(e, item))
        num_label.bind("<ButtonRelease-1>", lambda e: self.on_drag_release(e, item, idx))

        # 액션 설명
        display_text = self.flow_mgr.get_action_display_text(
            action, self.coord_mgr, self.excel_mgr, self.image_mgr
        )

        text_label = tk.Label(
            item,
            text=display_text,
            font=("맑은 고딕", 11),
            bg='#ecf0f1',
            anchor='w',
            justify='left'
        )
        text_label.pack(side='left', fill='both', expand=True, padx=5, pady=8)
        text_label.bind("<Button-1>", lambda e: self.on_drag_start(e, item, idx))
        text_label.bind("<B1-Motion>", lambda e: self.on_drag_motion(e, item))
        text_label.bind("<ButtonRelease-1>", lambda e: self.on_drag_release(e, item, idx))
        
        # 삭제 버튼만 유지 (드래그 앤 드롭으로 순서 변경 가능하므로 ▲▼ 버튼 제거)
        tk.Button(
            item,
            text="❌",
            font=("맑은 고딕", 9),
            width=3,
            height=1,
            bg='#e74c3c',
            fg='white',
            command=lambda: self.delete_action(action['id'])
        ).pack(side='right', padx=8, pady=5)
    
    # 좌표 추가
    def add_coordinate_dialog(self):
        """좌표 추가 다이얼로그"""
        dialog = tk.Toplevel(self.parent)
        dialog.title("좌표 추가")
        dialog.geometry("300x200")
        dialog.transient(self.parent)
        dialog.grab_set()
        dialog.attributes('-topmost', True)  # 추가
        tk.Label(
            dialog,
            text="좌표 추가",
            font=("맑은 고딕", 12, "bold")
        ).pack(pady=15)
        
        tk.Label(
            dialog,
            text="3초 후 마우스가 있는 위치의\n좌표가 저장됩니다.",
            font=("맑은 고딕", 9),
            fg='gray'
        ).pack(pady=5)
        
        tk.Label(
            dialog,
            text="원하는 위치로 마우스를 이동하세요!",
            font=("맑은 고딕", 9, "bold"),
            fg='#e74c3c'
        ).pack(pady=5)
        
        def start_capture():
            dialog.destroy()
            self.capture_coordinate()

        start_btn = tk.Button(
            dialog,
            text="시작",
            font=("맑은 고딕", 10),
            bg='#3498db',
            fg='white',
            padx=30,
            pady=8,
            command=start_capture
        )
        start_btn.pack(pady=15)

        # 스페이스바와 엔터키로 시작 가능
        dialog.bind('<space>', lambda e: start_capture())
        dialog.bind('<Return>', lambda e: start_capture())

        center_window_on_parent(dialog, self.parent)  # 이미 아이콘 설정 포함
        dialog.lift()
        dialog.focus_force()
        
    def capture_coordinate(self):
        """좌표 캡처"""
        # 카운트다운 창
        countdown_window = tk.Toplevel(self.parent)
        countdown_window.title("좌표 캡처")
        countdown_window.attributes('-topmost', True)
        countdown_window.attributes('-alpha', 0.9)
        
        # 화면 중앙에 배치
        window_width = 300
        window_height = 150
        screen_width = countdown_window.winfo_screenwidth()
        screen_height = countdown_window.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        countdown_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        countdown_window.overrideredirect(True)  # 테두리 제거
        countdown_window.configure(bg='#2c3e50')
        
        label = tk.Label(
            countdown_window,
            text="3",
            font=("맑은 고딕", 72, "bold"),
            fg='white',
            bg='#2c3e50'
        )
        label.pack(expand=True)
        
        info_label = tk.Label(
            countdown_window,
            text="마우스를 원하는 위치로 이동하세요",
            font=("맑은 고딕", 10),
            fg='#ecf0f1',
            bg='#2c3e50'
        )
        info_label.pack(pady=(0, 20))
        
        captured_data = {'x': None, 'y': None, 'thumbnail': None}
        
        def countdown(count):
            if count > 0:
                label.config(text=str(count))
                if count == 1:
                    label.config(fg='#e74c3c')
                countdown_window.after(1000, countdown, count-1)
            else:
                # 좌표 캡처
                x, y, thumbnail = self.coord_mgr.capture_current_position()
                captured_data['x'] = x
                captured_data['y'] = y
                captured_data['thumbnail'] = thumbnail
                
                countdown_window.destroy()
                
                # 이름 입력 (메인 스레드에서 안전하게)
                self.after(100, lambda: self.show_coordinate_name_dialog(
                    captured_data['x'],
                    captured_data['y'],
                    captured_data['thumbnail']
                ))
        
        countdown(3)
    
    def show_coordinate_name_dialog(self, x, y, thumbnail):
        """좌표 이름 입력 다이얼로그"""
        dialog = tk.Toplevel(self.parent)
        dialog.title("좌표 이름 입력")
        dialog.geometry("300x180")
        dialog.transient(self.parent)
        dialog.grab_set()
        dialog.attributes('-topmost', True)  # 추가
        tk.Label(
            dialog,
            text=f"좌표: ({x}, {y})",
            font=("맑은 고딕", 10, "bold"),
            fg='#27ae60'
        ).pack(pady=10)
        
        tk.Label(
            dialog,
            text="이름:",
            font=("맑은 고딕", 10)
        ).pack(anchor='w', padx=30, pady=(0,5))
        
        name_entry = tk.Entry(
            dialog,
            font=("맑은 고딕", 11),
            width=30
        )
        name_entry.pack(padx=30, pady=(0, 10))
        name_entry.focus()
        
        def save_coord():
            name = name_entry.get().strip()
            if not name:
                messagebox.showwarning("경고", "이름을 입력하세요.", parent=dialog)
                return
            
            self.coord_mgr.add_coordinate(name, x, y, thumbnail=thumbnail)
            self.refresh_coordinate_list()
            dialog.destroy()
            messagebox.showinfo("완료", f"좌표 '{name}'이(가) 추가되었습니다!")
        
        def on_cancel():
            dialog.destroy()
        
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=10)
        
        tk.Button(
            btn_frame,
            text="저장",
            font=("맑은 고딕", 10),
            bg='#27ae60',
            fg='white',
            padx=20,
            pady=5,
            command=save_coord
        ).pack(side='left', padx=5)
        
        tk.Button(
            btn_frame,
            text="취소",
            font=("맑은 고딕", 10),
            bg='#95a5a6',
            fg='white',
            padx=20,
            pady=5,
            command=on_cancel
        ).pack(side='left', padx=5)
        center_window_on_parent(dialog, self.parent)
        dialog.lift()
        dialog.focus_force()
        # Enter 키 바인딩
        name_entry.bind('<Return>', lambda e: save_coord())
    
    def delete_coordinate(self, coord_id):
        """좌표 삭제"""
        if messagebox.askyesno("확인", "이 좌표를 삭제하시겠습니까?"):
            self.coord_mgr.remove_coordinate(coord_id)
            self.refresh_coordinate_list()
 
        # 엑셀 추가
    def add_excel_dialog(self):
        """엑셀 추가 다이얼로그"""
        filepath = filedialog.askopenfilename(
            parent=self.parent,
            title="엑셀 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
            initialdir=os.path.expanduser("~")
        )
        
        if not filepath:
            return
        
        # 시트 선택
        sheets = self.excel_mgr.get_sheet_names(filepath)
        if not sheets:
            messagebox.showerror("오류", "엑셀 파일을 읽을 수 없습니다.")
            return
        
        sheet_name = sheets[0] if len(sheets) == 1 else self.select_sheet_dialog(sheets)
        if not sheet_name:
            return
        
        # 칼럼 선택
        columns = self.excel_mgr.get_columns(filepath, sheet_name)
        if not columns:
            messagebox.showerror("오류", "칼럼을 읽을 수 없습니다.")
            return
        
        selected_columns = self.select_columns_dialog(columns)
        if not selected_columns:
            return
        
        # 이름 입력 (커스텀 다이얼로그 사용)
        dialog = NameInputDialog(
            self.parent,
            title="데이터 소스 이름",
            message="이 엑셀 데이터의 이름을 입력하세요:",
            initial_value=""
        )
        self.parent.wait_window(dialog)

        name = dialog.result
        if not name:
            return

        # 추가 (기본값으로 자동 최신 파일 선택 비활성화, 중복 제거 활성화)
        source = self.excel_mgr.add_excel_source(
            name, filepath, sheet_name, selected_columns,
            auto_latest=False,
            auto_directory=None,
            auto_prefix='list',
            remove_empty_rows=True,
            remove_duplicates=True
        )
        if source:
            self.refresh_excel_list()
            messagebox.showinfo("완료", f"엑셀 데이터 '{name}'이(가) 추가되었습니다.\n\n행 수: {source['row_count']}\n칼럼 수: {len(selected_columns)}")
        else:
            messagebox.showerror("오류", "엑셀 데이터를 추가할 수 없습니다.")
    
    def select_sheet_dialog(self, sheets):
        """시트 선택 다이얼로그"""
        dialog = tk.Toplevel(self.parent)
        dialog.title("시트 선택")
        dialog.geometry("300x400")
        dialog.transient(self.parent)
        dialog.grab_set()
        dialog.attributes('-topmost', True)  # 추가
        
        result = [None]
        
        tk.Label(
            dialog,
            text="시트를 선택하세요",
            font=("맑은 고딕", 11, "bold")
        ).pack(pady=10)
        
        listbox = tk.Listbox(dialog, font=("맑은 고딕", 10))
        listbox.pack(fill='both', expand=True, padx=20, pady=10)
        
        for sheet in sheets:
            listbox.insert(tk.END, sheet)
        
        listbox.selection_set(0)
        
        def on_select():
            selection = listbox.curselection()
            if selection:
                result[0] = sheets[selection[0]]
                dialog.destroy()
        
        tk.Button(
            dialog,
            text="선택",
            command=on_select,
            bg='#3498db',
            fg='white',
            padx=20,
            pady=5
        ).pack(pady=10)
        
        # 더블클릭, Enter 키 바인딩 추가
        listbox.bind('<Double-Button-1>', lambda e: on_select())
        listbox.bind('<Return>', lambda e: on_select())
        
        # 중앙 배치 추가
        center_window_on_parent(dialog, self.parent)
        dialog.lift()
        dialog.focus_force()
        listbox.focus_set()
        
        dialog.wait_window()
        return result[0]

    
    def select_columns_dialog(self, columns):
        """칼럼 선택 다이얼로그 (개선됨)"""
        dialog = tk.Toplevel(self.parent)
        dialog.title("칼럼 선택")
        dialog.geometry("300x500")
        dialog.transient(self.parent)
        dialog.grab_set()
        
        result = [None]
        
        # 제목
        tk.Label(
            dialog,
            text="사용할 칼럼을 선택하세요",
            font=("맑은 고딕", 12, "bold")
        ).pack(pady=15)
        
        # 전체 선택/해제 버튼
        button_frame = tk.Frame(dialog)
        button_frame.pack(fill='x', padx=20, pady=(0, 10))
        
        vars_dict = {}  # 미리 선언 (함수에서 사용)
        
        def select_all():
            for var in vars_dict.values():
                var.set(True)
        
        def deselect_all():
            for var in vars_dict.values():
                var.set(False)
        
        tk.Button(
            button_frame,
            text="✅ 전체 선택",
            font=("맑은 고딕", 9),
            bg='#27ae60',
            fg='white',
            padx=15,
            pady=5,
            command=select_all
        ).pack(side='left', padx=5)
        
        tk.Button(
            button_frame,
            text="❌ 전체 해제",
            font=("맑은 고딕", 9),
            bg='#e74c3c',
            fg='white',
            padx=15,
            pady=5,
            command=deselect_all
        ).pack(side='left', padx=5)
        
        # 체크박스 리스트
        list_frame = tk.Frame(dialog)
        list_frame.pack(fill='both', expand=True, padx=0, pady=10)
        
        canvas = tk.Canvas(list_frame, bg='white', highlightthickness=1, highlightbackground='#bdc3c7')
        scrollbar = tk.Scrollbar(list_frame, orient='vertical', command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg='white')
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 체크박스 생성
        for col in columns:
            var = tk.BooleanVar(value=True)
            vars_dict[col] = var
            
            cb = tk.Checkbutton(
                scrollable_frame,
                text=col,
                variable=var,
                font=("맑은 고딕", 10),
                bg='white',
                anchor='w'
            )
            cb.pack(anchor='w', padx=10, pady=3, fill='x')
        
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # 선택 개수 표시
        count_label = tk.Label(
            dialog,
            text=f"선택된 칼럼: {len(columns)}개 / 전체: {len(columns)}개",
            font=("맑은 고딕", 9),
            fg='#7f8c8d'
        )
        count_label.pack(pady=5)
        
        def update_count(*args):
            selected_count = sum(1 for var in vars_dict.values() if var.get())
            count_label.config(text=f"선택된 칼럼: {selected_count}개 / 전체: {len(columns)}개")
        
        # 체크박스 변경 시 카운트 업데이트
        for var in vars_dict.values():
            var.trace_add('write', update_count)
        
        # 확인/취소 버튼
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=15)
        
        def on_confirm():
            selected = [col for col, var in vars_dict.items() if var.get()]
            if not selected:
                messagebox.showwarning("경고", "최소 하나의 칼럼을 선택하세요.", parent=dialog)
                return
            result[0] = selected
            dialog.destroy()
        
        def on_cancel():
            result[0] = None
            dialog.destroy()
        
        tk.Button(
            btn_frame,
            text="확인",
            font=("맑은 고딕", 10, "bold"),
            bg='#3498db',
            fg='white',
            padx=30,
            pady=8,
            command=on_confirm
        ).pack(side='left', padx=5)
        
        tk.Button(
            btn_frame,
            text="취소",
            font=("맑은 고딕", 10),
            bg='#95a5a6',
            fg='white',
            padx=30,
            pady=8,
            command=on_cancel
        ).pack(side='left', padx=5)
        
        # Enter 키로 확인, ESC 키로 취소
        dialog.bind('<Return>', lambda e: on_confirm())
        dialog.bind('<Escape>', lambda e: on_cancel())

        # 중앙 배치 및 포커스 (추가)
        dialog.attributes('-topmost', True)
        center_window_on_parent(dialog, self.parent)
        dialog.lift()
        dialog.focus_force()

        dialog.wait_window()
        return result[0]

    def select_data_options_dialog(self):
        """데이터 처리 옵션 선택 다이얼로그"""
        dialog = tk.Toplevel(self.parent)
        dialog.title("데이터 처리 옵션")
        dialog.geometry("400x300")
        dialog.transient(self.parent)
        dialog.grab_set()
        dialog.attributes('-topmost', True)

        result = [None]

        # 제목
        tk.Label(
            dialog,
            text="📊 데이터 처리 옵션",
            font=("맑은 고딕", 14, "bold")
        ).pack(pady=20)

        # 설명
        tk.Label(
            dialog,
            text="엑셀 데이터 로드 시 적용할 옵션을 선택하세요",
            font=("맑은 고딕", 9),
            fg='#7f8c8d'
        ).pack(pady=(0, 20))

        # 옵션 프레임
        options_frame = tk.Frame(dialog, bg='white', padx=20, pady=20)
        options_frame.pack(fill='both', expand=True, padx=20)

        # 공백 행 제거 옵션
        remove_empty_var = tk.BooleanVar(value=True)

        empty_frame = tk.Frame(options_frame, bg='white')
        empty_frame.pack(fill='x', pady=10)

        tk.Checkbutton(
            empty_frame,
            text="🧹 상위 공백 행 제거",
            variable=remove_empty_var,
            font=("맑은 고딕", 11, "bold"),
            bg='white'
        ).pack(anchor='w')

        tk.Label(
            empty_frame,
            text="   데이터 시작 전의 빈 행을 자동으로 제거합니다",
            font=("맑은 고딕", 9),
            fg='#7f8c8d',
            bg='white'
        ).pack(anchor='w')

        # 중복 제거 옵션
        remove_dup_var = tk.BooleanVar(value=False)

        dup_frame = tk.Frame(options_frame, bg='white')
        dup_frame.pack(fill='x', pady=10)

        tk.Checkbutton(
            dup_frame,
            text="🔄 중복 행 제거",
            variable=remove_dup_var,
            font=("맑은 고딕", 11, "bold"),
            bg='white'
        ).pack(anchor='w')

        tk.Label(
            dup_frame,
            text="   선택한 컬럼 기준으로 중복된 행을 제거합니다",
            font=("맑은 고딕", 9),
            fg='#7f8c8d',
            bg='white'
        ).pack(anchor='w')

        # 버튼
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=20)

        def on_confirm():
            result[0] = {
                'remove_empty': remove_empty_var.get(),
                'remove_duplicates': remove_dup_var.get()
            }
            dialog.destroy()

        def on_cancel():
            result[0] = None
            dialog.destroy()

        tk.Button(
            btn_frame,
            text="확인",
            font=("맑은 고딕", 10, "bold"),
            bg='#3498db',
            fg='white',
            padx=30,
            pady=8,
            command=on_confirm
        ).pack(side='left', padx=5)

        tk.Button(
            btn_frame,
            text="취소",
            font=("맑은 고딕", 10),
            bg='#95a5a6',
            fg='white',
            padx=30,
            pady=8,
            command=on_cancel
        ).pack(side='left', padx=5)

        # 중앙 배치
        center_window_on_parent(dialog, self.parent)
        dialog.lift()
        dialog.focus_force()

        dialog.wait_window()
        return result[0]


    def delete_excel(self, source_id):
        """엑셀 삭제"""
        if messagebox.askyesno("확인", "이 엑셀 데이터를 삭제하시겠습니까?"):
            self.excel_mgr.remove_excel_source(source_id)
            self.refresh_excel_list()
    
    
    def add_image_from_file(self):
        """파일에서 이미지 추가"""
        filepath = filedialog.askopenfilename(
            parent=self.parent,
            title="이미지 파일 선택",
            filetypes=[("Image files", "*.png *.jpg *.jpeg"), ("All files", "*.*")],
            initialdir=os.path.expanduser("~")
        )
        
        if not filepath:
            return
        
        # 이미지를 base64로 변환
        import base64
        with open(filepath, 'rb') as f:
            img_data = base64.b64encode(f.read()).decode()
        
        # 이름 및 정확도 입력
        from ui.dialogs import ImageNameDialog
        dialog = ImageNameDialog(
            self.parent,
            title="이미지 등록",
            initial_name="",
            initial_confidence=80
        )
        self.parent.wait_window(dialog)

        if not dialog.result:
            return

        name = dialog.result['name']
        confidence = dialog.result['confidence']

        # 추가
        image = self.image_mgr.add_image(name, img_data, confidence=confidence)
        if image:
            self.refresh_image_list()
            messagebox.showinfo("완료", f"이미지 '{name}'이(가) 추가되었습니다.\n정확도: {int(confidence*100)}%")
        else:
            messagebox.showerror("오류", "이미지를 추가할 수 없습니다.")

    def delete_image(self, image_id):
        """이미지 삭제"""
        if messagebox.askyesno("확인", "이 이미지를 삭제하시겠습니까?"):
            self.image_mgr.remove_image(image_id)
            self.refresh_image_list()
    
    # 플로우 관리
    def add_action_menu(self):
        """액션 추가 메뉴"""
        from ui.dialogs import ActionSelectDialog
        
        dialog = ActionSelectDialog(
            self.parent,
            self.coord_mgr,
            self.excel_mgr,
            self.image_mgr
        )
        self.parent.wait_window(dialog)
        
        if dialog.result:
            action = self.flow_mgr.add_action(
                dialog.result['type'],
                dialog.result['params']
            )
            self.refresh_flow_list()
    
    def move_action_up(self, action_id):
        """액션 위로 이동"""
        if self.flow_mgr.move_action_up(action_id):
            self.refresh_flow_list()
    
    def move_action_down(self, action_id):
        """액션 아래로 이동"""
        if self.flow_mgr.move_action_down(action_id):
            self.refresh_flow_list()
    
    def delete_action(self, action_id):
        """액션 삭제"""
        if messagebox.askyesno("확인", "이 액션을 삭제하시겠습니까?"):
            self.flow_mgr.remove_action(action_id)
            self.refresh_flow_list()
    
    # 프로젝트 관리
    def save_project(self):
        """프로젝트 저장"""
        self.project_data['coordinates'] = self.coord_mgr.to_list()
        self.project_data['excel_sources'] = self.excel_mgr.to_list()
        self.project_data['images'] = self.image_mgr.to_list()
        self.project_data['flow_sequence'] = self.flow_mgr.to_list()
        
        if ProjectManager.save_project(self.filepath, self.project_data):
            messagebox.showinfo("완료", "프로젝트가 저장되었습니다.")
        else:
            messagebox.showerror("오류", "프로젝트를 저장할 수 없습니다.")
    
    
    def finish_editing(self):
        """편집 완료"""
        self.save_project()
        self.app.show_start_screen()


    def get_action_color(self, action_type):
        """액션 타입별 색상 반환"""
        color_map = {
            # 마우스 동작 - 파란색
            'click_coord': '#3498db',
            'click_image': '#3498db',
            'mouse_scroll': '#3498db',

            # 키보드 동작 - 녹색
            'type_text': '#27ae60',
            'type_variable': '#27ae60',
            'key_press': '#27ae60',
            'paste': '#27ae60',

            # 제어 동작 - 빨간색
            'delay': '#e74c3c',
            'wait_image': '#e74c3c',

            # 기타 - 노란색
            'screenshot': '#f39c12',
        }

        return color_map.get(action_type, '#95a5a6')

    def get_action_category(self, action_type):
        """액션 카테고리 반환"""
        categories = {
            # 마우스 동작
            'click_coord': '🖱️ 마우스',
            'click_image': '🖱️ 마우스',
            'mouse_scroll': '🖱️ 마우스',

            # 키보드 동작
            'type_text': '⌨️ 키보드',
            'type_variable': '⌨️ 키보드',
            'key_press': '⌨️ 키보드',
            'paste': '⌨️ 키보드',

            # 제어 동작
            'delay': '⏱️ 제어',
            'wait_image': '⏱️ 제어',

            # 기타
            'screenshot': '💾 기타',
        }

        return categories.get(action_type, '❓ 기타')

    # 드래그 앤 드롭 이벤트 핸들러
    def on_drag_start(self, event, item, idx):
        """드래그 시작"""
        # 버튼 클릭은 드래그 무시
        if event.widget.winfo_class() == 'Button':
            return

        self.drag_data["item"] = item
        self.drag_data["index"] = idx
        self.drag_data["y"] = event.y_root

        # 드래그 중 시각적 피드백
        item.config(bg='#3498db', relief='sunken')

    def on_drag_motion(self, event, item):
        """드래그 중"""
        if self.drag_data["item"] is None:
            return

        # 드래그 중인 아이템만 색상 유지
        if item == self.drag_data["item"]:
            item.config(bg='#3498db')

    def on_drag_release(self, event, item, idx):
        """드롭"""
        if self.drag_data["item"] is None:
            return

        # 버튼 클릭은 드래그 무시
        if event.widget.winfo_class() == 'Button':
            self.drag_data["item"].config(bg='#ecf0f1', relief='raised')
            self.drag_data["item"] = None
            return

        drag_idx = self.drag_data["index"]

        # 드롭 위치 계산
        children = list(self.flow_list_frame.winfo_children())
        drop_idx = None

        for i, child in enumerate(children):
            if child.winfo_y() <= event.y_root - self.flow_list_frame.winfo_rooty() <= child.winfo_y() + child.winfo_height():
                drop_idx = i
                break

        if drop_idx is None:
            # 마지막 위치로 드롭
            drop_idx = len(children) - 1

        # 드래그 상태 초기화 (UI 새로고침 전에 먼저 해야 함)
        self.drag_data["item"].config(bg='#ecf0f1', relief='raised')
        self.drag_data["item"] = None

        # 순서 변경
        if drag_idx != drop_idx:
            # flow_manager에서 순서 변경
            action = self.flow_mgr.flow_sequence.pop(drag_idx)
            self.flow_mgr.flow_sequence.insert(drop_idx, action)

            # UI 새로고침
            self.refresh_flow_list()

