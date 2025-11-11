"""
프로젝트 편집 화면 - 좌표/엑셀/이미지/플로우 관리
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from datetime import datetime
import main

from core.project_manager import ProjectManager
from core.coordinate_manager import CoordinateManager
from core.excel_manager import ExcelManager
from core.image_manager import ImageManager
from core.flow_manager import FlowManager
from ui.dialogs import ActionSelectDialog, NameInputDialog

def set_dialog_icon(dialog):
    """다이얼로그에 아이콘 설정"""
    try:
        if hasattr(main, 'ICON_PATH') and main.ICON_PATH:
            dialog.iconbitmap(main.ICON_PATH)
    except Exception as e:
        print(f"⚠️ 다이얼로그 아이콘 설정 실패: {e}")

def center_dialog(dialog, parent):
    """다이얼로그를 부모 창 중앙에 배치"""
    dialog.update_idletasks()
    
    # 부모 창의 위치와 크기
    parent_x = parent.winfo_x()
    parent_y = parent.winfo_y()
    parent_width = parent.winfo_width()
    parent_height = parent.winfo_height()
    
    # 다이얼로그 크기
    dialog_width = dialog.winfo_width()
    dialog_height = dialog.winfo_height()
    
    # 중앙 위치 계산
    x = parent_x + (parent_width - dialog_width) // 2
    y = parent_y + (parent_height - dialog_height) // 2
    
    dialog.geometry(f'+{x}+{y}')

    set_dialog_icon(dialog)


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
        ).pack(side='left', padx=20, pady=15)
        
        # 헤더 버튼들
        btn_frame = tk.Frame(header, bg='#34495e')
        btn_frame.pack(side='right', padx=20)
        
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
        main_paned = tk.PanedWindow(self, orient='horizontal', sashwidth=3, bg='#bdc3c7')
        main_paned.pack(fill='both', expand=True)
        
        # 좌측: 리소스 관리 (컴팩트하게)
        left_frame = tk.Frame(main_paned, width=280, bg='#ecf0f1')
        main_paned.add(left_frame, minsize=250)
        
        self.setup_resource_panel(left_frame)
        
        # 우측: 플로우 에디터 (넓게)
        right_frame = tk.Frame(main_paned, bg='white')
        main_paned.add(right_frame, minsize=600)
        
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
            text="📍 좌표 목록",
            font=("맑은 고딕", 11, "bold"),
            bg='#ecf0f1',
            padx=10,
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
            text="📊 엑셀 데이터",
            font=("맑은 고딕", 11, "bold"),
            bg='#ecf0f1',
            padx=10,
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
            text="🖼️ 이미지 템플릿",
            font=("맑은 고딕", 11, "bold"),
            bg='#ecf0f1',
            padx=10,
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
        title_frame.pack(fill='x', padx=20, pady=10)
        
        tk.Label(
            title_frame,
            text="⚙️ 플로우 시퀀스",
            font=("맑은 고딕", 13, "bold"),
            bg='white'
        ).pack(side='left')
        
        # 플로우 리스트 (스크롤 가능)
        list_frame = tk.Frame(parent, bg='white')
        list_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        canvas = tk.Canvas(list_frame, bg='white', highlightthickness=0)
        scrollbar = tk.Scrollbar(list_frame, orient='vertical', command=canvas.yview)
        self.flow_list_frame = tk.Frame(canvas, bg='white')
        
        self.flow_list_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.flow_list_frame, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # 액션 추가 버튼
        btn_frame = tk.Frame(parent, bg='white')
        btn_frame.pack(fill='x', padx=20, pady=10)
        
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
            ).pack(pady=5)
            return
        
        for coord in self.coord_mgr.coordinates:
            self.create_coordinate_item(coord)
    
    def create_coordinate_item(self, coord):
        """좌표 아이템 생성"""
        item = tk.Frame(self.coord_list_frame, bg='white', relief='ridge', borderwidth=1)
        item.pack(fill='x', pady=2)
        
        info_frame = tk.Frame(item, bg='white')
        info_frame.pack(side='left', fill='both', expand=True, padx=5, pady=5)
        
        tk.Label(
            info_frame,
            text=f"{coord['id']}. {coord['name']}",
            font=("맑은 고딕", 9, "bold"),
            bg='white',
            anchor='w',
            wraplength=200  # 좁게 조정
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
        ).pack()

    
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
            ).pack(pady=5)
            return
        
        for source in self.excel_mgr.excel_sources:
            self.create_excel_item(source)
    
    def create_excel_item(self, source):
        """엑셀 아이템 생성"""
        item = tk.Frame(self.excel_list_frame, bg='white', relief='ridge', borderwidth=1)
        item.pack(fill='x', pady=2)
        
        info_frame = tk.Frame(item, bg='white')
        info_frame.pack(side='left', fill='both', expand=True, padx=5, pady=5)
        
        tk.Label(
            info_frame,
            text=f"{source['id']}. {source['name']}",
            font=("맑은 고딕", 9, "bold"),
            bg='white',
            anchor='w',
            wraplength=200  # 좁게 조정
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
        ).pack()
    
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
            ).pack(pady=5)
            return
        
        for image in self.image_mgr.images:
            self.create_image_item(image)
    
    def create_image_item(self, image):
        """이미지 아이템 생성"""
        item = tk.Frame(self.image_list_frame, bg='white', relief='ridge', borderwidth=1)
        item.pack(fill='x', pady=2)
        
        info_frame = tk.Frame(item, bg='white')
        info_frame.pack(side='left', fill='both', expand=True, padx=5, pady=5)
        
        tk.Label(
            info_frame,
            text=f"{image['id']}. {image['name']}",
            font=("맑은 고딕", 9, "bold"),
            bg='white',
            anchor='w',
            wraplength=200  # 좁게 조정
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
        ).pack()
    
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
            ).pack(pady=20)
            return
        
        for idx, action in enumerate(self.flow_mgr.flow_sequence):
            self.create_flow_item(idx, action)
    
    def create_flow_item(self, idx, action):
        """플로우 아이템 생성 (넓게)"""
        item = tk.Frame(self.flow_list_frame, bg='#ecf0f1', relief='raised', borderwidth=1)
        item.pack(fill='x', pady=3, padx=10)
        
        # 번호
        tk.Label(
            item,
            text=f"{idx+1}",
            font=("맑은 고딕", 11, "bold"),
            bg='#3498db',
            fg='white',
            width=4,
            height=2
        ).pack(side='left', padx=(8, 15), pady=8)
        
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
        text_label.pack(side='left', fill='both', expand=True, padx=10, pady=8)
        
        # 버튼
        btn_frame = tk.Frame(item, bg='#ecf0f1')
        btn_frame.pack(side='right', padx=8, pady=5)
        
        tk.Button(
            btn_frame,
            text="▲",
            font=("맑은 고딕", 9),
            width=3,
            height=1,
            command=lambda: self.move_action_up(action['id'])
        ).pack(side='left', padx=3)
        
        tk.Button(
            btn_frame,
            text="▼",
            font=("맑은 고딕", 9),
            width=3,
            height=1,
            command=lambda: self.move_action_down(action['id'])
        ).pack(side='left', padx=3)
        
        tk.Button(
            btn_frame,
            text="❌",
            font=("맑은 고딕", 9),
            width=3,
            height=1,
            bg='#e74c3c',
            fg='white',
            command=lambda: self.delete_action(action['id'])
        ).pack(side='left', padx=3)
    
    # 좌표 추가
    def add_coordinate_dialog(self):
        """좌표 추가 다이얼로그"""
        dialog = tk.Toplevel(self.parent)
        dialog.title("좌표 추가")
        dialog.geometry("350x180")
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
        
        tk.Button(
            dialog,
            text="시작",
            font=("맑은 고딕", 10),
            bg='#3498db',
            fg='white',
            padx=30,
            pady=8,
            command=start_capture
        ).pack(pady=15)

        center_dialog(dialog, self.parent)  # 이미 아이콘 설정 포함
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
        dialog.geometry("350x200")
        dialog.transient(self.parent)
        dialog.grab_set()
        dialog.attributes('-topmost', True)  # 추가
        tk.Label(
            dialog,
            text=f"좌표: ({x}, {y})",
            font=("맑은 고딕", 10, "bold"),
            fg='#27ae60'
        ).pack(pady=15)
        
        tk.Label(
            dialog,
            text="이름:",
            font=("맑은 고딕", 10)
        ).pack(anchor='w', padx=30, pady=(10, 5))
        
        name_entry = tk.Entry(
            dialog,
            font=("맑은 고딕", 11),
            width=30
        )
        name_entry.pack(padx=30, pady=(0, 15))
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
        btn_frame.pack(pady=15)
        
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
        center_dialog(dialog, self.parent)
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
            title="엑셀 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
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
        
        # 추가
        source = self.excel_mgr.add_excel_source(name, filepath, sheet_name, selected_columns)
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
        center_dialog(dialog, self.parent)
        dialog.lift()
        dialog.focus_force()
        listbox.focus_set()
        
        dialog.wait_window()
        return result[0]

    
    def select_columns_dialog(self, columns):
        """칼럼 선택 다이얼로그 (개선됨)"""
        dialog = tk.Toplevel(self.parent)
        dialog.title("칼럼 선택")
        dialog.geometry("400x550")
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
        list_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
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
        center_dialog(dialog, self.parent)
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
            title="이미지 파일 선택",
            filetypes=[("Image files", "*.png *.jpg *.jpeg"), ("All files", "*.*")]
        )
        
        if not filepath:
            return
        
        # 이미지를 base64로 변환
        import base64
        with open(filepath, 'rb') as f:
            img_data = base64.b64encode(f.read()).decode()
        
        # 이름 입력 (커스텀 다이얼로그 사용)
        dialog = NameInputDialog(
            self.parent,
            title="이미지 이름",
            message="이미지 이름을 입력하세요:",
            initial_value=""
        )
        self.parent.wait_window(dialog)
        
        name = dialog.result
        if not name:
            return
        
        # 추가
        image = self.image_mgr.add_image(name, img_data)
        if image:
            self.refresh_image_list()
            messagebox.showinfo("완료", f"이미지 '{name}'이(가) 추가되었습니다.")
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
    
    def test_flow(self):
        """플로우 테스트"""
        messagebox.showinfo("테스트", "플로우 테스트 기능은 추후 구현 예정입니다.")
    
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
            
            # 키보드 동작 - 녹색
            'type_text': '#27ae60',
            'type_variable': '#27ae60',
            'key_press': '#27ae60',
            'paste': '#27ae60',
            
            # 제어 동작 - 빨간색
            'delay': '#e74c3c',
            'wait_image': '#e74c3c',
            
            # 지능형 동작 - 보라색 ← 추가
            'ocr_delay': '#9b59b6',
            
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
            
            # 키보드 동작
            'type_text': '⌨️ 키보드',
            'type_variable': '⌨️ 키보드',
            'key_press': '⌨️ 키보드',
            'paste': '⌨️ 키보드',
            
            # 제어 동작
            'delay': '⏱️ 제어',
            'wait_image': '⏱️ 제어',
            
            # 지능형 동작 ← 추가
            'ocr_delay': '🤖 지능형',
            
            # 기타
            'screenshot': '💾 기타',
        }
        
        return categories.get(action_type, '❓ 기타')

