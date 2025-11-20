"""
시작 화면 - 프로젝트 생성/불러오기
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import json
from datetime import datetime



class StartScreen(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.parent = parent
        
        self.setup_ui()
        self.load_recent_projects()
    
    def setup_ui(self):
        """UI 구성"""
        # 제목
        title_frame = tk.Frame(self, bg='#2c3e50', height=80)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)
        
        title_label = tk.Label(
            title_frame, 
            text="dMax MacroBuilder", 
            font=("맑은 고딕", 24, "bold"),
            bg='#2c3e50',
            fg='white'
        )
        title_label.pack(pady=20)
        
        # 버튼 섹션 (먼저 pack - 맨 아래 고정) ⭐
        button_frame = tk.Frame(self, bg='#F0F0F0')
        button_frame.pack(side='bottom', fill='x', padx=20, pady=(0, 20))
        
        # 새 프로젝트 버튼
        new_btn = tk.Button(
            button_frame,
            text="새 프로젝트 만들기",
            font=("맑은 고딕", 14, "bold"),
            bg='#3498db',
            fg='white',
            padx=30,
            pady=15,
            cursor='hand2',
            command=self.create_new_project
        )
        new_btn.pack(side='left', expand=True, fill='both', padx=(0, 10))
        
        # 프로젝트 열기 버튼
        open_btn = tk.Button(
            button_frame,
            text="프로젝트 불러오기",
            font=("맑은 고딕", 14, "bold"),
            bg='#2ecc71',
            fg='white',
            padx=30,
            pady=15,
            cursor='hand2',
            command=self.load_project
        )
        open_btn.pack(side='left', expand=True, fill='both', padx=(10, 0))
        
        # 메인 컨텐츠 (나머지 공간 차지) ⭐
        content_frame = tk.Frame(self, bg='white')
        content_frame.pack(fill='both', expand=True, padx=20, pady=(30, 20))
        
        # 최근 프로젝트 섹션
        recent_label = tk.Label(
            content_frame,
            text="최근 프로젝트",
            font=("맑은 고딕", 14, "bold"),
            bg='white'
        )
        recent_label.pack(fill='x', anchor='w', padx=20, pady=10)
        
        # 최근 프로젝트 리스트 (스크롤 추가) ⭐
        self.setup_scrollable_projects(content_frame)
    
    def setup_scrollable_projects(self, parent):
        """스크롤 가능한 프로젝트 리스트 생성"""
        # 컨테이너 프레임
        container = tk.Frame(parent, bg='white')
        container.pack(fill='both', expand=True)
        
        # 캔버스
        canvas = tk.Canvas(container, bg='white', highlightthickness=0)
        scrollbar = tk.Scrollbar(container, orient='vertical', command=canvas.yview)
        
        # 스크롤 가능한 프레임
        self.recent_frame = tk.Frame(canvas, bg='white')
        
        canvas_id = canvas.create_window((0, 0), window=self.recent_frame, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # ⭐ 핵심: 프레임 너비를 캔버스 너비에 동기화
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        def on_canvas_configure(event):
            # 캔버스 너비에 맞춰 프레임 너비 설정
            canvas_width = event.width
            self.recent_frame.config(width=canvas_width)
        
        self.recent_frame.bind("<Configure>", on_frame_configure)
        canvas.bind("<Configure>", on_canvas_configure)
        
        # 마우스 휠 스크롤 지원
        def _on_mousewheel(event):
            try:
                # 캔버스가 아직 존재하는지 확인
                if canvas.winfo_exists():
                    canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except tk.TclError:
                # 캔버스가 삭제되었을 때 에러 무시
                pass
        
        canvas.bind("<MouseWheel>", _on_mousewheel)
        
        # 배치
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
    
    def load_recent_projects(self):
        """최근 프로젝트 목록 로드"""
        # 기존 위젯 제거
        for widget in self.recent_frame.winfo_children():
            widget.destroy()
        
        # projects 폴더에서 .json 파일 찾기
        if not os.path.exists('projects'):
            no_project_label = tk.Label(
                self.recent_frame,
                text="최근 프로젝트가 없습니다.",
                font=("맑은 고딕", 10),
                fg='gray',
                bg='white'
            )
            no_project_label.pack(fill='x',pady=10)
            return
        
        json_files = [f for f in os.listdir('projects') if f.endswith('.json')]
        
        if not json_files:
            no_project_label = tk.Label(
                self.recent_frame,
                text="최근 프로젝트가 없습니다.",
                font=("맑은 고딕", 10),
                fg='gray',
                bg='white'
            )
            no_project_label.pack(fill='x',pady=10)
            return
        
        # 수정 시간 기준 정렬
        json_files.sort(
            key=lambda x: os.path.getmtime(os.path.join('projects', x)),
            reverse=True
        )
        
        # 전체 프로젝트 표시 (5개 제한 제거) ⭐
        for idx, filename in enumerate(json_files):
            self.create_project_card(filename, idx)
    
    def create_project_card(self, filename, index):
        """프로젝트 카드 생성"""
        filepath = os.path.join('projects', filename)
        
        # 프로젝트 정보 읽기
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                project_data = json.load(f)
            
            project_name = project_data.get('name', filename.replace('.json', ''))
            description = project_data.get('description', '설명 없음')
            modified_time = datetime.fromtimestamp(os.path.getmtime(filepath))
            
        except Exception as e:
            print(f"프로젝트 로드 오류: {e}")
            return
        
        # 카드 프레임 (⭐ fill='both', expand=True 추가)
        card = tk.Frame(
            self.recent_frame,
            bg='#ecf0f1',
            relief='raised',
            borderwidth=1
        )
        card.pack(fill='both', expand=True, padx=15,pady=(0,10))  # ⭐ padx=0으로 변경
        
        # 좌측 정보 영역
        info_frame = tk.Frame(card, bg='#ecf0f1')
        info_frame.pack(side='left', fill='both', expand=True, padx=(20,30), pady=10)

        # 프로젝트 이름
        name_label = tk.Label(
            info_frame,
            text=f"📁 {project_name}",
            font=("맑은 고딕", 12, "bold"),
            bg='#ecf0f1',
            anchor='w'
        )
        name_label.pack(anchor='w')

        # 설명
        desc_label = tk.Label(
            info_frame,
            text=description,
            font=("맑은 고딕", 9),
            fg='#7f8c8d',
            bg='#ecf0f1',
            anchor='w'
        )
        desc_label.pack(anchor='w', pady=(2, 0))

        # 수정 시간
        time_label = tk.Label(
            info_frame,
            text=f"마지막 수정: {modified_time.strftime('%Y-%m-%d %H:%M')}",
            font=("맑은 고딕", 8),
            fg='#95a5a6',
            bg='#ecf0f1',
            anchor='w'
        )
        time_label.pack(anchor='w', pady=(2, 0))

        # 우측 버튼 영역
        button_frame = tk.Frame(card, bg='#ecf0f1')
        button_frame.pack(side='right', padx=(30,20), pady=10)

        # 열기 버튼
        open_btn = tk.Button(
            button_frame,
            text="열기",
            font=("맑은 고딕", 9),
            bg='#3498db',
            fg='white',
            padx=15,
            pady=5,
            cursor='hand2',
            command=lambda: self.open_project(filepath)
        )
        open_btn.pack(side='left', padx=2)
        
        # 편집 버튼
        edit_btn = tk.Button(
            button_frame,
            text="편집",
            font=("맑은 고딕", 9),
            bg='#95a5a6',
            fg='white',
            padx=15,
            pady=5,
            cursor='hand2',
            command=lambda: self.edit_project(filepath)
        )
        edit_btn.pack(side='left', padx=2)
        
        # 제거 버튼
        remove_btn = tk.Button(
            button_frame,
            text="제거",
            font=("맑은 고딕", 9),
            bg='#e74c3c',
            fg='white',
            padx=15,
            pady=5,
            cursor='hand2',
            command=lambda: self.remove_project(filepath, project_name)
        )
        remove_btn.pack(side='left', padx=2)
    
    def remove_project(self, filepath, project_name):
        """프로젝트 제거"""
        result = messagebox.askyesnocancel(
            "프로젝트 제거",
            f"'{project_name}' 프로젝트를 어떻게 처리하시겠습니까?\n\n"
            f"예: 파일 완전 삭제\n"
            f"아니오: 목록에서만 제거 (파일 유지)\n"
            f"취소: 작업 취소"
        )
        
        if result is None:
            return
        elif result:
            try:
                os.remove(filepath)
                messagebox.showinfo("완료", f"'{project_name}' 프로젝트가 삭제되었습니다.")
            except Exception as e:
                messagebox.showerror("오류", f"프로젝트 삭제 중 오류가 발생했습니다:\n{str(e)}")
                return
        else:
            try:
                hidden_path = filepath + '.hidden'
                os.rename(filepath, hidden_path)
                messagebox.showinfo("완료", f"'{project_name}' 프로젝트가 목록에서 제거되었습니다.\n(파일은 보존됨)")
            except Exception as e:
                messagebox.showerror("오류", f"프로젝트 제거 중 오류가 발생했습니다:\n{str(e)}")
                return
        
        self.load_recent_projects()
    
    def open_project(self, filepath):
        """프로젝트 열기 (실행 화면)"""
        from ui.project_runner import ProjectRunner
        from core.project_manager import ProjectManager
        
        project_data = ProjectManager.load_project(filepath)
        if not project_data:
            messagebox.showerror("오류", "프로젝트를 불러올 수 없습니다.")
            return
        
        for widget in self.parent.winfo_children():
            widget.destroy()
        
        runner = ProjectRunner(self.parent, self.app, project_data, filepath)
        runner.pack(fill='both', expand=True)
    
    def edit_project(self, filepath):
        """프로젝트 편집"""
        from ui.project_editor import ProjectEditor
        from core.project_manager import ProjectManager
        
        project_data = ProjectManager.load_project(filepath)
        if not project_data:
            messagebox.showerror("오류", "프로젝트를 불러올 수 없습니다.")
            return
        
        for widget in self.parent.winfo_children():
            widget.destroy()
        
        editor = ProjectEditor(self.parent, self.app, project_data, filepath)
        editor.pack(fill='both', expand=True)
    
    def create_new_project(self):
        """새 프로젝트 만들기"""
        from ui.dialogs import NewProjectDialog
        from core.project_manager import ProjectManager
        
        dialog = NewProjectDialog(self.parent)
        self.parent.wait_window(dialog)
        
        if dialog.result:
            name = dialog.result['name']
            description = dialog.result['description']
            
            filename = f"{name.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            filepath = os.path.join('projects', filename)
            
            project_data = ProjectManager.create_empty_project(name, description)
            
            if ProjectManager.save_project(filepath, project_data):
                messagebox.showinfo("완료", f"프로젝트 '{name}'이(가) 생성되었습니다!")
                self.edit_project(filepath)
            else:
                messagebox.showerror("오류", "프로젝트 생성에 실패했습니다.")
    
    def load_project(self):
        """프로젝트 불러오기"""
        filepath = filedialog.askopenfilename(
            title="프로젝트 선택",
            initialdir="projects",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if filepath:
            self.open_project(filepath)
