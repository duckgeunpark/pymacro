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
            text=" 매크로 빌더", 
            font=("맑은 고딕", 24, "bold"),
            bg='#2c3e50',
            fg='white'
        )
        title_label.pack(pady=20)
        
        # 메인 컨텐츠
        content_frame = tk.Frame(self, bg='white')
        content_frame.pack(fill='both', expand=True, padx=40, pady=30)
        
        # 최근 프로젝트 섹션
        recent_label = tk.Label(
            content_frame,
            text="최근 프로젝트",
            font=("맑은 고딕", 14, "bold"),
            bg='white'
        )
        recent_label.pack(anchor='w', pady=(0, 10))
        
        # 최근 프로젝트 리스트
        self.recent_frame = tk.Frame(content_frame, bg='white')
        self.recent_frame.pack(fill='x', pady=(0, 30))
        
        # 구분선
        separator = ttk.Separator(content_frame, orient='horizontal')
        separator.pack(fill='x', pady=20)
        
        # 버튼 섹션
        button_frame = tk.Frame(content_frame, bg='white')
        button_frame.pack()
        
        # 새 프로젝트 버튼
        new_btn = tk.Button(
            button_frame,
            text="새 프로젝트 만들기",
            font=("맑은 고딕", 12),
            bg='#3498db',
            fg='white',
            padx=30,
            pady=15,
            cursor='hand2',
            command=self.create_new_project
        )
        new_btn.grid(row=0, column=0, padx=10)
        
        # 프로젝트 열기 버튼
        open_btn = tk.Button(
            button_frame,
            text="프로젝트 불러오기",
            font=("맑은 고딕", 12),
            bg='#2ecc71',
            fg='white',
            padx=30,
            pady=15,
            cursor='hand2',
            command=self.load_project
        )
        open_btn.grid(row=0, column=1, padx=10)
    
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
            no_project_label.pack(pady=10)
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
            no_project_label.pack(pady=10)
            return
        
        # 수정 시간 기준 정렬 (최근 5개)
        json_files.sort(
            key=lambda x: os.path.getmtime(os.path.join('projects', x)),
            reverse=True
        )
        recent_files = json_files[:5]
        
        for idx, filename in enumerate(recent_files):
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
        
        # 카드 프레임
        card = tk.Frame(
            self.recent_frame,
            bg='#ecf0f1',
            relief='raised',
            borderwidth=2
        )
        card.pack(fill='x', pady=5)
        
        # 좌측 정보 영역
        info_frame = tk.Frame(card, bg='#ecf0f1')
        info_frame.pack(side='left', fill='both', expand=True, padx=15, pady=10)
        
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
        button_frame.pack(side='right', padx=10, pady=10)
        
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
        
        # 제거 버튼 (새로 추가)
        remove_btn = tk.Button(
            button_frame,
            text="🗑️",
            font=("맑은 고딕", 9),
            bg='#e74c3c',
            fg='white',
            padx=8,
            pady=5,
            cursor='hand2',
            command=lambda: self.remove_project(filepath, project_name)
        )
        remove_btn.pack(side='left', padx=2)
    
    def remove_project(self, filepath, project_name):
        """프로젝트 제거"""
        # 확인 다이얼로그
        result = messagebox.askyesnocancel(
            "프로젝트 제거",
            f"'{project_name}' 프로젝트를 어떻게 처리하시겠습니까?\n\n"
            f"예: 파일 완전 삭제\n"
            f"아니오: 목록에서만 제거 (파일 유지)\n"
            f"취소: 작업 취소"
        )
        
        if result is None:  # 취소
            return
        elif result:  # 예 - 파일 삭제
            try:
                # JSON 파일 삭제
                os.remove(filepath)
                
                # 관련 이미지 파일들도 삭제
                try:
                    with open(filepath, 'r', encoding='utf-8') as f:
                        project_data = json.load(f)
                    
                    # 이미지 파일 삭제
                    images = project_data.get('images', [])
                    for img in images:
                        img_path = img.get('path', '')
                        if img_path and os.path.exists(img_path):
                            os.remove(img_path)
                except:
                    pass
                
                messagebox.showinfo("완료", f"'{project_name}' 프로젝트가 삭제되었습니다.")
                
            except Exception as e:
                messagebox.showerror("오류", f"프로젝트 삭제 중 오류가 발생했습니다:\n{str(e)}")
                return
        else:  # 아니오 - 목록에서만 제거
            # 파일을 다른 폴더로 이동하거나 숨김 속성 설정
            # 여기서는 파일명에 .hidden 추가
            try:
                hidden_path = filepath + '.hidden'
                os.rename(filepath, hidden_path)
                messagebox.showinfo("완료", f"'{project_name}' 프로젝트가 목록에서 제거되었습니다.\n(파일은 보존됨)")
            except Exception as e:
                messagebox.showerror("오류", f"프로젝트 제거 중 오류가 발생했습니다:\n{str(e)}")
                return
        
        # 목록 새로고침
        self.load_recent_projects()
    
    def open_project(self, filepath):
        """프로젝트 열기 (실행 화면)"""
        from ui.project_runner import ProjectRunner
        from core.project_manager import ProjectManager
        
        project_data = ProjectManager.load_project(filepath)
        if not project_data:
            messagebox.showerror("오류", "프로젝트를 불러올 수 없습니다.")
            return
        
        # 화면 전환
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
        
        # 화면 전환
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
            
            # 프로젝트 파일 생성
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
