"""
프로젝트 실행 화면 - 매크로 실행 및 모니터링
"""
import tkinter as tk
from tkinter import ttk, messagebox
import threading
from pynput import keyboard  # ← 추가
import main

from core.project_manager import ProjectManager
from core.coordinate_manager import CoordinateManager
from core.excel_manager import ExcelManager
from core.image_manager import ImageManager
from core.flow_manager import FlowManager
from core.executor import MacroExecutor

class ProjectRunner(tk.Frame):
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
        
        # 실행 엔진
        self.executor = MacroExecutor(
            project_data,
            self.coord_mgr,
            self.excel_mgr,
            self.image_mgr,
            self.flow_mgr
        )
        
        self.executor.set_callbacks(
            log_cb=self.add_log,
            progress_cb=self.update_progress,
            error_cb=self.report_error
        )
        
        self.is_running = False
        
        # 단축키 리스너
        self.hotkey_listener = None
        self.setup_hotkeys()
        
        self.setup_ui()
    
    def setup_hotkeys(self):
        """단축키 설정"""
        settings = self.project_data.get('settings', {})
        hotkeys = settings.get('hotkeys', {
            'start': 'f8',
            'pause': 'f9',
            'stop': 'f10',
            'focus': 'f12'
        })
        
        # 단축키를 pynput 형식으로 변환
        self.hotkey_map = {}
        for action, key in hotkeys.items():
            # F8, F9 등을 keyboard.Key.f8 형식으로 변환
            if key.lower().startswith('f') and key[1:].isdigit():
                # F키
                key_obj = getattr(keyboard.Key, key.lower(), None)
            else:
                # 일반 키
                key_obj = key.lower()
            
            if key_obj:
                self.hotkey_map[action] = key_obj
        
        # 리스너 시작
        self.start_hotkey_listener()
    
    def start_hotkey_listener(self):
        """단축키 리스너 시작"""
        def on_press(key):
            try:
                # F키 처리
                if hasattr(key, 'name'):
                    key_name = key
                else:
                    key_name = key.char.lower() if hasattr(key, 'char') else None
                
                # 단축키 매칭
                if key_name == self.hotkey_map.get('start'):
                    self.parent.after(0, self.start_macro)
                elif key_name == self.hotkey_map.get('pause'):
                    self.parent.after(0, self.pause_macro)
                elif key_name == self.hotkey_map.get('stop'):
                    self.parent.after(0, self.stop_macro)
                elif key_name == self.hotkey_map.get('focus'):
                    self.parent.after(0, self.bring_to_front)  # 맨 앞으로
            except Exception as e:
                print(f"단축키 처리 오류: {e}")
        
        # 리스너 시작
        self.hotkey_listener = keyboard.Listener(on_press=on_press)
        self.hotkey_listener.daemon = True
        self.hotkey_listener.start()
    
    def stop_hotkey_listener(self):
        """단축키 리스너 중지"""
        if self.hotkey_listener:
            self.hotkey_listener.stop()
    
    def setup_ui(self):
        """UI 구성"""
        # 상단 헤더
        header = tk.Frame(self, bg='#2c3e50', height=60)
        header.pack(fill='x')
        header.pack_propagate(False)
        
        tk.Label(
            header,
            text=f"▶️ {self.project_data['name']}",
            font=("맑은 고딕", 14, "bold"),
            bg='#2c3e50',
            fg='white'
        ).pack(side='left', padx=20, pady=15)
        
        btn_frame = tk.Frame(header, bg='#2c3e50')
        btn_frame.pack(side='right', padx=20)
        
        tk.Button(
            btn_frame,
            text="✏️ 편집",
            font=("맑은 고딕", 10),
            bg='#95a5a6',
            fg='white',
            padx=15,
            pady=5,
            command=self.edit_project
        ).pack(side='left', padx=5)
        
        tk.Button(
            btn_frame,
            text="⚙️ 설정",
            font=("맑은 고딕", 10),
            bg='#34495e',
            fg='white',
            padx=15,
            pady=5,
            command=self.show_settings
        ).pack(side='left', padx=5)
        
        tk.Button(
            btn_frame,
            text="🏠 홈",
            font=("맑은 고딕", 10),
            bg='#7f8c8d',
            fg='white',
            padx=15,
            pady=5,
            command=self.go_home
        ).pack(side='left', padx=5)
        
        # 메인 컨텐츠
        content = tk.Frame(self, bg='white')
        content.pack(fill='both', expand=True, padx=30, pady=20)
        
        # 단축키 정보 표시
        hotkey_frame = tk.LabelFrame(
            content,
            text="⌨️ 단축키",
            font=("맑은 고딕", 11, "bold"),
            bg='white',
            padx=15,
            pady=10
        )
        hotkey_frame.pack(fill='x', pady=(0, 10))

        hotkeys = self.project_data.get('settings', {}).get('hotkeys', {})
        hotkey_text = f"시작: {hotkeys.get('start', 'F9').upper()}  |  " \
                    f"일시정지: {hotkeys.get('pause', 'F10').upper()}  |  " \
                    f"중지: {hotkeys.get('stop', 'F11').upper()}  |  " \
                    f"맨 앞으로: {hotkeys.get('focus', 'F12').upper()}"

        tk.Label(
            hotkey_frame,
            text=hotkey_text,
            font=("맑은 고딕", 9),
            bg='white',
            fg='#2c3e50'
        ).pack()
        
        # 프로젝트 정보
        info_frame = tk.LabelFrame(
            content,
            text="📋 프로젝트 정보",
            font=("맑은 고딕", 11, "bold"),
            bg='white',
            padx=15,
            pady=15
        )
        info_frame.pack(fill='x', pady=(0, 20))
        
        info_grid = tk.Frame(info_frame, bg='white')
        info_grid.pack()
        
        tk.Label(info_grid, text="• 좌표:", font=("맑은 고딕", 10), bg='white').grid(row=0, column=0, sticky='w', padx=5)
        tk.Label(info_grid, text=f"{len(self.coord_mgr.coordinates)}개", font=("맑은 고딕", 10), bg='white').grid(row=0, column=1, sticky='w')
        
        tk.Label(info_grid, text="• 엑셀:", font=("맑은 고딕", 10), bg='white').grid(row=1, column=0, sticky='w', padx=5)
        excel_info = f"{len(self.excel_mgr.excel_sources)}개"
        if self.excel_mgr.excel_sources:
            excel_info += f" ({self.excel_mgr.excel_sources[0]['row_count']} rows)"
        tk.Label(info_grid, text=excel_info, font=("맑은 고딕", 10), bg='white').grid(row=1, column=1, sticky='w')
        
        tk.Label(info_grid, text="• 이미지:", font=("맑은 고딕", 10), bg='white').grid(row=2, column=0, sticky='w', padx=5)
        tk.Label(info_grid, text=f"{len(self.image_mgr.images)}개", font=("맑은 고딕", 10), bg='white').grid(row=2, column=1, sticky='w')
        
        tk.Label(info_grid, text="• 플로우:", font=("맑은 고딕", 10), bg='white').grid(row=3, column=0, sticky='w', padx=5)
        tk.Label(info_grid, text=f"{len(self.flow_mgr.flow_sequence)}개 액션", font=("맑은 고딕", 10), bg='white').grid(row=3, column=1, sticky='w')
        
        # 진행 상황
        progress_frame = tk.LabelFrame(
            content,
            text="📊 진행 상황",
            font=("맑은 고딕", 11, "bold"),
            bg='white',
            padx=15,
            pady=15
        )
        progress_frame.pack(fill='x', pady=(0, 20))
        
        self.progress_label = tk.Label(
            progress_frame,
            text="대기 중...",
            font=("맑은 고딕", 10),
            bg='white'
        )
        self.progress_label.pack(pady=(0, 10))
        
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            mode='determinate',
            length=500
        )
        self.progress_bar.pack(fill='x')
        
        self.status_label = tk.Label(
            progress_frame,
            text="",
            font=("맑은 고딕", 9),
            fg='gray',
            bg='white'
        )
        self.status_label.pack(pady=(5, 0))
        
        # 로그
        log_frame = tk.LabelFrame(
            content,
            text="📝 실행 로그",
            font=("맑은 고딕", 11, "bold"),
            bg='white',
            padx=15,
            pady=15
        )
        log_frame.pack(fill='both', expand=True, pady=(0, 20))
        
        log_scroll = tk.Scrollbar(log_frame)
        log_scroll.pack(side='right', fill='y')
        
        self.log_text = tk.Text(
            log_frame,
            font=("Consolas", 9),
            height=10,
            yscrollcommand=log_scroll.set,
            bg='#f8f9fa',
            state='disabled'
        )
        self.log_text.pack(fill='both', expand=True)
        log_scroll.config(command=self.log_text.yview)
        
        # 컨트롤 버튼
        control_frame = tk.Frame(content, bg='white')
        control_frame.pack()
        
        self.start_btn = tk.Button(
            control_frame,
            text="▶️ 시작",
            font=("맑은 고딕", 12, "bold"),
            bg='#27ae60',
            fg='white',
            padx=40,
            pady=15,
            command=self.start_macro
        )
        self.start_btn.pack(side='left', padx=5)
        
        self.pause_btn = tk.Button(
            control_frame,
            text="⏸️ 일시정지",
            font=("맑은 고딕", 12, "bold"),
            bg='#f39c12',
            fg='white',
            padx=30,
            pady=15,
            state='disabled',
            command=self.pause_macro
        )
        self.pause_btn.pack(side='left', padx=5)
        
        self.stop_btn = tk.Button(
            control_frame,
            text="⏹️ 중지",
            font=("맑은 고딕", 12, "bold"),
            bg='#e74c3c',
            fg='white',
            padx=40,
            pady=15,
            state='disabled',
            command=self.stop_macro
        )
        self.stop_btn.pack(side='left', padx=5)
    
    def add_log(self, message):
        """로그 추가"""
        self.log_text.config(state='normal')
        self.log_text.insert('end', message + '\n')
        self.log_text.see('end')
        self.log_text.config(state='disabled')
    
    def update_progress(self, current, total, status=""):
        """진행상황 업데이트"""
        if total > 0:
            percentage = (current / total) * 100
            self.progress_bar['value'] = percentage
            self.progress_label.config(text=f"{current}/{total} ({percentage:.1f}%)")
        else:
            self.progress_bar['mode'] = 'indeterminate'
            self.progress_bar.start()
            self.progress_label.config(text=f"반복 {current}")
        
        self.status_label.config(text=status)
    
    def report_error(self, error_msg, screenshot=None):
        """에러 보고"""
        pass
    
    def start_macro(self):
        """매크로 시작"""
        if self.is_running:
            return
        
        # 유효성 검사
        if not self.flow_mgr.flow_sequence:
            messagebox.showerror("오류", "플로우가 비어있습니다.")
            return
        
        settings = self.project_data.get('settings', {}).get('execution', {})
        if settings.get('mode') == 'excel_loop' and not self.excel_mgr.excel_sources:
            messagebox.showerror("오류", "엑셀 행 반복 모드는 엑셀 데이터가 필요합니다.")
            return
        
        # 버튼 상태 변경
        self.start_btn.config(state='disabled')
        self.pause_btn.config(state='normal')
        self.stop_btn.config(state='normal')
        
        self.is_running = True
        self.add_log("="*50)
        self.add_log("매크로 실행 준비 중...")
        
        # 3초 카운트다운
        self.countdown(3)
    
    def countdown(self, count):
        """카운트다운"""
        if count > 0:
            self.add_log(f"시작까지 {count}초...")
            self.after(1000, self.countdown, count-1)
        else:
            self.add_log("시작!")
            # 별도 스레드에서 실행
            thread = threading.Thread(target=self.run_executor, daemon=True)
            thread.start()
    
    def run_executor(self):
        """실행 엔진 실행"""
        try:
            self.executor.start()
            self.after(100, self.on_execution_finished)
        except Exception as e:
            self.add_log(f"실행 오류: {str(e)}")
            self.after(100, self.on_execution_finished)
    
    def on_execution_finished(self):
        """실행 완료"""
        self.is_running = False
        self.start_btn.config(state='normal')
        self.pause_btn.config(state='disabled')
        self.stop_btn.config(state='disabled')
        
        messagebox.showinfo("완료", "매크로 실행이 완료되었습니다!")
    
    def pause_macro(self):
        """매크로 일시정지/재개"""
        if self.executor.is_paused:
            self.executor.resume()
            self.pause_btn.config(text="⏸️ 일시정지")
        else:
            self.executor.pause()
            self.pause_btn.config(text="▶️ 재개")
    
    def stop_macro(self):
        """매크로 중지 (확인 없이 즉시 중지)"""
        self.executor.stop()
        self.add_log("⏹️ 사용자가 중지를 요청했습니다.")

    def bring_to_front(self):
        """프로그램 창을 맨 앞으로 가져오기"""
        try:
            # 창을 최소화 상태에서 복원
            self.parent.state('normal')
            
            # 맨 앞으로 가져오기
            self.parent.lift()
            self.parent.focus_force()
            self.parent.attributes('-topmost', True)
            self.parent.after(100, lambda: self.parent.attributes('-topmost', False))
            
            self.add_log("🔼 프로그램이 맨 앞으로 이동했습니다")
        except Exception as e:
            print(f"맨 앞으로 가져오기 오류: {e}")
    
    def edit_project(self):
        """프로젝트 편집"""
        self.stop_hotkey_listener()
        
        from ui.project_editor import ProjectEditor
        
        for widget in self.parent.winfo_children():
            widget.destroy()
        
        editor = ProjectEditor(self.parent, self.app, self.project_data, self.filepath)
        editor.pack(fill='both', expand=True)
    
    def show_settings(self):
        """설정 창"""
        dialog = tk.Toplevel(self.parent)
        dialog.title("실행 설정")
        dialog.geometry("450x700")  # 높이 증가
        dialog.transient(self.parent)
        dialog.grab_set()
        
        try:
            if hasattr(main, 'ICON_PATH') and main.ICON_PATH:
                dialog.iconbitmap(main.ICON_PATH)
        except:
            pass

        tk.Label(
            dialog,
            text="⚙️ 실행 설정",
            font=("맑은 고딕", 14, "bold")
        ).pack(pady=15)
        
        # 단축키 설정
        hotkey_frame = tk.LabelFrame(
            dialog,
            text="⌨️ 단축키 설정",
            font=("맑은 고딕", 11, "bold"),
            padx=20,
            pady=15
        )
        hotkey_frame.pack(fill='x', padx=20, pady=10)
        
        hotkeys = self.project_data.get('settings', {}).get('hotkeys', {
            'start': 'F9',
            'pause': 'F10',
            'stop': 'F11',
            'focus': 'F12'
        })
        
        hotkey_entries = {}
        
        # 단축키 목록에 '맨 앞으로' 추가
        for idx, (action, label) in enumerate([
            ('start', '시작'),
            ('pause', '일시정지'),
            ('stop', '중지'),
            ('focus', '맨 앞으로')  # 추가
        ]):
            tk.Label(
                hotkey_frame,
                text=f"{label}:",
                font=("맑은 고딕", 10)
            ).grid(row=idx, column=0, sticky='w', pady=5)
            
            entry = tk.Entry(hotkey_frame, font=("맑은 고딕", 10), width=15)
            entry.insert(0, hotkeys.get(action, ''))
            entry.grid(row=idx, column=1, padx=10, pady=5)
            hotkey_entries[action] = entry
        
        tk.Label(
            hotkey_frame,
            text="※ 예: F9, F10, F11, F12, ctrl+shift+s 등",
            font=("맑은 고딕", 8),
            fg='gray'
        ).grid(row=4, column=0, columnspan=2, pady=(10, 0))
        
        # 실행 모드
        mode_frame = tk.LabelFrame(
            dialog,
            text="🎯 실행 모드",
            font=("맑은 고딕", 11, "bold"),
            padx=20,
            pady=15
        )
        mode_frame.pack(fill='x', padx=20, pady=10)
        
        mode_var = tk.StringVar(
            value=self.project_data.get('settings', {}).get('execution', {}).get('mode', 'flow_repeat')
        )
        
        modes = [
            ('flow_repeat', '플로우 반복 실행'),
            ('excel_loop', '엑셀 행 반복'),
            ('infinite', '무한 반복 (중지할 때까지)')
        ]
        
        for value, text in modes:
            tk.Radiobutton(
                mode_frame,
                text=text,
                variable=mode_var,
                value=value,
                font=("맑은 고딕", 9)
            ).pack(anchor='w', padx=10, pady=2)
        
        # 반복 횟수
        repeat_frame = tk.Frame(mode_frame)
        repeat_frame.pack(fill='x', padx=10, pady=(10, 0))
        
        tk.Label(
            repeat_frame,
            text="플로우 반복 횟수:",
            font=("맑은 고딕", 9)
        ).pack(side='left')
        
        repeat_entry = tk.Entry(repeat_frame, font=("맑은 고딕", 9), width=10)
        repeat_entry.insert(0, str(
            self.project_data.get('settings', {}).get('execution', {}).get('repeat_count', 1)
        ))
        repeat_entry.pack(side='left', padx=10)
        
        # 엑셀 무한반복 체크박스 추가
        excel_infinite_var = tk.BooleanVar(
            value=self.project_data.get('settings', {}).get('execution', {}).get('excel_infinite_loop', False)
        )
        
        excel_infinite_check = tk.Checkbutton(
            mode_frame,
            text="🔄 엑셀 행 무한반복 (마지막 행 후 처음부터 다시)",
            variable=excel_infinite_var,
            font=("맑은 고딕", 9, "bold"),
            fg='#e74c3c'
        )
        excel_infinite_check.pack(anchor='w', padx=10, pady=(15, 5))
        
        tk.Label(
            mode_frame,
            text="※ 엑셀 모드에서만 적용됩니다",
            font=("맑은 고딕", 8),
            fg='gray'
        ).pack(anchor='w', padx=30)
        
        def save_settings():
            # 단축키 저장
            if 'settings' not in self.project_data:
                self.project_data['settings'] = {}
            if 'hotkeys' not in self.project_data['settings']:
                self.project_data['settings']['hotkeys'] = {}
            
            for action, entry in hotkey_entries.items():
                self.project_data['settings']['hotkeys'][action] = entry.get().strip()
            
            # 실행 설정 저장
            if 'execution' not in self.project_data['settings']:
                self.project_data['settings']['execution'] = {}
            
            self.project_data['settings']['execution']['mode'] = mode_var.get()
            self.project_data['settings']['execution']['excel_infinite_loop'] = excel_infinite_var.get()
            
            try:
                repeat_count = int(repeat_entry.get())
                self.project_data['settings']['execution']['repeat_count'] = repeat_count
            except:
                pass
            
            # 저장
            ProjectManager.save_project(self.filepath, self.project_data)
            
            # 단축키 리스너 재시작
            self.stop_hotkey_listener()
            self.setup_hotkeys()
            
            dialog.destroy()
            messagebox.showinfo("완료", "설정이 저장되었습니다!\n단축키가 적용되었습니다.")
        
        tk.Button(
            dialog,
            text="💾 저장",
            command=save_settings,
            bg='#27ae60',
            fg='white',
            padx=30,
            pady=10,
            font=("맑은 고딕", 11, "bold")
        ).pack(pady=20)

    
    def update_hotkey_display(self):
        """단축키 표시 업데이트"""
        # UI의 단축키 정보를 업데이트하려면 화면을 다시 그려야 함
        # 여기서는 간단히 메시지로 알림
        pass
    
    def go_home(self):
        """홈으로"""
        if self.is_running:
            messagebox.showwarning("경고", "매크로 실행 중에는 홈으로 갈 수 없습니다.")
            return
        
        self.stop_hotkey_listener()
        self.app.show_start_screen()
