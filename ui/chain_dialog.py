"""
매크로 체인 설정 다이얼로그
"""
import tkinter as tk
from tkinter import messagebox
from utils.ui_helpers import set_dialog_icon, center_window_on_parent


class MacroChainDialog(tk.Toplevel):
    """매크로 체인 설정 다이얼로그"""

    def __init__(self, parent, project_manager):
        super().__init__(parent)
        self.title("매크로 체인 실행 설정")
        self.geometry("500x600")
        self.resizable(False, False)

        self.project_manager = project_manager
        self.result = None
        self.chain_items = []  # [{'project_name': str, 'filepath': str, 'repeat_count': int}]

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
            text="🔗 매크로 체인 실행",
            font=("맑은 고딕", 14, "bold")
        ).pack(pady=(0, 10))

        tk.Label(
            main_frame,
            text="여러 매크로를 순차적으로 실행합니다.\n각 매크로의 반복 횟수를 설정할 수 있습니다.",
            font=("맑은 고딕", 9),
            fg='#7f8c8d',
            justify='left'
        ).pack(anchor='w', pady=(0, 15))

        # 체인 리스트
        list_frame = tk.LabelFrame(
            main_frame,
            text="체인 순서",
            font=("맑은 고딕", 11, "bold"),
            padx=10,
            pady=10
        )
        list_frame.pack(fill='both', expand=True, pady=(0, 10))

        # 스크롤 가능한 리스트
        canvas = tk.Canvas(list_frame, bg='white', highlightthickness=1, highlightbackground='#bdc3c7')
        scrollbar = tk.Scrollbar(list_frame, orient='vertical', command=canvas.yview)
        self.chain_list_frame = tk.Frame(canvas, bg='white')

        self.chain_list_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.chain_list_frame, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        # 버튼들
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(fill='x', pady=(10, 0))

        tk.Button(
            btn_frame,
            text="+ 매크로 추가",
            font=("맑은 고딕", 10),
            bg='#3498db',
            fg='white',
            padx=15,
            pady=8,
            command=self.add_macro_to_chain
        ).pack(side='left', padx=5)

        tk.Button(
            btn_frame,
            text="✅ 실행",
            font=("맑은 고딕", 10),
            bg='#27ae60',
            fg='white',
            padx=20,
            pady=8,
            command=self.on_execute
        ).pack(side='right', padx=5)

        tk.Button(
            btn_frame,
            text="취소",
            font=("맑은 고딕", 10),
            bg='#95a5a6',
            fg='white',
            padx=20,
            pady=8,
            command=self.on_cancel
        ).pack(side='right', padx=5)

        self.refresh_chain_list()

    def add_macro_to_chain(self):
        """체인에 매크로 추가"""
        # 프로젝트 선택 다이얼로그
        projects = self.project_manager.get_project_list()

        if not projects:
            messagebox.showwarning("경고", "등록된 프로젝트가 없습니다.", parent=self)
            return

        # 프로젝트 선택 다이얼로그
        dialog = tk.Toplevel(self)
        dialog.title("매크로 선택 및 실행 설정")
        dialog.geometry("520x700")
        dialog.transient(self)
        dialog.grab_set()
        dialog.attributes('-topmost', True)
        set_dialog_icon(dialog)

        result = [None]

        tk.Label(
            dialog,
            text="실행할 매크로를 선택하세요",
            font=("맑은 고딕", 12, "bold")
        ).pack(pady=15)

        # 프로젝트 리스트박스
        listbox_frame = tk.Frame(dialog)
        listbox_frame.pack(fill='both', expand=True, padx=20, pady=10)

        scrollbar = tk.Scrollbar(listbox_frame)
        scrollbar.pack(side='right', fill='y')

        project_listbox = tk.Listbox(
            listbox_frame,
            font=("맑은 고딕", 10),
            yscrollcommand=scrollbar.set,
            selectmode='single'
        )
        project_listbox.pack(side='left', fill='both', expand=True)
        scrollbar.config(command=project_listbox.yview)

        # 프로젝트 필터링 (type='project'만 선택)
        filtered_projects = [proj for proj in projects if proj.get('type') != 'chain']

        for proj in filtered_projects:
            project_listbox.insert(tk.END, proj['name'])

        # 선택된 프로젝트의 실행 설정 표시 영역
        settings_frame = tk.LabelFrame(
            dialog,
            text="실행 설정 (수정 가능)",
            font=("맑은 고딕", 10, "bold"),
            padx=20,
            pady=15
        )
        settings_frame.pack(fill='x', padx=20, pady=10)

        # 설정 정보를 담을 변수들
        mode_var = tk.StringVar(value='flow_repeat')
        repeat_var = tk.StringVar(value='1')
        excel_infinite_var = tk.BooleanVar(value=False)

        # 설정 UI 컨테이너
        settings_container = tk.Frame(settings_frame, bg='white')
        settings_container.pack(fill='x')

        # 실행 모드 라디오 버튼들
        modes = [
            ('flow_repeat', '플로우 반복 실행'),
            ('excel_loop', '엑셀 행 반복'),
            ('infinite', '무한 반복 (중지할 때까지)')
        ]

        for value, text in modes:
            tk.Radiobutton(
                settings_container,
                text=text,
                variable=mode_var,
                value=value,
                font=("맑은 고딕", 9),
                bg='white'
            ).pack(anchor='w', padx=10, pady=3)

        # 반복 횟수 입력
        repeat_frame = tk.Frame(settings_container, bg='white')
        repeat_frame.pack(fill='x', padx=10, pady=(10, 5))

        tk.Label(
            repeat_frame,
            text="플로우 반복 횟수:",
            font=("맑은 고딕", 9),
            bg='white'
        ).pack(side='left')

        repeat_entry = tk.Entry(repeat_frame, textvariable=repeat_var, font=("맑은 고딕", 9), width=10)
        repeat_entry.pack(side='left', padx=10)

        # 엑셀 무한반복 체크박스
        excel_check = tk.Checkbutton(
            settings_container,
            text="🔄 엑셀 행 무한반복 (마지막 행 후 처음부터 다시)",
            variable=excel_infinite_var,
            font=("맑은 고딕", 9),
            bg='white'
        )
        excel_check.pack(anchor='w', padx=10, pady=(10, 5))

        tk.Label(
            settings_container,
            text="※ 선택한 프로젝트의 현재 설정입니다. 수정하면 프로젝트 파일에 저장됩니다.",
            font=("맑은 고딕", 8),
            fg='#7f8c8d',
            bg='white'
        ).pack(anchor='w', padx=10, pady=(5, 0))

        # 프로젝트 선택 시 설정 로드
        def on_project_select(event):
            selection = project_listbox.curselection()
            if selection:
                selected_project = filtered_projects[selection[0]]
                import json

                # 프로젝트 파일 로드
                try:
                    with open(selected_project['filepath'], 'r', encoding='utf-8') as f:
                        project_data = json.load(f)

                    # 설정값 로드
                    settings = project_data.get('settings', {})
                    execution = settings.get('execution', {})

                    mode_var.set(execution.get('mode', 'flow_repeat'))
                    repeat_var.set(str(execution.get('repeat_count', 1)))
                    excel_infinite_var.set(execution.get('excel_infinite_loop', False))

                except Exception as e:
                    print(f"프로젝트 설정 로드 실패: {e}")

        project_listbox.bind('<<ListboxSelect>>', on_project_select)

        def on_ok():
            selection = project_listbox.curselection()
            if not selection:
                messagebox.showwarning("경고", "프로젝트를 선택하세요.", parent=dialog)
                return

            import json
            from core.project_manager import ProjectManager

            selected_project = filtered_projects[selection[0]]

            # 반복 횟수 검증
            try:
                repeat_count = int(repeat_var.get())
                if repeat_count < 1:
                    raise ValueError
            except ValueError:
                messagebox.showerror("오류", "반복 횟수는 1 이상의 정수여야 합니다.", parent=dialog)
                return

            # 프로젝트 파일 로드 및 설정 저장
            try:
                with open(selected_project['filepath'], 'r', encoding='utf-8') as f:
                    project_data = json.load(f)

                # 설정 업데이트
                if 'settings' not in project_data:
                    project_data['settings'] = {}
                if 'execution' not in project_data['settings']:
                    project_data['settings']['execution'] = {}

                project_data['settings']['execution']['mode'] = mode_var.get()
                project_data['settings']['execution']['repeat_count'] = repeat_count
                project_data['settings']['execution']['excel_infinite_loop'] = excel_infinite_var.get()

                # 프로젝트 파일에 저장
                ProjectManager.save_project(selected_project['filepath'], project_data)

                result[0] = {
                    'project_name': selected_project['name'],
                    'filepath': selected_project['filepath'],
                    'execution_mode': 'use_project_settings'  # 항상 프로젝트 설정 사용
                }

                dialog.destroy()

            except Exception as e:
                messagebox.showerror("오류", f"설정 저장 실패: {e}", parent=dialog)

        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=15)

        tk.Button(
            btn_frame,
            text="추가",
            font=("맑은 고딕", 10),
            bg='#27ae60',
            fg='white',
            padx=20,
            pady=8,
            command=on_ok
        ).pack(side='left', padx=5)

        tk.Button(
            btn_frame,
            text="취소",
            font=("맑은 고딕", 10),
            bg='#95a5a6',
            fg='white',
            padx=20,
            pady=8,
            command=dialog.destroy
        ).pack(side='left', padx=5)

        center_window_on_parent(dialog, self)
        dialog.lift()
        dialog.focus_force()

        self.wait_window(dialog)

        if result[0]:
            self.chain_items.append(result[0])
            self.refresh_chain_list()

    def refresh_chain_list(self):
        """체인 리스트 새로고침"""
        # 기존 위젯 제거
        for widget in self.chain_list_frame.winfo_children():
            widget.destroy()

        if not self.chain_items:
            tk.Label(
                self.chain_list_frame,
                text="매크로를 추가하세요",
                font=("맑은 고딕", 10),
                fg='#95a5a6',
                bg='white'
            ).pack(pady=50)
            return

        for idx, item in enumerate(self.chain_items):
            self.create_chain_item_widget(idx, item)

    def create_chain_item_widget(self, idx, item):
        """체인 아이템 위젯 생성"""
        item_frame = tk.Frame(self.chain_list_frame, bg='white', relief='solid', borderwidth=1)
        item_frame.pack(fill='x', padx=5, pady=5)

        # 순서 번호
        tk.Label(
            item_frame,
            text=f"{idx + 1}.",
            font=("맑은 고딕", 11, "bold"),
            bg='white',
            fg='#3498db',
            width=3
        ).pack(side='left', padx=(10, 5))

        # 정보
        info_frame = tk.Frame(item_frame, bg='white')
        info_frame.pack(side='left', fill='x', expand=True, padx=10, pady=10)

        tk.Label(
            info_frame,
            text=item['project_name'],
            font=("맑은 고딕", 10, "bold"),
            bg='white',
            anchor='w'
        ).pack(anchor='w')

        # 실행 모드 표시
        execution_mode = item.get('execution_mode', 'custom_repeat')
        if execution_mode == 'use_project_settings':
            mode_text = "프로젝트 설정 사용"
        else:
            repeat_count = item.get('repeat_count', 1)
            mode_text = f"반복: {repeat_count}회" if repeat_count else "프로젝트 설정 사용"

        tk.Label(
            info_frame,
            text=mode_text,
            font=("맑은 고딕", 9),
            bg='white',
            fg='#7f8c8d',
            anchor='w'
        ).pack(anchor='w')

        # 버튼들
        btn_frame = tk.Frame(item_frame, bg='white')
        btn_frame.pack(side='right', padx=10)

        # 위로 이동
        if idx > 0:
            tk.Button(
                btn_frame,
                text="▲",
                font=("맑은 고딕", 8),
                width=3,
                command=lambda: self.move_item_up(idx)
            ).pack(side='left', padx=2)

        # 아래로 이동
        if idx < len(self.chain_items) - 1:
            tk.Button(
                btn_frame,
                text="▼",
                font=("맑은 고딕", 8),
                width=3,
                command=lambda: self.move_item_down(idx)
            ).pack(side='left', padx=2)

        # 삭제
        tk.Button(
            btn_frame,
            text="✕",
            font=("맑은 고딕", 8),
            bg='#e74c3c',
            fg='white',
            width=3,
            command=lambda: self.remove_item(idx)
        ).pack(side='left', padx=2)

    def move_item_up(self, idx):
        """아이템 위로 이동"""
        if idx > 0:
            self.chain_items[idx], self.chain_items[idx-1] = self.chain_items[idx-1], self.chain_items[idx]
            self.refresh_chain_list()

    def move_item_down(self, idx):
        """아이템 아래로 이동"""
        if idx < len(self.chain_items) - 1:
            self.chain_items[idx], self.chain_items[idx+1] = self.chain_items[idx+1], self.chain_items[idx]
            self.refresh_chain_list()

    def remove_item(self, idx):
        """아이템 삭제"""
        del self.chain_items[idx]
        self.refresh_chain_list()

    def on_execute(self):
        """실행 버튼"""
        if not self.chain_items:
            messagebox.showwarning("경고", "실행할 매크로를 추가하세요.", parent=self)
            return

        # 체인 저장 여부 묻기
        save = messagebox.askyesnocancel(
            "체인 저장",
            "이 체인을 저장하시겠습니까?\n\n"
            "예: 저장 후 실행\n"
            "아니오: 저장 없이 실행\n"
            "취소: 돌아가기",
            parent=self
        )

        if save is None:  # 취소
            return
        elif save:  # 저장
            if self.save_chain():
                self.result = self.chain_items
                self.destroy()
        else:  # 저장 안 함
            self.result = self.chain_items
            self.destroy()

    def save_chain(self):
        """체인 저장"""
        from ui.dialogs import NameInputDialog
        from core.chain_manager import ChainManager

        # 이름 입력
        name_dialog = NameInputDialog(
            self,
            title="체인 저장",
            message="체인 이름을 입력하세요:",
            initial_value="체인"
        )
        self.wait_window(name_dialog)

        if not name_dialog.result:
            return False

        name = name_dialog.result

        # 파일명 생성
        from datetime import datetime
        import os
        filename = f"{name.replace(' ', '_')}_chain_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        filepath = os.path.join('projects', filename)

        # 체인 데이터 생성
        chain_data = ChainManager.create_empty_chain(name, "")
        chain_data['chain_items'] = self.chain_items

        # 저장
        if ChainManager.save_chain(filepath, chain_data):
            messagebox.showinfo("완료", f"체인 '{name}'이(가) 저장되었습니다!", parent=self)
            return True
        else:
            messagebox.showerror("오류", "체인 저장에 실패했습니다.", parent=self)
            return False

    def on_cancel(self):
        """취소 버튼"""
        self.result = None
        self.destroy()
