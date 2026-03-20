"""
매크로 체인 설정 다이얼로그
"""
import tkinter as tk
import customtkinter as ctk
from tkinter import messagebox
from utils.ui_helpers import set_dialog_icon, center_window_on_parent
from ui.theme import Colors, Fonts, Sizes, Styles


class MacroChainDialog(ctk.CTkToplevel):
    """매크로 체인 설정 다이얼로그"""

    def __init__(self, parent, project_manager):
        super().__init__(parent)
        self.title("매크로 체인 실행 설정")
        self.geometry("520x640")
        self.resizable(False, False)

        self.project_manager = project_manager
        self.result = None
        self.chain_items = []  # [{'project_name': str, 'filepath': str, 'repeat_count': int}]

        # 모달 설정
        self.withdraw()
        self.transient(parent)
        self.grab_set()
        self.attributes('-topmost', True)
        set_dialog_icon(self)

        self.setup_ui()

        def _show():
            center_window_on_parent(self, parent)
            self.deiconify()
            self.lift()
            self.focus_force()
        self.after(50, _show)

    def setup_ui(self):
        """UI 구성"""
        main_frame = ctk.CTkFrame(self, fg_color=Colors.BG_SECONDARY)
        main_frame.pack(fill='both', expand=True)

        # ── 헤더 카드 ──
        header_card = Styles.card(main_frame, corner_radius=0, border_width=0)
        header_card.pack(fill='x')

        header_inner = ctk.CTkFrame(header_card, fg_color="transparent")
        header_inner.pack(fill='x', padx=Sizes.PAD_XL, pady=Sizes.PAD_LG)

        ctk.CTkLabel(
            header_inner,
            text="매크로 체인 실행",
            font=Fonts.HEADING,
            text_color=Colors.TEXT_PRIMARY,
        ).pack(anchor='w')

        ctk.CTkLabel(
            header_inner,
            text="여러 매크로를 순차적으로 실행합니다.\n각 매크로의 반복 횟수를 설정할 수 있습니다.",
            font=Fonts.CAPTION,
            text_color=Colors.TEXT_MUTED,
            justify='left',
        ).pack(anchor='w', pady=(Sizes.PAD_XS, 0))

        # ── 체인 리스트 영역 ──
        body = ctk.CTkFrame(main_frame, fg_color="transparent")
        body.pack(fill='both', expand=True, padx=Sizes.PAD_LG, pady=Sizes.PAD_MD)

        Styles.section_title(body, text="체인 순서").pack(anchor='w', pady=(0, Sizes.PAD_XS))

        # 스크롤 가능한 리스트
        self.chain_scroll_frame = ctk.CTkScrollableFrame(
            body,
            fg_color=Colors.BG_CARD,
            corner_radius=Sizes.RADIUS_MD,
            border_width=1,
            border_color=Colors.BORDER,
        )
        self.chain_scroll_frame.pack(fill='both', expand=True)

        # chain_list_frame은 scroll_frame 내부
        self.chain_list_frame = self.chain_scroll_frame

        # ── 하단 버튼 ──
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill='x', padx=Sizes.PAD_LG, pady=(0, Sizes.PAD_LG))

        Styles.primary_button(
            btn_frame,
            text="+ 매크로 추가",
            command=self.add_macro_to_chain,
            height=Sizes.BTN_HEIGHT,
        ).pack(side='left', padx=(0, Sizes.PAD_XS))

        Styles.secondary_button(
            btn_frame,
            text="취소",
            command=self.on_cancel,
        ).pack(side='right', padx=(Sizes.PAD_XS, 0))

        ctk.CTkButton(
            btn_frame,
            text="실행",
            font=Fonts.BODY_BOLD,
            fg_color=Colors.SUCCESS,
            hover_color=Colors.SUCCESS_HOVER,
            text_color=Colors.TEXT_INVERSE,
            height=Sizes.BTN_HEIGHT,
            corner_radius=Sizes.RADIUS_SM,
            command=self.on_execute,
        ).pack(side='right', padx=(Sizes.PAD_XS, 0))

        self.refresh_chain_list()

    def add_macro_to_chain(self):
        """체인에 매크로 추가"""
        # 프로젝트 선택 다이얼로그
        projects = self.project_manager.get_project_list()

        if not projects:
            messagebox.showwarning("경고", "등록된 프로젝트가 없습니다.", parent=self)
            return

        # 프로젝트 선택 다이얼로그
        dialog = ctk.CTkToplevel(self)
        dialog.title("매크로 선택 및 실행 설정")
        dialog.geometry("540x640")
        dialog.transient(self)
        dialog.grab_set()
        dialog.attributes('-topmost', True)
        set_dialog_icon(dialog)

        result = [None]

        dialog_main = ctk.CTkFrame(dialog, fg_color=Colors.BG_SECONDARY)
        dialog_main.pack(fill='both', expand=True)

        # ── 헤더 ──
        dlg_header = Styles.card(dialog_main, corner_radius=0, border_width=0)
        dlg_header.pack(fill='x')

        ctk.CTkLabel(
            dlg_header,
            text="실행할 매크로를 선택하세요",
            font=Fonts.SUBHEADING,
            text_color=Colors.TEXT_PRIMARY,
        ).pack(padx=Sizes.PAD_XL, pady=Sizes.PAD_MD)

        # ── 프로젝트 리스트 ──
        body = ctk.CTkFrame(dialog_main, fg_color="transparent")
        body.pack(fill='both', expand=True, padx=Sizes.PAD_LG, pady=Sizes.PAD_SM)

        Styles.section_title(body, text="프로젝트 목록").pack(anchor='w', pady=(0, Sizes.PAD_XS))

        listbox_card = Styles.card(body)
        listbox_card.pack(fill='both', expand=True)

        listbox_inner = ctk.CTkFrame(listbox_card, fg_color="transparent")
        listbox_inner.pack(fill='both', expand=True, padx=2, pady=2)

        scrollbar = tk.Scrollbar(listbox_inner)
        scrollbar.pack(side='right', fill='y')

        project_listbox = tk.Listbox(
            listbox_inner,
            font=Fonts.BODY,
            yscrollcommand=scrollbar.set,
            selectmode='single',
            bg=Colors.BG_INPUT,
            fg=Colors.TEXT_PRIMARY,
            selectbackground=Colors.PRIMARY,
            selectforeground=Colors.TEXT_INVERSE,
            borderwidth=0,
            highlightthickness=0,
            relief='flat',
        )
        project_listbox.pack(side='left', fill='both', expand=True)
        scrollbar.config(command=project_listbox.yview)

        # 모든 프로젝트 표시 (체인과 매크로 구분)
        for proj in projects:
            proj_type = proj.get('type', 'project')
            if proj_type == 'chain':
                display_name = f"[체인] {proj['name']}"
            else:
                display_name = f"[매크로] {proj['name']}"
            project_listbox.insert(tk.END, display_name)

        # ── 실행 설정 영역 ──
        settings_section = ctk.CTkFrame(dialog_main, fg_color="transparent")
        settings_section.pack(fill='x', padx=Sizes.PAD_LG, pady=(Sizes.PAD_SM, 0))

        Styles.section_title(settings_section, text="실행 설정 (수정 가능)").pack(
            anchor='w', pady=(0, Sizes.PAD_XS)
        )

        settings_card = Styles.card(settings_section)
        settings_card.pack(fill='x')

        # 설정 정보를 담을 변수들
        repeat_var = tk.StringVar(value='1')
        infinite_var = tk.BooleanVar(value=False)

        # 설정 UI 컨테이너
        settings_container = ctk.CTkFrame(settings_card, fg_color="transparent")
        settings_container.pack(fill='x', padx=Sizes.PAD_MD, pady=Sizes.PAD_MD)

        # 반복 횟수 입력
        repeat_frame = ctk.CTkFrame(settings_container, fg_color="transparent")
        repeat_frame.pack(fill='x', pady=(0, Sizes.PAD_SM))

        repeat_title_label = ctk.CTkLabel(
            repeat_frame,
            text="반복 횟수:",
            font=Fonts.SMALL_BOLD,
            text_color=Colors.TEXT_SECONDARY,
            width=80,
            anchor='w',
        )
        repeat_title_label.pack(side='left')

        repeat_entry = Styles.input_field(
            repeat_frame,
            textvariable=repeat_var,
            width=100,
        )
        repeat_entry.pack(side='left', padx=Sizes.PAD_SM)

        repeat_label = ctk.CTkLabel(
            repeat_frame,
            text="회",
            font=Fonts.SMALL,
            text_color=Colors.TEXT_SECONDARY,
        )
        repeat_label.pack(side='left')

        # 무한 반복 체크박스 및 토글 함수
        def toggle_repeat_entry():
            """무한 반복 체크 시 반복 횟수 입력 비활성화"""
            if infinite_var.get():
                repeat_title_label.configure(text_color=Colors.TEXT_MUTED)
                repeat_entry.configure(state='disabled')
                repeat_label.configure(text_color=Colors.TEXT_MUTED)
            else:
                repeat_title_label.configure(text_color=Colors.TEXT_PRIMARY)
                repeat_entry.configure(state='normal')
                repeat_label.configure(text_color=Colors.TEXT_PRIMARY)

        infinite_check = ctk.CTkCheckBox(
            settings_container,
            text="무한 반복",
            variable=infinite_var,
            font=Fonts.SMALL_BOLD,
            text_color=Colors.DANGER,
            command=toggle_repeat_entry,
        )
        infinite_check.pack(anchor='w', pady=(0, Sizes.PAD_SM))

        ctk.CTkLabel(
            settings_container,
            text="※ 엑셀 데이터가 있으면 자동으로 엑셀 행별 반복이 적용됩니다",
            font=Fonts.CAPTION,
            text_color=Colors.TEXT_MUTED,
        ).pack(anchor='w', pady=(Sizes.PAD_XS, 0))

        ctk.CTkLabel(
            settings_container,
            text="※ 수정하면 프로젝트 파일에 저장됩니다",
            font=Fonts.TINY,
            text_color=Colors.TEXT_MUTED,
        ).pack(anchor='w', pady=(2, 0))

        # 프로젝트 선택 시 설정 로드
        def on_project_select(event):
            selection = project_listbox.curselection()
            if selection:
                selected_project = projects[selection[0]]
                import json

                # 프로젝트 파일 로드
                try:
                    with open(selected_project['filepath'], 'r', encoding='utf-8') as f:
                        project_data = json.load(f)

                    # 설정값 로드
                    settings = project_data.get('settings', {})
                    execution = settings.get('execution', {})

                    mode = execution.get('mode', 'flow_repeat')
                    repeat_count = execution.get('repeat_count', 1)

                    # 무한 반복 여부 판단
                    is_infinite = (mode == 'infinite' or execution.get('excel_infinite_loop', False))

                    infinite_var.set(is_infinite)
                    repeat_var.set(str(repeat_count))

                    # UI 상태 업데이트
                    toggle_repeat_entry()

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

            selected_project = projects[selection[0]]

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

                # 반복 횟수 저장
                project_data['settings']['execution']['repeat_count'] = repeat_count

                # 엑셀 데이터 확인
                has_excel = bool(project_data.get('excel_sources', []))

                # 실행 모드 결정
                if infinite_var.get():
                    # 무한 반복
                    if has_excel:
                        # 엑셀이 있으면 엑셀 무한 반복
                        project_data['settings']['execution']['mode'] = 'excel_loop'
                        project_data['settings']['execution']['excel_infinite_loop'] = True
                    else:
                        # 엑셀이 없으면 플로우 무한 반복
                        project_data['settings']['execution']['mode'] = 'infinite'
                        project_data['settings']['execution']['excel_infinite_loop'] = False
                else:
                    # 횟수 지정 반복
                    if has_excel:
                        # 엑셀이 있으면 엑셀 행 반복
                        project_data['settings']['execution']['mode'] = 'excel_loop'
                        project_data['settings']['execution']['excel_infinite_loop'] = False
                    else:
                        # 엑셀이 없으면 플로우 반복
                        project_data['settings']['execution']['mode'] = 'flow_repeat'
                        project_data['settings']['execution']['excel_infinite_loop'] = False

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

        # ── 하단 버튼 ──
        btn_frame = ctk.CTkFrame(dialog_main, fg_color="transparent")
        btn_frame.pack(fill='x', padx=Sizes.PAD_LG, pady=Sizes.PAD_MD)

        ctk.CTkButton(
            btn_frame,
            text="추가",
            font=Fonts.BODY_BOLD,
            fg_color=Colors.SUCCESS,
            hover_color=Colors.SUCCESS_HOVER,
            text_color=Colors.TEXT_INVERSE,
            height=Sizes.BTN_HEIGHT,
            corner_radius=Sizes.RADIUS_SM,
            command=on_ok,
        ).pack(side='right', padx=(Sizes.PAD_XS, 0))

        Styles.secondary_button(
            btn_frame,
            text="취소",
            command=dialog.destroy,
        ).pack(side='right', padx=(Sizes.PAD_XS, 0))

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
            ctk.CTkLabel(
                self.chain_list_frame,
                text="매크로를 추가하세요",
                font=Fonts.SMALL,
                text_color=Colors.TEXT_MUTED,
            ).pack(pady=50)
            return

        for idx, item in enumerate(self.chain_items):
            self.create_chain_item_widget(idx, item)

    def get_execution_mode_text(self, item):
        """프로젝트의 실행 설정 텍스트 가져오기"""
        execution_mode = item.get('execution_mode', 'custom_repeat')

        if execution_mode == 'custom_repeat':
            # 커스텀 반복 횟수 사용
            repeat_count = item.get('repeat_count', 1)
            return f"반복: {repeat_count}회"
        else:
            # 프로젝트 설정 사용 - 파일에서 읽어오기
            try:
                import json
                filepath = item.get('filepath', '')
                if filepath:
                    with open(filepath, 'r', encoding='utf-8') as f:
                        project_data = json.load(f)

                    settings = project_data.get('settings', {}).get('execution', {})
                    mode = settings.get('mode', 'flow_repeat')
                    repeat_count = settings.get('repeat_count', 1)
                    excel_infinite = settings.get('excel_infinite_loop', False)

                    if mode == 'infinite':
                        return "무한 반복"
                    elif mode == 'excel_loop':
                        if excel_infinite:
                            return f"엑셀 행 반복 (무한)"
                        else:
                            return f"엑셀 행 반복 ({repeat_count}회)"
                    else:  # flow_repeat
                        return f"{repeat_count}회 반복"
            except Exception as e:
                print(f"설정 읽기 오류: {e}")

            return "프로젝트 설정 사용"

    def create_chain_item_widget(self, idx, item):
        """체인 아이템 위젯 생성"""
        item_frame = ctk.CTkFrame(
            self.chain_list_frame,
            fg_color=Colors.BG_INPUT,
            corner_radius=Sizes.RADIUS_SM,
            border_width=1,
            border_color=Colors.BORDER,
        )
        item_frame.pack(fill='x', padx=Sizes.PAD_XS, pady=Sizes.PAD_XS)

        # 순서 번호 배지
        badge = ctk.CTkLabel(
            item_frame,
            text=f" {idx + 1} ",
            font=Fonts.CAPTION_BOLD,
            text_color=Colors.TEXT_INVERSE,
            fg_color=Colors.PRIMARY,
            corner_radius=Sizes.RADIUS_SM,
            width=28,
            height=28,
        )
        badge.pack(side='left', padx=(Sizes.PAD_SM, Sizes.PAD_XS), pady=Sizes.PAD_SM)

        # 정보
        info_frame = ctk.CTkFrame(item_frame, fg_color="transparent")
        info_frame.pack(side='left', fill='x', expand=True, padx=Sizes.PAD_SM, pady=Sizes.PAD_SM)

        ctk.CTkLabel(
            info_frame,
            text=item['project_name'],
            font=Fonts.SMALL_BOLD,
            text_color=Colors.TEXT_PRIMARY,
            anchor='w',
        ).pack(anchor='w')

        # 실행 설정 표시 (실제 프로젝트 설정 읽기)
        mode_text = self.get_execution_mode_text(item)

        ctk.CTkLabel(
            info_frame,
            text=mode_text,
            font=Fonts.CAPTION,
            text_color=Colors.PRIMARY_LIGHT,
            anchor='w',
        ).pack(anchor='w')

        # 버튼들
        btn_frame = ctk.CTkFrame(item_frame, fg_color="transparent")
        btn_frame.pack(side='right', padx=Sizes.PAD_SM)

        # 위로 이동
        if idx > 0:
            Styles.ghost_button(
                btn_frame,
                text="▲",
                command=lambda: self.move_item_up(idx),
            ).configure(width=30, height=Sizes.BTN_HEIGHT_XS, font=Fonts.TINY)
            btn_frame.winfo_children()[-1].pack(side='left', padx=2)

        # 아래로 이동
        if idx < len(self.chain_items) - 1:
            Styles.ghost_button(
                btn_frame,
                text="▼",
                command=lambda: self.move_item_down(idx),
            ).configure(width=30, height=Sizes.BTN_HEIGHT_XS, font=Fonts.TINY)
            btn_frame.winfo_children()[-1].pack(side='left', padx=2)

        # 삭제
        Styles.danger_button(
            btn_frame,
            text="X",
            command=lambda: self.remove_item(idx),
        ).configure(width=30, height=Sizes.BTN_HEIGHT_XS, font=Fonts.TINY)
        btn_frame.winfo_children()[-1].pack(side='left', padx=2)

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

        self.save_chain()
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
