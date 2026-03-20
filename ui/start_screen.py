"""
시작 화면 - 프로젝트 생성/불러오기
"""

import customtkinter as ctk
from tkinter import messagebox, filedialog
import os
import json
from datetime import datetime
from ui.theme import Colors, Fonts, Sizes, Styles


class StartScreen(ctk.CTkFrame):
    def __init__(self, parent, app):
        super().__init__(parent, fg_color=Colors.BG_PRIMARY)
        self.app = app
        self.parent = parent

        self.setup_ui()
        self.load_recent_projects()

    def setup_ui(self):
        """UI 구성"""
        # ─── Header ───
        header = ctk.CTkFrame(
            self,
            fg_color=Colors.BG_SECONDARY,
            height=Sizes.HEADER_HEIGHT,
            corner_radius=0,
            border_width=0,
        )
        header.pack(fill="x")
        header.pack_propagate(False)

        # Header bottom accent line (gradient-like effect)
        header_inner = ctk.CTkFrame(header, fg_color="transparent")
        header_inner.pack(fill="both", expand=True)

        title_label = ctk.CTkLabel(
            header_inner,
            text="MacroBuilder",
            font=Fonts.DISPLAY,
            text_color=Colors.PRIMARY_LIGHT,
        )
        title_label.pack(side="left", padx=Sizes.PAD_XL, pady=Sizes.PAD_MD)

        # Utility buttons (ghost style)
        util_frame = ctk.CTkFrame(header_inner, fg_color="transparent")
        util_frame.pack(side="right", padx=Sizes.PAD_XL)

        Styles.ghost_button(
            util_frame,
            text="단축키",
            command=self.show_hotkey_settings,
            width=80,
        ).pack(side="right", padx=(Sizes.PAD_XS, 0))

        Styles.ghost_button(
            util_frame,
            text="Auto",
            command=self.show_autostart_settings,
            width=72,
            text_color=Colors.SUCCESS,
        ).pack(side="right")

        # Header accent bar
        accent_bar = ctk.CTkFrame(
            self,
            fg_color=Colors.PRIMARY,
            height=2,
            corner_radius=0,
        )
        accent_bar.pack(fill="x")

        # ─── Bottom action bar (pack first to anchor at bottom) ───
        bottom_bar = ctk.CTkFrame(
            self,
            fg_color=Colors.BG_SECONDARY,
            corner_radius=0,
        )
        bottom_bar.pack(side="bottom", fill="x")

        bottom_inner = ctk.CTkFrame(bottom_bar, fg_color="transparent")
        bottom_inner.pack(fill="x", padx=Sizes.PAD_XL, pady=Sizes.PAD_LG)

        Styles.primary_button(
            bottom_inner,
            text="+ 새 프로젝트",
            command=self.create_new_project,
        ).pack(side="left", expand=True, fill="both", padx=(0, Sizes.PAD_SM))

        Styles.primary_button(
            bottom_inner,
            text="+ 체인 생성",
            command=self.start_chain_execution,
            fg_color=Colors.ACCENT,
            hover_color=Colors.ACCENT_HOVER,
        ).pack(side="left", expand=True, fill="both", padx=(Sizes.PAD_SM, 0))

        # Bottom bar top border
        bottom_border = ctk.CTkFrame(
            bottom_bar,
            fg_color=Colors.BORDER,
            height=1,
            corner_radius=0,
        )
        bottom_border.pack(side="top", fill="x")
        bottom_border.lift()

        # ─── Main content area ───
        content_frame = ctk.CTkFrame(self, fg_color="transparent")
        content_frame.pack(
            fill="both",
            expand=True,
            padx=Sizes.PAD_XL,
            pady=(Sizes.PAD_LG, Sizes.PAD_SM),
        )

        # ─── Search bar (full width, rounded) ───
        search_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
        search_frame.pack(fill="x", pady=(0, Sizes.PAD_LG))

        self.search_var = ctk.StringVar()
        self.search_var.trace_add("write", lambda *args: self.load_recent_projects())

        search_entry = Styles.input_field(
            search_frame,
            textvariable=self.search_var,
            placeholder_text="프로젝트 검색...",
            height=Sizes.BTN_HEIGHT_LG,
            corner_radius=Sizes.RADIUS_XL,
        )
        search_entry.pack(fill="x")

        # ─── Section header ───
        section_header = ctk.CTkFrame(content_frame, fg_color="transparent")
        section_header.pack(fill="x", pady=(0, Sizes.PAD_SM))

        Styles.section_title(
            section_header,
            text="프로젝트",
            font=Fonts.SMALL_BOLD,
        ).pack(side="left")

        # ─── Scrollable project list ───
        self.setup_scrollable_projects(content_frame)

    def setup_scrollable_projects(self, parent):
        """스크롤 가능한 프로젝트 리스트 생성"""
        self.recent_frame = ctk.CTkScrollableFrame(
            parent,
            fg_color="transparent",
            scrollbar_button_color=Colors.GRAY_700,
            scrollbar_button_hover_color=Colors.GRAY_600,
            corner_radius=0,
        )
        self.recent_frame.pack(fill="both", expand=True)

    def load_recent_projects(self):
        """최근 프로젝트 목록 로드"""
        # 기존 위젯 제거
        for widget in self.recent_frame.winfo_children():
            widget.destroy()

        # projects 폴더에서 .json 파일 찾기
        if not os.path.exists("projects"):
            self._show_empty_state()
            return

        json_files = [f for f in os.listdir("projects") if f.endswith(".json")]

        if not json_files:
            self._show_empty_state()
            return

        # 수정 시간 기준 정렬
        json_files.sort(
            key=lambda x: os.path.getmtime(os.path.join("projects", x)), reverse=True
        )

        # 검색 필터 적용
        search_term = ""
        if hasattr(self, "search_var"):
            search_term = self.search_var.get().strip().lower()

        # 전체 프로젝트 표시 (검색 필터 포함)
        for idx, filename in enumerate(json_files):
            if search_term:
                # 파일명 또는 프로젝트명에 검색어 포함 여부
                filepath = os.path.join("projects", filename)
                try:
                    with open(filepath, "r", encoding="utf-8") as f:
                        data = json.load(f)
                    project_name = data.get("name", "").lower()
                    if (
                        search_term not in project_name
                        and search_term not in filename.lower()
                    ):
                        continue
                except Exception:
                    continue
            self.create_project_card(filename, idx)

    def _show_empty_state(self):
        """빈 상태 표시"""
        empty_frame = ctk.CTkFrame(self.recent_frame, fg_color="transparent")
        empty_frame.pack(fill="x", pady=Sizes.PAD_2XL)

        ctk.CTkLabel(
            empty_frame,
            text="프로젝트가 없습니다",
            font=Fonts.BODY,
            text_color=Colors.TEXT_MUTED,
        ).pack()

        ctk.CTkLabel(
            empty_frame,
            text="아래 버튼으로 새 프로젝트를 만들어 보세요",
            font=Fonts.CAPTION,
            text_color=Colors.GRAY_600,
        ).pack(pady=(Sizes.PAD_XS, 0))

    def create_project_card(self, filename, index):
        """프로젝트 카드 생성"""
        filepath = os.path.join("projects", filename)

        # 프로젝트 정보 읽기
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                project_data = json.load(f)

            project_name = project_data.get("name", filename.replace(".json", ""))
            description = project_data.get("description", "설명 없음")
            modified_time = datetime.fromtimestamp(os.path.getmtime(filepath))

            # 타입 확인 (프로젝트인지 체인인지)
            item_type = project_data.get("type", "project")

        except Exception as e:
            print(f"프로젝트 로드 오류: {e}")
            return

        # 자동시작 프로젝트인지 확인
        is_autostart = False
        try:
            with open("settings.json", "r", encoding="utf-8") as f:
                settings = json.load(f)
                autostart_path = settings.get("autostart")
                if autostart_path and autostart_path == filepath:
                    is_autostart = True
        except (FileNotFoundError, json.JSONDecodeError, OSError):
            pass

        # ─── Card frame ───
        if is_autostart:
            card = Styles.card(
                self.recent_frame,
                fg_color=Colors.AUTOSTART_BG,
                border_color=Colors.AUTOSTART_BORDER,
                border_width=1,
            )
        else:
            card = Styles.card(self.recent_frame)

        card.pack(fill="x", padx=0, pady=(0, Sizes.PAD_SM))

        # Card inner padding
        card_inner = ctk.CTkFrame(card, fg_color="transparent")
        card_inner.pack(fill="both", expand=True, padx=Sizes.PAD_LG, pady=Sizes.PAD_MD)

        # ─── Left: info area ───
        info_frame = ctk.CTkFrame(card_inner, fg_color="transparent")
        info_frame.pack(side="left", fill="both", expand=True)

        # Type icon + name row
        name_row = ctk.CTkFrame(info_frame, fg_color="transparent")
        name_row.pack(anchor="w")

        icon = "C" if item_type == "chain" else "P"

        # Type icon badge
        badge_color = Colors.ACCENT if item_type == "chain" else Colors.PRIMARY
        icon_label = ctk.CTkLabel(
            name_row,
            text=icon,
            font=Fonts.CAPTION_BOLD,
            fg_color=badge_color,
            text_color=Colors.TEXT_INVERSE,
            width=22,
            height=22,
            corner_radius=4,
        )
        icon_label.pack(side="left", padx=(0, Sizes.PAD_SM))

        # Autostart indicator
        if is_autostart:
            auto_badge = ctk.CTkLabel(
                name_row,
                text="A",
                font=Fonts.CAPTION,
                text_color=Colors.AUTOSTART_FG,
                width=16,
            )
            auto_badge.pack(side="left", padx=(0, Sizes.PAD_XS))

        # Project name (bold)
        name_label = ctk.CTkLabel(
            name_row,
            text=project_name,
            font=Fonts.BODY_BOLD,
            text_color=Colors.AUTOSTART_FG if is_autostart else Colors.TEXT_PRIMARY,
        )
        name_label.pack(side="left")

        # Modified time (muted, below name)
        time_label = ctk.CTkLabel(
            info_frame,
            text=modified_time.strftime("%Y-%m-%d %H:%M"),
            font=Fonts.TINY,
            text_color=Colors.TEXT_MUTED,
        )
        time_label.pack(anchor="w", padx=(24 + Sizes.PAD_SM, 0), pady=(Sizes.PAD_XS, 0))

        # ─── Right: action buttons ───
        button_frame = ctk.CTkFrame(card_inner, fg_color="transparent")
        button_frame.pack(side="right", padx=(Sizes.PAD_SM, 0))

        # Open button (primary small)
        ctk.CTkButton(
            button_frame,
            text="열기",
            font=Fonts.CAPTION_BOLD,
            fg_color=Colors.PRIMARY,
            hover_color=Colors.PRIMARY_HOVER,
            text_color=Colors.TEXT_INVERSE,
            command=lambda: self.open_item(filepath, item_type),
            width=56,
            height=Sizes.BTN_HEIGHT_SM,
            corner_radius=Sizes.RADIUS_SM,
        ).pack(side="left", padx=(0, Sizes.PAD_XS))

        # Edit button (ghost) - only for projects
        if item_type == "project":
            Styles.ghost_button(
                button_frame,
                text="편집",
                command=lambda: self.edit_project(filepath),
                width=56,
                font=Fonts.CAPTION,
            ).pack(side="left", padx=(0, Sizes.PAD_XS))

        # Remove button (ghost with danger text)
        Styles.ghost_button(
            button_frame,
            text="제거",
            command=lambda: self.remove_project(filepath, project_name),
            width=56,
            font=Fonts.CAPTION,
            text_color=Colors.DANGER,
            hover_color="#fef2f2",
        ).pack(side="left")

    def remove_project(self, filepath, project_name):
        """프로젝트 제거"""
        result = messagebox.askyesnocancel(
            "프로젝트 제거",
            f"'{project_name}' 프로젝트를 어떻게 처리하시겠습니까?\n\n"
            f"예: 파일 완전 삭제\n"
            f"아니오: 목록에서만 제거 (파일 유지)\n"
            f"취소: 작업 취소",
        )

        if result is None:
            return
        elif result:
            try:
                os.remove(filepath)
                messagebox.showinfo(
                    "완료", f"'{project_name}' 프로젝트가 삭제되었습니다."
                )
            except Exception as e:
                messagebox.showerror(
                    "오류", f"프로젝트 삭제 중 오류가 발생했습니다:\n{str(e)}"
                )
                return
        else:
            try:
                hidden_path = filepath + ".hidden"
                os.rename(filepath, hidden_path)
                messagebox.showinfo(
                    "완료",
                    f"'{project_name}' 프로젝트가 목록에서 제거되었습니다.\n(파일은 보존됨)",
                )
            except Exception as e:
                messagebox.showerror(
                    "오류", f"프로젝트 제거 중 오류가 발생했습니다:\n{str(e)}"
                )
                return

        self.load_recent_projects()

    def open_item(self, filepath, item_type):
        """프로젝트 또는 체인 열기"""
        if item_type == "chain":
            # 체인 로드 및 체인 러너 실행
            from core.chain_manager import ChainManager
            from ui.chain_runner import ChainRunner

            chain_data = ChainManager.load_chain(filepath)
            if not chain_data:
                messagebox.showerror("오류", "체인을 불러올 수 없습니다.")
                return

            chain_items = chain_data.get("chain_items", [])

            for widget in self.parent.winfo_children():
                widget.destroy()

            runner = ChainRunner(self.parent, self.app, chain_items)
            runner.pack(fill="both", expand=True)
        else:
            # 프로젝트 열기
            self.open_project(filepath)

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
        runner.pack(fill="both", expand=True)

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
        editor.pack(fill="both", expand=True)

    def create_new_project(self):
        """새 프로젝트 만들기"""
        from ui.dialogs import NewProjectDialog
        from core.project_manager import ProjectManager

        dialog = NewProjectDialog(self.parent)
        self.parent.wait_window(dialog)

        if dialog.result:
            name = dialog.result["name"]
            description = dialog.result["description"]

            filename = f"{name.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            filepath = os.path.join("projects", filename)

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
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )

        if filepath:
            self.open_project(filepath)

    def start_chain_execution(self):
        """매크로 체인 실행"""
        from ui.chain_dialog import MacroChainDialog
        from ui.chain_runner import ChainRunner
        from core.project_manager import ProjectManager

        # 체인 설정 다이얼로그
        dialog = MacroChainDialog(self.parent, ProjectManager)
        self.parent.wait_window(dialog)

        if dialog.result:
            # 체인 러너 화면으로 전환
            for widget in self.parent.winfo_children():
                widget.destroy()

            runner = ChainRunner(self.parent, self.app, dialog.result)
            runner.pack(fill="both", expand=True)

    def show_hotkey_settings(self):
        """전역 단축키 설정"""
        from ui.hotkey_settings_dialog import HotkeySettingsDialog

        HotkeySettingsDialog(self.parent)

    def show_autostart_settings(self):
        """자동시작 프로젝트 설정"""
        from ui.dialogs import AutostartSelectDialog
        import json

        dialog = AutostartSelectDialog(self.parent)
        self.parent.wait_window(dialog)

        if dialog.result is not None:
            # settings.json 읽기
            try:
                with open("settings.json", "r", encoding="utf-8") as f:
                    settings = json.load(f)
            except (FileNotFoundError, json.JSONDecodeError, OSError):
                settings = {"hotkeys": {}}

            # 자동시작 설정 업데이트
            if dialog.result["filepath"] is None:
                # 자동시작 해제
                settings["autostart"] = None
                messagebox.showinfo("완료", "자동시작이 해제되었습니다.")
            else:
                # 자동시작 설정
                settings["autostart"] = dialog.result["filepath"]
                messagebox.showinfo(
                    "완료",
                    f"'{dialog.result['name']}'이(가) 자동시작으로 설정되었습니다.",
                )

            # settings.json 저장
            try:
                with open("settings.json", "w", encoding="utf-8") as f:
                    json.dump(settings, f, ensure_ascii=False, indent=2)

                # 프로젝트 목록 새로고침 (자동시작 표시 업데이트)
                self.load_recent_projects()
            except Exception as e:
                messagebox.showerror(
                    "오류", f"설정 저장 중 오류가 발생했습니다:\n{str(e)}"
                )
