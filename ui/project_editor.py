"""
프로젝트 편집 화면 - 좌표/엑셀/이미지/플로우 관리
"""

import tkinter as tk
import customtkinter as ctk
from tkinter import messagebox, filedialog
from datetime import datetime
import os
from utils.ui_helpers import set_dialog_icon, center_window_on_parent
from ui.theme import Colors, Fonts, Sizes, Styles

from core.project_manager import ProjectManager
from core.coordinate_manager import CoordinateManager
from core.excel_manager import ExcelManager
from core.image_manager import ImageManager
from core.flow_manager import FlowManager
from ui.dialogs import ActionSelectDialog, NameInputDialog


class ProjectEditor(ctk.CTkFrame):
    def __init__(self, parent, app, project_data, filepath):
        super().__init__(parent, fg_color=Colors.BG_PRIMARY)
        self.app = app
        self.parent = parent
        self.project_data = project_data
        self.filepath = filepath

        # 관리자 초기화
        self.coord_mgr = CoordinateManager()
        self.coord_mgr.load_from_list(project_data.get("coordinates", []))

        self.excel_mgr = ExcelManager()
        self.excel_mgr.load_from_list(project_data.get("excel_sources", []))

        self.image_mgr = ImageManager()
        self.image_mgr.load_from_list(project_data.get("images", []))

        self.flow_mgr = FlowManager()
        self.flow_mgr.load_from_list(project_data.get("flow_sequence", []))

        # 드래그 앤 드롭 관련 변수
        self.drag_data = {"item": None, "index": None, "y": 0}
        self.drop_indicator = None  # 드롭 위치 표시 라인

        self.setup_ui()

    def setup_ui(self):
        """UI 구성"""
        # 상단 헤더
        header = ctk.CTkFrame(
            self,
            fg_color=Colors.BG_CARD,
            height=Sizes.HEADER_HEIGHT,
            corner_radius=0,
            border_width=0,
        )
        header.pack(fill="x")
        header.pack_propagate(False)

        # 헤더 좌측: 프로젝트 이름 (prominent)
        name_label = ctk.CTkLabel(
            header,
            text=self.project_data["name"],
            font=Fonts.HEADING,
            text_color=Colors.TEXT_PRIMARY,
        )
        name_label.pack(side="left", padx=Sizes.PAD_XL, pady=Sizes.PAD_MD)

        # 헤더 우측: 완료 버튼 (primary style)
        btn_frame = ctk.CTkFrame(header, fg_color="transparent")
        btn_frame.pack(side="right", padx=Sizes.PAD_XL)

        Styles.primary_button(
            btn_frame,
            text="완료",
            command=self.finish_editing,
            width=100,
            height=Sizes.BTN_HEIGHT,
        ).pack(side="left", padx=Sizes.PAD_XS)

        # 헤더 하단 구분선
        separator = ctk.CTkFrame(
            self, fg_color=Colors.BORDER, height=1, corner_radius=0
        )
        separator.pack(fill="x")

        # 메인 컨텐츠 (좌우 분할)
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True)

        # 좌측: 리소스 관리 (SIDEBAR_WIDTH)
        left_frame = ctk.CTkFrame(
            main_frame,
            width=Sizes.SIDEBAR_WIDTH,
            fg_color=Colors.BG_SECONDARY,
            corner_radius=0,
            border_width=0,
        )
        left_frame.pack(side="left", fill="y")
        left_frame.pack_propagate(False)

        # 좌측 패널과 우측 패널 사이 구분선
        v_separator = ctk.CTkFrame(
            main_frame, fg_color=Colors.BORDER, width=1, corner_radius=0
        )
        v_separator.pack(side="left", fill="y")

        self.setup_resource_panel(left_frame)

        # 우측: 플로우 에디터
        right_frame = ctk.CTkFrame(
            main_frame, fg_color=Colors.BG_PRIMARY, corner_radius=0
        )
        right_frame.pack(side="left", fill="both", expand=True)

        self.setup_flow_panel(right_frame)

    def setup_resource_panel(self, parent):
        """리소스 패널 구성"""
        # CTkScrollableFrame 사용
        scrollable_frame = ctk.CTkScrollableFrame(
            parent,
            fg_color=Colors.BG_SECONDARY,
            corner_radius=0,
            scrollbar_button_color=Colors.GRAY_600,
            scrollbar_button_hover_color=Colors.GRAY_500,
        )
        scrollable_frame.pack(fill="both", expand=True, padx=0, pady=0)

        # 좌표 섹션
        self.setup_coordinate_section(scrollable_frame)

        # 엑셀 섹션
        self.setup_excel_section(scrollable_frame)

        # 이미지 섹션
        self.setup_image_section(scrollable_frame)

    def setup_coordinate_section(self, parent):
        """좌표 섹션"""
        # 섹션 타이틀
        Styles.section_title(parent, text="좌표").pack(
            fill="x", padx=Sizes.PAD_MD, pady=(Sizes.PAD_MD, Sizes.PAD_XS)
        )

        section = ctk.CTkFrame(
            parent, fg_color="transparent", corner_radius=0
        )
        section.pack(fill="x", padx=Sizes.PAD_SM, pady=(0, Sizes.PAD_SM))

        # 좌표 리스트
        self.coord_list_frame = ctk.CTkFrame(section, fg_color="transparent")
        self.coord_list_frame.pack(fill="x")

        # 추가 버튼 (PRIMARY 색상)
        ctk.CTkButton(
            section,
            text="+ 새 좌표 추가",
            font=Fonts.CAPTION_BOLD,
            fg_color=Colors.PRIMARY,
            hover_color=Colors.PRIMARY_HOVER,
            text_color=Colors.TEXT_INVERSE,
            height=Sizes.BTN_HEIGHT_SM,
            corner_radius=Sizes.RADIUS_SM,
            command=self.add_coordinate_dialog,
        ).pack(fill="x", padx=Sizes.PAD_XS, pady=(Sizes.PAD_SM, Sizes.PAD_XS))

        self.refresh_coordinate_list()

    def setup_excel_section(self, parent):
        """엑셀 섹션"""
        # 구분선
        ctk.CTkFrame(
            parent, fg_color=Colors.BORDER, height=1, corner_radius=0
        ).pack(fill="x", padx=Sizes.PAD_MD, pady=(Sizes.PAD_XS, 0))

        # 섹션 타이틀
        Styles.section_title(parent, text="엑셀").pack(
            fill="x", padx=Sizes.PAD_MD, pady=(Sizes.PAD_MD, Sizes.PAD_XS)
        )

        section = ctk.CTkFrame(
            parent, fg_color="transparent", corner_radius=0
        )
        section.pack(fill="x", padx=Sizes.PAD_SM, pady=(0, Sizes.PAD_SM))

        # 엑셀 리스트
        self.excel_list_frame = ctk.CTkFrame(section, fg_color="transparent")
        self.excel_list_frame.pack(fill="x")

        # 추가 버튼 (SUCCESS 색상)
        ctk.CTkButton(
            section,
            text="+ 새 엑셀 추가",
            font=Fonts.CAPTION_BOLD,
            fg_color=Colors.SUCCESS,
            hover_color=Colors.SUCCESS_HOVER,
            text_color=Colors.TEXT_INVERSE,
            height=Sizes.BTN_HEIGHT_SM,
            corner_radius=Sizes.RADIUS_SM,
            command=self.add_excel_dialog,
        ).pack(fill="x", padx=Sizes.PAD_XS, pady=(Sizes.PAD_SM, Sizes.PAD_XS))

        self.refresh_excel_list()

    def setup_image_section(self, parent):
        """이미지 섹션"""
        # 구분선
        ctk.CTkFrame(
            parent, fg_color=Colors.BORDER, height=1, corner_radius=0
        ).pack(fill="x", padx=Sizes.PAD_MD, pady=(Sizes.PAD_XS, 0))

        # 섹션 타이틀
        Styles.section_title(parent, text="이미지").pack(
            fill="x", padx=Sizes.PAD_MD, pady=(Sizes.PAD_MD, Sizes.PAD_XS)
        )

        section = ctk.CTkFrame(
            parent, fg_color="transparent", corner_radius=0
        )
        section.pack(fill="x", padx=Sizes.PAD_SM, pady=(0, Sizes.PAD_SM))

        # 이미지 리스트
        self.image_list_frame = ctk.CTkFrame(section, fg_color="transparent")
        self.image_list_frame.pack(fill="x")

        # 추가 버튼 (PURPLE 색상)
        ctk.CTkButton(
            section,
            text="+ 새 이미지 추가",
            font=Fonts.CAPTION_BOLD,
            fg_color=Colors.PURPLE,
            hover_color=Colors.PURPLE_HOVER,
            text_color=Colors.TEXT_INVERSE,
            height=Sizes.BTN_HEIGHT_SM,
            corner_radius=Sizes.RADIUS_SM,
            command=self.add_image_from_file,
        ).pack(fill="x", padx=Sizes.PAD_XS, pady=(Sizes.PAD_SM, Sizes.PAD_XS))

        self.refresh_image_list()

    def setup_flow_panel(self, parent):
        """플로우 패널 구성"""
        # 제목
        title_frame = ctk.CTkFrame(parent, fg_color="transparent")
        title_frame.pack(fill="x", padx=Sizes.PAD_LG, pady=(Sizes.PAD_MD, Sizes.PAD_SM))

        ctk.CTkLabel(
            title_frame,
            text="플로우 시퀀스",
            font=Fonts.SUBHEADING,
            text_color=Colors.TEXT_PRIMARY,
        ).pack(side="left")

        # 플로우 리스트 (CTkScrollableFrame)
        self.flow_scrollable = ctk.CTkScrollableFrame(
            parent,
            fg_color=Colors.BG_PRIMARY,
            corner_radius=0,
            scrollbar_button_color=Colors.GRAY_600,
            scrollbar_button_hover_color=Colors.GRAY_500,
        )
        self.flow_scrollable.pack(
            fill="both", expand=True, padx=Sizes.PAD_SM, pady=(0, Sizes.PAD_SM)
        )

        self.flow_list_frame = self.flow_scrollable

        # 액션 추가 버튼 - Large primary style, centered
        btn_frame = ctk.CTkFrame(parent, fg_color="transparent")
        btn_frame.pack(fill="x", padx=Sizes.PAD_XL, pady=Sizes.PAD_LG)

        Styles.primary_button(
            btn_frame,
            text="  액션 추가",
            command=self.add_action_menu,
            width=200,
            height=Sizes.BTN_HEIGHT_LG,
            corner_radius=Sizes.RADIUS_MD,
        ).pack()

        self.refresh_flow_list()

    def refresh_coordinate_list(self):
        """좌표 목록 새로고침"""
        for widget in self.coord_list_frame.winfo_children():
            widget.destroy()

        if not self.coord_mgr.coordinates:
            ctk.CTkLabel(
                self.coord_list_frame,
                text="좌표가 없습니다",
                font=Fonts.CAPTION,
                text_color=Colors.TEXT_MUTED,
            ).pack(fill="x", pady=Sizes.PAD_SM)
            return

        for coord in self.coord_mgr.coordinates:
            self.create_coordinate_item(coord)

    def create_coordinate_item(self, coord):
        """좌표 아이템 생성 - compact card"""
        item = ctk.CTkFrame(
            self.coord_list_frame,
            fg_color=Colors.BG_INPUT,
            corner_radius=Sizes.RADIUS_SM,
            border_width=0,
        )
        item.pack(fill="x", pady=2, padx=Sizes.PAD_XS)

        info_frame = ctk.CTkFrame(item, fg_color="transparent")
        info_frame.pack(
            side="left", fill="both", expand=True, padx=Sizes.PAD_SM, pady=Sizes.PAD_XS
        )

        # 제목 길이 제한 (최대 10자)
        MAX_TITLE_LENGTH = 10
        display_name = (
            coord["name"]
            if len(coord["name"]) <= MAX_TITLE_LENGTH
            else coord["name"][:MAX_TITLE_LENGTH] + "..."
        )

        ctk.CTkLabel(
            info_frame,
            text=f"{coord['id']}. {display_name}",
            font=Fonts.CAPTION_BOLD,
            text_color=Colors.TEXT_PRIMARY,
            anchor="w",
        ).pack(anchor="w", fill="x")

        ctk.CTkLabel(
            info_frame,
            text=f"({coord['x']}, {coord['y']})",
            font=Fonts.TINY,
            text_color=Colors.TEXT_MUTED,
            anchor="w",
        ).pack(anchor="w")

        # 삭제 버튼 - ghost style, danger color, icon "✕"
        Styles.ghost_button(
            item,
            text="\u2715",
            command=lambda: self.delete_coordinate(coord["id"]),
            width=24,
            height=22,
            text_color=Colors.DANGER,
            hover_color=Colors.BG_CARD,
            font=Fonts.TINY,
        ).pack(side="right", padx=(0, Sizes.PAD_XS), pady=Sizes.PAD_XS)

    def refresh_excel_list(self):
        """엑셀 목록 새로고침"""
        for widget in self.excel_list_frame.winfo_children():
            widget.destroy()

        if not self.excel_mgr.excel_sources:
            ctk.CTkLabel(
                self.excel_list_frame,
                text="엑셀 데이터가 없습니다",
                font=Fonts.CAPTION,
                text_color=Colors.TEXT_MUTED,
            ).pack(fill="x", pady=Sizes.PAD_SM)
            return

        for source in self.excel_mgr.excel_sources:
            self.create_excel_item(source)

    def create_excel_item(self, source):
        """엑셀 아이템 생성 - compact card"""
        item = ctk.CTkFrame(
            self.excel_list_frame,
            fg_color=Colors.BG_INPUT,
            corner_radius=Sizes.RADIUS_SM,
            border_width=0,
        )
        item.pack(fill="x", pady=2, padx=Sizes.PAD_XS)

        info_frame = ctk.CTkFrame(item, fg_color="transparent")
        info_frame.pack(
            side="left", fill="both", expand=True, padx=Sizes.PAD_SM, pady=Sizes.PAD_XS
        )

        # 이름만 표시 (제목 길이 제한)
        MAX_TITLE_LENGTH = 12
        display_name = (
            source["name"]
            if len(source["name"]) <= MAX_TITLE_LENGTH
            else source["name"][:MAX_TITLE_LENGTH] + "..."
        )

        ctk.CTkLabel(
            info_frame,
            text=display_name,
            font=Fonts.CAPTION_BOLD,
            text_color=Colors.TEXT_PRIMARY,
            anchor="w",
        ).pack(anchor="w", fill="x")

        ctk.CTkLabel(
            info_frame,
            text=f"{source['row_count']} rows, {len(source['columns'])} cols",
            font=Fonts.TINY,
            text_color=Colors.TEXT_MUTED,
            anchor="w",
        ).pack(anchor="w")

        # 삭제 버튼 - ghost style, danger color, icon "✕"
        Styles.ghost_button(
            item,
            text="\u2715",
            command=lambda: self.delete_excel(source["id"]),
            width=24,
            height=22,
            text_color=Colors.DANGER,
            hover_color=Colors.BG_CARD,
            font=Fonts.TINY,
        ).pack(side="right", padx=(0, Sizes.PAD_XS), pady=Sizes.PAD_XS)

    def refresh_image_list(self):
        """이미지 목록 새로고침"""
        for widget in self.image_list_frame.winfo_children():
            widget.destroy()

        if not self.image_mgr.images:
            ctk.CTkLabel(
                self.image_list_frame,
                text="이미지가 없습니다",
                font=Fonts.CAPTION,
                text_color=Colors.TEXT_MUTED,
            ).pack(fill="x", pady=Sizes.PAD_SM)
            return

        for image in self.image_mgr.images:
            self.create_image_item(image)

    def create_image_item(self, image):
        """이미지 아이템 생성 - compact card"""
        item = ctk.CTkFrame(
            self.image_list_frame,
            fg_color=Colors.BG_INPUT,
            corner_radius=Sizes.RADIUS_SM,
            border_width=0,
        )
        item.pack(fill="x", pady=2, padx=Sizes.PAD_XS)

        info_frame = ctk.CTkFrame(item, fg_color="transparent")
        info_frame.pack(
            side="left", fill="both", expand=True, padx=Sizes.PAD_SM, pady=Sizes.PAD_XS
        )

        # 제목 길이 제한 (최대 10자)
        MAX_TITLE_LENGTH = 10
        display_name = (
            image["name"]
            if len(image["name"]) <= MAX_TITLE_LENGTH
            else image["name"][:MAX_TITLE_LENGTH] + "..."
        )

        ctk.CTkLabel(
            info_frame,
            text=f"{image['id']}. {display_name}",
            font=Fonts.CAPTION_BOLD,
            text_color=Colors.TEXT_PRIMARY,
            anchor="w",
        ).pack(anchor="w", fill="x")

        ctk.CTkLabel(
            info_frame,
            text=f"정확도: {int(image['confidence']*100)}%",
            font=Fonts.TINY,
            text_color=Colors.TEXT_MUTED,
            anchor="w",
        ).pack(anchor="w")

        # 삭제 버튼 - ghost style, danger color, icon "✕"
        Styles.ghost_button(
            item,
            text="\u2715",
            command=lambda: self.delete_image(image["id"]),
            width=24,
            height=22,
            text_color=Colors.DANGER,
            hover_color=Colors.BG_CARD,
            font=Fonts.TINY,
        ).pack(side="right", padx=(0, Sizes.PAD_XS), pady=Sizes.PAD_XS)

    def refresh_flow_list(self):
        """플로우 목록 새로고침"""
        for widget in self.flow_list_frame.winfo_children():
            widget.destroy()

        if not self.flow_mgr.flow_sequence:
            ctk.CTkLabel(
                self.flow_list_frame,
                text="액션을 추가하세요",
                font=Fonts.SMALL,
                text_color=Colors.TEXT_MUTED,
            ).pack(fill="x", padx=Sizes.PAD_XS, pady=Sizes.PAD_LG)
            return

        for idx, action in enumerate(self.flow_mgr.flow_sequence):
            self.create_flow_item(idx, action)

    def create_flow_item(self, idx, action):
        """플로우 아이템 생성 - Card with BG_CARD"""
        item = ctk.CTkFrame(
            self.flow_list_frame,
            fg_color=Colors.BG_CARD,
            corner_radius=Sizes.RADIUS_MD,
            border_width=1,
            border_color=Colors.BORDER,
        )
        item.pack(fill="x", pady=3, padx=Sizes.PAD_SM)

        # 드래그 앤 드롭 이벤트 바인딩
        item.bind("<Button-1>", lambda e: self.on_drag_start(e, item, idx))
        item.bind("<B1-Motion>", lambda e: self.on_drag_motion(e, item))
        item.bind("<ButtonRelease-1>", lambda e: self.on_drag_release(e, item, idx))

        # 액션 타입별 색상 가져오기
        action_color = self.get_action_color(action.get("type", ""))

        # 번호 배지 (action type color)
        num_label = ctk.CTkLabel(
            item,
            text=f"{idx+1}",
            font=Fonts.BODY_BOLD,
            fg_color=action_color,
            text_color=Colors.TEXT_INVERSE,
            width=36,
            height=28,
            corner_radius=Sizes.RADIUS_SM,
        )
        num_label.pack(side="left", padx=(Sizes.PAD_SM, Sizes.PAD_XS), pady=Sizes.PAD_SM)
        num_label.bind("<Button-1>", lambda e: self.on_drag_start(e, item, idx))
        num_label.bind("<B1-Motion>", lambda e: self.on_drag_motion(e, item))
        num_label.bind(
            "<ButtonRelease-1>", lambda e: self.on_drag_release(e, item, idx)
        )

        # 액션 설명
        display_text = self.flow_mgr.get_action_display_text(
            action, self.coord_mgr, self.excel_mgr, self.image_mgr
        )

        text_label = ctk.CTkLabel(
            item,
            text=display_text,
            font=Fonts.BODY,
            text_color=Colors.TEXT_PRIMARY,
            anchor="w",
            justify="left",
        )
        text_label.pack(
            side="left", fill="both", expand=True, padx=Sizes.PAD_XS, pady=Sizes.PAD_SM
        )
        text_label.bind("<Button-1>", lambda e: self.on_drag_start(e, item, idx))
        text_label.bind("<B1-Motion>", lambda e: self.on_drag_motion(e, item))
        text_label.bind(
            "<ButtonRelease-1>", lambda e: self.on_drag_release(e, item, idx)
        )

        # 삭제 버튼 - ghost style, danger
        Styles.ghost_button(
            item,
            text="\u2715",
            command=lambda: self.delete_action(action["id"]),
            width=30,
            height=28,
            text_color=Colors.DANGER,
            hover_color=Colors.BG_ELEVATED,
            font=Fonts.CAPTION,
        ).pack(side="right", padx=(0, Sizes.PAD_SM), pady=Sizes.PAD_XS)

        # 복제 버튼 - ghost style
        Styles.ghost_button(
            item,
            text="CP",
            command=lambda: self.duplicate_action(action["id"]),
            width=34,
            height=28,
            text_color=Colors.TEXT_SECONDARY,
            hover_color=Colors.BG_ELEVATED,
            font=Fonts.CAPTION,
        ).pack(side="right", padx=(0, 2), pady=Sizes.PAD_XS)

    # 좌표 추가
    def add_coordinate_dialog(self):
        """좌표 추가 다이얼로그"""
        dialog = ctk.CTkToplevel(self.parent)
        dialog.title("좌표 추가")
        dialog.geometry("300x200")
        dialog.withdraw()
        dialog.transient(self.parent)
        dialog.grab_set()
        dialog.attributes("-topmost", True)
        dialog.configure(fg_color=Colors.BG_SECONDARY)

        ctk.CTkLabel(
            dialog,
            text="좌표 추가",
            font=Fonts.SUBHEADING,
            text_color=Colors.TEXT_PRIMARY,
        ).pack(pady=Sizes.PAD_MD)

        ctk.CTkLabel(
            dialog,
            text="3초 후 마우스가 있는 위치의\n좌표가 저장됩니다.",
            font=Fonts.CAPTION,
            text_color=Colors.TEXT_MUTED,
        ).pack(pady=Sizes.PAD_XS)

        ctk.CTkLabel(
            dialog,
            text="원하는 위치로 마우스를 이동하세요!",
            font=Fonts.CAPTION_BOLD,
            text_color=Colors.DANGER,
        ).pack(pady=Sizes.PAD_XS)

        def start_capture():
            dialog.destroy()
            self.capture_coordinate()

        Styles.primary_button(
            dialog,
            text="시작",
            command=start_capture,
            width=120,
            height=Sizes.BTN_HEIGHT,
        ).pack(pady=Sizes.PAD_MD)

        # 스페이스바와 엔터키로 시작 가능
        dialog.bind("<space>", lambda e: start_capture())
        dialog.bind("<Return>", lambda e: start_capture())

        def _show():
            center_window_on_parent(dialog, self.parent)
            dialog.deiconify()
            dialog.lift()
            dialog.focus_force()
        dialog.after(50, _show)

    def capture_coordinate(self):
        """좌표 캡처"""
        # 카운트다운 창 - tk.Toplevel 사용 (overrideredirect 필요)
        countdown_window = tk.Toplevel(self.parent)
        countdown_window.title("좌표 캡처")
        countdown_window.attributes("-topmost", True)
        countdown_window.attributes("-alpha", 0.9)

        # 화면 중앙에 배치
        window_width = 300
        window_height = 200
        screen_width = countdown_window.winfo_screenwidth()
        screen_height = countdown_window.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        countdown_window.geometry(f"{window_width}x{window_height}+{x}+{y}")

        countdown_window.overrideredirect(True)  # 테두리 제거
        countdown_window.configure(bg=Colors.BG_PRIMARY)

        label = tk.Label(
            countdown_window,
            text="3",
            font=Fonts.COUNTER,
            fg=Colors.PRIMARY_LIGHT,
            bg=Colors.BG_PRIMARY,
        )
        label.pack(expand=True)

        info_label = tk.Label(
            countdown_window,
            text="마우스를 원하는 위치로 이동하세요",
            font=Fonts.SMALL,
            fg=Colors.TEXT_SECONDARY,
            bg=Colors.BG_PRIMARY,
        )
        info_label.pack(pady=(0, 20))

        captured_data = {"x": None, "y": None, "thumbnail": None}

        def countdown(count):
            if count > 0:
                label.config(text=str(count))
                if count == 1:
                    label.config(fg=Colors.DANGER)
                countdown_window.after(1000, countdown, count - 1)
            else:
                # 좌표 캡처
                x, y, thumbnail = self.coord_mgr.capture_current_position()
                captured_data["x"] = x
                captured_data["y"] = y
                captured_data["thumbnail"] = thumbnail

                countdown_window.destroy()

                # 이름 입력 (메인 스레드에서 안전하게)
                self.after(
                    100,
                    lambda: self.show_coordinate_name_dialog(
                        captured_data["x"],
                        captured_data["y"],
                        captured_data["thumbnail"],
                    ),
                )

        countdown(3)

    def show_coordinate_name_dialog(self, x, y, thumbnail):
        """좌표 이름 입력 다이얼로그"""
        dialog = ctk.CTkToplevel(self.parent)
        dialog.title("좌표 이름 입력")
        dialog.geometry("300x200")
        dialog.withdraw()
        dialog.transient(self.parent)
        dialog.grab_set()
        dialog.attributes("-topmost", True)
        dialog.configure(fg_color=Colors.BG_SECONDARY)

        ctk.CTkLabel(
            dialog,
            text=f"좌표: ({x}, {y})",
            font=Fonts.SMALL_BOLD,
            text_color=Colors.SUCCESS,
        ).pack(pady=Sizes.PAD_SM)

        ctk.CTkLabel(
            dialog, text="이름:", font=Fonts.SMALL, text_color=Colors.TEXT_PRIMARY
        ).pack(anchor="w", padx=Sizes.PAD_XL, pady=(0, Sizes.PAD_XS))

        name_entry = Styles.input_field(
            dialog,
            width=240,
        )
        name_entry.pack(padx=Sizes.PAD_XL, pady=(0, Sizes.PAD_SM))
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

        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(pady=Sizes.PAD_SM)

        Styles.primary_button(
            btn_frame,
            text="저장",
            command=save_coord,
            width=80,
            height=Sizes.BTN_HEIGHT_SM,
            fg_color=Colors.SUCCESS,
            hover_color=Colors.SUCCESS_HOVER,
        ).pack(side="left", padx=Sizes.PAD_XS)

        Styles.secondary_button(
            btn_frame,
            text="취소",
            command=on_cancel,
            width=80,
            height=Sizes.BTN_HEIGHT_SM,
        ).pack(side="left", padx=Sizes.PAD_XS)

        def _show():
            center_window_on_parent(dialog, self.parent)
            dialog.deiconify()
            dialog.lift()
            dialog.focus_force()
        dialog.after(50, _show)
        # Enter 키 바인딩
        name_entry.bind("<Return>", lambda e: save_coord())

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
            initialdir=os.path.expanduser("~"),
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
            initial_value="",
        )
        self.parent.wait_window(dialog)

        name = dialog.result
        if not name:
            return

        # 추가 (기본값으로 자동 최신 파일 선택 비활성화, 중복 제거 활성화)
        source = self.excel_mgr.add_excel_source(
            name,
            filepath,
            sheet_name,
            selected_columns,
            auto_latest=False,
            auto_directory=None,
            auto_prefix="list",
            remove_empty_rows=True,
            remove_duplicates=True,
        )
        if source:
            self.refresh_excel_list()
            messagebox.showinfo(
                "완료",
                f"엑셀 데이터 '{name}'이(가) 추가되었습니다.\n\n행 수: {source['row_count']}\n칼럼 수: {len(selected_columns)}",
            )
        else:
            messagebox.showerror("오류", "엑셀 데이터를 추가할 수 없습니다.")

    def select_sheet_dialog(self, sheets):
        """시트 선택 다이얼로그"""
        dialog = ctk.CTkToplevel(self.parent)
        dialog.title("시트 선택")
        dialog.geometry("300x400")
        dialog.withdraw()
        dialog.transient(self.parent)
        dialog.grab_set()
        dialog.attributes("-topmost", True)
        dialog.configure(fg_color=Colors.BG_SECONDARY)

        result = [None]

        ctk.CTkLabel(
            dialog,
            text="시트를 선택하세요",
            font=Fonts.BODY_BOLD,
            text_color=Colors.TEXT_PRIMARY,
        ).pack(pady=Sizes.PAD_SM)

        # 버튼 리스트로 대체
        list_frame = ctk.CTkScrollableFrame(
            dialog,
            fg_color=Colors.BG_CARD,
            corner_radius=Sizes.RADIUS_MD,
            border_width=1,
            border_color=Colors.BORDER,
        )
        list_frame.pack(fill="both", expand=True, padx=Sizes.PAD_LG, pady=Sizes.PAD_SM)

        selected_idx = [0]
        sheet_buttons = []

        def select_sheet(idx):
            selected_idx[0] = idx
            for i, btn in enumerate(sheet_buttons):
                if i == idx:
                    btn.configure(fg_color=Colors.PRIMARY)
                else:
                    btn.configure(fg_color=Colors.BG_INPUT)

        def on_select():
            result[0] = sheets[selected_idx[0]]
            dialog.destroy()

        for i, sheet in enumerate(sheets):
            btn = ctk.CTkButton(
                list_frame,
                text=sheet,
                font=Fonts.SMALL,
                fg_color=Colors.PRIMARY if i == 0 else Colors.BG_INPUT,
                hover_color=Colors.PRIMARY_HOVER,
                text_color=Colors.TEXT_PRIMARY,
                height=Sizes.BTN_HEIGHT,
                corner_radius=Sizes.RADIUS_SM,
                anchor="w",
                command=lambda idx=i: select_sheet(idx),
            )
            btn.pack(fill="x", pady=2)
            btn.bind(
                "<Double-Button-1>", lambda e, idx=i: (select_sheet(idx), on_select())
            )
            sheet_buttons.append(btn)

        Styles.primary_button(
            dialog,
            text="선택",
            command=on_select,
            width=100,
            height=Sizes.BTN_HEIGHT,
        ).pack(pady=Sizes.PAD_SM)

        dialog.bind("<Return>", lambda e: on_select())

        def _show():
            center_window_on_parent(dialog, self.parent)
            dialog.deiconify()
            dialog.lift()
            dialog.focus_force()
        dialog.after(50, _show)

        dialog.wait_window()
        return result[0]

    def select_columns_dialog(self, columns):
        """칼럼 선택 다이얼로그 (개선됨)"""
        dialog = ctk.CTkToplevel(self.parent)
        dialog.title("칼럼 선택")
        dialog.geometry("300x500")
        dialog.withdraw()
        dialog.transient(self.parent)
        dialog.grab_set()
        dialog.configure(fg_color=Colors.BG_SECONDARY)

        result = [None]

        # 제목
        ctk.CTkLabel(
            dialog,
            text="사용할 칼럼을 선택하세요",
            font=Fonts.SUBHEADING,
            text_color=Colors.TEXT_PRIMARY,
        ).pack(pady=Sizes.PAD_MD)

        # 전체 선택/해제 버튼
        button_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        button_frame.pack(fill="x", padx=Sizes.PAD_LG, pady=(0, Sizes.PAD_SM))

        vars_dict = {}  # 미리 선언 (함수에서 사용)

        def select_all():
            for var in vars_dict.values():
                var.set(True)

        def deselect_all():
            for var in vars_dict.values():
                var.set(False)

        ctk.CTkButton(
            button_frame,
            text="전체 선택",
            font=Fonts.CAPTION_BOLD,
            fg_color=Colors.SUCCESS,
            hover_color=Colors.SUCCESS_HOVER,
            text_color=Colors.TEXT_INVERSE,
            width=80,
            height=Sizes.BTN_HEIGHT_SM,
            corner_radius=Sizes.RADIUS_SM,
            command=select_all,
        ).pack(side="left", padx=Sizes.PAD_XS)

        ctk.CTkButton(
            button_frame,
            text="전체 해제",
            font=Fonts.CAPTION_BOLD,
            fg_color=Colors.DANGER,
            hover_color=Colors.DANGER_HOVER,
            text_color=Colors.TEXT_INVERSE,
            width=80,
            height=Sizes.BTN_HEIGHT_SM,
            corner_radius=Sizes.RADIUS_SM,
            command=deselect_all,
        ).pack(side="left", padx=Sizes.PAD_XS)

        # 체크박스 리스트
        scrollable_frame = ctk.CTkScrollableFrame(
            dialog,
            fg_color=Colors.BG_CARD,
            corner_radius=Sizes.RADIUS_MD,
            border_width=1,
            border_color=Colors.BORDER,
        )
        scrollable_frame.pack(
            fill="both", expand=True, padx=Sizes.PAD_LG, pady=Sizes.PAD_SM
        )

        # 체크박스 생성
        for col in columns:
            var = tk.BooleanVar(value=True)
            vars_dict[col] = var

            cb = ctk.CTkCheckBox(
                scrollable_frame,
                text=col,
                variable=var,
                font=Fonts.SMALL,
                text_color=Colors.TEXT_PRIMARY,
                fg_color=Colors.PRIMARY,
                hover_color=Colors.PRIMARY_HOVER,
                corner_radius=4,
            )
            cb.pack(anchor="w", padx=Sizes.PAD_SM, pady=3)

        # 선택 개수 표시
        count_label = ctk.CTkLabel(
            dialog,
            text=f"선택된 칼럼: {len(columns)}개 / 전체: {len(columns)}개",
            font=Fonts.CAPTION,
            text_color=Colors.TEXT_MUTED,
        )
        count_label.pack(pady=Sizes.PAD_XS)

        def update_count(*args):
            selected_count = sum(1 for var in vars_dict.values() if var.get())
            count_label.configure(
                text=f"선택된 칼럼: {selected_count}개 / 전체: {len(columns)}개"
            )

        # 체크박스 변경 시 카운트 업데이트
        for var in vars_dict.values():
            var.trace_add("write", update_count)

        # 확인/취소 버튼
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(pady=Sizes.PAD_MD)

        def on_confirm():
            selected = [col for col, var in vars_dict.items() if var.get()]
            if not selected:
                messagebox.showwarning(
                    "경고", "최소 하나의 칼럼을 선택하세요.", parent=dialog
                )
                return
            result[0] = selected
            dialog.destroy()

        def on_cancel():
            result[0] = None
            dialog.destroy()

        Styles.primary_button(
            btn_frame,
            text="확인",
            command=on_confirm,
            width=100,
            height=Sizes.BTN_HEIGHT,
        ).pack(side="left", padx=Sizes.PAD_XS)

        Styles.secondary_button(
            btn_frame,
            text="취소",
            command=on_cancel,
            width=100,
            height=Sizes.BTN_HEIGHT,
        ).pack(side="left", padx=Sizes.PAD_XS)

        # Enter 키로 확인, ESC 키로 취소
        dialog.bind("<Return>", lambda e: on_confirm())
        dialog.bind("<Escape>", lambda e: on_cancel())

        # 중앙 배치 및 포커스
        dialog.attributes("-topmost", True)
        def _show_cols():
            center_window_on_parent(dialog, self.parent)
            dialog.deiconify()
            dialog.lift()
            dialog.focus_force()
        dialog.after(50, _show_cols)

        dialog.wait_window()
        return result[0]

    def select_data_options_dialog(self):
        """데이터 처리 옵션 선택 다이얼로그"""
        dialog = ctk.CTkToplevel(self.parent)
        dialog.title("데이터 처리 옵션")
        dialog.geometry("400x320")
        dialog.withdraw()
        dialog.transient(self.parent)
        dialog.grab_set()
        dialog.attributes("-topmost", True)
        dialog.configure(fg_color=Colors.BG_SECONDARY)

        result = [None]

        # 제목
        ctk.CTkLabel(
            dialog,
            text="데이터 처리 옵션",
            font=Fonts.HEADING,
            text_color=Colors.TEXT_PRIMARY,
        ).pack(pady=Sizes.PAD_LG)

        # 설명
        ctk.CTkLabel(
            dialog,
            text="엑셀 데이터 로드 시 적용할 옵션을 선택하세요",
            font=Fonts.CAPTION,
            text_color=Colors.TEXT_MUTED,
        ).pack(pady=(0, Sizes.PAD_LG))

        # 옵션 프레임
        options_frame = Styles.card(dialog)
        options_frame.pack(fill="both", expand=True, padx=Sizes.PAD_LG)

        # 공백 행 제거 옵션
        remove_empty_var = tk.BooleanVar(value=True)

        empty_frame = ctk.CTkFrame(options_frame, fg_color="transparent")
        empty_frame.pack(fill="x", padx=Sizes.PAD_LG, pady=Sizes.PAD_SM)

        ctk.CTkCheckBox(
            empty_frame,
            text="상위 공백 행 제거",
            variable=remove_empty_var,
            font=Fonts.BODY_BOLD,
            text_color=Colors.TEXT_PRIMARY,
            fg_color=Colors.PRIMARY,
            hover_color=Colors.PRIMARY_HOVER,
            corner_radius=4,
        ).pack(anchor="w")

        ctk.CTkLabel(
            empty_frame,
            text="   데이터 시작 전의 빈 행을 자동으로 제거합니다",
            font=Fonts.CAPTION,
            text_color=Colors.TEXT_MUTED,
        ).pack(anchor="w")

        # 중복 제거 옵션
        remove_dup_var = tk.BooleanVar(value=False)

        dup_frame = ctk.CTkFrame(options_frame, fg_color="transparent")
        dup_frame.pack(fill="x", padx=Sizes.PAD_LG, pady=Sizes.PAD_SM)

        ctk.CTkCheckBox(
            dup_frame,
            text="중복 행 제거",
            variable=remove_dup_var,
            font=Fonts.BODY_BOLD,
            text_color=Colors.TEXT_PRIMARY,
            fg_color=Colors.PRIMARY,
            hover_color=Colors.PRIMARY_HOVER,
            corner_radius=4,
        ).pack(anchor="w")

        ctk.CTkLabel(
            dup_frame,
            text="   선택한 컬럼 기준으로 중복된 행을 제거합니다",
            font=Fonts.CAPTION,
            text_color=Colors.TEXT_MUTED,
        ).pack(anchor="w")

        # 버튼
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(pady=Sizes.PAD_LG)

        def on_confirm():
            result[0] = {
                "remove_empty": remove_empty_var.get(),
                "remove_duplicates": remove_dup_var.get(),
            }
            dialog.destroy()

        def on_cancel():
            result[0] = None
            dialog.destroy()

        Styles.primary_button(
            btn_frame,
            text="확인",
            command=on_confirm,
            width=100,
            height=Sizes.BTN_HEIGHT,
        ).pack(side="left", padx=Sizes.PAD_XS)

        Styles.secondary_button(
            btn_frame,
            text="취소",
            command=on_cancel,
            width=100,
            height=Sizes.BTN_HEIGHT,
        ).pack(side="left", padx=Sizes.PAD_XS)

        # 중앙 배치
        def _show_opts():
            center_window_on_parent(dialog, self.parent)
            dialog.deiconify()
            dialog.lift()
            dialog.focus_force()
        dialog.after(50, _show_opts)

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
            initialdir=os.path.expanduser("~"),
        )

        if not filepath:
            return

        # 이미지를 base64로 변환
        import base64

        with open(filepath, "rb") as f:
            img_data = base64.b64encode(f.read()).decode()

        # 이름 및 정확도 입력
        from ui.dialogs import ImageNameDialog

        dialog = ImageNameDialog(
            self.parent, title="이미지 등록", initial_name="", initial_confidence=80
        )
        self.parent.wait_window(dialog)

        if not dialog.result:
            return

        name = dialog.result["name"]
        confidence = dialog.result["confidence"]

        # 추가
        image = self.image_mgr.add_image(name, img_data, confidence=confidence)
        if image:
            self.refresh_image_list()
            messagebox.showinfo(
                "완료",
                f"이미지 '{name}'이(가) 추가되었습니다.\n정확도: {int(confidence*100)}%",
            )
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
            self.parent, self.coord_mgr, self.excel_mgr, self.image_mgr
        )
        self.parent.wait_window(dialog)

        if dialog.result:
            action = self.flow_mgr.add_action(
                dialog.result["type"], dialog.result["params"]
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

    def duplicate_action(self, action_id):
        """액션 복제"""
        new_action = self.flow_mgr.duplicate_action(action_id)
        if new_action:
            self.refresh_flow_list()

    def delete_action(self, action_id):
        """액션 삭제"""
        if messagebox.askyesno("확인", "이 액션을 삭제하시겠습니까?"):
            self.flow_mgr.remove_action(action_id)
            self.refresh_flow_list()

    # 프로젝트 관리
    def save_project(self):
        """프로젝트 저장"""
        self.project_data["coordinates"] = self.coord_mgr.to_list()
        self.project_data["excel_sources"] = self.excel_mgr.to_list()
        self.project_data["images"] = self.image_mgr.to_list()
        self.project_data["flow_sequence"] = self.flow_mgr.to_list()

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
            "click_coord": Colors.ACTION_MOUSE,
            "click_image": Colors.ACTION_MOUSE,
            "mouse_scroll": Colors.ACTION_MOUSE,
            # 키보드 동작 - 녹색
            "type_text": Colors.ACTION_KEYBOARD,
            "type_variable": Colors.ACTION_KEYBOARD,
            "key_press": Colors.ACTION_KEYBOARD,
            "paste": Colors.ACTION_KEYBOARD,
            # 제어 동작 - 빨간색
            "delay": Colors.ACTION_CONTROL,
            "wait_image": Colors.ACTION_CONTROL,
            # 기타 - 노란색
            "screenshot": Colors.ACTION_OTHER,
        }

        return color_map.get(action_type, Colors.ACTION_DEFAULT)

    def get_action_category(self, action_type):
        """액션 카테고리 반환"""
        categories = {
            # 마우스 동작
            "click_coord": "마우스",
            "click_image": "마우스",
            "mouse_scroll": "마우스",
            # 키보드 동작
            "type_text": "키보드",
            "type_variable": "키보드",
            "key_press": "키보드",
            "paste": "키보드",
            # 제어 동작
            "delay": "제어",
            "wait_image": "제어",
            # 기타
            "screenshot": "기타",
        }

        return categories.get(action_type, "기타")

    # 드래그 앤 드롭 이벤트 핸들러
    def _get_drop_index(self, event):
        """마우스 위치에서 드롭 인덱스 계산"""
        children = list(self.flow_list_frame.winfo_children())
        # 드롭 인디케이터는 제외
        children = [c for c in children if c != self.drop_indicator]
        if not children:
            return 0

        mouse_y = event.y_root - self.flow_list_frame.winfo_rooty()

        for i, child in enumerate(children):
            child_mid = child.winfo_y() + child.winfo_height() // 2
            if mouse_y < child_mid:
                return i
        return len(children)

    def _show_drop_indicator(self, drop_idx):
        """드롭 위치에 파란 라인 표시"""
        # 기존 인디케이터 제거
        if self.drop_indicator:
            self.drop_indicator.destroy()
            self.drop_indicator = None

        self.drop_indicator = ctk.CTkFrame(
            self.flow_list_frame,
            fg_color=Colors.DRAG_INDICATOR,
            height=3,
            corner_radius=2,
        )

        children = [c for c in self.flow_list_frame.winfo_children()
                     if c != self.drop_indicator]

        if drop_idx >= len(children):
            self.drop_indicator.pack(fill="x", padx=Sizes.PAD_SM, pady=1)
        else:
            self.drop_indicator.pack(
                fill="x", padx=Sizes.PAD_SM, pady=1,
                before=children[drop_idx],
            )

    def _remove_drop_indicator(self):
        """드롭 인디케이터 제거"""
        if self.drop_indicator:
            self.drop_indicator.destroy()
            self.drop_indicator = None

    def on_drag_start(self, event, item, idx):
        """드래그 시작"""
        # 버튼 클릭은 드래그 무시
        widget_class = event.widget.winfo_class()
        if widget_class in ("Button", "CTkButton"):
            return

        self.drag_data["item"] = item
        self.drag_data["index"] = idx
        self.drag_data["y"] = event.y_root

        # 드래그 중 시각적 피드백
        item.configure(
            fg_color=Colors.BG_ELEVATED,
            border_color=Colors.PRIMARY_LIGHT,
            border_width=2,
        )

    def on_drag_motion(self, event, item):
        """드래그 중"""
        if self.drag_data["item"] is None:
            return

        # 드롭 위치 인디케이터 표시
        drop_idx = self._get_drop_index(event)
        self._show_drop_indicator(drop_idx)

    def on_drag_release(self, event, item, idx):
        """드롭"""
        # 드롭 인디케이터 제거
        self._remove_drop_indicator()

        if self.drag_data["item"] is None:
            return

        # 버튼 클릭은 드래그 무시
        widget_class = event.widget.winfo_class()
        if widget_class in ("Button", "CTkButton"):
            self.drag_data["item"].configure(
                fg_color=Colors.BG_CARD, border_color=Colors.BORDER, border_width=1
            )
            self.drag_data["item"] = None
            return

        drag_idx = self.drag_data["index"]
        drop_idx = self._get_drop_index(event)

        # 드래그 인덱스 이후면 보정
        if drop_idx > drag_idx:
            drop_idx = min(drop_idx, len(self.flow_mgr.flow_sequence)) - 1
        drop_idx = max(0, min(drop_idx, len(self.flow_mgr.flow_sequence) - 1))

        # 드래그 상태 초기화
        self.drag_data["item"].configure(
            fg_color=Colors.BG_CARD, border_color=Colors.BORDER, border_width=1
        )
        self.drag_data["item"] = None

        # 순서 변경
        if drag_idx != drop_idx:
            action = self.flow_mgr.flow_sequence.pop(drag_idx)
            self.flow_mgr.flow_sequence.insert(drop_idx, action)
            self.refresh_flow_list()
