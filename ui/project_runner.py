"""
프로젝트 실행 화면 - 매크로 실행 및 모니터링
"""

import customtkinter as ctk
from tkinter import messagebox
import threading
from pynput import keyboard  # ← 추가
from core.project_manager import ProjectManager
from utils.ui_helpers import set_dialog_icon, center_window_on_parent
from core.coordinate_manager import CoordinateManager
from core.excel_manager import ExcelManager
from core.image_manager import ImageManager
from core.flow_manager import FlowManager
from core.executor import MacroExecutor
from ui.theme import Colors, Fonts, Sizes, Styles


class ProjectRunner(ctk.CTkFrame):
    def __init__(self, parent, app, project_data, filepath, autostart=False):
        super().__init__(parent, fg_color=Colors.BG_PRIMARY)
        self.app = app
        self.parent = parent
        self.project_data = project_data
        self.filepath = filepath
        self.autostart = autostart

        # 관리자 초기화
        self.coord_mgr = CoordinateManager()
        self.coord_mgr.load_from_list(project_data.get("coordinates", []))

        self.excel_mgr = ExcelManager()
        self.excel_mgr.load_from_list(project_data.get("excel_sources", []))

        self.image_mgr = ImageManager()
        self.image_mgr.load_from_list(project_data.get("images", []))

        self.flow_mgr = FlowManager()
        self.flow_mgr.load_from_list(project_data.get("flow_sequence", []))

        # 실행 엔진
        self.executor = MacroExecutor(
            project_data,
            self.coord_mgr,
            self.excel_mgr,
            self.image_mgr,
            self.flow_mgr,
            filepath,  # 프로젝트 파일 경로 전달
        )

        self.executor.set_callbacks(
            log_cb=self.add_log,
            # progress_cb=self.update_progress,
            error_cb=self.report_error,
        )

        self.is_running = False

        # 단축키 리스너
        self.hotkey_listener = None
        self.setup_hotkeys()

        self.setup_ui()

    def setup_hotkeys(self):
        """단축키 설정 (전역 설정 사용)"""
        from core.settings_manager import SettingsManager

        hotkeys = SettingsManager.get_hotkeys()

        # 단축키를 pynput 형식으로 변환
        self.hotkey_map = {}
        for action, key in hotkeys.items():
            # 빈 문자열이나 None 값 체크
            if not key:
                continue

            key = key.strip().lower()  # 공백 제거 후 소문자 변환

            # F8, F9 등을 keyboard.Key.f8 형식으로 변환
            if key.startswith("f") and len(key) > 1 and key[1:].isdigit():
                # F키 (f1 ~ f24)
                key_obj = getattr(keyboard.Key, key, None)
            elif key in ["enter", "tab", "esc", "space", "backspace", "delete"]:
                # 특수 키
                key_obj = getattr(keyboard.Key, key, None)
            else:
                # 일반 키
                key_obj = key

            if key_obj:
                self.hotkey_map[action] = key_obj

        # 리스너 시작
        self.start_hotkey_listener()

    def start_hotkey_listener(self):
        """단축키 리스너 시작"""

        def on_press(key):
            try:
                # 현재 누른 키 이름 가져오기
                if hasattr(key, "name"):
                    # F키 등 special keys
                    key_name = key
                else:
                    # 일반 문자 키
                    key_name = (
                        key.char.lower() if hasattr(key, "char") and key.char else None
                    )

                # None 체크
                if key_name is None:
                    return

                # 단축키 매칭
                if key_name == self.hotkey_map.get("start"):
                    self.parent.after(0, self.start_macro)
                elif key_name == self.hotkey_map.get("pause"):
                    self.parent.after(0, self.pause_macro)
                elif key_name == self.hotkey_map.get("stop"):
                    self.parent.after(0, self.stop_macro)
                elif key_name == self.hotkey_map.get("focus"):
                    self.parent.after(0, self.bring_to_front)
            except AttributeError as e:
                # None에서 메서드 호출하려고 할 때 발생
                pass
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
        # ─── 상단 헤더 ───
        header = ctk.CTkFrame(
            self, fg_color=Colors.BG_CARD, height=Sizes.HEADER_HEIGHT, corner_radius=0
        )
        header.pack(fill="x")
        header.pack_propagate(False)

        ctk.CTkLabel(
            header,
            text=self.project_data['name'],
            font=Fonts.HEADING,
            text_color=Colors.TEXT_PRIMARY,
        ).pack(side="left", padx=Sizes.PAD_XL, pady=Sizes.PAD_MD)

        btn_frame = ctk.CTkFrame(header, fg_color="transparent")
        btn_frame.pack(side="right", padx=Sizes.PAD_XL)

        Styles.ghost_button(
            btn_frame, text="설정", command=self.show_settings, width=60,
        ).pack(side="left", padx=Sizes.PAD_XS)

        Styles.ghost_button(
            btn_frame, text="편집", command=self.edit_project, width=60,
        ).pack(side="left", padx=Sizes.PAD_XS)

        Styles.ghost_button(
            btn_frame, text="홈", command=self.go_home, width=50,
        ).pack(side="left", padx=Sizes.PAD_XS)

        # ─── 메인 컨텐츠 ───
        content = ctk.CTkFrame(self, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=Sizes.PAD_XL, pady=Sizes.PAD_LG)

        # ─── 단축키 정보 카드 ───
        hotkey_card = Styles.card(content, border_color=Colors.BORDER_SUBTLE)
        hotkey_card.pack(fill="x", pady=(0, Sizes.PAD_LG))

        hotkey_inner = ctk.CTkFrame(hotkey_card, fg_color="transparent")
        hotkey_inner.pack(fill="x", padx=Sizes.PAD_LG, pady=Sizes.PAD_MD)

        Styles.section_title(hotkey_inner, text="단축키").pack(anchor="w", pady=(0, Sizes.PAD_SM))

        # 전역 설정에서 단축키 가져오기
        from core.settings_manager import SettingsManager

        hotkeys = SettingsManager.get_hotkeys()

        hotkey_row = ctk.CTkFrame(hotkey_inner, fg_color="transparent")
        hotkey_row.pack(fill="x")

        key_items = [
            ("시작", hotkeys.get("start", "F9")),
            ("일시정지", hotkeys.get("pause", "F10")),
            ("중지", hotkeys.get("stop", "F11")),
            ("맨 앞으로", hotkeys.get("focus", "F12")),
        ]

        for i, (label, key_val) in enumerate(key_items):
            item_frame = ctk.CTkFrame(hotkey_row, fg_color="transparent")
            item_frame.pack(side="left", padx=(0, Sizes.PAD_XL))

            ctk.CTkLabel(
                item_frame,
                text=label,
                font=Fonts.CAPTION,
                text_color=Colors.TEXT_MUTED,
            ).pack(side="left", padx=(0, Sizes.PAD_XS))

            ctk.CTkLabel(
                item_frame,
                text=key_val.upper(),
                font=Fonts.MONO,
                text_color=Colors.PRIMARY_LIGHT,
            ).pack(side="left")

        # ─── 프로젝트 정보 카드 ───
        info_card = Styles.card(content, border_color=Colors.BORDER_SUBTLE)
        info_card.pack(fill="x", pady=(0, Sizes.PAD_LG))

        info_inner = ctk.CTkFrame(info_card, fg_color="transparent")
        info_inner.pack(fill="x", padx=Sizes.PAD_LG, pady=Sizes.PAD_MD)

        Styles.section_title(info_inner, text="프로젝트 정보").pack(anchor="w", pady=(0, Sizes.PAD_SM))

        info_grid = ctk.CTkFrame(info_inner, fg_color="transparent")
        info_grid.pack(fill="x")
        info_grid.columnconfigure(1, weight=1)

        # 좌표 행
        ctk.CTkLabel(
            info_grid, text="좌표", font=Fonts.SMALL, text_color=Colors.TEXT_MUTED
        ).grid(row=0, column=0, sticky="w", padx=(0, Sizes.PAD_LG), pady=Sizes.PAD_XS)
        ctk.CTkLabel(
            info_grid,
            text=f"{len(self.coord_mgr.coordinates)}개",
            font=Fonts.SMALL_BOLD,
            text_color=Colors.TEXT_PRIMARY,
        ).grid(row=0, column=1, sticky="w", pady=Sizes.PAD_XS)

        # ─── 엑셀 정보 표시 (플로우에 사용된 순서대로) ───
        if self.excel_mgr.excel_sources:
            # 각 엑셀별로 사용하는 컬럼 추출
            import re

            excel_usage = {}  # {excel_id: set(columns)}
            excel_order = []  # 플로우에서 사용된 순서대로 excel_id

            for action in self.flow_mgr.flow_sequence:
                action_type = action.get("type", "")

                # type_variable 액션에서 엑셀 컬럼 추출
                if action_type == "type_variable":
                    params = action.get("params", {})
                    if params.get("var_type") == "excel":
                        var_name = params.get("var_name", "")
                        excel_id = params.get("excel_id", 1)  # 기본값 1
                        if var_name:
                            if excel_id not in excel_usage:
                                excel_usage[excel_id] = set()
                                excel_order.append(excel_id)
                            excel_usage[excel_id].add(var_name)

                # 다른 액션에서 {{엑셀ID_컬럼명}} 패턴 찾기
                elif action_type in ["type", "click", "hotkey", "wait"]:
                    value = action.get("value", "")
                    if isinstance(value, str) and "{{엑셀" in value:
                        matches = re.findall(r"\{\{엑셀(\d+)_([^}]+)\}\}", value)
                        for excel_id_str, col_name in matches:
                            excel_id = int(excel_id_str)
                            if excel_id not in excel_usage:
                                excel_usage[excel_id] = set()
                                excel_order.append(excel_id)
                            excel_usage[excel_id].add(col_name)

            # 플로우에 사용된 엑셀만 순서대로 표시
            if excel_usage:
                ctk.CTkLabel(
                    info_grid,
                    text="엑셀",
                    font=Fonts.SMALL,
                    text_color=Colors.TEXT_MUTED,
                ).grid(
                    row=1, column=0, sticky="nw", padx=(0, Sizes.PAD_LG), pady=Sizes.PAD_XS
                )

                excel_container = ctk.CTkFrame(info_grid, fg_color="transparent")
                excel_container.grid(
                    row=1, column=1, columnspan=2, sticky="w", pady=Sizes.PAD_XS
                )

                for order_idx, excel_id in enumerate(excel_order):
                    # excel_id로 source 찾기
                    source = None
                    source_name = f"엑셀{excel_id}"
                    for src in self.excel_mgr.excel_sources:
                        if src.get("id", 0) == excel_id:
                            source = src
                            source_name = src["name"]
                            break

                    # 각 엑셀을 한 줄로 표시
                    row_frame = ctk.CTkFrame(excel_container, fg_color="transparent")
                    row_frame.pack(fill="x", pady=2)

                    # 순번
                    ctk.CTkLabel(
                        row_frame,
                        text=f"{order_idx + 1}.",
                        font=Fonts.CAPTION,
                        text_color=Colors.TEXT_MUTED,
                        width=20,
                    ).pack(side="left")

                    # 엑셀 경로 표시 (자동 선택 vs 수동 선택 구분)
                    if source and source.get("auto_latest", False):
                        # 자동 파일 선택: [디렉토리, 접두사, 모드] 형식
                        auto_dir = source.get("auto_directory", "")
                        auto_prefix = source.get("auto_prefix", "")
                        auto_mode = source.get("auto_mode", "modified")
                        mode_text = "최근수정" if auto_mode == "modified" else "최근번호"
                        prefix_text = f", {auto_prefix}" if auto_prefix else ""
                        excel_path = f"[{auto_dir}{prefix_text}] {mode_text}"
                        path_color = Colors.PRIMARY_LIGHT
                    elif source:
                        # 수동 파일 선택: 전체 경로
                        excel_path = f"[{source.get('filepath', 'N/A')}]"
                        path_color = Colors.TEXT_PRIMARY
                    else:
                        excel_path = "[N/A]"
                        path_color = Colors.TEXT_MUTED

                    # 사용 컬럼
                    used_columns = sorted(list(excel_usage.get(excel_id, set())))
                    col_text = ", ".join(used_columns) if used_columns else ""

                    # 경로 표시
                    ctk.CTkLabel(
                        row_frame,
                        text=excel_path,
                        font=Fonts.CAPTION,
                        text_color=path_color,
                        anchor="w",
                    ).pack(side="left", padx=(0, Sizes.PAD_XS))

                    # 사용 컬럼 표시
                    if col_text:
                        ctk.CTkLabel(
                            row_frame,
                            text=f"({col_text})",
                            font=Fonts.CAPTION,
                            text_color=Colors.TEXT_MUTED,
                            anchor="w",
                        ).pack(side="left", fill="x", expand=True)

                    # 설정 버튼 (ghost style with gear icon)
                    if source:
                        Styles.ghost_button(
                            row_frame,
                            text="SET",
                            command=lambda s=source: self.show_excel_config(s),
                            width=30,
                            height=Sizes.BTN_HEIGHT_XS,
                            font=Fonts.SMALL,
                        ).pack(side="right")
            else:
                ctk.CTkLabel(
                    info_grid,
                    text="엑셀",
                    font=Fonts.SMALL,
                    text_color=Colors.TEXT_MUTED,
                ).grid(row=1, column=0, sticky="w", padx=(0, Sizes.PAD_LG), pady=Sizes.PAD_XS)
                ctk.CTkLabel(
                    info_grid,
                    text="사용 안 함",
                    font=Fonts.CAPTION,
                    text_color=Colors.TEXT_MUTED,
                ).grid(row=1, column=1, sticky="w", pady=Sizes.PAD_XS)
        else:
            ctk.CTkLabel(
                info_grid,
                text="엑셀",
                font=Fonts.SMALL,
                text_color=Colors.TEXT_MUTED,
            ).grid(row=1, column=0, sticky="w", padx=(0, Sizes.PAD_LG), pady=Sizes.PAD_XS)
            ctk.CTkLabel(
                info_grid, text="없음", font=Fonts.CAPTION, text_color=Colors.TEXT_MUTED
            ).grid(row=1, column=1, sticky="w", pady=Sizes.PAD_XS)

        # 이미지 행
        ctk.CTkLabel(
            info_grid,
            text="이미지",
            font=Fonts.SMALL,
            text_color=Colors.TEXT_MUTED,
        ).grid(row=2, column=0, sticky="w", padx=(0, Sizes.PAD_LG), pady=Sizes.PAD_XS)
        ctk.CTkLabel(
            info_grid,
            text=f"{len(self.image_mgr.images)}개",
            font=Fonts.SMALL_BOLD,
            text_color=Colors.TEXT_PRIMARY,
        ).grid(row=2, column=1, sticky="w", pady=Sizes.PAD_XS)

        # 플로우 행
        ctk.CTkLabel(
            info_grid,
            text="플로우",
            font=Fonts.SMALL,
            text_color=Colors.TEXT_MUTED,
        ).grid(row=3, column=0, sticky="w", padx=(0, Sizes.PAD_LG), pady=Sizes.PAD_XS)
        ctk.CTkLabel(
            info_grid,
            text=f"{len(self.flow_mgr.flow_sequence)}개 액션",
            font=Fonts.SMALL_BOLD,
            text_color=Colors.TEXT_PRIMARY,
        ).grid(row=3, column=1, sticky="w", pady=Sizes.PAD_XS)

        # ─── 로그 카드 ───
        log_card = Styles.card(content, border_color=Colors.BORDER_SUBTLE)
        log_card.pack(fill="both", expand=True, pady=(0, Sizes.PAD_LG))

        log_header = ctk.CTkFrame(log_card, fg_color="transparent")
        log_header.pack(fill="x", padx=Sizes.PAD_LG, pady=(Sizes.PAD_MD, 0))

        Styles.section_title(log_header, text="실행 로그").pack(side="left")

        # 로그 내보내기 버튼
        Styles.ghost_button(
            log_header,
            text="로그 저장",
            command=self.export_log,
            width=80,
        ).pack(side="right")

        self.log_text = ctk.CTkTextbox(
            log_card,
            font=Fonts.MONO_SMALL,
            height=200,
            fg_color=Colors.BG_INPUT,
            text_color=Colors.TEXT_PRIMARY,
            corner_radius=Sizes.RADIUS_SM,
            border_width=1,
            border_color=Colors.BORDER,
            state="disabled",
        )
        self.log_text.pack(
            fill="both",
            expand=True,
            padx=Sizes.PAD_LG,
            pady=(Sizes.PAD_SM, Sizes.PAD_LG),
        )

        # 자동시작인 경우 바로 실행
        if self.autostart:
            self.after(500, self.start_macro)

    def add_log(self, message):
        """로그 추가"""
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    # def update_progress(self, current, total, status=""):
    #     """진행상황 업데이트"""
    #     if total > 0:
    #         percentage = (current / total) * 100
    #         self.progress_bar['value'] = percentage
    #         self.progress_label.config(text=f"{current}/{total} ({percentage:.1f}%)")
    #     else:
    #         self.progress_bar['mode'] = 'indeterminate'
    #         self.progress_bar.start()
    #         self.progress_label.config(text=f"반복 {current}")

    #     self.status_label.config(text=status)

    def report_error(self, error_msg, screenshot=None):
        """에러 보고"""
        pass

    def export_log(self):
        """로그를 파일로 내보내기"""
        from tkinter import filedialog
        from datetime import datetime

        self.log_text.configure(state="normal")
        log_content = self.log_text.get("1.0", "end").strip()
        self.log_text.configure(state="disabled")
        if not log_content:
            messagebox.showinfo("알림", "저장할 로그가 없습니다.")
            return

        default_name = f"macro_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        filepath = filedialog.asksaveasfilename(
            parent=self.parent,
            title="로그 저장",
            initialfile=default_name,
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        )

        if filepath:
            try:
                with open(filepath, "w", encoding="utf-8") as f:
                    f.write(log_content)
                messagebox.showinfo("완료", f"로그가 저장되었습니다.\n{filepath}")
            except (OSError, IOError) as e:
                messagebox.showerror("오류", f"로그 저장 실패: {e}")

    def start_macro(self):
        """매크로 시작"""
        if self.is_running:
            return

        # 유효성 검사
        if not self.flow_mgr.flow_sequence:
            messagebox.showerror("오류", "플로우가 비어있습니다.")
            return

        settings = self.project_data.get("settings", {}).get("execution", {})
        if settings.get("mode") == "excel_loop" and not self.excel_mgr.excel_sources:
            messagebox.showerror(
                "오류", "엑셀 행 반복 모드는 엑셀 데이터가 필요합니다."
            )
            return

        self.is_running = True
        self.add_log("=" * 50)
        self.add_log("매크로 실행 준비 중...")

        # 3초 카운트다운
        self.countdown(3)

    def countdown(self, count):
        """카운트다운"""
        if count > 0:
            self.add_log(f"시작까지 {count}초...")
            self.after(1000, self.countdown, count - 1)
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
        messagebox.showinfo("완료", "매크로 실행이 완료되었습니다!")

    def pause_macro(self):
        """매크로 일시정지/재개"""
        if self.executor.is_paused:
            self.executor.resume()
            self.pause_btn.configure(text="|| 일시정지")
        else:
            self.executor.pause()
            self.pause_btn.configure(text="> 재개")

    def stop_macro(self):
        """매크로 중지 (확인 없이 즉시 중지)"""
        self.executor.stop()
        self.add_log("[STOP] 사용자가 중지를 요청했습니다.")

    def bring_to_front(self):
        """프로그램 창을 맨 앞으로 가져오기"""
        try:
            # 창을 최소화 상태에서 복원
            self.parent.state("normal")

            # 맨 앞으로 가져오기
            self.parent.lift()
            self.parent.focus_force()
            self.parent.attributes("-topmost", True)
            self.parent.after(100, lambda: self.parent.attributes("-topmost", False))

            self.add_log("[INFO] 프로그램이 맨 앞으로 이동했습니다")
        except Exception as e:
            print(f"맨 앞으로 가져오기 오류: {e}")

    def edit_project(self):
        """프로젝트 편집"""
        self.stop_hotkey_listener()

        from ui.project_editor import ProjectEditor

        for widget in self.parent.winfo_children():
            widget.destroy()

        editor = ProjectEditor(self.parent, self.app, self.project_data, self.filepath)
        editor.pack(fill="both", expand=True)

    def show_excel_config(self, source):
        """특정 엑셀 소스의 자동 최신 파일 및 데이터 처리 옵션 설정"""
        from ui.auto_latest_dialog import AutoLatestFileDialog
        from core.project_manager import ProjectManager
        import os

        # 현재 설정값 가져오기
        current_dir = source.get("auto_directory", "")
        if not current_dir and source.get("filepath"):
            # 기존 파일 경로에서 디렉토리 추출
            current_dir = os.path.dirname(source.get("filepath", ""))

        # 다이얼로그 생성 (source를 전달하여 데이터 처리 옵션 표시)
        dialog = AutoLatestFileDialog(
            self.parent, default_directory=current_dir, source=source
        )

        # 현재 설정값으로 초기화
        if source.get("auto_latest", False):
            dialog.auto_var.set(True)
            dialog.toggle_inputs()
            dialog.dir_entry.delete(0, "end")
            dialog.dir_entry.insert(0, source.get("auto_directory", ""))
            dialog.prefix_entry.delete(0, "end")
            dialog.prefix_entry.insert(0, source.get("auto_prefix", ""))
            dialog.mode_var.set(source.get("auto_mode", "modified"))

        # 다이얼로그 대기
        self.wait_window(dialog)

        # 취소하지 않았으면 설정 저장
        if not dialog.cancelled:
            # 소스 데이터 업데이트 (AutoLatestFileDialog에서 이미 source 객체를 수정함)
            source["auto_latest"] = dialog.auto_latest
            source["auto_directory"] = dialog.auto_directory
            source["auto_prefix"] = dialog.auto_prefix
            source["auto_mode"] = dialog.auto_mode

            # project_data의 excel_sources에서 해당 소스 업데이트
            excel_sources = self.project_data.get("excel_sources", [])
            for i, src in enumerate(excel_sources):
                if src.get("id") == source.get("id"):
                    excel_sources[i] = source
                    break

            self.project_data["excel_sources"] = excel_sources

            # 프로젝트 파일 저장
            ProjectManager.save_project(self.filepath, self.project_data)

            # ExcelManager 갱신
            self.excel_mgr.load_from_list(excel_sources)

            # UI 전체 갱신
            for widget in self.winfo_children():
                widget.destroy()
            self.setup_ui()

            messagebox.showinfo(
                "완료",
                f"'{source.get('name', '엑셀')}' 설정이 저장되었습니다.",
                parent=self.parent,
            )

    def show_settings(self):
        """설정 창"""
        dialog = ctk.CTkToplevel(self.parent)
        dialog.title("실행 설정")
        dialog.geometry("400x300")
        dialog.transient(self.parent)
        dialog.grab_set()
        dialog.attributes("-topmost", True)
        dialog.configure(fg_color=Colors.BG_PRIMARY)

        set_dialog_icon(dialog)
        center_window_on_parent(dialog, self.parent)

        # ─── 다이얼로그 헤더 ───
        dialog_header = ctk.CTkFrame(dialog, fg_color="transparent")
        dialog_header.pack(fill="x", padx=Sizes.PAD_XL, pady=(Sizes.PAD_XL, Sizes.PAD_LG))

        ctk.CTkLabel(
            dialog_header,
            text="실행 설정",
            font=Fonts.HEADING,
            text_color=Colors.TEXT_PRIMARY,
        ).pack(anchor="w")

        ctk.CTkLabel(
            dialog_header,
            text="매크로 실행 옵션을 설정합니다",
            font=Fonts.CAPTION,
            text_color=Colors.TEXT_MUTED,
        ).pack(anchor="w", pady=(Sizes.PAD_XS, 0))

        # ─── 실행 설정 카드 ───
        mode_card = Styles.card(dialog, border_color=Colors.BORDER_SUBTLE)
        mode_card.pack(fill="x", padx=Sizes.PAD_XL, pady=(0, Sizes.PAD_LG))

        mode_inner = ctk.CTkFrame(mode_card, fg_color="transparent")
        mode_inner.pack(fill="x", padx=Sizes.PAD_LG, pady=Sizes.PAD_MD)

        Styles.section_title(mode_inner, text="반복 실행").pack(anchor="w", pady=(0, Sizes.PAD_SM))

        # 반복 횟수
        repeat_frame = ctk.CTkFrame(mode_inner, fg_color="transparent")
        repeat_frame.pack(fill="x", pady=(0, Sizes.PAD_SM))

        repeat_title_label = ctk.CTkLabel(
            repeat_frame,
            text="반복 횟수",
            font=Fonts.SMALL,
            text_color=Colors.TEXT_SECONDARY,
            width=80,
            anchor="w",
        )
        repeat_title_label.pack(side="left")

        # 기존 설정 읽기
        settings = self.project_data.get("settings", {}).get("execution", {})
        mode = settings.get("mode", "flow_repeat")
        repeat_count = settings.get("repeat_count", 1)
        is_infinite = mode == "infinite"

        repeat_entry = Styles.input_field(repeat_frame, width=80)
        repeat_entry.insert(0, str(repeat_count))
        repeat_entry.pack(side="left", padx=Sizes.PAD_SM)

        repeat_label = ctk.CTkLabel(
            repeat_frame, text="회", font=Fonts.SMALL, text_color=Colors.TEXT_SECONDARY
        )
        repeat_label.pack(side="left")

        # 무한 반복 체크박스
        infinite_var = ctk.BooleanVar(value=is_infinite)

        def toggle_repeat_entry():
            """무한 반복 체크 시 반복 횟수 입력 비활성화"""
            if infinite_var.get():
                repeat_title_label.configure(text_color=Colors.TEXT_MUTED)
                repeat_entry.configure(state="disabled")
                repeat_label.configure(text_color=Colors.TEXT_MUTED)
            else:
                repeat_title_label.configure(text_color=Colors.TEXT_SECONDARY)
                repeat_entry.configure(state="normal")
                repeat_label.configure(text_color=Colors.TEXT_SECONDARY)

        infinite_check = ctk.CTkCheckBox(
            mode_inner,
            text="무한 반복",
            variable=infinite_var,
            font=Fonts.SMALL_BOLD,
            text_color=Colors.DANGER,
            fg_color=Colors.DANGER,
            hover_color=Colors.DANGER_HOVER,
            border_color=Colors.BORDER_SUBTLE,
            checkmark_color=Colors.TEXT_PRIMARY,
            corner_radius=Sizes.RADIUS_SM,
            command=toggle_repeat_entry,
        )
        infinite_check.pack(
            anchor="w", pady=(0, Sizes.PAD_XS)
        )

        # 초기 상태 설정
        toggle_repeat_entry()

        def save_settings():
            # 실행 설정 저장
            if "settings" not in self.project_data:
                self.project_data["settings"] = {}
            if "execution" not in self.project_data["settings"]:
                self.project_data["settings"]["execution"] = {}

            # 반복 횟수 저장
            try:
                rc = int(repeat_entry.get())
                if rc < 1:
                    rc = 1
                self.project_data["settings"]["execution"]["repeat_count"] = rc
            except (ValueError, TypeError):
                self.project_data["settings"]["execution"]["repeat_count"] = 1

            # 실행 모드 결정
            if infinite_var.get():
                # 무한 반복
                if self.excel_mgr.excel_sources:
                    # 엑셀이 있으면 엑셀 무한 반복
                    self.project_data["settings"]["execution"]["mode"] = "excel_loop"
                    self.project_data["settings"]["execution"][
                        "excel_infinite_loop"
                    ] = True
                else:
                    # 엑셀이 없으면 플로우 무한 반복
                    self.project_data["settings"]["execution"]["mode"] = "infinite"
                    self.project_data["settings"]["execution"][
                        "excel_infinite_loop"
                    ] = False
            else:
                # 횟수 지정 반복
                if self.excel_mgr.excel_sources:
                    # 엑셀이 있으면 엑셀 행 반복
                    self.project_data["settings"]["execution"]["mode"] = "excel_loop"
                    self.project_data["settings"]["execution"][
                        "excel_infinite_loop"
                    ] = False
                else:
                    # 엑셀이 없으면 플로우 반복
                    self.project_data["settings"]["execution"]["mode"] = "flow_repeat"
                    self.project_data["settings"]["execution"][
                        "excel_infinite_loop"
                    ] = False

            # 저장
            ProjectManager.save_project(self.filepath, self.project_data)

            # 단축키 리스너 재시작
            self.stop_hotkey_listener()
            self.setup_hotkeys()

            dialog.destroy()
            messagebox.showinfo("완료", "설정이 저장되었습니다!")

        # ─── 저장 버튼 ───
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(fill="x", padx=Sizes.PAD_XL, pady=(0, Sizes.PAD_XL))

        Styles.primary_button(
            btn_frame,
            text="저장",
            command=save_settings,
            fg_color=Colors.SUCCESS,
            hover_color=Colors.SUCCESS_HOVER,
            width=120,
        ).pack(anchor="e")

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
