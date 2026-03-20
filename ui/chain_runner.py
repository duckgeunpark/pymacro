"""
매크로 체인 실행 화면
"""
import customtkinter as ctk
from tkinter import messagebox
import threading
from pynput import keyboard
from core.chain_executor import ChainExecutor
from ui.theme import Colors, Fonts, Sizes, Styles


class ChainRunner(ctk.CTkFrame):
    """매크로 체인 실행 화면"""

    def __init__(self, parent, app, chain_items, autostart=False):
        super().__init__(parent, fg_color=Colors.BG_PRIMARY)
        self.app = app
        self.parent = parent
        self.chain_items = chain_items  # [{'project_name': str, 'filepath': str, 'repeat_count': int}]
        self.autostart = autostart

        # 실행 엔진
        self.executor = ChainExecutor(chain_items)
        self.executor.set_callbacks(
            log_cb=self.add_log,
            # progress_cb=self.update_progress,
            error_cb=self.report_error
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
            if not key:
                continue

            key = key.strip().lower()

            # F키 변환
            if key.startswith('f') and len(key) > 1 and key[1:].isdigit():
                key_obj = getattr(keyboard.Key, key, None)
            elif key in ['enter', 'tab', 'esc', 'space', 'backspace', 'delete']:
                key_obj = getattr(keyboard.Key, key, None)
            else:
                key_obj = key

            if key_obj:
                self.hotkey_map[action] = key_obj

        # 리스너 시작
        self.start_hotkey_listener()

    def start_hotkey_listener(self):
        """단축키 리스너 시작"""
        def on_press(key):
            try:
                if key == self.hotkey_map.get('start'):
                    self.parent.after(0, self.start_chain)
                elif key == self.hotkey_map.get('pause'):
                    self.parent.after(0, self.pause_chain)
                elif key == self.hotkey_map.get('stop'):
                    self.parent.after(0, self.stop_chain)
                elif key == self.hotkey_map.get('focus'):
                    self.parent.after(0, self.bring_to_front)
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
        # ── 상단 헤더 ──
        header = Styles.card(self, corner_radius=0, border_width=0, height=Sizes.HEADER_HEIGHT)
        header.pack(fill='x')
        header.pack_propagate(False)

        ctk.CTkLabel(
            header,
            text="매크로 체인 실행",
            font=Fonts.HEADING,
            text_color=Colors.TEXT_PRIMARY,
        ).pack(side='left', padx=Sizes.PAD_XL, pady=Sizes.PAD_MD)

        # 뒤로가기 버튼 (ghost)
        Styles.ghost_button(
            header,
            text="← 뒤로",
            command=self.go_back,
        ).pack(side='right', padx=Sizes.PAD_LG, pady=Sizes.PAD_MD)

        # ── 메인 컨텐츠 ──
        main_frame = ctk.CTkFrame(self, fg_color="transparent", corner_radius=0)
        main_frame.pack(fill='both', expand=True, padx=Sizes.PAD_LG, pady=Sizes.PAD_MD)

        # ── 상측: 체인 정보 카드 ──
        top_card = Styles.card(main_frame, height=300)
        top_card.pack(fill='x', pady=(0, Sizes.PAD_SM))
        top_card.pack_propagate(False)

        self.setup_chain_info(top_card)

        # ── 하측: 로그 카드 ──
        bottom_card = Styles.card(main_frame)
        bottom_card.pack(fill='both', expand=True, pady=(Sizes.PAD_SM, 0))

        self.setup_log_and_controls(bottom_card)

        # 자동시작인 경우 바로 실행
        if self.autostart:
            self.after(500, self.start_chain)

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
                from core.project_manager import ProjectManager
                filepath = item.get('filepath', '')
                if filepath:
                    project_data = ProjectManager.load_project(filepath)
                    if project_data:
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

    def setup_chain_info(self, parent):
        """체인 정보 패널"""
        # ── 단축키 카드 ──
        hotkey_container = ctk.CTkFrame(parent, fg_color="transparent")
        hotkey_container.pack(fill='x', padx=Sizes.PAD_LG, pady=(Sizes.PAD_MD, Sizes.PAD_XS))

        Styles.section_title(hotkey_container, text="단축키").pack(anchor='w', pady=(0, Sizes.PAD_XS))

        hotkey_card = ctk.CTkFrame(
            hotkey_container,
            fg_color=Colors.BG_INPUT,
            corner_radius=Sizes.RADIUS_SM,
            border_width=1,
            border_color=Colors.BORDER,
        )
        hotkey_card.pack(fill='x')

        from core.settings_manager import SettingsManager
        hotkeys = SettingsManager.get_hotkeys()
        hotkey_text = (
            f"시작: {hotkeys.get('start', 'F9').upper()}  |  "
            f"일시정지: {hotkeys.get('pause', 'F10').upper()}  |  "
            f"중지: {hotkeys.get('stop', 'F11').upper()}  |  "
            f"맨 앞으로: {hotkeys.get('focus', 'F12').upper()}"
        )

        ctk.CTkLabel(
            hotkey_card,
            text=hotkey_text,
            font=Fonts.MONO_SMALL,
            text_color=Colors.TEXT_SECONDARY,
        ).pack(padx=Sizes.PAD_MD, pady=Sizes.PAD_SM)

        # ── 실행 순서 ──
        Styles.section_title(parent, text="실행 순서").pack(
            anchor='w', padx=Sizes.PAD_LG, pady=(Sizes.PAD_MD, Sizes.PAD_XS)
        )

        # 스크롤 가능한 리스트
        list_frame = ctk.CTkScrollableFrame(parent, fg_color="transparent", corner_radius=0)
        list_frame.pack(fill='both', expand=True, padx=Sizes.PAD_MD)

        # 체인 항목 표시 (grid layout)
        for idx, item in enumerate(self.chain_items):
            row = idx // 2
            col = idx % 2

            item_frame = ctk.CTkFrame(
                list_frame,
                fg_color=Colors.BG_INPUT,
                corner_radius=Sizes.RADIUS_SM,
                border_width=1,
                border_color=Colors.BORDER,
            )
            item_frame.grid(row=row, column=col, padx=Sizes.PAD_XS, pady=Sizes.PAD_XS, sticky='nsew')

            # 번호 배지
            badge = ctk.CTkLabel(
                item_frame,
                text=f" {idx + 1} ",
                font=Fonts.CAPTION_BOLD,
                text_color=Colors.TEXT_INVERSE,
                fg_color=Colors.PRIMARY,
                corner_radius=Sizes.RADIUS_SM,
                width=26,
                height=26,
            )
            badge.pack(side='left', padx=Sizes.PAD_SM, pady=Sizes.PAD_SM)

            info = ctk.CTkFrame(item_frame, fg_color="transparent")
            info.pack(side='left', fill='x', expand=True, padx=Sizes.PAD_SM, pady=Sizes.PAD_SM)

            ctk.CTkLabel(
                info,
                text=item['project_name'],
                font=Fonts.SMALL_BOLD,
                text_color=Colors.TEXT_PRIMARY,
                anchor='w',
            ).pack(anchor='w')

            # 실행 설정 표시 (실제 프로젝트 설정 읽기)
            mode_text = self.get_execution_mode_text(item)

            ctk.CTkLabel(
                info,
                text=mode_text,
                font=Fonts.CAPTION,
                text_color=Colors.PRIMARY_LIGHT,
                anchor='w',
            ).pack(anchor='w')

        for col in range(2):
            list_frame.grid_columnconfigure(col, weight=1, minsize=230, uniform="chain")

    def setup_log_and_controls(self, parent):
        """로그 및 제어 패널"""
        log_container = ctk.CTkFrame(parent, fg_color="transparent")
        log_container.pack(fill='both', expand=True, padx=Sizes.PAD_MD, pady=Sizes.PAD_MD)

        Styles.section_title(log_container, text="실행 로그").pack(anchor='w', pady=(0, Sizes.PAD_XS))

        self.log_text = ctk.CTkTextbox(
            log_container,
            font=Fonts.MONO_SMALL,
            fg_color=Colors.BG_INPUT,
            text_color=Colors.TEXT_PRIMARY,
            wrap='word',
            corner_radius=Sizes.RADIUS_SM,
            border_width=1,
            border_color=Colors.BORDER,
            state='disabled',
        )
        self.log_text.pack(fill='both', expand=True)

    def add_log(self, message):
        """로그 추가"""
        self.log_text.configure(state='normal')
        self.log_text.insert('end', message + '\n')
        self.log_text.see('end')
        self.log_text.configure(state='disabled')

    # def update_progress(self, current, total, status=""):
    #     """진행 상황 업데이트"""
    #     if total > 0:
    #         progress = (current / total) * 100
    #         self.progress_bar['value'] = progress
    #         self.progress_label.config(text=f"{status} ({current}/{total})")
    #     else:
    #         self.progress_label.config(text=status)

    def report_error(self, error_msg):
        """에러 보고"""
        self.add_log(f"[ERR] 에러: {error_msg}")

    def start_chain(self):
        """체인 실행 시작"""
        if self.is_running:
            return

        self.is_running = True

        # 별도 스레드에서 실행
        thread = threading.Thread(target=self.executor.start, daemon=True)
        thread.start()

    def pause_chain(self):
        """일시정지/재개"""
        if not self.is_running:
            return

        if self.executor.is_paused:
            self.executor.resume()
            self.add_log("[PLAY] 재개")
        else:
            self.executor.pause()
            self.add_log("[PAUSE] 일시정지")

    def stop_chain(self):
        """중지"""
        if not self.is_running:
            return

        self.executor.stop()
        self.is_running = False
        self.add_log("[STOP] 중지")

    def bring_to_front(self):
        """프로그램 창을 맨 앞으로 가져오기"""
        try:
            root = self.winfo_toplevel()
            if root.state() == 'iconic':
                root.state('normal')

            root.lift()
            root.focus_force()
            root.attributes('-topmost', True)
            root.after(100, lambda: root.attributes('-topmost', False))

            self.add_log("[INFO] 프로그램이 맨 앞으로 이동했습니다")
        except Exception as e:
            print(f"[WARNING] 맨 앞으로 가져오기 오류: {e}")

    def go_back(self):
        """뒤로가기"""
        if self.is_running:
            if not messagebox.askyesno("확인", "실행 중입니다. 중지하고 돌아가시겠습니까?"):
                return
            self.executor.stop()

        # 리스너 중지
        self.stop_hotkey_listener()

        # 시작 화면으로 돌아가기
        self.app.show_start_screen()
