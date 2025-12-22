"""
화면 메크로 빌더 - 메인 실행 파일
"""
import tkinter as tk
from tkinter import messagebox
import sys
import os

# 프로젝트 내부 모듈
from core.config import config
from ui.start_screen import StartScreen


class MacroBuilderApp:
    """메인 애플리케이션 클래스"""

    def __init__(self):
        # 설정 초기화
        config.initialize()

        # 메인 윈도우 설정
        self.root = tk.Tk()
        self.root.title("dMax MacroBuilder")
        self.root.geometry("550x670")
        self.root.minsize(550, 670)

        # 윈도우 아이콘 설정
        self.set_window_icon()

        # 이벤트 바인딩
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.bind('<F12>', lambda e: self.bring_to_front())

        # 최초 실행 플래그
        self.is_initial_launch = True

        # 시작 화면 표시
        self.show_start_screen()

    def set_window_icon(self):
        """윈도우 아이콘 설정"""
        try:
            icon_path = config.icon_path
            if icon_path and os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
                print(f"[OK] 아이콘 설정: {icon_path}")
            else:
                print("[WARNING] 아이콘 파일 없음")
        except Exception as e:
            print(f"[WARNING] 아이콘 설정 실패: {e}")
        
    def show_start_screen(self):
        """시작 화면 표시"""
        # 기존 위젯 제거
        for widget in self.root.winfo_children():
            widget.destroy()

        # 최초 실행 시에만 자동시작 확인
        autostart_executed = False
        if self.is_initial_launch:
            autostart_executed = self.check_autostart()
            # 최초 실행 플래그 해제
            self.is_initial_launch = False

        # 자동시작이 실행되지 않은 경우에만 StartScreen 표시
        if not autostart_executed:
            start_screen = StartScreen(self.root, self)
            start_screen.pack(fill='both', expand=True)
    
    def check_autostart(self):
        """자동시작 프로젝트/체인 확인 및 실행

        Returns:
            bool: 자동시작이 실행되었으면 True, 아니면 False
        """
        import json
        from tkinter import messagebox

        try:
            # settings.json 읽기
            with open('settings.json', 'r', encoding='utf-8') as f:
                settings = json.load(f)

            autostart_path = settings.get('autostart')

            # 자동시작 설정이 비어있거나 None인 경우
            if not autostart_path:
                return False

            # 파일이 존재하지 않는 경우
            if not os.path.exists(autostart_path):
                print(f"[WARNING] 자동시작 파일이 존재하지 않습니다: {autostart_path}")

                # 사용자에게 알림 (after를 사용하여 UI 생성 후 표시)
                self.root.after(100, lambda: messagebox.showwarning(
                    "자동시작 실패",
                    "자동시작으로 설정된 프로젝트/체인을 찾을 수 없습니다.\n자동시작 설정이 해제되었습니다."
                ))

                # settings.json의 autostart 값 비우기
                settings['autostart'] = None
                with open('settings.json', 'w', encoding='utf-8') as f:
                    json.dump(settings, f, ensure_ascii=False, indent=2)
                return False

            # 파일이 존재하는 경우
            print(f"[INFO] 자동시작 실행: {autostart_path}")

            from core.project_manager import ProjectManager

            # 파일 로드
            project_data = ProjectManager.load_project(autostart_path)
            if not project_data:
                print(f"[WARNING] 자동시작 파일 로드 실패: {autostart_path}")

                # 사용자에게 알림 (after를 사용하여 UI 생성 후 표시)
                self.root.after(100, lambda: messagebox.showerror(
                    "자동시작 실패",
                    "자동시작 파일을 불러올 수 없습니다.\n자동시작 설정이 해제되었습니다."
                ))

                # settings.json의 autostart 값 비우기
                settings['autostart'] = None
                with open('settings.json', 'w', encoding='utf-8') as f:
                    json.dump(settings, f, ensure_ascii=False, indent=2)
                return False

            # 타입 확인 (프로젝트 또는 체인)
            item_type = project_data.get('type', 'project')

            if item_type == 'chain':
                # 체인 실행
                from ui.chain_runner import ChainRunner

                chain_items = project_data.get('chain_items', [])
                runner = ChainRunner(self.root, self, chain_items, autostart=True)
                runner.pack(fill='both', expand=True)
            else:
                # 프로젝트 실행
                from ui.project_runner import ProjectRunner

                runner = ProjectRunner(self.root, self, project_data, autostart_path, autostart=True)
                runner.pack(fill='both', expand=True)

            return True  # 자동시작 실행됨

        except Exception as e:
            print(f"[WARNING] 자동시작 확인 중 오류: {e}")
            return False

    def bring_to_front(self):
        """프로그램 창을 맨 앞으로 가져오기"""
        try:
            # 최소화 상태면 복원
            if self.root.state() == 'iconic':
                self.root.state('normal')

            # 맨 앞으로 가져오기
            self.root.lift()
            self.root.focus_force()

            # 잠깐 topmost 설정 후 해제 (확실하게 앞으로)
            self.root.attributes('-topmost', True)
            self.root.after(100, lambda: self.root.attributes('-topmost', False))

            print("[INFO] 프로그램이 맨 앞으로 이동했습니다")
        except Exception as e:
            print(f"[WARNING] 맨 앞으로 가져오기 오류: {e}")
    
    def on_closing(self):
        """프로그램 종료 확인"""
        if messagebox.askokcancel("종료", "프로그램을 종료하시겠습니까?"):
            self.root.destroy()
            sys.exit(0)
    
    def run(self):
        """프로그램 실행"""
        self.root.mainloop()


def main():
    """메인 함수"""
    # 설정 초기화 및 디렉토리 생성
    config.initialize()
    os.chdir(config.app_path)
    config.create_directories()
    config.create_settings_file()  # settings.json 파일 생성

    # 앱 실행
    app = MacroBuilderApp()
    app.run()


if __name__ == "__main__":
    main()
