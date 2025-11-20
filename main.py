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
        
        # 시작 화면 로드
        start_screen = StartScreen(self.root, self)
        start_screen.pack(fill='both', expand=True)
    
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

    # 앱 실행
    app = MacroBuilderApp()
    app.run()


if __name__ == "__main__":
    main()
