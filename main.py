"""
화면 메크로 빌더 - 메인 실행 파일
"""
import tkinter as tk
from tkinter import messagebox
import sys
import os


# 프로젝트 내부 모듈
from ui.start_screen import StartScreen


class MacroBuilderApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("DMAX 매크로 빌더")
        self.root.geometry("900x650")
        self.root.minsize(800, 600)
        
        # 프로그램 종료 시 확인
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # F12 단축키로 맨 앞으로 가져오기
        self.root.bind('<F12>', lambda e: self.bring_to_front())
        
        # 시작 화면 표시
        self.show_start_screen()
        
    def show_start_screen(self):
        """시작 화면 표시"""
        # 기존 위젯 제거
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # 시작 화면 로드
        start_screen = StartScreen(self.root, self)
        start_screen.pack(fill='both', expand=True)
    
    def bring_to_front(self):
        """프로그램 창을 맨 앞으로 가져오기 (F12)"""
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
            
            print("🔼 프로그램이 맨 앞으로 이동했습니다 (F12)")
        except Exception as e:
            print(f"⚠️ 맨 앞으로 가져오기 오류: {e}")
    
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
    # 작업 디렉토리 설정 (exe 실행 시 경로 문제 해결)
    if getattr(sys, 'frozen', False):
        # exe로 실행 중
        application_path = os.path.dirname(sys.executable)
    else:
        # 스크립트로 실행 중
        application_path = os.path.dirname(os.path.abspath(__file__))
    
    os.chdir(application_path)
    
    # 필요한 폴더 생성
    os.makedirs('projects', exist_ok=True)
    os.makedirs('projects/images', exist_ok=True)
    os.makedirs('projects/excel', exist_ok=True)
    os.makedirs('logs', exist_ok=True)
    os.makedirs('logs/screenshots', exist_ok=True)
    
    # 앱 실행
    app = MacroBuilderApp()
    app.run()


if __name__ == "__main__":
    main()
