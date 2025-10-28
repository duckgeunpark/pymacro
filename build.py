"""
PyInstaller를 사용한 exe 빌드 스크립트
"""
import os
import subprocess
import shutil

def build_exe():
    """exe 파일 빌드"""
    print("="*50)
    print("매크로 빌더 exe 빌드 시작")
    print("="*50)
    
    # PyInstaller 명령
    cmd = [
            'pyinstaller',
            '--onefile',
            '--windowed',
            '--name=MacroBuilder',
            '--icon=resources/icon.ico',              # EXE 파일 아이콘
            '--add-data=resources;resources',         # resources 폴더 포함 ⭐
            '--add-data=ui;ui',
            '--add-data=core;core',
            '--hidden-import=PIL._tkinter_finder',
            '--hidden-import=openpyxl',
            '--hidden-import=pandas',
            '--hidden-import=pyautogui',
            '--hidden-import=pynput',
            '--hidden-import=pyperclip',
            'main.py'
    ]
    
    # 실행
    try:
        subprocess.run(cmd, check=True)
        print("\n" + "="*50)
        print("빌드 완료!")
        print("실행 파일: dist/MacroBuilder.exe")
        print("="*50)
        
        # 필요한 폴더를 dist에 복사
        print("\n필요한 폴더 복사 중...")
        dist_path = 'dist'
        
        # 빈 프로젝트 폴더 생성
        os.makedirs(os.path.join(dist_path, 'projects'), exist_ok=True)
        os.makedirs(os.path.join(dist_path, 'projects', 'images'), exist_ok=True)
        os.makedirs(os.path.join(dist_path, 'projects', 'excel'), exist_ok=True)
        os.makedirs(os.path.join(dist_path, 'logs'), exist_ok=True)
        os.makedirs(os.path.join(dist_path, 'logs', 'screenshots'), exist_ok=True)
        
        print("완료!")
        print(f"\n배포 파일 위치: {os.path.abspath(dist_path)}")
        
    except subprocess.CalledProcessError as e:
        print(f"\n빌드 오류: {e}")
        return False
    
    return True

if __name__ == "__main__":
    build_exe()
