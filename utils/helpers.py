"""
유틸리티 함수들
"""
import os
import sys
from datetime import datetime

def get_application_path():
    """애플리케이션 경로 반환"""
    if getattr(sys, 'frozen', False):
        # exe로 실행 중
        return os.path.dirname(sys.executable)
    else:
        # 스크립트로 실행 중
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def ensure_directories():
    """필요한 디렉토리 생성"""
    dirs = [
        'projects',
        'projects/images',
        'logs',
        'logs/screenshots'
    ]
    for d in dirs:
        os.makedirs(d, exist_ok=True)

def format_timestamp(dt=None):
    """타임스탬프 포맷"""
    if dt is None:
        dt = datetime.now()
    return dt.strftime('%Y-%m-%d %H:%M:%S')

def sanitize_filename(filename):
    """파일명에서 사용 불가능한 문자 제거"""
    invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    return filename
