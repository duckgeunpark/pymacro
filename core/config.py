"""
프로젝트 설정 관리 모듈
"""
import os
import sys


class AppConfig:
    """애플리케이션 설정 관리 클래스"""

    _instance = None
    _icon_path = None
    _app_path = None
    _initialized = False

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def initialize(self):
        """설정 초기화 (앱 시작 시 한 번만 호출)"""
        if self._initialized:
            return

        self._app_path = self._get_application_path()
        self._icon_path = self._find_icon_path()
        self._initialized = True

    @property
    def icon_path(self):
        """아이콘 파일 경로"""
        if not self._initialized:
            self.initialize()
        return self._icon_path

    @property
    def app_path(self):
        """애플리케이션 경로"""
        if not self._initialized:
            self.initialize()
        return self._app_path

    def _get_application_path(self):
        """애플리케이션 경로 반환"""
        try:
            if getattr(sys, 'frozen', False):
                # exe로 실행 중
                return os.path.dirname(sys.executable)
            else:
                # 스크립트로 실행 중
                return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        except:
            return os.getcwd()

    def _find_icon_path(self):
        """아이콘 파일 경로 찾기"""
        try:
            icon_path = os.path.join(self._app_path, 'resources', 'icon.ico')
            if os.path.exists(icon_path):
                return icon_path
            return None
        except:
            return None

    def create_directories(self):
        """필요한 디렉토리 생성"""
        directories = [
            'projects',
            'projects/images',
            'projects/excel',
            'projects/logs',
            'projects/logs/screenshots'
        ]

        for directory in directories:
            os.makedirs(directory, exist_ok=True)


# 전역 설정 인스턴스
config = AppConfig()