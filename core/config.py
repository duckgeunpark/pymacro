"""
프로젝트 설정 관리 모듈
"""
import os
import sys
import json


class AppConfig:
    """애플리케이션 설정 관리 클래스"""

    _instance = None
    _icon_path = None
    _app_path = None
    _initialized = False

    # 앱 버전
    VERSION = "1.0.0"

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
                return os.path.dirname(sys.executable)
            else:
                return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        except (OSError, AttributeError):
            return os.getcwd()

    def _find_icon_path(self):
        """아이콘 파일 경로 찾기"""
        try:
            icon_path = os.path.join(self._app_path, 'resources', 'icon.ico')
            if os.path.exists(icon_path):
                return icon_path
            return None
        except (OSError, TypeError):
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

    def create_settings_file(self):
        """settings.json 파일 생성 (없는 경우)"""
        settings_path = 'settings.json'

        if not os.path.exists(settings_path):
            default_settings = {
                "hotkeys": {
                    "start": "f9",
                    "pause": "f10",
                    "stop": "f11",
                    "focus": "f12"
                },
                "autostart": None
            }

            try:
                with open(settings_path, 'w', encoding='utf-8') as f:
                    json.dump(default_settings, f, ensure_ascii=False, indent=2)
                print(f"[INFO] settings.json 파일이 생성되었습니다.")
            except (OSError, IOError) as e:
                print(f"[WARNING] settings.json 생성 실패: {e}")


# 전역 설정 인스턴스
config = AppConfig()
