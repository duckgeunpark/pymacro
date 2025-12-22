"""
전역 설정 관리자
"""
import json
import os


class SettingsManager:
    """전역 설정 관리"""

    SETTINGS_FILE = "settings.json"

    @staticmethod
    def get_settings_path():
        """설정 파일 경로 (현재 작업 디렉토리 기준)"""
        # main.py에서 os.chdir(config.app_path)를 호출하므로
        # 현재 작업 디렉토리에서 settings.json을 찾음
        return SettingsManager.SETTINGS_FILE

    @staticmethod
    def load_settings():
        """설정 로드"""
        filepath = SettingsManager.get_settings_path()
        if not os.path.exists(filepath):
            return SettingsManager.get_default_settings()

        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                settings = json.load(f)
                # 기본값과 병합
                default_settings = SettingsManager.get_default_settings()
                for key in default_settings:
                    if key not in settings:
                        settings[key] = default_settings[key]
                return settings
        except Exception as e:
            print(f"설정 로드 실패: {e}")
            return SettingsManager.get_default_settings()

    @staticmethod
    def save_settings(settings):
        """설정 저장"""
        filepath = SettingsManager.get_settings_path()
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"설정 저장 실패: {e}")
            return False

    @staticmethod
    def get_default_settings():
        """기본 설정"""
        return {
            'hotkeys': {
                'start': 'f9',
                'pause': 'f10',
                'stop': 'f11',
                'focus': 'f12'
            }
        }

    @staticmethod
    def get_hotkeys():
        """단축키 설정 가져오기"""
        settings = SettingsManager.load_settings()
        return settings.get('hotkeys', SettingsManager.get_default_settings()['hotkeys'])

    @staticmethod
    def save_hotkeys(hotkeys):
        """단축키 설정 저장"""
        settings = SettingsManager.load_settings()
        settings['hotkeys'] = hotkeys
        return SettingsManager.save_settings(settings)
