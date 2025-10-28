"""
프로젝트 저장/로드 관리
"""
import json
import os
from datetime import datetime

class ProjectManager:
    """프로젝트 파일 관리 클래스"""
    
    @staticmethod
    def save_project(filepath, project_data):
        """프로젝트 저장"""
        try:
            # 수정 시간 업데이트
            project_data['modified_at'] = datetime.now().isoformat()
            
            # JSON 저장
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(project_data, f, ensure_ascii=False, indent=4)
            
            return True
        except Exception as e:
            print(f"프로젝트 저장 오류: {e}")
            return False
    
    @staticmethod
    def load_project(filepath):
        """프로젝트 로드"""
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                project_data = json.load(f)
            return project_data
        except Exception as e:
            print(f"프로젝트 로드 오류: {e}")
            return None
    
    @staticmethod
    def create_empty_project(name, description=""):
        """빈 프로젝트 생성"""
        return {
            'name': name,
            'description': description,
            'coordinates': [],
            'excel_sources': [],
            'images': [],
            'flow_sequence': [],
            'settings': {
                'hotkeys': {
                    'start': 'F9',
                    'pause': 'F10',
                    'stop': 'F11',
                    'focus': 'F12' 
                },
                'execution': {
                    'mode': 'flow_repeat',  # excel_loop, flow_repeat, infinite
                    'repeat_count': 1,
                    'excel_start_row': 1,
                    'excel_end_row': None,
                    'excel_infinite_loop': False,
                    'on_error': 'skip',  # skip, stop, retry
                    'retry_count': 3,
                    'speed': 'normal'  # slow, normal, fast
                }
            },
            'created_at': datetime.now().isoformat(),
            'modified_at': datetime.now().isoformat()
        }
