"""
매크로 체인 관리 클래스
"""
import json
import os
from datetime import datetime

class ChainManager:
    """매크로 체인 관리"""

    def __init__(self):
        self.chain_items = []  # [{'project_name': str, 'filepath': str, 'repeat_count': int}]

    def add_item(self, project_name, filepath, repeat_count):
        """체인에 항목 추가"""
        item = {
            'project_name': project_name,
            'filepath': filepath,
            'repeat_count': repeat_count
        }
        self.chain_items.append(item)
        return item

    def remove_item(self, index):
        """체인에서 항목 제거"""
        if 0 <= index < len(self.chain_items):
            del self.chain_items[index]
            return True
        return False

    def move_item_up(self, index):
        """항목을 위로 이동"""
        if index > 0 and index < len(self.chain_items):
            self.chain_items[index], self.chain_items[index-1] = \
                self.chain_items[index-1], self.chain_items[index]
            return True
        return False

    def move_item_down(self, index):
        """항목을 아래로 이동"""
        if index >= 0 and index < len(self.chain_items) - 1:
            self.chain_items[index], self.chain_items[index+1] = \
                self.chain_items[index+1], self.chain_items[index]
            return True
        return False

    def get_items(self):
        """체인 항목 리스트 반환"""
        return self.chain_items

    def clear(self):
        """체인 초기화"""
        self.chain_items = []

    def get_total_executions(self):
        """총 실행 횟수 계산"""
        return sum(item['repeat_count'] for item in self.chain_items)

    @staticmethod
    def save_chain(filepath, chain_data):
        """체인 저장"""
        try:
            # 수정 시간 업데이트
            chain_data['modified_at'] = datetime.now().isoformat()

            # JSON 저장
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(chain_data, f, ensure_ascii=False, indent=4)

            return True
        except Exception as e:
            print(f"체인 저장 오류: {e}")
            return False

    @staticmethod
    def load_chain(filepath):
        """체인 불러오기"""
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                chain_data = json.load(f)
            return chain_data
        except Exception as e:
            print(f"체인 로드 오류: {e}")
            return None

    @staticmethod
    def create_empty_chain(name, description=""):
        """빈 체인 생성"""
        return {
            'name': name,
            'description': description,
            'type': 'chain',  # 'project' 또는 'chain' 구분
            'chain_items': [],
            'created_at': datetime.now().isoformat(),
            'modified_at': datetime.now().isoformat()
        }

    @staticmethod
    def get_chain_list():
        """체인 목록 반환"""
        chains = []

        if not os.path.exists('projects'):
            return chains

        json_files = [f for f in os.listdir('projects') if f.endswith('.json') and not f.endswith('.hidden')]

        for filename in json_files:
            filepath = os.path.join('projects', filename)
            try:
                data = ChainManager.load_chain(filepath)
                if data and data.get('type') == 'chain':
                    chains.append({
                        'name': data.get('name', 'Unknown'),
                        'filepath': filepath,
                        'modified_at': data.get('modified_at', ''),
                        'type': 'chain'
                    })
            except Exception as e:
                print(f"체인 로드 오류 ({filename}): {e}")

        # 수정 시간 순 정렬 (최신순)
        chains.sort(key=lambda x: x.get('modified_at', ''), reverse=True)

        return chains
