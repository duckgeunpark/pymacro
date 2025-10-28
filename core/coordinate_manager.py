"""
좌표 관리 모듈 (수정됨)
"""
import pyautogui
import time
from PIL import ImageGrab, Image
import io

class CoordinateManager:
    """좌표 관리 클래스"""
    
    def __init__(self):
        self.coordinates = []
        self.next_id = 1
    
    def add_coordinate(self, name, x, y, description="", thumbnail=None):
        """좌표 추가"""
        coord = {
            'id': self.next_id,
            'name': name,
            'x': x,
            'y': y,
            'description': description,
            'thumbnail': thumbnail
        }
        self.coordinates.append(coord)
        self.next_id += 1
        return coord
    
    def remove_coordinate(self, coord_id):
        """좌표 삭제"""
        self.coordinates = [c for c in self.coordinates if c['id'] != coord_id]
    
    def get_coordinate(self, coord_id):
        """ID로 좌표 찾기"""
        for coord in self.coordinates:
            if coord['id'] == coord_id:
                return coord
        return None
    
    def update_coordinate(self, coord_id, **kwargs):
        """좌표 업데이트"""
        coord = self.get_coordinate(coord_id)
        if coord:
            coord.update(kwargs)
            return True
        return False
    
    @staticmethod
    def capture_current_position():
        """현재 마우스 위치 캡처"""
        x, y = pyautogui.position()
        
        # 주변 영역 스크린샷 (50x50)
        try:
            screenshot = ImageGrab.grab(bbox=(x-25, y-25, x+25, y+25))
            # PIL Image를 base64로 변환
            import base64
            buffered = io.BytesIO()
            screenshot.save(buffered, format="PNG")
            img_str = base64.b64encode(buffered.getvalue()).decode()
            return x, y, img_str
        except Exception as e:
            print(f"썸네일 캡처 오류: {e}")
            return x, y, None
    
    def load_from_list(self, coord_list):
        """리스트에서 좌표 로드"""
        self.coordinates = coord_list
        if coord_list:
            self.next_id = max(c['id'] for c in coord_list) + 1
        else:
            self.next_id = 1
    
    def to_list(self):
        """리스트로 변환"""
        return self.coordinates
