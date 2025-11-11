"""
좌표 관리 모듈 (수정됨) - OCR 영역 기능 추가
"""
import pyautogui
import time
from PIL import ImageGrab, Image
import io
import mss
import numpy as np


class CoordinateManager:
    """좌표 관리 클래스"""
    
    def __init__(self):
        self.coordinates = []
        self.ocr_regions = []  # ← 추가
        self.next_id = 1
        self.next_region_id = 1  # ← 추가
    
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
    
    # ===== OCR 영역 관리 (추가) =====
    
    def add_ocr_region(self, name, x, y, w, h, mode='remaining'):
        """OCR 영역 추가"""
        region = {
            'id': self.next_region_id,
            'name': name,
            'x': x,
            'y': y,
            'w': w,
            'h': h,
            'mode': mode,  # 'remaining' 또는 'total'
            'thumbnail': self._capture_region_thumbnail(x, y, w, h)
        }
        self.ocr_regions.append(region)
        self.next_region_id += 1
        return region
    
    def remove_ocr_region(self, region_id):
        """OCR 영역 삭제"""
        self.ocr_regions = [r for r in self.ocr_regions if r['id'] != region_id]
    
    def get_ocr_region(self, region_id):
        """ID로 OCR 영역 찾기"""
        for region in self.ocr_regions:
            if region['id'] == region_id:
                return region
        return None
    
    def update_ocr_region(self, region_id, **kwargs):
        """OCR 영역 업데이트"""
        region = self.get_ocr_region(region_id)
        if region:
            region.update(kwargs)
            return True
        return False
    
    @staticmethod
    def _capture_region_thumbnail(x, y, w, h):
        """영역 썸네일 캡처"""
        try:
            screenshot = ImageGrab.grab(bbox=(x, y, x+w, y+h))
            import base64
            buffered = io.BytesIO()
            screenshot.save(buffered, format="PNG")
            img_str = base64.b64encode(buffered.getvalue()).decode()
            return img_str
        except Exception as e:
            print(f"영역 썸네일 캡처 오류: {e}")
            return None

            
    @staticmethod
    def capture_ocr_region_screenshot():
        """
        안전하게 오버레이 Toplevel만 생성&파괴(메인 루트 건드리지 않음)
        """
        import tkinter as tk

        selected_region = [None]
        start = {'x': None, 'y': None}
        rect_id = [None]
        is_done = [False]

        # 이미 _default_root가 있다면 재활용, 없으면 새로 생성(자동 mainloop됨)
        root = tk._default_root if tk._default_root else tk.Tk()
        root.withdraw()

        overlay = tk.Toplevel(root)
        overlay.attributes('-fullscreen', True)
        overlay.attributes('-alpha', 0.3)
        overlay.attributes('-topmost', True)
        overlay.configure(bg='red')
        screen_w, screen_h = overlay.winfo_screenwidth(), overlay.winfo_screenheight()
        canvas = tk.Canvas(
            overlay, bg='red', width=screen_w, height=screen_h,
            highlightthickness=0, cursor='tcross'
        )
        canvas.pack(fill='both', expand=True)

        canvas.create_text(
            24, 20,
            text="드래그해서 영역 지정 | ESC:취소",
            anchor='nw', fill='#ffff96', font=("Arial", 12, "bold")
        )

        def on_press(e):
            start['x'] = e.x
            start['y'] = e.y
            if rect_id[0]:
                canvas.delete(rect_id[0])

        def on_drag(e):
            if start['x'] is not None and start['y'] is not None:
                if rect_id[0]:
                    canvas.delete(rect_id[0])
                rect_id[0] = canvas.create_rectangle(
                    start['x'], start['y'], e.x, e.y,
                    outline='#ffff00', width=3
                )

        def on_release(e):
            x1, y1 = start['x'], start['y']
            x2, y2 = e.x, e.y
            x, y, w, h = min(x1, x2), min(y1, y2), abs(x2 - x1), abs(y2 - y1)
            if w > 10 and h > 10:
                selected_region[0] = (x, y, w, h)
            is_done[0] = True
            overlay.destroy()

        def on_cancel(e=None):
            is_done[0] = True
            overlay.destroy()

        canvas.bind('<Button-1>', on_press)
        canvas.bind('<B1-Motion>', on_drag)
        canvas.bind('<ButtonRelease-1>', on_release)
        canvas.bind('<Escape>', on_cancel)
        overlay.bind('<Escape>', on_cancel)

        overlay.update()
        overlay.focus_force()
        overlay.grab_set()  # 이 윈도우가 마우스‧키보드 모두 가져감

        # 추가: 메인 root는 절대 destroy/quit 안함!
        while not is_done[0]:
            overlay.update_idletasks()
            overlay.update()

        # 주의! root.destroy() X, root.quit() X
        return selected_region[0]




    def capture_region_screenshot(self):
        """영역 선택 스크린샷"""
        return self.capture_ocr_region_screenshot()
    
    # ===== 저장/로드 =====
    
    def load_from_list(self, coord_list, region_list=None):
        """리스트에서 좌표와 OCR 영역 로드"""
        self.coordinates = coord_list
        self.ocr_regions = region_list if region_list else []
        
        if coord_list:
            self.next_id = max(c['id'] for c in coord_list) + 1
        else:
            self.next_id = 1
        
        if region_list:
            self.next_region_id = max(r['id'] for r in region_list) + 1
        else:
            self.next_region_id = 1
    
    def to_list(self):
        """리스트로 변환 (좌표)"""
        return self.coordinates
    
    def to_region_list(self):
        """OCR 영역 리스트로 변환"""
        return self.ocr_regions
