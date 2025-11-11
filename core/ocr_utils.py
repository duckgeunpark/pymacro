"""
OCR 유틸리티 - 이미지에서 텍스트와 숫자 인식
"""
import pytesseract
from PIL import Image
import re
import cv2
import numpy as np
import time


class OCRUtils:
    """OCR 기능"""
    
    @staticmethod
    def extract_text(image_path):
        """이미지에서 텍스트 추출"""
        try:
            image = Image.open(image_path)
            text = pytesseract.image_to_string(image)
            return text.strip()
        except Exception as e:
            print(f"OCR 오류: {e}")
            return None
    
    @staticmethod
    def extract_region_text(image_data, region):
        """스크린의 특정 영역에서 텍스트 추출"""
        try:
            # numpy 배열로 변환
            image_array = np.frombuffer(image_data, dtype=np.uint8)
            image = cv2.imdecode(image_array, cv2.IMREAD_COLOR)
            
            # 영역 추출
            x, y, w, h = region
            roi = image[y:y+h, x:x+w]
            
            # 이미지 전처리
            gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
            _, thresh = cv2.threshold(gray, 127, 255, cv2.THRESH_BINARY)
            
            # 이미지 확대 (정확도 향상)
            thresh = cv2.resize(thresh, None, fx=2, fy=2, interpolation=cv2.INTER_LINEAR)
            
            # OCR 실행
            text = pytesseract.image_to_string(Image.fromarray(thresh))
            return text.strip()
        except Exception as e:
            print(f"영역 OCR 오류: {e}")
            return None
    
    @staticmethod
    def extract_time_from_text(text):
        """
        텍스트에서 MM:SS 또는 MM:SS/MM:SS 형식 추출
        예: "00:30/12:34" → (30, 754) [남은 시간, 전체 시간]
        """
        try:
            # MM:SS/MM:SS 형식 인식
            pattern = r'(\d{1,2}):(\d{2})\s*/\s*(\d{1,2}):(\d{2})'
            match = re.search(pattern, text)
            
            if match:
                current_min, current_sec, total_min, total_sec = match.groups()
                current_seconds = int(current_min) * 60 + int(current_sec)
                total_seconds = int(total_min) * 60 + int(total_sec)
                remaining_seconds = total_seconds - current_seconds
                
                return {
                    'current': current_seconds,
                    'total': total_seconds,
                    'remaining': max(remaining_seconds, 0)  # 음수 방지
                }
            
            # 단순 MM:SS 형식
            pattern = r'(\d{1,2}):(\d{2})'
            match = re.search(pattern, text)
            
            if match:
                minutes, seconds = match.groups()
                total = int(minutes) * 60 + int(seconds)
                return {
                    'current': 0,
                    'total': total,
                    'remaining': total
                }
            
            return None
        except Exception as e:
            print(f"시간 추출 오류: {e}")
            return None
