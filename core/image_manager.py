"""
이미지 템플릿 관리 (화면 인식용)
"""
import pyautogui
from PIL import Image, ImageGrab
import io
import base64
import os


class ImageManager:
    """이미지 템플릿 관리 클래스"""
    
    def __init__(self):
        self.images = []
        self.next_id = 1
        
        # images 폴더 생성
        self.ensure_images_folder()
    
    def ensure_images_folder(self):
        """images 폴더 확인 및 생성"""
        images_dir = os.path.join('projects', 'images')
        os.makedirs(images_dir, exist_ok=True)
    
    def add_image(self, name, image_data, confidence=0.8, description=""):
        """이미지 추가
        
        Args:
            name: 이미지 이름
            image_data: base64 encoded image string
            confidence: 매칭 정확도 (0.0 ~ 1.0)
            description: 설명
        """
        # 안전한 파일명 생성
        safe_name = "".join(c for c in name if c.isalnum() or c in (' ', '-', '_')).strip()
        safe_name = safe_name.replace(' ', '_')
        
        # 이미지 파일 저장
        image_filename = f"img_{self.next_id}_{safe_name}.png"
        image_path = os.path.join('projects', 'images', image_filename)
        
        try:
            # data URI 형식인 경우 처리
            if isinstance(image_data, str) and image_data.startswith('data:image'):
                image_data = image_data.split(',')[1]
            
            # base64 디코딩 및 이미지 저장
            img_bytes = base64.b64decode(image_data)
            img = Image.open(io.BytesIO(img_bytes))
            
            # PNG 형식으로 저장
            img.save(image_path, 'PNG')
            
            # 성공 로그
            print(f"✅ 이미지 저장 완료: {image_path}")
            
        except Exception as e:
            print(f"❌ 이미지 저장 오류: {e}")
            return None
        
        # 이미지 객체 생성 (data 필드에 base64 저장)
        image_obj = {
            'id': self.next_id,
            'name': name,
            'filename': image_filename,
            'path': image_path,
            'data': image_data,  # base64 데이터 저장 (executor에서 사용)
            'confidence': confidence,
            'description': description
        }
        
        self.images.append(image_obj)
        self.next_id += 1
        return image_obj
    
    def remove_image(self, image_id):
        """이미지 삭제"""
        image = self.get_image(image_id)
        if image:
            # 파일 삭제
            try:
                if os.path.exists(image['path']):
                    os.remove(image['path'])
                    print(f"🗑️ 이미지 파일 삭제: {image['path']}")
            except Exception as e:
                print(f"⚠️ 이미지 파일 삭제 실패: {e}")
            
            self.images = [img for img in self.images if img['id'] != image_id]
    
    def get_image(self, image_id):
        """ID로 이미지 찾기"""
        for img in self.images:
            if img['id'] == image_id:
                return img
        return None
    
    def update_image(self, image_id, **kwargs):
        """이미지 업데이트"""
        image = self.get_image(image_id)
        if image:
            # confidence 등의 속성 업데이트
            for key, value in kwargs.items():
                if key in ['name', 'confidence', 'description']:
                    image[key] = value
            return True
        return False
    
    @staticmethod
    def capture_region(x1, y1, x2, y2):
        """화면 영역 캡처"""
        try:
            screenshot = ImageGrab.grab(bbox=(x1, y1, x2, y2))
            # base64로 변환
            buffered = io.BytesIO()
            screenshot.save(buffered, format="PNG")
            img_str = base64.b64encode(buffered.getvalue()).decode()
            return img_str
        except Exception as e:
            print(f"❌ 화면 캡처 오류: {e}")
            return None
    
    @staticmethod
    def find_image_on_screen(image_path, confidence=0.8):
        """화면에서 이미지 찾기 (파일 경로 사용)"""
        try:
            location = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if location:
                center = pyautogui.center(location)
                return center.x, center.y
            return None
        except Exception as e:
            print(f"❌ 이미지 찾기 오류: {e}")
            return None
    
    @staticmethod
    def find_image_from_data(image_data_b64, confidence=0.8):
        """화면에서 이미지 찾기 (base64 데이터 사용)"""
        import tempfile
        
        try:
            # base64를 임시 파일로 저장
            img_bytes = base64.b64decode(image_data_b64)
            
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_file:
                tmp_file.write(img_bytes)
                tmp_path = tmp_file.name
            
            # 이미지 찾기
            location = pyautogui.locateOnScreen(tmp_path, confidence=confidence)
            
            # 임시 파일 삭제
            os.unlink(tmp_path)
            
            if location:
                center = pyautogui.center(location)
                return center.x, center.y
            return None
            
        except Exception as e:
            print(f"❌ 이미지 찾기 오류: {e}")
            return None
    
    def load_from_list(self, image_list):
        """리스트에서 이미지 로드"""
        self.images = []
        
        for img_data in image_list:
            # 파일 경로 확인 (하위 호환성)
            if 'path' in img_data and os.path.exists(img_data['path']):
                # 파일이 존재하면 base64로 변환하여 저장
                if 'data' not in img_data:
                    try:
                        with open(img_data['path'], 'rb') as f:
                            img_bytes = f.read()
                            img_data['data'] = base64.b64encode(img_bytes).decode()
                    except Exception as e:
                        print(f"⚠️ 이미지 로드 실패: {img_data['path']} - {e}")
            
            self.images.append(img_data)
        
        if self.images:
            self.next_id = max(img['id'] for img in self.images) + 1
        else:
            self.next_id = 1
    
    def to_list(self):
        """리스트로 변환 (JSON 저장용)"""
        # data 필드 포함하여 반환 (base64 문자열)
        return self.images
