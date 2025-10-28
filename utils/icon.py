from PIL import Image
import os

# resources 폴더 생성
os.makedirs('resources', exist_ok=True)

# 선택한 PNG 파일
input_file = 'logo.png'  # 파일명
output_file = 'resources/icon.ico'

try:
    # ICO 변환 (여러 크기 포함)
    img = Image.open(input_file)
    img.save(output_file, format='ICO', sizes=[
        (256, 256),
        (128, 128),
        (64, 64),
        (48, 48),
        (32, 32),
        (16, 16)
    ])
    
    print(f"✅ 아이콘 생성 완료: {output_file}")
    
except FileNotFoundError:
    print(f"❌ 오류: '{input_file}' 파일을 찾을 수 없습니다.")
    print(f"현재 폴더에 '{input_file}' 파일이 있는지 확인하세요.")
    
except Exception as e:
    print(f"❌ 오류 발생: {e}")
