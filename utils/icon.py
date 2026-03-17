from PIL import Image
import os

# 현재 작업 디렉토리와 파일 목록 확인
print(f"현재 작업 디렉토리: {os.getcwd()}")
print("현재 폴더 파일 목록:")
for f in os.listdir('.'):
    print(f"  - {f}")
    
input_file = '9461e7c2-f72d-46ad-9878-e131cfa24800.png'
if not os.path.exists(input_file):
    print(f"❌ '{input_file}' 파일이 현재 디렉토리에 없습니다.")
    exit(1)

# resources 폴더 생성
os.makedirs('resources', exist_ok=True)

# 선택한 PNG 파일
input_file = '9461e7c2-f72d-46ad-9878-e131cfa24800.png'  # 파일명
output_file = 'resources/9461e7c2-f72d-46ad-9878-e131cfa24800.ico'

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
