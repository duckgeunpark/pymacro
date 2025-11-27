"""
엑셀 데이터 소스 관리
"""
import pandas as pd
import traceback
import os
import shutil
import re
from datetime import datetime


class ExcelManager:
    """엑셀 데이터 관리 클래스"""

    def __init__(self):
        self.excel_sources = []
        self.next_id = 1
        self.excel_folder = 'projects/excel'  # 엑셀 저장 폴더

        # 중복 체크용 배열 (엑셀 ID별로 붙여넣은 값들을 추적)
        self.pasted_values = {}  # {excel_id: set()}
    
    def copy_excel_to_project(self, source_filepath):
        """엑셀 파일을 프로젝트 폴더로 복사"""
        try:
            # 폴더 생성 (없으면)
            os.makedirs(self.excel_folder, exist_ok=True)
            
            # 파일명 추출
            filename = os.path.basename(source_filepath)
            
            # 중복 방지: 같은 이름이 있으면 번호 추가
            dest_path = os.path.join(self.excel_folder, filename)
            if os.path.exists(dest_path):
                name, ext = os.path.splitext(filename)
                counter = 1
                while os.path.exists(dest_path):
                    new_filename = f"{name}_{counter}{ext}"
                    dest_path = os.path.join(self.excel_folder, new_filename)
                    counter += 1
            
            # 파일 복사
            shutil.copy2(source_filepath, dest_path)
            print(f"📋 엑셀 파일 복사: {source_filepath} → {dest_path}")
            
            return dest_path
            
        except Exception as e:
            print(f"❌ 파일 복사 오류: {e}")
            traceback.print_exc()
            return None
    
    def add_excel_source(self, name, filepath, sheet_name, columns, auto_latest=False, auto_directory=None, auto_prefix='list', remove_empty_rows=True, remove_duplicates=True):
        """
        엑셀 소스 추가 (파일 복사)

        Args:
            name: 소스 이름
            filepath: 엑셀 파일 경로
            sheet_name: 시트 이름
            columns: 사용할 컬럼 리스트
            auto_latest: 자동 최신 파일 선택 여부
            auto_directory: 자동 선택 시 검색할 디렉토리
            auto_prefix: 자동 선택 시 파일명 prefix
            remove_empty_rows: 상위 공백 행 제거 여부
            remove_duplicates: 중복 행 제거 여부
        """
        try:
            print(f"📊 엑셀 파일 처리 시작: {filepath}")

            # 1. 파일을 프로젝트 폴더로 복사
            copied_filepath = self.copy_excel_to_project(filepath)
            if not copied_filepath:
                raise Exception("파일 복사 실패")

            # 2. 복사된 파일에서 데이터 읽기
            print(f"   시트: {sheet_name}")
            print(f"   선택 컬럼: {columns}")

            # 상위 공백 행 제거를 위해 header=None으로 먼저 읽기
            if remove_empty_rows:
                df = pd.read_excel(copied_filepath, sheet_name=sheet_name, header=None)

                # 첫 번째 유효한 행 찾기
                first_valid_idx = None
                for idx, row in df.iterrows():
                    if row.notna().any() and row.astype(str).str.strip().ne('').any():
                        first_valid_idx = idx
                        break

                if first_valid_idx is None:
                    raise ValueError("유효한 데이터 행을 찾을 수 없습니다")

                if first_valid_idx > 0:
                    df = df.iloc[first_valid_idx:].reset_index(drop=True)
                    print(f"   🧹 상위 공백 행 {first_valid_idx}개 제거됨")

                # 첫 번째 행을 헤더로 설정
                df.columns = df.iloc[0]
                df = df[1:].reset_index(drop=True)
                print(f"   📋 헤더 설정 완료: {list(df.columns)}")
            else:
                df = pd.read_excel(copied_filepath, sheet_name=sheet_name)

            print(f"   ✅ 전체 {len(df)}행, {len(df.columns)}컬럼 읽기 완료")
            print(f"   전체 컬럼: {list(df.columns)}")

            # 3. 선택된 칼럼만 추출
            if columns:
                missing_cols = [col for col in columns if col not in df.columns]
                if missing_cols:
                    raise ValueError(f"존재하지 않는 컬럼: {missing_cols}")

                df = df[columns]
                print(f"   ✅ {len(columns)}개 컬럼 필터링 완료")

            # 4. 중복 행 제거 (옵션)
            if remove_duplicates:
                original_count = len(df)
                # 선택한 컬럼 기준으로 중복 제거 (첫 번째 발견된 행 유지)
                df = df.drop_duplicates(subset=columns if columns else None, keep='first').reset_index(drop=True)
                removed = original_count - len(df)
                if removed > 0:
                    print(f"   🔄 중복 행 {removed}개 제거됨 (남은 행: {len(df)})")
                else:
                    print(f"   ℹ️ 중복 행 없음")

            row_count = len(df)

            # 4. 파일명만 저장 (projects/excel/ 기준)
            filename = os.path.basename(copied_filepath)

            source = {
                'id': self.next_id,
                'name': name,
                'filepath': filename,  # 파일명만 저장
                'sheet_name': sheet_name,
                'columns': columns if columns else list(df.columns),
                'row_count': row_count,
                'auto_latest': auto_latest,  # 자동 최신 파일 선택 여부
                'auto_directory': auto_directory,  # 검색 디렉토리
                'auto_prefix': auto_prefix,  # 파일명 prefix
                'remove_empty_rows': remove_empty_rows,  # 공백 행 제거 옵션
                'remove_duplicates': remove_duplicates  # 중복 행 제거 옵션
            }

            self.excel_sources.append(source)
            self.next_id += 1

            print(f"✅ 엑셀 소스 '{name}' 추가 완료! (파일: {filename})")
            if auto_latest:
                print(f"   🔄 자동 최신 파일 선택 활성화 (디렉토리: {auto_directory}, prefix: {auto_prefix})")
            return source

        except Exception as e:
            print(f"❌ 엑셀 로드 오류: {e}")
            traceback.print_exc()
            return None
    
    def remove_excel_source(self, source_id):
        """엑셀 소스 삭제 (파일도 삭제)"""
        source = self.get_excel_source(source_id)
        if source:
            # 파일 삭제
            try:
                filepath = os.path.join(self.excel_folder, source['filepath'])
                if os.path.exists(filepath):
                    os.remove(filepath)
                    print(f"🗑️ 엑셀 파일 삭제: {filepath}")
            except Exception as e:
                print(f"⚠️ 파일 삭제 실패: {e}")
        
        self.excel_sources = [s for s in self.excel_sources if s['id'] != source_id]
        print(f"🗑️ 엑셀 소스 ID {source_id} 삭제됨")
    
    def get_excel_source(self, source_id):
        """ID로 엑셀 소스 찾기"""
        for source in self.excel_sources:
            if source['id'] == source_id:
                return source
        return None
    
    def load_excel_data(self, source_id):
        """엑셀 데이터 전체 로드"""
        source = self.get_excel_source(source_id)
        if not source:
            print(f"❌ 엑셀 소스 ID {source_id}를 찾을 수 없습니다.")
            return None

        try:
            print(f"📊 엑셀 데이터 로드 중: {source['name']}")

            # 자동 최신 파일 선택 모드인 경우
            if source.get('auto_latest', False):
                auto_directory = source.get('auto_directory')
                auto_prefix = source.get('auto_prefix', '')
                auto_mode = source.get('auto_mode', 'modified')  # 'modified' or 'number'

                if not auto_directory:
                    raise ValueError("자동 최신 파일 선택이 활성화되었지만 디렉토리가 지정되지 않았습니다")

                mode_text = "수정일" if auto_mode == "modified" else "숫자"
                print(f"   🔄 자동 최신 파일 검색 중... (기준: {mode_text})")
                print(f"   디렉토리: {auto_directory}")
                if auto_prefix:
                    print(f"   Prefix: {auto_prefix}")

                # 최신 파일 찾기
                latest_filepath = self.find_latest_file(auto_directory, auto_prefix, auto_mode)

                if not latest_filepath:
                    prefix_msg = f"'{auto_prefix}'로 시작하는 " if auto_prefix else ""
                    raise FileNotFoundError(f"{prefix_msg}최신 파일을 찾을 수 없습니다: {auto_directory}")

                filepath = latest_filepath
                print(f"   ✅ 최신 파일 로드: {os.path.basename(filepath)}")

            else:
                # 일반 모드: projects/excel/ 폴더에서 파일 찾기
                filepath = os.path.join(self.excel_folder, source['filepath'])
                print(f"   경로: {filepath}")

                # 파일 존재 확인
                if not os.path.exists(filepath):
                    raise FileNotFoundError(f"엑셀 파일을 찾을 수 없습니다: {filepath}")

            # 상위 공백 행 제거가 필요한 경우 header=None으로 읽기
            if source.get('remove_empty_rows', True):
                # header 없이 읽어서 공백 행 찾기
                df = pd.read_excel(filepath, sheet_name=source['sheet_name'], header=None)

                original_count = len(df)
                first_valid_idx = None
                for idx, row in df.iterrows():
                    if row.notna().any() and row.astype(str).str.strip().ne('').any():
                        first_valid_idx = idx
                        break

                if first_valid_idx is not None and first_valid_idx > 0:
                    # 공백 행 제거
                    df = df.iloc[first_valid_idx:].reset_index(drop=True)
                    removed = original_count - len(df)
                    print(f"   🧹 상위 공백 행 {removed}개 제거됨")

                # 첫 번째 행을 헤더로 설정
                df.columns = df.iloc[0]
                df = df[1:].reset_index(drop=True)
                print(f"   📋 헤더 설정 완료: {list(df.columns)}")
            else:
                # 공백 행 제거가 필요 없으면 일반적으로 읽기
                df = pd.read_excel(filepath, sheet_name=source['sheet_name'])

            # 2. 컬럼 필터링 (공백 행 제거 후 올바른 컬럼명 사용)
            if source['columns']:
                df = df[source['columns']]

            # 3. 중복 행 제거 (저장된 옵션 적용, 기본값 True)
            if source.get('remove_duplicates', True):
                original_count = len(df)

                # 디버깅: 중복 제거 전 데이터 확인
                print(f"   🔍 중복 제거 시작 - 원본 행 수: {original_count}")
                if source['columns']:
                    print(f"   📋 중복 검사 기준 컬럼: {source['columns']}")
                    # 각 컬럼별 값 확인
                    for col in source['columns']:
                        unique_count = df[col].nunique()
                        print(f"      - {col}: {unique_count}개 고유값")
                else:
                    print(f"   📋 전체 컬럼 기준 중복 검사")

                # 중복 제거
                df = df.drop_duplicates(subset=source['columns'] if source['columns'] else None, keep='first').reset_index(drop=True)
                removed = original_count - len(df)

                if removed > 0:
                    print(f"   🔄 중복 행 {removed}개 제거됨 (남은 행: {len(df)})")
                else:
                    print(f"   ℹ️ 중복 행 없음 (총 {len(df)}행)")

            print(f"✅ {len(df)}행 로드 완료")
            return df

        except Exception as e:
            print(f"❌ 엑셀 데이터 로드 오류: {e}")
            traceback.print_exc()
            return None
    
    def get_row_data(self, source_id, row_index):
        """특정 행 데이터 가져오기"""
        df = self.load_excel_data(source_id)
        if df is not None and row_index < len(df):
            return df.iloc[row_index].to_dict()
        return None
    
    @staticmethod
    def get_sheet_names(filepath):
        """엑셀 파일의 시트 이름 목록"""
        try:
            print(f"📄 시트 목록 읽기: {filepath}")
            xl_file = pd.ExcelFile(filepath)
            sheets = xl_file.sheet_names
            print(f"✅ 시트 목록: {sheets}")
            return sheets
            
        except Exception as e:
            print(f"❌ 시트 목록 로드 오류: {e}")
            traceback.print_exc()
            return []
    
    @staticmethod
    def get_columns(filepath, sheet_name):
        """시트의 칼럼 목록 (상위 공백 행 제거 후 헤더 찾기)"""
        try:
            print(f"📋 컬럼 목록 읽기: {filepath} - 시트: {sheet_name}")

            # 먼저 header=None으로 전체 데이터 읽기
            df = pd.read_excel(filepath, sheet_name=sheet_name, header=None)

            # 첫 번째 유효한 행 찾기 (공백 행 건너뛰기)
            first_valid_idx = None
            for idx, row in df.iterrows():
                # 행에 유효한 데이터가 있는지 확인 (null 아니고 공백 문자열 아닌 값)
                if row.notna().any() and row.astype(str).str.strip().ne('').any():
                    first_valid_idx = idx
                    break

            if first_valid_idx is None:
                print(f"⚠️ 유효한 데이터 행을 찾을 수 없습니다")
                return []

            # 유효한 첫 행을 헤더로 사용
            if first_valid_idx > 0:
                print(f"   🧹 상위 공백 행 {first_valid_idx}개 건너뜀")

            columns = df.iloc[first_valid_idx].tolist()

            processed_columns = []
            for i, col in enumerate(columns):
                col_str = str(col).strip()

                if not col_str or col_str in ['nan', 'None'] or col_str.startswith('Unnamed'):
                    processed_columns.append(f"컬럼_{i+1}")
                else:
                    processed_columns.append(col_str)

            print(f"✅ 컬럼 목록 ({len(processed_columns)}개): {processed_columns}")
            return processed_columns
            
        except FileNotFoundError:
            print(f"❌ 파일을 찾을 수 없습니다: {filepath}")
            return []
        except PermissionError:
            print(f"❌ 파일 접근 권한이 없습니다: {filepath}")
            return []
        except Exception as e:
            print(f"❌ 칼럼 목록 로드 오류: {e}")
            print(f"   에러 타입: {type(e).__name__}")
            traceback.print_exc()
            return []
    
    def load_from_list(self, source_list):
        """리스트에서 엑셀 소스 로드"""
        self.excel_sources = source_list
        if source_list:
            self.next_id = max(s['id'] for s in source_list) + 1
        else:
            self.next_id = 1
        
        if source_list:
            print(f"[Excel] {len(source_list)}개의 엑셀 소스 로드됨")
    
    def to_list(self):
        """리스트로 변환"""
        return self.excel_sources

    @staticmethod
    def find_latest_file(directory, prefix='', mode='modified'):
        """
        디렉토리에서 조건에 맞는 최신 엑셀 파일 찾기

        Args:
            directory: 검색할 디렉토리 경로
            prefix: 파일명 prefix (빈 문자열이면 모든 엑셀 파일 대상)
            mode: 'modified' (파일 수정일 기준) 또는 'number' (파일명 숫자 기준)

        Returns:
            최신 파일의 전체 경로 또는 None
        """
        import glob

        try:
            if not os.path.exists(directory):
                print(f"❌ 디렉토리가 존재하지 않습니다: {directory}")
                return None

            # 엑셀 파일 패턴 검색
            if prefix:
                pattern_xlsx = os.path.join(directory, f"{prefix}*.xlsx")
                pattern_xls = os.path.join(directory, f"{prefix}*.xls")
            else:
                pattern_xlsx = os.path.join(directory, "*.xlsx")
                pattern_xls = os.path.join(directory, "*.xls")

            files = glob.glob(pattern_xlsx)
            files.extend([f for f in glob.glob(pattern_xls) if not f.endswith('.xlsx')])

            if not files:
                prefix_msg = f"'{prefix}'로 시작하는 " if prefix else ""
                print(f"❌ {prefix_msg}엑셀 파일이 없습니다: {directory}")
                return None

            prefix_msg = f"'{prefix}'로 시작하는 " if prefix else ""
            print(f"📁 {prefix_msg}엑셀 파일 {len(files)}개 발견")

            if mode == "modified":
                # 수정일 기준 - 가장 최근 수정된 파일
                latest_file = max(files, key=os.path.getmtime)
                mtime = datetime.fromtimestamp(os.path.getmtime(latest_file))
                print(f"✅ 최신 파일 선택 (수정일 기준): {os.path.basename(latest_file)}")
                print(f"   수정일: {mtime.strftime('%Y-%m-%d %H:%M:%S')}")
                return latest_file

            else:
                # 숫자 기준 - 파일명에서 가장 큰 숫자를 가진 파일
                def extract_number(filepath):
                    filename = os.path.basename(filepath)
                    # 모든 숫자 추출 (괄호 안 포함)
                    # 패턴: 20250120, 250120, (20250120), [250120], 20250120132234, 1, 2, 3 등
                    numbers = re.findall(r'\d+', filename)
                    if numbers:
                        # 가장 긴 숫자를 선택 (날짜+시간이 보통 가장 김)
                        return max(int(n) for n in numbers)
                    return 0

                # 숫자가 있는 파일만 필터링
                files_with_numbers = [(f, extract_number(f)) for f in files]
                files_with_numbers = [(f, n) for f, n in files_with_numbers if n > 0]

                if not files_with_numbers:
                    print(f"❌ 숫자가 포함된 파일이 없습니다")
                    return None

                # 숫자 기준 정렬
                latest_file, latest_number = max(files_with_numbers, key=lambda x: x[1])
                print(f"✅ 최신 파일 선택 (숫자 기준): {os.path.basename(latest_file)}")
                print(f"   추출된 숫자: {latest_number}")
                return latest_file

        except Exception as e:
            print(f"❌ 최신 파일 검색 오류: {e}")
            traceback.print_exc()
            return None

    def is_duplicate(self, excel_id, value):
        """중복 값인지 체크"""
        if excel_id not in self.pasted_values:
            self.pasted_values[excel_id] = set()

        return value in self.pasted_values[excel_id]

    def add_pasted_value(self, excel_id, value):
        """붙여넣은 값 기록"""
        if excel_id not in self.pasted_values:
            self.pasted_values[excel_id] = set()

        self.pasted_values[excel_id].add(value)

    def reset_pasted_values(self, excel_id):
        """특정 엑셀의 중복 체크 배열 초기화"""
        if excel_id in self.pasted_values:
            self.pasted_values[excel_id].clear()
            print(f"🔄 엑셀 ID {excel_id} 중복 체크 배열 초기화")
