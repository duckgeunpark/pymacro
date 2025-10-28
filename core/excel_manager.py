"""
엑셀 데이터 소스 관리
"""
import pandas as pd
import traceback
import os
import shutil


class ExcelManager:
    """엑셀 데이터 관리 클래스"""
    
    def __init__(self):
        self.excel_sources = []
        self.next_id = 1
        self.excel_folder = 'projects/excel'  # 엑셀 저장 폴더
    
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
    
    def add_excel_source(self, name, filepath, sheet_name, columns):
        """엑셀 소스 추가 (파일 복사)"""
        try:
            print(f"📊 엑셀 파일 처리 시작: {filepath}")
            
            # 1. 파일을 프로젝트 폴더로 복사
            copied_filepath = self.copy_excel_to_project(filepath)
            if not copied_filepath:
                raise Exception("파일 복사 실패")
            
            # 2. 복사된 파일에서 데이터 읽기
            print(f"   시트: {sheet_name}")
            print(f"   선택 컬럼: {columns}")
            
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
                'preview': df.head(5).to_dict('records')
            }
            
            self.excel_sources.append(source)
            self.next_id += 1
            
            print(f"✅ 엑셀 소스 '{name}' 추가 완료! (파일: {filename})")
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
            
            # projects/excel/ 폴더에서 파일 찾기
            filepath = os.path.join(self.excel_folder, source['filepath'])
            print(f"   경로: {filepath}")
            
            # 파일 존재 확인
            if not os.path.exists(filepath):
                raise FileNotFoundError(f"엑셀 파일을 찾을 수 없습니다: {filepath}")
            
            df = pd.read_excel(filepath, sheet_name=source['sheet_name'])
            
            if source['columns']:
                df = df[source['columns']]
            
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
        """시트의 칼럼 목록"""
        try:
            print(f"📋 컬럼 목록 읽기: {filepath} - 시트: {sheet_name}")
            
            df = pd.read_excel(filepath, sheet_name=sheet_name, nrows=0)
            columns = df.columns.tolist()
            
            processed_columns = []
            for i, col in enumerate(columns):
                col_str = str(col).strip()
                
                if not col_str or col_str.startswith('Unnamed'):
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
            print(f"📚 {len(source_list)}개의 엑셀 소스 로드됨")
    
    def to_list(self):
        """리스트로 변환"""
        return self.excel_sources
