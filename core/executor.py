"""
매크로 실행 엔진
"""
import pyautogui
import time
import pyperclip
from datetime import datetime
import os
import base64
import tempfile
import cv2
import numpy as np
from PIL import ImageGrab

class MacroExecutor:
    """매크로 실행 엔진"""

    def __init__(self, project_data, coord_mgr, excel_mgr, image_mgr, flow_mgr, project_filepath=None):
        self.project_data = project_data
        self.coord_mgr = coord_mgr
        self.excel_mgr = excel_mgr
        self.image_mgr = image_mgr
        self.flow_mgr = flow_mgr
        self.project_filepath = project_filepath

        self.is_running = False
        self.is_paused = False
        self.should_stop = False

        self.current_row = 0
        self.current_action = 0

        # 엑셀 루프 실행 시 사용되는 데이터 (action_type_variable에서 접근)
        self.all_excel_data = {}  # {excel_id: dataframe}
        self.excel_row_indices = {}  # {excel_id: current_row_index}
        self.excel_start_row = 0
        self.infinite_loop = False

        # 로그
        self.log_callback = None
        self.progress_callback = None
        self.error_callback = None
    
    def set_callbacks(self, log_cb=None, progress_cb=None, error_cb=None):
        """콜백 함수 설정"""
        self.log_callback = log_cb
        self.progress_callback = progress_cb
        self.error_callback = error_cb
    
    def log(self, message):
        """로그 출력"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_msg = f"[{timestamp}] {message}"
        print(log_msg)
        if self.log_callback:
            self.log_callback(log_msg)
    
    def update_progress(self, current, total, status=""):
        """진행상황 업데이트"""
        if self.progress_callback:
            self.progress_callback(current, total, status)
    
    def report_error(self, error_msg, screenshot=None):
        """에러 보고"""
        self.log(f"❌ 에러: {error_msg}")
        if self.error_callback:
            self.error_callback(error_msg, screenshot)
    
    def start(self):
        """매크로 실행 시작"""
        self.is_running = True
        self.should_stop = False
        self.log("🚀 매크로 실행 시작")
        
        settings = self.project_data.get('settings', {}).get('execution', {})
        mode = settings.get('mode', 'excel_loop')

        if mode == 'excel_loop' and not self.excel_mgr.excel_sources:
            self.log("⚠️ 엑셀 데이터가 없어 단순 플로우 반복 모드로 전환합니다.")
            mode = 'flow_repeat'
            if 'repeat_count' not in settings:
                settings['repeat_count'] = 1
        
        try:
            if mode == 'excel_loop':
                self.execute_excel_loop(settings)
            elif mode == 'flow_repeat':
                self.execute_flow_repeat(settings)
            elif mode == 'infinite':
                self.execute_infinite(settings)
            
            self.log("✅ 매크로 실행 완료")
        
        except Exception as e:
            self.report_error(f"실행 중 오류 발생: {str(e)}")
        
        finally:
            self.is_running = False
    
    def pause(self):
        """일시정지"""
        self.is_paused = True
        self.log("⏸️ 일시정지")
    
    def resume(self):
        """재개"""
        self.is_paused = False
        self.log("▶️ 재개")
    
    def stop(self):
        """중지"""
        self.should_stop = True
        self.log("⏹️ 중지 요청")
    
    def execute_excel_loop(self, settings):
        """엑셀 행 반복 모드 (무한반복 지원, 각 엑셀 독립 인덱스)"""
        # 모든 엑셀 소스 로드
        if not self.excel_mgr.excel_sources:
            self.report_error("엑셀 데이터 소스가 없습니다.")
            return

        # 첫 번째 엑셀의 행 수를 기준으로 반복
        primary_source = self.excel_mgr.excel_sources[0]
        primary_df = self.excel_mgr.load_excel_data(primary_source['id'])

        if primary_df is None:
            self.report_error("메인 엑셀 데이터를 로드할 수 없습니다.")
            return

        # 모든 엑셀 데이터를 딕셔너리로 관리 {excel_id: dataframe}
        self.all_excel_data = {primary_source.get('id', 1): primary_df}

        # 나머지 엑셀 소스들도 로드
        for source in self.excel_mgr.excel_sources[1:]:
            df = self.excel_mgr.load_excel_data(source['id'])
            if df is not None:
                self.all_excel_data[source.get('id', len(self.all_excel_data) + 1)] = df
            else:
                self.log(f"⚠️ 엑셀{source.get('id')} 로드 실패: {source['name']}")

        start_row = settings.get('excel_start_row', 1) - 1  # 0-based index
        self.excel_start_row = start_row
        end_row = settings.get('excel_end_row', None)
        if end_row is None:
            end_row = len(primary_df)

        total_rows = end_row - start_row
        self.infinite_loop = settings.get('excel_infinite_loop', False)  # 무한반복 옵션

        # 각 엑셀별 독립 인덱스 관리 {excel_id: current_row_index}
        self.excel_row_indices = {excel_id: start_row for excel_id in self.all_excel_data.keys()}

        if self.infinite_loop:
            self.log(f"📊 엑셀 무한반복 모드: 메인 엑셀 {start_row+1}행 ~ {end_row}행 (각 엑셀 독립 반복)")
            # 각 엑셀의 행 수 출력
            for excel_id, df in self.all_excel_data.items():
                source_name = self._get_excel_source_name(excel_id)
                self.log(f"   - {source_name}: {len(df)}행")
        else:
            self.log(f"📊 엑셀 행 반복 모드: {start_row+1}행 ~ {end_row}행 (총 {total_rows}행)")

        loop_count = 0  # 반복 횟수

        while True:  # 무한 루프
            loop_count += 1

            if self.infinite_loop:
                self.log(f"\n🔄 === 반복 {loop_count}회차 시작 ===")

            for row_idx in range(start_row, end_row):
                if self.should_stop:
                    self.log(f"⏹️ 중지됨 (반복 {loop_count}회차, 행 {row_idx + 1})")
                    return

                self.current_row = row_idx + 1

                # 각 엑셀의 현재 행 데이터를 딕셔너리로 구성 (독립 인덱스 사용)
                all_row_data = {}
                for excel_id, df in self.all_excel_data.items():
                    # 각 엑셀의 독립적인 행 인덱스 사용
                    excel_row_idx = self.excel_row_indices[excel_id]

                    if excel_row_idx < len(df):
                        all_row_data[excel_id] = df.iloc[excel_row_idx].to_dict()
                    else:
                        # 행 인덱스가 끝나면 처음으로 돌아감
                        self.excel_row_indices[excel_id] = start_row
                        all_row_data[excel_id] = df.iloc[start_row].to_dict()

                    # 다음 행으로 인덱스 증가
                    self.excel_row_indices[excel_id] += 1

                    # 끝에 도달하면 처음으로 (다음 반복을 위해)
                    if self.excel_row_indices[excel_id] >= len(df):
                        if self.infinite_loop:
                            self.excel_row_indices[excel_id] = start_row
                            source_name = self._get_excel_source_name(excel_id)
                            self.log(f"   🔄 {source_name} 끝 도달 → 처음부터 반복")
                            # 해당 엑셀의 중복 체크 배열 초기화
                            self.excel_mgr.reset_pasted_values(excel_id)

                # 현재 각 엑셀의 행 번호 로그
                row_info_parts = []
                for excel_id in self.all_excel_data.keys():
                    # 현재 사용된 행 번호 (증가 전 값)
                    used_row = self.excel_row_indices[excel_id]
                    if used_row == start_row:
                        # 방금 리셋되었으면 마지막 행이었던 것
                        df = self.all_excel_data[excel_id]
                        used_row = len(df)
                    source_name = self._get_excel_source_name(excel_id)
                    row_info_parts.append(f"{source_name}:{used_row}")

                self.log(f"\n--- 메인행 {self.current_row} 처리 ({', '.join(row_info_parts)}) ---")

                if self.infinite_loop:
                    status = f"반복 {loop_count}회차 - 행 {self.current_row}/{end_row} 처리 중"
                else:
                    status = f"행 {self.current_row} 처리 중"

                self.update_progress(row_idx - start_row + 1, total_rows, status)

                # 플로우 실행
                try:
                    self.execute_flow(all_row_data)
                except Exception as e:
                    on_error = settings.get('on_error', 'skip')
                    if on_error == 'stop':
                        self.report_error(f"행 {self.current_row}에서 오류 발생. 중지합니다.")
                        return
                    elif on_error == 'skip':
                        self.report_error(f"행 {self.current_row}에서 오류 발생. 건너뜁니다: {str(e)}")
                        continue
                    elif on_error == 'retry':
                        retry_count = settings.get('retry_count', 3)
                        for attempt in range(retry_count):
                            self.log(f"재시도 {attempt+1}/{retry_count}")
                            try:
                                self.execute_flow(all_row_data)
                                break
                            except:
                                if attempt == retry_count - 1:
                                    self.report_error(f"행 {self.current_row} 재시도 실패. 건너뜁니다.")

            # 무한반복이 아니면 한 번만 실행하고 종료
            if not self.infinite_loop:
                break

            # 무한반복일 경우 다시 처음부터 (메인 엑셀만 리셋, 나머지는 독립 인덱스 유지)
            if self.infinite_loop:
                self.log(f"✅ 반복 {loop_count}회차 완료. 메인 엑셀 처음부터 다시 시작...")
                time.sleep(0.5)  # 약간의 딜레이

    def _get_excel_source_name(self, excel_id):
        """엑셀 ID로 소스 이름 조회"""
        for source in self.excel_mgr.excel_sources:
            if source.get('id') == excel_id:
                return source.get('name', f'엑셀{excel_id}')
        return f'엑셀{excel_id}'

    
    def execute_flow_repeat(self, settings):
        """플로우 반복 모드"""
        repeat_count = settings.get('repeat_count', 1)
        self.log(f"🔁 플로우 반복 모드: {repeat_count}회")
        
        for i in range(repeat_count):
            if self.should_stop:
                break
            
            self.log(f"\n--- 반복 {i+1}/{repeat_count} ---")
            self.update_progress(i+1, repeat_count, f"반복 {i+1} 실행 중")
            
            try:
                self.execute_flow()
            except Exception as e:
                self.report_error(f"반복 {i+1}에서 오류: {str(e)}")
    
    def execute_infinite(self, settings):
        """무한 반복 모드"""
        self.log("♾️ 무한 반복 모드 (중지할 때까지 계속)")
        
        iteration = 0
        while not self.should_stop:
            iteration += 1
            self.log(f"\n--- 반복 {iteration} ---")
            self.update_progress(iteration, -1, f"반복 {iteration} 실행 중")
            
            try:
                self.execute_flow()
            except Exception as e:
                self.report_error(f"반복 {iteration}에서 오류: {str(e)}")
    
    def execute_flow(self, row_data=None):
        """플로우 시퀀스 실행"""
        for idx, action in enumerate(self.flow_mgr.flow_sequence):
            # 일시정지 체크
            while self.is_paused and not self.should_stop:
                time.sleep(0.1)
            
            if self.should_stop:
                break
            
            self.current_action = idx + 1
            
            try:
                self.execute_action(action, row_data)
            except Exception as e:
                raise Exception(f"액션 {idx+1} 실행 오류: {str(e)}")
    
    def execute_action(self, action, row_data=None):
        """개별 액션 실행"""
        action_type = action['type']
        params = action['params']
        
        # 액션 로그
        display_text = self.flow_mgr.get_action_display_text(
            action, self.coord_mgr, self.excel_mgr, self.image_mgr
        )
        self.log(f"  ▶ {display_text}")
        
        if action_type == 'click_coord':
            self.action_click_coord(params)
        
        elif action_type == 'click_image':
            self.action_click_image(params)
        
        elif action_type == 'type_text':
            self.action_type_text(params)
        
        elif action_type == 'type_variable':
            self.action_type_variable(params, row_data)
        
        elif action_type == 'key_press':
            self.action_key_press(params)
        
        elif action_type == 'hotkey':
            self.action_hotkey(params)
        
        elif action_type == 'paste':
            self.action_paste()

        elif action_type == 'mouse_scroll':
            self.action_mouse_scroll(params)

        elif action_type == 'delay':
            self.action_delay(params)
        
        elif action_type == 'wait_image':
            self.action_wait_image(params)
        
        elif action_type == 'screenshot':
            self.action_screenshot(params)
        
        elif action_type == 'memo':
            pass  # 메모는 실행하지 않음

        
        else:
            self.log(f"    ⚠️ 알 수 없는 액션 타입: {action_type}")
    
    def action_click_coord(self, params):
        """좌표 클릭"""
        coord_id = params.get('coord_id')
        coord = self.coord_mgr.get_coordinate(coord_id)
        
        if not coord:
            raise Exception(f"좌표 ID {coord_id}를 찾을 수 없습니다.")
        
        x, y = coord['x'], coord['y']
        click_type = params.get('click_type', 'left')
        click_count = params.get('click_count', 1)
        
        pre_delay = params.get('pre_delay', 0.2)
        post_delay = params.get('post_delay', 0.2)
        
        time.sleep(pre_delay)
        
        if click_type == 'left':
            pyautogui.click(x, y, clicks=click_count)
        elif click_type == 'right':
            pyautogui.rightClick(x, y)
        elif click_type == 'middle':
            pyautogui.middleClick(x, y)
        
        time.sleep(post_delay)
    
    def action_click_image(self, params):
        """이미지 클릭 (수정됨)"""
        image_id = params.get('image_id')
        image = self.image_mgr.get_image(image_id)
        
        if not image:
            raise Exception(f"이미지 ID {image_id}를 찾을 수 없습니다.")
        
        self.log(f"    🔍 이미지 '{image['name']}' 찾는 중...")
        
        try:
            # base64 이미지를 임시 파일로 저장
            img_data = base64.b64decode(image['data'])
            
            # 임시 파일 생성
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_file:
                tmp_file.write(img_data)
                tmp_path = tmp_file.name
            
            # PyAutoGUI로 이미지 찾기
            confidence = image.get('confidence', 0.8)
            location = pyautogui.locateOnScreen(tmp_path, confidence=confidence)
            
            # 임시 파일 삭제
            os.unlink(tmp_path)
            
            if location:
                # 중심점 클릭
                center = pyautogui.center(location)
                self.log(f"    ✅ 이미지 발견: ({center.x}, {center.y})")
                
                time.sleep(0.2)
                pyautogui.click(center.x, center.y)
                time.sleep(0.2)
            else:
                raise Exception(f"이미지 '{image['name']}'을(를) 찾을 수 없습니다.")
        
        except Exception as e:
            raise Exception(f"이미지 클릭 오류: {str(e)}")
    
    def action_type_text(self, params):
        """텍스트 타이핑 (한글/영문 모두 지원 - pyperclip 사용)"""
        text = params.get('text', '')
        
        try:
            # 클립보드로 복사 후 붙여넣기 (모든 언어 지원)
            pyperclip.copy(text)
            time.sleep(0.1)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.2)
        except Exception as e:
            self.log(f"    ⚠️ 타이핑 오류: {e}")
            raise Exception(f"텍스트 타이핑 실패: {str(e)}")
    
    def action_type_variable(self, params, row_data):
        """변수 타이핑 (한글/영문 모두 지원 - pyperclip 사용)"""
        var_type = params.get('var_type')
        var_name = params.get('var_name', '')
        excel_id = params.get('excel_id', 1)  # 기본값 1 (하위 호환성)

        if var_type == 'excel' and row_data:
            # row_data가 딕셔너리의 딕셔너리인 경우 (다중 엑셀)
            if isinstance(row_data, dict) and excel_id in row_data:
                text = str(row_data[excel_id].get(var_name, ''))
            # row_data가 단순 딕셔너리인 경우 (기존 단일 엑셀, 하위 호환성)
            elif isinstance(row_data, dict) and var_name in row_data:
                text = str(row_data.get(var_name, ''))
            else:
                text = ''

            # 엑셀 데이터인 경우 중복 체크 및 다음 행 찾기
            if text and var_type == 'excel' and excel_id in self.all_excel_data:
                skip_count = 0

                # 중복이 아닌 값이 나올 때까지 다음 행 탐색
                while self.excel_mgr.is_duplicate(excel_id, text):
                    skip_count += 1
                    self.log(f"    ⏭️ 중복 데이터 스킵 ({skip_count}번째): {text}")

                    # 다음 행으로 인덱스 증가
                    df = self.all_excel_data[excel_id]
                    self.excel_row_indices[excel_id] += 1

                    # 끝에 도달하면 처음으로
                    if self.excel_row_indices[excel_id] >= len(df):
                        if self.infinite_loop:
                            self.excel_row_indices[excel_id] = self.excel_start_row
                            source_name = self._get_excel_source_name(excel_id)
                            self.log(f"    🔄 {source_name} 끝 도달 → 처음부터 반복")
                            # 해당 엑셀의 중복 체크 배열 초기화
                            self.excel_mgr.reset_pasted_values(excel_id)
                        else:
                            # 무한 루프가 아니면 더 이상 진행 불가
                            self.log(f"    ⚠️ 엑셀 끝에 도달했지만 중복이 아닌 데이터 없음")
                            return

                    # 다음 행 데이터 가져오기
                    next_row_idx = self.excel_row_indices[excel_id]
                    if next_row_idx < len(df):
                        text = str(df.iloc[next_row_idx][var_name])
                    else:
                        self.log(f"    ⚠️ 다음 행 데이터를 가져올 수 없음")
                        return

                if skip_count > 0:
                    self.log(f"    ✅ 중복 아닌 데이터 발견 (총 {skip_count}개 스킵): {text}")

        elif var_type == 'counter':
            text = str(self.current_row)
        elif var_type == 'timestamp':
            text = datetime.now().strftime('%Y%m%d_%H%M%S')
        else:
            text = ''

        try:
            # 클립보드로 복사 후 붙여넣기 (모든 언어 지원)
            pyperclip.copy(text)
            time.sleep(0.1)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.2)

            # 엑셀 데이터를 성공적으로 붙여넣었으면 기록
            if var_type == 'excel' and text:
                self.excel_mgr.add_pasted_value(excel_id, text)

        except Exception as e:
            self.log(f"    ⚠️ 변수 타이핑 오류: {e}")
            raise Exception(f"변수 타이핑 실패: {str(e)}")
    
    def action_key_press(self, params):
        """키 입력"""
        key = params.get('key', '')

        # 조합키 처리 (예: ctrl+v, shift+a)
        if '+' in key:
            keys = key.split('+')
            pyautogui.hotkey(*keys)
        else:
            # 단일 키
            pyautogui.press(key)

        time.sleep(0.2)
    
    def action_hotkey(self, params):
        """단축키"""
        keys = params.get('keys', [])
        pyautogui.hotkey(*keys)
        time.sleep(0.2)
    
    def action_paste(self):
        """붙여넣기"""
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.2)

    def action_mouse_scroll(self, params):
        """마우스 스크롤"""
        direction = params.get('direction', 'down')
        amount = params.get('amount', 3)

        # pyautogui.scroll: 양수는 위로, 음수는 아래로
        # amount를 100 단위로 변환 (더 부드러운 스크롤)
        scroll_amount = amount * 100 if direction == 'up' else -amount * 100

        pyautogui.scroll(scroll_amount)
        time.sleep(0.2)

    def action_delay(self, params):
        """대기"""
        seconds = params.get('seconds', 1)
        time.sleep(seconds)
        
    def action_wait_image(self, params):
        """이미지가 나타날 때까지 대기 (OpenCV 기반 수정)"""
        image_id = params.get('image_id')
        timeout = params.get('timeout', 10)

        image = self.image_mgr.get_image(image_id)
        if not image:
            raise Exception(f"이미지 ID {image_id}를 찾을 수 없습니다.")

        self.log(f"   ⏳ 이미지 '{image['name']}' 대기 중... (최대 {timeout}초)")

        try:
            # base64 디코딩 후 numpy 배열로 변환
            img_data = base64.b64decode(image['data'])
            nparr = np.frombuffer(img_data, np.uint8)
            template = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
            template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)
            w, h = template_gray.shape[::-1]

            start_time = time.time()
            confidence_threshold = image.get('confidence', 0.6)

            while time.time() - start_time < timeout:
                if self.should_stop:
                    raise Exception("사용자가 중지했습니다.")

                # 화면 캡처 후 그레이스케일 변환
                screen = np.array(ImageGrab.grab())
                screen_gray = cv2.cvtColor(screen, cv2.COLOR_BGR2GRAY)

                # 템플릿 매칭
                res = cv2.matchTemplate(screen_gray, template_gray, cv2.TM_CCOEFF_NORMED)
                _, max_val, _, max_loc = cv2.minMaxLoc(res)

                # 매칭 신뢰도가 기준 이상이면 위치 반환
                if max_val >= confidence_threshold:
                    center_x = max_loc[0] + w // 2
                    center_y = max_loc[1] + h // 2
                    elapsed = time.time() - start_time
                    self.log(f"   ✅ 이미지 발견! ({center_x}, {center_y}) - {elapsed:.1f}초 소요")
                    return

                time.sleep(0.5)

            raise Exception(f"이미지 '{image['name']}'을(를) {timeout}초 내에 찾을 수 없습니다.")

        except Exception as e:
            raise Exception(f"이미지 대기 오류: {str(e)}")
    
    def action_screenshot(self, params):
        """스크린샷 저장"""
        base_filename = params.get('filename', 'screenshot')

        # 확장자가 있으면 제거
        if base_filename.endswith('.png'):
            base_filename = base_filename[:-4]

        # 타임스탬프 추가
        filename = f"{base_filename}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"

        # 프로젝트 폴더 내부에 logs/screenshots 폴더 생성
        if self.project_filepath:
            # 프로젝트 파일이 있는 디렉토리 경로
            project_dir = os.path.dirname(self.project_filepath)
            screenshot_dir = os.path.join(project_dir, 'logs', 'screenshots')
        else:
            # 프로젝트 파일 경로가 없으면 현재 디렉토리 사용
            screenshot_dir = os.path.join('logs', 'screenshots')

        os.makedirs(screenshot_dir, exist_ok=True)

        filepath = os.path.join(screenshot_dir, filename)

        screenshot = pyautogui.screenshot()
        screenshot.save(filepath)
        self.log(f"    💾 스크린샷 저장: {filepath}")


