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


class MacroExecutor:
    """매크로 실행 엔진"""
    
    def __init__(self, project_data, coord_mgr, excel_mgr, image_mgr, flow_mgr):
        self.project_data = project_data
        self.coord_mgr = coord_mgr
        self.excel_mgr = excel_mgr
        self.image_mgr = image_mgr
        self.flow_mgr = flow_mgr
        
        self.is_running = False
        self.is_paused = False
        self.should_stop = False
        
        self.current_row = 0
        self.current_action = 0
        
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
        """엑셀 행 반복 모드 (무한반복 지원)"""
        # 엑셀 소스 가져오기
        if not self.excel_mgr.excel_sources:
            self.report_error("엑셀 데이터 소스가 없습니다.")
            return
        
        excel_source = self.excel_mgr.excel_sources[0]
        df = self.excel_mgr.load_excel_data(excel_source['id'])
        
        if df is None:
            self.report_error("엑셀 데이터를 로드할 수 없습니다.")
            return
        
        start_row = settings.get('excel_start_row', 1) - 1  # 0-based index
        end_row = settings.get('excel_end_row', None)
        if end_row is None:
            end_row = len(df)
        
        total_rows = end_row - start_row
        infinite_loop = settings.get('excel_infinite_loop', False)  # 무한반복 옵션
        
        if infinite_loop:
            self.log(f"📊 엑셀 무한반복 모드: {start_row+1}행 ~ {end_row}행 (중지할 때까지 반복)")
        else:
            self.log(f"📊 엑셀 행 반복 모드: {start_row+1}행 ~ {end_row}행 (총 {total_rows}행)")
        
        loop_count = 0  # 반복 횟수
        
        while True:  # 무한 루프
            loop_count += 1
            
            if infinite_loop:
                self.log(f"\n🔄 === 반복 {loop_count}회차 시작 ===")
            
            for row_idx in range(start_row, end_row):
                if self.should_stop:
                    self.log(f"⏹️ 중지됨 (반복 {loop_count}회차, 행 {row_idx + 1})")
                    return
                
                self.current_row = row_idx + 1
                row_data = df.iloc[row_idx].to_dict()
                
                self.log(f"\n--- 행 {self.current_row} 처리 시작 ---")
                
                if infinite_loop:
                    status = f"반복 {loop_count}회차 - 행 {self.current_row}/{end_row} 처리 중"
                else:
                    status = f"행 {self.current_row} 처리 중"
                
                self.update_progress(row_idx - start_row + 1, total_rows, status)
                
                # 플로우 실행
                try:
                    self.execute_flow(row_data)
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
                                self.execute_flow(row_data)
                                break
                            except:
                                if attempt == retry_count - 1:
                                    self.report_error(f"행 {self.current_row} 재시도 실패. 건너뜁니다.")
            
            # 무한반복이 아니면 한 번만 실행하고 종료
            if not infinite_loop:
                break
            
            # 무한반복일 경우 다시 처음부터
            if infinite_loop:
                self.log(f"✅ 반복 {loop_count}회차 완료. 처음부터 다시 시작합니다...")
                time.sleep(0.5)  # 약간의 딜레이

    
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
        
        if var_type == 'excel' and row_data:
            text = str(row_data.get(var_name, ''))
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
        except Exception as e:
            self.log(f"    ⚠️ 변수 타이핑 오류: {e}")
            raise Exception(f"변수 타이핑 실패: {str(e)}")
    
    def action_key_press(self, params):
        """키 입력"""
        key = params.get('key', '')
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
    
    def action_delay(self, params):
        """대기"""
        seconds = params.get('seconds', 1)
        time.sleep(seconds)
    
    def action_wait_image(self, params):
        """이미지가 나타날 때까지 대기 (수정됨)"""
        image_id = params.get('image_id')
        timeout = params.get('timeout', 10)
        
        image = self.image_mgr.get_image(image_id)
        if not image:
            raise Exception(f"이미지 ID {image_id}를 찾을 수 없습니다.")
        
        self.log(f"    ⏳ 이미지 '{image['name']}' 대기 중... (최대 {timeout}초)")
        
        try:
            # base64 이미지를 임시 파일로 저장
            img_data = base64.b64decode(image['data'])
            
            # 임시 파일 생성
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_file:
                tmp_file.write(img_data)
                tmp_path = tmp_file.name
            
            # 타임아웃까지 반복 검색
            start_time = time.time()
            confidence = image.get('confidence', 0.8)
            
            while time.time() - start_time < timeout:
                if self.should_stop:
                    os.unlink(tmp_path)
                    raise Exception("사용자가 중지했습니다.")
                
                location = pyautogui.locateOnScreen(tmp_path, confidence=confidence)
                
                if location:
                    center = pyautogui.center(location)
                    elapsed = time.time() - start_time
                    self.log(f"    ✅ 이미지 발견! ({center.x}, {center.y}) - {elapsed:.1f}초 소요")
                    os.unlink(tmp_path)
                    return
                
                time.sleep(0.5)  # 0.5초마다 체크
            
            # 임시 파일 삭제
            os.unlink(tmp_path)
            
            raise Exception(f"이미지 '{image['name']}'을(를) {timeout}초 내에 찾을 수 없습니다.")
        
        except Exception as e:
            raise Exception(f"이미지 대기 오류: {str(e)}")
    
    def action_screenshot(self, params):
        """스크린샷 저장"""
        filename = params.get('filename', f'screenshot_{datetime.now().strftime("%Y%m%d_%H%M%S")}.png')
        
        # logs/screenshots 폴더 생성
        screenshot_dir = os.path.join('logs', 'screenshots')
        os.makedirs(screenshot_dir, exist_ok=True)
        
        filepath = os.path.join(screenshot_dir, filename)
        
        screenshot = pyautogui.screenshot()
        screenshot.save(filepath)
        self.log(f"    💾 스크린샷 저장: {filepath}")
