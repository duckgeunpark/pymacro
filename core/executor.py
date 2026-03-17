"""
매크로 실행 엔진
"""
import pyautogui
import time
import threading
import pyperclip
from datetime import datetime
import os
import base64
import tempfile
import cv2
import numpy as np
from PIL import ImageGrab

# FAILSAFE 활성화 (마우스를 좌상단 모서리로 이동하면 긴급 중지)
pyautogui.FAILSAFE = True

# 실행 관련 기본값 (설정으로 오버라이드 가능)
DEFAULT_PRE_DELAY = 0.2
DEFAULT_POST_DELAY = 0.2
DEFAULT_TYPING_DELAY = 0.1
DEFAULT_PASTE_DELAY = 0.2
DEFAULT_IMAGE_TIMEOUT = 10
DEFAULT_IMAGE_POLL_INTERVAL = 0.5
DEFAULT_IMAGE_CONFIDENCE = 0.8
DEFAULT_WAIT_IMAGE_CONFIDENCE = 0.6
PAUSE_CHECK_INTERVAL = 0.1


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
        # threading.Event로 변경 (thread-safe)
        self._pause_event = threading.Event()
        self._pause_event.set()  # 초기 상태: 실행 중 (not paused)
        self._stop_event = threading.Event()

        self.current_row = 0
        self.current_action = 0

        # 엑셀 루프 실행 시 사용되는 데이터
        self.all_excel_data = {}  # {excel_id: dataframe}
        self.excel_row_indices = {}  # {excel_id: current_row_index}
        self.excel_start_row = 0
        self.infinite_loop = False

        # 이미지 캐시 (base64 디코딩 결과 재사용)
        self._image_cache = {}  # {image_id: {'tmp_path': str, 'np_array': ndarray, 'gray': ndarray}}

        # 로그
        self.log_callback = None
        self.progress_callback = None
        self.error_callback = None

    # threading.Event 기반 프로퍼티 (하위 호환성)
    @property
    def is_paused(self):
        return not self._pause_event.is_set()

    @is_paused.setter
    def is_paused(self, value):
        if value:
            self._pause_event.clear()
        else:
            self._pause_event.set()

    @property
    def should_stop(self):
        return self._stop_event.is_set()

    @should_stop.setter
    def should_stop(self, value):
        if value:
            self._stop_event.set()
        else:
            self._stop_event.clear()

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
        self.log(f"에러: {error_msg}")
        if self.error_callback:
            self.error_callback(error_msg, screenshot)

    def _get_cached_image_path(self, image):
        """이미지 캐시에서 임시 파일 경로 반환 (없으면 생성)"""
        image_id = image['id']
        if image_id not in self._image_cache:
            img_data = base64.b64decode(image['data'])
            tmp_file = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
            tmp_file.write(img_data)
            tmp_file.close()

            # OpenCV용 numpy 배열도 캐시
            nparr = np.frombuffer(img_data, np.uint8)
            template = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
            template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)

            self._image_cache[image_id] = {
                'tmp_path': tmp_file.name,
                'np_array': template,
                'gray': template_gray,
            }

        return self._image_cache[image_id]

    def _cleanup_image_cache(self):
        """이미지 캐시 정리 (임시 파일 삭제)"""
        for cache in self._image_cache.values():
            try:
                if os.path.exists(cache['tmp_path']):
                    os.unlink(cache['tmp_path'])
            except OSError:
                pass
        self._image_cache.clear()

    def start(self):
        """매크로 실행 시작"""
        self.is_running = True
        self._stop_event.clear()
        self._pause_event.set()
        self.log("매크로 실행 시작")

        settings = self.project_data.get('settings', {}).get('execution', {})
        mode = settings.get('mode', 'excel_loop')

        if mode == 'excel_loop' and not self.excel_mgr.excel_sources:
            self.log("엑셀 데이터가 없어 단순 플로우 반복 모드로 전환합니다.")
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

            self.log("매크로 실행 완료")

        except Exception as e:
            self.report_error(f"실행 중 오류 발생: {str(e)}")

        finally:
            self.is_running = False
            self._cleanup_image_cache()

    def pause(self):
        """일시정지"""
        self._pause_event.clear()
        self.log("일시정지")

    def resume(self):
        """재개"""
        self._pause_event.set()
        self.log("재개")

    def stop(self):
        """중지"""
        self._stop_event.set()
        self._pause_event.set()  # pause 상태에서도 즉시 중지 가능하도록
        self.log("중지 요청")

    def execute_excel_loop(self, settings):
        """엑셀 행 반복 모드 (무한반복 지원, 각 엑셀 독립 인덱스)"""
        if not self.excel_mgr.excel_sources:
            self.report_error("엑셀 데이터 소스가 없습니다.")
            return

        primary_source = self.excel_mgr.excel_sources[0]
        primary_df = self.excel_mgr.load_excel_data(primary_source['id'])

        if primary_df is None:
            self.report_error("메인 엑셀 데이터를 로드할 수 없습니다.")
            return

        self.all_excel_data = {primary_source.get('id', 1): primary_df}

        for source in self.excel_mgr.excel_sources[1:]:
            df = self.excel_mgr.load_excel_data(source['id'])
            if df is not None:
                self.all_excel_data[source.get('id', len(self.all_excel_data) + 1)] = df
            else:
                self.log(f"엑셀{source.get('id')} 로드 실패: {source['name']}")

        start_row = settings.get('excel_start_row', 1) - 1
        self.excel_start_row = start_row
        end_row = settings.get('excel_end_row', None)
        if end_row is None:
            end_row = len(primary_df)

        total_rows = end_row - start_row
        self.infinite_loop = settings.get('excel_infinite_loop', False)
        repeat_count = settings.get('repeat_count', 1)

        self.excel_row_indices = {excel_id: start_row for excel_id in self.all_excel_data.keys()}

        if self.infinite_loop:
            self.log(f"엑셀 무한반복 모드: 메인 엑셀 {start_row+1}행 ~ {end_row}행 (각 엑셀 독립 반복)")
            for excel_id, df in self.all_excel_data.items():
                source_name = self._get_excel_source_name(excel_id)
                self.log(f"   - {source_name}: {len(df)}행")
        else:
            self.log(f"엑셀 행 반복 모드: {start_row+1}행 ~ {end_row}행 (총 {total_rows}행) x {repeat_count}회")

        loop_count = 0

        while True:
            loop_count += 1

            if self.infinite_loop:
                self.log(f"\n=== 반복 {loop_count}회차 시작 ===")
            else:
                if loop_count > 1:
                    self.log(f"\n=== 반복 {loop_count}/{repeat_count}회차 시작 ===")

            for row_idx in range(start_row, end_row):
                if self.should_stop:
                    self.log(f"중지됨 (반복 {loop_count}회차, 행 {row_idx + 1})")
                    return

                self.current_row = row_idx + 1

                all_row_data = {}
                for excel_id, df in self.all_excel_data.items():
                    excel_row_idx = self.excel_row_indices[excel_id]

                    if excel_row_idx < len(df):
                        all_row_data[excel_id] = df.iloc[excel_row_idx].to_dict()
                    else:
                        self.excel_row_indices[excel_id] = start_row
                        all_row_data[excel_id] = df.iloc[start_row].to_dict()

                    self.excel_row_indices[excel_id] += 1

                    if self.excel_row_indices[excel_id] >= len(df):
                        if self.infinite_loop:
                            self.excel_row_indices[excel_id] = start_row
                            source_name = self._get_excel_source_name(excel_id)
                            self.log(f"   {source_name} 끝 도달 -> 처음부터 반복")
                            self.excel_mgr.reset_pasted_values(excel_id)

                row_info_parts = []
                for excel_id in self.all_excel_data.keys():
                    used_row = self.excel_row_indices[excel_id]
                    if used_row == start_row:
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
                            except Exception:
                                if attempt == retry_count - 1:
                                    self.report_error(f"행 {self.current_row} 재시도 실패. 건너뜁니다.")

            if not self.infinite_loop:
                if loop_count >= repeat_count:
                    self.log(f"설정된 반복 횟수({repeat_count}회) 완료")
                    break
                else:
                    self.excel_row_indices = {excel_id: start_row for excel_id in self.all_excel_data.keys()}
                    for excel_id in self.all_excel_data.keys():
                        self.excel_mgr.reset_pasted_values(excel_id)
                    self.log(f"반복 {loop_count}/{repeat_count}회차 완료. 다시 시작...")
                    time.sleep(DEFAULT_PASTE_DELAY)
            else:
                self.excel_row_indices = {excel_id: start_row for excel_id in self.all_excel_data.keys()}
                for excel_id in self.all_excel_data.keys():
                    self.excel_mgr.reset_pasted_values(excel_id)
                self.log(f"반복 {loop_count}회차 완료. 메인 엑셀 처음부터 다시 시작...")
                time.sleep(DEFAULT_PASTE_DELAY)

    def _get_excel_source_name(self, excel_id):
        """엑셀 ID로 소스 이름 조회"""
        for source in self.excel_mgr.excel_sources:
            if source.get('id') == excel_id:
                return source.get('name', f'엑셀{excel_id}')
        return f'엑셀{excel_id}'

    def execute_flow_repeat(self, settings):
        """플로우 반복 모드"""
        repeat_count = settings.get('repeat_count', 1)
        self.log(f"플로우 반복 모드: {repeat_count}회")

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
        self.log("무한 반복 모드 (중지할 때까지 계속)")

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
            # 일시정지 체크 (Event.wait으로 thread-safe 대기)
            self._pause_event.wait()

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

        display_text = self.flow_mgr.get_action_display_text(
            action, self.coord_mgr, self.excel_mgr, self.image_mgr
        )
        self.log(f"  > {display_text}")

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
            pass
        else:
            self.log(f"    알 수 없는 액션 타입: {action_type}")

    def action_click_coord(self, params):
        """좌표 클릭"""
        coord_id = params.get('coord_id')
        coord = self.coord_mgr.get_coordinate(coord_id)

        if not coord:
            raise Exception(f"좌표 ID {coord_id}를 찾을 수 없습니다.")

        x, y = coord['x'], coord['y']
        click_type = params.get('click_type', 'left')
        click_count = params.get('click_count', 1)

        pre_delay = params.get('pre_delay', DEFAULT_PRE_DELAY)
        post_delay = params.get('post_delay', DEFAULT_POST_DELAY)

        time.sleep(pre_delay)

        if click_type == 'left':
            pyautogui.click(x, y, clicks=click_count)
        elif click_type == 'right':
            pyautogui.rightClick(x, y)
        elif click_type == 'middle':
            pyautogui.middleClick(x, y)

        time.sleep(post_delay)

    def action_click_image(self, params):
        """이미지 클릭 (캐시 사용)"""
        image_id = params.get('image_id')
        image = self.image_mgr.get_image(image_id)

        if not image:
            raise Exception(f"이미지 ID {image_id}를 찾을 수 없습니다.")

        self.log(f"    이미지 '{image['name']}' 찾는 중...")

        try:
            cache = self._get_cached_image_path(image)
            tmp_path = cache['tmp_path']

            confidence = image.get('confidence', DEFAULT_IMAGE_CONFIDENCE)
            location = pyautogui.locateOnScreen(tmp_path, confidence=confidence)

            if location:
                center = pyautogui.center(location)
                self.log(f"    이미지 발견: ({center.x}, {center.y})")

                time.sleep(DEFAULT_PRE_DELAY)
                pyautogui.click(center.x, center.y)
                time.sleep(DEFAULT_POST_DELAY)
            else:
                raise Exception(f"이미지 '{image['name']}'을(를) 찾을 수 없습니다.")

        except Exception as e:
            raise Exception(f"이미지 클릭 오류: {str(e)}")

    def action_type_text(self, params):
        """텍스트 타이핑 (한글/영문 모두 지원 - pyperclip 사용)"""
        text = params.get('text', '')

        try:
            pyperclip.copy(text)
            time.sleep(DEFAULT_TYPING_DELAY)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(DEFAULT_PASTE_DELAY)
        except Exception as e:
            self.log(f"    타이핑 오류: {e}")
            raise Exception(f"텍스트 타이핑 실패: {str(e)}")

    def action_type_variable(self, params, row_data):
        """변수 타이핑 (한글/영문 모두 지원 - pyperclip 사용)"""
        var_type = params.get('var_type')
        var_name = params.get('var_name', '')
        excel_id = params.get('excel_id', 1)

        if var_type == 'excel' and row_data:
            if isinstance(row_data, dict) and excel_id in row_data:
                text = str(row_data[excel_id].get(var_name, ''))
            elif isinstance(row_data, dict) and var_name in row_data:
                text = str(row_data.get(var_name, ''))
            else:
                text = ''

            if text and var_type == 'excel' and excel_id in self.all_excel_data:
                skip_count = 0

                while self.excel_mgr.is_duplicate(excel_id, text):
                    skip_count += 1
                    self.log(f"    중복 데이터 스킵 ({skip_count}번째): {text}")

                    df = self.all_excel_data[excel_id]
                    self.excel_row_indices[excel_id] += 1

                    if self.excel_row_indices[excel_id] >= len(df):
                        if self.infinite_loop:
                            self.excel_row_indices[excel_id] = self.excel_start_row
                            source_name = self._get_excel_source_name(excel_id)
                            self.log(f"    {source_name} 끝 도달 -> 처음부터 반복")
                            self.excel_mgr.reset_pasted_values(excel_id)
                        else:
                            self.log(f"    엑셀 끝에 도달했지만 중복이 아닌 데이터 없음")
                            return

                    next_row_idx = self.excel_row_indices[excel_id]
                    if next_row_idx < len(df):
                        text = str(df.iloc[next_row_idx][var_name])
                    else:
                        self.log(f"    다음 행 데이터를 가져올 수 없음")
                        return

                if skip_count > 0:
                    self.log(f"    중복 아닌 데이터 발견 (총 {skip_count}개 스킵): {text}")

        elif var_type == 'counter':
            text = str(self.current_row)
        elif var_type == 'timestamp':
            text = datetime.now().strftime('%Y%m%d_%H%M%S')
        else:
            text = ''

        try:
            pyperclip.copy(text)
            time.sleep(DEFAULT_TYPING_DELAY)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(DEFAULT_PASTE_DELAY)

            if var_type == 'excel' and text:
                self.excel_mgr.add_pasted_value(excel_id, text)

        except Exception as e:
            self.log(f"    변수 타이핑 오류: {e}")
            raise Exception(f"변수 타이핑 실패: {str(e)}")

    def action_key_press(self, params):
        """키 입력"""
        key = params.get('key', '')

        if '+' in key:
            keys = key.split('+')
            pyautogui.hotkey(*keys)
        else:
            pyautogui.press(key)

        time.sleep(DEFAULT_POST_DELAY)

    def action_hotkey(self, params):
        """단축키"""
        keys = params.get('keys', [])
        pyautogui.hotkey(*keys)
        time.sleep(DEFAULT_POST_DELAY)

    def action_paste(self):
        """붙여넣기"""
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(DEFAULT_POST_DELAY)

    def action_mouse_scroll(self, params):
        """마우스 스크롤"""
        direction = params.get('direction', 'down')
        amount = params.get('amount', 3)

        scroll_amount = amount * 100 if direction == 'up' else -amount * 100

        pyautogui.scroll(scroll_amount)
        time.sleep(DEFAULT_POST_DELAY)

    def action_delay(self, params):
        """대기"""
        seconds = params.get('seconds', 1)
        time.sleep(seconds)

    def action_wait_image(self, params):
        """이미지가 나타날 때까지 대기 (OpenCV 기반, 캐시 사용)"""
        image_id = params.get('image_id')
        timeout = params.get('timeout', DEFAULT_IMAGE_TIMEOUT)

        image = self.image_mgr.get_image(image_id)
        if not image:
            raise Exception(f"이미지 ID {image_id}를 찾을 수 없습니다.")

        self.log(f"   이미지 '{image['name']}' 대기 중... (최대 {timeout}초)")

        try:
            cache = self._get_cached_image_path(image)
            template_gray = cache['gray']
            w, h = template_gray.shape[::-1]

            start_time = time.time()
            confidence_threshold = image.get('confidence', DEFAULT_WAIT_IMAGE_CONFIDENCE)

            while time.time() - start_time < timeout:
                if self.should_stop:
                    raise Exception("사용자가 중지했습니다.")

                screen = np.array(ImageGrab.grab())
                screen_gray = cv2.cvtColor(screen, cv2.COLOR_BGR2GRAY)

                res = cv2.matchTemplate(screen_gray, template_gray, cv2.TM_CCOEFF_NORMED)
                _, max_val, _, max_loc = cv2.minMaxLoc(res)

                if max_val >= confidence_threshold:
                    center_x = max_loc[0] + w // 2
                    center_y = max_loc[1] + h // 2
                    elapsed = time.time() - start_time
                    self.log(f"   이미지 발견! ({center_x}, {center_y}) - {elapsed:.1f}초 소요")
                    return

                time.sleep(DEFAULT_IMAGE_POLL_INTERVAL)

            raise Exception(f"이미지 '{image['name']}'을(를) {timeout}초 내에 찾을 수 없습니다.")

        except Exception as e:
            raise Exception(f"이미지 대기 오류: {str(e)}")

    def action_screenshot(self, params):
        """스크린샷 저장"""
        base_filename = params.get('filename', 'screenshot')

        if base_filename.endswith('.png'):
            base_filename = base_filename[:-4]

        filename = f"{base_filename}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"

        if self.project_filepath:
            project_dir = os.path.dirname(self.project_filepath)
            screenshot_dir = os.path.join(project_dir, 'logs', 'screenshots')
        else:
            screenshot_dir = os.path.join('logs', 'screenshots')

        os.makedirs(screenshot_dir, exist_ok=True)

        filepath = os.path.join(screenshot_dir, filename)

        mode = params.get('mode', 'full')
        region = params.get('region')

        if mode == 'region' and region:
            screenshot = pyautogui.screenshot(region=(
                region['x'],
                region['y'],
                region['width'],
                region['height']
            ))
            self.log(f"    영역 스크린샷: ({region['x']}, {region['y']}, {region['width']}x{region['height']})")
        else:
            screenshot = pyautogui.screenshot()
            self.log(f"    전체 화면 스크린샷")

        screenshot.save(filepath)
        self.log(f"    저장 위치: {filepath}")
