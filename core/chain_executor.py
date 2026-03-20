"""
매크로 체인 실행 엔진
"""
import json
from datetime import datetime
from core.executor import MacroExecutor
from core.coordinate_manager import CoordinateManager
from core.excel_manager import ExcelManager
from core.image_manager import ImageManager
from core.flow_manager import FlowManager


class ChainExecutor:
    """매크로 체인 실행 엔진"""

    def __init__(self, chain_items):
        """
        Args:
            chain_items: [{'project_name': str, 'filepath': str, 'repeat_count': int}, ...]
        """
        self.chain_items = chain_items
        self.is_running = False
        self.should_stop = False
        self.is_paused = False

        self.current_chain_index = 0
        self.current_repeat = 0

        # 현재 실행 중인 executor 저장
        self.current_executor = None

        # 콜백
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

    def report_error(self, error_msg):
        """에러 보고"""
        self.log(f"[ERR] 에러: {error_msg}")
        if self.error_callback:
            self.error_callback(error_msg)

    def start(self):
        """체인 실행 시작"""
        self.is_running = True
        self.should_stop = False

        # 총 실행 횟수 계산 (프로젝트 설정 사용은 1회로 카운트)
        total_executions = 0
        for item in self.chain_items:
            if item.get('execution_mode') == 'use_project_settings':
                total_executions += 1  # 프로젝트 설정 사용은 1회로 카운트
            else:
                total_executions += item.get('repeat_count', 1)

        current_execution = 0

        self.log("===== 매크로 체인 실행 시작 =====")
        self.log(f"총 {len(self.chain_items)}개의 매크로 실행 예정")

        try:
            for chain_idx, item in enumerate(self.chain_items):
                if self.should_stop:
                    self.log("[STOP] 사용자가 중지했습니다.")
                    break

                self.current_chain_index = chain_idx
                project_name = item['project_name']
                filepath = item['filepath']
                execution_mode = item.get('execution_mode', 'custom_repeat')
                repeat_count = item.get('repeat_count', 1)

                # 로그 메시지 생성
                if execution_mode == 'use_project_settings':
                    self.log(f"\n[{chain_idx + 1}/{len(self.chain_items)}] '{project_name}' 실행 (프로젝트 설정 사용)")
                else:
                    self.log(f"\n[{chain_idx + 1}/{len(self.chain_items)}] '{project_name}' 실행 (총 {repeat_count}회)")

                # 프로젝트 로드
                try:
                    with open(filepath, 'r', encoding='utf-8') as f:
                        project_data = json.load(f)
                except Exception as e:
                    self.report_error(f"프로젝트 '{project_name}' 로드 실패: {str(e)}")
                    continue

                # 관리자 초기화
                coord_mgr = CoordinateManager()
                coord_mgr.load_from_list(project_data.get('coordinates', []))

                excel_mgr = ExcelManager()
                excel_mgr.load_from_list(project_data.get('excel_sources', []))

                image_mgr = ImageManager()
                image_mgr.load_from_list(project_data.get('images', []))

                flow_mgr = FlowManager()
                flow_mgr.load_from_list(project_data.get('flow_sequence', []))

                # 실행 모드 확인
                execution_mode = item.get('execution_mode', 'custom_repeat')

                if execution_mode == 'use_project_settings':
                    # 프로젝트 설정 사용
                    self.log(f"  프로젝트 실행 설정 사용")

                    # 매크로 실행기 생성
                    self.current_executor = MacroExecutor(
                        project_data,
                        coord_mgr,
                        excel_mgr,
                        image_mgr,
                        flow_mgr,
                        filepath
                    )

                    # 콜백 전달
                    self.current_executor.set_callbacks(
                        log_cb=self.log_callback,
                        progress_cb=None,  # 체인의 진행률로 대체
                        error_cb=self.error_callback
                    )

                    try:
                        # 프로젝트 설정에 따라 실행
                        self.current_executor.start()
                        current_execution += 1
                    except Exception as e:
                        self.report_error(f"실행 중 오류: {str(e)}")
                    finally:
                        self.current_executor = None

                else:
                    # 반복 횟수 직접 지정
                    for repeat_idx in range(repeat_count):
                        if self.should_stop:
                            break

                        # 일시정지 체크
                        while self.is_paused and not self.should_stop:
                            import time
                            time.sleep(0.1)

                        self.current_repeat = repeat_idx + 1
                        current_execution += 1

                        self.log(f"  반복 {repeat_idx + 1}/{repeat_count}")
                        self.update_progress(
                            current_execution,
                            total_executions,
                            f"'{project_name}' {repeat_idx + 1}/{repeat_count} 실행 중"
                        )

                        # 매크로 실행
                        self.current_executor = MacroExecutor(
                            project_data,
                            coord_mgr,
                            excel_mgr,
                            image_mgr,
                            flow_mgr,
                            filepath
                        )

                        # 콜백 전달
                        self.current_executor.set_callbacks(
                            log_cb=self.log_callback,
                            progress_cb=None,  # 체인의 진행률로 대체
                            error_cb=self.error_callback
                        )

                        try:
                            # 단순 플로우 1회 실행
                            self.current_executor.execute_flow()
                        except Exception as e:
                            self.report_error(f"실행 중 오류: {str(e)}")
                            # 계속 진행
                        finally:
                            self.current_executor = None

                self.log(f"[OK] '{project_name}' 완료")

            if not self.should_stop:
                self.log("\n===== 매크로 체인 실행 완료 =====")
            else:
                self.log("\n[STOP] ===== 매크로 체인 중지됨 =====")

        except Exception as e:
            self.report_error(f"체인 실행 중 오류 발생: {str(e)}")

        finally:
            self.is_running = False

    def pause(self):
        """일시정지"""
        self.is_paused = True
        self.log("[PAUSE] 일시정지")

    def resume(self):
        """재개"""
        self.is_paused = False
        self.log("[PLAY] 재개")

    def stop(self):
        """중지"""
        self.should_stop = True
        self.log("[STOP] 중지 요청")

        # 현재 실행 중인 executor도 중지
        if self.current_executor:
            self.current_executor.stop()
