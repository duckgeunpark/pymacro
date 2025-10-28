"""
플로우 시퀀스 관리
"""

class FlowManager:
    """플로우 관리 클래스"""
    
    def __init__(self):
        self.flow_sequence = []
        self.next_id = 1
    
    def add_action(self, action_type, params):
        """액션 추가
        
        Args:
            action_type: 액션 타입 (click, type, delay, paste, wait_image 등)
            params: 액션별 파라미터
        """
        action = {
            'id': self.next_id,
            'type': action_type,
            'params': params
        }
        self.flow_sequence.append(action)
        self.next_id += 1
        return action
    
    def remove_action(self, action_id):
        """액션 삭제"""
        self.flow_sequence = [a for a in self.flow_sequence if a['id'] != action_id]
    
    def move_action_up(self, action_id):
        """액션 위로 이동"""
        idx = self.get_action_index(action_id)
        if idx is not None and idx > 0:
            self.flow_sequence[idx], self.flow_sequence[idx-1] = \
                self.flow_sequence[idx-1], self.flow_sequence[idx]
            return True
        return False
    
    def move_action_down(self, action_id):
        """액션 아래로 이동"""
        idx = self.get_action_index(action_id)
        if idx is not None and idx < len(self.flow_sequence) - 1:
            self.flow_sequence[idx], self.flow_sequence[idx+1] = \
                self.flow_sequence[idx+1], self.flow_sequence[idx]
            return True
        return False
    
    def get_action(self, action_id):
        """ID로 액션 찾기"""
        for action in self.flow_sequence:
            if action['id'] == action_id:
                return action
        return None
    
    def get_action_index(self, action_id):
        """액션의 인덱스 찾기"""
        for idx, action in enumerate(self.flow_sequence):
            if action['id'] == action_id:
                return idx
        return None
    
    def update_action(self, action_id, params):
        """액션 업데이트"""
        action = self.get_action(action_id)
        if action:
            action['params'].update(params)
            return True
        return False
    
    def load_from_list(self, action_list):
        """리스트에서 플로우 로드"""
        self.flow_sequence = action_list
        if action_list:
            self.next_id = max(a['id'] for a in action_list) + 1
        else:
            self.next_id = 1
    
    def to_list(self):
        """리스트로 변환"""
        return self.flow_sequence
    
    @staticmethod
    def get_action_display_text(action, coord_mgr=None, excel_mgr=None, image_mgr=None):
        """액션을 사람이 읽을 수 있는 텍스트로 변환"""
        action_type = action['type']
        params = action['params']
        
        if action_type == 'click_coord':
            coord_id = params.get('coord_id')
            coord_name = "알 수 없음"
            if coord_mgr:
                coord = coord_mgr.get_coordinate(coord_id)
                if coord:
                    coord_name = coord['name']
            click_type = params.get('click_type', 'left')
            return f"[좌표:{coord_name}] {click_type} 클릭"
        
        elif action_type == 'click_image':
            image_id = params.get('image_id')
            image_name = "알 수 없음"
            if image_mgr:
                img = image_mgr.get_image(image_id)
                if img:
                    image_name = img['name']
            return f"[이미지:{image_name}] 클릭"
        
        elif action_type == 'type_text':
            text = params.get('text', '')
            return f"[타이핑] {text[:30]}..."
        
        elif action_type == 'type_variable':
            var_type = params.get('var_type')  # excel, counter, timestamp
            var_name = params.get('var_name', '')
            return f"[타이핑] {{{var_type}:{var_name}}}"
        
        elif action_type == 'key_press':
            key = params.get('key', '')
            return f"[키입력] {key}"
        
        elif action_type == 'hotkey':
            keys = params.get('keys', [])
            return f"[단축키] {'+'.join(keys)}"
        
        elif action_type == 'paste':
            return f"[붙여넣기] Ctrl+V"
        
        elif action_type == 'delay':
            seconds = params.get('seconds', 0)
            return f"[대기] {seconds}초"
        
        elif action_type == 'wait_image':
            image_id = params.get('image_id')
            timeout = params.get('timeout', 10)
            image_name = "알 수 없음"
            if image_mgr:
                img = image_mgr.get_image(image_id)
                if img:
                    image_name = img['name']
            return f"[이미지 대기] {image_name} (최대 {timeout}초)"
        
        elif action_type == 'screenshot':
            filename = params.get('filename', 'screenshot.png')
            return f"[스크린샷] {filename}"
        
        elif action_type == 'memo':
            text = params.get('text', '')
            return f"[메모] {text[:50]}..."
        
        else:
            return f"[알 수 없는 액션] {action_type}"
