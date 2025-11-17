"""
UI 관련 유틸리티 함수들
"""
from core.config import config


def set_dialog_icon(dialog):
    """다이얼로그에 아이콘 설정하는 공통 함수"""
    try:
        icon_path = config.icon_path
        if icon_path:
            dialog.iconbitmap(icon_path)
    except Exception:
        pass  # 조용히 실패


def center_window_on_parent(dialog, parent):
    """다이얼로그를 부모 창 중앙에 배치"""
    dialog.update_idletasks()

    # 부모 창의 위치와 크기
    parent_x = parent.winfo_x()
    parent_y = parent.winfo_y()
    parent_width = parent.winfo_width()
    parent_height = parent.winfo_height()

    # 다이얼로그 크기
    dialog_width = dialog.winfo_width()
    dialog_height = dialog.winfo_height()

    # 중앙 위치 계산
    x = parent_x + (parent_width - dialog_width) // 2
    y = parent_y + (parent_height - dialog_height) // 2

    dialog.geometry(f'{dialog_width}x{dialog_height}+{x}+{y}')


def center_window_on_screen(window):
    """창을 화면 중앙에 배치"""
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f'{width}x{height}+{x}+{y}')