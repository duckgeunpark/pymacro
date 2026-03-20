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


def _get_root_window(widget):
    """최상위 앱 윈도우를 찾아 반환"""
    root = widget.winfo_toplevel()
    # CTkToplevel의 경우 master를 따라가며 메인 윈도우 찾기
    while hasattr(root, 'master') and root.master is not None:
        parent = root.master
        try:
            parent_top = parent.winfo_toplevel()
            if parent_top == root:
                break
            root = parent_top
        except Exception:
            break
    return root


def center_window_on_parent(dialog, parent):
    """다이얼로그를 앱 메인 윈도우 중앙에 배치"""
    dialog.update_idletasks()

    # 앱 메인 윈도우 기준으로 센터링
    app_window = _get_root_window(parent)

    parent_x = app_window.winfo_x()
    parent_y = app_window.winfo_y()
    parent_width = app_window.winfo_width()
    parent_height = app_window.winfo_height()

    # 다이얼로그 크기
    dialog_width = dialog.winfo_width()
    dialog_height = dialog.winfo_height()

    # 중앙 위치 계산
    x = parent_x + (parent_width - dialog_width) // 2
    y = parent_y + (parent_height - dialog_height) // 2

    # 화면 밖으로 나가지 않도록 보정
    screen_w = dialog.winfo_screenwidth()
    screen_h = dialog.winfo_screenheight()
    x = max(0, min(x, screen_w - dialog_width))
    y = max(0, min(y, screen_h - dialog_height))

    dialog.geometry(f'{dialog_width}x{dialog_height}+{x}+{y}')


def center_window_on_screen(window):
    """창을 화면 중앙에 배치"""
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f'{width}x{height}+{x}+{y}')


def setup_modal(dialog, parent, width=None, height=None):
    """CTkToplevel 모달 다이얼로그 공통 설정"""
    dialog.transient(parent)
    dialog.grab_set()
    dialog.attributes('-topmost', True)
    # 초기에 화면 밖에 배치하여 깜빡임 방지
    dialog.withdraw()
    set_dialog_icon(dialog)
    if width and height:
        dialog.geometry(f"{width}x{height}")

    def _show_centered():
        center_window_on_parent(dialog, parent)
        dialog.deiconify()
        dialog.lift()
        dialog.focus_force()

    dialog.after(50, _show_centered)