"""
MacroBuilder Design System - Clean Light Theme
"""

import customtkinter as ctk

# 기본 모드 설정
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class Colors:
    # ─── Brand ───
    PRIMARY = "#2563eb"        # Clean blue
    PRIMARY_HOVER = "#1d4ed8"
    PRIMARY_LIGHT = "#3b82f6"
    PRIMARY_SURFACE = "#eff6ff"

    # ─── Accent ───
    ACCENT = "#0ea5e9"         # Sky blue
    ACCENT_HOVER = "#0284c7"
    SUCCESS = "#16a34a"
    SUCCESS_HOVER = "#15803d"
    WARNING = "#ca8a04"
    WARNING_HOVER = "#a16207"
    DANGER = "#dc2626"
    DANGER_HOVER = "#b91c1c"
    PURPLE = "#7c3aed"
    PURPLE_HOVER = "#6d28d9"
    ROSE = "#e11d48"

    # ─── Neutral ───
    GRAY_50 = "#f8fafc"
    GRAY_100 = "#f1f5f9"
    GRAY_200 = "#e2e8f0"
    GRAY_300 = "#cbd5e1"
    GRAY_400 = "#94a3b8"
    GRAY_500 = "#64748b"
    GRAY_600 = "#475569"
    GRAY_700 = "#334155"
    GRAY_800 = "#1e293b"
    GRAY_900 = "#0f172a"
    GRAY_950 = "#020617"

    # ─── Surfaces ───
    BG_PRIMARY = "#ffffff"     # White background
    BG_SECONDARY = "#f8fafc"   # Light gray
    BG_CARD = "#ffffff"        # Card surface
    BG_CARD_HOVER = "#f1f5f9"  # Card hover
    BG_ELEVATED = "#f1f5f9"    # Elevated panels
    BG_INPUT = "#ffffff"       # Input fields
    BG_INPUT_FOCUS = "#f8fafc"

    # ─── Text ───
    TEXT_PRIMARY = "#1e293b"
    TEXT_SECONDARY = "#475569"
    TEXT_MUTED = "#94a3b8"
    TEXT_INVERSE = "#ffffff"

    # ─── Border ───
    BORDER = "#e2e8f0"
    BORDER_SUBTLE = "#cbd5e1"
    BORDER_FOCUS = "#2563eb"

    # ─── Status ───
    AUTOSTART_BG = "#f0fdf4"
    AUTOSTART_FG = "#16a34a"
    AUTOSTART_BORDER = "#86efac"

    # ─── Action Type Colors ───
    ACTION_MOUSE = "#2563eb"
    ACTION_KEYBOARD = "#16a34a"
    ACTION_CONTROL = "#ea580c"
    ACTION_OTHER = "#7c3aed"
    ACTION_DEFAULT = "#94a3b8"

    # ─── Gradients (for labels/badges) ───
    BADGE_BLUE = "#2563eb"
    BADGE_GREEN = "#16a34a"
    BADGE_ORANGE = "#ea580c"
    BADGE_PURPLE = "#7c3aed"

    # ─── Drag & Drop ───
    DRAG_INDICATOR = "#2563eb"
    DROP_TARGET = "#1d4ed8"
    DROP_TARGET_BG = "#eff6ff"


class Fonts:
    FAMILY = "Pretendard"
    FAMILY_FALLBACK = "맑은 고딕"
    MONO_FAMILY = "JetBrains Mono"
    MONO_FALLBACK = "Consolas"

    # ─── Scale ───
    DISPLAY = (FAMILY_FALLBACK, 28, "bold")
    TITLE = (FAMILY_FALLBACK, 22, "bold")
    HEADING = (FAMILY_FALLBACK, 18, "bold")
    SUBHEADING = (FAMILY_FALLBACK, 16, "bold")
    BODY = (FAMILY_FALLBACK, 14)
    BODY_BOLD = (FAMILY_FALLBACK, 14, "bold")
    BODY_LG = (FAMILY_FALLBACK, 15)
    SMALL = (FAMILY_FALLBACK, 13)
    SMALL_BOLD = (FAMILY_FALLBACK, 13, "bold")
    CAPTION = (FAMILY_FALLBACK, 12)
    CAPTION_BOLD = (FAMILY_FALLBACK, 12, "bold")
    TINY = (FAMILY_FALLBACK, 11)
    MONO = (MONO_FALLBACK, 13)
    MONO_SMALL = (MONO_FALLBACK, 12)
    COUNTER = (FAMILY_FALLBACK, 56, "bold")


class Sizes:
    # ─── Spacing ───
    PAD_XS = 4
    PAD_SM = 8
    PAD_MD = 12
    PAD_LG = 16
    PAD_XL = 24
    PAD_2XL = 32

    # ─── Radius ───
    RADIUS_SM = 6
    RADIUS_MD = 10
    RADIUS_LG = 14
    RADIUS_XL = 20

    # ─── Button ───
    BTN_HEIGHT = 36
    BTN_HEIGHT_LG = 42
    BTN_HEIGHT_SM = 28
    BTN_HEIGHT_XS = 24

    # ─── Header ───
    HEADER_HEIGHT = 56

    # ─── Sidebar ───
    SIDEBAR_WIDTH = 200

    # ─── Window ───
    WINDOW_WIDTH = 520
    WINDOW_HEIGHT = 600


class Styles:
    """Pre-built widget style presets"""

    @staticmethod
    def card(parent, **kwargs):
        defaults = dict(
            fg_color=Colors.BG_CARD,
            corner_radius=Sizes.RADIUS_LG,
            border_width=1,
            border_color=Colors.BORDER,
        )
        defaults.update(kwargs)
        return ctk.CTkFrame(parent, **defaults)

    @staticmethod
    def section_title(parent, text, **kwargs):
        defaults = dict(
            text=text,
            font=Fonts.CAPTION_BOLD,
            text_color=Colors.TEXT_MUTED,
            anchor="w",
        )
        defaults.update(kwargs)
        return ctk.CTkLabel(parent, **defaults)

    @staticmethod
    def primary_button(parent, text, command=None, **kwargs):
        defaults = dict(
            text=text,
            font=Fonts.BODY_BOLD,
            fg_color=Colors.PRIMARY,
            hover_color=Colors.PRIMARY_HOVER,
            text_color=Colors.TEXT_INVERSE,
            height=Sizes.BTN_HEIGHT_LG,
            corner_radius=Sizes.RADIUS_MD,
            command=command,
        )
        defaults.update(kwargs)
        return ctk.CTkButton(parent, **defaults)

    @staticmethod
    def secondary_button(parent, text, command=None, **kwargs):
        defaults = dict(
            text=text,
            font=Fonts.SMALL_BOLD,
            fg_color=Colors.BG_ELEVATED,
            hover_color=Colors.GRAY_300,
            text_color=Colors.TEXT_SECONDARY,
            height=Sizes.BTN_HEIGHT,
            corner_radius=Sizes.RADIUS_SM,
            command=command,
        )
        defaults.update(kwargs)
        return ctk.CTkButton(parent, **defaults)

    @staticmethod
    def ghost_button(parent, text, command=None, **kwargs):
        defaults = dict(
            text=text,
            font=Fonts.SMALL,
            fg_color="transparent",
            hover_color=Colors.BG_ELEVATED,
            text_color=Colors.TEXT_SECONDARY,
            height=Sizes.BTN_HEIGHT_SM,
            corner_radius=Sizes.RADIUS_SM,
            command=command,
        )
        defaults.update(kwargs)
        return ctk.CTkButton(parent, **defaults)

    @staticmethod
    def danger_button(parent, text, command=None, **kwargs):
        defaults = dict(
            text=text,
            font=Fonts.SMALL_BOLD,
            fg_color=Colors.DANGER,
            hover_color=Colors.DANGER_HOVER,
            text_color=Colors.TEXT_INVERSE,
            height=Sizes.BTN_HEIGHT,
            corner_radius=Sizes.RADIUS_SM,
            command=command,
        )
        defaults.update(kwargs)
        return ctk.CTkButton(parent, **defaults)

    @staticmethod
    def input_field(parent, **kwargs):
        defaults = dict(
            font=Fonts.BODY,
            fg_color=Colors.BG_INPUT,
            text_color=Colors.TEXT_PRIMARY,
            placeholder_text_color=Colors.TEXT_MUTED,
            border_color=Colors.BORDER_SUBTLE,
            border_width=1,
            corner_radius=Sizes.RADIUS_SM,
            height=Sizes.BTN_HEIGHT,
        )
        defaults.update(kwargs)
        return ctk.CTkEntry(parent, **defaults)

    @staticmethod
    def combobox(parent, variable, values, **kwargs):
        defaults = dict(
            variable=variable,
            values=values,
            font=Fonts.BODY,
            dropdown_font=Fonts.BODY,
            state="readonly",
            fg_color=Colors.BG_INPUT,
            border_color=Colors.BORDER_SUBTLE,
            border_width=1,
            button_color=Colors.GRAY_200,
            button_hover_color=Colors.GRAY_300,
            dropdown_fg_color=Colors.BG_CARD,
            dropdown_hover_color=Colors.PRIMARY_SURFACE,
            dropdown_text_color=Colors.TEXT_PRIMARY,
            text_color=Colors.TEXT_PRIMARY,
            corner_radius=Sizes.RADIUS_SM,
            height=Sizes.BTN_HEIGHT,
        )
        defaults.update(kwargs)
        return ctk.CTkComboBox(parent, **defaults)
