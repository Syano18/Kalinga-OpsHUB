import threading, os, socket, sys, subprocess, tempfile, csv, json, certifi
from datetime import datetime
import importlib.util
from dotenv import load_dotenv

# ── Kivy bootstrap ────────────────────────────────────────────────────────────
from kivy.config import Config
Config.set('graphics', 'allow_high_dpi', '1')
Config.set('graphics', 'resizable', '1')
Config.set('input', 'mouse', 'mouse,disable_multitouch')
Config.set('graphics', 'minimum_width', '900')
Config.set('graphics', 'minimum_height', '600')
Config.set('graphics', 'width', '1000')
Config.set('graphics', 'height', '600')
Config.set('graphics', 'multisamples', '8')

from kivy.app import App
from kivy.lang import Builder
from kivy.uix.screenmanager import ScreenManager, Screen, FadeTransition
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.relativelayout import RelativeLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.anchorlayout import AnchorLayout
from kivy.uix.scrollview import ScrollView
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.uix.popup import Popup
from kivy.uix.image import Image as KivyImage
from kivy.uix.progressbar import ProgressBar
from kivy.uix.checkbox import CheckBox
from kivy.uix.widget import Widget
from kivy.clock import Clock
from kivy.core.window import Window
from kivy.metrics import dp, sp
from kivy.properties import StringProperty, BooleanProperty, ListProperty, NumericProperty
from kivy.graphics import Color, Rectangle, RoundedRectangle, Line, Ellipse
from kivy.utils import get_color_from_hex
from kivy.animation import Animation
from kivy.core.text import LabelBase

# ── Environment & Path Setup ──────────────────────────────────────────────────
def resource_path(p):
    """Get absolute path to resource, works for dev and for PyInstaller."""
    base = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable)
                   if getattr(sys, 'frozen', False) else os.path.abspath('.'))
    return os.path.join(base, p)

# Fix for SSL: CERTIFICATE_VERIFY_FAILED on Windows when packaged
os.environ['SSL_CERT_FILE'] = certifi.where()
os.environ['REQUESTS_CA_BUNDLE'] = certifi.where()

# 1. Try to load .env from bundled resources (internal) and current folder
load_dotenv(resource_path(".env"))
# 2. Also try loading from the directory where the .exe is located (external)
if getattr(sys, 'frozen', False):
    load_dotenv(os.path.join(os.path.dirname(sys.executable), ".env"))

# ── Feature flags ─────────────────────────────────────────────────────────────
GOOGLE_LOGIN_AVAILABLE = bool(
    importlib.util.find_spec("google_auth_oauthlib") and
    importlib.util.find_spec("google.oauth2")
)

# ── App constants (loaded after environment is ready) ─────────────────────────
TURSO_DB_URL        = os.getenv("TURSO_DB_URL")
TURSO_AUTH_TOKEN    = os.getenv("TURSO_AUTH_TOKEN")
FIREBASE_WEB_API_KEY= os.getenv("FIREBASE_WEB_API_KEY")
GITHUB_OWNER        = "Syano18"
GITHUB_REPO         = "Kalinga-OpsHUB"
CURRENT_VERSION     = "2.4"
SCOPES              = ['https://www.googleapis.com/auth/userinfo.email', 'openid']

USER_HEADERS = ["Email","Role","First Name","Middle Initial","Last Name",
                "Position","Salary","Salary Grade","Status"]
LOG_HEADERS  = ["Timestamp","Reference Number","Particulars","Addressee",
                "Transmitter","Section","Mode","Remarks","Encoded By"]

ATT_HEADERS = ["Full Name", "Date", "Time In AM", "Time Out AM", "Time In PM", "Time Out PM", "Notes"]
ATT_COL_WIDTHS = [dp(250), dp(120), dp(120), dp(120), dp(120), dp(120), dp(200)]
try:
    _bd = os.path.join(os.getenv("LOCALAPPDATA", os.path.expanduser("~")), "KalingaOpsHub")
    os.makedirs(_bd, exist_ok=True)
    SESSION_FILE     = os.path.join(_bd, "session.txt")
    GOOGLE_TOKEN_FILE= os.path.join(_bd, "token.json")
    CACHE_FILE       = os.path.join(_bd, "cache.json")
    THEME_FILE       = os.path.join(_bd, "theme.txt")
except Exception:
    SESSION_FILE = "user_session.txt"; GOOGLE_TOKEN_FILE = "token.json"; CACHE_FILE = "cache.json"
    THEME_FILE = "theme.txt"

LOGO_PATH   = resource_path("assets/Logo.png")
AGENCY_LOGO = resource_path("assets/PSA.png")

# ── Register Custom Emoji Font ────────────────────────────────────────────────
try:
    LabelBase.register(name='Emoji', fn_regular=resource_path('assets/emoji.ttf'))
except Exception as e:
    print("Could not load emoji font:", e)

# ── Palette (Light Mode) ──────────────────────────────────────────────────────
P_LIGHT = {
    "primary"   : get_color_from_hex("#2563eb"),   # Indigo 600
    "primary_d" : get_color_from_hex("#4338CA"),   # Indigo 700
    "primary_l" : get_color_from_hex("#EEF2FF"),   # Indigo 50
    "bg"        : get_color_from_hex("#F8FAFC"),   # Slate 50
    "sidebar"   : get_color_from_hex("#FFFFFF"),   # White
    "sidebar_l" : get_color_from_hex("#F1F5F9"),   # Slate 100
    "card"      : get_color_from_hex("#FFFFFF"),   # White
    "text"      : get_color_from_hex("#020202"),   # Slate 800
    "subtext"   : get_color_from_hex("#64748B"),   # Slate 500
    "success"   : get_color_from_hex("#16A34A"),   # Green 600
    "error"     : get_color_from_hex("#DC2626"),   # Red 600
    "warning"   : get_color_from_hex("#D97706"),   # Amber 600
    "border"    : get_color_from_hex("#94A3B8"),   # Slate 400
    "nav_item"  : get_color_from_hex("#64748B"),   # Slate 500
    "nav_active": get_color_from_hex("#EEF2FF"),   # Indigo 50
    "row_alt"   : get_color_from_hex("#F8FAFC"),   # Slate 50
    "selected"  : get_color_from_hex("#E0E7FF"),   # Indigo 100
    "header_bg" : get_color_from_hex("#ECF0F4"),   # Slate 100
    "google"    : get_color_from_hex("#EA4335"),
    "white"     : [1,1,1,1],
    "input_bg"  : get_color_from_hex("#FFFFFF"),   # White
    "disabled"  : get_color_from_hex("#D1D5DB"),   # Gray 300
    "overlay"   : [0, 0, 0, 0.5],
    "transparent": [0,0,0,0],
}

# Dark theme palette
# Dark theme palette
P_DARK = {
    "primary"   : get_color_from_hex("#2563eb"),   
    "primary_d" : get_color_from_hex("#4F46E5"),
    "primary_l" : get_color_from_hex("#0B1220"),
    "bg"        : get_color_from_hex("#111827"),   # Updated: bg-gray-900
    "sidebar"   : get_color_from_hex("#1f2937"),   # Updated: bg-gray-800
    "sidebar_l" : get_color_from_hex("#111827"),   # Updated: bg-gray-900
    "card"      : get_color_from_hex("#1f2937"),   # Updated: bg-gray-800
    "text"      : get_color_from_hex("#FFFFFF"),
    "subtext"   : get_color_from_hex("#9CA3AF"),   # text-gray-400
    "success"   : get_color_from_hex("#16A34A"),
    "error"     : get_color_from_hex("#EF4444"),
    "warning"   : get_color_from_hex("#F59E0B"),
    "border"    : get_color_from_hex("#374151"),   # border-gray-700
    "nav_item"  : get_color_from_hex("#9CA3AF"),   # text-gray-400
    "nav_active": get_color_from_hex("#111827"),   # bg-gray-900
    "row_alt"   : get_color_from_hex("#111827"),   # bg-gray-900
    "selected"  : get_color_from_hex("#374151"),   
    "header_bg" : get_color_from_hex("#18223780"),   # bg-gray-900
    "google"    : get_color_from_hex("#EA4335"),
    "white"     : [1,1,1,1],
    "input_bg"  : get_color_from_hex("#111827"),   # bg-gray-900 inside inputs
    "disabled"  : get_color_from_hex("#374151"),   # bg-gray-700
    "overlay"   : [0, 0, 0, 0.7],
    "transparent": [0,0,0,0],
}

P = P_LIGHT.copy()

CURRENT_THEME = "light"

def apply_theme(name="light"):
    global P, CURRENT_THEME
    CURRENT_THEME = name
    base = P_LIGHT if name == 'light' else P_DARK
    for k,v in base.items(): P[k] = v
    try: Window.clearcolor = list(P["bg"])
    except: pass
    try:
        app = App.get_running_app()
        if app and hasattr(app, 'sm'):
            for s in list(app.sm.screens):
                fn = getattr(s, '_apply_theme', None)
                if callable(fn):
                    try: fn()
                    except: pass
    except: pass

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  KV style definitions
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Builder.load_string("""
#:import dp kivy.metrics.dp
#:import sp kivy.metrics.sp
#:import get_color_from_hex kivy.utils.get_color_from_hex

<RoundBtn>:
    markup: True
    background_color: 0,0,0,0
    background_normal: ''
    color: 1,1,1,1
    font_size: sp(13)
    bold: True
    size_hint_y: None
    height: dp(44)
    canvas.before:
        Color:
            rgba: (self.btn_color if not self.hover else [min(1, c*0.85) for c in self.btn_color[:3]] + [1]) if not self.disabled else self.disabled_bg_color
        RoundedRectangle:
            pos: self.pos
            size: self.size
            radius: [dp(8)]

<GhostBtn>:
    markup: True
    background_color: 0,0,0,0
    background_normal: ''
    font_size: sp(12)
    size_hint_y: None
    height: dp(38)

<NavBtn>:
    markup: True
    background_color: 0,0,0,0
    background_normal: ''
    text_size: (self.width - dp(32), None)
    halign: 'left'
    valign: 'middle'
    font_size: sp(15)
    size_hint_y: None
    height: dp(50)
    padding: [dp(16), 0]
    canvas.before:
        Color:
            rgba: self.active_color if self.is_active else (self.hover_color if self.hover else [0,0,0,0])
        RoundedRectangle:
            pos: self.pos
            size: self.size
            radius: [dp(8)]

<CardBox>:
    canvas.before:
        Color:
            rgba: self.bg_color
        RoundedRectangle:
            pos: self.pos
            size: self.size
            radius: [dp(16)]
        Color:
            rgba: self.border_color
        Line:
            rounded_rectangle: self.x, self.y, self.width, self.height, dp(16)
            width: dp(1)

<SidebarBox>:
    canvas.before:
        Color:
            rgba: self.bg_color
        Rectangle:
            pos: self.pos
            size: self.size

<UserBox>:
    canvas.before:
        Color:
            rgba: self.bg_color
        RoundedRectangle:
            pos: self.pos
            size: self.size
            radius: [dp(10)]

<BgBox>:
    canvas.before:
        Color:
            rgba: self.bg_color
        Rectangle:
            pos: self.pos
            size: self.size
""")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  KV-declared widget classes
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

class RoundBtn(Button):
    btn_color = ListProperty([0,0,0,1])
    disabled_bg_color = ListProperty([0.8,0.8,0.8,1])
    role = StringProperty("primary")
    text_role = StringProperty(None, allownone=True)
    hover = BooleanProperty(False)
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.refresh_theme()
        self.bind(parent=self._bind_mouse)
    def _bind_mouse(self, instance, parent):
        if parent: Window.bind(mouse_pos=self._on_mouse_pos)
        else: Window.unbind(mouse_pos=self._on_mouse_pos)
    def _on_mouse_pos(self, instance, pos):
        if not self.get_root_window(): return
        self.hover = self.collide_point(*self.to_widget(*pos))
    def refresh_theme(self):
        if self.role in P: self.btn_color = list(P[self.role])
        if self.text_role and self.text_role in P:
            self.color = list(P[self.text_role])
        self.disabled_bg_color = list(P.get("disabled", [0.8,0.8,0.8,1]))
        self.disabled_color = [1,1,1,1]

class GhostBtn(Button): pass

class NavBtn(Button):
    is_active   = BooleanProperty(False)
    hover       = BooleanProperty(False)
    active_color= ListProperty([0,0,0,0])
    hover_color = ListProperty([0,0,0,0])
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.refresh_theme()
        self.bind(parent=self._bind_mouse)
    def _bind_mouse(self, instance, parent):
        if parent: Window.bind(mouse_pos=self._on_mouse_pos)
        else: Window.unbind(mouse_pos=self._on_mouse_pos)
    def _on_mouse_pos(self, instance, pos):
        if not self.get_root_window(): return
        self.hover = self.collide_point(*self.to_widget(*pos))
    def on_is_active(self, *_):
        self.color = list(P["primary"]) if self.is_active else list(P["nav_item"])
        self.bold  = self.is_active
    def refresh_theme(self):
        self.active_color = list(P["nav_active"])
        self.hover_color  = list(P["selected"])
        self.on_is_active()

class CardBox(BoxLayout):
    bg_color = ListProperty([1,1,1,1])
    border_color = ListProperty([0,0,0,0.1])
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.refresh_theme()
    def refresh_theme(self):
        self.bg_color = list(P["card"])
        self.border_color = list(P["border"])

class SidebarBox(BoxLayout):
    bg_color = ListProperty([1,1,1,1])
    border_color = ListProperty([0,0,0,0.1])
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.refresh_theme()
    def refresh_theme(self):
        self.bg_color = list(P["sidebar"])
        self.border_color = list(P.get("sidebar_l", P["border"]))

class UserBox(BoxLayout):
    bg_color = ListProperty([0.9,0.9,0.9,1])
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.refresh_theme()
    def refresh_theme(self):
        self.bg_color = list(P.get("sidebar_l", P["sidebar"]))

class BgBox(BoxLayout):
    bg_color = ListProperty([1,1,1,1])
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.refresh_theme()
    def refresh_theme(self):
        self.bg_color = list(P["bg"])

class FieldInput(BoxLayout):
    text = StringProperty("")
    def __init__(self, **kwargs):
        self.hint_text = kwargs.pop('hint_text', "")
        kwargs.setdefault('size_hint_y', None)
        kwargs.setdefault('height', dp(38))
        super().__init__(**kwargs)
        with self.canvas.before:
            self._bg_c = Color(*P.get("input_bg", [1,1,1,1]))
            self._bg_r = RoundedRectangle(pos=self.pos, size=self.size, radius=[dp(8)])
            self._bd_c = Color(*P.get("border", [0,0,0,0.2]))
            self._bd_l = Line(rounded_rectangle=(self.x, self.y, self.width, self.height, dp(8)), width=dp(1))
        self.bind(pos=self._upd, size=self._upd)
        self.input = TextInput(
            text=self.text, hint_text=self.hint_text, multiline=False, write_tab=False,
            background_normal='', background_active='', background_color=[0,0,0,0],
            foreground_color=P["text"], cursor_color=P["primary"], hint_text_color=P["subtext"],
            font_size=sp(14), padding=[dp(10), dp(10)]
        )
        self.input.bind(text=lambda i,v: setattr(self, 'text', v))
        self.input.bind(focus=self._on_focus)
        self.bind(text=lambda i,v: setattr(self.input, 'text', v))
        self.add_widget(self.input)
        Clock.schedule_once(self._upd_pad, 0)
    def _upd(self, *args):
        self._bg_r.pos = self.pos
        self._bg_r.size = self.size
        self._bd_l.rounded_rectangle = (self.x, self.y, self.width, self.height, dp(8))
        self._upd_pad()
    def _upd_pad(self, *args):
        if self.input.line_height and self.height>1:
            dy = (self.height - self.input.line_height)/2
            self.input.padding = [dp(10), dy, dp(10), dy]
    def _on_focus(self, instance, focused):
        self._bd_c.rgba = P["primary"] if focused else P["border"]
        self._bd_l.width = dp(2) if focused else dp(1)
    def refresh_theme(self):
        self._bg_c.rgba = P.get("input_bg")
        self._bd_c.rgba = P.get("border")
        self.input.foreground_color = P["text"]
        self.input.cursor_color = P["primary"]
        self.input.hint_text_color = P["subtext"]

class FloatingLabelInput(BoxLayout):
    text = StringProperty('')
    password = BooleanProperty(False)
    
    def __init__(self, label_text="", **kwargs):
        self.password = kwargs.pop('password', False)
        height = kwargs.pop('height', dp(58))
        super().__init__(orientation='vertical', size_hint_y=None, height=height, **kwargs)
        self._label_text = label_text
        self._focused = False
        self._build()
    
    def _build(self):
        self._container = FloatLayout(size_hint=(1, 1))
        
        with self._container.canvas.before:
            self._bg_color = Color(*P.get("row_alt", [0.9, 0.9, 0.9, 1])) 
            self._bg_rect = RoundedRectangle(pos=self._container.pos, size=self._container.size, radius=[dp(8), dp(8), 0, 0])
            self._line_color = Color(*P["border"]) 
            self._border_line = Rectangle(pos=(self._container.x, self._container.y),
                                          size=(self._container.width, dp(1)))
            
        self._container.bind(pos=self._update_graphics, size=self._update_graphics)
        
        self._input = TextInput(
            multiline=False, password=self.password,
            background_normal='', background_active='', background_color=[0, 0, 0, 0],
            foreground_color=P["text"], cursor_color=P["primary"],
            font_size=sp(16), padding=[dp(14), dp(28), dp(14), dp(8)],
            size_hint=(1, 1), pos_hint={'x': 0, 'y': 0}, write_tab=False,
        )
        self._input.bind(focus=self._on_focus, text=self._on_text_change)
        self._container.add_widget(self._input)
        
        self._label = Label(text=self._label_text, font_size=sp(14), color=list(P["subtext"]),
                            halign='left', valign='middle', size_hint=(None, None))
        self._label.role = 'ignore'
        self._label.texture_update()
        self._label.size = self._label.texture_size
        self._container.add_widget(self._label)
        
        Clock.schedule_once(self._position_label, 0.05)
        self.add_widget(self._container)
    
    def _update_graphics(self, *args):
        self._bg_rect.pos = self._container.pos
        self._bg_rect.size = self._container.size
        self._border_line.pos = (self._container.x, self._container.y)
        self._border_line.size = (self._container.width, dp(2) if self._focused else dp(1))
        self._position_label()

    def refresh_theme(self):
        try:
            self._bg_color.rgba = P.get("row_alt", [0.9, 0.9, 0.9, 1])
            self._line_color.rgba = P["primary"] if self._focused else P["border"]
            self._label.color = list(P.get("subtext", self._label.color))
            try:
                self._input.foreground_color = P.get("text", self._input.foreground_color)
                self._input.cursor_color = P.get("primary", self._input.cursor_color)
            except: pass
            self._position_label()
        except: pass
    
    def _position_label(self, *args):
        self._label.texture_update()
        self._label.size = self._label.texture_size
        
        if self._focused or self._input.text:
            self._label.font_size = sp(11)
            self._label.texture_update()
            self._label.size = self._label.texture_size
            self._label.color = list(P["primary"]) if self._focused else list(P["subtext"])
            self._label.x = self._container.x + dp(14)
            self._label.y = self._container.y + self._container.height - self._label.height - dp(6)
        else:
            self._label.font_size = sp(14)
            self._label.texture_update()
            self._label.size = self._label.texture_size
            self._label.color = list(P["subtext"])
            self._label.x = self._container.x + dp(14)
            self._label.center_y = self._container.center_y
    
    def _on_focus(self, instance, focused):
        self._focused = focused
        self._line_color.rgba = P["primary"] if focused else P["border"]
        self._border_line.size = (self._container.width, dp(2) if focused else dp(1))
            
        if focused or self._input.text:
            target_y = self._container.y + self._container.height - dp(20)
            anim = Animation(font_size=sp(11), x=self._container.x + dp(14), y=target_y, duration=0.15, t='out_quad')
        else:
            anim = Animation(font_size=sp(14), x=self._container.x + dp(14), center_y=self._container.center_y, duration=0.15, t='out_quad')
            
        anim.start(self._label)
        self._label.color = list(P["primary"]) if focused else list(P["subtext"])
    
    def _on_text_change(self, instance, value):
        self.text = value
        if not self._focused: self._position_label()
    
    @property
    def input(self): return self._input
    def bind_on_text_validate(self, callback): self._input.bind(on_text_validate=callback)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  Enhanced Input Widgets
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

class DropdownOption(Button):
    hover = BooleanProperty(False)

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ''
        self.background_down = ''
        self.background_color = [0,0,0,0]
        self.size_hint_y = None
        self.height = dp(46)
        self.halign = 'left'
        self.valign = 'middle'
        self.padding = [dp(20), 0]
        self.color = P['text']
        self.bind(size=self.setter('text_size'))
        self.bind(state=self._update_bg)
        self.bind(hover=self._update_bg)
        self.bind(parent=self._bind_mouse)

    def _bind_mouse(self, instance, parent):
        if parent: Window.bind(mouse_pos=self._on_mouse_pos)
        else: Window.unbind(mouse_pos=self._on_mouse_pos)

    def _on_mouse_pos(self, instance, pos):
        if not self.get_root_window(): return
        self.hover = self.collide_point(*self.to_widget(*pos))

    def _update_bg(self, *args):
        self.canvas.before.clear()
        with self.canvas.before:
            if self.state == 'down':
                Color(*P['selected'])
                RoundedRectangle(pos=(self.x+dp(4), self.y+dp(2)), size=(self.width-dp(8), self.height-dp(4)), radius=[dp(8)])
            elif self.hover:
                c = P["primary"]
                Color(c[0], c[1], c[2], 0.1 if CURRENT_THEME == 'light' else 0.25)
                RoundedRectangle(pos=(self.x+dp(4), self.y+dp(2)), size=(self.width-dp(8), self.height-dp(4)), radius=[dp(8)])

class SelectionDropdown(FloatingLabelInput):
    values = ListProperty([])
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.input.readonly = True
        self.input.cursor_color = [0,0,0,0]
        self.input.bind(focus=self._on_focus)
        self._arrow = Label(text="[font=Emoji]🔻[/font]", markup=True, size_hint=(None, None), size=(dp(30), dp(30)),
                            pos_hint={'right': 1, 'center_y': 0.5}, color=P['subtext'], font_size=sp(12))
        self._container.add_widget(self._arrow)

    def _on_focus(self, instance, is_focused):
        if is_focused: 
            self.input.focus = False
            self._open_popup()

    def _open_popup(self):
        sv = ScrollView(bar_width=dp(8), bar_margin=dp(4), do_scroll_x=False)
        box = GridLayout(cols=1, size_hint_y=None, spacing=dp(1))
        box.bind(minimum_height=box.setter('height'))
        
        popup_ref = [None]

        for val in self.values:
            btn = DropdownOption(text=str(val))
            btn.bind(on_release=lambda btn, v=val: self._select(v, popup_ref[0]))
            box.add_widget(btn)

        sv.add_widget(box)
        
        count = len(self.values)
        h_est = count * 50 + 90
        h_final = min(h_est, 550)
        
        popup_ref[0] = _make_popup(self._label_text, sv, w=340, h=h_final)
        popup_ref[0].open()
        
    def _select(self, value, popup):
        self.input.text = str(value)
        if not self.text: self._position_label()
        if popup: popup.dismiss()

class SimpleDropdown(BoxLayout):
    values = ListProperty([])
    text = StringProperty("")
    def __init__(self, title="Select Option", **kwargs):
        self.title = title
        self.values = kwargs.pop('values', [])
        h = kwargs.pop('height', dp(44))
        shy = kwargs.pop('size_hint_y', None)
        super().__init__(orientation='horizontal', size_hint_y=shy, height=h, **kwargs)
        with self.canvas.before:
            self._bg_c = Color(*P.get("input_bg", [1,1,1,1]))
            self._bg_rect = RoundedRectangle(pos=self.pos, size=self.size, radius=[dp(8)])
            self._bd_c = Color(*P.get("border", [0,0,0,0.2]))
            self._bd_line = Line(rounded_rectangle=(self.x, self.y, self.width, self.height, dp(8)), width=dp(1))
        self.bind(pos=self._upd, size=self._upd)
        
        self._input = TextInput(
            readonly=True, multiline=False, background_normal='', background_active='',
            background_color=[0,0,0,0], foreground_color=P["text"], cursor_color=[0,0,0,0],
            font_size=sp(14), padding=[dp(10), dp(10)], size_hint=(1, 1), write_tab=False
        )
        self._input.bind(focus=self._on_focus)
        self._input.bind(text=lambda i, v: setattr(self, 'text', v))
        self.add_widget(self._input)
        
        self._arrow = Label(text="[font=Emoji]🔻[/font]", markup=True, size_hint=(None, 1), width=dp(30),
                            color=P['subtext'], font_size=sp(10))
        self.add_widget(self._arrow)

    @property
    def input(self): return self._input

    def _upd(self, *args):
        self._bg_rect.pos = self.pos
        self._bg_rect.size = self.size
        self._bd_line.rounded_rectangle = (self.x, self.y, self.width, self.height, dp(8))
        if self._input.line_height and self.height > 1:
            dy = (self.height - self._input.line_height)/2
            self._input.padding = [dp(10), dy, dp(10), dy]

    def _on_focus(self, instance, focused):
        if focused:
            self._input.focus = False
            self._open_popup()

    def _open_popup(self):
        sv = ScrollView(bar_width=dp(8), bar_margin=dp(4), do_scroll_x=False)
        box = GridLayout(cols=1, size_hint_y=None, spacing=dp(1))
        box.bind(minimum_height=box.setter('height'))
        popup_ref = [None]
        for val in self.values:
            btn = DropdownOption(text=str(val))
            btn.bind(on_release=lambda btn, v=val: self._select(v, popup_ref[0]))
            box.add_widget(btn)
        sv.add_widget(box)
        h_final = min(len(self.values) * 50 + 90, 550)
        popup_ref[0] = _make_popup(self.title, sv, w=340, h=h_final)
        popup_ref[0].open()

    def _select(self, value, popup):
        self._input.text = str(value)
        if popup: popup.dismiss()

    def refresh_theme(self):
        self._bg_c.rgba = P.get("input_bg")
        self._bd_c.rgba = P.get("border")
        self._input.foreground_color = P["text"]
        self._arrow.color = P["subtext"]

class SearchField(BoxLayout):
    text = StringProperty("")
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.size_hint = (None, None)
        self.size = (dp(240), dp(40))
        self.orientation = 'horizontal'
        self.spacing = dp(4)
        self.padding = [dp(12), 0]
        
        with self.canvas.before:
            self._bg_c = Color(*P["input_bg"])
            self._bg_rect = RoundedRectangle(pos=self.pos, size=self.size, radius=[dp(20)])
            self._bd_c = Color(*P["border"])
            self._bd_line = Line(rounded_rectangle=(self.x, self.y, self.width, self.height, dp(20)), width=dp(1))
        self.bind(pos=self._upd, size=self._upd)
        
        self.icon = Label(text="[font=Emoji]🔍[/font]", markup=True, size_hint=(None, 1), width=dp(24), 
                          font_size=sp(16), color=P["subtext"])
        self.add_widget(self.icon)
        
        self.input = TextInput(
            hint_text="Search...", multiline=False, write_tab=False,
            background_normal='', background_active='', background_color=[0,0,0,0],
            foreground_color=P["text"], cursor_color=P["primary"], hint_text_color=P["subtext"],
            font_size=sp(13), size_hint=(1, 1), padding=[0, dp(11), 0, 0]
        )
        self.input.bind(text=lambda i,v: setattr(self, 'text', v))
        self.add_widget(self.input)

    def _upd(self, *args):
        self._bg_rect.pos = self.pos
        self._bg_rect.size = self.size
        self._bd_line.rounded_rectangle = (self.x, self.y, self.width, self.height, dp(20))

    def refresh_theme(self):
        self._bg_c.rgba = P["input_bg"]
        self._bd_c.rgba = P["border"]
        self.icon.color = P["subtext"]
        self.input.foreground_color = P["text"]
        self.input.cursor_color = P["primary"]
        self.input.hint_text_color = P["subtext"]

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  Backend helpers  (logic unchanged from original)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def run_query(sql, params=(), fetch=False, commit=False, retry=3):
    import libsql_client
    import time
    for attempt in range(retry):
        try:
            client = libsql_client.create_client_sync(TURSO_DB_URL, auth_token=TURSO_AUTH_TOKEN)
            rs = client.batch([(sql, params)])[0]
            client.close()
            return [tuple(r) for r in rs.rows] if fetch else None
        except Exception as e:
            print(f"DB Error ({attempt+1}/{retry}): {e}")
            if attempt < retry-1: time.sleep(2**attempt)
        finally:
            try: client.close()
            except: pass
    return None

def check_internet():
    try: socket.create_connection(("8.8.8.8",53),timeout=3); return True
    except OSError: return False

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  UI utility helpers
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _sep(height=1, color=None):
    w = Widget(size_hint_y=None, height=dp(height))
    c = color or P["border"]
    with w.canvas:
        Color(*c)
        rect = Rectangle(pos=w.pos, size=w.size)
    w.bind(pos=lambda i, v: setattr(rect, 'pos', v),
           size=lambda i, v: setattr(rect, 'size', v))
    return w

def _label(text, size=13, bold=False, color=None, halign='left', height=None, wrap=True, role=None, **kwargs):
    lbl = Label(
        text=text, font_size=sp(size), bold=bold,
        color=color or P["text"], halign=halign, valign='middle', **kwargs
    )
    lbl.role = role if role else ('text' if color is None else 'ignore')
    if wrap: lbl.bind(size=lbl.setter('text_size'))
    if height is not None: lbl.size_hint_y = None; lbl.height = dp(height)
    return lbl

def _make_popup(title, content_widget, w=420, h=260, overlay_alpha=None):
    outer = BoxLayout(orientation='vertical')
    with outer.canvas.before:
        Color(*P["card"])
        bg_rect = RoundedRectangle(pos=outer.pos, size=outer.size, radius=[dp(16)])
    outer.bind(pos=lambda i, v: setattr(bg_rect, 'pos', v),
               size=lambda i, v: setattr(bg_rect, 'size', v))
    
    tbar = Widget(size_hint_y=None, height=dp(56))
    with tbar.canvas.before:
        Color(*P["primary"])
        tbar_rect = RoundedRectangle(pos=tbar.pos, size=tbar.size, radius=[dp(16), dp(16), 0, 0])
    tbar.bind(pos=lambda i, v: setattr(tbar_rect, 'pos', v),
              size=lambda i, v: setattr(tbar_rect, 'size', v))
    title_lbl = _label(title, size=15, bold=True, color=list(P["white"]), height=56, role="white")
    title_lbl.pos_hint = {'center_x': 0.5, 'center_y': 0.5}
    title_lbl.halign = 'center'
    tbar_layout = AnchorLayout(anchor_x='center', anchor_y='center', size_hint_y=None, height=dp(56))
    tbar_layout.add_widget(title_lbl)
    
    tbar_box = BoxLayout(size_hint_y=None, height=dp(56))
    with tbar_box.canvas.before:
        Color(*P["primary"])
        tbar_bg = RoundedRectangle(pos=tbar_box.pos, size=tbar_box.size, radius=[dp(16), dp(16), 0, 0])
    tbar_box.bind(pos=lambda i, v: setattr(tbar_bg, 'pos', v),
                  size=lambda i, v: setattr(tbar_bg, 'size', v))
    tbar_box.add_widget(tbar_layout)
    outer.add_widget(tbar_box)
    outer.add_widget(content_widget)
    
    overlay_c = list(P.get("overlay", [0,0,0,0.5]))
    if overlay_alpha is not None:
        overlay_c[3] = overlay_alpha
    
    pop = Popup(title='', content=outer, size_hint=(None,None), size=(dp(w), dp(h)),
                separator_height=0, background='', 
                background_color=[0,0,0,0], # Make popup's own background transparent
                overlay_color=overlay_c)
    return pop

def show_message(title, message, on_dismiss=None, msg_type="info"):
    icons = {"info":"[font=Emoji]ℹ[/font]","error":"[font=Emoji]✖[/font]","warning":"[font=Emoji]⚠[/font]","success":"[font=Emoji]✔[/font]"}
    colors= {"info":P["primary"],"error":P["error"], "warning":P["warning"],"success":P["success"]}
    body = BoxLayout(orientation='vertical', padding=[dp(24),dp(16)], spacing=dp(12))
    
    icon_row = AnchorLayout(anchor_x='center', size_hint_y=None, height=dp(56))
    icon_circle = Widget(size_hint=(None, None), size=(dp(48), dp(48)))
    with icon_circle.canvas:
        Color(*colors.get(msg_type, P["primary"]))
        ic_ellipse = Ellipse(pos=icon_circle.pos, size=icon_circle.size)
    icon_circle.bind(pos=lambda i, v: setattr(ic_ellipse, 'pos', v), size=lambda i, v: setattr(ic_ellipse, 'size', v))
    icon_lbl = Label(text=icons.get(msg_type, "[font=Emoji]ℹ[/font]"), markup=True, font_size=sp(20), bold=True, color=list(P["white"]), 
                     size_hint=(None, None), size=(dp(48), dp(48)))
    icon_lbl.role = "white"
    icon_box = FloatLayout(size_hint=(None, None), size=(dp(48), dp(48)))
    icon_box.add_widget(icon_circle)
    icon_box.add_widget(icon_lbl)
    icon_row.add_widget(icon_box)
    body.add_widget(icon_row)
    
    msg_lbl = _label(message, size=13, color=P["text"], halign='center')
    body.add_widget(msg_lbl)
    body.add_widget(Widget())

    ok = RoundBtn(text="OK", btn_color=list(colors.get(msg_type, P["primary"])),
                  role=(msg_type if msg_type in P else "primary"), size_hint_y=None, height=dp(44))
    body.add_widget(ok)

    pop = _make_popup(title, body, w=400, h=220)
    def _ok(_): pop.dismiss(); on_dismiss and on_dismiss()
    ok.bind(on_press=_ok); pop.open(); return pop

def show_confirm(title, message, on_yes=None, on_no=None):
    body = BoxLayout(orientation='vertical', padding=[dp(28),dp(20)], spacing=dp(16))
    
    q_row = AnchorLayout(anchor_x='center', size_hint_y=None, height=dp(56))
    q_circle = Widget(size_hint=(None, None), size=(dp(48), dp(48)))
    with q_circle.canvas:
        Color(*P["warning"])
        qc_ellipse = Ellipse(pos=q_circle.pos, size=q_circle.size)
    q_circle.bind(pos=lambda i, v: setattr(qc_ellipse, 'pos', v), size=lambda i, v: setattr(qc_ellipse, 'size', v))
    q_lbl = Label(text="?", font_size=sp(24), bold=True, color=list(P["white"]), size_hint=(None, None), size=(dp(48), dp(48)))
    q_lbl.role = "white"
    q_box = FloatLayout(size_hint=(None, None), size=(dp(48), dp(48)))
    q_box.add_widget(q_circle)
    q_box.add_widget(q_lbl)
    q_row.add_widget(q_box)
    body.add_widget(q_row)
    
    body.add_widget(_label(message, size=13, color=P["text"], halign='center'))
    body.add_widget(Widget())
    
    row = BoxLayout(size_hint_y=None, height=dp(46), spacing=dp(12))
    yes = RoundBtn(text="Yes", btn_color=list(P["success"]), role="success")
    no  = RoundBtn(text="No",  btn_color=list(P["error"]), role="error")
    row.add_widget(no); row.add_widget(yes)
    body.add_widget(row)
    pop = _make_popup(title, body, w=420, h=260)
    pop.auto_dismiss = False
    def _y(_): pop.dismiss(); on_yes and on_yes()
    def _n(_): pop.dismiss(); on_no  and on_no()
    yes.bind(on_press=_y); no.bind(on_press=_n); pop.open(); return pop

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  MultiSelectDropdown
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
class MultiSelectItem(BoxLayout):
    active = BooleanProperty(False)
    text = StringProperty('')
    hover = BooleanProperty(False)

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.size_hint_y = None
        self.height = dp(48)
        self.padding = [dp(12), 0]
        self.spacing = dp(12)
        
        self.cb = CheckBox(active=self.active, size_hint_x=None, width=dp(28), color=P['primary'])
        self.cb.bind(active=self.on_active_prop)
        self.add_widget(self.cb)
        
        self.lbl = _label(self.text, size=13, color=P["text"])
        self.add_widget(self.lbl)
        
        self.bind(text=lambda i, v: setattr(self.lbl, 'text', v))
        self.bind(active=lambda i, v: setattr(self.cb, 'active', v))
        
        self.bind(parent=self._bind_mouse)
        self.bind(pos=self._update_bg, size=self._update_bg)

    def on_active_prop(self, instance, value):
        self.active = value

    def on_touch_down(self, touch):
        if self.collide_point(*touch.pos):
            self.active = not self.active
            return True
        return super().on_touch_down(touch)

    def _bind_mouse(self, instance, parent):
        if parent: Window.bind(mouse_pos=self._on_mouse_pos)
        else: Window.unbind(mouse_pos=self._on_mouse_pos)

    def _on_mouse_pos(self, instance, pos):
        if not self.get_root_window(): return
        is_hover = self.collide_point(*self.to_widget(*pos))
        if self.hover != is_hover:
            self.hover = is_hover
            self._update_bg()

    def _update_bg(self, *args):
        self.canvas.before.clear()
        c = [0,0,0,0]
        if self.hover:
            base = P["primary"]
            c = [base[0], base[1], base[2], 0.1 if CURRENT_THEME == 'light' else 0.25]
        
        with self.canvas.before:
            Color(*c)
            Rectangle(pos=self.pos, size=self.size)

class MultiSelectDropdown(FloatingLabelInput):
    values = ListProperty([])
    
    def __init__(self, label_text, options, **kwargs):
        self.values = options
        super().__init__(label_text=label_text, **kwargs)
        self.input.readonly = True
        self.input.cursor_color = [0,0,0,0]
        self.input.bind(focus=self._on_focus)
        self._arrow = Label(text="[font=Emoji]🔻[/font]", markup=True, size_hint=(None, None), size=(dp(30), dp(30)),
                            pos_hint={'right': 1, 'center_y': 0.5}, color=P['subtext'], font_size=sp(12))
        self._container.add_widget(self._arrow)
        self._selected = []

    def get_value(self): return ", ".join(self._selected)
    def set_value(self, v):
        self._selected = [x.strip() for x in v.split(',') if x.strip()]
        self.input.text = ", ".join(self._selected)
        self._position_label()

    def _on_focus(self, instance, is_focused):
        if is_focused: 
            self.input.focus = False
            self._open_popup()

    def _open_popup(self):
        form = BoxLayout(orientation='vertical', size_hint_y=None, spacing=dp(1))
        form.bind(minimum_height=form.setter('height'))
        
        checks = {}
        for opt in self.values:
            item = MultiSelectItem(text=opt, active=(opt in self._selected))
            form.add_widget(item)
            checks[opt] = item

        sv = ScrollView(bar_width=dp(8), bar_margin=dp(4))
        sv.add_widget(form)
        content = BoxLayout(orientation='vertical', spacing=dp(8), padding=[0, 0, 0, dp(12)])
        content.add_widget(sv)
        
        done_btn_container = BoxLayout(padding=[dp(12), 0], size_hint_y=None, height=dp(44))
        done = RoundBtn(text="Done", size_hint_y=None, height=dp(44))
        done_btn_container.add_widget(done)
        content.add_widget(done_btn_container)
        
        h_est = len(self.values) * 48 + 130
        h = min(h_est, 550)
        pop = _make_popup(self._label_text, content, w=360, h=h)
        pop.auto_dismiss = False

        def _done(_):
            self._selected = [o for o, item in checks.items() if item.active]
            self.input.text = ", ".join(self._selected)
            self._position_label()
            pop.dismiss()
        
        done.bind(on_press=_done)
        pop.open()

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  Log table (RecycleView-based)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

COL_WIDTHS = [dp(160), dp(170), dp(300), dp(180), dp(180), dp(120), dp(160), dp(250), dp(160)]

class LogTableRow(BoxLayout):
    hover = BooleanProperty(False)
    context_active = BooleanProperty(False)
    _popup_active = False
    def __init__(self, row_data, col_widths, idx, on_context_click=None, **kwargs):
        super().__init__(**kwargs)
        self.idx = idx
        self.row_data = row_data
        self.on_context_click = on_context_click
        self.size_hint_y = None
        self.height = dp(42)
        
        with self.canvas.before:
            self.bg_color = Color(*(P['row_alt'] if idx % 2 else P['card']))
            self.rect = Rectangle(pos=self.pos, size=self.size)
        self.bind(pos=self._update_rect, size=self._update_rect)
        self.bind(parent=self._bind_mouse)

        for i, c in enumerate(row_data):
            if i >= len(col_widths): break
            w = col_widths[i]
            lbl = Label(text=str(c).replace('\n', ' '), color=P['text'], font_size=sp(14),
                        size_hint_x=None, width=w, halign='left', valign='middle',
                        shorten=True, shorten_from='right', padding=(dp(8),0))
            lbl.text_size = (w, dp(42))
            lbl.bind(size=lbl.setter('text_size'))
            self.add_widget(lbl)

    def set_context_active(self, active):
        self.context_active = active
        LogTableRow._popup_active = active
        self._update_bg_color()

    def _bind_mouse(self, instance, parent):
        if parent: Window.bind(mouse_pos=self._on_mouse_pos)
        else: Window.unbind(mouse_pos=self._on_mouse_pos)

    def _on_mouse_pos(self, instance, pos):
        if not self.get_root_window(): return

        if LogTableRow._popup_active and not self.context_active:
            if self.hover:
                self.hover = False
                self._update_bg_color()
            return

        is_hover = self.collide_point(*self.to_widget(*pos))

        # If context menu is active, just track hover state without visual change
        if self.context_active:
            self.hover = is_hover
            return

        if self.hover != is_hover:
            self.hover = is_hover
            self._update_bg_color()

    def _update_bg_color(self, *args):
        if self.context_active or self.hover:
            self.bg_color.rgba = P['selected']
        else:
            self.bg_color.rgba = P['row_alt'] if self.idx % 2 else P['card']

    def _update_rect(self, *args):
        self.rect.pos = self.pos
        self.rect.size = self.size

    def on_touch_down(self, touch):
        if self.collide_point(*touch.pos):
            if getattr(touch, 'button', None) == 'right':
                if self.on_context_click:
                    self.on_context_click(self, touch.pos)
                    return True
        return super().on_touch_down(touch)

    def refresh_theme(self):
        self._update_bg_color()
        for child in self.children:
            if isinstance(child, Label):
                child.color = P['text']

class LogTableWidget(BoxLayout):
    def __init__(self, headers, col_widths=None, rows_per_page=None, **kw):
        super().__init__(orientation='vertical', **kw)
        self.headers = headers
        self.col_widths = col_widths if col_widths else COL_WIDTHS
        
        self.auto_paginate = (rows_per_page is None)
        self.rows_per_page = rows_per_page if rows_per_page else 14
        self.current_page = 1
        self.all_rows = []
        self._on_row_context = None
        
        self.total_w = sum(self.col_widths) + (len(headers) * dp(2))
        
        # 1. Sticky Header (Scrolls horizontally, fixed vertically)
        self.header_scroll = ScrollView(do_scroll_y=False, do_scroll_x=False, size_hint_y=None, height=dp(44), size_hint_x=1)
        self.header_grid = BoxLayout(size_hint_x=None, height=dp(44))
        
        with self.header_grid.canvas.before:
            self.header_bg_c = Color(*P['header_bg'])
            self.header_bg_r = Rectangle(pos=self.header_grid.pos, size=self.header_grid.size)
        self.header_grid.bind(pos=lambda i,v: setattr(self.header_bg_r, 'pos', v),
                              size=lambda i,v: setattr(self.header_bg_r, 'size', v))

        self.header_scroll.add_widget(self.header_grid)
        
        # Populate header immediately
        for i, h in enumerate(self.headers):
            w = self.col_widths[i] if i < len(self.col_widths) else dp(150)
            lbl = Label(text=h.upper(), bold=True, color=P['text'], font_size=sp(14),
                        size_hint_x=None, width=w, halign='left', valign='middle', padding=(dp(8), 0))
            lbl.text_size = (w, dp(44))
            lbl.bind(size=lbl.setter('text_size'))
            self.header_grid.add_widget(lbl)
            
        self.add_widget(self.header_scroll)
        
        # 2. Body ScrollView
        self.body_scroll = ScrollView(do_scroll_y=True, do_scroll_x=True, bar_width=dp(10), scroll_type=['bars', 'content'])
        self.grid = GridLayout(cols=1, size_hint_x=None, size_hint_y=None, spacing=dp(1))
        self.grid.bind(minimum_height=self.grid.setter('height'))
        self.body_scroll.add_widget(self.grid)
        # Bind body horizontal scroll to header horizontal scroll
        self.body_scroll.bind(scroll_x=self.header_scroll.setter('scroll_x'))
        self.add_widget(self.body_scroll)
        
        # 3. Pagination Footer
        self.footer = BoxLayout(size_hint_y=None, height=dp(48), spacing=dp(10), padding=[dp(10), dp(4)])
        self.btn_prev = RoundBtn(text="< Prev", size_hint_x=None, width=dp(60), height=dp(40), role="primary")
        self.btn_prev.bind(on_press=self.prev_page)
        self.lbl_page = Label(text="Page 1", color=P['text'], font_size=sp(13), halign='center')
        self.btn_next = RoundBtn(text="Next >", size_hint_x=None, width=dp(60),height=dp(40), role="primary")
        self.btn_next.bind(on_press=self.next_page)
        
        self.footer.add_widget(Widget())
        self.footer.add_widget(self.btn_prev)
        self.footer.add_widget(self.lbl_page)
        self.footer.add_widget(self.btn_next)
        self.footer.add_widget(Widget())
        self.add_widget(self.footer)
        
        if self.auto_paginate:
            self.bind(height=self._recalc_pagination)
        
        self.bind(size=self._update_widths)
        Clock.schedule_once(self._update_widths, 0)
        if self.auto_paginate:
            Clock.schedule_once(self._recalc_pagination, 0)
    
    def _update_widths(self, *args):
        w = max(self.total_w, self.width)
        self.header_grid.width = w
        self.grid.width = w

    def _recalc_pagination(self, *args):
        if not self.auto_paginate or self.height < dp(100): return
        
        # Heights: Header=44, Footer=48, Row=42+1(spacing)=43
        header_h = dp(44)
        footer_h = dp(48)
        row_h = dp(43) 
        
        # 1. Try fitting without footer
        avail_h = self.height - header_h
        capacity = int(max(1, avail_h // row_h))
        total = len(self.all_rows)
        
        if total <= capacity:
            # Fits on one page
            if self.footer in self.children: self.remove_widget(self.footer)
            self.rows_per_page = max(total, 1)
        else:
            # Needs pagination
            if self.footer not in self.children: self.add_widget(self.footer)
            avail_h = self.height - header_h - footer_h
            self.rows_per_page = int(max(1, avail_h // row_h))
            self.footer.opacity = 1
            self.footer.disabled = False

        self._update_view()

    def set_data(self, rows, on_row_context=None):
        self.all_rows = rows
        self._on_row_context = on_row_context
        self.current_page = 1
        if self.auto_paginate:
            self._recalc_pagination()
        else:
            self._update_view()
    
    def refresh_theme(self):
        # Header
        self.header_bg_c.rgba = P['header_bg']
        for child in self.header_grid.children:
            if isinstance(child, Label):
                child.color = P['text']
        # Footer
        if hasattr(self, 'lbl_page'): self.lbl_page.color = P['text']
        # Rows
        for child in self.grid.children:
            if hasattr(child, 'refresh_theme'):
                child.refresh_theme()

    def _update_view(self):
        self.grid.clear_widgets()
        total = len(self.all_rows)
        
        # Handle footer visibility for non-auto mode or explicit empty state
        if not self.auto_paginate:
            if total == 0:
                self.footer.opacity = 0
                self.footer.disabled = True
                return
            self.footer.opacity = 1
            self.footer.disabled = False
        
        total_pages = max(1, (total + self.rows_per_page - 1) // self.rows_per_page)
        if self.current_page > total_pages: self.current_page = total_pages
        if self.current_page < 1: self.current_page = 1
        
        start = (self.current_page - 1) * self.rows_per_page
        end = start + self.rows_per_page
        page_rows = self.all_rows[start:end]
        
        for idx, row in enumerate(page_rows):
            self.grid.add_widget(LogTableRow(row, self.col_widths, idx, on_context_click=self._on_row_context))
            
        self.lbl_page.text = f"Page {self.current_page} of {total_pages} ({total} records)"
        self.btn_prev.disabled = (self.current_page == 1)
        self.btn_next.disabled = (self.current_page == total_pages)
        self.body_scroll.scroll_y = 1

    def prev_page(self, *args):
        if self.current_page > 1:
            self.current_page -= 1
            self._update_view()
            
    def next_page(self, *args):
        total = len(self.all_rows)
        total_pages = (total + self.rows_per_page - 1) // self.rows_per_page
        if self.current_page < total_pages:
            self.current_page += 1
            self._update_view()

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  Login Screen
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

class LoginScreen(Screen):
    def __init__(self, **kw):
        super().__init__(**kw)
        self._build()

    def _build(self):
        root = FloatLayout()
        self._root = root
        with root.canvas.before:
            Color(*P["bg"]); self._bg_rect = Rectangle(pos=root.pos, size=root.size)
        root.bind(pos =lambda i,v: setattr(self._bg_rect,'pos',v), size=lambda i,v: setattr(self._bg_rect,'size',v))

        # REMOVED global padding so elements can touch the actual edges of the card
        card = CardBox(orientation='vertical', padding=0, spacing=0, size_hint=(None, None),
                       size=(dp(400), dp(570)), pos_hint={'center_x': 0.5, 'center_y': 0.5})
        self._card = card
        # Ensure the login card explicitly uses the current theme's card colors
        try:
            card.bg_color = list(P.get('card', card.bg_color))
            card.border_color = list(P.get('border', card.border_color))
        except Exception:
            pass

        # ── Theme Toggle (Absolutely positioned top-right of Window) ──
        icon = "[font=Emoji]☀️[/font]" if CURRENT_THEME == 'dark' else "[font=Emoji]🌙[/font]"
        lbl = "Light Mode" if CURRENT_THEME == 'dark' else "Dark Mode"
        self._theme_btn = RoundBtn(
            text=f"{icon} {lbl}", role="selected", text_role="text",
            size_hint=(None, None), size=(dp(120), dp(28)),
            pos_hint={'right': 0.98, 'top': 0.96}
        )
        self._theme_btn.bind(on_press=self._toggle_theme)
        root.add_widget(self._theme_btn)

        # ── Top Header (Logo) ──
        header = FloatLayout(size_hint_y=None, height=dp(60))

        # Centered Logo
        if os.path.exists(LOGO_PATH):
            logo = KivyImage(source=LOGO_PATH, size_hint=(None,None), size=(dp(100),dp(100)),
                             pos_hint={'center_x': 0.5, 'center_y': 1})
            header.add_widget(logo)
        else:
            lbl = _label("Kalinga OpsHUB", size=26, bold=True, color=P["primary"], halign='center', role="primary")
            lbl.pos_hint = {'center_x': 0.5, 'center_y': 0.4}
            header.add_widget(lbl)

        card.add_widget(header)

        # ── Content Area (Inputs & Buttons) ──
        # Reapply the 32dp padding here so the form elements stay properly spaced
        content = BoxLayout(orientation='vertical', padding=[dp(32), dp(0), dp(32), dp(32)], spacing=dp(15))

        content.add_widget(_label("Kalinga Operations HUB", size=20, color=P["text"], halign='center', height=dp(25), role="text", bold=True))

        content.add_widget(Widget())

        self._email_field = FloatingLabelInput(label_text="Email Address")
        content.add_widget(self._email_field)
        self._email = self._email_field  

        # ── Nested Password field with Eye icon ──
        p_outer = RelativeLayout(size_hint_y=None, height=dp(64))
        self._pass_field = FloatingLabelInput(label_text="Password", password=True)
        self._pass_field.bind_on_text_validate(lambda i: self.attempt_login())
        p_outer.add_widget(self._pass_field)
        self._pass = self._pass_field  
        
        self._eye_btn = GhostBtn(
            text="[font=Emoji]👁[/font]", font_size=sp(18), size_hint=(None, None), size=(dp(40), dp(40)),
            pos_hint={'right': 0.98, 'center_y': 0.4}, color=list(P["subtext"])
        )
        self._eye_btn.bind(on_press=self._toggle_pass)
        p_outer.add_widget(self._eye_btn)
        content.add_widget(p_outer)

        self._login_btn = RoundBtn(text="Log In")
        self._login_btn.bind(on_press=lambda i: self.attempt_login())
        content.add_widget(self._login_btn)

        if GOOGLE_LOGIN_AVAILABLE:
            div = BoxLayout(size_hint_y=None, height=dp(28), spacing=dp(10))
            div.add_widget(Widget())
            div.add_widget(_label("or continue with", size=11, color=P["subtext"], halign='center', height=28, wrap=False, role="subtext"))
            div.add_widget(Widget())
            content.add_widget(div)
            self._google_btn = RoundBtn(text="  Sign in with Google", btn_color=list(P["google"]), role="google")
            self._google_btn.bind(on_press=lambda i: self._login_google())
            content.add_widget(self._google_btn)

        fp = GhostBtn(text="Forgot password?", color=list(P["primary"]), font_size=sp(12), halign='right')
        fp.bind(on_press=lambda i: self._open_forgot())
        content.add_widget(fp)

        content.add_widget(_sep())
        footer = BoxLayout(size_hint_y=None, height=dp(40), spacing=dp(6))
        
        ver_box = BoxLayout(size_hint_x=None, width=dp(50), padding=[dp(6), dp(6)])
        ver_inner = BoxLayout(size_hint=(None, None), size=(dp(38), dp(22)))
        lbl_ver = Label(text=f"v{CURRENT_VERSION}", font_size=sp(13), color=list(P["text"]), bold=True); lbl_ver.role = "text"
        ver_inner.add_widget(lbl_ver)
        ver_box.add_widget(ver_inner)
        footer.add_widget(ver_box)
        
        footer.add_widget(Widget()) # Spacer

        self._upd_btn = GhostBtn(text="[font=Emoji]🔄[/font] Updates", color=list(P["primary"]), font_size=sp(13), size_hint_x=None, width=dp(90), height=dp(36))
        self._upd_btn.bind(on_press=lambda i: self._check_updates(self._upd_btn))
        footer.add_widget(self._upd_btn)
        
        if GOOGLE_LOGIN_AVAILABLE:
            self._unlink_btn = GhostBtn(text="[font=Emoji]🔓[/font] Unlink", color=list(P["primary"]), font_size=sp(13), size_hint_x=None, width=dp(80), height=dp(36))
            self._unlink_btn.bind(on_press=self._unlink_google)
            footer.add_widget(self._unlink_btn)

        ab = GhostBtn(text="[font=Emoji]ℹ[/font] About", color=list(P["primary"]), font_size=sp(13), size_hint_x=None, width=dp(70), height=dp(36))
        ab.bind(on_press=lambda i: self._show_about())
        footer.add_widget(ab)
        content.add_widget(footer)

        card.add_widget(content)

        root.add_widget(card)
        self.add_widget(root)
        apply_theme(CURRENT_THEME)
        Clock.schedule_once(lambda dt: self._apply_theme(), 0.05)

    def on_enter(self, *args):
        super().on_enter(*args)
        if hasattr(self, '_login_btn'): # Ensure widgets are built before resetting
            self.reset_form()
        self._card.opacity = 0
        self._card.center_y = self.height * 0.45 
        anim = Animation(opacity=1, center_y=self.height * 0.5, d=0.6, t='out_expo')
        anim.start(self._card)
        Window.bind(on_key_down=self._on_key_down)

    def on_leave(self, *args):
        Window.unbind(on_key_down=self._on_key_down)
        super().on_leave(*args)

    def _on_key_down(self, window, key, scancode, codepoint, modifiers):
        if 'ctrl' in modifiers and key == 114: # 'r'
            self.reset_form()
            if hasattr(self, '_email_field') and hasattr(self._email_field, 'input'): self._email_field.input.text = ""
            return True
        return False

    def _toggle_pass(self, *_):
        try:
            if getattr(self._pass_field, '_input', None) is not None:
                cur = self._pass_field._input.password
                self._pass_field._input.password = not cur
                self._eye_btn.text = "[font=Emoji]🙈[/font]" if not self._pass_field._input.password else "[font=Emoji]👁[/font]"
            else:
                cur = getattr(self._pass_field, 'password', False)
                setattr(self._pass_field, 'password', not cur)
                self._eye_btn.text = "[font=Emoji]🙈[/font]" if not getattr(self._pass_field, 'password') else "[font=Emoji]👁[/font]"
        except Exception: pass

    def _toggle_theme(self, *_):
        global CURRENT_THEME
        new = 'dark' if CURRENT_THEME == 'light' else 'light'
        try:
            with open(THEME_FILE, 'w') as f: f.write(new)
        except Exception: pass
        apply_theme(new)
        CURRENT_THEME = new
        try: Window.clearcolor = list(P.get('bg', Window.clearcolor))
        except: pass
        icon = "[font=Emoji]☀️[/font]" if new == 'dark' else "[font=Emoji]🌙[/font]"
        lbl = "Light Mode" if new == 'dark' else "Dark Mode"
        self._theme_btn.text = f"{icon} {lbl}"
        self._apply_theme()

    def _apply_theme(self):
        try:
            try:
                self._root.canvas.before.clear()
                with self._root.canvas.before:
                    Color(*P["bg"])
                    self._bg_rect = Rectangle(pos=self._root.pos, size=self._root.size)
            except: pass
            try:
                for w in list(self._root.walk()):
                    if hasattr(w, 'refresh_theme') and callable(w.refresh_theme):
                        try: w.refresh_theme()
                        except: pass
                    elif isinstance(w, Label):
                        role = getattr(w, 'role', 'text')
                        if role != 'ignore' and role in P:
                            w.color = P[role]
                    else:
                        try:
                            if hasattr(w, 'bg_color'): w.bg_color = list(P.get('card', w.bg_color))
                            if hasattr(w, 'border_color'): w.border_color = list(P.get('border', w.border_color))
                            if hasattr(w, 'btn_color') and str(getattr(w, 'text', '')).lower() not in ('logout','export csv'):
                                w.btn_color = list(P.get('primary', w.btn_color))
                        except: pass
            except: pass
            try:
                if hasattr(self, '_card'):
                    if hasattr(self._card, 'refresh_theme'):
                        self._card.refresh_theme()
            except: pass
            try:
                if hasattr(self, '_login_btn'): self._login_btn.btn_color = list(P["primary"])
            except: pass
        except: pass

    def _show_about(self):
        body = BoxLayout(orientation='vertical', padding=[dp(24), dp(20)], spacing=dp(12))
        if os.path.exists(LOGO_PATH):
            logo_wrap = AnchorLayout(anchor_x='center', size_hint_y=None, height=dp(90))
            logo_wrap.add_widget(KivyImage(source=LOGO_PATH, size_hint=(None, None), size=(dp(80), dp(80))))
            body.add_widget(logo_wrap)
        body.add_widget(_label("Kalinga OpsHUB", size=20, bold=True, color=P["primary"], halign='center', height=32))
        
        ver_row = AnchorLayout(anchor_x='center', size_hint_y=None, height=dp(28))
        ver_badge = BoxLayout(size_hint=(None, None), size=(dp(80), dp(24)), padding=[dp(12), dp(4)])
        
        body.add_widget(_label("A digital innovation tool combining an Outgoing Document Logbook and Leave Monitoring in one unified platform.", size=12, color=P["subtext"], halign='center', height=50))
        body.add_widget(_sep())
        
        dev_box = BoxLayout(orientation='vertical', size_hint_y=None, height=dp(60), spacing=dp(4))
        dev_box.add_widget(_label("Developed by", size=10, color=P["nav_item"], halign='center', height=18))
        dev_box.add_widget(_label("Chano, ISA II", size=14, bold=True, color=P["text"], halign='center', height=22))
        body.add_widget(dev_box)
        body.add_widget(Widget()) 
        close = RoundBtn(text="Close", btn_color=list(P["primary"]), size_hint_y=None, height=dp(44))
        body.add_widget(close)
        
        pop = _make_popup("About", body, w=380, h=450)
        close.bind(on_press=lambda _: pop.dismiss())
        pop.open()

    def _check_updates(self, btn):
        btn.text = "Checking…"; btn.disabled = True
        threading.Thread(target=self._update_thread, args=(btn,), daemon=True).start()

    def _update_thread(self, btn):
        import requests
        def _rst(dt): btn.text="[font=Emoji]🔄[/font] Updates"; btn.disabled=False
        if not check_internet():
            Clock.schedule_once(lambda dt: show_message("Error","No internet.",msg_type="error"),0)
            Clock.schedule_once(_rst,0); return
        try:
            data = requests.get(f"https://api.github.com/repos/Syano18/Kalinga-OpsHUB/releases/latest", timeout=10).json()
            ver  = data.get("tag_name","").lstrip("v").strip()
            url  = next((a["browser_download_url"] for a in data.get("assets",[]) if a.get("name","").endswith(".exe")), None)
            if not url: raise Exception("Invalid Url")
            if ver and ver != CURRENT_VERSION: Clock.schedule_once(lambda dt: self._prompt_update(ver,url),0)
            else: Clock.schedule_once(lambda dt: show_message("Up to Date","You have the latest version."),0)
        except Exception as e:
            err_msg = str(e)
            Clock.schedule_once(lambda dt: show_message("Error",err_msg,msg_type="error"),0)
        Clock.schedule_once(_rst,0)

    def _prompt_update(self, ver, url):
        show_confirm("Update Available", f"Version {ver} is available. Download and install now?", on_yes=lambda: self._download_update(url))

    def _download_update(self, url):
        pb_ref={}
        body = BoxLayout(orientation='vertical', padding=dp(20), spacing=dp(12))
        body.add_widget(_label("Downloading update…",size=14,bold=True,height=32))
        pb  = ProgressBar(max=100, size_hint_y=None, height=dp(22)); pb_ref['pb']=pb
        msg = _label("Initializing…", size=11, color=P["subtext"], height=26)
        body.add_widget(pb)
        body.add_widget(msg)
        pop = _make_popup("Updating…", body, w=370, h=200)
        pop.auto_dismiss = False
        pop.open()
        def _dl():
            import requests
            try:
                dest = os.path.join(tempfile.gettempdir(),"KalingaOpsHUB_Update.exe")
                r    = requests.get(url, stream=True)
                tot  = int(r.headers.get('content-length',0))
                with open(dest,'wb') as f:
                    dl=0
                    for chunk in r.iter_content(8192):
                        if chunk:
                            f.write(chunk); dl+=len(chunk)
                            if tot:
                                p=int(dl*100/tot)
                                Clock.schedule_once(lambda dt,pv=p:(setattr(pb_ref['pb'],'value',pv), setattr(msg,'text',f"{pv}% complete")),0)
                if os.path.getsize(dest)<1024: raise Exception("File too small.")
                Clock.schedule_once(lambda dt:(subprocess.Popen([dest]),App.get_running_app().stop()),0)
            except Exception as e:
                msg = str(e)
                Clock.schedule_once(lambda dt:(pop.dismiss(), show_message("Error",msg,msg_type="error")),0)
        threading.Thread(target=_dl, daemon=True).start()

    def _open_forgot(self):
        body = BoxLayout(orientation='vertical', padding=dp(20), spacing=dp(10))
        body.add_widget(_label("Enter your email and we'll send a reset link.", size=12, color=P["text"], height=10, halign='center', role="text"))
        ef_field = FloatingLabelInput(label_text="Email Address")
        body.add_widget(ef_field)
        send = RoundBtn(text="Send Reset Link", btn_color=list(P["success"]), role="success")
        body.add_widget(send) 
        pop = _make_popup("Reset Password", body, w=430, h=270, overlay_alpha=0)
        pop.auto_dismiss = False
        pop.open()
        def _go(_):
            em = ef_field.text.strip()
            if not em: return show_message("Error","Enter your email.",msg_type="warning")
            if not check_internet(): return show_message("Error","No internet.",msg_type="error")
            send.text="Sending…"; send.disabled=True
            threading.Thread(target=self._send_reset, args=(em,send,pop), daemon=True).start()
        send.bind(on_press=_go)

    def _send_reset(self, email, btn, pop):
        import requests
        try:
            requests.post(f"https://identitytoolkit.googleapis.com/v1/accounts:sendOobCode?key={FIREBASE_WEB_API_KEY}", json={"requestType":"PASSWORD_RESET","email":email})
            Clock.schedule_once(lambda dt:(show_message("Sent","Reset link sent to your email.",msg_type="success"), pop.dismiss()),0)
        except Exception:
            Clock.schedule_once(lambda dt:(show_message("Error","Failed to send reset link.",msg_type="error"), setattr(btn,'text','Send Reset Link'), setattr(btn,'disabled',False)),0)

    def _login_google(self):
        if not check_internet(): return show_message("Error","No internet.",msg_type="error")
        if not os.path.exists(resource_path("client_secret.json")): return show_message("Error","Google configuration missing.",msg_type="error")
        self._login_btn.disabled = True
        if hasattr(self, '_google_btn'): self._google_btn.disabled = True
        threading.Thread(target=self._google_flow, daemon=True).start()

    def _google_flow(self):
        import requests
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request
        creds=None
        if os.path.exists(GOOGLE_TOKEN_FILE):
            try: creds=Credentials.from_authorized_user_file(GOOGLE_TOKEN_FILE,SCOPES)
            except: creds=None
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try: creds.refresh(Request())
                except: creds=None
            if not creds:
                try:
                    flow=InstalledAppFlow.from_client_secrets_file(resource_path("client_secret.json"),SCOPES)
                    creds=flow.run_local_server(port=0,access_type='offline',prompt='consent')
                except Exception as e:
                    msg = f"Google Sign-In Error: {e}"; Clock.schedule_once(lambda dt:self._reset_ui(msg),0); return
            try:
                with open(GOOGLE_TOKEN_FILE,'w') as f: f.write(creds.to_json())
            except: pass
        try:
            info=requests.Session()
            info.headers.update({'Authorization':f'Bearer {creds.token}'})
            email=info.get('https://www.googleapis.com/oauth2/v1/userinfo').json().get('email')
            if email: Clock.schedule_once(lambda dt:self._process_google(email),0)
            else: raise Exception("Email not found.")
        except Exception as e:
            msg = f"Google Sign-In Error: {e}"; Clock.schedule_once(lambda dt:self._reset_ui(msg),0)

    def _process_google(self, email):
        def _db():
            rows=run_query('SELECT "Email","Role","First_Name","Middle_Name","Last_Name",'
                           '"Position","Salary","Salary_Grade","Status" FROM User_Permissions '
                           'WHERE LOWER(Email)=LOWER(?)',params=(email,),fetch=True)
            ud=dict(zip(USER_HEADERS,rows[0])) if rows else None
            Clock.schedule_once(lambda dt:self._finish_google(email,ud),0)
        threading.Thread(target=_db,daemon=True).start()

    def _finish_google(self, email, ud):
        if not ud: return self._reset_ui("Google account not registered.")
        if ud.get('Status')!="Active": return self._reset_ui("Account inactive.")
        try:
            with open(SESSION_FILE,'w') as f: f.write(email)
        except: pass
        App.get_running_app().launch_main(email,ud)

    def _reset_ui(self, err=None):
        if err: show_message("Login Error",err,msg_type="error")
        self._login_btn.disabled=False
        if hasattr(self, '_google_btn'): self._google_btn.disabled = False

    def reset_form(self):
        """Resets the login form to its initial state."""
        self._login_btn.text = "Log In"
        self._login_btn.disabled = False
        if hasattr(self, '_google_btn'): self._google_btn.disabled = False
        if hasattr(self, '_pass_field') and hasattr(self._pass_field, 'input'):
            self._pass_field.input.text = ""

    def attempt_login(self):
        em = self._email_field.text.strip()
        pw = self._pass_field.text
        if not em or not pw: return show_message("Login Error", "Please enter both email and password.", msg_type="warning")

        if not check_internet(): return show_message("Network Error","No internet connection.",msg_type="error")
        self._login_btn.text="Authenticating…"; self._login_btn.disabled=True
        threading.Thread(target=self._auth, args=(em, pw), daemon=True).start()

    def _auth(self, em, pw):
        import requests
        db={"ud":None}
        def _db():
            try:
                rows=run_query('SELECT "Email","Role","First_Name","Middle_Name","Last_Name",'
                               '"Position","Salary","Salary_Grade","Status" FROM User_Permissions '
                               'WHERE LOWER(Email)=LOWER(?)',params=(em,),fetch=True)
                if rows: db["ud"]=dict(zip(USER_HEADERS,rows[0]))
            except Exception as e: print(e)
        t=threading.Thread(target=_db); t.start()
        try:
            r=requests.post(f"https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key={FIREBASE_WEB_API_KEY}", json={"email":em,"password":pw,"returnSecureToken":True})
            t.join()
            if r.status_code==200:
                ud=db["ud"]
                if not ud: Clock.schedule_once(lambda dt:self._fail("Account not registered."),0); return
                if ud.get('Status')!="Active": Clock.schedule_once(lambda dt:self._fail("Account inactive."),0); return
                try:
                    with open(SESSION_FILE,'w') as f: f.write(em)
                except: pass
                Clock.schedule_once(lambda dt:App.get_running_app().launch_main(em,ud),0)
            else:
                msg="Invalid login details"
                try:
                    e=r.json().get('error',{}).get('message','')
                    if e=="USER_DISABLED": msg="Account is disabled."
                except: pass
                Clock.schedule_once(lambda dt:self._fail(msg),0)
        except Exception as e:
            msg = f"Connection Error: {e}"; Clock.schedule_once(lambda dt:self._fail(msg),0)

    def _fail(self, msg="Invalid login details"):
        show_message("Login Error",msg,msg_type="error")
        self._login_btn.text="Log In"; self._login_btn.disabled=False

    def _unlink_google(self, *args):
        if not os.path.exists(GOOGLE_TOKEN_FILE):
            show_message("Info", "No Google account is currently linked.", msg_type="info")
            return

        show_confirm(
            "Unlink Google Account?",
            "This will remove the link to your currently saved Google account. The next time you sign in with Google, you will be prompted to choose an account.",
            on_yes=self._perform_unlink
        )

    def _perform_unlink(self):
        try:
            os.remove(GOOGLE_TOKEN_FILE)
            show_message("Success", "Your Google account has been unlinked.", msg_type="success")
        except Exception as e:
            show_message("Error", f"Could not unlink account:\n{e}", msg_type="error")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  Main Screen
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

class MainScreen(Screen):
    def __init__(self, email, user_data=None, **kw):
        super().__init__(**kw)
        self.email   = email
        self.user_data= user_data
        self.role    = str(user_data.get('Role','Staff')).strip() if user_data else "Staff"

        self.master_logs        = []
        self.attendance_logs    = []
        self.headers            = LOG_HEADERS
        self.fetching_logs      = False
        self.fetching_attendance= False
        self.att_data_loaded    = False
        self.user_names         = []
        self.data_loaded_from_net=False
        self.connecting         = True
        self._monitor_ev        = None
        self._log_table         = None
        self.selected_att_employee = "All Employees"
        self.selected_att_month = datetime.now().strftime("%B %Y")
        self.last_attendance_update = None
        self._attendance_monitor_ev = None

        self._build()
        threading.Thread(target=self._connect_bg, daemon=True).start()
        self._monitor_ev = Clock.schedule_interval(lambda dt: threading.Thread(target=self._ping, daemon=True).start(), 5)
        Clock.schedule_once(lambda dt: self._apply_theme(), 0)

    def on_enter(self, *args):
        super().on_enter(*args)
        # Entry animations for Main Screen content
        self._content_area.opacity = 0
        self._content_area.y = -dp(50)
        anim = Animation(opacity=1, y=0, d=0.5, t='out_quad')
        anim.start(self._content_area)
        Window.bind(on_key_down=self._on_key_down)

    def on_leave(self, *args):
        Window.unbind(on_key_down=self._on_key_down)
        super().on_leave(*args)

    def _on_key_down(self, window, key, scancode, codepoint, modifiers):
        if 'ctrl' in modifiers and key == 114: # 'r'
            if self._nav_attendance.is_active:
                self._refresh_attendance()
            else: # Default to logbook
                self._refresh_logs()
            return True
        return False

    def _refresh_logs(self, *args):
        self.master_logs = []
        self.data_loaded_from_net = False
        self._draw_logs()

    def _refresh_attendance(self, *args):
        self.attendance_logs = []
        self.att_data_loaded = False
        self._draw_attendance()

    def _build(self):
        root = BoxLayout(orientation='horizontal')
        self._sidebar = self._make_sidebar()
        root.add_widget(self._sidebar)

        self._content_area = BgBox(orientation='vertical', bg_color=list(P["bg"]))
        root.add_widget(self._content_area)
        self.add_widget(root)
        self._draw_logs()

    def _make_sidebar(self):
        sb = SidebarBox(orientation='vertical', size_hint_x=None, width=dp(230), padding=[dp(12),dp(16)], spacing=dp(4))

        if os.path.exists(LOGO_PATH):
            w = AnchorLayout(anchor_x='center', size_hint_y=None, height=dp(100))
            w.add_widget(KivyImage(source=LOGO_PATH, size_hint=(None,None), size=(dp(90),dp(90))))
            sb.add_widget(w)

        sb.add_widget(Widget(size_hint_y=None, height=dp(30)))

        # Theme-aware: match the login button color (primary)
        self._nav_logbook = NavBtn(text="[font=Emoji]📁[/font]  Digital Logbook", is_active=True)
        self._nav_logbook.bind(on_press=lambda i: self._nav_to('logs'))
        self._nav_logbook.on_is_active()  # apply active/inactive coloring immediately
        sb.add_widget(self._nav_logbook)

        self._nav_attendance = NavBtn(text="[font=Emoji]📅[/font]  Attendance")
        self._nav_attendance.bind(on_press=lambda i: self._nav_to('attendance'))
        sb.add_widget(self._nav_attendance)

        sb.add_widget(Widget())

        user_box = UserBox(orientation='vertical', size_hint_y=None, height=dp(200), spacing=dp(6), padding=[dp(8), dp(8)])

        user_box.add_widget(_label("Logged in as", size=12, color=P["nav_item"], height=18, role="nav_item"))
        user_box.add_widget(_label(self.email, size=12, bold=True, color=P["text"], height=24, halign='left', role="text"))
        user_box.add_widget(Widget(size_hint_y=None, height=dp(4)))

        icon = "[font=Emoji]☀️[/font]" if CURRENT_THEME == 'dark' else "[font=Emoji]🌙[/font]"
        lbl = "Light Mode" if CURRENT_THEME == 'dark' else "Dark Mode"
        self._theme_btn = RoundBtn(text=f"{icon} {lbl}", role="selected", text_role="text",
                                 size_hint_y=None, height=dp(40))
        self._theme_btn.bind(on_press=self._toggle_theme)
        user_box.add_widget(self._theme_btn)

        logout = RoundBtn(text="[font=Emoji]🚪[/font] Logout", btn_color=list(P["error"]), role="error", size_hint_y=None, height=dp(40))
        logout.bind(on_press=lambda i: self._logout())
        user_box.add_widget(logout)

        self._status_dot = Widget(size_hint=(None, None), size=(dp(10), dp(10)), pos_hint={'center_y': 0.5})
        with self._status_dot.canvas:
            Color(*P["success"])
            self._status_ellipse = Ellipse(pos=self._status_dot.pos, size=self._status_dot.size)
        self._status_dot.bind(pos=lambda i, v: setattr(self._status_ellipse, 'pos', v), size=lambda i, v: setattr(self._status_ellipse, 'size', v))
        self._status_lbl = Label(text="ONLINE", font_size=sp(10), color=P["subtext"], halign='left', valign='middle', size_hint_y=None, height=dp(24))
        self._status_lbl.bind(size=self._status_lbl.setter('text_size'))
        s_row = BoxLayout(size_hint_y=None, height=dp(24), spacing=dp(8), padding=[dp(4), 0])
        s_row.add_widget(self._status_dot); s_row.add_widget(self._status_lbl)
        user_box.add_widget(s_row)
        sb.add_widget(user_box)
        return sb

    def _toggle_theme(self, *_):
        global CURRENT_THEME
        new = 'dark' if CURRENT_THEME == 'light' else 'light'
        try:
            with open(THEME_FILE, 'w') as f: f.write(new)
        except Exception: pass
        apply_theme(new)
        try: Window.clearcolor = list(P.get('bg', Window.clearcolor))
        except: pass
        icon = "[font=Emoji]☀️[/font]" if new == 'dark' else "[font=Emoji]🌙[/font]"
        lbl = "Light Mode" if new == 'dark' else "Dark Mode"
        self._theme_btn.text = f"{icon} {lbl}"
        self._apply_theme()

    def _ping(self):
        Clock.schedule_once(lambda dt: self._update_status(check_internet()),0)

    def _update_status(self, online):
        col = P["success"] if online else P["error"]
        self._status_dot.canvas.clear()
        with self._status_dot.canvas:
            Color(*col)
            self._status_ellipse = Ellipse(pos=self._status_dot.pos, size=self._status_dot.size)
        self._status_lbl.text  = "ONLINE" if online else "OFFLINE"
        self._status_lbl.color = P["subtext"] if online else P["error"]

    def _nav_to(self, page):
        self._nav_logbook.is_active = (page == 'logs')
        self._nav_attendance.is_active = (page == 'attendance')
        self._content_area.clear_widgets()

        # Stop attendance monitor if we navigate away from the attendance page
        if self._attendance_monitor_ev:
            self._attendance_monitor_ev.cancel()
            self._attendance_monitor_ev = None

        if page == 'logs':
            self._draw_logs()
        elif page == 'attendance':
            self._draw_attendance()
            # Start monitor when navigating to the attendance page
            if not self._attendance_monitor_ev:
                self._attendance_monitor_ev = Clock.schedule_interval(self._check_for_attendance_updates, 10) # Check every 10 seconds

    def _get_role_filtered_attendance(self):
        """Helper to filter attendance logs and columns based on user role and selected month."""
        is_admin = self.role in ("Super_Admin", "Admin")
        
        # Determine prefix for 'YYYY-MM' to match database strings
        try:
            target_prefix = datetime.strptime(self.selected_att_month, "%B %Y").strftime("%Y-%m")
        except:
            target_prefix = datetime.now().strftime("%Y-%m")

        # 1. Filter raw logs by the selected month first (r[1] is the Date column)
        logs_in_month = [r for r in self.attendance_logs if str(r[1]).startswith(target_prefix)]

        if is_admin:
            # Apply employee filter if a specific one is selected
            if self.selected_att_employee != "All Employees":
                return [r for r in logs_in_month if str(r[0]) == self.selected_att_employee]
            return logs_in_month

        # For PACD and User role: match name format (J.Cruz)
        fn = self.user_data.get('First Name', '').strip()
        ln = self.user_data.get('Last Name', '').strip()
        match_name = f"{fn[0]}.{ln}".lower() if fn and ln else ""

        data = []
        for r in logs_in_month:
            if str(r[0]).lower() == match_name:
                # New structure for staff: [Date, InAM, OutAM, InPM, OutPM, Remarks, ID]
                data.append(tuple(list(r[1:7]) + [r[7]]))
        return data

    def _draw_attendance(self):
        self._content_area.clear_widgets()
        is_admin = self.role in ("Super_Admin", "Admin")
        if not self.fetching_attendance and not self.att_data_loaded:
            self.fetching_attendance = True
            threading.Thread(target=self._fetch_attendance, daemon=True).start()

        topbar = SidebarBox(orientation='horizontal', size_hint_y=None, height=dp(68), padding=[dp(24),dp(12)], spacing=dp(10))
        topbar.add_widget(_label("Attendance", size=25, bold=True, role="text", height=40))
        topbar.add_widget(Widget())
        
        # Month Filter Dropdown
        months_found = {datetime.now().strftime("%B %Y")} # Always include current month
        for r in self.attendance_logs:
            try:
                dt = datetime.strptime(r[1], "%Y-%m-%d")
                months_found.add(dt.strftime("%B %Y"))
            except: pass
        sorted_months = sorted(list(months_found), key=lambda x: datetime.strptime(x, "%B %Y"), reverse=True)

        # Month Filter Label and Dropdown
        topbar.add_widget(_label("Filter Month:", size=14, color=P["subtext"], size_hint=(None, 1), width=dp(100), halign='right'))

        month_dropdown = SimpleDropdown(title="Filter Month", values=sorted_months, size_hint_x=None, width=dp(180), height=dp(50))
        month_dropdown.input.text = self.selected_att_month
        
        def _on_month_change(inst, val):
            if val != self.selected_att_month:
                self.selected_att_month = val
                self._att_table.set_data(self._get_role_filtered_attendance(), on_row_context=self._on_att_row_right_click)
        
        month_dropdown.bind(text=_on_month_change)
        topbar.add_widget(month_dropdown)

        # Employee Filter (Admin only)
        if is_admin:
            emps_found = {"All Employees"}
            for r in self.attendance_logs:
                if r[0]: emps_found.add(str(r[0]))
            sorted_emps = sorted(list(emps_found))
            
            topbar.add_widget(_label("Employee:", size=14, color=P["subtext"], size_hint=(None, 1), width=dp(80), halign='right'))
            emp_dropdown = SimpleDropdown(title="Filter Employee", values=sorted_emps, size_hint_x=None, width=dp(200), height=dp(50))
            emp_dropdown.input.text = self.selected_att_employee
            
            def _on_emp_change(inst, val):
                if val != self.selected_att_employee:
                    self.selected_att_employee = val
                    self._att_table.set_data(self._get_role_filtered_attendance(), on_row_context=self._on_att_row_right_click)
            emp_dropdown.bind(text=_on_emp_change)
            topbar.add_widget(emp_dropdown)

        search = SearchField()
        search.bind(text=self._filter_attendance)
        topbar.add_widget(search)

        self._content_area.add_widget(topbar)

        wrap = BoxLayout(orientation='vertical', padding=[dp(10),dp(12)], spacing=dp(10))
        
        headers = ATT_HEADERS if is_admin else ["Date", "Time In AM", "Time Out AM", "Time In PM", "Time Out PM", "Notes"]
        widths = ATT_COL_WIDTHS if is_admin else [dp(120), dp(120), dp(120), dp(120), dp(120), dp(200)]
        data = self._get_role_filtered_attendance()

        self._att_table = LogTableWidget(headers=headers, col_widths=widths)
        self._att_table.set_data(data, on_row_context=self._on_att_row_right_click)
        
        # Use FloatLayout to overlay the loading/empty message
        table_container = FloatLayout()
        self._att_table.pos_hint = {'x': 0, 'y': 0}
        table_container.add_widget(self._att_table)

        if not data:
            msg = "Loading data..." if self.fetching_attendance else "No attendance records found."
            lbl = _label(msg, size=16, color=P["subtext"], halign='center', role="subtext")
            lbl.pos_hint = {'center_x': 0.5, 'center_y': 0.4}
            table_container.add_widget(lbl)
            
        wrap.add_widget(table_container)
        self._content_area.add_widget(wrap)

    def _on_att_row_right_click(self, row_widget, pos):
        data = row_widget.row_data
        # Admin data: [Full Name, Date, InAM, OutAM, InPM, OutPM, Remarks, ID]
        # Staff data: [Date, InAM, OutAM, InPM, OutPM, Remarks, ID]
        box = BoxLayout(orientation='vertical', spacing=dp(10), padding=dp(10))
        btn_rem = RoundBtn(text="Add/Edit Notes")
        box.add_widget(btn_rem)
        
        pop = _make_popup("Options", box, w=300, h=170)
        
        row_widget.set_context_active(True)
        pop.bind(on_dismiss=lambda _: row_widget.set_context_active(False))
        
        def _do_rem(i):
            pop.dismiss()
            self._open_remarks_dialog(data)
        btn_rem.bind(on_press=_do_rem)
        pop.open()

    def _open_remarks_dialog(self, data):
        is_admin = self.role in ("Super_Admin", "Admin")
        current_rem = str(data[6 if is_admin else 5]) if data[6 if is_admin else 5] and str(data[6 if is_admin else 5]).lower() != 'none' else ""
        rec_id = data[7 if is_admin else 6]
        
        content = BoxLayout(orientation='vertical', spacing=dp(12), padding=dp(20))
        inp_rem = FloatingLabelInput(label_text="Notes")
        inp_rem.input.text = current_rem
        content.add_widget(inp_rem)
        
        btns = BoxLayout(size_hint_y=None, height=dp(44), spacing=dp(10))
        cancel = RoundBtn(text="Cancel", btn_color=list(P["error"]), role="error")
        save = RoundBtn(text="Save")
        btns.add_widget(cancel); btns.add_widget(save)
        content.add_widget(btns)
        
        pop = _make_popup("Update Notes", content, w=400, h=250)
        pop.auto_dismiss = False
        cancel.bind(on_press=pop.dismiss)
        
        def _go(_):
            new_rem = inp_rem.text.strip()
            save.text="Saving..."; save.disabled=True
            threading.Thread(target=self._save_remarks_thread, args=(rec_id, new_rem, pop, save), daemon=True).start()
        save.bind(on_press=_go)
        pop.open()

    def _save_remarks_thread(self, rec_id, remarks, pop, btn):
        import libsql_client
        try:
            cl = libsql_client.create_client_sync(TURSO_DB_URL, auth_token=TURSO_AUTH_TOKEN)
            cl.execute('UPDATE attendance SET Remarks=? WHERE id=?', (remarks, rec_id))
            cl.close()
            def _done(dt):
                pop.dismiss()
                show_message("Success", "Remarks updated!", msg_type="success")
                self.att_data_loaded = False
                self.attendance_logs = []
                self._draw_attendance()
            Clock.schedule_once(_done, 0)
        except Exception as e:
            Clock.schedule_once(lambda dt:(show_message("Error", str(e), msg_type="error"), setattr(btn,'text','Save'), setattr(btn,'disabled',False)), 0)

    def _connect_bg(self):
        if not check_internet():
            self.connecting=False; Clock.schedule_once(lambda dt:self._draw_logs(),0); return
        try:
            run_query("SELECT 1", fetch=True)
            if not self.user_data:
                try:
                    rows=run_query('SELECT "Email","Role","First_Name","Middle_Name","Last_Name",'
                                   '"Position","Salary","Salary_Grade","Status" FROM User_Permissions '
                                   'WHERE LOWER(Email)=LOWER(?)',params=(self.email,),fetch=True)
                    if rows:
                        self.user_data=dict(zip(USER_HEADERS,rows[0]))
                        self.role=str(self.user_data.get('Role','Staff')).strip()
                except: pass
            try: self._fetch_users()
            except: pass
            Clock.schedule_once(lambda dt:self._on_connected(),0)
        except Exception as e:
            print(f"Connection Error: {e}")
            self.connecting=False; Clock.schedule_once(lambda dt:self._draw_logs(),0)

    def _on_connected(self):
        self.connecting=False
        if not self.fetching_logs and not self.data_loaded_from_net:
            self.fetching_logs=True
            threading.Thread(target=self._fetch_logs, daemon=True).start()

    def _fetch_users(self):
        try:
            rows = run_query('SELECT "First_Name", "Middle_Name", "Last_Name" FROM User_Permissions', fetch=True)
            names = []
            if rows:
                for r in rows:
                    f = r[0].strip() if r[0] else ""
                    m = r[1].strip() if r[1] else ""
                    l = r[2].strip() if r[2] else ""
                    if m: m = f"{m[0]}."
                    full = " ".join([x for x in [f, m, l] if x])
                    if full: names.append(full)
            self.user_names = sorted(list(set(names)))
        except: pass

    def _fetch_logs(self):
        try:
            cy = datetime.now().strftime("%Y")
            sql = f'SELECT "Timestamp","REFERENCE_NUMBER","PARTICULARS","ADDRESSE","TRANSMITTER","SECTION","MODE_OF_TRANSMITTAL","REMARKS","ENCODED_BY" FROM Digital_Logbook WHERE "Timestamp" LIKE "{cy}%" ORDER BY rowid DESC'
            raw=run_query(sql,fetch=True)
            self.master_logs=raw if raw else []
        except Exception as e: print(f"Log fetch: {e}")
        self.fetching_logs=False; self.data_loaded_from_net=True
        Clock.schedule_once(lambda dt:self._draw_logs(),0)

    def _fetch_attendance(self):
        try:
            sql = 'SELECT full_name, date, time_in_am, time_out_am, time_in_pm, time_out_pm, Remarks, id FROM attendance ORDER BY date DESC'
            raw = run_query(sql, fetch=True)
            # Sanitize None values to empty strings for cleaner display
            self.attendance_logs = [tuple((x if x is not None else "") for x in r) for r in raw] if raw else []
        except Exception as e: print(f"Attendance fetch: {e}")
        self.fetching_attendance = False; self.att_data_loaded = True
        Clock.schedule_once(lambda dt: self._draw_attendance(), 0)

    def _check_for_attendance_updates(self, dt):
        """Called by a Clock schedule to check for new attendance records."""
        if self.fetching_attendance: return # Don't overlap checks
        threading.Thread(target=self._check_for_attendance_updates_thread, daemon=True).start()

    def _check_for_attendance_updates_thread(self):
        """Queries the database in a background thread for the latest update timestamp."""
        try:
            sql = "SELECT MAX(updated_at) FROM attendance"
            result = run_query(sql, fetch=True)
            
            if result and result[0] and result[0][0]:
                latest_update = result[0][0]
                
                # First time check, just store the timestamp
                if self.last_attendance_update is None:
                    self.last_attendance_update = latest_update
                    return

                # If a newer update is found, trigger a refresh on the main thread
                if latest_update > self.last_attendance_update:
                    self.last_attendance_update = latest_update
                    Clock.schedule_once(self._refresh_attendance)
        except Exception as e:
            print(f"Error checking for attendance updates: {e}")

    def _logout(self):
        try: os.remove(SESSION_FILE)
        except: pass
        if self._monitor_ev: self._monitor_ev.cancel()
        if self._attendance_monitor_ev: self._attendance_monitor_ev.cancel()
        App.get_running_app().show_login()

    def _open_export_dialog(self, *args):
        if not self.master_logs: return show_message("Info", "No data to export.", msg_type="info")
        
        years = set()
        for r in self.master_logs:
            if r and len(r) > 0:
                try: years.add(str(datetime.strptime(r[0], "%Y-%m-%d %H:%M").year))
                except: pass
        sorted_years = sorted(list(years), reverse=True)
        if not sorted_years: sorted_years = [datetime.now().strftime("%Y")]

        content = BoxLayout(orientation='vertical', spacing=dp(12), padding=dp(20))
        
        mode_spin = SelectionDropdown(label_text="Export Mode", values=["Current View", "By Year", "By Month"])
        mode_spin.input.text = "Current View"
        content.add_widget(mode_spin)

        opts_box = BoxLayout(orientation='vertical', spacing=dp(12), size_hint_y=None)
        opts_box.bind(minimum_height=opts_box.setter('height'))
        
        spin_year = SelectionDropdown(label_text="Select Year", values=sorted_years)
        spin_year.input.text = sorted_years[0]
        
        months = ["January","February","March","April","May","June","July","August","September","October","November","December"]
        spin_month = SelectionDropdown(label_text="Select Month", values=months)
        spin_month.input.text = datetime.now().strftime("%B")

        def _upd_opts(inst, val):
            opts_box.clear_widgets()
            if val in ("By Year", "By Month"):
                opts_box.add_widget(spin_year)
            if val == "By Month":
                opts_box.add_widget(spin_month)
        mode_spin.bind(text=_upd_opts)
        _upd_opts(None, mode_spin.text)
        content.add_widget(opts_box)
        content.add_widget(Widget())

        btns = BoxLayout(size_hint_y=None, height=dp(44), spacing=dp(10))
        cancel = RoundBtn(text="Cancel", btn_color=list(P["error"]), role="error")
        export = RoundBtn(text="Export CSV", btn_color=list(P["success"]), role="success")
        btns.add_widget(cancel); btns.add_widget(export)
        content.add_widget(btns)

        pop = _make_popup("Export Data", content, w=350, h=450)
        pop.auto_dismiss = False
        cancel.bind(on_press=pop.dismiss)

        def _do_export(_):
            mode, data, suffix = mode_spin.text, [], ""
            if mode == "Current View":
                data = [r for r in self.master_logs if len(r)>0 and str(r[0]).startswith(datetime.now().strftime("%Y"))]
                suffix = "CurrentView"
            elif mode == "By Year":
                data = [r for r in self.master_logs if len(r)>0 and str(r[0]).startswith(spin_year.text)]
                suffix = f"Year_{spin_year.text}"
            elif mode == "By Month":
                try:
                    mn = datetime.strptime(spin_month.text, "%B").month
                    data = [r for r in self.master_logs if len(r)>0 and str(r[0]).startswith(f"{spin_year.text}-{mn:02d}")]
                    suffix = f"{spin_month.text}_{spin_year.text}"
                except: pass
            
            if not data: return show_message("Info", "No records found.", msg_type="info")
            try:
                fn = f"Logbook_{suffix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
                path = os.path.join(os.path.expanduser("~"), "Desktop", fn)
                with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.writer(f); writer.writerow(self.headers); writer.writerows(data)
                pop.dismiss(); show_message("Success", f"Exported to Desktop:\n{fn}", msg_type="success")
                try: os.startfile(path)
                except: pass
            except Exception as e: show_message("Error", f"Failed to export: {e}", msg_type="error")
        export.bind(on_press=_do_export); pop.open()

    def _filter_logs(self, instance, text):
        if not hasattr(self, '_log_table') or not self._log_table: return
        text = text.lower().strip()
        current_year = datetime.now().strftime("%Y")
        
        rows = []
        if self.master_logs:
            matches = []
            for r in self.master_logs:
                if len(r) > 0 and str(r[0]).startswith(current_year):
                    if not text or any(text in str(c).lower() for c in r):
                        matches.append(r)
            rows = sorted(matches, key=lambda x: x[1] if len(x)>1 else "", reverse=True)
        self._log_table.set_data(rows, on_row_context=self._on_row_right_click)

    def _filter_attendance(self, instance, text):
        if not hasattr(self, '_att_table') or not self._att_table: return
        text = text.lower().strip()
        filtered_data = self._get_role_filtered_attendance()
        rows = []
        if filtered_data:
            if not text: rows = filtered_data
            else:
                rows = [r for r in filtered_data if any(text in str(c).lower() for c in r)]
        self._att_table.set_data(rows, on_row_context=self._on_att_row_right_click)

    def _draw_logs(self):
        self._content_area.clear_widgets()
        if not self.fetching_logs and not self.data_loaded_from_net:
            self.fetching_logs=True
            threading.Thread(target=self._fetch_logs, daemon=True).start()

        topbar = SidebarBox(orientation='horizontal', size_hint_y=None, height=dp(68), padding=[dp(24),dp(12)], spacing=dp(10))
        topbar.add_widget(_label("Digital Logbook", size=25, bold=True, role="text", height=40))
        topbar.add_widget(Widget())

        search = SearchField()
        search.bind(text=self._filter_logs)
        topbar.add_widget(search)

        add_btn = RoundBtn(text="[font=Emoji]➕[/font] New Record", size_hint_x=None, width=dp(130))
        add_btn.bind(on_press=self._open_add_dialog)
        topbar.add_widget(add_btn)

        exp_btn = RoundBtn(text="[font=Emoji]📥[/font] Export CSV", btn_color=list(P["success"]), role="success", size_hint_x=None, width=dp(120))
        exp_btn.bind(on_press=self._open_export_dialog); topbar.add_widget(exp_btn)

        self._content_area.add_widget(topbar)

        wrap = BoxLayout(orientation='vertical', padding=[dp(10), dp(12)], spacing=dp(10))

        # Always create table structure so headers are visible
        self._log_table = LogTableWidget(headers=self.headers)
        
        current_year = datetime.now().strftime("%Y")
        rows = []
        if self.master_logs:
            rows = sorted([r for r in self.master_logs if len(r)>0 and str(r[0]).startswith(current_year)], key=lambda x: x[1] if len(x)>1 else "", reverse=True)
        self._log_table.set_data(rows, on_row_context=self._on_row_right_click)

        # Use FloatLayout to overlay the "Loading" or "No Data" message on top of the table
        table_container = FloatLayout()
        self._log_table.pos_hint = {'x': 0, 'y': 0}
        table_container.add_widget(self._log_table)

        if not rows:
            msg = "Loading data…" if (self.fetching_logs or self.connecting) else ("No records found for " + current_year if self.master_logs else "No data available.")
            lbl = _label(msg, size=16, color=P["subtext"], halign='center', role="subtext")
            # Position slightly below center to not obscure headers
            lbl.pos_hint = {'center_x': 0.5, 'center_y': 0.4}
            table_container.add_widget(lbl)

        wrap.add_widget(table_container)

        self._content_area.add_widget(wrap)

    def _on_row_right_click(self, row_widget, pos):
        data = row_widget.row_data
        box = BoxLayout(orientation='vertical', spacing=dp(10), padding=dp(10))
        btn_edit = RoundBtn(text="Edit Record")
        box.add_widget(btn_edit)
        
        can_backdate = self.role in ("Super_Admin", "Admin", "PACD")
        if can_backdate:
            btn_sub  = RoundBtn(text="Insert Backdated Record", btn_color=list(P["error"]), role="error")
            box.add_widget(btn_sub)
        
        pop = _make_popup("Options", box, w=300, h=220 if can_backdate else 160)

        # Set row as active and bind to popup dismiss
        row_widget.set_context_active(True)
        def on_popup_dismiss(instance):
            row_widget.set_context_active(False)
        pop.bind(on_dismiss=on_popup_dismiss)

        def _do_edit(i):
            pop.dismiss()
            is_owner = len(data) > 8 and str(data[8]).strip().lower() == str(self.email).strip().lower()
            if self.role == "Super_Admin" or is_owner:
                self._open_edit_dialog(data)
            else:
                show_message("Access Denied", "You can only edit records you encoded.", msg_type="error")
            
        btn_edit.bind(on_press=_do_edit)
        
        if can_backdate:
            def _do_sub(i):
                pop.dismiss()
                self._open_backdate_dialog(data)
            btn_sub.bind(on_press=_do_sub)
            
        pop.open()

    def _open_backdate_dialog(self, data):
        base_ref = str(data[1])
        ts_orig = str(data[0])
        date_val = ts_orig.split(" ")[0] if " " in ts_orig else ts_orig
        time_val = ts_orig.split(" ")[1] if " " in ts_orig else ""

        content = BoxLayout(orientation='vertical', spacing=dp(12), padding=dp(25))
        sv = ScrollView(bar_width=dp(8), bar_margin=dp(4))
        form = BoxLayout(orientation='vertical', spacing=dp(15), size_hint_y=None, padding=[dp(10), dp(10), dp(15), 0])
        form.bind(minimum_height=form.setter('height'))
        
        # Combined Top Row for Ref and Date/Time
        top_row = BoxLayout(orientation='horizontal', spacing=dp(15), size_hint_y=None, height=dp(88), padding=[dp(10), 0, 0, 0])

        # Reference
        ref_container = BoxLayout(orientation='vertical', spacing=dp(4))
        ref_container.add_widget(_label("Reference No.:", size=13, bold=True, color=P["text"], height=18, role="text"))
        ref_row = BoxLayout(orientation='horizontal', spacing=dp(10))
        lbl_base = Label(text=base_ref, color=P["text"], font_size=sp(16), bold=True,
                         size_hint_x=None, width=dp(100), halign='left', valign='middle')
        lbl_base.bind(size=lbl_base.setter('text_size'))
        ref_row.add_widget(lbl_base)
        inp_suffix = FieldInput(hint_text="Ex. A", size_hint_x=None, pos_hint={'center_y': 0.5})
        ref_row.add_widget(inp_suffix)
        ref_container.add_widget(ref_row)
        top_row.add_widget(ref_container)

        # Timestamp
        ts_container = BoxLayout(orientation='vertical', spacing=dp(4))
        ts_container.add_widget(_label("Date & Time:", size=13, bold=True, color=P["text"], height=18, role="text"))
        ts_row = BoxLayout(orientation='horizontal')
        lbl_date = Label(text=date_val, color=P["text"], font_size=sp(16), bold=True, halign='left', valign='middle', size_hint_x=None, width=dp(100))
        lbl_date.bind(size=lbl_date.setter('text_size'))
        inp_time = FieldInput(hint_text="Time (HH:MM)", size_hint_x=None, pos_hint={'center_y': 0.5})
        inp_time.text = time_val
        ts_row.add_widget(lbl_date); ts_row.add_widget(inp_time)
        ts_container.add_widget(ts_row)
        top_row.add_widget(ts_container)
        form.add_widget(top_row)

        # Blank Fields
        inp_part = FloatingLabelInput(label_text="Particulars")
        form.add_widget(inp_part)
        inp_addr = FloatingLabelInput(label_text="Addressee")
        form.add_widget(inp_addr)
        inp_trans = SelectionDropdown(label_text="Transmitter", values=self.user_names)
        form.add_widget(inp_trans)
        inp_sec = SelectionDropdown(label_text="Section", values=["PSO", "Admin", "CRS", "PhilSys", "Statistical"])
        form.add_widget(inp_sec)
        inp_mode = MultiSelectDropdown("Mode of Transmittal", ["Email", "Walk-in", "Hand Carry", "JRS", "Routing", "Google Link"])
        form.add_widget(inp_mode)
        inp_rem = FloatingLabelInput(label_text="Remarks (Optional)")
        form.add_widget(inp_rem)
        
        sv.add_widget(form); content.add_widget(sv)
        
        btns = BoxLayout(size_hint_y=None, height=dp(44), spacing=dp(10))
        cancel = RoundBtn(text="Cancel", btn_color=list(P["error"]), role="error")
        save = RoundBtn(text="Save Record")
        btns.add_widget(cancel); btns.add_widget(save)
        content.add_widget(btns)
        
        pop = _make_popup(f"Insert Earlier or Later Record on this Ref. No: {data[1]}", content, w=560, h=720)
        pop.auto_dismiss = False; cancel.bind(on_press=pop.dismiss)
        
        def _go(_):
            suf = inp_suffix.text.strip()
            if not suf: return show_message("Error","Suffix required (e.g. A).",msg_type="warning")
            
            tv = inp_time.text.strip()
            try:
                th, tm = map(int, tv.split(':'))
                if not (0 <= th <= 23 and 0 <= tm <= 59): raise ValueError
            except: return show_message("Error","Invalid Time. Use HH:MM (00-23, 00-59).",msg_type="warning")

            ts = f"{date_val} {tv}".strip()
            if len(ts)<10: return show_message("Error","Date/Time required.",msg_type="warning")
            
            if ts == ts_orig: return show_message("Error","Time must be different from original.",msg_type="warning")
            
            d={"ref":base_ref+suf, "ts":ts, "part":inp_part.text.strip(),
               "addr":inp_addr.text.strip(), "trans":inp_trans.input.text.strip(),
               "sec":inp_sec.text, "mode":inp_mode.get_value(), "rem":inp_rem.text.strip()}
            
            if not all([d["part"],d["addr"],d["trans"],d["sec"],d["mode"]]):
                return show_message("Missing Fields","Please fill required fields.",msg_type="warning")
            
            save.text="Saving..."; save.disabled=True
            threading.Thread(target=self._save_backdate_thread, args=(d,pop,save), daemon=True).start()
        
        save.bind(on_press=_go); pop.open()

    def _save_backdate_thread(self, d, pop, btn):
        import libsql_client
        try:
            cl = libsql_client.create_client_sync(TURSO_DB_URL, auth_token=TURSO_AUTH_TOKEN)
            sql = 'INSERT INTO Digital_Logbook ("TIMESTAMP", "REFERENCE_NUMBER", "PARTICULARS", "ADDRESSE", "TRANSMITTER", "SECTION", "MODE_OF_TRANSMITTAL", "REMARKS", "ENCODED_BY") VALUES (?,?,?,?,?,?,?,?,?)'
            cl.batch([(sql, (d["ts"], d["ref"], d["part"], d["addr"], d["trans"], d["sec"], d["mode"], d["rem"], self.email))])
            cl.close()
            Clock.schedule_once(lambda dt:(pop.dismiss(), show_message("Success",f"Saved!\nRef: {d['ref']}",msg_type="success"), self._refresh_logs()), 0)
        except Exception as e:
            msg = str(e)
            Clock.schedule_once(lambda dt:(show_message("Error",msg,msg_type="error"), setattr(btn,'text','Save Record'), setattr(btn,'disabled',False)), 0)

    def _open_edit_dialog(self, data):
        content = BoxLayout(orientation='vertical', spacing=dp(12), padding=dp(25))
        sv = ScrollView(bar_width=dp(8), bar_margin=dp(4))
        form = BoxLayout(orientation='vertical', spacing=dp(12), size_hint_y=None, padding=[0, dp(10)])
        form.bind(minimum_height=form.setter('height'))

        inp_part = FloatingLabelInput(label_text="Particulars")
        inp_part.input.text = str(data[2])
        form.add_widget(inp_part)
        
        inp_addr = FloatingLabelInput(label_text="Addressee")
        inp_addr.input.text = str(data[3])
        form.add_widget(inp_addr)

        inp_trans = SelectionDropdown(label_text="Transmitter", values=self.user_names)
        inp_trans.input.text = str(data[4])
        form.add_widget(inp_trans)
        
        inp_sec = SelectionDropdown(label_text="Section", values=["PSO", "Admin", "CRS", "PhilSys", "Statistical"])
        inp_sec.input.text = str(data[5])
        form.add_widget(inp_sec)
        
        inp_mode = MultiSelectDropdown("Mode of Transmittal", ["Email", "Walk-in", "Hand Carry", "JRS", "Routing", "Google Link"])
        inp_mode.set_value(str(data[6]))
        form.add_widget(inp_mode)
        
        inp_rem = FloatingLabelInput(label_text="Remarks (Optional)")
        inp_rem.input.text = str(data[7])
        form.add_widget(inp_rem)
        
        sv.add_widget(form); content.add_widget(sv)
        
        btns = BoxLayout(size_hint_y=None, height=dp(44), spacing=dp(10))
        cancel = RoundBtn(text="Cancel", btn_color=list(P["error"]), role="error")
        save = RoundBtn(text="Update Record")
        save.disabled = True
        btns.add_widget(cancel)
        btns.add_widget(save)
        content.add_widget(btns)
        
        pop = _make_popup(f"Edit the Record with Ref. No: {data[1]}", content, w=520, h=650)
        pop.auto_dismiss = False; cancel.bind(on_press=pop.dismiss)
        
        orig_vals = {
            "part": str(data[2]).strip(), "addr": str(data[3]).strip(),
            "trans": str(data[4]).strip(), "sec": str(data[5]).strip(),
            "mode": str(data[6]).strip(), "rem": str(data[7]).strip()
        }
        def check_changes(*_):
            curr = {
                "part": inp_part.text.strip(), "addr": inp_addr.text.strip(),
                "trans": inp_trans.text.strip(), "sec": inp_sec.text.strip(),
                "mode": inp_mode.text.strip(), "rem": inp_rem.text.strip()
            }
            save.disabled = not any(curr[k] != orig_vals[k] for k in curr)
            
        for w in (inp_part, inp_addr, inp_trans, inp_sec, inp_mode, inp_rem):
            w.bind(text=check_changes)

        def _go(_):
            new_d = {"part":inp_part.text.strip(), "addr":inp_addr.text.strip(), "trans":inp_trans.input.text.strip(),
                     "sec":inp_sec.text, "mode":inp_mode.get_value(), "rem":inp_rem.text.strip(), "ref": data[1]}
            if not all([new_d["part"], new_d["addr"], new_d["trans"], new_d["sec"], new_d["mode"]]):
                return show_message("Missing Fields", "Please fill all required fields.", msg_type="warning")
            save.text = "Updating..."
            save.disabled = True
            threading.Thread(target=self._update_record_thread, args=(new_d, pop, save), daemon=True).start()
        save.bind(on_press=_go); pop.open()

    def _open_add_dialog(self, *args):
        content = BoxLayout(orientation='vertical', spacing=dp(12), padding=dp(25))
        sv = ScrollView(bar_width=dp(8), bar_margin=dp(4))
        form = BoxLayout(orientation='vertical', spacing=dp(12), size_hint_y=None, padding=[0, dp(10)])
        form.bind(minimum_height=form.setter('height'))
        
        inp_part = FloatingLabelInput(label_text="Particulars")
        form.add_widget(inp_part)
        inp_addr = FloatingLabelInput(label_text="Addressee")
        form.add_widget(inp_addr)
        
        # Transmitter
        inp_trans = SelectionDropdown(label_text="Transmitter", values=self.user_names)
        form.add_widget(inp_trans)
        
        inp_sec = SelectionDropdown(label_text="Section", values=["PSO", "Admin", "CRS", "PhilSys", "Statistical"])
        form.add_widget(inp_sec)
        
        inp_mode = MultiSelectDropdown("Mode of Transmittal", ["Email", "Walk-in", "Hand Carry", "JRS", "Routing", "Google Link"])
        form.add_widget(inp_mode)
        
        inp_rem = FloatingLabelInput(label_text="Remarks (Optional)")
        form.add_widget(inp_rem)
        
        sv.add_widget(form); content.add_widget(sv)
        
        btns = BoxLayout(size_hint_y=None, height=dp(44), spacing=dp(10))
        cancel = RoundBtn(text="Cancel", btn_color=list(P["error"]), role="error")
        save = RoundBtn(text="Save Record")
        btns.add_widget(cancel)
        btns.add_widget(save)
        content.add_widget(btns)
        
        pop = _make_popup("New Record", content, w=520, h=650)
        pop.auto_dismiss = False; cancel.bind(on_press=pop.dismiss)
        def _go(_):
            d={"part":inp_part.text.strip(),"addr":inp_addr.text.strip(),"trans":inp_trans.input.text.strip(),
               "sec":inp_sec.text,"mode":inp_mode.get_value(),"rem":inp_rem.text.strip()}
            if not all([d["part"],d["addr"],d["trans"],d["sec"],d["mode"]]):
                return show_message("Missing Fields","Please fill all required fields.",msg_type="warning")
            save.text="Saving..."; save.disabled=True
            threading.Thread(target=self._save_record_thread, args=(d,pop,save), daemon=True).start()
        save.bind(on_press=_go); pop.open()

    def _save_record_thread(self, data, pop, btn):
        import libsql_client
        try:
            cl = libsql_client.create_client_sync(TURSO_DB_URL, auth_token=TURSO_AUTH_TOKEN)
            rs = cl.batch([('INSERT INTO Digital_Logbook ("PARTICULARS","ADDRESSE","TRANSMITTER","SECTION","MODE_OF_TRANSMITTAL","REMARKS","ENCODED_BY") VALUES (?,?,?,?,?,?,?)',
                          (data["part"],data["addr"],data["trans"],data["sec"],data["mode"],data["rem"],self.email))])[0]
            ref = cl.execute('SELECT "REFERENCE_NUMBER" FROM Digital_Logbook WHERE rowid=?',(rs.last_insert_rowid,)).rows[0][0]
            cl.close()
            Clock.schedule_once(lambda dt:(pop.dismiss(), show_message("Success",f"Record Saved!\nReference No: {ref}",msg_type="success"), self._refresh_logs()),0)
        except Exception as e:
            msg = str(e)
            Clock.schedule_once(lambda dt:(show_message("Error",msg,msg_type="error"), setattr(btn,'text','Save Record'), setattr(btn,'disabled',False)),0)

    def _update_record_thread(self, d, pop, btn):
        import libsql_client
        try:
            cl = libsql_client.create_client_sync(TURSO_DB_URL, auth_token=TURSO_AUTH_TOKEN)
            sql = 'UPDATE Digital_Logbook SET "PARTICULARS"=?, "ADDRESSE"=?, "TRANSMITTER"=?, "SECTION"=?, "MODE_OF_TRANSMITTAL"=?, "REMARKS"=? WHERE "REFERENCE_NUMBER"=?'
            cl.execute(sql, (d["part"], d["addr"], d["trans"], d["sec"], d["mode"], d["rem"], d["ref"]))
            cl.close()
            Clock.schedule_once(lambda dt: (pop.dismiss(), show_message("Success", "Record Updated!", msg_type="success"), self._refresh_logs()), 0)
        except Exception as e:
            Clock.schedule_once(lambda dt: (show_message("Error", str(e), msg_type="error"), setattr(btn, 'text', 'Update Record'), setattr(btn, 'disabled', False)), 0)

    def _apply_theme(self):
        try:
            if hasattr(self, '_sidebar') and hasattr(self._sidebar, 'refresh_theme'): self._sidebar.refresh_theme()
            if hasattr(self, '_content_area') and hasattr(self._content_area, 'refresh_theme'): self._content_area.refresh_theme()

            for w in list(self.walk()):
                try:
                    if hasattr(w, 'refresh_theme') and callable(w.refresh_theme):
                        w.refresh_theme(); continue
                    if isinstance(w, Label):
                        role = getattr(w, 'role', 'text')
                        if role != 'ignore' and role in P:
                            w.color = P[role]
                    if hasattr(w, 'bg_color'): w.bg_color = list(P.get('card', w.bg_color))
                    if hasattr(w, 'border_color'): w.border_color = list(P.get('border', w.border_color))
                    if hasattr(w, 'btn_color') and str(getattr(w, 'text', '')).lower() not in ('logout','export csv'):
                        w.btn_color = list(P.get('primary', w.btn_color))
                except: pass

            try: self._update_status(check_internet())
            except: pass
        except: pass

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  App
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

class KalingaOpsHUBApp(App):
    title = "Kalinga OpsHUB"

    def build(self):
        self.icon = AGENCY_LOGO
        Clock.schedule_once(lambda dt: Window.maximize(), 0)
        try:
            if os.path.exists(THEME_FILE):
                t = open(THEME_FILE).read().strip()
                if t in ('light','dark'): apply_theme(t)
        except Exception: pass
        Window.clearcolor = list(P["bg"])

        self.sm = ScreenManager(transition=FadeTransition(duration=0.18))
        self.sm.add_widget(LoginScreen(name='login'))

        if os.path.exists(SESSION_FILE):
            try:
                saved = open(SESSION_FILE).read().strip()
                if saved:
                    ud=None
                    if os.path.exists(CACHE_FILE):
                        try: ud=json.load(open(CACHE_FILE)).get(saved)
                        except: pass
                    Clock.schedule_once(lambda dt:self.launch_main(saved,ud), 0.2)
            except: pass

        return self.sm

    def launch_main(self, email, user_data=None):
        if user_data:
            try:
                c={}
                if os.path.exists(CACHE_FILE):
                    try: c=json.load(open(CACHE_FILE))
                    except: pass
                c[email]=user_data
                json.dump(c, open(CACHE_FILE,'w'))
            except Exception as e: print(f"Cache write: {e}")

        if self.sm.has_screen('main'): self.sm.remove_widget(self.sm.get_screen('main'))
        self.sm.add_widget(MainScreen(email=email, user_data=user_data, name='main'))
        self.sm.current='main'

    def show_login(self):
        if self.sm.has_screen('main'): self.sm.remove_widget(self.sm.get_screen('main'))
        self.sm.current='login'


if __name__ == "__main__":
    KalingaOpsHUBApp().run()