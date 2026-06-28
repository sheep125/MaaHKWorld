"""
启动器辅助模块 - 处理游戏闪退后通过启动器重新启动游戏

王者荣耀世界直接启动游戏.exe会闪退，之后自动弹出启动器。
本模块负责：检测启动器窗口 → OCR识别"启动游戏"按钮 → 点击 → 等待游戏窗口出现
"""
import time
from utils.logger import log

try:
    import win32gui
    import win32con
    import win32api
    import win32com.client
    _WIN32_AVAILABLE = True
except ImportError:
    win32gui = None
    win32con = None
    win32api = None
    win32com = None
    _WIN32_AVAILABLE = False


# 可能的启动器窗口标题（按优先级）
LAUNCHER_WINDOW_TITLES = [
    "王者荣耀世界",
    "王者荣耀世界启动器",
    "Honor of Kings: World",
    "Launcher",
]

# 可能的启动器窗口类名
LAUNCHER_WINDOW_CLASSES = [
    "Qt5152QWindowIcon",
    "Qt5QWindowIcon",
    "#32770",
    None,  # 不限制class，只匹配title
]


def _find_window_by_title(keywords, exclude_class_regex=None):
    """按标题关键字查找窗口
    
    Args:
        keywords: 标题关键字列表
        exclude_class_regex: 要排除的窗口类名正则（如排除UnrealWindow游戏窗口）
    
    Returns:
        找到的窗口句柄，未找到返回None
    """
    found_hwnd = None
    
    def callback(hwnd, extra):
        nonlocal found_hwnd
        if not win32gui.IsWindowVisible(hwnd):
            return True
        title = win32gui.GetWindowText(hwnd)
        if not title:
            return True
        
        # 排除自身进程的窗口
        for kw in keywords:
            if kw in title:
                found_hwnd = hwnd
                return False  # 停止枚举
        return True
    
    win32gui.EnumWindows(callback, None)
    return found_hwnd


def find_launcher_window():
    """查找启动器窗口
    
    Returns:
        启动器窗口句柄，未找到返回None
    """
    if not _WIN32_AVAILABLE:
        log("[Launcher] win32gui not available")
        return None
    
    # 先尝试匹配特定窗口类 + 标题
    for cls in LAUNCHER_WINDOW_CLASSES:
        for title_kw in LAUNCHER_WINDOW_TITLES:
            try:
                if cls:
                    hwnd = win32gui.FindWindow(cls, title_kw)
                else:
                    hwnd = win32gui.FindWindow(None, title_kw)
                if hwnd:
                    actual_title = win32gui.GetWindowText(hwnd)
                    log(f"[Launcher] Found launcher window: HWND={hwnd}, class={cls}, title='{actual_title}'")
                    return hwnd
            except Exception:
                continue
    
    # 如果精确匹配失败，使用枚举方式查找
    # 查找标题包含关键字的窗口，但排除 UnrealWindow（游戏本体窗口）
    log("[Launcher] Exact match failed, trying window enumeration...")
    found = []
    
    def enum_callback(hwnd, extra):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        
        title = win32gui.GetWindowText(hwnd)
        cls_name = win32gui.GetClassName(hwnd)
        
        # 跳过游戏窗口（UnrealWindow）
        if "UnrealWindow" in cls_name:
            return True
        
        for kw in LAUNCHER_WINDOW_TITLES:
            if kw in title:
                found.append((hwnd, title, cls_name))
                break
        return True
    
    win32gui.EnumWindows(enum_callback, None)
    
    if found:
        hwnd, title, cls_name = found[0]
        log(f"[Launcher] Found via enumeration: HWND={hwnd}, title='{title}', class='{cls_name}'")
        return hwnd
    
    # 最后手段：查找任何包含"世界"或"World"的非游戏窗口
    def enum_callback2(hwnd, extra):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        
        title = win32gui.GetWindowText(hwnd)
        cls_name = win32gui.GetClassName(hwnd)
        
        if "UnrealWindow" in cls_name:
            return True
        
        if "世界" in title or "启动" in title:
            found.append((hwnd, title, cls_name))
        return True
    
    win32gui.EnumWindows(enum_callback2, None)
    
    if found:
        hwnd, title, cls_name = found[0]
        log(f"[Launcher] Found via fallback: HWND={hwnd}, title='{title}', class='{cls_name}'")
        return hwnd
    
    log("[Launcher] No launcher window found")
    return None


def is_game_window_alive():
    """检查游戏窗口是否还活着（没有闪退）
    
    Returns:
        bool: True 如果游戏窗口存在
    """
    if not _WIN32_AVAILABLE:
        return False
    
    try:
        hwnd = win32gui.FindWindow("UnrealWindow", "王者荣耀世界")
        if hwnd and win32gui.IsWindow(hwnd):
            return True
    except Exception:
        pass
    return False


def find_start_button_by_ocr():
    """使用OCR在桌面上查找"启动游戏"按钮
    
    由于启动器窗口可能不是标准Win32控件，使用全屏截图+OCR识别"启动游戏"文字
    
    Returns:
        (x, y) 按钮中心坐标，未找到返回None
    """
    log("[Launcher] Searching for '启动游戏' button via OCR...")
    
    try:
        from PIL import ImageGrab
        import win32gui
        import win32con
        
        launcher_hwnd = find_launcher_window()
        
        if launcher_hwnd:
            # 聚焦启动器窗口
            try:
                if win32gui.IsIconic(launcher_hwnd):
                    win32gui.ShowWindow(launcher_hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(launcher_hwnd)
                time.sleep(0.5)
            except Exception as e:
                log(f"[Launcher] Failed to focus launcher window: {e}")
            
            # 获取启动器窗口位置，裁剪截图区域
            try:
                rect = win32gui.GetWindowRect(launcher_hwnd)
                left, top, right, bottom = rect
                # 扩大一点区域
                left = max(0, left - 10)
                top = max(0, top - 10)
                right = right + 10
                bottom = bottom + 10
                log(f"[Launcher] Launcher window rect: ({left}, {top}, {right}, {bottom})")
                screenshot = ImageGrab.grab(bbox=(left, top, right, bottom))
                offset_x, offset_y = left, top
            except Exception as e:
                log(f"[Launcher] Failed to get launcher rect, using full screen: {e}")
                screenshot = ImageGrab.grab()
                offset_x, offset_y = 0, 0
        else:
            log("[Launcher] Launcher not found, using full screen OCR")
            screenshot = ImageGrab.grab()
            offset_x, offset_y = 0, 0
        
        # 尝试使用 MaaFramework 的 OCR (通过 import 方式，如果可用)
        # 这里使用一个更通用的方法：先尝试 maa 的 OCR，失败则用键盘导航
        try:
            import maa.toolkit
            from maa.resource import Resource
            from maa.controller import AdbController  # 这里不用ADB，只是尝试导入
        except ImportError:
            log("[Launcher] MaaFramework OCR not available for launcher detection")
        
        # 使用 Windows 可访问性 API 或简单的图像处理
        # 简化方案：直接使用键盘模拟操作
        
        # 先尝试用简单的图像匹配来找按钮
        import cv2
        import numpy as np
        
        img = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
        # 尝试使用 pytesseract 做 OCR
        try:
            import pytesseract
            data = pytesseract.image_to_data(gray, lang='chi_sim', output_type=pytesseract.Output.DICT)
            
            for i, text in enumerate(data['text']):
                if '启动' in text and ('游戏' in text or data['text'][i+1] == '游戏' if i+1 < len(data['text']) else False):
                    x = data['left'][i] + data['width'][i] // 2 + offset_x
                    y = data['top'][i] + data['height'][i] // 2 + offset_y
                    conf = data['conf'][i]
                    log(f"[Launcher] Found '启动游戏' at ({x}, {y}), confidence={conf}")
                    return (x, y)
            
            # 搜索单独的"启动"或"游戏"
            for i, text in enumerate(data['text']):
                if text.strip() in ['启动', '开始游戏', '进入游戏', '启动游戏']:
                    x = data['left'][i] + data['width'][i] // 2 + offset_x
                    y = data['top'][i] + data['height'][i] // 2 + offset_y
                    conf = data['conf'][i]
                    log(f"[Launcher] Found '{text}' at ({x}, {y}), confidence={conf}")
                    return (x, y)
                    
        except ImportError:
            log("[Launcher] pytesseract not available, using fallback method")
        except Exception as e:
            log(f"[Launcher] OCR failed: {e}")
        
        log("[Launcher] OCR-based button detection failed")
        return None
        
    except ImportError as e:
        log(f"[Launcher] Required module not available: {e}")
        return None
    except Exception as e:
        log(f"[Launcher] Error searching for start button: {e}")
        return None


def find_start_button_by_keyboard():
    """使用键盘导航方式的 fallback: 直接模拟Tab+Enter
    
    当OCR不可用时，尝试用键盘在启动器中操作。
    许多游戏启动器默认选中"启动游戏"按钮，按Enter即可。
    
    Returns:
        True 如果执行了键盘操作
    """
    log("[Launcher] Attempting keyboard-based launcher navigation...")
    
    try:
        import ctypes
        
        launcher_hwnd = find_launcher_window()
        if not launcher_hwnd:
            log("[Launcher] Cannot find launcher for keyboard navigation")
            return False
        
        # 聚焦启动器
        win32gui.SetForegroundWindow(launcher_hwnd)
        time.sleep(0.5)
        
        # 发送 Enter 键（大多数启动器默认按钮就是"启动游戏"）
        user32 = ctypes.windll.user32
        
        # 按下 Enter
        user32.keybd_event(0x0D, 0, 0, 0)  # VK_RETURN
        time.sleep(0.05)
        user32.keybd_event(0x0D, 0, 2, 0)  # KEYEVENTF_KEYUP
        
        log("[Launcher] Sent Enter key to launcher")
        return True
        
    except Exception as e:
        log(f"[Launcher] Keyboard navigation failed: {e}")
        return False


def click_at_position(x, y):
    """在指定位置执行鼠标点击
    
    Args:
        x, y: 屏幕坐标
    """
    if not _WIN32_AVAILABLE:
        log("[Launcher] win32api not available for clicking")
        return False
    
    try:
        import ctypes
        
        # 保存当前鼠标位置
        user32 = ctypes.windll.user32
        class POINT(ctypes.Structure):
            _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]
        old_pos = POINT()
        user32.GetCursorPos(ctypes.byref(old_pos))
        
        # 移动鼠标到目标位置
        win32api.SetCursorPos((int(x), int(y)))
        time.sleep(0.1)
        
        # 执行点击
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.05)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        
        log(f"[Launcher] Clicked at ({x}, {y})")
        
        # 恢复鼠标位置
        time.sleep(0.1)
        user32.SetCursorPos(old_pos.x, old_pos.y)
        
        return True
        
    except Exception as e:
        log(f"[Launcher] Click failed: {e}")
        return False


def wait_for_game_window(timeout=30.0, check_interval=0.5):
    """等待游戏窗口出现
    
    Args:
        timeout: 最大等待时间（秒）
        check_interval: 检查间隔（秒）
    
    Returns:
        bool: True 如果游戏窗口出现
    """
    log(f"[Launcher] Waiting for game window (timeout={timeout}s)...")
    
    start_time = time.time()
    
    while time.time() - start_time < timeout:
        if is_game_window_alive():
            elapsed = time.time() - start_time
            log(f"[Launcher] Game window appeared after {elapsed:.1f}s")
            
            # 激活游戏窗口
            try:
                hwnd = win32gui.FindWindow("UnrealWindow", "王者荣耀世界")
                if hwnd:
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.3)
            except Exception:
                pass
            
            return True
        
        # 检查启动器是否还在
        launcher = find_launcher_window()
        if not launcher:
            log("[Launcher] Launcher window disappeared, game may have started")
        
        time.sleep(check_interval)
    
    log(f"[Launcher] Timeout: game window did not appear after {timeout}s")
    return False


def handle_launcher_startup(attempts=3) -> bool:
    """处理通过启动器启动游戏的完整流程
    
    流程：
    1. 检查游戏窗口是否存在
    2. 如果不存在，查找启动器
    3. 在启动器中点击"启动游戏"
    4. 等待游戏窗口出现
    
    Args:
        attempts: 最大尝试次数
    
    Returns:
        bool: True 如果游戏成功启动
    """
    log("=" * 50)
    log("[Launcher] Starting launcher-based game startup...")
    
    # 第一步：检查游戏是否已经运行
    if is_game_window_alive():
        log("[Launcher] Game window already alive, no need to use launcher")
        return True
    
    log("[Launcher] Game window not found, searching for launcher...")
    
    for attempt in range(attempts):
        log(f"[Launcher] Attempt {attempt + 1}/{attempts}")
        
        # 查找启动器窗口
        launcher_hwnd = find_launcher_window()
        
        if not launcher_hwnd:
            # 启动器还没弹出来，等一下
            if attempt == 0:
                log("[Launcher] Launcher not found yet, waiting for it to appear...")
                time.sleep(3)
                launcher_hwnd = find_launcher_window()
            
            if not launcher_hwnd:
                log("[Launcher] Launcher not found")
                # 启动器找不到，可能已经在启动游戏了，等待一下游戏窗口
                if wait_for_game_window(timeout=10):
                    return True
                time.sleep(2)
                continue
        
        # 找到了启动器，尝试点击"启动游戏"按钮
        
        # 方法1：OCR识别按钮位置并点击
        button_pos = find_start_button_by_ocr()
        if button_pos:
            x, y = button_pos
            log(f"[Launcher] Clicking '启动游戏' at ({x}, {y})")
            click_at_position(x, y)
        else:
            # 方法2：OCR失败，使用键盘导航（Enter键）
            log("[Launcher] OCR failed, trying keyboard navigation...")
            if not find_start_button_by_keyboard():
                # 方法3：最后手段 - 在启动器窗口中央区域点击
                log("[Launcher] Keyboard navigation also failed, trying center-click...")
                try:
                    rect = win32gui.GetWindowRect(launcher_hwnd)
                    left, top, right, bottom = rect
                    # 点击启动器窗口下半部分的中央（"启动游戏"按钮通常在下方）
                    cx = (left + right) // 2
                    cy = top + int((bottom - top) * 0.75)  # 窗口75%高度位置
                    log(f"[Launcher] Clicking center of launcher at ({cx}, {cy})")
                    click_at_position(cx, cy)
                except Exception as e:
                    log(f"[Launcher] Center-click failed: {e}")
        
        # 等待游戏窗口
        if wait_for_game_window(timeout=20):
            log("[Launcher] Game started successfully via launcher!")
            return True
        
        # 这一轮失败了，等一会重试
        log(f"[Launcher] Attempt {attempt + 1} failed, retrying...")
        time.sleep(3)
    
    log("[Launcher] All attempts exhausted, game did not start")
    return False
