import ctypes
import json
import logging
import os
import threading
import time
from dataclasses import dataclass, asdict

import pystray
from PIL import Image, ImageDraw
import win32api  # pyright: ignore[reportMissingModuleSource]
from win32com.client import Dispatch


APP_NAME = "AwakeKeeper"
CONFIG_DIR = os.path.join(os.environ.get("APPDATA", "."), APP_NAME)
CONFIG_PATH = os.path.join(CONFIG_DIR, "config.json")
LOG_PATH = os.path.join(CONFIG_DIR, "awake_keeper.log")

ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002

VALID_METHODS = {"preventsleep", "scrolllock", "mousejiggle"}
VALID_PROFILES = {"Dashboard Mode", "Presentation Mode", "Custom"}
DEFAULT_METHOD = "preventsleep"
DEFAULT_PROFILE = "Dashboard Mode"


os.makedirs(CONFIG_DIR, exist_ok=True)

logging.basicConfig(
    filename=LOG_PATH,
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)
LOGGER = logging.getLogger(APP_NAME)


@dataclass
class AppConfig:
    idle_threshold_seconds: int = 30
    check_interval_seconds: int = 5
    method: str = DEFAULT_METHOD
    profile: str = DEFAULT_PROFILE
    start_enabled: bool = True


class AwakeKeeper:
    def __init__(self):
        self.config = self.load_config()
        self.enabled = self.config.start_enabled
        self.running = True
        self.last_keep_awake = None
        self.last_error = "None"
        self.sleep_override_active = False
        self.lock = threading.Lock()
        self.menu_state = None

        try:
            self.shell = Dispatch("WScript.Shell")
        except Exception:
            LOGGER.exception("Failed to initialize WScript.Shell; ScrollLock will fall back to mouse jiggle.")
            self.shell = None

        self.icon = pystray.Icon(
            APP_NAME,
            self.create_icon_image(),
            APP_NAME,
            menu=self.build_menu(),
        )

        self.worker_thread = threading.Thread(target=self.worker_loop, daemon=True)

    # ----------------------------
    # Config
    # ----------------------------
    def normalize_config(self, data: dict | None) -> AppConfig:
        if not isinstance(data, dict):
            data = {}

        idle_threshold = data.get("idle_threshold_seconds", AppConfig.idle_threshold_seconds)
        check_interval = data.get("check_interval_seconds", AppConfig.check_interval_seconds)
        method = str(data.get("method", DEFAULT_METHOD)).lower()
        profile = str(data.get("profile", DEFAULT_PROFILE))
        start_enabled = bool(data.get("start_enabled", True))

        try:
            idle_threshold = int(idle_threshold)
        except (TypeError, ValueError):
            idle_threshold = AppConfig.idle_threshold_seconds

        try:
            check_interval = int(check_interval)
        except (TypeError, ValueError):
            check_interval = AppConfig.check_interval_seconds

        idle_threshold = min(max(idle_threshold, 1), 3600)
        check_interval = min(max(check_interval, 1), 3600)

        if method not in VALID_METHODS:
            LOGGER.warning("Unknown method '%s' in config; resetting to %s.", method, DEFAULT_METHOD)
            method = DEFAULT_METHOD

        if profile not in VALID_PROFILES:
            profile = "Custom"

        return AppConfig(
            idle_threshold_seconds=idle_threshold,
            check_interval_seconds=check_interval,
            method=method,
            profile=profile,
            start_enabled=start_enabled,
        )

    def backup_invalid_config(self):
        if not os.path.exists(CONFIG_PATH):
            return

        timestamp = time.strftime("%Y%m%d-%H%M%S")
        backup_path = os.path.join(CONFIG_DIR, f"config.invalid.{timestamp}.json")
        try:
            os.replace(CONFIG_PATH, backup_path)
            LOGGER.warning("Moved invalid config to %s.", backup_path)
        except Exception:
            LOGGER.exception("Failed to back up invalid config.")

    def load_config(self) -> AppConfig:
        if not os.path.exists(CONFIG_PATH):
            return AppConfig()

        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            LOGGER.exception("Failed to load config from %s.", CONFIG_PATH)
            self.backup_invalid_config()
            return AppConfig()

        config = self.normalize_config(data)
        if asdict(config) != data:
            LOGGER.info("Normalized config values from %s.", CONFIG_PATH)
            self.persist_config(config)
        return config

    def persist_config(self, config: AppConfig):
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(asdict(config), f, indent=2)

    def save_config(self):
        try:
            self.persist_config(self.config)
        except Exception:
            LOGGER.exception("Failed to save config to %s.", CONFIG_PATH)
            self.last_error = "Config save failed"

    # ----------------------------
    # Idle detection
    # ----------------------------
    @staticmethod
    def get_idle_time_seconds() -> int:
        class LASTINPUTINFO(ctypes.Structure):
            _fields_ = [
                ("cbSize", ctypes.c_uint),
                ("dwTime", ctypes.c_uint),
            ]

        lii = LASTINPUTINFO()
        lii.cbSize = ctypes.sizeof(LASTINPUTINFO)

        if not ctypes.windll.user32.GetLastInputInfo(ctypes.byref(lii)):
            return 0

        tick_count = ctypes.windll.kernel32.GetTickCount()
        idle_ms = tick_count - lii.dwTime
        return max(0, idle_ms // 1000)

    # ----------------------------
    # Keep-awake methods
    # ----------------------------
    def notify(self, message: str):
        try:
            self.icon.notify(message, APP_NAME)
        except Exception:
            LOGGER.info("Notification: %s", message)

    def set_last_error(self, message: str):
        self.last_error = message
        LOGGER.error(message)

    def prevent_sleep_windows(self):
        result = ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
        )
        if result == 0:
            raise OSError("SetThreadExecutionState failed while enabling preventsleep.")

        self.sleep_override_active = True
        self.last_keep_awake = time.strftime("%H:%M:%S")
        self.last_error = "None"

    def clear_sleep_override(self):
        result = ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
        if result == 0:
            raise OSError("SetThreadExecutionState failed while clearing preventsleep.")

        self.sleep_override_active = False

    def scroll_lock_pulse(self):
        if self.shell is None:
            LOGGER.warning("ScrollLock requested without WScript.Shell; falling back to mouse jiggle.")
            self.notify("ScrollLock unavailable, using MouseJiggle.")
            self.mouse_jiggle()
            return

        try:
            self.shell.SendKeys("{SCROLLLOCK}")
            time.sleep(0.1)
            self.shell.SendKeys("{SCROLLLOCK}")
            self.last_keep_awake = time.strftime("%H:%M:%S")
            self.last_error = "None"
        except Exception:
            LOGGER.exception("ScrollLock pulse failed; falling back to mouse jiggle.")
            self.notify("ScrollLock failed, using MouseJiggle.")
            self.mouse_jiggle()

    def mouse_jiggle(self):
        x, y = win32api.GetCursorPos()
        win32api.SetCursorPos((x + 1, y))
        time.sleep(0.1)
        win32api.SetCursorPos((x, y))
        self.last_keep_awake = time.strftime("%H:%M:%S")
        self.last_error = "None"

    def keep_awake(self):
        method = self.config.method.lower()

        if method == "preventsleep":
            self.prevent_sleep_windows()
        elif method == "scrolllock":
            self.scroll_lock_pulse()
        elif method == "mousejiggle":
            self.mouse_jiggle()
        else:
            raise ValueError(f"Unsupported keep-awake method: {method}")

    # ----------------------------
    # Profiles
    # ----------------------------
    def set_dashboard_mode(self):
        with self.lock:
            self.config.idle_threshold_seconds = 30
            self.config.check_interval_seconds = 5
            self.config.method = "preventsleep"
            self.config.profile = "Dashboard Mode"
            self.save_config()
        self.refresh_menu(force=True)

    def set_presentation_mode(self):
        with self.lock:
            if self.sleep_override_active:
                self.clear_sleep_override()
            self.config.idle_threshold_seconds = 60
            self.config.check_interval_seconds = 10
            self.config.method = "mousejiggle"
            self.config.profile = "Presentation Mode"
            self.save_config()
        self.refresh_menu(force=True)

    # ----------------------------
    # UI helpers
    # ----------------------------
    def create_icon_image(self):
        width = 64
        height = 64
        image = Image.new("RGBA", (width, height), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)

        draw.rounded_rectangle((8, 8, 56, 56), radius=12, fill=(25, 95, 170, 255))
        draw.rectangle((20, 18, 44, 46), fill=(255, 255, 255, 255))
        draw.rectangle((24, 22, 40, 42), fill=(25, 95, 170, 255))
        return image

    def status_lines(self):
        idle = self.get_idle_time_seconds()
        state = "ON" if self.enabled else "OFF"
        last = self.last_keep_awake or "Never"
        override = "Active" if self.sleep_override_active else "Inactive"

        return {
            "state": f"Keeper: {state}",
            "idle": f"Idle: {idle}s",
            "threshold": f"Threshold: {self.config.idle_threshold_seconds}s",
            "interval": f"Check: {self.config.check_interval_seconds}s",
            "method": f"Method: {self.config.method}",
            "profile": f"Profile: {self.config.profile}",
            "last": f"Last action: {last}",
            "override": f"Sleep override: {override}",
            "error": f"Last error: {self.last_error}",
        }

    def build_menu(self):
        status = self.status_lines()

        return pystray.Menu(
            pystray.MenuItem(status["state"], None, enabled=False),
            pystray.MenuItem(status["idle"], None, enabled=False),
            pystray.MenuItem(status["threshold"], None, enabled=False),
            pystray.MenuItem(status["interval"], None, enabled=False),
            pystray.MenuItem(status["method"], None, enabled=False),
            pystray.MenuItem(status["profile"], None, enabled=False),
            pystray.MenuItem(status["override"], None, enabled=False),
            pystray.MenuItem(status["last"], None, enabled=False),
            pystray.MenuItem(status["error"], None, enabled=False),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(
                "Stop" if self.enabled else "Start",
                self.toggle_enabled,
            ),
            pystray.MenuItem(
                "Run Now",
                self.run_now,
                enabled=lambda item: self.enabled,
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(
                "Dashboard Mode",
                self.on_dashboard_mode,
                checked=lambda item: self.config.profile == "Dashboard Mode",
            ),
            pystray.MenuItem(
                "Presentation Mode",
                self.on_presentation_mode,
                checked=lambda item: self.config.profile == "Presentation Mode",
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(
                "Use PreventSleep",
                self.use_preventsleep,
                checked=lambda item: self.config.method == "preventsleep",
            ),
            pystray.MenuItem(
                "Use ScrollLock",
                self.use_scrolllock,
                checked=lambda item: self.config.method == "scrolllock",
            ),
            pystray.MenuItem(
                "Use MouseJiggle",
                self.use_mousejiggle,
                checked=lambda item: self.config.method == "mousejiggle",
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Exit", self.exit_app),
        )

    def menu_snapshot(self):
        status = self.status_lines()
        return tuple(status.values()) + (self.enabled,)

    def refresh_menu(self, force: bool = False):
        snapshot = self.menu_snapshot()
        if not force and snapshot == self.menu_state:
            return

        self.menu_state = snapshot
        self.icon.menu = self.build_menu()
        self.icon.title = f"{APP_NAME} [{'ON' if self.enabled else 'OFF'}]"
        try:
            self.icon.update_menu()
        except Exception:
            # Some backends do not need or implement explicit menu refresh.
            pass

    # ----------------------------
    # Menu actions
    # ----------------------------
    def set_custom_profile(self):
        if self.config.profile not in {"Dashboard Mode", "Presentation Mode"}:
            return
        self.config.profile = "Custom"

    def toggle_enabled(self, icon=None, item=None):
        with self.lock:
            self.enabled = not self.enabled
            if not self.enabled and self.sleep_override_active:
                self.clear_sleep_override()
            self.config.start_enabled = self.enabled
            self.save_config()

        self.notify(f"Keeper {'started' if self.enabled else 'stopped'}.")
        self.refresh_menu(force=True)

    def run_now(self, icon=None, item=None):
        if not self.enabled:
            return

        try:
            self.keep_awake()
        except Exception:
            LOGGER.exception("Run Now failed.")
            self.last_error = "Run Now failed"
        self.refresh_menu(force=True)

    def on_dashboard_mode(self, icon=None, item=None):
        self.set_dashboard_mode()

    def on_presentation_mode(self, icon=None, item=None):
        self.set_presentation_mode()

    def use_preventsleep(self, icon=None, item=None):
        with self.lock:
            self.config.method = "preventsleep"
            self.set_custom_profile()
            self.save_config()
        self.notify("Method set to PreventSleep.")
        self.refresh_menu(force=True)

    def use_scrolllock(self, icon=None, item=None):
        with self.lock:
            if self.sleep_override_active:
                self.clear_sleep_override()
            self.config.method = "scrolllock"
            self.set_custom_profile()
            self.save_config()
        self.notify("Method set to ScrollLock.")
        self.refresh_menu(force=True)

    def use_mousejiggle(self, icon=None, item=None):
        with self.lock:
            if self.sleep_override_active:
                self.clear_sleep_override()
            self.config.method = "mousejiggle"
            self.set_custom_profile()
            self.save_config()
        self.notify("Method set to MouseJiggle.")
        self.refresh_menu(force=True)

    def exit_app(self, icon=None, item=None):
        self.running = False
        try:
            if self.sleep_override_active:
                self.clear_sleep_override()
        except Exception:
            LOGGER.exception("Failed to clear sleep override on exit.")

        self.save_config()
        self.icon.stop()

    # ----------------------------
    # Worker loop
    # ----------------------------
    def worker_loop(self):
        while self.running:
            try:
                if self.enabled:
                    idle = self.get_idle_time_seconds()

                    if idle >= self.config.idle_threshold_seconds:
                        self.keep_awake()
                    elif self.config.method == "preventsleep" and self.sleep_override_active:
                        self.clear_sleep_override()
                elif self.sleep_override_active:
                    self.clear_sleep_override()
            except Exception:
                self.last_error = "Worker loop failed"
                LOGGER.exception("Worker loop iteration failed.")

            self.refresh_menu()
            time.sleep(max(1, self.config.check_interval_seconds))

    def run(self):
        LOGGER.info("Starting %s.", APP_NAME)
        self.worker_thread.start()
        self.icon.run()


if __name__ == "__main__":
    app = AwakeKeeper()
    app.run()
