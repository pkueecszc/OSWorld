import base64
import io
import time
from typing import Any, Dict, Tuple

import pyautogui


def _capture_screenshot_base64() -> str:
    image = pyautogui.screenshot()
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    return base64.b64encode(buffer.getvalue()).decode("utf-8")


def _build_response(success: bool, message: str, screenshot: str) -> Dict[str, Any]:
    return {
        "code": 0 if success else 1,
        "success": success,
        "message": message,
        "screenshot": screenshot,
    }


def _safe_int(value: Any) -> Tuple[bool, int]:
    try:
        return True, int(value)
    except Exception:
        return False, 0


def _maybe_move_to(data: Dict[str, Any]) -> None:
    if "coordinate_x" in data and "coordinate_y" in data:
        ok_x, x = _safe_int(data.get("coordinate_x"))
        ok_y, y = _safe_int(data.get("coordinate_y"))
        if ok_x and ok_y:
            pyautogui.moveTo(x, y)


def exec_gui_click(data: Dict[str, Any]) -> Tuple[Dict[str, Any], int]:
    ok_x, x = _safe_int(data.get("coordinate_x"))
    ok_y, y = _safe_int(data.get("coordinate_y"))
    if not ok_x or not ok_y:
        screenshot = _capture_screenshot_base64()
        return _build_response(False, "invalid coordinates", screenshot), 400
    pyautogui.click(x=x, y=y)
    screenshot = _capture_screenshot_base64()
    return _build_response(True, "ok", screenshot), 200


def exec_gui_drag(data: Dict[str, Any]) -> Tuple[Dict[str, Any], int]:
    ok_sx, sx = _safe_int(data.get("source_coordinate_x"))
    ok_sy, sy = _safe_int(data.get("source_coordinate_y"))
    ok_tx, tx = _safe_int(data.get("target_coordinate_x"))
    ok_ty, ty = _safe_int(data.get("target_coordinate_y"))
    if not (ok_sx and ok_sy and ok_tx and ok_ty):
        screenshot = _capture_screenshot_base64()
        return _build_response(False, "invalid coordinates", screenshot), 400
    pyautogui.moveTo(sx, sy)
    pyautogui.dragTo(tx, ty, duration=0.1, button="left")
    screenshot = _capture_screenshot_base64()
    return _build_response(True, "ok", screenshot), 200


def exec_gui_left_double(data: Dict[str, Any]) -> Tuple[Dict[str, Any], int]:
    ok_x, x = _safe_int(data.get("coordinate_x"))
    ok_y, y = _safe_int(data.get("coordinate_y"))
    if not ok_x or not ok_y:
        screenshot = _capture_screenshot_base64()
        return _build_response(False, "invalid coordinates", screenshot), 400
    pyautogui.doubleClick(x=x, y=y, button="left")
    screenshot = _capture_screenshot_base64()
    return _build_response(True, "ok", screenshot), 200


def exec_gui_right_single(data: Dict[str, Any]) -> Tuple[Dict[str, Any], int]:
    ok_x, x = _safe_int(data.get("coordinate_x"))
    ok_y, y = _safe_int(data.get("coordinate_y"))
    if not ok_x or not ok_y:
        screenshot = _capture_screenshot_base64()
        return _build_response(False, "invalid coordinates", screenshot), 400
    pyautogui.click(x=x, y=y, button="right")
    screenshot = _capture_screenshot_base64()
    return _build_response(True, "ok", screenshot), 200


def exec_gui_scroll(data: Dict[str, Any]) -> Tuple[Dict[str, Any], int]:
    direction = data.get("direction")
    if direction not in {"up", "down", "left", "right"}:
        screenshot = _capture_screenshot_base64()
        return _build_response(False, "invalid direction", screenshot), 400
    _maybe_move_to(data)
    amount = 120
    if direction == "up":
        pyautogui.scroll(amount)
    elif direction == "down":
        pyautogui.scroll(-amount)
    elif direction == "left":
        pyautogui.hscroll(-amount)
    else:
        pyautogui.hscroll(amount)
    screenshot = _capture_screenshot_base64()
    return _build_response(True, "ok", screenshot), 200


def exec_gui_move_to(data: Dict[str, Any]) -> Tuple[Dict[str, Any], int]:
    ok_x, x = _safe_int(data.get("coordinate_x"))
    ok_y, y = _safe_int(data.get("coordinate_y"))
    if not ok_x or not ok_y:
        screenshot = _capture_screenshot_base64()
        return _build_response(False, "invalid coordinates", screenshot), 400
    pyautogui.moveTo(x, y)
    screenshot = _capture_screenshot_base64()
    return _build_response(True, "ok", screenshot), 200


def exec_gui_mouse_down(data: Dict[str, Any]) -> Tuple[Dict[str, Any], int]:
    button = data.get("button")
    if button not in {"left", "right", "middle"}:
        screenshot = _capture_screenshot_base64()
        return _build_response(False, "invalid button", screenshot), 400
    _maybe_move_to(data)
    pyautogui.mouseDown(button=button)
    screenshot = _capture_screenshot_base64()
    return _build_response(True, "ok", screenshot), 200


def exec_gui_mouse_up(data: Dict[str, Any]) -> Tuple[Dict[str, Any], int]:
    button = data.get("button")
    if button not in {"left", "right", "middle"}:
        screenshot = _capture_screenshot_base64()
        return _build_response(False, "invalid button", screenshot), 400
    _maybe_move_to(data)
    pyautogui.mouseUp(button=button)
    screenshot = _capture_screenshot_base64()
    return _build_response(True, "ok", screenshot), 200


def exec_gui_type(data: Dict[str, Any]) -> Tuple[Dict[str, Any], int]:
    text = data.get("text")
    if text is None:
        screenshot = _capture_screenshot_base64()
        return _build_response(False, "text required", screenshot), 400
    pyautogui.write(str(text))
    screenshot = _capture_screenshot_base64()
    return _build_response(True, "ok", screenshot), 200


def exec_gui_hotkey(data: Dict[str, Any]) -> Tuple[Dict[str, Any], int]:
    keys = data.get("keys")
    if not isinstance(keys, list) or not keys:
        screenshot = _capture_screenshot_base64()
        return _build_response(False, "keys required", screenshot), 400
    pyautogui.hotkey(*[str(k) for k in keys])
    screenshot = _capture_screenshot_base64()
    return _build_response(True, "ok", screenshot), 200


def exec_gui_press(data: Dict[str, Any]) -> Tuple[Dict[str, Any], int]:
    key = data.get("key")
    if not key:
        screenshot = _capture_screenshot_base64()
        return _build_response(False, "key required", screenshot), 400
    pyautogui.press(str(key))
    screenshot = _capture_screenshot_base64()
    return _build_response(True, "ok", screenshot), 200


def exec_gui_release(data: Dict[str, Any]) -> Tuple[Dict[str, Any], int]:
    key = data.get("key")
    if not key:
        screenshot = _capture_screenshot_base64()
        return _build_response(False, "key required", screenshot), 400
    pyautogui.keyUp(str(key))
    screenshot = _capture_screenshot_base64()
    return _build_response(True, "ok", screenshot), 200


def exec_gui_wait(data: Dict[str, Any]) -> Tuple[Dict[str, Any], int]:
    ok, duration = _safe_int(data.get("duration"))
    if not ok or duration < 0:
        screenshot = _capture_screenshot_base64()
        return _build_response(False, "invalid duration", screenshot), 400
    time.sleep(duration)
    screenshot = _capture_screenshot_base64()
    return _build_response(True, "ok", screenshot), 200
