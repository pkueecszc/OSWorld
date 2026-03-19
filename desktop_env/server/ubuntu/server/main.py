import atexit
import concurrent.futures
import ctypes
import os
import platform
import queue
import re
import shlex
import shutil
import signal
import subprocess
import time
import uuid
from pathlib import Path
from typing import Any, Optional, Sequence
from typing import List, Dict, Tuple
from requests.utils import should_bypass_proxies

import Xlib
import lxml.etree
import pyautogui
import requests
from PIL import Image
from Xlib import display, X
from flask import Flask, request, jsonify, send_file, abort  # , send_from_directory
from jupyter_client.manager import KernelManager
from lxml.etree import _Element

platform_name: str = platform.system()

import pyatspi
from pyatspi import Accessible, StateType, STATE_SHOWING
from pyatspi import Action as ATAction
from pyatspi import Component  # , Document
from pyatspi import Text as ATText
from pyatspi import Value as ATValue

BaseWrapper = Any

from pyxcursor import Xcursor

# todo: need to reformat and organize this whole file

app = Flask(__name__)

pyautogui.PAUSE = 0
pyautogui.DARWIN_CATCH_UP_TIME = 0

runtime_path = os.path.dirname(os.path.abspath(__file__))
logger = app.logger
recording_process = None  # fixme: this is a temporary solution for recording, need to be changed to support multiple-process
recording_path = "/tmp/recording.mp4"

proxies = {
    "http": "http://seed_gui_osworld_proxy:2Gj6QEgYtInSL5Xx@id3473.http-sg-idc-idc-sg-flow.forward-proxy.byted.org:8080",
    "https": "http://seed_gui_osworld_proxy:2Gj6QEgYtInSL5Xx@id3473.http-sg-idc-idc-sg-flow.forward-proxy.byted.org:8080",
}

os.environ["NO_PROXY"] = "localhost,127.0.0.0/8,::1,.byted.org,.bytedance.net"
os.environ["no_proxy"] = "localhost,127.0.0.0/8,::1,.byted.org,.bytedance.net"

TMUX_SESSION = "vm_control_session"
# Global state to hold the result of the last executed command.
last_command_info = {"exit_code": None, "output": "", "markers": None}

def _do_prepare(cfg: Dict[str, Any]):
    result = ""
    for i in range(3):
        config_type: str = cfg["type"]
        parameters: Dict[str, Any] = cfg["parameters"]
        try:
            # Assumes all the setup the functions should follow this name
            # protocol
            if config_type == "launch":
                _do_launch(
                    command=parameters["command"], shell=parameters.get("shell", False)
                )
            elif config_type in ["execute", "command"]:
                _do_execute(parameters["command"], shell=parameters.get("shell", False))
            elif config_type == "download":
                if "files" in parameters:
                    _do_batch_download_file(parameters["files"])
                else:
                    _do_download_file(parameters["url"], parameters["path"])
            elif config_type == "activate_window":
                _do_activate_window(
                    parameters["window_name"],
                    parameters.get("strict", False),
                    parameters.get("by_class", False),
                )
            elif config_type == "chrome_open_tabs":
                _do_chrome_open_tabs(parameters.get("urls_to_open", []))
            elif config_type == "chrome_close_tabs":
                _do_chrome_close_tabs(parameters.get("urls_to_close", []))
            elif config_type == "sleep":
                time.sleep(parameters["seconds"])
            elif config_type == "open":
                _do_open_file(parameters["path"])
            elif config_type == "update_browse_history":
                _do_update_browse_history(parameters.get("history", []))
            else:
                raise NotImplementedError(
                    f"Setup type {config_type} is not implemented"
                )

            return "Success SETUP: %s(%s)" % (config_type, str(parameters))
        except Exception as e:
            result = "Failed to setup {:}. Caused by {:}".format(config_type, e)

    return result


@app.route("/setup/prepare_all", methods=["POST"])
def prepare_all():
    data = request.json
    config = data.get("config", [])
    results = []
    for cfg in config:
        results.append(_do_prepare(cfg))

    return jsonify({"status": "ok", "results": results})


def _do_execute(command: List[str] | str, shell: bool = False):
    if isinstance(command, str) and not shell:
        command = shlex.split(command)

    # Expand user directory
    for i, arg in enumerate(command):
        if arg.startswith("~/"):
            command[i] = os.path.expanduser(arg)

    # Execute the command with safety checks.
    result = subprocess.run(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        shell=shell,
        text=True,
        timeout=120,
        check=True
    )
    return result


@app.route("/setup/execute", methods=["POST"])
@app.route("/execute", methods=["POST"])
def execute_command():
    data = request.json
    # The 'command' key in the JSON request should contain the command to be executed.
    shell = data.get("shell", False)
    command = data.get("command", "" if shell else [])

    try:
        result = _do_execute(command, shell)
        return jsonify(
            {
                "status": "success",
                "output": result.stdout,
                "error": result.stderr,
                "returncode": result.returncode,
            }
        )
    except Exception as e:
        if isinstance(e, subprocess.CalledProcessError):
            error_message = (
                f"Command {e.cmd} failed with return code {e.returncode}. "
                f"Output: {e.stdout.strip() if e.stdout else 'None'}. "
                f"Error: {e.stderr.strip() if e.stderr else 'None'}."
            )
        elif isinstance(e, subprocess.TimeoutExpired):
            error_message = f"Command timed out after {e.timeout} seconds."
        else:
            error_message = f"{str(e)}"
        return jsonify({"status": "error", "message": error_message}), 500


def _get_machine_architecture() -> str:
    """Get the machine architecture, e.g., x86_64, arm64, aarch64, i386, etc."""
    architecture = platform.machine().lower()
    if architecture in [
        "amd32",
        "amd64",
        "x86",
        "x86_64",
        "x86-64",
        "x64",
        "i386",
        "i686",
    ]:
        return "amd"
    elif architecture in ["arm64", "aarch64", "aarch32"]:
        return "arm"
    else:
        return "unknown"


def _do_launch(command: List[str] | str, shell: bool = False):
    if isinstance(command, str) and not shell:
        command = shlex.split(command)

    # Expand user directory
    for i, arg in enumerate(command):
        if arg.startswith("~/"):
            command[i] = os.path.expanduser(arg)

    if "google-chrome" in command and _get_machine_architecture() == "arm":
        index = command.index("google-chrome")
        command[index] = (
            "chromium-browser"  # arm64 chrome is not available yet, can only use chromium
        )
        if "--proxy-server=http://sys-proxy-rd-relay.byted.org:8118" in command:
            index = command.index(
                "--proxy-server=http://sys-proxy-rd-relay.byted.org:8118"
            )
            global_proxy = proxies["http"]
            command[index] = f"--proxy-server={global_proxy}"

    subprocess.Popen(command, shell=shell)


@app.route("/setup/launch", methods=["POST"])
def launch_app():
    data = request.json
    shell = data.get("shell", False)
    command: List[str] = data.get("command", "" if shell else [])

    try:
        _do_launch(command, shell)
        return "{:} launched successfully".format(
            command if shell else " ".join(command)
        )
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/screenshot", methods=["GET"])
def capture_screen_with_cursor():
    # fixme: when running on virtual machines, the cursor is not captured, don't know why

    file_path = os.path.join(os.path.dirname(__file__), "screenshots", "screenshot.png")
    user_platform = platform.system()

    # Ensure the screenshots directory exists
    os.makedirs(os.path.dirname(file_path), exist_ok=True)

    # fixme: This is a temporary fix for the cursor not being captured on Windows and Linux
    if user_platform == "Windows":

        def _download_image(url, path):
            response = requests.get(url)
            with open(path, "wb") as file:
                file.write(response.content)

        cursor_path = os.path.join("screenshots", "cursor.png")
        if not os.path.exists(cursor_path):
            cursor_url = "https://vip.helloimg.com/images/2023/12/02/oQPzmt.png"
            _download_image(cursor_url, cursor_path)
        screenshot = pyautogui.screenshot()
        cursor_x, cursor_y = pyautogui.position()
        cursor = Image.open(cursor_path)
        # make the cursor smaller
        cursor = cursor.resize((int(cursor.width / 1.5), int(cursor.height / 1.5)))
        screenshot.paste(cursor, (cursor_x, cursor_y), cursor)
        screenshot.save(file_path)
    elif user_platform == "Linux":
        cursor_obj = None
        try:
            cursor_obj = Xcursor()
            imgarray = cursor_obj.getCursorImageArrayFast()
            cursor_img = Image.fromarray(imgarray)
            screenshot = pyautogui.screenshot()
            cursor_x, cursor_y = pyautogui.position()
            screenshot.paste(cursor_img, (cursor_x, cursor_y), cursor_img)
            screenshot.save(file_path)
        finally:
            if cursor_obj:
                cursor_obj.close()
    elif user_platform == "Darwin":  # (Mac OS)
        # Use the screencapture utility to capture the screen with the cursor
        subprocess.run(["screencapture", "-C", file_path])
    else:
        logger.warning(
            f"The platform you're using ({user_platform}) is not currently supported"
        )

    return send_file(file_path, mimetype="image/png")


def _has_active_terminal(desktop: Accessible) -> bool:
    """A quick check whether the terminal window is open and active."""
    for app in desktop:
        if app.getRoleName() == "application" and app.name == "gnome-terminal-server":
            for frame in app:
                if frame.getRoleName() == "frame" and frame.getState().contains(
                    pyatspi.STATE_ACTIVE
                ):
                    return True
    return False


@app.route("/terminal", methods=["GET"])
def get_terminal_output():
    user_platform = platform.system()
    output: Optional[str] = None
    try:
        if user_platform == "Linux":
            desktop: Accessible = pyatspi.Registry.getDesktop(0)
            if _has_active_terminal(desktop):
                desktop_xml: _Element = _create_atspi_node(desktop)
                # 1. the terminal window (frame of application is st:active) is open and active
                # 2. the terminal tab (terminal status is st:focused) is focused
                xpath = '//application[@name="gnome-terminal-server"]/frame[@st:active="true"]//terminal[@st:focused="true"]'
                terminals: List[_Element] = desktop_xml.xpath(
                    xpath, namespaces=_accessibility_ns_map_ubuntu
                )
                output = terminals[0].text.rstrip() if len(terminals) == 1 else None
        else:  # windows and macos platform is not implemented currently
            # raise NotImplementedError
            return (
                "Currently not implemented for platform {:}.".format(
                    platform.platform()
                ),
                500,
            )
        return jsonify({"output": output, "status": "success"})
    except Exception as e:
        logger.error("Failed to get terminal output. Error: %s", e)
        return jsonify({"status": "error", "message": str(e)}), 500


_accessibility_ns_map = {
    "ubuntu": {
        "st": "https://accessibility.ubuntu.example.org/ns/state",
        "attr": "https://accessibility.ubuntu.example.org/ns/attributes",
        "cp": "https://accessibility.ubuntu.example.org/ns/component",
        "doc": "https://accessibility.ubuntu.example.org/ns/document",
        "docattr": "https://accessibility.ubuntu.example.org/ns/document/attributes",
        "txt": "https://accessibility.ubuntu.example.org/ns/text",
        "val": "https://accessibility.ubuntu.example.org/ns/value",
        "act": "https://accessibility.ubuntu.example.org/ns/action",
    },
    "windows": {
        "st": "https://accessibility.windows.example.org/ns/state",
        "attr": "https://accessibility.windows.example.org/ns/attributes",
        "cp": "https://accessibility.windows.example.org/ns/component",
        "doc": "https://accessibility.windows.example.org/ns/document",
        "docattr": "https://accessibility.windows.example.org/ns/document/attributes",
        "txt": "https://accessibility.windows.example.org/ns/text",
        "val": "https://accessibility.windows.example.org/ns/value",
        "act": "https://accessibility.windows.example.org/ns/action",
        "class": "https://accessibility.windows.example.org/ns/class",
    },
    "macos": {
        "st": "https://accessibility.macos.example.org/ns/state",
        "attr": "https://accessibility.macos.example.org/ns/attributes",
        "cp": "https://accessibility.macos.example.org/ns/component",
        "doc": "https://accessibility.macos.example.org/ns/document",
        "txt": "https://accessibility.macos.example.org/ns/text",
        "val": "https://accessibility.macos.example.org/ns/value",
        "act": "https://accessibility.macos.example.org/ns/action",
        "role": "https://accessibility.macos.example.org/ns/role",
    },
}

_accessibility_ns_map_ubuntu = _accessibility_ns_map["ubuntu"]
_accessibility_ns_map_windows = _accessibility_ns_map["windows"]
_accessibility_ns_map_macos = _accessibility_ns_map["macos"]

# A11y tree getter for Ubuntu
libreoffice_version_tuple: Optional[Tuple[int, ...]] = None
MAX_DEPTH = 50
MAX_WIDTH = 1024
MAX_CALLS = 5000


def _get_libreoffice_version() -> Tuple[int, ...]:
    """Function to get the LibreOffice version as a tuple of integers."""
    result = subprocess.run(
        "libreoffice --version", shell=True, text=True, stdout=subprocess.PIPE
    )
    version_str = result.stdout.split()[
        1
    ]  # Assuming version is the second word in the command output
    return tuple(map(int, version_str.split(".")))


def _create_atspi_node(
    node: Accessible, depth: int = 0, flag: Optional[str] = None
) -> _Element:
    node_name = node.name
    attribute_dict: Dict[str, Any] = {"name": node_name}

    #  States
    states: List[StateType] = node.getState().get_states()
    for st in states:
        state_name: str = StateType._enum_lookup[st]
        state_name: str = state_name.split("_", maxsplit=1)[1].lower()
        if len(state_name) == 0:
            continue
        attribute_dict[
            "{{{:}}}{:}".format(_accessibility_ns_map_ubuntu["st"], state_name)
        ] = "true"

    #  Attributes
    attributes: Dict[str, str] = node.get_attributes()
    for attribute_name, attribute_value in attributes.items():
        if len(attribute_name) == 0:
            continue
        attribute_dict[
            "{{{:}}}{:}".format(_accessibility_ns_map_ubuntu["attr"], attribute_name)
        ] = attribute_value

    #  Component
    if (
        attribute_dict.get(
            "{{{:}}}visible".format(_accessibility_ns_map_ubuntu["st"]), "false"
        )
        == "true"
        and attribute_dict.get(
            "{{{:}}}showing".format(_accessibility_ns_map_ubuntu["st"]), "false"
        )
        == "true"
    ):
        try:
            component: Component = node.queryComponent()
        except NotImplementedError:
            pass
        else:
            bbox: Sequence[int] = component.getExtents(pyatspi.XY_SCREEN)
            attribute_dict[
                "{{{:}}}screencoord".format(_accessibility_ns_map_ubuntu["cp"])
            ] = str(tuple(bbox[0:2]))
            attribute_dict["{{{:}}}size".format(_accessibility_ns_map_ubuntu["cp"])] = (
                str(tuple(bbox[2:]))
            )

    text = ""
    #  Text
    try:
        text_obj: ATText = node.queryText()
        # only text shown on current screen is available
        # attribute_dict["txt:text"] = text_obj.getText(0, text_obj.characterCount)
        text: str = text_obj.getText(0, text_obj.characterCount)
        # if flag=="thunderbird":
        # appeared in thunderbird (uFFFC) (not only in thunderbird), "Object
        # Replacement Character" in Unicode, "used as placeholder in text for
        # an otherwise unspecified object; uFFFD is another "Replacement
        # Character", just in case
        text = text.replace("\ufffc", "").replace("\ufffd", "")
    except NotImplementedError:
        pass

    #  Image, Selection, Value, Action
    try:
        node.queryImage()
        attribute_dict["image"] = "true"
    except NotImplementedError:
        pass

    try:
        node.querySelection()
        attribute_dict["selection"] = "true"
    except NotImplementedError:
        pass

    try:
        value: ATValue = node.queryValue()
        value_key = f"{{{_accessibility_ns_map_ubuntu['val']}}}"

        for attr_name, attr_func in [
            ("value", lambda: value.currentValue),
            ("min", lambda: value.minimumValue),
            ("max", lambda: value.maximumValue),
            ("step", lambda: value.minimumIncrement),
        ]:
            try:
                attribute_dict[f"{value_key}{attr_name}"] = str(attr_func())
            except:
                pass
    except NotImplementedError:
        pass

    try:
        action: ATAction = node.queryAction()
        for i in range(action.nActions):
            action_name: str = action.getName(i).replace(" ", "-")
            attribute_dict[
                "{{{:}}}{:}_desc".format(
                    _accessibility_ns_map_ubuntu["act"], action_name
                )
            ] = action.getDescription(i)
            attribute_dict[
                "{{{:}}}{:}_kb".format(_accessibility_ns_map_ubuntu["act"], action_name)
            ] = action.getKeyBinding(i)
    except NotImplementedError:
        pass

    # Add from here if we need more attributes in the future...

    raw_role_name: str = node.getRoleName().strip()
    node_role_name = (raw_role_name or "unknown").replace(" ", "-")

    if not flag:
        if raw_role_name == "document spreadsheet":
            flag = "calc"
        if raw_role_name == "application" and node.name == "Thunderbird":
            flag = "thunderbird"

    xml_node = lxml.etree.Element(
        node_role_name, attrib=attribute_dict, nsmap=_accessibility_ns_map_ubuntu
    )

    if len(text) > 0:
        xml_node.text = text

    if depth == MAX_DEPTH:
        logger.warning("Max depth reached")
        return xml_node

    if flag == "calc" and node_role_name == "table":
        # Maximum column: 1024 if ver<=7.3 else 16384
        # Maximum row: 104 8576
        # Maximun sheet: 1 0000

        global libreoffice_version_tuple
        MAXIMUN_COLUMN = 1024 if libreoffice_version_tuple < (7, 4) else 16384
        MAX_ROW = 104_8576

        index_base = 0
        first_showing = False
        column_base = None
        for r in range(MAX_ROW):
            for clm in range(column_base or 0, MAXIMUN_COLUMN):
                child_node: Accessible = node[index_base + clm]
                showing: bool = child_node.getState().contains(STATE_SHOWING)
                if showing:
                    child_node: _Element = _create_atspi_node(
                        child_node, depth + 1, flag
                    )
                    if not first_showing:
                        column_base = clm
                        first_showing = True
                    xml_node.append(child_node)
                elif first_showing and column_base is not None or clm >= 500:
                    break
            if first_showing and clm == column_base or not first_showing and r >= 500:
                break
            index_base += MAXIMUN_COLUMN
        return xml_node
    else:
        try:
            for i, ch in enumerate(node):
                if i == MAX_WIDTH:
                    logger.warning("Max width reached")
                    break
                xml_node.append(_create_atspi_node(ch, depth + 1, flag))
        except:
            logger.warning(
                "Error occurred during children traversing. Has Ignored. Node: %s",
                lxml.etree.tostring(xml_node, encoding="unicode"),
            )
        return xml_node


# A11y tree getter for Windows
def _create_pywinauto_node(
    node, depth: int = 0, flag: Optional[str] = None
) -> _Element:
    attribute_dict: Dict[str, Any] = {"name": node.element_info.name}

    #  States
    for attr_name, attr_func in [
        ("enabled", lambda: node.is_enabled()),
        ("visible", lambda: node.is_visible()),
        ("active", lambda: node.is_active()),
        ("minimized", lambda: node.is_minimized()),
        ("maximized", lambda: node.is_maximized()),
        ("normal", lambda: node.is_normal()),
        ("unicode", lambda: node.is_unicode()),
        ("collapsed", lambda: node.is_collapsed()),
        ("checkable", lambda: node.is_checkable()),
        ("checked", lambda: node.is_checked()),
        ("focused", lambda: node.is_focused()),
        ("keyboard_focused", lambda: node.is_keyboard_focused()),
        ("selected", lambda: node.is_selected()),
        ("selection_required", lambda: node.is_selection_required()),
        ("pressable", lambda: node.is_pressable()),
        ("pressed", lambda: node.is_pressed()),
        ("expanded", lambda: node.is_expanded()),
        ("editable", lambda: node.is_editable()),
    ]:
        try:
            attribute_dict[f"{{{_accessibility_ns_map_windows['st']}}}{attr_name}"] = (
                str(attr_func()).lower()
            )
        except:
            pass

    #  Component
    try:
        rectangle = node.rectangle()
        attribute_dict[
            "{{{:}}}screencoord".format(_accessibility_ns_map_windows["cp"])
        ] = "({:d}, {:d})".format(rectangle.left, rectangle.top)
        attribute_dict["{{{:}}}size".format(_accessibility_ns_map_windows["cp"])] = (
            "({:d}, {:d})".format(rectangle.width(), rectangle.height())
        )

    except Exception as e:
        logger.error("Error accessing rectangle: ", e)

    #  Text
    text: str = node.window_text()
    if text == attribute_dict["name"]:
        text = ""

    #  Selection
    if hasattr(node, "select"):
        attribute_dict["selection"] = "true"

    # Value
    for attr_name, attr_funcs in [
        ("step", [lambda: node.get_step()]),
        (
            "value",
            [
                lambda: node.value(),
                lambda: node.get_value(),
                lambda: node.get_position(),
            ],
        ),
        ("min", [lambda: node.min_value(), lambda: node.get_range_min()]),
        ("max", [lambda: node.max_value(), lambda: node.get_range_max()]),
    ]:
        for attr_func in attr_funcs:
            if hasattr(node, attr_func.__name__):
                try:
                    attribute_dict[
                        f"{{{_accessibility_ns_map_windows['val']}}}{attr_name}"
                    ] = str(attr_func())
                    break  # exit once the attribute is set successfully
                except:
                    pass

    attribute_dict["{{{:}}}class".format(_accessibility_ns_map_windows["class"])] = str(
        type(node)
    )

    node_role_name: str = node.class_name().lower().replace(" ", "-")
    node_role_name = "".join(
        map(
            lambda _ch: (
                _ch if _ch.isidentifier() or _ch in {"-"} or _ch.isalnum() else "-"
            ),
            node_role_name,
        )
    )

    if node_role_name.strip() == "":
        node_role_name = "unknown"
    if not node_role_name[0].isalpha():
        node_role_name = "tag" + node_role_name

    xml_node = lxml.etree.Element(
        node_role_name, attrib=attribute_dict, nsmap=_accessibility_ns_map_windows
    )

    if text is not None and len(text) > 0 and text != attribute_dict["name"]:
        xml_node.text = text

    if depth == MAX_DEPTH:
        logger.warning("Max depth reached")
        return xml_node

    for i, ch in enumerate(node.children()):
        if i >= MAX_WIDTH:
            logger.warning("Max width reached")
            break
        xml_node.append(_create_pywinauto_node(ch, depth + 1, flag))
    return xml_node


# A11y tree getter for macOS


def _create_axui_node(node, nodes: set = None, depth: int = 0, bbox: tuple = None):
    nodes = nodes or set()
    if node in nodes:
        return
    nodes.add(node)

    reserved_keys = {
        "AXEnabled": "st",
        "AXFocused": "st",
        "AXFullScreen": "st",
        "AXTitle": "attr",
        "AXChildrenInNavigationOrder": "attr",
        "AXChildren": "attr",
        "AXFrame": "attr",
        "AXRole": "role",
        "AXHelp": "attr",
        "AXRoleDescription": "role",
        "AXSubrole": "role",
        "AXURL": "attr",
        "AXValue": "val",
        "AXDescription": "attr",
        "AXDOMIdentifier": "attr",
        "AXSelected": "st",
        "AXInvalid": "st",
        "AXRows": "attr",
        "AXColumns": "attr",
    }
    attribute_dict = {}

    if depth == 0:
        bbox = (
            node["kCGWindowBounds"]["X"],
            node["kCGWindowBounds"]["Y"],
            node["kCGWindowBounds"]["X"] + node["kCGWindowBounds"]["Width"],
            node["kCGWindowBounds"]["Y"] + node["kCGWindowBounds"]["Height"],
        )
        app_ref = ApplicationServices.AXUIElementCreateApplication(
            node["kCGWindowOwnerPID"]
        )
        error_code, app_wins_ref = ApplicationServices.AXUIElementCopyAttributeValue(
            app_ref, "AXWindows", None
        )
        if error_code:
            logger.error(
                "MacOS parsing %s encountered Error code: %d", app_ref, error_code
            )

        attribute_dict["name"] = node["kCGWindowOwnerName"]

        node = app_wins_ref[0]

    error_code, attr_names = ApplicationServices.AXUIElementCopyAttributeNames(
        node, None
    )

    if error_code:
        # -25202: AXError.invalidUIElement
        #         The accessibility object received in this event is invalid.
        return

    value = None

    if "AXFrame" in attr_names:
        error_code, attr_val = ApplicationServices.AXUIElementCopyAttributeValue(
            node, "AXFrame", None
        )
        rep = repr(attr_val)
        x_value = re.search(r"x:(-?[\d.]+)", rep)
        y_value = re.search(r"y:(-?[\d.]+)", rep)
        w_value = re.search(r"w:(-?[\d.]+)", rep)
        h_value = re.search(r"h:(-?[\d.]+)", rep)
        type_value = re.search(r"type\s?=\s?(\w+)", rep)
        value = {
            "x": float(x_value.group(1)) if x_value else None,
            "y": float(y_value.group(1)) if y_value else None,
            "w": float(w_value.group(1)) if w_value else None,
            "h": float(h_value.group(1)) if h_value else None,
            "type": type_value.group(1) if type_value else None,
        }

        if not any(v is None for v in value.values()):
            x_min = max(bbox[0], value["x"])
            x_max = min(bbox[2], value["x"] + value["w"])
            y_min = max(bbox[1], value["y"])
            y_max = min(bbox[3], value["y"] + value["h"])

            if x_min > x_max or y_min > y_max:
                # No intersection
                return

    role = None
    text = None

    for attr_name, ns_key in reserved_keys.items():
        if attr_name not in attr_names:
            continue

        if value and attr_name == "AXFrame":
            bb = value
            if not any(v is None for v in bb.values()):
                attribute_dict[
                    "{{{:}}}screencoord".format(_accessibility_ns_map_macos["cp"])
                ] = "({:d}, {:d})".format(int(bb["x"]), int(bb["y"]))
                attribute_dict[
                    "{{{:}}}size".format(_accessibility_ns_map_macos["cp"])
                ] = "({:d}, {:d})".format(int(bb["w"]), int(bb["h"]))
            continue

        error_code, attr_val = ApplicationServices.AXUIElementCopyAttributeValue(
            node, attr_name, None
        )

        full_attr_name = f"{{{_accessibility_ns_map_macos[ns_key]}}}{attr_name}"

        if attr_name == "AXValue" and not text:
            text = str(attr_val)
            continue

        if attr_name == "AXRoleDescription":
            role = attr_val
            continue

        # Set the attribute_dict
        if not (
            isinstance(attr_val, ApplicationServices.AXUIElementRef)
            or isinstance(attr_val, (AppKit.NSArray, list))
        ):
            if attr_val is not None:
                attribute_dict[full_attr_name] = str(attr_val)

    node_role_name = role.lower().replace(" ", "_") if role else "unknown_role"

    xml_node = lxml.etree.Element(
        node_role_name, attrib=attribute_dict, nsmap=_accessibility_ns_map_macos
    )

    if text is not None and len(text) > 0:
        xml_node.text = text

    for attr_name, ns_key in reserved_keys.items():
        if attr_name not in attr_names:
            continue

        error_code, attr_val = ApplicationServices.AXUIElementCopyAttributeValue(
            node, attr_name, None
        )
        if isinstance(attr_val, ApplicationServices.AXUIElementRef):
            _xml_node = _create_axui_node(attr_val, nodes, depth + 1, bbox)
            if _xml_node is not None:
                xml_node.append(_xml_node)

        elif isinstance(attr_val, (AppKit.NSArray, list)):
            for child in attr_val:
                _xml_node = _create_axui_node(child, nodes, depth + 1, bbox)
                if _xml_node is not None:
                    xml_node.append(_xml_node)

    return xml_node


@app.route("/accessibility", methods=["GET"])
def get_accessibility_tree():
    os_name: str = platform.system()

    # AT-SPI works for KDE as well
    if os_name == "Linux":
        global libreoffice_version_tuple
        libreoffice_version_tuple = _get_libreoffice_version()

        desktop: Accessible = pyatspi.Registry.getDesktop(0)
        xml_node = lxml.etree.Element(
            "desktop-frame", nsmap=_accessibility_ns_map_ubuntu
        )
        with concurrent.futures.ThreadPoolExecutor() as executor:
            futures = [
                executor.submit(_create_atspi_node, app_node, 1) for app_node in desktop
            ]
            for future in concurrent.futures.as_completed(futures):
                xml_tree = future.result()
                xml_node.append(xml_tree)
        return jsonify({"AT": lxml.etree.tostring(xml_node, encoding="unicode")})

    elif os_name == "Windows":
        # Attention: Windows a11y tree is implemented to be read through `pywinauto` module, however,
        # two different backends `win32` and `uia` are supported and different results may be returned
        desktop: Desktop = Desktop(backend="uia")
        xml_node = lxml.etree.Element("desktop", nsmap=_accessibility_ns_map_windows)
        with concurrent.futures.ThreadPoolExecutor() as executor:
            futures = [
                executor.submit(_create_pywinauto_node, wnd, 1)
                for wnd in desktop.windows()
            ]
            for future in concurrent.futures.as_completed(futures):
                xml_tree = future.result()
                xml_node.append(xml_tree)
        return jsonify({"AT": lxml.etree.tostring(xml_node, encoding="unicode")})

    elif os_name == "Darwin":
        xml_node = lxml.etree.Element("desktop", nsmap=_accessibility_ns_map_macos)

        with concurrent.futures.ThreadPoolExecutor() as executor:
            futures = [
                executor.submit(_create_axui_node, wnd, None, 0)
                for wnd in [
                    win
                    for win in Quartz.CGWindowListCopyWindowInfo(
                        (
                            Quartz.kCGWindowListExcludeDesktopElements
                            | Quartz.kCGWindowListOptionOnScreenOnly
                        ),
                        Quartz.kCGNullWindowID,
                    )
                    if win["kCGWindowLayer"] == 0
                    and win["kCGWindowOwnerName"] != "Window Server"
                ]
            ]

            for future in concurrent.futures.as_completed(futures):
                xml_tree = future.result()
                if xml_tree is not None:
                    xml_node.append(xml_tree)

        return jsonify({"AT": lxml.etree.tostring(xml_node, encoding="unicode")})

    else:
        return (
            "Currently not implemented for platform {:}.".format(platform.platform()),
            500,
        )


@app.route("/screen_size", methods=["POST"])
def get_screen_size():
    if platform_name == "Linux":
        d = display.Display()
        screen_width = d.screen().width_in_pixels
        screen_height = d.screen().height_in_pixels
    elif platform_name == "Windows":
        user32 = ctypes.windll.user32
        screen_width: int = user32.GetSystemMetrics(0)
        screen_height: int = user32.GetSystemMetrics(1)
    return jsonify({"width": screen_width, "height": screen_height})


@app.route("/window_size", methods=["POST"])
def get_window_size():
    if "app_class_name" in request.form:
        app_class_name = request.form["app_class_name"]
    else:
        return jsonify({"error": "app_class_name is required"}), 400

    d = display.Display()
    root = d.screen().root
    window_ids = root.get_full_property(
        d.intern_atom("_NET_CLIENT_LIST"), X.AnyPropertyType
    ).value

    for window_id in window_ids:
        try:
            window = d.create_resource_object("window", window_id)
            wm_class = window.get_wm_class()

            if wm_class is None:
                continue

            if app_class_name.lower() in [name.lower() for name in wm_class]:
                geom = window.get_geometry()
                return jsonify({"width": geom.width, "height": geom.height})
        except Xlib.error.XError:  # Ignore windows that give an error
            continue
    return None


@app.route("/desktop_path", methods=["POST"])
def get_desktop_path():
    # Get the home directory in a platform-independent manner using pathlib
    home_directory = str(Path.home())

    # Determine the desktop path based on the operating system
    desktop_path = {
        "Windows": os.path.join(home_directory, "Desktop"),
        "Darwin": os.path.join(home_directory, "Desktop"),  # macOS
        "Linux": os.path.join(home_directory, "Desktop"),
    }.get(platform.system(), None)

    # Check if the operating system is supported and the desktop path exists
    if desktop_path and os.path.exists(desktop_path):
        return jsonify(desktop_path=desktop_path)
    else:
        return (
            jsonify(error="Unsupported operating system or desktop path not found"),
            404,
        )


@app.route("/wallpaper", methods=["POST"])
def get_wallpaper():
    def get_wallpaper_windows():
        SPI_GETDESKWALLPAPER = 0x73
        MAX_PATH = 260
        buffer = ctypes.create_unicode_buffer(MAX_PATH)
        ctypes.windll.user32.SystemParametersInfoW(
            SPI_GETDESKWALLPAPER, MAX_PATH, buffer, 0
        )
        return buffer.value

    def get_wallpaper_macos():
        script = """
        tell application "System Events" to tell every desktop to get picture
        """
        process = subprocess.Popen(
            ["osascript", "-e", script], stdout=subprocess.PIPE, stderr=subprocess.PIPE
        )
        output, error = process.communicate()
        if error:
            app.logger.error("Error: %s", error.decode("utf-8"))
            return None
        return output.strip().decode("utf-8")

    def get_wallpaper_linux():
        try:
            output = subprocess.check_output(
                ["gsettings", "get", "org.gnome.desktop.background", "picture-uri"],
                stderr=subprocess.PIPE,
            )
            return (
                output.decode("utf-8").strip().replace("file://", "").replace("'", "")
            )
        except subprocess.CalledProcessError as e:
            app.logger.error("Error: %s", e)
            return None

    os_name = platform.system()
    wallpaper_path = None
    if os_name == "Windows":
        wallpaper_path = get_wallpaper_windows()
    elif os_name == "Darwin":
        wallpaper_path = get_wallpaper_macos()
    elif os_name == "Linux":
        wallpaper_path = get_wallpaper_linux()
    else:
        app.logger.error(f"Unsupported OS: {os_name}")
        abort(400, description="Unsupported OS")

    if wallpaper_path:
        try:
            # Ensure the filename is secure
            return send_file(wallpaper_path, mimetype="image/png")
        except Exception as e:
            app.logger.error(f"An error occurred while serving the wallpaper file: {e}")
            abort(500, description="Unable to serve the wallpaper file")
    else:
        abort(404, description="Wallpaper file not found")


@app.route("/list_directory", methods=["POST"])
def get_directory_tree():
    def _list_dir_contents(directory):
        """
        List the contents of a directory recursively, building a tree structure.

        :param directory: The path of the directory to inspect.
        :return: A nested dictionary with the contents of the directory.
        """
        tree = {
            "type": "directory",
            "name": os.path.basename(directory),
            "children": [],
        }
        try:
            # List all files and directories in the current directory
            for entry in os.listdir(directory):
                full_path = os.path.join(directory, entry)
                # If entry is a directory, recurse into it
                if os.path.isdir(full_path):
                    tree["children"].append(_list_dir_contents(full_path))
                else:
                    tree["children"].append({"type": "file", "name": entry})
        except OSError as e:
            # If the directory cannot be accessed, return the exception message
            tree = {"error": str(e)}
        return tree

    # Extract the 'path' parameter from the JSON request
    data = request.get_json()
    if "path" not in data:
        return jsonify(error="Missing 'path' parameter"), 400

    start_path = data["path"]
    # Ensure the provided path is a directory
    if not os.path.isdir(start_path):
        return jsonify(error="The provided path is not a directory"), 400

    # Generate the directory tree starting from the provided path
    directory_tree = _list_dir_contents(start_path)
    return jsonify(directory_tree=directory_tree)


@app.route("/file", methods=["POST"])
def get_file():
    # Retrieve filename from the POST request
    if "file_path" in request.form:
        file_path = os.path.expandvars(os.path.expanduser(request.form["file_path"]))
    else:
        return jsonify({"error": "file_path is required"}), 400

    try:
        # Check if the file exists and send it to the user
        return send_file(file_path, as_attachment=True)
    except FileNotFoundError:
        # If the file is not found, return a 404 error
        return jsonify({"error": "File not found"}), 404


@app.route("/setup/upload", methods=["POST"])
def upload_file():
    # Retrieve filename from the POST request
    if "file_path" in request.form and "file_data" in request.files:
        file_path = os.path.expandvars(os.path.expanduser(request.form["file_path"]))
        file = request.files["file_data"]
        file.save(file_path)
        return "File Uploaded"
    else:
        return jsonify({"error": "file_path and file_data are required"}), 400


@app.route("/platform", methods=["GET"])
def get_platform():
    return platform.system()


@app.route("/cursor_position", methods=["GET"])
def get_cursor_position():
    return pyautogui.position().x, pyautogui.position().y


@app.route("/setup/change_wallpaper", methods=["POST"])
def change_wallpaper():
    data = request.json
    path = data.get("path", None)

    if not path:
        return "Path not supplied!", 400

    path = Path(os.path.expandvars(os.path.expanduser(path)))

    if not path.exists():
        return f"File not found: {path}", 404

    try:
        user_platform = platform.system()
        if user_platform == "Windows":
            import ctypes

            ctypes.windll.user32.SystemParametersInfoW(20, 0, str(path), 3)
        elif user_platform == "Linux":
            import subprocess

            subprocess.run(
                [
                    "gsettings",
                    "set",
                    "org.gnome.desktop.background",
                    "picture-uri",
                    f"file://{path}",
                ]
            )
        elif user_platform == "Darwin":  # (Mac OS)
            import subprocess

            subprocess.run(
                [
                    "osascript",
                    "-e",
                    f'tell application "Finder" to set desktop picture to POSIX file "{path}"',
                ]
            )
        return "Wallpaper changed successfully"
    except Exception as e:
        return f"Failed to change wallpaper. Error: {e}", 500


def _do_download_file(url: str = None, path: str = None):
    if not url or not path:
        raise ValueError("Path or URL not supplied!")
    path = Path(os.path.expandvars(os.path.expanduser(path)))
    path.parent.mkdir(parents=True, exist_ok=True)

    max_retries = 3
    error: Optional[Exception] = None

    request_proxies = None if should_bypass_proxies(url, no_proxy=None) else proxies
    for i in range(max_retries):
        try:
            response = requests.get(url, stream=True, proxies=request_proxies)
            response.raise_for_status()

            with open(path, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            return f"File downloaded successfully {url} to {path}"

        except requests.RequestException as e:
            error = e
            logger.error(
                f"Failed to download {url}. Retrying... ({max_retries - i - 1} attempts left)"
            )

    if error:
        raise RuntimeError(f"Failed to download {url}. No retries left. Error: {error}")


def _do_batch_download_file(files: List[Dict[str, str]] = None):
    failed_files = []
    if not files:
        raise ValueError("Files not supplied!")
    for file in files:
        url = file.get("url", None)
        path = file.get("path", None)
        try:
            _do_download_file(url, path)
        except Exception as e:
            logger.error(f"Failed to download {url}. Error: {e}")
            failed_files.append({"file": file, "error": str(e)})

    if len(failed_files) > 0:
        raise RuntimeError(f"Failed to download some files. {failed_files}")


@app.route("/setup/download_file", methods=["POST"])
def download_file():
    data = request.json
    url = data.get("url", None)
    path = data.get("path", None)

    try:
        result = _do_download_file(url, path)
        return result
    except Exception as e:
        return str(e), 500


def _do_open_file(path: str = None):
    if not path:
        raise ValueError("Path not supplied!")
    path = Path(os.path.expandvars(os.path.expanduser(path)))
    if not path.exists():
        raise RuntimeError(f"File not found: {path}")

    open_cmd: str = "xdg-open"
    subprocess.Popen([open_cmd, str(path)])


@app.route("/setup/open_file", methods=["POST"])
def open_file():
    data = request.json
    path = data.get("path", None)

    try:
        _do_open_file(path)
        return "File opened successfully"
    except Exception as e:
        return f"Failed to open {path}. Error: {e}", 500


def _do_activate_window(window_name: str, strict: bool = False, by_class: bool = False):
    subprocess.run(
        [
            "wmctrl",
            "-{:}{:}a".format("x" if by_class else "", "F" if strict else ""),
            window_name,
        ]
    )


@app.route("/setup/activate_window", methods=["POST"])
def activate_window():
    data = request.json
    window_name = data.get("window_name", None)
    if not window_name:
        return "window_name required", 400
    strict: bool = data.get(
        "strict", False
    )  # compare case-sensitively and match the whole string
    by_class_name: bool = data.get("by_class", False)

    try:
        _do_activate_window(window_name, strict, by_class_name)
        return "Window activated successfully", 200
    except Exception as e:
        return str(e), 500


def _do_close_window(window_name: str, strict: bool = False, by_class: bool = False):
    subprocess.run(
        [
            "wmctrl",
            "-{:}{:}c".format("x" if by_class else "", "F" if strict else ""),
            window_name,
        ]
    )


@app.route("/setup/close_window", methods=["POST"])
def close_window():
    data = request.json
    if "window_name" not in data:
        return "window_name required", 400
    window_name: str = data["window_name"]
    strict: bool = data.get(
        "strict", False
    )  # compare case-sensitively and match the whole string
    by_class_name: bool = data.get("by_class", False)

    try:
        _do_close_window(window_name, strict, by_class_name)
        return "Window closed successfully.", 200
    except Exception as e:
        return str(e), 500


@app.route("/start_recording", methods=["POST"])
def start_recording():
    global recording_process
    if recording_process:
        return (
            jsonify(
                {"status": "error", "message": "Recording is already in progress."}
            ),
            400,
        )

    d = display.Display()
    screen_width = d.screen().width_in_pixels
    screen_height = d.screen().height_in_pixels

    start_command = f"ffmpeg -y -f x11grab -draw_mouse 1 -s {screen_width}x{screen_height} -i :0.0 -c:v libx264 -r 30 {recording_path}"

    recording_process = subprocess.Popen(
        shlex.split(start_command), stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
    )

    return jsonify({"status": "success", "message": "Started recording."})


@app.route("/end_recording", methods=["POST"])
def end_recording():
    global recording_process

    if not recording_process:
        return (
            jsonify(
                {"status": "error", "message": "No recording in progress to stop."}
            ),
            400,
        )

    recording_process.send_signal(signal.SIGINT)
    recording_process.wait()
    recording_process = None

    # return recording video file
    if os.path.exists(recording_path):
        return send_file(recording_path, as_attachment=True)
    else:
        return abort(404, description="Recording failed")


def compare_urls(url1, url2):
    from urllib.parse import urlparse, urlunparse

    if url1 is None or url2 is None:
        return url1 == url2

    def normalize_url(url):
        # Parse the URL
        parsed_url = urlparse(url)

        # If no scheme is present, assume 'http'
        scheme = parsed_url.scheme if parsed_url.scheme else "http"

        # Lowercase the scheme and netloc, remove 'www.', and handle trailing slash
        normalized_netloc = parsed_url.netloc.lower().replace("www.", "")
        normalized_path = parsed_url.path if parsed_url.path != "/" else ""

        # Reassemble the URL with normalized components
        normalized_parsed_url = parsed_url._replace(
            scheme=scheme.lower(), netloc=normalized_netloc, path=normalized_path
        )
        normalized_url = urlunparse(normalized_parsed_url)

        return normalized_url

    # Normalize both URLs for comparison
    norm_url1 = normalize_url(url1)
    norm_url2 = normalize_url(url2)

    # Compare the normalized URLs
    return norm_url1 == norm_url2


def _do_chrome_open_tabs(urls_to_open: List[str]):
    from playwright.sync_api import sync_playwright, TimeoutError

    local_debug_url = "http://localhost:1337"
    with sync_playwright() as p:
        browser = None
        for attempt in range(15):
            try:
                browser = p.chromium.connect_over_cdp(local_debug_url)
                break
            except Exception as e:
                if attempt < 14:
                    logger.error(
                        f"Attempt {attempt + 1}: Failed to connect, retrying. Error: {e}"
                    )
                    time.sleep(1)
                else:
                    logger.error(f"Failed to connect after multiple attempts: {e}")
                    raise e

        if not browser:
            raise RuntimeError(f"Failed to connect after multiple attempts: {e}")

        logger.info("Opening %s...", urls_to_open)
        for i, url in enumerate(urls_to_open):
            # Use the first context (which should be the only one if using default profile)
            if i == 0:
                context = browser.contexts[0]

            page = (
                context.new_page()
            )  # Create a new page (tab) within the existing context
            try:
                page.goto(url, timeout=60000)
            except:
                logger.warning(
                    "Opening %s exceeds time limit", url
                )  # only for human test
            logger.info(f"Opened tab {i + 1}: {url}")

            if i == 0:
                # clear the default tab
                default_page = context.pages[0]
                default_page.close()
        # Do not close the context or browser; they will remain open after script ends
        return browser, context


def _do_chrome_close_tabs(urls_to_close: List[str]):
    from playwright.sync_api import sync_playwright, TimeoutError

    local_debug_url = "http://localhost:1337"
    with sync_playwright() as p:
        browser = None
        for attempt in range(15):
            try:
                browser = p.chromium.connect_over_cdp(local_debug_url)
                break
            except Exception as e:
                if attempt < 14:
                    logger.error(
                        f"Attempt {attempt + 1}: Failed to connect, retrying. Error: {e}"
                    )
                    time.sleep(1)
                else:
                    logger.error(f"Failed to connect after multiple attempts: {e}")
                    raise e

        if not browser:
            raise RuntimeError(f"Failed to connect after multiple attempts: {e}")

        for i, url in enumerate(urls_to_close):
            # Use the first context (which should be the only one if using default profile)
            if i == 0:
                context = browser.contexts[0]

            for page in context.pages:

                # if two urls are the same, close the tab
                if compare_urls(page.url, url):
                    context.pages.pop(context.pages.index(page))
                    page.close()
                    logger.info(f"Closed tab {i + 1}: {url}")
                    break

        # Do not close the context or browser; they will remain open after script ends
        return browser, context


def _do_update_browse_history(history: List[Dict[str, Any]]):
    import sqlite3
    from datetime import datetime, timedelta

    # get the path of the history file according to the platform
    if "arm" in platform.machine():
        chrome_history_path = os.path.join(
            os.getenv("HOME"),
            "snap",
            "chromium",
            "common",
            "chromium",
            "Default",
            "History",
        )
    else:
        chrome_history_path = os.path.join(
            os.getenv("HOME"), ".config", "google-chrome", "Default", "History"
        )

    db_url = "https://drive.usercontent.google.com/u/0/uc?id=1Lv74QkJYDWVX0RIgg0Co-DUcoYpVL0oX&export=download"  # google drive
    max_retries = 3
    downloaded = False
    e = None

    request_proxies = None if should_bypass_proxies(db_url, no_proxy=None) else proxies
    for i in range(max_retries):
        try:
            response = requests.get(db_url, stream=True, proxies=request_proxies)
            response.raise_for_status()

            with open(chrome_history_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            logger.info("File downloaded successfully")
            downloaded = True
            break

        except requests.RequestException as e:
            logger.error(
                f"Failed to download {db_url} caused by {e}. Retrying... ({max_retries - i - 1} attempts left)"
            )
    if not downloaded:
        raise requests.RequestException(
            f"Failed to download {db_url}. No retries left. Error: {e}"
        )

    for history_item in history:
        url = history_item["url"]
        title = history_item["title"]
        visit_time = datetime.now() - timedelta(
            seconds=history_item["visit_time_from_now_in_seconds"]
        )

        # Chrome use ms from 1601-01-01 as timestamp
        epoch_start = datetime(1601, 1, 1)
        chrome_timestamp = int((visit_time - epoch_start).total_seconds() * 1000000)

        conn = sqlite3.connect(chrome_history_path)
        cursor = conn.cursor()

        cursor.execute(
            """
                INSERT INTO urls (url, title, visit_count, typed_count, last_visit_time, hidden)
                VALUES (?, ?, ?, ?, ?, ?)
            """,
            (url, title, 1, 0, chrome_timestamp, 0),
        )

        url_id = cursor.lastrowid

        cursor.execute(
            """
                INSERT INTO visits (url, visit_time, from_visit, transition, segment_id, visit_duration)
                VALUES (?, ?, ?, ?, ?, ?)
            """,
            (url_id, chrome_timestamp, 0, 805306368, 0, 0),
        )

        conn.commit()
        conn.close()

    logger.info("Fake browsing history added successfully.")

    # _do_execute(
    #     ["sudo chown -R user:user /home/user/.config/google-chrome/Default/History"],
    #     shell=True,
    # )

@app.route('/execute_bash', methods=['POST'])
def execute_bash():
    """
    Executes a bash command in a persistent tmux session.
    """
    data = request.json
    command = data.get("command", "")
    session_name = data.get("session_name", TMUX_SESSION)
    timeout = data.get("timeout", 600)
    start_time = time.time()

    # Check if the session exists. If not, create it.
    success, error = _ensure_tmux_session(session_name)
    if not success:
        return jsonify({
            "status": "error", "stdout": "", "stderr": error,
            "return_code": -1, "execution_time": time.time() - start_time
        }), 500

    try:
        # Case 1: The command is empty. Capture the entire pane for logging purposes.
        if not command:
            return jsonify({
                "status": "success", "stdout": last_command_info["output"], "stderr": "",
                "return_code": 0, "execution_time": time.time() - start_time
            })

        # Case 2: Interrupt command
        if command == "C-c":
            subprocess.run(["tmux", "send-keys", "-t", session_name, "C-c"], check=True)
            time.sleep(0.5)

            # Now, try to get the output using the markers from the interrupted command
            if not last_command_info.get("markers"):
                return jsonify({
                    "status": "success", "stdout": "No command seems to be running", "stderr": "",
                    "return_code": 0, "execution_time": time.time() - start_time
                })

            markers = last_command_info["markers"]
            start_marker = markers["start"]
            end_marker = markers["end"]
            exit_code_marker = markers["exit_code"]

            command_executed, pane_content = _poll_for_output(session_name, start_marker, end_marker, timeout=5)

            if not command_executed:
                return jsonify({
                    "status": "error", "stdout": "", "stderr": "Failed to get output of interrupted command after 5s.",
                    "return_code": -1, "execution_time": time.time() - start_time
                }), 500

            output, exit_code, error_message = _parse_output(pane_content, start_marker, end_marker, exit_code_marker)

            if error_message:
                last_command_info["markers"] = None  # Clear markers on error
                return jsonify({
                    "status": "error", "stdout": "", "stderr": error_message,
                    "return_code": -1, "execution_time": time.time() - start_time
                }), 500

            last_command_info["exit_code"] = exit_code
            last_command_info["output"] = output
            last_command_info["markers"] = None  # Clear markers on success
            return jsonify({
                "status": "success", "stdout": "", "stderr": "Command interrupted.",
                "return_code": 130, "execution_time": time.time() - start_time
            })

        # Case 3: A normal command is provided. Execute it and capture the output precisely using markers.
        output, exit_code, error_message = _execute_normal_command(command, session_name, timeout)
        execution_time = time.time() - start_time

        if error_message:
            last_command_info["exit_code"] = exit_code if exit_code is not None else -1
            last_command_info["output"] = error_message
            return jsonify({
                "status": "error", "stdout": "", "stderr": error_message,
                "return_code": exit_code if exit_code is not None else -1, "execution_time": execution_time
            }), 500

        last_command_info["exit_code"] = exit_code
        last_command_info["output"] = output

        if exit_code == 0:
            return jsonify({
                "status": "success", "stdout": output, "stderr": "",
                "return_code": exit_code, "execution_time": execution_time
            })
        else:
            return jsonify({
                "status": "error", "stdout": "", "stderr": output,
                "return_code": exit_code, "execution_time": execution_time
            }), 500

    except subprocess.CalledProcessError as e:
        execution_time = time.time() - start_time
        error_message = f"Command execution failed: {e.stderr if e.stderr else e.stdout}"
        return_code = e.returncode if hasattr(e, 'returncode') else 1
        last_command_info["exit_code"] = return_code
        last_command_info["output"] = error_message
        return jsonify({
            "status": "error", "stdout": "", "stderr": error_message,
            "return_code": return_code, "execution_time": execution_time
        }), 500
    except FileNotFoundError:
        execution_time = time.time() - start_time
        error_message = "tmux command not found. Is tmux installed and in PATH?"
        last_command_info["exit_code"] = -1
        last_command_info["output"] = error_message
        return jsonify({
            "status": "error", "stdout": "", "stderr": error_message,
            "return_code": -1, "execution_time": execution_time
        }), 500

def _ensure_tmux_session(session_name):
    """Ensures a tmux session exists, creating it if necessary."""
    try:
        subprocess.run(["tmux", "has-session", "-t", session_name], check=True, capture_output=True, text=True)
    except (subprocess.CalledProcessError, FileNotFoundError):
        try:
            subprocess.run(["tmux", "new-session", "-d", "-s", session_name], check=True)
            logger.info(f"Created tmux session '{session_name}'")
        except (subprocess.CalledProcessError, FileNotFoundError) as e:
            logger.error(f"Failed to create tmux session '{session_name}': {e}")
            return False, f"Failed to create tmux session: {e}"
    return True, None

def _poll_for_output(session_name, start_marker, end_marker, timeout):
    """Polls the tmux pane for the command output markers."""
    polling_end_time = time.time() + timeout + 2  # 2s buffer
    pane_content = ""
    while time.time() < polling_end_time:
        pane_content = subprocess.run(
            ["tmux", "capture-pane", "-p", "-S", "-10000", "-E", "-", "-t", session_name],
            capture_output=True, text=True, check=True
        ).stdout.strip()

        lines = pane_content.split('\n')
        last_start_line_idx = -1
        try:
            last_start_line_idx = len(lines) - 1 - lines[::-1].index(start_marker)
        except ValueError:
            time.sleep(0.2)
            continue

        if end_marker in lines[last_start_line_idx:]:
            return True, pane_content

        time.sleep(0.2)
    return False, pane_content

def _parse_output(pane_content, start_marker, end_marker, exit_code_marker):
    """Parses the tmux pane content to extract command output and exit code."""
    start_index = pane_content.rfind(start_marker)
    end_index = pane_content.rfind(end_marker)

    if start_index == -1 or end_index == -1 or end_index < start_index:
        error_message = "Could not reliably determine command output."
        debug_output = "\\n".join(pane_content.strip().split("\\n")[-10:])
        return None, None, f"{error_message}\\nDebug output:\\n{debug_output}"

    output_block = pane_content[start_index + len(start_marker) + 1:end_index]
    exit_code_marker_index = output_block.rfind(exit_code_marker)

    if exit_code_marker_index == -1:
        return None, None, "Could not find exit code marker in output."

    output = output_block[:exit_code_marker_index].strip()
    exit_code_str = output_block[exit_code_marker_index + len(exit_code_marker):].strip()

    try:
        exit_code = int(exit_code_str)
    except (ValueError, TypeError):
        return f"Could not parse exit code. Raw output block:\\n{output_block}", -1, None

    return output, exit_code, None


def _execute_normal_command(command, session_name, timeout):
    """Executes a normal command in the tmux session."""
    start_marker = f"---TMUX_OUTPUT_START_{uuid.uuid4()}---"
    end_marker = f"---TMUX_OUTPUT_END_{uuid.uuid4()}---"
    exit_code_marker = f"---TMUX_EXIT_CODE_{uuid.uuid4()}---"

    # Store markers for potential interruption
    last_command_info["markers"] = {
        "start": start_marker,
        "end": end_marker,
        "exit_code": exit_code_marker
    }

    timed_command = f"timeout -s INT -k 1s {timeout} {command}"
    wrapped_command = f"{{ {timed_command}; }} 2>&1; echo '{exit_code_marker}'$?"
    full_command = f"echo '{start_marker}'; {wrapped_command}; echo '{end_marker}'"

    subprocess.run(["tmux", "send-keys", "-t", session_name, full_command, "C-m"], check=True)

    command_executed, pane_content = _poll_for_output(session_name, start_marker, end_marker, timeout)

    if not command_executed:
        subprocess.run(["tmux", "send-keys", "-t", session_name, "C-c"], check=True)
        error_message = f"Command '{command}' timed out after {timeout} seconds (or polling failed)."
        return None, -1, error_message

    output, exit_code, error_message = _parse_output(pane_content, start_marker, end_marker, exit_code_marker)

    return output, exit_code, error_message

# --- JupyterCI Implementation (using jupyter_client) ---

km = None
kc = None

def start_kernel(start_time):
    """Starts and manages a single Jupyter kernel."""
    global km, kc
    if kc and kc.is_alive():
        return  # Kernel is already running and alive

    # If kernel died, clean up before restarting
    if kc:
        kc.stop_channels()
    if km and km.has_kernel:
        km.shutdown_kernel(now=True)

    # Start a new kernel
    km = KernelManager(kernel_name='python3')
    km.start_kernel()
    kc = km.client()
    kc.start_channels()

    # Wait for the kernel to be ready
    try:
        kc.wait_for_ready(timeout=60)
    except RuntimeError:
        stop_kernel()
        raise Exception("Jupyter kernel failed to start in 60 seconds.")

def stop_kernel():
    """Stops the Jupyter kernel gracefully."""
    global km, kc
    if kc:
        kc.stop_channels()
    if km and km.has_kernel:
        km.shutdown_kernel(now=True)

# Register the stop_kernel function to be called on exit
atexit.register(stop_kernel)

def _process_jupyter_messages(msg_id, timeout):
    """Processes messages from the Jupyter kernel's IOPub channel."""
    stdout = ""
    stderr = ""
    execution_successful = False

    while True:
        try:
            # Wait for a message on the IOPub channel
            msg = kc.get_iopub_msg(timeout=timeout)

            # Check if the message belongs to our execution request
            if msg.get('parent_header', {}).get('msg_id') != msg_id:
                continue

            msg_type = msg['header']['msg_type']
            content = msg['content']

            if msg_type == 'status':
                if content['execution_state'] == 'idle':
                    # Idle status means execution is complete
                    execution_successful = True
                    break
            elif msg_type == 'stream':
                stdout += content['text']
            elif msg_type == 'execute_result':
                if 'data' in content and 'text/plain' in content['data']:
                    stdout += content['data']['text/plain']
            elif msg_type == 'error':
                stderr += f"{content.get('ename', 'Error')}: {content.get('evalue', '')}\\n"
                stderr += "\\n".join(content.get('traceback', []))
                execution_successful = False
                break
        except queue.Empty:
            km.interrupt_kernel()
            raise TimeoutError(f"Code execution timed out after {timeout} seconds")
        except Exception as e:
            raise RuntimeError(f"An unexpected error occurred while processing kernel messages: {e}")

    return stdout, stderr, execution_successful

@app.route('/jupyter_ci', methods=['POST'])
def jupyter_ci():
    """Executes Python code in a stateful Jupyter kernel."""
    start_time = time.time()
    try:
        start_kernel(start_time)
    except Exception as e:
        return jsonify({
            "status": "error", "stdout": "", "stderr": f"Jupyter kernel error: {e}",
            "return_code": -1, "execution_time": time.time() - start_time
        }), 500

    data = request.json
    code = data.get("code")
    timeout = data.get("timeout", 60)

    if not code:
        return jsonify({
            "status": "error", "stdout": "", "stderr": "No code provided",
            "return_code": -1, "execution_time": time.time() - start_time
        }), 400

    try:
        msg_id = kc.execute(code)
        stdout, stderr, execution_successful = _process_jupyter_messages(msg_id, timeout)
        execution_time = time.time() - start_time

        if stderr:
            return jsonify({
                "status": "error", "stdout": stdout, "stderr": stderr,
                "return_code": 1, "execution_time": round(execution_time, 2)
            })
        
        if execution_successful:
            return jsonify({
                "status": "success", "stdout": stdout, "stderr": "",
                "return_code": 0, "execution_time": round(execution_time, 2)
            })
        else:
            return jsonify({
                "status": "error", "stdout": stdout,
                "stderr": "Execution failed without a specific error message.",
                "return_code": -1, "execution_time": round(execution_time, 2)
            })

    except TimeoutError as e:
        return jsonify({
            "status": "error", "stdout": "", "stderr": str(e),
            "return_code": -1, "execution_time": time.time() - start_time
        }), 408
    except (RuntimeError, Exception) as e:
        return jsonify({
            "status": "error", "stdout": "", "stderr": str(e),
            "return_code": -1, "execution_time": time.time() - start_time
        }), 500


# --- End JupyterCI Implementation ---


def _get_backup_path(file_path: str) -> str:
    """Constructs the backup file path."""
    return f"{file_path}.bak.vm_control_server"

def _handle_view(path, data):
    """Handle the 'view' command."""
    MAX_OUTPUT_LENGTH = 5000
    if not os.path.exists(path):
        return {"status": "error", "message": f"Path does not exist: {path}"}, 404

    content = ""
    if os.path.isfile(path):
        view_range = data.get("view_range")
        with open(path, 'r', encoding='utf-8') as f:
            lines = f.readlines()

        if view_range:
            try:
                start_line, end_line = view_range
                if not isinstance(start_line, int) or not isinstance(end_line, int):
                    return {
                        "status": "error", "stdout": "", "stderr": "view_range must contain integers."
                    }, 400
                if start_line < 1:
                    return {
                        "status": "error", "stdout": "", "stderr": "start_line must be 1 or greater."
                    }, 400

                start_index = start_line - 1

                if end_line == -1:
                    target_lines = lines[start_index:]
                else:
                    if end_line < start_line:
                        return jsonify({
                            "status": "error", "stdout": "", "stderr": "end_line must be greater than or equal to start_line."
                        }), 400
                    end_index = end_line
                    target_lines = lines[start_index:end_index]

                numbered_lines = [f"{i + start_line: >6}  {line}" for i, line in enumerate(target_lines)]
                content = "".join(numbered_lines)

            except (ValueError, TypeError, IndexError) as e:
                return {
                    "status": "error", "output": "", "error": f"Invalid view_range parameter: {e}"
                }, 400
        else:
            numbered_lines = [f"{i + 1: >6}  {line}" for i, line in enumerate(lines)]
            content = "".join(numbered_lines)

    elif os.path.isdir(path):
        dir_content = []
        for item1 in sorted(os.listdir(path)):
            if item1.startswith('.'):
                continue
            dir_content.append(item1)
            path1 = os.path.join(path, item1)
            if os.path.isdir(path1):
                for item2 in sorted(os.listdir(path1)):
                    if item2.startswith('.'):
                        continue
                    dir_content.append(os.path.join(item1, item2))
        content = "\n".join(dir_content)

    if len(content) > MAX_OUTPUT_LENGTH:
        content = content[:MAX_OUTPUT_LENGTH] + "\n<response clipped>"

    return {"status": "success", "stdout": content, "stderr": "", "return_code": 0}, 200

def _handle_create(path, data):
    """Handle the 'create' command."""
    if os.path.exists(path):
        return {
            "status": "error", "stdout": "", "stderr": f"File or directory already exists: {path}"
        }, 409
    file_text = data.get("file_text", "")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(file_text)
    return {"status": "success", "stdout": f"File created at {path}", "stderr": "", "return_code": 0}, 200


def _handle_str_replace(path, data, backup_path):
    """Handle the 'str_replace' command."""
    old_str = data.get("old_str")
    new_str = data.get("new_str", "")
    if old_str is None:
        return {
            "status": "error", "stdout": "", "stderr": "Missing 'old_str' for str_replace."
        }, 400

    if not os.path.isfile(path):
        return {"status": "error", "stdout": "", "stderr": f"Path is not a file: {path}"}, 400

    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()

    if content.count(old_str) != 1:
        return {
            "status": "error", "stdout": "", "stderr": f"Expected 'old_str' to appear once, but found {content.count(old_str)} times."
        }, 400

    shutil.copy2(path, backup_path)
    new_content = content.replace(old_str, new_str, 1)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(new_content)

    return {"status": "success", "stdout": "String replaced successfully.", "stderr": "", "return_code": 0}, 200


def _handle_insert(path, data, backup_path):
    """Handle the 'insert' command."""
    new_str = data.get("new_str")
    insert_line = data.get("insert_line")
    if new_str is None or insert_line is None:
        return {
            "status": "error", "stdout": "", "stderr": "Missing 'new_str' or 'insert_line' for insert."
        }, 400

    if not os.path.isfile(path):
        return {"status": "error", "stdout": "", "stderr": f"Path is not a file: {path}"}, 400

    shutil.copy2(path, backup_path)
    with open(path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    lines.insert(insert_line, new_str + '\n')
    with open(path, 'w', encoding='utf-8') as f:
        f.writelines(lines)

    return {"status": "success", "stdout": f"Content inserted after line {insert_line}.", "stderr": "", "return_code": 0}, 200


def _handle_undo_edit(path, backup_path):
    """Handle the 'undo_edit' command."""
    if not os.path.exists(backup_path):
        return {"status": "error", "stdout": "", "stderr": "No edit to undo."}, 404

    shutil.move(backup_path, path)
    return {"status": "success", "stdout": "Last edit has been undone.", "stderr": "", "return_code": 0}, 200

def _validate_path(path):
    """Validates the provided file path."""
    if not path:
        return "Missing 'path' parameter."
    if not os.path.isabs(path):
        return "Only absolute paths are allowed."
    if ".." in path.split(os.sep):
        return "Directory traversal is not allowed."
    return None

@app.route('/str_replace_editor', methods=['POST'])
def str_replace_editor():
    """
    A stateful tool to view, create, and edit plain text files.
    It supports undoing the last modification.
    """
    start_time = time.time()
    data = request.json
    command = data.get("command")
    path = data.get("path")

    path_error = _validate_path(path)
    if path_error:
        execution_time = time.time() - start_time
        return jsonify({"status": "error", "stdout": "", "stderr": path_error, "return_code": -1,
                        "execution_time": execution_time}), 400

    if not command:
        execution_time = time.time() - start_time
        return jsonify({"status": "error", "stdout": "", "stderr": "Missing 'command' parameter.", "return_code": -1,
                        "execution_time": execution_time}), 400

    backup_path = _get_backup_path(path)

    handlers = {
        "view": lambda: _handle_view(path, data),
        "create": lambda: _handle_create(path, data),
        "str_replace": lambda: _handle_str_replace(path, data, backup_path),
        "insert": lambda: _handle_insert(path, data, backup_path),
        "undo_edit": lambda: _handle_undo_edit(path, backup_path),
    }

    handler = handlers.get(command)
    if not handler:
        return jsonify({
            "status": "error",
            "stdout": "",
            "stderr": f"Unknown command: {command}",
            "return_code": -1,
            "execution_time": time.time() - start_time
        }), 400

    try:
        result, status_code = handler()
        execution_time = time.time() - start_time
        result["execution_time"] = execution_time
        return jsonify(result), status_code
    except Exception as e:
        return jsonify({
            "status": "error",
            "stdout": "",
            "stderr": str(e),
            "return_code": -1,
            "execution_time": time.time() - start_time
        }), 500

if __name__ == "__main__":
    app.run(debug=True, host="::")