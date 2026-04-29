from __future__ import annotations

import ctypes
from ctypes import wintypes
from dataclasses import dataclass
import logging
import os
import shutil
from pathlib import Path
import re
import subprocess
import time

from automation_components import SOFTWARE_SPECS, connect_software_window
from ui_automation import (
    UiAutomationError,
    accept_overwrite_confirmation,
    bring_window_to_front,
    click_control,
    click_text_control,
    control_exists,
    connect_window,
    dismiss_close_prompts,
    fill_file_dialog,
    find_labeled_edit,
    find_text_control,
    get_clipboard_text,
    get_foreground_window,
    minimize_window,
    paste_text_via_clipboard,
    request_window_close,
    send_hotkey,
    wait_for_desktop_text_control,
    wait_for_text_control,
    wait_for_file_dialog,
    wait_for_window,
)


logger = logging.getLogger(__name__)


OPEN_DIALOG_PATTERNS = (
    r"^Open$",
    r"^Select",
    r"^Browse",
    r"^\u6253\u5f00$",
    r"^\u6253\u5f00\u6587\u4ef6$",
)
SAVE_DIALOG_PATTERNS = (
    r"^Save$",
    r"^Save As$",
    r"^Browse",
    r"^\u4fdd\u5b58$",
    r"^\u53e6\u5b58\u4e3a$",
)
SHOTCUT_SAVE_DIALOG_PATTERNS = (
    r"^Export Video/Audio$",
    r"^Save$",
    r"^Save As$",
    r"^\u4fdd\u5b58$",
    r"^\u53e6\u5b58\u4e3a$",
)
SHOTCUT_SAVE_CONFIRM_PATTERNS = (r"^\u4fdd\u5b58", r"^Save")
SHOTCUT_TASK_TIME_PATTERN = r"^\d{2}:\d{2}:\d{2}$"
SHOTCUT_OPEN_BUTTON_PATTERNS = (r"^\u6253\u5f00\u6587\u4ef6$", r"^Open File$")
SHOTCUT_OUTPUT_BUTTON_PATTERNS = (r"^\u8f93\u51fa$", r"^Export$")
SHOTCUT_TASKS_BUTTON_PATTERNS = (r"^\u4efb\u52a1$", r"^Jobs$")
SHOTCUT_PAUSE_QUEUE_PATTERNS = (r"^\u6682\u505c\u961f\u5217$", r"^Pause Queue$")
SHOTCUT_SAVE_CHANGES_TEXT_PATTERNS = (
    r"^The project has been modified.*$",
    r"^Do you want to save your changes.*$",
    r"^Save your changes.*$",
    r"^.*\u9879\u76ee\u5df2\u88ab\u4fee\u6539.*$",
    r"^.*\u4f60\u60f3\u4fdd\u5b58.*\u4fee\u6539.*$",
)
SHOTCUT_RECOVERY_TEXT_PATTERNS = (
    r"^Autosave files exist.*$",
    r"^Do you want to recover.*$",
    r"^.*autosave.*recover.*$",
    r"^.*\u81ea\u52a8\u4fdd\u5b58.*\u6587\u4ef6.*$",
    r"^.*\u60a8\u60f3\u6062\u590d.*$",
    r"^.*\u6062\u590d\u5b83\u4eec.*$",
)
SHOTCUT_RECOVERY_DISMISS_PATTERNS = (
    r"^No(?:\([A-Z]\))?$",
    r"^\u5426(?:\([A-Z]\))?$",
    r"^Do not recover(?:\([A-Z]\))?$",
)
SHOTCUT_RECOVERY_DISMISS_KEY = "%n"
SHOTCUT_DONT_SAVE_DIRECT_KEYS = ("%n", "n")
SHOTCUT_DONT_SAVE_PATTERNS = (
    r"^No(?:\([A-Z]\))?$",
    r"^\u5426(?:\([A-Z]\))?$",
    r"^Don't Save(?:\([A-Z]\))?$",
    r"^Discard(?:\([A-Z]\))?$",
)
SHOTCUT_APPEND_TO_TIMELINE_SHORTCUT = "a"
SHOTCUT_CLOSE_HOTKEY = "%{F4}"

SHOTCUT_RECONNECT_TIMEOUT_SECONDS = 8.0
SHOTCUT_DIALOG_TIMEOUT_SECONDS = 5.0
SHOTCUT_CONTROL_TIMEOUT_SECONDS = 5.0
SHOTCUT_EXPORT_READY_TIMEOUT_SECONDS = 8.0
SHOTCUT_EXPORT_BUTTON_RETRY_COUNT = 2
SHOTCUT_EXPORT_BUTTON_POST_CLICK_SECONDS = 0.5
SHOTCUT_JOBS_TIMEOUT_SECONDS = 6.0
SHOTCUT_EXPORT_POLL_SECONDS = 1.0
SHOTCUT_SAVE_DIALOG_OVERWRITE_TIMEOUT_SECONDS = 0.5
WAIT_PROGRESS_LOG_INTERVAL_SECONDS = 15.0

AVIDEMUX_MAIN_WINDOW_CLASS = "MainWindow"
AVIDEMUX_OPEN_DIALOG_PATTERNS = (
    r"^Select Video File.*$",
    r"^Open(?: File)?$",
    r"^\u6253\u5f00(?:\u6587\u4ef6)?$",
)
AVIDEMUX_INFO_DIALOG_PATTERNS = (r"^Info$", r"^\u4fe1\u606f$")
AVIDEMUX_THANKS_DIALOG_PATTERNS = (r"^Thanks!?$",)
AVIDEMUX_DIALOG_OK_PATTERNS = (r"^OK$", r"^\u786e\u5b9a$")
AVIDEMUX_OPEN_FAILURE_PATTERNS = (
    r"^Could not open the file$",
    r"^Cannot find a demuxer for .*$",
    r"^\u65e0\u6cd5\u6253\u5f00\u8be5\u6587\u4ef6$",
)
AVIDEMUX_DEMUXER_FAILURE_PATTERNS = (r"^Cannot find a demuxer for .*$",)
AVIDEMUX_EXPORT_SUCCESS_PATTERNS = (
    r"^Done$",
    r"^File .+ has been successfully saved\.$",
    r"^\u5df2\u5b8c\u6210$",
    r"^.+\u5df2\u6210\u529f\u4fdd\u5b58.*$",
)
AVIDEMUX_CODEC_SELECTION_TIMEOUT_SECONDS = 8.0
AVIDEMUX_DIALOG_TIMEOUT_SECONDS = 8.0
AVIDEMUX_DIALOG_SETTLE_SECONDS = 0.6
AVIDEMUX_DIALOG_POLL_SECONDS = 0.2
AVIDEMUX_EXPORT_POLL_SECONDS = 2.0
AVIDEMUX_EXPORT_STABLE_ROUNDS = 5
AVIDEMUX_EXPORT_TIMEOUT_SECONDS = 7200.0
AVIDEMUX_SAVE_START_TIMEOUT_SECONDS = 20.0
AVIDEMUX_CLOSE_RETRY_COUNT = 3
AVIDEMUX_CLOSE_SETTLE_SECONDS = 0.6
AVIDEMUX_CLOSE_TIMEOUT_SECONDS = 2.5
AVIDEMUX_FILE_ENTRY_MIN_WIDTH = 300
AVIDEMUX_FILE_ENTRY_MIN_LEFT = 150
AVIDEMUX_FILE_ENTRY_TOP_MIN = 520
AVIDEMUX_FILE_ENTRY_TOP_MAX = 760
AVIDEMUX_LIST_ITEM_LEFT_OFFSET = 250
AVIDEMUX_LIST_ITEM_TOP_OFFSET = 120
AVIDEMUX_LIST_ITEM_BOTTOM_OFFSET = 200
AVIDEMUX_ACTION_BUTTON_LEFT_MIN = 700
AVIDEMUX_ACTION_BUTTON_TOP_MIN = 680
AVIDEMUX_ACTION_BUTTON_TOP_MAX = 820
AVIDEMUX_OPEN_CONFIRM_GAP_RATIO = 0.14
AVIDEMUX_OPEN_CONFIRM_MIN_GAP = 12
AVIDEMUX_VIDEO_CODEC_TEXT = "HEVC (x265)"
AVIDEMUX_MUXER_TEXT = "MKV Muxer"
AVIDEMUX_YES_PATTERNS = (r"^Yes$", r"^\u662f$")
AVIDEMUX_OVERWRITE_CONFIRM_PATTERNS = (
    r"^Yes(?:\([A-Z]\))?$",
    r"^Overwrite(?:\([A-Z]\))?$",
    r"^Replace(?:\([A-Z]\))?$",
    r"^\u662f(?:\([A-Z]\))?$",
    r"^\u8986\u5199(?:\([A-Z]\))?$",
    r"^\u8986\u76d6(?:\([A-Z]\))?$",
    r"^\u590d\u5199(?:\([A-Z]\))?$",
)
AVIDEMUX_NO_PATTERNS = (r"^No$", r"^\u5426$")
AVIDEMUX_CANCEL_PATTERNS = (r"^Cancel$", r"^\u53d6\u6d88$")
AVIDEMUX_SAVE_BUTTON_PATTERNS = (r"^Save(?:\([A-Z]\))?$", r"^\u4fdd\u5b58(?:\([A-Z]\))?$")
AVIDEMUX_OVERWRITE_PROMPT_PATTERNS = (
    r"^Overwrite file .*\?$",
    r"^\u8986\u5199\u6587\u4ef6 .*\?$",
    r"^.*overwrite.*$",
    r"^.*replace.*$",
    r"^.*\u8986\u5199.*$",
    r"^.*\u8986\u76d6.*$",
    r"^.*\u590d\u5199.*$",
)
AVIDEMUX_PROGRESS_DIALOG_TITLE_PATTERNS = (
    r"^Encoding.*$",
    r"^Saving.*$",
    r"^\u6b63\u5728\u7f16\u7801.*$",
    r"^\u4fdd\u5b58\u4e2d.*$",
)
AVIDEMUX_PROGRESS_DIALOG_CONTENT_PATTERNS = (
    r"^Output File:$",
    r"^Keep dialog open when finished$",
    r"^Stage:$",
    r"^Remaining time:?$",
    r"^\u8f93\u51fa\u6587\u4ef6[:\uff1a]?$",
    r"^\u5b8c\u6210\u540e.*$",
    r"^\u9636\u6bb5[:\uff1a]?$",
    r"^\u5269\u4f59\u65f6\u95f4[:\uff1a]?$",
)
AVIDEMUX_MINIMIZE_TO_TRAY_PATTERNS = (
    r"^Minimi[sz]e to tray(?: icon)?$",
    r"^Send to tray$",
    r"^To tray$",
    r"^\u7f29\u5230\u5de5\u5177\u680f\u4e0a$",
    r"^\u6700\u5c0f\u5316\u5230\u6258\u76d8$",
)
AVIDEMUX_TRAY_BUTTON_MAX_WIDTH = 220
AVIDEMUX_TRAY_BUTTON_MAX_HEIGHT = 80
AVIDEMUX_TRAY_BUTTON_BOTTOM_REGION_RATIO = 0.35
AVIDEMUX_TRAY_BUTTON_LEFT_REGION_RATIO = 0.45

HAND_BRAKE_SELECT_FILE_PATTERNS = (
    r"^\u9009\u62e9\u8981\u626b\u63cf\u7684\u6587\u4ef6$",
    r"^Open a Single Video File$",
)
HAND_BRAKE_OPEN_SOURCE_PATTERNS = (r"^\u6253\u5f00\u6e90$", r"^Open Source$")
HAND_BRAKE_START_BUTTON_PATTERNS = (r"^\u5f00\u59cb\u7f16\u7801$", r"^Start Encode$")
HAND_BRAKE_SUMMARY_TAB_PATTERNS = (r"^\u6458\u8981$", r"^Summary$")
HAND_BRAKE_VIDEO_TAB_PATTERNS = (r"^\u89c6\u9891$", r"^Video$")
HAND_BRAKE_PRESET_LABEL_PATTERNS = (r"^Preset[s]?:?$", r"^\u9884\u8bbe[:\uff1a]?$")
HAND_BRAKE_DESTINATION_BROWSE_PATTERNS = (
    r"^Browse(?:\([A-Z]\))?$",
    r"^Browse(?:\.{3})?$",
    r"^Save As(?:\([A-Z]\))?$",
    r"^\u6d4f\u89c8(?:\([A-Z]\))?$",
    r"^\u53e6\u5b58\u4e3a(?:\([A-Z]\))?$",
)
HAND_BRAKE_VIDEO_ENCODER_LABEL_PATTERNS = (r"^Video Encoder:?$", r"^\u89c6\u9891\u7f16\u7801\u5668[:\uff1a]?$")
HAND_BRAKE_DIALOG_OPEN_PATTERNS = (r"^\u6253\u5f00(?:\([A-Z]\))?$", r"^Open(?:\([A-Z]\))?$")
HAND_BRAKE_DIALOG_CANCEL_PATTERNS = (r"^\u53d6\u6d88$", r"^Cancel$")
HAND_BRAKE_X264_PATTERNS = (r"\bx264\b", r"^H\.264 \(x264\).*$")
HAND_BRAKE_ACTIVE_STATUS_PATTERNS = (r"^(?:Encoding|\u7f16\u7801)\s*[:\uff1a].*$",)
HAND_BRAKE_READY_STATUS_PATTERNS = (r"^(?:Ready|\u51c6\u5907\u5c31\u7eea)$",)
HAND_BRAKE_UPDATE_DIALOG_PATTERNS = (r"^Check for Updates\?$", r"^Update Check\??$", r"^\u68c0\u67e5\u66f4\u65b0\?$")
HAND_BRAKE_UPDATE_PROMPT_PATTERNS = (
    r"^Would you like HandBrake to check for updates automatically\?$",
    r"^\u60a8\u60f3\u8ba9\s*HandBrake\s*\u81ea\u52a8\u68c0\u67e5\u66f4\u65b0\u5417\uff1f?$",
)
HAND_BRAKE_OVERWRITE_TEXT_PATTERNS = (
    r"overwrite",
    r"replace",
    r"already exists",
    r"\u8986\u76d6",
    r"\u66ff\u6362",
    r"\u5df2\u5b58\u5728",
)
HAND_BRAKE_YES_PATTERNS = (r"^Yes(?:\b|[(])", r"^\u662f(?:$|[(])")
HAND_BRAKE_UPDATE_NO_PATTERNS = (r"^No(?:\([A-Z]\))?$", r"^\u5426(?:\([A-Z]\))?$")
HAND_BRAKE_RECOVERY_DIALOG_PATTERNS = (r"^\u6709\u53ef\u6062\u590d\u7684\u961f\u5217$", r"^Recoverable Queue$")
HAND_BRAKE_RECOVERY_NO_PATTERNS = (r"^\u5426(?:\([A-Z]\))?$", r"^No(?:\([A-Z]\))?$")
HAND_BRAKE_UPDATE_DECLINE_KEY = "%n"
HAND_BRAKE_DIALOG_TIMEOUT_SECONDS = 8.0
HAND_BRAKE_DIALOG_SETTLE_SECONDS = 0.6
HAND_BRAKE_PRESET_SELECTION_TIMEOUT_SECONDS = 10.0
HAND_BRAKE_SOURCE_LOAD_TIMEOUT_SECONDS = 60.0
HAND_BRAKE_START_TIMEOUT_SECONDS = 20.0
HAND_BRAKE_EXPORT_TIMEOUT_SECONDS = 7200.0
HAND_BRAKE_EXPORT_POLL_SECONDS = 2.0
HAND_BRAKE_EXPORT_STABLE_ROUNDS = 4
HAND_BRAKE_SOURCE_DESTINATION_SUFFIXES = (".mp4", ".m4v", ".mkv", ".webm")
HAND_BRAKE_DESTINATION_EDIT_MIN_WIDTH = 160
HAND_BRAKE_SOURCE_NAME_MIN_TEXT_LENGTH = 4
HAND_BRAKE_TARGET_PRESET_TEXT = "HQ 2160P60 4K HEVC Surround"
HAND_BRAKE_PRESET_VALUE_HINTS = (
    "2160p",
    "1080p",
    "720p",
    "576p",
    "480p",
    "hq",
    "fast",
    "super hq",
    "hevc",
    "h.265",
    "h265",
    "av1",
    "h.264",
    "h264",
    "surround",
    "preset",
)
HAND_BRAKE_TITLE_SELECTOR_VALUE_RE = r"^\d+\s*\(.+\)$"

KDENLIVE_OPEN_DIALOG_PATTERNS = (
    r"^Kdenlive$",
    r"^Open(?: File)?$",
    r"^Add Clip or Folder(?:…|\.{3})?$",
    r"^Import.*$",
    r"^\u6253\u5f00(?:\u6587\u4ef6)?$",
    r"^\u5bfc\u5165.*$",
)
KDENLIVE_OPEN_CONFIRM_PATTERNS = (
    r"^\u786e\u5b9a(?:\(|$)",
    r"^\u6253\u5f00(?:\(|$)",
    r"^OK$",
    r"^Open$",
    r"^Import(?: Selection)?(?:\s*\(.+\))?$",
    r"^\u5bfc\u5165(?:.*)?$",
)
KDENLIVE_PROJECT_MENU_PATTERNS = (r"^\u9879\u76ee(?:\(P\))?$", r"^Project(?:\(P\))?$")
KDENLIVE_ADD_CLIP_PATTERNS = (
    r"^\u6dfb\u52a0\u526a\u8f91\u6216\u6587\u4ef6\u5939.*$",
    r"^Add Clip or Folder.*$",
)
KDENLIVE_EXPORT_MENU_PATTERNS = (r"^\u5bfc\u51fa.*$", r"^Render.*$", r"^Export.*$")
KDENLIVE_INSERT_TO_TIMELINE_PATTERNS = (
    r"^\u63d2\u5165\u526a\u8f91\u533a\u6bb5\u5230\u65f6\u95f4\u8f74.*$",
    r"^Insert.*Timeline.*$",
)
KDENLIVE_RENDER_DIALOG_TITLE_RE = r".*(?:Rendering|Render Project|Render|Export|\u6e32\u67d3|\u5bfc\u51fa).*"
KDENLIVE_RENDER_TO_FILE_PATTERNS = (r"^Render to File$", r"^\u6e32\u67d3\u5230\u6587\u4ef6$")
KDENLIVE_RENDER_WINDOW_PATTERNS = (
    r"^Rendering(?: - Kdenlive)?$",
    r"^\u6e32\u67d3(?:.*Kdenlive)?$",
)
KDENLIVE_JOB_QUEUE_TAB_PATTERNS = (r"^Job Queue$", r"^\u4efb\u52a1\u961f\u5217$")
KDENLIVE_ACTIVE_RENDER_PATTERNS = (
    r"^Remaining time\b.*$",
    r".*\(frame \d+ @ \d+(?:\.\d+)? fps\).*$",
    r"^\u5269\u4f59\u65f6\u95f4.*$",
)
KDENLIVE_FINISHED_RENDER_PATTERNS = (
    r"^Rendering finished\b.*$",
    r"^\u6e32\u67d3.*\u5b8c\u6210.*$",
)
KDENLIVE_ABORT_JOB_PATTERNS = (r"^Abort Job$", r"^\u4e2d\u6b62\u4efb\u52a1$")
KDENLIVE_START_JOB_PATTERNS = (r"^Start Job$", r"^\u542f\u52a8\u4efb\u52a1$")
KDENLIVE_CLEAN_UP_PATTERNS = (r"^Clean Up$", r"^\u6e05\u7406$")
KDENLIVE_CLOSE_BUTTON_PATTERNS = (r"^Close$", r"^\u5173\u95ed$")
KDENLIVE_WARNING_DIALOG_PATTERNS = (r"^Warning(?: - Kdenlive)?$", r"^\u8b66\u544a(?: - Kdenlive)?$")
KDENLIVE_SAVE_CHANGES_TEXT_PATTERNS = (
    r"^Save changes to document\?$",
    r"^\u4fdd\u5b58.*\u6587\u6863.*\?$",
)
KDENLIVE_OVERWRITE_TEXT_PATTERNS = (
    r"^Output file already exists\..*$",
    r"^Do you want to overwrite it\?$",
    r"^.*overwrite.*$",
    r"^.*\u8f93\u51fa\u6587\u4ef6.*\u5df2\u5b58\u5728.*$",
    r"^.*\u8986\u76d6.*\?$",
)
KDENLIVE_OVERWRITE_CONFIRM_PATTERNS = (
    r"^Overwrite(?:\s*\(.+\))?$",
    r"^\u8986\u76d6(?:\s*\(.+\))?$",
)
KDENLIVE_DONT_SAVE_PATTERNS = (
    r"^Don't Save(?:\s*\(.+\))?$",
    r"^Do Not Save(?:\s*\(.+\))?$",
    r"^Discard(?:\s*\(.+\))?$",
    r"^No(?:\s*\(.+\))?$",
    r"^\u4e0d\u4fdd\u5b58(?:\s*\(.+\))?$",
    r"^\u5426(?:\s*\(.+\))?$",
)
KDENLIVE_DONT_SAVE_DIRECT_KEYS = ("%d", "%n", "d", "n")
KDENLIVE_DIALOG_CANCEL_PATTERNS = (r"^Cancel(?:\([A-Z]\))?$", r"^\u53d6\u6d88(?:\([A-Z]\))?$")
KDENLIVE_PROFILE_SWITCH_TEXT_PATTERNS = (
    r"^Switch to clip .*\bprofile\b.*\?$",
    r"^\u5207\u6362.*\u914d\u7f6e.*\?$",
)
KDENLIVE_RECOVERY_DIALOG_PATTERNS = (
    r"^File Recovery(?: - Kdenlive)?$",
    r"^\u6587\u4ef6\u6062\u590d(?: - Kdenlive)?$",
)
KDENLIVE_RECOVERY_TEXT_PATTERNS = (
    r"^Auto-saved file exists\..*$",
    r"^Do you want to recover now\?$",
    r"^.*\u81ea\u52a8\u4fdd\u5b58.*$",
    r"^.*\u6062\u590d.*\?$",
)
KDENLIVE_DO_NOT_RECOVER_PATTERNS = (
    r"^Do not recover(?:\s*\(.+\))?$",
    r"^Do Not Recover(?:\s*\(.+\))?$",
    r"^\u4e0d\u8981\u6062\u590d(?:\s*\(.+\))?$",
    r"^\u4e0d\u6062\u590d(?:\s*\(.+\))?$",
)
KDENLIVE_OUTPUT_FILE_LABEL_PATTERNS = (r"^Output file$", r"^\u8f93\u51fa\u6587\u4ef6$")
KDENLIVE_RENDER_LENGTH_PATTERNS = (r"^Rendered File Length:.*$", r"^\u6e32\u67d3\u6587\u4ef6\u65f6\u957f:.*$")
KDENLIVE_CLIP_CONTROL_TYPES = ("TreeItem", "ListItem", "DataItem", "Text")
KDENLIVE_CONTROL_TIMEOUT_SECONDS = 6.0
KDENLIVE_DIALOG_TIMEOUT_SECONDS = 6.0
KDENLIVE_RENDER_DIALOG_TIMEOUT_SECONDS = 8.0
KDENLIVE_POST_ACTION_SLEEP_SECONDS = 0.5
KDENLIVE_IMPORT_SETTLE_SECONDS = 1.0
KDENLIVE_RENDER_POLL_SECONDS = 1.0
KDENLIVE_RENDER_STABLE_ROUNDS = 3
KDENLIVE_UI_IDLE_ROUNDS = 2
KDENLIVE_FALLBACK_COMPLETION_STABLE_ROUNDS = 12
KDENLIVE_FALLBACK_COMPLETION_IDLE_ROUNDS = 12
KDENLIVE_MIN_OUTPUT_BYTES = 1024

SHUTTER_ENCODER_UPDATE_DIALOG_PATTERNS = (
    r"^Available update.*$",
    r"^Update.*$",
    r"^\u53ef\u7528\u66f4\u65b0.*$",
    r"^\u66f4\u65b0.*$",
)
SHUTTER_ENCODER_UPDATE_PROMPT_PATTERNS = (r"^Do you want to download\??$", r"^\u4f60\u8981\u4e0b\u8f7d\u5417\??$")
SHUTTER_ENCODER_UPDATE_VERSION_PATTERNS = (r"^Version\s+\d+(?:\.\d+)?\s+-.*$",)
SHUTTER_ENCODER_UPDATE_NO_PATTERNS = (r"^No(?:\(|$)", r"^\u5426(?:\(|$)")
SHUTTER_ENCODER_BROWSE_PATTERNS = (r"^Browse(?:\.{3}|\u2026)?$", r"^\u6d4f\u89c8(?:\.{3}|\u2026)?$")
SHUTTER_ENCODER_BROWSE_CONTROL_TYPES = ("Button", "Text", "Pane", "Group", "Custom", "Hyperlink")
SHUTTER_ENCODER_FILE_SECTION_PATTERNS = (r"^\u9009\u62e9\u6587\u4ef6$", r"^Choose Files?$")
SHUTTER_ENCODER_FILE_COUNT_PATTERNS = (
    r"\b[1-9]\d*\s*(?:files?|file)\b",
    r"[1-9]\d*\s*\u6587\u4ef6",
)
SHUTTER_ENCODER_DROP_PLACEHOLDER_PATTERNS = (r".*drag.*here.*", r"^\u6587\u4ef6\u62d6\u653e\u5230\u8fd9\u91cc$")
SHUTTER_ENCODER_CLEAR_PATTERNS = (r"^Clear$", r"^\u6e05\u9664$")
SHUTTER_ENCODER_IMPORTED_PATH_PATTERNS = (
    r"[A-Za-z]:\\",
    r"\\Users\\",
    r"\.mp4\b",
    r"\.mov\b",
    r"\.mkv\b",
)
SHUTTER_ENCODER_FUNCTION_PICKER_PATTERNS = (
    r"^Choose function$",
    r"^Select function$",
    r"^Function$",
    r"^\u9009\u62e9\u529f\u80fd$",
)
SHUTTER_ENCODER_TARGET_FUNCTION_TEXT = "H.265"
SHUTTER_ENCODER_TARGET_FUNCTION_PATTERNS = (r"^H\.265$",)
SHUTTER_ENCODER_START_BUTTON_PATTERNS = (
    r"^Start function$",
    r"^Start Function$",
    r"^Start$",
    r"^\u542f\u52a8\u529f\u80fd$",
    r"^\u542f\u52a8$",
)
SHUTTER_ENCODER_PROGRESS_COMPLETE_PATTERNS = (r"progress.*completed", r"^\u8fdb\u5ea6\u5df2\u5b8c\u6210$")
SHUTTER_ENCODER_PROGRESS_100_PATTERNS = (r"^100(?:[\.,]0+)?%$",)
SHUTTER_ENCODER_SOURCE_CONTROL_TYPES = ("Text", "ListItem", "DataItem", "TreeItem", "Edit", "ComboBox")
SHUTTER_ENCODER_DIALOG_TIMEOUT_SECONDS = 6.0
SHUTTER_ENCODER_CONTROL_TIMEOUT_SECONDS = 6.0
SHUTTER_ENCODER_FUNCTION_SELECTION_TIMEOUT_SECONDS = 8.0
SHUTTER_ENCODER_PROGRESS_TIMEOUT_SECONDS = 7200.0
SHUTTER_ENCODER_PROGRESS_POLL_SECONDS = 1.0
SHUTTER_ENCODER_PROGRESS_IDLE_ROUNDS = 2
SHUTTER_ENCODER_OUTPUT_STABLE_COMPLETION_ROUNDS = 3
SHUTTER_ENCODER_SETTLE_SECONDS = 0.5
SHUTTER_ENCODER_CLOSE_TIMEOUT_SECONDS = 5.0
SHUTTER_ENCODER_OUTPUT_SUFFIXES = (".mp4", ".mov", ".mkv", ".m4v")
SHUTTER_ENCODER_UPDATE_TIMEOUT_SECONDS = 3.0
SHUTTER_ENCODER_UPDATE_DIRECT_KEY = "{ESC}"
SHUTTER_ENCODER_UPDATE_UIA_PROBE_SECONDS = 0.8
SHUTTER_ENCODER_BROWSE_CLICK_X_RATIO = 0.22
SHUTTER_ENCODER_BROWSE_CLICK_Y_RATIO = 0.11
SHUTTER_ENCODER_BROWSE_BUTTON_DESIGN_X = 74
SHUTTER_ENCODER_BROWSE_BUTTON_DESIGN_Y = 62
SHUTTER_ENCODER_FUNCTION_EDITOR_CLICK_X_RATIO = 0.31
SHUTTER_ENCODER_FUNCTION_EDITOR_CLICK_Y_RATIO = 0.51
SHUTTER_ENCODER_FUNCTION_PICKER_CLICK_X_RATIO = 0.53
SHUTTER_ENCODER_FUNCTION_PICKER_CLICK_Y_RATIO = 0.51
SHUTTER_ENCODER_START_BUTTON_CLICK_X_RATIO = 0.31
SHUTTER_ENCODER_START_BUTTON_CLICK_Y_RATIO = 0.545
SHUTTER_ENCODER_QUIT_BUTTON_RIGHT_OFFSET = 12
SHUTTER_ENCODER_QUIT_BUTTON_TOP_OFFSET = 12
SHUTTER_ENCODER_FUNCTION_EDITOR_DESIGN_X = 102
SHUTTER_ENCODER_FUNCTION_EDITOR_DESIGN_Y = 372
SHUTTER_ENCODER_START_BUTTON_DESIGN_X = 102
SHUTTER_ENCODER_START_BUTTON_DESIGN_Y = 398
SHUTTER_ENCODER_FUNCTION_PICKER_DESIGN_HEIGHT = 22
SHUTTER_ENCODER_START_FROM_PICKER_CENTER_Y_OFFSET = (
    SHUTTER_ENCODER_START_BUTTON_DESIGN_Y - SHUTTER_ENCODER_FUNCTION_EDITOR_DESIGN_Y
)
SHUTTER_ENCODER_START_RETRY_SECONDS = 3.0
SHUTTER_ENCODER_IMPORT_TIMEOUT_SECONDS = 12.0


@dataclass(frozen=True)
class OperationProfile:
    software: str
    main_window_title_re: str
    output_suffix: str
    import_hotkey: str = ""
    import_clicks: tuple[tuple[str, ...], ...] = ()
    export_pre_save_clicks: tuple[tuple[str, ...], ...] = ()
    export_post_save_clicks: tuple[tuple[str, ...], ...] = ()
    save_hotkey: str = ""
    import_wait_seconds: float = 2.0
    export_wait_seconds: float = 1.0


OPERATION_PROFILES = {
    "shotcut": OperationProfile(
        software="shotcut",
        main_window_title_re=SOFTWARE_SPECS["shotcut"].main_window_title_re,
        output_suffix=".mp4",
        import_hotkey="^o",
        export_pre_save_clicks=((r"^Export$",), (r"^Export File$",)),
    ),
    "kdenlive": OperationProfile(
        software="kdenlive",
        main_window_title_re=SOFTWARE_SPECS["kdenlive"].main_window_title_re,
        output_suffix=".mp4",
        import_hotkey="^o",
        export_pre_save_clicks=((r"^Render$",),),
        export_post_save_clicks=((r"^Render to File$", r"^Render$"),),
    ),
    "shutter_encoder": OperationProfile(
        software="shutter_encoder",
        main_window_title_re=SOFTWARE_SPECS["shutter_encoder"].main_window_title_re,
        output_suffix=".mp4",
        import_clicks=((r"^Browse$",),),
        export_post_save_clicks=((r"^Start function$", r"^Start Function$", r"^Start$"),),
    ),
    "avidemux": OperationProfile(
        software="avidemux",
        main_window_title_re=SOFTWARE_SPECS["avidemux"].main_window_title_re,
        output_suffix=".mkv",
        import_hotkey="^o",
        save_hotkey="^s",
    ),
    "handbrake": OperationProfile(
        software="handbrake",
        main_window_title_re=SOFTWARE_SPECS["handbrake"].main_window_title_re,
        output_suffix=".mp4",
        import_clicks=((r"^Open Source$",),),
        export_pre_save_clicks=((r"^Browse$", r"^Save As$"),),
        export_post_save_clicks=((r"^Start Encode$", r"^Start$"),),
    ),
}


class SoftwareOperator:
    def __init__(self, profile: OperationProfile) -> None:
        self.profile = profile

    def _connect_main_window(self, timeout: float = 30.0):
        return connect_software_window(self.profile.software, timeout=timeout)

    def perform(self, input_video_path: Path, output_video_path: Path) -> None:
        logger.info("Starting %s operation. input=%s output=%s", self.profile.software, input_video_path, output_video_path)
        logger.info("Validating %s input and output paths.", self.profile.software)
        assert input_video_path.exists(), f"Input video does not exist: {input_video_path}"
        assert output_video_path.parent.exists(), f"Output directory does not exist: {output_video_path.parent}"
        assert output_video_path.suffix.lower() == self.profile.output_suffix, (
            f"Expected output suffix {self.profile.output_suffix}, got {output_video_path.suffix}"
        )

        logger.info("Connecting to %s main window.", self.profile.software)
        window = self._connect_main_window(timeout=30.0)
        bring_window_to_front(window, keep_topmost=False)
        self._import_input(window, input_video_path)
        self._start_export(window, output_video_path)

    def close(self) -> None:
        logger.info("Closing %s.", self.profile.software)
        try:
            window = self._connect_main_window(timeout=2.0)
        except UiAutomationError:
            logger.info("%s window is already closed.", self.profile.software)
            return
        bring_window_to_front(window, keep_topmost=False)
        try:
            window.close()
        except Exception as exc:
            logger.info("%s window.close() raised %s. Falling back to WM_CLOSE.", self.profile.software, exc)
            request_window_close(window)
        time.sleep(1.0)
        dismiss_close_prompts(owner_window=window)

    def _minimize_window_during_wait(self, window, *, phase: str) -> None:
        try:
            top_level = window.top_level_parent()
        except Exception:
            top_level = window
        logger.info("Minimizing %s during the %s wait.", self.profile.software, phase)
        try:
            minimize_window(top_level)
        except Exception as exc:
            logger.info(
                "Could not minimize %s during the %s wait: %s",
                self.profile.software,
                phase,
                exc,
            )

    def _begin_background_wait(
        self,
        window,
        *,
        phase: str,
        start_waiter=None,
        start_description: str | None = None,
    ) -> None:
        self._minimize_window_during_wait(window, phase=phase)
        if start_waiter is None:
            return
        if start_description:
            logger.info("Confirming that %s has started after minimizing.", start_description)
        start_waiter()

    def _import_input(self, window, input_video_path: Path) -> None:
        logger.info("Importing input for %s.", self.profile.software)
        for patterns in self.profile.import_clicks:
            click_text_control(window, patterns, control_types=("Button", "MenuItem", "Hyperlink"))
        if self.profile.import_hotkey:
            bring_window_to_front(window, keep_topmost=False)
            send_hotkey(self.profile.import_hotkey)
        fill_file_dialog(input_video_path, dialog_patterns=OPEN_DIALOG_PATTERNS, must_exist=True)
        time.sleep(self.profile.import_wait_seconds)

    def _start_export(self, window, output_video_path: Path) -> None:
        logger.info("Starting export flow for %s.", self.profile.software)
        for patterns in self.profile.export_pre_save_clicks:
            click_text_control(window, patterns, control_types=("Button", "MenuItem", "Hyperlink", "TabItem"))
            time.sleep(self.profile.export_wait_seconds)
        if self.profile.save_hotkey:
            bring_window_to_front(window, keep_topmost=False)
            send_hotkey(self.profile.save_hotkey)
        if self.profile.save_hotkey or self.profile.export_pre_save_clicks:
            fill_file_dialog(output_video_path, dialog_patterns=SAVE_DIALOG_PATTERNS, must_exist=False)
        for patterns in self.profile.export_post_save_clicks:
            click_text_control(window, patterns, control_types=("Button", "MenuItem", "Hyperlink"))
            time.sleep(self.profile.export_wait_seconds)


def _format_elapsed_hms(seconds: float) -> str:
    total_seconds = max(0, int(seconds))
    hours, remainder = divmod(total_seconds, 3600)
    minutes, secs = divmod(remainder, 60)
    return f"{hours:02d}:{minutes:02d}:{secs:02d}"


class _WaitProgressLogger:
    def __init__(self, label: str, *, interval_seconds: float = WAIT_PROGRESS_LOG_INTERVAL_SECONDS) -> None:
        self.label = label
        self.interval_seconds = max(1.0, interval_seconds)
        self.started_at = time.monotonic()
        self.next_log_at = self.started_at + self.interval_seconds

    def elapsed_seconds(self) -> float:
        return max(0.0, time.monotonic() - self.started_at)

    def elapsed_text(self) -> str:
        return _format_elapsed_hms(self.elapsed_seconds())

    def maybe_log(self, *, detail: str = "") -> None:
        now = time.monotonic()
        if now < self.next_log_at:
            return
        suffix = f" {detail}" if detail else ""
        logger.info("%s in progress. elapsed=%s%s", self.label, self.elapsed_text(), suffix)
        self.next_log_at = now + self.interval_seconds


class ShotcutOperator(SoftwareOperator):
    def _wrapper_text(self, wrapper) -> str:
        try:
            return (wrapper.window_text() or "").strip()
        except Exception:
            return ""

    def _process_id(self, window) -> int | None:
        try:
            process_id = getattr(getattr(window, "element_info", None), "process_id", None)
        except Exception:
            return None
        return int(process_id) if process_id else None

    def _iter_wrapper_tree(self, root):
        yield root
        for child in root.descendants():
            yield child

    def _matches_patterns(self, text: str, patterns: tuple[str, ...]) -> bool:
        return any(re.search(pattern, text, re.IGNORECASE) for pattern in patterns)

    def _iter_process_top_level_windows(self, process_id: int | None):
        if process_id is None:
            return []
        from pywinauto import Desktop

        matches = []
        for window in Desktop(backend="uia").windows():
            if self._process_id(window) != process_id:
                continue
            try:
                rect = window.rectangle()
                area = max(0, rect.right - rect.left) * max(0, rect.bottom - rect.top)
            except Exception:
                area = 0
            matches.append((-area, self._wrapper_text(window).casefold(), window))
        matches.sort(key=lambda item: (item[0], item[1]))
        return [item[2] for item in matches]

    def _dialog_matches(self, dialog, *, text_patterns: tuple[str, ...]) -> bool:
        for wrapper in self._iter_wrapper_tree(dialog):
            text = self._wrapper_text(wrapper)
            if text and self._matches_patterns(text, text_patterns):
                return True
        return False

    def _dismiss_save_changes_dialog_if_present(self, owner_window, *, timeout: float = 3.0) -> bool:
        process_id = self._process_id(owner_window)
        owner_handle = getattr(owner_window, "handle", None)
        deadline = time.monotonic() + timeout
        dismissed = False
        while time.monotonic() < deadline:
            dialog = None
            for candidate in self._iter_process_top_level_windows(process_id):
                if owner_handle is not None and getattr(candidate, "handle", None) == owner_handle:
                    continue
                if not self._dialog_matches(candidate, text_patterns=SHOTCUT_SAVE_CHANGES_TEXT_PATTERNS):
                    continue
                dialog = candidate
                break
            if dialog is None:
                return dismissed
            try:
                bring_window_to_front(dialog, keep_topmost=False)
            except Exception:
                pass
            try:
                button = find_text_control(dialog, SHOTCUT_DONT_SAVE_PATTERNS, control_types=("Button",))
            except Exception:
                logger.info(
                    "Shotcut save-changes dialog was found, but a discard-style button was not directly exposed. "
                    "Trying keyboard dismiss shortcuts before generic close-prompt dismissal."
                )
                for key in SHOTCUT_DONT_SAVE_DIRECT_KEYS:
                    try:
                        logger.info("Sending Shotcut discard shortcut: %s", key)
                        send_hotkey(key)
                    except Exception:
                        pass
                dismiss_close_prompts(timeout=0.8, owner_window=dialog)
                time.sleep(0.2)
                dismissed = True
                continue
            logger.info("Clicking the Shotcut discard-style button: %s", self._wrapper_text(button) or "<untitled>")
            try:
                click_control(button, post_click_sleep=0.5)
            except Exception as exc:
                logger.info(
                    "Clicking the Shotcut discard-style button failed. Falling back to keyboard shortcuts. details=%s",
                    exc,
                )
                for key in SHOTCUT_DONT_SAVE_DIRECT_KEYS:
                    try:
                        logger.info("Sending Shotcut discard shortcut: %s", key)
                        send_hotkey(key)
                    except Exception:
                        pass
                dismiss_close_prompts(timeout=0.8, owner_window=dialog)
            dismissed = True
        return dismissed

    def _dismiss_recovery_dialog_if_present(self, owner_window, *, timeout: float = 1.5) -> bool:
        process_id = self._process_id(owner_window)
        deadline = time.monotonic() + timeout
        dismissed = False
        while time.monotonic() < deadline:
            dialog = None
            for candidate in self._iter_process_top_level_windows(process_id):
                if not self._dialog_matches(candidate, text_patterns=SHOTCUT_RECOVERY_TEXT_PATTERNS):
                    continue
                dialog = candidate
                break
            if dialog is None:
                if self._dialog_matches(owner_window, text_patterns=SHOTCUT_RECOVERY_TEXT_PATTERNS):
                    dialog = owner_window
                else:
                    return dismissed
            try:
                bring_window_to_front(dialog, keep_topmost=False)
            except Exception:
                pass
            try:
                button = find_text_control(dialog, SHOTCUT_RECOVERY_DISMISS_PATTERNS, control_types=("Button",))
                logger.info("Clicking the Shotcut recovery-dismiss button: %s", self._wrapper_text(button) or "<untitled>")
                click_control(button, post_click_sleep=0.5)
            except Exception:
                logger.info("Shotcut recovery dialog was found, but the dismiss button was not directly exposed. Sending 'N'.")
                try:
                    send_hotkey(SHOTCUT_RECOVERY_DISMISS_KEY)
                except Exception:
                    pass
                time.sleep(0.2)
            dismissed = True
        return dismissed

    def perform(self, input_video_path: Path, output_video_path: Path) -> None:
        logger.info("Shotcut flow started. input=%s output=%s", input_video_path, output_video_path)
        logger.info("Validating Shotcut input and output paths.")
        assert input_video_path.exists(), f"Input video does not exist: {input_video_path}"
        assert output_video_path.parent.exists(), f"Output directory does not exist: {output_video_path.parent}"
        assert output_video_path.suffix.lower() == ".mp4", f"Shotcut export target must be an mp4 file: {output_video_path}"

        logger.info("Connecting to Shotcut main window.")
        window = self._connect_main_window(timeout=max(30.0, SOFTWARE_SPECS["shotcut"].startup_timeout_seconds))
        if self._dismiss_recovery_dialog_if_present(window, timeout=2.0):
            logger.info("Reconnecting to Shotcut after dismissing the autosave recovery dialog.")
            window = self._connect_main_window(timeout=SHOTCUT_RECONNECT_TIMEOUT_SECONDS)
        bring_window_to_front(window, keep_topmost=False)
        window = self._open_input_clip(window, input_video_path)
        logger.info("Appending the imported Shotcut clip to the timeline before export.")
        self._append_selected_clip_to_timeline(window)
        self._export_current_timeline(window, output_video_path)
        self.close()

    def close(self) -> None:
        logger.info("Closing Shotcut.")
        try:
            window = self._connect_main_window(timeout=2.0)
        except UiAutomationError:
            logger.info("Shotcut window is already closed.")
            return
        bring_window_to_front(window, keep_topmost=False)
        self._dismiss_save_changes_dialog_if_present(window, timeout=0.5)
        try:
            logger.info("Requesting Shotcut shutdown with hotkey: %s", SHOTCUT_CLOSE_HOTKEY)
            send_hotkey(SHOTCUT_CLOSE_HOTKEY)
        except Exception as exc:
            logger.info("Shotcut close hotkey raised %s. Falling back to WM_CLOSE.", exc)
        self._dismiss_save_changes_dialog_if_present(window, timeout=1.5)
        try:
            request_window_close(window)
        except Exception as exc:
            logger.info("Shotcut WM_CLOSE fallback raised %s.", exc)
        self._dismiss_save_changes_dialog_if_present(window, timeout=3.0)
        dismiss_close_prompts(timeout=1.5, owner_window=window)

    def _find_child_by_class_name(self, window, class_name: str):
        logger.info("Locating Shotcut child control by class name: %s", class_name)
        direct_child_classes: list[str] = []
        for child in window.children():
            child_class = (getattr(child.element_info, "class_name", "") or "").strip()
            if child_class:
                direct_child_classes.append(child_class)
            if child_class == class_name:
                return child

        logger.info(
            "Shotcut control '%s' was not found among direct children. Falling back to descendant search.",
            class_name,
        )
        descendant_class_sample: list[str] = []
        seen_descendant_classes: set[str] = set()
        for child in window.descendants():
            child_class = (getattr(child.element_info, "class_name", "") or "").strip()
            if child_class and child_class not in seen_descendant_classes and len(descendant_class_sample) < 12:
                seen_descendant_classes.add(child_class)
                descendant_class_sample.append(child_class)
            if child_class == class_name:
                logger.info("Resolved Shotcut control '%s' from descendants.", class_name)
                return child
        raise AssertionError(
            f"Shotcut control '{class_name}' was not found. "
            f"Direct child classes: {sorted(set(direct_child_classes))[:12]}. "
            f"Descendant class sample: {descendant_class_sample}."
        )

    def _find_child_by_class_name_if_present(self, window, class_name: str):
        try:
            return self._find_child_by_class_name(window, class_name)
        except AssertionError as exc:
            logger.info(
                "Shotcut control '%s' is not exposed as a dedicated class in this layout. "
                "Falling back to broader window search. details=%s",
                class_name,
                exc,
            )
            return None

    def _resolve_jobs_root(self, window):
        jobs_dock = self._find_child_by_class_name_if_present(window, "JobsDock")
        if jobs_dock is not None:
            return jobs_dock
        logger.info("Using the Shotcut main window as the Jobs search root.")
        return window

    def _append_selected_clip_to_timeline(self, window) -> None:
        logger.info("Appending the selected Shotcut clip to the timeline with shortcut '%s'.", SHOTCUT_APPEND_TO_TIMELINE_SHORTCUT.upper())
        bring_window_to_front(window, keep_topmost=False)
        send_hotkey(SHOTCUT_APPEND_TO_TIMELINE_SHORTCUT)

    def _open_input_clip(self, window, input_video_path: Path):
        if self._dismiss_save_changes_dialog_if_present(window, timeout=0.8):
            logger.info("Shotcut save-changes dialog was blocking the next import. Reconnecting to the main window.")
            window = self._connect_main_window(timeout=SHOTCUT_RECONNECT_TIMEOUT_SECONDS)
            bring_window_to_front(window, keep_topmost=False)
        logger.info("Locating Shotcut toolbar.")
        toolbar = self._find_child_by_class_name(window, "QToolBar")
        logger.info("Opening Shotcut input dialog.")
        self._open_input_dialog(window, toolbar)
        logger.info("Filling Shotcut input dialog with selected video.")
        fill_file_dialog(input_video_path, dialog_patterns=OPEN_DIALOG_PATTERNS, must_exist=True)
        logger.info("Reconnecting to Shotcut after input selection.")
        window = self._connect_main_window(timeout=SHOTCUT_RECONNECT_TIMEOUT_SECONDS)
        bring_window_to_front(window, keep_topmost=False)
        return window

    def _export_current_timeline(self, window, output_video_path: Path):
        logger.info("Locating Shotcut controls for export.")
        bring_window_to_front(window, keep_topmost=False)
        toolbar = self._find_child_by_class_name(window, "QToolBar")
        encode_dock = self._find_child_by_class_name(window, "EncodeDock")
        logger.info("Waiting for Shotcut export controls to become ready.")
        self._wait_for_export_ready(encode_dock)
        logger.info("Opening Shotcut output pane.")
        self._show_output_pane(toolbar, encode_dock)
        logger.info("Opening Shotcut export save dialog.")
        self._export_clip(encode_dock, output_video_path)
        self._begin_background_wait(
            window,
            phase="export",
        )
        logger.info("Waiting for Shotcut export completion.")
        self._wait_for_export_completion(output_video_path)
        return window

    def _wait_for_open_file_dialog(self, *, timeout: float) -> None:
        logger.info("Waiting for Shotcut open-file dialog.")
        dialog = wait_for_file_dialog(dialog_patterns=OPEN_DIALOG_PATTERNS, timeout=timeout)
        assert dialog is not None, "Shotcut did not open the input file dialog."

    def _open_input_dialog(self, window, toolbar) -> None:
        last_error: Exception | None = None
        current_window = window
        current_toolbar = toolbar
        for attempt_index in range(3):
            if self._dismiss_save_changes_dialog_if_present(current_window, timeout=0.4):
                logger.info(
                    "Shotcut save-changes dialog interrupted the next import. Reconnecting before retrying the open-file flow."
                )
                try:
                    current_window = self._connect_main_window(timeout=SHOTCUT_RECONNECT_TIMEOUT_SECONDS)
                    current_toolbar = self._find_child_by_class_name(current_window, "QToolBar")
                except Exception:
                    pass
            self._dismiss_recovery_dialog_if_present(current_window, timeout=0.4)
            logger.info("Waiting for Shotcut 'Open File' button.")
            open_button = wait_for_text_control(
                current_toolbar,
                SHOTCUT_OPEN_BUTTON_PATTERNS,
                control_types=("Button",),
                timeout=SHOTCUT_CONTROL_TIMEOUT_SECONDS,
            )
            assert open_button.is_visible(), "Shotcut toolbar 'Open File' button is not visible."
            logger.info("Clicking Shotcut 'Open File' button.")
            open_button.click_input()
            time.sleep(0.5)
            try:
                self._wait_for_open_file_dialog(timeout=SHOTCUT_DIALOG_TIMEOUT_SECONDS)
                return
            except Exception as exc:
                last_error = exc
                save_changes_dismissed = self._dismiss_save_changes_dialog_if_present(current_window, timeout=0.8)
                if save_changes_dismissed:
                    logger.info("Shotcut save-changes dialog interrupted the open-file flow. Retrying the import dialog.")
                recovery_dismissed = self._dismiss_recovery_dialog_if_present(current_window, timeout=0.8)
                if recovery_dismissed:
                    logger.info("Shotcut autosave recovery dialog interrupted the open-file flow. Retrying the import dialog.")
                logger.info("Shotcut open-file button path did not open the dialog. Trying Ctrl+O fallback. details=%s", exc)
                if save_changes_dismissed:
                    try:
                        current_window = self._connect_main_window(timeout=SHOTCUT_RECONNECT_TIMEOUT_SECONDS)
                        current_toolbar = self._find_child_by_class_name(current_window, "QToolBar")
                    except Exception:
                        pass
                try:
                    bring_window_to_front(current_window, keep_topmost=False)
                except Exception:
                    pass
                send_hotkey("^o")
                time.sleep(0.5)
                try:
                    self._wait_for_open_file_dialog(timeout=SHOTCUT_DIALOG_TIMEOUT_SECONDS)
                    return
                except Exception as hotkey_exc:
                    last_error = hotkey_exc
                    save_changes_dismissed = self._dismiss_save_changes_dialog_if_present(current_window, timeout=0.8)
                    if save_changes_dismissed:
                        logger.info(
                            "Shotcut save-changes dialog also interrupted the Ctrl+O fallback. Refreshing the main window before retry."
                        )
                    self._dismiss_recovery_dialog_if_present(current_window, timeout=0.8)
                    logger.info("Shotcut Ctrl+O fallback did not open the dialog. Refreshing the main window before retry. details=%s", hotkey_exc)
                    try:
                        current_window = self._connect_main_window(timeout=SHOTCUT_RECONNECT_TIMEOUT_SECONDS)
                        current_toolbar = self._find_child_by_class_name(current_window, "QToolBar")
                    except Exception:
                        pass
                    continue
        raise AssertionError("Shotcut did not open the input file dialog after button, recovery, and Ctrl+O retries.") from last_error

    def _wait_for_export_ready(self, encode_dock) -> None:
        logger.info("Waiting for Shotcut export button control to appear.")
        export_button = wait_for_text_control(
            encode_dock,
            r"^Export Video/Audio$",
            control_types=("Button",),
            timeout=SHOTCUT_CONTROL_TIMEOUT_SECONDS,
        )
        assert export_button.is_visible(), "Shotcut export button is not visible after import."

        logger.info(
            "Waiting up to %.1fs for Shotcut export button to become enabled.",
            SHOTCUT_EXPORT_READY_TIMEOUT_SECONDS,
        )
        deadline = time.monotonic() + SHOTCUT_EXPORT_READY_TIMEOUT_SECONDS
        while time.monotonic() < deadline:
            if export_button.is_enabled():
                logger.info("Shotcut export button is enabled.")
                return
            time.sleep(0.5)
        raise AssertionError("Shotcut export button did not become enabled after selecting the input video.")

    def _show_output_pane(self, toolbar, encode_dock) -> None:
        logger.info("Waiting for Shotcut 'Export' toolbar button.")
        output_button = wait_for_text_control(
            toolbar,
            SHOTCUT_OUTPUT_BUTTON_PATTERNS,
            control_types=("Button",),
            timeout=SHOTCUT_CONTROL_TIMEOUT_SECONDS,
        )
        assert output_button.is_visible(), "Shotcut toolbar 'Export' button is not visible."
        logger.info("Clicking Shotcut 'Export' toolbar button.")
        output_button.click_input()
        time.sleep(0.5)
        logger.info("Confirming that the Shotcut export pane is visible.")
        export_button = wait_for_text_control(
            encode_dock,
            r"^Export Video/Audio$",
            control_types=("Button",),
            timeout=SHOTCUT_CONTROL_TIMEOUT_SECONDS,
        )
        assert export_button.is_visible(), "Shotcut output pane did not expose 'Export Video/Audio'."

    def _export_clip(self, encode_dock, output_video_path: Path) -> None:
        dialog_error: UiAutomationError | None = None
        logger.info(
            "Trying to open Shotcut export save dialog. retries=%d output=%s",
            SHOTCUT_EXPORT_BUTTON_RETRY_COUNT,
            output_video_path,
        )
        for attempt_index in range(SHOTCUT_EXPORT_BUTTON_RETRY_COUNT):
            logger.info(
                "Shotcut export save dialog attempt %d/%d.",
                attempt_index + 1,
                SHOTCUT_EXPORT_BUTTON_RETRY_COUNT,
            )
            export_button = wait_for_text_control(
                encode_dock,
                r"^Export Video/Audio$",
                control_types=("Button",),
                timeout=SHOTCUT_CONTROL_TIMEOUT_SECONDS,
            )
            assert export_button.is_enabled(), "'Export Video/Audio' is disabled."
            logger.info("Clicking Shotcut 'Export Video/Audio' button.")
            export_button.click_input()
            time.sleep(SHOTCUT_EXPORT_BUTTON_POST_CLICK_SECONDS)
            try:
                logger.info("Waiting for Shotcut export save dialog and filling output path.")
                fill_file_dialog(
                    output_video_path,
                    dialog_patterns=SHOTCUT_SAVE_DIALOG_PATTERNS,
                    confirm_patterns=SHOTCUT_SAVE_CONFIRM_PATTERNS,
                    timeout=SHOTCUT_DIALOG_TIMEOUT_SECONDS,
                    must_exist=False,
                    wait_for_dialog_to_close=False,
                    overwrite_confirmation_timeout=SHOTCUT_SAVE_DIALOG_OVERWRITE_TIMEOUT_SECONDS,
                )
                logger.info("Shotcut export save dialog completed.")
                return
            except UiAutomationError as exc:
                dialog_error = exc
                logger.info("Shotcut export save dialog attempt failed: %s", exc)
                time.sleep(SHOTCUT_EXPORT_BUTTON_POST_CLICK_SECONDS)
        raise AssertionError("Shotcut did not open the export save dialog.") from dialog_error

    def _assert_task_queued_with_recovery(self, window, toolbar, output_video_path: Path):
        try:
            logger.info("Trying the current Shotcut window for Jobs pane inspection.")
            jobs_root = self._assert_task_queued(window, toolbar, output_video_path)
            return window, jobs_root
        except (AssertionError, UiAutomationError) as exc:
            logger.info(
                "Current Shotcut window could not confirm the queued export job. "
                "Reconnecting before retrying. details=%s",
                exc,
            )

        logger.info("Reconnecting to Shotcut to inspect Jobs pane.")
        window = self._connect_main_window(timeout=SHOTCUT_RECONNECT_TIMEOUT_SECONDS)
        bring_window_to_front(window, keep_topmost=False)
        toolbar = self._find_child_by_class_name(window, "QToolBar")
        jobs_root = self._assert_task_queued(window, toolbar, output_video_path)
        return window, jobs_root

    def _assert_task_queued(self, window, toolbar, output_video_path: Path):
        logger.info("Waiting for Shotcut 'Jobs' toolbar button.")
        tasks_button = wait_for_text_control(
            toolbar,
            SHOTCUT_TASKS_BUTTON_PATTERNS,
            control_types=("Button",),
            timeout=SHOTCUT_CONTROL_TIMEOUT_SECONDS,
        )
        assert tasks_button.is_visible(), "Shotcut toolbar 'Jobs' button is not visible."
        logger.info("Clicking Shotcut 'Jobs' toolbar button.")
        tasks_button.click_input()
        time.sleep(0.5)
        jobs_root = self._resolve_jobs_root(window)

        logger.info("Waiting for Shotcut queue controls to appear.")
        pause_button = wait_for_text_control(
            jobs_root,
            SHOTCUT_PAUSE_QUEUE_PATTERNS,
            control_types=("CheckBox", "Button"),
            timeout=SHOTCUT_JOBS_TIMEOUT_SECONDS,
        )
        assert pause_button.is_visible(), "Shotcut jobs pane did not show 'Pause Queue' after export started."

        logger.info("Waiting for Shotcut job row to appear: %s", output_video_path)
        task_item = wait_for_text_control(
            jobs_root,
            re.escape(str(output_video_path)),
            control_types=("TreeItem", "Text"),
            timeout=SHOTCUT_JOBS_TIMEOUT_SECONDS,
        )
        assert task_item.is_visible(), f"Shotcut jobs pane did not queue '{output_video_path}'."
        return jobs_root

    def _wait_for_export_output_start(self, output_video_path: Path, *, timeout: float = 20.0) -> None:
        logger.info("Waiting for Shotcut export output to start writing.")
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            try:
                window = self._connect_main_window(timeout=0.2)
            except Exception:
                window = None
            if window is not None:
                self._dismiss_recovery_dialog_if_present(window, timeout=0.2)
            if output_video_path.exists() and output_video_path.stat().st_size > 0:
                logger.info("Shotcut export output has started: %s", output_video_path)
                return
            time.sleep(0.2)
        raise AssertionError(f"Shotcut export output did not start within timeout: {output_video_path}")

    def _wait_for_export_completion(self, output_video_path: Path, jobs_dock=None) -> None:
        logger.info("Validating Shotcut output directory before waiting for completion.")
        assert output_video_path.parent.exists(), f"Output directory disappeared: {output_video_path.parent}"

        stable_rounds = 0
        last_size = -1
        deadline = time.monotonic() + 7200.0
        progress_logger = _WaitProgressLogger("Shotcut export")
        logger.info("Monitoring Shotcut export output until file size stabilizes.")
        while time.monotonic() < deadline:
            try:
                window = self._connect_main_window(timeout=0.2)
            except Exception:
                window = None
            if window is not None:
                self._dismiss_recovery_dialog_if_present(window, timeout=0.2)
            task_item_present = False
            time_item_present = False
            if jobs_dock is not None:
                task_item_present = control_exists(
                    jobs_dock,
                    re.escape(str(output_video_path)),
                    control_types=("TreeItem", "Text"),
                )
                time_item_present = control_exists(
                    jobs_dock,
                    SHOTCUT_TASK_TIME_PATTERN,
                    control_types=("TreeItem",),
                )
            if output_video_path.exists():
                current_size = output_video_path.stat().st_size
                if current_size > 0 and current_size == last_size:
                    stable_rounds += 1
                else:
                    stable_rounds = 0
                last_size = current_size
                progress_logger.maybe_log(
                    detail=(
                        f"output={output_video_path.name} size_bytes={current_size} "
                        f"stable_rounds={stable_rounds}"
                    )
                )
                if task_item_present and time_item_present and stable_rounds >= 3:
                    logger.info("Shotcut export completion conditions satisfied. elapsed=%s", progress_logger.elapsed_text())
                    break
                if jobs_dock is None and stable_rounds >= 3 and current_size > 0:
                    logger.info(
                        "Shotcut export completion detected from stable output growth without reopening the Jobs pane. elapsed=%s",
                        progress_logger.elapsed_text(),
                    )
                    break
            else:
                progress_logger.maybe_log(
                    detail=(
                        f"output={output_video_path.name} pending_creation=true "
                        f"stable_rounds={stable_rounds}"
                    )
                )
            time.sleep(SHOTCUT_EXPORT_POLL_SECONDS)
        else:
            raise AssertionError(f"Shotcut export did not finish within timeout: {output_video_path}")

        logger.info("Validating Shotcut export output file.")
        assert output_video_path.exists() and output_video_path.stat().st_size > 0, (
            f"Shotcut export output was not created correctly: {output_video_path}"
        )
        if jobs_dock is not None:
            logger.info("Validating Shotcut completed job row.")
            assert control_exists(
                jobs_dock,
                re.escape(str(output_video_path)),
                control_types=("TreeItem", "Text"),
            ), f"Shotcut completed row for '{output_video_path}' is missing from the jobs pane."
            logger.info("Validating Shotcut completed job duration.")
            assert control_exists(
                jobs_dock,
                SHOTCUT_TASK_TIME_PATTERN,
                control_types=("TreeItem",),
            ), "Shotcut jobs pane did not show the completed task duration."
        time.sleep(1.0)


class AvidemuxOperator(SoftwareOperator):
    def __init__(self, profile: OperationProfile) -> None:
        super().__init__(profile)
        self._active_process_id: int | None = None
        self._encode_dialog_minimized_to_tray = False

    def _wrapper_text(self, wrapper) -> str:
        try:
            return (wrapper.window_text() or "").strip()
        except Exception:
            return ""

    def _wrapper_control_type(self, wrapper) -> str:
        try:
            return (getattr(wrapper.element_info, "control_type", "") or "").strip()
        except Exception:
            return ""

    def _wrapper_class_name(self, wrapper) -> str:
        try:
            return (getattr(wrapper.element_info, "class_name", "") or "").strip()
        except Exception:
            return ""

    def _wrapper_rect(self, wrapper):
        return wrapper.rectangle()

    def _iter_wrapper_tree(self, root):
        yield root
        for child in root.descendants():
            yield child

    def _matches_patterns(self, text: str, patterns: tuple[str, ...]) -> bool:
        return any(re.search(pattern, text, re.IGNORECASE) for pattern in patterns)

    def _read_control_value(self, wrapper) -> str:
        readers = (
            lambda: wrapper.get_value(),
            lambda: wrapper.iface_value.CurrentValue,
            lambda: wrapper.window_text(),
            lambda: wrapper.texts()[0] if wrapper.texts() else "",
        )
        for reader in readers:
            try:
                value = reader()
            except Exception:
                continue
            if isinstance(value, str):
                return value.strip()
        return ""

    def _vertical_overlap(self, a, b) -> int:
        return max(0, min(a.bottom, b.bottom) - max(a.top, b.top))

    def _window_dimensions(self, window) -> tuple[int, int]:
        rect = self._wrapper_rect(window)
        return max(1, rect.right - rect.left), max(1, rect.bottom - rect.top)

    def _click_wrapper_center(self, wrapper, *, settle_seconds: float = 0.3) -> None:
        try:
            wrapper.click_input()
            time.sleep(settle_seconds)
            return
        except Exception as exc:
            logger.info(
                "Direct click failed for '%s'. Falling back to a center-coordinate click: %s",
                self._wrapper_text(wrapper) or "<untitled>",
                exc,
            )
        try:
            top_level = wrapper.top_level_parent()
        except Exception:
            top_level = wrapper
        rect = self._wrapper_rect(wrapper)
        top_rect = self._wrapper_rect(top_level)
        click_x_abs = rect.left + max(1, (rect.right - rect.left) // 2)
        click_y_abs = rect.top + max(1, (rect.bottom - rect.top) // 2)
        rel_x = max(1, min(click_x_abs - top_rect.left, (top_rect.right - top_rect.left) - 2))
        rel_y = max(1, min(click_y_abs - top_rect.top, (top_rect.bottom - top_rect.top) - 2))
        bring_window_to_front(top_level, keep_topmost=False)
        top_level.click_input(coords=(rel_x, rel_y))
        time.sleep(settle_seconds)

    def _list_main_windows(self):
        from pywinauto import Desktop

        candidates = []
        for window in Desktop(backend="uia").windows():
            title = self._wrapper_text(window)
            if "avidemux" not in title.lower():
                continue
            if self._wrapper_class_name(window) != AVIDEMUX_MAIN_WINDOW_CLASS:
                continue
            rect = self._wrapper_rect(window)
            area = max(0, rect.right - rect.left) * max(0, rect.bottom - rect.top)
            candidates.append(((area, rect.bottom, rect.right), window))
        candidates.sort(key=lambda item: item[0], reverse=True)
        return [window for _, window in candidates]

    def _connect_main_window(self, timeout: float = 15.0):
        deadline = time.monotonic() + timeout
        last_error: Exception | None = None
        while time.monotonic() < deadline:
            windows = self._list_main_windows()
            if windows:
                return windows[0]
            last_error = UiAutomationError("No Avidemux main window is currently visible.")
            time.sleep(0.25)
        raise UiAutomationError("Could not connect to the Avidemux main window.") from last_error

    def _process_id(self, window) -> int | None:
        try:
            process_id = getattr(getattr(window, "element_info", None), "process_id", None)
        except Exception:
            return None
        return int(process_id) if process_id else None

    def _is_visible(self, window) -> bool:
        try:
            return bool(window.is_visible())
        except Exception:
            return True

    def _iter_process_top_level_windows(self, process_id: int | None):
        if process_id is None:
            return []
        from pywinauto import Desktop

        matches = []
        for window in Desktop(backend="uia").windows():
            if self._process_id(window) != process_id:
                continue
            if not self._is_visible(window):
                continue
            try:
                rect = self._wrapper_rect(window)
            except Exception:
                continue
            area = max(0, rect.right - rect.left) * max(0, rect.bottom - rect.top)
            matches.append((rect.top, rect.left, -area, window))
        matches.sort(key=lambda item: (item[0], item[1], item[2]))
        return [item[-1] for item in matches]

    def _iter_desktop_top_level_windows(self):
        from pywinauto import Desktop

        matches = []
        for window in Desktop(backend="uia").windows():
            if not self._is_visible(window):
                continue
            try:
                rect = self._wrapper_rect(window)
            except Exception:
                continue
            area = max(0, rect.right - rect.left) * max(0, rect.bottom - rect.top)
            matches.append((rect.top, rect.left, -area, window))
        matches.sort(key=lambda item: (item[0], item[1], item[2]))
        return [item[-1] for item in matches]

    def _iter_process_window_handles(self, process_id: int | None, *, visible_only: bool | None = None) -> list[int]:
        if process_id is None:
            return []

        user32 = ctypes.windll.user32
        handles: list[int] = []
        callback_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

        def _callback(hwnd, _lparam):
            pid = wintypes.DWORD()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value != process_id:
                return True
            if visible_only is True and not user32.IsWindowVisible(hwnd):
                return True
            if visible_only is False and user32.IsWindowVisible(hwnd):
                return True
            handles.append(int(hwnd))
            return True

        user32.EnumWindows(callback_type(_callback), 0)
        return handles

    def _restore_process_windows(self, process_id: int | None, *, timeout: float = 2.0) -> bool:
        handles = self._iter_process_window_handles(process_id, visible_only=False)
        if not handles:
            return False

        logger.info("Restoring hidden Avidemux windows for process pid=%s.", process_id)
        user32 = ctypes.windll.user32
        restored = False
        for handle in handles:
            user32.ShowWindow(handle, 9)
            restored = True

        if not restored:
            return False

        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            if self._iter_process_top_level_windows(process_id):
                return True
            time.sleep(0.1)
        return False

    def _reset_runtime_state(self) -> None:
        self._active_process_id = None
        self._encode_dialog_minimized_to_tray = False

    def _first_matching_text(self, root, patterns: tuple[str, ...]) -> str:
        for wrapper in self._iter_wrapper_tree(root):
            text = self._wrapper_text(wrapper)
            if text and self._matches_patterns(text, patterns):
                return text
        return ""

    def _find_process_top_level_window(self, process_id: int | None, patterns: tuple[str, ...], *, timeout: float = 0.5):
        if process_id is None:
            return None

        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            matches = []
            for window in self._iter_process_top_level_windows(process_id):
                title = self._wrapper_text(window)
                if not title or not self._matches_patterns(title, patterns):
                    continue
                try:
                    rect = self._wrapper_rect(window)
                except Exception:
                    continue
                matches.append(((rect.top, rect.left), window))
            if matches:
                matches.sort(key=lambda item: item[0])
                return matches[0][1]
            time.sleep(0.1)
        return None

    def _find_process_top_level_window_by_content(
        self,
        process_id: int | None,
        patterns: tuple[str, ...],
        *,
        timeout: float = 0.5,
    ):
        if process_id is None:
            return None

        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            matches = []
            for window in self._iter_process_top_level_windows(process_id):
                matched_text = self._first_matching_text(window, patterns)
                if not matched_text:
                    continue
                try:
                    rect = self._wrapper_rect(window)
                except Exception:
                    continue
                matches.append(((rect.top, rect.left), window))
            if matches:
                matches.sort(key=lambda item: item[0])
                return matches[0][1]
            time.sleep(0.1)
        return None

    def _dismiss_thanks_popup_if_present(self, window, *, timeout: float = 0.5) -> bool:
        dialog = self._find_process_top_level_window(self._process_id(window), AVIDEMUX_THANKS_DIALOG_PATTERNS, timeout=timeout)
        if dialog is None:
            return False
        logger.info("Dismissing the Avidemux 'Thanks!' popup.")
        try:
            bring_window_to_front(dialog, keep_topmost=False)
            dialog.close()
        except Exception:
            request_window_close(dialog)
        time.sleep(0.5)
        return True

    def _dismiss_open_failure_dialog_if_present(self, window, *, timeout: float = 0.8) -> str:
        dialog = self._find_process_top_level_window(self._process_id(window), AVIDEMUX_INFO_DIALOG_PATTERNS, timeout=timeout)
        if dialog is None:
            return ""

        message = ""
        for wrapper in self._iter_wrapper_tree(dialog):
            text = self._wrapper_text(wrapper)
            if text and self._matches_patterns(text, AVIDEMUX_OPEN_FAILURE_PATTERNS):
                message = text
                break
        if not message:
            return ""

        logger.info("Dismissing the Avidemux open-file failure dialog: %s", message)
        try:
            ok_button = find_text_control(dialog, AVIDEMUX_DIALOG_OK_PATTERNS, control_types=("Button",))
            ok_button.click_input()
        except Exception:
            try:
                dialog.close()
            except Exception:
                request_window_close(dialog)
        time.sleep(0.5)
        return message

    def _dismiss_export_success_dialog_if_present(
        self,
        window,
        output_video_path: Path | None = None,
        *,
        timeout: float = 0.8,
    ) -> bool:
        process_id = self._process_id(window)
        dialog = self._find_process_top_level_window_by_content(
            process_id,
            AVIDEMUX_EXPORT_SUCCESS_PATTERNS,
            timeout=timeout,
        )
        if dialog is None:
            return False

        observed_texts: list[str] = []
        matched_text = ""
        output_name = output_video_path.name.casefold() if output_video_path is not None else ""
        for wrapper in self._iter_wrapper_tree(dialog):
            text = self._wrapper_text(wrapper)
            if not text:
                continue
            observed_texts.append(text)
            if self._matches_patterns(text, AVIDEMUX_EXPORT_SUCCESS_PATTERNS):
                matched_text = text
        if not matched_text:
            return False

        observed_blob = " | ".join(observed_texts).casefold()
        if output_name and output_name not in observed_blob and "successfully saved" not in observed_blob:
            return False

        logger.info(
            "Dismissing the Avidemux export completion dialog for '%s': %s",
            output_video_path.name if output_video_path is not None else "<unknown>",
            matched_text,
        )
        try:
            bring_window_to_front(dialog, keep_topmost=False)
            try:
                ok_button = find_text_control(dialog, AVIDEMUX_DIALOG_OK_PATTERNS, control_types=("Button", "Custom", "Pane"))
                ok_button.click_input()
            except Exception:
                send_hotkey("{ENTER}")
        except Exception:
            try:
                dialog.close()
            except Exception:
                request_window_close(dialog)
        time.sleep(0.5)
        return True

    def _wait_for_loaded_content(self, timeout: float = 15.0) -> bool:
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            try:
                window = self._connect_main_window(timeout=1.0)
            except UiAutomationError:
                time.sleep(0.2)
                continue
            zero_duration_visible = control_exists(
                window,
                re.escape("/ 00:00:00.000"),
                control_types=("Text",),
            )
            placeholder_visible = control_exists(
                window,
                r"^XXXX$",
                control_types=("Text",),
            )
            zero_tracks_visible = control_exists(
                window,
                r"^\(0 tracks\)$",
                control_types=("Text",),
            )
            if not zero_duration_visible and not placeholder_visible:
                return True
            if not zero_duration_visible and not zero_tracks_visible:
                return True
            time.sleep(0.2)
        return False

    def _find_export_progress_dialog(self, process_id: int | None, *, timeout: float = 2.0):
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            candidates = []
            seen_tokens: set[object] = set()
            search_sources = []
            if process_id is not None:
                search_sources.append((0, self._iter_process_top_level_windows(process_id)))
            search_sources.append((1, self._iter_desktop_top_level_windows()))

            for source_rank, windows in search_sources:
                for window in windows:
                    handle = getattr(window, "handle", None)
                    token = (handle, id(window)) if handle is not None else id(window)
                    if token in seen_tokens:
                        continue
                    seen_tokens.add(token)

                    title = self._wrapper_text(window)
                    title_matches = bool(title and self._matches_patterns(title, AVIDEMUX_PROGRESS_DIALOG_TITLE_PATTERNS))
                    content_match = self._first_matching_text(window, AVIDEMUX_PROGRESS_DIALOG_CONTENT_PATTERNS)
                    if not title_matches and not content_match:
                        continue

                    has_tray_action = False
                    try:
                        self._find_minimize_to_tray_control(window)
                        has_tray_action = True
                    except Exception:
                        try:
                            self._find_minimize_to_tray_geometry_candidate(window)
                            has_tray_action = True
                        except Exception:
                            pass

                    try:
                        rect = self._wrapper_rect(window)
                    except Exception:
                        continue
                    window_process_id = self._process_id(window)
                    candidates.append(
                        (
                            0 if has_tray_action else 1,
                            0 if process_id is not None and window_process_id == process_id else 1,
                            0 if title_matches else 1,
                            source_rank,
                            rect.top,
                            rect.left,
                            window,
                        )
                    )

            if candidates:
                candidates.sort(key=lambda item: item[:6])
                return candidates[0][6]
            time.sleep(0.1)
        return None

    def _find_minimize_to_tray_control(self, dialog):
        candidates = []
        for wrapper in self._iter_wrapper_tree(dialog):
            title = self._wrapper_text(wrapper)
            if not title or not self._matches_patterns(title, AVIDEMUX_MINIMIZE_TO_TRAY_PATTERNS):
                continue
            control_type = self._wrapper_control_type(wrapper).lower()
            if control_type and control_type not in {"button", "text", "pane", "custom", "group"}:
                continue
            rect = self._wrapper_rect(wrapper)
            candidates.append(((rect.top, rect.left, -(rect.right - rect.left)), wrapper))
        if not candidates:
            raise AssertionError("Avidemux tray-minimize control was not exposed by text.")
        candidates.sort(key=lambda item: item[0])
        return candidates[0][1]

    def _find_minimize_to_tray_geometry_candidate(self, dialog):
        dialog_rect = self._wrapper_rect(dialog)
        dialog_width = max(1, dialog_rect.right - dialog_rect.left)
        dialog_height = max(1, dialog_rect.bottom - dialog_rect.top)
        bottom_threshold = dialog_rect.bottom - int(dialog_height * AVIDEMUX_TRAY_BUTTON_BOTTOM_REGION_RATIO)
        left_threshold = dialog_rect.left + int(dialog_width * AVIDEMUX_TRAY_BUTTON_LEFT_REGION_RATIO)

        candidates = []
        for wrapper in self._iter_wrapper_tree(dialog):
            control_type = self._wrapper_control_type(wrapper).lower()
            if control_type and control_type not in {"button", "text", "pane", "custom", "group"}:
                continue
            rect = self._wrapper_rect(wrapper)
            width = rect.right - rect.left
            height = rect.bottom - rect.top
            if width <= 0 or height <= 0:
                continue
            if rect.left >= left_threshold:
                continue
            if rect.top < bottom_threshold:
                continue
            if width > AVIDEMUX_TRAY_BUTTON_MAX_WIDTH or height > AVIDEMUX_TRAY_BUTTON_MAX_HEIGHT:
                continue
            area = width * height
            candidates.append(((rect.top, rect.left, -area), wrapper))
        if not candidates:
            raise AssertionError("Avidemux tray-minimize geometry candidate was not found.")
        candidates.sort(key=lambda item: item[0])
        return candidates[0][1]

    def _minimize_encode_dialog_to_tray(self, window, *, timeout: float = 2.0) -> bool:
        process_id = self._process_id(window) or self._active_process_id
        dialog = self._find_export_progress_dialog(process_id, timeout=timeout)
        if dialog is None:
            logger.info("Avidemux encode-progress dialog was not found for tray minimization.")
            return False

        try:
            bring_window_to_front(dialog, keep_topmost=False)
        except Exception:
            pass

        try:
            tray_button = self._find_minimize_to_tray_control(dialog)
        except Exception as exc:
            logger.info("Avidemux tray-minimize text control was not exposed: %s", exc)
            try:
                tray_button = self._find_minimize_to_tray_geometry_candidate(dialog)
                logger.info(
                    "Using a geometry-based fallback for the Avidemux tray-minimize control: %s",
                    self._wrapper_text(tray_button) or "<untitled>",
                )
            except Exception as geometry_exc:
                logger.info("Avidemux tray-minimize geometry fallback failed: %s", geometry_exc)
                return False

        logger.info("Clicking the Avidemux tray-minimize button: %s", self._wrapper_text(tray_button) or "<untitled>")
        self._click_wrapper_center(tray_button, settle_seconds=0.3)
        self._active_process_id = process_id
        self._encode_dialog_minimized_to_tray = True
        return True

    def perform(self, input_video_path: Path, output_video_path: Path) -> None:
        logger.info("Avidemux flow started. input=%s output=%s", input_video_path, output_video_path)
        logger.info("Validating Avidemux input and output paths.")
        assert input_video_path.exists(), f"Input video does not exist: {input_video_path}"
        assert output_video_path.parent.exists(), f"Output directory does not exist: {output_video_path.parent}"
        assert output_video_path.suffix.lower() == ".mkv", f"Avidemux export target must be an mkv file: {output_video_path}"

        window = self._connect_main_window(timeout=30.0)
        self._active_process_id = self._process_id(window)
        bring_window_to_front(window, keep_topmost=False)
        self._dismiss_thanks_popup_if_present(window, timeout=1.0)

        logger.info("Opening the input clip in Avidemux.")
        self._open_input_clip(window, input_video_path)

        window = self._connect_main_window(timeout=AVIDEMUX_DIALOG_TIMEOUT_SECONDS)
        self._active_process_id = self._process_id(window) or self._active_process_id
        bring_window_to_front(window, keep_topmost=False)
        logger.info("Selecting the Avidemux video codec.")
        self._select_main_combo_value(window, combo_index=0, target_text=AVIDEMUX_VIDEO_CODEC_TEXT)
        logger.info("Selecting the Avidemux output container.")
        self._select_main_combo_value(window, combo_index=2, target_text=AVIDEMUX_MUXER_TEXT)

        logger.info("Saving the encoded Avidemux output.")
        minimized_to_tray = self._save_output(window, input_video_path, output_video_path)
        if not minimized_to_tray:
            self._begin_background_wait(window, phase="encode")
        logger.info("Waiting for the Avidemux encode to finish.")
        self._wait_for_export_completion(output_video_path)

        logger.info("Closing Avidemux after encode completion.")
        self.close()
        self._delete_generated_output_artifact(output_video_path, input_video_path)

    def close(self) -> None:
        logger.info("Closing Avidemux.")
        process_id = self._active_process_id
        if self._encode_dialog_minimized_to_tray and process_id is not None:
            self._restore_process_windows(process_id, timeout=1.5)
        try:
            window = self._connect_main_window(timeout=2.0)
        except UiAutomationError:
            if process_id is not None and self._restore_process_windows(process_id, timeout=1.5):
                try:
                    window = self._connect_main_window(timeout=2.0)
                except UiAutomationError:
                    visible_windows = self._iter_process_top_level_windows(process_id)
                    if not visible_windows:
                        logger.warning("Avidemux UI could not be restored for a graceful close. Terminating process pid=%s.", process_id)
                        subprocess.run(
                            ["taskkill", "/PID", str(process_id), "/T", "/F"],
                            capture_output=True,
                            text=True,
                            check=False,
                        )
                        self._reset_runtime_state()
                        return
                    window = visible_windows[0]
            else:
                logger.info("Avidemux window is already closed.")
                self._reset_runtime_state()
                return
        process_id = self._process_id(window) or process_id

        for attempt in range(1, AVIDEMUX_CLOSE_RETRY_COUNT + 1):
            try:
                self._dismiss_export_success_dialog_if_present(window, None, timeout=0.8)
            except Exception:
                logger.info("Ignoring a transient failure while dismissing the Avidemux completion dialog during close.")

            try:
                bring_window_to_front(window, keep_topmost=False)
            except Exception:
                pass
            try:
                window.close()
            except Exception as exc:
                logger.info("Avidemux window.close() raised %s. Falling back to WM_CLOSE.", exc)
                request_window_close(window)
            time.sleep(AVIDEMUX_CLOSE_SETTLE_SECONDS)
            dismiss_close_prompts(owner_window=window)

            remaining_windows = self._iter_process_top_level_windows(process_id)
            if not remaining_windows:
                logger.info("Avidemux UI closed successfully.")
                self._reset_runtime_state()
                return
            logger.info(
                "Avidemux still has visible windows after close attempt %d: %s",
                attempt,
                [self._wrapper_text(candidate) or "<untitled>" for candidate in remaining_windows[:5]],
            )

        if process_id is None:
            self._reset_runtime_state()
            return

        logger.warning("Avidemux still has visible UI after close retries. Terminating process pid=%s.", process_id)
        subprocess.run(
            ["taskkill", "/PID", str(process_id), "/T", "/F"],
            capture_output=True,
            text=True,
            check=False,
        )
        deadline = time.monotonic() + AVIDEMUX_CLOSE_TIMEOUT_SECONDS
        while time.monotonic() < deadline:
            if not self._iter_process_top_level_windows(process_id):
                logger.info("Avidemux UI closed after process termination.")
                self._reset_runtime_state()
                return
            time.sleep(0.1)
        self._reset_runtime_state()

    def _delete_generated_output_artifact(self, output_video_path: Path, input_video_path: Path) -> None:
        output_resolved = output_video_path.resolve(strict=False)
        input_resolved = input_video_path.resolve(strict=False)
        if output_resolved == input_resolved:
            logger.info("Skipping Avidemux output cleanup because the candidate matches the input video.")
            return
        if not output_video_path.exists():
            logger.info("Skipping Avidemux output cleanup because the file no longer exists: %s", output_video_path)
            return
        logger.info("Deleting the generated Avidemux output: %s", output_video_path)
        output_video_path.unlink()

    def _main_combos(self, window):
        combos = []
        for wrapper in self._iter_wrapper_tree(window):
            if self._wrapper_control_type(wrapper).lower() != "combobox":
                continue
            if self._wrapper_class_name(wrapper) != "QComboBox":
                continue
            combos.append(wrapper)
        combos.sort(key=lambda wrapper: (self._wrapper_rect(wrapper).top, self._wrapper_rect(wrapper).left))
        return combos

    def _dialog_filename_edit(self, window):
        candidates = []
        for wrapper in self._iter_wrapper_tree(window):
            control_type = self._wrapper_control_type(wrapper).lower()
            if control_type != "edit":
                continue
            rect = self._wrapper_rect(wrapper)
            width = rect.right - rect.left
            if rect.left < AVIDEMUX_FILE_ENTRY_MIN_LEFT:
                continue
            if not AVIDEMUX_FILE_ENTRY_TOP_MIN <= rect.top <= AVIDEMUX_FILE_ENTRY_TOP_MAX:
                continue
            if width < AVIDEMUX_FILE_ENTRY_MIN_WIDTH:
                continue
            class_name = self._wrapper_class_name(wrapper)
            if class_name == "SearchEditBox":
                continue
            score = (rect.top, rect.left, -width)
            candidates.append((score, wrapper))
        assert candidates, "Avidemux did not expose the embedded filename entry field."
        candidates.sort(key=lambda item: item[0])
        return candidates[0][1]

    def _focus_dialog_filename_entry(self, window) -> None:
        edit = self._dialog_filename_edit(window)
        rect = self._wrapper_rect(edit)
        main_rect = self._wrapper_rect(window)
        click_x_abs = rect.left + max(12, min(48, (rect.right - rect.left) // 10))
        click_y_abs = rect.top + max(1, (rect.bottom - rect.top) // 2)
        rel_x = max(1, min(click_x_abs - main_rect.left, (main_rect.right - main_rect.left) - 2))
        rel_y = max(1, min(click_y_abs - main_rect.top, (main_rect.bottom - main_rect.top) - 2))
        bring_window_to_front(window, keep_topmost=False)
        window.click_input(coords=(rel_x, rel_y))
        time.sleep(0.2)

    def _dialog_bottom_buttons(self, window):
        candidates = []
        for wrapper in self._iter_wrapper_tree(window):
            if self._wrapper_control_type(wrapper).lower() != "button":
                continue
            rect = self._wrapper_rect(wrapper)
            if rect.left < AVIDEMUX_ACTION_BUTTON_LEFT_MIN:
                continue
            if rect.top < AVIDEMUX_ACTION_BUTTON_TOP_MIN or rect.bottom > AVIDEMUX_ACTION_BUTTON_TOP_MAX:
                continue
            title = self._wrapper_text(wrapper)
            width = rect.right - rect.left
            score = (rect.bottom, width, rect.left)
            candidates.append((score, title, wrapper))
        candidates.sort(key=lambda item: item[0], reverse=True)
        return candidates

    def _find_bottom_button(self, window, patterns: tuple[str, ...]):
        matches = []
        for _, title, wrapper in self._dialog_bottom_buttons(window):
            if title and self._matches_patterns(title, patterns):
                matches.append(wrapper)
        if not matches:
            raise AssertionError(f"Avidemux did not expose a bottom-row button matching {patterns}.")
        matches.sort(key=lambda wrapper: (self._wrapper_rect(wrapper).bottom, self._wrapper_rect(wrapper).left), reverse=True)
        return matches[0]

    def _click_hidden_open_confirmation(self, window) -> None:
        cancel_button = self._find_bottom_button(window, AVIDEMUX_CANCEL_PATTERNS)
        cancel_rect = self._wrapper_rect(cancel_button)
        button_width = max(1, cancel_rect.right - cancel_rect.left)
        gap = max(AVIDEMUX_OPEN_CONFIRM_MIN_GAP, int(round(button_width * AVIDEMUX_OPEN_CONFIRM_GAP_RATIO)))
        click_x_abs = cancel_rect.left - gap - max(1, button_width // 2)
        click_y_abs = cancel_rect.top + max(1, (cancel_rect.bottom - cancel_rect.top) // 2)

        main_rect = self._wrapper_rect(window)
        rel_x = max(1, min(click_x_abs - main_rect.left, (main_rect.right - main_rect.left) - 2))
        rel_y = max(1, min(click_y_abs - main_rect.top, (main_rect.bottom - main_rect.top) - 2))

        logger.info("Clicking the hidden Avidemux open confirmation point. absolute=(%d,%d)", click_x_abs, click_y_abs)
        bring_window_to_front(window, keep_topmost=False)
        window.click_input(coords=(rel_x, rel_y))
        time.sleep(AVIDEMUX_DIALOG_SETTLE_SECONDS)

    def _visible_file_item(self, window, input_video_path: Path):
        candidates = []
        expected_names = {input_video_path.name.casefold(), input_video_path.stem.casefold()}
        main_rect = self._wrapper_rect(window)
        min_left = main_rect.left + AVIDEMUX_LIST_ITEM_LEFT_OFFSET
        min_top = main_rect.top + AVIDEMUX_LIST_ITEM_TOP_OFFSET
        max_bottom = main_rect.bottom - AVIDEMUX_LIST_ITEM_BOTTOM_OFFSET
        for wrapper in self._iter_wrapper_tree(window):
            if self._wrapper_control_type(wrapper).lower() != "listitem":
                continue
            title = self._wrapper_text(wrapper)
            if title.casefold() not in expected_names:
                continue
            rect = self._wrapper_rect(wrapper)
            if rect.left < min_left or rect.top < min_top or rect.bottom > max_bottom:
                continue
            area = max(0, rect.right - rect.left) * max(0, rect.bottom - rect.top)
            candidates.append(((area, rect.bottom, rect.left), wrapper))
        assert candidates, f"Avidemux did not show the visible file item for '{input_video_path.name}'."
        candidates.sort(key=lambda item: item[0], reverse=True)
        return candidates[0][1]

    def _wait_for_loaded_title(self, input_video_path: Path, timeout: float = 15.0) -> bool:
        deadline = time.monotonic() + timeout
        expected_name = input_video_path.name.casefold()
        while time.monotonic() < deadline:
            try:
                window = self._connect_main_window(timeout=1.0)
            except UiAutomationError:
                time.sleep(0.2)
                continue
            title = self._wrapper_text(window).casefold()
            if expected_name in title:
                return True
            time.sleep(0.2)
        return False

    def _open_input_clip(self, window, input_video_path: Path) -> None:
        self._dismiss_thanks_popup_if_present(window, timeout=1.0)
        bring_window_to_front(window, keep_topmost=False)
        send_hotkey("^o")
        time.sleep(AVIDEMUX_DIALOG_SETTLE_SECONDS)
        logger.info("Using the standard Avidemux file dialog to open the input clip.")
        fill_file_dialog(
            input_video_path,
            dialog_patterns=AVIDEMUX_OPEN_DIALOG_PATTERNS,
            must_exist=True,
            timeout=AVIDEMUX_DIALOG_TIMEOUT_SECONDS,
        )
        open_failure = self._dismiss_open_failure_dialog_if_present(window, timeout=1.5)
        if open_failure:
            if self._matches_patterns(open_failure, AVIDEMUX_DEMUXER_FAILURE_PATTERNS):
                raise AssertionError(
                    "Avidemux could not load the input clip because its demuxer plugins are unavailable. "
                    f"Dialog: {open_failure}"
                )
            raise AssertionError(f"Avidemux could not open the selected input video: {open_failure}")
        assert self._wait_for_loaded_content(timeout=15.0), (
            f"Avidemux did not finish loading the selected input video: {input_video_path}"
        )

    def _select_main_combo_value(self, window, *, combo_index: int, target_text: str) -> None:
        combos = self._main_combos(window)
        assert len(combos) >= combo_index + 1, (
            f"Avidemux did not expose the expected combo index {combo_index}. visible_combos={len(combos)}"
        )
        combo = combos[combo_index]
        try:
            if combo.selected_text() == target_text:
                return
        except Exception:
            pass

        bring_window_to_front(window, keep_topmost=False)
        combo.click_input()
        time.sleep(AVIDEMUX_DIALOG_SETTLE_SECONDS)

        combo_rect = self._wrapper_rect(combo)
        candidates = []
        for wrapper in self._iter_wrapper_tree(window):
            if self._wrapper_control_type(wrapper).lower() != "listitem":
                continue
            if self._wrapper_text(wrapper) != target_text:
                continue
            rect = self._wrapper_rect(wrapper)
            if rect.top < combo_rect.top - 20:
                continue
            score = (abs(rect.left - combo_rect.left), abs(rect.top - combo_rect.bottom), rect.top, rect.left)
            candidates.append((score, wrapper))
        assert candidates, f"Avidemux did not expose the list item '{target_text}'."
        candidates.sort(key=lambda item: item[0])
        candidates[0][1].click_input()
        time.sleep(AVIDEMUX_DIALOG_SETTLE_SECONDS)

        selected_text = ""
        deadline = time.monotonic() + AVIDEMUX_CODEC_SELECTION_TIMEOUT_SECONDS
        while time.monotonic() < deadline:
            try:
                selected_text = combo.selected_text()
            except Exception:
                selected_text = ""
            if selected_text == target_text:
                return
            time.sleep(0.2)
        raise AssertionError(f"Avidemux combo selection did not settle to '{target_text}'. Last value: '{selected_text}'.")

    def _save_output(self, window, input_video_path: Path, output_video_path: Path) -> bool:
        bring_window_to_front(window, keep_topmost=False)
        send_hotkey("^s")
        time.sleep(AVIDEMUX_DIALOG_SETTLE_SECONDS)

        logger.info("Focusing the Avidemux embedded save-filename field.")
        self._focus_dialog_filename_entry(window)
        paste_text_via_clipboard(str(output_video_path), replace_existing=True)
        time.sleep(0.5)

        save_button = self._find_bottom_button(window, AVIDEMUX_SAVE_BUTTON_PATTERNS)
        logger.info("Clicking the Avidemux save confirmation button: %s", self._wrapper_text(save_button) or "<untitled>")
        save_button.click_input()
        time.sleep(0.5)

        self._handle_export_overwrite_prompt(window, input_video_path, output_video_path)

        deadline = time.monotonic() + AVIDEMUX_SAVE_START_TIMEOUT_SECONDS
        while time.monotonic() < deadline:
            if output_video_path.exists() and output_video_path.stat().st_size > 0:
                logger.info("Avidemux started writing the output file: %s", output_video_path)
                logger.info("Trying to minimize Avidemux to tray immediately after export output starts.")
                minimized_to_tray = self._minimize_encode_dialog_to_tray(window, timeout=5.0)
                if minimized_to_tray:
                    logger.info("Avidemux was minimized to tray after export output started.")
                else:
                    logger.info("Avidemux tray minimization was not available immediately after export output started.")
                return minimized_to_tray
            time.sleep(0.5)
        raise AssertionError(f"Avidemux did not start writing the requested output: {output_video_path}")

    def _embedded_overwrite_prompt_text(self, window) -> str:
        for wrapper in self._iter_wrapper_tree(window):
            title = self._wrapper_text(wrapper)
            if title and self._matches_patterns(title, AVIDEMUX_OVERWRITE_PROMPT_PATTERNS):
                return title
        return ""

    def _handle_export_overwrite_prompt(
        self,
        window,
        input_video_path: Path,
        output_video_path: Path,
        *,
        timeout: float = 3.0,
    ) -> None:
        logger.info("Checking for Avidemux overwrite confirmation after clicking Save.")
        overwrite_accepted = accept_overwrite_confirmation(timeout=0.8, owner_window=window, poll_interval=0.05)
        if overwrite_accepted:
            logger.info("Accepted top-level Avidemux overwrite confirmation for '%s'.", output_video_path.name)
            return
        self._handle_embedded_overwrite_prompt(
            window,
            input_video_path,
            output_video_path,
            timeout=max(timeout - 0.8, 0.5),
        )

    def _handle_embedded_overwrite_prompt(
        self,
        window,
        input_video_path: Path,
        output_video_path: Path,
        *,
        timeout: float = 3.0,
    ) -> None:
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            prompt_text = self._embedded_overwrite_prompt_text(window)
            if not prompt_text:
                try:
                    if accept_overwrite_confirmation(timeout=0.2, owner_window=window, poll_interval=0.05):
                        logger.info("Accepted delayed top-level Avidemux overwrite confirmation for '%s'.", output_video_path.name)
                        return
                except Exception:
                    pass
                time.sleep(0.1)
                continue

            logger.info("Detected embedded Avidemux overwrite prompt: %s", prompt_text)
            expected_name = output_video_path.name.casefold()
            input_name = input_video_path.name.casefold()
            prompt_name = prompt_text.casefold()

            if input_name in prompt_name and expected_name not in prompt_name:
                no_button = self._find_bottom_button(window, AVIDEMUX_NO_PATTERNS)
                logger.info("Rejecting unexpected Avidemux overwrite target: %s", prompt_text)
                no_button.click_input()
                raise AssertionError(
                    f"Avidemux attempted to overwrite the input clip instead of '{output_video_path.name}': {prompt_text}"
                )

            confirm_button = self._find_bottom_button(window, AVIDEMUX_OVERWRITE_CONFIRM_PATTERNS)
            logger.info(
                "Accepting the Avidemux overwrite prompt for '%s' with button: %s",
                output_video_path.name,
                self._wrapper_text(confirm_button) or "<untitled>",
            )
            confirm_button.click_input()
            time.sleep(0.5)
            return
        logger.info("No embedded Avidemux overwrite prompt appeared.")

    def _wait_for_export_completion(self, output_video_path: Path, window=None) -> None:
        assert output_video_path.parent.exists(), f"Output directory disappeared: {output_video_path.parent}"

        stable_rounds = 0
        last_size = -1
        deadline = time.monotonic() + AVIDEMUX_EXPORT_TIMEOUT_SECONDS
        progress_logger = _WaitProgressLogger("Avidemux export")
        while time.monotonic() < deadline:
            if window is None:
                try:
                    window = self._connect_main_window(timeout=0.5)
                except UiAutomationError:
                    window = None
            if window is not None:
                try:
                    if self._dismiss_export_success_dialog_if_present(window, output_video_path, timeout=0.1):
                        logger.info("Avidemux export completion dialog detected.")
                        break
                except Exception as exc:
                    logger.info("Ignoring an Avidemux completion-dialog probe failure: %s", exc)
            current_size = output_video_path.stat().st_size if output_video_path.exists() else 0
            if current_size > 0 and current_size == last_size:
                stable_rounds += 1
            else:
                stable_rounds = 0
            last_size = current_size
            progress_logger.maybe_log(
                detail=(
                    f"output={output_video_path.name} size_bytes={current_size} "
                    f"stable_rounds={stable_rounds}"
                )
            )
            if current_size > 0 and stable_rounds >= AVIDEMUX_EXPORT_STABLE_ROUNDS:
                if window is not None:
                    try:
                        self._dismiss_export_success_dialog_if_present(window, output_video_path, timeout=0.1)
                    except Exception as exc:
                        logger.info("Ignoring a final Avidemux completion-dialog dismissal failure: %s", exc)
                logger.info("Avidemux export completion conditions satisfied. elapsed=%s", progress_logger.elapsed_text())
                break
            time.sleep(AVIDEMUX_EXPORT_POLL_SECONDS)
        else:
            raise AssertionError(f"Avidemux export did not finish within timeout: {output_video_path}")

        assert output_video_path.exists() and output_video_path.stat().st_size > 0, (
            f"Avidemux export output was not created correctly: {output_video_path}"
        )


class HandBrakeOperator(SoftwareOperator):
    def _wrapper_text(self, wrapper) -> str:
        try:
            return (wrapper.window_text() or "").strip()
        except Exception:
            return ""

    def _wrapper_control_type(self, wrapper) -> str:
        try:
            return (getattr(wrapper.element_info, "control_type", "") or "").strip()
        except Exception:
            return ""

    def _wrapper_rect(self, wrapper):
        return wrapper.rectangle()

    def _iter_wrapper_tree(self, root):
        yield root
        for child in root.descendants():
            yield child

    def _matches_patterns(self, text: str, patterns: tuple[str, ...]) -> bool:
        return any(re.search(pattern, text, re.IGNORECASE) for pattern in patterns)

    def _read_control_value(self, wrapper) -> str:
        readers = (
            lambda: wrapper.get_value(),
            lambda: wrapper.iface_value.CurrentValue,
            lambda: wrapper.window_text(),
            lambda: wrapper.texts()[0] if wrapper.texts() else "",
        )
        for reader in readers:
            try:
                value = reader()
            except Exception:
                continue
            if isinstance(value, str):
                return value.strip()
        return ""

    def _vertical_overlap(self, a, b) -> int:
        return max(0, min(a.bottom, b.bottom) - max(a.top, b.top))

    def _window_dimensions(self, window) -> tuple[int, int]:
        rect = self._wrapper_rect(window)
        return max(1, rect.right - rect.left), max(1, rect.bottom - rect.top)

    def _iter_desktop_wrappers(self):
        from pywinauto import Desktop

        for window in Desktop(backend="uia").windows():
            yield window
            for child in window.descendants():
                yield child

    def _iter_process_wrappers(self, process_id: int | None):
        if not process_id:
            return
        from pywinauto import Desktop

        for window in Desktop(backend="uia").windows():
            try:
                wrapper_process_id = int(getattr(getattr(window, "element_info", None), "process_id", 0) or 0)
            except Exception:
                continue
            if wrapper_process_id != process_id:
                continue
            yield window
            for child in window.descendants():
                yield child

    def _list_main_windows(self):
        from pywinauto import Desktop

        candidates = []
        for window in Desktop(backend="uia").windows():
            title = self._wrapper_text(window)
            if "handbrake" not in title.lower():
                continue
            rect = self._wrapper_rect(window)
            area = max(0, rect.right - rect.left) * max(0, rect.bottom - rect.top)
            candidates.append(((area, rect.bottom, rect.right), window))
        candidates.sort(key=lambda item: item[0], reverse=True)
        return [window for _, window in candidates]

    def _connect_main_window(self, timeout: float = 15.0):
        deadline = time.monotonic() + timeout
        last_error: Exception | None = None
        while time.monotonic() < deadline:
            self._dismiss_update_dialog_if_present(timeout=0.2)
            self._dismiss_recovery_dialog_if_present(timeout=0.2)
            windows = self._list_main_windows()
            if windows:
                return windows[0]
            last_error = UiAutomationError("No HandBrake main window is currently visible.")
            time.sleep(0.25)
        raise UiAutomationError("Could not connect to the HandBrake main window.") from last_error

    def _dismiss_update_dialog_if_present(self, timeout: float = 0.5) -> bool:
        try:
            dialog = wait_for_window(HAND_BRAKE_UPDATE_DIALOG_PATTERNS, timeout=timeout)
        except UiAutomationError:
            return False

        prompt_matches = self._find_matching_wrappers(
            dialog,
            HAND_BRAKE_UPDATE_PROMPT_PATTERNS,
            control_types=("Text", "Pane", "Group", "Document", "Custom"),
        )
        if not prompt_matches:
            logger.info(
                "HandBrake update dialog title matched, but the prompt text was not exposed. "
                "Proceeding with the dismissal anyway."
            )

        logger.info("Dismissing the HandBrake update-check prompt.")
        bring_window_to_front(dialog, keep_topmost=False)
        try:
            no_button = find_text_control(dialog, HAND_BRAKE_UPDATE_NO_PATTERNS, control_types=("Button",))
            click_control(no_button, post_click_sleep=HAND_BRAKE_DIALOG_SETTLE_SECONDS)
        except UiAutomationError:
            logger.info(
                "HandBrake update dialog negative action was not exposed through UIA. "
                "Falling back to key '%s'.",
                HAND_BRAKE_UPDATE_DECLINE_KEY,
            )
            send_hotkey(HAND_BRAKE_UPDATE_DECLINE_KEY)
            time.sleep(HAND_BRAKE_DIALOG_SETTLE_SECONDS)
        return True

    def _dismiss_recovery_dialog_if_present(self, timeout: float = 0.5) -> bool:
        try:
            dialog = wait_for_window(HAND_BRAKE_RECOVERY_DIALOG_PATTERNS, timeout=timeout)
        except UiAutomationError:
            return False

        logger.info("Dismissing the HandBrake recoverable-queue prompt.")
        bring_window_to_front(dialog, keep_topmost=False)
        no_button = find_text_control(dialog, HAND_BRAKE_RECOVERY_NO_PATTERNS, control_types=("Button",))
        no_button.click_input()
        time.sleep(0.8)
        return True

    def perform(self, input_video_path: Path, output_video_path: Path) -> None:
        logger.info("HandBrake flow started. input=%s output=%s", input_video_path, output_video_path)
        logger.info("Validating HandBrake input and output paths.")
        assert input_video_path.exists(), f"Input video does not exist: {input_video_path}"
        assert output_video_path.parent.exists(), f"Output directory does not exist: {output_video_path.parent}"
        assert output_video_path.suffix.lower() == ".mp4", f"HandBrake export target must be an mp4 file: {output_video_path}"

        window = self._connect_main_window(timeout=30.0)
        bring_window_to_front(window, keep_topmost=False)

        logger.info("Opening the source clip in HandBrake.")
        self._open_input_clip(window, input_video_path)

        window = self._connect_main_window(timeout=HAND_BRAKE_DIALOG_TIMEOUT_SECONDS)
        bring_window_to_front(window, keep_topmost=False)
        logger.info("Selecting the HandBrake preset: %s", HAND_BRAKE_TARGET_PRESET_TEXT)
        self._select_preset(window, HAND_BRAKE_TARGET_PRESET_TEXT)
        logger.info("Setting the HandBrake output path.")
        self._set_output_path(window, output_video_path)
        logger.info("Starting the HandBrake encode.")
        self._start_encode(window, output_video_path)
        self._begin_background_wait(window, phase="encode")
        logger.info("Waiting for the HandBrake encode to finish.")
        self._wait_for_export_completion(output_video_path, window)

        logger.info("Closing HandBrake after encode completion.")
        self.close()
        self._delete_generated_output_artifact(output_video_path, input_video_path)

    def close(self) -> None:
        logger.info("Closing HandBrake.")
        try:
            window = self._connect_main_window(timeout=2.0)
        except UiAutomationError:
            logger.info("HandBrake window is already closed.")
            return
        bring_window_to_front(window, keep_topmost=False)
        try:
            window.close()
        except Exception as exc:
            logger.info("HandBrake window.close() raised %s. Falling back to WM_CLOSE.", exc)
            request_window_close(window)
        time.sleep(1.0)
        dismiss_close_prompts(owner_window=window)

    def _delete_generated_output_artifact(self, output_video_path: Path, input_video_path: Path) -> None:
        output_resolved = output_video_path.resolve(strict=False)
        input_resolved = input_video_path.resolve(strict=False)
        if output_resolved == input_resolved:
            logger.info("Skipping HandBrake output cleanup because the candidate matches the input video.")
            return
        if not output_video_path.exists():
            logger.info("Skipping HandBrake output cleanup because the file no longer exists: %s", output_video_path)
            return
        logger.info("Deleting the generated HandBrake output: %s", output_video_path)
        output_video_path.unlink()

    def _find_matching_wrappers(
        self,
        window,
        patterns: tuple[str, ...],
        *,
        control_types: tuple[str, ...] = (),
    ):
        allowed_types = {control_type.lower() for control_type in control_types}
        matches = []
        for wrapper in self._iter_wrapper_tree(window):
            control_type = self._wrapper_control_type(wrapper).lower()
            if allowed_types and control_type not in allowed_types:
                continue
            title = self._wrapper_text(wrapper)
            if not title or not self._matches_patterns(title, patterns):
                continue
            matches.append(wrapper)
        return matches

    def _find_button(
        self,
        window,
        patterns: tuple[str, ...],
        *,
        top_min: int | None = None,
        top_max: int | None = None,
        left_min: int | None = None,
        left_max: int | None = None,
    ):
        candidates = []
        for wrapper in self._find_matching_wrappers(window, patterns, control_types=("Button",)):
            rect = self._wrapper_rect(wrapper)
            if top_min is not None and rect.top < top_min:
                continue
            if top_max is not None and rect.bottom > top_max:
                continue
            if left_min is not None and rect.left < left_min:
                continue
            if left_max is not None and rect.right > left_max:
                continue
            width = rect.right - rect.left
            candidates.append(((rect.bottom, width, rect.left), wrapper))
        assert candidates, f"HandBrake did not expose a button matching {patterns}."
        candidates.sort(key=lambda item: item[0], reverse=True)
        return candidates[0][1]

    def _find_labeled_combo(self, window, label_patterns: tuple[str, ...]):
        labels = self._find_matching_wrappers(
            window,
            label_patterns,
            control_types=("Text", "Pane", "Group", "Document", "Custom"),
        )
        combos = [
            wrapper
            for wrapper in self._iter_wrapper_tree(window)
            if self._wrapper_control_type(wrapper).lower() == "combobox"
        ]
        candidates = []
        for label in labels:
            label_rect = self._wrapper_rect(label)
            for combo in combos:
                combo_rect = self._wrapper_rect(combo)
                if combo_rect.left <= label_rect.right:
                    continue
                center_y_distance = abs(
                    ((combo_rect.top + combo_rect.bottom) // 2) - ((label_rect.top + label_rect.bottom) // 2)
                )
                if center_y_distance > max(40, (label_rect.bottom - label_rect.top) * 2):
                    continue
                width = combo_rect.right - combo_rect.left
                candidates.append(
                    (
                        center_y_distance,
                        combo_rect.left - label_rect.right,
                        -width,
                        combo_rect.top,
                        combo,
                    )
                )
        if not candidates:
            raise AssertionError(f"HandBrake did not expose a combo box near label patterns {label_patterns}.")
        candidates.sort(key=lambda item: item[:4])
        return candidates[0][4]

    def _find_labeled_row_control(
        self,
        window,
        label_patterns: tuple[str, ...],
        *,
        allowed_types: tuple[str, ...],
    ):
        labels = self._find_matching_wrappers(
            window,
            label_patterns,
            control_types=("Text", "Pane", "Group", "Document", "Custom"),
        )
        allowed_type_set = {value.lower() for value in allowed_types}
        candidates = []
        for label in labels:
            label_rect = self._wrapper_rect(label)
            label_text = self._normalized_ui_text(self._wrapper_text(label))
            for wrapper in self._iter_wrapper_tree(window):
                control_type = self._wrapper_control_type(wrapper).lower()
                if control_type not in allowed_type_set:
                    continue
                rect = self._wrapper_rect(wrapper)
                if rect.left <= label_rect.right:
                    continue
                overlap = self._vertical_overlap(label_rect, rect)
                if overlap <= 0:
                    continue
                width = rect.right - rect.left
                if width <= 0:
                    continue
                current_text = self._selector_selected_text(wrapper)
                normalized_text = self._normalized_ui_text(current_text)
                if normalized_text and normalized_text == label_text:
                    continue
                candidates.append(
                    (
                        0 if self._looks_like_preset_value(current_text) else 1,
                        abs(((rect.top + rect.bottom) // 2) - ((label_rect.top + label_rect.bottom) // 2)),
                        rect.left - label_rect.right,
                        -width,
                        rect.top,
                        current_text,
                        wrapper,
                    )
                )
        if not candidates:
            raise AssertionError(f"HandBrake did not expose a row control near label patterns {label_patterns}.")
        candidates.sort(key=lambda item: item[:5])
        return candidates[0][6]

    def _normalized_ui_text(self, text: str) -> str:
        return re.sub(r"\s+", " ", text).strip().casefold()

    def _combo_selected_text(self, combo) -> str:
        readers = (
            lambda: combo.selected_text(),
            lambda: combo.get_value(),
            lambda: combo.iface_value.CurrentValue,
            lambda: combo.window_text(),
            lambda: combo.texts()[0] if combo.texts() else "",
        )
        for reader in readers:
            try:
                value = reader()
            except Exception:
                continue
            if isinstance(value, str) and value.strip():
                return value.strip()
        return ""

    def _selector_selected_text(self, wrapper) -> str:
        return self._combo_selected_text(wrapper) or self._read_control_value(wrapper) or self._wrapper_text(wrapper)

    def _text_matches_target(self, text: str, target_text: str) -> bool:
        normalized_text = self._normalized_ui_text(text)
        normalized_target = self._normalized_ui_text(target_text)
        if not normalized_text or not normalized_target:
            return False
        return normalized_text == normalized_target or normalized_target in normalized_text

    def _looks_like_preset_value(self, text: str) -> bool:
        normalized_text = self._normalized_ui_text(text)
        if not normalized_text:
            return False
        return any(hint in normalized_text for hint in HAND_BRAKE_PRESET_VALUE_HINTS)

    def _looks_like_title_selector_value(self, text: str) -> bool:
        normalized_text = self._normalized_ui_text(text)
        if not normalized_text:
            return False
        return bool(re.match(HAND_BRAKE_TITLE_SELECTOR_VALUE_RE, normalized_text))

    def _set_combo_text_with_keyboard(self, window, combo, target_text: str) -> None:
        logger.info("Trying a keyboard-entry fallback for the HandBrake preset combo: %s", target_text)
        bring_window_to_front(window, keep_topmost=False)
        try:
            combo.click_input()
        except Exception:
            pass
        time.sleep(0.2)
        paste_text_via_clipboard(target_text, replace_existing=True)
        time.sleep(0.2)
        send_hotkey("{ENTER}")
        time.sleep(HAND_BRAKE_DIALOG_SETTLE_SECONDS)

    def _preset_combo(self, window):
        try:
            selector = self._find_labeled_row_control(
                window,
                HAND_BRAKE_PRESET_LABEL_PATTERNS,
                allowed_types=("Button", "ComboBox", "Edit", "Pane", "Custom", "Group"),
            )
            logger.info(
                "Resolved the HandBrake preset selector from the labeled row. type=%s text=%s",
                self._wrapper_control_type(selector) or "<unknown>",
                self._selector_selected_text(selector) or "<empty>",
            )
            return selector
        except AssertionError:
            pass

        try:
            return self._find_labeled_combo(window, HAND_BRAKE_PRESET_LABEL_PATTERNS)
        except AssertionError:
            pass

        window_rect = self._wrapper_rect(window)
        window_width, window_height = self._window_dimensions(window)
        candidates = []
        for wrapper in self._iter_wrapper_tree(window):
            if self._wrapper_control_type(wrapper).lower() != "combobox":
                continue
            rect = self._wrapper_rect(wrapper)
            width = rect.right - rect.left
            if width <= 0:
                continue
            relative_top = (rect.top - window_rect.top) / window_height
            relative_left = (rect.left - window_rect.left) / window_width
            current_text = self._selector_selected_text(wrapper)
            candidates.append(
                (
                    0 if self._looks_like_preset_value(current_text) else 1,
                    1 if self._looks_like_title_selector_value(current_text) else 0,
                    0 if relative_top <= 0.35 else 1,
                    0 if relative_left <= 0.55 else 1,
                    0 if current_text else 1,
                    relative_top,
                    -width,
                    rect.left,
                    current_text,
                    wrapper,
                )
            )
        assert candidates, "HandBrake did not expose a preset combo box."
        candidates.sort(key=lambda item: item[:8])
        logger.info(
            "HandBrake preset combo fallback candidates: %s",
            [
                {
                    "text": candidate[8] or "<empty>",
                    "top_bias": candidate[2],
                    "left_bias": candidate[3],
                    "looks_like_preset": candidate[0] == 0,
                    "looks_like_title_selector": candidate[1] == 1,
                }
                for candidate in candidates[:5]
            ],
        )
        return candidates[0][9]

    def _find_combo_dropdown_item(self, window, combo, target_text: str, *, timeout: float = 3.0):
        owner_process_id = getattr(getattr(window, "element_info", None), "process_id", None)
        combo_rect = self._wrapper_rect(combo)
        allowed_types_by_pass = (
            {"listitem", "dataitem", "treeitem", "button", "custom"},
            {"text", "listitem", "dataitem", "treeitem", "button", "custom", "pane", "group"},
        )

        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            for allowed_types in allowed_types_by_pass:
                candidates = []
                seen_keys: set[tuple[object, int]] = set()
                search_sources = [self._iter_wrapper_tree(window)]
                if owner_process_id:
                    search_sources.append(self._iter_process_wrappers(int(owner_process_id)))
                for source_rank, wrappers in enumerate(search_sources):
                    for wrapper in wrappers:
                        handle = getattr(wrapper, "handle", None)
                        dedupe_token = handle if handle is not None else id(wrapper)
                        seen_key = (dedupe_token, source_rank)
                        if seen_key in seen_keys:
                            continue
                        seen_keys.add(seen_key)

                        control_type = self._wrapper_control_type(wrapper).lower()
                        if control_type not in allowed_types:
                            continue

                        text = self._read_control_value(wrapper) or self._wrapper_text(wrapper)
                        if not self._text_matches_target(text, target_text):
                            continue

                        rect = self._wrapper_rect(wrapper)
                        area = max(0, rect.right - rect.left) * max(0, rect.bottom - rect.top)
                        wrapper_process_id = getattr(getattr(wrapper, "element_info", None), "process_id", None)
                        candidates.append(
                            (
                                0 if owner_process_id and wrapper_process_id == owner_process_id else 1,
                                0 if rect.top >= combo_rect.top - 20 else 1,
                                abs(rect.top - combo_rect.bottom),
                                abs(rect.left - combo_rect.left),
                                source_rank,
                                -area,
                                wrapper,
                            )
                        )
                if candidates:
                    candidates.sort(key=lambda item: item[:6])
                    return candidates[0][6]
            time.sleep(0.1)
        raise AssertionError(f"HandBrake did not expose the preset item '{target_text}'.")

    def _select_combo_value(self, window, combo, target_text: str) -> None:
        current_text = self._selector_selected_text(combo)
        if self._text_matches_target(current_text, target_text):
            return

        selected_via_dropdown_item = False
        try:
            combo.select(target_text)
            time.sleep(HAND_BRAKE_DIALOG_SETTLE_SECONDS)
            current_text = self._selector_selected_text(combo)
            if self._text_matches_target(current_text, target_text):
                return
        except Exception:
            pass

        bring_window_to_front(window, keep_topmost=False)
        combo.click_input()
        time.sleep(0.3)

        try:
            dropdown_item = self._find_combo_dropdown_item(window, combo, target_text)
            logger.info("Selecting the HandBrake combo item: %s", self._wrapper_text(dropdown_item) or target_text)
            dropdown_item.click_input()
            selected_via_dropdown_item = True
            time.sleep(HAND_BRAKE_DIALOG_SETTLE_SECONDS)
        except AssertionError as exc:
            logger.info("HandBrake preset dropdown item lookup failed. Falling back to keyboard entry: %s", exc)
            self._set_combo_text_with_keyboard(window, combo, target_text)

        deadline = time.monotonic() + HAND_BRAKE_PRESET_SELECTION_TIMEOUT_SECONDS
        last_selected_text = self._selector_selected_text(combo)
        while time.monotonic() < deadline:
            last_selected_text = self._selector_selected_text(combo)
            if self._text_matches_target(last_selected_text, target_text):
                return
            if selected_via_dropdown_item and not last_selected_text:
                logger.info(
                    "HandBrake selector text stayed empty after clicking the target preset item. "
                    "Assuming the preset selection succeeded based on the dropdown click."
                )
                return
            time.sleep(0.2)
        raise AssertionError(
            f"HandBrake combo selection did not settle to '{target_text}'. "
            f"Last value: '{last_selected_text or '<empty>'}'."
        )

    def _select_preset(self, window, target_text: str) -> None:
        preset_combo = self._preset_combo(window)
        logger.info("Opening the HandBrake preset combo.")
        self._select_combo_value(window, preset_combo, target_text)
        selected_text = self._selector_selected_text(preset_combo)
        if selected_text:
            assert self._text_matches_target(selected_text, target_text), (
                f"HandBrake preset selection did not settle to '{target_text}'. Last value: '{selected_text or '<empty>'}'."
            )
        else:
            logger.info("HandBrake preset selector does not expose a readable selected-text value after selection.")

    def _save_path_edit(self, window):
        window_rect = self._wrapper_rect(window)
        window_width, window_height = self._window_dimensions(window)
        browse_buttons = self._find_matching_wrappers(
            window,
            HAND_BRAKE_DESTINATION_BROWSE_PATTERNS,
            control_types=("Button",),
        )
        if browse_buttons:
            candidates = []
            for browse_button in browse_buttons:
                browse_rect = self._wrapper_rect(browse_button)
                for wrapper in self._iter_wrapper_tree(window):
                    control_type = self._wrapper_control_type(wrapper).lower()
                    if control_type not in {"edit", "combobox"}:
                        continue
                    rect = self._wrapper_rect(wrapper)
                    width = rect.right - rect.left
                    if width < HAND_BRAKE_DESTINATION_EDIT_MIN_WIDTH:
                        continue
                    if rect.right > browse_rect.left + 20:
                        continue
                    overlap = self._vertical_overlap(rect, browse_rect)
                    if overlap <= 0:
                        continue
                    value = self._read_control_value(wrapper)
                    looks_like_path = any(token in value for token in (":\\", "\\", "/", ".mp4", ".m4v", ".mkv", ".webm"))
                    candidates.append(
                        (
                            0 if looks_like_path else 1,
                            -overlap,
                            -(rect.right - rect.left),
                            -rect.bottom,
                            rect.left,
                            wrapper,
                        )
                    )
            if candidates:
                candidates.sort(key=lambda item: item[:5])
                return candidates[0][5]

        candidates = []
        for wrapper in self._iter_wrapper_tree(window):
            control_type = self._wrapper_control_type(wrapper).lower()
            if control_type not in {"edit", "combobox"}:
                continue
            rect = self._wrapper_rect(wrapper)
            width = rect.right - rect.left
            if width < HAND_BRAKE_DESTINATION_EDIT_MIN_WIDTH:
                continue
            value = self._read_control_value(wrapper)
            looks_like_path = any(token in value for token in (":\\", "\\", "/", ".mp4", ".m4v", ".mkv", ".webm"))
            relative_top = (rect.top - window_rect.top) / window_height
            relative_bottom = (rect.bottom - window_rect.top) / window_height
            lower_half_penalty = 0 if relative_top >= 0.45 else 1
            bottom_band_penalty = 0 if relative_bottom >= 0.60 else 1
            candidates.append(
                (
                    lower_half_penalty,
                    bottom_band_penalty,
                    0 if looks_like_path else 1,
                    -width,
                    -rect.bottom,
                    rect.left,
                    wrapper,
                )
            )
        assert candidates, "HandBrake did not expose the bottom destination-path field."
        candidates.sort(key=lambda item: item[:6])
        return candidates[0][6]

    def _focus_edit_left_side(self, window, edit_wrapper) -> None:
        rect = self._wrapper_rect(edit_wrapper)
        main_rect = self._wrapper_rect(window)
        click_x_abs = rect.left + max(12, min(40, (rect.right - rect.left) // 12))
        click_y_abs = rect.top + max(1, (rect.bottom - rect.top) // 2)
        rel_x = max(1, min(click_x_abs - main_rect.left, (main_rect.right - main_rect.left) - 2))
        rel_y = max(1, min(click_y_abs - main_rect.top, (main_rect.bottom - main_rect.top) - 2))
        bring_window_to_front(window, keep_topmost=False)
        window.click_input(coords=(rel_x, rel_y))
        time.sleep(0.2)

    def _read_save_path(self, window) -> str:
        try:
            return self._wrapper_text(self._save_path_edit(window))
        except Exception:
            return ""

    def _destination_matches_input(self, save_path: str, input_video_path: Path) -> bool:
        if not save_path:
            return False
        normalized_save_path = save_path.casefold()
        normalized_input_name = input_video_path.name.casefold()
        normalized_input_stem = input_video_path.stem.casefold()
        if normalized_save_path.endswith(normalized_input_name):
            return True
        try:
            destination_path = Path(save_path)
        except Exception:
            return False
        destination_stem = destination_path.stem.casefold()
        destination_suffix = destination_path.suffix.casefold()
        return (
            destination_suffix in HAND_BRAKE_SOURCE_DESTINATION_SUFFIXES
            and bool(normalized_input_stem)
            and normalized_input_stem in destination_stem
        )

    def _resolve_start_encode_button(self, window):
        candidates = []
        window_rect = self._wrapper_rect(window)
        window_width, window_height = self._window_dimensions(window)
        for wrapper in self._find_matching_wrappers(window, HAND_BRAKE_START_BUTTON_PATTERNS, control_types=("Button",)):
            rect = self._wrapper_rect(wrapper)
            width = rect.right - rect.left
            height = rect.bottom - rect.top
            relative_top = (rect.top - window_rect.top) / window_height
            relative_left = (rect.left - window_rect.left) / window_width
            try:
                enabled = bool(wrapper.is_enabled())
            except Exception:
                enabled = True
            candidates.append(
                (
                    0 if enabled else 1,
                    0 if relative_top <= 0.45 else 1,
                    0 if relative_left >= 0.25 else 1,
                    relative_top,
                    -width * height,
                    rect.left,
                    wrapper,
                )
            )
        assert candidates, f"HandBrake did not expose a button matching {HAND_BRAKE_START_BUTTON_PATTERNS}."
        candidates.sort(key=lambda item: item[:6])
        return candidates[0][6]

    def _start_encode_button_ready(self, window) -> bool:
        try:
            button = self._resolve_start_encode_button(window)
        except Exception:
            return False
        try:
            return bool(button.is_enabled())
        except Exception:
            return True

    def _source_name_visible(self, window, input_video_path: Path) -> bool:
        target_names = [
            input_video_path.name.casefold(),
            input_video_path.stem.casefold(),
        ]
        target_names = [value for value in target_names if len(value) >= HAND_BRAKE_SOURCE_NAME_MIN_TEXT_LENGTH]
        if not target_names:
            return False
        for wrapper in self._iter_wrapper_tree(window):
            text = self._read_control_value(wrapper).casefold()
            if not text:
                continue
            if any(target_name in text for target_name in target_names):
                return True
        return False

    def _source_loaded(self, window, input_video_path: Path) -> bool:
        save_path = self._read_save_path(window)
        if save_path and self._destination_matches_input(save_path, input_video_path):
            return True
        ready_status = self._status_line(window, HAND_BRAKE_READY_STATUS_PATTERNS)
        if not ready_status:
            return False
        if self._start_encode_button_ready(window):
            return True
        return self._source_name_visible(window, input_video_path)

    def _wait_for_source_loaded(self, input_video_path: Path, timeout: float = HAND_BRAKE_SOURCE_LOAD_TIMEOUT_SECONDS) -> None:
        deadline = time.monotonic() + timeout
        last_save_path = ""
        last_ready_status = ""
        last_start_button_ready = False
        while time.monotonic() < deadline:
            try:
                window = self._connect_main_window(timeout=1.0)
            except UiAutomationError:
                time.sleep(0.2)
                continue
            last_save_path = self._read_save_path(window)
            last_ready_status = self._status_line(window, HAND_BRAKE_READY_STATUS_PATTERNS)
            last_start_button_ready = self._start_encode_button_ready(window)
            if self._source_loaded(window, input_video_path):
                return
            time.sleep(0.3)
        raise AssertionError(
            f"HandBrake did not finish loading the selected source: {input_video_path}. "
            f"Last observed state: save_path={last_save_path or '<empty>'} "
            f"ready_status={last_ready_status or '<not-ready>'} "
            f"start_button_ready={last_start_button_ready}"
        )

    def _open_input_clip(self, window, input_video_path: Path) -> None:
        normalized_input_path = input_video_path.resolve(strict=False)
        start_button = None
        try:
            start_button = self._find_button(window, HAND_BRAKE_SELECT_FILE_PATTERNS)
        except Exception:
            start_button = None
        if start_button is None:
            start_button = self._find_button(window, HAND_BRAKE_OPEN_SOURCE_PATTERNS)
        logger.info("Clicking the HandBrake source-selection button: %s", self._wrapper_text(start_button) or "<untitled>")
        start_button.click_input()
        time.sleep(HAND_BRAKE_DIALOG_SETTLE_SECONDS)
        logger.info(
            "Selecting the HandBrake source clip through the standard Windows file dialog: %s",
            normalized_input_path,
        )
        fill_file_dialog(
            normalized_input_path,
            dialog_patterns=OPEN_DIALOG_PATTERNS,
            confirm_patterns=HAND_BRAKE_DIALOG_OPEN_PATTERNS,
            timeout=HAND_BRAKE_DIALOG_TIMEOUT_SECONDS,
            must_exist=True,
            allow_direct_selection=False,
        )
        self._wait_for_source_loaded(normalized_input_path)

    def _click_tab(self, window, patterns: tuple[str, ...]) -> None:
        candidates = []
        for wrapper in self._find_matching_wrappers(window, patterns, control_types=("TabItem",)):
            rect = self._wrapper_rect(wrapper)
            candidates.append(((rect.top, rect.left), wrapper))
        assert candidates, f"HandBrake did not expose a tab matching {patterns}."
        candidates.sort(key=lambda item: item[0])
        tab = candidates[0][1]
        logger.info("Clicking the HandBrake tab: %s", self._wrapper_text(tab) or "<untitled>")
        tab.click_input()
        time.sleep(HAND_BRAKE_DIALOG_SETTLE_SECONDS)

    def _summary_contains_x264(self, window) -> bool:
        for wrapper in self._iter_wrapper_tree(window):
            if self._wrapper_control_type(wrapper).lower() != "text":
                continue
            title = self._wrapper_text(wrapper)
            if title and self._matches_patterns(title, HAND_BRAKE_X264_PATTERNS):
                return True
        return False

    def _video_encoder_combo(self, window):
        try:
            return self._find_labeled_combo(window, HAND_BRAKE_VIDEO_ENCODER_LABEL_PATTERNS)
        except AssertionError:
            pass

        window_rect = self._wrapper_rect(window)
        _, window_height = self._window_dimensions(window)
        candidates = []
        for wrapper in self._iter_wrapper_tree(window):
            if self._wrapper_control_type(wrapper).lower() != "combobox":
                continue
            rect = self._wrapper_rect(wrapper)
            width = rect.right - rect.left
            if width <= 0:
                continue
            relative_top = (rect.top - window_rect.top) / window_height
            candidates.append(
                (
                    0 if 0.20 <= relative_top <= 0.75 else 1,
                    relative_top,
                    -width,
                    rect.left,
                    wrapper,
                )
            )
        assert candidates, "HandBrake did not expose the video-encoder combo box."
        candidates.sort(key=lambda item: item[:4])
        return candidates[0][4]

    def _ensure_x264_encoder(self, window) -> None:
        self._click_tab(window, HAND_BRAKE_SUMMARY_TAB_PATTERNS)
        if self._summary_contains_x264(window):
            logger.info("HandBrake summary already reports an x264 encoder.")
            return

        logger.info("HandBrake summary did not report x264. Trying the Video tab fallback.")
        self._click_tab(window, HAND_BRAKE_VIDEO_TAB_PATTERNS)
        video_combo = self._video_encoder_combo(window)
        bring_window_to_front(window, keep_topmost=False)
        video_combo.click_input()
        time.sleep(0.3)
        send_hotkey("{HOME}")
        time.sleep(0.1)
        send_hotkey("{ENTER}")
        time.sleep(HAND_BRAKE_DIALOG_SETTLE_SECONDS)

        self._click_tab(window, HAND_BRAKE_SUMMARY_TAB_PATTERNS)
        assert self._summary_contains_x264(window), "HandBrake did not resolve the selected encoder to x264."

    def _set_output_path(self, window, output_video_path: Path) -> None:
        save_edit = self._save_path_edit(window)
        self._focus_edit_left_side(window, save_edit)
        paste_text_via_clipboard(str(output_video_path), replace_existing=True)
        time.sleep(0.5)
        current_value = self._wrapper_text(save_edit)
        assert current_value.casefold() == str(output_video_path).casefold(), (
            f"HandBrake destination path did not update correctly. expected={output_video_path} actual={current_value}"
        )

    def _handle_overwrite_prompt(self, window, *, timeout: float = 3.0) -> None:
        accepted = accept_overwrite_confirmation(timeout=timeout, owner_window=window, poll_interval=0.05)
        if accepted:
            return

        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            prompt_detected = False
            for wrapper in self._iter_wrapper_tree(window):
                title = self._wrapper_text(wrapper)
                if not title or not self._matches_patterns(title, HAND_BRAKE_OVERWRITE_TEXT_PATTERNS):
                    continue
                prompt_detected = True
                break
            if not prompt_detected:
                time.sleep(0.1)
                continue

            for wrapper in self._find_matching_wrappers(window, HAND_BRAKE_YES_PATTERNS, control_types=("Button",)):
                logger.info("Accepting the HandBrake overwrite prompt.")
                wrapper.click_input()
                time.sleep(0.4)
                return
            time.sleep(0.1)

    def _start_encode(self, window, output_video_path: Path) -> None:
        start_button = self._resolve_start_encode_button(window)
        logger.info("Clicking the HandBrake start-encode button: %s", self._wrapper_text(start_button) or "<untitled>")
        start_button.click_input()
        self._handle_overwrite_prompt(window)

        deadline = time.monotonic() + HAND_BRAKE_START_TIMEOUT_SECONDS
        while time.monotonic() < deadline:
            if output_video_path.exists():
                return
            if self._status_line(window, HAND_BRAKE_ACTIVE_STATUS_PATTERNS):
                return
            time.sleep(0.5)
        raise AssertionError(f"HandBrake did not start encoding the requested output: {output_video_path}")

    def _status_line(self, window, patterns: tuple[str, ...]) -> str:
        for wrapper in self._iter_wrapper_tree(window):
            if self._wrapper_control_type(wrapper).lower() != "text":
                continue
            title = self._wrapper_text(wrapper)
            if title and self._matches_patterns(title, patterns):
                return title
        return ""

    def _wait_for_export_completion(self, output_video_path: Path, window=None) -> None:
        assert output_video_path.parent.exists(), f"Output directory disappeared: {output_video_path.parent}"

        stable_rounds = 0
        last_size = -1
        deadline = time.monotonic() + HAND_BRAKE_EXPORT_TIMEOUT_SECONDS
        progress_logger = _WaitProgressLogger("HandBrake encode")
        while time.monotonic() < deadline:
            if window is None:
                window = self._connect_main_window(timeout=2.0)
            current_size = output_video_path.stat().st_size if output_video_path.exists() else 0
            try:
                active_status = self._status_line(window, HAND_BRAKE_ACTIVE_STATUS_PATTERNS)
                ready_status = self._status_line(window, HAND_BRAKE_READY_STATUS_PATTERNS)
            except Exception:
                window = self._connect_main_window(timeout=2.0)
                active_status = self._status_line(window, HAND_BRAKE_ACTIVE_STATUS_PATTERNS)
                ready_status = self._status_line(window, HAND_BRAKE_READY_STATUS_PATTERNS)

            if current_size > 0 and current_size == last_size:
                stable_rounds += 1
            else:
                stable_rounds = 0
            last_size = current_size
            progress_logger.maybe_log(
                detail=(
                    f"output={output_video_path.name} size_bytes={current_size} stable_rounds={stable_rounds} "
                    f"active_status={active_status or '<idle>'} ready_status={ready_status or '<not-ready>'}"
                )
            )

            if current_size > 0 and not active_status and stable_rounds >= HAND_BRAKE_EXPORT_STABLE_ROUNDS:
                logger.info("HandBrake export completion conditions satisfied. elapsed=%s", progress_logger.elapsed_text())
                break
            time.sleep(HAND_BRAKE_EXPORT_POLL_SECONDS)
        else:
            raise AssertionError(f"HandBrake export did not finish within timeout: {output_video_path}")

        assert output_video_path.exists() and output_video_path.stat().st_size > 0, (
            f"HandBrake export output was not created correctly: {output_video_path}"
        )


class ShutterEncoderOperator(SoftwareOperator):
    def _wrapper_text(self, wrapper) -> str:
        try:
            return (wrapper.window_text() or "").strip()
        except Exception:
            return ""

    def _wrapper_control_type(self, wrapper) -> str:
        try:
            return (getattr(wrapper.element_info, "control_type", "") or "").strip()
        except Exception:
            return ""

    def _wrapper_rect(self, wrapper):
        return wrapper.rectangle()

    def _iter_wrapper_tree(self, root):
        yield root
        for child in root.descendants():
            yield child

    def _list_shutter_windows(self):
        from pywinauto import Desktop

        candidates = []
        for window in Desktop(backend="uia").windows():
            title = self._wrapper_text(window)
            if not title or not re.search(self.profile.main_window_title_re, title, re.IGNORECASE):
                continue
            class_name = (getattr(window.element_info, "class_name", "") or "").strip()
            if class_name != "SunAwtFrame":
                continue
            try:
                rect = self._wrapper_rect(window)
                area = max(0, rect.right - rect.left) * max(0, rect.bottom - rect.top)
            except Exception:
                area = 0
            pid = int(getattr(window.element_info, "process_id", 0) or 0)
            handle = int(getattr(window, "handle", 0) or 0)
            candidates.append(((area, pid, handle), window))
        candidates.sort(key=lambda item: item[0], reverse=True)
        return [window for _, window in candidates]

    def _connect_main_window(self, timeout: float = 15.0):
        deadline = time.monotonic() + timeout
        last_error: Exception | None = None
        while time.monotonic() < deadline:
            windows = self._list_shutter_windows()
            if windows:
                return windows[0]
            last_error = UiAutomationError("No Shutter Encoder main window is currently visible.")
            time.sleep(0.25)
        raise UiAutomationError("Could not connect to the Shutter Encoder main window.") from last_error

    def _process_id(self, window) -> int | None:
        try:
            process_id = getattr(getattr(window, "element_info", None), "process_id", None)
        except Exception:
            return None
        return int(process_id) if process_id else None

    def _is_visible(self, window) -> bool:
        try:
            return bool(window.is_visible())
        except Exception:
            return True

    def _iter_process_top_level_windows(self, process_id: int | None):
        if process_id is None:
            return []
        from pywinauto import Desktop

        matches = []
        for window in Desktop(backend="uia").windows():
            if self._process_id(window) != process_id:
                continue
            if not self._is_visible(window):
                continue
            try:
                rect = self._wrapper_rect(window)
                area = max(0, rect.right - rect.left) * max(0, rect.bottom - rect.top)
            except Exception:
                area = 0
            matches.append((area, self._wrapper_text(window).casefold(), window))
        matches.sort(key=lambda item: (-item[0], item[1]))
        return [item[-1] for item in matches]

    def _close_remaining_process_windows(self, process_id: int | None) -> None:
        remaining_windows = self._iter_process_top_level_windows(process_id)
        if not remaining_windows:
            return
        logger.info(
            "Closing remaining Shutter Encoder top-level windows: %s",
            [self._wrapper_text(candidate) or "<untitled>" for candidate in remaining_windows[:5]],
        )
        for candidate in remaining_windows:
            try:
                bring_window_to_front(candidate, keep_topmost=False)
            except Exception:
                pass
            try:
                candidate.close()
            except Exception:
                try:
                    request_window_close(candidate)
                except Exception:
                    pass
            try:
                dismiss_close_prompts(timeout=0.5, owner_window=candidate)
            except Exception:
                pass

    def _launch_application(self) -> None:
        shortcut_path = SOFTWARE_SPECS["shutter_encoder"].launch_path
        assert shortcut_path.exists(), f"Shutter Encoder shortcut does not exist: {shortcut_path}"
        logger.info("Launching Shutter Encoder from shortcut: %s", shortcut_path)
        os.startfile(str(shortcut_path))
        time.sleep(2.0)

    def _matches_patterns(self, text: str, patterns: tuple[str, ...]) -> bool:
        return any(re.search(pattern, text, re.IGNORECASE) for pattern in patterns)

    def _read_control_value(self, wrapper) -> str:
        readers = (
            lambda: wrapper.get_value(),
            lambda: wrapper.iface_value.CurrentValue,
            lambda: wrapper.window_text(),
            lambda: wrapper.texts()[0] if wrapper.texts() else "",
        )
        for reader in readers:
            try:
                value = reader()
            except Exception:
                continue
            if isinstance(value, str):
                return value.strip()
        return ""

    def _is_enabled(self, wrapper) -> bool | None:
        readers = (
            lambda: bool(wrapper.is_enabled()),
            lambda: bool(getattr(wrapper.element_info, "enabled")),
        )
        for reader in readers:
            try:
                return reader()
            except Exception:
                continue
        return None

    def _first_matching_control(self, root, patterns: tuple[str, ...], *, control_types: tuple[str, ...] = ()):
        allowed_types = {control_type.lower() for control_type in control_types}
        for wrapper in self._iter_wrapper_tree(root):
            control_type = self._wrapper_control_type(wrapper).lower()
            if allowed_types and control_type not in allowed_types:
                continue
            texts = (self._wrapper_text(wrapper), self._read_control_value(wrapper))
            if any(text and self._matches_patterns(text, patterns) for text in texts):
                return wrapper
        return None

    def _control_present(self, root, patterns: tuple[str, ...], *, control_types: tuple[str, ...] = ()) -> bool:
        return self._first_matching_control(root, patterns, control_types=control_types) is not None

    def _read_progress_ratio(self, window) -> float | None:
        best_candidate = None
        for wrapper in self._iter_wrapper_tree(window):
            if self._wrapper_control_type(wrapper).lower() != "progressbar":
                continue
            try:
                rect = self._wrapper_rect(wrapper)
            except Exception:
                continue
            current_value = None
            maximum_value = None
            minimum_value = 0.0
            for reader in (
                lambda: float(wrapper.iface_range_value.CurrentValue),
                lambda: float(wrapper.get_value()),
            ):
                try:
                    current_value = reader()
                    break
                except Exception:
                    continue
            try:
                maximum_value = float(wrapper.iface_range_value.CurrentMaximum)
            except Exception:
                maximum_value = None
            try:
                minimum_value = float(wrapper.iface_range_value.CurrentMinimum)
            except Exception:
                minimum_value = 0.0
            if current_value is None or maximum_value is None or maximum_value <= minimum_value:
                continue
            ratio = (current_value - minimum_value) / (maximum_value - minimum_value)
            if not 0.0 <= ratio <= 1.05:
                continue
            score = ((rect.right - rect.left) * (rect.bottom - rect.top), rect.bottom, rect.left)
            clamped_ratio = max(0.0, min(ratio, 1.0))
            if best_candidate is None or score > best_candidate[0]:
                best_candidate = (score, clamped_ratio)
        if best_candidate is None:
            return None
        return best_candidate[1]

    def _snapshot_output_candidates(self, input_video_path: Path) -> dict[Path, tuple[float, int]]:
        snapshot: dict[Path, tuple[float, int]] = {}
        for candidate in input_video_path.parent.iterdir():
            if not candidate.is_file() or candidate.suffix.lower() not in SHUTTER_ENCODER_OUTPUT_SUFFIXES:
                continue
            resolved = candidate.resolve(strict=False)
            stat = candidate.stat()
            snapshot[resolved] = (stat.st_mtime, stat.st_size)
        return snapshot

    def _resolve_output_candidate(self, input_video_path: Path, known_outputs: dict[Path, tuple[float, int]]) -> Path | None:
        candidates = []
        input_resolved = input_video_path.resolve(strict=False)
        for candidate in input_video_path.parent.iterdir():
            if not candidate.is_file() or candidate.suffix.lower() not in SHUTTER_ENCODER_OUTPUT_SUFFIXES:
                continue
            resolved = candidate.resolve(strict=False)
            if resolved == input_resolved:
                continue
            stat = candidate.stat()
            previous = known_outputs.get(resolved)
            changed = previous is None or stat.st_mtime > previous[0] or stat.st_size != previous[1]
            if not changed or stat.st_size <= 0:
                continue
            score = (1 if input_video_path.stem.casefold() in candidate.stem.casefold() else 0, stat.st_mtime, stat.st_size)
            candidates.append((score, candidate))
        if not candidates:
            return None
        candidates.sort(key=lambda item: item[0], reverse=True)
        return candidates[0][1]

    def _looks_like_update_dialog(self, window) -> bool:
        if window is None:
            return False
        title = self._wrapper_text(window)
        has_dialog_title = bool(title and self._matches_patterns(title, SHUTTER_ENCODER_UPDATE_DIALOG_PATTERNS))
        if not has_dialog_title:
            return False
        has_negative_action = self._control_present(
            window,
            SHUTTER_ENCODER_UPDATE_NO_PATTERNS,
            control_types=("Button", "Text", "Pane", "Group"),
        )
        has_prompt = self._control_present(
            window,
            SHUTTER_ENCODER_UPDATE_PROMPT_PATTERNS,
            control_types=("Text", "Pane", "Group"),
        )
        has_version = self._control_present(
            window,
            SHUTTER_ENCODER_UPDATE_VERSION_PATTERNS,
            control_types=("Text", "Pane", "Group"),
        )
        return has_negative_action or has_prompt or has_version

    def _click_wrapper_center(self, wrapper) -> None:
        try:
            wrapper.click_input()
            time.sleep(SHUTTER_ENCODER_SETTLE_SECONDS)
            return
        except Exception as exc:
            logger.info("Direct click failed for '%s'. Falling back to a center-coordinate click: %s", self._wrapper_text(wrapper) or "<untitled>", exc)
        try:
            top_level = wrapper.top_level_parent()
        except Exception:
            top_level = wrapper
        rect = self._wrapper_rect(wrapper)
        top_rect = self._wrapper_rect(top_level)
        click_x_abs = rect.left + max(1, (rect.right - rect.left) // 2)
        click_y_abs = rect.top + max(1, (rect.bottom - rect.top) // 2)
        rel_x = max(1, min(click_x_abs - top_rect.left, (top_rect.right - top_rect.left) - 2))
        rel_y = max(1, min(click_y_abs - top_rect.top, (top_rect.bottom - top_rect.top) - 2))
        bring_window_to_front(top_level, keep_topmost=False)
        top_level.click_input(coords=(rel_x, rel_y))
        time.sleep(SHUTTER_ENCODER_SETTLE_SECONDS)

    def _find_update_dialog(self):
        for main_window in self._list_shutter_windows():
            for child in main_window.children():
                if self._looks_like_update_dialog(child):
                    return child
        try:
            foreground = get_foreground_window()
            if self._looks_like_update_dialog(foreground):
                return foreground
        except UiAutomationError:
            pass

        try:
            dialog = wait_for_window(SHUTTER_ENCODER_UPDATE_DIALOG_PATTERNS, timeout=0.2)
            if self._looks_like_update_dialog(dialog):
                return dialog
        except UiAutomationError:
            pass
        raise UiAutomationError("No Shutter Encoder update dialog is currently visible.")

    def _dismiss_update_dialog_with_hotkey(self) -> bool:
        try:
            target = get_foreground_window()
            target_title = self._wrapper_text(target)
            if not self._matches_patterns(target_title, SHUTTER_ENCODER_UPDATE_DIALOG_PATTERNS):
                logger.info(
                    "Skipping the Shutter Encoder update-dialog fallback key because the foreground window does not look like an update dialog: %s",
                    target_title or "<untitled>",
                )
                return False
            logger.info(
                "Update dialog was not located via UIA. Sending fallback key '%s' to the foreground window: %s",
                SHUTTER_ENCODER_UPDATE_DIRECT_KEY,
                target_title or "<untitled>",
            )
            bring_window_to_front(target, keep_topmost=False)
        except UiAutomationError:
            logger.info(
                "Update dialog was not located via UIA and the foreground window could not be resolved. "
                "Skipping fallback key '%s'.",
                SHUTTER_ENCODER_UPDATE_DIRECT_KEY,
            )
            return False
        send_hotkey(SHUTTER_ENCODER_UPDATE_DIRECT_KEY)
        time.sleep(SHUTTER_ENCODER_SETTLE_SECONDS)
        return True

    def _dismiss_update_dialog_if_present(self, *, timeout: float = SHUTTER_ENCODER_UPDATE_TIMEOUT_SECONDS) -> None:
        logger.info("Checking for a Shutter Encoder update dialog.")
        deadline = time.monotonic() + timeout
        probe_deadline = min(deadline, time.monotonic() + SHUTTER_ENCODER_UPDATE_UIA_PROBE_SECONDS)
        while time.monotonic() < probe_deadline:
            try:
                dialog = self._find_update_dialog()
                break
            except UiAutomationError:
                time.sleep(0.1)
        else:
            if not self._dismiss_update_dialog_with_hotkey():
                logger.info("No Shutter Encoder update dialog detected.")
                return
            while time.monotonic() < deadline:
                try:
                    dialog = self._find_update_dialog()
                    break
                except UiAutomationError:
                    time.sleep(0.1)
            else:
                logger.info("No Shutter Encoder update dialog detected.")
                return
        logger.info("Shutter Encoder update dialog detected: %s", dialog.window_text() or "<untitled>")
        bring_window_to_front(dialog, keep_topmost=False)
        try:
            no_button = find_text_control(dialog, SHUTTER_ENCODER_UPDATE_NO_PATTERNS, control_types=("Button", "Text", "Pane", "Group"))
            logger.info("Clicking the Shutter Encoder update dialog negative action: %s", self._wrapper_text(no_button) or "<untitled>")
            self._click_wrapper_center(no_button)
        except UiAutomationError:
            logger.info("The Shutter Encoder update dialog negative action was not exposed through UIA. Falling back to key '%s'.", SHUTTER_ENCODER_UPDATE_DIRECT_KEY)
            self._dismiss_update_dialog_with_hotkey()

    def _find_browse_control(self, window):
        last_error: Exception | None = None
        search_windows = [window]
        try:
            refreshed = self._connect_main_window(timeout=1.0)
            if all(getattr(candidate, "handle", None) != getattr(refreshed, "handle", None) for candidate in search_windows):
                search_windows.append(refreshed)
        except UiAutomationError:
            pass

        for candidate_window in search_windows:
            try:
                return wait_for_text_control(
                    candidate_window,
                    SHUTTER_ENCODER_BROWSE_PATTERNS,
                    control_types=SHUTTER_ENCODER_BROWSE_CONTROL_TYPES,
                    timeout=1.0,
                    poll_interval=0.1,
                )
            except UiAutomationError as exc:
                last_error = exc

        try:
            from pywinauto import Desktop

            logger.info("Searching the desktop for a Browse control owned by the Shutter Encoder window.")
            candidates = []
            for desktop_window in Desktop(backend="uia").windows():
                title = self._wrapper_text(desktop_window)
                if not title or not re.search(self.profile.main_window_title_re, title, re.IGNORECASE):
                    continue
                for wrapper in self._iter_wrapper_tree(desktop_window):
                    control_type = self._wrapper_control_type(wrapper).lower()
                    if control_type not in {value.lower() for value in SHUTTER_ENCODER_BROWSE_CONTROL_TYPES}:
                        continue
                    text = self._wrapper_text(wrapper)
                    if not text or not self._matches_patterns(text, SHUTTER_ENCODER_BROWSE_PATTERNS):
                        continue
                    rect = self._wrapper_rect(wrapper)
                    candidates.append((rect.top, rect.left, wrapper, title))
            if candidates:
                candidates.sort(key=lambda item: (item[0], item[1]))
                _, _, browse_control, top_level_title = candidates[0]
                logger.info(
                    "Resolved the Shutter Encoder Browse control through filtered desktop search. top_level=%s",
                    top_level_title,
                )
                return browse_control
            raise UiAutomationError("Could not find a Shutter Encoder-owned Browse control through desktop search.")
        except Exception as exc:
            last_error = exc

        if last_error is not None:
            raise last_error
        raise UiAutomationError("Could not locate the Shutter Encoder Browse control.")

    def _button_enabled(self, root, patterns: tuple[str, ...]) -> bool:
        button = self._first_matching_control(root, patterns, control_types=("Button", "Text", "Pane", "Group", "Custom"))
        if button is None:
            return False
        enabled = self._is_enabled(button)
        return enabled is True

    def _wait_for_imported_source(self, window, input_video_path: Path) -> None:
        deadline = time.monotonic() + max(SHUTTER_ENCODER_IMPORT_TIMEOUT_SECONDS, SHUTTER_ENCODER_CONTROL_TIMEOUT_SECONDS)
        last_status = None
        while time.monotonic() < deadline:
            try:
                window = self._connect_main_window(timeout=2.0)
            except UiAutomationError:
                time.sleep(0.2)
                continue

            file_name_present = self._control_present(
                window,
                (re.escape(input_video_path.name), re.escape(input_video_path.stem)),
                control_types=SHUTTER_ENCODER_SOURCE_CONTROL_TYPES,
            )
            imported_path_present = self._control_present(
                window,
                SHUTTER_ENCODER_IMPORTED_PATH_PATTERNS,
                control_types=("Text", "Pane", "Group", "Custom", "ListItem", "DataItem", "TreeItem"),
            )
            positive_count_present = self._control_present(
                window,
                SHUTTER_ENCODER_FILE_COUNT_PATTERNS,
                control_types=("Text", "Pane", "Group", "Custom"),
            )
            placeholder_present = self._control_present(
                window,
                SHUTTER_ENCODER_DROP_PLACEHOLDER_PATTERNS,
                control_types=("Text", "Pane", "Group", "Custom"),
            )
            clear_enabled = self._button_enabled(window, SHUTTER_ENCODER_CLEAR_PATTERNS)

            status = (file_name_present, imported_path_present, positive_count_present, placeholder_present, clear_enabled)
            if status != last_status:
                logger.info(
                    "Shutter Encoder import state. file_name_present=%s imported_path_present=%s positive_count_present=%s placeholder_present=%s clear_enabled=%s",
                    file_name_present,
                    imported_path_present,
                    positive_count_present,
                    placeholder_present,
                    clear_enabled,
                )
                last_status = status

            if file_name_present or imported_path_present or positive_count_present or clear_enabled:
                logger.info("Shutter Encoder import completion conditions satisfied.")
                return

            time.sleep(0.2)

        logger.info(
            "Shutter Encoder did not expose the imported source '%s' through UIA after the file dialog closed. "
            "Continuing with the next step and letting the real start/progress checks validate the import.",
            input_video_path.name,
        )

    def _click_browse_fallback_area(self, window) -> None:
        try:
            top_level = window.top_level_parent()
        except Exception:
            top_level = window
        try:
            self._click_window_design_point_with_pyautogui(
                top_level,
                SHUTTER_ENCODER_BROWSE_BUTTON_DESIGN_X,
                SHUTTER_ENCODER_BROWSE_BUTTON_DESIGN_Y,
                "Browse button design point",
            )
        except Exception as exc:
            logger.info(
                "The Shutter Encoder Browse design-point fallback failed. Falling back to window-relative coordinates: %s",
                exc,
            )
            self._click_window_ratio_with_pyautogui(
                top_level,
                SHUTTER_ENCODER_BROWSE_CLICK_X_RATIO,
                SHUTTER_ENCODER_BROWSE_CLICK_Y_RATIO,
                "Browse button",
            )

    def _window_scale_factor(self, window) -> float:
        handle = getattr(window, "handle", None)
        if not handle:
            return 1.0
        try:
            scale = ctypes.windll.user32.GetDpiForWindow(handle) / 96.0
        except Exception:
            return 1.0
        return scale if scale > 0 else 1.0

    def _top_level_screen_metrics(self, window) -> tuple[int, int, int, int]:
        rect = self._wrapper_rect(window)
        scale = self._window_scale_factor(window)
        left = int(round(rect.left / scale))
        top = int(round(rect.top / scale))
        right = int(round(rect.right / scale))
        bottom = int(round(rect.bottom / scale))
        return left, top, max(1, right - left), max(1, bottom - top)

    def _click_window_ratio(self, window, x_ratio: float, y_ratio: float, description: str) -> None:
        try:
            top_level = window.top_level_parent()
        except Exception:
            top_level = window
        bring_window_to_front(top_level, keep_topmost=False)
        left, top, width, height = self._top_level_screen_metrics(top_level)
        click_x_abs = left + int(width * x_ratio)
        click_y_abs = top + int(height * y_ratio)
        rel_x = max(1, min(click_x_abs - left, width - 2))
        rel_y = max(1, min(click_y_abs - top, height - 2))
        logger.info(
            "Clicking the Shutter Encoder %s fallback area at window-relative ratios x=%.3f y=%.3f.",
            description,
            x_ratio,
            y_ratio,
        )
        top_level.click_input(coords=(rel_x, rel_y))
        time.sleep(SHUTTER_ENCODER_SETTLE_SECONDS)

    def _import_pyautogui(self):
        try:
            import pyautogui
        except ImportError as exc:
            raise UiAutomationError("pyautogui is not installed. Run: pip install pyautogui") from exc
        pyautogui.FAILSAFE = False
        pyautogui.PAUSE = 0.05
        return pyautogui

    def _click_window_ratio_with_pyautogui(self, window, x_ratio: float, y_ratio: float, description: str) -> None:
        pyautogui = self._import_pyautogui()
        try:
            top_level = window.top_level_parent()
        except Exception:
            top_level = window
        bring_window_to_front(top_level, keep_topmost=False)
        rect = self._wrapper_rect(top_level)
        width = max(1, rect.right - rect.left)
        height = max(1, rect.bottom - rect.top)
        click_x = rect.left + int(width * x_ratio)
        click_y = rect.top + int(height * y_ratio)
        logger.info(
            "Clicking the Shutter Encoder %s fallback area with pyautogui at screen coordinates x=%d y=%d.",
            description,
            click_x,
            click_y,
        )
        pyautogui.moveTo(click_x, click_y, duration=0.05)
        pyautogui.click(click_x, click_y)
        time.sleep(SHUTTER_ENCODER_SETTLE_SECONDS)

    def _click_window_design_point_with_pyautogui(self, window, x_design: int, y_design: int, description: str) -> None:
        try:
            top_level = window.top_level_parent()
        except Exception:
            top_level = window
        bring_window_to_front(top_level, keep_topmost=False)
        rect = self._wrapper_rect(top_level)
        scale = self._window_scale_factor(top_level)
        click_x = rect.left + int(round(x_design * scale))
        click_y = rect.top + int(round(y_design * scale))
        logger.info(
            "Clicking the Shutter Encoder %s with pyautogui at design coordinates x=%d y=%d scaled to screen coordinates x=%d y=%d.",
            description,
            x_design,
            y_design,
            click_x,
            click_y,
        )
        self._click_screen_point_with_pyautogui(click_x, click_y, description)

    def _click_screen_point_with_pyautogui(self, x: int, y: int, description: str) -> None:
        pyautogui = self._import_pyautogui()
        logger.info(
            "Clicking the Shutter Encoder %s with pyautogui at screen coordinates x=%d y=%d.",
            description,
            x,
            y,
        )
        pyautogui.moveTo(x, y, duration=0.05)
        pyautogui.click(x, y)
        time.sleep(SHUTTER_ENCODER_SETTLE_SECONDS)

    def _click_top_right_offset_with_pyautogui(self, window, right_offset: int, top_offset: int, description: str) -> None:
        try:
            top_level = window.top_level_parent()
        except Exception:
            top_level = window
        bring_window_to_front(top_level, keep_topmost=False)
        rect = self._wrapper_rect(top_level)
        scale = self._window_scale_factor(top_level)
        click_x = rect.right - int(round(right_offset * scale))
        click_y = rect.top + int(round(top_offset * scale))
        logger.info(
            "Clicking the Shutter Encoder %s near the custom title bar close area at screen coordinates x=%d y=%d.",
            description,
            click_x,
            click_y,
        )
        self._click_screen_point_with_pyautogui(click_x, click_y, description)

    def _wrapper_center_screen_point(self, wrapper) -> tuple[int, int]:
        rect = self._wrapper_rect(wrapper)
        center_x = rect.left + max(1, (rect.right - rect.left) // 2)
        center_y = rect.top + max(1, (rect.bottom - rect.top) // 2)
        return center_x, center_y

    def _predict_start_button_screen_point_from_picker(self, window) -> tuple[int, int]:
        picker = self._find_function_picker(window)
        rect = self._wrapper_rect(picker)
        picker_height = max(1, rect.bottom - rect.top)
        scale = picker_height / SHUTTER_ENCODER_FUNCTION_PICKER_DESIGN_HEIGHT
        picker_center_x, picker_center_y = self._wrapper_center_screen_point(picker)
        click_x = picker_center_x
        click_y = picker_center_y + int(round(SHUTTER_ENCODER_START_FROM_PICKER_CENTER_Y_OFFSET * scale))
        logger.info(
            "Predicted the Shutter Encoder Start function button position from the function picker. "
            "picker_center=(%d,%d) predicted_button_center=(%d,%d) scale=%.3f",
            picker_center_x,
            picker_center_y,
            click_x,
            click_y,
            scale,
        )
        return click_x, click_y

    def _find_start_button_candidate_near_picker(self, window):
        target_x, target_y = self._predict_start_button_screen_point_from_picker(window)
        try:
            picker = self._find_function_picker(window)
            picker_rect = self._wrapper_rect(picker)
            scale = max(0.75, (picker_rect.bottom - picker_rect.top) / SHUTTER_ENCODER_FUNCTION_PICKER_DESIGN_HEIGHT)
        except Exception:
            scale = 1.0

        candidates = []
        for wrapper in self._iter_wrapper_tree(window):
            control_type = self._wrapper_control_type(wrapper).lower()
            if control_type not in {"button", "text", "pane", "group", "custom"}:
                continue
            enabled = self._is_enabled(wrapper)
            if enabled is False:
                continue
            rect = self._wrapper_rect(wrapper)
            width = rect.right - rect.left
            height = rect.bottom - rect.top
            if width < int(90 * scale) or width > int(240 * scale):
                continue
            if height < int(14 * scale) or height > int(40 * scale):
                continue
            center_x, center_y = self._wrapper_center_screen_point(wrapper)
            dx = abs(center_x - target_x)
            dy = abs(center_y - target_y)
            if dx > int(85 * scale) or dy > int(40 * scale):
                continue
            score = (
                0 if control_type == "button" else 1,
                dy,
                dx,
                -width,
            )
            candidates.append((score, wrapper))

        if not candidates:
            raise UiAutomationError("Could not locate a geometry-based Shutter Encoder Start function button candidate.")

        candidates.sort(key=lambda item: item[0])
        _, candidate = candidates[0]
        candidate_center_x, candidate_center_y = self._wrapper_center_screen_point(candidate)
        logger.info(
            "Resolved a geometry-based Shutter Encoder Start function candidate. control_type=%s text=%s center=(%d,%d)",
            self._wrapper_control_type(candidate) or "<unknown>",
            self._wrapper_text(candidate) or "<empty>",
            candidate_center_x,
            candidate_center_y,
        )
        return candidate

    def _type_into_focused_control(
        self,
        value: str,
        *,
        replace_existing: bool = False,
        commit_with_enter: bool = False,
        prefer_clipboard: bool = False,
        dismiss_with_escape: bool = False,
    ) -> None:
        pyautogui = self._import_pyautogui()
        if prefer_clipboard:
            logger.info(
                "Setting the focused Shutter Encoder control through the clipboard to avoid IME interference: %s",
                value,
            )
            paste_text_via_clipboard(value, replace_existing=replace_existing)
        else:
            logger.info("Typing into the focused Shutter Encoder control with pyautogui: %s", value)
            if replace_existing:
                logger.info("Clearing the focused Shutter Encoder control before typing.")
                pyautogui.hotkey("ctrl", "a")
                time.sleep(0.05)
                pyautogui.press("backspace")
                time.sleep(0.05)
            if len(value) == 1 and value.isalpha() and value.upper() == value:
                pyautogui.keyDown("shift")
                try:
                    pyautogui.press(value.lower())
                finally:
                    pyautogui.keyUp("shift")
            else:
                pyautogui.write(value, interval=0.03)
        if commit_with_enter:
            logger.info("Committing the focused Shutter Encoder control value with Enter.")
            send_hotkey("{ENTER}")
        if dismiss_with_escape:
            logger.info("Dismissing any lingering IME or combo popup with Escape.")
            send_hotkey("{ESC}")
        time.sleep(SHUTTER_ENCODER_SETTLE_SECONDS)

    def _read_function_picker_value(self, window) -> str:
        picker = self._find_function_picker(window)
        return (self._read_control_value(picker) or self._wrapper_text(picker) or "").strip()

    def _read_function_picker_value_via_clipboard(self, window) -> str:
        window = self._connect_main_window(timeout=SHUTTER_ENCODER_CONTROL_TIMEOUT_SECONDS)
        bring_window_to_front(window, keep_topmost=False)
        try:
            picker = self._find_function_picker(window)
            logger.info("Refocusing the Shutter Encoder function picker for clipboard readback.")
            self._click_wrapper_center(picker)
        except Exception as exc:
            logger.info(
                "Could not refocus the Shutter Encoder function picker through UIA for clipboard readback. "
                "Falling back to the fixed design point: %s",
                exc,
            )
            self._click_window_design_point_with_pyautogui(
                window,
                SHUTTER_ENCODER_FUNCTION_EDITOR_DESIGN_X,
                SHUTTER_ENCODER_FUNCTION_EDITOR_DESIGN_Y,
                "function picker design point for clipboard readback",
            )
        time.sleep(0.1)
        send_hotkey("^a")
        time.sleep(0.05)
        send_hotkey("^c")
        time.sleep(0.1)
        clipboard_value = get_clipboard_text()
        logger.info("Shutter Encoder function picker clipboard readback. value=%s", clipboard_value or "<empty>")
        return clipboard_value

    def _wait_until_closed(
        self,
        timeout: float = SHUTTER_ENCODER_CLOSE_TIMEOUT_SECONDS,
        *,
        process_id: int | None = None,
    ) -> bool:
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            if process_id is not None:
                remaining_windows = self._iter_process_top_level_windows(process_id)
                if not remaining_windows:
                    return True
                for window in remaining_windows[:3]:
                    try:
                        dismiss_close_prompts(timeout=0.5, owner_window=window)
                    except Exception:
                        pass
            else:
                try:
                    window = self._connect_main_window(timeout=0.5)
                except UiAutomationError:
                    return True
                dismiss_close_prompts(timeout=0.5, owner_window=window)
            time.sleep(0.2)
        if process_id is not None:
            return not self._iter_process_top_level_windows(process_id)
        try:
            self._connect_main_window(timeout=0.5)
        except UiAutomationError:
            return True
        return False

    def _delete_generated_output_artifact(self, generated_output: Path, input_video_path: Path, output_video_path: Path) -> None:
        generated_resolved = generated_output.resolve(strict=False)
        input_resolved = input_video_path.resolve(strict=False)
        output_resolved = output_video_path.resolve(strict=False)
        if generated_resolved == input_resolved:
            logger.info("Skipping Shutter Encoder generated-file cleanup because the candidate matches the input video.")
            return
        if generated_resolved == output_resolved:
            logger.info("Skipping Shutter Encoder generated-file cleanup because the candidate matches the validation output path.")
            return
        if generated_output.parent.resolve(strict=False) != input_video_path.parent.resolve(strict=False):
            logger.info(
                "Skipping Shutter Encoder generated-file cleanup because the candidate is not in the source folder. candidate=%s",
                generated_output,
            )
            return
        if not generated_output.exists():
            logger.info("Skipping Shutter Encoder generated-file cleanup because the candidate no longer exists: %s", generated_output)
            return
        logger.info("Deleting the intermediate Shutter Encoder generated file from the source folder: %s", generated_output)
        generated_output.unlink()

    def _function_picker_is_target(self, window, target_text: str = SHUTTER_ENCODER_TARGET_FUNCTION_TEXT) -> bool:
        try:
            current_value = self._read_function_picker_value(window)
        except Exception:
            return False
        normalized_current = re.sub(r"\s+", " ", current_value).strip().casefold()
        normalized_target = re.sub(r"\s+", " ", target_text).strip().casefold()
        return bool(normalized_current and normalized_current == normalized_target)

    def _open_input_dialog(self, window) -> None:
        self._dismiss_update_dialog_if_present(timeout=1.0)
        dialog = None
        try:
            browse_button = self._find_browse_control(window)
            logger.info("Clicking the Shutter Encoder Browse control: %s", self._wrapper_text(browse_button) or "<untitled>")
            self._click_wrapper_center(browse_button)
            dialog = wait_for_file_dialog(
                dialog_patterns=OPEN_DIALOG_PATTERNS,
                timeout=2.0,
            )
        except Exception as exc:
            logger.info("The primary Shutter Encoder Browse click path did not open the file dialog: %s", exc)

        if dialog is None:
            logger.info("Retrying the Shutter Encoder Browse click with the fixed fallback area.")
            try:
                refreshed = connect_window(self.profile.main_window_title_re, timeout=2.0)
            except UiAutomationError:
                refreshed = window
            self._click_browse_fallback_area(refreshed)
            try:
                dialog = wait_for_file_dialog(
                    dialog_patterns=OPEN_DIALOG_PATTERNS,
                    timeout=2.0,
                )
            except UiAutomationError:
                logger.info("The first Shutter Encoder Browse fallback click did not open the file dialog. Retrying the same fallback click once more.")
                self._click_browse_fallback_area(refreshed)
                dialog = wait_for_file_dialog(
                    dialog_patterns=OPEN_DIALOG_PATTERNS,
                    timeout=SHUTTER_ENCODER_DIALOG_TIMEOUT_SECONDS,
                )
        assert dialog is not None, "Shutter Encoder did not open the file picker after Browse was clicked."

    def _find_function_picker(self, window):
        candidate = self._first_matching_control(
            window,
            SHUTTER_ENCODER_FUNCTION_PICKER_PATTERNS,
            control_types=("ComboBox", "Edit"),
        )
        if candidate is not None:
            return candidate

        window_rect = self._wrapper_rect(window)
        fallback_candidates = []
        for wrapper in self._iter_wrapper_tree(window):
            control_type = self._wrapper_control_type(wrapper).lower()
            if control_type not in {"combobox", "edit"}:
                continue
            rect = self._wrapper_rect(wrapper)
            width = rect.right - rect.left
            height = rect.bottom - rect.top
            if width < 140 or height < 18:
                continue
            if rect.top > window_rect.top + ((window_rect.bottom - window_rect.top) * 3 // 4):
                continue
            current_value = self._read_control_value(wrapper) or self._wrapper_text(wrapper)
            score = (
                0 if current_value and self._matches_patterns(current_value, SHUTTER_ENCODER_FUNCTION_PICKER_PATTERNS) else 1,
                rect.top,
                rect.left,
                -width,
            )
            fallback_candidates.append((score, wrapper, current_value, control_type))
        if not fallback_candidates:
            raise UiAutomationError("Could not locate the Shutter Encoder function picker.")
        fallback_candidates.sort(key=lambda item: item[0])
        _, picker, current_value, control_type = fallback_candidates[0]
        logger.info(
            "Resolved a fallback Shutter Encoder function picker. control_type=%s current_value=%s",
            control_type,
            current_value or "<empty>",
        )
        return picker

    def _select_function(self, window, function_name: str) -> None:
        window = self._connect_main_window(timeout=SHUTTER_ENCODER_CONTROL_TIMEOUT_SECONDS)
        bring_window_to_front(window, keep_topmost=False)
        normalized_target = re.sub(r"\s+", " ", function_name).strip().casefold()
        try:
            picker = self._find_function_picker(window)
        except Exception as exc:
            picker = None
            logger.info("The Shutter Encoder function picker was not exposed through UIA. Falling back to a coordinate click: %s", exc)

        if picker is not None:
            current_value = self._read_control_value(picker) or self._wrapper_text(picker)
            normalized_current = re.sub(r"\s+", " ", current_value).strip().casefold()
            if normalized_current and normalized_current == normalized_target:
                logger.info("Shutter Encoder already shows %s in the function picker. Skipping extra typing.", function_name)
                return
            logger.info("Focusing the Shutter Encoder function picker.")
            self._click_wrapper_center(picker)
        else:
            logger.info("Focusing the Shutter Encoder function picker through the fixed design point.")
            self._click_window_design_point_with_pyautogui(
                window,
                SHUTTER_ENCODER_FUNCTION_EDITOR_DESIGN_X,
                SHUTTER_ENCODER_FUNCTION_EDITOR_DESIGN_Y,
                "function picker design point",
            )

        self._type_into_focused_control(
            function_name,
            replace_existing=True,
            commit_with_enter=True,
            prefer_clipboard=True,
            dismiss_with_escape=True,
        )
        deadline = time.monotonic() + SHUTTER_ENCODER_FUNCTION_SELECTION_TIMEOUT_SECONDS
        last_value = None
        last_clipboard_value = None
        while time.monotonic() < deadline:
            window = self._connect_main_window(timeout=2.0)
            try:
                current_value = self._read_function_picker_value(window)
            except Exception:
                current_value = ""
            if current_value != last_value:
                logger.info("Shutter Encoder function picker state. current_value=%s", current_value or "<empty>")
                last_value = current_value
            normalized_current = re.sub(r"\s+", " ", current_value).strip().casefold()
            if normalized_current and normalized_current == normalized_target:
                logger.info("Verified that Shutter Encoder shows %s as the selected function.", function_name)
                return
            if not current_value:
                try:
                    clipboard_value = self._read_function_picker_value_via_clipboard(window)
                except Exception as exc:
                    clipboard_value = ""
                    logger.info("Shutter Encoder function picker clipboard readback failed: %s", exc)
                if clipboard_value != last_clipboard_value:
                    last_clipboard_value = clipboard_value
                normalized_clipboard = re.sub(r"\s+", " ", clipboard_value).strip().casefold()
                if normalized_clipboard and normalized_clipboard == normalized_target:
                    logger.info(
                        "Verified that Shutter Encoder shows %s as the selected function through clipboard readback.",
                        function_name,
                    )
                    return
            time.sleep(0.2)

        raise AssertionError(f"Shutter Encoder did not settle to {function_name} before Start function was allowed.")

    def _start_function(self, window) -> None:
        window = self._connect_main_window(timeout=SHUTTER_ENCODER_CONTROL_TIMEOUT_SECONDS)
        bring_window_to_front(window, keep_topmost=False)
        try:
            start_button = wait_for_text_control(
                window,
                SHUTTER_ENCODER_START_BUTTON_PATTERNS,
                control_types=("Button", "Text", "Pane", "Group", "Custom"),
                timeout=SHUTTER_ENCODER_CONTROL_TIMEOUT_SECONDS,
            )
            logger.info("Clicking the Shutter Encoder Start function button: %s", self._wrapper_text(start_button) or "<untitled>")
            enabled = self._is_enabled(start_button)
            assert enabled is not False, "The Shutter Encoder Start function button is disabled."
            self._click_wrapper_center(start_button)
        except Exception as exc:
            logger.info("The Shutter Encoder Start function button was not exposed through UIA. Falling back to a coordinate click: %s", exc)
            try:
                candidate = self._find_start_button_candidate_near_picker(window)
                bring_window_to_front(window, keep_topmost=False)
                click_x, click_y = self._wrapper_center_screen_point(candidate)
                self._click_screen_point_with_pyautogui(click_x, click_y, "geometry-based Start function button candidate")
            except Exception as candidate_exc:
                logger.info(
                    "The geometry-based Shutter Encoder Start function fallback failed. "
                    "Falling back to the picker-anchored click point: %s",
                    candidate_exc,
                )
                try:
                    click_x, click_y = self._predict_start_button_screen_point_from_picker(window)
                    self._click_screen_point_with_pyautogui(click_x, click_y, "picker-anchored Start function button point")
                except Exception as picker_exc:
                    logger.info(
                        "The picker-anchored Shutter Encoder Start function fallback failed. "
                        "Falling back to the fixed design point: %s",
                        picker_exc,
                    )
                    try:
                        self._click_window_design_point_with_pyautogui(
                            window,
                            SHUTTER_ENCODER_START_BUTTON_DESIGN_X,
                            SHUTTER_ENCODER_START_BUTTON_DESIGN_Y,
                            "Start function button design point",
                        )
                    except Exception as design_exc:
                        logger.info(
                            "The Shutter Encoder Start function design-point fallback failed. "
                            "Falling back to the old window-ratio click: %s",
                            design_exc,
                        )
                        try:
                            self._click_window_ratio_with_pyautogui(
                                window,
                                SHUTTER_ENCODER_START_BUTTON_CLICK_X_RATIO,
                                SHUTTER_ENCODER_START_BUTTON_CLICK_Y_RATIO,
                                "Start function button",
                            )
                        except Exception as click_exc:
                            logger.info("The Shutter Encoder Start function coordinate fallback failed. Falling back to Enter: %s", click_exc)
                            send_hotkey("{ENTER}")
                            time.sleep(SHUTTER_ENCODER_SETTLE_SECONDS)

    def _wait_for_job_to_start(
        self,
        input_video_path: Path,
        known_outputs: dict[Path, tuple[float, int]],
        *,
        timeout: float = 15.0,
    ) -> None:
        logger.info("Waiting for the Shutter Encoder job to actually start. timeout=%.1fs", timeout)
        deadline = time.monotonic() + timeout
        progress_logger = _WaitProgressLogger("Shutter Encoder start")
        while time.monotonic() < deadline:
            window = self._connect_main_window(timeout=2.0)
            progress_ratio = self._read_progress_ratio(window)
            output_candidate = self._resolve_output_candidate(input_video_path, known_outputs)
            started = (progress_ratio is not None and progress_ratio > 0.0) or output_candidate is not None
            progress_logger.maybe_log(
                detail=(
                    f"progress_ratio={'n/a' if progress_ratio is None else f'{progress_ratio:.3f}'} "
                    f"output_candidate={'<none>' if output_candidate is None else output_candidate.name}"
                )
            )
            if started:
                logger.info("Shutter Encoder start conditions satisfied. elapsed=%s", progress_logger.elapsed_text())
                return
            time.sleep(0.5)
        raise AssertionError("Shutter Encoder did not start the job after the start trigger.")

    def _wait_for_completion(self, input_video_path: Path, known_outputs: dict[Path, tuple[float, int]]) -> Path:
        completion_rounds = 0
        output_stable_rounds = 0
        last_output_signature = None
        deadline = time.monotonic() + SHUTTER_ENCODER_PROGRESS_TIMEOUT_SECONDS
        progress_logger = _WaitProgressLogger("Shutter Encoder transcode")
        while time.monotonic() < deadline:
            window = self._connect_main_window(timeout=2.0)
            complete_text = self._control_present(
                window,
                SHUTTER_ENCODER_PROGRESS_COMPLETE_PATTERNS,
                control_types=("Text", "Edit", "Pane", "Group", "Custom"),
            )
            progress_ratio = self._read_progress_ratio(window)
            percent_complete = self._control_present(
                window,
                SHUTTER_ENCODER_PROGRESS_100_PATTERNS,
                control_types=("Text", "Edit", "Pane", "Group"),
            )
            output_candidate = self._resolve_output_candidate(input_video_path, known_outputs)
            if output_candidate is not None:
                stat = output_candidate.stat()
                output_signature = (str(output_candidate), stat.st_size)
                if stat.st_size > 0 and output_signature == last_output_signature:
                    output_stable_rounds += 1
                else:
                    output_stable_rounds = 0
                last_output_signature = output_signature
            else:
                output_stable_rounds = 0
                last_output_signature = None
            progress_logger.maybe_log(
                detail=(
                    f"progress_ratio={'n/a' if progress_ratio is None else f'{progress_ratio:.3f}'} "
                    f"output_candidate={'<none>' if output_candidate is None else output_candidate.name} "
                    f"output_stable_rounds={output_stable_rounds}"
                )
            )

            has_uia_progress_signal = complete_text or percent_complete or progress_ratio is not None
            progress_complete = percent_complete or (progress_ratio is not None and progress_ratio >= 0.999)
            if complete_text and progress_complete:
                completion_rounds += 1
            else:
                completion_rounds = 0

            if completion_rounds >= SHUTTER_ENCODER_PROGRESS_IDLE_ROUNDS and output_candidate is not None and output_stable_rounds >= 1:
                logger.info(
                    "Shutter Encoder completion conditions satisfied. elapsed=%s generated_output=%s",
                    progress_logger.elapsed_text(),
                    output_candidate,
                )
                return output_candidate

            if (
                not has_uia_progress_signal
                and output_candidate is not None
                and output_stable_rounds >= SHUTTER_ENCODER_OUTPUT_STABLE_COMPLETION_ROUNDS
            ):
                logger.info(
                    "Shutter Encoder progress UI is not exposed through UIA. "
                    "Treating a stable generated output as completion. elapsed=%s generated_output=%s",
                    progress_logger.elapsed_text(),
                    output_candidate,
                )
                return output_candidate

            time.sleep(SHUTTER_ENCODER_PROGRESS_POLL_SECONDS)

        raise AssertionError(f"Shutter Encoder did not finish within timeout: {input_video_path}")

    def close(self) -> None:
        logger.info("Closing Shutter Encoder.")
        try:
            window = self._connect_main_window(timeout=2.0)
        except UiAutomationError:
            logger.info("Shutter Encoder window is already closed.")
            return
        process_id = self._process_id(window)
        bring_window_to_front(window, keep_topmost=False)

        try:
            logger.info("Trying pywinauto window.close() for Shutter Encoder.")
            window.close()
        except Exception as exc:
            logger.info("pywinauto window.close() did not complete Shutter Encoder shutdown: %s", exc)
            request_window_close(window)
        if self._wait_until_closed(timeout=1.5, process_id=process_id):
            logger.info("Shutter Encoder closed successfully.")
            return
        self._close_remaining_process_windows(process_id)
        if self._wait_until_closed(timeout=1.0, process_id=process_id):
            logger.info("Shutter Encoder closed successfully.")
            return

        try:
            logger.info("Trying the custom top-right Shutter Encoder close button.")
            self._click_top_right_offset_with_pyautogui(
                window,
                SHUTTER_ENCODER_QUIT_BUTTON_RIGHT_OFFSET,
                SHUTTER_ENCODER_QUIT_BUTTON_TOP_OFFSET,
                "custom close button",
            )
        except Exception as exc:
            logger.info("The custom top-right Shutter Encoder close button click failed: %s", exc)
        if self._wait_until_closed(timeout=1.5, process_id=process_id):
            logger.info("Shutter Encoder closed successfully.")
            return
        self._close_remaining_process_windows(process_id)
        if self._wait_until_closed(timeout=1.0, process_id=process_id):
            logger.info("Shutter Encoder closed successfully.")
            return

        logger.info("Posting WM_CLOSE for Shutter Encoder as a targeted close fallback.")
        bring_window_to_front(window, keep_topmost=False)
        request_window_close(window)
        if self._wait_until_closed(timeout=2.0, process_id=process_id):
            logger.info("Shutter Encoder closed successfully.")
            return
        self._close_remaining_process_windows(process_id)
        if self._wait_until_closed(timeout=1.0, process_id=process_id):
            logger.info("Shutter Encoder closed successfully.")
            return

        try:
            window = self._connect_main_window(timeout=1.0)
        except UiAutomationError:
            if self._wait_until_closed(timeout=0.5, process_id=process_id):
                logger.info("Shutter Encoder closed successfully.")
                return
            window = None
        if window is not None:
            try:
                logger.info("Retrying the custom top-right Shutter Encoder close button once more.")
                self._click_top_right_offset_with_pyautogui(
                    window,
                    SHUTTER_ENCODER_QUIT_BUTTON_RIGHT_OFFSET,
                    SHUTTER_ENCODER_QUIT_BUTTON_TOP_OFFSET,
                    "custom close button retry",
                )
            except Exception as exc:
                logger.info("The retry custom close button click failed: %s", exc)
            if self._wait_until_closed(timeout=2.0, process_id=process_id):
                logger.info("Shutter Encoder closed successfully.")
                return
            dismiss_close_prompts(timeout=1.5, owner_window=window)
        self._close_remaining_process_windows(process_id)
        if self._wait_until_closed(timeout=1.0, process_id=process_id):
            logger.info("Shutter Encoder closed successfully.")
            return

        if process_id is not None:
            logger.warning("Shutter Encoder still has visible UI after close retries. Terminating process pid=%s.", process_id)
            subprocess.run(
                ["taskkill", "/PID", str(process_id), "/T", "/F"],
                capture_output=True,
                text=True,
                check=False,
            )
            if self._wait_until_closed(timeout=2.0, process_id=process_id):
                logger.info("Shutter Encoder closed successfully.")
                return

        raise AssertionError("Shutter Encoder did not close after completion.")

    def perform(self, input_video_path: Path, output_video_path: Path) -> None:
        logger.info("Shutter Encoder flow started. input=%s output=%s", input_video_path, output_video_path)
        logger.info("Validating Shutter Encoder input and output paths.")
        assert input_video_path.exists(), f"Input video does not exist: {input_video_path}"
        assert output_video_path.parent.exists(), f"Output directory does not exist: {output_video_path.parent}"

        known_outputs = self._snapshot_output_candidates(input_video_path)
        try:
            logger.info("Checking whether Shutter Encoder is already running.")
            window = self._connect_main_window(timeout=2.0)
            logger.info("Using the existing Shutter Encoder window.")
        except UiAutomationError:
            logger.info("Shutter Encoder is not running. Launching it now.")
            self._launch_application()
            window = self._connect_main_window(timeout=30.0)
        bring_window_to_front(window, keep_topmost=False)

        self._dismiss_update_dialog_if_present(timeout=2.0)
        window = self._connect_main_window(timeout=SHUTTER_ENCODER_CONTROL_TIMEOUT_SECONDS)
        bring_window_to_front(window, keep_topmost=False)

        logger.info("Opening the Shutter Encoder source picker.")
        self._open_input_dialog(window)
        logger.info("Filling the Shutter Encoder source picker.")
        fill_file_dialog(
            input_video_path,
            dialog_patterns=OPEN_DIALOG_PATTERNS,
            timeout=SHUTTER_ENCODER_DIALOG_TIMEOUT_SECONDS,
            must_exist=True,
        )
        time.sleep(SHUTTER_ENCODER_SETTLE_SECONDS)

        logger.info("Reconnecting to Shutter Encoder after import.")
        window = self._connect_main_window(timeout=SHUTTER_ENCODER_CONTROL_TIMEOUT_SECONDS)
        bring_window_to_front(window, keep_topmost=False)

        logger.info("Validating the imported source is visible in Shutter Encoder.")
        self._wait_for_imported_source(window, input_video_path)

        logger.info(
            "Selecting %s in Shutter Encoder by writing '%s' into the function picker and confirming with Enter.",
            SHUTTER_ENCODER_TARGET_FUNCTION_TEXT,
            SHUTTER_ENCODER_TARGET_FUNCTION_TEXT,
        )
        self._select_function(window, SHUTTER_ENCODER_TARGET_FUNCTION_TEXT)
        logger.info("Starting the Shutter Encoder transcode.")
        self._start_function(window)
        self._begin_background_wait(
            window,
            phase="transcode",
            start_waiter=lambda: self._wait_for_job_to_start(input_video_path, known_outputs),
            start_description="the Shutter Encoder transcode",
        )
        logger.info("Waiting for the Shutter Encoder job to complete.")
        generated_output = self._wait_for_completion(input_video_path, known_outputs)

        logger.info("Closing Shutter Encoder after completion.")
        self.close()

        logger.info("Copying the generated Shutter Encoder output to the requested validation path. source=%s target=%s", generated_output, output_video_path)
        shutil.copy2(generated_output, output_video_path)
        assert output_video_path.exists() and output_video_path.stat().st_size > 0, (
            f"Shutter Encoder validation output was not copied correctly: {output_video_path}"
        )
        self._delete_generated_output_artifact(generated_output, input_video_path, output_video_path)


class KdenliveOperator(SoftwareOperator):
    def __init__(self, profile: OperationProfile) -> None:
        super().__init__(profile)
        self._main_window = None
        self._render_dialog = None

    def perform(self, input_video_path: Path, output_video_path: Path) -> None:
        output_video_path = output_video_path.resolve(strict=False)
        logger.info("Kdenlive flow started. input=%s output=%s", input_video_path, output_video_path)
        logger.info("Validating Kdenlive input and output paths.")
        assert input_video_path.exists(), f"Input video does not exist: {input_video_path}"
        assert output_video_path.parent.exists(), f"Output directory does not exist: {output_video_path.parent}"
        assert output_video_path.suffix.lower() == ".mp4", f"Kdenlive export target must be an mp4 file: {output_video_path}"

        logger.info("Connecting to Kdenlive main window.")
        window = self._connect_main_window(timeout=max(30.0, SOFTWARE_SPECS["kdenlive"].startup_timeout_seconds))
        self._main_window = window
        if self._dismiss_recovery_dialog_if_present(window, timeout=3.0):
            logger.info("Reconnecting to Kdenlive after dismissing the file-recovery dialog.")
            window = self._connect_main_window(timeout=KDENLIVE_CONTROL_TIMEOUT_SECONDS)
            self._main_window = window
        bring_window_to_front(window, keep_topmost=False)
        window = self._open_input_clip(window, input_video_path)
        logger.info("Inserting selected clip into timeline.")
        self._insert_clip_to_timeline(window)
        self._render_current_timeline(window, output_video_path)
        self.close()

    def close(self) -> None:
        logger.info("Closing Kdenlive render dialog and main window.")
        if self._render_dialog is not None:
            self._close_render_dialog_only()
        if self._main_window is not None:
            self._dismiss_profile_switch_prompt_if_present(self._main_window)
            try:
                bring_window_to_front(self._main_window, keep_topmost=False)
                request_window_close(self._main_window)
            except Exception as exc:
                logger.info("Posting WM_CLOSE to the Kdenlive main window raised %s.", exc)
                try:
                    request_window_close(self._main_window)
                except Exception:
                    pass
            logger.info("Handling any Kdenlive save-confirmation dialog triggered by closing the main window.")
            self._dismiss_kdenlive_save_dialog_if_present(self._main_window, timeout=3.0)
            dismiss_close_prompts(timeout=1.5, owner_window=self._main_window)
            if not self._wait_for_main_window_to_close(timeout=3.0):
                process_id = self._process_id(self._main_window)
                if process_id is not None:
                    logger.warning(
                        "Kdenlive still has visible UI after save-dismissal retries. Terminating process pid=%s.",
                        process_id,
                    )
                    subprocess.run(
                        ["taskkill", "/PID", str(process_id), "/T", "/F"],
                        capture_output=True,
                        text=True,
                        check=False,
                    )
            self._main_window = None
            return
        super().close()

    def _close_render_dialog_only(self) -> None:
        if self._render_dialog is None:
            return
        logger.info("Closing the Kdenlive render dialog while keeping the main window open.")
        render_dialog = self._render_dialog
        try:
            bring_window_to_front(render_dialog, keep_topmost=False)
            close_button = find_text_control(render_dialog, KDENLIVE_CLOSE_BUTTON_PATTERNS, control_types=("Button",))
            click_control(close_button, post_click_sleep=0.5)
        except Exception:
            try:
                render_dialog.close()
            except Exception:
                pass
        try:
            logger.info("Handling any Kdenlive save-confirmation dialog triggered by closing the render window.")
            self._dismiss_kdenlive_save_dialog_if_present(self._main_window or render_dialog, timeout=2.0)
            dismiss_close_prompts(timeout=1.5, owner_window=self._main_window or render_dialog)
            self._dismiss_kdenlive_save_dialog_if_present(self._main_window or render_dialog, timeout=1.0)
        except Exception:
            logger.exception("Could not dismiss Kdenlive save-confirmation dialog after closing the render window.")
        self._render_dialog = None
        if self._main_window is not None:
            try:
                self._main_window = self._connect_main_window(timeout=KDENLIVE_CONTROL_TIMEOUT_SECONDS)
            except Exception:
                pass

    def _open_input_clip(self, window, input_video_path: Path):
        if self._dismiss_kdenlive_save_dialog_if_present(window, timeout=1.0):
            logger.info("Kdenlive save-confirmation dialog was blocking the next import. Reconnecting to the main window.")
            try:
                window = self._connect_main_window(timeout=KDENLIVE_CONTROL_TIMEOUT_SECONDS)
                self._main_window = window
                bring_window_to_front(window, keep_topmost=False)
            except Exception:
                pass
        logger.info("Opening Kdenlive import command.")
        self._open_import_dialog(window)
        logger.info("Filling Kdenlive file picker.")
        fill_file_dialog(
            input_video_path,
            dialog_patterns=KDENLIVE_OPEN_DIALOG_PATTERNS,
            confirm_patterns=KDENLIVE_OPEN_CONFIRM_PATTERNS,
            timeout=KDENLIVE_DIALOG_TIMEOUT_SECONDS,
            must_exist=True,
        )
        time.sleep(KDENLIVE_IMPORT_SETTLE_SECONDS)

        logger.info("Reconnecting to Kdenlive after import.")
        window = self._connect_main_window(timeout=KDENLIVE_CONTROL_TIMEOUT_SECONDS)
        self._main_window = window
        bring_window_to_front(window, keep_topmost=False)

        logger.info("Waiting for imported clip in Sequences.")
        clip_item = self._wait_for_imported_clip(window, input_video_path)
        logger.info("Selecting imported clip.")
        click_control(clip_item, post_click_sleep=KDENLIVE_POST_ACTION_SLEEP_SECONDS)
        return window

    def _render_current_timeline(self, window, output_video_path: Path):
        logger.info("Opening Kdenlive rendering dialog.")
        render_dialog = self._open_render_dialog(window)
        logger.info("Setting render output path.")
        self._set_render_output_path(render_dialog, output_video_path)
        self._show_job_queue_tab(render_dialog)
        logger.info("Starting Kdenlive render.")
        self._start_render(render_dialog)
        self._begin_background_wait(
            render_dialog,
            phase="render",
            start_waiter=lambda: self._confirm_render_transition_after_minimizing(render_dialog),
            start_description="the Kdenlive render transition",
        )
        logger.info("Waiting for Kdenlive render completion.")
        self._wait_for_render_completion(render_dialog, output_video_path)
        if self._main_window is None:
            self._main_window = window
        return self._main_window

    def _process_id(self, window) -> int | None:
        try:
            process_id = getattr(getattr(window, "element_info", None), "process_id", None)
        except Exception:
            return None
        return int(process_id) if process_id else None

    def _iter_process_top_level_windows(self, process_id: int | None):
        if process_id is None:
            return []
        from pywinauto import Desktop

        matches = []
        for window in Desktop(backend="uia").windows():
            try:
                pid = int(getattr(getattr(window, "element_info", None), "process_id", 0) or 0)
            except Exception:
                continue
            if pid != process_id:
                continue
            try:
                if not window.is_visible():
                    continue
            except Exception:
                pass
            try:
                rect = self._wrapper_rect(window)
                area = max(0, rect.right - rect.left) * max(0, rect.bottom - rect.top)
            except Exception:
                area = 0
            matches.append((-area, self._wrapper_text(window).casefold(), window))
        matches.sort(key=lambda item: (item[0], item[1]))
        return [item[2] for item in matches]

    def _wait_for_main_window_to_close(self, timeout: float = 3.0) -> bool:
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            try:
                self._connect_main_window(timeout=0.5)
            except UiAutomationError:
                return True
            time.sleep(0.2)
        return False

    def _dialog_matches(self, dialog, *, title_patterns: tuple[str, ...], text_patterns: tuple[str, ...]) -> bool:
        title = self._wrapper_text(dialog)
        if title and self._matches_patterns(title, title_patterns):
            return True
        return self._control_present(
            dialog,
            text_patterns,
            control_types=("Text", "Pane", "Group", "Document", "Custom"),
        )

    def _dismiss_kdenlive_save_dialog_if_present(self, owner_window, *, timeout: float = 3.0) -> bool:
        process_id = self._process_id(owner_window)
        deadline = time.monotonic() + timeout
        search_roots = [owner_window]
        try:
            top_level_owner = owner_window.top_level_parent()
        except Exception:
            top_level_owner = owner_window
        if getattr(top_level_owner, "handle", None) != getattr(owner_window, "handle", None):
            search_roots.append(top_level_owner)
        for candidate in self._iter_process_top_level_windows(process_id):
            if getattr(candidate, "handle", None) in {
                getattr(root, "handle", None)
                for root in search_roots
            }:
                continue
            search_roots.append(candidate)
        dismissed = False
        while time.monotonic() < deadline:
            dialog = None
            for root in search_roots:
                for candidate in self._iter_wrapper_tree(root):
                    if getattr(candidate, "handle", None) == getattr(owner_window, "handle", None):
                        continue
                    if not self._dialog_matches(
                        candidate,
                        title_patterns=KDENLIVE_WARNING_DIALOG_PATTERNS,
                        text_patterns=KDENLIVE_SAVE_CHANGES_TEXT_PATTERNS,
                    ):
                        continue
                    if self._first_matching_control(candidate, KDENLIVE_DONT_SAVE_PATTERNS, control_types=("Button",)) is None:
                        if not self._control_present(
                            candidate,
                            KDENLIVE_SAVE_CHANGES_TEXT_PATTERNS,
                            control_types=("Text", "Pane", "Group", "Document", "Custom"),
                        ):
                            continue
                    dialog = candidate
                    break
                if dialog is not None:
                    break
            if dialog is None:
                return dismissed
            try:
                bring_window_to_front(dialog, keep_topmost=False)
            except Exception:
                pass
            button = self._first_matching_control(dialog, KDENLIVE_DONT_SAVE_PATTERNS, control_types=("Button",))
            if button is None:
                logger.info(
                    "Kdenlive save dialog was found, but a discard-style button was not directly exposed. "
                    "Trying keyboard dismiss shortcuts before generic close-prompt dismissal."
                )
                for key in KDENLIVE_DONT_SAVE_DIRECT_KEYS:
                    try:
                        logger.info("Sending Kdenlive discard shortcut: %s", key)
                        send_hotkey(key)
                    except Exception:
                        pass
                try:
                    dismiss_close_prompts(timeout=0.8, owner_window=dialog)
                except Exception:
                    logger.exception("Generic prompt dismissal failed while handling the Kdenlive save dialog.")
                time.sleep(0.2)
                dismissed = True
                continue
            logger.info("Clicking the Kdenlive discard-style button: %s", self._wrapper_text(button) or "<untitled>")
            try:
                click_control(button, post_click_sleep=0.5)
            except Exception as exc:
                logger.info(
                    "Clicking the Kdenlive discard-style button failed. Falling back to keyboard shortcuts. details=%s",
                    exc,
                )
                for key in KDENLIVE_DONT_SAVE_DIRECT_KEYS:
                    try:
                        logger.info("Sending Kdenlive discard shortcut: %s", key)
                        send_hotkey(key)
                    except Exception:
                        pass
                try:
                    dismiss_close_prompts(timeout=0.8, owner_window=dialog)
                except Exception:
                    logger.exception("Generic prompt dismissal failed while handling the Kdenlive save dialog.")
            dismissed = True
        return dismissed

    def _accept_kdenlive_overwrite_dialog_if_present(self, owner_window, *, timeout: float = 2.5) -> bool:
        process_id = self._process_id(owner_window)
        deadline = time.monotonic() + timeout
        search_roots = [owner_window]
        try:
            top_level_owner = owner_window.top_level_parent()
        except Exception:
            top_level_owner = owner_window
        if getattr(top_level_owner, "handle", None) != getattr(owner_window, "handle", None):
            search_roots.append(top_level_owner)
        for candidate in self._iter_process_top_level_windows(process_id):
            if getattr(candidate, "handle", None) in {
                getattr(root, "handle", None)
                for root in search_roots
            }:
                continue
            search_roots.append(candidate)
        while time.monotonic() < deadline:
            dialog = None
            for root in search_roots:
                for candidate in self._iter_wrapper_tree(root):
                    title = self._wrapper_text(candidate)
                    if not title or not self._matches_patterns(title, KDENLIVE_WARNING_DIALOG_PATTERNS):
                        continue
                    if not self._control_present(
                        candidate,
                        KDENLIVE_OVERWRITE_TEXT_PATTERNS,
                        control_types=("Text", "Pane", "Group", "Document", "Custom"),
                    ):
                        continue
                    if self._first_matching_control(
                        candidate,
                        KDENLIVE_OVERWRITE_CONFIRM_PATTERNS,
                        control_types=("Button",),
                    ) is None:
                        continue
                    dialog = candidate
                    break
                if dialog is not None:
                    break
            if dialog is None:
                time.sleep(0.1)
                continue
            try:
                bring_window_to_front(dialog.top_level_parent(), keep_topmost=False)
            except Exception:
                try:
                    bring_window_to_front(dialog, keep_topmost=False)
                except Exception:
                    pass
            button = self._first_matching_control(dialog, KDENLIVE_OVERWRITE_CONFIRM_PATTERNS, control_types=("Button",))
            assert button is not None, "Kdenlive overwrite dialog was found, but the overwrite button was not exposed."
            logger.info("Clicking the Kdenlive overwrite confirmation button: %s", self._wrapper_text(button) or "<untitled>")
            click_control(button, post_click_sleep=0.5)
            return True
        logger.info("No Kdenlive overwrite confirmation dialog appeared.")
        return False

    def _resolve_render_dialog_top_level(self, render_container) -> None:
        try:
            top_level = render_container.top_level_parent()
        except Exception:
            top_level = render_container
        try:
            bring_window_to_front(top_level, keep_topmost=False)
        except Exception:
            pass
        if self._main_window is not None and getattr(top_level, "handle", None) == getattr(self._main_window, "handle", None):
            self._render_dialog = None
        else:
            self._render_dialog = top_level
        logger.info("Resolved Kdenlive rendering container. top_level=%s", self._wrapper_text(top_level) or "<untitled>")

    def _iter_render_search_roots(self, window):
        yielded_handles: set[int | None] = set()
        for candidate in (window, *self._iter_process_top_level_windows(self._process_id(window))):
            handle = getattr(candidate, "handle", None)
            if handle in yielded_handles:
                continue
            yielded_handles.add(handle)
            yield candidate

    def _dismiss_profile_switch_prompt_if_present(self, window) -> bool:
        if not self._control_present(
            window,
            KDENLIVE_PROFILE_SWITCH_TEXT_PATTERNS,
            control_types=("Text", "Pane", "Group", "Document", "Custom"),
        ):
            return False
        cancel_button = self._first_matching_control(window, KDENLIVE_DIALOG_CANCEL_PATTERNS, control_types=("Button",))
        if cancel_button is None:
            logger.info("Kdenlive profile-switch prompt was detected, but the Cancel button was not exposed.")
            return False
        logger.info("Cancelling the Kdenlive profile-switch prompt before shutdown.")
        click_control(cancel_button, post_click_sleep=0.5)
        return True

    def _dismiss_recovery_dialog_if_present(self, owner_window, *, timeout: float = 3.0) -> bool:
        process_id = self._process_id(owner_window)
        owner_handle = getattr(owner_window, "handle", None)
        deadline = time.monotonic() + timeout
        dismissed = False
        while time.monotonic() < deadline:
            dialog = None
            for candidate in self._iter_process_top_level_windows(process_id):
                if owner_handle is not None and getattr(candidate, "handle", None) == owner_handle:
                    continue
                if not self._dialog_matches(
                    candidate,
                    title_patterns=KDENLIVE_RECOVERY_DIALOG_PATTERNS,
                    text_patterns=KDENLIVE_RECOVERY_TEXT_PATTERNS,
                ):
                    continue
                dialog = candidate
                break
            if dialog is None:
                return dismissed
            try:
                bring_window_to_front(dialog, keep_topmost=False)
            except Exception:
                pass
            button = self._first_matching_control(dialog, KDENLIVE_DO_NOT_RECOVER_PATTERNS, control_types=("Button",))
            assert button is not None, "Kdenlive file-recovery dialog was found, but the 'Do not recover' button was not exposed."
            logger.info("Clicking the Kdenlive recovery-dismiss button: %s", self._wrapper_text(button) or "<untitled>")
            click_control(button, post_click_sleep=0.5)
            dismissed = True
        return dismissed

    def _open_import_dialog(self, window) -> None:
        if self._dismiss_recovery_dialog_if_present(window, timeout=0.8):
            logger.info("Recovered Kdenlive main window focus after dismissing the recovery dialog during import setup.")
            window = self._connect_main_window(timeout=KDENLIVE_CONTROL_TIMEOUT_SECONDS)
            self._main_window = window
            bring_window_to_front(window, keep_topmost=False)
        try:
            logger.info("Trying Kdenlive toolbar add-clip button.")
            add_clip_button = wait_for_text_control(window, KDENLIVE_ADD_CLIP_PATTERNS, control_types=("Button",), timeout=2.0)
            click_control(add_clip_button, post_click_sleep=KDENLIVE_POST_ACTION_SLEEP_SECONDS)
            dialog = wait_for_file_dialog(
                dialog_patterns=KDENLIVE_OPEN_DIALOG_PATTERNS,
                confirm_patterns=KDENLIVE_OPEN_CONFIRM_PATTERNS,
                timeout=KDENLIVE_DIALOG_TIMEOUT_SECONDS,
            )
            assert dialog is not None, "Kdenlive did not open the import dialog from the toolbar button."
            return
        except Exception as exc:
            logger.info("Kdenlive toolbar import path unavailable: %s", exc)

        logger.info("Falling back to Project -> Add Clip or Folder.")
        click_text_control(window, KDENLIVE_PROJECT_MENU_PATTERNS, control_types=("MenuItem",))
        time.sleep(0.3)
        click_text_control(window, KDENLIVE_ADD_CLIP_PATTERNS, control_types=("MenuItem",))
        dialog = wait_for_file_dialog(
            dialog_patterns=KDENLIVE_OPEN_DIALOG_PATTERNS,
            confirm_patterns=KDENLIVE_OPEN_CONFIRM_PATTERNS,
            timeout=KDENLIVE_DIALOG_TIMEOUT_SECONDS,
        )
        assert dialog is not None, "Kdenlive did not open the import dialog."

    def _wait_for_imported_clip(self, window, input_video_path: Path):
        clip_item = wait_for_text_control(
            window,
            re.escape(input_video_path.name),
            control_types=KDENLIVE_CLIP_CONTROL_TYPES,
            timeout=KDENLIVE_CONTROL_TIMEOUT_SECONDS,
        )
        assert clip_item.is_visible(), f"Kdenlive did not show imported clip '{input_video_path.name}' in Project Bin."
        return clip_item

    def _insert_clip_to_timeline(self, window) -> None:
        logger.info("Sending Kdenlive shortcut 'v' to insert the selected clip into the timeline.")
        bring_window_to_front(window, keep_topmost=False)
        send_hotkey("v")
        time.sleep(1.0)

    def _wrapper_text(self, wrapper) -> str:
        try:
            return (wrapper.window_text() or "").strip()
        except Exception:
            return ""

    def _wrapper_control_type(self, wrapper) -> str:
        try:
            return (getattr(wrapper.element_info, "control_type", "") or "").strip()
        except Exception:
            return ""

    def _wrapper_rect(self, wrapper):
        return wrapper.rectangle()

    def _iter_wrapper_tree(self, root):
        yield root
        for child in root.descendants():
            yield child

    def _matches_patterns(self, text: str, patterns: tuple[str, ...]) -> bool:
        return any(re.search(pattern, text, re.IGNORECASE) for pattern in patterns)

    def _vertical_overlap(self, a, b) -> int:
        return max(0, min(a.bottom, b.bottom) - max(a.top, b.top))

    def _read_control_value(self, wrapper) -> str:
        readers = (
            lambda: wrapper.get_value(),
            lambda: wrapper.iface_value.CurrentValue,
            lambda: wrapper.window_text(),
            lambda: wrapper.texts()[0] if wrapper.texts() else "",
        )
        for reader in readers:
            try:
                value = reader()
            except Exception:
                continue
            if isinstance(value, str):
                return value.strip()
        return ""

    def _find_output_file_edit(self, render_dialog, *, timeout: float = KDENLIVE_CONTROL_TIMEOUT_SECONDS, log_result: bool = True):
        label = wait_for_text_control(
            render_dialog,
            KDENLIVE_OUTPUT_FILE_LABEL_PATTERNS,
            control_types=("Text", "Pane", "Group"),
            timeout=timeout,
            poll_interval=0.1,
        )
        label_rect = self._wrapper_rect(label)

        def collect_candidates(min_width: int):
            candidates = []
            for wrapper in self._iter_wrapper_tree(render_dialog):
                control_type = self._wrapper_control_type(wrapper).lower()
                if control_type not in {"edit", "combobox"}:
                    continue
                rect = self._wrapper_rect(wrapper)
                overlap = self._vertical_overlap(label_rect, rect)
                width = rect.right - rect.left
                if rect.left <= label_rect.right:
                    continue
                if overlap <= 0:
                    continue
                if width < min_width:
                    continue
                current_value = self._read_control_value(wrapper)
                looks_like_path = any(token in current_value for token in (":\\", "\\", "/", ".mp4", ".mov", ".mkv"))
                score = (
                    0 if looks_like_path else 1,
                    abs(rect.top - label_rect.top),
                    rect.left - label_rect.right,
                    -width,
                )
                candidates.append((score, wrapper, current_value, control_type, width))
            return candidates

        candidates = collect_candidates(140)
        if not candidates:
            candidates = collect_candidates(80)
        assert candidates, "Kdenlive output-path control was not found in the rendering container."
        candidates.sort(key=lambda item: item[0])
        _, wrapper, current_value, control_type, width = candidates[0]
        if log_result:
            logger.info(
                "Resolved Kdenlive output-path control. control_type=%s width=%s current_value=%s",
                control_type,
                width,
                current_value or "<empty>",
            )
        return wrapper

    def _container_has_render_controls(self, container) -> bool:
        try:
            self._find_render_to_file_button(container)
            wait_for_text_control(
                container,
                KDENLIVE_OUTPUT_FILE_LABEL_PATTERNS,
                control_types=("Text", "Pane", "Group"),
                timeout=0.2,
                poll_interval=0.05,
            )
            return True
        except Exception:
            return False

    def _find_render_container_from_anchor(self, anchor):
        current = anchor
        for _ in range(10):
            if current is None:
                break
            if self._container_has_render_controls(current):
                return current
            try:
                current = current.parent()
            except Exception:
                break
        return None

    def _locate_render_container(self, root, timeout: float):
        deadline = time.monotonic() + timeout
        last_error: Exception | None = None
        while time.monotonic() < deadline:
            try:
                anchor = wait_for_text_control(
                    root,
                    KDENLIVE_RENDER_WINDOW_PATTERNS,
                    control_types=("Text", "Pane", "Window"),
                    timeout=0.5,
                )
                container = self._find_render_container_from_anchor(anchor)
                if container is not None:
                    return container
            except Exception as exc:
                last_error = exc
            try:
                render_button = wait_for_text_control(
                    root,
                    KDENLIVE_RENDER_TO_FILE_PATTERNS,
                    control_types=("Button",),
                    timeout=0.5,
                )
                container = self._find_render_container_from_anchor(render_button)
                if container is not None:
                    return container
            except Exception as exc:
                last_error = exc
            time.sleep(0.1)
        raise AssertionError("Kdenlive rendering container was not found.") from last_error

    def _find_render_to_file_button(self, render_dialog):
        candidates = []
        for wrapper in self._iter_wrapper_tree(render_dialog):
            if self._wrapper_control_type(wrapper).lower() != "button":
                continue
            title = self._wrapper_text(wrapper)
            if not title or not self._matches_patterns(title, KDENLIVE_RENDER_TO_FILE_PATTERNS):
                continue
            try:
                rect = wrapper.rectangle()
            except Exception:
                continue
            candidates.append((rect.bottom, rect.left, wrapper))
        assert candidates, "Kdenlive 'Render to File' button was not found in the rendering container."
        candidates.sort(key=lambda item: (-item[0], item[1]))
        return candidates[0][2]

    def _first_matching_control(self, root, patterns: tuple[str, ...], *, control_types: tuple[str, ...] = ()): 
        allowed_types = {control_type.lower() for control_type in control_types}
        for wrapper in self._iter_wrapper_tree(root):
            control_type = self._wrapper_control_type(wrapper).lower()
            if allowed_types and control_type not in allowed_types:
                continue
            texts = (self._wrapper_text(wrapper), self._read_control_value(wrapper))
            if any(text and self._matches_patterns(text, patterns) for text in texts):
                return wrapper
        return None

    def _is_enabled(self, wrapper) -> bool | None:
        readers = (
            lambda: bool(wrapper.is_enabled()),
            lambda: bool(getattr(wrapper.element_info, "enabled")),
        )
        for reader in readers:
            try:
                return reader()
            except Exception:
                continue
        return None

    def _control_present(self, root, patterns: tuple[str, ...], *, control_types: tuple[str, ...] = ()) -> bool:
        return self._first_matching_control(root, patterns, control_types=control_types) is not None

    def _button_enabled(self, root, patterns: tuple[str, ...]) -> bool:
        button = self._first_matching_control(root, patterns, control_types=("Button",))
        if button is None:
            return False
        enabled = self._is_enabled(button)
        return True if enabled is None else enabled

    def _show_job_queue_tab(self, render_dialog) -> None:
        logger.info("Switching Kdenlive rendering panel to the Job Queue tab.")
        job_queue_tab = self._first_matching_control(
            render_dialog,
            KDENLIVE_JOB_QUEUE_TAB_PATTERNS,
            control_types=("TabItem", "Button", "Text"),
        )
        if job_queue_tab is None:
            logger.info("Kdenlive Job Queue tab was not found. Monitoring will continue on the current panel.")
            return
        try:
            click_control(job_queue_tab, post_click_sleep=0.3)
        except Exception as exc:
            logger.info("Could not switch to the Kdenlive Job Queue tab: %s", exc)

    def _read_progress_ratio(self, render_dialog) -> float | None:
        best_candidate = None
        for wrapper in self._iter_wrapper_tree(render_dialog):
            if self._wrapper_control_type(wrapper).lower() != "progressbar":
                continue
            try:
                rect = self._wrapper_rect(wrapper)
            except Exception:
                continue
            current_value = None
            maximum_value = None
            minimum_value = 0.0
            for reader in (
                lambda: float(wrapper.iface_range_value.CurrentValue),
                lambda: float(wrapper.get_value()),
            ):
                try:
                    current_value = reader()
                    break
                except Exception:
                    continue
            try:
                maximum_value = float(wrapper.iface_range_value.CurrentMaximum)
            except Exception:
                maximum_value = None
            try:
                minimum_value = float(wrapper.iface_range_value.CurrentMinimum)
            except Exception:
                minimum_value = 0.0
            if current_value is None or maximum_value is None or maximum_value <= minimum_value:
                continue
            ratio = (current_value - minimum_value) / (maximum_value - minimum_value)
            if not 0.0 <= ratio <= 1.05:
                continue
            score = ((rect.right - rect.left) * (rect.bottom - rect.top), rect.bottom, rect.left)
            clamped_ratio = max(0.0, min(ratio, 1.0))
            if best_candidate is None or score > best_candidate[0]:
                best_candidate = (score, clamped_ratio)
        if best_candidate is None:
            return None
        return best_candidate[1]

    def _render_still_active(self, render_dialog) -> tuple[bool, bool, bool, bool, float | None]:
        active_text_present = self._control_present(
            render_dialog,
            KDENLIVE_ACTIVE_RENDER_PATTERNS,
            control_types=("Text", "Pane", "Group", "Custom", "Document", "ListItem", "DataItem"),
        )
        abort_enabled = self._button_enabled(render_dialog, KDENLIVE_ABORT_JOB_PATTERNS)
        start_enabled = self._button_enabled(render_dialog, KDENLIVE_START_JOB_PATTERNS)
        clean_up_enabled = self._button_enabled(render_dialog, KDENLIVE_CLEAN_UP_PATTERNS)
        progress_ratio = self._read_progress_ratio(render_dialog)
        if progress_ratio is not None and progress_ratio < 0.999:
            active = True
        else:
            active = active_text_present or abort_enabled
        return active, abort_enabled, start_enabled, clean_up_enabled, progress_ratio

    def _render_marked_finished(self, render_dialog) -> bool:
        return self._control_present(
            render_dialog,
            KDENLIVE_FINISHED_RENDER_PATTERNS,
            control_types=("Text", "Pane", "Group", "Custom", "Document", "ListItem", "DataItem"),
        )

    def _open_render_dialog(self, window):
        logger.info("Opening Kdenlive rendering dialog with Ctrl+Enter.")
        bring_window_to_front(window, keep_topmost=False)
        send_hotkey("^{ENTER}")
        time.sleep(KDENLIVE_POST_ACTION_SLEEP_SECONDS)

        logger.info("Searching the desktop for the unique Kdenlive 'Render to File' button.")
        try:
            render_button = wait_for_desktop_text_control(
                KDENLIVE_RENDER_TO_FILE_PATTERNS,
                control_types=("Button",),
                timeout=KDENLIVE_RENDER_DIALOG_TIMEOUT_SECONDS,
                poll_interval=0.2,
            )
            logger.info("Desktop render-button search succeeded: %s", self._wrapper_text(render_button) or "<untitled>")
            render_container = self._find_render_container_from_anchor(render_button)
            assert render_container is not None, "Kdenlive rendering container could not be resolved from the Render to File button."
            self._resolve_render_dialog_top_level(render_container)
            return render_container
        except Exception as exc:
            logger.info(
                "Desktop render-button search did not succeed. Falling back to process-window render-container search. details=%s",
                exc,
            )

        last_error: Exception | None = None
        for candidate in self._iter_render_search_roots(window):
            try:
                render_container = self._locate_render_container(candidate, timeout=1.5)
                self._resolve_render_dialog_top_level(render_container)
                return render_container
            except Exception as exc:
                last_error = exc

        try:
            logger.info("Trying Kdenlive export action as a final fallback.")
            click_text_control(window, KDENLIVE_EXPORT_MENU_PATTERNS, control_types=("Button", "MenuItem"))
            time.sleep(KDENLIVE_POST_ACTION_SLEEP_SECONDS)
            for candidate in self._iter_render_search_roots(window):
                try:
                    render_container = self._locate_render_container(candidate, timeout=1.5)
                    self._resolve_render_dialog_top_level(render_container)
                    return render_container
                except Exception as exc:
                    last_error = exc
        except Exception as exc:
            last_error = exc

        raise AssertionError("Kdenlive rendering dialog was not found after shortcut and fallback searches.") from last_error

    def _focus_output_file_entry(self, render_dialog) -> None:
        label = wait_for_text_control(
            render_dialog,
            KDENLIVE_OUTPUT_FILE_LABEL_PATTERNS,
            control_types=("Text", "Pane", "Group"),
            timeout=KDENLIVE_CONTROL_TIMEOUT_SECONDS,
            poll_interval=0.1,
        )
        label_rect = self._wrapper_rect(label)
        candidates = []
        for wrapper in self._iter_wrapper_tree(render_dialog):
            control_type = self._wrapper_control_type(wrapper).lower()
            if control_type in {"button", "checkbox", "radiobutton", "tabitem", "menuitem"}:
                continue
            rect = self._wrapper_rect(wrapper)
            overlap = self._vertical_overlap(label_rect, rect)
            width = rect.right - rect.left
            if rect.left <= label_rect.right:
                continue
            if overlap <= 0:
                continue
            if width < 60:
                continue
            current_value = self._read_control_value(wrapper) or self._wrapper_text(wrapper)
            looks_like_path = any(token in current_value for token in (":\\", "\\", "/", ".mp4", ".mov", ".mkv"))
            score = (
                0 if looks_like_path else 1,
                -width,
                abs(rect.top - label_rect.top),
                rect.left,
            )
            candidates.append((score, wrapper, rect, current_value, control_type, width))

        try:
            top_level = render_dialog.top_level_parent()
        except Exception:
            top_level = render_dialog
        bring_window_to_front(top_level, keep_topmost=False)
        top_rect = self._wrapper_rect(top_level)

        if candidates:
            candidates.sort(key=lambda item: item[0])
            _, wrapper, rect, current_value, control_type, width = candidates[0]
            logger.info(
                "Focusing Kdenlive output-path row candidate. control_type=%s width=%s current_value=%s",
                control_type,
                width,
                current_value or "<empty>",
            )
            try:
                click_control(wrapper, post_click_sleep=0.25)
                return
            except Exception:
                click_x_abs = rect.left + max(10, min(40, width // 8))
                click_y_abs = rect.top + max(1, (rect.bottom - rect.top) // 2)
        else:
            render_button = self._find_render_to_file_button(render_dialog)
            button_rect = self._wrapper_rect(render_button)
            click_x_abs = min(label_rect.right + 180, button_rect.left - 120)
            click_y_abs = label_rect.top + max(1, (label_rect.bottom - label_rect.top) // 2)
            logger.info("Falling back to coordinate focus for the Kdenlive output-path row.")

        rel_x = max(1, min(click_x_abs - top_rect.left, (top_rect.right - top_rect.left) - 2))
        rel_y = max(1, min(click_y_abs - top_rect.top, (top_rect.bottom - top_rect.top) - 2))
        top_level.click_input(coords=(rel_x, rel_y))
        time.sleep(0.25)

    def _set_render_output_path(self, render_dialog, output_video_path: Path) -> None:
        output_video_path = output_video_path.resolve(strict=False)
        logger.info("Waiting for Kdenlive output-file edit control.")
        self._focus_output_file_entry(render_dialog)
        logger.info("Pasting Kdenlive render output path via clipboard: %s", output_video_path)
        paste_text_via_clipboard(str(output_video_path), replace_existing=True)
        render_length = wait_for_text_control(
            render_dialog,
            KDENLIVE_RENDER_LENGTH_PATTERNS,
            control_types=("Text", "Edit", "Pane"),
            timeout=KDENLIVE_CONTROL_TIMEOUT_SECONDS,
        )
        assert render_length.is_visible(), "Kdenlive render dialog did not show the rendered file length."

    def _start_render(self, render_dialog) -> None:
        try:
            bring_window_to_front(render_dialog.top_level_parent(), keep_topmost=False)
        except Exception:
            pass
        render_button = self._find_render_to_file_button(render_dialog)
        logger.info("Clicking Kdenlive 'Render to File' button: %s", self._wrapper_text(render_button) or "<untitled>")
        assert render_button.is_enabled(), "Kdenlive 'Render to File' button is disabled."
        click_control(render_button)
        time.sleep(0.2)

    def _confirm_render_transition_after_minimizing(self, render_dialog) -> None:
        overwrite_accepted = accept_overwrite_confirmation(timeout=0.4, owner_window=render_dialog, poll_interval=0.05)
        if not overwrite_accepted:
            self._accept_kdenlive_overwrite_dialog_if_present(render_dialog, timeout=0.8)

    def _wait_for_render_completion(self, render_dialog, output_video_path: Path) -> None:
        output_video_path = output_video_path.resolve(strict=False)
        logger.info("Validating Kdenlive output directory before render wait.")
        assert output_video_path.parent.exists(), f"Output directory disappeared: {output_video_path.parent}"

        stable_rounds = 0
        inactive_rounds = 0
        finished_without_output_rounds = 0
        last_size = -1
        deadline = time.monotonic() + 7200.0
        progress_logger = _WaitProgressLogger("Kdenlive render")
        while time.monotonic() < deadline:
            current_size = output_video_path.stat().st_size if output_video_path.exists() else 0
            if current_size > 0 and current_size == last_size:
                stable_rounds += 1
            else:
                stable_rounds = 0
            last_size = current_size

            render_active, abort_enabled, start_enabled, clean_up_enabled, progress_ratio = self._render_still_active(render_dialog)
            render_finished = self._render_marked_finished(render_dialog)
            if render_active:
                inactive_rounds = 0
            else:
                inactive_rounds += 1
            if render_finished and current_size < KDENLIVE_MIN_OUTPUT_BYTES:
                finished_without_output_rounds += 1
            else:
                finished_without_output_rounds = 0
            progress_logger.maybe_log(
                detail=(
                    f"output={output_video_path.name} size_bytes={current_size} stable_rounds={stable_rounds} "
                    f"inactive_rounds={inactive_rounds} active={render_active} finished={render_finished} "
                    f"progress_ratio={'n/a' if progress_ratio is None else f'{progress_ratio:.3f}'}"
                )
            )

            if render_finished and current_size >= KDENLIVE_MIN_OUTPUT_BYTES:
                logger.info("Kdenlive render completion detected from the Job Queue finished state. elapsed=%s", progress_logger.elapsed_text())
                break

            if render_finished and finished_without_output_rounds >= 5:
                raise AssertionError(
                    "Kdenlive reported that rendering finished, but the expected output file was not found: "
                    f"{output_video_path}"
                )

            if (
                current_size >= KDENLIVE_MIN_OUTPUT_BYTES
                and stable_rounds >= KDENLIVE_FALLBACK_COMPLETION_STABLE_ROUNDS
                and inactive_rounds >= KDENLIVE_FALLBACK_COMPLETION_IDLE_ROUNDS
                and not render_active
            ):
                logger.info(
                    "Kdenlive render completion detected from the stable-output fallback. "
                    "elapsed=%s size_bytes=%d stable_rounds=%d inactive_rounds=%d",
                    progress_logger.elapsed_text(),
                    current_size,
                    stable_rounds,
                    inactive_rounds,
                )
                break

            if (
                current_size >= KDENLIVE_MIN_OUTPUT_BYTES
                and stable_rounds >= KDENLIVE_RENDER_STABLE_ROUNDS
                and inactive_rounds >= KDENLIVE_UI_IDLE_ROUNDS
                and (clean_up_enabled or start_enabled or (progress_ratio is not None and progress_ratio >= 0.999))
            ):
                logger.info("Kdenlive render completion conditions satisfied. elapsed=%s", progress_logger.elapsed_text())
                break
            time.sleep(KDENLIVE_RENDER_POLL_SECONDS)
        else:
            raise AssertionError(f"Kdenlive render did not finish within timeout: {output_video_path}")

        logger.info("Validating final Kdenlive output file.")
        assert output_video_path.exists() and output_video_path.stat().st_size >= KDENLIVE_MIN_OUTPUT_BYTES, (
            f"Kdenlive render output was not created correctly: {output_video_path}"
        )
def build_operator(software: str) -> SoftwareOperator:
    assert software in OPERATION_PROFILES, f"No software operation profile defined for {software}"
    if software == "shotcut":
        return ShotcutOperator(OPERATION_PROFILES[software])
    if software == "avidemux":
        return AvidemuxOperator(OPERATION_PROFILES[software])
    if software == "handbrake":
        return HandBrakeOperator(OPERATION_PROFILES[software])
    if software == "shutter_encoder":
        return ShutterEncoderOperator(OPERATION_PROFILES[software])
    if software == "kdenlive":
        return KdenliveOperator(OPERATION_PROFILES[software])
    return SoftwareOperator(OPERATION_PROFILES[software])


def run_shotcut_operation(input_video_path: Path, output_video_path: Path) -> None:
    build_operator("shotcut").perform(input_video_path, output_video_path)


def run_kdenlive_operation(input_video_path: Path, output_video_path: Path) -> None:
    build_operator("kdenlive").perform(input_video_path, output_video_path)


def run_shutter_encoder_operation(input_video_path: Path, output_video_path: Path) -> None:
    build_operator("shutter_encoder").perform(input_video_path, output_video_path)


def run_avidemux_operation(input_video_path: Path, output_video_path: Path) -> None:
    build_operator("avidemux").perform(input_video_path, output_video_path)


def run_handbrake_operation(input_video_path: Path, output_video_path: Path) -> None:
    build_operator("handbrake").perform(input_video_path, output_video_path)










