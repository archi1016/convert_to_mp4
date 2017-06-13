Attribute VB_Name = "define"
Option Explicit

Public Const APPLICATION_ID = "convert_to_mp4"
Public Const APPLICATION_LOCATION = "\Hua-Yu Technology\" + APPLICATION_ID
Public Const APPLICATION_DB_EXT = ".db"

Public Const CONFIG_INI = "\" + APPLICATION_ID + ".ini"

Public Const CONFIG_KEY_OUTPUT_DIRECTORY = "output_directory"
Public Const CONFIG_KEY_VIDEO_EXT_NAMES = "video_ext_names"
Public Const CONFIG_KEY_CRF_OF_VIDEO = "crf_of_video"
Public Const CONFIG_KEY_BITRATE_OF_AUDIO = "bitrate_of_audio"
Public Const CONFIG_KEY_PRESET = "preset"
Public Const CONFIG_KEY_TUNE = "tune"
Public Const CONFIG_KEY_LEVEL = "level"
Public Const CONFIG_KEY_MAX_HD = "max_hd"
Public Const CONFIG_KEY_MAX_FPS = "max_fps"


Public WM_TASKBARCREATED As Long
Public Const TASKBAR_CREATED = "TaskbarCreated"
Public Const WM_TRAYICONCLICK = WM_USER + 168

Public Const WM_DROPFILES = &H233

Public Const WM_VIDEO_INFO = WM_USER + 100
'Public Const WM_UDP_NOTIFY = WM_USER + 161
'Public Const WM_NOTIFY_AGAIN = WM_USER + 163

'Public Const MASTER_SERVICE_LISTEN_PORT = "9981"
'Public Const NOTIFY_FROM_PORT = "3385"

Public Const CMD_SPLIT_CHAR = ","
Public Const CMD_NAME = 0

'Public Const CMD_CREATE_STARTUP_SHORTCUT = "/CREATE_STARTUP_SHORTCUT"
'Public Const CMD_NOTIFY_ORDER = "/NOTIFY_ORDER"
'Public Const CMD_NOTIFY_CALL = "/NOTIFY_CALL"

Public Const CMD_GET_VIDEO_INFO = "/GET_VIDEO_INFO"

Public Const GVI_ARGS_UBOUND = 3
Public Const GVI_ARGS_HWND = 1
Public Const GVI_ARGS_UID = 2
Public Const GVI_ARGS_FILE = 3

Public Const WEB_CMD_SPLIT_CHAR = vbTab

Public Type BASE_CONFIG
    OutputDirectory As String
    VideoExtNames As String
    CrfOfVideo As Long
    BitrateOfAudio As Long
    Preset As String
    Tune As String
    level As String
    MaxHD As String
    MaxFPS As String
    FFmpegExe As String
    MPlayerExe As String
End Type

Public Const PIXELS_OF_720W As Currency = 345600
Public Const PIXELS_OF_960W As Currency = 518400
Public Const PIXELS_OF_1280W As Currency = 921600
Public Const PIXELS_OF_1920W As Currency = 2073600

Public Const MAX_HD_OF_360P = "360p"
Public Const MAX_HD_OF_480P = "480p"
Public Const MAX_HD_OF_540P = "540p"
Public Const MAX_HD_OF_720P = "720p"
Public Const MAX_HD_OF_1080P = "1080p"

Public Const MAX_FPS_OF_3000 = "30.00"


Public Const VIDEO_INFO_SUCCESS = &HEEEEE000
Public Const VIDEO_INFO_NOT_EXIST = &HEEEEE001
Public Const VIDEO_INFO_FFMPEG_FAIL = &HEEEEE002
Public Const VIDEO_INFO_MPLAYER_FAIL = &HEEEEE003
Public Const VIDEO_INFO_NO_VIDEO = &HEEEEE004
Public Const VIDEO_INFO_NO_AUDIO = &HEEEEE005


