Attribute VB_Name = "define_taskbarlist3"
Option Explicit

Public WM_TASKBAR_BUTTON_CREATED As Long

Public Const ITaskbarList3_TaskbarButtonCreated = "TaskbarButtonCreated"

Public Const ITaskbarList3_QueryInterface = 0
Public Const ITaskbarList3_AddRef = 1
Public Const ITaskbarList3_Release = 2
Public Const ITaskbarList3_HrInit = 3
Public Const ITaskbarList3_AddTab = 4
Public Const ITaskbarList3_DeleteTab = 5
Public Const ITaskbarList3_ActivateTab = 6
Public Const ITaskbarList3_SetActiveAlt = 7
Public Const ITaskbarList3_MarkFullscreenWindow = 8
Public Const ITaskbarList3_SetProgressValue = 9
Public Const ITaskbarList3_SetProgressState = 10
Public Const ITaskbarList3_RegistrerTab = 11
Public Const ITaskbarList3_UnregisterTab = 12
Public Const ITaskbarList3_SetTabOrder = 13
Public Const ITaskbarList3_SetTabActive = 14
Public Const ITaskbarList3_TumbBarAddButtons = 15
Public Const ITaskbarList3_TumbBarUpdateButtons = 16
Public Const ITaskbarList3_TumbBarSetImageList = 17
Public Const ITaskbarList3_SetOverlayIcon = 18
Public Const ITaskbarList3_SetTumbnailTooltip = 19
Public Const ITaskbarList3_SetTumbnailClip = 20

Public Const TBPF_NOPROGRESS = &H0
Public Const TBPF_INDETERMINATE = &H1
Public Const TBPF_NORMAL = &H2
Public Const TBPF_ERROR = &H4
Public Const TBPF_PAUSED = &H8

Public Type CHANGEFILTERSTRUCT
    cbSize As Long
    ExtStatus As Long
End Type

Public Const MSGFLTINFO_NONE = 0
Public Const MSGFLTINFO_ALREADYALLOWED_FORWND = 1
Public Const MSGFLTINFO_ALREADYDISALLOWED_FORWND = 2
Public Const MSGFLTINFO_ALLOWED_HIGHER = 3

Public Declare Function ChangeWindowMessageFilterEx Lib "user32" _
    (ByVal hWnd As Long, _
     ByVal message As Long, _
     ByVal action As Long, _
     pChangeFilterStruct As CHANGEFILTERSTRUCT) As Long

Public Const MSGFLT_RESET = 0
Public Const MSGFLT_ALLOW = 1
Public Const MSGFLT_DISALLOW = 2

Public Sub InitTaskbarButtonCreatedMessage()
    WM_TASKBAR_BUTTON_CREATED = RegisterWindowMessageW(StrPtr(ITaskbarList3_TaskbarButtonCreated))
End Sub

Public Sub ChangeMessageFliter(ByVal hWnd As Long, ByVal wMessage As Long)
    Dim cfs As CHANGEFILTERSTRUCT
    
    If CURRENT_WINDOWS_VERSION >= WINDOWS_VERSION_7 Then
        If 0 <> wMessage Then
            cfs.cbSize = Len(cfs)
            ChangeWindowMessageFilterEx hWnd, wMessage, MSGFLT_ALLOW, cfs
        End If
    End If
End Sub
