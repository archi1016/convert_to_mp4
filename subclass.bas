Attribute VB_Name = "subclass"
Option Explicit

Public OldServerFormProc As Long
Public OldFilesViewerProc As Long

Public Function NewServerFormProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case WM_TRAYICONCLICK
            If lParam = WM_RBUTTONUP Then
                ShowWindow hWnd, SW_SHOWMAXIMIZED
                SetForegroundWindow hWnd
            End If
            
        Case WM_TASKBARCREATED
            ServerForm.CreateTrayIcon
            
        Case WM_VIDEO_INFO
            ServerForm.WmVideoInfo wParam, lParam
        
        'Case WM_UPLOAD_DONE
        '    ServerForm.WmUploadDone
            
        'Case WM_ERROR_CODE
        '    ServerForm.WmErrorCode wParam, lParam
    End Select
    
    NewServerFormProc = CallWindowProc(OldServerFormProc, hWnd, uMsg, wParam, lParam)
End Function

Public Function NewFilesViewerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case WM_DROPFILES
            ServerForm.CheckLvDropFiles hWnd, wParam
            
    End Select
    
    NewFilesViewerProc = CallWindowProc(OldFilesViewerProc, hWnd, uMsg, wParam, lParam)
End Function

