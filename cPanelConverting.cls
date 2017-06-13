VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPanelConverting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateHardLinkW Lib "Kernel32" _
    (ByVal lpFileName As Long, _
     ByVal lpExistingFileName As Long, _
     ByVal lpSecurityAttributes As Long) As Long
     
Const TIMES_PROCESS_DONE As Long = 15&

Dim DrawPanel As PictureBox
Dim BtnTermination As CommandButton

Dim LV As cListview

Dim CountForProcessDone As Long

Private Sub Class_Initialize()
    Set LV = New cListview
    
    CountForProcessDone = 0
End Sub

Private Sub Class_Terminate()
    Set LV = Nothing
    
    Set DrawPanel = Nothing
    Set BtnTermination = Nothing
End Sub

Public Sub Init(drPanel As PictureBox, bnTermination As CommandButton)
    Dim T As Long
    
    Set DrawPanel = drPanel
    Set BtnTermination = bnTermination
    
    DrawPanel.Visible = False
    
    T = DEFAULT_PANEL_MARGIN
    BtnTermination.Move DEFAULT_PANEL_MARGIN, T, DEFAULT_BUTTON_WIDTH, DEFAULT_BUTTON_HEIGHT
    
    With LV
        .Create DrawPanel.hWnd, 0, 0, 100, 100
        .InsertColumn "來源檔案", LV_CONVERTING_FILE, 240
        .InsertColumn "UID", LV_CONVERTING_UID, 0
        .InsertColumn "PROCESS_HANDLE", LV_CONVERTING_PROCESS_HANDLE, 0
        .InsertColumn "長度", LV_CONVERTING_DURATION, 100
        .InsertColumn "編碼", LV_CONVERTING_VIDEO_CODEC, 0
        .InsertColumnAlignRight "寬", LV_CONVERTING_VIDEO_WIDTH, 70
        .InsertColumnAlignRight "高", LV_CONVERTING_VIDEO_HEIGHT, 70
        .InsertColumnAlignRight "張數", LV_CONVERTING_VIDEO_FPS, 70
        .InsertColumnAlignRight "編碼率", LV_CONVERTING_VIDEO_BITRATE, 0
        .InsertColumn "編碼", LV_CONVERTING_AUDIO_CODEC, 0
        .InsertColumn "聲道", LV_CONVERTING_AUDIO_CHANNEL, 0
        .InsertColumnAlignRight "取樣", LV_CONVERTING_AUDIO_HZ, 0
        .InsertColumnAlignRight "編碼率", LV_CONVERTING_AUDIO_BITRATE, 0
        .InsertColumn "來源路徑", LV_CONVERTING_PATH, 400
        .InsertColumn "PASS_2", LV_CONVERTING_PASS_2, 0
    End With
End Sub

Public Sub MoveCtrl()
    Dim T As Long
    Dim W As Long
    Dim H As Long
    
    W = DrawPanel.Width - DEFAULT_PANEL_MARGIN * 2
    T = BtnTermination.Top + BtnTermination.Height + DEFAULT_PANEL_MARGIN
    H = DrawPanel.Height - T - DEFAULT_PANEL_MARGIN
    If H < 128 Then H = 128
    LV.Move DEFAULT_PANEL_MARGIN, T, W, H
End Sub

Public Sub CopyFrom(fromLV As cListview, panelLog As cPanelLog)
    Dim I As Long
    Dim AvInfo(AV_UBOUND) As String
    Dim srcFP As String
    Dim hlFP As String
    Dim FM As New cFFmpeg
    Dim convArgs As String
    Dim hProcess As Long
    
    With fromLV
        AvInfo(AV_DURATION) = .GetItemText(0, LV_VIDEO_DURATION)
        AvInfo(AV_VIDEO_CODEC) = .GetItemText(0, LV_VIDEO_VIDEO_CODEC)
        AvInfo(AV_VIDEO_WIDTH) = .GetItemText(0, LV_VIDEO_VIDEO_WIDTH)
        AvInfo(AV_VIDEO_HEIGHT) = .GetItemText(0, LV_VIDEO_VIDEO_HEIGHT)
        AvInfo(AV_VIDEO_FPS) = .GetItemText(0, LV_VIDEO_VIDEO_FPS)
        AvInfo(AV_VIDEO_BITRATE) = .GetItemText(0, LV_VIDEO_VIDEO_BITRATE)
        AvInfo(AV_AUDIO_CODEC) = .GetItemText(0, LV_VIDEO_AUDIO_CODEC)
        AvInfo(AV_AUDIO_CHANNEL) = .GetItemText(0, LV_VIDEO_AUDIO_CHANNEL)
        AvInfo(AV_AUDIO_HZ) = .GetItemText(0, LV_VIDEO_AUDIO_HZ)
        AvInfo(AV_AUDIO_BITRATE) = .GetItemText(0, LV_VIDEO_AUDIO_BITRATE)
        
        I = LV.InsertItem(.GetItemText(0, LV_VIDEO_FILE), 0, -1)
        LV.SetItemText I, LV_CONVERTING_UID, .GetItemText(0, LV_VIDEO_UID)
        LV.SetItemText I, LV_CONVERTING_DURATION, AvInfo(AV_DURATION)
        LV.SetItemText I, LV_CONVERTING_VIDEO_CODEC, AvInfo(AV_VIDEO_CODEC)
        LV.SetItemText I, LV_CONVERTING_VIDEO_WIDTH, AvInfo(AV_VIDEO_WIDTH)
        LV.SetItemText I, LV_CONVERTING_VIDEO_HEIGHT, AvInfo(AV_VIDEO_HEIGHT)
        LV.SetItemText I, LV_CONVERTING_VIDEO_FPS, AvInfo(AV_VIDEO_FPS)
        LV.SetItemText I, LV_CONVERTING_VIDEO_BITRATE, AvInfo(AV_VIDEO_BITRATE)
        LV.SetItemText I, LV_CONVERTING_AUDIO_CODEC, AvInfo(AV_AUDIO_CODEC)
        LV.SetItemText I, LV_CONVERTING_AUDIO_CHANNEL, AvInfo(AV_AUDIO_CHANNEL)
        LV.SetItemText I, LV_CONVERTING_AUDIO_HZ, AvInfo(AV_AUDIO_HZ)
        LV.SetItemText I, LV_CONVERTING_AUDIO_BITRATE, AvInfo(AV_AUDIO_BITRATE)
        LV.SetItemText I, LV_CONVERTING_PATH, .GetItemText(0, LV_VIDEO_PATH)
    End With
    
    hlFP = ReturnHardLinkFile(I)
    srcFP = LV.GetItemText(I, LV_CONVERTING_PATH) + "\" + LV.GetItemText(I, LV_CONVERTING_FILE)
    
    If CreateHardLinkFile(srcFP, hlFP) Then
        FM.GetConvertArguments hlFP, ReturnTempMp4File(I), AvInfo, convArgs
        hProcess = CreateProgram(BaseConfig.FFmpegExe, convArgs, SW_SHOWMINNOACTIVE)
        If 0 <> hProcess Then
            LV.SetItemText I, LV_CONVERTING_PROCESS_HANDLE, CStr(hProcess)
            panelLog.InsertLog "開始轉檔", convArgs
        Else
            panelLog.InsertLog "轉檔失敗", BaseConfig.FFmpegExe
        End If
    End If
    
    Set FM = Nothing
End Sub

Public Sub Termination(panelLog As cPanelLog)
    Dim I As Long
    Dim hProcess As Long
    
    With LV
        I = .GetSelectedIndex
        If INVALID_HANDLE_VALUE <> I Then
            If MsgQuestion("確定終止 " + .GetItemText(I, LV_CONVERTING_FILE) + " 的轉換？") Then
                hProcess = CLng(.GetItemText(I, LV_CONVERTING_PROCESS_HANDLE))
                TerminateProcess hProcess, 0
                CloseHandle hProcess
                DoEvents
                
                DeleteFile ReturnHardLinkFile(I)
                DeleteFile ReturnTempMp4File(I)
        
                panelLog.InsertLog "終止", .GetItemText(I, LV_CONVERTING_FILE)
        
                .DeleteItem I
            End If
        End If
    End With
End Sub

Public Function IsConverting() As Boolean
    IsConverting = (LV.GetItemCount > 0)
End Function

Public Function ReturnHardLinkFile(ByVal nIndex As Long) As String
    ReturnHardLinkFile = ""
    
    If INVALID_HANDLE_VALUE <> nIndex Then
        With LV
            ReturnHardLinkFile = .GetItemText(nIndex, LV_CONVERTING_PATH) + "\" + APPLICATION_ID + "." + CStr(GetCurrentProcessId) + "." + .GetItemText(nIndex, LV_CONVERTING_UID) + ".hardlink"
        End With
    End If
End Function

Public Function ReturnTempMp4File(ByVal nIndex As Long) As String
    ReturnTempMp4File = ""
    
    If INVALID_HANDLE_VALUE <> nIndex Then
        With LV
            ReturnTempMp4File = BaseConfig.OutputDirectory + "\" + APPLICATION_ID + "." + CStr(GetCurrentProcessId) + "." + .GetItemText(nIndex, LV_CONVERTING_UID) + ".converting"
        End With
    End If
End Function

Private Sub CheckProcessDone(panelLog As cPanelLog)
    Dim C As Long
    Dim I As Long
    Dim S As String
    Dim hProcess As Long
    Dim sPass2 As String
    Dim TmpMp4File As String
    Dim CurMp4File As String
    
    With LV
        C = .GetItemCount
        For I = (C - 1) To 0
            S = .GetItemText(I, LV_CONVERTING_PROCESS_HANDLE)
            If "" <> S Then
                hProcess = CLng(S)
                
                If WAIT_TIMEOUT <> WaitForSingleObject(hProcess, 0) Then
                    sPass2 = .GetItemText(I, LV_CONVERTING_PASS_2)
                    If "" = sPass2 Then
                        CloseHandle hProcess
                        
                        DeleteFile ReturnHardLinkFile(I)
                        
                        TmpMp4File = ReturnTempMp4File(I)
                        If GetFileSize64(TmpMp4File) > 0 Then
                            CurMp4File = BaseConfig.OutputDirectory + "\" + .GetItemText(I, LV_CONVERTING_FILE) + ".mp4"
                            MoveFile TmpMp4File, CurMp4File
                            panelLog.InsertLog "轉檔完成", CurMp4File
                        Else
                            DeleteFile TmpMp4File
                            panelLog.InsertLog "轉檔失敗", .GetItemText(I, LV_CONVERTING_PATH) + "\" + .GetItemText(I, LV_CONVERTING_FILE)
                        End If
                        
                        .DeleteItem I
                    Else
                        hProcess = CreateProgram(BaseConfig.FFmpegExe, sPass2, SW_SHOWMINNOACTIVE)
                        If 0 <> hProcess Then
                            LV.SetItemText I, LV_CONVERTING_PROCESS_HANDLE, CStr(hProcess)
                            LV.SetItemText I, LV_CONVERTING_PASS_2, ""
                            panelLog.InsertLog "開始轉檔", sPass2
                        Else
                            panelLog.InsertLog "轉檔失敗", BaseConfig.FFmpegExe
                        End If
                    End If
                End If
            End If
        Next
    End With
End Sub


Public Sub Heartbeat(panelLog As cPanelLog)
    CountForProcessDone = CountForProcessDone + 1
    If CountForProcessDone >= TIMES_PROCESS_DONE Then
        Call CheckProcessDone(panelLog)
        CountForProcessDone = 0
    End If
End Sub




