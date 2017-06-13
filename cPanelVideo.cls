VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPanelVideo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const TIMES_GET_VIDEO_INFO As Long = 2&
Const TIMES_LAUNCH_CONVERT As Long = 11&

Dim DrawPanel As PictureBox
Dim BtnAdd As CommandButton
Dim BtnView As CommandButton
Dim BtnCancel As CommandButton
Dim hWndOfForm As Long

Dim LV As cListview

Dim CountForGetVideoInfo As Long
Dim CountForLaunchConvert As Long

Dim UidCount As Long
Dim hProcess As Long

Private Sub Class_Initialize()
    Set LV = New cListview
    
    UidCount = 1001
    hProcess = 0
    CountForGetVideoInfo = 0
    CountForLaunchConvert = 0
End Sub

Private Sub Class_Terminate()
    SetWindowLong LV.hWnd, GWL_WNDPROC, OldFilesViewerProc
    Set LV = Nothing
    
    Set DrawPanel = Nothing
    Set BtnAdd = Nothing
    Set BtnView = Nothing
    Set BtnCancel = Nothing
End Sub

Public Sub Init(ByVal hWnd As Long, drPanel As PictureBox, bnAdd As CommandButton, bnView As CommandButton, bnCancel As CommandButton)
    Dim T As Long
    
    hWndOfForm = hWnd
    Set DrawPanel = drPanel
    Set BtnAdd = bnAdd
    Set BtnView = bnView
    Set BtnCancel = bnCancel
    
    DrawPanel.Visible = False
    
    T = DEFAULT_PANEL_MARGIN
    BtnAdd.Move DEFAULT_PANEL_MARGIN, T, DEFAULT_BUTTON_WIDTH, DEFAULT_BUTTON_HEIGHT
    BtnView.Move BtnAdd.Left + BtnAdd.Width + DEFAULT_PANEL_MARGIN, T, DEFAULT_BUTTON_WIDTH, DEFAULT_BUTTON_HEIGHT
    BtnCancel.Move 0, T, DEFAULT_BUTTON_WIDTH, DEFAULT_BUTTON_HEIGHT
    
    With LV
        .CreateMulti DrawPanel.hWnd, 0, 0, 100, 100
        .InsertColumn "來源檔案", LV_VIDEO_FILE, 200
        .InsertColumn "UID", LV_VIDEO_UID, 0
        .InsertColumn "TIMER", LV_VIDEO_TIMER, 0
        .InsertColumnAlignRight "長度", LV_VIDEO_DURATION, 100
        .InsertColumn "編碼", LV_VIDEO_VIDEO_CODEC, 100
        .InsertColumnAlignRight "寬", LV_VIDEO_VIDEO_WIDTH, 70
        .InsertColumnAlignRight "高", LV_VIDEO_VIDEO_HEIGHT, 70
        .InsertColumnAlignRight "張數", LV_VIDEO_VIDEO_FPS, 70
        .InsertColumnAlignRight "碼率", LV_VIDEO_VIDEO_BITRATE, 70
        .InsertColumn "編碼", LV_VIDEO_AUDIO_CODEC, 100
        .InsertColumn "聲道", LV_VIDEO_AUDIO_CHANNEL, 80
        .InsertColumnAlignRight "取樣", LV_VIDEO_AUDIO_HZ, 80
        .InsertColumnAlignRight "碼率", LV_VIDEO_AUDIO_BITRATE, 70
        .InsertColumn "來源路徑", LV_VIDEO_PATH, 400
    End With
    
    DragAcceptFiles LV.hWnd, &HFFFFFFFF
    Call ChangeMessageFliter(LV.hWnd, WM_DROPFILES)
    OldFilesViewerProc = SetWindowLong(LV.hWnd, GWL_WNDPROC, ReturnAddressOfFunction(AddressOf NewFilesViewerProc))
End Sub

Public Sub MoveCtrl()
    Dim T As Long
    Dim W As Long
    Dim H As Long
    
    W = DrawPanel.Width - DEFAULT_PANEL_MARGIN * 2
    T = BtnAdd.Top + BtnAdd.Height + DEFAULT_PANEL_MARGIN
    BtnCancel.Left = DrawPanel.Width - BtnCancel.Width - DEFAULT_PANEL_MARGIN
    
    H = DrawPanel.Height - T - DEFAULT_PANEL_MARGIN
    If H < 128 Then H = 128
    LV.Move DEFAULT_PANEL_MARGIN, T, W, H
End Sub

Public Sub CheckLvDropFiles(ByVal hwndFrom As Long, ByVal hDrop As Long)
    Dim C As Long
    Dim I As Long
    Dim F As String
    
    If hwndFrom = LV.hWnd Then
        C = DragQueryFileW(hDrop, &HFFFFFFFF, 0, 0)
        If C > 0 Then
            ReDim FFS(C - 1)
            For I = 0 To (C - 1)
                F = String$(MAX_PATH, vbNullChar)
                If 0 <> DragQueryFileW(hDrop, I, StrPtr(F), MAX_PATH) Then
                    F = StrCutNull(F)
                    Call AddVideoFile(F)
                End If
            Next
        End If

        DragFinish hDrop
    End If
End Sub

Public Sub Add()
    Dim FFS() As String
    Dim FFC As Long
    Dim I As Long
    
    If GetSelectedFiles(hWndOfForm, "請選取影片檔", "任何檔案", "*.*", FFS, FFC) Then
        For I = 0 To (FFC - 1)
            Call AddVideoFile(FFS(I))
        Next
    End If
    
    Erase FFS
End Sub

Public Sub Cancel()
    Dim nIndexs() As Long
    Dim nIndexsCount As Long
    Dim I As Long
    
    If LV.GetSelectedIndexs(nIndexs, nIndexsCount) Then
        If MsgQuestion("取消選取的 " + CStr(nIndexsCount) + " 個檔案？") Then
            For I = (nIndexsCount - 1) To 0 Step -1
                LV.DeleteItem nIndexs(I)
            Next
        End If
    End If
End Sub

Public Function GetVideoFile(ByVal nIndex As Long) As String
    GetVideoFile = ""
    
    If INVALID_HANDLE_VALUE <> nIndex Then
        With LV
            GetVideoFile = .GetItemText(nIndex, LV_VIDEO_PATH) + "\" + .GetItemText(nIndex, LV_VIDEO_FILE)
        End With
    End If
End Function

Public Sub SetVideoInfo(ByVal nIndex As Long, ByVal FP As String)
    Dim T As String
    Dim F() As String

    If INVALID_HANDLE_VALUE <> nIndex Then
        If ReadTextFromFile(FP, T) Then
            F = Split(T, vbTab)
            
            With LV
                .SetItemText nIndex, LV_VIDEO_DURATION, F(AV_DURATION)
                .SetItemText nIndex, LV_VIDEO_VIDEO_CODEC, F(AV_VIDEO_CODEC)
                .SetItemText nIndex, LV_VIDEO_VIDEO_WIDTH, F(AV_VIDEO_WIDTH)
                .SetItemText nIndex, LV_VIDEO_VIDEO_HEIGHT, F(AV_VIDEO_HEIGHT)
                .SetItemText nIndex, LV_VIDEO_VIDEO_FPS, F(AV_VIDEO_FPS)
                .SetItemText nIndex, LV_VIDEO_VIDEO_BITRATE, F(AV_VIDEO_BITRATE)
                .SetItemText nIndex, LV_VIDEO_AUDIO_CODEC, F(AV_AUDIO_CODEC)
                .SetItemText nIndex, LV_VIDEO_AUDIO_CHANNEL, F(AV_AUDIO_CHANNEL)
                .SetItemText nIndex, LV_VIDEO_AUDIO_HZ, F(AV_AUDIO_HZ)
                .SetItemText nIndex, LV_VIDEO_AUDIO_BITRATE, F(AV_AUDIO_BITRATE)
                .SetItemText nIndex, LV_VIDEO_TIMER, ""
            End With
        End If
    End If

    Erase F
End Sub

Public Sub RemoveVideoInfo(ByVal nIndex As Long)
    If INVALID_HANDLE_VALUE <> nIndex Then
        LV.DeleteItem nIndex
    End If
End Sub

Public Function ReturnHardLinkFile(ByVal nIndex As Long) As String
    ReturnHardLinkFile = ""
    
    If INVALID_HANDLE_VALUE <> nIndex Then
        With LV
            ReturnHardLinkFile = .GetItemText(nIndex, LV_VIDEO_PATH) + "\" + APPLICATION_ID + "." + .GetItemText(nIndex, LV_VIDEO_UID) + ".hardlink"
        End With
    End If
End Function

Public Function FindItemIndexFromUID(ByVal sID As String) As Long
    Dim C As Long
    Dim I As Long
    Dim A(GVI_ARGS_UBOUND) As String
    
    FindItemIndexFromUID = INVALID_HANDLE_VALUE
    
    With LV
        C = .GetItemCount
        If C > 0 Then
            For I = 0 To (C - 1)
                If sID = .GetItemText(I, LV_VIDEO_UID) Then
                    FindItemIndexFromUID = I
                    Exit For
                End If
            Next
        End If
    End With
End Function

Private Sub AddVideoFile(ByVal FP As String)
    Dim I As Long
    Dim FN As String
    Dim ExtName As String
    
    If IsFileExist(FP) Then
        I = InStrRev(FP, ".")
        If I > 0 Then
            FN = ReturnLastDirectoryName(FP)
            If INVALID_HANDLE_VALUE = LV.FindItem(FN) Then
                ExtName = "/" + LCase$(Right$(FP, Len(FP) - I)) + "/"
    
                If InStr(BaseConfig.VideoExtNames, ExtName) > 0 Then
                    LV.InsertItem FN, 0, -1
                    LV.SetLastInsertItemText LV_VIDEO_PATH, ReturnParentDirectory(FP)
                    LV.SetLastInsertItemText LV_VIDEO_UID, CStr(UidCount)
                    UidCount = UidCount + 1
                End If
            End If
        End If
    End If
End Sub

Private Sub NextGetVideInfo()
    Dim C As Long
    Dim I As Long
    Dim A(GVI_ARGS_UBOUND) As String
    Dim hlFP As String
    
    With LV
        C = .GetItemCount
        If C > 0 Then
            For I = 0 To (C - 1)
                If "" = .GetItemText(I, LV_VIDEO_VIDEO_CODEC) Then
                    If "" = .GetItemText(I, LV_VIDEO_AUDIO_CODEC) Then
                        If "" = .GetItemText(I, LV_VIDEO_TIMER) Then
                            hlFP = ReturnHardLinkFile(I)
                            If CreateHardLinkFile(.GetItemText(I, LV_VIDEO_PATH) + "\" + .GetItemText(I, LV_VIDEO_FILE), hlFP) Then
                                A(CMD_NAME) = CMD_GET_VIDEO_INFO
                                A(GVI_ARGS_HWND) = CStr(hWndOfForm)
                                A(GVI_ARGS_UID) = .GetItemText(I, LV_VIDEO_UID)
                                A(GVI_ARGS_FILE) = hlFP
                                
                                .SetItemText I, LV_VIDEO_TIMER, CStr(GetNowTimeMinutes)
                                
                                Call LaunchAnotherProcess(Join(A, CMD_SPLIT_CHAR))
                            End If

                            Exit For
                        End If
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Sub LaunchNextConvert(panelConverting As cPanelConverting, panelLog As cPanelLog)
    Dim C As Long
    
    With LV
        C = .GetItemCount
        If C > 0 Then
            If Not panelConverting.IsConverting Then
                If "" <> .GetItemText(0, LV_VIDEO_VIDEO_CODEC) Then
                    If "" <> .GetItemText(0, LV_VIDEO_AUDIO_CODEC) Then
                        panelConverting.CopyFrom LV, panelLog
                        .DeleteItem 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub Heartbeat(panelConverting As cPanelConverting, panelLog As cPanelLog)
    CountForGetVideoInfo = CountForGetVideoInfo + 1
    If CountForGetVideoInfo >= TIMES_GET_VIDEO_INFO Then
        Call NextGetVideInfo
        CountForGetVideoInfo = 0
    End If
    
    CountForLaunchConvert = CountForLaunchConvert + 1
    If CountForLaunchConvert >= TIMES_LAUNCH_CONVERT Then
        Call LaunchNextConvert(panelConverting, panelLog)
        CountForLaunchConvert = 0
    End If
End Sub


