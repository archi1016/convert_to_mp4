VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPanelLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const LV_TIME = 0
Const LV_STATUS = 1
Const LV_FILE = 2

Dim DrawPanel As PictureBox

Dim LV As cListview

Private Sub Class_Initialize()
    Set LV = New cListview
End Sub

Private Sub Class_Terminate()
    Set LV = Nothing
    
    Set DrawPanel = Nothing
End Sub

Public Sub Init(drPanel As PictureBox)
    Set DrawPanel = drPanel
    
    DrawPanel.Visible = False
    
    With LV
        .Create DrawPanel.hWnd, 0, 0, 100, 100
        .InsertColumn "®É¶¡", LV_TIME, 120
        .InsertColumn "ª¬ºA", LV_STATUS, 120
        .InsertColumn "ÀÉ®×", LV_FILE, 400
    End With
End Sub

Public Sub MoveCtrl()
    Dim T As Long
    Dim W As Long
    Dim H As Long
    
    W = DrawPanel.Width - DEFAULT_PANEL_MARGIN * 2
    T = DEFAULT_PANEL_MARGIN
    H = DrawPanel.Height - T - DEFAULT_PANEL_MARGIN
    If H < 128 Then H = 128
    LV.Move DEFAULT_PANEL_MARGIN, T, W, H
End Sub

Public Sub InsertLog(ByVal sStatus As String, ByVal sFile As String)
    Dim C As Long
    
    With LV
        .InsertItemByPos 0, GetNowTimeStr, 0, -1
        .SetItemText 0, LV_STATUS, sStatus
        .SetItemText 0, LV_FILE, sFile
        
        C = .GetItemCount
        If C > 256 Then
            .DeleteItem C - 1
        End If
    End With
    
    
End Sub
