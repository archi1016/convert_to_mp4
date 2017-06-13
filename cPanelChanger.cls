VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPanelChanger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim DrawPanel As PictureBox
Dim TmrToFirst As Timer
Dim TmrToNext As Timer

Dim TimesCount As Long
Dim NextPanel As PictureBox

Private Sub Class_Terminate()
    Set DrawPanel = Nothing
    Set NextPanel = Nothing
    Set TmrToFirst = Nothing
    Set TmrToNext = Nothing
End Sub

Public Sub Init(drPanel As PictureBox, firstPanel As PictureBox, tmToFirst As Timer, tmToNext As Timer)
    Set DrawPanel = drPanel
    Set NextPanel = firstPanel
    Set TmrToFirst = tmToFirst
    Set TmrToNext = tmToNext
    
    DrawPanel.Visible = False
    DrawPanel.Width = SECTION_PANEL_WIDTH
    
    With TmrToFirst
        .Enabled = False
        .Interval = 100
    End With
    With TmrToNext
        .Enabled = False
        .Interval = DEFAULT_PANEL_TIMER_INTERVAL
    End With
    
End Sub

Public Sub Hide()
    NextPanel.Visible = False
    DoEvents
End Sub

Public Sub SetNext(nxPanel As PictureBox)
    Set NextPanel = nxPanel
End Sub

Public Sub ToFirst()
    DrawPanel.Visible = True
    TmrToFirst.Enabled = True
End Sub

Public Sub ToNext()
    TimesCount = DEFAULT_PANEL_MOVE_TIMES
    NextPanel.Left = SECTION_PANEL_WIDTH + DEFAULT_PANEL_MOVE_STEPS * TimesCount
    NextPanel.Visible = True
    DoEvents
    TmrToNext.Enabled = True
End Sub

Public Sub TimerFirst()
    TmrToFirst.Enabled = False
    NextPanel.Visible = True
End Sub

Public Sub TimerNext()
    TimesCount = TimesCount - 1
    
    NextPanel.Left = SECTION_PANEL_WIDTH + DEFAULT_PANEL_MOVE_STEPS * TimesCount
    DoEvents
    
    If 0 >= TimesCount Then
        TmrToNext.Enabled = False
    End If
End Sub

