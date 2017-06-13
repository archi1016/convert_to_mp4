VERSION 5.00
Begin VB.Form ServerForm 
   Caption         =   "錯誤"
   ClientHeight    =   8700
   ClientLeft      =   225
   ClientTop       =   915
   ClientWidth     =   13800
   BeginProperty Font 
      Name            =   "微軟正黑體"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ServerForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   580
   ScaleMode       =   3  '像素
   ScaleWidth      =   920
   StartUpPosition =   3  '系統預設值
   WindowState     =   2  '最大化
   Begin VB.PictureBox ConfigPanel 
      Appearance      =   0  '平面
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   8415
      Left            =   8280
      ScaleHeight     =   561
      ScaleMode       =   3  '像素
      ScaleWidth      =   689
      TabIndex        =   32
      Top             =   2520
      Width           =   10335
      Begin VB.ComboBox configLevel 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   240
         Style           =   2  '單純下拉式
         TabIndex        =   17
         Top             =   6120
         Width           =   2835
      End
      Begin VB.ComboBox configMaxFPS 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   3660
         Style           =   2  '單純下拉式
         TabIndex        =   22
         Top             =   7260
         Width           =   2835
      End
      Begin VB.ComboBox configMaxHD 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   3960
         Style           =   2  '單純下拉式
         TabIndex        =   20
         Top             =   6060
         Width           =   2835
      End
      Begin VB.ComboBox configTune 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   3780
         Style           =   2  '單純下拉式
         TabIndex        =   16
         Top             =   4680
         Width           =   2835
      End
      Begin VB.ComboBox configPreset 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   180
         Style           =   2  '單純下拉式
         TabIndex        =   14
         Top             =   4680
         Width           =   2835
      End
      Begin VB.TextBox configBitrateOfAudio 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   3780
         TabIndex        =   12
         Top             =   3240
         Width           =   2955
      End
      Begin VB.CommandButton configBrowser 
         Caption         =   "瀏覽... (&B)"
         Height          =   615
         Left            =   3540
         TabIndex        =   6
         Top             =   780
         Width           =   1995
      End
      Begin VB.CommandButton configSave 
         Caption         =   "儲存 (&L)"
         Height          =   615
         Left            =   6840
         TabIndex        =   23
         Top             =   1620
         Width           =   1995
      End
      Begin VB.TextBox configCrfOfVideo 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   240
         TabIndex        =   10
         Top             =   3120
         Width           =   2955
      End
      Begin VB.TextBox configVideoExtNames 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   0
         TabIndex        =   8
         Top             =   1920
         Width           =   4215
      End
      Begin VB.TextBox configOutputDirectory 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   300
         TabIndex        =   5
         Top             =   780
         Width           =   4815
      End
      Begin VB.Label configDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "指定級別 (&D)"
         Height          =   360
         Index           =   6
         Left            =   420
         TabIndex        =   18
         Top             =   5640
         Width           =   1605
      End
      Begin VB.Label configDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "最高影格 (&Z)"
         Height          =   360
         Index           =   8
         Left            =   3780
         TabIndex        =   21
         Top             =   6840
         Width           =   1575
      End
      Begin VB.Label configDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "最高畫質 (&F)"
         Height          =   360
         Index           =   7
         Left            =   3960
         TabIndex        =   19
         Top             =   5580
         Width           =   1545
      End
      Begin VB.Label configDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "影片類型 (&S)"
         Height          =   360
         Index           =   5
         Left            =   3960
         TabIndex        =   15
         Top             =   4200
         Width           =   1560
      End
      Begin VB.Label configDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "壓縮品質 (&A)"
         Height          =   360
         Index           =   4
         Left            =   360
         TabIndex        =   13
         Top             =   4200
         Width           =   1590
      End
      Begin VB.Label configDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "聲道編碼率 kbit/s (&R)"
         Height          =   360
         Index           =   3
         Left            =   3900
         TabIndex        =   11
         Top             =   2820
         Width           =   2700
      End
      Begin VB.Label configDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "壓縮畫值 CRF (&E)"
         Height          =   360
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   2700
         Width           =   2145
      End
      Begin VB.Label configDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "判斷為影片檔的副檔名 (&W)"
         Height          =   360
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   1440
         Width           =   3390
      End
      Begin VB.Label configDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "輸出存放資料夾 (&Q)"
         Height          =   360
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   300
         Width           =   2475
      End
   End
   Begin VB.PictureBox ConvertingPanel 
      Appearance      =   0  '平面
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   7440
      ScaleHeight     =   161
      ScaleMode       =   3  '像素
      ScaleWidth      =   681
      TabIndex        =   34
      Top             =   5820
      Width           =   10215
      Begin VB.CommandButton convertingTermination 
         Caption         =   "終止 (&R)"
         Height          =   615
         Left            =   660
         TabIndex        =   24
         Top             =   540
         Width           =   1995
      End
   End
   Begin VB.PictureBox VideoPanel 
      Appearance      =   0  '平面
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   -720
      ScaleHeight     =   161
      ScaleMode       =   3  '像素
      ScaleWidth      =   681
      TabIndex        =   31
      Top             =   1560
      Width           =   10215
      Begin VB.CommandButton videoView 
         Caption         =   "檢視 (&W)"
         Height          =   615
         Left            =   2820
         TabIndex        =   26
         Top             =   540
         Width           =   1995
      End
      Begin VB.CommandButton videoCancel 
         Caption         =   "取消 (&R)"
         Height          =   615
         Left            =   7560
         TabIndex        =   27
         Top             =   540
         Width           =   1995
      End
      Begin VB.CommandButton videoAdd 
         Caption         =   "增加 (&Q)"
         Height          =   615
         Left            =   660
         TabIndex        =   25
         Top             =   540
         Width           =   1995
      End
   End
   Begin VB.PictureBox SectionPanel 
      Appearance      =   0  '平面
      BackColor       =   &H80000010&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   5355
      Left            =   0
      ScaleHeight     =   357
      ScaleMode       =   3  '像素
      ScaleWidth      =   184
      TabIndex        =   30
      Top             =   720
      Width           =   2760
      Begin VB.OptionButton PanelTab 
         Caption         =   "配置 (&4)"
         Height          =   600
         Index           =   3
         Left            =   300
         Style           =   1  '圖片外觀
         TabIndex        =   3
         Top             =   3000
         Width           =   2160
      End
      Begin VB.OptionButton PanelTab 
         Caption         =   "記錄 (&3)"
         Height          =   600
         Index           =   2
         Left            =   300
         Style           =   1  '圖片外觀
         TabIndex        =   2
         Top             =   2100
         Width           =   2160
      End
      Begin VB.Timer TimerForFirstPanel 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   0
         Top             =   0
      End
      Begin VB.Timer TimerForNextPanel 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   600
         Top             =   0
      End
      Begin VB.OptionButton PanelTab 
         Caption         =   "轉檔中 (&2)"
         Height          =   600
         Index           =   1
         Left            =   300
         Style           =   1  '圖片外觀
         TabIndex        =   1
         Top             =   1200
         Width           =   2160
      End
      Begin VB.OptionButton PanelTab 
         Caption         =   "來源檔 (&1)"
         Height          =   600
         Index           =   0
         Left            =   300
         Style           =   1  '圖片外觀
         TabIndex        =   0
         Top             =   300
         Value           =   -1  'True
         Width           =   2160
      End
   End
   Begin VB.PictureBox LogPanel 
      Appearance      =   0  '平面
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   2280
      ScaleHeight     =   161
      ScaleMode       =   3  '像素
      ScaleWidth      =   265
      TabIndex        =   33
      Top             =   3720
      Width           =   3975
   End
   Begin VB.Timer TimerForHeartbeat 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4800
      Top             =   1980
   End
   Begin VB.Timer TimerForHide 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4260
      Top             =   1980
   End
   Begin VB.PictureBox TitlePanel 
      Appearance      =   0  '平面
      BackColor       =   &H00B24A09&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   0
      ScaleHeight     =   48
      ScaleMode       =   3  '像素
      ScaleWidth      =   353
      TabIndex        =   28
      Top             =   0
      Width           =   5295
      Begin VB.Image TitleLogo 
         Height          =   480
         Left            =   120
         Picture         =   "ServerForm.frx":1CFA
         Top             =   120
         Width           =   480
      End
      Begin VB.Label LabTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "LabTitle"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   18
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Left            =   720
         TabIndex        =   29
         Top             =   120
         Width           =   1395
      End
   End
   Begin VB.Menu menuFile 
      Caption         =   "檔案(&F)"
      Begin VB.Menu menuFileExit 
         Caption         =   "真正的關閉本服務 (&X)"
      End
   End
End
Attribute VB_Name = "ServerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const SECTION_VIDEO = 0
Const SECTION_CONVERTING = 1
Const SECTION_LOG = 2
Const SECTION_CONFIG = 3

Dim TYI As cTrayIcon

Dim panelVideo As cPanelVideo
Dim panelConverting As cPanelConverting
Dim panelLog As cPanelLog
Dim panelConfig As cPanelConfig
Dim panelChanger As cPanelChanger

Dim IsCanUnload As Boolean

Private Sub Form_Load()
    SendMessageW Me.hWnd, WM_SETICON, ICON_BIG, LoadResPicture(132, vbResIcon)
    
    Call LoadMemory
    Call LoadControl

    Me.Caption = App.ProductName + " [" + CStr(GetCurrentProcessId) + "]"
    LabTitle.Caption = Me.Caption
    App.Title = Me.Caption

    Call CreateTrayIcon
    WM_TASKBARCREATED = RegisterWindowMessageW(StrPtr(TASKBAR_CREATED))
    
    OldServerFormProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, ReturnAddressOfFunction(AddressOf NewServerFormProc))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If IsCanUnload Then
        SetWindowLong Me.hWnd, GWL_WNDPROC, OldServerFormProc
        
        Call FreeControl
        Call FreeMemory
    Else
        Me.Hide
        Cancel = 1
    End If
End Sub

Private Sub Form_Resize()
    Dim L As Long
    Dim T As Long
    Dim W As Long
    Dim H As Long
    
    If vbMinimized <> Me.WindowState Then
        TitlePanel.Width = Me.ScaleWidth
        
        T = TitlePanel.Height
        H = Me.ScaleHeight - T
        SectionPanel.Height = H
        If H >= 128 Then
            L = SectionPanel.Width
            W = Me.ScaleWidth - L
            
            VideoPanel.Move L, T, W, H
            ConvertingPanel.Move L, T, W, H
            LogPanel.Move L, T, W, H
            ConfigPanel.Move L, T, W, H
            
            panelVideo.MoveCtrl
            panelConverting.MoveCtrl
            panelLog.MoveCtrl
            panelConfig.MoveCtrl
        End If
    End If

End Sub

Private Sub LoadMemory()
    IsCanUnload = False
End Sub

Private Sub FreeMemory()
    
End Sub

Private Sub LoadControl()
    Set TYI = New cTrayIcon

    Set panelVideo = New cPanelVideo
    panelVideo.Init Me.hWnd, VideoPanel, videoAdd, videoView, videoCancel
    
    Set panelConverting = New cPanelConverting
    panelConverting.Init ConvertingPanel, convertingTermination
    
    Set panelLog = New cPanelLog
    panelLog.Init LogPanel
    
    Set panelConfig = New cPanelConfig
    panelConfig.Init ConfigPanel, configDescription, configOutputDirectory, configVideoExtNames, configCrfOfVideo, configBitrateOfAudio, configPreset, configTune, configLevel, configMaxHD, configMaxFPS, configBrowser, configSave
    
    Set panelChanger = New cPanelChanger
    panelChanger.Init SectionPanel, VideoPanel, TimerForFirstPanel, TimerForNextPanel
    
    panelChanger.ToFirst
    
    TimerForHeartbeat.Enabled = True
    'TimerForHide.Enabled = True
End Sub

Private Sub FreeControl()
    TYI.Remove Me.hWnd
        
    Set TYI = Nothing
    Set panelVideo = Nothing
    Set panelConverting = Nothing
    Set panelLog = Nothing
    Set panelConfig = Nothing
    Set panelChanger = Nothing
End Sub

Public Sub CreateTrayIcon()
    TYI.Add Me.hWnd, Me.Icon.Handle, Me.Caption
End Sub

Public Sub CheckLvDropFiles(ByVal hwndFrom As Long, ByVal wParam As Long)
    panelVideo.CheckLvDropFiles hwndFrom, wParam
End Sub

Public Sub WmVideoInfo(ByVal wParam As Long, ByVal lParam As Long)
    Dim IsSuccess As Boolean
    Dim FP As String
    Dim nIndex As Long
    Dim VideoFile As String
    Dim S As String
    
    IsSuccess = False
    FP = GetInfoFile(lParam)
    nIndex = panelVideo.FindItemIndexFromUID(CStr(lParam))

    If INVALID_HANDLE_VALUE <> nIndex Then
        VideoFile = panelVideo.GetVideoFile(nIndex)

        Select Case wParam
            Case VIDEO_INFO_SUCCESS
                panelVideo.SetVideoInfo nIndex, FP
                IsSuccess = True
                
            Case VIDEO_INFO_NOT_EXIST
                S = "檔案不存在"
            
            Case VIDEO_INFO_FFMPEG_FAIL
                S = "沒有FFmpge"
            
            Case VIDEO_INFO_NO_VIDEO
                S = "沒有影像軌"
            
            Case VIDEO_INFO_NO_AUDIO
                S = "沒有聲音軌"
            
            Case Else
                S = ""

        End Select
        
        DeleteFile panelVideo.ReturnHardLinkFile(nIndex)

        If Not IsSuccess Then
            panelVideo.RemoveVideoInfo nIndex
            panelLog.InsertLog S, VideoFile
        End If
    End If
    
    DeleteFile FP
End Sub

Private Sub configBrowser_Click()
    panelConfig.Browser Me.hWnd
End Sub

Private Sub configSave_Click()
    panelConfig.Save
End Sub

Private Sub convertingTermination_Click()
    panelConverting.Termination panelLog
End Sub

Private Sub menuFileExit_Click()
    IsCanUnload = True
    Unload Me
End Sub

Private Sub PanelTab_Click(Index As Integer)
    With panelChanger
        .Hide
        Select Case Index
            Case SECTION_VIDEO
                .SetNext VideoPanel
            
            Case SECTION_CONVERTING
                .SetNext ConvertingPanel
            
            Case SECTION_LOG
                .SetNext LogPanel
                
            Case SECTION_CONFIG
                .SetNext ConfigPanel
                
        End Select
        .ToNext
    End With
End Sub

Private Sub TimerForHeartbeat_Timer()
    panelVideo.Heartbeat panelConverting, panelLog
    panelConverting.Heartbeat panelLog
End Sub

Private Sub TimerForHide_Timer()
    TimerForHide.Enabled = False
    Me.Hide
End Sub

Private Sub TimerForFirstPanel_Timer()
    panelChanger.TimerFirst
End Sub

Private Sub TimerForNextPanel_Timer()
    panelChanger.TimerNext
End Sub

Private Sub videoAdd_Click()
    panelVideo.Add
End Sub

Private Sub videoCancel_Click()
    panelVideo.Cancel
End Sub

Private Sub videoView_Click()
    Call ShellProgram("EXPLORER.EXE", "/n,""" + BaseConfig.OutputDirectory + """")
End Sub
