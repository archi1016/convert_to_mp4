VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPanelConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim DrawPanel As PictureBox
Dim LabDescriptions As Variant
Dim TxtOutputDirectory As TextBox
Dim TxtVideoExtNames As TextBox
Dim TxtCrfOfVideo As TextBox
Dim TxtBitrateOfAudio As TextBox
Dim CmbPreset As ComboBox
Dim CmbTune As ComboBox
Dim CmbLevel As ComboBox
Dim CmbMaxHD As ComboBox
Dim CmbMaxFPS As ComboBox
Dim BtnBrowser As CommandButton
Dim BtnSave As CommandButton

Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Set DrawPanel = Nothing
    Set LabDescriptions = Nothing
    Set TxtOutputDirectory = Nothing
    Set TxtVideoExtNames = Nothing
    Set TxtCrfOfVideo = Nothing
    Set TxtBitrateOfAudio = Nothing
    Set CmbPreset = Nothing
    Set CmbTune = Nothing
    Set CmbLevel = Nothing
    Set CmbMaxHD = Nothing
    Set CmbMaxFPS = Nothing
    Set BtnBrowser = Nothing
    Set BtnSave = Nothing
End Sub

Public Sub Init(drPanel As PictureBox, lbDescriptions As Variant, ttOutputDirectory As TextBox, ttVideoExtNames As TextBox, ttCrfOfvideo As TextBox, ttBitrateOfAudio As TextBox, cbPreset As ComboBox, cbTune As ComboBox, cbLevel As ComboBox, cbMaxHD As ComboBox, cbMaxFPS As ComboBox, bnBrowser As CommandButton, bnSave As CommandButton)
    Set DrawPanel = drPanel
    Set LabDescriptions = lbDescriptions
    Set TxtOutputDirectory = ttOutputDirectory
    Set TxtVideoExtNames = ttVideoExtNames
    Set TxtCrfOfVideo = ttCrfOfvideo
    Set TxtBitrateOfAudio = ttBitrateOfAudio
    Set CmbPreset = cbPreset
    Set CmbTune = cbTune
    Set CmbLevel = cbLevel
    Set CmbMaxHD = cbMaxHD
    Set CmbMaxFPS = cbMaxFPS
    Set BtnBrowser = bnBrowser
    Set BtnSave = bnSave
    
    DrawPanel.Visible = False
    
    With CmbPreset
        .AddItem "(無)"
        .AddItem "veryslow"
        .AddItem "slower"
        .AddItem "slow"
        .AddItem "medium"
        .AddItem "fast"
        .AddItem "faster"
        .AddItem "veryfast"
        .AddItem "superfast"
        .AddItem "ultrafast"
    End With
    
    With CmbTune
        .AddItem "(無)"
        .AddItem "film"
        .AddItem "animation"
        .AddItem "grain"
        .AddItem "stillimage"
        .AddItem "psnr"
        .AddItem "ssim"
        .AddItem "fastdecode"
        .AddItem "zerolatency"
    End With
    
    With CmbLevel
        .AddItem "(無)"
        .AddItem "Baseline 3.0"
        .AddItem "Baseline 3.1"
        .AddItem "Main 3.1"
        .AddItem "Main 4.0"
        .AddItem "High 4.0"
        .AddItem "High 4.1"
        .AddItem "High 4.2"
    End With
    
    With CmbMaxHD
        .AddItem "(無)"
        .AddItem MAX_HD_OF_360P
        .AddItem MAX_HD_OF_480P
        .AddItem MAX_HD_OF_540P
        .AddItem MAX_HD_OF_720P
        .AddItem MAX_HD_OF_1080P
    End With
    
    With CmbMaxFPS
        .AddItem "(無)"
        .AddItem MAX_FPS_OF_3000
    End With
    
    LabDescriptions(0).Move DEFAULT_PANEL_MARGIN, DEFAULT_PANEL_MARGIN
    TxtOutputDirectory.Move DEFAULT_PANEL_MARGIN, LabDescriptions(0).Top + LabDescriptions(0).Height + DEFAULT_PANEL_DESCRIPTION_MARGIN
    LabDescriptions(1).Move DEFAULT_PANEL_MARGIN, TxtOutputDirectory.Top + TxtOutputDirectory.Height + DEFAULT_PANEL_MARGIN
    TxtVideoExtNames.Move DEFAULT_PANEL_MARGIN, LabDescriptions(1).Top + LabDescriptions(1).Height + DEFAULT_PANEL_DESCRIPTION_MARGIN
    
    LabDescriptions(2).Move DEFAULT_PANEL_MARGIN, TxtVideoExtNames.Top + TxtVideoExtNames.Height + DEFAULT_PANEL_MARGIN
    TxtCrfOfVideo.Move DEFAULT_PANEL_MARGIN, LabDescriptions(2).Top + LabDescriptions(2).Height + DEFAULT_PANEL_DESCRIPTION_MARGIN, DEFAULT_BUTTON_WIDTH * 2
    LabDescriptions(3).Move TxtCrfOfVideo.Left + TxtCrfOfVideo.Width + DEFAULT_PANEL_MARGIN, LabDescriptions(2).Top
    TxtBitrateOfAudio.Move LabDescriptions(3).Left, LabDescriptions(3).Top + LabDescriptions(3).Height + DEFAULT_PANEL_DESCRIPTION_MARGIN, DEFAULT_BUTTON_WIDTH * 2
    
    LabDescriptions(4).Move DEFAULT_PANEL_MARGIN, TxtBitrateOfAudio.Top + TxtBitrateOfAudio.Height + DEFAULT_PANEL_MARGIN
    CmbPreset.Move DEFAULT_PANEL_MARGIN, LabDescriptions(4).Top + LabDescriptions(4).Height + DEFAULT_PANEL_DESCRIPTION_MARGIN, DEFAULT_BUTTON_WIDTH * 2
    LabDescriptions(5).Move CmbPreset.Left + CmbPreset.Width + DEFAULT_PANEL_MARGIN, LabDescriptions(4).Top
    CmbTune.Move LabDescriptions(5).Left, LabDescriptions(5).Top + LabDescriptions(5).Height + DEFAULT_PANEL_DESCRIPTION_MARGIN, DEFAULT_BUTTON_WIDTH * 2
    
    LabDescriptions(6).Move DEFAULT_PANEL_MARGIN, CmbTune.Top + CmbTune.Height + DEFAULT_PANEL_MARGIN
    CmbLevel.Move DEFAULT_PANEL_MARGIN, LabDescriptions(6).Top + LabDescriptions(6).Height + DEFAULT_PANEL_DESCRIPTION_MARGIN, DEFAULT_BUTTON_WIDTH * 2
    LabDescriptions(7).Move CmbLevel.Left + CmbLevel.Width + DEFAULT_PANEL_MARGIN, LabDescriptions(6).Top
    CmbMaxHD.Move LabDescriptions(7).Left, LabDescriptions(7).Top + LabDescriptions(7).Height + DEFAULT_PANEL_DESCRIPTION_MARGIN, DEFAULT_BUTTON_WIDTH * 2
    
    LabDescriptions(8).Move DEFAULT_PANEL_MARGIN, CmbMaxHD.Top + CmbMaxHD.Height + DEFAULT_PANEL_MARGIN
    CmbMaxFPS.Move LabDescriptions(8).Left, LabDescriptions(8).Top + LabDescriptions(8).Height + DEFAULT_PANEL_DESCRIPTION_MARGIN, DEFAULT_BUTTON_WIDTH * 2
    
    Call SetValues
End Sub

Public Sub MoveCtrl()
    Dim W As Long
    
    BtnBrowser.Move DrawPanel.Width - DEFAULT_BUTTON_WIDTH - DEFAULT_PANEL_MARGIN, TxtOutputDirectory.Top, DEFAULT_BUTTON_WIDTH, DEFAULT_BUTTON_HEIGHT
    W = BtnBrowser.Left - DEFAULT_PANEL_MARGIN * 2
    If W <= 200 Then W = 200
    TxtOutputDirectory.Width = W
    TxtVideoExtNames.Width = TxtOutputDirectory.Width
    
    BtnSave.Move BtnBrowser.Left, DrawPanel.Height - DEFAULT_BUTTON_HEIGHT - DEFAULT_PANEL_MARGIN
End Sub

Private Sub SetValues()
    With BaseConfig
        TxtOutputDirectory.Text = .OutputDirectory
        TxtVideoExtNames.Text = .VideoExtNames
        TxtCrfOfVideo.Text = CStr(.CrfOfVideo)
        TxtBitrateOfAudio.Text = CStr(.BitrateOfAudio)
        Call SetComboBoxListIndexByTitle(CmbPreset, .Preset)
        Call SetComboBoxListIndexByTitle(CmbTune, .Tune)
        Call SetComboBoxListIndexByTitle(CmbLevel, .level)
        Call SetComboBoxListIndexByTitle(CmbMaxHD, .MaxHD)
        Call SetComboBoxListIndexByTitle(CmbMaxFPS, .MaxFPS)
    End With
End Sub

Public Sub Browser(ByVal hWnd As Long)
    Dim FP As String
    
    FP = GetSelectedFolder(hWnd, "請指定輸出存放資料夾")
    If "" <> FP Then TxtOutputDirectory.Text = FP
End Sub

Public Sub Save()
    Call WriteStrToIniFile(CONFIG_INI, CONFIG_KEY_OUTPUT_DIRECTORY, TxtOutputDirectory.Text)
    Call WriteStrToIniFile(CONFIG_INI, CONFIG_KEY_VIDEO_EXT_NAMES, LCase$(TxtVideoExtNames.Text))
    Call WriteStrToIniFile(CONFIG_INI, CONFIG_KEY_CRF_OF_VIDEO, TxtCrfOfVideo.Text)
    Call WriteStrToIniFile(CONFIG_INI, CONFIG_KEY_BITRATE_OF_AUDIO, TxtBitrateOfAudio.Text)
    Call WriteStrToIniFile(CONFIG_INI, CONFIG_KEY_PRESET, GetComboBoxListTitle(CmbPreset))
    Call WriteStrToIniFile(CONFIG_INI, CONFIG_KEY_TUNE, GetComboBoxListTitle(CmbTune))
    Call WriteStrToIniFile(CONFIG_INI, CONFIG_KEY_LEVEL, GetComboBoxListTitle(CmbLevel))
    Call WriteStrToIniFile(CONFIG_INI, CONFIG_KEY_MAX_HD, GetComboBoxListTitle(CmbMaxHD))
    Call WriteStrToIniFile(CONFIG_INI, CONFIG_KEY_MAX_FPS, GetComboBoxListTitle(CmbMaxFPS))
    
    Call LoadBaseConfig
    Call SetValues
    
    Call MsgInfo("配置已生效！")
End Sub

Private Sub SetComboBoxListIndexByTitle(CMB As ComboBox, ByVal sTitle As String)
    Dim I As Long
    
    With CMB
        If "" <> sTitle Then
            For I = 0 To (.ListCount - 1)
                If sTitle = .List(I) Then
                    .ListIndex = I
                    Exit For
                End If
            Next
        Else
            .ListIndex = 0
        End If
    End With
End Sub

Private Function GetComboBoxListTitle(CMB As ComboBox) As String
    With CMB
        If .ListIndex > 0 Then
            GetComboBoxListTitle = .List(.ListIndex)
        Else
            GetComboBoxListTitle = ""
        End If
    End With
End Function
