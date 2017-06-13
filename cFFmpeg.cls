VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFFmpeg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateHardLinkW Lib "Kernel32" _
    (ByVal lpFileName As Long, _
     ByVal lpExistingFileName As Long, _
     ByVal lpSecurityAttributes As Long) As Long
     
Private Declare Function CreatePipe Lib "Kernel32" _
    (hReadPipe As Long, _
     hWritePipe As Long, _
     ByVal lpPipeAttributes As Long, _
     ByVal nSize As Long) As Long
     
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
     
Private Const CREATE_NO_WINDOW As Long = &H8000000

Public Sub GetConvertArguments(ByVal srcFP As String, ByVal desFP As String, AvInfo() As String, convArgs As String)
    Dim videoWidth As Long
    Dim videoHeight As Long
    Dim videoFPS As Single
    Dim videoRate As Long
    Dim IsNeedRS As Boolean
    Dim IsNeedRF As Boolean
    Dim newWidth As Long
    Dim newHeight As Long
    Dim newFPS As Single

    videoWidth = CLng(AvInfo(AV_VIDEO_WIDTH))
    videoHeight = CLng(AvInfo(AV_VIDEO_HEIGHT))
    videoFPS = CSng(AvInfo(AV_VIDEO_FPS))
    videoRate = CLng(AvInfo(AV_VIDEO_BITRATE))
    IsNeedRS = IsNeedRescale(videoWidth, videoHeight, newWidth, newHeight)
    IsNeedRF = IsNeedRefps(videoFPS, newFPS)

    convArgs = "-y -i """ + srcFP + """ -codec:v"
    If "h264" = AvInfo(AV_VIDEO_CODEC) Then
        If IsNeedRS Then
            convArgs = convArgs + " libx264"
            If IsNeedRF Then
                convArgs = convArgs + " -r " + CStr(newFPS)
            End If
            convArgs = convArgs + " -vf scale=" + CStr(newWidth) + ":" + CStr(newHeight)
            convArgs = convArgs + " -crf " + CStr(BaseConfig.CrfOfVideo) + ReturnMoreOptions
        Else
            If IsNeedRF Then
                convArgs = convArgs + " libx264" + " -r " + CStr(newFPS)
                convArgs = convArgs + " -crf " + CStr(BaseConfig.CrfOfVideo) + ReturnMoreOptions
            Else
                convArgs = convArgs + " copy"
            End If
        End If
    Else
        convArgs = convArgs + " libx264"
        If IsNeedRF Then
            convArgs = convArgs + " -r " + CStr(newFPS)
        End If
        If IsNeedRS Then
            convArgs = convArgs + " -vf scale=" + CStr(newWidth) + ":" + CStr(newHeight)
        End If
        convArgs = convArgs + " -crf " + CStr(BaseConfig.CrfOfVideo) + ReturnMoreOptions
    End If

    convArgs = convArgs + " -codec:a"
    If "aac" = AvInfo(AV_AUDIO_CODEC) Then
        convArgs = convArgs + " copy"
    Else
        convArgs = convArgs + " libvo_aacenc -b:a " + CStr(BaseConfig.BitrateOfAudio) + "k"
    End If
    If "stereo" <> AvInfo(AV_AUDIO_CHANNEL) Then
        convArgs = convArgs + " -ac 2"
    End If
    convArgs = convArgs + " -map_metadata -1 -movflags +faststart -f mp4 """ + desFP + """"
End Sub

Private Function ReturnMoreOptions() As String
    ReturnMoreOptions = " -pix_fmt yuv420p" + ReturnLevel()
    If "" <> BaseConfig.Preset Then ReturnMoreOptions = ReturnMoreOptions + " -preset " + BaseConfig.Preset
    If "" <> BaseConfig.Tune Then ReturnMoreOptions = ReturnMoreOptions + " -tune " + BaseConfig.Tune
End Function

Private Function ReturnLevel() As String
    Dim S As String
    
    ReturnLevel = ""
    If "" <> BaseConfig.level Then
        S = LCase$(BaseConfig.level)
        S = Replace(S, " ", " -level ")
        ReturnLevel = " -profile:v " + S
    End If
End Function

Private Function IsNeedRefps(ByVal srcFPS As Single, newFPS As Single) As Boolean
    IsNeedRefps = False
    
    Select Case BaseConfig.MaxFPS
        Case MAX_FPS_OF_3000
            newFPS = srcFPS
            Do While newFPS > 30
                newFPS = newFPS / 2
                IsNeedRefps = True
            Loop
        
    End Select
End Function

Private Function IsNeedRescale(ByVal srcWidth As Long, ByVal srcHeight As Long, newWidth As Long, newHeight As Long) As Boolean
    Dim V As Long
    
    IsNeedRescale = False
    newWidth = 0
    newHeight = 0
    
    Select Case BaseConfig.MaxHD
        Case MAX_HD_OF_360P
            If srcWidth > 640 Then newWidth = 640
            
        Case MAX_HD_OF_480P
            If srcWidth > 720 Then newWidth = 720
               
        Case MAX_HD_OF_540P
            If srcWidth > 960 Then newWidth = 960
            
        Case MAX_HD_OF_720P
            If srcWidth > 1280 Then newWidth = 1280
            
        Case MAX_HD_OF_1080P
            If srcWidth > 1920 Then newWidth = 1920
            
    End Select
    
    If 0 <> newWidth Then
        newHeight = newWidth * srcHeight \ srcWidth
        
        V = newHeight And 1
        If 0 <> V Then newHeight = newHeight + 1
        
        IsNeedRescale = True
    End If
End Function

Private Function CreateConsolePipe(ByVal CmdLine As String, rPipe As Long) As Boolean
    Dim sa As SECURITY_ATTRIBUTES
    Dim wPipe As Long
    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION
    
    CreateConsolePipe = False
    
    With sa
        .nLength = Len(sa)
        .bInheritHandle = &HFFFFFFFF
    End With
    CreatePipe rPipe, wPipe, VarPtr(sa), 0
    
    si.cb = Len(si)
    si.dwFlags = STARTF_USESTDHANDLES
    si.hStdInput = 0
    si.hStdOutput = wPipe
    si.hStdError = wPipe
    
    If 0 <> CreateProcessW(0, StrPtr(CmdLine), 0, 0, &HFFFFFFFF, CREATE_NO_WINDOW, 0, 0, si, pi) Then
        CloseHandle wPipe
        CreateConsolePipe = True
    End If
End Function

Private Function ReadPipeText(ByVal rPipe As Long) As String
    Dim Bin() As Byte
    Dim lpBin As Long
    Dim Ret As Long
    
    ReDim Bin(16383)
    lpBin = VarPtr(Bin(0))
    
    Do While (0 <> ReadFile(rPipe, lpBin, 4096, Ret, 0))
           lpBin = lpBin + Ret
    Loop
    
    ReadPipeText = StrConv(Bin, vbUnicode)
    ReadPipeText = StrCutNull(ReadPipeText)
    
    CloseHandle rPipe
    
    Erase Bin
End Function

'Public Function GetVideoInfoTextByMPlayer(ByVal FP As String, AvInfo() As String) As Boolean
'    Dim rPipe As Long
'    Dim T As String
'
'    GetVideoInfoTextByMPlayer = False
'
'    If CreateConsolePipe("""" + BaseConfig.MPlayerExe + """ -identify -frames 0 -vo null -ao null """ + FP + """", rPipe) Then
'        T = ReadPipeText(rPipe)
'        If "" <> T Then
'            Call GetVideoAndAudioInfoByMPlayer(T, AvInfo)
'            GetVideoInfoTextByMPlayer = True
'        End If
'    End If
'End Function

Public Function GetVideoInfoTextByFFmpeg(ByVal FP As String, AvInfo() As String) As Boolean
    Dim rPipe As Long
    Dim T As String
    
    GetVideoInfoTextByFFmpeg = False
    
    If CreateConsolePipe("""" + BaseConfig.FFmpegExe + """ -i """ + FP + """", rPipe) Then
        T = ReadPipeText(rPipe)
        If "" <> T Then
            Call GetVideoAndAudioInfoByFFmpeg(T, AvInfo)
            GetVideoInfoTextByFFmpeg = True
        End If
    End If
End Function

'Private Sub GetVideoAndAudioInfoByMPlayer(fromT As String, AvInfo() As String)
'    Dim L() As String
'    Dim U As Long
'    Dim I As Long
'    Dim Cols() As String
'    Dim S As Single
'
'    L = Split(fromT, vbCrLf)
'    U = UBound(L)
'    For I = 0 To U
'        If InStr(L(I), "ID_VIDEO_FPS") > 0 Then
'            Cols = Split(L(I), "=")
'            S = CSng(Cols(1))
'            If S <= 120 Then
'                If S <> CSng(AvInfo(AV_VIDEO_FPS)) Then
'                    AvInfo(AV_VIDEO_FPS) = CStr(S)
'                End If
'            End If
'
'            Exit For
'        End If
'    Next
'
'    Erase Cols
'    Erase L
'End Sub

Private Sub GetVideoAndAudioInfoByFFmpeg(fromT As String, AvInfo() As String)
    Dim L() As String
    Dim U As Long
    Dim I As Long
    Dim P As Long
    Dim Cols() As String
    Dim F() As String
    Dim R() As String
    
    L = Split(fromT, vbCrLf)
    U = UBound(L)
    For I = 0 To U
        P = InStr(L(I), "Duration: ")
        If P > 0 Then
            P = P + 10
            AvInfo(AV_DURATION) = Mid$(L(I), P, InStr(P, L(I), ",") - P)
            If InStr(AvInfo(AV_DURATION), ".") > 0 Then
                AvInfo(AV_DURATION) = Left$(AvInfo(AV_DURATION), InStrRev(AvInfo(AV_DURATION), ".") - 1)
            End If
        End If
        If InStr(L(I), "Stream #0") > 0 Then
            Cols = Split(L(I), ": ")
        
            If Cols(1) = "Video" Then
                F = Split(RemoveFieldDescription(Cols(2)), ", ")
                R = Split(GetFiledData(F(2)), "x")
                AvInfo(AV_VIDEO_CODEC) = GetFiledData(F(0))
                AvInfo(AV_VIDEO_WIDTH) = R(0)
                AvInfo(AV_VIDEO_HEIGHT) = R(1)
                AvInfo(AV_VIDEO_BITRATE) = GetFieldDataFromKey(F, "kb/s")
                AvInfo(AV_VIDEO_FPS) = GetFieldDataFromKey(F, "tbr")
            End If
            If Cols(1) = "Audio" Then
                F = Split(RemoveFieldDescription(Cols(2)), ", ")
                AvInfo(AV_AUDIO_CODEC) = GetFiledData(F(0))
                AvInfo(AV_AUDIO_CHANNEL) = GetFiledData(F(2))
                AvInfo(AV_AUDIO_HZ) = GetFiledData(F(1))
                AvInfo(AV_AUDIO_BITRATE) = GetFieldDataFromKey(F, "kb/s")
            End If
        End If
        
        If "" <> AvInfo(AV_VIDEO_CODEC) Then
            If "" <> AvInfo(AV_AUDIO_CODEC) Then
                If "" <> AvInfo(AV_DURATION) Then
                    Exit For
                End If
            End If
        End If
    Next
    
    Erase Cols
    Erase F
    Erase R
End Sub

Private Function GetFieldDataFromKey(F() As String, ByVal sKey As String) As String
    Dim U As Long
    Dim I As Long
    
    GetFieldDataFromKey = "0"
    
    U = UBound(F)
    For I = 3 To U
        If InStr(F(I), sKey) > 0 Then
            GetFieldDataFromKey = GetFiledData(F(I))
            Exit For
        End If
    Next
End Function

Private Function GetFiledData(S As String) As String
    Dim B() As Byte
    Dim U As Long
    Dim I As Long
    
    GetFiledData = ""
    
    B = StrConv(S, vbFromUnicode)
    U = UBound(B)
    For I = 0 To U
        If &H20 = B(I) Then Exit For
        GetFiledData = GetFiledData + Chr$(B(I))
    Next
    
    Erase B
End Function

Private Function RemoveFieldDescription(S As String) As String
    Dim B() As Byte
    Dim U As Long
    Dim I As Long
    
    RemoveFieldDescription = ""
    
    B = StrConv(S, vbFromUnicode)
    U = UBound(B)
    For I = 0 To U
        Select Case B(I)
            'Case &H20
            
            Case &H2C ',
                RemoveFieldDescription = RemoveFieldDescription + ", "
            
            Case &H28 '( ) &H29
                Do While (&H29 <> B(I))
                    I = I + 1
                Loop
            
            Case &H5B '[ ] &H5d
                Do While (&H5D <> B(I))
                    I = I + 1
                Loop
            
            Case Else
                RemoveFieldDescription = RemoveFieldDescription + Chr$(B(I))

        End Select
    Next
    RemoveFieldDescription = Replace(RemoveFieldDescription, ",  ", ", ")

    Erase B
End Function
