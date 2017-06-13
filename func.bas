Attribute VB_Name = "func"
Option Explicit
    
Public Declare Function CreateHardLinkW Lib "Kernel32" _
    (ByVal lpFileName As Long, _
     ByVal lpExistingFileName As Long, _
     ByVal lpSecurityAttributes As Long) As Long
     
Public BaseConfig As BASE_CONFIG

Sub Main()
    Dim C As String
    
    Call InitCommonControls
    Call InitWindowsVersion
    Call LoadBaseConfig
    
    C = Trim$(Command)
    If "" = C Then
        ServerForm.Show
    Else
        Call DecodeCommand(C)
    End If
End Sub

Private Sub DecodeCommand(ByVal C As String)
    Dim A() As String
    Dim U As Long
    
    A = Split(C, CMD_SPLIT_CHAR)
    U = UBound(A)
    Select Case UCase$(A(CMD_NAME))
        Case CMD_GET_VIDEO_INFO
            Call CmdGetVideoInfo(A)
            
        'Case CMD_CREATE_STARTUP_SHORTCUT
        '    Call CmdCreateStartupShortcut
        
        'Case CMD_NOTIFY_ORDER
        '    Call CmdNotifyOrder
        
        'Case CMD_NOTIFY_CALL
        '    Call CmdNotifyCall
        
    End Select
End Sub

Private Sub CmdGetVideoInfo(A() As String)
    Dim FM As New cFFmpeg
    Dim hWnd As Long
    Dim uID As Long
    Dim FP As String
    Dim AvInfo(AV_UBOUND) As String
    Dim nStatus As Long

    FP = A(GVI_ARGS_FILE)
    hWnd = CLng(A(GVI_ARGS_HWND))
    uID = CLng(A(GVI_ARGS_UID))

    If IsFileExist(FP) Then
        If FM.GetVideoInfoTextByFFmpeg(FP, AvInfo) Then
            'If FM.GetVideoInfoTextByMPlayer(FP, AvInfo) Then
                If "" <> AvInfo(AV_VIDEO_CODEC) Then
                    If "" <> AvInfo(AV_AUDIO_CODEC) Then
                        WriteAnsiTextToFile GetInfoFile(uID), Join(AvInfo, vbTab)
                        nStatus = VIDEO_INFO_SUCCESS
                    Else
                        nStatus = VIDEO_INFO_NO_AUDIO
                    End If
                Else
                    nStatus = VIDEO_INFO_NO_VIDEO
                End If
            'Else
            '    nStatus = VIDEO_INFO_MPLAYER_FAIL
            'End If
        Else
            nStatus = VIDEO_INFO_FFMPEG_FAIL
        End If
    Else
        nStatus = VIDEO_INFO_NOT_EXIST
    End If
    SendMessageW hWnd, WM_VIDEO_INFO, nStatus, uID
    
    Set FM = Nothing
End Sub

Public Sub LoadBaseConfig()
    With BaseConfig
        .OutputDirectory = ReadStrFromIniFile(CONFIG_INI, CONFIG_KEY_OUTPUT_DIRECTORY)
        .VideoExtNames = ReadStrFromIniFile(CONFIG_INI, CONFIG_KEY_VIDEO_EXT_NAMES)
        .CrfOfVideo = CLng(SafeGetNumbersFromString(ReadStrFromIniFile(CONFIG_INI, CONFIG_KEY_CRF_OF_VIDEO)))
        .BitrateOfAudio = CLng(SafeGetNumbersFromString(ReadStrFromIniFile(CONFIG_INI, CONFIG_KEY_BITRATE_OF_AUDIO)))
        .Preset = ReadStrFromIniFile(CONFIG_INI, CONFIG_KEY_PRESET)
        .Tune = ReadStrFromIniFile(CONFIG_INI, CONFIG_KEY_TUNE)
        .Level = ReadStrFromIniFile(CONFIG_INI, CONFIG_KEY_LEVEL)
        .MaxHD = ReadStrFromIniFile(CONFIG_INI, CONFIG_KEY_MAX_HD)
        .MaxFPS = ReadStrFromIniFile(CONFIG_INI, CONFIG_KEY_MAX_FPS)
        
        .FFmpegExe = App.Path + "\ffmpeg.exe"
        '.MPlayerExe = App.Path + "\mplayer.exe"
        
        If "" = .OutputDirectory Then
            Call MsgError("尚未設定 ""輸出存放資料夾"" ！")
        Else
            If Not IsFolderExist(.OutputDirectory) Then
                Call MsgError("""輸出存放資料夾"" 不存在！")
            End If
        End If
        If "" = .VideoExtNames Then .VideoExtNames = "/mp4/m4v/h264/flv/f4v/webm/mov/hdmov/avi/wmv/mkv/bd/bdmv/bsf/rm/rmvb/mpg/mpeg/asf/ts/ogv/m1v/vob/m2v/m2p/m2t/m2ts/m4v/3g2/3gp/"
        If 0 = .CrfOfVideo Then .CrfOfVideo = 22
        If 0 = .BitrateOfAudio Then .BitrateOfAudio = 128
        
        If Not IsFileExist(.FFmpegExe) Then Call MsgError("請下載 ""ffmpeg.exe"" 與本程式放在同一資料夾！")
    End With
End Sub

Public Function GetInfoFile(ByVal uID As Long) As String
    GetInfoFile = App.Path + "\" + APPLICATION_ID + "." + CStr(uID) + ".tsv"
End Function

Public Function GetAddressInfo(retSegment As String, retIndex As Long) As Boolean
    Dim IpAdapterInfo As IP_ADAPTER_INFO
    Dim IpAdapterInfoLen As Long
    Dim lpIpAdapterInfo As Long
    Dim lpIpAddressListOfIpAdapterInfo As Long
    Dim IpAddrString As IP_ADDR_STRING
    Dim IpAddrStringLen As Long
    Dim lpIpAddrString As Long
    Dim nSize As Long
    Dim Buffer() As Byte
    Dim lpBuffer As Long
    Dim nAddress As Long
    Dim PhyAddr(1) As Long
    
    GetAddressInfo = False
    
    retSegment = ""
    retIndex = INVALID_HANDLE_VALUE
    
    nSize = 0
    If ERROR_BUFFER_OVERFLOW = GetAdaptersInfo(0, nSize) Then
        ReDim Buffer(nSize - 1)
        lpBuffer = VarPtr(Buffer(0))
        If ERROR_SUCCESS = GetAdaptersInfo(lpBuffer, nSize) Then
            IpAdapterInfoLen = Len(IpAdapterInfo)
            IpAddrStringLen = Len(IpAddrString)
            lpIpAdapterInfo = VarPtr(IpAdapterInfo)
            lpIpAddressListOfIpAdapterInfo = VarPtr(IpAdapterInfo.IpAddressList)
            lpIpAddrString = VarPtr(IpAddrString)
            
            CopyMemory lpIpAdapterInfo, lpBuffer, IpAdapterInfoLen
            Do
                If MIB_IF_TYPE_ETHERNET = IpAdapterInfo.type Then
                    CopyMemory VarPtr(PhyAddr(0)), VarPtr(IpAdapterInfo.Address(0)), IpAdapterInfo.AddressLength
                    If 0 <> PhyAddr(0) Then
                        If 0 <> PhyAddr(1) Then
                            CopyMemory lpIpAddrString, lpIpAddressListOfIpAdapterInfo, IpAddrStringLen
                            Do
                                nAddress = inet_addr_by_addr(IpAddrString.IpAddress(0))
                                If INADDR_ANY <> nAddress Then
                                    retSegment = ConvIpToStr(nAddress)
                                    retSegment = Left$(retSegment, InStrRev(retSegment, ".") - 1)
                                    
                                    nAddress = htonl(nAddress)
                                    retIndex = nAddress And 255
                                End If
                                
                                If IpAddrString.Next = 0 Then Exit Do
                                CopyMemory lpIpAddrString, IpAddrString.Next, IpAddrStringLen
                            Loop
                        End If
                    End If
                End If
                
                If IpAdapterInfo.Next = 0 Then Exit Do
                CopyMemory lpIpAdapterInfo, IpAdapterInfo.Next, IpAdapterInfoLen
            Loop
        End If
    End If
    
    GetAddressInfo = (INVALID_HANDLE_VALUE <> retIndex)
    
    Erase Buffer
End Function

Public Function GetFileLength(ByVal FP As String) As Long
    Dim hFile As Long
    Dim SizeH As Long
    
    GetFileLength = 0
    
    hFile = CreateFileW(StrPtr(FP), GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, 0)
    If INVALID_HANDLE_VALUE <> hFile Then
        GetFileLength = GetFileSize(hFile, SizeH)
        CloseHandle hFile
    End If
End Function

Public Function SafeGetNumbersFromString(ByVal S As String) As String
    Dim L As Long
    Dim I As Long
    Dim V As Long
    
    SafeGetNumbersFromString = ""
    If "" <> S Then
        L = Len(S)
        For I = 1 To L
            V = Asc(Mid$(S, I, 1))
            If V >= &H30 Then
                If V <= &H39 Then
                    SafeGetNumbersFromString = SafeGetNumbersFromString + ChrW$(V)
                End If
            End If
        Next
    End If
    
    If "" = SafeGetNumbersFromString Then SafeGetNumbersFromString = "0"
End Function

Public Function GetNowTimeStr() As String
    Dim ST As SYSTEMTIME
    
    Call GetLocalTime(ST)
    With ST
        GetNowTimeStr = ConvIntegerTo2Digi(.wHour) + ":" + ConvIntegerTo2Digi(.wMinute) + ":" + ConvIntegerTo2Digi(.wSecond)
    End With
End Function

Public Function ConvIntegerTo2Digi(ByVal V As Integer) As String
    ConvIntegerTo2Digi = Right$("0" + CStr(V), 2)
End Function

Public Function CreateHardLinkFile(ByVal srcFP As String, ByVal desFP As String) As Boolean
    DeleteFile desFP
    CreateHardLinkFile = (0 <> CreateHardLinkW(StrPtr(desFP), StrPtr(srcFP), 0))
End Function

Public Function GetFileSize64(ByVal FP As String) As Currency
    Dim hFile As Long
    
    GetFileSize64 = 0
    
    hFile = CreateFileW(StrPtr(FP), GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_FLAG_NO_BUFFERING, 0)
    If INVALID_HANDLE_VALUE <> hFile Then
        GetFileSizeEx hFile, GetFileSize64
        
        CloseHandle hFile
        
        GetFileSize64 = GetFileSize64 * 10000
    End If
End Function
