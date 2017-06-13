VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cListview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const SB_TOP = 6

Public hWnd As Long

Private Sub Class_Initialize()
    hWnd = 0
End Sub

Private Sub Class_Terminate()
    If hWnd <> 0 Then DestroyWindow hWnd
End Sub

Public Sub Create(ByVal hWndOfParent As Long, ByVal nX As Long, ByVal nY As Long, ByVal nWidth As Long, ByVal nHeight As Long)
    hWnd = CreateWindowExW( _
            WS_EX_CLIENTEDGE, _
            StrPtr(WC_LISTVIEW), _
            0, _
            WS_CHILD Or WS_VISIBLE Or LVS_REPORT Or LVS_SHOWSELALWAYS Or LVS_SHAREIMAGELISTS Or LVS_SINGLESEL, _
            nX, _
            nY, _
            nWidth, _
            nHeight, _
            hWndOfParent, _
            0, _
            App.hInstance, _
            0)
        
    Call SetFontFromhWnd(hWndOfParent)
    SendMessageW hWnd, LVM_SETTEXTBKCOLOR, 0, &HFFFFFFFF
    Call SetExStyle(LVS_EX_DOUBLEBUFFER Or LVS_EX_INFOTIP Or LVS_EX_FULLROWSELECT Or LVS_EX_GRIDLINES)
End Sub

Public Sub CreateMulti(ByVal hWndOfParent As Long, ByVal nX As Long, ByVal nY As Long, ByVal nWidth As Long, ByVal nHeight As Long)
    hWnd = CreateWindowExW( _
            WS_EX_CLIENTEDGE, _
            StrPtr(WC_LISTVIEW), _
            0, _
            WS_CHILD Or WS_VISIBLE Or LVS_REPORT Or LVS_SHOWSELALWAYS Or LVS_SHAREIMAGELISTS, _
            nX, _
            nY, _
            nWidth, _
            nHeight, _
            hWndOfParent, _
            0, _
            App.hInstance, _
            0)
        
    Call SetFontFromhWnd(hWndOfParent)
    SendMessageW hWnd, LVM_SETTEXTBKCOLOR, 0, &HFFFFFFFF
    Call SetExStyle(LVS_EX_DOUBLEBUFFER Or LVS_EX_INFOTIP Or LVS_EX_FULLROWSELECT Or LVS_EX_GRIDLINES)
End Sub

Public Sub CreateNoEdge(ByVal hWndOfParent As Long, ByVal nX As Long, ByVal nY As Long, ByVal nWidth As Long, ByVal nHeight As Long)
    hWnd = CreateWindowExW( _
            0, _
            StrPtr(WC_LISTVIEW), _
            0, _
            WS_CHILD Or WS_VISIBLE Or LVS_REPORT Or LVS_SHOWSELALWAYS Or LVS_SHAREIMAGELISTS Or LVS_SINGLESEL, _
            nX, _
            nY, _
            nWidth, _
            nHeight, _
            hWndOfParent, _
            0, _
            App.hInstance, _
            0)
        
    Call SetFontFromhWnd(hWndOfParent)
    SendMessageW hWnd, LVM_SETTEXTBKCOLOR, 0, &HFFFFFFFF
    Call SetExStyle(LVS_EX_DOUBLEBUFFER Or LVS_EX_INFOTIP Or LVS_EX_FULLROWSELECT Or LVS_EX_GRIDLINES)
End Sub

Public Sub SetExStyle(ByVal nExStyle As Long)
    SendMessageW hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, nExStyle Or SendMessageW(hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
End Sub

Public Sub SetFontFromhWnd(ByVal fromhWnd As Long)
    SendMessageW hWnd, WM_SETFONT, SendMessageW(fromhWnd, WM_GETFONT, 0, 0), 1
End Sub

Public Sub Move(ByVal nX As Long, ByVal nY As Long, ByVal nW As Long, ByVal nH As Long)
    MoveWindow hWnd, nX, nY, nW, nH, &HFFFFFFFF
End Sub

Public Sub Show()
    ShowWindow hWnd, SW_SHOW
    DoEvents
End Sub

Public Sub Hide()
    ShowWindow hWnd, SW_HIDE
End Sub

Public Sub SetFocus()
    SetFocusW hWnd
End Sub

Public Function GetItemCount() As Long
    GetItemCount = SendMessageW(hWnd, LVM_GETITEMCOUNT, 0, 0)
End Function

Public Function GetColumnCount() As Long
    GetColumnCount = SendMessageW(SendMessageW(hWnd, LVM_GETHEADER, 0, 0), HDM_GETITEMCOUNT, 0, 0)
End Function

Public Sub EnableGroupView()
    SendMessageW hWnd, LVM_ENABLEGROUPVIEW, &HFFFFFFFF, 0
End Sub

Public Sub RemoveAllGroups()
    SendMessageW hWnd, LVM_REMOVEALLGROUPS, 0, 0
    SendMessageW hWnd, LVM_ENABLEGROUPVIEW, 0, 0
End Sub

Public Sub DeleteAllItems()
    SendMessageW hWnd, LVM_DELETEALLITEMS, 0, 0
End Sub

Public Sub DeleteItem(ByVal ItemIndex As Long)
    SendMessageW hWnd, LVM_DELETEITEM, ItemIndex, 0
End Sub

Public Sub SetImageListOfReport(ByVal hImageList As Long)
    SendMessageW hWnd, LVM_SETIMAGELIST, LVSIL_SMALL, hImageList
End Sub

Public Sub SetImageListOfIcon(ByVal hImageList As Long)
    SendMessageW hWnd, LVM_SETIMAGELIST, LVSIL_NORMAL, hImageList
End Sub

Public Sub SetImageListOfState(ByVal hImageList As Long)
    SendMessageW hWnd, LVM_SETIMAGELIST, LVSIL_STATE, hImageList
End Sub

Public Sub InsertGroup(ByVal GroupName As String, ByVal GroupId As Long)
    Dim LVG As LVGROUP
    
    With LVG
        .cbSize = Len(LVG)
        .mask = LVGF_HEADER Or LVGF_GROUPID
        .pszHeader = StrPtr(GroupName)
        .iGroupId = GroupId
        
        If CURRENT_WINDOWS_VERSION >= WINDOWS_VERSION_VISTA Then
            .mask = .mask Or LVGF_STATE Or LVGF_ALIGN
            .state = LVGS_COLLAPSIBLE
            .uAlign = LVGA_HEADER_LEFT
        End If
    End With
    SendMessageW hWnd, LVM_INSERTGROUP, -1, VarPtr(LVG)
End Sub

Public Sub InsertColumn(ByVal ColumnName As String, ByVal ColumnIndex As Long, ByVal ColumnWidth As Long)
    Dim LvC As LVCOLUMN
    
    With LvC
        .mask = LVCF_FMT Or LVCF_WIDTH Or LVCF_TEXT
        .fmt = LVCFMT_LEFT
        .cx = ColumnWidth
        .pszText = StrPtr(ColumnName)
    End With
    SendMessageW hWnd, LVM_INSERTCOLUMN, ColumnIndex, VarPtr(LvC)
End Sub

Public Sub InsertColumnAlignRight(ByVal ColumnName As String, ByVal ColumnIndex As Long, ByVal ColumnWidth As Long)
    Dim LvC As LVCOLUMN
    
    With LvC
        .mask = LVCF_FMT Or LVCF_WIDTH Or LVCF_TEXT
        .fmt = LVCFMT_RIGHT
        .cx = ColumnWidth
        .pszText = StrPtr(ColumnName)
    End With
    SendMessageW hWnd, LVM_INSERTCOLUMN, ColumnIndex, VarPtr(LvC)
End Sub

Public Sub InsertColumnAlignCenter(ByVal ColumnName As String, ByVal ColumnIndex As Long, ByVal ColumnWidth As Long)
    Dim LvC As LVCOLUMN
    
    With LvC
        .mask = LVCF_FMT Or LVCF_WIDTH Or LVCF_TEXT
        .fmt = LVCFMT_CENTER
        .cx = ColumnWidth
        .pszText = StrPtr(ColumnName)
    End With
    SendMessageW hWnd, LVM_INSERTCOLUMN, ColumnIndex, VarPtr(LvC)
End Sub

Public Function InsertItem(ByVal ItemName As String, ByVal ItemIcon As Long, ByVal ItemGroupId As Long) As Long
    Dim LVI As LVITEM
    
    With LVI
        .mask = LVIF_GROUPID Or LVIF_IMAGE Or LVIF_TEXT
        .iItem = GetItemCount
        .iSubItem = 0
        .iGroupId = ItemGroupId
        .iImage = ItemIcon
        .pszText = StrPtr(ItemName)
    End With
    InsertItem = SendMessageW(hWnd, LVM_INSERTITEM, 0, VarPtr(LVI))
End Function

Public Function InsertItemByPos(ByVal nPos As Long, ByVal ItemName As String, ByVal ItemIcon As Long, ByVal ItemGroupId As Long) As Long
    Dim LVI As LVITEM
    
    With LVI
        .mask = LVIF_GROUPID Or LVIF_IMAGE Or LVIF_TEXT
        .iItem = nPos
        .iSubItem = 0
        .iGroupId = ItemGroupId
        .iImage = ItemIcon
        .pszText = StrPtr(ItemName)
    End With
    InsertItemByPos = SendMessageW(hWnd, LVM_INSERTITEM, 0, VarPtr(LVI))
End Function

Public Function InsertItem2(ByVal ItemName As String, ByVal ItemGroupId As Long) As Long
    Dim LVI As LVITEM
    
    With LVI
        .mask = LVIF_GROUPID Or LVIF_TEXT
        .iItem = GetItemCount
        .iSubItem = 0
        .iGroupId = ItemGroupId
        .pszText = StrPtr(ItemName)
    End With
    InsertItem2 = SendMessageW(hWnd, LVM_INSERTITEM, 0, VarPtr(LVI))
End Function

Public Sub ReListView()
    SendMessageW hWnd, LVM_ARRANGE, LVA_DEFAULT, 0
End Sub

Public Function GetItemIcon(ByVal ItemIndex As Long) As Long
    Dim LVI As LVITEM
    
    With LVI
        .mask = LVIF_IMAGE
        .iItem = ItemIndex
    End With
    
    SendMessageW hWnd, LVM_GETITEM, 0, VarPtr(LVI)
    GetItemIcon = LVI.iImage
End Function

Public Function GetItemGroupId(ByVal ItemIndex As Long) As Long
    Dim LVI As LVITEM
    
    With LVI
        .mask = LVIF_GROUPID
        .iItem = ItemIndex
    End With
    
    SendMessageW hWnd, LVM_GETITEM, 0, VarPtr(LVI)
    GetItemGroupId = LVI.iGroupId
End Function

Public Sub SetItemTextAndIcon(ByVal ItemIndex As Long, ByVal ItemName As String, ByVal ItemIcon As Long)
    Dim LVI As LVITEM
    
    With LVI
        .mask = LVIF_IMAGE Or LVIF_TEXT
        .iItem = ItemIndex
        .iImage = ItemIcon
        .pszText = StrPtr(ItemName)
    End With
    
    SendMessageW hWnd, LVM_SETITEM, 0, VarPtr(LVI)
End Sub

Public Sub SetItemIcon(ByVal ItemIndex As Long, ByVal ItemIcon As Long)
    Dim LVI As LVITEM
    
    With LVI
        .mask = LVIF_IMAGE
        .iItem = ItemIndex
        .iImage = ItemIcon
    End With
    
    SendMessageW hWnd, LVM_SETITEM, 0, VarPtr(LVI)
End Sub

Public Sub SetItemText(ByVal ItemIndex As Long, ByVal ItemCol As Long, ByVal ItemName As String)
    Dim LVI As LVITEM
    
    With LVI
        .mask = LVIF_TEXT
        .iItem = ItemIndex
        .iSubItem = ItemCol
        .pszText = StrPtr(ItemName)
    End With
    
    SendMessageW hWnd, LVM_SETITEM, 0, VarPtr(LVI)
End Sub

Public Sub SetLastInsertItemText(ByVal nColIndex As Long, ByVal ItemName As String)
    Dim LVI As LVITEM
    
    With LVI
        .mask = LVIF_TEXT
        .iItem = GetItemCount - 1
        .iSubItem = nColIndex
        .pszText = StrPtr(ItemName)
    End With
    
    SendMessageW hWnd, LVM_SETITEM, 0, VarPtr(LVI)
End Sub

Public Sub SetLastInsertItemGroupId(ByVal newGroupId As Long, ByVal defaultGroupId As Long)
    Dim LVG As LVGROUP
    Dim LVI As LVITEM
    
    With LVG
        .cbSize = Len(LVG)
        .mask = LVGF_GROUPID
    End With
    With LVI
        .mask = LVIF_GROUPID
        .iItem = GetItemCount - 1
        If SendMessageW(hWnd, LVM_GETGROUPINFO, newGroupId, VarPtr(LVG)) <> -1 Then
            .iGroupId = newGroupId
        Else
            .iGroupId = defaultGroupId
        End If
    End With
    SendMessageW hWnd, LVM_SETITEM, 0, VarPtr(LVI)
End Sub

Public Sub SetLastInsertItemOverlayImage(ByVal iOverlay As Long)
    Dim LVI As LVITEM

    With LVI
        .mask = LVIF_STATE
        .iSubItem = 0
        .stateMask = LVIS_OVERLAYMASK
        .state = iOverlay * 256
    End With
    SendMessageW hWnd, LVM_SETITEMSTATE, GetItemCount - 1, VarPtr(LVI)
End Sub

Public Sub SetItemOverlayImage(ByVal ItemIndex As Long, ByVal iOverlay As Long)
    Dim LVI As LVITEM

    With LVI
        .mask = LVIF_STATE
        .iSubItem = 0
        .stateMask = LVIS_OVERLAYMASK
        .state = iOverlay * 256
    End With
    SendMessageW hWnd, LVM_SETITEMSTATE, ItemIndex, VarPtr(LVI)
End Sub


Public Sub Cls()
    SendMessageW hWnd, LVM_DELETEALLITEMS, 0, 0
End Sub

Public Function GetSelectedIndex() As Long
    GetSelectedIndex = SendMessageW(hWnd, LVM_GETNEXTITEM, -1, LVNI_SELECTED)
End Function

Public Function GetItemText(ByVal ItemIndex As Long, ByVal nColIndex As Long) As String
    Dim LVI As LVITEM
    
    GetItemText = String$(2048, vbNullChar)
    With LVI
        .mask = LVIF_TEXT
        .iItem = ItemIndex
        .iSubItem = nColIndex
        .pszText = StrPtr(GetItemText)
        .cchTextMax = 2048
    End With
    
    SendMessageW hWnd, LVM_GETITEM, 0, VarPtr(LVI)
    GetItemText = StrCutNull(GetItemText)
End Function

Public Sub SetBackgroundColor(ByVal nColor As Long)
    SendMessageW hWnd, LVM_SETBKCOLOR, 0, nColor
End Sub

Public Sub SetTextColor(ByVal nColor As Long)
    SendMessageW hWnd, LVM_SETTEXTCOLOR, 0, nColor
End Sub

Public Sub SetColumnWidth(ByVal nColIndex As Long, ByVal nWidth As Long)
    SendMessageW hWnd, LVM_SETCOLUMNWIDTH, nColIndex, nWidth
End Sub

Public Sub SetSelectedIndex(ByVal nIndex As Long)
    Dim LI As LVITEM
    
    With LI
        .stateMask = LVIS_SELECTED Or LVIS_FOCUSED
        .state = LVIS_SELECTED Or LVIS_FOCUSED
    End With
    SendMessageW hWnd, LVM_SETITEMSTATE, nIndex, VarPtr(LI)
    SendMessageW hWnd, LVM_ENSUREVISIBLE, nIndex, 0
End Sub

Public Sub SetTileViewInfo(ByVal nWidth As Long, ByVal nHeight As Long, ByVal nLines As Long)
    Dim LTVI As LVTILEVIEWINFO

    With LTVI
        .cbSize = Len(LTVI)
        .dwMask = LVTVIM_TILESIZE Or LVTVIM_COLUMNS
        .dwFlags = LVTVIF_FIXEDSIZE
        .sizeWidth = nWidth
        .sizeHeight = nHeight
        .cLines = nLines
    End With
    SendMessageW hWnd, LVM_SETTILEVIEWINFO, 0, VarPtr(LTVI)
End Sub

Public Sub SetTileInfo(ByVal ItemIndex As Long, ByVal nColumns As Long, puC() As Long)
    Dim LVI As LVTILEINFO

    With LVI
        .cbSize = Len(LVI)
        .iItem = ItemIndex
        .cColumns = nColumns
        .puColumns = VarPtr(puC(0))
    End With
    SendMessageW hWnd, LVM_SETTILEINFO, 0, VarPtr(LVI)
End Sub

Public Sub SetViewToTile()
    SendMessageW hWnd, LVM_SETVIEW, LV_VIEW_TILE, 0
End Sub

Public Function GetSelectedIndexs(nIndexs() As Long, nIndexsCount As Long) As Boolean
    Dim Ret As Long
    
    GetSelectedIndexs = False
    nIndexsCount = 0
    ReDim nIndexs(GetItemCount)
    Ret = SendMessageW(hWnd, LVM_GETNEXTITEM, -1, LVNI_SELECTED)
    Do While (Ret <> -1)
        nIndexs(nIndexsCount) = Ret
        nIndexsCount = nIndexsCount + 1
        Ret = SendMessageW(hWnd, LVM_GETNEXTITEM, Ret, LVNI_SELECTED)
    Loop
    GetSelectedIndexs = (nIndexsCount > 0)
End Function

Public Sub SetGroupName(ByVal GroupName As String, ByVal GroupId As Long)
    Dim LVG As LVGROUP
    Dim tmpGroupId As Long
    
    tmpGroupId = 99
    With LVG
        .cbSize = Len(LVG)
        .mask = LVGF_HEADER Or LVGF_GROUPID
        .pszHeader = StrPtr(GroupName)
        .iGroupId = tmpGroupId
    End With
    SendMessageW hWnd, LVM_SETGROUPINFO, GroupId, VarPtr(LVG)
    
    With LVG
        .mask = LVGF_GROUPID
        .iGroupId = GroupId
    End With
    SendMessageW hWnd, LVM_SETGROUPINFO, tmpGroupId, VarPtr(LVG)
End Sub

Public Function IsItemChecked(ByVal ItemIndex As Long) As Boolean
    Dim nState As Long
    
    nState = SendMessageW(hWnd, LVM_GETITEMSTATE, ItemIndex, LVIS_STATEIMAGEMASK)
    nState = nState And &H1000
    IsItemChecked = (nState = 0)
End Function

Public Function IsItemSelected(ByVal ItemIndex As Long) As Boolean
    Dim nState As Long
    
    nState = SendMessageW(hWnd, LVM_GETITEMSTATE, ItemIndex, LVIS_SELECTED)
    IsItemSelected = (0 <> nState)
End Function

Public Function IsVerticalScrollBar() As Boolean
    Dim nStyle As Long
    
    nStyle = GetWindowLong(hWnd, GWL_STYLE)
    nStyle = nStyle And WS_VSCROLL
    
    IsVerticalScrollBar = (0 <> nStyle)
End Function

Public Function IsHorizontalScrollBar() As Boolean
    Dim nStyle As Long
    
    nStyle = GetWindowLong(hWnd, GWL_STYLE)
    nStyle = nStyle And WS_HSCROLL
    
    IsHorizontalScrollBar = (0 <> nStyle)
End Function

Public Function IsChanged(ByVal lParam As Long, newIndex As Long) As Boolean
    Dim NML As NMLISTVIEW

    IsChanged = False

    CopyMemory VarPtr(NML), lParam, Len(NML)
    With NML
        If -1 <> .iItem Then
            .uChanged = .uChanged And LVIF_STATE
            If 0 <> .uChanged Then
                .uNewState = .uNewState And LVIS_SELECTED
                If 0 <> .uNewState Then
                    newIndex = .iItem
                    IsChanged = True
                End If
            End If
        End If
    End With
End Function

Public Function IsClick(ByVal lParam As Long, newIndex As Long) As Boolean
    Dim NMIA As NMITEMACTIVATE
    
    IsClick = False
    
    CopyMemory VarPtr(NMIA), lParam, Len(NMIA)
    With NMIA
        If -1 <> .iItem Then
            newIndex = .iItem
            IsClick = True
        End If
    End With
End Function

Public Function IsDoubleClick(ByVal lParam As Long, newIndex As Long) As Boolean
    Dim NMIA As NMITEMACTIVATE
    
    IsDoubleClick = False
    
    CopyMemory VarPtr(NMIA), lParam, Len(NMIA)
    With NMIA
        If -1 <> .iItem Then
            newIndex = .iItem
            IsDoubleClick = True
        End If
    End With
End Function

Public Function IsRightClick(ByVal lParam As Long, newIndex As Long) As Boolean
    Dim NMIA As NMITEMACTIVATE
    
    IsRightClick = False
    
    CopyMemory VarPtr(NMIA), lParam, Len(NMIA)
    With NMIA
        If -1 <> .iItem Then
            newIndex = .iItem
            IsRightClick = True
        End If
    End With
End Function

Public Function SetItemChecked(ByVal ItemIndex As Long, ByVal IsChecked As Boolean)
    Dim LI As LVITEM
    
    With LI
        .mask = LVIF_STATE
        .stateMask = LVIS_STATEIMAGEMASK
        If IsChecked Then
            .state = &H2000
        Else
            .state = &H1000
        End If
    End With
    SendMessageW hWnd, LVM_SETITEMSTATE, ItemIndex, VarPtr(LI)
End Function

Public Function InsertItemByRow(ByVal nRow As Long, ByVal ItemName As String) As Long
    Dim LVI As LVITEM
    
    With LVI
        .mask = LVIF_TEXT
        .iItem = nRow
        .pszText = StrPtr(ItemName)
    End With
    InsertItemByRow = SendMessageW(hWnd, LVM_INSERTITEM, 0, VarPtr(LVI))
End Function

Public Function GetColumnName(ByVal nCol As Long) As String
    Dim LvC As LVCOLUMN
    Dim pszText As String
    
    pszText = String$(MAX_PATH, vbNullChar)
    With LvC
        .mask = LVCF_TEXT
        .cchTextMax = MAX_PATH
        .pszText = StrPtr(pszText)
    End With
    SendMessageW hWnd, LVM_GETCOLUMN, nCol, VarPtr(LvC)
    GetColumnName = StrCutNull(pszText)
End Function

Public Sub DeleteColumn(ByVal nCol As Long)
    SendMessageW hWnd, LVM_DELETECOLUMN, nCol, 0
End Sub

Public Sub DeleteAllColumns()
    Dim nCount As Long
    Dim I As Long
    
    nCount = GetColumnCount
    If nCount > 0 Then
        For I = (nCount - 1) To 0 Step -1
            Call DeleteColumn(I)
        Next
    End If
End Sub

Public Function FindItem(ByVal ItemName As String) As Long
    Dim LFI As LVFINDINFO
    
    With LFI
        .flags = LVFI_STRING
        .psz = StrPtr(ItemName)
    End With
    
    FindItem = SendMessageW(hWnd, LVM_FINDITEM, -1, VarPtr(LFI))
End Function

Public Sub SetIconSpacing(ByVal nX As Long, ByVal nY As Long)
    CopyMemory VarPtr(nX) + 2, VarPtr(nY), 2
    SendMessageW hWnd, LVM_SETICONSPACING, 0, nX
End Sub

Public Sub SetLastInsertItemSubitemIcon(ByVal ItemCol As Long, ByVal ItemIcon As Long)
    Call SetSubitemIcon(GetItemCount - 1, ItemCol, ItemIcon)
End Sub

Public Sub SetSubitemIcon(ByVal ItemIndex As Long, ByVal ItemCol As Long, ByVal ItemIcon As Long)
    Dim LVI As LVITEM
    
    With LVI
        .mask = LVIF_IMAGE
        .iItem = ItemIndex
        .iSubItem = ItemCol
        .iImage = ItemIcon
    End With
    
    SendMessageW hWnd, LVM_SETITEM, 0, VarPtr(LVI)
End Sub

Public Sub SetGroupTextColor(ByVal nColor As Long)
    Dim LGM As LVGROUPMETRICS
    
    With LGM
        .cbSize = Len(LGM)
        .mask = LVGMF_TEXTCOLOR
        .crHeader = nColor
    End With
    SendMessageW hWnd, LVM_SETGROUPMETRICS, 0, VarPtr(LGM)
End Sub

Private Sub ChangeHeaderStyle()
    Dim hHeader As Long
    Dim nStyle As Long
    
    hHeader = SendMessageW(hWnd, LVM_GETHEADER, 0, 0)
    nStyle = GetWindowLong(hHeader, GWL_STYLE)
    nStyle = nStyle Xor HDS_BUTTONS
    SetWindowLong hHeader, GWL_STYLE, nStyle
End Sub

Public Sub ScrollToItem(ByVal ItemIndex As Long)
    Dim pt As POINTAPI
    
    SendMessageW hWnd, LVM_GETITEMPOSITION, ItemIndex, VarPtr(pt)
    SendMessageW hWnd, WM_VSCROLL, SB_TOP, 0
    SendMessageW hWnd, LVM_SCROLL, 0, pt.nY
End Sub

Public Function ExportToCsvFile(ByVal FP As String) As Boolean
    Dim hFile As Long
    Dim Rows As Long
    Dim Cols As Long
    Dim F() As String
    Dim I As Long
    Dim A() As Byte
    Dim J As Long
    Dim Ret As Long
    
    ExportToCsvFile = False
    
    DeleteFileW StrPtr(FP)
    hFile = CreateFileW(StrPtr(FP), GENERIC_WRITE, FILE_SHARE_READ, 0, CREATE_NEW, FILE_FLAG_SEQUENTIAL_SCAN, 0)
    If hFile <> INVALID_HANDLE_VALUE Then
        Cols = GetColumnCount
        ReDim F(Cols - 1)
        For I = 0 To (Cols - 1)
            F(I) = GetColumnName(I)
        Next
        A = StrConv(Join(F, vbTab) + vbCrLf, vbFromUnicode)
        WriteFile hFile, VarPtr(A(0)), UBound(A) + 1, Ret, 0
        
        Rows = GetItemCount
        If Rows > 0 Then
            For I = 0 To (Rows - 1)
                ReDim F(Cols - 1)
                For J = 0 To (Cols - 1)
                    F(J) = GetItemText(I, J)
                Next
                A = StrConv(Join(F, vbTab) + vbCrLf, vbFromUnicode)
                WriteFile hFile, VarPtr(A(0)), UBound(A) + 1, Ret, 0
            Next
        End If
        
        CloseHandle hFile
        ExportToCsvFile = True
    End If
    
    Erase F
    Erase A
End Function

Public Function GetRowIndexsByGroupId(ByVal GroupId As Long, RowIndexs() As Long, RowIndexCount As Long) As Boolean
    Dim C As Long
    Dim I As Long
    
    RowIndexCount = 0
    C = GetItemCount
    If C > 0 Then
        ReDim RowIndexs(C - 1)
        For I = 0 To (C - 1)
            If GetItemGroupId(I) = GroupId Then
                RowIndexs(RowIndexCount) = I
                RowIndexCount = RowIndexCount + 1
            End If
        Next
    End If
    GetRowIndexsByGroupId = (RowIndexCount > 0)
End Function
