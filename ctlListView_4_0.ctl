VERSION 5.00
Begin VB.UserControl ListView 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "ListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-Callback declarations for Paul Caton thunking magic----------------------------------------------
Private z_CbMem   As Long                                                       'Callback allocated memory address
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
'-------------------------------------------------------------------------------------------------

'-----------------------------------HookWindow-------------------------------------------------
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const GWL_WNDPROC = (-4)
Private Const WM_NCDESTROY = &H82

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202

Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208


Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Dim m_hwnd As Long, m_NewProc As Long, m_OldProc As Long


Public Event DblClick(ByVal Button As Integer)                                  '��� = 1 �Ҽ� = 2 �м� = 3
Public Event MouseDown(ByVal Button As Integer)                                 '��� = 1 �Ҽ� = 2 �м� = 3
Public Event MouseUP(ByVal Button As Integer)                                   '��� = 1 �Ҽ� = 2 �м� = 3





'-----------------------------------HookWindow-------------------------------------------------













Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, LpParam As Any) As Long
'Private Declare Function CreateWindow Lib "user32" Alias "CreateWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Private Declare Function SendLongMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CoInitialize Lib "ole32.dll" (ByVal pvReserved As Long) As Long
Private Declare Sub CoUninitialize Lib "ole32.dll" ()
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long




Const WC_LISTVIEWA = "SysListView32"

Private Const WS_CHILD = &H40000000
Private Const WS_VISIBLE = &H10000000

Const WM_SETFOCUS = &H7

Private Const LVM_FIRST = &H1000
Private Const LVM_INSERTCOLUMN = (LVM_FIRST + 27)
Const LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54)
Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
Const LVM_GETITEMSTATE = (LVM_FIRST + 44)
Const LVM_INSERTITEM = (LVM_FIRST + 7)
Const LVM_SETITEMTEXT = (LVM_FIRST + 46)
Const LVM_GETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 55)
Const LVM_GETITEM = (LVM_FIRST + 5)
Const LVM_GETITEMTEXT = (LVM_FIRST + 45)
Const LVM_SETITEM = (LVM_FIRST + 6)
Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Const LVM_SETCOLUMN = (LVM_FIRST + 26)
Const LVM_GETCOLUMNWIDTH = (LVM_FIRST + 29)
Const LVM_DELETEITEM = (LVM_FIRST + 8)
Const LVM_GETITEMCOUNT = (LVM_FIRST + 4)
Const LVM_DELETEALLITEMS = (LVM_FIRST + 9)
Const LVM_DELETECOLUMN = (LVM_FIRST + 28)
Const LVM_GETSELECTIONMARK = (LVM_FIRST + 66)
Const LVM_ENSUREVISIBLE = (LVM_FIRST + 19)
Const LVM_SETSELECTIONMARK = (LVM_FIRST + 67)
'Const LVM_SETITEMSTATE = (LVM_FIRST + 43)

Const LVS_EX_FULLROWSELECT = &H20
Const LVS_EX_GRIDLINES = &H1
Const LVS_EX_CHECKBOXES = &H4
Private Const LVS_REPORT = &H1

Const LVCFMT_LEFT = &H0

Const LVCF_FMT = &H1
Const LVCF_TEXT = &H4
Const LVCF_WIDTH = &H2

' LVITEM mask
Private Const LVIF_TEXT = &H1                                                   ' ������Ч
Private Const LVIF_IMAGE = &H2                                                  ' ͼƬ��Ч
Private Const LVIF_PARAM = &H4                                                  ' ������Ч
Private Const LVIF_STATE = &H8                                                  ' ״̬(����)��Ч
Private Const LVIF_INDENT = &H10                                                ' ͼ��������Ч
Private Const LVIF_NORECOMPUTE = &H800

' LVITEM state
Private Const LVIS_FOCUSED = &H1                                                '
Private Const LVIS_SELECTED = &H2
Private Const LVIS_CUT = &H4
Private Const LVIS_DROPHILITED = &H8
Private Const LVIS_ACTIVATING = &H20
Private Const LVIS_SELCHECK = &H2000
Private Const LVIS_OVERLAYMASK = &HF00
Private Const LVIS_STATEIMAGEMASK = &HF000

Dim hListView As Long

' ListView �¼���Ϣ
Const LVN_FIRST = (-100)
Const LVN_COLUMNCLICK = (LVN_FIRST - 8)

Public Event ColumnClick()

Private Type LV_COLUMN
    mask As Integer
    fmt As Long
    cx As Long
    pszText As String
    cchTextMax As Long
    iSubItem As Long
    iImage As Long
    iOrder As Long
    cxMin As Long
    cxDefault As Long
    cxIdeal As Long
End Type

Private Type LV_ITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    State As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type
Public RowCount As Long
Public ColumnCount As Long
Private nCheckState As Boolean
'Private Type LV_COLUMN
'    mask As Long
'    iItem As Long
'    iSubItem As Long
'    State As Long
'    stateMask As Long
'    pszText As String
'    cchTextMax As Long
'    iImage As Long
'    lParam As Long
'    iIndent As Long
'End Type


Private Sub ListView_SetExtendedListViewStyleEx(hWnd As Long, ByVal lStyle As Long) ', ByVal lStyleNot As Long)
    Dim lNewStyle As Long
    
    lNewStyle = SendMessageLong(hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
    lNewStyle = lNewStyle Or lStyle
    SendMessageLong hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, lNewStyle
End Sub

Public Function ListView_InsertColumn(hWnd As Long, iCol As Long, pszColumnText As String, Optional mnWidth As Long = 88)
    Dim pCol As LV_COLUMN
    With pCol
        .mask = LVCF_FMT Or LVCF_TEXT Or LVCF_WIDTH
        .fmt = LVCFMT_LEFT
        .cx = mnWidth
        .pszText = pszColumnText
    End With
    
    SendMessage hWnd, LVM_INSERTCOLUMN, iCol, pCol
    'Debug.Print SendMessage(hWnd, LVM_INSERTCOLUMN, iCol, pCol)
End Function

Public Function ListView_SetColumnWidth(hWnd As Long, iCol As Long, mnWidth As Long)

    SendMessage hListView, LVM_SETCOLUMNWIDTH, iCol, ByVal mnWidth
    'Debug.Print SendMessage(hListView, LVM_SETCOLUMNWIDTH, iCol, ByVal mnWidth)
    
End Function
Public Function ListView_GetColumnWidth(hWnd As Long, iCol As Long) As Long

    ListView_GetColumnWidth = SendMessage(hListView, LVM_GETCOLUMNWIDTH, iCol, 0)
    
End Function

Public Function ListView_DeleteColumn(hWnd As Long, iCol As Long) As Boolean

    ListView_DeleteColumn = SendMessage(hListView, LVM_DELETECOLUMN, iCol, 0)
    
End Function

Public Function ListView_GetSelectionMark(hWnd As Long) As Long
    
    ListView_GetSelectionMark = SendMessage(hListView, LVM_GETSELECTIONMARK, 0, 0)
    
End Function

' ��ʱ��ȡ����Column�ı�
'Public Function ListView_SetColumnText(hWnd As Long, iCol As Long, pszText As String)
' Dim pCol As LV_COLUMN
'    With pCol
'        .mask = LVIF_TEXT
'        .pszText = pszText
'        .cchTextMax = Len(pszText)
'    End With
'
'   Debug.Print SendMessage(hWnd, LVM_SETCOLUMN, iCol, pCol)
'End Function

Public Function ListView_InsertItem(hWnd As Long, nItem As Long, ItemText As String, Optional State As Long)
    Dim pItem As LV_ITEM
    With pItem
        .mask = LVIF_TEXT Or LVIF_STATE
        .iItem = nItem
        .pszText = ItemText
        .State = State
        .stateMask = LVIS_STATEIMAGEMASK
    End With
    
    SendMessage hWnd, LVM_INSERTITEM, 0, pItem
End Function

Public Function ListView_SetItemText(hWnd As Long, nItem As Long, iSubItem As Long, pszText As String)
 Dim pItem As LV_ITEM
    With pItem
        .mask = LVIF_TEXT Or LVIF_STATE
        .iItem = nItem
        .pszText = pszText
        .iSubItem = iSubItem
    End With
    
    SendMessage hWnd, LVM_SETITEMTEXT, nItem, pItem
End Function

Public Function ListView_GetItemText(hWnd As Long, nItem As Long, iSubItem As Long) As String
    Dim lpPitem As LV_ITEM
    Dim SubItemText As String
    
    SubItemText = String$(1024, 0)
    lpPitem.iSubItem = iSubItem
    lpPitem.cchTextMax = 1024
    lpPitem.pszText = SubItemText
    
    SendMessage hWnd, LVM_GETITEMTEXT, ByVal nItem, lpPitem
    ListView_GetItemText = Left$(lpPitem.pszText, InStr(lpPitem.pszText, vbNullChar) - 1)
End Function

Public Function ListView_GetItemCount(hWnd As Long) As Long

    ListView_GetItemCount = SendMessageLong(hWnd, LVM_GETITEMCOUNT, 0, 0)
    
End Function

Public Function ListView_DeleteItem(hWnd As Long, nItem As Long) As Boolean

    ListView_DeleteItem = SendMessageLong(hWnd, LVM_DELETEITEM, nItem, 0)
    
End Function

Public Function ListView_DeleteAllItems(hWnd As Long) As Boolean

    ListView_DeleteAllItems = SendMessageLong(hWnd, LVM_DELETEALLITEMS, 0, 0)
    
End Function

Public Function ListView_SetItem(hWnd As Long, nItem As Long, strItemText As String)
    Dim pItem As LV_ITEM
    With pItem
        .mask = LVIF_TEXT Or LVIF_STATE
        .iItem = nItem
        .pszText = strItemText
    End With
    
    SendMessage hWnd, LVM_SETITEM, 0, pItem
End Function

Public Function ListView_GetItem(hWnd As Long, nItem As Long) As String
    Dim lpItem As LV_ITEM, ItemText As String
    ItemText = String$(260, 0)
    
    With lpItem
        .mask = LVIF_TEXT
        .iItem = nItem
        .pszText = ItemText
        .cchTextMax = 256
        .iSubItem = 0
    End With
    
    SendMessage hWnd, LVM_GETITEM, 0, lpItem
    ListView_GetItem = Left$(lpItem.pszText, InStr(lpItem.pszText, vbNullChar) - 1)
End Function




Public Property Get CheckBox() As Boolean
    CheckBox = nCheckState
    If nCheckState = False Then
        ListView_SetExtendedListViewStyleEx hListView, LVS_EX_GRIDLINES Or LVS_EX_FULLROWSELECT
      Else
        ListView_SetExtendedListViewStyleEx hListView, LVS_EX_GRIDLINES Or LVS_EX_FULLROWSELECT Or LVS_EX_CHECKBOXES
    End If
End Property
Public Property Let CheckBox(ByVal CheckState As Boolean)
    nCheckState = CheckState
    PropertyChanged "CheckBox"
    If nCheckState = False Then
        ListView_SetExtendedListViewStyleEx hListView, LVS_EX_GRIDLINES Or LVS_EX_FULLROWSELECT
    Else
        ListView_SetExtendedListViewStyleEx hListView, LVS_EX_GRIDLINES Or LVS_EX_FULLROWSELECT Or LVS_EX_CHECKBOXES
    End If
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    nCheckState = PropBag.ReadProperty("CheckBox", False)
    If nCheckState = False Then
        ListView_SetExtendedListViewStyleEx hListView, LVS_EX_GRIDLINES Or LVS_EX_FULLROWSELECT
    Else
        ListView_SetExtendedListViewStyleEx hListView, LVS_EX_GRIDLINES Or LVS_EX_FULLROWSELECT Or LVS_EX_CHECKBOXES
    End If
    
End Sub

Private Sub UserControl_Show()
    Call InitCallBack
    If Ambient.UserMode = True Then
        ' ��ʱ��ֻ�������й�����
        Call Bind(hListView)
    End If
    
End Sub

Private Sub UserControl_Terminate()
    Call Unbind
    Call zTerminate
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("CheckBox", nCheckState, False)
    If nCheckState = False Then
        ListView_SetExtendedListViewStyleEx hListView, LVS_EX_GRIDLINES Or LVS_EX_FULLROWSELECT
    Else
        ListView_SetExtendedListViewStyleEx hListView, LVS_EX_GRIDLINES Or LVS_EX_FULLROWSELECT Or LVS_EX_CHECKBOXES
    End If
End Sub
Private Sub UserControl_Initialize()
    Call CoInitialize(0)
    Call InitCommonControls
    hListView = CreateWindowEx(0, _
    WC_LISTVIEWA, _
0, _
    LVS_REPORT Or WS_CHILD Or WS_VISIBLE, _
0, _
0, _
0, _
0, _
    UserControl.hWnd, _
0, _
0, _
0)
    If nCheckState = False Then
        ListView_SetExtendedListViewStyleEx hListView, LVS_EX_GRIDLINES Or LVS_EX_FULLROWSELECT
    Else
        ListView_SetExtendedListViewStyleEx hListView, LVS_EX_GRIDLINES Or LVS_EX_FULLROWSELECT Or LVS_EX_CHECKBOXES
    End If
    
End Sub
Private Sub UserControl_Resize()
    MoveWindow hListView, 0, 0, (Extender.Width) / 15, (Extender.Height) / 15, True
End Sub

Public Function GetRowText(ByVal Row As Long)
    GetRowText = GetItemText(Row, 0)
End Function
Public Sub AddListHead(ByVal HeadName As String, Optional ByVal clmWidth As Long = 100)
    ListView_InsertColumn hListView, ColumnCount, HeadName, clmWidth
    ColumnCount = ColumnCount + 1
End Sub
Public Function JoinColumns(ByVal Row As String, ByVal Columns As String, Optional SplitDelimiter As String = ",", Optional ByVal Delimiter As String = "----") As String
    Dim i As Long, tmpStr As String, SplitInfo() As String
    SplitInfo = Split(Columns, SplitDelimiter)
    For i = 1 To UBound(SplitInfo) + 1
        JoinColumns = JoinColumns & GetItemText(Row, i) & Delimiter
    Next i
    JoinColumns = Left(JoinColumns, Len(JoinColumns) - Len(Delimiter))
End Function
Public Function JoinAllColumn(ByVal Row As Long, Optional ByVal Delimiter As String = "----") As String
    Dim i As Long
    For i = 1 To ColumnCount
        JoinAllColumn = JoinAllColumn & GetItemText(Row, i) & Delimiter
    Next i
    JoinAllColumn = Left(JoinAllColumn, Len(JoinAllColumn) - Len(Delimiter))
End Function

Public Function SetItemText(ByVal Row As Long, ByVal Column As Long, ByVal sValue As String)
    ListView_SetItemText hListView, Row - 1, Column, sValue
End Function
Public Function GetItemText(ByVal Row As Long, ByVal Column As Long) As String
    GetItemText = ListView_GetItemText(hListView, Row - 1, Column)
End Function
Public Function GetSelectionRow() As Long
    Dim strItem As String
    Dim nSelectedItem As Long
    nSelectedItem = ListView_GetSelectionMark(hListView)
    GetSelectionRow = nSelectedItem + 1
End Function
Public Function AddList(ByVal HeadText As String)
            ListView_InsertItem hListView, RowCount, HeadText
            RowCount = RowCount + 1
End Function
Public Function EnsureVisible(ByVal Row As Long, ByVal fPartialOK As Boolean) As Boolean
    EnsureVisible = SendMessage(hListView, LVM_ENSUREVISIBLE, Row - 1, fPartialOK)
    SetRowSelected Row
End Function
Public Function DeleteColumn(iCol As Long) As Boolean
    DeleteColumn = SendMessage(hListView, LVM_DELETECOLUMN, iCol, 0)
End Function
Public Function DeleteRow(nItem As Long) As Boolean
    DeleteRow = SendMessageLong(hListView, LVM_DELETEITEM, nItem - 1, 0)
    RowCount = RowCount - 1
End Function
Public Function DeleteAllRow() As Boolean
    DeleteAllRow = SendMessageLong(hListView, LVM_DELETEALLITEMS, 0, 0)
    RowCount = 0
End Function
Public Function SetRowSelected(ByVal Row As Long) As Boolean
    Dim lpItem As LV_ITEM
    SetFocus hListView
    With lpItem
    .State = LVIS_FOCUSED Or LVIS_SELECTED
    .stateMask = LVIS_FOCUSED Or LVIS_SELECTED
    End With
    SetRowSelected = SendMessage(hListView, LVM_SETITEMSTATE, ByVal Row - 1, lpItem)
End Function
Public Function SetCheckState(ByVal Row As Long, ByVal CheckState As Boolean) As Long
    Dim lpItem As LV_ITEM
    With lpItem
        If CheckState = True Then .State = INDEXTOSTATEIMAGEMASK(2) Else .State = INDEXTOSTATEIMAGEMASK(1)
        .stateMask = LVIS_STATEIMAGEMASK
    End With
    
    SetCheckState = SendMessage(hListView, LVM_SETITEMSTATE, ByVal Row - 1, lpItem)
End Function
Public Function IsChecked(ByVal Row As Long) As Boolean
    '    Dim lpItem As LV_ITEM
    '    With lpItem
    '        .State = INDEXTOSTATEIMAGEMASK(2)
    '        .stateMask = LVIS_STATEIMAGEMASK
    '    End With
    '    IsChecked = SendMessage(hListView, LVM_GETITEMSTATE, ByVal Row - 1, lpItem)
    Dim nRet As Long
    Const MaskBit As Long = &H1000                                              '(2 ^ 12)
    
    nRet = SendMessage(hListView, LVM_GETITEMSTATE, Row - 1, ByVal LVIS_STATEIMAGEMASK)
    
    IsChecked = (((nRet \ MaskBit) - 1) <> 0)
End Function



Public Function INDEXTOSTATEIMAGEMASK(iIndex As Long) As Long
' #define INDEXTOSTATEIMAGEMASK(i) ((i) << 12)
  INDEXTOSTATEIMAGEMASK = iIndex * (2 ^ 12)
End Function


Public Sub SortList(ByVal SortColumn As Long, Optional AscendingSort As Boolean = True)
    Dim tmpStr As String, NumberIndex As String, Row As Long, Column As Long, OrgList() As String, SortList() As String, SplitInfo() As String, tmpArray() As String
    ReDim OrgList(1 To RowCount, 1 To ColumnCount)
    ReDim SortList(1 To RowCount, 1 To ColumnCount)
    ReDim tmpArray(1 To RowCount)
    For Row = 1 To RowCount
        For Column = 1 To ColumnCount
            OrgList(Row, Column) = GetItemText(Row, Column)
        Next Column
        tmpArray(Row) = OrgList(Row, SortColumn)
    Next Row
    tmpStr = Join(tmpArray, ",")
    tmpStr = SortNumberEx(tmpStr, , , NumberIndex, AscendingSort)
    SplitInfo = Split(NumberIndex, ",")
    For Row = 1 To RowCount
        For Column = 1 To ColumnCount
            SetItemText Row, Column, OrgList(SplitInfo(Row - 1), Column)
        Next Column
    Next Row
End Sub
Public Function SortNumberEx(ByVal nSortNumber As String, Optional Delimiter As String = ",", Optional RemoveDuplicate As Boolean = False, Optional ByRef GetIndexFromStr As String, Optional AscendingSort As Boolean = False) As String
    Dim tmpStr() As String, SplitInfo() As String, arrAscendingSort() As String, i As Long, MaxIndex As Long, MixIndex As Long, tMaxIndex As Long, tMixIndex As Long, MaxNum As String, MixNum As String, tStr As String, tStrB As String, tIndex As Long
    Dim gMaxIndex As String, gMixIndex As String
    
    SplitInfo = Split(nSortNumber, Delimiter)
    ReDim tmpStr(0 To UBound(SplitInfo))
    MaxIndex = UBound(SplitInfo)
    MixIndex = LBound(SplitInfo)
    tStr = ""
    For i = 0 To UBound(SplitInfo)
            tStr = GetStrFromArray(SplitInfo(), Delimiter)
            MaxNum = GetMaxNumber(tStr, Delimiter, tMaxIndex)
            MixNum = GetMixNumber(tStr, Delimiter, tMixIndex)
            If MaxNum <> "" Then
                tmpStr(MaxIndex) = MaxNum
                MaxIndex = MaxIndex - 1
                SplitInfo(tMaxIndex) = ""
                gMaxIndex = CStr(tMaxIndex + 1) & Delimiter & gMaxIndex
            End If
            If MixNum <> "" Then
                tmpStr(MixIndex) = MixNum
                MixIndex = MixIndex + 1
                SplitInfo(tMixIndex) = ""
                gMixIndex = gMixIndex & CStr(tMixIndex + 1) & Delimiter
            End If
    Next i
    tStr = GetStrFromArray(tmpStr, Delimiter)
    SortNumberEx = tStr
    GetIndexFromStr = gMixIndex & gMaxIndex
    GetIndexFromStr = Left(GetIndexFromStr, Len(GetIndexFromStr) - 1)
    If RemoveDuplicate = True Then
        SplitInfo = Split(tStr, Delimiter)
        tStr = ""
        tStrB = ""
        For tIndex = 0 To UBound(SplitInfo)
                If SplitInfo(tIndex) <> tStrB Then
                        tStrB = SplitInfo(tIndex)
                        tStr = tStr & tStrB & Delimiter
                End If
        Next tIndex
        SortNumberEx = Left(tStr, Len(tStr) - 1)
    End If
    If AscendingSort = True Then
        SplitInfo = Split(SortNumberEx, Delimiter)
        ReDim arrAscendingSort(0 To UBound(SplitInfo))
        For i = 0 To UBound(SplitInfo)
            arrAscendingSort(i + (UBound(SplitInfo) - i) - i) = SplitInfo(i)
        Next i
        SortNumberEx = GetStrFromArray(arrAscendingSort, Delimiter)
        
        SplitInfo = Split(GetIndexFromStr, Delimiter)
        ReDim arrAscendingSort(0 To UBound(SplitInfo))
        For i = 0 To UBound(SplitInfo)
            arrAscendingSort(i + (UBound(SplitInfo) - i) - i) = SplitInfo(i)
        Next i
        GetIndexFromStr = GetStrFromArray(arrAscendingSort, Delimiter)
        
    End If
End Function
Public Function GetStrFromArray(ByRef cArray() As String, Optional Delimiter As String = "") As String
    Dim i As Long
    For i = LBound(cArray) To UBound(cArray)
            GetStrFromArray = GetStrFromArray & cArray(i) & Delimiter
    Next i
    GetStrFromArray = Left(GetStrFromArray, Len(GetStrFromArray) - 1)
End Function
Public Function GetMaxNumber(ByVal tStr As String, Optional Delimiter As String = ",", Optional ByRef index As Long) As String
    Dim SplitInfo() As String, i As Long, tmpStr As String
    SplitInfo() = Split(tStr, Delimiter)
    tmpStr = SplitInfo(0)
    index = 0
    For i = LBound(SplitInfo) To UBound(SplitInfo)
            If SplitInfo(i) <> "" Then
                If Val(SplitInfo(i)) > Val(tmpStr) Then
                    tmpStr = SplitInfo(i)
                    index = i
                End If
            End If
    Next i
    GetMaxNumber = tmpStr
End Function
Public Function GetMixNumber(ByVal tStr As String, Optional Delimiter As String = ",", Optional ByRef index As Long) As String
    Dim SplitInfo() As String, i As Long, tmpStr As String
    SplitInfo() = Split(tStr, Delimiter)
    tmpStr = SplitInfo(0)
    index = 0
    For i = LBound(SplitInfo) To UBound(SplitInfo)
            If SplitInfo(i) <> "" Then
                If Val(SplitInfo(i)) < Val(tmpStr) Or SplitInfo(i) = "0" Then
                    tmpStr = SplitInfo(i)
                    index = i
                End If
            End If
    Next i
    GetMixNumber = tmpStr
End Function


Public Sub SplitListFromStr(What As String, Optional SplitChr As String = "----", Optional SplitLine As String = vbNewLine)
    Dim Info, Line, i As Long, j As Long
    On Error Resume Next
    If What <> "" Then
        Info = Split(What, SplitLine)
        For i = 0 To UBound(Info)
            If Info(i) <> "" Then
                Line = Split(Info(i), SplitChr)
                AddList RowCount + 1
                For j = 0 To UBound(Line)
                    SetItemText RowCount, (j + 1), Line(j)
                Next j
            End If
        Next i
    End If
End Sub
Public Sub SplitListFromFile(FilePath As String, Optional SplitChr As String = "----")
    Dim Info, Line, i As Long, j As Long
    Dim itmX As Object
    If Dir(FilePath) <> "" Then
        If ReadFile(FilePath) <> "" Then
            Info = Split(ReadFile(FilePath), vbNewLine)
            For i = 0 To UBound(Info)
                If Info(i) <> "" Then
                    Line = Split(Info(i), SplitChr)
                    AddList RowCount + 1
                    For j = 0 To UBound(Line)
                        SetItemText RowCount, (j + 1), Line(j)
                    Next j
                End If
            Next i
        End If
    End If
End Sub




Public Function ReadFile(FilePath As String) As String
    If Dir(FilePath) = "" Then Exit Function
    Dim FileBytes() As Byte
    Dim tmpStrA As Byte
    Dim i As Long
    On Error Resume Next
    ReDim FileBytes(FileLen(FilePath) - 1)
    Open FilePath For Binary Access Read As #1
    Get #1, , FileBytes
    Close #1
    If FileBytes(0) = &HFF And FileBytes(1) = &HFE Then
        ReadFile = StrConv(FileBytes, vbNarrow)
        ReadFile = Replace(ReadFile, "?", "")
    ElseIf FileBytes(0) = &HFE And FileBytes(1) = &HFF Then
        For i = 0 To UBound(FileBytes) Step 2
            tmpStrA = FileBytes(i)
            FileBytes(i) = FileBytes(i + 1)
            FileBytes(i + 1) = tmpStrA
        Next i
        ReadFile = StrConv(FileBytes, vbNarrow)
        ReadFile = Replace(ReadFile, "?", "")
    Else
        ReadFile = StrConv(FileBytes, vbUnicode)
    End If
End Function
Public Function WriteFile(FilePath As String, What As String) As String
    If Dir(FilePath) <> "" Then Kill FilePath
    Open FilePath For Binary Access Write As #1
    Put #1, , What
    Close #1
End Function






'============================================================================================
' /////////////////// �ص���������ʽת������ \\\\\\\\\\\\\\\\\\\
'============================================================================================

'*************************************************************************************************
'* cCallback - ��ͨ�õĻص�ģ��
'��*
'*ע�⣺
'*Ϊһ�࣬������û��ؼ��Ļص������ʹ�������ȫһ���ġ�
'*�ص������������ʹ�����Թ�ͬ��������������ʹ��롣
'�������������͵Ĵ�����һ���ļ��У���*..
'*ɾ���ظ��������ʹ��룬��Ctrl+ F5�ͻᷢ������Ϊ��
'*Ҫ�ر�ע���nOrdinal������zAddressOf
'��*
'��* Paul_Caton@hotmail.com
'����Ȩ��ѵģ�����Ϊ���ʵ�ʹ�ú����á�
'��*
'*1.0���ԭ........................................... .......................... 20060408
'* v1.1�����thunk��֧��........................................ ................ 20060409
'*1.2�������˿�ѡ��IDE����......................................... ........... 20060411
'* V1.3������һ����ѡ�Ļص�Ŀ�����....................................... .. 20060413
'*************************************************************************************************

'-�ص�����-----------------------------------------------------------------------------------
Private Function zb_AddressOf(ByVal nOrdinal As Long, _
    ByVal nParamCount As Long, _
    Optional ByVal nThunkNo As Long = 0, _
    Optional ByVal oCallback As Object = Nothing, _
    Optional ByVal bIdeSafety As Boolean = True) As Long                        '���ص�ַָ���Ļص���thunk
    '*************************************************************************************************
    '* nOrdinal - �ز���ŵģ������˽�з������1�����ڶ������2���ȵ�..
    '* nParamCount - ���ص��Ĳ���
    '* nThunkNo - ��ѡ��ͬʱ�������ص����ò�ͬ��thunk... ������MAX_THUNKS const�������Ҫͬʱʹ���������ϵ�thunk
    '* oCallback - ��ѡ�������ջص��Ķ������δ���壬�ص������͵������ʵ��
    '* bIdeSafety - ��ѡ������Ϊfalse������IDE������
    '*************************************************************************************************
    Const MAX_FUNKS   As Long = 2                                               'ͬʱ���е�thunk������������ζ��
    Const FUNK_LONGS  As Long = 22                                              '��ͷ����thunk
    Const FUNK_LEN    As Long = FUNK_LONGS * 4                                  'һ��thunk�е��ֽ�
    Const MEM_LEN     As Long = MAX_FUNKS * FUNK_LEN                            '��Ҫ���ڴ��ֽڵĻص���thunk
    Const PAGE_RWX    As Long = &H40&                                           '�����ִ���ڴ�
    Const MEM_COMMIT  As Long = &H1000&                                         '�ύ������ڴ�
    Dim nAddr       As Long
    Dim nOffset     As Long
    Dim z_Cb()      As Long                                                     'Callback thunk array
    
    If nThunkNo < 0 Or nThunkNo > (MAX_FUNKS - 1) Then
        MsgBox "nThunkNo doesn't exist.", vbCritical + vbApplicationModal, "Error in " & TypeName(Me) & ".cb_Callback"
        Exit Function
    End If
    
    If oCallback Is Nothing Then                                                '����û���û��ָ���Ļص�����
        Set oCallback = Me                                                      'Ȼ��������
    End If
    
    nAddr = zAddressOf(oCallback, nOrdinal)                                     '��ȡָ����ŵĻص���ַ
    If nAddr = 0 Then
        MsgBox "Callback address not found.", vbCritical + vbApplicationModal, "Error in " & TypeName(Me) & ".cb_Callback"
        Exit Function
    End If
    
    If z_CbMem = 0 Then                                                         '����ڴ�û�б�����
        ReDim z_Cb(0 To FUNK_LONGS - 1, 0 To MAX_FUNKS - 1) As Long             '����������������
        z_CbMem = VirtualAlloc(z_CbMem, MEM_LEN, MEM_COMMIT, PAGE_RWX)          '�����ִ���ڴ�
        
        If bIdeSafety Then                                                      '����û���ҪIDE����
            z_Cb(2, 0) = GetProcAddress(GetModuleHandleA("vba6"), "EbMode")     'EbMode��ַ
        End If
        z_Cb(3, 0) = GetProcAddress(GetModuleHandleA("kernel32"), "IsBadCodePtr")
        z_Cb(4, 0) = &HBB60E089
        z_Cb(6, 0) = &H73FFC589: z_Cb(7, 0) = &HC53FF04: z_Cb(8, 0) = &H7B831F75: z_Cb(9, 0) = &H20750008: z_Cb(10, 0) = &HE883E889: z_Cb(11, 0) = &HB9905004: z_Cb(13, 0) = &H74FF06E3: z_Cb(14, 0) = &HFAE2008D: z_Cb(15, 0) = &H53FF33FF: z_Cb(16, 0) = &HC2906104: z_Cb(18, 0) = &H830853FF: z_Cb(19, 0) = &HD87401F8: z_Cb(20, 0) = &H4589C031: z_Cb(21, 0) = &HEAEBFC
        
        For nOffset = 1 To MAX_FUNKS - 1                                        ' ÿ��thunk�ģ����ƵĻ�����thunk
            CopyMemory z_Cb(0, nOffset), z_Cb(0, 0), FUNK_LEN
        Next
        CopyMemory ByVal z_CbMem, z_Cb(0, 0), MEM_LEN                           '����thunk�����ִ���ڴ�
    End If
    
    nOffset = z_CbMem + nThunkNo * FUNK_LEN
    CopyMemory ByVal nOffset, ObjPtr(oCallback), 4&                             '���Ƶ����thunk��VMEM��objPtr
    CopyMemory ByVal nOffset + 4, nAddr, 4&                                     '�ص���ַ���Ƶ�VMEM
    CopyMemory ByVal nOffset + 20, nOffset, 4&                                  '���Ƶ�VMEM���thunk�Ŀ�ʼ
    CopyMemory ByVal nOffset + 48, nParamCount, 4&                              '�����Ƶ�VMEM�Ĳ�����
    CopyMemory ByVal nOffset + 68, nParamCount * 4, 4&                          '���Ƶ�VMEM�������ܳ���
    zb_AddressOf = nOffset + 16                                                 '������VMEM����������Ա���Ϊ
    
End Function

'���صĻص�����ĵ�ַָ����ŵķ�����1 =���һ��˽�з�����2=�����ڶ���˽�з�����
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
    Dim bSub  As Byte                                                           '�ļ�ֵ������ϣ���ҵ�һ���麯����ķ�������ָ����
    Dim bVal  As Byte
    Dim nAddr As Long                                                           '���麯����ĵ�ַ
    Dim i     As Long                                                           'ѭ������
    Dim j     As Long                                                           'ѭ������
    
    CopyMemory nAddr, ByVal ObjPtr(oCallback), 4                                '��ȡ�ص������ʵ���ĵ�ַ
    If Not zProbe(nAddr + &H1C, i, bSub) Then                                   'һ�෽����̽��
        If Not zProbe(nAddr + &H6F8, i, bSub) Then                              '����ʽ������̽��
            If Not zProbe(nAddr + &H7A4, i, bSub) Then                          '�����û����Ʒ�����̽��
                Exit Function                                                   '����...
            End If
        End If
    End If
    
    i = i + 4                                                                   '��ײ����һ��
    j = i + 1024                                                                '����һ��������޶ȣ�ɨ��256���麯�������Ŀ
    Do While i < j
        CopyMemory nAddr, ByVal i, 4                                            '��ȡ��ַ�洢�����vtable��
        
        If IsBadCodePtr(nAddr) Then                                             '����һ����Ч�Ĵ����ַ��
            CopyMemory zAddressOf, ByVal i - (nOrdinal * 4), 4                  '����ָ�����麯�������ڵ�ַ
            Exit Do                                                             '����ķ���ǩ�����˳�ѭ��
        End If
        
        CopyMemory bVal, ByVal nAddr, 1                                         '�õ����麯������ָ����ֽ�
        If bVal <> bSub Then                                                    '������ֽڲ�ƥ��Ԥ��ֵ...
            CopyMemory zAddressOf, ByVal i - (nOrdinal * 4), 4                  '����ָ�����麯�������ڵ�ַ
            Exit Do                                                             '����ķ���ǩ�����˳�ѭ��
        End If
        
        i = i + 4                                                               '��һ��vtable��
    Loop
End Function

'��ָ������ʼ��ַ���ڷ���ǩ����̽��
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
    Dim bVal    As Byte
    Dim nAddr   As Long
    Dim nLimit  As Long
    Dim nEntry  As Long
    
    nAddr = nStart                                                              '��ʼ��ַ
    nLimit = nAddr + 32                                                         '�˸���Ŀ��̽
    Do While nAddr < nLimit                                                     '��Ȼ���ǻ�û�дﵽ���ǵ�̽�����
        CopyMemory nEntry, ByVal nAddr, 4                                       'vtable��
        
        If nEntry <> 0 Then                                                     '���û��ʵ�ֽӿ�
            CopyMemory bVal, ByVal nEntry, 1                                    '�õ���ֵ��ָ���vtable��
            If bVal = &H33 Or bVal = &HE9 Then                                  '��鱾����P��ķ���ǩ��
                nMethod = nAddr                                                 '�洢vtable��
                bSub = bVal                                                     '�洢�ҵ��ķ���ǩ��
                zProbe = True                                                   '��ʾ�ɹ�
                Exit Function                                                   '����
            End If
        End If
        
        nAddr = nAddr + 4                                                       '��һ��vtable��
    Loop
End Function

Private Sub zTerminate()
    
    Const MEM_RELEASE As Long = &H8000&                                         '�ͷŷ�����ڴ��־
    If Not z_CbMem = 0 Then                                                     '����ڴ����
        VirtualFree z_CbMem, 0, MEM_RELEASE
        z_CbMem = 0                                                             '����;��ʾ�ڴ��ͷ�
    End If
End Sub

Private Function InitCallBack()
    m_NewProc = zb_AddressOf(1, 4)
End Function

Private Function Bind(ByVal hWnd As Long) As Boolean
    Call Unbind
    If IsWindow(hWnd) Then m_hwnd = hWnd
    m_OldProc = SetWindowLong(m_hwnd, GWL_WNDPROC, m_NewProc)
    Bind = CBool(m_OldProc)
End Function

Private Function Unbind() As Boolean
    If m_OldProc <> 0 Then Unbind = CBool(SetWindowLong(m_hwnd, GWL_WNDPROC, m_OldProc))
    m_OldProc = 0
End Function

Private Function WindowProcCallBack(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim bCallNext As Boolean, lReturn As Long
    bCallNext = True
    Select Case Msg
    Case WM_LBUTTONDBLCLK
        RaiseEvent DblClick(1)
    Case WM_LBUTTONDOWN
        RaiseEvent MouseDown(1)
    Case WM_LBUTTONUP
        RaiseEvent MouseUP(1)
        
    Case WM_RBUTTONDBLCLK
        RaiseEvent DblClick(2)
    Case WM_RBUTTONDOWN
        RaiseEvent MouseDown(2)
    Case WM_RBUTTONUP
        RaiseEvent MouseUP(2)
        
    Case WM_MBUTTONDBLCLK
        RaiseEvent DblClick(3)
        
    Case WM_MBUTTONDOWN
        RaiseEvent MouseDown(3)
    Case WM_MBUTTONUP
        RaiseEvent MouseUP(3)
        
        
    End Select
    
    
    
    
    
    'RaiseEvent WindowProc(Msg, wParam, lParam, bCallNext, lReturn)
    If bCallNext Then
        WindowProcCallBack = CallWindowProc(m_OldProc, hWnd, Msg, wParam, lParam)
    Else
        WindowProcCallBack = lReturn
    End If
    If hWnd = m_hwnd And Msg = WM_NCDESTROY Then Call Unbind
End Function












