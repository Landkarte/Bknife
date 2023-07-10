VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MGSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Download by http://www.codesc.net
Private Const GWL_WNDPROC    As Long = (-4)
Private Const WM_DESTROY     As Long = &H2

Public Enum EErrorWindowProc
    eeBaseWindowProc = 13080
    eeCantSubclass
    eeAlreadyAttached
    eeInvalidWindow
    eeNoExternalWindow
End Enum

Private Declare Function EbMode_vb5 Lib "vba5" Alias "EbMode" () As Long

Private Declare Function EbMode_vb6 Lib "vba6" Alias "EbMode" () As Long

Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, _
                                                                ByVal lpString As String) As Long

Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, _
                                                                ByVal lpString As String, _
                                                                ByVal hData As Long) As Long

Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, _
                                                                      ByVal lpString As String) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                                                              ByVal hwnd As Long, _
                                                                              ByVal Msg As Long, _
                                                                              ByVal wParam As Long, _
                                                                              ByVal lParam As Long) As Long

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, _
                                                                lpdwProcessId As Long) As Long

Private Declare Function GetCurrentProcessId Lib "KERNEL32" () As Long

Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (lpvDest As Any, _
                                                                     lpvSource As Any, _
                                                                     ByVal cbCopy As Long)

Private Declare Function GlobalAlloc Lib "KERNEL32" (ByVal wFlags As Long, _
                                                     ByVal dwBytes As Long) As Long

Private Declare Function GlobalFree Lib "KERNEL32" (ByVal hMem As Long) As Long

Private Declare Function GetProcAddress Lib "KERNEL32" (ByVal hModule As Long, _
                                                        ByVal lpProcName As String) As Long

Private Declare Function GetModuleHandle Lib "KERNEL32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private m_lAdd        As Long
Private m_lClassAddr  As Long
Private m_lCurrMsg    As Long


Private Property Get p_MessageClass(ByVal lHwnd As Long, _
                                    ByVal lMsg As Long, _
                                    ByVal lIndex As Long) As Long

Dim sName As String

    sName = CStr(lHwnd) + Chr$(35) + CStr(lMsg) + Chr$(35) + CStr(lIndex)
    p_MessageClass = GetProp(lHwnd, sName)

End Property

Private Property Let p_MessageClass(ByVal lHwnd As Long, _
                                    ByVal lMsg As Long, _
                                    ByVal lIndex As Long, _
                                    ByVal lClass As Long)

Dim sName As String

    sName = CStr(lHwnd) + Chr$(35) + CStr(lMsg) + Chr$(35) + CStr(lIndex)
    m_lAdd = SetProp(lHwnd, sName, lClass)
    If lClass = 0 Then RemoveProp lHwnd, sName

End Property

Private Property Get p_MessageCount(ByVal lHwnd As Long) As Long

Dim sName As String

    sName = Chr$(67) + CStr(lHwnd)
    p_MessageCount = GetProp(lHwnd, sName)

End Property

Private Property Let p_MessageCount(ByVal lHwnd As Long, _
                                    ByVal lCount As Long)

Dim sName As String

    m_lAdd = 1
    sName = Chr$(67) + CStr(lHwnd)
    m_lAdd = SetProp(lHwnd, sName, lCount)
    If lCount = 0 Then RemoveProp lHwnd, sName

End Property

Public Property Get p_OldWindowProc(ByVal lHwnd As Long) As Long

Dim sName As String

    sName = lHwnd
    p_OldWindowProc = GetProp(lHwnd, sName)

End Property

Private Property Let p_OldWindowProc(ByVal lHwnd As Long, _
                                     ByVal lPtr As Long)

Dim sName As String

    m_lAdd = 1
    sName = lHwnd
    m_lAdd = SetProp(lHwnd, sName, lPtr)
    If lPtr = 0 Then RemoveProp lHwnd, sName

End Property

Private Property Get p_MessageClassCount(ByVal lHwnd As Long, _
                                         ByVal lMsg As Long) As Long

Dim sName As String

    sName = CStr(lHwnd) + Chr$(35) + CStr(lMsg) + Chr$(67)
    p_MessageClassCount = GetProp(lHwnd, sName)

End Property

Private Property Let p_MessageClassCount(ByVal lHwnd As Long, _
                                         ByVal lMsg As Long, _
                                         ByVal lCount As Long)

Dim sName As String

    sName = CStr(lHwnd) + Chr$(35) + CStr(lMsg) + Chr$(67)
    m_lAdd = SetProp(lHwnd, sName, lCount)
    If lCount = 0 Then RemoveProp lHwnd, sName

End Property

Public Property Get p_CurrentMessage() As Long

    CopyMemory p_CurrentMessage, ByVal m_lCurrMsg, 4

End Property

Private Sub Class_Initialize()

'/* This is the subclasser code. The patches are were we will write
'/* addresses of funtions so we can call them from the assembler.
'/* The exceptions are PATCH_0D and PATCH_0E, which are actually
'/* the address of the heap-allocated long where we will store
'/* the value of the current message.

Dim i          As Long
Dim nLen       As Long
Dim sHex       As String
Dim sCode      As String
Dim lIDEMode   As Long
Dim nOffset    As Long

Const CODE_STR As String = "5589E581C4E0FFFFFF5752515331C08945F88945" & _
"E068240000006800000000E8xxxxx08x89C78945EC8B4508E8DE010000897DE8FF7" & _
"5ECFF7508E8xxxxx02x8945FC09C00F8417010000EB1EE8xxxxx01x3D020000000F" & _
"842501000085C0750AE826010000E9F70000008B7DE8B2238817478B450CE895010" & _
"000B243881747B2008817897DE8FF75ECFF7508E8xxxxx03x8945F03D000000000F" & _
"8E570100008B7DE84FB223881747BB0100000089D8E85B010000FF75ECFF7508E8x" & _
"xxxx04x8945F409C07439B8xxxxx0Dx8B4D0C890868040000006800000000E8xxxx" & _
"x0Fx8945E0508B45F4508B00FF501C8B45E08B00A8027405E8DC000000E84101000" & _
"0433B5DF07F308B7DE889D8E800010000FF75ECFF7508E8xxxxx05x8945F409C074" & _
"DEB8xxxxx0Ex8B4D0C8908E810010000E9CAFFFFFF8B45F8A9FFFFFFFF75228B45F" & _
"409C0741B8B45E0508B45F4508B00FF501C8B45E08B00A8017405E875000000FF75" & _
"ECE8xxxxx09x8B45E009C0740650E8xxxxx10x5B595A5F8B45E4C9C21000E850000" & _
"000E9D6FFFFFFFF75FC68FCFFFFFFFF7508E8xxxxx07x8B7DECB2438817478B4508" & _
"E864000000FF75ECFF7508E8xxxxx06x3D000000007CA48B55EC4252FF7508E8xxx" & _
"xx0AxFF75ECFF7508E8xxxxx0BxC3FF7514FF7510FF750CFF7508FF75FCE8xxxxx0" & _
"Cx8945E4B8FFFFFFFF8945F8C38B450C3D020000007505E886FFFFFFE8CCFFFFFFE" & _
"952FFFFFF5331C931D23D000000007D07B22D881747F7D8BB0A00000031D2F7F352" & _
"4109C075F15A80C230881747E2F7B20088175BC38B45E050FF7514FF7510FF750CF" & _
"F75088B45F4508B00FF50248B45E08B008945E4C3"
    
    '/* convert patch to byte array
    sHex = CODE_STR
    nLen = LenB(sHex)
    For i = 1 To nLen Step 2
        sCode = sCode & ChrB$(Val("&H" & Mid$(sHex, i, 2)))
    Next i
    
    '/* Get the machine code length
    nLen = LenB(sCode)
    '/* Allocate fixed memory for machine code buffer
    m_lClassAddr = GlobalAlloc(0, nLen)
    '/* Allocate fixed memory for the p_CurrentMessage property
    m_lCurrMsg = GlobalAlloc(0, 4)
    '/* Copy the code to allocated memory
    CopyMemory ByVal m_lClassAddr, ByVal StrPtr(sCode), nLen
    
    '/* test in ide status (is anyone still using vb5?)
    i = Function_Address("vba6", "EbMode")
    If i = 0 Then
        i = Function_Address("vba5", "EbMode")
        If Not i = 0 Then lIDEMode = EbMode_vb5()
    Else
        lIDEMode = EbMode_vb6()
    End If
    Debug.Assert i
    
    If Not lIDEMode = 0 Then
        '/* If In_IDE Then Get the IDE patch offset
        nOffset = InStr(1, CODE_STR, "EB1E") - 1
        nOffset = nOffset \ 2
        '/* Patch the jmp short (EB1E) with two nop's (90)
        '/* enabling the IDE breakpoint/stop checking code
        CopyMemory ByVal m_lClassAddr + nOffset, &H9090, 2
        Patch_Address 77, i
    End If

    '/* patch library addresses used by assembler
    Patch_Address 59, Function_Address("user32", "GetPropA")
    Patch_Address 141, Function_Address("user32", "GetPropA")
    Patch_Address 187, Function_Address("user32", "GetPropA")
    Patch_Address 278, Function_Address("user32", "GetPropA")
    Patch_Address 434, Function_Address("user32", "GetPropA")
    Patch_Address 407, Function_Address("user32", "SetWindowLongA")
    Patch_Address 32, Function_Address("oleaut32", "SysAllocStringByteLen")
    Patch_Address 357, Function_Address("oleaut32", "SysFreeString")
    Patch_Address 454, Function_Address("user32", "RemovePropA")
    Patch_Address 465, Function_Address("user32", "RemovePropA")
    Patch_Address 486, Function_Address("user32", "CallWindowProcA")
    Patch_Address 219, Function_Address("kernel32", "GlobalAlloc")
    Patch_Address 370, Function_Address("kernel32", "GlobalFree")
    Patch_Value 199, m_lCurrMsg
    Patch_Value 290, m_lCurrMsg

End Sub

Public Sub Attach_Message(ByRef ISC As MISubclass, _
                          ByVal lHwnd As Long, _
                          ByVal lMsg As Long)

Dim lProcOld       As Long
Dim lMsgCount      As Long
Dim lMsgClassCount As Long
Dim lMsgClass      As Long

    '/* validate window
    If IsWindow(lHwnd) = False Then
        Error_Raise eeInvalidWindow
    Else
        If Not Window_Local(lHwnd) Then
            Error_Raise eeNoExternalWindow
            Exit Sub
        End If
        
        '/* check if this class is already attached for this message
        lMsgClassCount = p_MessageClassCount(lHwnd, lMsg)
        If lMsgClassCount > 0 Then
            For lMsgClass = 1 To lMsgClassCount
                If (p_MessageClass(lHwnd, lMsg, lMsgClass) = ObjPtr(ISC)) Then
                    Error_Raise eeAlreadyAttached
                    Exit Sub
                End If
            Next lMsgClass
        End If
        
        '/* associate this class with this message for this window:
        p_MessageClassCount(lHwnd, lMsg) = p_MessageClassCount(lHwnd, lMsg) + 1
        If m_lAdd = 0 Then
            '/* out of memory
            Error_Raise 5
            Exit Sub
        End If
        
        '/* associate the class pointer
        p_MessageClass(lHwnd, lMsg, p_MessageClassCount(lHwnd, lMsg)) = ObjPtr(ISC)
        If m_lAdd = 0 Then
            '/* out of memory
            p_MessageClassCount(lHwnd, lMsg) = p_MessageClassCount(lHwnd, lMsg) - 1
            Error_Raise 5
            Exit Sub
        End If
        
        '/* get the message count
        lMsgCount = p_MessageCount(lHwnd)
        If lMsgCount = 0 Then
            '/* subclass window by installing window procedure
            lProcOld = SetWindowLong(lHwnd, GWL_WNDPROC, m_lClassAddr)
            If lProcOld = 0 Then
                '/* remove class
                p_MessageClass(lHwnd, lMsg, p_MessageClassCount(lHwnd, lMsg)) = 0
                '/* remove class count
                p_MessageClassCount(lHwnd, lMsg) = p_MessageClassCount(lHwnd, lMsg) - 1
                Error_Raise eeCantSubclass
                Exit Sub
            End If
            
            '/* associate old procedure with handle
            p_OldWindowProc(lHwnd) = lProcOld
            If m_lAdd = 0 Then
                '/* failed to SetProp, put the old window proc back again
                SetWindowLong lHwnd, GWL_WNDPROC, lProcOld
                '/* remove class:
                p_MessageClass(lHwnd, lMsg, p_MessageClassCount(lHwnd, lMsg)) = 0
                '/* remove class count
                p_MessageClassCount(lHwnd, lMsg) = p_MessageClassCount(lHwnd, lMsg) - 1
                '/* raise an error
                Error_Raise 5
                Exit Sub
            End If
        End If
        
        '/* count this message
        p_MessageCount(lHwnd) = p_MessageCount(lHwnd) + 1
        If m_lAdd = 0 Then
            '/* failed to SetProp, remove class
            p_MessageClass(lHwnd, lMsg, p_MessageClassCount(lHwnd, lMsg)) = 0
            '/* remove class count contribution
            p_MessageClassCount(lHwnd, lMsg) = p_MessageClassCount(lHwnd, lMsg) - 1
            '/* if we haven't any messages on this window then remove the subclass
            If p_MessageCount(lHwnd) = 0 Then
                '/* put old window proc back again
                lProcOld = p_OldWindowProc(lHwnd)
                If Not (lProcOld = 0) Then
                    SetWindowLong lHwnd, GWL_WNDPROC, lProcOld
                    p_OldWindowProc(lHwnd) = 0
                End If
            End If
            Error_Raise 5
        End If
    End If

End Sub

Public Sub Detach_Message(ByRef ISC As MISubclass, _
                          ByVal lHwnd As Long, _
                          ByVal lMsg As Long)

Dim lMsgClassCount As Long
Dim lMsgClass      As Long
Dim lMsgClassIndex As Long
Dim lMsgCount      As Long
Dim lProcOld       As Long

    '/* validate window
    If Not IsWindow(lHwnd) = False Then
        If Not Window_Local(lHwnd) Then
            Exit Sub
        End If
        '/* check if this message is attached for this class
        lMsgClassCount = p_MessageClassCount(lHwnd, lMsg)
        If lMsgClassCount > 0 Then
            lMsgClassIndex = 0
            
            For lMsgClass = 1 To lMsgClassCount
                If (p_MessageClass(lHwnd, lMsg, lMsgClass) = ObjPtr(ISC)) Then
                    lMsgClassIndex = lMsgClass
                    Exit For
                End If
            Next lMsgClass
            
            If lMsgClassIndex = 0 Then
                '/* fail silently
                Exit Sub
            Else
                '/* remove this message class, shift up index
                For lMsgClass = lMsgClassIndex To lMsgClassCount - 1
                    p_MessageClass(lHwnd, lMsg, lMsgClass) = p_MessageClass(lHwnd, lMsg, lMsgClass + 1)
                Next lMsgClass
                '/* message class at the end can be removed
                p_MessageClass(lHwnd, lMsg, lMsgClassCount) = 0
                '/* reduce the message class count
                p_MessageClassCount(lHwnd, lMsg) = p_MessageClassCount(lHwnd, lMsg) - 1
            End If
        Else
            Exit Sub
        End If
        
        '/* reduce the message count
        lMsgCount = p_MessageCount(lHwnd)
        If lMsgCount = 1 Then
            '/* remove the subclass
            lProcOld = p_OldWindowProc(lHwnd)
            If Not (lProcOld = 0) Then
                '/* unsubclass by reassigning old window procedure
                SetWindowLong lHwnd, GWL_WNDPROC, lProcOld
            End If
            '/* remove the old window proc
            p_OldWindowProc(lHwnd) = 0
        End If
        p_MessageCount(lHwnd) = p_MessageCount(lHwnd) - 1
    End If

End Sub

Public Function Window_Local(ByVal lHwnd As Long) As Boolean

Dim idWnd As Long

    GetWindowThreadProcessId lHwnd, idWnd
    Window_Local = (idWnd = GetCurrentProcessId())

End Function

Private Function Function_Address(ByVal sDLL As String, _
                                  ByVal sProc As String) As Long

    Function_Address = GetProcAddress(GetModuleHandle(sDLL), sProc)
    Debug.Assert Function_Address

End Function

Private Sub Patch_Address(ByVal lOffset As Long, _
                          ByVal lTarget As Long)

    CopyMemory ByVal (m_lClassAddr + lOffset), lTarget - m_lClassAddr - lOffset - 4, 4

End Sub

Private Sub Patch_Value(ByVal lOffset As Long, _
                        ByVal lValue As Long)

    CopyMemory ByVal (m_lClassAddr + lOffset), lValue, 4

End Sub

Public Function Get_OldProc(ByVal lHwnd As Long, _
                            ByVal lMsg As Long, _
                            ByVal wParam As Long, _
                            ByVal lParam As Long) As Long

Dim lProcOld As Long

    lProcOld = p_OldWindowProc(lHwnd)
    If Not (lProcOld = 0) Then
        Get_OldProc = CallWindowProc(lProcOld, lHwnd, lMsg, wParam, lParam)
    End If

End Function

Private Function Ide_True(bValue As Boolean) As Boolean

    Ide_True = True
    bValue = True

End Function

Public Function In_IDE() As Boolean

    Debug.Assert Ide_True(In_IDE)

End Function

Private Sub Error_Raise(ByVal lErr As Long)

Dim sText   As String
Dim sSource As String

    If lErr > 1000 Then
        sSource = App.EXEName + ".WindowProc"
        Select Case lErr
        Case eeCantSubclass
            sText = "Can't subclass window"
        Case eeAlreadyAttached
            sText = "Message already handled by another class"
        Case eeInvalidWindow
            sText = "Invalid window"
        Case eeNoExternalWindow
            sText = "Can't modify external window"
        End Select
        Err.Raise lErr Or vbObjectError, sSource, sText
    Else
        Err.Raise lErr, sSource
    End If

End Sub

Public Sub Clean_Up(ByVal lHwnd As Long)

Dim lMsgCount As Long
Dim lProcOld  As Long

    '/* explicit (forced) exit
    lMsgCount = p_MessageCount(lHwnd)
    If lMsgCount > 0 Then
        '/* remove the subclass:
        lProcOld = p_OldWindowProc(lHwnd)
        If Not (lProcOld = 0) Then
            '/* unsubclass by reassigning old window procedure
            SetWindowLong lHwnd, GWL_WNDPROC, lProcOld
        End If
        '/* remove the old window proc
        p_OldWindowProc(lHwnd) = 0
        p_MessageCount(lHwnd) = 0
    End If

End Sub

Private Sub Class_Terminate()

    GlobalFree m_lCurrMsg
    GlobalFree m_lClassAddr

End Sub

