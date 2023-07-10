VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MISubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum EMsgResponse
    emrConsume = 0&      ' Process instead of original WindowProc
    emrPostProcess = 1&  ' Process after original WindowProc
    emrPreProcess = 2&   ' Process before original WindowProc
End Enum

Public MsgResponse As EMsgResponse


Function WindowProc(ByVal hwnd As Long, _
                    ByVal iMsg As Long, _
                    ByVal wParam As Long, _
                    ByVal lParam As Long) As Long
End Function

