VERSION 5.00
Begin VB.Form frmRunVB 
   Caption         =   "执行VB脚本"
   ClientHeight    =   4380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   8760
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "执行VB脚本"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmRunVB.frx":0000
      Top             =   120
      Width           =   8415
   End
End
Attribute VB_Name = "frmRunVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub Command1_Click()
    Dim VBScript As String, Error As String
    Dim StartTime As Long, EndTime As Long
    Dim Paras() As Variant
    Dim Result As Variant
    
    StartTime = GetTickCount
    
    VBScript = Text1.Text
    ReDim Paras(1)
    Paras(0) = 1
    Paras(1) = 2
    Error = ""
    If RunVBFunction(VBScript, "Add", Paras, Result, Error) = 0 Then
        MsgBox Error
    Else
        Debug.Print "Result:" & Result
        EndTime = GetTickCount
        MsgBox "执行成功！共花费时间" & Round((EndTime - StartTime) / 1000, 3) & "秒"
    End If
    
    
End Sub
