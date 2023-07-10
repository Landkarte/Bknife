VERSION 5.00
Begin VB.Form frmWindows 
   Caption         =   "Windows函数示例"
   ClientHeight    =   5445
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   9120
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtLog 
      Height          =   4335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   960
      Width           =   8775
   End
   Begin VB.CommandButton cmdDo 
      Caption         =   "执行"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.ComboBox cmbCommand 
      Height          =   300
      ItemData        =   "frmWindows.frx":0000
      Left            =   240
      List            =   "frmWindows.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "日志"
      Height          =   180
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   360
   End
End
Attribute VB_Name = "frmWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NowPath As String
    
Private Sub cmdDo_Click()
    Dim StrScript As String
    Dim StrError As String
    Dim Paras() As Variant
    Dim Result As Variant
    
    txtLog.Text = ""
    
    cmdDo.Enabled = False
    
    Select Case cmbCommand.List(cmbCommand.ListIndex)
           Case "Win_CopyFileToClipBoard"
                txtLog.Text = Win_CopyFileToClipBoard(NowPath & "\测试字符串.txt")
           Case "RunVBScript"
                StrScript = "msgbox(""Hello!"")" & vbCrLf
                If RunVBScript(StrScript, StrError) = 0 Then
                    txtLog.Text = StrScript & "执行错误：" & StrError
                Else
                    txtLog.Text = StrScript & "执行成功！"
                End If
           Case "RunVBFunction"
                StrScript = "function Add(a,b)" & vbCrLf & "Add=a+b" & vbCrLf & "end function" & vbCrLf
                ReDim Paras(1)
                Paras(0) = 1
                Paras(1) = 2
                If RunVBFunction(StrScript, "add", Paras, Result, StrError) = 0 Then
                    txtLog.Text = StrScript & "执行错误：" & StrError
                Else
                    txtLog.Text = StrScript & "执行成功！执行结果：" & Result
                End If
          Case "SetFormIcon"
                SetFormIcon hwnd, "hello.ico"
                txtLog.Text = "设置窗体图标。"
    End Select
    
    cmdDo.Enabled = True
End Sub

Private Sub Form_Load()
    NowPath = App.Path
    cmbCommand.ListIndex = 0
End Sub
