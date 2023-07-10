VERSION 5.00
Begin VB.Form frmNet 
   Caption         =   "网络处理函数示例"
   ClientHeight    =   5325
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   9060
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox cmbCommand 
      Height          =   300
      ItemData        =   "frmNet.frx":0000
      Left            =   240
      List            =   "frmNet.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton cmdDo 
      Caption         =   "执行"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtLog 
      Height          =   4335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   840
      Width           =   8775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "日志"
      Height          =   180
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   360
   End
End
Attribute VB_Name = "frmNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NowPath As String

Private Sub cmdDo_Click()
    On Error GoTo ErrorHand
    
    Dim Str_A As String
    Dim Str_B As String
    Dim OutStr As String
    Dim ErrText As String
    Dim Total As Long
    Dim RequestHeader As String, ResponseHeaders As String, Result As String
    
    txtLog.Text = ""
    
    cmdDo.Enabled = False

    Select Case cmbCommand.List(cmbCommand.ListIndex)
           Case "XMLHTTP_Get"
                txtLog.Text = XMLHTTP_Get("https://www.baidu.com")
           Case "XMLHTTP_Post"
                txtLog.Text = XMLHTTP_Post("https://api.apiopen.top/musicRankings", "")
           Case "OpenSSL_Get"
                txtLog.Text = OpenSSL_Get("https://www.baidu.com", , , , 0)
           Case "OpenSSL_Post"
                txtLog.Text = OpenSSL_Post("https://api.apiopen.top/getSingleJoke?sid=28654780", "")
    End Select

    cmdDo.Enabled = True
    Exit Sub
ErrorHand:
    MsgBox Err.Description
    Err.Clear
End Sub


Private Sub Form_Load()
    NowPath = App.Path
    cmbCommand.ListIndex = 0
End Sub

