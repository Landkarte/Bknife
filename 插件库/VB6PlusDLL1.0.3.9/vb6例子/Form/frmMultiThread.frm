VERSION 5.00
Begin VB.Form frmMultiThread 
   Caption         =   "多线程函数示例"
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
      ItemData        =   "frmMultiThread.frx":0000
      Left            =   240
      List            =   "frmMultiThread.frx":000A
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
Attribute VB_Name = "frmMultiThread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NowPath As String
    
Private Sub cmdDo_Click()
    Dim OutStr As String
    Dim ErrText As String
    Dim Total As Long
    Dim MTData() As String
    Dim FuncParas() As Variant
    Dim RunResults() As Variant
    
    txtLog.Text = ""
    
    cmdDo.Enabled = False
    
    Select Case cmbCommand.List(cmbCommand.ListIndex)
           Case "Net_MT"
                ReDim MTData(1, 5)
                MTData(0, 0) = "https://www.baidu.com/"
                MTData(0, 1) = "GET"
                MTData(0, 4) = "1"
                MTData(1, 0) = "https://www.baidu.com/"
                MTData(1, 1) = "GET"
                MTData(1, 4) = "1"
                If Net_MT(MTData) = 1 Then
                    OutStr = "【线程1数据摘要】" & vbCrLf & VBA.Left(MTData(0, 5), 500) & "..." & vbCrLf & vbCrLf
                    OutStr = OutStr & "【线程2数据摘要】" & vbCrLf & VBA.Left(MTData(1, 5), 500) & "..." & vbCrLf & vbCrLf
                End If
                Erase MTData
           Case "RunVBFunction_MT"
                ReDim MTData(1, 2)
                ReDim FuncParas(1, 3)
                ReDim RunResults(1)
                
                MTData(0, 0) = "function Add(a,b)" & vbCrLf & "Add=a+b" & vbCrLf & "end function" & vbCrLf
                MTData(0, 1) = "Add"
                FuncParas(0, 0) = 2 '参数个数
                FuncParas(0, 1) = 1 '参数1
                FuncParas(0, 2) = 2 '参数2
                
                MTData(1, 0) = "function Add(a,b,c)" & vbCrLf & "Add=a+b+c" & vbCrLf & "end function" & vbCrLf
                MTData(1, 1) = "Add"
                FuncParas(1, 0) = 3 '参数个数
                FuncParas(1, 1) = 1 '参数1
                FuncParas(1, 2) = 2 '参数2
                FuncParas(1, 3) = 3 '参数3
                
                If RunVBFunction_MT(MTData, FuncParas, RunResults) = 1 Then
                    OutStr = "【线程1】执行结果：" & MTData(0, 2) & "，返回数据：" & RunResults(0) & vbCrLf & vbCrLf
                    OutStr = OutStr & "【线程2】执行结果：" & MTData(1, 2) & "，返回数据：" & RunResults(1) & vbCrLf & vbCrLf
                End If
                                
                Erase MTData, FuncParas, RunResults
    End Select
    txtLog.Text = OutStr
    
    cmdDo.Enabled = True
End Sub


Private Sub Form_Load()
    NowPath = App.Path
    cmbCommand.ListIndex = 0
End Sub

