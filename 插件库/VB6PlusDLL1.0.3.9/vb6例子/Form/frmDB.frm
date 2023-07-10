VERSION 5.00
Begin VB.Form frmDB 
   Caption         =   "数据库处理函数示例"
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
      ItemData        =   "frmDB.frx":0000
      Left            =   240
      List            =   "frmDB.frx":000A
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
Attribute VB_Name = "frmDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NowPath As String
    
Private Sub cmdDo_Click()
    Dim Str_A As String
    Dim Str_B As String
    Dim OutStr As String
    Dim Total As Long
    Dim SQLiteDBLong As Long
    Dim LngResult As Long
    Dim Data() As String
    Dim StrErr As String
    Dim i As Long, j As Long
    
    txtLog.Text = ""
    
    cmdDo.Enabled = False
    
    Select Case cmbCommand.List(cmbCommand.ListIndex)
           Case "SQLite标准读取"
                LngResult = SQLite_Open(SQLiteDBLong, "DB/DB.DB", StrErr)
                If LngResult = 1 Then
                    LngResult = SQLite_ReadData(SQLiteDBLong, "SELECT * From [Test] limit 1000", Data, StrErr)
                    If LngResult = 1 Then
                        
                        OutStr = OutStr & "记录数:" & UBound(Data, 1) & ","
                        OutStr = OutStr & "列数:" & UBound(Data, 2) & vbCrLf
                        For i = 1 To UBound(Data, 1)
                            OutStr = OutStr & "第" & i & "条记录："
                            For j = 1 To UBound(Data, 2)
                                OutStr = OutStr & Data(i, j) & vbTab
                            Next
                            OutStr = OutStr & vbCrLf
                        Next
                        Erase Data
                    Else
                        OutStr = "数据读取错误:" & StrErr
                    End If
                    LngResult = SQLite_Close(SQLiteDBLong)
                Else
                    OutStr = "数据库打开失败！" & StrErr
                End If
                txtLog.Text = OutStr
           Case "SQLite执行"
                LngResult = SQLite_Open(SQLiteDBLong, "DB/DB.DB", StrErr)
                If LngResult = 1 Then
                    LngResult = SQLite_Execute(SQLiteDBLong, "INSERT INTO [Test]([Title])VALUES('标题')", StrErr)
                    If LngResult = 1 Then
                        OutStr = "执行成功！"
                    Else
                        OutStr = "数据执行错误:" & StrErr
                    End If
                    LngResult = SQLite_Close(SQLiteDBLong)
                Else
                    OutStr = "数据库打开失败！" & StrErr
                End If
                txtLog.Text = OutStr
           Case "Permutation"
                Str_A = "ABCDEFG"
                OutStr = Str_A & "排列结果：" & vbCrLf & Permutation(Str_A, vbCrLf, Total)
                txtLog.Text = OutStr & vbCrLf & "共有" & Total & "种排列结果"
           Case "Combination"
                Str_A = "ABCDEFG"
                OutStr = Str_A & "组合结果：" & vbCrLf & Combination(Str_A, vbCrLf, Total)
                txtLog.Text = OutStr & vbCrLf & "共有" & Total & "种组合结果"
           Case "StrToHex_GB"
                Str_A = "你好!"
                OutStr = StrToHex_GB(Str_A)
                txtLog.Text = "“" & Str_A & "”转换为HEX(GB2312)结果：" & vbCrLf & OutStr
           Case "StrToHex_UTF8"
                Str_A = "你好!"
                OutStr = StrToHex_UTF8(Str_A)
                txtLog.Text = "“" & Str_A & "”转换为HEX(UTF8)结果：" & vbCrLf & OutStr
           Case "HexToStr_GB"
                Str_A = "C4E3BAC321"
                OutStr = HexToStr_GB(Str_A)
                txtLog.Text = "“" & Str_A & "”GB的HEX解码结果：" & vbCrLf & OutStr
           Case "HexToStr_UTF8"
                Str_A = "E4BDA0E5A5BD21"
                OutStr = HexToStr_UTF8(Str_A)
                txtLog.Text = "“" & Str_A & "”UTF8的HEX解码结果：" & vbCrLf & OutStr
    End Select
    
    cmdDo.Enabled = True
End Sub


Private Sub Form_Load()
    NowPath = App.Path
    cmbCommand.ListIndex = 0
End Sub

