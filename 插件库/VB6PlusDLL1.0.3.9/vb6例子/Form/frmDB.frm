VERSION 5.00
Begin VB.Form frmDB 
   Caption         =   "���ݿ⴦����ʾ��"
   ClientHeight    =   5325
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   9060
   StartUpPosition =   3  '����ȱʡ
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
      Caption         =   "ִ��"
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
      Caption         =   "��־"
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
           Case "SQLite��׼��ȡ"
                LngResult = SQLite_Open(SQLiteDBLong, "DB/DB.DB", StrErr)
                If LngResult = 1 Then
                    LngResult = SQLite_ReadData(SQLiteDBLong, "SELECT * From [Test] limit 1000", Data, StrErr)
                    If LngResult = 1 Then
                        
                        OutStr = OutStr & "��¼��:" & UBound(Data, 1) & ","
                        OutStr = OutStr & "����:" & UBound(Data, 2) & vbCrLf
                        For i = 1 To UBound(Data, 1)
                            OutStr = OutStr & "��" & i & "����¼��"
                            For j = 1 To UBound(Data, 2)
                                OutStr = OutStr & Data(i, j) & vbTab
                            Next
                            OutStr = OutStr & vbCrLf
                        Next
                        Erase Data
                    Else
                        OutStr = "���ݶ�ȡ����:" & StrErr
                    End If
                    LngResult = SQLite_Close(SQLiteDBLong)
                Else
                    OutStr = "���ݿ��ʧ�ܣ�" & StrErr
                End If
                txtLog.Text = OutStr
           Case "SQLiteִ��"
                LngResult = SQLite_Open(SQLiteDBLong, "DB/DB.DB", StrErr)
                If LngResult = 1 Then
                    LngResult = SQLite_Execute(SQLiteDBLong, "INSERT INTO [Test]([Title])VALUES('����')", StrErr)
                    If LngResult = 1 Then
                        OutStr = "ִ�гɹ���"
                    Else
                        OutStr = "����ִ�д���:" & StrErr
                    End If
                    LngResult = SQLite_Close(SQLiteDBLong)
                Else
                    OutStr = "���ݿ��ʧ�ܣ�" & StrErr
                End If
                txtLog.Text = OutStr
           Case "Permutation"
                Str_A = "ABCDEFG"
                OutStr = Str_A & "���н����" & vbCrLf & Permutation(Str_A, vbCrLf, Total)
                txtLog.Text = OutStr & vbCrLf & "����" & Total & "�����н��"
           Case "Combination"
                Str_A = "ABCDEFG"
                OutStr = Str_A & "��Ͻ����" & vbCrLf & Combination(Str_A, vbCrLf, Total)
                txtLog.Text = OutStr & vbCrLf & "����" & Total & "����Ͻ��"
           Case "StrToHex_GB"
                Str_A = "���!"
                OutStr = StrToHex_GB(Str_A)
                txtLog.Text = "��" & Str_A & "��ת��ΪHEX(GB2312)�����" & vbCrLf & OutStr
           Case "StrToHex_UTF8"
                Str_A = "���!"
                OutStr = StrToHex_UTF8(Str_A)
                txtLog.Text = "��" & Str_A & "��ת��ΪHEX(UTF8)�����" & vbCrLf & OutStr
           Case "HexToStr_GB"
                Str_A = "C4E3BAC321"
                OutStr = HexToStr_GB(Str_A)
                txtLog.Text = "��" & Str_A & "��GB��HEX��������" & vbCrLf & OutStr
           Case "HexToStr_UTF8"
                Str_A = "E4BDA0E5A5BD21"
                OutStr = HexToStr_UTF8(Str_A)
                txtLog.Text = "��" & Str_A & "��UTF8��HEX��������" & vbCrLf & OutStr
    End Select
    
    cmdDo.Enabled = True
End Sub


Private Sub Form_Load()
    NowPath = App.Path
    cmbCommand.ListIndex = 0
End Sub

