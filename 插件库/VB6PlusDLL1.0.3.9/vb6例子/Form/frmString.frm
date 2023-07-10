VERSION 5.00
Begin VB.Form frmString 
   Caption         =   "�ַ�������ʾ��"
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
      ItemData        =   "frmString.frx":0000
      Left            =   240
      List            =   "frmString.frx":0019
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
Attribute VB_Name = "frmString"
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
    Dim ErrText As String
    Dim Total As Long
    
    txtLog.Text = ""
    
    cmdDo.Enabled = False
    
    Select Case cmbCommand.List(cmbCommand.ListIndex)
           Case "StrCompare"
                Str_A = "ABCDEFG"
                Str_B = "ABC12FG"
                txtLog.Text = Str_A & "��" & Str_B & "���ƶȣ�"
                txtLog.Text = txtLog.Text & StrCompare(Str_A, Str_B)
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

