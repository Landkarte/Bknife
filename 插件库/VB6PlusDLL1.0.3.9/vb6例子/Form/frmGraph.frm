VERSION 5.00
Begin VB.Form frmGraph 
   Caption         =   "ͼ�κ���ʾ��"
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
      ItemData        =   "frmGraph.frx":0000
      Left            =   240
      List            =   "frmGraph.frx":0010
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
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const HomePage As String = "1Ŀǰ�й�ʵʩ��������ұ�׼������2010�꣬���ڶԵ����ʡ�������������ؼ�ָ��涨���ͣ���������1986��ľɰ�������꣬����10����һֱ�ܵ�ҵ��������ߵ����ɡ�" 'http://www.pygzs.com/"
Dim NowPath As String
    
Private Sub cmdDo_Click()
    Dim Str_A As String
    Dim Str_B As String
    Dim ErrText As String
    
    txtLog.Text = ""
    
    cmdDo.Enabled = False
    
    Select Case cmbCommand.List(cmbCommand.ListIndex)
           Case "MakeQRCode"
                txtLog.Text = "����" & HomePage & "���ɶ�ά��ͼ(" & NowPath & "\QRCode.jpg" & ")��" & vbCrLf & "���ɽ����"
                txtLog.Text = txtLog.Text & MakeQRCode(HomePage, NowPath & "\QRCode.jpg")
           Case "ScanQRImage"
                txtLog.Text = "��������" & ScanQRImage(NowPath & "\QRCode.jpg", , ErrText, 0) & vbCrLf
                If Len(ErrText) > 0 Then txtLog.Text = txtLog.Text & "ʧ����Ϣ��" & ErrText
           Case "ImageToJPG"
                txtLog.Text = "ͼ���ļ�ת��JPG�ļ��Ľ����" & ImageToJPG(NowPath & "\QRCode.bmp", NowPath & "\QRCode.jpg")
           Case "ImageToBMP"
                txtLog.Text = "ͼ���ļ�ת��BMP�ļ��Ľ����" & ImageToBMP(NowPath & "\QRCode.jfif", NowPath & "\QRCode.bmp")
    End Select
    
    cmdDo.Enabled = True
End Sub

Private Sub Form_Load()
    NowPath = App.Path
    cmbCommand.ListIndex = 0
End Sub

