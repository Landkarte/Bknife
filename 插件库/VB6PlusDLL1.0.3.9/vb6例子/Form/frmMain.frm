VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "VB6Plus.dll示例"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   10995
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdHTMLDecode 
      Caption         =   "↑HTMLDecode"
      Height          =   495
      Left            =   9000
      TabIndex        =   10
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmdHTMLEncode 
      Caption         =   "HTMLEncode↓"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmdUnicodeEncode 
      Caption         =   "UnicodeEncode↓"
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdLen 
      Caption         =   "长度"
      Height          =   735
      Left            =   10560
      TabIndex        =   6
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton cmdURLDecodeGB 
      Caption         =   "↑UrlEncode_GB"
      Height          =   495
      Left            =   9000
      TabIndex        =   4
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdURLEncodeUTF8 
      Caption         =   "UrlEncode_UTF8↓"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   2415
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3600
      Width           =   10335
   End
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmMain.frx":0000
      Top             =   240
      Width           =   10335
   End
   Begin VB.CommandButton cmdURLEncodeGB 
      Caption         =   "UrlEncode_GB↓"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdURLDecodeUTF8 
      Caption         =   "↑UrlDecode_UTF8"
      Height          =   495
      Left            =   7320
      TabIndex        =   5
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdUnicodeDecode 
      Caption         =   "↑UnicodeDecode"
      Height          =   495
      Left            =   5640
      TabIndex        =   8
      Top             =   2640
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHTMLDecode_Click()
    Text1.Text = HTMLDecode(Text2.Text)
End Sub

Private Sub cmdHTMLEncode_Click()
    Text2.Text = HTMLEncode(Text1.Text)
End Sub

Private Sub cmdLen_Click()
    MsgBox Len(Text1.Text)
End Sub

Private Sub cmdUnicodeDecode_Click()
    Text1.Text = UnicodeDecode(Text2.Text)
End Sub

Private Sub cmdUnicodeEncode_Click()
    Text2.Text = UnicodeEncode(Text1.Text)
End Sub

Private Sub cmdURLDecodeGB_Click()
    Text1.Text = UrlDecode_GB(Text2.Text)
End Sub

Private Sub cmdURLDecodeUTF8_Click()
    Text1.Text = UrlDecode_UTF8(Text2.Text)
End Sub

Private Sub cmdURLEncodeGB_Click()
    Text2.Text = UrlEncode_GB(Text1.Text)
End Sub

Private Sub cmdURLEncodeUTF8_Click()
    Text2.Text = UrlEncode_UTF8(Text1.Text)
End Sub

