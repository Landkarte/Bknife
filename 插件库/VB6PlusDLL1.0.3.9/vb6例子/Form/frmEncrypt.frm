VERSION 5.00
Begin VB.Form frmEncrypt 
   Caption         =   "加密函数示例"
   ClientHeight    =   8985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12660
   LinkTopic       =   "Form1"
   ScaleHeight     =   8985
   ScaleWidth      =   12660
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdAESDecrypt_GB_HEX 
      Caption         =   "↑AESDecrypt_GB(HEX)"
      Height          =   495
      Left            =   8760
      TabIndex        =   19
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CommandButton cmdLen 
      Caption         =   "长度"
      Height          =   735
      Left            =   12240
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   4455
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   4080
      Width           =   12015
   End
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmEncrypt.frx":0000
      Top             =   240
      Width           =   12015
   End
   Begin VB.CommandButton cmdAESEncrypt_UTF8_Hex 
      Caption         =   "AESEncrypt_UTF8(HEX)↓"
      Height          =   495
      Left            =   2280
      TabIndex        =   17
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CommandButton cmdAESEncrypt_GB_Hex 
      Caption         =   "AESEncrypt_GB(HEX)↓"
      Height          =   495
      Left            =   240
      TabIndex        =   18
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton cmdAESEncrypt_UTF8 
      Caption         =   "AESEncrypt_UTF8↓"
      Height          =   495
      Left            =   4560
      TabIndex        =   13
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdAESEncrypt_GB 
      Caption         =   "AESEncrypt_GB↓"
      Height          =   495
      Left            =   2880
      TabIndex        =   11
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdMD516_UTF8 
      Caption         =   "MD516_UTF8↓"
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdMD516_GB 
      Caption         =   "MD516_GB↓"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAESEncryptFile 
      Caption         =   "AESEncryptTxtFile↓"
      Height          =   495
      Left            =   4080
      TabIndex        =   15
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cmdBase64Encode_UTF8 
      Caption         =   "Base64Encode_UTF8↓"
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cmdBase64Encode_GB 
      Caption         =   "Base64Encode_GB↓"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cmdAESDecrypt_UTF8_HEX 
      Caption         =   "↑AESDecrypt_UTF8(HEX)"
      Height          =   495
      Left            =   6480
      TabIndex        =   20
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CommandButton cmdMD532_GB 
      Caption         =   "↑MD532_GB"
      Height          =   495
      Left            =   11280
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdMD532_UTF8 
      Caption         =   "↑MD532_UTF8"
      Height          =   495
      Left            =   9840
      TabIndex        =   10
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdAESDecrypt_GB 
      Caption         =   "↑AESDecrypt_GB"
      Height          =   495
      Left            =   8160
      TabIndex        =   12
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdAESDecrypt_UTF8 
      Caption         =   "↑AESDecrypt_UTF8"
      Height          =   495
      Left            =   6480
      TabIndex        =   14
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdBase64Decode_GB 
      Caption         =   "↑Base64Decode_GB"
      Height          =   495
      Left            =   10680
      TabIndex        =   3
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton cmdBase64Decode_UTF8 
      Caption         =   "↑Base64Decode_UTF8"
      Height          =   495
      Left            =   8760
      TabIndex        =   6
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cmdAESDecryptFile 
      Caption         =   "↑AESDecryptTxtFile"
      Height          =   495
      Left            =   6720
      TabIndex        =   16
      Top             =   2640
      Width           =   2055
   End
End
Attribute VB_Name = "frmEncrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NowPath As String
    
Private Sub cmdAESDecrypt_GB_Click()
    Text1.Text = AESDecrypt_GB(Text2.Text, "123", , AESMode.ECB, 0)
End Sub

Private Sub cmdAESDecrypt_GB_HEX_Click()
    Text1.Text = AESDecrypt_GB(Text2.Text, "123", , AESMode.ECB, 0, AESCodeType.Hex)
End Sub

Private Sub cmdAESDecrypt_UTF8_Click()
    Text1.Text = AESDecrypt_UTF8(Text2.Text, "123", , AESMode.ECB, 0)
End Sub

Private Sub cmdAESDecrypt_UTF8_HEX_Click()
    Text1.Text = AESDecrypt_UTF8(Text2.Text, "123", , AESMode.ECB, 0, AESCodeType.Hex)
End Sub

Private Sub cmdAESDecryptFile_Click()
    Text2.Text = AESDecryptTxtFile_GB(NowPath & "\Encrypt.txt", NowPath & "\Decrypt.txt", "123", , AESMode.ECB)
End Sub

Private Sub cmdAESEncrypt_GB_Click()
    Text2.Text = AESEncrypt_GB(Text1.Text, "123", , AESMode.ECB, 0)
End Sub

Private Sub cmdAESEncrypt_GB_Hex_Click()
    Text2.Text = AESEncrypt_GB(Text1.Text, "123", , AESMode.ECB, 0, AESCodeType.Hex)
End Sub

Private Sub cmdAESEncrypt_UTF8_Click()
    Text2.Text = AESEncrypt_UTF8(Text1.Text, "123", , AESMode.ECB, 0)
End Sub

Private Sub cmdAESEncrypt_UTF8_Hex_Click()
    Text2.Text = AESEncrypt_UTF8(Text1.Text, "123", , AESMode.ECB, 0, AESCodeType.Hex)
End Sub

Private Sub cmdAESEncryptFile_Click()
    Text2.Text = AESEncryptTxtFile_GB(NowPath & "\测试字符串.txt", NowPath & "\Encrypt.txt", "123", , AESMode.ECB)
End Sub

Private Sub cmdBase64Decode_GB_Click()
    Text1.Text = Base64Decode_GB(Text2.Text)
End Sub

Private Sub cmdBase64Decode_UTF8_Click()
    Text1.Text = Base64Decode_UTF8(Text2.Text)
End Sub

Private Sub cmdBase64Encode_GB_Click()
    Text2.Text = Base64Encode_GB(Text1.Text)
End Sub

Private Sub cmdBase64Encode_UTF8_Click()
    Text2.Text = Base64Encode_UTF8(Text1.Text)
End Sub

Private Sub cmdLen_Click()
    MsgBox Len(Text1.Text)
End Sub

Private Sub cmdMD516_GB_Click()
    Dim StrResult As String
    StrResult = MD516_GB(Text1.Text)
    Debug.Print StrResult
    Text2.Text = StrResult
End Sub

Private Sub cmdMD516_UTF8_Click()
    Text2.Text = MD516_UTF8(Text1.Text)
End Sub

Private Sub cmdMD532_GB_Click()
    Text1.Text = MD532_GB(Text2.Text)
End Sub

Private Sub cmdMD532_UTF8_Click()
    Text1.Text = MD532_UTF8(Text2.Text)
End Sub

Private Sub Form_Load()
    NowPath = App.Path
End Sub
