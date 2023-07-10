VERSION 5.00
Begin VB.Form add 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "添加"
   ClientHeight    =   3150
   ClientLeft      =   9420
   ClientTop       =   3555
   ClientWidth     =   7170
   Icon            =   "add.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox mm 
      Height          =   315
      Left            =   5115
      TabIndex        =   6
      Text            =   "fone"
      Top             =   285
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定添加"
      Height          =   495
      Left            =   4230
      TabIndex        =   4
      Top             =   2535
      Width           =   2730
   End
   Begin VB.TextBox bz 
      Height          =   1335
      Left            =   705
      TabIndex        =   3
      Text            =   "一直很安静"
      Top             =   765
      Width           =   3825
   End
   Begin VB.TextBox url 
      Height          =   375
      Left            =   735
      TabIndex        =   1
      Text            =   "http://127.0.0.1/body.php"
      Top             =   300
      Width           =   3735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "仅支持PHP语言"
      Height          =   180
      Left            =   5010
      TabIndex        =   7
      Top             =   1290
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "密码:"
      Height          =   180
      Left            =   4605
      TabIndex        =   5
      Top             =   330
      Width           =   450
   End
   Begin VB.Label Label2 
      Caption         =   "备注:"
      Height          =   465
      Left            =   180
      TabIndex        =   2
      Top             =   795
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "url:"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   360
   End
End
Attribute VB_Name = "add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim x
Dim urltext As String
Dim bztext As String
Dim mima As String
urltext = url.text
bztext = bz.text
mima = mm.text
    x = Form1.ListView1.ListItems.Count + 1
    Form1.ListView1.ListItems.add , , x, , 1
    Form1.ListView1.ListItems(x).SubItems(1) = "PHP"
    Form1.ListView1.ListItems(x).SubItems(2) = urltext   '地址
    Form1.ListView1.ListItems(x).SubItems(3) = mima
    Form1.ListView1.ListItems(x).SubItems(4) = "错误"
    Form1.ListView1.ListItems(x).SubItems(5) = Now & Time
    Form1.ListView1.ListItems(x).SubItems(6) = Now & Time

    Form1.ListView1.ListItems(x).SubItems(7) = bztext
Me.Hide
End Sub

