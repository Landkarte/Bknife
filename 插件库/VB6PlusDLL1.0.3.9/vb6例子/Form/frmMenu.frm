VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "VB6Plus.dllʾ��������Ŀ¼"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   7320
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdMQTT 
      Caption         =   "������MQTT"
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   4560
      Width           =   3495
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "�ļ���������"
      Height          =   975
      Left            =   3720
      TabIndex        =   9
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton cmdMultiThread 
      Caption         =   "���̺߳���"
      Height          =   735
      Left            =   3720
      TabIndex        =   8
      Top             =   3720
      Width           =   3495
   End
   Begin VB.CommandButton cmdDB 
      Caption         =   "���ݿ⺯��"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   3495
   End
   Begin VB.CommandButton cmdNet 
      Caption         =   "���纯��"
      Height          =   735
      Left            =   3720
      TabIndex        =   6
      Top             =   2880
      Width           =   3495
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "�Ի�����"
      Height          =   735
      Left            =   3720
      TabIndex        =   5
      Top             =   2040
      Width           =   3495
   End
   Begin VB.CommandButton cmdGraph 
      Caption         =   "ͼ�κ���"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   3495
   End
   Begin VB.CommandButton cmdString 
      Caption         =   "�ַ�������"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton cmdWindows 
      Caption         =   "Windows����"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   3495
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "���ܺ���"
      Height          =   735
      Left            =   3720
      TabIndex        =   2
      Top             =   1200
      Width           =   3495
   End
   Begin VB.CommandButton cmdHTML 
      Caption         =   "HTML����"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   3495
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdDB_Click()
    frmDB.Show
    Unload Me
End Sub

Private Sub cmdDialog_Click()
    frmDialog.Show
    Unload Me
End Sub

Private Sub cmdEncrypt_Click()
    frmEncrypt.Show
    Unload Me
End Sub

Private Sub cmdFile_Click()
    frmFile.Show
    Unload Me
End Sub

Private Sub cmdGraph_Click()
    frmGraph.Show
    Unload Me
End Sub

Private Sub cmdHTML_Click()
    frmHTML.Show
    Unload Me
End Sub

Private Sub cmdMQTT_Click()
    frmMQTT.Show
    Unload Me
End Sub

Private Sub cmdMultiThread_Click()
    frmMultiThread.Show
    Unload Me
End Sub

Private Sub cmdNet_Click()
    frmNet.Show
    Unload Me
End Sub

Private Sub cmdString_Click()
    frmString.Show
    Unload Me
End Sub

Private Sub cmdWindows_Click()
    frmWindows.Show
    Unload Me
End Sub
