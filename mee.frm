VERSION 5.00
Begin VB.Form mee 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ڳ���"
   ClientHeight    =   4920
   ClientLeft      =   3585
   ClientTop       =   6975
   ClientWidth     =   5685
   Icon            =   "mee.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   5685
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "����İ�һֱ�ܰ���"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   1980
      TabIndex        =   1
      Top             =   4305
      Width           =   1620
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "VBд��Webshell ������"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   1725
      TabIndex        =   0
      Top             =   4020
      Width           =   2070
   End
   Begin VB.Image Image1 
      Height          =   3750
      Left            =   840
      Picture         =   "mee.frx":21E2
      Top             =   120
      Width           =   3750
   End
End
Attribute VB_Name = "mee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

