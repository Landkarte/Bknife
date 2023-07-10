VERSION 5.00
Begin VB.Form frmMQTT 
   Caption         =   "MQTT函数示例"
   ClientHeight    =   5670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11985
   Icon            =   "frmMQTT.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   11985
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtPassWord 
      Height          =   375
      Left            =   8280
      TabIndex        =   16
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox txtUserName 
      Height          =   375
      Left            =   5880
      TabIndex        =   14
      Top             =   720
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Left            =   11160
      Top             =   720
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭"
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "打开"
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton txtPub 
      Caption         =   "发布"
      Height          =   495
      Left            =   10680
      TabIndex        =   10
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox txtMsg 
      Height          =   495
      Left            =   1200
      TabIndex        =   9
      Text            =   "Hello World!"
      Top             =   4680
      Width           =   9495
   End
   Begin VB.TextBox txtReceivedMsgs 
      Height          =   3135
      Left            =   1200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1440
      Width           =   10455
   End
   Begin VB.TextBox txtTopic 
      Height          =   375
      Left            =   8280
      TabIndex        =   5
      Text            =   "MQTT Examples"
      Top             =   120
      Width           =   2895
   End
   Begin VB.TextBox txtClientID 
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Text            =   "ExampleClientSub"
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox txtAddress 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Text            =   "tcp://mqtt.eclipse.org:1883"
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "密码："
      Height          =   180
      Left            =   7800
      TabIndex        =   15
      Top             =   840
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户名："
      Height          =   180
      Left            =   5160
      TabIndex        =   13
      Top             =   840
      Width           =   720
   End
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      Caption         =   "消息发布："
      Height          =   180
      Left            =   240
      TabIndex        =   8
      Top             =   4800
      Width           =   900
   End
   Begin VB.Label lblReceivedMsgs 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "接收消息："
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "主题："
      Height          =   180
      Left            =   7680
      TabIndex        =   4
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户端ID:"
      Height          =   180
      Left            =   3960
      TabIndex        =   2
      Top             =   240
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "地址："
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   540
   End
End
Attribute VB_Name = "frmMQTT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MQTTClient As Long
Dim MQTTStatus As Integer '0-未开 1-开启

Private Sub cmdClose_Click()
    Dim StrErr As String
    
    cmdClose.Enabled = False
    Timer1.Interval = 0
    Timer1.Enabled = False
        
    MQTT_Close MQTTClient, StrErr
    
    MQTTStatus = 0
    
    If Len(StrErr) > 0 Then
        MsgBox "关闭成功！" & StrErr
    Else
        MsgBox "关闭成功！"
    End If

    FormEnabled False
    cmdOpen.Enabled = True
End Sub

Private Sub cmdOpen_Click()
    Dim StrErr As String
    
    cmdOpen.Enabled = False
    txtReceivedMsgs.Text = ""
    
    MQTTStatus = MQTT_Open(MQTTClient, txtAddress.Text, txtTopic.Text, txtClientID.Text, txtUserName.Text, txtPassWord.Text, , StrErr)
    If MQTTStatus = 0 Then
        MsgBox StrErr
        cmdOpen.Enabled = True
        Exit Sub
    End If

    Timer1.Interval = 500
    Timer1.Enabled = True
    
    FormEnabled True
    
    cmdClose.Enabled = True
End Sub

Private Sub Form_Load()
    FormEnabled False
    
    cmdClose.Enabled = False
End Sub

Sub FormEnabled(tEnabled As Boolean)
    lblReceivedMsgs.Enabled = tEnabled
    txtReceivedMsgs.Enabled = tEnabled
    lblMsg.Enabled = tEnabled
    txtMsg.Enabled = tEnabled
    txtPub.Enabled = tEnabled
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim StrErr As String
    
    If MQTTStatus = 1 Then
        Timer1.Interval = 0
        Timer1.Enabled = False
        MQTT_Close MQTTClient, StrErr
        MQTTStatus = 0
    End If
End Sub

Private Sub Timer1_Timer()
    Dim StrMsg As String
    Dim TestInfo As String
    
    Timer1.Enabled = False
        
    StrMsg = MQTT_GetNewMsg(MQTTClient)
    If Len(StrMsg) > 0 Then
        With txtReceivedMsgs
            .Text = .Text & Now & vbTab & StrMsg & vbCrLf
        End With
    End If
    
    Timer1.Enabled = True
End Sub

Private Sub txtPub_Click()
    Dim StrErr As String
    
    txtPub.Enabled = False
    
    If MQTT_PubMessage(MQTTClient, txtMsg.Text, , , StrErr) = 0 Then
        MsgBox StrErr
        txtPub.Enabled = True
        Exit Sub
    End If

    txtMsg.Text = ""
    txtPub.Enabled = True
End Sub
