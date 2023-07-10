Attribute VB_Name = "mData"
Option Explicit

'/* change the data structure to suit your application
'/* Download by http://www.codesc.net
'/* and rem the subitem1/subitem2 entries

Public Type HLISubItm
    lIcon       As Long
    Text()      As String
End Type

Public Type HLIStc
    Item()      As String
    lIcon()     As Long
    SubItem1()  As String
    SubItem2()  As String
    SubItem3()  As String
    SubItem4()  As String
    SubItem5()  As String
    SubItem6()  As String
    SubItem7()  As String
    SubItem8()  As String
    'SubItem()   As HLISubItm
End Type

Public Type HLIRtn
    RItem       As String
    RSubItem()  As String
End Type

