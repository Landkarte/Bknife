Attribute VB_Name = "Module1"
Option Explicit

'---------ÍøÂçº¯Êý-----------
Enum XMLType
    Microsoft_XMLHTTP = 0
    MSXML2_ServerXMLHTTP
    Msxml2_XMLHTTP_6_0
End Enum
Public Declare Function XMLHTTP_Get Lib "D:\Webshell (1)\VB6Plus.dll" (ByRef url As String, Optional ByRef RequestHeaders As String = "", Optional ByRef ResponseHeaders As String = "", Optional ByVal IsUTF8 As Integer = 1, Optional ByVal XMLType As XMLType = XMLType.Microsoft_XMLHTTP) As String
Public Declare Function XMLHTTP_Post Lib "D:\Webshell (1)\VB6Plus.dll" (ByRef url As String, ByRef PostDatas As String, Optional ByRef RequestHeaders As String = "Content-Type:application/x-www-form-urlencoded", Optional ByRef ResponseHeaders As String = "", Optional ByVal IsUTF8 As Integer = 1, Optional ByVal XMLType As XMLType = XMLType.Microsoft_XMLHTTP) As String

