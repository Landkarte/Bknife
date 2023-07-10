Attribute VB_Name = "mVB6OpenSSL"
Option Explicit

'---------ÍøÂçº¯Êý-----------
Public Declare Function OpenSSL_Get Lib "VB6OpenSSL.dll" (ByRef URL As String, Optional ByRef RequestHeaders As String = "", Optional ByRef ResponseHeaders As String = "", Optional ByVal HTTPVersion As Double = 1#, Optional ByVal IsUTF8 As Integer = 1, Optional ByVal TimeOut As Long = 10) As String
Public Declare Function OpenSSL_Post Lib "VB6OpenSSL.dll" (ByRef URL As String, ByRef PostDatas As String, Optional ByRef RequestHeaders As String = "content-type:application/x-www-form-urlencoded", Optional ByRef ResponseHeaders As String = "", Optional ByVal HTTPVersion As Double = 1#, Optional ByVal IsUTF8 As Integer = 1, Optional ByVal TimeOut As Long = 10) As String

