Attribute VB_Name = "mVB6Plus"
Option Explicit

'---------字符处理函数-----------
Public Declare Function StrCompare Lib "VB6Plus.dll" (ByRef Str_A As String, ByRef Str_B As String) As Double
Public Declare Function Permutation Lib "VB6Plus.dll" (ByRef IN_STR As String, Optional ByRef Separator As String = ",", Optional ByRef ResultTotal As Long = 0) As String
Public Declare Function Combination Lib "VB6Plus.dll" (ByRef IN_STR As String, Optional ByRef Separator As String = ",", Optional ByRef ResultTotal As Long = 0) As String
Public Declare Function StrToHex_GB Lib "VB6Plus.dll" (ByRef IN_STR As String, Optional ByVal IsUpper As Integer = 1) As String
Public Declare Function StrToHex_UTF8 Lib "VB6Plus.dll" (ByRef IN_STR As String, Optional ByVal IsUpper As Integer = 1) As String
Public Declare Function HexToStr_GB Lib "VB6Plus.dll" (ByRef IN_STR As String) As String
Public Declare Function HexToStr_UTF8 Lib "VB6Plus.dll" (ByRef IN_STR As String) As String

'---------文件操作函数-----------
Public Declare Function ReadINIValue Lib "VB6Plus.dll" (ByRef SectionName As String, ByRef KeyName As String, Optional ByRef DefaultValue As String = "", Optional ByRef INIFile As String = "Config.ini") As String
Public Declare Function WriteINIValue Lib "VB6Plus.dll" (ByRef SectionName As String, ByRef KeyName As String, ByRef Value As String, Optional ByRef INIFile As String = "Config.ini") As Boolean

'---------HTML函数-----------
Public Declare Function UrlEncode_GB Lib "VB6Plus.dll" (ByVal IN_STR As String) As String
Public Declare Function UrlDecode_GB Lib "VB6Plus.dll" (ByVal IN_STR As String) As String

Public Declare Function UrlEncode_UTF8 Lib "VB6Plus.dll" (ByVal IN_STR As String) As String
Public Declare Function UrlDecode_UTF8 Lib "VB6Plus.dll" (ByVal IN_STR As String) As String

Public Declare Function UnicodeEncode Lib "VB6Plus.dll" (ByVal IN_STR As String) As String
Public Declare Function UnicodeDecode Lib "VB6Plus.dll" (ByVal IN_STR As String) As String

Public Declare Function HTMLEncode Lib "VB6Plus.dll" (ByVal IN_STR As String) As String
Public Declare Function HTMLDecode Lib "VB6Plus.dll" (ByVal IN_STR As String) As String

'---------加密函数-----------
Public Declare Function Base64Encode_GB Lib "VB6Plus.dll" (ByVal IN_STR As String) As String
Public Declare Function Base64Decode_GB Lib "VB6Plus.dll" (ByVal IN_STR As String) As String

Public Declare Function Base64Encode_UTF8 Lib "VB6Plus.dll" (ByVal IN_STR As String) As String
Public Declare Function Base64Decode_UTF8 Lib "VB6Plus.dll" (ByVal IN_STR As String) As String

Public Declare Function MD516_GB Lib "VB6Plus.dll" (ByVal IN_STR As String) As String
Public Declare Function MD532_GB Lib "VB6Plus.dll" (ByVal IN_STR As String) As String

Public Declare Function MD516_UTF8 Lib "VB6Plus.dll" (ByVal IN_STR As String) As String
Public Declare Function MD532_UTF8 Lib "VB6Plus.dll" (ByVal IN_STR As String) As String

Enum AESMode
    ECB = 0
    CBC = 1
    CFB = 2
End Enum
Enum AESCodeType
    Base64 = 0
    Hex = 1
End Enum
Public Declare Function AESEncrypt_GB Lib "VB6Plus.dll" (ByVal IN_STR As String, Optional ByVal IN_PWD As String = "", Optional ByVal IN_IV As String = "gfdertfghjkuyrtg", Optional ByVal iMode As AESMode = AESMode.ECB, Optional ByVal Padding As Integer = 0, Optional ByVal OutType As AESCodeType = AESCodeType.Base64) As String
Public Declare Function AESDecrypt_GB Lib "VB6Plus.dll" (ByVal IN_STR As String, Optional ByVal IN_PWD As String = "", Optional ByVal IN_IV As String = "gfdertfghjkuyrtg", Optional ByVal iMode As AESMode = AESMode.ECB, Optional ByVal Padding As Integer = 0, Optional ByVal InType As AESCodeType = AESCodeType.Base64) As String
Public Declare Function AESEncrypt_UTF8 Lib "VB6Plus.dll" (ByVal IN_STR As String, Optional ByVal IN_PWD As String = "", Optional ByVal IN_IV As String = "gfdertfghjkuyrtg", Optional ByVal iMode As AESMode = AESMode.ECB, Optional ByVal Padding As Integer = 0, Optional ByVal OutType As AESCodeType = AESCodeType.Base64) As String
Public Declare Function AESDecrypt_UTF8 Lib "VB6Plus.dll" (ByVal IN_STR As String, Optional ByVal IN_PWD As String = "", Optional ByVal IN_IV As String = "gfdertfghjkuyrtg", Optional ByVal iMode As AESMode = AESMode.ECB, Optional ByVal Padding As Integer = 0, Optional ByVal InType As AESCodeType = AESCodeType.Base64) As String

Public Declare Function AESEncryptTxtFile_GB Lib "VB6Plus.dll" (ByVal IN_File As String, ByVal OUT_File As String, Optional ByVal IN_PWD As String = "", Optional ByVal IN_IV As String = "gfdertfghjkuyrtg", Optional ByVal iMode As AESMode = AESMode.ECB) As String
Public Declare Function AESDecryptTxtFile_GB Lib "VB6Plus.dll" (ByVal IN_File As String, ByVal OUT_File As String, Optional ByVal IN_PWD As String = "", Optional ByVal IN_IV As String = "gfdertfghjkuyrtg", Optional ByVal iMode As AESMode = AESMode.ECB) As String

'---------网络函数-----------
Enum XMLType
    Microsoft_XMLHTTP = 0
    MSXML2_ServerXMLHTTP
    Msxml2_XMLHTTP_6_0
End Enum
Public Declare Function XMLHTTP_Get Lib "VB6Plus.dll" (ByRef URL As String, Optional ByRef RequestHeaders As String = "", Optional ByRef ResponseHeaders As String = "", Optional ByVal IsUTF8 As Integer = 1, Optional ByVal XMLType As XMLType = XMLType.Microsoft_XMLHTTP) As String
Public Declare Function XMLHTTP_Post Lib "VB6Plus.dll" (ByRef URL As String, ByRef PostDatas As String, Optional ByRef RequestHeaders As String = "Content-Type:application/x-www-form-urlencoded", Optional ByRef ResponseHeaders As String = "", Optional ByVal IsUTF8 As Integer = 1, Optional ByVal XMLType As XMLType = XMLType.Microsoft_XMLHTTP) As String

'---------Windows函数-----------
Public Declare Function Win_CopyFileToClipBoard Lib "VB6Plus.dll" (ByVal IN_FileOrDir As String) As Long
Public Declare Function RunVBScript Lib "VB6Plus.dll" (ByRef VBScript As String, ByRef Error As String) As Long
Public Declare Function RunVBFunction Lib "VB6Plus.dll" (ByRef VBScript As String, ByRef RunFuncName As String, ByRef Params() As Variant, ByRef Result As Variant, ByRef Error As String) As Long
Public Declare Sub SetFormIcon Lib "VB6Plus.dll" (ByRef hwnd As Long, ByRef SkinIcoName As String, Optional ByRef SkinPath As String = "Default")

'---------对话框函数-----------
Public Declare Function ShowOpenFile Lib "VB6Plus.dll" (Optional ByVal m_hWnd As Long = 0, Optional ByRef StrFilter As String = "全部|*.*|文本文件|*.TXT|图片文件|*.BMP;*.PNG;*.JPG|", Optional ByRef StrInitDir As String = "", Optional ByRef StrTitle As String = "打开", Optional ByRef StrFileJoinStr As String = vbCrLf, Optional ByRef MultiSel As Integer = 0) As String
Public Declare Function ShowSaveFile Lib "VB6Plus.dll" (Optional ByVal m_hWnd As Long = 0, Optional ByRef StrFilter As String = "TXT文件|*.TXT|LOG文件|*.LOG|", Optional ByRef StrInitDir As String = "", Optional ByRef StrTitle As String = "另存为", Optional ByRef StrDefExt As String = "TXT") As String
Public Declare Function ShowBrowserFolder Lib "VB6Plus.dll" (Optional ByVal m_hWnd As Long = 0, Optional ByRef StrInitDir As String = "", Optional ByRef StrTitle As String = "打开") As String

'---------图像函数-----------
Enum QRecLevel
    QR_ECLEVEL_L = 0
    QR_ECLEVEL_M
    QR_ECLEVEL_Q
    QR_ECLEVEL_H
End Enum
Public Declare Function MakeQRCode Lib "VB6Plus.dll" (ByRef QRText As String, ByRef IMGFilePath As String, Optional ByVal Size As Integer = 0, Optional ByVal ErrRateLevel As QRecLevel = QRecLevel.QR_ECLEVEL_H, Optional ByVal Quality As Integer = 100) As String
Public Declare Function ScanQRImage Lib "VB6Plus.dll" (ByRef IMGFilePath As String, Optional ByVal hybrid As Boolean = False, Optional ByRef ErrText As String = "", Optional ByVal QRTextIsUTF8 As Integer = 0) As String
Public Declare Function ImageToJPG Lib "VB6Plus.dll" (ByRef IMGFilePath As String, ByRef JPGFilePath As String, Optional ByVal Quality As Integer = 95) As String
Public Declare Function ImageToBMP Lib "VB6Plus.dll" (ByRef IMGFilePath As String, ByRef BMPFilePath As String) As String

'---------数据库函数-----------
Public Declare Function SQLite_Open Lib "VB6Plus.dll" (ByRef SQLiteDB As Long, Optional ByRef DBFileName As String = "DB.DB", Optional ByRef StrErr As String = "") As Long
Public Declare Function SQLite_Close Lib "VB6Plus.dll" (ByRef SQLiteDB As Long) As Long
Public Declare Function SQLite_ReadData Lib "VB6Plus.dll" (ByRef SQLiteDB As Long, ByRef QuerySQL As String, ByRef Data() As String, Optional ByRef StrErr As String = "") As Long
Public Declare Function SQLite_Execute Lib "VB6Plus.dll" (ByRef SQLiteDB As Long, ByRef ExeSQL As String, Optional ByRef StrErr As String = "") As Long

'---------多线程函数-----------
Public Declare Function Net_MT Lib "VB6Plus.dll" (ByRef Data() As String) As Long
Public Declare Function RunVBFunction_MT Lib "VB6Plus.dll" (ByRef Data() As String, ByRef FuncParas() As Variant, ByRef RunResult() As Variant) As Long

