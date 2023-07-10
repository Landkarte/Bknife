Attribute VB_Name = "VB6MQTT"
Option Explicit


'---------MQTT´¦Àíº¯Êý-----------
Public Declare Function MQTT_Open Lib "VB6MQTT.dll" (ByRef MQTTClient As Long, ByRef Address As String, ByRef Topic As String, ByRef ClientID As String, Optional ByRef UserName As String = "", Optional ByRef PassWord As String = "", Optional ByVal Qos As Integer = 1, Optional ByRef StrErr As String = "") As Long
Public Declare Function MQTT_PubMessage Lib "VB6MQTT.dll" (ByRef MQTTClient As Long, ByRef Message As String, Optional ByVal Qos As Integer = 1, Optional ByVal TimeOut As Long = 5000, Optional ByRef StrErr As String = "") As Long
Public Declare Function MQTT_GetNewMsg Lib "VB6MQTT.dll" (ByRef MQTTClient As Long) As String
Public Declare Function MQTT_Close Lib "VB6MQTT.dll" (ByRef MQTTClient As Long, Optional ByRef StrErr As String = "") As Long

