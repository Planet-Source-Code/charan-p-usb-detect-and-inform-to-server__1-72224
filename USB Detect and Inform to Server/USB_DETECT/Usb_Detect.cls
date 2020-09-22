VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Usb_Detect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Dim fox(13)
Dim c, i As Byte

Public Function Usb_Detected() As Boolean
For i = 68 To 80
c = c + 1
fox(c) = Chr(i) & ":"
Next
For c = 1 To 13
If GetDriveType(fox(c)) = 2 Then
usb_drv = fox(c)
Usb_Detected = True
Exit Function
Else
Usb_Detected = False
usb_drv = ""
End If
Next
End Function
