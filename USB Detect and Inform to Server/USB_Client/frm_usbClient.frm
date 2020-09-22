VERSION 5.00
Object = "{24D4D63A-384B-4D2C-87E4-3FB10F395BDC}#41.0#0"; "SIContToWorld.ocx"
Begin VB.Form frm_usbCilient 
   Caption         =   "Form1"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin SIContToWorld.ContSI SI 
      Left            =   2520
      Top             =   1200
      _ExtentX        =   1402
      _ExtentY        =   1296
   End
   Begin VB.CommandButton cmdUn_Access 
      Caption         =   "To Proceed Further Click Here"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Timer Time 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   1320
   End
End
Attribute VB_Name = "frm_usbCilient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New Usb_Detect
Private Sub Form_Load()
SI.CloseSI
SI.Connect "200.100.40.160", 5666
a = fso.Usb_Detected
If a Then
Time.Enabled = True
Else
MsgBox "Usb Device Not Connected"
End
End If
End Sub
Private Sub cmdUn_Access_Click()
SI.SendData "Unauthorised Access to the Usb port...Allow or Deney the permission"
End Sub
Private Sub SI_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
SI.GetData Data
If Data = "Allowed" Then
frmBlank.Hide
Me.Hide
MsgBox Data
ElseIf Data = "Denyed" Then
frmBlank.Show
MsgBox Data & "Please contact your Admin..."
End If
End Sub

Private Sub Time_Timer()
cmdUn_Access.Value = True
MsgBox "You need permissin for connecting External devices...Please contact Your Network Admin..."
Time.Enabled = False
frmMain.cmdAccept.Value = True
End Sub
