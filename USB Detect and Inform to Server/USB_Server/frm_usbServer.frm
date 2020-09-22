VERSION 5.00
Object = "{24D4D63A-384B-4D2C-87E4-3FB10F395BDC}#41.0#0"; "SIContToWorld.ocx"
Begin VB.Form frm_usbServer 
   BackColor       =   &H80000004&
   Caption         =   "Form1"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin SIContToWorld.ContSI SI 
      Left            =   1320
      Top             =   1080
      _ExtentX        =   1402
      _ExtentY        =   1296
   End
End
Attribute VB_Name = "frm_usbServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
SI.CloseSI
SI.LocalPort = 5666
SI.Listen
End Sub

Private Sub SI_CloseSI()
MsgBox "Connection with the Client is closed."
SI.CloseSI
SI.Listen
End Sub
Private Sub SI_ConnectionRequest(ByVal requestID As Long)
SI.CloseSI
SI.Accept requestID
End Sub

Private Sub SI_DataArrival(ByVal bytesTotal As Long)
Dim sd As String
SI.GetData sd
MsgBox sd
frmDecision.Show
End Sub

Private Sub SI_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox 1, vbCritical, Error
End Sub


