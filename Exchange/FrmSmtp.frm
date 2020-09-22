VERSION 5.00
Object = "*\APOPDBExchge.vbp"
Begin VB.Form FrmSmtp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SMTP Reciever"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11415
   Icon            =   "FrmSmtp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   11415
   StartUpPosition =   3  'Windows Default
   Begin POPDB.SMTP SMTP1 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   12515
   End
End
Attribute VB_Name = "FrmSmtp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.Caption = "SMTP Receiver v1.1 " & "Hattech Technologies " & "Domain Name (" & FrmMain.DNS & ")"
SMTP1.DNS = FrmMain.DNS
SMTP1.DBPATH = FrmMain.DBPATH
SMTP1.Start_SMTPListener
Me.WindowState = vbMinimized
End Sub
Public Sub Stop_SMTP_Lister()
On Error Resume Next
Call SMTP1.Stop_SMTPLister
End Sub

