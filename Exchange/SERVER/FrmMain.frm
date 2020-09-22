VERSION 5.00
Object = "{249D2905-FB80-4B06-A1EF-C68AFD55E9B4}#28.0#0"; "POPDBExchge.ocx"
Begin VB.Form FrmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "POP3 Receiver (www.chris.hatton.com)"
   ClientHeight    =   7260
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11250
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   11250
   StartUpPosition =   2  'CenterScreen
   Begin POPDB.Exchange Exchange1 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   12726
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuStartPOP3 
         Caption         =   "Start POP3 Service"
      End
      Begin VB.Menu mnuStopPop3 
         Caption         =   "Stop POP3 Service"
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuSync 
         Caption         =   "&Syncronize Mail"
      End
      Begin VB.Menu mnuMailbox 
         Caption         =   "&Refresh MailBoxes"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRepair 
         Caption         =   "Detect and Repair"
      End
      Begin VB.Menu mnuCompact 
         Caption         =   "Optimize Mailboxes"
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEDITVIEWMail 
         Caption         =   "Edit/Add Mailboxes"
      End
      Begin VB.Menu mnuoptions 
         Caption         =   "&Options"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DNS As String
Public DBPATH As String

Private Sub Form_Load()
DBPATH = App.Path & "\DBSTORE.MDB"
Call LoadDomain
    FrmSMTPRelay.DBPATH = DBPATH
    Exchange1.DBPATH = DBPATH
    Exchange1.Start_OnlineStore
    Exchange1.DNS = DNS
    FrmSmtp.Show
    FrmSMTPRelay.Show
    FrmMain.Caption = "Mail Server v1." & App.Revision & " Hattech Technologies " & "Domain Name (" & DNS & ")"
    FrmSMTPRelay.StartServices
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmMain.WindowState = vbMinimized
    ShutDown ("Unloading Online Database store....")
    FrmShutdown.Show

DoEvents
    Exchange1.Stop_OnlineStore
    FrmSMTPRelay.mnuSmtpStop_Click
    ShutDown ("Stopping SMTP Connector.......")
    Unload FrmSmtp
    ShutDown ("Unloading SMTP Relay.......")
    Unload FrmSMTPRelay
    End
    
End Sub
Private Sub ShutDown(Status As String)
      FrmShutdown.Label3.Caption = Status


End Sub
Private Sub mnuCompact_Click()
    Exchange1.Stop_OnlineStore 'stop to compact database
    Exchange1.CompactJetDatabase DBPATH, True 'compact it
    Exchange1.Start_OnlineStore 'start the database
    Exchange1.AccStatus 'get mail box status's
    Exchange1.SMTPQue 'get the external stats'
End Sub
Private Sub LoadDomain()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
    cn.Provider = "Microsoft.Jet.OLEDB.4.0"
    cn.Open DBPATH
    rs.Open "select * from settings", cn, adOpenStatic, adLockReadOnly
        DNS = rs!DomainName
    rs.Close
Set rs = Nothing
Set cn = Nothing
End Sub
Private Sub mnuEDITVIEWMail_Click()
    FrmProfiles.Show
End Sub

Private Sub mnuExit_Click()
    Exchange1.Stop_OnlineStore
    Unload Me
End Sub

Public Sub mnuMailbox_Click()
    Exchange1.AccStatus 'get mail box status's
    Exchange1.SMTPQue 'get the external stats'
End Sub

Private Sub mnuoptions_Click()
    FrmOptions.Show 1
    
End Sub
Public Sub UpdateGlobalChanges() 'updates all changes to properties to all forms

    FrmSMTPRelay.PropertiesUpdate 'updates the relay and the direct connection variables
   
End Sub

Private Sub mnuRepair_Click()
Exchange1.RepairDB
MsgBox "Finished Repair", vbInformation

End Sub

Private Sub mnuStartPOP3_Click()
    Exchange1.Start_OnlineStore
End Sub

Private Sub mnuStopPop3_Click()
    Exchange1.Stop_OnlineStore
End Sub

Private Sub mnuSync_Click()
    Exchange1.ProcessPOPMail
End Sub
