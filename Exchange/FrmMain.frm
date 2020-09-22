VERSION 5.00
Object = "*\APOPDBExchge.vbp"
Begin VB.Form FrmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "POP3 Receiver (www.chris.hatton.com)"
   ClientHeight    =   7410
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11205
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin POPDB.Exchange Exchange1 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   12938
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
      Begin VB.Menu mnustartRelay 
         Caption         =   "Start SMTP Relay Service"
      End
      Begin VB.Menu mnuStopRelay 
         Caption         =   "Stop SMTP Relay Service"
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuDomain 
         Caption         =   "Domain Name"
      End
   End
   Begin VB.Menu MnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuMinimize 
         Caption         =   "Minimize SMTP Windows"
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
   Begin VB.Menu mnuAbout 
      Caption         =   "&Help"
      Begin VB.Menu mnuMailADO 
         Caption         =   "About POP3 with ADO"
      End
      Begin VB.Menu mnuPOP3 
         Caption         =   "Pop3 Structures (RFC 1939)"
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
On Error Resume Next
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
On Error Resume Next
    FrmMain.WindowState = vbMinimized
    ShutDown ("Unloading Online Database store....")
    FrmShutdown.Show

DoEvents: DoEvents: DoEvents: DoEvents
    Exchange1.Stop_OnlineStore
    FrmSMTPRelay.mnuSmtpStop_Click
DoEvents: DoEvents: DoEvents: DoEvents
    ShutDown ("Stopping SMTP Connector.......")
DoEvents: DoEvents: DoEvents: DoEvents
    FrmSmtp.Stop_SMTP_Lister
    Unload FrmSmtp
    ShutDown ("Unloading SMTP Relay.......")
    
    FrmSMTPRelay.mnuunload_Click
    Unload FrmSMTPRelay
    Unload FrmShutdown
    Unload Me
    End
DoEvents: DoEvents: DoEvents: DoEvents
    
    
End Sub
Private Sub ShutDown(Status As String)
      FrmShutdown.Label3.Caption = Status


End Sub
Private Sub mnuCompact_Click()
On Error Resume Next
    Exchange1.Stop_OnlineStore 'stop to compact database
    Unload FrmSMTPRelay
    Unload FrmSmtp
    Exchange1.CompactJetDatabase DBPATH, True 'compact it
    Exchange1.Start_OnlineStore 'start the database
    FrmSMTPRelay.Show
    FrmSmtp.Show
    FrmSMTPRelay.StartServices
    Exchange1.AccStatus 'get mail box status's
    Exchange1.SMTPQue 'get the external stats'
End Sub
Private Sub LoadDomain()
On Error Resume Next
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

Private Sub mnuDomain_Click()
Dim domain As String
    domain = InputBox("Enter in your new domain name", "Domain Name")
    If Len(domain) = 0 Then
        MsgBox "No Changes have been made", vbInformation
    Else
        Call EditDomain(domain)
        MsgBox "You Must Restart Mail Server for the changes to take Effect", vbInformation

    End If
       
End Sub
Private Sub EditDomain(EDNS As String)
On Error Resume Next
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
    cn.Provider = "Microsoft.Jet.OLEDB.4.0"
    cn.Open DBPATH
    rs.Open "select * from settings", cn, adOpenDynamic, adLockOptimistic
        rs!DomainName = EDNS
    rs.Update
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

Private Sub mnuMailADO_Click()
frmAbout.Show
End Sub

Public Sub mnuMailbox_Click()
On Error Resume Next
    Exchange1.AccStatus 'get mail box status's
    Exchange1.SMTPQue 'get the external stats'
End Sub

Private Sub mnuMinimize_Click()
FrmSmtp.WindowState = vbMinimized
FrmSMTPRelay.WindowState = vbMinimized

End Sub

Private Sub mnuoptions_Click()
On Error Resume Next
    FrmOptions.Show 1
    
End Sub
Public Sub UpdateGlobalChanges() 'updates all changes to properties to all forms
On Error Resume Next
    FrmSMTPRelay.PropertiesUpdate 'updates the relay and the direct connection variables
   
End Sub

Private Sub mnuPOP3_Click()
    Dim sTopic As String
    Dim sFile As String
    Dim sParams As String
    Dim sDirectory As String
    sTopic = "Open"
    sFile = "http://www.faqs.org/rfcs/rfc1939.html"
    sParams = 0&
    sDirectory = 0&
    
    RunShellExecute sTopic, sFile, sParams, sDirectory, SW_SHOWNORMAL

'
End Sub

Private Sub mnuRepair_Click()
Exchange1.RepairDB
MsgBox "Finished Repair", vbInformation

End Sub

Private Sub mnuStartPOP3_Click()
On Error Resume Next
    Exchange1.Start_OnlineStore
End Sub

Private Sub mnustartRelay_Click()
FrmSMTPRelay.StartServices
End Sub

Private Sub mnuStopPop3_Click()
On Error Resume Next
    Exchange1.Stop_OnlineStore
End Sub

Private Sub mnuStopRelay_Click()
FrmSMTPRelay.mnuSmtpStop_Click

End Sub

Private Sub mnuSync_Click()
On Error Resume Next
    Exchange1.ProcessPOPMail
End Sub
