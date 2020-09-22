VERSION 5.00
Object = "*\APOPDBExchge.vbp"
Begin VB.Form FrmSMTPRelay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SMTP Relay / Direct Connection Model"
   ClientHeight    =   7710
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   10110
   Icon            =   "FrmSMTPRelay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin POPDB.DBSMTPSNDER DBSMTPSNDER1 
      Height          =   7695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   13573
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   7560
      Top             =   4920
   End
   Begin VB.Menu mnuServices 
      Caption         =   "&Services"
      Begin VB.Menu mnuSmtpStart 
         Caption         =   "Start SMTP Connector"
      End
      Begin VB.Menu mnuSmtpStop 
         Caption         =   "S&top SMTP Connector"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuunload 
         Caption         =   "Unload SMTP Connector"
      End
   End
   Begin VB.Menu MnuView1 
      Caption         =   "&View"
      Begin VB.Menu MnuView 
         Caption         =   "View Message"
      End
      Begin VB.Menu mnuSMTPLOG 
         Caption         =   "View SMTP Logging"
      End
   End
   Begin VB.Menu mnuMail 
      Caption         =   "&Mail"
      Begin VB.Menu mnurefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnusep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMailDelete 
         Caption         =   "Purge Message"
      End
      Begin VB.Menu mnuPurgeALL 
         Caption         =   "Purge All Messages"
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReprocess 
         Caption         =   "Retry Bad Queue"
      End
   End
   Begin VB.Menu mnuAdvanced 
      Caption         =   "&Advanced"
      Begin VB.Menu mnuDirect 
         Caption         =   "Direct Connection SMTP"
      End
      Begin VB.Menu mnuSMTP 
         Caption         =   "SMTP Relay"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuWebsite 
         Caption         =   "Visit Author's Website"
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpDirect 
         Caption         =   "About SMTP Connections"
      End
   End
End
Attribute VB_Name = "FrmSMTPRelay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public DBPATH As String




Private Sub Form_Load()


DBSMTPSNDER1.DNS = FrmMain.DNS 'set dns to hatton.net
DBSMTPSNDER1.DBPATH = DBPATH
DBSMTPSNDER1.ADO_Connect
DBSMTPSNDER1.LoadTreeView
DBSMTPSNDER1.LoadTreeviewBADQUE

If DBSMTPSNDER1.smtprelay = False Then mnuDirect.Checked = True Else mnuDirect.Checked = False
If DBSMTPSNDER1.smtprelay = True Then mnuSMTP.Checked = True Else mnuSMTP.Checked = False
Me.WindowState = vbMinimized
Me.Caption = Me.Caption & ";  Domain Name (" & DNS & ")"
End Sub

Private Sub mnuDirect_Click()
    
If mnuDirect.Checked = True Then 'turn off direct goto relay
    DBSMTPSNDER1.smtprelay = True 'relay
    mnuSMTP.Checked = True
    mnuDirect.Checked = False
    Call FrmOptions.ADO_Connect
    Call FrmOptions.Save(True)
    Call FrmMain.UpdateGlobalChanges
    Exit Sub
ElseIf mnuDirect.Checked = False Then 'goto direct
    DBSMTPSNDER1.smtprelay = False 'direct
    mnuSMTP.Checked = False
    mnuDirect.Checked = True
    Call FrmOptions.ADO_Connect
    Call FrmOptions.Save(False)
    Call FrmMain.UpdateGlobalChanges

    Exit Sub
End If
End Sub

Private Sub mnuHelpDirect_Click()
    Dim sTopic As String
    Dim sFile As String
    Dim sParams As String
    Dim sDirectory As String
    sTopic = "Open"
    sFile = "www.chris.hatton.com/Whitepapers/MailServer.htm"
    sParams = 0&
    sDirectory = 0&
    
    RunShellExecute sTopic, sFile, sParams, sDirectory, SW_SHOWNORMAL

End Sub

Private Sub mnuMailDelete_Click()
DBSMTPSNDER1.mnupurge_Click
End Sub

Private Sub mnuPurgeALL_Click()
DBSMTPSNDER1.mnuPurgeALL_Click
End Sub

Private Sub mnuRefresh_Click()
DBSMTPSNDER1.Refresh
End Sub

Private Sub mnuReprocess_Click()
DBSMTPSNDER1.ProcessBAD
End Sub

Private Sub mnuSMTP_Click()
If mnuSMTP.Checked = True Then 'turn off relay
    DBSMTPSNDER1.smtprelay = False 'turn on direct
    Call FrmOptions.ADO_Connect
    Call FrmOptions.Save(False)
    Call FrmMain.UpdateGlobalChanges
    mnuSMTP.Checked = False
    mnuDirect.Checked = True 'show direct
    Exit Sub
ElseIf mnuSMTP.Checked = False Then 'turn off direct
    DBSMTPSNDER1.smtprelay = True 'turn on relay
    Call FrmOptions.ADO_Connect
    Call FrmOptions.Save(True)
    Call FrmMain.UpdateGlobalChanges
    mnuSMTP.Checked = True 'show relay
    mnuDirect.Checked = False
    
    Exit Sub
End If
End Sub

Private Sub mnuSMTPLOG_Click()
Shell "notepad.exe " & Left$(App.Path, Len(App.Path) - 6) & "SMTPRELAY.LOG", vbNormalFocus
End Sub
Public Sub StartServices()
    DBSMTPSNDER1.StartTimer
End Sub

Private Sub mnuSmtpStart_Click()
    DBSMTPSNDER1.StartTimer
End Sub

Public Sub mnuSmtpStop_Click()
    DBSMTPSNDER1.StopTimer
End Sub
Public Sub PropertiesUpdate() 'reread these variables
    
    DBSMTPSNDER1.ReadOptions
    If DBSMTPSNDER1.smtprelay = True Then mnuSMTP.Checked = True: mnuDirect.Checked = False: Exit Sub
    If DBSMTPSNDER1.smtprelay = False Then mnuSMTP.Checked = False: mnuDirect.Checked = True: Exit Sub
End Sub
Public Sub mnuunload_Click()
DBSMTPSNDER1.StopTimer
Unload Me
End Sub

Private Sub mnuView_Click()
DBSMTPSNDER1.ViewMessage
End Sub

Private Sub mnuWebsite_Click()
    Dim sTopic As String
    Dim sFile As String
    Dim sParams As String
    Dim sDirectory As String
    sTopic = "Open"
    sFile = "www.chris.hatton.com"
    sParams = 0&
    sDirectory = 0&
    
    RunShellExecute sTopic, sFile, sParams, sDirectory, SW_SHOWNORMAL

End Sub

Private Sub Timer1_Timer()
        If DBSMTPSNDER1.SMTPCon = False Then
            Me.Caption = "SMTP Relay / Direct Connection Model" & "     (SMTP Services are STOPPED)"
            Else
            Me.Caption = "SMTP Relay / Direct Connection Model" & "     (SMTP Services are in Progress)"
        End If
End Sub
