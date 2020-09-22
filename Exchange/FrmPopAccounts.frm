VERSION 5.00
Begin VB.Form FrmPopAccounts 
   Caption         =   "POP Account Object"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   6765
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   5040
         TabIndex        =   4
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   2415
         Index           =   4
         Left            =   240
         TabIndex        =   6
         Top             =   2280
         Width           =   6135
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   6135
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   6135
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   6135
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   6135
      End
   End
End
Attribute VB_Name = "FrmPopAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Username, EXTUSER, SMTPACC, POP3ACC As String
Private DNS As String

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim pop3Address As String
DNS = "YourDomainName.com"
pop3Address = Split(POP3ACC, ".")(0)
POP3ACC = Split(POP3ACC, pop3Address & ".")(1)
Frame1.Caption = " Object: " & Username & " "
Label1(0).Caption = " " & Username & "'s  External Address:  <" & EXTUSER & "@" & POP3ACC & ">"
Label1(1).Caption = " " & Username & "'s  Internal LAN Address:  <" & Username & "@" & LCase(DNS) & ">"
Label1(2).Caption = " " & Username & "'s  Perferred SMTP Provider is:  " & SMTPACC
Label1(3).Caption = " " & Username & "'s  Perferred POP3 Provider is:  " & pop3Address & "." & POP3ACC


Label1(4).Caption = " Mail will be downloaded from  <" & pop3Address & "." & POP3ACC & ">  into the local mailbox" & _
" " & Username & " [" & LCase(DNS) & "] " & vbCrLf & vbCrLf & "Configure your local Email Client Account to your LAN Mail server." & vbCrLf & vbCrLf & _
"SMTP Provider: Your Local/LAN Server NetBIOS/DNS Name or IP Address" & vbCrLf & _
"POP3 Provider: Your Local/LAN Server NetBIOS/DNS Name or IP Address" & vbCrLf & _
"Account Name: " & Username & vbCrLf & _
"Password: " & "Your LAN Account password " & vbCrLf & vbCrLf & _
"Note# your Internal Password has to match your External Mail box at " & pop3Address & "." & POP3ACC



End Sub

