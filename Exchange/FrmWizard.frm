VERSION 5.00
Begin VB.Form FrmWizard 
   Caption         =   "New Email Account"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6600
      TabIndex        =   15
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmdfinish 
      Caption         =   "< Finish >"
      Height          =   495
      Left            =   8520
      TabIndex        =   14
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   10215
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   6240
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   3480
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   6240
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   2880
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   6240
         PasswordChar    =   "*"
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   2280
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   6240
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   6240
         PasswordChar    =   "*"
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   6240
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   600
         Width           =   3615
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00800000&
         Height          =   255
         Index           =   5
         Left            =   6960
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   7080
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   6840
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   5760
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   7800
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   6360
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   240
         TabIndex        =   12
         Top             =   3360
         Width           =   5655
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   240
         TabIndex        =   11
         Top             =   2760
         Width           =   5655
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   2160
         Width           =   5655
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   5655
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   5655
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   4185
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   10170
      End
   End
End
Attribute VB_Name = "FrmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection

Private Sub Command1_Click()

End Sub

Private Sub cmdfinish_Click()
If checkaccount = True Then
    If InStr(1, Text1(0), "@") > 1 Then
        Text1(0) = Split(Text1(0), "@")(0)
    End If
    
    Call CreateNewUser(Text1(0), Text1(1), Text1(4), Text1(5), Text1(2), Text1(3))
    
    FrmProfiles.List_Profiles
    Unload Me
    
    Else
    
    MsgBox "Enter all details", vbCritical


End If
End Sub
Private Sub ADO_Connect()
Set cn = New ADODB.Connection
        cn.Provider = "Microsoft.Jet.OLEDB.4.0"
        cn.Open FrmMain.DBPATH
        cn.CursorLocation = adUseClient

End Sub
Private Sub ADO_Close()
On Error Resume Next
    cn.Close
    Set cn = Nothing
End Sub
Private Sub CreateNewUser(strUser, strPass, strExtPOP, strExtSmtp, strExtUser, strExtPass As String)
On Error GoTo HelpMe
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
    rs.Open "select * from profiles", cn, adOpenKeyset, adLockOptimistic
    rs.AddNew
    rs!user = Trim$(strUser)
    rs!PASS = Trim$(strPass)
    rs!ExternalPOP3 = Trim$(strExtPOP)
    rs!ExternalSMTP = Trim$(strExtSmtp)
    rs!Externaluser = Trim$(strExtUser)
    rs!ExternalPass = Trim$(strExtPass)
    rs.Update
    rs.Close
Set rs = Nothing
MsgBox "Email Profile Successfully Completed", vbInformation
Exit Sub

HelpMe:
On Error Resume Next
rs.Close
Set rs = Nothing

MsgBox "Could not Create Profile" & vbNewLine & Err.Description, vbCritical



End Sub
Private Function checkaccount() As Boolean

Dim i As Integer

For i = 1 To 5
    If Len(Text1(i)) = 0 Then
        checkaccount = False
        Exit Function
    End If

Next i

checkaccount = True

End Function

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

Call ADO_Connect
Label1 = ""


        Label2(0) = "  What is your Local (LAN) Email address going to be?"
        Label2(1) = "  Specify a Password for the local Email Account"
        Label2(2) = "  Give me any External email MailBox Login name?"
        Label2(3) = "  And the MailBox Password that matches the External email Login?"
        Label2(4) = "  What is the External ISP POP3 Internet Server?"
        Label2(5) = "  What is your External ISP SMTP Internet Server?"
        

        Label2(0).ToolTipText = "This can be anything you like. eg. User@MyDomain.com"
        Label2(1).ToolTipText = "This is what you would set this in your outlook express or any other email application."
        Label2(2).ToolTipText = "This can be an external email account from any domain that is trusted by your ISP"
        Label2(3).ToolTipText = "This is the password for your external email account."
        Label2(4).ToolTipText = "Your ISP provides this info, but generally is POP3.YourISP.com"
        Label2(5).ToolTipText = "Your ISP provides this info, but generally is SMTP.YourISP.com"
        
        Text1(0).ToolTipText = "joebloggs@microsoft.com"
        Text1(1).ToolTipText = "Password for joebloggs"
        Text1(2).ToolTipText = "ISP Logon Name, 'joebloggs' @yourisp.com.. This is where the mail server collects its mail for joebloggs"
        Text1(3).ToolTipText = "External Password for your ISP Email account"
        Text1(4).ToolTipText = "POP3.YourExternalPOP3Server.com"
        Text1(5).ToolTipText = "SMTP.YourExternalSmtpServer.com"
        
Dim i As Integer

For i = 0 To 5
   Shape1(i).Left = Label2(i).Left - 10
   Shape1(i).Top = Label2(i).Top - 10
   Shape1(i).Width = Label2(i).Width + 30
   Shape1(i).Height = Label2(i).Height + 10
   Shape1(i).Visible = False
   Text1(i).Text = ""
        
Next i







End Sub

Private Sub Form_Unload(Cancel As Integer)
Call ADO_Close
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Dim i As Integer

    For i = 0 To 5
        Label2(i).BackColor = vbWhite
        Shape1(i).Visible = False
    Next i

Label2(Index).BackColor = &HFFC0C0
Shape1(Index).Visible = True

End Sub
