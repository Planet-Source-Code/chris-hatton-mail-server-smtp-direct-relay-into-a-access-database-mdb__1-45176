VERSION 5.00
Begin VB.Form FrmProfiles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create/Delete Mailbox Accounts"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmProfiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclose 
      Caption         =   "OK"
      Height          =   495
      Left            =   4560
      TabIndex        =   11
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdWizard 
      Caption         =   "Account Wizard"
      Height          =   495
      Left            =   6120
      TabIndex        =   10
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4540
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7930
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   7935
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "User Profiles"
            Height          =   255
            Left            =   0
            TabIndex        =   3
            Top             =   120
            Width           =   7935
         End
      End
      Begin VB.ListBox List1 
         Height          =   4155
         Left            =   0
         TabIndex        =   1
         Top             =   360
         Width           =   7935
      End
   End
   Begin VB.Label Label2 
      Height          =   255
      Index           =   6
      Left            =   3960
      TabIndex        =   12
      Top             =   5280
      Width           =   3735
   End
   Begin VB.Label Label2 
      Height          =   255
      Index           =   5
      Left            =   3960
      TabIndex        =   9
      Top             =   5280
      Width           =   3975
   End
   Begin VB.Label Label2 
      Height          =   255
      Index           =   4
      Left            =   3960
      TabIndex        =   8
      Top             =   5040
      Width           =   3975
   End
   Begin VB.Label Label2 
      Height          =   255
      Index           =   3
      Left            =   3960
      TabIndex        =   7
      Top             =   4800
      Width           =   3975
   End
   Begin VB.Label Label2 
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   5280
      Width           =   3735
   End
   Begin VB.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   5040
      Width           =   3735
   End
   Begin VB.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   4800
      Width           =   3735
   End
   Begin VB.Menu mnu 
      Caption         =   "mnusys"
      Visible         =   0   'False
      Begin VB.Menu mnuCreate 
         Caption         =   "&Create MailBox"
      End
      Begin VB.Menu munRemove 
         Caption         =   "R&emove Mailbox"
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnurename 
         Caption         =   "&Rename Mailbox"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Re&fresh"
      End
      Begin VB.Menu mnusetPass 
         Caption         =   "&Set Password"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEXTpop3 
         Caption         =   "Set External POP3 Server"
      End
      Begin VB.Menu mnuEXTsmtp 
         Caption         =   "Set External SMTP Server"
      End
      Begin VB.Menu mnuUserLogon 
         Caption         =   "Set External User Logon"
      End
      Begin VB.Menu mnuExtPass 
         Caption         =   "Set External Password"
      End
   End
End
Attribute VB_Name = "FrmProfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Private DNS As String


Private Sub cmdclose_Click()
Call FrmMain.mnuMailbox_Click 'refresh the list
Unload Me
End Sub

Private Sub cmdWizard_Click()
FrmWizard.Show 1
End Sub

Private Sub Form_Load()
    DNS = FrmMain.DNS
    Call ADO_Connect
    Call List_Profiles
    Me.Caption = Me.Caption & ";  Domain Name (" & DNS & ")"
End Sub
Public Sub List_Profiles()
    Dim rs As ADODB.Recordset
    Dim i As Long
    Set rs = New ADODB.Recordset
        rs.Open "select * from profiles", cn, adOpenKeyset, adLockOptimistic
            List1.Clear
            For i = 1 To rs.RecordCount
                List1.AddItem "User=" & rs!user & "#Pass=" & rs!PASS & "#ExtPOP3=" & rs!ExternalPOP3 & "#ExtSMTP=" & rs!ExternalSMTP & "#ExtUser=" & rs!Externaluser & "#ExtPass=" & rs!ExternalPass
                rs.MoveNext
            Next i
        rs.Close
    Set rs = Nothing
End Sub
Private Sub ADO_Connect()
Set cn = New ADODB.Connection
        cn.Provider = "Microsoft.Jet.OLEDB.4.0"
        cn.Open FrmMain.DBPATH
        cn.CursorLocation = adUseClient

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
cn.Close
Set cn = Nothing
End Sub

Private Sub List1_Click()
Dim i As Long
For i = 0 To Label2.Count - 1
    Label2(i) = ""
Next i

Label2(0).Caption = "Local " & Split(List1.List(List1.ListIndex), "#")(0) & "@" & DNS
Label2(1).Caption = "Local " & Split(List1.List(List1.ListIndex), "#")(1)
Label2(2).Caption = "External POP3 Provider: " & Split(List1.List(List1.ListIndex), "#")(2)
Label2(3).Caption = "External SMTP Server: " & Split(List1.List(List1.ListIndex), "#")(3)
Label2(4).Caption = "External " & Split(List1.List(List1.ListIndex), "#")(4)
Label2(6).Caption = "External Password " & Split(List1.List(List1.ListIndex), "#")(5)



End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnu
    End If
End Sub

Private Sub mnuCreate_Click()
    On Error Resume Next
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim MBUser As String
        rs.Open "select * from profiles", cn, adOpenKeyset, adLockOptimistic
            rs.AddNew
                MBUser = InputBox("Enter Mailbox Name", "New Mailbox")
                rs!user = MBUser
            If Not Len(MBUser) = 0 Then rs.Update
        rs.Close
    Set rs = Nothing
    mnuRefresh_Click
    
End Sub

Private Sub mnuExtPass_Click()
    On Error Resume Next
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim MBPASS As String
    Dim strUser As String
        strUser = Split(List1.Text, "#")(0): strUser = Split(strUser, "=")(1)
        rs.Open "select * from profiles where user =" & Chr(34) & strUser & Chr(34), cn, adOpenKeyset, adLockOptimistic
                MBPASS = InputBox("Enter External ISP Password for Mailbox " & strUser, "Set Password for Mailbox")
                rs!ExternalPass = MBPASS
            If Not Len(MBPASS) = 0 Then rs.Update
        rs.Close
    Set rs = Nothing
    mnuRefresh_Click
End Sub

Private Sub mnuEXTpop3_Click()
    On Error Resume Next
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim MBPOP As String
    Dim strUser As String
        strUser = Split(List1.Text, "#")(0): strUser = Split(strUser, "=")(1)
        rs.Open "select * from profiles where user =" & Chr(34) & strUser & Chr(34), cn, adOpenKeyset, adLockOptimistic
                MBPOP = InputBox("Enter External POP3 Server for Mailbox " & strUser, "External Mailbox Provider")
                rs!ExternalPOP3 = MBPOP
            If Not Len(MBPOP) = 0 Then rs.Update
        rs.Close
    Set rs = Nothing
    mnuRefresh_Click

End Sub

Private Sub mnuEXTsmtp_Click()
    On Error Resume Next
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim MBSMTP As String
    Dim strUser As String
        strUser = Split(List1.Text, "#")(0): strUser = Split(strUser, "=")(1)
        rs.Open "select * from profiles where user =" & Chr(34) & strUser & Chr(34), cn, adOpenKeyset, adLockOptimistic
                MBSMTP = InputBox("Enter External SMTP Server for Mailbox " & strUser, "External Mailbox Transfer")
                rs!ExternalSMTP = MBSMTP
            If Not Len(MBSMTP) = 0 Then rs.Update
        rs.Close
    Set rs = Nothing
    mnuRefresh_Click
End Sub

Private Sub mnuRefresh_Click()
List1.Clear
Call List_Profiles
End Sub

Private Sub mnurename_Click()
    On Error Resume Next
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim MBUser As String
    Dim strUser As String
        strUser = Split(List1.Text, "#")(0): strUser = Split(strUser, "=")(1)
        rs.Open "select * from profiles where user =" & Chr(34) & strUser & Chr(34), cn, adOpenKeyset, adLockOptimistic
                MBUser = InputBox("Rename Mailbox for " & strUser, "Renaming Mailbox")
                rs!user = MBUser
            If Not Len(MBUser) = 0 Then rs.Update
        rs.Close
    Set rs = Nothing
    mnuRefresh_Click
End Sub

Private Sub mnusetPass_Click()
    On Error Resume Next
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim MBPASS As String
    Dim strUser As String
        strUser = Split(List1.Text, "#")(0): strUser = Split(strUser, "=")(1)
        rs.Open "select * from profiles where user =" & Chr(34) & strUser & Chr(34), cn, adOpenKeyset, adLockOptimistic
                MBPASS = InputBox("Enter Password for Mailbox " & strUser, "Set Password for Mailbox")
                rs!PASS = MBPASS
            If Not Len(MBPASS) = 0 Then rs.Update
        rs.Close
    Set rs = Nothing
    mnuRefresh_Click
End Sub
Private Sub RemoveAll_History(Mailbox As String)
    On Error Resume Next
    Screen.MousePointer = 11
    Dim rs As ADODB.Recordset
    Dim strQry As String
    Dim i As Long
    Set rs = New ADODB.Recordset
        strQry = "Select * from FirstStore where RCPT = '<" & Mailbox & "@" & LCase(DNS) & ">'"
        rs.Open strQry, cn, adOpenKeyset, adLockPessimistic
            For i = 0 To rs.RecordCount
                rs.Delete
                rs.Requery
            Next i
        rs.Update
        rs.Close
    Set rs = Nothing
    Screen.MousePointer = 0
    MsgBox "All history for Mailbox (" & Mailbox & ") has been Removed" & vbNewLine & "Optimize the Database now, for performance and minimize the Database size" & vbNewLine & "Tip. Make sure you close the User Profiles Window First", vbInformation, "Mailbox History Deletion"
End Sub

Private Sub mnuUserLogon_Click()
    On Error Resume Next
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim MBUser As String
    Dim strUser As String
        strUser = Split(List1.Text, "#")(0): strUser = Split(strUser, "=")(1)
        rs.Open "select * from profiles where user =" & Chr(34) & strUser & Chr(34), cn, adOpenKeyset, adLockOptimistic
                MBUser = InputBox("Enter External User ID Server for Mailbox " & strUser, "(This is for your External Account logon POP3! not local!)")
                rs!Externaluser = MBUser
            If Not Len(MBUser) = 0 Then rs.Update
        rs.Close
    Set rs = Nothing
    mnuRefresh_Click
End Sub

Private Sub munRemove_Click()
    On Error Resume Next
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim ComfirmDel, ComfirmHistoryDel As String
    Dim strUser As String
        strUser = Split(List1.Text, "#")(0): strUser = Split(strUser, "=")(1)
        rs.Open "select * from profiles where user =" & Chr(34) & strUser & Chr(34), cn, adOpenKeyset, adLockOptimistic
            ComfirmDel = MsgBox("Delete " & strUser & " Mailbox?", vbInformation + vbYesNo, "Comfirm Mailbox Deletion")
                If ComfirmDel = vbYes Then rs.Delete
        rs.Close
    Set rs = Nothing
        If ComfirmDel = vbYes Then
            ComfirmHistoryDel = MsgBox("Would you like to remove all History from Mailbox (" & strUser & ")? this will free up database space", vbInformation + vbYesNo, "Mailbox History Deletion")
            If ComfirmHistoryDel = vbYes Then Call RemoveAll_History(strUser)
        End If
    mnuRefresh_Click

End Sub
