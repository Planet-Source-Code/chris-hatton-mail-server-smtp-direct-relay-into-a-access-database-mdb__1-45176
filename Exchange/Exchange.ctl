VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl Exchange 
   ClientHeight    =   7905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11295
   ScaleHeight     =   7905
   ScaleWidth      =   11295
   ToolboxBitmap   =   "Exchange.ctx":0000
   Begin MSComctlLib.StatusBar StatusBar2 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   7395
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   19686
            MinWidth        =   19686
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame6 
      Height          =   6735
      Left            =   7800
      TabIndex        =   10
      Top             =   0
      Width           =   3375
      Begin MSComctlLib.ListView ListView2 
         Height          =   3210
         Left            =   5
         TabIndex        =   12
         Top             =   360
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   5662
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   16777215
         BackColor       =   10841658
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   5857
         EndProperty
      End
      Begin VB.Frame Frame7 
         Height          =   375
         Left            =   -5
         TabIndex        =   11
         Top             =   0
         Width           =   3375
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "SMTP QUE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   13
            Top             =   120
            Width           =   3375
         End
      End
      Begin VB.Frame Frame9 
         Height          =   375
         Left            =   -5
         TabIndex        =   17
         Top             =   3480
         Width           =   3375
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "POP3 Connection Status"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   23
            Top             =   120
            Width           =   3375
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Label7"
            ForeColor       =   &H00FFC0C0&
            Height          =   190
            Left            =   30
            TabIndex        =   22
            Top             =   140
            Visible         =   0   'False
            Width           =   3300
         End
         Begin VB.Shape Shape1 
            Height          =   230
            Left            =   10
            Top             =   120
            Visible         =   0   'False
            Width           =   3330
         End
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   2600
         Left            =   5
         TabIndex        =   20
         Top             =   3840
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   4577
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483643
         BackColor       =   10841658
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   5328
         EndProperty
      End
      Begin VB.Frame Frame10 
         Height          =   375
         Left            =   0
         TabIndex        =   18
         Top             =   6360
         Width           =   3375
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   120
            Width           =   3135
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6960
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Exchange.ctx":0312
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Exchange.ctx":0764
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Exchange.ctx":0BB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Exchange.ctx":1008
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Exchange.ctx":145A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Exchange.ctx":1774
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Exchange.ctx":1BC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Exchange.ctx":1EE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Exchange.ctx":21FA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   6960
      Top             =   5520
   End
   Begin VB.Frame Frame1 
      Height          =   6735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7800
      Begin MSWinsockLib.Winsock SckExch 
         Index           =   0
         Left            =   6480
         Top             =   6120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Frame Frame4 
         Height          =   6735
         Left            =   2150
         TabIndex        =   6
         Top             =   0
         Width           =   5655
         Begin VB.Timer Timer2 
            Interval        =   15000
            Left            =   4320
            Top             =   5520
         End
         Begin MSWinsockLib.Winsock sckPOP3 
            Index           =   0
            Left            =   3720
            Top             =   6120
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   393216
         End
         Begin VB.Frame Frame5 
            Height          =   375
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   5655
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "  Mailbox                             Total Items        Size (KB)                 Last Email"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   9
               Top             =   120
               Width           =   5655
            End
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   6315
            Left            =   35
            TabIndex        =   7
            Top             =   375
            Width           =   5580
            _ExtentX        =   9843
            _ExtentY        =   11139
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            FlatScrollBar   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            SmallIcons      =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "User Name"
               Object.Width           =   3246
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "TotalEmails"
               Object.Width           =   1834
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Size"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "LastLogon"
               Object.Width           =   2205
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Height          =   6735
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   2175
         Begin VB.Frame Frame3 
            Height          =   375
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   2175
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Internal Connections"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   4
               Top             =   120
               Width           =   1935
            End
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   2970
            Left            =   15
            TabIndex        =   5
            Top             =   345
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   5239
            _Version        =   393217
            LabelEdit       =   1
            Style           =   7
            ImageList       =   "ImageList1"
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComctlLib.TreeView TreeView2 
            Height          =   3060
            Left            =   15
            TabIndex        =   14
            Top             =   3630
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   5398
            _Version        =   393217
            LabelEdit       =   1
            Style           =   7
            ImageList       =   "ImageList1"
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Frame Frame8 
            Height          =   375
            Left            =   0
            TabIndex        =   15
            Top             =   3240
            Width           =   2175
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "External Connections"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Top             =   120
               Width           =   1935
            End
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7650
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3758
            MinWidth        =   3758
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3829
            MinWidth        =   3829
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5946
            MinWidth        =   5946
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Exchange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public cn As ADODB.Connection
Public DBPATH As String
Dim sckIndex As Long 'Winsock Sockets
Dim inData As String 'Incoming data
Dim strOut As String 'Output string for senddata
Public DNS As String 'hatton.net
Dim MsgID As Long 'message id (RETR 1)
Dim sckState As Integer 'list of socket states
Dim JoData(0 To 1) As String
Dim strCMDLNE As String
Dim popcounter As Long
Dim MaxconRetrys As Integer 'maxium connection retrys
Dim initpop3(0 To 3) 'pop3 counter; for checking pop account every so many minutes
Dim ConRetrys As Integer 'counter for how many retrys
Dim CheckWait As Long 'check for new message after so many minutes
Private Declare Function GetTickCount Lib "kernel32" () As Long 'system timer
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private User As String
Private Const MAX_PATH = 260
Private Enum POP3State
    userid = 0
    pass = 1
    stat = 2
    retr = 3
    dele = 4
    quit = 5
    Data = 6
    closecon = 7
End Enum
Dim doPOPState As POP3State
Dim objPop3 As POP3STREAM
Dim objPop3Session As New POP3GRP


 
Public Sub CompactJetDatabase(Location As String, Optional BackupOriginal As Boolean = True)
On Error Resume Next

User = "ADMIN"

'On Error GoTo CompactErr
    Screen.MousePointer = 11
On Error Resume Next
Dim strBackupFile As String
Dim strTempFile As String

'Check the database exists
If Len(Dir(Location)) Then

    ' If a backup is required, do it!
    If BackupOriginal = True Then
        strBackupFile = GetTemporaryPath & "backup.mdb"
        If Len(Dir(strBackupFile)) Then Kill strBackupFile
        FileCopy Location, strBackupFile
    End If

    ' Create temporary filename
    strTempFile = GetTemporaryPath & "temp.mdb"
    If Len(Dir(strTempFile)) Then Kill strTempFile

    ' Do the compacting via DBEngine
    DBEngine.CompactDatabase Location, strTempFile
    Kill Location
    FileCopy strTempFile, Location
    Kill strTempFile

Else
    Screen.MousePointer = 0
    MsgBox "Mailbox database not found" & vbNewLine & Location, vbCritical, "Database Optimizer"
    Exit Sub
    
End If
    Screen.MousePointer = 0
    MsgBox "Mailbox is finished Optimizing", vbInformation, "Database Optimizer"


User = ""
    Exit Sub
CompactErr:
    Screen.MousePointer = 0
    MsgBox err.Description
  
    Exit Sub

End Sub

Public Function GetTemporaryPath()
On Error Resume Next
Dim strFolder As String
Dim lngResult As Long

strFolder = String(MAX_PATH, 0)
lngResult = GetTempPath(MAX_PATH, strFolder)

If lngResult <> 0 Then
  GetTemporaryPath = Left(strFolder, InStr(strFolder, _
    Chr(0)) - 1)
Else
  GetTemporaryPath = ""
End If

End Function

Private Sub ADO_Connect()
On Error Resume Next
Set cn = New ADODB.Connection
        cn.Provider = "Microsoft.Jet.OLEDB.4.0"
        cn.Open DBPATH
        cn.CursorLocation = adUseClient
   
    If cn.State = 1 Then
        StatusBar1.Panels(1).Text = "Mailbox Database Ready"
        DoEvents
    Else
        StatusBar1.Panels(1).Text = "Mailbox Database Stopped"
    End If
End Sub

Public Sub Stop_OnlineStore()
On Error Resume Next
    Call ClearALL
    Timer2.Enabled = False
    cn.Close
    sckPOP3(0).Close
    SckExch(0).Close
    Set cn = Nothing
End Sub
Private Sub ClearALL()
On Error Resume Next
Dim i As Long
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    TreeView1.Nodes.Clear
    TreeView2.Nodes.Clear
For i = 1 To StatusBar1.Panels.Count
        StatusBar1.Panels.Item(i).Text = ""
Next i
End Sub
Public Sub Start_OnlineStore()
On Error Resume Next
    Call ADO_Connect
    Call Accept_IncomingCalls
    Call AccStatus 'get the account status of each user
    CheckWait = 120 'wait 2 minute before checking mail
    Timer2.Enabled = True
End Sub

Private Sub Accept_IncomingCalls()
On Error Resume Next
SckExch(0).LocalPort = 110
SckExch(0).listen

If SckExch(0).State = sckListening Then
    StatusBar1.Panels(2).Text = "POP3 Service Started"
    Else
    StatusBar1.Panels(2).Text = "Winsock/POP3 Service Stopped"
End If


End Sub
Private Function Get_NextPort() As Integer
Dim i As Integer
On Error Resume Next
For i = 1 To 255 'reuse only 255 Simultaneous connections at once, if above that then allocate extra connections
    If SckExch(i).State = sckClosed Then
        Get_NextPort = i
        Exit Function
    End If
Next i

End Function

Private Sub cmdcompact_Click()
On Error Resume Next
Call Stop_OnlineStore
Call CompactJetDatabase(DBPATH)
Call Start_OnlineStore
End Sub
Private Sub GetUserObject(strUser As String)
On Error Resume Next
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim strQry As String
    strQry = "select * from profiles where user = " & Chr(34) & strUser & Chr(34)
        rs.Open strQry, cn, adOpenStatic, adLockReadOnly
            With FrmPopAccounts
                .Username = rs!User
                .EXTUSER = rs!externaluser
                .POP3ACC = rs!externalpop3
                .SMTPACC = rs!externalsmtp
                .Show 1
            End With
        rs.Close
Set rs = Nothing
End Sub

Private Sub ListView1_Click()
TreeView1.SetFocus


End Sub

Private Sub ListView1_DblClick()
Call GetUserObject(ListView1.SelectedItem.Text)

End Sub

Private Sub ListView2_Click()
TreeView1.SetFocus

End Sub

Private Sub SckExch_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
If Get_NextPort = 0 Then
    sckIndex = sckIndex + 1
Else
    sckIndex = Get_NextPort
End If

Load SckExch(sckIndex)
SckExch(sckIndex).accept requestID
SckExch(sckIndex).SendData "+OK Hattech Technologies POP3 Server version " & App.Major & "." & App.Minor & "." & App.Revision & " (" & SckExch(sckIndex).LocalHostName & "." & LCase(DNS) & ") Ready" & vbCrLf

End Sub
Private Sub SetPassword(strPass As String, lngPort As Integer, tv As TreeView)
Dim i As Long
Dim strItem As Variant
On Error GoTo err
    For i = 1 To tv.Nodes.Count
        tv.Nodes.Item(i).Expanded = True
        strItem = tv.Nodes.Item(i).Text
        If strItem = lngPort Then
                tv.Nodes.Item(i).Parent.Tag = strPass 'set the password to the current logged in user for other access
        End If
    Next i
Exit Sub
err:

End Sub
Private Sub AddConnection(strUser As String, lngPort As Integer, tv As TreeView)
On Error Resume Next
Dim ConnGroup As Node
Dim tvItm As Long

For tvItm = 1 To tv.Nodes.Count
    If tv.Nodes.Item(tvItm).Text = strUser Then Exit Sub
Next tvItm

Set ConnGroup = tv.Nodes.Add(, , , strUser, 1)
Set ConnGroup = tv.Nodes.Add(ConnGroup, tvwChild, , lngPort, 2)

End Sub
Private Sub RemoveConnection(strUser As String, lngPort As Integer, tv As TreeView)
Dim tvItm As Long
On Error GoTo err
Dim strItem As Variant
For tvItm = 1 To tv.Nodes.Count
    strItem = tv.Nodes.Item(tvItm).Text
    If strItem = lngPort Then
        tv.Nodes.Remove (tvItm)
        tv.Nodes.Remove (tvItm - 1)
    End If
    
Next tvItm

Exit Sub
err:

End Sub
Private Function CheckUserName(strUser As String, lngPort As Integer) As Boolean
On Error Resume Next
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim strQry As String
    strQry = "select user from profiles where user = '" & strUser & "'"

    rs.Open strQry, cn, adOpenKeyset, adLockReadOnly
    If rs.RecordCount = 1 Then
        CheckUserName = True
        AddConnection strUser, lngPort, TreeView1  'add user name to the treeview
    End If
    rs.Close
Set rs = Nothing
End Function

Private Function CheckPass(strUser As String, Optional strPass As String, Optional lngPort As Integer) As Boolean
On Error Resume Next
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim strQry As String
    strQry = "select * from profiles where user = '" & strUser & "'"
    rs.Open strQry, cn, adOpenKeyset, adLockReadOnly
    If rs!pass = strPass Then
            CheckPass = True 'if password is correct allow logon access
        Else
    If rs!pass = Verify_TAG(lngPort, TreeView1) Then CheckPass = True    'allow commands to executed
                                                                         'only if password is correct on logon.
                                                                         'function only = true if port has been specified('lngport')
    End If
    rs.Close
Set rs = Nothing

End Function
Private Function Check_Logon(lngPort As Integer, tv As TreeView) As Boolean 'find out if user is already
Dim i As Long                                                               'logged in
Dim strItem As Variant
On Error GoTo err
    For i = 1 To tv.Nodes.Count
        tv.Nodes.Item(i).Expanded = True
        strItem = tv.Nodes.Item(i).Text
        If strItem = lngPort Then Check_Logon = True
    Next i
    
Exit Function
err:

End Function
Private Function Verify_TAG(lngPort As Integer, tv As TreeView) As String 'find out what port belongs to what user
Dim i As Long                                                              'that is logged in and verify password
Dim strItem As Variant                                                     '* Returns password of user in question
On Error GoTo err
    For i = 1 To tv.Nodes.Count
        tv.Nodes.Item(i).Expanded = True
        strItem = tv.Nodes.Item(i).Text
        If strItem = lngPort Then
                Verify_TAG = tv.Nodes.Item(i).Parent.Tag 'this should = the password of the authorised user
        End If
    Next i
    

Exit Function
err:

End Function
Private Function STAT_RTN(strUser As String)

Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim strQry As String
Dim i As Long
Dim X As Long ' total number of octet's
Dim z As Long ' total number of messages
On Error GoTo err
    strQry = "Select * from FirstStore where RCPT = '<" & strUser & "@" & UCase(DNS) & ">'"
        rs.Open strQry, cn, adOpenKeyset, adLockReadOnly
            For i = 1 To rs.RecordCount
                    z = rs.RecordCount
                    X = X + rs!octat
                rs.MoveNext
            Next i

        rs.Close
Set rs = Nothing

STAT_RTN = "+OK " & z & " " & X
Exit Function
err:
STAT_RTN = "+OK 0 0" 'if error has occurred then status = 0 prevents retrieval of emails

End Function
Private Sub MSGID_RTN(strUser As String)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim strQry As String
Dim i As Long
On Error GoTo err
    strQry = "Select * from FirstStore where RCPT = '<" & strUser & "@" & LCase(DNS) & ">'"
        rs.Open strQry, cn, adOpenKeyset, adLockOptimistic
            For i = 1 To rs.RecordCount
               rs!MsgID = i 'create an idenification for each message for separate users
                rs.MoveNext
            Next i
err: 'incase of err unload any leaky code
On Error Resume Next
        rs.Close
Set rs = Nothing

End Sub

Private Function LIST_RTN(strUser As String)

Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim strQry As String
Dim j As Long 'total messages
Dim X As Long 'total octet's
Dim z As Long 'indivial octet
Dim i As Long
Dim List As String
On Error GoTo err
    strQry = "Select octat, rcpt from FirstStore where RCPT = '<" & strUser & "@" & LCase(DNS) & ">'"
        rs.Open strQry, cn, adOpenKeyset, adLockOptimistic
            For i = 1 To rs.RecordCount
                j = j + 1
                X = X + rs!octat
                z = rs!octat
                    List = List & j & " " & z & vbCrLf
                rs.MoveNext
            Next i

        rs.Close
Set rs = Nothing

LIST_RTN = "+OK " & j & " " & X & vbCrLf & List & "." & vbCrLf

Call AccStatus 'get the account status of each user

Exit Function
err:


End Function
Private Function RETR_RTN(strUser As String, MsgID As Long) 'message header and body

Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim strQry As String
On Error GoTo err
    strQry = "Select * from FirstStore where RCPT = '<" & strUser & "@" & LCase(DNS) & ">'" & " And MsgID = " & MsgID
        rs.Open strQry, cn, adOpenKeyset, adLockReadOnly

                    RETR_RTN = "+OK" & vbCrLf & rs!body & vbCrLf & vbCrLf & "." & vbCrLf
        rs.Close
Set rs = Nothing

Exit Function
err:
RETR_RTN = "+OK" & vbCrLf & _
                    "Subject: POP3 Request Database Error" & vbCrLf & _
                    "From: Mail Administrator" & vbCrLf & _
                    "To: " & strUser & "@" & LCase(DNS) & vbCrLf & vbCrLf & _
                    "This is a automated message PLEASE DO NOT REPLY. " & vbCrLf & "There has been an error in the Requested Email " & vbCrLf & _
                    "Message Idenification: " & MsgID - 1 & " User: " & strUser & _
                    "" & vbCrLf & vbCrLf & "." & vbCrLf

End Function

Private Function DELE_RTN(strUser As String, MsgID As Long) 'delete message

Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim strQry As String
On Error GoTo err
    strQry = "Select * from FirstStore where RCPT = '<" & strUser & "@" & LCase(DNS) & ">'" & " And MsgID = " & MsgID
        rs.Open strQry, cn, adOpenKeyset, adLockOptimistic
            'rs.Move MsgID - 1
            rs.Delete 'MsgID

        rs.Close
Set rs = Nothing

Exit Function
err:
End Function
Private Function Verify_User(lngPort As Integer, tv As TreeView) As String 'find out what port belongs to what user
Dim i As Long                                                              'that is logged in
Dim strItem As Variant
On Error GoTo err
    For i = 1 To tv.Nodes.Count
        tv.Nodes.Item(i).Expanded = True
        strItem = tv.Nodes.Item(i).Text
        If strItem = lngPort Then
                Verify_User = tv.Nodes.Item(i).Parent
        End If
    Next i
    
Exit Function
err:

End Function
Public Sub RepairDB()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim strQry As String
On Error GoTo err
    strQry = "Select * from FirstStore"
        rs.Open strQry, cn, adOpenKeyset, adLockOptimistic
            Do Until rs.EOF = True
                If IsNull(rs!body) = True Then rs.Delete adAffectCurrent
                rs.MoveNext
            Loop

        rs.Close
Set rs = Nothing


Call AccStatus 'get the account status of each user
err:



End Sub
Public Sub AccStatus()
On Error Resume Next
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim strQry As String
    Set rs = New ADODB.Recordset
    Screen.MousePointer = 11
    On Error Resume Next
    ListView1.ListItems.Clear
    Dim i, j As Long
    Dim lngOctat As Long
    rs.Open "Select * from profiles", cn, adOpenStatic, adLockReadOnly
        For j = 1 To rs.RecordCount
            ListView1.ListItems.Add , , rs!User, , 9
            StatusBar1.Panels(3).Text = "MailBox Size: " & DBFileSize(App.Path & "\DBSTORE.MDB")
         Set rs2 = New ADODB.Recordset
         rs2.Open "Select * from FirstStore where RCPT = '<" & ListView1.ListItems(ListView1.ListItems.Count).Text & "@" & UCase(DNS) & ">'", cn, adOpenStatic, adLockReadOnly
             For i = 1 To rs2.RecordCount 'loop through all emails for total octats
                lngOctat = lngOctat + rs2!octat
                rs2.MoveNext
             Next i
                    rs2.Requery
             For i = 1 To rs2.RecordCount
                With ListView1.ListItems.Item(j).ListSubItems
                    .Add , , rs2.RecordCount
                    .Add , , lngOctat & " (KB)"
                    .Add , , rs2!Date
                End With
                    rs2.MoveNext
             Next i
                    lngOctat = 0
            rs2.Close
        Set rs2 = Nothing
            rs.MoveNext
        Next j
    rs.Close
    Set rs = Nothing
    Screen.MousePointer = 0

End Sub

Public Sub SMTPQue() 'get the status
    On Error Resume Next
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim Octets As Long
    Dim i As Long
        ListView2.ListItems.Clear
        rs.Open "select * from SMTPQUE", cn, adOpenKeyset, adLockReadOnly
            For i = 1 To rs.RecordCount
                Octets = Octets + rs!octat
                    ListView2.ListItems.Add , , "From: " & MailFrom(rs!body) & " " & rs!Date & " " & rs!octat & " (KB)", , 6
                    ListView2.ListItems.Item(i).Tag = rs!unquieid
                rs.MoveNext
            Next i
        StatusBar1.Panels(4).Text = "Total Outbox " & Octets & " (KB)"
        rs.Close
    Set rs = Nothing
End Sub
Private Function MailFrom(objBody As String)
On Error Resume Next
    Dim strTemp(0 To 1) As String
        strTemp(0) = Split(objBody, "From: """)(1)
        strTemp(1) = Split(strTemp(0), """")(0)
    MailFrom = strTemp(1)
End Function
Private Sub SckExch_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
SckExch(Index).GETDATA inData
  If InStr(1, inData, vbCrLf) > 0 Then
          strCMDLNE = strCMDLNE & inData
          inData = Split(strCMDLNE, vbCrLf)(0)
        Else
            strCMDLNE = strCMDLNE & inData 'joins charactors
        Exit Sub
    End If
    JoData(0) = JoData(0) & inData
    JoData(1) = JoData(0)
    JoData(0) = Split(JoData(0), " ")(0)
    JoData(1) = Split(JoData(1), " ")(1)

strCMDLNE = ""
        Select Case UCase(JoData(0))

                Case UCase("user")
                    If Check_Logon(Index, TreeView1) = True Then GoTo InvalidState
                    If CheckUserName(JoData(1), Index) = True Then
                        GoTo OK
                       
                    End If
                GoTo ClearBuffer

                Case UCase("pass")
                            
                    If Not CheckUserName(Verify_User(Index, TreeView1), Index) = True Then 'if false then reset
                            Call RemoveConnection(Verify_User(Index, TreeView1), Index, TreeView1)
                            GoTo InvalidState
                    End If
                            
                    If CheckPass(Verify_User(Index, TreeView1), JoData(1)) = True Then 'if user is true then check pass
                            strOut = "+OK User Successfully logged on." & vbCrLf
                                Call SetPassword(JoData(1), Index, TreeView1) 'save the password into memory
                                Call MSGID_RTN(Verify_User(Index, TreeView1)) 'index users messages ready for download
                        Else
                            strOut = "-ERR Logon failure: unknown user name or bad password." & vbCrLf
                                Call RemoveConnection(Verify_User(Index, TreeView1), Index, TreeView1) 'remove connection
                    End If
                    
                GoTo ClearBuffer

                Case UCase("Quit")
                    Call RemoveConnection(Verify_User(Index, TreeView1), Index, TreeView1)
                        SckExch(Index).SendData "Closing Connection." & vbCrLf
                        SckExch(Index).Close
                        Unload SckExch(Index)
                
                Case UCase("Noop")
                    If CheckPass(Verify_User(Index, TreeView1), Verify_TAG(Index, TreeView1), Index) = False Then GoTo InvalidState
                    GoTo OK
                    
                Case UCase("STAT")
                    If CheckPass(Verify_User(Index, TreeView1), Verify_TAG(Index, TreeView1), Index) = False Then GoTo InvalidState
                        strOut = STAT_RTN(Verify_User(Index, TreeView1)) & vbCrLf 'get statistics
                            GoTo ClearBuffer 'send message and clear variables
                
                Case UCase("DELE")
                    If CheckPass(Verify_User(Index, TreeView1), Verify_TAG(Index, TreeView1), Index) = False Then GoTo InvalidState
                        MsgID = JoData(1)
                        DELE_RTN Verify_User(Index, TreeView1), MsgID
                        strOut = "+OK" & vbCrLf
                            GoTo ClearBuffer 'send message and clear variables
                    
                Case UCase("LIST")
                    If CheckPass(Verify_User(Index, TreeView1), Verify_TAG(Index, TreeView1), Index) = False Then GoTo InvalidState
                        strOut = LIST_RTN(Verify_User(Index, TreeView1))
                            GoTo ClearBuffer 'send message and clear variables
                    
                Case UCase("RETR")
                    If CheckPass(Verify_User(Index, TreeView1), Verify_TAG(Index, TreeView1), Index) = False Then GoTo InvalidState
                            MsgID = JoData(1)
                            strOut = RETR_RTN(Verify_User(Index, TreeView1), MsgID)
                            GoTo ClearBuffer 'send message and clear variables

                   
        End Select

        strOut = "-ERR Protocol" & vbCrLf

ClearBuffer:

SckExch(Index).SendData strOut
Call Tidy_UP
Exit Sub

InvalidState:
SckExch(Index).SendData "-ERR Command is not valid in this state" & vbCrLf
Call Tidy_UP
Exit Sub

OK:
SckExch(Index).SendData "+OK" & vbCrLf
Call Tidy_UP

End Sub

Private Sub Tidy_UP()

On Error Resume Next
    JoData(0) = ""
    JoData(1) = ""

End Sub


Private Sub sckPOP3_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error GoTo conrequest
sckPOP3(Index).Close
sckPOP3(Index).accept requestID

Exit Sub
conrequest:
'MsgBox "Err: 601 " & vbNewLine & err.Description, vbCritical
StatusBar2.Panels(1).Text = "Last Error: 601"
End Sub
Private Sub ReadOptions()
On Error Resume Next
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
    rs.Open "select * from settings where configurationset = 1", cn, adOpenKeyset, adLockOptimistic
        MaxconRetrys = rs!ConnectionRetrys
        CheckWait = rs!pop3timer * 60
    rs.Close
Set rs = Nothing

End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If CheckWait = 0 Then Exit Sub

DoEvents
initpop3(3) = GetTickCount / 1000 'current timer
If Timer2.Enabled = False Then Exit Sub

If initpop3(2) < initpop3(3) Then
    initpop3(2) = GetTickCount / 1000 + CheckWait
    Call ProcessPOPMail   'exit if timer current is greater than timer in advance
End If



End Sub
Public Sub ProcessPOPMail()
On Error GoTo processmail
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
    rs.Open "select * from profiles", cn, adOpenStatic, adLockReadOnly
       If popcounter = rs.RecordCount Then popcounter = 0: GoTo finish
            Do
          DoEvents: DoEvents: DoEvents: DoEvents
           ' Debug.Print sckPOP3(0).State
                If sckPOP3(0).State = 9 Then sckPOP3(0).Close
                If sckPOP3(0).State = 8 Then sckPOP3(0).Close
                 If sckPOP3(0).State = 0 Then sckPOP3(0).Close
            Loop Until sckPOP3(0).State = sckClosed
                'popcounter = popcounter + 1
            Call Checkmail(popcounter)
finish:
    rs.Close
Set rs = Nothing
initpop3(2) = GetTickCount / 1000 + CheckWait 'timer in advance
Exit Sub
processmail:
'MsgBox "Err: 602 " & vbNewLine & err.Description, vbCritical
StatusBar2.Panels(1).Text = "Last Error: 602"
sckPOP3(0).Close

End Sub
Private Sub Checkmail(lngUser As Long)
On Error Resume Next
Call ReadOptions 'get the current pop3 timer options and maximum retries
'ListView3.ListItems.Clear
ListView3.ListItems.Add , , "Last Checked " & Time
ListView3.ListItems(ListView3.ListItems.Count).EnsureVisible
ConRetrys = 0
GetPop3Details (lngUser)
SumitConnection (lngUser)

If SumitConnection(lngUser) = True Then
 Label6.Caption = "POP3 is Healthy"
 Call AccStatus
 Call ProcessPOPMail ': popcounter = popcounter + 1 'process next pop account
 Else
 Label6.Caption = "POP3 has Finished"
    Label7.Visible = False
    Shape1.Visible = False

  Call ProcessPOPMail ': popcounter = popcounter + 1 'process next pop account
End If
End Sub
Private Sub GetPop3Details(inxUser As Integer)
 On Error Resume Next
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strQry As String
        strQry = "select * from profiles"
            rs.Open strQry, cn, adOpenStatic, adLockReadOnly
            rs.Move (inxUser)
             objPop3Session.objUser = rs!User
             objPop3Session.EXTUSER = rs!externaluser
             objPop3Session.objPASS = rs!externalpass
             objPop3Session.EXTPOP3 = rs!externalpop3
            rs.Close
    Set rs = Nothing
End Sub
Private Function SumitConnection(inxUser As Long) As Boolean
On Error GoTo sumitcon
If ConRetrys = MaxconRetrys Then Exit Function
DoEvents
sckPOP3(0).Close
sckPOP3(0).connect objPop3Session.EXTPOP3, 110
ConRetrys = ConRetrys + 1

If ConnectionWait(10, 0) = True Then 'it was successful connection
    SumitConnection = True
    ConRetrys = 0 'reset counter
    Label6.Caption = "POP3 Connected"
    Label7.Visible = True
    Shape1.Visible = True

Else
    Label6.Caption = "Connection Time-Out"
    sckPOP3(0).Close
    SumitConnection (inxUser) 'retry connecting again
    Label7.Visible = True
    Shape1.Visible = True

End If
Exit Function
sumitcon:
'MsgBox "Err: 603 " & vbNewLine & err.Description, vbCritical
StatusBar2.Panels(1).Text = "Last Error: 603"
sckPOP3(0).Close
    Label7.Visible = False
    Shape1.Visible = False

End Function

'POP3 Retriever ISP / External / Relay / Chained POP3
Private Sub sckPOP3_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Dim POP3Data As String
Dim strOut3 As String
Dim retSize As Long
Dim retTotal As Long
Dim strBody As String
sckPOP3(Index).GETDATA POP3Data

If InStr(1, POP3Data, vbCrLf) > 1 Then popdata = Split(POP3Data, vbCrLf)(0)
'Debug.Print POP3Data

ListView3.ListItems.Add , , Left$(Left$(POP3Data, Len(POP3Data) - 2), 40)
ListView3.ListItems.Item(ListView3.ListItems.Count).EnsureVisible
'MsgBox popdata
    Select Case doPOPState
             
        Case userid
          
           If InStr(1, popdata, "+OK") Then 'have been welcomed let send our user identification
                strOut3 = "User " & Trim$(objPop3Session.EXTUSER): doPOPState = pass 'ready for pass
                Call AddConnection(sckPOP3(Index).RemoteHostIP, Index, TreeView2)
                GoTo Parse
            End If
        
        Case pass
            If InStr(1, popdata, "+OK") Then 'we have been accepted
                strOut3 = "Pass " & Trim$(objPop3Session.objPASS): doPOPState = stat 'ready to get the mail status
                GoTo Parse
            End If
           
        Case stat
            If InStr(1, popdata, "+OK") Then
                strOut3 = "STAT": doPOPState = retr 'ok get the status of emails and goto the retreival mode
            GoTo Parse
            End If
            
        Case retr
            If InStr(1, popdata, "+OK") Then
                Set objPop3 = New POP3STREAM 'a new message there create a new memory space for it
                   On Error Resume Next
                   If InStr(1, popdata, "+OK ") Then objPop3Session.Octets = Split(popdata, " ")(2)  ' total octet's
                    If InStr(1, popdata, "+OK ") Then objPop3Session.Total = Split(popdata, " ")(1)  'total messages
                    If objPop3Session.Total = 0 Then
                        strOut3 = "NOOP": doPOPState = quit 'no more messages quit here
                        GoTo Parse
                    Else
                        strOut3 = "RETR " & objPop3Session.Total: doPOPState = dele 'call out for the new message
                        GoTo Parse
                    End If
            End If
            
        Case dele
            If InStr(1, popdata, "+OK") Then 'were actually recieving the message here and if OK then delete it from the server
                    If InStr(1, POP3Data, "octets") > 1 Then POP3Data = Split(POP3Data, "octets")(1)
                    If Mid$(POP3Data, 1, 2) = vbCrLf Then POP3Data = Right$(POP3Data, Len(POP3Data) - 2)
                        objPop3.objBody = objPop3.objBody & POP3Data 'save first block of data
                        doPOPState = Data 'ok now detect EOF and delete message on server
                        Call AccStatus
            End If
        
        
        Case quit
            If InStr(1, popdata, "+OK") Then
                strOut3 = "quit"
                doPOPState = closecon 'close the connection
                Call RemoveConnection(sckPOP3(Index).RemoteHostIP, Index, TreeView2)
                Call AccStatus
                Set objPop3Session = Nothing
                GoTo Parse
            End If
            
        Case Data
        
            objPop3.objBody = objPop3.objBody & POP3Data 'just exit and wait for next block
                If InStr(1, POP3Data, vbCrLf & "." & vbCrLf) = 1 Then 'check for EOF
                        Call SavePOP3DATA(objPop3Session.objUser, strBody, objPop3Session.Octets) 'message completed get another
                        strOut3 = "DELE " & objPop3Session.Total
                        objPop3Session.Total = objPop3Session.Total - 1 'ready for next message
                        doPOPState = retr 'go back to retr for next message
                        Set objPop3 = Nothing 'finished with this message create new instance
                        GoTo Parse
                ElseIf InStr(1, objPop3.objBody, vbCrLf & "." & vbCrLf) > 1 Then 'if the first block as eof then do delete
                        Call SavePOP3DATA(objPop3Session.objUser, objPop3.objBody, objPop3Session.Octets)
                        strOut3 = "DELE " & objPop3Session.Total 'process delete for this message
                        objPop3Session.Total = objPop3Session.Total - 1 'ready for next message
                        doPOPState = retr 'go back loop it for next message
                        Set objPop3 = Nothing 'distory it, already save in database
                        GoTo Parse 'ok process the delete message and once the next message come in process that too.
                End If
        Case closecon
             If InStr(1, popdata, "+OK") Then
                 sckPOP3(Index).Close
                 doPOPState = userid 'reset to the start
                Label6.Caption = "Closed"
                popcounter = popcounter + 1
             End If
        

    End Select

 If InStr(1, UCase(popdata), "-ERR") Then
    strOut3 = "Quit"
    doPOPState = closecon 'close the connection
    GoTo Parse
 End If
 
 
 
Exit Sub
Parse:


'Debug.Print strOut3
    sckPOP3(Index).SendData strOut3 & vbCrLf



End Sub
Private Function SavePOP3DATA(strUser, strBody As String, Octets As Long) As Boolean
On Error Resume Next
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim strQry As String
    strQry = "select * from firststore"
    rs.Open strQry, cn, adOpenKeyset, adLockOptimistic
        rs.AddNew
            rs!RCPT = "<" & strUser & "@" & LCase(DNS) & ">"
            rs!Date = Date
            rs!body = strBody
            rs!octat = Octets
    rs.Update
    rs.Close
Set rs = Nothing

If Len(err.Description) = 0 Then
    SavePOP3DATA = True
Else
    SavePOP3DATA = False
End If

End Function
Private Sub Timer1_Timer()
On Error Resume Next
        For sckState = 1 To SckExch.UBound
                If SckExch(sckState).State = 9 Then GoTo badstate
                If SckExch(sckState).State = 8 Then GoTo badstate
        Next sckState
    Exit Sub
badstate:
                    Call RemoveConnection(Verify_User(sckState, TreeView1), sckState, TreeView1) 'remove pending connection
                    SckExch(sckState).Close 'close the connection if still pending
                    Unload SckExch(sckState) 'no longer need a dead connection so we will free up some memory
End Sub
Private Function ConnectionWait(seconds As Long, sckIndex As Integer) As Boolean
On Error Resume Next
initpop3(0) = GetTickCount / 1000 + seconds 'timer in advance
Do

    initpop3(1) = GetTickCount / 1000 'current timer
    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents
   ' Debug.Print sckPOP3(sckIndex).State
    If sckPOP3(sckIndex).State = 4 Then Label6.Caption = "Resolving Host"
    If sckPOP3(sckIndex).State = 9 Then: sckPOP3(sckIndex).Close: ConnectionWait = False: Exit Function
    If sckPOP3(sckIndex).State = 8 Then: sckPOP3(sckIndex).Close: ConnectionWait = False: Exit Function
    If sckPOP3(sckIndex).State = 7 Then: ConnectionWait = True: Exit Function
    
Loop Until initpop3(1) > initpop3(0) 'exit if timer current is greater than timer in advance


End Function


