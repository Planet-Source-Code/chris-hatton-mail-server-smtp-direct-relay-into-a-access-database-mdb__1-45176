VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl SMTP 
   ClientHeight    =   7920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11415
   ScaleHeight     =   7920
   ScaleWidth      =   11415
   Begin VB.PictureBox Picture1 
      Height          =   735
      Index           =   1
      Left            =   3480
      Picture         =   "SMTP.ctx":0000
      ScaleHeight     =   675
      ScaleWidth      =   555
      TabIndex        =   11
      Top             =   6840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Index           =   0
      Left            =   2760
      Picture         =   "SMTP.ctx":0442
      ScaleHeight     =   675
      ScaleWidth      =   555
      TabIndex        =   10
      Top             =   6840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   4680
      Top             =   6120
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   7665
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20082
            Object.ToolTipText     =   "Status"
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock SckSMTP 
      Index           =   0
      Left            =   4200
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3480
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SMTP.ctx":0884
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SMTP.ctx":0CD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SMTP.ctx":0E30
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SMTP.ctx":0F8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SMTP.ctx":10E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SMTP.ctx":1536
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SMTP.ctx":1988
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SMTP.ctx":1AE2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   6735
      Left            =   2160
      TabIndex        =   4
      Top             =   0
      Width           =   9225
      Begin MSComctlLib.ListView ListView1 
         Height          =   2385
         Left            =   0
         TabIndex        =   7
         Top             =   360
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   4207
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   14887
         EndProperty
      End
      Begin VB.Frame Frame4 
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   9220
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "SMTP Data Log"
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
            TabIndex        =   6
            Top             =   120
            Width           =   9135
         End
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3945
         Left            =   0
         TabIndex        =   9
         Top             =   2760
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   6959
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "SMTP Message ID"
            Object.Width           =   5362
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Owner"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Size"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Submitted"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.Frame Frame2 
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   2175
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Connections"
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
            TabIndex        =   3
            Top             =   120
            Width           =   1935
         End
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   6360
         Left            =   15
         TabIndex        =   1
         Top             =   345
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   11218
         _Version        =   393217
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
   End
   Begin VB.Menu rootmnu 
      Caption         =   "smtpmnu"
      Visible         =   0   'False
      Begin VB.Menu mnuSend 
         Caption         =   "Release"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "SMTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim sckIndex As Integer
Dim strOut As String
Dim strCMDLNE As String
Public DNS As String
Public DBPATH As String
Dim objmail As New Dictionary
Dim objSMTPITM As New Dictionary
Dim Output As String
Dim RestartTimer As Boolean

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As TRAYCON) As Long

Private Type TRAYCON
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    sTip As String * 64
    End Type
    
    'contants
Private Const IC_ADD = &H0
Private Const IC_MODIFY = &H1
Private Const IC_DELETE = &H2
Private Const IC_MESSAGE = &H1
Private Const IC_ICON = &H2
Private Const IC_TIP = &H4
Private Const IC_DOALL = IC_MESSAGE Or IC_ICON Or IC_TIP
Private Const IC_RESTORE = 9
Private Const IC_MINIMIZE = 6
Private Const IC_MOUSEMOVE = &H200
Private Const IC_LBUTTONDCLK = &H201
Private Const IC_LBUTTONDBLCLK = &H203
Private Const IC_RBUTTONUP = &H205

Private IC As TRAYCON
Dim DC As Long
Public Sub AddIcon()
On Error Resume Next
IC.cbSize = Len(IC)
                IC.hWnd = UserControl.hWnd                     'set form handle as systray handle
                IC.uFlags = IC_DOALL                       'do all features
                IC.uCallbackMessage = IC_MOUSEMOVE         'call back on mouse move
                IC.hIcon = Picture1(1).Picture     'default icon
                IC.sTip = "Hattech Mail Server Spooler" & vbNullChar 'change tooltip
                DC = Shell_NotifyIcon(IC_ADD, IC)          'adds icon to tray
               
End Sub
Public Sub RemoveIcon() 'removes icon from tray
    DC = Shell_NotifyIcon(IC_DELETE, IC)
End Sub
Private Sub ChangeIcon(IconImage As PictureBox) 'usage: changeicon (picture1.picture)
             IC.hIcon = IconImage
             DC = Shell_NotifyIcon(IC_MODIFY, IC)
End Sub
Private Sub Accept_IncomingCalls()
On Error Resume Next
SckSMTP(0).LocalPort = 25 'start our smtp server to listen on port 25
SckSMTP(0).listen
End Sub
Public Sub Stop_SMTPLister()
On Error Resume Next
SckSMTP(0).Close
End Sub
Public Sub Start_SMTPListener()
    Call Accept_IncomingCalls
End Sub
Private Function Get_NextPort() As Integer
Dim i As Integer
On Error Resume Next
For i = 1 To 255 'reuse only 255 Simultaneous connections at once, if above that then allocate extra connections
    If SckSMTP(i).State = sckClosed Then
        Get_NextPort = i
        Exit Function
    End If
Next i
End Function

Private Sub AddConnection(strSession As String, lngPort As Integer, tv As TreeView)
On Error Resume Next
Dim ConnGroup As Node
Dim tvItm As Long
    For tvItm = 1 To tv.Nodes.Count
        If tv.Nodes.Item(tvItm).Text = strSession Then Exit Sub
    Next tvItm
Set ConnGroup = tv.Nodes.Add(, , , strSession, 1)
Set ConnGroup = tv.Nodes.Add(ConnGroup, tvwChild, , lngPort)
End Sub

Private Sub RemoveConnection(lngPort As Integer, tv As TreeView)
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
Private Function Verify_Session(lngPort As Integer, tv As TreeView) As String 'find out what port belongs to what session
Dim i As Long                                                                 'that is logged in
Dim strItem As Variant
On Error GoTo err
    For i = 1 To tv.Nodes.Count
        tv.Nodes.Item(i).Expanded = True
        strItem = tv.Nodes.Item(i).Text
            If strItem = lngPort Then
                Verify_Session = tv.Nodes.Item(i).Parent
            End If
    Next i
Exit Function
err:
End Function
Private Sub SMTP_Que(SessionID As Integer, sckIndex As Integer)  'sessionid is the Index of the email class values
On Error Resume Next
Dim strTemp As String 'external temporary users
If Not objmail(sckIndex).sckState = sckquit Then Exit Sub
Dim rsStream As SMTPSTREAM 'Local
Dim rsStreamEx As SMTPSTREAM 'External
Set rsStreamEx = New SMTPSTREAM
Dim LoUser, ExUser As Integer
DoEvents: DoEvents: DoEvents
    StatusBar1.Panels(1).Text = "Processing Session (" & SessionID & ")"
        For LoUser = 1 To objmail(sckIndex).ColRCPT.Count 'do each recipient indiviually
            If Not InStr(1, objmail(sckIndex).ColRCPT(LoUser), DNS) = 0 Then   'if the user is local then do this
                Set rsStream = New SMTPSTREAM
                Set objSMTPITM(SessionID) = rsStream
                    DoEvents: DoEvents: DoEvents
                    objSMTPITM(SessionID).TempPathFile = objmail(sckIndex).TempDumpFile 'get the temp file from the class to load
                    objSMTPITM(SessionID).StrRCPTTO = objmail(sckIndex).ColRCPT(LoUser) 'recipient
                    objSMTPITM(SessionID).DBPATH = DBPATH 'current database location
                    Call objSMTPITM(SessionID).ImportFlatFile 'process the email into memory
                    Call ShowQueLV(ListView2, SessionID, Split(StrReverse(objmail(sckIndex).TempDumpFile), "\")(0), "Spooled", "administrator", DBFileSize(objmail(sckIndex).TempDumpFile), Time)
                Set rsStream = Nothing
            Else
                    rsStreamEx.ColRCPT.Add objmail(sckIndex).ColRCPT(LoUser)
                End If
        Next LoUser
    If Not rsStreamEx.ColRCPT.Count = 0 Then
        For ExUser = 1 To rsStreamEx.ColRCPT.Count
            rsStreamEx.StrRCPTTO = rsStreamEx.StrRCPTTO & rsStreamEx.ColRCPT(ExUser) & ","
        Next ExUser
        rsStreamEx.DBPATH = DBPATH 'current database location
        rsStreamEx.TempPathFile = objmail(sckIndex).TempDumpFile 'get the temp file from the class to load

        rsStreamEx.StreamTEXT (objmail(sckIndex).BodyData), False 'External
        
    End If
    StatusBar1.Panels(1).Text = "Session Complete (" & SessionID & ")"
    Screen.MousePointer = 0
    Set objmail(sckIndex) = Nothing 'unload instance of email
    Set rsStreamEx = Nothing
    Set SessionMail = Nothing
End Sub
Private Sub SMTPStatus(lv As ListView, lvStatus As String)
On Error Resume Next
    lv.SelectedItem.ListSubItems.Item(1).Text = lvStatus
    lv.refresh
End Sub
Private Sub mnuDelete_Click()
On Error GoTo exitstat
    Call SMTPStatus(ListView2, "Deleting")
    Call DeleteItem(ListView2, ListView2.SelectedItem.Tag)
exitstat:
End Sub
Private Sub LoadSpooler()
On Error Resume Next
Dim fso As FileSystemObject
Set fso = New FileSystemObject
Dim strTemp As String
Dim fsoDrive As Drive
Dim fsoFile As File
Dim fsoFolder As Folder
Dim i As Long
Set Drive = fso.GetDrive(Split(App.Path, ":")(0))
If Drive.IsReady Then
        For Each File In Drive.RootFolder.SubFolders
            If InStr(1, File, App.Path) >= 1 Then
                Set Folder = fso.GetFolder(App.Path)
                      For Each fsoFile In Folder.Files
                        If InStr(1, fsoFile, ".tmp") > 1 Then
                            strTemp = Split(StrReverse(fsoFile), "\")(0)
                           Call ShowQueLV(ListView2, ListView2.ListItems.Count, strTemp, "Paused", "Unknown", DBFileSize(fsoFile), fsoFile.DateCreated)
                        End If
                Next
            End If
        Next
End If
    

Set fsoDrive = Nothing
Set fso = Nothing
Set Folder = Nothing
End Sub

Private Sub sckSMTP_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error GoTo SMTPCon
If Get_NextPort = 0 Then 'reuse a available port perrably a free lower port
    sckIndex = sckIndex + 1
Else
    sckIndex = Get_NextPort
End If

Load SckSMTP(sckIndex)
SckSMTP(sckIndex).accept requestID
SckSMTP(sckIndex).SendData "220 " & SckSMTP(sckIndex).LocalHostName & "." & LCase(DNS) & " Hattech Technologies SMTP Mail Service, Version: " & App.Major & "." & App.Minor & "." & App.Revision & " ready at " & Date & " " & Time & vbCrLf
Set SessionMail = New SMTPGRP
Set objmail(sckIndex) = SessionMail
    objmail(sckIndex).DNS = DNS
    objmail(sckIndex).RemoteHost = SckSMTP(Index).RemoteHostIP
    objmail(sckIndex).RemoteHostName = SckSMTP(Index).LocalHostName
    Call AddConnection(SckSMTP(Index).RemoteHostIP & "(" & sckIndex & ")", sckIndex, TreeView1) 'add the new connection to the tree view control
    Randomize
    objmail(sckIndex).SMTPID = Int((8192# * Rnd))
Exit Sub
SMTPCon:
'MsgBox "Err: 604 " & vbNewLine & err.Description, vbCritical
'StatusBar2.Panels(0).Text = "Last Error: 602"
End Sub

Private Sub ShowQueLV(lv As ListView, Index As Integer, MsgID, strStatus, strOwner, strSize, strSubmitted As String)
    On Error Resume Next
    Dim lvitem As ListItem
    Dim lvsubitem As ListSubItem
        Set lvitem = lv.ListItems.Add(, , StrReverse(MsgID))
        Set lvsubitem = lvitem.ListSubItems.Add(, , strStatus)
        lv.ListItems.Item(lv.ListItems.Count).Tag = Index
        With lvitem
            .ListSubItems.Add , , strOwner
            .ListSubItems.Add , , "" & Split(strSize, ".")(0) + " KB"
            .ListSubItems.Add , , strSubmitted
        End With
        

End Sub
Private Sub ListView1_DblClick()
On Error Resume Next
MsgBox ListView1.SelectedItem.Text
End Sub

Private Sub ListView2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    PopupMenu rootmnu
End If
End Sub

Private Sub mnuSend_Click()
On Error GoTo exitstat
    Call SMTPStatus(ListView2, "Sending") 'set the status to sending
    Call StartSend(ListView2, ListView2.SelectedItem.Tag) 'actually start polling start transfer
exitstat:
End Sub
Private Sub DeleteItem(lv As ListView, Index As Integer)
On Error Resume Next
If lv.SelectedItem.Selected = True Then
    Kill objSMTPITM(SessionID).TempPathFile 'remove temp file
    Call lv.ListItems.Remove(lv.SelectedItem.Index)
    Set objSMTPITM(Index) = Nothing 'unload the smtp stream item (note the smtpitem contains the mail item object)
End If

End Sub
Private Function MailTo(objBody As String)
On Error Resume Next
    Dim strTemp(0 To 1) As String
        strTemp(0) = Split(objBody, "To: ")(1)
        strTemp(1) = Split(strTemp(0), vbCrLf)(0)
    MailTo = "<" & Split(strTemp(1), "<")(1)
End Function

Private Sub StartSend(lv As ListView, Index As Integer)
On Error Resume Next
Call ChangeIcon(Picture1(0)) ' change light to sending
If lv.SelectedItem.ListSubItems(2).Text = "Unknown" Then
    Dim FileStream As New SMTPSTREAM
    FileStream.DBPATH = DBPATH
    FileStream.TempPathFile = App.Path & "\" & lv.SelectedItem.Text
    FileStream.ImportFlatFile
    FileStream.StrRCPTTO = MailTo(FileStream.objFlatFile)
    If Split(FileStream.StrRCPTTO, "@")(1) = DNS & ">" Then
        FileStream.StreamTEXT FileStream.objFlatFile, True
        Else
        FileStream.StreamTEXT FileStream.objFlatFile, False
    End If
    Kill FileStream.TempPathFile 'remove temp file
    Call lv.ListItems.Remove(lv.SelectedItem.Index)
Exit Sub

End If


If lv.SelectedItem.Selected = True Then
    objSMTPITM(Index).StreamTEXT (objSMTPITM(Index).objFlatFile), True  'actual object email of data, and now stream it to ADO
    Call lv.ListItems.Remove(lv.SelectedItem.Index)
    Set objSMTPITM(Index) = Nothing 'unload the smtp stream item (note the smtpitem contains the mail item object)
    Kill objSMTPITM(SessionID).TempPathFile 'remove temp file
    Call ChangeIcon(Picture1(1)) 'switch off the light

End If

End Sub

Private Sub SckSMTP_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim inData As String
SckSMTP(Index).GETDATA inData
Debug.Print "length " & Len(inData)
If Len(inData) >= 8192 Then SckSMTP(Index).SendData ""
On Error Resume Next

If objmail(Index).sckState = sckdata Then
    objmail(Index).BodyData = inData
End If

If objmail(Index).sckState = sckquit Or Mid$(UCase(inData), 1, 4) = "QUIT" Then
quit:
        
        SckSMTP(Index).Close
        Unload SckSMTP(Index)
        Call RemoveConnection(Index, TreeView1)
        Call SMTP_Que(objmail(sckIndex).SMTPID, Index) 'all sockets closed place it in the smtp que

        objmail(sckIndex).SMTPID = 0


    Exit Sub
End If


  If InStr(1, inData, vbCrLf) > 0 Then
    If objmail(Index).sckState = sckhelo Then ListView1.ListItems.Add , , "<-----------------Message-Break----------------->", , 4
    If InStr(1, inData, vbCrLf) > 0 Then ListView1.ListItems.Add , , "TCP " & Index & " Recv: " & inData, , 7
        objmail(Index).InPutParser = strCMDLNE & inData 'input to smtp instance
            Call objmail(Index).Parser
                Output = objmail(Index).strOut & vbCrLf
                        If Not Len(objmail(Index).strOut) = 0 Then SckSMTP(Index).SendData LTrim$(Output)
                        If Not Len(objmail(Index).strOut) = 0 Then ListView1.ListItems.Add , , "TCP " & Index & " Send : " & objmail(Index).strOut, , 8
                        
            strCMDLNE = ""
        Else
            strCMDLNE = strCMDLNE & inData
        Exit Sub
    End If

End Sub


Private Sub ReleaseSpooler(lv As ListView)
 On Error Resume Next
    Dim lngTotal As Long
    Dim i As Long
    lngTotal = lv.ListItems.Count
    If lngTotal = 0 Then Exit Sub
    
    If lngTotal >= 1 Then
        If RestartTimer = True Then GoTo StartTrans
        RestartTimer = True
        Exit Sub
    End If

StartTrans:
        For i = 1 To lngTotal
            If lv.ListItems(i).ListSubItems(1).Text = "Sending" Then Exit Sub 'if already one in the list sending the do nothing else
            If lv.ListItems(i).ListSubItems(1).Text = "Spooled" Then 'this one is ready to go
                lv.ListItems(i).Selected = True
                Call SMTPStatus(ListView2, "Sending") 'set the status to sending
                Call StartSend(ListView2, ListView2.SelectedItem.Tag) 'actually start polling start transfer
                lngTotal = lv.ListItems.Count
            End If
        
        Next i
               RestartTimer = False
               Call ChangeIcon(Picture1(1)) 'switch off the light

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Call ReleaseSpooler(ListView2)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim Sendmsg As Long

If UserControl.ScaleMode = vbPixels Then 'switch between pixels or twips so we can find the icon to click on
        Sendmsg = X
    Else
       Sendmsg = X / Screen.TwipsPerPixelX
    End If
        
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
Call LoadSpooler
End Sub

Private Sub UserControl_Show()
Call AddIcon
End Sub

Private Sub UserControl_Terminate()
Call RemoveIcon
End Sub
