VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl DBSMTPSNDER 
   ClientHeight    =   7770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10185
   ScaleHeight     =   7770
   ScaleWidth      =   10185
   Begin VB.Frame Frame1 
      Height          =   7695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8500
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "DBSMTPSNDER.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2880
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   2520
         Width           =   9975
      End
      Begin VB.Frame Frame2 
         Height          =   4815
         Left            =   0
         TabIndex        =   1
         Top             =   2760
         Width           =   9975
         Begin VB.Timer Timer3 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   1680
            Top             =   720
         End
         Begin VB.Timer Timer2 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   1200
            Top             =   720
         End
         Begin VB.Timer Timer1 
            Interval        =   10000
            Left            =   720
            Top             =   720
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   120
            Top             =   600
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   7
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "DBSMTPSNDER.ctx":030A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "DBSMTPSNDER.ctx":075C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "DBSMTPSNDER.ctx":2F0E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "DBSMTPSNDER.ctx":3360
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "DBSMTPSNDER.ctx":37B2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "DBSMTPSNDER.ctx":3ACC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "DBSMTPSNDER.ctx":3DE6
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   4665
            Left            =   15
            TabIndex        =   2
            Top             =   120
            Width           =   9930
            _ExtentX        =   17515
            _ExtentY        =   8229
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   18415
            EndProperty
         End
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   2475
         Left            =   15
         TabIndex        =   3
         Top             =   120
         Width           =   9930
         _ExtentX        =   17515
         _ExtentY        =   4366
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         SingleSel       =   -1  'True
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
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   8160
         TabIndex        =   6
         Text            =   "DNS SERVER"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin MSWinsockLib.Winsock sckSMTPSND 
      Index           =   0
      Left            =   8640
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu mnupurge 
         Caption         =   "&Purge Message"
      End
      Begin VB.Menu mnupurgeall 
         Caption         =   "Pur&ge all BAD Messages"
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView 
         Caption         =   "&View Message"
      End
      Begin VB.Menu mnucancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "DBSMTPSNDER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim cn As ADODB.Connection
Public DNS As String 'hatton.net domain name service
Public DBPATH As String
Dim objSMTP As New Dictionary
Private Declare Function GetTickCount Lib "kernel32" () As Long 'system timer
Dim initret4, initret3, initret2, initret1 As Single
Dim initsmtp(0 To 1) As Single
Dim MAXBUFFER As Long
Public SMTPRELAY As Boolean
Dim SENDQUIT As Boolean
Dim MaxconRetrys As Integer 'maxium connection retrys
Dim ConRetrys As Integer
Private Enum SmtpSend
    helo = 0
    mail = 1
    RCPT = 2
    Data = 3
End Enum
Dim smtpparse As SmtpSend
    Dim vRs As New ADODB.Recordset 'public recordset so we can interact with it publically.
    Dim vRSCancel As Boolean    'cancel the lookup for the current message (listview)
    Public SMTPCon As Boolean 'is the connector status
Private Const DNS_RECURSION As Byte = 1

Private Type DNS_HEADER
    qryID As Integer
    options As Byte
    response As Byte
    qdcount As Integer
    ancount As Integer
    nscount As Integer
    arcount As Integer
End Type

' Registry data types
Const REG_NONE = 0&
Const REG_SZ = 1&
Const REG_EXPAND_SZ = 2&
Const REG_BINARY = 3&
Const REG_DWORD = 4&
Const REG_DWORD_LITTLE_ENDIAN = 4&
Const REG_DWORD_BIG_ENDIAN = 5&
Const REG_LINK = 6&
Const REG_MULTI_SZ = 7&
Const REG_RESOURCE_LIST = 8&
Const REG_FULL_RESOURCE_DESCRIPTOR = 9&
Const REG_RESOURCE_REQUIREMENTS_LIST = 10&

' Registry access types
Const KEY_QUERY_VALUE = &H1&
Const KEY_SET_VALUE = &H2&
Const KEY_CREATE_SUB_KEY = &H4&
Const KEY_ENUMERATE_SUB_KEYS = &H8&
Const KEY_NOTIFY = &H10&
Const KEY_CREATE_LINK = &H20&
Const READ_CONTROL = &H20000
Const WRITE_DAC = &H40000
Const WRITE_OWNER = &H80000
Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_READ = READ_CONTROL
Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Const KEY_EXECUTE = KEY_READ

' Registry keys
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006


Const ERROR_SUCCESS = 0&

' Registry access functions
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

Dim PDNS As String
Dim sDNS As Variant
Dim DNSDIRECT As String


Private Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Private Sub GetDNSInfo()
    Dim hKey As Long
    Dim hError As Long
    Dim sdhcpBuffer As String
    Dim sBuffer As String
    Dim sFinalBuff As String
    
    sdhcpBuffer = Space(1000)
    sBuffer = Space(1000)
    If (RegOpenKeyEx(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Tcpip\Parameters", 0, KEY_READ, hKey) = ERROR_SUCCESS) Then
        RegQueryValueEx hKey, "DhcpNameServer", 0, REG_SZ, sdhcpBuffer, 1000
        RegQueryValueEx hKey, "NameServer", 0, REG_SZ, sBuffer, 1000
        RegCloseKey hKey
            
        sFinalBuff = Trim(StripTerminator(sBuffer) & " " & StripTerminator(sdhcpBuffer))
        sDNS = Split(sFinalBuff, " ")
    End If
End Sub
Private Sub ParseName(dnsReply() As Byte, iNdx As Integer, sName As String)
    Dim iCompress As Integer
    Dim iChCount As Integer
        
    While (dnsReply(iNdx) <> 0)
        iChCount = dnsReply(iNdx)
        If (iChCount = 192) Then
            iCompress = dnsReply(iNdx + 1)
            ParseName dnsReply(), iCompress, sName
            iNdx = iNdx + 2
            Exit Sub
        End If
        
        iNdx = iNdx + 1
        While (iChCount)
            sName = sName + Chr(dnsReply(iNdx))
            iChCount = iChCount - 1
            iNdx = iNdx + 1
        Wend
        If (dnsReply(iNdx) <> 0) Then sName = sName + "."
    Wend
End Sub
Private Function GetMXName(dnsReply() As Byte, iNdx As Integer, iAnCount As Integer) As String
    Dim iChCount As Integer
    Dim sTemp As String
    
    Dim iBestPref As Integer
    Dim sBestMX As String
    
    iBestPref = -1
    
    ParseName dnsReply(), iNdx, sTemp
    iNdx = iNdx + 2
    iNdx = iNdx + 6
    
    While (iAnCount)
        If (dnsReply(iNdx) = 15) Then
            Dim sName As String
            Dim iPref As Integer
            
            sName = ""
            
            iNdx = iNdx + 1 + 6
            
            iNdx = iNdx + 2
            
            MemCopy iPref, dnsReply(iNdx), 2
            iPref = ntohs(iPref)
            iNdx = iNdx + 2
            
            ParseName dnsReply(), iNdx, sName
'            lstMX.AddItem "[Preference = " & iPref & "] " & sName
            
            If (iBestPref = -1 Or iPref < iBestPref) Then
                iBestPref = iPref
                sBestMX = sName
            End If
            
            iNdx = iNdx + 3
        Else
            GetMXName = sBestMX
            Exit Function
        End If
        iAnCount = iAnCount - 1
    Wend
    
    GetMXName = sBestMX
End Function

Private Function MakeQName(sDomain As String) As String
    Dim iQCount As Integer
    Dim iNdx As Integer
    Dim iCount As Integer
    Dim sQName As String
    Dim sDotName As String
    Dim sChar As String
    
    iNdx = 1
    iQCount = 0
    iCount = Len(sDomain)

    While (iNdx <= iCount)

        sChar = Mid(sDomain, iNdx, 1)
  
        If (sChar = ".") Then
            sQName = sQName & Chr(iQCount) & sDotName
            iQCount = 0
            sDotName = ""
        Else
            sDotName = sDotName + sChar
            iQCount = iQCount + 1
        End If
        iNdx = iNdx + 1
    Wend
    
    sQName = sQName & Chr(iQCount) & sDotName
    
    MakeQName = sQName
End Function

Private Function MX_Query() As String
    Dim StartupData As WSADataType
    Dim SocketBuffer As sockaddr
    Dim IpAddr As Long
    Dim iRC As Integer
    Dim dnsHead As DNS_HEADER
    Dim iSock As Integer

    ' Initialize the Winsocket
    iRC = WSAStartup(&H101, StartupData)
    'iRC = WSAStartup(&H101, StartupData)
    If iRC = SOCKET_ERROR Then Exit Function
    
    ' Create a socket
    iSock = socket(AF_INET, SOCK_DGRAM, 0)
    If iSock = SOCKET_ERROR Then Exit Function
    
    IpAddr = GetHostByNameAlias(Text1)
    If IpAddr = -1 Then Exit Function
    
    ' Setup the connnection parameters
    SocketBuffer.sin_family = AF_INET
    SocketBuffer.sin_port = htons(53)
    SocketBuffer.sin_addr = IpAddr
    SocketBuffer.sin_zero = String$(8, 0)
    
    ' Set the DNS parameters
    dnsHead.qryID = htons(&H11DF)
    dnsHead.options = DNS_RECURSION
    dnsHead.qdcount = htons(1)
    dnsHead.ancount = 0
    dnsHead.nscount = 0
    dnsHead.arcount = 0

    Dim dnsQuery() As Byte
    Dim sQName As String
    Dim dnsQueryNdx As Integer
    Dim iTemp As Integer
    Dim iNdx As Integer
    dnsQueryNdx = 0
    
    ReDim dnsQuery(4000)

    MemCopy dnsQuery(dnsQueryNdx), dnsHead, 12
    dnsQueryNdx = dnsQueryNdx + 12
    
    
    sQName = MakeQName(PDNS)
    iNdx = 0
    While (iNdx < Len(sQName))
        dnsQuery(dnsQueryNdx + iNdx) = Asc(Mid(sQName, iNdx + 1, 1))
        iNdx = iNdx + 1
    Wend

    dnsQueryNdx = dnsQueryNdx + Len(sQName)
    
    dnsQuery(dnsQueryNdx) = &H0
    dnsQueryNdx = dnsQueryNdx + 1
    
    iTemp = htons(15)
    MemCopy dnsQuery(dnsQueryNdx), iTemp, Len(iTemp)
    dnsQueryNdx = dnsQueryNdx + Len(iTemp)
    
    iTemp = htons(1)
    MemCopy dnsQuery(dnsQueryNdx), iTemp, Len(iTemp)
    dnsQueryNdx = dnsQueryNdx + Len(iTemp)
    
    ReDim Preserve dnsQuery(dnsQueryNdx - 1)
    iRC = sendto(iSock, dnsQuery(0), dnsQueryNdx + 1, 0, SocketBuffer, Len(SocketBuffer))
    If (iRC = SOCKET_ERROR) Then
        MsgBox "Problem sending"
        Exit Function
    End If
    

    Dim dnsReply(2048) As Byte
     iRC = recvfrom(iSock, dnsReply(0), 2048, 0, SocketBuffer, Len(SocketBuffer))
    If (iRC = SOCKET_ERROR) Then
        MsgBox "Problem receiving"
        Exit Function
    End If
    
    Dim iAnCount As Integer
    MemCopy iAnCount, dnsReply(6), 2
    iAnCount = ntohs(iAnCount)
    MX_Query = GetMXName(dnsReply(), 12, iAnCount)
End Function

Private Sub GETLOOKUP()
    Dim sMX As String
        sMX = MX_Query
        If (Len(sMX) > 0) Then
            DNSDIRECT = sMX
        Else
            Call Writetolog("No mail routing information found, resorting back to default SMTP Relay for SMTP Lookup " & PDNS, " Port: unknown")
        End If

End Sub
Public Sub ADO_Connect()
Set cn = New ADODB.Connection
        cn.Provider = "Microsoft.Jet.OLEDB.4.0"
        cn.Open DBPATH
        cn.CursorLocation = adUseClient
        Call ReadOptions
End Sub
Public Sub ReadOptions() 'add to global options
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
    rs.Open "select * from settings where configurationset = 1", cn, adOpenKeyset, adLockOptimistic
        MaxconRetrys = rs!ConnectionRetrys
        If rs!SMTPRELAY = True Then SMTPRELAY = True
        If rs!SMTPRELAY = False Then SMTPRELAY = False
        MAXBUFFER = rs!buffer
        
    rs.Close
Set rs = Nothing

End Sub

Private Sub ADO_Disconnect()
    cn.Close
    Set cn = Nothing
End Sub
Private Sub NotifyUserBADMSG(clsIndex As Integer, strRCPT, StrSubject As String)
    
    Dim i, j As Long
    Dim MSG_RTN As String
                'message of failure
          MSG_RTN = "Message-ID: <001c01c2565b$b8730140$0d08a2c0@" & LCase(DNS) & "> " & vbCrLf & _
                    "Subject: Undeliverable: " & StrSubject & vbCrLf & _
                    "From: Mail Administrator" & vbCrLf & _
                    "To: " & strRCPT & vbCrLf & vbCrLf & _
                    "Your message did not reach some or all of the intended recipients." & vbCrLf & _
                    "Subject: " & StrSubject & vbCrLf & _
                    "This is a automated message PLEASE DO NOT REPLY. will try again in 1 hour" & vbCrLf & _
                    " " & _
                    "" & vbCrLf & vbCrLf & "." & vbCrLf
                    initret4 = GetTickCount / 1000 + 3600
                    Timer3.Enabled = True 'set the timer for 1 hour
                'wrapp the object
               Call objclsSMTP(clsIndex, "<administrator@" & DNS & ">", strRCPT, MSG_RTN, "127.0.0.1")


End Sub
Private Sub objclsSMTP(clsIndex As Integer, objstrFrom, objstrRcpt, objstrBody, objstrSMTP As String)
    Set objSMTPRelay = New SMTPRELAY
    Set objSMTP(clsIndex) = objSMTPRelay
    Dim sck As Long
    On Error Resume Next
                
                If InStr(1, objstrRcpt, ",") = 0 Then objSMTP(clsIndex).ColRCPT.Add objstrRcpt 'if only one recipient then add him
                
                For i = 1 To UBound(Split(objstrRcpt, ",")) 'loop 1 to x amount of recipents
                    objSMTP(clsIndex).ColRCPT.Add Split(objstrRcpt, ",")(i - 1) 'process through all recipients and add them to the collection
                Next i 'next recipient
                    
    Load sckSMTPSND(clsIndex) 'load the requested socket number and winsock. if fails good idea to get one at random
        If Not Len(err.Description) = 0 Then
            Load sckSMTPSND(clsIndex)  'load a new socket corresponding to the class
                Randomize (1000) 'if for any reason the socket is already in use just choose a random socket under 100, or 100000 or whatever.
                sck = Int(1000 * Rnd)
                Load sckSMTPSND(sck)
            err.Clear
        End If
                    
            objSMTP(clsIndex).objMailFrom = objstrFrom 'set the reverse path of the sender
            objSMTP(clsIndex).objSMTP = objstrSMTP            'set the smtp provider
            objSMTP(clsIndex).objBody = objstrBody 'set the body of email in the class
          
            sckSMTPSND(clsIndex).Close 'close the session before connecting, good winsock practice.
            ConRetrys = 0
            
            
            SumitConnection clsIndex, objSMTP(clsIndex).objSMTP 'read the smtp value and connect to it via the winsock
            
            If SumitConnection(clsIndex, objSMTP(clsIndex).objSMTP) = True Then
                Timer1.Enabled = True
            End If


End Sub
Private Function SumitConnection(inxConnection As Integer, SMTPSERVER As String) As Boolean
On Error GoTo errcon
If ConRetrys = MaxconRetrys Then Exit Function
DoEvents
sckSMTPSND(inxConnection).connect SMTPSERVER, "25"
Call refresh
ConRetrys = ConRetrys + 1

If ConnectionWait(10, 0) = True Then 'it was successful connection
    SumitConnection = True
    ConRetrys = 0 'reset counter
Else
    sckSMTPSND(0).Close
    SumitConnection inxConnection, SMTPSERVER 'retry connecting again
End If

Exit Function
errcon:

'MsgBox "Err: 600 " & vbNewLine & err.Description, vbCritical
StatusBar2.Panels(1).Text = "Last Error: 600"
End Function
Private Function ConnectionWait(seconds As Long, sckIndex As Integer) As Boolean
initsmtp(0) = GetTickCount / 1000 + seconds 'timer in advance

Do

    initsmtp(1) = GetTickCount / 1000 'current timer
    DoEvents: DoEvents: DoEvents
'    Debug.Print sckSMTPSND(sckIndex).State
    
    If sckSMTPSND(sckIndex).State = 9 Then sckSMTPSND(sckIndex).Close: Exit Function 'socket error
    If sckSMTPSND(sckIndex).State = 8 Then: sckSMTPSND(sckIndex).Close: Exit Function
    If sckSMTPSND(sckIndex).State = 7 Then: ConnectionWait = True: Exit Function
    
Loop Until initsmtp(1) > initsmtp(0) 'exit if timer current is greater than timer in advance

If sckSMTPSND(sckIndex).State = 0 Then: sckSMTPSND(sckIndex).Close: Exit Function
End Function

Public Sub SendMail(clsIndex As Integer, QUETYPE As String) 'clsindex has to be unique for every instance of the class and socket as they are linked.
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim i, j As Long
    Dim objstrRcpt, strRCPT As String
        Timer1.Enabled = False 'stop the timer intill the entire message has been sent, we want to conserve bandwith
        rs.Open "select * from " & QUETYPE, cn, adOpenKeyset, adLockReadOnly
            If rs.RecordCount = 0 Then GoTo EmptyQue
            objstrRcpt = rs!RCPT
               If SMTPRELAY = True Then 'smtp relay connection
                    Call objclsSMTP(clsIndex, MailFrom(rs!body), objstrRcpt, rs!body, GetSMTPProvider(MailFrom(rs!body)))
                    If Not Len(err.Description) = 0 Then err.Clear: Call Writetolog("Error Trying to read data and for providing the SMTP object, MAIL TO:, RCPT TO:, DATA," & "Class: " & clsIndex, clsIndex)
                
                Else 'smtp Direct Connection
                    
                    If InStr(1, objstrRcpt, ",") = 0 Then
                        Call objclsSMTP(clsIndex, MailFrom(rs!body), objstrRcpt, rs!body, GetSMTPProvider(MailFrom(rs!body)))
                        If Not Len(err.Description) = 0 Then err.Clear: Call Writetolog("Error Trying to read data and for providing the SMTP object, MAIL TO:, RCPT TO:, DATA," & "Class: " & clsIndex, clsIndex)
                    Else
                        For i = 1 To UBound(Split(objstrRcpt, ",")) 'loop 1 to x amount of recipents
                            strRCPT = Split(objstrRcpt, ",")(i - 1) 'process through all recipients and add them to the collection
                            PDNS = Split(strRCPT, "@")(1): If InStr(1, PDNS, ">") > 0 Then PDNS = Split(PDNS, ">")(0)
                            Call GETLOOKUP 'if no routing info found do the default smtp connection
                            If Len(DNSDIRECT) = 0 Then DNSDIRECT = GetSMTPProvider(MailFrom(rs!body))
                            Call objclsSMTP(clsIndex, MailFrom(rs!body), strRCPT, rs!body, DNSDIRECT)
                            If Not Len(err.Description) = 0 Then err.Clear: Call Writetolog("Error Trying to read data and for providing the SMTP object, MAIL TO:, RCPT TO:, DATA," & "Class: " & clsIndex, clsIndex)
                        Next i 'next recipient
                    End If
                    
               End If
        
        
        rs.Close
    Set rs = Nothing
    Exit Sub
EmptyQue:
    
    rs.Close
    Set rs = Nothing
    
End Sub
Private Function GetSMTPProvider(strUserID As String) ' return the users smtp server (his/her isp)
On Error Resume Next
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
    Dim strQry As String
    If Not InStr(1, strUserID, "@") = 0 Then strUserID = Split(strUserID, "@")(0) 'extract only the username from the address if it is a email address value
    strQry = "select user, externalsmtp from Profiles where user = " & Chr(34) & strUserID & Chr(34) 'select only the users record
    rs.Open strQry, cn, adOpenStatic, adLockReadOnly 'just open it static, no changes need to be changed
        GetSMTPProvider = rs!externalsmtp 'return the local users smtp server
       If Not Len(err.Description) = 0 Then Call Writetolog("Error Trying to read data and for providing the SMTP object DB:" & err.Description, 100)
       If Len(GetSMTPProvider) = 0 Then Call Writetolog("Cannot Find SMTP Provider for " & strUserID, 99)
    rs.Close
Set rs = Nothing

End Function

Private Sub Writetolog(txtDescription As String, FileNumber As Integer)
On Error Resume Next
    Open App.Path & "\SMTPRELAY.log" For Append As #FileNumber
    Print #FileNumber, Date & " " & Time & " " & txtDescription
    Close #FileNumber
    
End Sub
Public Sub RetryBadQUE(strMsgID As String)
On Error Resume Next
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim strQry As String
    strQry = "select * from SMTPQUE where body = " & Chr(34) & MsgID(strMsgID) & Chr(34)
    rs.Open strQry, cn, adOpenKeyset, adLockOptimistic
        SendMail rs.RecordCount, "SMTPBADQUE"
    rs.Close
Set rs = Nothing
If Not Len(err.Description) = 0 Then
    MsgBox "Error in Message " & vbCrLf & err.Description, vbCritical, "SMTP BAD QUE"
End If
End Sub

Private Sub Command1_Click()
Call CancelRSview
End Sub

Public Sub mnupurge_Click()
On Error Resume Next
Dim que As String
Dim response As String
If Not InStr(1, TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Text, "Message-ID:") = 1 Then Exit Sub
que = TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Tag

response = MsgBox("Delete Message?", vbQuestion + vbYesNo, "Purge Email Message")

If response = vbYes Then
    If que = "BAD" Then
        Call RemoveSentMsg(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Text, "SMTPBADQUE")
    Else
        Call RemoveSentMsg(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Text, "SMTPQUE")
    End If
End If
End Sub

Public Sub mnupurgeall_Click()
Dim que As String
Dim response As String
Dim i As Long
response = MsgBox("Delete ALL Message(s)?", vbQuestion + vbYesNo, "Purging Email Message(s)")
If response = vbYes Then
    For i = 1 To TreeView1.Nodes.Count
        If Not InStr(1, TreeView1.Nodes.Item(i).Text, "Message-ID:") = 1 Then GoTo Skip
         que = TreeView1.Nodes.Item(i).Tag
         If que = "BAD" Then
           Call RemoveSentMsg(TreeView1.Nodes.Item(i).Text, "SMTPBADQUE")
         End If
    
Skip:
    Next i

End If
Call refresh




End Sub


Private Sub sckSMTPSND_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Dim strInData As String
Dim initret As Long
Dim errRcpt As String
Dim i As Long
Dim j As Long
Dim k As Long
sckSMTPSND(Index).GETDATA strInData
    If InStr(1, strInData, vbCrLf) <> 0 Then strInData = Split(strInData, vbCrLf)(0)
    initret = Split(strInData, " ")(0)
        Select Case smtpparse
            Case helo 'hello complete
                If initret = "220" Then
                    smtpparse = mail
                    sckSMTPSND(Index).SendData "Mail From: " & objSMTP(Index).objMailFrom & vbCrLf
                Else
                    Call Writetolog(MsgID(objSMTP(Index).objBody) & " Error in Connection State: " & strInData, Index)  'if any errors it will be a type of connection error
                    objSMTP(Index).Error = True
                End If
            Case mail
                If initret = "250" Then 'mail from OK
                    If objSMTP(Index).ColRCPT.Count = 0 Then objSMTP(Index).Error = True: Call Writetolog(MsgID(objSMTP(Index).objBody) & " No recipients specified " & strInData, Index): GoTo Distroy
                        For i = 1 To objSMTP(Index).ColRCPT.Count 'send all the mail to all recipients
                            sckSMTPSND(Index).SendData "RCPT TO: " & objSMTP(Index).ColRCPT(i) & vbCrLf 'each recipient
                            k = i 'track the last recipient in the collection so we know if there was an error
                        Next i
                    smtpparse = RCPT 'all recipients are sent move onto the data
                Else
                    Call Writetolog(MsgID(objSMTP(Index).objBody) & " Error in Sender: " & strInData, Index) 'error in MAIL FROM
                    objSMTP(Index).Error = True
                End If
            Case RCPT 'recipients complete send for data
                If initret = "250" Then ' recipients were ok
                    sckSMTPSND(Index).SendData "DATA " & vbCrLf
                    smtpparse = Data
                Else
                   Call Writetolog(MsgID(objSMTP(Index).objBody) & " Error in RCPT Command: " & strInData, Index) '
                   objSMTP(Index).Error = True 'if error has occouried then move this message to the bad list
                End If
            Case Data
                If initret = "354" Then 'ok pokey signal, lets send data!
                     strBody = objSMTP(Index).objBody
                     For j = 1 To Len(strBody) Step MAXBUFFER 'send 4096 bytes at per chunk good for 10megabit network
                        strsend = Mid$(strBody, j, MAXBUFFER) 'use 8192 for 100mb for faster connections, um, 2048 for slower
                            sckSMTPSND(Index).SendData strsend 'connections like modem
                            Debug.Print strsend
                     Next j
                ElseIf initret = "250" Then 'will need to wait atleast 200 millseconds before unloading the socket
                        sckSMTPSND(Index).SendData "quit" & vbCrLf 'or else it will not send the quit command.
                        SENDQUIT = False
                        Timer2.Enabled = True
                        initret2 = GetTickCount / 1000 + 20
                        Do
                        
                        DoEvents: DoEvents
                        Loop Until SENDQUIT = True
                        
                        GoTo Distroy
                Else
               
                   If Not initret = "221" Then
                        If Not initret = 0 Then
                            Call Writetolog(MsgID(objSMTP(Index).objBody) & " Failed receiving DATA command to send this message: " & strInData, Index) '
                          '  objSMTP(Index).Error = True 'if error has occouried then move this message to the bad list
                        End If
                   End If
                End If

        End Select
    Debug.Print strInData
If objSMTP(Index).Error = True Then GoTo Distroy
Exit Sub
Distroy:
If objSMTP(Index).Error = False Then 'if theres no errors then log it anyway
    Call Writetolog(MsgID(objSMTP(Index).objBody) & " Message Sent", Index)
    Call RemoveSentMsg(MsgID(objSMTP(Index).objBody), "SMTPQUE") 'remove the sent message
Else
    Call MoveMsg(MsgID(objSMTP(Index).objBody), "SMTPQUE", "SMTPBADQUE") 'move bad message to badque
    'If Len(objSMTP(Index).ColRCPT(k)) = 0 Then errRcpt = "<MailServerError@" & LCase(DNS) & ">" Else errRcpt = objSMTP(Index).ColRCPT(k)
    Call NotifyUserBADMSG(100, "<" & objSMTP(Index).objMailFrom & ">", ExtractSubject(objSMTP(Index).objBody))   'notify user of bad message, via phantum smtp session
End If
smtpparse = helo 'reset to for the next msg
Debug.Print "Socket " & Index & " is closed and unloaded"
sckSMTPSND(Index).Close
Unload sckSMTPSND(Index)
Timer1.Enabled = True 'message finish now check for more, and send them too.

Set objSMTP(Index) = Nothing 'distroy the phantom smtp connector
Set objSMTPRelay = Nothing
Call refresh 'refresh the current view
End Sub
Private Function ExtractSubject(objBody As String)
    Dim strTemp(0 To 1) As String
        strTemp(0) = Split(LCase(objBody), LCase("Subject: "))(1)
         ExtractSubject = Split(strTemp(0), vbCrLf)(0)
   
End Function


Private Sub RemoveSentMsg(strMsgID, StrQueue As Variant) 'find the successfully sent message and remove from the DB smtpQUE.
On Error Resume Next
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim strQry As String
Dim i As Long
    strQry = "select * from " & StrQueue
    rs.Open strQry, cn, adOpenKeyset, adLockOptimistic
        For i = 1 To rs.RecordCount 'loop through all messages and try to find the specified message
                If InStr(1, strMsgID, MsgID(rs!body)) = 1 Then ' ok found the right message in the db
                    rs.Delete
                    Exit For
                End If
        Next i
    rs.Update
    rs.Close
Set rs = Nothing
End Sub
Public Sub ProcessBAD()
Dim j As Long
    For j = 1 To TreeView1.Nodes.Count
        Screen.MousePointer = 11
            If TreeView1.Nodes.Item(j).Tag = "BAD" Then
                Call MoveMsg(TreeView1.Nodes(j).Text, "SMTPBADQUE", "SMTPQUE") 'move message back to the smtp que for delivery
            End If
        Screen.MousePointer = 0
    Next j
End Sub
Private Sub MoveMsg(strMsgID As String, strQueueSRC, strQueueDEST As String) 'move messages back and forward from both queue's
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim j As Long
Dim strQry As String
    strQry = "select * from " & strQueueSRC
    rs.Open strQry, cn, adOpenKeyset, adLockOptimistic
       For j = 1 To rs.RecordCount
        If InStr(1, strMsgID, Split(rs!body, vbCrLf)(0)) = 1 Then 'get the message and copy it to smtpque, and delete this message
                Call CopyMsg(rs!Date, rs!RCPT, rs!body, strMsgID, strQueueSRC, strQueueDEST, rs!octat) 'get the selected message and copy it into memory and transfer it to the destination table
            Exit For
        End If
       Next j
    rs.Close
Set rs = Nothing
Call refresh
End Sub
Private Sub CopyMsg(strDate, RCPT, body, strMsgID, strSRC, strDEST As String, Octet As Long)  'copy's the db values into variables
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim strQry As String
    strQry = "select * from " & strDEST 'this message is the Destination
    rs.Open strQry, cn, adOpenKeyset, adLockOptimistic
       rs.AddNew
        rs!Date = "" & strDate
        rs!RCPT = "" & RCPT
        rs!body = "" & body
        rs!octat = "" & Octet
       rs.Update
    rs.Close
Set rs = Nothing

If Len(err.Description) = 0 Then 'query the removeMsg on the messageID and Source Table
    Call RemoveSentMsg(strMsgID, strSRC) 'remove the src of the message if there were no errors!!
End If

End Sub
Private Function MailFrom(objBody As String)
    Dim strTemp(0 To 1) As String
        strTemp(0) = Split(objBody, "From: ")(1)
        strTemp(1) = Split(strTemp(0), ">")(0)
    MailFrom = Split(strTemp(1), "<")(1)
End Function
Public Sub LoadTreeView()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim TGrp As Node
    Dim i As Long
        Set TGrp = TreeView1.Nodes.Add(, , , "Outbox", 1)
            TGrp.Expanded = True
            rs.Open "Select * from SMTPQUE", cn, adOpenKeyset, adLockReadOnly
                For i = 1 To rs.RecordCount
                    If i = 1 Then Set TGrp = TreeView1.Nodes.Add(TGrp, tvwChild, , MsgID(rs!body), 7) Else _
                    Set TGrp = TreeView1.Nodes.Add(TGrp, tvwNext, , MsgID(rs!body), 7)
                        TreeView1.Nodes.Item(TreeView1.Nodes.Count).Tag = "SMTPQUE"
                rs.MoveNext
                Next i
        rs.Close
    Set rs = Nothing
End Sub
Public Sub LoadTreeviewBADQUE()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim TGrp As Node
    Dim i As Long
        Set TGrp = TreeView1.Nodes.Add(, , , "Outbox BAD Queue", 4)
            TGrp.Expanded = True
            rs.Open "Select * from SMTPBADQUE", cn, adOpenKeyset, adLockReadOnly
                For i = 1 To rs.RecordCount
                    If i = 1 Then Set TGrp = TreeView1.Nodes.Add(TGrp, tvwChild, , MsgID(rs!body), 6) Else _
                    Set TGrp = TreeView1.Nodes.Add(TGrp, tvwNext, , MsgID(rs!body), 6)
                        TreeView1.Nodes.Item(TreeView1.Nodes.Count).Tag = "BAD"
                rs.MoveNext
                Next i
        rs.Close
    Set rs = Nothing
End Sub
Private Function MsgID(strDATA As String)
On Error Resume Next
Dim strTemp(0 To 1) As String
    strTemp(0) = Split(strDATA, "Message-ID:")(1)
    strTemp(1) = Split(strTemp(0), vbCrLf)(0)
        MsgID = "Message-ID:" & strTemp(1)
End Function
Public Sub refresh()
TreeView1.Nodes.Clear
ListView1.ListItems.Clear
LoadTreeView
LoadTreeviewBADQUE
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim strQry As String
Dim i As Integer
    SMTPCon = True
    strQry = "select * from SMTPQUE"
    rs.Open strQry, cn, adOpenKeyset, adLockOptimistic
        If Not rs.RecordCount = 0 Then
            i = rs.RecordCount
            Call SendMail(i, "SMTPQUE") 'allocates a free number for a decremeting number which will allocate the socket and smtp object
        End If
    rs.Close
Set rs = Nothing
'If Not i = 0 Then Call refresh
End Sub
Private Sub CancelRSview() 'public recordset just so we can cancel and interact with the record loading process
    vRs.Cancel
    vRSCancel = True
End Sub
Public Sub StartTimer()
    Timer1.Enabled = True
    SMTPCon = True
End Sub
Public Sub StopTimer()
    Timer1.Enabled = False
    SMTPCon = False
End Sub

Private Sub Timer2_Timer()
initret1 = GetTickCount / 1000
If initret1 > initret2 Then SENDQUIT = True: Timer2.Enabled = False

End Sub

Private Sub Timer3_Timer()
initret3 = GetTickCount / 1000
If initret3 > initret4 Then Call ProcessBAD: Timer2.Enabled = True: Timer3.Enabled = False

End Sub

Private Sub TreeView1_Click()
'MsgBox TreeView1.SelectedItem.Tag
End Sub
Public Sub ViewMessage()
   On Error Resume Next
Screen.MousePointer = 11
Command1.Visible = True
        Dim i, j, l As Long
        Dim strQry(0 To 1), strTemp(0 To 1) As String
        ListView1.ListItems.Clear
           If TreeView1.SelectedItem.Tag = "BAD" Then i = 0: strQry(i) = "Select body from SMTPBADQUE" Else _
                 i = 1: strQry(i) = "select body from SMTPQUE"
            vRs.Open strQry(i), cn, adOpenStatic, adLockReadOnly
                For l = 1 To vRs.RecordCount
                    If TreeView1.SelectedItem.Text = Split(vRs!body, vbCrLf)(0) Then
                        For j = 1 To UBound(Split(vRs!body, vbCrLf))
                          DoEvents: DoEvents: DoEvents
                          strTemp(0) = Split(vRs!body, vbCrLf)(j)
                          strTemp(1) = Split(strTemp(0), vbCrLf)(0)
                          ListView1.ListItems.Add , , strTemp(1)
                          If vRSCancel = True Then GoTo cancelQry
                        Next j
                     End If
                    vRs.MoveNext
                Next l
cancelQry:
vRSCancel = False
                 vRs.Close
    Set vRs = Nothing
Screen.MousePointer = 0
Command1.Visible = False

End Sub
Private Sub TreeView1_DblClick()
Call ViewMessage
End Sub

Private Sub TreeView1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then Call refresh
End Sub

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    PopupMenu mnu
End If
End Sub

Private Sub UserControl_Show()
On Error GoTo mailexit
Dim iNdx As Integer
PDNS = "mail.com"
  GetDNSInfo
  iNdx = 0
    While (iNdx <= UBound(sDNS))
   Text1 = sDNS(iNdx)
        iNdx = iNdx + 1
    Wend
Exit Sub
mailexit:
End Sub
