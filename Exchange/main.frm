VERSION 5.00
Begin VB.Form frmMain1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MX Record lookup"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin VB.ComboBox cmbDNS 
      Height          =   315
      Left            =   2160
      TabIndex        =   4
      Text            =   "cmbDNS"
      Top             =   600
      Width           =   3855
   End
   Begin VB.ListBox lstMX 
      Height          =   2010
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   6975
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go"
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtDomain 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Text            =   "mail.com"
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label lblMX 
      Caption         =   "DNS server to use:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblMX 
      Caption         =   "Domain name for MX query:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmMain1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

' The only registry error that I care about =)
Const ERROR_SUCCESS = 0&

' Registry access functions
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

' Variant (string array) that holds all the DNS servers found in the registry
Dim sDNS As Variant

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''  StripTerminator
'''''''''''
''' Remove the NULL character from the end of a string
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''  GetDNSInfo
'''''''''''
''' Read the registry to find all the DNS servers (DHCP and configured)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub GetDNSInfo()
    Dim hKey As Long
    Dim hError As Long
    Dim sdhcpBuffer As String
    Dim sBuffer As String
    Dim sFinalBuff As String
    
    sdhcpBuffer = Space(1000)
    sBuffer = Space(1000)
    If (RegOpenKeyEx(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Tcpip\Parameters", 0, KEY_READ, hKey) = ERROR_SUCCESS) Then
        ' DNS servers configured through DHCP
        RegQueryValueEx hKey, "DhcpNameServer", 0, REG_SZ, sdhcpBuffer, 1000
        ' DNS servers configured through Network control panel applet
        RegQueryValueEx hKey, "NameServer", 0, REG_SZ, sBuffer, 1000
        RegCloseKey hKey
            
        sFinalBuff = Trim(StripTerminator(sBuffer) & " " & StripTerminator(sdhcpBuffer))
        sDNS = Split(sFinalBuff, " ")
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''  ParseName
'''''''''''
''' Parse the server name out of the MX record, returns it in variable sName, iNdx is also
''' modified to point to the end of the parsed structure.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub ParseName(dnsReply() As Byte, iNdx As Integer, sName As String)
    Dim iCompress As Integer        ' Compression index (index into original buffer)
    Dim iChCount As Integer         ' Character count (number of chars to read from buffer)
        
    ' While we didn't encounter a null char (end-of-string specifier)
    While (dnsReply(iNdx) <> 0)
        ' Read the next character in the stream (length specifier)
        iChCount = dnsReply(iNdx)
        ' If our length specifier is 192 (0xc0) we have a compressed string
        If (iChCount = 192) Then
            ' Read the location of the rest of the string (offset into buffer)
            iCompress = dnsReply(iNdx + 1)
            ' Call ourself again, this time with the offset of the compressed string
            ParseName dnsReply(), iCompress, sName
            ' Step over the compression indicator and compression index
            iNdx = iNdx + 2
            ' After a compressed string, we are done
            Exit Sub
        End If
        
        ' Move to next char
        iNdx = iNdx + 1
        ' While we should still be reading chars
        While (iChCount)
            ' add the char to our string
            sName = sName + Chr(dnsReply(iNdx))
            iChCount = iChCount - 1
            iNdx = iNdx + 1
        Wend
        ' If the next char isn't null then the string continues, so add the dot
        If (dnsReply(iNdx) <> 0) Then sName = sName + "."
    Wend
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''  GetMXName
'''''''''''
''' Parses the buffer returned by the DNS server, returns the best MX server (lowest preference
''' number), iNdx is modified to point to current buffer position (should be the end of buffer
''' by the end, unless a record other than MX is found)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function GetMXName(dnsReply() As Byte, iNdx As Integer, iAnCount As Integer) As String
    Dim iChCount As Integer     ' Character counter
    Dim sTemp As String         ' Holds original query string
    
    Dim iBestPref As Integer    ' Holds the "best" preference number (lowest)
    Dim sBestMX As String       ' Holds the "best" MX record (the one with the lowest preference)
    
    iBestPref = -1
    
    ParseName dnsReply(), iNdx, sTemp
    ' Step over null
    iNdx = iNdx + 2
    
    ' Step over 6 bytes (not sure what the 6 bytes are, but all other
    '   documentation shows steping over these 6 bytes)
    iNdx = iNdx + 6
    
    While (iAnCount)
        ' Check to make sure we received an MX record
        If (dnsReply(iNdx) = 15) Then
            Dim sName As String
            Dim iPref As Integer
            
            sName = ""
            
            ' Step over the last half of the integer that specifies the record type (1 byte)
            ' Step over the RR Type, RR Class, TTL (3 integers - 6 bytes)
            iNdx = iNdx + 1 + 6
            
            ' Read the MX data length specifier
            '              (not needed, hence why it's commented out)
            'MemCopy iMXLen, dnsReply(iNdx), 2
            'iMXLen = ntohs(iMXLen)
            
            ' Step over the MX data length specifier (1 integer - 2 bytes)
            iNdx = iNdx + 2
            
            MemCopy iPref, dnsReply(iNdx), 2
            iPref = ntohs(iPref)
            ' Step over the MX preference value (1 integer - 2 bytes)
            iNdx = iNdx + 2
            
            ' Have to step through the byte-stream, looking for 0xc0 or 192 (compression char)
            ParseName dnsReply(), iNdx, sName
            lstMX.AddItem "[Preference = " & iPref & "] " & sName
            
            If (iBestPref = -1 Or iPref < iBestPref) Then
                iBestPref = iPref
                sBestMX = sName
            End If
            
            ' Step over 3 useless bytes
            iNdx = iNdx + 3
        Else
            GetMXName = sBestMX
            Exit Function
        End If
        iAnCount = iAnCount - 1
    Wend
    
    GetMXName = sBestMX
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''  MakeQName
'''''''''''
''' Takes sDomain and converts it to the QNAME-type string, returns that. QNAME is how a
''' DNS server expects the string.
'''
'''    Ex...    Pass -        mail.com
'''             Returns -     &H4mail&H3com
'''                            ^      ^
'''                            |______|____ These two are character counters, they count the
'''                                         number of characters appearing after them
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function MakeQName(sDomain As String) As String
    Dim iQCount As Integer      ' Character count (between dots)
    Dim iNdx As Integer         ' Index into sDomain string
    Dim iCount As Integer       ' Total chars in sDomain string
    Dim sQName As String        ' QNAME string
    Dim sDotName As String      ' Temp string for chars between dots
    Dim sChar As String         ' Single char from sDomain string
    
    iNdx = 1
    iQCount = 0
    iCount = Len(sDomain)
    ' While we haven't hit end-of-string
    While (iNdx <= iCount)
        ' Read a single char from our domain
        sChar = Mid(sDomain, iNdx, 1)
        ' If the char is a dot, then put our character count and the part of the string
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''  MakeQName
'''''''''''
''' Performs the actual IP work to contact the DNS server, calls the other functions to parse
''' and return the best server to send email through
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function MX_Query() As String
    Dim StartupData As WSADataType
    Dim SocketBuffer As sockaddr
    Dim IpAddr As Long
    Dim iRC As Integer
    Dim dnsHead As DNS_HEADER
    Dim iSock As Integer

    ' Initialize the Winsocket
    iRC = WSAStartup(&H101, StartupData)
    iRC = WSAStartup(&H101, StartupData)
    If iRC = SOCKET_ERROR Then Exit Function
    
    ' Create a socket
    iSock = socket(AF_INET, SOCK_DGRAM, 0)
    If iSock = SOCKET_ERROR Then Exit Function
    
    IpAddr = GetHostByNameAlias(cmbDNS.Text)
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
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    '' Variables
    ''''''''''''''''''''''''''''''''''''''''''''''
    Dim dnsQuery() As Byte
    Dim sQName As String
    Dim dnsQueryNdx As Integer
    Dim iTemp As Integer
    Dim iNdx As Integer
    dnsQueryNdx = 0
    
    ReDim dnsQuery(4000)
    
    ' Setup the dns structure to send the query in
    
    
    ' First goes the DNS header information
    MemCopy dnsQuery(dnsQueryNdx), dnsHead, 12
    dnsQueryNdx = dnsQueryNdx + 12
    
    ' Then the domain name (as a QNAME)
    sQName = MakeQName(txtDomain)
    iNdx = 0
    While (iNdx < Len(sQName))
        dnsQuery(dnsQueryNdx + iNdx) = Asc(Mid(sQName, iNdx + 1, 1))
        iNdx = iNdx + 1
    Wend

    dnsQueryNdx = dnsQueryNdx + Len(sQName)
    
    ' Null terminate the string
    dnsQuery(dnsQueryNdx) = &H0
    dnsQueryNdx = dnsQueryNdx + 1
    
    ' The type of query (15 means MX query)
    iTemp = htons(15)
    MemCopy dnsQuery(dnsQueryNdx), iTemp, Len(iTemp)
    dnsQueryNdx = dnsQueryNdx + Len(iTemp)
    
    ' The class of query (1 means INET)
    iTemp = htons(1)
    MemCopy dnsQuery(dnsQueryNdx), iTemp, Len(iTemp)
    dnsQueryNdx = dnsQueryNdx + Len(iTemp)
    
    ReDim Preserve dnsQuery(dnsQueryNdx - 1)
    ' Send the query to the DNS server
    iRC = sendto(iSock, dnsQuery(0), dnsQueryNdx + 1, 0, SocketBuffer, Len(SocketBuffer))
    If (iRC = SOCKET_ERROR) Then
        MsgBox "Problem sending"
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    '' Variables
    ''''''''''''''''''''''''''''''''''''''''''''''
    Dim dnsReply(2048) As Byte
    ' Wait for answer from the DNS server
    iRC = recvfrom(iSock, dnsReply(0), 2048, 0, SocketBuffer, Len(SocketBuffer))
    If (iRC = SOCKET_ERROR) Then
        MsgBox "Problem receiving"
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    '' Variables
    ''''''''''''''''''''''''''''''''''''''''''''''
    Dim iAnCount As Integer
    ' Get the number of answers
    MemCopy iAnCount, dnsReply(6), 2
    iAnCount = ntohs(iAnCount)
    ' Parse the answer buffer
    MX_Query = GetMXName(dnsReply(), 12, iAnCount)
End Function

Private Sub cmdClear_Click()
    lstMX.Clear
End Sub

Private Sub cmdGo_Click()
    Dim sMX As String
    
    If (cmbDNS.Text <> "") Then
        lstMX.AddItem "Mail routing information for " & txtDomain
        lstMX.AddItem "     using DNS server of " & cmbDNS.Text
        sMX = MX_Query
        If (Len(sMX) > 0) Then
            lstMX.AddItem "Best MX record to send through: " & sMX
        Else
            lstMX.AddItem "No mail routing information found"
        End If
    Else
        MsgBox "ERROR: DNS information not entered/selected"
    End If
End Sub

Private Sub Form_Load()
    Dim iNdx As Integer

    GetDNSInfo
    iNdx = 0
    While (iNdx <= UBound(sDNS))
        cmbDNS.AddItem sDNS(iNdx)
        iNdx = iNdx + 1
    Wend
    
    If (cmbDNS.ListCount > 0) Then cmbDNS.ListIndex = 0
End Sub
