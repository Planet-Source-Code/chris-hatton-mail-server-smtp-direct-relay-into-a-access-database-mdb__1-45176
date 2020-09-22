VERSION 5.00
Begin VB.Form FrmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Global Options"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   960
      TabIndex        =   13
      Top             =   1920
      Width           =   5775
      Begin VB.OptionButton Option2 
         Caption         =   "STMP Relay  (Recommend)"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   15
         Top             =   240
         Width           =   2775
      End
      Begin VB.OptionButton Option2 
         Caption         =   "STMP Direct Connection   (if you have a high bandwith connection and unlimited megabytes, beware this is generally slower!)"
         Height          =   675
         Index           =   1
         Left            =   480
         TabIndex        =   14
         Top             =   480
         Width           =   5175
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   1560
      TabIndex        =   12
      Top             =   375
      Width           =   5055
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      Picture         =   "FrmOptions.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   11
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3960
      TabIndex        =   10
      Top             =   795
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   1560
      TabIndex        =   9
      Top             =   1695
      Width           =   5055
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      Picture         =   "FrmOptions.frx":030A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   8
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Top             =   1170
      Width           =   375
   End
   Begin VB.Frame Frame3 
      Height          =   120
      Left            =   2160
      TabIndex        =   6
      Top             =   3615
      Width           =   4455
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   960
      TabIndex        =   3
      Top             =   3840
      Width           =   5775
      Begin VB.OptionButton Option1 
         Caption         =   "128K ISDN or 100k DSL"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   23
         Top             =   600
         Width           =   4935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "56k Modem Connection"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   5
         Top             =   240
         Width           =   4935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "DSL or T1 - T3 Connections or above"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   4
         Top             =   960
         Width           =   4935
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   5880
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      Picture         =   "FrmOptions.frx":0614
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Check for new messages every"
      Height          =   255
      Left            =   1440
      TabIndex        =   22
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "minute(s)"
      Height          =   255
      Left            =   4440
      TabIndex        =   21
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "POP3 Connector"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "SMTP Connector"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Connection attempts before trying the next POP3 Account"
      Height          =   255
      Left            =   2280
      TabIndex        =   18
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label Label6 
      Caption         =   "Try"
      Height          =   255
      Left            =   1440
      TabIndex        =   17
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "External Connection Type"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   3600
      Width           =   2295
   End
End
Attribute VB_Name = "FrmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Public Sub ReadOptions()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
    rs.Open "select * from settings where configurationset = 1", cn, adOpenKeyset, adLockOptimistic
        Text1 = "" & rs!pop3timer
        Text2 = "" & rs!ConnectionRetrys
        If rs!smtprelay = True Then Option2(0) = True
        If rs!smtprelay = False Then Option2(1) = True
        If rs!buffer = 2048 Then Option1(0) = True
        If rs!buffer = 4096 Then Option1(1) = True
        If rs!buffer = 8192 Then Option1(2) = True
    rs.Close
Set rs = Nothing

End Sub

Private Sub SaveOptions()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
    rs.Open "select * from settings where configurationset = 1", cn, adOpenKeyset, adLockOptimistic
        If Len(Text1) Then rs!pop3timer = "" & Text1
        If Len(Text2) Then rs!ConnectionRetrys = "" & Text2
        If Option2(0) = True Then rs!smtprelay = True
        If Option2(1) = True Then rs!smtprelay = False
        If Option1(0) = True Then rs!buffer = "" & "2048"
        If Option1(1) = True Then rs!buffer = "" & "4096"
        If Option1(2) = True Then rs!buffer = "" & "8192"
    rs.Update
    rs.Close
Set rs = Nothing
Call FrmMain.UpdateGlobalChanges
Unload Me
End Sub
Public Sub Save(Relay As Boolean)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
    rs.Open "select * from settings where configurationset = 1", cn, adOpenKeyset, adLockOptimistic
        rs!smtprelay = Relay
    rs.Update
    rs.Close
Set rs = Nothing
Call FrmMain.UpdateGlobalChanges
Call ADO_Close
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
    Call SaveOptions
End Sub

Private Sub Form_Load()
On Error Resume Next
Call ADO_Connect
Call ReadOptions
End Sub
Public Sub ADO_Connect()
Set cn = New ADODB.Connection
        cn.Provider = "Microsoft.Jet.OLEDB.4.0"
        cn.Open FrmMain.DBPATH
        cn.CursorLocation = adUseClient

End Sub
Public Sub ADO_Close()
On Error Resume Next
    cn.Close
    Set cn = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    cn.Close
    Set cn = Nothing
End Sub

