VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "POP3GRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'this object is created for each pop3 retrieval session (per pop3 account)
Public objUser As String
Public objPASS As String
Public Octets As Long
Public Total As Long 'total number of messages
Public EXTPOP3 As String
Public EXTSMTP As String
Public EXTUSER As String
