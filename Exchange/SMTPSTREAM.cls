VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SMTPSTREAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim cn As ADODB.Connection
Public StrRCPTTO, StrFROMNAME, StrSubject As String
Public ColRCPT As New Collection
Public TempPathFile As String
Public objFlatFile As String
Public DBPATH As String

Public Sub StreamTEXT(StrStream As String, QueLocal As Boolean)
    Dim rs As ADODB.Recordset
    Dim strQry As String
    Set cn = New ADODB.Connection
        On Error Resume Next
        cn.Provider = "Microsoft.Jet.OLEDB.4.0"
        cn.Open DBPATH
        cn.CursorLocation = adUseClient
    If QueLocal = True Then strQry = "select * from Firststore" Else strQry = "select * from SMTPQUE"
    Set rs = New ADODB.Recordset
        rs.Open strQry, cn, adOpenKeyset, adLockOptimistic
            rs.AddNew
            rs![Date] = Date
            rs!RCPT = StrRCPTTO
            If QueLocal = True Then
                rs!body = objFlatFile 'load from temp file
            Else
                rs!body = StrStream
            End If
            rs!octat = Len(rs!body) * 6 / 8 / 255 / 3 'convert octets to byte in a very basic mathmatical form.

        rs.Update
        rs.Close
        Set rs = Nothing
        
        Call RemoveTempFile 'clean up
        
        
End Sub
Private Sub RemoveTempFile()
Kill TempPathFile
End Sub
Public Sub ImportFlatFile()
                Open TempPathFile For Input As #3
                      objFlatFile = Input$(LOF(3), #3)
                Close #3
End Sub


