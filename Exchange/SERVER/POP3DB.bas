Attribute VB_Name = "POP3DB"
Public Function DBFileSize(FileName) As String
    On Error Resume Next
    Dim strMDB As String
    strMDB = FileLen(FileName)

    If strMDB >= "1024" Then
        strMDB = CCur(strMDB / 1024) & "KB"
    Else
        
        If strMDB >= "1048576" Then
            strMDB = CCur(strMDB / (1024 * 1024)) & "KB"
        Else
            strMDB = CCur(strMDB) & "B"
        End If
    End If
    DBFileSize = strMDB
End Function

