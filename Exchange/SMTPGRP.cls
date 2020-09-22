VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SMTPGRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum SmtpStages 'major
    sckhelo = 0
    sckmailfrom = 1
    sckrcpt = 2
    sckdata = 3
    sckdataready = 4
    sckquit = 5
    
End Enum


Private varGetData As String 'body of email
Private StrBuffer As String
Private strInData As String
Public strMail As String 'mail from
Public ColRCPT As New Collection
Public strDATA As String
Public strOut As String
Dim GETDATA As Boolean
Public SMTPID As Long
Public TempDumpFile As String
Public DNS As String 'current DNS
Public RemoteHost As String 'ip address
Public RemoteHostName  As String 'remote host name
Public sckState As SmtpStages 'major stage
Public strRCPT As String 'indiviual recipients

Public Property Get BodyData() As Variant
    
If InStr(1, varGetData, "Message-ID:") = 0 Then
    Randomize (123456789)
    BodyData = "Message-ID: <" & Int((1234568910112# * Rnd)) & "$" & Int((123456891# * Rnd)) & "$" & Int((12345689 * Rnd)) & "@" & DNS & ">" & vbCrLf & varGetData
    Exit Sub
Else
    BodyData = varGetData
End If
                                
If InStr(1, varGetData, "From:") = 0 Then
    BodyData = "From: " & LCase(strMail) & vbCrLf & varGetData
Else
    BodyData = varGetData
End If
End Property

Public Property Let BodyData(ByVal vNewValue As Variant)
    varGetData = varGetData & vNewValue
  '   Call TEMPSMTP(vNewValue)

End Property
Public Sub Parser()
If sckState = sckhelo Then
If Not SMTPID = 0 Then
    TempDumpFile = App.Path & "\" & Year(Now) & Month(Now) & Day(Now) & SMTPID & ".tmp"
End If
Open TempDumpFile For Append As 1#
End If
TopLine:


Select Case sckState
    
            Case sckhelo
                    If Mid(UCase(strInData), 1, 4) = UCase("helo") Then
                            If InStr(1, strInData, vbCrLf) Then strInData = Split(strInData, vbCrLf)(0)
                                strOut = "250 " & RemoteHostName & "." & DNS & " Hello [" & RemoteHost & "]"
                                sckState = sckmailfrom: Exit Sub
                        ElseIf Mid(UCase(strInData), 1, 5) = UCase("esmtp") Then
                            strOut = "ESMTP not supported": Exit Sub
                        Else
                            If InStr(1, UCase(strInData), UCase("Mail From:")) = 1 Then
                                sckState = sckmailfrom
                                GoTo TopLine
                            End If
                    End If
                    
            Case sckmailfrom
                    If InStr(1, UCase(strInData), UCase("Mail From:")) = 1 Then
                        strMail = Split(strInData, ":")(1)
                            If InStr(1, LCase(strMail), "@") = 0 Then  'add dns
                                If InStr(1, strMail, vbCrLf) Then strMail = Split(strMail, vbCrLf)(0)
                                    strMail = strMail & "@" & DNS
                                    If chkFormat(strMail) = True Then
                                        strOut = "250 OK " & LTrim$(LCase(strMail)) '& "....Sender OK"
                                        sckState = sckrcpt
                                        Else
                                        strOut = "501 Invalid Command or Unregconized Parameter: " & LCase(strMail)
                                    End If
                                    Exit Sub
                                Else
                                If InStr(1, strMail, vbCrLf) Then strMail = Split(strMail, vbCrLf)(0)
                                    If chkFormat(strMail) = True Then
                                        strOut = "250 OK " & LTrim$(LCase(strMail)) '& "....Sender OK"
                                        sckState = sckrcpt
                                        Else
                                        strOut = "501 Invalid Command or Unregconized Parameter: " & LCase(strMail)
                                    End If
                                   Exit Sub
                              End If
                    End If
            
            Case sckrcpt
                    If InStr(1, UCase(strInData), UCase("Mail From:")) = 1 Then
                        strOut = "503 Sender already specified": Exit Sub
                        ElseIf InStr(1, UCase(strInData), UCase("RCPT TO:")) = 1 Then
                            If Not InStr(1, strInData, "@") = 0 Then strRCPT = Split(strInData, ": ")(1)
                            If InStr(1, strRCPT, vbCrLf) Then strRCPT = Split(strRCPT, vbCrLf)(0)
                                If Left$(strRCPT, 1) = "@" Then strRCPT = Split(Left$(strInData, Len(strInData) - 2), ": ")(1) + "@" & DNS
                                If Left$(strRCPT, 1) = "" Then strRCPT = Split(Left$(strInData, Len(strInData) - 2), ": ")(1) + "@" & DNS
 
                                If ColRCPT.Count >= 60 Then
                                    strOut = "452 Too many recipients"
                                    Exit Sub
                                End If
                            If InStr(1, strRCPT, vbCrLf) Then strRCPT = Split(strRCPT, vbCrLf)(0)
                                If chkFormat(strRCPT) = True Then
                                        If InStr(1, strRCPT, "@") = 0 Then strRCPT = strRCPT & "@" & DNS 'treat it as local if no domain is specified.
                                        strOut = "250 OK " & LTrim$(LCase(strRCPT))
                                        ColRCPT.Add LTrim$(LCase(strRCPT)): strRCPT = "" 'address is ok wait for another one
                                    Exit Sub
                                Else
                                    strOut = "501 Invaild address"
                                    Exit Sub
                                End If
                                     sckState = sckdata
                            End If
            Case sckdata
                        
                        If strOut = "354 Start Mail input; end with <CRLF> . <CRLF>" Then strOut = ""
                         Print #1, strInData
                            If Right(strInData, 5) = vbCrLf & "." & vbCrLf Then 'check for "" . "" EOF
                                strOut = "250 OK Mail is placed in the que"
                                sckState = sckquit
                                Close 1#
                            End If
                            
                    Exit Sub
    End Select
    
                If InStr(1, UCase(strInData), UCase("HELO")) = 1 Then
                        strOut = "250 " & RemoteHostName & "." & DNS & " Hello [" & RemoteHost & "]"
                    
                    ElseIf InStr(1, UCase(strInData), UCase("RCPT TO:")) = 1 Then
                        If sckState = sckmailfrom Then strOut = "503 Need Mail From: first"
                    ElseIf InStr(1, UCase(strInData), UCase("QUIT")) = 1 Then
                            sckState = sckquit
                            strOut = "Bye Bye": Exit Sub
                    ElseIf InStr(1, UCase(strInData), UCase("RSET")) = 1 Then
                             sckState = sckhelo
                             strOut = "250 Resetting"
                             strMail = "": strDATA = ""
                             Do
                               If Not ColRCPT.Count = 0 Then ColRCPT.Remove (ColRCPT.Count) 'remove recipients in memory
                             Loop Until ColRCPT.Count = 0
                    ElseIf InStr(1, UCase(strInData), UCase("DATA")) = 1 Then
                        If sckState = sckmailfrom Then
                             strOut = "503 Need Mail command"
                    
                        ElseIf sckState = sckrcpt And Not InStr(1, UCase(strInData), UCase("DATA")) = 1 Then
                             strOut = "503 Need Rcpt command"
                        ElseIf sckState = sckrcpt And InStr(1, UCase(strInData), UCase("DATA")) Then
                           sckState = sckdata
                           strOut = "354 Start Mail input; end with <CRLF> . <CRLF>"
                            
                            Exit Sub
                       End If
                    Else
                        strOut = "500 Unrecognized command"
                End If
    
End Sub


Private Function chkFormat(strString As String) As Boolean
   Dim formatCol As New Collection
        With formatCol
            .Add "?": .Add ":": .Add "[": .Add "`": .Add "*": .Add "/": .Add "\": .Add "|": .Add "]"
            .Add "=": .Add "+": .Add "!": .Add "#": .Add "$": .Add "^": .Add "&": .Add "(": .Add ")": .Add "~": .Add "{": .Add "}"
        End With
   
   For i = 1 To formatCol.Count
        If InStr(1, strString, LTrim(RTrim(formatCol(i)))) = 0 Then
            chkFormat = True 'good email parameters
        Debug.Print formatCol(i)
        Else
            chkFormat = False
            Exit Function
        End If
   Next i
Set formatCol = Nothing
End Function
Public Property Let InPutParser(ByVal vNewValue As Variant)
strInData = vNewValue

'MsgBox strInData
End Property
