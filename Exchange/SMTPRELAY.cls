VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SMTPRELAY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'this module provides the esentials of the external or direct smtp relay model.
Public ColRCPT As New Collection
Private strSMTP As String
Private strBody As String
Private strMail As String
Public Error As Boolean 'mark if there has been an error with sending this smtp message.

Public Property Get objSMTP() As Variant
    objSMTP = strSMTP
End Property

Public Property Let objSMTP(ByVal vNewValue As Variant)
    strSMTP = vNewValue
End Property

Public Property Get objBody() As Variant
    objBody = strBody
End Property

Public Property Let objBody(ByVal vNewValue As Variant)
    strBody = vNewValue
End Property

Public Property Get objMailFrom() As Variant
    objMailFrom = strMail
End Property

Public Property Let objMailFrom(ByVal vNewValue As Variant)
    strMail = vNewValue
End Property


