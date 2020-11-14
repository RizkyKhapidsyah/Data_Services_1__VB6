Attribute VB_Name = "MHelper"
Option Explicit

' ------------------------------------------------------------
'  Copyright ©2001 Mike G --> IvbNET.COM
'  All Rights Reserved, http://www.ivbnet.com
'  EMAIL : webmaster@ivbnet.com
' ------------------------------------------------------------
'  You are free to use this code within your own applications,
'  but you are forbidden from selling or distributing this
'  source code without prior written consent.
' ------------------------------------------------------------
'You need ref to MTS as ADO


Public Function RaiseError(Module As String, FunctionName As String)
    Dim lErr As Long
    Dim sErr As String
    'Set the default to disable transaction.
    'Now unless someone does a SetComplete the transaction will abort.
    'This is just like calling SetAbort,
    'but has it doesn't destroy the Err object if we are in a transaction.
    GetObjectContext.DisableCommit

    lErr = VBA.Err.Number
    sErr = VBA.Err.Description
    Err.Raise lErr, sErr
End Function

' mp is short for MakeParameter - does typesafe array creation for use with Run* functions
Public Function mp(ByVal PName As String, ByVal PType As ADODB.DataTypeEnum, ByVal PSize As Integer, ByVal PValue As Variant)
    mp = Array(PName, PType, PSize, PValue)
End Function

'Converts a variant into a string. If the varaint is null, it is converted to the empty string.
Public Function ConvertToString(v As Variant) As String
    If IsNull(v) Then
        ConvertToString = ""
    Else
        ConvertToString = CStr(v)
    End If
End Function

'Converts any Null variant to null, otherwise it returns the existing value.
Public Function NullsToZero(v As Variant) As Variant
    If IsNull(v) Then
        NullsToZero = 0
    Else
        NullsToZero = v
    End If
End Function
