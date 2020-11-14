VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 4  'RequiresNewTransaction
END
Attribute VB_Name = "clsData_Objects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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


Private Const m_modName    As String = "Data_Services.clsData_Objects"

'Run SQL and return rs back
Public Function RunSQLReturnRS(ByVal sConnStr As String, _
                               ByVal sSQL As String) As ADODB.Recordset

    On Error GoTo errorHandler


    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    'Run sql
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockBatchOptimistic
    rs.ActiveConnection = sConnStr
    'Execute
    rs.Open sSQL

    ' Disconnect the recordsets and cleanup
    Set rs.ActiveConnection = Nothing
    Set RunSQLReturnRS = rs
    Set rs = Nothing
    GetObjectContext.SetComplete

    Exit Function

errorHandler:
    Set rs = Nothing
    RaiseError m_modName, "RunSQLReturnRS"
End Function


'Run SQL and return rs true or false
Public Function RunSQLReturnBool(ByVal sConnStr As String, _
                                 ByVal sSQL As String) As Boolean

    On Error GoTo errorHandler

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset


    'Run sql
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockBatchOptimistic
    rs.ActiveConnection = sConnStr
    'Execute
    rs.Open sSQL, , , , adExecuteNoRecords

    ' Disconnect the recordsets and cleanup
    Set rs.ActiveConnection = Nothing
    RunSQLReturnBool = True
    Set rs = Nothing
    GetObjectContext.SetComplete

    Exit Function

errorHandler:
    RunSQLReturnBool = False
    Set rs = Nothing
    RaiseError m_modName, "RunSQLReturnBool"
End Function


Function RunSPRetBool(ByVal sConn As String, ByVal strSP As String, _
                            ParamArray params() As Variant) As Boolean
    On Error GoTo errorHandler

    ' Create the ADO objects
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command

    ' Init the ADO objects & the stored proc parameters
    cmd.ActiveConnection = sConn
    cmd.CommandText = strSP
    cmd.CommandType = adCmdStoredProc
    CollectParams cmd, params

    ' Execute the query without returning a recordset
    cmd.Execute , , adExecuteNoRecords

    ' Cleanup
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    RunSPRetBool = True
    GetObjectContext.SetComplete
    Exit Function

errorHandler:
    RunSPRetBool = False
    Set cmd = Nothing
    RaiseError m_modName, "RunSPRetBool(" & strSP & ", ...)"
End Function


Function RunSPRetRS(ByVal sConn As String, ByVal strSP As String, _
                      ParamArray params() As Variant) As ADODB.Recordset
    On Error GoTo errorHandler

    ' Create the ADO objects
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command

    ' Init the ADO objects & the stored proc parameters
    cmd.ActiveConnection = sConn
    cmd.CommandText = strSP
    cmd.CommandType = adCmdStoredProc
    CollectParams cmd, params

    'Return rs
    Set RunSPRetRS = cmd.Execute()

    ' Cleanup
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    GetObjectContext.SetComplete
    Exit Function

errorHandler:
    Set cmd = Nothing
    RaiseError m_modName, "RunSPRetBool(" & strSP & ", ...)"
End Function

'Return string
Function RunSPRetString_Params(ByVal sConn As String, ByVal strSP As String, _
                               ParamArray params() As Variant) As String

    On Error GoTo errorHandler
    ' Create the ADO objects
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command

    ' Init the ADO objects & the stored proc parameters
    cmd.ActiveConnection = sConn
    cmd.CommandText = strSP
    cmd.CommandType = adCmdStoredProc
    CollectParams cmd, params

    'Return rs
    Set rs = cmd.Execute()
    RunSPRetString_Params = rs.GetString(, , , vbTab)

    ' Cleanup
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set rs = Nothing
    Exit Function

errorHandler:
    Set rs = Nothing
    Set cmd = Nothing
    RaiseError m_modName, "RunSPRetString_Params"
End Function

'Return string
Function RunSPRetString_NoParams(ByVal sConn As String, _
                               ByVal strSP As String) As String

    On Error GoTo errorHandler
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    rs.CursorLocation = adUseClient
    rs.CacheSize = 1
    'Run sp
    rs.Open strSP, sConn, _
            adOpenForwardOnly, adLockReadOnly
    'Ret string
    RunSPRetString_NoParams = rs.GetString(, , , vbTab)
    'Clean up
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
    GetObjectContext.SetComplete
    Exit Function

errorHandler:
    Set rs = Nothing
    RaiseError m_modName, "RunSPRetString_NoParams"
End Function

'Collect params
Private Function CollectParams(ByRef cmd As ADODB.Command, _
                               ParamArray argparams() As Variant)
    Dim params As Variant, v As Variant
    Dim i As Integer, l As Integer, u As Integer

    params = argparams(0)
    For i = LBound(params) To UBound(params)
        l = LBound(params(i))
        u = UBound(params(i))
        ' Check for nulls.
        If u - l = 3 Then
            If VarType(params(i)(3)) = vbString Then
                v = IIf(params(i)(3) = "", Null, params(i)(3))
            Else
                v = params(i)(3)
            End If
            cmd.Parameters.Append cmd.CreateParameter(params(i)(0), _
                    params(i)(1), adParamInput, params(i)(2), v)
        Else
            RaiseError m_modName, "CollectParams(...): incorrect # of parameters"
        End If
    Next i
End Function

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
