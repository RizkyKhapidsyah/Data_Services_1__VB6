VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "clsBus"
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
'You need ref to MTS as ADO and Data_Services

Const ODBCName = "DSN=Pubs"
Const m_name = "clsBus"

Implements ObjectControl


Public Function Test_SQL_With_Rs() As ADODB.Recordset
    On Error GoTo FixError
    Dim obj As Data_Services.clsData_Objects
    Set obj = _
            GetObjectContext.CreateInstance("Data_Services.clsData_Objects")
    Set Test_SQL_With_Rs = _
            obj.RunSQLReturnRS(ODBCName, "SELECT * FROM Authors")

    GetObjectContext.SetComplete
    Set obj = Nothing
    Exit Function
FixError:
    RaiseError m_name, "Test_SQL_With_RS"
End Function

Public Function Test_SQL_With_Out_Rs() As Variant
    On Error GoTo FixError
    Dim obj As Data_Services.clsData_Objects
    Set obj = _
            GetObjectContext.CreateInstance("Data_Services.clsData_Objects")

    Test_SQL_With_Out_Rs = _
            obj.RunSQLReturnBool(ODBCName, "SELECT * FROM Authors")

    GetObjectContext.SetComplete
    Set obj = Nothing
    Exit Function
FixError:
    RaiseError m_name, "Test_SQL_With_Out_Rs"
End Function



Public Function Exec_Sp_With_No_Ret() As Variant
    On Error GoTo FixError
    Dim obj As Data_Services.clsData_Objects
    Set obj = _
            GetObjectContext.CreateInstance("Data_Services.clsData_Objects")

    Exec_Sp_With_No_Ret = _
            obj.RunSPRetBool(ODBCName, _
            "sp_Select_By_State", mp("@State", adVarChar, 2, "CA"))
    Set obj = Nothing
    GetObjectContext.SetComplete
    Exit Function
FixError:
    Exec_Sp_With_No_Ret = False
    RaiseError m_name, "Exec_Sp_With_No_Ret"
End Function

Public Function Exec_Sp_With_RS() As ADODB.Recordset
    On Error GoTo FixError
    Dim obj As Data_Services.clsData_Objects
    Set obj = _
            GetObjectContext.CreateInstance("Data_Services.clsData_Objects")

    Set Exec_Sp_With_RS = _
            obj.RunSPRetRS(ODBCName, _
            "sp_Select_By_State", mp("@State", adVarChar, 2, "CA"))
    Set obj = Nothing
    GetObjectContext.SetComplete
    Exit Function
FixError:
    RaiseError m_name, "Exec_Sp_With_RS"
End Function


Public Function Exec_Sp_Test() As ADODB.Recordset
    On Error GoTo FixError
    Dim obj As Data_Services.clsData_Objects
    Set obj = _
            GetObjectContext.CreateInstance("Data_Services.clsData_Objects")

    Set Exec_Sp_Test = _
            obj.RunSPRetRS(ODBCName, _
            "sp_test", mp("@test", adInteger, 0, 1), _
                       mp("@State", adVarChar, 2, "CA"))
    Set obj = Nothing
    GetObjectContext.SetComplete
    Exit Function
FixError:
    RaiseError m_name, "Exec_Sp_Test"
End Function

Private Sub ObjectControl_Activate()

End Sub

Private Function ObjectControl_CanBePooled() As Boolean
   ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
   
End Sub
