VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBienesNavidad"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0
Option Explicit
Dim dbConec As DConecta
Dim psSql As String
Dim prs   As ADODB.Recordset
Dim sConexion As String

Public Function CargaBeneficiarios(psTipo As String, pdFecha As Date, Optional psPersNombre As String = "", Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
   On Error GoTo CargaBeneficiariosErr
      psSql = " SELECT cPersNombre, nCanasta, nPavo, nJuguete, nIgv, nCanasta + nPavo + nJuguete + nIgv as nTotal " _
            & " FROM BienesNavidad WHERE cTipo = '" & psTipo & "' and datediff(d,dFecha,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 " _
            & IIf(psPersNombre = "", "", "and cPersNombre LIKE '" & psPersNombre & "%'")
            
      Set CargaBeneficiarios = dbConec.CargaRecordSet(psSql, LockType)
   Exit Function
CargaBeneficiariosErr:
   Call RaiseError(MyUnhandledError, "DMarcas:CargaBeneficiarios Method")
End Function

Public Sub InsertaBeneficiario(psTipo As String, pdFecha As Date, psPersNombre As String, pnCanasta As Currency, pnPavo As Currency, pnJuguete As Currency, pnIGV As Currency)
   On Error GoTo InsertaMarcaErr
      psSql = "INSERT INTO BienesNavidad (cTipo, dFecha, cPersNombre, nCanasta, nPavo, nJuguete, nIgv) " _
            & "VALUES ('" & psTipo & "','" & Format(pdFecha, gsFormatoFecha) & "','" & psPersNombre & "'," & pnCanasta & "," & pnPavo & "," & pnJuguete & "," & pnIGV & ")"
      dbConec.Ejecutar psSql
   Exit Sub
InsertaMarcaErr:
   Call RaiseError(MyUnhandledError, "DMarcas:InsertaMarca Method")
End Sub

Public Function VerificaBeneficiario(psTipo As String, pdFecha As Date, psPersNombre As String) As Boolean
   On Error GoTo VerificaBeneficiarioErr
   Dim rs As New ADODB.Recordset
      VerificaBeneficiario = False
      psSql = "   SELECT cPersNombre FROM  BienesNavidad " _
            & "   Where cTipo = '" & psTipo & "' and datediff(d,dFecha,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and cPersNombre = '" & psPersNombre & "'  "
      Set rs = dbConec.CargaRecordSet(psSql)
      If Not rs.EOF And Not rs.BOF Then
         VerificaBeneficiario = True
      End If
      RSClose rs
   Exit Function
VerificaBeneficiarioErr:
   Call RaiseError(MyUnhandledError, "DMarcas:VerificaBeneficiario Method")
End Function

Public Sub ActualizaBeneficiario(psTipo As String, pdFecha As Date, psPersNombre As String, pnCanasta As Currency, pnPavo As Currency, pnJuguete As Currency, pnIGV As Currency)
   On Error GoTo ActualizaBeneficiarioErr
      psSql = "UPDATE BienesNavidad SET cPersNombre = '" & psPersNombre & "', nCanasta = " & pnCanasta & ", nPavo = " & pnPavo & ", nJuguete = " & pnJuguete & ", nIgv = " & pnIGV & " " _
        & "   WHERE cTipo = '" & psTipo & "' and datediff(d,dFecha,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and cPersNombre = '" & psPersNombre & "' "
      dbConec.Ejecutar psSql
   Exit Sub
ActualizaBeneficiarioErr:
   Call RaiseError(MyUnhandledError, "DMarcas:ActualizaBeneficiario Method")
End Sub

Public Sub EliminaBeneficiario(psTipo As String, pdFecha As Date, psPersNombre As String)
   On Error GoTo EliminaBeneficiarioErr
      psSql = "DELETE FROM BienesNavidad WHERE cTipo = '" & psTipo & "' and datediff(d,dFecha,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and cPersNombre = '" & psPersNombre & "' "
      dbConec.Ejecutar psSql
   Exit Sub
EliminaBeneficiarioErr:
   Call RaiseError(MyUnhandledError, "DMarcas:EliminaBeneficiario Method")
End Sub

Private Sub Class_Initialize()
   Set dbConec = New DConecta
   dbConec.AbreConexion
End Sub

Private Sub Class_Terminate()
dbConec.CierraConexion
Set dbConec = Nothing
End Sub


