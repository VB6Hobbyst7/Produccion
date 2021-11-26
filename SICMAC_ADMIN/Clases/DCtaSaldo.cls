VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCtaSaldo"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A82E2C1032C"
Option Base 0
Option Explicit
Dim dbConec As DConecta
Dim psSql As String
Dim prs   As ADODB.Recordset

Dim sConexion As String
Dim sCentralCom As String

Public Function CargaCtaSaldo(Optional psCtaContCod As String = "", Optional psFecha As String = "", Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
Dim sCond As String
   On Error GoTo CargaCtaSaldoErr
   sCond = ""
   If psCtaContCod <> "" Then
      sCond = " WHERE cCtaContCod = '" & psCtaContCod & "'"
   End If
   If psFecha <> "" Then
      sCond = sCond & IIf(sCond = "", "WHERE", "AND") & " dCtaSaldoFecha = (SELECT MAX(dCtaSaldoFecha) FROM CtaSaldo cs1 WHERE cs1.cCtaContCod = cs.cCtaContCod and dCtaSaldoFecha <= '" & psFecha & "')"
   End If
   psSql = "SELECT cCtaContCod, dCtaSaldoFecha, nCtaSaldoImporte FROM CtaSaldo cs " & sCond
   Set CargaCtaSaldo = dbConec.CargaRecordSet(psSql, LockType)
   Exit Function
CargaCtaSaldoErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:CargaCtaSaldo Method")
End Function

Public Sub InsertaCtaSaldo(psCtaContCod As String, psFecha As String, pnImporte As Currency, pnImporteME As Currency)
   On Error GoTo InsertaCtaSaldoErr
      psSql = "INSERT CtaSaldo (cCtaContCod, dCtaSaldoFecha, nCtaSaldoImporte, nCtaSaldoImporteME) " _
            & "VALUES ('" & psCtaContCod & "', '" & psFecha & "'," & pnImporte & "," & pnImporteME & ")"
      dbConec.Ejecutar psSql
   Exit Sub
InsertaCtaSaldoErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:InsertaCtaSaldo Method")
End Sub

Public Sub ActualizaCtaSaldo(psCtaContCod As String, psFecha As String, pnImporte As Currency, pnImporteME As Currency)
   On Error GoTo ActualizaCtaSaldoErr
      psSql = "UPDATE CtaSaldo SET nCtaSaldoImporte = " & pnImporte & ", nCtaSaldoImporteME = " & pnImporteME _
            & "WHERE cCtaContCod = '" & psCtaContCod & "' and dCtaSaldoFecha = '" & psFecha & "'"
      dbConec.Ejecutar psSql
   Exit Sub
ActualizaCtaSaldoErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:ActualizaCtaSaldo Method")
End Sub
Public Sub EliminaCtaSaldo(Optional psCtaContCod As String = "", Optional psFecha As String = "", Optional psSimbolo As String = "=")
Dim sCondCta   As String
Dim sCondFecha As String
   On Error GoTo EliminaCtaSaldoErr
      If psCtaContCod <> "" Then
         sCondCta = "WHERE cCtaContCod = '" & psCtaContCod & "' "
      End If
      If psFecha <> "" Then
         sCondFecha = IIf(sCondCta = "", "WHERE", "AND") & " dCtaSaldoFecha " & psSimbolo & " '" & psFecha & "'"
      End If
      psSql = "DELETE CtaSaldo " & sCondCta & sCondFecha
      dbConec.Ejecutar psSql
   Exit Sub
EliminaCtaSaldoErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:EliminaCtaSaldo Method")
End Sub

Public Function CargaCtaObjSaldo(Optional psCtaContCod As String = "", Optional psFecha As String = "", Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
Dim sCond As String
   On Error GoTo CargaCtaObjSaldoErr
   sCond = ""
   If psCtaContCod <> "" Then
      sCond = " WHERE cCtaContCod = '" & psCtaContCod & "'"
   End If
   If psFecha <> "" Then
      sCond = sCond & IIf(sCond = "", "WHERE", "AND") & " dCtaSaldoFecha = (SELECT MAX(dCtaObjSaldoFecha) FROM CtaObjSaldo cs1 WHERE cs1.cCtaContCod = cs.cCtaContCod and cs1.cObjetoCod = cs.cObjetoCod and dCtaObjSaldoFecha <= '" & psFecha & "')"
   End If
   psSql = "SELECT cCtaContCod, cCtaObjOrden, cObjetoCod, dCtaObjSaldoFecha, nCtaObjSaldoImporte, nCtaObjSaldoImporteME FROM CtaObjSaldo cs " & sCond
   Set CargaCtaObjSaldo = dbConec.CargaRecordSet(psSql, LockType)
   Exit Function
CargaCtaObjSaldoErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:CargaCtaObjSaldo Method")
End Function

Public Sub InsertaCtaSaldoEstad(psCtaSaldoFecha, psCodUser As String, psCtaSaldoFecIni As String, psCtaSaldoFecFin As String, pnValido As Integer)
   On Error GoTo InsertaCtaSaldoEstadErr
   psSql = "INSERT CtaSaldoEstad (dCtaSaldoFecha, cCodUser, dCtaSaldoFechaIni, dCtaSaldoFechaFin, bCtaSaldoValido) " _
         & "VALUES ('" & psCtaSaldoFecha & "','" & psCodUser & "', '" & psCtaSaldoFecIni & "','" & psCtaSaldoFecFin & "', '" & pnValido & "')"
   dbConec.Ejecutar psSql
   Exit Sub
InsertaCtaSaldoEstadErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:InsertaCtaSaldoEstad Method")
End Sub

Public Sub ActualizaCtaSaldoEstad(psCtaSaldoFecha, psCodUser As String, pnValido As Integer)
   On Error GoTo ActualizaCtaSaldoEstadErr
   psSql = "UPDATE CtaSaldoEstad set bCtaSaldoValido = " & pnValido & " WHERE dCtaSaldoFecha = '" & psCtaSaldoFecha & "' and cCodUser = '" & psCodUser & "'"
   dbConec.Ejecutar psSql
   Exit Sub
ActualizaCtaSaldoEstadErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:ActualizaCtaSaldoEstad Method")
End Sub

Public Sub InsertaCtaObjSaldo()
   On Error GoTo InsertaCtaObjSaldoErr
   
   Exit Sub
InsertaCtaObjSaldoErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:InsertaCtaObjSaldo Method")
End Sub

Public Sub ActualizaCtaObjSaldo()
   On Error GoTo ActualizaCtaObjSaldoErr
   
   Exit Sub
ActualizaCtaObjSaldoErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:ActualizaCtaObjSaldo Method")
End Sub

Public Sub EliminaCtaObjSaldo(Optional psCtaContCod As String = "", Optional psFecha As String = "", Optional psSimbolo As String = "=")
Dim sCondCta   As String
Dim sCondFecha As String
   On Error GoTo EliminaCtaObjSaldoErr
      If psCtaContCod <> "" Then
         sCondCta = "WHERE cCtaContCod = '" & psCtaContCod & "' "
      End If
      If psFecha <> "" Then
         sCondFecha = IIf(sCondCta = "", "WHERE", "AND") & " dCtaObjSaldoFecha " & psSimbolo & " '" & psFecha & "'"
      End If
      psSql = "DELETE CtaObjSaldo " & sCondCta & sCondFecha
      dbConec.Ejecutar psSql
   Exit Sub
EliminaCtaObjSaldoErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:EliminaCtaObjSaldo Method")
End Sub

Public Function GeneraSaldosContables(psFecha As String) As Boolean
   On Error GoTo GeneraSaldosContablesErr
      dbConec.Ejecutar "spSaldosContab '" & psFecha & "'"
Exit Function
GeneraSaldosContablesErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:GeneraSaldosContables Method")
End Function

Public Function PermiteActualizarSaldos(psCodUser As String) As Boolean
   On Error GoTo PermiteActualizarSaldosErr
   psSql = "SELECT cs.bCtaSaldoValido, cs.cCodUser, p.cPersNombre " _
         & "FROM CtaSaldoEstad cs JOIN RRHH RH ON RH.cUser = cs.cCodUser " _
         & "     JOIN Persona P ON P.cPersCod = RH.cPersCod " _
         & "WHERE  dCtaSaldoFecha = (SELECT MAX(dCtaSaldoFecha) FROM CtaSaldoEstad)"
   Set prs = dbConec.CargaRecordSet(psSql)
   If Not prs.EOF Then
      If Not prs!bCtaSaldoValido Then
         If prs!cCodUser <> psCodUser Then
            PermiteActualizarSaldos = False
            Err.Raise "50001", "DCtaSaldo", UCase(prs!cPersNombre) & " se encuentra Actualizando Saldos"
         Else
            PermiteActualizarSaldos = True
         End If
      Else
         PermiteActualizarSaldos = True
      End If
   Else
      PermiteActualizarSaldos = True
   End If
Exit Function
PermiteActualizarSaldosErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:PermiteActualizarSaldos Method")
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

Dim oIni As New ClasIni
   sConexion = oIni.CadenaConexion
   sCentralCom = oIni.BaseComunes
Set oIni = Nothing
   Set dbConec = New DConecta
   dbConec.AbreConexion sConexion
End Sub

Private Sub Class_Terminate()
dbConec.CierraConexion
Set dbConec = Nothing
End Sub

