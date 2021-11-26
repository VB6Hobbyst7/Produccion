VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DDocRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***********************************************************************
'** Nombre : DDocRec
'** Descripción : Clase de Cheques  creado según RFC117-2012
'** Creación : EJVG, 20121124 09:00:00 AM
'***********************************************************************
Option Explicit
Dim oConn As DConecta
Dim fsSql As String

Private Sub Class_Initialize()
    Set oConn = New DConecta
    If Not oConn.AbreConexion Then
        Call RaiseError(MyUnhandledError, "DDocRec:Initialize Method. Error en Conexion a Base de datos")
    End If
End Sub
Private Sub Class_Terminate()
    oConn.CierraConexion
    Set oConn = Nothing
End Sub
Public Sub dBeginTrans()
    oConn.BeginTrans
End Sub
Public Sub dRollbackTrans()
    oConn.RollbackTrans
    oConn.Ejecutar "SET TRANSACTION ISOLATION LEVEL READ COMMITTED"
End Sub
Public Sub dCommitTrans()
    oConn.CommitTrans
End Sub
Public Function InsertaChequeTalonario(ByVal psIFTpo As String, ByVal psPersCod As String, ByVal psCtaIFCod As String, ByVal psUser As String, ByVal pdFecha As Date) As Long
On Error GoTo ErrInsertaChequeTalonario
    Dim rs As New ADODB.Recordset
    fsSql = "Exec stp_ins_InsertaTalonarioCheque '" & psIFTpo & "','" & psPersCod & "','" & psCtaIFCod & "','" & psUser & "','" & Format(pdFecha, "yyyymmdd hh:mm:ss") & "'"
    Set rs = oConn.CargaRecordSet(fsSql)
    If Not RSVacio(rs) Then
        InsertaChequeTalonario = rs!nTalonarioId
    Else
        InsertaChequeTalonario = 0
    End If
    Set rs = Nothing
    Exit Function
ErrInsertaChequeTalonario:
    Err.Raise Err.Number, "InsertaChequeTalonario", Err.Description
End Function
Public Sub InsertaChequeTalonarioDet(ByVal pnTalonarioId As Long, ByVal pnNroCheque As Long)
On Error GoTo ErrInsertaChequeTalonarioDet
    fsSql = "Exec stp_ins_InsertaTalonarioChequeDet " & pnTalonarioId & ",'" & Format(pnNroCheque, "00000000") & "'"
    oConn.Ejecutar fsSql
    Exit Sub
ErrInsertaChequeTalonarioDet:
    Err.Raise Err.Number, "InsertaChequeTalonarioDet", Err.Description
End Sub
Public Function RecuperaChequesTalonario(ByVal psIFTpo As String, ByVal psPersCod As String, ByVal psCtaIFCod As String, ByVal pnNroChequeIni As Long, ByVal pnNroChequeFin As Long) As ADODB.Recordset
On Error GoTo ErrRecuperaChequesTalonario
    Dim oConecta As New DConecta
    Dim lsSql As String
    
    lsSql = "Exec stp_sel_RecuperaChequesTalonario '" & psIFTpo & "','" & psPersCod & "','" & psCtaIFCod & "'," & pnNroChequeIni & "," & pnNroChequeFin
    oConecta.AbreConexion
    Set RecuperaChequesTalonario = oConecta.CargaRecordSet(lsSql)
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function
ErrRecuperaChequesTalonario:
    Err.Raise Err.Number, "RecuperaChequesTalonario", Err.Description
End Function
Public Function RecuperaChequeMovDoc(ByVal psIFTpo As String, ByVal psPersCod As String, ByVal psCtaIFCod As String, ByVal pcNroCheque As String) As ADODB.Recordset
On Error GoTo ErrRecuperaChequeMovDoc
    Dim oConecta As New DConecta
    Dim lsSql As String
    
    lsSql = "Exec stp_sel_RecuperaChequeMovDoc '" & psIFTpo & "','" & psPersCod & "','" & psCtaIFCod & "','" & pcNroCheque & "'"
    oConecta.AbreConexion
    Set RecuperaChequeMovDoc = oConecta.CargaRecordSet(lsSql)
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function
ErrRecuperaChequeMovDoc:
    Err.Raise Err.Number, "RecuperaChequeMovDoc", Err.Description
End Function
Public Function RecuperaTalonarioChequePaArbol(ByVal psIFTpo As String, ByVal psPersCod As String, ByVal psCtaIFCod As String) As ADODB.Recordset
On Error GoTo ErrRecuperaTalonarioCheque
    Dim oConecta As New DConecta
    Dim lsSql As String
    
    lsSql = "Exec stp_sel_TalonarioChequeArbol '" & psIFTpo & "','" & psPersCod & "','" & psCtaIFCod & "'"
    oConecta.AbreConexion
    Set RecuperaTalonarioChequePaArbol = oConecta.CargaRecordSet(lsSql)
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function
ErrRecuperaTalonarioCheque:
    Err.Raise Err.Number, "RecuperaTalonarioCheque", Err.Description
End Function
Public Function RecuperaChequesMantenimiento(ByVal pnTodosIF As Integer, ByVal psIFTpo As String, ByVal psPersCod As String, ByVal psCtaIFCod As String, ByVal pnTodosNroCheque As Integer, ByVal pnNroChequeIni As Long, ByVal pnNroChequeFin As Long, ByVal pnTodosFechaReg As Integer, ByVal pdFechaReg As Date) As ADODB.Recordset
On Error GoTo ErrRecuperaChequesMantenimiento
    Dim oConecta As New DConecta
    Dim lsSql As String
    
    lsSql = "Exec stp_sel_ChequesMantenimiento " & pnTodosIF & ",'" & psIFTpo & "','" & psPersCod & "','" & psCtaIFCod & "'," & pnTodosNroCheque & "," & pnNroChequeIni & "," & pnNroChequeFin & "," & pnTodosFechaReg & ",'" & Format(pdFechaReg, "yyyymmdd") & "'"
    oConecta.AbreConexion
    Set RecuperaChequesMantenimiento = oConecta.CargaRecordSet(lsSql)
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function
ErrRecuperaChequesMantenimiento:
    Err.Raise Err.Number, "RecuperaChequesMantenimiento", Err.Description
End Function
Public Sub ActualizaCheque(ByVal psIFTpo As String, ByVal psPersCod As String, ByVal psCtaIFCod As String, ByVal psNroCheque As String, Optional ByVal pnEstado As Integer = 0, Optional ByVal psMotivoElimina As String = "", Optional ByVal pdFechaElimina As Date = "1900-01-01", Optional ByVal psMotivoAnula As String = "", Optional ByVal pdFechaAnula As Date = "1900-01-01")
On Error GoTo ErrActualizaCheque
    Dim oConecta As New DConecta
    Dim lsSql As String
    Dim lsCampos As String
    
    If pnEstado <> 0 Then
        lsCampos = "D.nEstado = " & pnEstado & ","
    End If
    If psMotivoElimina <> "" Then
        lsCampos = lsCampos & "D.cMotivoElimina = '" & psMotivoElimina & "',"
    End If
    If DateDiff("D", pdFechaElimina, "1900-01-01") <> 0 Then
        lsCampos = lsCampos & "D.dFechaElimina = '" & Format(pdFechaElimina, "yyyymmdd hh:mm:ss") & "',"
    End If
    If psMotivoAnula <> "" Then
        lsCampos = lsCampos & "D.cMotivoAnula = '" & psMotivoAnula & "',"
    End If
    If DateDiff("D", pdFechaAnula, "1900-01-01") <> 0 Then
        lsCampos = lsCampos & "D.dFechaAnula = '" & Format(pdFechaAnula, "yyyymmdd hh:mm:ss") & "',"
    End If

    If Len(lsCampos) = 0 Then
        Exit Sub
    Else
        lsCampos = Mid(lsCampos, 1, Len(lsCampos) - 1)
    End If
    
    lsSql = "UPDATE D Set " & lsCampos & " FROM ChequeTalonario C INNER JOIN ChequeTalonarioDet D on C.nTalonarioId = D.nTalonarioId WHERE C.cIFTpo = '" & psIFTpo & "' And C.cPersCod = '" & psPersCod & "' And C.cCtaIFCod = '" & psCtaIFCod & "' And D.cNroCheque = '" & psNroCheque & "'"
    oConecta.AbreConexion
    oConecta.Ejecutar (lsSql)
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Sub
ErrActualizaCheque:
    Err.Raise Err.Number, "ActualizaCheque", Err.Description
End Sub
'EJVG20140125 ***
Public Function ListaChequeNegxValorizacion(ByVal pnMoneda As Integer) As ADODB.Recordset
On Error GoTo ErrListaChequeNegxValorizacion
    Dim oConecta As New DConecta
    Dim lsSql As String
    lsSql = "Exec stp_sel_ERS1262013_ListaChequexValorizacion " & pnMoneda
    oConecta.AbreConexion
    Set ListaChequeNegxValorizacion = oConecta.CargaRecordSet(lsSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrListaChequeNegxValorizacion:
    Err.Raise Err.Number, "ListaChequeNegxValorizacion", Err.Description
End Function
Public Function ListaChequeNegxExtornoValorizacion(ByVal pnMoneda As Moneda, ByVal pdFechaValorizacion As Date) As ADODB.Recordset
On Error GoTo ErrListaChequeNegxExtornoValorizacion
    Dim oConecta As New DConecta
    Dim lsSql As String
    lsSql = "Exec stp_sel_ERS1262013_ListaChequexExtornoValorizacion " & pnMoneda & ",'" & Format(pdFechaValorizacion, "yyyymmdd") & "'"
    oConecta.AbreConexion
    Set ListaChequeNegxExtornoValorizacion = oConecta.CargaRecordSet(lsSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrListaChequeNegxExtornoValorizacion:
    Err.Raise Err.Number, "ListaChequeNegxExtornoValorizacion", Err.Description
End Function
Public Function ListaChequeNegxExtornoRechazo(ByVal pnMoneda As Moneda, ByVal pdFechaRechazo As Date) As ADODB.Recordset
On Error GoTo ErrListaChequeNegxExtornoRechazo
    Dim oConecta As New DConecta
    Dim lsSql As String
    lsSql = "Exec stp_sel_ERS1262013_ListaChequexExtornoRechazo " & pnMoneda & ",'" & Format(pdFechaRechazo, "yyyymmdd") & "'"
    oConecta.AbreConexion
    Set ListaChequeNegxExtornoRechazo = oConecta.CargaRecordSet(lsSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrListaChequeNegxExtornoRechazo:
    Err.Raise Err.Number, "ListaChequeNegxExtornoRechazo", Err.Description
End Function
'END EJVG *******
