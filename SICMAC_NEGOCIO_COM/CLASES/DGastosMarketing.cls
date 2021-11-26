VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DGastosMarketing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'VAPI SEGUN ERS 082-2014

Option Explicit
    Dim oCon As New DConecta

Public Function getOcon() As DConecta
    Set getOcon = oCon
End Function

Public Function RecuperaBienesServicios(ByVal pCodigoComodin, Optional pConSaldo As Integer, Optional pnAlma As Integer) As ADODB.Recordset
On Error GoTo ErrRecuperaBienesServicios
    Dim sqlA  As String
    If oCon.AbreConexion = False Then Exit Function
    sqlA = "EXEC stp_selBienesServicios " & "'" & pCodigoComodin & "', " & pConSaldo & ", " & pnAlma
    Set RecuperaBienesServicios = oCon.CargaRecordSet(sqlA)
    oCon.CierraConexion
    Exit Function
ErrRecuperaBienesServicios:
    Err.Raise Err.Number, "Recupera los bienes y servicios", Err.Description
End Function

Public Function RecuperaCampanas() As ADODB.Recordset
On Error GoTo ErrRecuperaCampanas
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    If oCon.AbreConexion = False Then Exit Function
    Dim sSql As String
    sSql = "Exec stp_sel_ObtieneCampanas"
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set RecuperaCampanas = rs
    Exit Function
ErrRecuperaCampanas:
    Err.Raise Err.Number, "Recupera las campañas", Err.Description
End Function

Public Function RecuperaCombosXCampana(ByVal nIdCampana As String) As ADODB.Recordset
On Error GoTo ErrRecuperaCampanas
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    If oCon.AbreConexion = False Then Exit Function
    Dim sSql As String
    sSql = "Exec ERS0822014_stp_sel_comboXcampana " & nIdCampana
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set RecuperaCombosXCampana = rs
    Exit Function
ErrRecuperaCampanas:
    Err.Raise Err.Number, "Recupera los combos por campañas", Err.Description
End Function
Public Function InsertaComboCampana(ByVal idCampana As Integer, ByVal cComboDescripcion As String, ByVal bDesembolso As Integer, ByVal bApertura As Integer, ByVal bSoles As Integer, ByVal bDolares As Integer, ByVal nMinSoles As String, ByVal nMaxSoles As String, ByVal nMinDolares As String, ByVal nMaxDolares As String) As Integer
    On Error GoTo ErrInsertaComboCampana
    Dim sql As String
    Dim RsSource  As ADODB.Recordset
    Set RsSource = New ADODB.Recordset
    sql = "EXEC ERS0822014_stp_ins_comboxcampana " _
                    & idCampana & ", " _
             & " '" & cComboDescripcion & "', " _
                    & bDesembolso & ", " _
                    & bApertura & ", " _
                    & bSoles & ", " _
                    & bDolares & ", " _
                    & nMinSoles & ", " _
                    & nMaxSoles & ", " _
                    & nMinDolares & ", " _
                    & nMaxDolares
    Set RsSource = oCon.CargaRecordSet(sql)
    InsertaComboCampana = RsSource!Cod
Exit Function
ErrInsertaComboCampana:
    Err.Raise Err.Number, "Insertar combo de campañas", Err.Description
End Function

Public Function InsertaDetalleCombo(ByVal nIdCombo As String, ByVal cBSCod As String, ByVal nCantidad As String) As Integer
On Error GoTo ErrInsertaDetalleCombo:
    Dim RsSource  As ADODB.Recordset
    Set RsSource = New ADODB.Recordset
    Dim sql As String
    sql = "EXEC ERS0822014_stp_ins_comboBienes " _
                            & nIdCombo & ", " _
                            & " '" & cBSCod & "', " _
                            & nCantidad
   Set RsSource = oCon.CargaRecordSet(sql)
   InsertaDetalleCombo = RsSource!Cod
Exit Function
ErrInsertaDetalleCombo:
    Err.Raise Err.Number, "Insertar los detalles combo de campañas", Err.Description
End Function

Public Function RecuperaComboBienes(ByVal nIdCombo As String) As ADODB.Recordset
On Error GoTo ErrRecuperaComboBienes
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    If oCon.AbreConexion = False Then Exit Function
    Dim sSql As String
    sSql = "Exec ERS0822014_stp_sel_comboBienes " & nIdCombo
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set RecuperaComboBienes = rs
    Exit Function
ErrRecuperaComboBienes:
    Err.Raise Err.Number, "Recupera los bienes por combos", Err.Description
End Function

Public Sub EliminaBienEnCombo(ByVal nIdComboBienes As String)
On Error GoTo ErrEliminaBienEnCombo:
     If oCon.AbreConexion = False Then Exit Sub
    Dim sql As String
    sql = "EXEC ERS0822014_stp_del_comboBienes " & nIdComboBienes
    oCon.Ejecutar (sql)
    oCon.CierraConexion
Exit Sub
ErrEliminaBienEnCombo:
    Err.Raise Err.Number, "Elimina un bien por combo", Err.Description
End Sub

Public Sub ActualizaComboxCampana(ByVal nIdCombo As Integer, ByVal idCampana As Integer, ByVal cComboDescripcion As String, ByVal bDesembolso As Integer, ByVal bApertura As Integer, ByVal bSoles As Integer, ByVal bDolares As Integer, ByVal nMinSoles As String, ByVal nMaxSoles As String, ByVal nMinDolares As String, ByVal nMaxDolares As String)
    On Error GoTo ErrActualizaComboxCampana:
     If oCon.AbreConexion = False Then Exit Sub
    Dim sql As String
    sql = "EXEC ERS0822014_stp_upd_comboxcampana " _
                    & nIdCombo & ", " _
                    & idCampana & ", " _
             & " '" & cComboDescripcion & "', " _
                    & bDesembolso & ", " _
                    & bApertura & ", " _
                    & bSoles & ", " _
                    & bDolares & ", " _
                    & nMinSoles & ", " _
                    & nMaxSoles & ", " _
                    & nMinDolares & ", " _
                    & nMaxDolares
    oCon.Ejecutar (sql)
    oCon.CierraConexion
Exit Sub
ErrActualizaComboxCampana:
    Err.Raise Err.Number, "Actualiza un combo de la campana", Err.Description
End Sub

Public Function ExisteDescripcionCombo(ByVal nIdCombo As String, ByVal nIdCampana As String, ByVal cComboDescripcion As String) As Integer
On Error GoTo ErrExisteDescripcionCombo:
    Dim RsSource  As ADODB.Recordset
    Set RsSource = New ADODB.Recordset
    If oCon.AbreConexion = False Then Exit Function
    Dim sql As String
    
    If nIdCombo = "" Then
        sql = "EXEC ERS0822014_stp_sel_ExisteDescripcionCombo null," & nIdCampana & "," & "'" & cComboDescripcion & "'"
    Else
        sql = "EXEC ERS0822014_stp_sel_ExisteDescripcionCombo " & nIdCombo & "," & nIdCampana & "," & "'" & cComboDescripcion & "'"
    End If
    
   Set RsSource = oCon.CargaRecordSet(sql)
   ExisteDescripcionCombo = RsSource!existe
   oCon.CierraConexion
Exit Function
ErrExisteDescripcionCombo:
    Err.Raise Err.Number, "Verifica si existe una descripción del combo", Err.Description
End Function

Public Sub EliminaComboCampana(ByVal nIdCombo As String)
On Error GoTo ErrEliminaComboCampana:
     If oCon.AbreConexion = False Then Exit Sub
    Dim sql As String
    sql = "EXEC ERS0822014_stp_del_comboxcampana " & nIdCombo
    oCon.Ejecutar (sql)
    oCon.CierraConexion
Exit Sub
ErrEliminaComboCampana:
    Err.Raise Err.Number, "Elimina un combo de campana", Err.Description
End Sub


Public Function InsertaEntregaCampana(ByVal cAgencia As String, ByVal cPersCodUser As String, ByVal cPersCodCliente As String, ByVal fecha As String, ByVal cGlosa As String, ByVal sMovNro As String, ByVal bDesembolso As Integer) As Integer
    On Error GoTo ErrInsertaEntregaCampana
    Dim sql As String
    Dim RsSource  As ADODB.Recordset
    Set RsSource = New ADODB.Recordset
    sql = "EXEC ERS0822014_stp_ins_entregacampana " _
                & " '" & cAgencia & "', " _
                & " '" & cPersCodUser & "', " _
                       & " '', " _
                & " '" & cPersCodCliente & "', " _
                & " '" & fecha & "', " _
                & " '" & cGlosa & "'," _
                & " '" & sMovNro & "'," _
                       & "0" & "," _
                       & bDesembolso

    Set RsSource = oCon.CargaRecordSet(sql)
    InsertaEntregaCampana = RsSource!Cod
Exit Function
ErrInsertaEntregaCampana:
    Err.Raise Err.Number, "Inserta una entrega de campana", Err.Description
End Function

Public Function InsertaDetalleEntregaCampana(ByVal nIdEntregaCampana As Integer, ByVal cBSCod As String, ByVal nIdCombo As String, ByVal nIdCampana As String, ByVal nCantidad As String, ByVal bCombo As Integer) As Integer
On Error GoTo ErrInsertaEntregaCombobienes:
    Dim RsSource  As ADODB.Recordset
    Set RsSource = New ADODB.Recordset
    Dim sql As String
    sql = "EXEC ERS0822014_stp_ins_detalleentregacampana " _
                           & nIdEntregaCampana & ", " _
                           & " '" & cBSCod & "', " _
                           & nIdCombo & ", " _
                           & nIdCampana & ", " _
                           & nCantidad & ", " _
                           & bCombo
   Set RsSource = oCon.CargaRecordSet(sql)
    InsertaDetalleEntregaCampana = RsSource!Cod
Exit Function
ErrInsertaEntregaCombobienes:
    Err.Raise Err.Number, "Inserta los detalles de la entregas", Err.Description
End Function

Public Function RecuperaComboBienesInserta(ByVal nIdCombo As String) As ADODB.Recordset
On Error GoTo ErrRecuperaComboBienes
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sSql As String
    sSql = "Exec ERS0822014_stp_sel_comboBienes " & nIdCombo
    Set rs = oCon.CargaRecordSet(sSql)
    Set RecuperaComboBienesInserta = rs
    Exit Function
ErrRecuperaComboBienes:
    Err.Raise Err.Number, "Recupera los bienes por combos al insertar", Err.Description
End Function

Public Function RecuperaCampanaPorCuenta(ByVal cCtaCod As String, ByVal bDesembolso) As ADODB.Recordset
    On Error GoTo ErrRecuperaIdCampanaPorCuenta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    If oCon.AbreConexion = False Then Exit Function
    Dim sSql As String
    sSql = "Exec ERS0822014_stp_sel_idCampanaPorcCta  '" & cCtaCod & "', " & IIf(bDesembolso, 1, 0)
    Set rs = oCon.CargaRecordSet(sSql)
    Set RecuperaCampanaPorCuenta = rs
    oCon.CierraConexion
    Exit Function
ErrRecuperaIdCampanaPorCuenta:
    Err.Raise Err.Number, "Recupera el idCampaña por Nro de Cuenta", Err.Description
End Function

Public Function RecuperaCombosXCampanaCondicion(ByVal nIdCampana As String, ByVal esDesembolso As Boolean, ByVal Moneda As Integer, ByVal Monto) As ADODB.Recordset
On Error GoTo ErrRecuperaCampanas
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    If oCon.AbreConexion = False Then Exit Function
    Dim sSql As String
    
    Dim PesDesembolso As Integer: PesDesembolso = IIf(esDesembolso, 1, 0)
    
    sSql = "Exec ERS0822014_stp_sel_comboXcampanaCondicion " _
                                                            & nIdCampana & ", " _
                                                            & PesDesembolso & ", " _
                                                            & Moneda & ", " _
                                                            & Monto
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set RecuperaCombosXCampanaCondicion = rs
    Exit Function
ErrRecuperaCampanas:
    Err.Raise Err.Number, "Recupera los combos por campañas", Err.Description
End Function
