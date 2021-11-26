VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMRCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
Dim oCon As COMConecta.DCOMConecta
Dim sSQL As String

Private Sub Class_Initialize()
Set oCon = New COMConecta.DCOMConecta
If oCon.AbreConexion = False Then
    Err.Raise Err.Number, "Error al Abrir la Conexion", Err.Description
End If

End Sub

Private Sub Class_Terminate()
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Function EliminaParametro(ByVal psServConsol As String, _
                                ByVal psMes As String)

On Error GoTo ErrorEliminaParametro
sSQL = "Delete " & psServConsol & "RCDParametro WHERE CMES='" & psMes & "' "
        
oCon.Ejecutar (sSQL)
Exit Function

ErrorEliminaParametro:
    Err.Raise Err.Number, "Elimina Parametro", Err.Description
End Function

Public Function InsertaParametro(ByVal psServConsol As String, _
                                ByVal pcMes As String, _
                                ByVal pdFecha As Date, _
                                ByVal pnMontoMin As Double, _
                                ByVal pnCambioFijo As Double, _
                                ByVal pnUIT As Double, _
                                ByVal pcCodUsu As String, _
                                ByVal pdFecMod As Date)

Dim oGen As COMDConstSistema.DCOMGeneral
On Error GoTo ErrorInsertaParametro

Set oGen = New COMDConstSistema.DCOMGeneral

sSQL = " INSERT INTO " & psServConsol & "RCDParametro (cMes,dFecha,nMontoMin,nCambioFijo,nUIT,cCodUsu,dFecMod) " _
            & " VALUES('" & pcMes & "','" & Format(pdFecha, "mm/dd/yyyy") & "'," _
            & pnMontoMin & "," & pnCambioFijo & "," & pnUIT & ",'" & pcCodUsu & "','" _
            & oGen.FechaHora(pdFecMod) & "')"

Set oGen = Nothing
oCon.Ejecutar (sSQL)
Exit Function

ErrorInsertaParametro:
    Err.Raise Err.Number, "Inserta Parametro", Err.Description
End Function

Public Function ModificaParametro(ByVal psServConsol As String, _
                                ByVal pcMes As String, _
                                ByVal pdFecha As Date, _
                                ByVal pnMontoMin As Double, _
                                ByVal pnCambioFijo As Double, _
                                ByVal pnUIT As Double, _
                                ByVal pcCodUsu As String, _
                                ByVal pdFecMod As Date)

Dim oGen As COMDConstSistema.DCOMGeneral

On Error GoTo ErrorModificaParametro
                
Set oGen = New COMDConstSistema.DCOMGeneral

sSQL = "UPDATE " & psServConsol & "RCDParametro SET " _
        & "cMes='" & pcMes & "'," & "dFecha='" & Format(pdFecha, "mm/dd/yyyy") & "'," _
        & "nMontoMin=" & pnMontoMin & "," & "nCambioFijo=" & pnCambioFijo & "," _
        & "nUIT =" & pnUIT & "," _
        & "cCodUsu='" & pcCodUsu & "'," & "dFecMod='" & oGen.FechaHora(pdFecMod) & "' " _
        & "WHERE cMes='" & pcMes & "' "

Set oGen = Nothing

oCon.Ejecutar (sSQL)
Exit Function

ErrorModificaParametro:
    Err.Raise Err.Number, "Modifica Parametro", Err.Description
End Function

Public Function VerificaDuplicado(ByVal psServConsol As String, _
                                ByVal psMes As String) As Boolean
Dim rs As ADODB.Recordset

On Error GoTo ErrorVerificaDuplicado

sSQL = "Select cMes From " & psServConsol & "RCDParametro where cMes='" & psMes & "'"

Set rs = oCon.CargaRecordSet(sSQL)
    
If rs.BOF And rs.EOF Then
    VerificaDuplicado = False
Else
    VerificaDuplicado = True
End If
rs.Close
Set rs = Nothing
Exit Function

ErrorVerificaDuplicado:
    Err.Raise Err.Number, "Verifica Duplicado", Err.Description

End Function

Public Function ObtenerParametros(ByVal psServConsol As String) As ADODB.Recordset

On Error GoTo ErrorObtenerParametro
sSQL = "Select * from " & psServConsol & "RCDParametro Order By dFecha Desc"
Set ObtenerParametros = oCon.CargaRecordSet(sSQL)

Exit Function

ErrorObtenerParametro:
    Err.Raise Err.Number, "Obtener Parametro", Err.Description
End Function

Public Function ActualizarMesAnterior(ByVal pdFecDataFM As Date, _
                                      ByVal psServConsol As String, _
                                      ByRef psMensaje As String)

Dim lsSQL As String
Dim rs As ADODB.Recordset
Dim rsCodAux As ADODB.Recordset
Dim lnTotal As Long, J As Long
Dim lnNuevos As Long, lnModif As Long
Dim rs1 As ADODB.Recordset
Dim lbNuevoInsert As Boolean
Dim lsCodigoPersona As String
Dim gdFecDataFMAnt As Date

On Error GoTo ErrorActualizarMesAnterior
    
    gdFecDataFMAnt = DateAdd("m", -1, pdFecDataFM)
    lsSQL = "SELECT * FROM " & psServConsol & "sysobjects WHERE name = '" & "RCDvc" & Format(gdFecDataFMAnt, "yyyymm") & "01'"
    
    Set rs = oCon.CargaRecordSet(lsSQL)
    
    If rs.EOF And rs.BOF Then
        psMensaje = "Tabla [RCDvc" & Format(gdFecDataFMAnt, "yyyymm") & "01]  No existe, Por favor comuniquese con el Dpto de Sistemas"
        Set rs = Nothing
        Exit Function
    End If

    Set rs = Nothing
    
    lsSQL = "SELECT * FROM " & psServConsol & "RCDvc" & Format(gdFecDataFMAnt, "yyyymm") & "01 "
    
    Set rs = oCon.CargaRecordSet(lsSQL)
    
    lnNuevos = 0
    lnModif = 0
    If Not (rs.BOF And rs.EOF) Then
        Do While Not rs.EOF
            J = J + 1
            '*** BUSCO EL CODIGO DE PERSONA (TABLA AUXILIAR)
            lsSQL = "SELECT * FROM " & psServConsol & "RCDCodigoAux WHERE cCodAux='" & Trim(rs!cPersCod) & "'"
            
            Set rsCodAux = oCon.CargaRecordSet(lsSQL)
            If Not (rsCodAux.BOF And rsCodAux.EOF) Then
                lsCodigoPersona = Trim(rsCodAux!cPersCod)
            Else
                lsCodigoPersona = Trim(rs!cPersCod)
            End If
            rsCodAux.Close
            Set rsCodAux = Nothing
            '********
            
            lsSQL = "SELECT cPersCod FROM " & psServConsol & "RCDMaestroPersona WHERE cPersCod ='" & lsCodigoPersona & "' "
            Set rs1 = oCon.CargaRecordSet(lsSQL)
            If rs1.BOF And rs1.EOF Then ' No existe
                lbNuevoInsert = True
            Else
                lbNuevoInsert = False
            End If
            rs1.Close
            Set rs1 = Nothing
            
            If lbNuevoInsert = True Then
                lnNuevos = lnNuevos + 1
                
                lsSQL = "INSERT INTO " & psServConsol & "RCDMaestroPersona (cPersCod, cCodUnico, cCodSBS, cPersNom, " _
                    & " cActEcon, cCodRegPub, cTidoTr, cNudoTr, cTiDoci, cNuDoci, cTipPers, cResid, " _
                    & " cMagEmp, cAccionista, cRelInst, cPaisNac, cSiglas) " _
                    & " VALUES ('" & lsCodigoPersona & "','" & Trim(rs!cPersCod) & "','" & Trim(rs!cCodSBS) & "','" _
                    & Trim(Replace(rs!cPersNom, "'", "''")) & "','" & Trim(rs!cActEcon) & "','" & IIf(IsNull(rs!ccodregpub), "", rs!ccodregpub) & "','" _
                    & IIf(IsNull(rs!cTidoTr), "", IIf(Trim(rs!cTidoTr) = "4", "", Trim(rs!cTidoTr))) & "','" _
                    & IIf(IsNull(rs!cNudoTr), "", Trim(rs!cNudoTr)) & "','" _
                    & IIf(IsNull(rs!ctidoci), "", IIf(Trim(rs!ctidoci) = "9", "", Trim(rs!ctidoci))) & "','" _
                    & IIf(IsNull(rs!cnudoci), "", Trim(rs!cnudoci)) & "','" _
                    & Trim(rs!cTipPers) & "','" & Trim(rs!cResid) & "','" & IIf(IsNull(rs!cMagEmp), "", Trim(rs!cMagEmp)) & "','" _
                    & IIf(IsNull(rs!cAccionista), "", Trim(rs!cAccionista)) & "','" & IIf(IsNull(rs!cRelInst), "", Trim(rs!cRelInst)) & "','" _
                    & IIf(IsNull(rs!cPaisNac), "", Trim(rs!cPaisNac)) & "','" _
                    & IIf(IsNull(rs!cSiglas), "", Trim(rs!cSiglas)) & "' ) "
                
                oCon.Ejecutar (lsSQL)
            Else
                lnModif = lnModif + 1
                
            End If
            rs.MoveNext

        Loop
    End If
    rs.Close
    Set rs = Nothing

Exit Function

ErrorActualizarMesAnterior:
    Err.Raise Err.Number, "Error", Err.Description
        
End Function

Public Function ModificarMaestroPersona(ByVal psServConsol As String, _
                                        ByVal psCodSBS As String, _
                                        ByVal psCodDEUDOR As String)

On Error GoTo ErrorModificarMaestroPersona

sSQL = "UPDATE " & psServConsol & "RCDMaestroPersona SET cCodSBS ='" & psCodSBS & "' " _
                    & " WHERE cCodUnico='" & psCodDEUDOR & "' "

oCon.Ejecutar (sSQL)
Exit Function

ErrorModificarMaestroPersona:
    Err.Raise Err.Number, "Modificar Maestro Persona", Err.Description
                
End Function

Public Function ModificarMaestroPersonaLote(ByVal psServConsol As String, _
                                            ByVal prsDatos As ADODB.Recordset)
On Error GoTo ErrorModificarMaestroPersonaLote

prsDatos.MoveFirst
While Not prsDatos.EOF
    Call ModificarMaestroPersona(psServConsol, prsDatos("cCodSBS"), prsDatos("cCodUnico"))
    prsDatos.MoveNext
Wend
Exit Function
ErrorModificarMaestroPersonaLote:
    Err.Raise Err.Number, "Modificar Maestro Persona Lote", Err.Description
End Function

Public Function ActualizaPersonas_RCDMaestro(ByVal psServConsol As String)

Dim rs As ADODB.Recordset
Dim rsCodigo As ADODB.Recordset
Dim lnContTotal  As Long, lnCont As Long
Dim lsCodPers As String

On Error GoTo ErrorActualizaPersonas_RCDMaestro

sSQL = "Select * From " & psServConsol & "RCDMaestroPersona "
    Set rs = oCon.CargaRecordSet(sSQL)

    If Not rs.EOF Then
        Do While Not rs.EOF
            
            ' Emite Codigo de Persona
            sSQL = "Select * From " & psServConsol & "RCDCodigoAux Where cCodAux='" & Trim(rs!cCodUnico) & "' "
            Set rsCodigo = oCon.CargaRecordSet(sSQL)
            If rsCodigo.BOF And rsCodigo.EOF Then  ' No existe coge el codigo Persona
                lsCodPers = Trim(rs!cPersCod)
            Else ' Existe coge el Codigo de la tabla Auxiliar
                lsCodPers = Trim(rsCodigo!cPersCod)
            End If
            rsCodigo.Close
            Set rsCodigo = Nothing
            ' Codigo SBS
            sSQL = "UPDATE PERSONA SET cPersCodSbs = '" & Trim(rs!cCodSBS) & "' " _
                & " WHERE cPersCod = '" & Trim(lsCodPers) & "' "
            
            oCon.Ejecutar (sSQL)
            
            rs.MoveNext
        Loop
    End If
rs.Close
Set rs = Nothing
Exit Function

ErrorActualizaPersonas_RCDMaestro:
    Err.Raise Err.Number, "Actualiza Personas desde RCDMaestro", Err.Description
End Function

Public Function CargarCtaNoPref(ByVal psServConsol As String) As String

Dim rs As ADODB.Recordset
Dim SQL As String

SQL = "SELECT cCodCta,cOpeCod FROM " & psServConsol & "OPECUENTAGAR WHERE COPECOD IN ('001002') and cDEBEHAB='H'"

Set rs = oCon.CargaRecordSet(SQL)
If Not rs.BOF And Not rs.EOF Then
    CargarCtaNoPref = Trim(rs!cCodCta)
End If
rs.Close
Set rs = Nothing

End Function

Public Function ModificarRCDMaestroPersona(ByVal psServConsol As String, _
                                            ByVal psPersNom As String, _
                                            ByVal psCodSBS As String, _
                                            ByVal psActEcon As String, _
                                            ByVal psTidoTr As String, _
                                            ByVal psNudoTr As String, _
                                            ByVal psTiDoci As String, _
                                            ByVal psNuDoci As String, _
                                            ByVal psCodRegPub As String, _
                                            ByVal psMagEmp As String, _
                                            ByVal psTipPers As String, _
                                            ByVal psSiglas As String, _
                                            ByVal psPersCod As String)
Dim lsSQL  As String

On Error GoTo ErrorGrabaMaestro


        lsSQL = "UPDATE " & psServConsol & "RCDMaestroPersona SET " _
                & "cpersnom='" & psPersNom & "'," _
                & "cCodSBS='" & psCodSBS & "'," _
                & "cActEcon='" & psActEcon & "'," _
                & "cTidoTr=" & IIf(Len(psTidoTr) = 0, "Null", "'" & psTidoTr & "'") & "," _
                & "cNudoTr=" & IIf(Len(psNudoTr) = 0, "Null", "'" & psNudoTr & "'") & "," _
                & "cTiDoci=" & IIf(Len(psTiDoci) = 0, "Null", "'" & psTiDoci & "'") & "," _
                & "cNuDoci=" & IIf(Len(psNuDoci) = 0, "Null", "'" & psNuDoci & "'") & "," _
                & "cCodRegPub=" & IIf(Len(psCodRegPub) = 0, "Null", "'" & psCodRegPub & "'") & "," _
                & "cMagEmp=" & IIf(Len(psMagEmp) = 0, "Null", "'" & psMagEmp & "'") & "," _
                & "cTipPers='" & psTipPers & "'," _
                & "cSiglas=" & IIf(Len(Trim(psSiglas)) = 0, "Null", "'" & psSiglas & "'") & " " _
                & "WHERE cPersCod='" & psPersCod & "'"
            
            oCon.Ejecutar (lsSQL)
            
Exit Function
ErrorGrabaMaestro:
    Err.Raise Err.Number, "Modificar RCDMaestroPersona", Err.Description

End Function

Public Function CargaClientes(ByVal psServConsol As String, _
                            ByVal pnSeleccion As Integer, _
                            ByVal psPersCod As String) As ADODB.Recordset
Dim lsSQL  As String

Select Case pnSeleccion
    Case 0
        lsSQL = "Select cPersCod From " & psServConsol & "RCDMaestroPersona Where cPersCod= '" & psPersCod & "'"
    Case 1
        lsSQL = "Select cPersCod From " & psServConsol & "RCDMaestroPersona Where cCodSBS = '" & psPersCod & "'"
End Select

Set CargaClientes = oCon.CargaRecordSet(lsSQL)
       
End Function

Public Function CargaDatosCliente(ByVal psServConsol As String, _
                                ByVal psPersCod As String) As ADODB.Recordset
Dim lsSQL As String

lsSQL = "Select * From " & psServConsol & "RCDMaestroPersona where cPersCod='" & psPersCod & "'"

Set CargaDatosCliente = oCon.CargaRecordSet(lsSQL)

End Function
