VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DColPActualizaBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'* Modulo de Colocaciones Pignoraticio
'* Clase con todas las Actualizaciones a BD
Option Explicit
Dim csConexion As String
Dim csNegocio As String
Dim csCentralPer As String
Dim csCentralCom As String
Dim csCentralImg As String
Dim csAdminist As String

Public coConex As DConecta

Private Sub Class_Initialize()
    Dim loIni As COMConecta.DCOMClasIni
    
    Set loIni = New COMConecta.DCOMClasIni
        csConexion = loIni.CadenaConexion
        csNegocio = loIni.BaseNegocio
        csCentralPer = loIni.BasePersonas
        csCentralCom = loIni.BaseComunes
        csCentralImg = loIni.BaseImagenes
        csAdminist = loIni.BaseAdministracion
    Set loIni = Nothing
    
Set coConex = New DConecta
If coConex.AbreConexion(csConexion) = False Then
    Call RaiseError(MyUnhandledError, "DColPCalculos:Initialize. Error en Conexion a Base de datos")
End If

End Sub

Private Sub Class_Terminate()
    coConex.CierraConexion
    Set coConex = Nothing
End Sub

Public Sub dBeginTrans()
    coConex.BeginTrans
End Sub

Public Sub dRollbackTrans()
    coConex.RollbackTrans
End Sub
Public Sub dCommitTrans()
    coConex.CommitTrans
End Sub

Public Function dEjecutar(psSql As String) As ADODB.Recordset
On Error GoTo ErrEjecutar
Set dEjecutar = coConex.Ejecutar(psSql)
Exit Function
ErrEjecutar:
   Call RaiseError(MyUnhandledError, "DConecta:Ejecutar Method")
End Function

'Ejecuta procesos Batch
Public Function dEjecutaBatch() As Integer
On Error GoTo EjecutaBatchErr
    coConex.EjecutarBatch
    Exit Function
EjecutaBatchErr:
    Call RaiseError(MyUnhandledError, "DColPActualizaBD - Ejecutar Batch")
End Function

Public Sub dInsertCredPignoraticio(ByVal psCuenta As String, _
        ByVal pnTasaInteres As Double, ByVal pnSaldo As Currency, _
        ByVal psFechaHora As String, ByVal pnPlazo As Integer, ByVal psFecVenc As String, _
        ByVal pnOroBruto As Double, ByVal pnOroNeto As Double, ByVal pnTasacion As Currency, _
        ByVal pnPiezas As Integer, ByVal psTipoContrato As String, ByVal psLote As String, _
        ByVal pnki14 As Double, ByVal pnki16 As Double, ByVal pnKi18 As Double, _
        ByVal pnKi21 As Double, ByVal psMovNro As String, _
        Optional ByVal pbEjecBatch As Boolean = False)


Dim lsSQL As String

'**Insert Producto
lsSQL = "INSERT Producto (cCtaCod,nTasaInteres,nSaldo,nPrdEstado,dPrdEstado,nTransacc) " _
    & "VALUES ('" & psCuenta & "'," & pnTasaInteres & "," & pnSaldo & "," & gColPEstRegis & ",'" & psFechaHora & "',1)"
If pbEjecBatch Then
    coConex.AdicionaCmdBatch lsSQL
Else
    coConex.Ejecutar lsSQL
End If
'**Insert Colocaciones
lsSQL = "INSERT Colocaciones (cCtaCod,nPlazo,dVenc,nMontoCol,nColocCalendCod,cUltimaActualizacion,cLineaCred,dVigencia) " _
    & "VALUES ('" & psCuenta & "'," & pnPlazo & ",'" & psFecVenc & "'," & pnSaldo & "," & gColocCalendCodPFCF & ",'" & psMovNro & "','" & "01011130501" & "','" & psFechaHora & "' )"
If pbEjecBatch Then
    coConex.AdicionaCmdBatch lsSQL
Else
    coConex.Ejecutar lsSQL
End If
'**Insert ColocPignoraticio
lsSQL = "INSERT ColocPignoraticio (cCtaCod,nOroBruto,nOroNeto,nPiezas,nTasacion,nNroRenov,nNroRemate,nNroDuplic,cPrdCtaTpo,cAgeBoveda,cLote) " _
     & "VALUES ('" & psCuenta & "'," & pnOroBruto & "," & pnOroNeto & "," & pnPiezas & "," & pnTasacion & ",0,0,0,'" & psTipoContrato & "','" & Mid(psCuenta, 4, 2) & "','" & psLote & "') "
If pbEjecBatch Then
    coConex.AdicionaCmdBatch lsSQL
Else
    coConex.Ejecutar lsSQL
End If
'**Insert ColocPigJoya
If pnki14 > 0 Then
    lsSQL = "INSERT ColocPigJoya (cCtaCod,cKilataje,nPesoOro) " _
          & "VALUES ('" & psCuenta & "','14'," & pnki14 & " )"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End If
If pnki16 > 0 Then
    lsSQL = "INSERT ColocPigJoya (cCtaCod,cKilataje,nPesoOro) " _
          & "VALUES ('" & psCuenta & "','16'," & pnki16 & " )"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End If
If pnKi18 > 0 Then
    lsSQL = "INSERT ColocPigJoya (cCtaCod,cKilataje,nPesoOro) " _
          & "VALUES ('" & psCuenta & "','18'," & pnKi18 & " )"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End If
If pnKi21 > 0 Then
    lsSQL = "INSERT ColocPigJoya (cCtaCod,cKilataje,nPesoOro) " _
          & "VALUES ('" & psCuenta & "','21'," & pnKi21 & " )"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End If

End Sub


Public Sub dInsertProducto(ByVal psctacod As String, ByVal pnTasaInteres As ColocTipoTasa, _
        ByVal pnSaldo As Currency, ByVal pnEstado As ColocEstado, ByVal psFecEstado As String, _
        ByVal pnTransac As Integer, _
        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT Producto (cCtaCod, nTasaInteres, nSaldo, nPrdEstado, dPrdEstado, nTransacc ) " _
        & "VALUES ('" & psctacod & "'," & pnTasaInteres & "," & pnSaldo & "," & pnEstado & ",'" & psFecEstado & "'," & pnTransac & ")"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dInsertColocaciones(ByVal psctacod As String, ByVal pnPlazo As Integer, _
        ByVal psFecVenc As String, ByVal pnMontoCol As Currency, ByVal pnTipoCalend As ColocTipoCalend, _
        ByVal psMovNro As String, ByVal psLineaCred As String, ByVal psFecVigencia As String, _
        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "INSERT Colocaciones (cCtaCod,nPlazo,dVenc,nMontoCol,nColocCalendCod,cUltimaActualizacion,cLineaCred, dVigencia) " _
        & "VALUES ('" & psctacod & "'," & pnPlazo & ",'" & psFecVenc & "'," & pnMontoCol & "," _
        & pnTipoCalend & ",'" & psMovNro & "','" & psLineaCred & "',''  )"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

Public Sub dInsertColocPignoraticio(ByVal psctacod As String, ByVal pnOroBruto As Currency, _
        ByVal pnOroNeto As Currency, ByVal pnPiezas As Integer, ByVal pnTasacion As Currency, _
        ByVal pnNroRenovac As Integer, ByVal pnNroDuplic As Integer, ByVal pnTipoCalend As ColocTipoCalend, _
        ByVal psAgeBoveda As String, ByVal psLote As String, Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "INSERT ColocPignoraticio (cCtaCod,nOroBruto,nOroNeto,nPiezas,nTasacion, " _
         & "nNroRenov,nNroRemate,nNroDuplic,cPrdCtaTpo,cAgeBoveda,cLote) " _
         & "VALUES ('" & psctacod & "'," & pnOroBruto & "," & pnOroNeto & "," & pnPiezas & "," & pnTasacion _
         & "," & pnNroRenovac & "," & pnNroDuplic & "," & pnTipoCalend & ",'" & psAgeBoveda & "','" & Replace(psLote, "'", "''") & "' )"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

Public Sub dInsertColocPigJoya(ByVal psctacod As String, ByVal psKilataje As String, _
        ByVal pnPesoOro As Currency, Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "INSERT ColocPigJoya (cCtaCod,cKilataje,nPesoOro )" _
         & "VALUES ('" & psctacod & "','" & psKilataje & "'," & pnPesoOro & " )"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

Public Sub dInsertCredPignoraticioEstado(ByVal psCuenta As String, ByVal pdFecha As Date, _
        ByVal pnEstado As String, ByVal pnCuotas As Integer, ByVal pnMonto As Double, _
        ByVal psDescrip As String, ByVal pnCalendario As Integer, ByVal pnPerFecFija As Integer, _
        ByVal pnPerGracia, ByVal pnPlazo As Integer, pnTipGracia As Integer, _
        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT ColocacEstado (cCtaCod,dPrdEstado,nPrdEstado,nCuotas,nMonto,cDescripcion,nColocCalendCod,nPeriodoFechaFija,nPeriodoGracia,nPlazo,nTipoGracia) " _
        & "VALUES ('" & psCuenta & "','" & pdFecha & "'," & pnEstado & "," & pnCuotas & "," & pnMonto & ",'" & psDescrip & "'," & pnCalendario & "," & pnPerFecFija & "," & pnPerGracia & "," & pnPlazo & "," & pnTipGracia & ")"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dInsertProductoPersona(ByVal psCuenta As String, ByVal psPersona As String, ByVal pnRelacion As ColocRelacPers, _
    Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT ProductoPersona (cCtaCod,cPersCod,nPrdPersRelac) " _
        & "VALUES ('" & psCuenta & "','" & psPersona & "'," & pnRelacion & ")"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    
End Sub

Public Sub dDeleteProductoPersona(ByVal psCuenta As String, ByVal psPersona As String, ByVal pnRelacion As ColocRelacPers, _
    Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = " DELETE ProductoPersona WHERE cCtaCod ='" & psCuenta & "' " _
          & " AND nPrdPersRelac = " & pnRelacion & " " _
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    
End Sub

'Insert ColocacEstado
Public Sub dInsertColocacEstado(ByVal psctacod As String, ByVal psFechaPrdEstado As String, _
        ByVal psPrdEstado As ColocEstado, ByVal pnCuotas As Integer, ByVal pnMonto As Currency, _
        ByVal psDescripcion As String, ByVal pnTipoCalendario As ColocTipoCalend, _
        ByVal pnPeriodoFechaFija As Integer, ByVal pnPeriodoGracia As Integer, _
        ByVal pnPlazo As Integer, ByVal pnTipoGracia As Integer, ByVal pnTipoDesembolso As Integer, _
        ByVal pnProxMes As Integer, ByVal pnCalendDinamico As Integer, ByVal pnMotivoRechazo As Integer, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT ColocacEstado (cCtaCod, dPrdEstado, nPrdEstado, nCuotas, nMonto, cDescripcion, " _
        & "nColocCalendCod, nPeriodoFechaFija, nPeriodoGracia, nPlazo, nTipoGracia, nTipoDesembolso, " _
        & "nProxMes, nCalendDinamico, nMotivoRechazo ) " _
        & "VALUES ('" & psctacod & "','" & psFechaPrdEstado & "'," & psPrdEstado & "," & pnCuotas & "," _
        & pnMonto & ",'" & psDescripcion & "'," & pnTipoCalendario & "," & pnPeriodoFechaFija & "," _
        & pnPeriodoGracia & "," & pnPlazo & "," & pnTipoGracia & "," & pnTipoDesembolso & "," _
        & pnProxMes & "," & pnCalendDinamico & "," & pnMotivoRechazo & ") "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub
'Insert ColocCalendario
Public Sub dInsertColocCalendario(ByVal psctacod As String, ByVal pnNroCalend As Integer, ByVal pnColocCalendApl As ColocCalendApl, _
        ByVal pnCuota As Integer, ByVal psFecVenc As String, ByVal pnColocCalendEstado As Integer, ByVal psDescrip As String, _
        ByVal psColocCalenFlag As String, Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "INSERT ColocCalendario (cCtaCod,nNroCalen,nColocCalendApl,nCuota,dVenc,nColocCalendEstado,cDescripcion,cColocCalenFlag) " _
         & "VALUES ('" & psctacod & "'," & pnNroCalend & "," & pnColocCalendApl & "," & pnCuota & ",'" & psFecVenc & "'," & pnColocCalendEstado & ",'" _
         & psDescrip & "','" & psColocCalenFlag & "' )"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

'Insert ColocCalendDet
Public Sub dInsertColocCalendDet(ByVal psctacod As String, ByVal pnNroCalend As Integer, ByVal pnColocCalendApl As ColocCalendApl, _
        ByVal pnCuota As Integer, ByVal pnColocConceptoCod As ColocConcepto, ByVal pnMonto As Currency, ByVal pnMontoPagado As Currency, _
        ByVal psFlag As String, Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "INSERT ColocCalendDet (cCtaCod,nNroCalen,nColocCalendApl,nCuota,nPrdConceptoCod,nMonto,nMontoPagado,cFlag) " _
         & "VALUES ('" & psctacod & "'," & pnNroCalend & "," & pnColocCalendApl & "," & pnCuota & "," & pnColocConceptoCod & "," & pnMonto & "," _
         & pnMontoPagado & ",'" & psFlag & "' ) "
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

'Insert ColocCalendDetPig
Public Sub dInsertColocCalendDetPig(ByVal psctacod As String, ByVal pnNroCalend As Integer, ByVal pnColocCalendApl As ColocCalendApl, _
        ByVal pnCuota As Integer, ByVal pnColocConceptoCod As ColocConcepto, ByVal pnMonto As Currency, ByVal pnMontoPagado As Currency, _
        ByVal psFlag As String, Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "INSERT ColocCalendDetPig (cCtaCod,nNroCalen,nColocCalendApl,nCuota,nPrdConceptoCod,nMonto,nMontoPagado,cFlag) " _
         & "VALUES ('" & psctacod & "'," & pnNroCalend & "," & pnColocCalendApl & "," & pnCuota & "," & pnColocConceptoCod & "," & pnMonto & "," _
         & pnMontoPagado & ",'" & psFlag & "' ) "
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

'Inserta Mov
Public Sub dInsertMov(ByVal psMovNro As String, ByVal psOpeCod As String, _
    ByVal psMovDesc As String, ByVal pnMovEstado As MovEstado, _
    Optional pnMovFlag As MovFlag = gMovFlagVigente, _
    Optional pbEjecBatch As Boolean = False)
    
Dim lsSQL As String
On Error GoTo InsertMovErr
    
    lsSQL = "INSERT Mov (cMovNro,cOpeCod,cMovDesc,nMovEstado, nMovFlag) " _
          & "VALUES ('" & psMovNro & "','" & psOpeCod & "','" & Replace(psMovDesc, "'", "''") & "'," _
          & pnMovEstado & "," & pnMovFlag & ")"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub

InsertMovErr:
    Call RaiseError(MyUnhandledError, "DColPActualizaBD Insert Mov")
End Sub

'Inserta el movimiento de lavado de dinero
Public Sub dInsertaMovLavDinero(ByVal nMovNro As Long, ByVal sPersCod As String)
Dim lsSQL As String

On Error GoTo InsertMovLav
lsSQL = "INSERT MovLavDinero (nMovNro,cPersCod) " _
    & "VALUES (" & nMovNro & ",'" & sPersCod & "')"
coConex.Ejecutar lsSQL
Exit Sub

InsertMovLav:
    Call RaiseError(MyUnhandledError, "DColPActualizaBD Insert Mov Lab")

End Sub

'Obtiene el nMovNro  apartir del cMovNro
Public Function dGetnMovNro(ByVal psMovNro As String) As Long
Dim lsSQL As String
Dim lrs As ADODB.Recordset
Set lrs = New ADODB.Recordset

lsSQL = "Select nMovNro From Mov where cMovNro ='" & psMovNro & "'"
Set lrs = coConex.CargaRecordSet(lsSQL)
If Not lrs.EOF And Not lrs.BOF Then
    dGetnMovNro = lrs!nMovNro
End If
lrs.Close
Set lrs = Nothing
End Function



'Inserta MovRef
Public Sub dInsertMovRef(ByVal pnMovNro As Long, ByVal pnMovNroRef As Long, _
    Optional pbEjecBatch As Boolean = False)
    
Dim lsSQL As String
On Error GoTo InsertMovRefErr
    
    lsSQL = "INSERT MovRef (nMovNro,nMovNroRef) " _
          & "VALUES (" & pnMovNro & "," & pnMovNroRef & ")"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub

InsertMovRefErr:
    Call RaiseError(MyUnhandledError, "DColPActualizaBD Insert Mov Ref")
End Sub

Public Sub dInsertMovCol(ByVal pnMovNro As Long, ByVal psOperacion As String, _
        ByVal psCuenta As String, ByVal pnNroCalend As Integer, ByVal pnMonto As Currency, _
        ByVal pnDiasMora As Integer, ByVal psMetLiq As String, ByVal pnPlazo As Integer, _
        ByVal pnCredEstado As Integer, ByVal pnSaldoCap As Currency, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT MovCol (nMovNro,cOpeCod,cCtaCod,nNroCalen,nMonto,nDiasMora,cMetLiq,nPlazo, nCredEstado, nSaldoCap) " _
        & "VALUES (" & pnMovNro & ",'" & psOperacion & "','" & psCuenta & "'," & pnNroCalend & "," _
        & pnMonto & "," & pnDiasMora & ",'" & psMetLiq & "'," & pnPlazo & "," & pnCredEstado & "," _
        & pnSaldoCap & " ) "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dInsertMovColDet(ByVal pnMovNro As Long, ByVal psOperacion As String, _
        ByVal psctacod As String, pnNroCalend As Integer, ByVal pnConcepto As ColocConcepto, _
        ByVal pnNroCuota As Integer, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT MovColDet (nMovNro,cOpeCod,cCtaCod,nNroCalen,nPrdConceptoCod,nNroCuota,nMonto) " _
        & "VALUES (" & pnMovNro & ",'" & psOperacion & "','" & psctacod & "'," & pnNroCalend & "," _
        & pnConcepto & "," & pnNroCuota & "," & pnMonto & " ) "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub


Public Sub dInsertMovDoc(ByVal pnMovNro As Long, ByVal pnDocTpo As Integer, ByVal psDocNro As String, _
        ByVal psFecDoc As String, Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT MovDoc (nMovNro, nDocTpo, cDocNro, dDocFecha) " _
        & "VALUES (" & pnMovNro & "," & pnDocTpo & ",'" & psDocNro & "','" & psFecDoc & "' ) "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub


Public Sub dInsertMovCMAC(ByVal pnMovNro As Long, ByVal psPersCod As String, _
        ByVal psIFTipo As String, ByVal pnMoneda As Moneda, ByVal psCtaIf As String, _
        ByVal psDocumento As String, ByVal psOpeCod As String, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False, Optional sCliente As String = "")

Dim lsSQL As String

    lsSQL = "INSERT MovCMAC (nMovNro, cPersCod, cIFTpo, nMoneda, cCtaIF, cDocumento, cOpeCod, nMonto, cCliente) " _
        & "VALUES (" & pnMovNro & ",'" & psPersCod & "','" & psIFTipo & "'," & pnMoneda & ",'" _
        & psCtaIf & "','" & psDocumento & "','" & psOpeCod & "'," & pnMonto & ", '" & sCliente & "') "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dInsertMovOpeVarias(ByVal pnMovNro As Long, ByVal pnMovImporte As Double, _
        ByVal psNroDoc As String, ByVal psReferencia As String, ByVal pnMoneda As Moneda, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "Insert MovOpeVarias (nMovNro,nMovImporte,cNroDoc,cReferencia,nMoneda)" _
        & " Values(" & pnMovNro & "," & pnMovImporte & ",'" & psNroDoc & "','" & psReferencia & "'," & pnMoneda & ")"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub


Public Sub dInsertProductoBloqueos(ByVal psctacod As String, ByVal pnBlqTpo As Integer, ByVal pnBlqMotivo As String, _
        ByVal psMovNro As String, ByVal psComentario As String, Optional ByVal psMovNroDbl As String = "@", _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT ProductoBloqueos (cCtaCod,nBlqTpo,nBlqMotivo,cMovNro,cComentario) " _
        & "VALUES ('" & psctacod & "'," & pnBlqTpo & "," & pnBlqMotivo & ",'" & psMovNro & "','" _
        & psComentario & "') "  ' --' ," & IIf(psMovNroDbl = "@", Null, ",'" & psMovNroDbl & "' ") & " ) "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dUpdateProductoBloqueos(ByVal psctacod As String, ByVal psMovNro As String, _
        Optional ByVal psMovNroDbl As String = "@", Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "UPDATE ProductoBloqueos SET "
    
    If psMovNroDbl <> "@" Then
        lsSQL = lsSQL & " cMovNroDbl = '" & psMovNroDbl & "',"
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cCtaCod ='" & psctacod & "' AND cMovNro ='" & psMovNro & "' "
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub


'**Update Producto
Public Sub dUpdateProducto(ByVal psctacod As String, Optional ByVal pnTasaInteres As Double = -1, _
        Optional ByVal pnSaldo As Currency = -1, Optional ByVal pnPrdEstado As ColocEstado = -1, _
        Optional ByVal psFechaHora As String = "@", Optional ByVal pnTransac As Integer = -1, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "UPDATE Producto SET "

    If pnTasaInteres <> -1 Then
        lsSQL = lsSQL & " nTasaInteres = " & pnTasaInteres & ","
    End If
    If pnSaldo <> -1 Then
        lsSQL = lsSQL & " nSaldo = " & pnSaldo & ","
    End If
    If pnPrdEstado <> -1 Then
        lsSQL = lsSQL & " nPrdEstado = " & pnPrdEstado & ","
    End If
    If psFechaHora <> "@" Then
        lsSQL = lsSQL & " dPrdEstado = '" & psFechaHora & "',"
    End If
    If pnTransac = -2 Then ' Si se le envia (-2) aumenta el nro de transaccion en uno
        lsSQL = lsSQL & " nTransacc = nTransacc + 1,"
    ElseIf pnTransac <> -1 Then
        lsSQL = lsSQL & " nTransacc = " & pnTransac & ","
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cCtaCod ='" & psctacod & "' "
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Colocaciones
Public Sub dUpdateColocaciones(ByVal psctacod As String, Optional ByVal pnPlazo As Integer = -1, _
        Optional ByVal psFecVenc As String = "@", Optional ByVal pnMontoCol As Currency = -1, Optional ByVal pnColocCalendCod As Integer = -1, _
        Optional ByVal psUltimaActualizacion As String = "@", Optional ByVal psLineaCred As String = "@", _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "UPDATE Colocaciones SET "

    If pnPlazo <> -1 Then
        lsSQL = lsSQL & " nPlazo = " & pnPlazo & ","
    End If
    If psFecVenc <> "@" Then
        lsSQL = lsSQL & " dVenc = '" & psFecVenc & "',"
    End If
    If pnMontoCol <> -1 Then
        lsSQL = lsSQL & " nMontoCol = " & pnMontoCol & ","
    End If
    If pnColocCalendCod <> -1 Then
        lsSQL = lsSQL & " nColocCalendCod = " & pnColocCalendCod & ","
    End If
    If psUltimaActualizacion <> "@" Then
        lsSQL = lsSQL & " cUltimaActualizacion = '" & psUltimaActualizacion & "',"
    End If
    If psLineaCred <> "@" Then
        lsSQL = lsSQL & " cLineaCred = '" & psLineaCred & "',"
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cCtaCod ='" & psctacod & "' "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

' Update ColocPignoraticio
Public Sub dUpdateColocPignoraticio(ByVal psctacod As String, Optional ByVal pnOroBruto As Double = -1, _
        Optional ByVal pnOroNeto As Double = -1, Optional ByVal pnPiezas As Integer = -1, Optional ByVal pnTasacion As Currency = -1, _
        Optional ByVal pnNroRenov As Integer = -1, Optional ByVal pnNroRemate As Integer = -1, Optional ByVal pnNroDuplic As Integer = -1, _
        Optional ByVal psPrdCtaTpo As String = "@", Optional ByVal psAgeBoveda As String = "@", Optional ByVal psLote As String = "@", _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "UPDATE ColocPignoraticio SET "

    If pnOroBruto <> -1 Then
        lsSQL = lsSQL & " nOroBruto = " & pnOroBruto & ","
    End If
    If pnOroNeto <> -1 Then
        lsSQL = lsSQL & " nOroNeto = " & pnOroNeto & ","
    End If
    If pnPiezas <> -1 Then
        lsSQL = lsSQL & " nPiezas = " & pnPiezas & ","
    End If
    If pnTasacion <> -1 Then
        lsSQL = lsSQL & " nTasacion = " & pnTasacion & ","
    End If
    If pnNroRenov <> -1 Then
        lsSQL = lsSQL & " nNroRenov = " & pnNroRenov & ","
    End If
    If pnNroRemate <> -1 Then
        lsSQL = lsSQL & " nNroRemate = " & pnNroRemate & ","
    End If
    If pnNroDuplic <> -1 Then
        lsSQL = lsSQL & " nNroDuplic = " & pnNroDuplic & ","
    End If
    If psPrdCtaTpo <> "@" Then
        lsSQL = lsSQL & " cPrdCtaTpo = '" & psPrdCtaTpo & "',"
    End If
    If psAgeBoveda <> "@" Then
        lsSQL = lsSQL & " cAgeBoveda = '" & psAgeBoveda & "',"
    End If
    If psLote <> "@" Then
        lsSQL = lsSQL & " cLote = '" & psLote & "',"
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cCtaCod ='" & psctacod & "' "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Colocaciones
Public Sub dUpdateMov(ByVal pnMovNro As Long, Optional ByVal psOpeCod As String = "@", _
    Optional ByVal psMovDesc As String = "@", Optional ByVal pnMovEstado As MovEstado = -1, _
    Optional ByVal pnMovFlag As MovFlag = -1, Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "UPDATE Mov SET "

    If psOpeCod <> "@" Then
        lsSQL = lsSQL & " cOpeCod = '" & psOpeCod & "',"
    End If
    If psMovDesc <> "@" Then
        lsSQL = lsSQL & " cMovDesc = '" & psMovDesc & "',"
    End If
    If pnMovEstado <> -1 Then
        lsSQL = lsSQL & " nMovEstado = " & pnMovEstado & ","
    End If
    If pnMovFlag <> -1 Then
        lsSQL = lsSQL & " nMovFlag = " & pnMovFlag & ","
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE nMovNro =" & pnMovNro & " "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub


Public Function dUltimaTransaccionProducto(ByVal psctacod As String) As Integer
Dim lsSQL As String
Dim lrs As ADODB.Recordset

On Error GoTo dError

lsSQL = "SELECT IsNull(nTransacc, 0 ) as nTransacc FROM Producto where cCtaCod ='" & psctacod & "' "

Set lrs = New ADODB.Recordset


Set lrs = coConex.CargaRecordSet(lsSQL)
    dUltimaTransaccionProducto = lrs!nTransacc
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Ultima Transaccion en <<dObtieneDatosContrato>>", Err.Description
    

End Function

'Insert ColocPigRecGar
Public Sub dInsertColocPigRecGar(ByVal psTpoProceso As String, ByVal psNroProceso As String, _
        ByVal psFecProceso As String, ByVal pnRecGar As Integer, ByVal psProcesoCadaAgencia As String, _
        Optional ByVal psFecParaAdjudicar As String = "@", Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT ColocPigRecGar (cTpoProceso, cNroProceso, dProceso, nRGEstado, cCodAge "
    If psFecParaAdjudicar = "@" Then
        lsSQL = lsSQL & " ) "
    Else
        lsSQL = lsSQL & " , dParaAdjudicar) "
    End If
    lsSQL = lsSQL & " VALUES ('" & psTpoProceso & "','" & psNroProceso & "','" & psFecProceso & "'," & pnRecGar & ",'" & psProcesoCadaAgencia & "'"
    If psFecParaAdjudicar = "@" Then
        lsSQL = lsSQL & " ) "
    Else
        lsSQL = lsSQL & " ,'" & psFecParaAdjudicar & "'  ) "
    End If
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

' Update ColocPigRecGar
Public Sub dUpdateColocPigRecGar(ByVal psTpoProceso As String, ByVal psNroProceso As String, _
        ByVal psProcesoCadaAgencia As String, Optional ByVal psFecProceso As String = "@", _
        Optional ByVal pnRGEstado As Integer = "-1", Optional ByVal psFecParaAdjudicar As String = "@", _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "UPDATE ColocPigRecGar SET "

    If psFecProceso <> "@" Then
        lsSQL = lsSQL & " dProceso = '" & psFecProceso & "',"
    End If
    If pnRGEstado <> -1 Then
        lsSQL = lsSQL & " nRGEstado = " & pnRGEstado & ","
    End If
    If psFecParaAdjudicar <> "@" Then
        lsSQL = lsSQL & " dParaAdjudicar = '" & psFecParaAdjudicar & "',"
    End If
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cTpoProceso ='" & psTpoProceso & "' AND cNroProceso = '" & psNroProceso & "' " _
                  & " AND cCodAge = '" & psProcesoCadaAgencia & "' "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Insert ColocPigRGPrecio
Public Sub dInsertColocPigRGPrecio(ByVal psTpoProceso As String, ByVal psNroProceso As String, _
        ByVal psTpoPrecio As String, ByVal pnPrecio As Currency, _
        ByVal psProcesoCadaAgencia As String, Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT ColocPigRGPrecio (cTpoProceso, cNroProceso, cTpoPrecio, nPrecio, cCodAge) " _
        & "VALUES ('" & psTpoProceso & "','" & psNroProceso & "','" & psTpoPrecio & "'," & pnPrecio & ",'" & psProcesoCadaAgencia & "' )"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

' Update ColocPigRGPrecio
Public Sub dUpdateColocPigRGPrecio(ByVal psTpoProceso As String, ByVal psNroProceso As String, _
        ByVal psTpoPrecio As String, Optional ByVal pnPrecio As Currency = -1, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "UPDATE ColocPigRecGar SET "

    If pnPrecio <> -1 Then
        lsSQL = lsSQL & " nPrecio = " & pnPrecio & ","
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cTpoProceso ='" & psTpoProceso _
        & "' AND cNroProceso ='" & psNroProceso & "' AND cTpoPrecio ='" & psTpoPrecio & "' "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Delete ColocPigRGPrecio
Public Sub dDeleteColocPigRGPrecio(ByVal psTpoProceso As String, ByVal psNroProceso As String, _
        ByVal psProcesoCadaAgencia As String, Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "DELETE ColocPigRGPrecio WHERE cTpoProceso ='" & psTpoProceso & "' " _
        & " AND cNroProceso ='" & psNroProceso & "' " & " AND cCodAge = '" & psProcesoCadaAgencia & "' "
 
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub


'Insert ColocPigRGDet
Public Sub dInsertColocPigRGDet(ByVal psTpoProceso As String, ByVal psNroProceso As String, _
        ByVal psctacod As String, ByVal pnRGDetEstado As Integer, ByVal pnMontoProceso As Currency, _
        ByVal pnRemSubBaseVta As Currency, ByVal pnAdjValMercado As Currency, _
        ByVal pnAdjValRegistro As Currency, ByVal pnDeuda As Currency, ByVal psProcesoCadaAgencia As String, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT ColocPigRGDet (cTpoProceso,cNroProceso,cCtaCod,nRGDetEstado,nMontoProceso,nRemSubBaseVta,nAdjValMercado," _
        & "nAdjValRegistro,nDeuda,cCodAge)  " _
        & "VALUES ('" & psTpoProceso & "','" & psNroProceso & "','" & psctacod & "'," & pnRGDetEstado & "," _
        & pnMontoProceso & "," & pnRemSubBaseVta & "," & pnAdjValMercado & "," & pnAdjValRegistro _
        & "," & pnDeuda & ",'" & psProcesoCadaAgencia & "')"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

' Update ColocPigRGPrecio
Public Sub dUpdateColocPigRGDet(ByVal psTpoProceso As String, ByVal psNroProceso As String, _
        ByVal psctacod As String, Optional ByVal pnRGDetEstado As Integer = -1, Optional ByVal pnMontoProceso As Currency = -1, _
        Optional ByVal pnRemSubBaseVta As Currency = -1, Optional ByVal pnAdjValMercado As Currency = -1, _
        Optional ByVal pnAdjValRegistro As Currency = -1, Optional ByVal pnMovVta As Long, Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "UPDATE ColocPigRGDet SET "

    If pnRGDetEstado <> -1 Then
        lsSQL = lsSQL & " nRGDetEstado = " & pnRGDetEstado & ","
    End If
    If pnMontoProceso <> -1 Then
        lsSQL = lsSQL & " nMontoProceso  = " & pnMontoProceso & ","
    End If
    If pnRemSubBaseVta <> -1 Then
        lsSQL = lsSQL & " nRemSubBaseVta = " & pnRemSubBaseVta & ","
    End If
    If pnAdjValMercado <> -1 Then
        lsSQL = lsSQL & " nAdjValMercado = " & pnAdjValMercado & ","
    End If
    If pnAdjValRegistro <> -1 Then
        lsSQL = lsSQL & " nAdjValRegistro = " & pnAdjValRegistro & ","
    End If
    If pnMovVta <> -1 Then
        lsSQL = lsSQL & " nMovVta = " & pnMovVta & ","
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cTpoProceso ='" & psTpoProceso & "' " _
        & "AND cNroProceso ='" & psNroProceso & "' AND cCtaCod ='" & psctacod & "' "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Insert ColocPigRecupVta
Public Sub dInsertColocPigRecupVta(ByVal psTpoProceso As String, ByVal psNroProceso As String, _
        ByVal psctacod As String, ByVal pnDocTpo As Integer, ByVal psNroDoc As String, ByVal psPersCod As String, _
        ByVal pnPrecioVta As Currency, ByVal pnComision As Currency, ByVal pnSobrante As Currency, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT ColocPigRecupVta (cTpoProceso, cNroProceso, cCtaCod, nDocTpo, cNroDoc, cPersCod, nPrecioVta, nComision, nSobrante) " _
        & "VALUES ('" & psTpoProceso & "','" & psNroProceso & "','" & psctacod & "'," & pnDocTpo & ",'" _
        & psNroDoc & "','" & psPersCod & "'," & pnPrecioVta & "," & pnComision & "," & pnSobrante & " ) "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Update ColocPigRecupVta
Public Sub dUpdateColocPigRecupVta(ByVal psTpoProceso As String, ByVal psNroProceso As String, _
        ByVal psctacod As String, Optional ByVal pnDocTpo As Integer = -1, Optional ByVal psNroDoc As String = "@", _
        Optional ByVal psPersCod As String = "@", _
        Optional ByVal pnPrecioVta As Currency = -1, Optional ByVal pnComision As Currency = -1, _
        Optional ByVal pnSobrante As Currency = -1, Optional ByVal psPagoSobrante As String = "@", _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "UPDATE ColocPigRecupVta SET "

    If pnDocTpo <> -1 Then
        lsSQL = lsSQL & " nDocTpo = " & pnDocTpo & ","
    End If
    If psNroDoc <> "@" Then
        lsSQL = lsSQL & " cNroDoc  = '" & psNroDoc & "',"
    End If
    If psPersCod <> "@" Then
        lsSQL = lsSQL & " cPersCod = '" & psPersCod & "',"
    End If
    If pnPrecioVta <> -1 Then
        lsSQL = lsSQL & " nPrecioVta = " & pnPrecioVta & ","
    End If
    If pnComision <> -1 Then
        lsSQL = lsSQL & " nComision = " & pnComision & ","
    End If
    If pnSobrante <> -1 Then
        lsSQL = lsSQL & " nSobrante = " & pnSobrante & ","
    End If
    If psPagoSobrante <> "@" Then
        lsSQL = lsSQL & " dPagoSobrante = '" & Format(psPagoSobrante, "yyyy/mm/dd") & "',"
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cTpoProceso ='" & psTpoProceso & "' " _
        & "AND cNroProceso ='" & psNroProceso & "' AND cCtaCod ='" & psctacod & "' "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub
'**Delete ColocPigRecupVta
Public Sub dDeleteColocPigRecupVta(ByVal psTpoProceso As String, ByVal psNroProceso As String, _
        ByVal psctacod As String, Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "DELETE ColocPigRecupVta WHERE cTpoProceso ='" & psTpoProceso & "' " _
        & "AND cNroProceso ='" & psNroProceso & "' AND cCtaCod ='" & psctacod & "' "
        
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'********************************************************************
'***  RECUPERACIONES  ***********************************************
'********************************************************************

'Insert ColocRecup
Public Sub dInsertColocRecup(ByVal psctacod As String, ByVal psFecIngRecup As String, _
        ByVal pnComisionCod As Integer, ByVal pnSaldoIntComp As Currency, ByVal pnSaldoIntMor As Currency, _
        ByVal pnSaldoSaldoGasto As Currency, ByVal pnIntCompGen As Currency, ByVal psMetLiquid As String, _
        ByVal pnTipCj As Integer, ByVal pnDemanda As Integer, ByVal pnNroCalen As Integer, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT ColocRecup (cCtaCod, dIngRecup, nComisionCod, nSaldoIntComp, nSaldoIntMor, " _
        & "nSaldoGasto, nIntCompGen,cMetLiquid,nTipCj,nDemanda,nNroCalen ) " _
        & "VALUES ('" & psctacod & "','" & psFecIngRecup & "'," & pnComisionCod & "," & pnSaldoIntComp & "," _
        & pnSaldoIntMor & "," & pnSaldoSaldoGasto & "," & pnIntCompGen & ",'" & psMetLiquid & "'," _
        & pnTipCj & "," & pnDemanda & "," & pnNroCalen & " )"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'**Update ColocRecup
Public Sub dUpdateColocRecup(ByVal psctacod As String, Optional ByVal psFecIngRecup As String = "@", _
        Optional ByVal pnComisionCod As Integer = "-1", Optional ByVal pnSaldoIntComp As Currency = -1, _
        Optional ByVal pnSaldoIntMor As Currency = -1, Optional ByVal pnSaldoGasto As Currency = -1, _
        Optional ByVal pnIntCompGen As Currency = -1, Optional ByVal psMetLiquid As String = "@", _
        Optional ByVal pnTipCj As Integer = -1, _
        Optional ByVal pnDemanda As Integer = -1, Optional ByVal pnNroCalend As Integer = -1, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "UPDATE ColocRecup SET "

    If psFecIngRecup <> "@" Then
        lsSQL = lsSQL & " dIngRecup = '" & psFecIngRecup & "',"
    End If
    If pnComisionCod <> -1 Then
        lsSQL = lsSQL & " nComisionCod = " & pnComisionCod & ","
    End If
    
    If pnSaldoIntComp = -4 Then ' Si se le envia (-4) disminuye la Cantidad
        lsSQL = lsSQL & " nSaldoIntComp  = nSaldoIntComp - " & pnSaldoIntComp & ","
    ElseIf pnSaldoIntComp = -3 Then ' Si se le envia (-3) aumenta la Cantidad
        lsSQL = lsSQL & " nSaldoIntComp  = nSaldoIntComp + " & pnSaldoIntComp & ","
    ElseIf pnSaldoIntComp <> -1 Then
        lsSQL = lsSQL & " nSaldoIntComp = " & pnSaldoIntComp & ","
    End If
    
    If pnSaldoIntMor = -4 Then ' Si se le envia (-4) disminuye la Cantidad
        lsSQL = lsSQL & " nSaldoIntMor = nSaldoIntMor - " & pnSaldoIntMor & ","
    ElseIf pnSaldoIntMor = -3 Then ' Si se le envia (-3) aumenta la Cantidad
        lsSQL = lsSQL & " nSaldoIntMor = nSaldoIntMor + " & pnSaldoIntMor & ","
    ElseIf pnSaldoIntMor <> -1 Then
        lsSQL = lsSQL & " nSaldoIntMor = " & pnSaldoIntMor & ","
    End If
    
    If pnSaldoGasto = -4 Then ' Si se le envia (-4) disminuye la cantidad
        lsSQL = lsSQL & " nSaldoGasto = nSaldoGasto - " & pnSaldoGasto & ","
    ElseIf pnSaldoGasto = -3 Then ' Si se le envia (-3) aumenta la cantidad
        lsSQL = lsSQL & " nSaldoGasto = nSaldoGasto + " & pnSaldoGasto & ","
    ElseIf pnSaldoGasto <> -1 Then
        lsSQL = lsSQL & " nSaldoGasto = " & pnSaldoGasto & ","
    End If
    
    If pnIntCompGen = -3 Then 'Si se le envia (-3) aumenta la cantidad
        lsSQL = lsSQL & " nIntCompGen = nIntCompGen + " & pnIntCompGen & ","
    ElseIf pnIntCompGen <> -1 Then
        lsSQL = lsSQL & " nIntCompGen  = " & pnIntCompGen & ","
    End If
    
    If psMetLiquid <> "@" Then
        lsSQL = lsSQL & " cMetLiquid = '" & psMetLiquid & "',"
    End If
    If pnTipCj <> -1 Then
        lsSQL = lsSQL & " nTipCJ = " & pnTipCj & ","
    End If
    If pnDemanda <> -1 Then
        lsSQL = lsSQL & " nDemanda = " & pnDemanda & ","
    End If
    If pnNroCalend <> -1 Then
        lsSQL = lsSQL & " nNroCalen = " & pnNroCalend & ","
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cCtaCod ='" & psctacod & "' "
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub


'Insert ColocRecupEstado
Public Sub dInsertColocRecupEstado(ByVal psctacod As String, ByVal psFecEstado As String, _
        ByVal pnEstado As ColocEstado, ByVal psFlag As String, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT ColocRecupEstado (cCtaCod, dEstado, nEstado, cFlag )" _
        & "VALUES ('" & psctacod & "','" & psFecEstado & "'," & pnEstado & ",'" & psFlag & "') "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Insert ColocRecComision
Public Sub dInsertColocRecComision(ByVal pnComisionCod As Integer, ByVal psCodAbogado As String, _
        ByVal pnTipComis As Integer, ByVal pnValor As Double, ByVal pnRangIni As Currency, ByVal pnRangFin As Currency, _
        ByVal pnCategoria As Integer, Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    
    lsSQL = "INSERT ColocRecupComision (nComisionCod, cPersCod, nTipComis, nValor, nRangIni, nRangFin, nCategoria) " _
        & " VALUES (" & pnComisionCod & ",'" & psCodAbogado & "'," & pnTipComis & "," & pnValor & "," _
        & pnRangIni & "," & pnRangFin & "," & pnCategoria & " )"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Update ColocRecComision
Public Sub dUpdateColocRecComision(ByVal pnComisionCod As Integer, _
        Optional ByVal pnTipComis As Integer = -1, Optional ByVal pnValor As Double = "-1", _
        Optional ByVal pnRangIni As Currency = -1, Optional ByVal pnRangFin As Currency = -1, _
        Optional ByVal pnCategoria As Integer = -1, Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "UPDATE ColocRecupComision SET "

    If pnTipComis <> -1 Then
        lsSQL = lsSQL & " nTipComis = " & pnTipComis & ","
    End If
    If pnValor <> -1 Then
        lsSQL = lsSQL & " nValor = " & pnValor & ","
    End If
    If pnRangIni <> -1 Then
        lsSQL = lsSQL & " nRangIni = " & pnRangIni & ","
    End If
    If pnRangFin <> -1 Then
        lsSQL = lsSQL & " nRangFin = " & pnRangFin & ","
    End If
    If pnCategoria <> -1 Then
        lsSQL = lsSQL & " nCategoria = " & pnCategoria & ","
    End If

    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE nComisionCod = " & pnComisionCod & " "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Insert ColocRecExpediente
Public Sub dInsertColocRecExpediente(ByVal psctacod As String, ByVal psNumExp As String, ByVal pnMonPetit As Currency, _
        ByVal pnMoneda As Integer, ByVal pnViaProce As Integer, ByVal pnEstadoProceso As Integer, _
        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    
    lsSQL = "INSERT ColocRecupExpediente (cCtaCod, cNumExp, nMonPetit, nMoneda,nViaProce, nEstadoProceso ) " _
        & " VALUES ('" & psctacod & "','" & Trim(psNumExp) & "'," & pnMonPetit & "," & pnMoneda & "," _
        & pnViaProce & "," & pnEstadoProceso & ") "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Update ColocRecExpediente
Public Sub dUpdateColocRecupExpediente(ByVal psctacod As String, Optional ByVal psNumExp As String = "@", _
        Optional ByVal pnMonPetit As Currency = -1, Optional ByVal pnMoneda As Integer = -1, _
        Optional ByVal pnViaProce As Integer = -1, Optional ByVal pnEstadoProceso As Integer = -1, _
        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "UPDATE ColocRecupExpediente SET "

    If psNumExp <> "@" Then
        lsSQL = lsSQL & " cNumExp = '" & psNumExp & "',"
    End If
    If pnMonPetit <> -1 Then
        lsSQL = lsSQL & " nMonPetit = " & pnMonPetit & ","
    End If
    If pnMoneda <> -1 Then
        lsSQL = lsSQL & " nMoneda = " & pnMoneda & ","
    End If
    If pnViaProce <> -1 Then
        lsSQL = lsSQL & " nViaProce = " & pnViaProce & ","
    End If
    'If pnEstadoProceso <> -1 Then
    '    lsSQL = lsSQL & " nEstadoViaProce = " & pnViaProce & ","
    'End If

    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cCtaCod = '" & psctacod & "' "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Insert ColocRecExpedienteInf
Public Sub dInsertColocRecExpedienteInf(ByVal psctacod As String, ByVal pnTipoInf As Integer, ByVal psInforme As String, _
        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    
    lsSQL = "INSERT ColocRecupExpedienteInf (cCtaCod, nTipoInf, mInforme) " _
        & " VALUES ('" & psctacod & "'," & pnTipoInf & ",'" & psInforme & "' ) "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Update ColocRecExpedienteInf
Public Sub dUpdateColocRecupExpedienteInf(ByVal psctacod As String, ByVal pnTipoInf As ColocRecExpedTipoInf, _
        Optional ByVal psInforme As String = "@", Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "DELETE ColocRecupExpedienteInf WHERE cCtaCod = '" & psctacod & "' And nTipoInf = " & pnTipoInf & " "
    lsSQL = lsSQL & "INSERT ColocRecupExpedienteInf (cCtaCod, nTipoInf, mInforme) VALUES ('" & psctacod & "'," & pnTipoInf & ",'" & psInforme & "' ) "

    'If psInforme <> "@" Then
    '    lsSQL = lsSQL & " mInforme = '" & psInforme & "',"
    'End If

    'lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    'lsSQL = lsSQL & " WHERE cCtaCod = '" & psCtaCod & "' And nTipoInf = " & pnTipoInf & " "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Insert ColocRecComisionHist
Public Sub dInsertColocRecComisionHist(ByVal psctacod As String, ByVal pnComisionCod As Integer, _
        ByVal pnMovNro As Long, Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    
    lsSQL = "INSERT ColocRecupComisionHist (cCtaCod, nComisionCod, nMovNro) " _
        & " VALUES ('" & psctacod & "'," & pnComisionCod & "," & pnMovNro & " ) "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Insert ColocRecGastos
Public Sub dInsertColocRecupGastos(ByVal psctacod As String, ByVal pnNroGastoCta As Integer, _
        ByVal pnPrdConceptoCod As Integer, _
        ByVal psFecAsigna As String, ByVal pnMonto As Currency, ByVal pnMontoPagado As Currency, _
        ByVal pnColocRecGastoEstado As ColocRecGastoEstado, ByVal psMotivoGasto As String, _
        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    
    lsSQL = "INSERT ColocRecupGastos (cCtaCod,nNroGastoCta, nPrdConceptoCod, dAsigna,nMonto, " _
        & " nMontoPagado, nColocRecGastoEstado,cMotivoGasto ) " _
        & " VALUES ('" & psctacod & "'," & pnNroGastoCta & "," & pnPrdConceptoCod & ",'" & psFecAsigna & "'," & pnMonto & "," _
        & pnMontoPagado & "," & pnColocRecGastoEstado & ",'" & psMotivoGasto & "'  )"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Update ColocRecGastos
Public Sub dUpdateColocRecupGastos(ByVal psctacod As String, ByVal pnNroGastoCta As Integer, _
        Optional ByVal psFecAsigna As String = "@", Optional ByVal pnMonto As Currency = -1, Optional ByVal pnMontoPagado As Currency = -1, _
        Optional ByVal pnColocRecGastoEstado As ColocRecGastoEstado = gColRecGastoEstPendiente, Optional ByVal psMotivoGasto As String = "@", _
        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    
    lsSQL = "UPDATE ColocRecupGastos SET "

    If psFecAsigna <> "@" Then
        lsSQL = lsSQL & " dAsigna = '" & psFecAsigna & "',"
    End If
    If pnMonto <> -1 Then
        lsSQL = lsSQL & " nMonto = " & pnMonto & ","
    End If
    If pnMontoPagado <> -1 Then
        lsSQL = lsSQL & " nMontoPagado = " & pnMontoPagado & ","
    End If
    If pnColocRecGastoEstado <> -1 Then
        lsSQL = lsSQL & " nColocRecGastoEstado = " & pnColocRecGastoEstado & ","
    End If
    If psMotivoGasto <> "@" Then
        lsSQL = lsSQL & " cMotivoGasto = '" & psMotivoGasto & "',"
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cCtaCod = '" & psctacod & "' AND  nNroGastoCta = " & pnNroGastoCta & " "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub


'Delete ColocRecExpedienteInf
Public Sub dDeleteColocRecupExpedienteInf(ByVal psctacod As String, Optional ByVal pnTipoInf As ColocRecExpedTipoInf = -1, _
    Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "DELETE ColocRecupExpedienteInf WHERE cCtaCod = '" & psctacod & "'  "
    If pnTipoInf <> -1 Then
        lsSQL = lsSQL & " And nTipoInf = " & pnTipoInf & " "
    End If

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Delete ColocRecExpediente
Public Sub dDeleteColocRecupExpediente(ByVal psctacod As String, Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "DELETE ColocRecupExpediente WHERE cCtaCod = '" & psctacod & "' "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Delete ColocRecGastos
Public Sub dDeleteColocRecupGastos(ByVal psctacod As String, Optional ByVal pnNroGastoCta As Integer = -1, _
        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    
    lsSQL = "DELETE ColocRecupGastos WHERE cCtaCod = '" & psctacod & "' "
    If pnNroGastoCta <> -1 Then
        lsSQL = lsSQL & " AND  nNroGastoCta = " & pnNroGastoCta & " "
    End If

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'**Delete ColocRecup
Public Sub dDeleteColocRecup(ByVal psctacod As String, Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "DELETE ColocRecup WHERE  cCtaCod ='" & psctacod & "' "
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dInsertCredPignoraticioDetalle(ByVal psCuenta As String, _
        ByVal pnTasaInteres As Double, ByVal pnSaldo As Currency, _
        ByVal psFechaHora As String, ByVal pnPlazo As Integer, ByVal psFecVenc As String, _
        ByVal pnOroBruto As Double, ByVal pnOroNeto As Double, ByVal pnTasacion As Currency, _
        ByVal pnPiezas As Integer, ByVal psTipoContrato As String, ByVal psLote As String, _
        ByVal pnki14 As Double, ByVal pnki16 As Double, ByVal pnKi18 As Double, _
        ByVal pnKi21 As Double, ByVal psMovNro As String, ByVal prJoyas As ADODB.Recordset, _
        Optional ByVal pbEjecBatch As Boolean = False)

Dim loParam As DColPCalculos, ntasamor As Double, ntasagra As Double
Set loParam = New DColPCalculos

    ntasamor = loParam.dObtieneTasaInteres("01011130501", "03")
    ntasagra = loParam.dObtieneTasaInteres("01011130501", "05")
    
Set loParam = Nothing

Dim lsSQL As String

'**Insert Producto
lsSQL = "INSERT Producto (cCtaCod,nTasaInteres,nSaldo,nPrdEstado,dPrdEstado,nTransacc) " _
    & "VALUES ('" & psCuenta & "'," & pnTasaInteres & "," & pnSaldo & "," & gColPEstRegis & ",'" & psFechaHora & "',1)"
If pbEjecBatch Then
    coConex.AdicionaCmdBatch lsSQL
Else
    coConex.Ejecutar lsSQL
End If
'2004/10/14 MPBR y LAYG
'***
lsSQL = "INSERT INTO PRODUCTOTASAINTERES(cCtaCod,nTasaInteres,nPrdTasaInteres) " _
      & "VALUES('" & psCuenta & "'," & pnTasaInteres & ",1)"
If pbEjecBatch Then
    coConex.AdicionaCmdBatch lsSQL
Else
    coConex.Ejecutar lsSQL
End If

lsSQL = "INSERT INTO PRODUCTOTASAINTERES(cCtaCod,nTasaInteres,nPrdTasaInteres) " _
      & "VALUES('" & psCuenta & "'," & ntasagra & ",5)"
If pbEjecBatch Then
    coConex.AdicionaCmdBatch lsSQL
Else
    coConex.Ejecutar lsSQL
End If

lsSQL = "INSERT INTO PRODUCTOTASAINTERES(cCtaCod,nTasaInteres ,nPrdTasaInteres) " _
      & "VALUES('" & psCuenta & "'," & ntasamor & ",3)"
If pbEjecBatch Then
    coConex.AdicionaCmdBatch lsSQL
Else
    coConex.Ejecutar lsSQL
End If






'**Insert Colocaciones
lsSQL = "INSERT Colocaciones (cCtaCod,nPlazo,dVenc,nMontoCol,nColocCalendCod,cUltimaActualizacion,cLineaCred,dVigencia) " _
    & "VALUES ('" & psCuenta & "'," & pnPlazo & ",'" & psFecVenc & "'," & pnSaldo & "," & gColocCalendCodPFCF & ",'" & psMovNro & "','" & "01011130501" & "','" & psFechaHora & "' )"
If pbEjecBatch Then
    coConex.AdicionaCmdBatch lsSQL
Else
    coConex.Ejecutar lsSQL
End If
'**Insert ColocPignoraticio
lsSQL = "INSERT ColocPignoraticio (cCtaCod,nOroBruto,nOroNeto,nPiezas,nTasacion,nNroRenov,nNroRemate,nNroDuplic,cPrdCtaTpo,cAgeBoveda,cLote) " _
     & "VALUES ('" & psCuenta & "'," & pnOroBruto & "," & pnOroNeto & "," & pnPiezas & "," & pnTasacion & ",0,0,0,'" & psTipoContrato & "','" & Mid(psCuenta, 1, 5) & "','" & psLote & "') "
If pbEjecBatch Then
    coConex.AdicionaCmdBatch lsSQL
Else
    coConex.Ejecutar lsSQL
End If
'**Insert ColocPigJoya
If pnki14 > 0 Then
    lsSQL = "INSERT ColocPigJoya (cCtaCod,cKilataje,nPesoOro) " _
          & "VALUES ('" & psCuenta & "','14'," & pnki14 & " )"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End If
If pnki16 > 0 Then
    lsSQL = "INSERT ColocPigJoya (cCtaCod,cKilataje,nPesoOro) " _
          & "VALUES ('" & psCuenta & "','16'," & pnki16 & " )"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End If
If pnKi18 > 0 Then
    lsSQL = "INSERT ColocPigJoya (cCtaCod,cKilataje,nPesoOro) " _
          & "VALUES ('" & psCuenta & "','18'," & pnKi18 & " )"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End If
If pnKi21 > 0 Then
    lsSQL = "INSERT ColocPigJoya (cCtaCod,cKilataje,nPesoOro) " _
          & "VALUES ('" & psCuenta & "','21'," & pnKi21 & " )"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End If
'inserta Joyas para cada registro de prJoyas
Do While Not prJoyas.EOF
    lsSQL = "INSERT ColocPigJoyaDet (cCtaCod,cKilataje, nItem, nPiezas, " _
          & "nPesoBruto, nPesoNeto, nValTasac, cDescrip ) " _
          & "VALUES ('" & psCuenta & "','" & Trim(Str(Right(prJoyas("Material"), 2))) & "'," _
          & prJoyas.AbsolutePosition & "," & prJoyas("Pzas") & "," _
          & prJoyas("PBruto") & "," & prJoyas("PNeto") & "," _
          & prJoyas("Tasac") & ",'" & prJoyas("Descripcion") & "' )"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    prJoyas.MoveNext
Loop
End Sub




'Insert ColocRecComisionHist
Public Sub dInsertColocRecupActProcesales(ByVal psctacod As String, ByVal psMovNro As String, _
        ByVal psComenta As String, Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    
    lsSQL = "INSERT ColocRecupActProcesales (cCtaCod, cMovNro, cComenta) " _
        & " VALUES ('" & psctacod & "','" & psMovNro & "','" & psComenta & "' ) "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub



Public Sub dInsertHistColocPigJoyaDetRet(ByVal prJoyas As ADODB.Recordset, ByVal pnNumRetasacion As Integer, ByVal psMovNro As String, ByVal psctacod As String)
'inserta Joyas para cada registro de prJoyas
Dim lsSQL As String
Do While Not prJoyas.EOF
    lsSQL = "INSERT HistColocPigJoyaDetRet (cCtaCod,cKilataje, nItem, nPiezas, " _
          & "nPesoBruto, nPesoNeto, nValTasac, cDescrip,nnumretasacion,cmovnro ) " _
          & "VALUES ('" & psctacod & "','" & Trim(Str(Left(prJoyas.Fields("ckilataje"), 2))) & "'," _
          & prJoyas.Fields("nitem") & "," & prJoyas.Fields("npiezas") & "," _
          & prJoyas.Fields("nPesoBruto") & "," & prJoyas.Fields("nPesoNeto") & "," _
          & prJoyas.Fields("nValTasac") & ",'" & prJoyas.Fields("cDescrip") & "'," & pnNumRetasacion & ",'" & psMovNro & "')"
    
        coConex.Ejecutar lsSQL
    
    prJoyas.MoveNext
Loop
End Sub

Public Sub dUpdateColocPigJoyaDet(ByVal prJoyas As ADODB.Recordset, ByVal psctacod As String)
'inserta Joyas para cada registro de prJoyas
Dim lsSQL As String
Do While Not prJoyas.EOF
    lsSQL = " Update ColocPigJoyaDet "
    lsSQL = lsSQL & " Set cKilataje='" & prJoyas.Fields("Material") & "', nPesoBruto= " & prJoyas.Fields("pBruto") & ",nPesoNeto=" & prJoyas.Fields("pNeto")
    lsSQL = lsSQL & " where cctacod='" & psctacod & "' and nitem=" & prJoyas.AbsolutePosition
    coConex.Ejecutar lsSQL
    
    prJoyas.MoveNext
Loop

End Sub


Public Sub dUpdateColocPigJoya(ByVal psctacod As String, Optional n14k As Double = 0, Optional n16k As Double = 0, Optional n18k As Double = 0, Optional n21k As Double = 0)
'inserta Joyas para cada registro de prJoyas
    Dim lsSQL As String
    lsSQL = "Delete ColocPigJoya where cctacod='" & psctacod & "'"
    coConex.Ejecutar lsSQL

If n14k > 0 Then
    lsSQL = "Insert into ColocPigJoya "
    lsSQL = lsSQL & " values('" & psctacod & "','14'," & n14k & " )"
    coConex.Ejecutar lsSQL
End If
    
If n16k > 0 Then
    lsSQL = "insert into ColocPigJoya "
    lsSQL = lsSQL & " values('" & psctacod & "','16'," & n16k & " )"
    coConex.Ejecutar lsSQL
End If
    
If n18k > 0 Then
    lsSQL = "insert into ColocPigJoya "
    lsSQL = lsSQL & " values('" & psctacod & "','18'," & n18k & " )"
    coConex.Ejecutar lsSQL
End If
    
If n21k > 0 Then
    lsSQL = "insert into ColocPigJoya "
    lsSQL = lsSQL & " values('" & psctacod & "','21'," & n21k & " )"
    coConex.Ejecutar lsSQL
End If
    
    

End Sub

