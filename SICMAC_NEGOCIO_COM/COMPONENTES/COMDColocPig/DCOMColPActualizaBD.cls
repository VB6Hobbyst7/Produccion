VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMColPActualizaBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'* Modulo de Colocaciones Pignoraticio
'* Clase con todas las Actualizaciones a BD
Option Explicit
Dim coConex As New COMConecta.DCOMConecta
Dim oError As New COMConecta.COMErrorHandling
Dim csConexion As String
Dim csNegocio As String
Dim csCentralPer As String
Dim csCentralCom As String
Dim csCentralImg As String
Dim csAdminist As String
Private cnConecOpen As Integer


Public Property Get cnConecOpenValor() As Variant
    cnConecOpenValor = cnConecOpen
End Property

Public Property Let cnConecOpenValor(ByVal vNewValue As Variant)
    cnConecOpen = vNewValue
End Property

Public Function dEjecutar(psSql As String) As ADODB.Recordset
On Error GoTo ErrEjecutar
Set dEjecutar = coConex.Ejecutar(psSql)
Exit Function
ErrEjecutar:
   oError.RaiseError oError.MyUnhandledError, "COMConecta.DCOMConecta:Ejecutar Method"
End Function

'Ejecuta procesos Batch
Public Function dEjecutaBatch() As Integer
On Error GoTo EjecutaBatchErr
    coConex.EjecutarBatch
    Exit Function
EjecutaBatchErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DColPActualizaBD - Ejecutar Batch")
End Function

Public Sub dInsertCredPignoraticio(ByVal psCuenta As String, _
        ByVal pnTasaInteres As Double, ByVal pnSaldo As Currency, _
        ByVal psFechaHora As String, ByVal pnPlazo As Integer, ByVal psFecVenc As String, _
        ByVal pnOroBruto As Double, ByVal pnOroNeto As Double, ByVal pnTasacion As Currency, _
        ByVal pnPiezas As Integer, ByVal psTipoContrato As String, ByVal psLote As String, _
        ByVal pnki14 As Double, ByVal pnki16 As Double, ByVal pnKi18 As Double, _
        ByVal pnKi21 As Double, ByVal psMovNro As String, _
        Optional ByVal pbEjecBatch As Boolean = False)

Dim gColocLineaCredPig As String
Dim lsSQL As String
Dim psColocLineaCredPig As String

'psColocLineaCredPig = "0101113050101"
'BAS II
psColocLineaCredPig = "0101117550101"

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
    & "VALUES ('" & psCuenta & "'," & pnPlazo & ",'" & psFecVenc & "'," & pnSaldo & "," & gColocCalendCodPFCF & ",'" & psMovNro & "','" & psColocLineaCredPig & "','" & psFechaHora & "' )"
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

Public Sub dInsertProducto(ByVal psCtaCod As String, ByVal pnTasaInteres As ColocTipoTasa, _
        ByVal pnSaldo As Currency, ByVal pnEstado As ColocEstado, ByVal psFecEstado As String, _
        ByVal pnTransac As Integer, _
        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT Producto (cCtaCod, nTasaInteres, nSaldo, nPrdEstado, dPrdEstado, nTransacc ) " _
        & "VALUES ('" & psCtaCod & "'," & pnTasaInteres & "," & pnSaldo & "," & pnEstado & ",'" & psFecEstado & "'," & pnTransac & ")"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dInsertColocaciones(ByVal psCtaCod As String, ByVal pnPlazo As Integer, _
        ByVal psFecVenc As String, ByVal pnMontoCol As Currency, ByVal pnTipoCalend As ColocTipoCalend, _
        ByVal psMovNro As String, ByVal psLineaCred As String, ByVal psFecVigencia As String, _
        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "INSERT Colocaciones (cCtaCod,nPlazo,dVenc,nMontoCol,nColocCalendCod,cUltimaActualizacion,cLineaCred, dVigencia) " _
        & "VALUES ('" & psCtaCod & "'," & pnPlazo & ",'" & psFecVenc & "'," & pnMontoCol & "," _
        & pnTipoCalend & ",'" & psMovNro & "','" & psLineaCred & "',''  )"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

Public Sub dInsertColocPignoraticio(ByVal psCtaCod As String, ByVal pnOroBruto As Currency, _
        ByVal pnOroNeto As Currency, ByVal pnPiezas As Integer, ByVal pnTasacion As Currency, _
        ByVal pnNroRenovac As Integer, ByVal pnNroDuplic As Integer, ByVal pnTipoCalend As ColocTipoCalend, _
        ByVal psAgeBoveda As String, ByVal psLote As String, Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "INSERT ColocPignoraticio (cCtaCod,nOroBruto,nOroNeto,nPiezas,nTasacion, " _
         & "nNroRenov,nNroRemate,nNroDuplic,cPrdCtaTpo,cAgeBoveda,cLote) " _
         & "VALUES ('" & psCtaCod & "'," & pnOroBruto & "," & pnOroNeto & "," & pnPiezas & "," & pnTasacion _
         & "," & pnNroRenovac & "," & pnNroDuplic & "," & pnTipoCalend & ",'" & psAgeBoveda & "','" & Replace(psLote, "'", "''") & "' )"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

Public Sub dInsertColocPigJoya(ByVal psCtaCod As String, ByVal psKilataje As String, _
        ByVal pnPesoOro As Currency, Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "INSERT ColocPigJoya (cCtaCod,cKilataje,nPesoOro )" _
         & "VALUES ('" & psCtaCod & "','" & psKilataje & "'," & pnPesoOro & " )"
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

'*** PEAC 20161111
Public Sub dInsertColocPigPagosParciales( _
    ByVal psCtaCod As String, ByVal pcFecVenc As String, ByVal pnPrdConceptoCod As Integer, _
    ByVal pnMonto As Currency, ByVal pnMontoPndte As Currency, ByVal pnMovNro As Long, _
    ByVal pnEstado As Integer, Optional ByVal pbEjecBatch As Boolean = False)

    Dim lsSQL As String

    lsSQL = "exec stp_ins_ColocPigPagosParciales '" & psCtaCod & "','" & pcFecVenc & "'," & pnPrdConceptoCod & "," & pnMonto & "," & pnMontoPndte & "," & pnEstado & "," & pnMovNro

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'*** PEAC 20161116
Public Sub dUpdateColocPigPagosParciales( _
    ByVal psCtaCod As String, ByVal pnMovNroCanc As Long, Optional ByVal pbEjecBatch As Boolean = False)

    Dim lsSQL As String

    lsSQL = "exec stp_upd_ColocPigPagosParciales '" & psCtaCod & "'," & pnMovNroCanc

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'*** PEAC 20161116
Public Sub dUpdateColocPigPagosParcialesExtorno( _
    ByVal psCtaCod As String, ByVal pnMovNroExt As Long, Optional ByVal pbEjecBatch As Boolean = False)

    Dim lsSQL As String

    lsSQL = "exec stp_upd_ColocPigPagosParcialesExtorno '" & psCtaCod & "'," & pnMovNroExt

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'********* GEMO 20210723
Public Sub dDeleteColocPigPagosParcialesExtorno( _
    ByVal psCtaCod As String, ByVal pnMovNroAnt As Long, ByVal pnMovNroExt As Long, Optional ByVal pbEjecBatch As Boolean = False)

    Dim lsSQL As String

    lsSQL = "exec stp_var_ColocPigPagosParcialesExtorno '" & psCtaCod & "'," & pnMovNroAnt & "," & pnMovNroExt
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
Public Sub dInsertColocacEstado(ByVal psCtaCod As String, ByVal psFechaPrdEstado As String, _
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
        & "VALUES ('" & psCtaCod & "','" & psFechaPrdEstado & "'," & psPrdEstado & "," & pnCuotas & "," _
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
Public Sub dInsertColocCalendario(ByVal psCtaCod As String, ByVal pnNroCalend As Integer, ByVal pnColocCalendApl As ColocCalendApl, _
        ByVal pnCuota As Integer, ByVal psFecVenc As String, ByVal pnColocCalendEstado As Integer, ByVal psDescrip As String, _
        ByVal psColocCalenFlag As String, Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "INSERT ColocCalendario (cCtaCod,nNroCalen,nColocCalendApl,nCuota,dVenc,nColocCalendEstado,cDescripcion,cColocCalenFlag) " _
         & "VALUES ('" & psCtaCod & "'," & pnNroCalend & "," & pnColocCalendApl & "," & pnCuota & ",'" & psFecVenc & "'," & pnColocCalendEstado & ",'" _
         & psDescrip & "','" & psColocCalenFlag & "' )"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

'Insert ColocCalendDet
Public Sub dInsertColocCalendDet(ByVal psCtaCod As String, ByVal pnNroCalend As Integer, ByVal pnColocCalendApl As ColocCalendApl, _
        ByVal pnCuota As Integer, ByVal pnColocConceptoCod As ColocConcepto, ByVal pnMonto As Currency, ByVal pnMontoPagado As Currency, _
        ByVal psFlag As String, Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "INSERT ColocCalendDet (cCtaCod,nNroCalen,nColocCalendApl,nCuota,nPrdConceptoCod,nMonto,nMontoPagado,cFlag) " _
         & "VALUES ('" & psCtaCod & "'," & pnNroCalend & "," & pnColocCalendApl & "," & pnCuota & "," & pnColocConceptoCod & "," & pnMonto & "," _
         & pnMontoPagado & ",'" & psFlag & "' ) "
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

'Insert ColocCalendDetPig
Public Sub dInsertColocCalendDetPig(ByVal psCtaCod As String, ByVal pnNroCalend As Integer, ByVal pnColocCalendApl As ColocCalendApl, _
        ByVal pnCuota As Integer, ByVal pnColocConceptoCod As ColocConcepto, ByVal pnMonto As Currency, ByVal pnMontoPagado As Currency, _
        ByVal psFlag As String, Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "INSERT ColocCalendDetPig (cCtaCod,nNroCalen,nColocCalendApl,nCuota,nPrdConceptoCod,nMonto,nMontoPagado,cFlag) " _
         & "VALUES ('" & psCtaCod & "'," & pnNroCalend & "," & pnColocCalendApl & "," & pnCuota & "," & pnColocConceptoCod & "," & pnMonto & "," _
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
    Call oError.RaiseError(oError.MyUnhandledError, "DColPActualizaBD Insert Mov")
End Sub

'Inserta el movimiento de lavado de dinero
Public Sub dInsertaMovLavDinero(ByVal nMovNro As Long, ByVal sPersCod As String, _
                                Optional psPersCodTit As String, _
                                Optional psPersCodOrd As String, _
                                Optional psPersCodRea As String, _
                                Optional psPersCodBen As String, _
                                Optional psPersCodVis As String, _
                                Optional nTipoREU As Integer = 1, Optional nMontoAcu As Double = 0, Optional sOrigen As String = "")
                                'By Capi 18022008 se agrego 4 parametros
                                'ALPA 20081009 se agrego 3 parametros
Dim lsSQL As String

On Error GoTo InsertMovLav
lsSQL = "INSERT MovLavDinero (nMovNro,cPersCod,cPersCodOrd,cPersCodRea,cPersCodBen,cPersCodVisto,nTipoReu,nMontoAcum,cOrigenEfectivo) " _
    & "VALUES (" & nMovNro & ",'" & psPersCodTit & "','" & psPersCodOrd & "','" & psPersCodRea & "','" & psPersCodBen & "','" & psPersCodVis & "'," & nTipoREU & ", " & nMontoAcu & ",'" & sOrigen & "')"
coConex.Ejecutar lsSQL
Exit Sub

InsertMovLav:
    Call oError.RaiseError(oError.MyUnhandledError, "DColPActualizaBD Insert Mov Lab")

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

'ALPA20130827*********
Dim oRS As ADODB.Recordset
Set oRS = New ADODB.Recordset
lsSQL = "Select * From MovRef where nMovNro=" & pnMovNro & " and nMovNroRef=" & pnMovNroRef & ""
Set oRS = coConex.CargaRecordSet(lsSQL)
If Not (oRS.BOF Or oRS.EOF) Then
    Exit Sub
End If
'*********************

        lsSQL = "INSERT MovRef (nMovNro,nMovNroRef) " _
              & "VALUES (" & pnMovNro & "," & pnMovNroRef & ")"
        
        If pbEjecBatch Then
            coConex.AdicionaCmdBatch lsSQL
        Else
            coConex.Ejecutar lsSQL
        End If

    Exit Sub

InsertMovRefErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DColPActualizaBD Insert Mov Ref")
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
        ByVal psCtaCod As String, pnNroCalend As Integer, ByVal pnConcepto As ColocConcepto, _
        ByVal pnNroCuota As Integer, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT MovColDet (nMovNro,cOpeCod,cCtaCod,nNroCalen,nPrdConceptoCod,nNroCuota,nMonto) " _
        & "VALUES (" & pnMovNro & ",'" & psOperacion & "','" & psCtaCod & "'," & pnNroCalend & "," _
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


Public Sub dInsertProductoBloqueos(ByVal psCtaCod As String, ByVal pnBlqTpo As Integer, ByVal pnBlqMotivo As String, _
        ByVal psMovNro As String, ByVal psComentario As String, Optional ByVal psMovNroDbl As String = "@", _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT ProductoBloqueos (cCtaCod,nBlqTpo,nBlqMotivo,cMovNro,cComentario) " _
        & "VALUES ('" & psCtaCod & "'," & pnBlqTpo & "," & pnBlqMotivo & ",'" & psMovNro & "','" _
        & psComentario & "') "  ' --' ," & IIf(psMovNroDbl = "@", Null, ",'" & psMovNroDbl & "' ") & " ) "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dUpdateProductoBloqueos(ByVal psCtaCod As String, ByVal psMovNro As String, _
        Optional ByVal psMovNroDbl As String = "@", Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "UPDATE ProductoBloqueos SET "
    
    If psMovNroDbl <> "@" Then
        lsSQL = lsSQL & " cMovNroDbl = '" & psMovNroDbl & "',"
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cCtaCod ='" & psCtaCod & "' AND cMovNro ='" & psMovNro & "' "
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub


'**Update Producto
Public Sub dUpdateProducto(ByVal psCtaCod As String, Optional ByVal pnTasaInteres As Double = -1, _
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
    lsSQL = lsSQL & " WHERE cCtaCod ='" & psCtaCod & "' "
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Colocaciones
Public Sub dUpdateColocaciones(ByVal psCtaCod As String, Optional ByVal pnPlazo As Integer = -1, _
        Optional ByVal psFecVenc As String = "@", Optional ByVal pnMontoCol As Currency = -1, Optional ByVal pnColocCalendCod As Integer = -1, _
        Optional ByVal psUltimaActualizacion As String = "@", Optional ByVal psLineaCred As String = "@", _
        Optional pbEjecBatch As Boolean = False, Optional psFechaVencAnt As String = "@", Optional pnPigExt As Integer = 0)

'*** PEAC 20161213 - se agrego psFechaVencAnt,pnPigExt

Dim lsSQL As String
    lsSQL = "UPDATE Colocaciones SET "

    If pnPlazo <> -1 Then
        lsSQL = lsSQL & " nPlazo = " & pnPlazo & ","
    End If
    If psFecVenc <> "@" Then
        '*** PEAC 20161213
        'lsSQL = lsSQL & " dVenc = '" & psFecVenc & "',"
        If pnPigExt = 1 Then
            lsSQL = lsSQL & " dVenc = dVencUlt, "
        Else
            lsSQL = lsSQL & " dVenc = '" & psFecVenc & "',"
        End If
        '*** FIN PEAC
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
    '*** PEAC 20161213
    If psFechaVencAnt <> "@" Then
        lsSQL = lsSQL & " dVencUlt = '" & psFechaVencAnt & "',"
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cCtaCod ='" & psCtaCod & "' "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

' Update ColocPignoraticio
Public Sub dUpdateColocPignoraticio(ByVal psCtaCod As String, _
        Optional ByVal pnOroBruto As Double = -1, Optional ByVal pnOroNeto As Double = -1, _
        Optional ByVal pnPiezas As Integer = -1, Optional ByVal pnTasacion As Currency = -1, _
        Optional ByVal pnNroRenov As Integer = -1, Optional ByVal pnNroRemate As Integer = -1, _
        Optional ByVal pnNroDuplic As Integer = -1, Optional ByVal psPrdCtaTpo As String = "@", _
        Optional ByVal psAgeBoveda As String = "@", Optional ByVal psLote As String = "@", _
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
    
'    If psCredAntiguo <> "@" Then
        lsSQL = lsSQL & " cCredB = 'B'," ' peac 20070925
'    End If
    
'    lsSQL = lsSQL & " ncodnotifiadj = 0," ' peac 20080515
    
    If psLote <> "@" Then
        lsSQL = lsSQL & " cLote = '" & psLote & "',"
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cCtaCod ='" & psCtaCod & "' "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'*** PEAC 20080412
Public Sub PignoExcluyeAdjudica(ByVal pcCta As String, ByVal psNumProce As String)

    Dim lsSQL As String
    lsSQL = "exec stp_upd_ExcluyeAdjudica '" & pcCta & "','" & psNumProce & "'"
    coConex.Ejecutar lsSQL

End Sub

'*** PEAC 20080412
Public Sub dActualizaPignoAdjudica(ByVal pdFecCorte As Date, ByVal psAge As String, _
    ByVal pdFechaLey As Date, Optional pdFecNotifi As Date, Optional pnNotifica As Integer, _
    Optional pnEstado As Integer, Optional psMovNroNotifi As String, Optional psNumProceso As String)

    Dim lsSQL As String
    lsSQL = "exec stp_upd_NotificaAdjudica '" & Format(pdFecCorte, "yyyymmdd") & "','" & psAge & "','" & Format(pdFechaLey, "yyyymmdd") & "','" & Format(pdFecNotifi, "yyyymmdd") & "'," & pnNotifica & "," & pnEstado & ",'" & psMovNroNotifi & "','" & psNumProceso & "'"
    coConex.Ejecutar lsSQL

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


Public Function dUltimaTransaccionProducto(ByVal psCtaCod As String) As Integer
Dim lsSQL As String
Dim lrs As ADODB.Recordset

On Error GoTo dError

lsSQL = "SELECT IsNull(nTransacc, 0 ) as nTransacc FROM Producto where cCtaCod ='" & psCtaCod & "' "

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
    lsSQL = "UPDATE ColocPigRGPrecio SET "

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
        ByVal psCtaCod As String, ByVal pnRGDetEstado As Integer, ByVal pnMontoProceso As Currency, _
        ByVal pnRemSubBaseVta As Currency, ByVal pnAdjValMercado As Currency, _
        ByVal pnAdjValRegistro As Currency, ByVal pnDeuda As Currency, ByVal psProcesoCadaAgencia As String, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT ColocPigRGDet (cTpoProceso,cNroProceso,cCtaCod,nRGDetEstado,nMontoProceso,nRemSubBaseVta,nAdjValMercado," _
        & "nAdjValRegistro,nDeuda,cCodAge)  " _
        & "VALUES ('" & psTpoProceso & "','" & psNroProceso & "','" & psCtaCod & "'," & pnRGDetEstado & "," _
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
        ByVal psCtaCod As String, Optional ByVal pnRGDetEstado As Integer = -1, Optional ByVal pnMontoProceso As Currency = -1, _
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
        & "AND cNroProceso ='" & psNroProceso & "' AND cCtaCod ='" & psCtaCod & "' "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Insert ColocPigRecupVta
Public Sub dInsertColocPigRecupVta(ByVal psTpoProceso As String, ByVal psNroProceso As String, _
        ByVal psCtaCod As String, ByVal pnDocTpo As Integer, ByVal psNroDoc As String, ByVal psPersCod As String, _
        ByVal pnPrecioVta As Currency, ByVal pnComision As Currency, ByVal pnSobrante As Currency, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT ColocPigRecupVta (cTpoProceso, cNroProceso, cCtaCod, nDocTpo, cNroDoc, cPersCod, nPrecioVta, nComision, nSobrante) " _
        & "VALUES ('" & psTpoProceso & "','" & psNroProceso & "','" & psCtaCod & "'," & pnDocTpo & ",'" _
        & psNroDoc & "','" & psPersCod & "'," & pnPrecioVta & "," & pnComision & "," & pnSobrante & " ) "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Update ColocPigRecupVta
Public Sub dUpdateColocPigRecupVta(ByVal psTpoProceso As String, ByVal psNroProceso As String, _
        ByVal psCtaCod As String, Optional ByVal pnDocTpo As Integer = -1, Optional ByVal psNroDoc As String = "@", _
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
        & "AND cNroProceso ='" & psNroProceso & "' AND cCtaCod ='" & psCtaCod & "' "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub
'**Delete ColocPigRecupVta
Public Sub dDeleteColocPigRecupVta(ByVal psTpoProceso As String, ByVal psNroProceso As String, _
        ByVal psCtaCod As String, Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "DELETE ColocPigRecupVta WHERE cTpoProceso ='" & psTpoProceso & "' " _
        & "AND cNroProceso ='" & psNroProceso & "' AND cCtaCod ='" & psCtaCod & "' "
        
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
Public Sub dInsertColocRecup(ByVal psCtaCod As String, ByVal psFecIngRecup As String, _
        ByVal pnComisionCod As Integer, ByVal pnSaldoIntComp As Currency, ByVal pnSaldoIntMor As Currency, _
        ByVal pnSaldoSaldoGasto As Currency, ByVal pnIntCompGen As Currency, ByVal psMetLiquid As String, _
        ByVal pnTipCj As Integer, ByVal pnDemanda As Integer, ByVal pnNroCalen As Integer, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT ColocRecup (cCtaCod, dIngRecup, nComisionCod, nSaldoIntComp, nSaldoIntMor, " _
        & "nSaldoGasto, nIntCompGen,cMetLiquid,nTipCj,nDemanda,nNroCalen ) " _
        & "VALUES ('" & psCtaCod & "','" & psFecIngRecup & "'," & pnComisionCod & "," & pnSaldoIntComp & "," _
        & pnSaldoIntMor & "," & pnSaldoSaldoGasto & "," & pnIntCompGen & ",'" & psMetLiquid & "'," _
        & pnTipCj & "," & pnDemanda & "," & pnNroCalen & " )"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'**Update ColocRecup
Public Sub dUpdateColocRecup(ByVal psCtaCod As String, Optional ByVal psFecIngRecup As String = "@", _
        Optional ByVal pnComisionCod As Integer = "-1", Optional ByVal pnSaldoIntComp As Currency = -1, _
        Optional ByVal pnSaldoIntMor As Currency = -1, Optional ByVal pnSaldoGasto As Currency = -1, _
        Optional ByVal pnIntCompGen As Currency = -1, Optional ByVal psMetLiquid As String = "@", _
        Optional ByVal pnTipCj As Integer = -1, _
        Optional ByVal pnDemanda As Integer = -1, Optional ByVal pnNroCalend As Integer = -1, _
        Optional pbEjecBatch As Boolean = False, Optional ByVal psComentario As String = "", _
        Optional ByVal pnVisMetodo As Integer = -1, Optional ByVal pnVisNegociacion As Integer = -1, _
        Optional ByVal pnVisCancelacion As Integer = -1, _
        Optional ByVal pnIntMoraGen As Currency = -1) 'DAOR 20070809, pnIntMoraGen

Dim lsSQL As String

Dim lsSQL1 As String
Dim nVistoMetodo As Integer
Dim nVisNegociacion As Integer
Dim nVisCancelacion As Integer
Dim rs As New ADODB.Recordset

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
    
    '**DAOR 20070809 ****************
    If pnIntMoraGen <> -1 Then
        lsSQL = lsSQL & " nIntMoraGen  = " & pnIntMoraGen & ","
    End If
    '********************************
    
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
      
'    lsSQL1 = "Select nVisMetodo,nVisNegociacion,nVisCancelacion From ColocRecup WHERE cCtaCod ='" & psCtaCod & "'"
'    Set rs = coConex.CargaRecordSet(lsSQL1)
'    If Not (rs.EOF And rs.BOF) Then
'        nVistoMetodo = IIf(IsNull(rs!nVisMetodo), 0, rs!nVisMetodo)
'        nVisNegociacion = IIf(IsNull(rs!nVisNegociacion), 0, rs!nVisNegociacion)
'        nVisCancelacion = IIf(IsNull(rs!nVisCancelacion), 0, rs!nVisCancelacion)
'    End If
    '--- Visto de Recuperaciones AVMM --- 13-12-2006   -------
    If pnVisMetodo <> -1 Then
        lsSQL = lsSQL & " nVisMetodo = " & pnVisMetodo & ","
    End If
    If pnVisNegociacion <> -1 Then
        lsSQL = lsSQL & " nVisNegociacion = " & pnVisNegociacion & ","
    End If
    If pnVisCancelacion <> -1 Then
        lsSQL = lsSQL & " nVisCancelacion = " & pnVisCancelacion & ","
    End If
    '---------------------------------------------------------
      
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    
    lsSQL = lsSQL & " , cComentario = '" & psComentario & "'"
    lsSQL = lsSQL & " WHERE cCtaCod ='" & psCtaCod & "'"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub


'Insert ColocRecupEstado
Public Sub dInsertColocRecupEstado(ByVal psCtaCod As String, ByVal psFecEstado As String, _
        ByVal pnEstado As ColocEstado, ByVal psFlag As String, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT ColocRecupEstado (cCtaCod, dEstado, nEstado, cFlag )" _
        & "VALUES ('" & psCtaCod & "','" & psFecEstado & "'," & pnEstado & ",'" & psFlag & "') "

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
Public Sub dInsertColocRecExpediente(ByVal psCtaCod As String, ByVal psNumExp As String, ByVal pnMonPetit As Currency, _
        ByVal pnMoneda As Integer, ByVal pnViaProce As Integer, ByVal pnEstadoProceso As Integer, _
        Optional ByVal pbEjecBatch As Boolean = False, Optional ByVal pnTipoProceso As Integer = 0, Optional pnTipMedCau As Integer = 0) 'DAOR 20070417, pnTipMedCau

Dim lsSQL As String
    
    If pnTipoProceso <> 0 Then
        lsSQL = "INSERT ColocRecupExpediente (cCtaCod, cNumExp, nMonPetit, nMoneda,nViaProce, nEstadoProceso,nTipoProceso,nTipMedCau ) " _
            & " VALUES ('" & psCtaCod & "','" & Trim(psNumExp) & "'," & pnMonPetit & "," & pnMoneda & "," _
            & pnViaProce & "," & pnEstadoProceso & "," & pnTipoProceso & "," & pnTipMedCau & ") "
    Else
        lsSQL = "INSERT ColocRecupExpediente (cCtaCod, cNumExp, nMonPetit, nMoneda,nViaProce, nEstadoProceso,nTipMedCau) " _
            & " VALUES ('" & psCtaCod & "','" & Trim(psNumExp) & "'," & pnMonPetit & "," & pnMoneda & "," _
            & pnViaProce & "," & pnEstadoProceso & "," & pnTipMedCau & ") "
    End If
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Update ColocRecExpediente
Public Sub dUpdateColocRecupExpediente(ByVal psCtaCod As String, Optional ByVal psNumExp As String = "@", _
        Optional ByVal pnMonPetit As Currency = -1, Optional ByVal pnMoneda As Integer = -1, _
        Optional ByVal pnViaProce As Integer = -1, Optional ByVal pnEstadoProceso As Integer = -1, _
        Optional ByVal pbEjecBatch As Boolean = False, Optional ByVal pnTipoProceso As Integer = 0, _
        Optional ByVal pnTipMedCau As Integer = 0) 'DAOR 20070417, pnTipMedCau

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
    If pnTipoProceso <> 0 Then
        lsSQL = lsSQL & " nTipoProceso = " & pnTipoProceso & ","
    End If
    'If pnEstadoProceso <> -1 Then
    '    lsSQL = lsSQL & " nEstadoViaProce = " & pnViaProce & ","
    'End If
    
    '**DAOR 20070417*************************************************
    If pnTipMedCau <> 0 Then
        lsSQL = lsSQL & " nTipMedCau = " & pnTipMedCau & ","
    End If
    '****************************************************************

    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cCtaCod = '" & psCtaCod & "' "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Insert ColocRecExpedienteInf
Public Sub dInsertColocRecExpedienteInf(ByVal psCtaCod As String, ByVal pnTipoInf As Integer, ByVal psInforme As String, _
        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    
    lsSQL = "INSERT ColocRecupExpedienteInf (cCtaCod, nTipoInf, mInforme) " _
        & " VALUES ('" & psCtaCod & "'," & pnTipoInf & ",'" & psInforme & "' ) "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Update ColocRecExpedienteInf
Public Sub dUpdateColocRecupExpedienteInf(ByVal psCtaCod As String, ByVal pnTipoInf As ColocRecExpedTipoInf, _
        Optional ByVal psInforme As String = "@", Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "DELETE ColocRecupExpedienteInf WHERE cCtaCod = '" & psCtaCod & "' And nTipoInf = " & pnTipoInf & " "
    lsSQL = lsSQL & "INSERT ColocRecupExpedienteInf (cCtaCod, nTipoInf, mInforme) VALUES ('" & psCtaCod & "'," & pnTipoInf & ",'" & psInforme & "') "

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
Public Sub dInsertColocRecComisionHist(ByVal psCtaCod As String, ByVal pnComisionCod As Integer, _
        ByVal pnMovNro As Long, Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    
    lsSQL = "INSERT ColocRecupComisionHist (cCtaCod, nComisionCod, nMovNro) " _
        & " VALUES ('" & psCtaCod & "'," & pnComisionCod & "," & pnMovNro & " ) "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Insert ColocRecGastos
Public Sub dInsertColocRecupGastos(ByVal psCtaCod As String, ByVal pnNroGastoCta As Integer, _
        ByVal pnPrdConceptoCod As Integer, _
        ByVal psFecAsigna As String, ByVal pnMonto As Currency, ByVal pnMontoPagado As Currency, _
        ByVal pnColocRecGastoEstado As ColocRecGastoEstado, ByVal psMotivoGasto As String, _
        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    
    lsSQL = "INSERT ColocRecupGastos (cCtaCod,nNroGastoCta, nPrdConceptoCod, dAsigna,nMonto, " _
        & " nMontoPagado, nColocRecGastoEstado,cMotivoGasto ) " _
        & " VALUES ('" & psCtaCod & "'," & pnNroGastoCta & "," & pnPrdConceptoCod & ",'" & psFecAsigna & "'," & pnMonto & "," _
        & pnMontoPagado & "," & pnColocRecGastoEstado & ",'" & psMotivoGasto & "'  )"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Update ColocRecGastos
Public Sub dUpdateColocRecupGastos(ByVal psCtaCod As String, ByVal pnNroGastoCta As Integer, _
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
    lsSQL = lsSQL & " WHERE cCtaCod = '" & psCtaCod & "' AND  nNroGastoCta = " & pnNroGastoCta & " "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub


'Delete ColocRecExpedienteInf
Public Sub dDeleteColocRecupExpedienteInf(ByVal psCtaCod As String, Optional ByVal pnTipoInf As ColocRecExpedTipoInf = -1, _
    Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "DELETE ColocRecupExpedienteInf WHERE cCtaCod = '" & psCtaCod & "'  "
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
Public Sub dDeleteColocRecupExpediente(ByVal psCtaCod As String, Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "DELETE ColocRecupExpediente WHERE cCtaCod = '" & psCtaCod & "' "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Delete ColocRecGastos
Public Sub dDeleteColocRecupGastos(ByVal psCtaCod As String, Optional ByVal pnNroGastoCta As Integer = -1, _
        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    
    lsSQL = "DELETE ColocRecupGastos WHERE cCtaCod = '" & psCtaCod & "' "
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
Public Sub dDeleteColocRecup(ByVal psCtaCod As String, Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "DELETE ColocRecup WHERE  cCtaCod ='" & psCtaCod & "' "
    
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
        Optional ByVal pbEjecBatch As Boolean = False, _
        Optional ByVal psAgeCodAct As String = "", _
        Optional ByVal psTpoProdCod As String = "", _
        Optional ByVal pbMalCalificacion As Boolean = False, _
        Optional ByVal pnTasaIntVencidoCrePigCliMalCal = 0#, Optional psLineaCredito As String = "", Optional pnTasaGracia As Double = 0#, Optional pnTasaMora As Double = 0#, Optional pnHolograma As Long = 0#)
        'JUCS Agregó pnHolograma TI ERS-063-2017
        'MAVM se agregarron las var: psAgeCodAct, psTpoProdCod BAS II
        '***Agregado por ELRO el parametro "pbMalCalificacion" el 20120120, según Acta N° 002-2012/TI-D
Dim loParam As DCOMColPCalculos, nTasaMor As Double, nTasaGra As Double
Set loParam = New DCOMColPCalculos

    'ALPA 20150701**************************************************
    'ntasamor = loParam.dObtieneTasaInteres("0101113050101", "03")
    'ntasagra = loParam.dObtieneTasaInteres("0101113050101", "06")
    'BAS II
    
    '***Modificado por ELRO el 20120120, según Acta N° 002-2012/TI-D
    'ntasamor = loParam.dObtieneTasaInteres("0101117550101", "03")
    'ntasagra = loParam.dObtieneTasaInteres("0101117550101", "06")
    
'    If pbMalCalificacion Then
'        ntasamor = loParam.dObtieneTasaInteres("0101117550101", "03")
'        ntasagra = pnTasaIntVencidoCrePigCliMalCal
'    Else
'        ntasamor = loParam.dObtieneTasaInteres("0101117550101", "03")
'        ntasagra = loParam.dObtieneTasaInteres(psLineaCredito, "06")
'    End If
    '***Fin Modificado por ELRO*************************************
    'ALPA 20150620**************************************************
    nTasaMor = pnTasaMora
    nTasaGra = pnTasaGracia
    '***************************************************************
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
      & "VALUES('" & psCuenta & "'," & nTasaGra & ",6)"
If pbEjecBatch Then
    coConex.AdicionaCmdBatch lsSQL
Else
    coConex.Ejecutar lsSQL
End If

lsSQL = "INSERT INTO PRODUCTOTASAINTERES(cCtaCod,nTasaInteres ,nPrdTasaInteres) " _
      & "VALUES('" & psCuenta & "'," & nTasaMor & ",3)"
If pbEjecBatch Then
    coConex.AdicionaCmdBatch lsSQL
Else
    coConex.Ejecutar lsSQL
End If

'**Insert Colocaciones
lsSQL = "INSERT Colocaciones (cCtaCod,nPlazo,dVenc,nMontoCol,nColocCalendCod,cUltimaActualizacion,cLineaCred,dVigencia, cAgeCodAct, cTpoProdCod, cTpoCredCod) " _
    & "VALUES ('" & psCuenta & "'," & pnPlazo & ",'" & psFecVenc & "'," & pnSaldo & "," & gColocCalendCodPFCF & ",'" & psMovNro & "','" & "0101117550101" & "','" & psFechaHora & "','" & psAgeCodAct & "','" & psTpoProdCod & "', 755)"
If pbEjecBatch Then
    coConex.AdicionaCmdBatch lsSQL
Else
    coConex.Ejecutar lsSQL
End If
'**Insert ColocPignoraticio
'Modificación JUCS TI ERS 063-2017 AGREGÓ pnHolograma
lsSQL = "INSERT ColocPignoraticio (cCtaCod,nOroBruto,nOroNeto,nPiezas,nTasacion,nNroRenov,nNroRemate,nNroDuplic,cPrdCtaTpo,cAgeBoveda,cLote,nHolograma) " _
     & "VALUES ('" & psCuenta & "'," & pnOroBruto & "," & pnOroNeto & "," & pnPiezas & "," & pnTasacion & ",0,0,0,'" & psTipoContrato & "','" & Mid(psCuenta, 1, 5) & "','" & psLote & "','" & pnHolograma & "') "
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
          & "VALUES ('" & psCuenta & "','21'," & pnKi21 & ")"
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

'Insert ColocRecupActProcesales
'** Modificado por DAOR 20070124
'** Se aumentó el parametro tipo de actuación procesal
Public Sub dInsertColocRecupActProcesales(ByVal psCtaCod As String, ByVal psMovNro As String, ByVal psComenta As String, _
        Optional ByVal pbEjecBatch As Boolean = False, Optional ByVal pdFechaA As Date, Optional ByVal pdFechaV As Date, _
        Optional ByVal pbAlerta As Integer, Optional ByVal gsProyectoActual As String, Optional ByVal pnTipoAct As Integer)

Dim lsSQL As String
    If gsProyectoActual = "H" Then
        lsSQL = "INSERT ColocRecupActProcesales (cCtaCod, cMovNro, cComenta,dFechaAviso,dFechaVencimiento,bAlerta) " _
            & " VALUES ('" & psCtaCod & "','" & psMovNro & "','" & psComenta & "','" & Format(pdFechaA, "yyyymmdd") & "','" & Format(pdFechaV, "yyyymmdd") & "', " & pbAlerta & " ) "
    Else
        lsSQL = "INSERT ColocRecupActProcesales (cCtaCod, cMovNro, cComenta,dFechaAviso,dFechaVencimiento,nTipoAct) " _
            & " VALUES ('" & psCtaCod & "','" & psMovNro & "','" & psComenta & "','" & Format(pdFechaA, "yyyymmdd") & "','" & Format(pdFechaV, "yyyymmdd") & "'," & pnTipoAct & ") "
    End If
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub


Public Sub dInsertHistColocPigJoyaDetRet(ByVal prJoyas As ADODB.Recordset, ByVal pnNumRetasacion As Integer, ByVal psMovNro As String, ByVal psCtaCod As String)
'inserta Joyas para cada registro de prJoyas
Dim lsSQL As String
Do While Not prJoyas.EOF
    lsSQL = "INSERT HistColocPigJoyaDetRet (cCtaCod,cKilataje, nItem, nPiezas, " _
          & "nPesoBruto, nPesoNeto, nValTasac, cDescrip,nnumretasacion,cmovnro ) " _
          & "VALUES ('" & psCtaCod & "','" & Trim(Str(Left(prJoyas.Fields("ckilataje"), 2))) & "'," _
          & prJoyas.Fields("nitem") & "," & prJoyas.Fields("npiezas") & "," _
          & prJoyas.Fields("nPesoBruto") & "," & prJoyas.Fields("nPesoNeto") & "," _
          & prJoyas.Fields("nValTasac") & ",'" & prJoyas.Fields("cDescrip") & "'," & pnNumRetasacion & ",'" & psMovNro & "')"
    
        coConex.Ejecutar lsSQL
    
    prJoyas.MoveNext
Loop
End Sub

Public Sub dUpdateColocPigJoyaDet(ByVal prJoyas As ADODB.Recordset, ByVal psCtaCod As String)
'inserta Joyas para cada registro de prJoyas
Dim lsSQL As String
Do While Not prJoyas.EOF
    lsSQL = " Update ColocPigJoyaDet "
    lsSQL = lsSQL & " Set cKilataje='" & prJoyas.Fields("Material") & "', nPesoBruto= " & prJoyas.Fields("pBruto") & ",nPesoNeto=" & prJoyas.Fields("pNeto")
    lsSQL = lsSQL & " where cctacod='" & psCtaCod & "' and nitem=" & prJoyas.AbsolutePosition
    coConex.Ejecutar lsSQL
    
    prJoyas.MoveNext
Loop

End Sub


Public Sub dUpdateColocPigJoya(ByVal psCtaCod As String, Optional n14k As Double = 0, Optional n16k As Double = 0, Optional n18k As Double = 0, Optional n21k As Double = 0)
'inserta Joyas para cada registro de prJoyas
    Dim lsSQL As String
    lsSQL = "Delete ColocPigJoya where cctacod='" & psCtaCod & "'"
    coConex.Ejecutar lsSQL

If n14k > 0 Then
    lsSQL = "Insert into ColocPigJoya "
    lsSQL = lsSQL & " values('" & psCtaCod & "','14'," & n14k & " )"
    coConex.Ejecutar lsSQL
End If
    
If n16k > 0 Then
    lsSQL = "insert into ColocPigJoya "
    lsSQL = lsSQL & " values('" & psCtaCod & "','16'," & n16k & " )"
    coConex.Ejecutar lsSQL
End If
    
If n18k > 0 Then
    lsSQL = "insert into ColocPigJoya "
    lsSQL = lsSQL & " values('" & psCtaCod & "','18'," & n18k & " )"
    coConex.Ejecutar lsSQL
End If
    
If n21k > 0 Then
    lsSQL = "insert into ColocPigJoya "
    lsSQL = lsSQL & " values('" & psCtaCod & "','21'," & n21k & " )"
    coConex.Ejecutar lsSQL
End If
    
End Sub

'Insert ColocRecExpedienteRelacionado
Public Sub dInsertColocRecExpedienteRelacionado(ByVal psCtaCod As String, ByVal psNumExp As String, ByVal pnMonPetit As Currency, _
        ByVal pnMoneda As Integer, ByVal pnViaProce As Integer, ByVal pnEstadoProceso As Integer, _
        Optional ByVal pbEjecBatch As Boolean = False, Optional ByVal pnNroCorrelativo As Integer, Optional ByVal pnTipoProceso As Integer, Optional ByVal pnTipoProcesal As Integer)

Dim lsSQL As String
    
    lsSQL = " INSERT ColocRecupExpedienteRelacionados (cCtaCod, cNumExp, nMonPetit, nMoneda,nViaProce, nEstadoProceso ,nNroCorreRelacion,nTipoProceso,nTipoProcesal) " & _
            " VALUES ('" & psCtaCod & "','" & Trim(psNumExp) & "'," & pnMonPetit & "," & pnMoneda & ", " & pnViaProce & "," & pnEstadoProceso & "," & pnNroCorrelativo & "," & pnTipoProceso & "," & pnTipoProcesal & ") "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Update ColocRecExpedienteRelacionado
Public Sub dUpdateColocRecupExpedienteRelacionado(ByVal psCtaCod As String, Optional ByVal psNumExp As String = "@", _
        Optional ByVal pnMonPetit As Currency = -1, Optional ByVal pnMoneda As Integer = -1, _
        Optional ByVal pnViaProce As Integer = -1, Optional ByVal pnEstadoProceso As Integer = -1, _
        Optional ByVal pbEjecBatch As Boolean = False, Optional ByVal pnNroCorrelativo As String, Optional ByVal pnTipoProceso As Integer, Optional ByVal pnTipoProcesal As Integer)

Dim lsSQL As String
    lsSQL = "UPDATE ColocRecupExpedienteRelacionados SET "

    If psNumExp <> "@" Then
        lsSQL = lsSQL & " cNumExp = '" & Trim(psNumExp) & "',"
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
    
    lsSQL = lsSQL & " nTipoProceso = " & pnTipoProceso & ","
    
    lsSQL = lsSQL & " nTipoProcesal = " & pnTipoProcesal & ","
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cCtaCod = '" & psCtaCod & "' and  nNroCorreRelacion = " & pnNroCorrelativo & " "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'Update ColocRecExpedienteInfRelacionado
Public Sub dUpdateColocRecupExpedienteInfRelacionado(ByVal psCtaCod As String, ByVal pnTipoInf As ColocRecExpedTipoInf, _
        Optional ByVal psInforme As String = "@", Optional ByVal pbEjecBatch As Boolean = False, Optional ByVal pnNroCorrelativo As Integer)

Dim lsSQL As String
    lsSQL = "DELETE ColocRecupExpedienteRelacionadosInf WHERE cCtaCod = '" & psCtaCod & "' And nTipoInf = " & pnTipoInf & " "
    lsSQL = lsSQL & "INSERT ColocRecupExpedienteRelacionadosInf (cCtaCod, nTipoInf, mInforme,nNroCorreRelacion) VALUES ('" & psCtaCod & "'," & pnTipoInf & ",'" & psInforme & "'," & pnNroCorrelativo & " ) "

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

Public Sub dInsertExpedientePersona(ByVal psCuenta As String, ByVal psPersona As String, ByVal pnCorrelativo As Integer, ByVal pnRelacion As ColocRelacPers)

Dim lsSQL As String

    lsSQL = "INSERT ColocRecupExpedientePersona (cCtaCod,cPersCod,nNroCorreRelacion,nPrdPersRelac) " _
        & "VALUES ('" & psCuenta & "','" & psPersona & "'," & pnCorrelativo & "," & pnRelacion & ")"
    
    coConex.Ejecutar lsSQL
End Sub

Public Sub dDeleteExpedientePersona(ByVal psCuenta As String, ByVal psPersona As String, ByVal pnRelacion As ColocRelacPers, _
    Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = " DELETE ColocRecupExpedientePersona WHERE cCtaCod ='" & psCuenta & "' " _
          & " AND nPrdPersRelac = " & pnRelacion & " " _
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    
End Sub

'**DAOR 20070414, Metodo que inserta la distibución de los montos de CIMG a cobrar
'**para una determinada cuenta
Public Sub dInsertColocRecupDistribucionCobr(ByVal psCtaCod As String, pdFecha As Date, _
        ByVal pnCapital As Currency, ByVal pnIntComp As Currency, ByVal pnMora As Currency, _
        ByVal pnGasto As Currency, pnComisionAbog As Currency, ByVal psUltimaActualizacion As String, _
        Optional ByVal pnRegCanc As Integer) '** Juez 20120421 Se agregó pnRegCanc

Dim lsSQL As String

    lsSQL = "Insert into ColocRecupDistribucionCobr (cCtaCod, dFecha, nCapital,nIntComp, nMora,nGasto,nComiAbog,cUltimaActualizacion"
    If pnRegCanc = 1 Then '** Juez 20120421
        lsSQL = lsSQL & ",nRegCancelacion"
    End If
    
    lsSQL = lsSQL & ") Values ('" & psCtaCod & "','" & Format(pdFecha, "yyyymmdd") & "'," & pnCapital & "," & pnIntComp & "," _
        & pnMora & "," & pnGasto & "," & pnComisionAbog & ",'" & psUltimaActualizacion & "'"
    If pnRegCanc = 1 Then '** Juez 20120421
        lsSQL = lsSQL & "," & pnRegCanc
    End If
    
    lsSQL = lsSQL & ")"
        
    coConex.Ejecutar lsSQL
End Sub

'**DAOR 20070414, Metodo que elimina el registro de distibución de los montos de CIMG a cobrar
'**para una determinada cuenta
Public Sub dDeleteColocRecupDistribucionCobr(ByVal psCtaCod As String, pdFecha As Date)

Dim lsSQL As String

    lsSQL = " Delete ColocRecupDistribucionCobr Where cCtaCod ='" & psCtaCod & "' " _
          & " AND dFecha='" & Format(pdFecha, "yyyymmdd") & "'"
        coConex.Ejecutar lsSQL
End Sub

Public Sub dBeginTrans()
    coConex.BeginTrans
    cnConecOpenValor = 1 'RECO----*****
End Sub

Public Sub dRollbackTrans()
    coConex.RollbackTrans
    cnConecOpenValor = 0 'RECO----*****
End Sub
Public Sub dCommitTrans()
    coConex.CommitTrans
    cnConecOpenValor = 0 'RECO----*****
End Sub

Private Sub Class_Initialize()
    Dim loIni As New COMConecta.DCOMClasIni
    csConexion = loIni.CadenaConexion
    csNegocio = loIni.BaseNegocio
    csCentralPer = loIni.BasePersonas
    csCentralCom = loIni.BaseComunes
    csCentralImg = loIni.BaseImagenes
    csAdminist = loIni.BaseAdministracion
    If coConex.AbreConexion(csConexion) = False Then
        Call oError.RaiseError(oError.MyUnhandledError, "DColPContrato:Initialize. Error en Conexion a Base de datos")
    End If
End Sub

Private Sub Class_Terminate()
    coConex.CierraConexion
    Set coConex = Nothing
End Sub

'*** PEAC 20080711
Public Sub QuitaNotificaAdjudica()
    Dim lsSQL As String
        lsSQL = " exec stp_upd_QuitaNotificaAdjudica "
    coConex.Ejecutar lsSQL
End Sub

'*** PEAC 20080715 - Update ColocPigNotifi
Public Sub dUpdateColocPigNotifi(ByVal psCtaCod As String, _
        Optional ByVal pnNotifi As Integer = -1, Optional ByVal pnEstado As Integer = -1, _
        Optional ByVal pcMovNroNotif As String = "@", _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

'*** PEAC 20101230
'    lsSQL = "UPDATE ColocPigNotifi SET "
'
'    If pnNotifi <> -1 Then
'        lsSQL = lsSQL & " nNotifi = " & pnNotifi & ","
'    End If
'
'    If pnEstado <> -1 Then
'        lsSQL = lsSQL & " nEstado = " & pnEstado & ","
'    End If
'
'    If pcMovNroNotif <> "@" Then
'        lsSQL = lsSQL & " cMovNroNotif = '" & pcMovNroNotif & "',"
'    End If
'
'    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
'    lsSQL = lsSQL & " WHERE cCtaCod ='" & psCtaCod & "' and cNroProceso=(select max(cNroProceso) from ColocPigNotifi where cCtaCod = '" & psCtaCod & "') "

    lsSQL = " exec stp_upd_dUpdateColocPigNotifi '" & psCtaCod & "'," & pnNotifi & "," & pnEstado & ",'" & pcMovNroNotif & "'"
'*** FIN PEAC

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub
'''**---

''*** PEAC 20090509
Public Function CapAperturaCuenta(ByVal nProducto As Producto, ByVal nMoneda As Moneda, ByVal rsPers As ADODB.Recordset, ByVal sMovNro As String, _
        ByVal sAgencia As String, ByVal nOperacion As CaptacOperacion, ByVal nTasa As Double, ByVal nSaldo As Double, _
        ByVal dFecha As Date, ByVal nFirmas As Integer, ByVal nPersoneria As PersPersoneria, ByVal sGlosa As String, _
        ByVal nTipoCuenta As ProductoCuentaTipo, _
        ByVal nTipoTasa As CaptacTipoTasa, Optional bCheque As Boolean = False, _
        Optional sNroDoc As String, Optional sCodIF As String, _
        Optional ByVal bOrdPag As Boolean = False, Optional nPlazo As Long = 0, _
        Optional nFormaRetiro As CaptacPFFormaRetiro = gCapPFFormRetMensual, _
        Optional bCtaAboInt As Boolean = False, Optional sCtaCodAbono As String = "", _
        Optional nPorcRetCTS As Double = 0, Optional sInstitucion As String = "", _
        Optional sPersLavDinero As String = "", _
        Optional ByVal pnPrograma As Integer = 0, Optional ByVal pnMontoAbonar As Double, Optional ByVal pnPlazoAbono As Integer, Optional ByVal psPromotor As String) As String

Dim nErrNum As Integer
Dim nErrDesc As String, sCuenta As String
Dim bTrans As Boolean
Dim nMovNro As Long

bTrans = False

bTrans = True
sCuenta = AgregaNuevaCaptacion(nProducto, nMoneda, Mid(sMovNro, 15, 3) & sAgencia, nTasa, nSaldo, dFecha, nFirmas, nPersoneria, _
    nTipoCuenta, sMovNro, nTipoTasa, bCheque, bOrdPag, nPlazo, nFormaRetiro, bCtaAboInt, sCtaCodAbono, nPorcRetCTS, sInstitucion, _
    pnPrograma, pnMontoAbonar, pnPlazoAbono, psPromotor)
'AgregaMov sMovNro, nOperacion, sGlosa
nMovNro = GetnMovNro(sMovNro)
Select Case nOperacion
    Case gAhoApeChq, gPFApeChq, gCTSApeChq
        AgregaMovCap nMovNro, nOperacion, sCuenta, nSaldo, 0, nSaldo
        AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nSaldo
        AgregaCuentaDocumento sCuenta, TpoDocCheque, sNroDoc, sCodIF, sMovNro, nMovNro
    Case gAhoApeTransf, gPFApeTransf, gCTSApeTransf
        AgregaMovCap nMovNro, nOperacion, sCuenta, nSaldo, nSaldo, nSaldo
        AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nSaldo
        'Aqui debe registrarse la referencia con la transferencia de caja general
    Case Else
        AgregaMovCap nMovNro, nOperacion, sCuenta, nSaldo, nSaldo, nSaldo
        AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nSaldo
End Select
rsPers.MoveFirst
Do While Not rsPers.EOF
    AgregaNuevoProdPers sCuenta, rsPers("Codigo"), CLng(Trim(Right(rsPers("Relacion"), 4)))
    rsPers.MoveNext
Loop

If psPromotor <> "" Then 'JOEP20190730 Mejora Boton Apertura
    AgregaNuevoProdPers sCuenta, psPromotor, gCapRelPersPromotor
End If 'JOEP20190730 Mejora Boton Apertura

CapAperturaCuenta = sCuenta

End Function

'''*** PEAC 20090509
Public Function AgregaNuevaCaptacion(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
        ByVal sAgencia As String, ByVal nTasa As Double, ByVal nSaldo As Double, _
        ByVal dFecha As Date, ByVal nFirmas As Integer, ByVal nPersoneria As PersPersoneria, _
        ByVal nTipoCuenta As ProductoCuentaTipo, ByVal sMovNro As String, _
        ByVal nTipoTasa As CaptacTipoTasa, Optional bCheque As Boolean = False, _
        Optional ByVal bOrdPag As Boolean = False, Optional nPlazo As Long = 0, _
        Optional nFormaRetiro As CaptacPFFormaRetiro = gCapPFFormRetMensual, _
        Optional bCtaAboInt As Boolean = False, Optional sCtaCodAbono As String = "", _
        Optional nPorcRetCTS As Double = 0, Optional sInstitucion As String = "", _
        Optional ByVal pnPrograma As Integer = 0, Optional ByVal pnMontoAbonar As Double, Optional ByVal pnPlazoAbono As Integer, Optional ByVal psPromotor As String) As String

'ARCV 13-02-2007
'Se Agrego para los cambios en Captaciones pnPrograma,pnMontoAbonar,pnPlazoAbono
'--------
Dim clsGen As COMDConstSistema.DCOMGeneral
Dim sCuenta As String, sFecha As String
Dim sFecUltCierre As String, sOrdPag As String, sCtaAbonoIntPF As String
Dim nSaldoDisp As Double, nSaldRetCTS As Double
Dim sSql As String
Set clsGen = New COMDConstSistema.DCOMGeneral
sCuenta = clsGen.GeneraNuevaCuenta(sAgencia, nProducto, nMoneda)
Set clsGen = Nothing
sFecha = Format$(dFecha, "mm/dd/yyyy") & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)
sSql = "Insert Producto (cCtaCod,nTasaInteres,nSaldo,nPrdEstado,dPrdEstado,nTransacc) " _
    & "Values ('" & sCuenta & "'," & nTasa & "," & nSaldo & "," & gCapEstActiva & ",'" & sFecha & "',0)"
coConex.ConexionActiva.Execute sSql
nSaldoDisp = IIf(bCheque, 0, nSaldo)
sFecUltCierre = Format$(DateAdd("d", -1, dFecha), "mm/dd/yyyy")

'ARCV 13-02-2007
If nProducto <> gCapAhorros Then
    sSql = "Insert Captaciones (cCtaCod,nSaldoDisp,nPersoneria,nFirmas,nIntAcum,dUltCierre,dApertura,nPrdCtaTpo,nPrdTasaInteres,cUltimaActualizacion) " _
    & "Values ('" & sCuenta & "'," & nSaldoDisp & "," & nPersoneria & "," & nFirmas & ",0,'" & sFecUltCierre & "','" & sFecha & "'," & nTipoCuenta & "," _
    & nTipoTasa & ",'" & sMovNro & "')"
    
    coConex.ConexionActiva.Execute sSql
Else
    sSql = "Insert Captaciones (cCtaCod,nSaldoDisp,nPersoneria,nFirmas,nIntAcum,dUltCierre,dApertura,nPrdCtaTpo,nPrdTasaInteres,cUltimaActualizacion,nTpoPrograma) " _
        & "Values ('" & sCuenta & "'," & nSaldoDisp & "," & nPersoneria & "," & nFirmas & ",0,'" & sFecUltCierre & "','" & sFecha & "'," & nTipoCuenta & "," _
        & nTipoTasa & ",'" & sMovNro & "'," & pnPrograma & ")"
        
    coConex.ConexionActiva.Execute sSql
End If
'-------

Select Case nProducto
    Case gCapAhorros
        sOrdPag = IIf(bOrdPag, "1", "0")
        'ARCV 13-02-2007
        sSql = "Insert CaptacAhorros (cCtaCod,nSaldoAnterior,bOrdPag,nSobregiro,dUltContacto,nPlazo,nMontoAbono) " _
            & "Values ('" & sCuenta & "',0," & sOrdPag & ",0,'" & sFecha & "'," & pnPlazoAbono & "," & pnMontoAbonar & ")"
        '------
    Case gCapPlazoFijo
        sCtaAbonoIntPF = IIf(bCtaAboInt, "1", "0")
        sSql = "Insert CaptacPlazoFijo (cCtaCod,nPlazo,nIntPag,dRenovacion,nApertura,dAuxiliar,nFormaRetiro,nDuplicado,bAbonoIntCtaAho) " _
            & "Values ('" & sCuenta & "'," & nPlazo & ",0,'" & sFecha & "'," & nSaldo & ",'" & sFecha & "'," & nFormaRetiro & ",0," & sCtaAbonoIntPF & ")"
    Case gCapCTS
        nSaldRetCTS = Round(nSaldo * nPorcRetCTS / 100, 2)
        sSql = "Insert CaptacCTS (cCtaCod,nSaldRetiro,nIntSaldo,cCodInst) " _
            & "Values ('" & sCuenta & "'," & nSaldRetCTS & ",0,'" & sInstitucion & "')"
End Select
coConex.ConexionActiva.Execute sSql
If nProducto = gCapPlazoFijo And bCtaAboInt Then
    sSql = "Insert CaptacCtaAboIntPF (cCtaCod,cCtaCodAbono) " _
        & "Values ('" & sCuenta & "','" & sCtaCodAbono & "')"
    coConex.ConexionActiva.Execute sSql
End If
AgregaNuevaCaptacion = sCuenta
End Function

'*** PEAC 20090509
Public Function CapAbonoCuentaAho(ByRef pnMovNro As Long, ByRef pMatAho As Variant, ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nOperacion As CaptacOperacion, ByVal sMovNro As String, ByVal sGlosa As String, _
            Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", Optional sCodIF As String = "", _
            Optional dFechaValor As Date, Optional bActivaCta As Boolean = True, Optional bNumExtracto As Boolean = True, _
            Optional dFecSis As Date = CDate("01/01/1950"), Optional sNomAge As String = "", _
            Optional pbITFAplica As Boolean = True, Optional pnITFValor As Double = 0, Optional pbITFAsumido As Boolean = False, _
            Optional psItfOperacion As String = "", Optional ByVal pnEstadoTransf As ColocEstado, Optional ByVal pnCapPag As Currency) As Double
            'FRHU 20150530 ERS022-2015: Se agrego pnEstadoTransf y pnCapPag
Dim rsCta As ADODB.Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim nSaldoInac As Double
Dim dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean
Dim nNumExtracto As Long, nDiasTranscurridos As Long, nExtracto As Long
Dim sMsgOpe As String
Dim nMovNro As Long
bTrans = False
'Obtiene los datos para el calculo

Set rsCta = GetDatosCuentaAho(sCuenta)

nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
bInactiva = IIf(rsCta("bInactiva") = 0, False, True)
nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing

On Error GoTo ErrGraba
If sGlosa = "" Then sGlosa = "Abono Cuenta = " & sCuenta
If bActivaCta Then
    If bInactiva Then
        nSaldoInac = nSaldoCnt
    End If
End If
'Calcula intereses
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
pMatAho(2) = CDbl(Format(nIntGanado, "#0.00"))

dUltMov = DateAdd("d", -1, dFecSis)
nSaldoCnt = nSaldoCnt + nMonto
nMovNro = pnMovNro

If sNroDoc = "" Then ' Si la operacion es en efectivo
    Call ActualizaSaldoAnteriorAho(sCuenta, nSaldoDisp)
    nSaldoDisp = nSaldoDisp + nMonto
    Call ActualizaAbonoCaptacion(sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True)
    sMsgOpe = "Depósito Efectivo"
End If
    Call AgregaMovCap(nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt)

If bInactiva Then
    Call AgregaMovCap(nMovNro, gAhoEstInacAct, sCuenta, nSaldoInac, nSaldoDisp, nSaldoCnt)
    Call AgregaMovCapDet(nMovNro, gAhoEstInacAct, sCuenta, gConcCapital, nSaldoInac)
    Call ActualizaEstadoCuenta(sCuenta, gCapEstActiva, dFecSis, sMovNro)
End If

Call AgregaMovCapDet(nMovNro, nOperacion, sCuenta, gConcCapital, nMonto)
Call UltimaActualizacionCuenta(sCuenta, sMovNro)
CapAbonoCuentaAho = nSaldoCnt

pMatAho(8) = nSaldoCnt
pMatAho(9) = nSaldoCnt

If pbITFAplica And pnITFValor > 0 Then
    If pbITFAsumido Then
        AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp, nSaldoCnt
        AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFAsumido, pnITFValor
    Else
        AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp - pnITFValor, nSaldoCnt - pnITFValor
        AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFCliente, pnITFValor
    End If
End If

'FRHU 20150530 ERS022-2015: Movimiento de Reversión
If nOperacion = "200222" Or nOperacion = "200223" Or nOperacion = "200224" Then
    If pnEstadoTransf = gColocEstRecVigJud Or pnEstadoTransf = gColocEstRecRefiJud Then
        AgregaMovCap nMovNro, 107315, sCuenta, pnCapPag, nSaldoDisp, nSaldoCnt
        AgregaMovCapDet nMovNro, 107315, sCuenta, 0, pnCapPag
    ElseIf pnEstadoTransf = gColocEstVigVenc Or pnEstadoTransf = gColocEstRefVenc Then
        AgregaMovCap nMovNro, 107316, sCuenta, pnCapPag, nSaldoDisp, nSaldoCnt
        AgregaMovCapDet nMovNro, 107316, sCuenta, 0, pnCapPag
    End If
End If
'FIN FRHU 20150530
'If (pbITFAplica And pnITFValor > 0) And (Not pbITFAsumido) And (Val(psItfOperacion) = gITFCobroCargo Or Val(psItfOperacion) = gITFCobroCMACCargo) Then
If (pbITFAplica And pnITFValor > 0) And (Not pbITFAsumido) And (Val(psItfOperacion) = gITFCobroCargoPigno) Then '' Or Val(psItfOperacion) = gITFCobroCMACCargo) Then
    ActualizaCargoCaptacion sCuenta, pnITFValor, pnITFValor, nIntGanado, dUltMov, sMovNro, bNumExtracto
    CapAbonoCuentaAho = nSaldoCnt - pnITFValor
    pMatAho(8) = nSaldoCnt - pnITFValor
    pMatAho(9) = nSaldoCnt - pnITFValor
End If


Exit Function
ErrGraba:
    Err.Raise Err.Number, "", Err.Description
    CapAbonoCuentaAho = 0
End Function

''*** PEAC 20090509
Public Sub ExoneraInactivas(ByVal sCuenta As String, ByVal sMovNro As String)
Dim sSql As String

sSql = "Update CaptacAhorros Set bNoCobroInactivas =1" & ",cNoCobroInactivas= '" & sMovNro & "' WHERE cCtaCod = '" & sCuenta & "'"
coConex.ConexionActiva.Execute sSql
End Sub

''*** PEAC 20090509
Public Sub dInsertColocCargoAutoma(ByVal psCtaCod As String, ByVal pnOrden As Integer, ByVal psCodCtaAho As String, Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String

    lsSQL = "INSERT INTO ColocCargoAutoma(cCtaCod,nOrden,cCodCtaAho,cEstado)"
    lsSQL = lsSQL & " VALUES('" & psCtaCod & "'," & pnOrden & ",'" & psCodCtaAho & "',1)"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

''*** PEAC 20090509
'CTI4 ERS0112020 Comento
'Public Function GetDatosCuentaAho(ByVal sCuenta As String) As ADODB.Recordset
'    Dim sSql As String
'Dim rsCta As ADODB.Recordset
'
'sSql = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, " _
'    & "C.nSaldoDisp, C.nPersoneria, C.nFirmas, C.nIntAcum, C.dUltCierre, C.dApertura, " _
'    & "C.nPrdCtaTpo, C.nPrdTasaInteres, A.bOrdPag, A.dUltContacto, T.cConsDescripcion cEstado, " _
'    & "T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa, A.bInactiva, A.bInmovilizada, A.nSobregiro FROM Producto P " _
'    & "INNER JOIN Captaciones C INNER JOIN CaptacAhorros A ON C.cCtaCod = A.cCtaCod ON " _
'    & "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
'    & "INNER JOIN Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN Constante T2 " _
'    & "ON C.nPrdTasaInteres = T2.nConsValor WHERE P.cCtaCod = '" & sCuenta & "' AND " _
'    & "T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
'    & "AND T2.nConsCod = " & gCaptacTipoTasa
'
''sSql = " exec stp_sel_ObtieneDatosParaDesembCtaAho "
'
'Set rsCta = New ADODB.Recordset
'rsCta.CursorLocation = adUseClient
'rsCta.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
'Set rsCta.ActiveConnection = Nothing
'Set GetDatosCuentaAho = rsCta
'Set rsCta = Nothing
'End Function


''*** PEAC 20090509
Public Function GetInteres(ByVal nCapital As Double, ByVal nTasa As Double, _
            ByVal nPlazo As Long, Optional nTipoInteres As TipoCalculoInteres = TpoCalcIntSimple) As Double
If nTipoInteres = TpoCalcIntSimple Then
    GetInteres = Round((nTasa / 36000) * nPlazo * nCapital, 2)
ElseIf nTipoInteres = TpoCalcIntCompuesto Then
    GetInteres = Round((((nTasa / 36000) + 1) ^ nPlazo - 1) * nCapital, 2)
End If
End Function

''***PEAC 20090509
Public Function GetnMovNro(ByVal sMovNro As String) As Long
Dim rsMov As ADODB.Recordset
Dim sSql As String

sSql = "Select nMovNro From Mov Where cMovNro = '" & sMovNro & "'"
Set rsMov = New ADODB.Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
If rsMov.EOF And rsMov.BOF Then
    GetnMovNro = 0
Else
    GetnMovNro = rsMov("nMovNro")
End If
rsMov.Close
End Function

''*** PEAC 20090509
Public Sub ActualizaSaldoAnteriorAho(ByVal sCuenta As String, ByVal nSaldoAnt As Double)
Dim sSql As String

sSql = "Update CaptacAhorros Set nSaldoAnterior = " & nSaldoAnt & "  WHERE cCtaCod = '" & sCuenta & "'"
coConex.ConexionActiva.Execute sSql
End Sub

''*** PEAC 20090509
Public Sub ActualizaAbonoCaptacion(ByVal sCuenta As String, ByVal nMontoCnt As Double, _
            ByVal nMontoDisp As Double, ByVal nIntGanado As Double, ByVal dUltMov As Date, _
            ByVal sMovNro As String, Optional bActExtracto As Boolean = True)
Dim sSql As String

If bActExtracto Then
    sSql = "Update Producto Set nSaldo = nSaldo + " & nMontoCnt & ", nTransacc = nTransacc + 1 " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
Else
    sSql = "Update Producto Set nSaldo = nSaldo + " & nMontoCnt & " " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
End If
coConex.ConexionActiva.Execute sSql
sSql = "Update Captaciones Set nSaldoDisp = nSaldoDisp + " & nMontoDisp & ", nIntAcum = nIntAcum + " & nIntGanado _
    & ", dUltCierre = '" & Format$(dUltMov, "mm/dd/yyyy") & "', cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
coConex.ConexionActiva.Execute sSql

End Sub

''*** PEAC 2009509
Public Sub AgregaCuentaDocumento(ByVal sCuenta As String, ByVal nTpoDoc As TpoDoc, _
        ByVal sNroDoc As String, ByVal sCodIF As String, ByVal sMovNro As String, _
        ByVal nMovNro As Long, Optional nMonto As Double = 0, _
        Optional nEstadoOP As CaptacOrdPagoEstado = gCapOPEstCobrada)
Dim sFecha As String
Dim sSql As String
sFecha = Mid(sMovNro, 1, 4) & "/" & Mid(sMovNro, 5, 2) & "/" & Mid(sMovNro, 7, 2) & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)
Select Case nTpoDoc
    Case TpoDocCheque
        sSql = "INSERT DocRecCapta (nTpoDoc,cNroDoc,cPersCod,cIFTpo,nMovNro,cCtaCod,nMonto) " _
            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCodIF & "','" & Format$(gTpoIFBanco, "00") & "'," & nMovNro & ",'" & sCuenta & "'," & nMonto & ")"
    Case TpoDocOrdenPago
        sSql = "INSERT DocRecOP (nTpoDoc,cNroDoc,cCtaCod,nMonto,cIFCodPers,nEstado,cMovNro) " _
            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCuenta & "'," & nMonto & ",'" & sCodIF & "'," & nEstadoOP & ",'" & sMovNro & "')"
    Case TpoDocNotaAbono, TpoDocNotaCargo
End Select
coConex.ConexionActiva.Execute sSql
sSql = "INSERT MovDoc (nMovNro,nDocTpo,cDocNro,dDocFecha) " _
    & "VALUES (" & nMovNro & "," & nTpoDoc & ",'" & sNroDoc & "','" & sFecha & "')"
coConex.ConexionActiva.Execute sSql
End Sub

''*** PEAC 20090409
Public Sub AgregaMovCap(ByVal nMovNro As Long, ByVal nOperacion As CaptacOperacion, _
        ByVal sCuenta As String, ByVal nMonto As Double, _
        ByVal nSaldoDisp As Double, ByVal nSaldoCnt As Double)
Dim sSql As String

sSql = "INSERT MovCap (nMovNro,cOpeCod,cCtaCod,nMonto,nSaldoDisponible,nSaldoContable) " _
    & "VALUES (" & nMovNro & ",'" & nOperacion & "','" & sCuenta & "'," & nMonto & "," & nSaldoDisp & "," & nSaldoCnt & ")"
coConex.ConexionActiva.Execute sSql
End Sub

''*** PEAC 20090427
Public Sub AgregaMovCapDet(ByVal nMovNro As Long, ByVal nOperacion As CaptacOperacion, _
        ByVal sCuenta As String, ByVal nConcepto As CaptacConcepto, ByVal nMonto As Double)
Dim sSql As String

sSql = "INSERT MovCapDet (nMovNro,cOpeCod,cCtaCod,nConceptoCod,nMonto) " _
    & "VALUES (" & nMovNro & ",'" & nOperacion & "','" & sCuenta & "'," & nConcepto & "," & nMonto & ")"
coConex.ConexionActiva.Execute sSql
End Sub

''*** PEAC 20090509
Public Sub ActualizaEstadoCuenta(ByVal sCuenta As String, ByVal nEstado As CaptacEstado, _
        ByVal dFecha As Date, ByVal sMovNro As String)
'Actualiza el ultimo estado a la tabla de producto
Dim sSql As String
sSql = "Update Producto Set nPrdEstado = " & nEstado & ", " _
    & "dPrdEstado = '" & Format$(dFecha, "mm/dd/yyyy") & "' WHERE " _
    & "cCtaCod = '" & sCuenta & "'"
coConex.ConexionActiva.Execute sSql
'Agrega un registro mas de estado a la historia de estados de la cuenta
sSql = "INSERT CaptacEstado (cCtaCod,nPrdEstado,cMovNro) " _
    & "VALUES ('" & sCuenta & "'," & nEstado & ",'" & sMovNro & "')"
coConex.ConexionActiva.Execute sSql
End Sub

''*** PEAC 20090509
Public Sub UltimaActualizacionCuenta(ByVal sCuenta As String, ByVal sMovNro As String)
'Actualiza la fecha de ultima actualizacion de la cuenta de captaciones
Dim sSql As String
sSql = "Update Captaciones Set cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
coConex.ConexionActiva.Execute sSql
End Sub

''*** PEAC 20090511
Public Sub ActualizaCargoCaptacion(ByVal sCuenta As String, ByVal nMontoCnt As Double, _
            ByVal nMontoDisp As Double, ByVal nIntGanado As Double, ByVal dUltMov As Date, _
            ByVal sMovNro As String, Optional bActExtracto As Boolean = True)

Dim sSql As String

If bActExtracto Then
    sSql = "Update Producto Set nSaldo = nSaldo - " & nMontoCnt & ", nTransacc = nTransacc + 1 " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
Else
    sSql = "Update Producto Set nSaldo = nSaldo - " & nMontoCnt & " " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
End If
coConex.ConexionActiva.Execute sSql
sSql = "Update Captaciones Set nSaldoDisp = nSaldoDisp - " & nMontoDisp & ", nIntAcum = nIntAcum + " & nIntGanado _
    & ", dUltCierre = '" & Format$(dUltMov, "mm/dd/yyyy") & "', cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
coConex.ConexionActiva.Execute sSql
End Sub

''*** PEAC 20090515
Public Sub AgregaNuevoProdPers(ByVal sCuenta As String, ByVal sPersona As String, ByVal nRelacion As CaptacRelacPersona)
Dim sSql As String
sSql = "INSERT ProductoPersona (cCtaCod,cPersCod,nPrdPersRelac) " _
    & "VALUES ('" & sCuenta & "','" & sPersona & "'," & nRelacion & ")"
coConex.ConexionActiva.Execute sSql
End Sub

'*** PEAC 20090609
Public Function CapExtornoAbonoAho(ByVal sMovNroCap As String, ByVal pnMovNro As Long, ByVal nMovNroBus As Long, ByVal nOperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
            Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional sNomAge As String, Optional sLpt As String = "") As Double

Dim rsCta As ADODB.Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean
Dim nDiasTranscurridos As Long, i As Long
Dim nExtracto As Long, nMovNro As Long
Dim sMsgOpe As String, sCodUser As String
Dim dFecSis As Date
'Dim oFun As COMNContabilidad.NCOMContFunciones

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False
'Obtiene los datos para el calculo

Set rsCta = GetDatosCuentaAho(sCuenta)

nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
bInactiva = IIf(rsCta("bInactiva") = 0, False, True)

nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing
'Inicia la transaccion


bTrans = True

'Randomize
'For i = 0 To Rnd(2000) * 1000
'Next i

'If Not ValidaSaldoCuenta(sCuenta, nMonto) Then
'
'    Err.Raise 1000, "CapExtornoAbonoAho", "Cuenta NO Posee Saldo Suficiente"
'    CapExtornoAbonoAho = 0
'    Exit Function
'End If
sGlosa = sGlosa & "Extorno Abono Cuenta = " & sCuenta
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
dUltMov = DateAdd("d", -1, dFecSis)
nSaldoCnt = nSaldoCnt - nMonto

bTrans = True

ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp

If sNroDoc <> "" Then
    If nTipoDoc = TpoDocCheque Then
        nSaldoDisp = nSaldoDisp - 0
        ActualizaCargoCaptacion sCuenta, nMonto, 0, nIntGanado, dUltMov, sMovNro, True
        ActualizaEstadoDocRecEst TpoDocCheque, sNroDoc, sCodIF, sMovNro, gsChqEstExtornado
        sMsgOpe = "Ext. Depósito Chq."
    ElseIf nTipoDoc = TpoDocNotaCargo Then
        nSaldoDisp = nSaldoDisp - nMonto
        sMsgOpe = "Ext. Depósito NC"
    End If
Else
    nSaldoDisp = nSaldoDisp - nMonto
    ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True

End If
ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
nMovNro = pnMovNro
If nOperacion <> gAhoExtDepPagFocmacm Then 'FRHU 20150520 ERS022-2015
    AgregaMov sMovNroCap, nOperacion, sGlosa, gMovEstContabMovContable, gMovFlagDeExtorno
    nMovNro = GetnMovNro(sMovNroCap)
End If
AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
'ALPA20130827**********************************************************
'dInsertMovRef nMovNro, pnMovNro, False
dInsertMovRef nMovNro, IIf(pnMovNro = 0, nMovNroBus, pnMovNro), False
'**********************************************************************
UltimaActualizacionCuenta sCuenta, sMovNro

CapExtornoAbonoAho = nSaldoCnt

Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sModDep = sMsgOpe
sTipApe = "EXTORNO AHORROS"

End Function

'*** PEAC 20090609
Public Function CapExtornoCargoAho(ByVal pnMovNro As Long, ByVal nMovNroBus As Long, ByVal nOperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
            Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional sNomAge As String = "", Optional sLpt As String = "") As Double

Dim rsCta As ADODB.Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean
Dim nNumExtracto As Long, nDiasTranscurridos As Long, nExtracto As Long
Dim nMovNro As Long
Dim sMsgOpe As String, sCodUser As String
Dim dFecSis As Date
dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False
'Obtiene los datos para el calculo
Set rsCta = GetDatosCuentaAho(sCuenta)

nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
bInactiva = IIf(rsCta("bInactiva") = 1, True, False)
nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing

'Inicia la transaccion

bTrans = True
If sGlosa = "" Then sGlosa = "Extorno Cargo Cuenta = " & sCuenta
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
dUltMov = DateAdd("d", -1, dFecSis)
nSaldoCnt = nSaldoCnt + nMonto
nSaldoDisp = nSaldoDisp + nMonto
ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
ActualizaAbonoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
If sNroDoc = "" Then
    sMsgOpe = "Ext. Retiro Efectivo"
Else
    If nTipoDoc = TpoDocOrdenPago Then
        If nOperacion = gAhoExtRetOPCanje Then
            sMsgOpe = "Ext. Retiro OP Canje"
        Else
            sMsgOpe = "Ext. Retiro OP"
        End If
        AgregaOrdenPagoEstado sCuenta, sNroDoc, sMovNro, nMonto, gCapOPEstExtornada
    ElseIf nTipoDoc = TpoDocNotaCargo Then
        sMsgOpe = " Ext. Retiro NC"
    End If
End If
ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
AgregaMov sMovNro, nOperacion, sGlosa, , gMovFlagDeExtorno
nMovNro = GetnMovNro(sMovNro)
'nMovNro = pnMovNro
AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
UltimaActualizacionCuenta sCuenta, sMovNro

CapExtornoCargoAho = nSaldoCnt

End Function

'*** PEAC 20090609
Public Sub ActualizaEstadoDocRecEst(ByVal nTipoDoc As TpoDoc, ByVal sNroDoc As String, _
        ByVal sCodIF As String, ByVal sMovNro As String, ByVal nEstado As ChequeEstado)
'Actualiza el ultimo estado a la tabla de Documento Recibidos
Dim sSql As String
sSql = "Update DocRec Set nEstado = " & nEstado & " WHERE " _
    & "nTpoDoc = '" & nTipoDoc & "' AND cNroDoc = '" & sNroDoc & "' AND cPersCod = '" & sCodIF & "'"
coConex.ConexionActiva.Execute sSql
'Agrega un registro mas de estado a la historia de estados del documento recibido
sSql = "INSERT DocRecEst (nTpoDoc,cNroDoc,cPersCod,cMovNro,nEstado) " _
    & "VALUES ('" & nTipoDoc & "','" & sNroDoc & "','" & sCodIF & "','" & sMovNro & "'," & nEstado & ")"
coConex.ConexionActiva.Execute sSql
End Sub

'*** `PEAC 20090609
Public Sub AgregaMov(ByVal sMovNro As String, ByVal nOperacion As CaptacOperacion, _
        ByVal sGlosa As String, Optional nMovEstado As MovEstado = gMovEstContabMovContable, _
        Optional nMovFlag As MovFlag = gMovFlagVigente)
Dim sSql As String
sSql = "INSERT Mov (cMovNro,cOpeCod,cMovDesc,nMovEstado,nMovFlag) " _
    & "VALUES ('" & sMovNro & "','" & nOperacion & "','" & sGlosa & "'," & nMovEstado & "," & nMovFlag & ")"
coConex.ConexionActiva.Execute sSql
End Sub

'*** PEAC 20090609
Public Sub ActualizaEstadoMov(ByVal nMovNro As Long, ByVal nMovFlag As MovFlag)
Dim sSql As String
sSql = "Update Mov Set nMovFlag = " & nMovFlag & " Where nMovNro = " & nMovNro
coConex.ConexionActiva.Execute sSql
End Sub

'*** PEAC 20090609
Public Sub AgregaOrdenPagoEstado(ByVal sCuenta As String, ByVal sNroDoc As String, _
        ByVal sMovNro As String, ByVal nMonto As Double, ByVal nEstadoOP As CaptacOrdPagoEstado)
Dim sSql As String
sSql = "UPDATE DocRecOP Set nMonto = " & nMonto & " WHERE nTpoDoc = " & TpoDocOrdenPago & " " _
    & "AND cNroDoc = '" & sNroDoc & "' And cCtaCod = '" & sCuenta & "'"
coConex.ConexionActiva.Execute sSql
sSql = "INSERT DocRecOPEst (nTpoDoc,cNroDoc,cCtaCod,cMovNro,nMonto,nEstado) " _
    & "VALUES (" & TpoDocOrdenPago & ",'" & sNroDoc & "','" & sCuenta & "','" & sMovNro & "'," & nMonto & "," & nEstadoOP & ")"
coConex.ConexionActiva.Execute sSql
End Sub

'*** MAVM 20100717
Public Sub dInsertarJoyaGarantia(ByVal psJoyGarCod As String, ByVal psPersCod As String, _
        ByVal pnTasacion As Currency, ByVal pnPiezas As Integer, ByVal pnPesoBruto As Currency, _
        ByVal pnPesoNeto As Currency, ByVal psMovNro As String, ByVal prJoyas As ADODB.Recordset, _
        Optional ByVal pbEjecBatch As Boolean = False)
Dim lsSQL As String

lsSQL = "INSERT JoyaGarantia (cJoyGarCod, cPersCod, nPiezas, nPesoBruto, nPesoNeto, nTasacion, cMovNro, nJoyGarEstado, cMovNroRes) " _
    & "VALUES ('" & psJoyGarCod & "','" & psPersCod & "'," & pnPiezas & "," & pnPesoBruto & "," & pnPesoNeto & "," & pnTasacion & ",'" & psMovNro & "', 0, '')"
If pbEjecBatch Then
    coConex.AdicionaCmdBatch lsSQL
Else
    coConex.Ejecutar lsSQL
End If

Do While Not prJoyas.EOF
    lsSQL = "INSERT JoyaGarantiaDet (cJoyGarCod, cKilataje, nItem, nPiezas, " _
          & "nPesoBruto, nPesoNeto, nTasacion, cDescripcion ) " _
          & "VALUES ('" & psJoyGarCod & "','" & Trim(Str(Right(prJoyas("Material"), 2))) & "'," _
          & prJoyas.AbsolutePosition & "," & prJoyas("Pzas") & "," _
          & prJoyas("PBruto") & "," & prJoyas("PNeto") & "," & prJoyas("Tasac") & "," _
          & "'" & prJoyas("Descripcion") & "' )"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    prJoyas.MoveNext
Loop
End Sub

Public Sub dUpdateJoyaGarantia(ByVal psJoyGarCod As String, ByVal pnEstado As Integer, Optional psMovNroRes As String = "", Optional ByVal pnTransac As Integer = -1, _
        Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String
    lsSQL = "UPDATE JoyaGarantia SET "
    lsSQL = lsSQL & " nJoyGarEstado = " & pnEstado & ""
    If psMovNroRes <> "" Then
    lsSQL = lsSQL & " , cMovNroRes = '" & psMovNroRes & "'"
    End If
    lsSQL = lsSQL & " WHERE cJoyGarCod ='" & psJoyGarCod & "' "
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

'*** PEAC 20120709
Public Sub InsertColocRecupVisitaGestores(ByVal pnMovNro As Long, ByVal psCtaCod As String, ByVal psFechaVisita As String, ByVal pnPersContac As Integer, _
        ByVal psFecCompromiso As String, ByVal pnResulGestion As Integer, ByVal psZona As String, ByVal pnLugarContac As Integer, _
        ByVal pnMontoCompromiso As Double, ByVal pnCondicion As Integer, ByVal psComentario As String)

Dim lsSQL As String
       
    lsSQL = "exec stp_ins_VisitaGestores " & pnMovNro & ",'" & psCtaCod & "','" & psFechaVisita & "'," & pnPersContac & ",'" & psFecCompromiso & "'," & pnResulGestion & ",'" & psZona & "'," & pnLugarContac & "," & pnMontoCompromiso & "," & pnCondicion & ",'" & psComentario & "'"
    
    coConex.Ejecutar lsSQL
End Sub

'******RECO 2013-07-19******
'//Registra una nueva autorizacion para un cuenta aplicando pago por judicial
Public Sub dInsertColocRecupJudAutorizado(cCtaCod As String, cMovnro As String, dFecRegistro As String, cMetLiquid As String, nEstado As Integer, cMetLiquidAut As String)
    Dim lsSQL As String
        lsSQL = "UPDATE ColocRecupJudAutorizado SET nEstado = 0 WHERE cCtaCod = '" & cCtaCod & "'"
        lsSQL = lsSQL & "INSERT INTO ColocRecupJudAutorizado (cCtaCod,cMovNro ,dFecRegistro,cMetLiquid, nEstado ,cMetLiquidAut) VALUES ('" _
        & cCtaCod & "','" & cMovnro & "','" & dFecRegistro & "','" & cMetLiquid & "'," & nEstado & ",'" & cMetLiquidAut & "')"
        coConex.Ejecutar lsSQL
End Sub
Public Function dValidaDiaAutorizacionPagoJud(ByVal cCtaCod As String, ByVal cFecSist As String) As ADODB.Recordset
    Dim sSql As String
    Dim rsCta As ADODB.Recordset
    
    sSql = "SELECT cCtaCod ,dFecRegistro  FROM ColocRecupJudAutorizado WHERE cCtaCod = '" & cCtaCod & "' AND nEstado = 1"
    Set rsCta = New ADODB.Recordset
    rsCta.CursorLocation = adUseClient
    rsCta.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    Set rsCta.ActiveConnection = Nothing
    Set dValidaDiaAutorizacionPagoJud = rsCta
    Set rsCta = Nothing
End Function
Public Function dActualizarMetLiquidPagoJud(ByVal cCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim rsCta As ADODB.Recordset
    
    sSql = "UPDATE ColocRecupJudAutorizado SET nEstado = 0 WHERE cCtaCod = '" & cCtaCod & "'"
    Set rsCta = New ADODB.Recordset
    rsCta.CursorLocation = adUseClient
    rsCta.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    Set rsCta.ActiveConnection = Nothing
    Set rsCta = Nothing
End Function
'**********END RECO****************
'WIOR 20140128 **********************************************************
Public Sub dDeleteCredVinculados(ByVal psCtaCod As String, ByVal pnEstado As Integer)
Dim lsSQL As String
On Error GoTo error
    
    lsSQL = "EXEC stp_del_CredAsignacionSaldo '" & psCtaCod & "'," & pnEstado
    
    coConex.Ejecutar lsSQL

    Exit Sub
error:
    Err.Raise Err.Number, "Error en Proceso (dDeleteColocConvBcoNac)", Err.Description
End Sub
Public Sub dInsertCredVinculados(ByVal psCtaCod As String, ByVal pnMonto As Double, ByVal psMovNroSol As String, ByVal pnTipoVinc As Integer, ByVal pnEstado As Integer)
Dim lsSQL As String
On Error GoTo error
    
    lsSQL = "EXEC stp_ins_CredAsignacionSaldo '" & psCtaCod & "'," & pnMonto & ",'" & psMovNroSol & "'," & pnTipoVinc & "," & pnEstado
    
    coConex.Ejecutar lsSQL

    Exit Sub
error:
    Err.Raise Err.Number, "Error en proceso (dInsertCredVinculados)", Err.Description
End Sub
Public Sub ActualizaCredVinculados(ByVal psCtaCod As String, ByVal psMovAprueba As String, ByVal pnEstadoA As Integer, ByVal pnEstadoB As Integer)
Dim lsSQL As String
On Error GoTo error

    lsSQL = "EXEC stp_upd_CredAsignacionSaldo '" & psCtaCod & "','" & psMovAprueba & "'," & pnEstadoA & "," & pnEstadoB
    
    coConex.Ejecutar lsSQL

    Exit Sub
error:
    Err.Raise Err.Number, "Error en Proceso(ActualizaCredVinculados)", Err.Description
End Sub

'WIOR FIN ******************************************************************
'RECO20140208 ERS002********************************************************
Public Function ObtieneCtaRescateXAmpliacion(ByVal pnMovNro As Long) As ADODB.Recordset
    Dim sSql As String
    Dim rsCta As ADODB.Recordset
    
    sSql = "stp_sel_ObtieneCtaRescateXAmpliacion " & pnMovNro
    Set rsCta = New ADODB.Recordset
    rsCta.CursorLocation = adUseClient
    rsCta.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    Set ObtieneCtaRescateXAmpliacion = rsCta
    Set rsCta.ActiveConnection = Nothing
    Set rsCta = Nothing
End Function

Public Function dRecuperaDatosExtornoCancCredPig(ByVal psCtaCod As String, ByVal pnMovNroAnt As Long) As ADODB.Recordset
    Dim sSql As String
    Dim rsCta As ADODB.Recordset
    
    sSql = "exec stp_sel_RecuperaDatosExtornoCancCredPig2 '" & psCtaCod & "'," & pnMovNroAnt
    Set rsCta = New ADODB.Recordset
    rsCta.CursorLocation = adUseClient
    rsCta.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    Set dRecuperaDatosExtornoCancCredPig = rsCta
    Set rsCta.ActiveConnection = Nothing
    Set rsCta = Nothing
End Function
'RECO FIN*******************************************************************
'EJVG20140303 ***
Public Sub dInsertaMovDocRec(ByVal pnMovNro As Long, ByVal pnDocTpo As TpoDoc, ByVal psNroDoc As String, ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psIFCta As String)
    On Error GoTo InsertaMovDocErr
    Dim sql As String
    sql = "Exec stp_ins_ERS1262013_MovDocRec " & pnMovNro & "," & pnDocTpo & ",'" & psNroDoc & "','" & psPersCod & "','" & Format(psIFTpo, "00") & "','" & psIFCta & "'"
    coConex.Ejecutar sql
    Exit Sub
InsertaMovDocErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovDocRec Method")
End Sub
Public Sub dInsertaCuentaDocumento(ByVal pnMovNro As Long, ByVal pnTpoDoc As Integer, ByVal psNroDoc As String, ByVal psPersCod As String, ByVal psIFTpo As String, _
                                    ByVal psIFCta As String, ByVal psCtaCod As String, ByVal pnMonto As Currency)
    Dim lsSQL As String
    On Error GoTo ErrordInsertaCuentaDocumento
    lsSQL = "Exec stp_ins_ERS1262013_DocRecColoc " & pnTpoDoc & ",'" & psNroDoc & "','" & psPersCod & "','" & Format(psIFTpo, "00") & "','" & psIFCta & "'," & pnMovNro & ",'" & psCtaCod & "'," & pnMonto
    coConex.Ejecutar lsSQL
    Exit Sub
ErrordInsertaCuentaDocumento:
    Err.Raise Err.Number, "Error En Proceso dInsertaCuentaDocumento", Err.Description
End Sub
'END EJVG *******

'RIRO20140610 ERS017 *************************
Public Sub actualizarRelacionExtornoVoucherCaptacion(ByVal pnMovNroOpe As Long)
    Dim lsSQL As String
    lsSQL = "exec stp_upd_ActualizarCapVoucherDeposito_3  " & pnMovNroOpe & ""
    coConex.Ejecutar lsSQL
End Sub
Public Sub ActualizaMovPendientesOpe(ByVal pnMovNroOpe As Long, ByVal pnMonto As Currency)

    Dim lsSQL As String
    lsSQL = "exec STP_UPD_ACTUALIZAPENDOPE  " & pnMovNroOpe & ", " & pnMonto
    coConex.Ejecutar lsSQL
    
End Sub
'END RIRO ************************************
'RECO20140722 ERS114-2014************************************************
Public Function DevuelveDistritoXProvincia(ByVal psProvinciaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim rsDis As ADODB.Recordset
    
    sSql = "exec stp_sel_DevuelveDistritoXProvincia '" & psProvinciaCod & "'"
    Set rsDis = New ADODB.Recordset
    rsDis.CursorLocation = adUseClient
    rsDis.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    Set DevuelveDistritoXProvincia = rsDis
    Set rsDis.ActiveConnection = Nothing
    Set rsDis = Nothing
End Function
Public Sub ActualizaTarifarioCartanotariaMinka(ByVal psUbiGeoCod As String, ByVal pnValor As Currency)
    Dim lsSQL As String
    lsSQL = "exec stp_upd_ActualizaTarifarioCartanotariaMinka  '" & psUbiGeoCod & "', " & pnValor
    coConex.Ejecutar lsSQL
End Sub
Public Function DevuelveValorNotificacionCarNotMinka(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim rsDatos As ADODB.Recordset
    
    sSql = "exec stp_sel_DevuelveValorNotificacionCarNotMinka '" & psCtaCod & "'"
    Set rsDatos = New ADODB.Recordset
    rsDatos.CursorLocation = adUseClient
    rsDatos.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    Set DevuelveValorNotificacionCarNotMinka = rsDatos
    Set rsDatos.ActiveConnection = Nothing
    Set rsDatos = Nothing
End Function

'RECO FIN****************************************************************
'FRHU 20150415 ERS022-2015
Public Sub dAgregaMovOpeVarias(ByVal pnMovNro As Long, ByVal pnMovImporte As Double, _
        ByVal psNroDoc As String, ByVal psReferencia As String, ByVal pnMoneda As Moneda, _
        Optional nOperacion As CaptacOperacion, Optional ByVal pbEjecBatch As Boolean = False)
    Dim sql As String
    
    sql = " Insert MovOpeVarias (nMovNro,nMovImporte,cNroDoc,cReferencia,nMoneda,cOpeCod)" _
        & " Values(" & pnMovNro & "," & pnMovImporte & ",'" & psNroDoc & "','" & psReferencia & "'," & pnMoneda & ",'" & Trim(nOperacion) & "')"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sql
    Else
        coConex.Ejecutar sql
    End If
End Sub
Public Sub dUpdateColocTransferido(ByVal psCtaCod As String, Optional ByVal psFecIngRecup As String = "@", _
        Optional ByVal pnComisionCod As Integer = "-1", Optional ByVal pnSaldoIntComp As Currency = -1, _
        Optional ByVal pnSaldoIntMor As Currency = -1, Optional ByVal pnSaldoGasto As Currency = -1, _
        Optional ByVal psMetLiquid As String = "@", Optional ByVal pnNroCalend As Integer = -1, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

Dim lsSQL1 As String
Dim nVistoMetodo As Integer
Dim nVisNegociacion As Integer
Dim nVisCancelacion As Integer
Dim rs As New ADODB.Recordset

    lsSQL = "UPDATE ColocTransferido SET "

    If psFecIngRecup <> "@" Then
        lsSQL = lsSQL & " dIngRecup = '" & psFecIngRecup & "',"
    End If
    
    If pnSaldoIntComp = -4 Then ' Si se le envia (-4) disminuye la Cantidad
        lsSQL = lsSQL & " nIntComp  = nIntComp - " & pnSaldoIntComp & ","
    ElseIf pnSaldoIntComp = -3 Then ' Si se le envia (-3) aumenta la Cantidad
        lsSQL = lsSQL & " nIntComp  = nIntComp + " & pnSaldoIntComp & ","
    ElseIf pnSaldoIntComp <> -1 Then
        lsSQL = lsSQL & " nIntComp = " & pnSaldoIntComp & ","
    End If
    
    If pnSaldoIntMor = -4 Then ' Si se le envia (-4) disminuye la Cantidad
        lsSQL = lsSQL & " nIntMora = nIntMora - " & pnSaldoIntMor & ","
    ElseIf pnSaldoIntMor = -3 Then ' Si se le envia (-3) aumenta la Cantidad
        lsSQL = lsSQL & " nIntMora = nIntMora + " & pnSaldoIntMor & ","
    ElseIf pnSaldoIntMor <> -1 Then
        lsSQL = lsSQL & " nIntMora = " & pnSaldoIntMor & ","
    End If
    
    If pnSaldoGasto = -4 Then ' Si se le envia (-4) disminuye la cantidad
        lsSQL = lsSQL & " nGasto = nGasto - " & pnSaldoGasto & ","
    ElseIf pnSaldoGasto = -3 Then ' Si se le envia (-3) aumenta la cantidad
        lsSQL = lsSQL & " nGasto = nGasto + " & pnSaldoGasto & ","
    ElseIf pnSaldoGasto <> -1 Then
        lsSQL = lsSQL & " nGasto = " & pnSaldoGasto & ","
    End If
    
    If psMetLiquid <> "@" Then
        lsSQL = lsSQL & " cMetLiquid = '" & psMetLiquid & "',"
    End If
    If pnNroCalend <> -1 Then
        lsSQL = lsSQL & " nNroCalen = " & pnNroCalend & ","
    End If
      
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    
    lsSQL = lsSQL & " WHERE cCtaCod ='" & psCtaCod & "'"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub
Public Sub dInsertColocTransferidoGastos(ByVal psCtaCod As String, ByVal pnNroGastoCta As Integer, _
        ByVal pnPrdConceptoCod As Integer, _
        ByVal psFecAsigna As String, ByVal pnMonto As Currency, ByVal pnMontoPagado As Currency, _
        ByVal pnColocRecGastoEstado As ColocRecGastoEstado, ByVal psMotivoGasto As String, _
        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    
    lsSQL = "INSERT ColocTransferidoGastos (cCtaCod,nNroGastoCta, nPrdConceptoCod, dAsigna,nMonto, " _
        & " nMontoPagado, nColocTransfGastoEstado,cMotivoGasto ) " _
        & " VALUES ('" & psCtaCod & "'," & pnNroGastoCta & "," & pnPrdConceptoCod & ",'" & psFecAsigna & "'," & pnMonto & "," _
        & pnMontoPagado & "," & pnColocRecGastoEstado & ",'" & psMotivoGasto & "'  )"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub
Public Sub dInsertColocTransferidoDet(ByVal pnMovNro As Long, ByVal pcCtaCod As String, _
        ByVal pnSaldoCapital As Currency, ByVal pnIntComp As Currency, ByVal pnIntMora As Currency, ByVal pnGasto As Currency, _
        ByVal pnSaldoCapitalNew As Currency, ByVal pnIntCompNew As Currency, ByVal pnIntMoraNew As Currency, ByVal pnGastoNew As Currency, _
        Optional ByVal pbEjecBatch As Boolean = False)
        'FRHU 20150701: Incidente, Se modifico los tipos de datos de los parametros pnSaldoCapital(de Integer a Currency) y pnIntComp(de string a Currency)
Dim lsSQL As String
    
    lsSQL = "EXEC stp_ins_ColocTransferidoDet " & pnMovNro & ",'" & pcCtaCod & "'," & pnSaldoCapital & "," & pnIntComp & "," & pnIntMora & "," & pnGasto & "," _
            & pnSaldoCapitalNew & "," & pnIntCompNew & "," & pnIntMoraNew & "," & pnGastoNew

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub
Public Sub dUpdateColocTransferidoDet(ByVal pnMovNro As Long, Optional ByVal pbEjecBatch As Boolean = False)
    Dim lsSQL As String
    
    lsSQL = "EXEC stp_upd_ColocTransferidoDet " & pnMovNro
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub
Public Sub dUpdateColocTransferidoGastos(ByVal psCtaCod As String, ByVal pnNroGastoCta As Integer, _
        Optional ByVal psFecAsigna As String = "@", Optional ByVal pnMonto As Currency = -1, Optional ByVal pnMontoPagado As Currency = -1, _
        Optional ByVal pnColocRecGastoEstado As ColocRecGastoEstado = gColRecGastoEstPendiente, Optional ByVal psMotivoGasto As String = "@", _
        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    
    lsSQL = "UPDATE ColocTransferidoGastos SET "

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
        lsSQL = lsSQL & " nColocTransfGastoEstado = " & pnColocRecGastoEstado & ","
    End If
    If psMotivoGasto <> "@" Then
        lsSQL = lsSQL & " cMotivoGasto = '" & psMotivoGasto & "',"
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cCtaCod = '" & psCtaCod & "' AND  nNroGastoCta = " & pnNroGastoCta & " "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub
Public Sub dInsertColocTransfDistribucionCobr(ByVal psCtaCod As String, pdFecha As Date, _
        ByVal pnCapital As Currency, ByVal pnIntComp As Currency, ByVal pnMora As Currency, _
        ByVal pnGasto As Currency, ByVal psUltimaActualizacion As String, Optional ByVal pnRegCanc As Integer)

    Dim lsSQL As String

    lsSQL = "EXEC stp_ins_ColocTransfDistribucionCobr '" & psCtaCod & "','" & Format(pdFecha, "yyyymmdd") & "'," _
            & pnCapital & "," & pnIntComp & "," & pnMora & "," & pnGasto & ",'" _
            & psUltimaActualizacion & "'," & pnRegCanc
    coConex.Ejecutar lsSQL
End Sub
Public Sub dDeleteColocTransfDistribucionCobr(ByVal psCtaCod As String, pdFecha As Date)
    Dim lsSQL As String
    
    lsSQL = "EXEC stp_del_ColocTransfDistribucionCobr '" & psCtaCod & "','" & Format(pdFecha, "yyyymmdd") & "'"
    coConex.Ejecutar lsSQL
End Sub
Public Sub dInsertColocTransfAutorizado(cCtaCod As String, cMovnro As String, dFecRegistro As String, cMetLiquid As String, nEstado As Integer, cMetLiquidAut As String, cGlosa As String)
    Dim lsSQL As String
    
    lsSQL = "EXEC stp_ins_upd_ColocTransfAutorizado '" & cCtaCod & "','" & cMovnro & "','" & Format(dFecRegistro, "yyyymmdd") & "','" & cMetLiquid & "'," _
                                                       & nEstado & ",'" & cMetLiquidAut & "','" & cGlosa & "'"
    coConex.Ejecutar lsSQL
End Sub
Public Function dActualizarColocTransfAutorizado(ByVal cCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim rsCta As ADODB.Recordset
    
    sSql = "EXEC stp_upd_ColocTransfAutrizado '" & cCtaCod & "'"
    Set rsCta = New ADODB.Recordset
    rsCta.CursorLocation = adUseClient
    rsCta.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    Set rsCta.ActiveConnection = Nothing
    Set rsCta = Nothing
End Function
'FIN FRHU 20150415
'EJVG20150822 ***
Public Sub dLiberarGarantia(pnMovNro As Long, ByVal psOpeTpo As String, ByVal psCtaCod As String, ByVal pnMontoCapAmortiza As Currency, ByVal pnEstadoCred As ColocEstado, Optional pbEjecBatch As Boolean = False)
    Dim lsSQL As String
    On Error GoTo ErrLiberarGarantia
    lsSQL = "EXEC spt_ins_ERS0632014_LiberaGarantia " & pnMovNro & ",'" & psOpeTpo & "','" & psCtaCod & "'," & pnMontoCapAmortiza & "," & pnEstadoCred
    coConex.Ejecutar lsSQL
    Exit Sub
ErrLiberarGarantia:
    Err.Raise Err.Number, "Error al liberar las garantías en el proceso de Pago de Créditos", Err.Description
End Sub
Public Sub dExtornoLiberarGarantia(pnMovNro As Long, ByVal psCtaCod As String)
    Dim lsSQL As String
    On Error GoTo ErrLiberarGarantia
    lsSQL = "EXEC spt_ins_ERS0632014_ExtornoLiberaGarantia " & pnMovNro & ",'" & psCtaCod & "'"
    coConex.Ejecutar lsSQL
    Exit Sub
ErrLiberarGarantia:
    Err.Raise Err.Number, "Error al liberar las garantías en el proceso de Pago de Créditos", Err.Description
End Sub
'END EJVG *******

'*** GIPO 11-10-2017 ERS048-2017
Public Function InsertMovSobrantes(ByVal cMovNroAdjud As String, ByVal cOpeCod As String, ByVal cMovDescrip As String, _
    ByVal nMovEstado As Integer) As Long
    On Error GoTo GeneraMovErr
    Dim rs As ADODB.Recordset
    Dim oConect As COMConecta.DCOMConecta
    Dim sql As String
    Set oConect = New COMConecta.DCOMConecta
    Set rs = New ADODB.Recordset
    If oConect.AbreConexion = False Then Exit Function
    
    sql = "stp_ins_ERS0482017_MovSobrantes '" & cMovNroAdjud & "','" & cOpeCod & "','" & cMovDescrip & "'," & nMovEstado
   
    Set rs = oConect.Ejecutar(sql)
    If Not rs.EOF Then
        InsertMovSobrantes = rs.Fields(0)
    End If
    rs.Close
    Set rs = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Function
GeneraMovErr:
    Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones:GeneraMovNro Method")
End Function
Public Sub RegistrarSobrantesPorCuenta(ByVal pcCta As String, ByVal pnMovNro, ByVal psNroAdju As String, ByVal psDeuda As Currency)
    Dim lsSQL As String
    lsSQL = "exec stp_ins_ERS0482017_SobrantePorCuenta '" & pcCta & "'," & pnMovNro & "," & psDeuda & ",'" & psNroAdju & "'"
    coConex.Ejecutar lsSQL
End Sub

'CTI2 20181116 INI ******
Public Sub ActualizarRelacionVoucherCaptacion(ByVal pnMovNroOpe As Long, _
                                              ByVal pnMovNroRVD As Long)
    Dim lsSQL As String
    lsSQL = "exec stp_upd_ActualizarCapVoucherDeposito_2  " & pnMovNroOpe & ", " & pnMovNroRVD & ""
    coConex.CargaRecordSet lsSQL
    Exit Sub
End Sub
Public Sub ActualizaMovPendientesRend(ByVal nMovNroPend As Long, ByVal pnMonto As Currency)
    Dim lsSQL As String
        
    lsSQL = " UPDATE MovPendientesRend " _
          & " Set nSaldo = nSaldo - " & pnMonto & "" _
          & " Where nMovNro = " & nMovNroPend & ""
    
    coConex.CargaRecordSet lsSQL
    Exit Sub
End Sub
'CTI2 20181116 END ******
'*** CTI3 (ferimoro) 16102018
Public Sub dInsertDetExtMov(ByVal pnMovNro As Long, ByVal psOperacion As String, _
        ByVal pnEstado As Integer, Optional psCuenta As String, _
        Optional pbEjecBatch As Boolean = False, _
        Optional ByVal psMotExt As String, _
        Optional ByVal psDetMotExt As String)

Dim lsSQL As String
Dim lrs As ADODB.Recordset
Set lrs = New ADODB.Recordset
 
    lsSQL = "INSERT MovDetExtornos (nMovNro,cOpedCod,cCtaCod,cMotExtorno,cDetMotExto) " _
          & "VALUES (" & pnMovNro & ",'" & psOperacion & "','" & psCuenta & "','" & psMotExt & "','" & UCase(psDetMotExt) & "')"

    Set lrs = coConex.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrs = Nothing

End Sub
'*** GIPO 02-01-2019 ERS001-2018
Public Function InsertCodigosJoyasAdjud(ByVal psNroProceso As String, ByVal nMovNroAdjud As Long) As Long
    On Error GoTo GeneraMovErr
    Dim rs As ADODB.Recordset
    Dim oConect As COMConecta.DCOMConecta
    Dim sql As String
    Set oConect = New COMConecta.DCOMConecta
    Set rs = New ADODB.Recordset
    If oConect.AbreConexion = False Then Exit Function

    sql = "stp_ins_ERS012018_InsertaCodigoJoyasAdjud '" & psNroProceso & "'," & nMovNroAdjud
   
    Set rs = oConect.Ejecutar(sql)
    If Not rs.EOF Then
        InsertCodigosJoyasAdjud = rs.Fields(0)
    End If
    rs.Close
    Set rs = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Function
GeneraMovErr:
    Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones:GeneraMovNro Method")
End Function
'CTI4 ERS0112020
Public Function CapCargoCuentaAho(ByRef pMatAho As Variant, ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nOperacion As CaptacOperacion, ByVal sMovNro As String, ByVal sGlosa As String, ByVal pdFecSis As Date, _
            Optional nTipoDoc As TpoDoc = TpoDocNotaCargo, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional bOPExiste As Boolean = False, _
            Optional bActivaCta As Boolean = True, Optional bNumExtracto As Boolean = True, _
            Optional sCodCMAC As String = "", Optional sNomCmac As String = "", _
            Optional sNomAge As String = "", Optional sLpt As String = "", _
            Optional sPersLavDinero As String = "", Optional bCancelaciones As Boolean = False, _
            Optional pnMonedaCred As Moneda = 0, _
            Optional pbITFAplica As Boolean = True, Optional pnITFValor As Double = 0, Optional pbITFAsumido As Boolean = False, _
            Optional psItfOperacion As String = "", Optional pbRetiroEfectivo As Boolean = False) As Double
            'FRHU RQ14008 Se agrego: pbRetiroEfectivo

Dim rsCta As ADODB.Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim nSaldoInac As Double
Dim dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean
Dim nNumExtracto As Long, nDiasTranscurridos As Long, i As Long
Dim nExtracto As Long, nMovNro As Long
Dim sMsgOpe As String, sCodUser As String
Dim dFecSis As Date
Dim lnMontoOri As Double
Dim oGen As COMDConstSistema.DCOMGeneral
Dim lnTipCambioV As Currency
Dim lnTipCambioC As Currency
Dim lnTipCambio As Currency

'Dim clsMov As COMNContabilidad.NCOMContFunciones

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False
'Obtiene los datos para el calculo

Set rsCta = GetDatosCuentaAho(sCuenta)

nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
bInactiva = IIf(rsCta("bInactiva") = 0, False, True)
nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing

Set oGen = New COMDConstSistema.DCOMGeneral

'GetTipCambio pdFecSis
lnTipCambioC = oGen.EmiteTipoCambio(pdFecSis, TCCompra)
lnTipCambioV = oGen.EmiteTipoCambio(pdFecSis, TCVenta)
lnTipCambio = oGen.EmiteTipoCambio(pdFecSis, TCFijoDia)

'RIRO 20200706 ****************
Dim sMensajeReactiva As String
sMensajeReactiva = oGen.LeeConstSistema("725")
If Not Len(Trim(sMensajeReactiva)) > 0 Then
    sMensajeReactiva = "Retiro Comisión Reactiva"
End If
'END RIRO *********************

If pnMonedaCred = 0 Then pnMonedaCred = Mid(sCuenta, 9, 1)

If pnMonedaCred <> Mid(sCuenta, 9, 1) Then
    lnMontoOri = nMonto
    If pnMonedaCred = gMonedaExtranjera And Mid(sCuenta, 9, 1) = gMonedaNacional Then
        nMonto = Round(nMonto * lnTipCambioV, 2)
    Else
        nMonto = Round(nMonto / lnTipCambioC, 2)
    End If
Else
    lnMontoOri = 0
End If

bTrans = True
Randomize
For i = 0 To Rnd(2000) * 1000
Next i

If Not ValidaSaldoCuenta(sCuenta, nMonto) Then 'Valida el saldo de la cuenta nuevamente
    Err.Raise 1000, "CapCargoCuentaAho", "Cuenta NO Posee Saldo Suficiente"
    CapCargoCuentaAho = 0
    Exit Function
End If
If bActivaCta Then
    If bInactiva Then
        nSaldoInac = nSaldoCnt
    End If
End If
'Calcula los intereses
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
If Not bCancelaciones Then
    pMatAho(3) = Format(nIntGanado, "#0.00")
Else
    pMatAho(4) = Format(nIntGanado, "#0.00")
End If

dUltMov = DateAdd("d", -1, dFecSis)
nSaldoCnt = nSaldoCnt - nMonto

ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp

nSaldoDisp = nSaldoDisp - nMonto
If Not bCancelaciones Then
    pMatAho(10) = nSaldoDisp
    pMatAho(11) = nSaldoCnt
Else
    pMatAho(12) = nSaldoDisp
    pMatAho(13) = nSaldoCnt
End If

'FRHU20140228 RQ14006
If pbRetiroEfectivo Then
    pMatAho(15) = nSaldoDisp
    pMatAho(16) = nSaldoCnt
    pMatAho(17) = Format(nIntGanado, "#0.00")
End If
'FIN FRHU20140228 RQ14006
ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True

If sNroDoc <> "" Then
    If nTipoDoc = TpoDocOrdenPago Then 'Si es con orden de pago
        If bOPExiste Then
            AgregaOrdenPagoEstado sCuenta, sNroDoc, sMovNro, nMonto, gCapOPEstCobrada
        Else
            AgregaCuentaDocumento sCuenta, nTipoDoc, sNroDoc, sCodIF, sMovNro, nMonto, gCapOPEstCobrada
        End If
        If nOperacion = gAhoRetOPCanje Then
            sMsgOpe = "Retiro OP Canje"
        ElseIf nOperacion = gAhoRetOP Then
            sMsgOpe = "Retiro OP"
        ElseIf nOperacion = gAhoOPCertificacion Then
            sMsgOpe = "Certificacion OP"
        
        'RIRO 20200630 ******
        ElseIf nOperacion = "100949" Then
            sMsgOpe = sMensajeReactiva
        'END RIRO ***********
            
        End If
    ElseIf nTipoDoc = TpoDocNotaCargo Then
        sMsgOpe = "Retiro NC"
    End If
Else
        'RIRO 20200630 ******
        If nOperacion = "100949" Then
            sMsgOpe = sMensajeReactiva
        'END RIRO ***********
        Else
            sMsgOpe = "Retiro Efectivo"
        End If
End If

'AgregaMov sMovNro, nOperacion, sGlosa

nMovNro = GetnMovNro(sMovNro)
AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
If bInactiva Then
    AgregaMovCap nMovNro, gAhoEstInacAct, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
    AgregaMovCapDet nMovNro, gAhoEstInacAct, sCuenta, gConcCapital, nSaldoInac
    ActualizaEstadoCuenta sCuenta, gCapEstActiva, dFecSis, sMovNro
    ActualizaInactivaAct (sCuenta) 'JUEZ 20150127
End If

'RIRO 20200611 Comisión Reactiva ****
If nOperacion = "100949" Then
    AgregaMovCapDet nMovNro, nOperacion, sCuenta, 1251, nMonto
Else
    AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
End If
'End RIRO 20200611 ******************

'AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto RIRO20200611 Comentado
'si el pago de credito es en diferente moneda a la cuenta de ahorros
Dim lnMontoCV As Double
If pnMonedaCred <> Mid(sCuenta, 9, 1) Then
    If Mid(sCuenta, 9, 1) = gMonedaExtranjera Then
        lnMontoCV = Round((lnMontoOri / lnTipCambioC) * lnTipCambio, 2)
        AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeCompra, lnMontoCV
        AgregaMovTipoCambio nMovNro, lnTipCambioC
        
        If lnTipCambioC < lnTipCambio Then
            AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - nMonto)
        Else
            AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - nMonto)
        End If
    Else
        lnMontoCV = Round(lnMontoOri * lnTipCambio, 2)
        AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeVenta, lnMontoCV
        AgregaMovTipoCambio nMovNro, lnTipCambioV
        
        If lnTipCambioV > lnTipCambio Then
            AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - Round(lnMontoOri * lnTipCambioV, 2))
        Else
            AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - Round(lnMontoOri * lnTipCambioV, 2))
        End If
    End If
End If
CapCargoCuentaAho = nSaldoCnt

If pbITFAplica And pnITFValor > 0 Then
    If pbITFAsumido Then
        AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp, nSaldoCnt
        AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFAsumido, pnITFValor
    Else
        ActualizaCargoCaptacion sCuenta, pnITFValor, pnITFValor, nIntGanado, dUltMov, sMovNro, bNumExtracto
        AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp - pnITFValor, nSaldoCnt - pnITFValor
        AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFCliente, pnITFValor
    End If
    
    nSaldoDisp = nSaldoDisp - pnITFValor
    If Not bCancelaciones Then
        pMatAho(10) = nSaldoDisp
        pMatAho(11) = nSaldoCnt - pnITFValor
    Else
        pMatAho(12) = nSaldoDisp
        pMatAho(13) = nSaldoCnt - pnITFValor
    End If
    'FRHU20140429 INCIDENTE: Se agrego para que en el vouche muestre el saldo correcto. cuando hace el retiro
    If pbRetiroEfectivo Then
        pMatAho(15) = nSaldoDisp
        pMatAho(16) = nSaldoCnt - pnITFValor
        pMatAho(17) = Format(nIntGanado, "#0.00")
    End If
    'FIN FRHU20140429
    CapCargoCuentaAho = nSaldoCnt - pnITFValor
End If
UltimaActualizacionCuenta sCuenta, sMovNro

bTrans = False
End Function
Public Function ValidaSaldoCuenta(ByVal sCuenta As String, ByVal nMonto As Double) As Boolean
Dim nMoneda As Moneda
Dim rsCta As ADODB.Recordset
Dim nSaldo As Double, nMontoMinimo As Double

Set rsCta = GetDatosCuenta(sCuenta)
nSaldo = rsCta("nSaldoDisp")
rsCta.Close
Set rsCta = Nothing

nMoneda = CLng(Mid(sCuenta, 9, 1))
If nMoneda = gMonedaNacional Then
    nMontoMinimo = GetCapParametro(gSaldMinAhoMN)
Else
    nMontoMinimo = GetCapParametro(gSaldMinAhoME)
End If
If nSaldo - nMontoMinimo - nMonto >= 0 Then
    ValidaSaldoCuenta = True
Else
    ValidaSaldoCuenta = False
End If
End Function
Public Sub AgregaMovTipoCambio(ByVal nMovNro As Long, ByVal pnTipoCambio As Currency)
Dim sSql As String

sSql = " INSERT MovTpoCambio (nMovNro, nMovTpoCambio) " _
     & " VALUES (" & nMovNro & "," & pnTipoCambio & ")"

coConex.ConexionActiva.Execute sSql
End Sub
Public Sub ActualizaInactivaAct(ByVal sCuenta As String)
Dim sSql As String
    sSql = "Update CaptacAhorros Set bInactiva = 0 WHERE cCtaCod = '" & sCuenta & "'"
    coConex.ConexionActiva.Execute sSql
End Sub
Public Function GetDatosCuenta(ByVal sCuenta As String) As ADODB.Recordset
Dim rsCta As ADODB.Recordset
Dim nProd As Producto
nProd = CInt(Mid(sCuenta, 6, 3))

Select Case nProd
    Case gCapAhorros
        Set rsCta = GetDatosCuentaAho(sCuenta)
        
    Case gCapPlazoFijo
        Set rsCta = GetDatosCuentaPF(sCuenta)
        
    Case gCapCTS
        Set rsCta = GetDatosCuentaCTS(sCuenta)
        
End Select
Set GetDatosCuenta = rsCta
Set rsCta = Nothing

End Function
Public Function GetDatosCuentaPF(ByVal sCuenta As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, P.dPrdEstado, " _
    & "C.nSaldoDisp, C.nPersoneria, C.nFirmas, C.nIntAcum, C.dUltCierre, C.dApertura, AG.cAgeDescripcion, " _
    & "C.nPrdCtaTpo, C.nPrdTasaInteres, A.nPlazo, A.nIntPag, A.dRenovacion, A.nApertura, " _
    & "A.nFormaRetiro, A.dAuxiliar, A.nDuplicado, T.cConsDescripcion cEstado, A.bBusCredPend, " _
    & "T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa, T3.cConsDescripcion cRetiro, " _
    & "A.dUltCierreAnt, A.dAuxiliarAnt, ISNULL(CI.cCtaCodAbono,'') cCtaCodAbono,a.nformaretiro,c.calias,c.nfirmasmin FROM Producto P " _
    & "INNER JOIN Captaciones C INNER JOIN CaptacPlazoFijo A LEFT JOIN CaptacCtaAboIntPF CI " _
    & "ON A.cCtaCod = CI.cCtaCod ON C.cCtaCod = A.cCtaCod ON " _
    & "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
    & "INNER JOIN Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN Constante T2 " _
    & "ON C.nPrdTasaInteres = T2.nConsValor INNER JOIN Constante T3 ON A.nFormaRetiro = T3.nConsValor " _
    & "INNER JOIN Agencias AG ON SUBSTRING(P.cCtaCod,4,2) = AG.cAgeCod WHERE P.cCtaCod = '" & sCuenta & "' AND " _
    & "T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
    & "AND T2.nConsCod = " & gCaptacTipoTasa & " AND T3.nConsCod = " & gCaptacPFFormaRetiro


Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set GetDatosCuentaPF = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function
Public Function GetDatosCuentaCTS(ByVal sCuenta As String) As ADODB.Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset

sSql = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, P.dPrdEstado, " _
    & "C.nSaldoDisp, C.nPersoneria, C.nFirmas,c.nfirmasmin ,C.nIntAcum, dUltCierre=( case when dbo.ultimodiames(dUltCierre)=1 then dUltCierre else dateadd(dd,1, dUltCierre) end )  , C.dApertura, AG.cAgeDescripcion, " _
    & "C.nPrdCtaTpo, C.nPrdTasaInteres, A.nSaldRetiro, A.nIntSaldo, PE.cPersNombre cInstitucion, " _
    & "T.cConsDescripcion cEstado, T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa, A.cCodInst,C.cAlias " _
    & "FROM Producto P INNER JOIN Captaciones C INNER JOIN CaptacCTS A INNER JOIN " _
    & "Persona PE ON A.cCodInst = PE.cPersCod ON C.cCtaCod = A.cCtaCod ON " _
    & "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
    & "INNER JOIN Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN Constante T2 " _
    & "ON C.nPrdTasaInteres = T2.nConsValor INNER JOIN Agencias AG ON SUBSTRING(P.cCtaCod,4,2) = AG.cAgeCod " _
    & "WHERE P.cCtaCod = '" & sCuenta & "' AND T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
    & "AND T2.nConsCod = " & gCaptacTipoTasa

Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetDatosCuentaCTS = rsCta
Set rsCta = Nothing

End Function
Public Function GetCapParametro(ByVal nParametro As CaptacParametro) As Double
Dim rsVar As ADODB.Recordset
Dim sSql As String
sSql = "SELECT nParValor FROM Parametro WHERE nParCod = " & nParametro & " " _
    & "And nParProd = " & gPrdParamCaptac
Set rsVar = New ADODB.Recordset
rsVar.CursorLocation = adUseClient
rsVar.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
Set rsVar.ActiveConnection = Nothing
If rsVar.EOF And rsVar.BOF Then
    GetCapParametro = 0
Else
    GetCapParametro = rsVar("nParValor")
End If
rsVar.Close
Set rsVar = Nothing
End Function
Public Function RecuperaMovimientoCapataciones(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String

'    sSql = "Select M.cMovNro, MC.cOpeCod, MC.cCtaCod, MC.nMonto from MovCap MC "
'    sSql = sSql & " Inner Join Mov M ON M.nMovNro = MC.nMovNro "
'    sSql = sSql & " Where MC.nMovNro = " & pnMovNro

    sSql = "exec stp_sel_RecuperaMovimientoCapataciones " & pnMovNro 'CTI4 ERS0112020
    
    Set RecuperaMovimientoCapataciones = New ADODB.Recordset
    
    RecuperaMovimientoCapataciones.CursorLocation = adUseClient
    RecuperaMovimientoCapataciones.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    RecuperaMovimientoCapataciones.ActiveConnection = Nothing

End Function
Public Function GetDatosCuentaAho(ByVal sCuenta As String) As ADODB.Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset

sSql = "Exec sp_sel_DatosCuentaAhorro '" & sCuenta & "' , " & gCaptacEstado & " , " & gProductoCuentaTipo & " , " & gCaptacTipoTasa

Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetDatosCuentaAho = rsCta
Set rsCta = Nothing
End Function
'END CTI4

'*** JAOR 20210223
Public Sub dUpdateColocPigDescInteres(ByVal psCtaCod As String, ByVal pnPorcDesc As Integer)
    
    Dim lsSQL As String
    lsSQL = "exec stp_upd_ColocPigPorcDescuento '" & psCtaCod & "' , " & pnPorcDesc
       
    coConex.CargaRecordSet lsSQL
    Exit Sub
   

End Sub

'JOEP20210915 campana prendario
Public Sub CampPrendExtorno(ByVal psCtaCod As String, ByVal pdFechaSis As String, ByVal pnModulo As Integer, ByVal pcMovnro As String)
Dim lsSQL As String
On Error GoTo error
    lsSQL = "EXEC stp_ins_CampanaPrendarioDesembolsoExtorno '" & psCtaCod & "','" & Format(pdFechaSis, "yyyymmdd") & "'," & pnModulo & ",'" & pcMovnro & "'"
    coConex.Ejecutar lsSQL
    Exit Sub
error:
    Err.Raise Err.Number, "Error en Proceso(ExtornoCampPrend)", Err.Description
End Sub
'JOEP20210915 campana prendario
