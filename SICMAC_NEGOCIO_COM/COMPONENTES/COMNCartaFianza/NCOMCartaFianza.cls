VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMCartaFianza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim mnEjecutaBatch As Integer
Dim mbTrans As Boolean


Public Sub nCFRegistraSolicitud(ByVal psNuevaCta As String, ByVal psCodAge As String, ByVal pnTipoProd As Producto, _
    ByVal pnMoneda As Moneda, ByVal pnCondicion As ColocCredCondicion, _
    ByVal pnModalidad As ColCFModalidad, ByVal psFinalidad As String, _
    ByVal pdRegistro As Date, ByVal oRelPers As Variant, _
    ByVal psPersCodAna As String, ByVal pdFecAsig As Date, _
    ByVal pdFecVenc As Date, ByVal pnMonto As Double, psMovAct As String, _
    ByVal psNumFuente As String, ByVal pdPersEval As Date, ByVal psPersCodAcre As String, Optional ByVal pbRenovacion As Boolean = False, Optional ByVal psCtaAntigua As String = "", _
    Optional ByVal psAgeCodAct As String = "", Optional ByVal psTpoProdCod As String = "", Optional ByVal psTpoCredCod As String = "", Optional ByVal psPersCodAval As String = "", Optional ByVal pbConsorcio As Boolean = False, Optional ByVal psConsorcio As String = "", Optional ByVal psOtrsModalidades As String = "") 'WIOR 20120328
    'MADM 20111020 psPersCodAval - 'MAVM 20100604 Se agrego los par: psAgeCodAct, psTpoProdCod, psTpoCredCod BAS II
    'JOEP20181219 CP psOtrsModalidades
    
Dim oBase As COMDCredito.DCOMCredActBD
Dim sMovNro As String
Dim i As Integer
Dim nPersCodTitular As String  'WIOR 20120328

    On Error GoTo ErrorNuevaCF
    Set oBase = New COMDCredito.DCOMCredActBD
    oBase.dBeginTrans

    Call oBase.dInsertProducto(psNuevaCta, 0#, 0#, gColocEstSolic, pdRegistro, 0, False)
'    oRelPers.IniciarMatriz
'    Do While Not oRelPers.EOF
'        Call oBase.dInsertProductoPersona(psNuevaCta, oRelPers.ObtenerCodigo, CInt(oRelPers.ObtenerValorRelac), False)
'        oRelPers.siguiente
'    Loop
    For i = 0 To UBound(oRelPers) - 1
      If CInt(oRelPers(i, 1)) <> gColRelPersAval Then 'WIOR 20120328
        Call oBase.dInsertProductoPersona(psNuevaCta, oRelPers(i, 0), CInt(oRelPers(i, 1)), False)
        If CInt(oRelPers(i, 1)) = 20 Then 'WIOR 20120328
            nPersCodTitular = oRelPers(i, 0) 'WIOR 20120328
        End If 'WIOR 20120328
      End If
    Next
    Call oBase.dInsertProductoPersona(psNuevaCta, psPersCodAcre, gColRelPersAcreedor, False)
    
    Call oBase.dInsertProductoPersona(psNuevaCta, psPersCodAna, gColRelPersAnalista, False)
    'MAVM 2010604 Se agregaron los par: psAgeCodAct, psTpoProdCod, psTpoCredCod BAS II ***
    Call oBase.dInsertColocaciones(psNuevaCta, 0, pdFecVenc, pnMonto, gColocCalendCodPFCF, "", pdFecAsig, psMovAct, False, psAgeCodAct, psTpoProdCod, psTpoCredCod)
    '***
    'Call oBase.dInsertColocFteIngreso(psNumFuente, psNuevaCta, pdPersEval) 'LUCV20160919, Segun ERS004-2016
    'Call oBase.dInsertColocCartaFianza(psNuevaCta, pnCondicion, pnModalidad, pdFecAsig, pdFecVenc, psFinalidad) 'COmento JOEP20190208 CP
    Call oBase.dInsertColocCartaFianza(psNuevaCta, pnCondicion, pnModalidad, pdFecAsig, pdFecVenc, psFinalidad, , , psOtrsModalidades) 'JOEP20181219 CP psOtrsModalidades
    
    'FRHU20131126 se agrego el parametro: pdFecAsig
    Call oBase.dInsertColocCFEstado(psNuevaCta, pdRegistro, gColocEstSolic, pnMonto, pdFecVenc, "Registro Carta Fianza", 0, False, , pdFecAsig)
    'FIN FRHU20131126
    If pbRenovacion = True Then
        Call oBase.dInsertColocCFRenovacion(psCtaAntigua, psNuevaCta)
    End If
    'MADM 20111020
    
    If pbConsorcio = True Then 'WIOR 20120328
        Call oBase.dInsertProductoPersona2(psNuevaCta, nPersCodTitular, gColRelPersAval, psConsorcio, False) 'WIOR 20120328
    Else
        If psPersCodAval <> "" Then
            Call oBase.dInsertProductoPersona(psNuevaCta, psPersCodAval, gColRelPersAval, False)
        End If
    End If
    'END MADM
    oBase.dCommitTrans
    Set oBase = Nothing

    Exit Sub

ErrorNuevaCF:
    oBase.dRollbackTrans
    Set oBase = Nothing
End Sub

Public Sub nCFActualizaSolicitud(ByVal psNuevaCta As String, ByVal psCodAge As String, ByVal pnTipoProd As Producto, _
    ByVal pnMoneda As Moneda, ByVal pnCondicion As ColocCredCondicion, _
    ByVal pnModalidad As ColCFModalidad, ByVal psFinalidad As String, _
    ByVal pdRegistro As Date, ByVal oRelPers As Variant, _
    ByVal psPersCodAna As String, ByVal pdFecAsig As Date, _
    ByVal pdFecVenc As Date, ByVal pnMonto As Double, psMovAct As String, _
    ByVal psNumFuente As String, ByVal pdPersEval As Date, ByVal psPersCodAcre As String, Optional ByVal pbRenovacion As Boolean = False, Optional ByVal psCtaAntigua As String = "", _
    Optional ByVal psAgeCodAct As String = "", Optional ByVal psTpoProdCod As String = "", Optional ByVal psTpoCredCod As String = "", Optional ByVal psPersCodAval As String = "", Optional ByVal pbConsorcio As Boolean = False, Optional ByVal psConsorcio As String = "", Optional ByVal psOtrsModalidades As String = "") 'WIOR 20120328
    'MADM 20111020 psAval -- 'MAVM 20100604 Se agrego los par: psAgeCodAct, psTpoProdCod, psTpoCredCod BAS II
'JOEP20181219 CP psOtrsModalidades

Dim oBase As COMDCredito.DCOMCredActBD
Dim sMovNro As String
Dim i As Integer
Dim nPersCodTitular As String   'WIOR 20120328
    
     
    Set oBase = New COMDCredito.DCOMCredActBD
    oBase.dBeginTrans
    
    Call oBase.dDeleteProductoPersona(psNuevaCta)
    Call oBase.dDeleteConsorcio(psNuevaCta) 'WIOR 20120330
'    oRelPers.IniciarMatriz
'    Do While Not oRelPers.EOF
'        Call oBase.dInsertProductoPersona(psNuevaCta, oRelPers.ObtenerCodigo, CInt(oRelPers.ObtenerValorRelac), False)
'        oRelPers.siguiente
'    Loop
    For i = 0 To UBound(oRelPers) - 1
        If CInt(oRelPers(i, 1)) <> gColRelPersAval Then 'WIOR 20120328
            Call oBase.dInsertProductoPersona(psNuevaCta, oRelPers(i, 0), CInt(oRelPers(i, 1)), False)
            If CInt(oRelPers(i, 1)) = 20 Then 'WIOR 20120328
                nPersCodTitular = oRelPers(i, 0) 'WIOR 20120328
            End If 'WIOR 20120328
        End If
    Next

    Call oBase.dInsertProductoPersona(psNuevaCta, psPersCodAcre, gColRelPersAcreedor, False)
    Call oBase.dInsertProductoPersona(psNuevaCta, psPersCodAna, gColRelPersAnalista, False)
    
    'MAVM 2010604 Se agregaron los par: psAgeCodAct, psTpoProdCod, psTpoCredCod BAS II ***
    Call oBase.dUpdateColocaciones(psNuevaCta, , pdFecVenc, pnMonto, , , pdFecAsig, , , , psAgeCodAct, psTpoProdCod, psTpoCredCod)
    '***
    'Call oBase.dDeleteColocFteIngreso(psNuevaCta) 'LUCV20160919, Comentó
    'Call oBase.dInsertColocFteIngreso(psNumFuente, psNuevaCta, pdPersEval) 'LUCV20160919, Comentó
    'Call oBase.dUpdateColocCartaFianza(psNuevaCta, pnCondicion, pnModalidad, pdFecAsig, pdFecVenc, psFinalidad)'Comento JOEP20181219 CP
    Call oBase.dUpdateColocCartaFianza(psNuevaCta, pnCondicion, pnModalidad, pdFecAsig, pdFecVenc, psFinalidad, , , , , , psOtrsModalidades) 'JOEP20181219 CP psOtrsModalidades
    Call oBase.dDeleteColocCFEstado(psNuevaCta, gColocEstSolic)
    'FRHU20131126 AGREGO el parametro pdFecAsig
    Call oBase.dInsertColocCFEstado(psNuevaCta, pdRegistro, gColocEstSolic, pnMonto, pdFecVenc, "Registro Carta Fianza", 0, False, , pdFecAsig)
    'FIN FRHU20131126
    
    If pbRenovacion = True Then
        Call oBase.dUpdateColocCFRenovacion(psCtaAntigua, psNuevaCta)
    End If
    
    'MADM 20111020
    If pbConsorcio = True Then 'WIOR 20120328
        Call oBase.dInsertProductoPersona2(psNuevaCta, nPersCodTitular, gColRelPersAval, psConsorcio, False) 'WIOR 20120328
    Else
        If psPersCodAval <> "" Then
            Call oBase.dInsertProductoPersona(psNuevaCta, psPersCodAval, gColRelPersAval, False)
        End If
    End If
    'END MADM
    
    oBase.dCommitTrans
    Set oBase = Nothing

End Sub

'** Sugerencia de Carta Fianza
Public Sub nCFSugerencia(ByVal psCtaCod As String, ByVal psFechaHora As String, ByVal pdFecAsigSug As Date, _
    ByVal pdFecVencSug As Date, ByVal pnMontoSug As Double, Optional ByVal psTpoCredCod As String = "", Optional psTipoInstCod As String = "", _
    Optional ByVal pvGravamen As Variant = Nothing, Optional ByVal psOtrsModalidades As String = "")
'FRHU20131126 se agrego el parametro: ByVal pdFecAsigSug As Date
'MAVM 20100605 se agrego la var: psTpoCredCod
'JOEP20181219 CP psOtrsModalidades
Dim loBase As COMDCredito.DCOMCredActBD
Dim lnMovNro As Long
Dim lvGravamen() As tGarantiaGravamen

'On Error GoTo ErrorModCF
Set loBase = New COMDCredito.DCOMCredActBD
    loBase.dBeginTrans
    mbTrans = True

    '** Actualiza Producto
    Call loBase.dUpdateProducto(psCtaCod, , , gColocEstSug, psFechaHora, -2, False)       ' (-2) aumenta el ste transac

    '** Actualiza Colocaciones
    Call loBase.dUpdateColocaciones(psCtaCod, , pdFecVencSug, pnMontoSug, , , , False, , , , , psTpoCredCod, psTipoInstCod)

    '** Actualiza CartaFianza
    'Call loBase.dUpdateColocCartaFianza(psCtaCod, , , , pdFecVencSug, , , False)'COmento JOEP20190208 CP
    Call loBase.dUpdateColocCartaFianza(psCtaCod, , , , pdFecVencSug, , , False, , , , psOtrsModalidades) 'JOEP20181219 CP
    
    '** Insert ColocCFEstado
    Call loBase.dDeleteColocCFEstado(psCtaCod, gColocEstSug)
    'FRHU20131126 Se agrego: pdFecAsigSug
    Call loBase.dInsertColocCFEstado(psCtaCod, psFechaHora, gColocEstSug, pnMontoSug, pdFecVencSug, "Sugerencia Carta Fianza", 0, False, , pdFecAsigSug)
    
    'EJVG20150715 ***
    If IsArray(pvGravamen) Then
        lvGravamen = pvGravamen
        If Not GrabarCoberturaGarantia(psCtaCod, "514", pnMontoSug, lvGravamen, loBase) Then
            loBase.dRollbackTrans
            Set loBase = Nothing
            Exit Sub
        End If
    End If
    'END EJVG *******

    '** Inserta Mov
    'Call loBase.dInsertMov(psMovNro, lsOpeCod, "Entrega Joya Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    ' Obtiene nMovNro
    'lnMovNro = loBase.dGetnMovNro(psMovNro)

    'mnEjecutaBatch = loBase.dEjecutaBatch
    loBase.dCommitTrans
    mbTrans = False

Set loBase = Nothing

Exit Sub
ErrorModCF:
    If mbTrans Then
        loBase.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise Err.Number, "Error nCFSugerencia ", "Error en Funcion de Sugerencia CF "
End Sub
'MADM 20111020
'** Aprobacion de Carta Fianza
'** Aprobacion de Carta Fianza
Public Sub nCFAprobacion(ByVal psCtaCod As String, ByVal psFechaHora As String, _
    ByVal pdFecVencApr As Date, ByVal pnMontoApr As Double, ByVal psApoderado As String, Optional ByVal psFinalidad As String, _
    Optional ByVal psAval As String = "", Optional ByVal pdFecVigeApr As Date, Optional ByVal pvGravamen As Variant = Nothing, Optional ByVal pcOtrsModalidades As String = "") 'FRHU20131126 Agregue el parametro: pdFecVigeApr
     'JOEP20181221 CP pcOtrsModalidades
Dim loBase As COMDCredito.DCOMCredActBD
Dim lnMovNro As Long
Dim lvGravamen() As tGarantiaGravamen

'On Error GoTo ErrorModCF
Set loBase = New COMDCredito.DCOMCredActBD
    loBase.dBeginTrans
    mbTrans = True

    '** Actualiza Producto
    Call loBase.dUpdateProducto(psCtaCod, , pnMontoApr, gColocEstAprob, psFechaHora, -2, False)          ' (-2) aumenta el ste transac
    '*** Insert ProductoPersona - Apoderado
    Call loBase.dInsertProductoPersona(psCtaCod, psApoderado, gColRelPersApoderado, False)

    If psAval <> "" Then
        Call loBase.dUpdatProductoPersona(psCtaCod, False, gColRelPersAval, psAval)
    End If
    
    '** Actualiza Colocaciones
    Call loBase.dUpdateColocaciones(psCtaCod, , pdFecVencApr, pnMontoApr, , , pdFecVigeApr, False) 'FRHU20131126 - Agregue: pdFecVigeApr como parametro

    '** Actualiza CartaFianza
    'By Capi Acta 035-2007 para que grabe modificacion en finalidad
    
    
    'Call loBase.dUpdateColocCartaFianza(psCtaCod, , , psFechaHora, pdFecVencApr, psFinalidad, , False)'FRHU20131126
    'Call loBase.dUpdateColocCartaFianza(psCtaCod, , , pdFecVigeApr, pdFecVencApr, psFinalidad, , False) 'FRHU20131126 - Agregue: pdFecVigeApr como parametro
     Call loBase.dUpdateColocCartaFianza(psCtaCod, , , pdFecVigeApr, pdFecVencApr, psFinalidad, , False, , , , pcOtrsModalidades) 'JOEP20190208 CP
    
    '** Insert ColocCFEstado
    'FRHU20131126 se agrego: pdFecVigeApr
    Call loBase.dInsertColocCFEstado(psCtaCod, psFechaHora, gColocEstAprob, pnMontoApr, pdFecVencApr, "Aprobacion Carta Fianza", 0, False, , pdFecVigeApr)

    'EJVG20150715 ***
    If IsArray(pvGravamen) Then
        lvGravamen = pvGravamen
        If Not GrabarCoberturaGarantia(psCtaCod, "514", pnMontoApr, lvGravamen, loBase) Then
            loBase.dRollbackTrans
            Set loBase = Nothing
            Exit Sub
        End If
    End If
    'END EJVG *******
    '** Inserta Mov
    'Call loBase.dInsertMov(psMovNro, lsOpeCod, "Entrega Joya Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    ' Obtiene nMovNro
    'lnMovNro = loBase.dGetnMovNro(psMovNro)

    'mnEjecutaBatch = loBase.dEjecutaBatch
    loBase.dCommitTrans
    mbTrans = False

Set loBase = Nothing

Exit Sub
ErrorModCF:
    If mbTrans Then
        loBase.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise Err.Number, "Error nCFSugerencia ", "Error en Funcion de Sugerencia CF "
End Sub

'** Rechazo de Carta Fianza
Public Sub nCFRechazo(ByVal psCtaCod As String, ByVal psFechaHora As String, _
    ByVal psComenta As String, ByVal pnMotivoRechazo As Integer)

Dim loBase As COMDCredito.DCOMCredActBD 'DCredActualizaBD
Dim lnMovNro As Long

'EJVG20151001 Se comentó ya que se actualizá los datos de garantía con el procedimiento [EliminarCoberturaGarantia]
'''*** PEAC 20090522 - PARA ACTUALIZAR LAS GARANTIAS
    'Dim oDCred As COMDCredito.DCOMCredito
    'Dim R As ADODB.Recordset

    'Set oDCred = New COMDCredito.DCOMCredito
    'Set R = oDCred.RecuperaColocGarantia(psCtaCod)
    'Set oDCred = Nothing
'*******************************************************

'On Error GoTo ErrorModCF
Set loBase = New COMDCredito.DCOMCredActBD
    loBase.dBeginTrans
    mbTrans = True

    '** Actualiza Producto
    Call loBase.dUpdateProducto(psCtaCod, , 0, gColocEstRech, psFechaHora, -2, False)            ' (-2) aumenta el ste transac

    '** Actualiza Colocaciones
    'Call loBase.dUpdateColocaciones(psCtaCod, , , , ,  ,  , False)

    '** Actualiza CartaFianza
    'Call loBase.dUpdateColocCartaFianza(psCtaCod, , , , pdFecVencApr, , , False)

    '** Insert ColocCFEstado
    Call loBase.dInsertColocCFEstado(psCtaCod, psFechaHora, gColocEstRech, 0, psFechaHora, "Rechazo Carta Fianza", pnMotivoRechazo, False)

    '** Inserta Mov
    'Call loBase.dInsertMov(psMovNro, lsOpeCod, "Entrega Joya Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    ' Obtiene nMovNro
    'lnMovNro = loBase.dGetnMovNro(psMovNro)

'''*** PEAC 20090522
    'Do While Not R.EOF
    '    Call loBase.dUpdateGarantias(R!cTpoDoc, R!cNroDoc, R!nGravado, 3, False)
    '    Call loBase.dDeleteColocGarantia(R!cNumGarant, psCtaCod, False)
    '    R.MoveNext
    'Loop
    'R.Close
    'Set R = Nothing
''***************************
    Call loBase.EliminarCoberturaGarantia(psCtaCod) 'EJVG20151001

    'mnEjecutaBatch = loBase.dEjecutaBatch
    loBase.dCommitTrans
    mbTrans = False

Set loBase = Nothing

Exit Sub
ErrorModCF:
    If mbTrans Then
        loBase.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise Err.Number, "Error nCFSugerencia ", "Error en Funcion de Sugerencia CF "
End Sub

'** Emsión de Carta Fianza
Public Sub nCFEmision(ByVal psCtaCod As String, ByVal psFechaHora As String, ByVal psMovNro As String, _
    ByVal pdFecVencEmi As Date, ByVal pnMontoEmi As Double, ByVal pdFecAsigEmi As Date, Optional ByVal pnTipo As Integer = 0, Optional ByVal pnTipoRiesgo As Integer = 0)
'JOEP20180622 Acta 122-2018 Agrego: Optional ByVal pnTipo As Integer = 0
'FRHU20131126 AGREGO: ByVal pdFecAsigEmi As Date
'WIOR 20120806 AGREGO PARAMETRO pnTipo
Dim loBase As COMDCredito.DCOMCredActBD
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = gColCFOpeEmision

'On Error GoTo ErrorModCF
Set loBase = New COMDCredito.DCOMCredActBD
    loBase.dBeginTrans
    mbTrans = True

    'MAVM 200907
    Call BloqueoPF(psCtaCod, psMovNro, psFechaHora)
    
    '** Actualiza Producto 'WIOR 20120806
    If pnTipo = 0 Then
        Call loBase.dUpdateProducto(psCtaCod, , , gColocEstVigNorm, psFechaHora, -2, False)              ' (-2) aumenta el ste transac
    ElseIf pnTipo = 1 Then
        Call loBase.dUpdateProducto(psCtaCod, , , gColocEstAprob, psFechaHora, -2, False)              ' (-2) aumenta el ste transac
    End If
    
    '** Actualiza Colocaciones
    'Call loBase.dUpdateColocaciones(psCtaCod, , , , , , psFechaHora, False) 'FRHU20131126
     Call loBase.dUpdateColocaciones(psCtaCod, , , , , , pdFecAsigEmi, False) 'FRHU20131126
    '** Actualiza CartaFianza
    'Call loBase.dUpdateColocCartaFianza(psCtaCod, , , ,  pdFecVencEmi, , , False)

    '** Insert ColocCFEstado'WIOR 20120806
    If pnTipo = 0 Then
        'FRHU20131126 SE AGRAGEO: pdFecAsigEmi
        Call loBase.dInsertColocCFEstado(psCtaCod, psFechaHora, gColocEstVigNorm, pnMontoEmi, pdFecVencEmi, "Emision Carta Fianza", 0, False, , pdFecAsigEmi, pnTipoRiesgo) 'JOEP20180622 Acta 122-2018 Agrego: pnTipoRiesgo
    ElseIf pnTipo = 1 Then
        'FRHU20131126 SE AGRAGEO: pdFecAsigEmi
        Call loBase.dInsertColocCFEstado(psCtaCod, psFechaHora, gColocEstModificada, pnMontoEmi, pdFecVencEmi, "Modificación de Carta Fianza", 0, False, , pdFecAsigEmi, pnTipoRiesgo) 'JOEP20180622 Acta 122-2018 Agrego: pnTipoRiesgo
    End If
    '** Inserta Mov'WIOR 20120806
    If pnTipo = 0 Then
        Call loBase.dInsertMov(psMovNro, lsOpeCod, "Emision Carta Fianza", gMovEstContabMovContable, gMovFlagVigente, False)
        
        ' Obtiene nMovNro
        lnMovNro = loBase.dGetnMovNro(psMovNro)
    
        '** Inserta MovCol
        Call loBase.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 0, pnMontoEmi, 0, "", 0, 0, gColocEstVigNorm, False)
    
        '** Inserta MovColDet -  Monto Carta Fianza
        Call loBase.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 0, gColCFConceptoCodCapital, 0, pnMontoEmi, False)
        
        Call loBase.dAsientoGarantiaDesembolso(lnMovNro, psCtaCod) 'EJVG20151117
    End If
    'mnEjecutaBatch = loBase.dEjecutaBatch
    loBase.dCommitTrans
    mbTrans = False

Set loBase = Nothing

Exit Sub
ErrorModCF:
    If mbTrans Then
        loBase.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise Err.Number, "Error nCFSugerencia ", "Error en Funcion de Sugerencia CF "
End Sub

'MAVM 200907 Bloqueo de Cuentas de PF
Public Sub BloqueoPF(ByVal psCtaCod As String, ByVal sMovNroBloq As String, ByVal psFechaHora As Date)
    Dim oBase As COMDCredito.DCOMCredActBD
    Dim rPlzFjoGar As ADODB.Recordset
    Dim sOperacionBloq As String
    Dim nMovNroBloq As Long
    
    Set oBase = New COMDCredito.DCOMCredActBD
    Set rPlzFjoGar = oBase.RecuperaPlazosFijosAsocComoGarantia(psCtaCod)
    While Not rPlzFjoGar.EOF
        oBase.NuevoBloqueoRetiro rPlzFjoGar!cCtaCod, 3, IIf(Mid(rPlzFjoGar!cCtaCod, 6, 3) = "234", "CTS", "Plazo Fijo") & " Asociado como garantía del Crédito " & psCtaCod, sMovNroBloq
        sOperacionBloq = IIf(Mid(rPlzFjoGar!cCtaCod, 6, 3) = "232", "201101", IIf(Mid(rPlzFjoGar!cCtaCod, 6, 3) = "233", "211101", "221101"))
        oBase.AgregaMov sMovNroBloq, sOperacionBloq, IIf(Mid(rPlzFjoGar!cCtaCod, 6, 3) = "234", "CTS", "Plazo Fijo") & " Asociado como garantía del Crédito " & psCtaCod, gMovEstContabMovContable, gMovFlagVigente
        nMovNroBloq = oBase.GetnMovNro(sMovNroBloq)
        oBase.AgregaMovCap nMovNroBloq, sOperacionBloq, rPlzFjoGar!cCtaCod, rPlzFjoGar!nSaldo, rPlzFjoGar!nSaldoDisp, rPlzFjoGar!nSaldo
        oBase.AgregaMovCapDet nMovNroBloq, sOperacionBloq, rPlzFjoGar!cCtaCod, gConcCapital, rPlzFjoGar!nSaldo
        oBase.ActualizaEstadoCuenta rPlzFjoGar!cCtaCod, 1200, Format(psFechaHora, "mm/dd/yyyy"), sMovNroBloq
    rPlzFjoGar.MoveNext
    Wend
    
End Sub

'** Comision de Carta Fianza
Public Sub nCFComision(ByVal psCtaCod As String, ByVal psFechaHora As String, ByVal psMovNro As String, _
    ByVal pnMontoComi As Double, _
    ByVal pbITFAplica As Boolean, ByVal pbITFAsumidoCreditos As Boolean, _
    ByVal pnMontoITf As Double, Optional ByVal pnTipo As Integer = 0, _
    Optional ByVal pbRenova As Boolean = False, Optional ByVal psMovAut As String = "") 'WIOR 20130313 AGREGO pbRenova Y psMovAut
'WIOR 20120807 Agrego el parametro pnTipo
Dim loBase As COMDCredito.DCOMCredActBD
Dim lnMovNro As Long
Dim lsOpeCod As String

'WIOR 20120807 agrego el pnTipo para comparacion
If pnTipo = 0 Then
    'lsOpeCod = gColCFOpeComisEfe
    lsOpeCod = gComiCredEmisionRenovCF 'JUEZ 20150928
ElseIf pnTipo = 1 Then
    'lsOpeCod = gColCFOpeComisMod
    lsOpeCod = gComiCredModifCF 'JUEZ 20150928
End If
    
'On Error GoTo ErrorModCF
Set loBase = New COMDCredito.DCOMCredActBD
    loBase.dBeginTrans
    mbTrans = True

    '** Actualiza Producto
    If pnTipo = 0 Then
        Call loBase.dUpdateProducto(psCtaCod, , , , psFechaHora, -2, False)              ' (-2) aumenta el ste transac
    ElseIf pnTipo = 1 Then
        Call loBase.dUpdateProducto(psCtaCod, , , gColocEstModificada, psFechaHora, -2, False)
    End If
    '** Actualiza Colocaciones
    'Call loBase.dUpdateColocaciones(psCtaCod, , , , , ,  psFechaHora, False)

    '** Actualiza CartaFianza
    'Call loBase.dUpdateColocCartaFianza(psCtaCod, , , ,  pdFecVencEmi, , , False)

    '** Inserta Mov
    'WIOR 20120807 agrego el pnTipo para comparacion
    If pnTipo = 0 Then
        Call loBase.dInsertMov(psMovNro, lsOpeCod, "Comision Carta Fianza", gMovEstContabMovContable, gMovFlagVigente, False)
    ElseIf pnTipo = 1 Then
        Call loBase.dInsertMov(psMovNro, lsOpeCod, "Comision por Modificacion de Carta Fianza", gMovEstContabMovContable, gMovFlagVigente, False)
    End If
    ' Obtiene nMovNro
    lnMovNro = loBase.dGetnMovNro(psMovNro)

    '** Inserta MovCol
    Call loBase.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 0, pnMontoComi, 0, "", 0, 0, 0, False)

    '** Inserta MovColDet -  Monto Carta Fianza
    'WIOR 20120807 agrego el pnTipo para comparacion
    If pnTipo = 0 Then
        Call loBase.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 0, gColCFConceptoCodComision, 0, pnMontoComi, False)
    ElseIf pnTipo = 1 Then
        Call loBase.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 0, gColCFConceptoCodComisionMod, 0, pnMontoComi, False)
    End If



    If pbITFAplica Then
        If Not pbITFAsumidoCreditos Then
            '** Inserta MovCol
            Call loBase.dInsertMovCol(lnMovNro, gITFCobroEfectivo, psCtaCod, 0, pnMontoITf, 0, "", 0, 0, 0, False)
            '** Inserta MovColDet -  Monto Carta Fianza
            Call loBase.dInsertMovColDet(lnMovNro, gITFCobroEfectivo, psCtaCod, 0, gConcITFCliente, 0, pnMontoITf, False)
        Else
            '** Inserta MovCol
            Call loBase.dInsertMovCol(lnMovNro, gITFCobroEfectivo, psCtaCod, 0, pnMontoITf, 0, "", 0, 0, 0, False)
            '** Inserta MovColDet -  Monto Carta Fianza
            Call loBase.dInsertMovColDet(lnMovNro, gITFCobroEfectivo, psCtaCod, 0, gConcITFAsumido, 0, pnMontoITf, False)
        End If
    End If

    'WIOR 20130313 *****************************************
    If pbRenova Then
        Dim oCFRenova As COMDCartaFianza.DCOMCartaFianza
        Set oCFRenova = New COMDCartaFianza.DCOMCartaFianza
        
        Call oCFRenova.InsActAutRenovacionCF(2, psCtaCod, , , , , , , , psMovAut, , lnMovNro, , 1)
        pbRenova = False
    End If
    'WIOR FIN **********************************************
    
    'mnEjecutaBatch = loBase.dEjecutaBatch
    loBase.dCommitTrans
    mbTrans = False

Set loBase = Nothing

Exit Sub
ErrorModCF:
    If mbTrans Then
        loBase.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise Err.Number, "Error nCFSugerencia ", "Error en Funcion de Sugerencia CF "
End Sub

Public Sub nCFExtornoComision(ByVal psCtaCod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False, _
        Optional ByVal psMotExtorno As Variant)

'** Actualiza Producto ( NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion)
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSql As String
Dim loBase As COMDCredito.DCOMCredActBD
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = gColCFOpeExtComis

'On Error GoTo ErrorModCF
Set loBase = New COMDCredito.DCOMCredActBD
    loBase.dBeginTrans
    mbTrans = True

    '** Actualiza Producto
    Call loBase.dUpdateProducto(psCtaCod, , , , psFechaHora, -2, False)        ' (-2) aumenta el ste transac

    '** Actualiza Colocaciones
    'Call loBase.dUpdateColocaciones(psCtaCod, , , , , , psMovNro, , False)

    '** Inserta Mov
    Call loBase.dInsertMov(psMovNro, lsOpeCod, "Extorno Comision CF", gMovEstContabMovContable, gMovFlagVigente, False)
    
        ' Obtiene nMovNro
    lnMovNro = loBase.dGetnMovNro(psMovNro)
    
    '**CTI3 (ferimoro) 02102018
    Call loBase.dInsertDetExtMov(lnMovNro, lsOpeCod, 0, "", False, psMotExtorno(0), psMotExtorno(1))
     
    '** Inserta MovCol
    Call loBase.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 0, pnMonto, 0, "", 0, 0, 0, False)

    '** Update Mov Anterior
    Call loBase.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loBase.dInsertMovRef(lnMovNro, pnMovNroAnt, False)
    
    'WIOR 20130313 *************************************************************
    Dim oCFRenova As COMDCartaFianza.DCOMCartaFianza
    Set oCFRenova = New COMDCartaFianza.DCOMCartaFianza
    Call oCFRenova.ExtornarPagoAutRenovaCF(psCtaCod, pnMovNroAnt, 0)
    'WIOR 20130313 *************************************************************
    'mnEjecutaBatch = loBase.dEjecutaBatch
    loBase.dCommitTrans
    mbTrans = False
Set loBase = Nothing

Exit Sub
ErrorModCF:
    If mbTrans Then
        loBase.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise Err.Number, "Error nCFExtornoComision", "Error en Funcion de Extorno de Comision"

End Sub

'** Retirar Carta Fianza Aprobada
Public Sub nCFRetirar(ByVal psCtaCod As String, ByVal psFechaHora As String, _
    ByVal psComenta As String, ByVal pnMotivoRechazo As Integer)

Dim loBase As COMDCredito.DCOMCredActBD
Dim lnMovNro As Long

'On Error GoTo ErrorModCF

'EJVG20151001 Se comentó ya que se actualizá los datos de garantía con el procedimiento [EliminarCoberturaGarantia]
'''*** PEAC 20090522 - PARA ACTUALIZAR LAS GARANTIAS
    'Dim oDCred As COMDCredito.DCOMCredito
    'Dim R As ADODB.Recordset

    'Set oDCred = New COMDCredito.DCOMCredito
    'Set R = oDCred.RecuperaColocGarantia(psCtaCod)
    'Set oDCred = Nothing
'*******************************************************


Set loBase = New COMDCredito.DCOMCredActBD
    loBase.dBeginTrans
    mbTrans = True

    '** Actualiza Producto
    Call loBase.dUpdateProducto(psCtaCod, , 0, gColocEstRetirado, psFechaHora, -2, False)             ' (-2) aumenta el ste transac

    '** Actualiza Colocaciones
    'Call loBase.dUpdateColocaciones(psCtaCod, , , , ,  ,  , False)

    '** Actualiza CartaFianza
    'Call loBase.dUpdateColocCartaFianza(psCtaCod, , , , pdFecVencApr, , , False)

    '** Insert ColocCFEstado
    Call loBase.dInsertColocCFEstado(psCtaCod, psFechaHora, gColocEstRetirado, 0, psFechaHora, "Rechazo Carta Fianza", pnMotivoRechazo, False)

    '** Inserta Mov
    'Call loBase.dInsertMov(psMovNro, lsOpeCod, "Entrega Joya Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
    ' Obtiene nMovNro
    'lnMovNro = loBase.dGetnMovNro(psMovNro)

'''*** PEAC 20090522
    'Do While Not R.EOF
    '    Call loBase.dUpdateGarantias(R!cTpoDoc, R!cNroDoc, R!nGravado, 3, False)
    '    Call loBase.dDeleteColocGarantia(R!cNumGarant, psCtaCod, False)
    '    R.MoveNext
    'Loop
    'R.Close
    'Set R = Nothing
''***************************
    Call loBase.EliminarCoberturaGarantia(psCtaCod) 'EJVG20151001

    'mnEjecutaBatch = loBase.dEjecutaBatch
    loBase.dCommitTrans
    mbTrans = False

Set loBase = Nothing

Exit Sub
ErrorModCF:
    If mbTrans Then
        loBase.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise Err.Number, "Error nCFSugerencia ", "Error en Funcion de Sugerencia CF "
End Sub

'** Devolución de Carta Fianza Emitida
Public Sub nCFDevolucion(ByVal psCtaCod As String, ByVal psFechaHora As String, ByVal psMovNro As String, _
    ByVal psComenta As String, ByVal pnMotivoDevolucion As Integer, ByVal pnMontoCF As Double)

Dim loBase As COMDCredito.DCOMCredActBD
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = gColCFOpeDevol

'On Error GoTo ErrorModCF
'EJVG20151001 Se comentó ya que se actualizá los datos de garantía con el procedimiento [EliminarCoberturaGarantia]
'''*** PEAC 20090522 - PARA ACTUALIZAR LAS GARANTIAS
    'Dim oDCred As COMDCredito.DCOMCredito
    'Dim R As ADODB.Recordset

    'Set oDCred = New COMDCredito.DCOMCredito
    'Set R = oDCred.RecuperaColocGarantia(psCtaCod)
    'Set oDCred = Nothing
'*******************************************************


Set loBase = New COMDCredito.DCOMCredActBD
    loBase.dBeginTrans
    mbTrans = True

    '** Actualiza Producto
    Call loBase.dUpdateProducto(psCtaCod, , , gColocEstDevuelta, psFechaHora, -2, False)                ' (-2) aumenta el ste transac

    '** Actualiza Colocaciones
    'Call loBase.dUpdateColocaciones(psCtaCod, , , , ,  ,  , False)

    '** Actualiza CartaFianza
    Call loBase.dUpdateColocCartaFianza(psCtaCod, , , , , , psFechaHora, False)

    '** Insert ColocCFEstado
    Call loBase.dInsertColocCFEstado(psCtaCod, psFechaHora, gColocEstDevuelta, pnMontoCF, psFechaHora, "Devolucion Carta Fianza", pnMotivoDevolucion, False)

    '** Inserta Mov
    Call loBase.dInsertMov(psMovNro, lsOpeCod, "Devolucion Carta Fianza", gMovEstContabMovContable, gMovFlagVigente, False)
    ' Obtiene nMovNro
    lnMovNro = loBase.dGetnMovNro(psMovNro)

    '** Inserta MovCol
    Call loBase.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 0, pnMontoCF, 0, "", 0, 0, gColocEstDevuelta, False)

    '** Inserta MovColDet -  Monto Carta Fianza
    Call loBase.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 0, gColCFConceptoCodCapital, 0, pnMontoCF, False)

'''*** PEAC 20090522
    'Do While Not R.EOF
    '    Call loBase.dUpdateGarantias(R!cTpoDoc, R!cNroDoc, R!nGravado, 3, False)
    '    Call loBase.dDeleteColocGarantia(R!cNumGarant, psCtaCod, False)
    '    R.MoveNext
    'Loop
    'R.Close
    'Set R = Nothing
''***************************
    Call loBase.EliminarCoberturaGarantia(psCtaCod) 'EJVG20151001

    'mnEjecutaBatch = loBase.dEjecutaBatch
    loBase.dCommitTrans
    mbTrans = False

Set loBase = Nothing

Exit Sub
ErrorModCF:
    If mbTrans Then
        loBase.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise Err.Number, "Error nCFDevolucion ", "Error en Funcion de Devolucion CF "
End Sub

Public Sub nCFHonrar(ByVal psCtaCod As String, ByVal psFechaHora As String, ByVal psMovNro As String, _
    ByVal psComenta As String, ByVal pnMotivoDevolucion As Integer, ByVal pnMontoCF As Double)

Dim loBase As COMDCredito.DCOMCredActBD
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = "140700"

'On Error GoTo ErrorModCF
Set loBase = New COMDCredito.DCOMCredActBD
    loBase.dBeginTrans
    mbTrans = True

    '** Actualiza Producto
    Call loBase.dUpdateProducto(psCtaCod, , 0, gColocEstHonrada, psFechaHora, -2, False)                ' (-2) aumenta el ste transac

    '** Actualiza Colocaciones
    'Call loBase.dUpdateColocaciones(psCtaCod, , , , ,  ,  , False)

    '** Actualiza CartaFianza
    Call loBase.dUpdateColocCartaFianza(psCtaCod, , , , , , psFechaHora, False)

    '** Insert ColocCFEstado
    Call loBase.dInsertColocCFEstado(psCtaCod, psFechaHora, gColocEstHonrada, pnMontoCF, psFechaHora, "Honrar Carta Fianza", pnMotivoDevolucion, False)

    '** Inserta Mov
    Call loBase.dInsertMov(psMovNro, lsOpeCod, "Honrar Carta Fianza", gMovEstContabMovContable, gMovFlagVigente, False)
    ' Obtiene nMovNro
    lnMovNro = loBase.dGetnMovNro(psMovNro)

    '** Inserta MovCol
    Call loBase.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 0, pnMontoCF, 0, "", 0, 0, gColocEstDevuelta, False)

    '** Inserta MovColDet -  Monto Carta Fianza
    Call loBase.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 0, gColCFConceptoCodCapital, 0, pnMontoCF, False)

    Call loBase.dUpdateCoberturaGarantiaxCredito(psCtaCod) 'EJVG20151014
    
    'mnEjecutaBatch = loBase.dEjecutaBatch
    loBase.dCommitTrans
    mbTrans = False

Set loBase = Nothing

Exit Sub
ErrorModCF:
    If mbTrans Then
        loBase.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise Err.Number, "Error nCFHonrar ", "Error en Funcion de Honrar CF "
End Sub

'** Relaciona Credito con Carta Fianza a Honrar
Public Sub nCFHonrarCredito(ByVal psCtaCod As String, ByVal psFechaHora As String, ByVal psMovNro As String, _
    ByVal pmCreditos As Variant)

Dim loBase As COMDCredito.DCOMCredActBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim lnCred As Integer
lsOpeCod = "140800"

'On Error GoTo ErrorModCF
Set loBase = New COMDCredito.DCOMCredActBD
    loBase.dBeginTrans
    mbTrans = True

    '** Actualiza Producto
    'Call loBase.dUpdateProducto(psCtaCod, , 0, gColocEstHonrada, psFechaHora, -2, False)                ' (-2) aumenta el ste transac
    '** Actualiza Colocaciones
    'Call loBase.dUpdateColocaciones(psCtaCod, , , , ,  ,  , False)
    '** Actualiza CartaFianza
    'Call loBase.dUpdateColocCartaFianza(psCtaCod, , , , , , psFechaHora, False)
    '** Insert ColocCFEstado
    'Call loBase.dInsertColocCFEstado(psCtaCod, psFechaHora, gColocEstHonrada, pnMontoCF, psFechaHora, "Honrar Carta Fianza", pnMotivoDevolucion, False)

    '** Inserta ColoCFCredito
    Call loBase.dDeleteColocCFCredito(psCtaCod)
    For lnCred = 0 To UBound(pmCreditos) - 1
        Call loBase.dInsertColocCFCredito(psCtaCod, pmCreditos(lnCred, 0))
    Next
    '** Inserta Mov
    Call loBase.dInsertMov(psMovNro, lsOpeCod, "Honrar Carta Fianza - Credito", gMovEstContabMovContable, gMovFlagVigente, False)
    ' Obtiene nMovNro
    lnMovNro = loBase.dGetnMovNro(psMovNro)

    '** Inserta MovCol
    'Call loBase.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 0, pnMontoCF, 0, "", 0, 0, gColocEstDevuelta, False)

    '** Inserta MovColDet -  Monto Carta Fianza
    'Call loBase.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 0, gColCFConceptoCodCapital, 0, pnMontoCF, False)

    'mnEjecutaBatch = loBase.dEjecutaBatch
    loBase.dCommitTrans
    mbTrans = False

Set loBase = Nothing

Exit Sub
ErrorModCF:
    If mbTrans Then
        loBase.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise Err.Number, "Error nCFHonrar ", "Error en Funcion de Honrar CF "
End Sub

'****** Procedimiento que se correra en el cierre diari/mes para Liberar Garantia *******
'****** AVMM-28/07/2006 ******

Public Function CierreCartaFianza(ByVal pdFecha As Date)
   Dim coConex As COMConecta.DCOMConecta
   Dim rsTC As ADODB.Recordset
   Dim rs As ADODB.Recordset
   Dim rsG As ADODB.Recordset
   Dim sSql As String
   Dim nOpeTpo As String
   Dim nMontoCol As Double
   Dim nSumMontoGra As Double
   Dim nPagado As Double
   Dim nMovNro As Long
   Dim loBase As COMDCredito.DCOMCredActBD

   Set coConex = New COMConecta.DCOMConecta
        If coConex.AbreConexion = True Then
            'Obtener total de Cartas Horadas a una Fecha
            'sSql = " Select P.cCtaCod From Producto P " & _
            '       " Join Colocaciones C on P.cCtaCod = C.cCtaCod " & _
            '       " Where nPrdEstado in(2020,2021,2022,2090,2091) and Substring(cultimaActualizacion,1,8) = '" & Format(pdFecha, "yyyymmdd") & "'" & _
            '       " AND SUBSTRING(P.cCtaCod,6,3) IN('121','221')"    'ARCV 30-06-2007
            
            '**Comentado por DAOR 20071005
            'sSql = " Select P.cCtaCod From Producto P " & _
            '       " Join Colocaciones C on P.cCtaCod = C.cCtaCod " & _
            '       " Where nPrdEstado=2050 and dPrdEstado = '" & Format(pdFecha, "yyyymmdd") & "'" & _
            '       " AND SUBSTRING(P.cCtaCod,6,3) IN('121','221')"    'ARCV 30-06-2007
            
            '**DAOR 20071005, Sòlo considerar cartas fianzas cancelados de forma automática (en proceso de cierre)
'            sSql = " Select P.cCtaCod From Producto P " & _
'                   "    inner join Colocaciones C on P.cCtaCod = C.cCtaCod " & _
'                   "    inner join ColocCFEstado CE on C.cCtaCod=CE.cCtaCod and CE.nPrdEstado=2050 and CE.cDescripcion like '%Cancela%Cierre%' " & _
'                   " where P.nPrdEstado=2050 and P.dPrdEstado = '" & Format(pdFecha, "yyyymmdd") & "'" & _
'                   "    and substring(P.cCtaCod,6,3) IN ('121','221')"
            'ALPA 20100626*******************************
            sSql = " Select P.cCtaCod From Producto P " & _
                   "    inner join Colocaciones C on P.cCtaCod = C.cCtaCod " & _
                   "    inner join ColocCFEstado CE on C.cCtaCod=CE.cCtaCod and CE.nPrdEstado=2050 and CE.cDescripcion like '%Cancela%Cierre%' " & _
                   " where P.nPrdEstado=2050 and P.dPrdEstado = '" & Format(pdFecha, "yyyymmdd") & "'" & _
                   "    and substring(C.cTpoCredCod,2,2) IN ('81')"
                   
                   
            Set rsTC = New ADODB.Recordset
            rsTC.CursorLocation = adUseClient
            coConex.ConexionActiva.CommandTimeout = 1200
            rsTC.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
            Set rsTC.ActiveConnection = Nothing
            
            If Not (rsTC.EOF And rsTC.BOF) Then
                Do Until rsTC.EOF
                   'Obtener Operacion
                   sSql = " select cOpeCod from Garantias G " & _
                          " Join ColocGarantia CG on G.cNumGarant=CG.cNumGarant" & _
                          " Join GarantiaOpe GO on GO.nTpoGarantia =G.nTpoGarantia " & _
                          " where cCtaCod='" & rsTC!cCtaCod & "' and GO.nMoneda=substring('" & rsTC!cCtaCod & "',9,1) and bReversion=1"
                   Set rs = New ADODB.Recordset
                   rs.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
                   If Not (rs.EOF And rs.BOF) Then
                      nOpeTpo = rs!cOpeCod
                   End If
                   rs.Close
                   
                   'Obtener nMontoCol
                   Set rs = New ADODB.Recordset
                   sSql = "select nMontoCol from Colocaciones where cctacod='" & rsTC!cCtaCod & "'"
                   rs.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
                   If Not (rs.EOF And rs.BOF) Then
                      nMontoCol = rs!nMontoCol
                   End If
                   rs.Close
                   
                   'Obtener Sumatoria de Grantias
                   Set rs = New ADODB.Recordset
                   sSql = "select sum(nGravado)Total from ColocGarantia  where cctacod='" & rsTC!cCtaCod & "'"
                   rs.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
                   If Not (rs.EOF And rs.BOF) Then
                      nSumMontoGra = rs!Total
                   End If
                   rs.Close
                   
                   'Obtener total de garantias Credito
                   Set rsG = New ADODB.Recordset
                   sSql = "  select nGravado,G.cNumGarant from Garantias G " & _
                          "  Join ColocGarantia CG on G.cNumGarant=CG.cNumGarant" & _
                          "  where cCtaCod='" & rsTC!cCtaCod & "'"
                   rsG.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
                   If Not (rsG.EOF And rsG.BOF) Then
                      Do Until rsG.EOF
                        
                        nPagado = nMontoCol * (rsG!nGravado / nSumMontoGra)
                        
                        sSql = "Update Garantias set nGravament= nGravament - " & nPagado & " where cNumGarant ='" & rsG!cNumGarant & "'"
                        coConex.Ejecutar sSql
                        
                        sSql = "Update ColocGarantia  set nEstado=0 where cCtaCod='" & rsTC!cCtaCod & "' "
                        coConex.Ejecutar sSql
                        
                        'obtener mov
                        Set rs = New ADODB.Recordset
                        sSql = " Select top 1 M.nMovNro From Mov M" & _
                               " Join MovCol MC ON M.nMovNro =MC.nMovNro" & _
                               " Where cCtaCod='" & rsTC!cCtaCod & "'order by M.nMovNro desc"
                        rs.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
                        If Not (rs.EOF And rs.BOF) Then
                            nMovNro = rs!nMovNro
                        End If
                        rs.Close
                        
                        'Insertar en el MovCol y MovColDet
                        Set loBase = New COMDCredito.DCOMCredActBD
                            '** Inserta MovCol
                            Call loBase.dInsertMovCol(nMovNro, nOpeTpo, rsTC!cCtaCod, 0, nPagado, 0, "", 0, 0, gColocEstHonrada, False)
                        
                            '** Inserta MovColDet -  Monto Carta Fianza
                            Call loBase.dInsertMovColDet(nMovNro, nOpeTpo, rsTC!cCtaCod, 0, 4000, 0, nPagado, False)

                        Set loBase = Nothing
                        rsG.MoveNext
                      Loop
                   End If
                   rsG.Close
                   rsTC.MoveNext
                Loop
            End If
        End If
   Set coConex = Nothing
   
End Function


'** Rechazo de Carta Fianza
'By Capi Acta 035-2007 se añadio parametro opcional pnMonto
Public Sub nCFRenovacion(ByVal psCtaCod As String, ByVal pdNewFechaVec As Date, _
                         ByVal pnRenovacion As Integer, ByVal pdFecRenovacion As Date, Optional ByVal pnMonto)

Dim loBase As COMDCredito.DCOMCredActBD 'DCredActualizaBD
Dim lnMovNro As Long
Dim rs As ADODB.Recordset
Dim lnMonto As Double
Dim lsSql As String
'On Error GoTo ErrorModCF
Set loBase = New COMDCredito.DCOMCredActBD

    'Bay Capi Acta 035-2007 para añadir parametro pnMonto
     '** Actualiza Producto
    Call loBase.dUpdateProducto(psCtaCod, , pnMonto, gColocEstRenovada, pdFecRenovacion, -2, False)
    'By Capi 05032008 para que tambien actualice tabla colocaciones
    Call loBase.dUpdateColocaciones(psCtaCod, , pdNewFechaVec, pnMonto, , , pdFecRenovacion, False, False, "")
    
    '** Actualiza CartaFianza
    Call loBase.dUpdateColocCartaFianza(psCtaCod, , , pdFecRenovacion, pdNewFechaVec, , , , pnRenovacion, pdFecRenovacion)
    
    lsSql = "select nSaldo from Producto where cCtaCod ='" & psCtaCod & "'"
    Set rs = New ADODB.Recordset
        Set rs = loBase.CargaRecordSet(lsSql)
        If Not (rs.EOF And rs.BOF) Then
            lnMonto = rs!nSaldo
        End If
    Set rs = Nothing
    
    loBase.dBeginTrans
    mbTrans = True
    
    
    
     '** Insert ColocCFEstado
    Call loBase.dDeleteColocCFEstado(psCtaCod, gColocEstRenovada)
    Call loBase.dInsertColocCFEstado(psCtaCod, pdFecRenovacion, gColocEstRenovada, lnMonto, pdNewFechaVec, "Renovacion de Carta Fianza", 0, False)


    loBase.dCommitTrans
    mbTrans = False

Set loBase = Nothing

Exit Sub
ErrorModCF:
    If mbTrans Then
        loBase.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise Err.Number, "Error nCFSugerencia ", "Error en Funcion de Sugerencia CF "
End Sub

'MAVM 20201216
Public Sub nCFRegistraHistorial(ByVal psCtaCod As String, ByVal pdFecVencAnt As Date, ByVal pnPeriodo As Integer, _
    ByVal pdFecVenc As Date, ByVal pnMonto As Double)

    Dim oBase As COMDCredito.DCOMCredActBD
    Set oBase = New COMDCredito.DCOMCredActBD
    
On Error GoTo ErrorHistorialCF
    oBase.dBeginTrans
    Call oBase.dInsertColocCFHistorial(psCtaCod, pdFecVencAnt, pnPeriodo, pdFecVenc, pnMonto, False)
    oBase.dCommitTrans
    Set oBase = Nothing
    Exit Sub

ErrorHistorialCF:
    oBase.dRollbackTrans
    Set oBase = Nothing
End Sub

'MADM 20110205 - Extorno Aprobacion
'** Rechazo de Carta Fianza
Public Sub nCFExtornoApro(ByVal psCtaCod As String, ByVal psFechaHora As String)

Dim loBase As COMDCredito.DCOMCredActBD 'DCredActualizaBD
Dim lnMovNro As Long

Dim oDCred As COMDCredito.DCOMCredito

On Error GoTo ErrorModCF
Set loBase = New COMDCredito.DCOMCredActBD
    loBase.dBeginTrans
    mbTrans = True

    '** Actualiza Producto
    Call loBase.dUpdateProducto(psCtaCod, , 0, gColocEstSug, psFechaHora, -1)               ' (-2) aumenta el ste transac

    'dDeleteProductoPersona
    Call loBase.dDeleteProductoPersonaRelac(psCtaCod, gColRelPersApoderado)
    
    '** Insert ColocCFEstado
     Call loBase.dDeleteColocCFEstado(psCtaCod, gColocEstAprob)
    '** Actualiza Colocaciones
    Call loBase.dUpdateColocaciones(psCtaCod, , , 0, , , , False)

    '** Actualiza CartaFianza
    Call loBase.dUpdateColocCartaFianza(psCtaCod, , , , , "", , False)
        
    loBase.dCommitTrans
    mbTrans = False

Set loBase = Nothing
Exit Sub
ErrorModCF:
    If mbTrans Then
        loBase.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise Err.Number, "Error nCFSugerencia ", "Error en Funcion de Sugerencia CF "
End Sub
'EJVG20150715 ***
Public Function GrabarCoberturaGarantia(ByVal psCtaCod As String, ByVal psTpoProdCod As String, ByVal pnMontoColocado As Currency, ByRef pvDatos() As tGarantiaGravamen, Optional ByVal pConnBase As COMDCredito.DCOMCredActBD = Nothing) As Boolean
    Dim oBase As COMDCredito.DCOMCredActBD
    Dim bTrans As Boolean
    Dim i As Integer
    
    On Error GoTo ErrGrabarCoberturaGarantia
    
    GrabarCoberturaGarantia = True
    If UBound(pvDatos) <= 0 Then Exit Function
    
    GrabarCoberturaGarantia = False
    
    If pConnBase Is Nothing Then
        Set oBase = New COMDCredito.DCOMCredActBD
        oBase.dBeginTrans
    Else
        Set oBase = pConnBase
    End If

    bTrans = True
    
    oBase.EliminarCoberturaGarantia psCtaCod
    oBase.RegistrarColocGravamen psCtaCod, psTpoProdCod, pnMontoColocado
    For i = 1 To UBound(pvDatos)
        oBase.RegistrarCoberturaGarantia pvDatos(i).sNumgarant, pvDatos(i).nMoneda, psCtaCod, pvDatos(i).dFechaValorizacion, pvDatos(i).nSaldoCobertura, pvDatos(i).nRatio, pvDatos(i).nMaxCobertura, pvDatos(i).nCobertura, pvDatos(i).nOrden, pvDatos(i).dFechaTramiteLegal
    Next

    If pConnBase Is Nothing Then
        oBase.dCommitTrans
        Set oBase = Nothing
    End If
    bTrans = False
    
    GrabarCoberturaGarantia = True
    Exit Function
ErrGrabarCoberturaGarantia:
    GrabarCoberturaGarantia = False
    If bTrans Then
        If pConnBase Is Nothing Then
            oBase.dRollbackTrans
            Set oBase = Nothing
        End If
    End If
    Err.Raise Err.Number, "Ocurrió un error al grabar la cobertura de garantías del Producto con Id: " & psCtaCod, Err.Description
End Function
'END EJVG *******
'RECO20160310 ERS012-2016*******************************************************************
Public Function ObtieneDatosAprobacionCF(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oCF As New COMDCartaFianza.DCOMCartaFianza
    Set ObtieneDatosAprobacionCF = oCF.ObtieneDatosAprobacionCF(psCtaCod)
    Set oCF = Nothing
End Function
Public Function ObtieneTasaCompCF(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oCF As New COMDCartaFianza.DCOMCartaFianza
    Set ObtieneTasaCompCF = oCF.ObtieneTasaCompCF(psCtaCod)
    Set oCF = Nothing
End Function
Public Function ObtieneTitAvaPagare(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oCF As New COMDCartaFianza.DCOMCartaFianza
    Set ObtieneTitAvaPagare = oCF.ObtieneTitAvaPagare(psCtaCod)
    Set oCF = Nothing
End Function
'RECO FIN***********************************************************************************
