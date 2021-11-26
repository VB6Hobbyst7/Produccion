VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMColRecCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim mnEjecutaBatch As Integer
Dim mbTrans As Boolean


Public Sub nCambiaMetodoLiquidCredRecup(ByVal psCtaCod As String, ByVal psMetliquid As String, _
        ByVal psFechaHora As String, ByVal psMovNro As String, _
        Optional pbEjecBatch As Boolean = False)

'** Actualiza Producto (NroTrans )
'** Actualiza ColocCredRecup (Metodo Liquid)
'************************************

Dim lsSQL As String
Dim loRegRec As COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
'Codigo de Operacion
'lsOpeCod = geColPImpDuplicado

'On Error GoTo ErrorModRec
Set loRegRec = New COMDColocPig.DCOMColPActualizaBD
    loRegRec.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loRegRec.dUpdateProducto(psCtaCod, , , , psFechaHora, -2, False)           ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    'Call loRegRec.dUpdateColocaciones(psCtaCod, , , , ,  psMovNro, , False)

    '** Actualiza ColocRecup
    Call loRegRec.dUpdateColocRecup(psCtaCod, , , , , , , psMetliquid, , , , False)

    '** Inserta Mov
    'Call loRegRec.dInsertMov(psMovNro, lsOpeCod, "Cambio Metodo Liquid", gMovEstContabNoContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    'lnMovNro = loRegrec.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    'Call loRegRec.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 1, pnMontoTransac, 0, "", 0, False)
    
    'mnEjecutaBatch = loRegRec.dEjecutaBatch
    loRegRec.dCommitTrans
    mbTrans = False
Set loRegRec = Nothing

Exit Sub
ErrorModRec:
    If mbTrans Then
        loRegRec.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nCambiaMetodoLiquidCredRecup", "Error en Funcion de Cambio Metodo Liquid Recuperaciones "

End Sub

Public Sub nGastoRecupAsignaNuevo(ByVal psCtaCod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnNroGastoCta As Integer, ByVal psFecAsigna As String, ByVal pnColocConceptoCod As Integer, _
        ByVal pnMonto As Currency, ByVal pnMontoPagado As Currency, ByVal pnColocRecGastoEstado As ColocRecGastoEstado, _
        ByVal psMotivoGasto As String, ByVal pnNewSaldoGasto As Currency, Optional pbEjecBatch As Boolean = False, _
        Optional ByVal pbCredTransferido As Boolean = False)
'FRHU 20150214 ERS022-2015: Se agrego pbCredTransferido
'** Inserta ColocRecGastos
'** Actualiza ColocRecup
'** dInsertMov
'** dInsertMovCol
'** dInsertMovColDet
'************************************

Dim lsSQL As String
Dim loReg As COMDColocPig.DCOMColPActualizaBD
'JUEZ 20121013 *********************************************************************
Dim loRec As COMDColocRec.DCOMColRecRConsulta
Dim lrs As ADODB.Recordset
Dim sMensaje As String
'END JUEZ **************************************************************************
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = 130801 'JUEZ 20121013

On Error GoTo ControlError
Set loReg = New COMDColocPig.DCOMColPActualizaBD
'JUEZ 20121013 *********************************************************************
Set loRec = New COMDColocRec.DCOMColRecRConsulta
Set lrs = loRec.dObtieneDatosAsigCredRecup(psCtaCod)
Set loRec = Nothing
'END JUEZ **************************************************************************
    loReg.dBeginTrans
    mbTrans = True
    
    If pbCredTransferido Then 'FRHU 20150415 ERS022-2015: Se agrego IF
        Call loReg.dUpdateColocTransferido(psCtaCod, , , , , pnNewSaldoGasto, , , False)
        Call loReg.dInsertColocTransferidoGastos(psCtaCod, pnNroGastoCta, pnColocConceptoCod, psFecAsigna, pnMonto, pnMontoPagado, _
            pnColocRecGastoEstado, psMotivoGasto, False)
    Else
    '** Actualiza Producto
    'Call loReg.dUpdateProducto(psCtaCod, , , , psFechaHora, -2, False)     ' (-2) aumenta el ste transac
    
    '** Actualiza ColocRecup
    'Comentado Por MAVM 20110418
    'Call loReg.dUpdateColocRecup(psCtaCod, , , , , pnNewSaldoGasto, , , , , , False)
    'Call loReg.dUpdateColocRecup(psCtaCod, , , , , pnNewSaldoGasto + pnMonto, , , , , , False)
    'Comentado x JUEZ 20121013
    Call loReg.dUpdateColocRecup(psCtaCod, , , , , pnNewSaldoGasto, , , , , , False)

    '** Inserta ColocRecupGastos
    Call loReg.dInsertColocRecupGastos(psCtaCod, pnNroGastoCta, pnColocConceptoCod, psFecAsigna, pnMonto, pnMontoPagado, _
            pnColocRecGastoEstado, psMotivoGasto, False)

    'Descomentado por JUEZ 20121013 ****************************************************************
    If lrs!nPrdEstado = gColocEstRecVigCast Or lrs!nPrdEstado = "2206" Then
        '** Inserta Mov
        Call loReg.dInsertMov(psMovNro, lsOpeCod, "Asignacion de Gasto: " & psCtaCod, gMovEstContabMovContable, gMovFlagVigente, False)
        
        ' Obtiene nMovNro
        lnMovNro = loReg.dGetnMovNro(psMovNro)
        
        '** Inserta MovCol
        Call loReg.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 0, pnMonto, 0, "", 0, lrs!nPrdEstado, lrs!nSaldo)
        
        '** Inserta MovColDet
        Call loReg.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 0, pnColocConceptoCod, 0, pnMonto)
    End If
    Set lrs = Nothing
    'END JUEZ **************************************************************************************
    End If
    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loReg.dCommitTrans
    mbTrans = False
Set loReg = Nothing

Exit Sub
ControlError:
    If mbTrans Then
        loReg.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nGastoRecupAsignaNuevo", "Error en Funcion de Asignacion de Gastos "

End Sub

Public Sub nGastoRecupModifica(ByVal psCtaCod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnNroGastoCta As Integer, ByVal psFecAsigna As String, ByVal pnColocConceptoCod As Integer, _
        ByVal pnMonto As Currency, ByVal pnMontoPagado As Currency, ByVal pnColocRecGastoEstado As ColocRecGastoEstado, _
        ByVal psMotivoGasto As String, ByVal pnNewSaldoGasto As Currency, Optional pbEjecBatch As Boolean = False, _
        Optional ByVal pbCredTransferido As Boolean = False)
'FRHU 20150214 ERS022-2015: Se agrego pbCredTransferido
'** Update ColocRecGastos
'** Actualiza ColocRecup
'** dInsertMov
'** dInsertMovCol
'** dInsertMovColDet
'************************************

Dim lsSQL As String
Dim loReg As COMDColocPig.DCOMColPActualizaBD
'JUEZ 20121013 *********************************************************************
Dim loRec As COMDColocRec.DCOMColRecRConsulta
Dim loRecCons As COMDColocRec.DCOMColRecRConsulta
Dim lrs As ADODB.Recordset
Dim lrsGasto As ADODB.Recordset
Dim sMensaje As String
Dim nNewMonto As Double
Dim lsOpeCodAum As String, lsOpeCodDism As String
'END JUEZ **************************************************************************
Dim lnMovNro As Long
'Dim lsOpeCod As String

lsOpeCodAum = 130801 'JUEZ 20121013
lsOpeCodDism = 130802 'JUEZ 20121013

On Error GoTo ControlError
Set loReg = New COMDColocPig.DCOMColPActualizaBD
'JUEZ 20121013 *********************************************************************
Set loRec = New COMDColocRec.DCOMColRecRConsulta
Set loRecCons = New COMDColocRec.DCOMColRecRConsulta

Set lrs = loRec.dObtieneDatosAsigCredRecup(psCtaCod)
Set lrsGasto = loRecCons.dObtieneGastoRecuperacionXNroReg(psCtaCod, pnNroGastoCta)
Set loRec = Nothing
Set loRecCons = Nothing
'END JUEZ **************************************************************************
    loReg.dBeginTrans
    mbTrans = True
    
    If pbCredTransferido Then 'FRHU 20150415 ERS022-2015: Se agrego IF
        Call loReg.dUpdateColocTransferido(psCtaCod, , , , , pnNewSaldoGasto, , , False)
        Call loReg.dUpdateColocTransferidoGastos(psCtaCod, pnNroGastoCta, , pnMonto, , , psMotivoGasto, False)
    Else
        '** Actualiza Producto
        'Call loReg.dUpdateProducto(psCtaCod, , , , psFechaHora, -2, False)     ' (-2) aumenta el ste transac
        
        '** Actualiza ColocRecup
        Call loReg.dUpdateColocRecup(psCtaCod, , , , , pnNewSaldoGasto, , , , , , False)
    
        '** Update ColocRecupGastos
        Call loReg.dUpdateColocRecupGastos(psCtaCod, pnNroGastoCta, , pnMonto, , , psMotivoGasto, False)
    
        'Descomentado por JUEZ 20121013 ****************************************************************
        If lrs!nPrdEstado = gColocEstRecVigCast Or lrs!nPrdEstado = "2206" Then
            '** Inserta Mov
            nNewMonto = lrsGasto!nMonto - pnMonto
            
            Call loReg.dInsertMov(psMovNro, IIf(nNewMonto < 0, lsOpeCodAum, lsOpeCodDism), "Modificacion de Gasto " & pnNroGastoCta & ": " & psCtaCod, gMovEstContabMovContable, gMovFlagVigente, False)
            
            ' Obtiene nMovNro
            lnMovNro = loReg.dGetnMovNro(psMovNro)
            
            '** Inserta MovCol
            Call loReg.dInsertMovCol(lnMovNro, IIf(nNewMonto < 0, lsOpeCodAum, lsOpeCodDism), psCtaCod, 0, nNewMonto * IIf(nNewMonto < 0, -1, 1), 0, "", 0, lrs!nPrdEstado, lrs!nSaldo)
            
            '** Inserta MovColDet
            Call loReg.dInsertMovColDet(lnMovNro, IIf(nNewMonto < 0, lsOpeCodAum, lsOpeCodDism), psCtaCod, 0, lrsGasto!nPrdConceptoCod, 0, nNewMonto * IIf(nNewMonto < 0, -1, 1))
        End If
    End If
    Set lrs = Nothing
    Set lrsGasto = Nothing
    'END JUEZ **************************************************************************************
 
    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loReg.dCommitTrans
    mbTrans = False
Set loReg = Nothing

Exit Sub
ControlError:
    If mbTrans Then
        loReg.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nGastoRecupModifica", "Error en Funcion de Modifica Gasto Recup "

End Sub

Public Sub nGastoRecupEliminaGasto(ByVal psCtaCod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnNroGastoCta As Integer, _
        ByVal pnColocRecGastoEstado As ColocRecGastoEstado, _
        ByVal pnNewSaldoGasto As Currency, Optional pbEjecBatch As Boolean = False, _
        Optional ByVal pbCredTransferido As Boolean = False)
'FRHU 20150214 ERS022-2015: Se agrego pbCredTransferido
'** Actualiza ColocRecup
'** Update ColocRecGastos
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loReg As COMDColocPig.DCOMColPActualizaBD
'JUEZ 20121013 *********************************************************************
Dim loRec As COMDColocRec.DCOMColRecRConsulta
Dim loRecCons As COMDColocRec.DCOMColRecRConsulta
Dim lrs As ADODB.Recordset
Dim lrsGasto As ADODB.Recordset
Dim sMensaje As String
'END JUEZ **************************************************************************
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = 130802 'JUEZ 20121013

On Error GoTo ControlError
Set loReg = New COMDColocPig.DCOMColPActualizaBD
'JUEZ 20121013 *********************************************************************
Set loRec = New COMDColocRec.DCOMColRecRConsulta
Set loRecCons = New COMDColocRec.DCOMColRecRConsulta

Set lrs = loRec.dObtieneDatosAsigCredRecup(psCtaCod)
Set lrsGasto = loRecCons.dObtieneGastoRecuperacionXNroReg(psCtaCod, pnNroGastoCta)
Set loRec = Nothing
Set loRecCons = Nothing
'END JUEZ **************************************************************************
    loReg.dBeginTrans
    mbTrans = True
    
    If pbCredTransferido Then 'FRHU 20150415 ERS022-2015: Se agrego IF
        Call loReg.dUpdateColocTransferido(psCtaCod, , , , , pnNewSaldoGasto, , , False)
        Call loReg.dUpdateColocTransferidoGastos(psCtaCod, pnNroGastoCta, , , , pnColocRecGastoEstado, , False)
    Else
        '** Actualiza Producto
        'Call loReg.dUpdateProducto(psCtaCod, , , , psFechaHora, -2, False)     ' (-2) aumenta el ste transac
        
        '** Actualiza ColocRecup
        Call loReg.dUpdateColocRecup(psCtaCod, , , , , pnNewSaldoGasto, , , , , , False)
    
        '** Update ColocRecupGastos
        Call loReg.dUpdateColocRecupGastos(psCtaCod, pnNroGastoCta, , , , pnColocRecGastoEstado, , False)
        
        'Descomentado por JUEZ 20121013 ****************************************************************
        If lrs!nPrdEstado = gColocEstRecVigCast Or lrs!nPrdEstado = "2206" Then
            '** Inserta Mov
            Call loReg.dInsertMov(psMovNro, lsOpeCod, "Eliminacion de Gasto " & pnNroGastoCta & ": " & psCtaCod, gMovEstContabMovContable, gMovFlagVigente, False)
            
            ' Obtiene nMovNro
            lnMovNro = loReg.dGetnMovNro(psMovNro)
            
            '** Inserta MovCol
            Call loReg.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 0, lrsGasto!nMonto, 0, "", 0, lrs!nPrdEstado, lrs!nSaldo)
            
            '** Inserta MovColDet
            Call loReg.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 0, lrsGasto!nPrdConceptoCod, 0, lrsGasto!nMonto)
        End If
    End If
    Set lrs = Nothing
    Set lrsGasto = Nothing
    'END JUEZ **************************************************************************************
    

    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loReg.dCommitTrans
    mbTrans = False
Set loReg = Nothing

Exit Sub
ControlError:
    If mbTrans Then
        loReg.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nGastoRecupEliminaGasto", "Error en Funcion de Elimina Gasto Recuperaciones"

End Sub

Public Sub nCastigaCredRecup(ByVal psCtaCod As String, ByVal psMetliquid As String, _
        ByVal psFechaHora As String, ByVal psMovNro As String, ByVal pnSaldoCap As Double, _
        ByVal pnSaldoIntComp As Double, ByVal pnSaldoIntMora As Double, ByVal pnSaldoGasto As Double, _
        ByVal pnIntComGen As Double, ByVal pnDemanda As ColocRecDemandado, ByVal pnProvision As Double, _
        Optional pbEjecBatch As Boolean = False)

'** Actualiza Producto (NroTrans )
'** Actualiza ColocCredRecup (Metodo Liquid)
'** Actualiza Mov (Metodo Liquid)
'************************************

Dim lsSQL As String
Dim loRegRec As COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim lnGasto4 As Double, lnIngreso6 As Double
'Codigo de Operacion
If pnDemanda = gColRecDemandaNo Then
    lsOpeCod = gColRecOpeCastigSinDem     ' Cambiar  la constante
ElseIf pnDemanda = gColRecDemandaSi Then
    lsOpeCod = gColRecOpeCastigConDem
End If

' Determina Gastos o Ingresos
If pnSaldoCap > pnProvision Then
    lnGasto4 = pnSaldoCap - pnProvision
    lnIngreso6 = 0
ElseIf pnProvision > pnSaldoCap Then
    lnIngreso6 = pnProvision - pnSaldoCap
    lnGasto4 = 0
End If
        
'On Error GoTo ErrorModRec
Set loRegRec = New COMDColocPig.DCOMColPActualizaBD
    loRegRec.dBeginTrans
    mbTrans = True

    '** Actualiza Producto
    Call loRegRec.dUpdateProducto(psCtaCod, , , gColocEstRecVigCast, psFechaHora, -2, False)              ' (-2) aumenta el ste transac

    '** Actualiza Colocaciones
    Call loRegRec.dUpdateColocaciones(psCtaCod, , , , , psMovNro, , False)

    '** Actualiza ColocRecup
    '**Modificado por DAOR 20070827
    'Call loRegRec.dUpdateColocRecup(psCtaCod, , , pnSaldoIntComp, , , pnIntComGen, psMetLiquid, , , , False)
    Call loRegRec.dUpdateColocRecup(psCtaCod, , , pnSaldoIntComp, pnSaldoIntMora, pnSaldoGasto, pnIntComGen, psMetliquid, , , , False)
 
    '**DAOR 20070806**************************************************************
    Call loRegRec.dInsertColocacEstado(psCtaCod, psFechaHora, gColocEstRecVigCast, 0, pnSaldoCap, "Castigo de crédito", _
            0, 0, 0, 0, 0, 0, 0, 0, 0)
    '*****************************************************************************
    
    '** Inserta Mov
    Call loRegRec.dInsertMov(psMovNro, lsOpeCod, "Castigar Credito", gMovEstContabMovContable, gMovFlagVigente, False)

    ' Obtiene nMovNro
    lnMovNro = loRegRec.dGetnMovNro(psMovNro)

    '** Inserta MovCol
    Call loRegRec.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 0, pnSaldoCap, 0, "", 0, 0, pnSaldoCap, False)
    
    '** Inserta MovColDet - Capital
    Call loRegRec.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 0, gColRecConceptoCodCapital, 0, pnSaldoCap, False)
    
    If lnGasto4 > 0 Then  ' Ingresos
        Call loRegRec.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 0, gColRecConceptoCodCastigGasto4, 0, lnGasto4, False)
    End If
    
    If lnIngreso6 > 0 Then  ' Gastos
        Call loRegRec.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 0, gColRecConceptoCodCastigIngreso6, 0, lnIngreso6, False)
    End If
    
    ' Capital
    Call loRegRec.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 0, gColRecConceptoCodCastigCtaOrdenCap, 0, pnSaldoCap, False)
    
    ' Int Comp
    If pnSaldoIntComp > 0 Then
        Call loRegRec.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 0, gColRecConceptoCodCastigCtaOrdenIntCom, 0, pnSaldoIntComp, False)
    End If
    
    ' Int Morat
    If pnSaldoIntMora > 0 Then
        Call loRegRec.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 0, gColRecConceptoCodCastigCtaOrdenIntMor, 0, pnSaldoIntMora, False)
    End If
    
    If pnSaldoGasto > 0 Then
        Call loRegRec.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 0, gColRecConceptoCodCastigCtaOrdenGasto, 0, pnSaldoGasto, False)
    End If

    'mnEjecutaBatch = loRegRec.dEjecutaBatch
    loRegRec.dCommitTrans
    mbTrans = False
Set loRegRec = Nothing

Exit Sub
ErrorModRec:
    If mbTrans Then
        loRegRec.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nCambiaMetodoLiquidCredRecup", "Error en Funcion de Cambio Metodo Liquid Recuperaciones "

End Sub

Public Sub nRegistraCredEnRecup(ByVal psCtaCod As String, ByVal pnTasaInteres As Double, ByVal pnSaldo As Currency, _
        ByVal pnPlazo As Integer, ByVal psLineaCred, ByVal prClientes As ADODB.Recordset, ByVal psFecIngRecup As String, _
        ByVal pnSaldoIntComp As Currency, ByVal pnSaldoIntMor As Currency, ByVal pnSaldoGasto As Currency, _
        ByVal pnDemanda As Integer, ByVal psFechaHora As String, ByVal psMovNro As String, _
        Optional pbEjecBatch As Boolean = False)

'** Insert Producto
'** Insert Colocaciones
'** Insert ColocacEstado
'** Insert ColocRecup
'** Insert ColocRecupPersona
'** Insert ColocRecupEstado
'** Insert Mov
'** Insert MovCol
'************************************

Dim lsSQL As String
Dim loRegRec As COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
'Codigo de Operacion
lsOpeCod = gColPOpeImpDuplicado

'On Error GoTo ErrorModRec
Set loRegRec = New COMDColocPig.DCOMColPActualizaBD
    loRegRec.dBeginTrans
    mbTrans = True
    
    '** Inserta Producto
    Call loRegRec.dInsertProducto(psCtaCod, pnTasaInteres, pnSaldo, gColocEstRecVigJud, psFechaHora, 1, False)
    
    '** Inserta Colocaciones
    Call loRegRec.dInsertColocaciones(psCtaCod, 30, psFecIngRecup, pnSaldo, gColocCalendCodPFCF, psMovNro, psLineaCred, psFechaHora, False)
    
    '** Inserta ColocEstado
    Call loRegRec.dInsertColocacEstado(psCtaCod, psFechaHora, gColocEstRecVigJud, 1, pnSaldo, "Ingreso a Recuperaciones", gColocCalendCodPFCF, 0, 0, 0, 0, 0, 0, 0, 0, False)
    
    '** Inserta ColocRecup
    Call loRegRec.dInsertColocRecup(psCtaCod, psFecIngRecup, -1, pnSaldoIntComp, pnSaldoIntMor, pnSaldoGasto, 0, "GCIM", -1, pnDemanda, 1, False)
    
    '** Inserta ColocRecupEstado
    Call loRegRec.dInsertColocRecupEstado(psCtaCod, psFechaHora, gColocEstRecVigJud, "", False)
    
    '** Inserta ColocRecupExpediente
    'Call loRegRec.dInsertColocRecupExpediente(psCtaCod, psFechaHora, gColocEstRecVigJud, "", False)
    
    '** Inserta ColocRecupComisionHist
    'Call loRegRec.dInsertColocRecupExpediente(psCtaCod, psFechaHora, gColocEstRecVigJud, "", False)
    
    '** Inserta ColocRecupPersonas
    'inserta Clientes para cada registro de prClientes
    Do While Not prClientes.EOF
        Call loRegRec.dInsertProductoPersona(psCtaCod, prClientes("cPersCod"), prClientes("cPersRela"), False)
        prClientes.MoveNext
    Loop
    
    '** Inserta Mov
    Call loRegRec.dInsertMov(psMovNro, lsOpeCod, "Registro En Recuperaciones", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegRec.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegRec.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 1, pnSaldo, 0, "", 0, gColocEstRecVigJud, pnSaldo, False)
    
    '** Inserta MovColDet
    Call loRegRec.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 1, gColRecConceptoCodCapital, 0, pnSaldo, False)
    
    'mnEjecutaBatch = loRegRec.dEjecutaBatch
    loRegRec.dCommitTrans
    mbTrans = False
Set loRegRec = Nothing

Exit Sub
ErrorModRec:
    If mbTrans Then
        loRegRec.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nCambiaMetodoLiquidCredRecup", "Error en Funcion de Cambio Metodo Liquid Recuperaciones "

End Sub


Public Sub nRegistraExpedienteRecup(ByVal psCtaCod As String, ByVal psFechaHora As String, ByVal psMovNro As String, _
        ByVal pnComisionCod As Integer, ByVal pnTipCj As Integer, ByVal psNumExp As String, ByVal pnMonPetit As Currency, _
        ByVal pnMoneda As Integer, ByVal prClientes As ADODB.Recordset, ByVal psPetit As String, ByVal psHechos As String, _
        ByVal psFundJur As String, ByVal psMedProb As String, ByVal psDatComp As String, _
        ByVal pnViaProce As Integer, ByVal pnEstadoProceso As Integer, ByVal pbCambioComision As Boolean, _
        ByVal pbNuevoExpediente As Boolean, Optional pbEjecBatch As Boolean = False, Optional ByVal pnTipoProceso As Integer, _
        Optional ByVal psTipoProyecto As String, Optional pnDemanda As Integer, Optional pnTipMedCau As Integer = 0) 'DAOR 20070417, Se agregó parametro pnTipMedCau (Tipo de Medida cautelar)

'** Actualiza ColocRecup
'** Actualiza ColocRecupExpediente
'** Actualiza ColocRecupComisionHist
'** Inserta Mov
'************************************

Dim lsSQL As String
Dim loRegRec As COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
'Codigo de Operacion
lsOpeCod = 135001

'On Error GoTo ErrorModRec
Set loRegRec = New COMDColocPig.DCOMColPActualizaBD
    loRegRec.dBeginTrans
    mbTrans = True
    
    '** Actualiza ColocRecup
    Call loRegRec.dUpdateColocRecup(psCtaCod, , pnComisionCod, , , , , , pnTipCj, pnDemanda, , False)
    
    
    '** Actualiza ColocRecupExpediente
    If pbNuevoExpediente = False Then ' Existe Expediente
        If psTipoProyecto = "H" Then
            Call loRegRec.dInsertColocRecExpediente(psCtaCod, psNumExp, pnMonPetit, pnMoneda, _
                pnViaProce, pnEstadoProceso, False, pnTipoProceso)
        Else
            'Comentado por DAOR 20070417
            'Call loRegRec.dInsertColocRecExpediente(psCtaCod, psNumExp, pnMonPetit, pnMoneda, _
            '    pnViaProce, pnEstadoProceso, False)
            
            '**DAOR 20070417******************************************************
            Call loRegRec.dInsertColocRecExpediente(psCtaCod, psNumExp, pnMonPetit, pnMoneda, _
                pnViaProce, pnEstadoProceso, False, , pnTipMedCau)
            '*********************************************************************
        End If
    Else
        If psTipoProyecto = "H" Then
            Call loRegRec.dUpdateColocRecupExpediente(psCtaCod, psNumExp, pnMonPetit, pnMoneda, _
                pnViaProce, pnEstadoProceso, False, pnTipoProceso)
        Else
            'Comentado por DAOR 20070417
            'Call loRegRec.dUpdateColocRecupExpediente(psCtaCod, psNumExp, pnMonPetit, pnMoneda, _
            '    pnViaProce, pnEstadoProceso, False)
            
            '**DAOR 20070417*******************************************************
            Call loRegRec.dUpdateColocRecupExpediente(psCtaCod, psNumExp, pnMonPetit, pnMoneda, _
                pnViaProce, pnEstadoProceso, False, , pnTipMedCau)
            '**********************************************************************
        End If
    End If
    '** Registra el cambio de comision
    If pbCambioComision = True Then
        '** Actualiza ColocRecupComisionHist
        Call loRegRec.dInsertColocRecComisionHist(psCtaCod, pnComisionCod, lnMovNro, False)
    End If

    '** Borra las personas relacionadas
    Call loRegRec.dDeleteProductoPersona(psCtaCod, "", gColRelPersEstudioJuridico, False)
    Call loRegRec.dDeleteProductoPersona(psCtaCod, "", gColRelPersJuzgado, False)
    Call loRegRec.dDeleteProductoPersona(psCtaCod, "", gColRelPersJuez, False)
    Call loRegRec.dDeleteProductoPersona(psCtaCod, "", gColRelPersSecretario, False)
    '** Registra las personas relacionadas
    prClientes.MoveFirst
    Do While Not prClientes.EOF
        Call loRegRec.dInsertProductoPersona(psCtaCod, prClientes("cPersCod"), prClientes("cPersRelac"), False)
        prClientes.MoveNext
    Loop
    
    If Len(psPetit) > 0 Then
        Call loRegRec.dUpdateColocRecupExpedienteInf(psCtaCod, gColRecExpedTipoInfPetit, psPetit, False)
    End If
    If Len(psHechos) > 0 Then
        Call loRegRec.dUpdateColocRecupExpedienteInf(psCtaCod, gColRecExpedTipoInfHechos, psHechos, False)
    End If
    If Len(psFundJur) > 0 Then
        Call loRegRec.dUpdateColocRecupExpedienteInf(psCtaCod, gColRecExpedTipoInfFundJur, psFundJur, False)
    End If
    If Len(psMedProb) > 0 Then
        Call loRegRec.dUpdateColocRecupExpedienteInf(psCtaCod, gColRecExpedTipoInfMedProb, psMedProb, False)
    End If
    If Len(psDatComp) > 0 Then
        Call loRegRec.dUpdateColocRecupExpedienteInf(psCtaCod, gColRecExpedTipoInfDatComp, psDatComp, False)
    End If
    
    '** Inserta Mov
    Call loRegRec.dInsertMov(psMovNro, lsOpeCod, "Modifica Exped Judic", gMovEstContabNoContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegRec.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    'Call loRegRec.dInsertMovCol( lnMovNro, lsOpeCod, psCtaCod, 1, pnMontoTransac, 0, "", 0, False)
    
    'mnEjecutaBatch = loRegRec.dEjecutaBatch
    loRegRec.dCommitTrans
    mbTrans = False
Set loRegRec = Nothing

Exit Sub
ErrorModRec:
    If mbTrans Then
        loRegRec.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraExpedienteRecup", "Error en Funcion de Registro Expediente Recuperaciones "

End Sub

Public Sub nExtornoPagoCreditoRecup(ByVal psCtaCod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False, Optional ByRef psmensaje As String, _
        Optional ByVal psMotExtorno As Variant)

'** Actualiza Producto ( Estado )
'** Actualiza ColocRecup
'** Update ColocRecupGasto(Comision) <Flag ='X'>
'** Update ColocRecupGasto Estado , (-) Monto Pagado
'** Update ColocRecupGasto (Administrativo )
'** dInsertMov
'** dInsertMovRef
'************************************

Dim lsSQL As String, lsSQL2 As String, lsSQL3 As String
Dim loBase As COMDColocPig.DCOMColPActualizaBD

Dim lnMovNro As Long
Dim lsOpeCod As String
Dim loDataExt As COMConecta.DCOMConecta
Dim lrDataExt As ADODB.Recordset
Dim lrDataG As ADODB.Recordset
Dim loFunciones As COMNContabilidad.NCOMContFunciones

Dim lnSaldCap As Currency, lnSaldIntComp As Currency, lnSaldIntMor As Currency, lnSaldGasto As Currency
Dim lnCapPag As Currency, lnIntComPag As Currency, lnIntMorPag As Currency, lnGastoPag As Currency
Dim lsUltMov As String
'**DAOR 20070809***************
Dim lnIntMoraGenerado As Double
Dim lnMora As Double
Dim lnComiAbo As Double
Dim ldFecha As Date
Dim lnGasto As Double
'*******************************
'19-05-2006
Dim nEstadoAnterior As Integer

lsOpeCod = gColRecOpeExtPagRecup



' *** Obtiene Datos de Montos Cubiertos

'**Modificado por DAOR 20070809, se agrego el campo nIntMoraGen
'JUEZ 20141213 Se agregó cOpeCod = mc.cOpeCod en las consultas a MovColDet para obtener sólo montos del pago, mas no de la condonacion
lsSQL = "SELECT p.nSaldo, cr.nSaldoIntComp, cr.nSaldoIntMor, cr.nSaldoGasto, cr.nIntCompGen,cr.nIntMoraGen,MC.nCredEstado, " _
    & " CapPag = (Select ISNULL(SUM(nMonto),0) From MovColDet mcd  Where mcd.nMovNro = " & pnMovNroAnt _
    & "           And mc.cCtaCod = '" & psCtaCod & "' and cOpeCod = mc.cOpeCod and mcd.nPrdConceptoCod in (" & gColRecConceptoCodCapital & " )  ) , " _
    & " IntComPag = (Select ISNULL(SUM(nMonto),0) From MovColDet mcd  Where mcd.nMovNro = " & pnMovNroAnt _
    & "           And mc.cCtaCod =  '" & psCtaCod & "' and cOpeCod = mc.cOpeCod and mcd.nPrdConceptoCod in (" & gColRecConceptoCodInteresCompensatorio & " )  ) , " _
    & " IntMorPag = (Select ISNULL(SUM(nMonto),0) From MovColDet mcd  Where mcd.nMovNro = " & pnMovNroAnt _
    & "           And mc.cCtaCod =  '" & psCtaCod & "' and cOpeCod = mc.cOpeCod and mcd.nPrdConceptoCod in (" & gColRecConceptoCodInteresMoratorio & " )  ) , " _
    & " GastoPag = (Select ISNULL(SUM(nMonto),0) From MovColDet mcd  Where mcd.nMovNro = " & pnMovNroAnt _
    & "           And mc.cCtaCod =  '" & psCtaCod & "' and cOpeCod = mc.cOpeCod and mcd.nPrdConceptoCod like '32%'  ),  " _
    & " cUltMov = (SELECT MAX(M.cMOVNRO) FROM MOV M JOIN MOVCOL MC1 ON MC1.NMOVNRO = M.NMOVNRO " _
    & "             WHERE  M.NMOVNRO <" & pnMovNroAnt & " AND M.NMOVFLAG =0 AND MC1.CCTACOD ='" & psCtaCod & "'), " _
    & " nTasaIntMor = (SELECT ISNULL(nTasaInteres, 0) From ProductoTasaInteres PT  WHERE PT.cCtaCod = P.cCtaCod and PT.nPrdTasaInteres = " & gColocLineaCredTasasIntMoratNormal & " ),  " _
    & " C.cUltimaActualizacion" _
    & " From MovCol mc " _
    & " Inner join Producto p on p.cCtaCod = mc.cCtaCod " _
    & " Inner join Colocaciones c on p.cCtaCod = c.cCtaCod " _
    & " Inner join ColocRecup cr on p.cCtaCod = cr.cCtaCod " _
    & " Where mc.cCtaCod = '" & psCtaCod & "' And mc.nmovnro =  " & pnMovNroAnt & " and mc.cOpeCod like '13[06]%' " 'JUEZ 20130921 Se agregó el filtro "13[06]%" para obtener sólo datos del pago
    'in (" & gColRecConceptoCodGastoCobranza & " )
lsSQL2 = "SELECT distinct mcd.nPrdConceptoCod, mcd.nNroCuota nNroGastoCta, mcd.nMonto nMontoPag, " _
    & " nMontoGasto = (Select CRG.nMontoPagado From ColocRecupGastos CRG Where CRG.cCtaCod = '" & psCtaCod & "' And CRG.nNroGastoCta = mcd.nNroCuota )" _
    & " FROM Mov m INNER JOIN MovCol mc on m.nMovNro = mc.nMovNro " _
    & " INNER JOIN MovColDet mcd on mc.nMovNro = mcd.nMovNro " _
    & " WHERE mc.cCtaCod = '" & psCtaCod & "'  and m.nmovnro = " & pnMovNroAnt & " " _
    & " AND mcd.nPrdConceptoCod like '32%' "
    'AND mcd.nPrdConceptoCod in (" & gColRecConceptoCodGastoCobranza & "," & gColRecConceptoCodGastoComisionAbog & "," & gColRecConceptoCodGastoAdmin & " ) "
Set loDataExt = New COMConecta.DCOMConecta
    loDataExt.AbreConexion
    '** Carga Conceptos Pagados
    Set lrDataExt = loDataExt.CargaRecordSet(lsSQL)
    If lrDataExt Is Nothing Then
        psmensaje = "ERROR: al Buscar datos para Extorno "
        Exit Sub
    End If
    If lrDataExt.BOF And lrDataExt.EOF Then
        psmensaje = "ERROR: No encontro datos de la Operacion a extornar "
        Exit Sub
    End If
    ' Asigna los valores
    lnSaldCap = lrDataExt!nSaldo
    lnSaldIntComp = lrDataExt!nSaldoIntComp - lrDataExt!nIntCompGen
    lnSaldIntMor = lrDataExt!nSaldoIntMor
    lnSaldGasto = lrDataExt!nSaldoGasto
    lnCapPag = lrDataExt!CapPag
    lnIntComPag = lrDataExt!IntComPag
    lnIntMorPag = lrDataExt!IntMorPag
    lnGastoPag = lrDataExt!GastoPag
    lsUltMov = IIf(IsNull(lrDataExt!cUltMov), "@", lrDataExt!cUltMov)
    
    lnIntMoraGenerado = lrDataExt!nIntMoraGen 'DAOR 20070809
    
'*********** Nuevos Cambios para el extorno AVMM 27-04-2007 *********************
'Proceso que permite calcular la mora generada apartir del saldo capital

'**Comentado por DAOR 20070809, ya no es necesario, el Interes Moratorio generado se obtiene
'**de la base de datos, campo nIntMoraGen *****************************

'Dim loParam As COMDConstSistema.NCOMConstSistema
'Dim loCalcula As NCOMColRecCalculos
'Dim oFunFecha As New COMFunciones.FCOMFechas
'Dim lnTipoCalcIntMora As String, lnFormaCalcIntMora As String
'Dim lnIntMoraGenerado As Double
'Dim lnTasaMora As Double
'Dim lnSaldoCap As Double
'Dim lnDiasUltTrans As Integer
'Dim lsFecUltPago As String
'Dim lnMora As Double
'Dim lnComiAbo As Double
'Dim ldFecha As Date
'Dim lnGasto As Double
'Set loCalcula = New NCOMColRecCalculos
'Set loParam = New COMDConstSistema.NCOMConstSistema
'   lnTipoCalcIntMora = loParam.LeeConstSistema(152)
'   lnFormaCalcIntMora = loParam.LeeConstSistema(203)
'Set loParam = Nothing
'
'lnTasaMora = lrDataExt!nTasaIntMor
'lnSaldoCap = lnSaldCap + lnCapPag
'lsFecUltPago = CDate(oFunFecha.fgFechaHoraGrab(lrDataExt!cUltimaActualizacion))
'lnDiasUltTrans = CDate(Format(psFechaHora, "dd/mm/yyyy")) - CDate(Format(lsFecUltPago, "dd/mm/yyyy"))
'
'If lnTipoCalcIntMora = 0 Then  ' NoCalcula
'    lnIntMoraGenerado = 0
'ElseIf lnTipoCalcIntMora = 1 Then ' En base al capital
'    If lnFormaCalcIntMora = 1 Then 'INTERES COMPUESTO
'        lnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGenerado(lnDiasUltTrans, lnTasaMora, lnSaldoCap)
'    Else
'        'INTERES SIMPLE
'        lnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGeneradoICA(lnDiasUltTrans, lnTasaMora, lnSaldoCap)
'    End If
'End If
'***********************************************************************

'Proceso para el verdadero valor que debe de ir en nSaldoIntMor de la Tabla ColocRecup***
lnMora = (lnSaldIntMor + lnIntMorPag) - lnIntMoraGenerado
lnSaldIntMor = lnMora

'Proceso Para Obeneter el monto del Abogado
ldFecha = CDate(Format(psFechaHora, "dd/mm/yyyy"))
'**Modificado por DAOR 20070809
'lsSQL3 = "select top 1 nMonto from ColocRecupGastos where nPrdConceptoCod=" & gColRecConceptoCodGastoComisionAbog & " and datediff(day,dAsigna,'" & Format(ldFecha, "yyyy/mm/dd") & "')=0 "
lsSQL3 = "select top 1 nMonto from ColocRecupGastos where nPrdConceptoCod=" & gColRecConceptoCodGastoComisionAbog & " and datediff(day,dAsigna,'" & Format(ldFecha, "yyyy/mm/dd") & "')=0 and cCtaCod='" & psCtaCod & "'"
Set lrDataG = New ADODB.Recordset
Set lrDataG = loDataExt.CargaRecordSet(lsSQL3)
    If Not (lrDataG.EOF And lrDataG.BOF) Then
       lnComiAbo = lrDataG("nMonto")
    Else
       lnComiAbo = 0
    End If
Set lrDataG = Nothing
    
lnGasto = (lnSaldGasto + lnGastoPag) - lnComiAbo
lnSaldGasto = lnGasto
'***************************************************************************************
    '19-05-2006
    nEstadoAnterior = lrDataExt!nCredEstado
    
    
    Set lrDataExt = Nothing
    '** Carga los Gastos Cobrados ' Concepto / NroGasto / MontoPagado
    Set lrDataExt = loDataExt.CargaRecordSet(lsSQL2)
    
    
    
Set loDataExt = Nothing

' *** Realiza el Extorno
'On Error GoTo Error
Set loBase = New COMDColocPig.DCOMColPActualizaBD
    loBase.dBeginTrans
    mbTrans = True
    
    '** Update Producto ---------> Actualizamos el Estado
    Call loBase.dUpdateProducto(psCtaCod, , lnSaldCap + lnCapPag, nEstadoAnterior, psFechaHora, -2, False)
    
    '** Update COLOCACIONES
    Call loBase.dUpdateColocaciones(psCtaCod, , , , , lsUltMov, , False)
    
    '** Update ColocRecup
    'Call loBase.dUpdateColocRecup(psCtaCod, , , lnSaldIntComp + lnIntComPag, lnSaldIntMor + lnIntMorPag, lnSaldGasto + lnGastoPag, 0, , , , , False)
    Call loBase.dUpdateColocRecup(psCtaCod, , , lnSaldIntComp + lnIntComPag, lnSaldIntMor, lnSaldGasto, 0, , , , , False)
    
    'Regresa los Gastos a su estado anterior
    Do While Not lrDataExt.EOF
        If lrDataExt!nPrdConceptoCod = gColRecConceptoCodGastoComisionAbog Then
            Call loBase.dUpdateColocRecupGastos(psCtaCod, lrDataExt!nNroGastoCta, , , , gColRecGastoEstEliminado, , False)
        ElseIf lrDataExt!nPrdConceptoCod = gColRecConceptoCodGastoComisionAbog Then
            Call loBase.dUpdateColocRecupGastos(psCtaCod, lrDataExt!nNroGastoCta, , , , gColRecGastoEstEliminado, , False)
        Else
            If lrDataExt!nMontoGasto = 0 Then
                Call loBase.dUpdateColocRecupGastos(psCtaCod, lrDataExt!nNroGastoCta, , , lrDataExt!nMontoGasto, gColRecGastoEstPendiente, , False)
            Else
                Call loBase.dUpdateColocRecupGastos(psCtaCod, lrDataExt!nNroGastoCta, , , lrDataExt!nMontoGasto - lrDataExt!nMontoPag, gColRecGastoEstPendiente, , False)
            End If
        End If
        lrDataExt.MoveNext
    Loop
    
    '** Inserta Mov
    Call loBase.dInsertMov(psMovNro, lsOpeCod, "Ext.Pago.Credito.Recup", gMovEstContabNoContable, gMovFlagDeExtorno, False)
            
    ' Obtiene nMovNro
    lnMovNro = loBase.dGetnMovNro(psMovNro)
    
    '***CTI3 (ferimoro) 17102018
    Call loBase.dInsertDetExtMov(lnMovNro, lsOpeCod, 0, psCtaCod, False, psMotExtorno(0), psMotExtorno(1))
    
    '** Inserta MovCol
    Call loBase.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 0, pnMonto, 0, "", 0, 0, 0, False)
    
    '** Update Mov Anterior
    Call loBase.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loBase.dInsertMovRef(lnMovNro, pnMovNroAnt, False)
    
    '**
    'mnEjecutaBatch = lobase.dEjecutaBatch
    
    'RIRO20140610 ERS017 *********
    loBase.actualizarRelacionExtornoVoucherCaptacion pnMovNroAnt
    loBase.ActualizaMovPendientesOpe pnMovNroAnt, pnMonto
    'END RIRO ********************
    loBase.dExtornoLiberarGarantia pnMovNroAnt, psCtaCod 'EJVG20150822
    
    loBase.dCommitTrans
    mbTrans = False
Set loBase = Nothing

Exit Sub
Error:
    If mbTrans Then
        loBase.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nExtornoTransFRecup", "Error en Funcion de Extorno de Transferencia"

End Sub



Public Sub nCancelacionCreditoRecup(ByVal psCtaCod As String, ByVal psFechaHora As String, ByVal psOpeCod As String, _
        ByVal psMovNro As String, ByVal pnPrdEstado As Integer, _
        ByVal pnNewSaldoCap As Currency, ByVal pnNewSaldoIntCom As Currency, _
        ByVal pnNewIntComGener As Currency, ByVal pnNroCalend As Integer, _
        Optional pbEjecBatch As Boolean = False, Optional ByVal psComentario As String = "")

'** Update Producto
'** Update Colocaciones
'** Update ColocRecup
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loReg As COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String


lsOpeCod = psOpeCod

'On Error GoTo ControlError
Set loReg = New COMDColocPig.DCOMColPActualizaBD
    loReg.dBeginTrans
    mbTrans = True
    
    '** Actualiza Producto
    Call loReg.dUpdateProducto(psCtaCod, , , pnPrdEstado, psFechaHora, -2, False)               ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loReg.dUpdateColocaciones(psCtaCod, , , , , psMovNro, , False)
    
    '** Actualiza ColocRecup
    Call loReg.dUpdateColocRecup(psCtaCod, , , pnNewSaldoIntCom, , , pnNewIntComGener, , , , , False, psComentario)

    '** Inserta Mov
    Call loReg.dInsertMov(psMovNro, lsOpeCod, "Cancelacion Cred Recup", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loReg.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loReg.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, pnNewSaldoCap, 0, "", 0, 0, pnNewSaldoCap, False)
    
    '** Inserta MovColDet - Capital
    'If pnCapPag > 0 Then
    '    Call loReg.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gColRecConceptoCodCapital, 0, pnCapPag, False)
    'End If
    
    'mnEjecutaBatch = loReg.dEjecutaBatch
    loReg.dCommitTrans
    mbTrans = False
Set loReg = Nothing

Exit Sub
ControlError:
    If mbTrans Then
        loReg.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nCancelacionCreditoRecup", "Error en Funcion Cancelacion Credito Recup "

End Sub


Public Sub nCierreMesRecuperaciones(ByVal psFechaHora As String, ByVal psOpeCod As String, _
        ByVal psMovNro As String, ByVal pnTipoCalcIntComp As Integer, ByVal pnTipoCalcIntMora As Integer, _
        Optional pbEjecBatch As Boolean = False, Optional psmensaje As String)

'** Actualiza Producto (NroTrans )
'** Actualiza ColocCredRecup (Metodo Liquid)
'************************************

Dim lsSQL As String
Dim loRegRec As COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim loRec As COMDColocRec.DCOMColRecCredito
Dim lrRec As ADODB.Recordset
Dim loCalcula As NCOMColRecCalculos
Dim oFun
Dim lsFecUltPago As String
Dim lnDiasUltTrans As Integer
Dim oFec As New COMFunciones.FCOMFechas

Dim lnNewSaldoIntCom As Currency, lnNewSaldoIntMor As Currency, lnNewSaldoIntComGener As Currency
Dim lnIntCompGenerado As Currency, lnIntMoraGenerado As Currency
Dim lsmensaje As String
'Codigo de Operacion
lsOpeCod = "131000"

'On Error GoTo ErrorModRec


    Set loRec = New COMDColocRec.DCOMColRecCredito
        Set lrRec = loRec.dObtieneDatosCierreMes(lsmensaje)
        psmensaje = lsmensaje
    Set loRec = Nothing
    If lrRec Is Nothing Then Exit Sub
    If lrRec.BOF And lrRec.EOF Then Exit Sub
    Set loRegRec = New COMDColocPig.DCOMColPActualizaBD
    loRegRec.dBeginTrans
    mbTrans = True
    
    Do While Not lrRec.EOF
        lnIntCompGenerado = 0
        lnIntMoraGenerado = 0
        lsFecUltPago = CDate(oFec.fgFechaHoraGrab(lrRec!cUltimaActualizacion))
                            
        'lnDiasUltTrans = CDate(Format(gdFecSis, "dd/mm/yyyy")) - CDate(Format(fsFecUltPago, "dd/mm/yyyy"))
        lnDiasUltTrans = CDate(oFec.fgFechaHoraGrab(psMovNro)) - CDate(Format(lsFecUltPago, "dd/mm/yyyy"))
        
        'Calcula los intereses Generados
        Set loCalcula = New NCOMColRecCalculos
            If pnTipoCalcIntComp = 0 Then ' NoCalcula
                lnIntCompGenerado = 0
            ElseIf pnTipoCalcIntComp = 1 Then ' En base al capital
                lnIntCompGenerado = loCalcula.nCalculaIntCompGeneradoICA(lnDiasUltTrans, lrRec!nTasaIntComp, lrRec!nSaldo)
            ElseIf pnTipoCalcIntComp = 2 Then ' En base a capit + int Comp
                lnIntCompGenerado = loCalcula.nCalculaIntCompGeneradoICA(lnDiasUltTrans, lrRec!nTasaIntComp, lrRec!nSaldo + lrRec!nSaldoIntComp)
            ElseIf pnTipoCalcIntComp = 3 Then ' En base a capit + int Comp + int Morat
                lnIntCompGenerado = loCalcula.nCalculaIntCompGeneradoICA(lnDiasUltTrans, lrRec!nTasaIntComp, lrRec!nSaldo + lrRec!nSaldoIntComp + lrRec!nSaldoIntMor)
            End If
            If pnTipoCalcIntMora = 0 Then  ' NoCalcula
                lnIntMoraGenerado = 0
            ElseIf pnTipoCalcIntMora = 1 Then ' En base al capital
                lnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGeneradoICA(lnDiasUltTrans, lrRec!nTasaIntMorat, lrRec!nSaldo)
            ElseIf pnTipoCalcIntMora = 2 Then ' En base a capit + int Comp
                lnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGeneradoICA(lnDiasUltTrans, lrRec!nTasaIntMorat, lrRec!nSaldo + lrRec!nSaldoIntComp)
            ElseIf pnTipoCalcIntMora = 3 Then ' En base a capit + int Comp + int Morat
                lnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGeneradoICA(lnDiasUltTrans, lrRec!nTasaIntMorat, lrRec!nSaldo + lrRec!nSaldoIntComp + lrRec!nSaldoIntMor)
            End If
        Set loCalcula = Nothing
        
        lnNewSaldoIntComGener = lrRec!nIntCompGen + Format(lnIntCompGenerado, "#.00") + Format(lnIntMoraGenerado, "#.00")
        lnNewSaldoIntCom = lrRec!nSaldoIntComp + Format(lnIntCompGenerado, "#.00")
        lnNewSaldoIntMor = lrRec!nSaldoIntMor + Format(lnIntMoraGenerado, "#.00")
        '******
        
        '** Actualiza Producto
        Call loRegRec.dUpdateProducto(lrRec!cCtaCod, , , , psFechaHora, -2, False)           ' (-2) aumenta el ste transac
        
        '** Actualiza Colocaciones
        Call loRegRec.dUpdateColocaciones(lrRec!cCtaCod, , , , , psMovNro, , False)
    
        '** Actualiza ColocRecup
        Call loRegRec.dUpdateColocRecup(lrRec!cCtaCod, , , lnNewSaldoIntCom, lnNewSaldoIntMor, , lnNewSaldoIntComGener, , , , , False)
    
        '** Inserta Mov
        Call loRegRec.dInsertMov(psMovNro, lsOpeCod, "Cierre Mensual", gMovEstContabNoContable, gMovFlagVigente, False)
        
        ' Obtiene nMovNro
        lnMovNro = loRegRec.dGetnMovNro(psMovNro)
        
        '** Inserta MovCol
        Call loRegRec.dInsertMovCol(lnMovNro, lsOpeCod, lrRec!cCtaCod, 0, lnNewSaldoIntComGener, 0, "", 0, 0, 0, False)
        
        '** Inserta MovColDet - Int Comp
        If lnIntCompGenerado > 0 Then
            Call loRegRec.dInsertMovColDet(lnMovNro, lsOpeCod, lrRec!cCtaCod, 0, gColRecConceptoCodInteresCompensatorio, 0, lnIntCompGenerado, False)
        End If
        
        '** Inserta MovColDet - Int Comp
        If lnIntMoraGenerado > 0 Then
            Call loRegRec.dInsertMovColDet(lnMovNro, lsOpeCod, lrRec!cCtaCod, 0, gColRecConceptoCodInteresMoratorio, 0, lnIntMoraGenerado, False)
        End If
        
        'mnEjecutaBatch = loRegRec.dEjecutaBatch
        
        lrRec.MoveNext
    Loop
    
    
    loRegRec.dCommitTrans
    mbTrans = False

Exit Sub
ErrorModRec:
    If mbTrans Then
        loRegRec.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nCambiaMetodoLiquidCredRecup", "Error en Funcion de Cambio Metodo Liquid Recuperaciones "

End Sub

'** Modificado por DAOR 20070124
'** Se aumentó el parametro tipo de actuación procesal
Public Sub nRegistraActProcesales(ByVal psCtaCod As String, ByVal psFechaHora As String, ByVal psMovNro As String, _
        ByVal psComenta As String, Optional pbEjecBatch As Boolean = False, Optional pdFechaA As Date, Optional pdFechaV As Date, _
        Optional gsProyectoActual As String, Optional pbAlerta As Integer, Optional pnTipoAct As Integer)

'** Inserta Mov
'** Inserta ColocRecupActProc
'************************************

Dim lsSQL As String
Dim loRegRec As COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
'Codigo de Operacion
lsOpeCod = 135003

'On Error GoTo ErrorModRec
Set loRegRec = New COMDColocPig.DCOMColPActualizaBD
    loRegRec.dBeginTrans
    mbTrans = True
    
    '** Inserta Mov
    Call loRegRec.dInsertMov(psMovNro, lsOpeCod, "Modifica Exped Judic", gMovEstContabNoContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegRec.dGetnMovNro(psMovNro)
    
     If gsProyectoActual = "H" Then
        '** Inserta ColocRecupActProcesales
        Call loRegRec.dInsertColocRecupActProcesales(psCtaCod, psMovNro, psComenta, False, pdFechaA, pdFechaV, pbAlerta, gsProyectoActual)
     Else
        '** Inserta ColocRecupActProcesales
        Call loRegRec.dInsertColocRecupActProcesales(psCtaCod, psMovNro, psComenta, False, pdFechaA, pdFechaV, , gsProyectoActual, pnTipoAct)   'ARCV 24012007
        'Call loRegRec.dInsertColocRecupActProcesales(psCtaCod, psMovNro, psComenta, False, pdFechaA, pdFechaV, , gsProyectoActual)
     End If
    'mnEjecutaBatch = loRegRec.dEjecutaBatch
    loRegRec.dCommitTrans
    mbTrans = False
Set loRegRec = Nothing

Exit Sub
ErrorModRec:
    If mbTrans Then
        loRegRec.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraActProcesales", "Error en Funcion de Registro Actuaciones Procesales"

End Sub

Public Sub nRegistraExpedienteRecupRelacionados(ByVal psCtaCod As String, ByVal psMovNro As String, _
        ByVal pnComisionCod As Integer, ByVal psNumExp As String, ByVal pnMonPetit As Currency, _
        ByVal pnMoneda As Integer, ByVal prClientes As ADODB.Recordset, ByVal psPetit As String, ByVal psHechos As String, _
        ByVal psFundJur As String, ByVal psMedProb As String, ByVal psDatComp As String, _
        ByVal pnViaProce As Integer, ByVal pnEstadoProceso As Integer, ByVal pbCambioComision As Boolean, _
        ByVal pbNuevoExpediente As Boolean, Optional pbEjecBatch As Boolean = False, Optional pnNroCorrelativo As Integer, Optional ByVal pnTipoProceso As Integer, Optional ByVal pnTipoProcesal As Integer)

'** Actualiza ColocRecup
'** Actualiza ColocRecupExpediente
'** Actualiza ColocRecupComisionHist
'** Inserta Mov
'************************************

Dim lsSQL As String
Dim loRegRec As COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
'Codigo de Operacion
lsOpeCod = 135001

'On Error GoTo ErrorModRec
Set loRegRec = New COMDColocPig.DCOMColPActualizaBD
    loRegRec.dBeginTrans
    mbTrans = True
    
    '** Actualiza ColocRecupExpediente
    If pbNuevoExpediente = False Then ' Existe Expediente
        Call loRegRec.dInsertColocRecExpedienteRelacionado(psCtaCod, psNumExp, pnMonPetit, pnMoneda, _
                pnViaProce, pnEstadoProceso, False, pnNroCorrelativo, pnTipoProceso, pnTipoProcesal)
    Else
        Call loRegRec.dUpdateColocRecupExpedienteRelacionado(psCtaCod, psNumExp, pnMonPetit, pnMoneda, _
                pnViaProce, pnEstadoProceso, False, pnNroCorrelativo, pnTipoProceso, pnTipoProcesal)
    End If
    ' Obtiene nMovNro
    lnMovNro = loRegRec.dGetnMovNro(psMovNro)
    
    '** Borra las personas relacionadas
    Call loRegRec.dDeleteExpedientePersona(psCtaCod, "", gColRelPersEstudioJuridico, False)
    Call loRegRec.dDeleteExpedientePersona(psCtaCod, "", gColRelPersJuzgado, False)
    Call loRegRec.dDeleteExpedientePersona(psCtaCod, "", gColRelPersJuez, False)
    Call loRegRec.dDeleteExpedientePersona(psCtaCod, "", gColRelPersSecretario, False)
    Call loRegRec.dDeleteExpedientePersona(psCtaCod, "", gColRelPersEncargado, False)
    
    '** Registra las personas relacionadas
    prClientes.MoveFirst
    Do While Not prClientes.EOF
        Call loRegRec.dInsertExpedientePersona(psCtaCod, prClientes("cPersCod"), pnNroCorrelativo, prClientes("cPersRelac"))
        prClientes.MoveNext
    Loop
    
    If Len(psPetit) > 0 Then
        Call loRegRec.dUpdateColocRecupExpedienteInfRelacionado(psCtaCod, gColRecExpedTipoInfPetit, psPetit, False, pnNroCorrelativo)
    End If
    If Len(psHechos) > 0 Then
        Call loRegRec.dUpdateColocRecupExpedienteInfRelacionado(psCtaCod, gColRecExpedTipoInfHechos, psHechos, False, pnNroCorrelativo)
    End If
    If Len(psFundJur) > 0 Then
        Call loRegRec.dUpdateColocRecupExpedienteInfRelacionado(psCtaCod, gColRecExpedTipoInfFundJur, psFundJur, False, pnNroCorrelativo)
    End If
    If Len(psMedProb) > 0 Then
        Call loRegRec.dUpdateColocRecupExpedienteInfRelacionado(psCtaCod, gColRecExpedTipoInfMedProb, psMedProb, False, pnNroCorrelativo)
    End If
    If Len(psDatComp) > 0 Then
        Call loRegRec.dUpdateColocRecupExpedienteInfRelacionado(psCtaCod, gColRecExpedTipoInfDatComp, psDatComp, False, pnNroCorrelativo)
    End If
    
    '** Inserta Mov
    Call loRegRec.dInsertMov(psMovNro, lsOpeCod, "Modifica Exped Judic", gMovEstContabNoContable, gMovFlagVigente, False)
    
  
    '** Inserta MovCol
    'Call loRegRec.dInsertMovCol( lnMovNro, lsOpeCod, psCtaCod, 1, pnMontoTransac, 0, "", 0, False)
    
    'mnEjecutaBatch = loRegRec.dEjecutaBatch
    loRegRec.dCommitTrans
    mbTrans = False
Set loRegRec = Nothing

Exit Sub
ErrorModRec:
    If mbTrans Then
        loRegRec.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraExpedienteRecupRelacionados", "Error en Funcion de Registro Expediente Recuperaciones "

End Sub



Public Sub nExtornoTransfRecup(ByVal psCtaCod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False, Optional ByRef psmensaje As String, _
        Optional ByVal psMotExtorno As Variant)

'** Actualiza Producto ( Estado )
'** Elimina ColocRecupExpedInf
'** Elimina ColocRecupExped
'** Elimina ColocRecupGasto
'** Elimina ColocRecup
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loBase As COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim loDataExt As COMConecta.DCOMConecta
Dim lrDataExt As ADODB.Recordset
Dim loFunciones As COMNContabilidad.NCOMContFunciones

lsOpeCod = gColRecOpeExtTransfRecup

' *** Obtiene Datos para el Extorno

lsSQL = "SELECT CE.nPrdEstado " _
      & "FROM ColocacEstado CE WHERE CE.cCtaCod = '" & psCtaCod & "' " _
      & "AND CE.nPrdEstado <> " & gColocEstRecVigJud & "   " _
      & "ORDER BY CE.dPrdEstado DESC "

Set loDataExt = New COMConecta.DCOMConecta
    loDataExt.AbreConexion
    Set lrDataExt = loDataExt.CargaRecordSet(lsSQL)
Set loDataExt = Nothing
If lrDataExt Is Nothing Then
    psmensaje = "ERROR: al Buscar datos para Extorno "
    Exit Sub
End If
If lrDataExt.BOF And lrDataExt.EOF Then
    psmensaje = "ERROR: No encontro datos de la Operacion a extornar "
    Exit Sub
End If
' Asigna los valores
Dim lnAntEstado As Integer

lnAntEstado = lrDataExt!nPrdEstado

' *** Realiza el Extorno
'On Error GoTo Error
Set loBase = New COMDColocPig.DCOMColPActualizaBD
    loBase.dBeginTrans
    mbTrans = True
    
    '** Update Producto
    Call loBase.dUpdateProducto(psCtaCod, , , lnAntEstado, psFechaHora, -2, False)
    '** Delete ColocRecupExpedienteInf
    Call loBase.dDeleteColocRecupExpedienteInf(psCtaCod, , False)
    '** Delete ColocRecupExpediente
    Call loBase.dDeleteColocRecupExpediente(psCtaCod, False)
    '** Delete ColocRecupGastos
    Call loBase.dDeleteColocRecupGastos(psCtaCod, , False)
    '** Delete ColocRecup
    Call loBase.dDeleteColocRecup(psCtaCod, False)
    
    '** Insert ColocEstado
    Call loBase.dInsertColocacEstado(psCtaCod, psFechaHora, lnAntEstado, 1, pnMonto, "Extorno Transferencia Recup", _
          0, 0, 0, 0, 0, 0, 0, 0, 0, False)

    '** Inserta Mov
    Call loBase.dInsertMov(psMovNro, lsOpeCod, "Ext Transferencia Recup", gMovEstContabNoContable, gMovFlagDeExtorno, False)
    '***CTI3 (ferimoro) 17102018
    Call loBase.dInsertDetExtMov(psMovNro, lsOpeCod, 0, psCtaCod, False, psMotExtorno(0), psMotExtorno(1))
    ' Obtiene nMovNro
    lnMovNro = loBase.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loBase.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 0, pnMonto, 0, "", 0, 0, 0, False)

    '** Update Mov Anterior
    Call loBase.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loBase.dInsertMovRef(lnMovNro, pnMovNroAnt, False)
    
    '**
    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    loBase.dCommitTrans
    mbTrans = False
Set loBase = Nothing

Exit Sub
Error:
    If mbTrans Then
        loBase.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nExtornoTransFRecup", "Error en Funcion de Extorno de Transferencia"

End Sub
'ALPA 20081009***********************************************************************************
'''Public Sub nPagoCreditoRecup(ByVal psCtaCod As String, ByVal psFechaHora As String, ByVal psOpeCod As String, _
'''        ByVal psMovNro As String, ByVal pnMontoPagar As Currency, ByVal psMetLiquid As String, _
'''        ByVal pnNewSaldoCap As Currency, ByVal pnNewSaldoIntCom As Currency, _
'''        ByVal pnNewSaldoIntMor As Currency, ByVal pnNewSaldoGasto As Currency, _
'''        ByVal pnNewSaldoIntComGener As Currency, ByVal pnComisionAbog As Currency, ByVal pnNroGastoCta As Integer, _
'''        ByVal pnGastoAdmin As Currency, ByVal pnCapPag As Currency, _
'''        ByVal pnIntCompPag As Currency, ByVal pnIntMoratPag As Currency, _
'''        ByVal pnGastoPag As Currency, ByVal pmGastos As Variant, ByVal pnNroCalend As Integer, ByVal pnEstadoNew As ColocEstado, _
'''        Optional pbEjecBatch As Boolean = False, Optional psPersLavDinero As String = "", Optional ByVal pnITF As Double = 0#, Optional pnDocTpo As Integer = 0, Optional psNroDoc As String = "", _
'''        Optional ByVal pOpeITFCheqEfec As String = "", Optional ByVal psPersCmac As String, _
'''        Optional ByVal pnNewSaldoIntMorGener As Currency = 0, _
'''        Optional psTitPersLavDinero As String = "", Optional psOrdPersLavDinero As String = "", _
'''        Optional psReaPersLavDinero As String = "", Optional psBenPersLavDinero As String = "", _
'''        Optional psVisPersLavDinero As String = "")
Public Sub nPagoCreditoRecup(ByVal psCtaCod As String, ByVal psFechaHora As String, ByVal psOpeCod As String, _
        ByVal psMovNro As String, ByVal pnMontoPagar As Currency, ByVal psMetliquid As String, _
        ByVal pnNewSaldoCap As Currency, ByVal pnNewSaldoIntCom As Currency, _
        ByVal pnNewSaldoIntMor As Currency, ByVal pnNewSaldoGasto As Currency, _
        ByVal pnNewSaldoIntComGener As Currency, ByVal pnComisionAbog As Currency, ByVal pnNroGastoCta As Integer, _
        ByVal pnGastoAdmin As Currency, ByVal pnCapPag As Currency, _
        ByVal pnIntCompPag As Currency, ByVal pnIntMoratPag As Currency, _
        ByVal pnGastoPag As Currency, ByVal pmGastos As Variant, ByVal pnNroCalend As Integer, ByVal pnEstadoNew As ColocEstado, _
        Optional pbEjecBatch As Boolean = False, Optional psPersLavDinero As String = "", Optional ByVal pnITF As Double = 0#, Optional pnDocTpo As Integer = 0, Optional psNroDoc As String = "", _
        Optional ByVal pOpeITFCheqEfec As String = "", Optional ByVal psPersCmac As String, _
        Optional ByVal pnNewSaldoIntMorGener As Currency = 0, _
        Optional psTitPersLavDinero As String = "", Optional psOrdPersLavDinero As String = "", _
        Optional psReaPersLavDinero As String = "", Optional psBenPersLavDinero As String = "", _
        Optional psVisPersLavDinero As String = "", _
        Optional pnMovNro As Long = 0, Optional ByVal psPersCod As String = "", Optional ByVal psIFTpo As String = "01", Optional ByVal psIFCta As String = "", _
        Optional nMovNroRVD As Long = 0, Optional nMovNroRVDPen As Long = 0)
        'By Capi 18022008 se agrego los 5 ultimos parametros
        'RIRO20140530 ERS017, Se agrego nMovNroRVD
        
'** Update Producto
'** Update Colocaciones
'** Update ColocRecup
'** Insert ColocRecupGasto (Comision)
'** Insert ColocRecupGasto (Gasto Admin)
'** Update ColocRecGastos
'** dInsertMov
'** dInsertMovCol
'** dInsertMovColDet
'************************************

Dim lsSQL As String
Dim loReg As COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim lnNroGastoSgte As Integer, I As Integer
Dim lnNroGastoComis As Integer, lnNroGastoAdmin As Integer

Dim oConecta As COMConecta.DCOMConecta
Dim Rs As ADODB.Recordset
Dim nEstadoActual As Integer
'JUEZ 20130429 *************
Dim nMontoCondCap As Double
Dim nMontoCondInt As Double
Dim nMontoCondGas As Double 'JIPR20201128 GastosCond

'END JUEZ ******************

'WIOR 20150421 ***
Dim oCred As COMDColocRec.DCOMColRecCredito
Dim nPrestamo As Currency
Dim nIntDiferido As Currency
Dim nRevIntDif As Currency
Dim sOpeRevIntDif As String
Dim nIntDiferidoPend As Currency

nRevIntDif = 0
sOpeRevIntDif = ""
nIntDiferidoPend = 0
'WIOR FIN ********

lsOpeCod = psOpeCod

lnNroGastoSgte = pnNroGastoCta

On Error GoTo ControlError

'WIOR 20150421 ***
Set oCred = New COMDColocRec.DCOMColRecCredito
Set Rs = oCred.RecuperaColocaciones(psCtaCod)
nPrestamo = Rs!nMontoCol
Set Rs = Nothing

Set Rs = oCred.ObtenerInteresDiferidoCredito(psCtaCod)
If Rs.RecordCount > 0 Then
    nIntDiferido = IIf(IsNull(Rs!nInteresDif), 0, Rs!nInteresDif)
    sOpeRevIntDif = gCredRevIntDifJud
    If nIntDiferido > 0 Then
        Set Rs = oCred.ObtenerInteresDiferidoPendCredito(psCtaCod)
        If Rs.RecordCount > 0 Then
            nIntDiferidoPend = IIf(IsNull(Rs!nIntDifPend), 0, Rs!nIntDifPend)
        End If
    End If
End If
Set oCred = Nothing
Set Rs = Nothing
'WIOR FIN ********

Set loReg = New COMDColocPig.DCOMColPActualizaBD
    loReg.dBeginTrans
    mbTrans = True
    
    '19-05-2006
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    lsSQL = "SELECT nPrdEstado FROM Producto WHERE cCtaCod='" & psCtaCod & "'"
    Set Rs = oConecta.CargaRecordSet(lsSQL)
    nEstadoActual = Rs!nPrdEstado
    oConecta.CierraConexion
    Set oConecta = Nothing
    '*****************
    '** Actualiza Producto
    Call loReg.dUpdateProducto(psCtaCod, , pnNewSaldoCap, pnEstadoNew, psFechaHora, -2, False)          ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loReg.dUpdateColocaciones(psCtaCod, , , , , psMovNro, , False)
    
    '** Actualiza ColocRecup
    'Modificado por DAOR 20070809
    'Call loReg.dUpdateColocRecup(psCtaCod, , , pnNewSaldoIntCom, pnNewSaldoIntMor, pnNewSaldoGasto, pnNewSaldoIntComGener, , , , , False)
    Call loReg.dUpdateColocRecup(psCtaCod, , , pnNewSaldoIntCom, pnNewSaldoIntMor, pnNewSaldoGasto, pnNewSaldoIntComGener, , , , , False, , , , , pnNewSaldoIntMorGener)

    'Si tiene Comision
    If pnComisionAbog > 0 Then
        '** Insert ColocRecupGastos
        lnNroGastoSgte = lnNroGastoSgte + 1
        lnNroGastoComis = lnNroGastoSgte
        Call loReg.dInsertColocRecupGastos(psCtaCod, lnNroGastoComis, gColRecConceptoCodGastoComisionAbog, psFechaHora, pnComisionAbog, pnComisionAbog, gColRecGastoEstPagado, "Comision Abogado", False)
    End If
    
    'Para cada Gasto que se tenga
    For I = 0 To UBound(pmGastos) - 1
        If pmGastos(I, 5) = "S" Then  ' Si ha sido modificado
            '** Update ColocRecupGastos
            Call loReg.dUpdateColocRecupGastos(psCtaCod, CInt(pmGastos(I, 1)), , , Format(CDbl(pmGastos(I, 6)), "#0.00"), CInt(pmGastos(I, 4)), , False)
        End If
    Next I

    'Si tiene Gasto Administrativo
    If pnGastoAdmin > 0 Then
        '** Insert ColocRecupGastos
        lnNroGastoSgte = lnNroGastoSgte + 1
        lnNroGastoAdmin = lnNroGastoSgte
        Call loReg.dInsertColocRecupGastos(psCtaCod, lnNroGastoAdmin, gColRecConceptoCodGastoAdmin, psFechaHora, pnComisionAbog, pnGastoAdmin, gColRecGastoEstPagado, "Gasto Administrativ", False)
    End If

    '** Inserta Mov
    Call loReg.dInsertMov(psMovNro, lsOpeCod, "Pago Credito Recup", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loReg.dGetnMovNro(psMovNro)
    'ALPA 20081010
    pnMovNro = lnMovNro
    
    If psNroDoc <> "" Then
        loReg.dInsertMovDoc lnMovNro, pnDocTpo, psNroDoc, psFechaHora
        'EJVG20140303 ***
        loReg.dInsertaMovDocRec lnMovNro, pnDocTpo, psNroDoc, psPersCod, psIFTpo, psIFCta
        loReg.dInsertaCuentaDocumento lnMovNro, pnDocTpo, psNroDoc, psPersCod, psIFTpo, psIFCta, psCtaCod, pnMontoPagar + pnITF
        'END EJVG *******
    End If
    
    'RIRO 20140530 ERS017 ***
    If nMovNroRVD > 0 Then
        'CTI2 20181207 COMENTADO ...
        'ActualizarRelacionVoucherCaptacion lnMovNro, nMovNroRVD
        'ActualizaMovPendientesRend nMovNroRVDPen, pnMontoPagar + pnITF ' monto a pagar mas el itf
        'END CTI2
        
        'CTI2 20181207 ********
        loReg.ActualizarRelacionVoucherCaptacion lnMovNro, nMovNroRVD
        loReg.ActualizaMovPendientesRend nMovNroRVDPen, pnMontoPagar + pnITF ' monto a pagar mas el itf
        'END ******************
    End If
    'END RIRO ***************
    
    '*********************************************
    '***** 19-05-2006 (Agregamos el nCredEstado)
    'ITF
    If pnITF > 0 Then
        If gColRecOpePagJudSDEnOtCjEfe = lsOpeCod Then
           Call loReg.dInsertMovCol(lnMovNro, gITFCobroCMACCred, psCtaCod, pnNroCalend, pnITF, 0, psMetliquid, 0, nEstadoActual, pnNewSaldoCap, False)
           Call loReg.dInsertMovColDet(lnMovNro, gITFCobroCMACCred, psCtaCod, pnNroCalend, gConcITFCliente, 0, pnITF, False)
        ElseIf gColRecOpePagJudCDEnOtCjEfe = lsOpeCod Then
           Call loReg.dInsertMovCol(lnMovNro, gITFCobroCMACCred, psCtaCod, pnNroCalend, pnITF, 0, psMetliquid, 0, nEstadoActual, pnNewSaldoCap, False)
           Call loReg.dInsertMovColDet(lnMovNro, gITFCobroCMACCred, psCtaCod, pnNroCalend, gConcITFCliente, 0, pnITF, False)
        ElseIf gColRecOpePagJudCastEnOtCjEfe = lsOpeCod Then
           Call loReg.dInsertMovCol(lnMovNro, gITFCobroCMACCred, psCtaCod, pnNroCalend, pnITF, 0, psMetliquid, 0, nEstadoActual, pnNewSaldoCap, False)
           Call loReg.dInsertMovColDet(lnMovNro, gITFCobroCMACCred, psCtaCod, pnNroCalend, gConcITFCliente, 0, pnITF, False)
        Else
           Call loReg.dInsertMovCol(lnMovNro, pOpeITFCheqEfec, psCtaCod, pnNroCalend, pnITF, 0, psMetliquid, 0, nEstadoActual, pnNewSaldoCap, False)
           Call loReg.dInsertMovColDet(lnMovNro, pOpeITFCheqEfec, psCtaCod, pnNroCalend, gConcITFCliente, 0, pnITF, False)
        End If
    End If
    
    '** Inserta MovCol
    Call loReg.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, pnMontoPagar, 0, psMetliquid, 0, nEstadoActual, pnNewSaldoCap, False)
    
    '** Inserta MovColDet - Capital
    If pnCapPag > 0 Then
        Call loReg.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gColRecConceptoCodCapital, 0, pnCapPag, False)
    End If
    
    'WIOR 20150421 ***
    If pnCapPag > 0 Then
        If nIntDiferido > 0 And Trim(sOpeRevIntDif) <> "" And nIntDiferidoPend > 0 Then
            nRevIntDif = CCur(Round((pnCapPag / nPrestamo) * nIntDiferido, 2))
            nIntDiferidoPend = nIntDiferidoPend - nRevIntDif
            
            If nIntDiferidoPend < 0 Then
                nRevIntDif = nRevIntDif + nIntDiferidoPend
            End If
                
            Call loReg.dInsertMovCol(lnMovNro, sOpeRevIntDif, psCtaCod, pnNroCalend, nRevIntDif, 0, psMetliquid, 0, nEstadoActual, pnNewSaldoCap, False)
            Call loReg.dInsertMovColDet(lnMovNro, sOpeRevIntDif, psCtaCod, pnNroCalend, gColocConceptoCodInteresDiferidoAmp, 0, nRevIntDif, False)
        End If
    End If
    'WIOR FIN ********
    
    '** Inserta MovColDet - Interes Comp
    If pnIntCompPag > 0 Then
        Call loReg.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gColRecConceptoCodInteresCompensatorio, 0, pnIntCompPag, False)
    End If
    
    '** Inserta MovColDet - Interes Morat
    If pnIntMoratPag > 0 Then
        Call loReg.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gColRecConceptoCodInteresMoratorio, 0, pnIntMoratPag, False)
    End If
    
    '** Inserta MovColDet - Gastos (Sin Comision, ni Gasto Administrativo )
    'If pnGastoPag > 0 Then
    '    Call loReg.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gColRecConceptoCodGastoCobranza, 0, pnGastoPag, False)
    'End If
    
    'Para cada Gasto que se tenga (Sin Comision, ni Gasto Administrativo )
    For I = 0 To UBound(pmGastos) - 1
        If pmGastos(I, 5) = "S" Then  ' Si ha sido modificado
            '** Inserta MovColDet - Gastos (Sin Comision, ni Gasto Administrativo )
            Call loReg.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, pmGastos(I, 7), pmGastos(I, 1), Format(CDbl(pmGastos(I, 6)), "#0.00"), False)
        'ElseIf pmGastos(I, 5) = "N" Then
            'Call loReg.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, pmGastos(I, 7), pmGastos(I, 1), Format(CDbl(pmGastos(I, 6)), "#0.00"), False)
        End If
    Next I
    
    '** Inserta MovColDet - Comision Abogado
    If pnComisionAbog > 0 Then
        Call loReg.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gColRecConceptoCodGastoComisionAbog, lnNroGastoComis, pnComisionAbog, False)
    End If
    
    '** Inserta MovColDet - Gasto Administ
    If pnGastoAdmin > 0 Then
        Call loReg.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gColRecConceptoCodGastoAdmin, lnNroGastoAdmin, pnGastoAdmin, False)
    End If
'ALPA 20081010************************************************************
    ''''''''''''''''''''
    ' Lavado de Dinero
''    If psPersLavDinero <> "" Then
''        'By Capi 18022008
''        'Call loReg.dInsertaMovLavDinero(lnMovNro, psPersLavDinero)
''        'ALPA 20081009*******************************************************************
''        'Call loReg.dInsertaMovLavDinero(lnMovNro, psTitPersLavDinero, psOrdPersLavDinero, psReaPersLavDinero, psBenPersLavDinero, psVisPersLavDinero)
''        Call loReg.dInsertaMovLavDinero(lnMovNro, psTitPersLavDinero, psOrdPersLavDinero, psReaPersLavDinero, psBenPersLavDinero, psVisPersLavDinero, nTipoREU, nMontoAcu, sOrigen)
''        '********************************************************************************
''    End If
    
    ''''''''''''''''''''
    
    If gColRecOpePagJudSDEnOtCjEfe = lsOpeCod Then
        Call loReg.dInsertMovCMAC(lnMovNro, psPersCmac, Format$(gTpoIFCmac, "00"), CInt(Mid(psCtaCod, 9, 1)), "", "", lsOpeCod, pnMontoPagar, False)
    ElseIf gColRecOpePagJudCDEnOtCjEfe = lsOpeCod Then
        Call loReg.dInsertMovCMAC(lnMovNro, psPersCmac, Format$(gTpoIFCmac, "00"), CInt(Mid(psCtaCod, 9, 1)), "", "", lsOpeCod, pnMontoPagar, False)
    ElseIf gColRecOpePagJudCastEnOtCjEfe = lsOpeCod Then
        Call loReg.dInsertMovCMAC(lnMovNro, psPersCmac, Format$(gTpoIFCmac, "00"), CInt(Mid(psCtaCod, 9, 1)), "", "", lsOpeCod, pnMontoPagar, False)
    End If
    
    'JUEZ 20130429 ****************************************************************
    If pnEstadoNew = gColocEstRecCanJud Or pnEstadoNew = gColocEstRecCanCast Then
        nMontoCondCap = pnNewSaldoCap
        nMontoCondInt = pnNewSaldoIntCom + pnNewSaldoIntMor
        nMontoCondGas = pnNewSaldoGasto 'JIPR20201128 GastosCond
        
        If nMontoCondCap > 0 Or nMontoCondInt > 0 Or nMontoCondGas > 0 Then 'JIPR20201128 GastosCond
        'If nMontoCondCap > 0 Or nMontoCondInt > 0 Then
        
            Call loReg.dInsertMovCol(lnMovNro, gAsientoCondonacionRecup, psCtaCod, 0, nMontoCondCap + nMontoCondInt + nMontoCondGas, 0, "", 0, 0, 0, False) 'JIPR20201128 GastosCond
            'Call loReg.dInsertMovCol(lnMovNro, gAsientoCondonacionRecup, psCtaCod, 0, nMontoCondCap + nMontoCondInt, 0, "", 0, 0, 0, False)
            If nMontoCondCap > 0 Then
                Call loReg.dInsertMovColDet(lnMovNro, gAsientoCondonacionRecup, psCtaCod, 0, gColRecConceptoCodCapital, 0, nMontoCondCap, False)
            End If
            If nMontoCondInt > 0 Then
                Call loReg.dInsertMovColDet(lnMovNro, gAsientoCondonacionRecup, psCtaCod, 0, gColRecConceptoCodInteresCompensatorio, 0, nMontoCondInt, False)
            End If
            'JIPR20201128 GastosCond Inicio
            If nMontoCondGas > 0 Then
                Call loReg.dInsertMovColDet(lnMovNro, gAsientoCondonacionRecup, psCtaCod, 0, gColRecConceptoCodGastoCobranza, 0, nMontoCondGas, False)
            End If
            'JIPR20201128 GastosCond Fin
            
        End If
    End If
    'END JUEZ *********************************************************************
    'EJVG20150822 ***
    If pnCapPag > 0 Then
        Call loReg.dLiberarGarantia(lnMovNro, lsOpeCod, psCtaCod, pnCapPag, pnEstadoNew, False)
    End If
    'END EJVG *******
    loReg.dCommitTrans
    mbTrans = False
Set loReg = Nothing

Exit Sub
ControlError:
    If mbTrans Then
        loReg.dRollbackTrans
        pnMovNro = 0
        mbTrans = False
        Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "
    End If
    

End Sub

Public Sub bVistoRecuperaciones(ByVal psCtaCod As String, Optional ByVal pnMetodo As Integer = -1, Optional ByVal pnNegociacion As Integer = -1, _
                                Optional ByVal pnCancelacion As Integer = -1)

'** Update ColocRecup

Dim lsSQL As String
Dim loReg As COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

'On Error GoTo ControlError
Set loReg = New COMDColocPig.DCOMColPActualizaBD
    loReg.dBeginTrans
    mbTrans = True
    
    '** Actualiza ColocRecup

    Call loReg.dUpdateColocRecup(psCtaCod, , , , , , , , , , , , , pnMetodo, pnNegociacion, pnCancelacion)
    
    'mnEjecutaBatch = loReg.dEjecutaBatch
    loReg.dCommitTrans
    mbTrans = False
Set loReg = Nothing

Exit Sub
ControlError:
    If mbTrans Then
        loReg.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nCancelacionCreditoRecup", "Error en Funcion Cancelacion Credito Recup "

End Sub

'**DAOR 20070414, Metodo que registra las Distribución de los montos CIMG a cobrar
'**de una determinada Cuenta, para una determinada fecha
Public Sub nRegistraDistribucionCIMGCobranza(ByVal psCtaCod As String, pdFecha As Date, _
        ByVal pnCapital As Currency, ByVal pnIntComp As Currency, ByVal pnMora As Currency, _
        ByVal pnGasto As Currency, ByVal pnComisionAbog As Currency, ByVal psUltimaActualizacion As String, _
        Optional ByVal pnIntCompGen As Currency = 0, Optional ByVal pnIntMoraGen As Currency = 0) 'DAOR 20070809

Dim lsSQL As String
Dim loRegRec As COMDColocPig.DCOMColPActualizaBD

'On Error GoTo ErrornRegistraDistribucionCIMGCobranza
Set loRegRec = New COMDColocPig.DCOMColPActualizaBD
    loRegRec.dBeginTrans
    mbTrans = True
    
    Call loRegRec.dDeleteColocRecupDistribucionCobr(psCtaCod, pdFecha)
    Call loRegRec.dInsertColocRecupDistribucionCobr(psCtaCod, pdFecha, pnCapital, pnIntComp, pnMora, pnGasto, pnComisionAbog, psUltimaActualizacion)
    
    '**DAOR 20070809 ******************************
    Call loRegRec.dUpdateColocRecup(psCtaCod, , , , , , pnIntCompGen, , , , , , , , , , pnIntMoraGen)
    '**********************************************

    loRegRec.dCommitTrans
    mbTrans = False
Set loRegRec = Nothing

Exit Sub
ErrornRegistraDistribucionCIMGCobranza:
    If mbTrans Then
        loRegRec.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraDistribucionCIMGCobranza", "Error en Funcion de Distribución de Montos CIMG de Cobranza "
End Sub

'******************************************************
'MADM 20110510 - END MADM
Public Function GetBuscaGastoCreditoFinanciero(ByVal psCtaCod As String, ByVal pnNroGastoCta As Integer) As Boolean
Dim ssql As String
Dim Rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta

GetBuscaGastoCreditoFinanciero = False

ssql = " EXEC stp_val_BuscaGastoCreditoFinanciero '" & psCtaCod & "', " & pnNroGastoCta

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set Rs = oCon.CargaRecordSet(ssql)

If Not Rs.EOF Then
    If Rs!nMovNro > 0 Then
        GetBuscaGastoCreditoFinanciero = True
    End If
    Rs.Close
    Set Rs = Nothing
End If

oCon.CierraConexion
Set oCon = Nothing
End Function
'******************************************************

'JACA 20111907************************************************************
Public Function ObtenerCuentaEmbargo(ByVal psCtaCod As String) As Recordset
    Dim oColRec As COMDColocRec.DCOMColRecCredito
    Set oColRec = New COMDColocRec.DCOMColRecCredito
    Set ObtenerCuentaEmbargo = oColRec.getCuentaEmbargo(psCtaCod)
End Function
Public Function ObtenerConsValorEmbargo(ByVal nConsCod As Integer, ByVal nConsValorLeft As String, ByVal nConsValorRight As String) As Recordset
    Dim oColRec As COMDColocRec.DCOMColRecCredito
    Set oColRec = New COMDColocRec.DCOMColRecCredito
    Set ObtenerConsValorEmbargo = oColRec.getConsValorEmbargo(nConsCod, nConsValorLeft, nConsValorRight)

End Function
Public Function ObtenerAlmacenEmbargo(ByVal psCodAge As String, ByVal psCodAlmacen As String) As Recordset
    Dim oColRec As COMDColocRec.DCOMColRecCredito
    Set oColRec = New COMDColocRec.DCOMColRecCredito
    Set ObtenerAlmacenEmbargo = oColRec.getAlmacenEmbargo(psCodAge, psCodAlmacen)

End Function
Public Function ObtenerAlmacenCorrelativo(ByVal psCodAge As String) As String
    Dim oColRec As COMDColocRec.DCOMColRecCredito
    Set oColRec = New COMDColocRec.DCOMColRecCredito
    ObtenerAlmacenCorrelativo = oColRec.getAlmacenCorrelativo(psCodAge)

End Function
Public Sub guardarAlmacenEmbargo(ByVal psCodAlmacen As String, ByVal psNomAlmacen As String, ByVal psDescripcion As String, ByVal psNroMov As String)
    Dim oColRec As COMDColocRec.DCOMColRecCredito
    Set oColRec = New COMDColocRec.DCOMColRecCredito
    oColRec.insertarAlmacenEmbargo psCodAlmacen, psNomAlmacen, psDescripcion, psNroMov

End Sub
Public Function ObtenerBienesEmbargo(ByVal psSubTpoBien As String) As Recordset
    Dim oColRec As COMDColocRec.DCOMColRecCredito
    Set oColRec = New COMDColocRec.DCOMColRecCredito
    Set ObtenerBienesEmbargo = oColRec.getBienesEmbargo(psSubTpoBien)

End Function
Public Sub guardarBienEmbargo(ByVal pnConsValor As Integer, ByVal pnCodBien As Integer, ByVal psNomBien As String, ByVal psNroMov As String)
    Dim oColRec As COMDColocRec.DCOMColRecCredito
    Set oColRec = New COMDColocRec.DCOMColRecCredito
    oColRec.insertarBienEmbargo pnConsValor, pnCodBien, psNomBien, psNroMov

End Sub
Public Function ObtenerBienCorrelativo(ByVal pnConsValor As Integer) As Integer
    Dim oColRec As COMDColocRec.DCOMColRecCredito
    Set oColRec = New COMDColocRec.DCOMColRecCredito
    ObtenerBienCorrelativo = oColRec.getBienCorrelativo(pnConsValor)

End Function
Public Sub modificarBienEmbargo(ByVal pnConsValor As Integer, ByVal pnCodBien As Integer)
    Dim oColRec As COMDColocRec.DCOMColRecCredito
    Set oColRec = New COMDColocRec.DCOMColRecCredito
    oColRec.actualizarBienEmbargo pnConsValor, pnCodBien

End Sub
Public Sub modificarAlmacenEmbargo(ByVal psCodAlmacen As String)
    Dim oColRec As COMDColocRec.DCOMColRecCredito
    Set oColRec = New COMDColocRec.DCOMColRecCredito
    oColRec.actualizarAlmacenEmbargo psCodAlmacen

End Sub
Public Sub guardarEmbargo(ByVal pnMovNro As Long, ByVal psCtaCod As String, ByVal pnCapital As Currency, ByVal pnInteres As Currency, ByVal pnGastos As Currency, _
                          ByVal psNroExpediente As String, ByVal psNroResolucion As String, ByVal pdFecEmbargo As Date, ByVal psUltimaActualizacion As String)
    Dim oColRec As COMDColocRec.DCOMColRecCredito
    Set oColRec = New COMDColocRec.DCOMColRecCredito
    oColRec.insertarEmbargo pnMovNro, psCtaCod, pnCapital, pnInteres, pnGastos, psNroExpediente, psNroResolucion, pdFecEmbargo, psUltimaActualizacion

End Sub
Public Sub guardarEmbargoDetalle(ByVal psCtaCod As String, ByVal pnNroBienEmbargo As Integer, ByVal psDocSalida As String, ByVal pdFecEmbargo As Date, _
                                ByVal pnCodSubTpoBien As Integer, ByVal pnCodBien As Integer, ByVal pnEstadoBien As Integer, ByVal pnEstadoEmbargo As Integer, ByVal psCodAlmacen As String, _
                                ByVal psBienDescripcion As String, ByVal pnCantidad As String, ByVal psColor As String, ByVal psMarca As String, ByVal psModelo As String, _
                                ByVal psSerie As String, ByVal psMotor As String, ByVal pnTasacion As Currency, ByVal psPlaca As String, ByVal psPartida As String, ByVal psUltimaActualizacion As String, _
                                Optional pdFecSalida As String = "", Optional pdFecTasacion As String = "", Optional psMonTasacion As String = "", _
                                Optional psDepositario As String = "", Optional psTitularBien As String = "")

        Dim oColRec As COMDColocRec.DCOMColRecCredito
        Set oColRec = New COMDColocRec.DCOMColRecCredito
        oColRec.insertarEmbargoDetalle psCtaCod, pnNroBienEmbargo, psDocSalida, pdFecEmbargo, _
                                        pnCodSubTpoBien, pnCodBien, pnEstadoBien, pnEstadoEmbargo, psCodAlmacen, _
                                        psBienDescripcion, pnCantidad, psColor, psMarca, psModelo, _
                                        psSerie, psMotor, pnTasacion, psPlaca, psPartida, psUltimaActualizacion, _
                                        pdFecSalida, pdFecTasacion, psMonTasacion, psDepositario, psTitularBien
        
        
    
End Sub
'*** PEAC 20130618
'Public Function obtenerEmbargosListar(ByVal pdFecha As Date) As Recordset
Public Function obtenerEmbargosListar(ByVal psFecDel As String, ByVal psFecAl As String, ByVal psNumCta As String, ByVal psCodCli As String, ByVal psNroExp As String, ByVal psNroRes As String) As Recordset

    Dim oColRec As COMDColocRec.DCOMColRecCredito
    Set oColRec = New COMDColocRec.DCOMColRecCredito
    'Set obtenerEmbargosListar = oColRec.getEmbargosListar(pdFecha)
    Set obtenerEmbargosListar = oColRec.getEmbargosListar(psFecDel, psFecAl, psNumCta, psCodCli, psNroExp, psNroRes)
    
End Function
Public Function obtenerEmbargo(ByVal psCtaCod As String) As Recordset
    Dim oColRec As COMDColocRec.DCOMColRecCredito
    Set oColRec = New COMDColocRec.DCOMColRecCredito
    Set obtenerEmbargo = oColRec.getEmbargo(psCtaCod)
End Function
Public Function obtenerEmbargoDetalle(ByVal psCtaCod As String) As Recordset
    Dim oColRec As COMDColocRec.DCOMColRecCredito
    Set oColRec = New COMDColocRec.DCOMColRecCredito
    Set obtenerEmbargoDetalle = oColRec.getEmbargoDetalle(psCtaCod)
End Function
Public Sub modificarEmbargo(ByVal psCtaCod As String, Optional pnNroBienEmbargo As Integer = 0)
    Dim oColRec As COMDColocRec.DCOMColRecCredito
    Set oColRec = New COMDColocRec.DCOMColRecCredito
    oColRec.actualizarEmbargo psCtaCod, pnNroBienEmbargo
End Sub
Public Function verificarEmbargo(ByVal psCtaCod As String) As Recordset
    Dim oColRec As COMDColocRec.DCOMColRecCredito
    Set oColRec = New COMDColocRec.DCOMColRecCredito
    Set verificarEmbargo = oColRec.getVerificarEmbargo(psCtaCod)
End Function
Public Function obtenerDistribucionSaldos(ByVal psCtaCod As String) As Recordset
    Dim oColRec As COMDColocRec.DCOMColRecCredito
    Set oColRec = New COMDColocRec.DCOMColRecCredito
    Set obtenerDistribucionSaldos = oColRec.getDistribucionSaldos(psCtaCod)
End Function

Public Sub guardarPagoCredAdjudicacion(ByVal psOpeCod As String, ByVal psCtaCod As String, ByVal pnMonto As Currency, ByVal pdFecha As Date, ByVal pnMoneda As Integer, ByVal psMovNro As String)

        Dim oColRec As COMDColocRec.DCOMColRecCredito
        Set oColRec = New COMDColocRec.DCOMColRecCredito
        oColRec.insertarPagoCredAdjudicacion psOpeCod, psCtaCod, pnMonto, pdFecha, pnMoneda, psMovNro
        End Sub
Public Function ObtenerPagoCredAdjudicacion(ByVal psCtaCod As String, ByVal psOpeCod As String) As Recordset
    Dim oColRec As COMDColocRec.DCOMColRecCredito
    Set oColRec = New COMDColocRec.DCOMColRecCredito
    Set ObtenerPagoCredAdjudicacion = oColRec.getPagoCredAdjudicacion(psCtaCod, psOpeCod)
End Function
Public Function ObtenerNroBienesEmbargados(ByVal psCtaCod As String) As Integer
    Dim oColRec As COMDColocRec.DCOMColRecCredito
    Set oColRec = New COMDColocRec.DCOMColRecCredito
     ObtenerNroBienesEmbargados = oColRec.getNroBienesEmbargados(psCtaCod)
End Function
Public Function ObtenerAgenciaCiudad(ByVal psAgeCod As String) As String
     Dim oColRec As COMDColocRec.DCOMColRecCredito
     Set oColRec = New COMDColocRec.DCOMColRecCredito
     ObtenerAgenciaCiudad = oColRec.getAgenciaCiudad(psAgeCod)
End Function
Public Function guardarMovPendientesRend(ByVal pnMovNro As Long, ByVal psCtaContCod As String, ByVal pnSaldo As Currency) As Integer
    Dim oColRec As COMDColocRec.DCOMColRecCredito
    Set oColRec = New COMDColocRec.DCOMColRecCredito
    guardarMovPendientesRend = oColRec.InsertaMovPendientesRend(pnMovNro, psCtaContCod, pnSaldo)
End Function
'JACA END*****************************************************************

'**Juez 20120421 ****************************************************************
Public Sub nRegistraPagoCancelacionCredJudicial(ByVal psCtaCod As String, pdFecha As Date, _
        ByVal pnCapital As Currency, ByVal pnIntComp As Currency, ByVal pnMora As Currency, _
        ByVal pnGasto As Currency, ByVal pnComisionAbog As Currency, psMetliquid, _
        ByVal psGlosa As String, ByVal psNroCalen As Integer, ByVal psMovNro As String, ByVal psOpeCod As String, _
        Optional ByVal pnIntCompGen As Currency = 0, Optional ByVal pnIntMoraGen As Currency = 0) 'ByRef lsmensaje As String)

Dim lsSQL As String
Dim loRegRec As COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Set loRegRec = New COMDColocPig.DCOMColPActualizaBD
    loRegRec.dBeginTrans
    mbTrans = True
    
    Call loRegRec.dDeleteColocRecupDistribucionCobr(psCtaCod, pdFecha)
    Call loRegRec.dInsertColocRecupDistribucionCobr(psCtaCod, pdFecha, pnCapital, pnIntComp, pnMora, pnGasto, pnComisionAbog, psMovNro, 1)
        
    '** Actualiza Colocaciones
    'Call loRegRec.dUpdateColocaciones(psCtaCod, , , , , psMovNro, , False)
    
    '** Actualiza ColocRecup
    Call loRegRec.dUpdateColocRecup(psCtaCod, , , , , , pnIntCompGen, psMetliquid, , , , False, psGlosa, , , , pnIntMoraGen)

    '** Inserta Mov
    Call loRegRec.dInsertMov(psMovNro, psOpeCod, "Grabar Cancelación Credito con Pago Judicial", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loRegRec.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loRegRec.dInsertMovCol(lnMovNro, psOpeCod, psCtaCod, psNroCalen, pnCapital, 0, "", 0, 0, pnCapital, False)
    
    loRegRec.dCommitTrans
    
    mbTrans = False
Set loRegRec = Nothing

Exit Sub
ErrornRegistraPagCancCredJudicial:
    If mbTrans Then
        loRegRec.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error ErrornRegistraPagCancCredJudicial", "Error en Funcion de Registro de Pago de Cancelación para Judicial "
End Sub
'**End Juez ****************************************************************

'*** PEAC 20120709
Public Sub RegistraVisitaGestores(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psCtaCod As String, ByVal psFechaVisita As String, ByVal pnPersContac As Integer, _
        ByVal psFecCompromiso As String, ByVal pnResulGestion As Integer, ByVal psZona As String, ByVal pnLugarContac As Integer, _
        ByVal pnMontoCompromiso As Double, ByVal pnCondicion As Integer, ByVal psComentario As String)

'** Inserta Mov
'** Inserta ColocRecupVisGest
'************************************

Dim lsSQL As String
Dim loRegRec As COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long

Set loRegRec = New COMDColocPig.DCOMColPActualizaBD
    loRegRec.dBeginTrans
    mbTrans = True
    
    '** Inserta Mov
    Call loRegRec.dInsertMov(psMovNro, psOpeCod, "Registro de visita de gestores", gMovEstContabNoContable, gMovFlagVigente, False)

    ' Obtiene nMovNro
    lnMovNro = loRegRec.dGetnMovNro(psMovNro)

    'Call loRegRec.dInsertColocRecupActProcesales(psCtaCod, psMovNro, psComenta, False, pdFechaA, pdFechaV, pbAlerta, gsProyectoActual)
    Call loRegRec.InsertColocRecupVisitaGestores(lnMovNro, psCtaCod, psFechaVisita, pnPersContac, _
        psFecCompromiso, pnResulGestion, psZona, pnLugarContac, _
        pnMontoCompromiso, pnCondicion, psComentario)
    
    loRegRec.dCommitTrans
    mbTrans = False
Set loRegRec = Nothing

Exit Sub
ErrorModRec:
    If mbTrans Then
        loRegRec.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraActProcesales", "Error en Funcion COMNColocRec/RegistraVisitaGestores"

End Sub

'*******RECO 2013-07-20***********
Public Sub nRegistraAutorizacionPagoJud(ByVal cCtaCod As String, ByVal cMovNro As String, ByVal dFecRegistro As String, _
    ByVal cMetLiquid As String, ByVal nEstado As Integer, ByVal cMetLiquidNuevo As String)
Dim lsSQL As String
Dim loRegRec As COMDColocPig.DCOMColPActualizaBD

On Error GoTo ErrornRegistraAutorizacionPagoJud
Set loRegRec = New COMDColocPig.DCOMColPActualizaBD
    loRegRec.dBeginTrans
    mbTrans = True
    
    Call loRegRec.dInsertColocRecupJudAutorizado(cCtaCod, cMovNro, Format(dFecRegistro, "yyyy/MM/dd"), cMetLiquid, nEstado, cMetLiquidNuevo)
    'Call loRegRec.dActualizaMetLiquidPagoJud(cCtaCod, cMetLiquidNuevo)
    
    loRegRec.dCommitTrans
    mbTrans = False
Set loRegRec = Nothing

Exit Sub
ErrornRegistraAutorizacionPagoJud:
    If mbTrans Then
        loRegRec.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraAutorizacionPagoJud", "Error en Funcion de Distribución de Montos CIMG de Cobranza "
End Sub


Public Function nValidaDiaAutorizacionPagoJud(ByVal cCtaCod As String, ByVal cFecSist As String) As Boolean
Dim loRegRec As COMDColocPig.DCOMColPActualizaBD
Dim dr As ADODB.Recordset
On Error GoTo ErrornValidaDiaAutorizacionPagoJud
Set loRegRec = New COMDColocPig.DCOMColPActualizaBD
    loRegRec.dBeginTrans
    mbTrans = True
    
    Set dr = loRegRec.dValidaDiaAutorizacionPagoJud(cCtaCod, cFecSist)
        
    If dr.RecordCount <> 0 Then
        If dr!dFecRegistro = Format(cFecSist, "yyyy-MM-dd") Then
            nValidaDiaAutorizacionPagoJud = True
        Else
            nValidaDiaAutorizacionPagoJud = False
        End If
    End If
    loRegRec.dCommitTrans
    mbTrans = False
Set loRegRec = Nothing

Exit Function
ErrornValidaDiaAutorizacionPagoJud:
    If mbTrans Then
        loRegRec.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nValidaDiaAutorizacionPagoJud", "Error en Funcion de Distribución de Montos CIMG de Cobranza "
End Function
Public Sub nActualizarMetLiquidPagoJud(ByVal cCtaCod As String)
    Dim loRegRec As COMDColocPig.DCOMColPActualizaBD
    Dim dr As ADODB.Recordset
    Set loRegRec = New COMDColocPig.DCOMColPActualizaBD
    
    On Error GoTo ErrornActualizarMetLiquidPagoJud
    loRegRec.dBeginTrans
    mbTrans = True
    
    Call loRegRec.dActualizarMetLiquidPagoJud(cCtaCod)
        
    loRegRec.dCommitTrans
    mbTrans = False
    Set loRegRec = Nothing

Exit Sub
ErrornActualizarMetLiquidPagoJud:
    If mbTrans Then
        loRegRec.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nActualizarMetLiquidPagoJud", "Error en Funcion de Distribución de Montos CIMG de Cobranza "
End Sub
'*********END RECO***************
'RIRO20140530 ERS017 ***
Public Sub ActualizarRelacionVoucherCaptacion(ByVal pnMovNroOpe As Long, ByVal pnMovNroRVD As Long)
    Dim oCol As New COMDColocRec.DCOMColRecCredito
    oCol.ActualizarRelacionVoucherCaptacion pnMovNroOpe, pnMovNroRVD
End Sub
Public Sub ActualizaMovPendientesRend(ByVal pnMovNroOpe As Long, ByVal pnMonto As Currency)
    Dim oCol As New COMDColocRec.DCOMColRecCredito
    oCol.ActualizaMovPendientesRend pnMovNroOpe, pnMonto
End Sub
'END RIRO **************
'FRHU 20150415 ERS022-2015
Public Sub nRegistraPagoCancelacionCredTransferidos(ByVal psCtaCod As String, pdFecha As Date, _
        ByVal pnCapital As Currency, ByVal pnIntComp As Currency, ByVal pnMora As Currency, ByVal pnGasto As Currency, _
        ByVal psMetliquid As String, ByVal psGlosa As String, ByVal psNroCalen As Integer, ByVal psMovNro As String, _
        Optional ByVal pbCancelacion As Boolean = False)

    Dim lsSQL As String
    Dim loRegRec As COMDColocPig.DCOMColPActualizaBD
    Dim lnMovNro As Long
    
    Set loRegRec = New COMDColocPig.DCOMColPActualizaBD
        loRegRec.dBeginTrans
        mbTrans = True
        
        Call loRegRec.dDeleteColocTransfDistribucionCobr(psCtaCod, pdFecha)
        Call loRegRec.dInsertColocTransfDistribucionCobr(psCtaCod, pdFecha, pnCapital, pnIntComp, pnMora, pnGasto, psMovNro, IIf(pbCancelacion = True, 1, 0))
        Call loRegRec.dUpdateColocTransferido(psCtaCod, , , , , , psMetliquid, , False)
        
        loRegRec.dCommitTrans
        mbTrans = False
    Set loRegRec = Nothing

Exit Sub
ErrornRegistraPagoCancelacionCredTransferidos:
    If mbTrans Then
        loRegRec.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error ErrornRegistraPagoCancelacionCredTransferidos", "Error en Funcion de Registro de Pago de Cancelación para Creditos Transferidos "
End Sub
Public Sub nRegistraAutorizacionPagoTransferencia(ByVal cCtaCod As String, ByVal cMovNro As String, ByVal dFecRegistro As String, _
                        ByVal cMetLiquid As String, ByVal nEstado As Integer, ByVal cMetLiquidNuevo As String, ByVal pcGlosa As String)
    Dim lsSQL As String
    Dim loRegRec As COMDColocPig.DCOMColPActualizaBD

On Error GoTo ErrornRegistraAutorizacionPagoJud
    Set loRegRec = New COMDColocPig.DCOMColPActualizaBD
        loRegRec.dBeginTrans
        mbTrans = True
        
        Call loRegRec.dInsertColocTransfAutorizado(cCtaCod, cMovNro, Format(dFecRegistro, "yyyy/MM/dd"), cMetLiquid, nEstado, cMetLiquidNuevo, pcGlosa)
        
        loRegRec.dCommitTrans
        mbTrans = False
    Set loRegRec = Nothing

Exit Sub
ErrornRegistraAutorizacionPagoJud:
    If mbTrans Then
        loRegRec.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraAutorizacionPagoJud", "Error en Funcion de Distribución de Montos CIMG de Cobranza "
End Sub
Public Sub nPagoCreditoTransferido(ByVal psCtaCod As String, ByVal psFechaHora As String, ByVal psOpeCod As String, _
        ByVal psMovNro As String, ByVal pnMontoPagar As Currency, ByVal psMetliquid As String, _
        ByVal pnNewSaldoCap As Currency, ByVal pnNewSaldoIntCom As Currency, _
        ByVal pnNewSaldoIntMor As Currency, ByVal pnNewSaldoGasto As Currency, ByVal pnNroGastoCta As Integer, _
        ByVal pnCapPag As Currency, ByVal pnIntCompPag As Currency, ByVal pnIntMoratPag As Currency, ByVal pnGastoPag As Currency, _
        ByVal pmGastos As Variant, ByVal pnNroCalend As Integer, ByVal pnEstadoNew As ColocEstado, _
        Optional pbEjecBatch As Boolean = False, Optional psPersLavDinero As String = "", Optional ByVal pnITF As Double = 0#, Optional pnDocTpo As Integer = 0, Optional psNroDoc As String = "", _
        Optional ByVal pOpeITFCheqEfec As String = "", Optional ByVal psPersCmac As String, _
        Optional ByVal pnNewSaldoIntMorGener As Currency = 0, _
        Optional psTitPersLavDinero As String = "", Optional psOrdPersLavDinero As String = "", _
        Optional psReaPersLavDinero As String = "", Optional psBenPersLavDinero As String = "", _
        Optional psVisPersLavDinero As String = "", _
        Optional pnMovNro As Long = 0, Optional ByVal psPersCod As String = "", Optional ByVal psIFTpo As String = "01", Optional ByVal psIFCta As String = "", _
        Optional nMovNroRVD As Long = 0, Optional nMovNroRVDPen As Long = 0, Optional psCtaAhorro As String = "", Optional pMatDatosAho As Variant = "", _
        Optional ByVal pnEstadoTransf As ColocEstado)
        'By Capi 18022008 se agrego los 5 ultimos parametros
        'RIRO20140530 ERS017, Se agrego nMovNroRVD

Dim lsSQL As String
Dim loReg As COMDColocPig.DCOMColPActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim lnNroGastoSgte As Integer, I As Integer
Dim lnNroGastoComis As Integer, lnNroGastoAdmin As Integer
Dim oConecta As COMConecta.DCOMConecta
Dim Rs As ADODB.Recordset
Dim nEstadoActual As Integer
Dim lnMoneda As Integer
Dim lsOpeCodCaptac As String
Dim lsOpeCodItf As String
'JUEZ 20130429 *************
Dim nMontoCondCap As Double
Dim nMontoCondInt As Double
'END JUEZ ******************

lsOpeCod = psOpeCod
lnMoneda = IIf(Mid(psCtaCod, 9, 1) = "1", 1, 2)
lnNroGastoSgte = pnNroGastoCta

On Error GoTo ControlError
Set loReg = New COMDColocPig.DCOMColPActualizaBD
    loReg.dBeginTrans
    mbTrans = True
    
    '19-05-2006
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    lsSQL = "SELECT nPrdEstado FROM Producto WHERE cCtaCod='" & psCtaCod & "'"
    Set Rs = oConecta.CargaRecordSet(lsSQL)
    nEstadoActual = Rs!nPrdEstado
    oConecta.CierraConexion
    Set oConecta = Nothing
    '*****************
    '** Actualiza Producto
    Call loReg.dUpdateProducto(psCtaCod, , pnNewSaldoCap, pnEstadoNew, psFechaHora, -2, False)          ' (-2) aumenta el ste transac
    '** Actualiza Colocaciones
    Call loReg.dUpdateColocaciones(psCtaCod, , , , , psMovNro, , False)
    
    Call loReg.dUpdateColocTransferido(psCtaCod, , , pnNewSaldoIntCom, pnNewSaldoIntMor, pnNewSaldoGasto, , , False)
    
    'Para cada Gasto que se tenga
    For I = 0 To UBound(pmGastos) - 1
        If pmGastos(I, 5) = "S" Then  ' Si ha sido modificado
            '** Update ColocTransferidoGastos
            Call loReg.dUpdateColocTransferidoGastos(psCtaCod, CInt(pmGastos(I, 1)), , , Format(CDbl(pmGastos(I, 6)), "#0.00"), CInt(pmGastos(I, 4)), , False)
        End If
    Next I
    
    '** Inserta Mov
    Call loReg.dInsertMov(psMovNro, lsOpeCod, "Pago Credito Transferido FOCMACM", gMovEstContabMovContable, gMovFlagVigente, False)
    
    ' Obtiene nMovNro
    lnMovNro = loReg.dGetnMovNro(psMovNro)
    'ALPA 20081010
    pnMovNro = lnMovNro
    
    'Inserta ColocTransferidoDet
    Call loReg.dInsertColocTransferidoDet(lnMovNro, psCtaCod, pnNewSaldoCap + pnCapPag, pnNewSaldoIntCom + pnIntCompPag, pnNewSaldoIntMor + pnIntMoratPag, pnNewSaldoGasto + pnGastoPag, pnNewSaldoCap, pnNewSaldoIntCom, pnNewSaldoIntMor, pnNewSaldoGasto, False)
    
    If psNroDoc <> "" Then
        loReg.dInsertMovDoc lnMovNro, pnDocTpo, psNroDoc, psFechaHora
        'EJVG20140303 ***
        loReg.dInsertaMovDocRec lnMovNro, pnDocTpo, psNroDoc, psPersCod, psIFTpo, psIFCta
        loReg.dInsertaCuentaDocumento lnMovNro, pnDocTpo, psNroDoc, psPersCod, psIFTpo, psIFCta, psCtaCod, pnMontoPagar + pnITF
        'END EJVG *******
    End If
    
    'RIRO 20140530 ERS017 ***
    If nMovNroRVD > 0 Then
        'CTI2 20181207 Comentado *******
        'ActualizarRelacionVoucherCaptacion lnMovNro, nMovNroRVD
        'ActualizaMovPendientesRend nMovNroRVDPen, pnMontoPagar + pnITF ' monto a pagar mas el itf
        'CTI2 END **********************
        
        'ADD CTI2 20181207 - incluir dentro de la transaccion de pago
            loReg.ActualizarRelacionVoucherCaptacion lnMovNro, nMovNroRVD
            loReg.ActualizaMovPendientesRend nMovNroRVDPen, pnMontoPagar + pnITF
        'END ***
    End If
    'END RIRO ***************

    'ITF
    If pnITF > 0 Then
        Call loReg.dInsertMovCol(lnMovNro, "990122", psCtaCod, pnNroCalend, pnITF, 0, psMetliquid, 0, nEstadoActual, pnNewSaldoCap, False)
        Call loReg.dInsertMovColDet(lnMovNro, "990122", psCtaCod, pnNroCalend, gConcITFCliente, 0, pnITF, False)
    End If
    
    '** Inserta MovCol
    Call loReg.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, pnMontoPagar, 0, psMetliquid, 0, nEstadoActual, pnNewSaldoCap, False)
    '** Inserta MovColDet - Capital
    If pnCapPag > 0 Then
        Call loReg.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gColTransfConceptoCodCapital, 0, pnCapPag, False)
        Call loReg.dAgregaMovOpeVarias(lnMovNro, Round((pnCapPag * 0.03), 2), psCtaCod, "COMISION RECAUDO FOCMACM", lnMoneda, gOtrOpeComisionRecaudoFocmacm, False)
    End If
    '** Inserta MovColDet - Interes Comp
    If pnIntCompPag > 0 Then
        Call loReg.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gColTransfConceptoCodInteresCompensatorio, 0, pnIntCompPag, False)
    End If
    '** Inserta MovColDet - Interes Morat
    If pnIntMoratPag > 0 Then
        Call loReg.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, gColTransfConceptoCodInteresMoratorio, 0, pnIntMoratPag, False)
    End If
    'Para cada Gasto que se tenga (Sin Comision, ni Gasto Administrativo )
    For I = 0 To UBound(pmGastos) - 1
        If pmGastos(I, 5) = "S" Then  ' Si ha sido modificado
            Call loReg.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, pnNroCalend, pmGastos(I, 7), pmGastos(I, 1), Format(CDbl(pmGastos(I, 6)), "#0.00"), False)
        End If
    Next I
    'Abonar a la Cuenta de Ahorro del FOCMACM
    If lsOpeCod = gColRecOpePagTransfFocMacmEfe Then
        lsOpeCodCaptac = "200222"
        lsOpeCodItf = gITFCobroEfectivo
    ElseIf lsOpeCod = gColRecOpePagTransfFocMacmChq Then
        lsOpeCodCaptac = "200223"
        lsOpeCodItf = gITFCobroCheque
    ElseIf lsOpeCod = gColRecOpePagTransfFocMacmVou Then
        lsOpeCodCaptac = "200224"
        lsOpeCodItf = gITFCobroVoucher
    End If
    Call loReg.CapAbonoCuentaAho(lnMovNro, pMatDatosAho, psCtaAhorro, pnMontoPagar, lsOpeCodCaptac, psMovNro, "Deposito Pago FOCMACM", , , , , , , CDate(psFechaHora), "", True, pnITF, False, lsOpeCodItf, pnEstadoTransf, pnCapPag)
    
    loReg.dCommitTrans
    mbTrans = False
Set loReg = Nothing

Exit Sub
ControlError:
    If mbTrans Then
        loReg.dRollbackTrans
        pnMovNro = 0
        mbTrans = False
        Err.Raise vbObjectError + 100, "Error nPagoCreditoTransferido", "Error en Pago Credito Transferido "
    End If
End Sub
Public Sub nActualizarColocTransfAutorizado(ByVal cCtaCod As String)
    Dim loRegRec As COMDColocPig.DCOMColPActualizaBD
    Dim dr As ADODB.Recordset
    Set loRegRec = New COMDColocPig.DCOMColPActualizaBD
    
    On Error GoTo ErrornActualizarColocTransfAutorizado
    loRegRec.dBeginTrans
    mbTrans = True
    
    Call loRegRec.dActualizarColocTransfAutorizado(cCtaCod)
        
    loRegRec.dCommitTrans
    mbTrans = False
    Set loRegRec = Nothing

Exit Sub
ErrornActualizarColocTransfAutorizado:
    If mbTrans Then
        loRegRec.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nActualizarColocTransfAutorizado", "Error en Actualizacion Pago Transferido "
End Sub

Public Sub nExtornoPagoCreditoTransferidos(ByVal psCtaCod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False, Optional ByRef psmensaje As String, _
        Optional ByVal psMovNroCap As String = "", Optional ByVal psCtaAhorro As String = "", _
        Optional ByVal pnMontoCtaAhorro As Currency, Optional ByVal psMotExtorno As Variant)

Dim lsSQL As String, lsSQL2 As String, lsSQL3 As String
Dim loBase As COMDColocPig.DCOMColPActualizaBD

Dim lnMovNro As Long
Dim lsOpeCod As String
Dim loDataExt As COMConecta.DCOMConecta
Dim lrDataExt As ADODB.Recordset
Dim lrDataG As ADODB.Recordset
Dim loFunciones As COMNContabilidad.NCOMContFunciones

Dim lnSaldCap As Currency, lnSaldIntComp As Currency, lnSaldIntMor As Currency, lnSaldGasto As Currency
Dim lnCapPag As Currency, lnIntComPag As Currency, lnIntMorPag As Currency, lnGastoPag As Currency
Dim lsUltMov As String
'**DAOR 20070809***************
Dim lnIntMoraGenerado As Double
Dim lnMora As Double
Dim lnComiAbo As Double
Dim ldFecha As Date
Dim lnGasto As Double
'*******************************
'19-05-2006
Dim nEstadoAnterior As Integer

lsOpeCod = gColRecOpeExtPagRecup

' *** Obtiene Datos de Montos Cubiertos
lsSQL = "EXEC stp_sel_ColocTransfExtCargarConceptosPagados '" & psCtaCod & "'," & pnMovNroAnt
    
lsSQL2 = "EXEC stp_sel_ColocTransfExtCargasConceptosGastosPagados '" & psCtaCod & "'," & pnMovNroAnt
    
Set loDataExt = New COMConecta.DCOMConecta
    loDataExt.AbreConexion
    '** Carga Conceptos Pagados
    Set lrDataExt = loDataExt.CargaRecordSet(lsSQL)
    If lrDataExt Is Nothing Then
        psmensaje = "ERROR: al Buscar datos para Extorno "
        Exit Sub
    End If
    If lrDataExt.BOF And lrDataExt.EOF Then
        psmensaje = "ERROR: No encontro datos de la Operacion a extornar "
        Exit Sub
    End If
    ' Asigna los valores
    lnSaldCap = lrDataExt!nSaldo
    lnSaldIntComp = lrDataExt!nSaldoIntComp
    lnSaldIntMor = lrDataExt!nSaldoIntMor
    lnSaldGasto = lrDataExt!nSaldoGasto
    lnCapPag = lrDataExt!CapPag
    lnIntComPag = lrDataExt!IntComPag
    lnIntMorPag = lrDataExt!IntMorPag
    lnGastoPag = lrDataExt!GastoPag
    lsUltMov = IIf(IsNull(lrDataExt!cUltMov), "@", lrDataExt!cUltMov)

'Proceso para el verdadero valor que debe de ir en nSaldoIntMor de la Tabla ColocRecup***
lnMora = (lnSaldIntMor + lnIntMorPag)
lnSaldIntMor = lnMora
    
lnGasto = (lnSaldGasto + lnGastoPag)
lnSaldGasto = lnGasto
'***************************************************************************************
    nEstadoAnterior = lrDataExt!nCredEstado
    Set lrDataExt = Nothing
    '** Carga los Gastos Cobrados ' Concepto / NroGasto / MontoPagado
    Set lrDataExt = loDataExt.CargaRecordSet(lsSQL2)
    
Set loDataExt = Nothing

' *** Realiza el Extorno
'On Error GoTo Error
Set loBase = New COMDColocPig.DCOMColPActualizaBD
    loBase.dBeginTrans
    mbTrans = True
    
    '** Actualiza ColocTransferidoDet
    Call loBase.dUpdateColocTransferidoDet(pnMovNroAnt)
    
    '** Update Producto ---------> Actualizamos el Estado
    Call loBase.dUpdateProducto(psCtaCod, , lnSaldCap + lnCapPag, nEstadoAnterior, psFechaHora, -2, False)
    
    '** Update COLOCACIONES
    Call loBase.dUpdateColocaciones(psCtaCod, , , , , lsUltMov, , False)
    
    '** Update ColocRecup
    'Call loBase.dUpdateColocRecup(psCtaCod, , , lnSaldIntComp + lnIntComPag, lnSaldIntMor, lnSaldGasto, 0, , , , , False)
    Call loBase.dUpdateColocTransferido(psCtaCod, , , lnSaldIntComp + lnIntComPag, lnSaldIntMor, lnSaldGasto, , , False)
    'Regresa los Gastos a su estado anterior
    Do While Not lrDataExt.EOF
        If lrDataExt!nMontoGasto = 0 Then
            Call loBase.dUpdateColocTransferidoGastos(psCtaCod, lrDataExt!nNroGastoCta, , , lrDataExt!nMontoGasto, gColRecGastoEstPendiente, , False)
        Else
            Call loBase.dUpdateColocTransferidoGastos(psCtaCod, lrDataExt!nNroGastoCta, , , lrDataExt!nMontoGasto - lrDataExt!nMontoPag, gColRecGastoEstPendiente, , False)
        End If
        lrDataExt.MoveNext
    Loop
    
    '** Inserta Mov
    Call loBase.dInsertMov(psMovNro, lsOpeCod, "Ext.Pago.Credito.Transferido", gMovEstContabNoContable, gMovFlagDeExtorno, False)
    
    ' Obtiene nMovNro
    lnMovNro = loBase.dGetnMovNro(psMovNro)
    
    '***CTI3 (ferimoro) 17102018
    Call loBase.dInsertDetExtMov(lnMovNro, lsOpeCod, 0, psCtaCod, False, psMotExtorno(0), psMotExtorno(1))
    
    '** Inserta MovCol
    Call loBase.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 0, pnMonto, 0, "", 0, 0, 0, False)
    
    '** Update Mov Anterior
    Call loBase.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loBase.dInsertMovRef(lnMovNro, pnMovNroAnt, False)
    
    '**
    'mnEjecutaBatch = lobase.dEjecutaBatch
    
    'RIRO20140610 ERS017 *********
    loBase.actualizarRelacionExtornoVoucherCaptacion pnMovNroAnt
    loBase.ActualizaMovPendientesOpe pnMovNroAnt, pnMonto
    'END RIRO ********************
    
    'Extorna Monto de la cuenta de la FOCMACM
    If psMovNroCap <> "" Then
        Call loBase.CapExtornoAbonoAho(psMovNroCap, lnMovNro, pnMovNroAnt, gAhoExtDepPagFocmacm, psCtaAhorro, psMovNro, "", pnMontoCtaAhorro)
    End If
    loBase.dCommitTrans
    mbTrans = False
Set loBase = Nothing

Exit Sub
Error:
    If mbTrans Then
        loBase.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nExtornoTransFRecup", "Error en Funcion de Extorno de Transferencia"

End Sub
Public Function ObtenerCtaAhorroPorCredito(ByVal psCtaCod As String) As Recordset
    Dim oColRec As COMDColocRec.DCOMColRecCredito
    Set oColRec = New COMDColocRec.DCOMColRecCredito
    Set ObtenerCtaAhorroPorCredito = oColRec.getCtaAhorroPorCredito(psCtaCod)
End Function
'FIN FRHU 20150415
