VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCartaFianza"
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
    ByVal pdRegistro As Date, ByVal oRelPers As UCredRelacion, _
    ByVal psPersCodAna As String, ByVal pdFecAsig As Date, _
    ByVal pdFecVenc As Date, ByVal pnMonto As Double, psMovAct As String, _
    ByVal psNumFuente As String, ByVal pdPersEval As Date, ByVal psPersCodAcre As String)
    
Dim oBase As DCredActualizaBD
Dim sMovNro As String

    On Error GoTo ErrorNuevaCF
    Set oBase = New DCredActualizaBD
    oBase.dBeginTrans
        
    Call oBase.dInsertProducto(psNuevaCta, 0#, 0#, gColocEstSolic, pdRegistro, 0, False)
    oRelPers.IniciarMatriz
    Do While Not oRelPers.EOF
        Call oBase.dInsertProductoPersona(psNuevaCta, oRelPers.ObtenerCodigo, CInt(oRelPers.ObtenerValorRelac), False)
        oRelPers.siguiente
    Loop
    Call oBase.dInsertProductoPersona(psNuevaCta, psPersCodAcre, gColRelPersAcreedor, False)
    
    Call oBase.dInsertProductoPersona(psNuevaCta, psPersCodAna, gColRelPersAnalista, False)
    Call oBase.dInsertColocaciones(psNuevaCta, 0, pdFecVenc, pnMonto, gColocCalendCodPFCF, "", pdFecAsig, psMovAct, False)
    Call oBase.dInsertColocFteIngreso(psNumFuente, psNuevaCta, pdPersEval)
    Call oBase.dInsertColocCartaFianza(psNuevaCta, pnCondicion, pnModalidad, pdFecAsig, pdFecVenc, psFinalidad)
    Call oBase.dInsertColocCFEstado(psNuevaCta, pdRegistro, gColocEstSolic, pnMonto, pdFecVenc, "Registro Carta Fianza", 0, False)
    
    
    oBase.dCommitTrans
    Set oBase = Nothing
    
    
    Exit Sub
    
ErrorNuevaCF:
    oBase.dRollbackTrans
    Set oBase = Nothing
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Public Sub nCFActualizaSolicitud(ByVal psNuevaCta As String, ByVal psCodAge As String, ByVal pnTipoProd As Producto, _
    ByVal pnMoneda As Moneda, ByVal pnCondicion As ColocCredCondicion, _
    ByVal pnModalidad As ColCFModalidad, ByVal psFinalidad As String, _
    ByVal pdRegistro As Date, ByVal oRelPers As UCredRelacion, _
    ByVal psPersCodAna As String, ByVal pdFecAsig As Date, _
    ByVal pdFecVenc As Date, ByVal pnMonto As Double, psMovAct As String, _
    ByVal psNumFuente As String, ByVal pdPersEval As Date, ByVal psPersCodAcre As String)

Dim oBase As DCredActualizaBD
Dim sMovNro As String

    Set oBase = New DCredActualizaBD
    oBase.dBeginTrans
    
    Call oBase.dDeleteProductoPersona(psNuevaCta)
    oRelPers.IniciarMatriz
    Do While Not oRelPers.EOF
        Call oBase.dInsertProductoPersona(psNuevaCta, oRelPers.ObtenerCodigo, CInt(oRelPers.ObtenerValorRelac), False)
        oRelPers.siguiente
    Loop
    Call oBase.dInsertProductoPersona(psNuevaCta, psPersCodAcre, gColRelPersAcreedor, False)
    Call oBase.dInsertProductoPersona(psNuevaCta, psPersCodAna, gColRelPersAnalista, False)
    Call oBase.dUpdateColocaciones(psNuevaCta, , pdFecVenc, pnMonto, , , pdFecAsig)
    Call oBase.dDeleteColocFteIngreso(psNuevaCta)
    Call oBase.dInsertColocFteIngreso(psNumFuente, psNuevaCta, pdPersEval)
    Call oBase.dUpdateColocCartaFianza(psNuevaCta, pnCondicion, pnModalidad, pdFecAsig, pdFecVenc, psFinalidad)
    Call oBase.dDeleteColocCFEstado(psNuevaCta, gColocEstSolic)
    Call oBase.dInsertColocCFEstado(psNuevaCta, pdRegistro, gColocEstSolic, pnMonto, pdFecVenc, "Registro Carta Fianza", 0, False)
    
    oBase.dCommitTrans
    Set oBase = Nothing
    
End Sub

'** Sugerencia de Carta Fianza
Public Sub nCFSugerencia(ByVal psCtaCod As String, ByVal psFechaHora As String, _
    ByVal pdFecVencSug As Date, ByVal pnMontoSug As Double)

Dim loBase As DCredActualizaBD
Dim lnMovNro As Long


'On Error GoTo ErrorModCF
Set loBase = New DCredActualizaBD
    loBase.dBeginTrans
    mbTrans = True

    '** Actualiza Producto
    Call loBase.dUpdateProducto(psCtaCod, , , gColocEstSug, psFechaHora, -2, False)       ' (-2) aumenta el ste transac
    
    '** Actualiza Colocaciones
    Call loBase.dUpdateColocaciones(psCtaCod, , pdFecVencSug, pnMontoSug, , , , False)
    
    '** Actualiza CartaFianza
    Call loBase.dUpdateColocCartaFianza(psCtaCod, , , , pdFecVencSug, , , False)

    '** Insert ColocCFEstado
    Call loBase.dDeleteColocCFEstado(psCtaCod, gColocEstSug)
    Call loBase.dInsertColocCFEstado(psCtaCod, psFechaHora, gColocEstSug, pnMontoSug, pdFecVencSug, "Sugerencia Carta Fianza", 0, False)
    
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
    Err.Raise vbObjectError + 100, "Error nCFSugerencia ", "Error en Funcion de Sugerencia CF "
End Sub

'** Aprobacion de Carta Fianza
Public Sub nCFAprobacion(ByVal psCtaCod As String, ByVal psFechaHora As String, _
    ByVal pdFecVencApr As Date, ByVal pnMontoApr As Double, ByVal psApoderado As String)

Dim loBase As DCredActualizaBD
Dim lnMovNro As Long

'On Error GoTo ErrorModCF
Set loBase = New DCredActualizaBD
    loBase.dBeginTrans
    mbTrans = True

    '** Actualiza Producto
    Call loBase.dUpdateProducto(psCtaCod, , pnMontoApr, gColocEstAprob, psFechaHora, -2, False)          ' (-2) aumenta el ste transac
    '*** Insert ProductoPersona - Apoderado
    Call loBase.dInsertProductoPersona(psCtaCod, psApoderado, gColRelPersApoderado, False)
    
    '** Actualiza Colocaciones
    Call loBase.dUpdateColocaciones(psCtaCod, , pdFecVencApr, pnMontoApr, , , , False)
    
    '** Actualiza CartaFianza
    Call loBase.dUpdateColocCartaFianza(psCtaCod, , , , pdFecVencApr, , , False)

    '** Insert ColocCFEstado
    Call loBase.dInsertColocCFEstado(psCtaCod, psFechaHora, gColocEstAprob, pnMontoApr, pdFecVencApr, "Aprobacion Carta Fianza", 0, False)
    
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
    Err.Raise vbObjectError + 100, "Error nCFSugerencia ", "Error en Funcion de Sugerencia CF "
End Sub

'** Rechazo de Carta Fianza
Public Sub nCFRechazo(ByVal psCtaCod As String, ByVal psFechaHora As String, _
    ByVal psComenta As String, ByVal pnMotivoRechazo As Integer)

Dim loBase As DCredActualizaBD
Dim lnMovNro As Long

'On Error GoTo ErrorModCF
Set loBase = New DCredActualizaBD
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
    Err.Raise vbObjectError + 100, "Error nCFSugerencia ", "Error en Funcion de Sugerencia CF "
End Sub


'** Emsión de Carta Fianza
Public Sub nCFEmision(ByVal psCtaCod As String, ByVal psFechaHora As String, ByVal psMovNro As String, _
    ByVal pdFecVencEmi As Date, ByVal pnMontoEmi As Double)

Dim loBase As DCredActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = gColCFOpeEmision

'On Error GoTo ErrorModCF
Set loBase = New DCredActualizaBD
    loBase.dBeginTrans
    mbTrans = True

    '** Actualiza Producto
    Call loBase.dUpdateProducto(psCtaCod, , , gColocEstVigNorm, psFechaHora, -2, False)              ' (-2) aumenta el ste transac
   
    '** Actualiza Colocaciones
    Call loBase.dUpdateColocaciones(psCtaCod, , , , , , psFechaHora, False)
    
    '** Actualiza CartaFianza
    'Call loBase.dUpdateColocCartaFianza(psCtaCod, , , ,  pdFecVencEmi, , , False)

    '** Insert ColocCFEstado
    Call loBase.dInsertColocCFEstado(psCtaCod, psFechaHora, gColocEstVigNorm, pnMontoEmi, pdFecVencEmi, "Emision Carta Fianza", 0, False)
    
    '** Inserta Mov
    Call loBase.dInsertMov(psMovNro, lsOpeCod, "Emision Carta Fianza", gMovEstContabMovContable, gMovFlagVigente, False)
    ' Obtiene nMovNro
    lnMovNro = loBase.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loBase.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 0, pnMontoEmi, 0, "", 0, 0, gColocEstVigNorm, False)

    '** Inserta MovColDet -  Monto Carta Fianza
    Call loBase.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 0, gColCFConceptoCodCapital, 0, pnMontoEmi, False)

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
    Err.Raise vbObjectError + 100, "Error nCFSugerencia ", "Error en Funcion de Sugerencia CF "
End Sub

'** Comision de Carta Fianza
Public Sub nCFComision(ByVal psCtaCod As String, ByVal psFechaHora As String, ByVal psMovNro As String, _
    ByVal pnMontoComi As Double, _
    ByVal pbITFAplica As Boolean, ByVal pbITFAsumidoCreditos As Boolean, _
    ByVal pnMontoITf As Double)

Dim loBase As DCredActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = gColCFOpeComisEfe

'On Error GoTo ErrorModCF
Set loBase = New DCredActualizaBD
    loBase.dBeginTrans
    mbTrans = True

    '** Actualiza Producto
    Call loBase.dUpdateProducto(psCtaCod, , , , psFechaHora, -2, False)              ' (-2) aumenta el ste transac
   
    '** Actualiza Colocaciones
    'Call loBase.dUpdateColocaciones(psCtaCod, , , , , ,  psFechaHora, False)
    
    '** Actualiza CartaFianza
    'Call loBase.dUpdateColocCartaFianza(psCtaCod, , , ,  pdFecVencEmi, , , False)

    '** Inserta Mov
    Call loBase.dInsertMov(psMovNro, lsOpeCod, "Comision Carta Fianza", gMovEstContabMovContable, gMovFlagVigente, False)
    ' Obtiene nMovNro
    lnMovNro = loBase.dGetnMovNro(psMovNro)
    
    '** Inserta MovCol
    Call loBase.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 0, pnMontoComi, 0, "", 0, 0, 0, False)

    '** Inserta MovColDet -  Monto Carta Fianza
    Call loBase.dInsertMovColDet(lnMovNro, lsOpeCod, psCtaCod, 0, gColCFConceptoCodComision, 0, pnMontoComi, False)
    
    
    
    
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
    Err.Raise vbObjectError + 100, "Error nCFSugerencia ", "Error en Funcion de Sugerencia CF "
End Sub

Public Sub nCFExtornoComision(ByVal psCtaCod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnMovNroAnt As Long, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False)

'** Actualiza Producto ( NroTrans)
'** Actualiza Colocaciones (Ultima Modificacion)
'** dInsertMov
'** dInsertMovCol
'************************************

Dim lsSQL As String
Dim loBase As DCredActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = gColCFOpeExtComis

'On Error GoTo ErrorModCF
Set loBase = New DCredActualizaBD
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
    
    '** Inserta MovCol
    Call loBase.dInsertMovCol(lnMovNro, lsOpeCod, psCtaCod, 0, pnMonto, 0, "", 0, 0, 0, False)
    
    '** Update Mov Anterior
    Call loBase.dUpdateMov(pnMovNroAnt, , , , gMovFlagExtornado, False)

    '** Insert Mov Ref
    Call loBase.dInsertMovRef(lnMovNro, pnMovNroAnt, False)

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
    Err.Raise vbObjectError + 100, "Error nCFExtornoComision", "Error en Funcion de Extorno de Comision"

End Sub


'** Retirar Carta Fianza Aprobada
Public Sub nCFRetirar(ByVal psCtaCod As String, ByVal psFechaHora As String, _
    ByVal psComenta As String, ByVal pnMotivoRechazo As Integer)

Dim loBase As DCredActualizaBD
Dim lnMovNro As Long

'On Error GoTo ErrorModCF
Set loBase = New DCredActualizaBD
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
    Err.Raise vbObjectError + 100, "Error nCFSugerencia ", "Error en Funcion de Sugerencia CF "
End Sub

'** Devolución de Carta Fianza Emitida
Public Sub nCFDevolucion(ByVal psCtaCod As String, ByVal psFechaHora As String, ByVal psMovNro As String, _
    ByVal psComenta As String, ByVal pnMotivoDevolucion As Integer, ByVal pnMontoCF As Double)

Dim loBase As DCredActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = gColCFOpeDevol

'On Error GoTo ErrorModCF
Set loBase = New DCredActualizaBD
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
    Err.Raise vbObjectError + 100, "Error nCFDevolucion ", "Error en Funcion de Devolucion CF "
End Sub


'** Honrar Carta Fianza
Public Sub nCFHonrar(ByVal psCtaCod As String, ByVal psFechaHora As String, ByVal psMovNro As String, _
    ByVal psComenta As String, ByVal pnMotivoDevolucion As Integer, ByVal pnMontoCF As Double)

Dim loBase As DCredActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = "140700"

'On Error GoTo ErrorModCF
Set loBase = New DCredActualizaBD
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
    Err.Raise vbObjectError + 100, "Error nCFHonrar ", "Error en Funcion de Honrar CF "
End Sub


'** Relaciona Credito con Carta Fianza a Honrar
Public Sub nCFHonrarCredito(ByVal psCtaCod As String, ByVal psFechaHora As String, ByVal psMovNro As String, _
    ByVal pmCreditos As Variant)

Dim loBase As DCredActualizaBD
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim lnCred As Integer
lsOpeCod = "140800"

'On Error GoTo ErrorModCF
Set loBase = New DCredActualizaBD
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
    Err.Raise vbObjectError + 100, "Error nCFHonrar ", "Error en Funcion de Honrar CF "
End Sub







