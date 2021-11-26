VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NRepFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
'** Nombre : NRepFormula
'** Descripción : Módulo de clase para Notas Estado segun ERS052-2013
'** Creación : EJVG, 20130416 09:00:00 AM
'********************************************************************

Option Explicit

Public Function RegistrarNotasEstado(ByVal psOpeCod As String, ByVal pMatNotas As Variant, pdFecha As Date, ByVal psAgeCod As String, psUserCod As String) As Boolean
    Dim oMov As New DMov
    Dim oRep As DRepFormula
    Dim lsMovNro As String
    Dim iCab As Long, iDet As Long, lnID As Long
    Dim lbTrans As Boolean
    Dim MatNotasDet As Variant
    
    'Nota Estado *******************
    Dim lsDescripcion As String
    Dim lsFormula As String
    Dim lbPeriodo As Boolean
    Dim lbComentario As Boolean
    Dim lnNotaEstado As Integer 'EJVG20140102
    'Nota Estado Detalle
    Dim lnTipoDetalle As Integer
    Dim lsDesc As String
    Dim lnNivel As Integer
    Dim lbNegrita As Boolean
    Dim lsFormula1 As String
    Dim lsFormula1_2012 As String
    Dim lsFormula2 As String
    Dim lsFormula2_2012 As String
    Dim lsFormula3 As String
    Dim lsFormula3_2012 As String
    Dim lsFormula4 As String
    Dim lsFormula4_2012 As String
    Dim lsFormula5 As String
    Dim lsFormula5_2012 As String
    '*******************************
    
    On Error GoTo ErrRegistrarNotasEstado

    lsMovNro = oMov.GeneraMovNro(pdFecha, psAgeCod, psUserCod)

    Set oRep = New DRepFormula
    Call oRep.dBeginTrans
    lbTrans = True

    lnID = oRep.InsertarNotasEstadoConfig(psOpeCod, lsMovNro)
    If lnID = 0 Then Exit Function
    
    For iCab = 1 To UBound(pMatNotas, 2)
        lsDescripcion = pMatNotas(1, iCab)
        lsFormula = pMatNotas(2, iCab)
        lbPeriodo = pMatNotas(3, iCab)
        lbComentario = pMatNotas(4, iCab)
        lnNotaEstado = pMatNotas(6, iCab) 'EJVG20140102
        
        Call oRep.InsertarNotasEstadoConfigDet(lnID, psOpeCod, iCab, lsDescripcion, lsFormula, lbPeriodo, lbComentario)
        If lnNotaEstado > 0 Then 'EJVG20140102
            Call oRep.ActualizaNotaEstadoxNotaComplementaria(lnID, psOpeCod, iCab, lnNotaEstado)
        End If
        
        Set MatNotasDet = Nothing
        MatNotasDet = pMatNotas(5, iCab)
        
        For iDet = 1 To UBound(MatNotasDet, 2)
            lnTipoDetalle = CInt(MatNotasDet(1, iDet))
            lsDesc = Trim(CStr(MatNotasDet(2, iDet)))
            lnNivel = CInt(MatNotasDet(3, iDet))
            lbNegrita = IIf(MatNotasDet(4, iDet) = "1", True, False)
            lsFormula1 = Trim(CStr(MatNotasDet(5, iDet)))
            lsFormula1_2012 = Trim(CStr(MatNotasDet(6, iDet)))
            lsFormula2 = Trim(CStr(MatNotasDet(7, iDet)))
            lsFormula2_2012 = Trim(CStr(MatNotasDet(8, iDet)))
            lsFormula3 = Trim(CStr(MatNotasDet(9, iDet)))
            lsFormula3_2012 = Trim(CStr(MatNotasDet(10, iDet)))
            lsFormula4 = Trim(CStr(MatNotasDet(11, iDet)))
            lsFormula4_2012 = Trim(CStr(MatNotasDet(12, iDet)))
            lsFormula5 = Trim(CStr(MatNotasDet(13, iDet)))
            lsFormula5_2012 = Trim(CStr(MatNotasDet(14, iDet)))
            
            Call oRep.InsertarNotasEstadoConfigDetDet(lnID, psOpeCod, iCab, iDet, lnTipoDetalle, lsDesc, lbNegrita, lnNivel, lsFormula1, lsFormula2, lsFormula3, lsFormula4, lsFormula5, lsFormula1_2012, lsFormula2_2012, lsFormula3_2012, lsFormula4_2012, lsFormula5_2012)
        Next
    Next

    oRep.dCommitTrans
    lbTrans = False
    RegistrarNotasEstado = True
    
    Set oMov = Nothing
    Set oRep = Nothing
    Exit Function
ErrRegistrarNotasEstado:
    If lbTrans Then
        oRep.dRollbackTrans
        Set oRep = Nothing
    End If
End Function
Public Function RecuperaConfigRepNotasEstado(ByVal psOpeCod As String) As ADODB.Recordset
    Dim oRep As New DRepFormula
    Set RecuperaConfigRepNotasEstado = oRep.RecuperaConfigRepNotasEstado(psOpeCod)
    Set oRep = Nothing
End Function
Public Function RecuperaConfigRepNotasEstadoDetalle(ByVal psOpeCod As String, ByVal pnId As Long, ByVal pnNotaEstado As Long) As ADODB.Recordset
    Dim oRep As New DRepFormula
    Set RecuperaConfigRepNotasEstadoDetalle = oRep.RecuperaConfigRepNotasEstadoDetalle(psOpeCod, pnId, pnNotaEstado)
    Set oRep = Nothing
End Function
Public Function ObtenerUltimoNivelConfig(ByVal psOpeCod As String) As Integer
    Dim oRep As New DRepFormula
    ObtenerUltimoNivelConfig = oRep.ObtenerUltimoNivelConfig(psOpeCod)
    Set oRep = Nothing
End Function
'FRHU 20131223 RQ13657
Public Function RecuperaConfigFlujoEfectivo() As ADODB.Recordset
    Dim oRep As New DRepFormula
    Set RecuperaConfigFlujoEfectivo = oRep.RecuperaConfigFlujoEfectivo()
    Set oRep = Nothing
End Function
Public Sub RegistrarConfigFlujoEfectivo(ByVal pnId As Integer, ByVal pcDescripcion As String, ByVal pnNivel As Integer, ByVal pnTipo As Integer, ByVal pcValor As String, ByVal pbPeriodo As Boolean, ByVal pbEstado As Boolean, ByVal pnOrden As Integer)
    Dim oRep As New DRepFormula
    Dim bValor As Boolean
    Call oRep.RegistrarConfigFlujoEfectivo(pnId, pcDescripcion, pnNivel, pnTipo, pcValor, pbPeriodo, pbEstado, pnOrden)
    Set oRep = Nothing
End Sub
Public Function RecuperaHojaTrabajoFE(ByVal pnTipo As Integer) As ADODB.Recordset
    Dim oRep As New DRepFormula
    Set RecuperaHojaTrabajoFE = oRep.RecuperaHojaTrabajoFE(pnTipo)
    Set oRep = Nothing
End Function
Public Function RegistrarHojaTrabajoFE(ByVal pMatNotas As Variant) As Boolean
    Dim oRep As DRepFormula
    Dim iCab As Long
    'Flujo de Efectivo *******************
    Dim lnID As Integer
    Dim lnTipo As Integer
    Dim lnNivel As Integer
    Dim lsDescripcion As String
    Dim lsForMen As String 'lsForMen2013
    Dim lsForMay As String 'lsForMay2013
    Dim lsAjusDebe As String
    Dim lsAjusHaber As String
    Dim lsOperDebe As String
    Dim lsOperHaber As String
    Dim lsInveDebe As String
    Dim lsInveHaber As String
    Dim lsFinaDebe As String
    Dim lsFinaHaber As String
    Dim lbEstado As Boolean
    Dim lnOrden As Integer
    Dim lbMov As Boolean
    '*******************************
    Set oRep = New DRepFormula
    For iCab = 1 To UBound(pMatNotas, 2)
        lnID = CInt(pMatNotas(1, iCab))
        lnTipo = CInt(pMatNotas(2, iCab))
        lnNivel = CInt(pMatNotas(3, iCab))
        lsDescripcion = Trim(CStr(pMatNotas(4, iCab)))
        lsForMen = Trim(CStr(pMatNotas(5, iCab)))
        lsForMay = Trim(CStr(pMatNotas(6, iCab)))
        lsAjusDebe = Trim(CStr(pMatNotas(7, iCab)))
        lsAjusHaber = Trim(CStr(pMatNotas(8, iCab)))
        lsOperDebe = Trim(CStr(pMatNotas(9, iCab)))
        lsOperHaber = Trim(CStr(pMatNotas(10, iCab)))
        lsInveDebe = Trim(CStr(pMatNotas(11, iCab)))
        lsInveHaber = Trim(CStr(pMatNotas(12, iCab)))
        lsFinaDebe = Trim(CStr(pMatNotas(13, iCab)))
        lsFinaHaber = Trim(CStr(pMatNotas(14, iCab)))
        lbEstado = pMatNotas(15, iCab)
        lnOrden = CInt(pMatNotas(16, iCab))
        lbMov = pMatNotas(17, iCab)
        
        Call oRep.InsertarHojasDeTrabajoFE(lnID, lnTipo, lnNivel, lsDescripcion, lsForMen, lsForMay, lsAjusDebe, lsAjusHaber, lsOperDebe, lsOperHaber, lsInveDebe, lsInveHaber, lsFinaDebe, lsFinaHaber, lbEstado, lnOrden, lbMov)
    Next
    RegistrarHojaTrabajoFE = True
    
    Set oRep = Nothing
    Exit Function
End Function
Public Sub ModificarEstadoHojaTrabajo(ByVal pnId As Integer)
    Dim oRep As New DRepFormula
    Dim bValor As Boolean
    Call oRep.ModificarEstadoHojaTrabajo(pnId)
    Set oRep = Nothing
End Sub
'FIN FRHU 20131223 RQ13657
'EJVG20140102 ***
Public Function GuardarNotasComplementarias(ByVal psOpeCod As String, ByRef pMatDatos() As Variant) As Boolean
    Dim oRep As New DRepFormula
    Dim lbTrans As Boolean
    Dim lnNotaEstado As Integer, lnSeccion As Integer, lnColumna As Integer
    Dim lsFormula As String, lsFormula_2012 As String
    Dim i As Integer
    Dim lsOpeCod As String
    
    On Error GoTo ErrGuardarNotasComplementaria
    Call oRep.dBeginTrans
    lbTrans = True
    
    oRep.LimpiaNotaEstadosxNuevasNotasComplementaria psOpeCod
    For i = 1 To UBound(pMatDatos, 2)
        lnNotaEstado = pMatDatos(1, i)
        lnSeccion = pMatDatos(2, i)
        lnColumna = pMatDatos(3, i)
        lsFormula = pMatDatos(4, i)
        lsFormula_2012 = pMatDatos(5, i)
        oRep.ActualizaNotaComplementaria psOpeCod, lnNotaEstado, lnSeccion, lnColumna, lsFormula, lsFormula_2012
    Next
    oRep.dCommitTrans
    lbTrans = False
    GuardarNotasComplementarias = True

    Set oRep = Nothing
    Exit Function
ErrGuardarNotasComplementaria:
    If lbTrans Then
        oRep.dRollbackTrans
        Set oRep = Nothing
    End If
    GuardarNotasComplementarias = False
End Function
'Public Function ListaNotaComplementarias(ByVal psOpeCod As String, ByVal pnSeccion As Integer) As ADODB.Recordset
'    Dim oRep As New DRepFormula
'    Set ListaNotaComplementarias = oRep.ListaNotaComplementarias(psOpeCod, pnSeccion)
'    Set oRep = Nothing
'End Function
Public Function ListaNotaComplementariasAll(ByVal psOpeCod As String) As ADODB.Recordset
    Dim oRep As New DRepFormula
    Set ListaNotaComplementariasAll = oRep.ListaNotaComplementariasAll(psOpeCod)
    Set oRep = Nothing
End Function
Public Function ListaCalidadCrediticiaActivosFinancieros(ByVal pdFecha As Date, ByVal pbResumen As Boolean) As ADODB.Recordset
    Dim oRep As New DRepFormula
    Set ListaCalidadCrediticiaActivosFinancieros = oRep.ListaCalidadCrediticiaActivosFinancieros(pdFecha, pbResumen)
    Set oRep = Nothing
End Function
Public Function ListaCreditosVencidosxDiasAtraso(ByVal pdFecha As Date) As ADODB.Recordset
    Dim oRep As New DRepFormula
    Set ListaCreditosVencidosxDiasAtraso = oRep.ListaCreditosVencidosxDiasAtraso(pdFecha)
    Set oRep = Nothing
End Function
Public Function ListaConcentracionRiesgoxSector(ByVal pdFecha As Date) As ADODB.Recordset
    Dim oRep As New DRepFormula
    Set ListaConcentracionRiesgoxSector = oRep.ListaConcentracionRiesgoxSector(pdFecha)
    Set oRep = Nothing
End Function
'END EJVG *******
'EJVG20140909 ***
Public Function GrabarProyeccionxOperacion(ByVal psOpeCod As String, ByVal pnAnio As Integer, ByVal poRS As ADODB.Recordset) As Boolean
    Dim obj As New DRepFormula
    Dim bTrans As Boolean
    On Error GoTo ErrGrabarProyecionxOperacion
    
    obj.dBeginTrans
    bTrans = True
    
    obj.LimpiarProyeccionxOperacion psOpeCod, pnAnio
    Do While Not poRS.EOF
        obj.InsertarProyeccionxOperacion psOpeCod, pnAnio, poRS!Nota, 1, CCur(poRS!Enero)
        obj.InsertarProyeccionxOperacion psOpeCod, pnAnio, poRS!Nota, 2, CCur(poRS!Febrero)
        obj.InsertarProyeccionxOperacion psOpeCod, pnAnio, poRS!Nota, 3, CCur(poRS!Marzo)
        obj.InsertarProyeccionxOperacion psOpeCod, pnAnio, poRS!Nota, 4, CCur(poRS!Abril)
        obj.InsertarProyeccionxOperacion psOpeCod, pnAnio, poRS!Nota, 5, CCur(poRS!Mayo)
        obj.InsertarProyeccionxOperacion psOpeCod, pnAnio, poRS!Nota, 6, CCur(poRS!Junio)
        obj.InsertarProyeccionxOperacion psOpeCod, pnAnio, poRS!Nota, 7, CCur(poRS!Julio)
        obj.InsertarProyeccionxOperacion psOpeCod, pnAnio, poRS!Nota, 8, CCur(poRS!Agosto)
        obj.InsertarProyeccionxOperacion psOpeCod, pnAnio, poRS!Nota, 9, CCur(poRS!Septiembre)
        obj.InsertarProyeccionxOperacion psOpeCod, pnAnio, poRS!Nota, 10, CCur(poRS!Octubre)
        obj.InsertarProyeccionxOperacion psOpeCod, pnAnio, poRS!Nota, 11, CCur(poRS!Noviembre)
        obj.InsertarProyeccionxOperacion psOpeCod, pnAnio, poRS!Nota, 12, CCur(poRS!Diciembre)
        poRS.MoveNext
    Loop
    
    obj.dCommitTrans
    bTrans = False
    GrabarProyeccionxOperacion = True
    Exit Function
ErrGrabarProyecionxOperacion:
    GrabarProyeccionxOperacion = False
    If bTrans Then
        obj.dRollbackTrans
        Set obj = Nothing
    End If
End Function
'END EJVG *******
