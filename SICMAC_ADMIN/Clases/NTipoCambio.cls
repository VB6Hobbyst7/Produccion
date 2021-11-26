VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NTipoCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim vsConexion As String
Dim vsServerCom As String
Dim vsServerPers As String
Dim vsFormatoFecha As String

Public Function CargaTipoCambio() As ADODB.Recordset
'Carga en un Data los Registros existentes en la Tabla Tipo Cambio
Dim RegTC As New ADODB.Recordset
Dim lnDiasTipoCambio As Integer
Dim oConsSit As NConstSistemas
Dim sql As String
Dim rs As ADODB.Recordset
Dim oconect As DConecta

Set oConsSit = New NConstSistemas
Set oconect = New DConecta
lnDiasTipoCambio = 0
lnDiasTipoCambio = oConsSit.LeeConstSistema(gConstSistDiasTipoCambioMant)

sql = "  SET ROWCOUNT " & lnDiasTipoCambio & "  " _
    & "  Select Convert(VarChar(12), dFecCamb,103)  as  Fecha,  nValVent as Venta, " _
    & "         nValComp as Compra,nValFijoDia as FijoDia, nValFijo as Fijo, dFecCamb  " _
    & "  From " & vsServerCom & "TipoCambio Order BY dFecCamb Desc " _
    & " SET ROWCOUNT 0 "
If oconect.AbreConexion = False Then Exit Function
Set rs = oconect.CargaRecordSet(sql)
Set CargaTipoCambio = rs
Set oConsSit = Nothing
oconect.CierraConexion
Set oconect = Nothing
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

Dim oIni As ClasIni
Set oIni = New ClasIni
vsConexion = oIni.CadenaConexion
vsServerCom = oIni.BaseComunes
vsServerPers = oIni.BasePersonas
Set oIni = Nothing
End Sub
Public Function ExisteTipoCambio(ByVal pdFecha As String) As Boolean
'Verifica si el Tipo de Cambio para la fecha enviada como parámetro ya ha sido ingresada
    Dim rs As New ADODB.Recordset
    Dim oCon As New DConecta
    Dim sql As String
    
    Set oCon = New DConecta
    
    If oCon.AbreConexion = False Then Exit Function
    sql = "Select dFecCamb from " & vsServerCom & "TipoCambio Where DateDiff(day, dFecCamb, '" & pdFecha & "') = 0"
    
    Set rs = oCon.CargaRecordSet(sql)
    If rs.EOF And rs.BOF Then
       ExisteTipoCambio = False
    Else
       ExisteTipoCambio = True
    End If
    rs.Close
    Set rs = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function EmiteTipoCambio(ByVal dFecha As Date, ByVal nTpoTipoCambio As TipoCambio) As Double
    Dim sql As String
    Dim rs As Recordset
    Dim lbCentral As Boolean
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    EmiteTipoCambio = 0
    Set rs = New Recordset
    
    If oCon.AbreConexion = False Then Exit Function
    rs.CursorLocation = adUseClient
    
    sql = " Select nConsSisValor From ConstSistema where nConsSisCod = " & gConstSistBitCentral
    Set rs = oCon.CargaRecordSet(sql)
    
    If rs.EOF And rs.BOF Then
        lbCentral = True
    Else
       If rs.Fields(0) = 1 Then
           lbCentral = True
       Else
           lbCentral = False
       End If
    End If
    
    rs.Close
    
    If lbCentral Then
        'Centralizado
        sql = "Select IsNull(nValFijo,0) nValFijo, IsNull(nValFijoDia,0) nValFijoDia, IsNull(nValVent,0) nValVent, IsNull(nValComp,0) nValComp, IsNull(nValVentEsp,0) nValVentEsp, IsNull(nValCompEsp,0) nValCompEsp, IsNull(nValPond,0) nValPond From TipoCambio " _
            & " WHERE dFecCamb = (   Select Max(dFecCamb)" _
            & "                      From TipoCambio " _
            & "                      Where datediff(day,dFecCamb,'" & Format$(dFecha, gsFormatoFecha) & "')=0)"
        Set rs = oCon.CargaRecordSet(sql)
    Else
        'Distribuido
        sql = "Select nValFijo, nValFijoDia, nValVent, nValComp From DbComunes..TipCambio " _
           & " WHERE dFecCamb = (   Select Max(dFecCamb)" _
           & "                      From DbComunes..TipCambio " _
           & "                      Where datediff(day,dFecCamb,'" & Format$(dFecha, gsFormatoFecha) & "')=0)"
        Dim oConecta As DConecta
        Set oConecta = New DConecta
        
        oConecta.AbreConexion 'Remota "01"
        Set rs = oConecta.CargaRecordSet(sql)
    End If
    

    Set rs.ActiveConnection = Nothing
    If Not rs.EOF And Not rs.BOF Then
        Select Case nTpoTipoCambio
            Case TCFijoMes
                EmiteTipoCambio = rs("nValFijo")
            Case TCFijoDia
                EmiteTipoCambio = rs("nValFijoDia")
            Case TCVenta
                EmiteTipoCambio = rs("nValVent")
            Case TCCompra
                EmiteTipoCambio = rs("nValComp")
            Case TCVentaEsp
                EmiteTipoCambio = rs("nValVentEsp")
            Case TCCompraEsp
                EmiteTipoCambio = rs("nValCompEsp")
            Case TCPonderado
                EmiteTipoCambio = rs("nValPond")
        End Select
    End If
    rs.Close
    Set rs = Nothing
End Function

Public Function GrabaTipoCambio(ByVal psFormatoFecha As String, ByVal pdFecCamb As Date, ByVal pnValVent As Currency, ByVal pnValComp As Currency, _
                                    ByVal pnValVentEsp As Currency, ByVal pnValCompEsp As Currency, _
                                    ByVal pnValFijo As Currency, ByVal pnValFijoDia As Currency, ByVal pnValPonderado As Currency, _
                                    ByVal psUltimaActualizacion As String, _
                                    Optional ByVal pbEjectBatch As Boolean)
    Dim oTipoCambio As dTipoCambio
    Set oTipoCambio = New dTipoCambio
    oTipoCambio.Inicio psFormatoFecha
    oTipoCambio.ActualizaTipoCambioDiario pdFecCamb, psUltimaActualizacion, pnValFijoDia, pbEjectBatch
    oTipoCambio.InsertaTipoCambio pdFecCamb, pnValVent, pnValComp, pnValVentEsp, pnValCompEsp, pnValFijo, pnValFijoDia, pnValPonderado, psUltimaActualizacion, pbEjectBatch
    
    If pbEjectBatch Then
        GrabaTipoCambio = oTipoCambio.EjecutaBatch
    Else
        GrabaTipoCambio = 0
    End If
End Function

Public Function ActualizaTipoCambio(ByVal psFormatoFecha As String, ByVal pdFecCamb As Date, ByVal pnValVent As Currency, ByVal pnValComp As Currency, _
                                    ByVal pnValVentEnt As Currency, ByVal pnValCompEst As Currency, _
                                    ByVal pnValFijo As Currency, ByVal pnValFijoDia As Currency, ByVal pnValPonderado As Currency, ByVal psUltimaActualizacion As String, _
                                    Optional ByVal pbEjectBatch As Boolean)
Dim oTipoCambio As dTipoCambio
Set oTipoCambio = New dTipoCambio

oTipoCambio.Inicio psFormatoFecha
oTipoCambio.ActualizaTipoCambio pdFecCamb, psUltimaActualizacion, pnValVent, pnValComp, pnValVentEnt, pnValCompEst, pnValFijo, pnValFijoDia, pnValPonderado, pbEjectBatch

If pbEjectBatch Then
    ActualizaTipoCambio = oTipoCambio.EjecutaBatch
Else
    ActualizaTipoCambio = 0
End If

End Function




