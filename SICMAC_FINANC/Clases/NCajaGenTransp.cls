VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCajaGenTransp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim vsServerCom As String
Dim vsServerPers As String
Enum TiposMontosTransp
    NetoServicio = 0
    IGV = 1
    MontoServicio = 2
End Enum
Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As ClasIni
Set oIni = New ClasIni
vsServerCom = oIni.BaseComunes
vsServerPers = oIni.BasePersonas
Set oIni = Nothing
End Sub
Public Function GetTranspTipo(ByVal pnTipoTrans As CGTipoTransporte) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta

Set oCon = New DConecta
Set rs = New ADODB.Recordset

If oCon.AbreConexion = False Then Exit Function
sql = "Select   CONVERT(CHAR(200),UPPER(ISNULL(P.cPersNombre , T.cDescTrans))) AS Descripcion, T.cTranspCod " _
    & " From    Transporte T " _
    & "         JOIN CONSTANTE C ON C.nConsValor = T.cTpoTrans " _
    & "         LEFT JOIN PERSONA P ON P.cPersCod = T.cPersCod " _
    & " WHERE   C.nCONSCOD = '" & gCGTipoTransporte & "' AND T.cTpoTrans='" & pnTipoTrans & "' "

Set rs = oCon.CargaRecordSet(sql)
Set GetTranspTipo = rs
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function GetDatosTransp(ByVal psCodTrans As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta

Set oCon = New DConecta
Set rs = New ADODB.Recordset

If oCon.AbreConexion = False Then Exit Function

sql = " Select   T.cTranspCod , T.cTpoTrans, P.cPersNombre , T.cDescTrans, cConsDescripcion, " _
    & "         nMontoMin , nCostoMontoMin, nCantidad, nCostoCant " _
    & " From    Transporte T " _
    & "         JOIN CONSTANTE C ON C.nConsValor = T.cTpoTrans " _
    & "         LEFT JOIN PERSONA P ON P.cPersCod = T.cPersCod " _
    & " WHERE   C.nCONSCOD = 4023 AND T.cTranspCod='" & psCodTrans & "' "

Set rs = oCon.CargaRecordSet(sql)
Set GetDatosTransp = rs
oCon.CierraConexion
Set oCon = Nothing
End Function

Function GetMontosServicio(ByVal psCodTransp As String, ByVal pnMontoTrans As Currency, ByVal pnTipoMonto As TiposMontosTransp, ByVal pnMoneda As Moneda, ByVal pnTipoCambio As Currency) As Currency
Dim rs As ADODB.Recordset
Dim lnNetoServ As Currency
Dim lnIGV As Currency
Dim lnMontoServ As Currency
Dim lnMontoMin As Currency
Dim lnCostoMontoMin As Currency
Dim lnCantidad As Currency
Dim lnCostoCant As Currency
Dim lnMontoTrans As Currency

Dim oImpuesto As DImpuesto
Set oImpuesto = New DImpuesto

Dim oCon As NConstSistemas
Set oCon = New NConstSistemas

Dim lnIGVValor As Currency
lnIGVValor = 1 + (oImpuesto.CargaImpuestoFechaValor(oCon.LeeConstSistema(40), CDate(oCon.LeeConstSistema(15))) / 100)

Set rs = New ADODB.Recordset
lnMontoMin = 0
lnCostoMontoMin = 0
lnCantidad = 0
lnCostoCant = 0
lnNetoServ = 0
lnMontoServ = 0
lnIGV = 0
Set rs = GetDatosTransp(Trim(psCodTransp))
If Not rs.EOF And Not rs.BOF Then
    lnMontoMin = rs!nMontoMin
    lnCostoMontoMin = rs!nCostoMontoMin
    lnCantidad = rs!nCantidad
    lnCostoCant = rs!nCostoCant
End If
rs.Close
Set rs = Nothing
If pnMoneda = gMonedaNacional Then
    lnMontoTrans = pnMontoTrans
Else
    lnMontoTrans = pnMontoTrans * pnTipoCambio
End If
If CCur(lnMontoTrans) <= lnMontoMin Then
    lnNetoServ = lnCostoMontoMin
Else
    lnNetoServ = lnMontoTrans * lnCostoCant
End If
lnMontoServ = Format(lnNetoServ * lnIGVValor, "#,#0.00")
lnIGV = Format(lnMontoServ - lnNetoServ, "#,#0.00")
GetMontosServicio = 0
Select Case pnTipoMonto
    Case NetoServicio
        GetMontosServicio = lnNetoServ
    Case IGV
        GetMontosServicio = lnIGV
    Case MontoServicio
        GetMontosServicio = lnMontoServ
End Select
End Function


