VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nCredCrediPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function VerificaReporteGenerado(ByVal pdFecSis As Date) As Boolean
Dim Sql As String
Dim Co As DConecta
Dim Rs As New ADODB.Recordset

Set Co = New DConecta
Sql = "Select Count(cCtaCod) as Contad From ColocCrediPagoArcCobranza Where Datediff(d,dFecha,'" & Format(pddFecSis, "mm/dd/yyyy") & "') = 0 "
Co.AbreConexion
Set Rs = Co.CargaRecordSet(Sql)
VerificaReporteGenerado = False
If Rs!Contad > 0 Then
    If MsgBox("Ya se ha generado Archivo de cobranza para hoy día. Desea Reemplazarlo ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
        lsSQL = "Delete ColocCrediPagoArcCobranza Where Datediff(d,dFecha,'" & Format(pdFecSis, "mm/dd/yyyy") & "' ) = 0 "
        Co.Ejecutar (Sql)
        VerificaReporteGenerado = True
    Else
        VerificaReporteGenerado = False
        Exit Function
    End If
End If
Co.CierraConexion
Rs.Close
Set Co = Nothing
Set Rs = Nothing
End Function

Public Sub DeleteColocCrediPagoArcResultado(ByVal pdFechaArchivoTXT As Date)
Dim Sql As String
Dim Co As DConecta
Dim Rs As New ADODB.Recordset

Set Co = New DConecta
Sql = "DELETE ColocCrediPagoArcResultado Where datediff(d, dFecha, '" & Format(pdFechaArchivoTXT, "mm/dd/yyyy") & "' ) = 0"
Co.AbreConexion
Co.Ejecutar (Sql)
Co.CierraConexion
Rs.Close
Set Co = Nothing
Set Rs = Nothing
End Sub


Public Function ObtieneMoraProyectada(ByVal psCodCta As String, _
            ByVal pnNroCuo As Integer, pnDiasAtraso As Integer, _
            ByVal pnNroCalen As Integer) As Double

Dim lsSQL As String
Dim reg As New ADODB.Recordset
Dim regGasto As New ADODB.Recordset
Dim lnMontoDesemb As Double, lnMontoMora As Double
Dim lnMontoGasto As Double
Dim lnMontoMoraProyectada As Double
Dim Co As DConecta
Dim nCred As NCredito
Set nCred = New NCredito

Set Co = New DConecta


lnMontoGasto = 0
'Credito
lsSQL = " Select Col.cCtaCod, Col.nMontoCol, Cred.nDiasAtraso, Cred.cProtesto,"
lsSQL = lsSQL & " ntasaMorat =(Select nTasaini from ColocLineaCreditoTasa where nColocLinCredTasaTpo=3 and cLineaCred=Col.cLineaCred),"
lsSQL = lsSQL & " nCapCuoPend =(Select (nMonto-nMontopagado) from  ColocCalendDet where cCtacod='" & psCodCta & "' and nNroCalen=" & pnNroCalen & " and nCuota=" & pnNroCuo & " and  nPrdConceptoCod=1000)"
lsSQL = lsSQL & " from Colocaciones Col"
lsSQL = lsSQL & " Inner Join ColocacCred Cred on Cred.cCtaCod=Col.cCtaCod"
lsSQL = lsSQL & " Where Col.cCtaCod='" & psCodCta & "'"

Co.AbreConexion
Set reg = Co.CargaRecordSet(lsSQL)
Co.CierraConexion

If reg.BOF And reg.EOF Then
    lnMontoMoraProyectada = 0
Else
    lnMontoDesemb = reg!nMontoCol
    If pnDiasAtraso <= -2 Then
        lnMontoMoraProyectada = 0
    Else
        'Mora Proyectada         - Mora x (Dias Atraso + 2)
        lnMontoMora = (reg!nTasaMorat * 1 / 100 * (reg!nCapCuoPend)) * (pnDiasAtraso + 2)
        'Gastos Dias Generacion del Gastos
        If pnDiasAtraso > -2 And pnDiasAtraso <= 0 Then
            'Gastos 1 Dia de Atraso
            
            'lsSQL = " Select IsNull(nValorGas,0) nValorGas " _
            '      & " From Gastos where cCodGasto like '018010%'  AND nRanIni <= " & lnMontoDesemb _
            '      & " AND nRanFin >= " & lnMontoDesemb & " AND cMoneda = '" & Mid(psCodCta, 6, 1) & "' "
            
            lsSQL = "Select isnull(nValor,0) nValorGas"
            lsSQL = lsSQL & " from ProductoConcepto where nPrdConceptoCod like '12__'"
            lsSQL = lsSQL & " and nInicial<=" & lnMontoDesemb & " and nFinal>=" & lnMontoDesemb & " and nMoneda=substring(psCodCta,9,1)"
            Co.AbreConexion
            Set regGasto = Co.CargaRecordSet(lsSQL)
            Co.CierraConexion
            lnMontoGasto = regGasto!nValorGas
            regGasto.Close
        ElseIf pnDiasAtraso > 28 And pnDiasAtraso <= 29 Then
            'Gastos 30 Dia de Atraso
            'lsSQL = " Select IsNull(nValor,0) nValorGas " _
            '      & " From Gastos where cCodGasto like '018011%'  AND nRanIni <= " & lnMontoDesemb _
            '      & " AND nRanFin >= " & lnMontoDesemb & " AND cMoneda = '" & Mid(psCodCta, 6, 1) & "' "
            'regGasto.Open lsSQL, pcConex, adOpenStatic, adLockReadOnly, adCmdText
            
            lsSQL = "Select isnull(nValor,0) nValorGas "
            lsSQL = lsSQL & " from ProductoConcepto where nPrdConceptoCod like '1235%'"
            lsSQL = lsSQL & " and nInicial<=" & lnMontoDesemb & " and nFinal>=" & lnMontoDesemb & " and nMoneda=substring(psCodCta,9,1)"
            Co.AbreConexion
            Set regGasto = Co.CargaRecordSet(lsSQL)
            Co.CierraConexion
            
            lnMontoGasto = regGasto!nValorGas
            regGasto.Close
        End If
        If pnDiasAtraso = 7 And IsNull(reg!cProtesto) Then
            'Protesto
            lsSQL = " Select IsNull(nValorGas,0) nValorGas " _
                  & " From ProductoConcepto " _
                  & " Where nPrdConceptoCod = '" & IIf(Mid(reg!cCtaCod, 9, 1) = "1", "1219", "1220") & "' "
            
            Co.AbreConexion
            Set regGasto = Co.CargaRecordSet(lsSQL)
            Co.CierraConexion
        End If
        lnMontoMoraProyectada = lnMontoGasto + lnMontoMora
    End If
End If
reg.Close

ObtieneMoraProyectada = lnMontoMoraProyectada

End Function


Public Function ProcesaDatos(ByVal pdFecSis As Date, ByVal psAgencias As String) As ADODB.Recordset
Dim Sql As String
Dim Co As DConecta
Set Co = New DConecta
Sql = "Select Cred.cCtaCod,  Per.cPersNombre,"
Sql = Sql & "       Cred.nNroProxCuota,"
Sql = Sql & "       Cred.nNroCalen, Cal.nCuota, Cal.dVenc,"
Sql = Sql & " (Select Max(nCuota) From ColocCalendario where cCtaCod=Pro.cCtaCod and  nNroCalen=Cred.nNroCalen) NroCuotas,"
Sql = Sql & " (Select (sum(nMonto) - sum (nMontoPagado) ) Cuota From ColocCalendDet Where nNroCalen=Cred.nNroCalen and cCtaCod=Cred.cCtaCod and nCuota=Cal.nCuota) Cuota"
Sql = Sql & "       from  ColocacCred Cred"
Sql = Sql & "  Inner Join ColocCalendario Cal  on Cal.cCtaCod = Cred.cCtaCod"
Sql = Sql & "  Inner Join ColocCrediPago Credi on Credi.cCtaCod=Cred.cCtaCod"
Sql = Sql & "  Inner Join Producto Pro on Pro.cCtaCod = Cred.cCtaCod"
Sql = Sql & "  Inner Join ProductoPersona PP on PP.cCtaCod =  Cred.cCtaCod"
Sql = Sql & "  Inner Join Persona Per on Per.cPersCod=PP.cPersCod"
Sql = Sql & "  Where Cal.nNroCalen = Cred.nNroCalen And Cal.nColocCalendApl = 1 And Cal.nColocCalendEstado = 0"
Sql = Sql & "  and Pro.nPrdEstado in (" & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor & ") and Credi.cCredipago='S'"
Sql = Sql & "  and ( Cred.nNroProxCuota=Cal.nCuota or Cal.dVenc<='" & Format(pdFecSis, "mm/dd/yyyy") & "') and PP.nPrdPersRelac='20'"
Sql = Sql & "  and substring(cred.cCtaCod,4,2) in " & psAgencias
Sql = Sql & "  Order by Cred.cCtaCod, Cal.nNroCalen, Cal.nCuota"
Co.AbreConexion
Set Rs = Co.CargaRecordSet(Sql)
Co.CierraConexion
Set ProcesaDatos = Rs
Set Rs = Nothing
Set Co = Nothing
End Function


Public Function Busca_x_Cliente(ByVal psCodPers As String) As ADODB.Recordset
Dim Sql As String
Dim Co As DConecta
Dim Rs As New ADODB.Recordset
Set Co = New DConecta

Sql = "Select Pro.cCtaCod  from Producto Pro " & _
        " Inner Join ProductoPersona PP on Pro.cCtaCod=PP.cCtaCod " & _
        " Where Pro.nPrdEstado in (" & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor & "," & gColocEstSug & "," & gColocEstAprob & ") and " & _
        " PP.nPrdPersRelac=" & ColocRelacPers.gColRelPersTitular & " and PP.cPersCod='" & psCodPers & "'"
Co.AbreConexion
Set Rs = Co.CargaRecordSet(Sql)
Co.CierraConexion
Set Busca_x_Cliente = Rs
Set Rs = Nothing
Set Co = Nothing
End Function

Public Function EstadoCred(ByVal psCodCta As String) As ADODB.Recordset
Dim Sql As String
Dim Co As DConecta
Dim Rs As New ADODB.Recordset
Set Co = New DConecta
Sql = "SELECT * FROM Producto WHERE cCtaCod = '" & psCodCta & "'"
Co.AbreConexion
Set Rs = Co.CargaRecordSet(Sql)
Co.CierraConexion
Set EstadoCred = Rs
Set Rs = Nothing
Set Rs = Nothing
End Function

Public Function DatosCred(ByVal psCodCta As String) As ADODB.Recordset
Dim Sql As String
Dim Co As DConecta
Dim Rs As New ADODB.Recordset
Set Co = New DConecta
Sql = "Select Pro.cCtaCod, Pro.nSaldo nSaldoCap, Col.nMonto nMontoApr, Cred.cMetLiquidacion cMetLiquid, " & _
        " Cred.nNroProxCuota,Col.nCuotas nCuotasApr,(Select cUser From RRHH where  cPersCod=PP.cPersCod) cCodAnalista " & _
        " From Producto Pro " & _
        " Inner Join ColocacCred Cred on Pro.cCtaCod=Cred.cCtaCod " & _
        " Inner join ColocacEstado Col on Col.cCtaCod=Cred.cCtaCod " & _
        " Inner join ProductoPersona PP on PP.cCtaCod=Cred.cCtaCod " & _
        " Where PP.nPrdPersRelac ='" & ColocRelacPers.gColRelPersAnalista & "' And Col.nPrdEstado = 2002  and Pro.cCtaCod='" & psCodCta & "'"

Co.AbreConexion
Set Rs = Co.CargaRecordSet(Sql)
Co.CierraConexion
Set DatosCred = Rs
Set Rs = Nothing
End Function

Public Function GetCrediPago(ByVal psCodCta As String) As ADODB.Recordset
Dim Sql As String
Dim Co As DConecta
Dim Rs As New ADODB.Recordset
Set Co = New DConecta
Sql = "SELECT *  " _
         & "FROM ColocCrediPago WHERE cCtaCod='" & psCodCta & "'"
Co.AbreConexion
Set Rs = Co.CargaRecordSet(Sql)
Co.CierraConexion
Set GetCrediPago = Rs
Set Rs = Nothing
Set Co = Nothing
End Function

Public Sub ActualizaCrediPago(ByVal psCodCta As String, psCredipago As String)
Dim Sql As String
Dim Co As DConecta
Set Co = New DConecta
Co.AbreConexion
Sql = "UPDATE ColocCrediPago SET cCrediPago = '" & IIf(psCredipago = "1", "S", "N") & "'  WHERE cCtaCod ='" & psCodCta & "'"
Co.Ejecutar (Sql)
Co.CierraConexion
Set Co = Nothing
End Sub
Public Sub InsertaColocCrediPagoArcCobranza(ByVal pdFecSis As Date, ByVal psCtaCod As String, _
ByVal pdFecVenc As Date, ByVal pnMonto As Double, ByVal pnMora As Double, _
ByVal pnNroCuo As Integer)
Dim Sql As String
Dim Co As DConecta
Sql = " INSERT INTO ColocCrediPagoArcCobranza (dFecha, cCodCta, dFecVenc, nMonto, nMora, nNroCuo) " _
                & " VALUES ('" & Format(pdFecSis, "mm/dd/yyyy") & "','" & psCtaCod & "','" _
                & Format(pdFecVenc, "mm/dd/yyyy") & "'," & pnMonto & "," _
                & pnMora & ", 0   ) "

Set Co = New DConecta
Co.AbreConexion
Co.Ejecutar (Sql)
Co.CierraConexion
Set Co = Nothing
End Sub

Public Sub InsertCrediPago(ByVal psCodCta As String, ByVal psCredipago As String, _
ByVal psCodUser As String, ByVal pdFecSis As Date)
Dim Sql As String
Dim Co As DConecta
Set Co = New DConecta
Sql = "INSERT INTO ColocCrediPago (cCtaCod, cCrediPago, cCodUsu, dFecMod) " _
& " Values ('" & psCodCta & "','" & IIf(psCredipago = "1", "S", "N") & "','" _
& psCodUser & "','" & Format(pdFecSis, "mm/dd/yyyy") & "' ) "
Co.AbreConexion
Co.Ejecutar (Sql)
Co.CierraConexion
Set Co = Nothing
End Sub

Public Sub InsertColocCrediPagoArcResultado(ByVal pdFecha As Date, ByVal psCodCta As String, ByVal pdFecPago As Date, _
ByVal pnMonto As Double, ByVal pnMora As Double, ByVal pnMontPgo As Double, ByVal psSucursal As String, _
ByVal psReferncia As String)
Dim Sql As String
Dim Co As DConecta
Set Co = New DConecta
lsSQL = "Insert into ColocCrediPagoArcResultado " _
                & " (dFecha, cCtaCod, dFecPago, nMonto, nMora, nMontoPag, cSucursal, cReferencia ) " _
                & " VALUES ('" & Format(pdFecha, "mm/dd/yyyy") & "','" & Trim(lstCampos.ListItems(i)) & "','" _
                & Format(Trim(pdFecPago), "mm/dd/yyyy") & "'," & Trim(pnMonto) & "," _
                & Trim(pnMora) & "," & Trim(pnMontoPag) & "," _
                & "'" & psSucursal & "','" & Trim(psReferncia) & "' ) "
Co.AbreConexion
Co.Ejecutar (Sql)
Co.CierraConexion
Set Co = Nothing
End Sub


Public Function GetNroCta(Optional ByVal pBOpt As Boolean = True) As String
Dim Sql As String
Dim Rs As New ADODB.Recordset
Dim RCD As nRcdReportes
Set RCD = New nRcdReportes

If pBOpt Then
    Set Rs = RCD.GetQuery("Select nConsSisValor from ConstSistema where nConsSisCod=73 ")
Else
    Set Rs = RCD.GetQuery("Select nConsSisValor from ConstSistema where nConsSisCod=74 ")
End If

If Rs.EOF And Rs.BOF Then
    GetNroCta = ""
Else
    GetNroCta = Rs!nConsSisValor
End If
Set Rs = Nothing
Set RCD = Nothing
End Function
