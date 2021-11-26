VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMPoliza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim oCon As COMConecta.DCOMConecta

Public Function RecuperaNumeroPoliza() As String

Dim rs As ADODB.Recordset


Set rs = oCon.CargaRecordSet("SELECT cNumPoliza=ISNULL(MAX(convert(int,cNumPoliza)),0) FROM Poliza")

If rs.EOF Then
    RecuperaNumeroPoliza = "00000001"
Else
    RecuperaNumeroPoliza = Format(rs!cNumPoliza + 1, "00000000")
End If


End Function

'peac 20071224
'Public Function RecuperaDatosConstantes(ByVal pNum As Integer)
'Dim sSql As Integer
'
'sSql = " INSERT INTO Poliza(cNumPoliza,cPersCodContr,cPersCodAseg,nMontoPrimaTotal,nTipoPoliza,nEstado,dRegistro, bPolizaExterna, nSumaAsegurada,cNumCertif,nPrimaNeta) "
'sSql = sSql & " VALUES('" & pcNumPoliza & "','" & pcPersCodContr & "','" & pcPersCodAseg & "'," & pnMontoPrima & "," & pnTipoPoliza & ",2,'" & Format(pdRegistro, "mm/dd/yyyy") & "'," & pnPolizaExterna & "," & pnSumaAsegurada & "," & pcNumCertif & "," & pnPrimaNeta & " )"
'
'oCon.Ejecutar sSql

'End Function


Public Function RegistroPoliza(ByVal pcNumPoliza As String, _
                                ByVal pcPersCodContr As String, _
                                ByVal pcPersCodAseg As String, _
                                ByVal pnMontoPrima As Double, _
                                ByVal pnTipoPoliza As Integer, _
                                ByVal pdRegistro As Date, _
                                ByVal pnPolizaExterna As Integer, _
                                ByVal pnSumaAsegurada As Double, _
                                ByVal pcNumCertif As String, _
                                ByVal pnPrimaNeta As Double, _
                                ByVal pnMoneda As Integer, ByVal pnNumAnio As Integer, ByVal iTipoPago As Integer)
                                'peac 20071128 se agrego "pcNumCertif y pnPrimaNeta"

Dim sSql As String

'peac 20071128 se agregó " & pcNumCertif & "," & pnPrimaNeta & " )"
'sSql = " INSERT INTO Poliza(cNumPoliza,cCodPolizaAseg,cPersCodContr,cPersCodAseg,nMontoPrimaTotal,nTipoPoliza,nEstado,dRegistro, bPolizaExterna, nSumaAsegurada,nPrimaNeta,nMoneda,nNumAnio) "
'sSql = sSql & " VALUES('" & pcNumPoliza & "','" & pcNumCertif & "','" & pcPersCodContr & "','" & pcPersCodAseg & "'," & pnMontoPrima & "," & pnTipoPoliza & ",2,'" & Format(pdRegistro, "mm/dd/yyyy") & "'," & pnPolizaExterna & "," & pnSumaAsegurada & "," & pnPrimaNeta & "," & pnMoneda & "," & pnNumAnio & " )"

'*** PEAC 20080619
sSql = "exec stp_ins_Polizas '" & pcNumPoliza & "','" & pcNumCertif & "','" & pcPersCodContr & "','" & pcPersCodAseg & "'," & pnMontoPrima & "," & pnTipoPoliza & ",'" & Format(pdRegistro, "yyyymmdd") & "'," & pnPolizaExterna & "," & pnSumaAsegurada & "," & pnPrimaNeta & "," & pnMoneda & "," & pnNumAnio & ", " & iTipoPago & ""

oCon.Ejecutar sSql

End Function

'*** PEAC 20080412
Public Function VerificaTasasPolizas(ByVal pcAgeCod As String, ByVal pnInmueble As Integer, _
                                    ByVal pnAlternativa As Integer, Optional ByVal pnMueble As Integer = 0) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    sSql = "exec stp_sel_VerificaTasasPolizas '" & pcAgeCod & "'," & pnInmueble & "," & pnAlternativa & "," & pnMueble

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set VerificaTasasPolizas = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function


'peac 20071226
Public Function RegistraTasasPolizas(ByVal pcAgeCod As String, _
                                ByVal pnInmueble As Integer, ByVal pnMonto As Double, _
                                ByVal pnPriMinima As Double, ByVal pnDereEmi As Double, _
                                ByVal pnAlternativa As Integer, ByVal pdRegistro As Date, _
                                Optional ByVal pnMueble As Integer = 0)

Dim sSql As String

sSql = " INSERT INTO PolizaCotiza(cAgeCod,nInmueble,nTna,nPrimaNetaMinima,ndrchoEmision,nTipo,dregistro,nMueble) "
sSql = sSql & " VALUES('" & pcAgeCod & "'," & pnInmueble & "," & pnMonto & "," & pnPriMinima & "," & pnDereEmi & "," & pnAlternativa & ",'" & Format(pdRegistro, "mm/dd/yyyy") & "'" & "," & pnMueble & " )"

oCon.Ejecutar sSql

End Function

'peac 20071226
Public Function ModificaTasasPolizas(ByVal pcAgeCod As String, _
                                ByVal pnInmueble As Integer, ByVal pnMonto As Double, _
                                ByVal pnPriMinima As Double, ByVal pnDereEmi As Double, _
                                ByVal pnAlternativa As Integer, ByVal pdRegistro As Date, _
                                Optional ByVal pnMueble As Integer = 0)
                                '*** BRGO 20111205 Se agregó el parámetro pnMueble

Dim sSql As String

sSql = " UPDATE PolizaCotiza SET "
sSql = sSql & " nTna=" & pnMonto & ", nPrimaNetaMinima=" & pnPriMinima & ", ndrchoEmision=" & pnDereEmi & ", dregistro='" & Format(pdRegistro, "mm/dd/yyyy") & "' "

sSql = Left(sSql, Len(sSql) - 1)
sSql = sSql & " WHERE cAgeCod='" & pcAgeCod & "' and nInmueble=" & pnInmueble & " and nMueble = " & pnMueble
oCon.Ejecutar sSql

End Function

Public Function ModificaPoliza(ByVal pcNumPoliza As String, _
                                Optional ByVal pcPersCodContr As String = "", _
                                Optional ByVal pcPersCodAseg As String = "", _
                                Optional ByVal pnMontoPrima As Double = -1, _
                                Optional ByVal pnTipoPoliza As Integer = -1, _
                                Optional ByVal pnEstado As Integer = -1, _
                                Optional ByVal pdRegistro As Date = 0, _
                                Optional ByVal pdAprobacion As Date = 0, _
                                Optional ByVal pdPago As Date = 0, _
                                Optional ByVal pdVigenciaAseg As Date = 0, _
                                Optional ByVal pdVencAseg As Date = 0, _
                                Optional ByVal pcCodigoPoliza As String = "", _
                                Optional ByVal pnPolizaExterna As Integer = -1, _
                                Optional ByVal pnSumaAsegurada As Double = -1, _
                                Optional ByVal pcNumCertif As String = "", _
                                Optional ByVal pnPrimaNeta As Double = -1, _
                                Optional ByVal pnMoneda As Integer = -1, _
                                Optional ByVal pnNumAnio As Integer = -1, _
                                Optional ByVal pnTotalTC As Double = -1, _
                                Optional ByVal pnTipoPago As Integer = -1)
                                'MAVM 20090820 pnTipoPago, pcCtaCod
                                '*** PEAC 20080512 pnMoneda

Dim sSql As String

sSql = " UPDATE Poliza SET "
If pcPersCodContr <> "" Then
    sSql = sSql & " cPersCodContr='" & pcPersCodContr & "',"
End If
If pcPersCodAseg <> "" Then
    sSql = sSql & " cPersCodAseg='" & pcPersCodAseg & "',"
End If
If pnMontoPrima <> -1 Then
    sSql = sSql & " nMontoPrimaTotal=" & pnMontoPrima & ","
End If
If pnTipoPoliza <> -1 Then
    sSql = sSql & " nTipoPoliza=" & pnTipoPoliza & ","
End If
'*** PEAC 20080512
If pnMoneda <> -1 Then
    sSql = sSql & " nMoneda=" & pnMoneda & ","
End If
If pnMoneda <> -1 Then
    sSql = sSql & " nNumAnio=" & pnNumAnio & ","
End If

If pnEstado <> -1 Then
    sSql = sSql & " nEstado=" & pnEstado & ","
End If
If pdRegistro <> 0 Then
    sSql = sSql & " dRegistro='" & Format(pdRegistro, "mm/dd/yyyy") & "',"
End If
If pdAprobacion <> 0 Then
    sSql = sSql & " dAprobacion='" & Format(pdAprobacion, "mm/dd/yyyy") & "',"
End If
If pdPago <> 0 Then
    sSql = sSql & " dPago='" & Format(pdPago, "mm/dd/yyyy") & "',"
End If
If pdVigenciaAseg <> 0 Then
    sSql = sSql & " dVigenciaAseg='" & Format(pdVigenciaAseg, "mm/dd/yyyy") & "',"
End If
If pdVencAseg <> 0 Then
    sSql = sSql & " dVencAseg='" & Format(pdVencAseg, "mm/dd/yyyy") & "',"
End If

'If pcCodigoPoliza <> "" Then
'    sSql = sSql & " cCodPolizaAseg='" & pcCodigoPoliza & "',"
'End If

If pnPolizaExterna <> -1 Then
    sSql = sSql & " bPolizaExterna=" & pnPolizaExterna & ","
End If

'peac 20071128
If pcNumCertif <> "" Then
    sSql = sSql & " cCodPolizaAseg='" & pcNumCertif & "',"
End If

'peac 20071128
If pnPrimaNeta <> -1 Then
    sSql = sSql & " nPrimaNeta=" & pnPrimaNeta & ","
End If

'MAVM 20090820
If pnTipoPago <> -1 Then
    sSql = sSql & " iTipoPago=" & pnTipoPago & ","
End If
'Fin MAVM 20090820

'*** PEAC 20080626
If pnTotalTC <> -1 Then
    sSql = sSql & " nTotal=" & pnTotalTC & ","
End If

If pnSumaAsegurada <> -1 Then
    sSql = sSql & " nSumaAsegurada=" & pnSumaAsegurada & ","
End If

sSql = Left(sSql, Len(sSql) - 1)
sSql = sSql & " WHERE cNumPoliza='" & pcNumPoliza & "'"
oCon.Ejecutar sSql

End Function

Public Function EliminaPoliza(ByVal pcNumPoliza As String)

Dim sSql As String


sSql = " DELETE FROM Poliza" & _
       " WHERE cNumPoliza='" & pcNumPoliza & "'"

oCon.Ejecutar sSql

End Function

'peac 20071226
Public Function EliminaTasasPoliza(ByVal pcAgeCod As String, _
                                ByVal pnInmueble As Integer, Optional ByVal pnMueble As Integer = 0)
                                'BRGO 20111205 Se agregó pnMueble

Dim sSql As String

sSql = " DELETE FROM PolizaCotiza " & _
       " WHERE cAgeCod='" & pcAgeCod & "' and nInmueble=" & pnInmueble & " and nMueble = " & pnMueble
oCon.Ejecutar sSql

End Function

'*** BRGO 27/08/2010
'MADM 20110511 - 20110430
Public Function RecuperaPolizasCuenta(ByVal pcCtaCod As String) As ADODB.Recordset
    Dim sSql As String
'    sSql = " Select P.cNumPoliza, cCodPolizaAseg, C.cCtaCod "
'    sSql = sSql & " From Colocaciones C "
'    sSql = sSql & " INNER JOIN ColocGarantia CG on C.cCtaCod = CG.cCtaCod "
'    sSql = sSql & " INNER JOIN Garantias G on G.cNumGarant = CG.cNumGarant "
'    sSql = sSql & " INNER JOIN GarantPoliza GP on GP.cNumGarant = G.cNumGarant and GP.nEstado = 1 "
'    sSql = sSql & " INNER JOIN Poliza P on P.cNumPoliza = GP.cNumPoliza and P.nEstado = 2 "
'    sSql = sSql & " INNER JOIN Persona Pe ON P.cPersCodContr=Pe.cPersCod "
'    sSql = sSql & " Where C.cCtaCod = '" & pcCtaCod & "' "
    sSql = "EXEC stp_sel_RecuperaPolizasCuenta '" & pcCtaCod & "' "
    
    Set RecuperaPolizasCuenta = oCon.CargaRecordSet(sSql)
End Function
'END MADM
'*** End BRGO

Public Function RecuperaPolizas(ByVal pnTipoBusqueda As Integer, _
                                Optional ByVal pcPersCodCont As String = "", _
                                Optional ByVal pdInicio As Date = 0, _
                                Optional ByVal pdFin As Date = 0, _
                                Optional ByVal pnEstadoPoliza As Integer = -1) As ADODB.Recordset
Dim sSql As String

sSql = "SELECT cNumPoliza,cCodPolizaAseg,cPersNombre,cConsDescripcion,CONVERT(varchar(10),dRegistro,103),P.cPersCod FROM Poliza " & _
           " INNER JOIN Persona P ON  cPersCodContr=p.cPersCod INNER JOIN Constante ON nConsValor= nTipoPoliza AND nConsCod=9066 " & _
        " WHERE "
If pnTipoBusqueda = 0 Then
    sSql = sSql & " P.cPersCod='" & pcPersCodCont & "' "
Else
    sSql = sSql & " DATEDIFF(dd,dRegistro,'" & Format(pdInicio, "mm/dd/yyyy") & "') <=0 AND DATEDIFF(dd,dRegistro,'" & Format(pdFin, "mm/dd/yyyy") & "') >=0 "
    'sSQL = sSQL & " DATEDIFF(dd,dRegistro,'" & pdInicio & "') >=0 AND DATEDIFF(dd,dRegistro,'" & pdFin & "') <=0 "
End If

If pnEstadoPoliza <> -1 Then
    sSql = sSql & " AND Poliza.nEstado=" & pnEstadoPoliza
End If

Set RecuperaPolizas = oCon.CargaRecordSet(sSql)

End Function
'MADM 20110005
Public Function RecuperaEstadoPolizasListado(ByVal pcNumGarant As String) As ADODB.Recordset
Dim sSql As String

sSql = " exec stp_sel_RecuperaEstadoPolizasListado '" & pcNumGarant & "' "

Set RecuperaEstadoPolizasListado = oCon.CargaRecordSet(sSql)

End Function

Public Function ActualizaEstadoPolizasListado(ByVal pcNumGarant As String) As ADODB.Recordset
Dim sSql As String

sSql = " exec stp_upd_EstadoGarantPoliza '" & pcNumGarant & "' "

oCon.Ejecutar sSql

End Function
'END MADM

'peac 20071226
Public Function RecuperaTasasPolizas() As ADODB.Recordset 'BRGO Esto es para inmuebles
Dim sSql As String

sSql = "exec stp_sel_CargaDatosRecuperaTasasPolizas"

Set RecuperaTasasPolizas = oCon.CargaRecordSet(sSql)

End Function
'*** BRGO 20111205 ***********************************************
Public Function RecuperaTasasPolizasMobiliaria() As ADODB.Recordset
Dim sSql As String

sSql = "exec stp_sel_CargaDatosRecuperaTasasPolizasMobiliaria"

Set RecuperaTasasPolizasMobiliaria = oCon.CargaRecordSet(sSql)

End Function
'*** END BRGO *****************************************************

'peac 20071226
Public Function CargaDatosTasasPoliza(ByVal pcAgencia As String, ByVal pnInmueble As Integer, Optional ByVal pnMueble As Integer = 0) As ADODB.Recordset
Dim sSql As String

sSql = " select a.cAgeDescripcion,pc.* "
sSql = sSql & " from PolizaCotiza pc "
sSql = sSql & " inner join agencias a on pc.cAgeCod=a.cAgeCod "
sSql = sSql & " WHERE pc.cAgeCod='" & pcAgencia & "' and pc.nInmueble=" & pnInmueble & " and pc.nMueble=" & pnMueble 'BRGO 20111205

Set CargaDatosTasasPoliza = oCon.CargaRecordSet(sSql)
End Function

'peac 20071226
Public Function CargaDatosPoliza(ByVal pcNumPoliza As String) As ADODB.Recordset
Dim sSql As String

sSql = "exec stp_sel_CargaDatosPoliza '" & pcNumPoliza & "'"

Set CargaDatosPoliza = oCon.CargaRecordSet(sSql)
End Function
'CargaDatosPrimaNetaPoliza

'peac 20071228
Public Function CargaDatosPrimaNetaPoliza(ByVal pcNumPoliza As String, ByVal pnSumaAseg As Double, ByVal pdFecSis As Date) As ADODB.Recordset
Dim sSql As String

sSql = "exec stp_sel_CargaDatosPrimaNetaPoliza '" & pcNumPoliza & "'," & pnSumaAseg & ",'" & Format(pdFecSis, "yyyymmdd") & "'"

Set CargaDatosPrimaNetaPoliza = oCon.CargaRecordSet(sSql)

End Function

Public Function Garantias_x_Titular(ByVal pcPersCodContr As String) As ADODB.Recordset
Dim sSql As String

'sSql = "SELECT GR.cNumGarant,P.cPersNombre,CONVERT(varchar(10),GRT.dTasacion,103) FROM GarantReal GR INNER JOIN GarantRealTasacion GRT ON GR.cNumGarant = GRT.cNumGarant AND GRT.dTasacion=(SELECT MAX(dTasacion)FROM GarantRealTasacion WHERE cNumGarant=GR.cNumGarant)" & _
           " INNER JOIN PersGarantia PG ON PG.cNumGarant = GR.cNumGarant AND nRelacion =1 " & _
           " INNER JOIN Persona P ON  PG.cPersCod=P.cPersCod " & _
           " WHERE PG.cPersCod='" & pcPersCodContr & "'"

'*** PEAC 20090730
sSql = " exec stp_sel_Garantias_Por_Titular '" & pcPersCodContr & "' "

Set Garantias_x_Titular = oCon.CargaRecordSet(sSql)
End Function
'MADM 20111005 - STORE
Public Function Garantias_x_Poliza(ByVal pcNumPoliza As String) As ADODB.Recordset
Dim sSql As String

'sSql = "SELECT GP.cNumGarant,C.cPersNombre,CONVERT(varchar(10),GP.dTasacion,103) FROM Poliza P" & _
'           " INNER JOIN GarantPoliza GP ON P.cNumPoliza = GP.cNumPoliza " & _
'           " INNER JOIN Persona C ON  C.cPersCod= P.cPersCodContr " & _
'           " WHERE P.cNumPoliza='" & pcNumPoliza & "' AND GP.nEstado=1"
           
 sSql = "Exec stp_sel_ObtenerGarantiasxPoliza '" & pcNumPoliza & "' "
           
Set Garantias_x_Poliza = oCon.CargaRecordSet(sSql)
End Function

'MADM 20111005 - STORE - MAVM 20090920 Arreglo Mod Polizas
Public Function GarantiasPorPersona(ByVal psPersCod As String) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
    
On Error GoTo GarantiasPorPersona
'        sSql = "Select G.cNumGarant, cDescripcion , P.cPersNombre"
'        sSql = sSql & ", dTasacion = Convert(varchar(10),IsNull((SELECT MAX(dTasacion)FROM GarantRealTasacion WHERE cNumGarant = G.cNumGarant), '1990-01-01'),103)"
'        sSql = sSql & " From Garantias G Inner Join PersGarantia PG on G.cNumGarant = PG.cNumGarant"
'        sSql = sSql & "                  Inner Join Persona P ON P.cPersCod = PG.cPersCod"
'        sSql = sSql & " WHERE PG.nRelacion = 1 and  G.nTpoGarantia = 1 and G.nEstado not in (6) and PG.cPersCod = '" & psPersCod & "'"
         
         sSql = "Exec stp_sel_ObtenerGarantiasxPersona '" & psPersCod & "' "
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set GarantiasPorPersona = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
GarantiasPorPersona:
        Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'MAVM 20091004
Public Function ValidaTasacion(ByVal psNumGarantia As String) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
    
On Error GoTo ValidaTasacion
        'sSql = "Select cNumGarant, dTasacion "
        'sSql = sSql & " from GarantRealTasacion "
        'sSql = sSql & " where cnumgarant = '" & psNumGarantia & "'"
        sSql = "exec stp_sel_ObtenerPolizaRealTasa '" & psNumGarantia & "' "
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set ValidaTasacion = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ValidaTasacion:
        Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function Adiciona_GarantPoliza(ByVal pcNumPoliza As String, _
                                    ByVal pcNumGarant As String, _
                                    ByVal pdTasacion As Date, _
                                    ByVal pdEstado As Date) As Boolean

Dim sSql As String
Dim rs As ADODB.Recordset

Adiciona_GarantPoliza = True

Set rs = oCon.CargaRecordSet("SELECT nEstado FROM GarantPoliza WHERE cNumPoliza='" & pcNumPoliza & "' AND cNumGarant='" & pcNumGarant & "' AND dTasacion='" & Format(pdTasacion, "mm/dd/yyyy") & "'")

If Not rs.EOF Then
    If rs!nEstado = 1 Then
        Adiciona_GarantPoliza = False
    Else
        sSql = "UPDATE GarantPoliza SET nEstado=1 WHERE cNumPoliza='" & pcNumPoliza & "' AND cNumGarant='" & pcNumGarant & "' AND dTasacion='" & Format(pdTasacion, "mm/dd/yyyy") & "'"
        oCon.Ejecutar sSql
    End If
Else
    sSql = " INSERT INTO GarantPoliza" & _
       " VALUES('" & pcNumPoliza & "','" & pcNumGarant & "','" & Format(pdTasacion, "mm/dd/yyyy") & "','" & Format(pdEstado, "mm/dd/yyyy") & "',1)"

    oCon.Ejecutar sSql
End If
End Function


Public Function Poliza_para_Credito(ByVal pcCtaCod As String, _
                                    ByVal pnMontoPrestamo As Double) As Boolean
Dim rs As ADODB.Recordset
Dim sSql As String

Poliza_para_Credito = False

'sSql = "SELECT nSaldoK=SUM(dbo.ColocObtieneMontoEnDolares(P.nSaldo,SUBSTRING(P.cCtaCod,9,1))) " & _
     " + dbo.ColocObtieneMontoEnDolares(" & pnMontoPrestamo & ",SUBSTRING(P.cCtaCod,9,1))" & _
     " FROM Poliza P INNER JOIN GarantPoliza GP ON P.cNumPoliza = GP.cNumPoliza " & _
     " INNER JOIN GarantRealTasacion GRT ON GRT.cNumGarant= GP.cNumGarant AND GRT.dTasacion = GP.dTasacion " & _
     " INNER JOIN ColocGarantia CG ON CG.cNumGarant = GRT.cNumGarant " & _
     " INNER JOIN Producto P ON P.cCtaCod= CG.cCtaCod " & _
     " WHERE CG.cCtaCod= '" & pcCtaCod & "'"

sSql = "SELECT nSaldoK=SUM(dbo.ColocObtieneMontoEnDolares(P.nSaldo,SUBSTRING(P.cCtaCod,9,1))) " & _
     " + dbo.ColocObtieneMontoEnDolares(" & pnMontoPrestamo & ",SUBSTRING(P.cCtaCod,9,1))" & _
     " FROM Producto P " & _
     " WHERE P.cCtaCod= '" & pcCtaCod & "' GROUP BY P.cCtaCod"

Set rs = oCon.CargaRecordSet(sSql)
If rs!nSaldoK >= 15000 Then    '15000 dolares es el monto minimo
    Poliza_para_Credito = True
End If

End Function

Public Function Elimina_GarantPoliza(ByVal pcNumPoliza As String, _
                                    ByVal pcNumGarant As String, _
                                    ByVal pdTasacion As Date)

Dim sSql As String


sSql = " UPDATE GarantPoliza" & _
       " SET nEstado=0 WHERE cNumPoliza='" & pcNumPoliza & "' AND cNumGarant='" & pcNumGarant & "'"
oCon.Ejecutar sSql

End Function

Private Sub Class_Initialize()
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
End Sub

Private Sub Class_Terminate()
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'WIOR 20120601**********************************************************
Public Function GenerarTramaMicroseguroAF(ByVal psAgencia As String, ByVal pnMoneda As Integer, ByVal pdFechaIni As Date, ByVal pdFechaFin As Date, ByVal pdTipoM As String) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
    
On Error GoTo GenerarTramaMicroseguroAF

        sSql = "exec stp_sel_GenerarTramaMicroseguroAF '" & psAgencia & "'," & pnMoneda & ",'" & Format(pdFechaIni, "mm/dd/yyyy") & "','" & Format(pdFechaFin, "mm/dd/yyyy") & "'," & CInt(pdTipoM)
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set GenerarTramaMicroseguroAF = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        
        Exit Function
GenerarTramaMicroseguroAF:
        Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function GenerarTramaMicroseguroCA(ByVal psAgencia As String, ByVal pnMoneda As Integer, ByVal pdFechaIni As Date, ByVal pdFechaFin As Date, ByVal pdTipoM As String) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
    
On Error GoTo GenerarTramaMicroseguroAF

        sSql = "exec stp_sel_GenerarTramaMicroseguroCA '" & psAgencia & "'," & pnMoneda & ",'" & Format(pdFechaIni, "mm/dd/yyyy") & "','" & Format(pdFechaFin, "mm/dd/yyyy") & "'," & CInt(pdTipoM)
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set GenerarTramaMicroseguroCA = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        
        Exit Function
GenerarTramaMicroseguroAF:
        Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function GenerarTramaMultiriesgoAF(ByVal psAgencia As String, ByVal pnMoneda As Integer, ByVal pdFechaIni As Date, ByVal pdFechaFin As Date) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
    
On Error GoTo GenerarTramaMultiriesgoAF

        sSql = "exec stp_sel_GenerarTramaMultiriesgoAF '" & psAgencia & "'," & pnMoneda & ",'" & Format(pdFechaIni, "mm/dd/yyyy") & "','" & Format(pdFechaFin, "mm/dd/yyyy") & "'"
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set GenerarTramaMultiriesgoAF = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        
        Exit Function
GenerarTramaMultiriesgoAF:
        Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function GenerarTramaMultiriesgoCA(ByVal psAgencia As String, ByVal pnMoneda As Integer, ByVal pdFechaIni As Date, ByVal pdFechaFin As Date) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
    
On Error GoTo GenerarTramaMultiriesgoCA

        sSql = "exec stp_sel_GenerarTramaMultiriesgoCA '" & psAgencia & "'," & pnMoneda & ",'" & Format(pdFechaIni, "mm/dd/yyyy") & "','" & Format(pdFechaFin, "mm/dd/yyyy") & "'"
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set GenerarTramaMultiriesgoCA = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        
        Exit Function
GenerarTramaMultiriesgoCA:
        Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'WIOR FIN **************************************************************

'WIOR 20140603 *********************************************************
Public Function CargaDatosPrimaNetaPolizaRenova(ByVal pnSumaAseg As Double, ByVal pnInmueble As Integer, ByVal psAgeCod As String) As ADODB.Recordset
Dim sSql As String

sSql = "exec stp_sel_CargaDatosPrimaNetaPolizaRenova " & pnSumaAseg & "," & pnInmueble & ",'" & psAgeCod & "'"

Set CargaDatosPrimaNetaPolizaRenova = oCon.CargaRecordSet(sSql)

End Function
Public Function MontoPolizaRenova(ByVal pnMoneda As Integer, ByVal pnPrima As Double, ByVal pnPrimaTC As Double) As ADODB.Recordset
Dim sSql As String

sSql = "exec stp_sel_MontoPolizaRenova " & pnMoneda & "," & pnPrima & "," & pnPrimaTC

Set MontoPolizaRenova = oCon.CargaRecordSet(sSql)

End Function
'WIOR FIN *************************************************************
'EJVG20150702 ***
Public Function DamePrimaPolizaxCuota(ByVal psCtaCod As String, ByVal pnNroCuota As Integer, ByVal pnUltimaCuota As Integer, ByVal pnPrdConceptoCod As ColocConcepto) As Currency
    Dim rs As ADODB.Recordset
    Dim sSql As String
    On Error GoTo ErrDamePrimaPolizaxCuota
    sSql = "EXEC stp_sel_ERS0632014_MontoPagarPolizaxCuota '" & psCtaCod & "'," & pnNroCuota & "," & pnUltimaCuota & "," & pnPrdConceptoCod
    Set rs = oCon.CargaRecordSet(sSql)
    If Not rs.EOF Then
        DamePrimaPolizaxCuota = rs!nMontoPagar
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
ErrDamePrimaPolizaxCuota:
    Err.Raise Err.Number, "Error En Proceso (DamePrimaPolizaxCuota)", Err.Description
End Function
Public Function RecuperaPolizasActivasxGarantia(ByVal psNumGarant As String) As ADODB.Recordset
    Dim sSql As String
    sSql = "EXEC stp_sel_ERS0632014_PolizasActivasxGarantia '" & psNumGarant & "' "
    Set RecuperaPolizasActivasxGarantia = oCon.CargaRecordSet(sSql)
End Function
'END EJVG *******

'INICIO EAAS20180405
Public Function RecuperaPolizasInternasCuenta(ByVal pcCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    sSql = "EXEC stp_sel_RecuperaPolizasInternasCuenta '" & pcCtaCod & "' "
    Set RecuperaPolizasInternasCuenta = oCon.CargaRecordSet(sSql)
End Function
'FIN EAAS20180405
'LUCV20180601, Según ERS022-2018
Public Function DamePrimaPolizaxCuotaTasaSegInc(ByVal psCtaCod As String, ByVal pnNroCuota As Integer, ByVal pnUltimaCuota As Integer, ByVal pnPrdConceptoCod As ColocConcepto) As ADODB.Recordset
    'Referencia: DamePrimaPolizaxCuota
    Dim rs As ADODB.Recordset
    Dim sSql As String
On Error GoTo ErrorDamePrimaPolizaxCuotaTasaSegInc
    sSql = "EXEC stp_sel_ERS0632014_MontoPagarPolizaxCuota '" & psCtaCod & "'," & pnNroCuota & "," & pnUltimaCuota & "," & pnPrdConceptoCod
    Set rs = oCon.CargaRecordSet(sSql)
    
    Set DamePrimaPolizaxCuotaTasaSegInc = oCon.CargaRecordSet(sSql)

    rs.Close
    Set rs = Nothing
    Exit Function
ErrorDamePrimaPolizaxCuotaTasaSegInc:
    Err.Raise Err.Number, "Error En Proceso (DamePrimaPolizaxCuotaTasaSegInc)", Err.Description
End Function
'Fin LUCV20180601
