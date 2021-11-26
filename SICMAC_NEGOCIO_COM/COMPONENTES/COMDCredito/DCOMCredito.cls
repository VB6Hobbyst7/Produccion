VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************************
'***     Rutina:           DCredito
'***     Descripcion:      Clase que permite Recuperar Datos del Credito
'***     Creado por:        NSSE
'***     Maquina:           07SISTDES03
'***     Fecha-Tiempo:         12/07/2001 12:06:08 AM
'***     Ultima Modificacion: Inicio de descripicion
'*****************************************************************************************

Option Explicit

Private gConsPersona As String
Private gConsComunes As String
Private gConsImagenes As String
Private Conn As COMConecta.DCOMConecta
Dim loConecta As COMConecta.DCOMConecta 'JACA 20111228
Dim objProducto As COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018

Public Function RecuperaBancosCtas(ByVal psCodPers As String) As ADODB.Recordset
Dim sSql As String
Dim C As COMConecta.DCOMConecta

    sSql = " select CI.cPersCod, CI.cCtaIFDesc "
    sSql = sSql & " from Ctaif CI"
    sSql = sSql & " Inner Join InstitucionFinanc I ON CI.cPerscod = I.cPerscod"
    sSql = sSql & " Where CI.cIFTpo in ('01','02') and CI.cCtaIFEstado=1 and I.cIFTpo='01' and LEN(CI.cCtaIFCod)=7"
    sSql = sSql & " AND CI.cPersCod = '" & psCodPers & "'"

    
    Set C = New COMConecta.DCOMConecta
    C.AbreConexion
    Set RecuperaBancosCtas = C.CargaRecordSet(sSql)
    C.CierraConexion
    Set C = Nothing
    
End Function

Public Function RecuperaBancos() As ADODB.Recordset
Dim sSql As String
Dim C As COMConecta.DCOMConecta

    sSql = " Select P.cPersNombre, P.cPersCod "
    sSql = sSql & " from InstitucionFinanc I"
    sSql = sSql & " Inner Join Persona P ON P.cPerscod = I.cPersCod"
    sSql = sSql & " Where cIFTpo = '01'"
    
    Set C = New COMConecta.DCOMConecta
    C.AbreConexion
    Set RecuperaBancos = C.CargaRecordSet(sSql)
    C.CierraConexion
    Set C = Nothing
    
End Function

Public Function UsuarioPerteneceACargo(ByVal psCodCargo As String, ByVal psCodusu As String, _
                                   ByVal psCodAge As String) As Boolean
Dim sSql As String
Dim C As COMConecta.DCOMConecta

    sSql = "Select cCodCargo "
    sSql = sSql & " From CredCargosUsu"
    sSql = sSql & " Where ccodUsu = '" & psCodusu & "' and cCodCargo = '" & psCodCargo & "'"
    '08-05-2005
    sSql = sSql & " AND (cAgeCod='" & psCodAge & "' OR ISNULL(cAgeCod,'')='')"
    Set C = New COMConecta.DCOMConecta
    C.AbreConexion
    If C.CargaRecordSet(sSql).RecordCount > 0 Then
        UsuarioPerteneceACargo = True
    Else
        UsuarioPerteneceACargo = False
    End If
    C.CierraConexion
    Set C = Nothing
    
    
End Function

Public Sub NuevaAutorizacion(ByVal pcCtaCod As String, ByVal pcCodCargo As String, ByVal pcCodUsu As String, _
    ByVal pcNomUsu As String, ByVal pdFecApr As Date, ByVal psComen As String)

Dim B As DCOMCredActBD
    
    Set B = New DCOMCredActBD
    Call B.dInsertaCredAprobacion(pcCtaCod, pcCodCargo, pcCodUsu, pcNomUsu, pdFecApr, psComen)
    Set B = Nothing
    
End Sub

Public Function CargaDatosAutorizacionCredito(ByVal psCodCta As String, ByRef prsDatos As ADODB.Recordset, ByRef pR As ADODB.Recordset) As String

Dim sSql As String
Dim C As COMConecta.DCOMConecta
'Dim R As ADODB.Recordset

sSql = " Select P.cPersNombre as cAnalista, P2.cPersNombre as cCliente, CN.cConsDescripcion cProd,"
sSql = sSql & " CN2.cConsDescripcion cMoneda, A.cAgeDescripcion, CE.nCuotas, CE.nPlazo, CE.nMonto, CE.nPeriodoFechaFija,cCodAnalista=P.cPersCod "
'08-05-2006
sSql = sSql & " ,Gracia=ISNULL(CE.nPeriodoGracia,0),Linea=L.cDescripcion,Pro.nTasaInteres, MontoCuota=SUM(CD.nMonto) "
'sSQL = sSQL & " from Colocaciones C"
sSql = sSql & " from Producto Pro INNER JOIN Colocaciones C ON Pro.cCtaCod=C.cCtaCod"
sSql = sSql & " INNER JOIN ColocacCred CC ON CC.cCtaCod = C.cCtaCod"
sSql = sSql & " INNER JOIN ColocLineaCredito L ON C.cLineaCred = L.cLineaCred"
'-------------
sSql = sSql & " Inner Join ProductoPersona PP ON PP.cCtaCod = C.cCtaCod and PP.nPrdPersRelac = 28"
sSql = sSql & " Inner Join Persona P ON P.cPersCod = PP.cPersCod"
sSql = sSql & " Inner Join ProductoPersona PP2 ON PP2.cCtaCod = C.cCtaCod and PP2.nPrdPersRelac = 20"
sSql = sSql & " Inner Join Persona P2 ON P2.cPersCod = PP2.cPersCod"
'sSql = sSql & " Inner Join Constante CN ON CN.nConsValor = SUBSTRING(C.cCtaCod, 6,3) and CN.nConsCod = 1001"
sSql = sSql & " Inner Join Constante CN ON CN.nConsValor = C.cTpoCredCod and CN.nConsCod = 3034"

sSql = sSql & " Inner Join Constante CN2 ON CN2.nConsValor = SUBSTRING(C.cCtaCod, 9,1) and CN2.nConsCod = 1011"
'sSql = sSql & " Inner Join Agencias A ON A.cAgeCod = SUBSTRING(C.cCtaCod, 4,2)"
sSql = sSql & " Inner Join Agencias A ON A.cAgeCod = C.cAgeCodAct"
sSql = sSql & " Inner Join ColocacEstado CE ON CE.cCtaCod = C.cCtaCod AND CE.nPrdEstado = " & gColocEstAprob
'08-05-2006
sSql = sSql & " Inner Join ColocCalendario CL ON CL.nNroCalen =CC.nNroCalen AND CE.cCtaCod = CL.cCtaCod"
sSql = sSql & " AND CL.nColocCalendApl=1 AND CL.nCuota=1 "
sSql = sSql & " INNER JOIN ColocCalendDet CD ON CL.cCtaCod=CD.cCtaCod AND CL.nNroCalen=CD.nNroCalen "
sSql = sSql & " AND CD.nColocCalendApl=1 AND CD.nCuota=1 "
sSql = sSql & " Where C.cCtaCod = '" & psCodCta & "'"
sSql = sSql & " GROUP BY P.cPersNombre,P2.cPersNombre,CN.cConsDescripcion,CN2.cConsDescripcion,A.cAgeDescripcion,"
sSql = sSql & " CE.nCuotas, CE.nPlazo, CE.nMonto, CE.nPeriodoFechaFija,P.cPersCod,CE.nPeriodoGracia,L.cDescripcion,Pro.nTasaInteres"
'---------------
'sSQL = sSQL & " Where C.cCtaCod = '" & psCodCta & "'"

Set C = New COMConecta.DCOMConecta
C.AbreConexion
Set prsDatos = C.CargaRecordSet(sSql)

If prsDatos.RecordCount > 0 Then
     
    sSql = "Select CA.cCtaCod, CA.cCodCargo, C.cNomCargo, CA.cCodUsu, CA.bComen, CA.cComen"
    '06-05-2006
    sSql = sSql & ",cAgeCod=ISNULL(RTRIM(cAgeCod),'') "
    '*************
    sSql = sSql & "   from CredAprobacion CA"
    'ALPA 20100616 *********
    sSql = sSql & " Inner Join Colocacaciones Col ON CA.cCtaCod = Col.cCtaCod"
    '***********************
    sSql = sSql & " Inner Join CredCargos C ON C.cCodCargo = CA.cCodCargo"
    '06-05-2006
    sSql = sSql & " LEFT JOIN CredCargosUsu CU ON CU.cCodCargo=C.cCodCargo "
    '*************
    sSql = sSql & " Where cCtaCod = '" & psCodCta & "'"
    'sSql = sSql & " AND ((cAgeCod='" & Mid(psCodCta, 4, 2) & "'AND C.bTodaAgencia=0) OR (ISNULL(cAgeCod,'')=''AND C.bTodaAgencia=1))"
    sSql = sSql & " AND ((cAgeCod = Col.cAgeCodAct AND C.bTodaAgencia=0) OR (ISNULL(cAgeCod,'')=''AND C.bTodaAgencia=1))"
    
    Set pR = C.CargaRecordSet(sSql)
    
    CargaDatosAutorizacionCredito = "OK"
Else
    CargaDatosAutorizacionCredito = ""
    
End If

'R.Close
'Set R = Nothing
C.CierraConexion
Set C = Nothing

End Function

Public Sub EliminarNivel(ByVal pnItem As Integer, ByVal pcCodRango As String, ByVal pcCodCargo As String)
Dim B As DCOMCredActBD

    Set B = New DCOMCredActBD
    Call B.dEliminarNivel(pnItem, pcCodRango, pcCodCargo)
    Set B = Nothing
    
    
End Sub

Public Sub EliminarCargo(ByVal psCodCargo As String)
Dim B As DCOMCredActBD

    Set B = New DCOMCredActBD
    Call B.dEliminaCargo(psCodCargo)
    Set B = Nothing
End Sub

Public Sub EliminarRango(ByVal psCodRango As String)
Dim B As DCOMCredActBD

    Set B = New DCOMCredActBD
    Call B.dEliminaRango(psCodRango)
    Set B = Nothing
End Sub


Public Sub EliminarCargoResponsable(ByVal psCodCargo As String, ByVal psCodusu As String)
Dim B As DCOMCredActBD

    Set B = New DCOMCredActBD
    Call B.dEliminaCargoResponsable(psCodCargo, psCodusu)
    Set B = Nothing
End Sub

Public Sub CargaCombosCredNivAprob(ByRef pProd As Variant, ByRef pTipoCred As Variant, ByRef pMoneda As Variant, ByRef pCargos As Variant)
Dim C As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset
Dim MatProd() As String
Dim MatTipoCred() As String
Dim MatTipoMon() As String
Dim MatCargos() As String
Dim Conex As COMConecta.DCOMConecta
Dim i As Integer
Dim sSql As String


    Set C = New COMDConstantes.DCOMConstantes
    Set R = C.RecuperaConstantes_2(1001, 1)
        
    ReDim MatProd(R.RecordCount)
    For i = 0 To R.RecordCount - 1
        MatProd(i) = Trim(R!cConsDescripcion) & Space(100) & Trim(R!nConsvalor)
        R.MoveNext
    Next i
        
    Set R = C.RecuperaConstantes_2(1050)
    ReDim MatTipoCred(R.RecordCount)
    For i = 0 To R.RecordCount - 1
        MatTipoCred(i) = Trim(R!cConsDescripcion) & Space(100) & Trim(R!nConsvalor)
        R.MoveNext
    Next i
    
    Set R = C.RecuperaConstantes_2(1011)
    ReDim MatTipoMon(R.RecordCount)
    For i = 0 To R.RecordCount - 1
        MatTipoMon(i) = Trim(R!cConsDescripcion) & Space(100) & Trim(R!nConsvalor)
        R.MoveNext
    Next i
    
    Set Conex = New COMConecta.DCOMConecta
    sSql = "Select cCodCargo, cNomCargo From CredCargos order by cNomCargo"
    Conex.AbreConexion
    Set R = Conex.CargaRecordSet(sSql)
    ReDim MatCargos(R.RecordCount)
    Do While Not R.EOF
        MatCargos(R.Bookmark - 1) = R!cNomCargo & Space(100) & R!cCodCargo
        R.MoveNext
    Loop
    R.Close
    Conex.CierraConexion
    Set Conex = Nothing
    
    
    Set R = Nothing
    Set C = Nothing
    
    pProd = MatProd
    pTipoCred = MatTipoCred
    pMoneda = MatTipoMon
    pCargos = MatCargos
    
End Sub

Public Sub InsertaNiveles(ByVal psCodRango As String, ByVal psCodCargo As String, _
        ByVal pbComen As Boolean)
        
Dim B As DCOMCredActBD
Dim C As COMConecta.DCOMConecta
Dim R As ADODB.Recordset
Dim nItem As Integer
Dim sSql As String

    sSql = "Select MAX(nItem) as nItem From CredRangoCargo Where cCodRango ='" & psCodRango & "'"
    Set C = New COMConecta.DCOMConecta
    C.AbreConexion
    Set R = C.CargaRecordSet(sSql)
    C.CierraConexion
    Set C = Nothing
    
    If Not IsNull(R!nItem) Then
        nItem = R!nItem + 1
    Else
        nItem = 1
    End If
    
    Set B = New DCOMCredActBD
    Call B.dInsertaNiveles(nItem, psCodRango, psCodCargo, pbComen)
    Set B = Nothing
    
    Set R = Nothing
    
End Sub


Public Sub IngresaRangos(ByVal sTipoProd As String, ByVal sTipoCred As String, ByVal sTipoMon As String, ByVal sMontoMin As String, ByVal sMontoMax As String)
Dim B As DCOMCredActBD
Dim C As COMConecta.DCOMConecta
Dim R As ADODB.Recordset
Dim sCodRango As String
Dim sSql As String

    sSql = "Select MAX(cCodRango) as cCodRangoMax From CredRangos"
    Set C = New COMConecta.DCOMConecta
    C.AbreConexion
    Set R = C.CargaRecordSet(sSql)
    C.CierraConexion
    Set C = Nothing
    
    sCodRango = Trim(Str(CInt(Trim(IIf(IsNull(R!cCodRangoMax), 0, R!cCodRangoMax))) + 1))
    sCodRango = Right("0000" & sCodRango, 4)
    Set B = New DCOMCredActBD
    Call B.dInsertaCredRangos(sCodRango, sTipoProd, CInt(sTipoMon), CInt(sTipoCred), CDbl(sMontoMax), CDbl(sMontoMin))
    Set B = Nothing
    
End Sub

Public Sub IngresaCreCargoResponsable(ByVal psCodCargo As String, ByVal psCodusu As String, ByVal psNombreUsu As String, ByVal pcAgeCod As String)
Dim B As DCOMCredActBD
Dim sSql As String

    Set B = New DCOMCredActBD
    Call B.dInsertaCredCargosUsu(psCodCargo, psCodusu, psNombreUsu, pcAgeCod)
    Set B = Nothing
    
End Sub


Public Sub IngresaCreCargo(ByVal psNomCargo As String, _
                            ByVal pbTodaAgencia As Integer)
Dim B As DCOMCredActBD
Dim C As COMConecta.DCOMConecta
Dim R As ADODB.Recordset
Dim sSql As String
Dim sCodCar As String

    sSql = "Select MAX(cCodCargo) as sCargoNew from credcargos"
    Set C = New COMConecta.DCOMConecta
    C.AbreConexion
    Set R = C.CargaRecordSet(sSql)
    C.CierraConexion
    Set C = Nothing
    
    If R.RecordCount > 0 Then
        sCodCar = Right("000" & Trim(Str(CInt(IIf(IsNull(R!sCargoNew), 0, R!sCargoNew)) + 1)), 3)
    Else
        sCodCar = "001"
    End If
    
    Set B = New DCOMCredActBD
    Call B.dInsertaCredCargos(sCodCar, psNomCargo, pbTodaAgencia)
    Set B = Nothing
End Sub

'*** PEAC 20080412
Public Function RecuperaClientesRelacionados(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
         
    sSql = "exec stp_sel_ObtieneClientesRelacionados '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaClientesRelacionados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function



'*** PEAC 20080412
Public Function RecuperaDatosClientesRelacionados(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
'    sSql = "Select Pers.cPersCod, Pers.cPersNombre, PID.cPersIDNro DNI,PID2.cPersIDNro RUC, "
'    sSql = sSql & " Pers.cPersDireccDomicilio, PNat.nPersnatHijos, Pers.dPersNacCreac, Pers.nperspersoneria, PNat.nPersNatEstCiv, "
'    sSql = sSql & " CN.cConsDescripcion cEstadoCivil, CN2.cConsDescripcion cDestino, "
'    sSql = sSql & " CN3.cConsDescripcion cCondicion, CN5.cConsDescripcion cTipoCred, L.cDescripcion cLinea, "
'    sSql = sSql & " CN6.cConsDescripcion cPersoneria "
'    sSql = sSql & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtacod AND PP.nPrdPersRelac = " & gColRelPersTitular
'    sSql = sSql & "                 Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod"
'    sSql = sSql & "                 Left Join PersID PID ON PID.cPersCod = PP.cPersCod AND PID.cPersIDTpo = " & gPersIdDNI
'    sSql = sSql & "                 Left Join PersID PID2 ON PID2.cPersCod = PP.cPersCod AND PID2.cPersIDTpo = " & gPersIdRUC
'    sSql = sSql & "                 Left Join PersonaNat PNat ON PNat.cPersCod = PP.cPersCod"
'    sSql = sSql & "                 Left Join Constante CN ON CN.nConsValor = PNat.nPersNatEstCiv AND CN.nConsCod = " & gPersEstadoCivil
'    sSql = sSql & "                 Inner Join Colocaciones C ON C.cCtaCod = P.cCtaCod "
'    sSql = sSql & "                 Left Join ColocLineaCredito L ON C.cLineaCred = L.cLineaCred "
'    sSql = sSql & "                 Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod "
'    sSql = sSql & "                 Left Join Constante CN2 ON CN2.nConsValor = CC.nColocDestino AND CN2.nConsCod = " & gColocDestino
'    sSql = sSql & "                 Left Join Constante CN3 ON CN3.nConsValor = CC.nColocCondicion AND CN3.nConsCod = " & gColocCredCondicion
'    sSql = sSql & "                 Left Join Constante CN5 ON Convert(int, SubString(P.cCtaCod,6,3)) = CN5.nConsValor AND CN5.nConsCod = " & gProducto
'    sSql = sSql & "                 Left Join Constante CN6 ON CN6.nConsValor = Pers.nPersPersoneria AND CN6.nConsCod = " & gPersPersoneria
'    sSql = sSql & " Where P.cCtaCod = '" & psCtaCod & "'"
    
    sSql = "exec stp_sel_ObtieneDatosClientesRelacionados '" & psCtaCod & "'"
        
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosClientesRelacionados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function




Public Function RecuperaCredCargos() As ADODB.Recordset
Dim C As COMConecta.DCOMConecta
Dim sSql As String

    Set C = New COMConecta.DCOMConecta
    sSql = "Select cCodCargo, cNomCargo,bTodaAgencia= CASE WHEN ISNULL(bTodaAgencia,0)=1 THEN 'SI' ELSE 'NO' END  from credcargos Order by cNomCargo "
    C.AbreConexion
    Set RecuperaCredCargos = C.CargaRecordSet(sSql)
    C.CierraConexion
    Set C = Nothing

End Function
Public Function RecuperaCredNiveles(ByVal psCodRango As String) As ADODB.Recordset
Dim sSql As String
Dim C As COMConecta.DCOMConecta

    sSql = "Select N.nItem, N.cCodRango, N.cCodCargo, CC.cNomCargo, CASE WHEN N.bComentario = 1 THEN 'SI' ELSE 'NO' END as cComen "
    sSql = sSql & " From CredRangoCargo N "
    sSql = sSql & " Inner Join CredCargos CC ON CC.cCodCargo = N.cCodCargo "
    sSql = sSql & " Where N.cCodRango = '" & psCodRango & "'"

    Set C = New COMConecta.DCOMConecta
    C.AbreConexion
    Set RecuperaCredNiveles = C.CargaRecordSet(sSql)
    C.CierraConexion
    Set C = Nothing
    
End Function


Public Function RecuperaCredRangos() As ADODB.Recordset
Dim C As COMConecta.DCOMConecta
Dim sSql As String

    Set C = New COMConecta.DCOMConecta
    C.AbreConexion
    
    sSql = "Select CR.cCodRango, C.cConsDescripcion,C2.cConsDescripcion as cMonedaDesc, CR.cCodRango, CR.cCodProd, CR.nMoneda, CR.nTipoCred, CR.nMontoMax, CR.nMontoMin, C3.cConsDescripcion as cTipoCredDesc"
    sSql = sSql & " From CredRangos CR"
    sSql = sSql & " Inner Join Constante C ON C.nConsValor = CR.cCodProd"
    sSql = sSql & " Inner Join Constante C2 ON C2.nConsValor = CR.nMoneda"
    sSql = sSql & " Inner Join Constante C3 ON C3.nConsValor = CR.nTipoCred and C3.nConsCod = 1050"
    sSql = sSql & " Where C.nConsCod = 1001 and C.nConsValor like '_0[123456789]'"
    sSql = sSql & " AND C2.nConsCod = 1011"
    Set RecuperaCredRangos = C.CargaRecordSet(sSql)
    
    C.CierraConexion
    Set C = Nothing

End Function

Public Function RecuperaCredCargosResponsable(ByVal psCodCargo As String) As ADODB.Recordset
Dim C As COMConecta.DCOMConecta
Dim sSql As String

    Set C = New COMConecta.DCOMConecta
    sSql = "Select cCodCargo, cCodUsu, cNomUsu,cAgeCod=ISNULL(cAgeCod,'') From CredCargosUsu where  cCodCargo ='" & psCodCargo & "' Order by cNomUsu"
    C.AbreConexion
    Set RecuperaCredCargosResponsable = C.CargaRecordSet(sSql)
    C.CierraConexion
    Set C = Nothing

End Function

Public Function DatosPosicionClienteComentarios(ByVal psPersCod As String) As ADODB.Recordset
Dim Conn As COMConecta.DCOMConecta
Dim sSql As String

On Error GoTo ErrorDatosPosicionClienteComentario

    Set Conn = New COMConecta.DCOMConecta
    sSql = "Select C.cPersCod, C.cMovNro, C.cComentario From Persona P JOIN PersComentario C ON " _
        & "P.cPersCod = C.cPersCod WHERE P.cPersCod = '" & psPersCod & "' ORDER BY C.cMovNro DESC"
        
    Conn.AbreConexion
    Set DatosPosicionClienteComentarios = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function

ErrorDatosPosicionClienteComentario:
    Err.Raise Err.Number, "Credito Posicion Cliente", Err.Description
End Function

Public Sub Grabar_Comentarios(pcPersCod As String, pcComentario As String, pdFecSis As Date, psCodAge As String, psCodUser As String)

    Dim oCont As COMNContabilidad.NCOMContFunciones
    Dim sMovNro As String, sTemp As String
    Dim oImpre As COMFunciones.FCOMImpresion

    Set oCont = New COMNContabilidad.NCOMContFunciones
    sMovNro = oCont.GeneraMovNro(pdFecSis, psCodAge, psCodUser)
    Set oCont = Nothing
    
    sTemp = pcComentario
    Set oImpre = New COMFunciones.FCOMImpresion
    
    pcComentario = Left(oImpre.JustificaTextoCadena(sTemp, 140), 255)
    
    AgregaPersComentario pcPersCod, sMovNro, pcComentario
    
End Sub

Public Sub AgregaPersComentario(ByVal sPersCod As String, ByVal sMovNro As String, _
        ByVal sComentario As String)

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorAgregaPersComentario
    sSql = "INSERT PersComentario (cPersCod,cMovNro,cComentario) " _
        & "Values ('" & sPersCod & "','" & sMovNro & "','" & sComentario & "')"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorAgregaPersComentario:
    Err.Raise Err.Number, "Agrega Persona Comentario", Err.Description
End Sub

Public Function RecuperaProductosDeCredito(Optional ByVal psCab As String = "") As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    'sSQL = " Select cConsDescripcion, nConsValor from Constante"
    'sSQL = sSQL & " where nConscod = 1001 and (nConsValor % 10) <> 0 and (nConsValor < 232 or nConsValor >= 300)"
    'sSQL = sSQL & " Order by cConsDescripcion DESC"

    sSql = " Select C.cConsDescripcion, C.nConsValor, CF.cAbrev from Constante C "
    sSql = sSql & " Inner Join ColocCredConsFiltro CF ON C.nConsCod = CF.nConsCod AND C.nConsValor = CF.nConsValor"
    sSql = sSql & " Where CF.nCodFiltro = " & gCredFiltroProd
    
    'ALPA BAII 20100603**********
    sSql = sSql & " AND C.nConsCod =3034"
    '****************************

    If psCab <> "" Then
        sSql = sSql & " AND C.nConsValor like '" & Mid(psCab, 1, 1) & "%'"
    End If
    sSql = sSql & " Order by C.nConsValor  "

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaProductosDeCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

'*** PEAC 20080515
Public Function RecuperaReferenciaPersComer(ByVal psPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaReferenciaPersComer
      
    sSql = " exec stp_sel_recuperaReferenciaPersComer '" & psPersCod & "'"
       
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaReferenciaPersComer = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaReferenciaPersComer:
    Err.Raise Err.Number, "Error En Proceso de Recuperar Referencias Personal/Comercial.", Err.Description
End Function

Public Function RecuperaDatosFteIng(ByVal psPersCod As String, ByVal psNumFuente As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaDatosFteIng
      
    sSql = "exec stp_sel_ObtieneDatosFteIng '" & psPersCod & "','" & psNumFuente & "'"
       
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosFteIng = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaDatosFteIng:
    Err.Raise Err.Number, "Error En Proceso de Recuperar Datos de fuente de ingreso.", Err.Description

End Function

Public Function RecuperaHisCred(ByVal psPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaHisCred
           
    sSql = "exec stp_sel_RecuperaHisCred '" & psPersCod & "'"
           
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaHisCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaHisCred:
    Err.Raise Err.Number, "Error En Proceso de Recuperar historia crediticia.", Err.Description
End Function

Public Function RecuperaEvolucionEconomica(ByVal psPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaEvolucionEco
      
    sSql = "exec stp_sel_ReporteEvolucionEconomica '" & psPersCod & "'"
                 
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.CargaRecordSet (sSql)
    Set RecuperaEvolucionEconomica = oConecta.CargaRecordSet("select * from TablaRepo ")
    oConecta.CargaRecordSet (" Truncate Table TablaRepo ")
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaEvolucionEco:
    Err.Raise Err.Number, "Error En Proceso de Recuperar historia crediticia.", Err.Description
End Function

Public Function RecuperaDatosFinan(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorRecuperaDatosFinan
      
    sSql = "exec stp_sel_ObtieneDatosFinancieros '" & psCtaCod & "'"
       
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosFinan = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaDatosFinan:
    Err.Raise Err.Number, "Error En Proceso.", Err.Description
End Function


Public Function RecuperaRelacPers(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorRecuperaRelacPers

'    sSql = "Select P.dPersNacCreac, PN.cPersNatSexo, P.nPersPersoneria, Prd.nPrdEstado, PP.cPersCod, RTRIM(P.cPersNombre) cPersNombre, C.cConsDescripcion, C.nConsValor, DNI = (Select Top 1 cPersIDnro from PersID Where cPersCod = PP.cPersCod  AND cPersIDTpo = '" & COMDConstantes.gPersIdDNI & "'),"
'    sSql = sSql & " RUC = (Select Top 1 cPersIDnro from PersID Where cPersCod = PP.cPersCod AND cPersIDTpo = '" & COMDConstantes.gPersIdRUC & "'), "
'    sSql = sSql & " CASE WHEN PP.nPrdPersRelac = 20 THEN 1 WHEN PP.nPrdPersRelac in (22,21) THEN 2 WHEN PP.nPrdPersRelac = 25 THEN 3 ELSE 99 END nOrden, cPersDireccDomicilio,cPersTelefono,cpersdireccubigeo "
'    sSql = sSql & " From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod "
'    sSql = sSql & "                         LEFT JOIN PersonaNat PN ON PN.cPersCod = P.cPersCod "
'    sSql = sSql & "                         INNER JOIN Producto Prd ON Prd.cCtaCod = PP.cCtaCod "
'    sSql = sSql & "                         INNER JOIN Constante C ON PP.nPrdPersRelac = C.nConsValor "
'    sSql = sSql & " WHERE C.nConsCod = " & COMDConstantes.gColocRelacPers & " And C.nConsValor not in (" & COMDConstantes.gColRelPersAnalista & "," & COMDConstantes.gColRelPersApoderado & "," & COMDConstantes.gColRelPersAcreedor & ") AND PP.cCtaCod = '" & psCtaCod & "' "
'    sSql = sSql & " ORDER BY nOrden "

    sSql = "exec stp_sel_RecuperaRelacPers '" & psCtaCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaRelacPers = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaRelacPers:
    Err.Raise Err.Number, "Error En Proceso de Recuperar Relaciones de Personas con el Credito", Err.Description
End Function

'*** PEAC 20080412
Public Function ObtieneCodEvaluacion(Optional nTipoEval As Integer = 2) As ADODB.Recordset

    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrHandler

    sSql = "exec stp_sel_MaquetaHojaEvaluacion '" & nTipoEval & "'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ObtieneCodEvaluacion = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
    
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set ObtieneCodEvaluacion = Null
End Function


'madm 20101012 - PARAMETRO BUSQUEDA CODAGE - MADM 20101108
Public Function GetPersDevConv(ByVal psPersCod As String, ByVal pnMoneda As COMDConstantes.Moneda, Optional psCodAge As String = "") As ADODB.Recordset
Dim SQL As String
Dim SQLAux As String
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
'ALPA 20091111********************************************************************
'SQL = "select   C.cPersCod, C.cPersCodIns, p.cPersNombre," _
'    & "         CONVERT(CHAR(10), c.dRegistro,103) AS dRegistro, c.cNroDoc, c.nMonto,C.nMovNro as nMovNroReg " _
'    & " from    ColocacConvenioRegDev C " _
'    & "         JOIN PERSONA P ON C.cPersCodIns = P.cPersCod " _
'    & "         JOIN CONSTANTE CC ON CC.nConsValor = C.nMoneda and CC.NCONSCOD =1011 " _
'    & " WHERE   (C.nPendiente IS NULL OR C.nPendiente<>'1')  and C.cPersCod ='" & psPersCod & "' and C.nMoneda=" & pnMoneda & " " _
'    & " order by C.dRegistro"

SQL = "exec stp_sel_ObtenerDevolucionConvenio '" & psPersCod & "'," & pnMoneda & ", '" & psCodAge & "'"

oCon.AbreConexion
'**********************************************************************************
Set GetPersDevConv = oCon.CargaRecordSet(SQL)
oCon.CierraConexion
Set oCon = Nothing
End Function

'MADM 20111001
Public Function GetPersDevConvFecha(ByVal pdFecIni As Date, ByVal pdFecFin As Date, ByVal pnMoneda As COMDConstantes.Moneda, Optional psCodAge As String = "") As ADODB.Recordset
Dim SQL As String
Dim SQLAux As String
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
'ALPA 20091111********************************************************************
'SQL = "select   C.cPersCod, C.cPersCodIns, p.cPersNombre," _
'    & "         CONVERT(CHAR(10), c.dRegistro,103) AS dRegistro, c.cNroDoc, c.nMonto,C.nMovNro as nMovNroReg " _
'    & " from    ColocacConvenioRegDev C " _
'    & "         JOIN PERSONA P ON C.cPersCodIns = P.cPersCod " _
'    & "         JOIN CONSTANTE CC ON CC.nConsValor = C.nMoneda and CC.NCONSCOD =1011 " _
'    & " WHERE   (C.nPendiente IS NULL OR C.nPendiente<>'1')  and C.cPersCod ='" & psPersCod & "' and C.nMoneda=" & pnMoneda & " " _
'    & " order by C.dRegistro"

SQL = "exec stp_sel_ObtenerCastigoDevolucionConvenio '" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "'," & pnMoneda & ", '" & psCodAge & "'"

oCon.AbreConexion
'**********************************************************************************
Set GetPersDevConvFecha = oCon.CargaRecordSet(SQL)
oCon.CierraConexion
Set oCon = Nothing
End Function
'END MADM

''MADM 20101108
'Public Function PertenecePersDevConv(ByVal psPersCod As String, ByVal psPersCodInst As String) As Boolean
'Dim SQL As String
'Dim oCon As COMConecta.DCOMConecta
'Dim R As ADODB.Recordset
'
'Set oCon = New COMConecta.DCOMConecta
'SQL = "exec stp_sel_PertenecePersDevolucionConvenio '" & psPersCod & "','" & psPersCodInst & "'"
'oCon.AbreConexion
''**********************************************************************************
'   Set R = oCon.CargaRecordSet(SQL)
'    If R.BOF And R.EOF Then
'        PertenecePersDevConv = False
'    Else
'        PertenecePersDevConv = True
'    End If
'
'    R.Close
'    Set R = Nothing
'oCon.CierraConexion
'Set oCon = Nothing
'End Function

'MADM 20110224
Public Function GetCreditoPersInstitucion(ByVal psPersCod As String, ByVal psPersCodInst As String) As ADODB.Recordset
Dim SQL As String
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
SQL = "exec stp_sel_PertenecePersDevolucionConvenio '" & psPersCod & "','" & psPersCodInst & "'"

oCon.AbreConexion
Set GetCreditoPersInstitucion = oCon.CargaRecordSet(SQL)
oCon.CierraConexion
Set oCon = Nothing
End Function
'MADM 20110225
Public Sub InsertaColocacConvenioRegDevolucion(ByVal psPersCodIns As String, ByVal psPersCod As String, ByVal pdRegistro As Date, ByVal psCodAge As String, ByVal pnMoneda As COMDConstantes.Moneda, ByVal pnMonto As Currency, ByVal psNroDoc As String, Optional ByVal pnMovNroReg As Long = 0, Optional ByVal pcCtaCod As String = "", _
                                                Optional ByVal pnTpoDoc As Integer = 47, Optional ByVal psIFPersCod As String, Optional ByVal psIFTpo As String = "", Optional ByVal psIFCta As String = "")
Dim SQL As String
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

oCon.AbreConexion

'SQL = "exec stp_ins_DevolucionCreditoinstitucion '" & psPersCodIns & "','" & psPersCod & "','" & Format(pdRegistro, "mm/dd/yyyy") & "','" & psCodAge & "'," & pnMoneda & "," & pnMonto & ",'" & psNroDoc & "'," & pnMovNroReg & ", '" & pcCtaCod & "' "
SQL = "exec stp_ins_DevolucionCreditoinstitucion '" & psPersCodIns & "','" & psPersCod & "','" & Format(pdRegistro, "mm/dd/yyyy") & "','" & psCodAge & "'," & pnMoneda & "," & pnMonto & ",'" & psNroDoc & "'," & pnMovNroReg & ", '" & pcCtaCod & "'," & pnTpoDoc & ",'" & psIFPersCod & "','" & psIFTpo & "','" & psIFCta & "'" 'EJVG20140228

    'Sql = "INSERT INTO ColocacConvenioRegDev (cPersCodIns, cPersCod, dRegistro, cCodAge, nMoneda, nMonto, cNroDoc, nMovNroReg)"
    '    & " Values ('" & psPersCodIns & "','" & psPersCod & "','" & Format(pdRegistro, "mm/dd/yyyy") & "','"
    '    & psCodAge & "'," & pnMoneda & "," & pnMonto & ",'" & psNroDoc & "'," & pnMovNroReg & ")"

    oCon.ConexionActiva.Execute SQL
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Sub

oCon.CierraConexion
End Sub

Public Sub ActualizaColocacConvenioRegDevOpe(ByVal psPersCodIns As String, ByVal psPersCod As String, ByVal pdRegistro As Date, ByVal pnMovNro As Currency, Optional ByVal pnCastDev As Integer = 0, Optional ByVal pnMovNroReg As Long = 0)
Dim SQL As String
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

oCon.AbreConexion

    'MADM 20110930
    If pnCastDev = 1 Then
    SQL = "UPDATE ColocacConvenioRegDev SET nMovNro=" & pnMovNro & ", nPendiente = 1,bPagoCastigo = 1  " _
        & " WHERE nMovNroReg =" & pnMovNroReg & ""
    Else
    SQL = "UPDATE ColocacConvenioRegDev SET nMovNro=" & pnMovNro & ", nPendiente = 1  " _
        & " WHERE cPersCodIns ='" & psPersCodIns & "' and cPersCod='" & psPersCod & "' and dRegistro='" & Format(pdRegistro, "mm/dd/yyyy") & "'"
    End If
    ''MADM 20110930
    
    oCon.ConexionActiva.Execute SQL
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Sub

oCon.CierraConexion
End Sub

Public Function RecuperaGarantiasCreditoConsol(ByVal psPersCod As String, ByVal pdFecha As Date) As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String
Dim oGeneral As COMDConstSistema.DCOMGeneral
Dim nTipoCambioFijo As Double

    On Error GoTo ErrorRecuperaGarantiasCredito

    Set oGeneral = New COMDConstSistema.DCOMGeneral
    nTipoCambioFijo = oGeneral.EmiteTipoCambio(pdFecha, TCFijoMes)
    nTipoCambioFijo = CDbl(Format(nTipoCambioFijo, "#0.00"))
    Set oGeneral = Nothing

    sSql = "Select SUM( "
    sSql = sSql & "         CASE WHEN  G.nMoneda = 1 THEN G.nPorGravar "
    sSql = sSql & "              WHEN  G.nMoneda = 2 THEN G.nPorGravar * " & Format(nTipoCambioFijo, "#0.00") & " END "
    sSql = sSql & "   ) nPorGravar,  "
    sSql = sSql & "    SUM( "
    sSql = sSql & "         CASE WHEN  G.nMoneda = 1 THEN G.nRealizacion "
    sSql = sSql & "              WHEN  G.nMoneda = 2 THEN G.nRealizacion * " & Format(nTipoCambioFijo, "#0.00") & " END "
    sSql = sSql & "  ) nRealizacion,  "
    sSql = sSql & "    SUM( "
    sSql = sSql & "         CASE WHEN  G.nMoneda = 1 THEN G.nTasacion "
    sSql = sSql & "              WHEN  G.nMoneda = 2 THEN G.nTasacion * " & Format(nTipoCambioFijo, "#0.00") & " END "
    sSql = sSql & "  ) nTasacion  "
    sSql = sSql & " From Garantias G  "
    sSql = sSql & "                       Inner Join Constante CN ON G.nTpoGarantia = CN.nConsvalor And CN.nConsCod = " & Trim(Str(COMDConstantes.gPersGarantia))
    sSql = sSql & "                       Inner Join PersGarantia PG ON G.cNumGarant = PG.cNumGarant AND PG.nRelacion = " & Trim(Str(COMDConstantes.gPersRelGarantiaTitular))
    sSql = sSql & "                       Inner Join Persona P ON PG.cPersCod = P.cPersCod "
    sSql = sSql & " WHERE PG.cPersCod = '" & psPersCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaGarantiasCreditoConsol = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaGarantiasCredito:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function NumerosCredEnJudicial(ByVal psPersCod As String)
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim R As ADODB.Recordset
    sSql = " select ISNULL(COUNT(P.cCtaCod),0) as NumCredJud from Producto P"
    sSql = sSql & "     Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod"
    sSql = sSql & " where PP.nPrdPersRelac = 20 AND P.nPrdEstado in (2201,2202)"
    sSql = sSql & " AND PP.cPersCod = '" & psPersCod & "'"

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    If R.RecordCount > 0 Then
        NumerosCredEnJudicial = R!NumCredJud
    Else
        NumerosCredEnJudicial = 0
    End If
    R.Close


End Function

Public Function RecuperaUltimoMovimiento(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

     'CTI6-20210503-ERS032-2019
'    sSql = " Select MC.nMovNro, MC.cOpeCod, dbo.FechaMov(M.cMovNro) as dFecPago,"
'    sSql = sSql & " nNroCuota = (Select MAX(nNroCuota) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro),"
'    sSql = sSql & " nMontoPagado = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro),"
'    sSql = sSql & " nCapital = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro AND nPrdConceptoCod = 1000 ),"
'    sSql = sSql & " nInteres = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod in (1100,1102,1103,1104,1105)),"
'    sSql = sSql & " nMora = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro AND nPrdConceptoCod = 1101),"
'    sSql = sSql & " nGastos = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod like  '12%'),"
'    sSql = sSql & " MC.nDiasMora , MC.nSaldoCap"
'    sSql = sSql & " from MovCol MC Inner Join Mov M ON M.nMovNro = MC.nMovNro"
'    sSql = sSql & " where MC.cOpeCod like '100%' AND MC.cCtaCod = '" & psCtaCod & "' "
'    sSql = sSql & " AND (SUBSTRING(MC.cOpeCod,1,4) in ('1002','1003','1004','1005','1006','1007','1070'))"
'    sSql = sSql & " AND MC.nMovNro = (Select MAX(nMovNro) From MovCol Where cCtaCod = '" & psCtaCod & "' AND (SUBSTRING(cOpeCod,1,4) in ('1002','1003','1004','1005','1006','1007','1070')) ) "
    
    sSql = "exec stp_sel_RecuperaUltimoMovimiento '" & psCtaCod & "'"
    'Fin CTI6-20210503

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaUltimoMovimiento = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function RecuperaCreditosVigentesGrabados(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer


    sSql = "Select cCtaCod, cCtaCodRef From ColocCredCredVig Where cCtaCod = '" & psCtaCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditosVigentesGrabados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing


End Function


Public Function RecuperaChequeCreditos(ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
    ByVal pnMoneda As Integer) As ADODB.Recordset

Dim sSql As String
Dim oConec As COMConecta.DCOMConecta
Dim sFecha As String


'*** PEAC 20120223
'    sSql = "select  '' , '.' as OPT,P.cPersNombre, DR.cIFCta, DR.cNroDoc,  DR.nMonto, C3.cConsDescripcion cProducto, "
'    sSql = sSql & " C2.cConsDescripcion cEstado, DR.dValorizacion, DR.nEstado,DR.cPersCod "
'    sSql = sSql & " from DocRec DR"
'    sSql = sSql & " Inner Join Persona P ON P.cPersCod = DR.cPersCod"
'    sSql = sSql & " Inner Join Constante C2 ON C2.nConsValor = DR.nEstado AND C2.nConsCod = 1007"
'    'ALPA 20100616***********
'    'sSql = sSql & " Inner Join ColocCredConsFiltro CF ON CF.nConsValor = DR.nProducto AND CF.nConsCod = 1001"
'    sSql = sSql & " Inner Join ColocCredConsFiltro CF ON CF.nConsValor = DR.nProducto AND CF.nConsCod = 3033"
'    'sSql = sSql & " Inner Join Constante C3 ON C3.nConsCod = 1001 AND C3.nConsValor = CF.nConsValor"
'    '************************
'    sSql = sSql & " Inner Join Constante C3 ON C3.nConsCod = 3033 AND C3.nConsValor = CF.nConsValor"
'    sSql = sSql & " Inner Join DocRecEst  RE on RE.cNroDoc=DR.cNroDoc and RE.nTpoDoc=DR.nTpoDoc and RE.cPersCod=DR.cPersCod and RE.cIFTpo=DR.cIFTpo"
'    sSql = sSql & " Inner Join Mov M on  RE.cMovNro =M.cMovNro"
'    sSql = sSql & " Where  dbo.FechaMov(RE.cMovNro) >= '" & Format(pdFecIni, "dd/mm/yyyy") & "' AND dbo.FechaMov(RE.cMovNro) <= '" & Format(pdFecFin, "dd/mm/yyyy") & "' "
'    sSql = sSql & " AND DR.nMoneda = " & pnMoneda & " AND DR.nTpoDoc = 47 and nMovFlag=0 "
'    sSql = sSql & " Order by P.cPersNombre,DR.dValorizacion"

    sSql = "exec stp_sel_RecuperaChequeCreditos '" & Format(pdFecIni, "dd/mm/yyyy") & "','" & Format(pdFecFin, "dd/mm/yyyy") & "'," & pnMoneda
'*** FIN PEAC

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set RecuperaChequeCreditos = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing

End Function

Public Function RecuperaPagosRealizados(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

'    sSQL = " Select MC.nMovNro, MC.cOpeCod, dbo.FechaMov(M.cMovNro) as dFecPago,"
'    sSQL = sSQL & " nNroCuota = (Select MAX(nNroCuota) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro),"
'    sSQL = sSQL & " nMontoPagado = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro),"
'    sSQL = sSQL & " nCapital = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro AND nPrdConceptoCod = 1000 ),"
'    sSQL = sSQL & " nInteres = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod in (1100,1102,1103,1104,1105)),"
'    sSQL = sSQL & " nMora = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro AND nPrdConceptoCod = 1101),"
'    sSQL = sSQL & " nGastos = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod like  '12%'),"
'    sSQL = sSQL & " nITF = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod in  (20, 21)),  "
'    sSQL = sSQL & " MC.nDiasMora , MC.nSaldoCap"
'    sSQL = sSQL & " from MovCol MC Inner Join Mov M ON M.nMovNro = MC.nMovNro"
'    sSQL = sSQL & " where MC.cOpeCod like '100%' AND MC.cCtaCod = '" & psCtaCod & "' "
'    sSQL = sSQL & " AND (SUBSTRING(MC.cOpeCod,1,4) in ('1002','1003','1004','1005','1006','1007','1070'))"
'    sSQL = sSQL & " Order by MC.nMovNro "

 sSql = "Select MC.nMovNro, MC.cOpeCod, dbo.FechaMov(M.cMovNro) as dFecPago, "
 sSql = sSql & " nNroCuota = (Select MAX(nNroCuota) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro),"
 sSql = sSql & " nMontoPagado = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro and cOpeCod not like '107[123456789]%')+(Select isnull(Sum(nMovImporte),0) From movOpeVarias Where nMovNro= MC.nMovNro),"
 sSql = sSql & " nCapital = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro AND nPrdConceptoCod in(1000,3000) and cOpeCod not like '107[123456789]%'),"
 sSql = sSql & " nInteres = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod in (1100,1102,1103,1104,1105,3100)),"
 sSql = sSql & " nMora = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro AND nPrdConceptoCod in(1101,3101)),"
 sSql = sSql & " nGastos = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod like  '12%'),"
 sSql = sSql & " nITF = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod in  (20, 21)),"
 sSql = sSql & " MC.nDiasMora , MC.nSaldoCap from MovCol MC Inner Join Mov M ON M.nMovNro = MC.nMovNro"
 sSql = sSql & " where (MC.cOpeCod like '100%' OR MC.cOpeCod = '105001') AND MC.cCtaCod = '" & psCtaCod & "'  AND M.nMovFlag<>2 AND "
 sSql = sSql & "     (SUBSTRING(MC.cOpeCod,1,4) in ('1002','1003','1004','1005','1006','1007','1070','1050')) AND "
 sSql = sSql & " (Select MAX(nNroCuota) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro) IS NOT NULL"
 sSql = sSql & " Order by (Select MAX(nNroCuota) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro)"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaPagosRealizados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function UltimaCuotaCanceladaOriginal(ByVal psCtaCod As String) As Integer
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    sSql = "Select ISNULL(MAX(nCuotaOrig),0) as nUltCuota From ColocCalifMiViv where cCtaCod = '" & psCtaCod & "' AND nColocCalendEstado = " & gColocCalendEstadoPagado
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    UltimaCuotaCanceladaOriginal = R!nUltCuota
    R.Close
    Set R = Nothing
End Function

Public Function SaldoPactadoCredito(ByVal psCtaCod As String) As Double
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    sSql = "Select ISNULL(SUM(CD.nMonto),0) as nSaldo "
    sSql = sSql & " from ColocCalendDet CD"
    sSql = sSql & " Where CD.nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = CD.cCtaCod )"
    sSql = sSql & " AND CD.nPrdConceptoCod = 1000 AND CD.nColocCalendApl = 1"
    sSql = sSql & " AND CD.cCtaCod = '" & psCtaCod & "'"

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    SaldoPactadoCredito = R!nSaldo
    R.Close
End Function

Public Function CuotaAprobada(ByVal psCtaCod As String) As Double
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim R As ADODB.Recordset


    sSql = "Select dbo.CuotaAprobada('" & psCtaCod & "') as nCuotaAprobada "
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    CuotaAprobada = IIf(IsNull(R!nCuotaAprobada), 0, R!nCuotaAprobada)

End Function

Public Function ExisteMetaAnalista(ByVal psPersCod As String, ByVal pdFechaIni As Date, _
                                ByVal pdFechaFin As Date, ByVal pnMoneda As Integer, ByVal pnTipoMeta As Integer) As Boolean

Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    sSql = " Select * from ColocMetasAnalista "
    sSql = sSql & "  Where cPersCod = '" & psPersCod & "' AND nTipoMeta = " & pnTipoMeta
    sSql = sSql & "  AND dInicial <= '" & Format(pdFechaIni, "mm/dd/yyyy") & "' AND dFinal >= '" & Format(pdFechaIni, "mm/dd/yyyy") & "' "
    sSql = sSql & "  and nMoneda = " & pnMoneda

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    If R.RecordCount > 0 Then
        ExisteMetaAnalista = True
    Else
        ExisteMetaAnalista = False
    End If
    R.Close
End Function

Public Function CreditosCanceladoConDesembolso(ByVal psCtaCod As String, ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConec As COMConecta.DCOMConecta

    sSql = "Select cCtaCod, nMonto from MovCol where nMovnro = " & pnMovNro & " and cCtaCod <> '" & psCtaCod & "' and copecod like '100[23456]%'"
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set CreditosCanceladoConDesembolso = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function PerteneceADesembolsoConCancelacion(ByVal pnMovNro As Long, ByVal psCtaCod As String) As Boolean
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String
Dim R As ADODB.Recordset
    sSql = "Select cCtaCod,cOpeCod from MovCol where nMovnro = " & pnMovNro & " and cOpeCod in ('" & gCredDesembEfec & "','" & gCredDesembCtaNueva & "','" & gCredDesembCtaExist & "','" & gCredDesembCtaExistDOA & "','" & gCredDesembCtaNuevaDOA & "') and cCtaCod <> '" & psCtaCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    If R.RecordCount > 0 Then
        PerteneceADesembolsoConCancelacion = True
    Else
        PerteneceADesembolsoConCancelacion = False
    End If
    R.Close
    Set R = Nothing

End Function

Public Function RecuperaCuentasAho(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConec As COMConecta.DCOMConecta

    sSql = "Select cCodCtaAho from ColocCargoAutoma Where cCtaCod = '" & psCtaCod & "' "
    sSql = sSql & " AND cestado = '1' Order by nOrden "

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set RecuperaCuentasAho = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing

End Function

Public Function RecuperaCreditosCargoAutomatico(ByVal pdHoy As Date, ByVal pnMoneda As COMDConstantes.Moneda) As ADODB.Recordset
Dim sSql As String
Dim oConec As COMConecta.DCOMConecta


    sSql = "        Select  ' ', C.cCtaCod, P.cPersNombre,  "
    sSql = sSql & "         nMonto = (Select SUM(nMonto-nMontoPagado) From ColocCalendDet CD"
    sSql = sSql & "                   Where cCtaCod = CC.cCtaCod And nNroCalen = CC.nNroCalen And nColocCalendApl = CC.nColocCalendApl "
    sSql = sSql & "                         AND nCuota = CC.nCuota), "
    sSql = sSql & "         Prd.nPrdEstado, C.cMetLiquidacion, CG.cCodCtaAho  "
    sSql = sSql & " from    Producto Prd Inner Join ColocacCred C ON Prd.cCtaCod = C.cCtaCod "
    sSql = sSql & "         INNER Join ColocCalendario CC ON CC.cCtaCod = C.cCtaCod AND CC.nNroCalen = C.nNroCalen"
    sSql = sSql & "         INNER Join ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "         INNER Join Persona P ON PP.cPersCod = P.cPersCod"
    sSql = sSql & "         INNER Join ColocCargoAutoma CG ON CG.cCtaCod = C.cCtaCod "
    sSql = sSql & " Where   CC.nColocCalendEstado = " & gColocCalendEstadoPendiente & " AND CC.nColocCalendApl = " & gColocCalendAplCuota & " AND CC.dVenc = '" & Format(pdHoy, "mm/dd/yyyy") & "' "
    sSql = sSql & "         AND SUBSTRING(C.cCtaCod,9,1)='" & pnMoneda & "' "
    sSql = sSql & " ORDER BY P.cPersNombre "

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set RecuperaCreditosCargoAutomatico = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing

End Function

Public Function CodigosModulares(ByVal psPersCodTitular As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    sSql = "Select C.cCodModular From ColocacConvenio C Inner Join ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular & " AND PP.cPersCod = '" & psPersCodTitular & "' Order by C.cCodModular"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CodigosModulares = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function DefineCondicionCredito_ALPA(ByVal psPersCod As String, Optional ByVal pnProducto As Producto = 0, _
Optional ByVal pdFecha As Date, Optional ByVal pbRefinaciado As Boolean = False) As COMDConstantes.ColocCredCondicion
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConn As COMConecta.DCOMConecta
Dim nCredNumVig As Integer
Dim nCredNumCancelados As Integer
Dim cCalif As String
    'ALPA 20130123****************************************************************************
    pnProducto = IIf(Len(Trim(pnProducto)) = 3, CInt(Mid(CStr(pnProducto), 1, 2)), pnProducto)
    '*****************************************************************************************
    'Define la Condicion del Credito por Producto
    If pnProducto <> 0 Then
        sSql = "Select Count(PP.cctaCod) as nTotal from ProductoPersona PP Inner Join Producto P ON P.cCtaCod = PP.cCtaCod Inner Join Colocaciones C on C.cCtaCod = PP.cCtaCod Where PP.cPersCod = '" & psPersCod & "' "
        sSql = sSql & " AND PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
        sSql = sSql & " AND left(cTpoProdCod,len(" & Trim(Left(pnProducto, 2)) & ")) = '" & Trim(Str(Left(pnProducto, 2))) & "' "
        sSql = sSql & " AND P.nPrdEstado in (" & COMDConstantes.gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor & "," & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & ")"
        'MAVM 20110628 ***
        'sSql = sSql & " AND SubString(PP.cCtaCod,6,3)<>'305'"
        'ALPA 20130118 Comentado por RQ12214-2012
        'sSql = sSql & " AND (cTpoProdCod <> '703') AND (cTpoProdCod <> '705')"
        '***
        
        
    Else
        sSql = "Select Count(PP.cctaCod) as nTotal from ProductoPersona PP Inner Join Producto P ON P.cCtaCod = PP.cCtaCod Inner Join Colocaciones C on C.cCtaCod = PP.cCtaCod Where cPersCod = '" & psPersCod & "' "
        sSql = sSql & " AND PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
        sSql = sSql & " AND P.nPrdEstado in (" & COMDConstantes.gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor & "," & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & ")"
        'MAVM 20110628 ***
        'sSql = sSql & " AND SubString(PP.cCtaCod,6,3)<>'305'"
        'ALPA 20130118 Comentado por RQ12214-2012
        'sSql = sSql & " AND (cTpoProdCod <> '703') AND (cTpoProdCod <> '705') And (cTpoProdCod <> '514')"
        'sSql = sSql & " And (cTpoProdCod <> '514')" '**ARLO20180712 ERS042 - 2018
        sSql = sSql & " And (cTpoProdCod not in (select valor from fn_getProdEquivTbl('514')))" '**ARLO20180712 ERS042 - 2018
        '***
    End If
    Set oConn = New COMConecta.DCOMConecta
    oConn.AbreConexion
    Set R = oConn.CargaRecordSet(sSql)
    nCredNumVig = IIf(IsNull(R!nTotal), 0, R!nTotal)
    R.Close
    
    '*** PEAC 20100525
'    If pnProducto <> 0 Then
'        sSql = "Select Count(PP.cctaCod) as nTotal from ProductoPersona PP Inner Join Producto P ON P.cCtaCod = PP.cCtaCod "
'        sSql = sSql & " inner join colocaciones c on c.cctacod=pp.cctacod"
'        sSql = sSql & " Where cPersCod = '" & psPersCod & "' "
'        sSql = sSql & " AND PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
'        sSql = sSql & " AND SUBSTRING(P.cCtaCod,6,2) = '" & Trim(Str(pnProducto)) & "' "
'        sSql = sSql & " AND P.nPrdEstado not in (" & COMDConstantes.gColocEstSolic & "," & COMDConstantes.gColocEstSug & "," & COMDConstantes.gColocEstAprob & "," & gColocEstRetirado & "," & COMDConstantes.gColocEstRech & "," & COMDConstantes.gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor & ")"
'        sSql = sSql & " AND DATEDIFF(day,C.dVenc,'" & CStr(Format(pdFecha, "MM/dd/yyyy")) & "')<=360"
'        sSql = sSql & " AND SubString(PP.cCtaCod,6,3)<>'305'"
'    Else
'        sSql = "Select Count(PP.cctaCod) as nTotal from ProductoPersona PP Inner Join Producto P ON P.cCtaCod = PP.cCtaCod "
'        sSql = sSql & " Inner Join Colocaciones C on C.cCtaCod=PP.cCtaCod"
'        sSql = sSql & " Where cPersCod = '" & psPersCod & "' "
'        sSql = sSql & " AND PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
'        sSql = sSql & " AND P.nPrdEstado not in (" & COMDConstantes.gColocEstSolic & "," & COMDConstantes.gColocEstSug & "," & COMDConstantes.gColocEstAprob & "," & gColocEstRetirado & "," & COMDConstantes.gColocEstRech & "," & COMDConstantes.gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor & ")"
'        sSql = sSql & " AND DATEDIFF(day,C.dVenc,'" & CStr(Format(pdFecha, "MM/dd/yyyy")) & "')<=360"
'
'    End If
    
    sSql = " exec stp_sel_ObtieneNumCredCancelados " & pnProducto & ",'" & psPersCod & "','" & CStr(Format(pdFecha, "yyyymmdd")) & "' "
    '*** FIN PEAC
    
    Set R = oConn.CargaRecordSet(sSql)
    nCredNumCancelados = IIf(IsNull(R!nTotal), 0, R!nTotal)
    R.Close
    oConn.CierraConexion
    Set oConn = Nothing
    
    'si no tiene ningun credito es normal
    If nCredNumVig = 0 And nCredNumCancelados = 0 Then
        If pbRefinaciado = True Then
            DefineCondicionCredito_ALPA = COMDConstantes.gColocCredCondRefinan
            Exit Function
        End If
        
        'MAVM 20110613 ***
        sSql = " exec stp_sel_ObtieneCalificacion '" & psPersCod & "'"
        Set oConn = New COMConecta.DCOMConecta
        oConn.AbreConexion
        Set R = oConn.CargaRecordSet(sSql)
        If R.RecordCount <> 0 Then
            cCalif = IIf(IsNull(R!cCalif), "", R!cCalif)
        End If
        R.Close
        oConn.CierraConexion
        Set oConn = Nothing
        If cCalif = "D" Or cCalif = "E" Then
            DefineCondicionCredito_ALPA = COMDConstantes.gColocCredCondRecurrente
        Else
            DefineCondicionCredito_ALPA = COMDConstantes.gColocCredCondNormal
        End If
        '***
    Else
        'REFINANCIADO
        If pbRefinaciado = True Then
            DefineCondicionCredito_ALPA = COMDConstantes.gColocCredCondRefinan
            Exit Function
        End If
        'si tiene creditos vigentes
        If nCredNumVig > 0 And nCredNumCancelados = 0 Or nCredNumVig > 0 And nCredNumCancelados > 0 Then
            DefineCondicionCredito_ALPA = COMDConstantes.gColocCredCondParalelo
        Else
            DefineCondicionCredito_ALPA = COMDConstantes.gColocCredCondRecurrente
        End If
    End If

End Function
Public Function DefineCondicionCredito(ByVal psPersCod As String, Optional ByVal pnProducto As Producto = 0, _
Optional ByVal pdFecha As Date, Optional ByVal pbRefinaciado As Boolean = False, Optional ByVal pnNroCuotasNewCredito As Integer = 0) As COMDConstantes.ColocCredCondicion
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConn As COMConecta.DCOMConecta
Dim nCredNumVig As Integer
Dim nCredNumCancelados As Integer
Dim cCalif As String

    'Define la Condicion del Credito por Producto
    If pnProducto <> 0 Then
        sSql = "Select Count(PP.cctaCod) as nTotal from ProductoPersona PP Inner Join Producto P ON P.cCtaCod = PP.cCtaCod Inner Join Colocaciones C on C.cCtaCod = PP.cCtaCod Where PP.cPersCod = '" & psPersCod & "' "
        sSql = sSql & " AND PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
        sSql = sSql & " AND (cTpoProdCod) = '" & Trim(Str(pnProducto)) & "' "
        sSql = sSql & " AND P.nPrdEstado in (" & COMDConstantes.gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor & ")"
        'MAVM 20110628 ***
        'sSql = sSql & " AND SubString(PP.cCtaCod,6,3)<>'305'"
        'sSql = sSql & " AND (cTpoProdCod <> '703') AND (cTpoProdCod <> '705')" 'ARLO20180905
        sSql = sSql & " AND (cTpoProdCod not in (select valor from fn_getProdEquivTbl('703,705')))" 'ARLO20180905
        '***
    Else
        sSql = "Select Count(PP.cctaCod) as nTotal from ProductoPersona PP Inner Join Producto P ON P.cCtaCod = PP.cCtaCod Inner Join Colocaciones C on C.cCtaCod = PP.cCtaCod Where cPersCod = '" & psPersCod & "' "
        sSql = sSql & " AND PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
        sSql = sSql & " AND P.nPrdEstado in (" & COMDConstantes.gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor & ")"
        'MAVM 20110628 ***
        'sSql = sSql & " AND SubString(PP.cCtaCod,6,3)<>'305'"
        'sSql = sSql & " AND (cTpoProdCod <> '703') AND (cTpoProdCod <> '705') And (cTpoProdCod <> '514')" 'ARLO20180905
        sSql = sSql & " AND (cTpoProdCod not in (select valor from fn_getProdEquivTbl('703,705,514')))" 'ARLO20180905
        '***
    End If
    Set oConn = New COMConecta.DCOMConecta
    oConn.AbreConexion
    Set R = oConn.CargaRecordSet(sSql)
    nCredNumVig = IIf(IsNull(R!nTotal), 0, R!nTotal)
    R.Close
      
    sSql = " exec stp_sel_ObtieneNumCredCancelados " & pnProducto & ",'" & psPersCod & "','" & CStr(Format(pdFecha, "yyyymmdd")) & "' "
    
    Set R = oConn.CargaRecordSet(sSql)
    nCredNumCancelados = IIf(IsNull(R!nTotal), 0, R!nTotal)
    R.Close
    oConn.CierraConexion
    Set oConn = Nothing
    
    'si no tiene ningun credito es normal
    If nCredNumVig = 0 And nCredNumCancelados = 0 Then
        If pbRefinaciado = True Then
            DefineCondicionCredito = COMDConstantes.gColocCredCondRefinan
            Exit Function
        End If
        
        'MAVM 20110613 ***
        sSql = " exec stp_sel_ObtieneCalificacion '" & psPersCod & "'"
        Set oConn = New COMConecta.DCOMConecta
        oConn.AbreConexion
        Set R = oConn.CargaRecordSet(sSql)
        If R.RecordCount <> 0 Then
            cCalif = IIf(IsNull(R!cCalif), "", R!cCalif)
        End If
        R.Close
        oConn.CierraConexion
        Set oConn = Nothing
        If cCalif = "D" Or cCalif = "E" Then
            DefineCondicionCredito = COMDConstantes.gColocCredCondRecurrente
        Else
            DefineCondicionCredito = COMDConstantes.gColocCredCondNormal
        End If
        '***
    Else
        'REFINANCIADO
        If pbRefinaciado = True Then
            DefineCondicionCredito = COMDConstantes.gColocCredCondRefinan
            Exit Function
        End If
        'si tiene creditos vigentes
        If nCredNumVig > 0 And nCredNumCancelados = 0 Or nCredNumVig > 0 And nCredNumCancelados > 0 Then
            If pnNroCuotasNewCredito = 1 Then
            DefineCondicionCredito = COMDConstantes.gColocCredCondParalelo
            Else
                DefineCondicionCredito = COMDConstantes.gColocCredCondAdicional
            End If
        Else
            DefineCondicionCredito = COMDConstantes.gColocCredCondRecurrente
        End If
    End If
End Function
Public Function VerificaCombinacionProd(ByVal psPersCod As String, Optional ByVal pnProducto As COMDConstantes.Producto = 0, _
Optional ByVal pbSustitucion As Boolean = False) As String
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConn As COMConecta.DCOMConecta
Dim nCredNumProd As Integer
Dim nProductoCom As Integer
Dim sVerificaCombinacionProd As String
Dim nCheck As Boolean
Dim rs As ADODB.Recordset


    If pnProducto = 20 Then
       nProductoCom = 10
       sVerificaCombinacionProd = "El Titular presenta en su historia Créditos Comerciales"

       'Define la Condicion del Credito por Producto '
           sSql = "Select Count(PP.cctaCod) as nTotal from ProductoPersona PP Inner Join Producto P ON P.cCtaCod = PP.cCtaCod Where PP.cPersCod = '" & psPersCod & "' "
           sSql = sSql & " AND PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
           sSql = sSql & " AND SUBSTRING(P.cCtaCod,6,2) = '" & Trim(Str(nProductoCom)) & "' "
           sSql = sSql & " AND P.nPrdEstado in(2020,2021,2022,2030,2032,2031,2032)"

       Set oConn = New COMConecta.DCOMConecta
       oConn.AbreConexion
       Set R = oConn.CargaRecordSet(sSql)
       nCredNumProd = IIf(IsNull(R!nTotal), 0, R!nTotal)
       R.Close
       oConn.CierraConexion
       Set oConn = Nothing

     If pbSustitucion = False Then
            'si no tiene ningun credito c    on el producto Combinacion '
            If nCredNumProd = 0 Then
               VerificaCombinacionProd = ""
            Else 'si tiene creditos con el producto Combinacion '
               VerificaCombinacionProd = sVerificaCombinacionProd
            End If
      Else
          sSql = "Select nParamValor From ColocParametro Where nParamVar = 3102"
          Set oConn = New COMConecta.DCOMConecta
          oConn.AbreConexion
          Set rs = oConn.CargaRecordSet(sSql)
          oConn.CierraConexion
          Set oConn = Nothing

          If Not rs.EOF And Not rs.BOF Then
             If rs!nParamValor = 1 Then
                nCheck = True
             Else
                nCheck = False
             End If
          End If
          Set rs = Nothing
           If nCheck = True Then ' tonces no se verifica el creditos de comercial a ames
                VerificaCombinacionProd = ""
           Else
                If nCredNumProd = 0 Then
                    VerificaCombinacionProd = ""
                Else 'si tiene creditos con el producto Combinacion '
                    VerificaCombinacionProd = sVerificaCombinacionProd
                 End If
           End If
      End If
    End If


End Function


Public Function RecuperaNivelesAprUsuario(ByVal psCodUser As String, ByVal pnMonto As Double, _
    psMoneda As String, ByVal psProducto As String, ByVal pbRefinanc As Boolean) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta


    sSql = " Select CN.cCodNiv, CA.cProduct, CA.cTpoCred, CA.nNivel, CA.nMontoMax, CA.nMontoMin"
    sSql = sSql & " From ColocCredPersNivelesApr CN Inner Join RRHH RH ON CN.cPersCod = RH.cPersCod"
    sSql = sSql & " Inner join ColocCredNivelesApr CA ON CN.cCodNiv = CA.cCodNiv"
    sSql = sSql & " Where RH.cUser = '" & psCodUser & "' AND CA.cProduct = '" & psProducto & "' AND CA.cTpoCred = '" & IIf(pbRefinanc, Trim(Str(gColocCredTipoRefinanciado)), Trim(Str(gColocCredTipoNormal))) & "' "
    sSql = sSql & " AND CA.nMontoMax >= " & Format(pnMonto, "#0.00") & " AND CA.nMontoMin <= " & Format(pnMonto, "#0.00") & " AND CA.nMoneda = " & psMoneda

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaNivelesAprUsuario = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function RecuperaMovimientosAhorros(ByVal pnMov As Long, Optional ByVal pbSinMovRef As Boolean = False) As ADODB.Recordset
Dim sSql As String
Dim oConec As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    sSql = "Select * from Movref Where nMovNroRef = " & pnMov
    Set R = oConec.CargaRecordSet(sSql, adLockReadOnly)

    If R.RecordCount > 0 And Not pbSinMovRef Then
        sSql = "Select MC.nMovNro,MC.cOpeCod, MC.nMonto, MC.cCtaCod "
        'sSql = "Select MC.cOpeCod, SUM(MC.nMonto) nMonto "
        sSql = sSql & " From MovCap MC"
        sSql = sSql & " Inner Join MovRef MR ON MC.nMovNro = MR.nMovNro"
        sSql = sSql & " Where MR.nMovNroRef = " & pnMov
        sSql = sSql & " Order By MC.nMovNro "
        'sSql = sSql & " Group By MC.cOpeCod"
    Else
        sSql = "Select MC.nMovNro,MC.cOpeCod, MC.nMonto, MC.cCtaCod "
        'sSql = "Select MC.cOpeCod, SUM(MC.nMonto) nMonto "
        sSql = sSql & " From MovCap MC"
        sSql = sSql & " Where MC.nMovNro = " & pnMov
        sSql = sSql & " Order By MC.nMovNro "

    End If
    R.Close

    Set RecuperaMovimientosAhorros = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function RecuperaDatosExtorno(ByVal pnMovNro As Long, ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

'ARCV 06-09-2006
'--- Se modifico para manejar los Desembolsos
'--- con Abono a Cuenta y Cancelacion de Creditos

'JIPR20201124 Mejora Datos
    'ALPA20130723
'    sSql = "Select M.nCredEstado, SUM(MD.nMonto)+isnull(dbo.FncDevolverPagadoMiViviendaxnMovNro(MR.nMovNro),0) as nCapital, "
'    'sSql = "Select M.nCredEstado, SUM(MD.nMonto) as nCapital, "
'    sSql = sSql & " MIN(MD.nNroCuota) nMinCuota, M.nDiasMora, M.nFlag "
'    sSql = sSql & " ,nMovNroRef=ISNULL(MR.nMovNro,0)" 'ARCV 06-09
'    sSql = sSql & " From MovCol M Inner Join MovColDet MD ON M.nMovNro = MD.nMovNro "
'    sSql = sSql & " AND M.cOpeCod = MD.cOpeCod " 'ARCV 28-05-07
'    sSql = sSql & " AND MD.nPrdConceptoCod = 1000 "
'    sSql = sSql & "     Left Join MovRef MR ON M.nMovNro = MR.nMovNroRef " 'ARCV 06-09
'    'ALPA 20160324********
'    'sSql = sSql & " Where M.nMovNro = " & pnMovNro & " And MD.cCtaCod = '" & psCtaCod & "' AND SUBSTRING(CONVERT(varchar(4),MD.nPrdConceptoCod),1,2)<>'12' AND M.cCtaCod = '" & psCtaCod & "' and M.nCredEstado<>0  and MD.COPECOD LIKE '100[1234567]%' and M.cOpeCod not Like '107[123456789]%'"
'    sSql = sSql & " Where M.nMovNro = " & pnMovNro & " And MD.cCtaCod = '" & psCtaCod & "' AND SUBSTRING(CONVERT(varchar(4),MD.nPrdConceptoCod),1,2)<>'12' AND M.cCtaCod = '" & psCtaCod & "' and M.nCredEstado<>0  and (MD.COPECOD LIKE '100[1234567]%' or MD.COPECOD LIKE '1008[01][0123456]%') and M.cOpeCod not Like '107[123456789]%'"
'    '*********************
'    sSql = sSql & " Group By  M.nCredEstado,M.nDiasMora, M.nFlag "
'    sSql = sSql & " ,MR.nMovNro" 'ARCV 06-09
'    Set oConecta = New COMConecta.DCOMConecta
'    oConecta.AbreConexion
'    Set RecuperaDatosExtorno = oConecta.CargaRecordSet(sSql)

    sSql = "exec stp_sel_DatosExtorno '" & psCtaCod & "'," & pnMovNro & ""
    Set oConecta = New COMConecta.DCOMConecta
     oConecta.AbreConexion
    Set RecuperaDatosExtorno = oConecta.CargaRecordSet(sSql)
    

    If RecuperaDatosExtorno.RecordCount = 0 Then
        sSql = "Select M.nCredEstado, 0 as nCapital, "
        sSql = sSql & " MIN(MD.nNroCuota) nMinCuota, M.nDiasMora, M.nFlag "
        sSql = sSql & " ,nMovNroRef=ISNULL(MR.nMovNro,0)" 'ARCV 06-09
        sSql = sSql & " From MovCol M Inner Join MovColDet MD ON M.nMovNro = MD.nMovNro "
        sSql = sSql & " AND M.cOpeCod = MD.cOpeCod " 'ARCV 28-05-07
        sSql = sSql & " AND MD.nPrdConceptoCod = 1100 "
        sSql = sSql & "     Left Join MovRef MR ON M.nMovNro = MR.nMovNroRef " 'ARCV 06-09
        sSql = sSql & " Where M.nMovNro = " & pnMovNro & " And MD.cCtaCod = '" & psCtaCod & "' AND SUBSTRING(CONVERT(varchar(4),MD.nPrdConceptoCod),1,2)<>'12' AND M.cCtaCod = '" & psCtaCod & "' and M.cOpeCod not Like '107[123456789]%'"
        sSql = sSql & " Group By  M.nCredEstado,M.nDiasMora, M.nFlag "
        sSql = sSql & " ,MR.nMovNro" 'ARCV 06-09
        Set RecuperaDatosExtorno = oConecta.CargaRecordSet(sSql)
    End If

    If RecuperaDatosExtorno.RecordCount = 0 Then
        sSql = "Select M.nCredEstado, 0 as nCapital, "
        sSql = sSql & " MIN(MD.nNroCuota) nMinCuota, M.nDiasMora, M.nFlag "
        sSql = sSql & " ,nMovNroRef=ISNULL(MR.nMovNro,0)" 'ARCV 06-09
        sSql = sSql & " From MovCol M Inner Join MovColDet MD ON M.nMovNro = MD.nMovNro "
        sSql = sSql & " AND M.cOpeCod = MD.cOpeCod " 'ARCV 28-05-07
        sSql = sSql & " AND MD.nPrdConceptoCod = 1101 "
        sSql = sSql & "     Left Join MovRef MR ON M.nMovNro = MR.nMovNroRef " 'ARCV 06-09
        sSql = sSql & " Where M.nMovNro = " & pnMovNro & " And MD.cCtaCod = '" & psCtaCod & "' AND SUBSTRING(CONVERT(varchar(4),MD.nPrdConceptoCod),1,2)<>'12' AND M.cCtaCod = '" & psCtaCod & "' and M.cOpeCod not Like '107[123456789]%'"
        sSql = sSql & " Group By  M.nCredEstado,M.nDiasMora, M.nFlag "
        sSql = sSql & " ,MR.nMovNro" 'ARCV 06-09
        Set RecuperaDatosExtorno = oConecta.CargaRecordSet(sSql)
    End If

   ' Set RecuperaDatosExtorno = oConecta.CargaRecordSet(sSQL)
    If RecuperaDatosExtorno.RecordCount = 0 Then
        sSql = "Select M.nCredEstado, 0 as nCapital, "
        sSql = sSql & " MIN(MD.nNroCuota) nMinCuota, M.nDiasMora, M.nFlag "
        sSql = sSql & " ,nMovNroRef=ISNULL(MR.nMovNro,0)" 'ARCV 06-09
        sSql = sSql & " From MovCol M Inner Join MovColDet MD ON M.nMovNro = MD.nMovNro "
        sSql = sSql & " AND M.cOpeCod = MD.cOpeCod " 'ARCV 28-05-07
        sSql = sSql & " AND Left(Cast(nPrdConceptoCod as Varchar(10)),2)='12' "
        sSql = sSql & "     Left Join MovRef MR ON M.nMovNro = MR.nMovNroRef " 'ARCV 06-09
        sSql = sSql & " Where M.nMovNro = " & pnMovNro & " And MD.cCtaCod = '" & psCtaCod & "' AND Left(Cast(nPrdConceptoCod as Varchar(10)),2)='12' AND M.cCtaCod = '" & psCtaCod & "' and M.cOpeCod not Like '107[123456789]%'"
        sSql = sSql & " Group By  M.nCredEstado,M.nDiasMora, M.nFlag "
        sSql = sSql & " ,MR.nMovNro" 'ARCV 06-09
        Set RecuperaDatosExtorno = oConecta.CargaRecordSet(sSql)
    End If
    
    If RecuperaDatosExtorno.RecordCount = 0 Then 'LUCV20160530, Agrego - ERS004-2016
    sSql = "Select M.nCredEstado, SUM(MD.nMonto)+isnull(dbo.FncDevolverPagadoMiViviendaxnMovNro(MR.nMovNro),0) as nCapital, "
    sSql = sSql & " MIN(MD.nNroCuota) nMinCuota, M.nDiasMora, M.nFlag "
    sSql = sSql & " ,nMovNroRef=ISNULL(MR.nMovNro,0)"
    sSql = sSql & " From MovCol M Inner Join MovColDet MD ON M.nMovNro = MD.nMovNro "
    sSql = sSql & " AND M.cOpeCod = MD.cOpeCod "
    sSql = sSql & " AND MD.nPrdConceptoCod = 1000 "
    sSql = sSql & " Left Join MovRef MR ON M.nMovNro = MR.nMovNroRef "
    sSql = sSql & " Where M.nMovNro = " & pnMovNro & " And MD.cCtaCod = '" & psCtaCod & "' "
    sSql = sSql & " AND SUBSTRING(CONVERT(varchar(4),MD.nPrdConceptoCod),1,2)<>'12' AND M.cCtaCod = '" & psCtaCod & "' "
    sSql = sSql & " and M.nCredEstado<>0  and (MD.COPECOD LIKE '101[1234567]%' or MD.COPECOD LIKE '1008[01][0123456]%') and M.cOpeCod not Like '107[123456789]%'"
    sSql = sSql & " Group By  M.nCredEstado,M.nDiasMora, M.nFlag "
    sSql = sSql & " ,MR.nMovNro"
    Set RecuperaDatosExtorno = oConecta.CargaRecordSet(sSql)
    End If 'Fin->LUCV20160530, Agrego - ERS004-2016

    oConecta.CierraConexion
    Set oConecta = Nothing

End Function


Public Function RecuperaDatosMetasAnalistaReporte(ByVal psPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaDatosMetasAnalistaReporte
    sSql = "Select DISTINCT M.cPersCod,M.dInicial,M.dFinal,M.nTipoMeta, C.cConsDescripcion as cTipoMeta, M.nMonto, M.nTipoCred, C2.cConsDescripcion cTipoCred, M.nMontoAlc, C3.cConsDescripcion as sMoneda "
    sSql = sSql & " From ColocMetasAnalista M Inner Join Constante C ON M.nTipoMeta = C.nConsValor AND C.nConsCod = " & gColocTipoMetas
    sSql = sSql & "                           Inner Join Constante C2 ON M.nTipoCred = C2.nConsValor  AND C2.nConsCod = 1001"
    sSql = sSql & "                           Inner Join Constante C3 ON C3.nConsValor = M.nMoneda AND C3.nConsCod = " & COMDConstantes.gMoneda
    sSql = sSql & " Inner Join ColocCredConsFiltro CF ON CF.nConsCod = C2.nConsCod AND CF.nCodFiltro =1 "
    sSql = sSql & " Where M.cPersCod = '" & psPersCod & "'  ORDER BY M.dInicial,cTipoMeta,sMoneda,nTipoCred"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosMetasAnalistaReporte = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaDatosMetasAnalistaReporte:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaColocFteIngreso(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaColocacFteIngreso
    sSql = "Select * from ColocFteIngreso"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaColocFteIngreso = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaColocacFteIngreso:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function


Public Function RecuperaDatosMetasAnalistaCab(ByVal psPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaDatosMetasAnalistaCab
    sSql = " Select M.cPersCod,M.dInicial,M.dFinal,M.nTipoMeta, C.cConsDescripcion as cTipoMeta, CN.cConsDescripcion as sMoneda, M.nMoneda "
    sSql = sSql & " From ColocMetasAnalista M Inner Join Constante C ON M.nTipoMeta = C.nConsValor AND C. nConsCod = " & gColocTipoMetas
    sSql = sSql & "     Inner Join Constante CN ON CN.nConsValor = M.nMoneda AND CN.nConsCod = " & COMDConstantes.gMoneda
    sSql = sSql & " Where M.cPersCod = '" & psPersCod & "'"
    sSql = sSql & "Group By M.cPersCod,M.dInicial,M.dFinal,M.nTipoMeta,C.cConsDescripcion,CN.cConsDescripcion,M.nMoneda "
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosMetasAnalistaCab = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaDatosMetasAnalistaCab:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

'Modificado:    20060405
'Método que lista los montos colocados en función a las
'metas del analsita.
'Este método fue modificado para retornar el monto colocado
'y el número de créditos correspondiente al monto y condición del
'crédito.
'Autor:         Pedro Mucha
'Descripcion de cambio:     Se agregó un parámetro opcional
'rsMetaAnalista que contendrá los valores de monto y
'numero de creditos por condicion.
'Este parametro es opcional para fines de compatibilidad con
'el código que puede hacer referencia a este método.
Public Function MontoColocadoCredCondicion(ByVal pdFechaIni As Date, _
    ByVal pdFechaFin As Date, ByVal pnMoneda As Integer, ByVal psProducto As String, _
    ByVal psPersCod As String, ByVal pnCondicion As ColocCredCondicion, Optional ByVal pnTodos As Boolean = False, _
    Optional rsMetaAnalista As ADODB.Recordset = Nothing) As Double

Dim oCon As COMConecta.DCOMConecta
Dim sSql As String
Dim R As ADODB.Recordset
    
    '20060405
    'en esta parte se agrega un campo: COUNT(C.cCtaCod) as nNumeroCred
    'que contendrá el número de creditos correspondiente al
    'la condicion de credito.
    sSql = "Select ISNULL(SUM(C.nMontoCol),0) as nMontoCol, COUNT(C.cCtaCod) as nNumeroCred from Colocaciones C "
    sSql = sSql & " Inner Join ColocacCred CC ON C.cCtaCod = CC.cCtaCod "
    sSql = sSql & " Inner Join ProductoPersona PP ON PP.cCtaCod = C.cCtaCod AND PP.nPrdPersRelac = 28"
    If pnTodos Then
        sSql = sSql & " Where SUBSTRING(C.cCtaCod,6,3) = '" & psProducto & "' "
    Else
        sSql = sSql & " Where CC.nColocCondicion = " & pnCondicion & " AND SUBSTRING(C.cCtaCod,6,3) = '" & psProducto & "' "
    End If
    sSql = sSql & " AND SUBSTRING(C.cCtaCod,9,1) = " & pnMoneda & " AND ( C.dVigencia >= '" & Format(pdFechaIni, "mm/dd/yyyy") & "' AND C.dVigencia <= '" & Format(pdFechaFin, "mm/dd/yyyy") & "') "
    sSql = sSql & " AND PP.cPersCod = '" & psPersCod & "'"

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSql)
    
    '20060405
    'hacemos la verificacion del recordset
    If Not (rsMetaAnalista Is Nothing) Then
        'asignamos el recordset a rsMetaAnalista
        Set rsMetaAnalista = R
        
    End If
    
    oCon.CierraConexion
    Set oCon = Nothing

    MontoColocadoCredCondicion = IIf(IsNull(R!nMontoCol), 0, R!nMontoCol)

End Function


Public Function RecuperaDatosMetasAnalistaDetalle(ByVal psPersCod As String, ByVal pnTipo As Integer, _
    ByVal pdInicio As Date, ByVal pdFinal As Date, ByVal pnMoneda As Integer) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaDatosMetasAnalistaDetalle

    'sSql = "Select nTipoCred,dInicial,dFinal,nMonto"
    'sSql = sSql & " From ColocMetasAnalista Where cPersCod = '" & psPersCod & "'"
    'sSql = sSql & " AND nTipoMeta = " & pnTipo & " AND  dInicial = '" & Format(pdInicio, "mm/dd/yyyy")
    'sSql = sSql & "' AND dFinal = '" & Format(pdFinal, "mm/dd/yyyy") & "' AND nMoneda = " & pnMoneda

    'sSql = "Select DISTINCT nTipoCred,dInicial,dFinal,nMonto"
    sSql = "Select DISTINCT nTipoCred,nMonto"
    sSql = sSql & " From ColocMetasAnalista Where cPersCod = '" & psPersCod & "'"
    sSql = sSql & " AND nTipoMeta = " & pnTipo & " AND  dInicial = '" & Format(pdInicio, "mm/dd/yyyy") & "' "
    sSql = sSql & " AND nMoneda = " & pnMoneda & " AND dFinal = '" & Format(pdFinal, "mm/dd/yyyy") & "' "
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosMetasAnalistaDetalle = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaDatosMetasAnalistaDetalle:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function


Public Function RecuperaCreditosParaAsignarGasto(ByVal pnRanIni As Double, ByVal pnRanFin As Double, _
    ByVal pnAplicado As COMDConstantes.ColocCalendApl, ByVal psTipoProd As String, ByVal psMoneda As String, _
    ByVal psCodGasto As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCreditosParaAsignarGasto
    sSql = "Select distinct '',P.cCtaCod, RTRIM(Pers.cPersNombre), P.nSaldo"
    'ALPA 20100616****************
    'sSql = sSql & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
    sSql = sSql & " From Colocaciones Col inner join  Producto P on (Col.cCtaCod=P.cCtaCod) Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
    'sSql = sSql & "     AND substring(P.cCtaCod,6,1) = '" & psTipoProd & "' AND P.nPrdEstado in (" & gColocEstRefMor & "," & COMDConstantes.gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & ")"
    sSql = sSql & "     AND substring(Col.cTpoCredCod,1,1) = '" & psTipoProd & "' AND P.nPrdEstado in (" & gColocEstRefMor & "," & COMDConstantes.gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & ")"
    '****************************
    sSql = sSql & "  AND substring(P.cCtaCod,9,1)='" & psMoneda & "'"
    sSql = sSql & "                 Inner Join Persona Pers ON PP.cPersCod = Pers.cPersCod"
    sSql = sSql & "                 Inner Join ColocacCred CC ON P.cCtaCod = CC.cCtaCod"
    sSql = sSql & "                 Inner Join ColocCalendario CL ON P.cCtaCod = CL.cCtaCod AND CL.nNroCalen=CC.nNroCalen and CL.nColocCalendApl = " & pnAplicado & " AND CL.nColocCalendEstado = " & gColocCalendEstadoPendiente
    sSql = sSql & "                 Inner Join ColocCalendDet CD ON P.cCtaCod=CD.cCtaCod AND CD.nNroCalen = CC.nNroCalen AND CD.nColocCalendApl = " & pnAplicado & " And nPrdConceptoCod <> " & CLng(psCodGasto)
    sSql = sSql & " Group By P.cCtaCod, Pers.cPersNombre, P.nSaldo, CD.nCuota, CC.nNroCalen "
    sSql = sSql & " HAVING SUM(CD.nMonto-isnull(CD.nMontopagado,0)) >= " & Format(pnRanIni, "#0.00") & " AND SUM(CD.nMonto-isnull(CD.nMontopagado,0)) <= " & Format(pnRanFin, "#0.00")
    sSql = sSql & " and CD.nCuota not in (select nCuota from ColocCalenddet where nCuota = CD.nCuota and nNroCalen = CC.nNrocalen And cCtaCod = P.cCtaCod AND nPrdConceptoCod =" & psCodGasto & ")"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditosParaAsignarGasto = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCreditosParaAsignarGasto:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaCuotasParaAsignarGasto(ByVal pnRanIni As Double, ByVal pnRanFin As Double, _
    ByVal pnAplicado As COMDConstantes.ColocCalendApl, ByVal psCtaCod As String, ByVal psCodGasto As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCuotasParaAsignarGasto
    sSql = "Select distinct '',P.cCtaCod, Pers.cPersNombre, P.nSaldo, CD.nCuota, CC.nNroCalen, SUM(CD.nMonto-isnull(CD.nMontopagado,0)) as nMontoCuota "
    sSql = sSql & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
    sSql = sSql & "                 Inner Join Persona Pers ON PP.cPersCod = Pers.cPersCod"
    sSql = sSql & "                 Inner Join ColocacCred CC ON P.cCtaCod = CC.cCtaCod"
    sSql = sSql & "                 Inner Join ColocCalendario CL ON P.cCtaCod = CL.cCtaCod AND CL.nNroCalen=CC.nNroCalen and CL.nColocCalendApl = " & pnAplicado & " AND CL.nColocCalendEstado = " & gColocCalendEstadoPendiente
    sSql = sSql & "                 Inner Join ColocCalendDet CD ON P.cCtaCod=CD.cCtaCod AND CD.nNroCalen = CC.nNroCalen AND CD.nColocCalendApl = " & pnAplicado
    sSql = sSql & "  WHERE P.cCtaCod = '" & psCtaCod & "'"
    sSql = sSql & " Group By P.cCtaCod, Pers.cPersNombre, P.nSaldo, CD.nCuota, CC.nNroCalen "
    sSql = sSql & " HAVING SUM(CD.nMonto-isnull(CD.nMontopagado,0)) >= " & Format(pnRanIni, "#0.00") & " AND SUM(CD.nMonto-isnull(CD.nMontopagado,0)) <= " & Format(pnRanFin, "#0.00")
    sSql = sSql & " and CD.nCuota not in (select nCuota from ColocCalenddet where nCuota = CD.nCuota and nNroCalen = CC.nNrocalen And cCtaCod = P.cCtaCod AND nPrdConceptoCod =" & psCodGasto & ")"


    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCuotasParaAsignarGasto = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCuotasParaAsignarGasto:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaTiposCredito() As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String
    Set oConecta = New COMConecta.DCOMConecta
    'ALPA 20100616************
    'sSql = "Select nConsCod, nConsValor, cConsDescripcion from Constante where nConsCod = " & COMDConstantes.gProducto & " and (nConsValor % 100) = 0  order by  cConsDescripcion"
    sSql = "Select nConsCod, nConsValor, cConsDescripcion from Constante where nConsCod = " & COMDConstantes.gTpoCredito & " and (nConsValor % 100) = 0  order by  cConsDescripcion"
    '*************************
    oConecta.AbreConexion
    Set RecuperaTiposCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function RecuperaCreditosParaJudicalTotal(Optional ByVal psPersCod As String = "", Optional psCtaCod As String = "") As ADODB.Recordset
Dim oParam As DCOMParametro
Dim nValor As Double
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCreditosParaJudicalTotal
    Set oParam = New DCOMParametro
    nValor = oParam.RecuperaValorParametro(COMDConstantes.gColocEstJudicial)
    Set oParam = Nothing

    If psPersCod <> "" Then
        sSql = "Select P.cCtaCod, C.nMontoCol, P.nSaldo, P.nTasaInteres, P.nPrdEstado, CC.nNroCalen, "
        sSql = sSql & " cAnalista=(Select cPersNombre From Persona Where cPersCod = PP.cPersCod),"
        sSql = sSql & " cTitular=(Select cPersNombre From Persona Where cPersCod = PP2.cPersCod), CC.nDiasAtraso "
        sSql = sSql & " From Producto P Inner join Colocaciones C ON P.cCtaCod = C.cCtaCod "
        sSql = sSql & "                 Inner Join ColocacCred CC ON P.cCtaCod = CC.cCtaCod "
        sSql = sSql & "                 Left Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & COMDConstantes.gColRelPersAnalista
        sSql = sSql & "                 Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod And PP2.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
        sSql = sSql & " WHERE PP2.cPersCod = '" & psPersCod & "'"
        sSql = sSql & " AND P.nPrdEstado in (" & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstRefMor & "," & COMDConstantes.gColocEstRefNorm & "," & gColocEstRefVenc & ")"
        'sSql = sSql & " AND CC.nDiasAtraso > " & nValor
    Else
        If psCtaCod <> "" Then
            sSql = "Select P.cCtaCod, C.nMontoCol, P.nSaldo, P.nTasaInteres, P.nPrdEstado, CC.nNroCalen, "
            sSql = sSql & " cAnalista=(Select cPersNombre From Persona Where cPersCod = PP.cPersCod),"
            sSql = sSql & " cTitular=(Select cPersNombre From Persona Where cPersCod = PP2.cPersCod), CC.nDiasAtraso "
            sSql = sSql & " From Producto P Inner join Colocaciones C ON P.cCtaCod = C.cCtaCod "
            sSql = sSql & "                 Inner Join ColocacCred CC ON P.cCtaCod = CC.cCtaCod "
            sSql = sSql & "                 Left Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & COMDConstantes.gColRelPersAnalista
            sSql = sSql & "                 Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod And PP2.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
            sSql = sSql & " WHERE P.cCtaCod = '" & psCtaCod & "'"
            sSql = sSql & " AND P.nPrdEstado in (" & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstRefMor & "," & COMDConstantes.gColocEstRefNorm & "," & gColocEstRefVenc & ")"
        Else
            sSql = "Select P.cCtaCod, C.nMontoCol, P.nSaldo, P.nTasaInteres, P.nPrdEstado, CC.nNroCalen, "
            sSql = sSql & " cAnalista=(Select cPersNombre From Persona Where cPersCod = PP.cPersCod),"
            sSql = sSql & " cTitular=(Select cPersNombre From Persona Where cPersCod = PP2.cPersCod), CC.nDiasAtraso "
            sSql = sSql & " From Producto P Inner join Colocaciones C ON P.cCtaCod = C.cCtaCod "
            sSql = sSql & "                 Inner Join ColocacCred CC ON P.cCtaCod = CC.cCtaCod "
            sSql = sSql & "                 Left Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & COMDConstantes.gColRelPersAnalista
            sSql = sSql & "                 Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod And PP2.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
            sSql = sSql & " WHERE CC.nDiasAtraso > " & nValor
            sSql = sSql & " AND P.nPrdEstado in (" & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstRefMor & "," & COMDConstantes.gColocEstRefNorm & "," & gColocEstRefVenc & ")"
        End If
    End If
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditosParaJudicalTotal = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaCreditosParaJudicalTotal:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

'*** PEAC 20080530
Public Function RecuperaCapacidadPago(ByVal pcCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCapacidadPago
  
    sSql = "exec stp_sel_RecuperaCapacidadPago '" & pcCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCapacidadPago = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCapacidadPago:
                  Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

'*** PEAC 20080412
'MADM 20111215
Public Function RecuperaCaliSbs(ByVal psDocNat As String, ByVal psDocJur As String, Optional ByVal psTpoDoc As Integer = 1) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCaliSbs
  
    sSql = "exec stp_sel_ObtieneCaliSbs '" & psDocNat & "','" & psDocJur & "', " & psTpoDoc & " "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCaliSbs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCaliSbs:
                  Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function
'END MADM

Public Function RecuperaCreditosRefinanciados(ByVal psCtaCod As String, Optional ByVal pbSolicitados As Boolean = False) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCreditosRefinanciados
      'CTI6-20210503-ERS032-2019
'    If Not pbSolicitados Then
'        sSQL = "Select Distinct CR.cCtaCodRef, "
'        sSQL = sSQL & "nCapitalRef = (Select nMonto from ColocacRefinancDet Where cCtaCod = CR.cCtaCod AND cCtaCodRef = CR.cCtaCodRef AND nEstado = " & gColocEstadoRefinancAprobado & " AND nPrdConceptoCod = " & gColocConceptoCodCapital & "),"
'        sSQL = sSQL & "nInteresRef = (Select SUM(nMonto) from ColocacRefinancDet Where cCtaCod = CR.cCtaCod AND cCtaCodRef = CR.cCtaCodRef AND nEstado = " & gColocEstadoRefinancAprobado & " AND nPrdConceptoCod <> " & gColocConceptoCodCapital & ")"
'        sSQL = sSQL & " from ColocacRefinanc CR "
'        sSQL = sSQL & " WHERE CR.cCtaCod = '" & psCtaCod & "'"
'    Else
'        sSQL = "Select Distinct CR.cCtaCodRef, "
'        sSQL = sSQL & "nCapitalRef = (Select nMonto from ColocacRefinancDet Where cCtaCod = CR.cCtaCod AND cCtaCodRef = CR.cCtaCodRef AND nEstado = " & gColocEstadoRefinancSolicitado & " AND nPrdConceptoCod = " & gColocConceptoCodCapital & "),"
'        sSQL = sSQL & "nInteresRef = (Select SUM(nMonto) from ColocacRefinancDet Where cCtaCod = CR.cCtaCod AND cCtaCodRef = CR.cCtaCodRef AND nEstado = " & gColocEstadoRefinancSolicitado & " AND nPrdConceptoCod <> " & gColocConceptoCodCapital & ")"
'        sSQL = sSQL & " from ColocacRefinanc CR "
'        sSQL = sSQL & " WHERE CR.cCtaCod = '" & psCtaCod & "'"
'    End If
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = "exec stp_sel_DatosPosicionRecuperaCreditosRefinanciados '" & psCtaCod & "', " & IIf(pbSolicitados = True, 1, 0) & ""
    'Fin CTI6-20210503-ERS032-2019
    Set RecuperaCreditosRefinanciados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCreditosRefinanciados:
                  Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

'LUCV20160617, Segun ERS004-2016 **********
Public Function RecuperaColocacRefinancSolicitado(ByVal psCtaCod As String, ByVal pnEstado As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorRecuperaColocacRefinancSolicitado
    sSql = "Select Top 1 * from ColocacRefinanc where cCtacod = '" & psCtaCod & "' and nEstado = " & pnEstado & ""
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaColocacRefinancSolicitado = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function
ErrorRecuperaColocacRefinancSolicitado:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'Fin -> LUCV20160617 **********

Public Function RecuperaColocacRefinanc(ByVal psCtaCod As String, Optional ByVal nEstadoAproba = -1) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorRecuperaColocacRefinanc
    'CTI6-20210503-ERS032-2019 -(Optimizar Sugerencia)
'    sSql = "Select * from ColocacRefinanc where cCtacod = '" & psCtaCod & "'"
'    'LUCV20160816, Agrego Según ERS004-2016 ->**********
'    If nEstadoAproba <> -1 Then
'    sSql = sSql & " And nEstado = " & nEstadoAproba & " "
'    End If
'    'LUCV20160816, Fin <-**********
    
    sSql = "exec stp_sel_RecuperaColocacRefinanc '" & psCtaCod & "', " & nEstadoAproba & " "
    'Fin CTI6-20210503
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaColocacRefinanc = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaColocacRefinanc:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function


Public Function RecuperaColocaciones(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaColocaciones
    
    '*** peac 20071201
    'sSql = "Select * from Colocaciones where cCtaCod = '" & psCtaCod & "'"
    
    '->***** LUCV20190502, Comentó y agregó.
    '*** peac 20071204
'    sSql = " Select a.*, isnull(w.nmonto,0) nMontoColUltCal, isnull(w.dvenc,0) dMontoColUltCal"
'    sSql = sSql & " from Colocaciones a"
'    sSql = sSql & " Left Join"
'    sSql = sSql & " (SELECT cc.cctacod,cc.nmonto, ccc.dvenc"
'    sSql = sSql & " from coloccalenddet cc"
'    sSql = sSql & " inner join colocaccred c on cc.cctacod=c.cctacod and c.nnrocalen=cc.nnrocalen"
'    sSql = sSql & " inner join coloccalendario ccc on ccc.cctacod=cc.cctacod and ccc.nnrocalen=cc.nnrocalen and ccc.ncoloccalendapl=0"
'    sSql = sSql & " where cc.ncoloccalendapl=0 and nprdconceptocod=1000) w on  a.cctacod=w.cctacod"
'    sSql = sSql & " where a.cCtaCod = '" & psCtaCod & "'"
    
    sSql = "exec stp_sel_RecuperaColocaciones '" & psCtaCod & "'"
    'Fin LUCV20190502
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaColocaciones = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaColocaciones:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaGastosExonerados(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    sSql = "select * from ColocCredGastosExon Where cCtaCod = '" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaGastosExonerados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

'*** PEAC 20080412
Public Function RecuperaDatosInfCome(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer

    On Error GoTo ErrorRecuperaInfCome

    sSql = "exec stp_sel_ObtieneDatosInformeComercial '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosInfCome = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function
ErrorRecuperaInfCome:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

'*** PEAC 20080412
'Public Function RecuperaDatosBalGen(ByVal psNumFte As String) As ADODB.Recordset 'LUCV20160815, Comentó-> Según: ERS004-2016
Public Function RecuperaDatosBalGen(ByVal psCtaCod As String) As ADODB.Recordset 'LUCV20160815, Se agregó la CtaCod reemplazando a NumFte
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaBal
    'sSql = "exec stp_sel_ObtieneDatosBalance '" & psNumFte & "'"
    sSql = "exec stp_sel_ObtieneDatosBalance '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosBalGen = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function
ErrorRecuperaBal:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function


'*** PEAC 20080412
Public Function RecuperaDatosInfCome01(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer

    On Error GoTo ErrorRecuperaInfCome

    sSql = "exec stp_sel_ObtieneDatosInformeComercial '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosInfCome01 = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function
ErrorRecuperaInfCome:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function


'MARG 20160719
Public Function RecuperaDatosComunes(ByVal psCtaCod As String, Optional pnTodos As Boolean = True, Optional ByVal MatEst As Variant = Nothing) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer

Dim sEstados As String  ''20090212

    On Error GoTo ErrorRecuperaDatosComunes
    
''''    sSQL = "Select P.nPrdEstado, CC.nCalendDinamico, CC.bCuotaCom, CL.nTasaIni as nTasaMora, CC.nCalendDinamTipo, CC.bPrepago, CC.bMiVivienda, P.cCtaCod, CE.nMonto nMontoSol, C.dVigencia,C.nMontoCol, CC.cMetLiquidacion, P.nSaldo, P.nTasaInteres, CC.nNroProxCuota, "
''''    sSQL = sSQL & " cAnalista=(Select cPersNombre From Persona Where cPersCod = PP.cPersCod),"
''''    sSQL = sSQL & " Pers.cPersNombre cTitular, Pers.cPersCod,"
''''    sSQL = sSQL & " DNI=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = Pers.cPersCod AND cPersIDTpo = '" & gPersIdDNI & "'), "
''''    sSQL = sSQL & " RUC=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = Pers.cPersCod AND cPersIDTpo = '" & gPersIdRUC & "'), "
''''    sSQL = sSQL & " CC.nExoPenalidad, nNota = (select nColocNota  From ColocCalificacionAnalista Where cTipo='A' AND cCtaCod = C.cCtaCod And dColocNotaFecha = (select Max(dColocNotaFecha) from ColocCalificacionAnalista Where cCtaCod = C.cCtaCod AND cTipo='A')), "
''''    sSQL = sSQL & " CN2.cConsDescripcion as cTipoCredDescrip, "
''''    sSQL = sSQL & " CN.cConsDescripcion as cDestinoDescripcion, CC.nDiasAtrasoAcum, "
''''    sSQL = sSQL & " CN3.cConsDescripcion as cMoneda, C.nMontoCol, CN4.cConsDescripcion cEstado, CN5.cConsDescripcion cCondicion, "
''''    sSQL = sSQL & " cLineaCred = (Select cLineaCred from Colocaciones Where cCtaCod = C.cCtaCod), "
''''    sSQL = sSQL & " dFechaSol = (Select dPrdEstado from ColocacEstado where cCtaCod = P.cCtaCod And nPrdEstado = " & gColocEstSolic & "), "
''''    sSQL = sSQL & " dbo.ColocNotaCredito(P.cCtaCod) as nNotaSist "
''''    sSQL = sSQL & " From Producto P Inner join Colocaciones C ON P.cCtaCod = C.cCtaCod "
''''    sSQL = sSQL & "                 Inner Join ColocacCred CC ON P.cCtaCod = CC.cCtaCod "
''''    sSQL = sSQL & "                 Left Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & gColRelPersAnalista
''''    sSQL = sSQL & "                 Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod And PP2.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
''''    sSQL = sSQL & "                 Inner Join Persona Pers ON PP2.cPersCod = Pers.cPersCod"
''''    sSQL = sSQL & "                 Left Join Constante CN2 ON convert(int,substring(C.cCtaCod,6,3)) = CN2.nConsValor AND CN2.nConsCod = " & COMDConstantes.gProducto
''''    sSQL = sSQL & "                 Left Join Constante CN ON CC.nColocDestino = CN.nConsValor AND CN.nConsCod = " & Trim(Str(COMDConstantes.gColocDestino))
''''    sSQL = sSQL & "                 Left Join Constante CN3 ON convert(int,substring(C.cCtaCod,9,1)) = CN3.nConsValor AND CN3.nConsCod = " & COMDConstantes.gMoneda
''''    sSQL = sSQL & "                 Left Join Constante CN4 ON P.nPrdEstado = CN4.nConsValor AND CN4.nConsCod = " & COMDConstantes.gColocEstado
''''    sSQL = sSQL & "                 Left Join Constante CN5 ON CN5.nConsValor = CC.nColocCondicion AND CN5.nConsCod = " & COMDConstantes.gColocCredCondicion
''''    sSQL = sSQL & "                 Inner Join ColocacEstado CE ON CE.cCtaCod = P.cCtaCod AND CE.nPrdEstado =" & gColocEstSolic
''''    sSQL = sSQL & "                 Left Join ColocLineaCreditoTasa CL ON CL.cLineaCred = C.cLineaCred AND nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal
''''    sSQL = sSQL & " WHERE P.cCtaCod = '" & psCtaCod & "' AND CC.nCalendDinamico = 0 " 'AND P.nPrdEstado in (2020,2021,2022,2030,2031,2032)"
''''    If Not pnTodos Then
''''        sSQL = sSQL & " AND P.nPrdEstado in (" & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstRefMor & "," & COMDConstantes.gColocEstRefNorm & "," & gColocEstRefVenc & ")"
''''    Else
''''        If IsArray(MatEst) Then
''''            sSQL = sSQL & " AND P.nPrdEstado in ("
''''            For I = 0 To UBound(MatEst)
''''                sSQL = sSQL & MatEst(I) & ","
''''            Next I
''''            sSQL = Mid(sSQL, 1, Len(sSQL) - 1) & ")"
''''        End If
''''    End If
    'CE2.nPeriodoFechaFija


''***    PEAC 20090212 - COMENTADO PORQUE SE CREO UN STP

'    sSql = "Select P.nPrdEstado, CC.nCalendDinamico, CC.bCuotaCom, CL.nTasaIni as nTasaMora, CC.nCalendDinamTipo, CC.bPrepago, CC.bMiVivienda, P.cCtaCod, CE.nMonto nMontoSol, C.dVigencia,C.nMontoCol, CC.cMetLiquidacion, P.nSaldo, P.nTasaInteres, CC.nNroProxCuota, "
'    sSql = sSql & " cAnalista=(Select cPersNombre From Persona Where cPersCod = PP.cPersCod),"
'    sSql = sSql & " Pers.cPersNombre cTitular, Pers.cPersCod, "
'    sSql = sSql & " DNI=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = Pers.cPersCod AND cPersIDTpo = '" & COMDConstantes.gPersIdDNI & "'), "
'    sSql = sSql & " RUC=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = Pers.cPersCod AND cPersIDTpo = '" & COMDConstantes.gPersIdRUC & "'), "
'    sSql = sSql & " CC.nExoPenalidad, nNota = (select nColocNota  From ColocCalificacionAnalista Where cTipo='A' AND cCtaCod = C.cCtaCod And dColocNotaFecha = (select Max(dColocNotaFecha) from ColocCalificacionAnalista Where cCtaCod = C.cCtaCod AND cTipo='A')), "
'    sSql = sSql & " CC.nExoPenalidad, nNota = (select nColocNota  From ColocCalificacionAnalista Where cTipo='A' AND cCtaCod = C.cCtaCod And dColocNotaFecha = (select Max(dColocNotaFecha) from ColocCalificacionAnalista Where cCtaCod = C.cCtaCod AND cTipo='A')), "
'    sSql = sSql & " CN2.cConsDescripcion as cTipoCredDescrip, "
'    sSql = sSql & " CN.cConsDescripcion as cDestinoDescripcion, CC.nDiasAtrasoAcum, "
'    sSql = sSql & " CN3.cConsDescripcion as cMoneda, C.nMontoCol,"
'    sSql = sSql & " nMontoCol = CASE WHEN ISNULL(MC.nMovNro,0)<>0 THEN MC.nMonto ELSE C.nMontoCol END,"
'    sSql = sSql & " CN4.cConsDescripcion cEstado, CN5.cConsDescripcion cCondicion, "
'    sSql = sSql & " cLineaCred = (Select cLineaCred from Colocaciones Where cCtaCod = C.cCtaCod), "
'    sSql = sSql & " dFechaSol = (Select dPrdEstado from ColocacEstado where cCtaCod = P.cCtaCod And nPrdEstado = " & COMDConstantes.gColocEstSolic & "), "
'    sSql = sSql & " dbo.ColocNotaCredito(P.cCtaCod) as nNotaSist,C.nPlazo,dVenc," 'CUSCO nPlazo
'    sSql = sSql & " nCuotaSug = isnull ((select sum(CA.nMonto) from ColocCalenddet CA inner join ColocacEstado CE  ON CA.cCtaCod=CE.cCtaCod"
'    sSql = sSql & "                      where  CA.cCtaCod = C.cCtaCod and  nColocCalendCod in (10,11,40,41) and nColocCalendApl=1 and nCuota=1 and nNroCalen=3 and nPrdEstado=2020),0),"
'    sSql = sSql & " nSeguroD = isnull((select TOP 1 nMonto from ColocCalenddet CCD where CCD.cctacod =C.cCtaCod and ncuota=1 "
'    sSql = sSql & "            and nnrocalen =(select nnrocalen From ColocacCred CC1 where CC1.cctacod=CCD.cCtaCod)"
'    sSql = sSql & "            and nPrdConceptoCod in(select nParamValor from ColocParametro where nParamVar in (3060,3063))),0),"
'    sSql = sSql & " nSeguroDH = isnull((select TOP 1 nMonto from ColocCalenddet CCD1 where CCD1.cctacod =C.cCtaCod and ncuota=1 "
'    sSql = sSql & "            and nnrocalen =(select nnrocalen From ColocacCred CC2 where CC2.cctacod=CCD1.cCtaCod)"
'    sSql = sSql & "            and nPrdConceptoCod in(select nParamValor from ColocParametro where nParamVar in (3061,3064))),0),"
'    sSql = sSql & " nSeguroM = isnull((select TOP 1 nMonto from ColocCalenddet CCD2 where CCD2.cctacod =C.cCtaCod and ncuota=1 "
'    sSql = sSql & "            and nnrocalen =(select nnrocalen From ColocacCred CC3 where CC3.cctacod=CCD2.cCtaCod)"
'    sSql = sSql & "            and nPrdConceptoCod in(select nParamValor from ColocParametro where nParamVar in (3062,3065))),0)"
'    sSql = sSql & " ,dVencPagare= ( SELECT MIN(dVenc)FROM ColocCalendario WHERE nColocCalendApl=1 AND cCtaCod= P.cCtaCod AND nNroCalen=CC.nNroCalen)" 'ARCV 20-01-2007
'    sSql = sSql & " ,nMontoPagare= ( SELECT nMonto FROM ColocacEstado WHERE cCtaCod= P.cCtaCod AND nPrdEstado=P.nPrdEstado AND dPrdEstado=(SELECT MAX(dPrdEstado)FROM ColocacEstado WHERE cCtaCod=P.cCtaCod AND nPrdEstado=P.nPrdEstado )) " 'ARCV 29-01-2007
'    sSql = sSql & " ,nFecPgoCtas = case when substring(P.cCtaCod,6,3) in ('401','403','423') then  " 'PEAC 20070801 - se agrego para mostrar en la hoja resumen la fecha de pago de las cuotas
'    sSql = sSql & " (select (case "
'    sSql = sSql & "          when nColocCalendCod between '10' and '39' then nPlazo"
'    sSql = sSql & "          when nColocCalendCod between '40' and '69' then nPeriodoFechaFija"
'    sSql = sSql & "      end) from ColocacEstado where cctacod=p.cctacod and nPrdEstado='2002' and cDescripcion like 'Aprobacion%')"
'    sSql = sSql & "else 0 End"
'    sSql = sSql & " From Producto P Inner join Colocaciones C ON P.cCtaCod = C.cCtaCod "
'    sSql = sSql & "                 Inner Join ColocacCred CC ON P.cCtaCod = CC.cCtaCod "
'    sSql = sSql & "                 Left Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & COMDConstantes.gColRelPersAnalista
'    sSql = sSql & "                 Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod And PP2.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
'    sSql = sSql & "                 Inner Join Persona Pers ON PP2.cPersCod = Pers.cPersCod"
'    sSql = sSql & "                 Left Join Constante CN2 ON convert(int,substring(C.cCtaCod,6,3)) = CN2.nConsValor AND CN2.nConsCod = " & COMDConstantes.gProducto
'    sSql = sSql & "                 Left Join Constante CN ON CC.nColocDestino = CN.nConsValor AND CN.nConsCod = " & Trim(Str(COMDConstantes.gColocDestino))
'    sSql = sSql & "                 Left Join Constante CN3 ON convert(int,substring(C.cCtaCod,9,1)) = CN3.nConsValor AND CN3.nConsCod = " & COMDConstantes.gMoneda
'    sSql = sSql & "                 Left Join Constante CN4 ON P.nPrdEstado = CN4.nConsValor AND CN4.nConsCod = " & COMDConstantes.gColocEstado
'    sSql = sSql & "                 Left Join Constante CN5 ON CN5.nConsValor = CC.nColocCondicion AND CN5.nConsCod = " & COMDConstantes.gColocCredCondicion
'    sSql = sSql & "                 Inner Join ColocacEstado CE ON CE.cCtaCod = P.cCtaCod AND CE.nPrdEstado =" & COMDConstantes.gColocEstSolic
''    sSql = sSql & "                 Inner Join ColocacEstado CE2 ON CE2.cCtaCod = P.cCtaCod AND CE2.nPrdEstado =" & COMDConstantes.gColocEstVigNorm ' Add By GITU 15-08-08
'    sSql = sSql & "                 Left Join ColocLineaCreditoTasa CL ON CL.cLineaCred = C.cLineaCred AND nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal
'    sSql = sSql & "                 LEFT JOIN MovCol MC ON MC.cOpeCod='100902' AND MC.cCtaCod= P.cCtaCod  "
'    sSql = sSql & "                 AND MC.nMovNro=(SELECT MAX(nMovNro) FROM MovCol WHERE cOpeCod='100902' AND cCtaCod=P.cCtaCod)"
'    sSql = sSql & " WHERE P.cCtaCod = '" & psCtaCod & "' " 'AND P.nPrdEstado in (2020,2021,2022,2030,2031,2032)"

'    If Not pnTodos Then
'        sSql = sSql & " AND P.nPrdEstado in (" & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstRefMor & "," & COMDConstantes.gColocEstRefNorm & "," & gColocEstRefVenc & ")"
'    Else
'        If IsArray(MatEst) Then
'            sSql = sSql & " AND P.nPrdEstado in ("
'            For i = 0 To UBound(MatEst)
'                sSql = sSql & MatEst(i) & ","
'            Next i
'            sSql = Mid(sSql, 1, Len(sSql) - 1) & ")"
'        End If
'    End If


'*** PEAC 20090212 -

'    If IsArray(MatEst) Then
'            sSql = sSql & " AND P.nPrdEstado in ("
'            For i = 0 To UBound(MatEst)
'                sSql = sSql & MatEst(i) & ","
'            Next i
'            sSql = Mid(sSql, 1, Len(sSql) - 1) & ")"
'
'            sEstados = sSql & " AND P.nPrdEstado in ("
'            For i = 0 To UBound(MatEst)
'                sSql = sSql & MatEst(i) & ","
'            Next i
'            sSql = Mid(sSql, 1, Len(sSql) - 1) & ")"
'
'
'    End If

    If Not pnTodos Then
        sEstados = "2022,2020,2021,2032,2030,2031"
    Else
        If IsArray(MatEst) Then
            For i = 0 To UBound(MatEst)
                If i = 0 Then
                    sEstados = "" & MatEst(i) & ""
                Else
                    sEstados = sEstados & "," & MatEst(i) & ""
                End If
            Next i
        Else
            sEstados = " "
        End If
    End If

    sSql = " exec Stp_Sel_RecuperaDatosParaDuplicadosDocs2 '" & psCtaCod & "','" & sEstados & "'"

'*** FIN PEAC

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosComunes = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function
ErrorRecuperaDatosComunes:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaCreditosVigentesRelacCred(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCreditosVigentesRelacCred
    
'    sSql = "Select  Pers.cPersNombre, CN.cConsDescripcion cRelacion, P.cCtaCod, CC.nDiasAtraso"
'    sSql = sSql & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac not in (" & COMDConstantes.gColRelPersAnalista & "," & COMDConstantes.gColRelPersApoderado & "," & COMDConstantes.gColRelPersTitular & ")"
'    sSql = sSql & "                 Inner Join Constante CN ON CN.nConsValor = PP.nPrdPersRelac AND CN.nConsCod = " & COMDConstantes.gColocRelacPers
'    sSql = sSql & "                 Inner join Persona Pers ON Pers.cPersCod = PP.cPersCod "
'    sSql = sSql & "                 Inner join ColocacCred CC ON CC.cCtaCod = P.cCtaCod "
'    sSql = sSql & " Where Pers.cPersCod In (Select cPersCod from ProductoPersona Where cCtaCod = '" & psCtaCod & "' and nPrdPersRelac not in (" & COMDConstantes.gColRelPersAnalista & "," & COMDConstantes.gColRelPersApoderado & "," & COMDConstantes.gColRelPersTitular & "))"
'    sSql = sSql & "AND P.nPrdEstado in (" & gColocEstRefMor & "," & COMDConstantes.gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & ") "
    
    sSql = "exec stp_sel_RecuperaCreditosVigentesRelacCred '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditosVigentesRelacCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCreditosVigentesRelacCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description


End Function

Public Function RecuperaEstadosdelCredito(ByVal psCtaCod As String, ByVal pbSolicitud As Boolean, _
    ByVal pbSugerencia As Boolean, ByVal pbAprobacion As Boolean) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaEstadosdelCredito
    sSql = ""
    If pbSolicitud Then
        'MAVM 20121220 ***
        'sSql = sSql & "Select CE.nPrdEstado, CN.cConsDescripcion cEstado, CE.nMonto, CE.nCuotas, CE.nPlazo, NULL as nMontoCuota "
        sSql = sSql & "Select CE.nPrdEstado, CN.cConsDescripcion cEstado, CE.nMonto, CE.nCuotas, CE.nPlazo, NULL as nMontoCuota, CE.nTipoGracia,isnull(CE.nPeriodoGracia,0)nPeriodoGracia " 'WIOR 20130223 AGREGO nPeriodoGracia
        '***
        sSql = sSql & " From ColocacEstado CE Inner Join Constante CN ON CE.nPrdEstado = CN.nConsValor AND CN.nConsCod = " & COMDConstantes.gColocEstado
        sSql = sSql & " Where CE.cCtaCod = '" & psCtaCod & "' AND CE.nPrdEstado = " & COMDConstantes.gColocEstSolic
    End If
    If pbSugerencia Then
        If Len(sSql) > 0 Then
            sSql = sSql & " Union "
        End If
        sSql = sSql & "Select CE.nPrdEstado,CN.cConsDescripcion cEstado, CE.nMonto, CE.nCuotas, CE.nPlazo, "
        sSql = sSql & " nMontoCuota = (Select SUM(CD.nMonto) from ColocCalendDet CD where nNrocalen=1 and cCtaCod = CE.cCtaCod AND nColocCalendApl = 1 AND nCuota = 1) "
        'MAVM 20121220 ***
        sSql = sSql & ", CE.nTipoGracia,isnull(CE.nPeriodoGracia,0)nPeriodoGracia " 'WIOR 20130223 AGREGO nPeriodoGracia
        '***
        sSql = sSql & " From ColocacEstado CE Inner Join Constante CN ON CE.nPrdEstado = CN.nConsValor AND CN.nConsCod = " & COMDConstantes.gColocEstado
        'ARCV 30-06-2007
        sSql = sSql & " AND dPrdEstado =(SELECT MAX(dPrdEstado)FROM ColocacEstado WHERE cCtaCod=CE.cCtaCod AND nPrdEstado=CE.nPrdEstado)"
        '--------
        sSql = sSql & " Where CE.cCtaCod = '" & psCtaCod & "' AND  CE.nPrdEstado = " & COMDConstantes.gColocEstSug
    End If
    If pbAprobacion Then
        If Len(sSql) > 0 Then
            sSql = sSql & " Union "
        End If
        sSql = sSql & "Select CE.nPrdEstado, CN.cConsDescripcion cEstado, CE.nMonto, CE.nCuotas, CE.nPlazo, "
        sSql = sSql & " nMontoCuota = (Select SUM(CD.nMonto) from ColocCalendDet CD where nNrocalen=(Select nNroCalen From ColocacCred Where cCtaCod = CE.cCtaCod) and cCtaCod = CE.cCtaCod AND nColocCalendApl = 1 AND nCuota = 1) "
        'MAVM 20121220 ***
        sSql = sSql & ", CE.nTipoGracia,isnull(CE.nPeriodoGracia,0)nPeriodoGracia " 'WIOR 20130223 AGREGO nPeriodoGracia
        '***
        sSql = sSql & " From ColocacEstado CE Inner Join Constante CN ON CE.nPrdEstado = CN.nConsValor AND CN.nConsCod = " & COMDConstantes.gColocEstado
        sSql = sSql & " Where CE.cCtaCod = '" & psCtaCod & "' AND  CE.nPrdEstado = " & COMDConstantes.gColocEstAprob
    End If
    sSql = sSql & " Order By Ce.nprdEstado"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaEstadosdelCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaEstadosdelCredito:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaColocacEstado(ByVal psCtaCod As String, ByVal pnEstado As ColocEstado, Optional pbTodos As Boolean = False) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaColocacEstado
    '***COMMENT BY MARG20171205************************************
'    If Not pbTodos Then
'        sSql = "Select * from ColocacEstado Where cCtaCod = '" & psCtaCod & "' AND nPrdEstado =  " & pnEstado
'    Else
'        sSql = "Select * from ColocacEstado Where cCtaCod = '" & psCtaCod & "' Order By nPrdEstado"
'    End If
    '***END MARG*************************************************
'**MARG20171205***
    sSql = "EXEC stp_sel_RecuperaColocacEstado '" & psCtaCod & "'" & "," & pnEstado & "," & IIf(pbTodos, 1, 0)
'***END MARG*****************************************************
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaColocacEstado = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaColocacEstado:
    Err.Raise Err.Number, "Error En Proceso", Err.Description


End Function

'***MARG20171205***
Public Function GetPeriodoGraciaCalculado(ByVal psCtaCod As String) As Integer
Dim R As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorEstadoActual
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = "exec stp_sel_GetPeriodoGraciaCalculado '" & psCtaCod & "'"
    Set R = oConecta.CargaRecordSet(sSql)
    If Not R.BOF And Not R.EOF Then
        GetPeriodoGraciaCalculado = R!nPeriodoGracia
    Else
        GetPeriodoGraciaCalculado = 0
    End If
    R.Close
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorEstadoActual:
                  Err.Raise Err.Number, "Error al calcular Periodo de Gracia", Err.Description
End Function
'***END MARG*******************

Public Function RecuperaMontoGarantiaCredito(ByVal psCtaCod As String, ByVal pdFecha As Date, Optional ByVal bAmpliado As Boolean) As Double
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim oGeneral As COMDConstSistema.DCOMGeneral
Dim nTipoCambioFijo As Double
Dim nMontoMN As Double
Dim nMontoME As Double
Dim oAmpliado As DCOMAmpliacion
Dim R As ADODB.Recordset
Dim cCtaCodOld As String

    On Error GoTo ErrorRecuperaMontoGarantiaCredito
    Set oGeneral = New COMDConstSistema.DCOMGeneral
    
    'By Capi 03062008 para que jale el tipo de cambio fijo del dia
    'debido a que generaba distorsiones en la conversion de garantias
    'nTipoCambioFijo = oGeneral.EmiteTipoCambio(pdFecha, TCFijoMes)
    nTipoCambioFijo = oGeneral.EmiteTipoCambio(pdFecha, TCFijoDia)
    '
    nTipoCambioFijo = CDbl(Format(nTipoCambioFijo, "#0.0000"))
    Set oGeneral = Nothing

    'Moneda Nacional
    If bAmpliado = True Then
        ' se verifica que si el un credito ampliado o no?
        Set oAmpliado = New DCOMAmpliacion

'            cCtaCodOld = oAmpliado.GetcCtaCodPorAmpliar(psCtaCod)
'            nMontoMN = oAmpliado.EmularGarantia(cCtaCodOld, gMonedaNacional)
'            nMontoMN = Format(nMontoMN, "#0.00")
             nMontoMN = oAmpliado.ObtenerMontoGarantAmpliado(psCtaCod, COMDConstantes.gMonedaNacional)
             Set oAmpliado = Nothing
             nMontoMN = Format(nMontoMN, "#0.00")
    End If


        'sSql = "Select sum(nGravado) as nTotal from ColocGarantia Where cCtaCod = '" & psCtaCod & "' And nMoneda = " & COMDConstantes.gMonedaNacional
        sSql = "EXEC stp_sel_ERS0632014_RecuperaCobertura '" & psCtaCod & "'," & COMDConstantes.gMonedaNacional 'EJVG20150515


        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set R = oConecta.CargaRecordSet(sSql)
        If Not R.BOF And Not R.EOF Then
            nMontoMN = IIf(IsNull(R!nTotal), 0, R!nTotal) + nMontoMN
            nMontoMN = CDbl(Format(nMontoMN, "#0.00"))
        End If
        R.Close
        Set R = Nothing


    'Moneda Extrangera
    If bAmpliado = True Then
        ' se verifica que si el un credito ampliado o no?
        Set oAmpliado = New DCOMAmpliacion
            'cCtaCodOld = oAmpliado.GetcCtaCodPorAmpliar(psCtaCod)
           ' nMontoME = oAmpliado.EmularGarantia(cCtaCodOld, gMonedaExtranjera)
            nMontoME = oAmpliado.ObtenerMontoGarantAmpliado(psCtaCod, gMonedaExtranjera)
            nMontoME = Format(nMontoME, "#0.00")
        Set oAmpliado = Nothing
    End If

        'sSql = "Select sum(nGravado) as nTotal from ColocGarantia Where cCtaCod = '" & psCtaCod & "' And nMoneda = " & gMonedaExtranjera
        sSql = "EXEC stp_sel_ERS0632014_RecuperaCobertura '" & psCtaCod & "'," & COMDConstantes.gMonedaExtranjera 'EJVG20150515
        Set R = oConecta.CargaRecordSet(sSql)
        If Not R.BOF And Not R.EOF Then
            nMontoME = IIf(IsNull(R!nTotal), 0, R!nTotal) + nMontoME
            nMontoME = CDbl(Format(nMontoME, "#0.00"))
        End If
        R.Close
        Set R = Nothing
        oConecta.CierraConexion


    Set oConecta = Nothing
    If CInt(Mid(psCtaCod, 9, 1)) = COMDConstantes.gMonedaNacional Then
        RecuperaMontoGarantiaCredito = 0
        If nMontoMN > 0 Then
            RecuperaMontoGarantiaCredito = RecuperaMontoGarantiaCredito + nMontoMN
        End If
        If nMontoME > 0 Then
            RecuperaMontoGarantiaCredito = RecuperaMontoGarantiaCredito + (nMontoME * nTipoCambioFijo)
        End If
    End If

    If CInt(Mid(psCtaCod, 9, 1)) = COMDConstantes.gMonedaExtranjera Then
        RecuperaMontoGarantiaCredito = 0
        If nMontoMN > 0 Then
            RecuperaMontoGarantiaCredito = RecuperaMontoGarantiaCredito + (nMontoMN / nTipoCambioFijo)
        End If
        If nMontoME > 0 Then
            RecuperaMontoGarantiaCredito = RecuperaMontoGarantiaCredito + nMontoME
        End If
    End If
    
    RecuperaMontoGarantiaCredito = Round(RecuperaMontoGarantiaCredito, 2) 'EJVG20160412
    
    Exit Function

ErrorRecuperaMontoGarantiaCredito:
        Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaDatosPrepago(ByVal psCtaCod As String, ByVal pdHoy As Date) As Double
Dim sSql As String
Dim oConec As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    sSql = " Select SUM(MCD.nMonto) as nMontoPagado from MovCol MC Inner Join Mov M ON MC.nMovNro = M.nMovNro "
    sSql = sSql & " Inner Join MovColDet MCD ON MC.nMovNro = MCD.nMovNro "
    sSql = sSql & " Where MC.nPrepago = 1 AND MC.cCtaCod = '" & psCtaCod & "' AND "
    sSql = sSql & " M.cMovnro like '" & Year(pdHoy) & "%' "
    sSql = sSql & " AND MCD.nPrdConceptoCod = 1000 AND MCD.cOpeCod not in ('" & gCredDesembEfec & "','" & gCredDesembCtaNueva & "','" & gCredDesembCtaExist & "','" & gCredDesembCtaExistDOA & "','" & gCredDesembCtaNuevaDOA & "')"
    sSql = sSql & " AND MC.nNroCalen = (select nNroCalen From ColocacCred Where cCtaCod = MC.cCtaCod)"
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set R = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    RecuperaDatosPrepago = Format(IIf(IsNull(R!nMontoPagado), 0, R!nMontoPagado), "0.00")

End Function

Public Function RecuperaDatosCreditoVigente(ByVal psCtaCod As String, Optional ByVal pdHoy As Date = CDate("01/01/1900"), _
    Optional ByVal pbIncluirbAprobados As Boolean = False) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaDatosCreditoVigente
    'COMENTADO POR JUEZ 20150420 Convertido a SP ****************************************************************
'    'ALPA 20100607 B2**************************************************************
'    'sSql = "Select PP.cPersCod, CC.cMetLiquidacion, dbo.CuotaAprobada(P.cCtaCod) as CuotaAprobada, CC.bCuotaCom, CC.bMiVivienda, CC.nCalendDinamTipo, CC.bPrepago, CC.nCalPago, Pers.cPersNombre, C.cLineaCred,C.nMontoCol, P.nSaldo, "
'    sSql = "Select C.cAgeCodAct,PP.cPersCod, CC.cMetLiquidacion, dbo.CuotaAprobada(P.cCtaCod) as CuotaAprobada, CC.bCuotaCom, CC.bMiVivienda, CC.nCalendDinamTipo, CC.bPrepago, CC.nCalPago, Pers.cPersNombre, C.cLineaCred,C.nMontoCol, P.nSaldo, Pers.nPersPersoneria, " 'JUEZ 20130925 Se agregó Pers.nPersPersoneria
'    '******************************************************************************
'    sSql = sSql & " cMoneda = (Select cConsDescripcion from Constante where nConsCod = " & COMDConstantes.gMoneda & " AND nConsValor = " & CInt(Mid(psCtaCod, 9, 1)) & "),"
'    sSql = sSql & " CE.nCuotas nCuotasApr, CE.nColocCalendCod " & ", "
'    sSql = sSql & " CC.nDiasAtraso, CC.cMetLiquidacion, P.nTransacc, CC.nCalendDinamico, CC.nIntPend, CC.nNroCalen, CC.nNroProxCuota, CC.nNroProxDesemb "
'    If pdHoy <> CDate("01/01/1900") Then
'        sSql = sSql & ", DATEDIFF(year,C.dVigencia,'" & Format(pdHoy, "mm/dd/yyyy") & "') as nPlazoTranscurrido "
'    End If
'    'ALPA 20100607 B2*******************
'    sSql = sSql & ", isnull(cTpoProdCod,0) cTpoProdCod,isnull((select cConsDescripcion from constante where nConsValor=cTpoProdCod and nConsCod=3033),'') cTpoProdDes "
'    sSql = sSql & ", isnull(cTpoCredCod,0) cTpoCredCod,isnull((select cConsDescripcion from constante where nConsValor=cTpoCredCod and nConsCod=3034),'') cTpoCredDes "
'    '***********************************
'    'BRGO 20111227 *********************
'    sSql = sSql & ", CE.dPrdEstado dFecAprobacion "
'    '***********************************
'    sSql = sSql & " From Producto P Inner join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
'    sSql = sSql & "                 Inner Join Persona Pers ON PP.cPersCod = Pers.cPersCod "
'    sSql = sSql & "                 Inner Join Colocaciones C ON C.cCtaCod =  P.cCtaCod "
'    sSql = sSql & "                 Inner Join ColocacCred CC ON CC.cCtaCod =  P.cCtaCod "
'    sSql = sSql & "                 Inner Join ColocacEstado CE ON CE.cCtaCod = P.cCtaCod And CE.nPrdEstado = " & COMDConstantes.gColocEstAprob
'    sSql = sSql & " WHERE P.cCtaCod = '" & psCtaCod & "'"
'
'    If pbIncluirbAprobados Then
'        sSql = sSql & " AND P.nPrdEstado in (" & COMDConstantes.gColocEstRefMor & "," & COMDConstantes.gColocEstRefNorm & "," & COMDConstantes.gColocEstRefVenc & "," & COMDConstantes.gColocEstVigMor & "," & COMDConstantes.gColocEstVigNorm & "," & COMDConstantes.gColocEstVigVenc & "," & COMDConstantes.gColocEstAprob & ")"
'    Else
'        sSql = sSql & " AND P.nPrdEstado in (" & COMDConstantes.gColocEstRefMor & "," & COMDConstantes.gColocEstRefNorm & "," & COMDConstantes.gColocEstRefVenc & "," & COMDConstantes.gColocEstVigMor & "," & COMDConstantes.gColocEstVigNorm & "," & COMDConstantes.gColocEstVigVenc & ")"
'    End If
     
    sSql = "EXEC stp_sel_RecuperaDatosCreditoVigente '" & psCtaCod & "','" & Format(pdHoy, "yyyyMMdd") & "'," & IIf(pbIncluirbAprobados, 1, 0)
    'END JUEZ ***************************************************************************************************

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosCreditoVigente = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaDatosCreditoVigente:
    Err.Raise Err.Number, "Error En Proceso RecuperaDatosCreditoVigente", Err.Description

End Function

Public Sub ActualizaCreditoAnalista(ByVal psCtaCod As String, ByVal psPerscodActual As String, _
                                    Optional ByVal psPersCodAnterior As String = "", Optional ByVal pdFecSis As Date = 0, _
                                    Optional ByVal pnSalCap As Double = 0, Optional ByVal psTpoCred As String = "", Optional ByVal pnDiasAtraso As Integer = 0, _
                                    Optional ByVal pnMora830 As Double = 0, Optional ByVal pnMoraM30 As Double = 0, Optional ByVal pnMora815 As Double = 0, _
                                    Optional ByVal pnMoraM15 As Double = 0, _
                                    Optional ByVal psAgeCod As String = "") 'WIOR 20131119 AGREGÓ pnSalCap, psTpoCred, pnDiasAtraso, pnMora830, pnMoraM30, pnMora815 y pnMoraM15
                                    'JUEZ 20160425 Se agregó psAgeCod
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim oFechas As COMDConstSistema.DCOMGeneral

    On Error GoTo ErrorActualizaCreditoAnalista
    Set oConecta = New COMConecta.DCOMConecta
'    sSql = "UPDATE ProductoPersona SET cPersCod = '" & psPerscodActual & "' "
'    sSql = sSql & " Where nPrdPersRelac = " & Trim(Str(COMDConstantes.gColRelPersAnalista)) & " AND cCtaCod = '" & psCtaCod & "'"
     sSql = "EXEC STP_UPD_ACTUALIZACREDITOANALISTA '" & psPerscodActual & "','" & psCtaCod & "'" 'APRI20170529 ERS026-2017
       
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    If psPersCodAnterior <> "" Then
'       COMENTADO POR ARLO20180414 TIC1804110003
'        Set oFechas = New COMDConstSistema.DCOMGeneral
'        'sSql = "INSERT INTO ColocacReasignaCartera(cCtaCod,dReasignacion,cPersCodAnterior,cPersCodActual)" & _
'               " VALUES ('" & psCtaCod & "','" & Format(oFechas.FechaHora(pdFecSis), "mm/dd/yyyy hh:mm:ss") & "','" & psPersCodAnterior & "','" & psPerscodActual & "')" 'WIOR 20131118 COMENTÓ
'        'WIOR 20131118 ***********************************************
'        sSql = "INSERT INTO ColocacReasignaCartera(cCtaCod,dReasignacion,cPersCodAnterior,cPersCodActual,nSalCap,cTpoCredCod,nDiasAtraso,nMora830,nMoraM30,nMora815,nMoraM15,cAgeCodAnt,cAgeCodAct)" & _
'               " VALUES ('" & psCtaCod & "','" & Format(pdFecSis, "yyyymmdd hh:mm:ss") & "','" & psPersCodAnterior & "','" & psPerscodActual & "'," & pnSalCap & ",'" & psTpoCred & "'," & pnDiasAtraso & "," & pnMora830 & "," & pnMoraM30 & "," & pnMora815 & "," & pnMoraM15 & ",'" & psAgeCod & "','" & psAgeCod & "')"
'        'WIOR FIN ****************************************************
'       Set oFechas = Nothing
'       COMENTADO POR ARLO20180414 TIC1804110003
        sSql = "EXEC stp_ins_ReasignaCarteraAnalista '" & psCtaCod & "','" & psPersCodAnterior & "','" & psPerscodActual & "'," & pnSalCap & ",'" & psTpoCred & "'," & pnDiasAtraso & "," & pnMora830 & "," & pnMoraM30 & "," & pnMora815 & "," & pnMoraM15 & ",'" & psAgeCod & "'" 'ARLO 20181404 TIC1804110003

        oConecta.ConexionActiva.Execute sSql
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorActualizaCreditoAnalista:
    Set oConecta = Nothing
    Err.Raise Err.Number, "Error En Proceso ActualizaCreditoAnalista", Err.Description


End Sub

Public Function CarteraInstitucion(ByVal psPersCod As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorCarteraInstitucion
    sSql = "Select C.cCtaCod, P.cPersNombre from ColocacConvenio C inner join ProductoPersona PP ON C.cCtaCod = PP.cCtaCod and PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
    sSql = sSql & "             Inner Join Persona P ON PP.cPersCod = P.cPersCod"
    sSql = sSql & " where C.cPersCod = '" & Trim(psPersCod) & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CarteraInstitucion = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorCarteraInstitucion:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function CarteraAnalista(ByVal psPersCod As String, Optional ByVal pCodAgen As String, _
Optional ByVal psCodUbiGeografica As String = "", Optional ByVal psCodTipoProducto As String = "") As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorCarteraAnalista

    Set oConecta = New COMConecta.DCOMConecta
    'WIOR 20131118 comentó para agregar en un procedimiento
'    sSql = "Select C.cCtaCod, PER.cPersNombre "
'    sSql = sSql & " from    ColocacCred C inner join ProductoPersona PP ON C.cCtaCod = PP.cCtaCod "
'    sSql = sSql & " inner join Producto PRD ON C.cCtaCod = PRD.cCtaCod "
'    sSql = sSql & " inner join ProductoPersona PP2 ON C.cCtaCod = PP2.cCtaCod "
'    sSql = sSql & " inner join Persona PER ON PP2.cPersCod = PER.cPersCod "
'
'    If Len(psCodUbiGeografica) > 0 Then
'        If Mid(psCodUbiGeografica, 1, 1) = "3" Then
'            sSql = sSql & " Inner Join UbicacionGeoGrafica U on SubString(U.cUbiGeoCod,1,1)='3' and SubString(U.cUbiGeoCod,2,6)=SubString(Per.cPersDireccUbiGeo,2,6) and Substring(U.cUbiGeoCod,8,5)='00000'"
'        Else
'            sSql = sSql & "Inner Join UbicacionGeoGrafica U on U.cUbiGeoCod=Per.cPersDireccUbiGeo"
'        End If
'    End If
'
'    sSql = sSql & " Where PP.nPrdPersRelac = 28 and PRD.nPrdEstado IN('2020','2021','2022','2030','2031','2032','2201','2205') and PP2.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular & " and PP.cPersCod = '" & psPersCod & "'"
'    sSql = sSql & " and substring(prd.cctacod,4,2)='" & pCodAgen & "'"
'
'    If Len(psCodUbiGeografica) > 0 Then
'        sSql = sSql & " and U.cUbiGeoCod='" & psCodUbiGeografica & "'"
'    End If
'
'    If Len(psCodTipoProducto) > 0 Then
'       sSql = sSql & " and SubString(Prd.cCtaCod,6,3)='" & psCodTipoProducto & "'"
'    End If
    'WIOR 20131118 *****************************
    sSql = "EXEC stp_sel_ObtenerCarteraAnalista '" & Trim(psPersCod) & "','" & Trim(pCodAgen) & "','" & Trim(psCodUbiGeografica) & "','" & Trim(psCodTipoProducto) & "'"
    'WIOR FIN **********************************
    
    oConecta.AbreConexion
    Set CarteraAnalista = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorCarteraAnalista:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    Set oConecta = Nothing

End Function

Public Sub EliminaSolicitud(ByVal psCtaCod As String)
Dim sSql As String
Dim Conn As COMConecta.DCOMConecta
Dim bTran As Boolean

    On Error GoTo ErrorEliminaSolicitud
    bTran = False
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    Conn.ConexionActiva.BeginTrans
    bTran = True

    'Elimina ProductoPersona
    sSql = "Delete ProductoPersona Where cCtaCod = '" & psCtaCod & "'"
    Conn.ConexionActiva.Execute sSql

    'Elimina ColocEstado
    sSql = "Delete ColocacEstado Where cCtaCod = '" & psCtaCod & "'"
    Conn.ConexionActiva.Execute sSql

    'Elimina ColocFteIngreso
    sSql = "Delete ColocFteIngreso Where cCtaCod = '" & psCtaCod & "'"
    Conn.ConexionActiva.Execute sSql

    'Elimina ColocacConvenio
    sSql = "Delete ColocacConvenio Where cCtaCod = '" & psCtaCod & "'"
    Conn.ConexionActiva.Execute sSql

    'Elimina ColocacCred
    sSql = "Delete ColocacCred Where cCtaCod = '" & psCtaCod & "'"
    Conn.ConexionActiva.Execute sSql

    'Elimina Colocaciones
    sSql = "Delete Colocaciones Where cCtaCod = '" & psCtaCod & "'"
    Conn.ConexionActiva.Execute sSql

    'Elimina Producto
    sSql = "Delete Producto Where cCtaCod = '" & psCtaCod & "'"
    Conn.ConexionActiva.Execute sSql

    'ELimina si es credito ampliado
    sSql = "Delete ColocacAmpliado Where cCtaCod='" & psCtaCod & "'"
    Conn.ConexionActiva.Execute sSql

    Conn.ConexionActiva.CommitTrans
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Sub

ErrorEliminaSolicitud:
    If bTran Then
        Conn.RollbackTrans
        Set Conn = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso EliminaSolicitud", Err.Description
End Sub

'ARCV 27-10-2006
' psCargo As String
' psCARBEN As String
' psT_Plani As String
'Se agregaron para los Convenios

Public Sub ActualizaSolicitud(ByVal oRelPersCred As Variant, ByVal psCtaCod As String, _
                          ByVal psCondCred As String, ByVal psCondCredOtra As String, ByVal pnMontoSol As Double, _
                          ByVal pnCuotasSol As Integer, ByVal pnPlazo As Integer, ByVal psDestCred As String, _
                          ByVal psPerscodAnalista As String, ByVal psPersCodInst As String, ByVal psCodModular As String, _
                          ByVal dFecSol As Date, ByVal psMovAct As String, ByVal bRefinanciar As Boolean, ByVal pnCondProd As Integer, Optional ByVal MatCredRef As Variant, Optional pnCapitalInt As Boolean = False, Optional bSustiDeudor As Boolean = False, Optional ByVal nCampanaCod As Integer, _
                          Optional ByVal pdFecSis As Date = 0, Optional ByVal psCodUser As String = "", Optional ByVal psCodCmac As String = "", Optional ByVal psCodAge As String = "", Optional ByVal psCadenaCta As String = "", _
                          Optional ByVal psCargo As String = "", Optional ByVal psCARBEN As String = "", Optional ByVal psT_Plani As String = "", _
                          Optional ByVal pnMotivoRef As Integer, Optional ByVal psTpoProdCod As Producto, Optional ByVal nCasaCod As Integer = 0, Optional ByVal psVendedor As String = "", Optional ByVal psMotivoDetRef As String = "", Optional ByVal pbGrabaPromotor As Boolean = False, Optional ByVal psCodPromotor As String = "", Optional ByRef pbEliminaCobertura As Boolean = False, _
                          Optional ByRef pvListaCompraDeuda As Variant, Optional ByVal pnMontoExpEsteCred As Double = 0#, Optional ByVal pbEliminarEvaluacion As Boolean = False, Optional ByRef pvListaAguasaneamiento As Variant, _
                          Optional ByVal pnSubDestino As Long = 0, Optional ByVal pnTpDoc As Long = 0, Optional ByVal pnTpIngr As Long = 0, Optional ByVal pnTpInt As Long = 0, Optional ByRef pvListaCreditoVerde As Variant, _
                          Optional ByRef pvTablasRSE As Variant) 'EJVG20160203 ERS002-2016 'EAAS20180912 SEGUN ERS-054-2018 pvListaCompraDeuda 'EAAS20191401 pvListaCreditoVerde SEGUN 018-GM-DI_CMACM
'By capi 28102008 se agrego parametro pnMotivoRef
'MIOL 20130626 RQ133335 se agrego parametro psVendedor
'MADM 20100719 - Optional ByVal nCasaCod as Integer
'WIOR 20140509 AGREGO  pbGrabaPromotor,psCodPromotor
'EJVG20160712 Agregó pnMontoExpEsteCred y pbEliminarEvaluacion
'Public Sub ActualizaSolicitud(ByVal oRelPersCred As Variant, ByVal psCtaCod As String, _
'                          ByVal psCondCred As String, ByVal psCondCredOtra As String, ByVal pMatFuentes As Variant, ByVal pnMontoSol As Double, _
'                          ByVal pnCuotasSol As Integer, ByVal pnPlazo As Integer, ByVal psDestCred As String, _
'                          ByVal psPerscodAnalista As String, ByVal psPersCodInst As String, ByVal psCodModular As String, _
'                          ByVal dFecSol As Date, ByVal psMovAct As String, ByVal bRefinanciar As Boolean, ByVal pnCondProd As Integer, Optional ByVal MatCredRef As Variant, Optional pnCapitalInt As Boolean = False, Optional bSustiDeudor As Boolean = False, Optional ByVal nCampanaCod As Integer, _
'                          Optional ByVal pdFecSis As Date = 0, Optional ByVal psCodUser As String = "", Optional ByVal psCodCmac As String = "", Optional ByVal psCodAge As String = "", Optional ByVal psCadenaCta As String = "", _
'                          Optional ByVal psCargo As String = "", Optional ByVal psCARBEN As String = "", Optional ByVal psT_Plani As String = "")

'Agrego JOEP20190123 CP Optional ByVal pnSubDestino As Long = 0, Optional ByVal pnTpDoc As Long = 0, Optional ByVal pnTpIngr As Long = 0, Optional ByVal pnTpInt As Long = 0

Dim bTran As Boolean
Dim i As Integer
Dim sSql As String
Dim Conn As COMConecta.DCOMConecta
Dim oBase As DCOMCredActBD
Dim lnInteresDiferido As Currency 'ALPA 20150826********
Dim lnRefinanciadoDeVigenteCancelado As Currency 'ALPA 20150909
Dim oRs As ADODB.Recordset 'ALPA 20150909
Dim lvListaCompraDeuda() As TCompraDeuda 'EJVG20160203 ERS002-2016
Dim lvListaAguaSaneamiento() As TAguaSaneamiento 'EAAS20180815 ERS054-2018
Dim lvListaCreditoVerde() As TCreditoVerde 'EAAS20180815 SEGUN 018-GM-DI_CMACM
Dim ixCD As Integer 'EJVG20160203 ERS002-2016
Dim lsPersTitular As String 'ALPA20160702
'Dim oGen As COMDConstSistema.DCOMGeneral

    On Error GoTo ErrorActualizaSolicitud

    bTran = False
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    
    'If psCtaCod = "" Then
    '    Set oGen = New COMDConstSistema.DCOMGeneral
    '    psCtaCod = gsCodCMAC & oGen.GeneraNuevaCuenta(psCodAge, pProducto, pnMoneda)
    '    Set oGen = Nothing
    'End If
    
    psMovAct = GeneraMovNroActualiza(pdFecSis, psCodUser, psCodCmac, psCodAge)
    
    
    Conn.ConexionActiva.BeginTrans
    bTran = True

    'Actualiza Producto
    sSql = "UPDATE Producto SET "
    sSql = sSql & " nPrdEstado = " & Trim(Str(gColocEstSolic)) & ", "
    sSql = sSql & " dPrdEstado = '" & Format(dFecSol, "mm/dd/yyyy") & "'"
    sSql = sSql & " WHERE cCtaCod = '" & psCtaCod & "'"
    Conn.ConexionActiva.Execute sSql

    'Actualiza Nuevas Personas Relacionadas con el Credito
'    oRelPersCred.IniciarMatriz
'
'    Do While Not oRelPersCred.EOF
'            sSql = "DELETE ProductoPersona "
'            sSql = sSql & " WHERE cCtaCod = '" & psCtaCod & "' AND cPersCod = '" & oRelPersCred.ObtenerCodigo & "' AND nPrdPersRelac = " & oRelPersCred.ObtenerValorRelacAnt
'            Conn.ConexionActiva.Execute sSql
'
'            sSql = "INSERT INTO ProductoPersona(cCtaCod,cPersCod,nPrdPersRelac) "
'            sSql = sSql & " VALUES('" & psCtaCod & "','" & oRelPersCred.ObtenerCodigo & "'," & oRelPersCred.ObtenerValorRelac & ")"
'            Conn.ConexionActiva.Execute sSql
'
'        oRelPersCred.siguiente
'    Loop
    For i = 0 To UBound(oRelPersCred) - 1
        'FRHU 20160712 INCIDENTE
        'sSql = "DELETE ProductoPersona "
        'sSql = sSql & " WHERE cCtaCod = '" & psCtaCod & "' AND cPersCod = '" & oRelPersCred(i, 0) & "' AND nPrdPersRelac = " & oRelPersCred(i, 2)
        'Conn.ConexionActiva.Execute sSql
        If i = 0 Then
            sSql = "DELETE ProductoPersona WHERE cCtaCod = '" & psCtaCod & "' And nPrdPersRelac <> " & gColRelPersAnalista
            Conn.ConexionActiva.Execute sSql
        End If
        'FIN FRHU 20160712
        sSql = "INSERT INTO ProductoPersona(cCtaCod,cPersCod,nPrdPersRelac) "
        sSql = sSql & " VALUES('" & psCtaCod & "','" & oRelPersCred(i, 0) & "'," & oRelPersCred(i, 1) & ")"
        Conn.ConexionActiva.Execute sSql
        
        'ALPA20160702*****************************
        If oRelPersCred(i, 1) = 20 Then
            lsPersTitular = oRelPersCred(i, 0)
        End If
        '****************************************
    Next
    
    'Actualiza Analista
    sSql = "UPDATE ProductoPersona SET "
    sSql = sSql & " cPersCod = '" & psPerscodAnalista & "' "
    sSql = sSql & " WHERE cCtaCod = '" & psCtaCod & "' AND nPrdPersRelac = " & gColRelPersAnalista

    Conn.ConexionActiva.Execute sSql
    
    'WIOR 20140509(PROMOTORES) **************
    If pbGrabaPromotor Then
        sSql = "EXEC stp_ins_upd_CredPromotores '" & psCtaCod & "','" & psCodPromotor & "'"
        Conn.ConexionActiva.Execute sSql
    End If
    'WIOR FIN *******************************
    
    'Actualiza Registro de Colocaciones
    sSql = "UPDATE Colocaciones SET "
    sSql = sSql & " nPlazo = " & Trim(Str(pnPlazo)) & ","
    sSql = sSql & " nMontoCol = " & Format(pnMontoSol, "#0.00") & ","
    'ALPA 20100604 B2************************************************
    'sSql = sSql & " cUltimaActualizacion = '" & psMovAct & "' "
    sSql = sSql & " cUltimaActualizacion = '" & psMovAct & "', "
    sSql = sSql & " cTpoProdCod = '" & psTpoProdCod & "', "
    'CTI3 ERS0032020
    sSql = sSql & " nTpIngr = " & IIf(pnTpIngr = 0, "NULL", pnTpIngr) & " "
    '****************************************************************
    sSql = sSql & " WHERE cCtaCod = '" & psCtaCod & "'"
    Conn.ConexionActiva.Execute sSql

    'Actualiza Registro de ColocacCred pnCondProd
    sSql = "UPDATE ColocacCred SET "
    sSql = sSql & " nColocCondicion = " & psCondCred & ", "
    sSql = sSql & " nColocCondicionProd = " & pnCondProd & ", "
    sSql = sSql & " nColocDestino = '" & psDestCred & "', "
    sSql = sSql & " bRefCapInt = " & IIf(pnCapitalInt, "1 ", "0 ") & ","
    'sSql = sSql & " IdCampana=" & IIf(nCampanaCod = 0, "NULL", nCampanaCod) & "," 'Comento JOEP20180222
    sSql = sSql & " IdCampana=" & nCampanaCod & " ," 'AGREGO JOEP20180222
    sSql = sSql & " id_CasaCom=" & IIf(nCasaCod = 0, "NULL", nCasaCod) & ","
    sSql = sSql & " cVendedor = '" & psVendedor & "'," 'MIOL 20130626 RQ13335
    
    'JOEP20190919 ERS042 CP-2018
    sSql = sSql & " nSubDestino = " & IIf(pnSubDestino = 0, "NULL", pnSubDestino) & ","
    sSql = sSql & " nTpDoc = " & IIf(pnTpDoc = 0, "NULL", pnTpDoc) & ","
    sSql = sSql & " nTpIngr = " & IIf(pnTpIngr = 0, "NULL", pnTpIngr) & ","
    sSql = sSql & " nTpInt = " & IIf(pnTpInt = 0, "NULL", pnTpInt) & ""
    'JOEP20190919 ERS042 CP-2018
    
    sSql = sSql & " WHERE cCtaCod = '" & psCtaCod & "'"

    Conn.ConexionActiva.Execute sSql

    'Actualiza Registro en ColocEstado
    Set oBase = New DCOMCredActBD
    sSql = "UPDATE ColocacEstado SET "
    sSql = sSql & " dPrdEstado = '" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
    sSql = sSql & " nCuotas = " & Trim(Str(pnCuotasSol)) & ","
    sSql = sSql & " nMonto = " & Format(pnMontoSol, "#0.00") & ","
    sSql = sSql & " nPlazo = " & Trim(Str(pnPlazo))
    sSql = sSql & " WHERE cCtaCod = '" & psCtaCod & "' AND nPrdEstado = " & Trim(Str(gColocEstSolic))
    Set oBase = Nothing
    Conn.ConexionActiva.Execute sSql

    'Actualiza Registro en ColocFteIngreso
    'ARCV 30-12-2006
    'sSQL = "UPDATE ColocFteIngreso SET "
    'sSQL = sSQL & " cNumFuente = '" & psFuenteIng & "', "
    'sSQL = sSQL & " dPersEval = '" & Format(pdFIFecEval, "mm/dd/yyyy") & "' "
    'sSQL = sSQL & " WHERE cCtaCod = '" & psCtaCod & "' AND cNumFuente = '" & psFuenteIng & "'"
    'Conn.ConexionActiva.Execute sSQL


'*** COMENTADO POR PEAC 20080412 SEGUN ACTA 009-2008/TI-D
'*** ESTE PROCESO PASA A SUGERENCIA

'    sSql = "DELETE FROM ColocFteIngreso WHERE cCtaCod='" & psCtaCod & "'"
'    Conn.ConexionActiva.Execute sSql
'    For i = 0 To UBound(pMatFuentes) - 1
'        sSql = "INSERT INTO ColocFteIngreso(cNumFuente,dPersEval,cCtaCod)"
'            sSql = sSql & " VALUES('" & pMatFuentes(i, 0) & "',"
'            sSql = sSql & "'" & Format(pMatFuentes(i, 1), "mm/dd/yyyy") & "', "
'            sSql = sSql & "'" & psCtaCod & "')"
'            Conn.ConexionActiva.Execute sSql
'    Next i

    '-----------------
    'Nuevo Registro de ColocConvenio
    If Len(Trim(psPersCodInst)) > 0 Then
        sSql = "DELETE ColocacConvenio WHERE cCtaCod = '" & psCtaCod & "' AND cPersCod = '" & psPersCodInst & "'"
        Conn.ConexionActiva.Execute sSql

        sSql = "INSERT INTO ColocacConvenio(cCtaCod,cPersCod,cCodModular"
        sSql = sSql & ",cCargo,cCARBEN,cT_Plani )" 'ARCV 27-10-2006
        sSql = sSql & " VALUES('" & psCtaCod & "',"
        sSql = sSql & "'" & psPersCodInst & "',"
        sSql = sSql & "'" & psCodModular & "'"
        'ARCV 27-10-2006
        sSql = sSql & ",'" & psCargo & "','" & psCARBEN & "','" & psT_Plani & "')"
        '----------------
        Conn.ConexionActiva.Execute sSql
    End If

    'Si es una Refinanciancion
    'Si es una Refinanciacion Guardar los Datos
    If bRefinanciar Then
        If IsArray(MatCredRef) Then
            Set oBase = New DCOMCredActBD
            If UBound(MatCredRef) > 0 Then
                sSql = "Delete ColocacRefinancDet where cCtaCod = '" & psCtaCod & "'"
                Conn.ConexionActiva.Execute sSql
                sSql = "Delete ColocacRefinanc where cCtaCod = '" & psCtaCod & "'"
                Conn.ConexionActiva.Execute sSql
                For i = 0 To UBound(MatCredRef) - 1
                    'By Capi 28102008 se agrego codigo para que grabe motivo refinanciacion
                    
'                    sSql = "INSERT INTO ColocacRefinanc(cCtaCod,cCtaCodRef,nEstado,dEstado,nMontoRef)"
'                    sSql = sSql & " VALUES('" & psCtaCod & "','"
'                    sSql = sSql & MatCredRef(i, 0) & "',"
'                    sSql = sSql & gColocEstadoRefinancSolicitado & ","
'                    sSql = sSql & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
'                    sSql = sSql & MatCredRef(i, 1) & ")"
'                    Conn.ConexionActiva.Execute sSql
'
                    sSql = "INSERT INTO ColocacRefinanc(cCtaCod,cCtaCodRef,nEstado,dEstado,nMontoRef,nMotivoRef,sMotivoDetRef)"
                    sSql = sSql & " VALUES('" & psCtaCod & "','"
                    sSql = sSql & MatCredRef(i, 0) & "',"
                    sSql = sSql & gColocEstadoRefinancSolicitado & ","
                    sSql = sSql & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
                    sSql = sSql & MatCredRef(i, 1) & "," & pnMotivoRef
                    sSql = sSql & ",'" & psMotivoDetRef & "')" 'JAME20140509
                    Conn.ConexionActiva.Execute sSql
                    
                    'EndBy
                    
                    'Ingreso del Detalle
                    'Capital
                    sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                    sSql = sSql & " VALUES('" & psCtaCod & "','"
                    sSql = sSql & MatCredRef(i, 0) & "',"
                    sSql = sSql & gColocEstadoRefinancSolicitado & ","
                    sSql = sSql & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
                    sSql = sSql & gColocConceptoCodCapital & ","
                    sSql = sSql & MatCredRef(i, 2) & ","
                    sSql = sSql & "0.00)"
                    Conn.ConexionActiva.Execute sSql
                    
                    If pnCapitalInt Then 'MAVM 20110117
                        'Interes Compensatorio
                        If CDbl(MatCredRef(i, 3)) > 0 Then
                            sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                            sSql = sSql & " VALUES('" & psCtaCod & "','"
                            sSql = sSql & MatCredRef(i, 0) & "',"
                            sSql = sSql & gColocEstadoRefinancSolicitado & ","
                            sSql = sSql & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
                            sSql = sSql & gColocConceptoCodInteresCompensatorio & ","
                            sSql = sSql & MatCredRef(i, 3) & ","
                            sSql = sSql & "0.00)"
                            Conn.ConexionActiva.Execute sSql
                        End If
    
                        'Interes Moratorio
                        If CDbl(MatCredRef(i, 4)) > 0 Then
                            sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                            sSql = sSql & " VALUES('" & psCtaCod & "','"
                            sSql = sSql & MatCredRef(i, 0) & "',"
                            sSql = sSql & gColocEstadoRefinancSolicitado & ","
                            sSql = sSql & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
                            sSql = sSql & gColocConceptoCodInteresMoratorio & ","
                            sSql = sSql & MatCredRef(i, 4) & ","
                            sSql = sSql & "0.00)"
                            Conn.ConexionActiva.Execute sSql
                        End If
    
                        'Interes Gracia
                        If CDbl(MatCredRef(i, 5)) > 0 Then
                            sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                            sSql = sSql & " VALUES('" & psCtaCod & "','"
                            sSql = sSql & MatCredRef(i, 0) & "',"
                            sSql = sSql & gColocEstadoRefinancSolicitado & ","
                            sSql = sSql & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
                            sSql = sSql & gColocConceptoCodInteresGracia & ","
                            sSql = sSql & MatCredRef(i, 5) & ","
                            sSql = sSql & "0.00)"
                            Conn.ConexionActiva.Execute sSql
                        End If
    
                        'Interes Suspenso
                        If CDbl(MatCredRef(i, 6)) > 0 Then
                            sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                            sSql = sSql & " VALUES('" & psCtaCod & "','"
                            sSql = sSql & MatCredRef(i, 0) & "',"
                            sSql = sSql & gColocEstadoRefinancSolicitado & ","
                            sSql = sSql & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
                            sSql = sSql & gColocConceptoCodInteresSuspenso & ","
                            sSql = sSql & MatCredRef(i, 6) & ","
                            sSql = sSql & "0.00)"
                            Conn.ConexionActiva.Execute sSql
                        End If
    
                        'Interes Reprogramado
                        If CDbl(MatCredRef(i, 7)) > 0 Then
                            sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                            sSql = sSql & " VALUES('" & psCtaCod & "','"
                            sSql = sSql & MatCredRef(i, 0) & "',"
                            sSql = sSql & gColocEstadoRefinancSolicitado & ","
                            sSql = sSql & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
                            sSql = sSql & gColocConceptoCodInteresReprogramado & ","
                            sSql = sSql & MatCredRef(i, 7) & ","
                            sSql = sSql & "0.00)"
                            Conn.ConexionActiva.Execute sSql
                        End If
    
                        'Gastos
                        If CDbl(MatCredRef(i, 8)) > 0 Then
                            sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                            sSql = sSql & " VALUES('" & psCtaCod & "','"
                            sSql = sSql & MatCredRef(i, 0) & "',"
                            sSql = sSql & gColocEstadoRefinancSolicitado & ","
                            sSql = sSql & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
                            sSql = sSql & gColocConceptoCodGastoVarios & ","
                            sSql = sSql & MatCredRef(i, 8) & ","
                            sSql = sSql & "0.00)"
                            Conn.ConexionActiva.Execute sSql
                        End If
                    End If
                    'ALPA 20150803
                    lnInteresDiferido = lnInteresDiferido + MatCredRef(i, 7) + MatCredRef(i, 6) + MatCredRef(i, 5) + MatCredRef(i, 4) + MatCredRef(i, 3)
                    Set oRs = New ADODB.Recordset
                    sSql = "Exec stp_sel_obtenerDiferidoCreditoRefinanciadoDeVigenteCancelado '" & MatCredRef(i, 0) & "','" & Format(dFecSol, "YYYY/MM/DD") & "'," & Mid(psCtaCod, 9, 1)
                    Set oRs = Conn.ConexionActiva.Execute(sSql)
                    If Not (oRs.BOF Or oRs.EOF) Then
                        lnRefinanciadoDeVigenteCancelado = oRs!nMontoDiferidoCredCancelado
                    End If
                    lnInteresDiferido = lnInteresDiferido + lnRefinanciadoDeVigenteCancelado
                    Set oRs = Nothing
                    '*******************************
                Next i
                'ALPA 20150803**********************
                sSql = "Update ColocacCred "
                sSql = sSql & " Set nInteresDiferido=" & lnInteresDiferido
                sSql = sSql & " Where cCtaCod='" & psCtaCod & "'"
                Conn.ConexionActiva.Execute sSql
                '***********************************
            End If
            Set oBase = Nothing
        End If
    End If

    ' un credito automatico
'        oRelPersCred.IniciarMatriz
'        Do While Not oRelPersCred.EOF
'                If oRelPersCred.ObtenerValorRelac = 20 Then ' si es titular
'                    If psCondCredOtra = "3" Then
'                        sSql = "Update Persona set bPreferencial=1 Where cPersCod='" & oRelPersCred.ObtenerCodigo & "'"
'                    Else
'                        sSql = "Update Persona set bPreferencial=0 Where cPersCod='" & oRelPersCred.ObtenerCodigo & "'"
'                    End If
'                    Conn.ConexionActiva.Execute sSql
'                End If
'            oRelPersCred.siguiente
'        Loop
        For i = 0 To UBound(oRelPersCred) - 1
            If oRelPersCred(i, 1) = 20 Then ' si es titular
                If psCondCredOtra = "3" Then
                    sSql = "Update Persona set bPreferencial=1 Where cPersCod='" & oRelPersCred(i, 0) & "'"
                Else
                    sSql = "Update Persona set bPreferencial=0 Where cPersCod='" & oRelPersCred(i, 0) & "'"
                End If
                Conn.ConexionActiva.Execute sSql
            End If
        Next

     'AVMM-27-09-2006
    'EJVG20151015 ***
    'If bRefinanciar = True Then
    '   'Buscar Garantias e Inserta
    '    If psCadenaCta <> "" Then
    '        sSql = "DELETE ColocGarantia where cCtaCod ='" & psCtaCod & "'"
    '        Conn.ConexionActiva.Execute (sSql)
    '
    '        sSql = " INSERT INTO ColocGarantia" & _
    '               " SELECT cNumGarant,'" & psCtaCod & "'  ,nMoneda,nGravado,nEstado FROM dbo.ColocGarantia where cCtaCod in " & psCadenaCta
    '        Conn.ConexionActiva.Execute (sSql)
    '    End If
    'End If

    'ARCV 12-02-2007
    
    '---------------
    'EJVG20151015 *** Se va a eliminar porque la actualización afecta el [Registro de Cobertura]-> El Analista debe volver a registrar cobertura
    If Conn.CargaRecordSet("SELECT CG.cNumGarant FROM ColocGarantia CG WHERE cCtaCod = '" & psCtaCod & "'").RecordCount > 0 Then
        sSql = "EXEC stp_del_ERS0632014_ColocGarantia '" & psCtaCod & "'"
        Conn.ConexionActiva.Execute sSql
        pbEliminaCobertura = True
    End If
    'END EJVG *******
     'EJVG20160203 *** Compra Deuda otras IFIs
    If val(psDestCred) = ColocDestino.gColocDestinoCambEstructPasivo Then
        lvListaCompraDeuda = pvListaCompraDeuda
    Else
        ReDim lvListaCompraDeuda(0)
    End If
    
    sSql = "EXEC stp_del_ERS0022016_ColocCompraDeuda '" & psCtaCod & "'"
    Conn.ConexionActiva.Execute (sSql)
    
    For ixCD = 1 To UBound(lvListaCompraDeuda)
        sSql = "EXEC stp_ins_ERS0022016_ColocCompraDeuda '" & psCtaCod & "'," & ixCD & ",'" & lvListaCompraDeuda(ixCD).sIFICod & "'," & lvListaCompraDeuda(ixCD).nMoneda & ",'" & lvListaCompraDeuda(ixCD).sNroCredito & "'," & lvListaCompraDeuda(ixCD).nDestino & "," & IIf(lvListaCompraDeuda(ixCD).bRecompra, 1, 0) & ",'" & Format(lvListaCompraDeuda(ixCD).dDesembolso, "yyyymmdd") & "'," & lvListaCompraDeuda(ixCD).nMontoDesembolso & "," & lvListaCompraDeuda(ixCD).nNroCuotasPactadas & "," & lvListaCompraDeuda(ixCD).nNroCuotasPagadas & "," & lvListaCompraDeuda(ixCD).nSaldoComprar & "," & lvListaCompraDeuda(ixCD).nMontoCuota
        Conn.ConexionActiva.Execute (sSql)
    Next
    'END EJVG *******
    'INICIO EAAS20180815 SEGUN ERS-054-2018***

    If UBound(pvListaAguasaneamiento) > 0 Then
        lvListaAguaSaneamiento = pvListaAguasaneamiento
    Else
        ReDim lvListaAguaSaneamiento(0)
    End If
    
    sSql = "EXEC stp_del_ERS054_2018_ColocDestinoAguaSaneamiento '" & psCtaCod & "'"
    Conn.ConexionActiva.Execute (sSql)
    Dim nSumaTotalAguaSaneamiento As Double
    Dim nSuma As Double
    nSuma = 0
    Dim nSumaTotalBeneficiados As Integer
    Dim nSumaB As Integer
    nSumaB = 0
    
    For ixCD = 1 To UBound(lvListaAguaSaneamiento)
        nSumaTotalAguaSaneamiento = nSumaTotalAguaSaneamiento + lvListaAguaSaneamiento(ixCD).nMontoS
        nSumaTotalBeneficiados = nSumaTotalBeneficiados + lvListaAguaSaneamiento(ixCD).nBeneficia
    Next
    
    If (UBound(lvListaAguaSaneamiento) > 0) Then
        For ixCD = 1 To 1
            sSql = "EXEC stp_ins_ERS054_2018_ColocDestinoAguaSaneamiento '" & psCtaCod & "'," & ixCD & "," & lvListaAguaSaneamiento(ixCD).nDestinoCod & "," & pnMontoSol & "," & nSumaTotalAguaSaneamiento & ", " & lvListaAguaSaneamiento(ixCD).nBeneficia
            Conn.ConexionActiva.Execute (sSql)
        Next
    End If
    For ixCD = 1 To UBound(lvListaAguaSaneamiento)
        sSql = "EXEC stp_ins_ERS054_2018_ColocDestinoAguaSaneamientoDet '" & psCtaCod & "'," & ixCD & "," & lvListaAguaSaneamiento(ixCD).nSubDestinoCod & "," & lvListaAguaSaneamiento(ixCD).nMontoS & " ," & lvListaAguaSaneamiento(ixCD).nBeneficia
        Conn.ConexionActiva.Execute (sSql)
    Next
    'END EAAS SEGUN ERS-054-2018*******
    'INICIO EAAS20191401 SEGUN 018-GM-DI_CMACM***

    If UBound(pvListaCreditoVerde) > 0 Then
        lvListaCreditoVerde = pvListaCreditoVerde
    Else
        ReDim lvListaCreditoVerde(0)
    End If
    
    sSql = "EXEC stp_del_ERS054_2018_ColocDestinoCreditoVerde '" & psCtaCod & "'"
    Conn.ConexionActiva.Execute (sSql)
    Dim nSumaTotalCreditoVerde As Double
    Dim nSumaCV As Double
    nSumaCV = 0
    Dim nSumaTotalBeneficiadosCV As Integer
    For ixCD = 1 To UBound(lvListaCreditoVerde)
        nSumaTotalCreditoVerde = nSumaTotalCreditoVerde + lvListaCreditoVerde(ixCD).nMontoS
        nSumaTotalBeneficiadosCV = nSumaTotalBeneficiadosCV + lvListaCreditoVerde(ixCD).nBeneficia
    Next
    
    If (UBound(lvListaCreditoVerde) > 0) Then
        For ixCD = 1 To 1
            sSql = "EXEC stp_ins_ERS054_2018_ColocDestinoCreditoVerde '" & psCtaCod & "'," & ixCD & "," & lvListaCreditoVerde(ixCD).nDestinoCod & "," & pnMontoSol & "," & nSumaTotalCreditoVerde & ", " & lvListaCreditoVerde(ixCD).nBeneficia
            Conn.ConexionActiva.Execute (sSql)
        Next
    End If
    For ixCD = 1 To UBound(lvListaCreditoVerde)
        sSql = "EXEC stp_ins_ERS054_2018_ColocDestinoCreditoVerdeDet '" & psCtaCod & "'," & ixCD & "," & lvListaCreditoVerde(ixCD).nSubDestinoCod & "," & lvListaCreditoVerde(ixCD).nMontoS & " ," & lvListaCreditoVerde(ixCD).nBeneficia
        Conn.ConexionActiva.Execute (sSql)
    Next
    'END EAAS20191401 SEGUN 018-GM-DI_CMACM*******
    
'      'ALPA 20160702*************************
'      sSql = " exec stp_ins_CredEndeuLineaNoUtilizada_solicitud '" & lsPersTitular & "','" & psCtaCod & "'"
'      Conn.ConexionActiva.Execute (sSql)
'      '*************************************
        'CTI3 20210107*************************
        'ReDim pvTablasRSE(2)
        Set oRs = New ADODB.Recordset
        sSql = " exec stp_ins_RSE_Admision_Deuda '" & psCtaCod & "'"
        Set oRs = Conn.ConexionActiva.Execute(sSql)
        Dim sTabla As String
        Dim sEncontrado As String
        If Not (oRs.BOF Or oRs.EOF) Then
            sEncontrado = oRs!cEncontrado
            sTabla = oRs!cTabla
            pvTablasRSE(1) = sEncontrado
            pvTablasRSE(2) = sTabla
        End If
        If sEncontrado = "NO" Then
            If bTran Then
                Conn.ConexionActiva.RollbackTrans
                Set Conn = Nothing
            End If
            Exit Sub
        End If
        Set oRs = Nothing
      
        If UBound(lvListaCompraDeuda) > 0 Then
            sSql = " exec stp_upd_RSE_Deudas_SSFF_Admision '" & psCtaCod & "'"
            Conn.ConexionActiva.Execute (sSql)
        End If
    'EJVG20160712 ***
    sSql = "EXEC stp_upd_MontoExposicionEsteCredito '" & psCtaCod & "'," & pnMontoExpEsteCred
    Conn.ConexionActiva.Execute (sSql)
    
    If pbEliminarEvaluacion Then
        sSql = "EXEC spt_del_CredFormEval '" & psCtaCod & "'"
        Conn.ConexionActiva.Execute (sSql)
    End If
    'END EJVG *******
    Conn.ConexionActiva.CommitTrans
    Conn.CierraConexion
    Set Conn = Nothing

    Exit Sub

ErrorActualizaSolicitud:
    If bTran Then
        Conn.ConexionActiva.RollbackTrans
        Set Conn = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso ActualizaSolicitud", Err.Description

End Sub

'ARCV 27-10-2006
' psCargo As String
' psCARBEN As String
' psT_Plani As String
'Se agregaron para los Convenios

'ByVal psFteIngreso As String
'ByVal dFecEvalFte As Date,
Public Function TransferenciaCar(ByVal psCtaCodAnte As String, _
                            ByVal psPerscodAnalista As String, _
                            Optional ByVal pdFecSis As Date = 0, _
                            Optional ByVal psCodUser As String = "", _
                            Optional ByVal psCodCmac As String = "", _
                            Optional ByVal psCodAge As String = "", _
                            Optional ByVal pProducto As Producto, _
                            Optional ByVal pnMoneda As Moneda, _
                            Optional ByVal psMovAct As String = "", _
                            Optional ByVal nTipoTrans As Integer = 1, _
                            Optional ByVal nProductoNew As Integer, _
                            Optional ByRef cPersNombre As String = "") As String
    Dim pbTran As Boolean
    Dim Conn As COMConecta.DCOMConecta
    Dim sSql As String
    Dim psCtaCod As String
    Dim pnFlag As Integer 'Indica si el producto tiene equivalencia en la reclasificación
    pnFlag = 0
    psCtaCod = ""
    Dim oGen As COMDConstSistema.DCOMGeneral
    
On Error GoTo ErrorTransferenciaCar
    cPersNombre = RecuperaNombreTitularCredito(psCtaCodAnte)
    pbTran = False
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    
    If nTipoTrans = 2 Then
    pProducto = nProductoNew & Mid(pProducto, 2, 2)
    ''pProducto = 202
    End If
    Set oGen = New COMDConstSistema.DCOMGeneral
    psCtaCod = psCodCmac & oGen.GeneraNuevaCuenta(psCodAge, pProducto, pnMoneda)
    Set oGen = Nothing
    Dim rsCuentas As ADODB.Recordset
    Set rsCuentas = New ADODB.Recordset
    'Inicio Transferencia de carteras
    Conn.ConexionActiva.BeginTrans
    pbTran = True
    If nTipoTrans = 2 Then
        sSql = "exec stp_pro_ReclasificacionCreditos '" & psCtaCodAnte & "','" & psCtaCod & "','" & psPerscodAnalista & "','" & psCodUser & "','" & psMovAct & "'"
        pnFlag = 1
    Else
        sSql = "exec stp_pro_TraspasoCarteraAgencias '" & psCtaCodAnte & "','" & psCtaCod & "','" & psPerscodAnalista & "','" & psCodUser & "','" & psMovAct & "'"
        pnFlag = 1
    End If
    If pnFlag = 1 Then
        Conn.CargaRecordSet (sSql)
    End If
    Conn.ConexionActiva.CommitTrans
    sSql = "exec stp_sel_CuentaNueva '" & psCtaCodAnte & "'"
    Set rsCuentas = Conn.CargaRecordSet(sSql)
    If rsCuentas.RecordCount > 0 Then
        psCtaCod = rsCuentas!valor
    Else
        If pnFlag = 0 Then
          psCtaCod = "No tiene equivalencia"
        Else
         psCtaCod = psCtaCod
        End If
    End If
    Conn.CierraConexion
    TransferenciaCar = psCtaCod
    Set Conn = Nothing
    Exit Function

ErrorTransferenciaCar:
    If pbTran Then
        Conn.ConexionActiva.RollbackTrans
        Set Conn = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso Transferencia de Cartera", Err.Description
    TransferenciaCar = ""
End Function
Public Sub NuevaSolicitud(ByVal oRelPersCred As Variant, ByRef psCtaCod As String, _
                          ByVal psCondCred As String, ByVal psCondCredOtra As String, ByVal pnMontoSol As Double, _
                          ByVal pnCuotasSol As Integer, ByVal pnPlazo As Integer, ByVal psDestCred As String, _
                          ByVal psPerscodAnalista As String, ByVal psPersCodInst As String, ByVal psCodModular As String, _
                          ByVal dFecSol As Date, ByVal psMovAct As String, ByVal bRefinanciar As Boolean, _
                          ByVal pnCondProd As Integer, Optional ByVal MatCredRef As Variant, _
                          Optional pnCapitalInt As Boolean = False, Optional bSustiDeudor As Boolean = False, _
                          Optional ByVal bAmpliado As Boolean = False, Optional rsAmpliado As ADODB.Recordset, _
                          Optional ByVal nCampanaCod As Integer, Optional ByVal pdFecSis As Date = 0, _
                          Optional ByVal psCodUser As String = "", Optional ByVal psCodCmac As String = "", Optional ByVal psCodAge As String = "", _
                          Optional ByVal psTipoProdCod As Producto, Optional ByVal pnMoneda As Moneda, Optional ByVal psCadenaCta As String = "", _
                          Optional ByVal psCargo As String = "", Optional ByVal psCARBEN As String = "", Optional ByVal psT_Plani As String = "", _
                          Optional ByVal pnMotivoRef As Integer, Optional ByVal nCasaCod As Integer = 0, Optional ByVal psVendedor As String = "", _
                          Optional ByVal psMotivoDetRef As String = "", Optional ByVal pbGrabaPromotor As Boolean = False, Optional ByVal psCodPromotor As String = "", _
                          Optional ByRef pvListaCompraDeuda As Variant, Optional ByRef pvListaAguasaneamiento As Variant, _
                          Optional ByVal pnSubDestino As Long = 0, Optional ByVal pnTpDoc As Long = 0, Optional ByVal pnTpIngr As Long = 0, Optional ByVal pnTpInt As Long = 0, Optional ByRef pvListaCreditoVerde As Variant, _
                          Optional ByRef pvTablasRSE As Variant) 'EJVG20160203 ERS002-2016 Se agregó pvListaCompraDeuda 'EAAS20191401 SEGUN 018-GM-DI_CMACM pvListaCreditoVerde
    'By capi 28102008 se agrego parametro pnMotivoRef
    'MIOL 20130626 se agrego parametro psVendedor
'MADM 20100719 - Optional ByVal nCampanaCod as Integer
'WIOR 20140509 AGREGO  pbGrabaPromotor,psCodPromotor
'JOEP20190919 ERS042 CP-2018 Agrego pnSubDestino,pnTpDoc,pnTpIngr,pnTpInt

'Public Sub NuevaSolicitud(ByVal oRelPersCred As Variant, ByRef psCtaCod As String, _
'                          ByVal psCondCred As String, ByVal psCondCredOtra As String, ByVal pMatFuentes As Variant, ByVal pnMontoSol As Double, _
'                          ByVal pnCuotasSol As Integer, ByVal pnPlazo As Integer, ByVal psDestCred As String, _
'                          ByVal psPerscodAnalista As String, ByVal psPersCodInst As String, ByVal psCodModular As String, _
'                          ByVal dFecSol As Date, ByVal psMovAct As String, ByVal bRefinanciar As Boolean, _
'                          ByVal pnCondProd As Integer, Optional ByVal MatCredRef As Variant, _
'                          Optional pnCapitalInt As Boolean = False, Optional bSustiDeudor As Boolean = False, _
'                          Optional ByVal bAmpliado As Boolean = False, Optional rsAmpliado As ADODB.Recordset, _
'                          Optional ByVal nCampanaCod As Integer, Optional ByVal pdFecSis As Date = 0, _
'                          Optional ByVal psCodUser As String = "", Optional ByVal psCodCmac As String = "", Optional ByVal psCodAge As String = "", _
'                          Optional ByVal pProducto As Producto, Optional ByVal pnMoneda As Moneda, Optional ByVal psCadenaCta As String = "", _
'                          Optional ByVal psCargo As String = "", Optional ByVal psCARBEN As String = "", Optional ByVal psT_Plani As String = "")


Dim pbTran As Boolean
Dim i As Integer
Dim sSql As String
Dim Conn As COMConecta.DCOMConecta
Dim oBase As DCOMCredActBD
Dim dFecGrab As Date
Dim dGar As DCOMGarantia
Dim R As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim lcInstitucionCmac As String
Dim lsPersTitular As String 'ALPA20160702
'Dim LsNumGarant As String
Dim oGen As COMDConstSistema.DCOMGeneral
Dim lnInteresDiferido As Currency 'ALPA 20150826********
Dim lnRefinanciadoDeVigenteCancelado As Currency 'ALPA 20150909
Dim oRs As ADODB.Recordset 'ALPA 20150909
Dim lvListaCompraDeuda() As TCompraDeuda 'EJVG20160203 ERS002-2016
Dim lvListaAguaSaneamiento() As TAguaSaneamiento 'EAAS20180815 ERS054-2018
Dim lvListaCreditoVerde() As TCreditoVerde 'EAAS20191401 SEGUN 018-GM-DI_CMACM
Dim ixCD As Integer 'EJVG20160203 ERS002-2016
        Set oRs = New ADODB.Recordset
    On Error GoTo ErrorNuevaSolicitud

    pbTran = False
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion

    psMovAct = GeneraMovNroActualiza(pdFecSis, psCodUser, psCodCmac, psCodAge)
    
    Set oGen = New COMDConstSistema.DCOMGeneral
    psCtaCod = psCodCmac & oGen.GeneraNuevaCuenta(psCodAge, psTipoProdCod, pnMoneda)
    Set oGen = Nothing

    ' se obtiene el valor de la instucion de  la cmac ica para creditos de los trabajadores
    'ALPA 20100606 B2 *********************************
'    If Mid(psCtaCod, 6, 3) = "320" Then
'        ' si es credito de trabajadores y funcionarios
'        sSql = "Select nConsSisValor From ConstSistema where nConsSisCod=204"
'        Set rs = Conn.CargaRecordSet(sSql)
'        If Not rs.EOF And Not rs.BOF Then
'            lcInstitucionCmac = rs!nConsSisValor
'        End If
'        Set rs = Nothing
'    End If
   '**************************************************
    Conn.ConexionActiva.BeginTrans
    pbTran = True

    'Ingresa Nuevo Producto
    sSql = "INSERT INTO Producto(cCtaCod,nTasaInteres,nSaldo,nPrdEstado,dPrdEstado,nTransacc)"
    sSql = sSql & " VALUES('" & psCtaCod & "',"
    sSql = sSql & "0.00,"
    sSql = sSql & "0.00,"
    sSql = sSql & Trim(Str(gColocEstSolic)) & ","
    sSql = sSql & "'" & Format(dFecSol, "mm/dd/yyyy") & "',"
    sSql = sSql & "0)"
    Conn.ConexionActiva.Execute sSql

    'Ingresa Nuevas Personas Relacionadas con el Credito
'    oRelPersCred.IniciarMatriz
'    Do While Not oRelPersCred.EOF
'        sSql = "INSERT INTO ProductoPersona(cCtaCod,cPersCod,nPrdPersRelac) "
'        sSql = sSql & " VALUES('" & psCtaCod & "','" & oRelPersCred.ObtenerCodigo & "'," & oRelPersCred.ObtenerValorRelac & ")"
'        Conn.ConexionActiva.Execute sSql
'        oRelPersCred.siguiente
'    Loop
    For i = 0 To UBound(oRelPersCred) - 1
        sSql = "INSERT INTO ProductoPersona(cCtaCod,cPersCod,nPrdPersRelac) "
        sSql = sSql & " VALUES('" & psCtaCod & "','" & oRelPersCred(i, 0) & "'," & oRelPersCred(i, 1) & ")"
        Conn.ConexionActiva.Execute sSql
        'ALPA20160702*****************************
        If oRelPersCred(i, 1) = 20 Then
            lsPersTitular = oRelPersCred(i, 0)
        End If
        '****************************************
    Next

    'Ingresa el Analista
    sSql = "INSERT INTO ProductoPersona(cCtaCod,cPersCod,nPrdPersRelac) "
    sSql = sSql & " VALUES('" & psCtaCod & "','" & psPerscodAnalista & "'," & gColRelPersAnalista & ")"
    Conn.ConexionActiva.Execute sSql
    
    'WIOR 20140509(PROMOTORES) **************
    If pbGrabaPromotor Then
        sSql = "EXEC stp_ins_upd_CredPromotores '" & psCtaCod & "','" & psCodPromotor & "'"
        Conn.ConexionActiva.Execute sSql
    End If
    'WIOR FIN *******************************

    'Nuevo Registro de Colocaciones
    'sSql = "INSERT INTO Colocaciones(cCtaCod,nPlazo,dVenc,nMontoCol,nColocCalendCod,cUltimaActualizacion,cLineaCred)"
    'ALPA 20100604 BASII************************************
    sSql = "INSERT INTO Colocaciones(cCtaCod,nPlazo,dVenc,nMontoCol,nColocCalendCod,cUltimaActualizacion,cTpoProdCod,cAgeCodAct,cLineaCred,nTpIngr)"
    '*********************
    sSql = sSql & " VALUES('" & psCtaCod & "',"
    sSql = sSql & Trim(Str(pnPlazo)) & ","
    sSql = sSql & "NULL" & ","
    sSql = sSql & Format(pnMontoSol, "#0.00") & ","
    sSql = sSql & "0,"
    sSql = sSql & "'" & psMovAct & "',"
    'ALPA 20100604 BASII************************************
    sSql = sSql & "'" & psTipoProdCod & "', '" & psCodAge & "',"
    'sSql = sSql & "'')"
    '*******************************************************
    sSql = sSql & "'',"
    sSql = sSql & "" & IIf(pnTpIngr = 0, "NULL", pnTpIngr) & ")"
    
    '*******************************************************
    Conn.ConexionActiva.Execute sSql

    'Nuevo Registro de ColocacCred --- cambio metodo liquidacion MICG -- 18-05-2006 AVMM
    sSql = "INSERT INTO ColocacCred (cCtaCod,nDiasAtraso,nColocCondicion,nColocCondicion2,nColocDestino,"
    sSql = sSql & "cProtesto,bCargoAuto,cMetLiquidacion,bRefCapInt,nNroProxCuota,nIntPend,"
    sSql = sSql & "nExoPenalidad, nColocCondicionProd,IdCampana,id_CasaCom,cVendedor,nSubDestino,nTpDoc,nTpIngr,nTpInt)" 'JOEP20190919 ERS042 CP-2018 ,pnSubDestino,pnTpDoc,pnTpIngr
    sSql = sSql & "VALUES('" & psCtaCod & "',"
    sSql = sSql & "0,"
    sSql = sSql & psCondCred & ","
    'ARCV 04-02-2007 (TEMPORAL)
    'sSql = sSql & psCondCredOtra & ","
    sSql = sSql & psCondCred & ","
    '---------
    sSql = sSql & psDestCred & ","
    sSql = sSql & "0,"
    sSql = sSql & "NULL,"
    sSql = sSql & "'GMIC'," 'Por defecto en Maynas
    sSql = sSql & IIf(pnCapitalInt, "1,", "0,")
    sSql = sSql & "1,"
    sSql = sSql & "0,"
    sSql = sSql & "NULL," & pnCondProd & ","
    'sSQL = sSQL & IIf(nCampanaCod = 0, "NULL", nCampanaCod) & ")"
    sSql = sSql & nCampanaCod & " ,"
    sSql = sSql & IIf(nCasaCod = 0, "NULL", nCasaCod) & " ,"
    sSql = sSql & "'" & psVendedor & "',"   'MIOL 20130626, RQ13335
    'JOEP20190919 ERS042 CP-2018
    sSql = sSql & "" & IIf(pnSubDestino = 0, "NULL", pnSubDestino) & ","
    sSql = sSql & "" & IIf(pnTpDoc = 0, "NULL", pnTpDoc) & ","
    sSql = sSql & "" & IIf(pnTpIngr = 0, "NULL", pnTpIngr) & ","
    sSql = sSql & "" & IIf(pnTpInt = 0, "NULL", pnTpInt) & ")"
    'JOEP20190919 ERS042 CP-2018
    Conn.ConexionActiva.Execute sSql

    'Nuevo Registro en ColocEstado
    Set oBase = New DCOMCredActBD
    sSql = "INSERT INTO ColocacEstado(cCtaCod,dPrdEstado,nPrdEstado,nCuotas,nMonto,cDescripcion,"
    sSql = sSql & "nColocCalendCod,nPeriodoFechaFija,nPeriodoGracia,nPlazo,nTipoGracia)"
    sSql = sSql & " VALUES('" & psCtaCod & "',"
    sSql = sSql & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
    sSql = sSql & Trim(Str(gColocEstSolic)) & ","
    sSql = sSql & Trim(Str(pnCuotasSol)) & ","
    sSql = sSql & Format(pnMontoSol, "#0.00") & ","
    sSql = sSql & "'Nueva Solicitud de Credito',"
    sSql = sSql & "NULL,"
    sSql = sSql & "NULL,"
    sSql = sSql & "NULL,"
    sSql = sSql & Trim(Str(pnPlazo)) & ","
    sSql = sSql & "NULL)"
    Set oBase = Nothing
    Conn.ConexionActiva.Execute sSql

    'Nuevo Registro en ColocFteIngreso
'ARCV 30-12-2006
'    sSQL = "INSERT INTO ColocFteIngreso(cNumFuente,dPersEval,cCtaCod)"
'            sSQL = sSQL & " VALUES('" & psFteIngreso & "',"
'            sSQL = sSQL & "'" & Format(dFecEvalFte, "mm/dd/yyyy") & "', "
'            sSQL = sSQL & "'" & psCtaCod & "')"
'            Conn.ConexionActiva.Execute sSQL

'*** COMENTADO POR PEAC 20080412 segun Acta No 009-2008/ti-d
'*** este proceso se cambio a la sugerencia
'    For i = 0 To UBound(pMatFuentes) - 1
'        sSql = "INSERT INTO ColocFteIngreso(cNumFuente,dPersEval,cCtaCod)"
'            sSql = sSql & " VALUES('" & pMatFuentes(i, 0) & "',"
'            sSql = sSql & "'" & Format(pMatFuentes(i, 1), "mm/dd/yyyy") & "', "
'            sSql = sSql & "'" & psCtaCod & "')"
'            Conn.ConexionActiva.Execute sSql
'    Next i
'-----------------

    'Nuevo Registro de ColocConvenio
    If Len(Trim(psPersCodInst)) > 0 Then
        sSql = "INSERT INTO ColocacConvenio(cCtaCod,cPersCod,cCodModular"
        sSql = sSql & ",cCargo,cCARBEN,cT_Plani )" 'ARCV 27-10-2006
        sSql = sSql & " VALUES('" & psCtaCod & "',"
        sSql = sSql & "'" & psPersCodInst & "',"
        sSql = sSql & "'" & psCodModular & "'"
        'ARCV 27-10-2006
        sSql = sSql & ",'" & psCargo & "','" & psCARBEN & "','" & psT_Plani & "')"
        '----------------
        
        Conn.ConexionActiva.Execute sSql
    End If

    'Si es credito de funcionarios de cmac ica
'    If Mid(psCtaCod, 6, 3) = "320" And Len(lcInstitucionCmac) = 13 Then
'        sSql = "INSERT INTO ColocacConvenio(cCtaCod,cPersCod,cCodModular)"
'        sSql = sSql & " VALUES('" & psCtaCod & "',"
'        sSql = sSql & "'" & lcInstitucionCmac & "',"
'        sSql = sSql & "'" & " " & "')"
'        Conn.ConexionActiva.Execute sSql
'    End If
    'Si es una Refinanciacion Guardar los Datos
    lnRefinanciadoDeVigenteCancelado = 0
    If bRefinanciar Then
        If IsArray(MatCredRef) Then
            Set oBase = New DCOMCredActBD
            dFecGrab = CDate(Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "dd/mm/yyyy hh:mm:ss"))
            Set oBase = Nothing
            If UBound(MatCredRef) > 0 Then
                For i = 0 To UBound(MatCredRef) - 1
                    'By capi 28102008 se inserto para que grabe motivo de refinanciacion
                    
'                    sSql = "INSERT INTO ColocacRefinanc(cCtaCod,cCtaCodRef,nEstado,dEstado,nMontoRef)"
'                    sSql = sSql & " VALUES('" & psCtaCod & "','"
'                    sSql = sSql & MatCredRef(i, 0) & "',"
'                    sSql = sSql & gColocEstadoRefinancSolicitado & ","
'                    sSql = sSql & "'" & Format(dFecGrab, "mm/dd/yyyy hh:mm:ss") & "',"
'                    sSql = sSql & MatCredRef(i, 1) & ")"
'                    Conn.ConexionActiva.Execute sSql
                    
                    sSql = "INSERT INTO ColocacRefinanc(cCtaCod,cCtaCodRef,nEstado,dEstado,nMontoRef,nMotivoRef,sMotivoDetRef)"
                    sSql = sSql & " VALUES('" & psCtaCod & "','"
                    sSql = sSql & MatCredRef(i, 0) & "',"
                    sSql = sSql & gColocEstadoRefinancSolicitado & ","
                    sSql = sSql & "'" & Format(dFecGrab, "mm/dd/yyyy hh:mm:ss") & "',"
                    sSql = sSql & MatCredRef(i, 1) & "," & pnMotivoRef
                    sSql = sSql & ",'" & psMotivoDetRef & "')" 'JAME20140509
                    Conn.ConexionActiva.Execute sSql
                    'End by

                    'Ingreso del Detalle
                    'Capital
                    sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                    sSql = sSql & " VALUES('" & psCtaCod & "','"
                    sSql = sSql & MatCredRef(i, 0) & "',"
                    sSql = sSql & gColocEstadoRefinancSolicitado & ","
                    sSql = sSql & "'" & Format(dFecGrab, "mm/dd/yyyy hh:mm:ss") & "',"
                    sSql = sSql & gColocConceptoCodCapital & ","
                    sSql = sSql & MatCredRef(i, 2) & ","
                    sSql = sSql & "0.00)"
                    Conn.ConexionActiva.Execute sSql
                    
                    If pnCapitalInt Then 'MAVM 20110117
                        'Interes Compensatorio
                        If CDbl(MatCredRef(i, 3)) > 0 Then
                            sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                            sSql = sSql & " VALUES('" & psCtaCod & "','"
                            sSql = sSql & MatCredRef(i, 0) & "',"
                            sSql = sSql & gColocEstadoRefinancSolicitado & ","
                            sSql = sSql & "'" & Format(dFecGrab, "mm/dd/yyyy hh:mm:ss") & "',"
                            sSql = sSql & gColocConceptoCodInteresCompensatorio & ","
                            sSql = sSql & MatCredRef(i, 3) & ","
                            sSql = sSql & "0.00)"
                            Conn.ConexionActiva.Execute sSql
                        End If
    
                        'Interes Moratorio
                        If CDbl(MatCredRef(i, 4)) > 0 Then
                            sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                            sSql = sSql & " VALUES('" & psCtaCod & "','"
                            sSql = sSql & MatCredRef(i, 0) & "',"
                            sSql = sSql & gColocEstadoRefinancSolicitado & ","
                            sSql = sSql & "'" & Format(dFecGrab, "mm/dd/yyyy hh:mm:ss") & "',"
                            sSql = sSql & gColocConceptoCodInteresMoratorio & ","
                            sSql = sSql & MatCredRef(i, 4) & ","
                            sSql = sSql & "0.00)"
                            Conn.ConexionActiva.Execute sSql
                        End If
    
                        'Interes Gracia
                        If CDbl(MatCredRef(i, 5)) > 0 Then
                            sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                            sSql = sSql & " VALUES('" & psCtaCod & "','"
                            sSql = sSql & MatCredRef(i, 0) & "',"
                            sSql = sSql & gColocEstadoRefinancSolicitado & ","
                            sSql = sSql & "'" & Format(dFecGrab, "mm/dd/yyyy hh:mm:ss") & "',"
                            sSql = sSql & gColocConceptoCodInteresGracia & ","
                            sSql = sSql & MatCredRef(i, 5) & ","
                            sSql = sSql & "0.00)"
                            Conn.ConexionActiva.Execute sSql
                        End If
    
                        'Interes Suspenso
                        If CDbl(MatCredRef(i, 6)) > 0 Then
                            sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                            sSql = sSql & " VALUES('" & psCtaCod & "','"
                            sSql = sSql & MatCredRef(i, 0) & "',"
                            sSql = sSql & gColocEstadoRefinancSolicitado & ","
                            sSql = sSql & "'" & Format(dFecGrab, "mm/dd/yyyy hh:mm:ss") & "',"
                            sSql = sSql & gColocConceptoCodInteresSuspenso & ","
                            sSql = sSql & MatCredRef(i, 6) & ","
                            sSql = sSql & "0.00)"
                            Conn.ConexionActiva.Execute sSql
                        End If
    
                        'Interes Reprogramado
                        If CDbl(MatCredRef(i, 7)) > 0 Then
                            sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                            sSql = sSql & " VALUES('" & psCtaCod & "','"
                            sSql = sSql & MatCredRef(i, 0) & "',"
                            sSql = sSql & gColocEstadoRefinancSolicitado & ","
                            sSql = sSql & "'" & Format(dFecGrab, "mm/dd/yyyy hh:mm:ss") & "',"
                            sSql = sSql & gColocConceptoCodInteresReprogramado & ","
                            sSql = sSql & MatCredRef(i, 7) & ","
                            sSql = sSql & "0.00)"
                            Conn.ConexionActiva.Execute sSql
                        End If
    
                        'Gastos
                        If CDbl(MatCredRef(i, 8)) > 0 Then
                            sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                            sSql = sSql & " VALUES('" & psCtaCod & "','"
                            sSql = sSql & MatCredRef(i, 0) & "',"
                            sSql = sSql & gColocEstadoRefinancSolicitado & ","
                            sSql = sSql & "'" & Format(dFecGrab, "mm/dd/yyyy hh:mm:ss") & "',"
                            sSql = sSql & gColocConceptoCodGastoVarios & ","
                            sSql = sSql & MatCredRef(i, 8) & ","
                            sSql = sSql & "0.00)"
                            Conn.ConexionActiva.Execute sSql
                        End If
                    End If
                    'ALPA 20150803
                    lnInteresDiferido = lnInteresDiferido + MatCredRef(i, 7) + MatCredRef(i, 6) + MatCredRef(i, 5) + MatCredRef(i, 4) + MatCredRef(i, 3)
                    Set oRs = New ADODB.Recordset
                    sSql = "Exec stp_sel_obtenerDiferidoCreditoRefinanciadoDeVigenteCancelado '" & MatCredRef(i, 0) & "','" & Format(dFecGrab, "YYYY/MM/DD") & "'," & Mid(psCtaCod, 9, 1)
                    Set oRs = Conn.ConexionActiva.Execute(sSql)
                    If Not (oRs.BOF Or oRs.EOF) Then
                        lnRefinanciadoDeVigenteCancelado = oRs!nMontoDiferidoCredCancelado
                    End If
                    lnInteresDiferido = lnInteresDiferido + lnRefinanciadoDeVigenteCancelado
                    Set oRs = Nothing
                    '*******************************
                Next i
                'ALPA 20150803**********************
        
                
                sSql = "Update ColocacCred "
                sSql = sSql & " Set nInteresDiferido=" & lnInteresDiferido
                sSql = sSql & " Where cCtaCod='" & psCtaCod & "'"
                Conn.ConexionActiva.Execute sSql
                '***********************************
            End If
        End If
    End If

    'SI es Refinanciado Ingresar las Mismas Garantias de los Creditos Anteriores
    Dim cNumGaranOld As String
    Dim nNum As Integer
    'EJVG20160307 *** Comentar
    'If bRefinanciar Then

     ' For i = 0 To UBound(MatCredRef) - 1
     '       If i = 0 Then
     '           LsNumGarant = "'" & MatCredRef(i, 0) & "'"
     '       Else
     '           LsNumGarant = LsNumGarant & ",'" & MatCredRef(i, 0) & "'"
     '       End If
     ' Next i

      'sSql = "Select CN.nConsValor as nTpoGar, G.cNroDoc, CN.cConsDescripcion cTpoGarantia, G.cDescripcion, D.cDocDesc, Sum(CG.nGravado) as nGravado, CN2.cConsDescripcion cMoneda, SUM(G.nPorGravar) as nPorGravar, Sum(G.nGravaMent) as nGravaMent, CN3.cConsDescripcion cMonedaGar,G.cNumgarant, CG.nMoneda, CG.nEstado "
      'sSql = sSql & " From ColocGarantia CG"
      'sSql = sSql & " Inner Join Garantias G ON CG.cNumGarant = G.cNumGarant"
      'sSql = sSql & " Inner Join Constante CN ON G.nTpoGarantia = CN.nConsValor AND CN.nConsCod = 1027"
      'sSql = sSql & " Inner Join Constante CN2 ON CN2.nConsValor = CG.nMoneda AND CN2.nConsCod = 1011"
      'sSql = sSql & " Inner Join Constante CN3 ON CN3.nConsValor = G.nMoneda AND CN3.nConsCod = 1011"
      'sSql = sSql & " Inner Join Documento D ON D.nDocTpo = Convert(int,G.cTpoDoc)"
      'sSql = sSql & " Inner Join GarantDoc O ON D.nDocTpo = O.nDocTpo  and O.nConsValor=G.nTpoGarantia"
      'sSql = sSql & " WHERE CG.cCtaCod in(" & LsNumGarant & ")"
      'sSql = sSql & " GROUP BY CG.nEstado,CN.nConsValor,G.cNroDoc,CN.cConsDescripcion,G.cDescripcion,D.cDocDesc,CN2.cConsDescripcion,CN3.cConsDescripcion,G.cNumgarant,CG.nMoneda"
      'sSql = sSql & " Order By G.cNumgarant"

       ' Set R = Conn.CargaRecordSet(sSql)
       ' Do Until R.EOF
       ' sSql = "INSERT INTO COLOCGARANTIA(cNumGarant, cCtaCod, nMoneda, nGravado, nEstado) "
       ' sSql = sSql & " VALUES('" & R!cNumGarant & "','" & psCtaCod & "'," & R!nMoneda & "," & Format(R!nGravado, "#0.00") & "," & IIf(IsNull(R!nEstado), 1, R!nEstado) & ") "
       ' Conn.ConexionActiva.Execute sSql
       ' R.MoveNext
       ' Loop
       ' Set R = Nothing
    'End If
        ' un credito automatico
'        oRelPersCred.IniciarMatriz
'        Do While Not oRelPersCred.EOF
'                If oRelPersCred.ObtenerValorRelac = 20 Then ' si es titular
'                    If psCondCredOtra = "3" Then
'                        sSql = "Update Persona set bPreferencial=1 Where cPersCod='" & oRelPersCred.ObtenerCodigo & "'"
'                    Else
'                        sSql = "Update Persona set bPreferencial=0 Where cPersCod='" & oRelPersCred.ObtenerCodigo & "'"
'                    End If
'                    Conn.ConexionActiva.Execute sSql
'                End If
'            oRelPersCred.siguiente
'        Loop
    For i = 0 To UBound(oRelPersCred) - 1
        If oRelPersCred(i, 1) = 20 Then ' si es titular
            If psCondCredOtra = "3" Then
                sSql = "Update Persona set bPreferencial=1 Where cPersCod='" & oRelPersCred(i, 0) & "'"
            Else
                sSql = "Update Persona set bPreferencial=0 Where cPersCod='" & oRelPersCred(i, 0) & "'"
            End If
            Conn.ConexionActiva.Execute sSql
        End If
    Next

      ' AUTOR:LMMD
      ' SI SE TRATA DE UN CREDITO AMPLIADO
      If bAmpliado = True Then

            Dim nEstado As Integer
            'EJVG20160307 *** Comentar
            'rsAmpliado.MoveFirst
            
            'FRHU 20140424 TI-ERS015-2014
            'Dim slCuentas As String
            'slCuentas = ""
            'Do Until rsAmpliado.EOF
            '    slCuentas = slCuentas & rsAmpliado(0) & ","
            '    rsAmpliado.MoveNext
            'Loop
            'If slCuentas <> "" Then
            '    slCuentas = Mid(slCuentas, 1, Len(slCuentas) - 1)
            '    sSql = "EXEC stp_sel_ColocGarantiaParaAmpliado '" & slCuentas & "'"
            '    Set R = Conn.CargaRecordSet(sSql)
            '    Do Until R.EOF
            '        sSql = "INSERT INTO COLOCGARANTIA(cNumGarant, cCtaCod, nMoneda, nGravado, nEstado) "
            '        sSql = sSql & " VALUES('" & R!cNumGarant & "','" & psCtaCod & "'," & R!nMoneda & "," & Format(R!nGravado, "#0.00") & "," & IIf(IsNull(R!nEstado), 1, R!nEstado) & ") "
            '        Conn.ConexionActiva.Execute (sSql)
            '        R.MoveNext
            '    Loop
            '    Set R = Nothing
            'End If
            rsAmpliado.MoveFirst
            'FIN FRHU 20140424
            
            Do Until rsAmpliado.EOF

                sSql = "Select nPrdEstado From Producto Where cCtaCod='" & rsAmpliado(0) & "'"
                Set rs = New ADODB.Recordset
                Set rs = Conn.CargaRecordSet(sSql)

                'SE OBTIENE EL ESTADO
                If Not rs.EOF And Not rs.BOF Then
                    nEstado = rs!nPrdEstado
                End If

                sSql = "Insert Into ColocacAmpliado Values('" & psCtaCod & "','" & rsAmpliado(0) & "'," & nEstado & ","
                sSql = sSql & "'" & Format(pdFecSis, "MM/dd/yyyy") & "'," & rsAmpliado(3) & "," & pnMontoSol & ")"
                
                Conn.ConexionActiva.Execute (sSql)
                
                'ARCV 12-02-2007
                
                'FRHU 20140424 TI-ERS015-2014
                'sSql = "INSERT INTO ColocGarantia " & _
                '       " SELECT cNumGarant,'" & psCtaCod & "',nMoneda,nGravado,nEstado FROM ColocGarantia WHERE cCtaCod='" & rsAmpliado(0) & "' AND nEstado=1"
                'Conn.ConexionActiva.Execute (sSql)
                'FIN FRHU 20140424
                
                sSql = "INSERT INTO ColocCredCredVig VALUES('" & psCtaCod & "','" & rsAmpliado(0) & "')"
                Conn.ConexionActiva.Execute (sSql)
                '---------
                rsAmpliado.MoveNext
          Loop
      End If
      
      'ALPA 20160702*************************
'      sSql = " exec stp_ins_CredEndeuLineaNoUtilizada_solicitud '" & lsPersTitular & "','" & psCtaCod & "'"
'      Conn.ConexionActiva.Execute (sSql)
       'CTI3 20210107*************************
        'ReDim pvTablasRSE(2)
        Set oRs = New ADODB.Recordset
        sSql = " exec stp_ins_RSE_Admision_Deuda '" & psCtaCod & "'"
        Set oRs = Conn.ConexionActiva.Execute(sSql)
        Dim sTabla As String
        Dim sEncontrado As String
        If Not (oRs.BOF Or oRs.EOF) Then
            sEncontrado = oRs!cEncontrado
            sTabla = oRs!cTabla
            pvTablasRSE(1) = sEncontrado
            pvTablasRSE(2) = sTabla
        End If
        If sEncontrado = "NO" Then
            If pbTran Then
                Conn.ConexionActiva.RollbackTrans
                Set Conn = Nothing
            End If
            Exit Sub
        End If
        Set oRs = Nothing
'        sSql = " exec stp_ins_CredDeudaSistemaFinanciero_solicitud '" & lsPersTitular & "','" & psCtaCod & "'"
'        Conn.ConexionActiva.Execute (sSql)
       
      '*************************************
      '**ARLO20171127 ERS070 - 2017
      sSql = " exec stp_ins_CredRevolventesNoRevolventes_solicitud '" & lsPersTitular & "','" & psCtaCod & "'"
      Conn.ConexionActiva.Execute (sSql)
      '**ARLO20171127 ERS070 - 2017
    'AVMM-27-09-2006
    
'    If bRefinanciar = True Then
'       'Buscar Garantias e Inserta
'        sSQL = " INSERT INTO ColocGarantia" & _
'               " SELECT cNumGarant,'" & psCtacod & "'  ,nMoneda,nGravado,nEstado FROM dbo.ColocGarantia where cCtaCod in " & psCadenaCta
'        Conn.ConexionActiva.Execute (sSQL)
'    End If
'
    'EJVG20160203 *** Compra Deuda otras IFIs
    If val(psDestCred) = ColocDestino.gColocDestinoCambEstructPasivo Then
        lvListaCompraDeuda = pvListaCompraDeuda
    Else
        ReDim lvListaCompraDeuda(0)
    End If

    For ixCD = 1 To UBound(lvListaCompraDeuda)
        sSql = "EXEC stp_ins_ERS0022016_ColocCompraDeuda '" & psCtaCod & "'," & ixCD & ",'" & lvListaCompraDeuda(ixCD).sIFICod & "'," & lvListaCompraDeuda(ixCD).nMoneda & ",'" & lvListaCompraDeuda(ixCD).sNroCredito & "'," & lvListaCompraDeuda(ixCD).nDestino & "," & IIf(lvListaCompraDeuda(ixCD).bRecompra, 1, 0) & ",'" & Format(lvListaCompraDeuda(ixCD).dDesembolso, "yyyymmdd") & "'," & lvListaCompraDeuda(ixCD).nMontoDesembolso & "," & lvListaCompraDeuda(ixCD).nNroCuotasPactadas & "," & lvListaCompraDeuda(ixCD).nNroCuotasPagadas & "," & lvListaCompraDeuda(ixCD).nSaldoComprar & "," & lvListaCompraDeuda(ixCD).nMontoCuota
        Conn.ConexionActiva.Execute (sSql)
        
    Next
    
    'CTI3 ERS032020************************************
    If UBound(lvListaCompraDeuda) > 0 Then
        sSql = " exec stp_upd_RSE_Deudas_SSFF_Admision '" & psCtaCod & "'"
        Conn.ConexionActiva.Execute (sSql)
    End If
    
    sSql = " exec stp_ins_CredDeudaSistemaFinanciero_solicitud '" & lsPersTitular & "','" & psCtaCod & "'"
    Conn.ConexionActiva.Execute (sSql)
    '**************************************************
    'END EJVG *******
    'INICIO EAAS20180912 SEGUN ERS 054-2018 ***

    If UBound(pvListaAguasaneamiento) > 0 Then
        lvListaAguaSaneamiento = pvListaAguasaneamiento
    Else
        ReDim lvListaAguaSaneamiento(0)
    End If
    
    sSql = "EXEC stp_del_ERS054_2018_ColocDestinoAguaSaneamiento '" & psCtaCod & "'"
    Conn.ConexionActiva.Execute (sSql)
    Dim nSumaTotalAguaSaneamiento As Double
    Dim nSuma As Double
    nSuma = 0
    Dim nSumaTotalBeneficiados As Integer
    Dim nSumaB As Integer
    nSumaB = 0
    
    For ixCD = 1 To UBound(lvListaAguaSaneamiento)
        nSumaTotalAguaSaneamiento = nSumaTotalAguaSaneamiento + lvListaAguaSaneamiento(ixCD).nMontoS
        nSumaTotalBeneficiados = nSumaTotalBeneficiados + lvListaAguaSaneamiento(ixCD).nBeneficia
    Next
    
    If (UBound(lvListaAguaSaneamiento) > 0) Then
        For ixCD = 1 To 1
            sSql = "EXEC stp_ins_ERS054_2018_ColocDestinoAguaSaneamiento '" & psCtaCod & "'," & ixCD & "," & lvListaAguaSaneamiento(ixCD).nDestinoCod & "," & pnMontoSol & "," & nSumaTotalAguaSaneamiento & ", " & lvListaAguaSaneamiento(ixCD).nBeneficia
            Conn.ConexionActiva.Execute (sSql)
        Next
    End If
    For ixCD = 1 To UBound(lvListaAguaSaneamiento)
        sSql = "EXEC stp_ins_ERS054_2018_ColocDestinoAguaSaneamientoDet '" & psCtaCod & "'," & ixCD & "," & lvListaAguaSaneamiento(ixCD).nSubDestinoCod & "," & lvListaAguaSaneamiento(ixCD).nMontoS & " ," & lvListaAguaSaneamiento(ixCD).nBeneficia
        Conn.ConexionActiva.Execute (sSql)
    Next
    'END EAAS20180912 SEGUN ERS 054-2018 *******
            'INICIO EAAS20191401 SEGUN 018-GM-DI_CMACM***

    If UBound(pvListaCreditoVerde) > 0 Then
        lvListaCreditoVerde = pvListaCreditoVerde
    Else
        ReDim lvListaCreditoVerde(0)
    End If
    
    sSql = "EXEC stp_del_ERS054_2018_ColocDestinoCreditoVerde '" & psCtaCod & "'"
    Conn.ConexionActiva.Execute (sSql)
    Dim nSumaTotalCreditoVerde As Double
    Dim nSumaCV As Double
    nSumaCV = 0
    Dim nSumaTotalBeneficiadosCV As Integer
    
    For ixCD = 1 To UBound(lvListaCreditoVerde)
        nSumaTotalCreditoVerde = nSumaTotalCreditoVerde + lvListaCreditoVerde(ixCD).nMontoS
        nSumaTotalBeneficiadosCV = nSumaTotalBeneficiadosCV + lvListaCreditoVerde(ixCD).nBeneficia
    Next
    
    If (UBound(lvListaCreditoVerde) > 0) Then
        For ixCD = 1 To 1
            sSql = "EXEC stp_ins_ERS054_2018_ColocDestinoCreditoVerde '" & psCtaCod & "'," & ixCD & "," & lvListaCreditoVerde(ixCD).nDestinoCod & "," & pnMontoSol & "," & nSumaTotalCreditoVerde & ", " & lvListaCreditoVerde(ixCD).nBeneficia
            Conn.ConexionActiva.Execute (sSql)
        Next
    End If
    For ixCD = 1 To UBound(lvListaCreditoVerde)
        sSql = "EXEC stp_ins_ERS054_2018_ColocDestinoCreditoVerdeDet '" & psCtaCod & "'," & ixCD & "," & lvListaCreditoVerde(ixCD).nSubDestinoCod & "," & lvListaCreditoVerde(ixCD).nMontoS & " ," & lvListaCreditoVerde(ixCD).nBeneficia
        Conn.ConexionActiva.Execute (sSql)
    Next
    'END EAAS20191401 SEGUN 018-GM-DI_CMACM*******
    'EJVG20160712 ***
    sSql = "EXEC stp_upd_MontoExposicionEsteCredito '" & psCtaCod & "'"
    Conn.ConexionActiva.Execute (sSql)
    'END EJVG *******
    
    Conn.ConexionActiva.CommitTrans
    Conn.CierraConexion
    Set Conn = Nothing

    Exit Sub

ErrorNuevaSolicitud:
    If pbTran Then
        Conn.ConexionActiva.RollbackTrans
        Set Conn = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso NuevaSolicitud", Err.Description
End Sub

Public Sub ActuializaPersRelacCred(ByVal poRelPersCred As Variant, ByVal psCtaCod As String, ByVal pbTransac As Boolean)
Dim sSql As String
Dim i As Integer

    On Error GoTo ErrorActuializaPersRelacCred
    If pbTransac Then
        Set Conn = New COMConecta.DCOMConecta
        Conn.AbreConexion
        Conn.ConexionActiva.BeginTrans
    End If
'    poRelPersCred.IniciarMatriz
'    If Not poRelPersCred.EOF Then
'        sSql = "DELETE ProductoPersona Where cCtaCod = '" & psCtaCod & "' AND nPrdPersRelac not in (" & gColRelPersAnalista & "," & gColRelPersApoderado & ")"
'        Conn.ConexionActiva.Execute sSql
'    End If
    If UBound(poRelPersCred) > 0 Then
        'sSql = "DELETE ProductoPersona Where cCtaCod = '" & psCtaCod & "' AND nPrdPersRelac not in (" & gColRelPersAnalista & "," & gColRelPersApoderado & ")"
        'MADM 20101002 ***
        sSql = "DELETE ProductoPersona Where cCtaCod = '" & psCtaCod & "' AND nPrdPersRelac not in (" & gColRelPersAnalista & "," & gColRelPersApoderado & "," & gColRelPersAcreedor & ")"
        '***
        Conn.ConexionActiva.Execute sSql
    End If

'    Do While Not poRelPersCred.EOF
'        sSql = "INSERT INTO ProductoPersona(cCtaCod,cPersCod,nPrdPersRelac) "
'        sSql = sSql & "       VALUES('" & psCtaCod & "','" & poRelPersCred.ObtenerCodigo & "'," & poRelPersCred.ObtenerValorRelac & ")"
'        Conn.ConexionActiva.Execute sSql
'        poRelPersCred.siguiente
'    Loop
    For i = 0 To UBound(poRelPersCred) - 1
        sSql = "INSERT INTO ProductoPersona(cCtaCod,cPersCod,nPrdPersRelac) "
        sSql = sSql & "       VALUES('" & psCtaCod & "','" & poRelPersCred(i, 0) & "'," & poRelPersCred(i, 1) & ")"
        Conn.ConexionActiva.Execute sSql
    Next

    If pbTransac Then
        Conn.ConexionActiva.CommitTrans
        Conn.CierraConexion
        Set Conn = Nothing
    End If
    Exit Sub

ErrorActuializaPersRelacCred:
    If pbTransac Then
        Conn.ConexionActiva.RollbackTrans
        Set Conn = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso ActuializaPersRelacCred", Err.Description

End Sub


'*****************************************************************************************
'***     Rutina:            RecuperaCodigosModulares
'***     Descripcion:       Permite Obtener los Codigos Modulares de una persona
'***     Creado por:        NSSE
'***     Maquina:           07SIST_08
'***     Fecha-Tiempo:      31/05/2001 09:10:04 AM
'***     Ultima Modificacion: Creacion
'*****************************************************************************************

Public Function RecuperaProductoTasaInteres(ByVal psCtaCod As String, ByVal pnPrdTasaInteres As ColocLineaCredTasas) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaProductoTasaInteres
    sSql = "Select * from ProductoTasaInteres Where cCtaCod = '" & psCtaCod & "' AND nPrdTasaInteres = " & pnPrdTasaInteres
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaProductoTasaInteres = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaProductoTasaInteres:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaProducto(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaProducto
    sSql = "Select * from Producto where cCtacod = '" & psCtaCod & "' "
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaProducto = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaProducto:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaColocacCred(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaColocacCred
    'sSql = "Select * From ColocacCred Where cCtaCod = '" & psCtaCod & "'"'JOEP20200321 Comento Mejora Reprogramacion
    sSql = "exec stp_sel_ObtieneDatosColocacCred '" & psCtaCod & "'" 'JOEP20200321 Add Mejora Reprogramacion
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaColocacCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaColocacCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Sub ActualizaBloqueoCred(ByVal psCtaCod As String, ByVal pnBloqueo)
Dim oBase As DCOMCredActBD

    Set oBase = New DCOMCredActBD
    Call oBase.dUpdateColocacCred(psCtaCod, , , , , , , , , , , , , , , , , , , , , , , , , pnBloqueo)
    Set oBase = Nothing
    
End Sub

Public Function RecuperaColocacCredCampos(ByVal psCtaCod As String, ByVal psCampo As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaColocacCred
    sSql = "Select " & psCampo & " from ColocacCred where cCtacod = '" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaColocacCredCampos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaColocacCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function


Public Function RecuperaSugerencia(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaSugerencia

'    sSQL = "Select ISNULL(CE.cTipoGasto,'F') as cTipoGasto,CE.nPrdEstado, CE.cDescripcion, PRD.nTransacc, C.nColocCalendCod, CN2.cConsDescripcion as cTipoCredDescrip, C.nColocDestino, CN.cConsDescripcion as cDestinoDescripcion, P.cPersCod, P.cPersNombre, "
'    sSQL = sSQL & " DNI=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & gPersIdDNI & "'), "
'    sSQL = sSQL & " RUC=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & gPersIdRUC & "'), "
'    sSQL = sSQL & " CE.nProxMes,CE.nCalendDinamico, CE.nMonto, CE.nCuotas, CE.nPlazo, CE.dPrdEstado,  CE.nPeriodoFechaFija, CE.nPeriodoGracia, CE.nPlazo, CE.nTipoGracia, CE.nTipoDesembolso, "
'    sSQL = sSQL & " cAnalista = (select Top 1 cPersNombre from persona where cPersCod = (Select cPersCod From ProductoPersona Where cCtaCod = C.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ")), "
'    sSQL = sSQL & " nMontoSol = (Select nMonto from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstSolic & "), "
'    sSQL = sSQL & " nCuotasSol = (Select nCuotas from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstSolic & "),"
'    sSQL = sSQL & " nPlazoSol = (Select nPlazo from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstSolic & "),"
'    sSQL = sSQL & " nTasaComp = (Select nTasaInteres from ProductoTasaInteres Where cCtaCod = C.cCtaCod And nPrdTasaInteres = " & gColocLineaCredTasasIntCompNormal & "),"
'    sSQL = sSQL & " nTasaGracia = (Select nTasaInteres from ProductoTasaInteres Where cCtaCod = C.cCtaCod And nPrdTasaInteres = " & gColocLineaCredTasasIntGracia & "),"
'    sSQL = sSQL & " cLineaCred = (Select cLineaCred from Colocaciones Where cCtaCod = C.cCtaCod), "
'    sSQL = sSQL & " C.bMiVivienda, C.bCuotaCom, C.nNroCalPar, ISNULL(PT.nTasaInteres,0) as nTasaMora, ISNULL(PT2.nTasaInteres,0) as nTasaICVen "
'    sSQL = sSQL & " From ColocacCred C INNER JOIN ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
'    sSQL = sSQL & "                    INNER JOIN " & gConsPersona & "Persona P ON PP.cPersCod = P.cPersCod"
'    sSQL = sSQL & "                    INNER JOIN ColocacEstado CE ON C.cCtaCod = CE.cCtaCod AND CE.nPrdEstado in (" & COMDConstantes.gColocEstSug & "," & gColocEstSolic & ")"
'    sSQL = sSQL & "                    INNER JOIN Constante CN ON C.nColocDestino = CN.nConsValor AND CN.nConsCod = " & Trim(Str(COMDConstantes.gColocDestino))
'    sSQL = sSQL & "                    INNER JOIN Constante CN2 ON convert(int,substring(C.cCtaCod,6,3)) = CN2.nConsValor AND CN2.nConsCod = " & COMDConstantes.gProducto
'    sSQL = sSQL & "                    INNER JOIN Producto PRD ON C.cCtaCod = PRD.cCtaCod"
'    sSQL = sSQL & "                    LEFT JOIN ProductoTasaInteres PT ON C.cCtaCod = PT.cCtaCod AND PT.nPrdTasaInteres = 3 "
'    sSQL = sSQL & "                    LEFT JOIN ProductoTasaInteres PT2 ON C.cCtaCod = PT2.cCtaCod AND PT2.nPrdTasaInteres = 6 "
'    sSQL = sSQL & " WHERE C.cCtaCod = '" & psCtaCod & "' Order By CE.nPrdEstado DESC"

'    sSql = " Select ISNULL(CE.cTipoGasto,'F') as cTipoGasto,CE.nPrdEstado, CE.cDescripcion, PRD.nTransacc, C.nColocCalendCod, CN2.cConsDescripcion as cTipoCredDescrip, C.nColocDestino, isnull(C.nExpoAntMax,0)nExpoAntMax,CN.cConsDescripcion as cDestinoDescripcion, P.cPersCod, P.cPersNombre, "
'    sSql = sSql & " DNI=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & COMDConstantes.gPersIdDNI & "'), "
'    sSql = sSql & " RUC=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & COMDConstantes.gPersIdRUC & "'), "
'    sSql = sSql & " CE.nProxMes,CE.nCalendDinamico, CE.nMonto, CE.nCuotas, CE.nPlazo, CE.dPrdEstado,  CE.nPeriodoFechaFija, CE.nPeriodoGracia, CE.nPlazo, CE.nTipoGracia, CE.nTipoDesembolso, "
'    sSql = sSql & " cAnalista = (select Top 1 cPersNombre from persona where cPersCod = (Select cPersCod From ProductoPersona Where cCtaCod = C.cCtaCod AND nPrdPersRelac = " & COMDConstantes.gColRelPersAnalista & ")), "
'    sSql = sSql & " nMontoSol = (Select nMonto from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & COMDConstantes.gColocEstSolic & "), "
'    sSql = sSql & " nCuotasSol = (Select nCuotas from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & COMDConstantes.gColocEstSolic & "),"
'    sSql = sSql & " nPlazoSol = (Select nPlazo from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & COMDConstantes.gColocEstSolic & "),"
'    sSql = sSql & " nTasaComp = (Select nTasaInteres from ProductoTasaInteres Where cCtaCod = C.cCtaCod And nPrdTasaInteres = " & COMDConstantes.gColocLineaCredTasasIntCompNormal & "),"
'    sSql = sSql & " nTasaGracia = (Select nTasaInteres from ProductoTasaInteres Where cCtaCod = C.cCtaCod And nPrdTasaInteres = " & COMDConstantes.gColocLineaCredTasasIntGracia & "),"
'    sSql = sSql & " cLineaCred = (Select cLineaCred from Colocaciones Where cCtaCod = C.cCtaCod), "
'    'CMACICA_CSTS - 26/11/2003 --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'    'ARCV 30-12-2006
'    'ssql = ssql & " nPersFIDIngCli = (SELECT PFD.nPersIngCli FROM COLOCFTEINGRESO CFI INNER JOIN PERSFIDEPENDIENTE PFD ON CFI.cNumFuente = PFD.cNumFuente Where CFI.cCtaCod = C.cCtaCod AND PFD.dPersEval = ( SELECT MAX(PFD2.dPersEval) FROM COLOCFTEINGRESO CFI2 INNER JOIN PERSFIDEPENDIENTE PFD2 ON CFI2.cNumFuente = PFD2.cNumFuente WHERE CFI2.cCtaCod = CFI.cCtaCod )),"
'    'ssql = ssql & " cPersFIMoneda = (SELECT     PF.cPersFIMoneda FROM COLOCFTEINGRESO CFI INNER JOIN PERSFTEINGRESO PF ON CFI.cNumFuente = PF.cNumFuente Where CFI.cCtaCod = C.cCtaCod),"
'    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'    sSql = sSql & " C.bMiVivienda, C.bCuotaCom, C.nNroCalPar, ISNULL(PT.nTasaInteres,0) as nTasaMora, ISNULL(PT2.nTasaInteres,0) as nTasaICVen "
'    '   CIUU del Credito   ****
'    sSql = sSql & ", cPersCIIU=ISNULL(P.cPersCIIU,'') "
'    '**** Manejo de nuevas opciones de Gracia
'    sSql = sSql & ", nDiaFijo2=ISNULL(CE.nPeriodoFechaFija2,0), bIncremGraciaCap=CONVERT(int,ISNULL(bIncremGraciaCap,0)) "
'    '***************
'    '***** Manejo de Creditos VAC
'    sSql = sSql & ", bVAC=CONVERT(int,ISNULL(C.bVAC,0))"
'    '***************
'    '**DAOR 20061216, Numero de Consultas a la Central de Riesgo (Certicom)
'    sSql = sSql & ", C.nNumConCer"
'    '*******************************************************************
'    '**DAOR 20071207, Personeria ***************************************
'    sSql = sSql & ", P.nPersPersoneria, isnull(PPD.cPersCod,'') as cPersRepDesgrav "
'    '*******************************************************************
'    sSql = sSql & " From ColocacCred C INNER JOIN ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
'    sSql = sSql & "                    INNER JOIN " & gConsPersona & "Persona P ON PP.cPersCod = P.cPersCod"
'    sSql = sSql & "                    INNER JOIN ColocacEstado CE ON C.cCtaCod = CE.cCtaCod AND CE.nPrdEstado in (" & COMDConstantes.gColocEstSug & "," & COMDConstantes.gColocEstSolic & ")"
'    sSql = sSql & "                    INNER JOIN Constante CN ON C.nColocDestino = CN.nConsValor AND CN.nConsCod = " & Trim(Str(COMDConstantes.gColocDestino))
'    sSql = sSql & "                    INNER JOIN Constante CN2 ON convert(int,substring(C.cCtaCod,6,3)) = CN2.nConsValor AND CN2.nConsCod = " & COMDConstantes.gProducto
'    sSql = sSql & "                    INNER JOIN Producto PRD ON C.cCtaCod = PRD.cCtaCod"
'    sSql = sSql & "                    LEFT JOIN ProductoTasaInteres PT ON C.cCtaCod = PT.cCtaCod AND PT.nPrdTasaInteres = 3 "
'    sSql = sSql & "                    LEFT JOIN ProductoTasaInteres PT2 ON C.cCtaCod = PT2.cCtaCod AND PT2.nPrdTasaInteres = 6 "
'    '**DAOR 20071207 ***************************************************
'    sSql = sSql & "                    left join ProductoPersona PPD on C.cCtaCod=PPD.cCtaCod and PPD.cSgrDsg='S' "
'    '*******************************************************************
'    sSql = sSql & " WHERE C.cCtaCod = '" & psCtaCod & "' Order By CE.dPrdEstado DESC"

    '*** PEAC 20080412
    sSql = "exec stp_sel_RecuperaSugerencia '" & psCtaCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaSugerencia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaSugerencia:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaDatosDesembolso(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset

    On Error GoTo ErrorRecuperaDatosDesembolso
    '**Comentado por DAOR 20070831
'    sSql = "Select CC.nNroCalen, C.cLineaCred, L.cDescripcion, P.cPersCod, P.cPersNombre, P.cPersDireccDomicilio,CN.cConsDescripcion cMoneda,  "
'    sSql = sSql & " CC.nNroproxDesemb, nPrestamo = (Select nMonto From ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & COMDConstantes.gColocEstAprob & "),"
'    sSql = sSql & " nMontoADesemb = (Select CD.nMonto from ColocCalenddet CD "
'    sSql = sSql & "                 where CD.cCtaCod = C.cCtaCod and CD.nColocCalendApl=0 and nPrdConceptoCod=1000 and nCuota = CC.nNroproxDesemb "
'    sSql = sSql & "                 and CD.nnrocalen=(Select MAX(nNroCalen) From ColocCalendario Where cCtaCod = CD.cCtaCod AND nColocCalendApl=0) ),  "
'    sSql = sSql & " DNI=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & COMDConstantes.gPersIdDNI & "'), "
'    sSql = sSql & " RUC=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & COMDConstantes.gPersIdRUC & "'), "
'    sSql = sSql & " nTotalDesemb=(Select Count(cCtaCod) from ColocCalendario Where cCtaCod = C.cCtaCod AND nNroCalen = CC.nNroCalen AND nColocCalendApl = " & COMDConstantes.gColocCalendAplDesembolso & "), "
'    sSql = sSql & " CC.bMiVivienda, CC.bCuotaCom ,P.nPersPersoneria"
'    'ARCV 24-01-2007
'    sSql = sSql & " ,nMontoPrimaTotal=ISNULL(nMontoPrimaTotal,0)"
'    '------------
'    sSql = sSql & " from Colocaciones C Inner Join ColocLineaCredito L ON C.cLineaCred = L.cLineaCred "
'    sSql = sSql & "                       Inner Join ColocacCred CC ON C.cCtaCod = CC.cCtaCod"
'    sSql = sSql & "                       Inner Join ProductoPersona PP ON PP.cCtaCod = C.cCtaCod AND PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
'    sSql = sSql & "                       Inner Join Persona P ON PP.cPersCod = P.cPersCod"
'    sSql = sSql & "                       Inner Join Constante CN ON CN.nConsValor = Convert(int,substring(C.cCtaCod,9,1)) And CN.nConsCod = " & COMDConstantes.gMoneda
'    'ARCV 24-01-2007
'    sSql = sSql & " INNER JOIN ColocGarantia CG ON CG.cCtaCod = C.cCtaCod " & _
'                  " LEFT JOIN GarantRealTasacion GRT ON CG.cNumGarant = GRT.cNumGarant  " & _
'                  " LEFT JOIN GarantPoliza GP ON GRT.cNumGarant = GP.cNumGarant AND GP.dTasacion = GRT.dTasacion AND GP.dTasacion =  (SELECT MAX(dTasacion)FROM GarantPoliza WHERE cNumGarant=CG.cNumGarant) AND GP.nEstado=1 " & _
'                  " LEFT JOIN Poliza Pol ON Pol.cNumPoliza = GP.cNumPoliza  AND Pol.bPolizaExterna=0 "
'    '--------------
'    sSql = sSql & " WHERE C.cCtaCod = '" & psCtaCod & "'"

    '**DAOR 20070831, se corrigió la parte que obtiene el Monto de la Prima de la Poliza
'    sSql = "Select CC.nNroCalen, C.cLineaCred, L.cDescripcion, P.cPersCod, P.cPersNombre, P.cPersDireccDomicilio,CN.cConsDescripcion cMoneda,  "
'    sSql = sSql & " CC.nNroproxDesemb, nPrestamo = (Select nMonto From ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & COMDConstantes.gColocEstAprob & "),"
'    sSql = sSql & " nMontoADesemb = (Select CD.nMonto from ColocCalenddet CD "
'    sSql = sSql & "                 where CD.cCtaCod = C.cCtaCod and CD.nColocCalendApl=0 and nPrdConceptoCod=1000 and nCuota = CC.nNroproxDesemb "
'    sSql = sSql & "                 and CD.nnrocalen=(Select MAX(nNroCalen) From ColocCalendario Where cCtaCod = CD.cCtaCod AND nColocCalendApl=0) ),  "
'    sSql = sSql & " DNI=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & COMDConstantes.gPersIdDNI & "'), "
'    sSql = sSql & " RUC=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & COMDConstantes.gPersIdRUC & "'), "
'    sSql = sSql & " nTotalDesemb=(Select Count(cCtaCod) from ColocCalendario Where cCtaCod = C.cCtaCod AND nNroCalen = CC.nNroCalen AND nColocCalendApl = " & COMDConstantes.gColocCalendAplDesembolso & "), "
'    sSql = sSql & " CC.bMiVivienda, CC.bCuotaCom ,P.nPersPersoneria,"
'    sSql = sSql & " nMontoPrimaTotal=(select isnull(sum(nMontoPrimaTotal),0) from Poliza "
'    sSql = sSql & "                   Where cNumPoliza "
'    sSql = sSql & "                     in(select distinct GP.cNumPoliza from ColocGarantia CG inner join GarantRealTasacion GRT on CG.cNumGarant = GRT.cNumGarant "
'    sSql = sSql & "                             inner join GarantPoliza GP on GRT.cNumGarant = GP.cNumGarant and GP.dTasacion = GRT.dTasacion "
'    sSql = sSql & "                             and GP.dTasacion =(select max(dTasacion) from GarantPoliza where cNumGarant=CG.cNumGarant) and GP.nEstado=1 "
'    sSql = sSql & "                        Where CG.cCtaCod = C.cCtaCod "
'    sSql = sSql & "                       ) and bPolizaExterna=0 "
'    sSql = sSql & "                  ) "
'    sSql = sSql & " from Colocaciones C Inner Join ColocLineaCredito L ON C.cLineaCred = L.cLineaCred "
'    sSql = sSql & "                       Inner Join ColocacCred CC ON C.cCtaCod = CC.cCtaCod"
'    sSql = sSql & "                       Inner Join ProductoPersona PP ON PP.cCtaCod = C.cCtaCod AND PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
'    sSql = sSql & "                       Inner Join Persona P ON PP.cPersCod = P.cPersCod"
'    sSql = sSql & "                       Inner Join Constante CN ON CN.nConsValor = Convert(int,substring(C.cCtaCod,9,1)) And CN.nConsCod = " & COMDConstantes.gMoneda
'    sSql = sSql & " where C.cCtaCod = '" & psCtaCod & "'"
    
    '*** PEAC 20080117
    sSql = " exec stp_sel_RecuperaDatosDesembolso '" & psCtaCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    'Set RecuperaDatosDesembolso = oConecta.CargaRecordSet(sSql)
    Set rs = oConecta.CargaRecordSet(sSql)
    
    If rs.RecordCount > 1 Then
        If rs!nMontoPrimaTotal = 0 Then rs.MoveNext         'Si hay mas de una Garantia coger la que tiene Poliza en el caso de que exista..sino no hay problema
    End If
    Set RecuperaDatosDesembolso = rs
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
 
ErrorRecuperaDatosDesembolso:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function RecuperaDatosAprobacion(ByVal psCtaCod As String, Optional ByVal pnAgenciaCredEval As Integer = 0) As ADODB.Recordset 'JUEZ 20120917 Se agregó pnAgenciaCredEval
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaDatosAprobacion

'**Modficado por DAOR 20080109, Trabajar con procedimiento almacenado
'    sSql = "Select CC.cLineaCred as cLineaCredCod,  CE.nPrdEstado, C.nColocCalendCod, CN2.cConsDescripcion as cTipoCredDescrip, C.nColocDestino, CN.cConsDescripcion as cDestinoDescripcion, "
'    sSql = sSql & " ISNULL(CE.cTipogasto,'F') as cTipogasto, CE.nMonto, CE.nCuotas, CE.nPlazo, CE.nPeriodoGracia, CE.nPeriodoFechaFija, PRD.nTasaInteres, CE.nProxMes, CE.nCalendDinamico, CE.nTipoGracia, CE.nTipoDesembolso, CE.nPeriodoFechaFija2,bIncremGraciaCap=CONVERT(int,ISNULL(bIncremGraciaCap,0)),"
'    sSql = sSql & " cAnalista = (select Top 1 cPersNombre from persona where cPersCod = (Select cPersCod From ProductoPersona Where cCtaCod = C.cCtaCod AND nPrdPersRelac = " & COMDConstantes.gColRelPersAnalista & ")), "
'    sSql = sSql & " nMontoSol = (Select nMonto from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & COMDConstantes.gColocEstSolic & "), "
'    sSql = sSql & " nCuotasSol = (Select nCuotas from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & COMDConstantes.gColocEstSolic & "),"
'    sSql = sSql & " nPlazoSol = (Select nPlazo from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & COMDConstantes.gColocEstSolic & "),"
'    sSql = sSql & " nTasaComp = (Select nTasaInteres from ProductoTasaInteres Where cCtaCod = C.cCtaCod And nPrdTasaInteres = " & COMDConstantes.gColocLineaCredTasasIntCompNormal & "),"
'    sSql = sSql & " nTasaGracia = (Select nTasaInteres from ProductoTasaInteres Where cCtaCod = C.cCtaCod And nPrdTasaInteres = " & COMDConstantes.gColocLineaCredTasasIntGracia & "),"
'    sSql = sSql & " cLineaCred = (Select (cDescripcion +  space(150) + cLineaCred) from ColocLineaCredito Where cLineaCred = CC.cLineaCred), "
'    sSql = sSql & " P2.cPersNombre cFteIngreso, PF.cPersFIPersCod, PF.nPersTipFte, P.cPersCod, P.cPersNombre, P.cPersDireccDomicilio, "
'    sSql = sSql & " P.cPersCod, P.cPersNombre, P.cPersDireccDomicilio, "
'    sSql = sSql & " CN3.cConsDescripcion as cCondicionCred, CN4.cConsDescripcion as cCondicionOtraCred, C.nColocCondicion, C.nColocCondicion2, C.bRefCapInt, "
'    sSql = sSql & " nMaxCalend = (Select MAX(nNroCalen) From ColocCalendario Cal Where cCtaCod = C.cCtaCod), "
'    sSql = sSql & " C.bMiVivienda, C.bCuotaCom, "
'    sSql = sSql & " nTasaIntComp = (Select nTasaInteres from ProductoTasaInteres where cCtaCod = C.cCtaCod AND nPrdTasaInteres = " & COMDConstantes.gColocLineaCredTasasIntCompNormal & "), "
'    sSql = sSql & " nTasaIntMora = (Select nTasaInteres from ProductoTasaInteres where cCtaCod = C.cCtaCod AND nPrdTasaInteres = " & gColocLineaCredTasasIntMoratNormal & ") "
'    '-----Numero de Fuente de Ingreso
'    sSql = sSql & " ,cNumFuente = ISNULL(CF.cNumFuente,0),"
'    sSql = sSql & " nCuotaSug = isnull ((select sum(CA.nMonto) from ColocCalenddet CA inner join ColocacEstado CE  ON CA.cCtaCod=CE.cCtaCod"
'    sSql = sSql & "                      where  CA.cCtaCod = C.cCtaCod and  nColocCalendCod in (10,11,40,41) and nColocCalendApl=1 and nCuota=1),0),"
'    sSql = sSql & " dVenc"
'
'    'DAOR 20061216, Numero de Consultas a la Central de Riesgo (Certicom)
'    sSql = sSql & ", C.nNumConCer "
'    '*******************************************************************
'    sSql = sSql & ", CF.dPersEval " 'ARCV 30-12-2006
'
'    sSql = sSql & " From ColocacCred C "
'    sSql = sSql & " INNER JOIN Producto PRD ON C.cCtaCod = PRD.cCtaCod"
'    sSql = sSql & " INNER JOIN ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
'    sSql = sSql & " INNER JOIN Colocaciones CC ON CC.cCtaCod = C.cCtaCod"
'    sSql = sSql & " INNER JOIN ColocFteIngreso CF on CF.cCtaCod=PRD.cCtaCod AND CF.cNumFuente=(SELECT TOP 1 cNumFuente FROM ColocFteIngreso WHERE cCtaCod=C.cCtaCod )" 'ARCV 30-12-2006
'    sSql = sSql & " INNER JOIN PersFteIngreso PF ON PF.cPersCod = PP.cPersCod and PF.cNumFuente=CF.cNumFuente"
'    sSql = sSql & " INNER JOIN Persona P2 ON P2.cPersCod = PF.cPersFIPersCod"
'    sSql = sSql & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod"
'    sSql = sSql & " INNER JOIN ColocacEstado CE ON C.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = 2001"
'    sSql = sSql & " INNER JOIN Constante CN ON C.nColocDestino = CN.nConsValor AND CN.nConsCod = 3016"
'    sSql = sSql & " INNER JOIN Constante CN2 ON convert(int,substring(C.cCtaCod,6,3)) = CN2.nConsValor AND CN2.nConsCod = 1001"
'    sSql = sSql & " INNER JOIN Constante CN3 ON C.nColocCondicion = CN3.nConsValor AND CN3.nConsCod = 3015"
'    sSql = sSql & " INNER JOIN Constante CN4 ON C.nColocCondicion2 = CN4.nConsValor AND CN4.nConsCod = 3030"
'    sSql = sSql & " WHERE C.cCtaCod = '" & psCtaCod & "' And Prd.nPrdestado = " & COMDConstantes.gColocEstSug & " ORDER BY CE.dPrdEstado DESC"
    
    'JUEZ 20120917 *****************************************************************
    If pnAgenciaCredEval = 0 Then
        'madm 20100513 ----------------------------------------------------------------------------------------------
        '**ARLO20180712 ERS042 - 2018
        Set objProducto = New COMDCredito.DCOMCredito
        If objProducto.GetResultadoCondicionCatalogo("N0000096", Mid(psCtaCod, 6, 3)) Then
        'If Mid(psCtaCod, 6, 3) <> "703" Then
        '**ARLO20180712 ERS042 - 2018
            sSql = " exec stp_sel_RecuperaDatosAprobacion '" & psCtaCod & "'"
        Else
            sSql = " exec stp_sel_RecuperaDatosAprobacion_sinFTE '" & psCtaCod & "'"
        End If
        'end madm -----------------------------------------------------------------------------------------------------
    Else
        sSql = " exec stp_sel_RecuperaDatosAprobacion_sinFTE '" & psCtaCod & "'"
    End If
    'END JUEZ **********************************************************************
    
    Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
    Set RecuperaDatosAprobacion = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaDatosAprobacion:
    Err.Raise Err.Number, "Error En Proceso", Err.Description


End Function

' CMACICA_CSTS - 06/11/2003 ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function RecuperaDatosAprobados(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaDatosAprobados

    sSql = "Select C.nColocCalendCod, C.nTipoDesembolso, nMaxCalend = (  Select MAX(nNroCalen) From ColocCalendario Cal Where cCtaCod = C.cCtaCod), CC.cLineaCred as cLineaCredCod, "
    sSql = sSql & " CC.dVigencia, CC.nMontoCol, CC.nPlazo, cLineaCred = (  Select (cDescripcion +  space(150) + cLineaCred) From ColocLineaCredito Where cLineaCred = CC.cLineaCred), "
    sSql = sSql & " nMonto     = (  Select  CE.nMonto From  ColocacEstado CE Where  CE.cCtaCod = C.cCtaCod AND CE.nPrdEstado = " & COMDConstantes.gColocEstAprob & "), PRD.nPrdEstado, P.cPersCod, P.cPersNombre "
    sSql = sSql & " From ColocacCred C   INNER JOIN ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
    sSql = sSql & "                      INNER JOIN Colocaciones CC ON CC.cCtaCod = C.cCtaCod "
    sSql = sSql & "                      INNER JOIN Persona P ON PP.cPersCod = P.cPersCod "
    sSql = sSql & "                      INNER JOIN Producto PRD ON C.cCtaCod = PRD.cCtaCod "
    sSql = sSql & " WHERE C.cCtaCod = '" & psCtaCod & "'"

'    ssql = "Select CC.cLineaCred as cLineaCredCod, CC.dVigencia, CE.nPrdEstado, C.nColocCalendCod, CN2.cConsDescripcion as cTipoCredDescrip, C.nColocDestino, CN.cConsDescripcion as cDestinoDescripcion, "
'    ssql = ssql & " ISNULL(CE.cTipogasto,'F') as cTipogasto, CE.nMonto, CE.nCuotas, CE.nPlazo, CE.nPeriodoGracia, CE.nPeriodoFechaFija, PRD.nTasaInteres, CE.nProxMes, CE.nCalendDinamico, CE.nTipoGracia, CE.nTipoDesembolso, "
'    ssql = ssql & " cAnalista = (select Top 1 cPersNombre from persona where cPersCod = (Select cPersCod From ProductoPersona Where cCtaCod = C.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ")), "
'    ssql = ssql & " nMontoSol = (Select nMonto from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstSolic & "), "
'    ssql = ssql & " nCuotasSol = (Select nCuotas from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstSolic & "),"
'    ssql = ssql & " nPlazoSol = (Select nPlazo from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstSolic & "),"
'    ssql = ssql & " nTasaComp = (Select nTasaInteres from ProductoTasaInteres Where cCtaCod = C.cCtaCod And nPrdTasaInteres = " & gColocLineaCredTasasIntCompNormal & "),"
'    ssql = ssql & " nTasaGracia = (Select nTasaInteres from ProductoTasaInteres Where cCtaCod = C.cCtaCod And nPrdTasaInteres = " & gColocLineaCredTasasIntGracia & "),"
'    ssql = ssql & " cLineaCred = (Select (cDescripcion +  space(150) + cLineaCred) from ColocLineaCredito Where cLineaCred = CC.cLineaCred), "
'    ssql = ssql & " P2.cPersNombre cFteIngreso, PF.cPersFIPersCod, P.cPersCod, P.cPersNombre, P.cPersDireccDomicilio, "
'    ssql = ssql & " CN3.cConsDescripcion as cCondicionCred, CN4.cConsDescripcion as cCondicionOtraCred, C.nColocCondicion, C.nColocCondicion2, C.bRefCapInt, "
'    ssql = ssql & " nMaxCalend = (Select MAX(nNroCalen) From ColocCalendario Cal Where cCtaCod = C.cCtaCod), "
'    ssql = ssql & " C.bMiVivienda, C.bCuotaCom, "
'    ssql = ssql & " nTasaIntComp = (Select nTasaInteres from ProductoTasaInteres where cCtaCod = C.cCtaCod AND nPrdTasaInteres = " & gColocLineaCredTasasIntCompNormal & "), "
'    ssql = ssql & " nTasaIntMora = (Select nTasaInteres from ProductoTasaInteres where cCtaCod = C.cCtaCod AND nPrdTasaInteres = " & gColocLineaCredTasasIntMoratNormal & ") "
'    ssql = ssql & " From ColocacCred C INNER JOIN ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
'    ssql = ssql & "                    INNER JOIN Colocaciones CC ON CC.cCtaCod = C.cCtaCod "
'    ssql = ssql & "                    INNER JOIN PersFteIngreso PF ON PF.cPersCod = PP.cPersCod "
'    ssql = ssql & "                    INNER JOIN Persona P2 ON P2.cPersCod = PF.cPersFIPersCod "
'    ssql = ssql & "                    INNER JOIN " & gConsPersona & "Persona P ON PP.cPersCod = P.cPersCod"
'    ssql = ssql & "                    INNER JOIN ColocacEstado CE ON C.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstAprob
'    ssql = ssql & "                    INNER JOIN Constante CN ON C.nColocDestino = CN.nConsValor AND CN.nConsCod = " & Trim(Str(COMDConstantes.gColocDestino))
'    ssql = ssql & "                    INNER JOIN Constante CN2 ON convert(int,substring(C.cCtaCod,6,3)) = CN2.nConsValor AND CN2.nConsCod = " & COMDConstantes.gProducto
'    ssql = ssql & "                    INNER JOIN Constante CN3 ON C.nColocCondicion = CN3.nConsValor AND CN3.nConsCod = " & Trim(Str(COMDConstantes.gColocCredCondicion ))
'    ssql = ssql & "                    INNER JOIN Constante CN4 ON C.nColocCondicion2 = CN4.nConsValor AND CN4.nConsCod = " & Trim(Str(gColocCredCondicionOtra))
'    ssql = ssql & "                    INNER JOIN Producto PRD ON C.cCtaCod = PRD.cCtaCod"
'    ssql = ssql & " WHERE C.cCtaCod = '" & psCtaCod & "' And Prd.nPrdestado = " & gColocEstAprob

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosAprobados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaDatosAprobados:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function RecuperaSolicitud(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaSolicitud
    
    'By Capi 28102008 se modifico consulta para obtener motivo refinanciacion
    
'    sSql = "Select C.nColocDestino, C.nColocCondicion, C.nColocCondicion2, C.nColocCondicionProd, P.cPersCod, P.cPersNombre, "
'    sSql = sSql & " DNI=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & COMDConstantes.gPersIdDNI & "'), "
'    sSql = sSql & " RUC=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & COMDConstantes.gPersIdRUC & "'), "
'    sSql = sSql & " CE.nMonto, CE.nCuotas, CE.nPlazo, CE.dPrdEstado, "
'    sSql = sSql & " cAnalista = (Select cPersCod From ProductoPersona Where cCtaCod = C.cCtaCod AND nPrdPersRelac = " & COMDConstantes.gColRelPersAnalista & "),"
'    sSql = sSql & " CC.cPersCod as cPersConvenio, CC.cCodModular, CF.cNumFuente, CF.dPersEval dPersFIEvaluacion, C.bRefCapint "
'    sSql = sSql & " From ColocacCred C INNER JOIN ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
'    sSql = sSql & "                    INNER JOIN Producto Prd ON Prd.cCtaCod = C.cCtaCod "
'    sSql = sSql & "                    INNER JOIN " & gConsPersona & "Persona P ON PP.cPersCod = P.cPersCod"
'    sSql = sSql & "                    INNER JOIN ColocacEstado CE ON C.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & COMDConstantes.gColocEstSolic
'    sSql = sSql & "                    LEFT JOIN ColocacConvenio CC ON C.cCtaCod = CC.cCtaCod"
'    sSql = sSql & "                    LEFT join ColocFteIngreso CF ON C.cCtaCod = CF.cCtaCod "
'    sSql = sSql & " WHERE C.cCtaCod = '" & psCtaCod & "' AND Prd.nPrdEstado = " & COMDConstantes.gColocEstSolic

'WIOR 20140510 COMENTO ESTO Y AGREGO ENUN PROCEDIMIENTO ALMACENADO
'    'sSql = "Select C.nColocDestino, C.nColocCondicion, C.nColocCondicion2, C.nColocCondicionProd, P.cPersCod, P.cPersNombre, "
'    sSql = "Select isnull(C.IdCampana,0) IdCampana,C.nColocDestino, C.nColocCondicion, C.nColocCondicion2, C.nColocCondicionProd, P.cPersCod, P.cPersNombre, "
'    sSql = sSql & " DNI=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & COMDConstantes.gPersIdDNI & "'), "
'    sSql = sSql & " RUC=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & COMDConstantes.gPersIdRUC & "'), "
'    sSql = sSql & " CE.nMonto, CE.nCuotas, CE.nPlazo, CE.dPrdEstado, "
'    'JUEZ 20120917 *************************************************************************************
'    'sSql = sSql & " cAnalista = (Select cPersCod From ProductoPersona Where cCtaCod = C.cCtaCod AND nPrdPersRelac = " & COMDConstantes.gColRelPersAnalista & "),"
'    sSql = sSql & " cAnalista=R.cPersCod, UPPER(R.cUser) cUserAnalista, "
'    'END JUEZ ******************************************************************************************
'    sSql = sSql & " CC.cPersCod as cPersConvenio, CC.cCodModular, CF.cNumFuente, CF.dPersEval dPersFIEvaluacion, C.bRefCapint,CR.nMotivoRef,isnull(C.id_CasaCom,0) id_CasaCom,isnull(C.cVendedor,'') cVendedor   "  'MIOL 20130626 RQ133335 ,isnull(C.cVendedor,0) cVendedor
'    sSql = sSql & " From ColocacCred C INNER JOIN ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
'    'JUEZ 20120917 *************************************************************************************
'    sSql = sSql & "                    INNER JOIN ProductoPersona PP2 ON C.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & COMDConstantes.gColRelPersAnalista
'    sSql = sSql & "                    INNER JOIN RRHH R ON R.cPersCod = PP2.cPersCod "
'    'END JUEZ ******************************************************************************************
'    sSql = sSql & "                    INNER JOIN Producto Prd ON Prd.cCtaCod = C.cCtaCod "
'    sSql = sSql & "                    INNER JOIN " & gConsPersona & "Persona P ON PP.cPersCod = P.cPersCod"
'    sSql = sSql & "                    INNER JOIN ColocacEstado CE ON C.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & COMDConstantes.gColocEstSolic
'    sSql = sSql & "                    LEFT JOIN ColocacConvenio CC ON C.cCtaCod = CC.cCtaCod"
'    sSql = sSql & "                    LEFT join ColocFteIngreso CF ON C.cCtaCod = CF.cCtaCod "
'    sSql = sSql & "                    LEFT JOIN COLOCACREFINANC CR ON C.CCTACOD=CR.CCTACOD "
'    sSql = sSql & " WHERE C.cCtaCod = '" & psCtaCod & "' AND Prd.nPrdEstado = " & COMDConstantes.gColocEstSolic

    'WIOR 20140510 ********************************
    sSql = "EXEC stp_sel_RecuperaSolicitudCred '" & psCtaCod & "' "
    'WIOR FIN *************************************
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaSolicitud = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaSolicitud:
    Err.Raise Err.Number, "Recuperar Solicitud", Err.Description
End Function

Public Function ExisteGasto(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, _
        pnColocCalendApl As Integer, pnCuota As Integer, ByVal nPrdConceptoCod As Long) As Boolean
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset


    sSql = "Select nCuota From colocCalendDet Where cCtaCod = '" & psCtaCod & "' AND nNroCalen = " & pnNroCalen
    sSql = sSql & " AND nColocCalendApl = " & pnColocCalendApl & " AND nCuota = " & pnCuota & " AND nPrdConceptoCod = " & nPrdConceptoCod
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    If Not R.BOF And Not R.EOF Then
        ExisteGasto = True
    Else
        ExisteGasto = False
    End If
    R.Close
    Set R = Nothing


End Function

Public Function ExisteGarantia(ByVal psCtaCod As String, ByVal psNumGarant As String) As Boolean
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    On Error GoTo ErrorExisteGarantia
    sSql = "Select cNumGarant From ColocGarantia Where cCtaCod = '" & psCtaCod & "' AND cNumGarant = '" & psNumGarant & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    If Not R.BOF And Not R.EOF Then
        ExisteGarantia = True
    Else
        ExisteGarantia = False
    End If
    R.Close
    Set R = Nothing
    Exit Function

ErrorExisteGarantia:
                  Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

'*****************************************************************************************
'***     Rutina:           RecuperaDatosGarantiaCred
'***     Descripcion:      Recupera los datos de la solicitud del Credito
'                          a la Cual se le va a registrar sus Garantias.
'***     Modificado por:        NSSE
'***     Fecha-Tiempo:         11/06/2001 11:30:40 AM
'***     Ultima Modificacion: Lo Ultimo que se Modifico
'*****************************************************************************************

Public Function RecuperaDatosGarantiaCred(ByVal psCtaCod As String) As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String

    On Error GoTo ErrorRecuperaDatosGarantiaCred
    sSql = "Select isnull((select cConsDescripcion from constante where nConsCod=3033 and nConsValor=substring(cTpoProdCod,1,1)+'00'),'') cTipoProdDescrip,isnull(CN4.cConsDescripcion,'') as cTipoCredDescrip,CN3.cConsDescripcion as cSTipoProdDescrip, C.nColocDestino, CN.cConsDescripcion as cDestinoDescripcion , P.cPersCod, P.cPersNombre, CN2.cConsDescripcion as cMonedaDesc, "
    sSql = sSql & " DNI=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & COMDConstantes.gPersIdDNI & "'), "
    sSql = sSql & " CE.nMonto "
    sSql = sSql & " From ColocacCred C INNER JOIN ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
    'ALPA 20100605 B2*************************
    sSql = sSql & "                    INNER JOIN Colocaciones CO on C.cCtaCod = CO.cCtaCod "
    '*****************************************
    sSql = sSql & "                    INNER JOIN " & gConsPersona & "Persona P ON PP.cPersCod = P.cPersCod"
    sSql = sSql & "                    INNER JOIN ColocacEstado CE ON C.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = '" & COMDConstantes.gColocEstSolic & "' "
    sSql = sSql & "                    INNER JOIN Constante CN ON C.nColocDestino = CN.nConsValor AND CN.nConsCod = " & COMDConstantes.gColocDestino
    sSql = sSql & "                    INNER JOIN Constante CN2 ON convert(int,substring(C.cCtaCod,9,1)) = CN2.nConsValor AND CN2.nConsCod = " & COMDConstantes.gMoneda
    sSql = sSql & "                    INNER JOIN Constante CN3 ON convert(int,cTpoProdCod) = CN3.nConsValor AND CN3.nConsCod = " & COMDConstantes.gTpoProducto
    sSql = sSql & "                    LEFT JOIN Constante CN4 ON convert(int,cTpoCredCod) = CN4.nConsValor AND CN4.nConsCod = " & COMDConstantes.gTpoCredito

    sSql = sSql & " WHERE C.cCtaCod = '" & psCtaCod & "' "

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosGarantiaCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaDatosGarantiaCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'*****************************************************************************************
'***     Rutina:           RecuperaGarantiasCredito
'***     Descripcion:      Recupera las Garantias del Credito Registradas anteriormente
'***     Modificado por:        NSSE
'***     Fecha-Tiempo:         11/06/2001 11:31:46 AM
'***     Ultima Modificacion: Lo Ultimo que se Modifico
'*****************************************************************************************

Public Function EstadoActual(ByVal psCtaCod As String) As Integer
Dim R As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorEstadoActual
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = "Select nPrdEstado from Producto Where cCtaCod = '" & psCtaCod & "'"
    Set R = oConecta.CargaRecordSet(sSql)
    If Not R.BOF And Not R.EOF Then
        EstadoActual = R!nPrdEstado
    Else
        EstadoActual = vbNull
    End If
    R.Close
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorEstadoActual:
                  Err.Raise Err.Number, "Error En Proceso EstadoActual", Err.Description
End Function

Public Function GarantiaDeOtroCredito(ByVal psCtaCod As String, ByVal psNumGarant As String) As Boolean
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta

    sSql = " Select CG.cNumGarant "
    sSql = sSql & " from ColocacRefinanc CR Inner Join ColocGarantia CG ON CR.cCtaCodRef = CG.cCtaCod "
    sSql = sSql & " where CR.cCtaCod = '" & psCtaCod & "' And CG.cNumGarant = '" & psNumGarant & "'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set R = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing

    If R.RecordCount > 0 Then
        GarantiaDeOtroCredito = True
    Else
        GarantiaDeOtroCredito = False
    End If
    R.Close

End Function

Public Sub EliminarGarantia(ByVal psCtaCod As String, ByVal psNumGarant As String, _
     ByVal pnMontoGarantia As Double)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorEliminarGarantia
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.BeginTrans

'    'Si es garantia de otro credito por Refinanciacion
'    If Not GarantiaDeOtroCredito(psCtaCod, psNumGarant) Then
'        'JACA 20111011********************************************************************
'            If (pnGravament - pnMontoGarantia) < 0 Then
'                sSql = "Update Garantias set nGravament =0  where cNumgarant = '" & psNumGarant & "' "
'            Else
                sSql = "Update Garantias set nGravament = nGravament - " & Format(pnMontoGarantia, "#0.00") & " where cNumgarant = '" & psNumGarant & "' "
'            End If
        'JACA END**********************************************************************
        oConecta.ConexionActiva.Execute sSql
'    End If

    sSql = "DELETE ColocGarantia WHERE cCtaCod = '" & psCtaCod & "' AND cNumgarant = '" & psNumGarant & "' "
    oConecta.ConexionActiva.Execute sSql

    oConecta.ConexionActiva.CommitTrans
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Sub

ErrorEliminarGarantia:
                  Err.Raise Err.Number, "Error En Proceso EliminarGarantia", Err.Description


End Sub

Public Function RecuperaColocGarantia(ByVal psCtaCod As String, Optional poConex As ADODB.Connection) As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String

    On Error GoTo ErrorRecuperaColocGarantia
    sSql = "Select G.cTpoDoc, G.cNroDoc, G.nTpoGarantia, G.IdSupGarant, CG.cNumGarant, CG.cCtaCod, CG.nMoneda, CG.nGravado, ISNULL(G.nGravament,0) AS nGravament "
    sSql = sSql & " from ColocGarantia CG Inner Join Garantias G ON CG.cNumGarant = G.cNumGarant "
    sSql = sSql & " Where cCtaCod = '" & psCtaCod & "'"

   ' Set oConecta = New COMConecta.DCOMConecta
    'oConecta.AbreConexion
   ' If Not poConex Is Nothing Then
  '      oConecta.ConexionActiva = poConex
  '  End If
    If poConex Is Nothing Then
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set RecuperaColocGarantia = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
    Else
        Set RecuperaColocGarantia = poConex.Execute(sSql)
    End If

    'oConecta.CierraConexion
    'Set oConecta = Nothing
    Exit Function

ErrorRecuperaColocGarantia:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaMov(ByVal psCtaCod As String, ByVal psOpeCod As String) As Long
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    sSql = "Select nMovNro from MovCol Where cCtaCod = '" & psCtaCod & "'"
    sSql = sSql & " And cOpeCod = '" & psOpeCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    If R.RecordCount > 0 Then
        RecuperaMov = IIf(IsNull(R!nMovNro), 0, R!nMovNro)
    Else
        RecuperaMov = 0
    End If
    R.Close
    Set R = Nothing

End Function

Public Function RecuperaGarantiasCredito(ByVal psCtaCod As String) As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String

    On Error GoTo ErrorRecuperaGarantiasCredito
    
    '*** PEAC 20080515
'    sSql = "Select CG.cNumGarant, CG.nGravado, CASE WHEN CG.nMoneda='1' THEN 'SOLES' ELSE 'DOLAR' END nMoneda, G.nTpoGarantia, CN.cConsDescripcion cTpoGarDescripcion, "
'    sSql = sSql & " CASE WHEN G.nMoneda='1' THEN 'SOLES' ELSE 'DOLAR' END nMoneda , G.nTasacion, G.nRealizacion, G.nPorGravar, G.cDescripcion, P.cPersNombre, G.cTpoDoc, G.cNroDoc "
'    sSql = sSql & " From ColocGarantia CG Inner Join Garantias G ON CG.cNumGarant = G.cNumGarant "
'    sSql = sSql & "                       Inner Join Constante CN ON G.nTpoGarantia = CN.nConsvalor And CN.nConsCod = " & Trim(Str(COMDConstantes.gPersGarantia))
'    sSql = sSql & "                       Inner Join PersGarantia PG ON CG.cNumGarant = PG.cNumGarant AND PG.nRelacion = " & Trim(Str(COMDConstantes.gPersRelGarantiaTitular))
'    sSql = sSql & "                       Inner Join Persona P ON PG.cPersCod = P.cPersCod "
'    sSql = sSql & " WHERE CG.cCtaCod = '" & psCtaCod & "'"

    sSql = "exec stp_sel_RecuperaGarantiasCredito '" & psCtaCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaGarantiasCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaGarantiasCredito:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function GarantiaPerteneceACreditoAprobado(ByVal psCtaCod As String, ByVal psNumagarant As String) As Boolean
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String
Dim R As ADODB.Recordset

    sSql = "select P.nPrdestado from ColocGarantia CG Inner Join Producto P ON CG.cCtaCod = P.cCtaCod "
    sSql = sSql & " Where P.cCtaCod = '" & psCtaCod & "' AND CG.cNumgarant='" & psNumagarant & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion

    If R.RecordCount > 0 Then
        If R!nPrdEstado >= COMDConstantes.gColocEstAprob Then
            GarantiaPerteneceACreditoAprobado = True
        Else
            GarantiaPerteneceACreditoAprobado = False
        End If
    Else
        GarantiaPerteneceACreditoAprobado = False
    End If
End Function

Public Sub NuevaGarantia(ByVal psCtaCod As String, ByVal psNumGarant As String, psMoneda As String, ByVal pnGravado As Double)
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String
Dim bTran As Boolean
Dim dGar As DCOMGarantia
Dim R As ADODB.Recordset
Dim nGravado As Double

Dim nTpoGarantia As Integer
Dim cNroDoc As String

    On Error GoTo ErrorNuevaGarantia
    bTran = False


    Set dGar = New DCOMGarantia
    Set R = dGar.RecuperaGarantia(psNumGarant)
    Set dGar = Nothing
    If R.RecordCount > 0 Then
        nGravado = R!nGravament
        nTpoGarantia = R!nTpoGarantia
        cNroDoc = R!cNroDoc
    End If
    R.Close

    'Nuevo Registro en ColocGarantia
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.BeginTrans

    bTran = True
    sSql = "INSERT INTO ColocGarantia(cNumGarant, cCtaCod, nMoneda, nGravado, nEstado) "
    sSql = sSql & " VALUES('" & psNumGarant & "','" & psCtaCod & "'," & psMoneda & "," & Format(pnGravado, "#0.00") & ", 1)"
    Call oConecta.ConexionActiva.Execute(sSql)

    If nGravado = 0 Then
        sSql = "UPDATE Garantias SET "
        sSql = sSql & " nGravament = nGravament + " & Format(pnGravado, "#0.00") & ", nEstado = " & COMDConstantes.gPersGarantEstadoAsignado
        sSql = sSql & " WHERE cNumgarant = '" & psNumGarant & "'"
    Else
        sSql = "UPDATE Garantias SET "
        sSql = sSql & " nGravament = nGravament + " & Format(pnGravado, "#0.00") & ", nEstado = " & COMDConstantes.gPersGarantEstadoContabilizado
        sSql = sSql & " WHERE cNumgarant = '" & psNumGarant & "'"
    End If
    Call oConecta.ConexionActiva.Execute(sSql)
    
'    'ARCV 27-02-2007    (Bloquea Garantia)
'    If nTpoGarantia = 6 Then
'        Dim rsRet As New ADODB.Recordset
'        Dim rsTot As New ADODB.Recordset
'
'        'Llenar Recordset
'
'        With rsRet
'            'Crear RecordSet
'            .Fields.Append "Est", adInteger
'            .Fields.Append "Fecha Blq", adDate
'            .Fields.Append "Motivo", adVarChar, 100
'            .Fields.Append "Comentario", adVarChar, 100
'            .Fields.Append "cMovNro", adVarChar, 25
'            .Fields.Append "Flag", adChar, 1
'            .Open
'        End With
'
'        With rsTot
'            'Crear RecordSet
'            .Fields.Append "Est", adInteger
'            .Fields.Append "Fecha Blq", adDate
'            .Fields.Append "Motivo", adVarChar, 100
'            .Fields.Append "Comentario", adVarChar, 100
'            .Fields.Append "cMovNro", adVarChar, 25
'            .Fields.Append "Flag", adChar, 1
'            .Open
'            .AddNew
'
'        End With
'
'
'
'                .Fields("cCtaCod") = FECreditos.TextMatrix(i, 2)
'                .Fields("nSaldoCap") = CDbl(FECreditos.TextMatrix(i, 6))
'                .Fields("nDemanda") = IIf(FECreditos.TextMatrix(i, 3) = "SI", gColRecDemandaSi, gColRecDemandaNo)
'                .Fields("nDiasAtraso") = CInt(FECreditos.TextMatrix(i, 9))
'                .Fields("nNroCalen") = CInt(FECreditos.TextMatrix(i, 10))
'                .Fields("nPrdEstado") = CInt(FECreditos.TextMatrix(i, 11))
'
'        Set rsRet = grdRetiro.GetRsNew()
'        Set rsTot = grdTotal.GetRsNew()
'
'        Set rsRetTmp = rsRet
'        Set rsTotTmp = rsTot
        
'        Set clsMov = New COMNContabilidad.NCOMContFunciones
'        sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'        Set clsMov = Nothing
'        sCuenta = txtCuenta.NroCuenta
'        Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
'        clsMant.ActualizaBloqueos sCuenta, rsRet, rsTot, sMovNro, nEstado
'        Set clsMant = Nothing
'    End If
'    '-------
    Call oConecta.ConexionActiva.CommitTrans
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorNuevaGarantia:
    If bTran Then
        Call oConecta.ConexionActiva.RollbackTrans
        Set oConecta = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso NuevaGarantia", Err.Description
End Sub


Public Sub ActualizaGarantias(ByVal psCtaCod As String, ByVal psNumGarant As String, psMoneda As String, ByVal pnGravado As Double, ByVal pnGravadoAnt As Double)
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String
Dim bTran As Boolean

    On Error GoTo ErrorNuevaGarantia
    bTran = False
    'Actualiza Registro en ColocGarantia
    sSql = "UPDATE ColocGarantia SET "
    sSql = sSql & " nGravado = " & Format(pnGravado, "#0.00")
    sSql = sSql & " WHERE cNumGarant" & psNumGarant & "' AND cCtaCod = '" & psCtaCod & "' "
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.BeginTrans
    bTran = True
    Call oConecta.ConexionActiva.Execute(sSql)

    sSql = "UPDATE Garantias SET "
    sSql = sSql & " nGravament = nGravament - " & Format(pnGravadoAnt, "#0.00")
    sSql = sSql & " WHERE cNumgarant = '" & psNumGarant & "'"
    Call oConecta.ConexionActiva.Execute(sSql)
    sSql = "UPDATE Garantias SET "
    sSql = sSql & " nGravament = nGravament + " & Format(pnGravado, "#0.00") & ", "
    sSql = sSql & " nEstado = " & COMDConstantes.gPersGarantEstadoAsignado
    sSql = sSql & " WHERE cNumgarant = '" & psNumGarant & "'"
    Call oConecta.ConexionActiva.Execute(sSql)
    Call oConecta.ConexionActiva.CommitTrans
    Call oConecta.CierraConexion
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorNuevaGarantia:
    If bTran Then
        Call oConecta.ConexionActiva.RollbackTrans
        Set oConecta = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso NuevaGarantia", Err.Description
End Sub
Public Function RecuperaAgencia(ByVal psAge As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaAgencia
    sSql = "Select * from Agencias Where cAgeCod = '" & psAge & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaAgencia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaAgencia:
    Err.Raise Err.Number, "Error En Proceso", Err.Description


End Function

Public Function RecuperaCreditosVigentesCtas(ByVal psPersCod As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCreditosVigentes
    sSql = "Select Prd.cCtaCod "
    sSql = sSql & " FROM Producto Prd Inner Join ProductoPersona PP ON Prd.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
    sSql = sSql & " WHERE PP.cPersCod = '" & psPersCod & "' And Prd.nPrdEstado  in (" & gColocEstRefMor & "," & COMDConstantes.gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & ")"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditosVigentesCtas = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCreditosVigentes:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

'ALPA 20111209
'Public Function RecuperaCreditosVigentes(ByVal psPersCod As String, Optional ByVal psCtaCod As String = "", _
'    Optional pEstadosEspecificos As Variant = Nothing, Optional ByVal psMoneda As String = "") As ADODB.Recordset
Public Function RecuperaCreditosVigentes(ByVal psPersCod As String, Optional ByVal psCtaCod As String = "", _
    Optional pEstadosEspecificos As Variant = Nothing, Optional ByVal psMoneda As String = "", Optional ByVal bLeasing As Boolean = False) As ADODB.Recordset
Dim sSql As String
Dim sPrdEstados As String 'marg20170224
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
'EJVG20120710 Se agregó cAgeCodAct
    On Error GoTo ErrorRecuperaCreditosVigentes
'COMENTADO POR MARG20170224---
'''    sSql = "Select Prd.cCtaCod, CC.nDiasAtraso, Prd.nSaldo, CN.cConsDescripcion cEstado, Prd.nPrdEstado,"
'''    sSql = sSql & " nPrestamo = (Select nMonto From ColocacEstado Where cCtaCod = Prd.cCtaCod And nPrdEstado = " & COMDConstantes.gColocEstAprob & "), C.dVigencia, Prd.nTasaInteres, ISNULL(C.cAgeCodAct,'') cAgeCodAct "
'''    sSql = sSql & " FROM Producto Prd Inner Join ProductoPersona PP ON Prd.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
'''    sSql = sSql & "                   Inner Join Persona P ON PP.cPersCod = P.cPersCod "
'''    sSql = sSql & "                   Inner Join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod "
'''    sSql = sSql & "                   Inner Join Colocaciones C ON C.cCtaCod = Prd.cCtaCod "
'''    sSql = sSql & "                   Inner Join Constante CN ON CN.nConsValor = Prd.nPrdEstado And CN.nConsCod = " & COMDConstantes.gColocEstado
'''    sSql = sSql & " WHERE P.cPersCod = '" & psPersCod & "' "
'''    'ALPA 20111209**********
'''    If bLeasing = True Then
'''        sSql = sSql & "     AND SUBSTRING(Prd.cCtaCod,6,3) in ('515','516') "
'''    End If
'''    '***********************
'''    If psMoneda <> "" Then
'''        sSql = sSql & " AND SUBSTRING(Prd.cCtaCod,9,1) = '" & psMoneda & "' "
'''    End If
'''    If psCtaCod <> "" Then
'''        sSql = sSql & " AND Prd.cCtaCod <> '" & psCtaCod & "' "
'''    End If
'''
'''    If IsArray(pEstadosEspecificos) Then
'''        If UBound(pEstadosEspecificos) <> 0 Then
'''            sSql = sSql & " AND Prd.nPrdEstado in ("
'''            For i = 0 To UBound(pEstadosEspecificos)
'''                sSql = sSql & pEstadosEspecificos(i) & ","
'''            Next i
'''            sSql = Mid(sSql, 1, Len(sSql) - 1)
'''            sSql = sSql & ")"
'''        End If
'''    End If
'END MARG20170224---

'MARG20170224--
If IsArray(pEstadosEspecificos) Then
    If UBound(pEstadosEspecificos) <> 0 Then
        For i = 0 To UBound(pEstadosEspecificos)
            sPrdEstados = sPrdEstados & pEstadosEspecificos(i) & ","
        Next i
        sPrdEstados = Mid(sPrdEstados, 1, Len(sPrdEstados) - 1)
    End If
End If

sSql = "exec sp_sel_RecuperaCreditosVigentes '" & psPersCod & "','" & psCtaCod & "','" & psMoneda & "'," & IIf(bLeasing, 1, 0) & ",'" & sPrdEstados & "'"
'END MARG20170224---

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditosVigentes = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCreditosVigentes:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaConsultaCred(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorRecuperaConsultaCred
    
''COmento JOEP20201202 Segemntacion Riesgo
'    '**DAOR 20080409, Se aumentó el campo P.dPrdEstado
'    'ALPA 20091202****************************************************************
'    'sSql = "Select ISNULL(CC.nCalendDinamico,0) as nCalendDinamico, P.nPrdEstado, P.dPrdEstado, ISNULL(CC.bMiVivienda,0) as bMiVivienda, ISNULL(CC.bCuotaCom,0) as bCuotaCom, CC.nDiasAtraso, PFI.cPersNombre cFteIngreso, L.cDescripcion cLineaDesc, PersAna.cPersNombre cAnalista, PersApod.cPersNombre cApoderado, CN.cConsDescripcion cCondicion, Lin.nTasaIni as nTasaGracia, "
'    sSql = "Select ISNULL(CC.nCalendDinamico,0) as nCalendDinamico, P.nPrdEstado, P.dPrdEstado, ISNULL(CC.bMiVivienda,0) as bMiVivienda, ISNULL(CC.bCuotaCom,0) as bCuotaCom, CC.nDiasAtraso, PFI.cPersNombre cFteIngreso, L.cDescripcion cLineaDesc, PersAna.cPersNombre cAnalista, PersApod.cPersNombre cApoderado, CN.cConsDescripcion cCondicion, (Select nTasaInteres From ProductoTasaInteres Where nPrdTasaInteres  = 1 AND cCtaCod = P.cCtaCod) as nTasaGracia, "
'    '*****************************************************************************
'    sSql = sSql & " CN2.cConsDescripcion cDestino, P.nTasaInteres, CN3.cConsDescripcion cTipoCuota, C.dVigencia,"
'    sSql = sSql & " nNota = (select nColocNota  From ColocCalificacionAnalista Where cCtaCod = P.cCtaCod And dColocNotaFecha = (select Max(dColocNotaFecha) from ColocCalificacionAnalista Where cCtaCod = P.cCtaCod)), "
'    sSql = sSql & " CE.nMonto nMontoSol, CE.dPrdEstado dFecSol, CE.nCuotas nCuotasSol, CE.nPlazo nPlazoSol,"
'    sSql = sSql & " CE2.nMonto nMontoSug, CE2.dPrdEstado dFecSug, CE2.nCuotas nCuotasSug, CE2.nPlazo nPlazoSug, CE2.nPeriodoGracia nPeriodoGraciaSug,"
'    sSql = sSql & " CE3.nMonto nMontoApr, CE3.dPrdEstado dFecApr, CE3.nCuotas nCuotasApr, CE3.nPlazo nPlazoApr, CE3.nPeriodoGracia nPeriodoGraciaApr, CN8.cConsDescripcion cTipoGracia,"
'    sSql = sSql & " CN4.cConsDescripcion cMotivoRech,CE4.cDescripcion cGlosa, CE5.dPrdEstado dFecCancel, CE6.dPrdEstado dFecJud, CC.cMetLiquidacion, CC.cProtesto," ''FRHU 20130914 CE4.cDescripcion
'    sSql = sSql & " CE7.nPrdEstado nEstRefin, CC.bCargoAuto, CN5.cConsDescripcion cEstActual, "
'    'MAVM 20100611 BAS II ***
'    'sSql = sSql & " CN6.cConsDescripcion as cTipoCredDescrip, "
'    sSql = sSql & " isnull(CN4B.nConsValor,0) nTipoProdCod, ISNULL(CN4B.cConsDescripcion,'') as cTipoProdDescrip, isnull(CN5B.nConsValor,0) nTipoCredCod, ISNULL(CN5B.cConsDescripcion,'') as cTipoCredDescrip, "
'    '***
'    sSql = sSql & " CN7.cConsDescripcion cTipoDesemb, CC.nNroProxDesemb, "
'    sSql = sSql & " nCuotaSug = (Select sum(nMonto) from ColocCalendDet Cal Where Cal.cCtaCod = P.cCtaCod AND Cal.nNrocalen =1 AND Cal.nColocCalendApl=1 AND Cal.nCuota=1) " & " ,"
'    sSql = sSql & " nCuotaApr = (Select sum(nMonto) from ColocCalendDet Cal Where Cal.cCtaCod = P.cCtaCod AND Cal.nNrocalen =(Select nNroCalen From ColocacCred Where cCtaCod = Cal.cCtaCod) AND Cal.nColocCalendApl=1 AND Cal.nCuota=2), "
'    sSql = sSql & " nTasaComp = (Select nTasaInteres From ProductoTasaInteres Where nPrdTasaInteres  = 1 AND cCtaCod = P.cCtaCod), "
'    sSql = sSql & " nTasaCompVen = (Select nTasaInteres From ProductoTasaInteres Where nPrdTasaInteres  = 6 AND cCtaCod = P.cCtaCod), "
'    sSql = sSql & " dFecEEFF = isnull((select top 1 dfeceeff from persfiindependiente where cnumfuente=pf.cnumfuente order by dperseval desc),''),"
'    sSql = sSql & " nTasaMor = (Select nTasaInteres From ProductoTasaInteres Where nPrdTasaInteres  = 3 AND cCtaCod = P.cCtaCod) "
'    If CInt(Mid(psCtaCod, 6, 3)) = gColConsuDctoPlan Then
'        sSql = sSql & ", Conv.cCodModular, PConv.cPersNombre cConvenio"
'    End If
'    sSql = sSql & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod AND PP.nPrdpersRelac = " & COMDConstantes.gColRelPersTitular
'    sSql = sSql & "                 Left Join ColocFteIngreso FI ON  FI.cCtaCod = P.cCtaCod"
'    sSql = sSql & "                 Left Join PersFteIngreso PF ON PF.cNumFuente =  FI.cNumFuente "
'    sSql = sSql & "                 Left Join Persona PFI ON PFI.cPersCod = PF.cPersFIPersCod"
'    sSql = sSql & "                 Inner Join Colocaciones C ON C.cCtaCod = P.cCtaCod "
'    sSql = sSql & "                 Left Join ColocLineaCredito L ON L.cLineaCred = C.cLineaCred"
'    sSql = sSql & "                 Left Join ProductoPersona PP2 ON PP2.cCtaCod = P.cCtaCod AND PP2.nPrdPersRelac = " & COMDConstantes.gColRelPersAnalista
'    sSql = sSql & "                 Left Join Persona PersAna ON PersAna.cPersCod = PP2.cPersCod"
'    sSql = sSql & "                 Left Join ProductoPersona PP3 ON PP3.cCtaCod = P.cCtaCod AND PP3.nPrdPersRelac = " & COMDConstantes.gColRelPersApoderado
'    sSql = sSql & "                 Left Join Persona PersApod ON PersApod.cPersCod = PP3.cPersCod"
'    sSql = sSql & "                 Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod"
'    sSql = sSql & "                 Left Join Constante CN ON CN.nConsValor = CC.nColocCondicion AND CN.nConsCod = " & COMDConstantes.gColocCredCondicion
'    sSql = sSql & "                 Left Join Constante CN2 ON CN2.nConsValor = CC.nColocDestino AND CN2.nConsCod = " & COMDConstantes.gColocDestino
'    sSql = sSql & "                 Left Join Constante CN3 ON CN3.nConsValor = CC.nColocCalendCod AND CN3.nConsCod = " & COMDConstantes.gColocTipoCalend
'    sSql = sSql & "                 Left Join ColocacEstado CE ON CE.cCtaCod = P.cCtaCod AND CE.nPrdEstado = " & COMDConstantes.gColocEstSolic
'    sSql = sSql & "                 Left Join ColocacEstado CE2 ON CE2.cCtaCod = P.cCtaCod AND CE2.nPrdEstado = " & COMDConstantes.gColocEstSug
'    sSql = sSql & "                 Left Join ColocacEstado CE3 ON CE3.cCtaCod = P.cCtaCod AND CE3.nPrdEstado = " & COMDConstantes.gColocEstAprob
'    sSql = sSql & "                 Left Join ColocacEstado CE4 ON CE4.cCtaCod = P.cCtaCod AND CE4.nPrdEstado = " & COMDConstantes.gColocEstRech
'    sSql = sSql & "                 Left Join Constante CN4 ON CN4.nConsValor = CE4.nMotivoRechazo AND CN4.nConsCod = " & COMDConstantes.gColocMotivRechazo
'    sSql = sSql & "                 Left Join ColocacEstado CE5 ON CE5.cCtaCod = P.cCtaCod AND CE5.nPrdEstado = " & COMDConstantes.gColocEstCancelado
'    sSql = sSql & "                 Left Join ColocacEstado CE6 ON CE6.cCtaCod = P.cCtaCod AND CE6.nPrdEstado = " & COMDConstantes.gColocEstJudicial
'    sSql = sSql & "                 Left Join ColocacEstado CE7 ON CE7.cCtaCod = P.cCtaCod AND CE7.nPrdEstado = " & COMDConstantes.gColocEstRefNorm
'    sSql = sSql & "                 Left Join Constante CN5 ON CN5.nConsValor = P.nPrdEstado AND CN5.nConsCod = " & COMDConstantes.gColocEstado
'    'MAVM 20100611 BAS II ***
'    'sSql = sSql & "                 Left Join Constante CN6 ON convert(int,substring(P.cCtaCod,6,3)) = CN6.nConsValor AND CN6.nConsCod = " & COMDConstantes.gProducto
'    sSql = sSql & "                 LEFT JOIN Constante CN4B ON C.cTpoProdCod = CN4B.nConsValor AND CN4B.nConsCod = " & COMDConstantes.gTpoProducto
'    sSql = sSql & "                 LEFT JOIN Constante CN5B ON C.cTpoCredCod = CN5B.nConsValor AND CN5B.nConsCod = " & COMDConstantes.gTpoCredito
'    '***
'    sSql = sSql & "                 Left Join Constante CN7 ON CN7.nConsValor = CC.nTipoDesembolso AND CN7.nConsCod = " & COMDConstantes.gColocTiposDesembolso
'    sSql = sSql & "                 Left Join Constante CN8 ON CN8.nConsValor = CE3.nTipoGracia AND CN8.nConsCod = " & COMDConstantes.gColocTiposGracia
'    sSql = sSql & "                 Left Join ColocLineaCreditoTasa Lin ON C.cLineaCred = Lin.cLineaCred AND Lin.nColocLinCredTasaTpo = " & COMDConstantes.gColocLineaCredTasasIntGracia
'    If CInt(Mid(psCtaCod, 6, 3)) = gColConsuDctoPlan Then
'        sSql = sSql & "             Left Join ColocacConvenio Conv ON P.cCtaCod = Conv.cCtaCod"
'        sSql = sSql & "             Left Join Persona PConv ON PConv.cPersCod = Conv.cPersCod "
'    End If
'    sSql = sSql & " WHERE P.cCtaCod = '" & psCtaCod & "'"
''COmento JOEP20201202 Segemntacion Riesgo

    sSql = "Exec stp_sel_RecuperaConsultaCred '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaConsultaCred = oConecta.CargaRecordSet(sSql)  ' Seguimiento de eeff
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaConsultaCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaDacionesPago() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
                   
    sSql = " Select CG.nNroGarantRec, CG.cCtaCod, P.cPersNombre "
    sSql = sSql & " From ColocGarantRec CG Inner Join ProductoPersona PP ON CG.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
    sSql = sSql & "             Inner Join Persona P ON P.cPersCod = PP.cPersCod "
    sSql = sSql & "             Inner Join Colocaciones C ON CG.cCtaCod = C.cCtaCod Inner Join Producto Prd ON Prd.cCtaCod = CG.cCtaCod "
    sSql = sSql & "             Inner Join ColocacEstado CE ON CE.cCtaCod = CG.cCtaCod And CE.nPrdEstado = 2002"
    sSql = sSql & "    WHERE CG.nEstado = " & gColocGarantRecEstadoRegistrado

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDacionesPago = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

'Funciones de la Clase DCredGeneral
Public Function CargaAnalistas(Optional ByVal psAge As String = -1) As ADODB.Recordset 'PTI1 20170701 se Agredgo parametro psAge
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sAnalistas As String
Dim oGen As COMDConstSistema.DCOMGeneral
    On Error GoTo ERRORCargaAnalistas
    
    Set oGen = New COMDConstSistema.DCOMGeneral
    sAnalistas = oGen.LeeConstSistema(gConstSistRHCargoCodAnalistas)
    Set oGen = Nothing
    
'    sSql = "Select R.cPersCod, P.cPersNombre from RRHH R inner join Persona P ON R.cPersCod = P.cpersCod "
'    sSql = sSql & " AND nRHEstado = 201 "
'    sSql = sSql & " inner join RHCargos RC ON R.cPersCod = RC.cPersCod "
'    sSql = sSql & " where  RC.cRHCargoCod in (" & sAnalistas & ") AND RC.dRHCargoFecha = (select MAX(dRHCargoFecha) from RHCargos RHC2 where RHC2.cPersCod = RC.cPersCod) "
'    sSql = sSql & " order by P.cPersNombre "
    If Right(psAge, 2) = -1 Then 'PTI1 20170701
    sSql = "Exec STP_SEL_OBTENER_ANALISTAS" 'APRI20170527 ERS026-2017
    Else
    sSql = "stp_sel_ERS041_2017_CargaAnalistas'" & Right(psAge, 2) & "'" 'PTI1 20170701
    End If
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CargaAnalistas = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
ERRORCargaAnalistas:
    Err.Raise Err.Number, "Error", Err.Description

End Function

Public Function RecuperaParametro(Optional ByVal pnCodParametro As Integer = -1) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaParametro
    sSql = "Select * from ColocParametro "
    If pnCodParametro <> -1 Then
        sSql = sSql & " WHERE nParamVar = " & pnCodParametro
    End If
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaParametro = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaParametro:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaAgencias() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaAgencias
    sSql = "Select cAgeCod,cAgeDescripcion from Agencias"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaAgencias = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaAgencias:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function
'***************************************

Public Function RecuperaDacionPago(ByVal pnNroDacion As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    sSql = "Select C.*, P.cPersNombre, P.cPersCod From ColocGarantRec C Inner Join ProductoPersona PP ON C.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
    sSql = sSql & "     Inner Join Persona P ON PP.cPersCod = P.cPersCod "
    sSql = sSql & " Where nNroGarantRec = " & pnNroDacion
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDacionPago = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaDacionPagoDetalle(ByVal pnNroDacion As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    sSql = "Select * From ColocGarantRecdet Where nNroGarantRec = " & pnNroDacion
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDacionPagoDetalle = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaDatosDacionPagoCredito(ByVal pnDacion As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    sSql = "Select CG.cCtaCod, CG.nValorTotal, P.cPersNombre, CN.cConsDescripcion as cMoneda, C.nMontoCol,  "
    sSql = sSql & " Prd.nSaldo, CE.nCuotas, CC.nDiasAtraso, CC.cMetLiquidacion, Prd.nTransacc, CC.nCalendDinamico "
    sSql = sSql & " From ColocGarantRec CG Inner Join ProductoPersona PP ON CG.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
    sSql = sSql & "             Inner Join Persona P ON P.cPersCod = PP.cPersCod "
    sSql = sSql & "             Inner Join Constante CN ON Convert(int,SUBSTRING(CG.cCtaCod,9,1)) = CN.nConsValor And CN.nConsCod = " & COMDConstantes.gMoneda
    sSql = sSql & "             Inner Join Colocaciones C ON CG.cCtaCod = C.cCtaCod "
    sSql = sSql & "             Inner Join Producto Prd ON Prd.cCtaCod = CG.cCtaCod "
    sSql = sSql & "             Inner Join ColocacEstado CE ON CE.cCtaCod = CG.cCtaCod And CE.nPrdEstado = " & COMDConstantes.gColocEstAprob
    sSql = sSql & "             Inner Join ColocacCred CC ON CC.cCtaCod = CG.cCtaCod "
    sSql = sSql & " Where CG.nNroGarantRec = " & pnDacion
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosDacionPagoCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function


Public Function RecuperaProductosDeSolicitudDeCredito() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta


    sSql = " Select C.cConsDescripcion, C.nConsValor from Constante C "
    sSql = sSql & " Inner Join ColocCredConsFiltro CF ON C.nConsCod = CF.nConsCod AND C.nConsValor = CF.nConsValor"
    sSql = sSql & " Where CF.nCodFiltro = " & gCredFiltroProdCab & " and C.nConsCod=3033 "
    sSql = sSql & " Order by C.cConsDescripcion DESC"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaProductosDeSolicitudDeCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaUltimaCuotaEvaluada(ByVal psCtaCod As String) As Integer
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset
Dim nCuota1 As Integer

    sSql = "select ISNULL(MAX(nCuotaOrig),-1)+7 as nCuota from ColocCalifMiViv "
    sSql = sSql & "     where cCtaCod = '" & psCtaCod & "'"
    sSql = sSql & "   AND cColocMiVivEval='E'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    nCuota1 = R!nCuota

    sSql = "select ISNULL(MAX(nCuotaOrig),0) as nCuota from ColocCalifMiViv "
    sSql = sSql & "     where cCtaCod = '" & psCtaCod & "'"
    sSql = sSql & "   AND nColocCalendEstado=" & COMDConstantes.gColocCalendEstadoPagado
    Set R = oConecta.CargaRecordSet(sSql)
    RecuperaUltimaCuotaEvaluada = nCuota1 - R!nCuota
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaProductosDeSolicitudDeCF() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    sSql = " Select cConsDescripcion, nConsValor from Constante"
    'sSql = sSql & " where nConscod = 1001 and (nConsValor % 10) <> 0 and (nConsValor < 232 or nConsValor >= 300) AND nConsValor In (" & COMDConstantes.gColCFComercial & "," & COMDConstantes.gColCFPYME & ")"
    'MAVM *** 20100603 BAS II
    sSql = sSql & " where nConscod = " & COMDConstantes.gTpoProducto & " and nConsValor = " & COMDConstantes.gColCFTpoProducto & ""
    sSql = sSql & " Order by cConsDescripcion DESC"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaProductosDeSolicitudDeCF = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaTitularCredito(ByVal scCod As String) As String
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As New ADODB.Recordset

    sSql = " select cperscod from productopersona where cctacod='" & Trim(scCod) & "' and nprdpersrelac=" & COMDConstantes.gColRelPersTitular

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    If rs.BOF Then
        RecuperaTitularCredito = ""
    Else
        RecuperaTitularCredito = rs!cPersCod
    End If
    rs.Close
    Set rs = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Private Sub Class_Initialize()
Dim CIni As COMConecta.DCOMClasIni
    Set CIni = New COMConecta.DCOMClasIni
    gConsPersona = CIni.BasePersonas
    gConsComunes = CIni.BaseComunes
    gConsImagenes = CIni.BaseImagenes
    
    'JACA 20111228**************************************
    Set loConecta = New COMConecta.DCOMConecta
    If loConecta.AbreConexion = False Then
        Err.Raise Err.Number, "DCOMCredito:Initialize Method. Error en Conexion a Base de datos", Err.Description
    End If
    'JACA END*******************************************
    
    Set CIni = Nothing
End Sub

Public Function RecuperaCorresponCred(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCorresponCred
    sSql = "Select * from Corresponsalia where cCtacod = '" & psCtaCod & "' and nFlag=0"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCorresponCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCorresponCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function



Public Function ObtieneMonto_Validate(ByVal psCtaCod As String, ByVal pnMontoPago As Currency) As Boolean
    'Function que valida el monto no sobre pasa el pendiente
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs_Monto As ADODB.Recordset
    Dim nMonto As Currency
    On Error GoTo ErrHandler

    sSql = sSql & " SELECT SUM(CD.NMONTO-CD.NMONTOPAGADO) AS RESULT"
    sSql = sSql & " FROM COLOCCALENDDET CD"
    sSql = sSql & " INNER JOIN COLOCCALENDARIO C ON C.CCTACOD=CD.CCTACOD AND CD.NNROCALEN=C.NNROCALEN AND"
    sSql = sSql & " C.NCOLOCCALENDAPL = CD.NCOLOCCALENDAPL And C.nCuota = CD.nCuota"
    sSql = sSql & " WHERE C.CCTACOD='" & psCtaCod & "' AND C.NCOLOCCALENDAPL=1 AND C.NCOLOCCALENDESTADO=0"
        Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
        Set rs_Monto = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion

        If Not rs_Monto.EOF And Not rs_Monto.BOF Then
            nMonto = rs_Monto!result
        End If
        Set rs_Monto = Nothing

        If pnMontoPago > nMonto Then
           ObtieneMonto_Validate = False
        Else
            ObtieneMonto_Validate = True
        End If
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    If Not rs_Monto Is Nothing Then Set rs_Monto = Nothing
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function ActualizaColocacConvenioRegDevolucion(ByVal psPersCodIns As String, ByVal psPersCod As String, _
ByVal pdRegistro As Date, ByVal psCodAge As String, ByVal pnMoneda As COMDConstantes.Moneda, ByVal pnMonto As Currency, ByVal psNroDoc As String)
Dim SQL As String
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

oCon.AbreConexion

SQL = "UPDATE ColocacConvenioRegDev SET nMoneda=" & pnMoneda & ", nMonto = " & pnMonto & ", cNroDoc = '" & psNroDoc & "' " _
    & " WHERE cPersCodIns ='" & psPersCodIns & "' and cPersCod='" & psPersCod & "' and dRegistro='" & Format(pdRegistro, "mm/dd/yyyy") & "'"

oCon.Ejecutar SQL

oCon.CierraConexion
End Function

Public Function ExtornaRegDevConvOpe(ByVal pnMovNro As Currency)
Dim SQL As String
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

oCon.AbreConexion

SQL = "UPDATE ColocacConvenioRegDev SET nMovNro=NULL, nPendiente = 0  " _
    & " WHERE nMovNro = " & pnMovNro & ""

oCon.Ejecutar SQL

oCon.CierraConexion
End Function

Public Function EliminaColocacConvenioRegDevolucion(ByVal psPersCodIns As String, ByVal psPersCod As String, ByVal pdRegistro As Date, Optional ByVal pnMovNro As Long = 0) As Boolean 'PASI20141222 Agrego pnMovNro
Dim SQL As String
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta
Dim bTrans As Boolean
    
'MODIFICADO PASI20141222
'oCon.AbreConexion
'
'SQL = "DELETE FROM ColocacConvenioRegDev " _
'    & " WHERE cPersCodIns ='" & psPersCodIns & "' and cPersCod='" & psPersCod & "' and dRegistro='" & Format(pdRegistro, "mm/dd/yyyy") & "'"
'
'oCon.Ejecutar SQL
'
'oCon.CierraConexion

    On Error GoTo ErrEliminaColocacConvRegDev
    bTrans = False
    oCon.AbreConexion
    oCon.ConexionActiva.BeginTrans
    bTrans = True
    
    If pnMovNro > 0 Then
        SQL = "UPDATE M"
        SQL = SQL & " set M.nMovFlag = " & gMovFlagEliminado
        SQL = SQL & " FROM MOV M"
        SQL = SQL & " WHERE M.nMovNro = " & pnMovNro
        oCon.ConexionActiva.Execute SQL
    End If
    
    SQL = "DELETE FROM ColocacConvenioRegDev " _
    & " WHERE cPersCodIns ='" & psPersCodIns & "' and cPersCod='" & psPersCod & "' and dRegistro='" & Format(pdRegistro, "mm/dd/yyyy") & "'"
    oCon.ConexionActiva.Execute SQL
    oCon.ConexionActiva.CommitTrans
    bTrans = False
    oCon.CierraConexion
    EliminaColocacConvenioRegDevolucion = True
    Set oCon = Nothing
    Exit Function
ErrEliminaColocacConvRegDev:
    If bTrans Then oCon.ConexionActiva.RollbackTrans
    Set oCon = Nothing
    Err.Raise Err.Number, "Error En Proceso de Eliminar", Err.Description
    EliminaColocacConvenioRegDevolucion = False
'end PASI
End Function
'MADM 20110225 - STORE+CCTACOD
Public Function GetColocacConvRegDevolucion(ByVal psCodPersIns As String) As ADODB.Recordset
Dim SQL As String
Dim oCon As COMConecta.DCOMConecta

'Sql = "select  C.cPersCod, cPersNombre, CONVERT(CHAR(10), c.dRegistro,103) as dRegistro, c.nMonto, c.cNroDoc, " _
'    & "        RTRIM(cConsDescripcion) + SPACE(75) + CONVERT(CHAR(1),nMoneda) AS cMoneda, '1' as cNuevo, nMovNroReg " _
'    & " from    ColocacConvenioRegDev C " _
'    & "         JOIN PERSONA P ON C.cPerscod = P.cPersCod " _
'    & "         JOIN CONSTANTE CC ON CC.nConsValor = C.nMoneda and CC.NCONSCOD =1011 " _
'    & " WHERE   (C.nPendiente IS NULL OR C.nPendiente<>'1')  and C.cPersCodIns='" & psCodPersIns & "' " _
'    & "         and exists (SELECT M.nMovNro FROM MOV M WHERE M.nMovNro = nMovNroReg and M.nMovflag=0 ) " _
'    & " ORDER BY P.CPERSNOMBRE "
SQL = "exec stp_sel_CargaDevolucionesInstitucion '" & psCodPersIns & "'"

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set GetColocacConvRegDevolucion = oCon.CargaRecordSet(SQL)

oCon.CierraConexion
Set oCon = Nothing

End Function
'END MADM

Public Function RecuperaEstadoCredito(ByVal psCtaCod As String) As Integer
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset

    sSql = "Select nPrdEstado From Producto Where cCtaCod='" & psCtaCod & "'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing

    If Not rs.EOF And Not rs.BOF Then
        RecuperaEstadoCredito = rs!nPrdEstado
    End If
    Set rs = Nothing
End Function


Public Function ListaKardex(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta

    sSql = "Select C.nCuota, "
    sSql = sSql & " MontoCuota=(Select Sum(nMontoPagado) From ColocCalendDet Where cCtaCod=C.cCtaCod and nNroCalen=C.nNroCalen and nColocCalendApl=C.nColocCalendApl and nCuota=C.nCuota),"
    sSql = sSql & " Capital=(Select Sum(nMontoPagado) From ColocCalendDet Where cCtaCod=C.cCtaCod and nNroCalen=C.nNroCalen and nColocCalendApl=C.nColocCalendApl and nCuota=C.nCuota and nPrdConceptoCod in (1000,1010)),"
    sSql = sSql & " Interes=(Select Sum(nMontoPagado) From ColocCalendDet Where cCtaCod=C.cCtaCod and nNroCalen=C.nNroCalen and nColocCalendApl=C.nColocCalendApl and nCuota=C.nCuota and nPrdConceptoCod in (1100,1107)),"
    sSql = sSql & " Mora=(Select Sum(nMontoPagado) From ColocCalendDet Where cCtaCod=C.cCtaCod and nNroCalen=C.nNroCalen and nColocCalendApl=C.nColocCalendApl and nCuota=C.nCuota and nPrdConceptoCod in (1101,1108)),"
    sSql = sSql & " Gastos=(Select Sum(nMontoPagado) From ColocCalendDet Where cCtaCod=C.cCtaCod and nNroCalen=C.nNroCalen and nColocCalendApl=C.nColocCalendApl and nCuota=C.nCuota and nPrdConceptoCod like '12%'),"
    sSql = sSql & " Atraso = DateDiff(Day, dVenc , C.dPago)"
    sSql = sSql & " From Producto P"
    sSql = sSql & " Inner Join ColocCalendario C on C.cCtaCod=P.cCtaCod"
    sSql = sSql & " Where P.cCtaCod='" & psCtaCod & "' and C.nColocCalendEstado=1 and C.nColocCalendApl=1"
    sSql = sSql & " and C.nNroCalen=(Select Max(nNroCalen) From ColocCalendario Where cCtaCod='" & psCtaCod & "')"
    sSql = sSql & " Order By C.nCuota"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaKardex = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function CargarProductos() As ADODB.Recordset
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    
'    sSql = "Select nConsValor,cConsDescripcion"
'    sSql = sSql & " From Constante"
'    sSql = sSql & " Where nconsCod = 1001 And nConsValor <> 1001"
'    sSql = sSql & " Order By nConsValor"

    sSql = "Exec STP_SEL_OBTENER_PRODUCTO_CREDITOS" 'APRI20170527 ERS026-2017

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set CargarProductos = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function VerificarVencimientoCredito(ByVal psCtaCod As String) As Boolean
    ' funcion que verifica si el credito esta vencido o no
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim nDiasAtraso As Integer
    Dim nDiasMinimo As Integer

    sSql = "Select nParamValor From ColocParametro Where nParamVar=3101"
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
        Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing

    If Not rs.EOF And Not rs.BOF Then
        nDiasMinimo = rs!nParamValor
    End If
    Set rs = Nothing
    VerificarVencimientoCredito = False
    sSql = "Select nDiasAtraso From ColocacCred Where cCtaCod='" & psCtaCod & "'"
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    If Not rs.EOF And Not rs.BOF Then
        If rs!nDiasAtraso <= nDiasMinimo Then
            VerificarVencimientoCredito = True
        Else
            VerificarVencimientoCredito = False
        End If
    End If
    Set rs = Nothing
End Function

Public Function RecuperaDetallePago(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta

    sSql = "Select nMovNro,cUsuario,cOpeCod,dFecPago,nNroCuota,nMontoPagado,nCapital,nInteres,nMora,"
    sSql = sSql & "  nGastos , nITF, nDiasMora, nSaldoCap"
    sSql = sSql & " From ("
    sSql = sSql & " Select MC.nMovNro, right(M.cMovNro,4) as cUsuario, MC.cOpeCod, dbo.FechaMov(M.cMovNro) as dFecPago,"
    sSql = sSql & " MD.nNroCuota,"
    sSql = sSql & " nMontoPagado = (Select ISNULL(SUM(nMonto),0)"
    sSql = sSql & " From MovColDet"
    sSql = sSql & " Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro and cOpeCod not like '107[123456789]%' and nNroCuota=MD.nNroCuota)+(Select isnull(Sum(nMovImporte),0) From movOpeVarias Where nMovNro= MC.nMovNro),"
    sSql = sSql & " nCapital = (Select ISNULL(SUM(nMonto),0)"
    sSql = sSql & " From MovColDet"
    sSql = sSql & " Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro AND nPrdConceptoCod in(1000,3000) and cOpeCod not like '107[123456789]%' and nNroCuota=MD.nNroCuota),"
    sSql = sSql & " nInteres = (Select ISNULL(SUM(nMonto),0)"
    sSql = sSql & " From MovColDet"
    sSql = sSql & " Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod in (1100,1102,1103,1104,1105,3100) and nNroCuota=MD.nNroCuota),"
    sSql = sSql & " nMora = (Select ISNULL(SUM(nMonto),0)"
    sSql = sSql & " From MovColDet"
    sSql = sSql & " Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro AND nPrdConceptoCod in(1101,3101) and nNroCuota=MD.nNroCuota),"
    sSql = sSql & " nGastos = (Select ISNULL(SUM(nMonto),0)"
    sSql = sSql & " From MovColDet"
    sSql = sSql & " Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod like  '12%' and nNroCuota=MD.nNroCuota),"
    sSql = sSql & " nITF = (Select ISNULL(SUM(nMonto),0)"
    sSql = sSql & " From MovColDet"
    'sSQL = sSQL & " Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod in  (20, 21) and nNroCuota=MD.nNroCuota), MC.nDiasMora , MC.nSaldoCap"
     sSql = sSql & " Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod in  (20, 21) and "
     sSql = sSql & " (Select Max(nNroCuota) From MovColDet Where nMovNro=MC.nMovNro)=MD.nNroCuota), MC.nDiasMora , MC.nSaldoCap "
    sSql = sSql & " from MovCol MC"
    sSql = sSql & " Inner Join Mov M ON M.nMovNro = MC.nMovNro"
    sSql = sSql & " Inner Join MovColDet MD on MD.nMovNro=M.nMovNro and MC.cCtaCod=MD.cCtaCod"
    sSql = sSql & " where (MC.cOpeCod like '100%' OR MC.cOpeCod = '105001') AND MC.cCtaCod = '" & psCtaCod & "'  AND M.nMovFlag<>2 AND"
    sSql = sSql & " (SUBSTRING(MC.cOpeCod,1,4) in ('1002','1003','1004','1005','1006','1007','1070','1050'))"
    sSql = sSql & " ) X"
    sSql = sSql & " Where nNroCuota <> 0"
    sSql = sSql & " Group By nMovNro,cUsuario,cOpeCod,dFecPago,nNroCuota,nMontoPagado,nCapital,nInteres,nMora,"
    sSql = sSql & " nGastos , nITF, nDiasMora, nSaldoCap"
    sSql = sSql & " Order by nMovNro"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set RecuperaDetallePago = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function ValidaSugAprobacion(ByVal psCtaCod As String) As Boolean
'Funcion que valida que en el momento de sugerir pueda aprobar
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim bValor As Boolean

    bValor = True
    'CTI6-20210503-ERS032-2019 -(Optimizar Sugerencia)
    'sSql = "Select  Count(*) as nCantidad"
    'sSql = sSql & " From ColocacEstado"
    'sSql = sSql & " where cctacod='" & psCtaCod & "' and nPrdEstado=2002"

    sSql = "Execute spt_sel_ValidaSugAprobacion '" & psCtaCod & "'"
    'Fin CTI6-20210503
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing

    If Not rs.EOF And Not rs.BOF Then
        If rs!nCantidad > 0 Then
            bValor = False
        Else
            bValor = True
        End If
    End If
    Set rs = Nothing
    ValidaSugAprobacion = bValor
End Function


Public Function GetDatosCheque(ByVal psCodCheque As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta

    sSql = "Select MD.nMovNro,MD.cDocNro as cCheque,Pers.cPersNombre,AG.cAgeDescripcion as cAgencia,           "
    sSql = sSql & "            MOV.nMovImporte as nMonto,MOV.nMoneda,Ag.cAgeCod,MOI.cCtaIfCod,Pers.cPersCod"
    sSql = sSql & " From MovDoc MD"
    sSql = sSql & "           Inner Join MovObjIF MOI on MOI.nMovNro=MD.nMovNro"
    sSql = sSql & "           Inner Join Persona Pers on Pers.cPersCod=MOI.cPersCod"
    sSql = sSql & "           Inner Join MovObjAreaAgencia MOA on MOA.nMovNro=MD.nMovNro"
    sSql = sSql & "           Inner Join Agencias  AG on AG.cAgeCod=MOA.cAgeCod"
    sSql = sSql & "           Inner Join MovOpeVarias MOV on MOV.nMovNro=MD.nMovNro and MOV.cNroDoc= MD.cDocNro"
    sSql = sSql & "           Inner Join Mov M on M.nMovNro=MD.nMovNro"
    sSql = sSql & "  Where MD.cDocNro='" & psCodCheque & "' and MD.nDocTpo=47 and M.nMovFlag=0 and Not Exists(Select * From MovRef Where nMovNroRef=M.nMovNro)"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set GetDatosCheque = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function Cargar_Objetos_Controles(ByRef prsCondCred As ADODB.Recordset, _
                                        ByRef prsCondCred2 As ADODB.Recordset, _
                                        ByRef prsTipoCred As ADODB.Recordset, _
                                        ByRef prsDestCred As ADODB.Recordset, _
                                        ByRef prsAnalista As ADODB.Recordset, _
                                        ByRef pbMuestraSoloAnalistaActual As Integer, _
                                        ByRef prsMoneda As ADODB.Recordset, _
                                        ByRef prsInstituc As ADODB.Recordset, _
                                        ByVal psCodAge As String, _
                                        Optional ByRef prsMotivoRef As ADODB.Recordset, _
                                        Optional ByRef prsPromotores As ADODB.Recordset)
                                        'By Capi 28102008 se agrego parametro prsMotivoRef
'WIOR 20140509 AGREGO prsPromotores

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim oCons As COMDConstantes.DCOMConstantes
'Analistas
Dim sAnalistas As String
Dim oGen As COMDConstSistema.DCOMGeneral

Dim oPersonas As COMDPersona.DCOMPersonas

    On Error GoTo ErrorCargar_Objetos_Controles
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
    Set oCons = New COMDConstantes.DCOMConstantes
    
    Set prsCondCred = oCons.RecuperaConstantes(gColocCredCondicion)
      
    'Set prsCondCred2 = oCons.RecuperaConstantes(gColocCredCondicionOtra)
    '**Modificado por DAOR 20071102, No considerar campañas de ahorros
    'sSql = "Select nConsValor = IdCampana, cConsDescripcion = cDescripcion From Campanas Where bEstado = 1"
    
    '*** PEAC 20091119
    'sSql = "Select nConsValor = IdCampana, cConsDescripcion = cDescripcion From Campanas Where bEstado = 1 and (cProd <> 'A' or cProd is null) "
    If Len(Trim(psCodAge)) > 0 Then
        sSql = "exec stp_sel_ObtieneCampanas '" & psCodAge & "'"
    Else
        sSql = "exec stp_sel_ObtieneCampanas "
    End If
    '*** FIN PEAC
    
    Set prsCondCred2 = oConecta.CargaRecordSet(sSql, adLockReadOnly)
    
    Set prsMoneda = oCons.RecuperaConstantes(gMoneda)
    Set prsDestCred = oCons.RecuperaConstantes(gColocDestino)
    
    'By capi 28102008
    Set prsMotivoRef = oCons.RecuperaConstantes(gColocMotivoRef)
    '
    
    Set oCons = Nothing
    'ALPA 20100603 BASII*****************
    'Set prsTipoCred = RecuperaProductosDeSolicitudDeCredito
    Set prsTipoCred = RecuperaProductosCrediticios
    '*****************************
    
'************ Analistas *************

    Set oGen = New COMDConstSistema.DCOMGeneral
    sAnalistas = oGen.LeeConstSistema(gConstSistRHCargoCodAnalistas)
    pbMuestraSoloAnalistaActual = oGen.LeeConstSistema(58)
    Set oGen = Nothing
    
    sSql = "Select R.cPersCod, P.cPersNombre from RRHH R inner join Persona P ON R.cPersCod = P.cpersCod "
    sSql = sSql & " AND nRHEstado in (201,301) "
    sSql = sSql & " inner join RHCargos RC ON R.cPersCod = RC.cPersCod "
    sSql = sSql & " where  RC.cRHCargoCod in (" & sAnalistas & ") AND RC.dRHCargoFecha = (select MAX(dRHCargoFecha) from RHCargos RHC2 where RHC2.cPersCod = RC.cPersCod) "
    sSql = sSql & " and R.cAgenciaActual='" & psCodAge & "'"
    sSql = sSql & " order by P.cPersNombre "
    
    Set prsAnalista = oConecta.CargaRecordSet(sSql)
       
'********** Instituciones *******************
    Set oPersonas = New COMDPersona.DCOMPersonas
    Set prsInstituc = oPersonas.RecuperaPersonasTipo(gPersTipoConvenio)
    Set oPersonas = Nothing
    
    'WIOR 20140509 - PROMOTORES *********************
    sSql = "EXEC stp_sel_ObtenerPromotoresCreditos '" & psCodAge & "'"
    Set prsPromotores = oConecta.CargaRecordSet(sSql)
    'WIOR FIN ***************************************
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorCargar_Objetos_Controles:
    Err.Raise Err.Number, "Cargar Objetos Controles", Err.Description

End Function

Public Function GeneraMovNroActualiza(pdFecha As Date, psCodUser As String, psCodCmac As String, psCodAge As String) As String
Dim oConecta As New COMConecta.DCOMConecta
Dim oVar As New COMFunciones.FCOMVarPublicas

oConecta.AbreConexion
GeneraMovNroActualiza = Format(pdFecha & " " & oConecta.GetHoraServer, oVar.gsFormatoMovFechaHora) & psCodCmac & psCodAge & "00" & psCodUser

oConecta.CierraConexion
Set oConecta = Nothing
Set oVar = Nothing
End Function

Public Function CargarHistoriaCredito(ByVal psCtaCod As String, _
                                      ByVal pbPagoCuota As Boolean, _
                                      ByRef prsCredVig As ADODB.Recordset, _
                                      ByRef prsCalend As ADODB.Recordset)

Dim oCred As COMDCredito.DCOMCalendario
Dim nNroCalen As Integer
Dim i As Integer

'ARCV 26-06-2007
Dim oBase As COMDCredito.DCOMCredActBD
'----------
On Error GoTo ErrorCargarHistoriaCredito
        
        'ARCV 26-06-2007
        'Set prsCredVig = RecuperaDatosCreditoVigente(psCtaCod)
        Set oBase = New COMDCredito.DCOMCredActBD
        Set prsCredVig = RecuperaDatosCreditoVigente(psCtaCod)
        Set oBase = Nothing
        '------------------
        
        If prsCredVig.RecordCount > 0 Then
            
            i = prsCredVig!nNroCalen
            If Not pbPagoCuota Then
               nNroCalen = 1
            Else
               nNroCalen = i
            End If
            
            Set oCred = New COMDCredito.DCOMCalendario
            If pbPagoCuota = False Then
                Set prsCalend = oCred.RecuperaCalendarioPagos(psCtaCod, nNroCalen)
            Else
                Set prsCalend = oCred.RecuperaCalendarioPagos(psCtaCod, nNroCalen, , , , True)
            End If
            Set oCred = Nothing
        End If
        
    Exit Function
ErrorCargarHistoriaCredito:
    Err.Raise Err.Number, "Cargar Historia Credito", Err.Description

End Function

Public Function CargarControlesPagoLote(ByRef prsInstit As ADODB.Recordset, _
                                        ByRef prsTipoPago As ADODB.Recordset)

Dim oCons As COMDConstantes.DCOMConstantes
Dim oPersonas As COMDPersona.DCOMPersonas

On Error GoTo ErrorCargarControlesPagoLote
    
'    Call CargaComboPersonasTipo(gPersTipoConvenio, CmbInstitucion)
    'Carga Formas de Pago
'    Call CargaComboConstante(gColocTipoPago, CmbForPag)
    'Carga Cheques Recibidos
Set oCons = New COMDConstantes.DCOMConstantes
Set prsTipoPago = oCons.RecuperaConstantes(gColocTipoPago)
Set oCons = Nothing

Set oPersonas = New COMDPersona.DCOMPersonas
Set prsInstit = oPersonas.RecuperaPersonasTipo(gPersTipoConvenio)
Set oPersonas = Nothing
Exit Function

ErrorCargarControlesPagoLote:
    Err.Raise Err.Number, "Error", Err.Description

End Function

Public Function CargarControlesReasignacion(ByRef prsAnalista As ADODB.Recordset, _
                                            ByRef prsAgencia As ADODB.Recordset, _
                                            ByRef prsProducto As ADODB.Recordset, _
                                            Optional prsAnalistaEx As ADODB.Recordset = Nothing) 'PTI1 20170630 prsAnalistaEx

Dim sSql As String
Dim oConec As COMConecta.DCOMConecta

    On Error GoTo ErrorCargarControlesReasignacion

    Set prsAnalista = CargaAnalistas
    Set prsAnalistaEx = CargaAnalistasEx 'PTI1 20170630 prsAnalistaEx
    
    sSql = "select cagecod,cAgeDescripcion from agencias"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set prsAgencia = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    Set prsProducto = CargarProductos

    Exit Function
ErrorCargarControlesReasignacion:
    Err.Raise Err.Number, "Error", Err.Description
    
End Function


Public Function ReasignaCarteraAnalista(ByVal pMatCuentas As Variant, _
                                        ByVal psPerscodActual As String, _
                                        ByVal psPersCodAnterior As String, _
                                        ByVal pdFecSis As Date, _
                                        ByVal psAgeCod As String)
                                        'JUEZ 20160425 Se agregó psAgeCod
    
Dim i As Integer

On Error GoTo ErrorReasignaCarteraAnalista

For i = 1 To UBound(pMatCuentas)
    'Call ActualizaCreditoAnalista(pMatCuentas(i), psPerscodActual, psPersCodAnterior, pdFecSis) 'WIOR 20131119 COMENTÓ
    Call ActualizaCreditoAnalista(pMatCuentas(i, 1), psPerscodActual, psPersCodAnterior, pdFecSis, CDbl(pMatCuentas(i, 2)), _
    Trim(pMatCuentas(i, 3)), CInt(pMatCuentas(i, 4)), CDbl(pMatCuentas(i, 5)), CDbl(pMatCuentas(i, 6)), CDbl(pMatCuentas(i, 7)), _
    CDbl(pMatCuentas(i, 8)), psAgeCod) 'WIOR 20131119 AGREGÓ
    'JUEZ 20160425 Se agregó psAgeCod
Next i

Exit Function

ErrorReasignaCarteraAnalista:
    Err.Raise Err.Number, "Error", Err.Description

End Function

Public Function ReasignaInstituciones(ByVal pMatCuentas As Variant, _
                                        ByVal psPersCodInstitucion As String)

Dim oCredActBD As COMDCredito.DCOMCredActBD

Dim i As Integer

On Error GoTo ErrorReasignaInstituciones

Set oCredActBD = New COMDCredito.DCOMCredActBD

For i = 1 To UBound(pMatCuentas)
    Call oCredActBD.dUpdateColocacConvenio(pMatCuentas(i), psPersCodInstitucion, False)
Next i

Set oCredActBD = Nothing
Exit Function

ErrorReasignaInstituciones:
    Err.Raise Err.Number, "Reasigna Instituciones", Err.Description

End Function


Public Function ConsolidaDatosCliente(ByVal pdFecSis As Date, _
                                ByVal psCodPers As String, _
                                ByRef pnTipoCambio As Currency, _
                                ByRef pnSaldoCaptacS As Currency, _
                                ByRef pnSaldoCaptacD As Currency, _
                                ByRef pnSaldoColocS As Currency, _
                                ByRef pnSaldoColocD As Currency, _
                                ByRef pnSaldoCFS As Currency, _
                                ByRef pnSaldoCFD As Currency, _
                                ByRef pnSaldoCaptacTS As Currency, _
                                ByRef pnSaldoColocTS As Currency, _
                                ByRef pnSaldoCFTS As Currency, _
                                ByRef pnSaldoColocCFGarantizaS As Currency, _
                                ByRef pnSaldoColocCFGarantizaD As Currency, _
                                ByRef pnSaldoColocCFGarantizaTS As Currency)
                                'ByRef pnSaldoCFTS As Currency)
'EJVG20120430 Se agregó los parametros para los productos que el cliente garantiza

Dim oDatos As COMDCredito.DCOMCreditos
Dim oTipoCambio As COMDConstSistema.NCOMTipoCambio

On Error GoTo ErrorConsolidaDatosCliente


Set oTipoCambio = New COMDConstSistema.NCOMTipoCambio
    pnTipoCambio = oTipoCambio.EmiteTipoCambio(pdFecSis, TCFijoMes)
Set oTipoCambio = Nothing

Set oDatos = New COMDCredito.DCOMCreditos
pnSaldoCaptacS = oDatos.SaldoConsolCapta(psCodPers, 1)
pnSaldoCaptacD = oDatos.SaldoConsolCapta(psCodPers, 2)
pnSaldoColocS = oDatos.SaldoConsolColoc(psCodPers, 1)
pnSaldoColocD = oDatos.SaldoConsolColoc(psCodPers, 2)
pnSaldoCFS = oDatos.SaldoConsolCF(psCodPers, 1)
pnSaldoCFD = oDatos.SaldoConsolCF(psCodPers, 2)
pnSaldoColocCFGarantizaS = oDatos.SaldoConsolColocCFGarantiza(psCodPers, 1) 'EJVG20120430
pnSaldoColocCFGarantizaD = oDatos.SaldoConsolColocCFGarantiza(psCodPers, 2) 'EJVG20120430
pnSaldoCaptacTS = Round((pnSaldoCaptacS + (pnTipoCambio * pnSaldoCaptacD)), 2)
pnSaldoColocTS = Round((pnSaldoColocS + (pnTipoCambio * pnSaldoColocD)), 2)
pnSaldoCFTS = Round((pnSaldoCFS + (pnTipoCambio * pnSaldoCFD)), 2)
pnSaldoColocCFGarantizaTS = Round((pnSaldoColocCFGarantizaS + (pnTipoCambio * pnSaldoColocCFGarantizaD)), 2) 'EJVG20120430

Set oDatos = Nothing
Exit Function

ErrorConsolidaDatosCliente:
    Err.Raise Err.Number, "Error", Err.Description
End Function

'Modificado x JUEZ 20160316 **********************************************
Public Function CargarDatosReprogramacion(ByVal psCtaCod As String, ByRef prsCred As ADODB.Recordset, ByRef prsCal As ADODB.Recordset, _
                                          ByRef prsReprogApr As ADODB.Recordset, ByRef prsDatosReprog As ADODB.Recordset, ByRef pbAutorizacion As Boolean)

Dim oCalend As COMDCredito.DCOMCalendario
Dim rsAut As ADODB.Recordset, rsRep As ADODB.Recordset

'Parametrizar
Dim oParam As COMDCredito.DCOMParametro

    On Error GoTo ErrorCargarDatosReprogramacion
    
    'pbCredito = VerificarVencimientoCredito(psCtaCod) 'Comentado x JUEZ 20131104
    'If pbCredito = False Then Exit Function
    
    Set prsCred = RecuperaDatosComunes(psCtaCod, False)
    Set prsReprogApr = RecuperaColocacReprogramado(psCtaCod, gEstReprogAprobado)
    
    Set oCalend = New COMDCredito.DCOMCalendario
    Set prsCal = oCalend.RecuperaCalendarioPagos(psCtaCod)
    Set oCalend = Nothing
    
    Set prsDatosReprog = RecuperaColocacReprogramadoEstado(psCtaCod, gEstReprogSolicitado)
    'Set rsAut = RecuperaColocacReprogramadoEstado(psCtaCod, gEstReprogAutorizado)'Comento JOEP20171222 Acta 220 - 2017
    Set rsAut = RecuperaColocacReprogramadoEstado(psCtaCod, gEstReprogAprobado) 'Agrego JOEP20171222 Acta 220 - 2017
    
    If Not (rsAut.EOF And rsAut.BOF) Then
        If prsDatosReprog!bSolicAutorizacion And CDate(prsDatosReprog!dPrdEstado) <= CDate(rsAut!dPrdEstado) Then
            pbAutorizacion = True
        End If
    End If
    Set rsAut = Nothing
    
    'pnMontoApr = SaldoPactadoCredito(psCtaCod)
    'Parametrizar
    'Set oParam = New COMDCredito.DCOMParametro
    'pnReprogUltimaCuotaFija = oParam.RecuperaValorParametro(9003)
    'Set oParam = Nothing
    
    Exit Function

ErrorCargarDatosReprogramacion:
    Err.Raise Err.Number, "Cargar Datos Reprogramacion", Err.Description
End Function
'END JUEZ ****************************************************************

Public Function CargarControlesMetaAnalista(ByRef prsAna As ADODB.Recordset, _
                                            ByRef prsMon As ADODB.Recordset, _
                                            ByRef prsProdCred As ADODB.Recordset)

Dim oCons As COMDConstantes.DCOMConstantes

On Error GoTo ErrorCargarControlesMetaAnalista
    

    'Carga Analistas
    Set prsAna = CargaAnalistas
    
    Set oCons = New COMDConstantes.DCOMConstantes
    Set prsMon = oCons.RecuperaConstantes(gMoneda)
    Set oCons = Nothing
    
    Set prsProdCred = RecuperaProductosDeCredito
    
    Exit Function

ErrorCargarControlesMetaAnalista:
    Err.Raise Err.Number, "Cargar Datos Reprogramacion", Err.Description
End Function

Public Function CargaDatosPerdonarMora(ByVal psCtaCod As String, _
                                        ByRef pnMontoApr As Double, ByRef pnTasaInteres As Double, _
                                        ByRef prsCred As ADODB.Recordset, _
                                        ByRef prsCal As ADODB.Recordset, _
                                        ByRef pnNroCalen As Integer, _
                                        ByRef pnPrdEstado As Integer, _
                                        Optional ByRef pnDiasAtraso As Integer, Optional psTipoOpeCod As String = "")
'WIOR 20150331 SE AGREGO pnDiasAtraso
'NAGL 20190816 Según ERS036-2018 agregó psTipoOpeCod

Dim oCalend As COMDCredito.DCOMCalendario

On Error GoTo ErrorCargaDatosPerdonarMora

    Set oCalend = New COMDCredito.DCOMCalendario
    Set prsCal = oCalend.RecuperaCalendarioPagos(psCtaCod, , , , , , True, False, psTipoOpeCod)
    Set oCalend = Nothing
    
    Set prsCred = RecuperaColocacEstado(psCtaCod, gColocEstAprob)
    If prsCred.BOF Or prsCred.EOF Then
        pnMontoApr = 0
    Else
        pnMontoApr = CDbl(Format(IIf(IsNull(prsCred!nMonto), 0, prsCred!nMonto), "#0.00"))
    End If
    prsCred.Close
    
    Set prsCred = RecuperaProducto(psCtaCod)
    If prsCred.BOF Or prsCred.EOF Then
        pnTasaInteres = 0
    Else
        pnTasaInteres = CDbl(Format(prsCred!nTasaInteres, "#0.00"))
        'NAGL 20190716 Traslado esta parte
        pnPrdEstado = prsCred!nPrdEstado
        prsCred.Close
    End If

    Set prsCred = RecuperaColocacCred(psCtaCod)
    If prsCred.BOF Or prsCred.EOF Then
        pnNroCalen = 0
        pnDiasAtraso = 0 'WIOR 20150331
    Else
        pnNroCalen = prsCred!nNroCalen
        pnDiasAtraso = prsCred!nDiasAtraso 'WIOR 20150331
    End If
    
    Exit Function

ErrorCargaDatosPerdonarMora:
    Err.Raise Err.Number, "Error Carga Datos Perdonar Mora", Err.Description
End Function

Public Function CargaDatosReportesDuplicados(ByVal psCtaCod As String, _
                                        ByRef prsCred As ADODB.Recordset, _
                                        ByRef prsComun As ADODB.Recordset, _
                                        ByRef psEstado As String, Optional ByRef pnEstado As Long = 0) As Boolean

'*** PEAC 20091211 - se agrego el parametro pnEstado

Dim nEstado As Integer
Dim oCons As COMDConstantes.DCOMConstantes
    
On Error GoTo ErrorCargaDatosReportesDuplicados

    CargaDatosReportesDuplicados = True
    
    Set prsCred = RecuperaProducto(psCtaCod)
    
    'Set oCred = Nothing
    If Not prsCred.BOF And Not prsCred.EOF Then
        Set oCons = New COMDConstantes.DCOMConstantes
        psEstado = oCons.DameDescripcionConstante(gColocEstado, prsCred!nPrdEstado)
        Set oCons = Nothing
    
        nEstado = prsCred!nPrdEstado
        Set prsComun = RecuperaDatosComunes(psCtaCod, , Array(nEstado, 0))
        
        pnEstado = nEstado
        
    Else
        CargaDatosReportesDuplicados = False
    End If
    
    Exit Function

ErrorCargaDatosReportesDuplicados:
    CargaDatosReportesDuplicados = False
    Err.Raise Err.Number, "Error Carga Datos Reportes Duplicados", Err.Description
End Function

Public Sub RecuperaDatosMetasAnalistaDetalle_Montos(ByVal psPersCod As String, ByVal pnTipo As Integer, _
                        ByVal pdInicio As Date, ByVal pdFinal As Date, ByVal pnMoneda As Integer, _
                        ByRef prsDet As ADODB.Recordset, ByRef pMatMontos As Variant)

    '20060405
    'nueva variable rsMetaAnalista
    Dim rsMetaAnalista As New ADODB.Recordset
    Dim nValor As Double
    
    On Error GoTo ErrorRecuperaDatosMetasAnalistaDetalle_Montos

    Set prsDet = RecuperaDatosMetasAnalistaDetalle(psPersCod, pnTipo, pdInicio, pdFinal, pnMoneda)
    
    If prsDet.EOF Then Exit Sub
    
    'codigo obsoleto
    'ReDim pMatMontos(prsDet.RecordCount, 4)
    
    '20060405
    'redimencionamos el valor de la matriz a 8
    ReDim pMatMontos(prsDet.RecordCount, 7)
    
    prsDet.MoveFirst
    While Not prsDet.EOF
        '20060405
        'código obsoleto
        'pMatMontos(prsDet.Bookmark - 1, 0) = MontoColocadoCredCondicion(pdInicio, pdFinal, pnMoneda, prsDet!nTipoCred, psPersCod, gColocCredCondNormal, True)
        'pMatMontos(prsDet.Bookmark - 1, 1) = MontoColocadoCredCondicion(pdInicio, pdFinal, pnMoneda, prsDet!nTipoCred, psPersCod, gColocCredCondNormal)
        'pMatMontos(prsDet.Bookmark - 1, 2) = MontoColocadoCredCondicion(pdInicio, pdFinal, pnMoneda, prsDet!nTipoCred, psPersCod, gColocCredCondParalelo)
        'pMatMontos(prsDet.Bookmark - 1, 3) = MontoColocadoCredCondicion(pdInicio, pdFinal, pnMoneda, prsDet!nTipoCred, psPersCod, gColocCredCondRecurrente)
        'prsDet.MoveNext
        
        '20060405
        'nuevo código
        'creditos nuevos
        pMatMontos(prsDet.Bookmark - 1, 0) = MontoColocadoCredCondicion(pdInicio, pdFinal, pnMoneda, prsDet!nTipoCred, psPersCod, gColocCredCondNormal, True)
        
        nValor = MontoColocadoCredCondicion(pdInicio, pdFinal, pnMoneda, prsDet!nTipoCred, psPersCod, gColocCredCondNormal, , rsMetaAnalista)
        pMatMontos(prsDet.Bookmark - 1, 1) = rsMetaAnalista("nMontoCol")
        pMatMontos(prsDet.Bookmark - 1, 2) = rsMetaAnalista("nNumeroCred")
        
        'creditos paralelos
        nValor = MontoColocadoCredCondicion(pdInicio, pdFinal, pnMoneda, prsDet!nTipoCred, psPersCod, gColocCredCondParalelo, , rsMetaAnalista)
        pMatMontos(prsDet.Bookmark - 1, 3) = rsMetaAnalista("nMontoCol")
        pMatMontos(prsDet.Bookmark - 1, 4) = rsMetaAnalista("nNumeroCred")
        
        'creditos recurrentes
        nValor = MontoColocadoCredCondicion(pdInicio, pdFinal, pnMoneda, prsDet!nTipoCred, psPersCod, gColocCredCondRecurrente, , rsMetaAnalista)
        pMatMontos(prsDet.Bookmark - 1, 5) = rsMetaAnalista("nMontoCol")
        pMatMontos(prsDet.Bookmark - 1, 6) = rsMetaAnalista("nNumeroCred")
        
        prsDet.MoveNext
        
    Wend
    prsDet.MoveFirst
    Exit Sub

ErrorRecuperaDatosMetasAnalistaDetalle_Montos:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Sub

Public Sub CargarDatosDacion(ByVal pnDacion As Long, _
                            ByRef prs As ADODB.Recordset, _
                            ByRef prsCred As ADODB.Recordset, _
                            ByRef prsDet As ADODB.Recordset)

Dim sPersCod As String
On Error GoTo ErrorCargarDatosDacion

    Set prs = RecuperaDacionPago(pnDacion)
    If Not prs.BOF And Not prs.EOF Then
        sPersCod = prs!cPersCod
        Set prsCred = RecuperaCreditosVigentesCtas(sPersCod)
        Set prsDet = RecuperaDacionPagoDetalle(pnDacion)
    End If

    Exit Sub
ErrorCargarDatosDacion:
    Err.Raise Err.Number, "Cargar Datos Dacion", Err.Description
End Sub

Public Function ValidaTitularCredito(ByVal psPersCod As String) As String
Dim nNumCredJud As Integer
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset

'*** PEAC 20080813 - se comento porque se implemento un validador para el
'*** ingreso del visto electronico con respecto a clientes con historial negativo

'nNumCredJud = NumerosCredEnJudicial(psPersCod)
'If nNumCredJud > 0 Then
'    ValidaTitularCredito = "El Titular tiene creditos en judicial"
'    Exit Function
'End If

'*** FIN PEAC


sSql = "SELECT P.cPersCod FROM Persona P INNER JOIN RRHH RH ON P.cPersCod=RH.cPersCod " & _
       " AND nRHEstado NOT IN (101,501,701,702,703) INNER JOIN PersRelaciones PR " & _
       " ON P.cPersCod=PR.cPersRelacPersCod WHERE P.cPersCod='" & psPersCod & "'"

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing

If Not rs.EOF Then
    ValidaTitularCredito = "El Titular tiene parientes trabajando en la Empresa"
End If

End Function

Public Function VerificarClienteCreditos(ByVal psPersCod As String) As Boolean
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    SQL = " select COUNT (*) Moroso from Producto P " & _
          " JOIN ProductoPersona PP ON P.cCtaCod=PP.cCtacOD " & _
          " and nPrdPersRelac ='20'  and nPrdEstado in ('" & gColocEstRecVigJud & " ', '" & gColocEstRecVigCast & "') and PP.cPersCod='" & psPersCod & "'"
   Set rs = oCon.CargaRecordSet(SQL)
   If Not (rs.EOF And rs.BOF) Then
        If rs!Moroso > 0 Then
            VerificarClienteCreditos = True
        Else
            VerificarClienteCreditos = False
        End If
   End If

End Function

Public Function VerificarIndiceVAC(ByVal pdFecSis As Date) As Boolean
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    SQL = " SELECT * FROM IndiceVAC " & _
          " WHERE dIndiceVAC='" & Format(pdFecSis, "mm/dd/yyyy") & "'"
    Set rs = oCon.CargaRecordSet(SQL)
    VerificarIndiceVAC = Not rs.EOF
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ObtenerIndiceVAC(ByVal pdFecSis As Date) As Double
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    SQL = " SELECT nIndiceVac FROM IndiceVAC " & _
          " WHERE dIndiceVAC='" & Format(pdFecSis, "mm/dd/yyyy") & "'"
    Set rs = oCon.CargaRecordSet(SQL)
    If rs.EOF Then
        ObtenerIndiceVAC = 0
    Else
        ObtenerIndiceVAC = rs!nIndiceVac
    End If
    
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function RecuperaDatosMovDesembolso(ByVal psCtaCod As String, _
                                            ByRef pnMovNro As Long, _
                                            ByRef psOpeCod As String)

Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset

sSql = "SELECT nMovNro,cOpeCod FROM MovCol WHERE cCtaCod='" & psCtaCod & "' AND cOpeCod LIKE '1001%' "
Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSql)
If rs.EOF Then
    pnMovNro = 0
    psOpeCod = ""
Else
    pnMovNro = rs!nMovNro
    psOpeCod = rs!cOpeCod
    
End If
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function CargarDatosSugerencia(ByVal psCodCmac As String, ByRef prsCIUU As ADODB.Recordset, ByRef pnValorParamRCC As Double)
Dim oCIUU As COMDPersona.DCOMPersonas
Dim oParam As COMDCredito.DCOMParametro

Set oParam = New COMDCredito.DCOMParametro
pnValorParamRCC = oParam.RecuperaValorParametro(9004)
Set oParam = Nothing

Set oCIUU = New COMDPersona.DCOMPersonas
Set prsCIUU = oCIUU.Cargar_CIIU(psCodCmac)
Set oCIUU = Nothing

End Function

Public Function RecuperaSaldoDisponibleFteIngreso(ByVal psCtaCod As String) As Double
Dim oCon As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Dim sSql As String

'CTI6-20210503-ERS032-2019 -(Optimizar Sugerencia)
'sSql = "SELECT SaldoDisp= ((nPersFIVentas+nPersFIRecupCtasXCobrar+nPersIngFam)-(nPersFICostoVentas+nPersEgrFam+nPersFIEgresosOtros)) " & _
'       "     FROM ColocacCred C INNER JOIN ProductoPersona PP ON C.cCtaCod=PP.cCtaCod AND PP.nPrdPersRelac=20 " & _
'       " INNER JOIN Persona P ON PP.cPersCod=P.cPersCod INNER JOIN ColocFteIngreso CF ON C.cCtaCod=CF.cCtaCod " & _
'       " INNER JOIN PersFteIngreso F ON P.cPersCod=F.cPersCod AND F.cNumFuente = CF.cNumFuente" & _
'       " INNER JOIN PersFIIndependiente FI ON F.cNumFuente=FI.cNumFuente " & _
'       " WHERE C.cCtaCod='" & psCtaCod & "'"

sSql = "Execute spt_sel_RecuperaSaldoDisponibleFteIngreso '" & psCtaCod & "' "
'Fin CTI6-20210503

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSql)
If rs.EOF Then
    RecuperaSaldoDisponibleFteIngreso = 0
Else
    RecuperaSaldoDisponibleFteIngreso = rs!SaldoDisp
End If
oCon.CierraConexion
Set oCon = Nothing

End Function

Public Function RecuperaDatosCreditoECSA(ByVal psCtaCod As String, Optional ByVal pdHoy As Date = CDate("01/01/1900"), _
    Optional ByVal pbIncluirbAprobados As Boolean = False) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaDatosCreditoJudicial
    sSql = "Select PP.cPersCod, CC.cMetLiquidacion, dbo.CuotaAprobada(P.cCtaCod) as CuotaAprobada, CC.bCuotaCom, CC.bMiVivienda, CC.nCalendDinamTipo, CC.bPrepago, CC.nCalPago, Pers.cPersNombre, C.cLineaCred,C.nMontoCol, P.nSaldo, "
    sSql = sSql & " cMoneda = (Select cConsDescripcion from Constante where nConsCod = " & COMDConstantes.gMoneda & " AND nConsValor = " & CInt(Mid(psCtaCod, 9, 1)) & "),"
    sSql = sSql & " CE.nCuotas nCuotasApr, CE.nColocCalendCod " & ", "
    sSql = sSql & " CC.nDiasAtraso, CC.cMetLiquidacion, P.nTransacc, CC.nCalendDinamico, CC.nIntPend, CC.nNroCalen, CC.nNroProxCuota, CC.nNroProxDesemb "
    If pdHoy <> CDate("01/01/1900") Then
        sSql = sSql & ", DATEDIFF(year,C.dVigencia,'" & Format(pdHoy, "mm/dd/yyyy") & "') as nPlazoTranscurrido "
    End If
    sSql = sSql & " From Producto P Inner join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
    sSql = sSql & "                 Inner Join Persona Pers ON PP.cPersCod = Pers.cPersCod "
    sSql = sSql & "                 Inner Join Colocaciones C ON C.cCtaCod =  P.cCtaCod "
    sSql = sSql & "                 Inner Join ColocacCred CC ON CC.cCtaCod =  P.cCtaCod "
    sSql = sSql & "                 Inner Join ColocacEstado CE ON CE.cCtaCod = P.cCtaCod And CE.nPrdEstado = " & COMDConstantes.gColocEstAprob
    sSql = sSql & " WHERE P.cCtaCod = '" & psCtaCod & "'"

    If pbIncluirbAprobados Then
        sSql = sSql & " AND P.nPrdEstado in (" & COMDConstantes.gColocEstRefMor & "," & COMDConstantes.gColocEstRefNorm & "," & COMDConstantes.gColocEstRefVenc & "," & COMDConstantes.gColocEstVigMor & "," & COMDConstantes.gColocEstVigNorm & "," & COMDConstantes.gColocEstVigVenc & "," & COMDConstantes.gColocEstAprob & "," & COMDConstantes.gColocEstRecVigJud & ")"
    Else
        sSql = sSql & " AND P.nPrdEstado in (" & COMDConstantes.gColocEstRefMor & "," & COMDConstantes.gColocEstRefNorm & "," & COMDConstantes.gColocEstRefVenc & "," & COMDConstantes.gColocEstVigMor & "," & COMDConstantes.gColocEstVigNorm & "," & COMDConstantes.gColocEstVigVenc & "," & COMDConstantes.gColocEstRecVigJud & ")"
    End If

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosCreditoECSA = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaDatosCreditoJudicial:
    Err.Raise Err.Number, "Error En Proceso RecuperaDatosCreditoJudicial", Err.Description

End Function

'****17-05-2006
Public Function RecuperaDetallePagoHistorial(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
'sSQL = "SELECT dFecPago=dbo.FechaMov(M.cMovNro)," & _
'        "nNroCuota=(SELECT MAX(nNroCuota) FROM MovColDet WHERE nMovNro=M.nMovNro  AND nNroCuota>0), " & _
'        " nMontoPagado=ISNULL( (SELECT SUM(nMonto) FROM MovColDet WHERE nMovNro=M.nMovNro " & _
'        "    AND (nPrdConceptoCod IN(1000,1100,1102,1103,1104,1105,1106,3100,1101,3101,20) OR nPrdConceptoCod LIKE '12%')" & _
'        " AND (nNroCuota>0 OR nPrdConceptoCod=20) ),0), " & _
'        " nCapital= ISNULL( (SELECT SUM(nMonto)FROM MovColDet WHERE nMovNro=M.nMovNro AND nPrdConceptoCod=1000 " & _
'        " AND cOpeCod=M.cOpeCod AND nNroCuota>0),0), " & _
'        " nInteres=ISNULL( (SELECT SUM(nMonto)FROM MovColDet WHERE nMovNro=M.nMovNro " & _
'        " AND nPrdConceptoCod IN (1100,1102,1103,1104,1105,1106,3100) AND nNroCuota>0),0), " & _
'        " nMora= ISNULL( (SELECT SUM(nMonto)FROM MovColDet WHERE nMovNro=M.nMovNro AND nPrdConceptoCod IN(1101,3101) " & _
'        " AND nNroCuota>0),0), " & _
'        " nGastos= ISNULL( (SELECT SUM(nMonto)FROM MovColDet WHERE nMovNro=M.nMovNro AND nPrdConceptoCod LIKE '12%' " & _
'        " AND nNroCuota>0) ,0) , " & _
'        " nitf= ISNULL( (SELECT SUM(nMonto)FROM MovColDet WHERE nMovNro=M.nMovNro AND nPrdConceptoCod=20  ),0 ), " & _
'        " nDiasMora= ISNULL( (SELECT TOP 1 nDiasMora FROM MovCol WHERE nMovNro=M.nMovNro AND cOpeCod=M.cOpeCod ORDER BY nMovNro DESC),0 ), " & _
'        " nSaldoCap= ISNULL( (SELECT TOP 1 nSaldoCap FROM MovCol WHERE nMovNro=M.nMovNro AND cOpeCod=M.cOpeCod ORDER by nMovnro DESC), 0), " & _
'        " cUsuario = Right(M.cMovnro, 4) " & _
'        " FROM Mov M INNER JOIN MovCol MC ON M.nMovNro=MC.nMovNro AND M.nMovEstado=10 AND nMovFlag<>2 " & _
'        " AND M.cOpeCod=MC.cOpeCod AND MC.nNroCalen=(select nNroCalen from ColocacCred where cCtaCod = MC.cCtaCod) " & _
'        " AND M.cOpeCod LIKE '100[234567]%' " & _
'        " WHERE MC.cCtaCod='" & psCtacod & "'" & _
'        " GROUP BY M.nMovNro,M.cMovNro,M.cOpeCod,MC.nNroCalen ORDER BY M.nMovNro"

'sSQL = "SELECT dFecPago=dbo.FechaMov(Select top 1 cMovnro From Mov Where nMovnro = MC.nMovnro)," & _
'        "nNroCuota=(SELECT MAX(nNroCuota) FROM MovColDet WHERE nMovNro=MC.nMovNro  AND nNroCuota>0), " & _
'        " nMontoPagado=ISNULL( (SELECT SUM(nMonto) FROM MovColDet WHERE nMovNro=MC.nMovNro " & _
'        "    AND (nPrdConceptoCod IN(1000,1100,1102,1103,1104,1105,1106,3100,1101,3101,20) OR nPrdConceptoCod LIKE '12%')" & _
'        " AND (nNroCuota>0 OR nPrdConceptoCod=20) ),0), " & _
'        " nCapital= ISNULL( (SELECT SUM(nMonto)FROM MovColDet WHERE nMovNro=MC.nMovNro AND nPrdConceptoCod=1000 " & _
'        " AND cOpeCod=MC.cOpeCod AND nNroCuota>0),0), " & _
'        " nInteres=ISNULL( (SELECT SUM(nMonto)FROM MovColDet WHERE nMovNro=MC.nMovNro " & _
'        " AND nPrdConceptoCod IN (1100,1102,1103,1104,1105,1106,3100) AND nNroCuota>0),0), " & _
'        " nMora= ISNULL( (SELECT SUM(nMonto)FROM MovColDet WHERE nMovNro=MC.nMovNro AND nPrdConceptoCod IN(1101,3101) " & _
'        " AND nNroCuota>0),0), " & _
'        " nGastos= ISNULL( (SELECT SUM(nMonto)FROM MovColDet WHERE nMovNro=MC.nMovNro AND nPrdConceptoCod LIKE '12%' " & _
'        " AND nNroCuota>0) ,0) , " & _
'        " nitf= ISNULL( (SELECT SUM(nMonto)FROM MovColDet WHERE nMovNro=MC.nMovNro AND nPrdConceptoCod=20  ),0 ), " & _
'        " nDiasMora= ISNULL( (SELECT TOP 1 nDiasMora FROM MovCol WHERE nMovNro=MC.nMovNro AND cOpeCod=MC.cOpeCod ORDER BY nMovNro DESC),0 ), " & _
'        " nSaldoCap= ISNULL( (SELECT TOP 1 nSaldoCap FROM MovCol WHERE nMovNro=MC.nMovNro AND cOpeCod=MC.cOpeCod ORDER by nMovnro DESC), 0), " & _
'        " cUsuario = Right(Select top 1 cMovnro From Mov Where nMovnro = MC.nMovnro, 4) " & _
'        " FROM MovCol MC " & _
'        " WHERE MC.cCtaCod='" & psCtacod & "'" & _
'        " GROUP BY MC.nMovNro,MC.cOpeCod,MC.nNroCalen ORDER BY MC.nMovNro"
        
        'ARCV 12-06-2007 : Se agrego el Nro de credito en el Nro de cuota
'        sSql = " SELECT MC.cOpeCod, dFecPago = (Select TOP 1 Substring(cMovNro,7,2) + '/' + "
'        sSql = sSql & " SUBSTRING(cMovNro,5,2 ) + '/' + SUBSTRING(cMovNro,1,4) From Mov where nMovnro=MC.nMovnro), "
'        sSql = sSql & " nNroCuota=(SELECT MAX(nNroCuota) FROM MovColDet WHERE nMovNro=MC.nMovNro  AND ((nNroCuota=0 and substring(MC.cCtaCod,6,3) in ('515','516') and nPrdConceptoCod<>20) or nNroCuota>0) AND cCtaCod='" & psCtaCod & "'), "
'        sSql = sSql & " nMontoPagado=ISNULL( (SELECT SUM(CASE WHEN nPrdConceptoCod=1106 THEN nMonto*-1 ELSE nMonto END) FROM MovColDet WHERE cCtaCod = MC.cCtaCod AND nMovNro=MC.nMovNro     AND (nPrdConceptoCod IN(1000,1100,1102,1103,1104,1105,1106,3100,1101,3101,20) OR nPrdConceptoCod LIKE '12%') AND ((nNroCuota=0 and substring(MC.cCtaCod,6,3) in ('515','516') and nPrdConceptoCod<>20) or nNroCuota>0 OR nPrdConceptoCod=20)),0),"
'        sSql = sSql & " nCapital= ISNULL( (SELECT SUM(nMonto)FROM MovColDet WHERE cCtaCod = MC.cCtaCod AND nMovNro=MC.nMovNro AND nPrdConceptoCod=1000  AND cOpeCod=MC.cOpeCod AND ((nNroCuota=0 and substring(MC.cCtaCod,6,3) in ('515','516') and nPrdConceptoCod<>20) or nNroCuota>0)),0),"
'        sSql = sSql & " nInteres=ISNULL( (SELECT SUM(CASE WHEN nPrdConceptoCod=1106 THEN nMonto*-1 ELSE nMonto END)FROM MovColDet WHERE cCtaCod = MC.cCtaCod AND nMovNro=MC.nMovNro  AND nPrdConceptoCod IN (1100,1102,1103,1104,1105,1106,3100) AND nNroCuota>0),0),"
'        sSql = sSql & " nMora= ISNULL( (SELECT SUM(nMonto)FROM MovColDet WHERE cCtaCod = MC.cCtaCod AND nMovNro=MC.nMovNro AND nPrdConceptoCod IN(1101,3101)  AND ((nNroCuota=0 and substring(MC.cCtaCod,6,3) in ('515','516') and nPrdConceptoCod<>20) or nNroCuota>0)),0),"
'        sSql = sSql & " nGastos= ISNULL( (SELECT SUM(nMonto)FROM MovColDet WHERE cCtaCod = MC.cCtaCod AND nMovNro=MC.nMovNro AND nPrdConceptoCod LIKE '12%'  AND ( (nNroCuota=0 and substring(MC.cCtaCod,6,3) in ('515','516') and nPrdConceptoCod<>20) or nNroCuota>0)) ,0) ,"
'        'EJVG20140930 ***
'        sSql = sSql & " nGastoComision= ISNULL((SELECT SUM(nMonto)FROM MovColDet WHERE cCtaCod = MC.cCtaCod AND nMovNro=MC.nMovNro AND nPrdConceptoCod LIKE '12%' AND ((nNroCuota=0 and substring(MC.cCtaCod,6,3) in ('515','516') and nPrdConceptoCod<>20) or nNroCuota>0) AND nPrdConceptoCod Not In (1217,1231,1243)),0.00),"
'        sSql = sSql & " nGastoSegDes=ISNULL((SELECT SUM(nMonto)FROM MovColDet WHERE cCtaCod = MC.cCtaCod AND nMovNro=MC.nMovNro AND nPrdConceptoCod LIKE '12%' AND ((nNroCuota=0 and substring(MC.cCtaCod,6,3) in ('515','516') and nPrdConceptoCod<>20) or nNroCuota>0) AND nPrdConceptoCod = 1217),0.00),"
'        sSql = sSql & " nGastoPolizaIncendio=ISNULL((SELECT SUM(nMonto)FROM MovColDet WHERE cCtaCod = MC.cCtaCod AND nMovNro=MC.nMovNro AND nPrdConceptoCod LIKE '12%' AND ((nNroCuota=0 and substring(MC.cCtaCod,6,3) in ('515','516') and nPrdConceptoCod<>20) or nNroCuota>0) AND nPrdConceptoCod = 1231),0.00),"
'        sSql = sSql & " nGastoPolizaVehiculo=ISNULL((SELECT SUM(nMonto)FROM MovColDet WHERE cCtaCod = MC.cCtaCod AND nMovNro=MC.nMovNro AND nPrdConceptoCod LIKE '12%' AND ((nNroCuota=0 and substring(MC.cCtaCod,6,3) in ('515','516') and nPrdConceptoCod<>20) or nNroCuota>0) AND nPrdConceptoCod = 1243),0.00),"
'        'END EJVG *******
'        sSql = sSql & " nitf= ISNULL( (SELECT SUM(nMonto)FROM MovColDet WHERE cCtaCod = MC.cCtaCod AND nMovNro=MC.nMovNro AND nPrdConceptoCod=20  ),0 ),"
'        sSql = sSql & " nDiasMora= ISNULL( (SELECT TOP 1 nDiasMora FROM MovCol WHERE cCtaCod = MC.cCtaCod AND nMovNro=MC.nMovNro AND cOpeCod=MC.cOpeCod ORDER BY nMovNro DESC),0 ),"
'        sSql = sSql & " nSaldoCap= ISNULL( (SELECT TOP 1 nSaldoCap FROM MovCol WHERE cCtaCod = MC.cCtaCod AND nMovNro=MC.nMovNro AND cOpeCod=MC.cOpeCod ORDER by nMovnro DESC), 0),"
'        sSql = sSql & " cUsuario = SUBSTRING((Select top 1 cMovnro From Mov where nMovnro=MC.nMovnro),22,4 )"
'        sSql = sSql & " FROM MovCol MC INNER JOIN Mov M ON M.nMovNro=MC.nMovNro"
'        sSql = sSql & " WHERE MC.cCtaCod='" & psCtaCod & "' and (MC.cOpeCod LIKE '100[234567]%' OR MC.cOpeCod = '105001') AND M.nMovFlag<>2"
'        sSql = sSql & " GROUP BY MC.cCtaCod ,MC.nMovNro,MC.cOpeCod,MC.nNroCalen ORDER BY MC.nMovNro"
    'ALPA20160505**************
    '**************************
    sSql = "exec stp_sel_ObtenerHistorialPagosPosCli'" & psCtaCod & "'"
    'ALPAFIN
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set RecuperaDetallePagoHistorial = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
'**************


Public Function CargarCargosAprobacion() As ADODB.Recordset
Dim Conex As COMConecta.DCOMConecta
Dim sSql As String
    Set Conex = New COMConecta.DCOMConecta
    sSql = "Select cCodCargo, cNomCargo From CredCargos order by cNomCargo"
    Conex.AbreConexion
    Set CargarCargosAprobacion = Conex.CargaRecordSet(sSql)
    Conex.CierraConexion
    Set Conex = Nothing
End Function

Public Function RecuperaNomUbigeoAgencia(ByVal psCtaCod As String) As String
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As New ADODB.Recordset

    sSql = "Select cUbiGeoDescripcion from ubicacionGeografica where cUbigeoCod = (select cUbiGeoCod from AGENCIAS WHERE cAgeCod = '" & Mid(psCtaCod, 4, 2) & "')"
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    If Not (rs.EOF And rs.BOF) Then
        RecuperaNomUbigeoAgencia = rs!cUbiGeoDescripcion
    Else
        RecuperaNomUbigeoAgencia = ""
    End If
    oConec.CierraConexion
    Set oConec = Nothing
    Set rs = Nothing

End Function

Public Function RecuperaGarantes(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    'FRHU 20150303 MEMO 491-2015
    'sSql = " select cPersNombre, cPersDireccDomicilio,"
    'sSql = sSql & " IsNull((Select cPersIDnro from PersId where cPersIDTpo = '1' and cPersCod = P.cPersCod),'') DNI,"
    'sSql = sSql & " IsNull((Select cPersIDnro from PersId where cPersIDTpo = '2' and cPersCod = P.cPersCod),'') RUC"
    'sSql = sSql & " from ProductoPersona PP"
    'sSql = sSql & " Inner Join Persona P on PP.cPersCod = P.cPersCod and nPrdPersRelac = 25"
    'sSql = sSql & " where cctacod = '" & psCtaCod & "'"
     sSql = "exec stp_sel_RecuperaGarantes '" & psCtaCod & "'"
    'FIN FRHU 20150303
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set RecuperaGarantes = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing

End Function

Public Function VerificarClienteCredMorosos(ByVal psPersCod As String) As Boolean
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    SQL = " select COUNT (*) Moroso from Producto P " & _
          " JOIN ProductoPersona PP ON P.cCtaCod=PP.cCtacOD " & _
          " and nPrdPersRelac ='20'  and nPrdEstado in ('" & gColocEstVigMor & " ') and PP.cPersCod='" & psPersCod & "'"
   Set rs = oCon.CargaRecordSet(SQL)
   If Not (rs.EOF And rs.BOF) Then
        If rs!Moroso > 0 Then
            VerificarClienteCredMorosos = True
        Else
            VerificarClienteCredMorosos = False
        End If
   End If

End Function

Public Function RecuperaDatosFuenteingreso(ByVal pcNumFuente As String, _
                                            ByVal pnPersTipFte As Integer, _
                                            ByVal pdPersFecEval As Date) As ADODB.Recordset
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
        If pnPersTipFte = gPersFteIngresoTipoIndependiente Then
            sSql = "SELECT  PF.cNumFuente, PF.dPersEval, PF.dPersFICaduca, PF.nPersFIActivoDisp, "
            sSql = sSql & " PF.nPersFICtasxCobrar, PF.nPersFIInventarios, PF.nPersFIActivosFijos, "
            sSql = sSql & " PF.nPersFIProveedores, PF.nPersFICredCMACT, PF.nPersFICredOtros, "
            sSql = sSql & " PF.nPersFIVentas, PF.nPersFIPatrimonio, PF.nPersFICostoVentas, "
            sSql = sSql & " PF.nPersFIRecupCtasXCobrar, PF.nPersFIEgresosOtros, PF.nPersIngFam, "
            sSql = sSql & " PF.nPersEgrFam, PF.cUltimaActualizacion,FI.cPersfimoneda "
            sSql = sSql & " FROM PersFIIndependiente PF INNER JOIN PersFteIngreso FI ON FI.cNumFuente = PF.cNumFuente"
            sSql = sSql & " Where PF.cNumFuente = '" & pcNumFuente & "'"
            sSql = sSql & " AND PF.dPersEval='" & Format(pdPersFecEval, "mm/dd/yyyy") & "'"
        Else
            sSql = " Select PF.cNumFuente, PF.dPersEval, PF.nPersIngCli, PF.nPersGastoFam, PF.nPersIngCon, PF.nPersOtrIng, PF.dPersFICaduca "
            sSql = sSql & " ,FI.cPersFIMoneda,FI.cPersFIComentario,FI.cPersFICargo,FI.dPersFIInicio"
            sSql = sSql & " FROM PersFIDependiente PF "
            sSql = sSql & " INNER JOIN PersFteIngreso FI ON PF.cNumFuente= FI.cNumFuente "
            sSql = sSql & " Where PF.cNumFuente = '" & pcNumFuente & "'"
            sSql = sSql & " AND PF.dPersEval='" & Format(pdPersFecEval, "mm/dd/yyyy") & "'"
        End If
Set RecuperaDatosFuenteingreso = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function RecuperaDatosFlujoCaja(ByVal psCtaCod As String) As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim sSql As String

sSql = "SELECT CFI.cCtaCod,Per.cPersNombre,FII.cNumFuente, FII.dPersEval, dPersFICaduca, nPersFIActivoDisp, nPersFICtasxCobrar, nPersFIInventarios, nPersFIActivosFijos, nPersFIProveedores, nPersFICredCMACT, nPersFICredOtros, nPersFIVentas, nPersFIPatrimonio, nPersFICostoVentas, nPersFIRecupCtasXCobrar, " & _
       " nPersFIEgresosOtros=CASE WHEN ISNULL(nPersFIEgresosOtros,0)=0 THEN 0.01 ELSE nPersFIEgresosOtros END, nPersIngFam, nPersEgrFam, FII.cUltimaActualizacion " & _
        " FROM ColocFteIngreso CFI INNER JOIN Producto P ON CFI.cCtaCod= P.cCtaCod" & _
       " INNER JOIN PersFIIndependiente FII ON FII.cNumFuente = CFI.cNumFuente " & _
        " AND FII.dPersEval=(SELECT MAX(dPersEval)FROM PersFIIndependiente " & _
                " WHERE cNumFuente=FII.cNumFuente )" & _
       " INNER JOIN PersFteIngreso FI ON FII.cNumFuente = FI.cNumFuente  INNER JOIN Persona Per ON FI.cPersCod= Per.cPersCod " & _
       " WHERE P.cCtaCod='" & psCtaCod & "'"

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set RecuperaDatosFlujoCaja = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing

End Function

'**DAOR 20070202
'**Recupera datos para desembolso en el banco de la nación
Public Function RecuperaDatosDesembolsoBcoNac(ByVal psCtaCod As String) As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
Dim lsSQL As String

    On Error GoTo ErrorRecuperaDatosDesembolsoBcoNac
    'CTI6-20210503-ERS032-2019 -(Optimizar Sugerencia)
    'lsSQL = "Select cClave,cCodAgeBcoNac,nEstado from ColocConvBcoNac  "
    'lsSQL = lsSQL & " where cCtaCod = '" & psCtaCod & "'"
    
    lsSQL = "Execute spt_sel_RecuperaDatosDesembolsoBcoNac '" & psCtaCod & "'"
    
    'Fin CTI6-20210503
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosDesembolsoBcoNac = oConecta.CargaRecordSet(lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaDatosDesembolsoBcoNac:
    Err.Raise Err.Number, "Error En Proceso : RecuperaDatosDesembolsoBcoNac", Err.Description

End Function

'**DAOR 20070202
'**Obtiene la Tasa de Interes de una Cuenta
Public Function dObtieneProductoTasaInteres(ByVal psCtaCod As String, ByVal pnPrdTasaInteres As Integer) As Double
Dim oConecta As COMConecta.DCOMConecta
Dim lrs As ADODB.Recordset
Dim lsSQL As String

On Error GoTo ErrordObtieneProductoTasaInteres

    Set lrs = New Recordset
    Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion

        lsSQL = "select nTasaInteres from ProductoTasaInteres where cCtaCod ='" & psCtaCod & "' and nPrdTasaInteres =" & pnPrdTasaInteres & " "

    Set lrs = oConecta.CargaRecordSet(lsSQL)

        If lrs.BOF And lrs.EOF Then
            dObtieneProductoTasaInteres = 0
        Else
            dObtieneProductoTasaInteres = lrs("nTasaInteres")
        End If

        oConecta.CierraConexion
    Set lrs = Nothing
    Set oConecta = Nothing
    Exit Function

ErrordObtieneProductoTasaInteres:
    Err.Raise Err.Number, "Error en Obtener Tasa de Interes : dObtieneProductoTasaInteres ", Err.Description
End Function

'**DAOR 20070404
'**Función que recupera la Linea y Tasas de Interes de Creditos que han sido
'**refinanciados, Tomando como parametro la Cuenta del Nuevo Credito Refinanciado
Public Function RecuperaLineaTasaAnteriorDeRefinanciado(psCtaCod As String) As ADODB.Recordset
On Error GoTo ErrorRecuperaLineaTasaAnteriorDeRefinanciado
Dim oConecta As COMConecta.DCOMConecta
Dim lsSQL As String

   lsSQL = "Select  CR.cCtaCodRef,C.cLineaCred, " & _
    " nTasaComp=(select nTasaInteres from ProductoTasaInteres where nPrdTasaInteres=" & COMDConstantes.gColocLineaCredTasasIntCompNormal & " and cCtaCod=CR.cCtaCodRef), " & _
    " nTasaMora=(select nTasaInteres from ProductoTasaInteres where nPrdTasaInteres=" & COMDConstantes.gColocLineaCredTasasIntMoratNormal & " and cCtaCod=CR.cCtaCodRef), " & _
    " nTasaGracia=(select nTasaInteres from ProductoTasaInteres where nPrdTasaInteres=" & COMDConstantes.gColocLineaCredTasasIntGracia & " and cCtaCod=CR.cCtaCodRef) " & _
    " From ColocacRefinanc CR inner join Colocaciones C on CR.cCtaCodRef=C.cCtaCod " & _
    " Where CR.cCtaCod='" & psCtaCod & "' and CR.nEstado=(select max(nEstado) from ColocacRefinanc where cCtaCod='" & psCtaCod & "') "
    
    Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
    Set RecuperaLineaTasaAnteriorDeRefinanciado = oConecta.CargaRecordSet(lsSQL)
        oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaLineaTasaAnteriorDeRefinanciado:
    Err.Raise Err.Number, "Error en Recuperar Linea y Tasas de Interes de Credito Antiguo de Refinanciación : RecuperaLineaTasaAnteriorDeRefinanciado ", Err.Description
End Function

'**DAOR 20070410, Función que recupera datos necesarios para generar calendario de gastos
'**en proceso de reprogramación
Public Function RecuperaDatosParaGenerarGastosEnReprog(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer

On Error GoTo ErrorRecuperaDatosParaGenerarGastosEnReprog
    sSql = "Select ISNULL(CC.nExoSeguroDes,0) as nExoSeguroDes,CE.cTipoGasto,CE.nColocCalendCod "
    sSql = sSql & " From ColocacCred CC inner join ColocacEstado CE on CC.cCtaCod=CE.cCtaCod "
    sSql = sSql & " Where CE.cCtaCod = '" & psCtaCod & "' and CE.nPrdEstado =" & COMDConstantes.gColocEstAprob
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosParaGenerarGastosEnReprog = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorRecuperaDatosParaGenerarGastosEnReprog:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'**DAOR 20070512, Función que recupera los datos para generar el estado de cuenta
'**de un crédito
Public Function RecuperaDatosParaEstadoCuentaCredito(ByVal psCtaCod As String, Optional ByVal pMatProd As Variant, _
    Optional pMatAgencias As Variant, Optional pdDesembDe As Date, Optional pdDesembHasta As Date) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
Dim sCadProd As String
Dim sCadAge As String
    
    On Error GoTo ErrorRecuperaDatosParaEstadoCuentaCredito
    

        'sCadProd = " And SUBSTRING(P.cCtaCod,6,3) in ('"
        For i = 0 To UBound(pMatProd) - 1
            'sCadProd = sCadProd & pMatprod(i) & "','"
            sCadProd = sCadProd & pMatProd(i) & ","
        Next i
        'sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
        sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 1)
        
        'sCadAge = " And substring(P.cCtaCod,4,2) in ('"
        For i = 0 To UBound(pMatAgencias) - 1
            'sCadAge = sCadAge & pMatAgencias(i) & "','"
            sCadAge = sCadAge & pMatAgencias(i) & ","
        Next i
        'sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
        sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)
        
        'By Capi 28122007 se adiciono columna otros gastos
        sSql = " exec stp_sel_ReporteEstadoCuentaCredito '" & psCtaCod & "','" & sCadProd & "','" & sCadAge & "','" & Format(pdDesembDe, "yyyymmdd") & "','" & Format(pdDesembHasta, "yyyymmdd") & "'"
        

        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set RecuperaDatosParaEstadoCuentaCredito = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing

        Exit Function
ErrorRecuperaDatosParaEstadoCuentaCredito:
    Err.Raise Err.Number, "Error En Proceso (RecuperaDatosParaEstadoCuenta)", Err.Description

End Function

'**DAOR 20070920, Creditos para asignación de gasto en lote, sólo considera a créditos vigentes
Public Function RecuperaCreditosParaAsignarGastoLote(ByVal pnRanIni As Double, ByVal pnRanFin As Double, _
    ByVal pnAplicado As COMDConstantes.ColocCalendApl, ByVal psCodGasto As String, pdDesembDe As Date, pdDesembHasta As Date, _
    Optional ByVal psTipoProd As String = "", Optional ByVal psMoneda As String = "") As ADODB.Recordset
Dim sSql As String
Dim lsMoneda As String
Dim lsTipoProd As String
Dim oConecta As COMConecta.DCOMConecta
    
    If psTipoProd <> "" Then
        lsTipoProd = " and substring(P.cCtaCod,6,1) = '" & psTipoProd & "' "
    End If
    If psMoneda <> "" Then
        lsMoneda = " and substring(P.cCtaCod,9,1)='" & psMoneda & "' "
    End If

    On Error GoTo ErrorRecuperaCreditosParaAsignarGastoLote
    sSql = "select '',P.cCtaCod,RTRIM(PE.cPersNombre) as cPersNombre, P.nSaldo,CL.nCuota,sum(CD.nMonto-CD.nMontoPagado) as nMontoPendiente,CC.nNroCalen"
    sSql = sSql & " from Producto P inner join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
    sSql = sSql & "     inner join Persona PE ON PP.cPersCod = PE.cPersCod "
    sSql = sSql & "     inner join Colocaciones C ON P.cCtaCod = C.cCtaCod"
    sSql = sSql & "     inner join ColocacCred CC ON P.cCtaCod = CC.cCtaCod"
    sSql = sSql & "     inner join ColocCalendario CL ON P.cCtaCod = CL.cCtaCod AND CL.nNroCalen=CC.nNroCalen "
    sSql = sSql & "         and CL.nColocCalendApl = 1 AND CL.nColocCalendEstado = 0  and CC.nNroProxCuota=CL.nCuota "
    sSql = sSql & "     inner join ColocCalendDet CD ON P.cCtaCod=CD.cCtaCod AND CD.nNroCalen = CC.nNroCalen "
    sSql = sSql & "         and CD.nColocCalendApl = 1  and CL.nCuota=CD.nCuota "
    sSql = sSql & " where   P.nPrdEstado in (" & gColocEstVigNorm & "," & gColocEstVigMor & "," & gColocEstVigVenc & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor & ") "
    sSql = sSql & "     and CC.nNroProxCuota not in(select CDT.nCuota from ColocCalendDet CDT where CDT.cCtaCod=CC.cCtaCod and CDT.nNroCalen=CC.nNroCalen and CDT.nColocCalendApl=1 and CDT.nCuota=CC.nNroProxCuota and CDT.nPrdConceptoCod=" & psCodGasto & ")"
    sSql = sSql & "     and C.dVigencia between '" & Format(pdDesembDe, "yyyymmdd") & "' and '" & Format(pdDesembHasta, "yyyymmdd") & "' "
    sSql = sSql & lsTipoProd
    sSql = sSql & lsMoneda
    sSql = sSql & " group by P.cCtaCod, PE.cPersNombre, P.nSaldo, CL.nCuota,CC.nNroCalen "

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditosParaAsignarGastoLote = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCreditosParaAsignarGastoLote:
    Err.Raise Err.Number, "Error En Proceso: RecuperaCreditosParaAsignarGastoLote ", Err.Description

End Function

'**DAOR 20070926, Recupera cuota pendiente de pago para asignar gasto
Public Function RecuperaCuotaPendienteParaAsignarGasto(ByVal pnRanIni As Double, ByVal pnRanFin As Double, _
    ByVal pnAplicado As COMDConstantes.ColocCalendApl, ByVal psCtaCod As String, ByVal psCodGasto As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCuotaPendienteParaAsignarGasto
    
    sSql = "select P.cCtaCod,P.nSaldo,CC.nNroCalen,CL.nCuota,sum(CD.nMonto-isnull(CD.nMontoPagado,0)) as nMontoPendiente"
    sSql = sSql & " from Producto P inner join ColocacCred CC ON P.cCtaCod = CC.cCtaCod "
    sSql = sSql & "     inner join ColocCalendario CL ON P.cCtaCod = CL.cCtaCod AND CL.nNroCalen=CC.nNroCalen "
    sSql = sSql & "         and CL.nColocCalendApl = 1 AND CL.nColocCalendEstado = 0  and CC.nNroProxCuota=CL.nCuota "
    sSql = sSql & "     inner join ColocCalendDet CD ON P.cCtaCod=CD.cCtaCod AND CD.nNroCalen = CC.nNroCalen "
    sSql = sSql & "         and CD.nColocCalendApl = 1  and CL.nCuota=CD.nCuota "
    sSql = sSql & " where   P.nPrdEstado in (" & gColocEstVigNorm & "," & gColocEstVigMor & "," & gColocEstVigVenc & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor & ") "
    sSql = sSql & "     and CC.nNroProxCuota not in(select CDT.nCuota from ColocCalendDet CDT where CDT.cCtaCod=CC.cCtaCod and CDT.nNroCalen=CC.nNroCalen and CDT.nColocCalendApl=1 and CDT.nCuota=CC.nNroProxCuota and CDT.nPrdConceptoCod=" & psCodGasto & ") "
    sSql = sSql & "     and P.cCtaCod='" & psCtaCod & "' "
    sSql = sSql & " group by P.cCtaCod,CC.nNroCalen,P.nSaldo, CL.nCuota "

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCuotaPendienteParaAsignarGasto = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCuotaPendienteParaAsignarGasto:
    Err.Raise Err.Number, "Error En Proceso: RecuperaCuotaPendienteParaAsignarGasto", Err.Description

End Function

'**DAOR 20071207, Recupera personas relacionadas al créditos
'**Para representante del seguro de desgravamen (Solo para persona jurídica)
Public Function RecuperaPersRelacCredDesgrav(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaPersRelacCredDesgrav

    sSql = "exec stp_sel_PersonasRelacCredDesgravParaCombo '" & psCtaCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaPersRelacCredDesgrav = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaPersRelacCredDesgrav:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'*** PEAC 20080402
Public Function CargarObjetosControles(ByRef prsGrupoHojEval As ADODB.Recordset, Optional nTipoEval As Integer = 2)
    
    Dim oConec As COMConecta.DCOMConecta
    Dim oFinan As COMDPersona.DCOMInstFinac
    Dim oCons As COMDConstantes.DCOMConstantes
    Dim oUbic As COMDPersona.DCOMPersonas
    
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrorCargarObjetosControles
    
    Set prsGrupoHojEval = ListaGrupoHojEval(nTipoEval)
    'Set prsTipoEvaluacion = ListaTipoEvaluacion
    
    
    Exit Function
ErrorCargarObjetosControles:
    Set oConec = Nothing
    
End Function

'***PEAC 20080402
Public Function CargarRelEval(ByVal pIsSuperTipo As String, Optional nTipEval As Integer = 2) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrHandler

    sSql = "exec stp_sel_RecuperaGrupoHojEvalNivel1 3,'" & nTipEval & "','" & pIsSuperTipo & "'"
            
     Set oConec = New COMConecta.DCOMConecta
     oConec.AbreConexion
     Set CargarRelEval = oConec.CargaRecordSet(sSql)
     oConec.CierraConexion
     Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set CargarRelEval = Null
End Function

'***PEAC 20080402
Public Function CargarGruEval(ByVal pIsSuperGruEval As String, Optional nTipoEval As Integer = 2) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrHandler

    sSql = "exec stp_sel_RecuperaGrupoHojEvalNivel1 2,'" & nTipoEval & "','" & pIsSuperGruEval & "'"


     Set oConec = New COMConecta.DCOMConecta
     oConec.AbreConexion
     Set CargarGruEval = oConec.CargaRecordSet(sSql)
     oConec.CierraConexion
     Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set CargarGruEval = Null
End Function


'*** PEAC 20080402
Public Function ListaTipoEvaluacion() As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrHandler

    sSql = "exec stp_sel_RecuperaTipoEvaluacion"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaTipoEvaluacion = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set ListaTipoEvaluacion = Null
End Function
'*** PEAC 20080402
Public Function ListaGrupoHojEval(Optional nTEva As Integer = 2) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrHandler

    sSql = "exec stp_sel_RecuperaGrupoHojEvalNivel1 1,'" & nTEva & "'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaGrupoHojEval = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set ListaGrupoHojEval = Null
End Function



'*** PEAC 20080402
Public Function CargarTipoMone(ByVal pIsSuperTipoMone As Integer) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrHandler

    sSql = "exec stp_sel_RecuperaTipoImporte " & pIsSuperTipoMone

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set CargarTipoMone = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set CargarTipoMone = Null
End Function
'**ALPA******20080410****************************************************************************
Public Function CargarHojaEval(ByVal sNFuFinan As String, ByVal dFechaFin As Date) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrHandler

    sSql = "exec stp_sel_ReportesHojaEvaluacion '" & sNFuFinan & "','" & Format(dFechaFin, "yyyymmdd") & "'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set CargarHojaEval = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set CargarHojaEval = Null
End Function
'***End********************************************************************************************
'***PEAC 20080412
Public Function RecuperaDatosAprobacionCreditos(ByVal pNumCred As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "exec stp_sel_ObtieneDatosAprobacionCreditos '" & pNumCred & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set RecuperaDatosAprobacionCreditos = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function

'**ALPA******20080413****************************************************************************
Public Function CargarMaquetaHojaEval(Optional ByVal nTipoEval As Integer = 2) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrHandler

    sSql = "exec stp_sel_MaquetaHojaEvaluacion '" & nTipoEval & "'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set CargarMaquetaHojaEval = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set CargarMaquetaHojaEval = Null
End Function
'***End********************************************************************************************
'**ALPA******20080423****************************************************************************
Public Function ReportesHojaEvaluacionRatios(ByVal cNumFuente As String, ByVal dFecha As Date, Optional nTipoEval As Integer = 2) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrHandler

    sSql = "exec stp_sel_ReportesHojaEvaluacionRatios '" & cNumFuente & "','" & Format(dFecha, "YYYY/mm/dd") & "','" & nTipoEval & "'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ReportesHojaEvaluacionRatios = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set ReportesHojaEvaluacionRatios = Null
End Function
'**End********************************************************************************************

'**ALPA 20080512****************************************************************************
Public Function ReportesTransferenciaCartera(ByVal lsCtaCod As Variant) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrHandler

    sSql = "exec stp_sel_ObtenerListaTransferenciaCartera '" & lsCtaCod & "'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ReportesTransferenciaCartera = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set ReportesTransferenciaCartera = Null
End Function
'**End********************************************************************************************

'**ALPA 20080512****************************************************************************
Public Function ReporteRentabilidadCarteraXAnalista(ByVal dFechaIni As Date, ByVal dFechaFin As Date, ByVal nCanDiaMora As Integer, ByVal sAnslis As Variant, ByVal sAgencias As Variant, Optional sProductos As Variant = "") As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrHandler
    'ALPA 20090324*****************************************************************************
    'sSql = "exec stp_sel_ReporteRentabilidadCarteraXAnalista '" & Format(dFechaIni, "YYYYMMDD") & "','" & Format(dFechaFin, "YYYYMMDD") & "','" & nCanDiaMora & "','" & sAnslis & "','" & sAgencias & "','" & sProductos & "'"
    sSql = "exec stp_sel_ReporteRentabilidadCarteraXAnalista '" & Format(dFechaIni, "YYYYMMDD") & "','" & Format(dFechaFin, "YYYYMMDD") & "','" & nCanDiaMora & "','" & sAnslis & "','" & sAgencias & "','" & sProductos & "'"
    '******************************************************************************************

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ReporteRentabilidadCarteraXAnalista = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set ReporteRentabilidadCarteraXAnalista = Null
End Function
Public Function ReporteClientesConDeudaCanceladaNoVueltosAtender(ByVal dFechaIni As Date, ByVal dFechaFin As Date, ByVal nCanDiaMora As Integer, ByVal sAnslis As Variant, ByVal sAgencias As Variant) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrHandler
    
    sSql = "exec stp_sel_ReporteClientesConDeudaCanceladaNoVueltosAtender '" & Format(dFechaIni, "YYYYMMDD") & "','" & Format(dFechaFin, "YYYYMMDD") & "','" & nCanDiaMora & "','" & sAnslis & "','" & sAgencias & "'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ReporteClientesConDeudaCanceladaNoVueltosAtender = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set ReporteClientesConDeudaCanceladaNoVueltosAtender = Null
End Function
'***End********************************************************************************************

'*** PEAC 20080804
Public Function RecuperaDatosClientesHisNegativo(ByVal pMatAgencias As Variant) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
Dim sCadAge As String

    sCadAge = ""
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & ","
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)

sSql = "exec stp_sel_RecuperaDatosClientesHisNegativo '" & sCadAge & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion

Set RecuperaDatosClientesHisNegativo = oConecta.CargaRecordSet(sSql)

oConecta.CierraConexion
Set oConecta = Nothing

End Function

'***PEAC 20080805*
Public Function RecuperaClienteHisNegativo(ByVal pcCodCli As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "exec stp_sel_RecuperaClienteHisNegativo '" & pcCodCli & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set RecuperaClienteHisNegativo = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function
'****ALPA 20080805*********
Public Function RecuperaCreditosParaAdjudicados(Optional ByVal psPersCod As String = "", Optional psCtaCod As String = "") As ADODB.Recordset
Dim oParam As DCOMParametro
Dim nValor As Double
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCreditosParaAdjudicados
    Set oParam = New DCOMParametro
    nValor = oParam.RecuperaValorParametro(COMDConstantes.gColocEstJudicial)
    Set oParam = Nothing

    If psPersCod <> "" Then
''        sSql = "Select P.cCtaCod, C.nMontoCol, P.nSaldo, P.nTasaInteres, P.nPrdEstado, CC.nNroCalen, "
''        sSql = sSql & " cAnalista=(Select cPersNombre From Persona Where cPersCod = PP.cPersCod),"
''        sSql = sSql & " cTitular=(Select cPersNombre From Persona Where cPersCod = PP2.cPersCod), CC.nDiasAtraso "
''        sSql = sSql & " From Producto P Inner join Colocaciones C ON P.cCtaCod = C.cCtaCod "
''        sSql = sSql & "                 Inner Join ColocacCred CC ON P.cCtaCod = CC.cCtaCod "
''        sSql = sSql & "                 Left Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & COMDConstantes.gColRelPersAnalista
''        sSql = sSql & "                 Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod And PP2.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
''        sSql = sSql & " WHERE PP2.cPersCod = '" & psPersCod & "'"
''        sSql = sSql & " AND P.nPrdEstado in (" & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstRefMor & "," & COMDConstantes.gColocEstRefNorm & "," & gColocEstRefVenc & ")"
    sSql = "exec stp_sel_ReporteTransferirGarantiasParaAdjudicado " & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstRefMor & "," & COMDConstantes.gColocEstRefNorm & "," & gColocEstRefVenc & ",'" & psPersCod & "','" & psCtaCod & "'," & nValor & ",1,''"
    Else
        If psCtaCod <> "" Then
''            sSql = "Select P.cCtaCod, C.nMontoCol, P.nSaldo, P.nTasaInteres, P.nPrdEstado, CC.nNroCalen, "
''            sSql = sSql & " cAnalista=(Select cPersNombre From Persona Where cPersCod = PP.cPersCod),"
''            sSql = sSql & " cTitular=(Select cPersNombre From Persona Where cPersCod = PP2.cPersCod), CC.nDiasAtraso "
''            sSql = sSql & " From Producto P Inner join Colocaciones C ON P.cCtaCod = C.cCtaCod "
''            sSql = sSql & "                 Inner Join ColocacCred CC ON P.cCtaCod = CC.cCtaCod "
''            sSql = sSql & "                 Left Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & COMDConstantes.gColRelPersAnalista
''            sSql = sSql & "                 Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod And PP2.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
''            sSql = sSql & " WHERE P.cCtaCod = '" & psCtaCod & "'"
''            sSql = sSql & " AND P.nPrdEstado in (" & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstRefMor & "," & COMDConstantes.gColocEstRefNorm & "," & gColocEstRefVenc & ")"
    sSql = "exec stp_sel_ReporteTransferirGarantiasParaAdjudicado " & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstRefMor & "," & COMDConstantes.gColocEstRefNorm & "," & gColocEstRefVenc & ",'" & psPersCod & "','" & psCtaCod & "'," & nValor & ",2,''"
        Else
'            sSql = "Select P.cCtaCod, C.nMontoCol, P.nSaldo, P.nTasaInteres, P.nPrdEstado, CC.nNroCalen, "
'            sSql = sSql & " cAnalista=(Select cPersNombre From Persona Where cPersCod = PP.cPersCod),"
'            sSql = sSql & " cTitular=(Select cPersNombre From Persona Where cPersCod = PP2.cPersCod), CC.nDiasAtraso "
'            sSql = sSql & " From Producto P Inner join Colocaciones C ON P.cCtaCod = C.cCtaCod "
'            sSql = sSql & "                 Inner Join ColocacCred CC ON P.cCtaCod = CC.cCtaCod "
'            sSql = sSql & "                 Left Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & COMDConstantes.gColRelPersAnalista
'            sSql = sSql & "                 Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod And PP2.nPrdPersRelac = " & COMDConstantes.gColRelPersTitular
'            sSql = sSql & " WHERE CC.nDiasAtraso > " & nValor
'            sSql = sSql & " AND P.nPrdEstado in (" & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstRefMor & "," & COMDConstantes.gColocEstRefNorm & "," & gColocEstRefVenc & ")"
    sSql = "exec stp_sel_ReporteTransferirGarantiasParaAdjudicado " & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstRefMor & "," & COMDConstantes.gColocEstRefNorm & "," & gColocEstRefVenc & ",'" & psPersCod & "','" & psCtaCod & "'," & nValor & ",3,''"
        End If
    End If
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditosParaAdjudicados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaCreditosParaAdjudicados:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function
'*********************
'ALPA 20080926*************************************************************
Public Function ObtenerMontoTotalCalendarioXConcepto(ByVal psCtaCod As String, ByVal psConceptos As String) As ADODB.Recordset
Dim sSql As String
Dim C As COMConecta.DCOMConecta

    sSql = "exec stp_sel_ObtenerMontoTotalCalendarioXConcepto '" & psCtaCod & "','" & psConceptos & "'"

    
    Set C = New COMConecta.DCOMConecta
    C.AbreConexion
    Set ObtenerMontoTotalCalendarioXConcepto = C.CargaRecordSet(sSql)
    C.CierraConexion
    Set C = Nothing
    
End Function
'**************************************************************************
'ALPA 20080926*************************************************************
Public Function RecuperaTipoProducto() As ADODB.Recordset
Dim sSql As String
Dim C As COMConecta.DCOMConecta

    sSql = "exec stp_sel_RecuperaTipoProducto "
    
    Set C = New COMConecta.DCOMConecta
    C.AbreConexion
    Set RecuperaTipoProducto = C.CargaRecordSet(sSql)
    C.CierraConexion
    Set C = Nothing
    
End Function
'**************************************************************************

'MAVM 26112009 *********
Public Function RecuperaUbigeo(ByVal psAgeCod As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    sSql = "Select cUbiGeoDescripcion"
    sSql = sSql & " From Agencias A "
    sSql = sSql & " inner join UbicacionGeografica UG on A.cUbiGeoCod = UG.cUbiGeoCod "
    sSql = sSql & " Where cAgeCod = '" & psAgeCod & "'"
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set RecuperaUbigeo = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
'*********

'Constancia de NoAdeudo ******
'MAVM 20100412 ***
Public Sub RegistrarSolicitudConstNoAdeudo(ByVal sPersCod As String, ByVal sFSolicitud As String, ByVal sMotivo As String)
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_ins_SolicitudConstNoAdeudo '" & sPersCod & "', '" & sFSolicitud & "', '" & sMotivo & "'"
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub

Public Function ValidarConstNoAdeudoXPersona(ByVal sPersCod As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_sel_ValidarConstNoAdeudoXPersona '" & sPersCod & "'"
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set ValidarConstNoAdeudoXPersona = lrDatos
    Set lrDatos = Nothing
End Function

Public Function ListarConstNoAdeudo(ByVal sPersCod As String, ByVal sFI As String, ByVal sFF As String, ByRef psMensaje As String) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_ListarConstNoAdeudo '" & sPersCod & "', '" & Format(sFI, "yyyymmdd") & "', '" & Format(sFF, "yyyymmdd") & "'"
    
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    If lrDatos Is Nothing Then
        psMensaje = "No existen datos"
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        psMensaje = "No existen datos"
        Exit Function
    End If
    Set ListarConstNoAdeudo = lrDatos
    Set lrDatos = Nothing
End Function

Public Sub ModificarEstadoConstNoAdeudo(ByVal iConstNoAdeudoId As Integer, ByVal cPersCod As String, ByVal sFEntrega As String)
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_upd_EstadoConstNoAdeudo " & iConstNoAdeudoId & ", '" & cPersCod & "', '" & sFEntrega & "'"
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub

Public Function ValidarEntrega(ByVal iConstNoAdeudoId As Integer) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_sel_ValidarEntrega " & iConstNoAdeudoId & ""
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set ValidarEntrega = lrDatos
    Set lrDatos = Nothing
End Function

'*** PEAC 20100525 - MEMO 512-2010-GM-DI/CMACM - NO S PERMITE QUE UN MISMO ANALISTA APRUEBE UN CREDITO ASIGNADO A EL MISMO
'Public Function VerificaAprobacionDeAnalistaConSuCredito(ByVal psCtaCod As String, ByVal psUser As String) As Boolean
Public Function VerificaAprobacionDeAnalistaConSuCredito(ByVal psCtaCod As String, ByVal psCodUser As String) As Boolean 'MAVM 15112010
Dim lr As ADODB.Recordset
Dim lsSQL As String
Dim Co As COMConecta.DCOMConecta
Set Co = New COMConecta.DCOMConecta
    
    lsSQL = " exec stp_sel_VerificaAprobacionDeAnalistaConSuCredito '" & psCtaCod & "','" & psCodUser & "' "

    Co.AbreConexion
    Set lr = Co.CargaRecordSet(lsSQL)
        If lr!nNum > 0 Then
            VerificaAprobacionDeAnalistaConSuCredito = True
        Else
            VerificaAprobacionDeAnalistaConSuCredito = False
        End If
    Co.CierraConexion
    Set Co = Nothing
    lr.Close
End Function

'ALPA 20100603 BAS II**************************************************************
Public Function RecuperaProductosCrediticios() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta


    sSql = " Select C.cConsDescripcion, C.nConsValor from Constante C "
    sSql = sSql & " Inner Join ColocCredConsFiltro CF ON C.nConsCod = CF.nConsCod AND C.nConsValor = CF.nConsValor"
    sSql = sSql & " Where CF.nCodFiltro = " & gCredFiltroProdCab & " and C.nConsCod=3033 "
    sSql = sSql & " Order by C.nConsValor  "

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaProductosCrediticios = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

'Public Function RecuperaSubProductosCrediticios(Optional ByVal psCab As String = "", Optional psRHCargoCod As String = "") As ADODB.Recordset 'psRHCargoCod add by NAGL 20171121'Comento JOEP20190208 CP
Public Function RecuperaSubProductosCrediticios(Optional ByVal psCab As String = "", Optional psRHCargoCod As String = "", Optional pbOpeRefi As Boolean = False, Optional pbOpeAmpl As Boolean = False) As ADODB.Recordset 'JOEP20190118 CP
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    'sSql = " Select C.cConsDescripcion, C.nConsValor, CF.cAbrev from Constante C "
    'sSql = sSql & " Inner Join ColocCredConsFiltro CF ON C.nConsCod = CF.nConsCod AND C.nConsValor = CF.nConsValor"
    'sSql = sSql & " Where CF.nCodFiltro = " & gCredFiltroProd
    
    'sSql = sSql & " AND C.nConsCod =3033"
    
    'If psCab <> "" Then
        'sSql = sSql & " AND C.nConsValor like '" & Mid(psCab, 1, 1) & "%'"
    'End If
    'sSql = sSql & " Order by C.nConsValor " Comentado by NAGL 20171121
    
    'sSql = "Exec stp_sel_RecuperaSubProductosCrediticios '" & psCab & "', '" & psRHCargoCod & "'" 'NAGL 20171121'Comento JOEP20190208 CP
    sSql = "Exec stp_sel_RecuperaSubProductosCrediticios '" & psCab & "', '" & psRHCargoCod & "'," & IIf(pbOpeRefi = False, 0, 1) & "," & IIf(pbOpeAmpl = False, 0, 1) & "" 'JOEP20190118 CP
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaSubProductosCrediticios = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function
Public Function RecuperaTipoCreditos() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    'FRHU 20150512 SATI TIC1505080008
    'sSql = " Select C.cConsDescripcion, C.nConsValor from Constante C "
    'sSql = sSql & " Inner Join ColocCredConsFiltro CF ON C.nConsCod = CF.nConsCod AND C.nConsValor = CF.nConsValor"
    'sSql = sSql & " Where CF.nCodFiltro = " & gCredFiltroProdCab & " and C.nConsCod=3034 "
    'sSql = sSql & " Order by C.nConsValor  "
    sSql = "EXEC stp_sel_RecuperaTipoCreditos"
    'FIN FRHU 20150512

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaTipoCreditos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

'MIOL 20130612, SEGUN ERS076 ************************************************
Public Function RecuperaTipoCreditosxCod(ByVal pnConsValor As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    'FRHU 20150512 SATI TIC1505080008
    'sSql = " Select C.cConsDescripcion, C.nConsValor from Constante C "
    'sSql = sSql & " Inner Join ColocCredConsFiltro CF ON C.nConsCod = CF.nConsCod AND C.nConsValor = CF.nConsValor"
    'sSql = sSql & " Where CF.nCodFiltro = " & gCredFiltroProdCab & " and C.nConsCod=3034 "
    'If pnConsValor = 500 Then
        'sSql = sSql & " and C.nConsValor in (150,250,350,450,550)"
    'ElseIf pnConsValor = 600 Then
        'sSql = sSql & " and C.nConsValor in (350,450,550)"
    'ElseIf pnConsValor = 700 Then
        'sSql = sSql & " and C.nConsValor in (650,750)"
    'ElseIf pnConsValor = 800 Then
        'sSql = sSql & " and C.nConsValor in (850)"
    'End If
    'sSql = sSql & " Order by C.nConsValor  "
    sSql = "EXEC stp_sel_RecuperaTipoCreditosxCod " & pnConsValor
    'FIN FRHU 20150512
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaTipoCreditosxCod = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'END MIOL *******************************************************************

'Public Function RecuperaSubTipoCrediticios(Optional ByVal psCab As String = "", Optional psRHCargoCod As String = "") As ADODB.Recordset 'psRHCargoCod Add by NAGL 20171121'Comento JOEP20190208 CP
Public Function RecuperaSubTipoCrediticios(Optional ByVal psCab As String = "", Optional psRHCargoCod As String = "", Optional psCodProd As String = "") As ADODB.Recordset 'JOEP20190114 CP
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    'sSql = " Select C.cConsDescripcion, C.nConsValor, CF.cAbrev from Constante C "
    'sSql = sSql & " Inner Join ColocCredConsFiltro CF ON C.nConsCod = CF.nConsCod AND C.nConsValor = CF.nConsValor"
    'sSql = sSql & " Where CF.nCodFiltro = " & gCredFiltroProd
    'sSql = sSql & " AND C.nConsCod =3034"
    'If psCab <> "" Then
        'sSql = sSql & " AND C.nConsValor like '" & Mid(psCab, 1, 1) & "%'"
    'End If
    'sSql = sSql & " Order by C.nConsValor " Comentado by NAGL 20171121
    
    'sSql = "Exec stp_sel_RecuperaSubTipoCrediticios '" & psCab & "','" & psRHCargoCod & "'"  'NAGL 20171121'Comento JOEP20190208 CP
    sSql = "Exec stp_sel_RecuperaSubTipoCrediticios '" & psCab & "','" & psRHCargoCod & "','" & psCodProd & "'"  'JOEP20190114 CP
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaSubTipoCrediticios = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function ObtenerTipoCreditoxTipificacion(ByVal psPersCod As String) As Integer
Dim oConecta As COMConecta.DCOMConecta
Dim lrs As ADODB.Recordset
Dim lsSQL As String

On Error GoTo ErrorObtenerTipoCreditoxTipificacion

    Set lrs = New Recordset
    Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion

        lsSQL = " exec B2_stp_sel_GeneraTipoCredito '" & psPersCod & "'"

    Set lrs = oConecta.CargaRecordSet(lsSQL)

        If lrs.BOF And lrs.EOF Then
            ObtenerTipoCreditoxTipificacion = 0
        Else
            ObtenerTipoCreditoxTipificacion = lrs("cTipoCredCod")
        End If

        oConecta.CierraConexion
    Set lrs = Nothing
    Set oConecta = Nothing
    Exit Function

ErrorObtenerTipoCreditoxTipificacion:
    Err.Raise Err.Number, "Error en Obtener Tipo de Creditos : ObtenerTipoCredito ", Err.Description
End Function

'***************************************************************************

'MAVM 20100605
Public Function RecuperaSubTpoCrediticioCF() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    sSql = "Select cConsDescripcion, nConsValor"
    sSql = sSql & " From Constante"
    sSql = sSql & " Where nConsCod = 3034 and nConsValor = 181"
    sSql = sSql & " Union"
    sSql = sSql & " Select cConsDescripcion, nConsValor"
    sSql = sSql & " From Constante"
    sSql = sSql & " Where nConsCod = 3034 and nConsValor = 281"
    sSql = sSql & " Union"
    sSql = sSql & " Select cConsDescripcion, nConsValor"
    sSql = sSql & " From Constante"
    sSql = sSql & " Where nConsCod = 3034 and nConsValor = 381"
    sSql = sSql & " Union"
    sSql = sSql & " Select cConsDescripcion, nConsValor"
    sSql = sSql & " From Constante"
    sSql = sSql & " Where nConsCod = 3034 and nConsValor = 481"
    sSql = sSql & " Union"
    sSql = sSql & " Select cConsDescripcion, nConsValor"
    sSql = sSql & " From Constante"
    sSql = sSql & " Where nConsCod = 3034 and nConsValor = 581"
    sSql = sSql & " Order By nConsValor "
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaSubTpoCrediticioCF = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

'MAVM 20100826 ***
Public Function RecuperaCredMantPrepago(ByVal psCtaCod As String, ByVal psFecha As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCredMantPrepago
    sSql = "Select cCtaCod, cMovNro, nCalendDinamico, nMantNroCuota, nCuotasPend, nEstado From CredMantPrePago Where cCtacod = '" & psCtaCod & "' And SubString(cMovNro, 1, 8) = '" & Format(psFecha, "yyyymmdd") & "' And nEstado = 1"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCredMantPrepago = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCredMantPrepago:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function CuotaNormalPago(ByVal psCtaCod As String) As Double
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    sSql = "Select Top 1 nMontoCuota = (Isnull((select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in( 1000,1010)), 0) + IsNull((select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in( 1100,1107)), 0) + IsNull((select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1102), 0)) "
    sSql = sSql & " From ColocCalendario C Where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl = 1 And nNroCalen = (select nNroCalen from ColocacCred where cCtaCod = C.cCtaCod) And C.nColocCalendEstado = 0 Order by C.nCuota Asc"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    CuotaNormalPago = IIf(IsNull(R!nMontoCuota), 0, R!nMontoCuota)

End Function

Public Function CuotaNormalPagoSinGastos(ByVal psCtaCod As String) As Double
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    sSql = "Select Top 1 nMontoCuota = (Isnull((select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in( 1000,1010)), 0) + IsNull((select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in( 1100,1107)), 0) + IsNull((select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1102), 0) - IsNull((select sum(nMonto) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (Select nPrdConceptoCod from ProductoConcepto Where convert(varchar(6),nPrdConceptoCod) like '12%' )), 0)) "
    sSql = sSql & " From ColocCalendario C Where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl = 1 And nNroCalen = (select nNroCalen from ColocacCred where cCtaCod = C.cCtaCod) And C.nColocCalendEstado = 0 Order by C.nCuota Asc"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    If R.RecordCount <> 0 Then
        CuotaNormalPagoSinGastos = IIf(IsNull(R!nMontoCuota), 0, R!nMontoCuota)
    Else
        CuotaNormalPagoSinGastos = 0
    End If
End Function
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'************************************************************************************
'MADM 20101103 - 20100922 - PAGO VARIOS
Public Function ExistePagoVariosxDiaeInstitucion(ByVal psCodIns As String, ByVal pdFecSis As Date, Optional ByVal pbcab As Boolean) As Boolean
 Dim lrs As ADODB.Recordset
 Dim oConecta As COMConecta.DCOMConecta
 Dim lsSQL As String
 Set lrs = New ADODB.Recordset
 Set oConecta = New COMConecta.DCOMConecta
 ExistePagoVariosxDiaeInstitucion = False
 oConecta.AbreConexion

  If pbcab Then
    lsSQL = "Exec stp_sel_validaLogFechaClientesVarios '" & psCodIns & "','" & Format(pdFecSis, "yyyymmdd") & "'"
  Else
    ExistePagoVariosxDiaeInstitucion = False
    Exit Function
  End If
  
  Set lrs = oConecta.CargaRecordSet(lsSQL)

  If lrs.BOF And lrs.EOF Then
      ExistePagoVariosxDiaeInstitucion = False
  Else
      ExistePagoVariosxDiaeInstitucion = True
  End If

  oConecta.CierraConexion
  Set oConecta = Nothing
  lrs.Close
  Set lrs = Nothing
End Function

'MADM 20110307 - GENERA PAGO - 1
Public Sub InsertarDatosPagoVarios(ByVal prs As ADODB.Recordset, pPersCodCli As String, pdFecha As Date, pdMontoSol As Double, pdMontoDol As Double, cont As Integer, tr As Integer, pValcab As Boolean, nNumCab As Integer)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
        
    If pValcab Then
        sSql = "Exec stp_ins_LogClientesVarios  '" & prs("nTipoDoc") & "','" & prs("cNumDoc") & "','" & prs("nMoneda") & "','" & prs("cApeCliente") & "','" & prs("cNomCliente") & "','" & prs("cConcepto") & "','" & prs("cPeriodo") & "','" & prs("nImporteCuota") & "','" & pPersCodCli & "','" & Format(pdFecha, "yyyymmdd") & "', " & pdMontoSol & ", " & pdMontoDol & "," & cont & ", " & tr & " "
    Else
        sSql = "Exec stp_ins_LogClientesVariosDeudaLec " & nNumCab & ",'" & prs("nTipoDoc") & "','" & prs("cNumDoc") & "','" & prs("cApeCliente") & "','" & prs("cNomCliente") & "','" & prs("nMoneda") & "','" & prs("nImporteCuota") & "','" & prs("cConcepto") & "','" & prs("cPeriodo") & "'"
    End If
          
       Set oConecta = New COMConecta.DCOMConecta
       oConecta.AbreConexion
       oConecta.ConexionActiva.Execute sSql
       oConecta.CierraConexion
       Set oConecta = Nothing
End Sub
'MADM 20110307 - GENERA PAGO - 2
Public Function DevuelveMaxPagoInst(pPersCodCli As String, pdFecha As Date, tr As Integer) As Integer
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = "Exec stp_ins_maxDeudainstitucion  '" & pPersCodCli & "','" & Format(pdFecha, "yyyymmdd") & "'," & tr & ""
    Set lrDatos = oConecta.CargaRecordSet(sSql)
           
    If Not (lrDatos.BOF And lrDatos.EOF) Then
        DevuelveMaxPagoInst = IIf(IsNull(lrDatos!Codigo), 0, lrDatos!Codigo)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    lrDatos.Close
    Set lrDatos = Nothing
End Function

'MADM 20101110 - PROCESAR
'MADM 20110420
Public Function CargaDatosCabeceraPagoServicio(ByVal nImpToTS As String, nPersId As String, ByVal nNumReg As Long, ByVal fProceso As String, ByVal fCaducidad As String, ByVal tproceso As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim psFecha  As String
Dim psFecha1  As String

    If tproceso = 3 Then
        'excel
        psFecha = Mid(fProceso, 7, 4) & Mid(fProceso, 4, 2) & Mid(fProceso, 1, 2)
        psFecha1 = Mid(fCaducidad, 7, 4) & Mid(fCaducidad, 4, 2) & Mid(fCaducidad, 1, 2)
    Else
        'texto
        psFecha = fProceso
        psFecha1 = fCaducidad
    End If
    
    sSql = " exec stp_sel_LogPagoClientesVariosDeudaCab " & nImpToTS & ",'" & nPersId & "'," & nNumReg & " , '" & psFecha & "','" & psFecha1 & "'," & tproceso & ""
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CargaDatosCabeceraPagoServicio = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'Insertar Cabecera
Public Sub InsertarDatosPagoVariosCab(ByVal prs As ADODB.Recordset, ByVal tr As Integer)
    Dim sSql As String
    Dim psFecha As String
    Dim psFecha1 As String
    Dim oConecta As COMConecta.DCOMConecta
      
    If tr = 3 Then
        psFecha = Mid(prs("fProceso"), 7, 4) & Mid(prs("fProceso"), 4, 2) & Mid(prs("fProceso"), 1, 2)
        psFecha1 = Mid(prs("fCaducidad"), 7, 4) & Mid(prs("fCaducidad"), 4, 2) & Mid(prs("fCaducidad"), 1, 2)
    Else
        'tipo texto
        psFecha = prs("fProceso")
        psFecha1 = prs("fCaducidad")
    End If
    
    sSql = "Exec stp_ins_LogClientesVariosDeudaCab  '" & prs("cCodIns") & "','" & prs("nNumReg") & "','" & prs("nImpToTS") & "','" & prs("nImpToTD") & "','" & psFecha & "','" & psFecha1 & "'," & tr & ""
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
End Sub
'MADM 20110602 - 20110307 - MADM 20110420
Public Function InsertarDatosPagoVariosDet(ByVal prs As ADODB.Recordset, Optional ptr As Integer = 1) As Boolean
    Dim sSql As String
    Dim bTran  As Boolean
    Dim psFecha As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorInsertarDatosPagoVariosDet
    InsertarDatosPagoVariosDet = True
    bTran = False
    
    If ptr = 3 Then
        psFecha = Mid(prs("cPeriodo"), 7, 4) & Mid(prs("cPeriodo"), 4, 2) & Mid(prs("cPeriodo"), 1, 2)
    Else
        'texto
        psFecha = prs("cPeriodo")
    End If
    'cImprime
    sSql = "Exec stp_ins_LogClientesVariosDeudaLec  " & prs("cNumId") & ",'" & prs("cCodCli") & "','" & prs("cTpoDoc") & "','" & prs("cNumDoc") & "','" & prs("cApeCli") & "','" & prs("cNomCli") & "','" & prs("cMoneda") & "','" & prs("nImporte") & "','" & prs("cConcepto") & "','" & prs("cCodServicio") & "','" & psFecha & "','" & prs("cImprime") & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.BeginTrans
    bTran = True
    Call oConecta.ConexionActiva.Execute(sSql)
    
    oConecta.ConexionActiva.CommitTrans
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorInsertarDatosPagoVariosDet:
    If bTran Then
        InsertarDatosPagoVariosDet = False
        Call oConecta.ConexionActiva.RollbackTrans
        Set oConecta = Nothing
    End If
    Err.Raise Err.Number, "Error En Lectura de Archivo de Pago de Servicios", Err.Description
End Function
'Procesar Grilla - Cabecera
Public Function DevolverDatosCabeceraPagoTotal(ByVal cCodDeuCab As Long) As ADODB.Recordset
Dim R As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
     
     sSql = "Exec stp_sel_LogPagoClientesVariosDeudaCabTot  " & cCodDeuCab & ""
 
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Set DevolverDatosCabeceraPagoTotal = R
    Set R = Nothing
    Exit Function
End Function
'Procesar Grilla - Detalle
Public Function DevolverDatosDetPagoTotal(ByVal id As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogPagoClientesVariosDeudaDetTot " & id & "  "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverDatosDetPagoTotal = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'MADM 20111125 - 20110721 - 20110307 - 20101115 - Pago
Public Sub InsertarDatosPagoVariosServiciosDeuda(ByVal pnDeuNro As Long, ByVal pnImporteComision As Currency, ByVal pnImportePago As Currency, ByVal pdFechaPag As Date, ByVal pnMovNro As Long, ByVal pnMovNro2 As Long, ByVal psCuenta As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
        
        sSql = "Exec stp_ins_ClientesPagoClientesVariosPago " & pnDeuNro & "," & pnImporteComision & "," & pnImportePago & ",'" & Format(pdFechaPag, "YYYY/MM/DD hh:mm:ss") & "'," & pnMovNro & "," & pnMovNro2 & ", '" & psCuenta & "'"
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        oConecta.ConexionActiva.Execute sSql
        oConecta.CierraConexion
        Set oConecta = Nothing
End Sub
'20111206
Public Function DevolverDatosConceptosPagados(ByVal pnNumDeuSist As String) As ADODB.Recordset
Dim R As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim nCreditos As String
    
    nCreditos = ""

    If pnNumDeuSist <> "" Then
            If Len(pnNumDeuSist) > 1 Then
                nCreditos = Replace(Replace(pnNumDeuSist, "'", ""), " ", "")
            Else
                nCreditos = Replace(Replace(pnNumDeuSist, "'", ""), " ", "")
            End If
    End If
    
     sSql = "Exec stp_sel_CargaDatosConceptosPagados  '" & nCreditos & "'"
 
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Set DevolverDatosConceptosPagados = R
    Set R = Nothing
    Exit Function
End Function
'--
'MADM 20101109 - LISTAR reporte de registrados - y procesados
Public Function ReportesListaPagoVarios(ByVal ptr As Integer, ByVal pcInst As String, ByVal pdFecha As Date, Optional ByVal pcCab As Integer = 0) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrHandler

    If pcCab > 0 Then
        sSql = "exec stp_sel_CargaDatosPersPagoServicios " & ptr & "," & pcCab & ""
    Else
        sSql = "exec stp_Sel_Lista_LogPagoVarios " & ptr & ",'" & pcInst & "','" & Format(pdFecha, "yyyymmdd") & "'"
    End If
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ReportesListaPagoVarios = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set ReportesListaPagoVarios = Null
End Function

'MADM 20101203
Public Function ReportesListaPagoVariosPago(ByVal pcInst As String, ByVal pdFecha As Date, ByVal ind As Integer) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrHandler

   sSql = "exec stp_sel_LogPagoDeuda '" & pcInst & "','" & Format(pdFecha, "yyyymmdd") & "'," & ind & ""
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ReportesListaPagoVariosPago = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set ReportesListaPagoVariosPago = Null
End Function

'MADM 20110305
'EXCEL
'MADM 20110311 - PROCESAR
Public Function GetValConvenioInstitucionyAgencia(ByVal pcPersCod As String, ByVal pcCodAge As Integer) As Boolean
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim lr As ADODB.Recordset

    sSql = " exec stp_sel_CargaConvenioClientesVariosxInstyAgencia '" & pcPersCod & "'," & pcCodAge & ""
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set lr = oConecta.CargaRecordSet(sSql)
        If Not (lr.BOF And lr.EOF) Then
            GetValConvenioInstitucionyAgencia = True
        Else
            GetValConvenioInstitucionyAgencia = False
        End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    lr.Close
End Function
'15032011 - registro convenio - Tipos
Public Function GetValConvenioInstitucionxAgenciaTipo(ByVal pcPersCod As String) As Boolean
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim lr As ADODB.Recordset

    sSql = " exec stp_sel_CargaConvenioClientesVariosxInstyAgenciaTipoTotal '" & pcPersCod & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set lr = oConecta.CargaRecordSet(sSql)
        If Not (lr.BOF And lr.EOF) Then
            GetValConvenioInstitucionxAgenciaTipo = True
        Else
            GetValConvenioInstitucionxAgenciaTipo = False
        End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    lr.Close
End Function
'fecha venc convenio
Public Function GetValConvenioInstitucionxFechaVenc(ByVal pcPersCod As String) As Boolean
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim lr As ADODB.Recordset

    sSql = " exec stp_sel_CargaConvenioClientesVariosxInstyFecVenc '" & pcPersCod & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set lr = oConecta.CargaRecordSet(sSql)
        If Not (lr.BOF And lr.EOF) Then
            GetValConvenioInstitucionxFechaVenc = True
        Else
            GetValConvenioInstitucionxFechaVenc = False
        End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    lr.Close
End Function
Public Function GetValConvenioInstitucionxAgencia(ByVal pcPersCod As String, ByVal pcCodAge As Integer) As Boolean
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim lr As ADODB.Recordset

    sSql = " exec stp_sel_CargaConvenioClientesVariosxInstyAgencia '" & pcPersCod & "', " & pcCodAge & " "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set lr = oConecta.CargaRecordSet(sSql)
        If Not (lr.BOF And lr.EOF) Then
            GetValConvenioInstitucionxAgencia = True
        Else
            GetValConvenioInstitucionxAgencia = False
        End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    lr.Close
End Function
'stp_sel_CargaConvenioClientesVariosxInstyAgencia - ok
Public Function GetDevuelveConvenioInstitucionyAgenciayTipo(ByVal pcPersCod As String, ByVal pcCodAge As Integer, ByVal pnTipo As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    sSql = " exec stp_sel_CargaConvenioClientesVariosxInstyAgenciayTipo '" & pcPersCod & "'," & pcCodAge & ", " & pnTipo & ""
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set GetDevuelveConvenioInstitucionyAgenciayTipo = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
End Function
'ok
Public Function GetDevuelveAgenciaxTipoConvenio(ByVal filtro As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    sSql = "exec stp_sel_CargaAgenciaxTipoConvenio " & filtro & ""
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set GetDevuelveAgenciaxTipoConvenio = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
End Function
'ok
Public Function GetDevuelveConvenioAgenciayTipo(ByVal pcCodAge As Integer, ByVal filtro As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    sSql = "exec stp_sel_CargaConvenioClientesVariosxAgenciaxTipo " & pcCodAge & ", " & filtro & ""
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set GetDevuelveConvenioAgenciayTipo = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
End Function
'MADM 20110812ok
Public Function GetInstitucionesConvenioPago(ByVal pcCodAge As Integer, ByVal pdFecSis As Date) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String

   sSql = "exec stp_sel_LogInstPagoClientesVarios " & pcCodAge & ",'" & Format(pdFecSis, "yyyymmdd") & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set GetInstitucionesConvenioPago = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
End Function
Public Function GetInstitucionesPrevioPago(ByVal pcCodAge As Integer) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String

   sSql = "exec stp_sel_LogInstRegPago " & pcCodAge & " "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set GetInstitucionesPrevioPago = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
End Function
'MADM 20110602 - ByVal pnBusqueda As Integer
Public Sub InsertarDatosPagoConvenioInstyAgencia(ByVal pcPersCod As String, ByVal pnCodAge As Integer, ByVal pfRegistro As Date, ByVal pfVigencia As Date, ByVal pnTipoConv As Integer, ByVal pcObsConv As String, ByVal nImpComisionSol As Double, ByVal sCuenta As String, ByVal pnBusqueda As Integer)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
        
    sSql = "Exec stp_ins_LogPagoClientesConvenioVarios  '" & pcPersCod & "','" & pnCodAge & "','" & Format(pfRegistro, "yyyymmdd") & "','" & Format(pfVigencia, "yyyymmdd") & "','" & pnTipoConv & "', '" & pcObsConv & "'," & nImpComisionSol & ",'" & sCuenta & "', " & pnBusqueda
       
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
End Sub
'MADM 20110602 - ByVal pnBusqueda As Integer
Public Sub ActualizarDatosPagoConvenioInstyAgencia(ByVal pcPersCod As String, ByVal pcCodAge As Integer, ByVal pfVigencia As Date, pnTipoConv As Integer, pcObsConv As String, nImpComisionSol As Double, sCuenta As String, ByVal pnBusqueda As Integer)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
        
    sSql = "Exec stp_upd_LogPagoClientesConvenioVarios  '" & pcPersCod & "','" & pcCodAge & "','" & Format(pfVigencia, "yyyymmdd") & "'," & pnTipoConv & ", '" & pcObsConv & "'," & nImpComisionSol & ",'" & sCuenta & "', " & pnBusqueda
       
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
End Sub

Public Function GetDevuelveReporteConvenioAgencia(ByVal pcCodAge As Integer, ByVal pfiltro As Integer, ByVal ptodos As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    sSql = "exec stp_sel_ReporteconvenioClientesVariosxAgencia " & pcCodAge & "," & pfiltro & "," & ptodos & ""
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set GetDevuelveReporteConvenioAgencia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
End Function
'MADM 20110425
Public Sub ActualizarColocacEstadoDiasAtraso(ByVal psCtaCod As String, ByVal pnEstado As ColocEstado, Optional pnNewPeriodoGracia As Integer)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorActualizarColocacEstadoDiasAtraso
        sSql = "update ColocacEstado set nPeriodoGracia=" & pnNewPeriodoGracia & " Where cCtaCod = '" & psCtaCod & "' AND nPrdEstado =  " & pnEstado
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.CargaRecordSet (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorActualizarColocacEstadoDiasAtraso:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Sub
'MADM 20110712 ---------------------------------------------------------
Public Sub ActualizaDatosPagoVariosCabAnt(ByVal pdFecha As Date, ByVal pcPersCodCli As String)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
           
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = "Exec stp_val_LogPagoClientesVariosDeudaCabAnt  '" & pcPersCodCli & "','" & Format(pdFecha, "yyyymmdd") & "'"
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
End Sub
'END MADM
'END MADM

'************************************************************************************
Public Sub ActualizarColocacEstado(ByVal psCtaCod As String, ByVal pnEstado As ColocEstado, Optional pnPeriodoFechaFija As Integer)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorActualizarColocacEstado
        sSql = "update ColocacEstado set nPeriodoFechaFija=" & pnPeriodoFechaFija & " Where cCtaCod = '" & psCtaCod & "' AND nPrdEstado =  " & pnEstado
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.CargaRecordSet (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorActualizarColocacEstado:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Sub

'MAVM 20110514 ***
Public Function RecuperaNivApr(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaNivApr
    sSql = "Select cCtaCod, cCodCargo, vComentario, iEstado, cMovNro From ColocCredPreNivApr Where cCtaCod = '" & psCtaCod & "'" & " And iEstado in(0, 2)"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaNivApr = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaNivApr:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function RecuperaLlaveProceso(ByVal psConstante As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaLlaveProceso
    sSql = "Select nConsValor From Constante Where nConsCod = '" & psConstante & "' And nConsValor <> '" & psConstante & "'" & " And bEstado = 1"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaLlaveProceso = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaLlaveProceso:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function RecuperaMontoDesemb(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaMontoDesemb
    sSql = "Select CD.nMonto, cTpoCredCod From ColocCalenddet CD Inner Join Colocaciones C on CD.cCtaCod = C.cCtaCod Where CD.cCtaCod = '" & psCtaCod & "'" & " And CD.nColocCalendApl = 0 and nPrdConceptoCod = 1000 and nCuota = 1 And CD.nNroCalen = (Select MAX(nNroCalen) From ColocCalendario Where cCtaCod = CD.cCtaCod AND nColocCalendApl = 0)"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaMontoDesemb = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaMontoDesemb:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'***

Public Function RecuperaBancosCajasClasificacion() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaBancosCajasClasificacion
    sSql = "exec stp_sel_BancosCajasClasificacion"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaBancosCajasClasificacion = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaBancosCajasClasificacion:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function RecuperaCarteraCF(ByVal nTipoCambio As String, ByVal pdFecha As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCarteraCF
    sSql = "exec stp_sel_CarteraCF " & nTipoCambio & ",'" & Format(pdFecha, "YYYY/MM/DD") & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCarteraCF = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCarteraCF:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function RecuperaCarteraAutoliquidable(ByVal nTipoCambio As String, ByVal pnUIT As Currency, ByVal pdFecha As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCarteraAutoliquidable
    sSql = "exec stp_sel_ObtenerCarteraTotalAutoliquidable " & nTipoCambio & "," & pnUIT & ",'" & Format(pdFecha, "YYYY/MM/DD") & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCarteraAutoliquidable = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCarteraAutoliquidable:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function RecuperaCarteraCreditoPorTramos(ByVal nTipoCambio As String, ByVal pdFecha As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCarteraCreditoPorTramos
    'JUEZ 20131211 **************************************
'    sSql = "        select left(C.cTpoCredCod,1) cTpoCred,sum(C.nSaldoCap*case SUBSTRING(C.cCtaCod,9,1) when '1' then 1 else " & nTipoCambio & " end) K,"
'    sSql = sSql & "   sum((case when c.cCalGen <> '0' then C.nProvision else 0 end)*case SUBSTRING(C.cCtaCod,9,1) when '1' then 1 else " & nTipoCambio & " end) P,"
'    sSql = sSql & "   sum((isnull(C.nProvisionRCC,0))*case SUBSTRING(C.cCtaCod,9,1) when '1' then 1 else " & nTipoCambio & " end) PRCC, "
'    sSql = sSql & "   sum(isnull(R.nSaldo,0)) ID,"
'    sSql = sSql & "   cOrden='1'"
'    sSql = sSql & " from DBConsolidada..ColocCalifProvtotal C"
'    sSql = sSql & "     left join DBConsolidada..RCDvc" & Format(pdFecha, "YYYYMM") & "02 R on C.cCtaCod=R.cCtaCod and cCtaCnt like '14_8%'"
'    sSql = sSql & " where C.nPrdEstado IN (2020, 2021, 2022, 2030, 2031, 2032 , 2201,2205, 2101,2104, 2106, 2107)"
'    sSql = sSql & "     and DateDiff(d,C.dFecha,'" & Format(pdFecha, "YYYYMMDD") & "')=0"
'    sSql = sSql & "     and C.nDiasAtraso>90"
'    sSql = sSql & "     and c.cCalGen <> '0'"
'    sSql = sSql & "     and nProvision>=0.2*C.nSaldoCap"
'    sSql = sSql & "     and right(C.cTpoCredCod,2)<>'81'"
'    sSql = sSql & "     and left(C.cTpoCredCod,1)<>'8'"
'    sSql = sSql & " group by left(C.cTpoCredCod,1)"
'    sSql = sSql & " Union"
'    sSql = sSql & " select left(C.cTpoCredCod,1) cTpoCred,sum(C.nSaldoCap*case SUBSTRING(C.cCtaCod,9,1) when '1' then 1 else " & nTipoCambio & " end) K,"
'    sSql = sSql & "   sum((case when c.cCalGen <> '0' then C.nProvision else 0 end)*case SUBSTRING(C.cCtaCod,9,1) when '1' then 1 else " & nTipoCambio & " end) P,"
'    sSql = sSql & "   sum((isnull(C.nProvisionRCC,0))*case SUBSTRING(C.cCtaCod,9,1) when '1' then 1 else " & nTipoCambio & " end) PRCC, "
'    sSql = sSql & "   sum(isnull(R.nSaldo,0)) ID,"
'    sSql = sSql & "   cOrden='2'"
'    sSql = sSql & " from DBConsolidada..ColocCalifProvtotal C"
'    sSql = sSql & "     left join DBConsolidada..RCDvc" & Format(pdFecha, "YYYYMM") & "02 R on C.cCtaCod=R.cCtaCod and cCtaCnt like '14_8%'"
'    sSql = sSql & " where C.nPrdEstado IN (2020, 2021, 2022, 2030, 2031, 2032 , 2201,2205, 2101,2104, 2106, 2107)"
'    sSql = sSql & "     and C.nDiasAtraso<=90"
'    sSql = sSql & "     and DateDiff(d,C.dFecha,'" & Format(pdFecha, "YYYYMMDD") & "')=0"
'    sSql = sSql & "     and right(C.cTpoCredCod,2)<>'81'"
'    sSql = sSql & "     and left(C.cTpoCredCod,1)<>'8'"
'    sSql = sSql & " group by left(C.cTpoCredCod,1)"
'    sSql = sSql & " Union"
'    sSql = sSql & " select left(C.cTpoCredCod,1) cTpoCred,sum(C.nSaldoCap*case SUBSTRING(C.cCtaCod,9,1) when '1' then 1 else " & nTipoCambio & " end) K,"
'    sSql = sSql & "    sum((case when c.cCalGen <> '0' then C.nProvision else 0 end)*case SUBSTRING(C.cCtaCod,9,1) when '1' then 1 else " & nTipoCambio & " end) P,"
'    sSql = sSql & "   sum((isnull(C.nProvisionRCC,0))*case SUBSTRING(C.cCtaCod,9,1) when '1' then 1 else " & nTipoCambio & " end) PRCC, "
'    sSql = sSql & "    sum(isnull(R.nSaldo,0)) ID,"
'    sSql = sSql & "    cOrden='3'"
'    sSql = sSql & " from DBConsolidada..ColocCalifProvtotal C"
'    sSql = sSql & "     left join DBConsolidada..RCDvc" & Format(pdFecha, "YYYYMM") & "02 R on C.cCtaCod=R.cCtaCod and cCtaCnt like '14_8%'"
'    sSql = sSql & " where C.nPrdEstado IN (2020, 2021, 2022, 2030, 2031, 2032 , 2201,2205, 2101,2104, 2106, 2107)"
'    sSql = sSql & "     and C.nDiasAtraso>90"
'    sSql = sSql & "     and DateDiff(d,C.dFecha,'" & Format(pdFecha, "YYYYMMDD") & "')=0"
'    sSql = sSql & "     and c.cCalGen <> '0'"
'    sSql = sSql & "     and nProvision<0.2*C.nSaldoCap"
'    sSql = sSql & "     and right(C.cTpoCredCod,2)<>'81'"
'    sSql = sSql & "     and left(C.cTpoCredCod,1)<>'8'"
'    sSql = sSql & " group by left(C.cTpoCredCod,1)"
'    sSql = sSql & " Union"
'    sSql = sSql & " select left(C.cTpoCredCod,1) cTpoCred,sum(C.nSaldoCap*case SUBSTRING(C.cCtaCod,9,1) when '1' then 1 else " & nTipoCambio & " end) K,"
'    sSql = sSql & "   sum((case when c.cCalGen <> '0' then C.nProvision else 0 end)*case SUBSTRING(C.cCtaCod,9,1) when '1' then 1 else " & nTipoCambio & " end) P,"
'    sSql = sSql & "   sum((isnull(C.nProvisionRCC,0))*case SUBSTRING(C.cCtaCod,9,1) when '1' then 1 else " & nTipoCambio & " end) PRCC, "
'    sSql = sSql & "   sum(isnull(R.nSaldo,0)) ID,"
'    sSql = sSql & "   cOrden='4'"
'    sSql = sSql & " from DBConsolidada..ColocCalifProvtotal C"
'    sSql = sSql & "     left join DBConsolidada..RCDvc" & Format(pdFecha, "YYYYMM") & "02 R on C.cCtaCod=R.cCtaCod and cCtaCnt like '14_8%'"
'    sSql = sSql & " where C.nPrdEstado IN (2020, 2021, 2022, 2030, 2031, 2032 , 2201,2205, 2101,2104, 2106, 2107)"
'    sSql = sSql & "     and right(C.cTpoCredCod,2)<>'81'"
'    sSql = sSql & "     and DateDiff(d,C.dFecha,'" & Format(pdFecha, "YYYYMMDD") & "')=0"
'    sSql = sSql & "     and left(C.cTpoCredCod,1)<>'8'"
'    sSql = sSql & " group by left(C.cTpoCredCod,1)"
'    sSql = sSql & " Union"
'    sSql = sSql & " select cTpoCred='8',"
'    sSql = sSql & "   sum(nSaldoCap*case cMoneda when '1' then 1 else " & nTipoCambio & " end) K,"
'    sSql = sSql & "   sum(nProvision*case cMoneda when '1' then 1 else " & nTipoCambio & " end) P,"
'    sSql = sSql & "   sum(PRCC) PRCC, "
'    sSql = sSql & "   sum(nSaldo) ID,"
'    sSql = sSql & "   '1' cOrden"
'    sSql = sSql & " From"
'    sSql = sSql & " (   select"
'    sSql = sSql & "         C.nDiasAtraso,C.cCtaCod,"
'    sSql = sSql & "         case when C.cCalGen<>'0' then isnull(C.nProvision,0) else 0 end nProvision,"
'    sSql = sSql & "         sum((isnull(C.nProvisionRCC,0))*case SUBSTRING(C.cCtaCod,9,1) when '1' then 1 else " & nTipoCambio & " end) PRCC,"
'    sSql = sSql & "         isnull(C.nSaldoCap,0) nSaldoCap,"
'    sSql = sSql & "         SUBSTRING(C.cCtaCod,9,1)cMoneda,sum(G.nRealizacion) nRealizacion,"
'    sSql = sSql & "         ISNULL(R.nSaldo,0) nSaldo,"
'    sSql = sSql & "         C.cCalGen"
'    sSql = sSql & "     from DBConsolidada..ColocCalifProvtotal C"
'    sSql = sSql & "         inner join ColocGarantia CG on C.cCtaCod=CG.cCtaCod"
'    sSql = sSql & "         inner join Garantias G on CG.cNumGarant=G.cNumGarant"
'    sSql = sSql & "         left join DBConsolidada..RCDvc" & Format(pdFecha, "YYYYMM") & "02 R on C.cCtaCod=R.cCtaCod and cCtaCnt like '14_8%'"
'    sSql = sSql & "     where left(C.cTpoCredCod,1) in ('8')"
'    sSql = sSql & "         and DateDiff(d,C.dFecha,'" & Format(pdFecha, "YYYYMMDD") & "')=0"
'    sSql = sSql & "         and C.nPrdEstado IN (2020, 2021, 2022, 2030, 2031, 2032 , 2201,2205, 2101,2104, 2106, 2107)"
'    sSql = sSql & "     group by C.cCtaCod,C.nProvision,C.nSaldoCap,SUBSTRING(C.cCtaCod,9,1),C.nDiasAtraso,C.cCalGen,R.nSaldo"
'    sSql = sSql & " ) x"
'    sSql = sSql & " Where nDiasAtraso < 90 And nRealizacion >= nSaldoCap"
'    sSql = sSql & " Union"
'    sSql = sSql & " select cTpoCred='8',"
'    sSql = sSql & "   sum(nSaldoCap*case cMoneda when '1' then 1 else " & nTipoCambio & " end) K,"
'    sSql = sSql & "   sum(nProvision*case cMoneda when '1' then 1 else " & nTipoCambio & " end) P,"
'    sSql = sSql & "   sum(PRCC) PRCC, "
'    sSql = sSql & "   sum(nSaldo) ID,"
'    sSql = sSql & "   '2' cOrden"
'    sSql = sSql & " From"
'    sSql = sSql & " (   select"
'    sSql = sSql & "         C.nDiasAtraso,C.cCtaCod,"
'    sSql = sSql & "         case when C.cCalGen<>'0' then isnull(C.nProvision,0) else 0 end nProvision,"
'    sSql = sSql & "         sum((isnull(C.nProvisionRCC,0))*case SUBSTRING(C.cCtaCod,9,1) when '1' then 1 else " & nTipoCambio & " end) PRCC,"
'    sSql = sSql & "         isnull(C.nSaldoCap,0) nSaldoCap,"
'    sSql = sSql & "         SUBSTRING(C.cCtaCod,9,1)cMoneda,sum(G.nRealizacion) nRealizacion,"
'    sSql = sSql & "         ISNULL(R.nSaldo,0) nSaldo,"
'    sSql = sSql & "         C.cCalGen"
'    sSql = sSql & "     from DBConsolidada..ColocCalifProvtotal C"
'    sSql = sSql & "         inner join ColocGarantia CG on C.cCtaCod=CG.cCtaCod"
'    sSql = sSql & "         inner join Garantias G on CG.cNumGarant=G.cNumGarant"
'    sSql = sSql & "         left join DBConsolidada..RCDvc" & Format(pdFecha, "YYYYMM") & "02 R on C.cCtaCod=R.cCtaCod and cCtaCnt like '14_8%'"
'    sSql = sSql & "         where left(C.cTpoCredCod,1) in ('8')"
'    sSql = sSql & "             and DateDiff(d,C.dFecha,'" & Format(pdFecha, "YYYYMMDD") & "')=0"
'    sSql = sSql & "             and C.nPrdEstado IN (2020, 2021, 2022, 2030, 2031, 2032 , 2201,2205, 2101,2104, 2106, 2107)"
'    sSql = sSql & "     group by C.cCtaCod,C.nProvision,C.nSaldoCap,SUBSTRING(C.cCtaCod,9,1),C.nDiasAtraso,C.cCalGen,R.nSaldo"
'    sSql = sSql & " ) x"
'    sSql = sSql & " Where nDiasAtraso >= 90"
'    sSql = sSql & " Union"
'    sSql = sSql & " select cTpoCred='8',"
'    sSql = sSql & "   sum(nSaldoCap*case cMoneda when '1' then 1 else " & nTipoCambio & " end) K,"
'    sSql = sSql & "   sum(nProvision*case cMoneda when '1' then 1 else " & nTipoCambio & " end) P,"
'    sSql = sSql & "   sum(PRCC) PRCC, "
'    sSql = sSql & "   sum(nSaldo) ID,"
'    sSql = sSql & "   '3' cOrden"
'    sSql = sSql & " From"
'    sSql = sSql & " (select"
'    sSql = sSql & "    C.nDiasAtraso,C.cCtaCod,"
'    sSql = sSql & "    case when C.cCalGen<>'0' then isnull(C.nProvision,0) else 0 end nProvision,"
'    sSql = sSql & "    sum((isnull(C.nProvisionRCC,0))*case SUBSTRING(C.cCtaCod,9,1) when '1' then 1 else " & nTipoCambio & " end) PRCC,"
'    sSql = sSql & "    isnull(C.nSaldoCap,0) nSaldoCap,"
'    sSql = sSql & "    SUBSTRING(C.cCtaCod,9,1)cMoneda,sum(G.nRealizacion) nRealizacion,"
'    sSql = sSql & "    ISNULL(R.nSaldo,0) nSaldo,"
'    sSql = sSql & "    C.cCalGen"
'    sSql = sSql & " from DBConsolidada..ColocCalifProvtotal C"
'    sSql = sSql & "     inner join ColocGarantia CG on C.cCtaCod=CG.cCtaCod"
'    sSql = sSql & "     inner join Garantias G on CG.cNumGarant=G.cNumGarant"
'    sSql = sSql & "     left join DBConsolidada..RCDvc" & Format(pdFecha, "YYYYMM") & "02 R on C.cCtaCod=R.cCtaCod and cCtaCnt like '14_8%'"
'    sSql = sSql & " where left(C.cTpoCredCod,1) in ('8')"
'    sSql = sSql & "     and DateDiff(d,C.dFecha,'" & Format(pdFecha, "YYYYMMDD") & "')=0"
'    sSql = sSql & "     and C.nPrdEstado IN (2020,2021,2022,2030,2031,2032,2201,2205,2101,2104,2106,2107)"
'    sSql = sSql & " group by C.cCtaCod,C.nProvision,C.nSaldoCap,SUBSTRING(C.cCtaCod,9,1),C.nDiasAtraso,C.cCalGen,R.nSaldo"
'    sSql = sSql & " ) x"
'    sSql = sSql & " Where nDiasAtraso < 90 And nRealizacion < nSaldoCap"
'    sSql = sSql & " order by cTpoCred,cOrden"
    sSql = "exec stp_sel_RecuperaCarteraCreditoPorTramos '" & Format(pdFecha, "YYYYMMDD") & "'," & nTipoCambio
    'END JUEZ *******************************************
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCarteraCreditoPorTramos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCarteraCreditoPorTramos:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function RecuperaSaldosInteresDiferidos(ByVal nTipoCambio As String, ByVal pdFecha As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaSaldosInteresDiferidos
    sSql = "exec stp_sel_SaldoInteresesDiferidos '" & Format(pdFecha, "YYYY/MM/DD") & "'," & nTipoCambio
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaSaldosInteresDiferidos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaSaldosInteresDiferidos:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function RecuperaProvisionCartera(ByVal nTipoCambio As String, ByVal pdFecha As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaProvisionCartera
    sSql = "exec stp_sel_ObtenerProvision '" & Format(pdFecha, "YYYY/MM/DD") & "'," & nTipoCambio
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaProvisionCartera = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaProvisionCartera:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'ALPA 20111209
'ALPA 20110819***********************************************************************
Public Function RecuperaDatosPagoCuotaLeasing(ByVal psCtaCod As String, ByVal pdFecha As Date, ByVal pnOpcion As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaDatosPagoCuotaLeasing
    sSql = "exec stp_sel_ObtenerDatosPagoCuotaLeasing '" & psCtaCod & "','" & Format(pdFecha, "YYYY/MM/DD") & "'," & pnOpcion
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosPagoCuotaLeasing = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaDatosPagoCuotaLeasing:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'************************************************************************************

Public Function ObtenerProximaCuota(ByVal pnNroCalendario As Integer, ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtenerProximaCuota
    sSql = "exec stp_sel_ObtenerProximaCuota " & pnNroCalendario & ",'" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerProximaCuota = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtenerProximaCuota:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ObtenerProcesoLeasing(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtenerProcesoLeasing
    sSql = "exec B2_stp_sel_Desembolso_Log '" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerProcesoLeasing = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtenerProcesoLeasing:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'*** BRGO 20111103 *********************************************************
Public Function ValidaPersonaCOFIDE(ByVal psPersCod As String, Optional ByVal pnTipoPersona As Integer = -1) As String
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset

sSql = "SELECT cPersCod FROM PersonaCOFIDE WHERE cPersCod='" & psPersCod & "' "
If pnTipoPersona <> -1 Then
    sSql = sSql & " and nTipoPers = " & pnTipoPersona
End If

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing

If rs.EOF Then
    ValidaPersonaCOFIDE = "La persona no está Afiliado al Sistema Infogas"
End If
End Function


Public Function RecuperaRelacionesInfogas(ByVal psCtaCod As String, Optional ByVal psFiltro As String = "", Optional ByVal psFiltro2 As String = "") As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = " Select PP.cPersCod, ISNULL(P.cPersNombre,'') Nombre, "
    sSql = sSql & " (UPPER(RTRIM(T.cConsDescripcion)) + SPACE(75 - len((UPPER(RTRIM(T.cConsDescripcion))))) + CONVERT(Varchar(2),PP.nPrdPersRelac)) cRelacion,"
    sSql = sSql & " PP.nMontoAbono , PP.cCtaCodAbono"
    sSql = sSql & " FROM Persona P"
    sSql = sSql & " RIGHT JOIN ProductoPersonaInfogas PP ON P.cPersCod = PP.cPersCod"
    sSql = sSql & " INNER JOIN Constante T ON PP.nPrdPersRelac = T.nConsValor"
    sSql = sSql & " WHERE PP.cCtaCod = '" & psCtaCod & "' AND T.nConsCod = 3036"
    If psFiltro <> "" Then
        sSql = sSql & " AND PP.nPrdPersRelac NOT IN (" & psFiltro & ")"
    End If
    If psFiltro2 <> "" Then
        sSql = sSql & " AND PP.nPrdPersRelac IN (" & psFiltro2 & ")"
    End If
    sSql = sSql & " ORDER BY PP.nPrdPersRelac"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    If Not rs.EOF Then
        Set RecuperaRelacionesInfogas = rs
    End If
    Set rs = Nothing
End Function
'*** END BRGO *********************************************************
'EJVG20120612 ******************
Public Function RecuperaRecaudosParaAbono(ByVal pdFechaCarga As Date) As ADODB.Recordset
    Dim oCon As New COMConecta.DCOMConecta
    Dim sSql As String
    sSql = "stp_sel_ListaRecaudosxFechaCarga '" & Format(pdFechaCarga, "yyyymmdd hh:mm:ss") & "'"
    oCon.AbreConexion
    Set RecuperaRecaudosParaAbono = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function SeSubioArchivoRecaudo(ByVal pdFechaCarga As Date) As Boolean
    Dim oCon As New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim sSql As String
    Set rs = New ADODB.Recordset
    sSql = "stp_sel_SeSubioArchivoRecaudo '" & Format(pdFechaCarga, "yyyymmdd") & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    SeSubioArchivoRecaudo = rs!bSubio
    oCon.CierraConexion
    Set oCon = Nothing
    Set rs = Nothing
End Function
Public Function RecuperaCuotasAPagarCreditoEcotaxi(ByVal psAgeCod As String, ByVal pdFechaPago As Date) As ADODB.Recordset
    Dim oCon As New COMConecta.DCOMConecta
    Dim sSql As String
    sSql = "stp_sel_ListaCuotasPagarEcotaxi '" & psAgeCod & "','" & Format(pdFechaPago, "yyyymmdd") & "'"
    oCon.AbreConexion
    Set RecuperaCuotasAPagarCreditoEcotaxi = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'END EJVG **********************
'EJVG20111219 ******************
Public Function ObtenerVecesCreditoyAhorroPersona(ByVal psPersCod As String) As Integer
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    On Error GoTo ErrorObtenerVecesCreditoyAhorroPersona
    sSql = " exec stp_sel_CuentaCreditosAhorroPersona '" & psPersCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    Set rs = New ADODB.Recordset
    
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    ObtenerVecesCreditoyAhorroPersona = rs!nVeces
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtenerVecesCreditoyAhorroPersona:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'JACA 20111219********************************************************
Public Function getMetasAgenciasDet(psAgeCod As String, psAnio As String, psMes As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorMsg
    sSql = "exec stp_sel_MetasAgenciasDet '" & psAgeCod & "','" & psAnio & "','" & psMes & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set getMetasAgenciasDet = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function getSaldoCarteraAgencia(sAgeCod As String, dFecha As Date, nTpoCamb As Double) As Currency
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As Recordset
    On Error GoTo ErrorMsg
    sSql = "exec stp_sel_SaldoCarteraAgencia '" & Format(dFecha, "yyyymmdd") & "','" & sAgeCod & "'," & nTpoCamb & ""
    Set rs = New Recordset
    Set oConecta = New COMConecta.DCOMConecta
   
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    
    If Not (rs.EOF And rs.BOF) Then
        getSaldoCarteraAgencia = IIf(IsNull(rs!nSaldo), 0#, rs!nSaldo)
    Else
        getSaldoCarteraAgencia = 0#
    End If
    
    Set oConecta = Nothing
    Exit Function

ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Sub updEstadoMetasAgencias(psAgeCod As String, psAnio As String, psMes As String)
    Dim sSql As String
    On Error GoTo ErrorMsg
    
   sSql = "exec stp_upd_EstadoMetasAgencia '" & psAgeCod & "','" & psAnio & "','" & psMes & "'"
   loConecta.CargaRecordSet (sSql)
   
   Exit Sub

ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Sub

Public Sub insertarMetasAgencias(psAgeCod As String, psAnio As String, psMes As String, pnSaldoAnt As Currency, pnSaldoMeta As Currency, psMovNro As String)
    Dim sSql As String
    On Error GoTo ErrorMsg
    
   sSql = "exec stp_ins_MetasAgencia '" & psAgeCod & "','" & psAnio & "','" & psMes & "'," & pnSaldoAnt & "," & pnSaldoMeta & ",'" & psMovNro & "'"
   loConecta.CargaRecordSet (sSql)
   
   Exit Sub

ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Sub
Public Sub insertarMetasAgenciasDet(psMovNro As String, pnTpoCred As Integer, pnMonto As Currency)
    Dim sSql As String
    On Error GoTo ErrorMsg
    
   sSql = "exec stp_ins_MetasAgenciaDet '" & psMovNro & "'," & pnTpoCred & "," & pnMonto & ""
   loConecta.CargaRecordSet (sSql)
   
   Exit Sub

ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Sub
Public Sub BeginTrans()
    loConecta.BeginTrans
End Sub

Public Sub CommitTrans()
    loConecta.CommitTrans
End Sub
Public Sub RollbackTrans()
    loConecta.RollbackTrans
End Sub
Private Sub Class_Terminate()
    loConecta.CierraConexion
    Set loConecta = Nothing
End Sub
'JACA END*************************************************************

'MADM 20120125
Public Function getDatosCabeceraBCP(ByVal nImpToTal As Currency, ByVal nNumReg As Integer, ByVal fProceso As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCobrosBcoCredCabecera " & nImpToTal & "," & nNumReg & " , '" & fProceso & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set getDatosCabeceraBCP = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function getDatosCabeceraBCPCta(ByVal psCtaCod As String, ByVal pdFecha As Date, ByVal psPersNombre As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCobrosBcoBCPCtaCod '" & psCtaCod & "' ,'" & Format(pdFecha, "YYYYMMDD") & "' ,'" & psPersNombre & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set getDatosCabeceraBCPCta = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Sub InsertarDatosCabeceraBCP(ByVal prs As ADODB.Recordset, Optional ByVal nValor As Integer = 0)

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorInsertarDatosCabeceraBCP
   
   sSql = "Exec stp_ins_LogCobrosBcoBCP '" & prs("c1") & "','" & prs("c2") & "','" & prs("c3") & "','" & prs("c4") & "','" & prs("c5") & "', '" & prs("c6") & "'," & CInt(prs("c7")) & "," & prs("c8") & ", '" & prs("c9") & "', '" & prs("c10") & "', " & nValor & ""
       Set oConecta = New COMConecta.DCOMConecta
       oConecta.AbreConexion
       oConecta.ConexionActiva.Execute sSql
       oConecta.CierraConexion
       Set oConecta = Nothing
    Exit Sub
ErrorInsertarDatosCabeceraBCP:
    Err.Raise Err.Number, "Insertar Datos Cabecera", Err.Description
End Sub

Public Sub InsertarDatosDetalleBCP(ByVal id As Double, ByVal prs As ADODB.Recordset)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorInsertarDatosDetalleBCP
    
    sSql = "Exec stp_ins_LogCobrosBcoBCPDet " & id & ",'" & prs("cCodCtaF") & "'," & CInt(prs("Cuota")) & ",'" & prs("Titular") & "','" & prs("c1") & "','" & prs("c2") & "','" & prs("c3") & "', '" & prs("c4") & "', '"
    sSql = sSql & prs("c5") & "','" & prs("c6") & "', '" & prs("c7") & "','" & prs("c8") & "'," & prs("c9") & " ,"
    sSql = sSql & prs("c10") & "," & prs("c11") & ",'" & prs("c12") & "','" & prs("c13") & "', '"
    sSql = sSql & prs("c14") & "','" & prs("c15") & "','" & prs("c16") & "' ,'" & prs("c17") & "','" & prs("c18") & "','" & prs("c19") & "' ,'"
    sSql = sSql & prs("c20") & "','" & prs("c21") & "' "
       
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorInsertarDatosDetalleBCP:
    Err.Raise Err.Number, "Error Insertar Datos BCP Detalle", Err.Description
End Sub

Public Function getDatosDetalleBcoBCPxId(ByVal id_cab As Double, ByVal cCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCobrosBcoBCPDetalleID " & id_cab & " ,'" & cCtaCod & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set getDatosDetalleBcoBCPxId = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Sub ActualizarDetallexEstBCP(ByVal id As Double, ByVal cCodCtaF As String, ByVal nNroCta As String, ByVal dFechaCobro As String, ByVal nTipoPago As Integer, Optional ByVal bValorAct As Boolean = True)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim nInt As Integer
    On Error GoTo ErrorActualizarDetallexEstBCP
    
    sSql = "Exec stp_upd_LogCobrosBcoNacDet " & id & ",'" & cCodCtaF & "','" & nNroCta & "','" & dFechaCobro & "' ," & nInt & "  "
      
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorActualizarDetallexEstBCP:
    Err.Raise Err.Number, "Actualizar Datos Detalle", Err.Description
End Sub

Public Function getCabeceraBcoBCPxID(ByVal id As Double) As ADODB.Recordset
Dim R As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sAnalistas As String
Dim sSqlaux As String

    On Error GoTo ERRORgetCabeceraBcoBCPxID
    
    sSql = " exec stp_sel_ListaLogCobrosBcoCredCabecera " & id & " "
   
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Set getCabeceraBcoBCPxID = R
    Set R = Nothing
    Exit Function
ERRORgetCabeceraBcoBCPxID:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

Public Function getDetalleBcoBCPDetallexId(ByVal id As Double) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCobrosBcoBCPDetallexId " & id & "  "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set getDetalleBcoBCPDetallexId = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function getDetalleBcoBCPDetallexIdNoPagados(ByVal id As Double) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCobrosBcoBCPDetallexIdNoPagados " & id & "  "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set getDetalleBcoBCPDetallexIdNoPagados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function getDevolverDatosReporteBCP(ByVal dfpago As Date, Optional est As Integer = 0) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCorresponsaliaBcoBCPReporte  '" & Format(dfpago, "YYYYMMDD") & "'," & est & " "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set getDevolverDatosReporteBCP = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function getDevolverDatosDatosReporteBCP(ByVal id As Double) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCorresponsaliaBcoBCPDetalle " & id & "  "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set getDevolverDatosDatosReporteBCP = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

'madm 20100225
Public Sub InsertarDatosBCPSobrantes(ByVal cCtaCod As String, ByVal nCuota As Integer, ByVal nMontoDif As Double, ByVal dfpago As Date, ByVal cNomCliente As String, ByVal nMovNro As Long)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorInsertarDatosBCPSobrantes
   
   sSql = "Exec stp_ins_MovCobrosBcoBCPSobrante '" & cCtaCod & "'," & nCuota & "," & nMontoDif & ", '" & Format(dfpago, "YYYYMMDD") & "','" & cNomCliente & "'," & nMovNro & ""
       Set oConecta = New COMConecta.DCOMConecta
       oConecta.AbreConexion
       oConecta.ConexionActiva.Execute sSql
       oConecta.CierraConexion
       Set oConecta = Nothing
    Exit Sub
ErrorInsertarDatosBCPSobrantes:
    Err.Raise Err.Number, "Cargar Datos Generales", Err.Description
End Sub

Public Sub ActualizarDetallexEstMovBCP(ByVal id As Double, ByVal nMovNro As String, ByVal nMovNroR As String, ByVal nMovNroC As String)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorActualizarDetallexEstMovBCP
 
   sSql = "Exec stp_upd_LogCobrosBcoBCPDet_Mov " & id & ",'" & nMovNro & "','" & nMovNroR & "','" & nMovNroC & "'"
       Set oConecta = New COMConecta.DCOMConecta
       oConecta.AbreConexion
       oConecta.ConexionActiva.Execute sSql
       oConecta.CierraConexion
       Set oConecta = Nothing
    Exit Sub
ErrorActualizarDetallexEstMovBCP:
    Err.Raise Err.Number, "Actualizar Datos Detalle", Err.Description
End Sub

Public Function ReportesListaRecuperacionesBCP(ByVal pdFecha As Date, Optional ByVal psNumCab As Double = 0) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrHandler

    sSql = "exec stp_sel_lista_recuperaciones_Cred '" & Format(pdFecha, "yyyymmdd") & "', " & psNumCab & ""
      
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ReportesListaRecuperacionesBCP = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set ReportesListaRecuperacionesBCP = Null
End Function

Public Function getDevolverDatosDatosReporteBCPNoProcesados(ByVal id As Double) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCobrosBcoBCPDetalleNoPagados " & id & " "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set getDevolverDatosDatosReporteBCPNoProcesados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Sub EliminaCabceraBCP(ByVal id As Double)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorEliminaCabceraBCP
 
   sSql = "Exec elimina_cabeceraBCP " & id & ""
       Set oConecta = New COMConecta.DCOMConecta
       oConecta.AbreConexion
       oConecta.ConexionActiva.Execute sSql
       oConecta.CierraConexion
       Set oConecta = Nothing
    Exit Sub
ErrorEliminaCabceraBCP:
    Err.Raise Err.Number, "Elimina Datos Detalle", Err.Description
End Sub

'END MADM

'ALPA 20120216**************************************************************************
Public Function ObtenerCicloEconomico(pdFecha As Date, pnTipoCambio As Currency) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorMsgSE
        Set oConecta = New COMConecta.DCOMConecta
        
        oConecta.AbreConexion
        sSql = "stp_sel_CicloEconomicoRCDVcFecha02 '" & Format(pdFecha, "YYYY/MM/DD") & "'"
        oConecta.CargaRecordSet (sSql)
        
        sSql = "exec stp_sel_CicloEconomico '" & Format(pdFecha, "YYYY/MM/DD") & "'," & pnTipoCambio
'        Set oConecta = New COMConecta.DCOMConecta
        
        Set ObtenerCicloEconomico = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
    Exit Function

ErrorMsgSE:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ObtenerRiesgoConcentracionIndividual(pnTipoCambio As Currency) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorMsgRCI
        sSql = "exec stp_sel_ColocCalifPro100PrincipalesDeudas " & pnTipoCambio
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set ObtenerRiesgoConcentracionIndividual = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
    Exit Function

ErrorMsgRCI:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ObtenerRiesgoConcentracionSectorial(psFechaRep As Date, psFechaInicio As Date, pnTipoCambio As Currency) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorMsgRCS
        sSql = "exec B2_stp_sel_RiesgoConcentracionSectorial '" & Format(psFechaRep, "YYYY/MM/DD") & "', '" & Format("01/" & Month(psFechaInicio) & "/" & Year(psFechaInicio), "YYYY/MM/DD") & "'," & pnTipoCambio
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set ObtenerRiesgoConcentracionSectorial = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
    Exit Function

ErrorMsgRCS:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ObtenerRiesgoConcentracionRegional(pnTipoCambio As Currency) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorMsgRCR
        sSql = "exec stp_sel_RiesgoConcentracionRegional " & pnTipoCambio
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set ObtenerRiesgoConcentracionRegional = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
    Exit Function

ErrorMsgRCR:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ObtenerReporteRiesgoCambiario(psAnio As String, psMes As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorMsgRCCC
        sSql = "exec stp_sel_ReporteRiesgoCambiario '" & psAnio & "','" & psMes & "'"
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set ObtenerReporteRiesgoCambiario = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
    Exit Function

ErrorMsgRCCC:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'***************************************************************************************

'JACA 20120112*****************************************************************
Public Function getFechaCompromisoPago(psCtaCod As String, pnNroCuota As Integer) As Recordset
       
    Dim sSql As String
    On Error GoTo ErrorMsg
    
   sSql = "exec stp_sel_FechaCompromisoPago '" & psCtaCod & "'," & pnNroCuota & ""
   Set getFechaCompromisoPago = loConecta.CargaRecordSet(sSql)
   
   Exit Function

ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function
Public Sub updEstadoFecCompromisoPago(psCtaCod As String, nCuota As Integer)
    Dim sSql As String
    On Error GoTo ErrorMsg
    
   sSql = "exec stp_upd_EstadoFecCompromisoPago '" & psCtaCod & "'," & nCuota & ""
   loConecta.CargaRecordSet (sSql)
   
   Exit Sub

ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Sub

Public Sub insertarFecCompromisoPago(psCtaCod As String, pnCuota As Integer, pdFecPago As Date, psMovNro As String)
    Dim sSql As String
    On Error GoTo ErrorMsg
    
   sSql = "exec stp_ins_FecCompromisoPago '" & psCtaCod & "'," & pnCuota & ",'" & Format(pdFecPago, "yyyymmdd") & "','" & psMovNro & "'"
   loConecta.CargaRecordSet (sSql)
   
   Exit Sub

ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Sub
Public Function getPermisoFechaCompPago(psCargoCod As String, psPersCod As String) As Boolean
   Dim sSql As String
   Dim rs As New Recordset
    On Error GoTo ErrorMsg
    
   sSql = "exec stp_sel_PermisoFechaCompPago '" & psCargoCod & "','" & psPersCod & "'"
   Set rs = loConecta.CargaRecordSet(sSql)
   
   If Not (rs.EOF And rs.BOF) Then
        getPermisoFechaCompPago = True
   Else
        getPermisoFechaCompPago = False
   End If
   
   Exit Function

ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function GetCargoActualizarCompPago(ByVal psCtaCod As String, ByVal psPersCod As String) As Boolean
   Dim sSql As String
   Dim rs As New Recordset
    On Error GoTo ErrorMsg
    
   sSql = "exec stp_sel_GetCargoActualizarCompPago '" & psCtaCod & "','" & psPersCod & "'"
   Set rs = loConecta.CargaRecordSet(sSql)
   
   If Not (rs.EOF And rs.BOF) Then
        GetCargoActualizarCompPago = True
   Else
        GetCargoActualizarCompPago = False
   End If
   
   Exit Function

ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function
'JACA END**********************************************************************
'ALPA 20120328*****************************************************************
Public Function ObtnerUIT(pnPeriodo As Integer) As Currency
    Dim sSql As String
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    On Error GoTo ErrorMsg
    
    sSql = "exec stp_sel_UIT '" & pnPeriodo & "'"
    Set oRs = loConecta.CargaRecordSet(sSql)
        ObtnerUIT = 0#
        If Not (oRs.EOF Or oRs.BOF) Then
            ObtnerUIT = oRs!nUIT
        End If
   Exit Function

ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'******************************************************************************

'WIOR 20120424********************************************************************
Public Function CargaCuentasPersona(ByVal psPersCod As String, ByVal psEstado As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As Recordset
    Dim sSql As String
    Set oConec = New COMConecta.DCOMConecta
    Set rs = New Recordset
    On Error GoTo ErrorMsg
    
    If oConec.AbreConexion() Then
        sSql = "exec stp_sel_VerCuentasPersona '" & psPersCod & "','" & psEstado & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set CargaCuentasPersona = rs
    Exit Function
ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function HayRelacionCreditoGarantia(ByVal psCtaCod As String, ByVal psNumGarant As String, ByVal psEstado As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As Recordset
    Dim sSql As String
    Set oConec = New COMConecta.DCOMConecta
    Set rs = New Recordset
    On Error GoTo ErrorMsg
    
    If oConec.AbreConexion() Then
        sSql = "exec stp_sel_HayRelacionCreditoGarantia '" & psCtaCod & "','" & psNumGarant & "','" & psEstado & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set HayRelacionCreditoGarantia = rs
    Exit Function
ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Sub InsertarRelacionCredGarantPoliza(ByVal psCtaCod As String, ByVal psNumPoliza As String, ByVal pdFecha As Date)
    Dim sSql As String
    On Error GoTo ErrorMsg
    
    sSql = "exec stp_ins_RelacionCredPoliza '" & psCtaCod & "','" & psNumPoliza & "','" & Format(pdFecha, "yyyymmdd") & "'"
   loConecta.CargaRecordSet (sSql)
   
   Exit Sub
ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Sub

Public Sub ActualizarRelacionCredGarantPoliza(ByVal psCtaCod As String, ByVal psNumPoliza As String, ByVal psEstado As Integer)
    Dim sSql As String
    On Error GoTo ErrorMsg
    
   sSql = "exec stp_upd_RelacionCredPoliza '" & psCtaCod & "','" & psNumPoliza & "'," & psEstado
   loConecta.CargaRecordSet (sSql)
   

   Exit Sub
ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Sub

Public Function RecuperaRelacionCreditoPoliza(ByVal psCtaCod As String, ByVal psNumPoliza As String, ByVal psEstado As Integer) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As Recordset
    Dim sSql As String
    Set oConec = New COMConecta.DCOMConecta
    Set rs = New Recordset
    On Error GoTo ErrorMsg
    
    If oConec.AbreConexion() Then
        sSql = "exec stp_sel_RelacionColocPol '" & psCtaCod & "','" & psNumPoliza & "'," & psEstado
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set RecuperaRelacionCreditoPoliza = rs
    Exit Function
ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function TienePoliza(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As Recordset
    Dim sSql As String
    Set oConec = New COMConecta.DCOMConecta
    Set rs = New Recordset
    On Error GoTo ErrorMsg
    
    If oConec.AbreConexion() Then
        sSql = "exec stp_sel_TienePoliza '" & psCtaCod & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set TienePoliza = rs
    Exit Function
ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Sub CancelarPoliza(ByVal psNumPoliza As String)
    Dim sSql As String
    On Error GoTo ErrorMsg
    
   sSql = "exec stp_upd_CancelarPoliza '" & psNumPoliza & "'"
   loConecta.CargaRecordSet (sSql)
   
   Exit Sub
ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Sub


Public Function VerificarRefinanciados(ByVal psCtaCod As String, ByVal pnEstado As Integer) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As Recordset
    Dim sSql As String
    Set oConec = New COMConecta.DCOMConecta
    Set rs = New Recordset
    On Error GoTo ErrorMsg
    
    If oConec.AbreConexion() Then
        sSql = "exec stp_sel_VerificaRefinanciado '" & psCtaCod & "'," & pnEstado
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set VerificarRefinanciados = rs
    Exit Function
ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function VerificarAmpliados(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As Recordset
    Dim sSql As String
    Set oConec = New COMConecta.DCOMConecta
    Set rs = New Recordset
    On Error GoTo ErrorMsg
    
    If oConec.AbreConexion() Then
        sSql = "exec stp_sel_VerificaAmpliado '" & psCtaCod & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set VerificarAmpliados = rs
    Exit Function
ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ObtenerRelacionCredGarantPol(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As Recordset
    Dim sSql As String
    Set oConec = New COMConecta.DCOMConecta
    Set rs = New Recordset
    On Error GoTo ErrorMsg
    
    If oConec.AbreConexion() Then
        sSql = "exec stp_sel_RelacionColGarPolDetalle '" & psCtaCod & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set ObtenerRelacionCredGarantPol = rs
    Exit Function
ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'WIOR FIN *************************************************************************
'WIOR 20120517*********************************************************************
Public Function ObtenerMicroseguro(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As Recordset
    Dim sSql As String
    Set oConec = New COMConecta.DCOMConecta
    Set rs = New Recordset
    On Error GoTo ErrorMsg
    
    If oConec.AbreConexion() Then
        'CTI6-20210503-ERS032-2019 -(Optimizar Sugerencia)
        'sSql = "select cCtaCod,nTipo,ISNULL(cNumCert,'') cNumCert,dEstado,nEstado from ColocacMicroseguro WHERE cCtaCod='" & psCtaCod & "'"
        sSql = "Execute spt_sel_ObtenerMicroseguro '" & psCtaCod & "'"
        'Fin CTI6-20210503
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set ObtenerMicroseguro = rs
    Exit Function
ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ObtenerMultiriesgo(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As Recordset
    Dim sSql As String
    Set oConec = New COMConecta.DCOMConecta
    Set rs = New Recordset
    On Error GoTo ErrorMsg
    
    If oConec.AbreConexion() Then
        sSql = "stp_sel_ObtenerMultiriesgo '" & psCtaCod & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set ObtenerMultiriesgo = rs
    Exit Function
ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'WIOR FIN *************************************************************************
'RECO20150525 ERS023-2015**********************************************************
Public Function ObtenerMultiriesgoMYPE(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As Recordset
    Dim sSql As String
    Set oConec = New COMConecta.DCOMConecta
    Set rs = New Recordset
    On Error GoTo ErrorMsg
    
    If oConec.AbreConexion() Then
        sSql = "stp_sel_ObtieneValorMultiriesgoMYPE '" & psCtaCod & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set ObtenerMultiriesgoMYPE = rs
    Exit Function
ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'RECO FIN *************************************************************************

'MAVM 20120528 Correccion de Asiento CTB***
Public Sub ActualizarMontoColocacEstadoXRefin(ByVal psCtaCod As String, ByVal pnMonto As Double)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorActualizarMontoColocacEstadoXRefin
    sSql = "Update ColocacEstado set nMonto= " & pnMonto & " Where cCtaCod = '" & psCtaCod & "' AND nPrdEstado in (2000,2001)"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.CargaRecordSet (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorActualizarMontoColocacEstadoXRefin:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Sub
'***

'MAVM 20120614 ***
Public Function bEsFeriadoDomingo(ByVal dFecha As Date, ByVal psAgeCod As String) As Boolean
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

On Error GoTo ErrorbEsFeriadoDomingo
    sSql = "Select dbo.fnc_EsFeriadoDomingo ('" & Format(dFecha, "YYYYMMDD") & "', '" & psAgeCod & "')"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    If Not R.BOF And Not R.EOF Then
        bEsFeriadoDomingo = True
    Else
        bEsFeriadoDomingo = False
    End If
    R.Close
    Set R = Nothing
    Exit Function

ErrorbEsFeriadoDomingo:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'***
'EJVG20120620 ****************
Public Function RecuperaDatosCredEcoTaxixPenalidad(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRecuperaDatosCredEcoTaxixPenalidad
    sSql = "Exec stp_sel_ObtieneDatosEcoTaxixPenalidad '" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosCredEcoTaxixPenalidad = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorRecuperaDatosCredEcoTaxixPenalidad:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function DamePenalidadxMotivo(ByVal psCtaCod As String, ByVal pnMotivoPenalidad As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorDamePenalidadxMotivo
    sSql = "Exec stp_sel_PenalidadxCtaCodyMotivoPenalidad '" & psCtaCod & "'," & pnMotivoPenalidad
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set DamePenalidadxMotivo = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorDamePenalidadxMotivo:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'END EJVG ********************
'WIOR 20120727**************************************************
Public Sub OpeInformeRiesgo(ByVal psCtaCod As String, ByVal pnTipo As Integer, _
                            Optional ByVal pcMovIngExp As String = "", _
                            Optional ByVal pcMovSalExp As String = "", _
                            Optional ByVal pnNivel As Integer = 0, _
                            Optional ByVal pcGlosa As String = "", _
                            Optional ByVal pnEstado As Integer = 0, _
                            Optional ByVal pnRiesgo As Integer, _
                            Optional ByVal pnInformeID As Long = 0)
'WIOR 20141013
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorOpeInformeRiesgo
    
    If pnTipo = 1 Then
        'sSql = "Insert into InformeRiesgo(cCtaCod,nEstado,nRiesgo)values ('" & psCtaCod & "'," & pnEstado & "," & pnRiesgo & ")"'WIOR 20141013 COMENTO
        sSql = "EXEC stp_ins_InformeRiesgo '" & psCtaCod & "'," & pnEstado & "," & pnRiesgo 'WIOR 20141013
    ElseIf pnTipo = 2 Then
        'WIOR 20141013 COMENTO TODO ESTA PARTE
        'sSql = "Update InformeRiesgo Set "
        
        'If pcMovIngExp <> "" Then
        '    sSql = sSql & " cMovIngExp='" & pcMovIngExp & "',"
        'End If
        
        'If pcMovSalExp <> "" Then
        '    sSql = sSql & " cMovSalExp='" & pcMovSalExp & "',"
        'End If
        
        'If pnNivel <> 0 Then
        '    sSql = sSql & " nNivel=" & pnNivel & ","
        'End If
        
        'If pcGlosa <> "" Then
        '    sSql = sSql & " cGlosa='" & pcGlosa & "',"
        'End If
        
        'If pnEstado <> 0 Then
        '    sSql = sSql & " nEstado=" & pnEstado & ","
        'End If
        
        'If Trim(Right(sSql, 4)) = "Set" Then
        '    sSql = ""
        'Else
        '    sSql = Mid(sSql, 1, Len(sSql) - 1)
        '    sSql = sSql & " WHERE cCtaCod='" & psCtaCod & "' AND nEstado<>3" 'WIOR 20130411 AGREGO AND nEstado<>3
        'End If
        sSql = "EXEC stp_upd_InformeRiesgo '" & psCtaCod & "','" & pcMovIngExp & "','" & pcMovSalExp & "'," & pnNivel & ",'" & pcGlosa & "'," & pnEstado & "," & pnInformeID 'WIOR 20141013
    ElseIf pnTipo = 3 Then
        'sSql = "delete from InformeRiesgo where cCtaCod='" & psCtaCod & "' AND nEstado <> 3" 'WIOR 20130411 SE AGREGO AND nEstado <> 3'WIOR 20141013 COMENTO
        sSql = "EXEC stp_del_InformeRiesgo '" & psCtaCod & "'" 'WIOR 20141013
    End If
        
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.CargaRecordSet (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorOpeInformeRiesgo:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Sub

Public Function ObtenerInformeRiesgo(ByVal psCtaCod As String, Optional ByVal pnTipo As Integer = 0, Optional ByVal pbAproba As Boolean = False) As ADODB.Recordset 'WIOR 20130411 AGREGO pbAproba
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorObtenerInformeRiesgo
    If pnTipo = 0 Then
    'WIOR 20141013 COMENTO
    'sSql = "select IR.cCtaCod,IR.cMovIngExp,IR.cMovSalExp,IR.nNivel,IR.cGlosa,IR.nEstado,IR.nRiesgo " & _
    '       " from InformeRiesgo IR where IR.cCtaCod='" & psCtaCod & "'"
    ''WIOR 20130411 *****************************
    'If pbAproba Then
    '    sSql = sSql & " AND CONVERT(FLOAT,LEFT(ISNULL(IR.cMovIngExp,''),14))=(SELECT max(CONVERT(FLOAT,LEFT(ISNULL(X.cMovIngExp,''),14))) FROM InformeRiesgo X WHERE X.CCTACOD=IR.CCTACOD)"
    'Else
    '    sSql = sSql & " AND IR.nEstado<>3"
    'End If
    ''WIOR FIN **********************************
        sSql = "EXEC stp_sel_InformeRiesgo '" & psCtaCod & "'," & IIf(pbAproba, 1, 0) 'WIOR 20141013
    ElseIf pnTipo = 1 Then
    'WIOR 20130411 comento lineas abajo
    'sSql = "select IR.cCtaCod,isnull(dbo.FechaHoraMov(IR.cMovIngExp),'') Ingreso,isnull(dbo.FechaHoraMov(IR.cMovSalExp),'') Salida, " & _
    '       " CASE WHEN ISNULL(IR.nNivel,0)=0 then 'Sin Nivel' else CONS1.cConsDescripcion end Nivel, isnull(IR.cGlosa,'') Glosa " & _
    '       " from InformeRiesgo IR " & _
    '       " LEFT JOIN Constante CONs1 ON CONS1.nConsValor=IR.nNivel and CONs1.nConsCod =9999 " & _
    '       " WHERE cCtaCod='" & psCtaCod & "'"
    'WIOR FIN ******************
        sSql = "EXEC stp_sel_InformeRiesgoDatos '" & psCtaCod & "'" 'WIOR 20130411
    End If
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set ObtenerInformeRiesgo = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtenerInformeRiesgo:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function CreditoInformeRiesgoPend() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorCreditoInformeRiesgoPend

    sSql = "exec stp_sel_CreditosInformeRiesgoPend"
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set CreditoInformeRiesgoPend = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
ErrorCreditoInformeRiesgoPend:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'WIOR FIN ******************************************************

'JUEZ 20120908 *************************************************************************
Public Function RecuperaColocacCredEval(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_RecuperaColocacCredEval '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set RecuperaColocacCredEval = rs
    'End If
    Set rs = Nothing
End Function

Public Function RecuperaColocacCredEvalAprobacion(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_RecuperaColocacCredEvalAprobacion '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set RecuperaColocacCredEvalAprobacion = rs
    'End If
    Set rs = Nothing
End Function

Public Function RecuperaDatosIndicadCredEval(ByVal pnCodicion As Integer, ByVal pnFormato As Integer, ByVal pnPorcBusqueda As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_RecuperaDatosIndicadCredEval " & pnCodicion & "," & pnFormato & "," & pnPorcBusqueda
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set RecuperaDatosIndicadCredEval = rs
    'End If
    Set rs = Nothing
End Function

Public Function RecuperaDatosGrillaCredEval(ByVal psCtaCod As String, ByVal pnTipoGrilla As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    '1 Gastos Negocio,2 Gastos Familiares,3 Otros Ingresos
    sSql = "exec stp_sel_RecuperaDatosGrillaCredEval '" & psCtaCod & "'," & pnTipoGrilla
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set RecuperaDatosGrillaCredEval = rs
    'End If
    Set rs = Nothing
End Function

Public Function RecuperaDatosRef(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_RecuperaDatosRef '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set RecuperaDatosRef = rs
    'End If
    Set rs = Nothing
End Function

Public Function RecuperaDatosVentaSem(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_RecuperaDatosVentaSem '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set RecuperaDatosVentaSem = rs
    'End If
    Set rs = Nothing
End Function

Public Function ListaCredEvalParamTipos() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ListaCredEvalParamTipos"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set ListaCredEvalParamTipos = rs
    'End If
    Set rs = Nothing
End Function

Public Function ListaCredEvalParamIndicadores() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ListaCredEvalParamIndicadores"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set ListaCredEvalParamIndicadores = rs
    'End If
    Set rs = Nothing
End Function
'END JUEZ ******************************************************************************
'WIOR 20120914 *************************************************************************
Public Function AsignarFormato(ByVal pcPrd As String, ByVal pcSPrd As String, ByVal pnMonto As Double, Optional ByRef pnMin As Double, Optional ByRef pnMax As Double) As Integer
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset

    sSql = "stp_sel_AsignarFormatoEval '" & pcPrd & "','" & pcSPrd & "'," & pnMonto

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    If rs.RecordCount > 0 Then
        If Not (rs.EOF And rs.BOF) Then
            AsignarFormato = CInt(rs!nTpoEval)
            pnMin = CDbl(rs!nMin)
            pnMax = CDbl(rs!nMax)
        Else
            AsignarFormato = 0
            pnMin = 0
            pnMax = 0
        End If
    Else
        AsignarFormato = 0
        pnMin = 0
        pnMax = 0
    End If
    
    Set rs = Nothing
End Function
Public Function RecuperaSolicitudDatoBasicos(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaSolicitudDatoBasicos
    
    sSql = "stp_sel_DatosBasicosSolicitud '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaSolicitudDatoBasicos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaSolicitudDatoBasicos:
    Err.Raise Err.Number, "Recuperar Solicitud Datos Basicos", Err.Description
End Function
Public Function RecuperaCredVerif(ByVal psAgeCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCredVerif
    
    sSql = "stp_sel_CreditosVerificados '" & psAgeCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCredVerif = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCredVerif:
    Err.Raise Err.Number, "Recuperar Creditos Verififcados", Err.Description
End Function
Public Sub ExtornarVerificacion(ByVal psCtaCod As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorExtornarVerificacion
    
    sSql = "stp_upd_ExtornoVerficacionCred '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.Ejecutar sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorExtornarVerificacion:
    Err.Raise Err.Number, "Extornar Credito Verificado", Err.Description
End Sub

Public Sub EliminaDatosFormato(ByVal psCtaCod As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorEliminaDatosFormato
    
    sSql = "stp_del_EliminarDatosFormato '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.Ejecutar sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorEliminaDatosFormato:
    Err.Raise Err.Number, "Elimina Todos los Datos del Formato", Err.Description
End Sub
Public Function RecuperaDatosCredEvalRatios(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_RecuperaDatosCredEvalRatios '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperaDatosCredEvalRatios = rs
    Set rs = Nothing
End Function
Public Function RecuperaDatosCredEvalEstadosGP(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_RecuperaDatosCredEvalEstadosGP '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperaDatosCredEvalEstadosGP = rs
    Set rs = Nothing
End Function

Public Function RecuperaDatosCredEvalTasasIngEg(ByVal psCtaCod As String, ByVal pnTipo As Integer) As Double
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim nTasa As Double

    sSql = "exec stp_sel_RecuperaDatosCredEvalTasasIngEg '" & psCtaCod & "'," & pnTipo
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    If Not (rs.EOF And rs.BOF) Then
        nTasa = CDbl(rs!nTasa)
    Else
        nTasa = 0
    End If
    oCon.CierraConexion
    Set oCon = Nothing
    
    RecuperaDatosCredEvalTasasIngEg = nTasa
    Set rs = Nothing
End Function
Public Function RecuperaDatosCredEvalGastosNeg(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_RecuperaDatosCredEvalGastosNeg '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperaDatosCredEvalGastosNeg = rs
    Set rs = Nothing
End Function



Public Function RecuperaDatosCredEvalPDT(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_RecuperaDatosCredEvalPDT '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperaDatosCredEvalPDT = rs
    Set rs = Nothing
End Function
Public Function RecuperaDatosCredEvalPDTDet(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_RecuperaDatosCredEvalPDTDet '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperaDatosCredEvalPDTDet = rs
    Set rs = Nothing
End Function
Public Function RecuperaDatosCredEvalActivosPasivos(ByVal psCtaCod As String, ByVal pnTipo As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
 
    sSql = "exec stp_sel_RecuperaDatosCredEvalActivosPasivos '" & psCtaCod & "'," & pnTipo
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperaDatosCredEvalActivosPasivos = rs
    Set rs = Nothing
End Function
Public Function RecuperaDatosCredEvalIngNeg(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
 
    sSql = "exec stp_sel_RecuperaDatosCredEvalIngNeg '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperaDatosCredEvalIngNeg = rs
    Set rs = Nothing
End Function
Public Function ListaCredEvalParamEspecializacion() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ListaCredEvalParamEspecializacion"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Set ListaCredEvalParamEspecializacion = rs
    Set rs = Nothing
End Function
Public Function RecuperaEndeudamientoPersonal(ByVal psPersCod As String, ByVal pdFecha As Date) As Double
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset

    sSql = "stp_sel_TotalEndeudamientoPersonal '" & psPersCod & "','" & Format(pdFecha, "yyyymmdd") & "'"

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    If rs.RecordCount > 0 Then
        If Not (rs.EOF And rs.BOF) Then
            RecuperaEndeudamientoPersonal = CDbl(rs!Final)
        Else
            RecuperaEndeudamientoPersonal = 0
        End If
    Else
        RecuperaEndeudamientoPersonal = 0
    End If
    
    Set rs = Nothing
End Function
Public Function AsignarEspecializacionCred(ByVal pcPrd As String, ByVal pcSPrd As String, ByVal pcCred As String, ByVal pcSCred As String, ByVal pnMonto As Double) As Integer
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset

    sSql = "stp_sel_AsignarEspecializacion '" & pcPrd & "','" & pcSPrd & "','" & pcCred & "','" & pcSCred & "'," & pnMonto

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    If rs.RecordCount > 0 Then
        If Not (rs.EOF And rs.BOF) Then
            AsignarEspecializacionCred = CInt(rs!nTpoEsp)
        Else
            AsignarEspecializacionCred = 0
        End If
    Else
        AsignarEspecializacionCred = 0
    End If
    
    Set rs = Nothing
End Function
Public Sub InsertaEspecializacionCred(ByVal psCtaCod As String, ByVal pnEspecializacion As Integer)
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

    sSql = "UPDATE ColocacCred set  nEspecializacion=" & pnEspecializacion & " where cCtaCod='" & psCtaCod & "'"

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.Ejecutar (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'WIOR **********************************************************************************

'JUEZ 20120921 *************************************************************************
Public Function VerificaSiEsCancelacionAnticipada(ByVal psCtaCod As String, ByVal psFecSis As String) As Boolean
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    VerificaSiEsCancelacionAnticipada = False
    
    sSql = "stp_sel_VerificaSiEsCancelacionAnticipada '" & psCtaCod & "','" & Format(psFecSis, "yyyymmdd") & "'"

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    If Not (rs.EOF And rs.BOF) Then
        If rs!nCancAnt = 1 Then
            VerificaSiEsCancelacionAnticipada = True
        Else
            VerificaSiEsCancelacionAnticipada = False
        End If
    Else
        VerificaSiEsCancelacionAnticipada = False
    End If
    
    Set rs = Nothing
End Function
'END JUEZ ******************************************************************************
'EJVG20121211 ***
'Public Function RecuperaMovReprogxDesastreNatxMes(ByVal psCtaCod As String, ByVal pnAnio As Integer, pnMes As Integer) As ADODB.Recordset
Public Function RecuperaMovReprogxDesastreNatxMes(ByVal psCtaCod As String, ByVal pnAnio As Integer, pnMes As Integer, ByVal psOpeCod As String) As ADODB.Recordset 'JUEZ 20131022 Se agregó psOpeCod
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    'sSql = "Exec stp_sel_MovReprogramacionxDesastreNatEnMes '" & psCtaCod & "','" & Format(pnAnio, "0000") & "','" & Format(pnMes, "00") & "'"
    sSql = "Exec stp_sel_MovReprogramacionxDesastreNatEnMes '" & psCtaCod & "','" & Format(pnAnio, "0000") & "','" & Format(pnMes, "00") & "','" & psOpeCod & "'" 'JUEZ 20131022 Se agregó psOpeCod
    oCon.AbreConexion
    Set RecuperaMovReprogxDesastreNatxMes = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'END EJVG *******

'JUEZ 20121218 ***************************************************************
Public Function RecuperaCampanas(ByVal psCodAge As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "exec stp_sel_ObtieneCampanas '" & psCodAge & "'"
    oCon.AbreConexion
    Set RecuperaCampanas = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'END JUEZ ********************************************************************
'ALPA 20121226****************************************************************
Public Function RecuperaRecaudosParaReporte(ByVal pdFecha As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "exec stp_sel_ListaRecaudo '" & Format(pdFecha, "YYYYMMDD") & "'"
    oCon.AbreConexion
    Set RecuperaRecaudosParaReporte = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'*****************************************************************************

'*** PEAC 20130227
Public Function RecuperaMantCreditos(ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "exec stp_sel_ObtieneMantCreditos '" & psPersCod & "'"
    oCon.AbreConexion
    Set RecuperaMantCreditos = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

'*** PEAC 20130308
Public Function CambiaEstadoCred(ByVal psCodCta As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "exec stp_sel_EstadoCredito '" & psCodCta & "' "
    oCon.AbreConexion
    Set CambiaEstadoCred = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

'*** PEAC 20130312
Public Function ValidaUsuarioPermitido(ByVal psUser As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "exec stp_sel_UsuarioPermitido '" & psUser & "' "
    oCon.AbreConexion
    Set ValidaUsuarioPermitido = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

'*** Plan 003 - peac
'*** PEAC 20130510
Public Sub RegDetalleCuentasReasignacion(ByVal psMovNro As String, ByVal psCtaCod As String)
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

    sSql = "exec stp_ins_CuentasReasignacion '" & psMovNro & "','" & psCtaCod & "'"

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.Ejecutar (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

'*** PEAC
Public Sub RegReasignacionCartera(ByVal psMovNro As String, ByVal pcAnaAsignado As String, ByVal pcAnaReasignado As String, ByVal pnMovFlag As Integer, ByVal pnEstadoAsigna As Integer, ByVal psMovNroResuelto As String)
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

    sSql = "exec stp_ins_TransferenciaCartera '" & psMovNro & "','" & pcAnaAsignado & "','" & pcAnaReasignado & "'," & pnMovFlag & "," & pnEstadoAsigna & ",'" & psMovNroResuelto & "'"

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.Ejecutar (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'*** PEAC 20130515
Public Sub MantReasignacionCartera(ByVal psMovNro As String, ByVal pnMovFlag As Integer, ByVal pnEstadoAsigna As Integer, ByVal psMovNroResuelto As String)
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

    sSql = "exec stp_upd_ReasignacionCartera '" & psMovNro & "'," & pnMovFlag & "," & pnEstadoAsigna & ",'" & psMovNroResuelto & "'"

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.Ejecutar (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

'*** PEAC 20130515
Public Function ConsultaReasignacionCartera(ByVal psMovNro As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta

    sSql = "exec stp_sel_verificaReasignacion '" & psMovNro & "'"

    oCon.AbreConexion
    Set ConsultaReasignacionCartera = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

'*** PEAC 20130515
'Public Function ObtieneParaReasignacionCartera(ByVal pcAge As String) As ADODB.Recordset
Public Function ObtieneParaReasignacionCartera(ByVal pcAge As String, ByVal pnTipo As Integer, ByVal pcPersCod As String) As ADODB.Recordset 'FRHU 20140220 RQ14011
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta

    'sSql = "exec stp_sel_ReasignacionCartera '" & pcAge & "' "
    sSql = "exec stp_sel_ReasignacionCartera '" & pcAge & "'," & pnTipo & ",'" & pcPersCod & "'" 'FRHU 20140220 RQ14011
    
    oCon.AbreConexion
    Set ObtieneParaReasignacionCartera = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

'*** PEAC 20130517
Public Function ValidaUsuarioReasigCartPermitido(ByVal psUser As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "exec stp_sel_UsuarioReasignaCarteraPermitido '" & psUser & "' "
    oCon.AbreConexion
    Set ValidaUsuarioReasigCartPermitido = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'*** fin plan 003 - peac

'JUEZ 20130411 ***************************************************************
Public Function RecuperaDatosComision(ByVal psNumDoc As String, ByVal pnTipoCom As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    If pnTipoCom = 1 Then
        sSql = "exec stp_sel_RecuperaDatosComisionReprogCredito '" & psNumDoc & "'"
    ElseIf pnTipoCom = 2 Then
        sSql = "exec stp_sel_RecuperaDatosComision '" & psNumDoc & "'"
    End If
    
    oCon.AbreConexion
    Set RecuperaDatosComision = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function ExisteComisionVigente(ByVal psNumDoc As String, ByVal psOpeCod As CaptacOperacion) As Boolean
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oCon As New COMConecta.DCOMConecta
    ExisteComisionVigente = False
    sSql = "exec stp_sel_ExisteCredComisionVigente '" & psNumDoc & "','" & psOpeCod & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    If Not rs.EOF Then
        ExisteComisionVigente = True
    End If
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function RecuperaProductoConcepto(ByVal pnPrdConceptoCod As String) As ADODB.Recordset 'CTI6-20210503-ERS032-2019 -(Optimizar Sugerencia)
'Public Function RecuperaProductoConcepto(ByVal pnPrdConceptoCod As Integer) As ADODB.Recordset
'pnPrdConceptoCod as Integer CTI6-20210503-ERS032-2019 -(Optimizar Sugerencia)
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    'CTI6-20210503-ERS032-2019 -(Optimizar Sugerencia)
    'sSql = "SELECT * FROM ProductoConcepto WHERE nPrdConceptoCod = '" & pnPrdConceptoCod & "'"
    sSql = "Execute spt_sel_RecuperaProductoConcepto " & CInt(pnPrdConceptoCod) & " "
    'Fin CTI6-20210503
    oCon.AbreConexion
    Set RecuperaProductoConcepto = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ExisteCreditoTitular(ByVal psPersCod As String, Optional ByVal pbAprobados As Boolean = False, Optional ByVal pbVigentes As Boolean = False, Optional ByVal pbCancelados As Boolean = False, _
                                    Optional ByVal pbSugerido As Boolean = False) As Boolean 'WIOR 20130829 SE AGREGÓ pbSugerido
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oCon As New COMConecta.DCOMConecta
    Dim lsEstados As String
    ExisteCreditoTitular = False
    If pbAprobados Then lsEstados = lsEstados & "," & gColocEstAprob
    If pbCancelados Then lsEstados = lsEstados & "," & gColocEstCancelado
    If pbVigentes Then
        lsEstados = lsEstados & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor & "," & _
                    gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor & "," & gColPEstDesem & "," & _
                    gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & "," & gColocEstRecVigJud & "," & _
                    gColocEstRecVigCast & "," & 2205 & "," & 2206
    End If
    If pbSugerido Then lsEstados = lsEstados & "," & gColocEstSug 'WIOR 20130829
    lsEstados = Mid(lsEstados, 2, Len(lsEstados) - 1)
    sSql = "exec stp_sel_ExisteCreditoTitular '" & psPersCod & "','" & lsEstados & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    If Not rs.EOF Then
        ExisteCreditoTitular = True
    End If
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'END JUEZ ********************************************************************

'WIOR 20130411 **********************************************************************************
Public Function CreditosInformeRiesgo(ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "exec stp_sel_CreditosInformeRiesgo '" & psPersCod & "' "
    oCon.AbreConexion
    Set CreditosInformeRiesgo = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Sub ExoneraPersSegDesgOpe(ByVal pnTipo As Integer, ByVal pcCtaCod As String, Optional pcPersCod As String = "%", Optional pbExonera As Boolean = False, Optional pcGlosa As String = "", _
                                Optional pbExcluidoSis As Boolean = False)
                                'WIOR 20140825 AGREGO pbExcluidoSis
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta
On Error GoTo ErrorExoneraPersSegDesg
    
If pnTipo = 1 Then
    sSql = "EXEC stp_ins_ExoneradosSegDesg '" & pcCtaCod & "','" & pcPersCod & "'," & IIf(pbExonera, 1, 0) & ",'" & pcGlosa & "'," & IIf(pbExcluidoSis, 1, 0) 'WIOR 20140825 AGREGO pbExcluidoSis
ElseIf pnTipo = 2 Then
    'sSql = "DELETE FROM ExoneraSegDes WHERE CCTACOD='" & pcCtaCod & "'" 'WIOR 20130902 COMENTO
    sSql = "EXEC stp_del_ExoneradosSegDesg '" & pcCtaCod & "','" & pcPersCod & "'" 'WIOR 20130902
End If
    
    oCon.AbreConexion
    oCon.Ejecutar (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
Exit Sub
ErrorExoneraPersSegDesg:
    Err.Raise Err.Number, "DCOMCredito:ExoneraPersSegDesg", Err.Description
End Sub

Public Function ListaPersonaExonerasSegDesg(ByVal pcCtaCod As String, Optional pcPersCod As String = "", Optional psExonera As String = "0,1") As ADODB.Recordset
'WIOR 20140825 AGREGO psExonera
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta
On Error GoTo ErrorListaPersonaExonerasSegDesg

sSql = "Exec stp_sel_ExoneradosSegDesg '" & pcCtaCod & "','" & pcPersCod & "','" & psExonera & "'" 'WIOR 20140825 AGREGO psExonera

oCon.AbreConexion
Set ListaPersonaExonerasSegDesg = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing

Exit Function
ErrorListaPersonaExonerasSegDesg:
    Err.Raise Err.Number, "DCOMCredito:ListaPersonaExonerasSegDesg", Err.Description
End Function
'WIOR FIN ***************************************************************************************

'MAVM 20130620 ***
Public Function RecuperaTpoProd_AplicaPenalidad(ByVal sConsCod As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaTpoProd_AplicaPenalidad
    sSql = "Select nConsSisValor from ConstSistema where nConsSisCod = " & sConsCod
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaTpoProd_AplicaPenalidad = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaTpoProd_AplicaPenalidad:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaTpoProd_XValor(ByVal psCtaCod As String, ByVal psValor As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaTpoProd_XValor
    sSql = "Select cTpoProdCod From Colocaciones Where cCtaCod = '" & psCtaCod & "' and cTpoProdCod in (Select Valor From dbo.fnc_getTblValoresTexto('" & psValor & "'))"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaTpoProd_XValor = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaTpoProd_XValor:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function CalculaMontoPenalidad(ByVal psCtaCod As String, ByVal nDiasAtraso As Integer, ByVal nTipoCambio As Double) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorCalculaMontoPenalidad
    sSql = "Select nImporte/(Case When Substring('" & psCtaCod & "', 9, 1) = 1 Then 1 Else " & nTipoCambio & " End) nImporte From ColocacCredPenalidadAtraso Where " & nDiasAtraso & " Between nDiasAtrasoDesde And nDiasAtrasoHasta"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CalculaMontoPenalidad = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorCalculaMontoPenalidad:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function DevolverTCMoneda(ByVal pdFecha As Date) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorDevolverTCMoneda
    sSql = "Select nVenta From TCMoneda Where datediff(day,dFecha,'" & Format$(pdFecha, "mm/dd/yyyy") & "') = 0 And nIdRangoDet = 1"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevolverTCMoneda = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorDevolverTCMoneda:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'***

'JUEZ 20130529 *****************************************************************
Public Function ExisteComisionEnvioEstadoCtaCalendario(ByVal psCtaCod As String, ByVal pnCuota As Integer) As Boolean
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oCon As New COMConecta.DCOMConecta
    ExisteComisionEnvioEstadoCtaCalendario = False
    sSql = "exec stp_sel_ExisteComisionEnvioEstadoCtaCalendario '" & psCtaCod & "'," & pnCuota
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    If Not rs.EOF Then
        ExisteComisionEnvioEstadoCtaCalendario = True
    End If
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function ExisteComisionEnvioEstadoCta(ByVal psCtaCod As String, ByVal pnCuota As Integer) As Boolean
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oCon As New COMConecta.DCOMConecta
    ExisteComisionEnvioEstadoCta = False
    sSql = "exec stp_sel_ExisteComisionEnvioEstadoCta '" & psCtaCod & "'," & pnCuota
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    If Not rs.EOF Then
        ExisteComisionEnvioEstadoCta = True
    End If
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function ExisteCuotaCalendario(ByVal psCtaCod As String, ByVal pnCuota As Integer) As Boolean
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oCon As New COMConecta.DCOMConecta
    ExisteCuotaCalendario = False
    sSql = "exec stp_sel_ExisteCuotaCalendario '" & psCtaCod & "'," & pnCuota
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    If Not rs.EOF Then
        ExisteCuotaCalendario = True
    End If
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function DevolverValorGasto(ByVal pnCodigo As Long) As Double
    Dim rsVar As New ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    If oCon.AbreConexion = False Then Exit Function
    
    sSql = "select nValor from ProductoConcepto Where nPrdConceptoCod=" & pnCodigo
    
    DevolverValorGasto = 0
    Set rsVar = oCon.CargaRecordSet(sSql)
    If Not (rsVar.EOF And rsVar.BOF) Then
        DevolverValorGasto = rsVar("nValor")
    End If
    rsVar.Close
    Set rsVar = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function MostarTextoComisionEnvioEstCredito(ByVal psCtaCod As String, ByVal pnTipo As Integer) As String
    Dim rsVar As New ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    If oCon.AbreConexion = False Then Exit Function
    
    sSql = "EXEC stp_sel_MostarTextoComisionEnvioEstCredito '" & psCtaCod & "'," & pnTipo
    
    MostarTextoComisionEnvioEstCredito = ""
    Set rsVar = oCon.CargaRecordSet(sSql)
    If Not (rsVar.EOF And rsVar.BOF) Then
        MostarTextoComisionEnvioEstCredito = rsVar("Texto")
    End If
    rsVar.Close
    Set rsVar = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'END JUEZ **********************************************************************
'WIOR 20130723 ******************************************************************
Public Sub InsActDatosAgricolasCred(ByVal psCtaCod As String, ByVal pnTipo As Integer, ByVal pnSubTipo As Integer, ByVal pnTotalHect As Double, ByVal pnHectProd As Double, ByVal pnAnimal As Long, ByVal psCodCooperativa As String)
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    On Error GoTo ErrorInsActDatoAgricolasCred

        sSql = "EXEC stp_ins_upd_CredAgricola '" & psCtaCod & "'," & pnTipo & "," & pnSubTipo & "," & pnTotalHect & "," & pnHectProd & "," & pnAnimal & ",'" & psCodCooperativa & "'"
        
        oCon.AbreConexion
        oCon.Ejecutar sSql
        oCon.CierraConexion
        Set oCon = Nothing
        
    Exit Sub
ErrorInsActDatoAgricolasCred:
        Err.Raise Err.Number, "DCOMCredito:InsActDatoAgricolasCred", Err.Description
End Sub

Public Sub EliminaAgricolasCred(ByVal psCtaCod As String)
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    On Error GoTo ErrorEliminaAgricolasCred

        sSql = "EXEC stp_del_CredAgricola '" & psCtaCod & "'"
        
        oCon.AbreConexion
        oCon.Ejecutar sSql
        oCon.CierraConexion
        Set oCon = Nothing
        
    Exit Sub
ErrorEliminaAgricolasCred:
        Err.Raise Err.Number, "DCOMCredito:EliminaAgricolasCred", Err.Description
End Sub
Public Function MostrarDatosAgricolasCred(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorMostrarDatosAgricolasCred
    sSql = "exec stp_sel_CredAgricola '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set MostrarDatosAgricolasCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorMostrarDatosAgricolasCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'WIOR FIN ***********************************************************************

'WIOR 20130924 ******************************************************************
Public Function MostrarDatosCreditosVisita(ByVal psPersCod As String, ByVal pdFecha As Date) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorMostrarDatosCreditosVisita
    sSql = "exec stp_sel_MostrarCreditosVisita '" & psPersCod & "','" & Format(pdFecha, "yyyymmdd") & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set MostrarDatosCreditosVisita = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorMostrarDatosCreditosVisita:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function MostrarDeudaCentralRiesgo(ByVal pNumDocID As String, ByVal pnTipoDocID As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorMostrarDeudaCentralRiesgo
    sSql = "exec stp_sel_DeudaCentralRiesgo '" & pNumDocID & "'," & pnTipoDocID
    
    Set oConecta = New COMConecta.DCOMConecta
    
    oConecta.AbreConexion
    Set MostrarDeudaCentralRiesgo = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorMostrarDeudaCentralRiesgo:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function


Public Function RegistrarVisitaAnalista(ByVal psPersCod As String, ByVal pdFecVisita As Date, ByVal psEntrevistado As String, _
                                        ByVal psRelacion As String, ByVal psAnalista As String, ByVal pnNumVisita As Long, _
                                        ByVal psComentarioAnalista As String, ByVal psUltimoMov As String) As Long
Dim Conn As New COMConecta.DCOMConecta
Dim rsDatos As ADODB.Recordset
Dim sSql As String

sSql = "EXEC stp_ins_VisitaCliente '" & psPersCod & "','" & Format(pdFecVisita, "yyyymmdd") & "','" & psEntrevistado & "','" & psRelacion & "','" & _
        psAnalista & "'," & pnNumVisita & ",'" & psComentarioAnalista & "','" & psUltimoMov & "'"

Conn.AbreConexion
Set rsDatos = Conn.CargaRecordSet(sSql)
Conn.CierraConexion

If Not (rsDatos.BOF And rsDatos.EOF) Then
    RegistrarVisitaAnalista = CLng(rsDatos!nCodVisita)
Else
    RegistrarVisitaAnalista = 0
End If

Set Conn = Nothing
Set rsDatos = Nothing
End Function

Public Sub RegistrarVisitaAnalistaDetalle(ByVal pnCodVisita As Long, ByVal psCtaCod As String, ByVal psObservacion As String)
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta
On Error GoTo ErrorRegistrarVisitaAnalistaDetalle

sSql = "EXEC stp_ins_VisitaClienteDet " & pnCodVisita & ",'" & psCtaCod & "','" & psObservacion & "'"

oCon.AbreConexion
oCon.Ejecutar sSql
oCon.CierraConexion
Set oCon = Nothing
    
Exit Sub
ErrorRegistrarVisitaAnalistaDetalle:
        Err.Raise Err.Number, "DCOMCredito:RegistrarVisitaAnalistaDetalle", Err.Description
End Sub
Public Function ExisteVisitaAnalista(ByVal psPersCod As String, ByVal pdFecVisita As Date) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorExisteVisitaAnalista
    sSql = "exec stp_sel_VisitaClienteExiste '" & psPersCod & "','" & Format(pdFecVisita, "yyyymmdd") & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    
    oConecta.AbreConexion
    Set ExisteVisitaAnalista = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorExisteVisitaAnalista:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function VisitaAlCliente(Optional ByVal psPersCod As String = "", Optional ByVal pnCod As Long = 0) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorVisitaAlCliente
    
    If pnCod = 0 Then
        sSql = "exec stp_sel_VisitaAlCliente '" & psPersCod & "'"
    Else
        sSql = "exec stp_sel_VisitaAlClienteXCod " & pnCod
    End If
    
    Set oConecta = New COMConecta.DCOMConecta
    
    oConecta.AbreConexion
    Set VisitaAlCliente = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorVisitaAlCliente:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Sub RegistrarVisitaJefe(ByVal pnCodVisita As Integer, ByVal pdFecVisitaJefe As Date, ByVal psJefe As String, ByVal pnNumVisitaJefe As Long, _
                                ByVal psComentarioJefe As String, ByVal psMovVisitaJefe As String)
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta
On Error GoTo ErrorRegistrarVisitaAnalistaDetalle

    
sSql = "EXEC stp_upd_VisitaCliente " & pnCodVisita & ",'" & Format(pdFecVisitaJefe, "yyyymmdd") & "','" & psJefe & "'," & pnNumVisitaJefe & ",'" & _
    psComentarioJefe & "','" & psMovVisitaJefe & "'"

oCon.AbreConexion
oCon.Ejecutar sSql
oCon.CierraConexion
Set oCon = Nothing
    
Exit Sub
ErrorRegistrarVisitaAnalistaDetalle:
        Err.Raise Err.Number, "DCOMCredito:RegistrarVisitaAnalistaDetalle", Err.Description
End Sub
Public Function VisitaAlClienteDetalle(ByVal pnCod As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorVisitaAlClienteDetalle
    
    sSql = "exec stp_sel_VisitaAlClienteDet " & pnCod

    Set oConecta = New COMConecta.DCOMConecta
    
    oConecta.AbreConexion
    Set VisitaAlClienteDetalle = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorVisitaAlClienteDetalle:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'WIOR FIN ***********************************************************************
'ALPA20130928********************************************************************
Public Function RecuperaNombreTitularCredito(ByVal scCod As String) As String
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As New ADODB.Recordset

    sSql = "select cpersNombre from productopersona PP inner join Persona P on PP.cPersCod=P.cPersCod where PP.cctacod='" & Trim(scCod) & "' and PP.nprdpersrelac= " & COMDConstantes.gColRelPersTitular

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    If rs.BOF Then
        RecuperaNombreTitularCredito = ""
    Else
        RecuperaNombreTitularCredito = rs!cPersNombre
    End If
    rs.Close
    Set rs = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function
'********************************************************************************
'JUEZ 20130925 ******************************************************************
Public Function ExisteExoneracionPreCancelacion(ByVal psCtaCod As String, ByVal pdFecSis As Date) As Boolean
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    ExisteExoneracionPreCancelacion = False
    
    sSql = "stp_sel_ExisteExoneracionPreCancelacion '" & psCtaCod & "','" & Format(pdFecSis, "yyyymmdd") & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    If Not (rs.EOF And rs.BOF) Then
        ExisteExoneracionPreCancelacion = True
    Else
        ExisteExoneracionPreCancelacion = False
    End If
    Set rs = Nothing
End Function
'END JUEZ ***********************************************************************
'JUEZ 20131022 ******************************************************************
Public Function ExisteReprogramacionxDesastreNat(ByVal psCtaCod As String, ByVal psOpeCod As String, Optional ByVal nReversion As Integer = 0) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "Exec stp_sel_ExisteReprogramacionxDesastreNat '" & psCtaCod & "','" & psOpeCod & "'," & nReversion
    oCon.AbreConexion
    Set ExisteReprogramacionxDesastreNat = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'END JUEZ ***********************************************************************
'JUEZ 20131111 ******************************************************************
Public Function CalculaMoraBN(ByVal pnTasaMora As Double, ByVal pnDiasAtrCuota As Integer, ByVal pnCapital As Double, _
                              ByVal pnIntComp As Double, ByVal pnIntGracia As Double) As Double
Dim sSql As String
Dim oConec As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    sSql = " exec sp_sel_CalculaMoraBN " & pnTasaMora & "," & pnDiasAtrCuota & "," & pnCapital & "," & pnIntComp & "," & pnIntGracia
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set R = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    CalculaMoraBN = R!nMora
End Function
'END JUEZ ***********************************************************************

'WIOR 20131011 ******************************************************************
Public Function ObtenerConfigCuotaBalon(ByVal pnTipo As Integer, ByVal psAgencia As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    If pnTipo = 1 Then
        sSql = "EXEC stp_sel_CredConfigCuotaBalonProd"
    ElseIf pnTipo = 2 Then
        sSql = "EXEC stp_sel_CredConfigCuotaBalonMontos"
    End If
    
    sSql = sSql & " '" & psAgencia & "'"
    
    oCon.AbreConexion
    Set ObtenerConfigCuotaBalon = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Sub OperacionConfigCuotaBalonSubProd(ByVal pnOpe As Integer, ByVal psAgencia As String, Optional ByVal psTpoProd As String = "")
Dim sSql As String
On Error GoTo ErrorOperacionConfigCuotaBalonSubProd

If pnOpe = 1 Then 'REGISTRA DATOS
    sSql = "EXEC stp_ins_CredConfigCuotaBalonProd '" & psAgencia & "','" & psTpoProd & "'"
ElseIf pnOpe = 2 Then 'ELIMINA LOS DATOS
    sSql = "EXEC stp_del_CredConfigCuotaBalonProd '" & psAgencia & "'"
End If

loConecta.Ejecutar sSql
    
Exit Sub
ErrorOperacionConfigCuotaBalonSubProd:
        Err.Raise Err.Number, "DCOMCredito:OperacionConfigCuotaBalonSubProd", Err.Description
End Sub
Public Sub OperacionConfigCuotaBalonRangosMontos(ByVal pnOpe As Integer, ByVal psAgencia As String, _
                                    Optional ByVal pnMontoASol As Double = 0, Optional ByVal pnMontoBSol As Double = 0, _
                                    Optional ByVal pnMontoADol As Double = 0, Optional ByVal pnMontoBDol As Double = 0)
Dim sSql As String
On Error GoTo ErrorOperacionConfigCuotaBalonRangosMontos


If pnOpe = 1 Then 'REGISTRA DATOS
    sSql = "EXEC stp_ins_CredConfigCuotaBalonMontos '" & psAgencia & "'," & pnMontoASol & "," & pnMontoBSol & "," & pnMontoADol & "," & pnMontoBDol
ElseIf pnOpe = 2 Then 'ELIMINA LOS DATOS
    sSql = "EXEC stp_del_CredConfigCuotaBalonMontos '" & psAgencia & "'"
End If

loConecta.Ejecutar sSql
    
Exit Sub
ErrorOperacionConfigCuotaBalonRangosMontos:
        Err.Raise Err.Number, "DCOMCredito:OperacionConfigCuotaBalonRangosMontos", Err.Description
End Sub

Public Function AplicaCuotaBalon(ByVal psAgencia As String, ByVal psTpoProd As String, ByVal pnMonto As Double, ByVal pnMoneda As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_PerteneceACuotaBalon '" & psAgencia & "','" & psTpoProd & "'," & pnMonto & "," & pnMoneda
    oCon.AbreConexion
    Set AplicaCuotaBalon = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'WIOR FIN ***********************************************************************
'JUEZ 20131211 ******************************************************************
Public Function RecuperaCarteraCreditoNoRevolvPorTramos(ByVal nTipoCambio As String, ByVal pdFecha As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCarteraCreditoNoRevolvPorTramos
    
    sSql = "exec stp_sel_RecuperaCarteraCreditoNoRevolvPorTramos '" & Format(pdFecha, "YYYYMMDD") & "'," & nTipoCambio
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCarteraCreditoNoRevolvPorTramos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCarteraCreditoNoRevolvPorTramos:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function RecuperaCarteraCreditoPorTramosMayor90Dias(ByVal nTipoCambio As String, ByVal pdFecha As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaRecuperaCarteraCreditoPorTramosMayor90Dias
    
    sSql = "exec stp_sel_RecuperaRecuperaCarteraCreditoPorTramosMayor90Dias '" & Format(pdFecha, "YYYYMMDD") & "'," & nTipoCambio
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCarteraCreditoPorTramosMayor90Dias = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaRecuperaCarteraCreditoPorTramosMayor90Dias:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function RecuperaCarteraCreditoHipotecarioPorTramos(ByVal nTipoCambio As String, ByVal pdFecha As String, ByVal pbMiVivTechoProp As Boolean) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCarteraCreditoHipotecarioPorTramos
    
    sSql = "exec stp_sel_RecuperaCarteraCreditoHipotecarioPorTramos '" & Format(pdFecha, "YYYYMMDD") & "'," & nTipoCambio & "," & IIf(pbMiVivTechoProp, 1, 0)
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCarteraCreditoHipotecarioPorTramos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCarteraCreditoHipotecarioPorTramos:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'END JUEZ ***********************************************************************

Public Function DetalleCarteraCrediHipotecarioporTramos(ByVal pdFecha As String, ByVal pnTpoCambio As Currency, ByVal pbMiVivTechoProp As Boolean) As ADODB.Recordset
Dim sSql As String
Dim oConecta As New COMConecta.DCOMConecta
   On Error GoTo ErrorDetalleCarteraCrediHipotecarioporTramos
   sSql = "Exec stp_sel_DetalleCarteraCreditoHipotecarioPorTramos '" & Format(pdFecha, "yyyymmdd") & "', " & pnTpoCambio & "," & IIf(pbMiVivTechoProp, 1, 0)
   oConecta.AbreConexion
   Set DetalleCarteraCrediHipotecarioporTramos = oConecta.CargaRecordSet(sSql)
   oConecta.CierraConexion
   Set oConecta = Nothing
   Exit Function
ErrorDetalleCarteraCrediHipotecarioporTramos:
   Err.Raise Err.Number, "Observaciones en el Proceso", Err.Description
End Function '****NAGL ERS020-2018

'WIOR 20131219 ******************************************************************
Public Function CreditosCanceladosSinHonramiento(ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_CreditosCanceladosSinHonramiento '" & psPersCod & "'"
    oCon.AbreConexion
    Set CreditosCanceladosSinHonramiento = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function DatosCreditoCanceladoAHonrar(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_CreditoCancelParaHonrar '" & psCtaCod & "'"
    oCon.AbreConexion
    Set DatosCreditoCanceladoAHonrar = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Sub OperacionCredHonrado(ByVal nOpe As Integer, ByVal psCtaCod As String, Optional ByVal pnMontoHonrado As Double = 0, Optional ByVal pnMontoDevuelto As Double = 0, _
                                Optional ByVal psMovHonra As String = "", Optional ByVal pdFecHonramiento As Date = "01/01/1900", Optional ByVal pnCodigo As Long = 0)
Dim sSql As String
On Error GoTo ErrorOperacionCredHonrado


Select Case nOpe
    Case 1: 'Registrar Datos
            sSql = "EXEC stp_ins_CredHonrado '" & psCtaCod & "'," & pnMontoHonrado & "," & pnMontoDevuelto & ",'" & Format(pdFecHonramiento, "yyyymmdd hh:mm:ss") & "','" & psMovHonra & "'"
    Case 2: 'Actualiza Monto Devuelto
            sSql = "EXEC stp_upd_ActualizaMontoDev '" & psCtaCod & "'," & pnMontoDevuelto
    Case 3: 'Extorno del Crédito Honrado
            sSql = "EXEC stp_upd_ExtornoCredHonrados " & pnCodigo
End Select

loConecta.Ejecutar sSql
    
Exit Sub
ErrorOperacionCredHonrado:
        Err.Raise Err.Number, "DCOMCredito:OperacionCredHonrado", Err.Description
End Sub

Public Function ObtenerCreditosHonrados(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_CreditosHonrados '" & psCtaCod & "'"
    oCon.AbreConexion
    Set ObtenerCreditosHonrados = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ObtenerCodigoCredHonrado(ByVal pnCodigo As Long, ByVal pdFecha As Date) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_CredHonradosXCtaYFecha " & pnCodigo & ",'" & Format(pdFecha, "yyyymmdd") & "'"
    oCon.AbreConexion
    Set ObtenerCodigoCredHonrado = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function CreditosHonradosAPagar(ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_CreditosHonradosPagar '" & psPersCod & "'"
    oCon.AbreConexion
    Set CreditosHonradosAPagar = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ObtenerPorcComisionHonrado(ByVal pnMontoHonrado As Double) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_ObtenerComisionPagoHonrado " & pnMontoHonrado
    oCon.AbreConexion
    Set ObtenerPorcComisionHonrado = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'WIOR FIN ***********************************************************************

'JUEZ 20140103 ******************************************************************
Public Function RecuperaDPFParaGarantia(ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "exec stp_sel_RecuperaDPFParaGarantia '" & psPersCod & "'"
    oCon.AbreConexion
    Set RecuperaDPFParaGarantia = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function RecuperaCreditosGarantiaDPF(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "exec stp_sel_RecuperaCreditosGarantiaDPF '" & psCtaCod & "'"
    oCon.AbreConexion
    Set RecuperaCreditosGarantiaDPF = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'END JUEZ ***********************************************************************
'***********************************************************
'Agregado por PASI20131127 segun TI-ERS136-2013
Public Function RecuperaExisteDesembolso(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim C As COMConecta.DCOMConecta
    sSql = "Select * from ColocConvBcoNac c " _
        & "inner join Producto p on p.cCtaCod = c.cCtaCod " _
        & "and p.nPrdEstado = 2002 " _
        & "where p.cCtaCod = '" & psCtaCod & "'"
    Set C = New COMConecta.DCOMConecta
    C.AbreConexion
    Set RecuperaExisteDesembolso = C.CargaRecordSet(sSql)
    C.CierraConexion
    Set C = Nothing
End Function
'END PASI***************************************************

'WIOR 20140201 *********************************************
Public Function ObtenerSaldoTrabDirVinc() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_SaldosTrabDirVinc"
    oCon.AbreConexion
    Set ObtenerSaldoTrabDirVinc = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ObtenerSaldoAsignaEstado(ByVal pnEstado As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_SaldoAsignacion " & pnEstado
    oCon.AbreConexion
    Set ObtenerSaldoAsignaEstado = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ObtenerSaldoAsignaTrabEstado(ByVal psPersCod As String, ByVal pnEstado As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_SaldoAsignacionTrabVinc '" & psPersCod & "'," & pnEstado
    oCon.AbreConexion
    Set ObtenerSaldoAsignaTrabEstado = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ObtenerCreditosAAsignar(ByVal pnTipo As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_CreditosAsignarSaldos " & pnTipo
    oCon.AbreConexion
    Set ObtenerCreditosAAsignar = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ObtenerCreditosTrabVinc(ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_CreditosTrabajadorVinculados '" & psPersCod & "'"
    oCon.AbreConexion
    Set ObtenerCreditosTrabVinc = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Sub ActualizaCredVinculados(ByVal psCtaCod As String, ByVal psMovAprueba As String, ByVal pnEstadoA As Integer, ByVal pnEstadoB As Integer, Optional ByVal pnNoTransac As Boolean = False)
Dim lsSQL As String
Dim oConInterno As COMConecta.DCOMConecta
On Error GoTo Error

    lsSQL = "EXEC stp_upd_CredAsignacionSaldo '" & psCtaCod & "','" & psMovAprueba & "'," & pnEstadoA & "," & pnEstadoB
    
    If pnNoTransac Then
        Set oConInterno = New COMConecta.DCOMConecta
        oConInterno.AbreConexion
        oConInterno.Ejecutar lsSQL
        oConInterno.CierraConexion
        Set oConInterno = Nothing
    Else
        loConecta.Ejecutar lsSQL
    End If
    
    Exit Sub
Error:
    Err.Raise Err.Number, "Error en Proceso(ActualizaCredVinculados)", Err.Description
End Sub
Public Function ObtenerDatosAsignacionSaldo(ByVal psCtaCod As String, ByVal pnEstado As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_CredAsignacionSaldo '" & psCtaCod & "'," & pnEstado
    oCon.AbreConexion
    Set ObtenerDatosAsignacionSaldo = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'WIOR FIN **************************************************
'ORCR20140314 ***
Public Function ReporteSaldosTrabDirVinc(ByVal UltFinMes As Integer) As ADODB.Recordset
On Error GoTo ErrorReporteSaldosTrabDirVinc
    Dim oConec As New COMConecta.DCOMConecta
    Dim sSql As String
    sSql = "exec stp_sel_ReporteSaldosTrabDirVinc " & UltFinMes 'ORCR20140621
    oConec.AbreConexion
    Set ReporteSaldosTrabDirVinc = oConec.CargaRecordSet(sSql, adLockReadOnly)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrorReporteSaldosTrabDirVinc:
    Err.Raise Err.Number, "DCOMCredito:ReporteSaldosTrabDirVinc"", Err.Description"
End Function
Public Function ReporteSaldoAsignaEstado(ByVal pnEstado As Integer, ByVal UltFinMes As Integer) As ADODB.Recordset
On Error GoTo ErrorReporteSaldoAsignaEstado
    Dim oConec As New COMConecta.DCOMConecta
    Dim sSql As String
    sSql = "exec stp_sel_ReporteSaldoAsignacion " & pnEstado & ", " & UltFinMes 'ORCR20140621
    oConec.AbreConexion
    Set ReporteSaldoAsignaEstado = oConec.CargaRecordSet(sSql, adLockReadOnly)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrorReporteSaldoAsignaEstado:
    Err.Raise Err.Number, "DCOMCredito:ReporteSaldoAsignaEstado"", Err.Description"
End Function
'END ORCR *******
'FRHU 20140220 RQ14011
Public Function ValidarJefeTerritorial(ByVal psUser As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_ValidarPersonaAgenciaJefeTerritorial '" & psUser & "'"
    oCon.AbreConexion
    Set ValidarJefeTerritorial = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'FIN FRHU 20140220
'RECO20140312 ERS174-2014**************************************
Public Function ObtieneCalifSBSRelacionCred(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As Recordset
    Dim sSql As String
    Set oConec = New COMConecta.DCOMConecta
    Set rs = New Recordset
    On Error GoTo ErrorMsg
    
    If oConec.AbreConexion() Then
        sSql = "exec stp_sel_ObtieneCalifSBSRelacionCred '" & psCtaCod & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set ObtieneCalifSBSRelacionCred = rs
    'oConec.CierraConexion
    Exit Function
ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function ObtieneDatasGarantiaCred(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As Recordset
    Dim sSql As String
    Set oConec = New COMConecta.DCOMConecta
    Set rs = New Recordset
    On Error GoTo ErrorMsg
    
    If oConec.AbreConexion() Then
        sSql = "exec stp_sel_ObtieneDatasGarantiaCred '" & psCtaCod & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set ObtieneDatasGarantiaCred = rs
    'oConec.CierraConexion
    Exit Function
ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function ObtieneDatosConvenio(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As Recordset
    Dim sSql As String
    Set oConec = New COMConecta.DCOMConecta
    Set rs = New Recordset
    On Error GoTo ErrorMsg
    
    If oConec.AbreConexion() Then
        sSql = "exec stp_sel_ObtieneDatosConvenio '" & psCtaCod & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set ObtieneDatosConvenio = rs
    'oConec.CierraConexion
    Exit Function
ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function RecuperaDatosRelacCred(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As Recordset
    Dim sSql As String
    Set oConec = New COMConecta.DCOMConecta
    Set rs = New Recordset
    On Error GoTo ErrorMsg
    
    If oConec.AbreConexion() Then
        sSql = "exec stp_sel_RecuperaDatosRelacCred '" & psCtaCod & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set RecuperaDatosRelacCred = rs
    'oConec.CierraConexion
    Exit Function
ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function RecuperaOpinionRiesgo(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As Recordset
    Dim sSql As String
    Set oConec = New COMConecta.DCOMConecta
    Set rs = New Recordset
    On Error GoTo ErrorMsg
    
    If oConec.AbreConexion() Then
        sSql = "exec stp_sel_RecuperaOpinionRiesgo '" & psCtaCod & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set RecuperaOpinionRiesgo = rs
    'oConec.CierraConexion
    Exit Function
ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function RecuperaComentarioAnalistaSugerencia(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As Recordset
    Dim sSql As String
    Set oConec = New COMConecta.DCOMConecta
    Set rs = New Recordset
    On Error GoTo ErrorMsg
    
    If oConec.AbreConexion() Then
        sSql = "exec stp_sel_RecuperaComentarioAnalistaSugerencia '" & psCtaCod & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set RecuperaComentarioAnalistaSugerencia = rs
    'oConec.CierraConexion
    Exit Function
ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'RECO FIN******************************************************
'FRHU 20140320 ERS172-2013 RQ13874
Public Function GetFechaSemanaProyectadoAge(ByVal pdFechaSist As Date, ByVal psAgeCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_GetFechaSemana '" & Format(pdFechaSist, "yyyymmdd") & "','" & psAgeCod & "'"
     oCon.AbreConexion
    Set GetFechaSemanaProyectadoAge = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Sub InsertarProyeccionColocAge(ByVal psAgeCod As String, ByVal pdFechaSemana As Date, _
                                      ByVal pnDesembMN As Double, ByVal pnDesembME As Double, _
                                      ByVal pnCrecimMN As Double, ByVal pnCrecimME As Double, _
                                      ByVal pnCarteraMN As Double, ByVal pnCarteraME As Double, _
                                      ByVal psMovNro As String)
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_upd_ProyeccionColocAge '" & psAgeCod & "','" & Format(pdFechaSemana, "yyyymmdd") & "', " & _
                                    pnDesembMN & "," & pnDesembME & "," & pnCrecimMN & "," & pnCrecimME & "," & _
                                    pnCarteraMN & "," & pnCarteraME & ",'" & psMovNro & "'"
    oCon.AbreConexion
    oCon.Ejecutar sSql
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
Public Function ValidarCargoProyeccionColocAge(ByVal psCargo As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_ValidarCargoParaProyecciones '" & psCargo & "'"
    oCon.AbreConexion
    Set ValidarCargoProyeccionColocAge = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'FIN FRHU 20140320 RQ13874

'FRHU 20140324 ERS172-2013 RQ13875
Public Function ObtenerSeguimientoProyeccionSemanal(ByVal psAnio As String, ByVal psMes As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_ObtenerProyeccionSeguimientoSemanal '" & psAnio & "','" & psMes & "'"
    oCon.AbreConexion
    Set ObtenerSeguimientoProyeccionSemanal = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function ObtenerProyeccionColocAge(ByVal psAgeCod As String, ByVal pdFechaSemana As Date) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_ObtenerProyeccionColocAge '" & psAgeCod & "','" & Format(pdFechaSemana, "yyyymmdd") & "'"
    oCon.AbreConexion
    Set ObtenerProyeccionColocAge = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'FIN FRHU 20140324 RQ13875
'FRHU 20140326 ERS172-2013 RQ13876
Public Function ObtenerProyectadoEjecutadoXAgencia(ByVal psAgeCod As String, ByVal pnMes As Integer, ByVal pnAnio As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "exec stp_sel_ProyectadoVsEjecutadoXAgencia '" & psAgeCod & "'," & pnMes & "," & pnAnio
    oCon.AbreConexion
    Set ObtenerProyectadoEjecutadoXAgencia = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'FIN FRHU 20140326 RQ13876
'FRHU 20140329 TI-ERS042-2014 RQ14177 RQ14178
Public Sub OpeComunicarRiesgo(ByVal psCtaCod As String, ByVal pnTipo As Integer, Optional psMovReg As String = "", _
                              Optional ByVal psMovCon As String = "", Optional ByVal pnEstado As Integer = 0)
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrorOpeComunicarRiesgo
    
    sSql = "exec stp_ins_upd_OpeComunicarRiesgo '" & psCtaCod & "','" & psMovReg & "','" & psMovCon & "'," & pnEstado & "," & pnTipo
    
    oConecta.AbreConexion
    oConecta.CargaRecordSet (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorOpeComunicarRiesgo:
    Err.Raise Err.Number, "Error en Proceso", Err.Description
End Sub
Public Function ObtenerComunicarRiesgo(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrorObtenerComunicarRiesgo
    
    sSql = "EXEC stp_sel_ComunicarRiesgo '" & psCtaCod & "'"
    
    oConecta.AbreConexion
    Set ObtenerComunicarRiesgo = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Exit Function
ErrorObtenerComunicarRiesgo:
    Err.Raise Err.Number, "Error en Proceso", Err.Description
End Function
'RQ14179
Public Function ObtenerComunicarRiesgoPend() As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrorObtenerComunicarRiesgoPend
    
    sSql = "EXEC stp_sel_ComunicarRiesgoPendiente "
    
    oConecta.AbreConexion
    Set ObtenerComunicarRiesgoPend = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Exit Function
ErrorObtenerComunicarRiesgoPend:
    Err.Raise Err.Number, "Error en Proceso", Err.Description
End Function
'RQ14180
Public Function ValidarComunicarRiesgoPend(ByVal psCtaCod As String) As Boolean
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oConecta As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_ComunicarRiesgo '" & psCtaCod & "'"
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    If Not rs.EOF And Not rs.BOF Then
        ValidarComunicarRiesgoPend = True
    Else
        ValidarComunicarRiesgoPend = False
    End If
    
End Function
'FIN FRHU TI-ERS042-2014
'FRHU 20140401 TI-ERS027-2014
'RQ14132
Public Function ObtenerClientePorDireccion(ByVal psDireccion As String, ByVal psDistrito As String) As ADODB.Recordset
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oConecta As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_BuscarPersonaPorDireccion '" & psDireccion & "','" & psDistrito & "'"
    oConecta.AbreConexion
    Set ObtenerClientePorDireccion = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'FIN FRHU 20140401 TI-ERS027-2014
'FRHU 20140402 ERS026-2014
'RQ14129
Public Function ObtenerUltimaActualizacion(ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_ObtenerUltimaActualizacion '" & psPersCod & "'"
    oConecta.AbreConexion
    Set ObtenerUltimaActualizacion = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'RQ14131
Public Function ObtenerUltimaActualizacionXCredito(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_ObtenerUltimaActualizacionXCredito '" & psCtaCod & "'"
    oConecta.AbreConexion
    Set ObtenerUltimaActualizacionXCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'FIN FRHU 20140402
'JAME 20140328
Public Function ObtenerMotivoRefinanciamiento(ByVal cCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_MotivoRefinanciamiento '" & cCtaCod & "'"
    oCon.AbreConexion
    Set ObtenerMotivoRefinanciamiento = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'FIN JAME
'WIOR 20140502 *****************************************************************
Public Function VerificaAmpliacionPlazoCreditoConvenio(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    sSql = "exec stp_sel_VerificaAmpliacionPlazoCreditoConvenio '" & psCtaCod & "'"
    oConecta.AbreConexion
    Set VerificaAmpliacionPlazoCreditoConvenio = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function RecuperaCreditoConvenioAmpliacionPlazo(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    sSql = "exec stp_sel_RecuperaCreditoConvenioAmpliacionPlazo '" & psCtaCod & "'"
    oConecta.AbreConexion
    Set RecuperaCreditoConvenioAmpliacionPlazo = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'END WIOR **********************************************************************
'ALPA2014011*********************************************************************
Public Function RecuperaDatosCreditosLeasingPreCancelados(ByVal psCtaCod As String, ByVal pnNroCalen As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "exec stp_sel_ObtenerDatosParaPagoPreCanceladosLeasing '" & psCtaCod & "'," & pnNroCalen
    oCon.AbreConexion
    Set RecuperaDatosCreditosLeasingPreCancelados = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ObtenerProximaCuotaPago(ByVal pnNroCalendario As Integer, ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtenerProximaCuotaPago
    sSql = "exec stp_sel_ObtenerProximaCuotaPago " & pnNroCalendario & ",'" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerProximaCuotaPago = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtenerProximaCuotaPago:
    Err.Raise Err.Number, "Error En Proceso: DCOMCredito:::ObtenerProximaCuotaPago", Err.Description
End Function
'********************************************************************************
'ALPA20150113*******************************************************************
'Public Function RecuperaAutorizacionCreditos(ByVal psPersCod As String, ByVal psTpoProdCod As String, ByVal pnCampania As Integer, ByVal pnEdad As Integer, ByVal pbPlazo As Integer, ByVal pnPlazoValor As Integer, ByVal psAgencia As String, ByVal pnDestino As Integer, ByVal cMoneda As String, ByVal nMonto As Currency) As ADODB.Recordset 'Comentado JOEP20180202
Public Function RecuperaAutorizacionCreditos(ByVal psPersCod As String, ByVal psTpoProdCod As String, ByVal pnCampania As Integer, ByVal pnEdad As Integer, ByVal pbPlazo As Integer, ByVal pnPlazoValor As Integer, ByVal psAgencia As String, ByVal pnDestino As Integer, ByVal cMoneda As String, ByVal nMonto As Currency, Optional pnCondicion As Integer = 0, Optional pnPersCodRelac As Integer = 0) As ADODB.Recordset 'JOEP20180202 Acta 021-2018 Agrego Optional pnCondicion As Integer = 0 - JOEP20180115 Acta 158-2018 pnPersCodRelac
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    'sSql = "exec stp_sel_AutorizarSolicitudSugerenciaCreditos '" & psPersCod & "','" & psTpoProdCod & "'," & pnCampania & ", " & pnEdad & ", " & pbPlazo & ", " & pnPlazoValor & ", '" & psAgencia & "', " & pnDestino & ",'" & cMoneda & "'," & nMonto 'Comentado JOEP20180202
     sSql = "exec stp_sel_AutorizarSolicitudSugerenciaCreditos '" & psPersCod & "','" & psTpoProdCod & "'," & pnCampania & ", " & pnEdad & ", " & pbPlazo & ", " & pnPlazoValor & ", '" & psAgencia & "', " & pnDestino & ",'" & cMoneda & "'," & nMonto & "," & pnCondicion & "," & pnPersCodRelac & "" 'JOEP20180202 Acta 021-2018 Agrego Optional pnCondicion As Integer = 0 - JOEP20180115 Acta 158-2018 pnPersCodRelac
    oConecta.AbreConexion
    Set RecuperaAutorizacionCreditos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'END ALPA **********************************************************************
'WIOR 20140509 **********************
Public Function VerificaGruposRegPromotores(ByVal psGrupos As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_VerificaGruposRegPromotoresCreditos '" & psGrupos & "'"
    oConecta.AbreConexion
    Set VerificaGruposRegPromotores = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function RecuperaCredGarantiaPoliza(ByVal psCtaCod As String) As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String

    On Error GoTo ErrorRecuperaCredGarantiaPoliza
    

    sSql = "EXEC stp_sel_CredGarantiasConPoliza '" & psCtaCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCredGarantiaPoliza = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaCredGarantiaPoliza:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function RecuperaCredPagarAutomatico(ByVal pdFecha As Date) As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String

    On Error GoTo ErrorRecuperaCredPagarAutomatico

    sSql = "EXEC stp_sel_CreditosApagarAutomatico '" & Format(pdFecha, "yyyymmdd") & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCredPagarAutomatico = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaCredPagarAutomatico:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function RecuperaCuentasAhorroPagarAutomatico(ByVal psCtaCod As String) As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String

    On Error GoTo ErrorRecuperaCuentasAhorroPagarAutomatico

    sSql = "EXEC stp_sel_RecuperaCuentasADebitar '" & psCtaCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCuentasAhorroPagarAutomatico = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaCuentasAhorroPagarAutomatico:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'WIOR FIN ***************************

'WIOR 20140725 *********************************************
Public Function VerificaGarantiaDPF(ByVal psCtaCod As String) As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String

    On Error GoTo ErrorVerificaGarantiaDPF

    sSql = "EXEC stp_sel_CredVerificaGarantiaDPF '" & psCtaCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set VerificaGarantiaDPF = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorVerificaGarantiaDPF:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'WIOR FIN **************************************************

'WIOR 20140820 *********************************************
Public Function ConfigEdadProducto(ByVal psTpoCred As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_CredConfigEdadProd '" & psTpoCred & "'"
    oConecta.AbreConexion
    Set ConfigEdadProducto = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function ConfigPeriodoProducto(ByVal psTpoCred As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_CredConfigPeriodoProd '" & psTpoCred & "'"
    oConecta.AbreConexion
    Set ConfigPeriodoProducto = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function RecuperaRelacPersSegDes(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorRecuperaRelacPersSegDes

    sSql = "EXEC stp_sel_RecuperaRelacPersSegDes '" & psCtaCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaRelacPersSegDes = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaRelacPersSegDes:
    Err.Raise Err.Number, "Error En Proceso de Recuperar Relaciones de Personas con el Credito para el Seg. Desgravamente", Err.Description
End Function
'WIOR FIN **************************************************

'JAME 20140826************ '
Public Sub RegistroHistorialConvenio(ByVal psCtaCod As String, ByVal psCuser As String, ByVal psFecha As String, ByVal psConvenioOrigen As String, ByVal psConvenioDestino As String, ByVal pnTipo As Integer)
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrorOpeComunicarRiesgo
    
    sSql = "exec stp_ins_RegistroHistorialConvenio '" & psCtaCod & "','" & psCuser & "','" & psFecha & "','" & psConvenioOrigen & "','" & psConvenioDestino & "'," & pnTipo
    
    oConecta.AbreConexion
    oConecta.CargaRecordSet (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorOpeComunicarRiesgo:
    Err.Raise Err.Number, "Error en Proceso", Err.Description
End Sub

Public Function ObtenerHistorialConvenio(ByVal sConvenioOrigen As String, ByVal sConvenioDestino As String, sFechaDesde As String, sFechaHasta As String, nTipoBusqueda As Integer, nTipo As Integer, ByVal psAgeCad As String) As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String

    On Error GoTo ErrorVerificaGarantiaDPF

    sSql = "EXEC stp_sel_ObtenerHistorialConvenio '" & sConvenioOrigen & "','" & sConvenioDestino & "','" & Format(sFechaDesde, "yyyy/MM/dd") & "','" & Format(sFechaHasta, "yyyy/MM/dd") & "'," & nTipoBusqueda & "," & nTipo & ",'" & psAgeCad & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerHistorialConvenio = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorVerificaGarantiaDPF:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function HistorialConvenioReasigna(ByVal sConvenioOrigen As String, ByVal sConvenioDestino As String, ByVal nTpoBusqueda As Integer) As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String

    On Error GoTo ErrorVerificaGarantiaDPF

    sSql = "EXEC stp_sel_HistorialConvenioReasigna '" & sConvenioOrigen & "','" & sConvenioDestino & "'," & nTpoBusqueda

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set HistorialConvenioReasigna = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorVerificaGarantiaDPF:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'JAME FIN**********'
'ALPA 20141201************************************
Public Function RecuperaVerEntidades(ByVal lsCtaCod As String, ByVal lsPersCod As String, ByVal lsDocumento As String) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    Set oConecta = New COMConecta.DCOMConecta
    sSql = "exec stp_sel_EndeudamientoEntidades '" & lsCtaCod & "','" & lsPersCod & "','" & lsDocumento & "'"
    oConecta.AbreConexion
    Set RecuperaVerEntidades = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function CantidadCreditosVerEntidades(ByVal lsCtaCod As String) As Integer
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oConecta = New COMConecta.DCOMConecta
    sSql = "exec stp_sel_CantidadCreditosVerEntidades '" & lsCtaCod & "'"
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    If Not rs.EOF And Not rs.BOF Then
        CantidadCreditosVerEntidades = rs!nCantidad
    Else
        CantidadCreditosVerEntidades = 0
    End If
    Set oConecta = Nothing
End Function
Public Function ObtenerPorcentajeMoraAgencia(ByVal pdFecha As Date, ByVal psAgeCod As String) As Double
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oConecta = New COMConecta.DCOMConecta
    sSql = "exec stp_sel_PorcentajeMoraAgencia '" & Format(pdFecha, "YYYY/MM/DD") & "', '" & psAgeCod & "'"
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    If Not rs.EOF And Not rs.BOF Then
        ObtenerPorcentajeMoraAgencia = rs!nPorcentaje
    Else
        ObtenerPorcentajeMoraAgencia = 0
    End If
    Set oConecta = Nothing
End Function
Public Function ObtenerExcedenteFI(ByVal cNumFuente As String) As Long
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
        sSql = "exec stp_sel_ObtieneExcedenteFI '" & cNumFuente & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    If Not rs.EOF And Not rs.BOF Then
        ObtenerExcedenteFI = rs!nExcedente
    Else
        ObtenerExcedenteFI = 0
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function ObtenerCalificacionNormal(ByVal cNumeroDocumento As String) As Integer
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
        sSql = "exec stp_sel_CalificacionNormal '" & cNumeroDocumento & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    If Not rs.EOF And Not rs.BOF Then
        ObtenerCalificacionNormal = rs!nNormal
    Else
        ObtenerCalificacionNormal = -1
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'*************************************************
Public Sub GrabarVerEntidadesEndeudamiento(ByVal psCtaCod As String, ByVal prsVerEntidad As ADODB.Recordset, ByVal pdHoy As String, ByVal psCodAge As String, ByVal psCodUser As String)

Dim oCredito As COMDCredito.DCOMCredActBD
Dim i As Integer
Dim lsMovNro As String

On Error GoTo ErrorGrabarVerEntidadesEndeudamiento


    Set oCredito = New COMDCredito.DCOMCredActBD
    oCredito.dBeginTrans
        
    lsMovNro = oCredito.GeneraMovNro(pdHoy, psCodAge, psCodUser)

    Call oCredito.dDeleteEndeudamientoEntidades(psCtaCod, False)
    If prsVerEntidad.RecordCount > 0 Then
        If (prsVerEntidad.BOF = False Or prsVerEntidad.EOF = True) Then
            prsVerEntidad.MoveFirst
        End If
    End If
    If Not (prsVerEntidad.BOF Or prsVerEntidad.EOF) Then
        
        For i = 0 To prsVerEntidad.RecordCount - 1
           ' If prsExonera!Solicitar = 1 Then
                Call oCredito.GrabarVerEntidades(CStr(prsVerEntidad!tipo), CStr(psCtaCod), CStr(prsVerEntidad!cPersCod), CInt(prsVerEntidad!Codigo), CInt(prsVerEntidad!bAnulacion), CStr(prsVerEntidad!cod_Edu), CStr(lsMovNro), CStr(prsVerEntidad!cDocumento), False)
           ' End If
             prsVerEntidad.MoveNext
        Next i
    End If
    oCredito.dCommitTrans
    Set oCredito = Nothing

    Exit Sub

ErrorGrabarVerEntidadesEndeudamiento:
    oCredito.dRollbackTrans
    Err.Raise Err.Number, "Error En Proceso GrabarVerEntidadesEndeudamiento", Err.Description
End Sub
'ALPA 20150113********************************************************************************
Public Function RecuperaProductoCredicioActivo(ByVal psTpoCredCod As String, ByVal pnCampania As Integer) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    Set oConecta = New COMConecta.DCOMConecta
    sSql = "exec stp_sel_ProductoCrediticioActivo '" & psTpoCredCod & "'," & pnCampania & ""
    oConecta.AbreConexion
    Set RecuperaProductoCredicioActivo = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecValidaProcentajeCredito(ByVal pdFecha As Date, ByVal psAgeCod As String, ByVal pnMonto As Currency, ByVal psLineaCod As String) As ADODB.Recordset

Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String

    On Error GoTo ErrorRecValidaProcentajeCredito

    sSql = "EXEC stp_sel_ValidaProcentajeCredito '" & Format(pdFecha, "yyyymmdd") & "','" & psAgeCod & "'," & pnMonto & ",'" & psLineaCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecValidaProcentajeCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecValidaProcentajeCredito:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'END ALPA *************************************************************************************
'WIOR 20150401 ************************
Public Function ObtenerSolPerdonMoraCampRecupCred(ByVal pcCtaCod As String, ByVal pnCampRecup As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

Set oConecta = New COMConecta.DCOMConecta
sSql = "EXEC stp_sel_ObtenerSolPerdonMoraCampRecupCred '" & pcCtaCod & "'," & pnCampRecup & ""

oConecta.AbreConexion
Set ObtenerSolPerdonMoraCampRecupCred = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion

Set oConecta = Nothing
End Function

Public Function ObtenerConfigCampRecup(ByVal pnCampRecup As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

Set oConecta = New COMConecta.DCOMConecta
sSql = "EXEC stp_sel_ObtenerConfigCampRecup " & pnCampRecup & ""

oConecta.AbreConexion
Set ObtenerConfigCampRecup = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion

Set oConecta = Nothing
End Function

Public Function ObtenerModalidadesCampRecup(ByVal pnCampRecup As Integer, ByVal pnDiasAtraso As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

Set oConecta = New COMConecta.DCOMConecta
sSql = "EXEC stp_sel_ObtenerModalidadesCampRecup " & pnCampRecup & "," & pnDiasAtraso

oConecta.AbreConexion
Set ObtenerModalidadesCampRecup = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion

Set oConecta = Nothing
End Function

Public Function GrabarPerdonMoraCampRecup(ByVal pcCtaCod As String, ByVal pnNroCalen As Integer, ByVal nCampRecup As Integer, ByVal nModalidad As Integer, ByVal nCuotasDejadas As Integer, _
                                          ByVal pnPorcentaje As Double, ByVal nMontoPerdon As Double, ByVal cMovNroSol As String, ByVal pdFecha As Date) As Double
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As New ADODB.Recordset

Set oConecta = New COMConecta.DCOMConecta
sSql = "EXEC stp_ins_RegistrarSolPerdonMoraCampRecup '" & pcCtaCod & "'," & pnNroCalen & "," & nCampRecup & "," & nModalidad & "," & nCuotasDejadas & "," & pnPorcentaje & "," & nMontoPerdon & ",'" & cMovNroSol & "','" & Format(pdFecha, "yyyymmdd") & "'"

oConecta.AbreConexion
Set rs = oConecta.CargaRecordSet(sSql)
If Not rs.EOF And Not rs.BOF Then
    GrabarPerdonMoraCampRecup = rs!id
Else
    GrabarPerdonMoraCampRecup = 0
End If
oConecta.CierraConexion

Set oConecta = Nothing
End Function

Public Sub GrabarPerdonMoraCampRecupDet(ByVal pnId As Double, ByVal pnCuota As Integer, ByVal pnMonto As Double, ByVal pnMontoPerdonado As Double)
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta
On Error GoTo ErrorGrabarPerdonMoraCampRecupDet

sSql = "EXEC stp_ins_RegistrarSolPerdonMoraCampRecupDet " & pnId & "," & pnCuota & "," & pnMonto & "," & pnMontoPerdonado

oCon.AbreConexion
oCon.Ejecutar sSql
oCon.CierraConexion
Set oCon = Nothing
    
Exit Sub
ErrorGrabarPerdonMoraCampRecupDet:
        Err.Raise Err.Number, "DCOMCredito:ErrorGrabarPerdonMoraCampRecupDet", Err.Description
End Sub

Public Sub ActualizaMontosPerdonMoraCampRecup(ByVal pnId As Double)
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta
On Error GoTo ErrorActualizaMontosPerdonMoraCampRecup

sSql = "EXEC stp_upd_ActualizaMontoSolPerdonMoraCampRecup " & pnId

oCon.AbreConexion
oCon.Ejecutar sSql
oCon.CierraConexion
Set oCon = Nothing
    
Exit Sub
ErrorActualizaMontosPerdonMoraCampRecup:
        Err.Raise Err.Number, "DCOMCredito:ActualizaMontosPerdonMoraCampRecup", Err.Description
End Sub

Public Sub ObtenerPerdonMoraCampRecupXCta(ByVal pcCtaCod As String, ByVal pdFecha As Date, ByRef pnId As Double, ByRef pnMontoMin As Double, ByRef pnMontoMax As Double)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As New ADODB.Recordset

Set oConecta = New COMConecta.DCOMConecta
sSql = "EXEC stp_sel_SolPerdonMoraCampRecupXCta '" & pcCtaCod & "','" & Format(pdFecha, "yyyymmdd") & "'"

oConecta.AbreConexion
Set rs = oConecta.CargaRecordSet(sSql)
If Not (rs.EOF And rs.BOF) Then
    pnId = CDbl(rs!nId)
    pnMontoMin = CDbl(rs!nMontoMin)
    pnMontoMax = CDbl(rs!nMontoMax)
Else
    pnId = 0
    pnMontoMin = 0
    pnMontoMax = 0
End If
oConecta.CierraConexion

Set oConecta = Nothing
End Sub
Public Sub ActualizaPerdonMoraCampRecupXPago(ByVal pnId As Double, ByVal pcMovNro As String, ByVal pnMovNro As Double)
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta
On Error GoTo ErrorActualizaPerdonMoraCampRecupXPago

sSql = "EXEC stp_upd_ActPerdonMoraCampRecupXPago " & pnId & ",'" & pcMovNro & "'," & pnMovNro

oCon.AbreConexion
oCon.Ejecutar sSql
oCon.CierraConexion
Set oCon = Nothing
    
Exit Sub
ErrorActualizaPerdonMoraCampRecupXPago:
        Err.Raise Err.Number, "DCOMCredito:ActualizaPerdonMoraCampRecupXPago", Err.Description
End Sub
'WIOR FIN *****************************
'RECO20150318 ERS010-2015***********************************************************
'MARG 20160719*************************
Public Function AdmCredListaCredAprobados(ByVal psCadAge As String) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrorAdmCredListaCredAprobados
    sSql = "EXEC stp_sel_AdmCredListaCredAprobados2 '" & psCadAge & "'" '
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set AdmCredListaCredAprobados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorAdmCredListaCredAprobados:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Sub RegistraControlDesembolso(ByVal psCtaCod As String, ByVal psIngreso As String, ByVal psUser As String)
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrorRegistraControlDesembolso
    
    sSql = "exec stp_ins_RegistraControlDesembolso '" & psCtaCod & "','" & psIngreso & "','" & psUser & "'"
    
    oConecta.AbreConexion
    oConecta.CargaRecordSet (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorRegistraControlDesembolso:
    Err.Raise Err.Number, "Error en Proceso", Err.Description
End Sub
Public Sub RegistraControlDesembolsoHistoral(ByVal nIdControl As Long, ByVal nTpoOpe As Integer, ByVal dUltMov As String)
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrorRegistraControlDesembolso
    
    sSql = "exec stp_ins_RegistraControlDesembolsoHistorial " & nIdControl & "," & nTpoOpe & "','" & dUltMov & "'"
    
    oConecta.AbreConexion
    oConecta.CargaRecordSet (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorRegistraControlDesembolso:
    Err.Raise Err.Number, "Error en Proceso", Err.Description
End Sub
Public Sub RegistroCredAdmSeccion(ByVal psTpoCred As String, ByVal psDescripcion As String, ByVal psUltimaActualizacion As String, ByVal pnEstado As Integer)
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrorRegistroCredAdmSeccione
    
    sSql = "exec stp_ins_RegistroCredAdmSeccion '" & psTpoCred & "','" & psDescripcion & "','" & psUltimaActualizacion & "'," & pnEstado
    
    oConecta.AbreConexion
    oConecta.CargaRecordSet (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorRegistroCredAdmSeccione:
    Err.Raise Err.Number, "Error en Proceso", Err.Description
End Sub
'Public Function ListaCredAdmSecciones(ByVal psTpoCred As String) As ADODB.Recordset'Comento JOEP20190208 CP
Public Function ListaCredAdmSecciones(ByVal psCtaCod As String, ByVal psTpoCateg As String, ByVal psTpoProd As String, ByVal pnMonto As Currency, ByVal psTpoCred As String, Optional ByVal psTpoOpe As Integer = 0) As ADODB.Recordset 'JOEP20181229 CP
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrorListaCredAdmSecciones
    'sSql = "EXEC stp_sel_CredAdmSecciones '" & psTpoCred & "'"'Comento JOEP20190208 CP
    sSql = "EXEC stp_sel_CatalogoCheckListObtieneCabecera '" & psCtaCod & "','" & psTpoCateg & "','" & psTpoProd & "'," & pnMonto & ",'" & psTpoCred & "'," & psTpoOpe & "" 'JOEP29122018 CP
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ListaCredAdmSecciones = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorListaCredAdmSecciones:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function


'JOEP20190107 CP
Public Function ListaCredAdmNiveles(ByVal psTpoCateg As String, ByVal psTpoProd As String, ByVal pcItem As String, ByVal pnCantConfig As Integer) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrorListaCredAdmNiveles
    sSql = "EXEC stp_sel_CatalogoCheckListObtieneNiveles '" & psTpoCateg & "','" & psTpoProd & "','" & pcItem & "'," & pnCantConfig & ""
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ListaCredAdmNiveles = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorListaCredAdmNiveles:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'JOEP20190107 CP

Public Sub CredAdmActualizaSeccion(ByVal pnIdSeccion As Long, ByVal pnEstado As Integer, ByVal psUltActualizacion As String)
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrorCredAdmActualizaSeccion
    
    sSql = "exec stp_Upd_CredAdmActualizaSeccion " & pnIdSeccion & "," & pnEstado & ",'" & psUltActualizacion & "'"
    
    oConecta.AbreConexion
    oConecta.CargaRecordSet (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorCredAdmActualizaSeccion:
    Err.Raise Err.Number, "Error en Proceso", Err.Description
End Sub
Public Sub RegistroCredAdmRequisito(ByVal nIdSeccion As Long, ByVal psDescripcion As String, ByVal psUltimaActualizacion As String, ByVal pnEstado As Integer)
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrorRegistroCredAdmRequisito
    
    sSql = "exec stp_ins_RegistroCredAdmRequisitos " & nIdSeccion & ",'" & psDescripcion & "','" & psUltimaActualizacion & "'," & pnEstado
    
    oConecta.AbreConexion
    oConecta.CargaRecordSet (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorRegistroCredAdmRequisito:
    Err.Raise Err.Number, "Error en Proceso", Err.Description
End Sub
'Public Function ListaCredAdmRequisito(ByVal pnIdSeccion As String) As ADODB.Recordset'Comento JOEP20190208 CP
Public Function ListaCredAdmRequisito(ByVal psTpoCateg As String, ByVal psTpoProd As String, ByVal psItem As String, ByVal pnCantConf As Long) As ADODB.Recordset 'JOEP20181229 CP
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrorListaCredAdmRequisito
    'sSql = "EXEC stp_sel_CredAdmRequisitos '" & pnIdSeccion & "'"'Comento JOEP20190208 CP
    sSql = "EXEC stp_sel_CatalogoCheckListObtieneCabDetalle '" & psTpoCateg & "','" & psTpoProd & "','" & psItem & "'," & pnCantConf & "" 'JOEP20190208 CP
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ListaCredAdmRequisito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorListaCredAdmRequisito:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Sub CredAdmActualizaRequisito(ByVal pnIdRequisito As Long, ByVal pnEstado As Integer, ByVal psUltActualizacion As String)
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrorCredAdmActualizaRequisito
    
    sSql = "exec stp_Upd_CredAdmActualizaRequisito " & pnIdRequisito & "," & pnEstado & ",'" & psUltActualizacion & "'"
    
    oConecta.AbreConexion
    oConecta.CargaRecordSet (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorCredAdmActualizaRequisito:
    Err.Raise Err.Number, "Error en Proceso", Err.Description
End Sub
Public Function SegSolicitudCIHistorial(ByVal pnIdSolicitud As Long, Optional ByVal pnTpoSeguro As Integer = 0) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrorSegSolicitudCIHistorial
    sSql = "EXEC stp_sel_SegSolicitudCIHistorial '" & pnIdSolicitud & "'," & pnTpoSeguro
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set SegSolicitudCIHistorial = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorSegSolicitudCIHistorial:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Sub ActualizaControlDesembolso(ByVal pnIdControl As Long, ByVal psIngreso As String, ByVal pnTpoOpe As Integer)
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrorCredAdmActualizaRequisito
    
    sSql = "exec stp_upd_ActualizaControlDesembolso " & pnIdControl & ",'" & psIngreso & "'," & pnTpoOpe
    
    oConecta.AbreConexion
    oConecta.CargaRecordSet (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorCredAdmActualizaRequisito:
    Err.Raise Err.Number, "Error en Proceso", Err.Description
End Sub
Public Sub RegistraRequisitoCta(ByVal psCtaCod As String, ByVal pnIdRequisito As Long, ByVal pnEstado As Integer, ByVal pnTpoOpe As Integer)
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrorCredAdmActualizaRequisito
    oConecta.AbreConexion
    sSql = "exec stp_ins_RegistraRequisitoCta '" & psCtaCod & "'," & pnIdRequisito & "," & pnEstado
    oConecta.CargaRecordSet (sSql)
    
    If pnTpoOpe <> 2 Or pnTpoOpe <> 5 Then
        sSql = "stp_ins_RegistraRequisitoCtaHistorico '" & psCtaCod & "'," & pnIdRequisito & "," & pnEstado & "," & pnTpoOpe
    End If
    
    oConecta.CargaRecordSet (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorCredAdmActualizaRequisito:
    Err.Raise Err.Number, "Error en Proceso", Err.Description
End Sub
Public Function ObtieneDatosAutorizaChekList(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrorObtieneDatosAutorizaChekList
    sSql = "EXEC stp_sel_ObtieneDatosAutorizaChekList '" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneDatosAutorizaChekList = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtieneDatosAutorizaChekList:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Sub CredAdmRegistraAutorizacionChekList(ByVal psCtaCod As String, ByVal psPersCodAuto As String, ByVal psGlosa As String, ByVal psFecAuto As String, ByVal pnEstado As Integer)
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrorCredAdmRegistraAutorizacionChekList
    
    sSql = "exec stp_ins_CredAdmRegistraAutorizacionChekList '" & psCtaCod & "','" & psPersCodAuto & "','" & psGlosa & "','" & psFecAuto & "'," & pnEstado
    
    oConecta.AbreConexion
    oConecta.CargaRecordSet (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorCredAdmRegistraAutorizacionChekList:
    Err.Raise Err.Number, "Error en Proceso", Err.Description
End Sub
Public Function CredAdmObtieneCtaAutorizadaChekList(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrorCredAdmObtieneCtaAutorizadaChekList
    sSql = "EXEC stp_sel_CredAdmObtieneCtaAutorizadaChekList '" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CredAdmObtieneCtaAutorizadaChekList = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorCredAdmObtieneCtaAutorizadaChekList:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'Public Function ObtieneRequisitosTpoCred(ByVal psTpoCred As String, ByVal psCtaCod As String) As ADODB.Recordset'Comento JOEP20190208 CP
Public Function ObtieneRequisitosTpoCred(ByVal psTpoCatg As String, ByVal psTpoProd As String, ByVal psCtaCod As String, ByVal pnMonto As Currency, ByVal psTpoCred As String, Optional ByVal psTpoOpe As Integer = 0) As ADODB.Recordset 'JOEP20190102 CP
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrorObtieneRequisitosTpoCred
    'sSql = "EXEC stp_sel_ObtieneRequisitosTpoCred '" & psTpoCred & "','" & psCtaCod & "'"'Comento JOEP20190208 CP
    sSql = "EXEC stp_sel_CatalogoCheckListObtieneDetalleProd '" & psCtaCod & "','" & psTpoCatg & "','" & psTpoProd & "'," & pnMonto & ",'" & psTpoCred & "'," & psTpoOpe & ""   'JOEP20190102 CP
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneRequisitosTpoCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtieneRequisitosTpoCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function DevuelveDatosAutorizacion(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrorDevuelveDatosAutorizacion
    sSql = "EXEC stp_sel_DevuelveDatosAutorizacion '" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevuelveDatosAutorizacion = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorDevuelveDatosAutorizacion:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function DevuelveDatosCred(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrorDevuelveDatosCred
    sSql = "EXEC stp_sel_DevuelveDatosCred '" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevuelveDatosCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorDevuelveDatosCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Sub EliminaCheckListCta(ByVal psCtaCod As String)
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrorCredAdmRegistraAutorizacionChekList
    
    sSql = "exec stp_del_EliminaCheckListCta '" & psCtaCod & "'"
    
    oConecta.AbreConexion
    oConecta.CargaRecordSet (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorCredAdmRegistraAutorizacionChekList:
    Err.Raise Err.Number, "Error en Proceso", Err.Description
End Sub
Public Function ObtenerDatosCredChekList(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrorDevuelveDatosCred
    sSql = "EXEC stp_sel_ObtenerDatosCredChekList '" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerDatosCredChekList = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorDevuelveDatosCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'RECO FIN***************************************************************************
'WIOR 20150525 ***
Public Function RecuperaTipoCredCabecera() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    sSql = "EXEC stp_sel_TipoCreditoCabecera"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaTipoCredCabecera = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function RecuperarCampanaRecup(Optional ByVal pnCod As Long = 0) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_CampanasRecup " & pnCod
    
    oCon.AbreConexion
    Set RecuperarCampanaRecup = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function RecuperarCampanaRecupAgencias(ByVal pnCod As Long) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_CampanasRecupAgencias " & pnCod
    
    oCon.AbreConexion
    Set RecuperarCampanaRecupAgencias = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function RecuperarCampanaRecupTpoCred(ByVal pnCod As Long) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_CampanasRecupTpoCred " & pnCod
    
    oCon.AbreConexion
    Set RecuperarCampanaRecupTpoCred = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function RecuperarCampanaRecupSubCamp(ByVal pnCod As Long) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_CampanasRecupSubCamp " & pnCod
    
    oCon.AbreConexion
    Set RecuperarCampanaRecupSubCamp = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function RecuperarCampanaRecupNivelesApr(ByVal pnCod As Long) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_CampanasRecupNivelesApr " & pnCod
    
    oCon.AbreConexion
    Set RecuperarCampanaRecupNivelesApr = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function RecuperarCampanaRecupDescActivos() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_CampanasRecupConfActDescuento"
    
    oCon.AbreConexion
    Set RecuperarCampanaRecupDescActivos = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function RegistrarCampanaRecup(ByVal psNombre As String, ByVal psAprobado As String, ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal pnNumMax As Integer) As Long
Dim oRs As New ADODB.Recordset
Dim lsSQL As String

lsSQL = "EXEC stp_ins_CampanasRecup '" & Replace(psNombre, "'", "") & "','" & Replace(psAprobado, "'", "") & "','" & Format(pdDesde, "yyyymmdd") & "','" & _
        Format(pdHasta, "yyyymmdd") & "'," & pnNumMax

Set oRs = loConecta.CargaRecordSet(lsSQL)
If Not (oRs.EOF And oRs.BOF) Then
    RegistrarCampanaRecup = CLng(oRs!nId)
End If
Set oRs = Nothing
End Function

Public Sub ActualizaCampanaRecup(ByVal pnCod As Long, ByVal psNombre As String, ByVal psAprobado As String, ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal pnNumMax As Integer)
Dim lsSQL As String

lsSQL = "EXEC stp_upd_CampanasRecup " & pnCod & ",'" & Replace(psNombre, "'", "") & "','" & Replace(psAprobado, "'", "") & "','" & Format(pdDesde, "yyyymmdd") & "','" & _
        Format(pdHasta, "yyyymmdd") & "'," & pnNumMax

loConecta.Ejecutar (lsSQL)

End Sub
Public Sub EliminaAgenciasCampanaRecup(ByVal pnCod As Long)
Dim lsSQL As String

lsSQL = "EXEC stp_del_CampanasRecupAgencias " & pnCod

loConecta.Ejecutar (lsSQL)
End Sub
Public Sub EliminaTpoCredCampanaRecup(ByVal pnCod As Long)
Dim lsSQL As String

lsSQL = "EXEC stp_del_CampanasRecupTpoCred " & pnCod

loConecta.Ejecutar (lsSQL)
End Sub

Public Sub RegistraAgenciasCampanaRecup(ByVal pnCod As Long, ByVal psAge As String)
Dim lsSQL As String

lsSQL = "EXEC stp_ins_CampanasRecupAgencias " & pnCod & ",'" & psAge & "'"

loConecta.Ejecutar (lsSQL)
End Sub
Public Sub RegistraTpoCredCampanaRecup(ByVal pnCod As Long, ByVal psTpoCred As String)
Dim lsSQL As String

lsSQL = "EXEC stp_ins_CampanasRecupTpoCred " & pnCod & ",'" & psTpoCred & "'"

loConecta.Ejecutar (lsSQL)
End Sub

Public Sub DarBajaSubCampCampanaRecup(ByVal pnCod As Long, ByVal psSubCamp As String, ByVal pnEstado As Integer)
Dim lsSQL As String

lsSQL = "EXEC stp_upd_CampanasRecupSubCampDarBaja " & pnCod & ",'" & psSubCamp & "'," & pnEstado

loConecta.Ejecutar (lsSQL)
End Sub
Public Function RegistraSubCampCampanaRecup(ByVal pnCod As Long, ByVal pnDiasAtrasoIni As Integer, ByVal pnDiasAtrasoFin As Integer, _
                                        ByVal pnConsidera As Integer, ByVal pnTpoConsidera As Integer, ByVal pnDescCap As Double, ByVal pnDescInt As Double, _
                                        ByVal pnDescMora As Double, ByVal pnDescGasto As Double, ByVal pbVencidos As Integer, ByVal pbTransferidos As Integer, ByVal pnTpoGarantia As Integer, _
                                        ByVal pnDescIcv As Double) As Long 'JOEP
Dim oRs As New ADODB.Recordset
Dim lsSQL As String

lsSQL = "EXEC stp_ins_CampanasRecupSubCamp " & pnCod & "," & pnDiasAtrasoIni & "," & pnDiasAtrasoFin & "," & pnConsidera & "," & pnTpoConsidera & "," & _
        pnDescCap & "," & pnDescInt & "," & pnDescMora & "," & pnDescGasto & "," & pbVencidos & "," & pbTransferidos & "," & pnTpoGarantia & "," & _
        pnDescIcv & "" 'JOEP

Set oRs = loConecta.CargaRecordSet(lsSQL)
If Not (oRs.EOF And oRs.BOF) Then
    RegistraSubCampCampanaRecup = CLng(oRs!nIdSubCamp)
End If
Set oRs = Nothing
End Function

Public Sub ActualizaSubCampCampanaRecup(ByVal pnCod As Long, ByVal pnSubCod As Long, ByVal pnDiasAtrasoIni As Integer, ByVal pnDiasAtrasoFin As Integer, _
                                        ByVal pnConsidera As Integer, ByVal pnTpoConsidera As Integer, ByVal pnDescCap As Double, ByVal pnDescInt As Double, _
                                        ByVal pnDescMora As Double, ByVal pnDescGasto As Double, ByVal pbVencidos As Integer, ByVal pbTransferidos As Integer, ByVal pnTpoGarantia As Integer, _
                                        ByVal pnDescIcv As Double) 'JOEP
Dim lsSQL As String

lsSQL = "EXEC stp_upd_CampanasRecupSubCamp " & pnCod & "," & pnSubCod & "," & pnDiasAtrasoIni & "," & pnDiasAtrasoFin & "," & pnConsidera & "," & pnTpoConsidera & "," & _
        pnDescCap & "," & pnDescInt & "," & pnDescMora & "," & pnDescGasto & "," & pbVencidos & "," & pbTransferidos & "," & pnTpoGarantia & "," & _
        pnDescIcv & "" 'JOEP

loConecta.Ejecutar (lsSQL)
End Sub

Public Sub EliminaNivelesAprCampanaRecup(ByVal pnCod As Long)
Dim lsSQL As String

lsSQL = "EXEC stp_del_CampanasRecupNivelesApr " & pnCod

loConecta.Ejecutar (lsSQL)
End Sub
Public Sub RegistraNivelesAprCampanaRecup(ByVal pnCod As Long, ByVal pnNivel As Integer, ByVal pnDesde As Double, ByVal pnHasta As Double)
Dim lsSQL As String

lsSQL = "EXEC stp_ins_CampanasRecupNivelesApr " & pnCod & "," & pnNivel & "," & pnDesde & "," & pnHasta

loConecta.Ejecutar (lsSQL)
End Sub
Public Function CreditosHabilitadosCampRecup(ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_CreditosParaCampRecup '" & psPersCod & "'"
    oCon.AbreConexion
    Set CreditosHabilitadosCampRecup = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function RecuperarCampanaRecupActivas(ByVal pdFecha As Date, ByVal pcCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_CampanaRecupActivas '" & Format(pdFecha, "yyyymmdd") & "','" & pcCtaCod & "'"
    
    oCon.AbreConexion
    Set RecuperarCampanaRecupActivas = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function RecuperarDatosCredCampanaRecup(ByVal pcCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "EXEC stp_sel_DatosCredParaCampRecup '" & pcCtaCod & "'"
    
    oCon.AbreConexion
    Set RecuperarDatosCredCampanaRecup = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function RecuperarCampanaRecupCantVeces(ByVal pnCod As Long, ByVal pcCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "EXEC stp_sel_CampanaRecupCantVeces " & pnCod & ",'" & pcCtaCod & "'"
    
    oCon.AbreConexion
    Set RecuperarCampanaRecupCantVeces = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function RecuperarCampanaRecupSubCampCred(ByVal pnCod As Long, ByVal pcCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "EXEC stp_sel_CampanaRecupObtenerSubCampCred " & pnCod & ",'" & pcCtaCod & "'"
    
    oCon.AbreConexion
    Set RecuperarCampanaRecupSubCampCred = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function RecuperarCampanaRecupCuotasVencCred(ByVal pcCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "EXEC stp_sel_CampanaRecupCuotasVencidas '" & pcCtaCod & "'"
    
    oCon.AbreConexion
    Set RecuperarCampanaRecupCuotasVencCred = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function AcogerCredCampanaRecup(ByVal pcCtaCod As String, ByVal pnSaldoCap As Double, ByVal pnDiasAtraso As Integer, ByVal pnCuotaVenc As Integer, ByVal pnNroCalen As Integer, _
                                          ByVal pnCampana As Long, ByVal pnSubCamp As Long, ByVal pnMontoPerdon As Double, ByVal pnMontoPagar As Double, ByVal pnMontoMin As Double, _
                                          ByVal pnMontoMax As Double, ByVal pnNivelApr As Integer, ByVal pnMontoPend As Double, ByVal pnMontoDeuda As Double, ByVal pnPorcDescCap As Double, _
                                          ByVal pnPorcDescInt As Double, ByVal pnPorcDescMora As Double, ByVal pnPorcDescGasto As Double, ByVal psMovNro As String, _
                                          ByVal pnPorcDesIcv As Double) As Long 'JOEP
Dim oRs As New ADODB.Recordset
Dim lsSQL As String

lsSQL = "EXEC stp_ins_CampanasRecupCred '" & pcCtaCod & "'," & pnSaldoCap & "," & pnDiasAtraso & "," & pnCuotaVenc & "," & pnNroCalen & "," & pnCampana & "," & pnSubCamp & "," & _
                                            pnMontoPerdon & "," & pnMontoPagar & "," & pnMontoMin & "," & pnMontoMax & "," & pnNivelApr & "," & pnMontoPend & "," & pnMontoDeuda & "," & _
                                            pnPorcDescCap & "," & pnPorcDescInt & "," & pnPorcDescMora & "," & pnPorcDescGasto & ",'" & psMovNro & "'," & _
                                            pnPorcDesIcv 'JOEP

Set oRs = loConecta.CargaRecordSet(lsSQL)
If Not (oRs.EOF And oRs.BOF) Then
    AcogerCredCampanaRecup = CLng(oRs!nId)
End If
Set oRs = Nothing
End Function
Public Function RecuperarCampanaRecupNivelApr(ByVal pnCod As Long, ByVal pnMonto As Double) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "EXEC stp_sel_CampanasRecupNivelApr " & pnCod & "," & pnMonto
    
    oCon.AbreConexion
    Set RecuperarCampanaRecupNivelApr = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Sub AcogerCredCampanaRecupCuota(ByVal pnCod As Long, ByVal pnCuota As Integer, ByVal nCap As Double, ByVal nCapPag As Double, ByVal nCapPerd As Double, ByVal nInt As Double, _
                                        ByVal nIntPag As Double, ByVal nIntPerd As Double, ByVal nMora As Double, ByVal nMoraPag As Double, ByVal nMoraPerd As Double, _
                                        ByVal nGasto As Double, ByVal nGastoPag As Double, ByVal nGastoPerd As Double, ByVal nIntGraPend As Double, _
                                        ByVal nIcv As Double, ByVal nIcvPag As Double, ByVal nIcvPerd As Double) 'JOEP
Dim lsSQL As String

lsSQL = "EXEC stp_ins_CampanasRecupCredDet " & pnCod & "," & pnCuota & "," & nCap & "," & nCapPag & "," & nCapPerd & "," & nInt & "," & nIntPag & "," & nIntPerd & "," & nMora & "," & _
                                            nMoraPag & "," & nMoraPerd & "," & nGasto & "," & nGastoPag & "," & nGastoPerd & "," & nIntGraPend & _
                                            "," & nIcv & "," & nIcvPag & "," & nIcvPerd 'JOEP

loConecta.Ejecutar (lsSQL)
End Sub

Public Sub AnularCampanaRecupCred(ByVal pdFecha As Date, ByVal pcCtaCod As String)
Dim lsSQL As String

lsSQL = "EXEC stp_upd_CampanasRecupCredAnularSol '" & Format(pdFecha, "yyyymmdd") & "','" & pcCtaCod & "'"

loConecta.Ejecutar (lsSQL)
End Sub

Public Sub AprocacionDirectaCampanaRecupCred(ByVal pnCod As Integer, ByVal pnNivel As Integer, ByVal pcRHCargoCod As String, ByVal pcMovNroApr As String)
Dim lsSQL As String

lsSQL = "EXEC stp_ins_CampanasRecupAprobacionDirecta " & pnCod & "," & pnNivel & ",'" & pcRHCargoCod & "','" & pcMovNroApr & "'"

loConecta.Ejecutar (lsSQL)
End Sub
Public Function RecuperarCampanaRecupCreditosAAprobar(ByVal pdFecha As Date, ByVal psCargo As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "EXEC stp_sel_CampanasRecupCredParaAprobar '" & Format(pdFecha, "yyyymmdd") & "','" & psCargo & "'"
    
    oCon.AbreConexion
    Set RecuperarCampanaRecupCreditosAAprobar = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Sub RegistraAutorizacionCredCampanaRecup(ByVal pnCod As Long, ByVal psCodCargo As String, ByVal psMovNro As String)
Dim lsSQL As String

lsSQL = "EXEC stp_ins_CampanasRecupAprobaciones " & pnCod & ",'" & psCodCargo & "','" & psMovNro & "'"

loConecta.Ejecutar (lsSQL)
End Sub

Public Sub ActualizaEstCredCampanaRecup(ByVal pnCod As Long)
Dim lsSQL As String

lsSQL = "EXEC stp_upd_CampanasRecupCredActEstado " & pnCod

loConecta.Ejecutar (lsSQL)
End Sub
Public Function RecuperarCampanaRecupXCredEstadoFecha(ByVal pcCtaCod As String, ByVal pnEstado As Integer, ByVal pdFecha As Date) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "EXEC stp_sel_CampanasRecupXCredEstadoFecha '" & pcCtaCod & "'," & pnEstado & ",'" & Format(pdFecha, "yyyymmdd") & "'"
    
    oCon.AbreConexion
    Set RecuperarCampanaRecupXCredEstadoFecha = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function RecuperarCampanaRecupDetXID(ByVal pnCod As Long) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "EXEC stp_sel_CampanasRecupDetXID " & pnCod
    
    oCon.AbreConexion
    Set RecuperarCampanaRecupDetXID = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Sub CampanasRecupPagoConPerdon(ByVal pnId As Double, ByVal pcMovNro As String, ByVal pnMovNro As Double, ByVal pnEstado As Integer)
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta
On Error GoTo ErrorCampanasRecupPagoConPerdon

sSql = "EXEC stp_upd_CampanasRecupPagoConPerdon " & pnId & ",'" & pcMovNro & "'," & pnMovNro & "," & pnEstado

oCon.AbreConexion
oCon.Ejecutar sSql
oCon.CierraConexion
Set oCon = Nothing
    
Exit Sub
ErrorCampanasRecupPagoConPerdon:
        Err.Raise Err.Number, "DCOMCredito:CampanasRecupPagoConPerdon", Err.Description
End Sub
'WIOR FIN ********
'FRHU 20150415 ERS022-2015
Public Function VerificaSiEsCreditoTransferido(ByVal psCtaCod As String) As Boolean
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "stp_sel_VerificarSiEsCreditoTransferido '" & psCtaCod & "'"

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    If Not (rs.EOF And rs.BOF) Then
        VerificaSiEsCreditoTransferido = True
    Else
        VerificaSiEsCreditoTransferido = False
    End If
    
    Set rs = Nothing
End Function
'FIN FRHU 20150415

'WIOR 20150421 ***
Public Function ObtenerInteresDiferidoCredito(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "Exec stp_sel_ObtenerInteresDiferido '" & psCtaCod & "'"
    oCon.AbreConexion
    Set ObtenerInteresDiferidoCredito = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ObtenerInteresDiferidoPendCredito(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "Exec stp_sel_ObtenerIntDiferidoPend '" & psCtaCod & "'"
    oCon.AbreConexion
    Set ObtenerInteresDiferidoPendCredito = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'WIOR FIN ********
'ALPA 20150730****
Public Function RecuperaPorcentajeDiferidos(ByVal psCtaCod As String) As Currency
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim lrs As ADODB.Recordset
Set lrs = New ADODB.Recordset
    On Error GoTo ErrorRecuperaPorcentajeDiferidos
    RecuperaPorcentajeDiferidos = 0
    sSql = "exec stp_sel_ObtenerPorcentajeDiferidos '" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set lrs = oConecta.CargaRecordSet(sSql)
    If Not (lrs.EOF Or lrs.BOF) Then
    RecuperaPorcentajeDiferidos = lrs!nMontoPorcentaje
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorRecuperaPorcentajeDiferidos:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function
'*****************
'RECO20151019 ERS034-2015***************************************
Public Sub RegistraPagoActivacionSeguro(ByVal pnMovNro As Long, ByVal psCtaCod As String, ByVal pnTpoSeguro As Integer)
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    On Error GoTo ErrorRegistraPagoActivacionSeguro
    
    sSql = "EXEC stp_ins_RegistraPagoActivacionSeguro " & pnMovNro & ",'" & psCtaCod & "'," & pnTpoSeguro
    
    oCon.AbreConexion
    oCon.Ejecutar sSql
    oCon.CierraConexion
    Set oCon = Nothing
        
    Exit Sub
ErrorRegistraPagoActivacionSeguro:
            Err.Raise Err.Number, "DCOMCredito:RegistraPagoActivacionSeguro", Err.Description
End Sub
'RECO FIN ******************************************************

'WIOR 20151220 ***
Public Function EsCredMIVIVENDA(ByVal cTpoProd As String, ByVal cTpoCred As String, Optional ByVal pnTipo As Integer = 1) As Boolean
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset

sSql = "EXEC stp_sel_EsCreditoMIVIVIENDA '" & cTpoProd & "','" & cTpoCred & "'," & pnTipo

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing

If Not (rs.EOF And rs.BOF) Then
    If CBool(rs!Resultado) Then
        EsCredMIVIVENDA = True
    Else
        EsCredMIVIVENDA = False
    End If
Else
    EsCredMIVIVENDA = False
End If

Set rs = Nothing
End Function

Public Function ObtenerCredAlertaMIVIVIENDACred(ByVal psCtaCod As String, ByVal pnMonto As Double) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "EXEC stp_sel_CredMiViviendaAlertasPago '" & psCtaCod & "'," & pnMonto
    oCon.AbreConexion
    Set ObtenerCredAlertaMIVIVIENDACred = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ObtenerCredAlertaMIVIVIENDA() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "EXEC stp_sel_ObtenerPendCredMiViviendaAlertasPago"
    oCon.AbreConexion
    Set ObtenerCredAlertaMIVIVIENDA = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Sub RegistroMiViviendaAlertasPago(ByVal psCtaCod As String, ByVal pnMonto As Double, ByVal psMovNroReg As String, ByVal MatDatVivienda As Variant)
Dim lsSQL As String

lsSQL = "EXEC stp_ins_CredMiViviendaAlertasPago '" & psCtaCod & "'," & pnMonto & ",'" & psMovNroReg & "','" & MatDatVivienda(0) & "','" & MatDatVivienda(1) & "'"

loConecta.Ejecutar (lsSQL)
End Sub

Public Sub ActualizaMiViviendaAlertasPago(ByVal pnId As Long, ByVal psMovNroApr As String, ByVal pnEstado As Integer)
Dim lsSQL As String

lsSQL = "EXEC stp_upd_CredMiViviendaAlertasPago " & pnId & ",'" & psMovNroApr & "'," & pnEstado
    
loConecta.Ejecutar (lsSQL)
End Sub

Public Sub EliminarMiViviendaAlertasPago(ByVal psCtaCod As String, ByVal psMovNroApr As String)
Dim lsSQL As String

lsSQL = "EXEC stp_upd_EliminarCredMiViviendaAlertasPago '" & psCtaCod & "','" & psMovNroApr & "'"
    
loConecta.Ejecutar (lsSQL)
End Sub

Public Function ObtieneConfigMonedaTpoProd(ByVal psTpoProd As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "EXEC stp_sel_CredConfigTpoProdMoneda '" & psTpoProd & "'"
    oCon.AbreConexion
    Set ObtieneConfigMonedaTpoProd = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'CTI50012020****************************************************************************
Public Function ObtenerValoresNuevoMIVIVIENDA(ByVal psMoneda As String, ByVal pdFecha As Date, ByVal pnMontoVivienda As Double, Optional ByVal pnCuotaInicial As Double = 0) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta

    sSql = "EXEC stp_sel_ObtenerValoresCredMIVIVIENDA '" & psMoneda & "','" & Format(pdFecha, "yyyymmdd") & "'," & pnMontoVivienda & "," & pnCuotaInicial

    oCon.AbreConexion
    Set ObtenerValoresNuevoMIVIVIENDA = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function ObtenerValoresNuevoMIVIVIENDA_NEW(ByVal psMoneda As String, ByVal pdFecha As Date, ByVal pnMontoVivienda As Double, Optional ByVal pnCuotaInicial As Double = 0, Optional ByVal psTpoProdCod As String = "854", Optional ByVal pnDestino As Integer = 0, Optional ByVal pnGastoCierre As Double = 0) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta

    sSql = "EXEC stp_sel_ObtenerValoresCredMIVIVIENDA_NEW '" & psMoneda & "','" & Format(pdFecha, "yyyymmdd") & "'," & pnMontoVivienda & "," & pnCuotaInicial & ",'" & psTpoProdCod & "'," & pnDestino & "," & pnGastoCierre
    
    oCon.AbreConexion
    Set ObtenerValoresNuevoMIVIVIENDA_NEW = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'***************************************************************************************

Public Function ObtenerDatosNuevoMIVIVIENDA(ByVal psCtaCod As String, ByVal pnEstado As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta

    sSql = "EXEC stp_sel_CredMiVivienda '" & psCtaCod & "'," & pnEstado
    
    oCon.AbreConexion
    Set ObtenerDatosNuevoMIVIVIENDA = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Sub EliminarDatosNuevoMIVIVIENDA(ByVal psCtaCod As String, ByVal pnEstado As Integer)
Dim lsSQL As String

lsSQL = "EXEC stp_del_CredMiVivienda '" & psCtaCod & "'," & pnEstado

loConecta.Ejecutar (lsSQL)
End Sub

'CTI5 ER001-2020*************************
'Public Sub RegistrarDatosNuevoMIVIVIENDA(ByVal psCtaCod As String, ByVal pnPrdEstado As Integer, ByVal pnMontoVivienda As Double, _
'                                        ByVal pnCuotaInicial As Double, ByVal pnBonoOtorgado As Double, ByVal pnMontoCred As Double, _
'                                        ByVal pnUIT As Double, ByVal pnDesde As Integer, ByVal pnHasta As Integer, _
'                                        ByVal pnBono As Double, ByVal pnMinCredUIT As Double, ByVal pnPeriodoPerdBono As Integer, ByVal pnGastosCierre As Double)
Public Sub RegistrarDatosNuevoMIVIVIENDA(ByVal psCtaCod As String, ByVal pnPrdEstado As Integer, ByVal pnMontoVivienda As Double, _
                                        ByVal pnCuotaInicial As Double, ByVal pnBonoOtorgado As Double, ByVal pnMontoCred As Double, _
                                        ByVal pnUIT As Double, ByVal pnDesde As Long, ByVal pnHasta As Long, _
                                        ByVal pnBono As Double, ByVal pnMinCredUIT As Double, ByVal pnPeriodoPerdBono As Integer, Optional ByVal pnGastosCierre As Double = 0#)

Dim lsSQL As String

lsSQL = "EXEC stp_ins_CredMiVivienda '" & psCtaCod & "'," & pnPrdEstado & "," & pnMontoVivienda & "," & pnCuotaInicial & _
                                        "," & pnBonoOtorgado & "," & pnMontoCred & "," & pnUIT & "," & pnDesde & "," & pnHasta & _
                                        "," & pnBono & "," & pnMinCredUIT & "," & pnPeriodoPerdBono & "," & pnGastosCierre

loConecta.Ejecutar (lsSQL)
End Sub

'******************************

Public Function CreditosMIVIVIENDAPersona(ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_CreditosMIVIVIENDAPers '" & psPersCod & "'"
    oCon.AbreConexion
    Set CreditosMIVIVIENDAPersona = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'WIOR FIN ********
'JUEZ 20160216 **********************************************************
Public Function ObtienePromedioDiasAtraso(ByVal psCtaCod As String, ByVal pdFecSis As Date, ByVal nCantMesesAnalizar As Integer) As Integer
Dim oCon As New COMConecta.DCOMConecta
Dim sSql As String
Dim rs As ADODB.Recordset

    sSql = "Exec stp_sel_PromedioDiasAtrasoCredito '" & psCtaCod & "','" & Format(pdFecSis, "yyyyMMdd") & "'," & nCantMesesAnalizar
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    If Not rs.EOF And Not rs.BOF Then
        ObtienePromedioDiasAtraso = rs!nPromedioDias
    Else
        ObtienePromedioDiasAtraso = 0
    End If
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'Public Function RecuperaColocacReprogramado(Optional ByVal psCtaCod As String = "", Optional ByVal psEstados As String = "") As ADODB.Recordset
Public Function RecuperaColocacReprogramado(Optional ByVal psCtaCod As String = "", Optional ByVal psEstados As String = "", Optional ByVal pnOpcion As Integer = -1) As ADODB.Recordset 'Add JOEP20210306 garantia covid
'Trabaja con el ultimo estado del crédito
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    'sSql = "Exec stp_sel_ColocacReprogramado '" & psCtaCod & "','" & psEstados & "'"
    sSql = "Exec stp_sel_ColocacReprogramado '" & psCtaCod & "','" & psEstados & "'," & pnOpcion & "" 'Add JOEP20210306 garantia covid
    oCon.AbreConexion
    Set RecuperaColocacReprogramado = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function RecuperaColocacReprogramadoEstado(ByVal psCtaCod As String, ByVal pnEstado As ColocacReprogEstado) As ADODB.Recordset
'Trabaja con el estado que se envia
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    sSql = "Exec stp_sel_ColocacReprogramadoEstado '" & psCtaCod & "'," & pnEstado
    oCon.AbreConexion
    Set RecuperaColocacReprogramadoEstado = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function RecuperaDatosPropuestaReprogramacion(ByVal psCtaCod As String, ByVal pdFecha As Date) As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    sSql = "Exec stp_sel_RecuperaDatosPropuestaReprogramacion '" & psCtaCod & "','" & Format(pdFecha, "yyyyMMdd") & "'"
    oCon.AbreConexion
    Set RecuperaDatosPropuestaReprogramacion = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'Public Function RecuperaColocacReprogramadoPorAprobar(ByVal pnEstado As ColocacReprogEstado) As ADODB.Recordset 'Comento JOEP20171222 Acta 220-2017
Public Function RecuperaColocacReprogramadoPorAprobar(ByVal pnEstado As ColocacReprogEstado, Optional ByVal pnTpPermiso As Integer = 0) As ADODB.Recordset 'Agregar JOEP20171222 Acta 220-2017
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    'sSql = "Exec stp_sel_ColocacReprogramadoPorAprobar " & pnEstado'comento JOEP20171222 Acta 220-2017
    sSql = "Exec stp_sel_ColocacReprogramadoPorAprobar " & pnEstado & "," & pnTpPermiso & "" 'Agregar JOEP20171222 Acta 220-2017
    oCon.AbreConexion
    Set RecuperaColocacReprogramadoPorAprobar = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

'Public Function RecuperaColocReprogExonera() As ADODB.Recordset 'Comento JOEP20171222 ACTA 220-2017
Public Function RecuperaColocReprogExonera(ByVal npTipoOp As Integer) As ADODB.Recordset 'JOEP20171222 ACTA 220-2017
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    'sSql = "Exec stp_sel_ColocReprogExonera" 'Comento JOEP20171222 ACTA 220-2017
    sSql = "Exec stp_sel_ColocReprogExonera " & npTipoOp & "" 'JOEP20171222 ACTA 220-2017
    oCon.AbreConexion
    Set RecuperaColocReprogExonera = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function RecuperaColocacReprogExoneraSolicitud(ByVal psCtaCod As String, ByVal pdFecha As Date) As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    sSql = "Exec stp_sel_ColocacReprogExoneraSolicitud '" & psCtaCod & "','" & Format(pdFecha, "yyyyMMdd") & "'"
    oCon.AbreConexion
    Set RecuperaColocacReprogExoneraSolicitud = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'END JUEZ ***************************************************************

'WIOR 20160317 ***
Public Function GestionSiniestroCredCliente(ByVal psPersCod As String, ByVal pdFecha As Date) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "EXEC stp_sel_GestionSiniestroCredCliente '" & psPersCod & "','" & Format(pdFecha, "yyyymmdd") & "'"
    oCon.AbreConexion
    Set GestionSiniestroCredCliente = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function GestionSiniestroReg(ByVal pnTipoSeguro As Long, ByVal pcPersCodCiaSeguro As String, ByVal pcPersCodCliente As String, ByVal pdFechaSiniestro As Date, _
                                          ByVal pnMoneda As Integer, ByVal pnSumaSol As Currency, ByVal pcNumCarta As String, ByVal pcGlosa As String, ByVal pdFechaReg As Date, ByVal pcUserReg As String) As Long
Dim oRs As New ADODB.Recordset
Dim lsSQL As String

lsSQL = "EXEC stp_ins_GestionSiniestro " & pnTipoSeguro & ",'" & pcPersCodCiaSeguro & "','" & pcPersCodCliente & "','" & Format(pdFechaSiniestro, "yyyymmdd") & "'," & pnMoneda & "," & pnSumaSol & ",'" & _
                                        Replace(pcNumCarta, "'", "") & "','" & Replace(pcGlosa, "'", "") & "','" & Format(pdFechaReg + loConecta.GetHoraServer, "yyyymmdd hh:mm:ss") & "','" & pcUserReg & "'"

Set oRs = loConecta.CargaRecordSet(lsSQL)
If Not (oRs.EOF And oRs.BOF) Then
    GestionSiniestroReg = CLng(oRs!nIdSiniestro)
End If
Set oRs = Nothing
End Function

Public Sub GestionSiniestroRegDocsAdj(ByVal pnIdSiniestro As Long, ByVal pnCodDocAdj As Long, ByVal pnTipo As Integer)
Dim lsSQL As String

lsSQL = "EXEC stp_ins_GestionSiniestroDocsAdj " & pnIdSiniestro & "," & pnCodDocAdj & "," & pnTipo

loConecta.Ejecutar (lsSQL)
End Sub

Public Sub GestionSiniestroRegCreditos(ByVal pnIdSiniestro As Long, ByVal pdFecha As Date, ByVal pcCtaCod As String, ByVal pnCapital As Currency, ByVal pnIntComp As Currency, ByVal pnIntMora As Currency, ByVal pnGastos As Currency)
Dim lsSQL As String

lsSQL = "EXEC stp_ins_GestionSiniestroCreditos " & pnIdSiniestro & ",'" & Format(pdFecha, "yyyymmdd") & "','" & pcCtaCod & "'," & pnCapital & "," & pnIntComp & "," & pnIntMora & "," & pnGastos

loConecta.Ejecutar (lsSQL)
End Sub

Public Function GestionSiniestroEnProceso() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "EXEC stp_sel_GestionSiniestroEnProceso"
    oCon.AbreConexion
    Set GestionSiniestroEnProceso = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function GestionSiniestroRechazados(ByVal pbCliente As Boolean, ByVal pcPersCod As String, ByVal pbFecha As Boolean, ByVal pdFechaA As Date, ByVal pdFechaB As Date) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "EXEC stp_sel_GestionSiniestroRechazados " & IIf(pbCliente, 1, 0) & ",'" & pcPersCod & "'," & IIf(pbFecha, 1, 0) & ",'" & Format(pdFechaA, "yyyymmdd") & "','" & Format(pdFechaB, "yyyymmdd") & "'"
    oCon.AbreConexion
    Set GestionSiniestroRechazados = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function GestionSiniestroLiquidados(ByVal pbCliente As Boolean, ByVal pcPersCod As String, ByVal pbFecha As Boolean, ByVal pdFechaA As Date, ByVal pdFechaB As Date) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "EXEC stp_sel_GestionSiniestroLiquidados " & IIf(pbCliente, 1, 0) & ",'" & pcPersCod & "'," & IIf(pbFecha, 1, 0) & ",'" & Format(pdFechaA, "yyyymmdd") & "','" & Format(pdFechaB, "yyyymmdd") & "'"
    oCon.AbreConexion
    Set GestionSiniestroLiquidados = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function GestionSiniestroDetalle(ByVal pnId As Long) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "EXEC stp_sel_GestionSiniestroDetalle " & pnId
    oCon.AbreConexion
    Set GestionSiniestroDetalle = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function GestionSiniestroDetalleDocs(ByVal pnId As Long) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "EXEC stp_sel_GestionSiniestroDetalleDocs " & pnId
    oCon.AbreConexion
    Set GestionSiniestroDetalleDocs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function GestionSiniestroDetalleRegAten(ByVal pnId As Long) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "EXEC stp_sel_GestionSiniestroDetalleRegAten " & pnId
    oCon.AbreConexion
    Set GestionSiniestroDetalleRegAten = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function GestionSiniestroDetalleCred(ByVal pnId As Long, ByVal pdFecha As Date) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "EXEC stp_sel_GestionSiniestroDetalleCred " & pnId & ",'" & Format(pdFecha, "yyyymmdd") & "'"
    oCon.AbreConexion
    Set GestionSiniestroDetalleCred = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function GestionSiniestroCred(ByVal pnId As Long) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "EXEC stp_sel_GestionSiniestroCred " & pnId
    oCon.AbreConexion
    Set GestionSiniestroCred = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function GestionSiniestroRegAtencion(ByVal pnIdSiniestro As Long, ByVal pnTipo As Integer, ByVal pnImporte As Currency, ByVal pnMoneda As Integer, _
                                            ByVal pcDocReferencia As String, ByVal pdFechaRecepcion As Date, ByVal pcAgeCod As String, ByVal pcDocEmitir As String, _
                                            ByVal pcObservacion As String, ByVal pdFechaReg As Date, ByVal pcUserReg As String) As Long
Dim oRs As New ADODB.Recordset
Dim lsSQL As String

lsSQL = "EXEC stp_ins_GestionSiniestroRegAtencion " & pnIdSiniestro & "," & pnTipo & "," & pnImporte & "," & pnMoneda & ",'" & Replace(Replace(pcDocReferencia, "'", ""), vbCrLf, "") & "','" & _
                                        Format(pdFechaRecepcion, "yyyymmdd") & "','" & pcAgeCod & "','" & Replace(Replace(pcDocEmitir, "'", ""), vbCrLf, "") & "','" & Replace(Replace(pcObservacion, "'", ""), vbCrLf, "") & "','" & _
                                        Format(pdFechaReg + loConecta.GetHoraServer, "yyyymmdd hh:mm:ss") & "','" & pcUserReg & "'"

Set oRs = loConecta.CargaRecordSet(lsSQL)
If Not (oRs.EOF And oRs.BOF) Then
    GestionSiniestroRegAtencion = CLng(oRs!nIdRegAtencion)
End If
Set oRs = Nothing
End Function

Public Sub GestionSiniestroRegAtencionDistMontos(ByVal pnIdRegAtencion As Long, ByVal pcCtaCod As String, ByVal pnMontoDevolver As Currency, ByVal pnCapital As Currency, _
                                                ByVal pnIntComp As Currency, ByVal pnIntMora As Currency, ByVal pnGastos As Currency, ByVal pdFechaReg As Date, ByVal pcUserReg As String)
Dim lsSQL As String

lsSQL = "EXEC stp_ins_GestionSiniestroDistMontos " & pnIdRegAtencion & ",'" & pcCtaCod & "'," & pnMontoDevolver & "," & pnCapital & "," & pnIntComp & "," & pnIntMora & "," & _
        pnGastos & ",'" & Format(pdFechaReg + loConecta.GetHoraServer, "yyyymmdd hh:mm:ss") & "','" & pcUserReg & "'"


loConecta.Ejecutar (lsSQL)
End Sub

Public Sub GestionSiniestroRegAtencionDistMontosActualizar(ByVal pnIdRegAtencion As Long, ByVal pcCtaCod As String, ByVal pnMontoDevolver As Currency, ByVal pnCapital As Currency, _
                                                ByVal pnIntComp As Currency, ByVal pnIntMora As Currency, ByVal pnGastos As Currency, ByVal pdFechaReg As Date, ByVal pcUserReg As String)
Dim lsSQL As String

lsSQL = "EXEC stp_upd_GestionSiniestroDistMontos " & pnIdRegAtencion & ",'" & pcCtaCod & "'," & pnMontoDevolver & "," & pnCapital & "," & pnIntComp & "," & pnIntMora & "," & _
        pnGastos & ",'" & Format(pdFechaReg + loConecta.GetHoraServer, "yyyymmdd hh:mm:ss") & "','" & pcUserReg & "'"


loConecta.Ejecutar (lsSQL)
End Sub

Public Function GestionSiniestroRegAtencionID(ByVal pnId As Long) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "EXEC stp_sel_GestionSiniestroRegAtencionId " & pnId
    oCon.AbreConexion
    Set GestionSiniestroRegAtencionID = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function GestionSiniestroDistMontos(ByVal pnIdRegAtencion As Long, ByVal pcCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "EXEC stp_sel_GestionSiniestroDistMontos " & pnIdRegAtencion & ",'" & pcCtaCod & "'"
    oCon.AbreConexion
    Set GestionSiniestroDistMontos = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function GestionSiniestroDistMontosFecha(ByVal pnIdSiniestro As Long, ByVal pcCtaCod As String, ByVal pdFecha As Date) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "EXEC stp_sel_GestionSiniestroDistMontosFecha " & pnIdSiniestro & ",'" & pcCtaCod & "','" & Format(pdFecha, "yyyymmdd") & "'"
     
    oCon.AbreConexion
    Set GestionSiniestroDistMontosFecha = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function CreditosALiquidarSegDesgravamen(ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_CreditosALiquidarSegDesgravamen '" & psPersCod & "'"
    oCon.AbreConexion
    Set CreditosALiquidarSegDesgravamen = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ObtenerDatosALiquidarCred(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_ObtenerDatosALiquidarCred '" & psCtaCod & "'"
    oCon.AbreConexion
    Set ObtenerDatosALiquidarCred = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Sub GestionSiniestroActualizaEstado(ByVal pnId As Long, ByVal pnEstado As Integer)
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "EXEC stp_upd_GestionSiniestroEstado " & pnId & "," & pnEstado
    oCon.AbreConexion
    oCon.Ejecutar sSql
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'WIOR FIN ********
'JUEZ 20160509 **************************************************************************
Public Function ValidaCreditoAmpliar(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_ValidaCreditoAmpliar '" & psCtaCod & "'"
    oCon.AbreConexion
    Set ValidaCreditoAmpliar = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Sub RegistraSolicitudAutorizacionAmpliacion(ByVal psCtaCod As String, ByVal pnMontoSol As Double, ByVal pnSaldoCapAmp As Double, _
                                                   ByVal psComent As String, ByVal psMovNro As String)
Dim lsSQL As String

lsSQL = "EXEC stp_ins_ColocacSolicAutAmp '" & psCtaCod & "'," & pnMontoSol & "," & pnSaldoCapAmp & ",'" & psComent & "','" & psMovNro & "'"

loConecta.Ejecutar (lsSQL)
End Sub
Public Function RecuperaSolicitudAutorizacionAmpliacion(Optional psCtaCod As String = "", Optional ByVal pnEstado As Integer = -1) As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_ColocacSolicAutAmp '" & psCtaCod & "'," & pnEstado
    oCon.AbreConexion
    Set RecuperaSolicitudAutorizacionAmpliacion = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Sub RegistraResultadoSolicitudAutorizacionAmpliacion(ByVal psCtaCod As String, ByVal pnResultado As Integer, ByVal psGlosa As String, ByVal psMovNro As String)
Dim lsSQL As String

lsSQL = "EXEC stp_upd_ColocacSolicAutAmp '" & psCtaCod & "'," & pnResultado & ",'" & psGlosa & "','" & psMovNro & "'"

loConecta.Ejecutar (lsSQL)
End Sub
'END JUEZ *******************************************************************************
'EJVG20160511 ***
Public Function CadenaPropietarioPendientexVincularRiesgo(ByVal psCtaCod As String) As String
    Dim oConec As New COMConecta.DCOMConecta
    Dim oRs As New ADODB.Recordset
    Dim sSql As String
    On Error GoTo ErrorMsg
    
    oConec.AbreConexion
    sSql = "EXEC stp_sel_ERS0282016_CadenaPropietarioPendientexVincularRiesgo '" & psCtaCod & "'"
    Set oRs = oConec.CargaRecordSet(sSql)
    If Not oRs.EOF Then
        CadenaPropietarioPendientexVincularRiesgo = oRs!cCadena
    End If
    oRs.Close
    Set oRs = Nothing
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso <RecuperaPropietarioVinculadoRiesgoPend>", Err.Description
End Function
Public Function ObtenerInformeRiesgoNEW(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta

On Error GoTo ErrObtenerInformeRiesgoNEW

    sSql = "EXEC stp_sel_InformeRiesgoNEW '" & psCtaCod & "'"

    oConecta.AbreConexion
    Set ObtenerInformeRiesgoNEW = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrObtenerInformeRiesgoNEW:
    Err.Raise Err.Number, "Error En Proceso <<ObtenerInformeRiesgoNEW>>", Err.Description
End Function
Public Function InsertarInformeRiesgo(ByVal psCtaCod As String, _
                            ByVal psMovRegInforme As String, _
                            ByVal pnRiesgo As Integer, _
                            ByVal pnExposicion As Double, _
                            ByVal pnProcesoGenera As Integer) As Long
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    Dim oRs As New ADODB.Recordset
    On Error GoTo ErrInsertarInformeRiesgo

    sSql = "EXEC stp_ins_InformeRiesgoNEW '" & psCtaCod & "','" & psMovRegInforme & "'," & pnRiesgo & "," & pnExposicion & "," & pnProcesoGenera

    oConecta.AbreConexion
    Set oRs = oConecta.CargaRecordSet(sSql)
    If Not oRs.EOF Then
        InsertarInformeRiesgo = oRs!nInformeID
    End If
    oRs.Close
    Set oRs = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrInsertarInformeRiesgo:
    Err.Raise Err.Number, "Error En Proceso <InsertarInformeRiesgo>", Err.Description
End Function
Public Sub ActualizarInformeRiesgo(ByVal psCtaCod As String, _
                            ByVal pnInformeID As Long, _
                            Optional ByVal pnEstado As Integer = 0, _
                            Optional ByVal pnRiesgo As Integer = 0, _
                            Optional ByVal pnNivel As Integer = 0, _
                            Optional ByVal pnExposicion As Double = 0#, _
                            Optional ByVal psMovUltActualizacion As String = "", _
                            Optional ByVal psMovIngExp As String = "", _
                            Optional ByVal psMovSalExp As String = "", _
                            Optional ByVal psGlosa As String = "", _
                            Optional ByVal psGlosa2 As String = "")
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrorOpeInformeRiesgo

    sSql = "EXEC stp_upd_InformeRiesgoNEW '" & psCtaCod & "'," & pnInformeID & "," & pnEstado & "," & pnRiesgo & "," & pnNivel & "," & pnExposicion & ",'" & psMovUltActualizacion & "','" & psMovIngExp & "','" & psMovSalExp & "','" & psGlosa & "','" & psGlosa2 & "'"

    oConecta.AbreConexion
    oConecta.CargaRecordSet (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorOpeInformeRiesgo:
    Err.Raise Err.Number, "Error En Proceso <ActualizarInformeRiesgo>", Err.Description
End Sub
Public Sub EliminarInformeRiesgo(ByVal psCtaCod As String, ByVal pnInformeID As Long)
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrEliminarInformeRiesgo

    sSql = "EXEC stp_del_InformeRiesgoNEW '" & psCtaCod & "'," & pnInformeID

    oConecta.AbreConexion
    oConecta.CargaRecordSet (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrEliminarInformeRiesgo:
    Err.Raise Err.Number, "Error En Proceso <EliminarInformeRiesgo>", Err.Description
End Sub
Public Function RecuperaColocacionesXInformeRiesgos(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrRecuperaColocacionesXInformeRiesgos

    sSql = "EXEC stp_sel_ColocacionesXInformeRiesgo '" & psCtaCod & "'"

    oConecta.AbreConexion
    Set RecuperaColocacionesXInformeRiesgos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrRecuperaColocacionesXInformeRiesgos:
    Err.Raise Err.Number, "Error En Proceso <RecuperaColocacionesXInformeRiesgos>", Err.Description
End Function
Public Function CadenaDescripcionInOutCredito(ByVal psListaCod As String, ByVal pnTipo As Integer) As String
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    Dim oRs As New ADODB.Recordset
    On Error GoTo ErrCadenaDescripcionInOutCredito

    sSql = "EXEC stp_sel_ERS0282016_CadenaDescripcionInOutCredito '" & psListaCod & "'," & pnTipo

    oConecta.AbreConexion
    Set oRs = oConecta.CargaRecordSet(sSql)
    If Not oRs.EOF Then
        CadenaDescripcionInOutCredito = oRs!cListaDescripcion
    End If
    oRs.Close
    Set oRs = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrCadenaDescripcionInOutCredito:
    Err.Raise Err.Number, "Error En Proceso <CadenaDescripcionInOutCredito>", Err.Description
End Function
Public Function ObtenerMontoExposicionRiesgoUnico(ByVal psCtaCod As String, Optional ByVal pnMontoCol As Double = 0#) As Double
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    Dim oRs As ADODB.Recordset

On Error GoTo ErrObtenerInformeRiesgoNEW

    sSql = "SELECT nExposicionRU = dbo.fnc_ObtenerExposicionxRiesgoUnico('" & psCtaCod & "'," & pnMontoCol & ")"

    oConecta.AbreConexion
    Set oRs = oConecta.CargaRecordSet(sSql)
    If Not oRs.EOF Then
        ObtenerMontoExposicionRiesgoUnico = oRs!nExposicionRU
    End If
    oRs.Close
    Set oRs = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrObtenerInformeRiesgoNEW:
    Err.Raise Err.Number, "Error En Proceso <<ObtenerMontoExposicionRiesgoUnico>>", Err.Description
End Function
Public Function CadenaPermiteModificarInterviniente(ByVal psCtaCod As String) As String
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    Dim oRs As New ADODB.Recordset
    On Error GoTo ErrCadenaPermiteModificarInterviniente

    sSql = "EXEC stp_sel_ERS00282016_CadenaPermiteModificarInterviniente '" & psCtaCod & "'"

    oConecta.AbreConexion
    Set oRs = oConecta.CargaRecordSet(sSql)
    If Not oRs.EOF Then
        CadenaPermiteModificarInterviniente = oRs!cMensaje
    End If
    oRs.Close
    Set oRs = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrCadenaPermiteModificarInterviniente:
    Err.Raise Err.Number, "Error En Proceso <CadenaPermiteModificarInterviniente>", Err.Description
End Function
'END EJVG *******
'ALPA 20160605***********************************************************
Public Function RecuperaCredEndeuCuotaSistFinanc(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    sSql = "Exec stp_sel_CredEndeuCuotaSistFinanc '" & psCtaCod & "'"
    oCon.AbreConexion
    Set RecuperaCredEndeuCuotaSistFinanc = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Sub ActualizaCredEndeuCuotaSistFinanc(pdFecha As Date, psCtaCod As String, psCod_Emp As String, pnDiaPago As Integer, psCtaContCod As String, pnDeudaIFi As Currency, pnCuotaEstimada As Currency, pnInteresDevengado As Currency)
Dim lsSQL As String
lsSQL = "EXEC stp_ins_CredEndeuCuotaSistFinanc '" & Format(pdFecha, "YYYY/MM/DD") & "', '" & psCtaCod & "', '" & psCod_Emp & "', " & pnDiaPago & ", '" & psCtaContCod & "', " & pnDeudaIFi & ", " & pnCuotaEstimada & ", " & pnInteresDevengado
loConecta.Ejecutar (lsSQL)
End Sub
Public Function RecuperaCredEndeuLineaNoUtilizada(ByVal psPersCod As String, ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    sSql = "Exec stp_sel_CredEndeuLineaNoUtilizada '" & psPersCod & "','" & psCtaCod & "'"
    oCon.AbreConexion
    Set RecuperaCredEndeuLineaNoUtilizada = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'CTI3 ERS032020**********************************************************************************************
Public Function RecuperaCredDeudaSistemaFinanciero(ByVal psPersCod As String, ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    sSql = "Exec stp_sel_CredDeudaSistemaFinanciero '" & psPersCod & "','" & psCtaCod & "'"
    oCon.AbreConexion
    Set RecuperaCredDeudaSistemaFinanciero = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Sub ActualizaCredDeudaSistemaFinanciero(pbActivo As Integer, pdFecha As Date, psCtaCod As String, psCod_Emp As String, psTipoDeuda As String, pnSaldoDeuda As Currency, psTitularTit As String)
Dim lsSQL As String
lsSQL = "EXEC stp_upd_CredDeudaSistemaFinanciero " & pbActivo & ",'" & Format(pdFecha, "YYYY/MM/DD") & "', '" & psCtaCod & "', '" & psCod_Emp & "','" & psTipoDeuda & "', " & pnSaldoDeuda & ",'" & psTitularTit & "'"
loConecta.Ejecutar (lsSQL)
End Sub
'****************************************************************************************************************
Public Sub ActualizaCredEndeuLineaNoUtilizada(pbActivo As Integer, pdFecha As Date, psCtaCod As String, psCod_Emp As String, psCtaContCod As String, pnSaldoTarjeta As Currency, pnCuota_LNU As Currency)
Dim lsSQL As String
lsSQL = "EXEC stp_ins_CredEndeuLineaNoUtilizada " & pbActivo & ",'" & Format(pdFecha, "YYYY/MM/DD") & "', '" & psCtaCod & "', '" & psCod_Emp & "','" & psCtaContCod & "', " & pnSaldoTarjeta & ", " & pnCuota_LNU
loConecta.Ejecutar (lsSQL)
End Sub
Public Function ObtenerCuotaEsperadaLC(pnValorDeuda As Currency) As Currency
Dim lsSQL As String
Dim oRs As ADODB.Recordset
Set oRs = New ADODB.Recordset
Dim oCon As New COMConecta.DCOMConecta

lsSQL = "EXEC stp_sel_ObtenerCuotaEsperadaLC " & pnValorDeuda & ""
oCon.AbreConexion
Set oRs = oCon.CargaRecordSet(lsSQL)
oCon.CierraConexion
Set oCon = Nothing

ObtenerCuotaEsperadaLC = 0
If Not (oRs.BOF Or oRs.EOF) Then
    ObtenerCuotaEsperadaLC = oRs!nCuota_LNU
End If
End Function
'************************************************************************
'EJVG20160203 ERS002-2016***
Public Function RecuperaCompraDeuda(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrRecuperaCompraDeuda

    sSql = "EXEC stp_sel_ERS0022016_ColocCompraDeuda '" & psCtaCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCompraDeuda = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrRecuperaCompraDeuda:
    Err.Raise Err.Number, "Error En Proceso de Recuperar Compra de Deuda del Credito", Err.Description
End Function
'END EJVG *******
'************************************************************************
'FRHU 20160702 ERS002-2016
'Public Sub RegistraAutorizacionesRequeridas(ByVal psFecSis As String, ByVal psCodUser As String, ByVal psCodAge As String, ByVal psNuevaCta As String)   'El nucleo'Comento JOEP20190208 CP
Public Sub RegistraAutorizacionesRequeridas(ByVal psFecSis As String, ByVal psCodUser As String, ByVal psCodAge As String, ByVal psNuevaCta As String, Optional ByVal pnDestino As Long = 0) 'JOEP20190123 CP
    Dim lsSQL As String
    'lsSQL = "EXEC stp_ins_ERS0022014_AutorizacionCredito '" & psFecSis & "','" & psCodUser & "','" & psCodAge & "','" & psNuevaCta & "'"'Comento JOEP20190208 CP
    lsSQL = "EXEC stp_ins_ERS0022014_AutorizacionCredito '" & psFecSis & "','" & psCodUser & "','" & psCodAge & "','" & psNuevaCta & "'," & pnDestino & "" 'JOEP20190123 CP
    loConecta.Ejecutar (lsSQL)
End Sub
Public Function verificarExisteAutorizaciones(ByVal psCtaCod As String) As Boolean
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim oRs As New ADODB.Recordset
    
    Set oConecta = New COMConecta.DCOMConecta
    sSql = "exec stp_sel_verificarExisteAutorizaciones '" & psCtaCod & "'"
    oConecta.AbreConexion
    Set oRs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    If Not (oRs.BOF Or oRs.EOF) Then
        verificarExisteAutorizaciones = True
    Else
        verificarExisteAutorizaciones = False
    End If
End Function
'** B y C
'FRHU 20160815 ANEXO002 ERS002-2016
'Public Function ValidarCondicionesDeIntervinientesEnElCredito(ByVal psCtaCod As String, ByVal psPersCod As String, ByVal pnCondicion As Integer) As ADODB.Recordset
'    Dim sSql As String
'    Dim oConecta As COMConecta.DCOMConecta
'
'    Set oConecta = New COMConecta.DCOMConecta
'    sSql = "exec stp_sel_ERS0022016_ValidarCondicionesIntervinientesEnElCredito '" & psCtaCod & "','" & psPersCod & "'," & pnCondicion
'    oConecta.AbreConexion
'    Set ValidarCondicionesDeIntervinientesEnElCredito = oConecta.CargaRecordSet(sSql)
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'End Function
'Public Function ValidarCondicionesCredAmpliados(ByVal psCtaCodAmp As String) As ADODB.Recordset
'    Dim sSql As String
'    Dim oConecta As COMConecta.DCOMConecta
'
'    Set oConecta = New COMConecta.DCOMConecta
'    sSql = "exec stp_sel_ERS0022016_ValidarCondicionesCredAmpliados '" & psCtaCodAmp & "'"
'    oConecta.AbreConexion
'    Set ValidarCondicionesCredAmpliados = oConecta.CargaRecordSet(sSql)
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'End Function
'Public Function ValidarCondicionesDeIntervinientesEnElCredito(ByVal psCtaCod As String, ByVal psPersCod As String, ByVal pnCondicion As Integer, ByVal psFecSis As String, ByVal psCodUser As String, ByVal psCodAge As String) As ADODB.Recordset
Public Function ValidarCondicionesDeIntervinientesEnElCredito(ByVal psCtaCod As String, ByVal psPersCod As String, ByVal pnCondicion As Integer, ByVal psFecSis As String, ByVal psCodUser As String, ByVal psCodAge As String, Optional ByVal psTipoProducto As String = "") As ADODB.Recordset 'FRHU 20160917 Observación
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    Set oConecta = New COMConecta.DCOMConecta
    'sSql = "exec stp_sel_ERS0022016_ValidarCondicionesIntervinientesEnElCredito '" & psCtaCod & "','" & psPersCod & "'," & pnCondicion & ",'" & psFecSis & "','" & psCodUser & "','" & psCodAge & "'"
    sSql = "exec stp_sel_ERS0022016_ValidarCondicionesIntervinientesEnElCredito '" & psCtaCod & "','" & psPersCod & "'," & pnCondicion & ",'" & psFecSis & "','" & psCodUser & "','" & psCodAge & "','" & psTipoProducto & "'"
    oConecta.AbreConexion
    Set ValidarCondicionesDeIntervinientesEnElCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function ValidarCondicionesCredAmpliados(ByVal psCtaCodAmp As String, ByVal psPersCod As String, ByVal psFecSis As String, ByVal psCodUser As String, ByVal psCodAge As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    Set oConecta = New COMConecta.DCOMConecta
    sSql = "exec stp_sel_ERS0022016_ValidarCondicionesCredAmpliados '" & psCtaCodAmp & "','" & psPersCod & "','" & psFecSis & "','" & psCodUser & "','" & psCodAge & "'"
    oConecta.AbreConexion
    Set ValidarCondicionesCredAmpliados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'FIN FRHU 20160815
'** P. Firma del conyugue
Public Function verificarAlertasRelaCred(ByVal psPersCodTitular As String, ByVal psPersCodConyuCode As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    Set oConecta = New COMConecta.DCOMConecta
    sSql = "exec stp_sel_ERS0022016_VerificarConyugueCodeudor '" & psPersCodTitular & "','" & psPersCodConyuCode & "'"
    oConecta.AbreConexion
    Set verificarAlertasRelaCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'FIN FRHU 20160702 ********************************************************
'EJVG20160711 ***
Public Function ObtenerMontoExposicionEsteCredito(ByVal psCtaCod As String, Optional ByVal pnMontoCol As Double = 0#) As Double
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    Dim oRs As ADODB.Recordset

    On Error GoTo ErrObtenerInformeRiesgoNEW

    sSql = "SELECT nExposicion = dbo.fnc_ObtenerExposicionEsteCredito('" & psCtaCod & "'," & pnMontoCol & ")"

    oConecta.AbreConexion
    Set oRs = oConecta.CargaRecordSet(sSql)
    If Not oRs.EOF Then
        ObtenerMontoExposicionEsteCredito = oRs!nExposicion
    End If
    oRs.Close
    Set oRs = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrObtenerInformeRiesgoNEW:
    Err.Raise Err.Number, "Error En Proceso <<ObtenerMontoExposicionEsteCredito>>", Err.Description
End Function
'END EJVG *******

'WIOR 20160609 ***
Public Function SobreEndObtenerCodigos(ByVal psCtaCod As String, ByVal pcNumFuente As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_SobreEndObtenerCodigos '" & psCtaCod & "','" & pcNumFuente & "'"
    oCon.AbreConexion
    Set SobreEndObtenerCodigos = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function SobreEndObtenerCodigosRegXCta(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_SobreEndCodigos '" & psCtaCod & "'"
    oCon.AbreConexion
    Set SobreEndObtenerCodigosRegXCta = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function SobreEndCreditosADesbloq(ByVal psCargo As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_SobreEndCredADesbloq '" & psCargo & "'"
    oCon.AbreConexion
    Set SobreEndCreditosADesbloq = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function SobreEndDatosCreditoADesbloq(ByVal psCtaCod As String, ByVal psCargo As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_SobreEndDatosCreditoADesbloq '" & psCtaCod & "','" & psCargo & "'"
    oCon.AbreConexion
    Set SobreEndDatosCreditoADesbloq = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function SobreEndDatosCreditoADesbloqCodigos(ByVal psCtaCod As String, Optional ByVal psCargo As String = "", Optional ByVal pbCargo As Boolean = False, Optional ByVal pbTodos As Boolean = False) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_SobreEndDatosCreditoADesbloqCodigos '" & psCtaCod & "','" & psCargo & "'," & IIf(pbCargo, 1, 0) & "," & IIf(pbTodos, 1, 0)
    oCon.AbreConexion
    Set SobreEndDatosCreditoADesbloqCodigos = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Sub SobreEndRegistrarAutCodigos(ByVal pnId As Long, ByVal pnCodigo As Integer, ByVal pnResultado As Integer, ByVal psPlan As String, ByVal pdFechaReg As Date, ByVal pcUserReg As String, ByVal psCargo As String)
Dim lsSQL As String

lsSQL = "EXEC stp_ins_SobreEndDesbloqueos " & pnId & "," & pnCodigo & "," & pnResultado & ",'" & psPlan & "','" & _
        Format(pdFechaReg + loConecta.GetHoraServer, "yyyymmdd hh:mm:ss") & "','" & pcUserReg & "','" & psCargo & "'"

loConecta.Ejecutar (lsSQL)
End Sub
Public Sub SobreEndRegularEstado(ByVal pcCtaCod As String)
Dim lsSQL As String

lsSQL = "EXEC stp_upd_SobreEndRegularEstado '" & pcCtaCod & "'"

loConecta.Ejecutar (lsSQL)
End Sub
'WIOR FIN ********
'RECO20160627 ERS002-2016 ***********************************************
Public Function RecuperaDatosCredBasicos(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta

    sSql = "Exec stp_sel_ERS0022016_RecuperaDatosCredBasicos '" & psCtaCod & "'"
    oCon.AbreConexion
    Set RecuperaDatosCredBasicos = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Sub ActualizaRatioCoberturaAnalista(ByVal psCtaCod As String, ByVal pnRatioCober As Currency)
Dim lsSQL As String
lsSQL = "EXEC stp_sel_ERS0022016_ActualizaRatioCoberturaAnalista '" & psCtaCod & "'," & pnRatioCober
loConecta.Ejecutar (lsSQL)
End Sub
'RECO FIN****************************************************************

'************************ LUCV20160513 - ERS2016-004 *****************************
'---------------------------------------------------------------------------------
'LUCV20160520 **********
Public Function RecuperaDatosRefinanciar(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset

    On Error GoTo ErrorRecuperaDatosRefinanciar
    sSql = " exec stp_sel_RecuperaDatosRefinanciar '" & psCtaCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    Set RecuperaDatosRefinanciar = rs
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
 
ErrorRecuperaDatosRefinanciar:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'Fin ->LUCV20160520 **********

'LUCV20160529 **********
Public Function RecuperaTotalRefinanciar(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset

    On Error GoTo ErrorRecuperaTotalRefinanciar
    sSql = " exec stp_sel_RecuperaTotalRefinanciar '" & psCtaCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    Set RecuperaTotalRefinanciar = rs
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
 
ErrorRecuperaTotalRefinanciar:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'Fin -> LUCV20160529 **********

'LUCV20160608,ERS004-2016 **********
Public Function RecuperaEstadoCreditoRefinanciado(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer

    On Error GoTo ErrorRecuperaEstadoCreditoRefinanciado

    sSql = "exec stp_sel_ObtieneEstadoCreditosRefinanciado '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaEstadoCreditoRefinanciado = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaEstadoCreditoRefinanciado:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'Fin LUCV20160608 **********


'LUCV20160614,ERS004-2016 **********
Public Function RecuperaCreditoRefinanciado(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer

    On Error GoTo ErrorRecuperaCreditoRefinanciado

    sSql = "exec stp_sel_ObtieneCreditoRefinanciado '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditoRefinanciado = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaCreditoRefinanciado:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'Fin LUCV20160614 **********

'FRHU 20160811 Anexo002 ERS002-2016
Public Function MostrarAutorizacionesHojaAprobacion(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_MostrarAutorizacionHojaAprobacion '" & psCtaCod & "'"
    oCon.AbreConexion
    Set MostrarAutorizacionesHojaAprobacion = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'FIN FRHU 20160811

'JOEP ERS066-20161110
Public Function ObtenerCampActiva(ByVal psCtaCod As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ObtenerCampActiva '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set ObtenerCampActiva = rs
    'End If
    Set rs = Nothing
    
End Function
Public Function ObtenerCampActivaApro(ByVal pnIdCamp As Integer) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "Select bEstado From Campanas"
    sSql = sSql & " where bEstado =1 and IdCampana =" & pnIdCamp & ""
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set ObtenerCampActivaApro = rs
    'End If
    Set rs = Nothing
    
End Function
Public Function ObtenerAsigCampActiva(ByVal pnIdCamp As Integer, ByVal pcCodAge As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_ValidaAsigCamap " & pnIdCamp & ",'" & pcCodAge & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set ObtenerAsigCampActiva = rs
    'End If
    Set rs = Nothing
    
End Function
'JOEP ERS066-20161110


'JOEP ERS059-2016
Public Function ObtenerCabeceraSobreEnd(ByVal pscPerscod As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ERS059_2016_ObtenerCabeceraSobreEnd '" & pscPerscod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set ObtenerCabeceraSobreEnd = rs
    'End If
    Set rs = Nothing
    
End Function

Public Function ObtenerIfis(ByVal pscPerscod As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ERS059_2016_MostrarIfis '" & pscPerscod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set ObtenerIfis = rs
    'End If
    Set rs = Nothing
    
End Function

Public Function ObtenerDiasAtraso(ByVal pscPerscod As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ERS059_2016_MostrarDiasAtraso '" & pscPerscod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set ObtenerDiasAtraso = rs
    'End If
    Set rs = Nothing
    
End Function

Public Function ObtenerEvalSobreEndAdm(ByVal pscPerscod As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ERS059_2016_MostrarEvalSobreEndAdm '" & pscPerscod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set ObtenerEvalSobreEndAdm = rs
    'End If
    Set rs = Nothing
    
End Function

Public Function ObtenerEvalSobreEndNuevo(ByVal pscPerscod As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ERS059_2016_SobreEndObtenerCodigosEvalNueva'" & pscPerscod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set ObtenerEvalSobreEndNuevo = rs
    'End If
    Set rs = Nothing
    
End Function

Public Function ObtenerFichaSobEndCabnId() As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
        
    sSql = "exec sp_sel_ObtenerFichaSobEndnId "
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set ObtenerFichaSobEndCabnId = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
            
End Function

Public Function ObtenerFichaSobEndMatnId(ByVal pscCtaCodEval As String, ByVal pcCodPers As String, ByVal pdFecSis As Date, ByVal pdFecCierre As Date) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
        
    sSql = "exec sp_sel_ObtenernIdMant '" & pscCtaCodEval & "','" & pcCodPers & "','" & Format(pdFecSis, "yyyyMMdd") & "','" & Format(pdFecCierre, "yyyyMMdd") & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set ObtenerFichaSobEndMatnId = rs
    'End If
    Set rs = Nothing
            
End Function

Public Sub GrabarCabFichaSobreEnd(ByVal pnTipoReg As Integer, ByVal pscCtaCodEval As String, ByVal pcCodPers As String, ByVal pcAna As String, ByVal pnAge As String, ByVal pnSalCapConso As Double, _
                                                ByVal pcCalfCamc As Integer, ByVal pnSalEndSistFin As Double, ByVal pnMarClieAval As Integer, ByVal pnMonCmac As Double, ByVal dFecSis As Date, ByVal nVecesReg As Integer, ByVal dFecCierre As Date)
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
     oCon.AbreConexion
                  
     If pnTipoReg = 1 Then
        sSql = "exec stp_ins_ERS059_2016_GuardarFichaSobEnd '" & pscCtaCodEval & "','" & pcCodPers & "','" & pcAna & "','" & pnAge & "'," & pnSalCapConso & "," & pcCalfCamc & "," & pnSalEndSistFin & "," & pnMarClieAval & "," & pnMonCmac & ",'" & Format(dFecSis + loConecta.GetHoraServer, "yyyyMMdd hh:mm:ss") & "'," & nVecesReg & ",'" & Format(dFecCierre, "yyyyMMdd") & "'"
    Else
        sSql = "exec sp_upd_ERS059_2016_ActualizarCabFichaSobEnd '" & pscCtaCodEval & "','" & pcCodPers & "','" & pcAna & "','" & pnAge & "'," & pnSalCapConso & "," & pcCalfCamc & "," & pnSalEndSistFin & "," & pnMarClieAval & "," & pnMonCmac & ",'" & Format(dFecSis + loConecta.GetHoraServer, "yyyymmdd hh:mm:ss") & "'," & nVecesReg & ",'" & Format(dFecCierre, "yyyyMMdd") & "'"
    End If
    
    oCon.ConexionActiva.Execute sSql
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub GuardarFichaIfis(ByVal nId As Integer, ByVal pscCtaCodEval As String, ByVal pcCodPers As String, ByVal pcIfis As String, ByVal pnMonto As Double, ByVal gdFecSis As Date, ByVal dFecCierre As Date)
                                                
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
     oCon.AbreConexion
     
    sSql = "exec stp_ins_ERS059_2016_GuardarIfisFichaSobEnd " & nId & ",'" & pscCtaCodEval & "','" & pcCodPers & "','" & pcIfis & "'," & pnMonto & ",'" & Format(gdFecSis + loConecta.GetHoraServer, "yyyymmdd hh:mm:ss") & "','" & Format(dFecCierre, "yyyyMMdd") & "'"
       
    oCon.ConexionActiva.Execute sSql
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub GuardarFichaDiasAtrs(ByVal nId As Integer, ByVal pscCtaCodEval As String, ByVal pcCodPers As String, ByVal pcCtaCod As String, ByVal pcCredEstado As String, ByVal pnDiasAtrs As String, ByVal gdFecSis As Date, ByVal dFecCierre As Date)
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
     oCon.AbreConexion
     
    sSql = "exec stp_ins_ERS059_2016_GuardarDiasAtraFichaSobEnd " & nId & ",'" & pscCtaCodEval & "','" & pcCodPers & "','" & pcCtaCod & "','" & pcCredEstado & "'," & pnDiasAtrs & ",'" & Format(gdFecSis + loConecta.GetHoraServer, "yyyymmdd hh:mm:ss") & "','" & Format(dFecCierre, "yyyyMMdd") & "'"
       
    oCon.ConexionActiva.Execute sSql
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub GuardarFichaNew(ByVal nId As Integer, ByVal pscCtaCodEval As String, ByVal pcCodPers As String, ByVal pcCtaCod As String, ByVal pnCodigo As Integer, ByVal pcEval As Integer, ByVal pcPlanMitig As String, ByVal gdFecSis As Date, ByVal dFecCierre As Date)
                                                
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
     oCon.AbreConexion
     
    sSql = "exec stp_ins_ERS059_2016_GuardarFichaSobEndNew " & nId & ",'" & pscCtaCodEval & "','" & pcCodPers & "','" & pcCtaCod & "'," & pnCodigo & "," & pcEval & ",'" & pcPlanMitig & "','" & Format(gdFecSis + loConecta.GetHoraServer, "yyyymmdd hh:mm:ss") & "','" & Format(dFecCierre, "yyyyMMdd") & "'"
       
    oCon.ConexionActiva.Execute sSql
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub GuardarFichaAdm(ByVal nId As Integer, ByVal pscCtaCodEval As String, ByVal pcCodPers As String, ByVal pcCtaCod As String, ByVal pnCodigo As Integer, ByVal pcEval As Integer, ByVal pcPlanMitig As String, _
                                                ByVal pcMantCodigo As String, ByVal gdFecSis As Date, ByVal dFecCierre As Date)
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
     oCon.AbreConexion
     
    sSql = "exec stp_ins_ERS059_2016_GuardarFichaSobEndAdm " & nId & ",'" & pscCtaCodEval & "','" & pcCodPers & "','" & pcCtaCod & "'," & pnCodigo & "," & pcEval & ",'" & pcPlanMitig & "','" & pcMantCodigo & "','" & Format(gdFecSis + loConecta.GetHoraServer, "yyyymmdd hh:mm:ss") & "','" & Format(dFecCierre, "yyyyMMdd") & "'"
       
    oCon.ConexionActiva.Execute sSql
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Function ObtenerCabFichaSobEnd(ByVal pscPerscod As String, ByVal dfechaReg As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec sp_sel_ERS059_2016_ObtenerCabeFichaSobEnd '" & pscPerscod & "','" & dfechaReg & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set ObtenerCabFichaSobEnd = rs
    'End If
    Set rs = Nothing
    
End Function

Public Function ObtenerFichaSobEndIfis(ByVal pscPerscod As String, ByVal dfechaReg As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec sp_sel_ERS059_2016_ObtenerFichaSobEndIFIS '" & pscPerscod & "','" & dfechaReg & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set ObtenerFichaSobEndIfis = rs
    'End If
    Set rs = Nothing
    
End Function

Public Function ObtenerFichaSobEndDiasAtrs(ByVal pscPerscod As String, ByVal dfechaReg As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec sp_sel_ERS059_2016_ObtenerFichaSobEndDiasAtrs '" & pscPerscod & "','" & dfechaReg & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set ObtenerFichaSobEndDiasAtrs = rs
    'End If
    Set rs = Nothing
    
End Function

Public Function PDFObtenerFichaSobEndNew(ByVal pscPerscod As String, ByVal dfechaReg As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec sp_sel_ERS059_2016_PDFObtenerFichaSobEndNew '" & pscPerscod & "','" & dfechaReg & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set PDFObtenerFichaSobEndNew = rs
    'End If
    Set rs = Nothing
    
End Function

Public Function ObtenerFichaSobEndNew(ByVal pscPerscod As String, ByVal dfechaReg As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec sp_sel_ERS059_2016_ObtenerFichaSobEndNew '" & pscPerscod & "','" & dfechaReg & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set ObtenerFichaSobEndNew = rs
    'End If
    Set rs = Nothing
    
End Function

Public Function PDFObtenerFichaSobEndAdm(ByVal pscPerscod As String, ByVal dfechaReg As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec sp_sel_ERS059_2016_PDFObtenerFichaSobEndAdm '" & pscPerscod & "','" & dfechaReg & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set PDFObtenerFichaSobEndAdm = rs
    'End If
    Set rs = Nothing
    
End Function

Public Function ObtenerFichaSobEndAdm(ByVal pscPerscod As String, ByVal dfechaReg As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec sp_sel_ERS059_2016_ObtenerFichaSobEndAdm '" & pscPerscod & "','" & dfechaReg & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set ObtenerFichaSobEndAdm = rs
    'End If
    Set rs = Nothing
    
End Function

Public Function ListarFichaSobreEnd() As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    sSql = "exec stp_sel_ERS059_2016_ListaFichaSobreEnd "

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ListarFichaSobreEnd = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function EliminarFichaSobEndIfis(ByVal pscCtaCodEval As String, ByVal pscPerscod As String, ByVal dfechaReg As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec sp_del_ERS059_2016_EliminarFichaSobEndIfis '" & pscCtaCodEval & "','" & pscPerscod & "','" & Format(dfechaReg, "yyyymmdd") & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set EliminarFichaSobEndIfis = rs
    'End If
    Set rs = Nothing
    
End Function

Public Function EliminarFichaSobEndDiasAtraso(ByVal pscCtaCodEval As String, ByVal pscPerscod As String, ByVal dfechaReg As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec sp_del_ERS059_2016_EliminarFichaSobEndDiasAtraso '" & pscCtaCodEval & "','" & pscPerscod & "','" & Format(dfechaReg, "yyyymmdd") & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set EliminarFichaSobEndDiasAtraso = rs
    'End If
    Set rs = Nothing
    
End Function

Public Function EliminarFichaSobEndNew(ByVal pscCtaCodEval As String, ByVal pscPerscod As String, ByVal dfechaReg As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec sp_del_ERS059_2016_EliminarFichaSobEndNew '" & pscCtaCodEval & "','" & pscPerscod & "','" & Format(dfechaReg, "yyyymmdd") & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set EliminarFichaSobEndNew = rs
    'End If
    Set rs = Nothing
    
End Function

Public Function EliminarFichaSobEndAdm(ByVal pscCtaCodEval As String, ByVal pscPerscod As String, ByVal dfechaReg As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec sp_del_ERS059_2016_EliminarFichaSobEndAdm '" & pscCtaCodEval & "','" & pscPerscod & "','" & Format(dfechaReg, "yyyymmdd") & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    'If Not rs.EOF Then
        Set EliminarFichaSobEndAdm = rs
    'End If
    Set rs = Nothing
    
End Function

Public Function ValidaReg(ByVal cCtaCodEval As String, ByVal cPersCod As String, ByVal dfechaReg As String) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    sSql = "exec SP_SEL_ERS059_2016_VALIDAREGISTRO '" & cCtaCodEval & "','" & cPersCod & "','" & Format(dfechaReg, "yyyymmdd") & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ValidaReg = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function ObtenerBusquedaFichaSobEnd(ByVal cPersCod As String, ByVal nMes As Integer, ByVal nAno As Integer) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    sSql = "exec sp_sel_ERS059_2016_BuscarFichaSobEnd '" & cPersCod & "'," & nMes & "," & nAno & ""

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerBusquedaFichaSobEnd = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function ValReg(ByVal cPersCod As String) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    sSql = "exec sp_sel_ValidaRegistro '" & cPersCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ValReg = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function MostrarComboFecha() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_MostraComboFecha"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set MostrarComboFecha = rs
    Set rs = Nothing

End Function
'JOEP ERS059-2016

'Inicio JOEP Reprogramacion

Public Function ObtenerDiasMes(ByVal pscCtaCod As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec sp_sel_nDiasMes '" & Format(pscCtaCod, "yyyymmdd") & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set ObtenerDiasMes = rs
        
    Set rs = Nothing
    
End Function

Public Function VerificaExisteInteresGracia(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, ByVal pnCuota As Integer, ByVal pnConceptoCod As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "Exec sp_sel_ExisteGraci '" & psCtaCod & "'," & pnNroCalen & "," & pnCuota & "," & pnConceptoCod & ""
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set VerificaExisteInteresGracia = rs
        
    Set rs = Nothing
End Function

Public Function IdentificarTipoPeriodo(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "Exec sp_sel_IdentTipodePeriodo '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set IdentificarTipoPeriodo = rs
        
    Set rs = Nothing
    
End Function
'Fin JOEP Reprogramacion


'JOEP ERS064-20161019 -- Segmentacion
Public Function CreditoInformeRiesgoPendSinCampana() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorCreditoInformeRiesgoPendSinCampana

    sSql = "exec stp_sel_ERS064_CreditosInformeRiesgoPendSinCampaña"
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set CreditoInformeRiesgoPendSinCampana = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
    
ErrorCreditoInformeRiesgoPendSinCampana:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function CreditoInformeRiesgoPendConCampana() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorCreditoInformeRiesgoPendConCampana

    sSql = "exec stp_sel_ERS064_CreditosInformeRiesgoPendConCampaña"
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set CreditoInformeRiesgoPendConCampana = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
    
ErrorCreditoInformeRiesgoPendConCampana:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function VerificaInfRisgNewMod(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorVerificaInfRisgNewMod

    sSql = "exec stp_sel_ERS064_VerificarInformeRisgModNew '" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set VerificaInfRisgNewMod = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
    
ErrorVerificaInfRisgNewMod:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function MostrarCreditoInformeRiesgoPend(ByVal psCtaCod As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ERS064_MostrarCreditosInformeRiesgoPend '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
       
    Set MostrarCreditoInformeRiesgoPend = rs
    Set rs = Nothing
    
End Function

Public Function GuardarRiesgoInformeMontos(ByVal psCtaCod As String, ByVal psAgencia As String, ByVal psPersCod As String, ByVal psTpProducto As String, ByVal psTpCredito As String, _
ByVal pdfechaExp As String, ByVal psDescripcion As String, ByVal pnMontProp As Double, ByVal pnCuotProp As Double, ByVal pnCuota As Integer, ByVal pnTEA As Double, ByVal pnTenAnt As Double, ByVal pnTEM As Double)
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    sSql = "exec stp_ins_ERS064_RiesgoInformeMonto '" & psCtaCod & "','" & psAgencia & "','" & psPersCod & "','" & psTpProducto & "','" & psTpCredito & "','" & Format(pdfechaExp, "yyyy/MM/dd hh:mm:hh") & "','" & psDescripcion & "'," & pnMontProp & "," & pnCuotProp & "," & pnCuota & "," & pnTEA & "," & pnTenAnt & "," & pnTEM & ""
    
    oCon.ConexionActiva.Execute sSql
    oCon.CierraConexion
    Set oCon = Nothing
    
    GuardarRiesgoInformeMontos = True
    
End Function

Public Function ActualizaRiesgoInformeMontos(ByVal psCtaCod As String, ByVal psMovSalExp As String, ByVal pnEstado As Integer)
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    sSql = "exec spt_upd_ERS064_ActInformeRiesgoMontos '" & psCtaCod & "','" & psMovSalExp & "'," & pnEstado & ""
    
    oCon.ConexionActiva.Execute sSql
    oCon.CierraConexion
    Set oCon = Nothing
    
End Function

Public Function MostrarMorarAgencia(ByVal pdFechaSis As String, ByVal psAgeCod As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ERS064_MoraAgeUltCierre '" & pdFechaSis & "','" & psAgeCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
       
    Set MostrarMorarAgencia = rs
    Set rs = Nothing
    
End Function

Public Function MostrarCategoriaAgencia(ByVal psAgeCod As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ERS064_CredCatAgencia '" & psAgeCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set MostrarCategoriaAgencia = rs
    Set rs = Nothing
    
End Function

Public Function IdentificarMoraAge(ByVal pnMoraAge As Double, ByVal pnCatAgeB As Double, ByVal pnCatAgeM As Double, ByVal pnCatAgeA As Double, ByVal pnCatAgeE As Double) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ERS064_IdentificarMoraAge " & pnMoraAge & "," & pnCatAgeB & "," & pnCatAgeM & "," & pnCatAgeA & "," & pnCatAgeE & ""
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
        
    Set IdentificarMoraAge = rs
    Set rs = Nothing
    
End Function

Public Function IdentificaMoraCos(ByVal pnMoraCos As Double) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ERS064_IdentificarMoraCosecha " & pnMoraCos & ""
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
        
    Set IdentificaMoraCos = rs
    Set rs = Nothing
    
End Function

Public Function MostrarMoraCosechaRiesgo(ByVal psCtaCod As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ERS064_ObtenerMoraCosecha '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
        
    Set MostrarMoraCosechaRiesgo = rs
    Set rs = Nothing
    
End Function

Public Function MostrarMontoRiesgo(ByVal psCatAge As Integer, ByVal psCriterio As Integer) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ERS064_MostrarMontoRiesgo " & psCatAge & "," & psCriterio & ""
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
        
    Set MostrarMontoRiesgo = rs
    Set rs = Nothing
    
End Function

Public Function LlenarConfMonto()
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorLlenarConfMonto

    sSql = "exec stp_sel_ERS064_MostraConfMonto"
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set LlenarConfMonto = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
    
ErrorLlenarConfMonto:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function LlenarConfCatAge()
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorLlenarConfCatAge

    sSql = "exec stp_sel_ERS064_ObtCatAge"
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set LlenarConfCatAge = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
    
ErrorLlenarConfCatAge:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function GrabarDatosConfMonto(ByVal pscAgecat As Integer, ByVal psCriterio As Integer, ByVal pnRiesgo As Integer, ByVal pnMonto As Double) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
 On Error GoTo ErrorGrabarDatosConfMonto
    sSql = "exec stp_upd_ERS064_ActualizarConfMonto " & pscAgecat & "," & psCriterio & "," & pnRiesgo & "," & pnMonto & ""
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
        
    Set GrabarDatosConfMonto = rs
    Set rs = Nothing
    
    Exit Function
    
ErrorGrabarDatosConfMonto:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function GrabarDatosConfCatAge(ByVal psCodAge As String, ByVal pnValor As Integer, ByVal pnMonto As Double) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
 On Error GoTo ErrorGrabarDatosConfMonto
    sSql = "exec stp_upd_ERS064_ActualizarConfCatAge '" & psCodAge & "'," & pnValor & "," & pnMonto & ""
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
        
    Set GrabarDatosConfCatAge = rs
    Set rs = Nothing
    
    Exit Function
    
ErrorGrabarDatosConfMonto:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ObtSaldObs(ByVal psCodAge As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
 On Error GoTo ErrorObtSaldObs
 
    sSql = "exec stp_sel_ERS064_ObtSalidaObs '" & psCodAge & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
        
    Set ObtSaldObs = rs
    Set rs = Nothing
    
    Exit Function
    
ErrorObtSaldObs:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'JOEP ERS064-20161019 -- Segmentacion ***********************************************

'==Inicio== JOEP ERS039-20170616 --Extorno Reprogramcion
Public Function ObtenerEstadoExtoRepg(ByVal pnBus As Integer, Optional ByVal pcUser As String, Optional ByVal pcCtaCod As String, Optional ByVal pcCodPers As String)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorObtenerEstadoExtoRepg

    sSql = "exec sp_sel_ERS039_ObtDatosRepro " & pnBus & ",'" & pcUser & "','" & pcCtaCod & "','" & pcCodPers & "'"
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set ObtenerEstadoExtoRepg = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
    
ErrorObtenerEstadoExtoRepg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ValidaPagoPrimerExtoRep(ByVal pcCtaCod As String)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorValidaPagoExtoRep

    sSql = "exec stp_sel_ERS039_ExtornoValidaPago '" & pcCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set ValidaPagoPrimerExtoRep = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
    
ErrorValidaPagoExtoRep:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ValidaPagoMesExtoRep(ByVal pcCtaCod As String)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorValidaPagoMesExtoRep

    sSql = "exec stp_sel_ERS039_ExtornoValidaMes '" & pcCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set ValidaPagoMesExtoRep = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
    
ErrorValidaPagoMesExtoRep:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ExtornoReprgActualizaMov(ByVal pcCtaCod As String, ByVal pcUser As String, ByVal pdFecSis As Date, ByVal pcCodAge As String, ByVal pnTCEA As Double) ''PRueba Begin tras
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorExtornoReprgActualizaMov

    sSql = "exec stp_upd_ERS039_ExtornaMov '" & pcCtaCod & "','" & pcUser & "','" & Format(pdFecSis + loConecta.GetHoraServer, "yyyyMMdd hh:mm:ss") & "','" & pcCodAge & "'," & pnTCEA & "" 'PRueba Begin tras
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set ExtornoReprgActualizaMov = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
    
ErrorExtornoReprgActualizaMov:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ExtornoReprgActualizaCalendAnt(ByVal pcCtaCod As String, ByVal pnTCEA As Double)

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorExtornoReprgActualizaCalendAnt

    sSql = "exec stp_upd_ERS039_ExtornaCalenAnt '" & pcCtaCod & "'," & pnTCEA & ""
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set ExtornoReprgActualizaCalendAnt = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
    
ErrorExtornoReprgActualizaCalendAnt:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ExtornoReprgActualizaEstado(ByVal pcCtaCod As String, ByVal pcUser As String, ByVal pdFecSis As Date, ByVal pcCodAge As String)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorExtornoReprgActualizaEstado

    sSql = "exec stp_upd_ERS039_ExtornaEstado '" & pcCtaCod & "','" & pcUser & "','" & Format(pdFecSis + loConecta.GetHoraServer, "yyyyMMdd hh:mm:ss") & "','" & pcCodAge & "'"
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set ExtornoReprgActualizaEstado = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
    
ErrorExtornoReprgActualizaEstado:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ExtornoReprgEliminaCaled(ByVal pcCtaCod As String)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorExtornoReprgEliminaCale

    sSql = "exec stp_del_ERS039_ExtornaCalend '" & pcCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set ExtornoReprgEliminaCaled = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
    
ErrorExtornoReprgEliminaCale:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ObtieneCalendario(ByVal pcCtaCod As String)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorObtieneCalendario

    sSql = "exec stp_sel_ERS039_ObtCalendario '" & pcCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set ObtieneCalendario = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
    
ErrorObtieneCalendario:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'==FIN== JOEP ERS039-20170616 --Extorno Reprogramcion
'PTI1 20170701 INICIO *************************************
Public Function CargaAnalistasEx(Optional ByVal psAge As String = -1) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sAnalistas As String
Dim oGen As COMDConstSistema.DCOMGeneral
    On Error GoTo ERRORCargaAnalistasEx
    Set oGen = New COMDConstSistema.DCOMGeneral
    Set oGen = Nothing
    sSql = "stp_sel_ERS041_2017_CargaAnalistasEx '" & Right(psAge, 2) & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CargaAnalistasEx = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
ERRORCargaAnalistasEx:
    Err.Raise Err.Number, "Error", Err.Description
End Function
'END PTI1 20170701 FIN ************************************************


''*** PEAC 20170711
Public Function RecuperaSaldosCartaFianzaCartera(ByVal nTipoCambio As String, ByVal pdFecha As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaSaldosCartaFianzaCartera
    sSql = "exec stp_sel_RecuperaSaldosCartaFianzaCartera '" & Format(pdFecha, "YYYY/MM/DD") & "'," & nTipoCambio
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaSaldosCartaFianzaCartera = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaSaldosCartaFianzaCartera:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function


''*** PEAC 20170717
Public Function RecuperaSaldosInteresDiferidosNuevo(ByVal nTipoCambio As String, ByVal pdFecha As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaSaldosInteresDiferidos
    sSql = "exec stp_sel_SaldoInteresesDiferidosNuevo '" & Format(pdFecha, "YYYY/MM/DD") & "'," & nTipoCambio
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaSaldosInteresDiferidosNuevo = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaSaldosInteresDiferidos:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

''*** PEAC 20170725
Public Function RecuperaSaldosCtasEmpreSistFinanc(ByVal pdFecha As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaSaldosCtasEmpreSistFinanc
    sSql = "exec stp_sel_ObtieneCtasBalEmpreSecFinan '" & pdFecha & "' "
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaSaldosCtasEmpreSistFinanc = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaSaldosCtasEmpreSistFinanc:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

''*** PEAC 20170725
Public Function RecuperaSaldosCtasCnt110803(ByVal pdFecha As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaSaldosCtasCnt110803
    sSql = "exec stp_sel_ObtieneCtasCnt110803 '" & pdFecha & "' "
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaSaldosCtasCnt110803 = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaSaldosCtasCnt110803:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'*** PEAC 20170725 - 20180119
Public Function RecuperaSaldosCtasCntAccionariales(ByVal pdFecha As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaSaldosCtasCntAccionariales
    sSql = "exec stp_sel_ObtieneCtasCntAccionariales '" & pdFecha & "' "
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaSaldosCtasCntAccionariales = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaSaldosCtasCntAccionariales:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'JOEP Codigos Sobreendeudamiento ERS44
Public Function MostrarCmbTipProd(ByVal nCod As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ERS044_2017_CargaTipoProd " & nCod & ""
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set MostrarCmbTipProd = rs
    Set rs = Nothing

End Function

Public Function MostrarCmbTipCalifSobEnd() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ERS044_2017_CargaTipoCalifCodSobre  "
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set MostrarCmbTipCalifSobEnd = rs
    Set rs = Nothing

End Function

Public Function MostrarDatosCodSobreFlexEdit(ByVal nValor As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ERS044_2017_CargaTipoCodigo " & nValor & ""
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set MostrarDatosCodSobreFlexEdit = rs
    Set rs = Nothing

End Function

Public Function MostrarDatosPlanMitig(ByVal nValor As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ERS044_2017_ObtenerDatosPlanMitigacion " & nValor & ""
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set MostrarDatosPlanMitig = rs
    Set rs = Nothing

End Function

Public Function GrabarDatosCodigosSob(ByVal pnId As Integer, ByVal psCod As String, Optional ByVal pnRangoIni As Double, Optional ByVal pnRangoFin As Double, Optional ByVal pnPuntaje As Integer, Optional ByVal pnTipProd As Integer, Optional ByVal pnTipCliente As Integer, Optional ByVal pnCondicion As String, Optional ByVal pnEntidad As Integer, Optional ByVal pnClisf As Integer) As ADODB.Recordset
                                                                                                                                                                                
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
 On Error GoTo ErrorGrabarDatosCodigosSob
    
    If psCod = 1 Then
        sSql = "exec stp_ins_ERS044_2017_InsertarCod " & pnId & "," & psCod & "," & pnRangoIni & "," & pnRangoFin & "," & pnPuntaje & "," & pnTipProd & ""
    ElseIf psCod = 2 Then
        sSql = "exec stp_ins_ERS044_2017_InsertarCod " & pnId & "," & psCod & "," & pnRangoIni & "," & pnRangoFin & "," & pnPuntaje & ""
    ElseIf psCod = 3 Then
        sSql = "exec stp_ins_ERS044_2017_InsertarCod " & pnId & "," & psCod & "," & pnRangoIni & "," & pnRangoFin & "," & pnPuntaje & "," & pnTipProd & ""
    ElseIf psCod = 4 Then
        sSql = "exec stp_ins_ERS044_2017_InsertarCod " & pnId & "," & psCod & "," & pnRangoIni & "," & pnRangoFin & "," & pnPuntaje & "," & pnTipProd & "," & pnTipCliente & "," & pnCondicion & "," & pnEntidad & ""
    End If
        
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
        
    Set GrabarDatosCodigosSob = rs
    Set rs = Nothing
    
    Exit Function
    
ErrorGrabarDatosCodigosSob:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function GrabarDatosModulosCodigosSob(ByVal pnId As Integer, ByVal pnModulo As Integer, Optional ByVal pcModulo As String, Optional ByVal pnCodigo As Integer, Optional ByVal pnRangoIni As Double, Optional ByVal pnRangoFin As Double, Optional ByVal pnClisf As Integer, Optional ByVal pnPorcentaje As Double) As ADODB.Recordset
                                                                                                                                                                                
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
 On Error GoTo ErrorGrabarDatosModulosCodigosSob
        
    sSql = "exec stp_ins_ERS044_2017_ModulosCodigosSobre " & pnId & "," & pnModulo & ",'" & pcModulo & "'," & pnCodigo & "," & pnRangoIni & "," & pnRangoFin & "," & pnClisf & "," & pnPorcentaje & ""
            
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
        
    Set GrabarDatosModulosCodigosSob = rs
    Set rs = Nothing
    
    Exit Function
    
ErrorGrabarDatosModulosCodigosSob:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function CargaDatosCodigos(ByVal pnCodigo As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ERS044_2017_CargaDatosCodigo  " & pnCodigo & ""
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set CargaDatosCodigos = rs
    Set rs = Nothing

End Function

'JOEP Codigos Sobreendeudamiento ERS44

'JOEP Codigos Sobreendeudamiento ERS47
Public Function CargaDatosLimiteZonaGeog() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec sp_sel_ERS047_CargaDatosLimiteZonaGeog "
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set CargaDatosLimiteZonaGeog = rs
    Set rs = Nothing

End Function

Public Function UpdateLimiteZonaGeog(ByVal pnCodZonaGeo As Integer, ByVal pnLimite As Double) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec sp_upd_ERS047_CargaDatosLimiteZonaGeog " & pnCodZonaGeo & "," & pnLimite & ""
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set UpdateLimiteZonaGeog = rs
    Set rs = Nothing

End Function

Public Function RecuperaSolicitudZonaGeog() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_sel_ERS047_RecuperaSolicitudZonaGeografica"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set RecuperaSolicitudZonaGeog = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function

Public Function VerificaSuperaUmbralZonaGeog(ByVal pcAgeCod As String, ByVal pnMontoSol As Double, ByVal pnMonedaSol As Integer) As Boolean
Dim sSql As String
Dim rs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_sel_ERS047_VerificaSuperaUmbralZonaGeog '" & pcAgeCod & "'," & pnMontoSol & "," & pnMonedaSol

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set rs = oConecta.CargaRecordSet(sSql)
If Not (rs.EOF And rs.BOF) Then
    If rs!Resultado <> "" Then
        VerificaSuperaUmbralZonaGeog = True
    End If
End If
oConecta.CierraConexion
Set oConecta = Nothing
End Function

Public Sub InsertarSolicitudAutorizacionZonaGeog(ByVal psCtaCod As String, ByVal pnMontoSol As Double)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_ins_ERS047_SolicitudAutorizacionZonaGeog '" & psCtaCod & "'," & pnMontoSol

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Call oConecta.Ejecutar(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Sub

Public Function VerificaSolicitudAutorizacionZonaGeog(ByVal pcCtaCod As String, ByVal pnProduto As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_sel_ERS047_VerificaSolicitudAutorizacionZonaGeog '" & pcCtaCod & "','" & pnProduto & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set VerificaSolicitudAutorizacionZonaGeog = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function

Public Sub ActualizarSolicitudAutorizacionZonaGeog(ByVal psCtaCod As String, ByVal psGlosa As String, ByVal pnEstado As Integer)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
'0 Solicitado
'1 Autorizado
'2 Rechazado
sSql = "exec stp_upd_ERS047_ActualizarSolicitudAutorizacionZonaGeog '" & psCtaCod & "','" & psGlosa & "'," & pnEstado

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Call oConecta.Ejecutar(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Sub

'Limite TpCredito
Public Function CargaDatosLimiteTpCredito() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec sp_sel_ERS047_CargaDatosLimiteTpProducto "
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set CargaDatosLimiteTpCredito = rs
    Set rs = Nothing

End Function

Public Function UpdateLimiteTpCredito(ByVal pnCodZonaGeo As Integer, ByVal pnLimite As Double) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec sp_upd_ERS047_CargaDatosLimiteTpProducto " & pnCodZonaGeo & "," & pnLimite & ""
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set UpdateLimiteTpCredito = rs
    Set rs = Nothing

End Function

Public Function RecuperaSolicitudTpCredito() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_sel_ERS047_RecuperaSolicitudTpCredito"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set RecuperaSolicitudTpCredito = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function

Public Sub ActualizarSolicitudAutorizacionTpCredito(ByVal psCtaCod As String, ByVal psGlosa As String, ByVal pnEstado As Integer)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
'0 Solicitado
'1 Autorizado
'2 Rechazado
sSql = "exec stp_upd_ERS047_ActualizarSolicitudAutorizacionTpCredito '" & psCtaCod & "','" & psGlosa & "'," & pnEstado

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Call oConecta.Ejecutar(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Sub

Public Function VerificaSuperaUmbralTpCredito(ByVal pcTpProducto As String, ByVal pnMontoSol As Double, ByVal pnMonedaSol As Integer) As Boolean
Dim sSql As String
Dim rs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_sel_ERS047_VerificaSuperaUmbralTpProducto '" & pcTpProducto & "'," & pnMontoSol & "," & pnMonedaSol

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set rs = oConecta.CargaRecordSet(sSql)
If Not (rs.EOF And rs.BOF) Then
    If rs!Resultado <> "" Then
        VerificaSuperaUmbralTpCredito = True
    End If
End If
oConecta.CierraConexion
Set oConecta = Nothing
End Function

Public Sub InsertarSolicitudAutorizacionTpCredito(ByVal psCtaCod As String, ByVal pnTpProducto As Integer, ByVal pnMontoSol As Double)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_ins_ERS047_SolicitudAutorizacionTpProducto '" & psCtaCod & "','" & pnTpProducto & "'," & pnMontoSol & ""

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Call oConecta.Ejecutar(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Sub

Public Function VerificaGarantAutoliq(ByVal pcCtaCod As String, ByVal pnValor As Integer) As Boolean
Dim sSql As String
Dim rs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec sp_sel_ERS047_ObtenerGarntAutoLiquixHipt '" & pcCtaCod & "'," & pnValor & ""

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set rs = oConecta.CargaRecordSet(sSql)
If Not (rs.EOF And rs.BOF) Then
    If rs!cNumGarant <> "" Then
        VerificaGarantAutoliq = True
    End If
End If
oConecta.CierraConexion
Set oConecta = Nothing
End Function

Public Sub EliminarSolicitudAutorizacionZonaxProduxGarant(ByVal psCtaCod As String, ByVal nTpProducto As String, ByVal pnValor As Integer)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_del_ERS047_SolicitudAutorizacionZonaxProductoxGarant '" & psCtaCod & "','" & nTpProducto & "'," & pnValor & ""

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Call oConecta.Ejecutar(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Sub

Public Function VerificaSolicitudAutorizacionProducto(ByVal pcCtaCod As String, ByVal pnTpProducto As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_sel_ERS047_VerificaSolicitudAutorizacionProducto '" & pcCtaCod & "','" & pnTpProducto & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set VerificaSolicitudAutorizacionProducto = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
'JOEP Codigos Sobreendeudamiento ERS47
'FRHU 20171110
Public Function HabilitarDesembolsoTercero(ByVal psCtaCod As String, ByVal pnDestinoCred As Integer, ByVal psTipoProd As String) As Boolean
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "EXEC stp_sel_HabilitarDesembolsoTercero '" & psCtaCod & "'," & pnDestinoCred & ",'" & psTipoProd & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    If rs!valor = 1 Then
        HabilitarDesembolsoTercero = True
    Else
        HabilitarDesembolsoTercero = False
    End If
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'FIN FRHU 20171110

'VAPA 20171115 FROM 60 Comentado by NAGL 20180609
'Public Function RecCreditoAmpliadoNew(ByVal psCtaCod As String) As ADODB.Recordset
'Dim sSql As String
'Dim oConecta As COMConecta.DCOMConecta
'
'    On Error GoTo ErrorCreditoAmpliadoNew
'
'    sSql = "exec stp_sel_CreditoAmpliadoNew '" & psCtaCod & "'"
'
'    Set oConecta = New COMConecta.DCOMConecta
'    oConecta.AbreConexion
'    Set RecCreditoAmpliadoNew = oConecta.CargaRecordSet(sSql)
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Function
'
'ErrorCreditoAmpliadoNew:
'    Err.Raise Err.Number, "Error En Proceso", Err.Description
'End Function
'Public Function ExistePreAmpliacion(ByVal psCtaCod As String, ByVal pdFecha As Date) As Boolean
'Dim sSql As String
'Dim oConecta As COMConecta.DCOMConecta
'Dim R As ADODB.Recordset
'
'    On Error GoTo ErrorPreAmpliacion
'    sSql = "exec stp_sel_ExistePreAmpliacion '" & psCtaCod & "','" & Format(pdFecha, "yyyymmdd") & "'"
'    Set oConecta = New COMConecta.DCOMConecta
'    oConecta.AbreConexion
'    Set R = oConecta.CargaRecordSet(sSql)
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'
'    If R!Cuenta > 0 Then
'    ExistePreAmpliacion = True
'    Else
'     ExistePreAmpliacion = False
'    End If
'    R.Close
'    Set R = Nothing
'    Exit Function
'
'ErrorPreAmpliacion:
'                  Err.Raise Err.Number, "Error En Proceso", Err.Description
'End Function
'Public Function ValidaRegIntDiferido(ByVal psCtaCod As String) As Boolean
'Dim sSql As String
'Dim nMovNroIntDifAmp, nMovNroIntDif As Long
'Dim oConecta As COMConecta.DCOMConecta
'Dim R As ADODB.Recordset
'
'    On Error GoTo ErrorRegIntDiferido
'    sSql = "exec stp_sel_ExisteRegIntDiferido '" & psCtaCod & "'"
'    Set oConecta = New COMConecta.DCOMConecta
'    oConecta.AbreConexion
'    Set R = oConecta.CargaRecordSet(sSql)
'
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'
'    If Not R.BOF And Not R.EOF Then
'    nMovNroIntDifAmp = R!nMovNroPreAmp
'    nMovNroIntDif = R!nMovNroPreAmp
'
'        If nMovNroIntDifAmp > 0 Then
'        ValidaRegIntDiferido = True
'        End If
'
'    Else
'    ValidaRegIntDiferido = False
'    End If
'
'
'
'    R.Close
'    Set R = Nothing
'    Exit Function
'
'ErrorRegIntDiferido:
'                  Err.Raise Err.Number, "Error En Proceso", Err.Description
'End Function
'Public Function RecuperaInteresDiferidoCont(ByVal psCtaCod As String) As ADODB.Recordset
'Dim sSql As String
'Dim oConecta As COMConecta.DCOMConecta
'
'    On Error GoTo ErrorInteresDiferidoCont
'     sSql = "exec STP_SEL_InteresDiferidoCont '" & psCtaCod & "'"
'    Set oConecta = New COMConecta.DCOMConecta
'    oConecta.AbreConexion
'    Set RecuperaInteresDiferidoCont = oConecta.CargaRecordSet(sSql)
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Function
'
'ErrorInteresDiferidoCont:
'    Err.Raise Err.Number, "Error En Proceso", Err.Description
'End Function
'Public Function EsCreditoAmpDesembolsado(ByVal psCtaCod As String) As Boolean
'Dim sSql As String
'Dim oConecta As COMConecta.DCOMConecta
'Dim R As ADODB.Recordset
'
'    On Error GoTo ErrorEsCreditoAmpDesembolsado
'    sSql = "exec stp_sel_VerificaSiEsAmplDesembolsado '" & psCtaCod & "'"
'    Set oConecta = New COMConecta.DCOMConecta
'    oConecta.AbreConexion
'    Set R = oConecta.CargaRecordSet(sSql)
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'
'    'If R!cCtaCod > 0 Then 'LUCV20180116, Comentó y agregó. desde IP:60
'    If (R.EOF And Not R.BOF) Then
'        EsCreditoAmpDesembolsado = True
'    Else
'        EsCreditoAmpDesembolsado = False
'    End If
'    R.Close
'    Set R = Nothing
'    Exit Function
'
'ErrorEsCreditoAmpDesembolsado:
'                  Err.Raise Err.Number, "Error En Proceso", Err.Description
'End Function
'Public Function ExtornaInteresDiferidoCont(ByVal psCtaCod As String) As ADODB.Recordset
'Dim sSql As String
'Dim oConecta As COMConecta.DCOMConecta
'
'    On Error GoTo ErrorExtornaInteresDiferidoCont
'     sSql = "exec stp_udp_ExtornaIntDiferidoCont '" & psCtaCod & "'"
'    Set oConecta = New COMConecta.DCOMConecta
'    oConecta.AbreConexion
'    Set ExtornaInteresDiferidoCont = oConecta.CargaRecordSet(sSql)
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Function
'
'ErrorExtornaInteresDiferidoCont:
'    Err.Raise Err.Number, "Error En Proceso", Err.Description
'End Function
'END VAPA from 60 END NAGL 20180609

'JOEP20171107 Acta201
Public Function ValidaPriFecPago(ByVal pdFechaDesembolso As Date, ByVal pdFechaPriPago As Date) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_sel_ValidaPrimeraFechaPago '" & Format(pdFechaDesembolso, "yyyymmdd") & "','" & Format(pdFechaPriPago, "yyyymmdd") & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ValidaPriFecPago = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
'JOEP20171107 Acta201

'JOEP20171107 3097-2017-GM Acta193
Public Function ValidadDestinoConsEmp(ByVal pnTpProducto As Integer, ByVal pnTpSubProducto As Integer, ByVal pnDestino As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_sel_Memo3097_ValidaDestino " & pnTpProducto & "," & pnTpSubProducto & "," & pnDestino & ""

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ValidadDestinoConsEmp = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
'JOEP20171107 3097-2017-GM Acta193

'JOEP20171222 ACTA 220-2017 Autorizacion Reprogramado
Public Function ObtieneNivApro(ByVal psNuevaCta As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    sSql = "exec stp_sel_NivAprobacion '" & psNuevaCta & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneNivApro = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function ObtieneCondiciones(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
        
    sSql = "exec stp_sel_ObtieneCondiciones '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneCondiciones = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'JOEP20171222 ACTA 220-2017 Autorizacion Reprogramado

'ARLO 20171127 ********************************************************************
Public Function VerificaCampania(ByVal psIdCampana As String) As Boolean
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As Recordset
Dim bRevisaDesemb As Boolean
On Error GoTo ErrorVerificaCampania
    
    sSql = " EXEC stp_sel_ERS0452017_VerificaCampania '" & psIdCampana & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If (rs.EOF And rs.BOF) Then
        VerificaCampania = False
    Else
        VerificaCampania = True
    End If
    Exit Function
ErrorVerificaCampania:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'**ARLO20171113 ERS070 - 2017
Public Sub ActualizaCredRevolventeNoRevolvente(pbActivo As Integer, pdFecha As Date, psCtaCod As String, psCod_Emp As String, psCtaContCod As String, pnSaldo As Currency, pnCuota_LNU As Currency)
Dim lsSQL As String
lsSQL = "EXEC stp_ins_CredRevolventeNoRevolvent " & pbActivo & ",'" & Format(pdFecha, "YYYY/MM/DD") & "', '" & psCtaCod & "', '" & psCod_Emp & "','" & psCtaContCod & "', " & pnSaldo & ", " & pnCuota_LNU
loConecta.Ejecutar (lsSQL)
End Sub
'**ARLO20171113 ERS070 - 2017

'NRLO20180306 ERS027-2017
Public Function ValidarVisitaDomiciliaria(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_sel_VerDomVerificarCreditoAprobacion '" & psCtaCod & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ValidarVisitaDomiciliaria = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
'NRLO20180306 ERS027-2017 END

'NRLO20180306 ERS027-2017
Public Function RegistrarVisitaDomiciliaria(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_ins_VerDomRegistrarCreditoSugerencia '" & psCtaCod & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set RegistrarVisitaDomiciliaria = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
'NRLO20180306 ERS027-2017 END

'NRLO20180306 ERS027-2017
Public Function ValidarVisitasMinimasPostDesembolso(ByVal psUser As String, ByVal psCargo As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_sel_VerDomVerificarVisitasMinPostDes '" & psUser & "', '" & psCargo & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ValidarVisitasMinimasPostDesembolso = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
'NRLO20180306 ERS027-2017 END

'***---> LUCV20180320, Agregó según ERS087-2017
Public Function CuotaFijaMensualPago(ByVal psCtaCod As String) As Double
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim rs As Recordset
    Dim bRevisaDesemb As Boolean
    On Error GoTo ErrorCuotaFijaMensualPago
    
    sSql = "EXEC stp_sel_ERS0872017_CuotaFijaMensualPago '" & psCtaCod & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    CuotaFijaMensualPago = IIf(IsNull(rs!nMontoCuota), 0, rs!nMontoCuota)
    
    Exit Function
ErrorCuotaFijaMensualPago:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'<---*** Fin LUCV20180320
'MARG ERS003-2018----------------------------------------------
Public Function existeDataExperianEnBD(cTipoid As String, cId As String, cApellido As String, cPersCodCliente As String, _
                                        Optional ByVal pbParamExtra As Boolean = False, _
                                        Optional ByVal pParamExtra As String = "") As Boolean
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As Recordset
On Error GoTo ErrorexisteDataExperianEnBD
    
    sSql = " EXEC stp_sel_DC_VerificarExisteDataExperian '" & cTipoid & "','" & cId & "','" & cApellido & "','" & cPersCodCliente & "'"
    
    'CTI1 20180809 ***
    If pbParamExtra Then
        sSql = sSql & ",1,'" & pParamExtra & "'"
    End If
    'CTI1 FIN ********
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If (rs.EOF And rs.BOF) Then
        existeDataExperianEnBD = False
    Else
        existeDataExperianEnBD = True
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
    
ErrorexisteDataExperianEnBD:
    Err.Raise Err.Number, "Error en existeDataExperianEnBD", Err.Description
End Function

Public Function esDataExperianActualizado(cTipoid As String, cId As String, cApellido As String, cPersCodCliente As String, _
                                        Optional ByVal pbParamExtra As Boolean = False, _
                                        Optional ByVal pParamExtra As String = "") As Boolean
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As Recordset
On Error GoTo ErrorexisteDataExperianEnBD
    
    sSql = " EXEC stp_sel_DC_VerificarEsDataExperianActualizado '" & cTipoid & "','" & cId & "','" & cApellido & "','" & cPersCodCliente & "'"
    
    'CTI1 20180809 ***
    If pbParamExtra Then
        sSql = sSql & ",1,'" & pParamExtra & "'"
    End If
    'CTI1 FIN ********
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If Not rs.BOF And Not rs.EOF Then
        esDataExperianActualizado = CBool(rs!nResultado)
    Else
        esDataExperianActualizado = False
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
    
ErrorexisteDataExperianEnBD:
    Err.Raise Err.Number, "Error en existeDataExperianEnBD", Err.Description
End Function

Public Function obtenerInforme(cTipoid As String, cId As String, cApellido As String, cPersCodCliente As String, _
                                Optional ByVal pbParamExtra As Boolean = False, _
                                Optional ByVal pParamExtra As String = "") As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = " EXEC stp_sel_DC_ObtenerInforme '" & cTipoid & "','" & cId & "','" & cApellido & "','" & cPersCodCliente & "'"
    
    'CTI1 20180809 ***
    If pbParamExtra Then
        sSql = sSql & ",1,'" & pParamExtra & "'"
    End If
    'CTI1 FIN ********
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set obtenerInforme = rs
    Set rs = Nothing
End Function
'Public Function obtenerScore(ByVal pnIdInforme As Integer) As ADODB.Recordset Comentado x JGPA20181029
Public Function obtenerScore(ByVal pnIdInforme As Long) As ADODB.Recordset 'Added by JGPA20181029 ACTA Nº 177-2018
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = " exec stp_sel_DC_ObtenerScore  " & pnIdInforme & ""
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set obtenerScore = rs
    Set rs = Nothing
End Function
Public Function obtenerCliente(ByVal pcPersCod As String, ByVal pcPersIDnro As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = " exec stp_sel_DC_ObtenerCliente '" & pcPersCod & "','" & pcPersIDnro & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set obtenerCliente = rs
    Set rs = Nothing
End Function
'Public Function RegistrarInformeCredito(pnIdInforme As Integer, pcCtaCod As String, pnPrdpersRelac As Integer) As Integer 'JGPA20181029
Public Function RegistrarInformeCredito(pnIdInforme As Long, pcCtaCod As String, pnPrdpersRelac As Integer) As Long 'Added by JGPA20181029 ACTA Nº 177-2018
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim oRs As New ADODB.Recordset

    On Error GoTo ErrorRegistrarInformeCredito

    sSql = " exec stp_ins_ERS0032018_RegistrarInformeCredito " & pnIdInforme & ",'" & pcCtaCod & "'," & pnPrdpersRelac
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set oRs = oCon.Ejecutar(sSql)
    If Not oRs.BOF And Not oRs.EOF Then
        If oRs.RecordCount > 0 Then
            RegistrarInformeCredito = oRs!nId
        Else
            RegistrarInformeCredito = 0
        End If
    Else
        RegistrarInformeCredito = 0
    End If
    
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
ErrorRegistrarInformeCredito:
    Err.Raise Err.Number, "RegistrarInformeCredito ", Err.Description
    RegistrarInformeCredito = 0
End Function
Public Sub EliminarInformeCredito(ByVal pcCtaCod As String)
    Dim lsSQL As String
    Dim coConex As COMConecta.DCOMConecta
    On Error GoTo ErrEliminarInformeCredito
    lsSQL = "EXEC stp_del_ERS0032018_EliminarInformeCredito '" & pcCtaCod & "'"
    Set coConex = New COMConecta.DCOMConecta
    coConex.AbreConexion
    coConex.Ejecutar (lsSQL)
    coConex.CierraConexion
    Set coConex = Nothing
    Exit Sub
ErrEliminarInformeCredito:
    Err.Raise Err.Number, "Error EliminarInformeCredito", Err.Description
End Sub
Public Function getDecisionScore(ByVal pcCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = " exec stp_sel_ERS0032018_getDecisionScore '" & pcCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set getDecisionScore = rs
    Set rs = Nothing
End Function
Public Function esAplicableValidacionScore(ByVal pcCtaCod As String) As Boolean
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As Recordset
On Error GoTo ErroresAplicableValidacionScore
    
    sSql = " exec stp_sel_ERS0032018_VerificarEsAplicableValidacionScore '" & pcCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If Not rs.BOF And Not rs.EOF Then
        esAplicableValidacionScore = CBool(rs!nResultado)
    Else
        esAplicableValidacionScore = False
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
    
ErroresAplicableValidacionScore:
    Err.Raise Err.Number, "Error en esAplicableValidacionScore", Err.Description
End Function
'END MARG------------------------------------------------------
'JIPR 22/06/2018
Public Sub InsertaPagoLoteDetalle(ByVal pnMovNro As String, ByVal psPersCodIns As String, ByVal pnTpFormaPago As Integer)
Dim SQL As String
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

    oCon.AbreConexion
        
        SQL = "exec stp_ins_PagoLoteDetalle " & pnMovNro & ",'" & psPersCodIns & "'," & pnTpFormaPago & ""

    oCon.ConexionActiva.Execute SQL
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Sub

oCon.CierraConexion
End Sub

Public Function RecuperaPagoLoteDetalle(ByVal pcMovNro As String, ByVal pcCodAge As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaProducto
    sSql = "exec stp_sel_PagoLoteDetalle '" & pcMovNro & "','" & pcCodAge & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaPagoLoteDetalle = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaProducto:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function
'JIPR 22/06/2018
'APRI20171212 ERS028-2017
Public Sub InsertaSegMultiriesgoDesgravamenTrama(ByVal psCtaCod As String)
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

    sSql = "exec stp_ins_SegMultiriesgoDesgravamenTrama '" & psCtaCod & "'"

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.Ejecutar (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
Public Sub ExtornaSegMultiriesgoDesgravamenTrama(ByVal psCtaCod As String)
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

    sSql = "exec stp_upd_ExSegMultiriesgoDesgravamenTrama '" & psCtaCod & "'"

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.Ejecutar (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'END APRI
'APRI20180620 ERS004-2018
Public Function ExisteSolicitudPendiente(ByVal pcPersCod As String) As Boolean
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oCon As New COMConecta.DCOMConecta
    ExisteSolicitudPendiente = False
    sSql = "exec stp_sel_ValidarConstNoAdeudoXPersona '" & pcPersCod & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    If Not rs.EOF Then
        If rs!nPendiente > 0 Then
            ExisteSolicitudPendiente = True
        End If
    End If
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'END APRI

'CTI1 20180809
Public Function ObtenerRelacionesFinalesExperian(ByVal pcCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oCon As New COMConecta.DCOMConecta

    sSql = "exec stp_sel_ObtenerRelacionesFinalesExperian '" & pcCtaCod & "'"
    
    oCon.AbreConexion
    Set ObtenerRelacionesFinalesExperian = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'CTI1 FIN
'**ARLO2080723 ERS037-2018
Public Function RecuperaSubProductosCrediticiosNEW(Optional ByVal psCab As String = "") As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta


    sSql = "Exec stp_sel_RecuperaSubProductosCrediticiosNEW '" & psCab & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaSubProductosCrediticiosNEW = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function
'**END ARLO

'JIPR 20180625
Public Sub InsertaFactElect(ByVal pnMovNro As String, ByVal pncCtaCod As String, ByVal pncPersCod As String, ByVal pnnTipoDoc As String, ByVal pndFechaPago As Date)
Dim SQL As String
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

    oCon.AbreConexion
        
    SQL = "exec stp_ins_FactElect " & pnMovNro & ",'" & pncCtaCod & "','" & pncPersCod & "','" & pnnTipoDoc & "','" & CStr(Format(pndFechaPago, "yyyymmdd")) & "' "

    oCon.ConexionActiva.Execute SQL
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Sub

oCon.CierraConexion
End Sub


Public Function RecuperaFactElectDetalle(ByVal pccCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    sSql = "exec stp_sel_FactElectDetalle '" & pccCtaCod & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
        
    Set RecuperaFactElectDetalle = rs
    Set rs = Nothing
    
End Function


Public Function GenerarCorrelativo(ByVal pnTipoDoc As String, ByVal pnAge As String, Optional ByVal pncCtaCod As String) As ADODB.Recordset 'ANPS
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    sSql = "exec stp_sel_ObtenerCorrelativo '" & pnTipoDoc & "','" & pnAge & "','" & pncCtaCod & "'" 'ANPS

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set GenerarCorrelativo = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function

Public Sub UpdateFactElect(ByVal pnMovNro As String, ByVal pSerie As String, ByVal pCorrelativo As String)
Dim SQL As String
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

    oCon.AbreConexion
        
    SQL = "exec sp_upd_CorrelativoFactElect  " & pnMovNro & ",'" & pSerie & "','" & pCorrelativo & "'"

    oCon.ConexionActiva.Execute SQL
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Sub

oCon.CierraConexion
End Sub

Public Sub UpdateEstadoCorrelativo(ByVal pnMovNro As String)
Dim SQL As String
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

    oCon.AbreConexion
        
    SQL = "exec stp_upd_EstCorrelativo  '" & pnMovNro & "'"

    oCon.ConexionActiva.Execute SQL
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Sub

oCon.CierraConexion
End Sub

Public Sub InsertarRegVentaFactElect(ByVal pnFecha As Date, ByVal pncPersCod As String, ByVal pncCtaCod As String, ByVal pnMovNro As String, ByVal pnMoneda As String, ByVal pnAge As String)
Dim SQL As String
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

    oCon.AbreConexion
        
    SQL = "exec sp_ins_RegVentaFactElect '" & CStr(Format(pnFecha, "yyyymmdd")) & "','" & pncPersCod & "','" & pncCtaCod & "','" & pnMovNro & "','" & pnMoneda & "','" & pnAge & "'"

    oCon.ConexionActiva.Execute SQL
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Sub

oCon.CierraConexion
End Sub


Public Sub VerificaInteresFactElect(ByVal nMovNro As String)
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion

    sSql = "stp_VerificaInteresFactElect '" & nMovNro & "'"
    oCon.ConexionActiva.Execute sSql
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Sub

    oCon.CierraConexion
    
End Sub
'JIPR 20180625
'MARG 20180618 Pag. Ant.-------
Public Function getDecisionPagoAnticipado(ByVal pcCtaCod As String, ByVal pnDiasAtraso As Integer, ByVal pbEsCancelacionTotal As Boolean, ByVal pbEsPagoAnticipado As Boolean, ByVal pbEsPrepararAmpliacion As Boolean) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    
    Dim rs As ADODB.Recordset

    sSql = " exec stp_sel_getDecisionPagoAnticipado '" & pcCtaCod & "'" & "," & pnDiasAtraso & "," & IIf(pbEsCancelacionTotal, 1, 0) & "," & IIf(pbEsPagoAnticipado, 1, 0) & "," & IIf(pbEsPrepararAmpliacion, 1, 0)

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set getDecisionPagoAnticipado = rs
    Set rs = Nothing
End Function

'END MARG----------------------

'INICIO EAAS20180516
Public Function PreparaAmpliacionGastosCero(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorPreparaAmpliacionGastosCero
     sSql = "exec stp_udp_PreparaAmpliacionGastosCero '" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set PreparaAmpliacionGastosCero = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorPreparaAmpliacionGastosCero:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'FIN EAAS20180516

'EAAS20180203 ERS054-2018***
Public Function RecuperaAguaSaneamiento(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrRecuperaAguaSaneamiento

    sSql = "EXEC stp_sel_ERS054_2018_ColocDestinoAguaSaneamiento '" & psCtaCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaAguaSaneamiento = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrRecuperaAguaSaneamiento:
    Err.Raise Err.Number, "Error En Proceso de Recuperar Agua y Saneamiento", Err.Description
End Function
Public Function RecuperaCargaSubDestinoAguaSaneamiento() As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrRecuperaCargaSubDestinoAguaSaneamiento

    sSql = "EXEC stp_sel_RecuperaCargaSubDestinoAguaSaneamiento_ERS_054_2018"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCargaSubDestinoAguaSaneamiento = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrRecuperaCargaSubDestinoAguaSaneamiento:
    Err.Raise Err.Number, "Error En Proceso de Recuperar Agua y Saneamiento", Err.Description
End Function
'RecuperaCargaSubDestinoAguaSaneamiento
'END EAAS *******
'INICIO EAAS20180815
Public Function ValidadDestinoConsEmpAguaSaneamiento(ByVal pnTpProducto As Integer, ByVal pnAguaSaneamiento As Integer, pCodigo, Optional pspnCuotas As Integer = -1) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_sel_ERS054_2018_ValidaDestinoAguaSaneamiento " & pnTpProducto & "," & pnAguaSaneamiento & ",'" & pCodigo & "'," & pspnCuotas & ""

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ValidadDestinoConsEmpAguaSaneamiento = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function

'INICIO EAAS20180815/CAMBIO DE CODIGO EN LA LOGICA SEGUN ACTA 147 EAAS20180904
Public Sub getMensajeValidacion(ByVal pnTpProducto As Integer, ByVal pnAguaSaneamiento As Integer, pCodigo As String, ByRef pmensaje As String, Optional pspnCuotas As Integer = -1)
    Dim oCon As New COMConecta.DCOMConecta
    On Error GoTo ErrgetMensajeValidacion
    Dim Cmd As New Command
    Dim Prm As ADODB.Parameter

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nTipoProducto", adInteger, adParamInput, , pnTpProducto) '0
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nDestino", adInteger, adParamInput, , pnAguaSaneamiento) '1
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cPersCod", adVarChar, adParamInput, 18, pCodigo) '2
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pspnCuotas", adInteger, adParamInput, , pspnCuotas) '3
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@mensaje", adVarChar, adParamOutput, 120) '4
    Cmd.Parameters.Append Prm
    
    oCon.AbreConexion

    Cmd.ActiveConnection = oCon.ConexionActiva
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "stp_sel_ERS054_2018_ValidaDestinoAguaSaneamiento_Solicitud"
    Cmd.Execute
    
    pmensaje = Cmd.Parameters(4).Value

    oCon.CierraConexion
    Set oCon = Nothing
        
  
    Exit Sub
ErrgetMensajeValidacion:
    Err.Raise Err.Number, Err.Source, Err.Description


End Sub
'FIN EAAS20180815/CAMBIO DE CODIGO EN LA LOGICA SEGUN ACTA 147 EAAS20180904

'FIN EAAS20180516
'INICIO EAAS20180827
Public Function GrabarSugerenciaAguaSaneamiento(ByRef pvListaAguasaneamiento As Variant, ByVal psCtaCod As String, pnMontoSol As Double) As ADODB.Recordset
Dim sSql As String
Dim Conn As COMConecta.DCOMConecta
Dim lvListaAguaSaneamiento() As TAguaSaneamiento
Dim ixCD As Integer
    On Error GoTo ErrorGrabarSugerenciaAguaSaneamiento
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    If UBound(pvListaAguasaneamiento) > 0 Then
        lvListaAguaSaneamiento = pvListaAguasaneamiento
    Else
        ReDim lvListaAguaSaneamiento(0)
    End If
    
    sSql = "EXEC stp_del_ERS054_2018_ColocDestinoAguaSaneamiento '" & psCtaCod & "'"
    Conn.ConexionActiva.Execute (sSql)
    Dim nSumaTotalAguaSaneamiento As Double
    Dim nSuma As Double
    nSuma = 0
    Dim nSumaTotalBeneficiados As Integer
    Dim nSumaB As Integer
    nSumaB = 0
    
    For ixCD = 1 To UBound(lvListaAguaSaneamiento)
        nSumaTotalAguaSaneamiento = nSumaTotalAguaSaneamiento + lvListaAguaSaneamiento(ixCD).nMontoS
        nSumaTotalBeneficiados = nSumaTotalBeneficiados + lvListaAguaSaneamiento(ixCD).nBeneficia
    Next
    
    If (UBound(lvListaAguaSaneamiento) > 0) Then
        For ixCD = 1 To 1
            sSql = "EXEC stp_ins_ERS054_2018_ColocDestinoAguaSaneamiento '" & psCtaCod & "'," & ixCD & "," & lvListaAguaSaneamiento(ixCD).nDestinoCod & "," & pnMontoSol & "," & nSumaTotalAguaSaneamiento & ", " & lvListaAguaSaneamiento(ixCD).nBeneficia
            Conn.ConexionActiva.Execute (sSql)
        Next
    End If
    For ixCD = 1 To UBound(lvListaAguaSaneamiento)
        sSql = "EXEC stp_ins_ERS054_2018_ColocDestinoAguaSaneamientoDet '" & psCtaCod & "'," & ixCD & "," & lvListaAguaSaneamiento(ixCD).nSubDestinoCod & "," & lvListaAguaSaneamiento(ixCD).nMontoS & " ," & lvListaAguaSaneamiento(ixCD).nBeneficia
        Conn.ConexionActiva.Execute (sSql)
    Next
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function

ErrorGrabarSugerenciaAguaSaneamiento:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'FIN EAAS20180827

'CROB20180813 ERS055-2018
Public Sub RegistrarCampanaMaximaAnual(ByVal pnCantMaxAnual As Byte, pcUserReg As String)
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

sSql = "EXEC stp_ins_RegistrarCampanasMaximaAnual " & pnCantMaxAnual & ",'" & pcUserReg & "'"

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
oCon.Ejecutar (sSql)
oCon.CierraConexion
Set oCon = Nothing
End Sub

Public Function ObtenerValorMaxCampanasAnual(ByVal psAnio As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_ObtenerValorMaxCampanasAnual '" & psAnio & "'"
    
    oCon.AbreConexion
    Set ObtenerValorMaxCampanasAnual = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ObtenerValorCantAdicionalDesctosCampana(ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_ObtenerValorCantAdicionalDesctosCampana '" & psPersCod & "'"
    
    oCon.AbreConexion
    Set ObtenerValorCantAdicionalDesctosCampana = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Sub RegistrarDescuentoCampanaAdicionalAnual(ByVal psPersCod As String, ByVal pnCantDescMax As Byte, ByVal psUsuarioReg As String)
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

sSql = "EXEC stp_ins_RegistrarDescuentoCampanaAdicionalAnual '" & psPersCod & "'," & pnCantDescMax & ",'" & psUsuarioReg & "'"

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
oCon.Ejecutar (sSql)
oCon.CierraConexion
Set oCon = Nothing
End Sub

Public Function RecuperarCantVecesAcogidasCampanasXcliente(ByVal pcCpersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "EXEC stp_sel_CampanaRecupCantVecesAcogidasXcliente '" & pcCpersCod & "'"
    
    oCon.AbreConexion
    Set RecuperarCantVecesAcogidasCampanasXcliente = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'CROB20180813 ERS055-2018

'**** 16102018
'*** CTI3 27082018
Public Function MostrarTodo(ByVal psNombre As String, Optional ByVal pnValor As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
         
        psNombre = "%" & psNombre & "%"
        sSql = "exec stp_sel_ObtenerSolicitudes '" & psNombre & "'," & pnValor & " "

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set MostrarTodo = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

'*** CTI3 27082018
Public Function MostrarAfps() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
         
    sSql = "exec stp_sel_afp"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set MostrarAfps = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function
'*** CIT3 27082018
Public Function BuscarCreditoAfp(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
         
    sSql = "exec stp_sel_BuscarSolicAfp '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set BuscarCreditoAfp = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

'*** CIT3 27082018
Public Function OperacionesVarios(ByVal psCtaCod As String, Optional psAfp As String, Optional pdFcarta As Date, Optional psDestino As String, Optional pnDisponible As Double, Optional pdFabono As Date, Optional ByVal pnValor As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
         
    If pnValor = 1 Then 'Insertar
        sSql = "exec stp_ins_CredSoliAfp '" & psCtaCod & "','" & psAfp & "','" & Format(pdFcarta, "yyyy-mm-dd") & "','" & psDestino & "'," & pnDisponible & ",'" & Format(pdFabono, "yyyy-mm-dd") & "'"
    ElseIf pnValor = 2 Then 'edicion
        sSql = "exec stp_upd_CredSoliAfp '" & psCtaCod & "','" & psAfp & "','" & Format(pdFcarta, "yyyy-mm-dd") & "','" & psDestino & "'," & pnDisponible & ",'" & Format(pdFabono, "yyyy-mm-dd") & "'"
    ElseIf pnValor = 3 Then 'eliminar
        sSql = "exec stp_del_CredSoliAfp '" & psCtaCod & "'"
    End If

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set OperacionesVarios = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function
'**************CTI3 *************** 11092018
Public Function RecuperaDatosAdicionales(ByVal psCtaCod As String, ByVal pnValor As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorRecuperaDatosFinan
      
    sSql = "exec stp_sel_ObtieneDatosPersonaJurDetalle '" & psCtaCod & "', " & pnValor & ""
       
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosAdicionales = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaDatosFinan:
    Err.Raise Err.Number, "Error En Proceso.", Err.Description
End Function
'********************************************

'CTI3 (ferimoro)
Public Function RecuperaAutoporExpo(ByVal pntRiesgo As Integer, ByVal psDestino As String, _
                                    ByVal psCampana As String, _
                                    ByVal pnEriesUni As Double, _
                                    ByVal psAgencia As String, _
                                    ByVal psCredito As String, _
                                    Optional ByVal pnTcliente As Integer, _
                                    Optional ByVal psAant As String, _
                                    Optional ByVal cNroCredito As String) As ADODB.Recordset
                                    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    sSql = "exec stp_sel_ObtieneNivelAutorizacionPorExposicion " & pntRiesgo & ",'" & psDestino & "',"
    sSql = sSql & "'" & psCampana & "'," & pnEriesUni & ",'" & psAgencia & "','" & psCredito & "',1,'" & psAant & "'"
    sSql = sSql & ", '" & cNroCredito & "'" 'JAOR ADD:cNroCredito
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaAutoporExpo = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

End Function

'*** CIT3 24082018
Public Function RecuperaCredito(ByVal psCtaCod As String, ByVal psDestino As String, ByVal pnValChk) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
         
    sSql = "exec stp_sel_CredHipoteca25porcAfp '" & psCtaCod & "','" & psDestino & "'," & pnValChk & ""
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function
'*** CIT3 24082018
Public Function RecuperaCctaReferencia(ByVal psCtaCod As String) As ADODB.Recordset
                                    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    sSql = "exec stp_sel_RecuperaRefinanciadosAprobados '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCctaReferencia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

End Function

'CTI3 - 29102018
Public Function validausuario(ByVal psUser As String) As ADODB.Recordset
                                    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    sSql = "exec stp_sel_verificausuario25AFP '" & psUser & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
    Set validausuario = oConecta.CargaRecordSet(sSql)
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

End Function

'CTI3 06112018 - Indica el tipo de Cliente para Duplicados
Public Function tpoPersonaNatJur(ByVal pscCtaCod As String) As ADODB.Recordset
                                    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    sSql = "exec stp_sel_personaNatJur '" & pscCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
    Set tpoPersonaNatJur = oConecta.CargaRecordSet(sSql)
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

End Function

'INICIO EAAS20181120 SEGUN ERS-072-2018
Public Function GetProtestosSinAclarar(ByVal psCta) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_sel_ERS0722018_GetProtestosSinAclarar '" & psCta & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set GetProtestosSinAclarar = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function

Public Function GetCobranzaCoactivaSunat(ByVal psCta) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_sel_ERS0722018_GetCobranzaCoactivaSunat '" & psCta & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set GetCobranzaCoactivaSunat = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function

Public Function ObtenerObligacionCerradaTDC(ByVal pnInforme) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_sel_DC_ERS072_ObtenerObligacionCerradaTDC '" & pnInforme & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtenerObligacionCerradaTDC = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function

'FIN EAAS20181120 SEGUN ERS-072-2018

'CTI2 20181206 *************************
'Public Function getSolicitudesPerdon(ByVal pcCtaCod As String, ByVal pnNroCalen As Integer, _
'                               ByVal pnnCuota As Integer) As ADODB.Recordset
'
'    Dim sSql As String
'
'    sSql = " exec stp_sel_SolicitudPerdon '" & pcCtaCod & "' , " & pnNroCalen & " , " & pnnCuota
'    Set getSolicitudes = coConex.CargaRecordSet(sSql)
'
'End Function
Public Function SetSolicitudPerdon(ByVal pcCtaCod As String, ByVal pnNroCalen As Integer, _
                                   ByVal pnnCuota As Integer, pcGlosaSolicitud As String, _
                                   ByVal pcGlosaAutorizacion As String, ByVal pbMora As Boolean, _
                                   ByVal pnMontoMora As Double, ByVal pnMontoMoraPerdon As Double, _
                                   ByVal pnMontoMoraPerdonPorc As Double, ByVal pbIntComp As Boolean, _
                                   ByVal pnMontoIntComp As Double, ByVal pnMontoIntCompPerdon As Double, _
                                   ByVal pnMontoIntCompPerdonPorc As Double, ByVal pbCapital As Boolean, _
                                   ByVal pnMontoCapital As Double, ByVal pnMontoCapitalPerdon As Double, _
                                   ByVal pcMovNroSolicitud As String, ByVal pcMovNroAutorizacion As String, _
                                   ByVal pcUsuarioSolicita As String, ByVal pcUsuarioAutoriza As String, _
                                   ByRef psObservacion As String) As Boolean

    Dim sSql As String
    Dim rsResult As ADODB.Recordset
    Dim bResultado As Boolean
    bResultado = False
       
    sSql = "exec stp_ins_SolicitudPerdon '" & _
            pcCtaCod & "' , " & _
            pnNroCalen & "," & _
            pnnCuota & ",'" & _
            pcGlosaSolicitud & "', '" & _
            pcGlosaAutorizacion & "', " & _
            IIf(pbMora, 1, 0) & ", " & _
            pnMontoMora & ", " & _
            pnMontoMoraPerdon & ", " & _
            pnMontoMoraPerdonPorc & "," & _
            IIf(pbIntComp, 1, 0) & ", " & _
            pnMontoIntComp & ", " & _
            pnMontoIntCompPerdon & "," & _
            pnMontoIntCompPerdonPorc & ", " & _
            IIf(pbCapital, 1, 0) & ", " & _
            pnMontoCapital & "," & _
            pnMontoCapitalPerdon & ",'" & _
            pcMovNroSolicitud & "', '" & _
            pcMovNroAutorizacion & "', '" & _
            pcUsuarioSolicita & "', '" & _
            pcUsuarioAutoriza & "'"
            
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rsResult = oConecta.CargaRecordSet(sSql)
    If Not rsResult Is Nothing Then
        If Not rsResult.EOF And Not rsResult.BOF Then
            If rsResult.RecordCount > 0 Then
                bResultado = rsResult!result
                psObservacion = rsResult!observ
            End If
        End If
    End If
    Set rsResult = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing
    SetSolicitudPerdon = bResultado
End Function
Sub CargaDatosSolicPerdon(ByRef rsCalSolic As ADODB.Recordset, ByVal sCuenta As String, ByVal nCalend As Integer)

Dim sSql As String
sSql = " exec stp_sel_RecuperaCalendarioSolicPerdon '" & sCuenta & "' , " & nCalend
Dim oConecta As COMConecta.DCOMConecta
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set rsCalSolic = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Sub
Sub ValidaSolicitudXcuota(ByVal cCtaCod As String, ByVal nNroCalen As Integer, ByVal nCuota As Integer, _
                          ByRef bFlagPagos As Boolean, ByRef bFlagSolicitud As Boolean)

Dim sSql As String
Dim rsCalSolic As ADODB.Recordset
sSql = " exec stp_sel_ValidaSolicitudDiaXcuota '" & cCtaCod & "' , " & nNroCalen & ", " & nCuota
bFlagPagos = False
bFlagSolicitud = False
Dim oConecta As COMConecta.DCOMConecta
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set rsCalSolic = oConecta.CargaRecordSet(sSql)
If Not rsCalSolic Is Nothing Then
    If Not rsCalSolic.EOF And Not rsCalSolic.BOF Then
        If rsCalSolic.RecordCount > 0 Then
            bFlagPagos = rsCalSolic!bFlagPagos
            bFlagSolicitud = rsCalSolic!bFlagSolicitud
        End If
    End If
End If
oConecta.CierraConexion
Set oConecta = Nothing

End Sub

'CTI2 END ******************************

'JOEP20180725 ERS034-2018
Public Function ValidadRigCambCred(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorValidadRigCambCred
    sSql = "exec stp_sel_ValidadRiesgoCambCred '" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ValidadRigCambCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorValidadRigCambCred:
    Err.Raise Err.Number, "Error, Validad Riesgo Cambiario Crediticio", Err.Description
End Function
Public Function ObtieneDatRigCambCred(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorObtieneDatRigCambCred
    sSql = "exec stp_sel_DatosPdfRiesgoCamb '" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneDatRigCambCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtieneDatRigCambCred:
    Err.Raise Err.Number, "Error, Datos Riesgo Cambiario Crediticio", Err.Description
End Function
Public Function GrabaRigCamCred(ByVal pcCtaCod As String, ByVal pnIngMN As Double, ByVal pnEgrMN As Double, ByVal pnIngME As Double, ByVal pnEgrME As Double) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim pcFechaEstado As String
Dim oBase As COMDCredito.DCOMCredActBD
Set oBase = New COMDCredito.DCOMCredActBD
       sSql = "exec stp_ins_GrabaFormIngrEgr '" & pcCtaCod & "'," & pnIngMN & "," & pnEgrMN & "," & pnIngME & "," & pnEgrME & ""
Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
Set GrabaRigCamCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
Set oConecta = Nothing
End Function
'JOEP20180725 ERS034-2018
'**ARLO20180712 ERS042 - 2018
Public Function GetResultadoCondicionCatalogo(ByVal cCodCondicion As String, ByVal cParametro As String) As Boolean
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String
Dim R As ADODB.Recordset
    sSql = "stp_sel_GetResultadoCondicionCatalogo '" & cCodCondicion & "','" & cParametro & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    If R.RecordCount > 0 Then
        If (R!bResultado = True) Then
            GetResultadoCondicionCatalogo = True
        Else
            GetResultadoCondicionCatalogo = False
        End If
    End If
    R.Close
    Set R = Nothing
End Function
'**ARLO20180712 ERS042 - 2018


'JOEP20180817 Catalogo
Public Function getCatalogoCombo(ByVal pcParProd As Integer, ByVal pnParCod As Long, Optional ByVal pnDest As Long = 0, Optional ByVal pnCond As Long = 0, Optional ByVal nIdCampana As Integer = 0) As ADODB.Recordset 'arlo20200429 add nIdCampana
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    sSql = " exec stp_sel_CatalogoCargaParametro " & pcParProd & "," & pnParCod & "," & pnDest & "," & pnCond & "," & nIdCampana & "" 'arlo20200429 add nIdCampana
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Set getCatalogoCombo = rs
    Set rs = Nothing
End Function
Public Function getCatalogoValidador(ByVal pcParProd As Integer, ByVal pnParCod As Long, ByVal pnMoneda As Integer, Optional ByVal nMonto As Double = 0, Optional ByVal nCuota As Integer = 0, Optional ByVal nPlazo As Integer = 0, Optional ByVal nDestino As Long = 0, Optional ByVal nSubDestino As Long = 0, Optional ByVal nDoc As Long = 0, Optional ByVal nTpIng As Long = 0, Optional ByVal cTpPla As String = "", Optional ByVal nModulo As Integer = 0, Optional ByVal nRefinanciado As Integer = -1, Optional ByVal nAmpliado As Integer = -1, Optional ByVal nIdCampana As Integer = 0) As ADODB.Recordset 'arlo20200430 add nIdCampana
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    sSql = " exec stp_sel_CatalogoValidador " & pcParProd & "," & pnParCod & "," & pnMoneda & "," & nMonto & "," & nCuota & "," & nPlazo & "," & nDestino & "," & nSubDestino & "," & nDoc & "," & nTpIng & ",'" & cTpPla & "'," & nModulo & "," & nRefinanciado & "," & nAmpliado & "," & nIdCampana & "" 'arlo20200430
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Set getCatalogoValidador = rs
    Set rs = Nothing
End Function
Public Function getCatalogoRangoFecha(ByVal pcCodProd As String, ByVal pnParCod As Long, ByVal pnTpPago As Integer, ByVal pdFechaPago As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    sSql = " exec stp_sel_CP_ValidaFecha '" & pcCodProd & "'," & pnParCod & "," & pnTpPago & ",'" & Format(pdFechaPago, "yyyymmdd") & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Set getCatalogoRangoFecha = rs
    Set rs = Nothing
End Function
Public Function getActivarCheckList(ByVal nCodProd As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    sSql = " exec stp_sel_Catalogo_getExisteCheckLitProd " & nCodProd & ""
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Set getActivarCheckList = rs
    Set rs = Nothing
End Function
Public Function ValidaIfExistReqCond(ByVal nParProd As Currency) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    sSql = " exec stp_sel_ExstReqCond " & nParProd & ""
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Set ValidaIfExistReqCond = rs
    Set rs = Nothing
End Function
Public Function ValidaRequisitos(ByVal cParProd As String, ByVal cCodPers As String, ByVal cCodPersRelc As String, ByVal cCondicion As String, Optional ByVal bAmpli As Long = 0, Optional ByVal nTpIngre As Long = 0, Optional ByVal nRefinanciado As Integer = -1, Optional ByVal nAmpliado As Integer = -1) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    sSql = " exec stp_sel_CatalogoValidadorRequisitos '" & cParProd & "','" & cCodPers & "','" & cCodPersRelc & "','" & cCondicion & "'," & bAmpli & "," & nTpIngre & "," & nRefinanciado & "," & nAmpliado & ""
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Set ValidaRequisitos = rs
    Set rs = Nothing
End Function
Public Function GetCatalogoParametro(ByVal nParCod As Currency) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    sSql = " exec stp_sel_GetCatalogoParametro " & nParCod & ""
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Set GetCatalogoParametro = rs
    Set rs = Nothing
End Function

Public Function CatalogoProCargaSubDestinos(ByVal cCodProd As String, ByVal nDestino As Integer, Optional ByVal cCodPers As String = "", Optional ByVal cCondicion As Integer = 0) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    sSql = " exec stp_sel_CatalogoProdCargaSubDestinos '" & cCodProd & "'," & nDestino & ",'" & cCodPers & "'," & cCondicion & ""
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Set CatalogoProCargaSubDestinos = rs
    Set rs = Nothing
End Function

Public Function CatalogoProDefaut(ByVal cCodProd As String, ByVal nParCod As Currency, Optional ByVal nMoneda As Integer = 0, Optional ByVal nTpDoc As Long = 0, Optional ByVal nTpOper As Long = 0, Optional ByVal nRefinanciado As Integer = -1, Optional ByVal nAmpliado As Integer = -1, Optional ByVal nIdCampana = 0) As ADODB.Recordset 'arlo20200429 add nIdCampana
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    sSql = " exec stp_sel_LLenaDefautControl '" & cCodProd & "'," & nParCod & "," & nMoneda & "," & nTpDoc & "," & nTpOper & "," & nRefinanciado & "," & nAmpliado & "," & nIdCampana & "" 'arlo20200429 add nIdCampana
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Set CatalogoProDefaut = rs
    Set rs = Nothing
End Function
Public Function CatalogoCondicionSeg(ByVal cCtaCod As String, ByVal cCodProd As String, ByVal nMonedad As String, ByVal pnMonto As String, ByVal pnCuota As String, ByVal pnPlazo As String, ByVal pnDestino As String, ByVal pnPeriodo As String, ByVal pnTpCredito As String, ByVal pnModulo As Integer, Optional ByVal lnCampanaId As Integer = 0) As ADODB.Recordset 'arlo20200429 add lnCampanaId
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    sSql = " exec stp_sel_CatalogoCondicionesSeg '" & cCtaCod & "','" & cCodProd & "'," & nMonedad & ",'" & Format(pnMonto, "#0.00") & "','" & pnCuota & "','" & pnPlazo & "','" & pnDestino & "','" & pnPeriodo & "','" & pnTpCredito & "'," & pnModulo & "," & lnCampanaId & "" 'arlo20200429
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Set CatalogoCondicionSeg = rs
    Set rs = Nothing
End Function
Public Function ValidaCondExtras(ByVal cParProd As String, ByVal cCodPers As String, ByVal cCodPersRelc As Integer, ByVal nCondicion As Integer, ByVal nCampana As Integer, ByVal nDestino As Integer, Optional ByVal cCtaAmpME As String = "", Optional ByVal pnTpDoc As Long = 0, Optional ByVal pnMoneda As Integer = 0, Optional ByVal pcCodAge As String = "", Optional ByVal nRefinanciado As Integer = -1, Optional ByVal nAmpliado As Integer = -1, Optional ByVal nIdCampana As Integer = 0) As ADODB.Recordset 'arlo add nIdCampana
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    sSql = " exec stp_sel_ValidaDatosExtras " & cParProd & ",'" & cCodPers & "','" & cCodPersRelc & "'," & nCondicion & "," & nCampana & "," & nDestino & ",'" & cCtaAmpME & "'," & pnTpDoc & "," & pnMoneda & ",'" & pcCodAge & "'," & nRefinanciado & "," & nAmpliado & "," & nIdCampana & "" 'arlo add nIdCampana
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Set ValidaCondExtras = rs
    Set rs = Nothing
End Function
Public Function ObtieneAporteCalifInter(ByVal cCodPers As String, ByVal cCodProd As String, Optional ByVal nSubDestino As Double = 0, Optional ByVal nDestino As Double = 0, Optional ByVal nCondicion As Double = 0, Optional ByVal nTpInteres As Double = 0, Optional ByVal cCtaCod As String = "", Optional ByVal nCampana As Integer = -1) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    sSql = " exec stp_sel_AporteClienteCalifInter '" & cCodPers & "','" & cCodProd & "'," & nSubDestino & "," & nDestino & "," & nCondicion & "," & nTpInteres & ",'" & cCtaCod & "'," & nCampana & ""
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Set ObtieneAporteCalifInter = rs
    Set rs = Nothing
End Function

Public Function CargaAporte(ByVal cCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    sSql = " exec stp_sel_CargaAporte '" & cCtaCod & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Set CargaAporte = rs
    Set rs = Nothing
End Function

Public Function ObtieneTpDoc(ByVal cCodProd As String, ByVal nParCod As Long) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    sSql = " exec stp_sel_ObtieneTpDoc '" & cCodProd & "'," & nParCod & ""
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Set ObtieneTpDoc = rs
    Set rs = Nothing
End Function
'CTI3 ERS003-2020
Public Function RSE_ObtieneTpDoc(ByVal cCodProd As String, ByVal nParCod As Long) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    sSql = " exec stp_RSE_sel_ObtieneTpDoc '" & cCodProd & "'," & nParCod & ""
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Set RSE_ObtieneTpDoc = rs
    Set rs = Nothing
End Function
Public Function RSE_ObtieneCreditoVigentexParalelo(ByVal cPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    sSql = " exec stp_RSE_sel_ObtieneCreditoVigentexParalelo '" & cPersCod & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Set RSE_ObtieneCreditoVigentexParalelo = rs
    Set rs = Nothing
End Function
'END
Public Sub CP_GrabaAporte(ByVal pcCtaCod As String, ByVal pnMonto As Currency, ByVal pnAporte As Currency, ByVal pnMontoSol As Currency, Optional ByVal pnMontoDisponible As Currency = 0)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    sSql = "exec stp_ins_GrabaDatosAporte '" & pcCtaCod & "'," & pnMonto & "," & pnAporte & "," & pnMontoSol & "," & pnMontoDisponible & ""
Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
Set oConecta = Nothing
End Sub

Public Sub CP_DeleteAporte(ByVal pcCtaCod As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    sSql = "exec stp_del_DeleteDatosAporte '" & pcCtaCod & "'"
Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
Set oConecta = Nothing
End Sub

Public Function CP_ObtieneAporte(ByVal pcCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    sSql = "exec stp_sel_ObtieneDatosAporte '" & pcCtaCod & "'"
Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
Set CP_ObtieneAporte = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
Set oConecta = Nothing
End Function

Public Function CP_IdentGarantiasCF(ByVal pcCtaCod As String, ByVal psNumagarant As String, ByVal psCodCargo As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    sSql = "exec stp_sel_Catalogo_IdentGarantia '" & pcCtaCod & "','" & psNumagarant & "','" & psCodCargo & "'"
Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
Set CP_IdentGarantiasCF = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
Set oConecta = Nothing
End Function

Public Function CP_getMensaje(ByVal pcCtaCod As String, ByVal psTpoCateg As String, ByVal psTpoProd As String, ByVal psItem As String, ByVal pnCantConf As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    sSql = "exec stp_sel_CatalogoCheckListObtieneMsgCondiciones '" & pcCtaCod & "','" & psTpoCateg & "','" & psTpoProd & "','" & psItem & "'," & pnCantConf & ""
Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
Set CP_getMensaje = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function CP_getDestinoRef(ByVal pcCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    sSql = "exec stp_sel_GetDestinoRef '" & pcCtaCod & "'"
Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
Set CP_getDestinoRef = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
Set oConecta = Nothing
End Function

Public Function GetMsgCheckList(ByVal psCodProd As String, ByVal pnIntem As Integer, ByVal pnCantConf As Integer) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrorGetMsgCheckList
    sSql = "EXEC stp_sel_CatalogoCheckListObtieneMsgAviso '" & psCodProd & "'," & pnIntem & "," & pnCantConf & ""
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set GetMsgCheckList = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorGetMsgCheckList:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function CP_getValCobGar(ByVal pcCtaCod As String) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrorCP_getValCobGar
    sSql = "EXEC stp_sel_ValidadRegCobGar '" & pcCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CP_getValCobGar = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorCP_getValCobGar:
    Err.Raise Err.Number, "Error en Validar Cobertura de Garantia", Err.Description
End Function
'JOEP20180817 Catalogo
'**ARLO20190308 ERS068 - 2018
Public Function getObtieneCuotasPagadasCreditoRefinan(ByVal pcCtaCod As String, ByVal pcTpoProdCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    sSql = "exec stp_sel_getObtieneCuotasPagadasCreditoRefinan '" & pcCtaCod & "','" & pcTpoProdCod & "'"
Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
Set getObtieneCuotasPagadasCreditoRefinan = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ValidaPropuestaRefinanciado(ByVal pcPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    sSql = "exec STP_SEL_VALIDAPROPUESTAREFINANCIADO '" & pcPersCod & "'"
Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
Set ValidaPropuestaRefinanciado = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ValidaCreditoRefinanciarAmpliado(ByVal pcPersCod As String, ByVal pnMonto As Double) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    sSql = "exec STP_SEL_VALIDACREDITOREFINANCIARAMPLIADO '" & pcPersCod & "'," & pnMonto & ""
Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
Set ValidaCreditoRefinanciarAmpliado = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function obtenerVistoBueno90Dias(ByVal pcCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    sSql = "exec STP_SEL_OBTENERVISTOBUENO90DIAS '" & pcCtaCod & "'"
Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
Set obtenerVistoBueno90Dias = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
Set oConecta = Nothing
End Function
'**ARLO END
'INICIO EAAS20191401 SEGUN 018-GM-DI_CMACM
Public Function ValidadDestinoConsEmpCreditoVerde(ByVal pnTpProducto As Integer, ByVal pnAguaSaneamiento As Integer, pCodigo As String, Optional pspnCuotas As Integer = -1) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_sel_ERS054_2018_ValidaDestinoCreditoVerde " & pnTpProducto & "," & pnAguaSaneamiento & ",'" & pCodigo & "'," & pspnCuotas & ""

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ValidadDestinoConsEmpCreditoVerde = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function

Public Function GrabarSugerenciaCreditoVerde(ByRef pvListaCreditoVerde As Variant, ByVal psCtaCod As String, pnMontoSol As Double) As ADODB.Recordset
Dim sSql As String
Dim Conn As COMConecta.DCOMConecta
Dim lvListaCreditoVerde() As TCreditoVerde
Dim ixCD As Integer
    On Error GoTo ErrorGrabarSugerenciaCreditoVerde
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    If UBound(pvListaCreditoVerde) > 0 Then
        lvListaCreditoVerde = pvListaCreditoVerde
    Else
        ReDim lvListaCreditoVerde(0)
    End If
    
    sSql = "EXEC stp_del_ERS054_2018_ColocDestinoCreditoVerde '" & psCtaCod & "'"
    Conn.ConexionActiva.Execute (sSql)
    Dim nSumaTotalCreditoVerde As Double
    Dim nSuma As Double
    nSuma = 0
    Dim nSumaTotalBeneficiados As Integer
    Dim nSumaB As Integer
    nSumaB = 0
    
    For ixCD = 1 To UBound(lvListaCreditoVerde)
        nSumaTotalCreditoVerde = nSumaTotalCreditoVerde + lvListaCreditoVerde(ixCD).nMontoS
        nSumaTotalBeneficiados = nSumaTotalBeneficiados + lvListaCreditoVerde(ixCD).nBeneficia
    Next
    
    If (UBound(lvListaCreditoVerde) > 0) Then
        For ixCD = 1 To 1
            sSql = "EXEC stp_ins_ERS054_2018_ColocDestinoCreditoVerde '" & psCtaCod & "'," & ixCD & "," & lvListaCreditoVerde(ixCD).nDestinoCod & "," & pnMontoSol & "," & nSumaTotalCreditoVerde & ", " & lvListaCreditoVerde(ixCD).nBeneficia
            Conn.ConexionActiva.Execute (sSql)
        Next
    End If
    For ixCD = 1 To UBound(lvListaCreditoVerde)
        sSql = "EXEC stp_ins_ERS054_2018_ColocDestinoCreditoVerdeDet '" & psCtaCod & "'," & ixCD & "," & lvListaCreditoVerde(ixCD).nSubDestinoCod & "," & lvListaCreditoVerde(ixCD).nMontoS & " ," & lvListaCreditoVerde(ixCD).nBeneficia
        Conn.ConexionActiva.Execute (sSql)
    Next
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function

ErrorGrabarSugerenciaCreditoVerde:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function


Public Function RecuperaCreditoVerde(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrRecuperaCreditoVerde

    sSql = "EXEC stp_sel_018GMDI_2018_ColocDestinoCreditoVerde '" & psCtaCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditoVerde = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrRecuperaCreditoVerde:
    Err.Raise Err.Number, "Error En Proceso de Recuperar Crédito Verde", Err.Description
End Function
Public Function RecuperaCargaSubDestinoCreditoVerde() As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrRecuperaCargaSubDestinoCreditoVerde

    sSql = "EXEC stp_sel_RecuperaCargaSubDestinoCreditoVerde"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCargaSubDestinoCreditoVerde = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrRecuperaCargaSubDestinoCreditoVerde:
    Err.Raise Err.Number, "Error En Proceso de Recuperar Agua y Saneamiento", Err.Description
End Function
'FIN EAAS20191401

'ACTA Nº 113 - 2019 - JOEP - SORO - Asignacion de Agencias para Pre-Desembolsos
Public Function AdmCred_CargaAgencia() As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta
On Error GoTo ErrAdmCred_CargaAgencia
    sSql = "EXEC stp_sel_RS_ListaAgencias "
    oCon.AbreConexion
    Set AdmCred_CargaAgencia = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
ErrAdmCred_CargaAgencia:
    Err.Raise Err.Number, "Error al obtener Agencias ", Err.Description
End Function

Public Function AdmCred_AsignaAgencia(ByVal pcAgeAsignados As String, ByVal pcAgeAgrega As String) As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta
On Error GoTo ErrAdmCred_AsignaAgencia
    sSql = "EXEC stp_sel_AgenciaAsignarAgregar '" & pcAgeAsignados & "','" & pcAgeAgrega & "'"
    oCon.AbreConexion
    Set AdmCred_AsignaAgencia = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
ErrAdmCred_AsignaAgencia:
    Err.Raise Err.Number, "Error al Cargar Busqueda de Agencias ", Err.Description
End Function

Public Function AdmCred_BusAgencia(ByVal pcUser As String, ByVal pnNew As Integer) As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta
On Error GoTo ErrAdmCred_BusAgencia
    sSql = "EXEC stp_sel_AdmCredAsignarUserPreDes '" & pcUser & "'," & pnNew & ""
    oCon.AbreConexion
    Set AdmCred_BusAgencia = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
ErrAdmCred_BusAgencia:
    Err.Raise Err.Number, "Error al Cargar Busqueda de Agencias ", Err.Description
End Function

Public Function AdmCred_ValidaCargo(ByVal pcUser As String) As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta
On Error GoTo ErrAdmCred_ValidaCargo
    sSql = "EXEC stp_sel_ValidaCargo '" & pcUser & "'"
    oCon.AbreConexion
    Set AdmCred_ValidaCargo = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
ErrAdmCred_ValidaCargo:
    Err.Raise Err.Number, "Error al Validad Cargo ", Err.Description
End Function

Public Function AdmCred_ValidaUserReg(ByVal pcUserSist As String) As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta
On Error GoTo ErrAdmCred_ValidaCargo
    sSql = "EXEC stp_sel_ValidaUserRegistro '" & pcUserSist & "'"
    oCon.AbreConexion
    Set AdmCred_ValidaUserReg = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
ErrAdmCred_ValidaCargo:
    Err.Raise Err.Number, "Error al Validad Cargo ", Err.Description
End Function

Public Function AdmCred_RegistraUserAsigPreDes(ByVal cAgeCod_Actual As String, ByVal pcAgeCod_GAsignados As String, ByVal pcMovNro As String) As Boolean
Dim lsSQL As String
Dim rs As ADODB.Recordset
Dim oCon As New COMConecta.DCOMConecta
On Error GoTo ErrAdmCred_RegistraUserAsigPreDes
    lsSQL = "EXEC stp_ins_Registrar '" & cAgeCod_Actual & "','" & pcAgeCod_GAsignados & "','" & pcMovNro & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(lsSQL)
If rs!Okey = 1 Then
    AdmCred_RegistraUserAsigPreDes = True
Else
    AdmCred_RegistraUserAsigPreDes = False
End If
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
ErrAdmCred_RegistraUserAsigPreDes:
    Err.Raise Err.Number, "Error al registar Usuarios Asignados para Autorizar Pre-Desembolsos", Err.Description
End Function
'ACTA Nº 113 - 2019 - JOEP - SORO - Asignacion de Agencias para Pre-Desembolsos

Public Function ObtieneHistorialCuotasCond(ByVal psCtaCod As String, ByVal pdFecha As Date) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtieneHistorialCuotasCond
    sSql = "Exec stp_sel_ObtieneHistorialCuotasCond '" & psCtaCod & "','" & Format(pdFecha, "yyyymmdd") & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneHistorialCuotasCond = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorObtieneHistorialCuotasCond:
    Err.Raise Err.Number, "Error En Proceso de Historial en Condonación del Crédito en cuestión", Err.Description
End Function 'NAGL 20190716 Según ERS036-2018

Public Function GetOpeCondonacionxEstado(ByVal psCtaCod As String, ByVal psOpeCond As String, Optional pnMovNro As Long = 0) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorGetOpeCondonacionxEstado
    Set oConecta = New COMConecta.DCOMConecta
    sSql = "Exec stp_sel_GetOpeCondonacionxEstado '" & psCtaCod & "'," & pnMovNro & ""
    oConecta.AbreConexion
    Set GetOpeCondonacionxEstado = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorGetOpeCondonacionxEstado:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function 'NAGL 20190716 Según ERS036-2018

Public Function esAplicableCuotaAPerdonar(ByVal psCtaCod As String, ByVal pnNroCuotaPerd As Integer) As Boolean
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String
Dim RPerd As ADODB.Recordset
    sSql = "Exec stp_sel_VerificaCuotaProxCondonar '" & psCtaCod & "'," & pnNroCuotaPerd & ""
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RPerd = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    If RPerd!cAplicaPerdon = "1" Then
        esAplicableCuotaAPerdonar = True
    Else
        esAplicableCuotaAPerdonar = False
    End If
    RPerd.Close
    Set RPerd = Nothing
End Function 'NAGL 20190716 Según ERS036-2018

Public Sub AgregaRegCredPerdonMora(ByVal pnMovNro As Long, ByVal pcCtaCod As String, ByVal pnNroCalen As Integer, ByVal pnNroCuota As Integer, ByVal pnMora As Double, _
        ByVal pnMoraPagadaAnt As Double, ByVal pnMoraAPerdonar As Double, ByVal pnMoraAPerdonarPorc As Double, ByVal pnMoraPerdon As Double, ByVal pnMoraPagada As Double, ByVal psMovNro As String, ByVal psOpeTpo As String, ByVal psGlosa As String, ByVal pnBit As Boolean)
        
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
On Error GoTo ErrorAgregaRegCredPerdonMora
    sSql = "Exec stp_ins_AgregarRegCredPerdonMora " & pnMovNro & ",'" & pcCtaCod & "'," & pnNroCalen & "," & pnNroCuota & "," & pnMora & ", " & pnMoraPagadaAnt & "," & pnMoraAPerdonar & "," & pnMoraAPerdonarPorc & ", " & pnMoraPerdon & "," & pnMoraPagada & ", '" & psMovNro & "','" & psOpeTpo & "','" & psGlosa & "', " & IIf(pnBit = True, 1, 0) & ""
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorAgregaRegCredPerdonMora:
    Err.Raise Err.Number, "Agrega Registro en Perdón de Mora", Err.Description
End Sub 'NAGL 20190716 Según ERS036-2018
'arlo20191120 Begin
Public Sub InsertarVBCuotas6MinimasRiesgos(ByVal pcPersCod As String, ByVal psAgeCod As String, ByVal pnMoneda As Integer, ByVal pnMontoSol As Currency, ByVal pnCuotasPagadas As Integer, ByVal pnTotalCuotas As Integer, ByVal psMovNro As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    sSql = "exec stp_ins_InsertarVBCuotas6MinimasRiesgos '" & pcPersCod & "','" & psAgeCod & "'," & pnMoneda & "," & pnMontoSol & "," & pnCuotasPagadas & "," & pnTotalCuotas & ",'" & psMovNro & "'"
Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
Set oConecta = Nothing
End Sub
Public Function getObtieneMensajeVB6Cuotas(ByVal pcPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    sSql = "exec stp_sel_getObtieneMensajeVB6Cuotas '" & pcPersCod & "'"
Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
Set getObtieneMensajeVB6Cuotas = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
Set oConecta = Nothing
End Function
'arlo20191120 End
'marg 201906
Public Function IsConsultaExperian() As Boolean
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As Recordset
On Error GoTo ErroresIsConsultaExperian
    
    sSql = " exec stp_sel_ObtieneExperianSwitch"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If Not rs.BOF And Not rs.EOF Then
        IsConsultaExperian = CBool(rs!bConsultarExperian)
    Else
        IsConsultaExperian = False
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
    
ErroresIsConsultaExperian:
    Err.Raise Err.Number, "Error en IsConsultaExperian", Err.Description
End Function

Public Function IsAprobacionSinScore() As Boolean
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As Recordset
On Error GoTo ErroresIsAprobacionSinScore
    
    sSql = " exec stp_sel_ObtieneExperianSwitchScore"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If Not rs.BOF And Not rs.EOF Then
        IsAprobacionSinScore = CBool(rs!bAprobarSinScore)
    Else
        IsAprobacionSinScore = False
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
    
ErroresIsAprobacionSinScore:
    Err.Raise Err.Number, "Error en IsAprobacionSinScore", Err.Description
End Function

Public Function EsAplicableValidacionScoreManual(ByVal pcCtaCod As String) As Boolean
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As Recordset
On Error GoTo ErroresEsAplicableValidacionScoreManual
    
    sSql = " exec stp_sel_ERS0032018_VerificarEsAplicableValidacionScoreManual '" & pcCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If Not rs.BOF And Not rs.EOF Then
        EsAplicableValidacionScoreManual = CBool(rs!nResultado)
    Else
        EsAplicableValidacionScoreManual = False
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
    
ErroresEsAplicableValidacionScoreManual:
    Err.Raise Err.Number, "Error en esAplicableValidacionScoreManual", Err.Description
End Function

Public Function getDecisionScoreManual(ByVal pcCtaCod As String, ByVal pnFormulario As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = " exec stp_sel_getDecisionScoreManual '" & pcCtaCod & "'," & pnFormulario
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set getDecisionScoreManual = rs
    Set rs = Nothing
End Function

Public Sub InsertarCreditoSinScore(ByVal pcCtaCod As String, ByVal pcUser As String)
    Dim lsSQL As String
    Dim coConex As COMConecta.DCOMConecta
    On Error GoTo ErrInsertarCreditoSinScore
    lsSQL = "EXEC stp_ins_InsertarCreditosSinScore '" & pcCtaCod & "','" & pcUser & "'"
    Set coConex = New COMConecta.DCOMConecta
    coConex.AbreConexion
    coConex.Ejecutar (lsSQL)
    coConex.CierraConexion
    Set coConex = Nothing
    Exit Sub
ErrInsertarCreditoSinScore:
    Err.Raise Err.Number, "Error InsertarCreditoSinScore", Err.Description
End Sub
'end marg

' RIRO 20200611 ****
Public Function EsReactiva(ByVal cCtaCod As String) As Boolean
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String
Dim R As ADODB.Recordset
    
    sSql = "stp_sel_EsReactiva '" & cCtaCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    If R.RecordCount > 0 Then
        If (R!bEsReactiva = True) Then
            EsReactiva = True
        Else
            EsReactiva = False
        End If
    End If
    R.Close
    Set R = Nothing
End Function
' END RIRO ***************
'JOEP20200705 Comabio ReactivaCovid
Public Function RestrincionCovidReact(ByVal psCtaCod As String, ByVal pcAgeCod As String, ByVal pcModulo As String, Optional ByVal pnCancelar As Integer = 0, Optional ByVal pnAdelantoCuota As Integer = 0, Optional ByVal pnPagoAnticipado As Integer = 0) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
        
    sSql = "exec stp_sel_RestrincionesPagoReactCovid '" & psCtaCod & "','" & pcAgeCod & "'," & pcModulo & "," & pnCancelar & "," & pnAdelantoCuota & "," & pnPagoAnticipado & ""
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RestrincionCovidReact = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'JOEP20200705 Comabio ReactivaCovid
'arlo20200716
Public Function ValidadCreditoRefinanciadoTitular(ByVal pcCtaCod As String, ByVal pcPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    sSql = "exec stp_sel_ValidadCreditoRefinanciadoTitular '" & pcCtaCod & "','" & pcPersCod & "'"
Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
Set ValidadCreditoRefinanciadoTitular = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
Set oConecta = Nothing
End Function
'arlo end

'JOEP20200705 Comabio ReactivaCovid
Public Function MsgBoxReprog(ByVal psCtaCod As String, ByVal pnValorCombo As Integer) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
        
    sSql = "exec stp_sel_MsgBoxReprogramacion '" & psCtaCod & "'," & pnValorCombo & ""
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set MsgBoxReprog = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function ReprogramacionValidaDatosCovid(ByVal psCtaCod As String, ByVal pnDiasReprog As Integer, ByVal pcFechaSol As String, ByVal pcFecCuotaVenc As String, ByVal pcMotivo As String, ByVal pnDiasAtraso As Integer) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
        
    'sSql = "exec stp_sel_ReprogramacionValidaDatosCovid '" & psCtaCod & "'," & pnDiasReprog & ",'" & Format(pcFechaSol, "yyyyMMdd") & "','" & Format(pcFecCuotaVenc, "yyyyMMdd") & "','" & pcMotivo & "'," & pnDiasAtraso & "" RIRO20200726 Comentado
    sSql = "exec stp_sel_ReprogramacionValidaDatosCovid '" & psCtaCod & "'," & pnDiasReprog & ",'" & pcFechaSol & "','" & pcFecCuotaVenc & "','" & pcMotivo & "'," & pnDiasAtraso & "" ' RIRO20200726 Comentado
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ReprogramacionValidaDatosCovid = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'JOEP20200705 Comabio ReactivaCovid
'add jhcu 12-09-2020
Public Function varConsReversion(ByVal psCtaCod As String, Optional ByVal psCuser As String, Optional ByVal psCodAge As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim i As Integer
    Dim sCadProd As String
    Dim sCadAge As String

    On Error GoTo ErrorVarConsReversion


    sSql = " exec varSP_sol_reversion '" & psCtaCod & "','" & psCuser & "','" & psCodAge & "'"
        

        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set varConsReversion = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing

        Exit Function
ErrorVarConsReversion:
    Err.Raise Err.Number, "Error En Proceso (varConsReversion)", Err.Description

End Function
'fin jhcu


'JOEP202008 Propuesta reprogramacion Actividad economica
Public Function ReprogramacionObtieneDatosIfis(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    sSql = "Exec stp_sel_ReprogramadoObtieneIfis '" & psCtaCod & "'"
    oCon.AbreConexion
    Set ReprogramacionObtieneDatosIfis = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ReprogramacionValDtsPropuesta(ByVal pcCtaCod As String, ByVal fnTipo As Integer, ByVal pcComentario As String, _
ByVal OpNF_S As Integer, ByVal OpNF_n As Integer, _
ByVal OpCA_S As Integer, ByVal OpCA_n As Integer, _
ByVal pcNewComentario As String, _
ByVal OpIngrAct_incr As Integer, ByVal OpIngrAct_mant As Integer, ByVal OpIngrAct_dis As Integer, _
ByVal OPDimicEs_temp As Integer, ByVal OPDimicEs_perm As Integer, _
ByVal OpDismPorc_has As Integer, ByVal OpDismPorc_hasel As Integer, ByVal OpDismPorc_NHN As Integer, _
ByVal OpIngActPer_inccuot As Integer, ByVal OpIngActPer_mant As Integer, ByVal OpIngActPer_dis As Integer, _
ByVal IngrVentNet As Currency, ByVal GastoTot As Currency, ByVal Exce As Currency, ByVal pMatzIfis As Integer, _
ByVal pModalidad As Integer, ByVal pValOcmSicm As Integer, Optional ByVal RezOtr As Integer = -1, Optional ByVal DiasNew As Integer = -1) As ADODB.Recordset

Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    sSql = "Exec spt_sel_reprogramacionPropuestaValiDts '" & pcCtaCod & "'," & fnTipo & ",'" & pcComentario & "'," _
    & OpNF_S & "," & OpNF_n & "," _
    & OpCA_S & "," & OpCA_n & ",'" _
    & pcNewComentario & "'," _
    & OpIngrAct_incr & "," & OpIngrAct_mant & "," & OpIngrAct_dis & "," _
    & OPDimicEs_temp & "," & OPDimicEs_perm & "," _
    & OpDismPorc_has & "," & OpDismPorc_hasel & "," & OpDismPorc_NHN & "," _
    & OpIngActPer_inccuot & "," & OpIngActPer_mant & "," & OpIngActPer_dis & "," _
    & IngrVentNet & "," & GastoTot & "," & Exce & "," & pMatzIfis & "," & pModalidad & "," & pValOcmSicm & "," & RezOtr & "," & DiasNew & ""
        
    oCon.AbreConexion
    Set ReprogramacionValDtsPropuesta = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'JOEP202008 Propuesta reprogramacion Actividad economica

'JOEP202008 Tasa ESpecial y reduccion de monto
'Public Function ReprogramacionObtTasaEspecial(ByVal pcCtaCod As String, ByVal pnSaldoCap As Currency) As ADODB.Recordset
'Public Function ReprogramacionObtTasaEspecial(ByVal pcCtaCod As String, ByVal pnSaldoCap As Currency, Optional ByVal pnMenuOpcion As Integer = -1) As ADODB.Recordset 'Add JOEP20210306 garantia covid
Public Function ReprogramacionObtTasaEspecial(ByVal pcCtaCod As String, ByVal pnSaldoCap As Currency, Optional ByVal pnMenuOpcion As Integer = -1, Optional ByVal nOpCovid As Integer = -1) As ADODB.Recordset 'Add JOEP20210310 identificar facilidad de reprogramacion
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta
                              
    'sSql = "Exec stp_sel_ReprogramacionObtieneTasaEspecial '" & pcCtaCod & "'," & pnSaldoCap & ""
    'sSql = "Exec stp_sel_ReprogramacionObtieneTasaEspecial '" & pcCtaCod & "'," & pnSaldoCap & "," & pnMenuOpcion & "" 'Add JOEP20210306 garantia covid
    sSql = "Exec stp_sel_ReprogramacionObtieneTasaEspecial '" & pcCtaCod & "'," & pnSaldoCap & "," & pnMenuOpcion & "," & nOpCovid & "" 'Add JOEP20210310 identificar facilidad de reprogramacion
    oCon.AbreConexion
    Set ReprogramacionObtTasaEspecial = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

'Public Function ReprogramacionObtMantenerCuota(ByVal pcCtaCod As String) As ADODB.Recordset
Public Function ReprogramacionObtMantenerCuota(ByVal pcCtaCod As String, Optional ByVal pnMenuOpcion As Integer = -1, Optional ByVal nOpCovid As Integer = -1) As ADODB.Recordset 'Add JOEP20210310 identificar facilidad de reprogramacion
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    'sSql = "Exec stp_sel_ReprogramacionObtieneMantenerCuota '" & pcCtaCod & "'"
    sSql = "Exec stp_sel_ReprogramacionObtieneMantenerCuota '" & pcCtaCod & "'," & pnMenuOpcion & "," & nOpCovid & "" 'Add JOEP20210310 identificar facilidad de reprogramacion
    oCon.AbreConexion
    Set ReprogramacionObtMantenerCuota = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

'Public Function ReprogramacionValidDatos(ByVal pcCtaCod As String, ByVal pcTxtGlosa As String, ByVal pnMontoCuota As Currency, ByVal pnCapitalReprg As Currency, ByVal pnComboEmerg As Integer, ByVal pnComboCovid As Integer, ByVal nOpCovid As Integer, ByVal nMontoCuotaCalend As Currency, ByVal nTasaEspecial As Currency, ByVal nTasa As Currency, ByVal pdFechaVencReprg As String, ByVal pnBtn As Integer) As ADODB.Recordset
Public Function ReprogramacionValidDatos(ByVal pcCtaCod As String, ByVal pcTxtGlosa As String, ByVal pnMontoCuota As Currency, ByVal pnCapitalReprg As Currency, ByVal pnComboEmerg As Integer, ByVal pnComboCovid As Integer, ByVal nOpCovid As Integer, ByVal nMontoCuotaCalend As Currency, ByVal nTasaEspecial As Currency, ByVal nTasa As Currency, ByVal pdFechaVencReprg As String, ByVal pnBtn As Integer, Optional pnMontNegativo As Integer = 0) As ADODB.Recordset 'Add JOEP20210306 garantia covid
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    'sSql = "Exec stp_sel_ReprogramacionValdDatsReprog '" & pcCtaCod & "','" & pcTxtGlosa & "'," & pnMontoCuota & "," & pnCapitalReprg & "," & pnComboEmerg & "," & pnComboCovid & "," & nOpCovid & "," & nMontoCuotaCalend & "," & nTasaEspecial & "," & nTasa & ",'" & Format(pdFechaVencReprg, "yyyyMMdd") & "'," & pnBtn & ""
    sSql = "Exec stp_sel_ReprogramacionValdDatsReprog '" & pcCtaCod & "','" & pcTxtGlosa & "'," & pnMontoCuota & "," & pnCapitalReprg & "," & pnComboEmerg & "," & pnComboCovid & "," & nOpCovid & "," & nMontoCuotaCalend & "," & nTasaEspecial & "," & nTasa & ",'" & Format(pdFechaVencReprg, "yyyyMMdd") & "'," & pnBtn & "," & pnMontNegativo & "" 'Add JOEP20210306 garantia covid
    oCon.AbreConexion
    Set ReprogramacionValidDatos = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ReprogramacionOCM(ByVal pcCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    sSql = "Exec stp_sel_ReprogramacionWeb '" & pcCtaCod & "'"
    oCon.AbreConexion
    Set ReprogramacionOCM = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ReprogramacionValDtsAprobacion(ByVal pcCtaCod As String, ByVal fnTipo As Integer, ByVal pGlosa As String) As ADODB.Recordset

Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    sSql = "Exec stp_sel_ReprogramacionValidaDatosAprobacion '" & pcCtaCod & "'," & fnTipo & ",'" & pGlosa & "'"
    oCon.AbreConexion
    Set ReprogramacionValDtsAprobacion = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ReprogramacionPropuestaObtieneModalidades(ByVal psCtaCod As String, ByVal pdFecha As Date) As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    sSql = "Exec stp_sel_ReprogPropuestaObtieneModalidades '" & psCtaCod & "','" & Format(pdFecha, "yyyyMMdd") & "'"
    oCon.AbreConexion
    Set ReprogramacionPropuestaObtieneModalidades = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

'Public Function SimuladoReprogramacion(ByVal pcCtaCod As String) As ADODB.Recordset
Public Function SimuladoReprogramacion(ByVal pcCtaCod As String, Optional ByVal pnMenuOp As Integer = -1) As ADODB.Recordset 'Add JOEP20210306 garantia covid
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    'sSql = "Exec stp_sel_SimuladorReprog '" & pcCtaCod & "'"
    sSql = "Exec stp_sel_SimuladorReprog '" & pcCtaCod & "'," & pnMenuOp & "" 'Add JOEP20210306 garantia covid
    oCon.AbreConexion
    Set SimuladoReprogramacion = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

'Public Function SimuladoReprogramacionCargarDatos(ByVal psCtaCod As String, ByRef prsCred As ADODB.Recordset, ByRef prsCal As ADODB.Recordset)
Public Function SimuladoReprogramacionCargarDatos(ByVal psCtaCod As String, ByRef prsCred As ADODB.Recordset, ByRef prsCal As ADODB.Recordset, Optional ByVal pnMenuOpcion As Integer = -1) 'Add JOEP20210306 garantia covid

Dim oCalend As COMDCredito.DCOMCalendario
Dim rsAut As ADODB.Recordset, rsRep As ADODB.Recordset

Dim oParam As COMDCredito.DCOMParametro

    On Error GoTo ErrorSimuladoReprogramacionCargarDatos
    
    'Set prsCred = SimuladoReprogramacionDatosComun(psCtaCod)
    Set prsCred = SimuladoReprogramacionDatosComun(psCtaCod, pnMenuOpcion) 'Add JOEP20210306 garantia covid
    
    Set oCalend = New COMDCredito.DCOMCalendario
    Set prsCal = oCalend.RecuperaCalendarioPagos(psCtaCod)
    Set oCalend = Nothing
    
    Exit Function

ErrorSimuladoReprogramacionCargarDatos:
    Err.Raise Err.Number, "Cargar Datos Reprogramacion", Err.Description
End Function

'Public Function SimuladoReprogramacionDatosComun(ByVal pcCtaCod As String) As ADODB.Recordset
Public Function SimuladoReprogramacionDatosComun(ByVal pcCtaCod As String, Optional ByVal pnMenuOpcion As Integer = -1) As ADODB.Recordset 'Add JOEP20210306 garantia covid
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    'sSql = "Exec stp_sel_SimuladorReprogramacionDatos '" & pcCtaCod & "'"
    sSql = "Exec stp_sel_SimuladorReprogramacionDatos '" & pcCtaCod & "'," & pnMenuOpcion & "" 'Add JOEP20210306 garantia covid
    oCon.AbreConexion
    Set SimuladoReprogramacionDatosComun = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ReprogPropuestaValOcmSicm(ByVal pcCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    sSql = "Exec stp_sel_ReprogramacionPropuestaValidaOcmSicmam '" & pcCtaCod & "'"
    oCon.AbreConexion
    Set ReprogPropuestaValOcmSicm = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'JOEP202008 Tasa ESpecial y reduccion de monto



'reversion JOEP20201124
Public Function ReprogRechazoCombo(ByVal pcCtaCod As String, ByVal pnOp As Integer) As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    sSql = "Exec stp_sel_ReprogramacionRechazo_ObtieneDias '" & pcCtaCod & "'," & pnOp & ""
    oCon.AbreConexion
    Set ReprogRechazoCombo = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ReprogRechazoLlenaCombo(ByVal pcCtaCod As String, ByVal pnConsCod As Integer) As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    sSql = "Exec stp_sel_ReprogramacionRechazo_LlenaCombo '" & pcCtaCod & "'," & pnConsCod & ""
    oCon.AbreConexion
    Set ReprogRechazoLlenaCombo = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'reversion JOEP20201124
'CTI5 ERS0032020 20200120 ***
Public Sub GuardarHistoricoConfiguracionSobreEnd(ByVal psUsuario As String, ByVal psGlosa As String, ByVal pnTipoConfiguracion As Integer)
    Dim lsSQL As String
    lsSQL = "EXEC [dbo].[stp_ins_HistoricoConfiguracionSobreEnd] @pcUsuario='" & psUsuario & "'" & _
         ",@pcGlosa='" & psGlosa & "',@pnTipoConfiguracion=" & pnTipoConfiguracion
    loConecta.Ejecutar (lsSQL)
End Sub
Public Sub EliminaTramoCPSobreEnd(ByVal pnTipoConfiguracion As Integer)
    Dim lsSQL As String
    lsSQL = "EXEC [dbo].[stp_del_TramosCapacidadPagoSobreEnd] " & pnTipoConfiguracion
    loConecta.Ejecutar (lsSQL)
End Sub
Public Sub EliminaNroEntidadesSobreEnd(ByVal pnTipoConfiguracion As Integer)
    Dim lsSQL As String
    lsSQL = "EXEC [dbo].[stp_del_TramosNroEntidadesSobreEnd] " & pnTipoConfiguracion
    loConecta.Ejecutar (lsSQL)
End Sub
Public Sub EliminaMatrizConfiguracionSobreEnd(ByVal pnTipoConfiguracion As Integer)
    Dim lsSQL As String
    lsSQL = "EXEC [dbo].[stp_del_MatrizConfiguracionSobreEnd] " & pnTipoConfiguracion
    loConecta.Ejecutar (lsSQL)
End Sub
Public Sub InsertaTramoCPSobreEnd(ByVal pnTipoConfiguracion As Integer, _
    ByVal pnOrdenTramoCP As Integer, _
    ByVal pnTramoCPInicial As Double, _
    ByVal pnTramoCPFinal As Double, _
    ByVal psDescripcionTramoCP As String)
    
    Dim lsSQL As String
    lsSQL = "EXEC [dbo].[stp_ins_TramosCapacidadPagoSobreEnd] @pnTipoConfiguracion=" & pnTipoConfiguracion _
        & ",@pnOrdenTramoCP=" & pnOrdenTramoCP _
        & ",@pnTramoCPInicial=" & pnTramoCPInicial _
        & ",@pnTramoCPFinal=" & pnTramoCPFinal _
        & ",@pcDescripcionTramoCP='" & psDescripcionTramoCP & "'"
    loConecta.Ejecutar (lsSQL)
End Sub
Public Sub InsertaNroEntidadesSobreEnd(ByVal pnTipoConfiguracion As Integer, _
    ByVal pnOrdenNroEntidad As Integer, _
    ByVal pnNroEntidadIni As Integer, _
    ByVal pnNroEntidadFin As Integer, _
    ByVal psDescripcionNroEntidad As String)
    
    Dim lsSQL As String
    lsSQL = "EXEC [dbo].[stp_ins_TramosNroEntidadesSobreEnd] @pnTipoConfiguracion=" & pnTipoConfiguracion _
        & ",@pnOrdenNroEntidad=" & pnOrdenNroEntidad _
        & ",@pnNroEntidadIni=" & pnNroEntidadIni _
        & ",@pnNroEntidadFin=" & pnNroEntidadFin _
        & ",@pcDescripcionNroEntidad='" & psDescripcionNroEntidad & "'"
    loConecta.Ejecutar (lsSQL)
End Sub
Public Sub InsertaMatrizConfiguracionSobreEnd(ByVal pnTipoConfiguracion As Integer, _
    ByVal pnOrdenTramoCP As Integer, _
    ByVal pnOrdenNroEntidad As Integer, _
    ByVal pnRSE As Integer)
    
    Dim lsSQL As String
    lsSQL = "EXEC [dbo].[stp_ins_MatrizConfiguracionSobreEnd] @pnTipoConfiguracion=" & pnTipoConfiguracion _
        & ",@pnOrdenTramoCP=" & pnOrdenTramoCP _
        & ",@pnOrdenNroEntidad=" & pnOrdenNroEntidad _
        & ",@pnRSE=" & pnRSE
    loConecta.Ejecutar (lsSQL)
End Sub

Public Function ObtenerCalendarioMIVIENDA(ByVal pcCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    sSql = "Exec stp_sel_CalendarioCredito '" & pcCtaCod & "'"
    oCon.AbreConexion
    Set ObtenerCalendarioMIVIENDA = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function ObtenerDatosCredito(ByVal psNumGarant As String) As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    sSql = "Exec stp_sel_ObtenerDatosCreditoxGarantia '" & psNumGarant & "'"
    oCon.AbreConexion
    Set ObtenerDatosCredito = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Sub ActualizaTCEA(ByVal psCtaCod As String, pnTCEA As Double)
Dim lsSQL As String
lsSQL = "EXEC sp_sel_ActualizarTCEA '" & psCtaCod & "'," & pnTCEA & ""
loConecta.Ejecutar (lsSQL)
End Sub

Public Function RecuperaCtaAhorroGarantiasCreditoTP(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_RecuperaCtaAhorroGarantiasCreditoTP '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperaCtaAhorroGarantiasCreditoTP = rs
  
    Set rs = Nothing
End Function
'END CTI5 *******

'CTI3 ERS0032020******************************
Public Function ObtenerDatosCreditoxVB(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    sSql = "Exec stp_sel_ObtenerDatosCreditoxVB '" & psCtaCod & "'"
    oCon.AbreConexion
    Set ObtenerDatosCreditoxVB = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ActualizarEstadoxVB(ByVal psCtaCod As String, ByVal pnEstadoVB As Integer) As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    sSql = "Exec stp_upd_ActualizarEstadoxVB '" & psCtaCod & "'," & pnEstadoVB
    oCon.AbreConexion
    Set ActualizarEstadoxVB = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'********************************************
'LARI20210203: VERIFICACIÓN DE AMPLIACIÓN -************************************
Public Function ValidaCreditoAmpliarPagos(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_CalifAmpliCredito '" & psCtaCod & "'"
    oCon.AbreConexion
    Set ValidaCreditoAmpliarPagos = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'****************************************************************

'LARI20210604: eliminacion de los restos de las cuotas reprogramadas -************************************
Public Function EliminarCuotaSobranteRep(psCtaCod As String, nCuotaF As Integer) As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC sp_Del_Calendario '" & psCtaCod & "'," & nCuotaF
    oCon.AbreConexion
    Set EliminarCuotaSobranteRep = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'****************************************************************
      

'JOEP20210211 Garantia Covid
Public Function ReprogramacionPropuestaMsgBox(ByVal pcCtaCod As String, ByVal pnMenuOpcion As Integer) As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    sSql = "Exec stp_sel_ReprogramacionMensaje '" & pcCtaCod & "'," & pnMenuOpcion & ""
    oCon.AbreConexion
    Set ReprogramacionPropuestaMsgBox = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ReprogramacionPropuestaDatosClienteSimulCalend(ByVal psCtaCod As String, ByVal pdFecha As Date) As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    sSql = "Exec stp_sel_ColocacReprogramadoDatosClienteSimulacionCronograma '" & psCtaCod & "','" & Format(pdFecha, "yyyyMMdd") & "'"
    oCon.AbreConexion
    Set ReprogramacionPropuestaDatosClienteSimulCalend = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ReprogramacionPropuestaDatosSimulCalend(ByVal psCtaCod As String, ByVal pdFecha As Date) As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    sSql = "Exec stp_sel_ColocacReprogramadoDatosSimulacionCronograma '" & psCtaCod & "','" & Format(pdFecha, "yyyyMMdd") & "'"
    oCon.AbreConexion
    Set ReprogramacionPropuestaDatosSimulCalend = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ReprogramacionPropuestaDatosTotalSimulCalend(ByVal psCtaCod As String, ByVal pdFecha As Date) As ADODB.Recordset
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

    sSql = "Exec stp_sel_ColocacReprogramadoDatostTotalesSimulacionCronograma '" & psCtaCod & "','" & Format(pdFecha, "yyyyMMdd") & "'"
    oCon.AbreConexion
    Set ReprogramacionPropuestaDatosTotalSimulCalend = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'JOEP20210211 Garantia Covid
