VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMColRecNegociacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim coConex As COMConecta.DCOMConecta
Dim oError As New COMConecta.COMErrorHandling
Private Sub Class_Initialize()
    Dim loIni As COMConecta.DCOMClasIni
    Dim csConexion As String
    Dim csNegocio As String
    Dim csCentralPer As String
    Dim csCentralCom As String
    Dim csCentralImg As String
    Dim csAdminist As String
    
    Set loIni = New COMConecta.DCOMClasIni
        csConexion = loIni.CadenaConexion
        csNegocio = loIni.BaseNegocio
        csCentralPer = loIni.BasePersonas
        csCentralCom = loIni.BaseComunes
        csCentralImg = loIni.BaseImagenes
        csAdminist = loIni.BaseAdministracion
    Set loIni = Nothing

Set coConex = New COMConecta.DCOMConecta
If coConex.AbreConexion(csConexion) = False Then
    Call oError.RaiseError(oError.MyUnhandledError, "DColPContrato:Initialize. Error en Conexion a Base de datos")
End If
End Sub

Private Sub Class_Terminate()
    coConex.CierraConexion
    Set coConex = Nothing
End Sub


Public Function ObtenerDatosCredparaNegociacion(ByVal psctacod As String) As ADODB.Recordset
   
    Dim lssql As String
    Dim lrs As New ADODB.Recordset
On Error GoTo dError

     lssql = "SELECT P.cCtaCod, P.nPrdEstado, P.nSaldo, " _
        & " CRec.nSaldoIntComp, CRec.nSaldoIntMor, CRec.nSaldoGasto, CRec.dIngRecup, " _
        & " nTipComis = ISNULL ( ( SELECT nTipComis FROM ColocRecupComision WHERE nComisionCod = CRec.nComisionCod ), '') , " _
        & " nComisionValor = ISNULL ( ( SELECT nValor FROM ColocRecupComision WHERE nComisionCod = CRec.nComisionCod ), 0 ) , " _
        & " nTasaInt = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " _
        & "             WHERE LCT.cLineaCred = Coloc.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntCompNormal & " ),  " _
        & " NomCliente = (Select P.cPersNombre From Persona P  " _
        & "               Inner Join ProductoPersona PC ON P.cPersCod = PC.cPersCod " _
        & "               Where PC.cCtaCod = '" & psctacod & "' And PC.nPrdPersRelac = " & gColRelPersTitular & " ) , " _
        & " NomEstJur = ISNULL ( ( SELECT cPersNombre FROM Persona P INNER JOIN ProductoPersona PP " _
        & "              ON P.cPersCod = PP.cPersCod WHERE PP.cCtaCod = '" & psctacod & "' " _
        & "              AND nPrdPersRelac = " & gColRelPersEstudioJuridico & " ), '' ),  " _
        & " nTasaInt = (SELECT ISNULL(nTasaInteres, 0) From ProductoTasaInteres LCT  WHERE LCT.cCtaCod = Coloc.cCtaCod and LCT.nPrdTasaInteres = " & gColocLineaCredTasasIntCompNormal & " ) , " _
        & " nTasaIntMor = (SELECT ISNULL(nTasaInteres, 0) From ProductoTasaInteres LCT  WHERE LCT.cCtaCod = Coloc.cCtaCod and LCT.nPrdTasaInteres = " & gColocLineaCredTasasIntMoratNormal & " ), " _
        & " C.cUltimaActualizacion" _
        & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
        & " Inner Join ColocRecup CRec on C.cCtaCod = CRec.cCtaCod " _
        & " Inner Join Colocaciones Coloc ON CRec.cCtaCod = Coloc.cCtaCod " _
        & " WHERE P.cCtaCod ='" & psctacod & "' "
   Set lrs = coConex.CargaRecordSet(lssql)
        Set ObtenerDatosCredparaNegociacion = lrs
   Set lrs = Nothing
   Exit Function
dError:
    Err.Raise Err.Number, "Error al Cargar Datos", Err.Description
End Function
   
  
Public Function ObtenerDatosCredNegociacion(ByVal psctacod As String) As ADODB.Recordset
   
    Dim lssql As String
    Dim lrs As New ADODB.Recordset
On Error GoTo dError

   lssql = "SELECT * FROM ColocRecupNegocia WHERE cCtaCod = '" & psctacod & "' And cEstado = 'V' "
   Set lrs = coConex.CargaRecordSet(lssql)
        Set ObtenerDatosCredNegociacion = lrs
   Set lrs = Nothing
   Exit Function
dError:
    Err.Raise Err.Number, "Error al Cargar Datos", Err.Description
End Function
   
Public Sub AnularNegociacion(ByVal psNroNegocia As String, ByVal lsFechaHora As Date, ByVal gscoduser As String)
    Dim lssql As String
    On Error GoTo dError
    lssql = "UPDATE ColocRecupNegocia Set cEstado ='L' , dFecMod = '" & Format(lsFechaHora, "yyyymmdd") & "', " _
                   & " cCodUsu = '" & gscoduser & "', cComenta = cComenta + ' - Anulada ' " _
                   & " Where cNroNeg = '" & psNroNegocia & "' "
    coConex.Ejecutar (lssql)
    Exit Sub
dError:
    Err.Raise Err.Number, "Error al Anular Datos", Err.Description
End Sub

Public Sub InsertarNegociacion(ByVal pdfecVig As Date, ByVal psFecAnt As String, ByVal psCodCta As String, psNroNegocia, _
                               ByVal pnNegMonto As Currency, ByVal pnNegCuotas As Currency, ByVal psNegComenta As String, _
                               ByVal gscoduser As String, ByVal pdfechaH As Date, ByVal nNumTranCta As Integer, ByVal psNegoOperac As String, _
                               ByVal gsCodAge As String, ByVal rs As ADODB.Recordset)
    Dim lssql As String
    Dim i As Integer
    On Error GoTo dError
     'Anula negociacion Anterior
      coConex.BeginTrans
        If Len(Trim(psFecAnt)) > 0 Then
            lssql = "UPDATE ColocRecupNegocia Set cEstado ='L' , dFecMod = '" & Format(pdfecVig, "mm/dd/yyyy") & "', " _
               & " cComenta = cComenta +  ' - Anulada por Nueva Negociacion ' + '" & pdfecVig & "' " _
               & " Where cNroNeg = '" & psFecAnt & "' "
         coConex.Ejecutar (lssql)
        End If
        ' Inserta ColocRecupNegocia
        lssql = "INSERT INTO ColocRecupNegocia ( cCtaCod, cNroNeg,cEstado, dFecVig, nMontoNeg, " _
            & " nCuotasNeg, cComenta ,dFecCancel , cCodUsu, dFecMod, nNumTranCta ) " _
            & " VALUES ('" & psCodCta & "','" & psNroNegocia & "','V','" & Format(pdfecVig, "mm/dd/yyyy") & "'," _
            & CDbl(pnNegMonto) & "," & Val(pnNegCuotas) & ",'" _
            & Trim(psNegComenta) & "',null, '" & gscoduser & "','" & Format(pdfechaH, "mm/dd/yyyy") & "'," & nNumTranCta & " ) "
         coConex.Ejecutar (lssql)
        ' Inserta NegocKardex
        lssql = " INSERT INTO ColocRecupNegociaKardex (cCtaCod, cNroNeg, cOpeCod, dFecTran, nNumTranCta, " _
            & " nMonTran, nDiasAtraso, cNumDoc, cEstado, cCodAge, cCodUsu, cFlag) " _
            & " VALUES ('" & psCodCta & "','" & psNroNegocia & "','" & psNegoOperac & "','" & Format(pdfechaH, "mm/dd/yyyy") & "'," & nNumTranCta & "," _
            & CCur(pnNegMonto) & ",null, null, null, '" & gsCodAge & "','" & gscoduser & "', null ) "
        coConex.Ejecutar (lssql)
         ' Inserta Plan de Pagos
   
        i = 1
        If Not (rs.EOF And rs.BOF) Then
            rs.MoveFirst
            Do Until rs.EOF
            ' Inserta Cuota
                lssql = " INSERT INTO ColocRecupNegocPlanPagos (cCtaCod,cNroNeg, nNroCuota, dFecVenc, dFecPag, nCapital, nMonto, " _
                    & " nCapPag, nMontoPag, cEstado) " _
                    & " VALUES ('" & psCodCta & "','" & psNroNegocia & "'," & i & ",'" & Format(rs!dfecha, "mm/dd/yyyy") & "', null, 0, " _
                    & CCur(rs!nMonto) & ", 0 , 0, 'P'  ) "
                coConex.Ejecutar (lssql)
                i = i + 1
                rs.MoveNext
            Loop
        End If
        'Actualiza Nro de Negociacion
        lssql = "UPDATE ColocRecup set cNroNeg = '" & psNroNegocia & "' Where cCtaCod ='" & psCodCta & "' "
        coConex.Ejecutar (lssql)
        coConex.CommitTrans
        Exit Sub
dError:
    Err.Raise Err.Number, "Error al Anular Datos", Err.Description
End Sub

Public Function ObtenerDatosCredito(ByVal psctacod As String) As ADODB.Recordset
    Dim lssql As String
    Dim lrs As New ADODB.Recordset
  On Error GoTo dError
    lssql = "SELECT P.cCtaCod, P.nPrdEstado, P.nSaldo,  " _
    & " cNomClie = ISNULL ( ( SELECT cPersNombre FROM Persona P INNER JOIN ProductoPersona PP " _
    & "              ON P.cPersCod = PP.cPersCod WHERE PP.cCtaCod = '" & psctacod & "' " _
    & "              AND nPrdPersRelac = " & gColRelPersTitular & " ), '' ) , " _
    & " cCodAbog = ISNULL ( ( SELECT cPersCod FROM ProductoPersona WHERE cCtaCod = '" & psctacod & "' " _
    & "              AND nPrdPersRelac = " & gColRelPersEstudioJuridico & " ), '' ) , " _
    & " cNomAbog = ISNULL ( ( SELECT cPersNombre FROM Persona P INNER JOIN ProductoPersona PP " _
    & "              ON P.cPersCod = PP.cPersCod WHERE PP.cCtaCod = '" & psctacod & "' " _
    & "              AND nPrdPersRelac = " & gColRelPersEstudioJuridico & " ), '' )  " _
    & " FROM Producto P Inner Join ColocRecup CR on P.cCtaCod = CR.cCtaCod  " _
    & " WHERE P.cCtaCod = '" & psctacod & "'"
     Set lrs = coConex.CargaRecordSet(lssql)
        Set ObtenerDatosCredito = lrs
     Set lrs = Nothing
     Exit Function
dError:
    Err.Raise Err.Number, "Error al cargar los Datos", Err.Description
End Function

Public Function ObtenerPlanPagosNegocia(ByVal psctacod As String, ByVal sNegNro As String) As ADODB.Recordset
    Dim lssql As String
    Dim lrs As New ADODB.Recordset
 On Error GoTo dError
    lssql = "SELECT * FROM ColocRecupNegocPlanPagos WHERE cCtaCod = '" & psctacod & "' " & _
                "And cNroNeg ='" & Trim(sNegNro) & "'ORDER BY nNroCuota"
    Set lrs = coConex.CargaRecordSet(lssql)
        Set ObtenerPlanPagosNegocia = lrs
    Set lrs = Nothing
    Exit Function
dError:
    Err.Raise Err.Number, "Error al cargar los Datos", Err.Description
End Function
Public Function ObtenerNroNeg(ByVal psctacod As String) As ADODB.Recordset
    Dim lssql As String
    Dim lrs As New ADODB.Recordset
  On Error GoTo dError
    lssql = " Select cNroNeg from ColocRecupNegocia Where cCtaCod ='" & psctacod & "' And cEstado ='V' "
    Set lrs = coConex.CargaRecordSet(lssql)
        Set ObtenerNroNeg = lrs
    Set lrs = Nothing
    Exit Function
dError:
    Err.Raise Err.Number, "Error al cargar los Datos", Err.Description
End Function

Public Function ObtenerNroNegMax() As ADODB.Recordset
    Dim lssql As String
    Dim lrs As New ADODB.Recordset
On Error GoTo dError
    lssql = " Select Max(cNroNeg) UltNeg From ColocRecupNegocia "
    Set lrs = coConex.CargaRecordSet(lssql)
        Set ObtenerNroNegMax = lrs
    Set lrs = Nothing
    Exit Function
dError:
    Err.Raise Err.Number, "Error al cargar los Datos", Err.Description
End Function

Public Sub ActualizarNegocPlanPagos(ByVal pdFechaHoraGrab As Date, ByVal psctacod As String, ByVal psNroNeg As String, ByVal rs As ADODB.Recordset)
    Dim lssql As String
    On Error GoTo dError
     
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do Until rs.EOF
            If rs(5) = True Then
                lssql = " UPDATE ColocRecupNegocPlanPagos SET " & _
                                " dFecPag = '" & Format(pdFechaHoraGrab, "yyyymmdd hh:mm:ss") & "' " & _
                                ",cEstado = '" & rs(4) & "'" & _
                                ",nMontoPag = nMontoPag + " & Format(rs(3), "#0.00") & _
                                "  WHERE cCtaCod = '" & psctacod & "' And cNroNeg ='" & psNroNeg & "' " & _
                                " And nNroCuota = " & Str(rs(0)) & " "
            End If
            rs.MoveNext
        Loop
    End If
    coConex.Ejecutar (lssql)
    Exit Sub
dError:
    Err.Raise Err.Number, "Error al Anular Datos", Err.Description
End Sub


Public Sub InsertaNegocPlanPagoskardex(ByVal psctacod As String, ByVal psNroNeg As String, ByVal psOpeCod As String, ByVal pdFechaHoraGrab As Date, ByVal pnMontoPag As Double, _
                                       ByVal fnDiasAtraso As String, ByVal psNumDoc As String, ByVal gsCodAge As String, ByVal gscoduser As String)
    Dim lssql As String
    On Error GoTo dError
     
   
    lssql = "Insert into ColocRecupNegociaKardex (cCtaCod, cNroNeg, cOpeCod, dFecTran, nNumTranCta, nMonTran, nDiasAtraso, " & _
           " cNumDoc, cEstado, cCodAge, cCodUsu, cFlag ) " & _
           " Values ( '" & psctacod & "','" & psNroNeg & "','" & psOpeCod & "','" & Format(pdFechaHoraGrab, "yyyymmdd hh:mm:ss") & _
           "',null," & pnMontoPag & "," & _
            fnDiasAtraso & ",'" & psNumDoc & "',null, '" & gsCodAge & "','" & gscoduser & "', null ) "
   
    coConex.Ejecutar (lssql)
    Exit Sub
dError:
    Err.Raise Err.Number, "Error al Anular Datos", Err.Description
End Sub



