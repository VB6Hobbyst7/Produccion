VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DColRecCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim csConexion As String
Dim csNegocio As String
Dim csCentralPer As String
Dim csCentralCom As String
Dim csCentralImg As String
Dim csAdminist As String

Dim coConex As DConecta
Private Sub Class_Initialize()
    Dim loIni As COMConecta.DCOMClasIni
    
    Set loIni = New COMConecta.DCOMClasIni
        csConexion = loIni.CadenaConexion
        csNegocio = loIni.BaseNegocio
        csCentralPer = loIni.BasePersonas
        csCentralCom = loIni.BaseComunes
        csCentralImg = loIni.BaseImagenes
        csAdminist = loIni.BaseAdministracion
    Set loIni = Nothing

Set coConex = New DConecta
If coConex.AbreConexion(csConexion) = False Then
    Call RaiseError(MyUnhandledError, "DColPContrato:Initialize. Error en Conexion a Base de datos")
End If
End Sub

Private Sub Class_Terminate()
    coConex.CierraConexion
    Set coConex = Nothing
End Sub


Public Function dObtieneCreditosDePersona(ByVal psPersCod As String, _
        ByVal psEstados As String, Optional ByVal psAgencia As String = "") As Recordset
'Obtiene Listado Creditos en Recuperaciones  de una Persona
Dim lrs As ADODB.Recordset
Dim lsSQL As String
Dim lsAgencia As String

If Trim(psAgencia) = "" Then
    lsAgencia = "__"
Else
    lsAgencia = Trim(psAgencia)
End If

On Error GoTo dError

lsSQL = "SELECT PP.cCtaCod, P.nPrdEstado, PP.nPrdPersRelac, T.cConsDescripcion cEstado, " _
        & "T1.cConsDescripcion cRelacion, UPPER(T2.cConsDescripcion) cProducto, " _
        & "UPPER(T3.cConsDescripcion) cMoneda " _
        & "FROM ProductoPersona PP INNER JOIN Producto P  " _
        & "ON P.cCtaCod = PP.cCtaCod INNER JOIN " _
        & csCentralCom & "Constante T ON P.nPrdEstado = T.nConsValor INNER JOIN " & csCentralCom & "" _
        & "Constante T1 ON PP.nPrdPersRelac = T1.nConsValor INNER JOIN " & csCentralCom & "Constante T2 " _
        & "ON SUBSTRING(PP.cCtaCod,6,3) = CONVERT(Varchar(3),T2.nConsValor) INNER JOIN " & csCentralCom & "" _
        & "Constante T3 ON SUBSTRING(PP.cCtaCod,9,1) = CONVERT(Varchar(1),T3.nConsValor) " _
        & "WHERE PP.cPersCod = '" & psPersCod & "' " _
        & "AND T1.nConsCod = " & gColocRelacPers & " AND T.nConsCod = " & gColocEstado & " AND " _
        & "T2.nConsCod = " & gProducto & " AND T3.nConsCod = " & gMoneda _
        & "AND P.cCtaCod like '___" & lsAgencia & "%' " _
        & "AND P.nPrdEstado in ( " & psEstados & " ,2205 , 2206  ) "

lsSQL = lsSQL & " ORDER BY PP.cCtaCod"

Set lrs = coConex.CargaRecordSet(lsSQL)

Set dObtieneCreditosDePersona = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
    
End Function

Public Function dObtieneListaGastosxCredito(ByVal psctacod As String) As Recordset

'************************************
Dim lsSQL As String
Dim loRegValida As dColPFunciones
Dim lrValida As ADODB.Recordset

    lsSQL = " SELECT G.dAsigna, G.nMonto, G.nMontoPagado, G.cMotivoGasto,  " _
          & " G.nPrdConceptoCod, G.nNroGastoCta, G.nColocRecGastoEstado  " _
          & " FROM ColocRecup CRec INNER JOIN ColocRecupGastos G " _
          & " ON CRec.cCtaCod = G.cCtaCod " _
          & " " _
          & " WHERE CRec.cCtaCod ='" & psctacod & "' " _
          & " AND G.nColocRecGastoEstado not in ( " & gColRecGastoEstEliminado & " ) " _
          & " ORDER BY G.nNroGastoCta "
          
    Set loRegValida = New dColPFunciones
        Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
        Exit Function
    End If
    
    Set dObtieneListaGastosxCredito = lrValida
    Set lrValida = Nothing
End Function


Public Function dObtieneDatosAsignaGastoCredRecup(ByVal psctacod As String) As Recordset
'Obtiene Datos de Credito en Recuperaciones
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New Recordset

On Error GoTo dError

lsSQL = " SELECT P.cCtaCod, P.nPrdEstado, C.nMontoCol, P.nSaldo,  " _
    & " CRec.dIngRecup, CRec.nSaldoIntComp, CRec.nSaldoIntMor, CRec.nSaldoGasto, " _
    & " Pers.cPersCod, Pers.cPersNombre, " _
    & " nUltGasto = (SELECT ISNULL(MAX(nNroGastoCta),0) FROM ColocRecupGastos " _
    & "              WHERE cCtaCod ='" & psctacod & "' ) " _
    & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
    & " Inner Join ColocRecup CRec ON C.cctacod = CRec.cctacod " _
    & " Inner Join ProductoPersona PP ON C.cCtaCod = PP.cCtaCod " _
    & " Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod " _
    & " WHERE P.cCtaCod ='" & psctacod & "' " _
    & " AND PP.nPrdPersRelac = " & gColRelPersTitular & " "

Set lrs = coConex.CargaRecordSet(lsSQL)

Set dObtieneDatosAsignaGastoCredRecup = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneDatosContrato>>", Err.Description
    
End Function

Public Function dObtieneGastosRecup() As Recordset
'Obtiene Datos de Credito en Recuperaciones
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New Recordset

On Error GoTo dError

lsSQL = " SELECT nPrdConceptoCod, cDescripcion " _
    & " FROM ProductoConcepto " _
    & " WHERE nPrdConceptoCod like '32%' AND cAplicaProceso = 'MA' " _
    & " ORDER BY nPrdConceptoCod "
    
Set lrs = coConex.CargaRecordSet(lsSQL)

Set dObtieneGastosRecup = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneDatosContrato>>", Err.Description
    
End Function


Public Function dObtieneDatosPagoCredRecup(ByVal psctacod As String) As Recordset
'Obtiene Datos para Pago de Credito en Recuperaciones
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New Recordset

On Error GoTo dError

lsSQL = " SELECT P.cCtaCod, PP.cPersCod, P.nPrdEstado, C.nMontoCol, P.nSaldo, C.cUltimaActualizacion,  " _
    & " CRec.nSaldoIntComp, CRec.nSaldoIntMor, CRec.nSaldoGasto, CRec.nIntCompGen, CRec.cNroNeg, " _
    & " CRec.nComisionCod, CRec.cMetLiquid, CRec.nTipCj, CRec.nDemanda, CRec.nNroCalen, IsNull(Com.nValor,0) nValorCom,  " _
    & " Pers.cPersCod, Pers.cPersNombre, Pers2.cPersCod cPersCodAbog, Pers2.cPersNombre cPersNombreAbog,  " _
    & " nUltGas = (Select Isnull(max(nNroGastoCta),0) From ColocRecupGastos CRG Where CRG.cCtaCod = p.cCtaCod ) , " _
    & " nTasaInt = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT  " _
    & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntCompNormal & " ) , " _
    & " nTasaIntMor = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT  " _
    & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " ) , " _
    & " cEstudioAbogCod = (SELECT cPersCod FROM ProductoPersona  " _
    & "           WHERE cCtaCod ='" & psctacod & "' AND nPrdPersRelac = " & gColRelPersEstudioJuridico & " ) " _
    & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
    & " Inner Join ColocRecup CRec ON C.cctacod = CRec.cctacod " _
    & " Inner Join ProductoPersona PP ON C.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & gColRelPersTitular _
    & " Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod " _
    & " Inner Join ProductoPersona PP2 ON C.cCtaCod = PP2.cCtaCod And PP2.nPrdPersRelac = " & gColRelPersEstudioJuridico _
    & " Inner Join Persona Pers2 ON Pers2.cPersCod = PP2.cPersCod " _
    & " Left Join ColocRecupComision Com ON Com.nComisionCod = CRec.nComisionCod " _
    & " WHERE P.cCtaCod ='" & psctacod & "'"

Set lrs = coConex.CargaRecordSet(lsSQL)
    
    If lrs.BOF And lrs.EOF Then
        MsgBox "No se ha realizado registro de expediente ", vbInformation, "Aviso"
        Exit Function
    End If

    If lrs!nPrdEstado = gColocEstRecCanJud Or lrs!nPrdEstado = gColocEstRecCanCast Then
        MsgBox "Credito se encuentra en estado " & mfgEstadoColocRecupDesc(lrs!nPrdEstado), vbInformation, "Aviso"
        Exit Function
    End If
    
    
Set dObtieneDatosPagoCredRecup = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneDatosPagoCredRecup>>", Err.Description
    
End Function


Public Function dObtieneDatosCancelaCredRecup(ByVal psctacod As String) As Recordset
'Obtiene Datos para Cancelacion de Credito en Recuperaciones
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New Recordset

On Error GoTo dError

lsSQL = " SELECT P.cCtaCod, P.nPrdEstado, P.nTasaInteres, C.nMontoCol, P.nSaldo, C.cUltimaActualizacion,  " _
    & " CRec.nSaldoIntComp, CRec.nSaldoIntMor, CRec.nSaldoGasto, CRec.nIntCompGen, CRec.dIngRecup,  " _
    & " CRec.nComisionCod, CRec.cMetLiquid, CRec.nTipCj, CRec.nDemanda, CRec.nNroCalen, IsNull(Com.nValor,0) nValorCom,  " _
    & " Pers.cPersCod, Pers.cPersNombre, Pers2.cPersCod cPersCodAbog, Pers2.cPersNombre cPersNombreAbog,  " _
    & " cEstudioAbogCod = (SELECT cPersCod FROM ProductoPersona  " _
    & "           WHERE cCtaCod ='" & psctacod & "' AND nPrdPersRelac = " & gColRelPersEstudioJuridico & " ) , " _
    & " nTasaInt = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " _
    & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntCompNormal & " )  " _
    & "  " _
    & "  " _
    & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
    & " Inner Join ColocRecup CRec ON C.cctacod = CRec.cctacod " _
    & " Inner Join ProductoPersona PP ON C.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & gColRelPersTitular _
    & " Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod " _
    & " Inner Join ProductoPersona PP2 ON C.cCtaCod = PP2.cCtaCod And PP2.nPrdPersRelac = " & gColRelPersEstudioJuridico _
    & " Inner Join Persona Pers2 ON Pers2.cPersCod = PP2.cPersCod " _
    & " Left Join ColocRecupComision Com ON Com.nComisionCod = CRec.nComisionCod " _
    & " WHERE P.cCtaCod ='" & psctacod & "'"

Set lrs = coConex.CargaRecordSet(lsSQL)

    If Not (lrs.EOF And lrs.BOF) Then
        If lrs!nPrdEstado = gColocEstRecCanJud Or lrs!nPrdEstado = gColocEstRecCanCast Then
            MsgBox "Credito ya se encuentra en estado " & mfgEstadoColocRecupDesc(lrs!nPrdEstado), vbInformation, "Aviso"
            Exit Function
        End If
    End If
    
Set dObtieneDatosCancelaCredRecup = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneDatosPagoCredRecup>>", Err.Description
    
End Function

Public Function dObtieneDatosCierreMes() As ADODB.Recordset
Dim lsSQL As String
Dim lrs As Recordset
Set lrs = New Recordset

On Error GoTo ErrorRecupera
    
lsSQL = " SELECT P.cCtaCod, P.nPrdEstado, P.nTasaInteres, C.nMontoCol, P.nSaldo, C.cUltimaActualizacion,  " _
    & " CRec.nSaldoIntComp, CRec.nSaldoIntMor, CRec.nSaldoGasto, CRec.nIntCompGen, CRec.dIngRecup,  " _
    & " CRec.nTipCj, CRec.nDemanda, CRec.nNroCalen, " _
    & " nTasaIntComp = (SELECT ISNULL(nTasaInteres, 0) From ProductoTasaInteres  PT    WHERE PT.cCtaCod = C.cCtacod  and  nPrdTasaInteres=1 )," _
    & " nTasaIntMorat = (SELECT ISNULL(nTasaInteres, 0) From ProductoTasaInteres  PT    WHERE PT.cCtaCod = C.cCtacod  and  nPrdTasaInteres=3)" _
    & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
    & " Inner Join ColocRecup CRec ON C.cctacod = CRec.cctacod " _
    & " WHERE P.nPrdEstado in (" & gColocEstRecVigJud & "," & gColocEstRecVigCast & " ) "
    
Set lrs = coConex.CargaRecordSet(lsSQL)

    If lrs.EOF And lrs.BOF Then
            MsgBox "No se encuentran Creditos en Recuperaciones ", vbInformation, "Aviso"
            Exit Function
    End If
    
Set dObtieneDatosCierreMes = lrs
Set lrs = Nothing
    
Exit Function

ErrorRecupera:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
        
    
End Function

Public Function dObtieneDatosNegociaCuotas(ByVal psCodCta As String, ByVal psNroNeg As String) As ADODB.Recordset
Dim lsSQL As String
Dim lrs As Recordset
Set lrs = New Recordset

On Error GoTo ErrorRecupera
    
lsSQL = "SELECT * FROM ColocRecupNegocPlanPagos WHERE cCtaCod ='" & psCodCta & "' " & _
      " And cNroNeg = '" & psNroNeg & "' " & _
      " And cEstado ='P' ORDER BY nNroCuota "
    
Set lrs = coConex.CargaRecordSet(lsSQL)
    If lrs.EOF And lrs.BOF Then
            MsgBox "No se encuentran Calendario de Negociacion ", vbInformation, "Aviso"
            Exit Function
    End If
Set dObtieneDatosNegociaCuotas = lrs
Set lrs = Nothing
    
Exit Function

ErrorRecupera:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
        
    
End Function

