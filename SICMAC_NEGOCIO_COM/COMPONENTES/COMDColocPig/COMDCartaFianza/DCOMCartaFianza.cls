VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMCartaFianza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private gConsPersona As String
Private gConsComunes As String
Private gConsImagenes As String
Dim coConex As New COMConecta.DCOMConecta
Dim oError As New COMConecta.COMErrorHandling


Public Function ExisteTarifario(ByVal pnCod As Integer) As Boolean
Dim sql As String
Dim Rs As New ADODB.Recordset
'Set coConex = New COMConecta.DCOMConecta
sql = "select cTarifCod from ColocCfTarifario where cTarifCod=" & pnCod
coConex.AbreConexion
Set Rs = coConex.CargaRecordSet(sql)
'Co.cierraConexion
If Rs.EOF And Rs.BOF Then
   ExisteTarifario = False
Else
   ExisteTarifario = True
End If
Set Rs = Nothing
Set coConex = Nothing
End Function
Public Sub InsertaTarifario(ByVal pnTasa, _
pnMonto As Double, ByVal pnCod As Integer, _
ByVal pnMoneda As Integer, ByVal pnModalidad As Integer, pnMontoMax As Double)
Dim sql As String
Dim R As ADODB.Recordset
Dim sNumCod As String
'Dim coConex As COMConecta.DCOMConecta
'Set coConex = New COMConecta.DCOMConecta

coConex.AbreConexion
sql = "Select ISNULL(MAX(cTarifCod),0) as nCod From ColocCfTarifario"
Set R = New ADODB.Recordset
R.Open sql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
sNumCod = R!nCod + 1
R.Close
Set R = Nothing

sql = " Insert ColocCfTarifario"
sql = sql & " (cTarifCod,nModalidad,nMoneda,nTasaTrim,nMontoMinimo,nMontoMax)"
sql = sql & " Values"
sql = sql & " (" & sNumCod & "," & pnModalidad & "," & pnMoneda & ","
sql = sql & pnTasa & "," & pnMonto & "," & pnMontoMax & ")"

coConex.BeginTrans
    coConex.Ejecutar (sql)
coConex.CommitTrans
'Co.cierraConexion
Set coConex = Nothing
End Sub
Public Sub ActualizaTarifario(ByVal pnTasa, _
pnMonto As Double, ByVal pnCod As Integer, ByVal pnMontoMax As Double)
'Dim coConex As COMConecta.DCOMConecta
Dim sql As String
'Set coConex = New COMConecta.DCOMConecta
sql = " Update ColocCfTarifario"
sql = sql & " set nTasaTrim=" & pnTasa
sql = sql & " , nMontoMinimo=" & pnMonto
sql = sql & " , nMontoMax=" & pnMontoMax
sql = sql & " where cTarifCod=" & pnCod

coConex.AbreConexion
coConex.BeginTrans
    coConex.Ejecutar (sql)
coConex.CommitTrans
'Co.cierraConexion
Set coConex = Nothing
End Sub

Private Sub Class_Initialize()
    Dim loIni As New COMConecta.DCOMClasIni
    Dim csConexion As String
    csConexion = loIni.CadenaConexion
    gConsPersona = loIni.BasePersonas
    gConsComunes = loIni.BaseComunes
    gConsImagenes = loIni.BaseImagenes
    Set loIni = Nothing
    If coConex.AbreConexion(csConexion) = False Then
        Call oError.RaiseError(oError.MyUnhandledError, "DColPContrato:Initialize. Error en Conexion a Base de datos")
    End If
End Sub
Public Function RecuperaCF_Tarifario() As ADODB.Recordset
'Dim coConex As COMConecta.DCOMConecta
Dim sql As String
Dim Rs As New ADODB.Recordset
'Set coConex = New COMConecta.DCOMConecta
sql = "select cTarifCod,"
sql = sql & " (select cConsDescripcion from Constante where nConsCod=3402 and nConsValor=nModalidad) Modalidad,"
sql = sql & " (Select cConsDescripcion from Constante where nConsCod=1011 and nConsValor=nMoneda) Moneda,"
sql = sql & " nTasaTrim, nMontoMinimo, nMontoMAX, nModalidad, nMoneda from ColocCfTarifario"
coConex.AbreConexion
Set Rs = coConex.CargaRecordSet(sql)
'Co.cierraConexion
Set RecuperaCF_Tarifario = Rs
Set coConex = Nothing
Set Rs = Nothing
End Function

'MADM 20120302 - 20111020
Public Function RecuperaCartaFianzaSolicitud(ByVal psCtaCod As String) As ADODB.Recordset
Dim lsSQL As String
'Dim oCon As COMConecta.DCOMConecta

lsSQL = " stp_sel_SolicitudCartaFianza '" & psCtaCod & "'"

'    lsSQL = "Select CCF.nCondicion nColocCondicion, P.cPersCod, P.cPersNombre, " _
'        & " CE.nMonto, CE.dVenc, CE.dPrdEstado, " _
'        & " cAnalista = (Select cPersCod From ProductoPersona Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & "), " _
'        & " CF.cNumFuente, PP2.cPersCod as cPersAcreedor, P2.cPersNombre as cPersNomAcre, " _
'        & " CCF.dAsignacion,CCF.dVencimiento,CCF.nModalidad, CCF.cFinalidad,  " _
'        & " cPersAvalado = (Select isnull(PP.cPersCod,'')  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAval & ") , " _
'        & " cAvalNombre = (Select isnull(cPersNombre,'')  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAval & ")  " _
'        & " From ColocCartaFianza CCF " _
'        & " INNER JOIN ProductoPersona PP ON CCF.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular _
'        & " INNER JOIN ProductoPersona PP2 ON CCF.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAcreedor _
'        & " INNER JOIN ProductoPersona PP3 ON CCF.cCtaCod = PP3.cCtaCod AND PP3.nPrdPersRelac = " & gColRelPersAval _
'        & " INNER JOIN Producto Prd ON Prd.cCtaCod = CCF.cCtaCod " _
'        & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod " _
'        & " INNER JOIN Persona P2 ON PP2.cPersCod = P2.cPersCod " _
'        & " INNER JOIN ColocCFEstado CE ON CCF.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstSolic _
'        & " INNER JOIN ColocFteIngreso CF ON CCF.cCtaCod = CF.cCtaCod " _
'        & "  " _
'        & "  " _
'        & " WHERE CCF.cCtaCod = '" & psCtaCod & "' AND Prd.nPrdEstado = " & gColocEstSolic
    'Set coConex = New COMConecta.DCOMConecta
    coConex.AbreConexion
    Set RecuperaCartaFianzaSolicitud = coConex.CargaRecordSet(lsSQL)
    'oCon.cierraConexion
    Set coConex = Nothing
End Function

'MADM 20101020
Public Function RecuperaCartaFianzaSugerencia(ByVal psCtaCod As String) As ADODB.Recordset
Dim lsSQL As String
'Dim oCon As COMConecta.DCOMConecta

'MAVM 20100608 BAS II
'->*****LUCV20160919, Comentó. Según ERS004-2016
''WIOR 20120329-INICIO-CONSORCIOS
'lsSQL = " declare @cPersCodTitular as varchar(13) "
'lsSQL = lsSQL & " set @cPersCodTitular  =(select cPersCod from ProductoPersona where cCtaCod ='" & psCtaCod & "' and nPrdPersRelac=20) "
''WIOR 20120329
'
'lsSQL = lsSQL & "Select CCF.nCondicion nColocCondicion, P.cPersCod, P.cPersNombre, " _
'        & " CE.nMonto as nMontoSol, CE.dVenc as dVencSol, CE.dPrdEstado, " _
'        & " cAnalista = (Select cPersNombre From ProductoPersona PP1 INNER JOIN Persona P1 ON PP1.cPersCod = P1.cPersCod " _
'        & " Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ") , " _
'        & " CF.cNumFuente, PP2.cPersCod as cPersAcreedor, P2.cPersNombre as cPersNomAcre, " _
'        & " CCF.dAsignacion,CCF.dVencimiento,CCF.nModalidad, CCF.cFinalidad ,  " _
'        & " nMontoSug = (Select IsNull(nMonto,CE.nMonto) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " ) , " _
'        & " dVencSug = (Select IsNull(dVenc,CE.dVenc) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " ),  " _
'        & " RTrim(ISNULL(cTpoCredCod, '')) cTpoCredCod, CO.cConsDescripcion, nTpoInstCorp" _
'        & " ,cPersAvalado = (Select isnull(PP.cPersCod,'')  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAval & ") , "
''lsSQL = lsSQL & "   cAvalNombre = (Select isnull(cPersNombre,'')  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAval & ")  "
''WIOR 20120329-INICIO-CONSORCIOS
'lsSQL = lsSQL & "   cAvalNombre = (Select case when PP.cPersCod =@cPersCodTitular AND CCF.cConsorcio IS NOT NULL then isnull(CCF.cConsorcio,'') Else isnull(cPersNombre,'') end  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAval & ")  "
''WIOR 20120329
'lsSQL = lsSQL & "   From ColocCartaFianza CCF " _
'        & " INNER JOIN ProductoPersona PP ON CCF.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular _
'        & " INNER JOIN ProductoPersona PP2 ON CCF.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAcreedor _
'        & " INNER JOIN Producto Prd ON Prd.cCtaCod = CCF.cCtaCod " _
'        & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod " _
'        & " INNER JOIN Persona P2 ON PP2.cPersCod = P2.cPersCod " _
'        & " INNER JOIN ColocCFEstado CE ON CCF.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstSolic _
'        & " INNER JOIN ColocFteIngreso CF ON CCF.cCtaCod = CF.cCtaCod " _
'        & " INNER JOIN Colocaciones C on CCF.cCtaCod = C.cCtaCod LEFT JOIN Constante CO on C.cTpoCredCod = CO.nConsValor and nConsCod = 3034 and nConsValor <>3034 " _
'        & " WHERE CCF.cCtaCod = '" & psCtaCod & "' AND Prd.nPrdEstado in (" & gColocEstSolic & "," & gColocEstSug & ") "
'
'    'Set coConex = New COMConecta.DCOMConecta
'    coConex.AbreConexion
'    Set RecuperaCartaFianzaSugerencia = coConex.CargaRecordSet(lsSQL)
'    'oCon.cierraConexion
'    Set coConex = Nothing
'<-*****Fin, LUCV20160919
    lsSQL = " exec stp_sel_RecuperaCartaFianzaSugerencia '" & psCtaCod & "' "
    coConex.AbreConexion
    Set RecuperaCartaFianzaSugerencia = coConex.CargaRecordSet(lsSQL)
    Set coConex = Nothing
End Function

Public Function RecuperaCartaFianzaAprobacion(ByVal psCtaCod As String) As ADODB.Recordset
Dim lsSQL As String
'Dim oCon As COMConecta.DCOMConecta
    'MADM 20111020 AVAL -- 'MAVM 20100605 BAS II
'->*****LUCV20160919, Comentó. Según ERS004-2016
'    'WIOR 20120329-INICIO-CONSORCIOS
'    lsSQL = " declare @cPersCodTitular as varchar(13) "
'    lsSQL = lsSQL & " set @cPersCodTitular  =(select cPersCod from ProductoPersona where cCtaCod ='" & psCtaCod & "' and nPrdPersRelac=20) "
'    'WIOR 20120329
'
'    lsSQL = lsSQL & " Select CCF.nCondicion nColocCondicion, P.cPersCod, P.cPersNombre, " _
'        & " CE.nMonto as nMontoSug, CE.dVenc as dVencSug, CE.dPrdEstado, ciiu.cCIIUdescripcion, " _
'        & " cAnalista = (Select cPersNombre  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ") , " _
'        'CF.cNumFuente, PFI.cRazSocDescrip, 'LUCV20160914
'        lsSQL = lsSQL & " PP2.cPersCod as cPersAcreedor, P2.cPersNombre as cPersNomAcre, " _
'        & " CCF.dAsignacion,CCF.dVencimiento,CCF.nModalidad, CCF.cFinalidad ,  " _
'        & " nMontoSug = (Select IsNull(nMonto,CE.nMonto) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " ) , " _
'        & " dVencSug = (Select IsNull(dVenc,CE.dVenc) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " ) , " _
'        & " cTpoCredCod, CO.cConsDescripcion, " _
'        & " cAvalCod = (Select isnull(PP.cPersCod,'')  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAval & ") , "
'        'lsSQL = lsSQL & " cAvalNombre = (Select isnull(cPersNombre,'')  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAval & ")  "
'        'WIOR 20120329-INICIO-CONSORCIOS
'        lsSQL = lsSQL & " cAvalNombre = (Select case when PP.cPersCod =@cPersCodTitular  AND CCF.cConsorcio IS NOT NULL then isnull(CCF.cConsorcio,'') Else isnull(cPersNombre,'') end From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAval & ")  "
'        'WIOR 20120329
'        lsSQL = lsSQL & " From ColocCartaFianza CCF " _
'        & " INNER JOIN ProductoPersona PP ON CCF.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular _
'        & " INNER JOIN ProductoPersona PP2 ON CCF.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAcreedor _
'        & " INNER JOIN Producto Prd ON Prd.cCtaCod = CCF.cCtaCod " _
'        & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod " _
'        & " INNER JOIN Persona P2 ON PP2.cPersCod = P2.cPersCod " _
'        & " INNER JOIN ColocCFEstado CE ON CCF.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstSug _
'        ' INNER JOIN ColocFteIngreso CF ON CCF.cCtaCod = CF.cCtaCod"  'LUCV20160914
'        ' INNER JOIN PersFteIngreso PFI ON PFI.cNumFuente = CF.cNumFuente 'LUCV20160914
'        lsSQL = lsSQL & " LEFT JOIN Ciiu ciiu ON P.cPersCIIU = ciiu.cCIIUcod " _
'        & " INNER JOIN Colocaciones C on CCF.cCtaCod = C.cCtaCod LEFT JOIN Constante CO on C.cTpoCredCod = CO.nConsValor and nConsCod = 3034 and nConsValor <>3034 " _
'        & " WHERE CCF.cCtaCod = '" & psCtaCod & "' AND Prd.nPrdEstado = " & gColocEstSug
'    'Set coConex = New COMConecta.DCOMConecta
'    coConex.AbreConexion
'    Set RecuperaCartaFianzaAprobacion = coConex.CargaRecordSet(lsSQL)
'    'oCon.cierraConexion
'    Set coConex = Nothing
'<-*****LUCV20160919

    lsSQL = " exec stp_sel_RecuperaCartaFianzaAprobacion '" & psCtaCod & "' "
    coConex.AbreConexion
    Set RecuperaCartaFianzaAprobacion = coConex.CargaRecordSet(lsSQL)
    Set coConex = Nothing
End Function

Public Function RecuperaCartaFianzaRechazo(ByVal psCtaCod As String) As ADODB.Recordset
Dim lsSQL As String
'Dim oCon As COMConecta.DCOMConecta
'->*****LUCV20160919, Comentó Según ERS004-2016
'    'MAVM 20100606 BAS II
'    lsSQL = "Select CCF.nCondicion nColocCondicion, P.cPersCod, P.cPersNombre, " _
'        & " CE.nMonto as nMontoSol, CE.dVenc as dVencSol, CE.dPrdEstado, " _
'        & " cAnalista = (Select cPersNombre  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ") , " _
'        & " CF.cNumFuente, PP2.cPersCod as cPersAcreedor, P2.cPersNombre as cPersNomAcre, " _
'        & " CCF.dAsignacion,CCF.dVencimiento,CCF.nModalidad, CCF.cFinalidad ,  " _
'        & " nMontoSug = (Select IsNull(nMonto,CE.nMonto) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " ) , " _
'        & " dVencSug = (Select IsNull(dVenc,CE.dVenc) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " ),  " _
'        & " cTpoCredCod, CO.cConsDescripcion" _
'        & " From ColocCartaFianza CCF " _
'        & " INNER JOIN ProductoPersona PP ON CCF.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular _
'        & " INNER JOIN ProductoPersona PP2 ON CCF.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAcreedor _
'        & " INNER JOIN Producto Prd ON Prd.cCtaCod = CCF.cCtaCod " _
'        & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod " _
'        & " INNER JOIN Persona P2 ON PP2.cPersCod = P2.cPersCod " _
'        & " INNER JOIN ColocCFEstado CE ON CCF.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstSolic _
'        & " INNER JOIN ColocFteIngreso CF ON CCF.cCtaCod = CF.cCtaCod " _
'        & " INNER JOIN Colocaciones C on CCF.cCtaCod = C.cCtaCod LEFT JOIN Constante CO on C.cTpoCredCod = CO.nConsValor and nConsCod = 3034 and nConsValor <>3034 " _
'        & " WHERE CCF.cCtaCod = '" & psCtaCod & "' AND Prd.nPrdEstado in ( " & gColocEstSolic & "," & gColocEstSug & ") "
'    'Set coConex = New COMConecta.DCOMConecta
'    coConex.AbreConexion
'    Set RecuperaCartaFianzaRechazo = coConex.CargaRecordSet(lsSQL)
'    'coConex.cierraConexion
'    Set coConex = Nothing
'<-*****LUCV20160919
   lsSQL = " exec stp_sel_RecuperaCartaFianzaRechazo '" & psCtaCod & "' "
    coConex.AbreConexion
    Set RecuperaCartaFianzaRechazo = coConex.CargaRecordSet(lsSQL)
    Set coConex = Nothing
End Function
    
Public Function RecuperaCartaFianzaComision(ByVal psCtaCod As String) As ADODB.Recordset
Dim lsSQL As String
Dim Rs As ADODB.Recordset
Dim lsEstado As String
'Dim coConex As COMConecta.DCOMConecta
coConex.AbreConexion

lsSQL = "select count(*) NumReg from ColocCFEstado where cCtaCod = '" & psCtaCod & "' and nPrdEstado='" & gColocEstRenovada & "'"
Set Rs = New ADODB.Recordset
   Set Rs = coConex.CargaRecordSet(lsSQL)
   If Rs!NumReg <> 0 Then
        lsEstado = Trim(gColocEstRenovada)
   Else
        'lsEstado = Trim(gColocEstAprob)
        'WIOR 20120828
        lsSQL = "select TOP 1 nPrdEstado from ColocCFEstado where cCtaCod ='" & psCtaCod & "' order by dPrdEstado DESC"
        Set Rs = New ADODB.Recordset
        Set Rs = coConex.CargaRecordSet(lsSQL)
        If Rs.RecordCount > 0 Then
            If Trim(Rs!nPrdEstado) = gColocEstModificada Then
                lsEstado = Trim(gColocEstModificada)
            Else
                lsEstado = Trim(gColocEstAprob)
            End If
        Else
             lsEstado = Trim(gColocEstAprob)
        End If
   End If
Set Rs = Nothing

    'MAVM 20100606 BAS II
    lsSQL = "Select CCF.cCtaCod, CCF.nCondicion nColocCondicion, P.cPersCod, P.cPersNombre, Prd.nPrdEstado, " _
        & " CE.nMonto as nMontoApr, CE.dVenc as dVencApr, CE.dPrdEstado, " _
        & " cAnalista = (Select cPersNombre  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ") , " _
        & " cApoderado= (Select cPersCod From ProductoPersona Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersApoderado & ") , " _
        & " PP2.cPersCod as cPersAcreedor, P2.cPersNombre as cPersNomAcre, " _
        & " dAsignacion=(CASE WHEN CE.dPrdEstado ='" & gColocEstRenovada & " ' THEN CCF.dFecRenova  ELSE  CCF.dAsignacion END )" _
        & " ,CCF.dVencimiento,CCF.nModalidad, CCF.cFinalidad, CCF.nRenovacion,  " _
        & " cTpoCredCod, CO.cConsDescripcion" _
        & " From ColocCartaFianza CCF " _
        & " INNER JOIN ProductoPersona PP ON CCF.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular _
        & " INNER JOIN ProductoPersona PP2 ON CCF.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAcreedor _
        & " INNER JOIN Producto Prd ON Prd.cCtaCod = CCF.cCtaCod " _
        & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod " _
        & " INNER JOIN Persona P2 ON PP2.cPersCod = P2.cPersCod " _
        & " INNER JOIN ColocCFEstado CE ON CCF.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = '" & lsEstado & "'" _
        & " INNER JOIN Colocaciones C on CCF.cCtaCod = C.cCtaCod INNER JOIN Constante CO on C.cTpoCredCod = CO.nConsValor and nConsCod = 3034 and nConsValor <>3034 " _
        & " WHERE CCF.cCtaCod = '" & psCtaCod & "' AND Prd.nPrdEstado in ( " & gColocEstAprob & "," & gColocEstVigNorm & "," & gColocEstRenovada & " ) ORDER BY CE.nPrdEstado"
        
    Set RecuperaCartaFianzaComision = coConex.CargaRecordSet(lsSQL)
    Set coConex = Nothing
End Function
   'MADM 20111025 - AVAL
Public Function RecuperaCartaFianzaEmision(ByVal psCtaCod As String) As ADODB.Recordset
Dim lsSQL As String
Dim Rs As ADODB.Recordset
Dim lsEstado As String

lsSQL = "select count(*) NumReg from ColocCFEstado where cCtaCod = '" & psCtaCod & "' and nPrdEstado='" & gColocEstRenovada & "'"
Set Rs = New ADODB.Recordset
   Set Rs = coConex.CargaRecordSet(lsSQL)
   If Rs!NumReg <> 0 Then
        lsEstado = Trim(gColocEstRenovada)
   Else
        'lsEstado = Trim(gColocEstAprob)
        'WIOR 20120828
        lsSQL = "select TOP 1 nPrdEstado from ColocCFEstado where cCtaCod ='" & psCtaCod & "' order by dPrdEstado DESC"
        Set Rs = New ADODB.Recordset
        Set Rs = coConex.CargaRecordSet(lsSQL)
        If Rs.RecordCount > 0 Then
            If Trim(Rs!nPrdEstado) = gColocEstModificada Then
                lsEstado = Trim(gColocEstModificada)
            Else
                lsEstado = Trim(gColocEstAprob)
            End If
        Else
             lsEstado = Trim(gColocEstAprob)
        End If
   End If
Set Rs = Nothing

    'WIOR 20120329-INICIO-CONSORCIOS
    lsSQL = " declare @cPersCodTitular as varchar(13) "
    lsSQL = lsSQL & " set @cPersCodTitular  =(select cPersCod from ProductoPersona where cCtaCod ='" & psCtaCod & "' and nPrdPersRelac=20) "
    'WIOR 20120329
    lsSQL = lsSQL & "Select CCF.nCondicion nColocCondicion, P.cPersCod, P.cPersNombre, " _
        & " CE.nMonto as nMontoApr, CE.dVenc as dVencApr, CE.dPrdEstado, " _
        & " cAnalista = (Select cPersNombre  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ") , " _
        & " cApoderado= (Select cPersNombre  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersApoderado & ") , " _
        & " CF.cNumFuente, PP2.cPersCod as cPersAcreedor, P2.cPersNombre as cPersNomAcre, " _
        & " CCF.dAsignacion,CCF.dVencimiento,CCF.nModalidad, CCF.cFinalidad,Prd.nPrdEstado,nRenovacion,   " _
        & " cTpoCredCod, CO.cConsDescripcion," _
        & " cAvalCod = (Select isnull(PP.cPersCod,'')  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAval & ") , " _
        '& " cAvalNombre = (Select isnull(cPersNombre,'')  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAval & ")  "
        'lsSQL = lsSQL & " cAvalNombre = (Select isnull(cPersNombre,'')  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAval & ")  "
        
        'WIOR 20120329-INICIO-CONSORCIOS
        lsSQL = lsSQL & " cAvalNombre = (Select case when PP.cPersCod =@cPersCodTitular  AND CCF.cConsorcio IS NOT NULL then isnull(CCF.cConsorcio,'') Else isnull(cPersNombre,'') end From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAval & ")  "
        'WIOR 20120329
        'JOEP20181222 CP
        lsSQL = lsSQL & ", OtrsModalidades = ISNULL(CCF.cOtrsModalidades,'')"
        'JOEP20181222 CP
        lsSQL = lsSQL & " From ColocCartaFianza CCF " _
        & " INNER JOIN ProductoPersona PP ON CCF.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular _
        & " INNER JOIN ProductoPersona PP2 ON CCF.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAcreedor _
        & " INNER JOIN Producto Prd ON Prd.cCtaCod = CCF.cCtaCod " _
        & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod " _
        & " INNER JOIN Persona P2 ON PP2.cPersCod = P2.cPersCod " _
        & " INNER JOIN ColocCFEstado CE ON CCF.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & lsEstado _
        & " LEFT JOIN ColocFteIngreso CF ON CCF.cCtaCod = CF.cCtaCod " _
        & " INNER JOIN Colocaciones C on CCF.cCtaCod = C.cCtaCod LEFT JOIN Constante CO on C.cTpoCredCod = CO.nConsValor and nConsCod = 3034 and nConsValor <>3034 " _
        & " WHERE CCF.cCtaCod = '" & psCtaCod & "' AND Prd.nPrdEstado IN  ( " & gColocEstAprob & ", " & gColocEstRenovada & ",2020) ORDER BY CE.nPrdEstado "
    
    'Set coConex = New COMConecta.DCOMConecta
    coConex.AbreConexion
    Set RecuperaCartaFianzaEmision = coConex.CargaRecordSet(lsSQL)
    'coConex.cierraConexion
    Set coConex = Nothing
End Function
        
Public Function RecuperaCartaFianzaRetirar(ByVal psCtaCod As String) As ADODB.Recordset
Dim lsSQL As String
'Dim coConex As COMConecta.DCOMConecta

'->***** LUCV20160919, Comentó. Según ERS004-2016
'MAVM 20100606 BAS II
'    lsSQL = "Select CCF.nCondicion nColocCondicion, P.cPersCod, P.cPersNombre, " _
'        & " CE.nMonto as nMontoSol, CE.dVenc as dVencSol, CE.dPrdEstado, " _
'        & " cAnalista = (Select cPersNombre  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ") , " _
'        & " CF.cNumFuente, PP2.cPersCod as cPersAcreedor, P2.cPersNombre as cPersNomAcre, " _
'        & " CCF.dAsignacion,CCF.dVencimiento,CCF.nModalidad, CCF.cFinalidad ,  " _
'        & " nMontoSug = (Select IsNull(nMonto,CE.nMonto) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " ) , " _
'        & " dVencSug = (Select IsNull(dVenc,CE.dVenc) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " ) , " _
'        & " cTpoCredCod, CO.cConsDescripcion" _
'        & " From ColocCartaFianza CCF " _
'        & " INNER JOIN ProductoPersona PP ON CCF.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular _
'        & " INNER JOIN ProductoPersona PP2 ON CCF.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAcreedor _
'        & " INNER JOIN Producto Prd ON Prd.cCtaCod = CCF.cCtaCod " _
'        & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod " _
'        & " INNER JOIN Persona P2 ON PP2.cPersCod = P2.cPersCod " _
'        & " INNER JOIN ColocCFEstado CE ON CCF.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstSolic _
'        & " INNER JOIN ColocFteIngreso CF ON CCF.cCtaCod = CF.cCtaCod " _
'        & " INNER JOIN Colocaciones C on CCF.cCtaCod = C.cCtaCod LEFT JOIN Constante CO on C.cTpoCredCod = CO.nConsValor and nConsCod = 3034 and nConsValor <>3034 " _
'        & " WHERE CCF.cCtaCod = '" & psCtaCod & "' AND Prd.nPrdEstado in ( " & gColocEstAprob & ") "
'    'Set coConex = New COMConecta.DCOMConecta
'    coConex.AbreConexion
'    Set RecuperaCartaFianzaRetirar = coConex.CargaRecordSet(lsSQL)
'    'coConex.cierraConexion
'    Set coConex = Nothing
'<-*****Fin, LUCV20160919
    lsSQL = " exec stp_sel_RecuperaCartaFianzaRetirar '" & psCtaCod & "' "
    coConex.AbreConexion
    Set RecuperaCartaFianzaRetirar = coConex.CargaRecordSet(lsSQL)
    Set coConex = Nothing
End Function
        
Public Function RecuperaCartaFianzaDevolucion(ByVal psCtaCod As String) As ADODB.Recordset
Dim lsSQL As String
'Dim coConex As COMConecta.DCOMConecta
'->*****LUCV20160919, Comentó. Según ERS004-2016
'    lsSQL = "Select CCF.nCondicion nColocCondicion, P.cPersCod, P.cPersNombre, Prd.nPrdEstado, " _
'        & " CE.nMonto as nMontoSol, CE.dVenc as dVencSol, CE.dPrdEstado, " _
'        & " cAnalista = (Select cPersNombre  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ") ,  " _
'        & " CF.cNumFuente, PP2.cPersCod as cPersAcreedor, P2.cPersNombre as cPersNomAcre, " _
'        & " CCF.dAsignacion,CCF.dVencimiento,CCF.nModalidad, CCF.cFinalidad ,  " _
'        & " nMontoSug = (Select IsNull(nMonto,CE.nMonto) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " ) , " _
'        & " dVencSug = (Select IsNull(dVenc,CE.dVenc) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " ) , " _
'        & " cTpoCredCod, CO.cConsDescripcion" _
'        & " From ColocCartaFianza CCF " _
'        & " INNER JOIN ProductoPersona PP ON CCF.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular _
'        & " INNER JOIN ProductoPersona PP2 ON CCF.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAcreedor _
'        & " INNER JOIN Producto Prd ON Prd.cCtaCod = CCF.cCtaCod " _
'        & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod " _
'        & " INNER JOIN Persona P2 ON PP2.cPersCod = P2.cPersCod " _
'        & " INNER JOIN ColocCFEstado CE ON CCF.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstSolic _
'        & " INNER JOIN ColocFteIngreso CF ON CCF.cCtaCod = CF.cCtaCod " _
'        & " INNER JOIN Colocaciones C on CCF.cCtaCod = C.cCtaCod LEFT JOIN Constante CO on C.cTpoCredCod = CO.nConsValor and nConsCod = 3034 and nConsValor <>3034 " _
'        & " WHERE CCF.cCtaCod = '" & psCtaCod & "' AND Prd.nPrdEstado in ( " & gColocEstVigNorm & "," & gColocEstRenovada & " ) "
'    'By Capi Octubre-2007 para que cancele tambien los que estan en estado renovado
'    'Set coConex = New COMConecta.DCOMConecta
'    coConex.AbreConexion
'        Set RecuperaCartaFianzaDevolucion = coConex.CargaRecordSet(lsSQL)
'    'coConex.cierraConexion
'    Set coConex = Nothing
'<-*****LUCV20160919
    lsSQL = " exec stp_sel_RecuperaCartaFianzaDevolucion '" & psCtaCod & "' "
    coConex.AbreConexion
    Set RecuperaCartaFianzaDevolucion = coConex.CargaRecordSet(lsSQL)
    Set coConex = Nothing
End Function

Public Function RecuperaDatosGarantiaCF(ByVal psCtaCod As String) As ADODB.Recordset
'Dim coConex As COMConecta.DCOMConecta
Dim lsSQL As String

On Error GoTo ErrorRecuperaDatosGarantiaCred
    'MAVM Se agrego la var: gTpoProducto BAS II
    lsSQL = "Select CN3.cConsDescripcion as cTipoCredDescrip, C.nModalidad as nColocDestino, CN.cConsDescripcion as cDestinoDescripcion , " _
    & " P.cPersCod, P.cPersNombre, CN2.cConsDescripcion as cMonedaDesc, " _
    & " DNI=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & gPersIdDNI & "'), " _
    & " CE.nMonto " _
    & " From ColocCartaFianza C INNER JOIN ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular _
    & "                    INNER JOIN " & gConsPersona & "Persona P ON PP.cPersCod = P.cPersCod " _
    & "                    INNER JOIN ColocCFEstado CE ON C.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = '" & gColocEstSolic & "' " _
    & "                    INNER JOIN Constante CN ON C.nModalidad = CN.nConsValor AND CN.nConsCod = " & gColCFModalidad _
    & "                    INNER JOIN Constante CN2 ON convert(int,substring(C.cCtaCod,9,1)) = CN2.nConsValor AND CN2.nConsCod = " & gMoneda _
    & "                    INNER JOIN Constante CN3 ON convert(int,substring(C.cCtaCod,6,3)) = CN3.nConsValor AND CN3.nConsCod in (" & gProducto & ", " & gTpoProducto & ")" _
    & " WHERE C.cCtaCod = '" & psCtaCod & "' "
    
    'Set coConex = New COMConecta.DCOMConecta
    coConex.AbreConexion
    Set RecuperaDatosGarantiaCF = coConex.CargaRecordSet(lsSQL)
    'coConex.cierraConexion
    Set coConex = Nothing
    
    Exit Function

ErrorRecuperaDatosGarantiaCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
   
Public Function RecuperaCartaFianzaHonrar(ByVal psCtaCod As String) As ADODB.Recordset
Dim lsSQL As String
'Dim coConex As COMConecta.DCOMConecta
'->*****LUCV20160919, Comentó. Según ERS004-2016
'    'MAVM 20100606 BAS II
'    lsSQL = "Select CCF.nCondicion nColocCondicion, P.cPersCod, P.cPersNombre, " _
'        & " CE.nMonto as nMontoSol, CE.dVenc as dVencSol, CE.dPrdEstado, " _
'        & " cAnalista = (Select cPersNombre  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ") , " _
'        & " CF.cNumFuente, PP2.cPersCod as cPersAcreedor, P2.cPersNombre as cPersNomAcre, " _
'        & " CCF.dAsignacion,CCF.dVencimiento,CCF.nModalidad, CCF.cFinalidad ,  " _
'        & " nMontoSug = (Select IsNull(nMonto,CE.nMonto) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " ) , " _
'        & " dVencSug = (Select IsNull(dVenc,CE.dVenc) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " ),  " _
'        & " cTpoCredCod, CO.cConsDescripcion" _
'        & " From ColocCartaFianza CCF " _
'        & " INNER JOIN ProductoPersona PP ON CCF.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular _
'        & " INNER JOIN ProductoPersona PP2 ON CCF.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAcreedor _
'        & " INNER JOIN Producto Prd ON Prd.cCtaCod = CCF.cCtaCod " _
'        & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod " _
'        & " INNER JOIN Persona P2 ON PP2.cPersCod = P2.cPersCod " _
'        & " INNER JOIN ColocCFEstado CE ON CCF.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstSolic _
'        & " INNER JOIN ColocFteIngreso CF ON CCF.cCtaCod = CF.cCtaCod " _
'        & " INNER JOIN Colocaciones C on CCF.cCtaCod = C.cCtaCod LEFT JOIN Constante CO on C.cTpoCredCod = CO.nConsValor and nConsCod = 3034 and nConsValor <>3034 " _
'        & " WHERE CCF.cCtaCod = '" & psCtaCod & "' AND Prd.nPrdEstado in ( " & gColocEstVigNorm & ") "
'
'    'Set coConex = New COMConecta.DCOMConecta
'    coConex.AbreConexion
'        Set RecuperaCartaFianzaHonrar = coConex.CargaRecordSet(lsSQL)
'    'coConex.cierraConexion
'    Set coConex = Nothing
'<-*****Fin, LUCV20160919
    lsSQL = " exec stp_sel_RecuperaCartaFianzaHonrar '" & psCtaCod & "' "
    coConex.AbreConexion
    Set RecuperaCartaFianzaHonrar = coConex.CargaRecordSet(lsSQL)
    Set coConex = Nothing
End Function

Public Function VerificaComision(ByVal psCodCta As String) As Long
Dim lsSQL As String
'Dim coConex As COMConecta.DCOMConecta
Dim Rs As ADODB.Recordset

    'lsSQL = "Select M.nMovNro from MOV M INNER JOIN MOVCOL MC ON M.nMovNro = MC.nMovNro " _
    '    & " Where M.cOpeCod = '" & gColCFOpeComisEfe & "' And nMovFlag = 0 AND cCtaCod = '" & psCodCta & "'"
    lsSQL = "Exec stp_sel_VerificaComisionEmisionCF '" & psCodCta & "'" 'JUEZ 20160226
    
    'Set coConex = New COMConecta.DCOMConecta
    'coConex.AbreConexion
        Set Rs = coConex.CargaRecordSet(lsSQL)
    If Rs.EOF And Rs.BOF Then
        VerificaComision = -1
    Else
        VerificaComision = Rs!nMovNro
    End If
    
    Set Rs = Nothing
    'coConex.cierraConexion
    Set coConex = Nothing

End Function

Public Function RecuperaCartaFianzaHonrarCredito(ByVal psCtaCod As String) As ADODB.Recordset
Dim lsSQL As String
'Dim coConex As COMConecta.DCOMConecta

'->*****LUCV20160919, Comentó según ERS004-2016
'MAVM BAS II
'    lsSQL = "Select CCF.nCondicion nColocCondicion, P.cPersCod, P.cPersNombre, " _
'        & " CE.nMonto as nMontoSol, CE.dVenc as dVencSol, CE.dPrdEstado, " _
'        & " cAnalista = (Select cPersNombre  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ") , " _
'        & " CF.cNumFuente, PP2.cPersCod as cPersAcreedor, P2.cPersNombre as cPersNomAcre, " _
'        & " CCF.dAsignacion,CCF.dVencimiento,CCF.nModalidad, CCF.cFinalidad ,  " _
'        & " nMontoSug = (Select IsNull(nMonto,CE.nMonto) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " ) , " _
'        & " dVencSug = (Select IsNull(dVenc,CE.dVenc) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " ) , " _
'        & " cTpoCredCod, CO.cConsDescripcion" _
'        & " From ColocCartaFianza CCF " _
'        & " INNER JOIN ProductoPersona PP ON CCF.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular _
'        & " INNER JOIN ProductoPersona PP2 ON CCF.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAcreedor _
'        & " INNER JOIN Producto Prd ON Prd.cCtaCod = CCF.cCtaCod " _
'        & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod " _
'        & " INNER JOIN Persona P2 ON PP2.cPersCod = P2.cPersCod " _
'        & " INNER JOIN ColocCFEstado CE ON CCF.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstSolic _
'        & " INNER JOIN ColocFteIngreso CF ON CCF.cCtaCod = CF.cCtaCod " _
'        & " INNER JOIN Colocaciones C on CCF.cCtaCod = C.cCtaCod LEFT JOIN Constante CO on C.cTpoCredCod = CO.nConsValor and nConsCod = 3034 and nConsValor <>3034 " _
'        & " WHERE CCF.cCtaCod = '" & psCtaCod & "' AND Prd.nPrdEstado in ( " & gColocEstHonrada & ") "
'
'    'Set coConex = New COMConecta.DCOMConecta
'    coConex.AbreConexion
'        Set RecuperaCartaFianzaHonrarCredito = coConex.CargaRecordSet(lsSQL)
'    'coConex.cierraConexion
'    Set coConex = Nothing
'<-*****Fin, LUCV20160919.
   lsSQL = " exec stp_sel_RecuperaCartaFianzaHonrarCredito '" & psCtaCod & "' "
    coConex.AbreConexion
    Set RecuperaCartaFianzaHonrarCredito = coConex.CargaRecordSet(lsSQL)
    Set coConex = Nothing
End Function

Public Function ValidadCreditoCF(ByVal psCtaCod As String) As Boolean
    Dim sSql As String
    'Dim coConex As COMConecta.DCOMConecta
    Dim Rs As ADODB.Recordset
    Dim bSalida As Boolean
    
    sSql = "Select Count(*) as nCantidad"
    sSql = sSql & " From Producto"
    sSql = sSql & " Where SubString(cCtaCod,7,2)='21' and cCtaCod='" & psCtaCod & "'"
    
    'Set coConex = New COMConecta.DCOMConecta
    'coConex.AbreConexion
    Set Rs = coConex.CargaRecordSet(sSql)
    'coConex.cierraConexion
    Set coConex = Nothing
    
    If Not Rs.EOF And Not Rs.BOF Then
        If Rs!nCantidad > 0 Then
            bSalida = True
        Else
            bSalida = False
        End If
    End If
    ValidadCreditoCF = bSalida
End Function

Public Function ValorCoberturaGarantia() As Double
    Dim sSql As String
    Dim nValor As Double
    Dim Rs As New ADODB.Recordset
    
    sSql = "Select nParamValor"
    sSql = sSql & " From ColocParametro"
    sSql = sSql & " Where nParamVar=1028"
    
    'Set coConex = New COMConecta.DCOMConecta
    coConex.AbreConexion
    Set Rs = coConex.CargaRecordSet(sSql)
    'oConec.cierraConexion
    
    If Not Rs.BOF And Not Rs.EOF Then
        nValor = Rs!nParamValor
    End If
    Set Rs = Nothing
    ValorCoberturaGarantia = nValor
End Function

'--------- Nuevo Proceso CMAC-CUSCO--------------

Public Function ObtieneCartaFianzaPersona(ByVal psPersCod As String) As ADODB.Recordset
    Dim sql As String
    Dim Rs As New ADODB.Recordset
    
    sql = " SELECT PP.cCtaCod, P.nPrdEstado, PP.nPrdPersRelac, T.cConsDescripcion cEstado," & _
          " T1.cConsDescripcion cRelacion, UPPER(T2.cConsDescripcion) cProducto," & _
          " UPPER(T3.cConsDescripcion) cMoneda" & _
          " FROM ProductoPersona PP " & _
          " INNER JOIN Producto P  ON P.cCtaCod = PP.cCtaCod" & _
          " INNER JOIN Constante T ON P.nPrdEstado = T.nConsValor" & _
          " INNER JOIN Constante T1 ON PP.nPrdPersRelac = T1.nConsValor" & _
          " INNER JOIN Constante T2 ON SUBSTRING(PP.cCtaCod,6,3) = CONVERT(Varchar(3),T2.nConsValor)" & _
          " INNER JOIN Constante T3 ON SUBSTRING(PP.cCtaCod,9,1) = CONVERT(Varchar(1),T3.nConsValor)" & _
          " WHERE cPersCod = '" & psPersCod & "' and T1.nConsCod = 3002 AND" & _
          " T.nConsCod = 3001 AND T2.nConsCod = 1001 AND" & _
          " T3.nConsCod = 1011 AND (P.cCtaCod like '_____121%'  or P.cCtaCod like '_____221%')" & _
          " AND PP.nPrdPersRelac ='20' AND  P.nPrdEstado='2020'"
    coConex.AbreConexion
    Set Rs = coConex.CargaRecordSet(sql)
        Set ObtieneCartaFianzaPersona = Rs
    Set Rs = Nothing
    
End Function

Public Function VerficarCartaFianzaRenovacion(ByVal psCtaCod As String) As Boolean
    Dim sql As String
    Dim Rs As New ADODB.Recordset
    
    sql = " select count(cCtaCodAnterior) from ColocCFRenovacion" & _
          " where cCtaCodAnterior = '" & psCtaCod & "'"
          
    Set Rs = coConex.CargaRecordSet(sql)
        If Rs(0) = 0 Then
            VerficarCartaFianzaRenovacion = True
        Else
            VerficarCartaFianzaRenovacion = False
        End If
    Set Rs = Nothing
End Function
Public Function ObtenerCartaFianzaRenovacion(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sql As String
    Dim Rs As New ADODB.Recordset
    
    sql = " select cCtaCodAnterior from ColocCFRenovacion" & _
          " where cCtaCodNueva = '" & psCtaCod & "'"
          
    coConex.AbreConexion
    Set Rs = coConex.CargaRecordSet(sql)
     Set ObtenerCartaFianzaRenovacion = Rs
    Set Rs = Nothing
    
End Function

Private Sub Class_Terminate()
    coConex.CierraConexion
    Set coConex = Nothing
End Sub

Public Function RecuperaCartaFianzaRenovacion(ByVal psCtaCod As String, ByVal dfecha As Date) As ADODB.Recordset
Dim lsSQL As String
Dim Rs As ADODB.Recordset
Dim lsEstado As String

lsSQL = "select count(*) NumReg from ColocCFEstado where cCtaCod = '" & psCtaCod & "' and nPrdEstado='" & gColocEstRenovada & "'"
Set Rs = New ADODB.Recordset
   Set Rs = coConex.CargaRecordSet(lsSQL)
   If Rs!NumReg <> 0 Then
        lsEstado = Trim(gColocEstRenovada)
   Else
        lsEstado = Trim(gColocEstAprob)
   End If
Set Rs = Nothing

    'RECO20150416 ERS001-2015***********************
        ''WIOR 20120330-INICIO-CONSORCIOS
        'lsSQL = " declare @cPersCodTitular as varchar(13) "
        'lsSQL = lsSQL & " set @cPersCodTitular  =(select cPersCod from ProductoPersona where cCtaCod ='" & psCtaCod & "' and nPrdPersRelac=20) "
        ''WIOR 20120330
        '
        ''MADM 20111020---WIOR 20130315 AGREGO C.nMontoCol,C.dVigencia
        '    lsSQL = lsSQL & " Select CCF.cCtaCod, CCF.nCondicion nColocCondicion, P.cPersCod, P.cPersNombre, Prd.nPrdEstado, " _
        '        & " CE.nMonto as nMontoApr, CE.dVenc as dVencApr, CE.dPrdEstado, " _
        '        & " cAnalista = (Select cPersNombre  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ") , " _
        '        & " cApoderado= (Select cPersCod From ProductoPersona Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersApoderado & ") , " _
        '        & " PP2.cPersCod as cPersAcreedor, P2.cPersNombre as cPersNomAcre, " _
        '        & " CCF.dAsignacion,CCF.dVencimiento,CCF.nModalidad, CCF.cFinalidad, CCF.nRenovacion , CP.nPoliza, C.nMontoCol,C.dVigencia dVigCol " _
        '        & " " _
        '        & " , cTpoCredCod, CO.cConsDescripcion," _
        '        & " cPersAvalado = (Select isnull(PP.cPersCod,'')  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAval & ") , "
        '    'lsSQL = lsSQL & " cAvalNombre = (Select isnull(cPersNombre,'')  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAval & ")  "
        '    'WIOR 20120330-INICIO-CONSORCIOS
        '    lsSQL = lsSQL & "   cAvalNombre = (Select case when PP.cPersCod =@cPersCodTitular  AND CCF.cConsorcio IS NOT NULL then isnull(CCF.cConsorcio,'') Else isnull(cPersNombre,'') end  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAval & ")  "
        '    'WIOR 20120330
        '    'WIOR 20120611 se agrego MO.cConsDescripcion sModalidad PARA MOSTRAR LA DESCRIPCION DE LA MODALIDAD PROVENIENTE DE LA CONSTANTE Constante MO 3402
        '    lsSQL = lsSQL & ",MO.cConsDescripcion sModalidad From ColocCartaFianza CCF " _
        '        & " INNER JOIN ProductoPersona PP ON CCF.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular _
        '        & " INNER JOIN ProductoPersona PP2 ON CCF.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAcreedor _
        '        & " INNER JOIN Producto Prd ON Prd.cCtaCod = CCF.cCtaCod " _
        '        & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod " _
        '        & " INNER JOIN Persona P2 ON PP2.cPersCod = P2.cPersCod " _
        '        & " INNER JOIN ColocCFEstado CE ON CCF.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & lsEstado _
        '        & " INNER JOIN ColocCFPoliza CP ON CCF.cCtaCod = CP.cCtaCod " _
        '        & " INNER JOIN Colocaciones C on CCF.cCtaCod = C.cCtaCod LEFT JOIN Constante CO on C.cTpoCredCod = CO.nConsValor and nConsCod = 3034 and nConsValor <> 3034 " _
        '        & " LEFT JOIN Constante MO on CCF.nModalidad = MO.nConsValor and MO.nConsCod = 3402 and MO.nConsValor <> 3402" _
        '        & " WHERE CCF.cCtaCod = '" & psCtaCod & "' AND Prd.nPrdEstado in ( " & gColocEstVigNorm & " ," & gColocEstRenovada & " )" _
        '        & " AND DateDiff(day,dVencimiento,'" & Format(dfecha, "yyyy/mm/dd") & "')  >=0"
        '        'WIOR 20130330 CAMBIO dAsignacion por dVencimiento y 1 por 0
     lsSQL = "Exec stp_sel_RecuperaCartaFianzaRenovacion '" & psCtaCod & "'," & lsEstado
    'RECO FIN**********************************************
        'By Capi 25052008 se modifico y comento para que no verifique vencimiento de la carta fianza segun solicitud de administracion de creditos
        '& " AND DateDiff(day,dVenc,'" & Format(dfecha, "yyyy/mm/dd") & "')  >=1"
        ' By Capi Acta 035-2007 se quito el estado aprobado
    'Set coConex = New COMConecta.DCOMConecta
    coConex.AbreConexion
    Set RecuperaCartaFianzaRenovacion = coConex.CargaRecordSet(lsSQL)
    'coConex.cierraConexion
    Set coConex = Nothing
End Function

'ALPA 20090331******************
Public Function CantidadCartaFianza(ByVal pcCtaCod As String, ByVal pnTipoGaran As Integer, ByVal pnTipoCambio As Double) As Double
Dim sql As String
Dim Rs As New ADODB.Recordset
sql = "exec stp_sel_ObtenerMontoCartaFianza '" & pcCtaCod & "'," & pnTipoGaran & "," & pnTipoCambio
coConex.AbreConexion
Set Rs = coConex.CargaRecordSet(sql)
If Rs.EOF And Rs.BOF Then
   CantidadCartaFianza = 0
Else
   CantidadCartaFianza = Rs!nMontoGarantia
End If
Set Rs = Nothing
Set coConex = Nothing
End Function
'*******************************

'MADM 20120105
Public Function RecuperaCartaFianzaExtornoAprobacion(ByVal psCtaCod As String) As ADODB.Recordset
Dim lsSQL As String
'Dim oCon As COMConecta.DCOMConecta

    'MAVM 20100606 BAS II
    lsSQL = " exec stp_sel_DevuelveCFExtornoApro '" & psCtaCod & "' "
    'Set coConex = New COMConecta.DCOMConecta
    coConex.AbreConexion
    Set RecuperaCartaFianzaExtornoAprobacion = coConex.CargaRecordSet(lsSQL)
    'coConex.cierraConexion
    Set coConex = Nothing
End Function

'WIOR 20120409
Public Function RecuperaCartaFianzaDetalle(ByVal psCtaCod As String) As ADODB.Recordset
Dim lsSQL As String
Dim Rs As ADODB.Recordset
Dim lsEstado As String

lsSQL = "exec stp_sel_DetalleCartaFianza '" & psCtaCod & "'"
       
    coConex.AbreConexion
    Set RecuperaCartaFianzaDetalle = coConex.CargaRecordSet(lsSQL)
    Set coConex = Nothing
End Function
'FIN
'WIOR 20120619 ******************************************************
Public Function RecuperaCartaFianzaDarBaja(ByVal psCtaCod As String) As ADODB.Recordset
Dim lsSQL As String

'->*****LUCV20160919, Comentó segun ERS004-2016
'    lsSQL = " declare @cPersCodTitular as varchar(13) "
'    lsSQL = lsSQL & " set @cPersCodTitular  =(select cPersCod from ProductoPersona where cCtaCod ='" & psCtaCod & "' and nPrdPersRelac=20) "
'    lsSQL = lsSQL & "Select CCF.nCondicion nColocCondicion, P.cPersCod, P.cPersNombre, " _
'        & " CE.nMonto as nMontoApr, CE.dVenc as dVencApr, CE.dPrdEstado, " _
'        & " cAnalista = (Select cPersNombre  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ") , " _
'        & " cApoderado= (Select cPersNombre  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersApoderado & ") , " _
'        & " CF.cNumFuente, PP2.cPersCod as cPersAcreedor, P2.cPersNombre as cPersNomAcre, " _
'        & " CCF.dAsignacion,CCF.dVencimiento,CCF.nModalidad, CCF.cFinalidad,Prd.nPrdEstado,nRenovacion,   " _
'        & " cTpoCredCod, CO.cConsDescripcion," _
'        & " cAvalCod = (Select isnull(PP.cPersCod,'')  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAval & ") , "
'        lsSQL = lsSQL & " cAvalNombre = (Select case when PP.cPersCod =@cPersCodTitular  AND CCF.cConsorcio IS NOT NULL then isnull(CCF.cConsorcio,'') Else isnull(cPersNombre,'') end From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAval & ")  "
'        lsSQL = lsSQL & " From ColocCartaFianza CCF " _
'        & " INNER JOIN ProductoPersona PP ON CCF.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular _
'        & " INNER JOIN ProductoPersona PP2 ON CCF.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAcreedor _
'        & " INNER JOIN Producto Prd ON Prd.cCtaCod = CCF.cCtaCod " _
'        & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod " _
'        & " INNER JOIN Persona P2 ON PP2.cPersCod = P2.cPersCod " _
'        & " INNER JOIN ColocCFEstado CE ON CCF.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstVigNorm _
'        & " INNER JOIN ColocFteIngreso CF ON CCF.cCtaCod = CF.cCtaCod " _
'        & " INNER JOIN Colocaciones C on CCF.cCtaCod = C.cCtaCod LEFT JOIN Constante CO on C.cTpoCredCod = CO.nConsValor and nConsCod = 3034 and nConsValor <>3034 " _
'        & " WHERE CCF.cCtaCod = '" & psCtaCod & "' AND Prd.nPrdEstado IN  ( " & gColocEstVigNorm & ") order by CE.dPrdEstado "
'
'    coConex.AbreConexion
'    Set RecuperaCartaFianzaDarBaja = coConex.CargaRecordSet(lsSQL)
'    coConex.CierraConexion
'    Set coConex = Nothing
'<-*****Fin, LUCV20160919
lsSQL = " exec stp_sel_RecuperaCartaFianzaDarBaja '" & psCtaCod & "' "
coConex.AbreConexion
Set RecuperaCartaFianzaDarBaja = coConex.CargaRecordSet(lsSQL)
Set coConex = Nothing
End Function
'WIOR FIN ***********************************************************

'WIOR 20120619 ******************************************************
Public Function RecuperaCartaFianzaAModificar(ByVal psCtaCod As String) As ADODB.Recordset
Dim lsSQL As String

'->*****LUCV20160919, Comentó Según ERS004-2016
'    lsSQL = " declare @cPersCodTitular as varchar(13) "
'    lsSQL = lsSQL & " set @cPersCodTitular  =(select cPersCod from ProductoPersona where cCtaCod ='" & psCtaCod & "' and nPrdPersRelac=20) "
'    lsSQL = lsSQL & "Select CCF.nCondicion nColocCondicion, P.cPersCod, P.cPersNombre, " _
'        & " CE.nMonto as nMontoApr, CE.dVenc as dVencApr, CE.dPrdEstado, " _
'        & " cAnalista = (Select cPersNombre  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ") , " _
'        & " cApoderado= (Select cPersNombre  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersApoderado & ") , " _
'        & " CF.cNumFuente, PP2.cPersCod as cPersAcreedor, P2.cPersNombre as cPersNomAcre, " _
'        & " CCF.dAsignacion,CCF.dVencimiento,CCF.nModalidad, CCF.cFinalidad,Prd.nPrdEstado,nRenovacion,   " _
'        & " cTpoCredCod, CO.cConsDescripcion," _
'        & " cAvalCod = (Select isnull(PP.cPersCod,'')  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAval & ") , "
'        lsSQL = lsSQL & " cAvalNombre = (Select case when PP.cPersCod =@cPersCodTitular  AND CCF.cConsorcio IS NOT NULL then isnull(CCF.cConsorcio,'') Else isnull(cPersNombre,'') end From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAval & ")  "
'        lsSQL = lsSQL & " From ColocCartaFianza CCF " _
'        & " INNER JOIN ProductoPersona PP ON CCF.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular _
'        & " INNER JOIN ProductoPersona PP2 ON CCF.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAcreedor _
'        & " INNER JOIN Producto Prd ON Prd.cCtaCod = CCF.cCtaCod " _
'        & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod " _
'        & " INNER JOIN Persona P2 ON PP2.cPersCod = P2.cPersCod " _
'        & " INNER JOIN ColocCFEstado CE ON CCF.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstVigNorm _
'        & " INNER JOIN ColocFteIngreso CF ON CCF.cCtaCod = CF.cCtaCod " _
'        & " INNER JOIN Colocaciones C on CCF.cCtaCod = C.cCtaCod LEFT JOIN Constante CO on C.cTpoCredCod = CO.nConsValor and nConsCod = 3034 and nConsValor <>3034 " _
'        & " WHERE CCF.cCtaCod = '" & psCtaCod & "' AND Prd.nPrdEstado IN  ( " & gColocEstModificada & ")  order by CE.dPrdEstado "
'
'    coConex.AbreConexion
'    Set RecuperaCartaFianzaAModificar = coConex.CargaRecordSet(lsSQL)
'    coConex.CierraConexion
'    Set coConex = Nothing
'<-*****Fin, LUCV20160919
    lsSQL = " exec stp_sel_RecuperaCartaFianzaAModificar '" & psCtaCod & "' "
    coConex.AbreConexion
    Set RecuperaCartaFianzaAModificar = coConex.CargaRecordSet(lsSQL)
    Set coConex = Nothing
End Function
'WIOR FIN ***********************************************************
'WIOR 20121003 ******************************************************
Public Function ObtenerModalidadesCF(Optional ByVal psEstado As String = "1") As ADODB.Recordset
On Error GoTo ErrorModalidad
Dim lsSQL As String

lsSQL = "EXEC stp_sel_ModalidadCF  " & gColCFModalidad & ",'" & psEstado & "'"

coConex.AbreConexion
Set ObtenerModalidadesCF = coConex.CargaRecordSet(lsSQL)
coConex.CierraConexion
Set coConex = Nothing
Exit Function
ErrorModalidad:
    Err.Raise Err.Number, "Error en Mostrar las Modalidades", Err.Description
End Function
Public Sub GrabarModalidadesCF(ByVal pnInsertActualiza As Integer, ByVal psDescripcion As String, ByVal pbEstado As Integer, Optional ByVal pnCod As Integer = 0)
On Error GoTo ErrorModalidad
Dim lsSQL As String

    If pnInsertActualiza = 1 Then
        lsSQL = "exec stp_ins_RegistrarModalidadCF "
    ElseIf pnInsertActualiza = 2 Then
        lsSQL = "exec stp_upd_ActualizarModalidadCF "
    End If
    
    lsSQL = lsSQL & gColCFModalidad & ",'" & psDescripcion & "'," & pbEstado & IIf(pnCod = 0, "", "," & pnCod)

coConex.AbreConexion
coConex.Ejecutar (lsSQL)
coConex.CierraConexion
Set coConex = Nothing
Exit Sub
ErrorModalidad:
    Err.Raise Err.Number, "Error en Registrar o Actualizar las Modalidades", Err.Description
End Sub
Public Function ExiteModalidaENCF(ByVal pnCod As Integer) As Long
Dim lsSQL As String
Dim R As ADODB.Recordset

On Error GoTo ErrorModalidad

lsSQL = "SELECT COUNT(nModalidad) CANT FROM ColocCartaFianza WHERE nModalidad =" & pnCod
coConex.AbreConexion
Set R = coConex.CargaRecordSet(lsSQL)

If R.RecordCount > 0 Then
    If Not (R.EOF And R.BOF) Then
        ExiteModalidaENCF = CLng(R!CANT)
    Else
        ExiteModalidaENCF = 0
    End If
Else
    ExiteModalidaENCF = 0
End If

coConex.CierraConexion
Set coConex = Nothing

Exit Function
ErrorModalidad:
    Err.Raise Err.Number, "Error Validar Exitencia Modalidad", Err.Description
End Function

'WIOR FIN ***********************************************************
'WIOR 20130312 ******************************************************
Public Sub InsActAutRenovacionCF(ByVal pnInsAct As Integer, ByVal psCtaCod As String, Optional ByVal pnMonto As Double = 0, _
                                 Optional ByVal pnComision As Double = 0, Optional ByVal pnPeriodo As Double = 0, _
                                 Optional ByVal dNEmision As Date = "01/01/1900", Optional ByVal dNVencimiento As Date = "01/01/1900", _
                                 Optional ByVal dAutorizado As Boolean = False, Optional ByVal pnRenovacion As Integer = 0, _
                                 Optional ByVal pcMovAut As String = "", Optional ByVal pcMovExt As String = "", _
                                 Optional ByVal pnMovNro As Double = -1, Optional ByVal psGlosa As String = "", _
                                 Optional ByVal pnEstado As Integer = -1)
                                 
On Error GoTo ErrorInsActAutRenovacionCF
Dim lsSQL As String

    If pnInsAct = 1 Then
        lsSQL = "Exec stp_ins_AutRenovacionCF '" & psCtaCod & "'," & pnMonto & "," & pnComision & "," & pnPeriodo & _
                ",'" & Format(dNEmision, "yyyymmdd") & "','" & Format(dNVencimiento, "yyyymmdd") & "'," & IIf(dAutorizado, 1, 0) & _
                "," & pnRenovacion & ",'" & pcMovAut & "','" & pcMovExt & "'," & pnMovNro & ",'" & psGlosa & "'," & pnEstado
                
    ElseIf pnInsAct = 2 Then
        lsSQL = "Exec stp_upd_AutRenovacionCF '" & psCtaCod & "','" & pcMovAut & "','" & pcMovExt & "'," & pnMovNro & ",'" & psGlosa & "'," & pnEstado
    End If
    
coConex.AbreConexion
coConex.Ejecutar (lsSQL)
coConex.CierraConexion
Set coConex = Nothing

Exit Sub
ErrorInsActAutRenovacionCF:
    Err.Raise Err.Number, "Error Insertar o Actualizar Autorizacion Renovacion CF", Err.Description
End Sub
Public Function ObternerAutRenovacion(Optional ByVal psCtaCod As String = "%", Optional ByVal psEstado As String = "0", Optional ByVal psRenovacion As String = "%") As ADODB.Recordset
On Error GoTo ErrorObternerAutRenovacion
Dim lsSQL As String

lsSQL = "EXEC stp_sel_ObternerAutRenovacion '" & psCtaCod & "','" & psEstado & "','" & psRenovacion & "'"

coConex.AbreConexion
Set ObternerAutRenovacion = coConex.CargaRecordSet(lsSQL)

coConex.CierraConexion
Set coConex = Nothing
Exit Function
ErrorObternerAutRenovacion:
    Err.Raise Err.Number, "Error en Obterner Autorizacion de Renovacion", Err.Description
End Function
Public Sub ExtornarPagoAutRenovaCF(ByVal psCtaCod As String, ByVal pnMovNro As Double, ByVal pnEstado As Integer)
                                 
On Error GoTo ErrorInsActAutRenovacionCF
Dim lsSQL As String
    
    lsSQL = "Exec stp_upd_ExtornaPagoAutRenovacionCF '" & psCtaCod & "'," & pnMovNro & "," & pnEstado
    
coConex.AbreConexion
coConex.Ejecutar (lsSQL)
coConex.CierraConexion
Set coConex = Nothing

Exit Sub
ErrorInsActAutRenovacionCF:
    Err.Raise Err.Number, "Error Insertar o Actualizar Autorizacion Renovacion CF", Err.Description
End Sub
Public Function ObternerAutRenovacionNro(ByVal psCtaCod As String, ByVal pnRenovacion As Integer, ByVal psEstado As String) As ADODB.Recordset
On Error GoTo ErrorObternerAutRenovacionNro
Dim lsSQL As String

lsSQL = "EXEC stp_sel_ObternerAutRenovacionNro '" & psCtaCod & "'," & pnRenovacion & ",'" & psEstado & "'"

coConex.AbreConexion
Set ObternerAutRenovacionNro = coConex.CargaRecordSet(lsSQL)

coConex.CierraConexion
Set coConex = Nothing
Exit Function
ErrorObternerAutRenovacionNro:
    Err.Raise Err.Number, "Error en Obtener Autorizacion de Renovacion Nro", Err.Description
End Function
Public Function ObternerRenovacionAntCF(ByVal psCtaCod As String, ByVal pnEstado As Integer) As ADODB.Recordset
On Error GoTo ErrorObternerRenovacionAntCF
Dim lsSQL As String

lsSQL = " SELECT TOP 1 dPrdEstado,nMotivoRechazo FROM coloccfestado WHERE CCTACOD ='" & psCtaCod & "' and nPrdEstado=" & pnEstado

coConex.AbreConexion
Set ObternerRenovacionAntCF = coConex.CargaRecordSet(lsSQL)

coConex.CierraConexion
Set coConex = Nothing
Exit Function
ErrorObternerRenovacionAntCF:
    Err.Raise Err.Number, "Error en Obtener Anterior Renovacion CF", Err.Description
End Function
Public Function OperacionesCFRestaura(ByVal pnTipo As Integer, ByVal psCtaCod As String, ByVal pnRenovacion As Integer, _
                                    Optional ByVal pnPeriodo As Integer = 0, Optional ByVal pnMonto As Double = 0, _
                                    Optional ByVal pdVigCol As Date = "01/01/1900", Optional ByVal pdEmision As Date = "01/01/1900", _
                                    Optional ByVal pdVencimiento As Date = "01/01/1900", Optional ByVal pnPrdEstado As Integer = 2020, _
                                    Optional ByVal pdPrdEstado As Date = "01/01/1900", Optional ByVal pnMotivoRechazo As Integer = 0) As ADODB.Recordset

On Error GoTo ErrorOperacionesCFRestaura
Dim lsSQL As String

'pnTipo=>1:Inserta,2:Selecciona,3:Elimina

Select Case pnTipo
    Case 1: lsSQL = "Exec stp_ins_RegistrarColocCFRestaura '" & psCtaCod & "'," & pnRenovacion & "," & pnPeriodo & "," & pnMonto & _
                    ",'" & Format(pdVigCol, "yyyymmdd") & "','" & Format(pdEmision, "yyyymmdd") & "','" & Format(pdVencimiento, "yyyymmdd") & "'," & _
                    pnPrdEstado & ",'" & Format(pdPrdEstado, "mm/dd/yyyy hh:mm:ss") & "'," & pnMotivoRechazo
    
    Case 2, 3: lsSQL = "Exec stp_sel_del_ColocCFRestaura " & pnTipo & ",'" & psCtaCod & "'," & pnRenovacion
End Select

coConex.AbreConexion

If pnTipo = 2 Then
    Set OperacionesCFRestaura = coConex.CargaRecordSet(lsSQL)
Else
    coConex.Ejecutar lsSQL
    Set OperacionesCFRestaura = Nothing
End If

coConex.CierraConexion
Set coConex = Nothing

Exit Function
ErrorOperacionesCFRestaura:
    Err.Raise Err.Number, "Error Operaciones CF Restauracion", Err.Description
End Function
Public Sub RegistroExtornoRenovacionCF(ByVal psCtaCod As String, ByVal pnRenovacion As Integer, ByVal pcMovExt As String, ByVal pcGlosa As String)
On Error GoTo ErrorRegistroExtornoRenovacionCF
Dim lsSQL As String
    
lsSQL = "Exec stp_ins_RegistrarExtornoRenovacion '" & psCtaCod & "'," & pnRenovacion & ",'" & pcMovExt & "','" & pcGlosa & "'"
    
coConex.AbreConexion
coConex.Ejecutar (lsSQL)
coConex.CierraConexion
Set coConex = Nothing

Exit Sub
ErrorRegistroExtornoRenovacionCF:
    Err.Raise Err.Number, "Error Registro Extorno Renovacion CF", Err.Description
End Sub
Public Function ObternerRenovacionesHechas() As ADODB.Recordset
On Error GoTo ErrorObternerAutRenovacion
Dim lsSQL As String

lsSQL = "EXEC stp_sel_ObternerRenovacionRealizada "

coConex.AbreConexion
Set ObternerRenovacionesHechas = coConex.CargaRecordSet(lsSQL)

coConex.CierraConexion
Set coConex = Nothing
Exit Function
ErrorObternerAutRenovacion:
    Err.Raise Err.Number, "Error en Obterner Autorizacion de Renovacion", Err.Description
End Function
'WIOR FIN ***********************************************************
'RECO20160310 ERS012-2016 *******************************************
Public Function ObtieneDatosAprobacionCF(ByVal psCtaCod As String) As ADODB.Recordset
    On Error GoTo ErrorObtieneDatosAprobacionCF
    Dim lsSQL As String

    lsSQL = "stp_sel_ObtieneDatosAprobacionCF '" & psCtaCod & "'"

    coConex.AbreConexion
    Set ObtieneDatosAprobacionCF = coConex.CargaRecordSet(lsSQL)

    coConex.CierraConexion
    Set coConex = Nothing
    Exit Function
ErrorObtieneDatosAprobacionCF:
    Err.Raise Err.Number, "Error en Obterner datos de aprobacion de carta fianza", Err.Description
End Function
Public Function ObtieneTasaCompCF(ByVal psCtaCod As String) As ADODB.Recordset
    On Error GoTo ErrorObtieneTasaCompCF
    Dim lsSQL As String

    lsSQL = "stp_sel_CF_ObtieneTasaComp '" & psCtaCod & "'"

    coConex.AbreConexion
    Set ObtieneTasaCompCF = coConex.CargaRecordSet(lsSQL)

    coConex.CierraConexion
    Set coConex = Nothing
    Exit Function
ErrorObtieneTasaCompCF:
    Err.Raise Err.Number, "Error en Obterner datos de la tasa de comisión de carta fianza", Err.Description
End Function
Public Function ObtieneTitAvaPagare(ByVal psCtaCod As String) As ADODB.Recordset
    On Error GoTo ErrorObtieneTitAvaPagare
    Dim lsSQL As String

    lsSQL = "stp_sel_CF_ObtieneTitAvaPagare '" & psCtaCod & "'"

    coConex.AbreConexion
    Set ObtieneTitAvaPagare = coConex.CargaRecordSet(lsSQL)

    coConex.CierraConexion
    Set coConex = Nothing
    Exit Function
ErrorObtieneTitAvaPagare:
    Err.Raise Err.Number, "Error en Obterner datos del titular pagare de carta fianza", Err.Description
End Function
'RECO FIN************************************************************

'<marg--------------
Public Function get_CredAdmControlDesembolso(ByVal psCtaCod As String) As ADODB.Recordset
Dim lsSQL As String
Dim Rs As ADODB.Recordset
Dim lsEstado As String

lsSQL = "exec get_CredAdmControlDesembolso '" & psCtaCod & "'"
       
    coConex.AbreConexion
    Set get_CredAdmControlDesembolso = coConex.CargaRecordSet(lsSQL)
    Set coConex = Nothing
End Function
Public Function get_ControlCreditosObsAdmCred(ByVal psCtaCod As String) As ADODB.Recordset
Dim lsSQL As String
Dim Rs As ADODB.Recordset
Dim lsEstado As String

lsSQL = "exec get_ControlCreditosObsAdmCred '" & psCtaCod & "'"
       
    coConex.AbreConexion
    Set get_ControlCreditosObsAdmCred = coConex.CargaRecordSet(lsSQL)
    Set coConex = Nothing
End Function
'</marg---------

'JOEP ERS047 20170904
Public Function get_VerificaCFAutoLiqxHipot(ByVal psCtaCod As String, ByVal pnTpProducto As String) As ADODB.Recordset
Dim lsSQL As String
Dim Rs As ADODB.Recordset
Dim lsEstado As String

lsSQL = "exec stp_sel_ERS047_VerificaSolicitudAutorizacionProducto '" & psCtaCod & "','" & pnTpProducto & "'"
       
    coConex.AbreConexion
    Set get_VerificaCFAutoLiqxHipot = coConex.CargaRecordSet(lsSQL)
    Set coConex = Nothing
End Function
'JOEP ERS047 20170904

'JOEP20180622 Acta 122-2018
Public Function get_ObtieneDatosConstantexCod(ByVal pnConsCod As Long) As ADODB.Recordset
Dim lsSQL As String
Dim Rs As ADODB.Recordset
Dim lsEstado As String

lsSQL = "exec stp_sel_obtieneConstanteXCod " & pnConsCod & ""
       
    coConex.AbreConexion
    Set get_ObtieneDatosConstantexCod = coConex.CargaRecordSet(lsSQL)
    Set coConex = Nothing
End Function
'JOEP20180622 Acta 122-2018

'JOEP20190124 CP
Public Function get_ValidadCobGar(ByVal pcCtaCod As String) As ADODB.Recordset
Dim lsSQL As String
Dim Rs As ADODB.Recordset
Dim lsEstado As String

lsSQL = "exec stp_sel_CatalogoCheckListValidaRegCob '" & pcCtaCod & "'"
       
    coConex.AbreConexion
    Set get_ValidadCobGar = coConex.CargaRecordSet(lsSQL)
    Set coConex = Nothing
End Function
Public Function get_ValidadCF(ByVal pcCtaCod As String, ByVal pnTpCred As Integer, ByVal pcCodCargo As String) As ADODB.Recordset
Dim lsSQL As String
Dim Rs As ADODB.Recordset
Dim lsEstado As String

lsSQL = "exec stp_sel_CatalogoCheckListValidacionesCF '" & pcCtaCod & "'," & pnTpCred & ",'" & pcCodCargo & "'"
       
    coConex.AbreConexion
    Set get_ValidadCF = coConex.CargaRecordSet(lsSQL)
    Set coConex = Nothing
End Function
'JOEP20190124 CP
