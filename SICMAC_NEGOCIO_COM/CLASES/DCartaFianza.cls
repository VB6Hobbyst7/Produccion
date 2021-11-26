VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCartaFianza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private gConsPersona As String
Private gConsComunes As String
Private gConsImagenes As String

Public Function ExisteTarifario(ByVal pnCod As Integer) As Boolean
Dim sql As String
Dim Co As DConecta
Dim rs As New ADODB.Recordset
Set Co = New DConecta
sql = "select cTarifCod from ColocCfTarifario where cTarifCod=" & pnCod
Co.AbreConexion
Set rs = Co.CargaRecordSet(sql)
Co.CierraConexion
If rs.EOF And rs.BOF Then
   ExisteTarifario = False
Else
   ExisteTarifario = True
End If
Set rs = Nothing
Set Co = Nothing
End Function
Public Sub InsertaTarifario(ByVal pnTasa, _
pnMonto As Double, ByVal pnCod As Integer, _
ByVal pnMoneda As Integer, ByVal pnModalidad As Integer)
Dim sql As String
Dim Co As DConecta
Set Co = New DConecta

sql = " Insert ColocCfTarifario"
sql = sql & " (cTarifCod,nModalidad,nMoneda,nTasaTrim,nMontoMinimo)"
sql = sql & " Values"
sql = sql & " (" & pnCod & "," & pnModalidad & "," & pnMoneda & ","
sql = sql & pnTasa & "," & pnMonto & ")"
Co.AbreConexion
Co.BeginTrans
Co.CommitTrans
Co.Ejecutar (sql)
Co.CierraConexion
Set Co = Nothing
End Sub
Public Sub ActualizaTarifario(ByVal pnTasa, _
pnMonto As Double, ByVal pnCod As Integer)
Dim Co As DConecta
Dim sql As String
Set Co = New DConecta
sql = " Update ColocCfTarifario"
sql = sql & " set nTasaTrim=" & pnTasa
sql = sql & " , nMontoMinimo=" & pnMonto
sql = sql & " where cTarifCod=" & pnCod
Co.AbreConexion
Co.BeginTrans
    Co.Ejecutar (sql)
Co.CommitTrans
Co.CierraConexion
Set Co = Nothing
End Sub

Private Sub Class_Initialize()
Dim CIni As COMConecta.DCOMClasIni
    Set CIni = New COMConecta.DCOMClasIni
    gConsPersona = CIni.BasePersonas
    gConsComunes = CIni.BaseComunes
    gConsImagenes = CIni.BaseImagenes
    Set CIni = Nothing
End Sub
Public Function RecuperaCF_Tarifario() As ADODB.Recordset
Dim Co As DConecta
Dim sql As String
Dim rs As New ADODB.Recordset
Set Co = New DConecta
sql = "select cTarifCod,"
sql = sql & " (select cConsDescripcion from Constante where nConsCod=3402 and nConsValor=nModalidad) Modalidad,"
sql = sql & " (Select cConsDescripcion from Constante where nConsCod=1011 and nConsValor=nMoneda) Moneda,"
sql = sql & " nTasaTrim, nMontoMinimo, nModalidad, nMoneda from ColocCfTarifario"
Co.AbreConexion
Set rs = Co.CargaRecordSet(sql)
Co.CierraConexion
Set RecuperaCF_Tarifario = rs
Set Co = Nothing
Set rs = Nothing
End Function


Public Function RecuperaCartaFianzaSolicitud(ByVal psctacod As String) As ADODB.Recordset
Dim lsSQL As String
Dim oCon As DConecta

    lsSQL = "Select CCF.nCondicion nColocCondicion, P.cPersCod, P.cPersNombre, " _
        & " CE.nMonto, CE.dVenc, CE.dPrdEstado, " _
        & " cAnalista = (Select cPersCod From ProductoPersona Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & "), " _
        & " CF.cNumFuente, PP2.cPersCod as cPersAcreedor, P2.cPersNombre as cPersNomAcre, " _
        & " CCF.dAsignacion,CCF.dVencimiento,CCF.nModalidad, CCF.cFinalidad   " _
        & " From ColocCartaFianza CCF " _
        & " INNER JOIN ProductoPersona PP ON CCF.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular _
        & " INNER JOIN ProductoPersona PP2 ON CCF.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAcreedor _
        & " INNER JOIN Producto Prd ON Prd.cCtaCod = CCF.cCtaCod " _
        & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod " _
        & " INNER JOIN Persona P2 ON PP2.cPersCod = P2.cPersCod " _
        & " INNER JOIN ColocCFEstado CE ON CCF.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstSolic _
        & " INNER JOIN ColocFteIngreso CF ON CCF.cCtaCod = CF.cCtaCod " _
        & " " _
        & "  " _
        & " WHERE CCF.cCtaCod = '" & psctacod & "' AND Prd.nPrdEstado = " & gColocEstSolic
    Set oCon = New DConecta
    oCon.AbreConexion
    Set RecuperaCartaFianzaSolicitud = oCon.CargaRecordSet(lsSQL)
    oCon.CierraConexion
    Set oCon = Nothing
End Function


Public Function RecuperaCartaFianzaSugerencia(ByVal psctacod As String) As ADODB.Recordset
Dim lsSQL As String
Dim oCon As DConecta
    
lsSQL = "Select CCF.nCondicion nColocCondicion, P.cPersCod, P.cPersNombre, " _
        & " CE.nMonto as nMontoSol, CE.dVenc as dVencSol, CE.dPrdEstado, " _
        & " cAnalista = (Select cPersNombre From ProductoPersona PP1 INNER JOIN Persona P1 ON PP1.cPersCod = P1.cPersCod " _
        & " Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ") , " _
        & " CF.cNumFuente, PP2.cPersCod as cPersAcreedor, P2.cPersNombre as cPersNomAcre, " _
        & " CCF.dAsignacion,CCF.dVencimiento,CCF.nModalidad, CCF.cFinalidad ,  " _
        & " nMontoSug = (Select IsNull(nMonto,CE.nMonto) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " ) , " _
        & " dVencSug = (Select IsNull(dVenc,CE.dVenc) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " )  " _
        & " From ColocCartaFianza CCF " _
        & " INNER JOIN ProductoPersona PP ON CCF.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular _
        & " INNER JOIN ProductoPersona PP2 ON CCF.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAcreedor _
        & " INNER JOIN Producto Prd ON Prd.cCtaCod = CCF.cCtaCod " _
        & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod " _
        & " INNER JOIN Persona P2 ON PP2.cPersCod = P2.cPersCod " _
        & " INNER JOIN ColocCFEstado CE ON CCF.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstSolic _
        & " INNER JOIN ColocFteIngreso CF ON CCF.cCtaCod = CF.cCtaCod " _
        & " WHERE CCF.cCtaCod = '" & psctacod & "' AND Prd.nPrdEstado in (" & gColocEstSolic & "," & gColocEstSug & ") "
    
    Set oCon = New DConecta
    oCon.AbreConexion
    Set RecuperaCartaFianzaSugerencia = oCon.CargaRecordSet(lsSQL)
    oCon.CierraConexion
    Set oCon = Nothing
End Function


Public Function RecuperaCartaFianzaAprobacion(ByVal psctacod As String) As ADODB.Recordset
Dim lsSQL As String
Dim oCon As DConecta

    lsSQL = "Select CCF.nCondicion nColocCondicion, P.cPersCod, P.cPersNombre, " _
        & " CE.nMonto as nMontoSug, CE.dVenc as dVencSug, CE.dPrdEstado, ciiu.cCIIUdescripcion, " _
        & " cAnalista = (Select cPersNombre  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ") , " _
        & " CF.cNumFuente, PFI.cRazSocDescrip, PP2.cPersCod as cPersAcreedor, P2.cPersNombre as cPersNomAcre, " _
        & " CCF.dAsignacion,CCF.dVencimiento,CCF.nModalidad, CCF.cFinalidad ,  " _
        & " nMontoSug = (Select IsNull(nMonto,CE.nMonto) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " ) , " _
        & " dVencSug = (Select IsNull(dVenc,CE.dVenc) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " )  " _
        & " From ColocCartaFianza CCF " _
        & " INNER JOIN ProductoPersona PP ON CCF.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular _
        & " INNER JOIN ProductoPersona PP2 ON CCF.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAcreedor _
        & " INNER JOIN Producto Prd ON Prd.cCtaCod = CCF.cCtaCod " _
        & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod " _
        & " INNER JOIN Persona P2 ON PP2.cPersCod = P2.cPersCod " _
        & " INNER JOIN ColocCFEstado CE ON CCF.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstSug _
        & " INNER JOIN ColocFteIngreso CF ON CCF.cCtaCod = CF.cCtaCod INNER JOIN PersFteIngreso PFI ON PFI.cNumFuente = CF.cNumFuente  LEFT JOIN Ciiu ciiu ON P.cPersCIIU = ciiu.cCIIUcod " _
        & " WHERE CCF.cCtaCod = '" & psctacod & "' AND Prd.nPrdEstado = " & gColocEstSug
    Set oCon = New DConecta
    oCon.AbreConexion
    Set RecuperaCartaFianzaAprobacion = oCon.CargaRecordSet(lsSQL)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function RecuperaCartaFianzaRechazo(ByVal psctacod As String) As ADODB.Recordset
Dim lsSQL As String
Dim oCon As DConecta

    lsSQL = "Select CCF.nCondicion nColocCondicion, P.cPersCod, P.cPersNombre, " _
        & " CE.nMonto as nMontoSol, CE.dVenc as dVencSol, CE.dPrdEstado, " _
        & " cAnalista = (Select cPersNombre  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ") , " _
        & " CF.cNumFuente, PP2.cPersCod as cPersAcreedor, P2.cPersNombre as cPersNomAcre, " _
        & " CCF.dAsignacion,CCF.dVencimiento,CCF.nModalidad, CCF.cFinalidad ,  " _
        & " nMontoSug = (Select IsNull(nMonto,CE.nMonto) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " ) , " _
        & " dVencSug = (Select IsNull(dVenc,CE.dVenc) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " )  " _
        & " From ColocCartaFianza CCF " _
        & " INNER JOIN ProductoPersona PP ON CCF.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular _
        & " INNER JOIN ProductoPersona PP2 ON CCF.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAcreedor _
        & " INNER JOIN Producto Prd ON Prd.cCtaCod = CCF.cCtaCod " _
        & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod " _
        & " INNER JOIN Persona P2 ON PP2.cPersCod = P2.cPersCod " _
        & " INNER JOIN ColocCFEstado CE ON CCF.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstSolic _
        & " INNER JOIN ColocFteIngreso CF ON CCF.cCtaCod = CF.cCtaCod " _
        & " WHERE CCF.cCtaCod = '" & psctacod & "' AND Prd.nPrdEstado in ( " & gColocEstSolic & "," & gColocEstSug & ") "
    Set oCon = New DConecta
    oCon.AbreConexion
    Set RecuperaCartaFianzaRechazo = oCon.CargaRecordSet(lsSQL)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
    
Public Function RecuperaCartaFianzaComision(ByVal psctacod As String) As ADODB.Recordset
Dim lsSQL As String
Dim oCon As DConecta

    lsSQL = "Select CCF.cCtaCod, CCF.nCondicion nColocCondicion, P.cPersCod, P.cPersNombre, Prd.nPrdEstado, " _
        & " CE.nMonto as nMontoApr, CE.dVenc as dVencApr, CE.dPrdEstado, " _
        & " cAnalista = (Select cPersNombre  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ") , " _
        & " cApoderado= (Select cPersCod From ProductoPersona Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersApoderado & ") , " _
        & " PP2.cPersCod as cPersAcreedor, P2.cPersNombre as cPersNomAcre, " _
        & " CCF.dAsignacion,CCF.dVencimiento,CCF.nModalidad, CCF.cFinalidad   " _
        & " " _
        & " From ColocCartaFianza CCF " _
        & " INNER JOIN ProductoPersona PP ON CCF.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular _
        & " INNER JOIN ProductoPersona PP2 ON CCF.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAcreedor _
        & " INNER JOIN Producto Prd ON Prd.cCtaCod = CCF.cCtaCod " _
        & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod " _
        & " INNER JOIN Persona P2 ON PP2.cPersCod = P2.cPersCod " _
        & " INNER JOIN ColocCFEstado CE ON CCF.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstAprob _
        & " WHERE CCF.cCtaCod = '" & psctacod & "' AND Prd.nPrdEstado in ( " & gColocEstAprob & "," & gColocEstVigNorm & " )"
    Set oCon = New DConecta
    oCon.AbreConexion
    Set RecuperaCartaFianzaComision = oCon.CargaRecordSet(lsSQL)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
        
Public Function RecuperaCartaFianzaEmision(ByVal psctacod As String) As ADODB.Recordset
Dim lsSQL As String
Dim oCon As DConecta

    lsSQL = "Select CCF.nCondicion nColocCondicion, P.cPersCod, P.cPersNombre, " _
        & " CE.nMonto as nMontoApr, CE.dVenc as dVencApr, CE.dPrdEstado, " _
        & " cAnalista = (Select cPersNombre  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ") , " _
        & " cApoderado= (Select cPersNombre  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersApoderado & ") , " _
        & " CF.cNumFuente, PP2.cPersCod as cPersAcreedor, P2.cPersNombre as cPersNomAcre, " _
        & " CCF.dAsignacion,CCF.dVencimiento,CCF.nModalidad, CCF.cFinalidad   " _
        & " " _
        & " " _
        & " From ColocCartaFianza CCF " _
        & " INNER JOIN ProductoPersona PP ON CCF.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular _
        & " INNER JOIN ProductoPersona PP2 ON CCF.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAcreedor _
        & " INNER JOIN Producto Prd ON Prd.cCtaCod = CCF.cCtaCod " _
        & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod " _
        & " INNER JOIN Persona P2 ON PP2.cPersCod = P2.cPersCod " _
        & " INNER JOIN ColocCFEstado CE ON CCF.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstAprob _
        & " INNER JOIN ColocFteIngreso CF ON CCF.cCtaCod = CF.cCtaCod " _
        & " WHERE CCF.cCtaCod = '" & psctacod & "' AND Prd.nPrdEstado = " & gColocEstAprob
        
    Set oCon = New DConecta
    oCon.AbreConexion
    Set RecuperaCartaFianzaEmision = oCon.CargaRecordSet(lsSQL)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
        
Public Function RecuperaCartaFianzaRetirar(ByVal psctacod As String) As ADODB.Recordset
Dim lsSQL As String
Dim oCon As DConecta

    lsSQL = "Select CCF.nCondicion nColocCondicion, P.cPersCod, P.cPersNombre, " _
        & " CE.nMonto as nMontoSol, CE.dVenc as dVencSol, CE.dPrdEstado, " _
        & " cAnalista = (Select cPersNombre  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ") , " _
        & " CF.cNumFuente, PP2.cPersCod as cPersAcreedor, P2.cPersNombre as cPersNomAcre, " _
        & " CCF.dAsignacion,CCF.dVencimiento,CCF.nModalidad, CCF.cFinalidad ,  " _
        & " nMontoSug = (Select IsNull(nMonto,CE.nMonto) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " ) , " _
        & " dVencSug = (Select IsNull(dVenc,CE.dVenc) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " )  " _
        & " From ColocCartaFianza CCF " _
        & " INNER JOIN ProductoPersona PP ON CCF.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular _
        & " INNER JOIN ProductoPersona PP2 ON CCF.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAcreedor _
        & " INNER JOIN Producto Prd ON Prd.cCtaCod = CCF.cCtaCod " _
        & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod " _
        & " INNER JOIN Persona P2 ON PP2.cPersCod = P2.cPersCod " _
        & " INNER JOIN ColocCFEstado CE ON CCF.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstSolic _
        & " INNER JOIN ColocFteIngreso CF ON CCF.cCtaCod = CF.cCtaCod " _
        & " WHERE CCF.cCtaCod = '" & psctacod & "' AND Prd.nPrdEstado in ( " & gColocEstAprob & ") "
    Set oCon = New DConecta
    oCon.AbreConexion
    Set RecuperaCartaFianzaRetirar = oCon.CargaRecordSet(lsSQL)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
        
Public Function RecuperaCartaFianzaDevolucion(ByVal psctacod As String) As ADODB.Recordset
Dim lsSQL As String
Dim oCon As DConecta

    lsSQL = "Select CCF.nCondicion nColocCondicion, P.cPersCod, P.cPersNombre, Prd.nPrdEstado, " _
        & " CE.nMonto as nMontoSol, CE.dVenc as dVencSol, CE.dPrdEstado, " _
        & " cAnalista = (Select cPersNombre  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ") ,  " _
        & " CF.cNumFuente, PP2.cPersCod as cPersAcreedor, P2.cPersNombre as cPersNomAcre, " _
        & " CCF.dAsignacion,CCF.dVencimiento,CCF.nModalidad, CCF.cFinalidad ,  " _
        & " nMontoSug = (Select IsNull(nMonto,CE.nMonto) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " ) , " _
        & " dVencSug = (Select IsNull(dVenc,CE.dVenc) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " )  " _
        & " From ColocCartaFianza CCF " _
        & " INNER JOIN ProductoPersona PP ON CCF.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular _
        & " INNER JOIN ProductoPersona PP2 ON CCF.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAcreedor _
        & " INNER JOIN Producto Prd ON Prd.cCtaCod = CCF.cCtaCod " _
        & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod " _
        & " INNER JOIN Persona P2 ON PP2.cPersCod = P2.cPersCod " _
        & " INNER JOIN ColocCFEstado CE ON CCF.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstSolic _
        & " INNER JOIN ColocFteIngreso CF ON CCF.cCtaCod = CF.cCtaCod " _
        & " WHERE CCF.cCtaCod = '" & psctacod & "' AND Prd.nPrdEstado in ( " & gColocEstVigNorm & ") "
    Set oCon = New DConecta
    oCon.AbreConexion
        Set RecuperaCartaFianzaDevolucion = oCon.CargaRecordSet(lsSQL)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function RecuperaDatosGarantiaCF(ByVal psctacod As String) As ADODB.Recordset
Dim oCon As DConecta
Dim lsSQL As String

On Error GoTo ErrorRecuperaDatosGarantiaCred
    lsSQL = "Select CN3.cConsDescripcion as cTipoCredDescrip, C.nModalidad as nColocDestino, CN.cConsDescripcion as cDestinoDescripcion , " _
    & " P.cPersCod, P.cPersNombre, CN2.cConsDescripcion as cMonedaDesc, " _
    & " DNI=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & gPersIdDNI & "'), " _
    & " CE.nMonto " _
    & " From ColocCartaFianza C INNER JOIN ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular _
    & "                    INNER JOIN " & gConsPersona & "Persona P ON PP.cPersCod = P.cPersCod " _
    & "                    INNER JOIN ColocCFEstado CE ON C.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = '" & gColocEstSolic & "' " _
    & "                    INNER JOIN Constante CN ON C.nModalidad = CN.nConsValor AND CN.nConsCod = " & gColCFModalidad _
    & "                    INNER JOIN Constante CN2 ON convert(int,substring(C.cCtaCod,9,1)) = CN2.nConsValor AND CN2.nConsCod = " & gMoneda _
    & "                    INNER JOIN Constante CN3 ON convert(int,substring(C.cCtaCod,6,3)) = CN3.nConsValor AND CN3.nConsCod = " & gProducto _
    & " WHERE C.cCtaCod = '" & psctacod & "' "
    
    Set oCon = New DConecta
    oCon.AbreConexion
    Set RecuperaDatosGarantiaCF = oCon.CargaRecordSet(lsSQL)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Exit Function

ErrorRecuperaDatosGarantiaCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
   
Public Function RecuperaCartaFianzaHonrar(ByVal psctacod As String) As ADODB.Recordset
Dim lsSQL As String
Dim oCon As DConecta

    lsSQL = "Select CCF.nCondicion nColocCondicion, P.cPersCod, P.cPersNombre, " _
        & " CE.nMonto as nMontoSol, CE.dVenc as dVencSol, CE.dPrdEstado, " _
        & " cAnalista = (Select cPersNombre  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ") , " _
        & " CF.cNumFuente, PP2.cPersCod as cPersAcreedor, P2.cPersNombre as cPersNomAcre, " _
        & " CCF.dAsignacion,CCF.dVencimiento,CCF.nModalidad, CCF.cFinalidad ,  " _
        & " nMontoSug = (Select IsNull(nMonto,CE.nMonto) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " ) , " _
        & " dVencSug = (Select IsNull(dVenc,CE.dVenc) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " )  " _
        & " From ColocCartaFianza CCF " _
        & " INNER JOIN ProductoPersona PP ON CCF.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular _
        & " INNER JOIN ProductoPersona PP2 ON CCF.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAcreedor _
        & " INNER JOIN Producto Prd ON Prd.cCtaCod = CCF.cCtaCod " _
        & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod " _
        & " INNER JOIN Persona P2 ON PP2.cPersCod = P2.cPersCod " _
        & " INNER JOIN ColocCFEstado CE ON CCF.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstSolic _
        & " INNER JOIN ColocFteIngreso CF ON CCF.cCtaCod = CF.cCtaCod " _
        & " WHERE CCF.cCtaCod = '" & psctacod & "' AND Prd.nPrdEstado in ( " & gColocEstVigNorm & ") "
        
    Set oCon = New DConecta
    oCon.AbreConexion
        Set RecuperaCartaFianzaHonrar = oCon.CargaRecordSet(lsSQL)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function VerificaComision(ByVal psCodCta As String) As Long
Dim lsSQL As String
Dim oCon As DConecta
Dim rs As Recordset

    lsSQL = "Select M.nMovNro from MOV M INNER JOIN MOVCOL MC ON M.nMovNro = MC.nMovNro " _
        & " Where M.cOpeCod = '" & gColCFOpeComisEfe & "' And nMovFlag = 0 AND cCtaCod = '" & psCodCta & "'"

    Set oCon = New DConecta
    oCon.AbreConexion
        Set rs = oCon.CargaRecordSet(lsSQL)
    If rs.EOF And rs.BOF Then
        VerificaComision = -1
    Else
        VerificaComision = rs!nMovNro
    End If
    
    Set rs = Nothing
    oCon.CierraConexion
    Set oCon = Nothing

End Function

Public Function RecuperaCartaFianzaHonrarCredito(ByVal psctacod As String) As ADODB.Recordset
Dim lsSQL As String
Dim oCon As DConecta

    lsSQL = "Select CCF.nCondicion nColocCondicion, P.cPersCod, P.cPersNombre, " _
        & " CE.nMonto as nMontoSol, CE.dVenc as dVencSol, CE.dPrdEstado, " _
        & " cAnalista = (Select cPersNombre  From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod Where cCtaCod = CCF.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ") , " _
        & " CF.cNumFuente, PP2.cPersCod as cPersAcreedor, P2.cPersNombre as cPersNomAcre, " _
        & " CCF.dAsignacion,CCF.dVencimiento,CCF.nModalidad, CCF.cFinalidad ,  " _
        & " nMontoSug = (Select IsNull(nMonto,CE.nMonto) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " ) , " _
        & " dVencSug = (Select IsNull(dVenc,CE.dVenc) From ColocCFEstado CE1 Where CE1.cCtaCod = CCF.cCtaCod and CE1.nPrdEstado = " & gColocEstSug & " )  " _
        & " From ColocCartaFianza CCF " _
        & " INNER JOIN ProductoPersona PP ON CCF.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular _
        & " INNER JOIN ProductoPersona PP2 ON CCF.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAcreedor _
        & " INNER JOIN Producto Prd ON Prd.cCtaCod = CCF.cCtaCod " _
        & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod " _
        & " INNER JOIN Persona P2 ON PP2.cPersCod = P2.cPersCod " _
        & " INNER JOIN ColocCFEstado CE ON CCF.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstSolic _
        & " INNER JOIN ColocFteIngreso CF ON CCF.cCtaCod = CF.cCtaCod " _
        & " WHERE CCF.cCtaCod = '" & psctacod & "' AND Prd.nPrdEstado in ( " & gColocEstHonrada & ") "
        
    Set oCon = New DConecta
    oCon.AbreConexion
        Set RecuperaCartaFianzaHonrarCredito = oCon.CargaRecordSet(lsSQL)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ValidadCreditoCF(ByVal psctacod As String) As Boolean
    Dim sSql As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim bSalida As Boolean
    
    sSql = "Select Count(*) as nCantidad"
    sSql = sSql & " From Producto"
    sSql = sSql & " Where SubString(cCtaCod,7,2)='21' and cCtaCod='" & psctacod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        If rs!nCantidad > 0 Then
            bSalida = True
        Else
            bSalida = False
        End If
    End If
    ValidadCreditoCF = bSalida
End Function

Public Function ValorCoberturaGarantia() As Double
    Dim sSql As String
    Dim oConec As DConecta
    Dim nValor As Double
    Dim rs As ADODB.Recordset
    
    sSql = "Select nParamValor"
    sSql = sSql & " From ColocParametro"
    sSql = sSql & " Where nParamVar=1028"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    If Not rs.BOF And Not rs.EOF Then
        nValor = rs!nParamValor
    End If
    Set rs = Nothing
    ValorCoberturaGarantia = nValor
End Function

