VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCajaCtasIF"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A80C40F0000"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Dim vsConexion As String
Dim vsServerAdmin As String
Dim vsServerComunes As String
Dim vsServerNegocio As String
Dim vsServerPersonas As String
Dim vsServerImagenes As Variant

Public Function GetCuentasCaptacionesCmacs(ByVal lsOpeCod As String, Optional ByVal pnBitCentral As Boolean = False) As ADODB.Recordset
Dim vsServCentral As String
Dim sSQL As String
Dim R As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta
    
    Set oConec = New COMConecta.DCOMConecta
    
    sSQL = "Select * From ConstSistema Where nConsSisCod = 43"
    oConec.AbreConexion
    Set R = oConec.CargaRecordSet(sSQL)
    If R.RecordCount > 0 Then
        vsServCentral = Trim(R!nConsSisValor)
    Else
        vsServCentral = ""
    End If
    Set R = Nothing
    
    sSQL = "Delete CtasCmacs"
    oConec.ConexionActiva.Execute sSQL
    
    If pnBitCentral = True Then
        
        sSQL = " INSERT INTO CtasCmacs "
        sSQL = sSQL & " Select T.cPersNombre cObjetoDesc, T.cCtaCod From ( "
        sSQL = sSQL & " Select P.cPersNombre, PC.cCtaCod From Persona P INNER JOIN " & vsServCentral & "productopersonaconsol PC "
        sSQL = sSQL & " INNER JOIN " & vsServCentral & "AhorroCConsol A ON PC.cCtaCod = A.cCtaCod ON P.cPersCod = PC.cPersCod "
        sSQL = sSQL & " Where A.nPersoneria IN (4,5,6,7) And A.nEstCtaAC NOT IN (1400, 1300) "
        sSQL = sSQL & " And PC.nPrdPersRelac = 10 Union "
        sSQL = sSQL & " Select P.cPersNombre, PC.cCtaCod From Persona P INNER JOIN " & vsServCentral & "productopersonaconsol PC "
        sSQL = sSQL & " INNER JOIN " & vsServCentral & "PlazoFijoConsol A ON PC.cCtaCod = A.cCtaCod ON P.cPersCod = PC.cPersCod "
        sSQL = sSQL & " Where A.nPersoneria IN (4,5,6) And A.nEstCtaPF NOT IN (1400, 1300) "
        sSQL = sSQL & " And PC.nPrdPersRelac =  10 ) T "
        sSQL = sSQL & " Group by  T.cPersNombre, T.cCtaCod Order by T.cPersNombre,T.cCtaCod "
        
        oConec.Ejecutar sSQL

        oConec.ConexionActiva.Execute sSQL
    
        sSQL = "Select cPersNombre From CtasCmacs Where substring(cCtaCod,9,1) = '" & Mid(lsOpeCod, 3, 1) & "' Group by cPersNombre Order by cPersNombre "

    Else
        sSQL = "INSERT INTO CtasCmacs "
        sSQL = sSQL & " Select T.cNomPers cObjetoDesc, T.cCodCta From ( "
        sSQL = sSQL & " Select P.cNomPers, PC.cCodCta From DBPersona.dbo.Persona P INNER JOIN " & vsServCentral & "PersCuentaConsol PC INNER JOIN "
        sSQL = sSQL & vsServCentral & "AhorroCConsol A ON PC.cCodCta = A.cCodCta ON P.cCodPers = PC.cCodPers "
        sSQL = sSQL & " Where A.cPersoneria IN ('4','5','6')  And PC.cRelaCta = 'TI' "
        sSQL = sSQL & " Union "
        sSQL = sSQL & " Select P.cNomPers, PC.cCodCta From DBPersona.dbo.Persona P INNER JOIN " & vsServCentral & "PersCuentaConsol PC INNER JOIN "
        sSQL = sSQL & vsServCentral & "PlazoFijoConsol A ON PC.cCodCta = A.cCodCta ON P.cCodPers = PC.cCodPers "
        sSQL = sSQL & " Where A.cPersoneria IN ('4','5','6') And PC.cRelaCta = 'TI' "
        sSQL = sSQL & "  ) T Group by  T.cNomPers, T.cCodCta Order by T.cNomPers,T.cCodCta "
        
        
        oConec.ConexionActiva.Execute sSQL
    
        sSQL = "Select cNomPers From CtasCmacs Where substring(ccodcta,6,1) = '" & Mid(lsOpeCod, 3, 1) & "' Group by cNompers Order by cNomPers "
    End If
    Set GetCuentasCaptacionesCmacs = oConec.CargaRecordSet(sSQL)
    
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function CargaCtaIFFiltro(ByVal psCtaCod As String, Optional psPersCod As String = "", Optional psIFTpo As String, Optional psCtaIfCod As String) As ADODB.Recordset
Dim lsFiltro As String
Dim sSQL     As String
Dim rs   As New ADODB.Recordset
On Error GoTo CargaCtaIFFiltroErr

If psPersCod <> "" Then
    lsFiltro = " and cif.cPersCod = '" & psPersCod & "' "
End If
If psIFTpo <> "" Then
    lsFiltro = lsFiltro & " and cif.cIFTpo = '" & psIFTpo & "' "
End If
If psCtaIfCod <> "" Then
    lsFiltro = lsFiltro & " and cif.cCtaIFCod = '" & psCtaIfCod & "' "
End If
sSQL = "SELECT cCtaContCod, 1 nCtaObjOrden, cif.cIFTpo + '.' + cif.cPersCod + CASE WHEN cif.cCtaIFCod = '' THEN '' ELSE '.' + cif.cCtaIFCod END cObjetoCod, RTRIM(P.cPersNombre) + ' ' + ISNULL(RTRIM(ci.cCtaIFDesc),'') cObjetoDesc, cCtaIFSubCta cSubCtaCod, convert(int,bUsoAgencia) bUsoAgencia " _
     & "FROM CtaIFFiltro cif LEFT JOIN CtaIF ci ON ci.cPersCod = cif.cPersCod and ci.cIFTpo = cif.cIFTpo and ci.cCtaIFCod = cif.cCtaIFCod " _
     & "        JOIN INSTITUCIONFINANC I ON I.cPersCod = cif.cPersCod and I.cIFTpo = cif.cIFTpo " _
     & "        JOIN PERSONA P ON P.CPERSCOD = I.CPERSCOD " _
     & "WHERE cif.cCtaContCod = '" & psCtaCod & "' " & lsFiltro & " ORDER BY cif.cIFTpo, cif.cPersCod, cif.cCtaIFCod "
    
    Dim oConect As COMConecta.DCOMConecta
    Set oConect = New COMConecta.DCOMConecta
    If oConect.AbreConexion = False Then Exit Function
    Set CargaCtaIFFiltro = oConect.CargaRecordSet(sSQL)
    RSClose rs
    oConect.CierraConexion: Set oConect = Nothing
    Exit Function
CargaCtaIFFiltroErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub EliminaCtaIFFiltro(ByVal psCtaCod As String, Optional psPersCod As String = "", Optional psIFTpo As String, Optional psCtaIfCod As String)
Dim lsFiltro As String
Dim sSQL     As String
On Error GoTo EliminaCtaIFFiltroErr

If psPersCod <> "" Then
    lsFiltro = " and cif.cPersCod = '" & psPersCod & "' "
End If
If psIFTpo <> "" Then
    lsFiltro = lsFiltro & " and cif.cIFTpo = '" & psIFTpo & "' "
End If
If psCtaIfCod <> "" Then
    lsFiltro = lsFiltro & " and cif.cCtaIFCod = '" & psCtaIfCod & "' "
End If
sSQL = "DELETE CtaIFFiltro " _
     & "WHERE cCtaContCod = '" & psCtaCod & "' " & lsFiltro
    
    Dim oConect As New COMConecta.DCOMConecta
    If oConect.AbreConexion = False Then Exit Sub
    oConect.Ejecutar sSQL
    oConect.CierraConexion: Set oConect = Nothing
    Exit Sub
EliminaCtaIFFiltroErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub InsertaCtaIFFiltro(ByVal psCtaCod As String, psPersCod As String, psIFTpo As String, psCtaIfCod As String, psCtaIfSubCta As String, Optional pnUsoAgencia As Integer = 0)
Dim lsFiltro As String
Dim sSQL     As String
On Error GoTo InsertaCtaIFFiltroErr

sSQL = "INSERT CtaIFFiltro (cPersCod, cIFTpo, cCtaIfCod, cCtaContCod, cCtaIFSubCta, bUsoAgencia) " _
        & "VALUES ('" & psPersCod & "', '" & psIFTpo & "','" & psCtaIfCod & "','" & psCtaCod & "', '" & psCtaIfSubCta & "'," & pnUsoAgencia & " )"
    
    Dim oConect As New COMConecta.DCOMConecta
    If oConect.AbreConexion = False Then Exit Sub
    oConect.Ejecutar sSQL
    oConect.CierraConexion: Set oConect = Nothing
    Exit Sub
InsertaCtaIFFiltroErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'devengado.
Public Function CargaCtasIF(Optional ByVal pnMoneda As Moneda = 0, Optional psFiltroCtaIF As String = "", Optional pnMuestraIF As MuestraIF = MuestraCuentas, Optional ByVal psCanje As String = "", Optional psTipoIF As CGTipoIF, Optional pbFiltraActivos As Boolean = True) As ADODB.Recordset
    On Error GoTo CargaCtasIFErr
    Dim oConec As COMConecta.DCOMConecta
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim lsIFFiltro As String
    Dim lsTpoCtaFiltro1 As String
    Dim lsTpoCtaFiltro As String
    Dim lsFiltroCanje  As String
    Dim lsFiltroCanje1 As String
    Dim lsFiltroTipoIF As String
    Dim Pos As String
    Dim lsCadAux As String
    Dim lsFiltroMoneda As String
        
    lsIFFiltro = ""
    If psFiltroCtaIF <> "" Then
        ' se filtran las cuentas segun tipo
        lsIFFiltro = " AND CI.cIFTpo+CI.cCtaIfCod LIKE '" & psFiltroCtaIF & "'"
        If pbFiltraActivos Then
            lsIFFiltro = lsIFFiltro & " AND CI.cCtaIFEstado in ('" & gEstadoCtaIFActiva & "','" & gEstadoCtaIFRegistrada & "') "
        End If
    End If
    lsFiltroCanje = ""
    If psCanje <> "" Then
        lsFiltroCanje = " And CTASIF.bCanje =" & psCanje & " "
    End If
    lsFiltroMoneda = ""
    If pnMoneda <> 0 Then
        lsFiltroMoneda = " AND SUBSTRING(CI.cCtaIfCod,3,1)='" & pnMoneda & "' "
    End If
    
    lsFiltroTipoIF = ""
    If psTipoIF <> 0 Then
      lsFiltroTipoIF = " AND I.cIFTpo = '" & Format(psTipoIF, "00") & "' "
    End If
    
    Set oConec = New COMConecta.DCOMConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion = False Then Exit Function
    
sql = " SELECT   CASE WHEN NIVEL =1 THEN CPERSCOD ELSE CPERSCOD + '.' + cCtaIFCod END AS CODIGO , " _
    & "          Convert(char(40),CTAIFDESC)  as CTAIFDESC, Nivel " _
    & " FROM ( SELECT  I.cIFTpo + '.' + CI.CPERSCOD as CPERSCOD, CI.cCtaIFCod, " _
    & "                 CONVERT(CHAR(40),ISNULL( (SELECT LEFT(cDescripcion,22)  from ctaifadeudados cia join coloclineacredito cl on cl.cLineaCred = cia.cCodLinCred " _
    & "                                           WHERE cia.cPersCod = ci.cPersCod and cia.cIFTpo = ci.cIFTpo and cia.cCtaIFCod = ci.cCtaIFCod ) + ' ','') + CI.cCtaIFDesc) AS CTAIFDESC, " _
    & "                 LEN(CI.cCtaIFCod) AS Nivel, I.cIFTpo , I.bCanje " _
    & "         FROM    INSTITUCIONFINANC I JOIN CTAIF CI ON CI.cPersCod = I.cPersCod AND I.cIFTpo= CI.cIFTpo " _
    & "         WHERE   SUBSTRING(CI.CCTAIFCOD,1,1) NOT IN('X') " & lsFiltroMoneda & lsIFFiltro & lsFiltroTipoIF _
    & "                 and ci.cctaIFEstado <> 3 and (LEN(ci.cCtaIFCod) > 3   or EXISTS( Select cIFTpo FROM CtaIF civ WHERE " & IIf(pbFiltraActivos, "civ.cCtaIFEstado = '" & gEstadoCtaIFActiva & "' and ", "") _
    & "                                       civ.cIFTpo = ci.cIFTpo and civ.cPersCod = ci.cPersCod and civ.cCtaIFCod like ci.cCtaIFCod + '_%' ) ) "
'and CI.cCtaIFEstado IN(0,1)
sql = sql & "   Union " _
    & "         SELECT  I.cIFTpo + '.' + I.CPERSCOD as CPERSCOD, '' AS CTAIF, P.CPERSNOMBRE , 1 AS NIVEL , " _
    & "                 I.cIFTpo , I.bCanje " _
    & "         FROM    INSTITUCIONFINANC I " _
    & "                 JOIN PERSONA P ON P.CPERSCOD = I.CPERSCOD " _
    & "                 JOIN (  SELECT  CI.cIFTpo, CI.CPERSCOD " _
    & "                         FROM    CTAIF CI " _
    & "                         WHERE   ci.cctaIFEstado <> 3 and SUBSTRING(CI.CCTAIFCOD,1,1) NOT IN('X') " & lsFiltroMoneda & lsIFFiltro _
    & "                         ) AS C1 " _
    & "                 ON  C1.cIFTpo=I.cIFTpo AND C1.CPERSCOD= I.CPERSCOD " _
    & "         UNION Select  Replace(Str(nConsValor,2,0),' ','0') as cPerscod, '' as CtaIf , " _
    & "                 cConsDescripcion , 0 AS  NIVEL, Replace(Str(nConsValor,2,0),' ','0') as cIFTpo, 0 as bCanje " _
    & "         From    Constante " _
    & "         Where   nConsCod Like " & gCGTipoIF & " And nConsValor <> " & gCGTipoIF & " " _
    & "                 AND  Replace(Str(nConsValor,2,0),' ','0') IN (  SELECT  DISTINCT I.cIFTpo " _
    & "                                                                 FROM    INSTITUCIONFINANC I " _
    & "                                                                         JOIN CTAIF CI ON CI.cPersCod = I.cPersCod AND I.cIFTpo= CI.cIFTpo "
sql = sql + "                                                           WHERE   SUBSTRING(CI.CCTAIFCOD,1,1) NOT IN('X') " & lsFiltroMoneda & lsIFFiltro & " )" _
    & "  ) AS CTASIF " _
    & "  Where Nivel <= " & pnMuestraIF & lsFiltroCanje _
    & "  ORDER BY CPERSCOD, cCtaIFCod "


    Set rs = oConec.CargaRecordSet(sql)
    Set CargaCtasIF = rs
    
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
CargaCtasIFErr:
    MsgBox Err.Description, vbCritical
End Function

'##ModelId=3A8AF4BD02A3
Public Function NombreIF(ByVal psPersCodIF As String) As String
    On Error GoTo NombreIFErr
    Dim sql As String
    Dim oConect As COMConecta.DCOMConecta
    Set oConect = New COMConecta.DCOMConecta
    If oConect.AbreConexion = False Then Exit Function
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    NombreIF = ""
    sql = "Select p.cPersNombre from InstitucionFinanc I JOIN  " & vsServerPersonas & "PERSONA P ON P.cPersCod=I.cPersCod  where I.cPersCod='" & psPersCodIF & "'"
    Set rs = oConect.CargaRecordSet(sql)
    If Not rs.EOF And Not rs.BOF Then
        NombreIF = Trim(rs!cpersNombre)
    End If
    rs.Close
    Set rs = Nothing
        
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Function
NombreIFErr:
    MsgBox Err.Description, vbCritical
End Function

Private Sub Class_Initialize()
    Dim oImp As COMDConstSistema.DCOMImpresoras
    Set oImp = New COMDConstSistema.DCOMImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As COMConecta.DCOMClasIni
Set oIni = New COMConecta.DCOMClasIni

vsServerAdmin = oIni.BaseAdministracion
vsServerComunes = oIni.BaseComunes
vsServerNegocio = oIni.BaseNegocio
vsServerPersonas = oIni.BasePersonas
vsServerImagenes = oIni.BaseImagenes

End Sub

Public Function GetLineaCredito() As ADODB.Recordset
Dim sSQL As String
Dim oCon As COMConecta.DCOMConecta

On Error GoTo GetLineaCreditoErr

sSQL = "SELECT DISTINCT cLineaCred ,RTRIM(p.cPersNombre) + ' - ' + cDesCripcion cDescripcion, len(cLineaCred) nObjNivel " & _
           " FROM ColocLineaCredito c JOIN Persona p ON p.cPersCod = c.cPersCod WHERE bEstado = 1 and NOT cLineaCred = '' AND LEN(cLineaCred)=2 " & _
           " Union " & _
           "SELECT DISTINCT cLineaCred ,RTRIM(p.cPersNombre) + ' - ' + cDesCripcion cDescripcion,len(cLineaCred) nObjNivel " & _
           " FROM ColocLineaCredito c JOIN Persona p ON p.cPersCod = c.cPersCod WHERE bEstado = 1 and NOT cLineaCred = '' AND LEN(cLineaCred)=4 " & _
           " Union " & _
           " SELECT DISTINCT LEFT(cLineaCred,4)+SUBSTRING(cLineaCred,10,2), RTRIM(p.cPersNombre) + ' - ' + 'PAQUETE '+ SUBSTRING(cLineaCred,10,2) cDescripcion,len(cLineaCred) nObjNivel " & _
           " FROM ColocLineaCredito c JOIN Persona p ON p.cPersCod = c.cPersCod WHERE bEstado = 1 and NOT cLineaCred = '' AND LEN(cLineaCred)=13 " & _
           " GROUP BY LEFT(cLineaCred,4)+SUBSTRING(cLineaCred,10,2),p.cPersNombre,c.cDescripcion,c.cLineaCred "

           
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set GetLineaCredito = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
    
GetLineaCreditoErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'Se agrego para manejar la tabla intermedia entre Creditos y Adeudados
Public Function GetCredSaldosAdeudo() As ADODB.Recordset
Dim sSQL As String
Dim oCon As COMConecta.DCOMConecta

On Error GoTo GetCredSaldosAdeudo

    sSQL = " SELECT CodLinea=cCodPaq,cDescripcion=cLinea,Documento=cdocumento, " & _
           " Fecha=dFecha,nSaldoPaq= nSaldo,Justifica=cJustifica, " & _
           " Moneda=cMoneda FROM CredSaldosAdeudo CSA "
           
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set GetCredSaldosAdeudo = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function

GetCredSaldosAdeudo:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function InsertaCredSaldosAdeudo(ByVal pcCodPaq As String, _
                                        ByVal pcLinea As String, _
                                        ByVal pcDocum As String, _
                                        ByVal pdFecha As Date, _
                                        ByVal pnSaldo As Double, _
                                        ByVal pcJusti As String, _
                                        ByVal pcMoneda As String)
                                                                                                                        
Dim sSQL As String
Dim oCon As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
     
    sSQL = " INSERT INTO CredSaldosAdeudo(cCodPaq,cLinea,cDocumento,dFecha,nSaldo,nSaldoPaqSol,nSaldoPaqDol,nSalColSol,nSalColDol,cJustifica,cMoneda) " & _
           " VALUES('" & pcCodPaq & "','" & pcLinea & "','" & pcDocum & "','" & Format(pdFecha, "mm/dd/yyyy") & "'," & pnSaldo & ",0,0,0,0,'" & pcJusti & "','" & pcMoneda & "')"
    oCon.ConexionActiva.Execute (sSQL)
    
    oCon.CierraConexion
    Set oCon = Nothing

End Function

Public Function ModificaCredSaldosAdeudo(ByVal pcCodPaq As String, _
                                        ByVal pnSaldoPaqSol As Double, _
                                        ByVal pnSaldoPaqDol As Double)
Dim sSQL As String
Dim oCon As COMConecta.DCOMConecta

On Error GoTo ModificaCredSaldosAdeudo
    sSQL = "UPDATE CredSaldosAdeudo " & _
           " SET nSaldoPaqSol= " & pnSaldoPaqSol & _
           "  ,nSaldoPaqDol= " & pnSaldoPaqDol & _
           " WHERE cCodPaq='" & pcCodPaq & "'"

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.ConexionActiva.Execute (sSQL)
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function

ModificaCredSaldosAdeudo:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function EliminaCredSaldosAdeudo(ByVal pcCodPaq As String)

Dim sSQL As String
Dim oCon As COMConecta.DCOMConecta

On Error GoTo EliminaCredSaldosAdeudo
    Call EliminaCredSaldosAdeudoDetalle(pcCodPaq)
    
    sSQL = "DELETE FROM CredSaldosAdeudo " & _
           " WHERE cCodPaq='" & pcCodPaq & "'"

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.ConexionActiva.Execute (sSQL)
    
    'Borrar las Lineas Asignadas
    sSQL = "DELETE FROM CredSaldosAdeudoLinea WHERE cCodPaq='" & pcCodPaq & "'"
    oCon.ConexionActiva.Execute (sSQL)
    
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function

EliminaCredSaldosAdeudo:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function ConsolidaCredSaldosAdeudo()
Dim sSQL As String
Dim oCon As COMConecta.DCOMConecta

On Error GoTo ConsolidaCredSaldosAdeudo
    sSQL = "UPDATE CredSaldosAdeudo " & _
           " Set nSaldoPaq = T.Saldo " & _
           " FROM CredSaldosAdeudo A INNER JOIN " & _
           " ( SELECT C.cCodPaq,Saldo=SUM(nSaldoCap)  FROM CredSaldosAdeudo C INNER JOIN CtaIFAdeudados A ON RIGHT(C.cCodPaq,7) LIKE A.cCtaIFCod+'%' " & _
           " GROUP BY C.cCodPaq )     T " & _
           " ON T.cCodPaq =A.cCodPaq " & _
           " WHERE A.nPaqConsol=1"

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.ConexionActiva.Execute (sSQL)
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function

ConsolidaCredSaldosAdeudo:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'******************************************************
Public Function GetSaldoCtaIFAdeudado(ByVal psPersCod As String, ByVal pnIFTpo As CGTipoIF, ByVal psCtaIfCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oCon As COMConecta.DCOMConecta

On Error GoTo GetLineaCreditoErr
    
    sSQL = " SELECT nSaldoCap FROM CtaIfAdeudados C WHERE " & _
            " C.cPersCod='" & psPersCod & "' and C.cIFTpo ='" & Format(pnIFTpo, "00") & "' and C.cCtaIFCod = '" & psCtaIfCod & "'"

        
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set GetSaldoCtaIFAdeudado = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
    
GetLineaCreditoErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetPersona_LineaCredito(ByVal psLineaCred As String) As String
Dim sSQL As String
Dim oCon As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
    
    sSQL = " SELECT cPersCod FROM ColocLineaCredito WHERE cLineaCred='" & psLineaCred & "'"
        
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    GetPersona_LineaCredito = ""
    If Not rs.EOF Then GetPersona_LineaCredito = rs!cPerscod
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
    
End Function
 'By Capi Oct 07 Caja General
Public Function GetInformacionAdeudados(ByVal pnTipoCambio, ByVal pnAccion As Integer) As ADODB.Recordset
    Dim sSQL As String
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSQL = " Exec Stp_Sel_GeneraInformacionAdeudados " & pnTipoCambio & "," & pnAccion
    Set GetInformacionAdeudados = oConecta.Ejecutar(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
 End Function


Public Function GetCredSaldosAdeudoDetalle(ByVal psLineaCredito As String, _
                                            ByVal psCodPaq As String) As ADODB.Recordset
Dim sSQL As String
Dim oCon As COMConecta.DCOMConecta
    
sSQL = " SELECT A.cCtaIFCod,A.nSaldoCap,Confirmado=CASE WHEN S.cPersCod IS NULL THEN 'SI' ELSE 'NO'END, " & _
            " Selec='1' , A.CMONEDAPAGO FROM CTAIF ci INNER JOIN CtaIFAdeudados A  ON ci.cPErscod =A.cPErscod and ci.cIFTPo=A.cIFTPo and ci.cCtaIFCod =A.cCtaIFCod INNER JOIN CredSaldosAdeudoDetalle C ON A.cCtaIFCod=C.cCtaIFCod AND C.cCodPaq='" & psCodPaq & "'" & _
            " LEFT JOIN CtaIFSaldo S ON A.cCtaIFCod=S.cCtaIFCod AND A.cIFTpo=S.cIFTpo AND A.cPersCod = S.cPersCod  " & _
            " WHERE ci.cCtaIFEstado =1 and  LEFT(A.cCodLinCred,4)='" & Left(psLineaCredito, 4) & "'" & _
            " UNION "
    
    sSQL = sSQL & " SELECT A.cCtaIFCod,A.nSaldoCap,Confirmado=CASE WHEN S.cPersCod IS NULL THEN 'SI' " & _
                   " ELSE 'NO' END,   Selec= '0', A.CMONEDAPAGO FROM CTAIF ci INNER JOIN CtaIFAdeudados A ON ci.cPErscod =A.cPErscod and ci.cIFTPo=A.cIFTPo and ci.cCtaIFCod =A.cCtaIFCod LEFT JOIN CtaIFSaldo S ON A.cCtaIFCod = s.cCtaIFCod And A.cIFTpo = s.cIFTpo And A.cPersCod = s.cPersCod " & _
                   " WHERE ci.cCtaIFEstado =1 and A.cCtaIFCod NOT IN( SELECT AD.cCtaIFcod FROM CredSaldosAdeudoDetalle " & _
                   " AD INNER JOIN CredSaldosAdeudoLinea AL ON AL.cCodPaq=AD.cCodPaq AND LEFT(AL.cLineaCred,4)= LEFT(A.cCodLincred,4)) " & _
                   " AND LEFT(A.cCodLinCred,4)='" & Left(psLineaCredito, 4) & "'"
    
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set GetCredSaldosAdeudoDetalle = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set oCon = Nothing
    
End Function

Public Function InsertaCredSaldosAdeudoDetalle(ByVal pcCodPaq As String, _
                                        ByVal pcCtaIFCod As String, _
                                        Optional pcJustifica As String = "NO")
                                        
Dim sSQL As String
Dim oCon As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
        
    sSQL = " INSERT INTO CredSaldosAdeudoDetalle(cCodPaq,cCtaIFCod,cJustifica) " & _
           " VALUES('" & pcCodPaq & "','" & pcCtaIFCod & "','" & pcJustifica & "')"

    oCon.ConexionActiva.Execute (sSQL)
       
    oCon.CierraConexion
    Set oCon = Nothing
    
End Function

Public Function EliminaCredSaldosAdeudoDetalle(ByVal pcCodPaq As String)

Dim sSQL As String
Dim oCon As COMConecta.DCOMConecta

    sSQL = "DELETE FROM CredSaldosAdeudoDetalle " & _
           " WHERE cCodPaq='" & pcCodPaq & "'"

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.ConexionActiva.Execute (sSQL)
    oCon.CierraConexion
    Set oCon = Nothing
    
End Function

Public Function ObtenerCorrelativo_Paquete() As String
Dim sSQL As String
Dim oCon As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    sSQL = "SELECT Valor=MAX(CONVERT(int,cCodPaq)) FROM CredSaldosAdeudo "
    Set rs = oCon.CargaRecordSet(sSQL)
    ObtenerCorrelativo_Paquete = IIf(IsNull(rs!Valor), "00001", Format(rs!Valor + 1, "00000"))
    oCon.CierraConexion
    Set oCon = Nothing
    
End Function
'By capi 22112007
Public Function ObtenerLineas_Codigo() As ADODB.Recordset
Dim sSQL As String
Dim oCon As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    sSQL = " Select rTrim(cLineaCred)+ '|' + rtrim(cDescripcion) From ColocLineaCredito Where Len(rTrim(cLineaCred))=4 Order by cLineaCred"
    Set ObtenerLineas_Codigo = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set oCon = Nothing
    
End Function
'By Capi 22112007

Public Function ObtenerLineas_Descripcion(psCodigo) As String
Dim sSQL As String
Dim oCon As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    sSQL = " Select cDescripcion Descripcion From ColocLineaCredito Where Len(rTrim(cLineaCred))= " & "'psCodigo'" & " Order by cLineaCred"
    Set rs = oCon.CargaRecordSet(sSQL)
    ObtenerLineas_Descripcion = rs!Descripcion
    oCon.CierraConexion
    Set oCon = Nothing
    
End Function

