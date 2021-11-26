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
'ALPA 20080828****************************************************************
 'Se agrego el parametro lbCoop
 '*****************************************************************************
Public Function GetCuentasCaptacionesCmacs(ByVal lsOpeCod As String, Optional ByVal pnBitCentral As Boolean = False, Optional lbCoop As Boolean = False) As ADODB.Recordset
Dim vsServCentral As String
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConec As DConecta
    
    Set oConec = New DConecta
    
    sSql = "Select * From ConstSistema Where nConsSisCod = 43"
    oConec.AbreConexion
    Set R = oConec.CargaRecordSet(sSql)
    If R.RecordCount > 0 Then
        vsServCentral = Trim(R!nConsSisValor)
    Else
        vsServCentral = ""
    End If
    Set R = Nothing
    
    sSql = "Delete CtasCmacs"
    oConec.ConexionActiva.Execute sSql
    
    If pnBitCentral = True Then
        
        sSql = " INSERT INTO CtasCmacs "
        sSql = sSql & " Select T.cPersNombre cObjetoDesc, T.cCtaCod From ( "
        sSql = sSql & " Select P.cPersNombre, PC.cCtaCod From Persona P INNER JOIN " & vsServCentral & "productopersonaconsol PC "
        sSql = sSql & " INNER JOIN " & vsServCentral & "AhorroCConsol A ON PC.cCtaCod = A.cCtaCod ON P.cPersCod = PC.cPersCod "
        sSql = sSql & " Where A.nPersoneria IN (4,5,6,7) And A.nEstCtaAC NOT IN (1400, 1300) "
        'ALPA 20080828****************************************************************
        If lbCoop = True And lsOpeCod = "761202" Then
            sSql = sSql & "and PC.cCtaCod not in ('109022321000055069','109012321000467294','109022321000001864','109032321000002445','109042321000004502','109062321000031348','109072321000006629','109092321000010243') "
        End If
        '*****************************************************************************
        sSql = sSql & " And PC.nPrdPersRelac = 10 Union "
        sSql = sSql & " Select P.cPersNombre, PC.cCtaCod From Persona P INNER JOIN " & vsServCentral & "productopersonaconsol PC "
        sSql = sSql & " INNER JOIN " & vsServCentral & "PlazoFijoConsol A ON PC.cCtaCod = A.cCtaCod ON P.cPersCod = PC.cPersCod "
        sSql = sSql & " Where A.nPersoneria IN (4,5,6) And A.nEstCtaPF NOT IN (1400, 1300) "
        sSql = sSql & " And PC.nPrdPersRelac =  10 ) T "
        sSql = sSql & " Group by  T.cPersNombre, T.cCtaCod Order by T.cPersNombre,T.cCtaCod "
        
        oConec.Ejecutar sSql

        oConec.ConexionActiva.Execute sSql
    
        sSql = "Select cPersNombre From CtasCmacs Where substring(cCtaCod,9,1) = '" & Mid(lsOpeCod, 3, 1) & "' Group by cPersNombre Order by cPersNombre "

    Else
        sSql = "INSERT INTO CtasCmacs "
        sSql = sSql & " Select T.cNomPers cObjetoDesc, T.cCodCta From ( "
        sSql = sSql & " Select P.cNomPers, PC.cCodCta From DBPersona.dbo.Persona P INNER JOIN " & vsServCentral & "PersCuentaConsol PC INNER JOIN "
        sSql = sSql & vsServCentral & "AhorroCConsol A ON PC.cCodCta = A.cCodCta ON P.cCodPers = PC.cCodPers "
        sSql = sSql & " Where A.cPersoneria IN ('4','5','6')  And PC.cRelaCta = 'TI' "
        sSql = sSql & " Union "
        sSql = sSql & " Select P.cNomPers, PC.cCodCta From DBPersona.dbo.Persona P INNER JOIN " & vsServCentral & "PersCuentaConsol PC INNER JOIN "
        sSql = sSql & vsServCentral & "PlazoFijoConsol A ON PC.cCodCta = A.cCodCta ON P.cCodPers = PC.cCodPers "
        sSql = sSql & " Where A.cPersoneria IN ('4','5','6') And PC.cRelaCta = 'TI' "
        sSql = sSql & "  ) T Group by  T.cNomPers, T.cCodCta Order by T.cNomPers,T.cCodCta "
        
        
        oConec.ConexionActiva.Execute sSql
    
        sSql = "Select cNomPers From CtasCmacs Where substring(ccodcta,6,1) = '" & Mid(lsOpeCod, 3, 1) & "' Group by cNompers Order by cNomPers "
    End If
    Set GetCuentasCaptacionesCmacs = oConec.CargaRecordSet(sSql)
    
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function GetCtaIfFiltro(ByVal psCtaContCod As String, psCtaIFCod As String, psPersCod As String, Optional lbMuestraCta As Boolean = True) As String
    On Error GoTo GetCtaIfFiltroErr
    Dim sql As String
    Dim rs   As New ADODB.Recordset
    Dim oConect As DConecta
    
    Set oConect = New DConecta
    If oConect.AbreConexion = False Then Exit Function
    GetCtaIfFiltro = ""
    sql = "SELECT cCtaContCod ,  cCtaIFSubCta " _
        & " FROM " & vsServerComunes & "CtaIFFiltro WHERE cCtaContCod in ('" & psCtaContCod & "') " _
        & "         and cCtaIFCod = '" & psCtaIFCod & "' and cPersCod ='" & psPersCod & "'"
    
    Set rs = oConect.CargaRecordSet(sql)
    If Not rs.EOF Then
        GetCtaIfFiltro = IIf(lbMuestraCta, rs!cCtaContCod & rs!cCtaIFSubCta, rs!cCtaContCod)
        If GetCtaIfFiltro = "" And lbMuestraCta Then
            GetCtaIfFiltro = rs!cCtaContCod
        End If
    Else
        If GetCtaIfFiltro = "" And lbMuestraCta Then
            GetCtaIfFiltro = psCtaContCod
        End If
    End If
    rs.Close: Set rs = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Function
GetCtaIfFiltroErr:
    Call RaiseError(MyUnhandledError, "NContFunciones:GetCtaObjFiltro Method")
End Function

Public Function CargaCtaIFFiltro(ByVal psCtaCod As String, Optional psPersCod As String = "", Optional psIFTpo As String, Optional psCtaIFCod As String) As ADODB.Recordset
Dim lsFiltro As String
Dim sSql     As String
Dim rs   As New ADODB.Recordset
On Error GoTo CargaCtaIFFiltroErr

If psPersCod <> "" Then
    lsFiltro = " and cif.cPersCod = '" & psPersCod & "' "
End If
If psIFTpo <> "" Then
    lsFiltro = lsFiltro & " and cif.cIFTpo = '" & psIFTpo & "' "
End If
If psCtaIFCod <> "" Then
    lsFiltro = lsFiltro & " and cif.cCtaIFCod = '" & psCtaIFCod & "' "
End If
sSql = "SELECT cCtaContCod, 1 nCtaObjOrden, cif.cIFTpo + '.' + cif.cPersCod + CASE WHEN cif.cCtaIFCod = '' THEN '' ELSE '.' + cif.cCtaIFCod END cObjetoCod, RTRIM(P.cPersNombre) + ' ' + ISNULL(RTRIM(ci.cCtaIFDesc),'') cObjetoDesc, cCtaIFSubCta cSubCtaCod, convert(int,bUsoAgencia) bUsoAgencia " _
     & "FROM CtaIFFiltro cif LEFT JOIN CtaIF ci ON ci.cPersCod = cif.cPersCod and ci.cIFTpo = cif.cIFTpo and ci.cCtaIFCod = cif.cCtaIFCod " _
     & "        JOIN INSTITUCIONFINANC I ON I.cPersCod = cif.cPersCod and I.cIFTpo = cif.cIFTpo " _
     & "        JOIN PERSONA P ON P.CPERSCOD = I.CPERSCOD " _
     & "WHERE cif.cCtaContCod = '" & psCtaCod & "' " & lsFiltro & " ORDER BY cif.cIFTpo, cif.cPersCod, cif.cCtaIFCod "
    
    Dim oConect As DConecta
    Set oConect = New DConecta
    If oConect.AbreConexion = False Then Exit Function
    Set CargaCtaIFFiltro = oConect.CargaRecordSet(sSql)
    RSClose rs
    oConect.CierraConexion: Set oConect = Nothing
    Exit Function
CargaCtaIFFiltroErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub EliminaCtaIFFiltro(ByVal psCtaCod As String, Optional psPersCod As String = "", Optional psIFTpo As String, Optional psCtaIFCod As String)
Dim lsFiltro As String
Dim sSql     As String
On Error GoTo EliminaCtaIFFiltroErr

If psPersCod <> "" Then
    lsFiltro = " and cif.cPersCod = '" & psPersCod & "' "
End If
If psIFTpo <> "" Then
    lsFiltro = lsFiltro & " and cif.cIFTpo = '" & psIFTpo & "' "
End If
If psCtaIFCod <> "" Then
    lsFiltro = lsFiltro & " and cif.cCtaIFCod = '" & psCtaIFCod & "' "
End If
sSql = "DELETE CtaIFFiltro " _
     & "WHERE cCtaContCod = '" & psCtaCod & "' " & lsFiltro
    
    Dim oConect As New DConecta
    If oConect.AbreConexion = False Then Exit Sub
    oConect.Ejecutar sSql
    oConect.CierraConexion: Set oConect = Nothing
    Exit Sub
EliminaCtaIFFiltroErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub InsertaCtaIFFiltro(ByVal psCtaCod As String, psPersCod As String, psIFTpo As String, psCtaIFCod As String, psCtaIfSubCta As String, Optional pnUsoAgencia As Integer = 0)
Dim lsFiltro As String
Dim sSql     As String
On Error GoTo InsertaCtaIFFiltroErr

sSql = "INSERT CtaIFFiltro (cPersCod, cIFTpo, cCtaIfCod, cCtaContCod, cCtaIFSubCta, bUsoAgencia) " _
        & "VALUES ('" & psPersCod & "', '" & psIFTpo & "','" & psCtaIFCod & "','" & psCtaCod & "', '" & psCtaIfSubCta & "'," & pnUsoAgencia & " )"
    
    Dim oConect As New DConecta
    If oConect.AbreConexion = False Then Exit Sub
    oConect.Ejecutar sSql
    oConect.CierraConexion: Set oConect = Nothing
    Exit Sub
InsertaCtaIFFiltroErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'devengado.
'ALPA 20110703
'EJVG2013118 Se agregó psListaPersCod
Public Function CargaCtasIF(Optional ByVal pnMoneda As Moneda = 0, Optional psFiltroCtaIF As String = "", Optional pnMuestraIF As MuestraIF = MuestraCuentas, Optional ByVal psCanje As String = "", Optional psTipoIF As CGTipoIF, Optional pbFiltraActivos As Boolean = True, Optional ByVal bMostrarSoloVigentes As Boolean = False, Optional ByVal psListaPersCod As String = "") As ADODB.Recordset
    On Error GoTo CargaCtasIFErr
    Dim oConec As DConecta
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
    Dim lsFiltroPersona As String, lsFiltroPersona2 As String 'EJVG20131118
            
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
    'EJVG20131118 ***
    If Len(Trim(psListaPersCod)) <> 0 Then
        lsFiltroPersona = " AND I.cPersCod IN (SELECT VALOR FROM DBO.fnc_getTblValoresTexto( '" & psListaPersCod & "')) "
        lsFiltroPersona2 = " AND CI.cPersCod IN (SELECT VALOR FROM DBO.fnc_getTblValoresTexto( '" & psListaPersCod & "')) "
    End If
    'END EJVG *******
    
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion = False Then Exit Function
'ALPA 20110702 **************************
If bMostrarSoloVigentes = True Then
    sql = "select  CODIGO,CTAIFDESC,Nivel "
    sql = sql + "From "
    sql = sql + "   ( "
    sql = sql + "    SELECT C.*,case when LEN(C.CODIGO)<24 then 0 else 1 end nLogico,isnull(nSaldo,0) nSaldo "
    sql = sql + "    From "
    sql = sql + "    ( "
End If
'****************************************
sql = sql + " SELECT   CASE WHEN NIVEL =1 THEN CPERSCOD ELSE CPERSCOD + '.' + cCtaIFCod END AS CODIGO , " _
    & "          Convert(char(40),CTAIFDESC)  as CTAIFDESC, Nivel " _
    & " FROM ( SELECT  I.cIFTpo + '.' + CI.CPERSCOD as CPERSCOD, CI.cCtaIFCod, " _
    & "                 CONVERT(CHAR(40),case LEN(CI.cCtaIFCod) when 7 then '' else ISNULL( (SELECT LEFT(cDescripcion,22)  from ctaifadeudados cia join coloclineacredito cl on cl.cLineaCred = cia.cCodLinCred " _
    & "                                           WHERE cia.cPersCod = ci.cPersCod and cia.cIFTpo = ci.cIFTpo and cia.cCtaIFCod = ci.cCtaIFCod ) + ' ','') end + CI.cCtaIFDesc) AS CTAIFDESC, " _
    & "                 LEN(CI.cCtaIFCod) AS Nivel, I.cIFTpo , I.bCanje " _
    & "         FROM    INSTITUCIONFINANC I JOIN CTAIF CI ON CI.cPersCod = I.cPersCod AND I.cIFTpo= CI.cIFTpo " _
    & "         WHERE   SUBSTRING(CI.CCTAIFCOD,1,1) NOT IN('X') " & lsFiltroMoneda & lsIFFiltro & lsFiltroTipoIF & lsFiltroPersona _
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
    & "                         WHERE   ci.cctaIFEstado <> 3 and SUBSTRING(CI.CCTAIFCOD,1,1) NOT IN('X') " & lsFiltroMoneda & lsIFFiltro & lsFiltroPersona2 _
    & "                         ) AS C1 " _
    & "                 ON  C1.cIFTpo=I.cIFTpo AND C1.CPERSCOD= I.CPERSCOD " _
    & "         UNION Select  Replace(Str(nConsValor,2,0),' ','0') as cPerscod, '' as CtaIf , " _
    & "                 cConsDescripcion , 0 AS  NIVEL, Replace(Str(nConsValor,2,0),' ','0') as cIFTpo, 0 as bCanje " _
    & "         From    Constante " _
    & "         Where   nConsCod Like " & gCGTipoIF & " And nConsValor <> " & gCGTipoIF & " " _
    & "                 AND  Replace(Str(nConsValor,2,0),' ','0') IN (  SELECT  DISTINCT I.cIFTpo " _
    & "                                                                 FROM    INSTITUCIONFINANC I " _
    & "                                                                         JOIN CTAIF CI ON CI.cPersCod = I.cPersCod AND I.cIFTpo= CI.cIFTpo "
sql = sql + "                                                           WHERE   SUBSTRING(CI.CCTAIFCOD,1,1) NOT IN('X') " & lsFiltroMoneda & lsIFFiltro & lsFiltroPersona & " )" _
    & "  ) AS CTASIF " _
    & "  Where Nivel <= " & pnMuestraIF & lsFiltroCanje _
    '& "  ORDER BY CPERSCOD, cCtaIFCod "
'ALPA 20110702 **************************
If bMostrarSoloVigentes = True Then
    sql = sql + " ) C left JOIN"
    sql = sql + "  ("
    sql = sql + "   SELECT  rtrim(ltrim(cIFTpo)) cIFTpo,"
    sql = sql + "           rtrim(ltrim(cPersCod)) cPersCod,"
    sql = sql + "           rtrim(ltrim(cCtaIFCod)) cCtaIFCod,"
    sql = sql + "           SUM(NCAPITAL) nSaldo"
    sql = sql + "   From CtaIFCalendario"
    sql = sql + "   WHERE bvigente=1 and cEstado='0'"
    sql = sql + "   group by cIFTpo,cPersCod,cCtaIFCod"
    sql = sql + "   Having Sum(nCapital) > 0"
    sql = sql + "   ) D on C.CODIGO=D.cIFTpo+'.'+D.cPersCod+'.'+D.cCtaIFCod"
    sql = sql + "   ) C"
    'sql = sql + "  Where nLogico = 0 Or nSaldo > 0"
    'EJVG 20111019 ******************
    'EJVG 20120127 Se agregó el código de la Financiera=02 y Adeudos =05
    sql = sql + "  Where((nLogico = 0 Or (nSaldo > 0  AND nLogico = 1) AND LEFT(CODIGO,2) NOT IN ('05','03','04','01','02'))"
    sql = sql + "  OR LEFT(CODIGO,2) IN ('05','03','04','01','02'))"
    sql = sql + "    ORDER BY CODIGO"
End If
'END ****************************************

    Set rs = oConec.CargaRecordSet(sql)
    Set CargaCtasIF = rs
    
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
CargaCtasIFErr:
    Call RaiseError(MyUnhandledError, "DCajaCtasIF:CargaCtasIF Function")
End Function

'##ModelId=3A8AF4BD02A3
Public Function NombreIF(ByVal psPersCodIF As String) As String
    On Error GoTo NombreIFErr
    Dim sql As String
    Dim oConect As DConecta
    Set oConect = New DConecta
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
    Call RaiseError(MyUnhandledError, "DCajaCtasIF:EliminaInteres Method")
End Function

Public Function SubCuentaIF(ByVal psPersCodIF As String) As String
    On Error GoTo SubCuentaIFErr
    Dim sql As String
    Dim oConect As DConecta
    Set oConect = New DConecta
    SubCuentaIF = ""
    If oConect.AbreConexion = False Then Exit Function
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = "Select cSubCtaContCod from InstitucionFinanc where cPersCod='" & psPersCodIF & "'"
    Set rs = oConect.CargaRecordSet(sql)
    If Not rs.EOF And Not rs.BOF Then
        SubCuentaIF = Trim(rs!cSubCtaContCod)
    End If
    rs.Close
    Set rs = Nothing
        
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Function
SubCuentaIFErr:
    Call RaiseError(MyUnhandledError, "DCajaCtasIF:EliminaInteres Method")
End Function

Public Function GetCodPersAuxIF(ByVal psObjetoCod As String) As String
    On Error GoTo GetCodPersAuxIFErr
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim oConect As New DConecta
    GetCodPersAuxIF = ""
    If oConect.AbreConexion = False Then Exit Function
    
    sql = "Select cCodPers from AuxIF where cObjetoCod ='" & psObjetoCod & "'"
    Set rs = oConect.CargaRecordSet(sql)
    If Not rs.EOF And Not rs.BOF Then
        GetCodPersAuxIF = Trim(rs!cCodPers)
    End If
    
    RSClose rs
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Function
GetCodPersAuxIFErr:
    Call RaiseError(MyUnhandledError, "DCajaCtasIF:GetCodPersAuxIF Method")
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As ClasIni
Set oIni = New ClasIni

vsServerAdmin = oIni.BaseAdministracion
vsServerComunes = oIni.BaseComunes
vsServerNegocio = oIni.BaseNegocio
vsServerPersonas = oIni.BasePersonas
vsServerImagenes = oIni.BaseImagenes

End Sub

'ccordova 21-03
'Modificado ccordova 03-05
'******************************************************
Public Function GetLineaCredito() As ADODB.Recordset
Dim sSql As String
Dim oCon As DConecta

On Error GoTo GetLineaCreditoErr
    
    'sSQL = "SELECT cLineaCred, RTRIM(p.cPersNombre) + ' - ' + cDesCripcion cDescripcion, len(cLineaCred) nObjNivel, c.cPersCod " _
         & "FROM ColocLineaCredito c JOIN Persona p ON p.cPersCod = c.cPersCod " _
        & "WHERE bEstado = 1 and NOT cLineaCred = ''  " _
        & "ORDER BY cLineaCred "
    '28-04
    'sSQL = "SELECT DISTINCT cLineaCred, RTRIM(p.cPersNombre) + ' - ' + cDesCripcion cDescripcion, len(cLineaCred) nObjNivel, c.cPersCod " _
         & "FROM ColocLineaCredito c JOIN Persona p ON p.cPersCod = c.cPersCod " _
        & "WHERE bEstado = 1 and NOT cLineaCred = '' AND LEN(cLineaCred)<=13 " _
        & "ORDER BY cLineaCred "
    
    sSql = "SELECT DISTINCT cLineaCred ,RTRIM(p.cPersNombre) + ' - ' + cDesCripcion cDescripcion, len(cLineaCred) nObjNivel " & _
           " FROM ColocLineaCredito c JOIN Persona p ON p.cPersCod = c.cPersCod WHERE bEstado = 1 and NOT cLineaCred = '' AND LEN(cLineaCred)=2 " & _
           " Union " & _
           "SELECT DISTINCT cLineaCred ,RTRIM(p.cPersNombre) + ' - ' + cDesCripcion cDescripcion,len(cLineaCred) nObjNivel " & _
           " FROM ColocLineaCredito c JOIN Persona p ON p.cPersCod = c.cPersCod WHERE bEstado = 1 and NOT cLineaCred = '' AND LEN(cLineaCred)=4 " & _
           " Union " & _
           " SELECT DISTINCT LEFT(cLineaCred,4)+SUBSTRING(cLineaCred,10,2), RTRIM(p.cPersNombre) + ' - ' + 'PAQUETE '+ SUBSTRING(cLineaCred,10,2) cDescripcion,len(cLineaCred) nObjNivel " & _
           " FROM ColocLineaCredito c JOIN Persona p ON p.cPersCod = c.cPersCod WHERE bEstado = 1 and NOT cLineaCred = '' AND LEN(cLineaCred)=13 " & _
           " GROUP BY LEFT(cLineaCred,4)+SUBSTRING(cLineaCred,10,2),p.cPersNombre,c.cDescripcion,c.cLineaCred "

           
    Set oCon = New DConecta
    oCon.AbreConexion
    Set GetLineaCredito = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
    
GetLineaCreditoErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'Se agrego para manejar la tabla intermedia entre Creditos y Adeudados
Public Function GetCredSaldosAdeudo() As ADODB.Recordset
Dim sSql As String
Dim oCon As DConecta

On Error GoTo GetCredSaldosAdeudo
    'sSQL = "SELECT DISTINCT " & _
       " CodLinea = CSA.cCodPaq, C.cDescripcion , nSaldoPaq " & _
       " FROM CredSaldosAdeudo CSA INNER JOIN CtaIFAdeudados I " & _
       " ON I.cCtaIFCod=RIGHT(CSA.cCodPaq,7) INNER JOIN ColocLineaCredito C " & _
       " ON C.cLineaCred=LEFT(CSA.cCodPaq,5)"
    'sSQL = "SELECT DISTINCT CodLinea=CSA.cCodPaq,LEFT(C.cLineaCred,5),C.cDescripcion,nSaldoPaq FROM CredSaldosAdeudo CSA INNER JOIN ColocLineaCredito C " & _
            " ON LEFT(C.cLineaCred,5)=LEFT(CSA.cCodPaq,5) AND C.cLineaCred=((SELECT TOP 1 cLineaCred FROM CredSaldosAdeudoLinea WHERE LEFT(C.cLineaCred,5)=LEFT(cLineaCred,5))"
            
    'sSQL = "SELECT DISTINCT CodLinea=CSA.cCodPaq,LEFT(C.cLineaCred,5),C.cDescripcion,nSaldoPaq FROM CredSaldosAdeudo CSA INNER JOIN ColocLineaCredito C"
    'sSQL = sSQL & " ON LEFT(C.cLineaCred,5)=LEFT(CSA.cCodPaq,5) AND "
    'sSQL = sSQL & " C.cLineaCred=(SELECT TOP 1 cLineaCred FROM CredSaldosAdeudoLinea WHERE LEFT(C.cLineaCred,5)=LEFT(cLineaCred,5)) "
    'By Capi 22112007 comentado porque no es necesario
'    sSQL = "SELECT DISTINCT CodLinea=CSA.cCodPaq,LEFT(CL.cLineaCred,4)+RIGHT(CL.cLineaCred,2)," & _
'           " cDescripcion=(SELECT cDescripcion FROM ColocLineaCredito WHERE cLineaCred=LEFT(CL.cLineaCred,4) )+ ' PAQUETE ' + RIGHT(CL.cLineaCred,2), " & _
'           " nSaldoPaq= ISNULL(nSaldoPaqSol,0) + ISNULL(nSaldoPaqDol,0) * (SELECT nValFijo FROM TipoCambio WHERE dFecCamb=(SELECT MAX(dFecCamb)FROM TipoCambio))" & _
'           " FROM CredSaldosAdeudo CSA INNER JOIN CredSaldosAdeudoLinea CL ON CSA.cCodPaq=CL.cCodPaq " & _
'           " AND CL.cLineaCred=(SELECT TOP 1 cLineaCred FROM CredSaldosAdeudoLinea WHERE cCodPaq=CSA.cCodPaq)"
'    'End By
    'By Capi 22112007
    sSql = " SELECT CodLinea=cCodPaq,cDescripcion=cLinea,Documento=cdocumento, " & _
           " Fecha=dFecha,nSaldoPaq= nSaldo,Justifica=cJustifica, " & _
           " Moneda=cMoneda FROM CredSaldosAdeudo CSA "
           
    'End By
    '(SELECT nValFijo FROM TipoCambio WHERE dFecCamb=(SELECT MAX(dFecCamb)FROM TipoCambio)), " & _

    Set oCon = New DConecta
    oCon.AbreConexion
    Set GetCredSaldosAdeudo = oCon.CargaRecordSet(sSql)
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
                                                                                                                        
Dim sSql As String
Dim oCon As DConecta
'Dim sCorrelativo As String
Dim rs As ADODB.Recordset
    
    Set oCon = New DConecta
    oCon.AbreConexion
    
    'sSQL = "SELECT Valor=MAX(CONVERT(int,RIGHT(cCodPaq,5))) FROM CredSaldosAdeudo "
    'Set Rs = oCon.CargaRecordSet(sSQL)
    'sCorrelativo = IIf(IsNull(Rs!Valor), "00001", Format(Rs!Valor, "00000"))

    'sSQL = " INSERT INTO CredSaldosAdeudo(cCodPaq,nSaldoPaq,nSalColSol,nSalColDol) " & _
           " VALUES('" & pcCodPaq & sCorrelativo & "'," & pnSaldoPaq & ",0,0)"
    'By Capi 22112007
    
    sSql = " INSERT INTO CredSaldosAdeudo(cCodPaq,cLinea,cDocumento,dFecha,nSaldo,nSaldoPaqSol,nSaldoPaqDol,nSalColSol,nSalColDol,cJustifica,cMoneda) " & _
           " VALUES('" & pcCodPaq & "','" & pcLinea & "','" & pcDocum & "','" & Format(pdFecha, "mm/dd/yyyy") & "'," & pnSaldo & ",0,0,0,0,'" & pcJusti & "','" & pcMoneda & "')"
    oCon.ConexionActiva.Execute (sSql)
    
    'InsertaCredSaldosAdeudo = pcCodPaq & sCorrelativo
    'Insertamos la Tabla de Relacion
    '03-05
    'sSQL = "SELECT cCodPaq FROM CredSaldosAdeudoLinea WHERE cCodPaq='" & pcCodPaq & "' AND cLineaCred='" & pcLineaCred & "'"
    'Set Rs = oCon.CargaRecordSet(sSQL)
    'If Rs.EOF Then 'Solo si no existe
        'Guardamos en Soles y Dolares
        
    'By Capi 22112007 comentado porque ya no es necesario
'        sSQL = " INSERT INTO CredSaldosAdeudoLinea(cCodPaq,cLineaCred) " & _
'           " VALUES('" & pcCodPaq & "','" & Left(pcLineaCred, 4) & "1" & Right(pcLineaCred, 2) & "')"
'        oCon.ConexionActiva.Execute (sSQL)
'        sSQL = " INSERT INTO CredSaldosAdeudoLinea(cCodPaq,cLineaCred) " & _
'           " VALUES('" & pcCodPaq & "','" & Left(pcLineaCred, 4) & "2" & Right(pcLineaCred, 2) & "')"
'        oCon.ConexionActiva.Execute (sSQL)
    'End By
    
    oCon.CierraConexion
    Set oCon = Nothing

End Function

Public Function ModificaCredSaldosAdeudo(ByVal pcCodPaq As String, _
                                        ByVal pnSaldoPaqSol As Double, _
                                        ByVal pnSaldoPaqDol As Double)
Dim sSql As String
Dim oCon As DConecta

On Error GoTo ModificaCredSaldosAdeudo
    sSql = "UPDATE CredSaldosAdeudo " & _
           " SET nSaldoPaqSol= " & pnSaldoPaqSol & _
           "  ,nSaldoPaqDol= " & pnSaldoPaqDol & _
           " WHERE cCodPaq='" & pcCodPaq & "'"

    Set oCon = New DConecta
    oCon.AbreConexion
    oCon.ConexionActiva.Execute (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function

ModificaCredSaldosAdeudo:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function EliminaCredSaldosAdeudo(ByVal pcCodPaq As String)

Dim sSql As String
Dim oCon As DConecta

On Error GoTo EliminaCredSaldosAdeudo
    Call EliminaCredSaldosAdeudoDetalle(pcCodPaq)
    
    sSql = "DELETE FROM CredSaldosAdeudo " & _
           " WHERE cCodPaq='" & pcCodPaq & "'"

    Set oCon = New DConecta
    oCon.AbreConexion
    oCon.ConexionActiva.Execute (sSql)
    
    'Borrar las Lineas Asignadas
    sSql = "DELETE FROM CredSaldosAdeudoLinea WHERE cCodPaq='" & pcCodPaq & "'"
    oCon.ConexionActiva.Execute (sSql)
    
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function

EliminaCredSaldosAdeudo:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function ConsolidaCredSaldosAdeudo()
Dim sSql As String
Dim oCon As DConecta

On Error GoTo ConsolidaCredSaldosAdeudo
    sSql = "UPDATE CredSaldosAdeudo " & _
           " Set nSaldoPaq = T.Saldo " & _
           " FROM CredSaldosAdeudo A INNER JOIN " & _
           " ( SELECT C.cCodPaq,Saldo=SUM(nSaldoCap)  FROM CredSaldosAdeudo C INNER JOIN CtaIFAdeudados A ON RIGHT(C.cCodPaq,7) LIKE A.cCtaIFCod+'%' " & _
           " GROUP BY C.cCodPaq )     T " & _
           " ON T.cCodPaq =A.cCodPaq " & _
           " WHERE A.nPaqConsol=1"

    Set oCon = New DConecta
    oCon.AbreConexion
    oCon.ConexionActiva.Execute (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function

ConsolidaCredSaldosAdeudo:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


'******************************************************
Public Function GetSaldoCtaIFAdeudado(ByVal psPersCod As String, ByVal pnIFTpo As CGTipoIF, ByVal psCtaIFCod As String) As ADODB.Recordset
Dim sSql As String
Dim oCon As DConecta

On Error GoTo GetLineaCreditoErr
    
    sSql = " SELECT nSaldoCap FROM CtaIfAdeudados C WHERE " & _
            " C.cPersCod='" & psPersCod & "' and C.cIFTpo ='" & Format(pnIFTpo, "00") & "' and C.cCtaIFCod = '" & psCtaIFCod & "'"

        
    Set oCon = New DConecta
    oCon.AbreConexion
    Set GetSaldoCtaIFAdeudado = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
    
GetLineaCreditoErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetPersona_LineaCredito(ByVal psLineaCred As String) As String
Dim sSql As String
Dim oCon As DConecta
Dim rs As ADODB.Recordset
    
    sSql = " SELECT cPersCod FROM ColocLineaCredito WHERE cLineaCred='" & psLineaCred & "'"
        
    Set oCon = New DConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    GetPersona_LineaCredito = ""
    If Not rs.EOF Then GetPersona_LineaCredito = rs!cPerscod
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
    
End Function
 'By Capi Oct 07 Caja General
Public Function GetInformacionAdeudados(ByVal pnTipoCambio, ByVal pnAccion As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As DConecta
    Set oConecta = New DConecta
    oConecta.AbreConexion
    sSql = " Exec Stp_Sel_GeneraInformacionAdeudados " & pnTipoCambio & "," & pnAccion
    Set GetInformacionAdeudados = oConecta.Ejecutar(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
 End Function


Public Function GetCredSaldosAdeudoDetalle(ByVal psLineaCredito As String, _
                                            ByVal psCodPaq As String) As ADODB.Recordset
Dim sSql As String
Dim oCon As DConecta
    
     
'     sSQL = " SELECT A.cCtaIFCod,A.nSaldoCap,Confirmado=CASE WHEN S.cPersCod IS NULL THEN 'SI' ELSE 'NO'END, " & _
'            " Selec='1' FROM CtaIFAdeudados A INNER JOIN CredSaldosAdeudoDetalle C ON A.cCtaIFCod=C.cCtaIFCod AND C.cCodPaq='" & psCodPaq & "'" & _
'            " LEFT JOIN CtaIFSaldo S ON A.cCtaIFCod=S.cCtaIFCod AND A.cIFTpo=S.cIFTpo AND A.cPersCod = S.cPersCod AND S.nSaldo <0 " & _
'            " WHERE LEFT(A.cCodLinCred,4)='" & Left(psLineaCredito, 4) & "'" & _
'            " UNION "
'
'    sSQL = sSQL & " SELECT A.cCtaIFCod,A.nSaldoCap,Confirmado=CASE WHEN S.cPersCod IS NULL THEN 'SI' " & _
'                   " ELSE 'NO' END,  Selec= '0' FROM CtaIFAdeudados A LEFT JOIN CtaIFSaldo S ON A.cCtaIFCod = s.cCtaIFCod And A.cIFTpo = s.cIFTpo And A.cPersCod = s.cPersCod " & _
'                   " AND S.nSaldo <0  WHERE A.cCtaIFCod NOT IN( SELECT AD.cCtaIFcod FROM CredSaldosAdeudoDetalle " & _
'                   " AD INNER JOIN CredSaldosAdeudoLinea AL ON AL.cCodPaq=AD.cCodPaq AND LEFT(AL.cLineaCred,4)= LEFT(A.cCodLincred,4)) " & _
'                   " AND LEFT(A.cCodLinCred,4)='" & Left(psLineaCredito, 4) & "'"

sSql = " SELECT A.cCtaIFCod,A.nSaldoCap,Confirmado=CASE WHEN S.cPersCod IS NULL THEN 'SI' ELSE 'NO'END, " & _
            " Selec='1' , A.CMONEDAPAGO FROM CTAIF ci INNER JOIN CtaIFAdeudados A  ON ci.cPErscod =A.cPErscod and ci.cIFTPo=A.cIFTPo and ci.cCtaIFCod =A.cCtaIFCod INNER JOIN CredSaldosAdeudoDetalle C ON A.cCtaIFCod=C.cCtaIFCod AND C.cCodPaq='" & psCodPaq & "'" & _
            " LEFT JOIN CtaIFSaldo S ON A.cCtaIFCod=S.cCtaIFCod AND A.cIFTpo=S.cIFTpo AND A.cPersCod = S.cPersCod  " & _
            " WHERE ci.cCtaIFEstado =1 and  LEFT(A.cCodLinCred,4)='" & Left(psLineaCredito, 4) & "'" & _
            " UNION "
    
    sSql = sSql & " SELECT A.cCtaIFCod,A.nSaldoCap,Confirmado=CASE WHEN S.cPersCod IS NULL THEN 'SI' " & _
                   " ELSE 'NO' END,   Selec= '0', A.CMONEDAPAGO FROM CTAIF ci INNER JOIN CtaIFAdeudados A ON ci.cPErscod =A.cPErscod and ci.cIFTPo=A.cIFTPo and ci.cCtaIFCod =A.cCtaIFCod LEFT JOIN CtaIFSaldo S ON A.cCtaIFCod = s.cCtaIFCod And A.cIFTpo = s.cIFTpo And A.cPersCod = s.cPersCod " & _
                   " WHERE ci.cCtaIFEstado =1 and A.cCtaIFCod NOT IN( SELECT AD.cCtaIFcod FROM CredSaldosAdeudoDetalle " & _
                   " AD INNER JOIN CredSaldosAdeudoLinea AL ON AL.cCodPaq=AD.cCodPaq AND LEFT(AL.cLineaCred,4)= LEFT(A.cCodLincred,4)) " & _
                   " AND LEFT(A.cCodLinCred,4)='" & Left(psLineaCredito, 4) & "'"
    
    
    Set oCon = New DConecta
    oCon.AbreConexion
    Set GetCredSaldosAdeudoDetalle = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
End Function

Public Function InsertaCredSaldosAdeudoDetalle(ByVal pcCodPaq As String, _
                                        ByVal pcCtaIFCod As String, _
                                        Optional pcJustifica As String = "NO")
                                        
Dim sSql As String
Dim oCon As DConecta
Dim rs As ADODB.Recordset
    
    Set oCon = New DConecta
    oCon.AbreConexion
        
    sSql = " INSERT INTO CredSaldosAdeudoDetalle(cCodPaq,cCtaIFCod,cJustifica) " & _
           " VALUES('" & pcCodPaq & "','" & pcCtaIFCod & "','" & pcJustifica & "')"

    oCon.ConexionActiva.Execute (sSql)
       
    oCon.CierraConexion
    Set oCon = Nothing
    
End Function

Public Function EliminaCredSaldosAdeudoDetalle(ByVal pcCodPaq As String)

Dim sSql As String
Dim oCon As DConecta

    sSql = "DELETE FROM CredSaldosAdeudoDetalle " & _
           " WHERE cCodPaq='" & pcCodPaq & "'"

    Set oCon = New DConecta
    oCon.AbreConexion
    oCon.ConexionActiva.Execute (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
End Function

Public Function ObtenerCorrelativo_Paquete() As String
Dim sSql As String
Dim oCon As DConecta
Dim rs As ADODB.Recordset
    
    Set oCon = New DConecta
    oCon.AbreConexion
    sSql = "SELECT Valor=MAX(CONVERT(int,cCodPaq)) FROM CredSaldosAdeudo "
    Set rs = oCon.CargaRecordSet(sSql)
    ObtenerCorrelativo_Paquete = IIf(IsNull(rs!Valor), "00001", Format(rs!Valor + 1, "00000"))
    oCon.CierraConexion
    Set oCon = Nothing
    
End Function
'By capi 22112007
Public Function ObtenerLineas_Codigo() As ADODB.Recordset
Dim sSql As String
Dim oCon As DConecta
Dim rs As ADODB.Recordset
    
    Set oCon = New DConecta
    oCon.AbreConexion
    sSql = " Select rTrim(cLineaCred)+ '|' + rtrim(cDescripcion) From ColocLineaCredito Where Len(rTrim(cLineaCred))=4 Order by cLineaCred"
    Set ObtenerLineas_Codigo = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
End Function
'By Capi 22112007

Public Function ObtenerLineas_Descripcion(psCodigo) As String
Dim sSql As String
Dim oCon As DConecta
Dim rs As ADODB.Recordset
    
    Set oCon = New DConecta
    oCon.AbreConexion
    sSql = " Select cDescripcion Descripcion From ColocLineaCredito Where Len(rTrim(cLineaCred))= " & "'psCodigo'" & " Order by cLineaCred"
    Set rs = oCon.CargaRecordSet(sSql)
    ObtenerLineas_Descripcion = rs!Descripcion
    oCon.CierraConexion
    Set oCon = Nothing
    
End Function
'ALPA 20110823
Public Sub ActualizarCtaIFCalendario(ByVal pnTpoValor As Integer, ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psCtaIFCod As String, ByVal pnNroCuota As Integer)
Dim sSql As String
Dim oCon As DConecta
Dim rs As ADODB.Recordset
    
    Set oCon = New DConecta
    oCon.AbreConexion
    sSql = "exec stp_upd_CtaIFCalendarioSaldoMes " & pnTpoValor & ", '" & psPersCod & "', '" & psIFTpo & "', '" & psCtaIFCod & "', " & pnNroCuota
    oCon.CargaRecordSet (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
End Sub

