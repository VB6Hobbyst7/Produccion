VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NContFunciones"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A836D4803B9"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim vsServerCom As String
Dim vsServerPers As String

Public Function GeneraDocNro(psDocTpo As TpoDoc, Optional DocSerie As String = "", Optional psMoneda As String = "", Optional psAgencia As String = "", Optional pbSegunFecha As Boolean = False) As String
Dim sSql As String, rs As New ADODB.Recordset
Dim oConect As DConecta
Set oConect = New DConecta

If psDocTpo > -1 Then
   If DocSerie <> "" Then
      sSql = "SELECT max(cDocNro) AS cDocCorrela FROM  movdoc md JOIN Mov m ON m.nMovNro = md.nMovNro " _
           & "WHERE nDocTpo = '" & psDocTpo & "' and substring(cDocNro,1," & Len(DocSerie) & ") = '" & DocSerie & "' " & IIf(psAgencia = "", "", " and SubString(m.cMovNro,18,2) = '" & Right(psAgencia, 2) & "' ")
   Else
      If psMoneda = "" Then
         If pbSegunFecha Then
            sSql = "select top 1 cDocNro as cDocCorrela from movdoc where ndoctpo = " & psDocTpo & " order by nmovnro desc "
         Else
            sSql = "SELECT max(cDocNro) AS cDocCorrela FROM  movdoc md JOIN Mov m ON m.nMovNro = md.nMovNro " _
                 & "WHERE nDocTpo = '" & psDocTpo & "' " & IIf(psAgencia = "", "", " and SubString(m.cMovNro,18,2) = '" & Right(psAgencia, 2) & "' ")
         End If
      Else
         sSql = "SELECT  cDocNro AS cDocCorrela , md.nMovnro " _
               & " FROM    movdoc md " _
               & " WHERE   nDocTpo = '" & psDocTpo & "' " _
               & "         and md.nMovNro = (  Select Max(MD1.nmovnro) " _
               & "                             From MovDoc MD1 JOIN Mov M on M.nMovNro= MD1.nMovnro " _
               & "                             WHERE MD1.nDocTpo = '" & psDocTpo & "' and Substring(M.cOpeCod,3,1) ='" & psMoneda & "'" & IIf(psAgencia = "", "", " and SubString(m.cMovNro,18,2) = '" & Right(psAgencia, 2) & "' ") & ") "
      End If
   End If
   oConect.AbreConexion
   Set rs = oConect.CargaRecordSet(sSql)
   If rs.EOF Then
      GeneraDocNro = ""
   Else
      If Not IsNull(rs!cDocCorrela) Then
         If DocSerie <> "" Then
            GeneraDocNro = DocSerie & "-" & Format(Val(Mid(rs!cDocCorrela, Len(DocSerie) + 2, 20)) + 1, String(8, "0"))
         Else
            If Mid(rs!cDocCorrela, 4, 1) = "-" Then
               GeneraDocNro = Mid(rs!cDocCorrela, 1, 3) & "-" & Format(Val(Mid(rs!cDocCorrela, 5, 20)) + 1, String(8, "0"))
            Else
               GeneraDocNro = Format(Val(rs!cDocCorrela) + 1, String(8, "0"))
            End If
         End If
      Else
         GeneraDocNro = ""
      End If
   End If
   If GeneraDocNro = "" Then
      GeneraDocNro = "00000001"
      If DocSerie <> "" Then
         GeneraDocNro = DocSerie & "-" & GeneraDocNro
      End If
   End If
    If psDocTpo = TpoDocOrdenPago Then
        'Dim oCaja As New nCajaGeneral
        'Para verificar las OP entregadas a Caja General
        'Cuando este implementado el Centralizado del NEGOCIO
        'GeneraDocNro = oCaja.VerificaOPRango(GeneraDocNro, psMoneda)
        'Set oCaja = Nothing
    End If
    oConect.CierraConexion
Else
   Err.Raise 50001, "NContFunciones: GeneraDocNro", "Especificar Tipo de Documento para Generar Número"
End If
End Function

'Public Function GeneraDocNro(psDocTpo As TpoDoc, psMoneda As Moneda, Optional psDocSerie As String = "") As String
'    On Error GoTo GeneraMovNroErr
'    Dim sSql As String
'    Dim rs As New ADODB.Recordset
'    Dim oConect As DConecta
'    Dim oCaja As nCajaGeneral
'
'    Set oCaja = New nCajaGeneral
'    Set oConect = New DConecta
'    If oConect.AbreConexion = False Then Exit Function
'
'        If psDocSerie <> "" Then
'            sSql = "SELECT max(cDocNro) AS cDocCorrela FROM Movdoc md JOIN Mov m ON m.nMovNro = md.nMovNro " _
'                & "WHERE   m.nMovFlag = ('" & gMovFlagVigente & "') and nDocTpo = '" & Format(psDocTpo, "00") & "' and substring(cDocNro,1," & Len(psDocSerie) & ") = '" & psDocSerie & "'"
'        Else
'            sSql = "SELECT  cDocNro AS cDocCorrela , md.nMovnro " _
'                & " FROM    MovDoc MD JOIN MOV M on M.nMovNro = md.nMovNro  " _
'                & " WHERE   nDocTpo = '" & Format(psDocTpo, "00") & "'  " _
'                & "         and md.nmovnro = (  Select Max(MD1.nmovnro)  " _
'                & "                             From MovDoc MD1 JOIN Mov M1 on M1.nMovNro= MD1.nMovnro " _
'                & "                             WHERE MD1.nDocTpo = '" & Format(psDocTpo, "00") & "' and Substring(M.cOpeCod,3,1) ='" & psMoneda & "'  )  "
'
'                'and  m.nMovFlag = ('" & gMovFlagVigente & "')
'                'and  m1.nMovFlag = ('" & gMovFlagVigente & "')
'        End If
'        Set rs = oConect.CargaRecordSet(sSql)
'        If Not rs.EOF And Not rs.BOF Then
'            If Not IsNull(rs!cDocCorrela) Then
'                If psDocSerie <> "" Then
'                   GeneraDocNro = psDocSerie & "-" & Format(Val(Mid(rs!cDocCorrela, Len(psDocSerie) + 2, 20)) + 1, String(8, "0"))
'                Else
'                   If Mid(rs!cDocCorrela, 4, 1) = "-" Then
'                      GeneraDocNro = Mid(rs!cDocCorrela, 1, 3) & "-" & Format(Val(Mid(rs!cDocCorrela, 5, 20)) + 1, String(8, "0"))
'                   Else
'                      GeneraDocNro = Format(Val(rs!cDocCorrela) + 1, String(8, "0"))
'                   End If
'                End If
'            Else
'                GeneraDocNro = "00000001"
'                If psDocSerie <> "" Then GeneraDocNro = psDocSerie & "-" & GeneraDocNro
'            End If
'        Else
'            GeneraDocNro = "00000001"
'            If psDocSerie <> "" Then GeneraDocNro = psDocSerie & "-" & GeneraDocNro
'        End If
'        rs.Close
'        Set rs = Nothing
'        oConect.CierraConexion
'        Set oConect = Nothing
'        If psDocTpo = TpoDocOrdenPago Then
'            GeneraDocNro = oCaja.VerificaOPRango(GeneraDocNro, psMoneda)
'        End If
'        Set oCaja = Nothing
'    Exit Function
'GeneraMovNroErr:
'    Call RaiseError(MyUnhandledError, "NContFunciones:GeneraMovNro Method")
'End Function

Public Function GetFiltroObjetos(ByVal pnTipoObj As TpoObjetos, ByVal psCtaContCod As String, ByVal psObjetoCod As String, Optional lbMuestraCta As Boolean = True, Optional lbExisteFiltro As Boolean = False) As String
    On Error GoTo GetCtaObjFiltroErr
    Dim sql As String
    Dim rs   As New ADODB.Recordset
    Dim oConect As DConecta
    Dim lsPersCtaIF As String
    Dim lsCtaIf As String
    Dim lsTpoIf As String
    
    Set oConect = New DConecta
    If oConect.AbreConexion = False Then Exit Function
    GetFiltroObjetos = ""
    Select Case pnTipoObj
        Case ObjCMACAgencias
            sql = "SELECT cAgeCod,  cSubCtaCod as SubCta  FROM " & vsServerCom & "Agencias where cAgeCod='" & psObjetoCod & "'"
        Case ObjCMACArea
            sql = "SELECT cAreaCod, cSubCtaCod  as SubCta  FROM " & vsServerCom & "AREAS where cAreaCod='" & psObjetoCod & "'"
        Case ObjCMACAgenciaArea
           sql = "     SELECT   AA.cAreaCod, AA.cAgeCod, " _
                & "             CASE " _
                & "                  WHEN CF.cSubctaCod IS NULL THEN AA.cSubCtaCod " _
                & "                  Else CF.cSubCtaCod END As SubCta " _
                & "    FROM     " & vsServerCom & "AREAAGENCIA AA " _
                & "             LEFT JOIN ( Select cAreaCod , cAgeCod , cSubCtaCod " _
                & "                          From " & vsServerCom & "CtaAreaAgeFiltro " _
                & "                          WHERE  cCtaContcod in ('" & psCtaContCod & "')) AS CF " _
                & "             ON CF.cAreacod = AA.cAreaCod and CF.cAgeCod = AA.cAgeCod " _
                & "    WHERE    (AA.cAgeCod = '" & Mid(psObjetoCod, 4, 2) & "' Or AA.cAgeCod = '') AND  AA.cAreaCod='" & Mid(psObjetoCod, 1, 3) & "' and Exists(Select 'X' FROM CtaObj WHERE cCtaContCod in ('" & psCtaContCod & "') and cObjetoCod = '" & pnTipoObj & "') "
                
        Case ObjEntidadesFinancieras
            'If Len(psObjetoCod) > 15 Then
                lsTpoIf = Mid(psObjetoCod, 1, 2)
                lsPersCtaIF = Mid(psObjetoCod, 4, 13)
                lsCtaIf = Mid(psObjetoCod, 18, 10)
            'Else
            '    If psPersCodIf <> "" Then Exit Function
            '    lsPersCtaIF = psPersCodIf
            '    lsCtaIf = psObjetoCod
            'End If
            If Len(psObjetoCod) = 13 Then
                sql = "select cPersCod, cSubCtaContcod SubCta FROM InstitucionFinanc where cPersCod = '" & psObjetoCod & "' "
            Else
                sql = "Select cPersCod, cCtaIfCod, cCtaContCod, cCtaIFSubCta as SubCta From " & vsServerCom & "CtaIFFiltro WHERE  cPersCod = '" & lsPersCtaIF & "' and cIFTpo='" & lsTpoIf & "' AND cCtaIfCod='" & lsCtaIf & "' AND cCtaContCod in ('" & psCtaContCod & "') "
                sql = sql & "UNION "
                sql = sql & "Select cPersCod, cCtaIfCod, cCtaContCod, cCtaIFSubCta as SubCta From " & vsServerCom & "CtaIFFiltro WHERE  cPersCod = '" & lsPersCtaIF & "' and cIFTpo='" & lsTpoIf & "' AND cCtaContCod in ('" & psCtaContCod & "') and cCtaIFCod = '' "
                
            End If
        Case Else
            sql = "SELECT cCtaObjSubCta as SubCta FROM " & vsServerCom & "CtaObjFiltro WHERE cCtaContCod in ('" & psCtaContCod & "') and cObjetoCod = '" & psObjetoCod & "'"
    End Select
    Set rs = oConect.CargaRecordSet(sql)
    If Not rs.EOF Then
        lbExisteFiltro = True
        GetFiltroObjetos = IIf(lbMuestraCta, psCtaContCod, "") & rs!SubCta
    End If
    If GetFiltroObjetos = "" Then
        GetFiltroObjetos = IIf(lbMuestraCta, psCtaContCod, "")
    End If
    Exit Function
GetCtaObjFiltroErr:
    Call RaiseError(MyUnhandledError, "NContFunciones:GetCtaObjFiltro Method")
End Function



Public Function ObtieneFechaMov(psMovNro As String, pbDia As Boolean) As String
   On Error GoTo ObtieneFechaMovErr
   If pbDia Then
      ObtieneFechaMov = Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Left(psMovNro, 4)
   Else
      ObtieneFechaMov = Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 7, 2) & "/" & Left(psMovNro, 4)
   End If
   Exit Function
ObtieneFechaMovErr:
    Call RaiseError(MyUnhandledError, "NContFunciones:ObtieneFechaMov Method")
End Function

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
Public Function GeneraMovNro(ByVal pdFecha As Date, Optional ByVal psCodAge As String = "07", Optional ByVal psUser As String = "SIST", Optional psMovNro As String = "") As String
    On Error GoTo GeneraMovNroErr
    Dim rs As ADODB.Recordset
    Dim oConect As DConecta
    Dim sql As String
    Set oConect = New DConecta
    Set rs = New ADODB.Recordset
    If oConect.AbreConexion = False Then Exit Function
    If psMovNro = "" Or Len(psMovNro) <> 25 Then
       sql = "sp_GeneraMovNro '" & Format(pdFecha & " " & oConect.GetHoraServer, "mm/dd/yyyy hh:mm:ss") & "','" & Right(psCodAge, 2) & "','" & psUser & "'"
    Else
       sql = "sp_GeneraMovNro '','','','" & psMovNro & "'"
    End If
    Set rs = oConect.Ejecutar(sql)
    If Not rs.EOF Then
        GeneraMovNro = rs.Fields(0)
    End If
    rs.Close
    Set rs = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Function
GeneraMovNroErr:
    Call RaiseError(MyUnhandledError, "NContFunciones:GeneraMovNro Method")
End Function
Public Function EmiteCtaContDesc(ByVal psCtaContCod As String) As String
    On Error GoTo EmiteCtaContDescErr
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oConect As DConecta
    
    Set oConect = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConect.AbreConexion = False Then Exit Function
        
    If Len(psCtaContCod) > 4 Then
       sSql = "SELECT RTRIM(ISNULL(c1.cCtaContDesc,'')) + ' ' + RTRIM(c.cCtaContDesc) as cCtaContDesc " _
              & "FROM   " & vsServerCom & "CtaCont c LEFT JOIN " & vsServerCom & "CtaCont c1 ON c1.cCtaContCod = substring(c.cCtaContCod,1," & Len(psCtaContCod) - 2 & ") " _
              & "WHERE  c.cCtaContCod = '" & psCtaContCod & "'"
    Else
       sSql = "SELECT cCtaContDesc FROM " & vsServerCom & "CtaCont WHERE cCtaContCod = '" & psCtaContCod & "'"
    End If
    EmiteCtaContDesc = ""
    Set rs = oConect.CargaRecordSet(sSql)
    If Not rs.BOF And Not rs.EOF Then
        EmiteCtaContDesc = Trim(rs!cCtaContDesc)
    End If
    rs.Close: Set rs = Nothing
    
    oConect.CierraConexion:  Set oConect = Nothing
    Exit Function
EmiteCtaContDescErr:
    Call RaiseError(MyUnhandledError, "DMov:EmiteCtaContDesc Method")
End Function
Public Function GetObjetos(ByVal pnObjetoCod As TpoObjetos, Optional pbObjMiembros As Boolean = True) As ADODB.Recordset
Dim sql As String
Dim oConect As DConecta
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

If pbObjMiembros Then
    sql = "Select  cObjetoCod, CONVERT(CHAR(40),cObjetoDesc) as Descripcion, nObjetoNiv  " _
         & " From    " & vsServerCom & "Objeto Where cObjetoCod Like '" & Format(pnObjetoCod, "00") & "__'  ORDER BY cObjetoCod "
Else
    sql = "Select  cObjetoCod, CONVERT(CHAR(40),cObjetoDesc) as Descripcion, nObjetoNiv  " _
         & " From    " & vsServerCom & "Objeto Where cObjetoCod Like '" & Format(pnObjetoCod, "00") & "%' ORDER BY cObjetoCod "
End If
Set rs = oConect.CargaRecordSet(sql)
Set GetObjetos = rs
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function GetSubCtaContFiltro(ByVal psCtaCod As String, ByVal psObjCod As String, ByVal psSubCta As String, Optional ByVal psFiltro As String = "") As ADODB.Recordset
Dim oCta    As New DCtaCont
Dim oConect As New DConecta
Dim prs     As ADODB.Recordset
Dim psSql   As String
On Error GoTo GetSubCtaContFiltroErr
If oConect.AbreConexion = False Then Exit Function
   Select Case Val(psObjCod)
      Case ObjCMACAgencias
         psSql = "SELECT cAgeCod cObjetoCod, cAgeDescripcion cObjetoDesc, '1' nObjetoNiv, cSubCtaCod FROM Agencias WHERE cSubCtaCod = LEFT('" & psSubCta & "',LEN(cAgeCod))"
         Set prs = oConect.CargaRecordSet(psSql)
      Case ObjCMACAgenciaArea
         psSql = "SELECT cf.cAreaCod+cf.cAgeCod cObjetoCod, RTRIM(ISNULL(age.cAgeDescripcion,''))+' '+RTRIM(ISNULL(area.cAreaDescripcion,'')) cObjetoDesc , '1' nObjetoNiv, cf.cSubCtaCod FROM CtaAreaAgeFiltro cf LEFT JOIN Areas area ON area.cAreaCod = cf.cAreaCod LEFT JOIN Agencias age ON age.cAgeCod = cf.cAgeCod WHERE cCtaContCod = '" & psCtaCod & "' and cObjetoCod = '" & psObjCod & "' and cf.cSubCtaCod = LEFT('" & psSubCta & "',LEN(cf.cSubCtaCod))"
         Set prs = oConect.CargaRecordSet(psSql)
         If prs.EOF Then
            psSql = "SELECT cf.cAreaCod+cf.cAgeCod cObjetoCod, RTRIM(ISNULL(age.cAgeDescripcion,''))+' '+RTRIM(ISNULL(area.cAreaDescripcion,'')) cObjetoDesc, '1' nObjetoNiv, cf.cSubCtaCod FROM AreaAgencia cf LEFT JOIN Areas area ON area.cAreaCod = cf.cAreaCod LEFT JOIN Agencias age ON age.cAgeCod = cf.cAgeCod WHERE cf.cSubCtaCod = LEFT('" & Right(psSubCta, 2) & "',LEN(cf.cSubCtaCod))"
            Set prs = oConect.CargaRecordSet(psSql)
         End If
      Case ObjCMACArea
         psSql = "SELECT cAreaCod cObjetoCod, cAreaDescripcion cObjetoDesc, '1' nObjetoNiv, cSubCtaCod FROM Areas WHERE cSubCtaCod = LEFT('" & psSubCta & "',LEN(cAgeCod))"
         Set prs = oConect.CargaRecordSet(psSql)
      Case ObjEntidadesFinancieras
         If psSubCta = "" Then
            Set prs = CargaCtasIF(Mid(psCtaCod, 3, 1), psFiltro)
         Else
            psSql = "SELECT cif.cIFTpo +'.'+ cif.cPersCod +'.'+ cif.cCtaIfCod cObjetoCod, cCtaIFDesc cObjetoDesc, '1' nObjetoNiv,  cCtaIFSubCta cSubCtaCod FROM CtaIF cif JOIN CtaIFFiltro ciff ON cif.cPersCod = ciff.cPersCod and cif.cCtaIfCod = ciff.cCtaIfCod " _
               & "WHERE cCtaIFSubCta = LEFT('" & psSubCta & "',LEN(cCtaIFSubCta)) AND CIF.CIFTPO+CIF.CCTAIFCOD LIKE '" & psFiltro & "' AND ciff.cctacontcod = '" & psCtaCod & "'"
            Set prs = oConect.CargaRecordSet(psSql)
         End If
      Case ObjDescomEfectivo
         Set prs = Nothing
      Case ObjPersona
         Set prs = Nothing
      Case Else
         psSql = "SELECT co.cObjetoCod, o.cObjetoDesc, '1' nObjetoNiv, co.cCtaObjSubCta cSubCtaCod FROM CtaObjFiltro co LEFT JOIN Objeto o ON o.cObjetoCod =  co.cObjetoCod  WHERE cCtaContCod = '" & psCtaCod & "'  AND '" & psSubCta & "' LIKE cCtaObjSubCta + '%' and co.cObjetoCod LIKE '" & psObjCod & "%'"
         Set prs = oConect.CargaRecordSet(psSql)
   End Select
   Set GetSubCtaContFiltro = prs
Set oCta = Nothing
Set oConect = Nothing
Set prs = Nothing
Exit Function
GetSubCtaContFiltroErr:
   Call RaiseError(MyUnhandledError, "NContFunciones: GetSubCtaContFiltro Method")
End Function

Private Function CargaCtasIF(Optional ByVal pnMoneda As Moneda = 0, Optional psFiltroCtaIF As String = "", Optional pnMuestraIF As MuestraIF = MuestraCuentas, Optional ByVal psCanje As String = "", Optional psTipoIF As CGTipoIF, Optional pbFiltraActivos As Boolean = True) As ADODB.Recordset
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
        
    lsIFFiltro = ""
    If psFiltroCtaIF <> "" Then
        ' se filtran las cuentas segun tipo
        lsIFFiltro = " AND CI.cIFTpo+CI.cCtaIfCod LIKE '" & psFiltroCtaIF & "'"
        If pbFiltraActivos Then
            lsIFFiltro = lsIFFiltro & " AND CI.cCtaIFEstado='" & gEstadoCtaIFActiva & "' "
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
    
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion = False Then Exit Function
    
sql = " SELECT   CASE WHEN NIVEL =1 THEN CPERSCOD ELSE CPERSCOD + '.' + cCtaIFCod END AS cObjetoCod , " _
    & "          Convert(char(40),CTAIFDESC)  as cObjetoDesc, NIVEL nObjetoNiv, '' cSubCtaCod " _
    & " FROM ( " _
    & "         SELECT  I.cIFTpo + '.' + CI.CPERSCOD as CPERSCOD, CI.cCtaIFCod, " _
    & "                 CONVERT(CHAR(40),CI.cCtaIFDesc) AS CTAIFDESC, " _
    & "                 LEN(CI.cCtaIFCod) AS Nivel, I.cIFTpo , I.bCanje " _
    & "         FROM    INSTITUCIONFINANC I JOIN CTAIF CI ON CI.cPersCod = I.cPersCod AND I.cIFTpo= CI.cIFTpo " _
    & "         WHERE   SUBSTRING(CI.CCTAIFCOD,1,1) NOT IN('X') " & lsFiltroMoneda & lsIFFiltro & lsFiltroTipoIF _
    & "         Union " _
    & "         SELECT  I.cIFTpo + '.' + I.CPERSCOD as CPERSCOD, '' AS CTAIF, P.CPERSNOMBRE , 1 AS NIVEL , " _
    & "                 I.cIFTpo , I.bCanje " _
    & "         FROM    INSTITUCIONFINANC I " _
    & "                 JOIN PERSONA P ON P.CPERSCOD = I.CPERSCOD " _
    & "                 JOIN (  SELECT  CI.cIFTpo, CI.CPERSCOD " _
    & "                         FROM    CTAIF CI " _
    & "                         WHERE   SUBSTRING(CI.CCTAIFCOD,1,1) NOT IN('X') " & lsFiltroMoneda & lsIFFiltro _
    & "                         ) AS C1 " _
    & "                 ON  C1.cIFTpo=I.cIFTpo AND C1.CPERSCOD= I.CPERSCOD " _
    & "         UNION Select  Replace(Str(nConsValor,2,0),' ','0') as cPerscod, '' as CtaIf , " _
    & "                 cConsDescripcion , 0 AS  NIVEL, Replace(Str(nConsValor,2,0),' ','0') as cIFTpo, 0 as bCanje " _
    & "         From    Constante " _
    & "         Where   nConsCod Like " & gCGTipoIF & " And nConsValor <> " & gCGTipoIF & " " _
    & "                 AND  Replace(Str(nConsValor,2,0),' ','0') IN (  SELECT  I.cIFTpo " _
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
    Call RaiseError(MyUnhandledError, "DCajaCtasIF:CargaCtasIF Function")
End Function

Public Sub UbicarEnRegistro(prs As Recordset, ByVal pnPos As Variant)
   On Error GoTo UbicarenRegistroErr
   If pnPos < prs.RecordCount Then
      prs.Bookmark = pnPos
   Else
      If prs.RecordCount > 0 Then
         prs.MoveLast
      End If
   End If
    Exit Sub
UbicarenRegistroErr:
    Call RaiseError(MyUnhandledError, "NContFunciones:UbicarenRegistro Method")
End Sub

Public Function PermiteModificarAsiento(psMovNro As String, Optional pbMsg As Boolean = True, Optional psFecha As String) As Boolean
Dim ldFecCie As Date
Dim ldFecMov As Date
   On Error GoTo PermiteModificarMovErr
   PermiteModificarAsiento = True
   Dim oGen As New NConstSistemas
   ldFecCie = CDate(oGen.LeeConstSistema(gConstSistCierreMensualCont))
   ldFecMov = CDate(ObtieneFechaMov(psMovNro, True))
   Set oGen = Nothing
   
   If ldFecMov <= ldFecCie Then
      PermiteModificarAsiento = False
      If pbMsg Then
         MsgBox "Asiento pertenece a mes ya Cerrado. Imposible realizar operación...!", vbInformation, "Aviso"
      End If
   End If
   Exit Function
PermiteModificarMovErr:
    Call RaiseError(MyUnhandledError, "NContFunciones:PermiteModificarMov Method")
End Function


Public Function FactorAjuste(ByVal dFecAdq As Date, ByVal dFecRep As Date, Optional nDeci As Integer = 4) As Double
Dim sSql As String
Dim R As New ADODB.Recordset
Dim FactAFecha As Double
Dim FacRep As Double
Dim oIPM As New DAjusteCont
   Set R = oIPM.CargaIPM(Format(dFecRep, gsFormatoFecha))
      If R.BOF And R.EOF Then
          R.Close: Set R = Nothing
          FactorAjuste = 0
          Exit Function
      Else
          FacRep = R!nValor
      End If
   R.Close
    
   Set R = oIPM.CargaIPM(Format(dFecAdq, gsFormatoFecha))
      If R.BOF And R.EOF Then
         R.Close: Set R = Nothing
         FactorAjuste = 0
         Exit Function
      Else
         FactAFecha = R!nValor
      End If
   R.Close
   Set R = Nothing
   Set oIPM = Nothing
   
   FactorAjuste = CDbl(Format(FacRep / FactAFecha, "#0." & String(nDeci, "0")))
   
End Function
Public Function GetSaldoAreaAgencia(ByVal psFormatoFecha As String, ByVal pdFecha As Date, _
                                    ByVal psCtaContCod As String, ByVal psAreaCod As String, ByVal psAgeCod As String) As Currency
Dim sql As String
Dim oConect As DConecta
Dim rs As ADODB.Recordset


Set oConect = New DConecta
Set rs = New ADODB.Recordset

sql = " SELECT   dAreaSaldoFecha, nAreaSaldoImporte , ISNULL(MOVDIA.NDEBE,0) AS DEBE, " _
    & "         ISNULL(MOVDIA.NHABER,0) AS HABER,  nAreaSaldoImporte + ISNULL(MOVDIA.NDEBE,0) + ISNULL(MOVDIA.NHABER,0) AS SALDODIA " _
    & " FROM    CtaObjAreaAgenciaSaldo CS " _
    & "         LEFT JOIN (SELECT   MC.CCTACONTCOD , " _
    & "                             SUM(CASE WHEN MC.NMOVIMPORTE>0 THEN ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) ELSE 0 END) AS NDEBE, " _
    & "                             SUM(CASE WHEN MC.NMOVIMPORTE<0 THEN ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) ELSE 0 END) AS NHABER " _
    & "                     FROM    MOV M " _
    & "                             JOIN MOVCTA MC ON MC.nMOVNRO = M.nMOVNRO " _
    & "                             LEFT JOIN MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM=MC.nMOVITEM " _
    & "                     WHERE   SUBSTRING(M.CMOVNRO,1,8)='" & Format(pdFecha, "yyyymmdd") & "' " _
    & "                             AND M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") " _
    & "                             AND M.nMOVESTADO='" & gMovEstContabMovContable & "' " _
    & "                     GROUP BY MC.CCTACONTCOD) AS MOVDIA " _
    & "         ON MOVDIA.CCTACONTCOD = CS.CCTACONTCOD " _
    & " WHERE   CS.cCtaContCod Like '" & psCtaContCod & "' AND CS.cAreaCod ='" & psAreaCod & "' AND CS.cAgeCod='" & psAgeCod & "' " _
    & "         AND dAreaSaldoFecha= (  SELECT  MAX(dAreaSaldoFecha) " _
    & "                                 FROM    CtaObjAreaAgenciaSaldo CS1 " _
    & "                                 Where   CS1.cCtaContCod = CS.cCtaContCod " _
    & "                                         AND CS.cAreaCod = CS1.cAreaCod AND CS.cAgeCod=CS1.cAgeCod " _
    & "                                         AND dAreaSaldoFecha<'" & Format(pdFecha, psFormatoFecha) & " 23:59:59')"

If oConect.AbreConexion = False Then Exit Function

Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetSaldoAreaAgencia = rs!SaldoDia
End If
rs.Close
Set rs = Nothing
End Function

Public Function GetSaldoCtaCont(ByVal pdFecha As Date, ByVal psCtaContCod As String, Optional ByVal pnMoneda As Integer = 1) As Currency
Dim sql As String
Dim oConect As DConecta
Dim rs As ADODB.Recordset

Set oConect = New DConecta
Set rs = New ADODB.Recordset

'sql = " SELECT  dCtaSaldoFecha, nCtaSaldoImporte , ISNULL(MOVDIA.NDEBE,0) AS DEBE, " _
'    & "         ISNULL(MOVDIA.NHABER,0) AS HABER,  nCtaSaldoImporte + ISNULL(MOVDIA.NDEBE,0) + ISNULL(MOVDIA.NHABER,0) AS SALDODIA " _
'    & " FROM    CtaSaldo CS " _
'    & "         LEFT JOIN (SELECT   MC.CCTACONTCOD , " _
'    & "                             SUM(CASE WHEN MC.NMOVIMPORTE>0 THEN ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) ELSE 0 END) AS NDEBE, " _
'    & "                             SUM(CASE WHEN MC.NMOVIMPORTE<0 THEN ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) ELSE 0 END) AS NHABER " _
'    & "                     FROM    MOV M " _
'    & "                             JOIN MOVCTA MC ON MC.nMOVNRO = M.nMOVNRO " _
'    & "                             LEFT JOIN MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM=MC.nMOVITEM " _
'    & "                     WHERE   SUBSTRING(M.CMOVNRO,1,8)='" & Format(pdFecha, "yyyymmdd") & "' " _
'    & "                             AND M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") " _
'    & "                             AND M.nMOVESTADO='" & gMovEstContabMovContable & "' " _
'    & "                     GROUP BY MC.CCTACONTCOD) AS MOVDIA " _
'    & "         ON MOVDIA.CCTACONTCOD = CS.CCTACONTCOD " _
'    & "         WHERE   CS.cCtaContCod Like '" & psCtaContCod & "' " _
'    & "                 AND dCtaSaldoFecha= (   SELECT  MAX(dCtaSaldoFecha) " _
'    & "                                         FROM    CtaSaldo CS1 " _
'    & "                                         Where   CS1.cCtaContCod = CS.cCtaContCod " _
'    & "                                                 AND dCtaSaldoFecha<'" & Format(pdFecha, psFormatoFecha) & " 23:59:59')"

sql = "SELECT dbo.GetSaldoCta('" & Format(pdFecha, gsFormatoFecha) & "','" & psCtaContCod & "', " & pnMoneda & ") as SaldoDia "
If oConect.AbreConexion = False Then Exit Function
Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetSaldoCtaCont = rs!SaldoDia
End If
rs.Close
Set rs = Nothing
End Function
Public Function ExisteMovimiento(psMovNro As String, psOpeCod As String) As Boolean
Dim psSql   As String
Dim oConect As New DConecta
Dim prs     As ADODB.Recordset
ExisteMovimiento = True
   psSql = "SELECT m.cMovNro " _
         & "FROM Mov m " _
         & "WHERE m.cMovNro LIKE '" & psMovNro & "%' and m.nMovEstado = '" & gMovEstContabMovContable & "'" _
         & " and not m.nMovFlag in ('" & gMovFlagEliminado & "','" & gMovFlagModificado & "') and cOpeCod Like '" & psOpeCod & "%'"
   oConect.AbreConexion
   Set prs = oConect.CargaRecordSet(psSql)
   If prs.EOF Then
      ExisteMovimiento = False
   End If
   Set oConect = Nothing
   prs.Close: Set prs = Nothing
End Function

Public Function GetObjetosArbol(ByVal pnObjetoCod As String, ByVal psObjFiltro As String, ByVal pnNivel As Integer) As ADODB.Recordset
Dim sql As String
Dim oConect As New DConecta
Dim rs As New ADODB.Recordset
   If oConect.AbreConexion = False Then Exit Function
   sql = vsServerCom & "spGetTreeObj '" & Format(pnObjetoCod, "00") & "', " & pnNivel & ", '" & psObjFiltro & "'"
   Set rs = oConect.CargaRecordSet(sql)
   Set GetObjetosArbol = rs
   oConect.CierraConexion: Set oConect = Nothing
Set rs = Nothing
End Function

'Public Function GetLibroDiario(ByVal dFecIni As Date, ByVal dFecFin As Date, psAge As String, psOpe As String) As Recordset
'Dim sAgeCond As String, sOpeCond As String
'Dim sql As String
'Dim oConect As New DConecta
' On Error GoTo GetLibroDiarioErr
'
' If oConect.AbreConexion = False Then Exit Function
'    sAgeCond = "": sOpeCond = ""
'    If psAge <> "" Then
'       sAgeCond = " And SubString(M.cMovNro,18,2) = '" & Right(psAge, 2) & "' "
'    End If
'    If psOpe <> "" Then
'       sOpeCond = " And O.cOpeGruCod = '" & Trim(psOpe) & "' "
'    End If
'
'    sql = "Select M1.nMovCorrela, M.cMovNro, O.cOpeGruCod, MC.cCtaContCod, C.cCtaContDesc, MC.nMovImporte, og.cOpeGruDesc " _
'        & "From Mov M JOIN MovCont M1 ON M1.nMovNro = M.nMovNro " _
'        & "           Join MovCta MC On M.nMovNro = MC.nMovNro " _
'        & "           Join " & vsServerCom & "CtaCont C On MC.cCtaContCod = C.cCtaContCod " _
'        & "           Join " & vsServerCom & "OpeTpo  O On M.cOpeCod = O.cOpeCod " _
'        & "           Join " & vsServerCom & "OpeGrupo Og ON og.cOpeGruCod = o.cOpeGruCod " _
'        & "Where M.nMovEstado = " & gMovEstContabMovContable & " and not M.nMovFlag IN (" & gMovFlagEliminado & "," & gMovFlagModificado & ") and substring(M.cMovNro,1,8) between " _
'        & "      '" & Format(dFecIni, "yyyymmdd") & "' And '" & Format(dFecFin, "yyyymmdd") _
'        & "'     AND mc.cCtaContCod <> '' " & sAgeCond & sOpeCond _
'        & "ORDER BY M1.nMovCorrela, substring(M.cMovNro,1,8), substring(M.cMovNro,18,2), O.cOpeGruCod, M.cMovNro, Mc.nMovItem "
'    Set GetLibroDiario = oConect.CargaRecordSet(sql)
'    oConect.CierraConexion: Set oConect = Nothing
'Exit Function
'GetLibroDiarioErr:
'   Call RaiseError(MyUnhandledError, "NContFunciones: GetLibroDiario Method")
'End Function

Public Function GetLibroDiario(ByVal dFecIni As Date, ByVal dFecFin As Date, psAge As String, psOpe As String, Optional pbHistoCtaContb As Boolean = False) As ADODB.Recordset 'JACA 20111229
Dim sAgeCond As String, sOpeCond As String
Dim sql As String
Dim oConect As New DConecta
Dim rsLibro As ADODB.Recordset
Set rsLibro = New ADODB.Recordset
Dim Conn As New Connection

Dim xConn As New DConecta

 On Error GoTo GetLibroDiarioErr
 If xConn.AbreConexion Then
    Set rsLibro = New ADODB.Recordset

'    sql = ("dbo.spGetLibroDiario '" & Format(dFecIni, "yyyymmdd") & "' ,'" & Format(dFecFin, "yyyymmdd") & "','" & Right(psAge, 2) & "' ,'" & Trim(psOpe) & "'")
    
    'JACA 20111229**************************************************
    If Not pbHistoCtaContb Then
        Set rsLibro = xConn.Ejecutar("dbo.spGetLibroDiario '" & Format(dFecIni, "yyyymmdd") & "' ,'" & Format(dFecFin, "yyyymmdd") & "','" & Right(psAge, 2) & "' ,'" & Trim(psOpe) & "'")
    Else
        'Descomentar cuando se tenga nuevo plan contable
        'Set rsLibro = xConn.Ejecutar("dbo.Histo_spGetLibroDiario '" & Format(dFecIni, "yyyymmdd") & "' ,'" & Format(dFecFin, "yyyymmdd") & "','" & Right(psAge, 2) & "' ,'" & Trim(psOpe) & "'")
        Set rsLibro = xConn.Ejecutar("dbo.spGetLibroDiario '" & Format(dFecIni, "yyyymmdd") & "' ,'" & Format(dFecFin, "yyyymmdd") & "','" & Right(psAge, 2) & "' ,'" & Trim(psOpe) & "'")
    End If
    'JACA END******************************************************
    
    If Not rsLibro.EOF Then
       Set GetLibroDiario = rsLibro
       Exit Function
    End If
 End If

Exit Function

GetLibroDiarioErr:
   Call RaiseError(MyUnhandledError, "NContFunciones: GetLibroDiario Method")
End Function

Public Function ValidaDocumento(nDocTpo As TpoDoc, sDocNro As String, sCodProv As String) As Boolean
Dim sSql As String
Dim prs  As ADODB.Recordset
Dim oConect As New DConecta

ValidaDocumento = True
oConect.AbreConexion

sSql = "SELECT a.nMovNro FROM MovGasto a JOIN MovDoc b ON b.nMovNro = a.nMovNro JOIN Mov m ON m.nMovNro = a.nMovNro " _
     & "WHERE  not m.nMovFlag IN ('" & gMovFlagEliminado & "','" & gMovFlagModificado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "') and a.cPersCod = '" & sCodProv & "' and b.nDocTpo = '" & nDocTpo & "' and b.cDocNro = '" & sDocNro & "'"
Set prs = oConect.CargaRecordSet(sSql)
If Not prs.EOF Then
   ValidaDocumento = False
   
End If
RSClose prs

End Function


Public Function EsFeriado(pdFecha As Date) As Boolean
Dim sSql As String
Dim prs  As ADODB.Recordset
Dim oConect As New DConecta

oConect.AbreConexion
sSql = "Select * from feriado where month(dFeriado) = " & Month(pdFecha) & " and day(dFeriado) = " & Day(pdFecha) & " and cDescrip <> 'DOMINGO' and not cDescrip LIKE 'FERIADO%'"
Set prs = oConect.CargaRecordSet(sSql)
EsFeriado = Not prs.EOF
RSClose prs


End Function

Public Function VerificaCtaDetraccion(ByVal CtaDetra As String) As Integer
Dim sSql As String
Dim prs  As ADODB.Recordset
Dim oConect As New DConecta

oConect.AbreConexion
'***Modificado por ELRO según Acta N° 328-2011/TI-D
'sSql = "Select * from CtaDetraccionesDet where cCtaContCod = '" & CtaDetra & "' "
sSql = "exec stp_sel_VerificarCtaDetraccion " & CtaDetra
'***Fin Modificado por ELRO************************

Set prs = oConect.CargaRecordSet(sSql)

If prs.BOF Or prs.EOF Then
    VerificaCtaDetraccion = 0
Else
    VerificaCtaDetraccion = 1
End If
RSClose prs
End Function

Public Function GetImpuestoCuentaDetrac(ByVal CtaDetra As String) As ADODB.Recordset
Dim sSql As String
Dim prs  As ADODB.Recordset
Dim oConect As New DConecta

oConect.AbreConexion
sSql = " Select distinct d.nDetraPorc" & _
       " from CtaDetracciones d INNER JOIN CtaDetraccionesDet dt on dt.nDetraCod =d.nDetraCod " & _
       " where dt.cCtaContCod =  '" & CtaDetra & "'  "

Set prs = oConect.CargaRecordSet(sSql)

If Not prs.EOF Then
   Set GetImpuestoCuentaDetrac = prs
   Exit Function
End If
    
RSClose prs
End Function


Public Function GetCuentaEmbargo(ByVal pOpeCod As String, ByVal pOrden As String) As ADODB.Recordset
Dim sSql As String
Dim prs  As ADODB.Recordset
Dim oConect As New DConecta

oConect.AbreConexion
sSql = " Select cCtaContcod" & _
       " from OpeCta " & _
       " where cOpeCod = '" & pOpeCod & "' and cOpeCtaORden =  '" & pOrden & "'  "

Set prs = oConect.CargaRecordSet(sSql)

If Not prs.EOF Then
   Set GetCuentaEmbargo = prs
   Exit Function
End If
    
RSClose prs
End Function


'*** PEAC 20101104
Public Function GetImptoIGV() As Currency
Dim sSql As String
Dim prs  As ADODB.Recordset
Dim oConect As New DConecta

oConect.AbreConexion

sSql = " exec stp_sel_ObtieneImptoIGV "

Set prs = oConect.CargaRecordSet(sSql)

If Not prs.EOF Then
    GetImptoIGV = IIf(IsNull(prs!nImpTasa), 0, prs!nImpTasa)
Else
    GetImptoIGV = 0
End If

RSClose prs
End Function

'*** PEAC 20110427
Public Function PermiteBorrarRegPorOpe(ByVal psOpeTpo As String, ByRef pdFecIni As Date, ByRef pdFecFin As Date) As Boolean
Dim sSql As String
Dim prs  As ADODB.Recordset
Dim oConect As New DConecta

oConect.AbreConexion

sSql = "exec stp_sel_BuscaPermisoEliminarRegPorOpe '" & psOpeTpo & "'"

Set prs = oConect.CargaRecordSet(sSql)

If Not prs.EOF Then
    If prs!nPermiso Then
        PermiteBorrarRegPorOpe = True
        pdFecIni = prs!dFecIni
        pdFecFin = prs!dFecFin
    Else
        PermiteBorrarRegPorOpe = False
    End If
Else
    PermiteBorrarRegPorOpe = False
End If

RSClose prs

End Function

'***Modificado por ELRO el 20110924, según Acta N° 263-2011/TI-D
Public Function verificarUltimoNivelCta(ByVal cCtaContCod As String) As Boolean
Dim sSql As String
Dim prs  As ADODB.Recordset
Dim oConect As New DConecta

oConect.AbreConexion

sSql = "exec stp_sel_VerificarUltimoNivelCta'" & cCtaContCod & "'"

Set prs = oConect.CargaRecordSet(sSql)

If Not prs.EOF Then
    If prs!nCantidad = 1 Then
        verificarUltimoNivelCta = True
    Else
        verificarUltimoNivelCta = False
    End If
Else
    verificarUltimoNivelCta = False
End If

RSClose prs

End Function
'***Fin Modificado por ELRO
'***Agregado por ELRO el 20120423, según Acta N° 328-2011/TI-D
Public Function obtenerDocumentos() As ADODB.Recordset
Dim sSql As String
Dim oDConecta As New DConecta

On Error GoTo obtenerDocumentosErr

    If oDConecta.AbreConexion = False Then Exit Function

    sSql = "exec stp_sel_ObtenerDocumentos"

    Set obtenerDocumentos = oDConecta.CargaRecordSet(sSql)

    Set oDConecta = Nothing
    Exit Function

obtenerDocumentosErr:
    Set oDConecta = Nothing
    Call RaiseError(MyUnhandledError, "NContFunciones:obtenerDocumentos Method")

End Function

Public Function obtenerConceptoDetracciones() As ADODB.Recordset
Dim sSql As String
Dim oDConecta As New DConecta

On Error GoTo obtenerConceptoDetraccionesErr

    If oDConecta.AbreConexion = False Then Exit Function

    sSql = "exec stp_sel_ObtenerConceptoDetracciones"

    Set obtenerConceptoDetracciones = oDConecta.CargaRecordSet(sSql)

    Set oDConecta = Nothing
    Exit Function

obtenerConceptoDetraccionesErr:
    Set oDConecta = Nothing
    Call RaiseError(MyUnhandledError, "NContFunciones:obtenerConceptoDetracciones Method")

End Function

Public Sub registrarConceptoDetracion(ByVal psDetraDes As String, _
                                      ByVal pnDetraPorc As Currency, _
                                      ByVal pnDocDetra As Integer, _
                                      ByVal pnRango1 As Currency, _
                                      ByVal pnRango2 As Currency, _
                                      ByVal psMov As String)
Dim lssql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta

On Error GoTo registrarConceptoDetracionErr

    If oDConecta.AbreConexion = False Then Exit Sub
    lssql = "exec stp_ins_RegistrarConceptoDetraccion '" & psDetraDes & "', " & pnDetraPorc & ", " & pnDocDetra & ", " & pnRango1 & ", " & pnRango2 & ", '" & psMov & "'"
    oDConecta.CargaRecordSet (lssql)

    Set oDConecta = Nothing
    Exit Sub

registrarConceptoDetracionErr:
    Set oDConecta = Nothing
    Call RaiseError(MyUnhandledError, "NContFunciones:registrarConceptoDetracion Method")
End Sub

Public Sub actualizarConceptoDetraccion(ByVal pnDetraCod As Integer, _
                                         ByVal psDetraDes As String, _
                                         ByVal pnDetraPorc As Currency, _
                                         ByVal pnDocDetra As Integer, _
                                         ByVal pnRango1 As Currency, _
                                         ByVal pnRango2 As Currency, _
                                         ByVal psMov As String)
Dim lssql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta

On Error GoTo actualizarConceptoDetraccionErr

    If oDConecta.AbreConexion = False Then Exit Sub
    lssql = "exec stp_upd_ActualizarConceptoDetraccion " & pnDetraCod & ", '" & psDetraDes & "', " & pnDetraPorc & ", " & pnDocDetra & ", " & pnRango1 & ", " & pnRango2 & ", '" & psMov & "'"
    oDConecta.CargaRecordSet (lssql)

    Set oDConecta = Nothing
    Exit Sub

actualizarConceptoDetraccionErr:
    Set oDConecta = Nothing
    Call RaiseError(MyUnhandledError, "NContFunciones:actualizarConceptoDetraccion Method")
End Sub

Public Sub eliminarConceptoDetraccion(ByVal pnDetraCod As Integer, _
                                         ByVal psMov As String)
Dim lssql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta

   On Error GoTo eliminarConceptoDetraccionErr

    If oDConecta.AbreConexion = False Then Exit Sub
    lssql = "exec stp_del_EliminarConceptoDetraccion " & pnDetraCod & ", '" & psMov & "'"
    oDConecta.CargaRecordSet (lssql)

    Set oDConecta = Nothing
    Exit Sub

eliminarConceptoDetraccionErr:
    Set oDConecta = Nothing
    Call RaiseError(MyUnhandledError, "NContFunciones:eliminarConceptoDetraccion Method")
End Sub

Public Function obtenerCtaDetraccionesDet(ByVal pnDetraCod As Integer) As ADODB.Recordset
Dim sSql As String
Dim oDConecta As New DConecta

On Error GoTo obtenerCtaDetraccionesDetErr

    If oDConecta.AbreConexion = False Then Exit Function

    sSql = "exec stp_sel_CtaDetraccionesDet " & pnDetraCod

    Set obtenerCtaDetraccionesDet = oDConecta.CargaRecordSet(sSql)

    Set oDConecta = Nothing
    Exit Function

obtenerCtaDetraccionesDetErr:
    Set oDConecta = Nothing
    Call RaiseError(MyUnhandledError, "NContFunciones:CtaDetraccionesDet Method")

End Function

Public Function obtenerNumeroItemCtaDetraccionesDet(ByVal pnDetraCod As Integer) As Integer
Dim sSql As String
Dim rsNumeroItem  As ADODB.Recordset
Set rsNumeroItem = New ADODB.Recordset
Dim oDConecta As New DConecta

On Error GoTo obtenerNumeroItemCtaDetraccionesDetErr

    If oDConecta.AbreConexion = False Then Exit Function

    sSql = "exec stp_sel_ConsultarNumeroItemCtaDetraccionesDet " & pnDetraCod

    Set rsNumeroItem = oDConecta.CargaRecordSet(sSql)

    If Not rsNumeroItem.BOF And Not rsNumeroItem.EOF Then
        obtenerNumeroItemCtaDetraccionesDet = rsNumeroItem.Fields(0)
    Else
        obtenerNumeroItemCtaDetraccionesDet = 1
    End If

    Set rsNumeroItem = Nothing
    Set oDConecta = Nothing
    Exit Function

obtenerNumeroItemCtaDetraccionesDetErr:
    Set rsNumeroItem = Nothing
    Set oDConecta = Nothing
    Call RaiseError(MyUnhandledError, "NContFunciones:obtenerNumeroItemCtaDetraccionesDet Method")
End Function

Public Sub registrarCtaDetraccionesDet(ByVal pnDetraCod As Integer, _
                                        ByVal pnItem As Integer, _
                                        ByVal psCtaContCod As Currency, _
                                        ByVal psMov As String)
Dim lssql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta

On Error GoTo registrarCtaDetraccionesDetErr

    If oDConecta.AbreConexion = False Then Exit Sub
    lssql = "exec stp_ins_RegistrarCtaDetraccionesDet " & pnDetraCod & ", " & pnItem & ", '" & psCtaContCod & "', '" & psMov & "'"
    oDConecta.CargaRecordSet (lssql)

    Set oDConecta = Nothing
    Exit Sub

registrarCtaDetraccionesDetErr:
    Set oDConecta = Nothing
    Call RaiseError(MyUnhandledError, "NContFunciones:registrarCtaDetraccionesDet Method")
End Sub

Public Sub actualizarCtaDetraccionesDet(ByVal pnDetraCod As Integer, _
                                         ByVal pnItem As Integer, _
                                         ByVal psCtaContCod As String, _
                                         ByVal psMov As String)
Dim lssql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta

On Error GoTo actualizarCtaDetraccionesDetErr

    If oDConecta.AbreConexion = False Then Exit Sub
    lssql = "exec stp_upd_ActualizarCtaDetraccionesDet " & pnDetraCod & ", " & pnItem & ", '" & psCtaContCod & "', '" & psMov & "'"
    oDConecta.CargaRecordSet (lssql)

    Set oDConecta = Nothing
    Exit Sub

actualizarCtaDetraccionesDetErr:
    Set oDConecta = Nothing
    Call RaiseError(MyUnhandledError, "NContFunciones:actualizarCtaDetraccionesDet Method")
End Sub

Public Sub eliminarCtaDetraccionesDet(ByVal pnDetraCod As Integer, _
                                       ByVal pnItem As Integer, _
                                       ByVal psMov As String)
Dim lssql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta

   On Error GoTo eliminarCtaDetraccionesDetErr

    If oDConecta.AbreConexion = False Then Exit Sub
    lssql = "exec stp_upd_EliminarCtaDetraccionesDet " & pnDetraCod & ", " & pnItem & ", '" & psMov & "'"
    oDConecta.CargaRecordSet (lssql)

    Set oDConecta = Nothing
    Exit Sub

eliminarCtaDetraccionesDetErr:
    Set oDConecta = Nothing
    Call RaiseError(MyUnhandledError, "NContFunciones:eliminarCtaDetraccionesDet Method")
End Sub

Public Function validarCuentaContableConsolidada(ByVal psBalanceAnio As String, _
                                                 ByVal psBalanceMes As String, _
                                                 ByVal psCtaContCod As String) As Integer
Dim sSql As String
Dim oDConecta As New DConecta
Dim rsCtaCont As New ADODB.Recordset

On Error GoTo validarCuentaContableConsolidadaErr

    If oDConecta.AbreConexion = False Then Exit Function

    sSql = "exec stp_sel_ValidarCuentaContableConsolidada '" & psBalanceAnio & "', '" & psBalanceMes & "', '" & psCtaContCod & "' "

    Set rsCtaCont = oDConecta.CargaRecordSet(sSql)

    If Not rsCtaCont.BOF And Not rsCtaCont.EOF Then
        validarCuentaContableConsolidada = rsCtaCont.Fields(0)
    Else
        validarCuentaContableConsolidada = 0
    End If

    Set rsCtaCont = Nothing
    Set oDConecta = Nothing
    Exit Function

validarCuentaContableConsolidadaErr:
    Set oDConecta = Nothing
    Call RaiseError(MyUnhandledError, "NContFunciones:validarCuentaContableConsolidada Method")

End Function

'***Fin Agregado por ELRO*******************************
'***Agregado por ELRO el 20120423, según OYP-RFC005-2012
Public Function obtenerSiFechaEsLaborable(ByVal pdFecha As Date, ByVal psAgeCod As String) As Boolean
Dim lssql As String
Dim rsFecha  As ADODB.Recordset
Set rsFecha = New ADODB.Recordset
Dim oDConecta As DConecta
Set oDConecta = New DConecta

If oDConecta.AbreConexion = False Then Exit Function
lssql = "exec stp_sel_ObtenerSiFechaEsLaborable '" & Format(pdFecha, "yyyyMMdd") & "', '" & psAgeCod & "'"
Set rsFecha = oDConecta.CargaRecordSet(lssql)

If Not rsFecha.BOF And Not rsFecha.EOF Then
    If rsFecha!cFechaLaborable = 0 Then
        obtenerSiFechaEsLaborable = True
    Else
        obtenerSiFechaEsLaborable = False
    End If
Else
    obtenerSiFechaEsLaborable = False
End If

Set rsFecha = Nothing
Set oDConecta = Nothing

End Function
'***Fin Agregado por ELRO*******************************

'***PEAC 20130129
Public Sub InsertaFiltroObjetos(psTpoIf, psPersCtaIF, psCtaIf, psCtaContCod, psTipoCtaIf)

    Dim lssql As String
    Dim oDConecta As DConecta
    Set oDConecta = New DConecta
    
    On Error GoTo InsertaFiltroObjetosErr
    
        If oDConecta.AbreConexion = False Then Exit Sub
        lssql = "exec stp_ins_InsertaFiltroObjetos '" & psTpoIf & "','" & psPersCtaIF & "','" & psCtaIf & "','" & psCtaContCod & "','" & psTipoCtaIf & "'"
        oDConecta.CargaRecordSet (lssql)
    
        Set oDConecta = Nothing
        Exit Sub
    
InsertaFiltroObjetosErr:
    Set oDConecta = Nothing
    Call RaiseError(MyUnhandledError, "NContFunciones:InsertaFiltroObjetos Method")

End Sub

Public Function ExisteMovProvEspecificaCOVID(psFecha As String, psMoneda As String) As Boolean
Dim psSql As String
Dim oConect As New DConecta
Dim prs As ADODB.Recordset
    psSql = "exec stp_sel_ExisteProvEspecificaCOVID '" & psFecha & "', '" & psMoneda & "'"
    oConect.AbreConexion
    Set prs = oConect.CargaRecordSet(psSql)
    If prs!psExiste = "1" Then
        ExisteMovProvEspecificaCOVID = True
    Else
        ExisteMovProvEspecificaCOVID = False
    End If
Set oConect = Nothing
prs.Close: Set prs = Nothing
End Function 'NAGL 202102 Según ACTA N°017-2021
