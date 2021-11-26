VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMContFunciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim vsServerCom As String
Dim vsServerPers As String
Dim oError As COMConecta.COMErrorHandling

Public Function GeneraDocNro(psDocTpo As TpoDoc, psMoneda As Moneda, Optional psDocSerie As String = "", Optional psAgencia As String = "") As String
    On Error GoTo GeneraMovNroErr
    Dim ssql As String
    Dim rs As New ADODB.Recordset
    Dim oconect As COMConecta.DCOMConecta
 
    Set oconect = New COMConecta.DCOMConecta
    If oconect.AbreConexion = False Then Exit Function
      
        If psDocSerie <> "" Then
            ssql = "SELECT max(cDocNro) AS cDocCorrela FROM Movdoc md JOIN Mov m ON m.nMovNro = md.nMovNro " _
                & "WHERE   m.nMovFlag = ('" & gMovFlagVigente & "') and nDocTpo = '" & Format(psDocTpo, "00") & "' and substring(cDocNro,1," & Len(psDocSerie) & ") = '" & psDocSerie & "'"
        Else
            ssql = "SELECT  cDocNro AS cDocCorrela , md.nMovnro " _
                & " FROM    MovDoc MD JOIN MOV M on M.nMovNro = md.nMovNro  " _
                & " WHERE   nDocTpo = '" & Format(psDocTpo, "00") & "'  " _
                & "         and md.nmovnro = (  Select Max(MD1.nmovnro)  " _
                & "                             From MovDoc MD1 JOIN Mov M1 on M1.nMovNro= MD1.nMovnro " _
                & "                             WHERE MD1.nDocTpo = '" & Format(psDocTpo, "00") & "' and Substring(M.cOpeCod,3,1) ='" & psMoneda & "'  )  "
                
                'and  m.nMovFlag = ('" & gMovFlagVigente & "')
                'and  m1.nMovFlag = ('" & gMovFlagVigente & "')
        End If
        Set rs = oconect.CargaRecordSet(ssql)
        If Not rs.EOF And Not rs.BOF Then
            If Not IsNull(rs!cDocCorrela) Then
                If psDocSerie <> "" Then
                   GeneraDocNro = psDocSerie & "-" & Format(Val(Mid(rs!cDocCorrela, Len(psDocSerie) + 2, 20)) + 1, String(8, "0"))
                Else
                   If Mid(rs!cDocCorrela, 4, 1) = "-" Then
                      GeneraDocNro = Mid(rs!cDocCorrela, 1, 3) & "-" & Format(Val(Mid(rs!cDocCorrela, 5, 20)) + 1, String(8, "0"))
                   Else
                      GeneraDocNro = Format(Val(rs!cDocCorrela) + 1, String(8, "0"))
                   End If
                End If
        End If
        
If psDocTpo > -1 Then
   If psDocSerie <> "" Then
      ssql = "SELECT max(cDocNro) AS cDocCorrela FROM  movdoc md JOIN Mov m ON m.nMovNro = md.nMovNro " _
           & "WHERE nDocTpo = '" & psDocTpo & "' and substring(cDocNro,1," & Len(psDocSerie) & ") = '" & psDocSerie & "' " & IIf(psAgencia = "", "", " and SubString(m.cMovNro,18,2) = '" & Right(psAgencia, 2) & "' ")
   Else
      If psMoneda = "" Then
         ssql = "SELECT max(cDocNro) AS cDocCorrela FROM  movdoc md JOIN Mov m ON m.nMovNro = md.nMovNro " _
               & "WHERE nDocTpo = '" & psDocTpo & "' " & IIf(psAgencia = "", "", " and SubString(m.cMovNro,18,2) = '" & Right(psAgencia, 2) & "' ")
      Else
         ssql = "SELECT  cDocNro AS cDocCorrela , md.nMovnro " _
               & " FROM    movdoc md " _
               & " WHERE   nDocTpo = '" & psDocTpo & "' " _
               & "         and md.nMovNro = (  Select Max(MD1.nmovnro) " _
               & "                             From MovDoc MD1 JOIN Mov M on M.nMovNro= MD1.nMovnro " _
               & "                             WHERE MD1.nDocTpo = '" & psDocTpo & "' and Substring(M.cOpeCod,3,1) ='" & psMoneda & "'" & IIf(psAgencia = "", "", " and SubString(m.cMovNro,18,2) = '" & Right(psAgencia, 2) & "' ") & ") "
      End If
   End If
   oconect.AbreConexion
   Set rs = oconect.CargaRecordSet(ssql)
   If rs.EOF Then
      GeneraDocNro = ""
   Else
      If Not IsNull(rs!cDocCorrela) Then
         If psDocSerie <> "" Then
            GeneraDocNro = psDocSerie & "-" & Format(Val(Mid(rs!cDocCorrela, Len(psDocSerie) + 2, 20)) + 1, String(8, "0"))
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
      If psDocSerie <> "" Then
         GeneraDocNro = psDocSerie & "-" & GeneraDocNro
      End If
   End If
    'If psDocTpo = TpoDocOrdenPago Then
        'Dim oCaja As New nCajaGeneral
        'Para verificar las OP entregadas a Caja General
        'Cuando este implementado el Centralizado del NEGOCIO
        'GeneraDocNro = oCaja.VerificaOPRango(GeneraDocNro, psMoneda)
        'Set oCaja = Nothing
    End If
    oconect.CierraConexion
End If

GeneraMovNroErr:
    Err.Raise 50001, "NContFunciones: GeneraDocNro", "Especificar Tipo de Documento para Generar Número"
    
End Function
'Public Function GetFiltroObjetos(ByVal pnTipoObj As TpoObjetos, ByVal psCtaContCod As String, ByVal psObjetoCod As String, Optional lbMuestraCta As Boolean = True, Optional lbExisteFiltro As Boolean = False) As String

'Public Function GeneraDocNro(psDocTpo As TpoDoc, psMoneda As Moneda, Optional psDocSerie As String = "") As String
'    On Error GoTo GeneraMovNroErr
'    Dim sSql As String
'    Dim rs As New ADODB.Recordset
'    Dim oConect As COMConecta.DCOMConecta
'    Dim oCaja As nCajaGeneral
'
'    Set oCaja = New nCajaGeneral
'    Set oConect = New COMConecta.DCOMConecta
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
    Dim Sql As String
    Dim rs   As New ADODB.Recordset
    Dim oconect As COMConecta.DCOMConecta
    Dim lsPersCtaIF As String
    Dim lsCtaIf As String
    Dim lsTpoIf As String
    
    Set oconect = New COMConecta.DCOMConecta
    If oconect.AbreConexion = False Then Exit Function
    GetFiltroObjetos = ""
    Select Case pnTipoObj
        Case ObjCMACAgencias
            Sql = "SELECT cAgeCod,  cSubCtaCod as SubCta  FROM " & vsServerCom & "Agencias where cAgeCod='" & psObjetoCod & "'"
        Case ObjCMACArea
            Sql = "SELECT cAreaCod, cSubCtaCod  as SubCta  FROM " & vsServerCom & "AREAS where cAreaCod='" & psObjetoCod & "'"
        Case ObjCMACAgenciaArea
           Sql = "     SELECT   AA.cAreaCod, AA.cAgeCod, " _
                & "             CASE " _
                & "                  WHEN CF.cSubctaCod IS NULL THEN AA.cSubCtaCod " _
                & "                  Else CF.cSubCtaCod END As SubCta " _
                & "    FROM     " & vsServerCom & "AREAAGENCIA AA " _
                & "             LEFT JOIN ( Select cAreaCod , cAgeCod , cSubCtaCod " _
                & "                          From " & vsServerCom & "CtaAreaAgeFiltro " _
                & "                          WHERE  cCtaContcod in ('" & psCtaContCod & "')) AS CF " _
                & "             ON CF.cAreacod = AA.cAreaCod and CF.cAgeCod = AA.cAgeCod " _
                & "    WHERE    AA.cAgeCod='" & Mid(psObjetoCod, 4, 2) & "' AND  AA.cAreaCod='" & Mid(psObjetoCod, 1, 3) & "'"
                
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
                Sql = "select cPersCod, cSubCtaContcod SubCta FROM InstitucionFinanc where cPersCod = '" & psObjetoCod & "' "
            Else
                Sql = "Select cPersCod, cCtaIfCod, cCtaContCod, cCtaIFSubCta as SubCta From " & vsServerCom & "CtaIFFiltro WHERE  cPersCod = '" & lsPersCtaIF & "' and cIFTpo='" & lsTpoIf & "' AND cCtaIfCod='" & lsCtaIf & "' AND cCtaContCod in ('" & psCtaContCod & "')"
            End If
        Case Else
            Sql = "SELECT cCtaObjSubCta as SubCta FROM " & vsServerCom & "CtaObjFiltro WHERE cCtaContCod in ('" & psCtaContCod & "') and cObjetoCod = '" & psObjetoCod & "'"
    End Select
    Set rs = oconect.CargaRecordSet(Sql)
    If Not rs.EOF Then
        lbExisteFiltro = True
        GetFiltroObjetos = IIf(lbMuestraCta, psCtaContCod, "") & rs!SubCta
    End If
    If GetFiltroObjetos = "" Then
        GetFiltroObjetos = IIf(lbMuestraCta, psCtaContCod, "")
    End If
    Exit Function
GetCtaObjFiltroErr:
    Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones:GetCtaObjFiltro Method")
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
    Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones:ObtieneFechaMov Method")
End Function

Private Sub Class_Initialize()
    Dim oImp As New COMDConstSistema.DCOMImpresoras
    Dim oImpresora As New COMFunciones.FCOMVarImpresion
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As COMConecta.DCOMClasIni
Set oIni = New COMConecta.DCOMClasIni
vsServerCom = oIni.BaseComunes
vsServerPers = oIni.BasePersonas

Set oIni = Nothing
End Sub
Public Function GeneraMovNro(ByVal pdFecha As Date, Optional ByVal psCodAge As String = "07", _
                             Optional ByVal psUser As String = "SIST", Optional psMovNro As String = "", _
                             Optional ByVal pnConnBase As COMConecta.DCOMConecta = Nothing) As String
                             'RIRO20170610 ADD pnConnBase
    On Error GoTo GeneraMovNroErr
    Dim rs As ADODB.Recordset
    Dim oconect As COMConecta.DCOMConecta
    Dim Sql As String
        
    'RIRO 20170610 ***
    If Not pnConnBase Is Nothing Then
        Set oconect = pnConnBase
    Else
        Set oconect = New COMConecta.DCOMConecta
    End If
    
    Set rs = New ADODB.Recordset
    If oconect.AbreConexion = False Then Exit Function
    If psMovNro = "" Or Len(psMovNro) <> 25 Then
       Sql = "sp_GeneraMovNro '" & Format(pdFecha & " " & oconect.GetHoraServer, "mm/dd/yyyy hh:mm:ss") & "','" & Right(psCodAge, 2) & "','" & psUser & "'"
    Else
       Sql = "sp_GeneraMovNro '','','','" & psMovNro & "'"
    End If
    Set rs = oconect.Ejecutar(Sql)
    If Not rs.EOF Then
        GeneraMovNro = rs.Fields(0)
    End If
    rs.Close
    Set rs = Nothing
    
    If pnConnBase Is Nothing Then
        oconect.CierraConexion
        Set oconect = Nothing
    End If
    
    Exit Function
GeneraMovNroErr:
    Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones:GeneraMovNro Method")
End Function
Public Function EmiteCtaContDesc(ByVal psCtaContCod As String) As String
    On Error GoTo EmiteCtaContDescErr
    Dim ssql As String
    Dim rs As ADODB.Recordset
    Dim oconect As COMConecta.DCOMConecta
    
    Set oconect = New COMConecta.DCOMConecta
    Set rs = New ADODB.Recordset
    
    If oconect.AbreConexion = False Then Exit Function
        
    If Len(psCtaContCod) > 4 Then
       ssql = "SELECT RTRIM(ISNULL(c1.cCtaContDesc,'')) + ' ' + RTRIM(c.cCtaContDesc) as cCtaContDesc " _
              & "FROM   " & vsServerCom & "CtaCont c LEFT JOIN " & vsServerCom & "CtaCont c1 ON c1.cCtaContCod = substring(c.cCtaContCod,1," & Len(psCtaContCod) - 2 & ") " _
              & "WHERE  c.cCtaContCod = '" & psCtaContCod & "'"
    Else
       ssql = "SELECT cCtaContDesc FROM " & vsServerCom & "CtaCont WHERE cCtaContCod = '" & psCtaContCod & "'"
    End If
    EmiteCtaContDesc = ""
    Set rs = oconect.CargaRecordSet(ssql)
    If Not rs.BOF And Not rs.EOF Then
        EmiteCtaContDesc = Trim(rs!cCtaContDesc)
    End If
    rs.Close: Set rs = Nothing
    
    oconect.CierraConexion:  Set oconect = Nothing
    Exit Function
EmiteCtaContDescErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:EmiteCtaContDesc Method")
End Function
Public Function GetObjetos(ByVal pnObjetoCod As TpoObjetos, Optional pbObjMiembros As Boolean = True) As ADODB.Recordset
Dim Sql As String
Dim oconect As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Set oconect = New COMConecta.DCOMConecta
If oconect.AbreConexion = False Then Exit Function

If pbObjMiembros Then
    Sql = "Select  cObjetoCod, CONVERT(CHAR(40),cObjetoDesc) as Descripcion, nObjetoNiv  " _
         & " From    " & vsServerCom & "Objeto Where cObjetoCod Like '" & Format(pnObjetoCod, "00") & "__'  ORDER BY cObjetoCod "
Else
    Sql = "Select  cObjetoCod, CONVERT(CHAR(40),cObjetoDesc) as Descripcion, nObjetoNiv  " _
         & " From    " & vsServerCom & "Objeto Where cObjetoCod Like '" & Format(pnObjetoCod, "00") & "%' ORDER BY cObjetoCod "
End If
Set rs = oconect.CargaRecordSet(Sql)
Set GetObjetos = rs
oconect.CierraConexion: Set oconect = Nothing
End Function
Public Function GetSubCtaContFiltro(ByVal psCtaCod As String, ByVal psObjCod As String, ByVal psSubCta As String) As ADODB.Recordset
Dim oCta    As New COMDContabilidad.DCOMCtaCont
Dim oconect As New COMConecta.DCOMConecta
Dim prs     As ADODB.Recordset
Dim psSql   As String
On Error GoTo GetSubCtaContFiltroErr
If oconect.AbreConexion = False Then Exit Function
   Select Case Val(psObjCod)
      Case ObjCMACAgencias
         psSql = "SELECT cAgeCod cObjetoCod, cAgeDesc cObjetoDesc, '1' nObjetoNiv, cSubCtaCod FROM Agencias WHERE cSubCtaCod = LEFT('" & psSubCta & "',LEN(cAgeCod))"
         Set prs = oconect.CargaRecordSet(psSql)
      Case ObjCMACAgenciaArea
         psSql = "SELECT cf.cAreaCod+cf.cAgeCod cObjetoCod, RTRIM(ISNULL(age.cAgeDescripcion,''))+' '+RTRIM(ISNULL(area.cAreaDescripcion,'')) cObjetoDesc , '1' nObjetoNiv, cf.cSubCtaCod FROM CtaAreaAgeFiltro cf LEFT JOIN Areas area ON area.cAreaCod = cf.cAreaCod LEFT JOIN Agencias age ON age.cAgeCod = cf.cAgeCod WHERE cCtaContCod = '" & psCtaCod & "' and cObjetoCod = '" & psObjCod & "' and cf.cSubCtaCod = LEFT('" & psSubCta & "',LEN(cf.cSubCtaCod))"
         Set prs = oconect.CargaRecordSet(psSql)
         If prs.EOF Then
            psSql = "SELECT cf.cAreaCod+cf.cAgeCod cObjetoCod, RTRIM(ISNULL(age.cAgeDescripcion,''))+' '+RTRIM(ISNULL(area.cAreaDescripcion,'')) cObjetoDesc, '1' nObjetoNiv, cf.cSubCtaCod FROM AreaAgencia cf LEFT JOIN Areas area ON area.cAreaCod = cf.cAreaCod LEFT JOIN Agencias age ON age.cAgeCod = cf.cAgeCod WHERE cf.cSubCtaCod = LEFT('" & psSubCta & "',LEN(cf.cSubCtaCod))"
            Set prs = oconect.CargaRecordSet(psSql)
         End If
      Case ObjCMACArea
         psSql = "SELECT cAreaCod cObjetoCod, cAreaDescripcion cObjetoDesc, '1' nObjetoNiv, cSubCtaCod FROM Areas WHERE cSubCtaCod = LEFT('" & psSubCta & "',LEN(cAgeCod))"
         Set prs = oconect.CargaRecordSet(psSql)
      Case ObjEntidadesFinancieras
         If psSubCta = "" Then
            psSql = ""
         Else
         psSql = "SELECT cif.cPersCod + cif.cCtaIfCod cObjetoCod, cCtaIFDesc, '1' nObjetoNiv,  cCtaIFSubCta cSubCtaCod FROM CtaIF cif JOIN CtaIFFiltro ciff ON cif.cPersCod = ciff.cPersCod and cif.cCtaIfCod = ciff.cCtaIfCod " _
               & "WHERE cCtaIFSubCta = LEFT('" & psSubCta & "',LEN(cCtaIFSubCta)) "
         End If
         Set prs = oconect.CargaRecordSet(psSql)
      Case ObjDescomEfectivo
         Set prs = Nothing
      Case objPersona
         Set prs = Nothing
      Case Else
         psSql = "SELECT co.cObjetoCod, o.cObjetoDesc, '1' nObjetoNiv, co.cCtaObjSubCta cSubCtaCod FROM CtaObjFiltro co LEFT JOIN Objeto o ON o.cObjetoCod =  co.cObjetoCod  WHERE cCtaContCod = '" & psCtaCod & "'  AND '" & psSubCta & "' LIKE cCtaObjSubCta + '%' and co.cObjetoCod LIKE '" & psObjCod & "%'"
         Set prs = oconect.CargaRecordSet(psSql)
   End Select
   Set GetSubCtaContFiltro = prs
Set oCta = Nothing
Set oconect = Nothing
Set prs = Nothing
Exit Function
GetSubCtaContFiltroErr:
   Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones: GetSubCtaContFiltro Method")
End Function

Public Sub UbicarEnRegistro(prs As ADODB.Recordset, ByVal pnPos As Variant)
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
    Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones:UbicarenRegistro Method")
End Sub

Public Function PermiteModificarAsiento(psMovNro As String, Optional pbMsg As Boolean = True) As Boolean
Dim ldFecCie As Date
Dim ldFecMov As Date
   On Error GoTo PermiteModificarMovErr
   PermiteModificarAsiento = True
   Dim oGen As New COMDConstSistema.NCOMConstSistema
   ldFecCie = CDate(oGen.LeeConstSistema(gConstSistCierreMensualCont))
   ldFecMov = CDate(ObtieneFechaMov(psMovNro, True))
   Set oGen = Nothing
   
   If ldFecMov <= ldFecCie Then
      PermiteModificarAsiento = False
      'If pbMsg Then
      '   MsgBox "Asiento pertenece a mes ya Cerrado. Imposible realizar operación...!", vbInformation, "Aviso"
      'End If
   End If
   Exit Function
PermiteModificarMovErr:
    Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones:PermiteModificarMov Method")
End Function


Public Function FactorAjuste(ByVal dFecAdq As Date, ByVal dFecRep As Date, Optional nDeci As Integer = 4) As Double
Dim ssql As String
Dim R As New ADODB.Recordset
Dim oFun As COMFunciones.FCOMVarPublicas
Dim FactAFecha As Double
Dim FacRep As Double
Dim oIPM As New COMDContabilidad.DCOMAjusteCont
   Set R = oIPM.CargaIPM(Format(dFecRep, oFun.gsFormatoFecha))
      If R.BOF And R.EOF Then
          R.Close: Set R = Nothing
          FactorAjuste = 0
          Exit Function
      Else
          FacRep = R!nValor
      End If
   R.Close
    
   Set R = oIPM.CargaIPM(Format(dFecAdq, oFun.gsFormatoFecha))
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
Dim Sql As String
Dim oconect As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset


Set oconect = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset

Sql = " SELECT   dAreaSaldoFecha, nAreaSaldoImporte , ISNULL(MOVDIA.NDEBE,0) AS DEBE, " _
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

If oconect.AbreConexion = False Then Exit Function

Set rs = oconect.CargaRecordSet(Sql)
If Not rs.EOF And Not rs.BOF Then
    GetSaldoAreaAgencia = rs!SaldoDia
End If
rs.Close
Set rs = Nothing
End Function

Public Function GetSaldoCtaCont(ByVal psFormatoFecha As String, ByVal pdFecha As Date, _
                                    ByVal psCtaContCod As String) As Currency
Dim Sql As String
Dim oconect As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset

Set oconect = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset

Sql = " SELECT  dCtaSaldoFecha, nCtaSaldoImporte , ISNULL(MOVDIA.NDEBE,0) AS DEBE, " _
    & "         ISNULL(MOVDIA.NHABER,0) AS HABER,  nCtaSaldoImporte + ISNULL(MOVDIA.NDEBE,0) + ISNULL(MOVDIA.NHABER,0) AS SALDODIA " _
    & " FROM    CtaSaldo CS " _
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
    & "         WHERE   CS.cCtaContCod Like '" & psCtaContCod & "' " _
    & "                 AND dCtaSaldoFecha= (   SELECT  MAX(dCtaSaldoFecha) " _
    & "                                         FROM    CtaSaldo CS1 " _
    & "                                         Where   CS1.cCtaContCod = CS.cCtaContCod " _
    & "                                                 AND dCtaSaldoFecha<'" & Format(pdFecha, psFormatoFecha) & " 23:59:59')"
If oconect.AbreConexion = False Then Exit Function
Set rs = oconect.CargaRecordSet(Sql)
If Not rs.EOF And Not rs.BOF Then
    GetSaldoCtaCont = rs!SaldoDia
End If
rs.Close
Set rs = Nothing
End Function
Public Function ExisteMovimiento(psMovNro As String, psOpeCod As String) As Boolean
Dim psSql   As String
Dim oconect As New COMConecta.DCOMConecta
Dim prs     As ADODB.Recordset
ExisteMovimiento = True
   psSql = "SELECT m.cMovNro FROM Mov m WHERE m.cMovNro LIKE '" & psMovNro & "%' and m.nMovEstado = '" & gMovEstContabMovContable & "' and not m.nMovFlag in ('" & gMovFlagEliminado & "','" & gMovFlagModificado & "') and cOpeCod = '" & psOpeCod & "'"
   oconect.AbreConexion
   Set prs = oconect.CargaRecordSet(psSql)
   If prs.EOF Then
      ExisteMovimiento = False
   End If
   Set oconect = Nothing
   prs.Close: Set prs = Nothing
End Function

Public Function GetObjetosArbol(ByVal pnObjetoCod As String, ByVal psObjFiltro As String, ByVal pnNivel As Integer) As ADODB.Recordset
Dim Sql As String
Dim oconect As New COMConecta.DCOMConecta
Dim rs As New ADODB.Recordset
   If oconect.AbreConexion = False Then Exit Function
   Sql = vsServerCom & "spGetTreeObj '" & Format(pnObjetoCod, "00") & "', " & pnNivel & ", '" & psObjFiltro & "'"
   Set rs = oconect.CargaRecordSet(Sql)
   Set GetObjetosArbol = rs
   oconect.CierraConexion: Set oconect = Nothing
Set rs = Nothing
End Function

Public Function GetLibroDiario(ByVal dFecIni As Date, ByVal dFecFin As Date, psAge As String, psOpe As String) As ADODB.Recordset
Dim sAgeCond As String, sOpeCond As String
Dim Sql As String
Dim oconect As New COMConecta.DCOMConecta
 On Error GoTo GetLibroDiarioErr
 
 If oconect.AbreConexion = False Then Exit Function
    sAgeCond = "": sOpeCond = ""
    If psAge <> "" Then
       sAgeCond = " And SubString(M.cMovNro,18,2) = '" & Right(psAge, 2) & "' "
    End If
    If psOpe <> "" Then
       sOpeCond = " And O.cOpeGruCod = '" & Trim(psOpe) & "' "
    End If

    Sql = "Select M1.nMovCorrela, M.cMovNro, O.cOpeGruCod, MC.cCtaContCod, C.cCtaContDesc, MC.nMovImporte, og.cOpeGruDesc " _
        & "From Mov M JOIN MovCont M1 ON M1.nMovNro = M.nMovNro " _
        & "           Join MovCta MC On M.nMovNro = MC.nMovNro " _
        & "           Join " & vsServerCom & "CtaCont C On MC.cCtaContCod = C.cCtaContCod " _
        & "           Join " & vsServerCom & "OpeTpo  O On M.cOpeCod = O.cOpeCod " _
        & "           Join " & vsServerCom & "OpeGrupo Og ON og.cOpeGruCod = o.cOpeGruCod " _
        & "Where M.nMovEstado = " & gMovEstContabMovContable & " and not M.nMovFlag IN (" & gMovFlagEliminado & "," & gMovFlagModificado & ") and substring(M.cMovNro,1,8) between " _
        & "      '" & Format(dFecIni, "yyyymmdd") & "' And '" & Format(dFecFin, "yyyymmdd") _
        & "'     AND mc.cCtaContCod <> '' " & sAgeCond & sOpeCond _
        & "ORDER BY M1.nMovCorrela, substring(M.cMovNro,1,8), substring(M.cMovNro,18,2), O.cOpeGruCod, M.cMovNro, Mc.nMovItem "
    Set GetLibroDiario = oconect.CargaRecordSet(Sql)
    oconect.CierraConexion: Set oconect = Nothing
Exit Function
GetLibroDiarioErr:
   Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones: GetLibroDiario Method")
End Function

Public Function ValidaDocumento(nDocTpo As TpoDoc, sDocNro As String, sCodProv As String) As Boolean
Dim ssql As String
Dim prs  As ADODB.Recordset
Dim oconect As New COMConecta.DCOMConecta

ValidaDocumento = True
oconect.AbreConexion

ssql = "SELECT a.nMovNro FROM MovGasto a JOIN MovDoc b ON b.nMovNro = a.nMovNro JOIN Mov m ON m.nMovNro = a.nMovNro " _
     & "WHERE  not m.nMovFlag IN ('" & gMovFlagEliminado & "','" & gMovFlagModificado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "') and a.cPersCod = '" & sCodProv & "' and b.nDocTpo = '" & nDocTpo & "' and b.cDocNro = '" & sDocNro & "'"
Set prs = oconect.CargaRecordSet(ssql)
If Not prs.EOF Then
   ValidaDocumento = False
   prs.Close
   Err.Raise 50001, "nContFunciones:ValidaDocumento", "Documento de Proveedor ya registrado..."
End If
 prs.Close
End Function

Public Function IncrementaMovNro(ByVal cMovnro As String) As String
Dim cTexto1 As String
Dim cTexto2 As String
Dim cTexto3 As String
Dim cNumero As Integer
Dim nNumero As Integer
Dim sNumero As String
cTexto1 = Mid(cMovnro, 1, 19)
cTexto2 = Mid(cMovnro, 20, 2)
cTexto3 = Mid(cMovnro, 22, 4)
nNumero = Val(cTexto2) + 1
sNumero = Right("00" & Trim(Str(nNumero)), 2)
IncrementaMovNro = cTexto1 + sNumero + cTexto3
End Function
'*** BRGO 20110511 *******************
Public Sub InsertarBienAdjudicado(ByVal sCuenta As String, ByVal nNumPartida As String, ByVal sTipoBien As String, _
    ByVal sDescripcion As String, ByVal nTipoAdjud As String, ByVal nValAdjud As Double, ByVal nValComercial As Double, _
    ByVal nValTasacion As Double, ByVal dFecAdjud As Date, ByVal dFecTasacion As Date, ByVal nMoneda As Integer, _
    ByVal nCapital As Double, ByVal nIntereses As Double, ByVal nGastos As Double, ByVal nEstado As Integer, ByVal sMovNro As String)
    
    Dim oconect As New COMConecta.DCOMConecta
    Dim ssql As String
    On Error GoTo RegBienAdjudErr
    oconect.AbreConexion
    
    ssql = "stp_ins_ColocBienesAdjudicados '" & sCuenta & "','" & Left(sTipoBien, 3) & "','" & sDescripcion & "','" & nNumPartida & "'," & nTipoAdjud & ", " & nValAdjud & ", " & nValComercial & ",'" & Format(dFecAdjud, "MM/dd/yyyy") & "', " & nValTasacion & ",'" & Format(dFecTasacion, "MM/dd/yyyy") & "'," & nCapital & "," & nMoneda & "," & nIntereses & "," & nGastos & "," & nEstado & ",'" & sMovNro & "'"
    oconect.CargaRecordSet (ssql)
    oconect.CierraConexion: Set oconect = Nothing
    Exit Sub
RegBienAdjudErr:
   Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones: GetLibroDiario Method")
End Sub
Public Sub ModificarBienAdjudicado(ByVal nBienAdjudCod As Integer, ByVal sCuenta As String, ByVal nNumPartida As String, ByVal sTipoBien As String, _
    ByVal sDescripcion As String, ByVal nTipoAdjud As String, ByVal nValAdjud As Double, ByVal nValComercial As Double, _
    ByVal nValTasacion As Double, ByVal dFecAdjud As Date, ByVal dFecTasacion As Date, ByVal nMoneda As Integer, _
    ByVal nCapital As Double, ByVal nIntereses As Double, ByVal nGastos As Double, ByVal nEstado As Integer, ByVal sMovNro As String)
    
    Dim oconect As New COMConecta.DCOMConecta
    Dim ssql As String
    On Error GoTo RegBienAdjudErr
    oconect.AbreConexion
    
    ssql = "stp_upd_ColocBienesAdjudicadosDatos " & nBienAdjudCod & ",'" & sCuenta & "','" & Left(sTipoBien, 3) & "','" & sDescripcion & "','" & nNumPartida & "'," & nTipoAdjud & ", " & nValAdjud & ", " & nValComercial & ",'" & Format(dFecAdjud, "MM/dd/yyyy") & "', " & nValTasacion & ",'" & Format(dFecTasacion, "MM/dd/yyyy") & "'," & nCapital & "," & nMoneda & "," & nIntereses & "," & nGastos & "," & nEstado & ",'" & sMovNro & "'"
    oconect.CargaRecordSet (ssql)
    oconect.CierraConexion: Set oconect = Nothing
    Exit Sub
RegBienAdjudErr:
   Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones: GetLibroDiario Method")
End Sub
Public Sub InsertarVentaBienAdjudicado(ByVal NumAdj As Integer, dFecVenta As Date, nValVenta As Double)
    
    Dim oconect As New COMConecta.DCOMConecta
    Dim ssql As String
    On Error GoTo RegBienAdjudErr
    oconect.AbreConexion
    
    ssql = "stp_upd_ColocBienesAdjudicados " & NumAdj & ",'" & Format(dFecVenta, "MM/dd/yyyy") & "'," & nValVenta
    oconect.CargaRecordSet (ssql)
    oconect.CierraConexion: Set oconect = Nothing
    Exit Sub
RegBienAdjudErr:
   Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones: GetLibroDiario Method")
End Sub
Public Sub CambiaEstadoBienAdjudicado(ByVal NumAdj As Integer, ByVal nEstado As Integer)
    
    Dim oconect As New COMConecta.DCOMConecta
    Dim ssql As String
    On Error GoTo RegBienAdjudErr
    oconect.AbreConexion
    
    ssql = "stp_upd_CambiaEstadoBienAdjudicado " & NumAdj & "," & nEstado
    oconect.CargaRecordSet (ssql)
    oconect.CierraConexion: Set oconect = Nothing
    Exit Sub
RegBienAdjudErr:
   Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones: GetLibroDiario Method")
End Sub

Public Function ObtenerDatosBienAdjudicado(ByVal NumAdj As Integer) As ADODB.Recordset
    Dim oconect As New COMConecta.DCOMConecta
    Dim ssql As String
    On Error GoTo RegBienAdjudErr
    oconect.AbreConexion
    
    ssql = "stp_sel_ObtenerBienAdjudicado " & NumAdj
    Set ObtenerDatosBienAdjudicado = oconect.CargaRecordSet(ssql)
    oconect.CierraConexion: Set oconect = Nothing
    Exit Function
RegBienAdjudErr:
   Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones: GetLibroDiario Method")
End Function
Public Function ObtenerListaBienesAdjudicado(ByVal psFecDel As String, ByVal psFecAl As String, nTipo As Integer) As ADODB.Recordset
    'ByVal sMes As String, ByVal sAnio As String, nTipo As Integer) As ADODB.Recordset

    Dim oconect As New COMConecta.DCOMConecta
    Dim ssql As String
    On Error GoTo RegBienAdjudErr
    oconect.AbreConexion
    
    '*** PEAC 20130617
    'sSQL = "stp_sel_ObtenerListaBienesAdjudicados '" & sMes & "','" & sAnio & "'," & nTipo
    ssql = "stp_sel_ObtenerListaBienesAdjudicados '" & psFecDel & "','" & psFecAl & "'," & nTipo
    '*** FIN PEAC
    
    Set ObtenerListaBienesAdjudicado = oconect.CargaRecordSet(ssql)
    oconect.CierraConexion: Set oconect = Nothing
    Exit Function
RegBienAdjudErr:
   Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones: GetLibroDiario Method")
End Function
'***** End BRGO ***************************

'JACA 20110809*****************************************************************

Public Sub InsertarBienEmbargoAdjudicado(ByVal sCuenta As String, ByVal nNumPartida As String, ByVal sTipoBien As String, _
    ByVal sDescripcion As String, ByVal nTipoAdjud As String, ByVal nValAdjud As Double, ByVal nValComercial As Double, _
    ByVal nValTasacion As Double, ByVal dFecAdjud As Date, ByVal nMoneda As Integer, _
    ByVal nCapital As Double, ByVal nIntereses As Double, ByVal nGastos As Double, ByVal nEstado As Integer, _
    Optional dFecTasacion As String = "")
    
    Dim oconect As New COMConecta.DCOMConecta
    Dim ssql As String
    On Error GoTo RegBienAdjudErr
    oconect.AbreConexion
    
    ssql = "stp_ins_ColocBienesEmbargoAdjudicados '" & sCuenta & "','" & Left(sTipoBien, 3) & "','" & sDescripcion & "','" & nNumPartida & "'," & nTipoAdjud & ", " & nValAdjud & ", " & nValComercial & ",'" & Format(dFecAdjud, "MM/dd/yyyy") & "', " & nValTasacion & ",'" & IIf(dFecTasacion <> "", Format(CDate(dFecTasacion), "MM/dd/yyyy"), "") & "'," & nCapital & "," & nMoneda & "," & nIntereses & "," & nGastos & "," & nEstado
    oconect.CargaRecordSet (ssql)
    oconect.CierraConexion: Set oconect = Nothing
    Exit Sub
RegBienAdjudErr:
   Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones: GetLibroDiario Method")
End Sub
'JACA END********************************************************************

'WIOR 20140114 **************************************************************
Public Function PatrimonioEfecAjustInfl(ByVal lsAnio As String, ByVal lsMes As String) As Double
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rs = CargaDatosPatrimonioEfec(lsAnio, lsMes, 0)
    
   PatrimonioEfecAjustInfl = GeneraPatrimonioEfecAjustInfl(CInt(lsAnio), CInt(lsMes), rs)
   'RSClose rs
   Set rs = Nothing
End Function

Public Function CargaDatosPatrimonioEfec(ByVal lsAnio As String, ByVal lsMes As String, ByVal lsTipoBalance As String) As ADODB.Recordset
    Dim oRep As COMDContabilidad.DCOMCtaCont
    Set oRep = New COMDContabilidad.DCOMCtaCont
    Set CargaDatosPatrimonioEfec = oRep.Reporte3PatrimonioEfec(lsAnio, lsMes, lsTipoBalance)
    Set oRep = Nothing
End Function

Public Function GeneraPatrimonioEfecAjustInfl(ByVal lnAnio_ As Integer, ByVal lnMes_ As Integer, ByVal rsCta As ADODB.Recordset) As Double
Dim nPatrimonioEfecNivel1 As Double
Dim nPatrimonioEfecNivel2 As Double
Dim nPatrimonioEfecNivel3 As Double
Dim nSaldo85  As Currency

On Error GoTo GeneraExcelErr

    nPatrimonioEfecNivel1 = 0
    nPatrimonioEfecNivel2 = 0
    nPatrimonioEfecNivel3 = 0
   
    Do While Not rsCta.EOF
        DoEvents
        

        If rsCta!cCtaContCod = "310101" Or rsCta!cCtaContCod = "310102" Or rsCta!cCtaContCod = "3201" Or rsCta!cCtaContCod = "320201" Or rsCta!cCtaContCod = "320202" Or rsCta!cCtaContCod = "320301" Or _
            rsCta!cCtaContCod = "320302" Or rsCta!cCtaContCod = "3301" Or rsCta!cCtaContCod = "3302" Or rsCta!cCtaContCod = "330301" Or rsCta!cCtaContCod = "380101" Or rsCta!cCtaContCod = "390101" Or _
            rsCta!cCtaContCod = "380104" Or rsCta!cCtaContCod = "280202" Or rsCta!cCtaContCod = "280203" Or rsCta!cCtaContCod = "280204" Or rsCta!cCtaContCod = "280601" Or rsCta!cCtaContCod = "280602" Or _
            rsCta!cCtaContCod = "280603" Or rsCta!cCtaContCod = "380201" Or rsCta!cCtaContCod = "380203" Or rsCta!cCtaContCod = "380204" Or rsCta!cCtaContCod = "380205" Or rsCta!cCtaContCod = "3902" Or _
            rsCta!cCtaContCod = "130308" Or rsCta!cCtaContCod = "13080308" Or rsCta!cCtaContCod = "13090308" Or rsCta!cCtaContCod = "13040801" Or rsCta!cCtaContCod = "1308040801" Or rsCta!cCtaContCod = "1309040801" Then
            
            nPatrimonioEfecNivel1 = nPatrimonioEfecNivel1 + rsCta!nSaldoFinImporte
        End If
        
        If rsCta!cCtaContCod = "330302" Or rsCta!cCtaContCod = "310103" Or rsCta!cCtaContCod = "310104" Or rsCta!cCtaContCod = "310109" Or rsCta!cCtaContCod = "320203" Or rsCta!cCtaContCod = "320204" Or _
            rsCta!cCtaContCod = "320209" Or rsCta!cCtaContCod = "320303" Or rsCta!cCtaContCod = "320304" Or rsCta!cCtaContCod = "320309" Or rsCta!cCtaContCod = "260202010201" Or rsCta!cCtaContCod = "2602020202" Or _
            rsCta!cCtaContCod = "2602020302" Or rsCta!cCtaContCod = "2603020102" Or rsCta!cCtaContCod = "2603020202" Or rsCta!cCtaContCod = "2603020302" Or rsCta!cCtaContCod = "2604020102" Or rsCta!cCtaContCod = "2604020202" Or _
            rsCta!cCtaContCod = "2604020302" Or rsCta!cCtaContCod = "2605010102" Or rsCta!cCtaContCod = "2605010202" Or rsCta!cCtaContCod = "2605010302" Or rsCta!cCtaContCod = "2606020102" Or rsCta!cCtaContCod = "2606020202" Or _
            rsCta!cCtaContCod = "2606020302" Or rsCta!cCtaContCod = "260701010201" Or rsCta!cCtaContCod = "2607010202" Or rsCta!cCtaContCod = "2607010302" Or rsCta!cCtaContCod = "280202" Or rsCta!cCtaContCod = "280203" Or _
            rsCta!cCtaContCod = "280204" Or rsCta!cCtaContCod = "280601" Or rsCta!cCtaContCod = "280602" Or rsCta!cCtaContCod = "280603" Or rsCta!cCtaContCod = "14090202" Or rsCta!cCtaContCod = "14090203" Or _
            rsCta!cCtaContCod = "14090206" Or rsCta!cCtaContCod = "14090302" Or rsCta!cCtaContCod = "14090303" Or rsCta!cCtaContCod = "14090306" Or rsCta!cCtaContCod = "14090402" Or rsCta!cCtaContCod = "14090403" Or _
            rsCta!cCtaContCod = "14090502" Or rsCta!cCtaContCod = "14090503" Or rsCta!cCtaContCod = "14090602" Or rsCta!cCtaContCod = "14090603" Or rsCta!cCtaContCod = "14090702" Or rsCta!cCtaContCod = "14090703" Or _
            rsCta!cCtaContCod = "14090802" Or rsCta!cCtaContCod = "14090803" Or rsCta!cCtaContCod = "14090902" Or rsCta!cCtaContCod = "14090903" Or rsCta!cCtaContCod = "14091002" Or rsCta!cCtaContCod = "14091003" Or _
            rsCta!cCtaContCod = "14091102" Or rsCta!cCtaContCod = "14091103" Or rsCta!cCtaContCod = "14091202" Or rsCta!cCtaContCod = "14091203" Or rsCta!cCtaContCod = "14091302" Or rsCta!cCtaContCod = "14091303" Or _
            rsCta!cCtaContCod = "14091306" Or rsCta!cCtaContCod = "1509040101" Or rsCta!cCtaContCod = "270102" Or rsCta!cCtaContCod = "270203" Or rsCta!cCtaContCod = "270204" Or rsCta!cCtaContCod = "130308" Or _
            rsCta!cCtaContCod = "13080308" Or rsCta!cCtaContCod = "13090308" Or rsCta!cCtaContCod = "13040801" Or rsCta!cCtaContCod = "1308040801" Or rsCta!cCtaContCod = "1309040801" Then
            
            nPatrimonioEfecNivel2 = nPatrimonioEfecNivel2 + rsCta!nSaldoFinImporte
        End If
        
        If rsCta!cCtaContCod = "310104" Or rsCta!cCtaContCod = "310109" Or rsCta!cCtaContCod = "320204" Or rsCta!cCtaContCod = "320209" Or rsCta!cCtaContCod = "320304" Or rsCta!cCtaContCod = "320309" Or _
            rsCta!cCtaContCod = "280203" Or rsCta!cCtaContCod = "280204" Or rsCta!cCtaContCod = "280601" Or rsCta!cCtaContCod = "280602" Or rsCta!cCtaContCod = "280603" Or rsCta!cCtaContCod = "130308" Or _
            rsCta!cCtaContCod = "13080308" Or rsCta!cCtaContCod = "13090308" Or rsCta!cCtaContCod = "13040801" Or rsCta!cCtaContCod = "1308040801" Or rsCta!cCtaContCod = "1309040801" Then
            
            nPatrimonioEfecNivel3 = nPatrimonioEfecNivel3 + rsCta!nSaldoFinImporte
        End If
        
        If rsCta!cCtaContCod = "190407" Then
            nSaldo85 = nSaldo85 + rsCta!nSaldoFinImporte
        End If
        
        If rsCta!cCtaContCod = "19040907" Then
            nSaldo85 = nSaldo85 - rsCta!nSaldoFinImporte
        End If
        
        If rsCta!cCtaContCod = "17" Then
            nSaldo85 = nSaldo85 + rsCta!nSaldoFinImporte
        End If

       rsCta.MoveNext
        If rsCta.EOF Then
           Exit Do
        End If
    Loop

    GeneraPatrimonioEfecAjustInfl = nSaldo85 + nPatrimonioEfecNivel1 + nPatrimonioEfecNivel2 + nPatrimonioEfecNivel3

Exit Function
GeneraExcelErr:
    MsgBox Err.Description, vbInformation, "Aviso"
    Exit Function
End Function

'WIOR FIN *******************************************************************
'PASI20161103 ERS0572016*********************************************************
Public Function EsGrupoHabilxProDepreAdj(ByVal psGrupoUsu As String) As Boolean
    Dim oconect As New COMConecta.DCOMConecta
    Dim ssql As String
    Dim rs As ADODB.Recordset
    oconect.AbreConexion
    ssql = "stp_sel_EsGrupoHabilxProDepreAdj '" & Replace(psGrupoUsu, "'", "") & "'"
    Set rs = oconect.CargaRecordSet(ssql)
    If Not rs.EOF And Not rs.BOF Then EsGrupoHabilxProDepreAdj = IIf(rs!nValor > 0, True, False)
End Function
Public Function RegistraProrrogaDepreAdjudicado(ByVal pnBienAdjCod As Integer, ByVal psMovNro As String) As Boolean
    Dim oconect As New COMConecta.DCOMConecta
    Dim ssql As String
    On Error GoTo RegProDepreAdj
        RegistraProrrogaDepreAdjudicado = False
        oconect.AbreConexion
        ssql = "stp_ins_RegistraProrrogaDepreAdjudicado " & pnBienAdjCod & ",'" & psMovNro & "'"
        oconect.Ejecutar ssql
        oconect.CierraConexion
        RegistraProrrogaDepreAdjudicado = True
        Exit Function
RegProDepreAdj:
    RegistraProrrogaDepreAdjudicado = False
   Call oError.RaiseError(oError.MyUnhandledError, "NCOMContFunciones: RegistraProrrogaDepreAdjudicado Function")
End Function
Public Function ExisteProrrogaDepreAdjudicado(ByVal pnBienAdjCod As Integer) As Boolean
    Dim oconect As New COMConecta.DCOMConecta
    Dim ssql As String
    Dim rs As ADODB.Recordset
    oconect.AbreConexion
    ssql = "stp_sel_ExisteProrrogaDepreAdjudicado " & pnBienAdjCod
    Set rs = oconect.CargaRecordSet(ssql)
    If Not rs.EOF And Not rs.BOF Then ExisteProrrogaDepreAdjudicado = IIf(rs!nValor > 0, True, False)
End Function
'PASI END************************************************************************


