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


Public Function GeneraDocNro(psDocTpo As TpoDoc, psMoneda As Moneda, Optional psDocSerie As String = "", Optional psAgencia As String = "") As String
    On Error GoTo GeneraMovNroErr
    Dim sSql As String
    Dim RS As New ADODB.Recordset
    Dim oConect As DConecta
    Dim oCaja As nCajaGeneral
    
    Set oCaja = New nCajaGeneral
    Set oConect = New DConecta
    If oConect.AbreConexion = False Then Exit Function
      
        If psDocSerie <> "" Then
            sSql = "SELECT max(cDocNro) AS cDocCorrela FROM Movdoc md JOIN Mov m ON m.nMovNro = md.nMovNro " _
                & "WHERE   m.nMovFlag = ('" & gMovFlagVigente & "') and nDocTpo = '" & Format(psDocTpo, "00") & "' and substring(cDocNro,1," & Len(psDocSerie) & ") = '" & psDocSerie & "'"
        Else
            sSql = "SELECT  cDocNro AS cDocCorrela , md.nMovnro " _
                & " FROM    MovDoc MD JOIN MOV M on M.nMovNro = md.nMovNro  " _
                & " WHERE   nDocTpo = '" & Format(psDocTpo, "00") & "'  " _
                & "         and md.nmovnro = (  Select Max(MD1.nmovnro)  " _
                & "                             From MovDoc MD1 JOIN Mov M1 on M1.nMovNro= MD1.nMovnro " _
                & "                             WHERE MD1.nDocTpo = '" & Format(psDocTpo, "00") & "' and Substring(M.cOpeCod,3,1) ='" & psMoneda & "'  )  "
                
                'and  m.nMovFlag = ('" & gMovFlagVigente & "')
                'and  m1.nMovFlag = ('" & gMovFlagVigente & "')
        End If
        Set RS = oConect.CargaRecordSet(sSql)
        If Not RS.EOF And Not RS.BOF Then
            If Not IsNull(RS!cDocCorrela) Then
                If psDocSerie <> "" Then
                   GeneraDocNro = psDocSerie & "-" & Format(Val(Mid(RS!cDocCorrela, Len(psDocSerie) + 2, 20)) + 1, String(8, "0"))
                Else
                   If Mid(RS!cDocCorrela, 4, 1) = "-" Then
                      GeneraDocNro = Mid(RS!cDocCorrela, 1, 3) & "-" & Format(Val(Mid(RS!cDocCorrela, 5, 20)) + 1, String(8, "0"))
                   Else
                      GeneraDocNro = Format(Val(RS!cDocCorrela) + 1, String(8, "0"))
                   End If
                End If
        End If
        
If psDocTpo > -1 Then
   If psDocSerie <> "" Then
      sSql = "SELECT max(cDocNro) AS cDocCorrela FROM  movdoc md JOIN Mov m ON m.nMovNro = md.nMovNro " _
           & "WHERE nDocTpo = '" & psDocTpo & "' and substring(cDocNro,1," & Len(psDocSerie) & ") = '" & psDocSerie & "' " & IIf(psAgencia = "", "", " and SubString(m.cMovNro,18,2) = '" & Right(psAgencia, 2) & "' ")
   Else
      If psMoneda = "" Then
         sSql = "SELECT max(cDocNro) AS cDocCorrela FROM  movdoc md JOIN Mov m ON m.nMovNro = md.nMovNro " _
               & "WHERE nDocTpo = '" & psDocTpo & "' " & IIf(psAgencia = "", "", " and SubString(m.cMovNro,18,2) = '" & Right(psAgencia, 2) & "' ")
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
   Set RS = oConect.CargaRecordSet(sSql)
   If RS.EOF Then
      GeneraDocNro = ""
   Else
      If Not IsNull(RS!cDocCorrela) Then
         If psDocSerie <> "" Then
            GeneraDocNro = psDocSerie & "-" & Format(Val(Mid(RS!cDocCorrela, Len(psDocSerie) + 2, 20)) + 1, String(8, "0"))
         Else
            If Mid(RS!cDocCorrela, 4, 1) = "-" Then
               GeneraDocNro = Mid(RS!cDocCorrela, 1, 3) & "-" & Format(Val(Mid(RS!cDocCorrela, 5, 20)) + 1, String(8, "0"))
            Else
               GeneraDocNro = Format(Val(RS!cDocCorrela) + 1, String(8, "0"))
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
    oConect.CierraConexion
End If

GeneraMovNroErr:
    Err.Raise 50001, "NContFunciones: GeneraDocNro", "Especificar Tipo de Documento para Generar Número"
    
End Function
'Public Function GetFiltroObjetos(ByVal pnTipoObj As TpoObjetos, ByVal psCtaContCod As String, ByVal psObjetoCod As String, Optional lbMuestraCta As Boolean = True, Optional lbExisteFiltro As Boolean = False) As String

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
    Dim RS   As New ADODB.Recordset
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
                sql = "select cPersCod, cSubCtaContcod SubCta FROM InstitucionFinanc where cPersCod = '" & psObjetoCod & "' "
            Else
                sql = "Select cPersCod, cCtaIfCod, cCtaContCod, cCtaIFSubCta as SubCta From " & vsServerCom & "CtaIFFiltro WHERE  cPersCod = '" & lsPersCtaIF & "' and cIFTpo='" & lsTpoIf & "' AND cCtaIfCod='" & lsCtaIf & "' AND cCtaContCod in ('" & psCtaContCod & "')"
            End If
        Case Else
            sql = "SELECT cCtaObjSubCta as SubCta FROM " & vsServerCom & "CtaObjFiltro WHERE cCtaContCod in ('" & psCtaContCod & "') and cObjetoCod = '" & psObjetoCod & "'"
    End Select
    Set RS = oConect.CargaRecordSet(sql)
    If Not RS.EOF Then
        lbExisteFiltro = True
        GetFiltroObjetos = IIf(lbMuestraCta, psCtaContCod, "") & RS!SubCta
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
Dim oIni As COMConecta.DCOMClasIni
Set oIni = New COMConecta.DCOMClasIni
vsServerCom = oIni.BaseComunes
vsServerPers = oIni.BasePersonas
Set oIni = Nothing
End Sub
Public Function GeneraMovNro(ByVal pdFecha As Date, Optional ByVal psCodAge As String = "07", Optional ByVal psUser As String = "SIST", Optional psMovNro As String = "") As String
    On Error GoTo GeneraMovNroErr
    Dim RS As ADODB.Recordset
    Dim oConect As DConecta
    Dim sql As String
    Set oConect = New DConecta
    Set RS = New ADODB.Recordset
    If oConect.AbreConexion = False Then Exit Function
    If psMovNro = "" Or Len(psMovNro) <> 25 Then
       sql = "sp_GeneraMovNro '" & Format(pdFecha & " " & oConect.GetHoraServer, "mm/dd/yyyy hh:mm:ss") & "','" & Right(psCodAge, 2) & "','" & psUser & "'"
    Else
       sql = "sp_GeneraMovNro '','','','" & psMovNro & "'"
    End If
    Set RS = oConect.Ejecutar(sql)
    If Not RS.EOF Then
        GeneraMovNro = RS.Fields(0)
    End If
    RS.Close
    Set RS = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Function
GeneraMovNroErr:
    Call RaiseError(MyUnhandledError, "NContFunciones:GeneraMovNro Method")
End Function
Public Function EmiteCtaContDesc(ByVal psCtaContCod As String) As String
    On Error GoTo EmiteCtaContDescErr
    Dim sSql As String
    Dim RS As ADODB.Recordset
    Dim oConect As DConecta
    
    Set oConect = New DConecta
    Set RS = New ADODB.Recordset
    
    If oConect.AbreConexion = False Then Exit Function
        
    If Len(psCtaContCod) > 4 Then
       sSql = "SELECT RTRIM(ISNULL(c1.cCtaContDesc,'')) + ' ' + RTRIM(c.cCtaContDesc) as cCtaContDesc " _
              & "FROM   " & vsServerCom & "CtaCont c LEFT JOIN " & vsServerCom & "CtaCont c1 ON c1.cCtaContCod = substring(c.cCtaContCod,1," & Len(psCtaContCod) - 2 & ") " _
              & "WHERE  c.cCtaContCod = '" & psCtaContCod & "'"
    Else
       sSql = "SELECT cCtaContDesc FROM " & vsServerCom & "CtaCont WHERE cCtaContCod = '" & psCtaContCod & "'"
    End If
    EmiteCtaContDesc = ""
    Set RS = oConect.CargaRecordSet(sSql)
    If Not RS.BOF And Not RS.EOF Then
        EmiteCtaContDesc = Trim(RS!cCtaContDesc)
    End If
    RS.Close: Set RS = Nothing
    
    oConect.CierraConexion:  Set oConect = Nothing
    Exit Function
EmiteCtaContDescErr:
    Call RaiseError(MyUnhandledError, "DMov:EmiteCtaContDesc Method")
End Function
Public Function GetObjetos(ByVal pnObjetoCod As TpoObjetos, Optional pbObjMiembros As Boolean = True) As ADODB.Recordset
Dim sql As String
Dim oConect As DConecta
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

If pbObjMiembros Then
    sql = "Select  cObjetoCod, CONVERT(CHAR(40),cObjetoDesc) as Descripcion, nObjetoNiv  " _
         & " From    " & vsServerCom & "Objeto Where cObjetoCod Like '" & Format(pnObjetoCod, "00") & "__'  ORDER BY cObjetoCod "
Else
    sql = "Select  cObjetoCod, CONVERT(CHAR(40),cObjetoDesc) as Descripcion, nObjetoNiv  " _
         & " From    " & vsServerCom & "Objeto Where cObjetoCod Like '" & Format(pnObjetoCod, "00") & "%' ORDER BY cObjetoCod "
End If
Set RS = oConect.CargaRecordSet(sql)
Set GetObjetos = RS
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function GetSubCtaContFiltro(ByVal psCtaCod As String, ByVal psObjCod As String, ByVal psSubCta As String) As ADODB.Recordset
Dim oCta    As New DCtaCont
Dim oConect As New DConecta
Dim prs     As ADODB.Recordset
Dim psSql   As String
On Error GoTo GetSubCtaContFiltroErr
If oConect.AbreConexion = False Then Exit Function
   Select Case Val(psObjCod)
      Case ObjCMACAgencias
         psSql = "SELECT cAgeCod cObjetoCod, cAgeDesc cObjetoDesc, '1' nObjetoNiv, cSubCtaCod FROM Agencias WHERE cSubCtaCod = LEFT('" & psSubCta & "',LEN(cAgeCod))"
         Set prs = oConect.CargaRecordSet(psSql)
      Case ObjCMACAgenciaArea
         psSql = "SELECT cf.cAreaCod+cf.cAgeCod cObjetoCod, RTRIM(ISNULL(age.cAgeDescripcion,''))+' '+RTRIM(ISNULL(area.cAreaDescripcion,'')) cObjetoDesc , '1' nObjetoNiv, cf.cSubCtaCod FROM CtaAreaAgeFiltro cf LEFT JOIN Areas area ON area.cAreaCod = cf.cAreaCod LEFT JOIN Agencias age ON age.cAgeCod = cf.cAgeCod WHERE cCtaContCod = '" & psCtaCod & "' and cObjetoCod = '" & psObjCod & "' and cf.cSubCtaCod = LEFT('" & psSubCta & "',LEN(cf.cSubCtaCod))"
         Set prs = oConect.CargaRecordSet(psSql)
         If prs.EOF Then
            psSql = "SELECT cf.cAreaCod+cf.cAgeCod cObjetoCod, RTRIM(ISNULL(age.cAgeDescripcion,''))+' '+RTRIM(ISNULL(area.cAreaDescripcion,'')) cObjetoDesc, '1' nObjetoNiv, cf.cSubCtaCod FROM AreaAgencia cf LEFT JOIN Areas area ON area.cAreaCod = cf.cAreaCod LEFT JOIN Agencias age ON age.cAgeCod = cf.cAgeCod WHERE cf.cSubCtaCod = LEFT('" & psSubCta & "',LEN(cf.cSubCtaCod))"
            Set prs = oConect.CargaRecordSet(psSql)
         End If
      Case ObjCMACArea
         psSql = "SELECT cAreaCod cObjetoCod, cAreaDescripcion cObjetoDesc, '1' nObjetoNiv, cSubCtaCod FROM Areas WHERE cSubCtaCod = LEFT('" & psSubCta & "',LEN(cAgeCod))"
         Set prs = oConect.CargaRecordSet(psSql)
      Case ObjEntidadesFinancieras
         If psSubCta = "" Then
            psSql = ""
         Else
         psSql = "SELECT cif.cPersCod + cif.cCtaIfCod cObjetoCod, cCtaIFDesc, '1' nObjetoNiv,  cCtaIFSubCta cSubCtaCod FROM CtaIF cif JOIN CtaIFFiltro ciff ON cif.cPersCod = ciff.cPersCod and cif.cCtaIfCod = ciff.cCtaIfCod " _
               & "WHERE cCtaIFSubCta = LEFT('" & psSubCta & "',LEN(cCtaIFSubCta)) "
         End If
         Set prs = oConect.CargaRecordSet(psSql)
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

Public Function PermiteModificarAsiento(psMovNro As String, Optional pbMsg As Boolean = True) As Boolean
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
                                    ByVal psCtaContCod As String, ByVal psAreaCod As String, ByVal psAgecod As String) As Currency
Dim sql As String
Dim oConect As DConecta
Dim RS As ADODB.Recordset


Set oConect = New DConecta
Set RS = New ADODB.Recordset

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
    & " WHERE   CS.cCtaContCod Like '" & psCtaContCod & "' AND CS.cAreaCod ='" & psAreaCod & "' AND CS.cAgeCod='" & psAgecod & "' " _
    & "         AND dAreaSaldoFecha= (  SELECT  MAX(dAreaSaldoFecha) " _
    & "                                 FROM    CtaObjAreaAgenciaSaldo CS1 " _
    & "                                 Where   CS1.cCtaContCod = CS.cCtaContCod " _
    & "                                         AND CS.cAreaCod = CS1.cAreaCod AND CS.cAgeCod=CS1.cAgeCod " _
    & "                                         AND dAreaSaldoFecha<'" & Format(pdFecha, psFormatoFecha) & " 23:59:59')"

If oConect.AbreConexion = False Then Exit Function

Set RS = oConect.CargaRecordSet(sql)
If Not RS.EOF And Not RS.BOF Then
    GetSaldoAreaAgencia = RS!SaldoDia
End If
RS.Close
Set RS = Nothing
End Function

Public Function GetSaldoCtaCont(ByVal psFormatoFecha As String, ByVal pdFecha As Date, _
                                    ByVal psCtaContCod As String, Optional ByVal pnMoneda As Integer = 1) As Currency
Dim sql As String
Dim oConect As DConecta
Dim RS As ADODB.Recordset

Set oConect = New DConecta
Set RS = New ADODB.Recordset

'sql = " SELECT  dCtaSaldoFecha, nCtaSaldoImporte , ISNULL(MOVDIA.NDEBE,0) AS DEBE, " _
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
sql = "SELECT dbo.GetSaldoCta('" & Format(pdFecha, psFormatoFecha) & "','" & psCtaContCod & "', " & pnMoneda & ") as SaldoDia "
If oConect.AbreConexion = False Then Exit Function
Set RS = oConect.CargaRecordSet(sql)
If Not RS.EOF And Not RS.BOF Then
    GetSaldoCtaCont = RS!SaldoDia
End If
RS.Close
Set RS = Nothing
End Function
Public Function ExisteMovimiento(psMovNro As String, psOpeCod As String) As Boolean
Dim psSql   As String
Dim oConect As New DConecta
Dim prs     As ADODB.Recordset
ExisteMovimiento = True
   psSql = "SELECT m.cMovNro FROM Mov m WHERE m.cMovNro LIKE '" & psMovNro & "%' and m.nMovEstado = '" & gMovEstContabMovContable & "' and not m.nMovFlag in ('" & gMovFlagEliminado & "','" & gMovFlagModificado & "') and cOpeCod = '" & psOpeCod & "'"
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
Dim RS As New ADODB.Recordset
   If oConect.AbreConexion = False Then Exit Function
   sql = vsServerCom & "spGetTreeObj '" & Format(pnObjetoCod, "00") & "', " & pnNivel & ", '" & psObjFiltro & "'"
   Set RS = oConect.CargaRecordSet(sql)
   Set GetObjetosArbol = RS
   oConect.CierraConexion: Set oConect = Nothing
Set RS = Nothing
End Function

Public Function GetLibroDiario(ByVal dFecIni As Date, ByVal dFecFin As Date, psAge As String, psOpe As String) As Recordset
Dim sAgeCond As String, sOpeCond As String
Dim sql As String
Dim oConect As New DConecta
 On Error GoTo GetLibroDiarioErr
 
 If oConect.AbreConexion = False Then Exit Function
    sAgeCond = "": sOpeCond = ""
    If psAge <> "" Then
       sAgeCond = " And SubString(M.cMovNro,18,2) = '" & Right(psAge, 2) & "' "
    End If
    If psOpe <> "" Then
       sOpeCond = " And O.cOpeGruCod = '" & Trim(psOpe) & "' "
    End If

    sql = "Select M1.nMovCorrela, M.cMovNro, O.cOpeGruCod, MC.cCtaContCod, C.cCtaContDesc, MC.nMovImporte, og.cOpeGruDesc " _
        & "From Mov M JOIN MovCont M1 ON M1.nMovNro = M.nMovNro " _
        & "           Join MovCta MC On M.nMovNro = MC.nMovNro " _
        & "           Join " & vsServerCom & "CtaCont C On MC.cCtaContCod = C.cCtaContCod " _
        & "           Join " & vsServerCom & "OpeTpo  O On M.cOpeCod = O.cOpeCod " _
        & "           Join " & vsServerCom & "OpeGrupo Og ON og.cOpeGruCod = o.cOpeGruCod " _
        & "Where M.nMovEstado = " & gMovEstContabMovContable & " and not M.nMovFlag IN (" & gMovFlagEliminado & "," & gMovFlagModificado & ") and substring(M.cMovNro,1,8) between " _
        & "      '" & Format(dFecIni, "yyyymmdd") & "' And '" & Format(dFecFin, "yyyymmdd") _
        & "'     AND mc.cCtaContCod <> '' " & sAgeCond & sOpeCond _
        & "ORDER BY M1.nMovCorrela, substring(M.cMovNro,1,8), substring(M.cMovNro,18,2), O.cOpeGruCod, M.cMovNro, Mc.nMovItem "
    Set GetLibroDiario = oConect.CargaRecordSet(sql)
    oConect.CierraConexion: Set oConect = Nothing
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
   RSClose prs
   Err.Raise 50001, "nContFunciones:ValidaDocumento", "Documento de Proveedor ya registrado..."
End If
RSClose prs
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

'***Agregado por ELRO el 20111024, según Acta N° 277-2011/TI-D
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

Public Function recuperarCtaConAgencia(ByVal psAgecod As String, _
                                       ByVal psMoneda As String) As ADODB.Recordset
    On Error GoTo recuperarCtaConAgenciaErr
    Dim sql As String
    Dim oConect As COMConecta.DCOMConecta
    Set oConect = New COMConecta.DCOMConecta
    
    If oConect.AbreConexion = False Then
        Exit Function
    End If
    
    sql = "exec stp_sel_RecuperarCtaConAgencia '" & psAgecod & "', '" & psMoneda & "' "
    Set recuperarCtaConAgencia = oConect.CargaRecordSet(sql)
    
    oConect.CierraConexion: Set oConect = Nothing
    Exit Function
    
recuperarCtaConAgenciaErr:
    Call RaiseError(MyUnhandledError, "DOperacion:recuperarCtaConAgencia Method")

End Function

Public Function recuperarCtaConAgencia_02(ByVal psAgecod As String, _
                                       ByVal psMoneda As String) As ADODB.Recordset
    On Error GoTo recuperarCtaConAgencia02Err
    Dim sql As String
    Dim oConect As COMConecta.DCOMConecta
    Set oConect = New COMConecta.DCOMConecta
    
    If oConect.AbreConexion = False Then
        Exit Function
    End If
    
    sql = "exec stp_sel_RecuperarCtaConAgencia_02 '" & psAgecod & "', '" & psMoneda & "' "
    Set recuperarCtaConAgencia_02 = oConect.CargaRecordSet(sql)
    
    oConect.CierraConexion: Set oConect = Nothing
    Exit Function
    
recuperarCtaConAgencia02Err:
    Call RaiseError(MyUnhandledError, "DOperacion:recuperarCtaConAgencia_02 Method")

End Function


'***Fin Agregado por ELRO*************************************
'EJVG20140904 ***
Public Function BuscaCtaEquivalente(ByVal psCtaContCod As String) As String
    On Error GoTo ErrBuscaCtaEquivalente
    Dim sql As String
    Dim RS As New ADODB.Recordset
    Dim oConect As New COMConecta.DCOMConecta
    oConect.AbreConexion
    sql = "exec stp_sel_ERS0252014_BuscaCtaEquivalente '" & psCtaContCod & "'"
    Set RS = oConect.CargaRecordSet(sql)
    If Not RS.EOF Then
        BuscaCtaEquivalente = RS!cCtaContCodN
    End If
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Function
ErrBuscaCtaEquivalente:
    Call RaiseError(MyUnhandledError, "NContFunciones:BuscaCtaEquivalente Method")
End Function
'END EJVG *******

