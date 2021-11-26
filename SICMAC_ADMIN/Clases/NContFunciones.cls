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
Dim vsConexion As String
Dim vsServerCom As String
Dim vsServerPers As String
Dim vsFormatoFecha As String
Public Sub Inicio(ByVal psFormatoFecha As String)
Attribute Inicio.VB_UserMemId = 0
vsFormatoFecha = psFormatoFecha
End Sub
Public Sub ExtornaAsiento()
    On Error GoTo ExtornaAsientoErr

    'your code goes here...

    Exit Sub
ExtornaAsientoErr:
    Call RaiseError(MyUnhandledError, "NContFunciones:ExtornaAsiento Method")
End Sub
'##ModelId=3A836D9600EA

Public Function GeneraDocNro(psDocTpo As TpoDoc, psMoneda As Moneda, Optional psDocSerie As String = "", Optional psTipo As String = "") As String
    'NAGL 20191212 Agregó Optional psTipo As String Según INC1912120012
    On Error GoTo GeneraMovNroErr
    Dim sSQL As String
    Dim rs As New ADODB.Recordset
    Dim oConect As DConecta
    Dim oCaja As nCajaGeneral

    Set oCaja = New nCajaGeneral
    Set oConect = New DConecta
    If oConect.AbreConexion = False Then Exit Function

        'If psDocSerie <> "" Then
            'sSql = "SELECT max(cDocNro) AS cDocCorrela FROM Movdoc md JOIN Mov m ON m.nMovNro = md.nMovNro " _
                '& "WHERE   m.nMovFlag NOT IN ('" & gMovFlagEliminado & "') and nDocTpo = '" & Format(psDocTpo, "00") & "' and substring(cDocNro,1," & Len(psDocSerie) & ") = '" & psDocSerie & "'"
        'Else
            'sSql = "SELECT  cDocNro AS cDocCorrela , md.nMovnro " _
                '& " FROM    movdoc md " _
                '& " WHERE   nDocTpo = '" & Format(psDocTpo, "00") & "' " _
                '& "         and md.nmovnro = (  Select Max(MD1.nmovnro)  " _
                '& "                             From MovDoc MD1 JOIN Mov M on M.nMovNro= MD1.nMovnro " _
                '& "                             WHERE MD1.nDocTpo = '" & Format(psDocTpo, "00") & "' and Substring(M.cOpeCod,3,1) ='" & psMoneda & "') "

        'End If
        '************Comentado by NAGL 20191212 *******************
        sSQL = "Exec stp_sel_ObtieneGeneraDocNro '" & Format(psDocTpo, "00") & "', '" & psMoneda & "', '" & psDocSerie & "', '" & psTipo & "'"
        '********Agregado by NAGL 20191212 Según INC1912120012*****
        
        Set rs = oConect.CargaRecordSet(sSQL)
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
            Else
                GeneraDocNro = "00000001"
                If psDocSerie <> "" Then GeneraDocNro = psDocSerie & "-" & GeneraDocNro
            End If
        Else
            GeneraDocNro = "00000001"
            If psDocSerie <> "" Then GeneraDocNro = psDocSerie & "-" & GeneraDocNro
        End If
        rs.Close
        Set rs = Nothing
        oConect.CierraConexion
        Set oConect = Nothing
        If psDocTpo = TpoDocOrdenPago Then
            GeneraDocNro = oCaja.VerificaOPRango(GeneraDocNro, psMoneda)
        End If
        Set oCaja = Nothing
    Exit Function
GeneraMovNroErr:
    Call RaiseError(MyUnhandledError, "NContFunciones:GeneraMovNro Method")
End Function




Public Function GetFiltroObjetos(ByVal pnTipoObj As TpoObjetos, ByVal psCtaContCod As String, ByVal psObjetoCod As String, Optional lbMuestraCta As Boolean = True) As String
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
            sql = "Select cPersCod, cCtaIfCod, cCtaContCod, cCtaIFSubCta as SubCta From " & vsServerCom & "CtaIFFiltro WHERE  cPersCod = '" & lsPersCtaIF & "' and cIFTpo='" & lsTpoIf & "' AND cCtaIfCod='" & lsCtaIf & "' AND cCtaContCod in ('" & psCtaContCod & "')"
        Case Else
            sql = "SELECT cCtaObjSubCta as SubCta FROM " & vsServerCom & "CtaObjFiltro WHERE cCtaContCod in ('" & psCtaContCod & "') and cObjetoCod = '" & psObjetoCod & "'"
    End Select
    Set rs = oConect.CargaRecordSet(sql)
    If Not rs.EOF Then
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



'##ModelId=3A8ACF8E033C
Public Sub ValidaEstadoMov()
    On Error GoTo ValidaEstadoMovErr

    'your code goes here...

    Exit Sub
ValidaEstadoMovErr:
    Call RaiseError(MyUnhandledError, "NContFunciones:ValidaEstadoMov Method")
End Sub

'##ModelId=3A9D4B16035B
Public Sub DefineTpoCambio()
    On Error GoTo DefineTpoCambioErr

    'your code goes here...

    Exit Sub
DefineTpoCambioErr:
    Call RaiseError(MyUnhandledError, "NContFunciones:DefineTpoCambio Method")
End Sub

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

Dim oIni As ClasIni
Set oIni = New ClasIni
vsConexion = oIni.CadenaConexion
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
       sql = "sp_GeneraMovNro '" & Format(pdFecha & " " & oConect.GetHoraServer, gsFormatoFechaHora) & "','" & psCodAge & "','" & psUser & "'"
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
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim oConect As DConecta
    
    Set oConect = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConect.AbreConexion = False Then Exit Function
        
    If Len(psCtaContCod) > 4 Then
       sSQL = "SELECT RTRIM(ISNULL(c1.cCtaContDesc,'')) + ' ' + RTRIM(c.cCtaContDesc) as cCtaContDesc " _
              & "FROM   " & vsServerCom & "CtaCont c LEFT JOIN " & vsServerCom & "CtaCont c1 ON c1.cCtaContCod = substring(c.cCtaContCod,1," & Len(psCtaContCod) - 2 & ") " _
              & "WHERE  c.cCtaContCod = '" & psCtaContCod & "'"
    Else
       sSQL = "SELECT cCtaContDesc FROM " & vsServerCom & "CtaCont WHERE cCtaContCod = '" & psCtaContCod & "'"
    End If
    EmiteCtaContDesc = ""
    Set rs = oConect.CargaRecordSet(sSQL)
    If Not rs.BOF And Not rs.EOF Then
        EmiteCtaContDesc = Trim(rs!cCtaContDesc)
    End If
    rs.Close: Set rs = Nothing
    
    oConect.CierraConexion:  Set oConect = Nothing
    Exit Function
EmiteCtaContDescErr:
    Call RaiseError(MyUnhandledError, "DMov:EmiteCtaContDesc Method")
End Function
Public Function GetValorMargenes(ByVal pTipoMargenes As TpoMargenes) As Integer
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Dim lsTipoMargen As String
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function
Set rs = New ADODB.Recordset
Select Case pTipoMargenes
    Case MargenSuperior
        lsTipoMargen = "nMargSup"
    Case MargenDerecho
        lsTipoMargen = "nMargDer"
    Case MargenIzquierdo
        lsTipoMargen = "nMargIzq"
    Case MargenInferior
        lsTipoMargen = "nMargInf"
End Select
GetValorMargenes = 0
sql = "Select cNomVar,cValorVar From Varsistema Where cCodProd ='CON' AND  cNomVar IN ('" & lsTipoMargen & "')"
Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetValorMargenes = Val(rs!cValorVar)
End If
rs.Close: Set rs = Nothing
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function GetObjetos(ByVal pnObjetoCod As TpoObjetos, Optional pbObjMiembros As Boolean = True, Optional pbLopProveLogAge As Boolean = False) As ADODB.Recordset
Dim sql As String
Dim oConect As DConecta
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

If pbObjMiembros Then
    If pbLopProveLogAge = True Then
        sql = "Select  O.cObjetoCod,Case When cmensaje='' Then CONVERT(CHAR(40),cObjetoDesc) else cMensaje end as Descripcion, nObjetoNiv  " _
             & " From    " & vsServerCom & "Objeto O INNER JOIN LogPoderesObjeto LPO  Where O.cObjetoCod Like '" & Format(pnObjetoCod, "00") & "__' "
        
    Else
        sql = "Select  cObjetoCod, CONVERT(CHAR(40),cObjetoDesc) as Descripcion, nObjetoNiv  " _
             & " From    " & vsServerCom & "Objeto Where cObjetoCod Like '" & Format(pnObjetoCod, "00") & "__' "
    End If
Else
    sql = "Select  cObjetoCod, CONVERT(CHAR(40),cObjetoDesc) as Descripcion, nObjetoNiv  " _
         & " From    " & vsServerCom & "Objeto Where cObjetoCod Like '" & Format(pnObjetoCod, "00") & "%' "
End If
Set rs = oConect.CargaRecordSet(sql)
Set GetObjetos = rs
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
         psSql = "SELECT cAgeCod cObjetoCod, cAgeDescripcion cObjetoDesc, '1' nObjetoNiv, cSubCtaCod FROM Agencias WHERE cSubCtaCod = LEFT('" & psSubCta & "',LEN(cAgeCod))"
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
         psSql = "SELECT cif.cPersCod + cif.cCtaIfCod cObjetoCod, cCtaIFDesc, cObjetoDesc '1' nObjetoNiv,  cCtaIFSubCta cSubCtaCod FROM CtaIF cif JOIN CtaIFFiltro ciff ON cif.cPersCod = ciff.cPersCod and cif.cCtaIfCod = ciff.cCtaIfCod " _
               & "WHERE cCtaIFSubCta = LEFT('" & psSubCta & "',LEN(cCtaIFSubCta)) "
         Set prs = oConect.CargaRecordSet(psSql)
      Case ObjDescomEfectivo
         Set prs = Nothing
      Case ObjPersona
         Set prs = Nothing
      Case Else
         Set prs = oCta.CargaCtaObjFiltro(psCtaCod, psObjCod, "'" & psSubCta & "' LIKE cCtaObjSubCta + '%'", True)
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
Public Function GetObjetoDatos(ByVal psMovNro As String, ByVal psMovItem As String, ByVal psMovObjOrden As String, ByVal psObjCod As String) As ADODB.Recordset
Dim oConect As New DConecta
Dim prs     As ADODB.Recordset
Dim psSql   As String
On Error GoTo GetObjetoDatosErr
If oConect.AbreConexion = False Then Exit Function
   Select Case Val(psObjCod)
      Case ObjCMACAgencias, ObjCMACAgenciaArea, ObjCMACArea
         psSql = "SELECT mo.cAreacod+mo.cAgeCod cObjetoCod, ISNULL(ag.cAgeDescripcion,'')+' '+ISNULL(a.cAreaDescripcion,'') cObjetoDesc " _
               & "FROM MovObjAreaAgencia mo JOIN AreaAgencia aa ON aa.cAreaCod = mo.cAreaCod and aa.cAgeCod = mo.cAgeCod " _
               & "     LEFT JOIN Areas a ON a.cAreaCod = aa.cAreaCod " _
               & "     LEFT JOIN Agencias ag ON ag.cAgeCod = aa.cAgeCod " _
               & "WHERE mo.cMovNro = '" & psMovNro & "' and mo.cMovITem = '" & psMovItem & "' and mo.cMovObjOrden = '" & psMovObjOrden & "' "
         Set prs = oConect.CargaRecordSet(psSql)
         
      Case ObjEntidadesFinancieras
         psSql = "SELECT cif.cPersCod + cif.cCtaIfCod cObjetoCod, RTRIM(p.cPersNombre) + ' ' + cCtaIFDesc cObjetoDesc " _
               & "FROM MovObjIF mif " _
               & "     JOIN " & vsServerPers & "Persona p ON p.cPersCod = mif.cPersCod " _
               & "     JOIN CtaIF cif ON cif.cPersCod = mif.cPersCod and cif.cCtaIFCod = mif.cCtaIFCod " _
               & "WHERE mif.cMovNro = '" & psMovNro & "' and mif.cMovITem = '" & psMovItem & "' and mif.cMovObjOrden = '" & psMovObjOrden & "' "
         Set prs = oConect.CargaRecordSet(psSql)
         
      Case ObjDescomEfectivo
         psSql = "SELECT moe.cEfectivoCod cObjetoCod, e.nEfectivoValor cObjetoDesc " _
               & "FROM   MovObjEfectivo moe JOIN Efectivo e ON e.cEfectivoCod = mif.cEfectivoCod " _
               & "WHERE moe.cMovNro = '" & psMovNro & "' and moe.cMovITem = '" & psMovItem & "' and moe.cMovObjOrden = '" & psMovObjOrden & "' "
         Set prs = oConect.CargaRecordSet(psSql)
      Case ObjPersona
         psSql = "SELECT mop.cPersCod cObjetoCod, p.cPersNombre cObjetoDesc " _
               & "FROM   MovObjPers mop JOIN Persona p ON p.cPersCod = mop.cPersCod " _
               & "WHERE mop.cMovNro = '" & psMovNro & "' and mop.cMovITem = '" & psMovItem & "' and mop.cMovObjOrden = '" & psMovObjOrden & "' "
         Set prs = oConect.CargaRecordSet(psSql)
      Case Else
         Set prs = Nothing
   End Select
   Set GetObjetoDatos = prs
Set oConect = Nothing
Set prs = Nothing
Exit Function
GetObjetoDatosErr:
   Call RaiseError(MyUnhandledError, "NContFunciones: GetObjetoDatos Method")
End Function
Public Function FactorAjuste(ByVal dFecAdq As Date, ByVal dFecRep As Date, Optional nDeci As Integer = 4) As Double
Dim sSQL As String
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
    & "                             JOIN MOVCTA MC ON MC.CMOVNRO = M.CMOVNRO " _
    & "                             LEFT JOIN MOVME ME ON ME.CMOVNRO = MC.CMOVNRO AND ME.CMOVITEM=MC.CMOVITEM " _
    & "                     WHERE   SUBSTRING(M.CMOVNRO,1,8)='" & Format(pdFecha, "yyyymmdd") & "' " _
    & "                             AND M.nMOVFLAG NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") " _
    & "                             AND M.nMOVESTADO=" & gMovEstContabMovContable & " " _
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
    GetSaldoAreaAgencia = rs!SALDODIA
End If
rs.Close
Set rs = Nothing
End Function
Public Function GetSaldoCtaCont(ByVal psFormatoFecha As String, ByVal pdFecha As Date, _
                                    ByVal psCtaContCod As String) As Currency
Dim sql As String
Dim oConect As DConecta
Dim rs As ADODB.Recordset

Set oConect = New DConecta
Set rs = New ADODB.Recordset

sql = " SELECT  dCtaSaldoFecha, nCtaSaldoImporte , ISNULL(MOVDIA.NDEBE,0) AS DEBE, " _
    & "         ISNULL(MOVDIA.NHABER,0) AS HABER,  nCtaSaldoImporte + ISNULL(MOVDIA.NDEBE,0) + ISNULL(MOVDIA.NHABER,0) AS SALDODIA " _
    & " FROM    CtaSaldo CS " _
    & "         LEFT JOIN (SELECT   MC.CCTACONTCOD , " _
    & "                             SUM(CASE WHEN MC.NMOVIMPORTE>0 THEN ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) ELSE 0 END) AS NDEBE, " _
    & "                             SUM(CASE WHEN MC.NMOVIMPORTE<0 THEN ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) ELSE 0 END) AS NHABER " _
    & "                     FROM    MOV M " _
    & "                             JOIN MOVCTA MC ON MC.CMOVNRO = M.CMOVNRO " _
    & "                             LEFT JOIN MOVME ME ON ME.CMOVNRO = MC.CMOVNRO AND ME.CMOVITEM=MC.CMOVITEM " _
    & "                     WHERE   SUBSTRING(M.CMOVNRO,1,8)='" & Format(pdFecha, "yyyymmdd") & "' " _
    & "                             AND M.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "') " _
    & "                             AND M.nMovEstado='" & gMovEstContabMovContable & "' " _
    & "                     GROUP BY MC.CCTACONTCOD) AS MOVDIA " _
    & "         ON MOVDIA.CCTACONTCOD = CS.CCTACONTCOD " _
    & "         WHERE   CS.cCtaContCod Like '" & psCtaContCod & "' " _
    & "                 AND dCtaSaldoFecha= (   SELECT  MAX(dCtaSaldoFecha) " _
    & "                                         FROM    CtaSaldo CS1 " _
    & "                                         Where   CS1.cCtaContCod = CS.cCtaContCod " _
    & "                                                 AND dCtaSaldoFecha<'" & Format(pdFecha, psFormatoFecha) & " 23:59:59')"


If oConect.AbreConexion = False Then Exit Function

Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetSaldoCtaCont = rs!SALDODIA
End If
rs.Close
Set rs = Nothing
End Function
Public Function VerAjuste(psMovNro As String, psTipoAjuste As String) As Boolean
Dim psSql   As String
Dim oConect As New DConecta
Dim prs     As ADODB.Recordset
VerAjuste = True
   psSql = "SELECT m.cMovNro FROM Mov m JOIN MovOtrosItem mo ON mo.nMovNro = m.nMovNro where m.cMovNro LIKE '" & psMovNro & "%' and m.nMovEstado = '" & gMovEstContabMovContable & "' and m.nMovFlag <> '" & gMovFlagEliminado & "' and cMovOtroVariable like '" & psTipoAjuste & "%'"
   oConect.AbreConexion
      Set prs = oConect.CargaRecordSet(psSql)
      If prs.EOF Then
         VerAjuste = False
      End If
   Set oConect = Nothing
   prs.Close: Set prs = Nothing
End Function

'****************************************
'***************** NUEVAS FUNCIONES
'****************************************
Public Function GetObjetosArbol(ByVal pnObjetoCod As TpoObjetos, ByVal psObjFiltro As String, ByVal pnNivel As Integer) As ADODB.Recordset
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
        & "Where M.nMovEstado = " & gMovEstContabMovContable & " and M.nMovFlag <> " & gMovFlagEliminado & " and substring(M.cMovNro,1,8) between " _
        & "      '" & Format(dFecIni, "yyyymmdd") & "' And '" & Format(dFecFin, "yyyymmdd") _
        & "'     AND mc.cCtaContCod <> '' " & sAgeCond & sOpeCond _
        & "ORDER BY M1.nMovCorrela, substring(M.cMovNro,1,8), substring(M.cMovNro,18,2), O.cOpeGruCod"
    Set GetLibroDiario = oConect.CargaRecordSet(sql)
    oConect.CierraConexion: Set oConect = Nothing
Exit Function
GetLibroDiarioErr:
   Call RaiseError(MyUnhandledError, "NContFunciones: GetLibroDiario Method")
End Function
