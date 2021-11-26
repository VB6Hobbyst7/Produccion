VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMCajaCtasIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Base 0
Option Explicit
Dim vsConexion As String
Dim vsServerAdmin As String
Dim vsServerComunes As String
Dim vsServerNegocio As String
Dim vsServerPersonas As String
Dim vsServerImagenes As Variant

Dim oError As New COMConecta.COMErrorHandling

Public Function GetCtaIfFiltro(ByVal psCtaContCod As String, psCtaIfCod As String, psPersCod As String, Optional lbMuestraCta As Boolean = True) As String
    On Error GoTo GetCtaIfFiltroErr
    Dim sql As String
    Dim rs   As New ADODB.Recordset
    Dim oConect As COMConecta.DCOMConecta
    
    Set oConect = New COMConecta.DCOMConecta
    If oConect.AbreConexion = False Then Exit Function
    GetCtaIfFiltro = ""
    sql = "SELECT cCtaContCod ,  cCtaIFSubCta " _
        & " FROM " & vsServerComunes & "CtaIFFiltro WHERE cCtaContCod in ('" & psCtaContCod & "') " _
        & "         and cCtaIFCod = '" & psCtaIfCod & "' and cPersCod ='" & psPersCod & "'"
    
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
    Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones:GetCtaObjFiltro Method")
End Function

Public Function CargaCtaIFFiltro(ByVal psCtaCod As String, Optional psPersCod As String = "", Optional psIFTpo As String, Optional psCtaIfCod As String) As ADODB.Recordset
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
If psCtaIfCod <> "" Then
    lsFiltro = lsFiltro & " and cif.cCtaIFCod = '" & psCtaIfCod & "' "
End If
sSql = "SELECT cCtaContCod, 1 nCtaObjOrden, cif.cIFTpo + '.' + cif.cPersCod + '.' + cif.cCtaIFCod cObjetoCod, RTRIM(P.cPersNombre) + ' ' + RTRIM(ci.cCtaIFDesc) cObjetoDesc, cCtaIFSubCta cSubCtaCod " _
     & "FROM CtaIFFiltro cif JOIN CtaIF ci ON ci.cPersCod = cif.cPersCod and ci.cIFTpo = cif.cIFTpo and ci.cCtaIFCod = cif.cCtaIFCod " _
     & "        JOIN INSTITUCIONFINANC I ON I.cPersCod = ci.cPersCod and I.cIFTpo = ci.cIFTpo " _
     & "        JOIN PERSONA P ON P.CPERSCOD = I.CPERSCOD " _
     & "WHERE cif.cCtaContCod = '" & psCtaCod & "' " & lsFiltro & " ORDER BY cif.cIFTpo, cif.cPersCod, cif.cCtaIFCod "
    
    Dim oConect As COMConecta.DCOMConecta
    Set oConect = New COMConecta.DCOMConecta
    If oConect.AbreConexion = False Then Exit Function
    Set CargaCtaIFFiltro = oConect.CargaRecordSet(sSql)
    'RSClose rs
    rs.Close
    oConect.CierraConexion: Set oConect = Nothing
    Exit Function
CargaCtaIFFiltroErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub EliminaCtaIFFiltro(ByVal psCtaCod As String, Optional psPersCod As String = "", Optional psIFTpo As String, Optional psCtaIfCod As String)
Dim lsFiltro As String
Dim sSql     As String
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
sSql = "DELETE CtaIFFiltro " _
     & "WHERE cCtaContCod = '" & psCtaCod & "' " & lsFiltro
    
    Dim oConect As New COMConecta.DCOMConecta
    If oConect.AbreConexion = False Then Exit Sub
    oConect.Ejecutar sSql
    oConect.CierraConexion: Set oConect = Nothing
    Exit Sub
EliminaCtaIFFiltroErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub InsertaCtaIFFiltro(ByVal psCtaCod As String, psPersCod As String, psIFTpo As String, psCtaIfCod As String, psCtaIfSubCta As String)
Dim lsFiltro As String
Dim sSql     As String
On Error GoTo InsertaCtaIFFiltroErr

sSql = "INSERT CtaIFFiltro (cPersCod, cIFTpo, cCtaIfCod, cCtaContCod, cCtaIFSubCta) " _
        & "VALUES ('" & psPersCod & "', '" & psIFTpo & "','" & psCtaIfCod & "','" & psCtaCod & "', '" & psCtaIfSubCta & "')"
    
    Dim oConect As New COMConecta.DCOMConecta
    If oConect.AbreConexion = False Then Exit Sub
    oConect.Ejecutar sSql
    oConect.CierraConexion: Set oConect = Nothing
    Exit Sub
InsertaCtaIFFiltroErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'devengado.
Public Function CargaCtasIF(Optional ByVal pnMoneda As COMDConstantes.Moneda = 0, Optional psFiltroCtaIF As String = "", Optional pnMuestraIF As MuestraIF = COMDConstantes.MuestraCuentas, Optional ByVal psCanje As String = "", Optional psTipoIF As COMDConstantes.CGTipoIF) As ADODB.Recordset
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
    Dim pos As String
    Dim lsCadAux As String
    Dim lsFiltroMoneda As String
        
    lsIFFiltro = ""
    If psFiltroCtaIF <> "" Then
        ' se filtran las cuentas segun tipo
        lsIFFiltro = " AND CI.cIFTpo+CI.cCtaIfCod LIKE '" & psFiltroCtaIF & "'"
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
    & " FROM ( " _
    & "         SELECT  I.cIFTpo + '.' + CI.CPERSCOD as CPERSCOD, CI.cCtaIFCod, " _
    & "                 CONVERT(CHAR(40),CI.cCtaIFDesc) AS CTAIFDESC, " _
    & "                 LEN(CI.cCtaIFCod) AS Nivel, I.cIFTpo , I.bCanje " _
    & "         FROM    INSTITUCIONFINANC I JOIN CTAIF CI ON CI.cPersCod = I.cPersCod AND I.cIFTpo= CI.cIFTpo " _
    & "         WHERE   SUBSTRING(CI.CCTAIFCOD,1,1) NOT IN('X') AND CI.cCtaIFEstado='" & gEstadoCtaIFActiva & "' " & lsFiltroMoneda & lsIFFiltro & lsFiltroTipoIF _
    & "         Union " _
    & "         SELECT  I.cIFTpo + '.' + I.CPERSCOD as CPERSCOD, '' AS CTAIF, P.CPERSNOMBRE , 1 AS NIVEL , " _
    & "                 I.cIFTpo , I.bCanje " _
    & "         FROM    INSTITUCIONFINANC I " _
    & "                 JOIN PERSONA P ON P.CPERSCOD = I.CPERSCOD " _
    & "                 JOIN (  SELECT  CI.cIFTpo, CI.CPERSCOD " _
    & "                         FROM    CTAIF CI " _
    & "                         WHERE   SUBSTRING(CI.CCTAIFCOD,1,1) NOT IN('X') AND CI.cCtaIFEstado='" & gEstadoCtaIFActiva & "' " & lsFiltroMoneda & lsIFFiltro _
    & "                         ) AS C1 " _
    & "                 ON  C1.cIFTpo=I.cIFTpo AND C1.CPERSCOD= I.CPERSCOD " _
    & "         UNION Select  Replace(Str(nConsValor,2,0),' ','0') as cPerscod, '' as CtaIf , " _
    & "                 cConsDescripcion , 0 AS  NIVEL, Replace(Str(nConsValor,2,0),' ','0') as cIFTpo, 0 as bCanje " _
    & "         From    Constante " _
    & "         Where   nConsCod Like " & gCGTipoIF & " And nConsValor <> " & gCGTipoIF & " " _
    & "                 AND  Replace(Str(nConsValor,2,0),' ','0') IN (  SELECT  I.cIFTpo " _
    & "                                                                 FROM    INSTITUCIONFINANC I " _
    & "                                                                         JOIN CTAIF CI ON CI.cPersCod = I.cPersCod AND I.cIFTpo= CI.cIFTpo "
sql = sql + "                                                           WHERE   SUBSTRING(CI.CCTAIFCOD,1,1) NOT IN('X') AND CI.cCtaIFEstado='" & gEstadoCtaIFActiva & "' " & lsFiltroMoneda & lsIFFiltro & " )" _
    & "  ) AS CTASIF " _
    & "  Where Nivel <= " & pnMuestraIF & lsFiltroCanje _
    & "  ORDER BY CPERSCOD, cCtaIFCod "


    Set rs = oConec.CargaRecordSet(sql)
    Set CargaCtasIF = rs
    
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
CargaCtasIFErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DCajaCtasIF:CargaCtasIF Function")
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
        NombreIF = Trim(rs!cpersnombre)
    End If
    rs.Close
    Set rs = Nothing
        
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Function
NombreIFErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DCajaCtasIF:EliminaInteres Method")
End Function

Public Function SubCuentaIF(ByVal psPersCodIF As String) As String
    On Error GoTo SubCuentaIFErr
    Dim sql As String
    Dim oConect As COMConecta.DCOMConecta
    Set oConect = New COMConecta.DCOMConecta
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
    Call oError.RaiseError(oError.MyUnhandledError, "DCajaCtasIF:EliminaInteres Method")
End Function

Public Function GetCodPersAuxIF(ByVal psObjetoCod As String) As String
    On Error GoTo GetCodPersAuxIFErr
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim oConect As New COMConecta.DCOMConecta
    GetCodPersAuxIF = ""
    If oConect.AbreConexion = False Then Exit Function
    
    sql = "Select cCodPers from AuxIF where cObjetoCod ='" & psObjetoCod & "'"
    Set rs = oConect.CargaRecordSet(sql)
    If Not rs.EOF And Not rs.BOF Then
        GetCodPersAuxIF = Trim(rs!cCodPers)
    End If
    rs.Close
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Function
GetCodPersAuxIFErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DCajaCtasIF:GetCodPersAuxIF Method")
End Function

Private Sub Class_Initialize()
Dim oIni As COMConecta.DCOMClasIni
Set oIni = New COMConecta.DCOMClasIni

Set oError = New COMConecta.COMErrorHandling

vsServerAdmin = oIni.BaseAdministracion
vsServerComunes = oIni.BaseComunes
vsServerNegocio = oIni.BaseNegocio
vsServerPersonas = oIni.BasePersonas
vsServerImagenes = oIni.BaseImagenes

End Sub

