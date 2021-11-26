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
Attribute vsServerComunes.VB_VarDescription = "Se guarda la caden aen la cual se guardan el nombre del Servidor y la base de datos que posee la tabla de comunes."
Dim vsServerNegocio As String
Dim vsServerPersonas As String
Dim vsServerImagenes As Variant


'##ModelId=3A8C25DF035B
Public Function GetCtaIfFiltro(ByVal psCtaContCod As String, psCtaIfCod As String, psPersCod As String, Optional lbMuestraCta As Boolean = True) As String
    On Error GoTo GetCtaIfFiltroErr
    Dim sql As String
    Dim rs   As New ADODB.Recordset
    Dim oconect As DConecta
    
    Set oconect = New DConecta
    If oconect.AbreConexion = False Then Exit Function
    GetCtaIfFiltro = ""
    sql = "SELECT cCtaContCod ,  cCtaIFSubCta " _
        & " FROM " & vsServerComunes & "CtaIFFiltro WHERE cCtaContCod in ('" & psCtaContCod & "') " _
        & "         and cCtaIFCod = '" & psCtaIfCod & "' and cPersCod ='" & psPersCod & "'"
    
    Set rs = oconect.CargaRecordSet(sql)
    If Not rs.EOF Then
        GetCtaIfFiltro = IIf(lbMuestraCta, rs!cCtaContCod & rs!cCtaIFSubCta, rs!cCtaContCod)
        If GetCtaIfFiltro = "" Then
            GetCtaIfFiltro = IIf(lbMuestraCta, rs!cCtaContCod, "")
        End If
    Else
        If GetCtaIfFiltro = "" Then
            GetCtaIfFiltro = psCtaContCod
        End If
    End If
    rs.Close: Set rs = Nothing
    oconect.CierraConexion
    Set oconect = Nothing
    Exit Function
GetCtaIfFiltroErr:
    Call RaiseError(MyUnhandledError, "NContFunciones:GetCtaObjFiltro Method")
End Function

'devengado.
'##ModelId=3A80C42E0163
Public Function CargaCtasIF(Optional ByVal pnMoneda As Moneda = 0, Optional psFiltroCtaIF As String = "", Optional pnMuestraIF As MuestraIF = MuestraCuentas, Optional ByVal psCanje As String = "") As ADODB.Recordset
    On Error GoTo CargaCtasIFErr
    Dim oConec As DConecta
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim lsIFFiltro As String
    Dim lsTpoCtaFiltro1 As String
    Dim lsTpoCtaFiltro As String
    Dim lsFiltroCanje  As String
    Dim lsFiltroCanje1 As String
    Dim Pos As String
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
        lsFiltroMoneda = "AND SUBSTRING(CI.cCtaIfCod,3,1)='" & pnMoneda & "' "
    End If
    
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() = False Then Exit Function
    
sql = " SELECT   CASE WHEN NIVEL =1 THEN CPERSCOD ELSE CPERSCOD + '.' + cCtaIFCod END AS CODIGO , " _
    & "          Convert(char(40),CTAIFDESC)  as CTAIFDESC, Nivel " _
    & " FROM ( " _
    & "         SELECT  I.cIFTpo + '.' + CI.CPERSCOD as CPERSCOD, CI.cCtaIFCod, " _
    & "                 CONVERT(CHAR(40),CI.cCtaIFDesc) AS CTAIFDESC, " _
    & "                 LEN(CI.cCtaIFCod) AS Nivel, I.cIFTpo , I.bCanje " _
    & "         FROM    INSTITUCIONFINANC I JOIN CTAIF CI ON CI.cPersCod = I.cPersCod AND I.cIFTpo= CI.cIFTpo " _
    & "         WHERE   SUBSTRING(CI.CCTAIFCOD,1,1) NOT IN('X') AND CI.cCtaIFEstado='" & gEstadoCtaIFActiva & "' " & lsFiltroMoneda & lsIFFiltro _
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
    Call RaiseError(MyUnhandledError, "DCajaCtasIF:CargaCtasIF Function")
End Function

'##ModelId=3A8AF4BD02A3
Public Function NombreIF(ByVal psPersCodIf As String) As String
    On Error GoTo NombreIFErr
    Dim sql As String
    Dim oconect As DConecta
    Set oconect = New DConecta
    If oconect.AbreConexion() = False Then Exit Function
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    NombreIF = ""
    sql = "Select p.cPersNombre from InstitucionFinanc I JOIN  " & vsServerPersonas & "PERSONA P ON P.cPersCod=I.cPersCod  where I.cPersCod='" & psPersCodIf & "'"
    Set rs = oconect.CargaRecordSet(sql)
    If Not rs.EOF And Not rs.BOF Then
        NombreIF = Trim(rs!cPersNombre)
    End If
    rs.Close
    Set rs = Nothing
        
    oconect.CierraConexion
    Set oconect = Nothing
    Exit Function
NombreIFErr:
    Call RaiseError(MyUnhandledError, "DCajaCtasIF:EliminaInteres Method")
End Function
Public Function SubCuentaIF(ByVal psPersCodIf As String) As String
    On Error GoTo SubCuentaIFErr
    Dim sql As String
    Dim oconect As DConecta
    Set oconect = New DConecta
    SubCuentaIF = ""
    If oconect.AbreConexion() = False Then Exit Function
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = "Select cSubCtaContCod from InstitucionFinanc where cPersCod='" & psPersCodIf & "'"
    Set rs = oconect.CargaRecordSet(sql)
    If Not rs.EOF And Not rs.BOF Then
        SubCuentaIF = Trim(rs!cSubCtaContCod)
    End If
    rs.Close
    Set rs = Nothing
        
    oconect.CierraConexion
    Set oconect = Nothing
    Exit Function
SubCuentaIFErr:
    Call RaiseError(MyUnhandledError, "DCajaCtasIF:EliminaInteres Method")
End Function
Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

Dim oIni As ClasIni
Set oIni = New ClasIni

vsConexion = oIni.CadenaConexion
vsServerAdmin = oIni.BaseAdministracion
vsServerComunes = oIni.BaseComunes
vsServerNegocio = oIni.BaseNegocio
vsServerPersonas = oIni.BasePersonas
vsServerImagenes = oIni.BaseImagenes

End Sub
