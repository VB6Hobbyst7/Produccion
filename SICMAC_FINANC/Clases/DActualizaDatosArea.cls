VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DActualizaDatosArea"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A9423060251"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Realiza el mantenimiento de las Areas Organizacionales
Option Base 0
Option Explicit

Private lsServerComunes As String
Private lsServerPersona As String
Private lsServerAdministracion As String
Private lsServerNegocio As String
Private lsServerImagenes As String

Public Sub Inicio(psServerComunes As String, psServerPersona As String, psServerAdministracion As String, psServerNegocio As String, psServerImagenes As String)
    lsServerComunes = psServerComunes
    lsServerPersona = psServerPersona
    lsServerAdministracion = psServerAdministracion
    lsServerNegocio = psServerNegocio
    lsServerImagenes = psServerImagenes
End Sub

Public Function GetAgencias(Optional ByVal psObjetoFiltro As String = "", Optional pbRaiz As Boolean = True, Optional pbEstadoActivo As Boolean = False) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Dim lsFiltro As String

Set oConect = New DConecta
Set rs = New ADODB.Recordset

If oConect.AbreConexion = False Then Exit Function
If psObjetoFiltro <> "" Then
    lsFiltro = " WHERE cAgeCod Like '" & psObjetoFiltro & "'  "
End If
If pbEstadoActivo Then
    lsFiltro = IIf(lsFiltro <> "", " and ", " where ") & " nEstado = 1 "
End If
sql = " Select Convert(Varchar(20),cAgeCod) AS Codigo, cAgeDescripcion AS Descripcion, 2  as Nivel  " _
    & " From " & lsServerComunes & "Agencias " & lsFiltro _
    & IIf(Not pbRaiz, "", "" _
    & " Union " _
    & " Select Convert(varchar(20),'0') as cCodAge , 'AGENCIAS' , 1  AS Nivel  " _
    & "" _
    & " ORDER BY Codigo ")


'sql = sql & " WHERE Convert(Varchar(20),cAgeCod) NOT IN ('04','11')"
Set rs = oConect.CargaRecordSet(sql)
Set GetAgencias = rs
oConect.CierraConexion
Set oConect = Nothing
End Function

Public Function GetAreas(Optional ByVal psObjetoFiltro As String = "") As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Dim lsFiltro As String
If psObjetoFiltro <> "" Then
    lsFiltro = " WHERE cAreaCod Like '" & psObjetoFiltro & "'  "
End If

Set oConect = New DConecta
Set rs = New ADODB.Recordset

If oConect.AbreConexion = False Then Exit Function

sql = " select Convert(varchar(20),'' + cAreaCod) as Codigo, cAreaDescripcion ,  LEN(cAreaEstruc) As Nivel from " & lsServerComunes & "Areas " & lsFiltro _
    & " Union " _
    & " Select Convert(varchar(20),'0') as cCodAge , 'AREAS' , 1  AS Nivel " _
    & "" _
    & "ORDER BY Codigo"

Set rs = oConect.CargaRecordSet(sql)
Set GetAreas = rs

oConect.CierraConexion
Set oConect = Nothing
End Function
Public Function GetNombreAgencia(ByVal psAgenciaCod As String) As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta

Set oConect = New DConecta
Set rs = New ADODB.Recordset

If oConect.AbreConexion = False Then Exit Function

sql = " Select cAgeCod, cAgeDescripcion from " & lsServerComunes & "Agencias where cAgeCod ='" & psAgenciaCod & "'"

GetNombreAgencia = ""
Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetNombreAgencia = rs!CAGEDESCRIPCION
End If
rs.Close
Set rs = Nothing

oConect.CierraConexion
Set oConect = Nothing
End Function
Public Function GetNombreAreas(ByVal psAreaCod As String) As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta

Set oConect = New DConecta
Set rs = New ADODB.Recordset

If oConect.AbreConexion = False Then Exit Function

sql = " Select cAreaCod, cAreaDescripcion from " & lsServerComunes & "Areas where cAreaCod ='" & psAreaCod & "'"

GetNombreAreas = ""
Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetNombreAreas = rs!cAreaDescripcion
End If
rs.Close
Set rs = Nothing

oConect.CierraConexion
Set oConect = Nothing
End Function
'Public Function GetAgenciasAreas(Optional ByVal psObjetoFiltro As String = "", Optional lnARendir As Long = -1) As ADODB.Recordset
'Dim sql As String
'Dim rs As ADODB.Recordset
'Dim oconect As DConecta
'Dim lsFiltro As String
'Dim lsFiltroArendir As String
'Set oconect = New DConecta
'Set rs = New ADODB.Recordset
'lsFiltro = ""
'If Trim(psObjetoFiltro) <> "" Then
'    lsFiltro = " WHERE Codigo LIKE '" & psObjetoFiltro & "'  "
'End If
'lsFiltroArendir = ""
'If lnARendir <> -1 Then
'    lsFiltroArendir = IIf(lsFiltro = "", "WHERE ", " and ") & " AgeArea.nArendir =" & lnARendir & ""
'End If
'If oconect.AbreConexion = False Then Exit Function
'
'sql = " SELECT Codigo, Descripcion, Nivel " _
'    & " From  (   Select    AA.cAreaCod as  Codigo , " _
'    & "                     A.cAreaDescripcion as Descripcion , " _
'    & "                     Len(A.cAreaEstruc) As Nivel, AA.nArendir " _
'    & "            From     " & lsServerComunes & "AreaAgencia AA " _
'    & "                     JOIN " & lsServerComunes & "Areas A on A.cAreaCod = AA.cAreaCod " _
'    & "            Union " _
'    & "            Select   AA.cAreaCod  + AA.cAgecod as  Codigo , " _
'    & "                     AG.cAgeDescripcion AS Descripcion, Len(A.cAreaEstruc+Ag.cAgeCod) As Nivel, AA.nArendir " _
'    & "            From     " & lsServerComunes & "AreaAgencia AA " _
'    & "                     JOIN " & lsServerComunes & "Areas A on A.cAreaCod = AA.cAreaCod " _
'    & "                     JOIN " & lsServerComunes & "Agencias AG on AA.cAgecod = Ag.cAgecod  ) AS AgeArea " _
'    & lsFiltro & lsFiltroArendir _
'    & " GROUP BY Codigo, Descripcion, Nivel  Order by Codigo, Nivel "
'
'Set rs = oconect.CargaRecordSet(sql)
'Set GetAgenciasAreas = rs
'
'oconect.CierraConexion
'Set oconect = Nothing
'End Function

Public Function GetAgenciasAreas(Optional ByVal psObjetoFiltro As String = "", Optional lnARendir As Long = -1, Optional pbOrganigrama As Boolean = False) As ADODB.Recordset
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim oConect As DConecta
    
    Set oConect = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConect.AbreConexion = False Then Exit Function

Dim lsFiltro As String
Dim lsFiltroArendir As String
lsFiltro = ""
If Trim(psObjetoFiltro) <> "" Then
    lsFiltro = " WHERE Codigo LIKE '" & psObjetoFiltro & "'  "
End If
lsFiltroArendir = ""
If lnARendir <> -1 Then
    lsFiltroArendir = IIf(lsFiltro = "", "WHERE ", " and ") & " AgeArea.nArendir =" & lnARendir & ""
End If
    
    If pbOrganigrama Then
        sql = " SELECT Codigo, Descripcion, Nivel " _
            & " From  (   Select    AA.cAreaCod as  Codigo , cAreaEstruc EST," _
            & "                     A.cAreaDescripcion as Descripcion , " _
            & "                     Len(A.cAreaEstruc) As Nivel, AA.nArendir " _
            & "            From     " & lsServerComunes & "AreaAgencia AA " _
            & "                     JOIN " & lsServerComunes & "Areas A on A.cAreaCod = AA.cAreaCod " _
            & "            Union " _
            & "            Select   AA.cAreaCod  + AA.cAgecod as  Codigo , cAreaEstruc EST," _
            & "                     AG.cAgeDescripcion AS Descripcion, 10 As Nivel, AA.nArendir " _
            & "            From     " & lsServerComunes & "AreaAgencia AA " _
            & "                     JOIN " & lsServerComunes & "Agencias AG on AA.cAgecod = Ag.cAgecod  JOIN " & lsServerComunes & "Areas ARE on ARE.cAreaCod = AA.cAreacod) AS AgeArea " _
            & lsFiltro & lsFiltroArendir _
            & " GROUP BY Codigo, Descripcion, Nivel, EST Order by EST, Nivel "
    Else
        sql = " SELECT Codigo, Descripcion, Nivel " _
            & " From  (   Select    AA.cAreaCod as  Codigo , A.nAgrupacion EST," _
            & "                     A.cAreaDescripcion as Descripcion , " _
            & "                     Len(AA.cAreaCod) As Nivel, AA.nArendir " _
            & "            From     " & lsServerComunes & "AreaAgencia AA " _
            & "                     JOIN " & lsServerComunes & "Areas A on A.cAreaCod = AA.cAreaCod " _
            & "            Union " _
            & "            Select   AA.cAreaCod  + AA.cAgecod as  Codigo , ARE.nAgrupacion EST," _
            & "                     AG.cAgeDescripcion AS Descripcion, 10 As Nivel, AA.nArendir " _
            & "            From     " & lsServerComunes & "AreaAgencia AA " _
            & "                     JOIN " & lsServerComunes & "Agencias AG on AA.cAgecod = Ag.cAgecod  JOIN " & lsServerComunes & "Areas ARE on ARE.cAreaCod = AA.cAreacod" _
            & "              Union" _
            & "              Select  right('0000' + rtrim(ltrim(str(CO.nConsValor))),4) Codigo , rtrim(ltrim(str(CO.nConsValor))) EST," _
            & "                      CO.cConsDescripcion AS Descripcion, 1 As Nivel, 0 nArendir " _
            & "              From     Constante CO" _
            & "         Where nConsCod Like '1043' And nConsCod <> nConsValor" _
            & " ) AS AgeArea " _
            & lsFiltro & lsFiltroArendir _
            & " GROUP BY Codigo, Descripcion, Nivel, EST Order by EST , Codigo, Nivel "
    End If
    
    Set rs = oConect.CargaRecordSet(sql)
    Set GetAgenciasAreas = rs
    
    oConect.CierraConexion
    Set oConect = Nothing
End Function

#If DebugMode Then
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

Public Function GetAreaAgenciasImporte(Optional ByVal psAreaCod As String = "", Optional psAgeCod As String = "") As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Dim lsFiltro As String

Set oConect = New DConecta
Set rs = New ADODB.Recordset

If oConect.AbreConexion = False Then Exit Function
lsFiltro = ""
If psAreaCod <> "" Then
   lsFiltro = lsFiltro & "WHERE aai.cAreaCod Like '" & psAreaCod & "' "
End If
If psAgeCod <> "" Then
   lsFiltro = lsFiltro & IIf(lsFiltro = "", "WHERE", "AND") & " aai.cAgeCod  Like '" & psAgeCod & "' "
End If

sql = " SELECT aai.cAreaCod, aai.cAgeCod, ISNULL(cAgeDescripcion,cAreaDescripcion) AS Descripcion,ISNULL(nCobertura,0) nCobertura " _
    & " FROM AreaAgenciaImporte aai JOIN Areas a ON a.cAreaCod = aai.cAreaCod LEFT JOIN Agencias ag ON ag.cAgeCod = aai.cAgeCod " & lsFiltro _
    & " ORDER BY aai.cAreaCod, aai.cAgeCod"
Set rs = oConect.CargaRecordSet(sql)
Set GetAreaAgenciasImporte = rs
oConect.CierraConexion
Set oConect = Nothing
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
End Sub
