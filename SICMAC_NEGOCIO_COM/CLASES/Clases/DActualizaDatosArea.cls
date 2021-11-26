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

'Realiza el mantenimiento de las Areas Organizacionales
Option Base 0
Option Explicit


Private lsServerComunes As String
Private lsServerPersona As String
Private lsServerAdministracion As String
Private lsServerNegocio As String
Private lsServerImagenes As String

Public Sub inicio(psServerComunes As String, psServerPersona As String, psServerAdministracion As String, psServerNegocio As String, psServerImagenes As String)
    lsServerComunes = psServerComunes
    lsServerPersona = psServerPersona
    lsServerAdministracion = psServerAdministracion
    lsServerNegocio = psServerNegocio
    lsServerImagenes = psServerImagenes
End Sub
'Agrega una nuevo area Organizacional
'##ModelId=3A9423E501DA
Public Function AgregaAreaOrg(pcAreaCod As String, pcAreaDescripcion As String, pcLogProcSeleccAut As String, cUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaAreaOrgErr

    'your code goes here...

    Exit Function
AgregaAreaOrgErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosArea:AgregaAreaOrg Method")
End Function

'Elimina Area Organizacional, la eliminacion es logica
'##ModelId=3A9423E5020C
Public Function EliminaAreaOrg(pcAreaCod As String) As Boolean
    On Error GoTo EliminaAreaOrgErr

    'your code goes here...

    Exit Function
EliminaAreaOrgErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosArea:EliminaAreaOrg Method")
End Function

'Actualiza Area Organizacional
'##ModelId=3A9423E5023E
Public Function ModificaAreaOrg(pcAreaCod As String, pcAreaDescripcion As String, pcLogProcSeleccAut As String, cUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaAreaOrgErr

    'your code goes here...

    Exit Function
ModificaAreaOrgErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosArea:ModificaAreaOrg Method")
End Function

'Devuelve las Areas organizacionales
'##ModelId=3A9424210050
Public Function GetAreasOrg() As Recordset
    Dim oCon As DConecta
    Dim sqlA As String
    Set oCon = New DConecta
    On Error GoTo GetAreasOrgErr
    
    sqlA = "Select cAreaDescripcion, cAreaCod from areas"
    If oCon.AbreConexion Then
        Set GetAreasOrg = oCon.CargaRecordSet(sqlA)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetAreasOrgErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosArea:GetAreasOrg Method")
End Function

Public Function GetAgencias(Optional ByVal psObjetoFiltro As String = "", Optional pbRaiz As Boolean = True) As ADODB.Recordset
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

sql = " Select Convert(Varchar(20),cAgeCod) AS Codigo, cAgeDescripcion AS Descripcion, 2  as Nivel  " _
    & " From " & lsServerComunes & "Agencias " & lsFiltro _
    & IIf(Not pbRaiz, "", "" _
    & " Union " _
    & " Select Convert(varchar(20),'0') as cCodAge , 'AGENCIAS' , 1  AS Nivel  " _
    & "" _
    & " ORDER BY Codigo ")

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
    GetNombreAgencia = rs!cAgeDescripcion
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
    GetNombreAreas = rs!CAREADESCRIPCION
End If
rs.Close
Set rs = Nothing

oConect.CierraConexion
Set oConect = Nothing
End Function
Public Function GetAgenciasAreas(Optional ByVal psObjetoFiltro As String = "", Optional lnARendir As Long = -1) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Dim lsFiltro As String
Dim lsFiltroArendir As String
Set oConect = New DConecta
Set rs = New ADODB.Recordset
lsFiltro = ""
If Trim(psObjetoFiltro) <> "" Then
    lsFiltro = " WHERE Codigo LIKE '" & psObjetoFiltro & "'  "
End If
lsFiltroArendir = ""
If lnARendir <> -1 Then
    lsFiltroArendir = IIf(lsFiltro = "", "WHERE ", " and ") & " AgeArea.nArendir =" & lnARendir & ""
End If
If oConect.AbreConexion = False Then Exit Function
   
sql = " SELECT Codigo, Descripcion, Nivel " _
    & " From  (   Select    AA.cAreaCod as  Codigo , " _
    & "                     A.cAreaDescripcion as Descripcion , " _
    & "                     Len(A.cAreaEstruc) As Nivel, AA.nArendir " _
    & "            From     " & lsServerComunes & "AreaAgencia AA " _
    & "                     JOIN " & lsServerComunes & "Areas A on A.cAreaCod = AA.cAreaCod " _
    & "            Union " _
    & "            Select   AA.cAreaCod  + AA.cAgecod as  Codigo , " _
    & "                     AG.cAgeDescripcion AS Descripcion, 10 As Nivel , AA.nArendir " _
    & "            From     " & lsServerComunes & "AreaAgencia AA " _
    & "                     JOIN " & lsServerComunes & "Agencias AG on AA.cAgecod = Ag.cAgecod  ) AS AgeArea " _
    & lsFiltro & lsFiltroArendir _
    & " GROUP BY Codigo, Descripcion, Nivel  Order by Codigo, Nivel "
 
Set rs = oConect.CargaRecordSet(sql)
Set GetAgenciasAreas = rs

oConect.CierraConexion
Set oConect = Nothing
End Function
Public Function GetCargosAreas() As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta

Set oConect = New DConecta
Set rs = New ADODB.Recordset

If oConect.AbreConexion = False Then Exit Function

sql = " SELECT Codigo, Descripcion, Nivel " _
    & " From (" _
    & "         Select     " _
    & "         Case                  WHEN AA.cAreaCod = '' THEN '00'" _
    & "                               Else AA.cAreaCod" _
    & "                               END AS Codigo, ISNULL(AG.cAreaDescripcion,'AREAS INDEPENDIENTES') AS Descripcion, 1 As Nivel      " _
    & " From        AreaCargo AA" _
    & " LEFT JOIN Areas AG on AA.cAreaCod = Ag.cAreaCod Where AG.nAreaEstado = " & AreasEstado.AreasEstadoActivo & " " _
    & " Union" _
    & " Select  Case      WHEN    AA.cRHCargoCod = '' THEN '00' + AA.cRHCargoCod" _
    & " Else AA.cAreaCod + AA.cRHCargoCod" _
    & " END AS Codigo , ISNULL(AG.cRHCargoDescripcion,'INDEPENDIENTE') as Descripcion" _
    & " , Len(AG.cRHCargoCod) As Nivel      From    AreaCargo AA" _
    & " JOIN rhcargostabla AG on AA.cRHCargoCod = AG.cRHCargoCod) AS AgeArea " _
    & " Order by Codigo, Nivel"

Set rs = oConect.CargaRecordSet(sql)
Set GetCargosAreas = rs

oConect.CierraConexion
Set oConect = Nothing
End Function






#If DebugMode Then
    '##ModelId=3AB902E90234
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If





