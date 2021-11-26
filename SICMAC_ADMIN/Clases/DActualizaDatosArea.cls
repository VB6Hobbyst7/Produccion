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

'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902E901DA
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB902E90338
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

Public Sub Inicio(psServerComunes As String, psServerPersona As String, psServerAdministracion As String, psServerNegocio As String, psServerImagenes As String)
    lsServerComunes = psServerComunes
    lsServerPersona = psServerPersona
    lsServerAdministracion = psServerAdministracion
    lsServerNegocio = psServerNegocio
    lsServerImagenes = psServerImagenes
End Sub

'##ModelId=3AB902E902FC
Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
    #End If
End Sub

'Agrega una nuevo area Organizacional
'##ModelId=3A9423E501DA
Public Function AgregaAreaOrg(psAreaCod As String, psAreaDescripsion As String, psLogProcSeleccAut As String, cUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaAreaOrgErr

    'your code goes here...

    Exit Function
AgregaAreaOrgErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosArea:AgregaAreaOrg Method")
End Function

'Elimina Area Organizacional, la eliminacion es logica
'##ModelId=3A9423E5020C
Public Function EliminaAreaOrg(psAreaCod As String) As Boolean
    On Error GoTo EliminaAreaOrgErr

    'your code goes here...

    Exit Function
EliminaAreaOrgErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosArea:EliminaAreaOrg Method")
End Function

'Actualiza Area Organizacional
'##ModelId=3A9423E5023E
Public Function ModificaAreaOrg(psAreaCod As String, psAreaDescripcion As String, psLogProcSeleccAut As String, cUltimaActualizacion As String) As Boolean
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

Public Function GetAgencias1(Optional pbSoloAgencias As Boolean = False) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oconect As DConecta

Set oconect = New DConecta
Set rs = New ADODB.Recordset

If oconect.AbreConexion = False Then Exit Function

If pbSoloAgencias = False Then
    sql = " Select Convert(varchar(20),cAgeCod) AS Codigo, cAgeDescripcion AS Descripcion, 2  as Nivel  " _
        & " From Agencias Where nEstado = 1 " _
        & " Union " _
        & " Select Convert(varchar(20),'0') as cCodAge , 'AGENCIAS' , 1  AS Nivel  ORDER BY Codigo "
Else
    sql = " Select cAgeCod, cAgeDescripcion AS Descripcion From Agencias Where nEstado = 1 ORDER BY cAgeCod"
End If

Set rs = oconect.CargaRecordSet(sql)
Set GetAgencias1 = rs
oconect.CierraConexion
Set oconect = Nothing
End Function


Public Function GetAgencias(Optional ByVal psObjetoFiltro As String = "", Optional pbRaiz As Boolean = True, Optional pbEstadoActivo As Boolean = False) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oconect As DConecta
Dim lsFiltro As String

Set oconect = New DConecta
Set rs = New ADODB.Recordset

If oconect.AbreConexion = False Then Exit Function
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


'sql = sql & " WHERE Convert(Varchar(20),cAgeCod) NOT IN ('04','11','02')"
Set rs = oconect.CargaRecordSet(sql)
Set GetAgencias = rs
oconect.CierraConexion
Set oconect = Nothing
End Function

Public Function GetNomArea(psAreaCod As String) As String
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim oconect As DConecta
    
    Set oconect = New DConecta
    Set rs = New ADODB.Recordset
    
    If oconect.AbreConexion = False Then Exit Function
    
    sql = " Select cAreaDescripcion From Areas Where cAreaCod = '" & psAreaCod & "'"
    Set rs = oconect.CargaRecordSet(sql)
    
    GetNomArea = rs.Fields(0)
    oconect.CierraConexion
    Set oconect = Nothing
End Function

Public Function GetNomAge(psAreaCod As String) As String
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim oconect As DConecta
    Set oconect = New DConecta
    Set rs = New ADODB.Recordset
    
    If oconect.AbreConexion = False Then Exit Function
    
    sql = " Select cAgeDescripcion From Agencias Where cAgeCod = '" & psAreaCod & "'"
    Set rs = oconect.CargaRecordSet(sql)
    
    GetNomAge = rs.Fields(0)
    oconect.CierraConexion
    Set oconect = Nothing
End Function

Public Function GetAreas() As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oconect As DConecta

Set oconect = New DConecta
Set rs = New ADODB.Recordset

If oconect.AbreConexion = False Then Exit Function

sql = " select Convert(varchar(20),'' + cAreaCod) as Codigo, ltrim(str(nagrupacion)) + ' '  +  cAreaDescripcion Descrip ,  LEN(cAreaCod) As Nivel from Areas " _
    & " Union " _
    & " Select Convert(varchar(20),'0' + ltrim(str(nConsValor))) as cCodAge , ltrim(str(nConsvalor)) + ' ' + cConsDescripcion Descrip , 2  AS Nivel  from constante where nconscod = 1043 and nconscod <> nconsValor " _
    & " Union " _
    & " Select Convert(varchar(20),'') as cCodAge , '0 AREAS' Descrip , 1  AS Nivel " _
    & " ORDER BY Descrip, nivel "

Set rs = oconect.CargaRecordSet(sql)
Set GetAreas = rs

oconect.CierraConexion
Set oconect = Nothing
End Function


Public Function GetNombreAgencia(ByVal psAgenciaCod As String) As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim oconect As DConecta

Set oconect = New DConecta
Set rs = New ADODB.Recordset

If oconect.AbreConexion = False Then Exit Function

sql = " Select cAgeCod, cAgeDescripcion from Agencias where cAgeCod ='" & psAgenciaCod & "'"

GetNombreAgencia = ""
Set rs = oconect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetNombreAgencia = rs!cAgeDescripcion
End If
rs.Close
Set rs = Nothing

oconect.CierraConexion
Set oconect = Nothing
End Function

Public Function GetDirAreaAgencia(Optional psAreaCod As String = "", Optional psAgenciaCod As String = "") As String
    Dim sql As String
    Dim oconect As DConecta
    Set oconect = New DConecta
    
    If psAreaCod <> "" Then
        sql = " Select cAgeDireccion from Agencias Where cAgeCod In" _
            & " (Select cUbicaCod from areaagencia where cAreaCod = '" & psAreaCod & "' And cAgeCod = '" & psAgenciaCod & "')"
    Else
        sql = " Select cAgeDireccion from Agencias Where cAgeCod = '" & psAgenciaCod & "'"
    End If
    
    If oconect.AbreConexion = False Then Exit Function
    GetDirAreaAgencia = oconect.CargaRecordSet(sql).Fields(0)
    
    oconect.CierraConexion
    Set oconect = Nothing
End Function

'Public Function GetAgenciasAreas(Optional pbOrganigrama As Boolean = False) As ADODB.Recordset
'    Dim sql As String
'    Dim rs As ADODB.Recordset
'    Dim oConect As DConecta
'
'    Set oConect = New DConecta
'    Set rs = New ADODB.Recordset
'
'    If oConect.AbreConexion = False Then Exit Function
'
'    If pbOrganigrama Then
'        sql = " SELECT Codigo, Descripcion, Nivel " _
'            & " From  (   Select    AA.cAreaCod as  Codigo , cAreaEstruc EST," _
'            & "                     A.cAreaDescripcion as Descripcion , " _
'            & "                     Len(A.cAreaEstruc) As Nivel " _
'            & "            From     " & lsServerComunes & "AreaAgencia AA " _
'            & "                     JOIN " & lsServerComunes & "Areas A on A.cAreaCod = AA.cAreaCod " _
'            & "            Union " _
'            & "            Select   AA.cAreaCod  + AA.cAgecod as  Codigo , cAreaEstruc EST," _
'            & "                     AG.cAgeDescripcion AS Descripcion, 10 As Nivel " _
'            & "            From     " & lsServerComunes & "AreaAgencia AA " _
'            & "                     JOIN " & lsServerComunes & "Agencias AG on AA.cAgecod = Ag.cAgecod  JOIN " & lsServerComunes & "Areas ARE on ARE.cAreaCod = AA.cAreacod) AS AgeArea " _
'            & " Order by EST, Nivel "
'    Else
'        sql = " SELECT Codigo, Descripcion, Nivel  " _
'            & " From  (" _
'            & "         Select    AA.cAreaCod as  Codigo , A.nAgrupacion EST," _
'            & "                      A.cAreaDescripcion as Descripcion ," _
'            & "                       Len(AA.cAreaCod) As Nivel" _
'            & "              From     AreaAgencia AA" _
'            & "                       JOIN Areas A on A.cAreaCod = AA.cAreaCod" _
'            & "              Union" _
'            & "              Select   AA.cAreaCod  + AA.cAgecod as  Codigo , ARE.nAgrupacion EST," _
'            & "                      AG.cAgeDescripcion AS Descripcion, 10 As Nivel" _
'            & "              From     AreaAgencia AA" _
'            & "                       JOIN Agencias AG on AA.cAgecod = Ag.cAgecod" _
'            & "               JOIN Areas ARE on ARE.cAreaCod = AA.cAreacod" _
'            & "              Union" _
'            & "              Select  right('0000' + rtrim(ltrim(str(CO.nConsValor))),4) Codigo , rtrim(ltrim(str(CO.nConsValor))) EST," _
'            & "                      CO.cConsDescripcion AS Descripcion, 1 As Nivel" _
'            & "              From     Constante CO" _
'            & "         Where nConsCod Like '1043' And nConsCod <> nConsValor" _
'            & " ) AS AgeArea  Order by EST, Codigo, Nivel"
'    End If
'
'    Set rs = oConect.CargaRecordSet(sql)
'    Set GetAgenciasAreas = rs
'
'    oConect.CierraConexion
'    Set oConect = Nothing
'End Function

Public Function GetCargosAreas() As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oconect As DConecta

Set oconect = New DConecta
Set rs = New ADODB.Recordset

If oconect.AbreConexion = False Then Exit Function

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

Set rs = oconect.CargaRecordSet(sql)
Set GetCargosAreas = rs

oconect.CierraConexion
Set oconect = Nothing
End Function

Public Function GetAreasRRHHEvaluacion(psRHEvaluacionNro As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oconect As DConecta

Set oconect = New DConecta
Set rs = New ADODB.Recordset

If oconect.AbreConexion = False Then Exit Function

sql = " Select AR.cAreaCod Codigo, cAreaDescripcion, len(AR.cAreaCod) Nivel " _
    & " From Areas AR" _
    & "     Inner Join RHEvaluacionDet RHE" _
    & "         On AR.cAreaCod = RHE.cAreaCod" _
    & "             Where cRHEvaluacionNro = '" & psRHEvaluacionNro & "'" _
    & " Union" _
    & " select RH.cAreaCod + RH.cPersCod Codigo, PE.cPersNombre, len(RH.cPersCod) Nivel" _
    & " From RRHH RH" _
    & "     Inner Join RHEvaluacionDet RHE" _
    & "         On RH.cAreaCod = RHE.cAreaCod" _
    & "     Inner Join Persona PE          " _
    & "         On RH.cPersCod = PE.cPersCod And dCese is null" _
    & "              Where cRHEvaluacionNro = '" & psRHEvaluacionNro & "' order by codigo"

Set rs = oconect.CargaRecordSet(sql)
Set GetAreasRRHHEvaluacion = rs

oconect.CierraConexion
Set oconect = Nothing
End Function

#If DebugMode Then
    '##ModelId=3AB902E90234
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

'Public Function GetAgenciasAreas(Optional ByVal psObjetoFiltro As String = "", Optional lnARendir As Long = -1, Optional pbOrganigrama As Boolean = False) As ADODB.Recordset
'    Dim sql As String
'    Dim rs As ADODB.Recordset
'    Dim oconect As DConecta
'
'    Set oconect = New DConecta
'    Set rs = New ADODB.Recordset
'
'    If oconect.AbreConexion = False Then Exit Function
'
'Dim lsFiltro As String
'Dim lsFiltroArendir As String
'lsFiltro = ""
'If Trim(psObjetoFiltro) <> "" Then
'    lsFiltro = " WHERE Codigo LIKE '" & psObjetoFiltro & "'  "
'End If
'lsFiltroArendir = ""
'If lnARendir <> -1 Then
'    lsFiltroArendir = IIf(lsFiltro = "", "WHERE ", " and ") & " AgeArea.nArendir =" & lnARendir & ""
'End If
'
'    If pbOrganigrama Then
'        sql = " SELECT Codigo, Descripcion, Nivel " _
'            & " From  (   Select    AA.cAreaCod as  Codigo , cAreaEstruc EST," _
'            & "                     A.cAreaDescripcion as Descripcion , " _
'            & "                     Len(A.cAreaEstruc) As Nivel, AA.nArendir " _
'            & "            From     " & lsServerComunes & "AreaAgencia AA " _
'            & "                     JOIN " & lsServerComunes & "Areas A on A.cAreaCod = AA.cAreaCod Where A.nAreaEstado = " & AreasEstado.AreasEstadoActivo & "" _
'            & "            Union " _
'            & "            Select   AA.cAreaCod  + AA.cAgecod as  Codigo , cAreaEstruc EST," _
'            & "                     AG.cAgeDescripcion AS Descripcion, 10 As Nivel, AA.nArendir " _
'            & "            From     " & lsServerComunes & "AreaAgencia AA " _
'            & "                     JOIN " & lsServerComunes & "Agencias AG on AA.cAgecod = Ag.cAgecod  JOIN " & lsServerComunes & "Areas ARE on ARE.cAreaCod = AA.cAreacod Where ARE.nAreaEstado = " & AreasEstado.AreasEstadoActivo & ") AS AgeArea Where " _
'            & lsFiltro & lsFiltroArendir _
'            & " GROUP BY Codigo, Descripcion, Nivel, EST Order by EST, Nivel "
'    Else
'        sql = " SELECT Codigo, Descripcion, Nivel " _
'            & " From  (   Select    AA.cAreaCod as  Codigo , A.nAgrupacion EST," _
'            & "                     A.cAreaDescripcion as Descripcion , " _
'            & "                     Len(AA.cAreaCod) As Nivel, AA.nArendir " _
'            & "            From     " & lsServerComunes & "AreaAgencia AA " _
'            & "                     JOIN " & lsServerComunes & "Areas A on A.cAreaCod = AA.cAreaCod Where A.nAreaEstado = " & AreasEstado.AreasEstadoActivo & " And A.bVisible = 1 And AA.bVisible = 1" _
'            & "            Union " _
'            & "            Select   AA.cAreaCod  + AA.cAgecod as  Codigo , ARE.nAgrupacion EST," _
'            & "                     AG.cAgeDescripcion AS Descripcion, 10 As Nivel, AA.nArendir " _
'            & "            From     " & lsServerComunes & "AreaAgencia AA " _
'            & "                     JOIN " & lsServerComunes & "Agencias AG on AA.cAgecod = Ag.cAgecod  JOIN " & lsServerComunes & "Areas ARE On ARE.cAreaCod = AA.cAreacod  Where ARE.nAreaEstado = " & AreasEstado.AreasEstadoActivo & " And AA.bVisible = 1" _
'            & "              Union" _
'            & "              Select  right('0000' + rtrim(ltrim(str(CO.nConsValor))),4) Codigo , rtrim(ltrim(str(CO.nConsValor))) EST," _
'            & "                      CO.cConsDescripcion AS Descripcion, 1 As Nivel, 0 nArendir " _
'            & "              From     Constante CO" _
'            & "         Where nConsCod Like " & gAreasAgrupacion & " And nConsCod <> nConsValor" _
'            & " ) AS AgeArea " _
'            & lsFiltro & lsFiltroArendir _
'            & " GROUP BY Codigo, Descripcion, Nivel, EST Order by EST , Codigo, Nivel "
'    End If
'
'    Set rs = oconect.CargaRecordSet(sql)
'    Set GetAgenciasAreas = rs
'
'    oconect.CierraConexion
'    Set oconect = Nothing
'End Function
'

Public Function GetAgenciasAreas(Optional ByVal psObjetoFiltro As String = "", Optional lnARendir As Long = -1, Optional pbOrganigrama As Boolean = False) As ADODB.Recordset
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim oconect As DConecta
    
    Set oconect = New DConecta
    Set rs = New ADODB.Recordset
    
    If oconect.AbreConexion = False Then Exit Function

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
            & "                     JOIN " & lsServerComunes & "Areas A on A.cAreaCod = AA.cAreaCod Where A.nAreaEstado = " & AreasEstado.AreasEstadoActivo & "" _
            & "            Union " _
            & "            Select   AA.cAreaCod  + AA.cAgecod as  Codigo , cAreaEstruc EST," _
            & "                     AG.cAgeDescripcion AS Descripcion, 10 As Nivel, AA.nArendir " _
            & "            From     " & lsServerComunes & "AreaAgencia AA " _
            & "                     JOIN " & lsServerComunes & "Agencias AG on AA.cAgecod = Ag.cAgecod  JOIN " & lsServerComunes & "Areas ARE on ARE.cAreaCod = AA.cAreacod Where ARE.nAreaEstado = " & AreasEstado.AreasEstadoActivo & ") AS AgeArea Where " _
            & lsFiltro & lsFiltroArendir _
            & " GROUP BY Codigo, Descripcion, Nivel, EST Order by EST, Nivel "
    Else
        sql = " SELECT Codigo, Descripcion, Nivel " _
            & " From  (   Select    AA.cAreaCod as  Codigo , A.nAgrupacion EST," _
            & "                     A.cAreaDescripcion as Descripcion , " _
            & "                     Len(AA.cAreaCod) As Nivel, AA.nArendir " _
            & "            From     " & lsServerComunes & "AreaAgencia AA " _
            & "                     JOIN " & lsServerComunes & "Areas A on A.cAreaCod = AA.cAreaCod Where A.nAreaEstado = " & AreasEstado.AreasEstadoActivo & " " _
            & "            Union " _
            & "            Select   AA.cAreaCod  + AA.cAgecod as  Codigo , ARE.nAgrupacion EST," _
            & "                     AG.cAgeDescripcion AS Descripcion, 10 As Nivel, AA.nArendir " _
            & "            From     " & lsServerComunes & "AreaAgencia AA " _
            & "                     JOIN " & lsServerComunes & "Agencias AG on AA.cAgecod = Ag.cAgecod  JOIN " & lsServerComunes & "Areas ARE On ARE.cAreaCod = AA.cAreacod  Where ARE.nAreaEstado = " & AreasEstado.AreasEstadoActivo & "" _
            & "              Union" _
            & "              Select  right('0000' + rtrim(ltrim(str(CO.nConsValor))),4) Codigo , rtrim(ltrim(str(CO.nConsValor))) EST," _
            & "                      CO.cConsDescripcion AS Descripcion, 1 As Nivel, 0 nArendir " _
            & "              From     Constante CO" _
            & "         Where nConsCod Like " & gAreasAgrupacion & " And nConsCod <> nConsValor" _
            & " ) AS AgeArea " _
            & lsFiltro & lsFiltroArendir _
            & " GROUP BY Codigo, Descripcion, Nivel, EST Order by EST , Codigo, Nivel "
    End If
    
    Set rs = oconect.CargaRecordSet(sql)
    Set GetAgenciasAreas = rs
    
    oconect.CierraConexion
    Set oconect = Nothing
End Function
