VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DLogProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim vsConexion As String
Dim vsCentralPer As String
Dim vsCentralCom As String
Dim vsCentralImg As String
Dim sSQL As String
Dim sSQLProvAG As String

Public Enum TpoProBS
    ProBSBienServicio = 1
    ProBSProveedor = 2
End Enum

Private Sub Class_Initialize()
    'Dim oImp As DImpresoras
    'Set oImp = New DImpresoras
    
   'oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    'Set oImp = Nothing

    'Dim oIni As ClasIni
    
    'Set oIni = New ClasIni
        'vsConexion = oIni.CadenaConexion
        'vsCentralPer = oIni.BasePersonas
        'vsCentralCom = oIni.BaseComunes
        'vsCentralImg = oIni.BaseImagenes
    'Set oIni = Nothing
End Sub

'Devuelve un valor si existe proveedor
Public Function IsExisProveedor(ByVal psPersCod As String, Optional pbConEstado As Boolean = False) As Boolean
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        If pbConEstado = False Then
            sSQL = "SELECT cPersCod FROM Proveedor " & _
                " WHERE  cPersCod = '" & psPersCod & "' "
        Else
            sSQL = "SELECT cPersCod FROM Proveedor " & _
                " WHERE  cPersCod = '" & psPersCod & "' AND nProvEstado = " & gLogProvEstadoActivado & " "
        End If
        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
    End If
    IsExisProveedor = IIf(rs.RecordCount > 0, True, False)
End Function

'Devuelva un RecordSet con los proveedores
Public Function CargaProveedor() As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
    'INICIO ORCR 20140829****************************
'        sSql = "SELECT pr.cPersCod Código, p.cPersNombre Nombre, p.cPersDireccDomicilio Dirección, " & _
'            " pi1.cPersIdNro RUC, pi2.cPersIdNro RUS, pr.nProvEstado , c.cConsDescripcion Estado, PR.cCtaCodMN Cta_MN, PR.cCtaCodME Cta_ME, ISNull((Select Max(PBC.cPersBCRuc) From PersID PID Inner Join PersSunatBC PBC On PID.cPersIDnro = PBC.cPersBCRuc And PID.cPersIDTpo = " & PersIdTipo.gPersIdRUC & " Where PID.cPersCod = pr.cPersCod),0) bAgenteRetencion,  Case When (Select Max(PBC.cPersBCRuc) From PersID PID Inner Join PersSunatBC PBC On PID.cPersIDnro = PBC.cPersBCRuc And PID.cPersIDTpo = " & PersIdTipo.gPersIdRUC & " Where PID.cPersCod = pr.cPersCod) Is Null Then 'No es Agente_Ret/BC' Else 'Agente_Ret/BC' End  [A.Ret/B.Cont] " & _
'            " FROM Proveedor PR JOIN " & vsCentralPer & "Persona P ON pr.cPersCod = p.cPersCod " & _
'            " LEFT JOIN " & vsCentralPer & "PersID PI1 ON pr.cPersCod = pi1.cPersCod AND pi1.cPersIdTpo = '" & gPersIdRUC & "'" & _
'            " LEFT JOIN " & vsCentralPer & "PersID PI2 ON pr.cPersCod = pi2.cPersCod AND pi2.cPersIdTpo = '" & gPersIdRUS & "'" & _
'            " LEFT JOIN " & vsCentralCom & "Constante C ON pr.nProvEstado = c.nConsValor AND c.nConsCod = " & gLogProvEstado & " " & _
'            "  " & _
'            "  " & _
'            " ORDER BY p.cPersNombre "
sSQL = "exec stp_sel_CargarProveedores"
'FIN ORCR 20140829****************************
        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
    End If
    
    Set CargaProveedor = rs
End Function

'Devuelve un RecordSet con los Bienes/Servicios que brinda un proveedor
'Devuelve un RecordSet con los Bienes/Servicios que brinda un proveedor
Public Function CargaProveedorBS(ByVal psTipo As TpoProBS, ByVal psCodigo As String, Optional pbBBNivelSuperior As Boolean = False) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    Dim nLong As Integer
    Dim cRaiz As String

    sSQL = ""
    If oConec.AbreConexion() Then
        If psTipo = ProBSBienServicio Then
            sSQL = "SELECT pbs.cBSCod Codigo1, pbs.cBSCod Codigo, bs.cBSDescripcion Descripción " & _
                " FROM ProveedorBS PBS with (nolock) JOIN BienesServicios BS with (nolock) ON pbs.cBSCod = bs.cBSCod " & _
                "       AND pbs.cPersCod = '" & psCodigo & "'" & _
                " Union " _
                & " Select c.cCtaContCod Codigo1, c.cCtaContCod Codigo, Descripcion=x.cCtaContDesc+' - '+ y.cAgeDescripcion+' - '+ b.cMovDesc " _
                & "  from movgasto a inner join mov b on a.nmovnro = b.nmovnro inner join movcta c on a.nmovnro = c.nmovnro " _
                & "       inner join CtaCont x on left(c.cCtaContCod,len(c.cCtaContCod)-2) = x.cCtaContCod " _
                & "       inner join Agencias y on right(c.cCtaContCod,2) = y.cAgeCod " _
                & " where cperscod = '" & psCodigo & "' And nmovflag <> 1 and nmovestado = 10 And cOpecod like '70_40_' and c.cctacontcod like '4%' Order by Codigo1 "

                'And Len(Trim(psCodigo)) > 2
        ElseIf psTipo = ProBSProveedor Then
               nLong = Len(psCodigo)
               If nLong = 1 Then
                  cRaiz = psCodigo
               Else
                  If nLong > 3 Then
                     cRaiz = Left(psCodigo, nLong - 3)
                  End If
               End If
                sSQL = "SELECT  distinct p.cPersCod Codigo, p.cPersNombre Nombre, p.cPersDireccDomicilio Domicilio, cPersTelefono Fono1, cPersTelefono2 Fono2, cPersEmail EMAIL " & _
                " FROM Persona P with (nolock) JOIN  ProveedorBS PBS with (nolock) ON p.cPersCod = pbs.cPersCod " & _
                " WHERE pbs.cBSCod Like '" & IIf(Not pbBBNivelSuperior, psCodigo, cRaiz) & "%'"
        End If
        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
    End If

    Set CargaProveedorBS = rs
End Function

'Public Function CargaProveedorBS(ByVal psTipo As TpoProBS, _
'ByVal psCodigo As String) As ADODB.Recordset
'    Dim oConec As DConecta
'    Dim rs As ADODB.Recordset
'    Set oConec = New DConecta
'    Set rs = New ADODB.Recordset
'
'    If oConec.AbreConexion() Then
'        If psTipo = ProBSBienServicio Then
'            sSql = "SELECT pbs.cBSCod Codigo, pbs.cBSCod Codigo, bs.cBSDescripcion Descripción " & _
'                " FROM ProveedorBS PBS JOIN BienesServicios BS ON pbs.cBSCod = bs.cBSCod " & _
'                "       AND pbs.cPersCod = '" & psCodigo & "'" & _
'                " ORDER BY pbs.cBSCod "
'        ElseIf psTipo = ProBSProveedor Then
'            sSql = "SELECT p.cPersCod, p.cPersNombre, p.cPersDireccDomicilio " & _
'                " FROM Persona P JOIN  ProveedorBS PBS ON p.cPersCod = pbs.cPersCod " & _
'                " WHERE pbs.cBSCod = SUBSTRING('" & psCodigo & "',1,LEN(pbs.cBSCod))"
'        End If
'        Set rs = oConec.CargaRecordSet(sSql)
'        oConec.CierraConexion
'    End If
'
'    Set CargaProveedorBS = rs
'End Function


'Activa y Desactiva estado del Proveedor
Public Function ActDesProveedor(ByVal psSql As String) As Integer
    Dim oConec As DConecta
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
        oConec.Ejecutar psSql
        oConec.CierraConexion
        
        ActDesProveedor = 0
    Else
        ActDesProveedor = 1
    End If
End Function

'Elimina un proveedor y su relación con los bienes/servicios que brinda
Public Function EliminaProveedor(ByVal psPersCod As String) As Integer
    Dim oConec As DConecta
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
        sSQL = "DELETE ProveedorBS WHERE cPersCod = '" & psPersCod & "'"
        oConec.Ejecutar sSQL
        
        sSQL = "DELETE Proveedor WHERE cPersCod = '" & psPersCod & "'"
        oConec.Ejecutar sSQL
        
        oConec.CierraConexion
        EliminaProveedor = 0
    Else
        EliminaProveedor = 1
    End If
End Function

'Elimina el bien/servicio que presta un proveedor
Public Function EliminaProveedorBS(ByVal psPersCod As String, ByVal psBSCod As String) As Integer
    Dim oConec As DConecta
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
        sSQL = "DELETE ProveedorBS WHERE cPersCod = '" & psPersCod & "'" & _
            " AND cBSCod = '" & psBSCod & "'"
        oConec.Ejecutar sSQL
        
        oConec.CierraConexion
        EliminaProveedorBS = 0
    Else
        EliminaProveedorBS = 1
    End If
End Function

'Agrega un proveedor/y a proveedorAG
'Public Function GrabaProveedor(ByVal psPersCod As String, ByVal psActualizacion As String, _
'Optional strCAgeCod As String) As Integer
'    Dim sActualiza As String
'    Dim oConec As DConecta
'    Dim rs As Recordset
'
'    Set oConec = New DConecta
'
'    If oConec.AbreConexion() Then
'
'        sSql = "SELECT cPersCod FROM Proveedor WHERE cPersCod = '" & psPersCod & "' "
'        Set rs = oConec.CargaRecordSet(sSql)
'
'        If (rs.BOF And rs.EOF) Then
'            sActualiza = Left(psActualizacion, 8) & Format(oConec.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
'            sSql = " INSERT INTO Proveedor(cPersCod, nProvEstado, cUltimaActualizacion,bhabilSUNAT, " & _
'                   " bhabilCONSUCODE,cComentario) " & _
'                   " VALUES ('" & psPersCod & "'," & gLogProvEstadoActivado & ",'" & sActualiza & "',1,1,'')"
'            oConec.Ejecutar sSql
'            GrabaProveedor = 0
'        End If
'
'        sSQLProvAG = " if not exists(select cPersCod from ProveedorAG where cPersCod='" & psPersCod & "' " & _
'                     " and cAgeCod='" & strCAgeCod & "') begin insert into proveedorAG (cPersCod,cAgeCod,cUltimaActualizacion) " & _
'                     " values ('" & psPersCod & "','" & strCAgeCod & "','" & sActualiza & "') end"
'        oConec.Ejecutar sSQLProvAG
'
'        GrabaProveedor = 2
'        oConec.CierraConexion
'
'    Else
'        GrabaProveedor = 1
'    End If
'End Function

Public Function GrabaProveedor(ByVal psPersCod As String, ByVal psActualizacion As String) As Integer
    Dim sActualiza As String
    Dim oConec As DConecta
    Dim rs As Recordset

    Set oConec = New DConecta

    If oConec.AbreConexion() Then
        sSQL = "SELECT cPersCod FROM Proveedor " & _
            " WHERE cPersCod = '" & psPersCod & "' "
        Set rs = oConec.CargaRecordSet(sSQL)

        If rs.RecordCount = 0 Then
            sActualiza = Left(psActualizacion, 8) & Format(oConec.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)

            sSQL = "INSERT INTO Proveedor(cPersCod, nProvEstado, cUltimaActualizacion) " & _
                " VALUES ('" & psPersCod & "'," & gLogProvEstadoActivado & ",'" & sActualiza & "')"
            oConec.Ejecutar sSQL
            GrabaProveedor = 0
        Else
            GrabaProveedor = 2
        End If
        oConec.CierraConexion
    Else
        GrabaProveedor = 1
    End If
End Function

Public Function CargaProveedores(ByVal AccesoLogistica As Integer, ByVal gsCodAgeN As String, _
ByVal sValorBusqueda As String, Optional numFilas As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    Select Case AccesoLogistica
    Case Is = 0
    'PR.bHabilCONSUCODE , PR.cComentario
    If oConec.AbreConexion() Then
    'case when & numFilas &<=10 then ROWCOUNT=0 when ROWCOUNT=" & numFilas & "
    'if numFilas<=10 begin ROWCOUNT=0 end else ROWCOUNT=" & numFilas & "
    'IF " & numFilas & " <> 10 BEGIN SET ROWCOUNT " & numFilas & " END ELSE SET ROWCOUNT 10
    
        'sSQL = " SELECT pr.cPersCod Código, p.cPersNombre Nombre, p.cPersDireccDomicilio Dirección, " & _
            " pi1.cPersIdNro RUC, P.cPersTelefono Telefono, pr.nProvEstado , c.cConsDescripcion Estado, " & _
            " PR.cCtaCodMN Cta_MN, PR.cCtaCodME Cta_ME, ISNull((Select Max(PBC.cPersBCRuc) " & _
            " From PersID PID Inner Join PersSunatBC PBC On PID.cPersIDnro = PBC.cPersBCRuc " & _
            " And PID.cPersIDTpo = 2 Where PID.cPersCod = pr.cPersCod),0) " & _
            " bAgenteRetencion,  Case When (Select Max(PBC.cPersBCRuc) " & _
            " From PersID PID Inner Join PersSunatBC PBC On PID.cPersIDnro = PBC.cPersBCRuc " & _
            " And PID.cPersIDTpo = 2 " & _
            " Where PID.cPersCod = pr.cPersCod) Is Null Then " & _
            " 'No es Agente_Ret/BC' Else 'Agente_Ret/BC' End  [A.Ret/B.Cont], " & _
            " case when PR.bhabilSUNAT =1 then 'Habilitado' " & _
            " when PR.bhabilSUNAT=0 then 'Inhabilitado' " & _
            " End as 'Estado Sunat', " & _
            " case when PR.bHabilCONSUCODE =1 then 'Habilitado' " & _
            " when PR.bHabilCONSUCODE=0 then 'Inhabilitado' " & _
            " End as 'Estado Consucode', " & _
            " PR.cComentario " & _
            " FROM Proveedor PR JOIN " & vsCentralPer & "Persona P ON pr.cPersCod = p.cPersCod " & _
            " LEFT JOIN " & vsCentralPer & "PersID PI1 ON pr.cPersCod = pi1.cPersCod AND pi1.cPersIdTpo = '" & gPersIdRUC & "'" & _
            " LEFT JOIN " & vsCentralPer & "PersID PI2 ON pr.cPersCod = pi2.cPersCod AND pi2.cPersIdTpo = '" & gPersIdRUS & "'" & _
            " LEFT JOIN " & vsCentralCom & "Constante C ON pr.nProvEstado = c.nConsValor AND c.nConsCod = " & gLogProvEstado & " " & _
            " WHERE SUBSTRING(PR.cUltimaActualizacion,18,2)='" & gsCodAgen & "' " & _
            " AND P.cPersNombre LIKE  '" & sValorBusqueda & "%' And P.cPersEstado NOT IN ('X') " & _
            " " ' ORDER BY p.cPersNombre "
        
       sSQL = " SELECT pr.cPersCod Código, p.cPersNombre Nombre, p.cPersDireccDomicilio Dirección, " & _
            " pi1.cPersIdNro RUC, P.cPersTelefono Telefono, pr.nProvEstado , c.cConsDescripcion Estado, " & _
            " PR.cCtaCodMN Cta_MN, PR.cCtaCodME Cta_ME, Case When AGRET.cPersBCRuc Is Null Then 0 Else 1 End bAgenteRetencion,  " & _
            " Case When AGRET.cPersBCRuc Is Null Then 'No es Agente_Ret/BC' Else 'Agente_Ret/BC' End [A.Ret/B.Cont], " & _
            " case when PR.bhabilSUNAT =1 then 'Habilitado' " & _
            " when PR.bhabilSUNAT=0 then 'Inhabilitado' " & _
            " End as 'Estado Sunat', " & _
            " case when PR.bHabilCONSUCODE =1 then 'Habilitado' " & _
            " when PR.bHabilCONSUCODE=0 then 'Inhabilitado' " & _
            " End as 'Estado Consucode', " & _
            " PR.cComentario " & _
            " FROM Proveedor PR JOIN " & vsCentralPer & "Persona P ON pr.cPersCod = p.cPersCod " & _
            " LEFT JOIN (Select PBC.cPersBCRuc From PersID PID " & _
            "            Inner Join PersSunatBC PBC On PID.cPersIDnro = PBC.cPersBCRuc  And PID.cPersIDTpo = 2) AGRET On pr.cPersCod = AGRET.cPersBCRuc " & _
            " LEFT JOIN " & vsCentralPer & "PersID PI1 ON pr.cPersCod = pi1.cPersCod AND pi1.cPersIdTpo = '" & gPersIdRUC & "'" & _
            " LEFT JOIN " & vsCentralPer & "PersID PI2 ON pr.cPersCod = pi2.cPersCod AND pi2.cPersIdTpo = '" & gPersIdRUS & "'" & _
            " LEFT JOIN " & vsCentralCom & "Constante C ON pr.nProvEstado = c.nConsValor AND c.nConsCod = " & gLogProvEstado & " Inner Join ProveedorAG PAG On PAG.cPersCod = PR.cPersCod  " & _
            " WHERE PAG.cAgeCod = '" & gsCodAgeN & "' " & _
            " AND P.cPersNombre LIKE  '" & sValorBusqueda & "%' And P.cPersEstado <> 'X' " & _
            " ORDER BY p.cPersNombre "
        
        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
    End If
    Set CargaProveedores = rs
    
    Case Is = 1
    
    If oConec.AbreConexion() Then
        'sSQL = " SELECT pr.cPersCod Código, p.cPersNombre Nombre, p.cPersDireccDomicilio Dirección, " & _
            " pi1.cPersIdNro RUC, P.cPersTelefono Telefono, pr.nProvEstado , c.cConsDescripcion Estado, " & _
            " PR.cCtaCodMN Cta_MN, PR.cCtaCodME Cta_ME, ISNull((Select Max(PBC.cPersBCRuc) " & _
            " From PersID PID Inner Join PersSunatBC PBC On PID.cPersIDnro = PBC.cPersBCRuc " & _
            " And PID.cPersIDTpo = 2 Where PID.cPersCod = pr.cPersCod),0) " & _
            " bAgenteRetencion,  Case When (Select Max(PBC.cPersBCRuc) " & _
            " From PersID PID Inner Join PersSunatBC PBC On PID.cPersIDnro = PBC.cPersBCRuc " & _
            " And PID.cPersIDTpo = 2 " & _
            " Where PID.cPersCod = pr.cPersCod) Is Null Then " & _
            " 'No es Agente_Ret/BC' Else 'Agente_Ret/BC' End  [A.Ret/B.Cont],  " & _
            " case when PR.bhabilSUNAT =1 then 'Habilitado' " & _
            " when PR.bhabilSUNAT=0 then 'Inhabilitado' " & _
            " End as 'Estado Sunat', " & _
            " case when PR.bHabilCONSUCODE =1 then 'Habilitado' " & _
            " when PR.bHabilCONSUCODE=0 then 'Inhabilitado' " & _
            " End as 'Estado Consucode', " & _
            " PR.cComentario " & _
            " FROM Proveedor PR JOIN " & vsCentralPer & "Persona P ON pr.cPersCod = p.cPersCod " & _
            " LEFT JOIN " & vsCentralPer & "PersID PI1 ON pr.cPersCod = pi1.cPersCod AND pi1.cPersIdTpo = '" & gPersIdRUC & "'" & _
            " LEFT JOIN " & vsCentralPer & "PersID PI2 ON pr.cPersCod = pi2.cPersCod AND pi2.cPersIdTpo = '" & gPersIdRUS & "'" & _
            " LEFT JOIN " & vsCentralCom & "Constante C ON pr.nProvEstado = c.nConsValor AND c.nConsCod = " & gLogProvEstado & " " & _
            " WHERE P.cPersNombre LIKE  '" & sValorBusqueda & "%' And P.cPersEstado <> 'X' " & _
            "" ' ORDER BY p.cPersNombre "
               
    sSQL = " SELECT pr.cPersCod Código, p.cPersNombre Nombre, p.cPersDireccDomicilio Dirección, " & _
            " pi1.cPersIdNro RUC, P.cPersTelefono Telefono, pr.nProvEstado , c.cConsDescripcion Estado, " & _
            " PR.cCtaCodMN Cta_MN, PR.cCtaCodME Cta_ME, Case When AGRET.cPersBCRuc Is Null Then 0 Else 1 End bAgenteRetencion,  " & _
            " Case When AGRET.cPersBCRuc Is Null Then 'No es Agente_Ret/BC' Else 'Agente_Ret/BC' End [A.Ret/B.Cont], " & _
            " case when PR.bhabilSUNAT =1 then 'Habilitado' " & _
            " when PR.bhabilSUNAT=0 then 'Inhabilitado' " & _
            " End as 'Estado Sunat', " & _
            " case when PR.bHabilCONSUCODE =1 then 'Habilitado' " & _
            " when PR.bHabilCONSUCODE=0 then 'Inhabilitado' " & _
            " End as 'Estado Consucode', " & _
            " PR.cComentario " & _
            " FROM Proveedor PR JOIN " & vsCentralPer & "Persona P ON pr.cPersCod = p.cPersCod " & _
            " LEFT JOIN (Select PBC.cPersBCRuc From PersID PID " & _
            "            Inner Join PersSunatBC PBC On PID.cPersIDnro = PBC.cPersBCRuc  And PID.cPersIDTpo = 2) AGRET On pr.cPersCod = AGRET.cPersBCRuc " & _
            " LEFT JOIN " & vsCentralPer & "PersID PI1 ON pr.cPersCod = pi1.cPersCod AND pi1.cPersIdTpo = '" & gPersIdRUC & "'" & _
            " LEFT JOIN " & vsCentralPer & "PersID PI2 ON pr.cPersCod = pi2.cPersCod AND pi2.cPersIdTpo = '" & gPersIdRUS & "'" & _
            " LEFT JOIN " & vsCentralCom & "Constante C ON pr.nProvEstado = c.nConsValor AND c.nConsCod = " & gLogProvEstado & " " & _
            " WHERE P.cPersNombre LIKE  '" & sValorBusqueda & "%' And P.cPersEstado <> 'X' " & _
            " ORDER BY p.cPersNombre "
            
        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
    End If
    
    Set CargaProveedores = rs
    End Select
End Function





Public Function HabDesSunat(ByVal psSql As String) As Integer
    Dim oConec As DConecta
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
        oConec.Ejecutar psSql
        oConec.CierraConexion
        HabDesSunat = 0
    Else
        HabDesSunat = 1
    End If
End Function
Public Function IngresaQuitaComent(ByVal psSql As String) As Integer
    Dim oConec As DConecta
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
        oConec.Ejecutar psSql
        oConec.CierraConexion
        
        IngresaQuitaComent = 0
    Else
        IngresaQuitaComent = 1
    End If
End Function
Public Function HabDesConsucode(ByVal psSql As String) As Integer
    Dim oConec As DConecta
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
        oConec.Ejecutar psSql
        oConec.CierraConexion
        
        HabDesConsucode = 0
    Else
        HabDesConsucode = 1
    End If
End Function
'Agrega un bien/servicio de proveedor
Public Function GrabaProveedorBS(ByVal psPersCod As String, ByVal psBSCod As String, ByVal psActualizacion As String) As Integer
    Dim sActualiza As String
    Dim oConec As DConecta
    Dim rs As Recordset

    Set oConec = New DConecta

    If oConec.AbreConexion() Then

        sSQL = "SELECT cBSCod FROM ProveedorBS " & _
            " WHERE cPersCod = '" & psPersCod & "' AND cBSCod = '" & psBSCod & "'"
        Set rs = oConec.CargaRecordSet(sSQL)

        If rs.RecordCount = 0 Then
            sActualiza = Left(psActualizacion, 8) & Format(oConec.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)

            sSQL = "INSERT INTO ProveedorBS(cPersCod, cBSCod, cUltimaActualizacion) " & _
                " VALUES ('" & psPersCod & "','" & psBSCod & "','" & sActualiza & "')"
            oConec.Ejecutar sSQL
            GrabaProveedorBS = 0
        Else
            GrabaProveedorBS = 2
        End If
        oConec.CierraConexion
    Else
        GrabaProveedorBS = 1
    End If
End Function

Public Function GetProvCtas(psPersCod As String, psAgeCod As String, pMoneda As Moneda, pgBitCentral As Boolean) As Recordset
    Dim Sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If Not pgBitCentral Then
        Sql = " Select AHO.cCodCta, '' from PersCuenta PC " _
            & " Inner Join AhorroC AHO On PC.cCodCta = AHO.cCodCta" _
            & " Where PC.cRelaCta = 'TI' And PC.cCodPers = '" & Mid(psPersCod, 4) & "' And Substring(AHO.cCodCta,6,1) = '" & pMoneda & "'" _
            & " And cEstCtaAC Not In ('C','U')"
        
        oCon.AbreConexion 'Remota psAgeCod
    Else
        Sql = " Select AHO.cCtaCod, '' from ProductoPersona PC " _
            & " Inner Join Producto AHO On PC.cCtaCod = AHO.cCtaCod And SubString(AHO.cCtaCod,6,3) = '" & Producto.gCapAhorros & "'" _
            & " Where PC.nPrdPersRelac = " & CaptacRelacPersona.gCapRelPersTitular & " And PC.cPersCod = '" & psPersCod & "' And Substring(AHO.cCtaCod,9,1) = '" & pMoneda & "'" _
            & " And AHO.nPrdEstado Not In ('" & CaptacEstado.gCapEstAnulada & "','" & CaptacEstado.gCapEstCancelada & "')"
        
        oCon.AbreConexion
    End If
    
    Set GetProvCtas = oCon.CargaRecordSet(Sql)
    
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function SetProvCtas(psPersCod As String, psCtaMN As String, psCtaME As String, psUltAct As String) As Boolean
    Dim Sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    Sql = " Update Proveedor " _
        & " Set cCtaCodMN = '" & psCtaMN & "' , cCtaCodME = '" & psCtaME & "', cUltimaActualizacion = '" & psUltAct & "'" _
        & " Where cPersCod = '" & psPersCod & "'"
    
    oCon.AbreConexion
    
    oCon.Ejecutar Sql
    
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function SetProvCtasBancos(psPersCod As String, psBancoCodMN As String, psCtaMN As String, psBancoCodME As String, psCtaME As String, psUltAct As String, _
                                                            psCtaCCIMN As String, psCtaCCIME As String, psBancoCodDetracMN As String, psCtaDetracMN As String) As Boolean
    Dim Sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    'Modificado PASIERS0472015********************************
'    sql = " Update Proveedor " _
'        & " Set cCtaCodMN = '" & psCtaMN & "' , cPersCodBancoMN ='" & psBancoCodMN & "' ,  cCtaCodME = '" & psCtaME & "', cPersCodBancoME ='" & psBancoCodME & "' ,cUltimaActualizacion = '" & psUltAct & "'" _
'        & " Where cPersCod = '" & psPersCod & "'"
    
    Sql = "stp_sel_ERS0472015_ActualizaProveedor '" & psPersCod & "','" & psCtaMN & "','" & psCtaME & "','" & psBancoCodMN & "','" & psBancoCodME & "','" & psCtaCCIMN & "','" & psCtaCCIME & "','" & psBancoCodDetracMN & "','" & psCtaDetracMN & "','" & psUltAct & "'"
    
    'end PASI*********************************************************
    oCon.AbreConexion
    
    oCon.Ejecutar Sql
    
    oCon.CierraConexion
    Set oCon = Nothing
End Function

'Public Function GetProveedorAgeRetBuenCont(psPersCod As String) As ADODB.Recordset
'    Dim sql As String
'    Dim oCon As DConecta
'    Set oCon = New DConecta
'
'    sql = " Select bAgenteRetencion, bBuenContribuyente, bHabilSUNAT, bHabilCONSUCODE, cComentario + ' (' +  dbo.FechaHoraNombreCortoMov(cUltimaActualizacion)  + ')' cComentario  From Proveedor " _
'        & " Where cPersCod = '" & psPersCod & "' And nProvEstado = " & gLogProvEstadoActivado
'
'    oCon.AbreConexion
'
'    Set GetProveedorAgeRetBuenCont = oCon.CargaRecordSet(sql)
'End Function

Public Function GetProveedorAgeRetBuenCont(psPersCod As String) As ADODB.Recordset
    Dim Sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    Sql = " Select bAgenteRetencion, bBuenContribuyente From Proveedor " _
        & " Where cPersCod = '" & psPersCod & "' And nProvEstado = " & gLogProvEstadoActivado
    
    oCon.AbreConexion
    
    Set GetProveedorAgeRetBuenCont = oCon.CargaRecordSet(Sql)
End Function


Public Sub ActualizaBSProveedores(psMovNro As String)
    Dim Sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
        
    Sql = " Insert Proveedor(cPersCod, nProvEstado, cCtaCodMN, cCtaCodMe, bAgenteRetencion, bBuenContribuyente, cUltimaActualizacion)" _
        & " Select distinct cPerscod, 1, Null, Null, 0, 0, '" & psMovNro & "' from mov a" _
        & " inner join movbs b on a.nmovnro = b.nmovnro" _
        & " inner join movgasto c on a.nmovnro = c.nmovnro" _
        & " where copecod like '591101' and nMovFlag <> 1 and cperscod not in (Select cPersCod from proveedor)"
    oCon.Ejecutar Sql
        
    Sql = " Insert ProveedorBS(cPersCod, cbscod, cUltimaActualizacion)" _
        & " Select distinct cperscod, cbscod, '" & psMovNro & "'  from mov a" _
        & " inner join movbs b on a.nmovnro = b.nmovnro" _
        & " inner join movgasto c on a.nmovnro = c.nmovnro" _
        & " where copecod like '591101' and nMovFlag <> 1 and cperscod + cbscod not in  (Select cperscod + cbscod from proveedorbs)" _
        & " and cbscod in (Select cbscod from bienesservicios)"
    oCon.Ejecutar Sql
    

        
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Function GetOCOSVencidas(pdFecha As Date) As ADODB.Recordset
    Dim Sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    oCon.AbreConexion
    
    Sql = " Select Distinct  Case MD.nDocTpo When 33 Then 'OC' Else 'OS' End + ' ' + MD.cDocNro cDocNro, dbo.PstaNombre(PE.cPersNombre,1) Nombre, Convert(Varchar(10),MC.dMovPlazo,103) Fecha, Case Substring(M.cOpeCod,3,1) When '2' Then 'DOLAR' Else 'SOLES' End Moneda, IsNull(PE.cPersTelefono,'') Fono  from Mov M" _
        & " Inner Join MovCotizac MC On M.nMovNro = MC.nMovNro" _
        & " Inner Join MovGasto MG On M.nMovNro = MG.nMovNro" _
        & " Inner Join MovDoc MD On M.nMovNro = MD.nMovNro" _
        & " Inner Join Persona PE On MG.cPersCod = PE.cPersCod" _
        & " Where MC.dMovPlazo Between DateAdd(Day,-3,'" & Format(pdFecha, gsFormatoFecha) & "') And DateAdd(Day,3,'" & Format(pdFecha, gsFormatoFecha) & "') And nMovEstado = 16  Order By Fecha"
    Set rs = oCon.CargaRecordSet(Sql)
    
    Set GetOCOSVencidas = rs
    
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function GetProveedorNoHabido(psPersRuc As String, ByRef lsMotivoNoHabido As String) As Boolean
    Dim Sql As String
    Dim oCon As DConecta
    Dim prs  As ADODB.Recordset
    Set oCon = New DConecta
    oCon.AbreConexion
    
    Sql = " Select cMotivo From PersContribuyeNoHabido " _
        & " Where cRuc = '" & psPersRuc & "' "
    Set prs = oCon.CargaRecordSet(Sql)
    If prs.EOF Then
      lsMotivoNoHabido = ""
      GetProveedorNoHabido = False
    Else
      lsMotivoNoHabido = prs!cMotivo
      GetProveedorNoHabido = True
    End If
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function GetProveedorPerceptores(psPersRuc As String, ByRef lsMotivoNoHabido As String) As Boolean
    Dim Sql As String
    Dim oCon As DConecta
    Dim prs  As ADODB.Recordset
    Set oCon = New DConecta
    oCon.AbreConexion
    
    Sql = " Select cRuc From PersContribuyePerceptores " _
        & " Where cRuc = '" & psPersRuc & "' "
    Set prs = oCon.CargaRecordSet(Sql)
    If prs.EOF Then
      lsMotivoNoHabido = ""
      GetProveedorPerceptores = False
    Else
      lsMotivoNoHabido = ""
      GetProveedorPerceptores = True
    End If
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function GetProveedorRUC(psPersCod As String) As ADODB.Recordset
    Dim Sql As String
    Dim oCon As DConecta
    Dim prs  As ADODB.Recordset
    Set oCon = New DConecta
    oCon.AbreConexion
    
    Sql = " Select isnull(cPersIDnro,'')cPersIDnro From PersID" _
        & " Where cPersCod = '" & psPersCod & "' And  cPersIDTpo=2 "
    
    Set GetProveedorRUC = oCon.CargaRecordSet(Sql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function



