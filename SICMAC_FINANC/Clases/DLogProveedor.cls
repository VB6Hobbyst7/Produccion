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
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim vsCentralPer As String
Dim vsCentralCom As String
Dim vsCentralImg As String
Dim sSql As String

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
    Dim oIni As ClasIni
    
    Set oIni = New ClasIni
        vsCentralPer = oIni.BasePersonas
        vsCentralCom = oIni.BaseComunes
        vsCentralImg = oIni.BaseImagenes
    Set oIni = Nothing
End Sub

'Devuelve un valor si existe proveedor
Public Function IsExisProveedor(ByVal psPerscod As String, Optional pbConEstado As Boolean = False) As Boolean
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        If pbConEstado = False Then
            sSql = "SELECT cPersCod FROM Proveedor " & _
                " WHERE  cPersCod = '" & psPerscod & "' "
        Else
            sSql = "SELECT cPersCod FROM Proveedor " & _
                " WHERE  cPersCod = '" & psPerscod & "' AND cProvEstado = '" & gLogProvEstadoActivado & "' "
        End If
        Set rs = oConec.CargaRecordSet(sSql)
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
        sSql = "SELECT pr.cPersCod Código, p.cPersNombre Nombre, p.cPersDireccDomicilio Dirección, " & _
            " pi1.cPersIdNro RUC, pi2.cPersIdNro RUS, pr.cProvEstado , c.cConsDescripcion Estado" & _
            " FROM Proveedor PR JOIN " & vsCentralPer & "Persona P ON pr.cPersCod = p.cPersCod " & _
            " LEFT JOIN " & vsCentralPer & "PersID PI1 ON pr.cPersCod = pi1.cPersCod AND pi1.cPersIdTpo = '" & gPersIdRUC & "'" & _
            " LEFT JOIN " & vsCentralPer & "PersID PI2 ON pr.cPersCod = pi2.cPersCod AND pi2.cPersIdTpo = '" & gPersIdRUS & "'" & _
            " LEFT JOIN " & vsCentralCom & "Constante C ON pr.cProvEstado = c.nConsValor AND c.nCONSCOD = '" & gLogProvEstado & "'" & _
            " ORDER BY p.cPersNombre "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set CargaProveedor = rs
End Function

'Devuelve un RecordSet con los Bienes/Servicios que brinda un proveedor
Public Function CargaProveedorBS(ByVal psPerscod As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sSql = "SELECT pbs.cBSCod Código, pbs.cBSCod Código, bs.cBSDescripcion Descripción " & _
            " FROM ProveedorBS PBS JOIN BienesServicios BS ON pbs.cBSCod = bs.cBSCod AND pbs.cPersCod = '" & psPerscod & "'" & _
            " ORDER BY pbs.cBSCod "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set CargaProveedorBS = rs
End Function

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
Public Function EliminaProveedor(ByVal psPerscod As String) As Integer
    Dim oConec As DConecta
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
        sSql = "DELETE ProveedorBS WHERE cPersCod = '" & psPerscod & "'"
        oConec.Ejecutar sSql
        
        sSql = "DELETE Proveedor WHERE cPersCod = '" & psPerscod & "'"
        oConec.Ejecutar sSql
        
        oConec.CierraConexion
        EliminaProveedor = 0
    Else
        EliminaProveedor = 1
    End If
End Function

'Elimina el bien/servicio que presta un proveedor
Public Function EliminaProveedorBS(ByVal psPerscod As String, ByVal psBSCod As String) As Integer
    Dim oConec As DConecta
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
        sSql = "DELETE ProveedorBS WHERE cPersCod = '" & psPerscod & "'" & _
            " AND cBSCod = '" & psBSCod & "'"
        oConec.Ejecutar sSql
        
        oConec.CierraConexion
        EliminaProveedorBS = 0
    Else
        EliminaProveedorBS = 1
    End If
End Function

'Agrega un proveedor
Public Function GrabaProveedor(ByVal psPerscod As String, ByVal psActualizacion As String) As Integer
    Dim sActualiza As String
    Dim oConec As DConecta
    Dim rs As Recordset

    Set oConec = New DConecta

    If oConec.AbreConexion() Then
        sSql = "SELECT cPersCod FROM Proveedor " & _
            " WHERE cPersCod = '" & psPerscod & "' "
        Set rs = oConec.CargaRecordSet(sSql)

        If rs.RecordCount = 0 Then
            sActualiza = Left(psActualizacion, 8) & Format(oConec.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)

            sSql = "INSERT INTO Proveedor(cPersCod, cProvEstado, cUltimaActualizacion) " & _
                " VALUES ('" & psPerscod & "','" & gLogProvEstadoActivado & "','" & sActualiza & "')"
            oConec.Ejecutar sSql
            GrabaProveedor = 0
        Else
            GrabaProveedor = 2
        End If
        oConec.CierraConexion
    Else
        GrabaProveedor = 1
    End If
End Function

'Agrega un bien/servicio de proveedor
Public Function GrabaProveedorBS(ByVal psPerscod As String, ByVal psBSCod As String, ByVal psActualizacion As String) As Integer
    Dim sActualiza As String
    Dim oConec As DConecta
    Dim rs As Recordset

    Set oConec = New DConecta

    If oConec.AbreConexion() Then

        sSql = "SELECT cBSCod FROM ProveedorBS " & _
            " WHERE cPersCod = '" & psPerscod & "' AND cBSCod = '" & psBSCod & "'"
        Set rs = oConec.CargaRecordSet(sSql)

        If rs.RecordCount = 0 Then
            sActualiza = Left(psActualizacion, 8) & Format(oConec.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)

            sSql = "INSERT INTO ProveedorBS(cPersCod, cBSCod, cUltimaActualizacion) " & _
                " VALUES ('" & psPerscod & "','" & psBSCod & "','" & sActualiza & "')"
            oConec.Ejecutar sSql
            GrabaProveedorBS = 0
        Else
            GrabaProveedorBS = 2
        End If
        oConec.CierraConexion
    Else
        GrabaProveedorBS = 1
    End If
End Function


Public Function GetProvCtas(psPerscod As String, psAgeCod As String, pMoneda As Moneda, pgBitCentral As Boolean) As Recordset
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If Not pgBitCentral Then
        sql = " Select AHO.cCodCta, '' from PersCuenta PC " _
            & " Inner Join AhorroC AHO On PC.cCodCta = AHO.cCodCta" _
            & " Where PC.cRelaCta = 'TI' And PC.cCodPers = '" & Mid(psPerscod, 4) & "' And Substring(AHO.cCodCta,6,1) = '" & pMoneda & "'" _
            & " And cEstCtaAC Not In ('C','U')"
        
        oCon.AbreConexion 'Remota psAgeCod
    Else
        sql = " Select AHO.cCtaCod, '' from ProductoPersona PC " _
            & " Inner Join Producto AHO On PC.cCtaCod = AHO.cCtaCod And SubString(AHO.cCtaCod,6,3) = '" & Producto.gCapAhorros & "'" _
            & " Where PC.nPrdPersRelac = " & CaptacRelacPersona.gCapRelPersTitular & " And PC.cPersCod = '" & psPerscod & "' And Substring(AHO.cCodCta,9,1) = '" & pMoneda & "'" _
            & " And AHO.nPrdEstado Not In ('" & CaptacEstado.gCapEstAnulada & "','" & CaptacEstado.gCapEstCancelada & "')"
        
        oCon.AbreConexion psAgeCod
    End If
    
    Set GetProvCtas = oCon.CargaRecordSet(sql)
    
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function SetProvCtas(psPerscod As String, psCtaMN As String, psCtaME As String, psUltAct As String) As Boolean
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sql = " Update Proveedor " _
        & " Set cCtaCodMN = '" & psCtaMN & "' , cCtaCodME = '" & psCtaME & "', cUltimaActualizacion = '" & psUltAct & "'" _
        & " Where cPersCod = '" & psPerscod & "'"
    
    oCon.AbreConexion
    
    oCon.Ejecutar sql
    
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function GetProveedorAgeRetBuenCont(psPerscod As String) As ADODB.Recordset
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sql = " Select bAgenteRetencion, bBuenContribuyente From Proveedor " _
        & " Where cPersCod = '" & psPerscod & "' And nProvEstado = " & gLogProvEstadoActivado
    
    oCon.AbreConexion
    
    Set GetProveedorAgeRetBuenCont = oCon.CargaRecordSet(sql)
End Function

Public Function GetProveedorNoHabido(psPersRuc As String, ByRef lsMotivoNoHabido As String) As Boolean
    Dim sql As String
    Dim oCon As DConecta
    Dim prs  As ADODB.Recordset
    Set oCon = New DConecta
    oCon.AbreConexion
    
    sql = " Select cMotivo From PersContribuyeNoHabido " _
        & " Where cRuc = '" & psPersRuc & "' "
    Set prs = oCon.CargaRecordSet(sql)
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

Public Function GetProveedorRUC(psPerscod As String) As ADODB.Recordset
    Dim sql As String
    Dim oCon As DConecta
    Dim prs  As ADODB.Recordset
    Set oCon = New DConecta
    oCon.AbreConexion
    
    sql = " Select isnull(cPersIDnro,'')cPersIDnro From PersID" _
        & " Where cPersCod = '" & psPerscod & "' And  cPersIDTpo=2 "
    
    Set GetProveedorRUC = oCon.CargaRecordSet(sql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'PASI20140612 TI-ERS061-2014
Public Function GetProveedorRUCoDNI(psPerscod As String) As ADODB.Recordset
    Dim sql As String
    Dim oCon As DConecta
    Dim prs  As ADODB.Recordset
    Set oCon = New DConecta
    oCon.AbreConexion
    
    sql = "exec stp_sel_ObtenerProveedorDNIoRUC '" & psPerscod & "'"
    
    Set GetProveedorRUCoDNI = oCon.CargaRecordSet(sql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'end PASI

Public Function GetProveedorDNI(psPerscod As String) As ADODB.Recordset
    Dim sql As String
    Dim oCon As DConecta
    Dim prs  As ADODB.Recordset
    Set oCon = New DConecta
    oCon.AbreConexion
    
    sql = " Select isnull(cPersIDnro,'')cPersIDnro From PersID" _
        & " Where cPersCod = '" & psPerscod & "' And  cPersIDTpo=1 "
    
    Set GetProveedorDNI = oCon.CargaRecordSet(sql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function GetProveedorCarnetExt(psPerscod As String) As ADODB.Recordset
    Dim sql As String
    Dim oCon As DConecta
    Dim prs  As ADODB.Recordset
    Set oCon = New DConecta
    oCon.AbreConexion
    
    sql = " Select isnull(cPersIDnro,'')cPersIDnro From PersID" _
        & " Where cPersCod = '" & psPerscod & "' And  cPersIDTpo=4 "
    
    Set GetProveedorCarnetExt = oCon.CargaRecordSet(sql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function GetProvDirector(pdFecha As Date) As ADODB.Recordset
    Dim sql As String
    Dim oCon As New DConecta
    oCon.AbreConexion
    sql = "Exec CargaDirectoresDietasxPagar '" & Format(pdFecha, "YYYYMMDD") & "'"
    Set GetProvDirector = oCon.CargaRecordSet(sql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function '****NAGL INC1712260008


