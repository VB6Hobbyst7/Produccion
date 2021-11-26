VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCredExtorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function ObtenerDatosExtorno(ByVal psctacod As String, ByVal pnTipoAprobRech As Integer) _
As Recordset
    Dim oConec As DConecta
    Dim sSql As String
    Dim lnPrdEstado As Integer
    On Error GoTo ErrHandler
        If pnTipoAprobRech = 1 Then 'Aprobado
            lnPrdEstado = 2002
        Else
            'rechazado
            lnPrdEstado = 2003
        End If
        sSql = "SELECT PE.CPERSNOMBRE AS CTITULAR,CN.CCONSDESCRIPCION AS ESTADO,"
        sSql = sSql & " SUBSTRING(P.CCTACOD,9,1) AS NMONEDA,' 'AS USUARIO,"
        sSql = sSql & " CAST(DATEPART(HOUR,CE.DPRDESTADO) AS CHAR(2))+':'+ CAST(DATEPART(MINUTE,CE.DPRDESTADO) AS CHAR(2))+':'+CAST(DATEPART(SECOND,CE.DPRDESTADO) AS CHAR(2)) AS CHORA,"
        sSql = sSql & " CE.DPRDESTADO AS DFECHA,CE.NMONTO AS NMONTO"
        sSql = sSql & " FROM PRODUCTO P"
        sSql = sSql & " INNER JOIN PRODUCTOPERSONA PP ON P.CCTACOD=PP.CCTACOD AND PP.nPrdPersRelac=20"
        sSql = sSql & " INNER JOIN PERSONA PE ON PE.CPERSCOD=PP.CPERSCOD"
        sSql = sSql & " INNER JOIN CONSTANTE CN ON P.NPRDESTADO=CN.NCONSVALOR AND CN.NCONSCOD=3001"
        sSql = sSql & " INNER JOIN COLOCACESTADO CE ON CE.CCTACOD=P.CCTACOD AND CE.NPRDESTADO=" & lnPrdEstado
        sSql = sSql & " WHERE P.CCTACOD='" & psctacod & "'"
         
        Set oConec = New DConecta
        oConec.AbreConexion
        Set ObtenerDatosExtorno = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
        Set oConec = Nothing
    Exit Function
ErrHandler:
    If oConec Is Nothing Then Set oConec = Nothing
    Set ObtenerDatosExtorno = Null
End Function
'1 verdad
'0 es falso
'-1 fallo
Public Function ValidaExtorno(ByVal psctacod As String, ByVal pnTipoAprobRech As Integer) As Integer
    Dim oConec As DConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim nCantidad As Integer
    Dim lnPrdEstado As Integer
    On Error GoTo ErrHandler
        If pnTipoAprobRech = 1 Then 'Aprobado
            lnPrdEstado = 2002
        Else
            'rechazado
            lnPrdEstado = 2003
        End If
    
        sSql = sSql & " SELECT COUNT(*) AS CANTIDAD"
        sSql = sSql & " From COLOCACESTADO"
        sSql = sSql & " WHERE CCTACOD='" & psctacod & "' AND NPRDESTADO=" & lnPrdEstado
        
        Set oConec = New DConecta
        oConec.AbreConexion
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
        Set oConec = Nothing
        If Not rs.EOF And Not rs.BOF Then
            nCantidad = rs!cantidad
            If nCantidad > 0 Then
                ValidaExtorno = 1
            Else
                ValidaExtorno = 0
            End If
        Else
            ValidaExtorno = -1
        End If
        Set rs = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    If Not rs Is Nothing Then Set rs = Nothing
    ValidaExtorno = -1
End Function
Public Function Extorno(ByVal psctacod As String) As Boolean
    Dim oConec As DConecta
    
    Dim sSql As String
    On Error GoTo ErrHandler
    Set oConec = New DConecta
        oConec.AbreConexion
        oConec.ConexionActiva.BeginTrans
        
        'Volviendo el credito a estado de sugerido
        sSql = "Update Producto set nPrdEstado=2001 Where cCtaCod='" & psctacod & "'"
        oConec.ConexionActiva.Execute sSql
        
        'Eliminado el estado de aprobado
        sSql = "Delete From ColocacEstado Where cCtaCod='" & psctacod & "' and nPrdEstado=2002"
        oConec.ConexionActiva.Execute sSql
        
        'Eliminado el ProductoPersona
        sSql = "Delete from ProductoPersona Where nPrdPersRelac=29 and cctacod='" & psctacod & "'"
        oConec.ConexionActiva.Execute sSql
        
        'Verificando si el tiene un credito mi vivienda
           If ObteMiVivienda(psctacod) = True Then
                sSql = "Delete From ColocCalifMiViv Where cCtaCod='" & psctacod & "'"
                oConec.ConexionActiva.Execute sSql
           End If
         oConec.ConexionActiva.CommitTrans
         oConec.CierraConexion
         Set oConec = Nothing
         Extorno = True
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then
        oConec.ConexionActiva.RollbackTrans
    End If
    If Not oConec Is Nothing Then Set oConec = Nothing
    Extorno = False
End Function

Function ObteMiVivienda(ByVal psctacod As String) As Boolean
    Dim oConec As DConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim nCantidad As Integer
    On Error GoTo ErrHandler
        sSql = "Select count(*) as Cantidad From ColocacCred Where bMiVivienda=1 and cCtaCod='" & psctacod & "'"
        Set oConec = New DConecta
        oConec.AbreConexion
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
        Set oConec = Nothing
        If Not rs.BOF And Not rs.EOF Then
            nCantidad = rs!cantidad
        Else
            nCantidad = 0
        End If
        Set rs = Nothing
        
        If nCantidad > 0 Then
            ObteMiVivienda = True
        Else
            ObteMiVivienda = False
        End If
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    ObteMiVivienda = False
End Function
