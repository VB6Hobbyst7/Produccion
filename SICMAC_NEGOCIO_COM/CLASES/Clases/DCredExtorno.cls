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
Public Function ObtenerDatosExtorno(ByVal psCtaCod As String, ByVal pnTipoAprobRech As Integer) _
As Recordset
    Dim oConec As DConecta
    Dim sSQL As String
    Dim lnPrdEstado As Integer
    On Error GoTo ErrHandler
        If pnTipoAprobRech = 1 Then 'Aprobado
            lnPrdEstado = 2002
        Else
            'rechazado
            lnPrdEstado = 2003
        End If
        sSQL = "SELECT PE.CPERSNOMBRE AS CTITULAR,CN.CCONSDESCRIPCION AS ESTADO,"
        sSQL = sSQL & " SUBSTRING(P.CCTACOD,9,1) AS NMONEDA,' 'AS USUARIO,"
        sSQL = sSQL & " CAST(DATEPART(HOUR,CE.DPRDESTADO) AS CHAR(2))+':'+ CAST(DATEPART(MINUTE,CE.DPRDESTADO) AS CHAR(2))+':'+CAST(DATEPART(SECOND,CE.DPRDESTADO) AS CHAR(2)) AS CHORA,"
        sSQL = sSQL & " CE.DPRDESTADO AS DFECHA,CE.NMONTO AS NMONTO"
        sSQL = sSQL & " FROM PRODUCTO P"
        sSQL = sSQL & " INNER JOIN PRODUCTOPERSONA PP ON P.CCTACOD=PP.CCTACOD AND PP.nPrdPersRelac=20"
        sSQL = sSQL & " INNER JOIN PERSONA PE ON PE.CPERSCOD=PP.CPERSCOD"
        sSQL = sSQL & " INNER JOIN CONSTANTE CN ON P.NPRDESTADO=CN.NCONSVALOR AND CN.NCONSCOD=3001"
        sSQL = sSQL & " INNER JOIN COLOCACESTADO CE ON CE.CCTACOD=P.CCTACOD AND CE.NPRDESTADO=" & lnPrdEstado
        sSQL = sSQL & " WHERE P.CCTACOD='" & psCtaCod & "'"
         
        Set oConec = New DConecta
        oConec.AbreConexion
        Set ObtenerDatosExtorno = oConec.CargaRecordSet(sSQL)
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
Public Function ValidaExtorno(ByVal psCtaCod As String, ByVal pnTipoAprobRech As Integer) As Integer
    Dim oConec As DConecta
    Dim sSQL As String
    Dim Rs As ADODB.Recordset
    Dim nCantidad As Integer
    Dim lnPrdEstado As Integer
    On Error GoTo ErrHandler
        If pnTipoAprobRech = 1 Then 'Aprobado
            lnPrdEstado = 2002
        Else
            'rechazado
            lnPrdEstado = 2003
        End If
    
        sSQL = sSQL & " SELECT COUNT(*) AS CANTIDAD"
        sSQL = sSQL & " From COLOCACESTADO"
        sSQL = sSQL & " WHERE CCTACOD='" & psCtaCod & "' AND NPRDESTADO=" & lnPrdEstado
        
        Set oConec = New DConecta
        oConec.AbreConexion
        Set Rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
        Set oConec = Nothing
        If Not Rs.EOF And Not Rs.BOF Then
            nCantidad = Rs!cantidad
            If nCantidad > 0 Then
                ValidaExtorno = 1
            Else
                ValidaExtorno = 0
            End If
        Else
            ValidaExtorno = -1
        End If
        Set Rs = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    If Not Rs Is Nothing Then Set Rs = Nothing
    ValidaExtorno = -1
End Function
Public Function Extorno(ByVal psCtaCod As String) As Boolean
    Dim oConec As DConecta
    
    Dim sSQL As String
    On Error GoTo ErrHandler
    Set oConec = New DConecta
        oConec.AbreConexion
        oConec.ConexionActiva.BeginTrans
        
        'Volviendo el credito a estado de sugerido
        sSQL = "Update Producto set nPrdEstado=2001 Where cCtaCod='" & psCtaCod & "'"
        oConec.ConexionActiva.Execute sSQL
        
        'Eliminado el estado de aprobado
        sSQL = "Delete From ColocacEstado Where cCtaCod='" & psCtaCod & "' and nPrdEstado=2002"
        oConec.ConexionActiva.Execute sSQL
        
        'Eliminado el ProductoPersona
        sSQL = "Delete from ProductoPersona Where nPrdPersRelac=29 and cctacod='" & psCtaCod & "'"
        oConec.ConexionActiva.Execute sSQL
        
        'Verificando si el tiene un credito mi vivienda
           If ObteMiVivienda(psCtaCod) = True Then
                sSQL = "Delete From ColocCalifMiViv Where cCtaCod='" & psCtaCod & "'"
                oConec.ConexionActiva.Execute sSQL
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

Function ObteMiVivienda(ByVal psCtaCod As String) As Boolean
    Dim oConec As DConecta
    Dim sSQL As String
    Dim Rs As ADODB.Recordset
    Dim nCantidad As Integer
    On Error GoTo ErrHandler
        sSQL = "Select count(*) as Cantidad From ColocacCred Where bMiVivienda=1 and cCtaCod='" & psCtaCod & "'"
        Set oConec = New DConecta
        oConec.AbreConexion
        Set Rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
        Set oConec = Nothing
        If Not Rs.BOF And Not Rs.EOF Then
            nCantidad = Rs!cantidad
        Else
            nCantidad = 0
        End If
        Set Rs = Nothing
        
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
