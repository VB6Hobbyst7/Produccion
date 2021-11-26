VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DLogBieSer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim vsConexion As String
Dim vsCentralPer As String
Dim vsCentralCom As String
Dim vsCentralImg As String
Dim sSql As String

Public Enum TpoCargaBS
    BsUnRegistro = 0
    BsTodosArbol = 1
    BsTodosFlex = 2
    BsSuperiorFlex = 3
    BsTodos = 4
End Enum

Private Sub Class_Initialize()
    Dim oini As ClasIni
    
    Set oini = New ClasIni
        vsConexion = oini.CadenaConexion
        vsCentralPer = oini.BasePersonas
        vsCentralCom = oini.BaseComunes
        vsCentralImg = oini.BaseImagenes
    Set oini = Nothing
End Sub

'Devuelva un RecordSet con los Bienes/Servicios
Public Function CargaBS(ByVal pnTpoCarga As TpoCargaBS, Optional ByVal psCodBS As String = "") As ADODB.Recordset
    Dim oConec As DConecta
    Dim Rs As ADODB.Recordset
    Set oConec = New DConecta
    Set Rs = New ADODB.Recordset
    
    If oConec.AbreConexion(vsConexion) Then
        If pnTpoCarga = BsUnRegistro Then
            'Un registro
            sSql = "SELECT bs.cBSCod, bs.cBSDescripcion, " & _
                "       IsNull((ltrim(c.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad , " & _
                "       bs.nBSStock " & _
                " FROM BienesServicios BS LEFT JOIN Constante C On bs.nBSUnidad = c.cConsValor " & _
                "           AND cConsCod = " & gUnidadMedida & " " & _
                " WHERE bs.cBSCod = '" & psCodBS & "' "
        ElseIf pnTpoCarga = BsTodosArbol Then
            'Todos los registros para el Arbol
            sSql = "SELECT bs.cBSCod, bs.cBSDescripcion Descripción, len(bs.cBSCod) " & _
                " FROM BienesServicios BS " & _
                " ORDER BY bs.cBSCod "
        ElseIf pnTpoCarga = BsTodosFlex Then
            'Todos los registros para flex
            sSql = "SELECT bs.cBSCod Código, bs.cBSCod Código, bs.cBSDescripcion Descripción " & _
                " FROM BienesServicios BS " & _
                " ORDER BY bs.cBSCod "
        ElseIf pnTpoCarga = BsSuperiorFlex Then
            'Registros superiores para flex
            sSql = "SELECT bs.cBSCod Código, bs.cBSCod Código, bs.cBSDescripcion Descripción " & _
                " FROM BienesServicios BS WHERE bs.cBSCod LIKE '" & psCodBS & "_%' AND " & _
                " LEN(bs.cBSCod) = (SELECT MIN(LEN(cBSCod)) FROM BienesServicios WHERE cBSCod LIKE '" & psCodBS & "_%') " & _
                " ORDER BY bs.cBSCod "
        ElseIf pnTpoCarga = BsTodos Then
            'Todos los registros para flex
            sSql = "SELECT bs.cBSCod , bs.cBSDescripcion, c.cConsDescripcion" & _
                " FROM BienesServicios BS LEFT JOIN Constante C On bs.nBSUnidad = c.cConsValor " & _
                "           AND cConsCod = " & gUnidadMedida & " " & _
                " ORDER BY bs.cBSCod "
        End If
        Set Rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set CargaBS = Rs
    Set Rs = Nothing
End Function

'Agrega un Bien/Servicio
Public Function GrabaBS(ByVal psBSCod As String, ByVal psDescripcion As String, _
ByVal pnUnidad As Integer, ByVal pnStock As Currency, ByVal psActualizacion As String) As Integer
    Dim sActualiza As String
    Dim oConec As DConecta
    Dim Rs As Recordset
    
    Set oConec = New DConecta
    
    If oConec.AbreConexion(vsConexion) Then
        sSql = "SELECT cBSCod FROM BienesServicios " & _
            " WHERE cBSCod = '" & psBSCod & "' "
        Set Rs = oConec.CargaRecordSet(sSql)
        
        If Rs.RecordCount = 0 Then
            sActualiza = Left(psActualizacion, 8) & Format(oConec.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
            
            sSql = "INSERT INTO BienesServicios(cBSCod, cBSDescripcion, " & _
                " nBSUnidad, nBSStock, cUltimaActualizacion) " & _
                " VALUES ('" & psBSCod & "','" & psDescripcion & "'," & _
                pnUnidad & "," & pnStock & ",'" & sActualiza & "')"
            oConec.Ejecutar sSql
            GrabaBS = 0
        Else
            GrabaBS = 2
        End If
        oConec.CierraConexion
    Else
        GrabaBS = 1
    End If
End Function

'Modifica un Bien/Servicio
Public Function ModificaBS(ByVal psBSCod As String, ByVal psDescripcion As String, _
ByVal pnUnidad As Integer, ByVal pnStock As Currency, ByVal psActualizacion As String) As Integer
    Dim sActualiza As String
    Dim oConec As DConecta
    Dim Rs As Recordset
    
    Set oConec = New DConecta
    
    If oConec.AbreConexion(vsConexion) Then
        sSql = "SELECT cBSCod FROM BienesServicios " & _
            " WHERE cBSCod = '" & psBSCod & "' "
        Set Rs = oConec.CargaRecordSet(sSql)
        
        If Rs.RecordCount = 1 Then
            sActualiza = Left(psActualizacion, 8) & Format(oConec.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
            
            sSql = "UPDATE BienesServicios SET cBSDescripcion = '" & psDescripcion & "', " & _
                " nBSUnidad = " & pnUnidad & ", nBSStock = " & pnStock & ", " & _
                " cUltimaActualizacion = '" & sActualiza & "' " & _
                " WHERE cBSCod = '" & psBSCod & "'"
            oConec.Ejecutar sSql
            ModificaBS = 0
        Else
            ModificaBS = 2
        End If
        oConec.CierraConexion
    Else
        ModificaBS = 1
    End If
End Function

'Elimina un Bien/Servicio y su relacion con ProveedorBS
Public Function EliminaBS(ByVal psBSCod As String) As Integer
    Dim oConec As DConecta
    Set oConec = New DConecta
    
    If oConec.AbreConexion(vsConexion) Then
        sSql = "DELETE ProveedorBS WHERE cBSCod = '" & psBSCod & "'"
        oConec.Ejecutar sSql
        
        sSql = "DELETE BienesServicios WHERE cBSCod = '" & psBSCod & "'"
        oConec.Ejecutar sSql
        
        oConec.CierraConexion
        EliminaBS = 0
    Else
        EliminaBS = 1
    End If
End Function

'Devuelve un nuevo código para los Bienes/Servicios
Public Function GeneraBSCodNue(ByVal psCodBSPadre As String) As String
    Dim oConec As DConecta
    Dim Rs As ADODB.Recordset
    Set oConec = New DConecta
    Set Rs = New ADODB.Recordset
    Dim sValor As String
    
    If oConec.AbreConexion(vsConexion) Then
        sSql = "SELECT isnull(MAX(bs.cBSCod),0) Campo" & _
            " FROM BienesServicios BS WHERE bs.cBSCod LIKE '" & psCodBSPadre & "_%' AND " & _
            " LEN(bs.cBSCod) = (SELECT MIN(LEN(cBSCod)) FROM BienesServicios WHERE cBSCod LIKE '" & psCodBSPadre & "_%')"
        
        Set Rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
        sValor = Rs!Campo
        If Val(sValor) > 0 Then
            GeneraBSCodNue = Val(sValor) + 1
        Else
            GeneraBSCodNue = psCodBSPadre + "01"
        End If
    Else
        GeneraBSCodNue = ""
    End If
End Function

