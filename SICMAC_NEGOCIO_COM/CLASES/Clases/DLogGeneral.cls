VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DLogGeneral"
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

Public Enum TpoCargaUsu
    UsuUnRegistro = 0
    UsuTodosArea = 1
    UsuTodos = 2
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

'Devuelve un String con un Nro de Mov
Public Function GeneraMov(ByVal pdFecSis As Date, ByVal psCodCmac As String, _
    ByVal psCodAge As String, ByVal psCodUser As String) As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    If oConec.AbreConexion(vsConexion) Then
        GeneraMov = Format(pdFecSis & " " & oConec.GetHoraServer, "yyyymmddhhmmss") & psCodCmac & Right(psCodAge, 2) & "00" & psCodUser
    Else
        GeneraMov = ""
    End If
    Set oConec = Nothing
End Function

'Devuelve un RecordSet con las Areas
Public Function CargaArea() As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion(vsConexion) Then
        'Todos las Areas
        sSql = "SELECT a.cAreaCod, a.cAreaDescripcion, a.cAreaEstruc " & _
            " FROM  Areas A " & _
            " ORDER BY a.cAreaDescripcion "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set CargaArea = rs
End Function

'Devuelve un RecordSet con la relación usuario-area
Public Function CargaUsuario(ByVal pnTpoCarga As TpoCargaUsu, ByVal psUserArea As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion(vsConexion) Then
        If pnTpoCarga = UsuUnRegistro Then
            'Un usuario
            sSql = "SELECT rh.cAreaCod , a.cAreaDescripcion, p.cPersNombre, a.cAreaEstruc " & _
                " FROM  RRHH RH INNER JOIN Areas A ON rh.cAreaCod = a.cAreaCod " & _
                "       JOIN " & vsCentralPer & "PERSONA P ON P.cPersCod = RH.cPersCod  " & _
                " WHERE  rh.cUser = '" & psUserArea & "' "
        ElseIf pnTpoCarga = UsuTodosArea Then
            'Todos los usuarios por area
            sSql = "SELECT rh.cPersCod, rh.cUser, p.cPersNombre, p.cPersEstado " & _
                " FROM  RRHH RH INNER JOIN Areas A ON rh.cAreaCod = a.cAreaCod and rh.cAreaCod = '" & psUserArea & "'" & _
                "       JOIN " & vsCentralPer & "PERSONA P ON P.cPersCod = RH.cPersCod  " & _
                " ORDER BY p.cPersNombre "
        End If
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set CargaUsuario = rs
End Function

'Devuelve un RecordSet de acuerdo a la constante deseada
Public Function CargaConstante(ByVal psConstante As String, Optional ByVal pbPrimeroValor As Boolean = True) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion(vsConexion) Then
        sSql = IIf(pbPrimeroValor, "SELECT c.cConsValor, c.cConsDescripcion ", "SELECT c.cConsDescripcion, c.cConsValor ")
        sSql = sSql & _
            " FROM Constante C " & _
            " WHERE cConsCod = '" & psConstante & "' AND Not cConsValor = '" & psConstante & "'" & _
            " ORDER BY c.cConsValor "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set CargaConstante = rs
End Function


'Devuelve un RecordSet con el Area siguiente para el trámite
Public Function CargaAreaSuperior(ByVal psAreaStru As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion(vsConexion) Then
        sSql = " SELECT cAreaCod, cAreaDescripcion " & _
            " FROM Areas " & _
            " WHERE cAreaEstruc = substring('" & psAreaStru & "',1,len('" & psAreaStru & "') - 2) "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set CargaAreaSuperior = rs
End Function



