VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DClassIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim psArchivoIni As String
Dim psServerComunes As String
Dim psServerPersonas As String
Dim psServerImagenes As String
Dim psServerNegocio As String
Dim psServerAdmin As String


Public Function CadenaConexion(Optional psArchivo As String = "", Optional psAplicacion As String = "SICMACT") As String
Dim lsProvider As String
Dim lsServer As String
Dim lsUser As String
Dim lsPassword As String
Dim lsDatabase As String
Dim lsCadena As String

On Error GoTo ErrorCadCon
If psArchivo = "" Then
    psArchivo = App.Path & "\SICMACTF.INI"
End If
   
lsProvider = LeerArchivoIni(Encripta(psAplicacion), Encripta("Provider"), psArchivo)
lsServer = LeerArchivoIni(Encripta(psAplicacion), Encripta("Server"), psArchivo)
lsUser = LeerArchivoIni(Encripta(psAplicacion), Encripta("User"), psArchivo)
lsPassword = LeerArchivoIni(Encripta(psAplicacion), Encripta("Password"), psArchivo)
lsDatabase = LeerArchivoIni(Encripta(psAplicacion), Encripta("DataBase"), psArchivo)
'ANTES
lsCadena = "PROVIDER=" & Trim(Encripta(lsProvider, False)) & ";User ID=" & Trim(Encripta(lsUser, False)) & ";Password=" & Trim(Encripta(lsPassword, False)) & ";INITIAL CATALOG=" & Trim(Encripta(lsDatabase, False)) & ";DATA SOURCE=" & Trim(Encripta(lsServer, False)) & ""
'CUSCO
'lsCadena = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & Trim(Encripta(lsDatabase, False)) & ";Data Source=" & Trim(Encripta(lsServer, False)) & ""
CadenaConexion = lsCadena

Exit Function
ErrorCadCon:
    Err.Raise Err.Number, "Cadena de Conexion", "Error al Generar cadena de Conexion. Revise su Archivo de Inicio" & vbCrLf & "DESCRIPCION DEL ERROR:" & vbCrLf & Err.Description
End Function
'Public Sub CrearArchivoIni(Optional psArchivoIni As String = "", Optional lsAplicacion As String = "SCIMACT")
'    If psArchivoIni = "" Then
'        psArchivoIni = App.Path & "\SICMACT.INI"
'    End If
'    frmIni.Inicio "SICMACT", psArchivoIni
'    psArchivoIni = frmIni.lsArchivoIni
'    Unload frmIni
'    Set frmIni = Nothing
'End Sub

Public Function BaseComunes(Optional psArchivo As String = "", Optional psAplicacion As String = "BASE COMUNES") As String
Dim lsBaseComunes As String

On Error GoTo ErrorCadCon
If psArchivo = "" Then
    psArchivo = App.Path & "\SICMACTF.INI"
End If
lsBaseComunes = Encripta(LeerArchivoIni(Encripta(psAplicacion), Encripta("dbComunes"), psArchivo), False)
BaseComunes = lsBaseComunes
Exit Function
ErrorCadCon:
    Err.Raise Err.Number, "Base Comunes", "Error en Base Comunes" & vbCrLf & "DESCRIPCION DEL ERROR:" & vbCrLf & Err.Description
End Function
Public Function BasePersonas(Optional psArchivo As String = "", Optional psAplicacion As String = "BASE COMUNES") As String
Dim lsBase As String

On Error GoTo ErrorCadCon
If psArchivo = "" Then
    psArchivo = App.Path & "\SICMACTF.INI"
End If
lsBase = Encripta(LeerArchivoIni(Encripta(psAplicacion), Encripta("dbPersonas"), psArchivo), False)
BasePersonas = lsBase
Exit Function
ErrorCadCon:
    Err.Raise Err.Number, "Base Personas", "Error en Base Personas" & vbCrLf & "DESCRIPCION DEL ERROR:" & vbCrLf & Err.Description
End Function
Public Function BaseImagenes(Optional psArchivo As String = "", Optional psAplicacion As String = "BASE COMUNES") As String
Dim lsBase As String

On Error GoTo ErrorCadCon
If psArchivo = "" Then
    psArchivo = App.Path & "\SICMACTF.INI"
End If
lsBase = Encripta(LeerArchivoIni(Encripta(psAplicacion), Encripta("dbImagenes"), psArchivo), False)
BaseImagenes = lsBase
Exit Function
ErrorCadCon:
    Err.Raise Err.Number, "Base Imagenes", "Error en Base Imagenes" & vbCrLf & "DESCRIPCION DEL ERROR:" & vbCrLf & Err.Description
End Function

Public Function BaseNegocio(Optional psArchivo As String = "", Optional psAplicacion As String = "BASE COMUNES") As String
Dim lsBase As String

On Error GoTo ErrorCadCon
If psArchivo = "" Then
    psArchivo = App.Path & "\SICMACTN.INI"
End If
lsBase = Encripta(LeerArchivoIni(Encripta(psAplicacion), Encripta("dbNegocio"), psArchivo), False)
BaseNegocio = lsBase
Exit Function
ErrorCadCon:
    Err.Raise Err.Number, "Base Negocios", "Error en Base de Negocios" & vbCrLf & "DESCRIPCION DEL ERROR:" & vbCrLf & Err.Description
End Function
Public Function BaseAdministracion(Optional psArchivo As String = "", Optional psAplicacion As String = "BASE COMUNES") As String
Dim lsBase As String

On Error GoTo ErrorCadCon
If psArchivo = "" Then
    psArchivo = App.Path & "\SICMACTA.INI"
End If
lsBase = Encripta(LeerArchivoIni(Encripta(psAplicacion), Encripta("dbAdmin"), psArchivo), False)
BaseAdministracion = lsBase
Exit Function
ErrorCadCon:
    Err.Raise Err.Number, "Base Administración", "Error en Base de Administración" & vbCrLf & "DESCRIPCION DEL ERROR:" & vbCrLf & Err.Description
End Function

Public Function LeerArchivoIni(Seccion As String, Dato As String, lsArchivo As String) As String
    Dim strValue As String
    Dim lngRetLen As Long
    
    strValue = String(READ_BUFF + 1, Space(1))
    lngRetLen = GetPrivateProfileStringByKeyName(Seccion, Dato, "", strValue, READ_BUFF, lsArchivo)
    If lngRetLen > 1 Then
        LeerArchivoIni = Left(strValue, lngRetLen)
    Else
        LeerArchivoIni = ""
    End If
End Function

Public Function Encripta(psTexto As String, Optional psValor As Boolean = True) As String
'true = encripta
'false = desencripta
Dim oEnc As cEncrypt
Set oEnc = New cEncrypt
Encripta = oEnc.ConvertirClave(psTexto, , psValor)
End Function



