VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DColRecComision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim csConexion As String
Dim csNegocio As String
Dim csCentralPer As String
Dim csCentralCom As String
Dim csCentralImg As String
Dim csAdminist As String

Dim coConex As DConecta
Private Sub Class_Initialize()
    Dim loIni As ClasIni
    
    Set loIni = New ClasIni
        csConexion = loIni.CadenaConexion
        csNegocio = loIni.BaseNegocio
        csCentralPer = loIni.BasePersonas
        csCentralCom = loIni.BaseComunes
        csCentralImg = loIni.BaseImagenes
        csAdminist = loIni.BaseAdministracion
    Set loIni = Nothing

Set coConex = New DConecta
If coConex.AbreConexion(csConexion) = False Then
    Call RaiseError(MyUnhandledError, "DColPContrato:Initialize. Error en Conexion a Base de datos")
End If
End Sub

Private Sub Class_Terminate()
    coConex.CierraConexion
    Set coConex = Nothing
End Sub

Public Function nGeneraNuevaComision() As Integer
Dim lrs As Recordset
Dim lsSQL As String
Dim lnCorrelativo As Integer
lsSQL = "SELECT ISNULL(Max(nComisionCod),0) + 1 Correlativo FROM ColocRecupComision"

Set lrs = coConex.CargaRecordSet(lsSQL)
    If Not (lrs.EOF And lrs.BOF) Then
        lnCorrelativo = lrs("Correlativo")
    End If
Set lrs = Nothing
nGeneraNuevaComision = lnCorrelativo
End Function




