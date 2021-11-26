VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMColRecComision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim coConex As COMConecta.DCOMConecta
Dim oError As New COMConecta.COMErrorHandling
Private Sub Class_Initialize()
    Dim loIni As COMConecta.DCOMClasIni
    Dim csConexion As String
    Dim csNegocio As String
    Dim csCentralPer As String
    Dim csCentralCom As String
    Dim csCentralImg As String
    Dim csAdminist As String
    
    Set loIni = New COMConecta.DCOMClasIni
        csConexion = loIni.CadenaConexion
        csNegocio = loIni.BaseNegocio
        csCentralPer = loIni.BasePersonas
        csCentralCom = loIni.BaseComunes
        csCentralImg = loIni.BaseImagenes
        csAdminist = loIni.BaseAdministracion
    Set loIni = Nothing

Set coConex = New COMConecta.DCOMConecta
If coConex.AbreConexion(csConexion) = False Then
    Call oError.RaiseError(oError.MyUnhandledError, "DColPContrato:Initialize. Error en Conexion a Base de datos")
End If
End Sub

Private Sub Class_Terminate()
    coConex.CierraConexion
    Set coConex = Nothing
End Sub

Public Function nGeneraNuevaComision() As Integer
Dim lrs As New ADODB.Recordset
Dim lssql As String
Dim lnCorrelativo As Integer
lssql = "SELECT ISNULL(Max(nComisionCod),0) + 1 Correlativo FROM ColocRecupComision"

Set lrs = coConex.CargaRecordSet(lssql)
    If Not (lrs.EOF And lrs.BOF) Then
        lnCorrelativo = lrs("Correlativo")
    End If
Set lrs = Nothing
nGeneraNuevaComision = lnCorrelativo
End Function


