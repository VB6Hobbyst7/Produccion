VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMTipoCambioEsp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim vsConexion As String
Dim vsServerCom As String
Dim vsServerPers As String
Dim vsFormatoFecha As String
Dim oCon As COMConecta.DCOMConecta
Dim oerror As New COMConecta.COMErrorHandling

Private Sub Class_Initialize()

    Set oCon = New COMConecta.DCOMConecta
    
    If oCon.AbreConexion = False Then
        oerror.RaiseError oerror.MyUnhandledError, "DTipoCambio:Initialize Method. Error en Conexion de Base de datos"
        Exit Sub
    End If
End Sub

Private Sub Class_Terminate()
oCon.CierraConexion
Set oCon = Nothing
End Sub
Public Function InsertaTipoCambio(ByVal pdFecCamb As Date, ByVal pnValVent As Currency, ByVal pnValComp As Currency, _
ByVal pnIdRango As Integer, ByVal pnIdRangoDet As Integer) As Integer
    Dim Sql As String
    InsertaTipoCambio = 1
    
    'ARCV 10-06-2007

    'Sql = " Insert INTO TCMoneda"
    'Sql = Sql & " (dFecha,nIdRango,nIdRangoDet,nCompra,nVenta) values"
    'Sql = Sql & " ('" & Format(pdFecCamb, "MM/DD/YYYY hh:mm:ss AMPM") & "'"
    'Sql = Sql & " ," & pnIdRango & "," & pnIdRangoDet & "," & pnValComp & "," & pnValVent & ")"
    
    Dim rs As ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    
    Set oCon = New COMConecta.DCOMConecta
    Call oCon.AbreConexion
    Set rs = oCon.CargaRecordSet("SELECT*FROM TCMoneda WHERE DATEDIFF(d,dFecha,'" & Format(pdFecCamb, "MM/DD/YYYY") & "')=0 And nIdRango = " & pnIdRango & _
                            " AND nIdRangoDet=" & pnIdRangoDet)
    
    If Not rs.EOF Then
    
        Sql = " DELETE FROM TCMoneda"
        Sql = Sql & " WHERE DATEDIFF(d,dFecha,'" & Format(pdFecCamb, "MM/DD/YYYY") & "')=0 And nIdRango = " & pnIdRango & _
                      " AND nIdRangoDet=" & pnIdRangoDet
        
        oCon.Ejecutar Sql
    
    End If
    
    Sql = " Insert INTO TCMoneda"
    Sql = Sql & " (dFecha,nIdRango,nIdRangoDet,nCompra,nVenta) values"
    Sql = Sql & " ('" & Format(pdFecCamb, "MM/DD/YYYY") & "'"
    Sql = Sql & " ," & pnIdRango & "," & pnIdRangoDet & "," & pnValComp & "," & pnValVent & ")"
    
    oCon.Ejecutar Sql
    
    Call oCon.CierraConexion
    Set oCon = Nothing
    
    
    InsertaTipoCambio = 0
End Function
''->*** LUCV20180223, Se cambio de DCOMTCEspPermiso a este módulo de clases
'Public Function InsertaTipoCambioNew(ByVal psMovNro As String, ByVal psCodOpe As String, _
'                                      ByVal pnTCReg As Currency, ByVal pdFecSis As Date, _
'                                      ByVal pnMontoReg As Currency, ByVal psPersCod As String) As Integer
'    Dim Sql As String
'    Dim dFechaReg As String
'    InsertaTipoCambioNew = 1
'    Dim oCon As COMConecta.DCOMConecta
'    Set oCon = New COMConecta.DCOMConecta
'
'    Call oCon.AbreConexion
'    dFechaReg = Format(pdFecSis & " " & oCon.GetHoraServer, "MM/DD/YYYY hh:mm:ss AMPM")
'    Sql = "EXEC stp_ins_CapAutorizacionTC '" & psMovNro & "','" & psCodOpe & "'," & pnTCReg & ",'" & dFechaReg & "'," & pnMontoReg & ",'" & psPersCod & "'," & 0
'
'    oCon.Ejecutar Sql
'    Call oCon.CierraConexion
'    Set oCon = Nothing
'    InsertaTipoCambioNew = 0
'End Function
''<-***** Fin LUCV20180223


Public Function GetTipoCambioCV(ByVal pnMonto As Currency) As ADODB.Recordset
Dim Sql As String

Sql = " select top 3 TM.dFecha,TM.nIDRango,TM.nIdRangoDet,TM.nCompra,TM.nVenta, RD.nHasta from TCMoneda TM"
Sql = Sql & " Inner Join TCRangoDet RD on RD.nIdRangoDet = TM.nIdRangoDet"
Sql = Sql & " where TM.dfecha = (select Max(dfecha) from TCMoneda)"
Sql = Sql & " Order by nHasta"

Set GetTipoCambioCV = oCon.CargaRecordSet(Sql)
End Function

Public Function GetFechaIng() As Date
Dim Sql As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Sql = " Select max(dfecha) Fecha from TCMoneda"
Set rs = oCon.CargaRecordSet(Sql)
If Not (rs.EOF And rs.BOF) Then
    GetFechaIng = rs!Fecha
End If
Set rs = Nothing
End Function
Public Function GetTiposCambios() As ADODB.Recordset
Dim Sql As String
Sql = "select nHasta,dFecha,M.nIdRango,M.nIdRangoDet,nCompra,nVenta"
Sql = Sql & " from tcmoneda M"
Sql = Sql & " inner Join TCRangoDet RD on RD.nIdRango = M.nIdRango and  RD.nIdRangoDet = M.nIdRangoDet"
Sql = Sql & " and dFecha = (select max(dFEcha) from TCMoneda)"
Sql = Sql & " order by RD.nIdRangoDet"
Set GetTiposCambios = oCon.CargaRecordSet(Sql)
End Function

Public Function GetEstructuraTC() As ADODB.Recordset
Dim Sql As String

Sql = " Select RD.nIDRangoDet, R.nIDRango, nHasta"
Sql = Sql & " from TCRangoDet RD"
Sql = Sql & " Inner Join TCRango R on R.nIDRango = RD.nIDRango and R.nIDRango = (select Max(nIDRango) from TCRango)"
Sql = Sql & " Order by nIDRangoDet"
Set GetEstructuraTC = oCon.CargaRecordSet(Sql)
End Function
