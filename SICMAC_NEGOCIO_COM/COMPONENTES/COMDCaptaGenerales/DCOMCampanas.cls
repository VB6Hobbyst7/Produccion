VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMCampanas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public dbCmact As ADODB.Connection
Dim sDBComunes As String
Dim sDBPersona As String
Dim sDBImagenes As String

Public Function GetCapCampanas(ByVal pcProd As String) As ADODB.Recordset
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    
    sSql = "Select cDescripcion, IdCampana, bEstado From Campanas Where cProd = '" & Trim(pcProd) & "' Order By cDescripcion"
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
    Set GetCapCampanas = rs
    Set rs.ActiveConnection = Nothing
End Function

Public Sub RegCampana(ByVal psDesc As String, ByVal pEstcamp As Integer, ByVal psProd As String)
    Dim sSql As String
    sSql = "Insert Into Campanas(cDescripcion, bEstado,cProd)"
    sSql = sSql & "Values ('" & psDesc & "'," & pEstcamp & ",'" & psProd & "')"
    dbCmact.Execute sSql
End Sub

Public Sub RegPremio(ByVal psDesc As String, ByVal pnCost As Double, ByVal pnCant As Integer, ByVal pnRefe As Integer)
    Dim sSql As String
    sSql = "Insert Into CapPremio(cDescripcion, nMontoPremio,nStockActual, nStockRef)"
    sSql = sSql & "Values ('" & psDesc & "'," & pnCost & "," & pnCant & "," & pnRefe & ")"
    dbCmact.Execute sSql
End Sub

Public Sub RegCampanaPremio(ByVal pnIdCamp As Integer, ByVal pnTipPre As Integer, ByVal pnCant As Integer)
    Dim sSql As String
    sSql = "Insert Into Capcampanapremio(nIdCampana, nTipoPremio, nCantidad)"
    sSql = sSql & "Values (" & pnIdCamp & "," & pnTipPre & "," & pnCant & ")"
    dbCmact.Execute sSql
End Sub

Public Function GetCapCampanaPremio(ByVal pnIdCamp As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    sSql = "Select r.nTipoPremio, r.nCantidad, p.cDescripcion, p.nMontoPremio "
    sSql = sSql & "From capcampanapremio r join cappremio p on r.nTipoPremio = p.nTipoPremio "
    sSql = sSql & "Where r.nIdCampana = " & pnIdCamp
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
    Set GetCapCampanaPremio = rs
    Set rs.ActiveConnection = Nothing
End Function

Public Function GetCapPremios() As ADODB.Recordset
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    sSql = "Select nTipoPremio, cDescripcion From CapPremio Order by cDescripcion"
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
    Set GetCapPremios = rs
    Set rs.ActiveConnection = Nothing
End Function

Public Sub RegCtaPremio(ByVal pCta As String, ByVal pdFecha As String, ByVal pnIdCamp As Integer, ByVal pnTipPre As Integer, ByVal psUser As String)
    Dim sSql As String
    
    On Error GoTo ErrGraba
    dbCmact.BeginTrans
        'Disminuimos la cantidad del premio en la campaña
        sSql = "Update CapCampanaPremio Set nCantidad = nCantidad - 1 Where nIdCampana = " & pnIdCamp & " And nTipoPremio = " & pnTipPre
        sSql = sSql & Chr(13)
    
        'Registramos el premio asignado a la Cuenta
        sSql = sSql & "Insert Into CapPremioDet(cCtaCod, dFechaEntrega, nIdCampana, nTipoPremio, nCantidad, cUserEntrega, nExonerado)"
        sSql = sSql & "Values('" & pCta & "','" & Format(pdFecha, "yyyymmdd") & "'," & pnIdCamp & "," & pnTipPre & ",1,'" & psUser & "',0)"
        dbCmact.Execute sSql
    dbCmact.CommitTrans
    
    Exit Sub
    
ErrGraba:
    dbCmact.RollbackTrans
    Err.Raise Err.Number, "", Err.Description
    
End Sub

Public Function GetMontoPremio(ByVal psCuenta As String) As Double
    Dim rs As New ADODB.Recordset
    Dim sSql As String
    
    Set rs = New ADODB.Recordset
    sSql = "Select IsNull(p.nMontoPremio,0) as nMontoPremio From CapPremio p Inner Join cappremiodet cp on p.nTipoPremio = cp.nTipoPremio "
    sSql = sSql & "where cp.cctacod = '" & psCuenta & "'"
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rs.BOF And rs.EOF) Then
        GetMontoPremio = rs.Fields("nMontoPremio")
    Else
        GetMontoPremio = 0
    End If
    Set rs = Nothing
    Exit Function
End Function

Private Sub Class_Initialize()
    Dim sConn As String
    Dim ClsIni As COMConecta.DCOMClasIni
    Set ClsIni = New COMConecta.DCOMClasIni
    sConn = ClsIni.CadenaConexion
    sDBComunes = ClsIni.BaseComunes
    sDBPersona = ClsIni.BasePersonas
    sDBImagenes = ClsIni.BaseImagenes
    Set ClsIni = Nothing
    Set dbCmact = New ADODB.Connection
    dbCmact.Open sConn
    dbCmact.Execute "SET DATEFORMAT MDY"
End Sub

Private Sub Class_Terminate()
    dbCmact.Close
    Set dbCmact = Nothing
End Sub
