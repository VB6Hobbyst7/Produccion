VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DResponsability"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim oConect As DConecta
Dim sSql As String
Public Sub dBeginTrans()
    oConect.BeginTrans
End Sub
Public Sub dCommitTrans()
    oConect.CommitTrans
End Sub
Public Sub dRollbackTrans()
    oConect.RollbackTrans
End Sub
Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    Set oImp = Nothing
Dim oIni As New ClasIni
   Set oIni = Nothing
   
   Set oConect = New DConecta
   If oConect.AbreConexion = False Then
        Call RaiseError(MyUnhandledError, "DResponsability:Initialize Method. Error en Conexion a Base de datos")
    End If
End Sub
Private Sub Class_Terminate()
    oConect.CierraConexion
    Set oConect = Nothing
End Sub
Public Function RegistrarConfigRepResponsability(ByVal psMes As String, ByVal psAnio As String, ByVal psMovNro As String, ByVal pnVig As Integer) As Integer
On Error GoTo ErrorRegistra
Dim rs As ADODB.Recordset
sSql = "stp_ins_ERS0872014_ConfigRepResponsability '" & psMes & "','" & psAnio & "','" & psMovNro & "'," & pnVig
    Set rs = oConect.CargaRecordSet(sSql)
    If Not (rs.EOF And rs.BOF) Then
             RegistrarConfigRepResponsability = rs!nIdRep
    End If
    Exit Function
ErrorRegistra:
    Err.Raise Err.Number, "Error al Registrar la Configuración", Err.Description
End Function
Public Function RegistrarConfigRepResponsabilityDet(ByVal pnIdRep As Integer, ByVal psConfRepDesc As String, ByVal psConfRepValor As String, Optional ByVal pnIdRepdetRef As Integer = 0) As Integer
On Error GoTo ErrorRegistra
Dim rs As ADODB.Recordset
sSql = "stp_ins_ERS0872014_ConfigRepResponsabilityDet " & pnIdRep & ",'" & psConfRepDesc & "','" & psConfRepValor & "'," & pnIdRepdetRef
    Set rs = oConect.CargaRecordSet(sSql)
    If Not rs.EOF And Not rs.BOF Then
             RegistrarConfigRepResponsabilityDet = rs!nIdRepDet
    End If
    Exit Function
ErrorRegistra:
    Err.Raise Err.Number, "Error al Registrar la Configuración", Err.Description
End Function
Public Function ExisteConfigRepResponsability(ByVal psMes As String, ByVal psAnio As String) As Boolean
    Dim rs As ADODB.Recordset
    sSql = "stp_sel_ERS0872014_ExisteConfigRepResponsability '" & psMes & "','" & psAnio & "'"
    oConect.AbreConexion
    Set rs = oConect.CargaRecordSet(sSql)
    If Not (rs.EOF And rs.BOF) Then
        ExisteConfigRepResponsability = IIf(rs!Existe <> 0, True, False)
    Else
        ExisteConfigRepResponsability = False
    End If
    Exit Function
End Function
Public Function ObtieneValorxRepResponsability(ByVal psAnio As String, ByVal psMes As String, ByVal psConfDesc As String) As String
    Dim rs As ADODB.Recordset
    sSql = "stp_sel_ERS0872014_ObtieneValorxRepResponsability '" & psAnio & "','" & psMes & "','" & psConfDesc & "'"
    Set rs = oConect.CargaRecordSet(sSql)
     If Not (rs.EOF And rs.BOF) Then
        ObtieneValorxRepResponsability = rs!Valor
    Else
        ObtieneValorxRepResponsability = ""
    End If
    Exit Function
End Function
Public Function ObtieneNIIFRepForResponsability(ByVal psOpecod As String, ByVal pnAnio As Integer, ByVal pnMes As Integer, ByVal pnCorreint As Integer) As String
    Dim rs As ADODB.Recordset
    sSql = "stp_sel_NIIFRepForResponsability '" & psOpecod & "'," & pnAnio & "," & pnMes & "," & pnCorreint
    Set rs = oConect.CargaRecordSet(sSql)
     If Not (rs.EOF And rs.BOF) Then
        ObtieneNIIFRepForResponsability = rs!cFormulaCons
    Else
        ObtieneNIIFRepForResponsability = ""
    End If
    Exit Function
End Function
'Número de prestatarios**************************
Public Function ObtieneValorNumeroPrestatario(ByVal psFecha As Date) As Currency
    Dim rs As ADODB.Recordset
    sSql = "stp_sel_ERS0872014_PrestAho_NumPresta '" & Format(psFecha, "yyyyMMdd") & "'"
    Set rs = oConect.CargaRecordSet(sSql)
    If Not (rs.EOF And rs.BOF) Then
        ObtieneValorNumeroPrestatario = CCur(rs!nResultado)
    Else
        ObtieneValorNumeroPrestatario = 0
    End If
    Exit Function
End Function
'********************************************
'Número de ahorrantes (excl. ahorros forzosos)**********
Public Function ObtieneValorNumeroAhorrantes(ByVal psFecha As Date) As Currency
    Dim rs As ADODB.Recordset
    sSql = "stp_sel_ERS0872014_PrestAho_NumAho '" & Format(psFecha, "yyyyMMdd") & "'"
    Set rs = oConect.CargaRecordSet(sSql)
    If Not (rs.EOF And rs.BOF) Then
        ObtieneValorNumeroAhorrantes = CCur(rs!nResultado)
    Else
        ObtieneValorNumeroAhorrantes = 0
    End If
    Exit Function
End Function
'********************************************
'Activos dado en garantia hacia los refinanciados**********
Public Function ObtieneValorActivGarantia(ByVal psFecha As Date) As Currency
Dim rs As ADODB.Recordset
sSql = "stp_sel_ERS0872014_GAP_ActivGarantia '" & Format(psFecha, "yyyyMMdd") & "'"
Set rs = oConect.CargaRecordSet(sSql)
    If Not (rs.EOF And rs.BOF) Then
        ObtieneValorActivGarantia = rs!nResultado
    Else
        ObtieneValorActivGarantia = 0
    End If
    Exit Function
End Function
'Resultado Neto del ejercicio despues de impuestos
Public Function ObtieneValorCalCartera_ResNetoEjercicioByVal(ByVal psFecha As Date) As Currency
Dim rs As ADODB.Recordset
sSql = "stp_sel_ERS0872014_CalCartera_ResNetoEjercicio '" & DatePart("YYYY", psFecha) & "','" & Format(DatePart("M", psFecha), "00") & "'"
Set rs = oConect.CargaRecordSet(sSql)
    If Not (rs.EOF And rs.BOF) Then
        ObtieneValorCalCartera_ResNetoEjercicioByVal = rs!nResultado
    Else
        ObtieneValorCalCartera_ResNetoEjercicioByVal = 0
    End If
    Exit Function
End Function

'PAR 1-30 Dias
Public Function ObtieneValorCalCartera_PAR1and30Dias(ByVal psFecha As Date) As Currency
Dim rs As ADODB.Recordset
sSql = "stp_sel_ERS0872014_CalCartera_PAR1and30Dias '" & Format(psFecha, "yyyyMMdd") & "'"
Set rs = oConect.CargaRecordSet(sSql)
    If Not (rs.EOF And rs.BOF) Then
        ObtieneValorCalCartera_PAR1and30Dias = rs!nResultado
    Else
        ObtieneValorCalCartera_PAR1and30Dias = 0
    End If
    Exit Function
End Function
'PAR>30 Dias
Public Function ObtieneValorCalCartera_PARMayor30Dias(ByVal psFecha As Date) As Currency
Dim rs As ADODB.Recordset
sSql = "stp_sel_ERS0872014_CalCartera_PARMayor30Dias '" & Format(psFecha, "yyyyMMdd") & "'"
Set rs = oConect.CargaRecordSet(sSql)
    If Not (rs.EOF And rs.BOF) Then
        ObtieneValorCalCartera_PARMayor30Dias = rs!nResultado
    Else
        ObtieneValorCalCartera_PARMayor30Dias = 0
    End If
    Exit Function
End Function
'Creditos reestructurados/reprogramados/refinanciados(no incluidos en PAR < 30)
Public Function ObtieneValorCalCartera_CredReesctruc(ByVal psFecha As Date) As Currency
Dim rs As ADODB.Recordset
sSql = "stp_sel_ERS0872014_CalCartera_CredReesctruc '" & Format(psFecha, "yyyyMMdd") & "'"
Set rs = oConect.CargaRecordSet(sSql)
    If Not (rs.EOF And rs.BOF) Then
        ObtieneValorCalCartera_CredReesctruc = rs!nResultado
    Else
        ObtieneValorCalCartera_CredReesctruc = 0
    End If
    Exit Function
End Function
'Total de Castigos
Public Function ObtieneValorCalCartera_TotalCastigos(ByVal psFecha As Date) As Currency
Dim rs As ADODB.Recordset
sSql = "stp_sel_ERS0872014_CalCartera_TotalCastigos '" & Format(psFecha, "yyyyMMdd") & "'"
Set rs = oConect.CargaRecordSet(sSql)
    If Not (rs.EOF And rs.BOF) Then
        ObtieneValorCalCartera_TotalCastigos = rs!nResultado
    Else
        ObtieneValorCalCartera_TotalCastigos = 0
    End If
    Exit Function
End Function
Public Function ListaRepResponsability() As ADODB.Recordset
    'Dim rs As ADODB.Recordset
    sSql = "stp_sel_ERS0872014_ListaRepResponsability"
    'Set rs = oConect.CargaRecordSet(sSql)
    If oConect.AbreConexion Then
        Set ListaRepResponsability = oConect.CargaRecordSet(sSql)
        oConect.CierraConexion
    Else
        Set ListaRepResponsability = Nothing
    End If
    Exit Function
End Function
Public Function ExtornaReporte(ByVal pnIdRep As Integer)
    sSql = "stp_sel_ERS0872014_ExtornaReporte " & pnIdRep
    oConect.Ejecutar sSql
    Exit Function
End Function
