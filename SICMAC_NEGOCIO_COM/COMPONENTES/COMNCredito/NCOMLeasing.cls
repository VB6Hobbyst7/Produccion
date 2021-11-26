VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMLeasing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function GeneraMovNro(ByVal pdFecha As Date, Optional ByVal psCodAge As String = "07", Optional ByVal psUser As String = "SIST", Optional psMovNro As String = "") As String
    On Error GoTo GeneraMovNroErr
    Dim rs As ADODB.Recordset
    Dim oConect As COMConecta.DCOMConecta
    Dim Sql As String
    Set oConect = New COMConecta.DCOMConecta
    Set rs = New ADODB.Recordset
    If oConect.AbreConexion = False Then Exit Function
    Sleep (1000)
    If psMovNro = "" Or Len(psMovNro) <> 25 Then
       Sql = "sp_GeneraMovNro '" & Format(pdFecha & " " & oConect.GetHoraServer, "mm/dd/yyyy hh:mm:ss") & "','" & Right(psCodAge, 2) & "','" & psUser & "'"
    Else
       Sql = "sp_GeneraMovNro '','','','" & psMovNro & "'"
    End If
    Set rs = oConect.Ejecutar(Sql)
    If Not rs.EOF Then
        GeneraMovNro = rs.Fields(0)
    End If
    rs.Close
    Set rs = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Function
GeneraMovNroErr:
    Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones:GeneraMovNro Method")
End Function

Public Function ObtenerComisionLeasingOtrasOperaciones(ByVal psCtaCod As String) As Currency
On Error GoTo ObtenerComisionLeasingOtrasOperacionesErr
    Dim oRs As ADODB.Recordset
    Dim lnMontoCuotaInicial As Currency
    
    Set oRs = New ADODB.Recordset
    Dim oLeasing As COMDCredito.DCOMleasing
    Set oLeasing = New COMDCredito.DCOMleasing
    Set oRs = oLeasing.ObtenerComisionLeasingOtrasOperacionesAD(psCtaCod)
    lnMontoCuotaInicial = 0
    If Not (oRs.BOF Or oRs.EOF) Then
        Do While Not oRs.EOF
            lnMontoCuotaInicial = lnMontoCuotaInicial + oRs!nMonto
            oRs.MoveNext
        Loop
    End If
    Set oRs = Nothing
    Set oLeasing = Nothing
    ObtenerComisionLeasingOtrasOperaciones = lnMontoCuotaInicial
    Exit Function
ObtenerComisionLeasingOtrasOperacionesErr:
    Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones:GeneraMovNro Method")

End Function

Public Function ValidaCreditoPersona(ByVal psCtaCod As String, ByVal psPersCod As String) As String
On Error GoTo ValidaCreditoPersonaErr
    Dim oRs As ADODB.Recordset
    
    Set oRs = New ADODB.Recordset
    Dim oLeasing As COMDCredito.DCOMleasing
    Set oLeasing = New COMDCredito.DCOMleasing
    Set oRs = oLeasing.ValidaCreditoPersona(psCtaCod, psPersCod)
    
    If Not (oRs.BOF Or oRs.EOF) Then
        ValidaCreditoPersona = ""
    Else
        ValidaCreditoPersona = "Operacion no pertenece al cliente "
    End If
    Set oRs = Nothing
    Set oLeasing = Nothing
    
    Exit Function
ValidaCreditoPersonaErr:
    Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones:GeneraMovNro Method")

End Function
'ORCR INICIO 20140414 ***
Public Function ObtenerMontoFinanciado(ByVal cCtaCod As String) As Currency
    Dim oDCOMleasing As New COMDCredito.DCOMleasing
On Error GoTo ErrorObtenerMontoFinanciado
    Dim lrDatos As ADODB.Recordset
    Set lrDatos = oDCOMleasing.ObtenerMontoFinanciado(cCtaCod)
    
    If lrDatos.RecordCount > 0 Then
        ObtenerMontoFinanciado = lrDatos!SAF_VALOR_BIENES
    Else
        ObtenerMontoFinanciado = 0
    End If
    
    Exit Function
ErrorObtenerMontoFinanciado:
    Err.Raise Err.Number, "NCOMLeasing:ObtenerMontoFinanciado", Err.Description
    ObtenerPatrimonioEfectivo = 0
End Function
'ORCR FIN 20140414 ***
