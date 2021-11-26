VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DLogSerCon"
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

Public Enum TpoCargaSC
    SCUnRegistro = 1
    SCTodosTpo = 5
End Enum

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

    Dim oIni As ClasIni
    Set oIni = New ClasIni
        vsConexion = oIni.CadenaConexion
        vsCentralPer = oIni.BasePersonas
        vsCentralCom = oIni.BaseComunes
        vsCentralImg = oIni.BaseImagenes
    Set oIni = Nothing
End Sub

'Devuelva un RecordSet con los Servicios-Contratos
Public Function CargaSerCon(ByVal pnTpoCarga As TpoCargaSC, ByVal pnTpo As Integer, _
Optional ByVal pnSerConNro As Long = 0) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset

    If oConec.AbreConexion() Then
        If pnTpoCarga = SCUnRegistro Then
            'Un registro
            sSql = "SELECT * " & _
                " FROM LogSerCon SC JOIN Mov M ON sc.nLogSerConNro = m.nMovNro " & _
                " WHERE nLogSerConTpo = " & pnTpo & " AND nLogSerConNro = " & pnSerConNro & "" & _
                "       AND m.nMovFlag <> " & gMovFlagEliminado & " "
        ElseIf pnTpoCarga = SCTodosTpo Then
            'Todos los registros para flex
            sSql = "SELECT nLogSerConNro, sc.cPersCod, p.cPersNombre, cLogSerConDescripcion, " & _
                "       Convert(varchar(12), dLogSerConInicio, 103), " & _
                "       Convert(varchar(12), dLogSerConFinal, 103), nLogSerConMonto " & _
                " FROM LogSerCon SC JOIN Mov M ON sc.nLogSerConNro = m.nMovNro " & _
                "       LEFT JOIN Persona P ON sc.cPersCod = p.cPersCod " & _
                " WHERE nLogSerConTpo = " & pnTpo & " AND m.nMovFlag <> " & gMovFlagEliminado & " " & _
                " ORDER BY nLogSerConNro "
        End If
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If

    Set CargaSerCon = rs
    Set rs = Nothing
End Function

'Devuelva un RecordSet con los Servicios-Contratos Detalle
Public Function CargaSerConDet(ByVal pnSerConNro As Long) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset

    If oConec.AbreConexion() Then
        sSql = "SELECT Convert(varchar(12), dLogSerConDetFecha, 103), nLogSerConDetImporte " & _
            " FROM LogSerConDetalle " & _
            " WHERE nLogSerConNro = " & pnSerConNro & " " & _
            " ORDER BY dLogSerConDetFecha "
        
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If

    Set CargaSerConDet = rs
    Set rs = Nothing
End Function

'Devuelva un RecordSet con los Servicios-Contratos Garantía
Public Function CargaSerConGar(ByVal pnSerConNro As Long) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset

    If oConec.AbreConexion() Then
        sSql = "SELECT nLogSerConGarTpo, c.cConsDescripcion, cLogSerConGarNro, cLogSerConGarDescripcion " & _
            " FROM LogSerConGarantia SCG LEFT JOIN Constante C ON scg.nLogSerConGarTpo = c.nConsValor " & _
            "       AND c.nConsCod = " & gPersGarantia & " " & _
            " WHERE nLogSerConNro = " & pnSerConNro & " "
        
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If

    Set CargaSerConGar = rs
    Set rs = Nothing
End Function
