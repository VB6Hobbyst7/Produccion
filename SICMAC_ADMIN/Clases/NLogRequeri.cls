VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NLogRequeri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Agrega un Requerimiento
Public Function GrabaRequeri(ByVal psReqNro As String, ByVal psPeriodo As String, _
    ByVal psReqTpo As String, ByVal psNecesi As String, ByVal psRequeri As String, _
    ByVal psReqTraNro As String, ByVal psAreaCod As String, ByVal psReqTraCom As String, _
    ByVal psReqTraEst As String, ByVal psReqTraFlu As String, _
    ByVal poBSDet As ADODB.Recordset, ByVal poBSMes As ADODB.Recordset, _
    ByVal psActualizacion As String) As Integer

    Dim sBSCod As String
    Dim nRefPrecio As Currency, nCant As Currency
    Dim nBS As Integer, nBSMes As Integer
    Dim clsDMov As DLogMov
    Set clsDMov = New DLogMov
    'Inicializa valor de resultado
    GrabaRequeri = 1
    
    clsDMov.InsertaMov psReqNro, Trim(Str(gLogOpeReqRegistro)), "", Trim(Str(gLogReqEstadoInicio))
    
    'XXX
    clsDMov.InsertaRequeri psReqNro, psPeriodo, Trim(Str(gLogReqTipoNormal)), _
        psNecesi, psRequeri, psActualizacion
        
    clsDMov.InsertaReqTramite psReqNro, psReqNro, psAreaCod, _
        "", Trim(Str(gLogReqEstadoInicio)), gLogReqFlujoSin, psActualizacion
    nBS = 0: nBSMes = 0
    For nBS = 1 To poBSDet.RecordCount
        sBSCod = poBSDet.TextMatrix(nBS, 1)
        nRefPrecio = CCur(IIf(poBSDet.TextMatrix(nBS, 5) = "", 0, poBSDet.TextMatrix(nBS, 5)))
        clsDMov.InsertaReqDetalle psReqNro, psReqNro, sBSCod, _
            Trim(Right(poBSDet.TextMatrix(nBS, 4), 2)), nRefPrecio, 0, "", psActualizacion
        For nBSMes = 1 To poBSMes.Fields.Count
            nCant = CCur(IIf(poBSMes.TextMatrix(nBS, nBSMes) = "", 0, poBSMes.TextMatrix(nBS, nBSMes)))
            If nCant > 0 Then
                clsDMov.InsertaReqDetMes psReqNro, psReqNro, sBSCod, _
                     Val(fgeBSMes.TextMatrix(nBSMes, 1)), nCant
            End If
        Next
    Next
    'Ejecuta todos los querys en una transacción
    GrabaRequeri = clsDMov.EjecutaBatch
    Set clsDMov = Nothing
End Function


