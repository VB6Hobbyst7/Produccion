VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UCredRelacion"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'*****************************************************************************************
'***     Rutina:           UCredRelacion
'***     Descripcion:       Manteniemiento de Personas Relacionadas con el Credito
'***     Creado por:        NSSE
'***     Maquina:           07SIST_08
'***     Fecha-Tiempo:      31/05/2001 11:43:14 AM
'***     Ultima Modificacion: Crecion
'*****************************************************************************************
Option Explicit

Enum TModifUCredRelacion
    NuevaRegistro = 1
    Modificado = 2
    Eliminado = 3
End Enum

Private Type TRelaPersCred
    sPersCod As String
    sPersNom As String
    sTipoRelac As String 'Descripicion de la Relacion
    sValorTipoRelac As String
    sValorTipoRelacAnt As String
    sDNI As String
    sRUC As String
    nPrdEstado As Integer
    sModificado As TModifUCredRelacion
End Type

Private MatRelaPersCred() As TRelaPersCred
Private nNumRelac As Integer
Public nPuntMat As Integer

Public Sub CargaRelacPersCred(ByVal psCtaCod As String)
Dim oCredito As DCredito
Dim R As ADODB.Recordset
    On Error GoTo ErrorCargaRelacPersCred
    Set oCredito = New DCredito
    Set R = oCredito.RecuperaRelacPers(psCtaCod)
    Set oCredito = Nothing
    
    Do While Not R.EOF
        Call AdicionaRelacion(R!cPersCod, R!cPersNombre, R!cConsDescripcion, R!nConsValor, R!nConsValor, IIf(IsNull(R!DNI), "", Trim(R!DNI)), IIf(IsNull(R!RUC), "", Trim(R!RUC)), NuevaRegistro, R!nPrdEstado)
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Exit Sub

ErrorCargaRelacPersCred:
    Err.Raise Err.Number, "Error En Proceso CargaRelacPersCred", Err.Description
End Sub
Public Sub ActualizaRelacion(ByVal psPersCod As String, ByVal psTipoRelac As String, ByVal psValorRelac As String, ByVal psValorRelacAnt As String)
Dim I As Integer
    For I = 0 To nNumRelac - 1
        If MatRelaPersCred(I).sPersCod = psPersCod And MatRelaPersCred(I).sValorTipoRelac = psValorRelacAnt Then
            MatRelaPersCred(I).sTipoRelac = psTipoRelac
            MatRelaPersCred(I).sValorTipoRelac = psValorRelac
        End If
    Next I
End Sub

Public Sub AdicionaRelacion(ByVal psPersCod As String, ByVal psPersNom As String, _
    ByVal psTipoRelac As String, ByVal psValorRelac As String, ByVal psValorRelacAnt As String, _
    ByVal psDNI As String, ByVal psRUC As String, ByVal pTipoAct As TModifUCredRelacion, ByVal nPrdEstado As Integer)

    nNumRelac = nNumRelac + 1
    ReDim Preserve MatRelaPersCred(nNumRelac)
    MatRelaPersCred(nNumRelac - 1).sPersCod = psPersCod
    MatRelaPersCred(nNumRelac - 1).sPersNom = psPersNom
    MatRelaPersCred(nNumRelac - 1).sTipoRelac = psTipoRelac
    MatRelaPersCred(nNumRelac - 1).sModificado = pTipoAct
    MatRelaPersCred(nNumRelac - 1).sValorTipoRelac = psValorRelac
    MatRelaPersCred(nNumRelac - 1).sValorTipoRelacAnt = psValorRelacAnt
    MatRelaPersCred(nNumRelac - 1).nPrdEstado = nPrdEstado
    MatRelaPersCred(nNumRelac - 1).sDNI = psDNI
    MatRelaPersCred(nNumRelac - 1).sRUC = psRUC
End Sub

Public Sub EliminarRelacion(ByVal psPersCod As String, ByVal psValorRelac As String)
Dim I As Integer
Dim j As Integer
    j = -1
    For I = 0 To nNumRelac - 1
        If MatRelaPersCred(I).sPersCod = psPersCod And Trim(MatRelaPersCred(I).sValorTipoRelac) = psValorRelac Then
            j = I
            Exit For
        End If
    Next I
    If j <> -1 Then
        For I = j To nNumRelac - 2
            MatRelaPersCred(I) = MatRelaPersCred(I + 1)
        Next I
        nNumRelac = nNumRelac - 1
        ReDim Preserve MatRelaPersCred(nNumRelac)
    End If
End Sub

Public Function ObtenerNombre() As String
    ObtenerNombre = MatRelaPersCred(nPuntMat).sPersNom
End Function

Public Function ObtenerCodigo() As String
    ObtenerCodigo = MatRelaPersCred(nPuntMat).sPersCod
End Function

Public Function ObtenerRelac() As String
    ObtenerRelac = MatRelaPersCred(nPuntMat).sTipoRelac
End Function

Public Function ObtenerDNI() As String
    ObtenerDNI = MatRelaPersCred(nPuntMat).sDNI
End Function

Public Function ObtenerRUC() As String
    ObtenerRUC = MatRelaPersCred(nPuntMat).sRUC
End Function

Public Function ObtenerValorRelac() As String
    ObtenerValorRelac = MatRelaPersCred(nPuntMat).sValorTipoRelac
End Function

Public Function ObtenerValorEstado() As Integer
    ObtenerValorEstado = MatRelaPersCred(nPuntMat).nPrdEstado
End Function

Public Function ObtenerValorRelacAnt() As String
    ObtenerValorRelacAnt = MatRelaPersCred(nPuntMat).sValorTipoRelacAnt
End Function

Public Function ObtenerTipoAct() As String
    ObtenerTipoAct = MatRelaPersCred(nPuntMat).sModificado
End Function

Public Sub siguiente()
    nPuntMat = nPuntMat + 1
End Sub

Public Sub IniciarMatriz()
    nPuntMat = 0
End Sub
Public Function EOF() As Boolean
    If nNumRelac > nPuntMat Then
        EOF = False
    Else
        EOF = True
    End If
End Function
Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

    nNumRelac = 0
    ReDim MatRelaPersCred(0)
    nPuntMat = -1
End Sub

Public Property Get NroRelaciones() As Variant
    NroRelaciones = nNumRelac
End Property

Public Property Let NroRelaciones(ByVal vNewValue As Variant)
    nNumRelac = NroRelaciones
End Property
