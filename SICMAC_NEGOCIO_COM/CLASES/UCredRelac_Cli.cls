VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UCredRelac_Cli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'*****************************************************************************************
'***     Rutina:           UCredRelac_Cli
'***     Descripcion:       Manteniemiento de Personas Relacionadas con el Credito
'                           manejadas en la Maquina Cliente
'***     Creado por:        ARCV
'***     Maquina:           07DESARROLLO03
'***     Fecha-Tiempo:      07/12/2005 06:05:14 PM
'***     Ultima Modificacion: Creacion
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
    sSexo As String
    nPersoneria As Integer
    dPersNac As Date
End Type

Private MatRelaPersCred() As TRelaPersCred
Private nNumRelac As Integer
Public nPuntMat As Integer

Public Function ExisteTitular() As Boolean
Dim i As Integer
Dim bEnc As Boolean
    bEnc = False
    For i = 0 To nNumRelac - 1
        If CInt(MatRelaPersCred(i).sValorTipoRelac) = COMDConstantes.gColRelPersTitular Then
            bEnc = True
            Exit For
        End If
    Next i
    ExisteTitular = bEnc
End Function

Public Function ExisteConyuge() As Boolean
Dim i As Integer
Dim bEnc As Boolean
    bEnc = False
    For i = 0 To nNumRelac - 1
        If CInt(MatRelaPersCred(i).sValorTipoRelac) = COMDConstantes.gColRelPersConyugue Then
            bEnc = True
            Exit For
        End If
    Next i
    ExisteConyuge = bEnc
End Function

Public Function ExisteMasDeUnTitular() As Boolean
Dim i As Integer
Dim bEnc As Boolean
    bEnc = False
    For i = 0 To nNumRelac - 1
        If CInt(MatRelaPersCred(i).sValorTipoRelac) <> COMDConstantes.gColRelPersTitular Then
            bEnc = True
            Exit For
        End If
    Next i
    ExisteMasDeUnTitular = bEnc
End Function


Public Function TitularPersCod() As String
Dim i As Integer
   TitularPersCod = ""
    For i = 0 To nNumRelac - 1
        If CInt(MatRelaPersCred(i).sValorTipoRelac) = COMDConstantes.gColRelPersTitular Then
            TitularPersCod = MatRelaPersCred(i).sPersCod
            Exit For
        End If
    Next i
End Function

Public Function TitularNombre() As String
Dim i As Integer
   TitularNombre = ""
    For i = 0 To nNumRelac - 1
        If CInt(MatRelaPersCred(i).sValorTipoRelac) = COMDConstantes.gColRelPersTitular Then
            TitularNombre = MatRelaPersCred(i).sPersNom
            Exit For
        End If
    Next i
End Function


Public Function PosicionTitular() As Integer
Dim i As Integer
PosicionTitular = -1
    For i = 0 To nNumRelac - 1
        If CInt(MatRelaPersCred(i).sValorTipoRelac) = COMDConstantes.gColRelPersTitular Then
            PosicionTitular = i
            Exit For
        End If
    Next i
    
End Function

Public Sub CargaRelacPersCred(ByVal psCtaCod As String, _
                            Optional ByVal prsRelac As ADODB.Recordset = Nothing)

Dim oCredito As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
    On Error GoTo ErrorCargaRelacPersCred
    
    If prsRelac Is Nothing Then
        Set oCredito = New COMDCredito.DCOMCredito
        Set R = oCredito.RecuperaRelacPers(psCtaCod)
        Set oCredito = Nothing
        
        Do While Not R.EOF
            Call AdicionaRelacion(R!cPersCod, R!cPersNombre, R!cConsDescripcion, R!nConsValor, R!nConsValor, IIf(IsNull(R!DNI), "", Trim(R!DNI)), IIf(IsNull(R!Ruc), "", Trim(R!Ruc)), NuevaRegistro, R!nPrdEstado, IIf(IsNull(R!cPersnatSexo), "M", R!cPersnatSexo), R!nPersPersoneria, IIf(IsNull(R!dPersNacCreac), CDate("01/01/1901"), R!dPersNacCreac))
            R.MoveNext
        Loop
        R.Close
        Set R = Nothing
    Else
        Do While Not prsRelac.EOF
            Call AdicionaRelacion(prsRelac!cPersCod, prsRelac!cPersNombre, prsRelac!cConsDescripcion, prsRelac!nConsValor, prsRelac!nConsValor, IIf(IsNull(prsRelac!DNI), "", Trim(prsRelac!DNI)), IIf(IsNull(prsRelac!Ruc), "", Trim(prsRelac!Ruc)), NuevaRegistro, prsRelac!nPrdEstado, IIf(IsNull(prsRelac!cPersnatSexo), "M", prsRelac!cPersnatSexo), prsRelac!nPersPersoneria, IIf(IsNull(prsRelac!dPersNacCreac), CDate("01/01/1901"), prsRelac!dPersNacCreac))
            prsRelac.MoveNext
        Loop
        Set prsRelac = Nothing
    End If
    
    Exit Sub
ErrorCargaRelacPersCred:
    err.Raise err.Number, "Error En Proceso CargaRelacPersCred", err.Description
End Sub
Public Sub ActualizaRelacion(ByVal psPersCod As String, ByVal psTipoRelac As String, ByVal psValorRelac As String, ByVal psValorRelacAnt As String)
Dim i As Integer
    For i = 0 To nNumRelac - 1
        If MatRelaPersCred(i).sPersCod = psPersCod And MatRelaPersCred(i).sValorTipoRelac = psValorRelacAnt Then
            MatRelaPersCred(i).sTipoRelac = psTipoRelac
            MatRelaPersCred(i).sValorTipoRelac = psValorRelac
        End If
    Next i
End Sub

Public Sub AdicionaRelacion(ByVal psPersCod As String, ByVal psPersNom As String, _
    ByVal psTipoRelac As String, ByVal psValorRelac As String, ByVal psValorRelacAnt As String, _
    ByVal psDNI As String, ByVal psRUC As String, ByVal pTipoAct As TModifUCredRelacion, _
    ByVal nPrdEstado As Integer, ByVal psSexo As String, ByVal pnPersoneria As Integer, _
    Optional ByVal pdPersNac As Date = CDate("01/01/1900"))

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
    MatRelaPersCred(nNumRelac - 1).sSexo = psSexo
    MatRelaPersCred(nNumRelac - 1).nPersoneria = pnPersoneria
    MatRelaPersCred(nNumRelac - 1).dPersNac = Format(pdPersNac, "dd/mm/yyyy")
End Sub

Public Sub EliminarRelacion(ByVal psPersCod As String, ByVal psValorRelac As String)
Dim i As Integer
Dim J As Integer
    J = -1
    For i = 0 To nNumRelac - 1
        If MatRelaPersCred(i).sPersCod = psPersCod And Trim(MatRelaPersCred(i).sValorTipoRelac) = psValorRelac Then
            J = i
            Exit For
        End If
    Next i
    If J <> -1 Then
        For i = J To nNumRelac - 2
            MatRelaPersCred(i) = MatRelaPersCred(i + 1)
        Next i
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

Public Function ObtenerValorSexo() As String
    ObtenerValorSexo = MatRelaPersCred(nPuntMat).sSexo
End Function

Public Function ObtenerValorPersoneria() As Integer
    ObtenerValorPersoneria = MatRelaPersCred(nPuntMat).nPersoneria
End Function

Public Function ObtenerValorFechaNac() As Date
    ObtenerValorFechaNac = MatRelaPersCred(nPuntMat).dPersNac
End Function

Public Function ObtenerValorEdad(ByVal pdFecSis As Date) As Integer
    ObtenerValorEdad = DateDiff("yyyy", MatRelaPersCred(nPuntMat).dPersNac, pdFecSis)
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

Public Function ObtenerMatrizRelaciones() As Variant
Dim MatTemp As Variant
Dim i As Integer

    IniciarMatriz
    ReDim MatTemp(NroRelaciones, 3)
    i = 0
    Do While Not EOF
        MatTemp(i, 0) = ObtenerCodigo
        MatTemp(i, 1) = ObtenerValorRelac
        MatTemp(i, 2) = ObtenerValorRelacAnt
        i = i + 1
        siguiente
    Loop
    ObtenerMatrizRelaciones = MatTemp
End Function

Public Function ObtenerMatrizRelacionesRelacion(pnRelacionBuscar, nTipoBusqueda) As String
Dim lsDato As String
Dim i As Integer
    i = 0
    Do While Not EOF
        If pnRelacionBuscar = ObtenerValorRelac Then
            If nTipoBusqueda = 1 Then
                lsDato = ObtenerCodigo
            ElseIf nTipoBusqueda = 2 Then
                lsDato = CStr(ObtenerValorFechaNac)
            End If
        Exit Do
        End If
        i = i + 1
        siguiente
    Loop
    ObtenerMatrizRelacionesRelacion = lsDato
End Function

Public Function ObtenerDocumento() As String
Dim MatTemp As Variant
Dim i As Integer

    ObtenerDocumento = ""
    i = 0
    Do While Not EOF
        If ObtenerValorRelac = 20 Or ObtenerValorRelac = 21 Then
            'nPuntMat = i
            If Len(Trim(ObtenerDNI)) > 0 Then 'DNI
                ObtenerDocumento = ObtenerDocumento & ObtenerDNI & ","
            End If
            If Len(Trim(ObtenerRUC)) > 0 Then 'DNI
                ObtenerDocumento = ObtenerDocumento & ObtenerRUC & ","
            End If
        End If
        i = i + 1
        siguiente

    Loop
    If Len(Trim(ObtenerDocumento)) > 0 Then
        ObtenerDocumento = Mid(ObtenerDocumento, 1, Len(Trim(ObtenerDocumento)) - 1)
    End If
End Function
