VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NPersona"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Function FillNum(intNumero As String, intLenNum As Integer, ChrFil As String) As String
  FillNum = Left(String(intLenNum, ChrFil), (Len(String(intLenNum, ChrFil)) - Len(Trim(intNumero)))) + Trim(intNumero)
End Function

Public Function DigitoChequeo(ByVal psCorrelativo As String) As Integer
Dim liFactor As Variant
Dim liCadena(1 To 7) As Integer
Dim liSum, I As Integer
Dim lnDigito As Integer
    liFactor = Array(2, 7, 6, 5, 4, 3, 2)
    liCadena(1) = Val(Mid(psCorrelativo, 1, 1))
    liCadena(2) = Val(Mid(psCorrelativo, 2, 1))
    liCadena(3) = Val(Mid(psCorrelativo, 3, 1))
    liCadena(4) = Val(Mid(psCorrelativo, 4, 1))
    liCadena(5) = Val(Mid(psCorrelativo, 5, 1))
    liCadena(6) = Val(Mid(psCorrelativo, 6, 1))
    liCadena(7) = Val(Mid(psCorrelativo, 7, 1))
    liSum = 0
    For I = 1 To 7
        liSum = liSum + liCadena(I) * liFactor(I - 1)
    Next I
    lnDigito = 11 - (liSum Mod 11)
    If lnDigito = 10 Then
        DigitoChequeo = 0
    ElseIf lnDigito = 11 Then
        DigitoChequeo = 1
    Else
        DigitoChequeo = lnDigito
    End If
End Function
Public Function GenerarNuevoCodigoPersona(ByVal poOpersona As DPersona, ByVal psCodAge As String) As String
Dim lsUltCorr As String
Dim lnCorr As Long
Dim lnCodchequeo As Integer
Dim lsCadFinal As String

    lsUltCorr = poOpersona.UltimoCorrelativo(psCodAge)
    lnCorr = Val(lsUltCorr) + 1
    lsUltCorr = FillNum(Trim(Str(lnCorr)), 7, "0")
    lnCodchequeo = DigitoChequeo(lsUltCorr)
    lsCadFinal = lsUltCorr & Trim(Str(lnCodchequeo))
    GenerarNuevoCodigoPersona = Trim(psCodAge) & Trim(lsCadFinal)
End Function

Public Sub ChequeoFuenteIngreso(ByRef poOpersona As DPersona, ByVal pnIndice As Integer, ByVal pnTipoFte As PersFteIngresoTipo, ByVal pnPuntDet As Integer)
Dim dFteFecEval As Date
Dim DiasLimite As Integer
Dim dFteFecCaduca As Date

    dFteFecEval = poOpersona.ObtenerFteIngFecEval(pnIndice, pnPuntDet, pnTipoFte)
    DiasLimite = poOpersona.TiempoFteIngreso
    dFteFecCaduca = CDate(Format(dFteFecEval, "dd/mm/yyyy")) + DiasLimite
    Call poOpersona.ActualizarFteIngFecCaduc(dFteFecCaduca, pnIndice, pnPuntDet, pnTipoFte)
End Sub

Public Function VerificandoDuplicidadDoc(ByVal popersona As DPersona) As Integer
Dim oPersonas As DPersonas
Dim I As Integer

    On Error GoTo ErrorVeirficandoDuplicidadDoc
    Set oPersonas = New DPersonas
    VerificandoDuplicidadDoc = -1
    For I = 0 To popersona.NumeroDocumentos - 1
           If popersona.ObtenerDocTipoAct(I) <> PersFilaSinCambios And popersona.ObtenerDocTipoAct(I) <> PersFilaEliminda Then
             If oPersonas.BuscaDocumento(CInt(Trim(Right(popersona.ObtenerTipoDoc(I), 15))), popersona.ObtenerNumeroDoc(I), popersona.PersCodigo) Then
                VerificandoDuplicidadDoc = I
                Set oPersonas = Nothing
                Exit Function
             End If
         End If
    Next I
    Set oPersonas = Nothing
    Exit Function
    
ErrorVeirficandoDuplicidadDoc:
    Set oPersonas = Nothing
    Err.Raise Err.Number, "BuscaDocumento", Err.Description
End Function

Public Function VerificandoTamanioDoc(ByVal popersona As DPersona) As Integer
Dim I As Integer

    On Error GoTo ErrorVerificandoTamanioDoc
    VerificandoTamanioDoc = -1
    For I = 0 To popersona.NumeroDocumentos - 1
         If popersona.ObtenerDocTipoAct(I) <> PersFilaSinCambios And popersona.ObtenerDocTipoAct(I) <> PersFilaEliminda Then
            If (Len(popersona.ObtenerNumeroDoc(I)) <> 8 And CInt(Trim(Right(popersona.ObtenerTipoDoc(I), 15))) = gPersIdDNI) Or (Len(popersona.ObtenerNumeroDoc(I)) <> 11 And CInt(Trim(Right(popersona.ObtenerTipoDoc(I), 15))) = gPersIdRUC) Then
               VerificandoTamanioDoc = I
               Exit Function
            End If
        End If
    Next I
    Exit Function
    
ErrorVerificandoTamanioDoc:
    Err.Raise Err.Number, "BuscaDocumento", Err.Description
End Function

Public Function VerificandoHomonimia(ByVal popersona As DPersona) As ADODB.Recordset
Dim oPersonas As DPersonas
Dim R As ADODB.Recordset
Dim sNomTmp As String

    On Error GoTo ErrorVerificandoTamanioDoc
    If popersona.Personeria = 1 Then
        sNomTmp = popersona.ApellidoPaterno & "/" & popersona.ApellidoMaterno
    Else
        sNomTmp = popersona.NombreCompleto
    End If
    Set oPersonas = New DPersonas
    Set R = oPersonas.BuscaCliente(sNomTmp, BusquedaNombre)
    Set VerificandoHomonimia = R
    Set oPersonas = Nothing
    Exit Function
    
ErrorVerificandoTamanioDoc:
    Err.Raise Err.Number, "BuscaDocumento", Err.Description
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

End Sub

