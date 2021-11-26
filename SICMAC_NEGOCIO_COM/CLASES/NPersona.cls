VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NPersona"
Attribute VB_GlobalNameSpace = False
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
Dim liSum, i As Integer
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
    For i = 1 To 7
        liSum = liSum + liCadena(i) * liFactor(i - 1)
    Next i
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
Dim i As Integer

    On Error GoTo ErrorVeirficandoDuplicidadDoc
    Set oPersonas = New DPersonas
    VerificandoDuplicidadDoc = -1
    For i = 0 To popersona.NumeroDocumentos - 1
           If popersona.ObtenerDocTipoAct(i) <> PersFilaSinCambios And popersona.ObtenerDocTipoAct(i) <> PersFilaEliminda Then
             If oPersonas.BuscaDocumento(CInt(Trim(Right(popersona.ObtenerTipoDoc(i), 15))), popersona.ObtenerNumeroDoc(i), popersona.PersCodigo) Then
                VerificandoDuplicidadDoc = i
                Set oPersonas = Nothing
                Exit Function
             End If
         End If
    Next i
    Set oPersonas = Nothing
    Exit Function
    
ErrorVeirficandoDuplicidadDoc:
    Set oPersonas = Nothing
    Err.Raise Err.Number, "BuscaDocumento", Err.Description
End Function

Public Function VerificandoTamanioDoc(ByVal popersona As DPersona) As Integer
Dim i As Integer

    On Error GoTo ErrorVerificandoTamanioDoc
    VerificandoTamanioDoc = -1
    For i = 0 To popersona.NumeroDocumentos - 1
         If popersona.ObtenerDocTipoAct(i) <> PersFilaSinCambios And popersona.ObtenerDocTipoAct(i) <> PersFilaEliminda Then
            If (Len(popersona.ObtenerNumeroDoc(i)) <> 8 And CInt(Trim(Right(popersona.ObtenerTipoDoc(i), 15))) = gPersIdDNI) Or (Len(popersona.ObtenerNumeroDoc(i)) <> 11 And CInt(Trim(Right(popersona.ObtenerTipoDoc(i), 15))) = gPersIdRUC) Then
               VerificandoTamanioDoc = i
               Exit Function
            End If
        End If
    Next i
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

'*** PEAC 20100406 - 20090718
'-- esta función está en NCOMPERSONA para eliminarlo

Public Function ImprimePersListaNegativa(ByVal pcCondicion As String, ByVal pcMovNro As String) As String
Dim vcondC As String
Dim rsListaNeg As ADODB.Recordset
Dim rsJuztificacion As ADODB.Recordset
Dim lsDelito As String
Dim oDPers As New clases.DPersonas
Dim oPers As New COMDpersona.DCOMPersonas

Dim sCadImp As String, lcTitulo As String, lcNombre As String
Dim i As Integer, J As Integer, Pos As Long 'madm 20091112
Dim nItem As Integer 'JACA 20110217
Dim nChangeSituacion As Integer

    Screen.MousePointer = 11
       
    Set oDPers = New clases.DPersonas
    Set rsListaNeg = oDPers.ObtenerPersListaNegativa(pcCondicion, pcMovNro)
    Set oDPers = Nothing

    If (rsListaNeg.EOF And rsListaNeg.BOF) Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Function
    End If
    
    Pos = 0 'madm 20091112
    nItem = 0
    lcNombre = ""
    sCadImp = ""
    sCadImp = sCadImp & "CAJA MAYNAS" & Chr(10)  ''lnSaltoLinDoc
    sCadImp = sCadImp & "RUC : 20103845328" & Chr(10) & Chr(10) ''lnSaltoLinDoc & lnSaltoLinDoc
    
    lcTitulo = "REGISTRO DE CLIENTES - LISTA NEGATIVA"
    sCadImp = sCadImp & Centra(lcTitulo, 130) & Chr(10)
    sCadImp = sCadImp & String(130, "-") & Chr(10) & Chr(10)   'lnSaltoLinDoc & lnSaltoLinDoc
        
'    sCadImp = sCadImp & "ITEM  TPO PERSONA   DOC.    NUMERO DOC.       NOMBRES/RAZON SOCIAL             JUSTIFICACION               FUENTE" & Chr(10)  'lnSaltoLinDoc
'    sCadImp = sCadImp & String(130, "-") & Chr(10) ' lnSaltoLinDoc
    
    vcondC = "" 'madm 20100308
    
    nChangeSituacion = 0
    
    Do While Not rsListaNeg.EOF
        
        If rsListaNeg!cCondicion <> vcondC Then 'madm 20100308
            nChangeSituacion = 1
        End If
                            
        If nChangeSituacion = 1 Then 'madm 20100308
            sCadImp = sCadImp & Chr(10) & Chr(10)
            sCadImp = sCadImp & "CONDICION  : " & rsListaNeg!cCondicion & Chr(10)
            sCadImp = sCadImp & String(130, "-") & Chr(10)
            
            sCadImp = sCadImp & "ITEM  TPO PERSONA   DOC.    NUMERO DOC.       NOMBRES/RAZON SOCIAL             JUSTIFICACION               FUENTE" & Chr(10)  'lnSaltoLinDoc
            sCadImp = sCadImp & String(130, "-") & Chr(10) ' lnSaltoLinDoc
            Pos = 0
            nChangeSituacion = 0
        End If
            
        Pos = Pos + 1 'madm 20091112
        nItem = nItem + 1 'JACA 20110217
        sCadImp = sCadImp & ImpreFormat(nItem, 5, 0) 'madm 20091112
        sCadImp = sCadImp & ImpreFormat(rsListaNeg!cTipoPersona, 12)
        sCadImp = sCadImp & ImpreFormat(rsListaNeg!cTipoDoc, 5, 0)
        sCadImp = sCadImp & ImpreFormat(Trim(rsListaNeg!cNumId), 15)
        sCadImp = sCadImp & ImpreFormat(Trim(rsListaNeg!cNombre), 30)
        'modificado para 1 o + justificaciones
'        sCadImp = sCadImp & ImpreFormat(0, 25)
'        sCadImp = sCadImp & ImpreFormat(0, 28)
        
        'JACA 20110217
        Set rsJuztificacion = oPers.CargaDatosPersNegativasJustificacion_MovNro(rsListaNeg!cMovNro)
         
        If Not (rsJuztificacion.EOF And rsJuztificacion.BOF) Then
                Do While Not rsJuztificacion.EOF
                    
                    If lcNombre = rsListaNeg!cNombre Then
                        sCadImp = sCadImp & String(73, " ")
                    End If
                    
                    sCadImp = sCadImp & ImpreFormat(rsJuztificacion!cDelito, 25)
                    sCadImp = sCadImp & ImpreFormat(rsJuztificacion!cOfMultiple, 28)
                    sCadImp = sCadImp & Chr(10)
                    Pos = Pos + 1
                    lcNombre = rsListaNeg!cNombre
                    rsJuztificacion.MoveNext
                Loop
        End If
        'JACA END
        sCadImp = sCadImp & Chr(10)
        
        vcondC = rsListaNeg!cCondicion 'madm 20100308
        rsListaNeg.MoveNext
    Loop
        
    ImprimePersListaNegativa = sCadImp
    Screen.MousePointer = 0

    rsListaNeg.Close
    Set rsListaNeg = Nothing
    rsJuztificacion.Close
    Set rsJuztificacion = Nothing


End Function

