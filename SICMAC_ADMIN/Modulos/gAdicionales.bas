Attribute VB_Name = "gAdicionales"
Option Explicit

Public sRutaImg As String
 
Global sCodIF As String
Global gsLpt As String

'PARA NEGOCIO
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Const CB_FINDSTRING = &H14C
Global Const gsConnServDBF = "DSN=DSNCmactServ"
Global gcCentralPers As String

Public Enum CostosControlOpe
    OpeCostosControlOpeNivUsu = 999901
    OpeCostosControlOpeNivArea = 999902
    OpeCostosControlOpeNivAge = 999903
    OpeCostosControlOpeNivCmac = 999904
End Enum

Public Enum CostosControlNiv
    OpeCostosControlNivUsu = 1
    OpeCostosControlNivArea = 2
    OpeCostosControlNivAge = 3
    OpeCostosControlNivCmac = 4
End Enum

Public Sub LetPicture(f As Variant, ByVal Pic As Image)
    Dim X() As Byte
    Dim N   As Long
    Dim ff  As Integer
    SavePicture Pic.Picture, "pic"
    ff = FreeFile
    Open "pic" For Binary Access Read As ff
    N = LOF(ff)
    If N Then
        ReDim X(N)
        Get ff, , X()
        f.AppendChunk X()
        Close ff
    End If
    Kill "pic"
End Sub

Public Sub GetPicture(f As ADODB.Field, Pic As Image)
'    Dim X() As Byte
'    Dim ff  As Integer
'    ff = FreeFile
'    Open "pic" For Binary Access Write As ff
'    X() = f.GetChunk(f.ActualSize)
'    Put ff, , X()
'    Close ff
'    Pic = LoadPicture("pic")
'    Kill "pic"
End Sub

Private Function GrabarFirmaenBD(ByRef RF As ADODB.Recordset, ByVal psPersCod As String, ByVal psUltimaAct As String) As Boolean
    Dim sSQL As String
    Dim RStream As ADODB.Stream

    On Error GoTo ErrorGrabarFirmaenBD
    If RF.BOF And RF.EOF Then
        RF.AddNew
        RF.Fields("cPersCod").value = psPersCod
        RF.Fields("cUltimaActualizacion").value = psUltimaAct
    End If
    
    Set RStream = New ADODB.Stream
    RStream.Type = adTypeBinary
    RStream.Open
    RStream.LoadFromFile sRutaImg
    RF.Fields("iPersFirma").value = RStream.Read
    RStream.Close
    Set RStream = Nothing
    GrabarFirmaenBD = True
    Exit Function
    
ErrorGrabarFirmaenBD:
    GrabarFirmaenBD = False
    
End Function

Public Function GetMeses(ByVal pdFIni As Date, ByVal pdFFin As Date) As Single
    Dim nMes As Single
    Dim dAux As Date
    'Calcular valor en meses
    nMes = 0
    dAux = CDate("01" & "/" & Month(pdFIni) & "/" & Year(pdFIni))
    dAux = DateAdd("m", 1, dAux)
    If Day(pdFIni) > 1 Then
        If dAux < pdFFin Then
            nMes = nMes + DateDiff("d", pdFIni, dAux) / 30
        Else
            nMes = nMes + (DateDiff("d", pdFIni, pdFFin) + 1) / 30
            GetMeses = nMes
            Exit Function
        End If
    Else
        nMes = nMes + 1
    End If
    dAux = DateAdd("m", 1, dAux)
    While DateDiff("d", -1, dAux) < pdFFin
        nMes = nMes + 1
        dAux = DateAdd("m", 1, dAux)
    Wend
    If DateAdd("d", -1, dAux) = pdFFin Then
        nMes = nMes + 1
    Else
        dAux = DateAdd("m", -1, dAux)
        nMes = nMes + (DateDiff("d", dAux, pdFFin) + 1) / 30
    End If
    GetMeses = nMes
End Function

'***************************************************
'* VERIFICA QUE SEA CORRECTO EL NRO DE CUENTA
'* VALIDANDO EL DIGITO DE CHEQUEO
'***************************************************
'FECHA CREACION : 24/06/99  -   JPNZ
'MODIFICACION:
'***************************************************
Public Function EsValido(ByVal psCadena As String) As Boolean
Dim liDigito As Integer
'Validar una Cadena con Módulo 11
liDigito = Val(Mid(psCadena, 6, 1))
If GetDigitoChequeo(psCadena) = liDigito Then
    EsValido = True
Else
    EsValido = False
End If
End Function


'* UTILIZA EL MODULO 11
'***************************************************
'FECHA CREACION : 24/06/99  -   JPNZ
'MODIFICACION:
'***************************************************
Public Function GetDigitoChequeo(ByVal psCadena As String) As Integer
Dim liFactor As Variant
Dim liCadena(1 To 5) As Integer
Dim liSum, i As Integer
Dim lnDigito As Integer
liFactor = Array(6, 5, 4, 3, 2)
liCadena(1) = Val(Mid(psCadena, 1, 1))
liCadena(2) = Val(Mid(psCadena, 2, 1))
liCadena(3) = Val(Mid(psCadena, 3, 1))
liCadena(4) = Val(Mid(psCadena, 4, 1))
liCadena(5) = Val(Mid(psCadena, 5, 1))
liSum = 0
For i = 1 To 5
    liSum = liSum + liCadena(i) * liFactor(i - 1)
Next i
lnDigito = 11 - (liSum Mod 11)
If lnDigito = 10 Then
    GetDigitoChequeo = 0
ElseIf lnDigito = 11 Then
    GetDigitoChequeo = 1
Else
    GetDigitoChequeo = lnDigito
End If

End Function

'Public Sub LetPicture1(f As Variant, ByVal Pic As Image)
'    Dim X() As Byte
'    Dim N   As Long
'    Dim ff  As Integer
'    Dim rValue&
'
'    ChDrive App.Path
'    ChDir App.Path
'
'    SavePicture Pic.Picture, App.Path & "\pic"
'
'    'Create Zip
'    rValue& = zCreateNewZip(App.Path & "\zip")
'    If rValue& <> 0 Then
'        MsgBox "Unable to create Zip", vbCritical, "Error!"
'        Exit Sub
'    End If
'
'    'Ask the VbZip Library to be prepared to process the File
'    rValue& = zOrderFile(App.Path & "\pic", "pic", 0&)
'    If rValue& <> 0 Then
'      MsgBox "Unable to Add " & App.Path & "\pic", vbCritical, "Error!"
'    End If
'
'    rValue& = zCompressFiles(App.Path, "", 2, 0, 0, "", 0)
'    If rValue& <> 0 Then
'      MsgBox "Unable to Execute Compress", vbCritical, "Error!"
'    End If
'
'    'Don't forget to close the Archive.
'    rValue& = zCloseZipFile
'    If rValue& <> 0 Then
'      MsgBox "Unable to close Zip", vbCritical, "Error!"
'      Exit Sub
'    End If
'
'    ff = FreeFile
'    Open App.Path & "\zip" For Binary Access Read As ff
'    N = LOF(ff)
'    If N Then
'        ReDim X(N)
'        Get ff, , X()
'        f.AppendChunk X()
'        Close ff
'    End If
'    Kill App.Path & "\zip"
'    Kill App.Path & "\pic"
'End Sub
'
'Public Sub GetPicture1(f As ADODB.Field, Pic As Image)
'    Dim X() As Byte
'    Dim ff  As Integer
'    ff = FreeFile
'
'    ChDrive App.Path
'    ChDir App.Path
'
'    Open App.Path & "\zip" For Binary Access Write As ff
'    X() = f.GetChunk(f.ActualSize)
'    Put ff, , X()
'    Close ff
'
'    If zOpenZipFile(App.Path & "\zip") <> 0 Then
'      MsgBox (StringFromPointer(zGetLastErrorAsText))
'      Exit Sub
'    End If
'
'    Call zExtractAll(App.Path & "\", "", 1, 1, 1, 0, 0)
'
'    zCloseZipFile
'
'    Pic = LoadPicture(App.Path & "\pic")
'    Kill App.Path & "\pic"
'    Kill App.Path & "\zip"
'End Sub

'Public Sub LetPictureActualiza(fIni As Variant, fFin As Variant)
'    Dim X() As Byte
'    Dim N   As Long
'    Dim ffIni  As Integer
'    Dim ffDes  As Integer
'    Dim rValue&
'
'    ChDrive App.Path
'    ChDir App.Path
'
'    ffIni = FreeFile
'    Open App.Path & "\pic" For Binary Access Write As ffIni
'    X() = fIni.GetChunk(fIni.ActualSize)
'    Put ffIni, , X()
'    Close ffIni
'
'    'Create Zip
'    rValue& = zCreateNewZip(App.Path & "\zip")
'    If rValue& <> 0 Then
'        MsgBox "Unable to create Zip", vbCritical, "Error!"
'        Exit Sub
'    End If
'
'    'Ask the VbZip Library to be prepared to process the File
'    rValue& = zOrderFile(App.Path & "\pic", "pic", 0&)
'    If rValue& <> 0 Then
'      MsgBox "Unable to Add " & App.Path & "\pic", vbCritical, "Error!"
'    End If
'
'    rValue& = zCompressFiles(App.Path, "", 3, 0, 0, "", 0)
'    If rValue& <> 0 Then
'      MsgBox "Unable to Execute Compress", vbCritical, "Error!"
'    End If
'
'    'Don't forget to close the Archive.
'    rValue& = zCloseZipFile
'    If rValue& <> 0 Then
'      MsgBox "Unable to close Zip", vbCritical, "Error!"
'      Exit Sub
'    End If
'
'    ffDes = FreeFile
'    Open App.Path & "\zip" For Binary Access Read As ffDes
'    N = LOF(ffDes)
'    If N Then
'        ReDim X(N)
'        Get ffDes, , X()
'        fFin.AppendChunk X()
'        Close ffDes
'    End If
'    Kill App.Path & "\zip"
'    Kill App.Path & "\pic"
'End Sub

