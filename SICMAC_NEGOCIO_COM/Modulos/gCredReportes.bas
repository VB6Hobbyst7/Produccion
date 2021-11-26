Attribute VB_Name = "gCredReportes"
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim objProducto As COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018

'*** PEAC 20080412
Public Function ImprimeInformeCriteriosAceptacionRiesgo(ByVal psCtaCod As String, _
ByVal psNomAge As String, ByVal psCodUsu As String, ByVal pR As ADODB.Recordset, ByVal pRSbs As ADODB.Recordset, ByVal pRDatFin As ADODB.Recordset, ByVal pRCap As ADODB.Recordset) As String

    Dim nTasaAnual As Double
    Dim nTasaMoraA As Double

    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range

    Dim RAge As ADODB.Recordset
    Dim sCompaSeguro As String
    Dim sNumeroPoliza As String
    Dim nParametro As Double
    Dim RRelaCred  As ADODB.Recordset
    Dim RSbs As ADODB.Recordset
    Dim RDatFin As ADODB.Recordset
    Dim RCap As ADODB.Recordset

    Dim nCosRef As Double
    Dim nComDsg As Double
    Dim nCosCroPag As Double
    Dim nCosHisPag As Double
    Dim nCosComPag As Double
    Dim nMontoInt As Double
    Dim nPriPol As Double
    Dim nCantAnt As Double
    Dim nCenCrg As Double
    Dim nSgrDsg1 As Double
    Dim nSgrDsg2 As Double
    Dim nSegMor As Double
    Dim nMinCuo As Double
    Dim nMaxCuo As Double
    Dim lcNonCal As String
    Dim i As Integer, lcCal As String
    Dim sArchivo As String

    Dim FecIni As Date, FecFin As Date
    Dim Meses As Integer, Anhos As Integer, Dias As Integer

    On Error GoTo ErrorCargaCAR

    If pRSbs.RecordCount > 0 Then
    lcCal = ""

    For i = 0 To 4
        If i = 0 Then
            lcNonCal = "NOR"
        ElseIf i = 1 Then
            lcNonCal = "POT"
        ElseIf i = 2 Then
            lcNonCal = "DEF"
        ElseIf i = 3 Then
            lcNonCal = "DUD"
        ElseIf i = 4 Then
            lcNonCal = "PER"
        End If
        If pRSbs.Fields.Item(i) > 0 Then
            lcCal = lcCal & lcNonCal & Str(pRSbs.Fields.Item(i)) & " "
        End If
    Next
    End If

Anhos = 0
Meses = 0

'--- PEAC 20080530 calcula Experiencia en sector en años meses y dias
If pRCap.RecordCount > 0 Then
    If Format(pRCap!dPersIniActi, "dd/mm/yyyy") > "01/01/1950" Then
        FecIni = Format(pRCap!dPersIniActi, "dd/mm/yyyy")
        FecFin = Format(Date, "dd/mm/yyyy")
        '** años
        Anhos = Year(Format(FecFin, "yyyy/mm/dd")) - Year(Format(FecIni, "yyyy/mm/dd"))
        '** meses
        If Month(FecFin) = Month(FecIni) Then
           Meses = 0
        ElseIf Month(FecFin) > Month(FecIni) Then
            Meses = Month(FecFin) - Month(FecIni)
        ElseIf Month(FecFin) < Month(FecIni) Then
            Meses = 12 + (Month(FecFin) - Month(FecIni))
            Anhos = Anhos - 1
        End If
        '** dias
        If Day(FecFin) >= Day(FecIni) Then
            Dias = Day(FecFin) - Day(FecIni)
        ElseIf Day(FecFin) < Day(FecIni) Then
            Meses = Meses - 1
            If Meses < 0 Then
                Anhos = Anhos - 1
                Meses = Meses + 12
            End If
            If Month(FecIni) = 1 Or Month(FecIni) = 3 Or Month(FecIni) = 5 Or Month(FecIni) = 7 Or Month(FecIni) = 8 Or Month(FecIni) = 10 Or Month(FecIni) = 12 Then
            Dias = Day(FecFin) - Day(FecIni) + 31
            ElseIf Month(FecIni) = 2 Then
                Dias = Day(FecFin) - Day(FecIni) + 28
                If Year(FecIni) Mod 4 = 0 Then Dias = Dias + 1
            Else
                Dias = Day(FecFin) - Day(FecIni) + 30
            End If
        End If
    End If
End If
'--- fin calcula Experiencia en sector en años meses y dias

    Set oWord = CreateObject("Word.Application")
        oWord.Visible = False

    Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\CriteriosAceptacionRiesgo.doc")

    sArchivo = App.Path & "\FormatoCarta\CAR_" & Replace(Left(Time, 5), ":", "") & ".doc"
    oDoc.SaveAs (sArchivo)

'*** Datos de Cabecera

    With oWord.Selection.Find
        .Text = "<<cOficina>>"
        '.Replacement.Text = psCtaCod
        .Replacement.Text = psNomAge
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cNomCliente>>"
        '.Replacement.Text = R!cTipoCredDescrip
        .Replacement.Text = IIf(Len(pR!nombre_deudor) = 0, ".....", pR!nombre_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cCuenta>>"
        '.Replacement.Text = R!cTipoCredDescrip
        .Replacement.Text = psCtaCod
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cCodCliente>>"
        '.Replacement.Text = R!cTipoCredDescrip
        .Replacement.Text = IIf(Len(pR!CodCliente) = 0, ".....", pR!CodCliente)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cDni>>"
        '.Replacement.Text = R!cTipoCredDescrip
        .Replacement.Text = IIf(Len(pR!dni_deudor) = 0, ".....", pR!dni_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cRuc>>"
        '.Replacement.Text = R!cTipoCredDescrip
        .Replacement.Text = IIf(Len(pR!ruc_deudor) = 0, ".....", pR!ruc_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cTipCred>>"
        '.Replacement.Text = R!cTipoCredDescrip
        .Replacement.Text = IIf(Len(pR!TipoCred) = 0, ".....", pR!TipoCred)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cUsu>>"
        '.Replacement.Text = R!cTipoCredDescrip
        .Replacement.Text = psCodUsu
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cAna>>"
        '.Replacement.Text = R!cTipoCredDescrip
        .Replacement.Text = IIf(Len(pR!Analista) = 0, ".....", pR!Analista)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

'*** Datos del contenido
    With oWord.Selection.Find
        .Text = "<<nCapa>>"
        If pRDatFin!Excedente = 0 Then
            .Replacement.Text = Str(Format(0, "#0.00")) & " %"
        Else
            .Replacement.Text = Str(Format((pRDatFin!ValorCuota / pRDatFin!Excedente) * 100, "#0.00")) & " %"
        End If
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<nCalificaSbs>>"
        .Replacement.Text = lcCal
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<nExperiencia>>"
        .Replacement.Text = Str(Anhos) & " Años, " & Str(Meses) & " Meses"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

'*** fin

    oDoc.Close
    Set oDoc = Nothing

    Set oWord = CreateObject("Word.Application")
    oWord.Visible = True
    Set oDoc = oWord.Documents.Open(sArchivo)
    Set oDoc = Nothing
    Set oWord = Nothing

    Exit Function
'PEAC 20090520
ErrorCargaCAR:
    MsgBox err.Description, vbCritical, "Aviso"
    Set oDoc = Nothing
    Set oWord = Nothing

End Function

'*** PEAC 20080412
Public Function ImprimeInformeVisitaCliente(ByVal psCtaCod As String, _
ByVal psNomAge As String, ByVal psCodUsu As String, ByVal pR As ADODB.Recordset) As String

    Dim R As ADODB.Recordset
    Dim nTasaAnual As Double
    Dim nTasaMoraA As Double

    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range

    Dim RAge As ADODB.Recordset
    Dim sCompaSeguro As String
    Dim sNumeroPoliza As String
    Dim nParametro As Double
    Dim RRelaCred  As ADODB.Recordset
    Dim oCred As COMNCredito.NCOMCredDoc

    Dim nCosRef As Double
    Dim nComDsg As Double
    Dim nCosCroPag As Double
    Dim nCosHisPag As Double
    Dim nCosComPag As Double
    Dim nMontoInt As Double
    Dim nPriPol As Double
    Dim nCantAnt As Double
    Dim nCenCrg As Double
    Dim nSgrDsg1 As Double
    Dim nSgrDsg2 As Double
    Dim nSegMor As Double
    Dim nMinCuo As Double
    Dim nMaxCuo As Double

    Dim sArchivo As String

    On Error GoTo ErrorVisitaCli

    Set oWord = CreateObject("Word.Application")
        oWord.Visible = False

    Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\InformeVisitaCliente.doc")

    sArchivo = App.Path & "\FormatoCarta\IVC_" & Replace(Left(Time, 5), ":", "") & ".doc"
    oDoc.SaveAs (sArchivo)

'*** Datos de Cabecera

    With oWord.Selection.Find
        .Text = "<<cOficina>>"
        .Replacement.Text = psNomAge
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cNomCliente>>"
        .Replacement.Text = IIf(Len(pR!nombre_deudor) = 0, ".....", pR!nombre_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cCuenta>>"
        .Replacement.Text = psCtaCod
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cCodCliente>>"
        .Replacement.Text = IIf(Len(pR!CodCliente) = 0, ".....", pR!CodCliente)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cDni>>"
        .Replacement.Text = IIf(Len(pR!dni_deudor) = 0, ".....", pR!dni_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cRuc>>"
        .Replacement.Text = IIf(Len(pR!ruc_deudor) = 0, ".....", pR!ruc_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cTipCred>>"
        .Replacement.Text = IIf(Len(pR!TipoCred) = 0, ".....", pR!TipoCred)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cUsu>>"
        .Replacement.Text = psCodUsu
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cAna>>"
        .Replacement.Text = IIf(Len(pR!Analista) = 0, ".....", pR!Analista)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

'*** fin

    oDoc.Close
    Set oDoc = Nothing

    Set oWord = CreateObject("Word.Application")
    oWord.Visible = True
    Set oDoc = oWord.Documents.Open(sArchivo)
    Set oDoc = Nothing
    Set oWord = Nothing

    Exit Function
'PEAC 20090520
ErrorVisitaCli:
    MsgBox err.Description, vbCritical, "Aviso"
    Set oDoc = Nothing
    Set oWord = Nothing


End Function


'***PEAC 20080402
Public Function ImprimeInformeComercial02(ByVal psCtaCod As String, _
ByVal psNomAge As String, ByVal psCodUsu As String, ByVal pR As ADODB.Recordset, _
ByVal pRB As ADODB.Recordset, ByVal pRDatFin As ADODB.Recordset, _
ByVal pRDatAdc As ADODB.Recordset, ByVal pRDatGarant As ADODB.Recordset) As String

    Dim R As ADODB.Recordset
    Dim RB As ADODB.Recordset
    Dim nTasaAnual As Double
    Dim nTasaMoraA As Double

    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range

    Dim RAge As ADODB.Recordset
    Dim sCompaSeguro As String
    Dim sNumeroPoliza As String
    Dim nParametro As Double
    Dim RRelaCred  As ADODB.Recordset
    Dim RDatFin As ADODB.Recordset
    Dim oCred As COMNCredito.NCOMCredDoc
    Dim oCredB As COMNCredito.NCOMCredDoc
    Dim oDCredC As COMDCredito.DCOMCredito

    Dim nCosRef As Double
    Dim nComDsg As Double
    Dim nCosCroPag As Double
    Dim nCosHisPag As Double
    Dim nCosComPag As Double
    Dim nMontoInt As Double
    Dim nPriPol As Double
    Dim nCantAnt As Double
    Dim nCenCrg As Double
    Dim nSgrDsg1 As Double
    Dim nSgrDsg2 As Double
    Dim nSegMor As Double
    Dim nMinCuo As Double
    Dim nMaxCuo As Double

    Dim sArchivo As String

    On Error GoTo ErrorInformeCom02

    Set oWord = CreateObject("Word.Application")
        oWord.Visible = False

    Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\InformeComercial.doc")

    sArchivo = App.Path & "\FormatoCarta\IC_" & psCtaCod & "_" & Replace(Left(Time, 5), ":", "") & ".doc"
    oDoc.SaveAs (sArchivo)

'*** Datos de Cabecera

    With oWord.Selection.Find
        .Text = "<<cOficina>>"
        .Replacement.Text = psNomAge
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cNomCliente>>"
        .Replacement.Text = IIf(Len(pR!nombre_deudor) = 0, ".......", pR!nombre_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cCuenta>>"
        .Replacement.Text = psCtaCod
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cCodCliente>>"
        .Replacement.Text = IIf(Len(pR!CodCliente) = 0, ".......", pR!CodCliente)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cDni>>"
        .Replacement.Text = IIf(Len(pR!dni_deudor) = 0, ".......", pR!dni_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cRuc>>"
        .Replacement.Text = IIf(Len(pR!ruc_deudor) = 0, ".......", pR!ruc_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cTipCred>>"
        .Replacement.Text = IIf(Len(pR!TipoCred) = 0, ".......", pR!TipoCred)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cUsu>>"
        .Replacement.Text = psCodUsu
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cAna>>"
        .Replacement.Text = IIf(Len(pR!Analista) = 0, ".......", pR!Analista)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

'*** Datos del Deudor

    With oWord.Selection.Find
        .Text = "<<cNombreDeudor>>"
        .Replacement.Text = pR!nombre_deudor
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cDomicilio>>"
        .Replacement.Text = IIf(Len(pR!dire_deudor) = 0, ".......", pR!dire_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cDni>>"
        .Replacement.Text = IIf(Len(pR!dni_deudor) = 0, ".......", pR!dni_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cCodSbs>>"
        .Replacement.Text = IIf(Len(pR!codsbs_deudor) = 0, ".......", pR!codsbs_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

'*****CTI3 (ferimoro) 19092018**********************************
Dim nSumArt As Double
Dim nSumProp As Double
Dim Patri As Double
nSumArt = 0
nSumProp = 0

Do While Not pRDatGarant.EOF

If pRDatGarant!nTpoGarantia = 4 Then
  nSumArt = nSumArt + pRDatGarant!TotalDJ
Else
  nSumProp = nSumProp + pRDatGarant!nTasacion
End If

    pRDatGarant.MoveNext
Loop

Patri = nSumArt + nSumProp
'*************************************************************

    With oWord.Selection.Find
        .Text = "<<nPatrimonio>>"
        .Replacement.Text = IIf(Patri > 0, Format(Patri, "#,##0.00"), ".......")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

'    With oWord.Selection.Find
'        .Text = "<<nPatrimonio>>"
'        .Replacement.Text = IIf(pRDatFin.RecordCount > 0, Format(pRDatFin!PatrimonioPerso, "#,##0.00"), ".......")
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With

    With oWord.Selection.Find
        .Text = "<<cEstadoCivil>>"
        .Replacement.Text = IIf(Len(pR!estadocivil_deudor) = 0, ".......", pR!estadocivil_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cCentroLaboral>>"
        .Replacement.Text = IIf(Len(pR!centrolaboral_deudor) = 0, ".......", pR!centrolaboral_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'fin peac

    With oWord.Selection.Find
        .Text = "<<cGiroEmpresaLabora>>"
        .Replacement.Text = "....."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<nAntiguedad>>"
        .Replacement.Text = "....."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

'*** Conyuge

    With oWord.Selection.Find
        .Text = "<<cNombreConyuge>>"
        .Replacement.Text = IIf(Len(pR!nombre_conyuge) = 0, ".......", pR!nombre_conyuge)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cDniConyuge>>"
        .Replacement.Text = IIf(Len(pR!dni_conyuge) = 0, ".......", pR!dni_conyuge)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cCodSbsConyuge>>"
        .Replacement.Text = IIf(Len(pR!codsbs_conyuge) = 0, ".......", pR!codsbs_conyuge)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cCentroLaboralConyuge>>"
        .Replacement.Text = IIf(Len(pR!CentroLaboral_Conyuge) = 0, ".......", pR!CentroLaboral_Conyuge)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

'**************CTI3

If pRDatAdc.RecordCount > 0 Then

Dim i As Integer
Dim Tabla As Word.Table
Dim filas, columnas As Integer

filas = pRDatAdc.RecordCount + 2
columnas = 6

oWord.Selection.GoTo WHAT:=wdGoToBookmark, Name:="Cuadros"  'se coloca en la primera fila al buscar el marcador cuadro
oWord.Selection.Font.Size = 8
oWord.Selection.Font.Bold = False

Set Tabla = oDoc.Tables.Add(oWord.Selection.Range, filas, columnas, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
wdAutoFitFixed)

'CABECERA DE LA TABLA
With Tabla
    .Borders.OutsideLineStyle = wdLineStyleSingle
    .Borders.InsideLineStyle = wdLineStyleSingle
    
    .Columns(1).Width = 150
    .Columns(2).Width = 140
    .Columns(3).Width = 50
    .Columns(4).Width = 60
    .Columns(5).Width = 25
    
    .Columns(6).Width = 55
    
    'combina celdas
    .Cell(1, 1).Merge .Cell(2, 1) 'COMBINA FILA 1 -2 , COLUMNA 1
    .Cell(1, 2).Merge .Cell(1, 6)
    
    .Cell(1, 1).Range.Text = "NOMBRES Y APELLIDOS"
    .Cell(1, 1).Range.Font.Bold = True
    .Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 2).Range.Text = "ENTIDADES EN LAS QUE TIENEN PARTICIPACION PATRIMONIAL"
    .Cell(1, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 2).Range.Font.Bold = True
    .Cell(2, 2).Range.Text = "RAZON SOCIAL"
    .Cell(2, 2).Range.Font.Bold = True
    .Cell(2, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    '.Cell(1, 2).Range.Font.Color = wdColorBlue 'cambia coolor de letra
    .Cell(2, 3).Range.Text = "RUC"
    .Cell(2, 3).Range.Font.Bold = True
    .Cell(2, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(2, 4).Range.Text = "Aportes S/."
    .Cell(2, 4).Range.Font.Bold = True
    .Cell(2, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(2, 5).Range.Text = "%"
    .Cell(2, 5).Range.Font.Bold = True
    .Cell(2, 5).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(2, 6).Range.Text = "Cod. SBS(1)"
    .Cell(2, 6).Range.Font.Bold = True
    .Cell(2, 6).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
    
    i = 2
    Do While Not pRDatAdc.EOF
    
        i = i + 1
        
            .Cell(i, 1).Range.Text = pRDatAdc!cNombreJur
            .Cell(i, 2).Range.Text = pRDatAdc!cRazonSocial
            .Cell(i, 3).Range.Text = pRDatAdc!vRuc
            .Cell(i, 4).Range.Text = Format(pRDatAdc!nApoSol, "#,#0.00")
            .Cell(i, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
            .Cell(i, 5).Range.Text = pRDatAdc!nApoPorc
            .Cell(i, 6).Range.Text = pRDatAdc!vCodSbs
    
        pRDatAdc.MoveNext
    Loop
      
End With

Else
filas = 3
columnas = 6

oWord.Selection.GoTo WHAT:=wdGoToBookmark, Name:="Cuadros"  'se coloca en la primera fila al buscar el marcador cuadro
oWord.Selection.Font.Size = 8
oWord.Selection.Font.Bold = False

Set Tabla = oDoc.Tables.Add(oWord.Selection.Range, filas, columnas, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
wdAutoFitFixed)

'CABECERA DE LA TABLA
With Tabla
    .Borders.OutsideLineStyle = wdLineStyleSingle
    .Borders.InsideLineStyle = wdLineStyleSingle
    
    .Columns(1).Width = 150
    .Columns(2).Width = 140
    .Columns(3).Width = 50
    .Columns(4).Width = 60
    .Columns(5).Width = 25
    
    .Columns(6).Width = 55
    
    'combina celdas
    .Cell(1, 1).Merge .Cell(2, 1) 'COMBINA FILA 1 -2 , COLUMNA 1
    .Cell(1, 2).Merge .Cell(1, 6)
    
    .Cell(1, 1).Range.Text = "NOMBRES Y APELLIDOS"
    .Cell(1, 1).Range.Font.Bold = True
    .Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 2).Range.Text = "ENTIDADES EN LAS QUE TIENEN PARTICIPACION PATRIMONIAL"
    .Cell(1, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 2).Range.Font.Bold = True
    .Cell(2, 2).Range.Text = "RAZON SOCIAL"
    .Cell(2, 2).Range.Font.Bold = True
    .Cell(2, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    '.Cell(1, 2).Range.Font.Color = wdColorBlue 'cambia coolor de letra
    .Cell(2, 3).Range.Text = "RUC"
    .Cell(2, 3).Range.Font.Bold = True
    .Cell(2, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(2, 4).Range.Text = "Aportes S/."
    .Cell(2, 4).Range.Font.Bold = True
    .Cell(2, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(2, 5).Range.Text = "%"
    .Cell(2, 5).Range.Font.Bold = True
    .Cell(2, 5).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(2, 6).Range.Text = "Cod. SBS(1)"
    .Cell(2, 6).Range.Font.Bold = True
    .Cell(2, 6).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
    
    i = 2
    Do While Not pRDatAdc.EOF
    
        i = i + 1
        
            .Cell(i, 1).Range.Text = "..................."
            .Cell(i, 2).Range.Text = "..................."
            .Cell(i, 3).Range.Text = "..................."
            .Cell(i, 4).Range.Text = "..................."
            .Cell(i, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
            .Cell(i, 5).Range.Text = "..................."
            .Cell(i, 6).Range.Text = "..................."
    
        pRDatAdc.MoveNext
    Loop
      
End With
    
End If

'******************

'*** Datos Familiares

    With oWord.Selection.Find
        .Text = "<<cNumDependientes>>"
        .Replacement.Text = pR!cNumDependientes
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cEdades>>"
        .Replacement.Text = "....."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

'*** Datos del Reporte

    With oWord.Selection.Find
        .Text = "<<cFecha>>"
        .Replacement.Text = Format(gdFecSis, "dddd") & ", " & Day(gdFecSis) & " de " & Format(gdFecSis, "mmmm") & " de " & Year(gdFecSis)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

'*** Fin

    oDoc.Close
    Set oDoc = Nothing

    Set oWord = CreateObject("Word.Application")
    oWord.Visible = True
    Set oDoc = oWord.Documents.Open(sArchivo)
    Set oDoc = Nothing
    Set oWord = Nothing

    Exit Function
'PEAC 20090520
ErrorInformeCom02:
    MsgBox err.Description, vbCritical, "Aviso"
    Set oDoc = Nothing
    Set oWord = Nothing


End Function

'***PEAC 20080402
Public Function ImprimeInformeComercial01(ByVal psCtaCod As String, _
ByVal psNomAge As String, ByVal psCodUsu As String, ByVal pR As ADODB.Recordset, _
ByVal pRB As ADODB.Recordset, ByVal pRDatFin As ADODB.Recordset, _
ByVal pRDatA As ADODB.Recordset, ByVal pRDatD As ADODB.Recordset, _
ByVal pRDatG As ADODB.Recordset, ByVal pRDatP As ADODB.Recordset, _
ByVal pRDatCa As ADODB.Recordset) As String

    Dim R As ADODB.Recordset
    Dim RB As ADODB.Recordset
    Dim nTasaAnual As Double
    Dim nTasaMoraA As Double

    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range

    Dim RAge As ADODB.Recordset
    Dim sCompaSeguro As String
    Dim sNumeroPoliza As String
    Dim nParametro As Double
    Dim RRelaCred  As ADODB.Recordset
    Dim RDatFin As ADODB.Recordset
    Dim oCred As COMNCredito.NCOMCredDoc
    Dim oCredB As COMNCredito.NCOMCredDoc
    Dim oDCred As COMDCredito.DCOMCredito

    Dim nCosRef As Double
    Dim nComDsg As Double
    Dim nCosCroPag As Double
    Dim nCosHisPag As Double
    Dim nCosComPag As Double
    Dim nMontoInt As Double
    Dim nPriPol As Double
    Dim nCantAnt As Double
    Dim nCenCrg As Double
    Dim nSgrDsg1 As Double
    Dim nSgrDsg2 As Double
    Dim nSegMor As Double
    Dim nMinCuo As Double
    Dim nMaxCuo As Double
    Dim nCont  As Double
    Dim nCrec As Double
    Dim CostoTot As Double

    Dim sArchivo As String

    On Error GoTo ErrorInformeCom01

    Set oWord = CreateObject("Word.Application")
        oWord.Visible = False
    If Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516" Then
        Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\InformeComercial01_AF.doc")
    Else
        Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\InformeComercial01.doc")
    End If
    'sArchivo = App.Path & "\FormatoCarta\IC_" & psCtaCod & "_" & Replace(Left(Time, 5), ":", "") & ".doc"'Comento JOEP20180811
    sArchivo = App.Path & "\FormatoCarta\IC_" & psCtaCod & "_" & Replace(Left(Time, 8), ":", "") & ".doc" 'JOEP20180811
    oDoc.SaveAs (sArchivo)

'*** Datos de Cabecera

    With oWord.Selection.Find
        .Text = "<<cOficina>>"
        .Replacement.Text = psNomAge
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cNomCliente>>"
        .Replacement.Text = IIf(Len(pR!nombre_deudor) = 0, ".......", pR!nombre_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cCuenta>>"
        .Replacement.Text = psCtaCod
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cCodCliente>>"
        .Replacement.Text = IIf(Len(pR!CodCliente) = 0, ".......", pR!CodCliente)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cDni>>"
        .Replacement.Text = IIf(Len(pR!dni_deudor) = 0, ".......", pR!dni_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cRuc>>"
        .Replacement.Text = IIf(Len(pR!ruc_deudor) = 0, ".......", pR!ruc_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cTipCred>>"
        .Replacement.Text = IIf(Len(pR!TipoCred) = 0, ".......", pR!TipoCred)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cUsu>>"
        .Replacement.Text = psCodUsu
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cAna>>"
        .Replacement.Text = IIf(Len(pR!Analista) = 0, ".......", pR!Analista)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

'*** Datos del Deudor

    With oWord.Selection.Find
        .Text = "<<cNombreDeudor>>"
        .Replacement.Text = pR!nombre_deudor
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cDomicilio>>"
        .Replacement.Text = IIf(Len(pR!dire_deudor) = 0, ".......", pR!dire_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cRuc>>"
        .Replacement.Text = IIf(Len(pR!ruc_deudor) = 0, ".......", pR!ruc_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cDni>>"
        .Replacement.Text = IIf(Len(pR!dni_deudor) = 0, ".......", pR!dni_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cCodSbs>>"
        .Replacement.Text = IIf(Len(pR!codsbs_deudor) = 0, ".......", pR!codsbs_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<nCapiSocial>>"
        .Replacement.Text = "....."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    If pRDatFin.RecordCount > 0 Then
        With oWord.Selection.Find
            .Text = "<<nPatrimonio>>"
            .Replacement.Text = Format(pRDatFin!PatrimonioTotal, "#,##0.00")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
        With oWord.Selection.Find
            .Text = "<<nPatriPer>>"
            .Replacement.Text = Format(pRDatFin!PatrimonioPerso, "#,##0.00")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
    Else
        With oWord.Selection.Find
            .Text = "<<nPatrimonio>>"
            .Replacement.Text = "......."
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
        With oWord.Selection.Find
            .Text = "<<nPatriPer>>"
            .Replacement.Text = "......."
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
    End If

    With oWord.Selection.Find
        .Text = "<<cActiGiro>>"
        .Replacement.Text = IIf(Len(pR!acti_giro) = 0, ".......", pR!acti_giro)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cActividad>>"
        .Replacement.Text = IIf(Len(pR!CIIU) = 0, ".......", pR!CIIU)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cCiiu>>"
        .Replacement.Text = IIf(Len(pR!codciiu) = 0, ".......", pR!codciiu)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cPersCod>>"
        .Replacement.Text = IIf(Len(pR!CodCliente) = 0, ".......", pR!CodCliente)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<nExperiencia>>"
        .Replacement.Text = pR!nExperiencia
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<nNumPuntos>>"
        .Replacement.Text = pR!nPuntosVta
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    If Len(Trim(pR!cNumeroFuente)) > 0 Then
        If Len(Trim(pR!cNumeroFuente)) > 0 And pRB.RecordCount > 0 Then
            If (pRB!VtasContado + pRB!VtasCredito) > 0 Then
                nCont = (pRB!VtasContado / (pRB!VtasContado + pRB!VtasCredito)) * 100
                nCrec = (pRB!VtasCredito / (pRB!VtasContado + pRB!VtasCredito)) * 100
            Else
                nCont = 0
                nCrec = 0
            End If
        Else
            nCont = 0
            nCrec = 0
        End If
    Else
        nCont = 0
        nCrec = 0
    End If

    With oWord.Selection.Find
        .Text = "<<nContado>>"
        .Replacement.Text = nCont
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<nCredito>>"
        .Replacement.Text = nCrec
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

        CostoTot = pRB!CostoVtas + pRB!OtrosEgresos

    With oWord.Selection.Find
        .Text = "<<nVariable>>"
        If CostoTot > 0 Then
            .Replacement.Text = Format((pRB!CostoVtas / CostoTot) * 100, "#0.00")
        Else
            .Replacement.Text = 0
        End If
        'Format(gdFecSis, "dddd")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<nFijos>>"
        If CostoTot > 0 Then
            .Replacement.Text = Format((pRB!OtrosEgresos / CostoTot) * 100, "#0.00")
        Else
            .Replacement.Text = 0
        End If
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cSistInfor>>"
        .Replacement.Text = IIf(Len(pR!cSistInfor) = 0, ".......", pR!cSistInfor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cCompetencia>>"
        .Replacement.Text = IIf(Len(pR!TipoCompe) = 0, ".......", pR!TipoCompe)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cCadenaPro>>"
        .Replacement.Text = IIf(Len(pR!cCadenaProd) = 0, ".......", pR!cCadenaProd)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cOtrasActiviComple>>"
        .Replacement.Text = IIf(Len(pR!cActiviComple) = 0, ".......", pR!cActiviComple)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

'*** Datos del Conyuge

    With oWord.Selection.Find
        .Text = "<<cNombreConyuge>>"
        .Replacement.Text = IIf(Len(pR!nombre_conyuge) = 0, ".......", pR!nombre_conyuge)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<DniConyuge>>"
        .Replacement.Text = IIf(Len(pR!dni_conyuge) = 0, ".......", pR!dni_conyuge)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With


'*** Datos del Reporte

    With oWord.Selection.Find
        .Text = "<<cFecha>>"
        .Replacement.Text = Format(gdFecSis, "dddd") & ", " & Day(gdFecSis) & " de " & Format(gdFecSis, "mmmm") & " de " & Year(gdFecSis)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

'*** Fin

'**************CTI3 ------FERIMORO12092018
'**************CTI3 ---Accionistas

If pRDatA.RecordCount > 0 Then

Dim i As Integer
Dim Tabla As Word.Table
Dim filas, columnas As Integer

filas = pRDatA.RecordCount + 1
columnas = 6

oWord.Selection.GoTo WHAT:=wdGoToBookmark, Name:="accionistas"  'se coloca en la primera fila al buscar el marcador cuadro
oWord.Selection.Font.Size = 10
oWord.Selection.Font.Bold = False

Set Tabla = oDoc.Tables.Add(oWord.Selection.Range, filas, columnas, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
wdAutoFitFixed)

'CABECERA DE LA TABLA
With Tabla

    .Borders.OutsideLineStyle = wdLineStyleSingle
    .Borders.InsideLineStyle = wdLineStyleSingle
    
    .Columns(1).Width = 200
    .Columns(2).Width = 60
    .Columns(3).Width = 50
    .Columns(4).Width = 120
    .Columns(5).Width = 60
    .Columns(6).Width = 25
    
    .Cell(1, 1).Range.Text = "NOMBRES o RAZON SOCIAL"
    .Cell(1, 1).Range.Font.Bold = True
    .Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 2).Range.Text = "RUC"
    .Cell(1, 2).Range.Font.Bold = True
    .Cell(1, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 3).Range.Text = "D.O.I"
    .Cell(1, 3).Range.Font.Bold = True
    .Cell(1, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 4).Range.Text = "NACIONALIDAD"
    .Cell(1, 4).Range.Font.Bold = True
    .Cell(1, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 5).Range.Text = "Aportes S/."
    .Cell(1, 5).Range.Font.Bold = True
    .Cell(1, 5).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 6).Range.Text = "%"
    .Cell(1, 6).Range.Font.Bold = True
    .Cell(1, 6).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
    i = 1
    Do While Not pRDatA.EOF
    
        i = i + 1
        
            .Cell(i, 1).Range.Text = pRDatA!cNombreJur
            .Cell(i, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Cell(i, 2).Range.Text = pRDatA!vRuc
            .Cell(i, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Cell(i, 3).Range.Text = pRDatA!vCi
            .Cell(i, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Cell(i, 4).Range.Text = pRDatA!cNacionalidad
            .Cell(i, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Cell(i, 5).Range.Text = Format(pRDatA!nApoSol, "#,#0.00")
            .Cell(i, 5).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
            .Cell(i, 6).Range.Text = pRDatA!nApoPorc
            .Cell(i, 6).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
        pRDatA.MoveNext
    Loop
      
End With

Else

'Dim i As Integer
'Dim Tabla As Word.Table
'Dim filas, columnas As Integer

filas = 3
columnas = 6

oWord.Selection.GoTo WHAT:=wdGoToBookmark, Name:="accionistas"  'se coloca en la primera fila al buscar el marcador cuadro
oWord.Selection.Font.Size = 10
oWord.Selection.Font.Bold = False

Set Tabla = oDoc.Tables.Add(oWord.Selection.Range, filas, columnas, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
wdAutoFitFixed)

'CABECERA DE LA TABLA
With Tabla

    .Borders.OutsideLineStyle = wdLineStyleSingle
    .Borders.InsideLineStyle = wdLineStyleSingle
    
    .Columns(1).Width = 200
    .Columns(2).Width = 60
    .Columns(3).Width = 50
    .Columns(4).Width = 120
    .Columns(5).Width = 60
    .Columns(6).Width = 25
    
    .Cell(1, 1).Range.Text = "NOMBRES o RAZON SOCIAL"
    .Cell(1, 1).Range.Font.Bold = True
    .Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 2).Range.Text = "RUC"
    .Cell(1, 2).Range.Font.Bold = True
    .Cell(1, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 3).Range.Text = "D.O.I"
    .Cell(1, 3).Range.Font.Bold = True
    .Cell(1, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 4).Range.Text = "NACIONALIDAD"
    .Cell(1, 4).Range.Font.Bold = True
    .Cell(1, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 5).Range.Text = "Aportes S/."
    .Cell(1, 5).Range.Font.Bold = True
    .Cell(1, 5).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 6).Range.Text = "%"
    .Cell(1, 6).Range.Font.Bold = True
    .Cell(1, 6).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
    i = 1
    Do While Not pRDatA.EOF
    
        i = i + 1
        
            .Cell(i, 1).Range.Text = "..........."
            .Cell(i, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Cell(i, 2).Range.Text = "..........."
            .Cell(i, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Cell(i, 3).Range.Text = "..........."
            .Cell(i, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Cell(i, 4).Range.Text = "..........."
            .Cell(i, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Cell(i, 5).Range.Text = Format(pRDatA!nApoSol, "#,#0.00")
            .Cell(i, 5).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
            .Cell(i, 6).Range.Text = "..........."
            .Cell(i, 6).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
        pRDatA.MoveNext
    Loop
      
End With
    
       
    
    
End If

'******************
'**************CTI3 ---Directorio

If pRDatD.RecordCount > 0 Then

'Dim i As Integer
'Dim Tabla As Word.Table
'Dim filas, columnas As Integer

filas = pRDatD.RecordCount + 1
columnas = 5

oWord.Selection.GoTo WHAT:=wdGoToBookmark, Name:="directorio"  'se coloca en la primera fila al buscar el marcador cuadro
oWord.Selection.Font.Size = 10
oWord.Selection.Font.Bold = False

Set Tabla = oDoc.Tables.Add(oWord.Selection.Range, filas, columnas, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
wdAutoFitFixed)

'CABECERA DE LA TABLA
With Tabla
    .Borders.OutsideLineStyle = wdLineStyleSingle
    .Borders.InsideLineStyle = wdLineStyleSingle
    
    .Columns(1).Width = 180
    .Columns(2).Width = 60
    .Columns(3).Width = 50
    .Columns(4).Width = 120
    .Columns(5).Width = 100
    
    .Cell(1, 1).Range.Text = "NOMBRES Y APELLIDOS"
    .Cell(1, 1).Range.Font.Bold = True
    .Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 2).Range.Text = "RUC"
    .Cell(1, 2).Range.Font.Bold = True
    .Cell(1, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 3).Range.Text = "D.O.I"
    .Cell(1, 3).Range.Font.Bold = True
    .Cell(1, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 4).Range.Text = "NACIONALIDAD"
    .Cell(1, 4).Range.Font.Bold = True
    .Cell(1, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 5).Range.Text = "CARGO"
    .Cell(1, 5).Range.Font.Bold = True
    
    i = 1
    Do While Not pRDatD.EOF
    
        i = i + 1
        
            .Cell(i, 1).Range.Text = pRDatD!cNombreJur
            .Cell(i, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Cell(i, 2).Range.Text = pRDatD!vRuc
            .Cell(i, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Cell(i, 3).Range.Text = pRDatD!vCi
            .Cell(i, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Cell(i, 4).Range.Text = pRDatD!cNacionalidad
            .Cell(i, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Cell(i, 5).Range.Text = pRDatD!cCargo
            .Cell(i, 5).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
    
        pRDatD.MoveNext
    Loop
      
End With

Else

  filas = 3
columnas = 5

oWord.Selection.GoTo WHAT:=wdGoToBookmark, Name:="directorio"  'se coloca en la primera fila al buscar el marcador cuadro
oWord.Selection.Font.Size = 10
oWord.Selection.Font.Bold = False

Set Tabla = oDoc.Tables.Add(oWord.Selection.Range, filas, columnas, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
wdAutoFitFixed)

'CABECERA DE LA TABLA
With Tabla
    .Borders.OutsideLineStyle = wdLineStyleSingle
    .Borders.InsideLineStyle = wdLineStyleSingle
    
    .Columns(1).Width = 180
    .Columns(2).Width = 60
    .Columns(3).Width = 50
    .Columns(4).Width = 120
    .Columns(5).Width = 100
    
    .Cell(1, 1).Range.Text = "NOMBRES Y APELLIDOS"
    .Cell(1, 1).Range.Font.Bold = True
    .Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 2).Range.Text = "RUC"
    .Cell(1, 2).Range.Font.Bold = True
    .Cell(1, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 3).Range.Text = "D.O.I"
    .Cell(1, 3).Range.Font.Bold = True
    .Cell(1, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 4).Range.Text = "NACIONALIDAD"
    .Cell(1, 4).Range.Font.Bold = True
    .Cell(1, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 5).Range.Text = "CARGO"
    .Cell(1, 5).Range.Font.Bold = True
    
    i = 1
    Do While Not pRDatD.EOF
    
        i = i + 1
        
            .Cell(i, 1).Range.Text = "............"
            .Cell(i, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Cell(i, 2).Range.Text = "............"
            .Cell(i, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Cell(i, 3).Range.Text = "............"
            .Cell(i, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Cell(i, 4).Range.Text = "............"
            .Cell(i, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Cell(i, 5).Range.Text = "............"
            .Cell(i, 5).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
    
        pRDatD.MoveNext
    Loop
      
End With
    
End If

'******************
'**************CTI3 ---gERENCIAS

If pRDatG.RecordCount > 0 Then

'Dim i As Integer
'Dim Tabla As Word.Table
'Dim filas, columnas As Integer

filas = pRDatG.RecordCount + 1
columnas = 4

oWord.Selection.GoTo WHAT:=wdGoToBookmark, Name:="gerencia"  'se coloca en la primera fila al buscar el marcador cuadro
oWord.Selection.Font.Size = 10
oWord.Selection.Font.Bold = False

Set Tabla = oDoc.Tables.Add(oWord.Selection.Range, filas, columnas, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
wdAutoFitFixed)

'CABECERA DE LA TABLA
With Tabla
    .Borders.OutsideLineStyle = wdLineStyleSingle
    .Borders.InsideLineStyle = wdLineStyleSingle
    
    .Columns(1).Width = 180
    .Columns(2).Width = 60
    .Columns(3).Width = 50
    .Columns(4).Width = 80
    
    .Cell(1, 1).Range.Text = "NOMBRES Y APELLIDOS"
    .Cell(1, 1).Range.Font.Bold = True
    .Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 2).Range.Text = "RUC"
    .Cell(1, 2).Range.Font.Bold = True
    .Cell(1, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 3).Range.Text = "D.O.I"
    .Cell(1, 3).Range.Font.Bold = True
    .Cell(1, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 4).Range.Text = "CARGO"
    .Cell(1, 4).Range.Font.Bold = True
    .Cell(1, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
    i = 1
    Do While Not pRDatG.EOF
    
        i = i + 1
        
            .Cell(i, 1).Range.Text = pRDatG!cNombreJur
            .Cell(i, 2).Range.Text = pRDatG!vRuc
            .Cell(i, 3).Range.Text = pRDatG!vCi
            .Cell(i, 4).Range.Text = pRDatG!cCargo
    
        pRDatG.MoveNext
    Loop
      
End With

Else
filas = 3
columnas = 4

oWord.Selection.GoTo WHAT:=wdGoToBookmark, Name:="gerencia"  'se coloca en la primera fila al buscar el marcador cuadro
oWord.Selection.Font.Size = 10
oWord.Selection.Font.Bold = False

Set Tabla = oDoc.Tables.Add(oWord.Selection.Range, filas, columnas, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
wdAutoFitFixed)

'CABECERA DE LA TABLA
With Tabla
    .Borders.OutsideLineStyle = wdLineStyleSingle
    .Borders.InsideLineStyle = wdLineStyleSingle
    
    .Columns(1).Width = 180
    .Columns(2).Width = 60
    .Columns(3).Width = 50
    .Columns(4).Width = 80
    
    .Cell(1, 1).Range.Text = "NOMBRES Y APELLIDOS"
    .Cell(1, 1).Range.Font.Bold = True
    .Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 2).Range.Text = "RUC"
    .Cell(1, 2).Range.Font.Bold = True
    .Cell(1, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 3).Range.Text = "D.O.I"
    .Cell(1, 3).Range.Font.Bold = True
    .Cell(1, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 4).Range.Text = "CARGO"
    .Cell(1, 4).Range.Font.Bold = True
    .Cell(1, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
    i = 1
    Do While Not pRDatG.EOF
    
        i = i + 1
        
            .Cell(i, 1).Range.Text = "..........."
            .Cell(i, 2).Range.Text = "..........."
            .Cell(i, 3).Range.Text = "..........."
            .Cell(i, 4).Range.Text = "..........."
    
        pRDatG.MoveNext
    Loop
      
End With
    
End If

'******************
'*'**************CTI3 --Patrimonio

If pRDatP.RecordCount > 0 Then

'Dim i As Integer
'Dim Tabla As Word.Table
'Dim filas, columnas As Integer

filas = pRDatP.RecordCount + 2
columnas = 6

oWord.Selection.GoTo WHAT:=wdGoToBookmark, Name:="patrimonio"  'se coloca en la primera fila al buscar el marcador cuadro
oWord.Selection.Font.Size = 10
oWord.Selection.Font.Bold = False

Set Tabla = oDoc.Tables.Add(oWord.Selection.Range, filas, columnas, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
wdAutoFitFixed)

'CABECERA DE LA TABLA
With Tabla
    .Borders.OutsideLineStyle = wdLineStyleSingle
    .Borders.InsideLineStyle = wdLineStyleSingle
    
    .Columns(1).Width = 180
    .Columns(2).Width = 100
    .Columns(3).Width = 70
    .Columns(4).Width = 80
    .Columns(5).Width = 30
    .Columns(6).Width = 50
    
    'combina celdas
    .Cell(1, 1).Merge .Cell(2, 1) 'COMBINA FILA 1 -2 , COLUMNA 1
    .Cell(1, 2).Merge .Cell(1, 6)
    
    .Cell(1, 1).Range.Text = "NOMBRES Y APELLIDOS"
    .Cell(1, 1).Range.Font.Bold = True
    .Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 2).Range.Text = "ENTIDADES EN LAS QUE TIENEN PARTICIPACION PATRIMONIAL"
    .Cell(1, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 2).Range.Font.Bold = True
    .Cell(2, 2).Range.Text = "RAZON SOCIAL"
    .Cell(2, 2).Range.Font.Bold = True
    .Cell(2, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    '.Cell(1, 2).Range.Font.Color = wdColorBlue 'cambia coolor de letra
    .Cell(2, 3).Range.Text = "RUC"
    .Cell(2, 3).Range.Font.Bold = True
    .Cell(2, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(2, 4).Range.Text = "APORTES S/."
    .Cell(2, 4).Range.Font.Bold = True
    .Cell(2, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(2, 5).Range.Text = "%"
    .Cell(2, 5).Range.Font.Bold = True
    .Cell(2, 5).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(2, 6).Range.Text = "Cod. SBS(1)"
    .Cell(2, 6).Range.Font.Bold = True
    .Cell(2, 6).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
    i = 2
    Do While Not pRDatP.EOF
    
        i = i + 1
        
            .Cell(i, 1).Range.Text = pRDatP!cNombreJur
            .Cell(i, 2).Range.Text = pRDatP!cRazonSocial
            .Cell(i, 3).Range.Text = pRDatP!vRuc
            .Cell(i, 4).Range.Text = Format(pRDatP!nApoSol, "#,#0.00")
            .Cell(i, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
            .Cell(i, 5).Range.Text = pRDatP!nApoPorc
            .Cell(i, 6).Range.Text = pRDatP!vCodSbs
    
        pRDatP.MoveNext
    Loop
      
End With

Else
filas = 3
columnas = 6

oWord.Selection.GoTo WHAT:=wdGoToBookmark, Name:="patrimonio"  'se coloca en la primera fila al buscar el marcador cuadro
oWord.Selection.Font.Size = 10
oWord.Selection.Font.Bold = False

Set Tabla = oDoc.Tables.Add(oWord.Selection.Range, filas, columnas, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
wdAutoFitFixed)

'CABECERA DE LA TABLA
With Tabla
    .Borders.OutsideLineStyle = wdLineStyleSingle
    .Borders.InsideLineStyle = wdLineStyleSingle
    
    .Columns(1).Width = 180
    .Columns(2).Width = 100
    .Columns(3).Width = 70
    .Columns(4).Width = 80
    .Columns(5).Width = 30
    .Columns(6).Width = 50
    
    'combina celdas
    .Cell(1, 1).Merge .Cell(2, 1) 'COMBINA FILA 1 -2 , COLUMNA 1
    .Cell(1, 2).Merge .Cell(1, 6)
    
    .Cell(1, 1).Range.Text = "NOMBRES Y APELLIDOS"
    .Cell(1, 1).Range.Font.Bold = True
    .Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 2).Range.Text = "ENTIDADES EN LAS QUE TIENEN PARTICIPACION PATRIMONIAL"
    .Cell(1, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 2).Range.Font.Bold = True
    .Cell(2, 2).Range.Text = "RAZON SOCIAL"
    .Cell(2, 2).Range.Font.Bold = True
    .Cell(2, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    '.Cell(1, 2).Range.Font.Color = wdColorBlue 'cambia coolor de letra
    .Cell(2, 3).Range.Text = "RUC"
    .Cell(2, 3).Range.Font.Bold = True
    .Cell(2, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(2, 4).Range.Text = "APORTES S/."
    .Cell(2, 4).Range.Font.Bold = True
    .Cell(2, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(2, 5).Range.Text = "%"
    .Cell(2, 5).Range.Font.Bold = True
    .Cell(2, 5).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(2, 6).Range.Text = "Cod. SBS(1)"
    .Cell(2, 6).Range.Font.Bold = True
    .Cell(2, 6).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
    i = 2
    Do While Not pRDatP.EOF
    
        i = i + 1
        
            .Cell(i, 1).Range.Text = "..........."
            .Cell(i, 2).Range.Text = "..........."
            .Cell(i, 3).Range.Text = "..........."
            .Cell(i, 4).Range.Text = "..........."
            .Cell(i, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
            .Cell(i, 5).Range.Text = "..........."
            .Cell(i, 6).Range.Text = "..........."
    
        pRDatP.MoveNext
    Loop
      
End With
    
End If

'****************** Cargos

If pRDatCa.RecordCount > 0 Then

'Dim i As Integer
'Dim Tabla As Word.Table
'Dim filas, columnas As Integer

filas = pRDatCa.RecordCount + 2
columnas = 5

oWord.Selection.GoTo WHAT:=wdGoToBookmark, Name:="cargos"  'se coloca en la primera fila al buscar el marcador cuadro
oWord.Selection.Font.Size = 10
oWord.Selection.Font.Bold = False

Set Tabla = oDoc.Tables.Add(oWord.Selection.Range, filas, columnas, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
wdAutoFitFixed)

'CABECERA DE LA TABLA
With Tabla
    .Borders.OutsideLineStyle = wdLineStyleSingle
    .Borders.InsideLineStyle = wdLineStyleSingle

    .Columns(1).Width = 180
    .Columns(2).Width = 140
    .Columns(3).Width = 50
    .Columns(4).Width = 140
    .Columns(5).Width = 25

    'combina celdas
    .Cell(1, 1).Merge .Cell(2, 1) 'COMBINA FILA 1 -2 , COLUMNA 1
    .Cell(1, 2).Merge .Cell(1, 5)

    .Cell(1, 1).Range.Text = "NOMBRES Y APELLIDOS"
    .Cell(1, 1).Range.Font.Bold = True
    .Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 2).Range.Text = "OTRAS ENTIDADES EN LAS QUE TIENEN PARTICIPACION PATRIMONIAL"
    .Cell(1, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 2).Range.Font.Bold = True
    .Cell(2, 2).Range.Text = "RAZON SOCIAL"
    .Cell(2, 2).Range.Font.Bold = True
    .Cell(2, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    '.Cell(1, 2).Range.Font.Color = wdColorBlue 'cambia coolor de letra
    .Cell(2, 3).Range.Text = "RUC"
    .Cell(2, 3).Range.Font.Bold = True
    .Cell(2, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(2, 4).Range.Text = " CARGO Y FUNCION "
    .Cell(2, 4).Range.Font.Bold = True
    .Cell(2, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(2, 5).Range.Text = "Cod. SBS(1)"
    .Cell(2, 5).Range.Font.Bold = True
    .Cell(2, 5).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter


    i = 2
    Do While Not pRDatCa.EOF

        i = i + 1

            .Cell(i, 1).Range.Text = pRDatCa!cNombreJur
            .Cell(i, 2).Range.Text = pRDatCa!cRazonSocial
            .Cell(i, 3).Range.Text = pRDatCa!vRuc
            .Cell(i, 4).Range.Text = pRDatCa!cCargo
            .Cell(i, 5).Range.Text = pRDatCa!vCodSbs

        pRDatCa.MoveNext
    Loop

End With

Else
filas = 3
columnas = 5

oWord.Selection.GoTo WHAT:=wdGoToBookmark, Name:="cargos"  'se coloca en la primera fila al buscar el marcador cuadro
oWord.Selection.Font.Size = 10
oWord.Selection.Font.Bold = False

Set Tabla = oDoc.Tables.Add(oWord.Selection.Range, filas, columnas, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
wdAutoFitFixed)

'CABECERA DE LA TABLA
With Tabla
    .Borders.OutsideLineStyle = wdLineStyleSingle
    .Borders.InsideLineStyle = wdLineStyleSingle

    .Columns(1).Width = 180
    .Columns(2).Width = 140
    .Columns(3).Width = 50
    .Columns(4).Width = 140
    .Columns(5).Width = 25

    'combina celdas
    .Cell(1, 1).Merge .Cell(2, 1) 'COMBINA FILA 1 -2 , COLUMNA 1
    .Cell(1, 2).Merge .Cell(1, 5)

    .Cell(1, 1).Range.Text = "NOMBRES Y APELLIDOS"
    .Cell(1, 1).Range.Font.Bold = True
    .Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 2).Range.Text = "OTRAS ENTIDADES EN LAS QUE TIENEN PARTICIPACION PATRIMONIAL"
    .Cell(1, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(1, 2).Range.Font.Bold = True
    .Cell(2, 2).Range.Text = "RAZON SOCIAL"
    .Cell(2, 2).Range.Font.Bold = True
    .Cell(2, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    '.Cell(1, 2).Range.Font.Color = wdColorBlue 'cambia coolor de letra
    .Cell(2, 3).Range.Text = "RUC"
    .Cell(2, 3).Range.Font.Bold = True
    .Cell(2, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(2, 4).Range.Text = " CARGO Y FUNCION "
    .Cell(2, 4).Range.Font.Bold = True
    .Cell(2, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Cell(2, 5).Range.Text = "Cod. SBS(1)"
    .Cell(2, 5).Range.Font.Bold = True
    .Cell(2, 5).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter


    i = 2
    Do While Not pRDatCa.EOF

        i = i + 1

            .Cell(i, 1).Range.Text = "............."
            .Cell(i, 2).Range.Text = "............."
            .Cell(i, 3).Range.Text = "............."
            .Cell(i, 4).Range.Text = "............."
            .Cell(i, 5).Range.Text = "............."

        pRDatCa.MoveNext
    Loop

End With

End If

'*********************************************************************
'****************************************

    oDoc.Close
    Set oDoc = Nothing

    Set oWord = CreateObject("Word.Application")
    oWord.Visible = True
    Set oDoc = oWord.Documents.Open(sArchivo)
    Set oDoc = Nothing
    Set oWord = Nothing
    Set Tabla = Nothing

    Exit Function
'PEAC 20090520
ErrorInformeCom01:
    MsgBox err.Description, vbCritical, "Aviso"
    Set oDoc = Nothing
    Set oWord = Nothing

End Function

'**DAOR 20080716, se copió de gCredFunciones
Public Function ImprimeCartillaCred(ByVal psCtaCod As String) As String

    Dim R As ADODB.Recordset
    Dim nTasaAnual As Double
    Dim nTasaMoraA As Double

    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range

    Dim RAge As ADODB.Recordset
    Dim sCompaSeguro As String
    Dim sNumeroPoliza As String
    Dim nParametro As Double
    Dim RRelaCred  As ADODB.Recordset
    Dim oCred As COMNCredito.NCOMCredDoc
    Dim nITF As Double 'EAAS 20180321
    Dim nCosRef As Double
    Dim nComDsg As Double
    Dim nCosCroPag As Double
    Dim nCosHisPag As Double
    Dim nCosComPag As Double
    Dim nMontoInt As Double
    Dim nPriPol As Double
    Dim nPriPolInc As Double 'MAVM 20120215
    Dim nCantAnt As Double
    Dim nCenCrg As Double
    Dim nSgrDsg1 As Double
    Dim nSgrDsg2 As Double
    Dim nSegMor As Double
    Dim nMinCuo As Double
    Dim nMaxCuo As Double
    'peac 20070815
    Dim nTotPagar As Double, nTasaIntMorEfeAnual As Double, i As Integer
    'fin peac
    Dim bPoliza As Boolean 'CTI5  ERS0012021
    'MAVM 20100412 ***
    Dim nCosComPag1 As Double
    '***
        
    'MAVM 25112009 ***
    Dim nConScore As Double
    '***
    'JUEZ 20150724 *************************
    Dim oCredRel As COMDCredito.UCOMCredRela
    Dim RSeg As ADODB.Recordset
    Dim nNroTitSegDes As Integer
    Dim nValorGastoSeg As Double
    'END JUEZ ******************************
    
    Dim sArchivo As String
    Dim nMonAseguradoraPol As String 'MADM 20111213
    'EJVG20131001 ***
    Dim oDCred As New COMDCredito.DCOMCredito
    Dim rsRelEcotaxi As New ADODB.Recordset
    Dim sNombreAseguradoraPol As String
    Dim nPriPolVehic As Double
    Dim lsTpoProdCod As String
    Dim lsEtiqDistrEcotaxi As String
    'END EJVG *******
    'INICIO EAAS20180405 CAMBIO HOJAS RESUMENES *******
    Dim bPolExterna As Boolean
    Dim rs As New ADODB.Recordset
    Dim oPol As New COMDCredito.DCOMPoliza
    Set oPol = New COMDCredito.DCOMPoliza
    Set rs = oPol.RecuperaPolizasInternasCuenta(psCtaCod)
        
        Dim rsmr As ADODB.Recordset 'gemo
    Dim cCodSegMR As String
    Dim oSM As COMDConstSistema.DCOMConstSistema   'gemo
    Set oSM = New COMDConstSistema.DCOMConstSistema 'gemo
    Set rsmr = oSM.RecuperaSeguroMultiriesgo(psCtaCod) ' gemo
        
        If Not (rsmr.BOF And rsmr.EOF) Then
        cCodSegMR = rsmr!nConsSisValor
    End If
        
    If Not rs.EOF Then
    Dim X As Integer
    For X = 0 To rs.RecordCount - 1
        If rs!bPolizaExterna <> "" Then
            bPolExterna = rs!bPolizaExterna
        End If
        rs.MoveNext
    Next
    End If
    'FIN EAAS20180405 CAMBIO HOJAS RESUMENES*******
    Set oPol = Nothing
    On Error GoTo ErrorCartillaCred

    Set oCred = New COMNCredito.NCOMCredDoc
    'MAVM 20120215 ***
    'Call oCred.RecuperaDatosHojaResumenCreditos(psCtaCod, R, RAge, RRelaCred, sCompaSeguro, sNumeroPoliza, nCosRef, nComDsg, nCosCroPag, nCosHisPag, nCosComPag, nMontoInt, nPriPol, nCantAnt, nCenCrg, nSgrDsg1, nSgrDsg2, nSegMor, nMinCuo, nMaxCuo, nConScore, nCosComPag1, nMonAseguradoraPol)
    'Call oCred.RecuperaDatosHojaResumenCreditos(psCtaCod, R, RAge, RRelaCred, sCompaSeguro, sNumeroPoliza, nCosRef, nComDsg, nCosCroPag, nCosHisPag, nCosComPag, nMontoInt, nPriPol, nCantAnt, nCenCrg, nSgrDsg1, nSgrDsg2, nSegMor, nMinCuo, nMaxCuo, nConScore, nCosComPag1, nMonAseguradoraPol, nPriPolInc)
    Call oCred.RecuperaDatosHojaResumenCreditos(psCtaCod, R, RAge, RRelaCred, sCompaSeguro, sNumeroPoliza, nCosRef, nComDsg, nCosCroPag, nCosHisPag, nCosComPag, nMontoInt, nPriPol, nCantAnt, nCenCrg, nSgrDsg1, nSgrDsg2, nSegMor, nMinCuo, nMaxCuo, nConScore, nCosComPag1, nMonAseguradoraPol, nPriPolInc, sNombreAseguradoraPol, nPriPolVehic, bPoliza) 'EJVG20131001
    '***
    
    Set oCred = Nothing

    If R.EOF And R.BOF Then
        MsgBox "No existen Datos para el Reporte...", vbInformation, "Aviso"
        Exit Function
    End If
    
    'JUEZ 20150724 ************************************
    Set oCredRel = New COMDCredito.UCOMCredRela
    Set RSeg = oCredRel.ObtenerTitularesYTasaSegDes(psCtaCod)
    
    If Not (RSeg.EOF And RSeg.BOF) Then
        nNroTitSegDes = CInt(RSeg!nAsegurados)
        nValorGastoSeg = RSeg!nTasa
        If nNroTitSegDes = 0 Then nValorGastoSeg = 0 'JUEZ 20150803
    Else
        nNroTitSegDes = 0
        nValorGastoSeg = 0
    End If
    'END JUEZ *****************************************
    
    Set oWord = CreateObject("Word.Application")
        oWord.Visible = False
    'peac 20070815
    'Set oDoc = oWord.Documents.Open(App.path & "\FormatoCarta\HojaResumen.doc")
    
    'If Mid(psCtaCod, 6, 3) >= "401" And Mid(psCtaCod, 6, 3) <= "423" Then
        'Set oDoc = oWord.Documents.Open(App.path & "\FormatoCarta\HojaResumenH.doc")
    'Else
    'Set oDoc = oWord.Documents.Open(App.path & "\FormatoCarta\HojaResumen.doc")
    'End If
    'fin peac
    'EJVG20130930 ***
    lsTpoProdCod = IIf(IsNull(R!cTpoProdCod), "", R!cTpoProdCod)
    '**ARLO20180712 ERS042 - 2018
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("N0000142", lsTpoProdCod) Then
    'If lsTpoProdCod = "517" Then
    '**ARLO20180712 ERS042 - 2018
        Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\HojaResumenEcotaxi.doc")
    'MAVM 20140319 ***
    ElseIf objProducto.GetResultadoCondicionCatalogo("N0000143", lsTpoProdCod) Then
    'ElseIf lsTpoProdCod = "704" Then
        If sNumeroPoliza = "" Then
            Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\HojaResumenConv.doc")
        Else
            Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\HojaResumenConvAse.doc")
        End If
    '***
    Else
                If lsTpoProdCod = "802" Or lsTpoProdCod = "806" Then
            If lsTpoProdCod = "802" Then
                  Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\HojaResumen_TechoPropio.doc")
            Else
                  Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\HojaResumen_MiVivienda.doc")
            End If
        Else
                        If sNumeroPoliza = "" Then  'EAAS20180405 CAMBIO HOJAS RESUMENES
                                If nNroTitSegDes = 0 Then
                                Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\HojaResumenSinDesgravamen.doc")
                                Else
                                Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\HojaResumen.doc")
                                End If
                        ElseIf bPolExterna Then 'EAAS20180405 CAMBIO HOJAS RESUMENES
                         Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\HojaResumenAseExterna.doc")
                        ElseIf nNroTitSegDes = 0 Then 'EAAS20180405 CAMBIO HOJAS RESUMENES
                         Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\HojaResumenAseSinDesgravamen.doc")
                        Else
                                Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\HojaResumenAse.doc")
                        End If
                End If
    End If
    'END EJVG *******

    'ARCV 05-05-2007
    'ARCV 03-05-2007
    'sArchivo = App.Path & "\FormatoCarta\HR_" & psCtaCod & "_" & Replace(Left(Time, 5), ":", "") & ".doc"'COmento JOEP20180811
    sArchivo = App.Path & "\FormatoCarta\HR_" & psCtaCod & "_" & Replace(Left(Time, 8), ":", "") & ".doc" 'JOEP20180811
    oDoc.SaveAs (sArchivo)
    '----------
    'CTI3 ERS0032020*******************************************
    If lsTpoProdCod = "802" Or lsTpoProdCod = "806" Then
            With oWord.Selection.Find
                .Text = "<<nBBP>>"
                .Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", gcPEN_SIMBOLO & " ", "US$") & " " & Format(R!pnMontoBonoMiVivienda, "#0.00")
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            
            If lsTpoProdCod = "806" Then
                With oWord.Selection.Find
                    .Text = "<<nPBP>>"
                    .Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", gcPEN_SIMBOLO & " ", "US$") & " " & Format(R!pnMontoPremioMiVivienda, "#0.00")
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
            End If
            
            If nPriPolInc = 0 Then
                 With oWord.Selection.Find
                    .Text = "<<cPriInc>>"
                    .Replacement.Text = "" 'marg ers044-2016
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With

            Else
                With oWord.Selection.Find
                    .Text = "<<cPriInc>>"
                    '''.Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", "S/.", "US$") & " " & Format(nPriPolInc, "#0.00") 'marg ers044-2016
                    .Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", gcPEN_SIMBOLO & " ", "US$") & " " & Format(nPriPolInc, "#0.00") 'marg ers044-2016
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
            End If
            
            Dim cComent1 As String
            Dim cComent2 As String
            Dim cCiaSeg2 As String
            cComent1 = ""
            cComent2 = ""
            cCiaSeg2 = ""
            If bPoliza = True Then
                cComent1 = "- 0.20%o (Incluido IGV) sobre la suma asegurada"
                cComent2 = "- Prima mínima (Incluido IGV): S/ 15.00  o US$ 4.55"
                cCiaSeg2 = sCompaSeguro
            End If
            With oWord.Selection.Find
                .Text = "<<cComent1>>"
                .Replacement.Text = cComent1
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With oWord.Selection.Find
                .Text = "<<cComent2>>"
                .Replacement.Text = cComent2
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With oWord.Selection.Find
                .Text = "<<pnPeriodoGracia>>"
                .Replacement.Text = CStr(R!pnPeriodoGRacia)
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
              With oWord.Selection.Find
                .Text = "<<cCiaSeg2>>"
                .Replacement.Text = cCiaSeg2
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            
            
    End If
    '**********************************************************
    'INICIO EAAS 20180321 CAMBIO HOJAS RESUMENES
    With oWord.Selection.Find
        .Text = "<<nCuotas>>"
        .Replacement.Text = R!nCuotas
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    If nNroTitSegDes = 0 Then
        With oWord.Selection.Find
            .Text = "<<cNumPolizaDesgravamen>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
    Else
                Dim sNumPolizaDesgravamen As String
                sNumPolizaDesgravamen = ""
                'If Not ((lsTpoProdCod = "802" Or lsTpoProdCod = "806") And bPoliza = False) Then
                        sNumPolizaDesgravamen = IIf(Mid(psCtaCod, 9, 1) = "1", "672017", "672029")
               ' End If
        
        With oWord.Selection.Find
            .Text = "<<cNumPolizaDesgravamen>>"
            .Replacement.Text = sNumPolizaDesgravamen
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
    End If
    
    nITF = gnITFPorcent * 100
    With oWord.Selection.Find
        .Text = "<<TasaITF>>"
        '.Replacement.Text = Format$(nitf, "0.00")
        .Replacement.Text = Trim(CStr(nITF))
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'FIN EAAS 20180321 CAMBIO HOJAS RESUMENES
    
    
    With oWord.Selection.Find
        .Text = "<<cCodCta>>"
        .Replacement.Text = psCtaCod
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cTipCta>>"
        'ALPA 20100607 B2*******************
        '.Replacement.Text = R!cTipoCredDescrip
        .Replacement.Text = R!cTpoCredDes
        '***********************************
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cCapDes>>"
        '''.Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", "S/.", "US$") & Format(R!nMontoPagare, "#0.00") 'marg ers044-2016
        .Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", gcPEN_SIMBOLO & " ", "US$") & Format(R!nMontoPagare, "#0.00") 'marg ers044-2016
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cMaxCuo>>"
        '''.Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", "S/.", "US$") & Format(nMaxCuo, "#0.00") 'marg ers044-2016
        .Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", gcPEN_SIMBOLO & " ", "US$") & Format(nMaxCuo, "#0.00") 'marg ers044-2016
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cMinCuo>>"
        '''.Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", "S/.", "US$") & Format(nMinCuo, "#0.00") 'marg ers044-2016
        .Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", gcPEN_SIMBOLO & " ", "US$") & Format(nMinCuo, "#0.00") 'marg ers044-2016
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'peac 20080715
    With oWord.Selection.Find
        .Text = "<<cFecPgo>>"
        .Replacement.Text = IIf(IsNull(R!nFecPgoCtas), "***", Str(R!nFecPgoCtas) & " DE CADA MES")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'fin peac

    With oWord.Selection.Find
        .Text = "<<cSegMor>>"
        '''.Replacement.Text = Format(nSegMor, "#0.00") & " Soles" 'MAVM 02112009 "%" 'marg ers044-2016
        .Replacement.Text = Format(nSegMor, "#0.00") & " " & StrConv(gcPEN_PLURAL, vbProperCase) 'MAVM 02112009 "%" 'marg ers044-2016
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cSgrDsg2>>"
        .Replacement.Text = Format(nSgrDsg2, "#0.0000") & "%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cSgrDsg1>>"
        .Replacement.Text = Format(nSgrDsg1, "#0.0000") & "%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    'JUEZ 20150724 *************************************
    'INICIO EAAS20180405 CAMBIO HOJAS RESUMENES
    With oWord.Selection.Find
            .Text = "<<cSgrDsg>>"
            .Replacement.Text = Format(nValorGastoSeg, "#0.0000") & "%"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
        .Text = "<<nNroTitSegDes>>"
        .Replacement.Text = CStr(nNroTitSegDes) & IIf(nNroTitSegDes > 1 Or nNroTitSegDes = 0, " titulares", " titular")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'FIN EAAS20180405 CAMBIO HOJAS RESUMENES
    'END JUEZ ******************************************

    With oWord.Selection.Find
        .Text = "<<cCenCrg>>"
        '.Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", "S/.", "US$") & Format(nCenCrg, "#0.00")
        '''.Replacement.Text = "S/." & Format(nCenCrg, "#0.00") 'marg ers044-2016
        .Replacement.Text = gcPEN_SIMBOLO & " " & Format(nCenCrg, "#0.00") 'marg ers044-2016
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
        .Text = "<<cConScore>>"
        '''.Replacement.Text = "S/." & Format(nConScore, "#0.00") 'marg ers044-2016
        .Replacement.Text = gcPEN_SIMBOLO & " " & Format(nConScore, "#0.00") 'marg ers044-2016
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cCantAnt>>"
        '''.Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", "S/.", "US$") & Format(nCantAnt, "#0.00") 'marg ers044-2016
        .Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", gcPEN_SIMBOLO & " ", "US$") & Format(nCantAnt, "#0.00") 'marg ers044-2016
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'INICIO EAAS20180405 CAMBIO HOJAS RESUMENES
        Dim sPriPol As String
    sPriPol = ""
'    If Not ((lsTpoProdCod = "802" Or lsTpoProdCod = "806") And bPoliza = False) Then
        sPriPol = IIf(Mid(psCtaCod, 9, 1) = "1", gcPEN_SIMBOLO & " ", "US$") & " " & Format(nPriPol, "#0.00")
'    End If
    With oWord.Selection.Find
            .Text = "<<cPriPol>>"
            .Replacement.Text = sPriPol
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
    'FIN EAAS20180405 CAMBIO HOJAS RESUMENES
    With oWord.Selection.Find
        .Text = "<<cCosComPag>>"
        .Replacement.Text = gcPEN_SIMBOLO & " " & Format(nCosComPag, "#0.00")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    'MAVM 20100412 ***
    With oWord.Selection.Find
        .Text = "<<cCosComPag1>>"
        '''.Replacement.Text = "S/." & Format(nCosComPag1, "#0.00") 'marg ers044-2016
        .Replacement.Text = gcPEN_SIMBOLO & " " & Format(nCosComPag1, "#0.00") 'marg ers044-2016
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    '***

    With oWord.Selection.Find
        .Text = "<<cCosHisPag>>"
        '.Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", "S/.", "US$") & Format(nCosHisPag, "#0.00")
        '''.Replacement.Text = "S/." & Format(nCosHisPag, "#0.00") 'marg ers044-2016
        .Replacement.Text = gcPEN_SIMBOLO & " " & Format(nCosHisPag, "#0.00") 'marg ers044-2016
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cCosCroPag>>"
        '.Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", "S/.", "US$") & Format(nCosCroPag, "#0.00")
        '''.Replacement.Text = "S/." & Format(nCosCroPag, "#0.00") 'marg ers044-2016
        .Replacement.Text = gcPEN_SIMBOLO & " " & Format(nCosCroPag, "#0.00") 'marg ers044-2016
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cComDsg>>"
        '.Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", "S/.", "US$") & Format(nComDsg, "#0.00")
        '''.Replacement.Text = "S/." & Format(nComDsg, "#0.00") 'marg ers044-2016
        .Replacement.Text = gcPEN_SIMBOLO & " " & Format(nComDsg, "#0.00") 'marg ers044-2016
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cCosRef>>"
        '.Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", "S/.", "US$") & Format(nCosRef, "#0.00")
        '''.Replacement.Text = "S/." & Format(nCosRef, "#0.00") 'marg ers044-2016
        .Replacement.Text = gcPEN_SIMBOLO & " " & Format(nCosRef, "#0.00") 'marg ers044-2016
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cMonInt>>"
        '''.Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", "S/.", "US$") & Format(nMontoInt, "#0.00") 'marg ers044-2016
        .Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", gcPEN_SIMBOLO & " ", "US$") & Format(nMontoInt, "#0.00") 'marg ers044-2016
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

'MADM 20101130
'    With oWord.Selection.Find
'        .Text = "<<cTasInt>>"
'        .Replacement.Text = Format(R!nTasaInteres, "#0.00") & " %"
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With

    With oWord.Selection.Find
        .Text = "<<cTEAInt>>"
        .Replacement.Text = Format(((1 + R!nTasaInteres / 100) ^ (360 / 30) - 1) * 100, "#.00") & " %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    nTasaMoraA = (((1 + R!nTasaMora / 100) ^ (30)) - 1) * 100
'MADM 20101130
'    With oWord.Selection.Find
'        .Text = "<<cTasMor>>"
'        '.Replacement.Text = Format(R!nTasaMora, "#0.00") & " %"
'        .Replacement.Text = Format(nTasaMoraA, "#0.00") & " %"
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With

    'peac 20070815
    'nTasaIntMorEfeAnual = (((1 + nTasaMoraA / 100) ^ (360 / 30) - 1) * 100) 'Comentado por JUEZ 20131121

    With oWord.Selection.Find
        .Text = "<<cTEAMor>>"
        'peac 20070815
        '.Replacement.Text = Format(((1 + nTasaMoraA / 100) ^ (360 / 30) - 1) * 100, "#.00") & " %"
        '.Replacement.Text = Format(nTasaIntMorEfeAnual, "#.00") & " %"
        .Replacement.Text = Format(R!nTEAMora, "#.00") & " %" 'JUEZ 20131121
        'end peac
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'peac 20070815
    nTotPagar = R!nMontoPagare + nMontoInt

    With oWord.Selection.Find
        .Text = "<<cTotPgar>>"
        '''.Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", "S/.", "US$") & Format(nTotPagar, "#0.00") 'marg ers044-2016
        .Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", gcPEN_SIMBOLO & " ", "US$") & Format(nTotPagar, "#0.00") 'marg ers044-2016
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'end peac

    With oWord.Selection.Find
        .Text = "<<FechaHoja>>"
        .Replacement.Text = Format(gdFecSis, "dddd") & " " & Day(gdFecSis) & " de " & Format(gdFecSis, "mmmm") & " de " & Year(gdFecSis) 'EAAS20180321 SE QUITO LA COMA A LA FECHA
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cNomAge>>"
        .Replacement.Text = RAge!cAgeDescripcion
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cDirAge>>"
        .Replacement.Text = RAge!cAgeDireccion
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cDisAge>>"
        .Replacement.Text = Trim(RAge!Dist)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cCosRef>>"
        '''.Replacement.Text = "S/." & CStr(Format(nParametro, "#0.00")) 'marg ers044-2016
        .Replacement.Text = gcPEN_SIMBOLO & " " & CStr(Format(nParametro, "#0.00")) 'marg ers044-2016
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cComDsg>>"
        '''.Replacement.Text = "S/." & CStr(Format(nParametro, "#0.00")) 'marg ers044-2016
        .Replacement.Text = gcPEN_SIMBOLO & " " & CStr(Format(nParametro, "#0.00")) 'marg ers044-2016
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cCosCroPag>>"
        '''.Replacement.Text = "S/." & CStr(Format(nParametro, "#0.00")) 'marg ers044-2016
        .Replacement.Text = gcPEN_SIMBOLO & " " & CStr(Format(nParametro, "#0.00")) 'marg ers044-2016
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cCosHisPag>>"
        '''.Replacement.Text = "S/." & CStr(Format(nParametro, "#0.00")) 'marg ers044-2016
        .Replacement.Text = gcPEN_SIMBOLO & " " & CStr(Format(nParametro, "#0.00")) 'marg ers044-2016
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cCosComPag>>"
        '''.Replacement.Text = "S/." & CStr(Format(nParametro, "#0.00")) 'marg ers044-2016
        .Replacement.Text = gcPEN_SIMBOLO & " " & CStr(Format(nParametro, "#0.00")) 'marg ers044-2016
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'INICIO EAAS20180405 CAMBIO HOJAS RESUMENES
    'If ((lsTpoProdCod = "802" Or lsTpoProdCod = "806") And bPoliza = False) Then
        'sCompaSeguro = ""
    'End If
    With oWord.Selection.Find
            .Text = "<<cCiaSeg>>"
            .Replacement.Text = sCompaSeguro
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
    End With
   'FIN EAAS20180405 CAMBIO HOJAS RESUMENES
'EAAS INICIO 20180323 CAMBIO HOJAS RESUMENES

'    With oWord.Selection.Find
'        .Text = "<<cNroPol>>"
'        .Replacement.Text = sNumeroPoliza
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
        Dim sCodSegMR As String
    sCodSegMR = ""
    If Not ((lsTpoProdCod = "802" Or lsTpoProdCod = "806") And bPoliza = False) Then
        sCodSegMR = cCodSegMR
    End If
    With oWord.Selection.Find
        .Text = "<<cNroPol>>"
        '.Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", "1841716", "1841712") 'COMENTADO GEMO 13/02/2020
                .Replacement.Text = sCodSegMR ' ADD GEMO 13/02/2020
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
'EAAS FIN 20180323 CAMBIO HOJAS RESUMENES

    
    'MAVM 20140319 ***
    '**ARLO20180712 ERS042 - 2018
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("N0000144", lsTpoProdCod) Then
    'If lsTpoProdCod = "704" Then
    '**ARLO20180712 ERS042 - 2018
    
        Dim rsConv As New ADODB.Recordset
        Dim oDCredDoc As New COMDCredito.DCOMCredDoc
        Dim sValor As String
        Dim nCasill As Integer 'CTI2 20190126 Mejora de Ingresos
        
        nCasill = 0 'CTI2 20190126 Mejora de Ingresos
        Set rsConv = oDCredDoc.RecuperaGastoConv(psCtaCod)
        While Not rsConv.EOF
            With oWord.Selection.Find
                .Text = "<<cInstConv>>"
                .Replacement.Text = Trim(rsConv!cInstConv)
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            
            If rsConv!nTpoValor = 1 Then
                'sValor = "S/." & rsConv!nValor
                '''sValor = "S/." & Format(rsConv!nValor, "#0.00") 'JUEZ 20150731 'marg ers044-2016
                sValor = gcPEN_SIMBOLO & " " & Format(rsConv!nValor, "#0.00") 'JUEZ 20150731 'marg ers044-2016
            ElseIf rsConv!nTpoValor = 2 Then
                'sValor = rsConv!nValor & "%"
                sValor = Format(rsConv!nValor, "#0.00") & "%" 'JUEZ 20150731
            Else 'JUEZ 20150121
                sValor = "0.00"
            End If
            'INICIO EAAS20180405
'            With oWord.Selection.Find
'                .Text = "<<cGastoConv>>"
'                .Replacement.Text = sValor
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .Execute Replace:=wdReplaceAll
'            End With
            'FIN EAAS 20180405
            
            'CTI2 20181219 ADD
            ' Aplica si el gasto es por casillero, en casi de haber otros gastos
            ' Se recomienda configurarlo desde esta linea
            If rsConv!nAplicaConvenio = 1 Then
                With oWord.Selection.Find
                    .Text = "<<cGastoConv>>"
                    .Replacement.Text = sValor
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                nCasill = nCasill + 1
            End If
            'CTI2 FIN
            rsConv.MoveNext
        Wend
        'COMENTADO POR CTI2
        'With oWord.Selection.Find
        '    .Text = "<<cGastoConv>>"
        '    .Replacement.Text = sValor
        '    .Forward = True
        '    .Wrap = wdFindContinue
        '    .Format = False
        '    .Execute Replace:=wdReplaceAll
        'End With
        
        ' CTI2 20190126 Mejora de Ingresos - INI
        If nCasill = 0 Then
            With oWord.Selection.Find
                .Text = "<<cGastoConv>>"
                .Replacement.Text = gcPEN_SIMBOLO & " 0.00"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
        End If
        'CTI2 20190126 Mejora de Ingresos - FIN
        
        Set rsConv = Nothing
        Set oDCredDoc = Nothing
    End If
    '***
    
    'MADM 20111213
    If sNumeroPoliza <> "" Then
    With oWord.Selection.Find
        .Text = "<<cCiaASegPol>>"
        .Replacement.Text = nMonAseguradoraPol
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    'MAVM 20120215 ***
    With oWord.Selection.Find
        .Text = "<<cPriInc>>"
        '''.Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", "S/.", "US$") & " " & Format(nPriPolInc, "#0.00") 'marg ers044-2016
        .Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", gcPEN_SIMBOLO & " ", "US$") & " " & Format(nPriPolInc, "#0.00") 'marg ers044-2016
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    '***
    'EJVG20131001 ***
    With oWord.Selection.Find
        .Text = "<<cCiaASegPolVehic>>"
        .Replacement.Text = sNombreAseguradoraPol
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cPriVehic>>"
        '''.Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", "S/.", "US$") & " " & Format(nPriPolVehic, "#0.00") 'marg ers044-2016
        .Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", gcPEN_SIMBOLO & " ", "US$") & " " & Format(nPriPolVehic, "#0.00") 'marg ers044-2016
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END EJVG *******
    
    End If
    'END MADM
    'EJVG20131002 ***
    '**ARLO20180712 ERS042 - 2018
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("N0000145", lsTpoProdCod) Then
    'If lsTpoProdCod = "517" Then
    '**ARLO20180712 ERS042 - 2018
        Set rsRelEcotaxi = oDCred.RecuperaRelacionesInfogas(psCtaCod)
        While Not rsRelEcotaxi.EOF
            lsEtiqDistrEcotaxi = ""
            Select Case CInt(Trim(Right(rsRelEcotaxi!cRelacion, 4)))
                Case 12: lsEtiqDistrEcotaxi = "<<cOpeCertif>>"
                Case 14: lsEtiqDistrEcotaxi = "<<cServNotarial>>"
                Case 15: lsEtiqDistrEcotaxi = "<<cSOAT>>"
            End Select
            If lsEtiqDistrEcotaxi <> "" Then
                With oWord.Selection.Find
                    .Text = lsEtiqDistrEcotaxi
                    '''.Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", "S/.", "US$") & " " & Format(rsRelEcotaxi!nMontoAbono, "#0.00") 'marg ers044-2016
                    .Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", gcPEN_SIMBOLO & " ", "US$") & " " & Format(rsRelEcotaxi!nMontoAbono, "#0.00") 'marg ers044-2016
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
            End If
            rsRelEcotaxi.MoveNext
        Wend
    End If
    'END EJVG *******
    'peac 20070815
    i = 0

    If Not (RRelaCred.EOF And RRelaCred.BOF) Then
       Do Until RRelaCred.EOF
          If RRelaCred!nConsValor = gColRelPersTitular Then
               With oWord.Selection.Find
                    .Text = "<<cNomCli>>"
                    .Replacement.Text = PstaNombre(RRelaCred!cPersNombre)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
               End With

               With oWord.Selection.Find
                    .Text = "<<cNuDoci>>"
                    .Replacement.Text = IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
               End With

          ElseIf RRelaCred!nConsValor = gColRelPersCodeudor Then
               With oWord.Selection.Find
                    .Text = "<<cNomCli2>>"
                    .Replacement.Text = PstaNombre(RRelaCred!cPersNombre)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
               End With

               With oWord.Selection.Find
                    .Text = "<<cNuDoci2>>"
                    .Replacement.Text = IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
               End With
        'peac 20070815
          ElseIf RRelaCred!nConsValor = gColRelPersGarante Then
                i = i + 1
               With oWord.Selection.Find
                    .Text = "<<cAval" & i & ">>"
                    .Replacement.Text = PstaNombre(RRelaCred!cPersNombre)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
               End With

               With oWord.Selection.Find
                    .Text = "<<cDocAval" & i & ">>"
                    .Replacement.Text = IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
               End With
            'end peac

            End If

          RRelaCred.MoveNext
       Loop
    End If

    'PEAC 20070815
    RRelaCred.MoveFirst
    Do Until RRelaCred.EOF
        If RRelaCred!nConsValor = gColRelPersConyugue Then
           With oWord.Selection.Find
                .Text = "<<cNomCli2>>"
                .Replacement.Text = PstaNombre(RRelaCred!cPersNombre)
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
           End With

           With oWord.Selection.Find
                .Text = "<<cNuDoci2>>"
                .Replacement.Text = IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI)
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
           End With
        End If
      RRelaCred.MoveNext
    Loop
    'END PEAC 20070815

    With oWord.Selection.Find
         .Text = "<<cNomCli2>>"
         .Replacement.Text = ""
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
         .Text = "<<cNuDoci2>>"
         .Replacement.Text = ""
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With

    'PEAC 20070815
    With oWord.Selection.Find
         .Text = "<<cAval1>>"
         .Replacement.Text = ""
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
         .Text = "<<cDocAval1>>"
         .Replacement.Text = ""
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
         .Text = "<<cAval2>>"
         .Replacement.Text = ""
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
         .Text = "<<cDocAval2>>"
         .Replacement.Text = ""
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
    'END PEAC 20070815

    'WIOR 20130625 ********************************************
    Dim oEnvio As COMDCaptaGenerales.DCOMCaptaGenerales
    'Dim oDCred As COMDCredito.DCOMCredito
    Dim sTextoEnvio As String
    Dim rsEnvio As ADODB.Recordset
    
    Set oEnvio = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rsEnvio = oEnvio.RecuperaDatosEnvioEstadoCta(psCtaCod)
    Set oDCred = New COMDCredito.DCOMCredito

    If Not (rsEnvio.BOF And rsEnvio.EOF) Then
        If rsEnvio.RecordCount > 0 Then
            sTextoEnvio = oDCred.MostarTextoComisionEnvioEstCredito(psCtaCod, CInt(rsEnvio!nModoEnvio))
        End If
    Else
        sTextoEnvio = ""
    End If
    
    With oWord.Selection.Find
        .Text = "<<EnvioEstado>>"
        .Replacement.Text = sTextoEnvio
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'WIOR FIN *************************************************
'ARCV 03-05-2007
'    oDoc.SaveAs (App.path & "\FormatoCarta\HR_" & psCtaCod & "_" & Replace(Left(Time, 5), ":", "") & ".doc")
    oDoc.Close
    Set oDoc = Nothing
'-------

    Set oWord = CreateObject("Word.Application")
    oWord.Visible = True
    'ARCV 05-05-2007
    'Set oDoc = oWord.Documents.Open(App.path & "\FormatoCarta\HR_" & psCtaCod & "_" & Replace(Left(Time, 5), ":", "") & ".doc")
    Set oDoc = oWord.Documents.Open(sArchivo)
    '--------
    Set oDoc = Nothing
    Set oWord = Nothing
    Set rsRelEcotaxi = Nothing
    Set oDCred = Nothing
    Exit Function

'PEAC 20090520
ErrorCartillaCred:
    MsgBox err.Description, vbCritical, "Aviso"
    Set oDoc = Nothing
    Set oWord = Nothing

End Function
'*** BRGO 20111125 ******************************************************
Public Function ImprimeCartasEcotaxi(ByVal psCtaCod As String) As String

    Dim R As ADODB.Recordset
    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range
    Dim oCred As COMDCredito.DCOMCredDoc

    Dim sArchivo As String

    On Error GoTo ErrorCartasAutorInfoGas

    Set oCred = New COMDCredito.DCOMCredDoc
    Set R = oCred.ObtieneDatosInfoGasParaCartas(psCtaCod)
    Set oCred = Nothing

    If R.EOF And R.BOF Then
        MsgBox "No existen Datos para el Reporte...", vbInformation, "Aviso"
        Exit Function
    End If

    Set oWord = CreateObject("Word.Application")
        oWord.Visible = False
        Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\CartasAutorizInfogas.doc")
    
    sArchivo = App.Path & "\FormatoCarta\CA_" & psCtaCod & "_" & Replace(Left(Time, 5), ":", "") & ".doc"
    oDoc.SaveAs (sArchivo)

    'EJVG20130715***
    With oWord.Selection.Find
        .Text = "<<cCodCtaAbo>>"
        .Replacement.Text = R!cCtaCodRecaudo
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END EJVG ******
    
    With oWord.Selection.Find
        .Text = "<<cCodCta>>"
        .Replacement.Text = psCtaCod
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<FechaHoja>>"
        .Replacement.Text = Format(gdFecSis, "dddd") & ", " & Day(gdFecSis) & " de " & Format(gdFecSis, "mmmm") & " de " & Year(gdFecSis)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
        .Text = "<<cNomCli>>"
        .Replacement.Text = R!nomcli
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cNuDoci>>"
        .Replacement.Text = Trim(R!DNICli)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cNomCony>>"
        If R!NomCony <> "" Then
            .Replacement.Text = R!NomCony
        Else
            .Text = "y su  CONYUGE  Sr. ( a)  con DNI N° ,"
            .Replacement.Text = ""
        End If
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cNuDocCony>>"
        If R!DNICony <> "" Then
            .Replacement.Text = R!DNICony
        Else
            .Replacement.Text = ""
        End If
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
        .Text = "<<cNomConcesionario>>"
        .Replacement.Text = R!NomCon
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
        .Text = "<<nMonConcesionario>>"
        .Replacement.Text = Format(R!MonCon, "#,##0.00")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
        .Text = "<<cNomOperador>>"
        .Replacement.Text = R!NomOpe
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<nMonOperador>>"
        .Replacement.Text = Format(R!MonOpe, "#,##0.00")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cNomNotario>>"
        .Replacement.Text = R!NomNota
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<nMonNotario>>"
        .Replacement.Text = Format(R!MonNota, "#,##0.00")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
        .Text = "<<cNomNotario>>"
        .Replacement.Text = R!NomNota
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cNomAsegurador>>"
        .Replacement.Text = R!NomAseg
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<nMonAsegurador>>"
        .Replacement.Text = Format(R!MonAsg, "#,##0.00")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<nMonAprobado>>"
        .Replacement.Text = Format(R!nMontoCol, "#,##0.00")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find 'Pendiente de definir
        .Text = "<<nCuoIni>>"
        .Replacement.Text = "" 'Format(0, "#,##0.00")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
        .Text = "<<nValComercial>>"
        .Replacement.Text = Format(R!ValComercial, "#,##0.00")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
        .Text = "<<nPlazo>>"
        .Replacement.Text = Format(R!nPlazo, "#,##0.00")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<nTEA>>"
        .Replacement.Text = Format(R!TEA, "#,##0.00")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
        .Text = "<<nTEM>>"
        .Replacement.Text = Format(R!TEM, "#,##0.00")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find 'Pendiente de definir
        .Text = "<<cClase>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find 'Pendiente de definir
        .Text = "<<cMarca>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find 'Pendiente de definir
        .Text = "<<cModelo>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find 'Pendiente de definir
        .Text = "<<cColor>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find 'Pendiente de definir
        .Text = "<<cSerie>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find 'Pendiente de definir
        .Text = "<<cMotor>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find 'Pendiente de definir
        .Text = "<<cAnioFab>>"
        .Replacement.Text = R!dAnioFab
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
 
    oDoc.Close
    Set oDoc = Nothing

    Set oWord = CreateObject("Word.Application")
    oWord.Visible = True
    Set oDoc = oWord.Documents.Open(sArchivo)
    Set oDoc = Nothing
    Set oWord = Nothing

    Exit Function

ErrorCartasAutorInfoGas:
    MsgBox err.Description, vbCritical, "Aviso"
    Set oDoc = Nothing
    Set oWord = Nothing
End Function
'*** END BRGO ******************************************************



'*** PEAC 20160703
'*** PEAC 20160703
Public Function ImprimeInformeCriteriosAceptacionRiesgoFormatoEval(ByVal psCtaCod As String, _
ByVal psNomAge As String, ByVal psCodUsu As String, ByVal pR As ADODB.Recordset, ByVal pRSbs As ADODB.Recordset, ByVal pRDatFin As ADODB.Recordset, Optional ByVal pnProtestosSinAclarar As Integer = -1, Optional ByVal pnCobranzaCoativaSunat As Integer = -1, Optional ByVal pnObligacionesCerradas As Integer = -1) As String
'EAAS 20181126 segun ERS-072-2018
    Dim nTasaAnual As Double
    Dim nTasaMoraA As Double

    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range

    Dim RAge As ADODB.Recordset
    Dim sCompaSeguro As String
    Dim sNumeroPoliza As String
    Dim nParametro As Double
    Dim RRelaCred  As ADODB.Recordset
    Dim RSbs As ADODB.Recordset
    Dim RDatFin As ADODB.Recordset
    Dim RCap As ADODB.Recordset

    Dim nCosRef As Double
    Dim nComDsg As Double
    Dim nCosCroPag As Double
    Dim nCosHisPag As Double
    Dim nCosComPag As Double
    Dim nMontoInt As Double
    Dim nPriPol As Double
    Dim nCantAnt As Double
    Dim nCenCrg As Double
    Dim nSgrDsg1 As Double
    Dim nSgrDsg2 As Double
    Dim nSegMor As Double
    Dim nMinCuo As Double
    Dim nMaxCuo As Double
    Dim lcNonCal As String
    Dim i As Integer, lcCal As String
    Dim sArchivo As String
    Dim cCumpleIncumCAR As String
    Dim FecIni As Date, FecFin As Date
    Dim Meses As Integer, Anhos As Integer, Dias As Integer
    Dim nIncumplimientos As Integer 'EAAS 20181126 SEGUN ERS-072-2018
    Dim cResultado As String 'EAAS 20181126 SEGUN ERS-072-2018
    nIncumplimientos = 0 'EAAS 20181126 SEGUN ERS-072-2018

    On Error GoTo ErrorCargaCAR

    If pRSbs.RecordCount > 0 Then
    lcCal = ""

    For i = 0 To 4
        If i = 0 Then
            lcNonCal = "NOR"
        ElseIf i = 1 Then
            lcNonCal = "POT"
        ElseIf i = 2 Then
            lcNonCal = "DEF"
        ElseIf i = 3 Then
            lcNonCal = "DUD"
        ElseIf i = 4 Then
            lcNonCal = "PER"
        End If
        If pRSbs.Fields.Item(i) > 0 Then
            lcCal = lcCal & lcNonCal & Str(pRSbs.Fields.Item(i)) & " "
        End If
    Next
    End If

Anhos = 0
Meses = 0
Dias = 0

'--- PEAC 20080530 calcula Experiencia en sector en años meses y dias
If pR.RecordCount > 0 Then
'    If Format(pR!dPersIniActi, "dd/mm/yyyy") > "01/01/1950" Then
'        FecIni = Format(pR!dPersIniActi, "dd/mm/yyyy")
'       FecFin = Format(Date, "dd/mm/yyyy")
'        '** años
'        Anhos = Year(Format(FecFin, "yyyy/mm/dd")) - Year(Format(FecIni, "yyyy/mm/dd"))
'        '** meses
'        If Month(FecFin) = Month(FecIni) Then
'           Meses = 0
'        ElseIf Month(FecFin) > Month(FecIni) Then
'            Meses = Month(FecFin) - Month(FecIni)
'        ElseIf Month(FecFin) < Month(FecIni) Then
'            Meses = 12 + (Month(FecFin) - Month(FecIni))
'            Anhos = Anhos - 1
'        End If
'        '** dias
'        If Day(FecFin) >= Day(FecIni) Then
'            Dias = Day(FecFin) - Day(FecIni)
'        ElseIf Day(FecFin) < Day(FecIni) Then
'            Meses = Meses - 1
'            If Meses < 0 Then
'                Anhos = Anhos - 1
'                Meses = Meses + 12
'            End If
'            If Month(FecIni) = 1 Or Month(FecIni) = 3 Or Month(FecIni) = 5 Or Month(FecIni) = 7 Or Month(FecIni) = 8 Or Month(FecIni) = 10 Or Month(FecIni) = 12 Then
'            Dias = Day(FecFin) - Day(FecIni) + 31
'            ElseIf Month(FecIni) = 2 Then
'                Dias = Day(FecFin) - Day(FecIni) + 28
'                If Year(FecIni) Mod 4 = 0 Then Dias = Dias + 1
'            Else
'                Dias = Day(FecFin) - Day(FecIni) + 30
'            End If
'        End If
'    End If

Anhos = CInt(Mid(pR!cExperSectorEco, 1, 2))
Meses = CInt(Mid(pR!cExperSectorEco, 4, 2))

End If
'--- fin calcula Experiencia en sector en años meses y dias

cCumpleIncumCAR = Trim(pR!cCumpleIncumCAR)


    Set oWord = CreateObject("Word.Application")
        oWord.Visible = False

    'Set oDoc = oWord.Documents.Open(App.path & "\FormatoCarta\CriteriosAceptacionRiesgo.doc")
    Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\CAR_CriteriosAceptacionRiesgo.doc")

    sArchivo = App.Path & "\FormatoCarta\CAR_" & Format(Date, "yyyymmdd") & "_" & Replace(Left(Time, 8), ":", "") & ".doc"
    oDoc.SaveAs (sArchivo)

'*** Datos de Cabecera

    With oWord.Selection.Find
        .Text = "<<cNomCliente>>"
        '.Replacement.Text = R!cTipoCredDescrip
        .Replacement.Text = IIf(Len(pR!nombre_deudor) = 0, ".....", pR!nombre_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cAna>>"
        '.Replacement.Text = R!cTipoCredDescrip
        .Replacement.Text = IIf(UCase(Len(pR!Analista)) = 0, ".....", UCase(pR!Analista))
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

'*** Datos del contenido
    With oWord.Selection.Find
        .Text = "<<nCapa>>"
        If pRDatFin!nRatCapaPago = 0 Then
            .Replacement.Text = Str(Format(0, "#0.00")) & " %"
        Else
            '.Replacement.Text = Str(Format((pRDatFin!nRatcapaPgoRDS / pRDatFin!Excedente) * 100, "#0.00")) & " %"
            .Replacement.Text = Str(Format((pRDatFin!nRatCapaPago * 100), "#0.00")) & " %"
        End If
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
        .Text = "<<nCapa01>>"
        .Replacement.Text = Str(IIf(pRDatFin!nCI = True, 1, 0))
        .Forward = True
        'INICIO EAAS20181114
        If (.Replacement.Text = 0) Then
        nIncumplimientos = nIncumplimientos + 1
        End If
        'FIN EAAS20181114
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
        .Text = "<<nCal_Inter>>"
        .Replacement.Text = Trim(pR!Calif_Interna)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<nCal_Inter01>>"
        '.Replacement.Text = IIf(Trim(pR!Calif_Interna) = "A" Or Trim(pR!Calif_Interna) = "B" Or Trim(pR!Calif_Interna) = "C", "1", "0")
        .Replacement.Text = Trim(pR!cCalif_Interna)
        .Forward = True
        'INICIO EAAS20181114
        If (.Replacement.Text = 0) Then
        nIncumplimientos = nIncumplimientos + 1
        End If
        'FIN EAAS20181114
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<nCalificaSbs>>"
        .Replacement.Text = IIf(Trim(lcCal) <> "", Trim(lcCal) & "%", "Sin Cal.")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
        .Text = "<<nCalificaSbs01>>"
        .Replacement.Text = IIf(Trim(lcCal) = "NOR 100" Or Trim(lcCal) = "", "1", "0")
        .Forward = True
        'INICIO EAAS20181114
        If (.Replacement.Text = 0) Then
        nIncumplimientos = nIncumplimientos + 1
        End If
        'FIN EAAS20181114
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
        .Text = "<<nEscalon>>"
        .Replacement.Text = Format(pR!nMontoEscal, "#0.00") & "%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cEscalon>>"
        '.Replacement.Text = IIf(pR!bMontoEscal = True, "1", "0")
        '.Replacement.Text = IIf(pR!nMontoEscal > 30, "0", "1")
        .Replacement.Text = IIf(pR!bMontoEscal = True, "0", "1")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
        .Text = "<<nEscalon2>>"
        .Replacement.Text = Format(pR!nCuotaEscal, "#0.00") & "%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cEscalon2>>"
        '.Replacement.Text = IIf(pR!bCuotaEscal = True, "1", "0")
        '.Replacement.Text = IIf(pR!nCuotaEscal < 30, "1", "0")
        .Replacement.Text = IIf(pR!bCuotaEscal = True, "0", "1")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
        .Text = "<<nExperiencia>>"
        .Replacement.Text = Str(Anhos) & " Años, " & Str(Meses) & " Meses"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<nExperiencia01>>"
        '.Replacement.Text = IIf((Anhos = 0 And Meses > 10) Or (Anhos > 0), "1", "0")
        .Replacement.Text = cCumpleIncumCAR '***PEAC 20161003
        .Forward = True
        'INICIO EAAS20181114
         If (.Replacement.Text = 0) Then
         nIncumplimientos = nIncumplimientos + 1
         End If
         'END EAAS20181114
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
        'INICIO EAAS20181114 SEGUN ERS-072-2018 solo debe mostrar un único valor (1: Cumple o 0: Incumple) de acuerdo a las condiciones ya establecidas.
        With oWord.Selection.Find
        .Text = "<<cEscalon3>>"
        .Replacement.Text = IIf(pR!bCuotaEscal = True And pR!bMontoEscal = True, "0", "1")
        .Forward = True
        If (.Replacement.Text = 0) Then
        nIncumplimientos = nIncumplimientos + 1
        End If
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
        With oWord.Selection.Find
        .Text = "<<cCentralRiesgo1>>"
        .Replacement.Text = IIf(pnProtestosSinAclarar = 1, "No Registra", "Si Registra")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
        With oWord.Selection.Find
        .Text = "<<cCentralRiesgo2>>"
        .Replacement.Text = IIf(pnObligacionesCerradas = 1, "No Registra", "Si Registra")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
        With oWord.Selection.Find
        .Text = "<<cCentralRiesgo3>>"
        .Replacement.Text = IIf(pnCobranzaCoativaSunat = 1, "No Registra", "Si Registra")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cCentralRiesgo4>>"
        .Replacement.Text = IIf(pnCobranzaCoativaSunat = 0 Or pnProtestosSinAclarar = 0 Or pnObligacionesCerradas = 0, "0", "1")
        .Forward = True
         If (.Replacement.Text = 0) Then
         nIncumplimientos = nIncumplimientos + 1
         End If
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    If (nIncumplimientos = 0) Then
    cResultado = "Riesgo Aceptable"
    ElseIf (nIncumplimientos = 1) Then
    cResultado = "Riesgo Moderado"
    Else
    cResultado = "Riesgo NO Aceptable"
    End If
    With oWord.Selection.Find
        .Text = "<<nIncumple>>"
        .Replacement.Text = nIncumplimientos
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
        With oWord.Selection.Find
        .Text = "<<cRespuesta>>"
        .Replacement.Text = cResultado
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END EAAS20181114

'*** fin

    oDoc.Close
    Set oDoc = Nothing
    
    oWord.Visible = True ' ya no lo hacemos visible, por las puras si va ser pdf
    
    'aqui convertimos nuestro word en pdf y lo mostramos
    
'    oWord.Documents.Open(sArchivo).ExportAsFixedFormat OutputFileName:= _
'        App.Path & "\spooler\CAR_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf", ExportFormat:= _
'        wdExportFormatPDF, OpenAfterExport:=True, OptimizeFor:= _
'        wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, To:=1, _
'        item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
'        CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
'        BitmapMissingFonts:=True, UseISO19005_1:=False

   ' oWord.Documents.Close 'cerramos el archivo word, para no tenerlo en memoria
    'Set oDoc = Nothing
    Set oDoc = oWord.Documents.Open(sArchivo)
    Set oWord = Nothing
  '  Kill sArchivo ' nos deshacemos de nuestro archivo word

    Exit Function
ErrorCargaCAR:
    MsgBox err.Description, vbCritical, "Aviso"
    oDoc.Close
    Set oDoc = Nothing
    
    oWord.Quit
    Set oWord = Nothing
End Function

'*** PEAC 20160812
Public Function GeneraHojaEvalReporte6(ByVal prsDatos As ADODB.Recordset, _
                                        Optional ByVal pRSActivos As ADODB.Recordset = Nothing, _
                                        Optional ByVal pRSPasivos As ADODB.Recordset = Nothing, _
                                        Optional ByVal pRSCoeFinan As ADODB.Recordset = Nothing, _
                                        Optional ByVal pRSFlujoCaja As ADODB.Recordset = Nothing, _
                                        Optional ByVal pRSRatios As ADODB.Recordset = Nothing, _
                                        Optional ByVal pRSGastosFam As ADODB.Recordset = Nothing, _
                                        Optional ByVal pRSOtrosIng As ADODB.Recordset = Nothing _
                                        ) As Boolean

    Dim oDoc  As cPDF
    Set oDoc = New cPDF

    Dim a As Integer
    Dim B As Integer
    Dim i As Integer
    Dim nFila As Integer
    
    Dim lcAgencia As String
    Dim lcAnalista As String
    Dim lcCodCliente As String
    Dim lcNomCliente As String
    Dim lcDNI As String
    Dim lcRUC As String
    Dim lnConsValor As Integer
    Dim lcCtaCod As String
    Dim lcMoneda As String
    Dim lnSumaTotPP As Currency
    Dim lnSumaTotPE As Currency
    Dim lnCapiTrabajo As Currency
    Dim lcTitRep As String
    Dim lnSuma As Currency
    
    a = 50
    B = 29

    GeneraHojaEvalReporte6 = False

    If (prsDatos.BOF Or prsDatos.EOF) Then Exit Function
    
    'Creación del Archivo
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Informe de Visita Nº " & prsDatos!cCtaCod
    oDoc.Title = "Informe de Visita Nº " & prsDatos!cCtaCod
    
    If Not oDoc.PDFCreate(App.Path & "\Spooler\HojaEvaluacion_" & Trim(prsDatos!cCtaCod) & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Function
    End If

    'Contenido
    oDoc.Fonts.Add "F1", "Courier New", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Courier New", TrueType, Bold, WinAnsiEncoding
    
    oDoc.LoadImageFromFile App.Path & "\logo_cmacmaynas.bmp", "Logo"

    'Tamaño de hoja A4
    oDoc.NewPage A4_Vertical

    lcAgencia = prsDatos!CAgencia
    lcAnalista = prsDatos!cAnalistanombre
    lcCodCliente = prsDatos!cperscod
    lcNomCliente = prsDatos!cPersNombre
    lcDNI = prsDatos!cPersDni
    lcRUC = prsDatos!cPersRuc
    lcCtaCod = prsDatos!cCtaCod
    '''lcMoneda = IIf(Mid(prsDatos!cCtaCod, 9, 1) = "1", "SOLES", "DOLARES") 'marg ers044-2016
    lcMoneda = IIf(Mid(prsDatos!cCtaCod, 9, 1) = "1", StrConv(gcPEN_PLURAL, vbUpperCase), "DOLARES") 'marg ers044-2016
    lcTitRep = "HOJA DE EVALUACION"

    '---------- cabecera
    oDoc.WImage 40, 60, 35, 105, "Logo"
    oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
    oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
    oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
    oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
    oDoc.WTextBox 90, 100, 10, 400, "HOJA DE EVALUACION", "F2", 10, hCenter
    oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
    oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
    oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
    oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", space(11), lcRUC)), "F1", 7.5, hLeft
    oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
    oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
    
    nFila = 120
    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "----------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
    '-------------- fin cabecera
    
    'Set prsDatos = Nothing Set prsDatos = Nothing 'CTI3 ERS03-2020-Error despues del pase
 
'--------------------------------------------------------------------------------------------------------------------------------
    
    oDoc.WTextBox nFila, 55, 1, 160, "ACTIVOS", "F2", 7.5, hjustify
    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
    oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
    oDoc.WTextBox nFila, 140, 1, 160, "P.P", "F2", 7.5, hRight
    oDoc.WTextBox nFila, 240, 1, 160, "P.E.", "F2", 7.5, hRight
    oDoc.WTextBox nFila, 340, 1, 160, "TOTAL", "F2", 7.5, hRight
    
    Do While Not pRSActivos.EOF
        If pRSActivos!PP + pRSActivos!PE <> 0 Then
            nFila = nFila + 10
            
        If nFila >= 800 Then
            oDoc.NewPage A4_Vertical
            '---------- cabecera
            oDoc.WImage 40, 60, 35, 105, "Logo"
            oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
            oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
            oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
            oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
            oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
            oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
            oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", space(11), lcRUC)), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
            
            nFila = 120
            nFila = nFila + 10
            oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
            nFila = nFila + 10
            '-------------- fin cabecera
            nFila = 130
        End If
            
            If Left(pRSActivos!Concepto, 5) = "TOTAL" Then
                oDoc.WTextBox nFila, 55, 15, 250, Left(pRSActivos!Concepto, 35), "F2", 7.5, hLeft
                oDoc.WTextBox nFila, 150, 15, 150, Format(pRSActivos!PP, "#,#0.00"), "F2", 7.5, hRight
                oDoc.WTextBox nFila, 250, 15, 150, Format(pRSActivos!PE, "#,#0.00"), "F2", 7.5, hRight
                oDoc.WTextBox nFila, 350, 15, 150, Format(pRSActivos!PP + pRSActivos!PE, "#,#0.00"), "F2", 7.5, hRight
            Else
                oDoc.WTextBox nFila, 55, 15, 250, Left(pRSActivos!Concepto, 35), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 150, 15, 150, Format(pRSActivos!PP, "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 250, 15, 150, Format(pRSActivos!PE, "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 350, 15, 150, Format(pRSActivos!PP + pRSActivos!PE, "#,#0.00"), "F1", 7.5, hRight
            End If
        End If
        pRSActivos.MoveNext
    Loop
        
    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
    
'--------------------------------------------------------------------------------------------------------------------------------
    
        If nFila >= 800 Then
            oDoc.NewPage A4_Vertical
            '---------- cabecera
            oDoc.WImage 40, 60, 35, 105, "Logo"
            oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
            oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
            oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
            oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
            oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
            oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
            oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", space(11), lcRUC)), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
            
            nFila = 120
            nFila = nFila + 10
            oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
            nFila = nFila + 10
            '-------------- fin cabecera
            nFila = 130
        End If
    
    oDoc.WTextBox nFila, 55, 1, 160, "PASIVOS", "F2", 7.5, hjustify
    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
    oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
    oDoc.WTextBox nFila, 140, 1, 160, "P.P", "F2", 7.5, hRight
    oDoc.WTextBox nFila, 240, 1, 160, "P.E.", "F2", 7.5, hRight
    oDoc.WTextBox nFila, 340, 1, 160, "TOTAL", "F2", 7.5, hRight
    
    Do While Not pRSPasivos.EOF
        If pRSPasivos!PP + pRSPasivos!PE <> 0 Then
            nFila = nFila + 10
            
        If nFila >= 800 Then
            oDoc.NewPage A4_Vertical
            '---------- cabecera
            oDoc.WImage 40, 60, 35, 105, "Logo"
            oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
            oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
            oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
            oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
            oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
            oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
            oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", space(11), lcRUC)), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
            
            nFila = 120
            nFila = nFila + 10
            oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
            nFila = nFila + 10
            '-------------- fin cabecera
            nFila = 130
        End If
            If Left(pRSPasivos!Concepto, 5) = "TOTAL" Then
                oDoc.WTextBox nFila, 55, 15, 250, Left(pRSPasivos!Concepto, 35), "F2", 7.5, hLeft
                oDoc.WTextBox nFila, 150, 15, 150, Format(pRSPasivos!PP, "#,#0.00"), "F2", 7.5, hRight
                oDoc.WTextBox nFila, 250, 15, 150, Format(pRSPasivos!PE, "#,#0.00"), "F2", 7.5, hRight
                oDoc.WTextBox nFila, 350, 15, 150, Format(pRSPasivos!PP + pRSPasivos!PE, "#,#0.00"), "F2", 7.5, hRight
            Else
                oDoc.WTextBox nFila, 55, 15, 250, Left(pRSPasivos!Concepto, 35), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 150, 15, 150, Format(pRSPasivos!PP, "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 250, 15, 150, Format(pRSPasivos!PE, "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 350, 15, 150, Format(pRSPasivos!PP + pRSPasivos!PE, "#,#0.00"), "F1", 7.5, hRight
            End If
        End If
        pRSPasivos.MoveNext
    Loop
    
    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
       
'--------------------------------------------------------------------------------------------------------------------------------
        If nFila >= 800 Then
            oDoc.NewPage A4_Vertical
            '---------- cabecera
            oDoc.WImage 40, 60, 35, 105, "Logo"
            oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
            oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
            oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
            oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
            oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
            oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
            oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", space(11), lcRUC)), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
            
            nFila = 120
            nFila = nFila + 10
            oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
            nFila = nFila + 10
            '-------------- fin cabecera
            nFila = 130
        End If

    oDoc.WTextBox nFila, 55, 1, 160, "FLUJO DE CAJA MENSUAL", "F2", 7.5, hjustify
    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "----------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
    oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
    oDoc.WTextBox nFila, 140, 1, 160, "MONTO", "F2", 7.5, hRight
    
    Do While Not pRSFlujoCaja.EOF
        nFila = nFila + 10
        
        If nFila >= 800 Then
            oDoc.NewPage A4_Vertical
            '---------- cabecera
            oDoc.WImage 40, 60, 35, 105, "Logo"
            oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
            oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
            oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
            oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
            oDoc.WTextBox 90, 100, 10, 400, "HOJA DE EVALUACION", "F2", 10, hCenter
            oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
            oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", space(11), lcRUC)), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
            
            nFila = 120
            nFila = nFila + 10
            oDoc.WTextBox nFila, 50, 1, 500, "----------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
            nFila = nFila + 10
            '-------------- fin cabecera

        End If
        
        oDoc.WTextBox nFila, 55, 15, 250, Left(pRSFlujoCaja!cConcepto, 35), "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 150, 15, 150, Format(pRSFlujoCaja!nMonto, "#,#0.00"), "F1", 7.5, hRight
        pRSFlujoCaja.MoveNext
    Loop

    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "----------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10

    '----------------------------------------------------------------------------------------------------------------

    'nFila = 140
    'oDoc.WTextBox nFila, 55 + 250, 1, 160, "GASTOS FAMILIARES", "F1", 7.5, hjustify
    '------------------------------------------------------------------------------------------------------
    lnCapiTrabajo = 0
    If Not (pRSCoeFinan.EOF And pRSCoeFinan.BOF) Then
        Do While Not pRSCoeFinan.EOF
            If pRSCoeFinan!nConsValor = 101 And pRSCoeFinan!nConsCod = 7029 Then
                lnCapiTrabajo = pRSCoeFinan!nMonto
            End If
            pRSCoeFinan.MoveNext
        Loop
    End If
    
        If nFila >= 800 Then
            oDoc.NewPage A4_Vertical
            '---------- cabecera
            oDoc.WImage 40, 60, 35, 105, "Logo"
            oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
            oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
            oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
            oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
            oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
            oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
            oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", space(11), lcRUC)), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
            
            nFila = 120
            nFila = nFila + 10
            oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
            nFila = nFila + 10
            '-------------- fin cabecera
            nFila = 130
        End If
    
    
    If Not (pRSGastosFam.EOF And pRSGastosFam.BOF) Then
        oDoc.WTextBox nFila, 55, 1, 160, "GASTOS FAMILIARES", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "----------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
        oDoc.WTextBox nFila, 140, 1, 160, "MONTO", "F2", 7.5, hRight
        
        lnSuma = 0
        
            pRSGastosFam.MoveFirst
            Do While Not pRSGastosFam.EOF
                nFila = nFila + 10
                
                If nFila >= 800 Then
                    oDoc.NewPage A4_Vertical
                    '---------- cabecera
                    oDoc.WImage 40, 60, 35, 105, "Logo"
                    oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
                    oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
                    oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
                    oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
                    oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
                    oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
                    oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
                    oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
                    oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", space(11), lcRUC)), "F1", 7.5, hLeft
                    oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
                    oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
                    
                    nFila = 120
                    nFila = nFila + 10
                    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
                    nFila = nFila + 10
                    '-------------- fin cabecera
                    nFila = 130
                End If
                
                oDoc.WTextBox nFila, 55, 15, 250, Left(pRSGastosFam!cConsDescripcion, 35), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 150, 15, 150, Format(pRSGastosFam!nMonto, "#,#0.00"), "F1", 7.5, hRight
                lnSuma = lnSuma + pRSGastosFam!nMonto
                pRSGastosFam.MoveNext
            Loop
        
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "----------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10

        oDoc.WTextBox nFila, 55, 15, 250, Left("Totales", 35), "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 150, 15, 150, Format(lnSuma, "#,#0.00"), "F2", 7.5, hRight

        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "----------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        
        
    End If
    '--------------------------------------------------------------------------------------------------------------------------------
    
        If nFila >= 800 Then
            oDoc.NewPage A4_Vertical
            '---------- cabecera
            oDoc.WImage 40, 60, 35, 105, "Logo"
            oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
            oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
            oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
            oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
            oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
            oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
            oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", space(11), lcRUC)), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
            
            nFila = 120
            nFila = nFila + 10
            oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
            nFila = nFila + 10
            '-------------- fin cabecera
            nFila = 130
        End If
    
    If Not (pRSOtrosIng.EOF And pRSOtrosIng.BOF) Then
        oDoc.WTextBox nFila, 55, 1, 160, "OTROS INGRESOS", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "----------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
        oDoc.WTextBox nFila, 140, 1, 160, "MONTO", "F2", 7.5, hRight
    
            lnSuma = 0
            pRSOtrosIng.MoveFirst
            Do While Not pRSOtrosIng.EOF
                nFila = nFila + 10
                
                If nFila >= 800 Then
                    oDoc.NewPage A4_Vertical
                    '---------- cabecera
                    oDoc.WImage 40, 60, 35, 105, "Logo"
                    oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
                    oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
                    oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
                    oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
                    oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
                    oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
                    oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
                    oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
                    oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", space(11), lcRUC)), "F1", 7.5, hLeft
                    oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
                    oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
                    
                    nFila = 120
                    nFila = nFila + 10
                    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
                    nFila = nFila + 10
                    '-------------- fin cabecera
                    nFila = 130
                End If
                
                oDoc.WTextBox nFila, 55, 15, 250, Left(pRSOtrosIng!cConsDescripcion, 35), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 150, 15, 150, Format(pRSOtrosIng!nMonto, "#,#0.00"), "F1", 7.5, hRight
                lnSuma = lnSuma + pRSOtrosIng!nMonto
                pRSOtrosIng.MoveNext
            Loop
        
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "----------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        
        oDoc.WTextBox nFila, 55, 15, 250, Left("Totales", 35), "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 150, 15, 150, Format(lnSuma, "#,#0.00"), "F2", 7.5, hRight

        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "----------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        
        
    End If
        '----------------------------------------------------------------------------------------------------------------
        
        If nFila >= 800 Then
            oDoc.NewPage A4_Vertical
            '---------- cabecera
            oDoc.WImage 40, 60, 35, 105, "Logo"
            oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
            oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
            oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
            oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
            oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
            oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
            oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", space(11), lcRUC)), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
            
            nFila = 120
            nFila = nFila + 10
            oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
            nFila = nFila + 10
            '-------------- fin cabecera
            nFila = 130
        End If
        
        oDoc.WTextBox nFila, 55, 1, 160, "RATIOS E INDICADORES", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "----------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        
        oDoc.WTextBox nFila, 55, 1, 160, "Capacidad de pago", "F1", 7.5, hjustify
        If Not (Left(prsDatos!cTpoProdCod, 1) = "7" Or Left(prsDatos!cTpoProdCod, 1) = "8") Then
            oDoc.WTextBox nFila + 10, 55, 1, 160, "Liquidez Cte.", "F1", 7.5, hjustify
            oDoc.WTextBox nFila + 20, 55, 1, 160, "Capital de Trabajo", "F1", 7.5, hjustify
            oDoc.WTextBox nFila + 30, 55, 1, 160, "Endeudamiento Pat.", "F1", 7.5, hjustify
            oDoc.WTextBox nFila + 40, 55, 1, 160, "Rentabilidad Pat.", "F1", 7.5, hjustify
            oDoc.WTextBox nFila + 50, 55, 1, 160, "Excedente", "F1", 7.5, hjustify
            oDoc.WTextBox nFila, 150, 15, 150, Format(pRSRatios!nCapPagNeta * 100, "#,#0.00") & "%", "F1", 7.5, hRight
            oDoc.WTextBox nFila + 10, 150, 15, 150, Format(pRSRatios!nLiquidezCte, "#,#0.00"), "F1", 7.5, hRight
            oDoc.WTextBox nFila + 30, 150, 15, 150, Format(pRSRatios!nEndeuPat * 100, "#,#0.00") & "%", "F1", 7.5, hRight
            oDoc.WTextBox nFila + 40, 150, 15, 150, Format(pRSRatios!nRentaPatri * 100, "#,#0.00") & "%", "F1", 7.5, hRight
            oDoc.WTextBox nFila + 50, 150, 15, 150, Format(pRSRatios!nExceMensual, "#,#0.00"), "F1", 7.5, hRight
        Else
            oDoc.WTextBox nFila + 10, 55, 1, 160, "Capital de Trabajo", "F1", 7.5, hjustify
            oDoc.WTextBox nFila + 20, 55, 1, 160, "Excedente", "F1", 7.5, hjustify
            oDoc.WTextBox nFila, 150, 15, 150, Format(pRSRatios!nCapPagNeta * 100, "#,#0.00") & "%", "F1", 7.5, hRight
            oDoc.WTextBox nFila + 10, 150, 15, 150, Format(lnCapiTrabajo, "#,#0.00"), "F1", 7.5, hRight
            oDoc.WTextBox nFila + 20, 150, 15, 150, Format(pRSRatios!nExceMensual, "#,#0.00"), "F1", 7.5, hRight
        End If
'        If Not (Left(prsDatos!cTpoProdCod, 1) = "7" Or Left(prsDatos!cTpoProdCod, 1) = "8") Then
'            oDoc.WTextBox nFila + 10, 55, 1, 160, "Liquidez Cte.", "F1", 7.5, hjustify
'        End If
'        oDoc.WTextBox nFila + 20, 55, 1, 160, "Capital de Trabajo", "F1", 7.5, hjustify
'        If Not (Left(prsDatos!cTpoProdCod, 1) = "7" Or Left(prsDatos!cTpoProdCod, 1) = "8") Then
'            oDoc.WTextBox nFila + 30, 55, 1, 160, "Endeudamiento Pat.", "F1", 7.5, hjustify
'            oDoc.WTextBox nFila + 40, 55, 1, 160, "Rentabilidad Pat.", "F1", 7.5, hjustify
'        End If
'        oDoc.WTextBox nFila + 50, 55, 1, 160, "Excedente", "F1", 7.5, hjustify
'
'        oDoc.WTextBox nFila, 150, 15, 150, Format(pRSRatios!nCapPagNeta * 100, "#,#0.00") & "%", "F1", 7.5, hRight
'        If Not (Left(prsDatos!cTpoProdCod, 1) = "7" Or Left(prsDatos!cTpoProdCod, 1) = "8") Then
'            oDoc.WTextBox nFila + 10, 150, 15, 150, Format(pRSRatios!nLiquidezCte, "#,#0.00"), "F1", 7.5, hRight
'        End If
'        oDoc.WTextBox nFila + 20, 150, 15, 150, Format(lnCapiTrabajo, "#,#0.00"), "F1", 7.5, hRight
'        If Not (Left(prsDatos!cTpoProdCod, 1) = "7" Or Left(prsDatos!cTpoProdCod, 1) = "8") Then
'            oDoc.WTextBox nFila + 30, 150, 15, 150, Format(pRSRatios!nEndeuPat * 100, "#,#0.00") & "%", "F1", 7.5, hRight
'            oDoc.WTextBox nFila + 40, 150, 15, 150, Format(pRSRatios!nRentaPatri * 100, "#,#0.00") & "%", "F1", 7.5, hRight
'        End If
        oDoc.WTextBox nFila + 50, 150, 15, 150, Format(pRSRatios!nExceMensual, "#,#0.00"), "F1", 7.5, hRight
            
        nFila = nFila + 60
        oDoc.WTextBox nFila, 50, 1, 500, "----------------------------------------------------", "F1", 7.5, hLeft
        
        '----------------------------------------------------------------------------------------------------------------
    Set prsDatos = Nothing 'Set prsDatos = Nothing 'CTI3 ERS03-2020-Error despues del pase
    oDoc.PDFClose
    oDoc.Show

    
    GeneraHojaEvalReporte6 = True

End Function

'*** PEAC 20160729
Public Function GeneraImpresionReporte6(ByVal prsDatos As ADODB.Recordset, _
                                        Optional ByVal pRSActivos As ADODB.Recordset = Nothing, _
                                        Optional ByVal pRSPasivos As ADODB.Recordset = Nothing, _
                                        Optional ByVal pRSDetalleAct As ADODB.Recordset = Nothing, _
                                        Optional ByVal pRSDetallePas As ADODB.Recordset = Nothing, _
                                        Optional ByVal pRSEstGanPer As ADODB.Recordset = Nothing, _
                                        Optional ByVal pRSCoeFinan As ADODB.Recordset = Nothing, _
                                        Optional ByVal pRSFlujoCaja As ADODB.Recordset = Nothing, _
                                        Optional ByVal pRSGastosFam As ADODB.Recordset = Nothing, _
                                        Optional ByVal pRSOtrosIng As ADODB.Recordset = Nothing, _
                                        Optional ByVal pRSDeclaraPDT As ADODB.Recordset = Nothing, _
                                        Optional ByVal pRSDatPDTDet As ADODB.Recordset = Nothing, _
                                        Optional ByVal pRSIfisFlujoCaja As ADODB.Recordset = Nothing, _
                                        Optional ByVal pRSIfisGastosFam As ADODB.Recordset = Nothing, _
                                        Optional ByVal pRSRatios As ADODB.Recordset = Nothing) As Boolean

    Dim oDoc  As cPDF
    Set oDoc = New cPDF

    Dim a As Integer
    Dim B As Integer
    Dim i As Integer
    Dim nFila As Integer
    
    Dim lcAgencia As String
    Dim lcAnalista As String
    Dim lcCodCliente As String
    Dim lcNomCliente As String
    Dim lcDNI As String
    Dim lcRUC As String
    Dim lnConsValor As Integer
    Dim lcCtaCod As String
    Dim lcMoneda As String
    Dim lnSumaTotPP As Currency
    Dim lnSumaTotPE As Currency
    Dim lcTitRep As String
    a = 50
    B = 29

    GeneraImpresionReporte6 = False

    If (prsDatos.BOF Or prsDatos.EOF) Then Exit Function
    
    'Creación del Archivo
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Informe de Visita Nº " & prsDatos!cCtaCod
    oDoc.Title = "Informe de Visita Nº " & prsDatos!cCtaCod
    
    If Not oDoc.PDFCreate(App.Path & "\Spooler\FormatoEvaluacion_" & Trim(prsDatos!cCtaCod) & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Function
    End If

    'Contenido
    oDoc.Fonts.Add "F1", "Courier New", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Courier New", TrueType, Bold, WinAnsiEncoding
    
    oDoc.LoadImageFromFile App.Path & "\logo_cmacmaynas.bmp", "Logo"

    'Tamaño de hoja A4
    oDoc.NewPage A4_Vertical

    'Call CabeceraImpCuadros(rsInfVisita)

    lcAgencia = prsDatos!CAgencia
    lcAnalista = prsDatos!cAnalistanombre
    lcCodCliente = prsDatos!cperscod
    lcNomCliente = prsDatos!cPersNombre
    lcDNI = prsDatos!cPersDni
    lcRUC = prsDatos!cPersRuc
    lcCtaCod = prsDatos!cCtaCod
    '''lcMoneda = IIf(Mid(prsDatos!cCtaCod, 9, 1) = "1", "SOLES", "DOLARES")'marg ers044-2016
    lcMoneda = IIf(Mid(prsDatos!cCtaCod, 9, 1) = "1", StrConv(gcPEN_PLURAL, vbUpperCase), "DOLARES") 'marg ers044-2016
    lcTitRep = "ESTADOS FINANCIEROS"

    '---------- cabecera
    oDoc.WImage 40, 60, 35, 105, "Logo"
    oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
    oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
    oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
    oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
    oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
    oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
    oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
    oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
    oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", space(11), lcRUC)), "F1", 7.5, hLeft
    oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
    oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
    
    nFila = 120
    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
    '-------------- fin cabecera
    
    Set prsDatos = Nothing 'CTI3 ERS03-2020-Error despues del pase

'--------------------------------------------------------------------------------------------------------------------------------
    
'    oDoc.WTextBox nFila, 55, 1, 160, "ACTIVOS", "F1", 7.5, hjustify
'    nFila = nFila + 10
'    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
'    nFila = nFila + 10
'    oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F1", 7.5, hjustify
'    oDoc.WTextBox nFila, 140, 1, 160, "P.P", "F1", 7.5, hRight
'    oDoc.WTextBox nFila, 240, 1, 160, "P.E.", "F1", 7.5, hRight
'    oDoc.WTextBox nFila, 340, 1, 160, "TOTAL", "F1", 7.5, hRight
'
'    Do While Not pRSActivos.EOF
'        nFila = nFila + 10
'        oDoc.WTextBox nFila, 55, 15, 250, Left(pRSActivos!Concepto, 35), "F1", 7.5, hLeft
'        oDoc.WTextBox nFila, 150, 15, 150, Format(pRSActivos!PP, "#,#0.00"), "F1", 7.5, hRight
'        oDoc.WTextBox nFila, 250, 15, 150, Format(pRSActivos!PE, "#,#0.00"), "F1", 7.5, hRight
'        oDoc.WTextBox nFila, 350, 15, 150, Format(pRSActivos!PP + pRSActivos!PE, "#,#0.00"), "F1", 7.5, hRight
'        pRSActivos.MoveNext
'    Loop
'
'    nFila = nFila + 10
'    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
'    nFila = nFila + 10
    
'--------------------------------------------------------------------------------------------------------------------------------
    
'    oDoc.WTextBox nFila, 55, 1, 160, "PASIVOS", "F1", 7.5, hjustify
'    nFila = nFila + 10
'    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
'    nFila = nFila + 10
'    oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F1", 7.5, hjustify
'    oDoc.WTextBox nFila, 140, 1, 160, "P.P", "F1", 7.5, hRight
'    oDoc.WTextBox nFila, 240, 1, 160, "P.E.", "F1", 7.5, hRight
'    oDoc.WTextBox nFila, 340, 1, 160, "TOTAL", "F1", 7.5, hRight
'
'    Do While Not pRSPasivos.EOF
'        nFila = nFila + 10
'        oDoc.WTextBox nFila, 55, 15, 250, Left(pRSPasivos!Concepto, 35), "F1", 7.5, hLeft
'        oDoc.WTextBox nFila, 150, 15, 150, Format(pRSPasivos!PP, "#,#0.00"), "F1", 7.5, hRight
'        oDoc.WTextBox nFila, 250, 15, 150, Format(pRSPasivos!PE, "#,#0.00"), "F1", 7.5, hRight
'        oDoc.WTextBox nFila, 350, 15, 150, Format(pRSPasivos!PP + pRSPasivos!PE, "#,#0.00"), "F1", 7.5, hRight
'        pRSPasivos.MoveNext
'    Loop
'
'    nFila = nFila + 10
'    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
'    nFila = nFila + 10
       
'--------------------------------------------------------------------------------------------------------------------------------

'    oDoc.NewPage A4_Vertical
'
'    '---------- cabecera
'    oDoc.WImage 40, 60, 35, 105, "Logo"
'    oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
'    oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
'    oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
'    oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
'    oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
'    oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
'    oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
'    oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
'    oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", Space(11), lcRUC)), "F1", 7.5, hLeft
'    oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
'    oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
'
'    nFila = 120
'    nFila = nFila + 10
'    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
'    nFila = nFila + 10
'    '-------------- fin cabecera
              
'--------------------------------------------------------------------------------------------------------------------------------
    
    oDoc.WTextBox nFila, 55, 1, 160, "ESTADO DE GANANCIAS Y PERDIDAS", "F1", 7.5, hjustify
    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
    oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F1", 7.5, hjustify
    oDoc.WTextBox nFila, 140, 1, 160, "TOTAL", "F1", 7.5, hRight

    Do While Not pRSEstGanPer.EOF
        nFila = nFila + 10
        
        If nFila >= 800 Then
            oDoc.NewPage A4_Vertical
            '---------- cabecera
            oDoc.WImage 40, 60, 35, 105, "Logo"
            oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
            oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
            oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
            oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
            oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
            oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
            oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", space(11), lcRUC)), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
            
            nFila = 120
            nFila = nFila + 10
            oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
            nFila = nFila + 10
            '-------------- fin cabecera
            nFila = 130
        End If
        
        oDoc.WTextBox nFila, 55, 15, 250, Left(pRSEstGanPer!Concepto, 35), "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 150, 15, 150, Format(pRSEstGanPer!nMonto, "#,#0.00"), "F1", 7.5, hRight
        pRSEstGanPer.MoveNext
    Loop
    
    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10

'----------------------------------------------------------------------------------------------------------------

        If nFila >= 800 Then
            oDoc.NewPage A4_Vertical
            '---------- cabecera
            oDoc.WImage 40, 60, 35, 105, "Logo"
            oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
            oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
            oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
            oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
            oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
            oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
            oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", space(11), lcRUC)), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
            
            nFila = 120
            nFila = nFila + 10
            oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
            nFila = nFila + 10
            '-------------- fin cabecera
            nFila = 130
        End If

    If Not (pRSDetalleAct.EOF And pRSDetalleAct.BOF) Then
    oDoc.WTextBox nFila, 55, 1, 160, "DETALLE DE ACTIVOS", "F1", 7.5, hjustify
    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
    oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F1", 7.5, hjustify
    oDoc.WTextBox nFila, 140, 1, 160, "P.P", "F1", 7.5, hRight
    oDoc.WTextBox nFila, 240, 1, 160, "P.E.", "F1", 7.5, hRight
    oDoc.WTextBox nFila, 340, 1, 160, "TOTAL", "F1", 7.5, hRight
    
    i = 0
    'lnSumaTotPP = 0: lnSumaTotPE = 0
    
    Do While Not pRSDetalleAct.EOF
        nFila = nFila + 10
        
        If nFila >= 800 Then
            oDoc.NewPage A4_Vertical
            '---------- cabecera
            oDoc.WImage 40, 60, 35, 105, "Logo"
            oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
            oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
            oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
            oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
            oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
            oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
            oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", space(11), lcRUC)), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
            
            nFila = 120
            nFila = nFila + 10
            oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
            nFila = nFila + 10
            '-------------- fin cabecera
            nFila = 130
        End If
        
        If i = 0 Then
            lnConsValor = pRSDetalleAct!nConsValor
            oDoc.WTextBox nFila, 55, 15, 250, Left(Trim(pRSDetalleAct!cConcepto), 35), "F1", 7.5, hLeft
            nFila = nFila + 10
        Else
            If lnConsValor <> pRSDetalleAct!nConsValor Then
                i = 0
                lnConsValor = pRSDetalleAct!nConsValor
                oDoc.WTextBox nFila, 55, 15, 250, Left(Trim(pRSDetalleAct!cConcepto), 35), "F1", 7.5, hLeft
                nFila = nFila + 10
            End If
        End If
        i = i + 1
        'oDoc.WTextBox nFila, 55, 15, 250, pRSDetalleAct!dFechaDet, "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 55, 15, 250, pRSDetalleAct!cDescripcionDet, "F1", 7.5, hLeft
        'oDoc.WTextBox nFila, 120, 15, 250, pRSDetalleAct!cDescripcionDet, "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 150, 15, 150, Format(pRSDetalleAct!PP, "#,#0.00"), "F1", 7.5, hRight
        oDoc.WTextBox nFila, 250, 15, 150, Format(pRSDetalleAct!PE, "#,#0.00"), "F1", 7.5, hRight
        oDoc.WTextBox nFila, 350, 15, 150, Format(pRSDetalleAct!PP + pRSDetalleAct!PE, "#,#0.00"), "F1", 7.5, hRight
        
        'lnSumaTotPP = lnSumaTotPP + pRSDetalleAct!PP
        'lnSumaTotPE = lnSumaTotPE + pRSDetalleAct!PE
        
        pRSDetalleAct.MoveNext
    Loop
'    nFila = nFila + 10
'    oDoc.WTextBox nFila, 55, 15, 250, "Totales-->", "F1", 7.5, hLeft
'    oDoc.WTextBox nFila, 150, 15, 150, Format(lnSumaTotPP, "#,#0.00"), "F2", 7.5, hRight
'    oDoc.WTextBox nFila, 250, 15, 150, Format(lnSumaTotPE, "#,#0.00"), "F2", 7.5, hRight

    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
    End If
'----------------------------------------------------------------------------------------------------------------

        If nFila >= 800 Then
            oDoc.NewPage A4_Vertical
            '---------- cabecera
            oDoc.WImage 40, 60, 35, 105, "Logo"
            oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
            oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
            oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
            oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
            oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
            oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
            oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", space(11), lcRUC)), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
            
            nFila = 120
            nFila = nFila + 10
            oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
            nFila = nFila + 10
            '-------------- fin cabecera
            
        End If

    If Not (pRSDetallePas.EOF And pRSDetallePas.BOF) Then
    oDoc.WTextBox nFila, 55, 1, 160, "DETALLE DE PASIVOS", "F1", 7.5, hjustify
    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
    oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F1", 7.5, hjustify
    oDoc.WTextBox nFila, 140, 1, 160, "P.P", "F1", 7.5, hRight
    oDoc.WTextBox nFila, 240, 1, 160, "P.E.", "F1", 7.5, hRight
    oDoc.WTextBox nFila, 340, 1, 160, "TOTAL", "F1", 7.5, hRight
    
    i = 0
    Do While Not pRSDetallePas.EOF
        nFila = nFila + 10
        
        If nFila >= 800 Then
            oDoc.NewPage A4_Vertical
            '---------- cabecera
            oDoc.WImage 40, 60, 35, 105, "Logo"
            oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
            oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
            oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
            oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
            oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
            oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
            oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", space(11), lcRUC)), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
            
            nFila = 120
            nFila = nFila + 10
            oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
            nFila = nFila + 10
            '-------------- fin cabecera

        End If

        If i = 0 Then
            lnConsValor = pRSDetallePas!nConsValor
            oDoc.WTextBox nFila, 55, 15, 250, Left(Trim(pRSDetallePas!cConcepto), 35), "F1", 7.5, hLeft
            nFila = nFila + 10
        Else
            If lnConsValor <> pRSDetallePas!nConsValor Then
                i = 0
                lnConsValor = pRSDetallePas!nConsValor
                oDoc.WTextBox nFila, 55, 15, 250, Left(Trim(pRSDetallePas!cConcepto), 35), "F1", 7.5, hLeft
                nFila = nFila + 10
            End If
        End If

        i = i + 1
        'oDoc.WTextBox nFila, 55, 15, 250, pRSDetallePas!dFechaDet, "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 55, 15, 250, pRSDetallePas!cDescripcionDet, "F1", 7.5, hLeft
        'oDoc.WTextBox nFila, 120, 15, 250, pRSDetallePas!cDescripcionDet, "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 150, 15, 150, Format(pRSDetallePas!PP, "#,#0.00"), "F1", 7.5, hRight
        oDoc.WTextBox nFila, 250, 15, 150, Format(pRSDetallePas!PE, "#,#0.00"), "F1", 7.5, hRight
        oDoc.WTextBox nFila, 350, 15, 150, Format(pRSDetallePas!PP + pRSDetallePas!PE, "#,#0.00"), "F1", 7.5, hRight
        pRSDetallePas.MoveNext
    Loop

    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
    
    End If
'----------------------------------------------------------------------------------------------------------------

        If nFila >= 800 Then
            oDoc.NewPage A4_Vertical
            '---------- cabecera
            oDoc.WImage 40, 60, 35, 105, "Logo"
            oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
            oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
            oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
            oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
            oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
            oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
            oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", space(11), lcRUC)), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
            
            nFila = 120
            nFila = nFila + 10
            oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
            nFila = nFila + 10
            '-------------- fin cabecera

        End If

    oDoc.WTextBox nFila, 55, 1, 160, "COEFICIENTE FINANCIERO", "F1", 7.5, hjustify
    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
    oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F1", 7.5, hjustify
    oDoc.WTextBox nFila, 140, 1, 160, "TOTAL", "F1", 7.5, hRight

    Do While Not pRSCoeFinan.EOF
        nFila = nFila + 10
        
        If nFila >= 800 Then
            oDoc.NewPage A4_Vertical
            '---------- cabecera
            oDoc.WImage 40, 60, 35, 105, "Logo"
            oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
            oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
            oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
            oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
            oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
            oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
            oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", space(11), lcRUC)), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
            
            nFila = 120
            nFila = nFila + 10
            oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
            nFila = nFila + 10
            '-------------- fin cabecera

        End If
        
        oDoc.WTextBox nFila, 55, 15, 250, Left(pRSCoeFinan!Concepto, 35), "F1", 7.5, hLeft
        If pRSCoeFinan!nConsValor = 201 Or pRSCoeFinan!nConsValor = 303 Or pRSCoeFinan!nConsValor = 304 Then
            oDoc.WTextBox nFila, 150, 15, 150, Format(pRSCoeFinan!nMonto, "#,#0.00") & "%", "F1", 7.5, hRight
        Else
            oDoc.WTextBox nFila, 150, 15, 150, Format(pRSCoeFinan!nMonto, "#,#0.00"), "F1", 7.5, hRight
        End If
        pRSCoeFinan.MoveNext
    Loop

    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
'----------------------------------------------------------------------------------------------------------------

'    If nFila >= 800 Then
'        oDoc.NewPage A4_Vertical
'        '---------- cabecera
'        oDoc.WImage 40, 60, 35, 105, "Logo"
'        oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
'        oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
'        oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
'        oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
'        oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
'        oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
'        oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
'        oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
'        oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", Space(11), lcRUC)), "F1", 7.5, hLeft
'        oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
'        oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
'
'        nFila = 120
'        nFila = nFila + 10
'        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
'        nFila = nFila + 10
'        '-------------- fin cabecera
'
'    End If
'
'    oDoc.WTextBox nFila, 55, 1, 160, "FLUJO DE CAJA MENSUAL", "F1", 7.5, hjustify
'    nFila = nFila + 10
'    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
'    nFila = nFila + 10
'    oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F1", 7.5, hjustify
'    oDoc.WTextBox nFila, 140, 1, 160, "MONTO", "F1", 7.5, hRight
'
'    Do While Not pRSFlujoCaja.EOF
'        nFila = nFila + 10
'
'        If nFila >= 800 Then
'            oDoc.NewPage A4_Vertical
'            '---------- cabecera
'            oDoc.WImage 40, 60, 35, 105, "Logo"
'            oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
'            oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
'            oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
'            oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
'            oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
'            oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
'            oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
'            oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
'            oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", Space(11), lcRUC)), "F1", 7.5, hLeft
'            oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
'            oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
'
'            nFila = 120
'            nFila = nFila + 10
'            oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
'            nFila = nFila + 10
'            '-------------- fin cabecera
'
'        End If
'
'        oDoc.WTextBox nFila, 55, 15, 250, Left(pRSFlujoCaja!cConcepto, 35), "F1", 7.5, hLeft
'        oDoc.WTextBox nFila, 150, 15, 150, Format(pRSFlujoCaja!nMonto, "#,#0.00"), "F1", 7.5, hRight
'        pRSFlujoCaja.MoveNext
'    Loop
'
'    nFila = nFila + 10
'    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
'    nFila = nFila + 10

'----------------------------------------------------------------------------------------------------------------

'    If nFila >= 800 Then
'        oDoc.NewPage A4_Vertical
'        '---------- cabecera
'        oDoc.WImage 40, 60, 35, 105, "Logo"
'        oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
'        oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
'        oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
'        oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
'        oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
'        oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
'        oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
'        oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
'        oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", Space(11), lcRUC)), "F1", 7.5, hLeft
'        oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
'        oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
'
'        nFila = 120
'        nFila = nFila + 10
'        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
'        nFila = nFila + 10
'        '-------------- fin cabecera
'
'    End If
'
'    oDoc.WTextBox nFila, 55, 1, 160, "GASTOS FAMILIARES", "F1", 7.5, hjustify
'    nFila = nFila + 10
'    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
'    nFila = nFila + 10
'    oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F1", 7.5, hjustify
'    oDoc.WTextBox nFila, 140, 1, 160, "MONTO", "F1", 7.5, hRight
'
'    Do While Not pRSGastosFam.EOF
'        nFila = nFila + 10
'
'        If nFila >= 800 Then
'            oDoc.NewPage A4_Vertical
'            '---------- cabecera
'            oDoc.WImage 40, 60, 35, 105, "Logo"
'            oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
'            oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
'            oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
'            oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
'            oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
'            oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
'            oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
'            oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
'            oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", Space(11), lcRUC)), "F1", 7.5, hLeft
'            oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
'            oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
'
'            nFila = 120
'            nFila = nFila + 10
'            oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
'            nFila = nFila + 10
'            '-------------- fin cabecera
'
'        End If
'
'        oDoc.WTextBox nFila, 55, 15, 250, Left(pRSGastosFam!cConsDescripcion, 35), "F1", 7.5, hLeft
'        oDoc.WTextBox nFila, 150, 15, 150, Format(pRSGastosFam!nMonto, "#,#0.00"), "F1", 7.5, hRight
'        pRSGastosFam.MoveNext
'    Loop
'
'    nFila = nFila + 10
'    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
'    nFila = nFila + 10

    '----------------------------------------------------------------------------------------------------------------
    
'    If nFila >= 800 Then
'        oDoc.NewPage A4_Vertical
'        '---------- cabecera
'        oDoc.WImage 40, 60, 35, 105, "Logo"
'        oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
'        oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
'        oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
'        oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
'        oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
'        oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
'        oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
'        oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
'        oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", Space(11), lcRUC)), "F1", 7.5, hLeft
'        oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
'        oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
'
'        nFila = 120
'        nFila = nFila + 10
'        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
'        nFila = nFila + 10
'        '-------------- fin cabecera
'
'    End If
'
'    oDoc.WTextBox nFila, 55, 1, 160, "OTROS INGRESOS", "F1", 7.5, hjustify
'    nFila = nFila + 10
'    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
'    nFila = nFila + 10
'    oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F1", 7.5, hjustify
'    oDoc.WTextBox nFila, 140, 1, 160, "MONTO", "F1", 7.5, hRight
'
'    Do While Not pRSOtrosIng.EOF
'        nFila = nFila + 10
'
'        If nFila >= 800 Then
'            oDoc.NewPage A4_Vertical
'        '---------- cabecera
'        oDoc.WImage 40, 60, 35, 105, "Logo"
'        oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
'        oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
'        oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
'        oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
'        oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
'        oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
'        oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
'        oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
'        oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", Space(11), lcRUC)), "F1", 7.5, hLeft
'        oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
'        oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
'
'        nFila = 120
'        nFila = nFila + 10
'        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
'        nFila = nFila + 10
'        '-------------- fin cabecera
'
'        End If
'
'        oDoc.WTextBox nFila, 55, 15, 250, Left(pRSOtrosIng!cConsDescripcion, 35), "F1", 7.5, hLeft
'        oDoc.WTextBox nFila, 150, 15, 150, Format(pRSOtrosIng!nMonto, "#,#0.00"), "F1", 7.5, hRight
'        pRSOtrosIng.MoveNext
'    Loop
'
'    nFila = nFila + 10
'    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
'    nFila = nFila + 10

    '----------------------------------------------------------------------------------------------------------------

    If nFila >= 800 Then
        oDoc.NewPage A4_Vertical
        '---------- cabecera
        oDoc.WImage 40, 60, 35, 105, "Logo"
        oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
        oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
        oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
        oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
        oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
        oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
        oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
        oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
        oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", space(11), lcRUC)), "F1", 7.5, hLeft
        oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
        oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
        
        nFila = 120
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        '-------------- fin cabecera
    End If
    
    oDoc.WTextBox nFila, 55, 1, 160, "DECLARACION PDT", "F1", 7.5, hjustify
    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
    oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F1", 7.5, hjustify
    oDoc.WTextBox nFila, 80 - 80, 1, 160, DevolverNombreMes(pRSDeclaraPDT!nMes1), "F1", 7.5, hRight
    oDoc.WTextBox nFila, 160 - 80, 1, 160, DevolverNombreMes(pRSDeclaraPDT!nMes2), "F1", 7.5, hRight
    oDoc.WTextBox nFila, 240 - 80, 1, 160, DevolverNombreMes(pRSDeclaraPDT!nMes3), "F1", 7.5, hRight
    oDoc.WTextBox nFila, 320 - 80, 1, 160, "PROMEDIO", "F1", 7.5, hRight
    oDoc.WTextBox nFila, 400 - 80, 1, 160, "% VENT.DECLA.", "F1", 7.5, hRight

'1   Ventas
'2   Compras

    Do While Not pRSDatPDTDet.EOF

        If nFila >= 800 Then
            oDoc.NewPage A4_Vertical
            '---------- cabecera
            oDoc.WImage 40, 60, 35, 105, "Logo"
            oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
            oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
            oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
            oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
            oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
            oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
            oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
            oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", space(11), lcRUC)), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
            oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
            
            nFila = 120
            nFila = nFila + 10
            oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
            nFila = nFila + 10
            '-------------- fin cabecera
        End If

'    For i = 1 To feDeclaracionPDT.Rows - 1
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 15, 250, IIf(pRSDatPDTDet!nConsValor = 1, "Ventas", "Compras"), "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 90 - 80, 15, 150, Format(pRSDatPDTDet!nMontoMes1, "#,#0.00"), "F1", 7.5, hRight
        oDoc.WTextBox nFila, 170 - 80, 15, 150, Format(pRSDatPDTDet!nMontoMes2, "#,#0.00"), "F1", 7.5, hRight
        oDoc.WTextBox nFila, 250 - 80, 15, 150, Format(pRSDatPDTDet!nMontoMes3, "#,#0.00"), "F1", 7.5, hRight
        oDoc.WTextBox nFila, 330 - 80, 15, 150, Format(pRSDatPDTDet!nPromedio, "#,#0.00"), "F1", 7.5, hRight
        oDoc.WTextBox nFila, 410 - 80, 15, 150, Format(pRSDatPDTDet!nPorcentajeVent, "#,#0.00"), "F1", 7.5, hRight
        
'    Next i
        pRSDatPDTDet.MoveNext
    Loop
    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10

    '--------------------------------------------------------------------------------------------------------------------------------
    

    '-----------------------------------------------------------------------------------------
'    If nFila >= 800 Then
'        oDoc.NewPage A4_Vertical
'        '---------- cabecera
'        oDoc.WImage 40, 60, 35, 105, "Logo"
'        oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
'        oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
'        oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
'        oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
'        oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
'        oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
'        oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
'        oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
'        oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", Space(11), lcRUC)), "F1", 7.5, hLeft
'        oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
'        oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
'
'        nFila = 120
'        nFila = nFila + 10
'        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
'        nFila = nFila + 10
'        '-------------- fin cabecera
'
'    End If
    '-----------------------------------------------------------------------------------------

'    If Not (pRSRatios.EOF And pRSRatios.BOF) Then
'
'        If nFila >= 750 Then '800 normal , 750 para ratios porque son fijos
'            oDoc.NewPage A4_Vertical
'            '---------- cabecera
'            oDoc.WImage 40, 60, 35, 105, "Logo"
'            oDoc.WTextBox 40, 60, 35, 390, UCase(lcAgencia), "F1", 7.5, hLeft
'            oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
'            oDoc.WTextBox 70, 400, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
'            oDoc.WTextBox 80, 400, 10, 200, "ANALISTA: " & Trim(lcAnalista), "F1", 7.5, hLeft
'            oDoc.WTextBox 90, 100, 10, 400, lcTitRep, "F2", 10, hCenter
'            oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(lcCodCliente), "F1", 7.5, hLeft
'            oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(lcNomCliente), "F1", 7.5, hLeft
'            oDoc.WTextBox 100, 400, 10, 200, "DNI: " & Trim(lcDNI) & "   ", "F1", 7.5, hLeft
'            oDoc.WTextBox 110, 400, 10, 200, "RUC: " & Trim(IIf(lcRUC = "-", Space(11), lcRUC)), "F1", 7.5, hLeft
'            oDoc.WTextBox 120, 55, 10, 300, "CREDITO: " & Trim(lcCtaCod), "F1", 7.5, hLeft
'            oDoc.WTextBox 120, 400, 10, 200, "MONEDA: " & Trim(lcMoneda), "F1", 7.5, hLeft
'
'            nFila = 120
'            nFila = nFila + 10
'            oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
'            nFila = nFila + 10
'            '-------------- fin cabecera
'
'        End If
'
'        '----------------------------------------------------------------------------------------------------------------
'        oDoc.WTextBox nFila, 55, 1, 160, "RATIOS E INDICADORES", "F1", 7.5, hjustify
'        nFila = nFila + 10
'        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
'        nFila = nFila + 10
'
'        oDoc.WTextBox nFila, 55, 1, 160, "Capacidad de pago", "F1", 7.5, hjustify
'        oDoc.WTextBox nFila + 10, 55, 1, 160, "Ingreso Neto", "F1", 7.5, hjustify
'        oDoc.WTextBox nFila + 20, 55, 1, 160, "Excedente Mensual", "F1", 7.5, hjustify
'
''        oDoc.WTextBox nFila, 55, 1, 160, "Capacidad de pago", "F1", 7.5, hjustify
''        oDoc.WTextBox nFila + 10, 55, 1, 160, "Endeudamiento Patrimonial", "F1", 7.5, hjustify
''        oDoc.WTextBox nFila + 20, 55, 1, 160, "Liquidez Cte.", "F1", 7.5, hjustify
''        oDoc.WTextBox nFila + 30, 55, 1, 160, "Rentabilidad Pat.", "F1", 7.5, hjustify
''        oDoc.WTextBox nFila + 40, 55, 1, 160, "Ingreso Neto", "F1", 7.5, hjustify
''        oDoc.WTextBox nFila + 50, 55, 1, 160, "Excedente Mensual", "F1", 7.5, hjustify
'
'        oDoc.WTextBox nFila, 150, 15, 150, Format(pRSRatios!nCapPagNeta * 100, "#,#0.00") & "%", "F1", 7.5, hRight
'        oDoc.WTextBox nFila + 10, 150, 15, 150, Format(pRSRatios!nIngreNeto, "#,#0.00"), "F1", 7.5, hRight
'        oDoc.WTextBox nFila + 20, 150, 15, 150, Format(pRSRatios!nExceMensual, "#,#0.00"), "F1", 7.5, hRight
'
''        oDoc.WTextBox nFila, 150, 15, 150, Format(pRSRatios!nCapPagNeta * 100, "#,#0.00") & "%", "F1", 7.5, hRight
''        oDoc.WTextBox nFila + 10, 150, 15, 150, Format(pRSRatios!nEndeuPat * 100, "#,#0.00") & "%", "F1", 7.5, hRight
''        oDoc.WTextBox nFila + 20, 150, 15, 150, Format(pRSRatios!nLiquidezCte, "#,#0.00"), "F1", 7.5, hRight
''        oDoc.WTextBox nFila + 30, 150, 15, 150, Format(pRSRatios!nRentaPatri, "#,#0.00"), "F1", 7.5, hRight
''        oDoc.WTextBox nFila + 40, 150, 15, 150, Format(pRSRatios!nIngreNeto, "#,#0.00"), "F1", 7.5, hRight
''        oDoc.WTextBox nFila + 50, 150, 15, 150, Format(pRSRatios!nExceMensual, "#,#0.00"), "F1", 7.5, hRight
'
'        oDoc.WTextBox nFila, 320, 1, 250, "EN RELACION A SU EXCEDENTE", "F1", 7.5, hLeft
'        'oDoc.WTextBox nFila + 10, 320, 1, 250, "EN RELACION A SU PATRIMONIO TOTAL", "F1", 7.5, hLeft
'        '----------------------------------------------------------------------------------------------------------------
'
'    End If
   ' Set prsDatos = Nothing 'CTI3 ERS03-2020-Error despues del pase
    oDoc.PDFClose
    oDoc.Show
       
    GeneraImpresionReporte6 = True

End Function

'*** PEAC 20160809
Public Function CargaInformeVisitaPDF(ByVal pRs As ADODB.Recordset) As Boolean
    CargaInformeVisitaPDF = False
    Dim oDoc  As cPDF
    Set oDoc = New cPDF
    Dim a As Integer
    Dim B As Integer
    a = 50
    B = 29

    'Creación del Archivo
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Informe de visita Nº " & pRs!cCtaCod
    oDoc.Title = "Informe de visita Nº " & pRs!cCtaCod
    
    If Not oDoc.PDFCreate(App.Path & "\Spooler\" & pRs!cFormato & "_" & pRs!cCtaCod & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Function
    End If
    
    'Contenido
    oDoc.Fonts.Add "F1", "Arial", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Arial", TrueType, Bold, WinAnsiEncoding
    oDoc.LoadImageFromFile App.Path & "\logo_cmacmaynas.bmp", "Logo"
        
       
    'Tamaño de hoja A4
    oDoc.NewPage A4_Vertical
    '35
    oDoc.WImage 53, 43, 50, 115, "Logo"
    oDoc.WTextBox 45, 60, 35, 390, pRs!CAgencia, "F2", 10, hLeft
    
    oDoc.WTextBox 40, 60, 35, 390, "FECHA", "F2", 10, hRight
    oDoc.WTextBox 40, 60, 35, 450, Format(gdFecSis, "dd/mm/yyyy"), "F2", 10, hRight
    oDoc.WTextBox 40, 60, 35, 490, Format(Time, "hh:mm:ss"), "F2", 10, hRight
    
    B = 29
    oDoc.WTextBox 90 - B, 60, 15, 160, "Cliente", "F2", 10, hLeft
    oDoc.WTextBox 90 - B, 60, 15, 80, ":", "F2", 10, hRight
    oDoc.WTextBox 90 - B, 150, 15, 500, pRs!cPersNombre, "F1", 10, hjustify
    
    oDoc.WTextBox 71, 365, 35, 390, "Analista", "F2", 10, hLeft
    oDoc.WTextBox 71, 43, 35, 390, ":", "F2", 10, hRight
    oDoc.WTextBox 71, 440, 35, 390, UCase(pRs!cUserAnalista), "F1", 10, hjustify
    
    oDoc.WTextBox 100 - B, 60, 15, 160, "Usuario", "F2", 10, hLeft
    oDoc.WTextBox 100 - B, 60, 15, 80, ":", "F2", 10, hRight
    oDoc.WTextBox 100 - B, 150, 15, 118, gsCodUser, "F1", 10, hjustify
    
    oDoc.WTextBox 61, 365, 35, 390, "Producto", "F2", 10, hLeft
    oDoc.WTextBox 61, 43, 35, 390, ":", "F2", 10, hRight
    oDoc.WTextBox 61, 440, 35, 390, pRs!cConsDescripcion, "F1", 10, hjustify
    
    oDoc.WTextBox 110 - B, 60, 15, 160, "Crédito", "F2", 10, hLeft
    oDoc.WTextBox 110 - B, 60, 15, 80, ":", "F2", 10, hRight
    oDoc.WTextBox 110 - B, 150, 15, 500, pRs!cCtaCod, "F1", 10, hjustify
    
    oDoc.WTextBox 120 - B, 60, 15, 160, "Cod. Cliente", "F2", 10, hLeft
    oDoc.WTextBox 120 - B, 60, 15, 80, ":", "F2", 10, hRight
    oDoc.WTextBox 120 - B, 150, 15, 500, pRs!cperscod, "F1", 10, hjustify
    
    oDoc.WTextBox 81, 365, 35, 390, "Doc. Natural", "F2", 10, hLeft
    oDoc.WTextBox 81, 43, 35, 390, ":", "F2", 10, hRight
    oDoc.WTextBox 81, 440, 35, 390, pRs!cPersDni, "F1", 10, hjustify
    
    oDoc.WTextBox 91, 365, 35, 390, "Doc. Jurídico", "F2", 10, hLeft
    oDoc.WTextBox 91, 43, 35, 390, ":", "F2", 10, hRight
    oDoc.WTextBox 91, 440, 35, 390, pRs!cPersRuc, "F1", 10, hjustify
    
    a = 50
                'bajar izq  ar  der
    oDoc.WTextBox 110, 100, 15, 400, "INFORME DE VISITA AL CLIENTE", "F2", 12, hCenter
    
    'cuadro de Fecha de visita
    oDoc.WTextBox 130, 50, 80, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
    '135
    oDoc.WTextBox 185 - a, 55, 15, 160, "Fecha de Visita :", "F1", 10, hLeft
    oDoc.WTextBox 185 - a, 190, 15, 500, Format(pRs!dFecVisita, "dd/mm/yyyy"), "F1", 10, hjustify
    
    oDoc.WTextBox 185 - a, 300, 15, 160, "Fecha de última visita :", "F1", 10, hLeft
    oDoc.WTextBox 185 - a, 420, 15, 160, "__ / __ / ____", "F1", 10, hLeft
    
    oDoc.WTextBox 200 - a, 55, 15, 160, "Persona(s) Entrevistada(s) :", "F1", 10, hLeft
    
    oDoc.WTextBox 215 - a, 55, 15, 160, "Sr.(a) :", "F1", 10, hLeft
    oDoc.WTextBox 215 - a, 300, 15, 160, "Cargo/Parentesco :", "F1", 10, hjustify
    oDoc.WTextBox 230 - a, 55, 15, 160, "Sr.(a) :", "F1", 10, hLeft
    oDoc.WTextBox 230 - a, 300, 15, 160, "Cargo/Parentesco :", "F1", 10, hjustify
    
    'cuadro de Tipo de Visita
    oDoc.WTextBox 245 - a, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
    
    oDoc.WTextBox 247 - a, 55, 15, 500, "Tipo de Visita :", "F2", 10, hLeft
    '
    'cuadro de Tipo de Visita: Contenido
    oDoc.WTextBox 260 - a, 50, 23, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    
    oDoc.WTextBox 265 - a, 55, 15, 10, "( )", "F1", 10, hjustify
    oDoc.WTextBox 265 - a, 70, 15, 500, "1° Evaluación (Cliente Nuevo)", "F1", 10, hjustify
    oDoc.WTextBox 265 - a, 210, 15, 500, "( )", "F1", 10, hjustify
    oDoc.WTextBox 265 - a, 225, 15, 500, "Paralelo", "F1", 10, hjustify
    oDoc.WTextBox 265 - a, 275, 15, 700, "( )", "F1", 10, hjustify
    oDoc.WTextBox 265 - a, 290, 15, 800, "Inspección de Garantías", "F1", 10, hjustify
    oDoc.WTextBox 265 - a, 405, 15, 900, "( )", "F1", 10, hjustify
    oDoc.WTextBox 265 - a, 420, 15, 110, "Represtamo", "F1", 10, hjustify
    oDoc.WTextBox 265 - a, 480, 15, 120, "( )", "F1", 10, hjustify
    oDoc.WTextBox 265 - a, 495, 15, 130, "Ampliación", "F1", 10, hjustify
    a = 67
    'cuadro de Sobre el Entorno Familiar del Cliente o Representante
    oDoc.WTextBox 300 - a, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    
    oDoc.WTextBox 302 - a, 55, 15, 500, "Sobre el Entorno Familiar del Cliente o Representante:", "F2", 10, hLeft
    
    'cuadro de Sobre el Entorno Familiar del Cliente o Representante : CONTENIDO
    oDoc.WTextBox 315 - a, 50, 50, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    oDoc.WTextBox 320 - a, 55, 10, 500, pRs!cEntornoFami, "F1", 10, hjustify
    
    'cuadro de Sobre el giro y la Ubicacion del Negocio
    oDoc.WTextBox 365 - a, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    
    oDoc.WTextBox 367 - a, 55, 15, 500, "Sobre el Giro y la Ubicación del Negocio:", "F2", 10, hLeft
    
    'cuadro de Sobre el giro y la Ubicacion del Negocio : CONTENIDO
    oDoc.WTextBox 380 - a, 50, 50, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    oDoc.WTextBox 382 - a, 55, 10, 500, pRs!cGiroUbica, "F1", 10, hjustify
    
    'cuadro de Sobre la Experiencia Crediticia
    oDoc.WTextBox 430 - a, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    oDoc.WTextBox 432 - a, 55, 15, 500, "Sobre la Experiencia Crediticia:", "F2", 10, hLeft
    
    'cuadro de Sobre la Experiencia Crediticia : CONTENIDO
    oDoc.WTextBox 445 - a, 50, 50, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    oDoc.WTextBox 447 - a, 55, 10, 500, pRs!cExpeCrediticia, "F1", 10, hjustify
    
    'cuadro de Sobre la consistencia de la informacion y la formalidad del negocio
    oDoc.WTextBox 495 - a, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    oDoc.WTextBox 497 - a, 55, 15, 500, "Sobre la Consistencia de la Información y la Formalidad del Negocio:", "F2", 10, hLeft
    
    'cuadro de Sobre la consistencia de la informacion y la formalidad del negocio : CONTENIDO
    oDoc.WTextBox 510 - a, 50, 50, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    oDoc.WTextBox 512 - a, 55, 10, 500, pRs!cFormalNegocio, "F1", 10, hjustify
    
    'cuadro de Sobre la Colaterales o Garantias
    oDoc.WTextBox 560 - a, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    oDoc.WTextBox 562 - a, 55, 15, 500, "Sobre los Colaterales o Garantías:", "F2", 10, hLeft
    
    'cuadro de Sobre la Colaterales o Garantias : CONTENIDO
    oDoc.WTextBox 575 - a, 50, 50, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    oDoc.WTextBox 577 - a, 55, 10, 500, pRs!cColateGarantia, "F1", 10, hjustify
    
    'cuadro de Sobre el Destino y el Impacto del Mismo
    oDoc.WTextBox 625 - a, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    oDoc.WTextBox 627 - a, 55, 15, 500, "Sobre el Destino y el Impacto del Mismo:", "F2", 10, hLeft
    
    'cuadro de Sobre el Destino y el Impacto del Mismo : CONTENIDO
    oDoc.WTextBox 640 - a, 50, 50, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    oDoc.WTextBox 642 - a, 55, 10, 500, pRs!cDestino, "F1", 10, hjustify
    
    'cuadro de VERIFICACION DE INMUEBLE
    oDoc.WTextBox 690 - a, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    oDoc.WTextBox 692 - a, 55, 15, 500, "Verificacion de Inmueble :", "F2", 10, hLeft
    
    'cuadro de VERIFICACION DE INMUEBLE:CONTENIDO
    oDoc.WTextBox 705 - a, 50, 70, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    
    oDoc.WTextBox 707 - a, 55, 15, 500, "Dirección :", "F1", 10, hLeft
    oDoc.WTextBox 720 - a, 55, 15, 500, "Referencia de Ubicación :", "F1", 10, hLeft
    
    oDoc.WTextBox 732 - a, 55, 15, 500, "Zona :", "F1", 10, hLeft
    oDoc.WTextBox 732 - a, 200, 15, 500, "( )", "F1", 10, hjustify
    oDoc.WTextBox 732 - a, 220, 50, 500, "Urbana", "F1", 10, hjustify
    oDoc.WTextBox 732 - a, 300, 60, 500, "( )", "F1", 10, hjustify
    oDoc.WTextBox 732 - a, 320, 70, 500, "Rural", "F1", 10, hjustify
    
    oDoc.WTextBox 745 - a, 55, 15, 500, "Tipo de Construcción :", "F1", 10, hLeft
    oDoc.WTextBox 745 - a, 200, 15, 500, "( )", "F1", 10, hjustify
    oDoc.WTextBox 745 - a, 220, 15, 500, "Material Noble", "F1", 10, hjustify
    oDoc.WTextBox 745 - a, 300, 15, 500, "( )", "F1", 10, hjustify
    oDoc.WTextBox 745 - a, 320, 15, 500, "Madera", "F1", 10, hjustify
    oDoc.WTextBox 745 - a, 380, 15, 500, "( )", "F1", 10, hjustify
    oDoc.WTextBox 745 - a, 400, 15, 500, "Otros", "F1", 10, hjustify
    
    oDoc.WTextBox 757 - a, 55, 15, 500, "Estado de la Vivienda :", "F1", 10, hLeft
    
    'cuadro de comentario
    oDoc.WTextBox 775 - a, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
    
    oDoc.WTextBox 775 - a, 55, 15, 500, "Comentario :", "F2", 10, hjustify
    
    'cuadro de comentario:Contenido
    oDoc.WTextBox 790 - a, 50, 40, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
    
    
    'cuadro de firma:Contenido
    oDoc.WTextBox 830 - a, 50, 50, 250, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 830 - a, 300, 50, 250, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
    
    oDoc.WTextBox 834 - a, 55, 15, 500, "Analista de Créditos :", "F2", 10, hjustify
    oDoc.WTextBox 834 - a, 320, 15, 500, "Jefe de Grupo :", "F2", 10, hjustify
    
    oDoc.PDFClose
    oDoc.Show

CargaInformeVisitaPDF = True
End Function

'*** PEAC 20160809
Public Function CargaInformeVisitaPDF_form6(ByVal pRs As ADODB.Recordset) As Boolean
    CargaInformeVisitaPDF_form6 = False
    Dim oDoc  As cPDF
    Set oDoc = New cPDF
    Dim a As Integer
    Dim B As Integer
    Dim nTama As Double
    a = 50
    B = 29

    'Creación del Archivo
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Informe de visita Nº " & pRs!cCtaCod
    oDoc.Title = "Informe de visita Nº " & pRs!cCtaCod
    
    If Not oDoc.PDFCreate(App.Path & "\Spooler\" & pRs!cFormato & "_" & pRs!cCtaCod & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Function
    End If
    
    'Contenido
    oDoc.Fonts.Add "F1", "Arial", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Arial", TrueType, Bold, WinAnsiEncoding
    oDoc.LoadImageFromFile App.Path & "\logo_cmacmaynas.bmp", "Logo"
        
       
    'Tamaño de hoja A4
    oDoc.NewPage A4_Vertical
    '35
    oDoc.WImage 53, 43, 50, 115, "Logo"
    oDoc.WTextBox 45, 60, 35, 390, pRs!CAgencia, "F2", 10, hLeft
    
    oDoc.WTextBox 40, 60, 35, 390, "FECHA", "F2", 10, hRight
    oDoc.WTextBox 40, 60, 35, 450, Format(gdFecSis, "dd/mm/yyyy"), "F2", 10, hRight
    oDoc.WTextBox 40, 60, 35, 490, Format(Time, "hh:mm:ss"), "F2", 10, hRight
    
    B = 29
    oDoc.WTextBox 90 - B, 60, 15, 160, "Cliente", "F2", 10, hLeft
    oDoc.WTextBox 90 - B, 60, 15, 80, ":", "F2", 10, hRight
    oDoc.WTextBox 90 - B, 150, 15, 500, pRs!cPersNombre, "F1", 10, hjustify
    
    oDoc.WTextBox 71, 365, 35, 390, "Analista", "F2", 10, hLeft
    oDoc.WTextBox 71, 43, 35, 390, ":", "F2", 10, hRight
    oDoc.WTextBox 71, 440, 35, 390, UCase(pRs!cUserAnalista), "F1", 10, hjustify
    
    oDoc.WTextBox 100 - B, 60, 15, 160, "Usuario", "F2", 10, hLeft
    oDoc.WTextBox 100 - B, 60, 15, 80, ":", "F2", 10, hRight
    oDoc.WTextBox 100 - B, 150, 15, 118, gsCodUser, "F1", 10, hjustify
    
    oDoc.WTextBox 61, 365, 35, 390, "Producto", "F2", 10, hLeft
    oDoc.WTextBox 61, 43, 35, 390, ":", "F2", 10, hRight
    oDoc.WTextBox 61, 440, 35, 390, pRs!cConsDescripcion, "F1", 10, hjustify
    
    oDoc.WTextBox 110 - B, 60, 15, 160, "Credito", "F2", 10, hLeft
    oDoc.WTextBox 110 - B, 60, 15, 80, ":", "F2", 10, hRight
    oDoc.WTextBox 110 - B, 150, 15, 500, pRs!cCtaCod, "F1", 10, hjustify
    
    oDoc.WTextBox 120 - B, 60, 15, 160, "Cod. Cliente", "F2", 10, hLeft
    oDoc.WTextBox 120 - B, 60, 15, 80, ":", "F2", 10, hRight
    oDoc.WTextBox 120 - B, 150, 15, 500, pRs!cperscod, "F1", 10, hjustify
    
    oDoc.WTextBox 81, 365, 35, 390, "Doc. Natural", "F2", 10, hLeft
    oDoc.WTextBox 81, 43, 35, 390, ":", "F2", 10, hRight
    oDoc.WTextBox 81, 440, 35, 390, pRs!cPersDni, "F1", 10, hjustify
    
    oDoc.WTextBox 91, 365, 35, 390, "Doc. Juridico", "F2", 10, hLeft
    oDoc.WTextBox 91, 43, 35, 390, ":", "F2", 10, hRight
    oDoc.WTextBox 91, 440, 35, 390, pRs!cPersRuc, "F1", 10, hjustify
    
    a = 50
                'bajar izq  ar  der
    oDoc.WTextBox 110, 100, 15, 400, "INFORME DE VISITA AL CLIENTE", "F2", 12, hCenter
    
    'cuadro de Fecha de visita
    oDoc.WTextBox 130, 50, 80, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
    '135
    oDoc.WTextBox 185 - a, 55, 15, 160, "Fecha de Visita :", "F1", 10, hLeft
    oDoc.WTextBox 185 - a, 190, 15, 500, Format(pRs!dFecVisita, "dd/mm/yyyy"), "F1", 10, hjustify
    
    oDoc.WTextBox 185 - a, 300, 15, 160, "Fecha de ultima visita :", "F1", 10, hLeft
    oDoc.WTextBox 185 - a, 420, 15, 160, "__ / __ / ____", "F1", 10, hLeft
    
    oDoc.WTextBox 200 - a, 55, 15, 160, "Persona(s) Entrevistada(s) :", "F1", 10, hLeft
    
    oDoc.WTextBox 215 - a, 55, 15, 160, "Sr.(a) :", "F1", 10, hLeft
    oDoc.WTextBox 215 - a, 300, 15, 160, "Cargo/Parentesco :", "F1", 10, hjustify
    oDoc.WTextBox 230 - a, 55, 15, 160, "Sr.(a) :", "F1", 10, hLeft
    oDoc.WTextBox 230 - a, 300, 15, 160, "Cargo/Parentesco :", "F1", 10, hjustify

    'cuadro de Tipo de Visita
    oDoc.WTextBox 245 - a, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack

    oDoc.WTextBox 247 - a, 55, 15, 500, "Tipo de Visita :", "F2", 10, hLeft
    '
    'cuadro de Tipo de Visita: Contenido
    oDoc.WTextBox 260 - a, 50, 40, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    
    oDoc.WTextBox 265 - a, 55, 15, 10, "( )", "F1", 10, hjustify
    oDoc.WTextBox 265 - a, 75, 15, 500, "1° Evaluacion (Cliente Nuevo)", "F1", 10, hjustify
    oDoc.WTextBox 265 - a, 250, 15, 500, "( )", "F1", 10, hjustify
    oDoc.WTextBox 265 - a, 270, 15, 500, "Paralelo", "F1", 10, hjustify
    oDoc.WTextBox 265 - a, 400, 15, 700, "( )", "F1", 10, hjustify
    oDoc.WTextBox 265 - a, 420, 15, 800, "Inspeccion de Garantias", "F1", 10, hjustify
    oDoc.WTextBox 280 - a, 55, 15, 900, "( )", "F1", 10, hjustify
    oDoc.WTextBox 280 - a, 75, 15, 110, "Represtamo", "F1", 10, hjustify
    oDoc.WTextBox 280 - a, 250, 15, 120, "( )", "F1", 10, hjustify
    oDoc.WTextBox 280 - a, 270, 15, 130, "Ampliacion", "F1", 10, hjustify
    
    'cuadro de Sobre el Entorno Familiar del Cliente o Representante
'    oDoc.WTextBox 300 - A, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    
    oDoc.WTextBox 302 - a, 55, 15, 500, "Sobre el Entorno Familiar del Cliente o Representante:", "F2", 10, hLeft
    
    'cuadro de Sobre el Entorno Familiar del Cliente o Representante : CONTENIDO
    oDoc.WTextBox 315 - a, 50, 50, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    oDoc.WTextBox 320 - a, 55, 10, 500, pRs!cEntornoFami, "F1", 10, hjustify
    
'    nTama = oDoc.GetTextWidth(pRs!cEntornoFami, "F1", 10)
    nTama = oDoc.GetCellHeight(pRs!cEntornoFami, "F1", 10, 500)
    
    'cuadro de Sobre el giro y la Ubicacion del Negocio
    oDoc.WTextBox 365 - a, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    
    oDoc.WTextBox 367 - a, 55, 15, 500, "Sobre el Giro y la Ubicacion del Negocio:", "F2", 10, hLeft
    
    'cuadro de Sobre el giro y la Ubicacion del Negocio : CONTENIDO
    oDoc.WTextBox 380 - a, 50, 50, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    oDoc.WTextBox 382 - a, 55, 10, 500, pRs!cGiroUbica, "F1", 10, hjustify
    
    'cuadro de Sobre la Experiencia Crediticia
    oDoc.WTextBox 430 - a, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    oDoc.WTextBox 432 - a, 55, 15, 500, "Sobre la Experiencia Crediticia:", "F2", 10, hLeft
    
    'cuadro de Sobre la Experiencia Crediticia : CONTENIDO
    oDoc.WTextBox 445 - a, 50, 50, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    oDoc.WTextBox 447 - a, 55, 10, 500, pRs!cExpeCrediticia, "F1", 10, hjustify
    
    'cuadro de Sobre la consistencia de la informacion y la formalidad del negocio
    oDoc.WTextBox 495 - a, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    oDoc.WTextBox 497 - a, 55, 15, 500, "Sobre la Consistencia de la Informacion y la Formalidad del Negocio:", "F2", 10, hLeft
    
    'cuadro de Sobre la consistencia de la informacion y la formalidad del negocio : CONTENIDO
    oDoc.WTextBox 510 - a, 50, 50, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    oDoc.WTextBox 512 - a, 55, 10, 500, pRs!cFormalNegocio, "F1", 10, hjustify
    
    'cuadro de Sobre la Colaterales o Garantias
    oDoc.WTextBox 560 - a, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    oDoc.WTextBox 562 - a, 55, 15, 500, "Sobre los Colaterales o Garantias:", "F2", 10, hLeft
    
    'cuadro de Sobre la Colaterales o Garantias : CONTENIDO
    oDoc.WTextBox 575 - a, 50, 50, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    oDoc.WTextBox 577 - a, 55, 10, 500, pRs!cColateGarantia, "F1", 10, hjustify
    
    'cuadro de Sobre el Destino y el Impacto del Mismo
    oDoc.WTextBox 625 - a, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    oDoc.WTextBox 627 - a, 55, 15, 500, "Sobre el Destino y el Impacto del Mismo:", "F2", 10, hLeft
    
    'cuadro de Sobre el Destino y el Impacto del Mismo : CONTENIDO
    oDoc.WTextBox 640 - a, 50, 50, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    oDoc.WTextBox 642 - a, 55, 10, 500, pRs!cDestino, "F1", 10, hjustify
    
    'cuadro de VERIFICACION DE INMUEBLE
    oDoc.WTextBox 690 - a, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    oDoc.WTextBox 692 - a, 55, 15, 500, "Verificacion de Inmueble :", "F2", 10, hLeft
    
    'cuadro de VERIFICACION DE INMUEBLE:CONTENIDO
    oDoc.WTextBox 705 - a, 50, 95, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
    
    oDoc.WTextBox 707 - a, 55, 15, 500, "Direccion :", "F1", 10, hLeft
    oDoc.WTextBox 720 - a, 55, 15, 500, "Referencia de Ubicacion :", "F1", 10, hLeft
    oDoc.WTextBox 732 - a, 55, 15, 500, "Zona :", "F1", 10, hLeft
    oDoc.WTextBox 740 - a, 200, 15, 500, "( )", "F1", 10, hjustify
    oDoc.WTextBox 740 - a, 220, 50, 500, "Urbana", "F1", 10, hjustify
    oDoc.WTextBox 740 - a, 280, 60, 500, "( )", "F1", 10, hjustify
    oDoc.WTextBox 740 - a, 300, 70, 500, "Rural", "F1", 10, hjustify
    oDoc.WTextBox 752 - a, 55, 15, 500, "Tipo de Construccion :", "F1", 10, hLeft
    oDoc.WTextBox 765 - a, 100, 15, 500, "( )", "F1", 10, hjustify
    oDoc.WTextBox 765 - a, 120, 15, 500, "Material Noble", "F1", 10, hjustify
    oDoc.WTextBox 765 - a, 200, 15, 500, "( )", "F1", 10, hjustify
    oDoc.WTextBox 765 - a, 220, 15, 500, "Madera", "F1", 10, hjustify
    oDoc.WTextBox 765 - a, 280, 15, 500, "( )", "F1", 10, hjustify
    oDoc.WTextBox 765 - a, 300, 15, 500, "Otros", "F1", 10, hjustify
    oDoc.WTextBox 780 - a, 55, 15, 500, "Estado de la Vivienda :", "F1", 10, hLeft
    
    'cuadro de VISTO BUENO
    oDoc.WTextBox 800 - a, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
    
    oDoc.WTextBox 800 - a, 55, 15, 500, "Comentario :", "F2", 10, hjustify
    
    'cuadro de VISTO BUENO:Contenido
    oDoc.WTextBox 765, 50, 40, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
    
    oDoc.WTextBox 805, 55, 15, 500, "Analista de Creditos :", "F2", 10, hjustify
    oDoc.WTextBox 805, 320, 15, 500, "Jefe de Grupo :", "F2", 10, hjustify
    
    oDoc.PDFClose
    oDoc.Show

CargaInformeVisitaPDF_form6 = True
End Function
'*** PEAC 20170621
Public Sub ImprimePagareCredPDF(ByVal psCtaCod As String, ByVal pnFormato As Integer, Optional ByVal psFecDes As String)

Dim oDoc  As cPDF
Dim sLugar As String
Dim sFecEmision As String
Dim sFecVenc As String
Dim sMoneda As String
Dim sImporte As String
Dim sImporteLetras As String
Dim RelaGar As COMDPersona.DCOMPersonas ' PTI1 20170315
Set RelaGar = New COMDPersona.DCOMPersonas  ' PTI1 20170315
'INICIO EAAS 20180201
Dim nNumRepresentantes As Integer
nNumRepresentantes = 0
Dim cNombreTitular As String
cNombreTitular = ""
Dim cDocTitular As String
cDocTitular = ""
Dim cDirTitular As String
cDirTitular = ""
'FIN EAAS 20180201
Dim nTasaMora As Double
Dim nTasaInteres As Double

Dim ssql As String
Dim sCadImp As String
Dim sNomAgencia As String
Dim sEmision As String
Dim lsCiudad As String

Dim nGaran As Integer
Dim nTitu As Integer
Dim nCode As Integer
Dim liPosicion As Integer

Dim nTasaCompAnual As Double

Dim R As ADODB.Recordset
Dim RRelaCred As ADODB.Recordset
Dim rsUbi As ADODB.Recordset
Dim RsGarantes As New ADODB.Recordset

Dim oDCred As COMDCredito.DCOMCredito
Dim oFun As COMFunciones.FCOMCadenas
Dim ObjCons As New COMDConstantes.DCOMAgencias
Dim ObjGarantes As New COMDCredito.DCOMCredActBD
    
sNomAgencia = ObjCons.NombreAgencia(Mid(psCtaCod, 4, 2))
Set ObjCons = Nothing
Set oDCred = New COMDCredito.DCOMCredito
Set RRelaCred = oDCred.RecuperaRelacPers(psCtaCod)

Set RsGarantes = oDCred.RecuperaGarantes(psCtaCod)
 Dim nCantGarant As Integer
Dim sPersCodR As String 'PRT120170222, Agregó
Dim RrelGar As ADODB.Recordset 'PRT120170222, Agregó
Set oDCred = New COMDCredito.DCOMCredito
Set R = oDCred.RecuperaDatosComunes(psCtaCod)

Set rsUbi = oDCred.RecuperaUbigeo(Mid(psCtaCod, 4, 2))
sEmision = rsUbi!cUbiGeoDescripcion
    
lsCiudad = Trim(sEmision)
liPosicion = InStr(lsCiudad, "(")
    
If liPosicion > 0 Then
    lsCiudad = Left(lsCiudad, liPosicion - 1)
End If
    
Set oDCred = Nothing
Set oFun = New COMFunciones.FCOMCadenas

Set oDoc = New cPDF

'Creación del Archivo
oDoc.Author = gsCodUser
oDoc.Creator = "SICMACT - Negocio"
oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
oDoc.Subject = "Pagaré de Crédito Nº " & psCtaCod
oDoc.Title = "Pagaré de Crédito Nº " & psCtaCod

If Not oDoc.PDFCreate(App.Path & "\Spooler\Pagare_" & psCtaCod & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
    Exit Sub
End If

sMoneda = IIf(Mid(psCtaCod, 9, 1) = 1, gcPEN_PLURAL, "DOLARES") 'MARG ERS044-2016
If Not (R.EOF And R.BOF) Then

    
    If psFecDes = Empty Then
        sFecEmision = Format(R!dVigencia, "DD/MM/YYYY")
    Else
        sFecEmision = Format(CDate(psFecDes), "DD/MM/YYYY")
    End If
    nTasaMora = R!nTEAMora 'JUEZ 20140630
    nTasaInteres = R!nTasaInteres
    nTasaCompAnual = Format(((1 + nTasaInteres / 100) ^ (360 / 30) - 1) * 100, "#.00")
    sImporteLetras = NumLet(IIf(IsNull(R!nMontoPagare), 0, R!nMontoPagare)) & IIf(Mid(psCtaCod, 9, 1) = "2", "", " y " & IIf(InStr(1, R!nMontoPagare, ".") = 0, "00", Left(Mid(IIf(IsNull(R!nMontoPagare), 0, R!nMontoPagare), InStr(1, IIf(IsNull(R!nMontoPagare), 0, R!nMontoPagare), ".") + 1, 2) & "00", 2)) & "/100") 'EJVG20130924
Else
    sFecVenc = ""
    sFecEmision = ""
    sImporte = ""
    nTasaMora = 0
    nTasaInteres = 0
    nTasaCompAnual = 0
    sImporteLetras = ""
End If

oDoc.Fonts.Add "F1", "arial narrow", TrueType, Normal, WinAnsiEncoding
oDoc.Fonts.Add "F2", "arial narrow", TrueType, Bold, WinAnsiEncoding

oDoc.LoadImageFromFile App.Path & "\Logo_2015.jpg", "Logo"

oDoc.NewPage A4_Vertical

oDoc.WImage 50, 494, 35, 73, "Logo"
oDoc.WTextBox 30, 40, 15, 500, "PAGARÉ", "F2", 12, hCenter
oDoc.WTextBox 60, 45, 15, 175, "LUGAR DE EMISIÓN", "F1", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
oDoc.WTextBox 60, 220, 15, 160, "FECHA DE EMISIÓN", "F1", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
oDoc.WTextBox 60, 380, 15, 187, "NÚMERO", "F1", 10, hCenter, vMiddle, vbBlack, 1, vbBlack

oDoc.WTextBox 75, 45, 15, 175, lsCiudad, "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
'oDoc.WTextBox 75, 220, 15, 160, sFecEmision, "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack 'comentado PTI1 29-03-2017
oDoc.WTextBox 75, 380, 15, 187, psCtaCod, "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack

oDoc.WTextBox 90, 45, 15, 175, "FECHA DE VENCIMIENTO", "F1", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
oDoc.WTextBox 90, 220, 15, 160, "MONEDA PAGARÉ", "F1", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
oDoc.WTextBox 90, 380, 15, 187, "IMPORTE PAGARÉ", "F1", 10, hCenter, vMiddle, vbBlack, 1, vbBlack

oDoc.WTextBox 105, 45, 15, 175, sFecVenc, "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
oDoc.WTextBox 105, 220, 15, 160, sMoneda, "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
oDoc.WTextBox 105, 380, 15, 187, sImporte, "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
'----------------PTI1----------------------------------------------
oDoc.WTextBox 130, 45, 20, 480, "Por este", "F1", 11, hjustify
oDoc.WTextBox 131, 80, 20, 480, "PAGARÉ", "F2", 10, hjustify
oDoc.WTextBox 130, 117, 20, 480, "prometo/prometemos  pagar solidariamente e incondicionalmente a la orden de la", "F1", 11, hjustify
oDoc.WTextBox 131, 447, 20, 480, "CAJA MUNICIPAL DE AHORRO", "F2", 10, hjustify
oDoc.WTextBox 141, 45, 30, 480, "Y CRÉDITO DE MAYNAS S.A.", "F2", 10, hjustify
oDoc.WTextBox 140, 158, 30, 480, ", con R.U.C N° 20103845328, en adelante ", "F1", 11, hjustify
oDoc.WTextBox 141, 330, 30, 480, "LA CAJA", "F2", 10, hjustify
oDoc.WTextBox 140, 363, 30, 480, ", en cualquiera de sus oficinas a nivel nacional, o a", "F1", 11, hjustify
oDoc.WTextBox 150, 45, 30, 480, "quien", "F1", 11, hjustify
oDoc.WTextBox 151, 68, 30, 480, "LA CAJA", "F2", 10, hjustify
oDoc.WTextBox 150, 104, 30, 480, "hubiera endosado el presente título valor", "F1", 11, hjustify
oDoc.WTextBox 150, 267, 30, 600, ", la suma de__________________________________________________", "F1", 11, hjustify
oDoc.WTextBox 160, 45, 30, 600, "_________________________________________________", "F1", 11, hjustify
oDoc.WTextBox 160, 290, 30, 480, ", importe de dinero que expresamente declaro/declaramos adeudar a", "F1", 11, hjustify
oDoc.WTextBox 171, 45, 30, 480, "LA CAJA", "F2", 10, hjustify
oDoc.WTextBox 170, 80, 30, 500, "y que me(nos) obligo/obligamos a pagar en la misma moneda antes expresada en la fecha de vencimiento consignada. ", "F1", 11, hjustify
'-------------------------------------------------------------------------------------------
oDoc.WTextBox 190, 45, 30, 540, "Queda " & String(0.5, vbTab) & "expresamente" & String(0.55, vbTab) & " estipulado" & String(0.55, vbTab) & " que" & String(0.55, vbTab) & " el" & String(0.55, vbTab) & " importe" & String(0.55, vbTab) & " de" & String(0.55, vbTab) & " este" & String(0.55, vbTab) & " Pagaré " & String(0.55, vbTab) & "devengará" & String(0.55, vbTab) & " desde" & String(0.55, vbTab) & " su" & String(0.55, vbTab) & " fecha" & String(0.55, vbTab) & " de" & String(0.55, vbTab) & " emisión " & String(0.55, vbTab) & "hasta" & String(0.55, vbTab) & " la" & String(0.55, vbTab) & " fecha" & String(0.55, vbTab) & " de " & String(0.55, vbTab) & "", "F1", 11, hjustify
oDoc.WTextBox 200, 45, 30, 540, "su" & String(0.55, vbTab) & " vencimiento" & String(0.55, vbTab) & " un" & String(0.55, vbTab) & " interés" & String(0.55, vbTab) & " compensatorio" & String(0.55, vbTab) & " a" & String(0.55, vbTab) & " una" & String(0.55, vbTab) & " tasa" & String(0.55, vbTab) & " efectiva" & String(0.55, vbTab) & " anual" & String(0.55, vbTab) & " del" & String(0.55, vbTab) & "", "F1", 11, hjustify
oDoc.WTextBox 200, 394, 30, 520, "y" & String(0.55, vbTab) & " a" & String(0.55, vbTab) & "  partir" & String(0.55, vbTab) & "  de" & String(0.55, vbTab) & " su" & String(0.55, vbTab) & "  vencimiento" & String(0.55, vbTab) & " se" & String(0.55, vbTab) & " cobrará", "F1", 11, hjustify
oDoc.WTextBox 201, 358, 30, 520, "" & nTasaCompAnual & "%", "F2", 10, hjustify
oDoc.WTextBox 210, 45, 30, 520, "adicionalmente" & String(0.54, vbTab) & "un" & String(0.54, vbTab) & "interés" & String(0.54, vbTab) & "moratorio" & String(0.54, vbTab) & "a" & String(0.54, vbTab) & " una" & String(0.54, vbTab) & " tasa" & String(0.54, vbTab) & " efectiva" & String(0.54, vbTab) & " anual" & String(0.54, vbTab) & " del" & String(0.54, vbTab) & "", "F1", 11, hjustify
'oDoc.WTextBox 211, 320, 30, 520, "" & nTasaMora & ".00%.", "F2", 10, hjustify RIRO COMENTADO 20210918
oDoc.WTextBox 211, 320, 30, 520, "" & Format(nTasaMora, "0.00") & "%.", "F2", 10, hjustify  'RIRO 20210918 ADD
oDoc.WTextBox 210, 358, 30, 400, " Ambas" & String(0.54, vbTab) & "tasas" & String(0.54, vbTab) & "de" & String(0.54, vbTab) & "interés" & String(0.54, vbTab) & " continuarán" & String(0.54, vbTab) & "devengándose", "F1", 11, hjustify
oDoc.WTextBox 220, 45, 30, 520, "por todo el tiempo que demore el pago de la presente obligación.", "F1", 11, hjustify
'-------------------------------------------------------------------------------------------
oDoc.WTextBox 240, 45, 30, 515, "Asimismo " & String(0.51, vbTab) & " autorizo(amos) " & String(0.51, vbTab) & " de " & String(0.51, vbTab) & " manera" & String(0.51, vbTab) & " expresa " & String(0.51, vbTab) & " el cobro" & String(0.51, vbTab) & " de penalidades, seguros, gastos " & String(0.51, vbTab) & " notariales, de " & String(0.51, vbTab) & " cobranza judicial y " & String(300, vbTab) & "", "F1", 11, hjustify
oDoc.WTextBox 250, 45, 30, 520, String(2, vbTab) & " extrajudicial, y en" & String(0.54, vbTab) & " general" & String(0.54, vbTab) & "los gastos" & String(0.54, vbTab) & "y comisiones que pudiéramos adeudar derivados del crédito representado en este", "F1", 11, hjustify
oDoc.WTextBox 250, 535, 30, 520, "Pagaré,", "F1", 11, hjustify
oDoc.WTextBox 260, 45, 30, 540, "y que se pudieran generar desde la fecha de emisión del presente Pagaré hasta la cancelación total de la presente obligación,", "F1", 11, hjustify
oDoc.WTextBox 260, 554, 30, 540, "sin", "F1", 11, hjustify
oDoc.WTextBox 270, 45, 30, 540, "que" & String(0.55, vbTab) & "sea necesario" & String(0.55, vbTab) & " requerimiento" & String(0.55, vbTab) & " alguno" & String(0.55, vbTab) & " de" & String(0.55, vbTab) & " pago " & String(0.55, vbTab) & "para", "F1", 11, hjustify
oDoc.WTextBox 270, 278, 30, 540, "constituirme/constituirnos" & String(1, vbTab) & " en" & String(2, vbTab) & " mora," & String(0.54, vbTab) & " pues" & String(2, vbTab) & " es" & String(0.54, vbTab) & " entendido" & String(0.54, vbTab) & " que" & String(0.54, vbTab) & " ésta se ", "F1", 11, hjustify
oDoc.WTextBox 280, 45, 30, 540, "producirá de modo automático por el solo hecho del vencimiento de éste Pagaré.", "F1", 11, hjustify
'-------------------------------------------------------------------------------------------
oDoc.WTextBox 300, 45, 30, 540, "Expresamente" & String(0.55, vbTab) & " acepto(amos) toda" & String(1, vbTab) & " variación" & String(1, vbTab) & " de" & String(1, vbTab) & " las " & String(0.5, vbTab) & "tasas" & String(0.5, vbTab) & " de interés, dentro de los límites legales autorizados, las mismas que se ", "F1", 11, hjustify
oDoc.WTextBox 310, 45, 30, 540, "aplicarán" & String(0.55, vbTab) & " luego" & String(0.55, vbTab) & " de" & String(0.55, vbTab) & " la" & String(0.55, vbTab) & " comunicación" & String(0.55, vbTab) & " efectuada" & String(0.55, vbTab) & " por" & String(0.55, vbTab) & " la ", "F1", 11, hjustify
oDoc.WTextBox 311, 274, 30, 480, "LA CAJA", "F2", 10, hjustify
oDoc.WTextBox 310, 308, 30, 480, ", conforme a ley. Se" & String(0.55, vbTab) & " deja constancia que el presente Pagaré " & """" & "no", "F1", 11, hjustify
oDoc.WTextBox 320, 45, 30, 540, "requiere" & String(0.6, vbTab) & " ser" & String(0.55, vbTab) & " protestado" & """" & " por" & String(1.4, vbTab) & " falta" & String(1.4, vbTab) & " de" & String(1.4, vbTab) & " pago, procediendo" & String(1.4, vbTab) & "su ejecución" & String(0.55, vbTab) & " por el solo mérito del vencimiento del plazo pactado, o de", "F1", 11, hjustify
oDoc.WTextBox 330, 45, 30, 520, "sus renovaciones o prórrogas de ser el caso.", "F1", 11, hjustify
'-------------------------------------------------------------------------------------------
oDoc.WTextBox 350, 45, 30, 540, "De acuerdo" & String(0.55, vbTab) & "a" & String(0.55, vbTab) & " lo dispuesto en el numeral 11) del artículo 132° de la Ley General del Sistema Financiero y del Sistema de", "F1", 11, hjustify
oDoc.WTextBox 350, 533, 30, 540, "Seguros ", "F1", 11, hjustify
oDoc.WTextBox 360, 45, 30, 520, "y Orgánica " & String(0.5, vbTab) & "de" & String(0.5, vbTab) & " la" & String(0.55, vbTab) & "Superintendencia" & String(0.55, vbTab) & "de" & String(0.55, vbTab) & " Banca y Seguros, autorizo(amos) a la", "F1", 11, hjustify
oDoc.WTextBox 361, 355, 30, 480, "LA CAJA", "F2", 10, hjustify
oDoc.WTextBox 360, 392, 30, 480, "para" & String(0.55, vbTab) & " que compense entre mis acreencias y ", "F1", 11, hjustify
oDoc.WTextBox 370, 45, 30, 540, "activos (cuentas, valores, depósitos en general, entre otros) que" & String(0.55, vbTab) & " mantenga en su poder, hasta por el importe" & String(0.55, vbTab) & " de éste pagaré más", "F1", 11, hjustify
oDoc.WTextBox 380, 45, 30, 540, "los intereses compensatorios, moratorios, gastos y cualquier otro concepto antes detallado en el presente título valor.", "F1", 11, hjustify
'-------------------------------------------------------------------------------------------
oDoc.WTextBox 400, 45, 30, 530, "De" & String(0.55, vbTab) & " conformidad" & String(0.55, vbTab) & " con" & String(0.55, vbTab) & " el" & String(0.55, vbTab) & " artículo" & String(0.55, vbTab) & " 1233°" & String(0.55, vbTab) & " del" & String(0.55, vbTab) & " Código" & String(0.55, vbTab) & " Civil, acepto(amos)" & String(0.55, vbTab) & " que" & String(0.55, vbTab) & " en" & String(0.55, vbTab) & " la" & String(0.55, vbTab) & " eventualidad" & String(0.55, vbTab) & " de" & String(0.55, vbTab) & " que" & String(0.55, vbTab) & " el" & String(0.55, vbTab) & " presente" & String(0.55, vbTab) & " título se ", "F1", 11, hjustify
oDoc.WTextBox 410, 562, 30, 530, "o", "F1", 11, hjustify
oDoc.WTextBox 420, 45, 30, 540, "destrucción" & String(0.55, vbTab) & " parcial, deterioro" & String(0.55, vbTab) & " total, extravío" & String(0.55, vbTab) & " y sustracción, se aplicará lo dispuesto en los artículos 101° al 107° de la Ley No.27287, en lo que resultase pertinente.", "F1", 11, hjustify
oDoc.WTextBox 410, 45, 30, 525, "perjudicara" & String(0.55, vbTab) & "por" & String(0.55, vbTab) & "cualquier" & String(0.55, vbTab) & "causa, tal" & String(1, vbTab) & "hecho" & String(1, vbTab) & "no extinguirá la obligación primitiva" & String(0.51, vbTab) & "u original. Asimismo, en caso de deterioro notable", "F1", 11, hjustify
'-------------------------------------------------------------------------------------------
oDoc.WTextBox 450, 45, 30, 540, "Me(nos)" & String(0.55, vbTab) & " someto(emos) expresamente" & String(0.55, vbTab) & " a" & String(0.55, vbTab) & " la" & String(0.55, vbTab) & " competencia" & String(0.55, vbTab) & " y" & String(0.55, vbTab) & " tribunales" & String(0.55, vbTab) & "de" & String(0.55, vbTab) & "esta ciudad, en" & String(0.55, vbTab) & " cuyo" & String(0.55, vbTab) & " efecto" & String(0.55, vbTab) & " renuncio/renunciamos" & String(0.55, vbTab) & "al ", "F1", 11, hjustify
oDoc.WTextBox 460, 45, 30, 540, "fuero de mi/nuestro domicilio. Señalo(amos) como domicilio aquel" & String(0.55, vbTab) & " indicado" & String(0.55, vbTab) & " en" & String(0.55, vbTab) & " este pagaré, a donde se efectuarán las diligencias", "F1", 11, hjustify
oDoc.WTextBox 470, 45, 30, 540, "notariales, judiciales y demás que fuesen necesarias para lo que", "F1", 11, hjustify
oDoc.WTextBox 471, 306, 30, 480, "LA CAJA", "F2", 10, hjustify
oDoc.WTextBox 470, 342, 30, 520, "considere pertinente. Cualquier cambio de domicilio que", "F1", 11, hjustify
oDoc.WTextBox 480, 45, 30, 540, "haga(mos), para su validez, lo haré(mos) mediante carta notarial y conforme a lo dispuesto en el artículo 40° del Código Civil.", "F1", 11, hjustify
'-------------------------------------------------------------------------------------------
oDoc.WTextBox 500, 45, 30, 530, "Declaro(amos)" & String(2, vbTab) & " estar" & String(2, vbTab) & " plenamente" & String(2, vbTab) & " facultado(s)" & String(2, vbTab) & " para" & String(2, vbTab) & " suscribir" & String(2, vbTab) & " y" & String(2, vbTab) & " emitir" & String(2, vbTab) & "  el" & String(1, vbTab) & " presente" & String(1, vbTab) & " Pagaré, asumiendo", "F1", 11, hjustify
oDoc.WTextBox 500, 492, 30, 480, "en" & String(1, vbTab) & " caso" & String(1, vbTab) & " contrario", "F1", 11, hjustify
oDoc.WTextBox 510, 45, 30, 540, "responsabilidad civil y/o penal a que hubiera lugar. Se deja constancia que la información proporcionada por el(los) emitente(s) en", "F1", 11, hjustify
oDoc.WTextBox 520, 45, 30, 540, "el presente documento, tiene" & String(0.4, vbTab) & " el" & String(0.54, vbTab) & " carácter de declaración jurada, de acuerdo con el artículo 179° de la Ley No. 26702 - Ley General del Sistema Financiero y del Sistema de Seguros y Orgánica de la Superintendencia de Banca y Seguros.", "F1", 11, hjustify
oDoc.WTextBox 550, 45, 30, 520, "Suscribimos el presente en señal de conformidad.", "F1", 11, hjustify
'-------------------------------------------------------------------------------------------
oDoc.WTextBox 570, 45, 15, 520, "Nombres y Apellidos/Razón Social:", "F2", 11
Dim h As Integer
h = 140

If Not (RRelaCred.EOF And RRelaCred.BOF) Then
    Do Until RRelaCred.EOF
        If RRelaCred!nConsValor = gColRelPersTitular And (RRelaCred!nPersPersoneria = 2 Or RRelaCred!nPersPersoneria = 3 Or RRelaCred!nPersPersoneria = 4 Or RRelaCred!nPersPersoneria = 5 Or RRelaCred!nPersPersoneria = 6 Or RRelaCred!nPersPersoneria = 7 Or RRelaCred!nPersPersoneria = 8 Or RRelaCred!nPersPersoneria = 9 Or RRelaCred!nPersPersoneria = 10) Then

            oDoc.WTextBox 460 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
            oDoc.WTextBox 470 + h, 45, 20, 250, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
            oDoc.WTextBox 495 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
            oDoc.WTextBox 513 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
            oDoc.WTextBox 435 + h, 45, 35, 215, QuitarCaracter(PstaNombre(RRelaCred!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
            oDoc.WTextBox 460 + h, 90, 15, 205, IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
            oDoc.WTextBox 513 + h, 95, 35, 205, RRelaCred!cPersDireccDomicilio, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4

            oDoc.WTextBox 575 + h, 45, 15, 255, "Firma:_________________________", "F1", 11
            oDoc.WTextBox 545 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
            'INICIO EAAS 20180530 MEJORA PAGARE
            cNombreTitular = QuitarCaracter(PstaNombre(RRelaCred!cPersNombre), "-*/\._")
            cDocTitular = IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI)
            cDirTitular = RRelaCred!cPersDireccDomicilio
            'FIN EAAS 20180530 MEJORA PAGARE
            nTitu = 1
            
            ElseIf RRelaCred!nConsValor = gColRelPersTitular Then
            oDoc.WTextBox 460 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3 '
            oDoc.WTextBox 475 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3 '
            oDoc.WTextBox 435 + h, 45, 35, 205, QuitarCaracter(PstaNombre(RRelaCred!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
            oDoc.WTextBox 460 + h, 90, 15, 205, IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
            oDoc.WTextBox 475 + h, 95, 35, 205, RRelaCred!cPersDireccDomicilio, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
            oDoc.WTextBox 570 + h, 45, 15, 255, "Firma:_________________________", "F1", 11
            oDoc.WTextBox 540 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
            nTitu = 1

            nCode = 1

            ElseIf (RRelaCred!nConsValor = gColRelPersConyugue Or RRelaCred!nConsValor = gColRelPersCodeudor) And (RRelaCred!nPersPersoneria = 2 Or RRelaCred!nPersPersoneria = 3 Or RRelaCred!nPersPersoneria = 4 Or RRelaCred!nPersPersoneria = 5 Or RRelaCred!nPersPersoneria = 6 Or RRelaCred!nPersPersoneria = 7 Or RRelaCred!nPersPersoneria = 8 Or RRelaCred!nPersPersoneria = 9 Or RRelaCred!nPersPersoneria = 10) Then
            sPersCodR = RRelaCred!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 475 + h, 330, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 485 + h, 375, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                   
                End If
            oDoc.WTextBox 570, 330, 15, 520, "Nombres y Apellidos/Razón Social:", "F2", 11
            oDoc.WTextBox 435 + h, 330, 35, 205, QuitarCaracter(PstaNombre(RRelaCred!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
            oDoc.WTextBox 460 + h, 375, 15, 205, IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                          
            oDoc.WTextBox 513 + h, 380, 35, 205, RRelaCred!cPersDireccDomicilio, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
            oDoc.WTextBox 460 + h, 330, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
            oDoc.WTextBox 470 + h, 330, 20, 250, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
            oDoc.WTextBox 495 + h, 330, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
            oDoc.WTextBox 513 + h, 330, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3

            oDoc.WTextBox 575 + h, 330, 15, 255, "Firma:_________________________", "F1", 11
            oDoc.WTextBox 545 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack

            nCode = 1
            
            ElseIf (RRelaCred!nConsValor = gColRelPersConyugue Or RRelaCred!nConsValor = gColRelPersCodeudor) Then
            sPersCodR = RRelaCred!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 475 + h, 330, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 485 + h, 375, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                   
                End If
            oDoc.WTextBox 570, 330, 15, 520, "Nombres y Apellidos/Razón Social:", "F2", 11
            oDoc.WTextBox 435 + h, 330, 35, 205, QuitarCaracter(PstaNombre(RRelaCred!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
            oDoc.WTextBox 460 + h, 375, 15, 205, IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
            oDoc.WTextBox 475 + h, 380, 35, 205, RRelaCred!cPersDireccDomicilio, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
            
            oDoc.WTextBox 460 + h, 330, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3

            oDoc.WTextBox 475 + h, 330, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3

            oDoc.WTextBox 570 + h, 330, 15, 255, "Firma:_________________________", "F1", 11
            oDoc.WTextBox 540 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack

            nCode = 1

            ElseIf RRelaCred!nConsValor = gColRelPersRepresentante Then
                Select Case nNumRepresentantes
                Case 0
                    oDoc.WTextBox 473 + h, 45, 35, 205, QuitarCaracter(PstaNombre(RRelaCred!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    oDoc.WTextBox 485 + h, 90, 35, 250, IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                Case 1
                    oDoc.WTextBox 435 + h, 330, 35, 215, cNombreTitular, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    oDoc.WTextBox 460 + h, 375, 15, 205, cDocTitular, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                                           
                    oDoc.WTextBox 570, 330, 15, 520, "Nombres y Apellidos/Razón Social:", "F2", 11
                    oDoc.WTextBox 473 + h, 330, 35, 205, QuitarCaracter(PstaNombre(RRelaCred!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    oDoc.WTextBox 495 + h, 375, 15, 205, IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                                  
                    oDoc.WTextBox 513 + h, 380, 35, 205, cDirTitular, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                    oDoc.WTextBox 460 + h, 330, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
                    oDoc.WTextBox 470 + h, 330, 20, 250, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    oDoc.WTextBox 495 + h, 330, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
                    oDoc.WTextBox 513 + h, 330, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
        
                    oDoc.WTextBox 575 + h, 330, 15, 255, "Firma:_________________________", "F1", 11
                    oDoc.WTextBox 545 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
                Case 2
                    oDoc.NewPage A4_Vertical
                    oDoc.WTextBox 75, 45, 35, 205, cNombreTitular, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    oDoc.WTextBox 100, 90, 15, 205, cDocTitular, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    oDoc.WTextBox 150, 95, 35, 205, cDirTitular, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                      
                    oDoc.WTextBox 70, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
                    oDoc.WTextBox 113, 45, 35, 205, QuitarCaracter(PstaNombre(RRelaCred!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    oDoc.WTextBox 125, 90, 35, 250, IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    
                    oDoc.WTextBox 100, 45, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
                    oDoc.WTextBox 110, 45, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    oDoc.WTextBox 135, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
                    oDoc.WTextBox 150, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    oDoc.WTextBox 210, 45, 15, 255, "Firma:___________________________", "F1", 11
                    oDoc.WTextBox 190, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
                Case 3
                    oDoc.WTextBox 75, 300, 35, 205, cNombreTitular, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    oDoc.WTextBox 100, 345, 15, 205, cDocTitular, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    oDoc.WTextBox 150, 350, 35, 205, cDirTitular, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                 
                    oDoc.WTextBox 70, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
                    oDoc.WTextBox 113, 300, 35, 205, QuitarCaracter(PstaNombre(RRelaCred!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    oDoc.WTextBox 135, 345, 15, 205, IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    
                    oDoc.WTextBox 100, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
                    oDoc.WTextBox 110, 300, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    oDoc.WTextBox 135, 300, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
                    oDoc.WTextBox 150, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    oDoc.WTextBox 210, 300, 15, 255, "Firma:___________________________", "F1", 11
                    oDoc.WTextBox 190, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
                Case 4
                    Dim e As Integer
                    e = 250
                    oDoc.WTextBox 75 + e, 45, 35, 205, cNombreTitular, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    oDoc.WTextBox 100 + e, 90, 15, 205, cDocTitular, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    oDoc.WTextBox 150 + e, 95, 35, 205, cDirTitular, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                      
                    oDoc.WTextBox 70 + e, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
                    oDoc.WTextBox 113 + e, 45, 35, 205, QuitarCaracter(PstaNombre(RRelaCred!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    oDoc.WTextBox 125 + e, 90, 35, 250, IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    
                    oDoc.WTextBox 100 + e, 45, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
                    oDoc.WTextBox 110 + e, 45, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    oDoc.WTextBox 135 + e, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
                    oDoc.WTextBox 150 + e, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    oDoc.WTextBox 210 + e, 45, 15, 255, "Firma:___________________________", "F1", 11
                    oDoc.WTextBox 190 + e, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
                Case 5
                    oDoc.WTextBox 75 + e, 300, 35, 205, cNombreTitular, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    oDoc.WTextBox 100 + e, 345, 15, 205, cDocTitular, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    oDoc.WTextBox 150 + e, 350, 35, 205, cDirTitular, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                 
                    oDoc.WTextBox 70 + e, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
                    oDoc.WTextBox 113 + e, 300, 35, 205, QuitarCaracter(PstaNombre(RRelaCred!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    oDoc.WTextBox 135 + e, 345, 15, 205, IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    
                    oDoc.WTextBox 100 + e, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
                    oDoc.WTextBox 110 + e, 300, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    oDoc.WTextBox 135 + e, 300, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
                    oDoc.WTextBox 150 + e, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                    oDoc.WTextBox 210 + e, 300, 15, 255, "Firma:___________________________", "F1", 11
                    oDoc.WTextBox 190 + e, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
                End Select
                nNumRepresentantes = nNumRepresentantes + 1
                
                nCode = 1
            'Exit Do
        End If
        RRelaCred.MoveNext
    Loop
End If

ImprimeFianza psCtaCod, lsCiudad, oDoc, RsGarantes, RelaGar

oDoc.PDFClose
oDoc.Show
End Sub

'*** PEAC - 20170605
Private Function ImprimeFianza(ByVal psCtaCod As String, ByVal lsCiudad As String, ByVal oDoc, ByVal RsGarantes, ByVal RelaGar)

Dim sPersCodR As String
Dim RrelGar As ADODB.Recordset
Dim h As Integer
h = 140
oDoc.NewPage A4_Vertical
oDoc.WImage 50, 494, 35, 73, "Logo"
oDoc.WTextBox 50, 45, 15, 500, "FIANZA SOLIDARIA", "F2", 11
'---------------------------------------------------------------------------------------
oDoc.WTextBox 69, 45, 15, 545, "Me/Nos constituyo/constituimos en fiador/es solidario/s del(os) emitente(s) de este Pagaré, en forma irrevocable, incondicionada, ilimitada e indefinida, a favor de la", "F1", 11, hjustify
oDoc.WTextBox 81, 185, 15, 520, "CAJA" & String(0.55, vbTab) & " MUNICIPAL" & String(0.55, vbTab) & " DE" & String(0.55, vbTab) & " AHORRO" & String(0.55, vbTab) & " Y CRÉDITO DE MAYNAS S.A.", "F2", 10, hjustify
oDoc.WTextBox 80, 430, 20, 480, ", con R.U.C. N° 20103845328, en", "F1", 11, hjustify
oDoc.WTextBox 90, 45, 20, 550, "adelante" & String(15, vbTab) & ", renunciando" & String(2, vbTab) & " expresamente" & String(2, vbTab) & " al" & String(2, vbTab) & " beneficio" & String(2, vbTab) & " de" & String(2, vbTab) & " excusión" & String(1, vbTab) & " por" & String(1, vbTab) & " obligaciones" & String(1, vbTab) & " contraídas" & String(1, vbTab) & " en" & String(1, vbTab) & " este" & String(0.55, vbTab) & " documento" & String(0.55, vbTab) & " obligándome/obligándonos" & String(0.55, vbTab) & " al" & String(0.55, vbTab) & " pago" & String(0.55, vbTab) & " de" & String(0.55, vbTab) & " la" & String(0.55, vbTab) & " cantidad" & String(0.55, vbTab) & " adeudada, intereses" & String(0.55, vbTab) & " compensatorios" & String(2, vbTab) & " y" & String(2, vbTab) & " moratorios, así como comisiones,", "F1", 11
oDoc.WTextBox 91, 82, 30, 480, "LA CAJA", "F2", 10
oDoc.WTextBox 110, 45, 20, 560, "penalidades, seguros, gastos notariales, de cobranza judicial y extrajudicial, que se" & String(0.55, vbTab) & " pudieran devengar desde la fecha de emisión ", "F1", 11, hjustify
oDoc.WTextBox 120, 45, 20, 520, "hasta la cancelación total de la presente obligación.", "F1", 11, hjustify
'---------------------------------------------------------------------------------------
oDoc.WTextBox 140, 45, 20, 540, "De" & String(0.55, vbTab) & "acuerdo" & String(0.55, vbTab) & " a" & String(0.55, vbTab) & " lo" & String(0.55, vbTab) & " dispuesto" & String(0.55, vbTab) & " en" & String(0.55, vbTab) & " el" & String(0.55, vbTab) & " numeral" & String(0.55, vbTab) & " 11)" & String(0.55, vbTab) & " del" & String(0.55, vbTab) & " articulo" & String(0.55, vbTab) & " 132°" & String(0.55, vbTab) & " de" & String(0.55, vbTab) & " la" & String(0.55, vbTab) & " Ley" & String(0.55, vbTab) & " No. 26702 - Ley General del" & String(1, vbTab) & " Sistema" & String(0.55, vbTab) & " Financiero", "F1", 11, hjustify
oDoc.WTextBox 150, 45, 20, 520, " y del Sistema de Seguros y Orgánica de la Superintendencia de Banca y Seguros, autorizo(amos) a", "F1", 11, hjustify
oDoc.WTextBox 151, 448, 30, 480, "LA CAJA", "F2", 10, hjustify
oDoc.WTextBox 150, 486, 20, 480, "para que compense", "F1", 11, hjustify
oDoc.WTextBox 160, 45, 30, 535, " entre" & String(0.55, vbTab) & " mis/nuestras" & String(0.55, vbTab) & " acreencias" & String(0.5, vbTab) & " y" & String(0.5, vbTab) & " activos (cuentas, valores, depósitos en general, entre otros) que" & String(0.55, vbTab) & " mantenga" & String(0.55, vbTab) & " en" & String(0.55, vbTab) & " su poder,", "F1", 11, hjustify
oDoc.WTextBox 160, 544, 30, 535, "hasta", "F1", 11, hjustify
oDoc.WTextBox 170, 45, 30, 540, "por" & String(0.55, vbTab) & " el" & String(0.55, vbTab) & " importe" & String(0.55, vbTab) & " adeudado" & String(0.55, vbTab) & " de" & String(0.5, vbTab) & " este pagaré más los intereses compensatorios, moratorios, gastos, y cualquier otro concepto que puedan generarse.", "F1", 11, hjustify
'---------------------------------------------------------------------------------------
oDoc.WTextBox 200, 45, 30, 530, "Asimismo, me(nos)" & String(0.55, vbTab) & " someto(emos)" & String(0.55, vbTab) & " expresamente" & String(0.55, vbTab) & " a" & String(0.55, vbTab) & " la" & String(0.55, vbTab) & " competencia" & String(0.55, vbTab) & " y" & String(2, vbTab) & " tribunales" & String(2, vbTab) & " de" & String(2, vbTab) & " esta" & String(2, vbTab) & " ciudad, en" & String(1, vbTab) & " cuyo" & String(1, vbTab) & " efecto" & String(1, vbTab) & " renuncio", "F1", 11, hjustify
oDoc.WTextBox 210, 45, 30, 540, "/renunciamos al fuero de mi/nuestro domicilio. Señalo(amos) como domicilio" & String(0.55, vbTab) & " aquel" & String(0.5, vbTab) & " indicado en este pagaré a donde se efectuarán las diligencias notariales, judiciales y demás que fuesen necesarias para lo que ", "F1", 11, hjustify
oDoc.WTextBox 221, 365, 30, 480, "LA CAJA", "F2", 10, hjustify
oDoc.WTextBox 220, 400, 30, 480, "considere pertinente.Cualquier cambio de", "F1", 11, hjustify
oDoc.WTextBox 230, 45, 30, 540, " domicilio" & String(0.55, vbTab) & " que" & String(0.55, vbTab) & " haga(mos), para" & String(0.5, vbTab) & " su" & String(0.53, vbTab) & "validez, lo" & String(0.55, vbTab) & " haré(mos), mediante" & String(0.55, vbTab) & "carta" & String(0.55, vbTab) & " notarial" & String(0.55, vbTab) & " y" & String(0.55, vbTab) & " conforme" & String(0.55, vbTab) & " a lo dispuesto en el artículo 40° ", "F1", 11, hjustify
oDoc.WTextBox 230, 555, 30, 34, "del", "F1", 11, hjustify
oDoc.WTextBox 240, 45, 30, 520, "Código Civil.", "F1", 11, hjustify
'---------------------------------------------------------------------------------------
oDoc.WTextBox 260, 45, 30, 540, "Declaro(amos)" & String(0.55, vbTab) & " estar plenamente facultado(s) para afianzar el presente Pagaré, asumiendo en caso contrario la responsabilidad civil y/o penal que hubiere lugar. Se" & String(0.5, vbTab) & " deja" & String(0.55, vbTab) & "constancia que la información proporciona por el(los) fiador(es) en el presente documento, tiene el carácter de declaración jurada, de acuerdo con el artículo 179° de la Ley N° 26702.", "F1", 11, hjustify
oDoc.WTextBox 305, 45, 30, 480, "Suscribimos el presente en señal de conformidad.", "F1", 11, hjustify


If lsCiudad = "MOYOBAMBA " Or lsCiudad = "CAJAMARCA " Or lsCiudad = "YURIMAGUAS " Or lsCiudad = "TINGO MARIA " Then
oDoc.WTextBox 330, 45, 30, 480, "Lugar y fecha de emisión:" & String(23, vbTab) & " ,___de_____________________de _____", "F1", 11, hjustify
oDoc.WTextBox 332, 96, 15, 160, lsCiudad, "F2", 10, vMiddle, vbBlack, 1, vbBlack
ElseIf lsCiudad = "PUERTO CALLAO " Or lsCiudad = "CERRO DE PASCO " Or lsCiudad = "TOCACHE NUEVO " Then
oDoc.WTextBox 330, 45, 30, 480, "Lugar y fecha de emisión:" & String(30, vbTab) & " ,____de_____________________de _____", "F1", 11, hjustify
oDoc.WTextBox 332, 105, 15, 160, lsCiudad, "F2", 10, vMiddle, vbBlack, 1, vbBlack
Else: oDoc.WTextBox 330, 45, 30, 480, "Lugar y fecha de emisión:" & String(20, vbTab) & " ,___de_____________________de _____", "F1", 11, hjustify
oDoc.WTextBox 332, 90, 15, 160, lsCiudad, "F2", 10, vMiddle, vbBlack, 1, vbBlack
End If
           
Dim nGaran As Integer
nGaran = 0

If Not (RsGarantes.EOF And RsGarantes.BOF) Then
        While Not RsGarantes.EOF
        '##################################### 1 Garante ##############################################
            If nGaran = 0 And (RsGarantes!nPersPersoneria = 2 Or RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 4 Or RsGarantes!nPersPersoneria = 5 Or RsGarantes!nPersPersoneria = 6 Or RsGarantes!nPersPersoneria = 7 Or RsGarantes!nPersPersoneria = 8 Or RsGarantes!nPersPersoneria = 9 Or RsGarantes!nPersPersoneria = 10) Then
                sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR) '----Recupera el Representante
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 285 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 295 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                End If
                
             oDoc.WTextBox 245 + h, 45, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 270 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 320 + h, 95, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '1
             oDoc.WTextBox 360, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 380, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11 '
             oDoc.WTextBox 270 + h, 45, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3 '
             oDoc.WTextBox 280 + h, 45, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3 '
             oDoc.WTextBox 305 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3 '
             oDoc.WTextBox 320 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
'
             oDoc.WTextBox 365 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 345 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             
            ElseIf nGaran = 0 Then
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 285 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 295 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                   
                End If
                
             oDoc.WTextBox 245 + h, 45, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 270 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 285 + h, 95, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '1
             oDoc.WTextBox 360, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 380, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11

             oDoc.WTextBox 270 + h, 45, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 285 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 360 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 340 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
              '##################################### 2 Garante ##############################################
            ElseIf nGaran = 1 And (RsGarantes!nPersPersoneria = 2 Or RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 4 Or RsGarantes!nPersPersoneria = 5 Or RsGarantes!nPersPersoneria = 6 Or RsGarantes!nPersPersoneria = 7 Or RsGarantes!nPersPersoneria = 8 Or RsGarantes!nPersPersoneria = 9 Or RsGarantes!nPersPersoneria = 10) Then
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 285 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 295 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 245 + h, 300, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 270 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 320 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                 '2
             oDoc.WTextBox 360, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 380, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 270 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 280 + h, 300, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 305 + h, 300, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 320 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 365 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 345 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             
            ElseIf nGaran = 1 Then
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 285 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 295 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 245 + h, 300, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 270 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 285 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                 '2
             oDoc.WTextBox 360, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 380, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11

             oDoc.WTextBox 270 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3

             oDoc.WTextBox 285 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3

             oDoc.WTextBox 360 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 340 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
          '##################################### 3 Garante ##############################################
            ElseIf nGaran = 2 And (RsGarantes!nPersPersoneria = 2 Or RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 4 Or RsGarantes!nPersPersoneria = 5 Or RsGarantes!nPersPersoneria = 6 Or RsGarantes!nPersPersoneria = 7 Or RsGarantes!nPersPersoneria = 8 Or RsGarantes!nPersPersoneria = 9 Or RsGarantes!nPersPersoneria = 10) Then
            
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 505 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 515 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 465 + h, 45, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 490 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 540 + h, 95, 35, 180, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '2
              oDoc.WTextBox 580, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
              oDoc.WTextBox 600, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 490 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 500 + h, 45, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 525 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 540 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 605 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 585 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack


            ElseIf nGaran = 2 Then
            
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 505 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 515 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 465 + h, 45, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 490 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 505 + h, 95, 35, 180, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '2
              oDoc.WTextBox 580, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
              oDoc.WTextBox 600, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 490 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 505 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 605 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 585 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
           '##################################### 4 Garante ##############################################
            ElseIf nGaran = 3 And (RsGarantes!nPersPersoneria = 2 Or RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 4 Or RsGarantes!nPersPersoneria = 5 Or RsGarantes!nPersPersoneria = 6 Or RsGarantes!nPersPersoneria = 7 Or RsGarantes!nPersPersoneria = 8 Or RsGarantes!nPersPersoneria = 9 Or RsGarantes!nPersPersoneria = 10) Then
            sPersCodR = RsGarantes!cperscod
            Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 505 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 515 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 465 + h, 300, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 490 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 538 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                
                '4
             oDoc.WTextBox 580, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 600, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 490 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 500 + h, 300, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 525 + h, 300, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 538 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 605 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 585 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             
             ElseIf nGaran = 3 Then
            sPersCodR = RsGarantes!cperscod
            Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 505 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 515 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 465 + h, 300, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 490 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 505 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                
                '4
             oDoc.WTextBox 580, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 600, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 490 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 505 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 605 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 585 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             '##################################### 5 Garante ##############################################
            
            ElseIf nGaran = 4 And (RsGarantes!nPersPersoneria = 2 Or RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 4 Or RsGarantes!nPersPersoneria = 5 Or RsGarantes!nPersPersoneria = 6 Or RsGarantes!nPersPersoneria = 7 Or RsGarantes!nPersPersoneria = 8 Or RsGarantes!nPersPersoneria = 9 Or RsGarantes!nPersPersoneria = 10) Then
            h = -170
            oDoc.NewPage A4_Vertical
            oDoc.WImage 50, 494, 35, 73, "Logo"
            
                sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 283 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 295 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                   
                End If
                            
                
             oDoc.WTextBox 245 + h, 45, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 270 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 320 + h, 95, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '1
             oDoc.WTextBox 50, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 70, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 270 + h, 45, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 280 + h, 45, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 305 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 320 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 380 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 360 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
            ElseIf nGaran = 4 Then
            h = -170
            oDoc.NewPage A4_Vertical
            
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 283 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 295 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                   
                End If
                            
                
             oDoc.WTextBox 245 + h, 45, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 270 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 285 + h, 95, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '1
             oDoc.WTextBox 50, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 70, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 270 + h, 45, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 285 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 380 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 360 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             '##################################### 6 Garante ##############################################
            ElseIf nGaran = 5 And (RsGarantes!nPersPersoneria = 2 Or RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 4 Or RsGarantes!nPersPersoneria = 5 Or RsGarantes!nPersPersoneria = 6 Or RsGarantes!nPersPersoneria = 7 Or RsGarantes!nPersPersoneria = 8 Or RsGarantes!nPersPersoneria = 9 Or RsGarantes!nPersPersoneria = 10) Then
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 283 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 295 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 245 + h, 300, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 270 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 320 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                 '2
             oDoc.WTextBox 50, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 70, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 270 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 280 + h, 300, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 305 + h, 300, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 320 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 380 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 360 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             
            ElseIf nGaran = 5 Then
            
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 283 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 295 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 245 + h, 300, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 270 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 285 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                 '2
             oDoc.WTextBox 50, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 70, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 270 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 285 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 380 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 360 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             '##################################### 7 Garante ##############################################
            ElseIf nGaran = 6 And (RsGarantes!nPersPersoneria = 2 Or RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 4 Or RsGarantes!nPersPersoneria = 5 Or RsGarantes!nPersPersoneria = 6 Or RsGarantes!nPersPersoneria = 7 Or RsGarantes!nPersPersoneria = 8 Or RsGarantes!nPersPersoneria = 9 Or RsGarantes!nPersPersoneria = 10) Then
            
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 523 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 535 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 485 + h, 45, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 510 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 560 + h, 95, 35, 180, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '2
              oDoc.WTextBox 290, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
              oDoc.WTextBox 310, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 510 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 520 + h, 45, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 545 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 560 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 625 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 605 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
            ElseIf nGaran = 6 Then
            h = -170
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 523 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 535 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 485 + h, 45, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 510 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 525 + h, 95, 35, 180, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '2
              oDoc.WTextBox 290, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
              oDoc.WTextBox 310, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 510 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 525 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 625 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 605 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             '##################################### 8 Garante ##############################################
           ElseIf nGaran = 7 And (RsGarantes!nPersPersoneria = 2 Or RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 4 Or RsGarantes!nPersPersoneria = 5 Or RsGarantes!nPersPersoneria = 6 Or RsGarantes!nPersPersoneria = 7 Or RsGarantes!nPersPersoneria = 8 Or RsGarantes!nPersPersoneria = 9 Or RsGarantes!nPersPersoneria = 10) Then
            sPersCodR = RsGarantes!cperscod
            Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 523 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 535 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 485 + h, 300, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 510 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 560 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                
                '8
             oDoc.WTextBox 290, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 310, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 510 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 520 + h, 300, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 545 + h, 300, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 560 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 625 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 605 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             
             ElseIf nGaran = 7 Then
            sPersCodR = RsGarantes!cperscod
            Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 523 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 535 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 485 + h, 300, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 510 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 525 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4

                '4
             oDoc.WTextBox 290, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 310, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 510 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 525 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 625 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 605 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
              '##################################### 9 Garante ##############################################
            ElseIf nGaran = 8 And (RsGarantes!nPersPersoneria = 2 Or RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 4 Or RsGarantes!nPersPersoneria = 5 Or RsGarantes!nPersPersoneria = 6 Or RsGarantes!nPersPersoneria = 7 Or RsGarantes!nPersPersoneria = 8 Or RsGarantes!nPersPersoneria = 9 Or RsGarantes!nPersPersoneria = 10) Then
            h = 110
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 503 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 515 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 465 + h, 45, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 490 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 540 + h, 95, 35, 180, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '2
              oDoc.WTextBox 550, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
              oDoc.WTextBox 570, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 490 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 500 + h, 45, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 525 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 540 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 605 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 585 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack



            ElseIf nGaran = 8 Then
            'h = 110
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 613, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 625 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 575, 45, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 600, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 615, 95, 35, 180, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '2
              oDoc.WTextBox 550, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
              oDoc.WTextBox 570, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 600, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 615, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 715, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 685, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
           '##################################### 10 Garante ##############################################
            ElseIf nGaran = 9 And (RsGarantes!nPersPersoneria = 2 Or RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 4 Or RsGarantes!nPersPersoneria = 5 Or RsGarantes!nPersPersoneria = 6 Or RsGarantes!nPersPersoneria = 7 Or RsGarantes!nPersPersoneria = 8 Or RsGarantes!nPersPersoneria = 9 Or RsGarantes!nPersPersoneria = 10) Then
            sPersCodR = RsGarantes!cperscod
            h = 110
            Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 503 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 515 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 465 + h, 300, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 490 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 538 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                
                '4
             oDoc.WTextBox 550, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 570, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 490 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 500 + h, 300, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 525 + h, 300, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 538 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 605 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 585 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             
             ElseIf nGaran = 9 Then
            sPersCodR = RsGarantes!cperscod
            Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 503 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 515 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 465 + h, 300, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 490 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 505 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                
                '4
             oDoc.WTextBox 580, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 600, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 490 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 505 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 605 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 585 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
            End If
            nGaran = nGaran + 1
            RsGarantes.MoveNext
        Wend
    Else
End If
'########################PTI1 20170315  ########################################
End Function

'*** PEAC 20170605
Private Function QuitarCaracter(ByVal psCadena As String, ByVal psCaracter As String) As String
Dim nPosicion As Integer
Dim nTamano As Integer
Dim i As Integer
Dim sResultado As String
Dim sTemp As String
sResultado = psCadena

nTamano = Len(psCaracter)
nPosicion = Len(psCaracter)
sTemp = psCaracter

For i = 0 To nTamano - 1
    sTemp = Mid(sTemp, i + 1, 1)
    sResultado = Replace(sResultado, sTemp, "")
    sTemp = psCaracter
Next i

QuitarCaracter = sResultado
End Function

'*** PEAC 20170621
Public Function VerificarExisteDesembolsoBcoNac(ByVal psCtaCod As String, ByRef sFecDes As String, ByVal pnOpcion As Integer) As Boolean
    Dim oDCred As COMNCredito.NCOMCredito
    Dim oCredDoc As COMDCredito.DCOMCredDoc
    Dim oDCred2 As COMDCredito.DCOMCredito
    Dim bValor As Boolean
    Dim R As ADODB.Recordset
    Set oDCred = New COMNCredito.NCOMCredito
    
    bValor = oDCred.VerificarExisteDesembolsoBcoNac(psCtaCod)
    If bValor = True Then
        If (MsgBox("El Crédito no ha sido Desembolsado por lo que no cuenta con fecha para la generación del documento; desea agregar manualmente?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmación") = vbYes) Then
            
            If pnOpcion = 1 Then
                Set oDCred2 = New COMDCredito.DCOMCredito
                Set R = oDCred2.RecuperaDatosComunes(psCtaCod)
                sFecDes = frmIngFechaGenDoc.Inicio(R!dVigencia)
            End If
            
            If pnOpcion = 2 Then
                Set oCredDoc = New COMDCredito.DCOMCredDoc
                Set R = oCredDoc.RecuperaDatosDocPlanPagos(psCtaCod)
                sFecDes = frmIngFechaGenDoc.Inicio(R!dFecVig)
            End If
            
            Set R = Nothing
            
            If (sFecDes = Empty) Then
                VerificarExisteDesembolsoBcoNac = False
            Else
                VerificarExisteDesembolsoBcoNac = True
            End If
        Else
        VerificarExisteDesembolsoBcoNac = False '********Con esto ocurre el proceso normal
        End If
    Else
        VerificarExisteDesembolsoBcoNac = False '********Con esto ocurre el proceso normal
    End If
End Function

'JOEP20171218 ERS082-2017 --AGREGADO DESDE LA 60
Public Function ImprimeInformeComercialPig(ByVal psCtaCod As String, _
ByVal psNomAge As String, ByVal psCodUsu As String, ByVal pR As ADODB.Recordset, _
ByVal pRB As ADODB.Recordset, ByVal pRDatFin As ADODB.Recordset) As String

    Dim R As ADODB.Recordset
    Dim RB As ADODB.Recordset
    Dim nTasaAnual As Double
    Dim nTasaMoraA As Double

    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range

    Dim RAge As ADODB.Recordset
    Dim sCompaSeguro As String
    Dim sNumeroPoliza As String
    Dim nParametro As Double
    Dim RRelaCred  As ADODB.Recordset
    Dim RDatFin As ADODB.Recordset
    Dim oCred As COMNCredito.NCOMCredDoc
    Dim oCredB As COMNCredito.NCOMCredDoc
    Dim oDCredC As COMDCredito.DCOMCredito

    Dim nCosRef As Double
    Dim nComDsg As Double
    Dim nCosCroPag As Double
    Dim nCosHisPag As Double
    Dim nCosComPag As Double
    Dim nMontoInt As Double
    Dim nPriPol As Double
    Dim nCantAnt As Double
    Dim nCenCrg As Double
    Dim nSgrDsg1 As Double
    Dim nSgrDsg2 As Double
    Dim nSegMor As Double
    Dim nMinCuo As Double
    Dim nMaxCuo As Double

    Dim sArchivo As String

    On Error GoTo ErrorInformeCom02

    Set oWord = CreateObject("Word.Application")
        oWord.Visible = False

    Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\InformeComercialPig.doc")

    sArchivo = App.Path & "\FormatoCarta\IC_" & psCtaCod & "_" & Replace(Left(Time, 5), ":", "") & ".doc"
    oDoc.SaveAs (sArchivo)

'*** Datos de Cabecera

    With oWord.Selection.Find
        .Text = "<<cOficina>>"
        .Replacement.Text = psNomAge
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cNomCliente>>"
        .Replacement.Text = IIf(Len(pR!nombre_deudor) = 0, ".......", pR!nombre_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cCuenta>>"
        .Replacement.Text = psCtaCod
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cCodCliente>>"
        .Replacement.Text = IIf(Len(pR!CodCliente) = 0, ".......", pR!CodCliente)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cDni>>"
        .Replacement.Text = IIf(Len(pR!dni_deudor) = 0, ".......", pR!dni_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cRuc>>"
        .Replacement.Text = IIf(Len(pR!ruc_deudor) = 0, ".......", pR!ruc_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cTipCred>>"
        .Replacement.Text = IIf(Len(pR!TipoCred) = 0, ".......", pR!TipoCred)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cUsu>>"
        .Replacement.Text = psCodUsu
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<cAna>>"
        .Replacement.Text = IIf(Len(pR!Analista) = 0, ".......", pR!Analista)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

'*** Datos del Deudor

    With oWord.Selection.Find
        .Text = "<<cNombreDeudor>>"
        .Replacement.Text = pR!nombre_deudor
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cDomicilio>>"
        .Replacement.Text = IIf(Len(pR!dire_deudor) = 0, ".......", pR!dire_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cDni>>"
        .Replacement.Text = IIf(Len(pR!dni_deudor) = 0, ".......", pR!dni_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cCodSbs>>"
        .Replacement.Text = IIf(Len(pR!codsbs_deudor) = 0, ".......", pR!codsbs_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<nPatrimonio>>"
        .Replacement.Text = IIf(pRDatFin.RecordCount > 0, Format(pRDatFin!PatrimonioPerso, "#,##0.00"), ".......")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cEstadoCivil>>"
        .Replacement.Text = IIf(Len(pR!estadocivil_deudor) = 0, ".......", pR!estadocivil_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cCentroLaboral>>"
        .Replacement.Text = IIf(Len(pR!centrolaboral_deudor) = 0, ".......", pR!centrolaboral_deudor)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'fin peac

    With oWord.Selection.Find
        .Text = "<<cGiroEmpresaLabora>>"
        .Replacement.Text = "....."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<nAntiguedad>>"
        .Replacement.Text = "....."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

'*** Conyuge

    With oWord.Selection.Find
        .Text = "<<cNombreConyuge>>"
        .Replacement.Text = IIf(Len(pR!nombre_conyuge) = 0, ".......", pR!nombre_conyuge)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cDniConyuge>>"
        .Replacement.Text = IIf(Len(pR!dni_conyuge) = 0, ".......", pR!dni_conyuge)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cCodSbsConyuge>>"
        .Replacement.Text = IIf(Len(pR!codsbs_conyuge) = 0, ".......", pR!codsbs_conyuge)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With


    With oWord.Selection.Find
        .Text = "<<cCentroLaboralConyuge>>"
        .Replacement.Text = IIf(Len(pR!CentroLaboral_Conyuge) = 0, ".......", pR!CentroLaboral_Conyuge)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

'*** Datos Familiares

    With oWord.Selection.Find
        .Text = "<<cNumDependientes>>"
        .Replacement.Text = pR!cNumDependientes
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cEdades>>"
        .Replacement.Text = "....."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

'*** Datos del Reporte

    With oWord.Selection.Find
        .Text = "<<cFecha>>"
        .Replacement.Text = Format(gdFecSis, "dddd") & ", " & Day(gdFecSis) & " de " & Format(gdFecSis, "mmmm") & " de " & Year(gdFecSis)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

'*** Fin

    oDoc.Close
    Set oDoc = Nothing

    Set oWord = CreateObject("Word.Application")
    oWord.Visible = True
    Set oDoc = oWord.Documents.Open(sArchivo)
    Set oDoc = Nothing
    Set oWord = Nothing

    Exit Function
'PEAC 20090520
ErrorInformeCom02:
    MsgBox err.Description, vbCritical, "Aviso"
    Set oDoc = Nothing
    Set oWord = Nothing
End Function
'JOEP20171218 ERS082-2017


'JOEP20180725 ERS034-2018
Public Function RiesgoCambCredPDF(ByVal pcCtaCod As String, ByVal pgdFecSis As String, ByVal pgsCodUser As String, ByVal pgsNomAge As String)
Dim oDoc  As cPDF
Dim oDCredRiesgoCamb As COMDCredito.DCOMCredito
Dim rsRiegCamb As ADODB.Recordset
Dim rsValRiegCamb As ADODB.Recordset
Set oDCredRiesgoCamb = New COMDCredito.DCOMCredito
Set oDoc = New cPDF

On Error GoTo ErrorRiesgoCambCredPDF
'Verifica si emite PDF
Set rsValRiegCamb = oDCredRiesgoCamb.ValidadRigCambCred(pcCtaCod)
If Not (rsValRiegCamb.BOF And rsValRiegCamb.EOF) Then
    If rsValRiegCamb!nApli = 0 Then
        Exit Function
    End If
End If
RSClose rsValRiegCamb
Set oDCredRiesgoCamb = Nothing
'****************************************************************

'Creacion PDF
oDoc.Author = pgsCodUser
oDoc.Creator = "SICMACT - NEGOCIO"
oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
oDoc.Subject = "Identificacion de Clientes Expuestos Nº " & pcCtaCod
oDoc.Title = "Riesgo Cambiario Crediticio Nº " & pcCtaCod
'Si existe
If Not oDoc.PDFCreate(App.Path & "\Spooler\RiesgoCambiarioCrediticio" & pcCtaCod & "_" & Format(pgdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
    Exit Function
End If
'TIpo de Letra y Imagen
oDoc.Fonts.Add "F1", "Arial Narrow", TrueType, Normal, WinAnsiEncoding
oDoc.Fonts.Add "F2", "Arial Narrow", TrueType, Bold, WinAnsiEncoding
oDoc.LoadImageFromFile "C:\SICMACM_DEV\SICMAC_NEGOCIO_COM" & "\logo_cmacmaynas.bmp", "Logo"
'Hoja Vertical o Horizontal
oDoc.NewPage A4_Vertical
'Cabecera
oDoc.WImage 50, 50, 40, 80, "Logo"
oDoc.WTextBox 70, 50, 10, 500, "IDENTIFICACIÓN DE CLIENTES EXPUESTOS A RIESGO CAMBIARIO CREDITICIO", "F2", 12, hCenter

Set oDCredRiesgoCamb = New COMDCredito.DCOMCredito
Set rsRiegCamb = oDCredRiesgoCamb.ObtieneDatRigCambCred(pcCtaCod)

If Not (rsRiegCamb.BOF And rsRiegCamb.EOF) Then
    oDoc.WTextBox 20, 10, 10, 500, "Fecha: " & Format(pgdFecSis, "DD/MM/YYYY") & " " & Format(Time, "hh:mm:ss"), "F1", 12, hRight
    oDoc.WTextBox 30, 10, 10, 446, "Usuario: " & UCase(IIf(IsNull(pgsCodUser), "", pgsCodUser)), "F1", 12, hRight
    
    'Contenido
    oDoc.WTextBox 110, 50, 10, 500, "NOMBRE DEL CLIENTE: ", "F1", 12, hLeft
    oDoc.WTextBox 110, 180, 10, 500, UCase(rsRiegCamb!cPersNombre), "F1", 12, hLeft
    oDoc.WTextBox 123, 50, 10, 500, "N° Crédito: ", "F1", 12, hLeft
    oDoc.WTextBox 123, 180, 10, 500, pcCtaCod, "F1", 12, hLeft
    oDoc.WTextBox 136, 50, 10, 500, "AGENCIA: ", "F1", 12, hLeft
    oDoc.WTextBox 136, 180, 10, 500, UCase(pgsNomAge), "F1", 12, hLeft
    oDoc.WTextBox 149, 50, 10, 500, "ACTIVIDAD: ", "F1", 12, hLeft
    oDoc.WTextBox 149, 180, 10, 500, UCase(rsRiegCamb!cActividad), "F1", 12, hLeft
    oDoc.WTextBox 162, 50, 10, 500, "ANALISTA: ", "F1", 12, hLeft
    oDoc.WTextBox 162, 180, 10, 500, UCase(rsRiegCamb!cUserAnal), "F1", 12, hLeft
    oDoc.WTextBox 175, 50, 10, 500, "FECHA DE EVALUACIÓN: ", "F1", 12, hLeft
    oDoc.WTextBox 175, 180, 10, 500, rsRiegCamb!dFecEval, "F1", 12, hLeft
    'Cuadro
    oDoc.WRectangle 208, 50, 30, 95, 1, vbBlack, False
    oDoc.WRectangle 208, 145, 30, 150, 1, vbBlack, False
    oDoc.WRectangle 208, 295, 30, 150, 1, vbBlack, False
    oDoc.WRectangle 238, 50, 30, 95, 1, vbBlack, False
    oDoc.WRectangle 238, 145, 30, 150, 1, vbBlack, False
    oDoc.WRectangle 238, 295, 30, 150, 1, vbBlack, False
    'Cuadro
    
    oDoc.WTextBox 210, 150, 10, 80, "OBLIGACIONES EN SOLES (S/)", "F1", 12, hLeft
    oDoc.WTextBox 223, 242, 10, 80, Format(rsRiegCamb!nTotalPasivo, "#,#0.00"), "F1", 12, hLeft
    oDoc.WTextBox 210, 300, 10, 80, "OBLIGACIONES EN DOLARES ($)", "F1", 12, hLeft
    oDoc.WTextBox 223, 390, 10, 80, Format(rsRiegCamb!nMonto, "#,#0.00"), "F1", 12, hLeft
    oDoc.WTextBox 210, 55, 10, 200, "INGRESOS (S/)", "F1", 12, hLeft
    oDoc.WTextBox 240, 60, 10, 200, Format(rsRiegCamb!nIngresos, "#,#0.00"), "F1", 12, hLeft
    
    If rsRiegCamb!nACCapPagoD = 0 And rsRiegCamb!nACCapPagV = 0 Then
        oDoc.WTextBox 240, 150, 10, 200, "EXPUESTO ", "F2", 12, hLeft
    ElseIf (rsRiegCamb!nACCapPagoD = 0 And rsRiegCamb!nACCapPagV = 1) Or (rsRiegCamb!nACCapPagoD = 1 And rsRiegCamb!nACCapPagV = 0) Then
        oDoc.WTextBox 240, 150, 10, 200, "EXPUESTO ", "F2", 12, hLeft
        oDoc.WTextBox 240, 300, 10, 200, "NO EXPUESTO ", "F2", 12, hLeft
    Else
        oDoc.WTextBox 240, 300, 10, 200, "NO EXPUESTO ", "F2", 12, hLeft
    End If
    
    'Cuadro
    oDoc.WRectangle 298, 145, 20, 230, 1, vbBlack, False
    oDoc.WRectangle 318, 145, 18, 230, 1, vbBlack, False
    oDoc.WRectangle 335, 145, 18, 230, 1, vbBlack, False
    oDoc.WRectangle 353, 145, 18, 230, 1, vbBlack, False
    oDoc.WRectangle 318, 210, 53, 1, 1, vbBlack, True
    oDoc.WRectangle 318, 285, 53, 1, 1, vbBlack, True
    'Cuadro
    
    oDoc.WTextBox 300, 10, 10, 500, "ESTRUCTURA DE INGRESOS Y EGRESOS", "F1", 12, hCenter
    oDoc.WTextBox 320, 220, 10, 500, "% SOLES", "F1", 12, hLeft
    oDoc.WTextBox 320, 290, 10, 500, "% DOLARES", "F1", 12, hLeft
    oDoc.WTextBox 335, 150, 10, 500, "INGRESOS", "F1", 12, hLeft
    oDoc.WTextBox 335, 70, 10, 210, rsRiegCamb!IngMN & "%", "F1", 12, hRight
    oDoc.WTextBox 335, 160, 10, 210, rsRiegCamb!IngME & "%", "F1", 12, hRight
    oDoc.WTextBox 353, 150, 10, 500, "EGRESOS", "F1", 12, hLeft
    oDoc.WTextBox 353, 70, 10, 210, rsRiegCamb!EgrMN & "%", "F1", 12, hRight
    oDoc.WTextBox 353, 160, 10, 210, rsRiegCamb!EgrME & "%", "F1", 12, hRight
    
    oDoc.WTextBox 402, 53, 10, 500, "TIPO DE CAMBIO", "F2", 11, hLeft
    oDoc.WTextBox 402, 50, 10, 110, Format(rsRiegCamb!TC, "#,#0.000"), "F2", 12, hRight
    'Cuadro
    oDoc.WRectangle 400, 50, 20, 115, 1, vbBlack, False
    oDoc.WRectangle 420, 50, 20, 375, 1, vbBlack, False
    oDoc.WRectangle 440, 50, 20, 375, 1, vbBlack, False
    oDoc.WRectangle 460, 50, 20, 375, 1, vbBlack, False
    oDoc.WRectangle 480, 50, 20, 375, 1, vbBlack, False
    oDoc.WRectangle 420, 165, 80, 1, 1, vbBlack, True
    oDoc.WRectangle 420, 275, 80, 1, 1, vbBlack, True
    oDoc.WRectangle 420, 350, 80, 1, 1, vbBlack, True
    'Cuadro
    
    oDoc.WTextBox 420, 170, 10, 500, "EN LA EVALUACIÓN", "F2", 12, hLeft
    oDoc.WTextBox 420, 280, 10, 500, "SHOCK 10%", "F2", 12, hLeft
    oDoc.WTextBox 420, 360, 10, 500, "SHOCK 20%", "F2", 12, hLeft
    oDoc.WTextBox 440, 53, 10, 500, "EXCENDENTE", "F1", 12, hLeft
    oDoc.WTextBox 440, 60, 10, 210, Format(rsRiegCamb!Excedente, "#,#0.00"), "F1", 12, hRight
    oDoc.WTextBox 440, 135, 10, 210, Format(rsRiegCamb!Excedente, "#,#0.00"), "F1", 12, hRight
    oDoc.WTextBox 440, 208, 10, 210, Format(rsRiegCamb!Excedente, "#,#0.00"), "F1", 12, hRight
    oDoc.WTextBox 463, 53, 10, 500, "CUOTA", "F1", 12, hLeft
    oDoc.WTextBox 463, 60, 10, 210, Format(rsRiegCamb!CuotaMensual, "#,#0.00"), "F1", 12, hRight
    oDoc.WTextBox 463, 135, 10, 210, Format(rsRiegCamb!CuotaD, "#,#0.00"), "F1", 12, hRight
    oDoc.WTextBox 463, 208, 10, 210, Format(rsRiegCamb!CuotaV, "#,#0.00"), "F1", 12, hRight
    oDoc.WTextBox 483, 53, 10, 500, "CUOTA / EXCENDENTE", "F2", 12, hLeft
    oDoc.WTextBox 483, 60, 10, 210, Format(rsRiegCamb!CutExcMens, "#,#0.00") & " %", "F2", 12, hRight
    oDoc.WTextBox 483, 135, 10, 210, Format(rsRiegCamb!CutExcD, "#,#0.00") & " %", "F2", 12, hRight
    oDoc.WTextBox 483, 208, 10, 210, Format(rsRiegCamb!CutExcV, "#,#0.00") & " %", "F2", 12, hRight
    
    'Cuadro
    oDoc.WRectangle 560, 130, 15, 70, 1, vbBlack, False
    oDoc.WRectangle 575, 130, 15, 70, 1, vbBlack, False
    'Cuadro
    
    If rsRiegCamb!nACCapPagoD = 1 And rsRiegCamb!nACCapPagV = 1 Then
        oDoc.WTextBox 560, 50, 10, 500, "Aceptable: ", "F1", 12, hLeft
        oDoc.WTextBox 560, 160, 10, 150, Format(rsRiegCamb!nCapPagoD, "#,#0.00") & " %", "F1", 12, hLeft
        oDoc.WTextBox 560, 210, 10, 150, "SHOCK 10%", "F1", 12, hLeft
        oDoc.WTextBox 575, 160, 10, 150, Format(rsRiegCamb!nCapPagoV, "#,#0.00") & " %", "F1", 12, hLeft
        oDoc.WTextBox 575, 210, 10, 150, "SHOCK 20%", "F1", 12, hLeft
    ElseIf (rsRiegCamb!nACCapPagoD = 1 And rsRiegCamb!nACCapPagV = 0) Or (rsRiegCamb!nACCapPagoD = 0 And rsRiegCamb!nACCapPagV = 1) Then
        If rsRiegCamb!nACCapPagoD = 1 Then
            oDoc.WTextBox 560, 50, 10, 500, "Aceptable: ", "F1", 12, hLeft
            oDoc.WTextBox 560, 160, 10, 150, Format(rsRiegCamb!nCapPagoD, "#,#0.00") & " %", "F1", 12, hLeft
            oDoc.WTextBox 560, 210, 10, 150, "SHOCK 10%", "F1", 12, hLeft
        Else
            oDoc.WTextBox 560, 50, 10, 500, "Critico: ", "F1", 12, hLeft
            oDoc.WTextBox 560, 160, 10, 150, Format(rsRiegCamb!nCapPagoD, "#,#0.00") & " %", "F1", 12, hLeft
            oDoc.WTextBox 560, 210, 10, 150, "SHOCK 20%", "F1", 12, hLeft
        End If
        If rsRiegCamb!nACCapPagV = 1 Then
            oDoc.WTextBox 560, 50, 10, 500, "Aceptable: ", "F1", 12, hLeft
            oDoc.WTextBox 575, 160, 10, 150, Format(rsRiegCamb!nCapPagoV, "#,#0.00") & " %", "F1", 12, hLeft
            oDoc.WTextBox 575, 210, 10, 150, "SHOCK 10%", "F1", 12, hLeft
        Else
            oDoc.WTextBox 560, 50, 10, 500, "Critico: ", "F1", 12, hLeft
            oDoc.WTextBox 575, 160, 10, 150, Format(rsRiegCamb!nCapPagoV, "#,#0.00") & " %", "F1", 12, hLeft
            oDoc.WTextBox 575, 210, 10, 150, "SHOCK 20%", "F1", 12, hLeft
        End If
    Else
        oDoc.WTextBox 560, 50, 10, 500, "Critico: ", "F1", 12, hLeft
        oDoc.WTextBox 560, 160, 10, 150, Format(rsRiegCamb!nCapPagoD, "#,#0.00") & " %", "F1", 12, hLeft
        oDoc.WTextBox 560, 210, 10, 150, "SHOCK 10%", "F1", 12, hLeft
        oDoc.WTextBox 575, 160, 10, 150, Format(rsRiegCamb!nCapPagoV, "#,#0.00") & " %", "F1", 12, hLeft
        oDoc.WTextBox 575, 210, 10, 150, "SHOCK 20%", "F1", 12, hLeft
    End If
    
    oDoc.WTextBox 730, 50, 10, 500, "-------------------------------------------------------", "F1", 12, hCenter
    oDoc.WTextBox 740, 50, 10, 490, "Firma del Analista de Créditos", "F1", 12, hCenter
    
    oDoc.PDFClose
    oDoc.Show
End If
    RSClose rsRiegCamb
    Set oDCredRiesgoCamb = Nothing
    Set oDoc = Nothing
Exit Function
ErrorRiesgoCambCredPDF:
    MsgBox err.Description, vbCritical, "Aviso"
End Function
'JOEP20180725 ERS034-2018
'JATO 20210408 ACTA ACTA Nº 042 - 2021
Public Sub ContratoHipotecario(psCtaCod As String, ByVal sTpoProdCod As String)
    Dim rs1 As ADODB.Recordset
    Dim lsAgencia As String
    Dim lsFecha As String
    Dim nPag As Integer
    Dim nDoc As Integer
    Dim lsArchivo As String
    Dim loAge As COMDConstantes.DCOMAgencias
    Dim lsModeloPlantilla As String
    Dim lsNomMaq As String
    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range
    Dim sArchivo As String
    Dim lsNom1, lsNom2, lsNom3, lsNom4 As String
    Dim lsDoc1, lsDoc2, lsDoc3, lsDoc4 As String
    Dim lsDir1, lsDir2, lsDir3, lsDir4 As String
    Dim oDCred As COMDCredito.DCOMCredito
    Set oDCred = New COMDCredito.DCOMCredito
    Dim RRelaCred  As ADODB.Recordset
    Set RRelaCred = oDCred.RecuperaRelacPers(psCtaCod)
    Set oDCred = Nothing
    Set loAge = New COMDConstantes.DCOMAgencias
    Set rs1 = New ADODB.Recordset
        Set rs1 = loAge.RecuperaAgencias(gsCodAge)
        If Not (rs1.EOF And rs1.BOF) Then
            lsAgencia = Trim(rs1("Dist"))
        End If
    Set loAge = Nothing

    Set oWord = CreateObject("Word.Application")
        oWord.Visible = False
    If sTpoProdCod = "806" Then
    Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\chipotecarioFMV" & gsCodAge & ".doc")
    Else
    Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\chipotecarioTP" & gsCodAge & ".doc")
    End If
    sArchivo = App.Path & "\FormatoCarta\ICC_" & psCtaCod & "_" & Replace(Left(Time, 5), ":", "") & ".doc"
    oDoc.SaveAs (sArchivo)
    Dim i As Integer, J As Integer
    Dim rsDir As ADODB.Recordset
    Dim oPers As COMDPersona.DCOMPersona
    

    
If Not (RRelaCred.EOF And RRelaCred.BOF) Then
       Do Until RRelaCred.EOF
          If RRelaCred!nConsValor = gColRelPersTitular Then
               With oWord.Selection.Find
                    .Text = "<<cNomCli>>"
                    .Replacement.Text = PstaNombre(RRelaCred!cPersNombre)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                    
               End With

               With oWord.Selection.Find
                    .Text = "<<cNuDoci>>"
                    .Replacement.Text = IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
               End With
               With oWord.Selection.Find
                    .Text = "<<DirTit1>>"
                    .Replacement.Text = Trim(RRelaCred!cPersDireccDomicilio)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                    
               End With
               
            ElseIf RRelaCred!nConsValor = gColRelPersRepresentante Then
               With oWord.Selection.Find
               J = J + 1
                    .Text = "<<Representante" & J & ">>"
                    .Replacement.Text = vbCr & "Representante: " & "<<NombreRepresentante" & J & ">>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = False
                    .Execute Replace:=wdReplaceAll
                    
               End With
               
               With oWord.Selection.Find
                    .Text = "<<NombreRepresentante" & J & ">>"
                    .Replacement.Text = PstaNombre(RRelaCred!cPersNombre) & "<<vDOIRepresentante" & J & ">>" & vbCr
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = True
                    .Execute Replace:=wdReplaceAll
               End With
               
               With oWord.Selection.Find
                    .Text = "<<vDOIRepresentante" & J & ">>"
                    .Replacement.Text = "     DOI: " & "<<cDOIRepresentante" & J & ">>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = False
                    .Execute Replace:=wdReplaceAll
               End With
               
               With oWord.Selection.Find
                    .Text = "<<cDOIRepresentante" & J & ">>"
                    .Replacement.Text = IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI) & " <<Representante" & J + 1 & ">>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = True
                    .Execute Replace:=wdReplaceAll
               End With
        
         ElseIf RRelaCred!nConsValor = gColRelPersCodeudor Then
         
         With oWord.Selection.Find
                    .Text = "<<vConyugeCodeudor1>>"
                    .Replacement.Text = vbCr & "Codeudor " & vbCr & "<<vFirmaConyugeCodeudor1>>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
               End With
               
               With oWord.Selection.Find
                    .Text = "<<vFirmaConyugeCodeudor1>>"
                    .Replacement.Text = vbCr & "Firma y huella: .........................................................." & "<<NombreConyugeCodeudor1>>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = False
                    .Execute Replace:=wdReplaceAll
               End With
                With oWord.Selection.Find
                    .Text = "<<NombreConyugeCodeudor1>>"
                    .Replacement.Text = vbCr & PstaNombre(RRelaCred!cPersNombre) & vbCr & "<<vDOIConyugeCodeudor1>>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = True
                    .Execute Replace:=wdReplaceAll
                End With

                With oWord.Selection.Find
                    .Text = "<<vDOIConyugeCodeudor1>>"
                    .Replacement.Text = "DOI: " & "<<cDOIConyugeCodeudor1>>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = False
                    .Execute Replace:=wdReplaceAll
                End With

                With oWord.Selection.Find
                    .Text = "<<cDOIConyugeCodeudor1>>"
                    .Replacement.Text = IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI) & "<<vDireccionConyugeCodeudor1>>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = True
                    .Execute Replace:=wdReplaceAll
                End With

                With oWord.Selection.Find
                    .Text = "<<vDireccionConyugeCodeudor1>>"
                    .Replacement.Text = vbCr & "Dirección: " & "<<cDireccionConyugeCodeudor1>>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = False
                    .Execute Replace:=wdReplaceAll
                End With

                With oWord.Selection.Find
                    .Text = "<<cDireccionConyugeCodeudor1>>"
                    .Replacement.Text = Trim(RRelaCred!cPersDireccDomicilio) & vbCr
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = True
                    .Execute Replace:=wdReplaceAll
                End With
         

          
          ElseIf RRelaCred!nConsValor = gColRelPersGarante Then
                i = i + 1

                With oWord.Selection.Find
                    .Text = "<<vFiador" & i & ">>"
                    .Replacement.Text = "FIADOR SOLIDARIO"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                
                With oWord.Selection.Find
                    .Text = "<<vFirma" & i & ">>"
                    .Replacement.Text = "Firma y huella: .........................................................."
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = False
                    .Execute Replace:=wdReplaceAll
                End With
                
                With oWord.Selection.Find
                    .Text = "<<cNombres" & i & ">>"
                    .Replacement.Text = PstaNombre(RRelaCred!cPersNombre)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                
                With oWord.Selection.Find
                    .Text = "<<vDOI" & i & ">>"
                    .Replacement.Text = "DOI:"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = False
                    .Execute Replace:=wdReplaceAll
                End With
                
                With oWord.Selection.Find
                    .Text = "<<cDOI" & i & ">>"
                    .Replacement.Text = IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = True
                    .Execute Replace:=wdReplaceAll
                End With
                
                With oWord.Selection.Find
                    .Text = "<<vDireccion" & i & ">>"
                    .Replacement.Text = "Dirección:"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = False
                    .Execute Replace:=wdReplaceAll
                End With
                
       If RRelaCred!nCount = RRelaCred!nMax And RRelaCred!nConsValor = gColRelPersGarante Then
                
                With oWord.Selection.Find
                    .Text = "<<cDireccion" & i & ">>"
                    .Replacement.Text = Trim(RRelaCred!cPersDireccDomicilio)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = True
                    .Execute Replace:=wdReplaceAll
                End With
        Else
        
            With oWord.Selection.Find
                        .Text = "<<cDireccion" & i & ">>"
                        .Replacement.Text = Trim(RRelaCred!cPersDireccDomicilio) & vbCr & vbCr & "<<vfiador" & i + 1 & ">>" & vbCr & vbCr & "<<vFirma" & i + 1 & ">>" & vbCr & "<<cNombres" & i + 1 & ">>" & vbCr & "<<vDOI" & i + 1 & ">> " & "<<cDOI" & i + 1 & ">>" & vbCr & "<<vDireccion" & i + 1 & ">> " & "<<cDireccion" & i + 1 & ">>" & vbCr
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Replacement.Font.Bold = True
                        .Execute Replace:=wdReplaceAll
                    End With
        
        End If

            End If
          RRelaCred.MoveNext
       Loop
    End If

    
    RRelaCred.MoveFirst
    Do Until RRelaCred.EOF
        If RRelaCred!nConsValor = gColRelPersConyugue Then
            With oWord.Selection.Find
                    .Text = "<<vConyugeCodeudor1>>"
                    .Replacement.Text = vbCr & "Conyuge" & vbCr & "<<vFirmaConyugeCodeudor1>>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
               End With

               With oWord.Selection.Find
                    .Text = "<<vFirmaConyugeCodeudor1>>"
                    .Replacement.Text = vbCr & "Firma y huella: .........................................................." & "<<NombreConyugeCodeudor1>>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = False
                    .Execute Replace:=wdReplaceAll
               End With

                With oWord.Selection.Find
                    .Text = "<<NombreConyugeCodeudor1>>"
                    .Replacement.Text = vbCr & PstaNombre(RRelaCred!cPersNombre) & vbCr & "<<vDOIConyugeCodeudor1>>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = True
                    .Execute Replace:=wdReplaceAll
                End With

                With oWord.Selection.Find
                    .Text = "<<vDOIConyugeCodeudor1>>"
                    .Replacement.Text = "DOI: " & "<<cDOIConyugeCodeudor1>>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = False
                    .Execute Replace:=wdReplaceAll
                End With

                With oWord.Selection.Find
                    .Text = "<<cDOIConyugeCodeudor1>>"
                    .Replacement.Text = IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI) & "<<vDireccionConyugeCodeudor1>>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = True
                    .Execute Replace:=wdReplaceAll
                End With

                With oWord.Selection.Find
                    .Text = "<<vDireccionConyugeCodeudor1>>"
                    .Replacement.Text = vbCr & "Dirección: " & "<<cDireccionConyugeCodeudor1>>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = False
                    .Execute Replace:=wdReplaceAll
                End With

                With oWord.Selection.Find
                    .Text = "<<cDireccionConyugeCodeudor1>>"
                    .Replacement.Text = Trim(RRelaCred!cPersDireccDomicilio) & vbCr
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = True
                    .Execute Replace:=wdReplaceAll
                End With
        
        End If
        
      RRelaCred.MoveNext
    Loop

    With oWord.Selection.Find
    Dim Text As String
    Text = "<<vFiador" & i + 1 & ">>"
        Do While (.Execute(findtext:="<<vFiador" & i + 1 & ">>", Forward:=True) = True) = True
    oWord.Selection.HomeKey Unit:=wdLine
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
        Loop
    End With
    

    With oWord.Selection.Find
        Do While (.Execute(findtext:="<<vFirma" & i + 1 & ">>", Forward:=True) = True) = True
    oWord.Selection.HomeKey Unit:=wdLine
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
        Loop
    End With
    

    With oWord.Selection.Find
        Do While (.Execute(findtext:="<<cNombres" & i + 1 & ">>", Forward:=True) = True) = True
    oWord.Selection.HomeKey Unit:=wdLine
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
        Loop
    End With
    

    With oWord.Selection.Find
        Do While (.Execute(findtext:="<<vDOI" & i + 1 & ">>", Forward:=True) = True) = True
    oWord.Selection.HomeKey Unit:=wdLine
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
        Loop
    End With
    

    With oWord.Selection.Find
    Do While (.Execute(findtext:="<<cDOI" & i + 1 & ">>", Forward:=True) = True) = True
    oWord.Selection.HomeKey Unit:=wdLine
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
        Loop
    End With
    

    With oWord.Selection.Find
    Do While (.Execute(findtext:="<<vDireccion" & i + 1 & ">>", Forward:=True) = True) = True
    oWord.Selection.HomeKey Unit:=wdLine
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
        Loop
    End With
    

    With oWord.Selection.Find
    Do While (.Execute(findtext:="<<cDireccion" & i + 1 & ">>", Forward:=True) = True) = True
    oWord.Selection.HomeKey Unit:=wdLine
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
        Loop
    End With

    With oWord.Selection.Find
         .Text = "<<vConyugeCodeudor1>>"
         .Replacement.Text = ""
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
         .Text = "<<Representante" & J + 1 & ">>"
         .Replacement.Text = ""
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
        .Text = "<<Zona>>"
        .Replacement.Text = lsAgencia
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<fDay>>"
        .Replacement.Text = Format(gdFecSis, "dd")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<fMes>>"
        .Replacement.Text = UCase(MonthName(Month(gdFecSis)))
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<fAnio>>"
        .Replacement.Text = Format(gdFecSis, "yyyy")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    oDoc.Close
    Set oDoc = Nothing

    Set oWord = CreateObject("Word.Application")
    oWord.Visible = True
    Set oDoc = oWord.Documents.Open(sArchivo)
    Set oDoc = Nothing
    Set oWord = Nothing
End Sub
'GEMO 12/11/2019 ACTA 141 - 2019 :  MEJORAS EN EL CONTRATO MULTIPRODUCTO ACTIVO

'EAAS20190523 Memorándum Nº 756-2019-GM-DI/CMACM
Public Sub ContratosAutomaticosDes(psCtaCod As String)

    
    Dim rs1 As ADODB.Recordset
    Dim lsAgencia As String
    Dim lsFecha As String
    Dim nPag As Integer
    Dim nDoc As Integer
    Dim lsArchivo As String
    Dim loAge As COMDConstantes.DCOMAgencias
    Dim lsModeloPlantilla As String
    Dim lsNomMaq As String
    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range
    Dim sArchivo As String
    Dim lsNom1, lsNom2, lsNom3, lsNom4 As String
    Dim lsDoc1, lsDoc2, lsDoc3, lsDoc4 As String
    Dim lsDir1, lsDir2, lsDir3, lsDir4 As String
    Dim oDCred As COMDCredito.DCOMCredito
    Set oDCred = New COMDCredito.DCOMCredito
    Dim RRelaCred  As ADODB.Recordset
    Set RRelaCred = oDCred.RecuperaRelacPers(psCtaCod)
    Set oDCred = Nothing
    
    
    Set loAge = New COMDConstantes.DCOMAgencias
    Set rs1 = New ADODB.Recordset
        Set rs1 = loAge.RecuperaAgencias(gsCodAge)
        If Not (rs1.EOF And rs1.BOF) Then
            lsAgencia = Trim(rs1("Dist"))
        End If
    Set loAge = Nothing

    Set oWord = CreateObject("Word.Application")
        oWord.Visible = False

    Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\cactivosDesembolsos" & gsCodAge & ".doc")

    sArchivo = App.Path & "\FormatoCarta\ICC_" & psCtaCod & "_" & Replace(Left(Time, 5), ":", "") & ".doc"
    oDoc.SaveAs (sArchivo)
    Dim i As Integer, J As Integer
    Dim rsDir As ADODB.Recordset
    Dim oPers As COMDPersona.DCOMPersona
    

    
If Not (RRelaCred.EOF And RRelaCred.BOF) Then
       Do Until RRelaCred.EOF
          If RRelaCred!nConsValor = gColRelPersTitular Then
               With oWord.Selection.Find
                    .Text = "<<cNomCli>>"
                    .Replacement.Text = PstaNombre(RRelaCred!cPersNombre)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                    
               End With

               With oWord.Selection.Find
                    .Text = "<<cNuDoci>>"
                    .Replacement.Text = IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
               End With
               With oWord.Selection.Find
                    .Text = "<<DirTit1>>"
                    .Replacement.Text = Trim(RRelaCred!cPersDireccDomicilio)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                    
               End With
               
            ElseIf RRelaCred!nConsValor = gColRelPersRepresentante Then
               With oWord.Selection.Find
               J = J + 1
                    .Text = "<<Representante" & J & ">>"
                    .Replacement.Text = vbCr & "Representante: " & "<<NombreRepresentante" & J & ">>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = False
                    .Execute Replace:=wdReplaceAll
                    
               End With
               
               With oWord.Selection.Find
                    .Text = "<<NombreRepresentante" & J & ">>"
                    .Replacement.Text = PstaNombre(RRelaCred!cPersNombre) & "<<vDOIRepresentante" & J & ">>" & vbCr
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = True
                    .Execute Replace:=wdReplaceAll
               End With
               
               With oWord.Selection.Find
                    .Text = "<<vDOIRepresentante" & J & ">>"
                    .Replacement.Text = "     DOI: " & "<<cDOIRepresentante" & J & ">>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = False
                    .Execute Replace:=wdReplaceAll
               End With
               
               With oWord.Selection.Find
                    .Text = "<<cDOIRepresentante" & J & ">>"
                    .Replacement.Text = IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI) & " <<Representante" & J + 1 & ">>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = True
                    .Execute Replace:=wdReplaceAll
               End With
        
         ElseIf RRelaCred!nConsValor = gColRelPersCodeudor Then
         
         With oWord.Selection.Find
                    .Text = "<<vConyugeCodeudor1>>"
                    .Replacement.Text = vbCr & "Codeudor " & vbCr & "<<vFirmaConyugeCodeudor1>>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
               End With
               
               With oWord.Selection.Find
                    .Text = "<<vFirmaConyugeCodeudor1>>"
                    .Replacement.Text = vbCr & "Firma y huella: .........................................................." & "<<NombreConyugeCodeudor1>>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = False
                    .Execute Replace:=wdReplaceAll
               End With
                With oWord.Selection.Find
                    .Text = "<<NombreConyugeCodeudor1>>"
                    .Replacement.Text = vbCr & PstaNombre(RRelaCred!cPersNombre) & vbCr & "<<vDOIConyugeCodeudor1>>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = True
                    .Execute Replace:=wdReplaceAll
                End With

                With oWord.Selection.Find
                    .Text = "<<vDOIConyugeCodeudor1>>"
                    .Replacement.Text = "DOI: " & "<<cDOIConyugeCodeudor1>>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = False
                    .Execute Replace:=wdReplaceAll
                End With

                With oWord.Selection.Find
                    .Text = "<<cDOIConyugeCodeudor1>>"
                    .Replacement.Text = IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI) & "<<vDireccionConyugeCodeudor1>>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = True
                    .Execute Replace:=wdReplaceAll
                End With

                With oWord.Selection.Find
                    .Text = "<<vDireccionConyugeCodeudor1>>"
                    .Replacement.Text = vbCr & "Dirección: " & "<<cDireccionConyugeCodeudor1>>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = False
                    .Execute Replace:=wdReplaceAll
                End With

                With oWord.Selection.Find
                    .Text = "<<cDireccionConyugeCodeudor1>>"
                    .Replacement.Text = Trim(RRelaCred!cPersDireccDomicilio) & vbCr
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = True
                    .Execute Replace:=wdReplaceAll
                End With
         

          
          ElseIf RRelaCred!nConsValor = gColRelPersGarante Then
                i = i + 1


' ********************************* INICIO GEMO *********************************************************************

                With oWord.Selection.Find
                    .Text = "<<vFiador" & i & ">>"
                    .Replacement.Text = "FIADOR SOLIDARIO"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                
                With oWord.Selection.Find
                    .Text = "<<vFirma" & i & ">>"
                    .Replacement.Text = "Firma y huella: .........................................................."
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = False
                    .Execute Replace:=wdReplaceAll
                End With
                
                With oWord.Selection.Find
                    .Text = "<<cNombres" & i & ">>"
                    .Replacement.Text = PstaNombre(RRelaCred!cPersNombre)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                
                With oWord.Selection.Find
                    .Text = "<<vDOI" & i & ">>"
                    .Replacement.Text = "DOI:"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = False
                    .Execute Replace:=wdReplaceAll
                End With
                
                With oWord.Selection.Find
                    .Text = "<<cDOI" & i & ">>"
                    .Replacement.Text = IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = True
                    .Execute Replace:=wdReplaceAll
                End With
                
                With oWord.Selection.Find
                    .Text = "<<vDireccion" & i & ">>"
                    .Replacement.Text = "Dirección:"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = False
                    .Execute Replace:=wdReplaceAll
                End With
                
       If RRelaCred!nCount = RRelaCred!nMax And RRelaCred!nConsValor = gColRelPersGarante Then
                
                With oWord.Selection.Find
                    .Text = "<<cDireccion" & i & ">>"
                    .Replacement.Text = Trim(RRelaCred!cPersDireccDomicilio)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = True
                    .Execute Replace:=wdReplaceAll
                End With
        Else
        
            With oWord.Selection.Find
                        .Text = "<<cDireccion" & i & ">>"
                        .Replacement.Text = Trim(RRelaCred!cPersDireccDomicilio) & vbCr & vbCr & "<<vfiador" & i + 1 & ">>" & vbCr & vbCr & "<<vFirma" & i + 1 & ">>" & vbCr & "<<cNombres" & i + 1 & ">>" & vbCr & "<<vDOI" & i + 1 & ">> " & "<<cDOI" & i + 1 & ">>" & vbCr & "<<vDireccion" & i + 1 & ">> " & "<<cDireccion" & i + 1 & ">>" & vbCr
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Replacement.Font.Bold = True
                        .Execute Replace:=wdReplaceAll
                    End With
        
        End If

' ********************************* FIN GEMO *********************************************************************

            End If
          RRelaCred.MoveNext
       Loop
    End If

    
    RRelaCred.MoveFirst
    Do Until RRelaCred.EOF
        If RRelaCred!nConsValor = gColRelPersConyugue Then
            With oWord.Selection.Find
                    .Text = "<<vConyugeCodeudor1>>"
                    .Replacement.Text = vbCr & "Conyuge" & vbCr & "<<vFirmaConyugeCodeudor1>>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
               End With

               With oWord.Selection.Find
                    .Text = "<<vFirmaConyugeCodeudor1>>"
                    .Replacement.Text = vbCr & "Firma y huella: .........................................................." & "<<NombreConyugeCodeudor1>>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = False
                    .Execute Replace:=wdReplaceAll
               End With

                With oWord.Selection.Find
                    .Text = "<<NombreConyugeCodeudor1>>"
                    .Replacement.Text = vbCr & PstaNombre(RRelaCred!cPersNombre) & vbCr & "<<vDOIConyugeCodeudor1>>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = True
                    .Execute Replace:=wdReplaceAll
                End With

                With oWord.Selection.Find
                    .Text = "<<vDOIConyugeCodeudor1>>"
                    .Replacement.Text = "DOI: " & "<<cDOIConyugeCodeudor1>>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = False
                    .Execute Replace:=wdReplaceAll
                End With

                With oWord.Selection.Find
                    .Text = "<<cDOIConyugeCodeudor1>>"
                    .Replacement.Text = IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI) & "<<vDireccionConyugeCodeudor1>>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = True
                    .Execute Replace:=wdReplaceAll
                End With

                With oWord.Selection.Find
                    .Text = "<<vDireccionConyugeCodeudor1>>"
                    .Replacement.Text = vbCr & "Dirección: " & "<<cDireccionConyugeCodeudor1>>"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = False
                    .Execute Replace:=wdReplaceAll
                End With

                With oWord.Selection.Find
                    .Text = "<<cDireccionConyugeCodeudor1>>"
                    .Replacement.Text = Trim(RRelaCred!cPersDireccDomicilio) & vbCr
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Replacement.Font.Bold = True
                    .Execute Replace:=wdReplaceAll
                End With
        
        End If
        
      RRelaCred.MoveNext
    Loop
' ************************************* INICIO GEMO **********************************************
    With oWord.Selection.Find
    Dim Text As String
    Text = "<<vFiador" & i + 1 & ">>"
        Do While (.Execute(findtext:="<<vFiador" & i + 1 & ">>", Forward:=True) = True) = True
    oWord.Selection.HomeKey Unit:=wdLine
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
        Loop
    End With
    

    With oWord.Selection.Find
        Do While (.Execute(findtext:="<<vFirma" & i + 1 & ">>", Forward:=True) = True) = True
    oWord.Selection.HomeKey Unit:=wdLine
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
        Loop
    End With
    

    With oWord.Selection.Find
        Do While (.Execute(findtext:="<<cNombres" & i + 1 & ">>", Forward:=True) = True) = True
    oWord.Selection.HomeKey Unit:=wdLine
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
        Loop
    End With
    

    With oWord.Selection.Find
        Do While (.Execute(findtext:="<<vDOI" & i + 1 & ">>", Forward:=True) = True) = True
    oWord.Selection.HomeKey Unit:=wdLine
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
        Loop
    End With
    

    With oWord.Selection.Find
    Do While (.Execute(findtext:="<<cDOI" & i + 1 & ">>", Forward:=True) = True) = True
    oWord.Selection.HomeKey Unit:=wdLine
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
        Loop
    End With
    

    With oWord.Selection.Find
    Do While (.Execute(findtext:="<<vDireccion" & i + 1 & ">>", Forward:=True) = True) = True
    oWord.Selection.HomeKey Unit:=wdLine
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
        Loop
    End With
    

    With oWord.Selection.Find
    Do While (.Execute(findtext:="<<cDireccion" & i + 1 & ">>", Forward:=True) = True) = True
    oWord.Selection.HomeKey Unit:=wdLine
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
    oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    oWord.Selection.Delete Unit:=wdCharacter, count:=1
        Loop
    End With

    With oWord.Selection.Find
         .Text = "<<vConyugeCodeudor1>>"
         .Replacement.Text = ""
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
         .Text = "<<Representante" & J + 1 & ">>"
         .Replacement.Text = ""
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
' ************************************************ FIN GEMO *********************************************
    
    
'    With oWord.Selection.Find
'         .Text = "<<cNomCli2>>"
'         .Replacement.Text = ""
'         .Forward = True
'         .Wrap = wdFindContinue
'         .Format = False
'         .Execute Replace:=wdReplaceAll
'    End With
'
'    With oWord.Selection.Find
'         .Text = "<<cNuDoci2>>"
'         .Replacement.Text = ""
'         .Forward = True
'         .Wrap = wdFindContinue
'         .Format = False
'         .Execute Replace:=wdReplaceAll
'    End With
'
'
'    With oWord.Selection.Find
'         .Text = "<<cAval1>>"
'         .Replacement.Text = ""
'         .Forward = True
'         .Wrap = wdFindContinue
'         .Format = False
'         .Execute Replace:=wdReplaceAll
'    End With
'
'    With oWord.Selection.Find
'         .Text = "<<cDocAval1>>"
'         .Replacement.Text = ""
'         .Forward = True
'         .Wrap = wdFindContinue
'         .Format = False
'         .Execute Replace:=wdReplaceAll
'    End With
'
'    With oWord.Selection.Find
'         .Text = "<<cAval2>>"
'         .Replacement.Text = ""
'         .Forward = True
'         .Wrap = wdFindContinue
'         .Format = False
'         .Execute Replace:=wdReplaceAll
'    End With
'
'    With oWord.Selection.Find
'         .Text = "<<cDocAval2>>"
'         .Replacement.Text = ""
'         .Forward = True
'         .Wrap = wdFindContinue
'         .Format = False
'         .Execute Replace:=wdReplaceAll
'    End With
'    With oWord.Selection.Find
'         .Text = "<<cDocAvalDir1>>"
'         .Replacement.Text = ""
'         .Forward = True
'         .Wrap = wdFindContinue
'         .Format = False
'         .Execute Replace:=wdReplaceAll
'    End With
'        With oWord.Selection.Find
'         .Text = "<<cDocAvalDir2>>"
'         .Replacement.Text = ""
'         .Forward = True
'         .Wrap = wdFindContinue
'         .Format = False
'         .Execute Replace:=wdReplaceAll
'    End With
'    With oWord.Selection.Find
'         .Text = "<<DirTit2>>"
'         .Replacement.Text = ""
'         .Forward = True
'         .Wrap = wdFindContinue
'         .Format = False
'         .Execute Replace:=wdReplaceAll
'    End With
 
    
    With oWord.Selection.Find
        .Text = "<<Zona>>"
        .Replacement.Text = lsAgencia
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<fDay>>"
        .Replacement.Text = Format(gdFecSis, "dd")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<fMes>>"
        '.Replacement.Text = Format(gdFecSis, "mm") comentado por gemo 12/11/2019
        .Replacement.Text = UCase(MonthName(Month(gdFecSis))) 'agregado por gemo 12/11/2019
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<fAnio>>"
        .Replacement.Text = Format(gdFecSis, "yyyy")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    oDoc.Close
    Set oDoc = Nothing

    Set oWord = CreateObject("Word.Application")
    oWord.Visible = True
    Set oDoc = oWord.Documents.Open(sArchivo)
    Set oDoc = Nothing
    Set oWord = Nothing
End Sub
'END EAAS20190523 SEGUN Memorándum Nº 756-2019-GM-DI/CMACM
' END GEMO 12/11/2019 ACTA 141 - 2019 : MEJORAS EN EL CONTRATO MULTIPRODUCTO ACTIVO

'*** BRGO 20111125 ******************************************************
Public Function ImprimeActivatePeru(ByVal psCtaCod As String) As String

    Dim R As ADODB.Recordset
    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range
    Dim oCred As COMDCredito.DCOMCredDoc

    Dim sArchivo As String

    On Error GoTo ErrorCartasAutorInfoGas

    Set oCred = New COMDCredito.DCOMCredDoc
    Set R = oCred.ObtieneDatosActivatePeru(psCtaCod)
    Set oCred = Nothing

    If R.EOF And R.BOF Then
        'MsgBox "No existen Datos para el Reporte...", vbInformation, "Aviso"
        Exit Function
    End If

    Set oWord = CreateObject("Word.Application")
        oWord.Visible = False
        Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\DJYADENDAACONTRATO-PROGRAMREACTIVAPERU.doc")
    
    sArchivo = App.Path & "\FormatoCarta\CA_" & psCtaCod & "_" & Replace(Left(Time, 10), ":", "") & ".doc"
    oDoc.SaveAs (sArchivo)
    
    With oWord.Selection.Find
        .Text = "<<NCREDITO>>"
        .Replacement.Text = Trim(R!cCtaCod)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    If R!REPRESENTANTE <> "-" Then  ' si tiene representante
        With oWord.Selection.Find
            .Text = "<<EMPRESA>>"
            .Replacement.Text = Trim(R!cPersNombre)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
        With oWord.Selection.Find
            .Text = "<<EMPRESA2>>"
            .Replacement.Text = "de la empresa " & Trim(R!cPersNombre)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
        With oWord.Selection.Find
            .Text = "<<RUCDNI2>>"
            .Replacement.Text = "R.U.C. " & Trim(R!Ruc)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
        With oWord.Selection.Find
            .Text = "<<RUCDNI>>"
            .Replacement.Text = "R.U.C. " & Trim(R!Ruc)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
                
        With oWord.Selection.Find
            .Text = "<<EMPRESADOMICILIO>>"
            .Replacement.Text = Trim(R!cPersNegocioDireccion)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
        With oWord.Selection.Find
            .Text = "<<EMPRESADOMICILIO2>>"
            .Replacement.Text = Trim(R!cPersNegocioDireccion)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
        With oWord.Selection.Find
            .Text = "<<REPRESENTANTE>>"
            .Replacement.Text = Trim(R!REPRESENTANTE)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
        With oWord.Selection.Find
            .Text = "<<CONDICION>>"
            .Replacement.Text = "Representante Legal"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
        With oWord.Selection.Find
            .Text = "<<DNI>>"
            .Replacement.Text = Trim(R!DNIR)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
        With oWord.Selection.Find
            .Text = "<<ANT>>"
            .Replacement.Text = "la empresa"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
    Else
        With oWord.Selection.Find
            .Text = "<<EMPRESA>>"
            .Replacement.Text = Trim(R!cPersNombre)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
        With oWord.Selection.Find
            .Text = "<<EMPRESA2>>"
            .Replacement.Text = "del Sr(a). " & Trim(R!cPersNombre)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
        With oWord.Selection.Find
            .Text = "<<RUCDNI2>>"
            .Replacement.Text = "DNI N° " & Trim(R!DNI) & ", R.U.C. " & Trim(R!Ruc)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
        With oWord.Selection.Find
            .Text = "<<RUCDNI>>"
            .Replacement.Text = "DNI N° " & Trim(R!DNI) & ", R.U.C. " & Trim(R!Ruc)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
                
        With oWord.Selection.Find
            .Text = "<<EMPRESADOMICILIO>>"
            .Replacement.Text = Trim(R!cPersNegocioDireccion)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
        With oWord.Selection.Find
            .Text = "<<EMPRESADOMICILIO2>>"
            .Replacement.Text = Trim(R!cPersNegocioDireccion)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
        With oWord.Selection.Find
            .Text = "<<REPRESENTANTE>>"
            .Replacement.Text = "______________________"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
        With oWord.Selection.Find
            .Text = "<<CONDICION>>"
            .Replacement.Text = "______________________"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
        With oWord.Selection.Find
            .Text = "<<DNI>>"
            .Replacement.Text = "_________"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
        With oWord.Selection.Find
            .Text = "<<ANT>>"
            .Replacement.Text = "el Sr(a)."
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
    End If
    
    With oWord.Selection.Find
        .Text = "<<CUIDAD>>"
        .Replacement.Text = Trim(R!cubigeodes)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
        .Text = "<<FECHA>>"
        .Replacement.Text = Format(gdFecSis, "d") & " días del mes de " & Format(gdFecSis, "mmmm") & " del " & Year(gdFecSis)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
 
    oDoc.Close
    Set oDoc = Nothing

    Set oWord = CreateObject("Word.Application")
    oWord.Visible = True
    Set oDoc = oWord.Documents.Open(sArchivo)
    Set oDoc = Nothing
    Set oWord = Nothing

    Exit Function

ErrorCartasAutorInfoGas:
    MsgBox err.Description, vbCritical, "Aviso"
    Set oDoc = Nothing
    Set oWord = Nothing
End Function
'*** END BRGO ******************************************************
