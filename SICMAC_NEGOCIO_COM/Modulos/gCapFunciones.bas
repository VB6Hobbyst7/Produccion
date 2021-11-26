Attribute VB_Name = "gCapFunciones"
Option Explicit

Public Function CabeRepoCaptac(ByVal sCabe01 As String, ByVal sCabe02 As String, _
        ByVal nCarLin As Long, ByVal sSeccion As String, ByVal sTitRp1 As String, _
        ByVal sTitRp2 As String, ByVal sMoneda As String, ByVal sNumPag As String, _
        ByVal sNomAge As String, ByVal dFecSis As Date) As String

Dim sTit1 As String, sTit2 As String
Dim sMon As String
Dim sCad As String
sTit1 = "": sTit2 = ""

CabeRepoCaptac = ""


' Definición de Cabecera 1
sMoneda = IIf(sMoneda = "", String(10, " "), " - " & sMoneda)
sCad = UCase(Trim(sNomAge)) & sMoneda
sCabe01 = sCad & String(50 - Len(sCad), " ")
sCabe01 = sCabe01 & Space((IIf(nCarLin <= 36, 80, nCarLin) - 36) - (Len(Mid(Trim(sCabe01), 1, 45)) - 2))
sCabe01 = sCabe01 & "PAGINA: " & sNumPag
sCabe01 = sCabe01 & Space(5) & "FECHA: " & Format$(dFecSis, "dd/mm/yyyy")

' Definición de Cabecera 2
sCabe02 = sSeccion & String(19 - Len(sSeccion), " ")
sCabe02 = sCabe02 & Space((IIf(nCarLin <= 19, 100, nCarLin) - 19) - (Len(sCabe02) - 2))
sCabe02 = sCabe02 & "HORA :   " & Format$(Now(), "hh:mm:ss")

' Definición del Titulo del Reporte
sTit1 = String(Int((IIf(nCarLin <= Len(sTitRp1), Len(sTitRp1) + 1, nCarLin) - Len(sTitRp1)) / 2), " ") & sTitRp1
sTit2 = String(Int((IIf(nCarLin <= Len(sTitRp2), Len(sTitRp2) + 1, nCarLin) - Len(sTitRp2)) / 2), " ") & sTitRp2
    
CabeRepoCaptac = CabeRepoCaptac & sCabe01 & Chr$(10)
CabeRepoCaptac = CabeRepoCaptac & sCabe02 & Chr$(10)
CabeRepoCaptac = CabeRepoCaptac & sTit1 & Chr$(10)
CabeRepoCaptac = CabeRepoCaptac & sTit2
End Function
Public Function CabeRepoCaptacSM(ByVal sCabe01 As String, ByVal sCabe02 As String, _
        ByVal nCarLin As Integer, ByVal sSeccion As String, ByVal sTitRp1 As String, _
        ByVal sTitRp2 As String, ByVal sMoneda As String, ByVal sNumPag As String, _
        ByVal sNomAge As String, ByVal dFecSis As Date) As String

Dim sTit1 As String, sTit2 As String
Dim sMon As String
Dim sCad As String
sTit1 = "": sTit2 = ""

' Definición de Cabecera 1
sMoneda = IIf(sMoneda = "", String(10, " "), " - " & sMoneda)
sCad = UCase(Trim(sNomAge)) & sMoneda
sCabe01 = sCad & String(50 - Len(sCad), " ")
sCabe01 = sCabe01 & Space((nCarLin - 36) - (Len(Mid(Trim(sCabe01), 1, 45)) - 2))
sCabe01 = sCabe01 & "PAGINA: " & sNumPag
sCabe01 = sCabe01 & Space(5) & "FECHA: " & Format$(dFecSis, "dd/mm/yyyy")

' Definición de Cabecera 2
sCabe02 = sSeccion & String(19 - Len(sSeccion), " ")
sCabe02 = sCabe02 & Space((nCarLin - 19) - (Len(sCabe02) - 2))
sCabe02 = sCabe02 & "HORA :   " & Format$(Now(), "hh:mm:ss")

' Definición del Titulo del Reporte
sTit1 = String(Int((nCarLin - Len(sTitRp1)) / 2), " ") & sTitRp1
sTit2 = String(Int((nCarLin - Len(sTitRp2)) / 2), " ") & sTitRp2
    
CabeRepoCaptacSM = CabeRepoCaptacSM & Space(30) & sCabe01 & Chr$(10)
CabeRepoCaptacSM = CabeRepoCaptacSM & Space(30) & sCabe02 & Chr$(10)
CabeRepoCaptacSM = CabeRepoCaptacSM & sTit1 & Chr$(10)
CabeRepoCaptacSM = CabeRepoCaptacSM & sTit2
End Function


Public Function ConvierteTNAaTEA(ByVal nTasa As Double) As Double
ConvierteTNAaTEA = ((1 + nTasa / 36000) ^ 360 - 1) * 100
End Function

Public Function ConvierteTEAaTNA(ByVal nTasa As Double) As Double
ConvierteTEAaTNA = ((1 + nTasa / 100) ^ (1 / 360) - 1) * 36000
End Function


'ARCV : Comentar desde aca para compilar clases

'***Agregado por ELRO el 20130326, según TI-ERS011-2013
Public Function ImpreCartillaAHLote2(prsCuentas As ADODB.Recordset, Optional nTipoAhorro As Integer = 0, Optional pcInstConvDep As String = "", Optional psMovNro As String = "")
    Dim objCaptaGenerales As COMDCaptaGenerales.DCOMCaptaGenerales
    Set objCaptaGenerales = New COMDCaptaGenerales.DCOMCaptaGenerales
    Dim objCaptac As COMNCaptaGenerales.NCOMCaptaGenerales
    Set objCaptac = New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim rsTitulares As ADODB.Recordset
    Set rsTitulares = New ADODB.Recordset

    Dim rs As ADODB.Recordset
    Dim nValor As Double
    Dim nMonMinCta As String
    Dim nMonxConsul As String
    Dim nMonComRet As String
    Dim nCosInac As Double
    Dim nMonOtraPlaza As String
    Dim nTasaITF As Double
    Dim nMonMinRetMN As String
    Dim nMonMinRetME As String
    Dim nParCod  As Long
    Dim lsFechas As String
    Dim lsNom1, lsNom2, lsNom3, lsNom4 As String
    Dim lsDoc1, lsDoc2, lsDoc3, lsDoc4 As String
    Dim lsDir1, lsDir2, lsDir3, lsDir4 As String
    Dim lsCad As String
    Dim nitf As Double
    Dim lsModeloPlantilla As String
    Dim i As Integer
    Dim sTipoEnvio As String 'APRI20180420 ERS036-2017

    Dim loRs As COMDConstSistema.DCOMGeneral

      On Error GoTo Error 'ADD PTI1 10/09/2018 INFORME N° 003-2017-AC-TI / CMACM

    Set loRs = New COMDConstSistema.DCOMGeneral


    Dim loAge As COMDConstantes.DCOMAgencias
    Dim rs1 As ADODB.Recordset
    Dim lsAgencia As String
    Dim lsAgeDir As String
    Dim nTREA As Currency

    Set loAge = New COMDConstantes.DCOMAgencias
    Set rs1 = New ADODB.Recordset
        Set rs1 = loAge.RecuperaAgencias(gsCodAge)
        If Not (rs1.EOF And rs1.BOF) Then
            lsAgencia = Trim(rs1("cUbiGeoDescripcion"))
            lsAgeDir = rs1("cAgeDireccion")
        End If

    'JUEZ 20150121 Nuevos Parámetros *******************
    Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
    Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
    Dim rsPar As ADODB.Recordset
    Set rsPar = clsDef.GetCapParametroNew(gCapAhorros, 0)
    'END JUEZ ******************************************
     'APRI20190109 ERS077-2018
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim nSaldoEquilibrio As Currency
    Dim nCostoMan As Currency
    'END APRI

    '------Saldo x Cuenta a Mantener por 3 Meses------
     prsCuentas.MoveFirst
     If Mid(prsCuentas!Cuenta, 9, 1) = 1 Then
        nParCod = 2091
     Else
        nParCod = 2092
     End If


     'nValor = loRs.GetParametro(2000, nParCod)
     nValor = 0  'IIf(Mid(prsCuentas!Cuenta, 9, 1) = 1, rsPar!nSaldoMinCtaSol, rsPar!nSaldoMinCtaDol) 'JUEZ 20150121 'APRI20190109 ERS077-2018
     '''nMonMinCta = IIf(Mid(prsCuentas!Cuenta, 9, 1) = 1, "S/. ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(prsCuentas!Cuenta, 9, 1) = 1, " (" & UnNumero(nValor) & " Nuevos Soles)", " (" & UnNumero(nValor) & " DOLARES)") 'marg ers044-2016
     nMonMinCta = IIf(Mid(prsCuentas!Cuenta, 9, 1) = 1, gcPEN_SIMBOLO & " ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(prsCuentas!Cuenta, 9, 1) = 1, " (" & UnNumero(nValor) & " " & StrConv(gcPEN_PLURAL, vbProperCase) & ")", " (" & UnNumero(nValor) & " DOLARES)") 'marg ers044-2016
     '----Fin Saldo x Cuenta a Mantener por 3 Meses----
     '------Comision Consulta Saldo Ventanilla------
       nParCod = 2106
     nValor = loRs.GetParametro(2000, nParCod)
     '''nMonxConsul = "S/." & Format$(nValor, "0.00") & " (" & UnNumero(nValor) & " Nuevos Soles)" 'marg ers044-2016
     nMonxConsul = gcPEN_SIMBOLO & Format$(nValor, "0.00") & " (" & UnNumero(nValor) & " " & StrConv(gcPEN_PLURAL, vbProperCase) & ")" 'marg ers044-2016
     '----Fin Comision Consulta Saldo Ventanilla----
     '-----Comision de retiro cuando el monto es menor a 1000-----
     Set rs = New ADODB.Recordset
     If Mid(prsCuentas!Cuenta, 9, 1) = 1 Then
        nParCod = gMontoMNx31Ope
     Else
        nParCod = gMontoMEx31Ope
     End If
     nValor = loRs.GetParametro(2000, nParCod)
     '''nMonComRet = IIf(Mid(prsCuentas!Cuenta, 9, 1) = 1, "S/. ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(prsCuentas!Cuenta, 9, 1) = 1, " (" & UnNumero(nValor) & " Nuevos Soles)", " (" & UnNumero(nValor) & " Dolares)") 'marg ers044-2016
     nMonComRet = IIf(Mid(prsCuentas!Cuenta, 9, 1) = 1, gcPEN_SIMBOLO & " ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(prsCuentas!Cuenta, 9, 1) = 1, " (" & UnNumero(nValor) & " " & StrConv(gcPEN_PLURAL, vbProperCase) & ")", " (" & UnNumero(nValor) & " Dolares)") 'marg ers044-2016
     '----Fin Comision de retiro cuando el monto es menor a 1000-----
     '-----------Costo de Inactivas-------------------
     nValor = loRs.GetParametro(2000, gMonDescInacME)
     nCosInac = Format$(nValor, "0.00")
     '---------Fin Costo de Inactivas-----------------

    '------Retiros/Depositos en Agencias de La Caja en Otras Plazas------
    If Mid(prsCuentas!Cuenta, 9, 1) = 1 Then
        nParCod = 2046
    Else
        nParCod = 2047
    End If
    nValor = loRs.GetParametro(2000, nParCod)
    '''nMonOtraPlaza = IIf(Mid(prsCuentas!Cuenta, 9, 1) = 1, "S/. ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(prsCuentas!Cuenta, 9, 1) = 1, " (" & UnNumero(nValor) & " Nuevos Soles)", " (" & UnNumero(nValor) & " Dolares)") 'marg ers044-2016
    nMonOtraPlaza = IIf(Mid(prsCuentas!Cuenta, 9, 1) = 1, gcPEN_SIMBOLO & " ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(prsCuentas!Cuenta, 9, 1) = 1, " (" & UnNumero(nValor) & " " & StrConv(gcPEN_PLURAL, vbProperCase) & ")", " (" & UnNumero(nValor) & " Dolares)") 'marg ers044-2016
    '----Fin Retiros/Depositos en Agencias de La Caja en Otras Plazas----
    '---------------------Monto Minimo de Retiro-------------------
    nParCod = 2027
    nValor = loRs.GetParametro(2000, nParCod)
    '''nMonMinRetMN = "S/. " & Format$(nValor, "0.00") & " (" & UnNumero(nValor) & " Nuevos Soles)" 'marg ers044-2016
    nMonMinRetMN = gcPEN_SIMBOLO & " " & Format$(nValor, "0.00") & " (" & UnNumero(nValor) & " " & StrConv(gcPEN_PLURAL, vbProperCase) & ")" 'marg ers044-2016
    nParCod = 2028
    nValor = loRs.GetParametro(2000, nParCod)
    nMonMinRetME = "$. " & Format$(nValor, "0.00") & " (" & UnNumero(nValor) & " Dolares)"
    '-------------------Fin Monto Minimo de Retiro-----------------

    Set loRs = Nothing
    Set loAge = Nothing
    Set rs1 = Nothing

    If nTipoAhorro = 0 Then
        lsModeloPlantilla = App.Path & "\FormatoCarta\CARTILLAAHORROC.doc"
    ElseIf nTipoAhorro = 5 Then
        lsModeloPlantilla = App.Path & "\FormatoCarta\CARTILLAAHORROCUENTASONADA.doc"
    ElseIf nTipoAhorro = 6 Then
        lsModeloPlantilla = App.Path & "\FormatoCarta\CARTILLACAJASUELDO.doc"
    ElseIf nTipoAhorro = 8 Then
        lsModeloPlantilla = App.Path & "\FormatoCarta\CARTILLAAHORROCONVENIO.doc"
    End If

    'Crea una clase que de Word Object
    Dim wApp As Word.Application
    Dim wAppSource As Word.Application
    'Create a new instance of word
    Set wApp = New Word.Application
    Set wAppSource = New Word.Application

    Dim RangeSource As Word.Range
    'Abre Documento Plantilla
    wAppSource.Documents.Open FileName:=lsModeloPlantilla
    Set RangeSource = wAppSource.ActiveDocument.Content
    'Lo carga en Memoria
    wAppSource.ActiveDocument.Content.Copy

    'Crea Nuevo Documento
    wApp.Documents.Add

    With wApp.ActiveDocument.PageSetup
'        .LeftMargin = CentimetersToPoints(1.5)
'        .RightMargin = CentimetersToPoints(1)
'        .TopMargin = CentimetersToPoints(1.5)
'        .BottomMargin = CentimetersToPoints(1)

    '   ADD PTI1 10/09/2018 INFORME N° 003-2017-AC-TI / CMACM
         .PaperSize = wAppSource.ActiveDocument.PageSetup.PaperSize
         .PageHeight = wAppSource.ActiveDocument.PageSetup.PageHeight
         .PageWidth = wAppSource.ActiveDocument.PageSetup.PageWidth

         .HeaderDistance = wAppSource.ActiveDocument.PageSetup.HeaderDistance
         .FooterDistance = wAppSource.ActiveDocument.PageSetup.FooterDistance

         .LeftMargin = wAppSource.ActiveDocument.PageSetup.LeftMargin
         .RightMargin = wAppSource.ActiveDocument.PageSetup.RightMargin
         .TopMargin = wAppSource.ActiveDocument.PageSetup.TopMargin
         .BottomMargin = wAppSource.ActiveDocument.PageSetup.BottomMargin
     '  END AGREGADO PTI1
    End With

    With wApp.Selection.Font
        .Name = "Arial Narrow"
        .Size = 12
    End With

    prsCuentas.MoveFirst
    Do While Not prsCuentas.EOF

    'APRI2018 ERS036-2017
    Set rs = objCaptaGenerales.RecuperaDatosEnvioEstadoCta(prsCuentas!Cuenta)
        If Not (rs.BOF And rs.EOF) Then
        If rs.RecordCount > 0 Then
          sTipoEnvio = objCaptaGenerales.MostarTextoComisionEnvioEstAhorros(rs!nModoEnvio, nTipoAhorro)
        End If
        Else
         sTipoEnvio = ""
        End If
     Set rs = Nothing
    'END APRI

    Set rsTitulares = objCaptaGenerales.devolverTitularCuenta(prsCuentas!Cuenta)
    'APRI20190109 ERS077-2018
    Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    nCostoMan = clsDef.GetCapCostoMantenimiento(232, nTipoAhorro, Mid(prsCuentas!Cuenta, 9, 1), prsCuentas!Monto)
    nSaldoEquilibrio = clsMant.ObtenerSaldoEquilibrio(nCostoMan, prsCuentas("TEA Destino"))
    Set clsDef = Nothing
    Set clsMant = Nothing
    'END APRI
    'nTREA = objCaptac.ObtenerTREA(Mid$(prsCuentas!Cuenta, 6, 3), IIf(prsCuentas!Monto = 0, 10, prsCuentas!Monto), prsCuentas("TEA Destino"), 0, 0)
    nTREA = objCaptac.ObtenerTREA(Mid$(prsCuentas!Cuenta, 6, 3), IIf(prsCuentas!Monto = 0, 10, prsCuentas!Monto), prsCuentas("TEA Destino"), 0, nCostoMan) 'APRI20190109 ERS077-2018

    wApp.Application.Selection.TypeParagraph
    wApp.Application.Selection.PasteAndFormat (wdPasteDefault)
    wApp.Application.Selection.InsertBreak
    wApp.Selection.SetRange start:=wApp.Selection.start, End:=wApp.ActiveDocument.Content.End
    wApp.Selection.MoveEnd

    If nTipoAhorro = 8 Then
        With wApp.Selection.Find
            .Text = "<<Nombre Empresa-Convenio>>"
            .Replacement.Text = pcInstConvDep
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
    End If
    'END APRI *****************************************************

    'APRI20180420 ERS036-2017************************************************
    'Cuenta
    lsCad = prsCuentas!Cuenta
    With wApp.Selection.Find
        .Text = "<<cCodCta>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END APRI *****************************************************

    lsCad = ""
    lsCad = Format(nTREA, "0.00") & " % (Fija)"
    With wApp.Selection.Find
        .Text = "<<TasaTrea>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'APRI20190109 ERS077-2018
    lsCad = ""
    lsCad = Format(nSaldoEquilibrio, "0.00")
    With wApp.Selection.Find
        .Text = "<<SaldoEquilibrio>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END APRI
    lsCad = ""
    lsCad = IIf(Mid(prsCuentas!Cuenta, 9, 1) = 1, "MN ", "ME ") & Format(prsCuentas!Monto, "0.00")
    With wApp.Selection.Find
        .Text = "<<Monto>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    lsCad = ""
    '''lsCad = Format(prsCuentas("TEA Destino"), "0.00") & " % " & IIf(Mid(prsCuentas!Cuenta, 9, 1) = 1, "Soles", "Dolares") 'marg ers044-2016
    lsCad = Format(prsCuentas("TEA Destino"), "0.00") & " % " & IIf(Mid(prsCuentas!Cuenta, 9, 1) = 1, StrConv(gcPEN_PLURAL, vbProperCase), "Dolares") 'marg ers044-2016
    With wApp.Selection.Find
        .Text = "<<Tasa>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Saldo de la cuenta por 3 meses
    With wApp.Selection.Find
        .Text = "<<MontoMin>>"
        .Replacement.Text = nMonMinCta
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Monto x Consulta
    With wApp.Selection.Find
        .Text = "<<MonConsul>>"
        .Replacement.Text = nMonxConsul
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Comision por Retiro
    With wApp.Selection.Find
        .Text = "<<MonComRet>>"
        .Replacement.Text = nMonComRet
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Costo de Inactivas
    With wApp.Selection.Find
        .Text = "<<CostoInac>>"
        .Replacement.Text = nCosInac
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Monto de Otra Plaza
    With wApp.Selection.Find
        .Text = "<<MonOtraPlaza>>"
        .Replacement.Text = nMonOtraPlaza
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Monto de Otra Plaza
    With wApp.Selection.Find
        .Text = "<<MonOtraPlaza>>"
        .Replacement.Text = nMonOtraPlaza
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Monto Minimo de Retiro en Soles
    With wApp.Selection.Find
        .Text = "<<MonMinRetMN>>"
        .Replacement.Text = nMonMinRetMN
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'ITF
    nitf = gnITFPorcent * 100
    With wApp.Selection.Find
        .Text = "<<TasaITF>>"
        .Replacement.Text = Trim(CStr(nitf))
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Monto Minimo de Retiro en Dolares
    With wApp.Selection.Find
        .Text = "<<MonMinRetME>>"
        .Replacement.Text = nMonMinRetME
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'APRI20180420 ERS036-2017
    'Tipo envio estado de cuenta
    With wApp.Selection.Find
        .Text = "<<TipoEnvio>>"
        .Replacement.Text = sTipoEnvio
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END APRI
    'Fecha
    lsFechas = Format(gdFecSis, "dd") & " de " & Format(gdFecSis, "mmmm") & " del " & Format(gdFecSis, "yyyy")
    With wApp.Selection.Find
        .Text = "<<FecActual>>"
        .Replacement.Text = lsFechas
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Lugar
    With wApp.Selection.Find
        .Text = "<<cLugar>>"
        .Replacement.Text = lsAgencia
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Direccion
    With wApp.Selection.Find
        .Text = "<<cDireccion>>"
        .Replacement.Text = lsAgeDir
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    i = 1
    lsNom1 = ""
    lsDoc1 = ""
    lsDir1 = ""
    lsNom2 = ""
    lsDoc2 = ""
    lsDir2 = ""
    lsNom3 = ""
    lsDoc3 = ""
    lsDir3 = ""
    lsNom4 = ""
    lsDoc4 = ""
    lsDir4 = ""
    Do While Not rsTitulares.EOF
        If i = 1 Then
            lsNom1 = "Nombre del Cliente: " & rsTitulares!cPersNombre
            lsDoc1 = "DNI/RUC: " & rsTitulares!cPersIDNro & Space(60) & "Firma: ______________________"
            lsDir1 = "Dirección: " & rsTitulares!cPersDireccDomicilio
            With wApp.Selection.Find
                .Text = "<<NomTit1>>"
                .Replacement.Text = lsNom1
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With wApp.Selection.Find
                .Text = "<<DocTit1>>"
                .Replacement.Text = lsDoc1
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With wApp.Selection.Find
                .Text = "<<DirTit1>>"
                .Replacement.Text = lsDir1
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
        ElseIf i = 2 Then
            lsNom2 = "Nombre del Cliente: " & rsTitulares!cPersNombre
            lsDoc2 = "DNI/RUC: " & rsTitulares!cPersIDNro & Space(60) & "Firma: ______________________"
            lsDir2 = "Dirección: " & rsTitulares!cPersDireccDomicilio
            With wApp.Selection.Find
                .Text = "<<NomTit2>>"
                .Replacement.Text = lsNom2
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With wApp.Selection.Find
                .Text = "<<DocTit2>>"
                .Replacement.Text = lsDoc2
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With wApp.Selection.Find
                .Text = "<<DirTit2>>"
                .Replacement.Text = lsDir2
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
        ElseIf i = 3 Then
            lsNom3 = "Nombre del Cliente: " & rsTitulares!cPersNombre
            lsDoc3 = "DNI/RUC: " & rsTitulares!cPersIDNro & Space(60) & "Firma: ______________________"
            lsDir3 = "Dirección: " & rsTitulares!cPersDireccDomicilio
            With wApp.Selection.Find
                .Text = "<<NomTit3>>"
                .Replacement.Text = lsNom3
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With wApp.Selection.Find
                .Text = "<<DocTit3>>"
                .Replacement.Text = lsDoc3
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With wApp.Selection.Find
                .Text = "<<DirTit3>>"
                .Replacement.Text = lsDir3
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
        ElseIf i = 4 Then
            lsNom4 = "Nombre del Cliente: " & rsTitulares!cPersNombre
            lsDoc4 = "DNI/RUC: " & rsTitulares!cPersIDNro & Space(60) & "Firma: ______________________"
            lsDir4 = "Dirección: " & rsTitulares!cPersDireccDomicilio
            With wApp.Selection.Find
                .Text = "<<NomTit4>>"
                .Replacement.Text = lsNom4
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With wApp.Selection.Find
                .Text = "<<DocTit4>>"
                .Replacement.Text = lsDoc4
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With wApp.Selection.Find
                .Text = "<<DirTit4>>"
                .Replacement.Text = lsDir4
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
        End If
        i = i + 1
        rsTitulares.MoveNext
    Loop

     With wApp.Selection.Find
        .Text = "<<NomTit1>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DocTit1>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DirTit1>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With wApp.Selection.Find
        .Text = "<<NomTit2>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DocTit2>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DirTit2>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With wApp.Selection.Find
        .Text = "<<NomTit3>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DocTit3>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DirTit3>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With wApp.Selection.Find
        .Text = "<<NomTit4>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DocTit4>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DirTit4>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

Set rsTitulares = Nothing
prsCuentas.MoveNext
Loop
Set prsCuentas = Nothing

wAppSource.ActiveDocument.Close
wAppSource.Quit
wApp.ActiveDocument.SaveAs (App.Path & "\SPOOLER\AperturaLoteAhorro_" & psMovNro & ".doc")
wApp.Visible = True

  'ADD PTI1 10/09/2018 INFORME N° 003-2017-AC-TI / CMACM
Exit Function
Error:
MsgBox "Error al generar cartillas", vbInformation, "Aviso"

End Function
'***Fin Agregado por ELRO el 20130326******************

'*** Funcion de Impresion de Cartillas en Lote segun Normas de SBS - AVMM - 10-08-2006 ***
Public Function ImprimeCartilla(MatTitular() As String, ByVal pnTipo As Integer, ByVal psCtaCod As String, ByVal pnTasa As Double, _
                           ByVal pnMonto As Double, ByVal pdFechaA As Date, Optional pnPlazo As Integer, Optional bOrdPag As Boolean = False, Optional pnITF As Integer, Optional pnMontoITf As Double) As String
' pnTipo=1--> Ahorro /pnTipo=2--> Plazo Fijo /pnTipo=3--> CTS
Dim lsCad As String
Dim nInteres As Double
Dim nMeses As Integer
Dim OiMO As Impresoras
Dim i As Byte
If pnITF = 0 Then
    pnMonto = pnMonto - pnMontoITf
Else
    pnMonto = pnMonto
End If
nInteres = pnMonto * (pnTasa / 100) * 30
lsCad = lsCad & oImpresora.gPrnCondensadaON
lsCad = lsCad & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
lsCad = lsCad & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
If pnTipo <> 1 Then lsCad = lsCad & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
If pnTipo = 1 Then
lsCad = lsCad & oImpresora.gPrnSaltoLinea
'IMPRESION CARTILLA DE AHORROS
    If bOrdPag = False Then
        lsCad = lsCad & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    End If
    lsCad = lsCad & Space(10) & "1. Cuenta : " & Mid(psCtaCod, 1, 3) & "-" & Mid(psCtaCod, 4, 2) & "-" & Mid(psCtaCod, 6, 3) & "-" & Mid(psCtaCod, 9, 10) & Space(3) & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "2. Monto  : " & Format(pnMonto, "0.00") & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "3. T.E.A  : " & Format(pnTasa, "0.00") & "%" & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "4. Fecha de Corte para el Abono de Intereses " & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(13) & "Capitalizacion : " & CDate(pdFechaA) + 30 & Space(2) & "Capitalizacion : " & CDate(pdFechaA) + 60 & Space(2) & "Capitalizacion : " & CDate(pdFechaA) + 90 & oImpresora.gPrnSaltoLinea
    '''lsCad = lsCad & Space(10) & "5. El saldo minimo en cuenta de ahorros sera de S/. 5.00 Nuevos Soles o US$. 1.50 Dolares Americanos" & oImpresora.gPrnSaltoLinea 'marg ers044-2016
    lsCad = lsCad & Space(10) & "5. El saldo minimo en cuenta de ahorros sera de " & gcPEN_SIMBOLO & " 5.00 " & StrConv(gcPEN_PLURAL, vbProperCase) & " o US$. 1.50 Dolares Americanos" & oImpresora.gPrnSaltoLinea  'marg ers044-2016
    lsCad = lsCad & Space(10) & "6. Comisiones y/o gastos aplicables: " & oImpresora.gPrnSaltoLinea
    '''lsCad = lsCad & Space(13) & "-Apartir del retiro Nro 31 se cargara a la cuenta un S/. 1.00 nuevos soles  o US$ 0.5 Dolares Americanos" & oImpresora.gPrnSaltoLinea 'marg ers044-2016
    lsCad = lsCad & Space(13) & "-Apartir del retiro Nro 31 se cargara a la cuenta un " & gcPEN_SIMBOLO & " 1.00 " & StrConv(gcPEN_PLURAL, vbLowerCase) & "  o US$ 0.5 Dolares Americanos" & oImpresora.gPrnSaltoLinea 'marg ers044-2016
    lsCad = lsCad & Space(14) & "durante el mes" & oImpresora.gPrnSaltoLinea
    '''lsCad = lsCad & Space(13) & "-Retiros y Depositos en otras plaza por montos mayores a S/. 5,000.00 Nuevos Soles o US$ 1,500.00 Dolares" & oImpresora.gPrnSaltoLinea 'marg ers044-2016
    lsCad = lsCad & Space(13) & "-Retiros y Depositos en otras plaza por montos mayores a " & gcPEN_SIMBOLO & " 5,000.00 " & StrConv(gcPEN_PLURAL, vbProperCase) & " o US$ 1,500.00 Dolares" & oImpresora.gPrnSaltoLinea 'marg ers044-2016
    lsCad = lsCad & Space(14) & "Americanos tendra comision del 0.1 %" & oImpresora.gPrnSaltoLinea
    '''lsCad = lsCad & Space(13) & "-La cuentas inactivas genera una comisión semestral de S/. 5.00 Nuevos Soles o US$ 1.5 Dolares Americanos" & oImpresora.gPrnSaltoLinea 'marg ers044-2016
    lsCad = lsCad & Space(13) & "-La cuentas inactivas genera una comisión semestral de " & gcPEN_SIMBOLO & " 5.00 " & StrConv(gcPEN_PLURAL, vbProperCase) & " o US$ 1.5 Dolares Americanos" & oImpresora.gPrnSaltoLinea 'marg ers044-2016
    lsCad = lsCad & Space(10) & "7. Se aplicara la tasa de 0.08 % por Transaccion efectuada por concepto de Impuesto a las Transacciones " & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "   Financieras (ITF). Excepto en el caso de cuentas exoneradas u operaciones inafectas de acuerdo a Ley." & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    If bOrdPag = True Then
        For i = 1 To 16
            lsCad = lsCad & oImpresora.gPrnSaltoLinea
        Next i
    Else
        For i = 1 To 14
            lsCad = lsCad & oImpresora.gPrnSaltoLinea
        Next i
    End If
ElseIf pnTipo = 2 Then
'IMPRESION CARTILLA DE PLAZO FIJO
    Dim sTaInEf As String
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
    Dim lnPlazo As Long
    lnPlazo = CLng(pnPlazo)
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    sTaInEf = Format$(((((((pnTasa / 12) / 3000) + 1) ^ lnPlazo) - 1) * 100), "#0.00")
    nInteres = Format$(clsMant.GetInteresPF(pnTasa, pnMonto, lnPlazo), "#,##0.00")
    Set clsMant = Nothing
    lsCad = lsCad & Space(10) & "1. Cuenta : " & Mid(psCtaCod, 1, 3) & "-" & Mid(psCtaCod, 4, 2) & "-" & Mid(psCtaCod, 6, 3) & "-" & Mid(psCtaCod, 9, 10) & Space(3) & "3. T.E.A : " & Format(sTaInEf, "0.00") & "%" & Space(2) & "Plazo : " & pnPlazo & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "2. Monto  : " & Format(pnMonto, "0.00") & Space(8) & "4. Monto Total : " & nInteres & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "5. Fecha de Corte para el Abono de Intereses: " & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(13) & "Apertura  : " & pdFechaA & oImpresora.gPrnSaltoLinea
    nMeses = pnPlazo / 30
    If nMeses <= 2 Then
        If nMeses = 1 Then
            lsCad = lsCad & Space(13) & "Provision : " & CDate(pdFechaA) + 30 & oImpresora.gPrnSaltoLinea
        Else
            lsCad = lsCad & Space(13) & "Provision : " & CDate(pdFechaA) + 30 & Space(2) & "Provision : " & CDate(pdFechaA) + pnPlazo & oImpresora.gPrnSaltoLinea
        End If
    Else
        If nMeses = 3 Then
            lsCad = lsCad & Space(13) & "Provision : " & CDate(pdFechaA) + 30 & Space(2) & "Provision : " & CDate(pdFechaA) + 60 & Space(2) & "Provision : " & CDate(pdFechaA) + pnPlazo & oImpresora.gPrnSaltoLinea
        Else
            lsCad = lsCad & Space(13) & "Provision : " & CDate(pdFechaA) + 30 & Space(2) & "Provision : " & CDate(pdFechaA) + 60 & Space(2) & "Provision : " & CDate(pdFechaA) + 90 & Space(2) & "...Provision : " & CDate(pdFechaA) + pnPlazo & oImpresora.gPrnSaltoLinea
        End If
    End If
    lsCad = lsCad & Space(10) & "6. Fecha de Vencimiento : " & CDate(pdFechaA) + pnPlazo & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "7. Los intereses podran ser retirados cada 30 días" & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "8. Se aplicara la tasa de 0.08 % por Transaccion efectuada por concepto de Impuesto a las Transacciones " & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "   Financieras (ITF). Excepto en el caso de cuentas exoneradas u operaciones inafectas de acuerdo a Ley. " & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    For i = 1 To 23
        lsCad = lsCad & oImpresora.gPrnSaltoLinea
    Next i
Else
'IMPRESION CARTILLA CTS
    lsCad = lsCad & Space(10) & "1. Cuenta : " & Mid(psCtaCod, 1, 3) & "-" & Mid(psCtaCod, 4, 2) & "-" & Mid(psCtaCod, 6, 3) & "-" & Mid(psCtaCod, 9, 10) & Space(3) & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "2. Monto  : " & Format(pnMonto, "0.00") & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "3. T.E.A : " & Format(pnTasa, "0.00") & "%" & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "4. Fecha de Corte para el Abono de Intereses " & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(13) & "Capitalizacion : " & CDate(pdFechaA) + 30 & Space(2) & "Capitalizacion : " & CDate(pdFechaA) + 60 & Space(2) & "Capitalizacion : " & CDate(pdFechaA) + 90 & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "5. Los Retiros se efectuaran conforme a los porcentajes de Disponibilidad establecida por ley." & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    For i = 1 To 21
        lsCad = lsCad & oImpresora.gPrnSaltoLinea
    Next i


End If

For i = 1 To 13
        lsCad = lsCad & oImpresora.gPrnSaltoLinea
Next i
'IMPRESION DE FIRMAS
If UBound(MatTitular) = 1 Then
    lsCad = lsCad & Space(10) & "---------------------------------------" & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & MatTitular(1) & oImpresora.gPrnSaltoLinea

ElseIf UBound(MatTitular) = 2 Then
    lsCad = lsCad & Space(10) & "---------------------------------------" & Space(10) & "---------------------------------------" & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & ImpreFormat(Left(MatTitular(1), 39), 39) & Space(10) & ImpreFormat(Left(MatTitular(2), 39), 39) & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & Left(MatTitular(1), 30) & Space(10) & Left(MatTitular(2), 30) & oImpresora.gPrnSaltoLinea

ElseIf UBound(MatTitular) = 3 Then
    lsCad = lsCad & Space(10) & "---------------------------------------" & Space(10) & "---------------------------------------" & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & ImpreFormat(Left(MatTitular(1), 39), 39) & Space(10) & ImpreFormat(Left(MatTitular(2), 39), 39) & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & Left(MatTitular(1), 30) & Space(10) & Left(MatTitular(2), 30) & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "---------------------------------------" & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & ImpreFormat(Left(MatTitular(3), 39), 39) & oImpresora.gPrnSaltoLinea

ElseIf UBound(MatTitular) = 4 Then
    lsCad = lsCad & Space(10) & "---------------------------------------" & Space(10) & "---------------------------------------" & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & ImpreFormat(Left(MatTitular(1), 39), 39) & Space(10) & ImpreFormat(Left(MatTitular(2), 39), 39) & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & Left(MatTitular(1), 30) & Space(10) & Left(MatTitular(2), 30) & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "---------------------------------------" & Space(10) & "---------------------------------------" & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & ImpreFormat(Left(MatTitular(3), 39), 39) & Space(10) & ImpreFormat(Left(MatTitular(4), 39), 39) & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & Left(MatTitular(3), 30) & Space(10) & Left(MatTitular(4), 30) & oImpresora.gPrnSaltoLinea
End If
lsCad = lsCad & oImpresora.gPrnCondensadaOFF
ImprimeCartilla = lsCad
End Function

Public Function ImpreCartillaAHLote(MatTitular() As String, MatNroCta() As String, Optional nTipoAhorro As Integer = 0, Optional pcInstConvDep As String = "", Optional psMovNro As String = "")

    Dim rs As ADODB.Recordset
    Dim nValor As Double
    Dim nMonMinCta As String
    Dim nMonxConsul As String
    Dim nMonComRet As String
    Dim nCosInac As Double
    Dim nMonOtraPlaza As String
    Dim nTasaITF As Double
    Dim nMonMinRetMN As String
    Dim nMonMinRetME As String
    Dim nParCod  As Long
    Dim lsFechas As String
    Dim lsNom1 As String
    Dim lsDoc1 As String
    Dim lsDir1 As String
    Dim lsCad As String
    Dim nitf As Double
    Dim lsModeloPlantilla As String
    Dim i As Integer

    Dim loRs As COMDConstSistema.DCOMGeneral
    Set loRs = New COMDConstSistema.DCOMGeneral


    Dim loAge As COMDConstantes.DCOMAgencias
    Dim rs1 As ADODB.Recordset
    Dim lsAgencia As String
    Dim lsAgeDir As String

    Set loAge = New COMDConstantes.DCOMAgencias
    Set rs1 = New ADODB.Recordset
        Set rs1 = loAge.RecuperaAgencias(gsCodAge)
        If Not (rs1.EOF And rs1.BOF) Then
            lsAgencia = Trim(rs1("cUbiGeoDescripcion"))
            lsAgeDir = rs1("cAgeDireccion")
        End If
    Set loAge = Nothing

    'JUEZ 20150121 Nuevos Parámetros *******************
    Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
    Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
    Dim rsPar As ADODB.Recordset
    Set rsPar = clsDef.GetCapParametroNew(gCapAhorros, 0)
    'END JUEZ ******************************************
    'APRI20190109 ERS077-2018
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim nSaldoEquilibrio As Currency
    Dim nCostoMan As Currency
    'END APRI
    '--------Saldo x Cuenta a Mantener por 3 Meses --

    If Mid(MatNroCta(1), 9, 1) = 1 Then
       nParCod = 2091
    Else
       nParCod = 2092
    End If
    'nValor = loRs.GetParametro(2000, nParCod)
    nValor = 0 'IIf(Mid(MatNroCta(1), 9, 1) = 1, rsPar!nSaldoMinCtaSol, rsPar!nSaldoMinCtaDol) 'JUEZ 20150121 'APRI20190109 ERS077-2018
    '''nMonMinCta = IIf(Mid(MatNroCta(1), 9, 1) = 1, "S/. ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(MatNroCta(1), 9, 1) = 1, " (" & UnNumero(nValor) & " Nuevos Soles)", " (" & UnNumero(nValor) & " DOLARES)") 'marg ers044-2016
    nMonMinCta = IIf(Mid(MatNroCta(1), 9, 1) = 1, gcPEN_SIMBOLO & " ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(MatNroCta(1), 9, 1) = 1, " (" & UnNumero(nValor) & " " & StrConv(gcPEN_PLURAL, vbProperCase) & ")", " (" & UnNumero(nValor) & " DOLARES)") 'marg ers044-2016

    '------------------------------------------------

    '-----Monto de Consulta por extracto de Cuenta --
    'By Capi 05032008

'    Set rs = New ADODB.Recordset
'    If Mid(MatNroCta(1), 9, 1) = 1 Then
'       nParCod = gDctoExtMNxPag
'    Else
'       nParCod = gDctoExtMExPag
'    End If
'    nValor = loRs.GetParametro(2000, nParCod)
'    nMonxConsul = IIf(Mid(MatNroCta(1), 9, 1) = 1, "S/. ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(MatNroCta(1), 9, 1) = 1, " (" & UnNumero(nValor) & " Nuevos Soles)", " (" & UnNumero(nValor) & " Dolares)")

    '------------------------------------------------
     '-----Comision Consulta Saldo Ventanilla --
    Set rs = New ADODB.Recordset
    nParCod = 2106
    nValor = loRs.GetParametro(2000, nParCod)
    '''nMonxConsul = "S/." & Format$(nValor, "0.00") & " (" & UnNumero(nValor) & " Nuevos Soles)" 'marg ers044-2016
    nMonxConsul = gcPEN_SIMBOLO & Format$(nValor, "0.00") & " (" & UnNumero(nValor) & " " & StrConv(gcPEN_PLURAL, vbProperCase) & ")" 'marg ers044-2016


    '-----Comision de retiro cuando el monto es menor a 1000 --
    Set rs = New ADODB.Recordset
    If Mid(MatNroCta(1), 9, 1) = 1 Then
       nParCod = gMontoMNx31Ope
    Else
       nParCod = gMontoMEx31Ope
    End If
    nValor = loRs.GetParametro(2000, nParCod)
    '''nMonComRet = IIf(Mid(MatNroCta(1), 9, 1) = 1, "S/. ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(MatNroCta(1), 9, 1) = 1, " (" & UnNumero(nValor) & " Nuevos Soles)", " (" & UnNumero(nValor) & " Dolares)") 'marg ers044-2016
    nMonComRet = IIf(Mid(MatNroCta(1), 9, 1) = 1, gcPEN_SIMBOLO & " ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(MatNroCta(1), 9, 1) = 1, " (" & UnNumero(nValor) & " " & StrConv(gcPEN_PLURAL, vbProperCase) & ")", " (" & UnNumero(nValor) & " Dolares)") 'marg ers044-2016
    '-----------------------------------------------------------

    '--------- Costo de Inactivas -------------------
    Set rs = New ADODB.Recordset
    nValor = loRs.GetParametro(2000, gMonDescInacME)
    nCosInac = Format$(nValor, "0.00")

    '------------------------------------------------

    '--Retiros/Depositos en Agencias de La Caja en Otras Plazas --
    Set rs = New ADODB.Recordset
    If Mid(MatNroCta(1), 9, 1) = 1 Then
       nParCod = 2046
    Else
       nParCod = 2047
    End If
    nValor = loRs.GetParametro(2000, nParCod)
    '''nMonOtraPlaza = IIf(Mid(MatNroCta(1), 9, 1) = 1, "S/. ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(MatNroCta(1), 9, 1) = 1, " (" & UnNumero(nValor) & " Nuevos Soles)", " (" & UnNumero(nValor) & " Dolares)") 'marg ers044-2016
    nMonOtraPlaza = IIf(Mid(MatNroCta(1), 9, 1) = 1, gcPEN_SIMBOLO & " ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(MatNroCta(1), 9, 1) = 1, " (" & UnNumero(nValor) & " " & StrConv(gcPEN_PLURAL, vbProperCase) & ")", " (" & UnNumero(nValor) & " Dolares)") 'marg ers044-2016

    '--------------------------------------------------------------

    '---------------------Monto Minimo de Retiro  -----------------
    Set rs = New ADODB.Recordset

    nParCod = 2027
    nValor = loRs.GetParametro(2000, nParCod)
    '''nMonMinRetMN = "S/. " & Format$(nValor, "0.00") & " (" & UnNumero(nValor) & " Nuevos Soles)" 'marg ers044-2016
    nMonMinRetMN = gcPEN_SIMBOLO & " " & Format$(nValor, "0.00") & " (" & UnNumero(nValor) & " " & StrConv(gcPEN_PLURAL, vbProperCase) & ")" 'marg ers044-2016

    nParCod = 2028
    nValor = loRs.GetParametro(2000, nParCod)
    nMonMinRetME = "$. " & Format$(nValor, "0.00") & " (" & UnNumero(nValor) & " Dolares)"

    'ALPA 20100108************************************************
    Dim objCaptac As COMNCaptaGenerales.NCOMCaptaGenerales
    Set objCaptac = New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim objCaptaGenerales As COMDCaptaGenerales.DCOMCaptaGenerales
    Set objCaptaGenerales = New COMDCaptaGenerales.DCOMCaptaGenerales
    Dim sListaCtas As String
    Dim lnMonto As Currency
    Dim lnMontoSald As Currency
    Dim lnTasaInteres As Currency
    Dim nI As Integer
    Dim rsCtas As ADODB.Recordset
    sListaCtas = ""
    sListaCtas = "'0"
    For nI = 1 To UBound(MatNroCta(), 1)
        sListaCtas = sListaCtas & "," & MatNroCta(nI)
    Next nI

        sListaCtas = sListaCtas & "'"
        Set rsCtas = objCaptaGenerales.ObtenerCtaPorPersonas(MatTitular(i, 9), sListaCtas)
        If Not (rsCtas.BOF And rsCtas.EOF) Then
            sListaCtas = rsCtas!cCtaCod
            lnMontoSald = rsCtas!nMonApert
            lnTasaInteres = rsCtas!nTasaInteres
        End If
'
    Dim pnPlazoT As Integer
    Dim nTREA As Currency
    nTREA = objCaptac.ObtenerTREA(Mid$(sListaCtas, 6, 3), IIf(lnMontoSald = 0, 10, lnMontoSald), lnTasaInteres, 0, 0)
    '*************************************************************


    '--------------------------------------------------------------
    'ALPA 20100108**************************************
    If nTipoAhorro = 0 Then
        lsModeloPlantilla = App.Path & "\FormatoCarta\CARTILLAAHORROCL.doc" 'JATO 20210407
    ElseIf nTipoAhorro = 5 Then
        lsModeloPlantilla = App.Path & "\FormatoCarta\CARTILLAAHORROCUENTASONADAL.doc" 'JATO 20210407
    ElseIf nTipoAhorro = 6 Then
        lsModeloPlantilla = App.Path & "\FormatoCarta\CARTILLACAJASUELDOL.doc" 'JATO 20210407
    '***Agregado por ELRO el 20130131, según TI-ERS020-2013
    ElseIf nTipoAhorro = 8 Then
        lsModeloPlantilla = App.Path & "\FormatoCarta\CARTILLAAHORROCONVENIOL.doc" 'JATO 20210407
    '***Fin Agregado por ELRO el 20130131******************
    End If
    '**************************************************

    'Crea una clase que de Word Object
    Dim wApp As Word.Application
    Dim wAppSource As Word.Application
    'Create a new instance of word
    Set wApp = New Word.Application
    Set wAppSource = New Word.Application

    Dim RangeSource As Word.Range
    'Abre Documento Plantilla
    wAppSource.Documents.Open FileName:=lsModeloPlantilla
    Set RangeSource = wAppSource.ActiveDocument.Content
    'Lo carga en Memoria
    wAppSource.ActiveDocument.Content.Copy

    'Crea Nuevo Documento
    wApp.Documents.Add

    With wApp.ActiveDocument.PageSetup
        .LeftMargin = CentimetersToPoints(1.5)
        .RightMargin = CentimetersToPoints(1)
        .TopMargin = CentimetersToPoints(1.5)
        .BottomMargin = CentimetersToPoints(1)
    End With

    With wApp.Selection.Font
        .Name = "Arial Narrow"
        .Size = 12
    End With

    For i = 1 To UBound(MatTitular)
        'ALPA 20100109**********************************
        'If nTipoAhorro = 6 Then

        Set rsCtas = objCaptaGenerales.ObtenerCtaPorPersonas(MatTitular(i, 9), sListaCtas)
        If Not (rsCtas.BOF And rsCtas.EOF) Then
            'sListaCtas = rsCtas!cCtaCod EAAS20180404
            lnMontoSald = MatTitular(i, 3)
            lnTasaInteres = MatTitular(i, 2)
        End If

'       Dim pnPlazoT As Integer
        'Dim nTREA As Currency

        'End If
        'APRI20190109 ERS077-2018
        Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
        Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
        nCostoMan = clsDef.GetCapCostoMantenimiento(232, nTipoAhorro, Mid(MatNroCta(1), 9, 1), lnMontoSald)
        nSaldoEquilibrio = clsMant.ObtenerSaldoEquilibrio(nCostoMan, lnTasaInteres)
        Set clsDef = Nothing
        Set clsMant = Nothing
        'END APRI

        'nTREA = objCaptac.ObtenerTREA(Mid$(sListaCtas, 6, 3), IIf(lnMontoSald = 0, 10, lnMontoSald), lnTasaInteres, 0, 0)
        nTREA = objCaptac.ObtenerTREA(Mid$(sListaCtas, 6, 3), IIf(lnMontoSald = 0, 10, lnMontoSald), lnTasaInteres, 0, nCostoMan) 'APRI20190109 ERS077-2018
        '***********************************************
        wApp.Application.Selection.TypeParagraph
        wApp.Application.Selection.PasteAndFormat (wdPasteDefault)
        wApp.Application.Selection.InsertBreak

        wApp.Selection.SetRange start:=wApp.Selection.start, End:=wApp.ActiveDocument.Content.End

        wApp.Selection.MoveEnd

        'ALPA 20100109**********************************
        '***Agregado por ELRO el 20130130, según TI-ERS020-2013
        'Institucion del Convenio
        If nTipoAhorro = 8 Then
            With wApp.Selection.Find
                .Text = "<<Nombre Empresa-Convenio>>"
                .Replacement.Text = pcInstConvDep
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
        End If
        '***Fin Agregado por ELRO el 20130130******************
        'INICIO EAAS20180403
        With wApp.Selection.Find
                .Text = "<<cCodCta>>"
                .Replacement.Text = rsCtas!cCtaCod
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
        End With

       'FIN EAAS20180403

        'If nTipoAhorro = 6 Then ' comentado por gitu 25-102012
        lsCad = ""
        lsCad = Format(nTREA, "0.00") & " % (Fija)"
        With wApp.Selection.Find
            .Text = "<<TasaTrea>>"
            .Replacement.Text = lsCad
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        '**********************************************
    'End If
    'APRI20190109 ERS077-2018
    lsCad = ""
    lsCad = Format(nSaldoEquilibrio, "0.00")
    With wApp.Selection.Find
        .Text = "<<SaldoEquilibrio>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END APRI
    'Monto
    lsCad = ""
    lsCad = IIf(Mid(MatNroCta(1), 9, 1) = 1, "MN ", "ME ") & Format(MatTitular(i, 3), "0.00")
    With wApp.Selection.Find
        .Text = "<<Monto>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Tasa
    lsCad = ""
    '''lsCad = Format(MatTitular(i, 2), "0.00") & " % " & IIf(Mid(MatNroCta(1), 9, 1) = 1, "Soles", "Dolares") 'marg ers044-2016
    lsCad = Format(MatTitular(i, 2), "0.00") & " % " & IIf(Mid(MatNroCta(1), 9, 1) = 1, StrConv(gcPEN_PLURAL, vbProperCase), "Dolares") 'marg ers044-2016
    With wApp.Selection.Find
        .Text = "<<Tasa>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Saldo de la cuenta por 3 meses
    With wApp.Selection.Find
        .Text = "<<MontoMin>>"
        .Replacement.Text = nMonMinCta
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Monto x Consulta
    With wApp.Selection.Find
        .Text = "<<MonConsul>>"
        .Replacement.Text = nMonxConsul
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Comision por Retiro
    With wApp.Selection.Find
        .Text = "<<MonComRet>>"
        .Replacement.Text = nMonComRet
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Costo de Inactivas
    With wApp.Selection.Find
        .Text = "<<CostoInac>>"
        .Replacement.Text = nCosInac
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Monto de Otra Plaza
    With wApp.Selection.Find
        .Text = "<<MonOtraPlaza>>"
        .Replacement.Text = nMonOtraPlaza
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Monto de Otra Plaza
    With wApp.Selection.Find
        .Text = "<<MonOtraPlaza>>"
        .Replacement.Text = nMonOtraPlaza
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Monto Minimo de Retiro en Soles
    With wApp.Selection.Find
        .Text = "<<MonMinRetMN>>"
        .Replacement.Text = nMonMinRetMN
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'ITF
    nitf = gnITFPorcent * 100
    With wApp.Selection.Find
        .Text = "<<TasaITF>>"
        '.Replacement.Text = Format$(nitf, "0.00")
        .Replacement.Text = Trim(CStr(nitf))
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Monto Minimo de Retiro en Dolares
    With wApp.Selection.Find
        .Text = "<<MonMinRetME>>"
        .Replacement.Text = nMonMinRetME
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Fecha
    lsFechas = Format(gdFecSis, "dd") & " de " & Format(gdFecSis, "mmmm") & " del " & Format(gdFecSis, "yyyy")
    With wApp.Selection.Find
        .Text = "<<FecActual>>"
        .Replacement.Text = lsFechas
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Lugar
    With wApp.Selection.Find
        .Text = "<<cLugar>>"
        .Replacement.Text = lsAgencia
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Direccion
    With wApp.Selection.Find
        .Text = "<<cDireccion>>"
        .Replacement.Text = lsAgeDir
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'APRI20180504 ERS036-2017
    'Tipo envio estado de cuenta
    With wApp.Selection.Find
        .Text = "<<TipoEnvio>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END APRI
    lsNom1 = "Nombre del Cliente: " & MatTitular(i, 1)
    lsDoc1 = "DNI/RUC: " & MatTitular(i, 4) & Space(60) & "Firma:______________________"
    lsDir1 = "Dirección: " & MatTitular(i, 5)

    With wApp.Selection.Find
        .Text = "<<NomTit1>>"
        .Replacement.Text = lsNom1
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DocTit1>>"
        .Replacement.Text = lsDoc1
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DirTit1>>"
        .Replacement.Text = lsDir1
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<NomTit2>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DocTit2>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DirTit2>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<NomTit3>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DocTit3>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DirTit3>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<NomTit4>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DocTit4>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DirTit4>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

  Next

wAppSource.ActiveDocument.Close
'***Agregado por ELRO el 20130131, según TI-ERS020-2013
wAppSource.Quit
wApp.ActiveDocument.SaveAs (App.Path & "\SPOOLER\AperturaLoteAhorro_" & psMovNro & ".doc")
'***Fin Agregado por ELRO el 20130131******************
wApp.Visible = True

End Function

'*** Funcion de Impresion de Cartillas segun Normas de SBS - AVMM - 10-08-2006 ***
Public Function ImprimeCartillaLote(pnPersNombre As String, ByVal pnTipo As Integer, ByVal psCtaCod As String, ByVal pnTasa As Double, _
                           ByVal pnMonto As Double, ByVal pdFechaA As Date, Optional pnPlazo As Integer) As String
' pnTipo=1--> Ahorro /pnTipo=2--> Plazo Fijo /pnTipo=3--> CTS
Dim lsCad As String
Dim nInteres As Double
Dim nMeses As Integer
Dim OiMO As Impresoras
Dim i As Byte
'nInteres = pnMonto * (pnTasa / 100)

lsCad = lsCad & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
lsCad = lsCad & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
lsCad = lsCad & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea

lsCad = lsCad & oImpresora.gPrnCondensadaON
If pnTipo = 1 Then
'IMPRESION CARTILLA DE AHORROS
    lsCad = lsCad & Space(10) & "1. Cuenta : " & Mid(psCtaCod, 1, 3) & "-" & Mid(psCtaCod, 4, 2) & "-" & Mid(psCtaCod, 6, 3) & "-" & Mid(psCtaCod, 9, 10) & Space(3) & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "2. Monto  : " & Format(pnMonto, "0.00") & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "3. T.E.A  : " & Format(pnTasa, "0.00") & "%" & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "4. Fecha de Corte para el Abonode de Intereses " & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(13) & "Capitalizacion : " & CDate(pdFechaA) + 30 & Space(2) & "Capitalizacion : " & CDate(pdFechaA) + 60 & Space(2) & "Capitalizacion : " & CDate(pdFechaA) + 90 & oImpresora.gPrnSaltoLinea
    '''lsCad = lsCad & Space(10) & "5. El saldo mínimo en cuenta de ahorros sera de S/. 5.00 Nuevos Soles o US$. 1.50 Dolares Americanos" & oImpresora.gPrnSaltoLinea 'marg ers044-2016
    lsCad = lsCad & Space(10) & "5. El saldo mínimo en cuenta de ahorros sera de " & gcPEN_SIMBOLO & " 5.00 " & StrConv(gcPEN_PLURAL, vbProperCase) & " o US$. 1.50 Dolares Americanos" & oImpresora.gPrnSaltoLinea 'marg ers044-2016
    lsCad = lsCad & Space(10) & "6. Comisiones y/o gastos aplicables: " & oImpresora.gPrnSaltoLinea
    '''lsCad = lsCad & Space(13) & "-Apartir del retrio Nro 31 se cargara a la cuenta un S/. 1.00 nuevos soles  o US$ 0.5 Dolares Americanos" & oImpresora.gPrnSaltoLinea 'marg ers044-2016
    lsCad = lsCad & Space(13) & "-Apartir del retrio Nro 31 se cargara a la cuenta un " & gcPEN_SIMBOLO & " 1.00 " & StrConv(gcPEN_PLURAL, vbLowerCase) & "  o US$ 0.5 Dolares Americanos" & oImpresora.gPrnSaltoLinea  'marg ers044-2016
    lsCad = lsCad & Space(14) & "durante el mes" & oImpresora.gPrnSaltoLinea
    '''lsCad = lsCad & Space(13) & "-Retiros y Depositos en otras plaza por montos mayores a S/. 5,000.00 Nuevos Soles o US$ 1,500.00 Dolares" & oImpresora.gPrnSaltoLinea 'marg ers044-2016
    lsCad = lsCad & Space(13) & "-Retiros y Depositos en otras plaza por montos mayores a " & gcPEN_SIMBOLO & " 5,000.00 " & StrConv(gcPEN_PLURAL, vbProperCase) & " o US$ 1,500.00 Dolares" & oImpresora.gPrnSaltoLinea 'marg ers044-2016
    lsCad = lsCad & Space(14) & "Americanos tendra una comision del 0.1 %" & oImpresora.gPrnSaltoLinea
    '''lsCad = lsCad & Space(13) & "-La cuentas inactivas genera una comisión semestral de S/. 5.00 Nuevos Soles o US$ 1.5 Dolares Americanos" & oImpresora.gPrnSaltoLinea 'marg ers044-2016
    lsCad = lsCad & Space(13) & "-La cuentas inactivas genera una comisión semestral de " & gcPEN_SIMBOLO & " 5.00 " & StrConv(gcPEN_PLURAL, vbProperCase) & " o US$ 1.5 Dolares Americanos" & oImpresora.gPrnSaltoLinea 'marg ers044-2016
    lsCad = lsCad & Space(10) & "7. Se aplicara la tasa de 0.08 % por Transaccion efectuada por concepto de Impuesto a las Transacciones Financieras (ITF)" & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(13) & "excepto  en el caso de cuentas exoneradas u operaciones inafectas de acuerdo a Ley" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    For i = 1 To 3
        lsCad = lsCad & oImpresora.gPrnSaltoLinea
    Next i

ElseIf pnTipo = 2 Then
'IMPRESION CARTILLA DE PLAZO FIJO
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
    Dim lnPlazo As Long
    Dim sTaInEf As String
    lnPlazo = CLng(pnPlazo)
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    sTaInEf = Format$(((((((pnTasa / 12) / 3000) + 1) ^ lnPlazo) - 1) * 100), "#0.00")
    nInteres = Format$(clsMant.GetInteresPF(pnTasa, pnMonto, lnPlazo), "#,##0.00")
    lsCad = lsCad & Space(10) & "1. Cuenta : " & Mid(psCtaCod, 1, 3) & "-" & Mid(psCtaCod, 4, 2) & "-" & Mid(psCtaCod, 6, 3) & "-" & Mid(psCtaCod, 9, 10) & Space(3) & "3. T.E.A : " & Format(sTaInEf, "0.00") & "%" & Space(2) & "Plazo : " & pnPlazo & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "2. Monto  : " & Format(pnMonto, "0.00") & Space(8) & "4. Monto Total : " & nInteres & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "5. Fecha de Corte para el Abono de Intereses: " & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(13) & "Apertura  : " & pdFechaA & oImpresora.gPrnSaltoLinea
    nMeses = pnPlazo / 30
    If nMeses <= 2 Then
        If nMeses = 1 Then
            lsCad = lsCad & Space(13) & "Provision : " & CDate(pdFechaA) + 30 & oImpresora.gPrnSaltoLinea
        Else
            lsCad = lsCad & Space(13) & "Provision : " & CDate(pdFechaA) + 30 & Space(2) & "Provision : " & CDate(pdFechaA) + pnPlazo & oImpresora.gPrnSaltoLinea
        End If
    Else
        If nMeses = 3 Then
            lsCad = lsCad & Space(13) & "Provision : " & CDate(pdFechaA) + 30 & Space(2) & "Provision : " & CDate(pdFechaA) + 60 & Space(2) & "Provision : " & CDate(pdFechaA) + pnPlazo & oImpresora.gPrnSaltoLinea
        Else
            lsCad = lsCad & Space(13) & "Provision : " & CDate(pdFechaA) + 30 & Space(2) & "Provision : " & CDate(pdFechaA) + 60 & Space(2) & "Provision : " & CDate(pdFechaA) + 90 & Space(2) & "...Provision : " & CDate(pdFechaA) + pnPlazo & oImpresora.gPrnSaltoLinea
        End If
    End If
    lsCad = lsCad & Space(10) & "6. Fecha de Vencimiento : " & CDate(pdFechaA) + pnPlazo & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "7. Los intereses podran ser retirados cada 30 días" & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "8. Se aplicara la tasa de 0.08 % por Transaccion efectuada por concepto de Impuesto a las Transacciones Financieras (ITF)" & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(13) & "excepto  en el caso de cuentas exoneradas u operaciones inafectas de acuerdo a Ley" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    For i = 1 To 3
        lsCad = lsCad & oImpresora.gPrnSaltoLinea
    Next i

Else
'IMPRESION CARTILLA CTS
    lsCad = lsCad & Space(10) & "1. Cuenta : " & Mid(psCtaCod, 1, 3) & "-" & Mid(psCtaCod, 4, 2) & "-" & Mid(psCtaCod, 6, 3) & "-" & Mid(psCtaCod, 9, 10) & Space(3) & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "2. Monto  : " & Format(pnMonto, "0.00") & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "3. T.E.A : " & Format(pnTasa, "0.00") & "%" & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "4. Fecha de Corte para el Abonode de Intereses " & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(13) & "Capitalizacion : " & CDate(pdFechaA) + 30 & Space(2) & "Capitalizacion : " & CDate(pdFechaA) + 60 & Space(2) & "Capitalizacion : " & CDate(pdFechaA) + 90 & oImpresora.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "5. Los Retiros se efectuaran conforme a los porcentajes de Disponibilidad establecida por ley." & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    For i = 1 To 13
        lsCad = lsCad & oImpresora.gPrnSaltoLinea
    Next i

End If

'lsCad = lsCad & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'lsCad = lsCad & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'lsCad = lsCad & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea

For i = 1 To 21
        lsCad = lsCad & oImpresora.gPrnSaltoLinea
Next i

'IMPRESION DE FIRMAS
lsCad = lsCad & Space(50) & "---------------------------------------" & oImpresora.gPrnSaltoLinea
lsCad = lsCad & Space(52) & pnPersNombre & oImpresora.gPrnSaltoLinea
lsCad = lsCad & oImpresora.gPrnCondensadaOFF
ImprimeCartillaLote = lsCad & oImpresora.gPrnCondensadaOFF
End Function


'Solo se necesita cuando utilizamod el Proyecto VBP
Public Sub MuestraFirma(psPersCod As String, Optional ByVal psCodAge As String = "", Optional ByVal pbComprobacionSinForm As Boolean = False, Optional ByRef pbTieneFirma = False) ' ande 20170914 se agregó los parametros pbComprobacionSinForm y pbTieneFirma
    Dim R As ADODB.Recordset
    Set R = New ADODB.Recordset
    Dim clsCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set clsCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
    'ande 20170914 agregado de la ver el resultado sin cargar el formulario de firma por medio del parametro pbTieneFirma
    Set R = clsCapMov.GetFirma(psPersCod, psCodAge)
    'frmMuestraFirma.psCodCli = psPersCod
    If R.State = adStateClosed Then
        If pbComprobacionSinForm Then
            pbTieneFirma = False
        Else
            MsgBox "Cliente SIN Firma", vbInformation, "Aviso"
        End If

    ElseIf Not (R.EOF And R.BOF) Then
        If R.RecordCount > 0 Then
            If pbComprobacionSinForm Then
                pbTieneFirma = True
            Else
             Call frmPersonaFirma.inicio(psPersCod, psCodAge, False)
             'Call frmMuestraFirma.IDBFirma.CargarFirma(R)
            End If
        End If
        R.Close
        Set clsCapMov = Nothing

        If pbComprobacionSinForm = False Then
            frmPersonaFirma.Show 1
        End If
    Else
        pbTieneFirma = False
    'end ande
    End If
End Sub


'************ IMPRESION DE CARTILLAS EN WORD 06-02-2007 ************************
'By Capi 03112008 se agrego parametros para panderito
'Public Sub ImpreCartillaAhoCorriente(MatTitular As Variant, ByVal psCtaCod As String, ByVal pnTasa As Double, ByVal pnMonto As Double, Optional ByVal pnTpoProgram As Integer = 0)
Public Sub ImpreCartillaAhoCorriente(MatTitular As Variant, ByVal psCtaCod As String, ByVal pnTasa As Double, ByVal pnMonto As Double, Optional ByVal pnTpoProgram As Integer = 0, Optional ByVal pnPlazo = 0, Optional ByVal pdFecSis, Optional nCostoMan As Currency = 0, Optional pcInstConvDep As String = "")
    Dim rs As ADODB.Recordset
    Dim nValor As Double
    Dim nMonMinCta As String
    Dim nMonxConsul As String
    Dim nMonComRet As String
    Dim nCosInac As Double
    Dim nMonOtraPlaza As String
    Dim nTasaITF As Double
    Dim nMonMinRetMN As String
    Dim nMonMinRetME As String
    Dim nMonMinRetEu As String '*** PEAC 20090727
    Dim nParCod  As Long
    Dim lsFechas As String
    Dim lsFechaVen As String
    Dim lsNom1 As String, lsNom2 As String, lsNom3 As String, lsNom4 As String
    Dim lsDoc1 As String, lsDoc2 As String, lsDoc3 As String, lsDoc4 As String
    Dim lsDir1 As String, lsDir2 As String, lsDir3 As String, lsDir4 As String
    Dim lsCad As String
    Dim lsCad2 As String
    Dim nitf As Double
    Dim nCosEnvioFisico As String 'JUEZ 20130520
    Dim sTipoEnvio As String 'JUEZ 20130520
    'INICIO EAAS20180530 MODIFICACION DE CARTILLAS
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim lnPlazo As Long
    lnPlazo = CLng(pnPlazo)
    Dim lnTasa As Double
    lnTasa = clsMant.GetTasaNominal(pnTasa, 360)
    Dim nInteres As Double
    nInteres = Format$(clsMant.GetInteresPF(lnTasa, pnMonto, lnPlazo), "#,##0.00")
    'FIN EAAS20180530 MODIFICACION DE CARTILLAS
    'APRI20190109 ERS077-2018
    Dim nSaldoEquilibrio As Currency
    nSaldoEquilibrio = clsMant.ObtenerSaldoEquilibrio(nCostoMan, pnTasa)
    'END APRI
    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range
    
    'CTI2 ferimoro: ERS030-2019 // JATO 20210407 ************************
    Dim nCCI As String
    nCCI = clsMant.recuperaCCI(psCtaCod)
    Dim nSubProducto As String
    nSubProducto = clsMant.recuperaSubProducto(gCapAhorros, pnTpoProgram)
    '**************************************************

    Dim loRs As COMDConstSistema.DCOMGeneral
    Set loRs = New COMDConstSistema.DCOMGeneral

    Dim loAge As COMDConstantes.DCOMAgencias
    Dim rs1 As ADODB.Recordset
    Dim lsAgencia As String
    Dim lsAgeDir As String
    Dim lsDesAgencia As String 'CTI2 FERIMORO : ERS030-2019 // JATO 20210407
    Set loAge = New COMDConstantes.DCOMAgencias
    Set rs1 = New ADODB.Recordset
        Set rs1 = loAge.RecuperaAgencias(gsCodAge)
        If Not (rs1.EOF And rs1.BOF) Then
            lsAgencia = Trim(rs1("cUbiGeoDescripcion"))
            lsDesAgencia = Trim(rs1("cAgeDescripcion")) 'CTI2 FERIMORO : ERS030-2019
            lsAgeDir = rs1("cAgeDireccion")
        End If
    Set loAge = Nothing

    'JUEZ 20141201 Nuevos Parámetros *******************
    Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
    Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
    Dim rsPar As ADODB.Recordset
    Set rsPar = clsDef.GetCapParametroNew(Mid(psCtaCod, 6, 3), pnTpoProgram)
    'END JUEZ ******************************************

    '--------Saldo x Cuenta a Mantener por 3 Meses --

    If Mid(psCtaCod, 9, 1) = 1 Then
       nParCod = 2091
    ElseIf Mid(psCtaCod, 9, 1) = 2 Then
       nParCod = 2092
    ElseIf Mid(psCtaCod, 9, 1) = 3 Then '*** PEAC 20090727
       nParCod = 2103
    End If
    'nValor = loRs.GetParametro(2000, nParCod)
    'JUEZ 20150121 **************************************
    If Mid(psCtaCod, 6, 3) = gCapAhorros Then
        nValor = 0 'IIf(Mid(psCtaCod, 9, 1) = 1, rsPar!nSaldoMinCtaSol, rsPar!nSaldoMinCtaDol) 'COMENTADO POR APRI20190109 ERS077-2018
    Else
        nValor = IIf(Mid(psCtaCod, 9, 1) = 1, rsPar!nAumCapMinSol, rsPar!nAumCapMinDol)
    End If
    If pnTpoProgram = 2 Then
        Dim clsCapGen As New COMNCaptaGenerales.NCOMCaptaGenerales
        nValor = clsCapGen.GetDatosCuenta(psCtaCod)!nMontoAbono
        Set clsCapGen = Nothing
    End If
    'END JUEZ *******************************************
    '*** PEAC 20090727 - SE AGREGO EUROS
    'nMonMinCta = IIf(Mid(psCtaCod, 9, 1) = 1, "S/. ", IIf(Mid(psCtaCod, 9, 1) = 2, "$. ", "Eu.")) & Format$(nValor, "0.00") & IIf(Mid(psCtaCod, 9, 1) = 1, " (" & UnNumero(nValor) & " Nuevos Soles)", IIf(Mid(psCtaCod, 9, 1) = 2, " (" & UnNumero(nValor) & " DOLARES)", " (" & UnNumero(nValor) & " EUROS)"))
    '''nMonMinCta = IIf(Mid(psCtaCod, 9, 1) = 1, "S/. ", IIf(Mid(psCtaCod, 9, 1) = 2, "$. ", "Eu.")) & Format$(nValor, "0.00") & IIf(Mid(psCtaCod, 9, 1) = 1, " (Nuevos Soles)", IIf(Mid(psCtaCod, 9, 1) = 2, " (DOLARES)", " (EUROS)")) 'JUEZ 20150121 'marg ers044-2016
    nMonMinCta = IIf(Mid(psCtaCod, 9, 1) = 1, gcPEN_SIMBOLO & " ", IIf(Mid(psCtaCod, 9, 1) = 2, "$. ", "Eu.")) & Format$(nValor, "0.00") & IIf(Mid(psCtaCod, 9, 1) = 1, " (" & StrConv(gcPEN_PLURAL, vbProperCase) & ")", IIf(Mid(psCtaCod, 9, 1) = 2, " (DOLARES)", " (EUROS)")) 'JUEZ 20150121 'marg ers044-2016

    '------------------------------------------------


    '-----Monto de Consulta por extracto de Cuenta --
    'By Capi 05032008

'    Set rs = New ADODB.Recordset
'    If Mid(psCtaCod, 9, 1) = 1 Then
'       nParCod = gDctoExtMNxPag
'    Else
'       nParCod = gDctoExtMExPag
'    End If
'    nValor = loRs.GetParametro(2000, nParCod)
'    nMonxConsul = IIf(Mid(psCtaCod, 9, 1) = 1, "S/. ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(psCtaCod, 9, 1) = 1, " (" & UnNumero(nValor) & " Nuevos Soles)", " (" & UnNumero(nValor) & " Dolares)")
'

     '-----Comision Consulta Saldo Ventanilla --
    Set rs = New ADODB.Recordset
    nParCod = 2106
    nValor = loRs.GetParametro(2000, nParCod)
    '''nMonxConsul = "S/." & Format$(nValor, "0.00") & " (" & UnNumero(nValor) & " Nuevos Soles)" 'marg ers044-2016
    nMonxConsul = gcPEN_SIMBOLO & Format$(nValor, "0.00") & " (" & UnNumero(nValor) & " " & StrConv(gcPEN_PLURAL, vbProperCase) & ")" 'marg ers044-2016


    '------------------------------------------------

    '-----Comision de retiro cuando el monto es menor a 1000 --
    Set rs = New ADODB.Recordset
    If Mid(psCtaCod, 9, 1) = 1 Then
       nParCod = gMontoMNx31Ope
    ElseIf Mid(psCtaCod, 9, 1) = 2 Then
       nParCod = gMontoMEx31Ope
    ElseIf Mid(psCtaCod, 9, 1) = 3 Then '*** PEAC 20090727
       nParCod = gMontoEurox31Ope
    End If
    nValor = loRs.GetParametro(2000, nParCod)
    '''nMonComRet = IIf(Mid(psCtaCod, 9, 1) = 1, "S/. ", IIf(Mid(psCtaCod, 9, 1) = 2, "$. ", "Eu.")) & Format$(nValor, "0.00") & IIf(Mid(psCtaCod, 9, 1) = 1, " (" & UnNumero(nValor) & " Nuevos Soles)", IIf(Mid(psCtaCod, 9, 1) = 2, " (" & UnNumero(nValor) & " DOLARES)", " (" & UnNumero(nValor) & " EUROS)")) 'marg ers044-2016
    nMonComRet = IIf(Mid(psCtaCod, 9, 1) = 1, gcPEN_SIMBOLO & " ", IIf(Mid(psCtaCod, 9, 1) = 2, "$. ", "Eu.")) & Format$(nValor, "0.00") & IIf(Mid(psCtaCod, 9, 1) = 1, " (" & UnNumero(nValor) & " " & StrConv(gcPEN_PLURAL, vbProperCase) & ")", IIf(Mid(psCtaCod, 9, 1) = 2, " (" & UnNumero(nValor) & " DOLARES)", " (" & UnNumero(nValor) & " EUROS)")) 'marg ers044-2016

    '-----------------------------------------------------------

    '--------- Costo de Inactivas -------------------
    Set rs = New ADODB.Recordset
    nValor = loRs.GetParametro(2000, gMonDescInacME)
    nCosInac = Format$(nValor, "0.00")

    '------------------------------------------------

    '--Retiros/Depositos en Agencias de La Caja en Otras Plazas --
    Set rs = New ADODB.Recordset
    If Mid(psCtaCod, 9, 1) = 1 Then
       nParCod = 2046
    ElseIf Mid(psCtaCod, 9, 1) = 2 Then
       nParCod = 2047
    ElseIf Mid(psCtaCod, 9, 1) = 3 Then '*** PEAC 20090727
       nParCod = 2109
    End If
    nValor = loRs.GetParametro(2000, nParCod)

    '''nMonOtraPlaza = IIf(Mid(psCtaCod, 9, 1) = 1, "S/. ", IIf(Mid(psCtaCod, 9, 1) = 2, "$. ", "Eu.")) & Format$(nValor, "0.00") & IIf(Mid(psCtaCod, 9, 1) = 1, " (" & UnNumero(nValor) & " Nuevos Soles)", IIf(Mid(psCtaCod, 9, 1) = 2, " (" & UnNumero(nValor) & " DOLARES)", " (" & UnNumero(nValor) & " EUROS)")) 'marg ers044-2016
    nMonOtraPlaza = IIf(Mid(psCtaCod, 9, 1) = 1, gcPEN_SIMBOLO & " ", IIf(Mid(psCtaCod, 9, 1) = 2, "$. ", "Eu.")) & Format$(nValor, "0.00") & IIf(Mid(psCtaCod, 9, 1) = 1, " (" & UnNumero(nValor) & " " & StrConv(gcPEN_PLURAL, vbProperCase) & ")", IIf(Mid(psCtaCod, 9, 1) = 2, " (" & UnNumero(nValor) & " DOLARES)", " (" & UnNumero(nValor) & " EUROS)")) 'marg ers044-2016

    '---------------------Monto Minimo de Retiro  -----------------
    Set rs = New ADODB.Recordset

       nParCod = 2027
       nValor = loRs.GetParametro(2000, nParCod)
       '''nMonMinRetMN = "S/. " & Format$(nValor, "0.00") & " (" & UnNumero(nValor) & " Nuevos Soles)" 'marg ers044-2016
       nMonMinRetMN = gcPEN_SIMBOLO & " " & Format$(nValor, "0.00") & " (" & UnNumero(nValor) & " " & StrConv(gcPEN_PLURAL, vbProperCase) & ")" 'marg ers044-2016

       nParCod = 2028
       nValor = loRs.GetParametro(2000, nParCod)
       nMonMinRetME = "$. " & Format$(nValor, "0.00") & " (" & UnNumero(nValor) & " Dolares)"

       nParCod = 2110
       nValor = loRs.GetParametro(2000, nParCod)
       nMonMinRetEu = "$. " & Format$(nValor, "0.00") & " (" & UnNumero(nValor) & " Dolares)"

    '--------------------------------------------------------------
       'JUEZ 20130520 ***************************************
       Set rs = New ADODB.Recordset
       nValor = loRs.GetParametro(2000, 1005)
       '''nCosEnvioFisico = "S/. " & Format$(nValor, "0.00") 'marg ers044-2016
       nCosEnvioFisico = gcPEN_SIMBOLO & " " & Format$(nValor, "0.00") 'marg ers044-2016

       Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
       Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
       Set rs = oCap.RecuperaDatosEnvioEstadoCta(psCtaCod)
       'CTI2 ERS030-2019 // JATO 20210407************************
       Dim nFormaEnvio As String
       Dim nEnvioSiNo As String
       '****************************************
       'sTipoEnvio = rs!cTipoEnvio
       'APRI2018 ERS036-2017
        If Not (rs.BOF And rs.EOF) Then
          If rs.RecordCount > 0 Then
            sTipoEnvio = oCap.MostarTextoComisionEnvioEstAhorros(rs!nModoEnvio, pnTpoProgram)
            nFormaEnvio = rs!cTipoEnvio 'CTI2 ERS030-2019 // JATO 20210407************************
            nEnvioSiNo = "Si" 'CTI2 ERS030-2019 // JATO 20210407************************
          'End If
          Else 'CTI2 ERS030-2019************************
            nFormaEnvio = "" 'CTI2 ERS030-2019 // JATO 20210407************************
            nEnvioSiNo = "No" 'CTI2 ERS030-2019 // JATO 20210407************************
          End If
       Else
            sTipoEnvio = ""
            nFormaEnvio = "" 'CTI2 ERS030-2019 // JATO 20210407************************
            nEnvioSiNo = "No" 'CTI2 ERS030-2019 // JATO 20210407************************
       End If
       'END APRI
       Set rs = Nothing
       'END JUEZ ********************************************
    'ALPA 20091118************************************************
    Dim objCaptac As COMNCaptaGenerales.NCOMCaptaGenerales
    Set objCaptac = New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim pnPlazoT As Integer
    Dim nTREA As Currency
    pnPlazoT = CInt(pnPlazo)
    nTREA = objCaptac.ObtenerTREA(Mid$(psCtaCod, 6, 3), IIf(pnMonto = 0, 10, pnMonto), pnTasa, pnPlazoT, nCostoMan)
    '*************************************************************
    Set oWord = CreateObject("Word.Application")
    oWord.Visible = True

    If pnTpoProgram = 0 Then
        Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\CARTILLAAHORROC.doc")
    'By Capi 20082008 para cuenta soñada
    ElseIf pnTpoProgram = 5 Then
        Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\CARTILLAAHORROCUENTASONADA.doc")
    'By capi 03112008 para panderito
    ElseIf pnTpoProgram = 2 Then
        'Set oDoc = oWord.Documents.Open(App.path & "\FormatoCarta\CARTILLAAHORROPanderito.doc")
        Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\CARTILLAAHORRODIARIO.doc") 'JUEZ 20150121
    'ALPA 20100106*********************
    ElseIf pnTpoProgram = 6 Then
        Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\CARTILLACAJASUELDO.doc")
    '**********************************
    'BRGO 20111230*********************
    ElseIf pnTpoProgram = 7 Then
        Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\CARTILLAAHORROECOTAXI.doc")
    '**********************************
    '***Agregado por ELRO el 20130130, según TI-ERS020-2013
    ElseIf pnTpoProgram = 8 Then
        Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\CARTILLAAHORROCONVENIO.doc")
    '***Fin Agregado por ELRO el 20130130******************
    Else
        'Set oDoc = oWord.Documents.Open(App.path & "\FormatoCarta\CARTILLAAHORROPanderito.doc")
        Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\CARTILLAAHORRODIARIO.doc") 'JUEZ 20150121
    End If
    '**********************JHCU-ACTA N° 050-2019**************************'
    oDoc.SaveAs App.Path & "\SPOOLER\" & psCtaCod & ".doc"
    '**********************JHCU**************************'

    '***Agregado por ELRO el 20130130, según TI-ERS020-2013
    'Institucion del Convenio
    If pnTpoProgram = 8 Then
        With oWord.Selection.Find
            .Text = "<<Nombre Empresa-Convenio>>"
            .Replacement.Text = pcInstConvDep
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
    End If
    '***Fin Agregado por ELRO el 20130130******************
    'CTI2 FERIMORO: ERS030-2019 ************************************************
    lsCad = lsDesAgencia
    With oWord.Selection.Find
        .Text = "<<Oficina>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    lsCad = Format(Time, "hh:mm:ss")
    With oWord.Selection.Find
        .Text = "<<hora>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    lsCad = gsCodUser
    With oWord.Selection.Find
        .Text = "<<user>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    lsCad = nEnvioSiNo
    With oWord.Selection.Find
        .Text = "<<Elección SI/NO>>"
        .Replacement.Text = lsCad & Space(10)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    lsCad = nFormaEnvio
    With oWord.Selection.Find
        .Text = "<<TipoEnvio>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    '***********************************************************************

    'JUEZ 20150121 ************************************************
    'Cuenta
    lsCad = psCtaCod
    With oWord.Selection.Find
        .Text = "<<cCodCta>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END JUEZ *****************************************************
        'CTI2 FERIMORO: ERS030-2019 ************************************************
    'Cuenta
    lsCad = nCCI
    With oWord.Selection.Find
        .Text = "<<CodCCI>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    lsCad = nSubProducto
    With oWord.Selection.Find
        .Text = "<<SubProducto>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END CTI2 FERIMORO *****************************************************

    'Monto
    lsCad = IIf(Mid(psCtaCod, 9, 1) = 1, "MN ", "ME ") & Format(pnMonto, "#,##0.00")
    With oWord.Selection.Find
        .Text = "<<Monto>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Tasa
    lsCad = ""
    lsCad = Format(pnTasa, "0.00") & " % (Fija)"
    With oWord.Selection.Find
        .Text = "<<Tasa>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    lsCad = ""
    lsCad = Format(nTREA, "0.00") & " % (Fija)"
    With oWord.Selection.Find
        .Text = "<<TasaTrea>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'APRI20190109 ERS077-2018
    lsCad = ""
    lsCad = Format(nSaldoEquilibrio, "#,##0.00")
    With oWord.Selection.Find
        .Text = "<<SaldoEquilibrio>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END APRI
    'JUEZ 20150121 ****************************
    If pnTpoProgram = 2 Then
        With oWord.Selection.Find
            .Text = "<<FormaRetiro>>"
            .Replacement.Text = "FINAL DEL PLAZO"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
    End If
    'END JUEZ *********************************
    'Saldo de la cuenta por 3 meses
    With oWord.Selection.Find
        .Text = "<<MontoMin>>"
        .Replacement.Text = nMonMinCta
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Monto x Consulta
    With oWord.Selection.Find
        .Text = "<<MonConsul>>"
        .Replacement.Text = nMonxConsul
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Comision por Retiro
    With oWord.Selection.Find
        .Text = "<<MonComRet>>"
        .Replacement.Text = nMonComRet
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Costo de Inactivas
    With oWord.Selection.Find
        .Text = "<<CostoInac>>"
        .Replacement.Text = nCosInac
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Monto de Otra Plaza
    With oWord.Selection.Find
        .Text = "<<MonOtraPlaza>>"
        .Replacement.Text = nMonOtraPlaza
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Monto de Otra Plaza
    With oWord.Selection.Find
        .Text = "<<MonOtraPlaza>>"
        .Replacement.Text = nMonOtraPlaza
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Monto Minimo de Retiro en Soles
    With oWord.Selection.Find
        .Text = "<<MonMinRetMN>>"
        .Replacement.Text = nMonMinRetMN
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'ITF
    If pnTpoProgram <> 6 Then 'MAVM 20110406 ***
        nitf = gnITFPorcent * 100
        With oWord.Selection.Find
            .Text = "<<TasaITF>>"
            '.Replacement.Text = Format$(nitf, "0.00")
            .Replacement.Text = Trim(CStr(nitf))
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
    End If 'MAVM 20110406 ***

    'Monto Minimo de Retiro en Dolares
    With oWord.Selection.Find
        .Text = "<<MonMinRetME>>"
        .Replacement.Text = nMonMinRetME
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Fecha
    lsFechas = Format(gdFecSis, "dd") & " de " & Format(gdFecSis, "mmmm") & " del " & Format(gdFecSis, "yyyy")
    With oWord.Selection.Find
        .Text = "<<FecActual>>"
        .Replacement.Text = lsFechas
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'By Capi 03112008 para panderito 08012009 ahora se llama diario
    lsFechaVen = Format(DateAdd("d", gdFecSis, pnPlazo), "dd") & " de " & Format(DateAdd("d", gdFecSis, pnPlazo), "mmmm") & " del " & Format(DateAdd("d", gdFecSis, pnPlazo), "yyyy")
    With oWord.Selection.Find
        .Text = "<<Vencimiento>>"
        .Replacement.Text = lsFechaVen
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<Plazo>>"
        .Replacement.Text = pnPlazo
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'End by

    'Lugar
    With oWord.Selection.Find
        .Text = "<<cLugar>>"
        .Replacement.Text = lsAgencia
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Direccion
    With oWord.Selection.Find
        .Text = "<<cDireccion>>"
        .Replacement.Text = lsAgeDir
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'JUEZ 20130520 ****************************
    'Tipo envio estado de cuenta
    With oWord.Selection.Find
        .Text = "<<TipoEnvio>>"
        .Replacement.Text = sTipoEnvio
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'Costo envio fisico
    With oWord.Selection.Find
        .Text = "<<CostoEnvioFisico>>"
        .Replacement.Text = nCosEnvioFisico
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END JUEZ *********************************

    'INICIO EAAS20180530 MODIFICACION DE CARTILLAS
    lsCad = ""
    lsCad = IIf(Mid(psCtaCod, 9, 1) = 1, gcPEN_SIMBOLO & " ", "$. ") & Format(nInteres, "#,##0.00")
    With oWord.Selection.Find
        .Text = "<<MonInteres>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'FIN EAAS20180530 MODIFICACION DE CARTILLAS

    Dim i As Integer
    Dim rsDir As ADODB.Recordset 'JUEZ 20130520
    Dim oPers As COMDpersona.DCOMPersona 'JUEZ 20130520
    
    'CTI2 ferimoro : ERS030-2019 //JATO 20210407
    Call compForCartillaP(MatTitular, oWord)

    'CTI2 // JATO 20210407
'    If UBound(MatTitular) = 2 Then
'
'        lsNom1 = "Nombre del Cliente: " & MatTitular(1, 1)
'        lsDoc1 = "DNI/RUC: " & MatTitular(1, 2) & space(60) & "Firma:______________________"
'        lsDir1 = "Dirección: " & MatTitular(1, 3)
'        'JUEZ 20130520 ****************************
'        Set oPers = New COMDPersona.DCOMPersona
'        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(1, 2))
'        lsDir1 = "Dirección: " & rsDir!cPersDireccDomicilio
'        'END JUEZ *********************************
'        With oWord.Selection.Find
'            .Text = "<<NomTit1>>"
'            .Replacement.Text = lsNom1
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DocTit1>>"
'            .Replacement.Text = lsDoc1
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DirTit1>>"
'            .Replacement.Text = lsDir1
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<NomTit2>>"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DocTit2>>"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DirTit2>>"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<NomTit3>>"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DocTit3>>"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DirTit3>>"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<NomTit4>>"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DocTit4>>"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DirTit4>>"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'
'    ElseIf UBound(MatTitular) = 3 Then
'        lsNom1 = "Nombre del Cliente: " & MatTitular(1, 1)
'        lsDoc1 = "DNI/RUC: " & MatTitular(1, 2) & space(60) & "Firma:______________________"
'        lsDir1 = "Dirección: " & MatTitular(1, 3)
'        'JUEZ 20130520 ****************************
'        Set oPers = New COMDPersona.DCOMPersona
'        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(1, 2))
'        lsDir1 = "Dirección: " & rsDir!cPersDireccDomicilio
'        'END JUEZ *********************************
'        lsNom2 = "Nombre del Cliente: " & MatTitular(2, 1)
'        lsDoc2 = "DNI/RUC: " & MatTitular(2, 2) & space(60) & "Firma:______________________"
'        lsDir2 = "Dirección: " & MatTitular(2, 3)
'        'JUEZ 20130520 ****************************
'        Set oPers = New COMDPersona.DCOMPersona
'        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(2, 2))
'        lsDir2 = "Dirección: " & rsDir!cPersDireccDomicilio
'        'END JUEZ *********************************
'
'        With oWord.Selection.Find
'            .Text = "<<NomTit1>>"
'            .Replacement.Text = lsNom1
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DocTit1>>"
'            .Replacement.Text = lsDoc1
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DirTit1>>"
'            .Replacement.Text = lsDir1
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<NomTit2>>"
'            .Replacement.Text = lsNom2
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DocTit2>>"
'            .Replacement.Text = lsDoc2
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DirTit2>>"
'            .Replacement.Text = lsDir2
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<NomTit3>>"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DocTit3>>"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DirTit3>>"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<NomTit4>>"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DocTit4>>"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DirTit4>>"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'    ElseIf UBound(MatTitular) = 4 Then
'        lsNom1 = "Nombre del Cliente: " & MatTitular(1, 1)
'        lsDoc1 = "DNI/RUC: " & MatTitular(1, 2) & space(60) & "Firma:______________________"
'        lsDir1 = "Dirección: " & MatTitular(1, 3)
'        'JUEZ 20130520 ****************************
'        Set oPers = New COMDPersona.DCOMPersona
'        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(1, 2))
'        lsDir1 = "Dirección: " & rsDir!cPersDireccDomicilio
'        'END JUEZ *********************************
'        lsNom2 = "Nombre del Cliente: " & MatTitular(2, 1)
'        lsDoc2 = "DNI/RUC: " & MatTitular(2, 2) & space(60) & "Firma:______________________"
'        lsDir2 = "Dirección: " & MatTitular(2, 3)
'        'JUEZ 20130520 ****************************
'        Set oPers = New COMDPersona.DCOMPersona
'        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(2, 2))
'        lsDir2 = "Dirección: " & rsDir!cPersDireccDomicilio
'        'END JUEZ *********************************
'        lsNom3 = "Nombre del Cliente: " & MatTitular(3, 1)
'        lsDoc3 = "DNI/RUC: " & MatTitular(3, 2) & space(60) & "Firma:______________________"
'        lsDir3 = "Dirección: " & MatTitular(3, 3)
'        'JUEZ 20130520 ****************************
'        Set oPers = New COMDPersona.DCOMPersona
'        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(3, 2))
'        lsDir3 = "Dirección: " & rsDir!cPersDireccDomicilio
'        'END JUEZ *********************************
'
'        With oWord.Selection.Find
'            .Text = "<<NomTit1>>"
'            .Replacement.Text = lsNom1
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DocTit1>>"
'            .Replacement.Text = lsDoc1
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DirTit1>>"
'            .Replacement.Text = lsDir1
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<NomTit2>>"
'            .Replacement.Text = lsNom2
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DocTit2>>"
'            .Replacement.Text = lsDoc2
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DirTit2>>"
'            .Replacement.Text = lsDir2
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<NomTit3>>"
'            .Replacement.Text = lsNom3
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DocTit3>>"
'            .Replacement.Text = lsDoc3
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DirTit3>>"
'            .Replacement.Text = lsDir3
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<NomTit4>>"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DocTit4>>"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DirTit4>>"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'
'    ElseIf UBound(MatTitular) = 5 Then
'        lsNom1 = "Nombre del Cliente: " & MatTitular(1, 1)
'        lsDoc1 = "DNI/RUC: " & MatTitular(1, 2) & space(60) & "Firma:______________________"
'        lsDir1 = "Dirección: " & MatTitular(1, 3)
'        'JUEZ 20130520 ****************************
'        Set oPers = New COMDPersona.DCOMPersona
'        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(1, 2))
'        lsDir1 = "Dirección: " & rsDir!cPersDireccDomicilio
'        'END JUEZ *********************************
'        lsNom2 = "Nombre del Cliente: " & MatTitular(2, 1)
'        lsDoc2 = "DNI/RUC: " & MatTitular(2, 2) & space(60) & "Firma:______________________"
'        lsDir2 = "Dirección: " & MatTitular(2, 3)
'        'JUEZ 20130520 ****************************
'        Set oPers = New COMDPersona.DCOMPersona
'        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(2, 2))
'        lsDir2 = "Dirección: " & rsDir!cPersDireccDomicilio
'        'END JUEZ *********************************
'        lsNom3 = "Nombre del Cliente: " & MatTitular(3, 1)
'        lsDoc3 = "DNI/RUC: " & MatTitular(3, 2) & space(60) & "Firma:______________________"
'        lsDir3 = "Dirección: " & MatTitular(3, 3)
'        'JUEZ 20130520 ****************************
'        Set oPers = New COMDPersona.DCOMPersona
'        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(3, 2))
'        lsDir3 = "Dirección: " & rsDir!cPersDireccDomicilio
'        'END JUEZ *********************************
'        lsNom4 = "Nombre del Cliente: " & MatTitular(4, 1)
'        lsDoc4 = "DNI/RUC: " & MatTitular(4, 2) & space(60) & "Firma:______________________"
'        lsDir4 = "Dirección: " & MatTitular(4, 3)
'        'JUEZ 20130520 ****************************
'        Set oPers = New COMDPersona.DCOMPersona
'        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(4, 2))
'        lsDir4 = "Dirección: " & rsDir!cPersDireccDomicilio
'        'END JUEZ *********************************
'
'        With oWord.Selection.Find
'            .Text = "<<NomTit1>>"
'            .Replacement.Text = lsNom1
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DocTit1>>"
'            .Replacement.Text = lsDoc1
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DirTit1>>"
'            .Replacement.Text = lsDir1
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<NomTit2>>"
'            .Replacement.Text = lsNom2
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DocTit2>>"
'            .Replacement.Text = lsDoc2
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DirTit2>>"
'            .Replacement.Text = lsDir2
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<NomTit3>>"
'            .Replacement.Text = lsNom3
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DocTit3>>"
'            .Replacement.Text = lsDoc3
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DirTit3>>"
'            .Replacement.Text = lsDir3
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<NomTit4>>"
'            .Replacement.Text = lsNom4
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DocTit4>>"
'            .Replacement.Text = lsDoc4
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<DirTit4>>"
'            .Replacement.Text = lsDir4
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'    End If
    'MEJORAS JHCU
    'oWord.Visible = True 'ADD JHCU 26-04-2019 ACTA 050-2019
    oDoc.Save
    oDoc.Close ' CTI2 // JATO 20210407
    'oDoc.Application.Quit 'COM CTI2 // JATO 20210407
    oWord.Quit

    Dim doc As New Word.Application
    With doc
        .Documents.Open App.Path & "\SPOOLER\" & psCtaCod & ".doc" 'abrimos "Mi documento"
        .Visible = True 'hacemos visible Word
    End With
    'FIN MEJORAS JHCU
    'oDoc.SaveAs App.Path & "\SPOOLER\" & psCtaCod & ".doc" 'comentado por JHCU
End Sub
''CTI2 // JATO 20210407
Sub compForCartillaP(MatTitular As Variant, ByVal oWord As Object, Optional isNanito As Boolean = False)
 Dim k As Integer
 
    Dim lsNom1 As String, lsNom2 As String, lsNom3 As String, lsNom4 As String
    Dim lsDoc1 As String, lsDoc2 As String, lsDoc3 As String, lsDoc4 As String
    Dim lsDir1 As String, lsDir2 As String, lsDir3 As String, lsDir4 As String
    Dim lsFir1 As String, lsFir2 As String, lsFir3 As String, lsFir4 As String
    Dim lsCorreo1 As String, lsCorreo2 As String, lsCorreo3 As String, lsCorreo4 As String
Dim lsRepre1 As String, lsCtaRuc1 As String
Dim lsDepProvDis1 As String

Dim nNroPers As Integer
nNroPers = UBound(MatTitular) - 1

'If UBound(MatTitular) = 2 Then
If isNanito = False Then
    lsNom1 = "Personería :" & Space(3) & IIf(MatTitular(1, 5) = 1, "PERSONA NATURAL", "PERSONA JURIDICA") & Space(30) & "Razón Social :" & Space(3) & IIf(MatTitular(1, 5) = 1, "----------", MatTitular(1, 1))
    lsDoc1 = "Tipo de Cuenta :" & Space(3) & Trim(Mid(MatTitular(1, 12), 1, 20)) & Space(30) & "RUC : " & Space(3) & IIf(MatTitular(1, 5) = 1, "----------", MatTitular(1, 2)) & Space(30) & "Regla de Poderes : " & Space(3) & IIf(MatTitular(1, 5) = 1, MatTitular(1, 13), MatTitular(1, 13))
                    
    With oWord.Selection.Find
        .Text = "<<PersRazon1>>"
        .Replacement.Text = lsNom1
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<CtaRucRegla1>>"
        .Replacement.Text = lsDoc1
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
Else
    lsNom1 = "Personería :" & Space(3) & IIf(MatTitular(1, 5) = 1, "PERSONA NATURAL", "-----") & Space(30) & "Tipo de Cuenta :" & Space(3) & Trim(Mid(MatTitular(1, 12), 1, 20))
    With oWord.Selection.Find
        .Text = "<<PersRazon1>>"
        .Replacement.Text = lsNom1
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<CtaRucRegla1>>"
        .Forward = True
        .ClearFormatting
        .MatchWholeWord = True
        .MatchCase = False
        .Wrap = wdFindContinue
        .Execute
        oWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
        oWord.Selection.TypeBackspace
    End With
End If

If MatTitular(1, 5) = 1 Then 'PERSONA NATURAL

        k = 1
        If k <= nNroPers Then
        'Natural
            If isNanito = False Then
            lsNom1 = "Titular o Representante Legal :" & Space(3) & MatTitular(k, 1) & Space(30) & "Grupo:" & Space(3) & MatTitular(k, 10)
            lsDoc1 = "DOI/RUC" & Space(38) & ": " & Space(3) & MatTitular(k, 2) & Space(30) & "Teléfono:" & Space(3) & MatTitular(k, 9)
            Else
            lsNom1 = "Nombres y Apellidos" & Space(16) & ": " & Space(3) & MatTitular(k, 1) & Space(30) & "Relación:" & Space(3) & MatTitular(k, 4)
            lsDoc1 = "Documento" & Space(33) & ": " & Space(3) & "DOI - " & MatTitular(k, 2) & Space(30) & "Teléfono:" & Space(3) & MatTitular(k, 9)
            End If
            lsCorreo1 = "Correo Electrónico" & Space(20) & ":" & Space(3) & MatTitular(k, 11)
            lsDir1 = "Dirección Legal" & Space(26) & ": " & Space(3) & MatTitular(k, 3)
            lsDepProvDis1 = "Departamento:" & Space(3) & MatTitular(k, 6) & Space(30) & "Provincia:" & Space(3) & MatTitular(k, 7) & Space(30) & "Distrito:" & Space(3) & MatTitular(k, 8)
        
            With oWord.Selection.Find
                .Text = "<<Vacio1>>"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With oWord.Selection.Find
                .Text = "<<DatRepre1>>"
                .Replacement.Text = lsNom1
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With oWord.Selection.Find
                .Text = "<<DatDocu1>>"
                .Replacement.Text = lsDoc1
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With oWord.Selection.Find
                .Text = "<<DatCorreo1>>"
                .Replacement.Text = lsCorreo1
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With oWord.Selection.Find
                .Text = "<<DatDirecc1>>"
                .Replacement.Text = lsDir1
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With oWord.Selection.Find
                .Text = "<<DepProvDis1>>"
                .Replacement.Text = lsDepProvDis1
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
        
            '******FIRMAS AL FINAL DEL DOCUMENTO *******
            If isNanito = False Then
                lsFir1 = "Firma:______________________"
                
                lsNom1 = MatTitular(k, 1)
                lsDoc1 = "DOI/RUC: " & Space(3) & MatTitular(k, 2)
                lsDir1 = ""
            Else
              lsFir1 = ""
              lsNom1 = ""
              lsDoc1 = ""
              lsDir1 = ""
            End If
            
           
            '**********************************
            
            With oWord.Selection.Find
                .Text = "<<Firma1>>"
                .Replacement.Text = lsFir1
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With oWord.Selection.Find
                .Text = "<<NomTit1>>"
                .Replacement.Text = lsNom1
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With oWord.Selection.Find
                .Text = "<<DocTit1>>"
                .Replacement.Text = lsDoc1
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With oWord.Selection.Find
                .Text = "<<RazonTit1>>"
                .Replacement.Text = lsDir1
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
                
            k = k + 1  '2 nNroPers
            If (k <= nNroPers) Then
                If isNanito = False Then
                    lsNom1 = "Titular o Representante Legal :" & Space(3) & MatTitular(k, 1) & Space(30) & "Grupo:" & Space(3) & MatTitular(k, 10)
                Else
                    lsNom1 = "Apoderado o Representante Legal :" & Space(3) & MatTitular(k, 1) & Space(30) & "Relación:" & Space(3) & MatTitular(k, 4)
                End If
                lsDoc1 = "DOI/RUC" & Space(38) & ": " & Space(3) & MatTitular(k, 2) & Space(30) & "Teléfono:" & Space(3) & MatTitular(k, 9)
                lsCorreo1 = "Correo Electrónico" & Space(20) & ":" & Space(3) & MatTitular(k, 11)
                lsDir1 = "Dirección Legal" & Space(26) & ": " & Space(3) & MatTitular(k, 3)
                lsDepProvDis1 = "Departamento:" & Space(3) & MatTitular(k, 6) & Space(30) & "Provincia:" & Space(3) & MatTitular(k, 7) & Space(30) & "Distrito:" & Space(3) & MatTitular(k, 8)
                
                With oWord.Selection.Find
                    .Text = "<<Vacio2>>"
                    .Replacement.Text = ""
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DatRepre2>>"
                    .Replacement.Text = lsNom1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DatDocu2>>"
                    .Replacement.Text = lsDoc1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DatCorreo2>>"
                    .Replacement.Text = lsCorreo1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DatDirecc2>>"
                    .Replacement.Text = lsDir1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DepProvDis2>>"
                    .Replacement.Text = lsDepProvDis1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                
                '******FIRMAS AL FINAL DEL DOCUMENTO *******
                lsFir1 = "Firma:______________________"
                lsNom1 = MatTitular(k, 1)
                lsDoc1 = "DOI/RUC: " & Space(3) & MatTitular(k, 2)
                lsDir1 = ""
                '**********************************
                With oWord.Selection.Find
                .Text = "<<Vacio5>>"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
                End With 'jato
                With oWord.Selection.Find
                    .Text = "<<Firma2>>"
                    .Replacement.Text = lsFir1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<NomTit2>>"
                    .Replacement.Text = lsNom1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DocTit2>>"
                    .Replacement.Text = lsDoc1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<RazonTit2>>"
                    .Replacement.Text = lsDir1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                
            Else
              Call compRellenarWord(k, oWord)
            End If
            
            k = k + 1 '3
            If (k <= nNroPers) Then
                If isNanito = False Then
                    lsNom1 = "Titular o Representante Legal :" & Space(3) & MatTitular(k, 1) & Space(30) & "Grupo:" & Space(3) & MatTitular(k, 10)
                Else
                    lsNom1 = "Apoderado o Representante Legal :" & Space(3) & MatTitular(k, 1) & Space(30) & "Relación:" & Space(3) & MatTitular(k, 4)
                End If
                lsDoc1 = "DOI/RUC" & Space(38) & ": " & Space(3) & MatTitular(k, 2) & Space(30) & "Teléfono:" & Space(3) & MatTitular(k, 9)
                lsCorreo1 = "Correo Electrónico" & Space(20) & ":" & Space(3) & MatTitular(k, 11)
                lsDir1 = "Dirección Legal" & Space(26) & ": " & Space(3) & MatTitular(k, 3)
                lsDepProvDis1 = "Departamento:" & Space(3) & MatTitular(k, 6) & Space(30) & "Provincia:" & Space(3) & MatTitular(k, 7) & Space(30) & "Distrito:" & Space(3) & MatTitular(k, 8)
                
                With oWord.Selection.Find
                    .Text = "<<Vacio3>>"
                    .Replacement.Text = ""
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DatRepre3>>"
                    .Replacement.Text = lsNom1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DatDocu3>>"
                    .Replacement.Text = lsDoc1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DatCorreo3>>"
                    .Replacement.Text = lsCorreo1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DatDirecc3>>"
                    .Replacement.Text = lsDir1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DepProvDis3>>"
                    .Replacement.Text = lsDepProvDis1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                '******FIRMAS AL FINAL DEL DOCUMENTO *******
                lsFir1 = "Firma:______________________"
                lsNom1 = MatTitular(k, 1)
                lsDoc1 = "DOI/RUC: " & Space(3) & MatTitular(k, 2)
                lsDir1 = ""
                '**********************************
                With oWord.Selection.Find
                .Text = "<<Vacio6>>"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
                End With 'jato
                With oWord.Selection.Find
                    .Text = "<<Firma3>>"
                    .Replacement.Text = lsFir1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<NomTit3>>"
                    .Replacement.Text = lsNom1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DocTit3>>"
                    .Replacement.Text = lsDoc1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<RazonTit3>>"
                    .Replacement.Text = lsDir1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
            Else
              Call compRellenarWord(k, oWord)
            End If
        
            k = k + 1 '4
            If (k <= nNroPers) Then
                If isNanito = False Then
                    lsNom1 = "Titular o Representante Legal :" & Space(3) & MatTitular(k, 1) & Space(30) & "Grupo:" & Space(3) & MatTitular(k, 10)
                Else
                    lsNom1 = "Apoderado o Representante Legal :" & Space(3) & MatTitular(k, 1) & Space(30) & "Relación:" & Space(3) & MatTitular(k, 4)
                End If
                lsDoc1 = "DOI/RUC" & Space(38) & ": " & Space(3) & MatTitular(k, 2) & Space(30) & "Teléfono:" & Space(3) & MatTitular(k, 9)
                lsCorreo1 = "Correo Electrónico" & Space(20) & ":" & Space(3) & MatTitular(k, 11)
                lsDir1 = "Dirección Legal" & Space(26) & ": " & Space(3) & MatTitular(k, 3)
                lsDepProvDis1 = "Departamento:" & Space(3) & MatTitular(k, 6) & Space(30) & "Provincia:" & Space(3) & MatTitular(k, 7) & Space(30) & "Distrito:" & Space(3) & MatTitular(k, 8)
                
                With oWord.Selection.Find
                    .Text = "<<Vacio4>>"
                    .Replacement.Text = ""
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DatRepre4>>"
                    .Replacement.Text = lsNom1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DatDocu4>>"
                    .Replacement.Text = lsDoc1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DatCorreo4>>"
                    .Replacement.Text = lsCorreo1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DatDirecc4>>"
                    .Replacement.Text = lsDir1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DepProvDis4>>"
                    .Replacement.Text = lsDepProvDis1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                
                '******FIRMAS AL FINAL DEL DOCUMENTO *******
                lsFir1 = "Firma:______________________"
                lsNom1 = MatTitular(k, 1)
                lsDoc1 = "DOI/RUC: " & Space(3) & MatTitular(k, 2)
                lsDir1 = ""
                '**********************************
                With oWord.Selection.Find
                .Text = "<<Vacio7>>"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
                End With 'jato
                With oWord.Selection.Find
                    .Text = "<<Firma4>>"
                    .Replacement.Text = lsFir1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<NomTit4>>"
                    .Replacement.Text = lsNom1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DocTit4>>"
                    .Replacement.Text = lsDoc1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<RazonTit4>>"
                    .Replacement.Text = lsDir1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                
            Else
              Call compRellenarWord(k, oWord)
            End If
        End If

Else  'PERSONA JURIDICA
        
        k = 2
        If k <= nNroPers Then
        'Natural
            lsNom1 = "Titular o Representante Legal :" & Space(3) & MatTitular(k, 1) & Space(30) & "Grupo:" & Space(3) & MatTitular(k, 10)
            lsDoc1 = "DOI/RUC" & Space(38) & ": " & Space(3) & MatTitular(k, 2) & Space(30) & "Teléfono:" & Space(3) & MatTitular(k, 9)
            lsCorreo1 = "Correo Electrónico" & Space(20) & ":" & Space(3) & MatTitular(k, 11)
            lsDir1 = "Dirección Legal" & Space(26) & ": " & Space(3) & MatTitular(k, 3)
            lsDepProvDis1 = "Departamento:" & Space(3) & MatTitular(k, 6) & Space(30) & "Provincia:" & Space(3) & MatTitular(k, 7) & Space(30) & "Distrito:" & Space(3) & MatTitular(k, 8)
        
            With oWord.Selection.Find
                .Text = "<<Vacio1>>"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With oWord.Selection.Find
                .Text = "<<DatRepre1>>"
                .Replacement.Text = lsNom1
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With oWord.Selection.Find
                .Text = "<<DatDocu1>>"
                .Replacement.Text = lsDoc1
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With oWord.Selection.Find
                .Text = "<<DatCorreo1>>"
                .Replacement.Text = lsCorreo1
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With oWord.Selection.Find
                .Text = "<<DatDirecc1>>"
                .Replacement.Text = lsDir1
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With oWord.Selection.Find
                .Text = "<<DepProvDis1>>"
                .Replacement.Text = lsDepProvDis1
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
        
            '******FIRMAS AL FINAL DEL DOCUMENTO *******
            lsFir1 = "Firma:______________________"
            lsNom1 = MatTitular(k, 1)
            lsDoc1 = "DOI/RUC: " & Space(3) & MatTitular(1, 2)
            lsDir1 = MatTitular(1, 1)
            '**********************************
        
            With oWord.Selection.Find
                .Text = "<<Firma1>>"
                .Replacement.Text = lsFir1
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With oWord.Selection.Find
                .Text = "<<NomTit1>>"
                .Replacement.Text = lsNom1
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With oWord.Selection.Find
                .Text = "<<DocTit1>>"
                .Replacement.Text = lsDoc1
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With oWord.Selection.Find
                .Text = "<<RazonTit1>>"
                .Replacement.Text = lsDir1
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
                
            k = k + 1  '2 nNroPers
            If (k <= nNroPers) Then
                lsNom1 = "Titular o Representante Legal :" & Space(3) & MatTitular(k, 1) & Space(30) & "Grupo:" & Space(3) & MatTitular(k, 10)
                lsDoc1 = "DOI/RUC" & Space(38) & ": " & Space(3) & MatTitular(k, 2) & Space(30) & "Teléfono:" & Space(3) & MatTitular(k, 9)
                lsCorreo1 = "Correo Electrónico" & Space(20) & ":" & Space(3) & MatTitular(k, 11)
                lsDir1 = "Dirección Legal" & Space(26) & ": " & Space(3) & MatTitular(k, 3)
                lsDepProvDis1 = "Departamento:" & Space(3) & MatTitular(k, 6) & Space(30) & "Provincia:" & Space(3) & MatTitular(k, 7) & Space(30) & "Distrito:" & Space(3) & MatTitular(k, 8)
                
                With oWord.Selection.Find
                    .Text = "<<Vacio2>>"
                    .Replacement.Text = ""
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DatRepre2>>"
                    .Replacement.Text = lsNom1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DatDocu2>>"
                    .Replacement.Text = lsDoc1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DatCorreo2>>"
                    .Replacement.Text = lsCorreo1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DatDirecc2>>"
                    .Replacement.Text = lsDir1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DepProvDis2>>"
                    .Replacement.Text = lsDepProvDis1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                
                '******FIRMAS AL FINAL DEL DOCUMENTO *******
                lsFir1 = "Firma:______________________"
                lsNom1 = MatTitular(k, 1)
                lsDoc1 = "DOI/RUC: " & Space(3) & MatTitular(1, 2)
                lsDir1 = MatTitular(1, 1)
                '**********************************
                
                With oWord.Selection.Find
                .Text = "<<Vacio5>>"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
                End With 'jato
                With oWord.Selection.Find
                    .Text = "<<Firma2>>"
                    .Replacement.Text = lsFir1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<NomTit2>>"
                    .Replacement.Text = lsNom1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DocTit2>>"
                    .Replacement.Text = lsDoc1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<RazonTit2>>"
                    .Replacement.Text = lsDir1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                
            Else
              Call compRellenarWord(k - 1, oWord)
            End If
            
            k = k + 1 '3
            If (k <= nNroPers) Then
                lsNom1 = "Titular o Representante Legal :" & Space(3) & MatTitular(k, 1) & Space(30) & "Grupo:" & Space(3) & MatTitular(k, 10)
                lsDoc1 = "DOI/RUC" & Space(38) & ": " & Space(3) & MatTitular(k, 2) & Space(30) & "Teléfono:" & Space(3) & MatTitular(k, 9)
                lsCorreo1 = "Correo Electrónico" & Space(20) & ":" & Space(3) & MatTitular(k, 11)
                lsDir1 = "Dirección Legal" & Space(26) & ": " & Space(3) & MatTitular(k, 3)
                lsDepProvDis1 = "Departamento:" & Space(3) & MatTitular(k, 6) & Space(30) & "Provincia:" & Space(3) & MatTitular(k, 7) & Space(30) & "Distrito:" & Space(3) & MatTitular(k, 8)
                
                With oWord.Selection.Find
                    .Text = "<<Vacio3>>"
                    .Replacement.Text = ""
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DatRepre3>>"
                    .Replacement.Text = lsNom1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DatDocu3>>"
                    .Replacement.Text = lsDoc1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DatCorreo3>>"
                    .Replacement.Text = lsCorreo1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DatDirecc3>>"
                    .Replacement.Text = lsDir1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DepProvDis3>>"
                    .Replacement.Text = lsDepProvDis1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                '******FIRMAS AL FINAL DEL DOCUMENTO *******
                lsFir1 = "Firma:______________________"
                lsNom1 = MatTitular(k, 1)
                lsDoc1 = "DOI/RUC: " & Space(3) & MatTitular(1, 2)
                lsDir1 = MatTitular(1, 1)
                '**********************************
                With oWord.Selection.Find
                .Text = "<<Vacio6>>"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
                End With 'jato
                With oWord.Selection.Find
                    .Text = "<<Firma3>>"
                    .Replacement.Text = lsFir1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<NomTit3>>"
                    .Replacement.Text = lsNom1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DocTit3>>"
                    .Replacement.Text = lsDoc1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<RazonTit3>>"
                    .Replacement.Text = lsDir1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
            Else
              Call compRellenarWord(k - 1, oWord)
            End If
        
            k = k + 1 '4
            If (k <= nNroPers) Then
                lsNom1 = "Titular o Representante Legal :" & Space(3) & MatTitular(k, 1) & Space(30) & "Grupo:" & Space(3) & MatTitular(k, 10)
                lsDoc1 = "DOI/RUC" & Space(38) & ": " & Space(3) & MatTitular(k, 2) & Space(30) & "Teléfono:" & Space(3) & MatTitular(k, 9)
                lsCorreo1 = "Correo Electrónico" & Space(20) & ":" & Space(3) & MatTitular(k, 11)
                lsDir1 = "Dirección Legal" & Space(26) & ": " & Space(3) & MatTitular(k, 3)
                lsDepProvDis1 = "Departamento:" & Space(3) & MatTitular(k, 6) & Space(30) & "Provincia:" & Space(3) & MatTitular(k, 7) & Space(30) & "Distrito:" & Space(3) & MatTitular(k, 8)
                
                With oWord.Selection.Find
                    .Text = "<<Vacio4>>"
                    .Replacement.Text = ""
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DatRepre4>>"
                    .Replacement.Text = lsNom1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DatDocu4>>"
                    .Replacement.Text = lsDoc1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DatCorreo4>>"
                    .Replacement.Text = lsCorreo1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DatDirecc4>>"
                    .Replacement.Text = lsDir1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DepProvDis4>>"
                    .Replacement.Text = lsDepProvDis1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                
                '******FIRMAS AL FINAL DEL DOCUMENTO *******
                lsFir1 = "Firma:______________________"
                lsNom1 = MatTitular(k, 1)
                lsDoc1 = "DOI/RUC: " & Space(3) & MatTitular(1, 2)
                lsDir1 = MatTitular(1, 1)
                '**********************************
                
                With oWord.Selection.Find
                .Text = "<<Vacio7>>"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
                End With 'jato
                With oWord.Selection.Find
                    .Text = "<<Firma4>>"
                    .Replacement.Text = lsFir1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<NomTit4>>"
                    .Replacement.Text = lsNom1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<DocTit4>>"
                    .Replacement.Text = lsDoc1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<RazonTit4>>"
                    .Replacement.Text = lsDir1
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                
            Else
              Call compRellenarWord(k - 1, oWord)
            End If
        End If

End If

End Sub

Sub compRellenarWord(ByVal nRoPers As Integer, ByVal noWord As Object)

If nRoPers = 2 Then
'**** 2
        With noWord.Selection.Find
            .Text = "<<Vacio2>>"
            .Forward = True
            .ClearFormatting
            .MatchWholeWord = True
            .MatchCase = False
            .Wrap = wdFindContinue
            .Execute
            noWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
            noWord.Selection.TypeBackspace
        End With
        With noWord.Selection.Find
            .Text = "<<DatRepre2>>"
            .Forward = True
            .ClearFormatting
            .MatchWholeWord = True
            .MatchCase = False
            .Wrap = wdFindContinue
            .Execute
            noWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
            noWord.Selection.TypeBackspace
        End With

        With noWord.Selection.Find
            .Text = "<<DatDocu2>>"
            .Forward = True
            .ClearFormatting
            .MatchWholeWord = True
            .MatchCase = False
            .Wrap = wdFindContinue
            .Execute
            noWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
            noWord.Selection.TypeBackspace
        End With

        With noWord.Selection.Find
            .Text = "<<DatCorreo2>>"
            .Forward = True
            .ClearFormatting
            .MatchWholeWord = True
            .MatchCase = False
            .Wrap = wdFindContinue
            .Execute
            noWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
            noWord.Selection.TypeBackspace
        End With

        With noWord.Selection.Find
            .Text = "<<DatDirecc2>>"
            .Forward = True
            .ClearFormatting
            .MatchWholeWord = True
            .MatchCase = False
            .Wrap = wdFindContinue
            .Execute
            noWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
            noWord.Selection.TypeBackspace
        End With

        With noWord.Selection.Find
            .Text = "<<DepProvDis2>>"
            .Forward = True
            .ClearFormatting
            .MatchWholeWord = True
            .MatchCase = False
            .Wrap = wdFindContinue
            .Execute
            noWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
            noWord.Selection.TypeBackspace
        End With
'fIRMAS 2:
        With noWord.Selection.Find
            .Text = "<<Vacio5>>"
            .Forward = True
            .ClearFormatting
            .MatchWholeWord = True
            .MatchCase = False
            .Wrap = wdFindContinue
            .Execute
            noWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
            noWord.Selection.TypeBackspace
        End With 'jato
        With noWord.Selection.Find
            .Text = "<<Firma2>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With

        With noWord.Selection.Find
            .Text = "<<NomTit2>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With

        With noWord.Selection.Find
            .Text = "<<DocTit2>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With noWord.Selection.Find
            .Text = "<<RazonTit2>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With

ElseIf nRoPers = 3 Then

'****3
        With noWord.Selection.Find
            .Text = "<<Vacio3>>"
            .Forward = True
            .ClearFormatting
            .MatchWholeWord = True
            .MatchCase = False
            .Wrap = wdFindContinue
            .Execute
            noWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
            noWord.Selection.TypeBackspace
        End With
        With noWord.Selection.Find
            .Text = "<<DatRepre3>>"
            .Forward = True
            .ClearFormatting
            .MatchWholeWord = True
            .MatchCase = False
            .Wrap = wdFindContinue
            .Execute
            noWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
            noWord.Selection.TypeBackspace
        End With

        With noWord.Selection.Find
            .Text = "<<DatDocu3>>"
            .Forward = True
            .ClearFormatting
            .MatchWholeWord = True
            .MatchCase = False
            .Wrap = wdFindContinue
            .Execute
            noWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
            noWord.Selection.TypeBackspace
        End With

        With noWord.Selection.Find
            .Text = "<<DatCorreo3>>"
            .Forward = True
            .ClearFormatting
            .MatchWholeWord = True
            .MatchCase = False
            .Wrap = wdFindContinue
            .Execute
            noWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
            noWord.Selection.TypeBackspace
        End With

        With noWord.Selection.Find
            .Text = "<<DatDirecc3>>"
            .Forward = True
            .ClearFormatting
            .MatchWholeWord = True
            .MatchCase = False
            .Wrap = wdFindContinue
            .Execute
            noWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
            noWord.Selection.TypeBackspace
        End With

        With noWord.Selection.Find
            .Text = "<<DepProvDis3>>"
            .Forward = True
            .ClearFormatting
            .MatchWholeWord = True
            .MatchCase = False
            .Wrap = wdFindContinue
            .Execute
            noWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
            noWord.Selection.TypeBackspace
        End With
'****4

'        With noWord.Selection.Find
'            .Text = "<<Vacio4>>"
'            .Forward = True
'            .ClearFormatting
'            .MatchWholeWord = True
'            .MatchCase = False
'            .Wrap = wdFindContinue
'            .Execute
'            noWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
'            noWord.Selection.TypeBackspace
'        End With
'        With noWord.Selection.Find
'            .Text = "<<DatRepre4>>"
'            .Forward = True
'            .ClearFormatting
'            .MatchWholeWord = True
'            .MatchCase = False
'            .Wrap = wdFindContinue
'            .Execute
'            noWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
'            noWord.Selection.TypeBackspace
'        End With
'
'        With noWord.Selection.Find
'            .Text = "<<DatDocu4>>"
'            .Forward = True
'            .ClearFormatting
'            .MatchWholeWord = True
'            .MatchCase = False
'            .Wrap = wdFindContinue
'            .Execute
'            noWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
'            noWord.Selection.TypeBackspace
'        End With
'
'        With noWord.Selection.Find
'            .Text = "<<DatCorreo4>>"
'            .Forward = True
'            .ClearFormatting
'            .MatchWholeWord = True
'            .MatchCase = False
'            .Wrap = wdFindContinue
'            .Execute
'            noWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
'            noWord.Selection.TypeBackspace
'        End With
'
'        With noWord.Selection.Find
'            .Text = "<<DatDirecc4>>"
'            .Forward = True
'            .ClearFormatting
'            .MatchWholeWord = True
'            .MatchCase = False
'            .Wrap = wdFindContinue
'            .Execute
'            noWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
'            noWord.Selection.TypeBackspace
'        End With
'
'        With noWord.Selection.Find
'            .Text = "<<DepProvDis4>>"
'            .Forward = True
'            .ClearFormatting
'            .MatchWholeWord = True
'            .MatchCase = False
'            .Wrap = wdFindContinue
'            .Execute
'            noWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
'            noWord.Selection.TypeBackspace
'        End With
        
'***FIRMAS 3
        With noWord.Selection.Find
            .Text = "<<Vacio6>>"
            .Forward = True
            .ClearFormatting
            .MatchWholeWord = True
            .MatchCase = False
            .Wrap = wdFindContinue
            .Execute
            noWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
            noWord.Selection.TypeBackspace
        End With 'jato
        With noWord.Selection.Find
            .Text = "<<Firma3>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With

        With noWord.Selection.Find
            .Text = "<<NomTit3>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With

        With noWord.Selection.Find
            .Text = "<<DocTit3>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With noWord.Selection.Find
            .Text = "<<RazonTit3>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
''***4
'        With noWord.Selection.Find
'            .Text = "<<Firma4>>"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'
'        With noWord.Selection.Find
'            .Text = "<<NomTit4>>"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'
'        With noWord.Selection.Find
'            .Text = "<<DocTit4>>"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With noWord.Selection.Find
'            .Text = "<<RazonTit4>>"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With

ElseIf nRoPers = 4 Then

        With noWord.Selection.Find
            .Text = "<<Vacio4>>"
            .Forward = True
            .ClearFormatting
            .MatchWholeWord = True
            .MatchCase = False
            .Wrap = wdFindContinue
            .Execute
            noWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
            noWord.Selection.TypeBackspace
        End With
        With noWord.Selection.Find
            .Text = "<<DatRepre4>>"
            .Forward = True
            .ClearFormatting
            .MatchWholeWord = True
            .MatchCase = False
            .Wrap = wdFindContinue
            .Execute
            noWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
            noWord.Selection.TypeBackspace
        End With

        With noWord.Selection.Find
            .Text = "<<DatDocu4>>"
            .Forward = True
            .ClearFormatting
            .MatchWholeWord = True
            .MatchCase = False
            .Wrap = wdFindContinue
            .Execute
            noWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
            noWord.Selection.TypeBackspace
        End With

        With noWord.Selection.Find
            .Text = "<<DatCorreo4>>"
            .Forward = True
            .ClearFormatting
            .MatchWholeWord = True
            .MatchCase = False
            .Wrap = wdFindContinue
            .Execute
            noWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
            noWord.Selection.TypeBackspace
        End With

        With noWord.Selection.Find
            .Text = "<<DatDirecc4>>"
            .Forward = True
            .ClearFormatting
            .MatchWholeWord = True
            .MatchCase = False
            .Wrap = wdFindContinue
            .Execute
            noWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
            noWord.Selection.TypeBackspace
        End With

        With noWord.Selection.Find
            .Text = "<<DepProvDis4>>"
            .Forward = True
            .ClearFormatting
            .MatchWholeWord = True
            .MatchCase = False
            .Wrap = wdFindContinue
            .Execute
            noWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
            noWord.Selection.TypeBackspace
        End With
'***FIRMAS 4
        With noWord.Selection.Find
            .Text = "<<Vacio7>>"
            .Forward = True
            .ClearFormatting
            .MatchWholeWord = True
            .MatchCase = False
            .Wrap = wdFindContinue
            .Execute
            noWord.Selection.EndOf Unit:=wdLine, Extend:=wdExtend
            noWord.Selection.TypeBackspace
        End With 'jato
        With noWord.Selection.Find
            .Text = "<<Firma4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With

        With noWord.Selection.Find
            .Text = "<<NomTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With

        With noWord.Selection.Find
            .Text = "<<DocTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With noWord.Selection.Find
            .Text = "<<RazonTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With

End If
''FIN CTI2 // JATO 20210407
End Sub

Public Sub ImpreCartillaPlazoFijo(MatTitular As Variant, ByVal psCtaCod As String, ByVal pnTasa As Double, ByVal pnMonto As Double, _
                                  ByVal pnPlazo As Integer, ByVal pdFechaA As Date, ByVal pnFormaRetiro As Integer, Optional ByVal pnTotIntMes As Double, Optional ByVal pnTpoProgram As Integer = 0)
                                  'BRGO 20111230 - Se agregó el parámetro "pnTpoProgram"
       'By Capi 20042008 se agrego parametro pnTotIntMes

    Dim rs As ADODB.Recordset
    Dim nValor As Double
    Dim lsFechas As String
    Dim lsNom1 As String, lsNom2 As String, lsNom3 As String, lsNom4 As String
    Dim lsDoc1 As String, lsDoc2 As String, lsDoc3 As String, lsDoc4 As String
    Dim lsDir1 As String, lsDir2 As String, lsDir3 As String, lsDir4 As String
    Dim lsCad As String
    Dim nitf As Double
    Dim sTaInEf As String
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
    Dim clsMantA As COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMantenimiento
    Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion 'CTI2 FERIMORO ERS030-2019
        Dim lnPlazo As Long
    Dim nInteres As Double
    Dim sPlazo As String
    Dim lnTasa As Double
    Dim nCosEnvioFisico As String 'JUEZ 20130520
    Dim sTipoEnvio As String 'JUEZ
        
        
        Dim nParMontoMinSol As Double 'CTI2 FERIMORO ERS030-2019
    Dim nParMontoMinDol As Double 'CTI2 FERIMORO ERS030-2019


    lnPlazo = CLng(pnPlazo)
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    'sTaInEf = Format$(((((((pnTasa / 12) / 3600) + 1) ^ 360) - 1) * 100), "#0.00")
    sTaInEf = pnTasa
    lnTasa = clsMant.GetTasaNominal(pnTasa, 360)
        
         'CTI2 ferimoro: ERS030-2019 ************************
    Dim nSubProducto As String
    nSubProducto = clsMant.recuperaSubProducto(gCapPlazoFijo, pnTpoProgram)
    
    Dim rsPar As ADODB.Recordset
    Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
    Set rsPar = clsDef.GetCapParametroNew(gCapPlazoFijo, pnTpoProgram)
    nParMontoMinSol = rsPar!nAumCapMinSol 'MOD BY JATO 20210503
    nParMontoMinDol = rsPar!nAumCapMinDol 'MOD BY JATO 20210503
    Set clsDef = Nothing
    '**************************************************

    If pnFormaRetiro = 4 Then
        sPlazo = "el mismo Dia"
        Set clsMantA = New COMNCaptaGenerales.NCOMCaptaMovimiento
            nInteres = Format$(clsMantA.GetInteres(pnMonto, lnTasa, lnPlazo, TpoCalcIntAdelantado), "#,##0.00")
        Set clsMantA = Nothing
    'By capi 07032008
    ElseIf pnFormaRetiro = 3 Then
        sPlazo = "Libre"
        nInteres = Format$(clsMant.GetInteresPF(lnTasa, pnMonto, lnPlazo), "#,##0.00")

    ElseIf pnFormaRetiro = 1 Then
        sPlazo = "Mensual"
        'By Capi 20042008
        'nInteres = Format$(clsMant.GetInteresPF(lnTasa, pnMonto, lnPlazo), "#,##0.00")
        nInteres = Format$(pnTotIntMes, "#,##0.00")
    '
    Else
        sPlazo = "al Final del Periodo"
        nInteres = Format$(clsMant.GetInteresPF(lnTasa, pnMonto, lnPlazo), "#,##0.00")
    End If
    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range

    Dim loRs As COMDConstSistema.DCOMGeneral
    Set loRs = New COMDConstSistema.DCOMGeneral

    sPlazo = loRs.GetConstante(gCaptacPFFormaRetiro, , CStr(pnFormaRetiro), 1)!CDescripcion 'JUEZ 20150121

    Dim loAge As COMDConstantes.DCOMAgencias
    Dim rs1 As ADODB.Recordset
    Dim lsAgencia As String
    Dim lsAgeDir As String
    Set loAge = New COMDConstantes.DCOMAgencias
    Dim lsDesAgencia As String 'CTI2 FERIMORO : ERS030-2019
    Set rs1 = New ADODB.Recordset
        Set rs1 = loAge.RecuperaAgencias(gsCodAge)
        If Not (rs1.EOF And rs1.BOF) Then
            lsAgencia = Trim(rs1("cUbiGeoDescripcion"))
            lsDesAgencia = Trim(rs1("cAgeDescripcion")) 'CTI2 FERIMORO : ERS030-2019
            lsAgeDir = rs1("cAgeDireccion")
        End If
    Set loAge = Nothing
    'JUEZ 20130520 ***************************************
    Set rs = New ADODB.Recordset
    nValor = loRs.GetParametro(2000, 1005)
    '''nCosEnvioFisico = "S/. " & Format$(nValor, "0.00") 'marg ers044-2016
    nCosEnvioFisico = gcPEN_SIMBOLO & " " & Format$(nValor, "0.00") 'marg ers044-2016
        'CTI2 ERS030-2019************************
    Dim nFormaEnvio As String
    Dim nEnvioSiNo As String
    '****************************************
    Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
    Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rs = oCap.RecuperaDatosEnvioEstadoCta(psCtaCod)
    ''CTI2 FERIMORO ERS030-2019
    ''If Not rs.EOF Then sTipoEnvio = rs!cTipoEnvio 'APRI20180530 ERS036-2017
    'CTI2 ERS030-2019************************
    If Not rs.EOF Then
        sTipoEnvio = rs!cTipoEnvio
        nFormaEnvio = rs!cTipoEnvio
        nEnvioSiNo = "Si"
    Else
            nFormaEnvio = ""
            nEnvioSiNo = "No"
    End If
    '************************
    'sTipoEnvio = rs!cTipoEnvio
    Set rs = Nothing
    'END JUEZ ********************************************
     'ALPA 20091118************************************************
    Dim objCaptac As COMNCaptaGenerales.NCOMCaptaGenerales
    Set objCaptac = New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim nTREA As Currency
    nTREA = objCaptac.ObtenerTREA(Mid$(psCtaCod, 6, 3), IIf(pnMonto = 0, 10, pnMonto), pnTasa, pnPlazo)
    '*************************************************************

    Set oWord = CreateObject("Word.Application")
    oWord.Visible = True
    '*** BRGO 20111230 ********************************
    If pnTpoProgram = 1 Then
        Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\CARTILLAPREMIUM.doc")
    Else
        Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\CARTILLAPF.doc")
    End If
    '**************************************************
        'CTI2 FERIMORO: ERS030-2019 ************************************************
    'FECHA DEL DOCUMENTO - CABECERA
    lsCad = lsDesAgencia
    With oWord.Selection.Find
        .Text = "<<Oficina>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    lsCad = Format(Time, "hh:mm:ss")
    With oWord.Selection.Find
        .Text = "<<hora>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    lsCad = gsCodUser
    With oWord.Selection.Find
        .Text = "<<user>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    lsCad = IIf(Mid(psCtaCod, 9, 1) = 1, "MN ", "ME ") & Format(nParMontoMinSol, "#,##0.00")
    With oWord.Selection.Find
        .Text = "<<MontoMin>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    '***********************************************************************
    'JUEZ 20150121 ************************************************
    'Cuenta
    lsCad = psCtaCod
    With oWord.Selection.Find
        .Text = "<<cCodCta>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END JUEZ *****************************************************

    'Monto
    lsCad = IIf(Mid(psCtaCod, 9, 1) = 1, "MN ", "ME ") & Format(pnMonto, "#,##0.00")
    With oWord.Selection.Find
        .Text = "<<Monto>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Plazo
    lsCad = ""
    lsCad = lnPlazo & " días"
    With oWord.Selection.Find
        .Text = "<<Plazo>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'ALPA 20091118****************************************
    lsCad = ""
    lsCad = nTREA & " % (Fija)"
    With oWord.Selection.Find
        .Text = "<<TasaTrea>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    '*****************************************************
    'Tasa
    lsCad = ""
    lsCad = sTaInEf & " % (Fija)"
    With oWord.Selection.Find
        .Text = "<<Tasa>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Monto de interes
    lsCad = ""
    '''lsCad = IIf(Mid(psCtaCod, 9, 1) = 1, "S/. ", "$. ") & Format(nInteres, "#,##0.00") 'marg ers044-2016
    lsCad = IIf(Mid(psCtaCod, 9, 1) = 1, gcPEN_SIMBOLO & " ", "$. ") & Format(nInteres, "#,##0.00") 'marg ers044-2016
    With oWord.Selection.Find
        .Text = "<<MonInteres>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Tiempo Retiro de Interes
    With oWord.Selection.Find
        .Text = "<<cPlazo>>"
        .Replacement.Text = sPlazo
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'Fecha Apertura
    With oWord.Selection.Find
        .Text = "<<FechaAper>>"
        .Replacement.Text = pdFechaA
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
        
            'Sub Producto
    With oWord.Selection.Find
        .Text = "<<SubProducto>>"
        .Replacement.Text = nSubProducto
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Fecha de Vencimiento
    Dim dFechaV As Date
    dFechaV = CDate(pdFechaA) + lnPlazo
    With oWord.Selection.Find
        .Text = "<<FechaVen>>"
        .Replacement.Text = dFechaV
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
'ALPA 20091119*************************************************
     'Fecha de Cancelacion
'    Dim dFechac As Date
'    dFechac = CDate(pdFechaA) + lnPlazo + 1
'    With oWord.Selection.Find
'        .Text = "<<FechaCan>>"
'        .Replacement.Text = dFechac
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'**************************************************************
    'JUEZ 20150121 ************************************************
    'Forma de Retiro
    lsCad = sPlazo
    With oWord.Selection.Find
        .Text = "<<cFormaRetiro>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END JUEZ *****************************************************
    'ITF
    nitf = CDbl(gnITFPorcent) * 100
    With oWord.Selection.Find
        .Text = "<<TasaITF>>"
        '.Replacement.Text = Format$(nitf, "0.00")
        .Replacement.Text = Trim(CStr(nitf))
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Fecha
    lsFechas = Format(gdFecSis, "dd") & " de " & Format(gdFecSis, "mmmm") & " del " & Format(gdFecSis, "yyyy")
    With oWord.Selection.Find
        .Text = "<<FecActual>>"
        .Replacement.Text = lsFechas
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

   'Lugar
    With oWord.Selection.Find
        .Text = "<<cLugar>>"
        .Replacement.Text = lsAgencia
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Direccion
    With oWord.Selection.Find
        .Text = "<<cDireccion>>"
        .Replacement.Text = lsAgeDir
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
         'CTI2 FERIMORO ERS030-2019
    With oWord.Selection.Find
        .Text = "<<Elección SI/NO>>"
        .Replacement.Text = nEnvioSiNo
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    '*****************************************
    'JUEZ 20130520 ****************************
    'Tipo envio estado de cuenta
    With oWord.Selection.Find
        .Text = "<<TipoEnvio>>"
        .Replacement.Text = sTipoEnvio
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'Costo envio fisico
    With oWord.Selection.Find
        .Text = "<<CostoEnvioFisico>>"
        .Replacement.Text = nCosEnvioFisico
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END JUEZ *********************************

    Dim i As Integer
    Dim rsDir As ADODB.Recordset 'JUEZ 20130520
    Dim oPers As COMDpersona.DCOMPersona 'JUEZ 20130520

      ''CTI2 FERIMORO ERS030-2019 : DOCUMENTOS VERSIONADOS
'***************************************************
Call compForCartillaP(MatTitular, oWord)
''**************************************************

''CTI2 FERIMORO ERS030-2019 : DOCUMENTOS VERSIONADOS
'***************************************************
''    If UBound(MatTitular) = 2 Then
''
''        lsNom1 = "Nombre del Cliente: " & MatTitular(1, 1)
''        lsDoc1 = "DNI/RUC: " & MatTitular(1, 2) & space(60) & "Firma:______________________"
''        lsDir1 = "Dirección: " & MatTitular(1, 3)
''        'JUEZ 20130520 ****************************
''        Set oPers = New COMDPersona.DCOMPersona
''        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(1, 2))
''        lsDir1 = "Dirección: " & rsDir!cPersDireccDomicilio
''        'END JUEZ *********************************
''
''        With oWord.Selection.Find
''            .Text = "<<NomTit1>>"
''            .Replacement.Text = lsNom1
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit1>>"
''            .Replacement.Text = lsDoc1
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit1>>"
''            .Replacement.Text = lsDir1
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<NomTit2>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit2>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit2>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<NomTit3>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit3>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit3>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<NomTit4>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit4>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit4>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''
''    ElseIf UBound(MatTitular) = 3 Then
''        lsNom1 = "Nombre del Cliente: " & MatTitular(1, 1)
''        lsDoc1 = "DNI/RUC: " & MatTitular(1, 2) & space(60) & "Firma:______________________"
''        lsDir1 = "Dirección: " & MatTitular(1, 3)
''        'JUEZ 20130520 ****************************
''        Set oPers = New COMDPersona.DCOMPersona
''        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(1, 2))
''        lsDir1 = "Dirección: " & rsDir!cPersDireccDomicilio
''        'END JUEZ *********************************
''        lsNom2 = "Nombre del Cliente: " & MatTitular(2, 1)
''        lsDoc2 = "DNI/RUC: " & MatTitular(2, 2) & space(60) & "Firma:______________________"
''        lsDir2 = "Dirección: " & MatTitular(2, 3)
''        'JUEZ 20130520 ****************************
''        Set oPers = New COMDPersona.DCOMPersona
''        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(2, 2))
''        lsDir2 = "Dirección: " & rsDir!cPersDireccDomicilio
''        'END JUEZ *********************************
''
''        With oWord.Selection.Find
''            .Text = "<<NomTit1>>"
''            .Replacement.Text = lsNom1
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit1>>"
''            .Replacement.Text = lsDoc1
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit1>>"
''            .Replacement.Text = lsDir1
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<NomTit2>>"
''            .Replacement.Text = lsNom2
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit2>>"
''            .Replacement.Text = lsDoc2
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit2>>"
''            .Replacement.Text = lsDir2
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<NomTit3>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit3>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit3>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<NomTit4>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit4>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit4>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''    ElseIf UBound(MatTitular) = 4 Then
''        lsNom1 = "Nombre del Cliente: " & MatTitular(1, 1)
''        lsDoc1 = "DNI/RUC: " & MatTitular(1, 2) & space(60) & "Firma:______________________"
''        lsDir1 = "Dirección: " & MatTitular(1, 3)
''        'JUEZ 20130520 ****************************
''        Set oPers = New COMDPersona.DCOMPersona
''        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(1, 2))
''        lsDir1 = "Dirección: " & rsDir!cPersDireccDomicilio
''        'END JUEZ *********************************
''        lsNom2 = "Nombre del Cliente: " & MatTitular(2, 1)
''        lsDoc2 = "DNI/RUC: " & MatTitular(2, 2) & space(60) & "Firma:______________________"
''        lsDir2 = "Dirección: " & MatTitular(2, 3)
''        'JUEZ 20130520 ****************************
''        Set oPers = New COMDPersona.DCOMPersona
''        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(2, 2))
''        lsDir1 = "Dirección: " & rsDir!cPersDireccDomicilio
''        'END JUEZ *********************************
''        lsNom3 = "Nombre del Cliente: " & MatTitular(3, 1)
''        lsDoc3 = "DNI/RUC: " & MatTitular(3, 2) & space(60) & "Firma:______________________"
''        lsDir3 = "Dirección: " & MatTitular(3, 3)
''        'JUEZ 20130520 ****************************
''        Set oPers = New COMDPersona.DCOMPersona
''        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(3, 2))
''        lsDir1 = "Dirección: " & rsDir!cPersDireccDomicilio
''        'END JUEZ *********************************
''
''
''        With oWord.Selection.Find
''            .Text = "<<NomTit1>>"
''            .Replacement.Text = lsNom1
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit1>>"
''            .Replacement.Text = lsDoc1
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit1>>"
''            .Replacement.Text = lsDir1
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<NomTit2>>"
''            .Replacement.Text = lsNom2
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit2>>"
''            .Replacement.Text = lsDoc2
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit2>>"
''            .Replacement.Text = lsDir2
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<NomTit3>>"
''            .Replacement.Text = lsNom3
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit3>>"
''            .Replacement.Text = lsDoc3
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit3>>"
''            .Replacement.Text = lsDir3
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<NomTit4>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit4>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit4>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''
''    ElseIf UBound(MatTitular) = 5 Then
''        lsNom1 = "Nombre del Cliente: " & MatTitular(1, 1)
''        lsDoc1 = "DNI/RUC: " & MatTitular(1, 2) & space(60) & "Firma:______________________"
''        lsDir1 = "Dirección: " & MatTitular(1, 3)
''        'JUEZ 20130520 ****************************
''        Set oPers = New COMDPersona.DCOMPersona
''        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(1, 2))
''        lsDir1 = "Dirección: " & rsDir!cPersDireccDomicilio
''        'END JUEZ *********************************
''        lsNom2 = "Nombre del Cliente: " & MatTitular(2, 1)
''        lsDoc2 = "DNI/RUC: " & MatTitular(2, 2) & space(60) & "Firma:______________________"
''        lsDir2 = "Dirección: " & MatTitular(2, 3)
''        'JUEZ 20130520 ****************************
''        Set oPers = New COMDPersona.DCOMPersona
''        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(2, 2))
''        lsDir2 = "Dirección: " & rsDir!cPersDireccDomicilio
''        'END JUEZ *********************************
''        lsNom3 = "Nombre del Cliente: " & MatTitular(3, 1)
''        lsDoc3 = "DNI/RUC: " & MatTitular(3, 2) & space(60) & "Firma:______________________"
''        lsDir3 = "Dirección: " & MatTitular(3, 3)
''        'JUEZ 20130520 ****************************
''        Set oPers = New COMDPersona.DCOMPersona
''        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(3, 2))
''        lsDir3 = "Dirección: " & rsDir!cPersDireccDomicilio
''        'END JUEZ *********************************
''        lsNom4 = "Nombre del Cliente: " & MatTitular(4, 1)
''        lsDoc4 = "DNI/RUC: " & MatTitular(4, 2) & space(60) & "Firma:______________________"
''        lsDir4 = "Dirección: " & MatTitular(4, 3)
''        'JUEZ 20130520 ****************************
''        Set oPers = New COMDPersona.DCOMPersona
''        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(4, 2))
''        lsDir4 = "Dirección: " & rsDir!cPersDireccDomicilio
''        'END JUEZ *********************************
''
''        With oWord.Selection.Find
''            .Text = "<<NomTit1>>"
''            .Replacement.Text = lsNom1
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit1>>"
''            .Replacement.Text = lsDoc1
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit1>>"
''            .Replacement.Text = lsDir1
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<NomTit2>>"
''            .Replacement.Text = lsNom2
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit2>>"
''            .Replacement.Text = lsDoc2
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit2>>"
''            .Replacement.Text = lsDir2
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<NomTit3>>"
''            .Replacement.Text = lsNom3
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit3>>"
''            .Replacement.Text = lsDoc3
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit3>>"
''            .Replacement.Text = lsDir3
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<NomTit4>>"
''            .Replacement.Text = lsNom4
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit4>>"
''            .Replacement.Text = lsDoc4
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit4>>"
''            .Replacement.Text = lsDir4
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''    End If



   oDoc.SaveAs App.Path & "\SPOOLER\" & psCtaCod & ".doc"
   
   oDoc.Close
   oWord.Quit
                Dim doc As New Word.Application
        With doc
            .Documents.Open App.Path & "\SPOOLER\" & psCtaCod & ".doc" 'abrimos "Mi documento"
            .Visible = True 'hacemos visible Word
        End With
   
End Sub

Public Sub ImpreCartillaAhoCorrienteOP(MatTitular As Variant, ByVal psCtaCod As String, ByVal pnTasa As Double, ByVal pnMonto As Double, ByVal pnPersoneria As Integer, Optional nCostoMan As Currency = 0)

    Dim rs As ADODB.Recordset
    Dim nValor As Double
    Dim nMonMinMN As String
    Dim nMonMinME As String
    Dim nMonMinCH As String
    Dim nMonxConsul As String
    Dim nCosInac As Double
    'Dim nCosTal As Double
    Dim nCosTal As String
    'By Capi 15042008
    Dim nCostalD As Double
    '
    Dim nMonReMN As String
    'By Capi 15042008
    Dim nMonReME As String
    Dim nMonReMNMin As String
    Dim nMonReMNMax As String
    Dim nTasaITF As Double
    Dim nMonOtraPlaza As String
    Dim nPlazo As Integer

    Dim nParCod  As Long
    Dim lsFechas As String
    Dim lsNom1 As String, lsNom2 As String, lsNom3 As String, lsNom4 As String
    Dim lsDoc1 As String, lsDoc2 As String, lsDoc3 As String, lsDoc4 As String
    Dim lsDir1 As String, lsDir2 As String, lsDir3 As String, lsDir4 As String
    Dim lsCad As String
    Dim nitf As Double
    Dim nCosEnvioFisico As String 'JUEZ 20130520
    Dim sTipoEnvio As String 'JUEZ 20130520
    'APRI20190109 ERS077-2018
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim nSaldoEquilibrio As Currency
    nSaldoEquilibrio = clsMant.ObtenerSaldoEquilibrio(nCostoMan, pnTasa)
    'END APRI
    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range

    Dim loRs As COMDConstSistema.DCOMGeneral
    Set loRs = New COMDConstSistema.DCOMGeneral

    Dim loAge As COMDConstantes.DCOMAgencias
    Dim rs1 As ADODB.Recordset
    Dim lsAgencia As String
    Dim lsAgeDir As String
    Set loAge = New COMDConstantes.DCOMAgencias
    Set rs1 = New ADODB.Recordset
        Set rs1 = loAge.RecuperaAgencias(gsCodAge)
        If Not (rs1.EOF And rs1.BOF) Then
            lsAgencia = Trim(rs1("cUbiGeoDescripcion"))
            lsAgeDir = rs1("cAgeDireccion")
        End If
    Set loAge = Nothing

    'JUEZ 20141201 Nuevos Parámetros *******************
    Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
    Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
    Dim rsPar As ADODB.Recordset
    Set rsPar = clsDef.GetCapParametroNew(gCapAhorros, 0)
    'END JUEZ ******************************************

    '----Plazo de la Cuenta de Ahorros con OP---
    nPlazo = 0
    '-------------------------------------------

    '--------Monto Minimos en SOLES y DOLARES --

    If Mid(psCtaCod, 9, 1) = 1 Or Mid(psCtaCod, 9, 1) = 2 Then
       If pnPersoneria <> 1 Then
         nParCod = gMinApeAhOPPJMN
       Else
         nParCod = gMinApeAhOPPNMN
       End If
       'nValor = loRs.GetParametro(2000, nParCod)
       nValor = 0 'rsPar!nSaldoMinCtaSol 'JUEZ 20141201 'APRI20190109 ERS077-2018
       '''nMonMinMN = "S/. " & Format$(nValor, "0.00") 'marg ers044-2016
       nMonMinMN = gcPEN_SIMBOLO & " " & Format$(nValor, "0.00") 'marg ers044-2016

      If pnPersoneria <> 1 Then
         nParCod = gMinApeAhOPPJME
       Else
         nParCod = gMinApeAhOPPNME
       End If
       'nValor = loRs.GetParametro(2000, nParCod)
        nValor = 0 'rsPar!nSaldoMinCtaDol 'JUEZ 20141201 'APRI20190109 ERS077-2018
       nMonMinME = "$. " & Format$(nValor, "0.00")
    End If

    '------------------------------------------------
    'ALPA 20091118************************************************
    Dim objCaptac As COMNCaptaGenerales.NCOMCaptaGenerales
    Set objCaptac = New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim nTREA As Currency
    nTREA = objCaptac.ObtenerTREA(Mid$(psCtaCod, 6, 3), IIf(pnMonto = 0, 10, pnMonto), pnTasa, , nCostoMan)
    '*************************************************************
    '-----Saldo Minimo en Cuenta Ahorros OP --
    Set rs = New ADODB.Recordset
    nParCod = 1003
    'nValor = loRs.GetParametro(2000, nParCod)
    nValor = 0 'rsPar!nSaldoMinCtaSol 'JUEZ 20141201 'APRI20190109 ERS077-2018
    '''nMonMinCH = IIf(Mid(psCtaCod, 9, 1) = "1", nMonMinMN, nMonMinME) & IIf(Mid(psCtaCod, 9, 1) = 1, " (Nuevos Soles)", " (DOLARES)") 'JUEZ 20150121 'marg ers044-2016
    nMonMinCH = IIf(Mid(psCtaCod, 9, 1) = "1", nMonMinMN, nMonMinME) & IIf(Mid(psCtaCod, 9, 1) = 1, " (" & StrConv(gcPEN_PLURAL, vbProperCase) & ")", " (DOLARES)") 'JUEZ 20150121 'marg ers044-2016

    '------------------------------------------------

    '-----Monto de Consulta por extracto de Cuenta --
    'By Capi 05032008
'    Set rs = New ADODB.Recordset
'    If Mid(psCtaCod, 9, 1) = 1 Then
'       nParCod = gDctoExtMNxPag
'    Else
'       nParCod = gDctoExtMExPag
'    End If
'    nValor = loRs.GetParametro(2000, nParCod)
'    nMonxConsul = IIf(Mid(psCtaCod, 9, 1) = 1, "S/. ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(psCtaCod, 9, 1) = 1, " (" & UnNumero(nValor) & " Nuevos Soles)", " (" & UnNumero(nValor) & " Dolares)")

    '------------------------------------------------
     '-----Comision Consulta Saldo Ventanilla --
    Set rs = New ADODB.Recordset
    nParCod = 2106
    nValor = loRs.GetParametro(2000, nParCod)
    '''nMonxConsul = "S/." & Format$(nValor, "0.00") & " (" & UnNumero(nValor) & " Nuevos Soles)" 'marg ers044-2016
    nMonxConsul = gcPEN_SIMBOLO & Format$(nValor, "0.00") & " (" & UnNumero(nValor) & " " & StrConv(gcPEN_PLURAL, vbProperCase) & ")" 'marg ers044-2016


    '--------- Costo de Inactivas -------------------
    Set rs = New ADODB.Recordset
    nValor = loRs.GetParametro(2000, gMonDescInacME)
    nCosInac = Format$(nValor, "0.00")

    '------------------------------------------------

    '--Costo de Talonario --
    Set rs = New ADODB.Recordset
'COMENTADO POR APRI20190109 ERS077-2018
'    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
'    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales

    'By Capi 15042008
    'Set rs = clsMant.GetTarifaOrdenPago(Mid(psCtaCod, 9, 1))
    Set rs = clsMant.GetTarifaOrdenPago(1)
    '
    If Not (rs.EOF And rs.BOF) Then
        nValor = rs!nCosto
        '''nCosTal = "S/. " & Format$(nValor, "0.00") 'marg ers044-2016
        nCosTal = gcPEN_SIMBOLO & " " & Format$(nValor, "0.00") 'marg ers044-2016
    End If
    ''
    Set rs = clsMant.GetTarifaOrdenPago(2)
    If Not (rs.EOF And rs.BOF) Then
        nValor = rs!nCosto
        nCostalD = Format$(nValor, "0.00")
    End If
    '
    Set clsMant = Nothing
    '--------------------------------------------------------------

     '--Retiros/Depositos en Agencias de La Caja en Otras Plazas --
    Set rs = New ADODB.Recordset
    If Mid(psCtaCod, 9, 1) = 1 Then
       nParCod = 2046
    Else
       nParCod = 2047
    End If
    nValor = loRs.GetParametro(2000, nParCod)
    '''nMonOtraPlaza = IIf(Mid(psCtaCod, 9, 1) = 1, "S/. ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(psCtaCod, 9, 1) = 1, " (" & UnNumero(nValor) & " Nuevos Soles)", " (" & UnNumero(nValor) & " Dolares)") 'marg ers044-2016
    nMonOtraPlaza = IIf(Mid(psCtaCod, 9, 1) = 1, gcPEN_SIMBOLO & " ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(psCtaCod, 9, 1) = 1, " (" & UnNumero(nValor) & " " & StrConv(gcPEN_PLURAL, vbProperCase) & ")", " (" & UnNumero(nValor) & " Dolares)") 'marg ers044-2016

    '--------------------------------------------------------------

    '------------ Monto de Rechazo por Odern de Pago --------------
    Set rs = New ADODB.Recordset
    'By Capi 15042008
    'nParCod = 2001
    nParCod = 2004
    nValor = loRs.GetParametro(2000, nParCod)
    nMonReMN = Format$(nValor, "0.00")
    'By Capi 15042008
    nParCod = 2003
    nValor = loRs.GetParametro(2000, nParCod)
    nMonReME = Format$(nValor, "0.00")
    '
    'By Capi 15052008
    'nParCod = 2002
    nParCod = 2107
    nValor = loRs.GetParametro(2000, nParCod)
    nMonReMNMin = Format$(nValor, "0.00")
    nParCod = 1004
    nValor = loRs.GetParametro(2000, nParCod)
    nMonReMNMax = Format$(nValor, "0.00")

    '--------------------------------------------------------------
    'JUEZ 20130520 ***************************************
    Set rs = New ADODB.Recordset
    nValor = loRs.GetParametro(2000, 1005)
    '''nCosEnvioFisico = "S/. " & Format$(nValor, "0.00") 'marg ers044-2016
    nCosEnvioFisico = gcPEN_SIMBOLO & " " & Format$(nValor, "0.00") 'marg ers044-2016

    Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
    Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rs = oCap.RecuperaDatosEnvioEstadoCta(psCtaCod)
    'sTipoEnvio = rs!cTipoEnvio
    'APRI2018 ERS036-2017
       If Not (rs.BOF And rs.EOF) Then
          If rs.RecordCount > 0 Then
            sTipoEnvio = oCap.MostarTextoComisionEnvioEstAhorros(rs!nModoEnvio, 0)
          End If
       Else
            sTipoEnvio = ""
       End If
       'END APRI
    Set rs = Nothing
    'END JUEZ ********************************************

    Set oWord = CreateObject("Word.Application")
    oWord.Visible = True

    Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\CARTILLAAHORROCOP.doc")

    'JUEZ 20150121 ************************************************
    'Cuenta
    lsCad = psCtaCod
    With oWord.Selection.Find
        .Text = "<<cCodCta>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'Tasa
    lsCad = Format(pnTasa, "0.00") & "% (Fija)"
    With oWord.Selection.Find
        .Text = "<<Tasa>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END JUEZ *****************************************************
    'Monto
    'lsCad = Format(pnTasa, "0.00") & "% (Fija)"
    lsCad = IIf(Mid(psCtaCod, 9, 1) = 1, "MN ", "ME ") & Format(pnMonto, "#,##0.00")
    With oWord.Selection.Find
        .Text = "<<Monto>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Plazo
    With oWord.Selection.Find
        .Text = "<<Plazo>>"
        .Replacement.Text = nPlazo
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'ALPA 20091118****************************************
    lsCad = ""
    lsCad = Format(nTREA, "0.00") & "% (Fija) "
    With oWord.Selection.Find
        .Text = "<<TasaTrea>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    '*****************************************************
    'APRI20190109 ERS077-2018
    lsCad = ""
    lsCad = Format(nSaldoEquilibrio, "0.00")
    With oWord.Selection.Find
        .Text = "<<SaldoEquilibrio>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END APRI
    'Monto Minimo en Soles de Apertura
    With oWord.Selection.Find
        .Text = "<<MonMinMN>>"
        .Replacement.Text = nMonMinMN
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Monto Minimo en Dolares de Apertura
    With oWord.Selection.Find
        .Text = "<<MonMinME>>"
        .Replacement.Text = nMonMinME
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Saldo Minimo en Cuenta de Ahorros
    With oWord.Selection.Find
        '.Text = "<<MonMinCH>>"
        .Text = "<<MontoMin>>" 'JUEZ 20150121
        .Replacement.Text = nMonMinCH
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Monto x Consulta
    With oWord.Selection.Find
        .Text = "<<MonConsul>>"
        .Replacement.Text = nMonxConsul
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Costo de Inactivas
    With oWord.Selection.Find
        .Text = "<<CostoInac>>"
        .Replacement.Text = nCosInac
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Monto de Otra Plaza
    With oWord.Selection.Find
        .Text = "<<MonOtraPlaza>>"
        .Replacement.Text = nMonOtraPlaza
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Costo de Talonario
    With oWord.Selection.Find
        .Text = "<<CostoTal>>"
        .Replacement.Text = nCosTal
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'By Capi 15042008
    With oWord.Selection.Find
        .Text = "<<CostoTalD>>"
        .Replacement.Text = nCostalD
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    '

    'Monto Rechazo x Orden de Pago MN
    With oWord.Selection.Find
        .Text = "<<MonReMN>>"
        .Replacement.Text = nMonReMN
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'By Capi 15042008
    'Monto Rechazo x Orden de Pago ME
    With oWord.Selection.Find
        .Text = "<<MonReME>>"
        .Replacement.Text = nMonReME
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    '

    'Monto Rechazo x Orden de Pago MN Min
    With oWord.Selection.Find
        'By Capi 15052008
        '.Text = "<<MonReMEMin>>"
        .Text = "<<MonReMNMin>>"
        .Replacement.Text = nMonReMNMin
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Monto Rechazo x Orden de Pago MN Min
    With oWord.Selection.Find
        'By Capi 15052008
        '.Text = "<<MonReMEMax>>"
        .Text = "<<MonReMNMax>>"
        .Replacement.Text = nMonReMNMax
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'ITF
    nitf = gnITFPorcent * 100
    With oWord.Selection.Find
        .Text = "<<TasaITF>>"
        '.Replacement.Text = Format$(nitf, "0.00")
        .Replacement.Text = Trim(CStr(nitf))
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Fecha
    lsFechas = Format(gdFecSis, "dd") & " de " & Format(gdFecSis, "mmmm") & " del " & Format(gdFecSis, "yyyy")
    With oWord.Selection.Find
        .Text = "<<FecActual>>"
        .Replacement.Text = lsFechas
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Lugar
    With oWord.Selection.Find
        .Text = "<<cLugar>>"
        .Replacement.Text = lsAgencia
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Direccion
    With oWord.Selection.Find
        .Text = "<<cDireccion>>"
        .Replacement.Text = lsAgeDir
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'JUEZ 20130520 ****************************
    'Tipo envio estado de cuenta
    With oWord.Selection.Find
        .Text = "<<TipoEnvio>>"
        .Replacement.Text = sTipoEnvio
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'Costo envio fisico
    With oWord.Selection.Find
        .Text = "<<CostoEnvioFisico>>"
        .Replacement.Text = nCosEnvioFisico
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END JUEZ *********************************

    Dim i As Integer
    Dim rsDir As ADODB.Recordset 'JUEZ 20130520
    Dim oPers As COMDpersona.DCOMPersona 'JUEZ 20130520

    If UBound(MatTitular) = 2 Then

        lsNom1 = "Nombre del Cliente: " & MatTitular(1, 1)
        lsDoc1 = "DNI/RUC: " & MatTitular(1, 2) & Space(60) & "Firma:______________________"
        lsDir1 = "Dirección: " & MatTitular(1, 3)
        'JUEZ 20130520 ****************************
        Set oPers = New COMDpersona.DCOMPersona
        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(1, 2))
        lsDir1 = "Dirección: " & rsDir!cPersDireccDomicilio
        'END JUEZ *********************************

        With oWord.Selection.Find
            .Text = "<<NomTit1>>"
            .Replacement.Text = lsNom1
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit1>>"
            .Replacement.Text = lsDoc1
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit1>>"
            .Replacement.Text = lsDir1
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit2>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit2>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit2>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit3>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit3>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit3>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With

    ElseIf UBound(MatTitular) = 3 Then
        lsNom1 = "Nombre del Cliente: " & MatTitular(1, 1)
        lsDoc1 = "DNI/RUC: " & MatTitular(1, 2) & Space(60) & "Firma:______________________"
        lsDir1 = "Dirección: " & MatTitular(1, 3)
        'JUEZ 20130520 ****************************
        Set oPers = New COMDpersona.DCOMPersona
        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(1, 2))
        lsDir1 = "Dirección: " & rsDir!cPersDireccDomicilio
        'END JUEZ *********************************
        lsNom2 = "Nombre del Cliente: " & MatTitular(2, 1)
        lsDoc2 = "DNI/RUC: " & MatTitular(2, 2) & Space(60) & "Firma:______________________"
        lsDir2 = "Dirección: " & MatTitular(2, 3)
        'JUEZ 20130520 ****************************
        Set oPers = New COMDpersona.DCOMPersona
        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(2, 2))
        lsDir2 = "Dirección: " & rsDir!cPersDireccDomicilio
        'END JUEZ *********************************

        With oWord.Selection.Find
            .Text = "<<NomTit1>>"
            .Replacement.Text = lsNom1
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit1>>"
            .Replacement.Text = lsDoc1
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit1>>"
            .Replacement.Text = lsDir1
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit2>>"
            .Replacement.Text = lsNom2
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit2>>"
            .Replacement.Text = lsDoc2
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit2>>"
            .Replacement.Text = lsDir2
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit3>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit3>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit3>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
    ElseIf UBound(MatTitular) = 4 Then
        lsNom1 = "Nombre del Cliente: " & MatTitular(1, 1)
        lsDoc1 = "DNI/RUC: " & MatTitular(1, 2) & Space(60) & "Firma:______________________"
        lsDir1 = "Dirección: " & MatTitular(1, 3)
        'JUEZ 20130520 ****************************
        Set oPers = New COMDpersona.DCOMPersona
        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(1, 2))
        lsDir1 = "Dirección: " & rsDir!cPersDireccDomicilio
        'END JUEZ *********************************
        lsNom2 = "Nombre del Cliente: " & MatTitular(2, 1)
        lsDoc2 = "DNI/RUC: " & MatTitular(2, 2) & Space(60) & "Firma:______________________"
        lsDir2 = "Dirección: " & MatTitular(2, 3)
        'JUEZ 20130520 ****************************
        Set oPers = New COMDpersona.DCOMPersona
        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(2, 2))
        lsDir2 = "Dirección: " & rsDir!cPersDireccDomicilio
        'END JUEZ *********************************
        lsNom3 = "Nombre del Cliente: " & MatTitular(3, 1)
        lsDoc3 = "DNI/RUC: " & MatTitular(3, 2) & Space(60) & "Firma:______________________"
        lsDir3 = "Dirección: " & MatTitular(3, 3)
        'JUEZ 20130520 ****************************
        Set oPers = New COMDpersona.DCOMPersona
        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(3, 2))
        lsDir3 = "Dirección: " & rsDir!cPersDireccDomicilio
        'END JUEZ *********************************

        With oWord.Selection.Find
            .Text = "<<NomTit1>>"
            .Replacement.Text = lsNom1
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit1>>"
            .Replacement.Text = lsDoc1
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit1>>"
            .Replacement.Text = lsDir1
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit2>>"
            .Replacement.Text = lsNom2
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit2>>"
            .Replacement.Text = lsDoc2
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit2>>"
            .Replacement.Text = lsDir2
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit3>>"
            .Replacement.Text = lsNom3
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit3>>"
            .Replacement.Text = lsDoc3
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit3>>"
            .Replacement.Text = lsDir3
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With

    ElseIf UBound(MatTitular) = 5 Then
        lsNom1 = "Nombre del Cliente: " & MatTitular(1, 1)
        lsDoc1 = "DNI/RUC: " & MatTitular(1, 2) & Space(60) & "Firma:______________________"
        lsDir1 = "Dirección: " & MatTitular(1, 3)
        'JUEZ 20130520 ****************************
        Set oPers = New COMDpersona.DCOMPersona
        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(1, 2))
        lsDir1 = "Dirección: " & rsDir!cPersDireccDomicilio
        'END JUEZ *********************************
        lsNom2 = "Nombre del Cliente: " & MatTitular(2, 1)
        lsDoc2 = "DNI/RUC: " & MatTitular(2, 2) & Space(60) & "Firma:______________________"
        lsDir2 = "Dirección: " & MatTitular(2, 3)
        'JUEZ 20130520 ****************************
        Set oPers = New COMDpersona.DCOMPersona
        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(2, 2))
        lsDir2 = "Dirección: " & rsDir!cPersDireccDomicilio
        'END JUEZ *********************************
        lsNom3 = "Nombre del Cliente: " & MatTitular(3, 1)
        lsDoc3 = "DNI/RUC: " & MatTitular(3, 2) & Space(60) & "Firma:______________________"
        lsDir3 = "Dirección: " & MatTitular(3, 3)
        'JUEZ 20130520 ****************************
        Set oPers = New COMDpersona.DCOMPersona
        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(3, 2))
        lsDir3 = "Dirección: " & rsDir!cPersDireccDomicilio
        'END JUEZ *********************************
        lsNom4 = "Nombre del Cliente: " & MatTitular(4, 1)
        lsDoc4 = "DNI/RUC: " & MatTitular(4, 2) & Space(60) & "Firma:______________________"
        lsDir4 = "Dirección: " & MatTitular(4, 3)
        'JUEZ 20130520 ****************************
        Set oPers = New COMDpersona.DCOMPersona
        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(4, 2))
        lsDir4 = "Dirección: " & rsDir!cPersDireccDomicilio
        'END JUEZ *********************************

        With oWord.Selection.Find
            .Text = "<<NomTit1>>"
            .Replacement.Text = lsNom1
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit1>>"
            .Replacement.Text = lsDoc1
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit1>>"
            .Replacement.Text = lsDir1
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit2>>"
            .Replacement.Text = lsNom2
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit2>>"
            .Replacement.Text = lsDoc2
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit2>>"
            .Replacement.Text = lsDir2
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit3>>"
            .Replacement.Text = lsNom3
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit3>>"
            .Replacement.Text = lsDoc3
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit3>>"
            .Replacement.Text = lsDir3
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit4>>"
            .Replacement.Text = lsNom4
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit4>>"
            .Replacement.Text = lsDoc4
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit4>>"
            .Replacement.Text = lsDir4
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
    End If



   oDoc.SaveAs App.Path & "\SPOOLER\" & psCtaCod & ".doc"
End Sub

Public Sub ImpreCartillaAhoPandero(MatTitular As Variant, ByVal psCtaCod As String, ByVal pnTasa As Double, ByVal pnMonto As Double, ByVal pdFechaA As Date, ByVal pnMontoDep As Double, ByVal pnPlazo As Integer, ByVal pnTpoPrograma As Integer, ByVal psInstitucion As String, Optional nCostoMan As Currency = 0)

    Dim rs As ADODB.Recordset
    Dim nValor As Double
    Dim nMonxConsul As String
    Dim nMonOtraPlaza As String
    Dim nMonMinDep As String
    Dim nTasaITF As Double
    Dim nParCod  As Long
    Dim lsFechas As String
    Dim lsNom1 As String, lsNom2 As String, lsNom3 As String, lsNom4 As String
    Dim lsDoc1 As String, lsDoc2 As String, lsDoc3 As String, lsDoc4 As String
    Dim lsDir1 As String, lsDir2 As String, lsDir3 As String, lsDir4 As String
    Dim lsCad As String
    Dim nitf As Double
    Dim sProducto As String
    Dim nCosEnvioFisico As String 'JUEZ 20130520
    Dim sTipoEnvio As String 'JUEZ 20130520
    'INICIO EAAS20180530 MODIFICACION DE CARTILLAS
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim lnPlazo As Long
    lnPlazo = CLng(pnPlazo)
    Dim lnTasa As Double
    lnTasa = clsMant.GetTasaNominal(pnTasa, 360)
    Dim nInteres As Double
    nInteres = Format$(clsMant.GetInteresPF(lnTasa, pnMonto, lnPlazo), "#,##0.00")
    'FIN EAAS20180530 MODIFICACION DE CARTILLAS
    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range

    Dim loRs As COMDConstSistema.DCOMGeneral
    Set loRs = New COMDConstSistema.DCOMGeneral

    Dim loAge As COMDConstantes.DCOMAgencias
    Dim rs1 As ADODB.Recordset
    Dim lsAgencia As String
    Dim lsAgeDir As String

    Set loAge = New COMDConstantes.DCOMAgencias
    Set rs1 = New ADODB.Recordset
        Set rs1 = loAge.RecuperaAgencias(gsCodAge)
        If Not (rs1.EOF And rs1.BOF) Then
            lsAgeDir = rs1("cAgeDireccion")
            lsAgencia = Trim(rs1("cUbiGeoDescripcion"))
        End If
    Set loAge = Nothing

    'JUEZ 20150121 Nuevos Parámetros *******************
    Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
    Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
    Dim rsPar As ADODB.Recordset
    Set rsPar = clsDef.GetCapParametroNew(gCapPlazoFijo, pnTpoPrograma)
    'END JUEZ ******************************************

    If pnTpoPrograma = 3 Then
        '*** PEAC 20090722
       'sProducto = "PANDERO"
       sProducto = "POCO A POCO AHORRO"
    Else
       sProducto = "DESTINO"
    End If


    '-----Monto de Consulta por extracto de Cuenta --
    'By Capi 05032008

'    Set rs = New ADODB.Recordset
'    If Mid(psCtaCod, 9, 1) = 1 Then
'       nParCod = gDctoExtMNxPag
'    Else
'       nParCod = gDctoExtMExPag
'    End If
'    nValor = loRs.GetParametro(2000, nParCod)
'    nMonxConsul = IIf(Mid(psCtaCod, 9, 1) = 1, "S/. ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(psCtaCod, 9, 1) = 1, " (" & UnNumero(nValor) & " Nuevos Soles)", " (" & UnNumero(nValor) & " Dolares)")

    '------------------------------------------------
     '-----Comision Consulta Saldo Ventanilla --
    Set rs = New ADODB.Recordset
    nParCod = 2106
    nValor = loRs.GetParametro(2000, nParCod)
    '''nMonxConsul = "S/." & Format$(nValor, "0.00") & " (" & UnNumero(nValor) & " Nuevos Soles)" 'marg ers044-2016
    nMonxConsul = gcPEN_SIMBOLO & Format$(nValor, "0.00") & " (" & UnNumero(nValor) & " " & StrConv(gcPEN_PLURAL, vbProperCase) & ")" 'marg ers044-2016


    '--Retiros/Depositos en Agencias de La Caja en Otras Plazas --
    Set rs = New ADODB.Recordset
    If Mid(psCtaCod, 9, 1) = 1 Then
       nParCod = 2046
    Else
       nParCod = 2047
    End If
    nValor = loRs.GetParametro(2000, nParCod)
    '''nMonOtraPlaza = IIf(Mid(psCtaCod, 9, 1) = 1, "S/. ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(psCtaCod, 9, 1) = 1, " (" & UnNumero(nValor) & " Nuevos Soles)", " (" & UnNumero(nValor) & " Dolares)") 'marg ers044-2016
    nMonOtraPlaza = IIf(Mid(psCtaCod, 9, 1) = 1, gcPEN_SIMBOLO & " ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(psCtaCod, 9, 1) = 1, " (" & UnNumero(nValor) & " " & StrConv(gcPEN_PLURAL, vbProperCase) & ")", " (" & UnNumero(nValor) & " Dolares)") 'marg ers044-2016

    '--------------------------------------------------------------
     '--------------------------------------------------------------
    'By capi 28022008
    '--------- Monto Minimo a Depositar Mensual -------------------
    Set rs = New ADODB.Recordset
    nParCod = 2093
    nValor = loRs.GetParametro(2000, nParCod)
    nValor = IIf(Mid(psCtaCod, 9, 1) = 1, rsPar!nAumCapMinSol, rsPar!nAumCapMinDol) 'JUEZ 20150121
    'JUEZ 20150121 **********************
    If pnTpoPrograma = 3 Then
        Dim clsCapGen As New COMNCaptaGenerales.NCOMCaptaGenerales
        nValor = clsCapGen.GetDatosCuenta(psCtaCod)!nMontoAbono
        Set clsCapGen = Nothing
    End If
    'END JUEZ ***************************
    nMonMinDep = Format$(nValor, "0.00")


    '--------------------------------------------------------------
    'JUEZ 20130520 ***************************************
    Set rs = New ADODB.Recordset
    nValor = loRs.GetParametro(2000, 1005)
    '''nCosEnvioFisico = "S/. " & Format$(nValor, "0.00") 'marg ers044-2016
    nCosEnvioFisico = gcPEN_SIMBOLO & " " & Format$(nValor, "0.00") 'marg ers044-2016

    Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
    Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rs = oCap.RecuperaDatosEnvioEstadoCta(psCtaCod)
    If Not rs.EOF Then sTipoEnvio = rs!cTipoEnvio 'APRI20180530 ERS036-2017
    'sTipoEnvio = rs!cTipoEnvio
    Set rs = Nothing
    'END JUEZ ********************************************
     'ALPA 20091118************************************************
    Dim objCaptac As COMNCaptaGenerales.NCOMCaptaGenerales
    Set objCaptac = New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim pnPlazoT As Integer
    Dim nTREA As Currency
    pnPlazoT = CInt(pnPlazo)
    nTREA = objCaptac.ObtenerTREA(Mid$(psCtaCod, 6, 3), IIf(pnMonto = 0, 10, pnMonto), pnTasa, pnPlazoT, nCostoMan)
    '*************************************************************
    Set oWord = CreateObject("Word.Application")
    oWord.Visible = True

    If pnTpoPrograma = 3 Then
        '*** PEAC 20090731
       'Set oDoc = oWord.Documents.Open(App.path & "\FormatoCarta\CARTILLAAHORRPandero.doc")
       Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\CARTILLAPOCOAPOCOAHORRO.doc")
    Else
       Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\CARTILLAAHORRDestino.doc")
    End If

    'Tipo de producto
    With oWord.Selection.Find
        .Text = "<<cProducto>>"
        .Replacement.Text = sProducto
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With



    'El Monto en letras
    With oWord.Selection.Find
        .Text = "<<cMontos>>"
        'By Capi 01042008 para que se visualice correctamente el importe en letras
        .Replacement.Text = ConversNL(Mid(psCtaCod, 9, 1), pnMonto)
        '.Replacement.Text = UnNumero(pnMonto) & IIf(Mid(psCtaCod, 9, 1) = "1", " 00/100 NUEVOS SOLES", "")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Monto
    '''lsCad = IIf(Mid(psCtaCod, 9, 1) = 1, "S/. ", "$. ") & Format(pnMonto, "0.00") 'marg ers044-2016
    lsCad = IIf(Mid(psCtaCod, 9, 1) = 1, gcPEN_SIMBOLO & " ", "$. ") & Format(pnMonto, "0.00") 'marg ers044-2016
    With oWord.Selection.Find
        .Text = "<<nMonto>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    If pnTpoPrograma = 3 Then
        With oWord.Selection.Find
            .Text = "<<nMontoD>>"
            '''.Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = 1, "S/. ", "$. ") & Format$(pnMontoDep, "0.00") & IIf(Mid(psCtaCod, 9, 1) = "1", "  NUEVOS SOLES", "  DOLARES") 'marg ers044-2016
            .Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = 1, gcPEN_SIMBOLO & " ", "$. ") & Format$(pnMontoDep, "0.00") & IIf(Mid(psCtaCod, 9, 1) = "1", "  " & StrConv(gcPEN_PLURAL, vbUpperCase), "  DOLARES") 'marg ers044-2016
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
    End If

    If pnTpoPrograma = 4 Then
        With oWord.Selection.Find
            .Text = "<<cInstitucion>>"
            .Replacement.Text = psInstitucion
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
    End If
   'ALPA 20091118****************************************
    lsCad = ""
    lsCad = Format(nTREA, "0.00") & " % (Fija)"
    With oWord.Selection.Find
        .Text = "<<TasaTrea>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    '*****************************************************
    'Fecha de Vencimiento
    With oWord.Selection.Find
        .Text = "<<nPlazo>>"
        .Replacement.Text = pnPlazo
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Cuenta
    With oWord.Selection.Find
        .Text = "<<Cuenta>>"
        .Replacement.Text = psCtaCod
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Monto
    lsCad = IIf(Mid(psCtaCod, 9, 1) = 1, "MN ", "ME ") & Format(pnMonto, "#,##0.00")
    With oWord.Selection.Find
        .Text = "<<Monto>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With


    'Monto Minimo Deposito
    '''lsCad = IIf(Mid(psCtaCod, 9, 1) = 1, "S/. ", "$. ") & Format$(nMonMinDep, "0.00") & IIf(Mid(psCtaCod, 9, 1) = 1, " Nuevos Soles", " Dolares") 'marg ers044-2016
    lsCad = IIf(Mid(psCtaCod, 9, 1) = 1, gcPEN_SIMBOLO & " ", "$. ") & Format$(nMonMinDep, "0.00") & IIf(Mid(psCtaCod, 9, 1) = 1, " " & StrConv(gcPEN_PLURAL, vbProperCase), " Dolares") 'marg ers044-2016
    With oWord.Selection.Find
        .Text = "<<MontoMinDep>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Tasa
    lsCad = ""
    lsCad = Format(pnTasa, "0.00") & " % (Fija)"
    With oWord.Selection.Find
        .Text = "<<Tasa>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

     'Fecha de Vencimiento
    Dim dFechaV As Date
    dFechaV = CDate(pdFechaA) + pnPlazo
    With oWord.Selection.Find
        .Text = "<<FechaVen>>"
        .Replacement.Text = dFechaV
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With


    'Monto x Consulta
    With oWord.Selection.Find
        .Text = "<<MonConsul>>"
        .Replacement.Text = nMonxConsul
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    If pnTpoPrograma = 3 Then
        'Monto de Otra Plaza
        With oWord.Selection.Find
            .Text = "<<MonOtraPlaza>>"
            .Replacement.Text = nMonOtraPlaza
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
    End If


    'ITF
    nitf = gnITFPorcent * 100
    With oWord.Selection.Find
        .Text = "<<TasaITF>>"
        '.Replacement.Text = Format$(nitf, "0.00")
        .Replacement.Text = Trim(CStr(nitf))
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Fecha
    lsFechas = Format(gdFecSis, "dd") & " de " & Format(gdFecSis, "mmmm") & " del " & Format(gdFecSis, "yyyy")
    With oWord.Selection.Find
        .Text = "<<FecActual>>"
        .Replacement.Text = lsFechas
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Lugar
    With oWord.Selection.Find
        .Text = "<<cLugar>>"
        .Replacement.Text = lsAgencia
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Direccion
    With oWord.Selection.Find
        .Text = "<<cDireccion>>"
        .Replacement.Text = lsAgeDir
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'JUEZ 20130520 ****************************
    'Tipo envio estado de cuenta
    With oWord.Selection.Find
        .Text = "<<TipoEnvio>>"
        .Replacement.Text = sTipoEnvio
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'Costo envio fisico
    With oWord.Selection.Find
        .Text = "<<CostoEnvioFisico>>"
        .Replacement.Text = nCosEnvioFisico
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END JUEZ *********************************

    'INICIO EAAS20180530 MODIFICACION DE CARTILLAS
    lsCad = ""
    lsCad = IIf(Mid(psCtaCod, 9, 1) = 1, gcPEN_SIMBOLO & " ", "$. ") & Format(nInteres, "#,##0.00")
    With oWord.Selection.Find
        .Text = "<<MonInteres>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'Plazo
    With oWord.Selection.Find
        .Text = "<<Plazo>>"
        .Replacement.Text = lnPlazo
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'FIN EAAS20180530 MODIFICACION DE CARTILLAS

    Dim i As Integer
    Dim rsDir As ADODB.Recordset 'JUEZ 20130520
    Dim oPers As COMDpersona.DCOMPersona 'JUEZ 20130520

    If UBound(MatTitular) = 2 Then

        lsNom1 = "Nombre del Cliente: " & MatTitular(1, 1)
        lsDoc1 = "DNI/RUC: " & MatTitular(1, 2) & Space(60) & "Firma:______________________"
        lsDir1 = "Dirección: " & MatTitular(1, 3)
        'JUEZ 20130520 ****************************
        Set oPers = New COMDpersona.DCOMPersona
        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(1, 2))
        lsDir1 = "Dirección: " & rsDir!cPersDireccDomicilio
        'END JUEZ *********************************

        With oWord.Selection.Find
            .Text = "<<NomTit1>>"
            .Replacement.Text = lsNom1
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit1>>"
            .Replacement.Text = lsDoc1
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit1>>"
            .Replacement.Text = lsDir1
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit2>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit2>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit2>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit3>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit3>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit3>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With

    ElseIf UBound(MatTitular) = 3 Then
        lsNom1 = "Nombre del Cliente: " & MatTitular(1, 1)
        lsDoc1 = "DNI/RUC: " & MatTitular(1, 2) & Space(60) & "Firma:______________________"
        lsDir1 = "Dirección: " & MatTitular(1, 3)
        'JUEZ 20130520 ****************************
        Set oPers = New COMDpersona.DCOMPersona
        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(1, 2))
        lsDir1 = "Dirección: " & rsDir!cPersDireccDomicilio
        'END JUEZ *********************************
        lsNom2 = "Nombre del Cliente: " & MatTitular(2, 1)
        lsDoc2 = "DNI/RUC: " & MatTitular(2, 2) & Space(60) & "Firma:______________________"
        lsDir2 = "Dirección: " & MatTitular(2, 3)
        'JUEZ 20130520 ****************************
        Set oPers = New COMDpersona.DCOMPersona
        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(2, 2))
        lsDir2 = "Dirección: " & rsDir!cPersDireccDomicilio
        'END JUEZ *********************************

        With oWord.Selection.Find
            .Text = "<<NomTit1>>"
            .Replacement.Text = lsNom1
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit1>>"
            .Replacement.Text = lsDoc1
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit1>>"
            .Replacement.Text = lsDir1
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit2>>"
            .Replacement.Text = lsNom2
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit2>>"
            .Replacement.Text = lsDoc2
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit2>>"
            .Replacement.Text = lsDir2
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit3>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit3>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit3>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
    ElseIf UBound(MatTitular) = 4 Then
        lsNom1 = "Nombre del Cliente: " & MatTitular(1, 1)
        lsDoc1 = "DNI/RUC: " & MatTitular(1, 2) & Space(60) & "Firma:______________________"
        lsDir1 = "Dirección: " & MatTitular(1, 3)
        'JUEZ 20130520 ****************************
        Set oPers = New COMDpersona.DCOMPersona
        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(1, 2))
        lsDir1 = "Dirección: " & rsDir!cPersDireccDomicilio
        'END JUEZ *********************************
        lsNom2 = "Nombre del Cliente: " & MatTitular(2, 1)
        lsDoc2 = "DNI/RUC: " & MatTitular(2, 2) & Space(60) & "Firma:______________________"
        lsDir2 = "Dirección: " & MatTitular(2, 3)
        'JUEZ 20130520 ****************************
        Set oPers = New COMDpersona.DCOMPersona
        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(2, 2))
        lsDir2 = "Dirección: " & rsDir!cPersDireccDomicilio
        'END JUEZ *********************************
        lsNom3 = "Nombre del Cliente: " & MatTitular(3, 1)
        lsDoc3 = "DNI/RUC: " & MatTitular(3, 2) & Space(60) & "Firma:______________________"
        lsDir3 = "Dirección: " & MatTitular(3, 3)
        'JUEZ 20130520 ****************************
        Set oPers = New COMDpersona.DCOMPersona
        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(3, 2))
        lsDir3 = "Dirección: " & rsDir!cPersDireccDomicilio
        'END JUEZ *********************************

        With oWord.Selection.Find
            .Text = "<<NomTit1>>"
            .Replacement.Text = lsNom1
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit1>>"
            .Replacement.Text = lsDoc1
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit1>>"
            .Replacement.Text = lsDir1
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit2>>"
            .Replacement.Text = lsNom2
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit2>>"
            .Replacement.Text = lsDoc2
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit2>>"
            .Replacement.Text = lsDir2
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit3>>"
            .Replacement.Text = lsNom3
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit3>>"
            .Replacement.Text = lsDoc3
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit3>>"
            .Replacement.Text = lsDir3
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With

    ElseIf UBound(MatTitular) = 5 Then
        lsNom1 = "Nombre del Cliente: " & MatTitular(1, 1)
        lsDoc1 = "DNI/RUC: " & MatTitular(1, 2) & Space(60) & "Firma:______________________"
        lsDir1 = "Dirección: " & MatTitular(1, 3)
        'JUEZ 20130520 ****************************
        Set oPers = New COMDpersona.DCOMPersona
        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(1, 2))
        lsDir1 = "Dirección: " & rsDir!cPersDireccDomicilio
        'END JUEZ *********************************
        lsNom2 = "Nombre del Cliente: " & MatTitular(2, 1)
        lsDoc2 = "DNI/RUC: " & MatTitular(2, 2) & Space(60) & "Firma:______________________"
        lsDir2 = "Dirección: " & MatTitular(2, 3)
        'JUEZ 20130520 ****************************
        Set oPers = New COMDpersona.DCOMPersona
        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(2, 2))
        lsDir2 = "Dirección: " & rsDir!cPersDireccDomicilio
        'END JUEZ *********************************
        lsNom3 = "Nombre del Cliente: " & MatTitular(3, 1)
        lsDoc3 = "DNI/RUC: " & MatTitular(3, 2) & Space(60) & "Firma:______________________"
        lsDir3 = "Dirección: " & MatTitular(3, 3)
        'JUEZ 20130520 ****************************
        Set oPers = New COMDpersona.DCOMPersona
        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(3, 2))
        lsDir3 = "Dirección: " & rsDir!cPersDireccDomicilio
        'END JUEZ *********************************
        lsNom4 = "Nombre del Cliente: " & MatTitular(4, 1)
        lsDoc4 = "DNI/RUC: " & MatTitular(4, 2) & Space(60) & "Firma:______________________"
        lsDir4 = "Dirección: " & MatTitular(4, 3)
        'JUEZ 20130520 ****************************
        Set oPers = New COMDpersona.DCOMPersona
        Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(4, 2))
        lsDir4 = "Dirección: " & rsDir!cPersDireccDomicilio
        'END JUEZ *********************************

        With oWord.Selection.Find
            .Text = "<<NomTit1>>"
            .Replacement.Text = lsNom1
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit1>>"
            .Replacement.Text = lsDoc1
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit1>>"
            .Replacement.Text = lsDir1
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit2>>"
            .Replacement.Text = lsNom2
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit2>>"
            .Replacement.Text = lsDoc2
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit2>>"
            .Replacement.Text = lsDir2
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit3>>"
            .Replacement.Text = lsNom3
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit3>>"
            .Replacement.Text = lsDoc3
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit3>>"
            .Replacement.Text = lsDir3
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit4>>"
            .Replacement.Text = lsNom4
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit4>>"
            .Replacement.Text = lsDoc4
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit4>>"
            .Replacement.Text = lsDir4
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
    End If



   oDoc.SaveAs App.Path & "\SPOOLER\" & psCtaCod & ".doc"
End Sub

Public Function ImpreCartillaAHPanderoLote(MatTitular() As String, MatNroCta() As String, ByVal pnTpoPrograma As Integer, ByVal psInstitucion As String)

    Dim rs As ADODB.Recordset
    Dim nValor As Double
    Dim nMonxConsul As String
    Dim nMonOtraPlaza As String
    Dim nTasaITF As Double
    Dim nParCod  As Long
    Dim lsFechas As String
    Dim lsNom1 As String
    Dim lsDoc1 As String
    Dim lsDir1 As String
    Dim lsCad As String
    Dim nitf As Double
    Dim lsModeloPlantilla As String
    Dim i As Integer
    Dim sProducto As String
    Dim loRs As COMDConstSistema.DCOMGeneral
    Set loRs = New COMDConstSistema.DCOMGeneral

    Dim loAge As COMDConstantes.DCOMAgencias
    Dim rs1 As ADODB.Recordset
    Dim lsAgencia As String
    Dim lsAgeDir As String
    Set loAge = New COMDConstantes.DCOMAgencias
    Set rs1 = New ADODB.Recordset
        Set rs1 = loAge.RecuperaAgencias(gsCodAge)
        If Not (rs1.EOF And rs1.BOF) Then
            lsAgencia = Trim(rs1("cUbiGeoDescripcion"))
            lsAgeDir = rs1("cAgeDireccion")
        End If
    Set loAge = Nothing

    If pnTpoPrograma = 3 Then
        '*** PEAC 20090722
       'sProducto = "PANDERO"
       sProducto = "POCO A POCO AHORRO"
    Else
       sProducto = "DESTINO"
    End If


    '-----Monto de Consulta por extracto de Cuenta --
    'By capi 05032008

'    Set rs = New ADODB.Recordset
'    If Mid(MatNroCta(1), 9, 1) = 1 Then
'       nParCod = gDctoExtMNxPag
'    Else
'       nParCod = gDctoExtMExPag
'    End If
'    nValor = loRs.GetParametro(2000, nParCod)
'    nMonxConsul = IIf(Mid(MatNroCta(1), 9, 1) = 1, "S/. ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(MatNroCta(1), 9, 1) = 1, " (" & UnNumero(nValor) & " Nuevos Soles)", " (" & UnNumero(nValor) & " Dolares)")

    '------------------------------------------------
     '-----Comision Consulta Saldo Ventanilla --
    Set rs = New ADODB.Recordset
    nParCod = 2106
    nValor = loRs.GetParametro(2000, nParCod)
    '''nMonxConsul = "S/." & Format$(nValor, "0.00") & " (" & UnNumero(nValor) & " Nuevos Soles)" 'marg ers044-2016
    nMonxConsul = gcPEN_SIMBOLO & Format$(nValor, "0.00") & " (" & UnNumero(nValor) & " " & StrConv(gcPEN_PLURAL, vbProperCase) & ")" 'marg ers044-2016

    '--Retiros/Depositos en Agencias de La Caja en Otras Plazas --
    Set rs = New ADODB.Recordset
    If Mid(MatNroCta(1), 9, 1) = 1 Then
       nParCod = 2046
    Else
       nParCod = 2047
    End If
    nValor = loRs.GetParametro(2000, nParCod)
    '''nMonOtraPlaza = IIf(Mid(MatNroCta(1), 9, 1) = 1, "S/. ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(MatNroCta(1), 9, 1) = 1, " (" & UnNumero(nValor) & " Nuevos Soles)", " (" & UnNumero(nValor) & " Dolares)") 'marg ers044-2016
    nMonOtraPlaza = IIf(Mid(MatNroCta(1), 9, 1) = 1, gcPEN_SIMBOLO & " ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(MatNroCta(1), 9, 1) = 1, " (" & UnNumero(nValor) & " " & StrConv(gcPEN_PLURAL, vbProperCase) & ")", " (" & UnNumero(nValor) & " Dolares)") 'marg ers044-2016

    '--------------------------------------------------------------
    If pnTpoPrograma = 3 Then
        '*** PEAC 20090731
        'lsModeloPlantilla = App.path & "\FormatoCarta\CARTILLAAHORRPandero.doc"
        lsModeloPlantilla = App.Path & "\FormatoCarta\CARTILLAPOCOAPOCOAHORROL.doc"
    Else
        lsModeloPlantilla = App.Path & "\FormatoCarta\CARTILLAAHORRDestino.doc"
    End If

    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range

    Set oWord = CreateObject("Word.Application")
    oWord.Visible = True
    'Crea una clase que de Word Object
    Dim wApp As Word.Application
    Dim wAppSource As Word.Application
    'Create a new instance of word
    Set wApp = New Word.Application
    Set wAppSource = New Word.Application

    Dim RangeSource As Word.Range
    'Abre Documento Plantilla
    wAppSource.Documents.Open FileName:=lsModeloPlantilla
    Set RangeSource = wAppSource.ActiveDocument.Content
    'Lo carga en Memoria
    wAppSource.ActiveDocument.Content.Copy

    'Crea Nuevo Documento
    wApp.Documents.Add


   For i = 1 To UBound(MatTitular)

        wApp.Application.Selection.TypeParagraph
        wApp.Application.Selection.Paste
        wApp.Application.Selection.InsertBreak
        wApp.Selection.SetRange start:=wApp.Selection.start, End:=wApp.ActiveDocument.Content.End
        wApp.Selection.MoveEnd


    'Tipo de producto
    With oWord.Selection.Find
        .Text = "<<cProducto>> "
        .Replacement.Text = sProducto
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'El Monto en letras
    With oWord.Selection.Find
        .Text = "<<cMontos>>"
        '''.Replacement.Text = UnNumero(MatTitular(i, 3)) & IIf(Mid(MatNroCta(i), 9, 1) = "1", "00/100 NUEVOS SOLES", "") 'marg ers044-2016
        .Replacement.Text = UnNumero(MatTitular(i, 3)) & IIf(Mid(MatNroCta(i), 9, 1) = "1", "00/100 " & StrConv(gcPEN_PLURAL, vbUpperCase), "") 'marg ers044-2016
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Monto
    '''lsCad = IIf(Mid(MatNroCta(i), 9, 1) = 1, "S/. ", "$. ") & Format(MatTitular(i, 3), "0.00") 'marg ers044-2016
    lsCad = IIf(Mid(MatNroCta(i), 9, 1) = 1, gcPEN_SIMBOLO & " ", "$. ") & Format(MatTitular(i, 3), "0.00") 'marg ers044-2016
    With oWord.Selection.Find
        .Text = "<<nMonto>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'Monto
    If pnTpoPrograma = 3 Then
        With oWord.Selection.Find
            .Text = "<<nMontoD>>"
            '''.Replacement.Text = MatTitular(i, 3) & IIf(Mid(MatNroCta(i), 9, 1) = "1", " NUEVOS SOLES", " DOLARES") 'marg ers044-2016
            .Replacement.Text = MatTitular(i, 3) & IIf(Mid(MatNroCta(i), 9, 1) = "1", " " & StrConv(gcPEN_PLURAL, vbUpperCase), " DOLARES") 'marg ers044-2016
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
    End If

     If pnTpoPrograma = 4 Then
        With oWord.Selection.Find
            .Text = "<<cInstitucion>>"
            .Replacement.Text = psInstitucion
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
    End If

    'plazo
    With oWord.Selection.Find
        .Text = "<<nPlazo>>"
        .Replacement.Text = MatTitular(i, 13)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Cuenta
    With wApp.Selection.Find
        .Text = "<<Cuenta>>"
        .Replacement.Text = MatNroCta(i)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Monto
    lsCad = IIf(Mid(MatNroCta(i), 9, 1) = 1, "MN ", "ME ") & Format(MatTitular(i, 3), "0.00")
    With wApp.Selection.Find
        .Text = "<<Monto>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Monto Minimo de Deposito
    '''lsCad = IIf(Mid(MatNroCta(i), 9, 1) = 1, "S/. ", "$. ") & Format$(MatTitular(i, 8), "0.00") & IIf(Mid(MatNroCta(i), 9, 1) = 1, " Nuevos Soles", " Dolares") 'marg ers044-2016
    lsCad = IIf(Mid(MatNroCta(i), 9, 1) = 1, gcPEN_SIMBOLO & " ", "$. ") & Format$(MatTitular(i, 8), "0.00") & IIf(Mid(MatNroCta(i), 9, 1) = 1, " " & StrConv(gcPEN_PLURAL, vbProperCase), " Dolares") 'marg ers044-2016
    With wApp.Selection.Find
        .Text = "<<MontoMinDep>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Tasa
    lsCad = ""
    '''lsCad = Format(MatTitular(i, 2), "0.00") & " % " & IIf(Mid(MatNroCta(1), 9, 1) = 1, "Soles", "Dolares") 'marg ers044-2016
    lsCad = Format(MatTitular(i, 2), "0.00") & " % " & IIf(Mid(MatNroCta(1), 9, 1) = 1, StrConv(gcPEN_PLURAL, vbProperCase), "Dolares") 'marg ers044-2016
    With wApp.Selection.Find
        .Text = "<<Tasa>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With


    'Monto x Consulta
    With wApp.Selection.Find
        .Text = "<<MonConsul>>"
        .Replacement.Text = nMonxConsul
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    If pnTpoPrograma = 3 Then
        'Monto de Otra Plaza
        With wApp.Selection.Find
            .Text = "<<MonOtraPlaza>>"
            .Replacement.Text = nMonOtraPlaza
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
    End If
    'ITF
    nitf = gnITFPorcent * 100
    With wApp.Selection.Find
        .Text = "<<TasaITF>>"
        '.Replacement.Text = Format$(nitf, "0.00")
        .Replacement.Text = Trim(CStr(nitf))
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

     'Fecha de Vencimiento
    Dim dFechaV As Date
    dFechaV = CDate(gdFecSis) + MatTitular(i, 7)
    With wApp.Selection.Find
        .Text = "<<FechaVen>>"
        .Replacement.Text = dFechaV
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Fecha
    lsFechas = Format(gdFecSis, "dd") & " de " & Format(gdFecSis, "mmmm") & " del " & Format(gdFecSis, "yyyy")
    With wApp.Selection.Find
        .Text = "<<FecActual>>"
        .Replacement.Text = lsFechas
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Lugar
    With oWord.Selection.Find
        .Text = "<<cLugar>>"
        .Replacement.Text = lsAgencia
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

     'Direccion
    With wApp.Selection.Find
        .Text = "<<cDireccion>>"
        .Replacement.Text = lsAgeDir
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With


    lsNom1 = "Nombre del Cliente: " & MatTitular(i, 1)
    lsDoc1 = "DNI/RUC: " & MatTitular(i, 4) & Space(60) & "Firma:______________________"
    lsDir1 = "Dirección: " & MatTitular(i, 5)

    With wApp.Selection.Find
        .Text = "<<NomTit1>>"
        .Replacement.Text = lsNom1
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DocTit1>>"
        .Replacement.Text = lsDoc1
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DirTit1>>"
        .Replacement.Text = lsDir1
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<NomTit2>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DocTit2>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DirTit2>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<NomTit3>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DocTit3>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DirTit3>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<NomTit4>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DocTit4>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DirTit4>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

  Next

wAppSource.ActiveDocument.Close
wApp.Visible = True

End Function

Public Function ImpreCartillaPFLote(MatTitular() As String, MatNroCta() As String)

    Dim rs As ADODB.Recordset
    Dim nValor As Double
    Dim lsFechas As String
    Dim lsNom1 As String
    Dim lsDoc1 As String
    Dim lsDir1 As String
    Dim lsCad As String
    Dim nitf As Double
    Dim lsModeloPlantilla As String
    Dim i As Integer
    Dim nCosEnvioFisico As String 'JUEZ 20150121
    Dim sTipoEnvio As String 'JUEZ 20150121

    Dim sTaInEf As String
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
    Dim lnPlazo As Long
    Dim nInteres As Double
    Dim lnFormaRetiro As Long


    Dim loRs As COMDConstSistema.DCOMGeneral
    Set loRs = New COMDConstSistema.DCOMGeneral

    Dim loAge As COMDConstantes.DCOMAgencias
    Dim rs1 As ADODB.Recordset
    Dim lsAgencia As String
    Dim lsAgeDir As String
    Set loAge = New COMDConstantes.DCOMAgencias
    Set rs1 = New ADODB.Recordset
        Set rs1 = loAge.RecuperaAgencias(gsCodAge)
        If Not (rs1.EOF And rs1.BOF) Then
            lsAgencia = Trim(rs1("cUbiGeoDescripcion"))
            lsAgeDir = rs1("cAgeDireccion")
        End If
    Set loAge = Nothing

    lsModeloPlantilla = App.Path & "\FormatoCarta\CARTILLAPFL.doc" 'JATO 20210407
 
    'Crea una clase que de Word Object
    Dim wApp As Word.Application
    Dim wAppSource As Word.Application
    'Create a new instance of word
    Set wApp = New Word.Application
    Set wAppSource = New Word.Application

    Dim RangeSource As Word.Range
    'Abre Documento Plantilla
    wAppSource.Documents.Open FileName:=lsModeloPlantilla
    Set RangeSource = wAppSource.ActiveDocument.Content
    'Lo carga en Memoria
    wAppSource.ActiveDocument.Content.Copy

    'Crea Nuevo Documento
    wApp.Documents.Add


   For i = 1 To UBound(MatTitular)

        wApp.Application.Selection.TypeParagraph
        wApp.Application.Selection.Paste
        wApp.Application.Selection.InsertBreak
        wApp.Selection.SetRange start:=wApp.Selection.start, End:=wApp.ActiveDocument.Content.End
        wApp.Selection.MoveEnd

    lnPlazo = CLng(MatTitular(i, 6))
    lnFormaRetiro = CLng(MatTitular(i, 10))
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    sTaInEf = Format$(((((((MatTitular(i, 2) / 12) / 3000) + 1) ^ lnPlazo) - 1) * 100), "#0.00")
    nInteres = Format$(clsMant.GetInteresPF(MatTitular(i, 2), MatTitular(i, 3), lnPlazo), "#,##0.00")

    'JUEZ 20150121 ************************************************
    Set rs = New ADODB.Recordset
    nValor = loRs.GetParametro(2000, 1005)
    '''nCosEnvioFisico = "S/. " & Format$(nValor, "0.00") 'marg ers044-2016
    nCosEnvioFisico = gcPEN_SIMBOLO & " " & Format$(nValor, "0.00") 'marg ers044-2016

    Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
    Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rs = oCap.RecuperaDatosEnvioEstadoCta(MatNroCta(i))
    If Not rs.EOF Then sTipoEnvio = rs!cTipoEnvio
    Set rs = Nothing

    'Cuenta
    lsCad = MatNroCta(i)
    With wApp.Selection.Find
        .Text = "<<cCodCta>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END JUEZ *****************************************************
    'JUEZ 20150121 ************************************************
    Dim objCaptac As COMNCaptaGenerales.NCOMCaptaGenerales
    Set objCaptac = New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim nTREA As Currency
    nTREA = objCaptac.ObtenerTREA(Mid$(MatNroCta(i), 6, 3), IIf(CCur(MatTitular(i, 3)) = 0, 10, CCur(MatTitular(i, 3))), CCur(sTaInEf), CInt(lnPlazo))
    'END JUEZ *****************************************************
   'Monto
   lsCad = IIf(Mid(MatNroCta(i), 9, 1) = 1, "MN ", "ME ") & Format(MatTitular(i, 3), "0.00")
    With wApp.Selection.Find
        .Text = "<<Monto>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Plazo
    With wApp.Selection.Find
        .Text = "<<Plazo>>"
        .Replacement.Text = lnPlazo
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Tasa
    lsCad = ""
    lsCad = sTaInEf & " %"
    With wApp.Selection.Find
        .Text = "<<Tasa>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'JUEZ 20150121 ***********************
    lsCad = ""
    lsCad = nTREA & " % (Fija)"
    With wApp.Selection.Find
        .Text = "<<TasaTrea>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END JUEZ ****************************
    'Monto de interes
    lsCad = ""
    '''lsCad = IIf(Mid(MatNroCta(i), 9, 1) = 1, "S/. ", "$. ") & nInteres 'marg ers044-2016
    lsCad = IIf(Mid(MatNroCta(i), 9, 1) = 1, gcPEN_SIMBOLO & " ", "$. ") & nInteres 'marg ers044-2016
    With wApp.Selection.Find
        .Text = "<<MonInteres>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Fecha Apertura
    With wApp.Selection.Find
        .Text = "<<FechaAper>>"
        .Replacement.Text = gdFecSis
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With 'ITF

    nitf = gnITFPorcent * 100
    With wApp.Selection.Find
        .Text = "<<TasaITF>>"
        '.Replacement.Text = Format$(nitf, "0.00")
        .Replacement.Text = Trim(CStr(nitf))
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

     'Fecha de Vencimiento
    Dim dFechaV As Date
    dFechaV = CDate(gdFecSis) + lnPlazo
    With wApp.Selection.Find
        .Text = "<<FechaVen>>"
        .Replacement.Text = dFechaV
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'JUEZ 20150121 ************************************************
    'Forma de Retiro
    lsCad = loRs.GetConstante(gCaptacPFFormaRetiro, , CStr(lnFormaRetiro), 1)!CDescripcion 'JUEZ 20150121
    With wApp.Selection.Find
        .Text = "<<cFormaRetiro>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END JUEZ *****************************************************
    'Fecha
    lsFechas = Format(gdFecSis, "dd") & " de " & Format(gdFecSis, "mmmm") & " del " & Format(gdFecSis, "yyyy")
    With wApp.Selection.Find
        .Text = "<<FecActual>>"
        .Replacement.Text = lsFechas
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

     'Lugar
    With wApp.Selection.Find
        .Text = "<<cLugar>>"
        .Replacement.Text = lsAgencia
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

     'Direccion
    With wApp.Selection.Find
        .Text = "<<cDireccion>>"
        .Replacement.Text = lsAgeDir
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'JUEZ 20150121 ****************************
    'Tipo envio estado de cuenta
    With wApp.Selection.Find
        .Text = "<<TipoEnvio>>"
        .Replacement.Text = sTipoEnvio
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'Costo envio fisico
    With wApp.Selection.Find
        .Text = "<<CostoEnvioFisico>>"
        .Replacement.Text = nCosEnvioFisico
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END JUEZ *********************************


    lsNom1 = "Nombre del Cliente: " & MatTitular(i, 1)
    lsDoc1 = "DNI/RUC: " & MatTitular(i, 4) & Space(60) & "Firma:______________________"
    lsDir1 = "Dirección: " & MatTitular(i, 5)

    With wApp.Selection.Find
        .Text = "<<NomTit1>>"
        .Replacement.Text = lsNom1
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DocTit1>>"
        .Replacement.Text = lsDoc1
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DirTit1>>"
        .Replacement.Text = lsDir1
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<NomTit2>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DocTit2>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DirTit2>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<NomTit3>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DocTit3>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DirTit3>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<NomTit4>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DocTit4>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DirTit4>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

  Next

wAppSource.ActiveDocument.Close
wApp.Visible = True

End Function


Public Sub ImpreCartillaCTS(MatTitular As Variant, ByVal psCtaCod As String, ByVal pnTasa As Double, ByVal pdFechaA As Date, Optional pnMonto As Currency = 0, Optional pnPlazo As Integer = 0, Optional nCostoMan As Currency = 0)

    Dim rs As ADODB.Recordset
    Dim nValor As Double
    Dim lsFechas As String
    Dim lsNom1 As String, lsNom2 As String, lsNom3 As String, lsNom4 As String
    Dim lsDoc1 As String, lsDoc2 As String, lsDoc3 As String, lsDoc4 As String
    Dim lsDir1 As String, lsDir2 As String, lsDir3 As String, lsDir4 As String
    Dim lsCad As String
    Dim nitf As Double
    Dim sTaInEf As String
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
    Dim lnPlazo As Long
    Dim nInteres As Double
    Dim sPlazo As String
    Dim nCosEnvioFisico As String 'JUEZ 20130520
    Dim sTipoEnvio As String 'JUEZ 20130520


    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
     'CTI2 ferimoro: ERS030-2019 ************************
    Dim nSubProducto As String
    nSubProducto = clsMant.recuperaSubProducto(gCapCTS, 1)
    '**************************************************

    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range

    Dim loRs As COMDConstSistema.DCOMGeneral
    Set loRs = New COMDConstSistema.DCOMGeneral

    Dim loAge As COMDConstantes.DCOMAgencias
    Dim rs1 As ADODB.Recordset
    Dim lsAgencia As String
    Dim lsAgeDir As String
    Dim lsDesAgencia As String 'CTI2 FERIMORO : ERS030-2019
    Set loAge = New COMDConstantes.DCOMAgencias
    Set rs1 = New ADODB.Recordset
        Set rs1 = loAge.RecuperaAgencias(gsCodAge)
        If Not (rs1.EOF And rs1.BOF) Then
            lsAgencia = Trim(rs1("cUbiGeoDescripcion"))
            lsDesAgencia = Trim(rs1("cAgeDescripcion")) 'CTI2 FERIMORO : ERS030-2019
            lsAgeDir = rs1("cAgeDireccion")
        End If
    Set loAge = Nothing
    'JUEZ 20130520 ***************************************
    Set rs = New ADODB.Recordset
    nValor = loRs.GetParametro(2000, 1005)
    '''nCosEnvioFisico = "S/. " & Format$(nValor, "0.00") 'marg ers044-2016
    nCosEnvioFisico = gcPEN_SIMBOLO & " " & Format$(nValor, "0.00") 'marg ers044-2016
     'CTI2 ERS030-2019************************
    Dim nFormaEnvio As String
    Dim nEnvioSiNo As String
    '****************************************
    Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
    Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rs = oCap.RecuperaDatosEnvioEstadoCta(psCtaCod)
    ''CTI2 FERIMORO ERS030-2019
    ''If Not rs.EOF Then sTipoEnvio = rs!cTipoEnvio 'APRI20180530 ERS036-2017
    'CTI2 ERS030-2019************************
    If Not rs.EOF Then
        sTipoEnvio = rs!cTipoEnvio
        nFormaEnvio = rs!cTipoEnvio
        nEnvioSiNo = "Si"
    Else
            nFormaEnvio = ""
            nEnvioSiNo = "No"
    End If
    '************************
    'sTipoEnvio = rs!cTipoEnvio
    Set rs = Nothing
    'END JUEZ ********************************************
    'ALPA 20091118************************************************
    Dim objCaptac As COMNCaptaGenerales.NCOMCaptaGenerales
    Set objCaptac = New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim pnPlazoT As Integer
    Dim nTREA As Currency
    pnPlazoT = CInt(pnPlazo)
    nTREA = objCaptac.ObtenerTREA(Mid$(psCtaCod, 6, 3), IIf(pnMonto = 0, 10, pnMonto), pnTasa, pnPlazoT, nCostoMan)
    '*************************************************************
    Set oWord = CreateObject("Word.Application")
    oWord.Visible = True

    Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\CARTILLACTS.doc")
     'CTI2 FERIMORO: ERS030-2019 ************************************************
    'FECHA DEL DOCUMENTO - CABECERA
    lsCad = lsDesAgencia
    With oWord.Selection.Find
        .Text = "<<Oficina>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    lsCad = Format(Time, "hh:mm:ss")
    With oWord.Selection.Find
        .Text = "<<hora>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    lsCad = gsCodUser
    With oWord.Selection.Find
        .Text = "<<user>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    '***********************************************************************
    
    'JUEZ 20150121 ************************************************
    'Cuenta
    lsCad = psCtaCod
    With oWord.Selection.Find
        .Text = "<<cCodCta>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'Monto
    lsCad = IIf(Mid(psCtaCod, 9, 1) = 1, "MN ", "ME ") & Format(pnMonto, "#,##0.00")
    With oWord.Selection.Find
        .Text = "<<Monto>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END JUEZ *****************************************************
'CTI2 FERIMORO ERS030-2019
    'Sub Producto
    With oWord.Selection.Find
        .Text = "<<SubProducto>>"
        .Replacement.Text = nSubProducto
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
'******************************************************************

    sTaInEf = pnTasa
    'Tasa
    lsCad = ""
    lsCad = sTaInEf & " % (Fija)"
    With oWord.Selection.Find
        .Text = "<<Tasa>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'ALPA 20091118****************************************
    lsCad = ""
    lsCad = nTREA & " % (Fija)"
    With oWord.Selection.Find
        .Text = "<<TasaTrea>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    '*****************************************************
    'Fecha
    lsFechas = Format(gdFecSis, "dd") & " de " & Format(gdFecSis, "mmmm") & " del " & Format(gdFecSis, "yyyy")
    With oWord.Selection.Find
        .Text = "<<FecActual>>"
        .Replacement.Text = lsFechas
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Lugar
    With oWord.Selection.Find
        .Text = "<<cLugar>>"
        .Replacement.Text = lsAgencia
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Direccion
    With oWord.Selection.Find
        .Text = "<<cDireccion>>"
        .Replacement.Text = lsAgeDir
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'JUEZ 20130520 ****************************
         'Tipo envio estado de cuenta
     'CTI2 FERIMORO ERS030-2019
    With oWord.Selection.Find
        .Text = "<<Elección SI/NO>>"
        .Replacement.Text = nEnvioSiNo
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    '**************************
    'Tipo envio estado de cuenta
    With oWord.Selection.Find
        .Text = "<<TipoEnvio>>"
        .Replacement.Text = sTipoEnvio
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'Costo envio fisico
    With oWord.Selection.Find
        .Text = "<<CostoEnvioFisico>>"
        .Replacement.Text = nCosEnvioFisico
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END JUEZ *********************************
    'ITF
    nitf = gnITFPorcent * 100
    With oWord.Selection.Find
        .Text = "<<TasaITF>>"
        '.Replacement.Text = Format$(nitf, "0.00")
        .Replacement.Text = Trim(CStr(nitf))
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    

    Dim i As Integer
    Dim rsDir As ADODB.Recordset 'JUEZ 20130520
    Dim oPers As COMDpersona.DCOMPersona 'JUEZ 20130520
    ''****CTI2 FERIMORO ERS030-2019
    Call compForCartillaP(MatTitular, oWord)
    ''*******************************************************
''        If UBound(MatTitular) = 2 Then
''
''            lsNom1 = "Nombre del Cliente: " & MatTitular(1, 1)
''            lsDoc1 = "DNI/RUC: " & MatTitular(1, 2) & space(60) & "Firma:______________________"
''            lsDir1 = "Dirección: " & MatTitular(1, 3)
''            'JUEZ 20130520 ****************************
''            Set oPers = New COMDPersona.DCOMPersona
''            Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(1, 2))
''            lsDir1 = "Dirección: " & rsDir!cPersDireccDomicilio
''            'END JUEZ *********************************
''
''            With oWord.Selection.Find
''                .Text = "<<NomTit1>>"
''                .Replacement.Text = lsNom1
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DocTit1>>"
''                .Replacement.Text = lsDoc1
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DirTit1>>"
''                .Replacement.Text = lsDir1
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<NomTit2>>"
''                .Replacement.Text = ""
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DocTit2>>"
''                .Replacement.Text = ""
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DirTit2>>"
''                .Replacement.Text = ""
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<NomTit3>>"
''                .Replacement.Text = ""
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DocTit3>>"
''                .Replacement.Text = ""
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DirTit3>>"
''                .Replacement.Text = ""
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<NomTit4>>"
''                .Replacement.Text = ""
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DocTit4>>"
''                .Replacement.Text = ""
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DirTit4>>"
''                .Replacement.Text = ""
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''
''        ElseIf UBound(MatTitular) = 3 Then
''            lsNom1 = "Nombre del Cliente: " & MatTitular(1, 1)
''            lsDoc1 = "DNI/RUC: " & MatTitular(1, 2) & space(60) & "Firma:______________________"
''            lsDir1 = "Dirección: " & MatTitular(1, 3)
''            'JUEZ 20130520 ****************************
''            Set oPers = New COMDPersona.DCOMPersona
''            Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(1, 2))
''            lsDir1 = "Dirección: " & rsDir!cPersDireccDomicilio
''            'END JUEZ *********************************
''            lsNom2 = "Nombre del Cliente: " & MatTitular(2, 1)
''            lsDoc2 = "DNI/RUC: " & MatTitular(2, 2) & space(60) & "Firma:______________________"
''            lsDir2 = "Dirección: " & MatTitular(2, 3)
''            'JUEZ 20130520 ****************************
''            Set oPers = New COMDPersona.DCOMPersona
''            Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(2, 2))
''            lsDir2 = "Dirección: " & rsDir!cPersDireccDomicilio
''            'END JUEZ *********************************
''
''            With oWord.Selection.Find
''                .Text = "<<NomTit1>>"
''                .Replacement.Text = lsNom1
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DocTit1>>"
''                .Replacement.Text = lsDoc1
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DirTit1>>"
''                .Replacement.Text = lsDir1
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<NomTit2>>"
''                .Replacement.Text = lsNom2
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DocTit2>>"
''                .Replacement.Text = lsDoc2
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DirTit2>>"
''                .Replacement.Text = lsDir2
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<NomTit3>>"
''                .Replacement.Text = ""
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DocTit3>>"
''                .Replacement.Text = ""
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DirTit3>>"
''                .Replacement.Text = ""
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<NomTit4>>"
''                .Replacement.Text = ""
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DocTit4>>"
''                .Replacement.Text = ""
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DirTit4>>"
''                .Replacement.Text = ""
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''        ElseIf UBound(MatTitular) = 4 Then
''            lsNom1 = "Nombre del Cliente: " & MatTitular(1, 1)
''            lsDoc1 = "DNI/RUC: " & MatTitular(1, 2) & space(60) & "Firma:______________________"
''            lsDir1 = "Dirección: " & MatTitular(1, 3)
''            'JUEZ 20130520 ****************************
''            Set oPers = New COMDPersona.DCOMPersona
''            Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(1, 2))
''            lsDir1 = "Dirección: " & rsDir!cPersDireccDomicilio
''            'END JUEZ *********************************
''            lsNom2 = "Nombre del Cliente: " & MatTitular(2, 1)
''            lsDoc2 = "DNI/RUC: " & MatTitular(2, 2) & space(60) & "Firma:______________________"
''            lsDir2 = "Dirección: " & MatTitular(2, 3)
''            'JUEZ 20130520 ****************************
''            Set oPers = New COMDPersona.DCOMPersona
''            Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(2, 2))
''            lsDir2 = "Dirección: " & rsDir!cPersDireccDomicilio
''            'END JUEZ *********************************
''            lsNom3 = "Nombre del Cliente: " & MatTitular(3, 1)
''            lsDoc3 = "DNI/RUC: " & MatTitular(3, 2) & space(60) & "Firma:______________________"
''            lsDir3 = "Dirección: " & MatTitular(3, 3)
''            'JUEZ 20130520 ****************************
''            Set oPers = New COMDPersona.DCOMPersona
''            Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(3, 2))
''            lsDir3 = "Dirección: " & rsDir!cPersDireccDomicilio
''            'END JUEZ *********************************
''
''            With oWord.Selection.Find
''                .Text = "<<NomTit1>>"
''                .Replacement.Text = lsNom1
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DocTit1>>"
''                .Replacement.Text = lsDoc1
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DirTit1>>"
''                .Replacement.Text = lsDir1
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<NomTit2>>"
''                .Replacement.Text = lsNom2
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DocTit2>>"
''                .Replacement.Text = lsDoc2
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DirTit2>>"
''                .Replacement.Text = lsDir2
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<NomTit3>>"
''                .Replacement.Text = lsNom3
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DocTit3>>"
''                .Replacement.Text = lsDoc3
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DirTit3>>"
''                .Replacement.Text = lsDir3
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<NomTit4>>"
''                .Replacement.Text = ""
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DocTit4>>"
''                .Replacement.Text = ""
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DirTit4>>"
''                .Replacement.Text = ""
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''
''        ElseIf UBound(MatTitular) = 5 Then
''            lsNom1 = "Nombre del Cliente: " & MatTitular(1, 1)
''            lsDoc1 = "DNI/RUC: " & MatTitular(1, 2) & space(60) & "Firma:______________________"
''            lsDir1 = "Dirección: " & MatTitular(1, 3)
''            'JUEZ 20130520 ****************************
''            Set oPers = New COMDPersona.DCOMPersona
''            Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(1, 2))
''            lsDir1 = "Dirección: " & rsDir!cPersDireccDomicilio
''            'END JUEZ *********************************
''            lsNom2 = "Nombre del Cliente: " & MatTitular(2, 1)
''            lsDoc2 = "DNI/RUC: " & MatTitular(2, 2) & space(60) & "Firma:______________________"
''            lsDir2 = "Dirección: " & MatTitular(2, 3)
''            'JUEZ 20130520 ****************************
''            Set oPers = New COMDPersona.DCOMPersona
''            Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(2, 2))
''            lsDir2 = "Dirección: " & rsDir!cPersDireccDomicilio
''            'END JUEZ *********************************
''            lsNom3 = "Nombre del Cliente: " & MatTitular(3, 1)
''            lsDoc3 = "DNI/RUC: " & MatTitular(3, 2) & space(60) & "Firma:______________________"
''            lsDir3 = "Dirección: " & MatTitular(3, 3)
''            'JUEZ 20130520 ****************************
''            Set oPers = New COMDPersona.DCOMPersona
''            Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(3, 2))
''            lsDir3 = "Dirección: " & rsDir!cPersDireccDomicilio
''            'END JUEZ *********************************
''            lsNom4 = "Nombre del Cliente: " & MatTitular(4, 1)
''            lsDoc4 = "DNI/RUC: " & MatTitular(4, 2) & space(60) & "Firma:______________________"
''            lsDir4 = "Dirección: " & MatTitular(4, 3)
''            'JUEZ 20130520 ****************************
''            Set oPers = New COMDPersona.DCOMPersona
''            Set rsDir = oPers.RecuperaPersonaEnvioEstadoCtaDoc(MatTitular(4, 2))
''            lsDir4 = "Dirección: " & rsDir!cPersDireccDomicilio
''            'END JUEZ *********************************
''
''            With oWord.Selection.Find
''                .Text = "<<NomTit1>>"
''                .Replacement.Text = lsNom1
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DocTit1>>"
''                .Replacement.Text = lsDoc1
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DirTit1>>"
''                .Replacement.Text = lsDir1
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<NomTit2>>"
''                .Replacement.Text = lsNom2
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DocTit2>>"
''                .Replacement.Text = lsDoc2
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DirTit2>>"
''                .Replacement.Text = lsDir2
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<NomTit3>>"
''                .Replacement.Text = lsNom3
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DocTit3>>"
''                .Replacement.Text = lsDoc3
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DirTit3>>"
''                .Replacement.Text = lsDir3
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<NomTit4>>"
''                .Replacement.Text = lsNom4
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DocTit4>>"
''                .Replacement.Text = lsDoc4
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''            With oWord.Selection.Find
''                .Text = "<<DirTit4>>"
''                .Replacement.Text = lsDir4
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
''            End With
''        End If



   oDoc.SaveAs App.Path & "\SPOOLER\" & psCtaCod & ".doc"
   oDoc.Close
   oWord.Quit
   
        Dim doc As New Word.Application
        With doc
            .Documents.Open App.Path & "\SPOOLER\" & psCtaCod & ".doc" 'abrimos "Mi documento"
            .Visible = True 'hacemos visible Word
        End With
   
End Sub

'***Agregado por ELRO el 20121201, según OYP-RFC101-2012
Public Sub ImpreCartillaCTSLote(prsCuenta As ADODB.Recordset, MatNroCta() As String, psMovNro As String)
    Dim rs As ADODB.Recordset
    Dim lsFechas As String
    Dim lsNom1 As String, lsDoc1 As String, lsDir1 As String
    Dim lsCad As String
    Dim sTaInEf As String
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim lsModeloPlantilla As String
    Dim lsCtaCod As String
    Dim lnMonto As Currency
    Dim lnTasa As Currency
    Dim nTREA As Currency
    Dim i As Integer
    Dim nValor As Double
    Dim nCosEnvioFisico As String 'JUEZ 20150121
    Dim sTipoEnvio As String 'JUEZ 20150121

    'Crea una clase que de Word Object
    Dim wApp As Word.Application
    Dim wAppSource As Word.Application
    'Create a new instance of word
    Set wApp = New Word.Application
    Set wAppSource = New Word.Application

    Dim loRs As COMDConstSistema.DCOMGeneral
    Set loRs = New COMDConstSistema.DCOMGeneral

    Dim loAge As COMDConstantes.DCOMAgencias
    Dim rs1 As ADODB.Recordset
    Dim lsAgencia As String
    Dim lsAgeDir As String

    Set loAge = New COMDConstantes.DCOMAgencias
    Set rs1 = New ADODB.Recordset
        Set rs1 = loAge.RecuperaAgencias(gsCodAge)
        If Not (rs1.EOF And rs1.BOF) Then
            lsAgencia = Trim(rs1("cUbiGeoDescripcion"))
            lsAgeDir = rs1("cAgeDireccion")
        End If
    Set loAge = Nothing

    Dim objCaptac As COMNCaptaGenerales.NCOMCaptaGenerales
    Set objCaptac = New COMNCaptaGenerales.NCOMCaptaGenerales

    lsModeloPlantilla = App.Path & "\FormatoCarta\CARTILLACTS2.doc"

    Dim RangeSource As Word.Range
    'Abre Documento Plantilla
    wAppSource.Documents.Open FileName:=lsModeloPlantilla
    Set RangeSource = wAppSource.ActiveDocument.Content
    'Lo carga en Memoria
    wAppSource.ActiveDocument.Content.Copy

    'Crea Nuevo Documento
    wApp.Documents.Add

'    With wApp.ActiveDocument.PageSetup
'        .LeftMargin = CentimetersToPoints(1.5)
'        .RightMargin = CentimetersToPoints(1)
'        .TopMargin = CentimetersToPoints(1.5)
'        .BottomMargin = CentimetersToPoints(1)
'    End With

    With wApp.Selection.Font
        .Name = "Arial Narrow"
        .Size = 12
    End With

    'JUEZ 20150121 *************************************
    With wApp.Selection.PageSetup
        .PaperSize = wdPaperLetter
        .TopMargin = wApp.CentimetersToPoints(1.75)
        .BottomMargin = wApp.CentimetersToPoints(1.5)
        .LeftMargin = wApp.CentimetersToPoints(1.3)
        .RightMargin = wApp.CentimetersToPoints(1.2)
    End With
    'END JUEZ ******************************************

    i = 1
    prsCuenta.MoveFirst
    Do While Not prsCuenta.EOF

        wApp.Application.Selection.TypeParagraph
        'wApp.Application.Selection.Paste
        wApp.Application.Selection.PasteAndFormat (wdPasteDefault)
        wApp.Application.Selection.InsertBreak
        wApp.Selection.SetRange start:=wApp.Selection.start, End:=wApp.ActiveDocument.Content.End
        wApp.Selection.MoveEnd

        lsCtaCod = CStr(MatNroCta(i))
        lnMonto = CCur(prsCuenta("Monto"))
        lnTasa = CCur(prsCuenta("Tasa"))

        'JUEZ 20150121 ***************************************
        Set rs = New ADODB.Recordset
        nValor = loRs.GetParametro(2000, 1005)
        '''nCosEnvioFisico = "S/. " & Format$(nValor, "0.00") 'marg ers044-2016
        nCosEnvioFisico = gcPEN_SIMBOLO & " " & Format$(nValor, "0.00") 'marg ers044-2016

        Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
        Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
        Set rs = oCap.RecuperaDatosEnvioEstadoCta(lsCtaCod)
        'sTipoEnvio = "Digital"
        sTipoEnvio = "" 'APRI20180530 ERS036-2017
        If rs Is Nothing Then sTipoEnvio = rs!cTipoEnvio
        Set rs = Nothing
        'END JUEZ ********************************************

        nTREA = objCaptac.ObtenerTREA(Mid$(lsCtaCod, 6, 3), IIf(lnMonto = 0, 10, lnMonto), lnTasa)

        'JUEZ 20150121 ************************************************
        'Cuenta
        lsCad = lsCtaCod
        With wApp.Selection.Find
            .Text = "<<cCodCta>>"
            .Replacement.Text = lsCad
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        'Monto
        lsCad = IIf(Mid(lsCtaCod, 9, 1) = 1, "MN ", "ME ") & Format(lnMonto, "#,##0.00")
        With wApp.Selection.Find
            .Text = "<<Monto>>"
            .Replacement.Text = lsCad
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        'END JUEZ *****************************************************
        sTaInEf = prsCuenta("Tasa")
        'Tasa
        lsCad = ""
        lsCad = sTaInEf & " % (Fija)"
        With wApp.Selection.Find
            .Text = "<<Tasa>>"
            .Replacement.Text = lsCad
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With

        lsCad = ""
        lsCad = nTREA & " % (Fija)"
        With wApp.Selection.Find
            .Text = "<<TasaTrea>>"
            .Replacement.Text = lsCad
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With

        'Fecha
        lsFechas = Format(gdFecSis, "dd") & " de " & Format(gdFecSis, "mmmm") & " del " & Format(gdFecSis, "yyyy")
        With wApp.Selection.Find
            .Text = "<<FecActual>>"
            .Replacement.Text = lsFechas
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With

        'Lugar
        With wApp.Selection.Find
            .Text = "<<cLugar>>"
            .Replacement.Text = lsAgencia
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With

        'Direccion
        With wApp.Selection.Find
            .Text = "<<cDireccion>>"
            .Replacement.Text = lsAgeDir
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        'JUEZ 20150121 ****************************
        'Tipo envio estado de cuenta
        With wApp.Selection.Find
            .Text = "<<TipoEnvio>>"
            .Replacement.Text = sTipoEnvio
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        'Costo envio fisico
        With wApp.Selection.Find
            .Text = "<<CostoEnvioFisico>>"
            .Replacement.Text = nCosEnvioFisico
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        'END JUEZ *********************************


        lsNom1 = "Nombre del Cliente: " & prsCuenta("Nombre")
        lsDoc1 = "DNI/RUC: " & prsCuenta.Fields(14) & Space(60) & "Firma:______________________"
        lsDir1 = "Dirección: " & IIf(prsCuenta.Fields(13) = "", prsCuenta.Fields(15), prsCuenta.Fields(15))

            With wApp.Selection.Find
                .Text = "<<NomTit1>>"
                .Replacement.Text = lsNom1
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With wApp.Selection.Find
                .Text = "<<DocTit1>>"
                .Replacement.Text = lsDoc1
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With wApp.Selection.Find
                .Text = "<<DirTit1>>"
                .Replacement.Text = lsDir1
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            i = i + 1
     prsCuenta.MoveNext
     Loop
    wAppSource.ActiveDocument.Close
    wAppSource.Quit
    wApp.ActiveDocument.SaveAs (App.Path & "\SPOOLER\AperturaLoteCTS_" & psMovNro & ".doc")
    wApp.Visible = True
End Sub
'***Fin Agregado por ELRO el 20121201*******************


Public Sub ImpreCartillaAhoNanito(MatTitular As Variant, ByVal psCtaCod As String, ByVal pnTasa As Double, ByVal pnMonto As Double, Optional pnPlazo As Integer = 0, Optional nCostoMan As Currency = 0)

    Dim rs As ADODB.Recordset
    Dim nValor As Double
    Dim nMonMinCta As String
    Dim nMonxConsul As String
    Dim nMonComRet As String
    Dim nCosInac As Double
    Dim nMonOtraPlaza As String
    Dim nTasaITF As Double
    Dim nMonMinRetMN As String
    Dim nMonMinRetME As String
    Dim nParCod  As Long
    Dim lsFechas As String
    Dim lsNom1 As String, lsNom2 As String, lsNom3 As String, lsNom4 As String
    Dim lsDoc1 As String, lsDoc2 As String, lsDoc3 As String, lsDoc4 As String
    Dim lsDir1 As String, lsDir2 As String, lsDir3 As String, lsDir4 As String
    Dim lsRel1 As String, lsRel2 As String, lsRel3 As String, lsRel4 As String 'JUEZ 20150121
    Dim lsCad As String
    Dim nitf As Double
    Dim nCosEnvioFisico As String 'JUEZ 20130520
    Dim sTipoEnvio As String 'JUEZ 20130520
    'APRI20190109 ERS077-2018
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim nSaldoEquilibrio As Currency
    nSaldoEquilibrio = clsMant.ObtenerSaldoEquilibrio(nCostoMan, pnTasa)
    'END APRI

    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range

    Dim loRs As COMDConstSistema.DCOMGeneral
    Set loRs = New COMDConstSistema.DCOMGeneral

    Dim loAge As COMDConstantes.DCOMAgencias
    Dim rs1 As ADODB.Recordset
    Dim lsAgencia As String
    Dim lsAgeDir As String
    Dim lsDesAgencia As String 'CTI2 FERIMORO : ERS030-2019 JATO 202103
    Set loAge = New COMDConstantes.DCOMAgencias
    Set rs1 = New ADODB.Recordset
        Set rs1 = loAge.RecuperaAgencias(gsCodAge)
        If Not (rs1.EOF And rs1.BOF) Then
            lsAgeDir = rs1("cAgeDireccion")
            lsDesAgencia = Trim(rs1("cAgeDescripcion")) 'CTI2 FERIMORO : ERS030-2019 JATO 202103
            lsAgencia = Trim(rs1("cUbiGeoDescripcion"))
        End If
    Set loAge = Nothing

    'CTI2 FERIMORO ER030-2019 JATO 202103
    Dim nSubProducto As String
    nSubProducto = clsMant.recuperaSubProducto(gCapAhorros, 1)
    '**************************************************
    'JUEZ 20150121 Nuevos Parámetros *******************
    Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
    Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
    Dim rsPar As ADODB.Recordset
    Set rsPar = clsDef.GetCapParametroNew(gCapAhorros, 0)
    'END JUEZ ******************************************

    '--------Saldo x Cuenta a Mantener por 3 Meses --

    If Mid(psCtaCod, 9, 1) = 1 Then
       nParCod = 2091
    Else
       nParCod = 2092
    End If
    'nValor = loRs.GetParametro(2000, nParCod)
    nValor = 0 'IIf(Mid(psCtaCod, 9, 1) = 1, rsPar!nSaldoMinCtaSol, rsPar!nSaldoMinCtaDol) 'JUEZ 20150121 'APRI20190109 ERS077-2018
    '''nMonMinCta = IIf(Mid(psCtaCod, 9, 1) = 1, "S/. ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(psCtaCod, 9, 1) = 1, " (" & UnNumero(nValor) & " Nuevos Soles)", " (" & UnNumero(nValor) & " DOLARES)") 'marg ers044-2016
    nMonMinCta = IIf(Mid(psCtaCod, 9, 1) = 1, gcPEN_SIMBOLO & " ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(psCtaCod, 9, 1) = 1, " (" & UnNumero(nValor) & " " & StrConv(gcPEN_PLURAL, vbProperCase) & ")", " (" & UnNumero(nValor) & " DOLARES)") 'marg ers044-2016

    '------------------------------------------------

    '-----Monto de Consulta por extracto de Cuenta --
    'By capi 05032008

'    Set rs = New ADODB.Recordset
'    If Mid(psCtaCod, 9, 1) = 1 Then
'       nParCod = gDctoExtMNxPag
'    Else
'       nParCod = gDctoExtMExPag
'    End If
'    nValor = loRs.GetParametro(2000, nParCod)
'    nMonxConsul = IIf(Mid(psCtaCod, 9, 1) = 1, "S/. ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(psCtaCod, 9, 1) = 1, " (" & UnNumero(nValor) & " Nuevos Soles)", " (" & UnNumero(nValor) & " Dolares)")

    '------------------------------------------------
     '-----Comision Consulta Saldo Ventanilla --
    Set rs = New ADODB.Recordset
    nParCod = 2106
    nValor = loRs.GetParametro(2000, nParCod)
    '''nMonxConsul = "S/." & Format$(nValor, "0.00") & " (" & UnNumero(nValor) & " Nuevos Soles)" 'marg ers044-2016
    nMonxConsul = gcPEN_SIMBOLO & Format$(nValor, "0.00") & " (" & UnNumero(nValor) & " " & StrConv(gcPEN_PLURAL, vbProperCase) & ")" 'marg ers044-2016


    '-----Comision de retiro cuando el monto es menor a 1000 --
    Set rs = New ADODB.Recordset
    If Mid(psCtaCod, 9, 1) = 1 Then
       nParCod = gMontoMNx31Ope
    Else
       nParCod = gMontoMEx31Ope
    End If
    nValor = loRs.GetParametro(2000, nParCod)
    '''nMonComRet = IIf(Mid(psCtaCod, 9, 1) = 1, "S/. ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(psCtaCod, 9, 1) = 1, " (" & UnNumero(nValor) & " Nuevos Soles)", " (" & UnNumero(nValor) & " Dolares)") 'marg ers044-2016
    nMonComRet = IIf(Mid(psCtaCod, 9, 1) = 1, gcPEN_SIMBOLO & " ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(psCtaCod, 9, 1) = 1, " (" & UnNumero(nValor) & " " & StrConv(gcPEN_PLURAL, vbProperCase) & ")", " (" & UnNumero(nValor) & " Dolares)") 'marg ers044-2016
    '-----------------------------------------------------------

    '--------- Costo de Inactivas -------------------
    Set rs = New ADODB.Recordset
    nValor = loRs.GetParametro(2000, gMonDescInacME)
    nCosInac = Format$(nValor, "0.00")

    '------------------------------------------------

    '--Retiros/Depositos en Agencias de La Caja en Otras Plazas --
    Set rs = New ADODB.Recordset
    If Mid(psCtaCod, 9, 1) = 1 Then
       nParCod = 2046
    Else
       nParCod = 2047
    End If
    nValor = loRs.GetParametro(2000, nParCod)
    '''nMonOtraPlaza = IIf(Mid(psCtaCod, 9, 1) = 1, "S/. ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(psCtaCod, 9, 1) = 1, " (" & UnNumero(nValor) & " Nuevos Soles)", " (" & UnNumero(nValor) & " Dolares)") 'marg ers044-2016
    nMonOtraPlaza = IIf(Mid(psCtaCod, 9, 1) = 1, gcPEN_SIMBOLO & " ", "$. ") & Format$(nValor, "0.00") & IIf(Mid(psCtaCod, 9, 1) = 1, " (" & UnNumero(nValor) & " " & StrConv(gcPEN_PLURAL, vbProperCase), " (" & UnNumero(nValor) & " Dolares)") 'marg ers044-2016

    '--------------------------------------------------------------

    '---------------------Monto Minimo de Retiro  -----------------
    Set rs = New ADODB.Recordset

       nParCod = 2027
       nValor = loRs.GetParametro(2000, nParCod)
       '''nMonMinRetMN = "S/. " & Format$(nValor, "0.00") & " (" & UnNumero(nValor) & " Nuevos Soles)" 'marg ers044-2016
       nMonMinRetMN = gcPEN_SIMBOLO & " " & Format$(nValor, "0.00") & " (" & UnNumero(nValor) & " " & StrConv(gcPEN_PLURAL, vbProperCase) & ")" 'marg ers044-2016
       nParCod = 2028
       nValor = loRs.GetParametro(2000, nParCod)
       nMonMinRetME = "$. " & Format$(nValor, "0.00") & " (" & UnNumero(nValor) & " Dolares)"


    '--------------------------------------------------------------
    'JUEZ 20130520 ***************************************
    Set rs = New ADODB.Recordset
    nValor = loRs.GetParametro(2000, 1005)
    '''nCosEnvioFisico = "S/. " & Format$(nValor, "0.00") 'marg ers044-2016
    nCosEnvioFisico = gcPEN_SIMBOLO & " " & Format$(nValor, "0.00") 'marg ers044-2016

    Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
    Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rs = oCap.RecuperaDatosEnvioEstadoCta(psCtaCod)
     'CTI2 ERS030-2019************************
    Dim nFormaEnvio As String
    Dim nEnvioSiNo As String
    '****************************************
    'sTipoEnvio = rs!cTipoEnvio
    'APRI2018 ERS036-2017
        If Not (rs.BOF And rs.EOF) Then
       If rs.RecordCount > 0 Then
         sTipoEnvio = oCap.MostarTextoComisionEnvioEstAhorros(rs!nModoEnvio, 1)
         'CTI2 ERS030-2019************************
         nFormaEnvio = rs!cTipoEnvio
         nEnvioSiNo = "Si"
         '************************
       Else 'CTI2 ERS030-2019************************
         nFormaEnvio = "" 'CTI2 ERS030-2019************************
         nEnvioSiNo = "No" 'CTI2 ERS030-2019************************
       End If
    Else
         sTipoEnvio = ""
         nFormaEnvio = "" 'CTI2 ERS030-2019************************
         nEnvioSiNo = "No" 'CTI2 ERS030-2019************************
    End If
    'END APRI
    Set rs = Nothing
    'END JUEZ ********************************************
    'ALPA 20091118************************************************
    Dim objCaptac As COMNCaptaGenerales.NCOMCaptaGenerales
    Set objCaptac = New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim pnPlazoT As Integer
    Dim nTREA As Currency
    pnPlazoT = CInt(pnPlazo)
    nTREA = objCaptac.ObtenerTREA(Mid$(psCtaCod, 6, 3), IIf(pnMonto = 0, 10, pnMonto), pnTasa, pnPlazoT, nCostoMan)
    '*************************************************************
    Set oWord = CreateObject("Word.Application")
    oWord.Visible = True

    Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\CARTILLAAHORRONanito.doc")
    oDoc.SaveAs App.Path & "\SPOOLER\" & psCtaCod & ".doc" 'ADD JHCU 27-04-2019

    'JUEZ 20150121 ************************************************
    'Cuenta
    lsCad = psCtaCod
    With oWord.Selection.Find
        .Text = "<<cCodCta>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END JUEZ *****************************************************
    'CTI2 FERIMORO: ERS030-2019 ************************************************
    lsCad = nSubProducto
    With oWord.Selection.Find
        .Text = "<<SubProducto>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    lsCad = lsDesAgencia
    With oWord.Selection.Find
        .Text = "<<Oficina>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    lsCad = Format(Time, "hh:mm:ss")
    With oWord.Selection.Find
        .Text = "<<hora>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    lsCad = gsCodUser
    With oWord.Selection.Find
        .Text = "<<user>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    lsCad = nEnvioSiNo
    With oWord.Selection.Find
        .Text = "<<Elección SI/NO>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    lsCad = nFormaEnvio
    With oWord.Selection.Find
        .Text = "<<TipoEnvio>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    '***********************************************************************
    
    'Monto
    lsCad = IIf(Mid(psCtaCod, 9, 1) = 1, "MN ", "ME ") & Format(pnMonto, "#,##0.00")
    With oWord.Selection.Find
        .Text = "<<Monto>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Tasa
    lsCad = ""
    lsCad = Format(pnTasa, "0.00") & " % (Fija)"
    With oWord.Selection.Find
        .Text = "<<Tasa>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'ALPA 20091118****************************************
    lsCad = ""
    lsCad = Format(nTREA, "0.00") & " % (Fija)"
    With oWord.Selection.Find
        .Text = "<<TasaTrea>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    '*****************************************************
    'APRI20190109 ERS077-2018
    lsCad = ""
    lsCad = Format(nSaldoEquilibrio, "0.00")
    With oWord.Selection.Find
        .Text = "<<SaldoEquilibrio>>"
        .Replacement.Text = lsCad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END APRI
    'Saldo de la cuenta por 3 meses
    With oWord.Selection.Find
        .Text = "<<MontoMin>>"
        .Replacement.Text = nMonMinCta
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Monto x Consulta
    With oWord.Selection.Find
        .Text = "<<MonConsul>>"
        .Replacement.Text = nMonxConsul
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Comision por Retiro
    With oWord.Selection.Find
        .Text = "<<MonComRet>>"
        .Replacement.Text = nMonComRet
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With


    'Monto de Otra Plaza
    With oWord.Selection.Find
        .Text = "<<MonOtraPlaza>>"
        .Replacement.Text = nMonOtraPlaza
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Monto de Otra Plaza
    With oWord.Selection.Find
        .Text = "<<MonOtraPlaza>>"
        .Replacement.Text = nMonOtraPlaza
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With


    'ITF
    nitf = gnITFPorcent * 100
    With oWord.Selection.Find
        .Text = "<<TasaITF>>"
        '.Replacement.Text = Format$(nitf, "0.00")
        .Replacement.Text = Trim(CStr(nitf))
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Fecha
    lsFechas = Format(gdFecSis, "dd") & " de " & Format(gdFecSis, "mmmm") & " del " & Format(gdFecSis, "yyyy")
    With oWord.Selection.Find
        .Text = "<<FecActual>>"
        .Replacement.Text = lsFechas
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Lugar
    With oWord.Selection.Find
        .Text = "<<cLugar>>"
        .Replacement.Text = lsAgencia
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Direccion
    With oWord.Selection.Find
        .Text = "<<cDireccion>>"
        .Replacement.Text = lsAgeDir
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'JUEZ 20130520 ****************************
    'Tipo envio estado de cuenta
    With oWord.Selection.Find
        .Text = "<<TipoEnvio>>"
        .Replacement.Text = sTipoEnvio
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'Costo envio fisico
    With oWord.Selection.Find
        .Text = "<<CostoEnvioFisico>>"
        .Replacement.Text = nCosEnvioFisico
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'END JUEZ *********************************

    Dim i As Integer

    ''CTI2 FERIMORO ERS030-2019
Call compForCartillaP(MatTitular, oWord, True)
''*****************************************
''    If UBound(MatTitular) = 2 Then
''        lsRel1 = Trim(MatTitular(1, 4))
''        lsNom1 = "Nombre del " & lsRel1 & ": " & MatTitular(1, 1)
''        lsDoc1 = "DNI/RUC: " & MatTitular(1, 2) & IIf(lsRel1 = "TITULAR", "", space(60) & "Firma:______________________")
''        lsDir1 = "Dirección: " & MatTitular(1, 3)
''
''        With oWord.Selection.Find
''            .Text = "<<NomTit1>>"
''            .Replacement.Text = lsNom1
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit1>>"
''            .Replacement.Text = lsDoc1
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit1>>"
''            .Replacement.Text = lsDir1
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<NomTit2>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit2>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit2>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<NomTit3>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit3>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit3>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<NomTit4>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit4>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit4>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''
''    ElseIf UBound(MatTitular) = 3 Then
''        lsRel1 = Trim(MatTitular(1, 4))
''        lsNom1 = "Nombre del " & lsRel1 & ": " & MatTitular(1, 1)
''        lsDoc1 = "DNI/RUC: " & MatTitular(1, 2) & IIf(lsRel1 = "TITULAR", "", space(60) & "Firma:______________________")
''        lsDir1 = "Dirección: " & MatTitular(1, 3)
''        lsRel2 = Trim(MatTitular(2, 4))
''        lsNom2 = "Nombre del " & lsRel2 & ": " & MatTitular(2, 1)
''        lsDoc2 = "DNI/RUC: " & MatTitular(2, 2) & IIf(lsRel2 = "TITULAR", "", space(60) & "Firma:______________________")
''        lsDir2 = "Dirección: " & MatTitular(2, 3)
''
''        With oWord.Selection.Find
''            .Text = "<<NomTit1>>"
''            .Replacement.Text = lsNom1
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit1>>"
''            .Replacement.Text = lsDoc1
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit1>>"
''            .Replacement.Text = lsDir1
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<NomTit2>>"
''            .Replacement.Text = lsNom2
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit2>>"
''            .Replacement.Text = lsDoc2
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit2>>"
''            .Replacement.Text = lsDir2
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<NomTit3>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit3>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit3>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<NomTit4>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit4>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit4>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''    ElseIf UBound(MatTitular) = 4 Then
''        lsRel1 = Trim(MatTitular(1, 4))
''        lsNom1 = "Nombre del " & lsRel1 & ": " & MatTitular(1, 1)
''        lsDoc1 = "DNI/RUC: " & MatTitular(1, 2) & IIf(lsRel1 = "TITULAR", "", space(60) & "Firma:______________________")
''        lsDir1 = "Dirección: " & MatTitular(1, 3)
''        lsRel2 = Trim(MatTitular(2, 4))
''        lsNom2 = "Nombre del " & lsRel2 & ": " & MatTitular(2, 1)
''        lsDoc2 = "DNI/RUC: " & MatTitular(2, 2) & IIf(lsRel2 = "TITULAR", "", space(60) & "Firma:______________________")
''        lsDir2 = "Dirección: " & MatTitular(2, 3)
''        lsRel3 = Trim(MatTitular(3, 4))
''        lsNom3 = "Nombre del " & lsRel3 & ": " & MatTitular(3, 1)
''        lsDoc3 = "DNI/RUC: " & MatTitular(3, 2) & IIf(lsRel3 = "TITULAR", "", space(60) & "Firma:______________________")
''        lsDir3 = "Dirección: " & MatTitular(3, 3)
''
''        With oWord.Selection.Find
''            .Text = "<<NomTit1>>"
''            .Replacement.Text = lsNom1
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit1>>"
''            .Replacement.Text = lsDoc1
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit1>>"
''            .Replacement.Text = lsDir1
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<NomTit2>>"
''            .Replacement.Text = lsNom2
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit2>>"
''            .Replacement.Text = lsDoc2
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit2>>"
''            .Replacement.Text = lsDir2
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<NomTit3>>"
''            .Replacement.Text = lsNom3
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit3>>"
''            .Replacement.Text = lsDoc3
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit3>>"
''            .Replacement.Text = lsDir3
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<NomTit4>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit4>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit4>>"
''            .Replacement.Text = ""
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''
''    ElseIf UBound(MatTitular) = 5 Then
''        lsRel1 = Trim(MatTitular(1, 4))
''        lsNom1 = "Nombre del " & lsRel1 & ": " & MatTitular(1, 1)
''        lsDoc1 = "DNI/RUC: " & MatTitular(1, 2) & IIf(lsRel1 = "TITULAR", "", space(60) & "Firma:______________________")
''        lsDir1 = "Dirección: " & MatTitular(1, 3)
''        lsRel2 = Trim(MatTitular(2, 4))
''        lsNom2 = "Nombre del " & lsRel2 & ": " & MatTitular(2, 1)
''        lsDoc2 = "DNI/RUC: " & MatTitular(2, 2) & IIf(lsRel2 = "TITULAR", "", space(60) & "Firma:______________________")
''        lsDir2 = "Dirección: " & MatTitular(2, 3)
''        lsRel3 = Trim(MatTitular(3, 4))
''        lsNom3 = "Nombre del " & lsRel3 & ": " & MatTitular(3, 1)
''        lsDoc3 = "DNI/RUC: " & MatTitular(3, 2) & IIf(lsRel3 = "TITULAR", "", space(60) & "Firma:______________________")
''        lsDir3 = "Dirección: " & MatTitular(3, 3)
''        lsRel4 = Trim(MatTitular(4, 4))
''        lsNom4 = "Nombre del " & lsRel4 & ": " & MatTitular(4, 1)
''        lsDoc4 = "DNI/RUC: " & MatTitular(4, 2) & IIf(lsRel4 = "TITULAR", "", space(60) & "Firma:______________________")
''        lsDir4 = "Dirección: " & MatTitular(4, 3)
''
''        With oWord.Selection.Find
''            .Text = "<<NomTit1>>"
''            .Replacement.Text = lsNom1
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit1>>"
''            .Replacement.Text = lsDoc1
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit1>>"
''            .Replacement.Text = lsDir1
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<NomTit2>>"
''            .Replacement.Text = lsNom2
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit2>>"
''            .Replacement.Text = lsDoc2
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit2>>"
''            .Replacement.Text = lsDir2
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<NomTit3>>"
''            .Replacement.Text = lsNom3
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit3>>"
''            .Replacement.Text = lsDoc3
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit3>>"
''            .Replacement.Text = lsDir3
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<NomTit4>>"
''            .Replacement.Text = lsNom4
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DocTit4>>"
''            .Replacement.Text = lsDoc4
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''        With oWord.Selection.Find
''            .Text = "<<DirTit4>>"
''            .Replacement.Text = lsDir4
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''            .Execute Replace:=wdReplaceAll
''        End With
''    End If

    'MEJORAS JHCU
    'oWord.Visible = True 'ADD JHCU 26-04-2019 ACTA 050-2019
    oDoc.Save
    oDoc.Application.Quit

    Dim doc As New Word.Application
    With doc
        .Documents.Open App.Path & "\SPOOLER\" & psCtaCod & ".doc" 'abrimos "Mi documento"
        .Visible = True 'hacemos visible Word
    End With
   'oDoc.SaveAs App.Path & "\SPOOLER\" & psCtaCod & ".doc" 'cOMENTADO POR JHCU 27-04-2019
End Sub
'FRHU 20140608 TI-ERS068-2014
Public Sub ImprimeSolicitudSiniestro(ByVal psNumSolicitud As String, ByVal psPersCod As String, ByVal psNomPers As String, ByVal pMatDocumento As Variant, ByVal pMatDocumentoDes As Variant, ByVal pdFecSis As Date, _
                                     ByVal pdFechaSiniestro As Date, ByVal psHoraSiniestro As String)
    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range
    Dim oUPCli As New UPersona_Cli
    Dim ClsPersona As New COMDpersona.DCOMPersonas
    Dim R As ADODB.Recordset

    Dim i As Integer, fila As Integer, X As Integer
    Dim lsFechas As String
    Dim NroSolicitud As String
    Dim sApePat As String, sApeMat As String, sNombres As String
    Dim sFechaSiniestro As String, sDocPersona As String

    Set oWord = CreateObject("Word.Application")
    oWord.Visible = True

    'Set oDoc = oWord.Documents.Open(App.path & "\FormatoCarta\CartaSiniestralidad.docx")
    Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\CartaSiniestralidad.doc") 'FRHU 20140717 Observacion

    Set R = ClsPersona.BuscaCliente(psPersCod, BusquedaCodigo)
    If Not R.BOF And Not R.EOF Then
        sDocPersona = R!cPersIDNroDNI
    End If

    sApePat = oUPCli.RecuperarDetalleNombrePersona(Trim(psNomPers), 1)
    sApeMat = oUPCli.RecuperarDetalleNombrePersona(Trim(psNomPers), 2)
    sNombres = oUPCli.RecuperarDetalleNombrePersona(Trim(psNomPers), 3)

    'Numero de Solicitud
    With oWord.Selection.Find
        .Text = "<<NroCarta>>"
        .Replacement.Text = psNumSolicitud
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'Fecha
    lsFechas = "Iquitos, " & Format(gdFecSis, "dd") & " de " & Format(gdFecSis, "mmmm") & " del " & Format(gdFecSis, "yyyy")
    With oWord.Selection.Find
        .Text = "<<FecActual>>"
        .Replacement.Text = lsFechas
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'Fecha Siniestro
    sFechaSiniestro = Format(pdFechaSiniestro, "dd") & " de " & Format(pdFechaSiniestro, "mmmm") & " del " & Format(pdFechaSiniestro, "yyyy")
    With oWord.Selection.Find
        .Text = "<<FecSiniestro>>"
        .Replacement.Text = sFechaSiniestro
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'Hora Siniestro
    With oWord.Selection.Find
        .Text = "<<HoraSiniestro>>"
        .Replacement.Text = psHoraSiniestro
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'Apellido Paterno
    With oWord.Selection.Find
        .Text = "<<ApellidoPaterno>>"
        .Replacement.Text = sApePat
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'Apellido Materno
    With oWord.Selection.Find
        .Text = "<<ApellidoMaterno>>"
        .Replacement.Text = sApeMat
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'Nombre Persona
    With oWord.Selection.Find
        .Text = "<<Nombre>>"
        .Replacement.Text = sNombres
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'Dni Persona
    With oWord.Selection.Find
        .Text = "<<Dni>>"
        .Replacement.Text = sDocPersona
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    'Documento
    i = 0
    For fila = LBound(pMatDocumento) To UBound(pMatDocumento)
        If pMatDocumento(fila) <> 0 Then
            i = i + 1
            With oWord.Selection.Find
            .Text = "<<Doc" & CStr(i) & ">>"
            .Replacement.Text = "- " & CStr(pMatDocumentoDes(fila))
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
            End With
        End If
    Next

    i = i + 1
    For X = i To 6
        With oWord.Selection.Find
            .Text = "<<Doc" & CStr(X) & ">>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
    Next
    oDoc.SaveAs App.Path & "\SPOOLER\SolicitudSiniestro" & psNumSolicitud & ".doc"
End Sub
'FIN FRHU 20140608
'******APRI 20170613 SEGUN RFC1705230001
Public Function PersonaActualizarCIIUExcel(ByVal pArrayPersona As Variant)

Dim ApExcel As Variant
Set ApExcel = CreateObject("Excel.application")

ApExcel.Workbooks.Add

ApExcel.Cells(2, 2).Formula = gsNomCmac
ApExcel.Cells(3, 2).Formula = gsNomAge
ApExcel.Cells(2, 5).Formula = Date + Time()
ApExcel.Cells(3, 5).Formula = gsCodUser
ApExcel.Range("B2", "H8").Font.Bold = True

ApExcel.Range("B2", "B3").HorizontalAlignment = xlLeft
ApExcel.Range("B8", "B8").HorizontalAlignment = xlLeft
ApExcel.Range("H2", "E3").HorizontalAlignment = xlRight
ApExcel.Range("B5", "E6").HorizontalAlignment = xlCenter
ApExcel.Range("B9", "D9").VerticalAlignment = xlCenter
ApExcel.Range("B9", "D9").Borders.LineStyle = 1

ApExcel.Cells(5, 2).Formula = "LISTADO DE PERSONA PARA ACTUALIZAR DATOS"

ApExcel.Range("B5", "E5").MergeCells = True
ApExcel.Range("B6", "E6").MergeCells = True

ApExcel.Cells(9, 2).Formula = "ITEM"
ApExcel.Cells(9, 3).Formula = "COD. PERSONA"
ApExcel.Cells(9, 4).Formula = "CLIENTE"

ApExcel.Range("B9", "D9").Interior.Color = RGB(10, 190, 160)
ApExcel.Range("B9", "D9").Font.Bold = True
ApExcel.Range("B9", "D9").HorizontalAlignment = 3
Dim i As Integer
Dim a As Integer
Dim Count As Integer
i = 10
Count = 0

For a = 1 To UBound(pArrayPersona)
 Count = Count + 1

 ApExcel.Cells(i, 2).Formula = Count
 ApExcel.Cells(i, 3).Formula = Str(pArrayPersona(Count, 1))
 ApExcel.Cells(i, 4).Formula = pArrayPersona(Count, 2)

 ApExcel.Range("C" & Trim(Str(i)) & ":" & "C" & Trim(Str(i))).NumberFormat = "#"
 ApExcel.Range("B" & Trim(Str(i)) & ":" & "D" & Trim(Str(i))).Borders.LineStyle = 1

 i = i + 1
Next a

ApExcel.Cells.Select
ApExcel.Cells.EntireColumn.AutoFit
ApExcel.Columns("B:B").ColumnWidth = 6#
ApExcel.Range("B2").Select

ApExcel.Visible = True
Set ApExcel = Nothing

End Function
'******END APRI*****************************
'***********APRI2018 ERS036-2017
Public Sub GeneraEstadoCuentaAhorros(ByVal FEDatos As Variant, ByVal sPeriodo As String)
    Dim oDoc As New cPDF
    Dim rs As New ADODB.Recordset
    Dim R As New ADODB.Recordset
    Dim nIndex As Integer
    Dim Contador As Integer
    Dim nCentrar As Integer
    Dim nTamTit As Integer
    Dim nTamSubTit As Integer
    Dim nTamLet As Integer
    Dim nTamLet2 As Integer
    Dim sCiudadEmite As String
    Dim nCantDoc As Integer
    Dim sParrafo1 As String
    Dim sParrafo2 As String
    Dim sParrafo21 As String
    Dim sParrafo22 As String
    Dim sParrafo23 As String
    Dim sParrafo3 As String


    nCantDoc = 1
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Operaciones"
    oDoc.Producer = gsNomCmac
    oDoc.Subject = "EMISIÓN DE ESTADO DE CUENTA"
    oDoc.Title = "EMISIÓN DE ESTADO DE CUENTA"

    If Not oDoc.PDFCreate(App.Path & "\Spooler\EstadoCuentaAhorro_" & Trim(FEDatos.TextMatrix(nIndex, 3)) & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then Exit Sub

    oDoc.Fonts.Add "F1", "Arial Narrow", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Arial Narrow", TrueType, Bold, WinAnsiEncoding

    oDoc.LoadImageFromFile App.Path & "\Logo_2015.jpg", "Logo"
    nTamTit = 16: nTamSubTit = 15: nTamLet = 11: nTamLet2 = 8: Contador = 0: nCentrar = 80

    Dim FechaIni As Date
    Dim FechaFin As Date
    Dim oConect As COMConecta.DCOMConecta
    Set oConect = New COMConecta.DCOMConecta
    oConect.AbreConexion
    FechaIni = Left(sPeriodo, 4) & "-" & Right(sPeriodo, 2) & "-01" 'CDate(sPeriodo & "01")
    Dim sql As String
    sql = "select dFecFin = CONVERT(DATE,DATEADD(DAY,-1*DAY('" & Format(FechaIni, "YYYY-MM-DD") & "'),DATEADD(M,1,'" & Format(FechaIni, "YYYY-MM-DD") & "')))"
    Set rs = oConect.CargaRecordSet(sql)
    If Not (rs.EOF And rs.BOF) Then
        FechaFin = rs!dFecFin
    End If
    Set rs = Nothing

            Dim oCapta As COMDCaptaGenerales.DCOMCaptaGenerales

    For nIndex = 1 To FEDatos.Rows - 1
        If FEDatos.TextMatrix(nIndex, 1) = "." Then

        oDoc.NewPage A4_Vertical

        Set oCapta = New COMDCaptaGenerales.DCOMCaptaGenerales
        Set rs = oCapta.ObtenerDatosAhorros(Trim(FEDatos.TextMatrix(nIndex, 3)))

        Contador = 0
        oDoc.WImage 110 + Contador, 440, 70, 120, "Logo"
        oDoc.WTextBox 60 + Contador, 25, 20, 250, "ESTADO DE CUENTA", "F2", nTamTit, hLeft, vMiddle, vbBlack ', 1, vbBlack

        oDoc.WTextBox 125 + Contador, 30, 20, 540, "INFORMACIÓN GENERAL", "F2", nTamLet, hLeft, vMiddle, vbWhite ', 1, vbBlack, True
        oDoc.WTextBox 125 + Contador, 25, 20, 540, "", "F1", nTamLet, hCenter, , , 1, vbBlack, 1, vbBlack, True


        oDoc.WTextBox 145 + Contador, 25, 15, 50, "Cliente", "F2", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 145 + Contador, 65, 15, 250, ":   " & FEDatos.TextMatrix(nIndex, 2), "F1", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 160 + Contador, 25, 15, 65, "Dirección", "F2", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 160 + Contador, 65, 15, 250, ":   " & FEDatos.TextMatrix(nIndex, 4), "F1", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 175 + Contador, 25, 15, 85, "Agencia de Apertura", "F2", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 175 + Contador, 115, 15, 250, ":   " & rs!cAgeApertura, "F1", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 190 + Contador, 25, 15, 85, "Tipo Producto", "F2", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 190 + Contador, 85, 15, 250, ":   " & rs!cTpoProducto, "F1", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 205 + Contador, 25, 15, 85, "Saldo Ult. Cierre", "F2", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 205 + Contador, 105, 15, 250, ":   " & Format(rs!nSaldoCierre, gsFormatoNumeroView), "F1", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack


        oDoc.WTextBox 145 + Contador, 390, 15, 50, "Cuenta", "F2", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 145 + Contador, 430, 15, 100, ":   " & FEDatos.TextMatrix(nIndex, 3), "F1", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 160 + Contador, 390, 15, 50, "Periodo ", "F2", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 160 + Contador, 430, 15, 150, ":   " & FechaIni & " Al " & FechaFin, "F1", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 175 + Contador, 390, 15, 50, "Moneda ", "F2", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 175 + Contador, 430, 15, 150, ":   " & rs!cMoneda, "F1", nTamLet, hLeft, vMiddle, vbBlack  ', 1, vbBlack
        oDoc.WTextBox 190 + Contador, 390, 15, 80, "Fecha de Apertura ", "F2", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 190 + Contador, 465, 15, 150, ":   " & Format(rs!dApertura, "dd/MM/yyyy"), "F1", nTamLet, hLeft, vMiddle, vbBlack  ', 1, vbBlack
        oDoc.WTextBox 205 + Contador, 390, 15, 80, "Saldo Actual ", "F2", nTamLet, hLeft, vMiddle, vbBlack ', 1, vbBlack
        oDoc.WTextBox 205 + Contador, 445, 15, 150, ":   " & Format(rs!nSaldoActual, gsFormatoNumeroView), "F1", nTamLet, hLeft, vMiddle, vbBlack  ', 1, vbBlack

        oDoc.WTextBox 230 + Contador, 30, 20, 540, "MOVIMIENTOS DE LA CUENTA", "F2", nTamLet, hLeft, vMiddle, vbWhite ', 1, vbBlack, True
        oDoc.WTextBox 230 + Contador, 25, 20, 540, "", "F1", nTamLet, hCenter, , , 1, vbBlack, 1, vbBlack, True

        oDoc.WTextBox 255 + Contador, 25, 15, 70, "Fecha", "F2", nTamLet2, hCenter, vMiddle, vbWhite, 1, vbBlack, True
        oDoc.WTextBox 255 + Contador, 98, 15, 215, "Tipo de Operación", "F2", nTamLet2, hCenter, vMiddle, vbWhite, 1, vbBlack, True
        oDoc.WTextBox 255 + Contador, 316, 15, 50, "Depósitos", "F2", nTamLet2, hCenter, vMiddle, vbWhite, 1, vbBlack, True
        oDoc.WTextBox 255 + Contador, 369, 15, 50, "Retiros", "F2", nTamLet2, hCenter, vMiddle, vbWhite, 1, vbBlack, True
        oDoc.WTextBox 255 + Contador, 422, 15, 60, "Sald. Contab.", "F2", nTamLet2, hCenter, vMiddle, vbWhite, 1, vbBlack, True
        oDoc.WTextBox 255 + Contador, 485, 15, 80, "Agencia", "F2", nTamLet2, hCenter, vMiddle, vbWhite, 1, vbBlack, True


        Set rs = Nothing
        Set rs = oCapta.GeneraExtractoCta(Trim(FEDatos.TextMatrix(nIndex, 3)), sPeriodo)
        Contador = 15
        Do While Not rs.EOF
        oDoc.WTextBox 255 + Contador, 25, 15, 70, rs!fecha, "F1", nTamLet2, hCenter, vMiddle, vbBlack
        oDoc.WTextBox 255 + Contador, 98, 15, 215, rs!Operacion, "F1", nTamLet2, hLeft, vMiddle, vbBlack
        oDoc.WTextBox 255 + Contador, 316, 15, 50, Format(rs!nabono, gsFormatoNumeroView), "F1", nTamLet2, hRight, vMiddle, vbBlack
        oDoc.WTextBox 255 + Contador, 369, 15, 50, Format(rs!nCargo, gsFormatoNumeroView), "F1", nTamLet2, hRight, vMiddle, vbBlack
        oDoc.WTextBox 255 + Contador, 422, 15, 60, Format(rs!nSaldoContable, gsFormatoNumeroView), "F1", nTamLet2, hRight, vMiddle, vbBlack
        oDoc.WTextBox 255 + Contador, 485, 15, 80, rs!cAgencia, "F1", nTamLet2, hCenter, vMiddle, vbBlack

        Contador = Contador + 10
        rs.MoveNext
            If rs.EOF Then
            Contador = Contador + 255
            End If
        Loop
        Set rs = Nothing

        sParrafo1 = "Estimado Cliente, se le recuerda que si existiera alguna diferencia entre su Saldo Contable y Saldo Actual se puede atribuir a las siguientes razones:"

        sParrafo2 = "Nuestro compromiso es servirle cada día mejor, por lo que si tuviese alguna consulta o desee efectuar un reclamo puede apersonarse a nuestra red de " & _
                    "agencias o llamar al Fono Maynas 065-58 18 00 (Loreto) o 0801-10700 (Lima y provincias) donde gustosamente le estaremos atendiendo. Adicionalmente, " & _
                    "usted podrá recurrir a Indecopi  o a la Plataforma de Atención al Usuario de la Superintendencia de Banca, Seguros y AFP. "
        sParrafo3 = "Le recordamos verificar la información contenida en el estado de cuenta, si tuviese  alguna observación al respecto,  le solicitamos " & _
                    "contactarse con nosotros a través de nuestra red de agencias dentro de los 30 días calendario siguientes a la recepción de este documento. " & _
                    "En caso contrario, daremos por conforme la cuenta y aprobado el saldo."


        sParrafo21 = "Retiros en Cajeros Automáticos"
        sParrafo22 = "Compras POS (Red Visa)"
        sParrafo23 = "Reembolsos VISA NET -  Establecimientos afiliados al servicio banco pagador"


        oDoc.WTextBox Contador + 30, 30, 15, 200, "INFORMACIÓN AL CLIENTE ", "F2", nTamLet2, hLeft, vMiddle, vbBlack ', 1, vbBlack
        If Mid(FEDatos.TextMatrix(nIndex, 3), 6, 3) = "232" Then
            oDoc.WTextBox Contador + 45, 30, 15, 535, sParrafo1, "F1", nTamLet2, hjustify, vMiddle, vbBlack ', 1, vbBlack
            oDoc.WTextBox Contador + 60, 30, 15, 535, sParrafo21, "F1", nTamLet2, hjustify, vMiddle, vbBlack ', 1, vbBlack
            oDoc.WTextBox Contador + 70, 30, 15, 535, sParrafo22, "F1", nTamLet2, hjustify, vMiddle, vbBlack ', 1, vbBlack
            oDoc.WTextBox Contador + 80, 30, 15, 535, sParrafo23, "F1", nTamLet2, hjustify, vMiddle, vbBlack ', 1, vbBlack
            oDoc.WTextBox Contador + 95, 30, 40, 535, sParrafo2, "F1", nTamLet2, hjustify, vMiddle, vbBlack ', 1, vbBlack
            oDoc.WTextBox Contador + 125, 30, 40, 535, sParrafo3, "F1", nTamLet2, hjustify, vMiddle, vbBlack ', 1, vbBlack

            oDoc.WTextBox Contador + 25, 25, 140, 540, "", "F1", nTamLet, hCenter, , , 1, vbBlack

        ElseIf Mid(FEDatos.TextMatrix(nIndex, 3), 6, 3) = "232" Then
            oDoc.WTextBox Contador + 45, 30, 15, 535, sParrafo1, "F1", nTamLet2, hjustify, vMiddle, vbBlack ', 1, vbBlack
            oDoc.WTextBox Contador + 60, 30, 15, 535, sParrafo21, "F1", nTamLet2, hjustify, vMiddle, vbBlack ', 1, vbBlack
            oDoc.WTextBox Contador + 70, 30, 15, 535, sParrafo22, "F1", nTamLet2, hjustify, vMiddle, vbBlack ', 1, vbBlack
            oDoc.WTextBox Contador + 85, 30, 40, 535, sParrafo2, "F1", nTamLet2, hjustify, vMiddle, vbBlack ', 1, vbBlack
            oDoc.WTextBox Contador + 115, 30, 40, 535, sParrafo3, "F1", nTamLet2, hjustify, vMiddle, vbBlack ', 1, vbBlack

            oDoc.WTextBox Contador + 25, 25, 130, 540, "", "F1", nTamLet, hCenter, , , 1, vbBlack

        Else
            oDoc.WTextBox Contador + 45, 30, 40, 535, sParrafo2, "F1", nTamLet2, hjustify, vMiddle, vbBlack ', 1, vbBlack
            oDoc.WTextBox Contador + 75, 30, 40, 535, sParrafo3, "F1", nTamLet2, hjustify, vMiddle, vbBlack ', 1, vbBlack

            oDoc.WTextBox Contador + 25, 25, 90, 540, "", "F1", nTamLet, hCenter, , , 1, vbBlack

        End If




         End If

        Next





    oDoc.PDFClose
    oDoc.Show
End Sub
'***********END APRI************
'INICIO EAAS20190523 Memorándum Nº 756-2019-GM-DI/CMACM
Public Sub AhorroApertura_ContratosAutomaticos(MatTitular As Variant, ByVal psCtaCod As String)
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

    Set loAge = New COMDConstantes.DCOMAgencias
    Set rs1 = New ADODB.Recordset
        Set rs1 = loAge.RecuperaAgencias(gsCodAge)
        If Not (rs1.EOF And rs1.BOF) Then
            lsAgencia = Trim(rs1("Dist"))
        End If
    Set loAge = Nothing

    Set oWord = CreateObject("Word.Application")
        oWord.Visible = False

    Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\cpasivosAhorros" & gsCodAge & ".doc")

    sArchivo = App.Path & "\FormatoCarta\ICA_" & psCtaCod & "_" & Replace(Left(Time, 5), ":", "") & ".doc"
    oDoc.SaveAs (sArchivo)
    Dim i As Integer
    Dim rsDir As ADODB.Recordset
    Dim oPers As COMDpersona.DCOMPersona

    If UBound(MatTitular) = 2 Then

        lsNom1 = MatTitular(1, 1)
        lsDoc1 = MatTitular(1, 2)
        lsDir1 = MatTitular(1, 3)
        With oWord.Selection.Find
            .Text = "<<NomTit1>>"
            .Replacement.Text = Trim(lsNom1)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit1>>"
            .Replacement.Text = Trim(lsDoc1)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit1>>"
            .Replacement.Text = Trim(lsDir1)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit2>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit2>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit2>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit3>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit3>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit3>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With

    ElseIf UBound(MatTitular) = 3 Then
        lsNom1 = MatTitular(1, 1)
        lsDoc1 = MatTitular(1, 2)
        lsDir1 = MatTitular(1, 3)
        lsNom2 = MatTitular(2, 1)
        lsDoc2 = MatTitular(2, 2)
        lsDir2 = MatTitular(2, 3)


        With oWord.Selection.Find
            .Text = "<<NomTit1>>"
            .Replacement.Text = Trim(lsNom1)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit1>>"
            .Replacement.Text = Trim(lsDoc1)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit1>>"
            .Replacement.Text = Trim(lsDir1)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit2>>"
            .Replacement.Text = Trim(lsNom2)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit2>>"
            .Replacement.Text = Trim(lsDoc2)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit2>>"
            .Replacement.Text = Trim(lsDir2)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit3>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit3>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit3>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
    ElseIf UBound(MatTitular) = 4 Then
        lsNom1 = MatTitular(1, 1)
        lsDoc1 = MatTitular(1, 2)
        lsDir1 = MatTitular(1, 3)

        lsNom2 = MatTitular(2, 1)
        lsDoc2 = MatTitular(2, 2)
        lsDir2 = MatTitular(2, 3)

        lsNom3 = MatTitular(3, 1)
        lsDoc3 = MatTitular(3, 2)
        lsDir3 = MatTitular(3, 3)

        With oWord.Selection.Find
            .Text = "<<NomTit1>>"
            .Replacement.Text = Trim(lsNom1)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit1>>"
            .Replacement.Text = Trim(lsDoc1)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit1>>"
            .Replacement.Text = Trim(lsDir1)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit2>>"
            .Replacement.Text = Trim(lsNom2)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit2>>"
            .Replacement.Text = Trim(lsDoc2)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit2>>"
            .Replacement.Text = Trim(lsDir2)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit3>>"
            .Replacement.Text = Trim(lsNom3)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit3>>"
            .Replacement.Text = Trim(lsDoc3)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit3>>"
            .Replacement.Text = Trim(lsDir3)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit4>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With

    ElseIf UBound(MatTitular) = 5 Then
        lsNom1 = MatTitular(1, 1)
        lsDoc1 = MatTitular(1, 2)
        lsDir1 = MatTitular(1, 3)

        lsNom2 = MatTitular(2, 1)
        lsDoc2 = MatTitular(2, 2)
        lsDir2 = MatTitular(2, 3)

        lsNom3 = MatTitular(3, 1)
        lsDoc3 = MatTitular(3, 2)
        lsDir3 = MatTitular(3, 3)

        lsNom4 = MatTitular(4, 1)
        lsDoc4 = MatTitular(4, 2)
        lsDir4 = MatTitular(4, 3)


        With oWord.Selection.Find
            .Text = "<<NomTit1>>"
            .Replacement.Text = Trim(lsNom1)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit1>>"
            .Replacement.Text = Trim(lsDoc1)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit1>>"
            .Replacement.Text = Trim(lsDir1)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit2>>"
            .Replacement.Text = Trim(lsNom2)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit2>>"
            .Replacement.Text = Trim(lsDoc2)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit2>>"
            .Replacement.Text = Trim(lsDir2)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit3>>"
            .Replacement.Text = Trim(lsNom3)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit3>>"
            .Replacement.Text = Trim(lsDoc3)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit3>>"
            .Replacement.Text = Trim(lsDir3)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<NomTit4>>"
            .Replacement.Text = Trim(lsNom4)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DocTit4>>"
            .Replacement.Text = Trim(lsDoc4)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DirTit4>>"
            .Replacement.Text = Trim(lsDir4)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
    End If

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
        .Replacement.Text = Format(gdFecSis, "mm")
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
'   Set oDoc = Nothing
'
'    Set oWord = CreateObject("Word.Application")
'    oWord.Visible = True
'    Set oDoc = oWord.Documents.Open(sArchivo)
'    Set oDoc = Nothing
'    Set oWord = Nothing
'wAppSource.Quit
oWord.Application.Quit
'wApp.Visible = True

'MEJORAS JHCU
 Dim doc As New Word.Application
 With doc
        .Documents.Open sArchivo 'abrimos "Mi documento"
        .Visible = True 'hacemos visible Word
 End With
'FIN MEJORAS JHCU
End Sub

Public Function AhorroApertura_ContratosAutomaticosLote(MatTitular As Variant, MatNroCta() As String, psMovNro As String)

    Dim rs As ADODB.Recordset
    Dim nParCod  As Long
    Dim lsFechas As String
    Dim lsNom1 As String
    Dim lsDoc1 As String
    Dim lsDir1 As String
    Dim lsCad As String
    Dim lsModeloPlantilla As String
    Dim i As Integer

    Dim loRs As COMDConstSistema.DCOMGeneral
    Set loRs = New COMDConstSistema.DCOMGeneral


    Dim loAge As COMDConstantes.DCOMAgencias
    Dim rs1 As ADODB.Recordset
    Dim lsAgencia As String
    Dim lsAgeDir As String

    Set loAge = New COMDConstantes.DCOMAgencias
    Set rs1 = New ADODB.Recordset
        Set rs1 = loAge.RecuperaAgencias(gsCodAge)
        If Not (rs1.EOF And rs1.BOF) Then
            lsAgencia = Trim(rs1("Dist"))
        End If
    Set loAge = Nothing


    Dim objCaptac As COMNCaptaGenerales.NCOMCaptaGenerales
    Set objCaptac = New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim objCaptaGenerales As COMDCaptaGenerales.DCOMCaptaGenerales
    Set objCaptaGenerales = New COMDCaptaGenerales.DCOMCaptaGenerales
    Dim sListaCtas As String
    Dim lnMonto As Currency
    Dim lnMontoSald As Currency
    Dim lnTasaInteres As Currency
    Dim nI As Integer
    Dim rsCtas As ADODB.Recordset
    sListaCtas = ""
    sListaCtas = "'0"
    Dim lsFecha As String
    Dim nPag As Integer
    Dim nDoc As Integer
    Dim lsArchivo As String
    Dim lsNomMaq As String
    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range
    Dim sArchivo As String
    Dim lsNom2, lsNom3, lsNom4 As String
    Dim lsDoc2, lsDoc3, lsDoc4 As String
    Dim lsDir2, lsDir3, lsDir4 As String

    Set loAge = New COMDConstantes.DCOMAgencias
    Set rs1 = New ADODB.Recordset
        Set rs1 = loAge.RecuperaAgencias(gsCodAge)
        If Not (rs1.EOF And rs1.BOF) Then
            lsAgencia = Trim(rs1("Dist"))
        End If
    Set loAge = Nothing



    For nI = 1 To UBound(MatNroCta(), 1)
        sListaCtas = sListaCtas & "," & MatNroCta(nI)
    Next nI

        sListaCtas = sListaCtas & "'"
        Set rsCtas = objCaptaGenerales.ObtenerCtaPorPersonas(MatTitular(i, 9), sListaCtas)
        If Not (rsCtas.BOF And rsCtas.EOF) Then
            sListaCtas = rsCtas!cCtaCod
        End If
    lsModeloPlantilla = App.Path & "\FormatoCarta\cpasivosAhorros" & gsCodAge & ".doc"
    'Crea una clase que de Word Object
    Dim wApp As Word.Application
    Dim wAppSource As Word.Application
    'Create a new instance of word
    Set wApp = New Word.Application
    Set wAppSource = New Word.Application

    Dim RangeSource As Word.Range
    'Abre Documento Plantilla
    wAppSource.Documents.Open FileName:=lsModeloPlantilla
    Set RangeSource = wAppSource.ActiveDocument.Content
    'Lo carga en Memoria
    wAppSource.ActiveDocument.Content.Copy

    'Crea Nuevo Documento
    wApp.Documents.Add

    With wApp.ActiveDocument.PageSetup
        .HeaderDistance = CentimetersToPoints(0)
        .FooterDistance = CentimetersToPoints(0)
        .LeftMargin = CentimetersToPoints(1.8)
        .RightMargin = CentimetersToPoints(1.34) 'EAAS20190612 SE CAMBIO EL VALOR DE MARGEN
        .TopMargin = CentimetersToPoints(0)
        .BottomMargin = CentimetersToPoints(1)
    End With

    With wApp.Selection.Font
        .Name = "Arial Narrow"
        .Size = 11 'EAAS20190612 SE CAMBIO EL VALOR a 11
    End With

    For i = 1 To UBound(MatTitular)
        Set rsCtas = objCaptaGenerales.ObtenerCtaPorPersonas(MatTitular(i, 9), sListaCtas)
        If Not (rsCtas.BOF And rsCtas.EOF) Then
            lnMontoSald = MatTitular(i, 3)
            lnTasaInteres = MatTitular(i, 2)
        End If
        '***********************************************
        wApp.Application.Selection.TypeParagraph
        wApp.Application.Selection.PasteAndFormat (wdPasteDefault) 'EAAS20190612 SE CAMBIO EL VALOR wdPasteDefault
        wApp.Application.Selection.InsertBreak

        wApp.Selection.SetRange start:=wApp.Selection.start, End:=wApp.ActiveDocument.Content.End

    wApp.Selection.MoveEnd
    lsNom1 = MatTitular(i, 1)
    lsDoc1 = MatTitular(i, 4)
    lsDir1 = MatTitular(i, 5)

    With wApp.Selection.Find
        .Text = "<<NomTit1>>"
        .Replacement.Text = lsNom1
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DocTit1>>"
        .Replacement.Text = lsDoc1
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DirTit1>>"
        .Replacement.Text = lsDir1
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<NomTit2>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DocTit2>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DirTit2>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<NomTit3>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DocTit3>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DirTit3>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<NomTit4>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DocTit4>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<DirTit4>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<Zona>>"
        .Replacement.Text = lsAgencia
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<fDay>>"
        .Replacement.Text = Format(gdFecSis, "dd")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False

        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<fMes>>"
        .Replacement.Text = Format(gdFecSis, "mm")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With wApp.Selection.Find
        .Text = "<<fAnio>>"
        .Replacement.Text = Format(gdFecSis, "yyyy")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

  Next

'modifcaciones jhcu mejoras cartillas 10-06-2019
wAppSource.ActiveDocument.Close
wAppSource.Application.Quit
wApp.ActiveDocument.SaveAs (App.Path & "\SPOOLER\AperturaLoteAhorro_Automatico" & psMovNro & ".doc")
wApp.Application.Quit
'wApp.Visible = True

'MEJORAS JHCU
 Dim doc As New Word.Application
 With doc
        .Documents.Open App.Path & "\SPOOLER\AperturaLoteAhorro_Automatico" & psMovNro & ".doc" 'abrimos "Mi documento"
        .Visible = True 'hacemos visible Word
 End With
'FIN MEJORAS JHCU
End Function
'END EAAS20190523 SEGUN Memorándum Nº 756-2019-GM-DI/CMACM
'APRI20190109 ERS077-2018
Public Sub ImpreCartaNotificacionTarifario(ByVal psPersCod As String, Optional ByVal psCtaCod As String, Optional ByVal pdFecSis)
    Dim rs As ADODB.Recordset
    Dim lsFechas As String
    Dim lsFechaCom As String
    Dim dFechCom As Date
    Dim sPersCod As String
    Dim sNombreCliente As String
    Dim sNombreCliente2 As String
    Dim sNombreCliente3 As String
    Dim sNombreCliente4 As String
    Dim sDireccion As String
    Dim sUbigeo As String
    Dim nPeriodo As Integer

    Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
    Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales

    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range

    Dim loRs As COMDConstSistema.DCOMGeneral
    Set loRs = New COMDConstSistema.DCOMGeneral


    Set oWord = CreateObject("Word.Application")
    oWord.Visible = True

    Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\CARTANOTIFICACIONTARIFARIO.doc")
    '**********************APRI20191218**************************'
    oDoc.SaveAs App.Path & "\SPOOLER\CARTACOMUNICACION-" & psCtaCod & ".doc"
    '**********************APRI**************************'
    nPeriodo = loRs.LeeConstSistema(192)
    dFechCom = DateAdd("d", nPeriodo, gdFecSis)
    dFechCom = DateSerial(Year(dFechCom), Month(dFechCom) + 1, 0)
    dFechCom = DateAdd("d", 1, dFechCom)
    lsFechaCom = Format(dFechCom, "dd") & " de " & Format(dFechCom, "mmmm") & " de " & Format(dFechCom, "yyyy")

    'Fecha
    lsFechas = Format(gdFecSis, "dd") & " de " & Format(gdFecSis, "mmmm") & " de " & Format(gdFecSis, "yyyy")

    With oWord.Selection.Find
        .Text = "<<dFechaActual>>"
        .Replacement.Text = lsFechas
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With


     With oWord.Selection.Find
        .Text = "<<dFechaComunicaion>>"
        .Replacement.Text = lsFechaCom
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With



    Dim i As Integer
    Dim rsPers As ADODB.Recordset
    Dim oPers As COMDpersona.DCOMPersona
    Set oPers = New COMDpersona.DCOMPersona
    Set rsPers = oPers.RecuperaDatosPersonaComunicacion(psCtaCod, psPersCod)
    Set oPers = Nothing
    i = 0


    Do While Not rsPers.EOF


        sPersCod = rsPers!cPersCod
        sNombreCliente = rsPers!cPersNombre
        sDireccion = Trim(rsPers!cPersDireccDomicilio)
        sUbigeo = Trim(rsPers!cDepartamento) & "/" & Trim(rsPers!cProvincia) & "/" & Trim(rsPers!cDistrito)

        With oWord.Selection.Find
            .Text = "<<nNombreCliente1>>"
            .Replacement.Text = sNombreCliente
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With

        If i = 1 Then
        sNombreCliente2 = rsPers!cPersNombre
        ElseIf i = 2 Then
        sNombreCliente3 = rsPers!cPersNombre
        ElseIf i = 3 Then
        sNombreCliente4 = rsPers!cPersNombre
        End If

        With oWord.Selection.Find
            .Text = "<<nNombreCliente2>>"
            .Replacement.Text = sNombreCliente2
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With

         With oWord.Selection.Find
            .Text = "<<nNombreCliente3>>"
            .Replacement.Text = sNombreCliente3
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With

        With oWord.Selection.Find
            .Text = "<<nNombreCliente4>>"
            .Replacement.Text = sNombreCliente4
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With

        If i = 0 Then
            With oWord.Selection.Find
                .Text = "<<cDireccion>>"
                .Replacement.Text = sDireccion
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With oWord.Selection.Find
                .Text = "<<cUbigeo>>"
                .Replacement.Text = sUbigeo
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
        End If

        oCap.ActualizarFechaComunicacion (sPersCod)
        i = i + 1
        rsPers.MoveNext
    Loop
    Set oCap = Nothing
    'MEJORAS APRI20191218
    oDoc.Save
    oDoc.Application.Quit

    Dim doc As New Word.Application
    With doc
        .Documents.Open App.Path & "\SPOOLER\CARTACOMUNICACION-" & psCtaCod & ".doc" 'abrimos "Mi documento"
        .Visible = True 'hacemos visible Word
    End With
   'FIN MEJORAS APRI
   'oDoc.SaveAs App.Path & "\SPOOLER\CARTACOMUNICACION-" & psCtaCod & ".doc"
End Sub
'END APRI
