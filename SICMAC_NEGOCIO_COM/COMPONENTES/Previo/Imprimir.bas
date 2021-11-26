Attribute VB_Name = "Imprimir"
Option Explicit
Public ArcSal As Integer
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
    
    Global Const SWP_NOMOVE = 2
    Global Const SWP_NOSIZE = 1
    Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    Global Const HWND_TOPMOST = -1
    Global Const HWND_NOTOPMOST = -2
    
'***************************************************
'* Inicia una impresión - Cabecera
'***************************************************
'FECHA CREACION : 11/07/99  -   MAVF
'MODIFICACION:
'Referencia : Global ArcSal As Integer
'***************************************************
Public Sub ImpreBegin(pbCondensado As Boolean, nLineas As Integer)
    ArcSal = FreeFile
    
    Open sLpt For Output As ArcSal
    Print #ArcSal, oImpresora.gPrnInicializa;            'Inicializa Impresora
    If pbCondensado Then
       Print #ArcSal, oImpresora.gPrnMargenIzq00; 'Tipo letra : 0,1,2 - Roman,SansS,Courier
       Print #ArcSal, oImpresora.gPrnMargenIzq00;            'Tamaño  : 80, 77, 103
       Print #ArcSal, oImpresora.gPrnCondensadaON;                       'Retorna al tipo de letra normal
    Else
       Print #ArcSal, oImpresora.gPrnCondensadaOFF;
    End If
    Print #ArcSal, oImpresora.gPrnEspaLineaN;                     'Espaciamiento lineas 1/6 pulg.1
    Print #ArcSal, oImpresora.gPrnTamPaginaCab & Chr$(nLineas);  '   Chr$(nLineas); 'Longitud de página a 66 líneas
    If Not pbCondensado Then
       Print #ArcSal, oImpresora.gPrnTpoLetraCurier;        'Tipo letra : 0,1,2 - Roman,SansS,Courier
       Print #ArcSal, oImpresora.gPrnTamLetra10CPI;                   'Tamaño  : 80, 77, 103
    End If
    Print #ArcSal, oImpresora.gPrnTpoLetraRoman1P;          'Draf : 1 pasada
   
End Sub
'***************************************************
'* Termina un impresión - Cola
'***************************************************
'FECHA CREACION : 11/07/99  -   MAVF
'MODIFICACION:
'Referencia : Global ArcSal As Integer
'***************************************************
Public Sub ImpreEnd()
    Print #ArcSal, oImpresora.gPrnSaltoPagina;   'Nueva página
    Print #ArcSal, oImpresora.gPrnCondensadaOFF;   'Retorna al tipo de letra normal
    Close ArcSal
End Sub
'***************************************************
'* Genera nueva página
'***************************************************
'FECHA CREACION : 11/07/99  -   MAVF
'MODIFICACION:
'Referencia : Global ArcSal As Integer
'***************************************************
Public Sub ImpreNewPage()
    Print #ArcSal, oImpresora.gPrnSaltoPagina;   'Nueva página
End Sub
'Prepara una cadena especial (cadena con caracteres con tilde y/o otros)
' para que se imprima en el modo FREEFILE.
Public Function ImpreCarEsp(ByVal vCadena As String) As String
    vCadena = Replace(vCadena, "á", Chr(160), , , vbTextCompare)
    vCadena = Replace(vCadena, "é", Chr(130), , , vbTextCompare)
    vCadena = Replace(vCadena, "í", Chr(161), , , vbTextCompare)
    vCadena = Replace(vCadena, "ó", Chr(162), , , vbTextCompare)
    vCadena = Replace(vCadena, "ú", Chr(163), , , vbTextCompare)
    vCadena = Replace(vCadena, "ñ", Chr(164), , , vbTextCompare)
    vCadena = Replace(vCadena, "Ñ", Chr(165), , , vbTextCompare)
    vCadena = Replace(vCadena, "°", Chr(248), , , vbTextCompare)
    vCadena = Replace(vCadena, "¦", Chr(179), , , vbTextCompare)
    ImpreCarEsp = vCadena
End Function
'Verifica la corrceta habilitación de la impresora
Public Function ImpreSensa() As Boolean
Dim lbArchAbierto As Boolean
On Error GoTo ControlError
    'frmMdiMain.staMain.Panels(2).Text = "Verificando Conexión con Impresora"
    ArcSal = FreeFile
    lbArchAbierto = True
    Open sLpt For Output As ArcSal
    Print #ArcSal, oImpresora.gPrnInicializa;             'Inicializa Impresora
    Close ArcSal
    lbArchAbierto = False
    ImpreSensa = True
    'frmMdiMain.staMain.Panels(2).Text = ""
    Exit Function
ControlError:   ' Rutina de control de errores.
 '   frmMdiMain.staMain.Panels(2).Text = ""
    If lbArchAbierto Then
        Close ArcSal
    End If
    MsgBox "Impresora no Encontrada ó Inactiva" & vbCr & "Verifique que la Conexión sea Correcta", vbExclamation, "Aviso de Precaución"
    ImpreSensa = False
End Function
Public Function PrnSet(Code As String, Optional nValor As Integer) As String
If nValor = 12 Or nValor = 10 Then
   nValor = nValor - 1
End If
Select Case Code
 Case "B+": PrnSet = oImpresora.gPrnBoldON 'Bold On
 Case "B-": PrnSet = oImpresora.gPrnBoldOFF 'Bold Off
 Case "U+": PrnSet = oImpresora.gPrnUnderLineONOFF  'Underline On
 Case "U-": PrnSet = oImpresora.gPrnUnderLineONOFF 'Chr$(27) & Chr$(46) 'Underline Off
 Case "I+": PrnSet = oImpresora.gPrnItalicON 'Italic On
 Case "I-": PrnSet = oImpresora.gPrnItalicOFF 'Italic Off
 Case "W+": PrnSet = oImpresora.gPrnDblAnchoON 'Doble Ancho On
 Case "W-": PrnSet = oImpresora.gPrnDblAnchoOFF 'Doble Ancho Off
 Case "C+": PrnSet = oImpresora.gPrnCondensadaON 'Condensado On
 Case "C-": PrnSet = oImpresora.gPrnCondensadaOFF 'Condensado Off
 Case "Rm": PrnSet = oImpresora.gPrnTpoLetraRoman 'Roman
 Case "Ss": PrnSet = oImpresora.gPrnTpoLetraSansSerif 'Sans Serif
 Case "Co": PrnSet = oImpresora.gPrnTpoLetraCurier 'Courier
 Case "1.5": PrnSet = oImpresora.gPrnUnoMedioEspacio ' 1 1/2 espacios
 Case "MI": PrnSet = oImpresora.gPrnMargenIzqCab 'Margen Izquierdo
 Case "MD": PrnSet = oImpresora.gPrnMargenDerCab & Chr$(nValor)   'Margen Derecho
 Case "10CPI": PrnSet = oImpresora.gPrnTamLetra12CPI
 Case "12CPI": PrnSet = oImpresora.gPrnTamLetra10CPI
 Case "15CPI": PrnSet = oImpresora.gPrnTamLetra15CPI
 Case "EspN": PrnSet = oImpresora.gPrnEspaLineaN     'Espaciado Normal 4.5/72
 Case "Esp":  PrnSet = oImpresora.gPrnEspaLineaValor & Chr$(nValor)  'Espaciado nValor/72 pulg.
End Select
End Function

Public Function ImpreFormat(ByVal pNumero As Variant, ByVal pLongitudEntera As Integer, _
        Optional ByVal pLongitudDecimal As Integer = 2, _
        Optional ByVal pMoneda As Boolean = False) As String
Dim vPosPto As Integer
Dim vParEnt As String
Dim vParDec As String
Dim vLonEnt As Integer
Dim vLonDec As Integer
Dim x As Integer

On Error GoTo ErrHandler
vParDec = ""
If IsNull(pNumero) Then
    If pLongitudDecimal > 0 Then vParDec = "." & String(pLongitudDecimal, "0")
    If pLongitudEntera <= 0 Then pLongitudEntera = 1
    ImpreFormat = String(pLongitudEntera - 1, " ") & "0" & vParDec
ElseIf VarType(pNumero) = 8 Then
    pNumero = Trim(pNumero)
    vLonEnt = Len(pNumero)
    If vLonEnt > pLongitudEntera Then
        pNumero = Left(pNumero, pLongitudEntera)
        vLonEnt = pLongitudEntera
    End If
    ImpreFormat = String(pLongitudDecimal, " ") & pNumero & String(pLongitudEntera - vLonEnt, " ")
Else
    vPosPto = InStr(Trim(CStr(pNumero)), ".")
    If vPosPto > 0 Then
        vParEnt = Trim(CStr(Left(pNumero, vPosPto - 1)))
        vParDec = Trim(CStr(Mid(pNumero, vPosPto + 1)))
        vLonEnt = Len(vParEnt)
        vLonDec = Len(vParDec)
    Else
        vParEnt = Trim(Str(pNumero))
        vParDec = ""
        vLonEnt = Len(vParEnt)
        vLonDec = 0
    End If
    If pMoneda And vLonEnt > 3 Then
        vParEnt = Format(vParEnt, "#,###,###")
        For x = 1 To Len(vParEnt)
            If Mid(vParEnt, x, 1) = "," Then pLongitudEntera = pLongitudEntera - 1
        Next x
    End If
    If vLonEnt > pLongitudEntera Then pLongitudEntera = vLonEnt + 1
    If vLonDec > pLongitudDecimal Then
        vLonDec = pLongitudDecimal
        vParDec = Left(vParDec, vLonDec)
    End If
    ImpreFormat = String(pLongitudEntera - vLonEnt, " ") & vParEnt
    If pLongitudDecimal > 0 Then
        ImpreFormat = ImpreFormat & "." & vParDec & String(pLongitudDecimal - vLonDec, "0")
    End If
End If
Exit Function

ErrHandler:     ' Errores obtenidos
    MsgBox " Operación no válida " & vbCr & _
        " Error " & Err.Number & " : " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " ! Aviso ! "
End Function


