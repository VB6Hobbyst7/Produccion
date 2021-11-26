VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRepResCVenta 
   Caption         =   "Resumen Diario de Compra Venta de Dólares"
   ClientHeight    =   4530
   ClientLeft      =   4950
   ClientTop       =   4065
   ClientWidth     =   4905
   Icon            =   "frmRepResCVenta.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkConsol 
      Caption         =   "&Consolidar Fechas"
      Height          =   225
      Left            =   120
      TabIndex        =   10
      Top             =   4110
      Width           =   1665
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   3720
      TabIndex        =   4
      Top             =   4140
      Width           =   1065
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   345
      Left            =   2610
      TabIndex        =   3
      Top             =   4140
      Width           =   1065
   End
   Begin VB.Frame fraAge 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   3255
      Left            =   120
      TabIndex        =   5
      Top             =   60
      Width           =   4665
      Begin VB.ListBox lstAge 
         Height          =   2535
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   600
         Width           =   4425
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A G E N C I A S"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   405
         Left            =   120
         TabIndex        =   9
         Top             =   210
         Width           =   4425
      End
   End
   Begin VB.Frame fraFechas 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   3285
      Width           =   4665
      Begin MSMask.MaskEdBox mskFecFin 
         Height          =   300
         Left            =   3240
         TabIndex        =   2
         Top             =   285
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFecIni 
         Height          =   300
         Left            =   1560
         TabIndex        =   1
         Top             =   270
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFecFin 
         Caption         =   "Al"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   2940
         TabIndex        =   8
         Top             =   330
         Width           =   285
      End
      Begin VB.Label lblFecIni 
         Caption         =   "Fechas  :   Del "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   315
         Width           =   1425
      End
   End
End
Attribute VB_Name = "frmRepResCVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs   As ADODB.Recordset
Dim sSql As String
Dim lsOpeCompra As String
Dim lsOpeVenta  As String
Dim lsOpeCompraFI As String 'MIOL 20121022, SEGUN RQ12338
Dim lsOpeVentaFI  As String 'MIOL 20121022, SEGUN RQ12338
Dim oBarra As clsProgressBar
'MIOL 20121022, SEGUN RQ12338****************************************
Private Function GeneraRCompraVenta(psOpeCompra As String, psOpeVenta As String, Optional psOpeCompraFI As String, Optional psOpeVentaFI As String) As String
'Private Function GeneraRCompraVenta(psOpeCompra As String, psOpeVenta As String) As String
    Dim i As Integer
    Dim lsCadena As String
    Dim lsAgeNom As String
    Dim j As Integer
    
    If ValidaD Then Exit Function
    Set oBarra = New clsProgressBar
    oBarra.ShowForm Me
    oBarra.CaptionSyle = eCap_CaptionPercent
    oBarra.Max = lstAge.ListCount
    lsCadena = ""
   For i = 0 To lstAge.ListCount - 1
        lsAgeNom = Trim(lstAge.List(i))
        oBarra.Progress i + 1, "CONSOLIDADO DE COMPRA-VENTA", lsAgeNom, "Procesando..."
         If lstAge.Selected(i) Then
            If Left(lstAge.List(i), 6) = "CONSOL" Then
               If chkConsol.value = vbChecked Then
                 
                 '*** PEAC 20120823
                  'lsCadena = lsCadena & GetMovCompraVentaME(psOpeCompra, psOpeVenta, CDate(Me.mskFecIni), CDate(Me.mskFecFin))
                  
                  For j = 0 To DateDiff("d", CDate(Me.mskFecIni), CDate(Me.mskFecFin))
                    oBarra.Progress i + 1, "CONSOLIDADO DE COMPRA-VENTA", lsAgeNom, "Procesando día " & DateAdd("d", j, CDate(Me.mskFecIni)) & "..."
                    DoEvents
                    If j Mod 5 = 0 And j <> 0 Then lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                        If CDate(Me.mskFecIni) <> DateAdd("d", j, CDate(Me.mskFecIni)) Then
                            lsCadena = lsCadena & GetMovCompraVentaME(psOpeCompra, psOpeVenta, CDate(Me.mskFecIni), DateAdd("d", j, CDate(Me.mskFecIni)), , , psOpeCompraFI, psOpeVentaFI)
                            'lsCadena = lsCadena & GetMovCompraVentaME(psOpeCompra, psOpeVenta, CDate(Me.mskFecIni), DateAdd("d", j, CDate(Me.mskFecIni)))
                        End If
                  Next j
                  '*** FIN PEAC
               Else
                  For j = 0 To DateDiff("d", CDate(Me.mskFecIni), CDate(Me.mskFecFin))
                    oBarra.Progress i + 1, "CONSOLIDADO DE COMPRA-VENTA", lsAgeNom, "Procesando día " & DateAdd("d", j, CDate(Me.mskFecIni)) & "..."
                    DoEvents
                    If j Mod 5 = 0 And j <> 0 Then lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                    'MIOL ***
                    lsCadena = lsCadena & GetMovCompraVentaME(psOpeCompra, psOpeVenta, DateAdd("d", j, CDate(Me.mskFecIni)), DateAdd("d", j, CDate(Me.mskFecIni)), , , psOpeCompraFI, psOpeVentaFI)
                    'lsCadena = lsCadena & GetMovCompraVentaME(psOpeCompra, psOpeVenta, DateAdd("d", j, CDate(Me.mskFecIni)), DateAdd("d", j, CDate(Me.mskFecIni)))
                    'END MIOL
                  Next j
               End If
            'MIOL 20121022, SEGUN RQ12338****************************************
            ElseIf Left(lstAge.List(i), 2) = "OF" Then
               psOpeCompra = "400011"
               psOpeVenta = "400012"
               If chkConsol.value = vbChecked Then
                 '*** PEAC 20120823
                  'lsCadena = lsCadena & GetMovCompraVentaME(psOpeCompra, psOpeVenta, CDate(Me.mskFecIni), CDate(Me.mskFecFin))
                  For j = 0 To DateDiff("d", CDate(Me.mskFecIni), CDate(Me.mskFecFin))
                    oBarra.Progress i + 1, "CONSOLIDADO DE COMPRA-VENTA", lsAgeNom, "Procesando día " & DateAdd("d", j, CDate(Me.mskFecIni)) & "..."
                    DoEvents
                    If j Mod 5 = 0 And j <> 0 Then lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                        If CDate(Me.mskFecIni) <> DateAdd("d", j, CDate(Me.mskFecIni)) Then
                            'MIOL ***
                            lsCadena = lsCadena & GetMovCompraVentaME(psOpeCompra, psOpeVenta, CDate(Me.mskFecIni), DateAdd("d", j, CDate(Me.mskFecIni)), Left(lstAge.List(i), 2), Trim(Mid(lsAgeNom, 3, 100)), psOpeCompraFI, psOpeVentaFI)
                            'lsCadena = lsCadena & GetMovCompraVentaME(psOpeCompra, psOpeVenta, CDate(Me.mskFecIni), DateAdd("d", j, CDate(Me.mskFecIni)), Left(lstAge.List(i), 2), Trim(Mid(lsAgeNom, 3, 100)))
                            'END MIOL
                        End If
                  Next j
                  '*** FIN PEAC
               Else
                  For j = 0 To DateDiff("d", CDate(Me.mskFecIni), CDate(Me.mskFecFin))
                    oBarra.Progress i + 1, "CONSOLIDADO DE COMPRA-VENTA", lsAgeNom, "Procesando día " & DateAdd("d", j, CDate(Me.mskFecIni)) & "..."
                    DoEvents
                    If j Mod 5 = 0 And j <> 0 Then lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                    'MIOL ***
                    lsCadena = lsCadena & GetMovCompraVentaME(psOpeCompra, psOpeVenta, DateAdd("d", j, CDate(Me.mskFecIni)), DateAdd("d", j, CDate(Me.mskFecIni)), Left(lstAge.List(i), 2), Trim(Mid(lsAgeNom, 3, 100)), psOpeCompraFI, psOpeVentaFI)
                    'lsCadena = lsCadena & GetMovCompraVentaME(psOpeCompra, psOpeVenta, DateAdd("d", j, CDate(Me.mskFecIni)), DateAdd("d", j, CDate(Me.mskFecIni)), Left(lstAge.List(i), 2), Trim(Mid(lsAgeNom, 3, 100)))
                    'END MIOL
                  Next j
               End If
            'END MIOL ***********************************************************
            Else
               If chkConsol.value = vbChecked Then
                  lsCadena = lsCadena & GetMovCompraVentaME(psOpeCompra, psOpeVenta, CDate(Me.mskFecIni), CDate(Me.mskFecFin), Left(lstAge.List(i), 2), Trim(Mid(lstAge.List(i), 3, 100)))
               Else
                  For j = 0 To DateDiff("d", CDate(Me.mskFecIni), CDate(Me.mskFecFin))
                    oBarra.Progress i + 1, "CONSOLIDADO DE COMPRA-VENTA", lsAgeNom, "Procesando día " & DateAdd("d", j, CDate(Me.mskFecIni)) & "..."
                    If j Mod 5 = 0 And j <> 0 Then lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                    'MIOL 20121022, SEGUN RQ12338****************************************
                    lsCadena = lsCadena & GetMovCompraVentaME(psOpeCompra, psOpeVenta, DateAdd("d", j, CDate(Me.mskFecIni)), DateAdd("d", j, CDate(Me.mskFecIni)), Left(lstAge.List(i), 2), Trim(Mid(lsAgeNom, 3, 100)), psOpeCompraFI, psOpeVentaFI)
                    'lsCadena = lsCadena & GetMovCompraVentaME(psOpeCompra, psOpeVenta, DateAdd("d", j, CDate(Me.mskFecIni)), DateAdd("d", j, CDate(Me.mskFecIni)), Left(lstAge.List(i), 2), Trim(Mid(lsAgeNom, 3, 100)))
                    'END MIOL ***********************************************************
                  Next j
               End If
            End If
        End If
    Next i
    oBarra.CloseForm Me
    Set oBarra = Nothing
    GeneraRCompraVenta = lsCadena
End Function
'END MIOL **********************************************

Private Function GetMovCompraVentaME(psOpeCompra As String, psOpeVenta As String, pdFecIni As Date, pdFecFin As Date, Optional psCodAge As String = "", Optional psAgeNom As String = "", Optional psOpeCompraFI As String, Optional psOpeVentaFI As String) As String
'Private Function GetMovCompraVentaME(psOpeCompra As String, psOpeVenta As String, pdFecIni As Date, pdFecFin As Date, Optional psCodAge As String = "", Optional psAgeNom As String = "") As String MIOL
    Dim lnMonComDol As Currency
    Dim lnMonVenDol As Currency
    Dim lnMonComSol As Currency
    Dim lnMonVenSol As Currency
    Dim lnPromCom As Currency
    Dim lnPromVen As Currency
    Dim lnCotMaxCom As Currency
    Dim lnCotMaxVen As Currency
    Dim lnCotMinCom As Currency
    Dim lnCotMinVen As Currency
    
    Dim rsRes As New ADODB.Recordset
            
    Dim oDol As New NCompraVenta
   lnMonComDol = 0
   lnMonComSol = 0
   lnMonVenDol = 0
   lnMonVenSol = 0
    
   'Promedio, Maximo, Minimo
   lnPromCom = 0
   lnCotMaxCom = 0
   lnCotMinCom = 0

   lnPromVen = 0
   lnCotMaxVen = 0
   lnCotMinVen = 0
   'MIOL 20121022, SEGUN RQ12338 ********************************************************
   If psOpeCompra = "400011" Then
        Set rsRes = oDol.GetImporteCompraVentaFinanzas(psOpeCompra, pdFecIni, pdFecFin, psCodAge)
   Else
        Set rsRes = oDol.GetImporteCompraVentaMasFinanzas(psOpeCompra, pdFecIni, pdFecFin, psCodAge, psOpeCompraFI)
'        Set rsRes = oDol.GetImporteCompraVenta(psOpeCompra, pdFecIni, pdFecFin, psCodAge)
   End If
   'Set rsRes = oDol.GetImporteCompraVenta(psOpeCompra, pdFecIni, pdFecFin, psCodAge)
   'END MIOL ****************************************************************************
   If Not rsRes.EOF Then
      lnMonComDol = rsRes!TotalDol
      lnMonComSol = rsRes!TotalSol
      lnPromCom = rsRes!TCPromedio
      lnCotMaxCom = rsRes!TCMaximo
      lnCotMinCom = rsRes!TCMinimo
   End If
   'MIOL 20121022, SEGUN RQ12338 ********************************************************
   If psOpeVenta = "400012" Then
        Set rsRes = oDol.GetImporteCompraVentaFinanzas(psOpeVenta, pdFecIni, pdFecFin, psCodAge)
   Else
        Set rsRes = oDol.GetImporteCompraVentaMasFinanzas(psOpeVenta, pdFecIni, pdFecFin, psCodAge, psOpeVentaFI)
'        Set rsRes = oDol.GetImporteCompraVenta(psOpeVenta, pdFecIni, pdFecFin, psCodAge)
   End If
   'Set rsRes = oDol.GetImporteCompraVenta(psOpeVenta, pdFecIni, pdFecFin, psCodAge)
   'END MIOL ****************************************************************************
   If Not rsRes.EOF Then
      lnMonVenDol = rsRes!TotalDol
      lnMonVenSol = rsRes!TotalSol
      lnPromVen = rsRes!TCPromedio
      lnCotMaxVen = rsRes!TCMaximo
      lnCotMinVen = rsRes!TCMinimo
   End If
   RSClose rsRes
Set oDol = Nothing
   GetMovCompraVentaME = DibujaCompraVenta(lnMonComDol, lnMonVenDol, lnMonComSol, lnMonVenSol, lnPromCom, lnPromVen, lnCotMaxCom, lnCotMaxVen, lnCotMinCom, lnCotMinVen, psCodAge, psAgeNom, pdFecIni, pdFecFin)
End Function


Private Function DibujaCompraVenta(pnMonComDol As Currency, pnMonVenDol As Currency, _
                                   pnMonComSol As Currency, pnMonVenSol As Currency, _
                                   pnPromCom As Currency, pnPromVen As Currency, _
                                   pnCotMaxCom As Currency, pnCotMaxVen As Currency, _
                                   pnCotMinCom As Currency, pnCotMinVen As Currency, _
                                   psCodAge As String, psAgeNom As String, pdFecIni As Date, pdFecFin As Date)

    Dim lsCadena As String
    Dim lnMargen As Integer
    Dim lsNombreAg As String
    If psCodAge <> "" Then
        lsNombreAg = psAgeNom
     Else
        lsNombreAg = "CONSOLIDADO"
    End If
    
    lnMargen = 5
    lsCadena = ""
    
    lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMargen) & Space(30) & "Listado de Compra Venta de Dolares - " & lsNombreAg & oImpresora.gPrnSaltoLinea
    If DateDiff("d", pdFecIni, pdFecFin) = 0 Then
        lsCadena = lsCadena & Space(lnMargen) & Space(55) & "Dia: " & Format(pdFecIni, "dd/mm/yyyy") & oImpresora.gPrnSaltoLinea
    Else
        lsCadena = lsCadena & Space(lnMargen) & Space(55) & "Rango: " & Format(pdFecIni, "dd/mm/yyyy") & " - " & Format(pdFecFin, "dd/mm/yyyy") & oImpresora.gPrnSaltoLinea
    End If
    lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnSaltoLinea
    
    lsCadena = lsCadena & Space(lnMargen) & "+" & String(6, "-") & "+" & String(9, "-") & "+" & String(18, "-") & "+" & String(19, "-") & "+" & String(18, "-") & "+" & String(18, "-") & "+" & String(10, "-") & "+" & String(15, "-") & "+" & String(15, "-") & "+" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMargen) & "¦" & "      " & "¦" & "         " & "¦" & "                  " & "¦" & "                   " & "¦" & "                  " & "¦" & " Monto en Moneda  " & "¦" & "          " & "¦" & "               " & "¦" & "               " & "¦" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMargen) & "¦" & " Fila " & "¦" & " Monedas " & "¦" & " Codigo de Moneda " & "¦" & " Tipo de Operacion " & "¦" & " Monto en Dolares " & "¦" & "     Nacional     " & "¦" & " Promedio " & "¦" & " Cotiz. Maxima " & "¦" & " Cotiz. Minima " & "¦" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMargen) & "+" & String(6, "-") & "+" & String(9, "-") & "+" & String(18, "-") & "+" & String(19, "-") & "+" & String(18, "-") & "+" & String(18, "-") & "+" & String(10, "-") & "+" & String(15, "-") & "+" & String(15, "-") & "+" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMargen) & "¦" & "  100 " & "¦" & " Datos   " & "¦" & " Dolar de N.A.    " & "¦" & " Compra            " & "¦" & Space(18 - Len(Format(pnMonComDol, "#,##0.00"))) & Format(pnMonComDol, "#,##0.00") & "¦" & Space(18 - Len(Format(pnMonComSol, "#,##0.00"))) & Format(pnMonComSol, "#,##0.00") & "¦" & Space(10 - Len(Format(pnPromCom, "#,##0.000000"))) & Format(pnPromCom, "#,##0.000000") & "¦" & Space(15 - Len(Format(pnCotMaxCom, "#,##0.000000"))) & Format(pnCotMaxCom, "#,##0.000000") & "¦" & Space(15 - Len(Format(pnCotMinCom, "#,##0.000000"))) & Format(pnCotMinCom, "#,##0.000000") & "¦" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMargen) & "¦" & "  101 " & "¦" & "         " & "¦" & " Dolar de N.A.    " & "¦" & " Venta             " & "¦" & Space(18 - Len(Format(pnMonVenDol, "#,##0.00"))) & Format(pnMonVenDol, "#,##0.00") & "¦" & Space(18 - Len(Format(pnMonVenSol, "#,##0.00"))) & Format(pnMonVenSol, "#,##0.00") & "¦" & Space(10 - Len(Format(pnPromVen, "#,##0.000000"))) & Format(pnPromVen, "#,##0.000000") & "¦" & Space(15 - Len(Format(pnCotMaxVen, "#,##0.000000"))) & Format(pnCotMaxVen, "#,##0.000000") & "¦" & Space(15 - Len(Format(pnCotMinVen, "#,##0.000000"))) & Format(pnCotMinVen, "#,##0.000000") & "¦" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMargen) & "+" & String(6, "-") & "+" & String(9, "-") & "+" & String(18, "-") & "+" & String(19, "-") & "+" & String(18, "-") & "+" & String(18, "-") & "+" & String(10, "-") & "+" & String(15, "-") & "+" & String(15, "-") & "+" & oImpresora.gPrnSaltoLinea
    
    DibujaCompraVenta = lsCadena
End Function


Private Sub cmdImprimir_Click()
Dim sImpre As String
lsOpeCompra = gOpeCajeroMECompra
lsOpeVenta = gOpeCajeroMEVenta
lsOpeCompraFI = gOpeMECompraAInst 'MIOL 20121020, SEGUN RQ12338
lsOpeVentaFI = gOpeMEVentaAInst 'MIOL 20121020, SEGUN RQ12338
sImpre = GeneraRCompraVenta(lsOpeCompra, lsOpeVenta, lsOpeCompraFI, lsOpeVentaFI)
'lsCadena, sTitulo, True, 66, gImpresora
     
EnviaPrevio sImpre, gsOpeDesc, gnLinPage, True
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
frmOperaciones.Enabled = False
Dim oAge As New DActualizaDatosArea
Set rs = oAge.GetAgencias(, False)
Do While Not rs.EOF
   lstAge.AddItem rs!Codigo & " " & rs!Descripcion
   rs.MoveNext
Loop
lstAge.AddItem "CONSOLIDADO"
lstAge.AddItem "OF FINANZAS" 'MIOL 20121020, SEGUN RQ12338
Set oAge = Nothing
RSClose rs
mskFecIni = gdFecSis
mskFecFin = gdFecSis
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmOperaciones.Enabled = True
End Sub

Private Sub MskFecFin_GotFocus()
fEnfoque mskFecFin
End Sub

Private Sub MskFecFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cmdImprimir.SetFocus
    End If
End Sub

Private Sub mskFecIni_GotFocus()
fEnfoque mskFecIni
End Sub

Private Sub mskFecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      mskFecFin.SetFocus
   End If
End Sub
