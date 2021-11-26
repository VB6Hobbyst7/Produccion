VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCredCargoAuto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cargo Automatico"
   ClientHeight    =   5145
   ClientLeft      =   2295
   ClientTop       =   2670
   ClientWidth     =   9210
   Icon            =   "frmCredCargoAuto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   690
      Left            =   105
      TabIndex        =   2
      Top             =   4185
      Width           =   9015
      Begin VB.CommandButton CmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   405
         Left            =   7320
         TabIndex        =   4
         Top             =   180
         Width           =   1605
      End
      Begin VB.CommandButton CmdPagar 
         Caption         =   "A&mortizar"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   180
         Width           =   1605
      End
      Begin VB.Label lblMensaje 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   390
         Left            =   1995
         TabIndex        =   12
         Top             =   180
         Width           =   5100
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4125
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   9045
      Begin MSComCtl2.DTPicker txtFecha 
         Height          =   300
         Left            =   3480
         TabIndex        =   9
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   529
         _Version        =   393216
         Format          =   66256897
         CurrentDate     =   38371
      End
      Begin VB.OptionButton optMon 
         Caption         =   "Dolares"
         Height          =   270
         Index           =   1
         Left            =   1035
         TabIndex        =   7
         Top             =   255
         Width           =   945
      End
      Begin VB.CommandButton CmdAplicar 
         Caption         =   "&Aplicar"
         Height          =   375
         Left            =   7575
         TabIndex        =   6
         Top             =   180
         Width           =   1320
      End
      Begin VB.OptionButton optMon 
         Caption         =   "Soles"
         Height          =   270
         Index           =   0
         Left            =   255
         TabIndex        =   5
         Top             =   255
         Value           =   -1  'True
         Width           =   945
      End
      Begin SICMACT.FlexEdit FECtas 
         Height          =   2940
         Left            =   135
         TabIndex        =   1
         Top             =   645
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   5186
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "I-OK-Credito-Titular-Monto-Estado-MetLiq-q"
         EncabezadosAnchos=   "400-400-2000-4000-1200-0-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-4-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-L-R-C-C-C"
         FormatosEdit    =   "0-0-0-0-2-0-0-0"
         TextArray0      =   "I"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label lblTotalPagar 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7050
         TabIndex        =   11
         Top             =   3675
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total a Amortizar :"
         Height          =   195
         Left            =   5205
         TabIndex        =   10
         Top             =   3720
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuotas Vencen al :"
         Height          =   195
         Left            =   2115
         TabIndex        =   8
         Top             =   293
         Width           =   1350
      End
   End
End
Attribute VB_Name = "frmCredCargoAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim R As ADODB.Recordset
Dim NumRepo As Integer
Dim oBase As COMDCredito.DCOMCredActBD

Dim fnTipCambioC As Double
Dim fnTipCambioV As Double

Private Sub cmdAplicar_Click()
Dim oDCred As COMDCredito.DCOMCredito
Dim ldFecha As Date
Dim lnMoneda As Moneda

    Set oDCred = New COMDCredito.DCOMCredito
    ldFecha = txtfecha.value
    lnMoneda = IIf(optMon(0).value = True, gMonedaNacional, gMonedaExtranjera)
    Set R = oDCred.RecuperaCreditosCargoAutomatico(ldFecha, lnMoneda)
    Set oDCred = Nothing
    FECtas.Clear
    FECtas.Rows = 2
    FECtas.FormaCabecera
    If Not R.EOF And Not R.BOF Then
        Set FECtas.Recordset = R
    Else
        MsgBox "No existe Cuotas pendiente a la Fecha seleccionada", vbInformation, "Aviso"
    End If
    R.Close
    
End Sub

'Private Sub GeneraReporte(ByVal MatRep As Variant)
'Dim sCadImp As String
'Dim i As Integer
'Dim nSuma1 As Double
'Dim nSuma2 As Double
'Dim nSuma3 As Double
'Dim oPrev As Previo.clsPrevio
'
'    sCadImp = Chr$(10)
'    sCadImp = sCadImp & Space(5) & gsNomAge & Chr$(10)
'    sCadImp = sCadImp & Space(5) & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time(), "hh:mm:ss") & Chr$(10) & Chr$(10)
'    sCadImp = sCadImp & Space(30) & UCase(gsNomCmac) & Chr$(10)
'    sCadImp = sCadImp & PrnSet("B+") & Space(20) & "REPORTE DE AMORTIZACIONES POR CARGO AUTOMATICO" & Chr$(10) & PrnSet("B-") & Chr$(10)
'    sCadImp = sCadImp & Space(16) & "AMORTIZACION DE CREDITOS EN " & IIf(optMon(0).value = True, "MONEDA NACIONAL", "MONEDA EXTRANJERA") & " AL " & gdFecSis & Chr$(10) & Chr$(10)
'    If optMon(0).value = False Then
'        sCadImp = sCadImp & Space(20) & "TIPO DE CAMBIO : VENTA =" & Format(gnTipCambioV, "#0.0000") & " COMPRA=" & Format(gnTipCambioC, "#0.0000") & Chr$(10)
'    End If
'    sCadImp = sCadImp & String(88, "-") & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("CREDITO", 20) & ImpreFormat("TITULAR", 33) & ImpreFormat("MONTO", 8) & ImpreFormat("PAGO", 8) & ImpreFormat("SALDO", 8) & Chr$(10)
'    sCadImp = sCadImp & String(88, "-") & Chr$(10)
'    nSuma1 = 0
'    nSuma2 = 0
'    For i = 0 To NumRepo - 1
'        sCadImp = sCadImp & ImpreFormat(MatRep(i, 0), 20) & ImpreFormat(MatRep(i, 1), 30) & ImpreFormat(CDbl(MatRep(i, 2)), 8) & PrnSet("B+") & ImpreFormat(CDbl(MatRep(i, 3)), 8) & PrnSet("B-") & ImpreFormat(CDbl(MatRep(i, 4)), 6) & Chr$(10)
'        nSuma1 = nSuma1 + CDbl(MatRep(i, 2))
'        nSuma2 = nSuma2 + CDbl(MatRep(i, 3))
'        nSuma3 = nSuma3 + CDbl(MatRep(i, 4))
'    Next i
'    sCadImp = sCadImp & String(88, "-") & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat(nSuma1, 62) & ImpreFormat(nSuma2, 8) & ImpreFormat(nSuma3, 6) & Chr$(10)
'    sCadImp = sCadImp & String(88, "-") & Chr$(10)
'
'    Set oPrev = New Previo.clsPrevio
'    oPrev.Show sCadImp, "Cargo Automatico"
'    Set oPrev = Nothing
'
'End Sub

Private Sub CmdPagar_Click()

'Dim i As Integer
'Dim nMontoPago As Double
'Dim nMontoAmort As Double
'Dim oCred As COMDCredito.DCOMCredito
'Dim R As ADODB.Recordset
'Dim RAho As ADODB.Recordset
'Dim oAho As COMDCredito.DCOMCredActBD
'Dim MatRepo() As String
'Dim oFunciones As COMNContabilidad.NCOMContFunciones
'Dim sMovNro As String
'Dim sMovNroTemp As String
'Dim nMontoAmortAcum As Double
'Dim sEstado As String
'Dim sConsCred As String
'Dim MatCalendTemp As Variant
'Dim MatCalendDistrTemp As Variant
'Dim oNCredito As COMNCredito.NCOMCredito
'Dim nMovNro As Long
'Dim bTran As Boolean
'Dim MatAhorros() As String
'Dim bAmort As Boolean
'Dim nMiViv As Integer
'Dim MatCalend_2 As Variant
'Dim MatCalendNormalT1 As Variant
'Dim MatCalendParalelo As Variant
'Dim MatCalendMiVivResult As Variant
'Dim MatCalendTmp As Variant
'Dim MatCalendDistribuido As Variant
'Dim oGastos As COMNCredito.NCOMGasto
'Dim MatGastosFinal As Variant
'Dim MatCalendDistribuido_2 As Variant
'Dim MatCalendDistribuidoParalelo As Variant
'Dim nNumGastosFinal As Integer
'Dim nCalPago As Integer
'Dim MatMovAho() As String
'Dim NumMatMovAho As Integer
'Dim K As Integer
'
'Dim nMontoGastoGen As Double
'Dim MatSaldosAho(100, 2) As String
'Dim NumMatSaldosAho As Integer
'Dim MatCargosAho() As String
'Dim NumMatCargosAho As Integer
'Dim CredDoc As COMNCredito.NCOMCredDoc

Dim oCred As COMNCredito.NCOMCredito
Dim oPrevio As previo.clsprevio
Dim MatCuentas As Variant
Dim MatRepBoleta As Variant
Dim MatRepBoletaAho As Variant
Dim sReporte As String
Dim i As Integer
Dim j As Integer
Dim MatNumCargosAho As Variant
Dim MatCuentasIndice() As Integer

'Para optimizar
Dim nNumCuentas As Integer

On Error GoTo ErrorPago

    If FECtas.Rows = 1 Then
        MsgBox "No Existen Registros que Amortizar", vbInformation, "Aviso"
        Exit Sub
    End If

    If FECtas.Rows > 1 Then
        If Trim(FECtas.TextMatrix(1, 2)) = "" Then
            MsgBox "No Existen Registros que Amortizar", vbInformation, "Aviso"
            Exit Sub
        End If
    End If

    If MsgBox("Desea Realizar la Amortizacion de Creditos con Cargo a Cuenta??", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    lblMensaje.Visible = True
    'lblMensaje.Refresh

    'Para Optimizar
    nNumCuentas = 0
    'ReDim MatCuentas(0, 0)
    For i = 1 To FECtas.Rows - 1
        If FECtas.TextMatrix(i, 1) = "." Then
            nNumCuentas = nNumCuentas + 1
            ReDim Preserve MatCuentasIndice(nNumCuentas)
            MatCuentasIndice(nNumCuentas) = i
        End If
    Next i
    
    ReDim MatRepBoleta(nNumCuentas)  '(FECtas.Rows - 1)
    ReDim MatCuentas(nNumCuentas, 6) '(FECtas.Rows - 1, 6)
    ReDim MatNumCargosAho(nNumCuentas)  '(FECtas.Rows - 1)
    ReDim MatRepBoletaAho(nNumCuentas, 100) '(FECtas.Rows - 1, 100)
    
    'Llenar Datos de las Cuentas
    For i = 1 To nNumCuentas
        For j = 2 To 6
            MatCuentas(i, j) = FECtas.TextMatrix(MatCuentasIndice(i), j)
        Next j
    Next i
    
    lblMensaje.Caption = "Amortizando Creditos Seleccionados"
    'lblMensaje.Refresh
        
    'For I = 1 To nNumCuentas  'FECtas.Rows - 1
    '    For j = 2 To 6  '1 to 6
    '        MatCuentas(I, j) = FECtas.TextMatrix(I, j)
    '    Next j
    'Next I

    Set oCred = New COMNCredito.NCOMCredito

    Call oCred.PagarCargoAutomatico(gsNomAge, gdFecSis, gsNomCmac, optMon(0).value, gsCodAge, _
                                    gsCodUser, sLpt, gsInstCmac, gsCodCMAC, fnTipCambioV, _
                                    fnTipCambioC, nNumCuentas, MatCuentas, MatRepBoleta, MatRepBoletaAho, _
                                    sReporte, MatNumCargosAho)  'FECtas.Rows,

    lblMensaje.Caption = "Amortización con Cargo Automático Culminado con éxito"
    'lblMensaje.Refresh

    Set oCred = Nothing

    Set oPrevio = New previo.clsprevio
    
    For i = 1 To nNumCuentas 'FECtas.Rows - 1
        If MatRepBoleta(i) <> "" Then
            'oPrevio.Show CStr(MatRepBoleta(i)), "Boleta"
            oPrevio.Show CStr(MatRepBoleta(i)), "Boleta", , , gImpresora
        End If
        For j = 0 To MatNumCargosAho(i) - 1
            If MatRepBoletaAho(i, j) <> "" Then
                oPrevio.Show CStr(MatRepBoletaAho(i, j)), "Boleta de Ahorros", , , gImpresora
            End If
        Next j
    Next i
            
    oPrevio.Show sReporte, "Cargo Automatico", , , gImpresora
    
    Set oPrevio = Nothing

    Set FECtas.Recordset = Nothing
    Call LimpiaFlex(FECtas)
    
    Screen.MousePointer = vbDefault
    lblMensaje = ""
    lblMensaje.Visible = False
    'lblMensaje.Refresh

    Exit Sub
    
ErrorPago:
    MsgBox Err.Description, vbInformation, "Aviso"

'
'    If FECtas.Rows = 1 Then
'        MsgBox "No Existen Registros que Amortizar", vbInformation, "Aviso"
'        Exit Sub
'    End If
'
'    If FECtas.Rows > 1 Then
'        If Trim(FECtas.TextMatrix(1, 2)) = "" Then
'            MsgBox "No Existen Registros que Amortizar", vbInformation, "Aviso"
'            Exit Sub
'        End If
'    End If
'
'    If MsgBox("Desea Realizar la Amortizacion de Creditos con Cargo a Cuenta??", vbQuestion + vbYesNo, "Aviso") = vbNo Then
'        Exit Sub
'    End If
'
'    Screen.MousePointer = 11
'    lblMensaje.Visible = True
'    lblMensaje.Refresh
'    ReDim MatRepo(FECtas.Rows - 1, 5)
'    NumRepo = 0
'    bAmort = False
'    For i = 1 To FECtas.Rows - 1
'        If FECtas.TextMatrix(i, 1) = "." Then
'             bAmort = True
'             bTran = False
'             nMontoPago = CDbl(FECtas.TextMatrix(i, 4))
'             sEstado = FECtas.TextMatrix(i, 5)
'             '*****************************************
'             'Recupero las Cuentas
'             '*****************************************
'             Set oCred = New COMDCredito.DCOMCredito
'             Set R = oCred.RecuperaDatosCreditoVigente(FECtas.TextMatrix(i, 2))
'             nMiViv = IIf(IsNull(R!bMiVivienda), 0, R!bMiVivienda)
'             nCalPago = IIf(IsNull(R!nCalPago), 0, R!nCalPago)
'             Set R = oCred.RecuperaCuentasAho(FECtas.TextMatrix(i, 2))
'             Set oCred = Nothing
'             nMontoAmortAcum = 0
'
'
'             '*****************************************************************
'             'CARGA MATRIZ DE PAGOS DE LOS CREDITOS
'             '*****************************************************************
'             If nMiViv = 1 Then
'                Set oNCredito = New COMNCredito.NCOMCredito
'                MatCalendTemp = oNCredito.RecuperaMatrizCalendarioPendiente(FECtas.TextMatrix(i, 2))
'                MatCalend_2 = MatCalendTemp
'                MatCalendNormalT1 = MatCalendTemp
'                MatCalendParalelo = oNCredito.RecuperaMatrizCalendarioPendiente(FECtas.TextMatrix(i, 2), True)
'                MatCalendMiVivResult = UnirMatricesMiViviendaAmortizacion(MatCalendTemp, MatCalendParalelo)
'                MatCalendTemp = MatCalendMiVivResult
'                MatCalendTmp = MatCalendTemp
'             Else
'                Set oNCredito = New COMNCredito.NCOMCredito
'                MatCalendTemp = oNCredito.RecuperaMatrizCalendarioPendiente(FECtas.TextMatrix(i, 2))
'            End If
'
'            MatCalendDistrTemp = oNCredito.CrearMatrizparaAmortizacion(MatCalendTemp)
'
'
'            '**************************************************
'             '**************    AHORROS
'             '**************************************************
'             Set oBase = New COMDCredito.DCOMCredActBD
'             bTran = True
'             ReDim MatMovAho(0)
'             NumMatMovAho = 0
'             NumMatSaldosAho = 0
'             nMontoPago = CDbl(FECtas.TextMatrix(i, 4))
'             nMontoAmortAcum = 0
'
'             Do While Not R.EOF
'                MatSaldosAho(NumMatSaldosAho, 0) = R!cCodCtaAho
'                Set RAho = oBase.GetSaldoFecha(R!cCodCtaAho, gdFecSis)
'                MatSaldosAho(NumMatSaldosAho, 1) = Format(RAho!nSaldoDisponible, "#0.00")
'                NumMatSaldosAho = NumMatSaldosAho + 1
'                R.MoveNext
'             Loop
'             R.Close
'
'            '**************************************************
'            'Halla el Calculo de Cuanto se va  a pagar segun saldos de cuentas de ahorros
'            '**************************************************
'            Dim nMontoPagoTC As Double
'
'            For K = 0 To NumMatSaldosAho - 1
'                 If nMontoPago > 0 Then
'                     '*****************************************
'                     'Recupero los Saldos de las Cuentas
'                     '*****************************************
'                     If Mid(FECtas.TextMatrix(i, 2), 9, 1) = Mid(MatSaldosAho(K, 0), 9, 1) Then
'                        nMontoAmort = 0
'                        If CDbl(MatSaldosAho(K, 1)) > 0 Then
'                            If CDbl(MatSaldosAho(K, 1)) >= nMontoPago Then
'                                nMontoAmort = nMontoPago
'                            Else
'                                nMontoAmort = CDbl(MatSaldosAho(K, 1))
'                            End If
'                            nMontoAmortAcum = nMontoAmortAcum + nMontoAmort
'                            nMontoPago = nMontoPago - nMontoAmort
'                        End If
'                    Else
'                        'si la cuenta de creditos es dolares y el cuenta en soles debe hacer una venta de dolares
'                        If Mid(FECtas.TextMatrix(i, 2), 9, 1) = gMonedaExtranjera And Mid(MatSaldosAho(K, 0), 9, 1) = gMonedaNacional Then
'                            nMontoAmort = 0
'                            If CDbl(MatSaldosAho(K, 1)) > 0 Then
'                                nMontoPagoTC = Round(nMontoPago * gnTipCambioV, 2)
'                                If CDbl(MatSaldosAho(K, 1)) >= nMontoPagoTC Then
'                                    nMontoAmort = nMontoPago
'                                Else
'                                    ' amortiza el saldo de ahorros en moneda nacional debe cambiarse al
'                                    'tipo de cambio venta para  calcular el monto del pago a amortizar en dolares
'                                    nMontoAmort = Round(CDbl(MatSaldosAho(K, 1)) / gnTipCambioV, 2)
'                                End If
'                                nMontoAmortAcum = nMontoAmortAcum + nMontoAmort
'                                nMontoPago = nMontoPago - nMontoAmort
'                            End If
'                        Else
'                            ' si la cuenta de creditos es en soles y la cuenta de ahorros en dolares
'                            'debe hacerse una compra dolares
'                            nMontoAmort = 0
'                            If CDbl(MatSaldosAho(K, 1)) > 0 Then
'                                nMontoPagoTC = Round(nMontoPago / gnTipCambioC, 2)
'                                If CDbl(MatSaldosAho(K, 1)) >= nMontoPagoTC Then
'                                    nMontoAmort = nMontoPago
'                                Else
'                                    'amortiza el saldo de ahorros en moneda extrajera debe cambiarse al
'                                    'tipo de cambio compra para  calcular el monto del pago a amortizar en soles
'                                    nMontoAmort = Round(CDbl(MatSaldosAho(K, 1)) * gnTipCambioC, 2)
'                                End If
'                                nMontoAmortAcum = nMontoAmortAcum + nMontoAmort
'                                nMontoPago = nMontoPago - nMontoAmort
'                            End If
'                        End If
'                    End If
'                End If
'            Next K
'
'            '*********************************************************
'            ' CREDITOS CARGA GASTOS
'            '*********************************************************
'            Set oGastos = New COMNCredito.NCOMGasto
'            Set oNCredito = New COMNCredito.NCOMCredito
'            MatGastosFinal = oGastos.GeneraCalendarioGastos(Array(0), Array(0), nNumGastosFinal, gdFecSis, FECtas.TextMatrix(i, 2), 1, "PA", , , nMontoAmortAcum, oNCredito.MatrizMontoCapitalAPagar(MatCalendTemp, gdFecSis), oNCredito.MatrizCuotaPendiente(MatCalendTemp, MatCalendDistrTemp))
'            'obtener el total del gastos
'            nMontoGastoGen = MontoTotalGastosGenerado(MatGastosFinal, nNumGastosFinal, Array("PA", "", ""))
'            'MatCalend = MatCalendTmp
'            MatCalendTemp(0, 9) = Format(CDbl(MatCalendTemp(0, 9)) + nMontoGastoGen, "#0.00")
'            Set oGastos = Nothing
'
'            nMontoPago = nMontoAmortAcum + nMontoGastoGen
'            nMontoAmortAcum = 0
'
'            '**************************************************
'            'Halla el ReCalculo de Cuanto se va  a pagar ya que se puede haber agregado gastos
'            '***************************************************
'            For K = 0 To NumMatSaldosAho - 1
'                If nMontoPago > 0 Then
'                     '*****************************************
'                     'Recupero los Saldos de las Cuentas
'                     '*****************************************
''                     nMontoAmort = 0
''                     If CDbl(MatSaldosAho(K, 1)) > 0 Then
''                         If CDbl(MatSaldosAho(K, 1)) >= nMontoPago Then
''                             nMontoAmort = nMontoPago
''                         Else
''                             nMontoAmort = CDbl(MatSaldosAho(K, 1))
''                         End If
''                         nMontoAmortAcum = nMontoAmortAcum + nMontoAmort
''                         nMontoPago = nMontoPago - nMontoAmort
''                     End If
'                    If Mid(FECtas.TextMatrix(i, 2), 9, 1) = Mid(MatSaldosAho(K, 0), 9, 1) Then
'                        nMontoAmort = 0
'                        If CDbl(MatSaldosAho(K, 1)) > 0 Then
'                            If CDbl(MatSaldosAho(K, 1)) >= nMontoPago Then
'                                nMontoAmort = nMontoPago
'                            Else
'                                nMontoAmort = CDbl(MatSaldosAho(K, 1))
'                            End If
'                            nMontoAmortAcum = nMontoAmortAcum + nMontoAmort
'                            nMontoPago = nMontoPago - nMontoAmort
'                        End If
'                    Else
'                        'si la cuenta de creditos es dolares y el cuenta en soles debe hacer una venta de dolares
'                        If Mid(FECtas.TextMatrix(i, 2), 9, 1) = gMonedaExtranjera And Mid(MatSaldosAho(K, 0), 9, 1) = gMonedaNacional Then
'                            nMontoAmort = 0
'                            If CDbl(MatSaldosAho(K, 1)) > 0 Then
'                                nMontoPagoTC = Round(nMontoPago * gnTipCambioV, 2)
'                                If CDbl(MatSaldosAho(K, 1)) >= nMontoPagoTC Then
'                                    nMontoAmort = nMontoPago
'                                Else
'                                    ' amortiza el saldo de ahorros en moneda nacional debe cambiarse al
'                                    'tipo de cambio venta para  calcular el monto del pago a amortizar en dolares
'                                    nMontoAmort = Round(CDbl(MatSaldosAho(K, 1)) / gnTipCambioV, 2)
'                                End If
'                                nMontoAmortAcum = nMontoAmortAcum + nMontoAmort
'                                nMontoPago = nMontoPago - nMontoAmort
'                            End If
'                        Else
'                            ' si la cuenta de creditos es en soles y la cuenta de ahorros en dolares
'                            'debe hacerse una compra dolares
'                            nMontoAmort = 0
'                            If CDbl(MatSaldosAho(K, 1)) > 0 Then
'                                nMontoPagoTC = Round(nMontoPago / gnTipCambioC, 2)
'                                If CDbl(MatSaldosAho(K, 1)) >= nMontoPagoTC Then
'                                    nMontoAmort = nMontoPago
'                                Else
'                                    'amortiza el saldo de ahorros en moneda extrajera debe cambiarse al
'                                    'tipo de cambio compra para  calcular el monto del pago a amortizar en soles
'                                    nMontoAmort = Round(CDbl(MatSaldosAho(K, 1)) * gnTipCambioC, 2)
'                                End If
'                                nMontoAmortAcum = nMontoAmortAcum + nMontoAmort
'                                nMontoPago = nMontoPago - nMontoAmort
'                            End If
'                        End If
'                    End If
'                End If
'            Next K
'            Sleep 1000
'
'            oBase.dBeginTrans
'            Set oFunciones = New COMNContabilidad.NCOMContFunciones
'            sMovNroTemp = oFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'            Set oFunciones = Nothing
'
'            'Distribuye Monto
'            If Mid(Trim(FECtas.TextMatrix(i, 6)), 3, 1) = "i" Or Mid(Trim(FECtas.TextMatrix(i, 6)), 3, 1) = "Y" Then
'                MatCalendDistrTemp = oNCredito.MatrizDistribuirCancelacion(Trim(FECtas.TextMatrix(i, 2)), MatCalendTemp, nMontoAmortAcum, FECtas.TextMatrix(i, 6), gdFecSis, True)
'            Else
'                MatCalendDistrTemp = oNCredito.MatrizDistribuirMonto(MatCalendTemp, nMontoAmortAcum, Trim(FECtas.TextMatrix(i, 6)))
'            End If
'
'            '*****************************************************************
'            'Amortiza Credito
'            '*****************************************************************
'            If nMiViv = 1 Then
'                MatCalendDistribuido_2 = oNCredito.CrearMatrizparaAmortizacion(MatCalendTemp)
'                MatCalendDistribuidoParalelo = oNCredito.CrearMatrizparaAmortizacion(MatCalendParalelo)
'                Call oNCredito.DistribuirMatrizMiVivEnDosCalendarios(MatCalendDistribuidoParalelo, MatCalendDistribuido_2, MatCalendDistrTemp, MatCalendParalelo, MatCalendNormalT1, 0)
'                MatCalendDistrTemp = MatCalendDistribuido_2
'                Call oNCredito.AmortizarCredito(FECtas.TextMatrix(i, 2), MatCalendTemp, MatCalendDistrTemp, nMontoAmortAcum, gdFecSis, FECtas.TextMatrix(i, 6), gColocTipoPagoCargoCta, gsCodAge, gsCodUser, , oBase, , , , , , , sMovNroTemp, MatGastosFinal, nNumGastosFinal, MatCalendDistribuidoParalelo, , MatCalendParalelo, 0)
'            Else
'                Call oNCredito.AmortizarCredito(FECtas.TextMatrix(i, 2), MatCalendTemp, MatCalendDistrTemp, nMontoAmortAcum, gdFecSis, FECtas.TextMatrix(i, 6), gColocTipoPagoCargoCta, gsCodAge, gsCodUser, , oBase, , , , , , , sMovNroTemp, MatGastosFinal, nNumGastosFinal)
'            End If
'
'
'            nMovNro = oBase.dGetnMovNro(sMovNroTemp)
'
'             '**************************************************
'             '**************    AHORROS
'             '**************************************************
'             nMontoPago = CDbl(FECtas.TextMatrix(i, 4)) + nMontoGastoGen
'             nMontoAmortAcum = 0
'             NumMatCargosAho = 0
'             ReDim MatCargosAho(100, 2)
'             Dim lnMontoCargo As Double
'             For K = 0 To NumMatSaldosAho - 1
'                 If nMontoPago > 0 Then
'                     '*****************************************
'                     'Recupero los Saldos de las Cuentas
'                     '*****************************************
''                     nMontoAmort = 0
''                     If CDbl(MatSaldosAho(K, 1)) > 0 Then
''                         If CDbl(MatSaldosAho(K, 1)) >= nMontoPago Then
''                             nMontoAmort = nMontoPago
''                         Else
''                             nMontoAmort = CDbl(MatSaldosAho(K, 1))
''                         End If
''                         nMontoAmortAcum = nMontoAmortAcum + nMontoAmort
''                         nMontoPago = nMontoPago - nMontoAmort
''                     End If
'                    If Mid(FECtas.TextMatrix(i, 2), 9, 1) = Mid(MatSaldosAho(K, 0), 9, 1) Then
'                        nMontoAmort = 0
'                        If CDbl(MatSaldosAho(K, 1)) > 0 Then
'                            If CDbl(MatSaldosAho(K, 1)) >= nMontoPago Then
'                                nMontoAmort = nMontoPago
'                            Else
'                                nMontoAmort = CDbl(MatSaldosAho(K, 1))
'                            End If
'                            nMontoAmortAcum = nMontoAmortAcum + nMontoAmort
'                            nMontoPago = nMontoPago - nMontoAmort
'                        End If
'                        lnMontoCargo = nMontoAmort
'                    Else
'                        'si la cuenta de creditos es dolares y el cuenta en soles debe hacer una venta de dolares
'                        If Mid(FECtas.TextMatrix(i, 2), 9, 1) = gMonedaExtranjera And Mid(MatSaldosAho(K, 0), 9, 1) = gMonedaNacional Then
'                            nMontoAmort = 0
'                            If CDbl(MatSaldosAho(K, 1)) > 0 Then
'                                nMontoPagoTC = Round(nMontoPago * gnTipCambioV, 2)
'                                If CDbl(MatSaldosAho(K, 1)) >= nMontoPagoTC Then
'                                    nMontoAmort = nMontoPago
'                                Else
'                                    ' amortiza el saldo de ahorros en moneda nacional debe cambiarse al
'                                    'tipo de cambio venta para  calcular el monto del pago a amortizar en dolares
'                                    nMontoAmort = Round(CDbl(MatSaldosAho(K, 1)) / gnTipCambioV, 2)
'                                End If
'                                nMontoAmortAcum = nMontoAmortAcum + nMontoAmort
'                                nMontoPago = nMontoPago - nMontoAmort
'                            End If
'                            lnMontoCargo = Round(nMontoAmort * gnTipCambioV, 2)
'                        Else
'                            ' si la cuenta de creditos es en soles y la cuenta de ahorros en dolares
'                            'debe hacerse una compra dolares
'                            nMontoAmort = 0
'                            If CDbl(MatSaldosAho(K, 1)) > 0 Then
'                                nMontoPagoTC = Round(nMontoPago / gnTipCambioC, 2)
'                                If CDbl(MatSaldosAho(K, 1)) >= nMontoPagoTC Then
'                                    nMontoAmort = nMontoPago
'                                Else
'                                    'amortiza el saldo de ahorros en moneda extrajera debe cambiarse al
'                                    'tipo de cambio compra para  calcular el monto del pago a amortizar en soles
'                                    nMontoAmort = Round(CDbl(MatSaldosAho(K, 1)) * gnTipCambioC, 2)
'                                End If
'                                nMontoAmortAcum = nMontoAmortAcum + nMontoAmort
'                                nMontoPago = nMontoPago - nMontoAmort
'                            End If
'                            lnMontoCargo = Round(nMontoAmort / gnTipCambioC, 2)
'                        End If
'                    End If
'
'                     '****************************************************************
'                     'Cargar a la Cuenta Cuenta
'                     '****************************************************************
'                     Select Case CInt(FECtas.TextMatrix(i, 5))
'                         Case gColocEstRefMor
'                             sConsCred = gCredPagRefMorCC
'
'                         Case gColocEstRefNorm
'                             sConsCred = gCredPagRefNorCC
'
'                         Case gColocEstRefVenc
'                             sConsCred = gCredPagRefVenCC
'
'                         'si es Credito Normal
'                         Case gColocEstVigMor
'                             sConsCred = gCredPagNorNorCC
'
'                         Case gColocEstVigNorm
'                             sConsCred = gCredPagNorMorCC
'
'                         Case gColocEstVigVenc
'                             sConsCred = gCredPagNorVenCC
'                     End Select
'
'
'                     'sMovNro = oBase.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'                     'Call oBase.dInsertMov(sMovNro, sConsCred, "", gMovEstContabMovContable, gMovFlagVigente, False)
'                     'nMovNro = oBase.dGetnMovNro(sMovNro)
'                     ReDim MatAhorros(14)
'                     MatAhorros(0) = MatSaldosAho(K, 0)
'                     MatAhorros(1) = MatSaldosAho(K, 0)
'
'                     oBase.CapCargoCuentaAho MatAhorros, MatSaldosAho(K, 0), CDbl(Format(nMontoAmort, "#0.00")), gCredDesembCtaRetiroCancelacion, sMovNroTemp, "PAGO CREDITO CARGO AUTOMATICO", TpoDocBolDep, "", "", False, True, , , , , , , , Mid(FECtas.TextMatrix(i, 2), 9, 1)
'
'                     NumMatMovAho = NumMatMovAho + 1
'                     ReDim Preserve MatMovAho(NumMatMovAho)
'                     MatMovAho(NumMatMovAho - 1) = Trim(Str(nMovNro))
'
'                     MatCargosAho(NumMatCargosAho, 0) = MatSaldosAho(K, 0)
'                     'MatCargosAho(NumMatCargosAho, 1) = Format(nMontoAmort, "#0.00")
'                     MatCargosAho(NumMatCargosAho, 1) = lnMontoCargo
'                     NumMatCargosAho = NumMatCargosAho + 1
'
'                     lblMensaje.Caption = "Amortizando Credito N°" & FECtas.TextMatrix(i, 2) & " En la Cuenta N°:" & MatSaldosAho(K, 0)
'                     lblMensaje.Refresh
'                 End If
'            Next K
'
'            oBase.dCommitTrans
'
'             '******************************************************************
'             'Imprimo las Boletas
'             '******************************************************************
'             'Credito
'             Set CredDoc = New COMNCredito.NCOMCredDoc
'             Set oNCredito = New COMNCredito.NCOMCredito
'
'             Call CredDoc.ImprimeBoleta(FECtas.TextMatrix(i, 2), FECtas.TextMatrix(i, 3), gsNomAge, IIf(Mid(FECtas.TextMatrix(i, 2), 9, 1) = "1", "SOLES", "DOLARES"), "", _
'             gdFecSis, Right(Format(FechaHora(gdFecSis), "dd/mm/yyyy hh:mm:ss"), 8), "", "", oNCredito.MatrizCapitalPagado(MatCalendDistrTemp), _
'             oNCredito.MatrizIntCompPagado(MatCalendDistrTemp), oNCredito.MatrizIntCompVencPagado(MatCalendDistrTemp), _
'             oNCredito.MatrizIntMorPagado(MatCalendDistrTemp), oNCredito.MatrizGastoPag(MatCalendDistrTemp), _
'             oNCredito.MatrizIntGraciaPagado(MatCalendDistrTemp), oNCredito.MatrizIntReprogPag(MatCalendDistrTemp), _
'             oNCredito.MatrizSaldoCapital(MatCalendTemp, MatCalendDistrTemp), oNCredito.MatrizFechaCuotaPendiente(MatCalendTemp, MatCalendDistrTemp), _
'             gsCodUser, sLpt, gsInstCmac, , , gsCodCMAC, , , , True)
'
'             'AHORROS
'             For K = 0 To NumMatCargosAho - 1
'                Call CredDoc.ImprimeBoletaAhorro("CARGO.PAGO.CRED.", "CARGO.PAGO.CRED.", "", MatCargosAho(K, 1), FECtas.TextMatrix(i, 3), MatCargosAho(K, 0), "", 0#, "", "", 0#, 0#, , , , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt)
'             Next K
'
'             Set oNCredito = Nothing
'             Set CredDoc = Nothing
'
'             '*******************************************************************
'             'Guarda Datos para Reporte
'             '*******************************************************************
'             MatRepo(NumRepo, 0) = FECtas.TextMatrix(i, 2) 'Cuenta
'             MatRepo(NumRepo, 1) = PstaNombre(FECtas.TextMatrix(i, 3)) 'Titular
'             MatRepo(NumRepo, 2) = FECtas.TextMatrix(i, 4) 'Monto de Pago
'             MatRepo(NumRepo, 3) = Format(nMontoAmortAcum, "#0.00") 'Pago
'             MatRepo(NumRepo, 4) = Format(nMontoPago, "#0.00") 'Saldo
'             NumRepo = NumRepo + 1
'        End If
'        DoEvents
'    Next i
'    Set oNCredito = Nothing
'    Screen.MousePointer = 0

'    lblMensaje.Caption = "Amortizacion con Cargo Automatico Culminado con exito"
'    lblMensaje.Refresh
'    '*******************************************************************
'    'Genera Reporte
'    '*******************************************************************
'    If bAmort Then
'        Call GeneraReporte(MatRepo)
'    End If
'
'    Set FECtas.Recordset = Nothing
'    Call LimpiaFlex(FECtas)
'
'    lblMensaje = ""
'    lblMensaje.Visible = False
'    lblMensaje.Refresh
'
'    Set oBase = Nothing
    
'    Exit Sub
    
'ErrorPago:
'    If bTran Then
'        oBase.dRollbackTrans
'        Set oBase = Nothing
'    End If
'    MsgBox Err.Description, vbInformation, "Aviso"
    
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub FECtas_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)

Dim i As Integer
Dim nTotCred As Long
Dim nMontoAPagar As Double
Dim lnMontoTC As Double

Set oBase = New COMDCredito.DCOMCredActBD
    If FECtas.TextMatrix(pnRow, pnCol) = "." Then
        
        If Mid(FECtas.TextMatrix(pnRow, 2), 9, 1) = Mid(FECtas.TextMatrix(pnRow, 7), 9, 1) Then
            lnMontoTC = CDbl(FECtas.TextMatrix(pnRow, 4))
        Else
            'SI EL CREDITO ES EN DOLARES Y LA CUENTA DE AHORROS ES EN SOLES..
            If Mid(FECtas.TextMatrix(pnRow, 2), 9, 1) = gMonedaExtranjera And Mid(FECtas.TextMatrix(pnRow, 7), 9, 1) = gMonedaNacional Then
                lnMontoTC = Round(CDbl(FECtas.TextMatrix(pnRow, 4)) * fnTipCambioV, 2)
            Else
                lnMontoTC = Round(CDbl(FECtas.TextMatrix(pnRow, 4)) / fnTipCambioC, 2)
            End If
        End If
        If oBase.ValidaSaldoCuenta(Trim(FECtas.TextMatrix(pnRow, 7)), lnMontoTC) = False Then
            FECtas.TextMatrix(pnRow, 1) = ""
            MsgBox "Cuenta de Ahorros no tiene saldo suficiente para pagar Cuota", vbInformation, "Aviso"
            Set oBase = Nothing
            Exit Sub
        End If
    End If
    nTotCred = 0
    nMontoAPagar = 0
    For i = 1 To FECtas.Rows - 1
        If FECtas.TextMatrix(i, 1) = "." Then
            nTotCred = nTotCred + 1
            nMontoAPagar = nMontoAPagar + CDbl(FECtas.TextMatrix(i, 4))
        End If
    Next i
    LblTotalPagar = Format(nMontoAPagar, "#0.00")
    Set oBase = Nothing
End Sub

Private Sub Form_Load()
    Dim oGen As COMDConstSistema.NCOMTipoCambio
    
    CentraForm Me
    txtfecha = gdFecSis
    'GetTipCambio gdFecSis
    Set oGen = New COMDConstSistema.NCOMTipoCambio
    fnTipCambioC = oGen.EmiteTipoCambio(gdFecSis, TCCompra)
    fnTipCambioV = oGen.EmiteTipoCambio(gdFecSis, TCVenta)

    Set oGen = Nothing
    
    Me.lblMensaje.Visible = False
End Sub

Private Sub optMon_Click(Index As Integer)
    FECtas.Clear
    FECtas.Rows = 2
    FECtas.FormaCabecera
    If optMon(0).value = True Then
        LblTotalPagar.BackColor = vbWhite
    Else
        LblTotalPagar.BackColor = &H80FF80
    End If
End Sub

Private Sub txtFecha_Click()
    FECtas.Clear
    FECtas.Rows = 2
    FECtas.FormaCabecera
End Sub
