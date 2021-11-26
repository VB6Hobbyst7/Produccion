VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPigExtornoOpe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extorno de Operaciones de Pignoraticios"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9405
   Icon            =   "frmPigExtornoOpe.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6255
      TabIndex        =   6
      Top             =   210
      Width           =   1005
   End
   Begin VB.Frame fraBuscar 
      Caption         =   "Buscar Por"
      Height          =   1125
      Left            =   120
      TabIndex        =   3
      Top             =   75
      Width           =   1815
      Begin VB.OptionButton opt 
         Caption         =   "Nro Cuenta"
         Height          =   195
         Index           =   0
         Left            =   285
         TabIndex        =   5
         Top             =   360
         Width           =   1245
      End
      Begin VB.OptionButton opt 
         Caption         =   "General"
         Height          =   195
         Index           =   3
         Left            =   300
         TabIndex        =   4
         Top             =   720
         Value           =   -1  'True
         Width           =   1245
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   7575
      TabIndex        =   0
      Top             =   -15
      Width           =   1695
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1245
      End
      Begin VB.CommandButton cmdExtorno 
         Caption         =   "&Extornar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1245
      End
   End
   Begin MSComctlLib.ListView lstExtorno 
      Height          =   3540
      Left            =   60
      TabIndex        =   7
      Top             =   1275
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   6244
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1335
      Top             =   585
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   17
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPigExtornoOpe.frx":030A
            Key             =   "Cuenta"
         EndProperty
      EndProperty
   End
   Begin SICMACT.ActXCodCta AXCodCta 
      Height          =   375
      Left            =   2295
      TabIndex        =   8
      Top             =   225
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      Texto           =   "Crédito"
      EnabledAge      =   -1  'True
   End
   Begin VB.Label lblMensaje 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ADVERTENCIA!. Verifique que la Operación sea la correcta. Una vez ejecutado el Extorno no es posible su recuperación."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   540
      Left            =   2730
      TabIndex        =   9
      Top             =   4905
      Width           =   6600
   End
End
Attribute VB_Name = "frmPigExtornoOpe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''********************************************************************************
''* EXTORNO DE OPERACIONES DE PIGNORATICIO - LIMA
''********************************************************************************
'Option Explicit
'Dim vNroContrato As String
'Dim vPosExtorno As Integer
'
'Dim fsFechaTransac As String
'Dim fnVarOpeCod As Long
'Dim fsVarOpeDesc As String
'
'Public Sub Inicio(ByVal pnOpeCod As Long, ByVal psOpeDesc As String, '        ByVal psPersCodCMAC As String, ByVal psNomCmac As String)
'
'    fnVarOpeCod = pnOpeCod
'    fsVarOpeDesc = psOpeDesc
'
'    Me.Caption = psOpeDesc
'    Me.Show 1
'
'End Sub
'
'Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then cmdBuscar.SetFocus
'End Sub
'
'Private Sub cmdBuscar_Click()
'
'Dim lrBusca As ADODB.Recordset
'Dim loValContrato As nPigValida
'Dim lnOpeCod As Long
'Dim lsmensaje As String
'
''On Error GoTo ControlError
'
'    'Valida Contrato
'    'Limpiar
'    Set lrBusca = New ADODB.Recordset
'    Set loValContrato = New nPigValida
'
'    Select Case fnVarOpeCod
'
'    Case 159200     'Extorno de Desembolso
'        lnOpeCod = 1502
'    Case 159300     'Extorno de Amortizacion
'        lnOpeCod = 1509
'    Case 159400     'Extorno de Cancelacion
'        lnOpeCod = 1510
'    Case 159500     'Extorno de Uso de Linea
'        lnOpeCod = 1506
'    Case 159600     'Extorno de Cobranza de Custodia
'        lnOpeCod = 1513
'    Case 159900     'Extorno de Pago de Sobrante
'        lnOpeCod = 1521
'    Case 159901     'Extorno de Pago de Sobrante de Piezas
'        lnOpeCod = 1529
'    Case 159100     'Extorno de Rescate de Joya'
'        lnOpeCod = 1512
'    End Select
'
'    If Me.opt(0).value = True Then ' Busca por Codigo
'       Set lrBusca = loValContrato.nBuscaOperacionesCredPigParaExtorno(fsFechaTransac, lnOpeCod, AXCodCta.NroCuenta, gsCodAge)
'    Else
'       Set lrBusca = loValContrato.nBuscaOperacionesCredPigParaExtorno(fsFechaTransac, lnOpeCod, , gsCodAge)
'    End If
'    If Trim(lsmensaje) <> "" Then
'         MsgBox lsmensaje, vbInformation, "Aviso"
'         Exit Sub
'    End If
'    Set loValContrato = Nothing
'
'    If lrBusca Is Nothing Then ' Hubo un Error
'        Set lrBusca = Nothing
'        Exit Sub
'    End If
'    If lrBusca.BOF And lrBusca.EOF Then
'        MsgBox "No Existen Operaciones para EXTORNAR", vbInformation, "Aviso"
'        Exit Sub
'    End If
'
'    lstExtorno.ListItems.Clear
'
'    Call LLenaLista(lrBusca)
'
'    Set lrBusca = Nothing
'
'Exit Sub
'
'ControlError:   ' Rutina de control de errores.
'    MsgBox "Error: " & Err.Number & " " & Err.Description & vbCr & '        "Avise al Area de Sistemas ", vbInformation, " Aviso "
'
'End Sub
'
'Private Sub cmdExtorno_Click()
''On Error GoTo ControlError
'
'Dim loContFunct As NContFunciones
'Dim loGrabarExt As NPigContrato
'Dim oDatos As DPigContrato
'
'Dim lsMovNro As String
'Dim lsFechaHoraGrab As String
'
'Dim lsNroContrato As String
'Dim lsOperacion As String
'Dim lnMovNroAExt As Long, lnMovAmort As Long, lnMovDesem As Long
'Dim lnSaldo As Currency
'Dim lnMonto As Currency
'Dim Fecha As String
'Dim lnNroCalen As Integer
'Dim oImprime As NPigImpre
'Dim lsDesOpe As String * 28
'Dim lnSaldoCap As Currency
'Dim lsCliente As String
'Dim lnNroTransac As Integer
'Dim lnMovExt As Long
'Dim oValida As nPigValida
'Dim lnMovNro As Long
'Dim lsmensaje As String
'
'If lstExtorno.ListItems.Count = 0 Then
'    cmdExtorno.Enabled = False
'    Exit Sub
'End If
'
'If lstExtorno.SelectedItem.SubItems(2) <> "1" Then
'    MsgBox " Debe Extornar el último movimiento del Contrato ", vbInformation, " Aviso "
'    Exit Sub
'Else
'    If Right(lstExtorno.SelectedItem.ListSubItems(1), 6) = gPigOpeReusoLinea Then
'        lnMovExt = Val(lstExtorno.SelectedItem.ListSubItems(5)) + 2
'    Else
'        lnMovExt = Val(lstExtorno.SelectedItem.ListSubItems(5))
'    End If
'
'    Set oValida = New nPigValida
'    If oValida.ValidaSinoOpePrevias(lnMovExt, Trim(lstExtorno.SelectedItem)) = True Then
'        MsgBox "No se puede realizar el extorno. Contrato posee operaciones posteriores, Por favor Verifique", vbInformation, "Aviso"
'        Exit Sub
'    End If
'    Set oValida = Nothing
'    If MsgBox(" Esta Ud seguro de Extornar dicha Operación ? ", vbQuestion + vbYesNo + vbDefaultButton2, " Aviso ") = vbNo Then
'        Exit Sub
'    Else
'        MsgBox " Prepare la impresora para imprimir " & vbCr & '        " el recibo del Extorno", vbInformation, " Aviso "
'    End If
'End If
'
''*** Obtiene Datos de Operacion
'lsNroContrato = Trim(lstExtorno.SelectedItem)
'lsOperacion = Right(lstExtorno.SelectedItem.ListSubItems(1), 6)
'lnMovNroAExt = Val(lstExtorno.SelectedItem.ListSubItems(5))
'lnMonto = CCur(lstExtorno.SelectedItem.ListSubItems(3))
'Fecha = lstExtorno.SelectedItem.ListSubItems(4)
'lnNroCalen = IIf(lstExtorno.SelectedItem.ListSubItems(7) = "", 0, lstExtorno.SelectedItem.ListSubItems(7))
'lnSaldoCap = lstExtorno.SelectedItem.ListSubItems(8)
'lsCliente = lstExtorno.SelectedItem.ListSubItems(9)
'lnNroTransac = lstExtorno.SelectedItem.ListSubItems(10) + 1
'
''*** Genera el Mov Nro
'Set loContFunct = New NContFunciones
'    lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'Set loContFunct = Nothing
'
'lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
'
'Set loGrabarExt = New NPigContrato
'
'    Select Case lsOperacion
'        '** Extornar un DESEMBOLSO PIGNORATICIO
'        Case gPigOpeDesembolsoEFE
'            lnMovNro = loGrabarExt.nExtornoDesembolsoCredPig(lsNroContrato, lsFechaHoraGrab, '                 lsMovNro, lnMovNroAExt, lnMonto, lnNroCalen, False)
'            lsDesOpe = Space(4) & "EXTORNO DE DESEMBOLSO" & Space(3)
'        '** Extornar una AMORTIZACION PIGNORATICIO
'        Case gPigOpeAmortNorEFE, gPigOpeAmortMorEFE
'            lnMovNro = loGrabarExt.nExtornoAmortizacionCredPig(lsNroContrato, lsFechaHoraGrab, '                 lsMovNro, lnMovNroAExt, lnMonto, lnNroCalen, False)
'            lsDesOpe = Space(3) & "EXTORNO DE AMORTIZACION" & Space(2)
'
'        '** Extornar una CANCELACION PIGNORATICIO
'        Case gPigOpeCancelNorEFE, gPigOpeCancelMorEFE
'            lnMovNro = loGrabarExt.nExtornoCancelacionCredPig(lsNroContrato, lsFechaHoraGrab, '                 lsMovNro, lnMovNroAExt, lnMonto, lnNroCalen, False)
'            lsDesOpe = Space(3) & "EXTORNO DE CANCELACION" & Space(3)
'
'        '** Extornar una UTILIZACION DE LINEA
'        Case gPigOpeReusoLinea
'
'            Set oDatos = New DPigContrato
'                lnMovAmort = oDatos.dObtieneMovExtornoUsoLinea(lnMovNroAExt, gPigOpeAmortizEFE)
'                lnMovDesem = oDatos.dObtieneMovExtornoUsoLinea(lnMovNroAExt, gPigOpeDesembolso)
'            Set oDatos = Nothing
'
'            lnMovNro = loGrabarExt.nExtornoUsoLineaCredPig(lsNroContrato, lsFechaHoraGrab, '                 lsMovNro, lnMovNroAExt, lnMovAmort, lnMovDesem, lnMonto, lnNroCalen)
'            lsDesOpe = Space(3) & "EXTORNO DE USO DE LINEA" & Space(2)
'
'        '** Extornar un PAGO DE SOBRANTES
'        Case gPigOpePagoSobrantes
'
'            Dim oPigGrabaExt As NPigRemate
'
'            Set oPigGrabaExt = New NPigRemate
'            lnMovNro = oPigGrabaExt.ExtornoPagoSobrante(lnMovNroAExt, lsMovNro, gdFecSis, gsCodAge, gsCodUser)
'            Set oPigGrabaExt = Nothing
'
'            lnSaldoCap = lnMonto
'
'        '** Extornar un DUPLICADO CONTRATO PIG
'        Case geColPImpDuplicado
'        '    Call loGrabarExt.nExtornoDuplicadoContratoCredPig(lsNroContrato, lsFechaHoraGrab, '        '         lsMovNro, lnMovNroAExt, lnMonto, False)
'
'        '** Extornar una DEVOLUCION DE JOYAS
'        Case geColPDevJoyas
'        '    Call loGrabarExt.nExtornoRescateJoyaCredPig(lsNroContrato, lsFechaHoraGrab, '        '         lsMovNro, lnMovNroAExt, lnMonto, False)
'
'        '** Extornar COBRO DE CUSTODIA DIFERIDA
'        Case geColPCobCusDiferida
'        '    Call loGrabarExt.nExtornoCustodiaDiferidaCredPig(lsNroContrato, lsFechaHoraGrab, '        '         lsMovNro, lnMovNroAExt, lnMonto, False)
'
'        '** Extorno VENTA EN REMATE
'        Case geColPVtaRemate
''            Call loGrabarExt.nExtornoCustodiaDiferidaCredPig(lsNroContrato, lsFechaHoraGrab, ''                 lsMovNro, lnMovNroAExt, lnMonto, False)
'
'
'      '** Extorno Rescate de Joya'
'        Case gPigOpeDevJoyas
'             Call loGrabarExt.nExtornoRescateJoyas(lsNroContrato, lsFechaHoraGrab, '                  lsMovNro, lnMovNroAExt, lnMonto, False)
'                  lsDesOpe = Space(1) & "EXTORNO DE RESCATE DE JOYA" & Space(2)
'
'        '** Extorno Rescate de Sobrante de Piezas'
'        Case gPigOpeSobraPieza
'             Call loGrabarExt.nExtornoRescateSobrantePiezas(lsNroContrato, lsFechaHoraGrab, '                  lsMovNro, lnMovNroAExt, lnMonto, gsCodAge, False)
'                  lsDesOpe = Space(1) & "EXTORNO DE RESCATE DE SOBRANTE DE PIEZAS" & Space(2)
'
'    End Select
'
'    Set oImprime = New NPigImpre
'    Call oImprime.ImprimeBoletaExtorno(gsInstCmac, gsNomAge, lsDesOpe, lsNroContrato, lsCliente, '                                       lsFechaHoraGrab, lnNroTransac, lnMonto, lnSaldoCap, gsCodUser, lnMovNro, sLpt)
'
'    Do While MsgBox("Desea Reimprimir Comprobante de Extorno ? ", vbYesNo + vbQuestion + vbDefaultButton1, "Aviso") = vbYes
'        Call oImprime.ImprimeBoletaExtorno(gsInstCmac, gsNomAge, lsDesOpe, lsNroContrato, lsCliente, '                                       lsFechaHoraGrab, lnNroTransac, lnMonto, lnSaldoCap, gsCodUser, lnMovNro, sLpt)
'    Loop
'
'    Set oImprime = Nothing
'
'    Set loGrabarExt = Nothing
'
'    Me.lstExtorno.ListItems.Clear
'    If lstExtorno.ListItems.Count = 0 Then
'        cmdExtorno.Enabled = False
'    End If
'
'    Exit Sub
'
'ControlError:   ' Rutina de control de errores.
'    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & '        " Avise al Area de Sistemas ", vbInformation, " Aviso "
'
'End Sub
'
'Private Sub cmdsalir_Click()
'    Unload Me
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF12 And AXCodCta.Enabled = True Then 'F12
'        Dim sCuenta As String
'        sCuenta = frmValTarCodAnt.Inicia(gColConsuPrendario, False)
'        If sCuenta <> "" Then
'            AXCodCta.NroCuenta = sCuenta
'            AXCodCta.SetFocusCuenta
'        End If
'    End If
'End Sub
'
'Private Sub Form_Load()
'
'    fsFechaTransac = Mid(Format$(gdFecSis, "dd/mm/yyyy"), 7, 4) & Mid(Format$(gdFecSis, "dd/mm/yyyy"), 4, 2) & Mid(Format$(gdFecSis, "dd/mm/yyyy"), 1, 2)
'
'    lstExtorno.ColumnHeaders.Add , , "NroCuenta", 2000
'    lstExtorno.ColumnHeaders.Add , , "Operación", 2200
'    lstExtorno.ColumnHeaders.Add , , "OpcExt.", 750, lvwColumnCenter
'    lstExtorno.ColumnHeaders.Add , , "Monto", 1100, lvwColumnRight
'    lstExtorno.ColumnHeaders.Add , , "Fecha de Movimiento", 1750, lvwColumnCenter
'    lstExtorno.ColumnHeaders.Add , , "N°Tran", 800, lvwColumnCenter
'    lstExtorno.ColumnHeaders.Add , , "Usuario", 800, lvwColumnCenter
'    lstExtorno.ColumnHeaders.Add , , "N°Calen", 800, lvwColumnCenter
'    lstExtorno.ColumnHeaders.Add , , "SaldoCap", 1100, lvwColumnCenter
'    lstExtorno.ColumnHeaders.Add , , "Nombre", 0, lvwColumnCenter
'    lstExtorno.ColumnHeaders.Add , , "Transac", 0, lvwColumnCenter
'
'    lstExtorno.View = lvwReport
'    Limpiar
'    Me.Icon = LoadPicture(App.path & "\bmps\cm.ico")
'End Sub
'
'Private Sub LLenaLista(myRs As Recordset)
'Dim litmX As ListItem
'Dim lsCtaCodAnterior As String
'
'Do While Not myRs.EOF
'    Set litmX = lstExtorno.ListItems.Add(, , myRs!cCtaCod, , "Cuenta")           'Nro de Cred Pig
'        litmX.SubItems(1) = Mid(myRs!cOpedesc, 1, 30) & Space(10) & myRs!cOpecod 'Operacion
'        litmX.SubItems(3) = Format(myRs!nMonto, "#0.00")                         'Monto Operacion
'        litmX.SubItems(4) = fgFechaHoraGrab(myRs!cmovnro)                        'Fecha/hora Operacion
'        litmX.SubItems(5) = Str(myRs!nmovnro)                                    'Nro Movimiento(nMovNro)
'        litmX.SubItems(6) = Mid(myRs!cmovnro, 22, 4)                             'Usuario
'        litmX.SubItems(7) = Format(myRs!NroCalen, "0")                          'Nro.Calendario
'        litmX.SubItems(8) = Format(myRs!nSaldoCap, "##0.00")                     'Saldo Capital
'        litmX.SubItems(9) = PstaNombre(myRs!cPersNombre)                        'Cliente
'        litmX.SubItems(10) = myRs!nTransacc                                     'Nro de Transaccion
'
'    If myRs!cCtaCod = lsCtaCodAnterior Then
'        litmX.SubItems(2) = "0"
'    Else
'        litmX.SubItems(2) = "1"
'    End If
'    lsCtaCodAnterior = myRs!cCtaCod
'    myRs.MoveNext
'Loop
'
'End Sub
'
''Valida el ListView lstExtorno
'Private Sub lstExtorno_GotFocus()
'If lstExtorno.ListItems.Count >= 0 Then
'   cmdExtorno.Enabled = True
'End If
'End Sub
'
'Private Sub lstExtorno_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'     If lstExtorno.ListItems.Count > 0 Then
'        cmdExtorno.Enabled = True
'        cmdExtorno.SetFocus
'     End If
'End If
'End Sub
'
'Private Sub opt_Click(Index As Integer)
'Limpiar
'
'Select Case Index
'    Case 0
'        AXCodCta.Visible = True
'        AXCodCta.EnabledCta = True
'        AXCodCta.SetFocusCuenta
'    Case 3
'        AXCodCta.Visible = False
'        cmdBuscar.SetFocus
'End Select
'cmdBuscar.Visible = True
'End Sub
'
''Inicializa variables
'Private Sub Limpiar()
'    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
'    lstExtorno.ListItems.Clear
'End Sub
Private Sub cmdExtorno_Click()

End Sub
