VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPigUsoLinea 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Utilización de Linea"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   FillStyle       =   0  'Solid
   Icon            =   "frmPigUsoLinea.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   5542
      TabIndex        =   29
      Top             =   4185
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4350
      TabIndex        =   28
      Top             =   4185
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6735
      TabIndex        =   27
      Top             =   4185
      Width           =   975
   End
   Begin VB.Frame fraContenedor 
      Height          =   4050
      Index           =   0
      Left            =   15
      TabIndex        =   0
      Top             =   -30
      Width           =   7815
      Begin VB.CommandButton cmdBuscar 
         Height          =   405
         Left            =   7110
         Picture         =   "frmPigUsoLinea.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Buscar ..."
         Top             =   270
         Width           =   525
      End
      Begin VB.Frame fraContenedor 
         Height          =   1185
         Index           =   5
         Left            =   135
         TabIndex        =   16
         Top             =   2700
         Width           =   7515
         Begin MSComCtl2.DTPicker DtpNvoVcto 
            Height          =   315
            Left            =   3570
            TabIndex        =   33
            Top             =   780
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Format          =   56426497
            CurrentDate     =   37497
         End
         Begin VB.TextBox txtLineaDisponible 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6150
            TabIndex        =   32
            Top             =   765
            Width           =   1215
         End
         Begin VB.TextBox txtMontoTasacion 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   6150
            MaxLength       =   9
            TabIndex        =   20
            Top             =   270
            Width           =   1215
         End
         Begin VB.TextBox TxtSaldoCapitalNuevo 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            Height          =   315
            Left            =   1260
            TabIndex        =   19
            Top             =   270
            Width           =   1215
         End
         Begin VB.TextBox TxtTotalDeuda 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            ForeColor       =   &H80000012&
            Height          =   315
            Left            =   3570
            TabIndex        =   18
            Top             =   270
            Width           =   1215
         End
         Begin VB.TextBox txtPlazoNuevo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1260
            TabIndex        =   17
            Top             =   780
            Width           =   480
         End
         Begin VB.Label Label1 
            Caption         =   "Disponible S/."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   195
            Index           =   1
            Left            =   4860
            TabIndex        =   31
            Top             =   840
            Width           =   1230
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Nuevo Plazo"
            Height          =   210
            Index           =   15
            Left            =   120
            TabIndex        =   25
            Top             =   825
            Width           =   930
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Tasacion (-)"
            Height          =   195
            Index           =   12
            Left            =   4860
            TabIndex        =   24
            Top             =   345
            Width           =   1035
         End
         Begin VB.Label Label1 
            Caption         =   "Nvo Prestamo"
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   330
            Width           =   1080
         End
         Begin VB.Label Label2 
            Caption         =   "Nvo Vcto"
            Height          =   210
            Left            =   2565
            TabIndex        =   22
            Top             =   825
            Width           =   720
         End
         Begin VB.Label Label4 
            Caption         =   "Total Deuda"
            Height          =   195
            Left            =   2535
            TabIndex        =   21
            Top             =   345
            Width           =   960
         End
      End
      Begin VB.Frame fraContenedor 
         Height          =   1005
         Index           =   1
         Left            =   135
         TabIndex        =   14
         Top             =   705
         Width           =   7515
         Begin MSComctlLib.ListView lstClientes 
            Height          =   735
            Left            =   90
            TabIndex        =   15
            Top             =   180
            Width           =   7395
            _ExtentX        =   13044
            _ExtentY        =   1296
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483627
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Codigo"
               Object.Width           =   2470
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Cliente"
               Object.Width           =   5468
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Doc Ident."
               Object.Width           =   1765
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Tipo de Cliente"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame fraContenedor 
         Height          =   975
         Index           =   2
         Left            =   135
         TabIndex        =   2
         Top             =   1725
         Width           =   7515
         Begin VB.TextBox txtPlazoActual 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   6045
            MaxLength       =   2
            TabIndex        =   7
            Top             =   180
            Width           =   555
         End
         Begin VB.TextBox txtNroMvtos 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1620
            TabIndex        =   6
            Top             =   585
            Width           =   690
         End
         Begin VB.TextBox txtNroUsoLinea 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   3975
            TabIndex        =   5
            Top             =   585
            Width           =   690
         End
         Begin VB.TextBox txtDiasAtraso 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3975
            TabIndex        =   4
            Top             =   165
            Width           =   690
         End
         Begin VB.TextBox TxtVctoActual 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1620
            TabIndex        =   3
            Top             =   180
            Width           =   1005
         End
         Begin VB.Label Label5 
            Caption         =   "Fec.Vencimiento"
            Height          =   165
            Left            =   135
            TabIndex        =   13
            Top             =   255
            Width           =   1215
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Plazo Actual"
            Height          =   210
            Index           =   11
            Left            =   4770
            TabIndex        =   12
            Top             =   255
            Width           =   960
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Dias Atraso"
            Height          =   210
            Index           =   7
            Left            =   2730
            TabIndex        =   11
            Top             =   270
            Width           =   960
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Nro Amortizaciones"
            Height          =   225
            Index           =   8
            Left            =   120
            TabIndex        =   10
            Top             =   645
            Width           =   1395
         End
         Begin VB.Label lblMoneda 
            Height          =   255
            Left            =   2070
            TabIndex        =   9
            Top             =   210
            Width           =   255
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Usos de Linea"
            Height          =   195
            Index           =   9
            Left            =   2730
            TabIndex        =   8
            Top             =   675
            Width           =   1185
         End
      End
      Begin SICMACT.ActXCodCta AxCodCta 
         Height          =   405
         Left            =   135
         TabIndex        =   1
         Top             =   270
         Width           =   3630
         _extentx        =   6403
         _extenty        =   714
         enabledcmac     =   -1
         enabledcta      =   -1
         enabledprod     =   -1
         enabledage      =   -1
      End
   End
   Begin VB.Label lblMensaje 
      Caption         =   "<F12> Número Contrato Antiguo"
      Height          =   285
      Left            =   90
      TabIndex        =   30
      Top             =   4275
      Width           =   2370
   End
End
Attribute VB_Name = "frmPigUsoLinea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fnVarPlzMin As Integer
Dim fnVarPlzMax As Integer

Dim fnVarOpeCod As Long
Dim fsVarOpeDesc As String
Dim fsVarPersCodCMAC As String
Dim fsVarNombreCMAC As String

Dim fnVarTipoCliente As Integer

Dim fnVarTasaInteres As Currency
Dim fnVarValorTasacion As Currency
Dim fnVarTipoTasacion As Integer
Dim fnVarPlazo As Integer
Dim fdVarFecVencimiento As Date
Dim fdVarFecUltPago As Date
Dim fnVarEstado As ColocEstado

Dim fnVarPorcPrestamo As Currency
Dim fnVarLineaDisponible As Currency
Dim fnVarLineaDisponibleMaximo As Currency
Dim fnVarPorcConserv As Currency
Dim fnVarNewPlazo As Integer
Dim fsVarNewFecVencimiento As String
Dim fnVarPrestamo As Currency                    'Prestamo Original del Contrato Activo
Dim fnVarSaldoCap As Currency                    'Saldo del Prestamo del Contrato Activo
Dim fnVarNewSaldoCap As Currency              'Saldo final despues de la amortizacion o pago
Dim fnVarCapitalPagado As Currency             'Monto de Capital que se esta amortizando o pagando
Dim fnVarNewPrestamo As Currency              'Nuevo Prestamo que se otorgara con la Utilizacion de Linea
Dim fnVarIntCompensatorio As Currency
Dim fnVarIntMoratorio As Currency
Dim fnVarNewComServicio As Currency
Dim fnVarNewIntCompensatorio As Currency
Dim fnVarComServicio As Currency
Dim fnVarComPenalidad As Currency
Dim fnVarComVencida As Currency
Dim fnVarDerRemate As Currency
Dim fnVarPorcCapMin As Currency

Dim fnVarDiasAtraso As Double
Dim fnVarDiasCambCart  As Currency
Dim fnVarDiasIntereses As Currency
Dim fnVarDeuda As Currency

Dim fnVarUsoLineaNro As Integer
Dim fnVarMaterial As Integer
Dim fnVarConservacion As Integer
Dim fnVarPesoNeto As Currency
Dim fnVarPrecioMaterial As Currency
Dim fnVarTasacion As Currency
Dim fnVarTasacionAdic As Currency
Dim fnVarComTasacion As Currency

'*********
Dim fnVarMontoAPagar As Currency
Dim fnVarNroCalend As Integer
Dim fnVarNroCalendDesem As Integer
Dim lnNroTransac As Integer
Dim sPersCod As String, sNombre As String, sDireccion As String, sDocId As String, sTipoCuenta As String, sOperacion As String, sCuenta As String


Public Sub Inicio(ByVal pnOpeCod As Long, ByVal psOpeDesc As String, _
        ByVal psPersCodCMAC As String, ByVal psNomCMAC As String)
Dim loFunct As DPigFunciones
      
    fnVarOpeCod = pnOpeCod
    fsVarOpeDesc = psOpeDesc
    fsVarPersCodCMAC = psPersCodCMAC
    fsVarNombreCMAC = psNomCMAC

    Set loFunct = New DPigFunciones
         fnVarPlzMin = loFunct.GetParamValor(8015)
         fnVarPlzMax = loFunct.GetParamValor(8016)
    Set loFunct = Nothing
    Limpiar
    Me.Show 1
End Sub
Private Function IniciaLavDinero() As String
Dim i As Long
Dim nRelacion As CaptacRelacPersona
Dim nMonto As Double
Dim oPersona As NCapMantenimiento
Dim oDatos As DPigContrato
Dim rsPersPigno As Recordset
'Dim sPersCod As String, sNombre As String, sDireccion As String, sDocId As String, sTipoCuenta As String, sOperacion As String, sCuenta As String

Set oDatos = New DPigContrato
    Set rsPersPigno = oDatos.dClientePigno(AXCodCta.NroCuenta)
    sCuenta = AXCodCta.NroCuenta
    sPersCod = rsPersPigno("cPersCod")
    sNombre = rsPersPigno("cPersnombre")
    sDireccion = rsPersPigno("cPersDireccDomicilio")
    sDocId = rsPersPigno("cPersIdNro")
    sTipoCuenta = ""
    sOperacion = ""
    Set oDatos = Nothing

    nMonto = CDbl(txtLineaDisponible.Text)
    IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, False, True, nMonto, "", sOperacion, False, sTipoCuenta)
       
End Function

Private Sub Limpiar()
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio   ' Procedimiento que arma los primeros campos de la cuenta
    AXCodCta.Age = ""
    lstClientes.ListItems.Clear
    TxtVctoActual.Text = Format("  /  /  ", "")
    txtDiasAtraso.Text = Format(0, "0")
    txtNroMvtos = Format(0, "0")
    txtNroUsoLinea = Format(0, "0")
    txtPlazoNuevo.Text = Format(0, "0")
    txtPlazoActual.Text = Format(0, "0")
    txtTotalDeuda.Text = Format(0, "#0.00")
    txtSaldoCapitalNuevo.Text = Format(0, "#0.00")
    DtpNvoVcto.value = Format(gdFecSis, "dd/mm/yyyy")
    txtSaldoCapitalNuevo.Text = Format(0, "#0.00")
    txtLineaDisponible = Format(0, "#0.00")
    txtMontoTasacion = Format(0, "#0.00")

    fnVarTipoCliente = 0
    fnVarUsoLineaNro = 0
    fnVarTipoTasacion = 0
    fnVarValorTasacion = 0
    fnVarPlazo = 0
    fnVarEstado = 0
    fnVarPorcPrestamo = 0
    fnVarLineaDisponible = 0
    fnVarLineaDisponibleMaximo = 0
    fnVarPorcConserv = 0
    fnVarNewPlazo = 0
    fnVarPrestamo = 0
    fnVarSaldoCap = 0
    fnVarNewSaldoCap = 0
    fnVarCapitalPagado = 0
    fnVarNewPrestamo = 0
    fnVarIntCompensatorio = 0
    fnVarIntMoratorio = 0
    fnVarNewComServicio = 0
    fnVarNewIntCompensatorio = 0
    fnVarComServicio = 0
    fnVarComPenalidad = 0
    fnVarComVencida = 0
    fnVarDerRemate = 0
    fnVarDiasAtraso = 0
    fnVarDiasIntereses = 0
    fnVarDeuda = 0
    fnVarMaterial = 0
    fnVarConservacion = 0
    fnVarPesoNeto = 0
    fnVarPrecioMaterial = 0
    fnVarTasacion = 0
    fnVarTasacionAdic = 0
    fnVarComTasacion = 0
    fnVarMontoAPagar = 0
    fnVarNroCalend = 0

End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call BuscaContrato(AXCodCta.NroCuenta)
    End If
End Sub

'Busca el contrato ingresado
Private Sub BuscaContrato(ByVal psNroContrato As String)
Dim lbOk As Boolean
Dim lstTmpCliente As ListItem
Dim lrFunctVal As ADODB.Recordset
Dim lrValida As ADODB.Recordset
Dim lrCredPigPersonas As ADODB.Recordset
Dim lrCredPigDeuda As ADODB.Recordset
Dim lrCredPigJoyas As ADODB.Recordset
Dim loValContrato As nPigValida
Dim loMuestraDatos As DPigContrato
Dim loObtieneJoyas As DPigContrato
Dim loFunct As NPigCalculos
Dim loValores As DPigFunciones

Dim lnDeuda As Currency
Dim lnDiasAtraso  As Integer

Dim lsmensaje As String

    'Valida Contrato
    Set lrValida = New ADODB.Recordset
    Set loValContrato = New nPigValida
        Set lrValida = loValContrato.nValidaAmortizacionCredPignoraticio(psNroContrato, gdFecSis, 0, lsmensaje)
        If Trim(lsmensaje) <> "" Then
             MsgBox lsmensaje, vbInformation, "Aviso"
             Exit Sub
        End If
    Set loValContrato = Nothing
    
    If (lrValida Is Nothing) Then
        Limpiar
        Set lrValida = Nothing
        AXCodCta.SetFocusAge
        Exit Sub
    End If
    
    fnVarPlazo = lrValida!nPlazo
    fnVarPrestamo = Format(lrValida!nMontoCol, "#0.00")
    fnVarSaldoCap = Format(lrValida!nSaldo, "#0.00")
    fnVarEstado = lrValida!nPrdEstado
    fdVarFecVencimiento = Format(lrValida!dvenc, "dd/mm/yyyy")
    fdVarFecUltPago = Format(lrValida!dPrdEstado, "dd/mm/yyyy")
    fnVarDiasIntereses = DateDiff("d", fdVarFecUltPago, Format(gdFecSis, "dd/mm/yyyy"))
    fnVarNewPlazo = lrValida!nPlazo
    fnVarDiasAtraso = lrValida!nDiasAtraso
    fnVarValorTasacion = lrValida!totTasacion
    fnVarNroCalend = lrValida!nNumCalend
    fnVarNroCalendDesem = lrValida!nNumCalendDesem
    fnVarTasaInteres = lrValida!nTasaInteres
    fnVarUsoLineaNro = lrValida!nUsoLineaNro
    lnNroTransac = lrValida!nTransacc
    
    txtDiasAtraso.Text = lrValida!nDiasAtraso
    txtNroUsoLinea = Format(fnVarUsoLineaNro, "0")
    txtPlazoActual.Text = lrValida!nPlazo
    txtNroMvtos = Format(lrValida!nNroAmort, "0")
    TxtVctoActual = Format(lrValida!dvenc, "dd/mm/yyyy")
    
     Set lrCredPigPersonas = New ADODB.Recordset
     Set loMuestraDatos = New DPigContrato
         Set lrCredPigPersonas = loMuestraDatos.dObtieneDatosCreditoPignoraticioPersonas(psNroContrato)
        
     If lrCredPigPersonas.BOF And lrCredPigPersonas.EOF Then
        MsgBox " Error al mostrar datos del cliente ", vbCritical, " Aviso "
    Else
        lstClientes.ListItems.Clear
        Do While Not lrCredPigPersonas.EOF
            Set lstTmpCliente = lstClientes.ListItems.Add(, , Trim(lrCredPigPersonas!cPersCod))
                  lstTmpCliente.SubItems(1) = Trim(PstaNombre(lrCredPigPersonas!cPersNombre, False))
                  lstTmpCliente.SubItems(2) = Trim(IIf(IsNull(lrCredPigPersonas!NroDNI), "", lrCredPigPersonas!NroDNI))
                  lstTmpCliente.SubItems(3) = Trim(IIf(IsNull(lrCredPigPersonas!DescCalif), "", lrCredPigPersonas!DescCalif))
                  fnVarTipoCliente = IIf(IsNull(lrCredPigPersonas!cCalifiCliente), 0, lrCredPigPersonas!cCalifiCliente)
            lrCredPigPersonas.MoveNext
        Loop
    End If
  
    ' MostrarDeuda
    Set loFunct = New NPigCalculos
    
    Set lrCredPigDeuda = New ADODB.Recordset
    Set loMuestraDatos = New DPigContrato
         Set lrCredPigDeuda = loMuestraDatos.dObtieneDatosPignoraticioDeuda(psNroContrato)
       
     If lrCredPigDeuda.BOF And lrCredPigDeuda.EOF Then
        MsgBox " Error al mostrar Deuda del cliente ", vbCritical, " Aviso "
    Else
        fnVarSaldoCap = Format(lrCredPigDeuda!Capital, "#0.00")
        fnVarPorcCapMin = lrCredPigDeuda!PorcCapitalMinimo
        fnVarCapitalPagado = Format(lrCredPigDeuda!Capital, "#0.00")
       
        '************* RECALCULA INTERES COMPENSATORIO POR LOS DIAS TRANSCURRIDOS ***************
        fnVarIntCompensatorio = loFunct.nCalculaIntCompensatorio(fnVarSaldoCap, fnVarTasaInteres, fnVarDiasIntereses)
        '*************
        Set lrFunctVal = New ADODB.Recordset
        Set loValores = New DPigFunciones
             Set lrFunctVal = loValores.GetConceptoValor(Val(gColPigConceptoCodComiServ))
        fnVarComServicio = loFunct.nCalculaConcepto(lrFunctVal!nTpoValor, lrFunctVal!nValor, IIf(IsNull(lrFunctVal!nMontoMin), 0, lrFunctVal!nMontoMin), _
                                                                         IIf(IsNull(lrFunctVal!nMontoMax), 0, lrFunctVal!nMontoMax), fnVarSaldoCap)  'Comision para el nuevo Calendario
        fnVarIntMoratorio = Format(lrCredPigDeuda!IntMoratorio, "#0.00")
        fnVarComPenalidad = 0
        fnVarComVencida = Format(lrCredPigDeuda!ComVencida, "#0.00")
        fnVarDerRemate = Format(lrCredPigDeuda!DerRemate, "#0.00")
        fnVarDeuda = fnVarSaldoCap + fnVarIntCompensatorio + fnVarIntMoratorio + fnVarComServicio + _
                            fnVarComPenalidad + fnVarComVencida + fnVarDerRemate
        txtTotalDeuda.Text = Format(fnVarDeuda, "###,##0.00")
        fnVarNewSaldoCap = fnVarSaldoCap - fnVarCapitalPagado
        txtSaldoCapitalNuevo.Text = Format(fnVarNewSaldoCap, "#0.00")
        txtPlazoNuevo.Text = lrValida!nPlazo
        DtpNvoVcto.value = Format(gdFecSis + Format(txtPlazoNuevo.Text, "#0"), "dd/mm/yyyy")
     End If
    
    ' Calcula Nuevo Prestamo
    
    Set lrCredPigJoyas = New ADODB.Recordset
    Set loObtieneJoyas = New DPigContrato
         Set lrCredPigJoyas = loObtieneJoyas.dObtieneDatosPignoraticioJoyas(psNroContrato) 'Obtiene Caracteristicas de las joyas
       
     If lrCredPigJoyas.BOF And lrCredPigJoyas.EOF Then
        MsgBox " Error al Obtener datos de las Joyas ", vbCritical, " Aviso "
    Else
        fnVarTipoTasacion = lrCredPigJoyas!nTipoTasacion
        If fnVarTipoCliente = 1 Then              ' Cliente A1
           fnVarPorcPrestamo = loValores.GetParamValor(8001)
        ElseIf fnVarTipoCliente = 2 Then        ' Cliente A
           fnVarPorcPrestamo = loValores.GetParamValor(8002)
        ElseIf fnVarTipoCliente = 3 Then        ' Cliente B
           fnVarPorcPrestamo = loValores.GetParamValor(8003)
        ElseIf fnVarTipoCliente = 4 Then        ' Cliente B1
           fnVarPorcPrestamo = loValores.GetParamValor(8004)
        Else
             MsgBox "Error: El Tipo de Cliente No ha sido Considerado", vbInformation, " Aviso "
        End If
        
        Do While Not lrCredPigJoyas.EOF
             fnVarMaterial = lrCredPigJoyas!nMaterial
             fnVarConservacion = lrCredPigJoyas!nConservacion
             fnVarPesoNeto = lrCredPigJoyas!npesoneto
             fnVarTasacionAdic = fnVarTasacionAdic + lrCredPigJoyas!nTasacionAdicional
             If fnVarConservacion = 1 Then
                fnVarPorcConserv = loValores.GetParamValor(8010)          ' Conservacion de Joya Buena
             ElseIf fnVarConservacion = 2 Then
                fnVarPorcConserv = loValores.GetParamValor(8011)          ' Conservacion de Joya Regular
             ElseIf fnVarConservacion = 3 Then
                fnVarPorcConserv = loValores.GetParamValor(8012)          ' Conservacion de Joya Malo
             Else
                  MsgBox "Estado de Conservación de la Joya No ha sido Considerado", vbInformation, " Aviso "
             End If
             fnVarPrecioMaterial = loValores.GetPrecioMaterial(1, fnVarMaterial, 1)
             fnVarTasacion = fnVarTasacion + Round(loFunct.CalcValorTasacion(fnVarPorcConserv, fnVarPesoNeto, fnVarPrecioMaterial), 2)
             
             lrCredPigJoyas.MoveNext
        Loop
        
        fnVarNewPrestamo = Round(loFunct.CalcValorPrestamo(fnVarPorcPrestamo, fnVarTasacion + fnVarTasacionAdic), 2)
             
        Set lrFunctVal = loValores.GetConceptoValor(Val(gColPigConceptoCodTasacion))
        fnVarComTasacion = loFunct.nCalculaConcepto(lrFunctVal!nTpoValor, lrFunctVal!nValor, IIf(IsNull(lrFunctVal!nMontoMin), 0, lrFunctVal!nMontoMin), _
                                                                    IIf(IsNull(lrFunctVal!nMontoMax), 0, lrFunctVal!nMontoMax), fnVarSaldoCap)
        fnVarLineaDisponible = Format(fnVarNewPrestamo - fnVarDeuda - fnVarComTasacion, "###,##0.00")
        If fnVarLineaDisponible > 0 Then
            txtMontoTasacion.Text = Format(fnVarComTasacion, "###,##0.00")
            txtSaldoCapitalNuevo.Text = Format(fnVarNewPrestamo, "###,##0.00")
            txtLineaDisponible.Text = Format(fnVarLineaDisponible, "###,##0.00")
            fnVarLineaDisponibleMaximo = fnVarLineaDisponible
            AXCodCta.Enabled = False
        Else
             MsgBox "Contrato no posee Saldo Disponible", vbInformation, " Aviso "
             Limpiar
             AXCodCta.SetFocusAge
        End If
        
        cmdGrabar.SetFocus
    End If
    
    Set lrValida = Nothing
    Set lrCredPigDeuda = Nothing
    Set lrCredPigJoyas = Nothing
    Set lrCredPigPersonas = Nothing
    Set loFunct = Nothing
    Set loValores = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & Err.Number & " " & Err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdBuscar_Click()

Dim loPers As UPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersContrato As DColPContrato
Dim lrContratos As ADODB.Recordset
Dim loCuentas As UProdPersona

On Error GoTo ControlError

Set loPers = New UPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
Set loPers = Nothing

' Selecciona Estados
lsEstados = gPigEstDesemb & "," & gPigEstReusoLin & "," & gPigEstRemat & "," & gPigEstRematPRes & "," & gPigEstAmortiz

If Trim(lsPersCod) <> "" Then
    Set loPersContrato = New DColPContrato
        Set lrContratos = loPersContrato.dObtieneCredPigDePersona(lsPersCod, lsEstados, Mid(gsCodAge, 4, 2))
    Set loPersContrato = Nothing
End If

Set loCuentas = New UProdPersona
    Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrContratos)
    If loCuentas.sCtaCod <> "" Then
        AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        AXCodCta.SetFocusCuenta
    End If
Set loCuentas = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

'Cancela el proceso actual y permite inicializar ls variables para otro proceso
Private Sub cmdCancelar_Click()
    Limpiar
    cmdGrabar.Enabled = False
    txtPlazoNuevo.Enabled = False
    AXCodCta.Enabled = True
    AXCodCta.SetFocusCuenta
End Sub

Private Sub cmdGrabar_Click()

Dim loContFunct As NContFunciones
Dim loGrabaRL As NPigContrato
Dim oImpre As NPigImpre
Dim oPrevio As Previo.clsPrevio
Dim loCalc As NPigCalculos

Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim lsFechaVenc As String
Dim lnMontoTransaccion As Currency
Dim lsCadImpAmort As String
Dim lsCadImpDesemb As String
Dim lsNombreCliente As String
Dim lnMovNro As Long
Dim lnNewPagoMin As Currency

lsFechaVenc = Format$(gdFecSis + fnVarNewPlazo, "mm/dd/yyyy")
lnMontoTransaccion = CCur(Me.txtTotalDeuda.Text)
lsNombreCliente = lstClientes.ListItems(1).ListSubItems.Item(1)

If Not ValidaSDisponible Then Exit Sub

If MsgBox(" Grabar Uso de Linea? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
' **********************************************************************
   'Realiza la Validación para el Lavado de Dinero
    Dim clsLav As nCapDefinicion
    Dim nPorcRetCTS As Double, nMontoLavDinero As Double, nTC As Double
    Dim sPersLavDinero As String
    Dim nMonto As Double
        sPersLavDinero = ""
        Set clsLav = New nCapDefinicion
        nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
        Set clsLav = Nothing
        
           Dim clsTC As nTipoCambio
            Set clsTC = New nTipoCambio
            nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
            Set clsTC = Nothing
        nMonto = CDbl(txtLineaDisponible.Text)
        If nMonto >= Round(nMontoLavDinero * nTC, 2) Then
            sPersLavDinero = IniciaLavDinero()
            If sPersLavDinero = "" Then Exit Sub
        End If
    ' ********************************************************************
                    
        Set loCalc = New NPigCalculos
             fnVarNewIntCompensatorio = loCalc.nCalculaIntCompensatorio(fnVarNewSaldoCap, fnVarTasaInteres, fnVarNewPlazo)  'Obtiene Interes y Comisiones para el nuevo Capital
        cmdGrabar.Enabled = False
        
        Set loContFunct = New NContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        lnNewPagoMin = Format((fnVarPrestamo * fnVarPorcCapMin / 100), "#0.00") + fnVarNewIntCompensatorio + fnVarComServicio
        Set loGrabaRL = New NPigContrato
        
        'Grabar Amortizacion Pignoraticio
        lnMovNro = loGrabaRL.nUsoLineaCredPignoraticio(AXCodCta.NroCuenta, Format(fnVarNewSaldoCap, "#0.00"), lsFechaHoraGrab, _
                 lsMovNro, lsFechaVenc, fnVarNewPlazo, Format(lnMontoTransaccion, "#0.00"), Format(fnVarCapitalPagado, "#0.00"), _
                 Format(fnVarIntCompensatorio, "#0.00"), Format(fnVarIntMoratorio, "#0.00"), Format(fnVarComServicio, "#0.00"), _
                 Format(fnVarComPenalidad, "#0.00"), Format(fnVarComVencida, "#0.00"), Format(fnVarDerRemate, "#0.00"), _
                 fnVarDiasAtraso, Val(txtNroMvtos.Text) + 1, fnVarDiasCambCart, Format(fnVarValorTasacion, "#0.00"), fnVarOpeCod, _
                 Format(fnVarNewIntCompensatorio, "#0.00"), fsVarOpeDesc, fsVarPersCodCMAC, fnVarNroCalend, _
                 fnVarPlazo, fnVarEstado, Format(fnVarSaldoCap, "#0.00"), Format(fnVarNewPrestamo, "#0.00"), _
                 fnVarTipoTasacion, fnVarUsoLineaNro, fnVarNroCalendDesem, Format(fnVarPrestamo, "#0.00"), gdFecSis, sPersLavDinero, sPersCod)
        
        Set loGrabaRL = Nothing
        Set loCalc = Nothing
        'IMPRESION DE LAVADO DE DINERO
        If sPersLavDinero <> "" Then
          Dim oBoleta As NCapImpBoleta
          Set oBoleta = New NCapImpBoleta
           Do
               oBoleta.ImprimeBoletaLavadoDinero gsNomCmac, gsNomAge, gdFecSis, sCuenta, sNombre, sDocId, sDireccion, _
                        sNombre, sDocId, sDireccion, sNombre, sDocId, sDireccion, sOperacion, nMonto, sLpt, , , Trim(Left("", 15))
            Loop Until MsgBox("¿Desea reimprimir Boleta de Lavado de Dinero?", vbQuestion + vbYesNo, "Aviso") = vbNo
            Set oBoleta = Nothing
       End If
    
        Set oImpre = New NPigImpre
        'Amortizacion
'        Call oImpre.ImpreReciboAmortizacion(gsInstCmac, gsNomAge, lsFechaHoraGrab, AxCodCta.NroCuenta, lsNombreCliente, _
'                                                            lsFechaVenc, fnVarDiasAtraso, fnVarCapitalPagado, fnVarIntCompensatorio, _
'                                                            fnVarIntMoratorio, fnVarComServicio, fnVarDerRemate, fnVarComVencida, _
'                                                            lnMontoTransaccion, fnVarNewPrestamo, gsCodUser, lnNroTransac + 1, lnMovNro, _
'                                                            lnNewPagoMin, fnVarDiasIntereses, sLpt, " ")
                                                   
        Call oImpre.ImpreReciboDesembolsoRL(gsInstCmac, gsNomAge, lsFechaHoraGrab, AXCodCta.NroCuenta, lsNombreCliente, fnVarNewPrestamo, lnMontoTransaccion, _
                                                fnVarComServicio, fnVarLineaDisponible, gsCodUser, lnMovNro, lnNroTransac + 1, sLpt, "", fnVarValorTasacion)
                
        Do While MsgBox("Desea Reimprimir Comprobante de Reuso de Linea? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes
'            Call oImpre.ImpreReciboAmortizacion(gsInstCmac, gsNomAge, lsFechaHoraGrab, AxCodCta.NroCuenta, lsNombreCliente, _
'                                                                 lsFechaVenc, fnVarDiasAtraso, fnVarCapitalPagado, fnVarIntCompensatorio, _
'                                                                 fnVarIntMoratorio, fnVarComServicio, fnVarDerRemate, fnVarComVencida, _
'                                                                 lnMontoTransaccion, fnVarNewSaldoCap, gsCodUser, lnNroTransac + 1, _
'                                                                 lnMovNro, fnVarDiasIntereses, lnNewPagoMin, sLpt, " ")
              Call oImpre.ImpreReciboDesembolsoRL(gsInstCmac, gsNomAge, lsFechaHoraGrab, AXCodCta.NroCuenta, lsNombreCliente, fnVarNewPrestamo, lnMontoTransaccion, _
                                                fnVarComServicio, fnVarLineaDisponible, gsCodUser, lnMovNro, lnNroTransac + 1, sLpt, "", fnVarValorTasacion)
            
        Loop
        
        Set oImpre = Nothing
        Limpiar
        AXCodCta.Enabled = True
        AXCodCta.SetFocus
        
Else
    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And AXCodCta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gColConsuPrendario, False)
        If sCuenta <> "" Then
            AXCodCta.NroCuenta = sCuenta
            AXCodCta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
Dim oDatos As DPigFunciones

Set oDatos = New DPigFunciones
    fnVarDiasCambCart = oDatos.GetParamValor(gPigParamDiasCambioCartera)
Set oDatos = Nothing
AXCodCta.Texto = "Crédito"
AXCodCta.Age = ""
AXCodCta.EnabledProd = False
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub txtLineaDisponible_Change()
    cmdGrabar.Enabled = True
End Sub

Private Sub txtLineaDisponible_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If ValidaSDisponible Then
            If cmdGrabar.Enabled Then cmdGrabar.SetFocus
        Else
            If txtLineaDisponible.Enabled Then txtLineaDisponible.SetFocus
        End If
    End If
End Sub

Private Sub txtPlazoNuevo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Val(txtPlazoNuevo.Text) >= fnVarPlzMin And Val(txtPlazoNuevo.Text) <= fnVarPlzMax Then
           fnVarNewPlazo = Val(txtPlazoNuevo.Text)
           DtpNvoVcto.value = Format(DateAdd("d", fnVarNewPlazo, gdFecSis), "dd/mm/yyyy")
       Else
           MsgBox "Plazo Fuera del Rango Permitido", vbInformation, " Aviso "
       End If
    End If
End Sub

Private Sub DtpNvoVcto_Change()
    If DateDiff("d", gdFecSis, DtpNvoVcto.value) >= fnVarPlzMin And DateDiff("d", gdFecSis, DtpNvoVcto.value) <= fnVarPlzMax Then
        txtPlazoNuevo.Text = DateDiff("d", gdFecSis, DtpNvoVcto.value)
        fnVarNewPlazo = Val(txtPlazoNuevo.Text)
        txtLineaDisponible.SetFocus
    Else
        DtpNvoVcto.value = Format(DateAdd("d", fnVarNewPlazo, gdFecSis), "dd/mm/yyyy")
        MsgBox "Plazo Fuera del Rango Permitido", vbInformation, " Aviso "
    End If
End Sub

Private Function ValidaSDisponible() As Boolean

    If Val(txtLineaDisponible.Text) > 0 And Val(txtLineaDisponible.Text) <= fnVarLineaDisponibleMaximo Then
        txtLineaDisponible.Text = Format(txtLineaDisponible.Text, "########,###.00")
        fnVarLineaDisponible = CCur(txtLineaDisponible.Text)
        fnVarNewPrestamo = fnVarLineaDisponible + fnVarDeuda + fnVarComTasacion
        txtSaldoCapitalNuevo.Text = Format(fnVarNewPrestamo, "###,##0.00")
        ValidaSDisponible = True
    Else
        txtLineaDisponible.Text = Format(fnVarLineaDisponibleMaximo, "########,###.00")
        fnVarLineaDisponible = CCur(txtLineaDisponible.Text)
        fnVarNewPrestamo = fnVarLineaDisponible + fnVarDeuda + fnVarComTasacion
        txtSaldoCapitalNuevo.Text = Format(fnVarNewPrestamo, "###,##0.00")
        ValidaSDisponible = False
        MsgBox "Linea disponible fuera del Rango Permitido para el Contrato...", vbInformation, " Aviso "
    End If
    
End Function

