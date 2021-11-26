VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPigAmortizacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Amortización de Contratos"
   ClientHeight    =   4965
   ClientLeft      =   3615
   ClientTop       =   2925
   ClientWidth     =   7860
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRebajar 
      Caption         =   "&Rebajar"
      Height          =   345
      Left            =   2640
      TabIndex        =   35
      Top             =   4470
      Width           =   975
   End
   Begin MSComCtl2.DTPicker DtpNvoVcto 
      Height          =   315
      Left            =   3780
      TabIndex        =   34
      Top             =   3420
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      CalendarTrailingForeColor=   -2147483635
      Format          =   64487425
      CurrentDate     =   37492
   End
   Begin VB.Frame fraContenedor 
      Height          =   4425
      Index           =   0
      Left            =   45
      TabIndex        =   3
      Top             =   -30
      Width           =   7740
      Begin VB.Frame fraContenedor 
         Height          =   975
         Index           =   2
         Left            =   105
         TabIndex        =   20
         Top             =   1740
         Width           =   7515
         Begin VB.TextBox txtLineaDisponible 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   6060
            TabIndex        =   32
            Top             =   585
            Width           =   1305
         End
         Begin VB.TextBox TxtVctoActual 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1620
            TabIndex        =   31
            Top             =   240
            Width           =   1005
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
            Height          =   300
            Left            =   3975
            TabIndex        =   30
            Top             =   225
            Width           =   690
         End
         Begin VB.TextBox txtNroUsoLinea 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   3975
            TabIndex        =   23
            Top             =   585
            Width           =   690
         End
         Begin VB.TextBox txtNroMvtos 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1620
            TabIndex        =   22
            Top             =   585
            Width           =   690
         End
         Begin VB.TextBox txtPlazoActual 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   6060
            MaxLength       =   2
            TabIndex        =   21
            Top             =   240
            Width           =   555
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
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   1
            Left            =   4770
            TabIndex        =   33
            Top             =   645
            Width           =   1230
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Usos de Linea"
            Height          =   195
            Index           =   9
            Left            =   2730
            TabIndex        =   29
            Top             =   675
            Width           =   1185
         End
         Begin VB.Label lblMoneda 
            Height          =   255
            Left            =   2070
            TabIndex        =   28
            Top             =   210
            Width           =   255
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Nro Amortizaciones"
            Height          =   225
            Index           =   8
            Left            =   120
            TabIndex        =   27
            Top             =   645
            Width           =   1395
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Dias Atraso"
            Height          =   210
            Index           =   7
            Left            =   2730
            TabIndex        =   26
            Top             =   270
            Width           =   960
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Plazo Actual"
            Height          =   210
            Index           =   11
            Left            =   4770
            TabIndex        =   25
            Top             =   255
            Width           =   960
         End
         Begin VB.Label Label5 
            Caption         =   "Fec.Vencimiento"
            Height          =   165
            Left            =   135
            TabIndex        =   24
            Top             =   255
            Width           =   1215
         End
      End
      Begin VB.Frame fraContenedor 
         Height          =   1005
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   7515
         Begin MSComctlLib.ListView lstClientes 
            Height          =   735
            Left            =   75
            TabIndex        =   19
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
      Begin SICMACT.ActXCodCta AxCodCta 
         Height          =   405
         Left            =   135
         TabIndex        =   17
         Top             =   270
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   714
         Texto           =   "Credito"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Frame fraContenedor 
         Height          =   1635
         Index           =   5
         Left            =   105
         TabIndex        =   5
         Top             =   2700
         Width           =   7515
         Begin VB.TextBox TxtApagarConITF 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   6165
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   40
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox TxtITF 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3630
            TabIndex        =   37
            Top             =   1185
            Width           =   1005
         End
         Begin VB.TextBox txtPlazoNuevo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1080
            TabIndex        =   36
            Top             =   750
            Width           =   480
         End
         Begin VB.TextBox TxtTotalDeuda 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            ForeColor       =   &H80000011&
            Height          =   285
            Left            =   1065
            TabIndex        =   16
            Top             =   375
            Width           =   1215
         End
         Begin VB.TextBox TxtMontoMinimoPagar 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   3630
            TabIndex        =   14
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox TxtSaldoCapitalNuevo 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   6165
            TabIndex        =   11
            Top             =   750
            Width           =   1215
         End
         Begin VB.TextBox txtMontoPagar 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   1305
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   6
            Top             =   1170
            Width           =   1215
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "A Pagar :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   5100
            TabIndex        =   39
            Top             =   1245
            Width           =   855
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "I.T.F."
            Height          =   195
            Index           =   0
            Left            =   2685
            TabIndex        =   38
            Top             =   1215
            Width           =   510
         End
         Begin VB.Label Label4 
            Caption         =   "Total Deuda"
            Height          =   195
            Left            =   105
            TabIndex        =   15
            Top             =   435
            Width           =   960
         End
         Begin VB.Label Label3 
            Caption         =   "Pago Rebajado"
            Height          =   210
            Left            =   2460
            TabIndex        =   13
            Top             =   420
            Width           =   1185
         End
         Begin VB.Label Label2 
            Caption         =   "Nvo Vcto"
            Height          =   210
            Left            =   2475
            TabIndex        =   12
            Top             =   825
            Width           =   720
         End
         Begin VB.Label Label1 
            Caption         =   "Nuevo Capital"
            Height          =   210
            Index           =   0
            Left            =   5070
            TabIndex        =   10
            Top             =   810
            Width           =   1080
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "A Amortizar S/."
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   8
            Top             =   1215
            Width           =   1230
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Nuevo Plazo"
            Height          =   210
            Index           =   15
            Left            =   120
            TabIndex        =   7
            Top             =   795
            Width           =   930
         End
      End
      Begin VB.CommandButton cmdBuscar 
         Height          =   405
         Left            =   7050
         Picture         =   "frmPigAmortizacion.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Buscar ..."
         Top             =   270
         Width           =   525
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   6735
      TabIndex        =   2
      Top             =   4470
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4380
      TabIndex        =   1
      Top             =   4470
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   5580
      TabIndex        =   0
      Top             =   4470
      Width           =   975
   End
   Begin VB.Label lblMensaje 
      Caption         =   "<F12> Número Contrato Antiguo"
      Height          =   285
      Left            =   90
      TabIndex        =   9
      Top             =   4560
      Width           =   2370
   End
End
Attribute VB_Name = "frmPigAmortizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************************'
'* AMORTIZACION DE CONTRATO DE PIGNORATICIO                                    *'
'* Archivo   :  frmPigAmortizacion.frm                                                              *'
'* Resumen:  Nos permite amortizar un contrato cambiando otorgandolo un nuevo *'
'*                 Plazo de Vencimiento por el saldo del credito                                  *'
'***************************************************************************************'
Option Explicit

Dim fnVarPlzMin As Integer
Dim fnVarPlzMax As Integer

Dim fnVarOpeCod As Long
Dim fsVarOpeDesc As String
Dim fsVarPersCodCMAC As String
Dim fsVarNombreCMAC As String

Dim fnVarPorcCapMin As Currency
Dim fnVarTasaInteres As Currency
Dim fnVarValorTasacion As Currency
Dim fnVarPlazo As Integer
Dim fdVarFecVencimiento As Date
Dim fdVarFecUltPago As Date
Dim fnVarEstado As ColocEstado
Dim lnNroTransac As Integer

Dim fnVarPrestamo As Currency
Dim fnVarSaldoCap As Currency
Dim fnVarNewSaldoCap As Currency
Dim fnVarNewPlazo As Integer
Dim fsVarNewFecVencimiento As String
Dim fnVarCapitalPagado As Currency   ' Capital a Pagar
Dim fnVarIntCompensatorio As Currency
Dim fnVarIntMoratorio As Currency
Dim fnVarNewComServicio As Currency
Dim fnVarNewIntCompensatorio As Currency
Dim fnVarComServicio As Currency
Dim fnVarComPenalidad As Currency
Dim fnVarComVencida As Currency
Dim fnVarDerRemate As Currency

Dim fnVarDiasAtraso As Double
Dim fnVarDiasCambCart  As Currency
Dim fnVarDiasIntereses As Currency
Dim fnVarDeuda As Currency

'LAVADO DE DINERO
Dim sPersCod As String, sNombre As String, sDireccion As String, sDocId As String
Dim sTipoCuenta As String, sOperacion As String, sCuenta As String
'*********
Dim fnVarMontoMinimo As Currency
Dim fnVarMontoAPagar As Currency
Dim fnVarCapMinimo As Currency        'Capital Minimo, almacena del Codigo 8017 del ColocParametro

Dim fnVarNroCalend As Integer

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

Private Sub Limpiar()
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio   ' Procedimiento que arma los primeros campos de la cuenta
    lstClientes.ListItems.Clear
    TxtVctoActual.Text = Format("  /  /  ", "")
    txtDiasAtraso.Text = Format(0, "0")
    txtNroMvtos = Format(0, "0")
    txtNroUsoLinea.Text = Format(0, "0")
    txtPlazoNuevo.Text = Format(0, "0")
    txtMontoPagar.Text = Format(0, "#0.00")
    txtTotalDeuda.Text = Format(0, "#0.00")
    txtMontoMinimoPagar.Text = Format(0, "#0.00")
    txtSaldoCapitalNuevo.Text = Format(0, "#0.00")
    txtLineaDisponible = ""
    DtpNvoVcto.value = Format(gdFecSis, "dd/mm/yyyy")
    txtSaldoCapitalNuevo.Text = Format(0, "#0.00")
    AXCodCta.Age = ""
    'If AxCodCta.EnabledAge Then AxCodCta.SetFocusAge
    
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call BuscaContrato(AXCodCta.NroCuenta)
    End If
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
    sOperacion = "PIGNORATICIO"
    Set oDatos = Nothing
    nMonto = CDbl(Me.txtMontoPagar.Text)
    IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, False, True, nMonto, "", sOperacion, False, sTipoCuenta)
       
End Function

Private Function EsExoneradaLavadoDinero() As Boolean
Dim bExito As Boolean
Dim clsExo As NCapServicios
bExito = True

    Set clsExo = New NCapServicios
    
    If Not clsExo.EsPersonaExoneradaLavadoDinero(sPersCod) Then bExito = False

    Set clsExo = Nothing
    EsExoneradaLavadoDinero = bExito
    
End Function

Private Sub BuscaContrato(ByVal psNroContrato As String)
Dim lbOk As Boolean
Dim lstTmpCliente As ListItem
Dim lrFunctVal As ADODB.Recordset
Dim lrValida As ADODB.Recordset
Dim lrCredPigPersonas As ADODB.Recordset
Dim lrCredPigDeuda As ADODB.Recordset
Dim loValContrato As nPigValida '
'Dim loValContrato As nColPValida

Dim loMuestraDatos As DPigContrato
Dim loFunct As DPigFunciones
Dim loCalc As NPigCalculos
Dim lnVarTipoTasacion As Integer
Dim lnDeuda As Currency, lnMinimoPagar As Currency
Dim lnDiasAtraso  As Integer
Dim rsJoyas As Recordset
Dim fnVarPorcConserv As Integer
Dim fnVarMaterial As Integer
Dim fnVarPesoNeto As Double
Dim fnVarTasacion As Currency
Dim fnVarTasacionAdic As Currency
Dim fnVarPrecioMaterial As Currency
Dim fnVarTipoTasacion As Integer
Dim fnVarTipoCliente As Integer
Dim fnVarPorcPrestamo As Double
Dim fnVarConservacion As Integer
Dim fnVarNewPrestamo As Currency
Dim fnVarComTasacion As Currency
Dim fnVarLineaDisponible As Currency

Dim lsmensaje As String

    'Valida Contrato
    Set lrValida = New ADODB.Recordset
    Set loValContrato = New nPigValida
    'Set loValContrato = New nColPValida
        Set lrValida = loValContrato.nValidaAmortizacionCredPignoraticio(psNroContrato, gdFecSis, 0, lsmensaje)
        If Trim(lsmensaje) <> "" Then
             MsgBox lsmensaje, vbInformation, "Aviso"
             Exit Sub
        End If
    Set loValContrato = Nothing
    
    If (lrValida Is Nothing) Then
        Limpiar
        Set lrValida = Nothing
        If AXCodCta.EnabledAge Then AXCodCta.SetFocusAge
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
    fnVarTasaInteres = lrValida!nTasaInteres
    lnNroTransac = lrValida!nTransacc
    
    'Muestra Datos
    txtDiasAtraso.Text = lrValida!nDiasAtraso
    txtNroUsoLinea = lrValida!nUsoLineaNro
    txtPlazoActual.Text = lrValida!nPlazo
    txtNroMvtos = Format(lrValida!nNroAmort, "0")
    TxtVctoActual = Format(lrValida!dvenc, "dd/mm/yyyy")
    
     ' Mostrar Clientes
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
    Set lrCredPigDeuda = New ADODB.Recordset
         Set lrCredPigDeuda = loMuestraDatos.dObtieneDatosPignoraticioDeuda(psNroContrato)
         Set loMuestraDatos = Nothing
     If lrCredPigDeuda.BOF And lrCredPigDeuda.EOF Then
        MsgBox " Error al mostrar Deuda del cliente ", vbCritical, " Aviso "
    Else
        fnVarSaldoCap = Format(lrCredPigDeuda!Capital, "#0.00")
        fnVarPorcCapMin = lrCredPigDeuda!PorcCapitalMinimo
        fnVarCapitalPagado = Format((fnVarPrestamo * fnVarPorcCapMin / 100), "#0.00")
        If fnVarCapitalPagado > fnVarSaldoCap Then
           fnVarCapitalPagado = fnVarSaldoCap
        End If
        fnVarCapMinimo = Format(lrCredPigDeuda!CapitalMinimo, "#0.00")
        
        Set loCalc = New NPigCalculos
        '************* RECALCULA INTERES COMPENSATORIO POR LOS DIAS TRANSCURRIDOS ***************
        fnVarIntCompensatorio = Format(loCalc.nCalculaIntCompensatorio(fnVarSaldoCap, fnVarTasaInteres, fnVarDiasIntereses), "#0.00")
        '*************
        
        Set lrFunctVal = New ADODB.Recordset
        Set loFunct = New DPigFunciones
              Set lrFunctVal = loFunct.GetConceptoValor(gColPigConceptoCodComiServ)
        fnVarComServicio = loCalc.nCalculaConcepto(lrFunctVal!nTpoValor, lrFunctVal!nValor, IIf(IsNull(lrFunctVal!nMontoMin), 0, lrFunctVal!nMontoMin), _
                                                  IIf(IsNull(lrFunctVal!nMontoMax), 0, lrFunctVal!nMontoMax), fnVarSaldoCap)  'Comision para el nuevo Calendario
        fnVarDiasCambCart = loFunct.GetParamValor(gPigParamDiasCambioCartera)
        fnVarIntMoratorio = Format(lrCredPigDeuda!IntMoratorio, "#0.00")
        fnVarComPenalidad = 0
        fnVarComVencida = Format(lrCredPigDeuda!ComVencida, "#0.00")
        fnVarDerRemate = Format(lrCredPigDeuda!DerRemate, "#0.00")
        fnVarDeuda = fnVarSaldoCap + fnVarIntCompensatorio + fnVarIntMoratorio + fnVarComServicio + _
                            fnVarComPenalidad + fnVarComVencida + fnVarDerRemate
        fnVarMontoAPagar = fnVarCapitalPagado + (fnVarIntCompensatorio + fnVarIntMoratorio + fnVarComServicio + _
                                                                      fnVarComPenalidad + fnVarComVencida + fnVarDerRemate)
        fnVarMontoMinimo = fnVarCapMinimo + (fnVarIntCompensatorio + fnVarIntMoratorio + fnVarComServicio + _
                                                                      fnVarComPenalidad + fnVarComVencida + fnVarDerRemate)
                                                                      
        txtTotalDeuda.Text = Format(fnVarDeuda, "#0.00")
        txtMontoMinimoPagar.Text = Format(fnVarMontoMinimo, "#0.00")
        
        txtMontoPagar.Text = Format(fnVarMontoAPagar, "#0.00")
        fnVarNewSaldoCap = fnVarSaldoCap - fnVarCapitalPagado
        txtSaldoCapitalNuevo.Text = Format(fnVarNewSaldoCap, "#0.00")
        txtPlazoNuevo.Text = lrValida!nPlazo
        DtpNvoVcto.value = Format(gdFecSis + Format(txtPlazoNuevo.Text, "#0"), "dd/mm/yyyy")
        txtPlazoNuevo.Enabled = True
        txtMontoPagar.Enabled = True
        
        '***** PARA EL CALCULO DEL MONTO DISPONIBLE PARA REUSO DE LINEA
        Set loMuestraDatos = New DPigContrato
        
        Set rsJoyas = loMuestraDatos.dObtieneDatosPignoraticioJoyas(psNroContrato) 'Obtiene Caracteristicas de las joyas
       
        If rsJoyas.BOF And rsJoyas.EOF Then
            MsgBox " Error al Obtener datos de las Joyas ", vbCritical, " Aviso "
        Else
            Set loFunct = New DPigFunciones
            fnVarTipoTasacion = rsJoyas!nTipoTasacion
            If fnVarTipoCliente = 1 Then              ' Cliente A1
               fnVarPorcPrestamo = loFunct.GetParamValor(8001)
            ElseIf fnVarTipoCliente = 2 Then        ' Cliente A
               fnVarPorcPrestamo = loFunct.GetParamValor(8002)
            ElseIf fnVarTipoCliente = 3 Then        ' Cliente B
               fnVarPorcPrestamo = loFunct.GetParamValor(8003)
            ElseIf fnVarTipoCliente = 4 Then        ' Cliente B1
               fnVarPorcPrestamo = loFunct.GetParamValor(8004)
            Else
                 MsgBox "Error: El Tipo de Cliente No ha sido Considerado", vbInformation, " Aviso "
        End If
        
        Do While Not rsJoyas.EOF
             fnVarMaterial = rsJoyas!nMaterial
             fnVarConservacion = rsJoyas!nConservacion
             fnVarPesoNeto = rsJoyas!npesoneto
             fnVarTasacionAdic = fnVarTasacionAdic + rsJoyas!nTasacionAdicional
             If fnVarConservacion = 1 Then
                fnVarPorcConserv = loFunct.GetParamValor(8010)          ' Conservacion de Joya Buena
             ElseIf fnVarConservacion = 2 Then
                fnVarPorcConserv = loFunct.GetParamValor(8011)          ' Conservacion de Joya Regular
             ElseIf fnVarConservacion = 3 Then
                fnVarPorcConserv = loFunct.GetParamValor(8012)          ' Conservacion de Joya Malo
             Else
                  MsgBox "Estado de Conservación de la Joya No ha sido Considerado", vbInformation, " Aviso "
             End If
             fnVarPrecioMaterial = loFunct.GetPrecioMaterial(1, fnVarMaterial, 1)
             fnVarTasacion = fnVarTasacion + Round(loCalc.CalcValorTasacion(fnVarPorcConserv, fnVarPesoNeto, fnVarPrecioMaterial), 2)
             
             rsJoyas.MoveNext
        Loop
        fnVarNewPrestamo = Round(loCalc.CalcValorPrestamo(fnVarPorcPrestamo, fnVarTasacion + fnVarTasacionAdic), 2)
                     
        Set lrFunctVal = loFunct.GetConceptoValor(gColPigConceptoCodTasacion)
        fnVarComTasacion = loCalc.nCalculaConcepto(lrFunctVal!nTpoValor, lrFunctVal!nValor, IIf(IsNull(lrFunctVal!nMontoMin), 0, lrFunctVal!nMontoMin), _
                                                                    IIf(IsNull(lrFunctVal!nMontoMax), 0, lrFunctVal!nMontoMax), fnVarSaldoCap)
        Set loFunct = Nothing
        fnVarLineaDisponible = Format(fnVarNewPrestamo - fnVarDeuda - fnVarComTasacion, "#0.00")
    
        If fnVarLineaDisponible > 0 Then
            txtLineaDisponible.Text = Format(fnVarLineaDisponible, "#0.00")
        Else
            fnVarLineaDisponible = 0
            txtLineaDisponible.Text = Format(fnVarLineaDisponible, "#0.00")
        End If
        
        TxtITF.Text = Format(fgITFCalculaImpuestoNOIncluido(CDbl(txtMontoPagar.Text)) - CDbl(txtMontoPagar.Text), "#0.00")
        TxtApagarConITF.Text = Format(CDbl(txtMontoPagar.Text) + CDbl(TxtITF.Text), "#0.00")
        
    End If
        
    End If
    
    AXCodCta.Enabled = False
    txtMontoPagar.SetFocus
   
    Set lrValida = Nothing
    Set lrCredPigDeuda = Nothing
    Set lrCredPigPersonas = Nothing
    Set loFunct = Nothing
    Set loCalc = Nothing
    
    txtPlazoNuevo.SetFocus
    AXCodCta.Enabled = False

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

Private Sub CmdCancelar_Click()
    Limpiar
    cmdGrabar.Enabled = False
    txtPlazoNuevo.Enabled = False
    txtMontoPagar.Enabled = False
    AXCodCta.Enabled = True
    If AXCodCta.EnabledAge Then AXCodCta.SetFocusAge
End Sub

Private Sub cmdGrabar_Click()

Dim loContFunct As NContFunciones
Dim loGrabarAmort As NPigContrato
Dim oImprime As NPigImpre
Dim loPrevio As Previo.clsPrevio
Dim loCalc As NPigCalculos

Dim lsMovNro As String
Dim lnMovNro As Long
Dim lsFechaHoraGrab As String
Dim lsFechaVenc As String
Dim lnMontoTransaccion As Currency
Dim lnNewPagoMin As Currency
Dim lsCadImprimir As String
Dim lsNombreCliente As String
Dim lnDiasIC As Integer
Dim sOpeCod As String


lsFechaVenc = Format$(gdFecSis + fnVarNewPlazo, "mm/dd/yyyy")
lnMontoTransaccion = CCur(Me.txtMontoPagar.Text)
lsNombreCliente = lstClientes.ListItems(1).ListSubItems.Item(1)
If txtDiasAtraso > 0 Then
    sOpeCod = gPigOpeAmortMorEFE
Else
    sOpeCod = gPigOpeAmortNorEFE
End If

If MsgBox(" Grabar Amortización de Contrato ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
 'Realiza la Validación para el Lavado de Dinero
    Dim clsLav As nCapDefinicion
    Dim nPorcRetCTS As Double, nMontoLavDinero As Double, nTC As Double
    Dim sPersLavDinero As String
    Dim nMonto As Double
    
    sPersLavDinero = ""
    Set clsLav = New nCapDefinicion
    
    If clsLav.EsOperacionEfectivo(Trim(sOpeCod)) Then
        If Not EsExoneradaLavadoDinero() Then
        
            nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
            Set clsLav = Nothing
            
                Dim clsTC As nTipoCambio
                Set clsTC = New nTipoCambio
                nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
                Set clsTC = Nothing
           
            nMonto = CDbl(Me.txtMontoPagar.Text)
            
            If nMonto >= Round(nMontoLavDinero * nTC, 2) Then
                sPersLavDinero = IniciaLavDinero()
                If sPersLavDinero = "" Then Exit Sub
            End If
        End If
    End If
    
    '********************************************************************
         Set loCalc = New NPigCalculos
        fnVarNewIntCompensatorio = Format(loCalc.nCalculaIntCompensatorio(fnVarNewSaldoCap, fnVarTasaInteres, fnVarNewPlazo), "#0.00")   'Obtiene Interes y Comisiones para el nuevo Capital
        cmdGrabar.Enabled = False
        
        'Genera el Mov Nro
        Set loContFunct = New NContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        lnNewPagoMin = Format((fnVarPrestamo * fnVarPorcCapMin / 100), "#0.00") + fnVarNewIntCompensatorio + fnVarComServicio
        Set loGrabarAmort = New NPigContrato
        
        'Grabar Amortizacion Pignoraticio
        lnMovNro = loGrabarAmort.nAmortizacionCredPignoraticio(AXCodCta.NroCuenta, Format(fnVarNewSaldoCap, "#0.00"), lsFechaHoraGrab, _
                                   lsMovNro, lsFechaVenc, Format(fnVarNewPlazo, "#0.00"), lnMontoTransaccion, Format(fnVarCapitalPagado, "#0.00"), _
                                   Format(fnVarIntCompensatorio, "#0.00"), Format(fnVarIntMoratorio, "#0.00"), Format(fnVarComServicio, "#0.00"), _
                                   Format(fnVarComPenalidad, "#0.00"), Format(fnVarComVencida, "#0.00"), Format(fnVarDerRemate, "#0.00"), _
                                   fnVarDiasAtraso, CLng(txtNroMvtos.Text) + 1, fnVarDiasCambCart, Format(fnVarValorTasacion, "#0.00"), fnVarOpeCod, _
                                   Format(fnVarNewIntCompensatorio, "#0.00"), fsVarOpeDesc, fsVarPersCodCMAC, fnVarNroCalend, fnVarPlazo, _
                                   fnVarEstado, Format(fnVarSaldoCap, "#0.00"), False, sPersLavDinero, sPersCod)
        Set loGrabarAmort = Nothing
        ' **************************************************
        
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
      
        
        Set oImprime = New NPigImpre
        Call oImprime.ImpreReciboAmortizacion(gsInstCmac, gsNomAge, lsFechaHoraGrab, AXCodCta.NroCuenta, lsNombreCliente, _
                                                            lsFechaVenc, fnVarDiasAtraso, fnVarCapitalPagado, _
                                                            fnVarIntCompensatorio, fnVarIntMoratorio, fnVarComServicio, fnVarDerRemate, _
                                                            fnVarComVencida, lnMontoTransaccion, fnVarNewSaldoCap, gsCodUser, _
                                                            lnNroTransac + 1, lnMovNro, lnNewPagoMin, fnVarDiasIntereses, sLpt, " ", gsCodCMAC, gcEmpresaRUC, CDbl(TxtITF.Text))
                                                            
        Do While MsgBox("Reimprimir Recibo de Amortización ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes
            Call oImprime.ImpreReciboAmortizacion(gsInstCmac, gsNomAge, lsFechaHoraGrab, AXCodCta.NroCuenta, lsNombreCliente, _
                                                            lsFechaVenc, fnVarDiasAtraso, fnVarCapitalPagado, _
                                                            fnVarIntCompensatorio, fnVarIntMoratorio, fnVarComServicio, fnVarDerRemate, _
                                                            fnVarComVencida, lnMontoTransaccion, fnVarNewSaldoCap, gsCodUser, _
                                                            lnNroTransac + 1, lnMovNro, lnNewPagoMin, fnVarDiasIntereses, sLpt, " ", gsCodCMAC, gcEmpresaRUC, CDbl(TxtITF.Text))
        Loop
    
        Set oImprime = Nothing
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
    Me.Icon = LoadPicture(App.path & "\bmps\cm.ico")
    AXCodCta.EnabledProd = False
    AXCodCta.Age = " "
    AXCodCta.Texto = "Crédito"
    
End Sub

Private Sub txtPlazoNuevo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If CInt(txtPlazoNuevo.Text) >= fnVarPlzMin And CInt(txtPlazoNuevo.Text) <= fnVarPlzMax Then
           fnVarNewPlazo = CInt(txtPlazoNuevo.Text)
           DtpNvoVcto.value = Format(DateAdd("d", fnVarNewPlazo, gdFecSis), "dd/mm/yyyy")
           txtMontoPagar.SetFocus
       Else
           MsgBox "Error: Plazo Fuera del Rango Permitido", vbInformation, " Aviso "
       End If
    End If
End Sub

Private Sub cmdRebajar_Click()
    fnVarCapitalPagado = fnVarCapMinimo
    fnVarNewSaldoCap = fnVarSaldoCap - fnVarCapitalPagado
    txtSaldoCapitalNuevo.Text = Format(fnVarNewSaldoCap, "#0.00")
    fnVarMontoMinimo = fnVarCapitalPagado + (fnVarIntCompensatorio + fnVarIntMoratorio + fnVarComServicio + _
                                                                  fnVarComPenalidad + fnVarComVencida + fnVarDerRemate)
    txtMontoMinimoPagar.Text = Format(fnVarMontoMinimo, "#0.00")
    txtMontoPagar.Text = Format(fnVarMontoMinimo, "#0.00")
    If txtMontoPagar.Enabled Then txtMontoPagar.SetFocus
End Sub

Private Sub DtpNvoVcto_Change()
    If DateDiff("d", gdFecSis, DtpNvoVcto.value) >= fnVarPlzMin And DateDiff("d", gdFecSis, DtpNvoVcto.value) <= fnVarPlzMax Then
        txtPlazoNuevo.Text = DateDiff("d", gdFecSis, DtpNvoVcto.value)
        fnVarNewPlazo = CInt(txtPlazoNuevo.Text)
        txtMontoPagar.SetFocus
    Else
        DtpNvoVcto.value = Format(DateAdd("d", fnVarNewPlazo, gdFecSis), "dd/mm/yyyy")
        MsgBox "Error: Plazo Fuera del Rango Permitido", vbInformation, " Aviso "
    End If
End Sub
'Valida el campo txtmontopagar
Private Sub txtMontoPagar_GotFocus()
    fEnfoque txtMontoPagar
End Sub

Private Sub txtMontoPagar_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMontoPagar, KeyAscii)
If KeyAscii = 13 Then
    If CCur(txtMontoPagar) >= CCur(txtTotalDeuda) Then      'Se debe ingresar por Cancelacion de Contrato
       MsgBox " Usar Transaccion de Cancelación de Pignoraticio", , " Aviso "
       txtMontoPagar.SetFocus
       cmdGrabar.Enabled = False
       Exit Sub
    End If
    If CCur(txtMontoPagar) < CCur(txtMontoMinimoPagar) Then      'Transaccion no Permitida
       MsgBox " Monto a Pagar debe ser Mayor que el Mìnimo", , " Aviso "
       txtMontoPagar.SetFocus
       cmdGrabar.Enabled = False
       Exit Sub
    End If
    
    'Distribuye los importes a las diferentes rubros
       
    If CCur(txtMontoPagar.Text) >= CCur(txtMontoMinimoPagar) Then
       fnVarCapitalPagado = CCur(txtMontoPagar.Text) - (fnVarIntCompensatorio + fnVarIntMoratorio + fnVarComServicio + _
                                                                            fnVarComPenalidad + fnVarComVencida + fnVarDerRemate)
    End If
    fnVarNewSaldoCap = Format(fnVarSaldoCap - fnVarCapitalPagado, "#0.00")
    txtSaldoCapitalNuevo.Text = fnVarNewSaldoCap
    cmdGrabar.Enabled = True
    cmdGrabar.SetFocus
End If
End Sub
