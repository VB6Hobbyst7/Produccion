VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPigCancelacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelación de Contratos"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000002&
   Icon            =   "frmPigCancelacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5580
      TabIndex        =   22
      Top             =   4185
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4380
      TabIndex        =   21
      Top             =   4185
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6735
      TabIndex        =   20
      Top             =   4185
      Width           =   975
   End
   Begin VB.Frame fraContenedor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4050
      Index           =   0
      Left            =   15
      TabIndex        =   0
      Top             =   -30
      Width           =   7815
      Begin VB.CommandButton cmdBuscar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7110
         Picture         =   "frmPigCancelacion.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Buscar ..."
         Top             =   270
         Width           =   525
      End
      Begin VB.Frame fraContenedor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Index           =   5
         Left            =   135
         TabIndex        =   16
         Top             =   2700
         Width           =   7515
         Begin VB.TextBox txtLineaDisponible 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   315
            Left            =   1605
            TabIndex        =   26
            Top             =   450
            Width           =   1320
         End
         Begin VB.TextBox txtMontoPagar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5550
            TabIndex        =   25
            Top             =   300
            Visible         =   0   'False
            Width           =   1680
         End
         Begin VB.Frame Frame1 
            Caption         =   "Frame1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   45
            Left            =   5325
            TabIndex        =   24
            Top             =   120
            Width           =   60
         End
         Begin VB.TextBox TxtTotalDeuda 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
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
            Height          =   315
            Left            =   6030
            TabIndex        =   17
            Top             =   720
            Width           =   1215
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
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   240
            TabIndex        =   27
            Top             =   525
            Width           =   1230
         End
         Begin VB.Label Label4 
            Caption         =   "Total Deuda  S/."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4110
            TabIndex        =   18
            Top             =   765
            Width           =   1845
         End
      End
      Begin VB.Frame fraContenedor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   1
         Left            =   135
         TabIndex        =   13
         Top             =   705
         Width           =   7515
         Begin MSComctlLib.ListView lstClientes 
            Height          =   735
            Left            =   75
            TabIndex        =   14
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   2
         Left            =   120
         TabIndex        =   1
         Top             =   1725
         Width           =   7515
         Begin VB.TextBox txtPlazoActual 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6390
            MaxLength       =   2
            TabIndex        =   6
            Top             =   255
            Width           =   765
         End
         Begin VB.TextBox txtNroMvtos 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1620
            TabIndex        =   5
            Top             =   615
            Width           =   690
         End
         Begin VB.TextBox txtNroUsoLinea 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   4065
            TabIndex        =   4
            Top             =   615
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
            Height          =   300
            Left            =   4065
            TabIndex        =   3
            Top             =   255
            Width           =   690
         End
         Begin VB.TextBox TxtVctoActual 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1605
            TabIndex        =   2
            Top             =   210
            Width           =   1005
         End
         Begin VB.Label lblNumDuplicado 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   6390
            TabIndex        =   29
            Top             =   615
            Width           =   780
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Nro. Duplicado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   5085
            TabIndex        =   28
            Top             =   660
            Width           =   1170
         End
         Begin VB.Label Label5 
            Caption         =   "Fec.Vencimiento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   135
            TabIndex        =   12
            Top             =   255
            Width           =   1215
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Plazo Actual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   11
            Left            =   5070
            TabIndex        =   11
            Top             =   285
            Width           =   960
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Dias Atraso"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   7
            Left            =   2835
            TabIndex        =   10
            Top             =   270
            Width           =   960
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Nro Amortizaciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   8
            Left            =   120
            TabIndex        =   9
            Top             =   645
            Width           =   1395
         End
         Begin VB.Label lblMoneda 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2070
            TabIndex        =   8
            Top             =   210
            Width           =   255
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Usos de Linea"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   2835
            TabIndex        =   7
            Top             =   675
            Width           =   1110
         End
      End
      Begin SICMACT.ActXCodCta AxCodCta 
         Height          =   405
         Left            =   135
         TabIndex        =   15
         Top             =   270
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   714
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
   End
   Begin VB.Label lblMensaje 
      Caption         =   "<F12> Número Contrato Antiguo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   165
      TabIndex        =   23
      Top             =   4260
      Width           =   2385
   End
End
Attribute VB_Name = "frmPigCancelacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************************'
'* CANCELACION DE CONTRATO DE PIGNORATICIO
'* Archivo   :  frmPigCancelacion.frm
'* Resumen:  Nos permite cancelar un credito y dejarlo pendiente de rescate
'***************************************************************************************'
Option Explicit

Dim fnVarOpeCod As Long
Dim fsVarOpeDesc As String
Dim fsVarPersCodCMAC As String
Dim fsVarNombreCMAC As String

Dim fnVarTasaInteres As Currency
Dim fnVarValorTasacion As Currency
Dim fnVarPlazo As Integer
Dim fdVarFecVencimiento As Date
Dim fdVarFecUltPago As Date
Dim fnVarEstado As ColocEstado
Dim fnVarNroRenovacion As Integer

Dim fnVarPrestamo As Currency
Dim fnVarSaldoCap As Currency
Dim fnVarNewSaldoCap As Currency
Dim fnVarCapitalPagado As Currency   ' Capital a Pagar
Dim fnVarIntCompensatorio As Currency
Dim fnVarIntMoratorio As Currency
Dim fnVarComServicio As Currency
Dim fnVarComPenalidad As Currency
Dim fnVarComVencida As Currency
Dim fnVarDerRemate As Currency

Dim fnVarDiasAtraso As Double
Dim fnVarDiasCambCart  As Currency
Dim fnVarDiasIntereses As Integer     '''''' Calculado desde el ultimo movimiento ya sea desembolso o pago
Dim fnVarDiasVigencia As Integer      '''''Para Determinar si se cobrara penalidad
Dim fnVarDeuda As Currency

'*********
Dim fnVarMontoMinimo As Currency
Dim fnVarMontoAPagar As Currency
Dim fnVarCapMinimo As Currency        'Capital Minimo, almacena del Codigo 8017 del ColocParametro

Dim fnVarNroCalend As Integer
Dim fnVarTipoCliente As Integer
Dim rsJoyas As Recordset
Dim fnVarTipoTasacion As Integer
Dim fnVarPorcPrestamo As Double
Dim fnVarMaterial As Integer
Dim fnVarConservacion As Integer
Dim fnVarPesoNeto As Double
Dim fnVarTasacionAdic As Currency
Dim fnVarPorcConserv As Double
Dim fnVarPrecioMaterial As Currency
Dim fnVarTasacion As Currency
Dim fnVarNewPrestamo As Currency
Dim fnVarComTasacion As Currency
Dim fnVarLineaDisponible As Currency
Dim lnNroTransac As Integer
'LAVADO DE DINERO
Dim sPersCod As String, sNombre As String, sDireccion As String, sDocId As String
Dim sTipoCuenta As String, sOperacion As String, sCuenta As String
'*********

Public Sub Inicio(ByVal pnOpeCod As Long, ByVal psOpeDesc As String, _
        ByVal psPersCodCMAC As String, ByVal psNomCMAC As String)

    fnVarOpeCod = pnOpeCod
    fsVarOpeDesc = psOpeDesc
    fsVarPersCodCMAC = psPersCodCMAC
    fsVarNombreCMAC = psNomCMAC
   Limpiar
    Me.Show 1

End Sub

Private Sub Limpiar()
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio   ' Procedimiento que arma los primeros campos de la cuenta
    AXCodCta.Age = ""
    lstClientes.ListItems.Clear
    TxtVctoActual.Text = Format("  /  /  ", "")
    txtDiasAtraso.Text = Format(0, "0")
    txtNroMvtos = Format(0, "0")
    txtNroUsoLinea = Format(0, "0")
    txtMontoPagar.Text = Format(0, "#0.00")
    txtTotalDeuda.Text = Format(0, "#0.00")
    txtLineaDisponible = ""
    fnVarComPenalidad = 0
    
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call BuscaContrato(AXCodCta.NroCuenta)
    End If
End Sub

Private Sub BuscaContrato(ByVal psNroContrato As String)
Dim lbOk As Boolean
Dim lstTmpCliente As ListItem
Dim lrValida As ADODB.Recordset
Dim lrCredPigPersonas As ADODB.Recordset
Dim lrCredPigDeuda As ADODB.Recordset
Dim lrFunctVal As ADODB.Recordset
Dim loValContrato As nPigValida
Dim loCalc As NPigCalculos
Dim loMuestraDatos As DPigContrato
Dim loFunct As DPigFunciones

Dim lsmensaje As String

Dim lnDeuda As Currency
Dim lnDiasAtraso  As Integer

    fnVarLineaDisponible = 0: fnVarTasacion = 0

    
    'Valida Contrato
    Set lrValida = New ADODB.Recordset
    Set loValContrato = New nPigValida
        Set lrValida = loValContrato.nValidaCancelacionCredPignoraticio(psNroContrato, gdFecSis, 0, lsmensaje)
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
    fnVarSaldoCap = Format(lrValida!nSaldo, "#0.00")
    fnVarEstado = lrValida!nPrdEstado
    fdVarFecVencimiento = Format(lrValida!dvenc, "dd/mm/yyyy")
    fdVarFecUltPago = Format(lrValida!dPrdEstado, "dd/mm/yyyy")
    fnVarDiasVigencia = DateDiff("d", Format(lrValida!dVigencia, "dd/mm/yyyy"), Format(gdFecSis, "dd/mm/yyyy"))
    fnVarDiasIntereses = DateDiff("d", fdVarFecUltPago, Format(gdFecSis, "dd/mm/yyyy"))
    fnVarDiasAtraso = lrValida!nDiasAtraso
    fnVarValorTasacion = lrValida!totTasacion
    fnVarNroCalend = lrValida!nNumCalend
    fnVarTasaInteres = lrValida!nTasaInteres
    lnDiasAtraso = lrValida!nDiasAtraso
    lnNroTransac = lrValida!nTransacc
    lblNumDuplicado = IIf(IsNull(lrValida!nNroDuplic), 0, lrValida!nNroDuplic)
    
    'Muestra Datos
    txtDiasAtraso.Text = lrValida!nDiasAtraso
    txtNroUsoLinea = lrValida!nUsoLineaNro
    txtPlazoActual.Text = lrValida!nPlazo
    txtNroMvtos = IIf(IsNull(lrValida!nNroAmort), 0, lrValida!nNroAmort)
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
                  fnVarTipoCliente = lrCredPigPersonas!cCalifiCliente
            lrCredPigPersonas.MoveNext
        Loop
    End If
  
    ' MostrarDeuda
    Set lrCredPigDeuda = New ADODB.Recordset
    Set loMuestraDatos = New DPigContrato
         Set lrCredPigDeuda = loMuestraDatos.dObtieneDatosPignoraticioDeuda(psNroContrato)
       
    If lrCredPigDeuda.BOF And lrCredPigDeuda.EOF Then
        MsgBox " Error al mostrar Deuda del cliente ", vbCritical, " Aviso "
    Else
        fnVarSaldoCap = lrCredPigDeuda!Capital
        fnVarCapitalPagado = lrCredPigDeuda!Capital
        
        Set loCalc = New NPigCalculos
        '************* RECALCULA INTERES COMPENSATORIO POR LOS DIAS TRANSCURRIDOS ***************
        fnVarIntCompensatorio = Format(loCalc.nCalculaIntCompensatorio(fnVarSaldoCap, fnVarTasaInteres, fnVarDiasIntereses), "#0.00")
        
        Set lrFunctVal = New ADODB.Recordset
        Set loFunct = New DPigFunciones
              Set lrFunctVal = loFunct.GetConceptoValor(Val(gColPigConceptoCodPenalidad))
        
        If fnVarDiasVigencia < loFunct.GetParamValor(8019) Then
                 fnVarComPenalidad = loCalc.nCalculaConcepto(lrFunctVal!nTpoValor, lrFunctVal!nValor, IIf(IsNull(lrFunctVal!nMontoMin), 0, lrFunctVal!nMontoMin), _
                                                            IIf(IsNull(lrFunctVal!nMontoMax), 0, lrFunctVal!nMontoMax), fnVarSaldoCap)
        End If
        
        Set lrFunctVal = loFunct.GetConceptoValor(Val(gColPigConceptoCodComiServ))
        fnVarComServicio = loCalc.nCalculaConcepto(lrFunctVal!nTpoValor, lrFunctVal!nValor, IIf(IsNull(lrFunctVal!nMontoMin), 0, lrFunctVal!nMontoMin), _
                                                   IIf(IsNull(lrFunctVal!nMontoMax), 0, lrFunctVal!nMontoMax), fnVarSaldoCap) 'Comision para el nuevo Calendario
        
        fnVarIntMoratorio = lrCredPigDeuda!IntMoratorio
        fnVarComVencida = lrCredPigDeuda!ComVencida
        fnVarDerRemate = lrCredPigDeuda!DerRemate
        fnVarDeuda = fnVarSaldoCap + fnVarIntCompensatorio + fnVarIntMoratorio + fnVarComServicio + _
                            fnVarComPenalidad + fnVarComVencida + fnVarDerRemate
        txtTotalDeuda.Text = Format(fnVarDeuda, "#0.00")
        txtMontoPagar.Text = Format(fnVarDeuda, "#0.00")
        fnVarNewSaldoCap = fnVarSaldoCap - fnVarCapitalPagado
    End If
    
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
                 
    Set lrFunctVal = loFunct.GetConceptoValor(Val(gColPigConceptoCodTasacion))
    fnVarComTasacion = loCalc.nCalculaConcepto(lrFunctVal!nTpoValor, lrFunctVal!nValor, IIf(IsNull(lrFunctVal!nMontoMin), 0, lrFunctVal!nMontoMin), _
                                                                IIf(IsNull(lrFunctVal!nMontoMax), 0, lrFunctVal!nMontoMax), fnVarSaldoCap)
    Set loFunct = Nothing
    fnVarLineaDisponible = Format(fnVarNewPrestamo - fnVarDeuda - fnVarComTasacion, "###,##0.00")

    If fnVarLineaDisponible > 0 Then
        txtLineaDisponible.Text = Format(fnVarLineaDisponible, "###,##0.00")
    Else
        fnVarLineaDisponible = 0
        txtLineaDisponible.Text = Format(fnVarLineaDisponible, "###,##0.00")
    End If

    AXCodCta.Enabled = False
   
    Set lrValida = Nothing
    Set lrCredPigDeuda = Nothing
    Set lrCredPigPersonas = Nothing
    Set loFunct = Nothing
    Set loCalc = Nothing
        
    AXCodCta.Enabled = False
    cmdGrabar.Enabled = True
    cmdGrabar.SetFocus

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

ControlError:
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdCancelar_Click()
    Limpiar
    cmdGrabar.Enabled = False
    AXCodCta.Enabled = True
    If AXCodCta.EnabledAge Then AXCodCta.SetFocusAge
End Sub

Private Sub cmdGrabar_Click()

'On Error GoTo ControlError
Dim loContFunct As NContFunciones
Dim loGrabarCan As NPigContrato
Dim oImprime As NPigImpre
Dim loPrevio As Previo.clsPrevio

Dim lsMovNro As String
Dim lnMovNro As Long
Dim lsFechaHoraGrab As String
Dim lsCuenta As String

Dim lnSaldoCap As Currency, lnInteresComp As Currency
Dim lnCostoTasacion As Currency, lnCostoCustodia As Currency
Dim lnMontoTransaccion As Currency

Dim lsCadImprimir As String
Dim lsNombreCliente As String
Dim sOpeCod As String

lnMontoTransaccion = CCur(Me.txtTotalDeuda)
lsNombreCliente = lstClientes.ListItems(1).ListSubItems.Item(1)
If txtDiasAtraso > 0 Then
    sOpeCod = gPigOpeCancelMorEFE
Else
    sOpeCod = gPigOpeCancelNorEFE
End If
If MsgBox(" Grabar Cancelacion de Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
    cmdGrabar.Enabled = False
        
    ' *******************************************************************************
    'Realiza la Validación para el Lavado de Dinero
    Dim clsLav As nCapDefinicion
    Dim nMontoLavDinero As Double, nTC As Double
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
      '*********************************************************************************
        
        'Genera el Mov Nro
        Set loContFunct = New NContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        Set loGrabarCan = New NPigContrato
        'Grabar Cancelacion Pignoraticio

        lnMovNro = loGrabarCan.nCancelacionCredPignoraticio(AXCodCta.NroCuenta, Format(fnVarNewSaldoCap, "#0.00"), lsFechaHoraGrab, lsMovNro, _
                                                    lnMontoTransaccion, Format(fnVarCapitalPagado, "#0.00"), Format(fnVarIntCompensatorio, "#0.00"), _
                                                    Format(fnVarIntMoratorio, "#0.00"), Format(fnVarComServicio, "#0.00"), Format(fnVarComPenalidad, "#0.00"), _
                                                    Format(fnVarComVencida, "#0.00"), Format(fnVarDerRemate, "#0.00"), Format(fnVarDiasAtraso, "#0.00"), _
                                                    Format(fnVarValorTasacion, "#0.00"), fnVarOpeCod, fsVarOpeDesc, fsVarPersCodCMAC, fnVarNroCalend, _
                                                    fnVarPlazo, fnVarEstado, Format(fnVarSaldoCap, "#0.00"), fnVarDiasCambCart, sPersLavDinero, sPersCod)
        Set loGrabarCan = Nothing

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

        'Impresión
        Set oImprime = New NPigImpre
        Call oImprime.ImpreReciboCancelacion(gsInstCmac, gsNomAge, lsFechaHoraGrab, AXCodCta.NroCuenta, lsNombreCliente, _
                                             fnVarDiasAtraso, fnVarCapitalPagado, fnVarIntCompensatorio, fnVarIntMoratorio, _
                                             fnVarComServicio, fnVarDerRemate, fnVarComVencida, fnVarComPenalidad, _
                                             lnMontoTransaccion, gsCodUser, lnMovNro, lnNroTransac + 1, sLpt, "")
                                             
        Do While MsgBox("Reimprimir Recibo de Cancelación ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes
            Call oImprime.ImpreReciboCancelacion(gsInstCmac, gsNomAge, lsFechaHoraGrab, AXCodCta.NroCuenta, lsNombreCliente, _
                                             fnVarDiasAtraso, fnVarCapitalPagado, fnVarIntCompensatorio, fnVarIntMoratorio, _
                                             fnVarComServicio, fnVarDerRemate, fnVarComVencida, fnVarComPenalidad, _
                                             lnMontoTransaccion, gsCodUser, lnMovNro, lnNroTransac + 1, sLpt, "")
        Loop
        
        Set oImprime = Nothing
        Limpiar
        AXCodCta.Enabled = True
        AXCodCta.SetFocus
        
Else
    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
End If
Exit Sub

ControlError:
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

    AXCodCta.Texto = "Créditos"
    AXCodCta.Age = ""

End Sub

Private Function IniciaLavDinero() As String
Dim i As Long
Dim nRelacion As CaptacRelacPersona
Dim nMonto As Double
Dim oPersona As NCapMantenimiento
Dim oDatos As DPigContrato
Dim rsPersPigno As Recordset

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

    nMonto = CDbl(txtTotalDeuda.Text)
    IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, False, True, nMonto, "", sOperacion, False, sTipoCuenta)
       
End Function

Private Function EsExoneradaLavadoDinero() As Boolean
Dim bExito As Boolean
Dim clsExo As NCapServicios
bExito = True

    sPersCod = lstClientes.ListItems(1)
    Set clsExo = New NCapServicios
    
    If Not clsExo.EsPersonaExoneradaLavadoDinero(sPersCod) Then bExito = False

    Set clsExo = Nothing
    EsExoneradaLavadoDinero = bExito
    
End Function
