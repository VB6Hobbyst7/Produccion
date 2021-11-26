VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPigUsoLineaPR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Uso de Linea - Contratos Pendientes de Rescate"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBuscar 
      Height          =   360
      Left            =   3840
      Picture         =   "FrmPigUsoLineaPR.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Buscar ..."
      Top             =   105
      Width           =   420
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   9255
      TabIndex        =   32
      Top             =   6495
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   10320
      TabIndex        =   31
      Top             =   6495
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   8175
      TabIndex        =   30
      Top             =   6495
      Width           =   975
   End
   Begin VB.Frame FraContenedor 
      Height          =   5925
      Index           =   0
      Left            =   15
      TabIndex        =   4
      Top             =   435
      Width           =   11460
      Begin VB.Frame FraContenedor 
         Caption         =   "Cliente(s)"
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
         Height          =   990
         Index           =   6
         Left            =   90
         TabIndex        =   28
         Top             =   150
         Width           =   11280
         Begin MSComctlLib.ListView lstCliente 
            Height          =   660
            Left            =   75
            TabIndex        =   29
            Top             =   240
            Width           =   11130
            _ExtentX        =   19632
            _ExtentY        =   1164
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
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
            NumItems        =   11
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Codigo del Cliente"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nombre / Razón Social del Cliente"
               Object.Width           =   5293
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Dirección"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Teléfono"
               Object.Width           =   1413
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Zona"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Doc.Civil"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Nro.Doc.Civil"
               Object.Width           =   2382
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Doc.Tributario"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Nro.Doc.Tributario"
               Object.Width           =   2382
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Evaluacion"
               Object.Width           =   2222
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "SBS"
               Object.Width           =   2222
            EndProperty
         End
      End
      Begin VB.Frame FraContenedor 
         Enabled         =   0   'False
         Height          =   540
         Index           =   2
         Left            =   90
         TabIndex        =   21
         Top             =   5325
         Width           =   11265
         Begin SICMACT.EditMoney lblNetoRecibir 
            Height          =   285
            Left            =   10110
            TabIndex        =   22
            Top             =   195
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   503
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            BackColor       =   12648447
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin SICMACT.EditMoney lblCostoTasacion 
            Height          =   285
            Left            =   7500
            TabIndex        =   23
            Top             =   195
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   503
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin SICMACT.EditMoney lblPagoMin 
            Height          =   285
            Left            =   1245
            TabIndex        =   24
            Top             =   165
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   503
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Neto Recibir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   20
            Left            =   8970
            TabIndex        =   27
            Top             =   225
            Width           =   1140
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Costo Tasación"
            Height          =   255
            Index           =   19
            Left            =   6270
            TabIndex        =   26
            Top             =   195
            Width           =   1200
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Pago Mínimo"
            Height          =   255
            Index           =   16
            Left            =   210
            TabIndex        =   25
            Top             =   195
            Width           =   975
         End
      End
      Begin VB.Frame FraContenedor 
         Caption         =   "Joya(s)"
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
         Height          =   3600
         Index           =   4
         Left            =   90
         TabIndex        =   12
         Top             =   1155
         Width           =   11310
         Begin VB.TextBox TxtTotalB 
            Appearance      =   0  'Flat
            Height          =   255
            Left            =   1335
            TabIndex        =   13
            Top             =   3255
            Width           =   705
         End
         Begin SICMACT.FlexEdit FEJoyas 
            Height          =   3000
            Left            =   60
            TabIndex        =   14
            Top             =   195
            Width           =   11190
            _ExtentX        =   19738
            _ExtentY        =   5292
            Cols0           =   13
            HighLight       =   1
            AllowUserResizing=   2
            EncabezadosNombres=   "Num-Tipo-SubTipo-Material-Estado-Observacion-PBruto-PNeto-Tasacion-TasAdicion-ObsAdicion-p-Item"
            EncabezadosAnchos=   "350-1030-1030-1030-1030-2200-650-650-900-850-1300-0-0"
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
            ColumnasAEditar =   "X-1-2-3-4-5-6-7-X-9-10-X-X"
            ListaControles  =   "0-3-3-3-3-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-L-L-L-R-R-C-R-L-R-C"
            FormatosEdit    =   "0-1-1-1-1-0-2-2-4-2-0-4-3"
            TextArray0      =   "Num"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   345
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin VB.Label Label1 
            Caption         =   "Total Balanza"
            Height          =   210
            Left            =   255
            TabIndex        =   20
            Top             =   3270
            Width           =   990
         End
         Begin VB.Label LblPBruto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   6735
            TabIndex        =   19
            Top             =   3255
            Width           =   675
         End
         Begin VB.Label LblPNeto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   7410
            TabIndex        =   18
            Top             =   3255
            Width           =   690
         End
         Begin VB.Label LblTasacion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   8085
            TabIndex        =   17
            Top             =   3255
            Width           =   720
         End
         Begin VB.Label LblTasacionAdic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   8790
            TabIndex        =   16
            Top             =   3255
            Width           =   795
         End
         Begin VB.Label lblPrestamo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   5940
            TabIndex        =   15
            Top             =   3255
            Visible         =   0   'False
            Width           =   705
         End
      End
      Begin VB.Frame FraContenedor 
         Height          =   600
         Index           =   1
         Left            =   90
         TabIndex        =   5
         Top             =   4740
         Width           =   11250
         Begin VB.ComboBox CboPlazo 
            Height          =   315
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   180
            Width           =   930
         End
         Begin SICMACT.EditMoney txtPrestamo 
            Height          =   285
            Left            =   9930
            TabIndex        =   6
            Top             =   165
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            BackColor       =   12648447
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox MskFechaVenc 
            Height          =   300
            Left            =   3660
            TabIndex        =   8
            Top             =   195
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            Caption         =   "&Plazo"
            Height          =   255
            Left            =   390
            TabIndex        =   11
            Top             =   210
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Vencimiento"
            Height          =   255
            Left            =   2205
            TabIndex        =   10
            Top             =   225
            Width           =   1410
         End
         Begin VB.Label Label4 
            Caption         =   "Préstamo"
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
            Left            =   8985
            TabIndex        =   9
            Top             =   210
            Width           =   870
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   60
      TabIndex        =   0
      Top             =   6360
      Width           =   3915
      Begin VB.CommandButton CmdVer 
         Caption         =   "Detalle"
         Height          =   285
         Left            =   3030
         TabIndex        =   1
         Top             =   210
         Width           =   765
      End
      Begin VB.Label LblDeudaP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1485
         TabIndex        =   3
         Top             =   180
         Width           =   1395
      End
      Begin VB.Label Label5 
         Caption         =   "Deuda Pendiente "
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   225
         Width           =   1305
      End
   End
   Begin SICMACT.ActXCodCta AxCodCta 
      Height          =   405
      Left            =   90
      TabIndex        =   34
      Top             =   60
      Width           =   3630
      _ExtentX        =   6403
      _ExtentY        =   714
      Texto           =   "Credito"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
End
Attribute VB_Name = "frmPigUsoLineaPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBuscar_Click()
Dim loPers As UPersona
Dim lsPersCod As String, lsPersNombre As String
Dim loColPFunc As dColPFunciones
Dim oPigFunciones As DPigFunciones
Dim oDatos As dPigContrato
Dim lstItemClte As ListItem
Dim rsTemp As Recordset

On Error GoTo ControlError

Set loPers = New UPersona
Set loPers = frmBuscaPersona.Inicio

If Not loPers Is Nothing Then
    lsPersCod = loPers.sPersCod
    
    Set lstItemClte = lstCliente.ListItems.Add(, , lsPersCod)
        lstItemClte.SubItems(1) = loPers.sPersNombre
        lstItemClte.SubItems(2) = loPers.sPersDireccDomicilio
        lstItemClte.SubItems(3) = loPers.sPersTelefono
        lstItemClte.SubItems(5) = gPersIdDNI
        lstItemClte.SubItems(6) = loPers.sPersIdnroDNI
        lstItemClte.SubItems(7) = gPersIdRUC
        lstItemClte.SubItems(8) = loPers.sPersIdnroRUC
        Set loColPFunc = New dColPFunciones
            lstItemClte.SubItems(4) = Trim(loColPFunc.dObtieneNombreZonaPersona(loPers.sPersCod))
        Set loColPFunc = Nothing

    'Carga Calificacion Pigno y Calificacion SBS de la Tabla ColocPigEvalCliente
    Set oPigFunciones = New DPigFunciones
    
    Set rsTemp = oPigFunciones.GetEvalCliente(lsPersCod)
'    If rsTemp.EOF And rsTemp.BOF Then
'        lstItemClte.SubItems(9) = "Cliente B1"
'        lnCalif = 4
'        lbNuevo = 0
'    Else
'        lstItemClte.SubItems(9) = rsTemp!cConsDescripcion
'        lbNuevo = 1
'        lnCalif = rsTemp!cCalifiCliente
'    End If

    Set rsTemp = Nothing
   
    Set rsTemp = oPigFunciones.GetCalifCliente(lsPersCod)
'    If rsTemp.EOF And rsTemp.BOF Then
'        lstItemClte.SubItems(10) = "Normal"
'        lnEval = 1
'    Else
'        lstItemClte.SubItems(10) = IIf(IsNull(rsTemp!cConsDescripcion), "Normal", rsTemp!cConsDescripcion)
'        lnEval = rsTemp!cEvalSBSCliente
'    End If
    
    Set rsTemp = Nothing
    Set oPigFunciones = Nothing
    
    '****** Deuda Pendiente del Cliente
    Set oDatos = New dPigContrato
'    lnDeudaPendiente = oDatos.dObtieneDeudaTotal(lsPersCod, gdFecSis)
'    LblDeudaP = lnDeudaPendiente
    Set oDatos = Nothing

Else
    Exit Sub
End If

Set loPers = Nothing


BuscaContratos (lsPersCod)
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Private Sub BuscaContratos(ByVal pscliente As String)
Dim lsEstados As String
Dim loPersContrato As DColPContrato
Dim loCuentas As UProdPersona
   
    lsEstados = gPigEstRematPRes
    
'    If Trim(lsPersCod) <> "" Then
'        Set loPersContrato = New DColPContrato
'            Set lrContratos = loPersContrato.dObtieneCredPigDePersona(lsPersCod, lsEstados)
'        Set loPersContrato = Nothing
'    End If
    
'    Set loCuentas = New UProdPersona
'        Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrContratos)
'        If loCuentas.sCtaCod <> "" Then
'            AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
'            AXCodCta.SetFocusCuenta
'        End If
'    Set loCuentas = Nothing
    
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & Err.Number & " " & Err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

