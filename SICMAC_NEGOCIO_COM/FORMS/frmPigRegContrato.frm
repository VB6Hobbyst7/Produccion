VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPigRegContrato 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Contrato"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11535
   Icon            =   "frmPigRegContrato.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   75
      TabIndex        =   33
      Top             =   6360
      Width           =   3915
      Begin VB.CommandButton CmdVer 
         Caption         =   "Detalle"
         Height          =   285
         Left            =   3030
         TabIndex        =   34
         Top             =   210
         Width           =   765
      End
      Begin VB.Label Label5 
         Caption         =   "Deuda Pendiente "
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   225
         Width           =   1305
      End
      Begin VB.Label LblDeudaP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1485
         TabIndex        =   35
         Top             =   180
         Width           =   1395
      End
   End
   Begin VB.Frame FraContenedor 
      Height          =   6390
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   -30
      Width           =   11430
      Begin VB.Frame FraContenedor 
         Height          =   600
         Index           =   1
         Left            =   75
         TabIndex        =   22
         Top             =   5205
         Width           =   11250
         Begin SICMACT.EditMoney txtPrestamo 
            Height          =   285
            Left            =   9930
            TabIndex        =   28
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
         Begin VB.ComboBox CboPlazo 
            Height          =   315
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   180
            Width           =   930
         End
         Begin MSMask.MaskEdBox MskFechaVenc 
            Height          =   300
            Left            =   3660
            TabIndex        =   25
            Top             =   195
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
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
            TabIndex        =   27
            Top             =   210
            Width           =   870
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Vencimiento"
            Height          =   255
            Left            =   2205
            TabIndex        =   26
            Top             =   225
            Width           =   1410
         End
         Begin VB.Label Label2 
            Caption         =   "&Plazo"
            Height          =   255
            Left            =   390
            TabIndex        =   23
            Top             =   210
            Width           =   495
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
         Height          =   3870
         Index           =   4
         Left            =   60
         TabIndex        =   12
         Top             =   1350
         Width           =   11310
         Begin VB.CommandButton CmdAgregarJ 
            Caption         =   "A&gregar"
            Height          =   315
            Left            =   9840
            TabIndex        =   21
            Top             =   3480
            Width           =   705
         End
         Begin VB.CommandButton cmdEliminarJ 
            Caption         =   "Eliminar"
            Enabled         =   0   'False
            Height          =   315
            Left            =   10560
            TabIndex        =   20
            Top             =   3480
            Width           =   675
         End
         Begin VB.TextBox TxtTotalB 
            Appearance      =   0  'Flat
            Height          =   255
            Left            =   1440
            TabIndex        =   15
            Top             =   3510
            Width           =   705
         End
         Begin SICMACT.FlexEdit FEJoyas 
            Height          =   3255
            Left            =   60
            TabIndex        =   13
            Top             =   195
            Width           =   11190
            _ExtentX        =   19738
            _ExtentY        =   5741
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
         Begin VB.Label lblPrestamo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   6045
            TabIndex        =   32
            Top             =   3510
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.Label LblTasacionAdic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   8895
            TabIndex        =   19
            Top             =   3510
            Width           =   795
         End
         Begin VB.Label LblTasacion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   8190
            TabIndex        =   18
            Top             =   3510
            Width           =   720
         End
         Begin VB.Label LblPNeto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   7515
            TabIndex        =   17
            Top             =   3510
            Width           =   690
         End
         Begin VB.Label LblPBruto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   6840
            TabIndex        =   16
            Top             =   3510
            Width           =   675
         End
         Begin VB.Label Label1 
            Caption         =   "Total Balanza"
            Height          =   210
            Left            =   360
            TabIndex        =   14
            Top             =   3525
            Width           =   990
         End
      End
      Begin VB.Frame FraContenedor 
         Enabled         =   0   'False
         Height          =   540
         Index           =   2
         Left            =   75
         TabIndex        =   8
         Top             =   5775
         Width           =   11265
         Begin SICMACT.EditMoney lblNetoRecibir 
            Height          =   285
            Left            =   10110
            TabIndex        =   29
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
            TabIndex        =   30
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
            TabIndex        =   31
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
            Caption         =   "Pago Mínimo"
            Height          =   255
            Index           =   16
            Left            =   210
            TabIndex        =   11
            Top             =   195
            Width           =   975
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Costo Tasación"
            Height          =   255
            Index           =   19
            Left            =   6270
            TabIndex        =   10
            Top             =   195
            Width           =   1200
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
            TabIndex        =   9
            Top             =   225
            Width           =   1140
         End
      End
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
         Height          =   1290
         Index           =   6
         Left            =   90
         TabIndex        =   4
         Top             =   90
         Width           =   11250
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            Enabled         =   0   'False
            Height          =   285
            Left            =   10200
            TabIndex        =   6
            Top             =   930
            Width           =   810
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "&Agregar"
            Height          =   285
            Left            =   9315
            TabIndex        =   5
            Top             =   930
            Width           =   810
         End
         Begin MSComctlLib.ListView lstCliente 
            Height          =   675
            Left            =   60
            TabIndex        =   7
            Top             =   210
            Width           =   11085
            _ExtentX        =   19553
            _ExtentY        =   1191
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
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   8205
      TabIndex        =   2
      Top             =   6525
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   10350
      TabIndex        =   1
      Top             =   6525
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   9285
      TabIndex        =   0
      Top             =   6525
      Width           =   975
   End
End
Attribute VB_Name = "frmPigRegContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'************************************************
'CMPCL - Registro de Contrato
'CAFF  - 23/08/2002
'************************************************
Dim lsPersCod As String

'Parametros para el Calculo del Valor de Tasacion
Dim pnPrestamoClteA1 As Double
Dim pnPrestamoClteA As Double
Dim pnPrestamoClteB As Double
Dim pnPrestamoClteB1 As Double

Dim pnPrestamoJoyaBue As Double
Dim pnPrestamoJoyaReg As Double
Dim pnPrestamoJoyaMal As Double

'Datos del Cliente
Dim lnCalif As Integer
Dim lnEval As Integer
Dim lnPOro As Currency
Dim lnNuevo As Integer

'** lstCliente - ListView
Dim lstItemClte As ListItem
Dim lsIteSel As String

'Carga datos de Joyas
Dim rsTipoJoya As Recordset
Dim rsSTipoJoya As Recordset
Dim rsMaterial As Recordset
Dim rsPlazos As Recordset
Dim rsEstadoJoya As Recordset

'Datos adicionales
Dim pnPagoMinimo As Double
Dim pnCostoTasac As Currency
Dim pnComSer As Currency
Dim pnPenalidad As Currency

Dim lnPBrutoT As Double
Dim lnPNetoT As Double
Dim lnTasacT As Currency
Dim lnTasacAdicT As Currency
Dim lnPrestamoT As Currency
Dim lbNuevo As Byte

Dim lnTasaComp As Double
Dim lnTasaMora As Double
Dim lsLineaCred As String
Dim lnPlazo As Integer

Dim lnPPTipoCte As Double
Dim lnPPEstJoya As Double
Dim lnEstJoya As Integer

Dim lsFonoAge As String
Dim nRanIni As Integer
Dim nRanFin As Integer
Dim lnJoyas As Integer

Private Sub cboPlazo_Click()
Dim lsFecha As String
Dim lnIntComp As Currency
Dim oPigCalculos As NPigCalculos
Dim oPigFunciones As DPigFunciones
Dim rs As Recordset
Dim lnMontoPrestamo As Currency

    If CboPlazo.Text <> "" And txtPrestamo > 0 Then
    
        lsFecha = DateAdd("d", CInt(CboPlazo.Text), gdFecSis)
        MskFechaVenc = lsFecha
        If lblCostoTasacion <> "" Then
            lblNetoRecibir = Format(CCur(txtPrestamo) - CCur(lblCostoTasacion), "#####,###.00")
        End If
        
        If txtPrestamo > 0 Then
            
            'Obtengo la Linea de credito con sus respectivas Tasas
            Set oPigFunciones = New DPigFunciones
            Set rs = oPigFunciones.GetLineaCredito(txtPrestamo)
            lnTasaComp = rs!TasaComp
            lnTasaMora = rs!TasaMora
            lsLineaCred = rs!cLineaCred
            Set oPigFunciones = Nothing
            lnPlazo = CboPlazo.Text
        
            Set oPigCalculos = New NPigCalculos
            lnMontoPrestamo = Round(CCur(txtPrestamo), 2)
            lnIntComp = Format(oPigCalculos.nCalculaIntCompensatorio(lnMontoPrestamo, lnTasaComp, lnPlazo), "#0.00")
            lblPagoMin = Format((oPigCalculos.CalcPagoMin(pnPagoMinimo, txtPrestamo, pnComSer) + lnIntComp), "#0.00")
            Set oPigCalculos = Nothing
        
        End If
        cmdGrabar.Enabled = True
        
    End If
End Sub

Private Sub cboplazo_KeyPress(KeyAscii As Integer)
Dim lsFecha As String
Dim oPigCalculos As NPigCalculos

If KeyAscii = 13 Then

    If CboPlazo.Text <> "" And txtPrestamo > 0 Then
        lsFecha = DateAdd("d", CInt(CboPlazo.Text), gdFecSis)
        MskFechaVenc = lsFecha
        lblNetoRecibir = CCur(txtPrestamo) - CCur(lblCostoTasacion)
        Set oPigCalculos = New NPigCalculos
        lblPagoMin = Format(oPigCalculos.CalcPagoMin(pnPagoMinimo, txtPrestamo, pnComSer), "#0.00")
        Set oPigCalculos = Nothing
        txtPrestamo.SetFocus
    End If
End If

End Sub

Private Sub CmdAgregar_Click()
Dim loPers As UPersona
Dim loColPFunc As dColPFunciones
Dim oPigFunciones As DPigFunciones
Dim oDatos As dPigContrato
Dim rsTemp As Recordset
Dim lnDeudaPendiente As Currency

lbNuevo = 0
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
    If rsTemp.EOF And rsTemp.BOF Then
        lstItemClte.SubItems(9) = "Cliente B1"
        lnCalif = 4
        lbNuevo = 0
    Else
        lstItemClte.SubItems(9) = rsTemp!cConsDescripcion
        lbNuevo = 1
        lnCalif = rsTemp!cCalifiCliente
    End If

    Set rsTemp = Nothing
   
    Set rsTemp = oPigFunciones.GetCalifCliente(lsPersCod)
    If rsTemp.EOF And rsTemp.BOF Then
        lstItemClte.SubItems(10) = "Normal"
        lnEval = 1
    Else
        lstItemClte.SubItems(10) = IIf(IsNull(rsTemp!cConsDescripcion), "Normal", rsTemp!cConsDescripcion)
        lnEval = rsTemp!cEvalSBSCliente
    End If
    
    Set rsTemp = Nothing
    Set oPigFunciones = Nothing
    
    '****** Deuda Pendiente del Cliente
    Set oDatos = New dPigContrato
    lnDeudaPendiente = oDatos.dObtieneDeudaTotal(lsPersCod, gdFecSis)
    LblDeudaP = lnDeudaPendiente
    Set oDatos = Nothing
Else
    Exit Sub
End If

Set loPers = Nothing

FraContenedor(4).Enabled = True

If lstCliente.ListItems(1).Text <> "" Then
    cmdAgregar.Enabled = False
    cmdEliminar.Enabled = True
End If
lnJoyas = 0
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Private Sub CmdAgregarJ_Click()

    If FEJoyas.Rows <= 20 Then
        If lnJoyas > 1 And FEJoyas.TextMatrix(FEJoyas.Row, 7) = "" Then
            MsgBox "Ingrese datos de la Joya anterior", vbInformation, "Aviso"
            Exit Sub
        Else
            If lnJoyas = 1 And FEJoyas.TextMatrix(FEJoyas.Row, 7) = "" Then
                MsgBox "Ingrese datos de la Joya anterior", vbInformation, "Aviso"
                Exit Sub
            Else
                lnJoyas = lnJoyas + 1
                FEJoyas.AdicionaFila
                If FEJoyas.Rows >= 2 Then
                   cmdEliminarJ.Enabled = True
                End If
                Dim oConst As DConstante
                Set oConst = New DConstante
                
                Set rsTipoJoya = oConst.RecuperaConstantes(gColocPigTipoJoya, , "C.cConsDescripcion")
                FEJoyas.CargaCombo rsTipoJoya
                Set rsTipoJoya = Nothing
                Set oConst = Nothing
                FEJoyas.SetFocus
                
            End If
        End If
    Else
        CmdAgregarJ.Enabled = False
        MsgBox "Sólo puede ingresar como máximo veinte piezas", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdCancelar_Click()
    Limpiar
End Sub

Private Sub cmdeliminar_Click()
Dim i As Integer, J As Integer

On Error GoTo ControlError

    If lstCliente.ListItems.Count = 0 Then
       MsgBox "No existen datos a eliminar", vbInformation, "Aviso"
       cmdEliminar.Enabled = False
       lstCliente.SetFocus
       Exit Sub
    Else
        i = 1
        lstCliente.ListItems.Remove (i)
    End If
    lstCliente.SetFocus
    If lstCliente.ListItems.Count = 0 Then
        cmdEliminar.Enabled = False
    End If
    
    cmdAgregar.Enabled = Not cmdAgregar.Enabled
    
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdEliminarJ_Click()
    
    FEJoyas.EliminaFila FEJoyas.Row
    If FEJoyas.Rows <= 20 Then
        CmdAgregarJ.Enabled = True
    End If
    lnJoyas = lnJoyas + 1
    SumaColumnas
    cboPlazo_Click
  
End Sub

Private Sub cmdGrabar_Click()
Dim oPigFunciones As DPigFunciones
Dim oPigCalculos As NPigCalculos
Dim oContFunc As NContFunciones
Dim oRegPig As NPigContrato
Dim oPrevio As Previo.clsPrevio
Dim oPigImpre As nPigImpre
Dim rs As Recordset
Dim rsPer As ADODB.Recordset
Dim rsJoyas As Recordset
Dim lnSaldo As Currency
Dim lnPlazo As Integer
Dim lsFechaVenc As String
Dim lnCostoTasac As Currency, lnMontoTasac As Currency, lnNetoRec As Currency
Dim lnIntComp As Currency
Dim lnPiezas As Integer
Dim lnPNeto As Double, lnPBruto As Double
Dim lsMovNro As String
Dim lsContrato As String
Dim lsCodPers As String, lsNomPers As String, lsDocId As String, lsFono As String, lsDir As String
Dim lsFechaHoraGrab As String
Dim lnMontoPrestamo As Currency
Dim lnTotalBalanza As Currency
Dim lsCadImpre As String
Dim lnComServ As Currency

If Not ValidaPrestamo Then Exit Sub

Set rsPer = fgGetCodigoPersonaListaRsNew(lstCliente)
lsCodPers = rsPer!cPersCod

Set rsPer = Nothing

lnSaldo = CCur(lblNetoRecibir)
lnPlazo = Val(CboPlazo.Text)
lsFechaVenc = Format$(Me.MskFechaVenc, "mm/dd/yyyy")
lnCostoTasac = Format(CCur(lblCostoTasacion), "#0.00")
lnMontoTasac = Format(CCur(LblTasacion) + CCur(IIf(LblTasacionAdic = "", 0, LblTasacionAdic)), "#0.00")
lnMontoPrestamo = Format(CCur(txtPrestamo), "#0.00")
lnTotalBalanza = Format(CCur(TxtTotalB), "#0.00")
lnNetoRec = Format(CCur(lblNetoRecibir), "#0.00")
lnPNeto = CDbl(LblPNeto)
lnPBruto = CDbl(LblPBruto)

Set oPigCalculos = New NPigCalculos
lnIntComp = Format(oPigCalculos.nCalculaIntCompensatorio(lnMontoPrestamo, lnTasaComp, lnPlazo), "#0.00")
If FEJoyas.TextMatrix(FEJoyas.Rows - 1, 8) = "" Then
    FEJoyas.EliminaFila (FEJoyas.Rows - 1)
End If
Set oPigCalculos = Nothing
lnPiezas = FEJoyas.TextMatrix(FEJoyas.Rows - 1, 0)
Set rsJoyas = FEJoyas.GetRsNew

If MsgBox("¿Desea Grabar Contrato Prestamo Pignoraticio? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    'Genera Mov Nro
    Set oContFunc = New NContFunciones
        lsMovNro = oContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set oContFunc = Nothing
    
    lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)

   If lbNuevo = 1 Then lnCalif = -1
    
    Set oRegPig = New NPigContrato
        lsContrato = oRegPig.nRegistraContratoPigno(gsCodCMAC & gsCodAge, gMonedaNacional, lnTasaComp, lnTasaMora, lnSaldo, _
            lsFechaHoraGrab, lsCodPers, lnPlazo, lsFechaVenc, lsMovNro, lnCostoTasac, rsJoyas, lnPiezas, gsCodAge, gsCodPersUser, _
            lsLineaCred, lnIntComp, lnMontoTasac, pnPenalidad, lnMontoPrestamo, lnTotalBalanza, lnCalif, lnEval)
        
    Set oRegPig = Nothing
    MsgBox "Se ha generado Contrato Nro " & lsContrato, vbInformation, "Aviso"

   If MsgBox("Imprimir Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
        'Datos de la Persona
        lsNomPers = lstCliente.ListItems(1).SubItems(1)
        lsFono = lstCliente.ListItems(1).SubItems(3)
        lsDir = Trim(lstCliente.ListItems(1).SubItems(2)) & Space(2) & Trim(lstCliente.ListItems(1).SubItems(4))
        lsDocId = "DNI" & Space(2) & lstCliente.ListItems(1).SubItems(6)
        
        Set oPigImpre = New nPigImpre
        
            lsCadImpre = oPigImpre.ImpreContratoPignoraticio(lsContrato, False, lsCodPers, lsNomPers, lsDocId, lsFono, lsDir, lnPlazo, _
                CStr(gdFecSis), lsFechaVenc, gsCodUser, "SOLES", gsNomAge, rsJoyas, lnMontoPrestamo, lnCostoTasac, lnNetoRec, lnPiezas, lnPBruto, _
                lnPNeto, lnMontoTasac, pnComSer, lnTasaComp, 0, lsFonoAge)
            
        Set oPigImpre = Nothing
        Set oPrevio = New Previo.clsPrevio
            oPrevio.PrintSpool sLpt, lsCadImpre, False, 66

            Do While True
                If MsgBox("Reimprimir Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                    oPrevio.PrintSpool sLpt, lsCadImpre, False, 66
                Else
                    Set oPrevio = Nothing
                    Exit Do
                End If
            Loop
    End If
End If
Limpiar

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
    Limpiar

End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub
Private Sub cmdVer_Click()
    frmPigDeudaPendiente.Inicia (lsPersCod)
End Sub

Private Sub FEJoyas_Click()
Dim oConst As DConstante

Set oConst = New DConstante

Select Case FEJoyas.Col
Case 1
    Set rsTipoJoya = oConst.RecuperaConstantes(gColocPigTipoJoya, , "C.cConsDescripcion")
    FEJoyas.CargaCombo rsTipoJoya
    Set rsTipoJoya = Nothing
Case 2
    Set rsSTipoJoya = oConst.RecuperaConstantes(gColocPigSubTipoJoya, , "C.cConsDescripcion")
    FEJoyas.CargaCombo rsSTipoJoya
    Set rsSTipoJoya = Nothing
Case 3
    Set rsMaterial = oConst.RecuperaConstantes(gColocPigMaterial, , "C.cConsDescripcion")
    FEJoyas.CargaCombo rsMaterial
    Set rsMaterial = Nothing
Case 4
    Set rsEstadoJoya = oConst.RecuperaConstantes(gColocPigEstConservaJoya, , "C.cConsDescripcion")
    FEJoyas.CargaCombo rsEstadoJoya
    Set rsEstadoJoya = Nothing
End Select

Set oConst = Nothing
End Sub

Private Sub feJoyas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If FEJoyas.Col = 1 Then
        If FEJoyas.TextMatrix(FEJoyas.Row, 9) <> "" Then
            CmdAgregarJ.SetFocus
            CmdAgregarJ_Click
        End If
    End If
End If
End Sub

Private Sub FEJoyas_OnCellChange(pnRow As Long, pnCol As Long)
Dim oPigCalculos As NPigCalculos
Dim oPigFunciones As DPigFunciones
Dim lnPrestamo As Currency
Dim lnTasacionT As Currency

    FEJoyas.TextMatrix(FEJoyas.Row, 12) = FEJoyas.TextMatrix(FEJoyas.Row, 0)
        
    If FEJoyas.Col = 3 Then     'Tipo de Material
        
        If FEJoyas.TextMatrix(FEJoyas.Row, 3) <> "" And FEJoyas.TextMatrix(FEJoyas.Row, 4) <> "" And FEJoyas.TextMatrix(FEJoyas.Row, 7) <> "" Then
    
            CalculaTasacion
        
            If FEJoyas.TextMatrix(FEJoyas.Row, 9) <> "" Then
                lnTasacionT = CCur(IIf(FEJoyas.TextMatrix(FEJoyas.Row, 8) = "", 0, FEJoyas.TextMatrix(FEJoyas.Row, 8))) + CCur(IIf(FEJoyas.TextMatrix(FEJoyas.Row, 9) = "", 0, FEJoyas.TextMatrix(FEJoyas.Row, 9)))
                Set oPigCalculos = New NPigCalculos
                FEJoyas.TextMatrix(FEJoyas.Row, 11) = Format$(oPigCalculos.CalcValorPrestamo(lnPPTipoCte, lnTasacionT), "#####.00")
                Set oPigCalculos = Nothing
            End If
        
        End If
        
    End If
    
    If FEJoyas.Col = 4 Then     'Estado de la Joya
    
        If FEJoyas.TextMatrix(FEJoyas.Row, 3) <> "" Then
            Set oPigFunciones = New DPigFunciones
            lnPOro = oPigFunciones.GetPrecioMaterial(1, Val(Right(FEJoyas.TextMatrix(FEJoyas.Row, 3), 3)), 1)
            Set oPigFunciones = Nothing
        End If
    
        If FEJoyas.TextMatrix(FEJoyas.Row, 3) <> "" And FEJoyas.TextMatrix(FEJoyas.Row, 4) <> "" And FEJoyas.TextMatrix(FEJoyas.Row, 7) <> "" Then
            
            CalculaTasacion
        
            If FEJoyas.TextMatrix(FEJoyas.Row, 9) <> "" Then
                
                lnTasacionT = CCur(IIf(FEJoyas.TextMatrix(FEJoyas.Row, 8) = "", 0, FEJoyas.TextMatrix(FEJoyas.Row, 8))) + CCur(IIf(FEJoyas.TextMatrix(FEJoyas.Row, 9) = "", 0, FEJoyas.TextMatrix(FEJoyas.Row, 9)))
                Set oPigCalculos = New NPigCalculos
                FEJoyas.TextMatrix(FEJoyas.Row, 11) = Format$(oPigCalculos.CalcValorPrestamo(lnPPTipoCte, lnTasacionT), "#####.00")
                Set oPigCalculos = Nothing
            
            End If
            
        End If
                        
        If FEJoyas.TextMatrix(FEJoyas.Row, 4) <> "" Then
             lnEstJoya = Val(Right(FEJoyas.TextMatrix(FEJoyas.Row, 4), 3))
            
             Select Case lnEstJoya   'Porcentaje del Prestamo por Estado de la Joya
                 Case 1
                     lnPPEstJoya = pnPrestamoJoyaBue
                 Case 2
                     lnPPEstJoya = pnPrestamoJoyaReg
                 Case 3
                     lnPPEstJoya = pnPrestamoJoyaMal
             End Select
        
        End If
        
    End If
    
    If FEJoyas.Col = 6 Then
        If FEJoyas.TextMatrix(FEJoyas.Row, 6) <> "" Then
            If CCur(FEJoyas.TextMatrix(FEJoyas.Row, 6)) < 0 Then
                MsgBox "Peso Bruto no puede ser negativo", vbInformation, "Aviso"
                FEJoyas.TextMatrix(FEJoyas.Row, 6) = 0
            End If
        End If
    End If
    
    If FEJoyas.Col = 7 Then     'Peso Neto

        If FEJoyas.TextMatrix(FEJoyas.Row, 7) <> "" Then
            If CCur(FEJoyas.TextMatrix(FEJoyas.Row, 7)) < 0 Then
                MsgBox "Peso Neto no puede ser negativo", vbInformation, "Aviso"
                FEJoyas.TextMatrix(FEJoyas.Row, 7) = 0
            Else
                If CCur(FEJoyas.TextMatrix(FEJoyas.Row, 7)) > CCur(FEJoyas.TextMatrix(FEJoyas.Row, 6)) Then
                    MsgBox "Peso Neto no puede ser mayor que Peso Bruto", vbInformation, "Aviso"
                    FEJoyas.TextMatrix(FEJoyas.Row, 7) = 0
                Else
                    CalculaTasacion
                End If
            End If
        End If
        cboPlazo_Click
        
    End If

    If FEJoyas.Col = 9 Then     'Valor de Tasación --- Tasacion Adicional
    
        If FEJoyas.TextMatrix(FEJoyas.Row, 9) <> "" Then
            If CCur(FEJoyas.TextMatrix(FEJoyas.Row, 9)) < 0 Then
                MsgBox "Valor de Tasacion no puede ser negativo", vbInformation, "Aviso"
                FEJoyas.TextMatrix(FEJoyas.Row, 9) = 0
            Else
                Select Case lnCalif 'Porcentaje de Prestamo por Tipo de Calificacion del Cliente
                Case 1 'A1
                    lnPPTipoCte = pnPrestamoClteA1
                Case 2
                    lnPPTipoCte = pnPrestamoClteA
                Case 3
                    lnPPTipoCte = pnPrestamoClteB
                Case 4
                    lnPPTipoCte = pnPrestamoClteB1
                End Select
        
                lnTasacionT = CCur(IIf(FEJoyas.TextMatrix(FEJoyas.Row, 8) = "", 0, FEJoyas.TextMatrix(FEJoyas.Row, 8))) + CCur(IIf(FEJoyas.TextMatrix(FEJoyas.Row, 9) = "", 0, FEJoyas.TextMatrix(FEJoyas.Row, 9)))
                Set oPigCalculos = New NPigCalculos
                FEJoyas.TextMatrix(FEJoyas.Row, 11) = Format$(oPigCalculos.CalcValorPrestamo(lnPPTipoCte, lnTasacionT), "#####.00")
                Set oPigCalculos = Nothing
                cboPlazo_Click
            End If
        End If
        
    End If

    SumaColumnas
    cboPlazo_Click
    
End Sub

Private Sub FEJoyas_RowColChange()
Dim oConst As DConstante
Dim lnTasacionT As Currency
Dim oPigCalculos As NPigCalculos


    Set oConst = New DConstante
    
    Select Case FEJoyas.Col
    Case 1
        Set rsTipoJoya = oConst.RecuperaConstantes(gColocPigTipoJoya, , "C.cConsDescripcion")
        FEJoyas.CargaCombo rsTipoJoya
        Set rsTipoJoya = Nothing
    Case 2
        Set rsSTipoJoya = oConst.RecuperaConstantes(gColocPigSubTipoJoya, , "C.cConsDescripcion")
        FEJoyas.CargaCombo rsSTipoJoya
        Set rsSTipoJoya = Nothing
    Case 3
        Set rsMaterial = oConst.RecuperaConstantes(gColocPigMaterial, , "C.cConsDescripcion")
        FEJoyas.CargaCombo rsMaterial
        Set rsMaterial = Nothing
        
        If FEJoyas.TextMatrix(FEJoyas.Row, 3) <> "" And FEJoyas.TextMatrix(FEJoyas.Row, 4) <> "" And FEJoyas.TextMatrix(FEJoyas.Row, 7) <> "" Then
    
            CalculaTasacion
        
            If FEJoyas.TextMatrix(FEJoyas.Row, 9) <> "" Then
                lnTasacionT = CCur(IIf(FEJoyas.TextMatrix(FEJoyas.Row, 8) = "", 0, FEJoyas.TextMatrix(FEJoyas.Row, 8))) + CCur(IIf(FEJoyas.TextMatrix(FEJoyas.Row, 9) = "", 0, FEJoyas.TextMatrix(FEJoyas.Row, 9)))
                Set oPigCalculos = New NPigCalculos
                FEJoyas.TextMatrix(FEJoyas.Row, 11) = Format$(oPigCalculos.CalcValorPrestamo(lnPPTipoCte, lnTasacionT), "#####.00")
                Set oPigCalculos = Nothing
            End If
        
        End If
        
        
    Case 4
        Set rsEstadoJoya = oConst.RecuperaConstantes(gColocPigEstConservaJoya, , "C.cConsDescripcion")
        FEJoyas.CargaCombo rsEstadoJoya
        Set rsEstadoJoya = Nothing
    
    End Select
    
    Set oConst = Nothing

End Sub

Private Sub Form_Load()

CargaPlazos
CargaParametros
lblCostoTasacion = pnCostoTasac
CboPlazo.ListIndex = -1
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub CargaPlazos()
Dim oPigFuncion As DPigFunciones
Dim i As Integer

Dim rs As Recordset

    Set oPigFuncion = New DPigFunciones
    
    nRanIni = oPigFuncion.GetParamValor(gPigParamPrestamoPlazoMin)
    nRanFin = oPigFuncion.GetParamValor(gPigParamPrestamoPlazoMax)
    
    Set rs = Nothing
    
    Set oPigFuncion = Nothing
    
    For i = nRanIni To nRanFin
        CboPlazo.AddItem i
    Next i
    
    CboPlazo.ListIndex = 0

End Sub

Private Sub SumaColumnas()
Dim i As Integer
Dim oPigCalculos As NPigCalculos
    
    lnPBrutoT = 0:      lnPNetoT = 0:       lnTasacT = 0:         lnTasacAdicT = 0:     lnPrestamoT = 0
    Select Case FEJoyas.Col
    
    Case 6      'PESO BRUTO
        lnPBrutoT = FEJoyas.SumaRow(6)
        LblPBruto.Caption = Format$(lnPBrutoT, "#0.00")
        TxtTotalB.Text = LblPBruto.Caption

    Case 7      'PESO NETO
        lnPNetoT = FEJoyas.SumaRow(7)
        LblPNeto.Caption = Format$(lnPNetoT, "#0.00")
        
        lnTasacT = FEJoyas.SumaRow(8)
        LblTasacion.Caption = Format$(lnTasacT, "#0.00")
    
        lnPrestamoT = FEJoyas.SumaRow(11)
        lblPrestamo.Caption = Format$(lnPrestamoT, "#0.00")
    
    Case 9      'TASACION ADICIONAL
        lnTasacAdicT = FEJoyas.SumaRow(9)
        LblTasacionAdic.Caption = Format$(lnTasacAdicT, "#0.00")
        lnPrestamoT = FEJoyas.SumaRow(11)
        lblPrestamo.Caption = Format$(lnPrestamoT, "#0.00")
        
    Case Else
        lnPNetoT = FEJoyas.SumaRow(7)
        LblPNeto.Caption = Format$(lnPNetoT, "#0.00")
        
        lnTasacT = FEJoyas.SumaRow(8)
        LblTasacion.Caption = Format$(lnTasacT, "#0.00")
    
        lnPrestamoT = FEJoyas.SumaRow(11)
        lblPrestamo.Caption = Format$(lnPrestamoT, "#0.00")
        
        lnTasacAdicT = FEJoyas.SumaRow(9)
        LblTasacionAdic.Caption = Format$(lnTasacAdicT, "#0.00")
        lnPrestamoT = FEJoyas.SumaRow(11)
        lblPrestamo.Caption = Format$(lnPrestamoT, "#0.00")
    
    End Select

    If lblPrestamo <> "" Then
        txtPrestamo = CCur(lblPrestamo)
    End If
    
End Sub

Private Sub CargaParametros()
Dim oPigFunciones As DPigFunciones
Dim rs As Recordset

    Set oPigFunciones = New DPigFunciones
    
    pnPrestamoClteA1 = oPigFunciones.GetParamValor(gPigParamPorPrestamoCteA1)
    pnPrestamoClteA = oPigFunciones.GetParamValor(gPigParamPorPrestamoCteA)
    pnPrestamoClteB = oPigFunciones.GetParamValor(gPigParamPorPrestamoCteB)
    pnPrestamoClteB1 = oPigFunciones.GetParamValor(gPigParamPorPrestamoCteB1)
    pnPrestamoJoyaBue = oPigFunciones.GetParamValor(gPigParamPorPrestamoJoyaBue)
    pnPrestamoJoyaReg = oPigFunciones.GetParamValor(gPigParamPorPrestamoJoyaReg)
    pnPrestamoJoyaMal = oPigFunciones.GetParamValor(gPigParamPorPrestamoJoyaMal)
    pnPagoMinimo = oPigFunciones.GetParamValor(gPigParamAmortizaMin)
    Set rs = oPigFunciones.GetConceptoValor(gColPigConceptoCodTasacion)
    pnCostoTasac = rs!nValor
    Set rs = Nothing
    Set rs = oPigFunciones.GetConceptoValor(gColPigConceptoCodComiServ)
    pnComSer = rs!nValor
    Set rs = Nothing
    Set rs = oPigFunciones.GetConceptoValor(gColPigConceptoCodPenalidad)
    pnPenalidad = rs!nValor
    Set rs = Nothing
    lsFonoAge = oPigFunciones.GetFonoAge(gsCodAge)
    Set oPigFunciones = Nothing
    
End Sub

Private Sub MskFechaVenc_KeyPress(KeyAscii As Integer)
Dim lnDias As Integer
    If KeyAscii = 13 Then
        If txtPrestamo > 0 Then
        
            If IsDate(MskFechaVenc.Text) Then
        
                Dim oPigCalculos As NPigCalculos
                Set oPigCalculos = New NPigCalculos

                lnDias = DateDiff("d", gdFecSis, MskFechaVenc.Text)
                If lnDias >= nRanIni Then
                    CboPlazo.ListIndex = lnDias - nRanIni
                Else
                    MsgBox "Plazo de Vencimiento no válido", vbInformation, "Aviso"
                    Exit Sub
                End If
                lblNetoRecibir = CCur(txtPrestamo) - CCur(pnCostoTasac)
                lblPagoMin = oPigCalculos.CalcPagoMin(pnPagoMinimo, txtPrestamo, pnComSer)
                cmdGrabar.Enabled = True
            Else
                MsgBox "Fecha no válida", vbInformation, "Aviso"
                MskFechaVenc.SetFocus
            End If
    End If
    End If
    
End Sub

Private Sub txtPrestamo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

    If txtPrestamo > 0 Then
        If ValidaPrestamo Then
            If CboPlazo.ListIndex <> -1 Then
                cboPlazo_Click
                cmdGrabar.Enabled = True
                cmdGrabar.SetFocus
            End If
        Else
            txtPrestamo.SetFocus
        End If
    Else
        MsgBox "Monto del Prestamo no puede ser menor que 1.00", vbInformation, "Aviso"
        txtPrestamo.SetFocus
    End If
End If
End Sub

Private Function ValidaPrestamo() As Boolean

    If txtPrestamo <> "" And CCur(txtPrestamo) <= Format(lnPrestamoT, "#0.00") Then
        ValidaPrestamo = True
    Else
        MsgBox "Monto del Prestamo no puede ser mayor que " & CStr(Format(lnPrestamoT, "#0.00")), vbInformation, "Aviso"
        txtPrestamo = lnPrestamoT
        ValidaPrestamo = False
    End If
    
End Function

Private Sub TxtTotalB_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    FraContenedor(1).Enabled = True
    CboPlazo.SetFocus
Else
    KeyAscii = NumerosDecimales(TxtTotalB, KeyAscii, , , False)
End If

End Sub

Private Sub Limpiar()

LblPBruto = ""
LblPNeto = ""
LblDeudaP = ""
lblNetoRecibir = ""
lblPagoMin = ""
lblPrestamo = ""
LblTasacion = ""
LblTasacionAdic = ""
FEJoyas.Clear
FEJoyas.Rows = 2
FEJoyas.FormaCabecera
lnNuevo = 0
TxtTotalB = ""
MskFechaVenc.Text = "__/__/____"
CboPlazo.ListIndex = -1
lstCliente.ListItems.Clear
txtPrestamo = ""
cmdAgregar.Enabled = True
cmdEliminar.Enabled = False
cmdGrabar.Enabled = False

End Sub

Private Sub CalculaTasacion()
Dim oPigCalculos As NPigCalculos
Dim lnPPTipoCte As Double
Dim lnPPEstJoya As Double
Dim lnEstJoya As Integer
Dim oPigFunciones As DPigFunciones
Dim lnPrestamo As Currency
Dim lnTasacionT As Currency

    Set oPigFunciones = New DPigFunciones
    lnPOro = oPigFunciones.GetPrecioMaterial(1, Val(Right(FEJoyas.TextMatrix(FEJoyas.Row, 3), 3)), 1)
    
    If lnPOro <= 0 Then
        FEJoyas.TextMatrix(FEJoyas.Row, 8) = 0
        FEJoyas.TextMatrix(FEJoyas.Row, 11) = 0
        MsgBox "Precio del Material No ha sido ingresado en el Tarifario, actualice el Tarifario", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Set oPigFunciones = Nothing
     
    lnEstJoya = Val(Right(FEJoyas.TextMatrix(FEJoyas.Row, 4), 3))
     
     Select Case lnEstJoya   'Porcentaje del Prestamo por Estado de la Joya
         Case 1
             lnPPEstJoya = pnPrestamoJoyaBue
         Case 2
             lnPPEstJoya = pnPrestamoJoyaReg
         Case 3
             lnPPEstJoya = pnPrestamoJoyaMal
     End Select
        
    Select Case lnCalif 'Porcentaje de Prestamo por Tipo de Calificacion del Cliente
        Case 1 'A1
            lnPPTipoCte = pnPrestamoClteA1
        Case 2
            lnPPTipoCte = pnPrestamoClteA
        Case 3
            lnPPTipoCte = pnPrestamoClteB
        Case 4
            lnPPTipoCte = pnPrestamoClteB1
    End Select
        
    Set oPigCalculos = New NPigCalculos     'calculo del valor de tasación
    FEJoyas.TextMatrix(FEJoyas.Row, 8) = Format$(oPigCalculos.CalcValorTasacion(lnPPEstJoya, Val(FEJoyas.TextMatrix(FEJoyas.Row, 7)), lnPOro), "#####.00")
    FEJoyas.TextMatrix(FEJoyas.Row, 11) = Format$(oPigCalculos.CalcValorPrestamo(lnPPTipoCte, FEJoyas.TextMatrix(FEJoyas.Row, 8)), "#####.00")
    Set oPigCalculos = Nothing
        
End Sub

