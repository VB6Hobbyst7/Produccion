VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmColRecNegCalculaCalendario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recuperaciones - Calendario"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   ForeColor       =   &H80000002&
   Icon            =   "frmColRecNegCalculaCalendario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraCalendarioLibre 
      Caption         =   "Calendario Libre"
      Height          =   2130
      Left            =   4395
      TabIndex        =   55
      Top             =   4140
      Width           =   2085
      Begin VB.CommandButton cmdCalAceptar 
         Caption         =   "&Aceptar"
         Height          =   330
         Left            =   360
         TabIndex        =   60
         Top             =   1710
         Width           =   1275
      End
      Begin VB.TextBox txtCalMonto 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   315
         TabIndex        =   58
         Top             =   1305
         Width           =   1305
      End
      Begin VB.CommandButton cmdCalEliminar 
         Caption         =   "&Eliminar"
         Height          =   330
         Left            =   315
         TabIndex        =   57
         Top             =   630
         Width           =   1275
      End
      Begin VB.CommandButton cmdCalAgregar 
         Caption         =   "A&gregar"
         Height          =   330
         Left            =   315
         TabIndex        =   56
         Top             =   270
         Width           =   1275
      End
      Begin MSMask.MaskEdBox txtCalFecha 
         Height          =   300
         Left            =   315
         TabIndex        =   59
         Top             =   1035
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Deuda Proyectada"
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
      Height          =   2055
      Left            =   4200
      TabIndex        =   40
      Top             =   4200
      Visible         =   0   'False
      Width           =   2400
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   960
         TabIndex        =   52
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Label lblComisionCalculada 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   960
         TabIndex        =   51
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblIntCalculado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   960
         TabIndex        =   50
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Comision"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Total"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   42
         Top             =   1080
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "Interes"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Saldos Actuales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   30
      Top             =   1920
      Width           =   2535
      Begin VB.Label lblDeudaFecha 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1080
         TabIndex        =   49
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblGastos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1080
         TabIndex        =   48
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblIntMor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1080
         TabIndex        =   47
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblIntComp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1080
         TabIndex        =   46
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblCapital 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1080
         TabIndex        =   45
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Deuda Fec"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   43
         Top             =   1680
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Int.Morat"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Int.Comp"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "De Capital"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Gastos"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   31
         Top             =   1320
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   405
      Left            =   5640
      TabIndex        =   12
      Top             =   6420
      Width           =   855
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   405
      Left            =   1380
      TabIndex        =   13
      Top             =   6420
      Width           =   855
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   405
      Left            =   4620
      TabIndex        =   11
      Top             =   6420
      Width           =   855
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   3540
      TabIndex        =   1
      Top             =   6420
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Frame fraCalendario 
      Caption         =   "Calendario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   2250
      Left            =   120
      TabIndex        =   15
      Top             =   4080
      Width           =   3975
      Begin MSComctlLib.ListView lvwCalendario 
         Height          =   1920
         Left            =   75
         TabIndex        =   27
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   3387
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Monto"
            Object.Width           =   1940
         EndProperty
      End
   End
   Begin VB.Frame fraNegociacion1 
      Caption         =   "Negociación"
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
      Height          =   2175
      Left            =   2790
      TabIndex        =   10
      Top             =   1920
      Width           =   3825
      Begin VB.CommandButton CmdCalcular 
         Caption         =   "Calcular"
         Height          =   315
         Left            =   2790
         TabIndex        =   54
         Top             =   555
         Width           =   810
      End
      Begin VB.CommandButton cmdDehacer 
         Caption         =   "&Dehacer"
         Height          =   345
         Left            =   2280
         TabIndex        =   9
         Top             =   1680
         Width           =   975
      End
      Begin VB.Frame fraTipCalendario 
         Height          =   945
         Left            =   120
         TabIndex        =   39
         Top             =   1155
         Width           =   1935
         Begin VB.OptionButton optCuotaLibre 
            Caption         =   "Cuota Libre"
            Height          =   195
            Left            =   135
            TabIndex        =   61
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtFechaFija 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1320
            MaxLength       =   2
            TabIndex        =   7
            Top             =   405
            Width           =   495
         End
         Begin VB.OptionButton optFecha 
            Caption         =   "Fecha Fija"
            Height          =   195
            Left            =   135
            TabIndex        =   6
            Top             =   450
            Width           =   1215
         End
         Begin VB.TextBox txtPeriodoFijo 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1320
            MaxLength       =   2
            TabIndex        =   5
            Text            =   "30"
            Top             =   120
            Width           =   495
         End
         Begin VB.OptionButton optPeriodo 
            Caption         =   "Periodo Fijo"
            Height          =   195
            Left            =   135
            TabIndex        =   4
            Top             =   180
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.TextBox txtNegCuotas 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1440
         TabIndex        =   3
         Top             =   840
         Width           =   1305
      End
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "A&plicar"
         Height          =   345
         Left            =   2265
         TabIndex        =   8
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtNegMonto 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1440
         TabIndex        =   2
         Top             =   540
         Width           =   1305
      End
      Begin MSMask.MaskEdBox TxtNegVigencia 
         Height          =   300
         Left            =   1440
         TabIndex        =   29
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label11 
         Caption         =   "Nro Cuotas"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   840
         Width           =   915
      End
      Begin VB.Label Label9 
         Caption         =   "Vigencia"
         Height          =   225
         Left            =   240
         TabIndex        =   24
         Top             =   300
         Width           =   705
      End
      Begin VB.Label Label7 
         Caption         =   "Monto Neg."
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   540
         Width           =   1425
      End
   End
   Begin VB.Frame fraCredito 
      Caption         =   "Credito"
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
      Height          =   1785
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar ..."
         Height          =   315
         Left            =   3840
         TabIndex        =   14
         Top             =   315
         Width           =   855
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   465
         Left            =   120
         TabIndex        =   53
         Top             =   240
         Width           =   3615
         _extentx        =   6376
         _extenty        =   820
         texto           =   "Crédito"
         enabledcta      =   -1  'True
         enabledprod     =   -1  'True
         enabledage      =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Comision"
         Height          =   240
         Index           =   5
         Left            =   4260
         TabIndex        =   37
         Top             =   900
         Width           =   735
      End
      Begin VB.Label lblComision 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   5040
         TabIndex        =   36
         Top             =   900
         Width           =   1140
      End
      Begin VB.Label lblTasaInteres 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   5040
         TabIndex        =   28
         Top             =   1170
         Width           =   1140
      End
      Begin VB.Label lblEstudioJuridico 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1080
         TabIndex        =   26
         Top             =   1140
         Width           =   3015
      End
      Begin VB.Label lblCondicionCredito 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1080
         TabIndex        =   25
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Condicion"
         Height          =   225
         Index           =   3
         Left            =   240
         TabIndex        =   23
         Top             =   1440
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Estado"
         Height          =   225
         Index           =   1
         Left            =   4320
         TabIndex        =   21
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "Est. Jurid."
         Height          =   240
         Index           =   4
         Left            =   195
         TabIndex        =   20
         Top             =   1140
         Width           =   735
      End
      Begin VB.Label lblNombreCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1080
         TabIndex        =   19
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Tasa Int."
         Height          =   240
         Index           =   2
         Left            =   4260
         TabIndex        =   18
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblEstadoCredito 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   5040
         TabIndex        =   17
         Top             =   1440
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente "
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   840
         Width           =   660
      End
   End
   Begin RichTextLib.RichTextBox rtfImp 
      Height          =   315
      Left            =   360
      TabIndex        =   38
      Top             =   5820
      Visible         =   0   'False
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   556
      _Version        =   393217
      TextRTF         =   $"frmColRecNegCalculaCalendario.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmColRecNegCalculaCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*  APLICACION : Recuperaciones - Negociaciones
'*  CREACION: 15/06/2004      AUTOR :LAYG
'*  RESUMEN: PERMITE REGISTRAR LA NEGOCIACION DIRECTA PARA RECUPARAR UN CREDITO
'*************************************************************************
Option Explicit

Dim dUltimaCuota As Date
Dim fsCodCta As String
Dim lbConexion As Boolean
Dim fbNuevaNegoc As Boolean
Dim fnContCuota As Integer
Public fntipo As Integer
Dim fsNegAnterior As String
Dim fnCuotaSeleccionada As Integer
Dim fnPorComisionAbo As Double
Dim fnFormaCalcIntComp As Integer, fnFormaCalcIntMora As Integer
Dim fnTipoCalcIntComp As Integer, fnTipoCalcIntMora As Integer



Public Function Inicio(ByVal pnTipo As Integer, Optional psCodCta As String = "")
 fntipo = pnTipo
 
 If pnTipo = 1 Then  ' Simulador
    LimpiaDatos
 ElseIf pnTipo = 2 Then ' Registro Negociacion
    cmdNuevo.Visible = False
    cmdSalir.Caption = "Aceptar"
    AXCodCta.NroCuenta = psCodCta
    AxCodCta_keypressEnter
 End If
 Me.Show 1
End Function
Private Sub AxCodCta_keypressEnter()
Dim fcConexionJud As ADODB.Connection
If Len(Trim(AXCodCta.NroCuenta)) = 18 Then
    fsCodCta = Trim(AXCodCta.NroCuenta)
    'Mostrar Datos del Credito
    Call MuestraDatos(fsCodCta, fcConexionJud, fbNuevaNegoc)
 End If
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
Dim fcConexionJud As ADODB.Connection
If Len(Trim(AXCodCta.NroCuenta)) = 18 Then
    fsCodCta = Trim(AXCodCta.NroCuenta)
    'Mostrar Datos del Credito
    Call MuestraDatos(fsCodCta, fcConexionJud, fbNuevaNegoc)
 End If
End Sub

Private Sub cmdAplicar_Click()
Dim lnIntCalculado As Double
Dim lnComisAbogCalculado As Double
Dim lnNroCuotas As Integer, i As Integer
Dim lnPeriodo As Integer
Dim lnMontoCuota As Double
Dim L As ListItem
Dim ldFecPago As Date
Dim lnComiAboxCuota As Double
If ValidaDatosNeg = False Then Exit Sub

'lnNroCuotas = Int(CDbl(Me.txtTotal) / CDbl(Me.txtNegMonto))
lnNroCuotas = Int(Me.txtNegCuotas.Text)
If optFecha.value = True Then
    lnPeriodo = 30
ElseIf optPeriodo.value = True Then
    lnPeriodo = Val(Me.txtPeriodoFijo.Text)
Else
    MsgBox "Debe ingresar Calendario ", vbInformation, "Aviso"
    Exit Sub
End If
'Me.lblIntCalculado.Caption = Format(lnIntCalculado, "#,##0.00")
'Me.lblComisionCalculada.Caption = Format(lnComisAbogCalculado, "#,##0.00")
'Me.lblTotal.Caption = Format((CDbl(Me.lblDeudaFecha) + lnIntCalculado + lnComisAbogCalculado), "#,##0.00")

'Calculo de Comision de Abogado
lnComiAboxCuota = CalculoComisionAbogado()
'Calcula el Monto de la Cuota
lnMontoCuota = Format((CDbl(Me.txtNegMonto) / Val(Me.txtNegCuotas)) + lnComiAboxCuota, "#,##0.00")

'LLena el Calendario
lvwCalendario.ListItems.Clear
ldFecPago = Me.TxtNegVigencia.Text
For i = 1 To lnNroCuotas
    Set L = lvwCalendario.ListItems.Add(, , Trim(Str(i)))
    'Fecha de Proxima Cuota
    If optPeriodo.value = True Then
        ldFecPago = ldFecPago + lnPeriodo
    ElseIf optFecha.value = True Then
        ldFecPago = fgProximaFechaFija(ldFecPago, Val(txtFechaFija.Text))
    End If
    L.SubItems(1) = Format(ldFecPago, "dd/mm/yyyy")
    L.SubItems(2) = Format(lnMontoCuota, "#,##0.00")
Next i


Me.txtNegMonto.Enabled = False
Me.txtNegCuotas.Enabled = False
cmdAplicar.Enabled = False
cmdDehacer.Enabled = True
fraTipCalendario.Enabled = False

End Sub

Private Sub cmdBuscar_Click()

Dim RegPerCta As New ADODB.Recordset
Dim sSQL As String
'Dim loConec As DConecta
Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersCredito  As COMDColocRec.DCOMColRecCredito
Dim lrCreditos As ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona

On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If Not loPers Is Nothing Then
        lsPersCod = loPers.sPersCod
        lsPersNombre = loPers.sPersNombre
    Else
        Exit Sub
    End If
Set loPers = Nothing

' Selecciona Estados
lsEstados = gColocEstRecVigJud & "," & gColocEstRecVigCast

If Trim(lsPersCod) <> "" Then
    Set loPersCredito = New COMDColocRec.DCOMColRecCredito
        Set lrCreditos = loPersCredito.dObtieneCreditosDePersona(lsPersCod, lsEstados)
    Set loPersCredito = Nothing
End If

Set loCuentas = New COMDPersona.UCOMProdPersona
    Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrCreditos)
    If loCuentas.sCtaCod <> "" Then
        AXCodCta.Enabled = True
        AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        AXCodCta.SetFocusCuenta
    End If
Set loCuentas = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Private Sub cmdCalAceptar_Click()
Dim L As ListItem
 
If Not IsDate(txtCalFecha) Then
    MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
    If txtCalFecha.Enabled Then txtCalFecha.SetFocus
    Exit Sub
ElseIf CDate(txtCalFecha) < gdFecSis Then
    MsgBox "Debe ingresar una fecha posterior a la fecha actual.", vbInformation, "Aviso"
    If txtCalFecha.Enabled Then txtCalFecha.SetFocus
    Exit Sub
ElseIf Not IsNumeric(txtCalMonto) Then
    MsgBox "Debe ingresar un monto valido.", vbInformation, "Aviso"
    txtCalMonto.SetFocus
    Exit Sub
ElseIf Not IsNumeric(txtCalMonto) Then
    MsgBox "Debe ingresar un monto mayor a 0.", vbInformation, "Aviso"
    txtCalMonto.SetFocus
    Exit Sub
End If
 
If Me.lvwCalendario.ListItems.Count > 0 Then
    If CDate(Me.lvwCalendario.ListItems(Me.lvwCalendario.ListItems.Count).ListSubItems(1)) >= CDate(txtCalFecha) Then
        MsgBox "Debe ingresar una fecha posterior a Ultima Cuota ingresada.", vbInformation, "Aviso"
        If txtCalFecha.Enabled Then txtCalFecha.SetFocus
        Exit Sub
    End If
    
End If

If Not VerificaCalendarioCLibre(CCur(txtCalMonto.Text), True) Then
    MsgBox "Los montos ingresados suman mas del monto indicado en el calendario.", vbInformation, "Aviso"
    txtCalMonto.SetFocus
    Exit Sub
End If
 

Set L = lvwCalendario.ListItems.Add(, , Trim(lvwCalendario.ListItems.Count) + 1)
    'Fecha de Proxima Cuota
    L.SubItems(1) = Format(Me.txtCalFecha, "dd/mm/yyyy")
    L.SubItems(2) = Format(Me.txtCalMonto, "#,##0.00")
    
Me.txtCalFecha = "__/__/____"
Me.txtCalMonto = "0"
End Sub

Private Sub cmdCalAgregar_Click()
    If Val(txtNegMonto.Text) = 0 Then
        MsgBox "Ingrese Monto del Prestamo", vbInformation, "Aviso"
        If txtNegMonto.Enabled Then
            txtNegMonto.SetFocus
        End If
        Exit Sub
    End If
    txtCalFecha.Enabled = True
    txtCalMonto.Enabled = True

    txtCalFecha.SetFocus
    
End Sub

Private Sub CmdCalcular_Click()
Dim nMontoNegociar As Double

If Not IsNumeric(txtNegCuotas) Then
    MsgBox "Ingrese un dato Numerico", vbInformation, "AVISO"
    txtNegCuotas.SetFocus
    Exit Sub
End If

If txtNegCuotas <= 0 Then
    MsgBox "Ingrese el nro de Cuotas (meses) mayor a 0", vbInformation, "AVISO"
    txtNegCuotas.SetFocus
    Exit Sub
End If
nMontoNegociar = Me.lblCapital + (Me.lblCapital * (Me.lbltasainteres / 100) * Me.txtNegCuotas) + Me.LblIntComp + Me.LblIntMor + Me.LblGastos
Me.txtNegMonto = Format(nMontoNegociar, "#0.00")

End Sub

Private Sub cmdCalEliminar_Click()
    If lvwCalendario.ListItems.Count = 0 Then Exit Sub
    If MsgBox("Se va Ha Eliminar la Cuota " & lvwCalendario.ListItems.Count & ", Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
       lvwCalendario.ListItems.Remove lvwCalendario.ListItems.Count
       fnCuotaSeleccionada = lvwCalendario.ListItems.Count
    End If

End Sub

Private Sub cmdCancelar_Click()
If fntipo = 1 Then ' Simulador
    FraCredito.Enabled = True
    fraNegociacion1.Enabled = False
    'fraCalendario.Enabled = True
    Call LimpiaDatos
ElseIf fntipo = 2 Then  ' Registro
    Unload Me
End If
End Sub

Private Sub cmdDehacer_Click()
fraNegociacion1.Enabled = True
cmdDehacer.Enabled = False
cmdAplicar.Enabled = True
Me.txtNegMonto.Enabled = True
Me.txtNegCuotas.Enabled = True
lvwCalendario.ListItems.Clear
fraTipCalendario.Enabled = True
End Sub

Private Sub cmdImprimir_Click()
Dim lsCadena As String
Dim rs As New ADODB.Recordset
Dim lcRecImp As COMNColocRec.NCOMColRecImpre
Dim i As Integer

If Me.lvwCalendario.ListItems.Count < 1 Then
    MsgBox "No se ha generado el Calendario de Pagos", vbInformation, "Aviso"
    Exit Sub
End If
With rs
    'Crear RecordSet
    .Fields.Append "nNro", adInteger
    .Fields.Append "dFecha", adDate
    .Fields.Append "nMonto", adCurrency
    .Open
    'Llenar Recordset
    For i = 1 To Me.lvwCalendario.ListItems.Count
        .AddNew
        .Fields("nNro") = lvwCalendario.ListItems(i)
        .Fields("dFecha") = lvwCalendario.ListItems(i).SubItems(1)
        .Fields("nMonto") = lvwCalendario.ListItems(i).SubItems(2)
    Next i
End With

Set lcRecImp = New COMNColocRec.NCOMColRecImpre
    lsCadena = lcRecImp.ImprimeNegociacion(gsNomCmac, gdFecSis, gsNomAge, gsCodUser, Me.AXCodCta.NroCuenta, Me.lblNombreCliente.Caption, Me.lblEstudioJuridico.Caption, Me.txtNegCuotas.Text, rs, fntipo, gImpresora)
    
Set lcRecImp = Nothing

Dim loPrevio As previo.clsPrevio
    If Len(Trim(lsCadena)) > 0 Then
        Set loPrevio = New previo.clsPrevio
        loPrevio.Show lsCadena, "Recuperaciones - Negociaciones ", True
        Set loPrevio = Nothing
    Else
        MsgBox "No Existen Datos para el reporte", vbInformation, "Aviso"
    End If
'rs.Close
End Sub

Private Sub cmdNuevo_Click()
    'fraCredito.Enabled = True
    'fraNegociacion.Enabled = True
    'fraCalendario.Enabled = True
    'fbNuevaNegoc = True
    'cmdBuscar_Click
End Sub

Private Sub CmdSalir_Click()
Dim LCal As ListItem
Dim i As Integer
'Verifica que el total de calendario se haya generado
' para cuota libre
If Me.optCuotaLibre.value = True Then
    If Not VerificaCalendarioCLibre Then
        MsgBox "No se genero el calendario Cuota Libre Correctamente", vbInformation, "Aviso"
        Exit Sub
    End If
    'Actualiza el nro de cuotas
    txtNegCuotas.Text = Me.lvwCalendario.ListItems(Me.lvwCalendario.ListItems.Count).Text
End If
If fntipo = 1 Then ' Simulador
    Unload Me
ElseIf fntipo = 2 Then
    'Datos de Negociacion
    frmColRecNegRegistro.txtNegMonto = Me.txtNegMonto
    frmColRecNegRegistro.txtNegCuotas = Me.txtNegCuotas
    frmColRecNegRegistro.TxtNegVigencia = Me.TxtNegVigencia
    'Asigna el Calendario
    For i = 1 To lvwCalendario.ListItems.Count
        Set LCal = frmColRecNegRegistro.lvwCalendario.ListItems.Add(, , Trim(Str(i)))
        LCal.SubItems(1) = lvwCalendario.ListItems(i).SubItems(1)
        LCal.SubItems(2) = lvwCalendario.ListItems(i).SubItems(2)
        LCal.SubItems(3) = 0
        LCal.SubItems(4) = "P"
    Next i
    Unload Me
End If

End Sub


Private Sub MuestraDatos(ByVal psCodCta As String, ByVal pConex As ADODB.Connection, ByVal pbNuevaNeg As Boolean)

'**** Datos del Credito
Call MuestraDatosCredito(psCodCta)
Me.TxtNegVigencia.Enabled = True
Me.TxtNegVigencia.Text = Format(gdFecSis, "dd/mm/yyyy")
fraNegociacion1.Enabled = True

If pbNuevaNeg = False Then  ' Muestra datos de Negociacion Actual
    'Call MuestraDatosNegocia(psCodCta, pConex)

Else ' Negociacion Nueva
    'Verifica si tiene Negociacion Activa
'    Dim lrNegAct As New ADODB.Recordset
'    Dim lsSQL As String
'    lsSQL = " Select * from ColocRecupNegocia Where cCodCta ='" & psCodCta & "' And cEstado ='V' "
'    lrNegAct.Open lsSQL, pConex, adOpenStatic, adLockReadOnly, adCmdText
'    If Not (lrNegAct.BOF And lrNegAct.EOF) Then
'        fsNegAnterior = lrNegAct!cNroNeg
'        If MsgBox("El cliente Tiene una negociacion activa, Desea Anularla ? ", vbInformation + vbYesNo, "Aviso") = vbNo Then
'            Call MuestraDatosNegocia(psCodCta, pConex)
'            'Call HabilitaControles(True, True, False, True, True, True, False, False, True)
'            Exit Sub
'        End If
'    Else
'        fsNegAnterior = ""
'    End If
'    Set lrNegAct = Nothing
'
    Me.txtNegMonto.Enabled = True
    Me.txtNegMonto.SetFocus
    'cmdCalendario.SetFocus
End If

End Sub


Private Sub MuestraDatosCredito(ByVal psCtaCod As String)
'On Error GoTo ControlError
Dim reg As New ADODB.Recordset
Dim lcRec As COMDColocRec.DCOMColRecNegociacion
Dim lsSQL As String
Dim lsComi As String
Dim lnIntComGenerado  As Double
Dim lnDiaUltPago As Integer
Dim loCalcula As COMNColocRec.NCOMColRecCalculos
Dim fnIntCompGenerado As Currency, fnIntMoraGenerado As Currency
Dim fsFecUltPago As String
Dim lnDiasUltTrans As Integer
Dim fnTasaInt As Double
Dim fnTasaIntMorat As Double
Dim fnSaldoCap As Double
    CargaParametros
    ' Busca el Credito
    Set lcRec = New COMDColocRec.DCOMColRecNegociacion
        Set reg = lcRec.ObtenerDatosCredparaNegociacion(psCtaCod)
    Set lcRec = Nothing
    
    If reg.BOF And reg.EOF Then
        reg.Close
        Set reg = Nothing
        MsgBox " No se encuentra el Credito " & fsCodCta, vbInformation, " Aviso "
        LimpiaDatos
        AXCodCta.Enabled = True
        Exit Sub
    Else
        ' Mostrar los datos del Credito
        Me.lblNombreCliente = PstaNombre(reg!NomCliente, False)
        Me.lblEstudioJuridico = PstaNombre(reg!NomEstJur, False)
        Me.lblEstadoCredito = IIf((reg!nPrdEstado = gColocEstRecVigJud Or reg!nPrdEstado = gColocEstRecCanJud), "Vigente", "Cancelado")
        Me.lblCondicionCredito = IIf(reg!nPrdEstado = gColocEstRecVigJud, "Judicial", "Castigado")
        Me.lbltasainteres = Format(reg!nTasaInt, "#0.00")
        
        fnTasaInt = IIf(IsNull(reg!nTasaInt), 0, reg!nTasaInt)
        fnTasaIntMorat = reg!nTasaIntMor
        fsFecUltPago = CDate(fgFechaHoraGrab(reg!cUltimaActualizacion))
        lnDiasUltTrans = CDate(Format(gdFecSis, "dd/mm/yyyy")) - CDate(Format(fsFecUltPago, "dd/mm/yyyy"))
        fnSaldoCap = reg!nSaldo
        'Interes Generado
        
        'Calcula el Int Comp Generado
        Set loCalcula = New COMNColocRec.NCOMColRecCalculos
            If fnTipoCalcIntComp = 0 Then ' NoCalcula
                fnIntCompGenerado = 0
            ElseIf fnTipoCalcIntComp = 1 Then ' En base al capital
                If fnFormaCalcIntComp = 1 Then 'INTERES COMPUESTO
                    fnIntCompGenerado = loCalcula.nCalculaIntCompGenerado(lnDiasUltTrans, fnTasaInt, fnSaldoCap)
                Else
                    'INTERES SIMPLE
                    fnIntCompGenerado = loCalcula.nCalculaIntCompGeneradoICA(lnDiasUltTrans, fnTasaInt, fnSaldoCap)
                End If
            End If
            
            If fnTipoCalcIntMora = 0 Then  ' NoCalcula
                fnIntMoraGenerado = 0
            ElseIf fnTipoCalcIntMora = 1 Then ' En base al capital
                If fnFormaCalcIntMora = 1 Then 'INTERES COMPUESTO
                    fnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGenerado(lnDiasUltTrans, fnTasaIntMorat, fnSaldoCap)
                Else
                    'INTERES SIMPLE
                    fnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGeneradoICA(lnDiasUltTrans, fnTasaIntMorat, fnSaldoCap)
                End If
            End If
              
        Set loCalcula = Nothing
       
        
        
        Me.lblCapital = Format(reg!nSaldo, "#,##0.00")
        Me.LblIntComp = Format(reg!nSaldoIntComp + fnIntCompGenerado, "#,##0.00")
        Me.LblIntMor = Format(reg!nSaldoIntMor + fnIntMoraGenerado, "#,##0.00")
        Me.LblGastos = Format(reg!nSaldoGasto, "#,##0.00")
        
        ' Calcula Int Comp Generado
        lnDiaUltPago = DateDiff("d", Format(reg!dIngRecup, "dd/mm/yyyy"), Format(gdFecSis, "dd/mm/yyyy"))
        If lnDiaUltPago > 0 Then
        'Aqui me quede Cambiar -- nCierreMesRecuperaciones
        
        '    lnIntComGenerado = CalculaIntComJudi(lnDiaUltPago, reg!nTasaInt, reg!nSaldo)
        'Else
        '    lnIntComGenerado = 0
        End If
        'Me.LblIntComp = Format(reg!nSaldoIntComp + lnIntComGenerado, "#,##0.00")
        Me.LblDeudaFecha = Format(reg!nSaldo + reg!nSaldoIntComp + reg!nSaldoIntMor + reg!nSaldoGasto + fnIntCompGenerado + fnIntMoraGenerado, "#,##0.00")
        
        lblComision.Caption = reg!nTipComis & " - " & reg!nComisionValor
        fnPorComisionAbo = reg!nComisionValor
        AXCodCta.Enabled = False
         
    End If
    Set reg = Nothing
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Private Sub MuestraDatosNegocia(ByVal psCodCta As String)
On Error GoTo ControlError
Dim lcRec As COMDColocRec.DCOMColRecNegociacion
Dim reg As New ADODB.Recordset
Dim lsSQL As String
Dim L As ListItem

    ' Busca la Negociacion
    Set lcRec = New COMDColocRec.DCOMColRecNegociacion
         Set reg = lcRec.ObtenerDatosCredNegociacion(psCodCta)
    Set lcRec = Nothing
    
   
    If reg.BOF And reg.EOF Then
        MsgBox " Credito No Tiene Negociaciones Vigentes para Credito  " & psCodCta, vbInformation, " Aviso "
        LimpiaDatos
        AXCodCta.Enabled = True
        Exit Sub
    Else
        ' Mostrar los datos de Negociacion
        'Me.txtNegNro.Text = reg!cNroNeg
        Me.TxtNegVigencia.Text = Format(reg!dFecVig, "dd/mm/yyyy")
        'Me.txtNegEstado.Text = IIf(reg!cEstado = "V", "Vigente", "Cancelado")
        Me.txtNegMonto.Text = Format(reg!nMontoNeg, "#0.00")
        Me.txtNegCuotas.Text = Format(reg!nCuotasNeg, "#0.00")
        'Me.txtNegComenta.Text = IIf(IsNull(reg!cComenta), "", reg!cComenta)
        reg.Close
        Set reg = Nothing
    End If
    ' Busca Plan de Pagos de Negociacion
    'lsSQL = "SELECT * FROM NegocPlanPagos WHERE cCtaCod = '" & psCodCta & "' " & _
            "And cNroNeg ='" & Trim(txtNegNro.Text) & "'  ORDER BY nNroCuota "
    
    'reg.CursorLocation = adUseClient
    'reg.Open lsSQL, pConex, adOpenStatic, adLockReadOnly, adCmdText
    'Set reg.ActiveConnection = Nothing
    
    'If reg.BOF And reg.EOF Then
    '    MsgBox " Negociacion No Posee Plan Pagos " & psCodCta, vbInformation, " Aviso "
    'Else
        ' Mostrar Plan de Pagos
    '    Do While Not reg.EOF
    '        Set L = lvwCalendario.ListItems.Add(, , Format(reg!dFecVenc))
    '        L.SubItems(1) = Trim(Str(reg!nNroCuota))
    '        L.SubItems(2) = Format(reg!nMonto, "#0.00")
    '        L.SubItems(3) = Format(reg!nMontoPag, "#0.00")
    '        reg.MoveNext
    '    Loop
    'End If
  
  '27-12
  'Set loConec = Nothing
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "


End Sub

Private Sub ActivaConexionJud()
'Dim lbCone As Boolean
'Dim lsAgeJud As String
'AbreConexion
'lsAgeJud = AgenciaJudicial
'If gsCodAge = lsAgeJud Then
'    Set pConexionJud = dbCmact
'    lbConexion = False
'Else
'    lbCone = AbreConeccion(Mid(lsAgeJud, 4), True, False)
'    Set pConexionJud = dbCmactN
'    If lbCone = False Then
'      MsgBox "No se Pudo Conectar con la Base de Judicial"
'      Unload Me
'      Exit Sub
'    End If
'    lbConexion = True
'End If
'
End Sub

'*  VALIDACION DE DATOS DEL FORMULARIO ANTES DE GRABAR
Function ValidaDatos() As Boolean

Dim reg As New ADODB.Recordset
Dim lsSQL As String
Dim MonGarant As Currency
 
'ValidaDatos = True
'    'valida fecha Vigencia
'    If ValidaFecha(Me.TxtFechaVigencia) <> "" Then
'        MsgBox Cad, vbInformation, "Aviso"
'        ValidaDatos = False
'        TxtFechaVigencia.SetFocus
'        Exit Function
'    End If
'    'verificando que Calendario sea igual a Monto de Negociacion
'
'
   
End Function

'****************************************************************
'*  LIMPIA LOS DATOS DE LA PANTALLA PARA UNA NUEVA APROBACION
'****************************************************************
Sub LimpiaDatos()
    AXCodCta.NroCuenta = ""
    lblNombreCliente.Caption = ""
    lblEstudioJuridico.Caption = ""
    lblCondicionCredito.Caption = ""
    lblComision.Caption = ""
    lbltasainteres.Caption = ""
    TxtNegVigencia.Text = "__/__/____"
    txtNegMonto.Text = ""
    txtNegCuotas.Text = ""
    lblCapital.Caption = ""
    LblIntComp.Caption = ""
    LblIntMor.Caption = ""
    LblGastos.Caption = ""
    LblDeudaFecha.Caption = ""
    lblTotal.Caption = ""
    lblEstadoCredito.Caption = ""
    lblIntCalculado.Caption = ""
    lblComisionCalculada.Caption = ""
    lvwCalendario.ListItems.Clear
    AXCodCta.NroCuenta = fgIniciaAxCuentaRecuperaciones
    AXCodCta.Enabled = True
    
End Sub



Private Sub Form_Load()
optPeriodo_Click
End Sub

Private Sub lvwCalendario_Click()
    If lvwCalendario.SelectedItem Is Nothing Then Exit Sub
    fnCuotaSeleccionada = lvwCalendario.SelectedItem + 1
End Sub

Private Sub optCuotaLibre_Click()
If optCuotaLibre.value = True Then
    lvwCalendario.ListItems.Clear
    fraCalendarioLibre.Visible = True
    
    txtPeriodoFijo.Text = ""
    txtFechaFija.Text = ""
    cmdCalAgregar.SetFocus
End If
End Sub

Private Sub optFecha_Click()
If optFecha.value = True Then
    txtPeriodoFijo.Text = ""
    txtFechaFija.Text = ""
    txtFechaFija.SetFocus
    fraCalendarioLibre.Visible = False
End If
End Sub

Private Sub optPeriodo_Click()
If optPeriodo.value = True Then
    txtFechaFija.Text = ""
    txtPeriodoFijo.Text = "30"
    'txtPeriodoFijo.SetFocus
    fraCalendarioLibre.Visible = False
End If
End Sub

Private Sub txtCalFecha_GotFocus()
    txtCalFecha.SelStart = 0
    txtCalFecha.SelLength = 50
End Sub

Private Sub txtCalFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCalMonto.SetFocus
End If
End Sub

Private Sub txtCalMonto_GotFocus()
    txtCalMonto.SelStart = 0
    txtCalMonto.SelLength = 50
End Sub

Private Sub txtCalMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdCalAceptar.SetFocus
Else
    KeyAscii = NumerosDecimales(txtCalMonto, KeyAscii)
End If
End Sub


Private Sub txtFechaFija_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(txtFechaFija)) > 0 Then
            If txtFechaFija < 0 Then
               MsgBox "El Numero debe ser Mayor a 0", vbInformation, "Aviso"
               txtFechaFija = ""
               txtFechaFija.SetFocus
               Exit Sub
            End If
        End If
    End If
End Sub

Private Sub txtNegCuotas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(txtNegCuotas)) > 0 Then
            If txtNegCuotas < 0 Then
               MsgBox "El Nro de Cuotas debe ser Mayor a 0", vbInformation, "Aviso"
               txtNegCuotas = ""
               txtNegCuotas.SetFocus
               Exit Sub
            End If
        End If
        cmdAplicar.SetFocus
    End If
End Sub

Private Sub txtNegMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(Trim(txtNegMonto)) > 0 Then
        If txtNegMonto < 0 Then
           MsgBox "El Monto debe ser Mayor a 0", vbInformation, "Aviso"
           txtNegMonto = ""
           txtNegMonto.SetFocus
           Exit Sub
        End If
    End If
    txtNegCuotas.Enabled = True
    txtNegCuotas.SetFocus
End If
End Sub


Private Function ValidaDatosNeg() As Boolean
ValidaDatosNeg = True
If Len(Me.txtNegMonto) = 0 Then
    MsgBox "Ingrese Monto de la cuota a Pagar", vbInformation, "Aviso"
    Me.txtNegMonto.SetFocus
    ValidaDatosNeg = False
    Exit Function
End If
If Len(Me.txtNegCuotas) = 0 Then
    MsgBox "Ingrese Nro de Cuotas", vbInformation, "Aviso"
    Me.txtNegCuotas.SetFocus
    ValidaDatosNeg = False
    Exit Function
End If
If optPeriodo.value = True And Len(Me.txtPeriodoFijo.Text) = 0 Then
    MsgBox "Ingrese el Periodo", vbInformation, "Aviso"
    Me.txtPeriodoFijo.SetFocus
    ValidaDatosNeg = False
    Exit Function
End If
If optFecha.value = True And Len(Trim(Me.txtFechaFija.Text)) = 0 Then
    MsgBox "Ingrese el dia ", vbInformation, "Aviso"
    Me.txtFechaFija.SetFocus
    ValidaDatosNeg = False
    Exit Function
End If
End Function

'Calcula la proxima fecha de pago, de un calendario con fecha fija
Public Function fgProximaFechaFija(ByVal pdFecha As Date, ByVal pnDia As Integer) As Date
Dim Mon As Integer
Dim Yea As Integer
Dim ldNewFecha As Date
    Mon = Month(pdFecha)
    Yea = Year(pdFecha)
    
    
    If Mon >= 12 Then
        Mon = 1
        Yea = Yea + 1
    Else
        Mon = Mon + 1
    End If
    'verificar que en febrero no pase de el dia 28 a una fecha fija
    If Mon = 2 And pnDia > 28 Then
        If Yea Mod 4 <> 0 Then
            'obtener nueva fecha para año no bisiesto
            ldNewFecha = CDate("28/" + Str(Mon) + "/" + Str(Yea))
        Else
            'obtener nueva fecha para año bisiesto 2000,2004,2008,....
            ldNewFecha = CDate("29/" + Str(Mon) + "/" + Str(Yea))
        End If
    Else
        If (Mon = 4 Or Mon = 6 Or Mon = 9 Or Mon = 11) And (pnDia > 30) Then
            'obtener nueva fecha valida para el mes
            ldNewFecha = CDate("30/" + Str(Mon) + "/" + Str(Yea))
        Else
            'obtener nueva fecha
            ldNewFecha = CDate(Str(pnDia) + "/" + Str(Mon) + "/" + Str(Yea))
        End If
    End If
fgProximaFechaFija = ldNewFecha

End Function

Public Function VerificaCalendarioCLibre(Optional pnMonto As Double = 0, Optional pbEsMayor As Boolean = False) As Boolean
Dim i As Integer
Dim lnTotalCalendario As Double
VerificaCalendarioCLibre = True
    For i = 1 To lvwCalendario.ListItems.Count
        lnTotalCalendario = lnTotalCalendario + lvwCalendario.ListItems(i).SubItems(2)
    Next i
If pbEsMayor Then
    If lnTotalCalendario + pnMonto > CDbl(Me.txtNegMonto.Text) Then
        VerificaCalendarioCLibre = False
    Else
        VerificaCalendarioCLibre = True
    End If
Else
    If lnTotalCalendario <> Val(Me.txtNegMonto) Then
        VerificaCalendarioCLibre = False
    Else
        VerificaCalendarioCLibre = True
    End If
End If
End Function

Private Sub txtPeriodoFijo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(txtPeriodoFijo)) > 0 Then
            If txtPeriodoFijo < 0 Then
               txtPeriodoFijo = ""
               MsgBox "El Periodo debe ser Mayor a 0", vbInformation, "Aviso"
               txtPeriodoFijo.SetFocus
               Exit Sub
            End If
        End If
    End If
End Sub

Private Function CalculoComisionAbogado() As Double
  Dim nMontoCIM As Double
  Dim nComisionAbog As Double
  Dim loCalculaComision As COMNColocRec.NCOMColRecCalculos
  
  Set loCalculaComision = New COMNColocRec.NCOMColRecCalculos
  nMontoCIM = CDbl(lblCapital) + CDbl(LblIntComp) + CDbl(LblIntMor)
  If txtNegMonto.Text >= nMontoCIM Then
     nComisionAbog = Round(loCalculaComision.nCalculaComisionAbogado(fnPorComisionAbo, nMontoCIM), 2)
     CalculoComisionAbogado = nComisionAbog / CDbl(txtNegCuotas.Text)
  Else
     nComisionAbog = Round(loCalculaComision.nCalculaComisionAbogado(fnPorComisionAbo, CDbl(txtNegMonto.Text)), 2)
     CalculoComisionAbogado = nComisionAbog / CDbl(txtNegCuotas.Text)
  End If
  
  Set loCalculaComision = Nothing
End Function

Private Sub CargaParametros()
Dim loParam As COMDConstSistema.NCOMConstSistema
Set loParam = New COMDConstSistema.NCOMConstSistema
    fnTipoCalcIntComp = loParam.LeeConstSistema(151)
    fnTipoCalcIntMora = loParam.LeeConstSistema(152)
    fnFormaCalcIntComp = loParam.LeeConstSistema(202) ' CMACICA
    fnFormaCalcIntMora = loParam.LeeConstSistema(203) ' CMACICA
    
  
    
Set loParam = Nothing
End Sub

