VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRepSBS13 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generación de Reporte SBS Nº 13"
   ClientHeight    =   10215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10650
   Icon            =   "frmRepSBS13.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10215
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Sicmact.FlexEdit feRepDet 
      Height          =   6735
      Left            =   120
      TabIndex        =   37
      Top             =   3360
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   11880
      Cols0           =   5
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Aspecto-Val 1-Val 2-R(V1/V2) %"
      EncabezadosAnchos=   "0-5200-1400-1400-1200"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-2-3-X"
      ListaControles  =   "0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-R-R-R"
      FormatosEdit    =   "0-0-2-2-2"
      CantEntero      =   18
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin MSComctlLib.ProgressBar prgProgress 
      Height          =   150
      Left            =   3720
      TabIndex        =   36
      Top             =   2970
      Visible         =   0   'False
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Max             =   21
      Scrolling       =   1
   End
   Begin VB.Frame fraBackground 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   3720
      TabIndex        =   34
      Top             =   2970
      Visible         =   0   'False
      Width           =   5580
   End
   Begin VB.CommandButton cmdRecalcular 
      Caption         =   "Recalcular"
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
      Left            =   9480
      TabIndex        =   27
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdRepSBS13 
      Caption         =   "Reporte SBS 13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   30
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdHojaTrabajo 
      Caption         =   "Hoja de Trabajo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   29
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
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
      Left            =   9480
      TabIndex        =   28
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
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
      Left            =   9480
      TabIndex        =   31
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Frame fraPatriEfectivo2 
      Caption         =   "Patrimonio Efectivo"
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
      Height          =   735
      Left            =   120
      TabIndex        =   24
      Top             =   2520
      Width           =   3495
      Begin VB.TextBox txtPEMesAnterior 
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
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   1440
         TabIndex        =   26
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "P.E. Mes Ant."
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
         Left            =   120
         TabIndex        =   25
         Top             =   300
         Width           =   1185
      End
   End
   Begin VB.Frame fraPatriEfectivo 
      Caption         =   "Patrimonio Efectivo"
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
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   10455
      Begin VB.TextBox txtPEPatrimonioContable 
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
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   8880
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtPEMesActual 
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
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   6000
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtPENivel2 
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
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   3600
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtPENivel1 
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
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   1080
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Patrimonio Contable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7680
         TabIndex        =   22
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label8 
         Caption         =   "P.E. Mes Actual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5040
         TabIndex        =   20
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nivel 2"
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
         Left            =   2640
         TabIndex        =   18
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nivel 1"
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
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame fraDeudaSubordinada 
      Caption         =   "Deuda Subordinada"
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
      Height          =   735
      Left            =   8160
      TabIndex        =   13
      Top             =   840
      Width           =   2415
      Begin VB.TextBox txtDeudaSubordinada 
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
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame fraReqPatrimonio 
      Caption         =   "Requerimiento de Patrimonio Efectivo por Tipo de Riesgo"
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
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   7935
      Begin VB.TextBox txtROcupacional 
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
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   6360
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtRMercado 
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
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   3600
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtRCredito 
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
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   1080
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "R. Operacional"
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
         Left            =   5040
         TabIndex        =   11
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "R. Mercado"
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
         Left            =   2520
         TabIndex        =   9
         Top             =   300
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "R. Crédito"
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
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   870
      End
   End
   Begin VB.Frame fraPeriodo 
      Caption         =   "Periodo"
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
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5730
      Begin VB.CommandButton cmdSeleccionar 
         Caption         =   "Seleccionar"
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
         Left            =   4200
         TabIndex        =   5
         Top             =   200
         Width           =   1335
      End
      Begin VB.TextBox txtAnio 
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
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox CboMes 
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
         ItemData        =   "frmRepSBS13.frx":030A
         Left            =   690
         List            =   "frmRepSBS13.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1830
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
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
         Left            =   2640
         TabIndex        =   3
         Top             =   280
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
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
         Left            =   165
         TabIndex        =   1
         Top             =   280
         Width           =   480
      End
   End
   Begin VB.Label lblPorcentajeProgreso 
      Caption         =   "100%"
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
      Height          =   165
      Left            =   9360
      TabIndex        =   33
      Top             =   2910
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblCalculoDesc 
      Caption         =   "CalculoDesc"
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
      Height          =   200
      Left            =   3720
      TabIndex        =   35
      Top             =   3090
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Label lblCalculando 
      Caption         =   "Calculando..."
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
      Height          =   165
      Left            =   3720
      TabIndex        =   32
      Top             =   2750
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmRepSBS13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'Nombre : frmRepSBS13
'Descripcion:Formulario para Generar el Reporte SBS 13
'Creacion: PASI TI-ERS024-2015
'********************************************************************
Option Explicit
Dim oBalan As DbalanceCont
Private Sub cboMes_Click()
     txtAnio.SetFocus
End Sub
Private Sub cmdCerrar_Click()
    Unload Me
End Sub
Private Sub cmdGuardar_Click()
    If Not ValidaInterfaz Then Exit Sub
    If Not ValidaDatos(True) Then Exit Sub
    If feRepDet.TextMatrix(feRepDet.row, 1) = "" Then
        MsgBox "Se necesita realizar el calculo de los ratios para poder guardar la información.", vbInformation
        cmdRecalcular.SetFocus
        Exit Sub
    End If
    If oBalan.ExisteConfigRepSBS13(Trim(txtAnio.Text), Format(CboMes.ListIndex + 1, "00")) Then
        If MsgBox("Ya existe informacion para el periodo. Esta Seguro de Actualizar la información.?", vbInformation + vbYesNo, "Aviso") = vbNo Then
            Exit Sub
        End If
        ActualizaDatos
    Else
        If MsgBox("Esta Seguro de Guardar los Datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then
            Exit Sub
        End If
        GuardaDatos
    End If
End Sub
Private Sub GuardaDatos()
    Dim index As Integer
    Dim nIdRep As Integer
    Dim oDMov As DMov
    Set oDMov = New DMov
    Dim lsMovNro As String
    Dim nValor As Double
    On Error GoTo ErrorGuarda
    cmdGuardar.Enabled = False
    lsMovNro = oDMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    nIdRep = 0
    nIdRep = oBalan.RegistrarConfigRepSBS13(Format(CboMes.ListIndex + 1, "00"), Trim(txtAnio.Text), lsMovNro, 1)
    oBalan.RegistrarConfigRepSBS13Det nIdRep, "R. Crédito", CStr(txtRCredito.Text), 0, 0
    oBalan.RegistrarConfigRepSBS13Det nIdRep, "R. Mercado", CStr(txtRMercado.Text), 0, 0
    oBalan.RegistrarConfigRepSBS13Det nIdRep, "R. Operacional", CStr(txtROcupacional.Text), 0, 0
    oBalan.RegistrarConfigRepSBS13Det nIdRep, "Deuda Subordinada", CStr(txtDeudaSubordinada.Text), 0, 0
    oBalan.RegistrarConfigRepSBS13Det nIdRep, "PE Nivel 1", CStr(txtPENivel1.Text), 0, 0
    oBalan.RegistrarConfigRepSBS13Det nIdRep, "PE Nivel 2", CStr(txtPENivel2.Text), 0, 0
    oBalan.RegistrarConfigRepSBS13Det nIdRep, "PE Mes Actual", CStr(txtPEMesActual.Text), 0, 0
    oBalan.RegistrarConfigRepSBS13Det nIdRep, "Patrimonio Contable", CStr(txtPEPatrimonioContable.Text), 0, 0
    oBalan.RegistrarConfigRepSBS13Det nIdRep, "PE Mes Ant", CStr(txtPEMesAnterior.Text), 0, 0
    For index = 1 To feRepDet.Rows - 1
        nValor = IIf(feRepDet.TextMatrix(index, 4) = "N/A", -1, (Replace(feRepDet.TextMatrix(index, 4), "%", "")))
        oBalan.RegistrarConfigRepSBS13Det nIdRep, feRepDet.TextMatrix(index, 1), feRepDet.TextMatrix(index, 2), feRepDet.TextMatrix(index, 3), nValor
    Next
    MsgBox "la información ha sido guardada con éxito", vbInformation
    cmdGuardar.Enabled = True
    Exit Sub
ErrorGuarda:
    MsgBox Err.Number & " - " & Err.Description, vbInformation, "Error"
End Sub
Private Sub ActualizaDatos()
    Dim index As Integer
    Dim nIdRep As Integer
    Dim oDMov As DMov
    Set oDMov = New DMov
    Dim lsMovNro As String
    Dim nValor As Double
    On Error GoTo ErrorGuarda
    cmdGuardar.Enabled = False
    lsMovNro = oDMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    nIdRep = 0
    nIdRep = oBalan.ActualizaConfigRepSBS13(Format(CboMes.ListIndex + 1, "00"), Trim(txtAnio.Text), lsMovNro, 1)
    oBalan.ActualizaConfigRepSBS13Det nIdRep, "R. Crédito", CStr(txtRCredito.Text), 0, 0
    oBalan.ActualizaConfigRepSBS13Det nIdRep, "R. Mercado", CStr(txtRMercado.Text), 0, 0
    oBalan.ActualizaConfigRepSBS13Det nIdRep, "R. Operacional", CStr(txtROcupacional.Text), 0, 0
    oBalan.ActualizaConfigRepSBS13Det nIdRep, "Deuda Subordinada", CStr(txtDeudaSubordinada.Text), 0, 0
    oBalan.ActualizaConfigRepSBS13Det nIdRep, "PE Nivel 1", CStr(txtPENivel1.Text), 0, 0
    oBalan.ActualizaConfigRepSBS13Det nIdRep, "PE Nivel 2", CStr(txtPENivel2.Text), 0, 0
    oBalan.ActualizaConfigRepSBS13Det nIdRep, "PE Mes Actual", CStr(txtPEMesActual.Text), 0, 0
    oBalan.ActualizaConfigRepSBS13Det nIdRep, "Patrimonio Contable", CStr(txtPEPatrimonioContable.Text), 0, 0
    oBalan.ActualizaConfigRepSBS13Det nIdRep, "PE Mes Ant", CStr(txtPEMesAnterior.Text), 0, 0
    For index = 1 To feRepDet.Rows - 1
        nValor = IIf(feRepDet.TextMatrix(index, 4) = "N/A", -1, (Replace(feRepDet.TextMatrix(index, 4), "%", "")))
        oBalan.ActualizaConfigRepSBS13Det nIdRep, feRepDet.TextMatrix(index, 1), feRepDet.TextMatrix(index, 2), feRepDet.TextMatrix(index, 3), nValor
    Next
    MsgBox "la información ha sido actualizada con éxito", vbInformation, "Aviso"
    cmdGuardar.Enabled = True
    Exit Sub
ErrorGuarda:
    MsgBox Err.Number & " - " & Err.Description, vbInformation, "Error"
End Sub
Private Sub cmdHojaTrabajo_Click()
    Dim celda As Excel.Range
    Dim sPathFormatoRep13Guarda As String
    Dim sPathFormatoRep13Lee As String
    Dim fs As New Scripting.FileSystemObject
    Dim obj_excel As Object, Libro As Object, Hoja As Object
    On Error GoTo error_sub
    If Not ValidaInterfaz Then Exit Sub
    If feRepDet.TextMatrix(feRepDet.row, 1) = "" Then
        MsgBox "Se necesita realizar el calculo de los ratios para poder crear la hoja de trabajo.", vbInformation
        cmdRecalcular.SetFocus
        Exit Sub
    End If
    cmdHojaTrabajo.Enabled = False
    Me.MousePointer = vbHourglass
    sPathFormatoRep13Guarda = App.path & "\Spooler\Rep13" & dameNombreMes(CboMes.ListIndex + 1, True) & Trim(txtAnio.Text) & ".xlsx"
    If fs.FileExists(sPathFormatoRep13Guarda) Then
        If ArchivoEstaAbierto(sPathFormatoRep13Guarda) Then
            If MsgBox("Debe Cerrar el Archivo: " + fs.GetFileName(sPathFormatoRep13Guarda) + " para continuar", vbRetryCancel) = vbCancel Then
                Me.MousePointer = vbDefault
                Exit Sub
            End If
            Me.MousePointer = vbHourglass
        End If
        fs.DeleteFile sPathFormatoRep13Guarda, True
    End If
    sPathFormatoRep13Lee = App.path & "\FormatoCarta\Reporte13.xlsx"
    If Len(Dir(sPathFormatoRep13Lee)) = 0 Then
           MsgBox "No se Pudo Encontrar el Archivo:" & sPathFormatoRep13Lee, vbCritical
           Me.MousePointer = vbDefault
           Exit Sub
    End If
    Set obj_excel = CreateObject("Excel.Application")
    obj_excel.DisplayAlerts = False
    Set Libro = obj_excel.Workbooks.Open(sPathFormatoRep13Lee)
    Set Hoja = Libro.ActiveSheet
    
    CargaDataExcel obj_excel
    Hoja.SaveAs sPathFormatoRep13Guarda
    Libro.Close
    obj_excel.Quit
    Set Hoja = Nothing
    Set Libro = Nothing
    Set obj_excel = Nothing
    Me.MousePointer = vbDefault
    
    Dim m_excel As New Excel.Application
    m_excel.Workbooks.Open (sPathFormatoRep13Guarda)
    m_excel.Visible = True
    cmdHojaTrabajo.Enabled = True
    Exit Sub
error_sub:
    MsgBox TextErr(Err.Description), vbInformation, "Aviso", "Aviso"
        Set Libro = Nothing
        Set obj_excel = Nothing
        Set Hoja = Nothing
        Me.MousePointer = vbDefault
End Sub
Private Sub CargaDataExcel(ByVal pobj_Excel As Excel.Application)
    Dim nfil As Integer
    Dim nIndex As Integer
    Dim celdaValor As Excel.Range
    nfil = 1
    Set celdaValor = pobj_Excel.Range("Rep!B" & nfil)
    celdaValor.value = LCase(Left(dameNombreMes(CboMes.ListIndex + 1, True), 3) & "-" & Right(Trim(txtAnio.Text), 2))
    nfil = 2
    Set celdaValor = pobj_Excel.Range("Rep!B" & nfil)
    celdaValor.value = Trim(txtPEMesAnterior.Text)
    nfil = 3
    Set celdaValor = pobj_Excel.Range("Rep!B" & nfil)
    celdaValor.value = Trim(txtPEPatrimonioContable.Text)
    nfil = 22
    Set celdaValor = pobj_Excel.Range("Rep!B" & nfil)
    celdaValor.value = Trim(txtPENivel1.Text)
    nfil = 27
    Set celdaValor = pobj_Excel.Range("Rep!B" & nfil)
    celdaValor.value = Trim(txtPENivel2.Text)
    nfil = 28
    Set celdaValor = pobj_Excel.Range("Rep!B" & nfil)
    celdaValor.value = Trim(txtDeudaSubordinada.Text)
    nfil = 30
    Set celdaValor = pobj_Excel.Range("Rep!B" & nfil)
    celdaValor.value = Trim(txtPEMesActual.Text)
    nfil = 4
    For nIndex = 1 To feRepDet.Rows - 1
        Set celdaValor = pobj_Excel.Range("Rep!B" & nfil)
        celdaValor.Formula = "=" & CDbl(feRepDet.TextMatrix(nIndex, 2)) & "/" & CDbl(feRepDet.TextMatrix(nIndex, 3))
        'celdaValor.value = feRepDet.TextMatrix(nIndex, 2)
        nfil = IIf(nfil = 21, nfil + 2, nfil + 1)
    Next
End Sub
Private Sub cmdRecalcular_Click()
    Dim i As Integer
    Dim sDesc As String
    Dim nParcialProgress As Currency
    Dim nMes As Integer: Dim nAnio As Integer: Dim dFecha As Date
    Dim rs As ADODB.Recordset
    Dim nPorcent As Double
    If Not ValidaInterfaz Then Exit Sub
    If Not ValidaDatos(False) Then Exit Sub
    
    On Error GoTo ErrorCalcula
    cmdRecalcular.Enabled = False
    
    LimpiaFlex feRepDet
    lblCalculando.Visible = True
    prgProgress.Visible = True
    fraBackground.Visible = True
    lblCalculoDesc.Visible = True
    lblPorcentajeProgreso.Visible = True
    lblCalculando = "Calculando..."
    lblCalculoDesc.Caption = ""
    lblPorcentajeProgreso.Caption = "0%"
    nParcialProgress = 0
    nPorcent = 4.76
    nMes = CboMes.ListIndex + 1: nAnio = txtAnio: dFecha = DateAdd("m", 1, "01/" & Format(nMes, "00") & "/" & Format(txtAnio, "0000")) - 1
    
    DoEvents
    For i = 1 To 21
        feRepDet.AdicionaFila
        Select Case i
            Case 1:
                sDesc = "1. Ratio de Capital Global"
                feRepDet.TextMatrix(i, 1) = sDesc
                feRepDet.TextMatrix(i, 2) = CDbl(txtPENivel1.Text) + CDbl(txtPENivel2.Text)
                feRepDet.TextMatrix(i, 3) = CDbl(txtRCredito.Text) + CDbl(txtRMercado.Text) + CDbl(txtROcupacional.Text)
                feRepDet.TextMatrix(i, 4) = CStr(Round(((CDbl(txtPENivel1.Text) + CDbl(txtPENivel2.Text)) / ((CDbl(txtRCredito.Text) + CDbl(txtRMercado.Text) + CDbl(txtROcupacional.Text)))) * 100, 2)) + "%"
            Case 2:
                sDesc = "2.2. Patrimonio Suplementario"
                feRepDet.TextMatrix(i, 1) = sDesc
                feRepDet.TextMatrix(i, 2) = CDbl(txtPENivel2.Text)
                feRepDet.TextMatrix(i, 3) = CDbl(txtPENivel1.Text)
                feRepDet.TextMatrix(i, 4) = CStr(Round((CDbl(txtPENivel2.Text) / CDbl(txtPENivel1.Text)) * 100, 2)) + "%"
            Case 3:
                sDesc = "2.3. Deuda Subordinada Redimible del Patrimonio de Nivel 2"
                feRepDet.TextMatrix(i, 1) = sDesc
                 feRepDet.TextMatrix(i, 2) = CDbl(txtDeudaSubordinada.Text)
                feRepDet.TextMatrix(i, 3) = CDbl(txtPENivel1.Text)
                feRepDet.TextMatrix(i, 4) = CStr(Round((CDbl(txtDeudaSubordinada.Text) / CDbl(txtPENivel1.Text)) * 100, 2)) + "%"
            Case 4:
                Dim nMontoA As Double: Dim nMontoB As Double: Dim nMontoTotal As Double
                sDesc = "4. Requerimiento mínimo de liquidez en Moneda Nacional"
                nMontoTotal = Round(ObtieneRequerimientoMinimodeLiquidez(dFecha, "1", "A1", nMontoA, nMontoB), 2)
                feRepDet.TextMatrix(i, 1) = sDesc
                feRepDet.TextMatrix(i, 2) = Round(nMontoA, 2)
                feRepDet.TextMatrix(i, 3) = Round(nMontoB, 2)
                feRepDet.TextMatrix(i, 4) = CStr(nMontoTotal) + "%"
            Case 5:
                Dim nMontoAME As Double: Dim nMontoBME As Double: Dim nMontoTotalME As Double
                sDesc = "5. Requerimiento mínimo de liquidez en Moneda Extranjera"
                nMontoTotalME = Round(ObtieneRequerimientoMinimodeLiquidez(dFecha, "2", "A1", nMontoAME, nMontoBME), 2)
                feRepDet.TextMatrix(i, 1) = sDesc
                 feRepDet.TextMatrix(i, 2) = Round(nMontoAME, 2)
                feRepDet.TextMatrix(i, 3) = Round(nMontoBME, 2)
                feRepDet.TextMatrix(i, 4) = CStr(nMontoTotalME) + "%"
            Case 6:
                Dim nMontoCap As Double
                sDesc = "6. Inversiones en Capital Social de Subsidiarias"
                nMontoCap = Round(CDbl(ObtieneMontosBalance(6, Trim(txtAnio.Text), Format(nMes, "00"))), 2)
                feRepDet.TextMatrix(i, 1) = sDesc
                feRepDet.TextMatrix(i, 2) = nMontoCap
                feRepDet.TextMatrix(i, 3) = CDbl(txtPEPatrimonioContable.Text)
                feRepDet.TextMatrix(i, 4) = CStr(Round((nMontoCap / CDbl(txtPEPatrimonioContable.Text)) * 100, 2)) + "%"
            Case 7:
                Dim nMontoOro As Double
                sDesc = "8. Tenencia de Oro"
                nMontoOro = Round(CDbl(ObtieneMontosBalance(7, Trim(txtAnio.Text), Format(nMes, "00"))), 2)
                feRepDet.TextMatrix(i, 1) = sDesc
                feRepDet.TextMatrix(i, 2) = nMontoOro
                feRepDet.TextMatrix(i, 3) = CDbl(txtPEMesAnterior.Text)
                feRepDet.TextMatrix(i, 4) = CStr(Round((nMontoOro / CDbl(txtPEMesAnterior.Text)) * 100, 2)) + "%"
            Case 8:
                Dim nMontoInm As Double
                sDesc = "10. Inversión en muebles e inmuebles"
                nMontoInm = Round(CDbl(ObtieneMontosBalance(8, Trim(txtAnio.Text), Format(nMes, "00"))), 2)
                feRepDet.TextMatrix(i, 1) = sDesc
                feRepDet.TextMatrix(i, 2) = nMontoInm
                feRepDet.TextMatrix(i, 3) = CDbl(txtPEMesAnterior.Text)
                feRepDet.TextMatrix(i, 4) = CStr(Round((nMontoInm / CDbl(txtPEMesAnterior.Text)) * 100, 2)) + "%"
            Case 9:
                Dim nMonto9 As Double
                sDesc = "11.1. Limite de la Posición Global de SobreVenta de ME"
                nMonto9 = Round((CDbl(ObtieneMontosBalance(9, Trim(txtAnio.Text), Format(nMes, "00")))), 2)
                feRepDet.TextMatrix(i, 1) = sDesc
                feRepDet.TextMatrix(i, 2) = nMonto9
                feRepDet.TextMatrix(i, 3) = CDbl(txtPEMesAnterior.Text)
                feRepDet.TextMatrix(i, 4) = CStr(IIf(nMonto9 = Format(0, "#,#0.00"), "N/A", Round((nMonto9 / CDbl(txtPEMesAnterior.Text)) * 100, 2))) + IIf(nMonto9 = Format(0, "#,#0.00"), "", "%")
            Case 10:
                Dim nMonto_10 As Double
                sDesc = "11.2. Limite de la Posición Global de SobreCompra de ME"
                nMonto_10 = Round((CDbl(ObtieneMontosBalance(10, Trim(txtAnio.Text), Format(nMes, "00")))), 2)
                feRepDet.TextMatrix(i, 1) = sDesc
                feRepDet.TextMatrix(i, 2) = nMonto_10
                feRepDet.TextMatrix(i, 3) = CDbl(txtPEMesAnterior.Text)
                feRepDet.TextMatrix(i, 4) = CStr(IIf(nMonto_10 = Format(0, "#,#0.00"), "N/A", Round((nMonto_10 / CDbl(txtPEMesAnterior.Text)) * 100))) + IIf(nMonto_10 = Format(0, "#,#0.00"), "", "%")
            Case 11:
                Dim nMontoCredDir As Double
                sDesc = "12. Créditos a Directores y trabajadores de la empresa"
                nMontoCredDir = Round(CDbl(ObtieneRepSBS13CredDirecTrabEmpresa(dFecha)), 2)
                feRepDet.TextMatrix(i, 1) = sDesc
                feRepDet.TextMatrix(i, 2) = nMontoCredDir
                feRepDet.TextMatrix(i, 3) = CDbl(txtPEMesAnterior.Text)
                feRepDet.TextMatrix(i, 4) = CStr(Round((nMontoCredDir / CDbl(txtPEMesAnterior.Text)) * 100, 2)) + "%"
            Case 12:
                Dim nMontoFinan As Double
                sDesc = "13. Financiamiento a personas vinculadas a la empresa"
                nMontoFinan = Round(CDbl(ObtieneFinanciamientoPersonasVincEmpresa(dFecha)), 2)
                feRepDet.TextMatrix(i, 1) = sDesc
                feRepDet.TextMatrix(i, 2) = nMontoFinan
                feRepDet.TextMatrix(i, 3) = CDbl(txtPEMesAnterior.Text)
                feRepDet.TextMatrix(i, 4) = CStr(Round((nMontoFinan / CDbl(txtPEMesAnterior.Text)) * 100, 2)) + "%"
            Case 13:
                Set rs = ObtieneRepSBS13FinanciamientoDirectoeIndirectoEmpSistFinan(dFecha)
                Do While Not rs.EOF
                    feRepDet.AdicionaFila
                    sDesc = "2. Financiamiento directo e indirecto a Empresas del sistema financiero (" & CStr(i - 12) & ")"
                    feRepDet.TextMatrix(i, 1) = sDesc
                    feRepDet.TextMatrix(i, 2) = Round(CDbl(rs!nMonto), 2)
                    feRepDet.TextMatrix(i, 3) = CDbl(txtPEMesAnterior.Text)
                    feRepDet.TextMatrix(i, 4) = CStr(Round((CDbl(rs!nMonto) / CDbl(txtPEMesAnterior.Text)) * 100, 2)) + "%"
                    lblCalculoDesc = sDesc
                    prgProgress.value = i
                    nParcialProgress = Round(nParcialProgress + (nPorcent), 2)
                    lblPorcentajeProgreso.Caption = CStr(IIf(nParcialProgress > 100, 100, nParcialProgress)) + "%"
                    i = i + 1
                    rs.MoveNext
                Loop
                i = 15
            Case 16:
                sDesc = "4. Financiamiento mediante créditos, inversiones y contingentes (1)"
                feRepDet.TextMatrix(i, 1) = sDesc
                feRepDet.TextMatrix(i, 2) = 0
                feRepDet.TextMatrix(i, 3) = 0
                feRepDet.TextMatrix(i, 4) = "0%"

            Case 17:
                sDesc = "4. Financiamiento mediante créditos, inversiones y contingentes (2)"
                feRepDet.TextMatrix(i, 1) = sDesc
                feRepDet.TextMatrix(i, 2) = 0
                feRepDet.TextMatrix(i, 3) = 0
                feRepDet.TextMatrix(i, 4) = "0%"
            Case 18:
                sDesc = "4. Financiamiento mediante créditos, inversiones y contingentes (3)"
                feRepDet.TextMatrix(i, 1) = sDesc
                feRepDet.TextMatrix(i, 2) = 0
                feRepDet.TextMatrix(i, 3) = 0
                feRepDet.TextMatrix(i, 4) = "0%"
            Case 19:
                Set rs = ObtieneRepSBS13CredDirecTrabEmpresaMax(dFecha)
                 Do While Not rs.EOF
                    feRepDet.AdicionaFila
                    sDesc = "8. Créditos a directores y trabajadores de la empresa (" & CStr(i - 18) & ")"
                    feRepDet.TextMatrix(i, 1) = sDesc
                    feRepDet.TextMatrix(i, 2) = Round(CDbl(rs!nMonto), 2)
                    feRepDet.TextMatrix(i, 3) = CDbl(txtPEMesAnterior.Text)
                    feRepDet.TextMatrix(i, 4) = CStr(Round((CDbl(rs!nMonto) / CDbl(txtPEMesAnterior.Text)) * 100, 2)) + "%"
                    lblCalculoDesc = sDesc
                    prgProgress.value = i
                    nParcialProgress = Round(nParcialProgress + (nPorcent), 2)
                    lblPorcentajeProgreso.Caption = CStr(IIf(nParcialProgress > 100, 100, nParcialProgress)) + "%"
                    i = i + 1
                    rs.MoveNext
                Loop
                i = 21
        End Select
        lblCalculoDesc = sDesc
        prgProgress.value = i
        nParcialProgress = Round(nParcialProgress + (nPorcent), 2)
        lblPorcentajeProgreso.Caption = CStr(IIf(nParcialProgress > 100, 100, nParcialProgress)) + "%"
    Next i
    lblCalculando.Visible = False
    prgProgress.Visible = False
    fraBackground.Visible = False
    lblCalculoDesc.Visible = False
    lblPorcentajeProgreso.Visible = False
    cmdRecalcular.Enabled = True
    Exit Sub
ErrorCalcula:
    MsgBox "Al parecer los datos son inconsistentes. Por ahora no se puede realizar el cálculo, por favor verifique.", vbInformation, "Aviso"
    LimpiaFlex feRepDet
    lblCalculando.Visible = False
    prgProgress.Visible = False
    fraBackground.Visible = False
    lblCalculoDesc.Visible = False
    lblPorcentajeProgreso.Visible = False
    cmdRecalcular.Enabled = True
End Sub
Private Function ValidaDatos(ByVal phOpe As Boolean) As Boolean
On Error GoTo ErrorValidaDatos
ValidaDatos = False
    If CDbl(txtRCredito.Text) = 0 Then
        If MsgBox("El valor del Riesgo de Credito es igual a cero. Esta seguro de " & IIf(phOpe, "grabar", "recalcular") & " los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        txtRCredito.SetFocus
        Exit Function
        End If
    End If
    If CDbl(txtRMercado.Text) = 0 Then
        If MsgBox("El valor del Riesgo de Mercado es igual a cero. Esta seguro de " & IIf(phOpe, "grabar", "recalcular") & " los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        txtRMercado.SetFocus
        Exit Function
        End If
    End If
    If CDbl(txtROcupacional.Text) = 0 Then
        If MsgBox("El valor del Riesgo Operacional es igual a cero. Esta seguro de " & IIf(phOpe, "grabar", "recalcular") & " los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        txtROcupacional.SetFocus
        Exit Function
        End If
    End If
     If CDbl(txtDeudaSubordinada.Text) = 0 Then
        If MsgBox("El valor de Deuda subordinada es igual a cero. Esta seguro de " & IIf(phOpe, "grabar", "recalcular") & " los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        txtDeudaSubordinada.SetFocus
        Exit Function
        End If
    End If
    If CDbl(txtPENivel1.Text) = 0 Then
        If MsgBox("El valor de Patrimonio Efectivo Nivel 1 es igual a cero. Esta seguro de " & IIf(phOpe, "grabar", "recalcular") & " los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        txtPENivel1.SetFocus
        Exit Function
        End If
    End If
    If CDbl(txtPENivel2.Text) = 0 Then
        If MsgBox("El valor de Patrimonio Efectivo Nivel 2 es igual a cero. Esta seguro de " & IIf(phOpe, "grabar", "recalcular") & " los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        txtPENivel2.SetFocus
        Exit Function
        End If
    End If
    If CDbl(txtPEMesActual.Text) = 0 Then
        If MsgBox("El valor de Patrimonio Efectivo Mes Actual es igual a cero. Esta seguro de " & IIf(phOpe, "grabar", "recalcular") & " los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        txtPEMesActual.SetFocus
        Exit Function
        End If
    End If
    If CDbl(txtPEPatrimonioContable.Text) = 0 Then
        If MsgBox("El valor de Patrimonio Efectivo Contable es igual a cero. Esta seguro de " & IIf(phOpe, "grabar", "recalcular") & " los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        txtPEPatrimonioContable.SetFocus
        Exit Function
        End If
    End If
     If CDbl(txtPEMesAnterior.Text) = 0 Then
        If MsgBox("El valor de Patrimonio Efectivo Mes Anterior es igual a cero. Esta seguro de " & IIf(phOpe, "grabar", "recalcular") & " los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        txtPEMesAnterior.SetFocus
        Exit Function
        End If
    End If
ValidaDatos = True
Exit Function
ErrorValidaDatos:
        MsgBox "Al parecer los datos son inconsistentes. Por ahora no se puede realizar el cálculo, por favor verifique.", vbInformation, "Aviso"
        ValidaDatos = False
End Function
Private Function ValidaInterfaz() As Boolean
    ValidaInterfaz = False
    If CboMes.ListIndex = -1 Then
        MsgBox "Asegurese de haber seleccionado el Mes del Periodo.", vbInformation, "Aviso"
        CboMes.SetFocus
        Exit Function
    End If
    If Len(Trim(txtAnio.Text)) = 0 Then
        MsgBox "Asegurese de haber ingresado correctamente el año del periodo.", vbInformation, "Aviso"
        txtAnio.SetFocus
        Exit Function
    End If
    If Len(Trim(txtRCredito.Text)) = 0 Then
        MsgBox "Asegurese de haber ingresado correctamente el Riesgo de Crédito.", vbInformation, "Aviso"
        txtRCredito.SetFocus
        Exit Function
    End If
     If Len(Trim(txtRMercado.Text)) = 0 Then
        MsgBox "Asegurese de haber ingresado correctamente el Riesgo de Mercado.", vbInformation, "Aviso"
        txtRMercado.SetFocus
        Exit Function
    End If
    If Len(Trim(txtROcupacional.Text)) = 0 Then
        MsgBox "Asegurese de haber ingresado correctamente el Riesgo Operacional.", vbInformation, "Aviso"
        txtROcupacional.SetFocus
        Exit Function
    End If
    If Len(Trim(txtDeudaSubordinada.Text)) = 0 Then
        MsgBox "Asegurese de haber ingresado correctamente la Deuda Subordinada.", vbInformation, "Aviso"
        txtDeudaSubordinada.SetFocus
        Exit Function
    End If
     If Len(Trim(txtPENivel1.Text)) = 0 Then
        MsgBox "Asegurese de haber ingresado correctamente Patrimonio Efectivo Nivel 1.", vbInformation, "Aviso"
        txtPENivel1.SetFocus
        Exit Function
    End If
     If Len(Trim(txtPENivel2.Text)) = 0 Then
        MsgBox "Asegurese de haber ingresado correctamente Patrimonio Efectivo Nivel 2.", vbInformation, "Aviso"
        txtPENivel2.SetFocus
        Exit Function
    End If
    If Len(Trim(txtPEMesActual.Text)) = 0 Then
        MsgBox "Asegurese de haber ingresado correctamente Patrimonio Efectivo Mes Actual.", vbInformation, "Aviso"
        txtPEMesActual.SetFocus
        Exit Function
    End If
    If Len(Trim(txtPEPatrimonioContable.Text)) = 0 Then
        MsgBox "Asegurese de haber ingresado correctamente Patrimonio Efectivo Contable.", vbInformation, "Aviso"
        txtPEPatrimonioContable.SetFocus
        Exit Function
    End If
    If Len(Trim(txtPEMesAnterior.Text)) = 0 Then
        MsgBox "Asegurese de haber ingresado correctamente Patrimonio Efectivo Mes Anterior.", vbInformation, "Aviso"
        txtPEMesAnterior.SetFocus
        Exit Function
    End If
    ValidaInterfaz = True
End Function
Private Sub cmdRepSBS13_Click()
     Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range
    Dim index As Integer
    Dim sValor As String
    If Not ValidaInterfaz Then Exit Sub
    If feRepDet.TextMatrix(feRepDet.row, 1) = "" Then
        MsgBox "Se necesita realizar el calculo de los ratios para poder crear el reporte SBS 13.", vbInformation, "Aviso"
        cmdRecalcular.SetFocus
        Exit Sub
    End If
    cmdRepSBS13.Enabled = False
    Set oWord = CreateObject("Word.Application")
    oWord.Visible = True
    Set oDoc = oWord.Documents.Open(App.path & "\FormatoCarta\Reporte13.docx")
    With oWord.Selection.Find
          .Text = "<<Periodo>>"
          .Replacement.Text = dameNombreMes(CboMes.ListIndex + 1, True) & "/" & Trim(txtAnio.Text)
          .Forward = True
          .Wrap = wdFindContinue
          .Format = False
          .Execute Replace:=wdReplaceAll
      End With
    For index = 1 To feRepDet.Rows - 1
        Select Case feRepDet.TextMatrix(index, 1)
            Case "1. Ratio de Capital Global":
                sValor = "<<RatCapGlo>>"
            Case "2.2. Patrimonio Suplementario":
                sValor = "<<PatSuple>>"
            Case "2.3. Deuda Subordinada Redimible del Patrimonio de Nivel 2":
                sValor = "<<DeudaSubor>>"
            Case "4. Requerimiento mínimo de liquidez en Moneda Nacional":
                sValor = "<<ReqMinLiqMN>>"
            Case "5. Requerimiento mínimo de liquidez en Moneda Extranjera":
                sValor = "<<ReqMinLiqME>>"
            Case "6. Inversiones en Capital Social de Subsidiarias":
                sValor = "<<InvCapSocSub>>"
            Case "8. Tenencia de Oro":
                sValor = "<<TenOro>>"
            Case "10. Inversión en muebles e inmuebles":
                sValor = "<<InvMuebInmue>>"
            Case "11.1. Limite de la Posición Global de SobreVenta de ME":
                sValor = "<<LimPosGloSobreVME>>"
            Case "11.2. Limite de la Posición Global de SobreCompra de ME":
                sValor = "<<LimPosGloSobreCME>>"
            Case "12. Créditos a Directores y trabajadores de la empresa":
                sValor = "<<CredDirTraEmp>>"
            Case "13. Financiamiento a personas vinculadas a la empresa":
                sValor = "<<FinanPersVincEmp>>"
            Case "2. Financiamiento directo e indirecto a Empresas del sistema financiero (1)":
                sValor = "<<FinanDirIndEmpSF1>>"
            Case "2. Financiamiento directo e indirecto a Empresas del sistema financiero (2)":
                sValor = "<<FinanDirIndEmpSF2>>"
            Case "2. Financiamiento directo e indirecto a Empresas del sistema financiero (3)":
                sValor = "<<FinanDirIndEmpSF3>>"
            Case "4. Financiamiento mediante créditos, inversiones y contingentes (1)":
                sValor = "<<FinanMedCredInvCt1>>"
            Case "4. Financiamiento mediante créditos, inversiones y contingentes (2)":
                sValor = "<<FinanMedCredInvCt2>>"
            Case "4. Financiamiento mediante créditos, inversiones y contingentes (3)":
                sValor = "<<FinanMedCredInvCt3>>"
            Case "8. Créditos a directores y trabajadores de la empresa (1)":
                sValor = "<<CredDirecTrabEmp81>>"
            Case "8. Créditos a directores y trabajadores de la empresa (2)":
                sValor = "<<CredDirecTrabEmp82>>"
            Case "8. Créditos a directores y trabajadores de la empresa (3)":
                sValor = "<<CredDirecTrabEmp83>>"
        End Select
        With oWord.Selection.Find
          .Text = sValor
          .Replacement.Text = Replace(IIf(feRepDet.TextMatrix(index, 4) = "N/A", "................", feRepDet.TextMatrix(index, 4)), "%", "")
          .Forward = True
          .Wrap = wdFindContinue
          .Format = False
          .Execute Replace:=wdReplaceAll
      End With
    Next
    oDoc.SaveAs App.path & "\SPOOLER\" & "Rep13" & dameNombreMes(CboMes.ListIndex + 1, True) & Trim(txtAnio.Text) & ".docx"
    cmdRepSBS13.Enabled = True
End Sub

Private Sub cmdSeleccionar_Click()
    LimpiaDatos
    If CboMes.ListIndex = -1 Then
        MsgBox "Asegurese de haber seleccionado el Mes del Periodo.", vbInformation
        CboMes.SetFocus
        Exit Sub
    End If
    If Len(Trim(txtAnio.Text)) = 0 Then
        MsgBox "Asegurese de haber ingresado correctamente el año del periodo.", vbInformation
        txtAnio.SetFocus
        Exit Sub
    End If
    If oBalan.ExisteConfigRepSBS13(Trim(txtAnio.Text), Format(CboMes.ListIndex + 1, "00")) Then
        EstableceDatos
    Else
        MsgBox "No existen datos para el periodo seleccionado.", vbInformation
    End If
    txtRCredito.SetFocus
End Sub
Private Sub LimpiaDatos()
    txtRCredito.Text = ""
    txtRMercado.Text = ""
    txtROcupacional.Text = ""
    txtDeudaSubordinada.Text = ""
    txtPENivel1.Text = ""
    txtPENivel2.Text = ""
    txtPEMesActual.Text = ""
    txtPEPatrimonioContable.Text = ""
    txtPEMesAnterior.Text = ""
    LimpiaFlex feRepDet
End Sub
Private Sub EstableceDatos()
    Dim rs As ADODB.Recordset
    Dim row As Integer
    txtRCredito.Text = Trim(oBalan.ObtieneValorxRepSBS13(Trim(txtAnio.Text), Format(CboMes.ListIndex + 1, "00"), "R. Crédito"))
    txtRMercado.Text = Trim(oBalan.ObtieneValorxRepSBS13(Trim(txtAnio.Text), Format(CboMes.ListIndex + 1, "00"), "R. Mercado"))
    txtROcupacional.Text = Trim(oBalan.ObtieneValorxRepSBS13(Trim(txtAnio.Text), Format(CboMes.ListIndex + 1, "00"), "R. Operacional"))
    txtDeudaSubordinada.Text = Trim(oBalan.ObtieneValorxRepSBS13(Trim(txtAnio.Text), Format(CboMes.ListIndex + 1, "00"), "Deuda Subordinada"))
    txtPENivel1.Text = Trim(oBalan.ObtieneValorxRepSBS13(Trim(txtAnio.Text), Format(CboMes.ListIndex + 1, "00"), "PE Nivel 1"))
    txtPENivel2.Text = Trim(oBalan.ObtieneValorxRepSBS13(Trim(txtAnio.Text), Format(CboMes.ListIndex + 1, "00"), "PE Nivel 2"))
    txtPEMesActual.Text = Trim(oBalan.ObtieneValorxRepSBS13(Trim(txtAnio.Text), Format(CboMes.ListIndex + 1, "00"), "PE Mes Actual"))
    txtPEPatrimonioContable.Text = Trim(oBalan.ObtieneValorxRepSBS13(Trim(txtAnio.Text), Format(CboMes.ListIndex + 1, "00"), "Patrimonio Contable"))
    txtPEMesAnterior.Text = Trim(oBalan.ObtieneValorxRepSBS13(Trim(txtAnio.Text), Format(CboMes.ListIndex + 1, "00"), "PE Mes Ant"))
    Set rs = oBalan.ObtieneValorxRepSBS13Ratios(Trim(txtAnio.Text), Format(CboMes.ListIndex + 1, "00"))
    Do While Not rs.EOF
        feRepDet.AdicionaFila
        row = feRepDet.row
        feRepDet.TextMatrix(row, 1) = rs!cConfRepDesc
        feRepDet.TextMatrix(row, 2) = rs!nConfRepVal1
        feRepDet.TextMatrix(row, 3) = rs!nConfRepVal2
        feRepDet.TextMatrix(row, 4) = (IIf(rs!nConfRepResultPorcent = -1, "N/A", CStr(rs!nConfRepResultPorcent) + "%"))
        rs.MoveNext
    Loop
End Sub
Private Sub feRepDet_OnCellChange(pnRow As Long, pnCol As Long)
    If pnCol = 2 Then
        feRepDet.TextMatrix(feRepDet.row, 3) = Round(IIf(feRepDet.TextMatrix(feRepDet.row, 3) = 0, 1, feRepDet.TextMatrix(feRepDet.row, 3)), 2)
        feRepDet.TextMatrix(feRepDet.row, 4) = CStr(Round((CDbl(feRepDet.TextMatrix(feRepDet.row, 2)) / CDbl(feRepDet.TextMatrix(feRepDet.row, 3)) * 100), 2)) & "%"
    End If
    If pnCol = 3 Then
        feRepDet.TextMatrix(feRepDet.row, 3) = Round(IIf(feRepDet.TextMatrix(feRepDet.row, 3) = 0, 1, feRepDet.TextMatrix(feRepDet.row, 3)), 2)
        feRepDet.TextMatrix(feRepDet.row, 4) = CStr(Round(((CDbl(feRepDet.TextMatrix(feRepDet.row, 2)) / CDbl(IIf(feRepDet.TextMatrix(feRepDet.row, 3) = 0, 1, feRepDet.TextMatrix(feRepDet.row, 3)))) * 100), 2)) & "%"
    End If
End Sub
'WIOR 20150819*****
Private Sub feRepDet_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)

If pnCol = 2 Or pnCol = 3 Then
    If IsNumeric(feRepDet.TextMatrix(feRepDet.row, pnCol)) = False Then
        Cancel = False
        SendKeys "{Tab}", True
        Exit Sub
    End If
End If

End Sub
'WIOR FIN ***

Private Sub Form_Load()
    Set oBalan = New DbalanceCont
    txtAnio.Text = DatePart("YYYY", gdFecSis)
End Sub
Private Sub txtAnio_GotFocus()
    fEnfoque txtAnio
End Sub
Private Sub txtAnio_KeyPress(KeyAscii As Integer)
    KeyAscii = TextBox_SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
       Me.cmdSeleccionar.SetFocus
    End If
End Sub
Private Sub txtAnio_LostFocus()
    If Trim(txtRCredito.Text) = "" Then
        txtAnio.Text = DatePart("YYYY", gdFecSis)
    End If
End Sub
Private Sub txtDeudaSubordinada_GotFocus()
    fEnfoque txtDeudaSubordinada
End Sub
Private Sub txtDeudaSubordinada_KeyPress(KeyAscii As Integer)
    KeyAscii = TextBox_SoloNumerosDecimales(KeyAscii)
    If KeyAscii = 13 Then
       Me.txtPENivel1.SetFocus
    End If
End Sub
Private Sub txtDeudaSubordinada_LostFocus()
    If Trim(txtDeudaSubordinada.Text) = "" Then
        txtDeudaSubordinada.Text = "0.00"
    End If
    txtDeudaSubordinada.Text = Format(txtDeudaSubordinada.Text, "#0.00")
End Sub
Private Sub txtPEMesActual_GotFocus()
    fEnfoque txtPEMesActual
End Sub
Private Sub txtPEMesActual_KeyPress(KeyAscii As Integer)
    KeyAscii = TextBox_SoloNumerosDecimales(KeyAscii)
    If KeyAscii = 13 Then
       Me.txtPEPatrimonioContable.SetFocus
    End If
End Sub
Private Sub txtPEMesActual_LostFocus()
    If Trim(txtPEMesActual.Text) = "" Then
        txtPEMesActual.Text = "0.00"
    End If
    txtPEMesActual.Text = Format(txtPEMesActual.Text, "#0.00")
End Sub
Private Sub txtPEMesAnterior_GotFocus()
    fEnfoque txtPEMesAnterior
End Sub
Private Sub txtPEMesAnterior_KeyPress(KeyAscii As Integer)
    KeyAscii = TextBox_SoloNumerosDecimales(KeyAscii)
    If KeyAscii = 13 Then
       Me.cmdRecalcular.SetFocus
    End If
End Sub
Private Sub txtPEMesAnterior_LostFocus()
     If Trim(txtPEMesAnterior.Text) = "" Then
        txtPEMesAnterior.Text = "0.00"
    End If
    txtPEMesAnterior.Text = Format(txtPEMesAnterior.Text, "#0.00")
End Sub
Private Sub txtPENivel1_GotFocus()
    fEnfoque txtPENivel1
End Sub
Private Sub txtPENivel1_KeyPress(KeyAscii As Integer)
    KeyAscii = TextBox_SoloNumerosDecimales(KeyAscii)
    If KeyAscii = 13 Then
       Me.txtPENivel2.SetFocus
    End If
End Sub
Private Sub txtPENivel1_LostFocus()
    If Trim(txtPENivel1.Text) = "" Then
        txtPENivel1.Text = "0.00"
    End If
    txtPEMesActual.Text = CDbl(txtPENivel1.Text) + CDbl(IIf(Len(txtPENivel2.Text) = 0, 0, txtPENivel2.Text))
    txtPENivel1.Text = Format(txtPENivel1.Text, "#0.00")
End Sub
Private Sub txtPENivel2_GotFocus()
    fEnfoque txtPENivel2
End Sub
Private Sub txtPENivel2_KeyPress(KeyAscii As Integer)
    KeyAscii = TextBox_SoloNumerosDecimales(KeyAscii)
    If KeyAscii = 13 Then
       Me.txtPEMesActual.SetFocus
    End If
End Sub
Private Sub txtPENivel2_LostFocus()
    If Trim(txtPENivel2.Text) = "" Then
        txtPENivel2.Text = "0.00"
    End If
     txtPEMesActual.Text = CDbl(IIf(Len(txtPENivel1.Text) = 0, 0, txtPENivel1.Text)) + CDbl(txtPENivel2.Text)
     txtPENivel2.Text = Format(txtPENivel2.Text, "#0.00")
End Sub
Private Sub txtPEPatrimonioContable_GotFocus()
    fEnfoque txtPEPatrimonioContable
End Sub
Private Sub txtPEPatrimonioContable_KeyPress(KeyAscii As Integer)
    KeyAscii = TextBox_SoloNumerosDecimales(KeyAscii)
    If KeyAscii = 13 Then
       Me.txtPEMesAnterior.SetFocus
    End If
End Sub
Private Sub txtPEPatrimonioContable_LostFocus()
     If Trim(txtPEPatrimonioContable.Text) = "" Then
        txtPEPatrimonioContable.Text = "0.00"
    End If
    txtPEPatrimonioContable.Text = Format(txtPEPatrimonioContable.Text, "#0.00")
End Sub
Private Sub txtRCredito_GotFocus()
    fEnfoque txtRCredito
End Sub
Private Sub txtRCredito_KeyPress(KeyAscii As Integer)
    KeyAscii = TextBox_SoloNumerosDecimales(KeyAscii)
    If KeyAscii = 13 Then
       Me.txtRMercado.SetFocus
    End If
End Sub
Private Sub txtRCredito_LostFocus()
    If Trim(txtRCredito.Text) = "" Then
        txtRCredito.Text = "0.00"
    End If
    txtRCredito.Text = Format(txtRCredito.Text, "#0.00")
End Sub
Private Sub txtRMercado_GotFocus()
    fEnfoque txtRMercado
End Sub
Private Sub txtRMercado_KeyPress(KeyAscii As Integer)
    KeyAscii = TextBox_SoloNumerosDecimales(KeyAscii)
    If KeyAscii = 13 Then
       Me.txtROcupacional.SetFocus
    End If
End Sub
Private Sub txtRMercado_LostFocus()
    If Trim(txtRMercado.Text) = "" Then
        txtRMercado.Text = "0.00"
    End If
    txtRMercado.Text = Format(txtRMercado.Text, "#0.00")
End Sub
Private Sub txtROcupacional_GotFocus()
    fEnfoque txtROcupacional
End Sub
Private Sub txtROcupacional_KeyPress(KeyAscii As Integer)
    KeyAscii = TextBox_SoloNumerosDecimales(KeyAscii)
    If KeyAscii = 13 Then
       Me.txtDeudaSubordinada.SetFocus
    End If
End Sub
Private Sub txtROcupacional_LostFocus()
    If Trim(txtROcupacional.Text) = "" Then
         txtROcupacional.Text = "0.00"
    End If
     txtROcupacional.Text = Format(txtROcupacional.Text, "#0.00")
End Sub
Public Function ObtieneRequerimientoMinimodeLiquidez(ByVal pdFecha As Date, ByVal psMoneda As String, ByVal psCaracter As String, Optional ByRef nMontoA As Double = 0, Optional ByRef nMontoB As Double = 0) As Double
Dim X As Integer: Dim nSaltoContadorX As Integer: Dim nSaltoContadorY As Integer
Dim nTotalA As Double: Dim nTotalB As Double:
Dim nMonto As Double
Dim bTotal As Boolean
Dim dFecha As Date
Dim rsRep15C As ADODB.Recordset

dFecha = DateAdd("d", -Day(pdFecha), pdFecha)
nSaltoContadorX = IIf(psMoneda = "1", 13, 51)
nSaltoContadorY = 4
bTotal = True
DoEvents
For X = 1 To Day(pdFecha)
    dFecha = DateAdd("d", 1, dFecha)
    nSaltoContadorX = IIf(psMoneda = "1", 13, 51)
    Set rsRep15C = oBalan.ObtenerListaReporte15B(Format(dFecha, "YYYY/MM/DD"), psMoneda, 0, psCaracter, Format(pdFecha, "YYYY/MM/DD"))
    Do While Not rsRep15C.EOF
        'nSaltoContadorY = 2 + X
        'If bTotal Then
            'nTotalA = nTotalA + rsRep15C!nSaldo
        'ElseIf Not bTotal And nSaltoContadorX <= IIf(psMoneda = "1", 30, 70) Then
            'nTotalB = nTotalB + rsRep15C!nSaldo
        'End If
        'If nSaltoContadorX = IIf(psMoneda = "1", 19, 59) Then
            'nSaltoContadorX = IIf(psMoneda = "1", 22, 62)
            'bTotal = False
        'ElseIf nSaltoContadorX = IIf(psMoneda = "1", 30, 70) Then
            'nSaltoContadorX = IIf(psMoneda = "1", 33, 74)
            'bTotal = False
        'ElseIf nSaltoContadorX = IIf(psMoneda = "1", 41, 83) Then
            'nSaltoContadorX = IIf(psMoneda = "1", 13, 51)
            'bTotal = True
        'Else
            'nSaltoContadorX = nSaltoContadorX + 1
        'End If
        'rsRep15C.MoveNext
        'If rsRep15C.EOF Then
            'Exit Do
        'End If 'Comentado by NAGL 20190914
        If rsRep15C!cCategoria = "AL" Then
            nTotalA = nTotalA + rsRep15C!nSaldo
        ElseIf rsRep15C!cCategoria = "PC" Then
            nTotalB = nTotalB + rsRep15C!nSaldo
        ElseIf rsRep15C!cCategoria = "RL" Then
            Exit Do
        End If
        rsRep15C.MoveNext
        'Agregado by NAGL 20190919 Según Correo e INC1909180007
    Loop
Next X
nMontoA = (nTotalA / Day(dFecha))
nMontoB = (nTotalB / Day(dFecha))
nMonto = (nMontoA / nMontoB) * 100
ObtieneRequerimientoMinimodeLiquidez = nMonto
End Function
Public Function ObtieneMontosBalance(ByVal pnOpcion As Integer, ByVal psAnio As String, ByVal PsMes As String)
    Dim nMonto As Double
    Dim nMontoXAlt As Double
    Dim nMontoYAlt As Double
    DoEvents
    Select Case pnOpcion
        Case 6:
                nMonto = oBalan.ObtieneRepSBS13InversionesCapitalSocialSubsidiarias(psAnio, PsMes)
        Case 7:
                nMonto = oBalan.ObtieneRepSBS13TenenciaOro(psAnio, PsMes)
        Case 8:
                nMonto = oBalan.ObtieneRepSBS13InversionMueblesInmuebles(psAnio, PsMes)
        Case 9:
                nMontoXAlt = oBalan.ObtieneRepSBS13LimitePosicionGlobal(psAnio, PsMes, "1")
                nMontoYAlt = oBalan.ObtieneRepSBS13LimitePosicionGlobal(psAnio, PsMes, "2")
                If nMontoXAlt > nMontoYAlt Then
                    nMonto = 0
                Else
                    nMonto = nMontoYAlt - nMontoXAlt
                End If
        Case 10:
                nMontoXAlt = oBalan.ObtieneRepSBS13LimitePosicionGlobal(psAnio, PsMes, "1")
                nMontoYAlt = oBalan.ObtieneRepSBS13LimitePosicionGlobal(psAnio, PsMes, "2")
                If nMontoYAlt > nMontoXAlt Then
                    nMonto = 0
                Else
                    nMonto = nMontoXAlt - nMontoYAlt
                End If
    End Select
    ObtieneMontosBalance = nMonto
End Function
Public Function ObtieneFinanciamientoPersonasVincEmpresa(ByVal pdFecha As Date) As Double
    Dim oCambio As nTipoCambio
    Dim nTipoCambio As Currency
    Dim nMonto As Double
    Set oCambio = New nTipoCambio
    DoEvents
    nTipoCambio = Format(oCambio.EmiteTipoCambio(pdFecha, TCFijoDia), "#,##0.00000")
    nMonto = oBalan.ObtieneRepSBS13FinanciamientoPersonasVincEmpresa(nTipoCambio, pdFecha)
    ObtieneFinanciamientoPersonasVincEmpresa = nMonto
End Function
Public Function ObtieneRepSBS13CredDirecTrabEmpresa(ByVal pdFecha As Date) As Double
    DoEvents
    ObtieneRepSBS13CredDirecTrabEmpresa = oBalan.ObtieneRepSBS13CredDirecTrabEmpresa(pdFecha)
End Function
Public Function ObtieneRepSBS13FinanciamientoDirectoeIndirectoEmpSistFinan(ByVal pdFecha As Date) As ADODB.Recordset
    DoEvents
    Set ObtieneRepSBS13FinanciamientoDirectoeIndirectoEmpSistFinan = oBalan.ObtieneRepSBS13FinanciamientoDirectoeIndirectoEmpSistFinan(pdFecha)
End Function
Public Function ObtieneRepSBS13CredDirecTrabEmpresaMax(ByVal pdFecha As Date) As ADODB.Recordset
    DoEvents
    Set ObtieneRepSBS13CredDirecTrabEmpresaMax = oBalan.ObtieneRepSBS13CredDirecTrabEmpresaMax(pdFecha)
End Function
