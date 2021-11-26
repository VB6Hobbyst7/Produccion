VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLogContAndRegComp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Logística: Adendas"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14475
   Icon            =   "frmLogContAndRegComp.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   14475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdComprobante 
      Caption         =   "C&omprobante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   11760
      TabIndex        =   8
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "S&alir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   13200
      TabIndex        =   7
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Frame frmLogContRegistro 
      Caption         =   "Actualizar"
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
      Height          =   1005
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   14280
      Begin VB.ComboBox cboTipoContrato 
         Height          =   315
         Left            =   4320
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   2100
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "A&ctualizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6600
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker txtFecIni 
         Height          =   315
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61734913
         CurrentDate     =   37156
      End
      Begin MSComCtl2.DTPicker txtFecFin 
         Height          =   315
         Left            =   2880
         TabIndex        =   4
         Top             =   360
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61734913
         CurrentDate     =   37156
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Final:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2400
         TabIndex        =   6
         Top             =   400
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Inicial:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   400
         Width           =   450
      End
   End
   Begin Sicmact.FlexEdit feCrontratos 
      Height          =   5595
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   14280
      _ExtentX        =   25188
      _ExtentY        =   9869
      Cols0           =   10
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Nº Adenda-Nº Contrato-Proveedor-Moneda-Monto-Desde-Hasta-Nº Cuota-Estado"
      EncabezadosAnchos=   "500-2000-2000-3500-1000-1200-1200-1200-1200-0"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-C-C-C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   7
      lbBuscaDuplicadoText=   -1  'True
      Appearance      =   0
      ColWidth0       =   495
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmLogContAndRegComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Public Sub Inicio()
Call CargarControles
Me.Show 1
End Sub

Private Sub CargarControles()
Dim oConst As DConstantes
Set oConst = New DConstantes

CargaCombo oConst.GetConstante(gsLogContTipoContratos), Me.cboTipoContrato
cboTipoContrato.AddItem Mid("TODOS" & Space(100), 1, 100) & "%"
cboTipoContrato.ListIndex = 2

txtFecIni.value = gdFecSis
txtFecFin.value = gdFecSis
End Sub

Private Sub cmdActualizar_Click()
Call CargarGrid
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub CargarGrid()
Dim oLog As DLogGeneral
Dim rsLog As ADODB.Recordset

Set oLog = New DLogGeneral
Set rsLog = oLog.ListarAdendasGeneral(Format(Me.txtFecIni.value, "DD/MM/YYYY"), Format(Me.txtFecFin.value, "DD/MM/YYYY"), Trim(Right(Me.cboTipoContrato.Text, 4)), "4,5")
Call LimpiaFlex(Me.feCrontratos)
If rsLog.RecordCount > 0 Then
    For i = 0 To rsLog.RecordCount - 1
        feCrontratos.AdicionaFila
        Me.feCrontratos.TextMatrix(i + 1, 0) = i + 1
        Me.feCrontratos.TextMatrix(i + 1, 1) = rsLog!cNAdenda
        Me.feCrontratos.TextMatrix(i + 1, 2) = rsLog!NContrato
        Me.feCrontratos.TextMatrix(i + 1, 3) = rsLog!Proveedor
        Me.feCrontratos.TextMatrix(i + 1, 4) = rsLog!Moneda
        Me.feCrontratos.TextMatrix(i + 1, 5) = rsLog!monto
        Me.feCrontratos.TextMatrix(i + 1, 6) = Format(rsLog!Desde, "dd/mm/yyyy")
        Me.feCrontratos.TextMatrix(i + 1, 7) = Format(rsLog!Hasta, "dd/mm/yyyy")
        Me.feCrontratos.TextMatrix(i + 1, 8) = oLog.ObtenerUltCuotaContratos(Trim(rsLog!NContrato))
        Me.feCrontratos.TextMatrix(i + 1, 9) = Trim(rsLog!nEstado)
        rsLog.MoveNext
    Next i
Else
    MsgBox "No hay datos.", vbInformation, "Aviso"
End If
End Sub
