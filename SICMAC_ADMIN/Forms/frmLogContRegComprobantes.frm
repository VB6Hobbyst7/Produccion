VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLogContRegComprobantes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Logística: Contratos"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12405
   Icon            =   "frmLogContRegComprobantes.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   12405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   11040
      TabIndex        =   9
      Top             =   6960
      Width           =   1215
   End
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
      Left            =   9600
      TabIndex        =   8
      Top             =   6960
      Width           =   1335
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
      Top             =   120
      Width           =   12120
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
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox cboTipoContrato 
         Height          =   315
         Left            =   4320
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   2100
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
         Format          =   61276161
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
         Format          =   61276161
         CurrentDate     =   37156
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Inicial:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   400
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Final:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2400
         TabIndex        =   1
         Top             =   400
         Width           =   375
      End
   End
   Begin Sicmact.FlexEdit feCrontratos 
      Height          =   5595
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   12120
      _extentx        =   21378
      _extenty        =   9869
      cols0           =   9
      highlight       =   1
      allowuserresizing=   3
      rowsizingmode   =   1
      encabezadosnombres=   "#-Nº Contrato-Proveedor-Moneda-Monto-Desde-Hasta-Nº Cuota-Estado"
      encabezadosanchos=   "500-2000-3500-1000-1200-1200-1200-1200-0"
      font            =   "frmLogContRegComprobantes.frx":030A
      font            =   "frmLogContRegComprobantes.frx":0332
      font            =   "frmLogContRegComprobantes.frx":035A
      font            =   "frmLogContRegComprobantes.frx":0382
      font            =   "frmLogContRegComprobantes.frx":03AA
      fontfixed       =   "frmLogContRegComprobantes.frx":03D2
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1  'True
      tipobusqueda    =   7
      columnasaeditar =   "X-X-X-X-X-X-X-X-X"
      textstylefixed  =   3
      listacontroles  =   "0-0-0-0-0-0-0-0-0"
      encabezadosalineacion=   "C-C-C-C-C-C-C-C-C"
      formatosedit    =   "0-0-0-0-0-0-0-0-0"
      textarray0      =   "#"
      lbeditarflex    =   -1  'True
      lbbuscaduplicadotext=   -1  'True
      appearance      =   0
      colwidth0       =   495
      rowheight0      =   300
      forecolorfixed  =   -2147483630
   End
End
Attribute VB_Name = "frmLogContRegComprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Public Function Inicio()
Call CargarControles
Me.Show 1
End Function

Private Sub cmdActualizar_Click()
Call CargarGrid
End Sub

Private Sub CargarGrid()
Dim oLog As DLogGeneral
Dim rsLog As ADODB.Recordset

Set oLog = New DLogGeneral
Set rsLog = oLog.ListarContratosGeneral(Format(Me.txtFecIni.value, "DD/MM/YYYY"), Format(Me.txtFecFin.value, "DD/MM/YYYY"), Trim(Right(Me.cboTipoContrato.Text, 4)), "4,5")
Call LimpiaFlex(Me.feCrontratos)
If rsLog.RecordCount > 0 Then
    For I = 0 To rsLog.RecordCount - 1
        feCrontratos.AdicionaFila
        Me.feCrontratos.TextMatrix(I + 1, 0) = I + 1
        Me.feCrontratos.TextMatrix(I + 1, 1) = rsLog!NContrato
        Me.feCrontratos.TextMatrix(I + 1, 2) = rsLog!Proveedor
        Me.feCrontratos.TextMatrix(I + 1, 3) = rsLog!Moneda
        Me.feCrontratos.TextMatrix(I + 1, 4) = rsLog!monto
        Me.feCrontratos.TextMatrix(I + 1, 5) = Format(rsLog!Desde, "dd/mm/yyyy")
        Me.feCrontratos.TextMatrix(I + 1, 6) = Format(rsLog!Hasta, "dd/mm/yyyy")
        Me.feCrontratos.TextMatrix(I + 1, 7) = oLog.ObtenerUltCuotaContratos(Trim(rsLog!NContrato))
        Me.feCrontratos.TextMatrix(I + 1, 8) = Trim(rsLog!nEstado)
        rsLog.MoveNext
    Next I
Else
    MsgBox "No hay datos.", vbInformation, "Aviso"
End If
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


Private Sub cmdSalir_Click()
Unload Me
End Sub
