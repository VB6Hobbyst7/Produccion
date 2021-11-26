VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAnexo16ALiquidezVenc 
   Caption         =   "Anexo 16A: Cuadro de Liquidez por Plazo de Vencimiento de Corto Plazo"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8115
   Icon            =   "frmAnexo16ALiquidezVenc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sTab 
      Height          =   3585
      Left            =   150
      TabIndex        =   23
      Top             =   1470
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   6324
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Conceptos "
      TabPicture(0)   =   "frmAnexo16ALiquidezVenc.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraConcepto"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Rangos   "
      TabPicture(1)   =   "frmAnexo16ALiquidezVenc.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraRango"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Cri&terios     "
      TabPicture(2)   =   "frmAnexo16ALiquidezVenc.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraCriterio"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraCriterio 
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
         Height          =   3165
         Left            =   -74940
         TabIndex        =   38
         Top             =   330
         Width           =   7725
         Begin VB.CommandButton cmdEditarCriterio 
            Caption         =   "&Editar..."
            Height          =   315
            Left            =   150
            TabIndex        =   42
            Top             =   2790
            Width           =   915
         End
         Begin VB.CommandButton cmdGrabaCriterio 
            Caption         =   "&Grabar"
            Height          =   315
            Left            =   5670
            TabIndex        =   40
            Top             =   2790
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.CommandButton cmdCancelaCriterio 
            Caption         =   "&Cancelar"
            Height          =   315
            Left            =   6600
            TabIndex        =   39
            Top             =   2790
            Visible         =   0   'False
            Width           =   915
         End
         Begin MSComctlLib.ListView lvConcepto 
            Height          =   2445
            Left            =   150
            TabIndex        =   41
            Top             =   300
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   4313
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Conceptos"
               Object.Width           =   4939
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Clase"
               Object.Width           =   353
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Grupo"
               Object.Width           =   705
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Concepto"
               Object.Width           =   705
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Tpo.Calculo"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Calculo"
               Object.Width           =   0
            EndProperty
         End
         Begin Sicmact.FlexEdit fgCriterio 
            Height          =   2445
            Left            =   5220
            TabIndex        =   49
            Top             =   300
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   4313
            Cols0           =   4
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "Rango-Rango1-Valor-Valor1"
            EncabezadosAnchos=   "800-0-1185-0"
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-2-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "L-L-R-L"
            FormatosEdit    =   "0-0-0-0"
            AvanceCeldas    =   1
            TextArray0      =   "Rango"
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   795
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame fraRango 
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
         Height          =   3165
         Left            =   -74940
         TabIndex        =   26
         Top             =   330
         Width           =   6795
         Begin VB.CommandButton cmdCancelaRango 
            Caption         =   "&Cancelar"
            Height          =   315
            Left            =   5730
            TabIndex        =   37
            Top             =   2790
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.CommandButton cmdGrabaRango 
            Caption         =   "&Grabar"
            Height          =   315
            Left            =   4800
            TabIndex        =   36
            Top             =   2790
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.CommandButton cmdEliminaRango 
            Caption         =   "&Eliminar"
            Height          =   315
            Left            =   2010
            TabIndex        =   35
            Top             =   2790
            Width           =   915
         End
         Begin VB.CommandButton cmdModificaRango 
            Caption         =   "&Modificar"
            Height          =   315
            Left            =   1080
            TabIndex        =   34
            Top             =   2790
            Width           =   915
         End
         Begin VB.CommandButton cmdNuevoRango 
            Caption         =   "&Nuevo"
            Height          =   315
            Left            =   150
            TabIndex        =   33
            Top             =   2790
            Width           =   915
         End
         Begin MSComctlLib.ListView lvRango 
            Height          =   2445
            Left            =   150
            TabIndex        =   27
            Top             =   300
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   4313
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Rango"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Descripción"
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Desde"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Hasta"
               Object.Width           =   1764
            EndProperty
         End
         Begin VB.Frame fraDatRango 
            Height          =   525
            Left            =   150
            TabIndex        =   28
            Top             =   2220
            Visible         =   0   'False
            Width           =   6495
            Begin VB.TextBox txtRangoDesc 
               Height          =   315
               Left            =   630
               MaxLength       =   50
               TabIndex        =   30
               Top             =   150
               Width           =   3585
            End
            Begin VB.TextBox txtDesde 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   4230
               MaxLength       =   7
               TabIndex        =   31
               Top             =   150
               Width           =   1095
            End
            Begin VB.TextBox txtRango 
               Height          =   315
               Left            =   60
               MaxLength       =   2
               TabIndex        =   29
               Top             =   150
               Width           =   555
            End
            Begin VB.TextBox txtHasta 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   5340
               MaxLength       =   7
               TabIndex        =   32
               Top             =   150
               Width           =   1095
            End
         End
      End
      Begin VB.Frame fraConcepto 
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
         Height          =   3165
         Left            =   60
         TabIndex        =   24
         Top             =   330
         Width           =   7725
         Begin VB.CommandButton cmdNuevoConcep 
            Caption         =   "&Nuevo"
            Height          =   315
            Left            =   150
            TabIndex        =   12
            Top             =   2790
            Width           =   915
         End
         Begin VB.CommandButton cmdModificaConcep 
            Caption         =   "&Modificar"
            Height          =   315
            Left            =   1080
            TabIndex        =   13
            Top             =   2790
            Width           =   915
         End
         Begin VB.CommandButton cmdEliminaConcep 
            Caption         =   "&Eliminar"
            Height          =   315
            Left            =   2010
            TabIndex        =   14
            Top             =   2790
            Width           =   915
         End
         Begin VB.CommandButton cmdGrabaConcep 
            Caption         =   "&Grabar"
            Height          =   315
            Left            =   5520
            TabIndex        =   15
            Top             =   2790
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.CommandButton cmdCancelaConcep 
            Caption         =   "&Cancelar"
            Height          =   315
            Left            =   6450
            TabIndex        =   16
            Top             =   2790
            Visible         =   0   'False
            Width           =   915
         End
         Begin MSComctlLib.ListView lvConcep 
            Height          =   2445
            Left            =   150
            TabIndex        =   4
            Top             =   300
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   4313
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Clase"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Grupo"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Concep"
               Object.Width           =   1323
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Descripción"
               Object.Width           =   4939
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Cálculo"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Cálculo"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Cuenta Contable"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Formula"
               Object.Width           =   2117
            EndProperty
         End
         Begin VB.Frame fraDatConcep 
            Height          =   525
            Left            =   150
            TabIndex        =   25
            Top             =   2220
            Visible         =   0   'False
            Width           =   7455
            Begin VB.TextBox txtFormula 
               Height          =   315
               Left            =   6300
               TabIndex        =   11
               Top             =   150
               Width           =   1095
            End
            Begin VB.TextBox txtCtaCod 
               Height          =   315
               Left            =   5280
               TabIndex        =   10
               Top             =   150
               Width           =   1005
            End
            Begin VB.ComboBox cboTpoCalculo 
               Height          =   315
               ItemData        =   "frmAnexo16ALiquidezVenc.frx":035E
               Left            =   4170
               List            =   "frmAnexo16ALiquidezVenc.frx":0360
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   150
               Width           =   1095
            End
            Begin VB.TextBox txtClase 
               Height          =   315
               Left            =   60
               MaxLength       =   1
               TabIndex        =   5
               Top             =   150
               Width           =   315
            End
            Begin VB.TextBox txtGrupo 
               Height          =   315
               Left            =   390
               MaxLength       =   2
               TabIndex        =   6
               Top             =   150
               Width           =   435
            End
            Begin VB.TextBox txtConcep 
               Height          =   315
               Left            =   840
               MaxLength       =   2
               TabIndex        =   7
               Top             =   150
               Width           =   435
            End
            Begin VB.TextBox txtConcepDesc 
               Height          =   315
               Left            =   1290
               MaxLength       =   50
               TabIndex        =   8
               Top             =   150
               Width           =   2865
            End
         End
      End
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
      Height          =   345
      Left            =   5205
      TabIndex        =   17
      Top             =   5160
      Width           =   1155
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   6420
      TabIndex        =   18
      Top             =   5160
      Width           =   1155
   End
   Begin VB.Frame fraRep 
      Height          =   1305
      Left            =   150
      TabIndex        =   19
      Top             =   90
      Width           =   7845
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de Cambio"
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
         Height          =   825
         Left            =   4320
         TabIndex        =   43
         Top             =   270
         Width           =   1635
         Begin VB.TextBox txtTipCambio 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   150
            TabIndex        =   3
            Top             =   300
            Width           =   1335
         End
      End
      Begin VB.Frame fraMes 
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
         Height          =   795
         Left            =   180
         TabIndex        =   20
         Top             =   270
         Width           =   4110
         Begin VB.TextBox txtAnio 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   630
            MaxLength       =   4
            TabIndex        =   1
            Top             =   300
            Width           =   855
         End
         Begin VB.ComboBox CboMes 
            Height          =   315
            ItemData        =   "frmAnexo16ALiquidezVenc.frx":0362
            Left            =   2160
            List            =   "frmAnexo16ALiquidezVenc.frx":038A
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   300
            Width           =   1800
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Año :"
            Height          =   195
            Left            =   180
            TabIndex        =   22
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Mes :"
            Height          =   195
            Left            =   1710
            TabIndex        =   21
            Top             =   390
            Width           =   390
         End
      End
      Begin VB.Frame fraFecha 
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
         Height          =   795
         Left            =   180
         TabIndex        =   47
         Top             =   270
         Visible         =   0   'False
         Width           =   4110
         Begin MSMask.MaskEdBox txtFechaAl 
            Height          =   345
            Left            =   2040
            TabIndex        =   0
            Top             =   270
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha al"
            Height          =   285
            Left            =   1020
            TabIndex        =   48
            Top             =   330
            Width           =   915
         End
      End
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   150
      TabIndex        =   44
      Top             =   5040
      Width           =   2625
      Begin VB.OptionButton OptOpc 
         Caption         =   "&Detallado"
         Height          =   225
         Index           =   1
         Left            =   1350
         TabIndex        =   46
         Top             =   180
         Width           =   1065
      End
      Begin VB.OptionButton OptOpc 
         Caption         =   "&Agrupado"
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   45
         Top             =   180
         Value           =   -1  'True
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmAnexo16ALiquidezVenc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim sSql       As String
Dim lNuevo     As Boolean
Dim nColRango  As Integer
Dim nTipCambio As Currency
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim lbMes       As Boolean
Dim lbIndicador As Boolean

Dim aPosicion() As String     'Guarda Filas llenadas
Dim nCont        As Integer   'Contador de Arreglo
Dim dbCmact As DConecta
Dim I As Integer, N As Integer
Dim oAnx        As DAnexoRiesgos

Public Sub Inicio(pbMes As Boolean, Optional pbIndicador As Boolean = False)
lbMes = pbMes
lbIndicador = pbIndicador
Me.Show 1
End Sub

Private Sub CboMes_Click()
If cboMes.ListIndex > -1 And txtAnio <> "" Then
    txtTipCambio = TipoCambioCierre(txtAnio, cboMes.ListIndex + 1)
    txtTipCambio.SetFocus
End If
End Sub

Private Sub cboMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtTipCambio = TipoCambioCierre(Year(Me.txtFechaAl), Month(txtFechaAl))
    txtTipCambio.SetFocus
End If
End Sub

Private Sub cboTpoCalculo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtCtaCod.SetFocus
End If
End Sub

Private Sub cmdCancelaConcep_Click()
HabilitaConcepto False
lvConcep.SetFocus
End Sub

Private Sub cmdCancelaCriterio_Click()
HabilitaCriterio False
End Sub

Private Sub cmdCancelaRango_Click()
HabilitaRango False
lvRango.SetFocus
End Sub

Private Sub cmdEditarCriterio_Click()
HabilitaCriterio True
fgCriterio.SetFocus
End Sub

Private Sub cmdEliminaConcep_Click()
If lvConcep.ListItems.Count = 0 Then
   Exit Sub
End If
If MsgBox("¿ Seguro que desea Eliminar Concepto ?", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
   Exit Sub
End If
oAnx.EliminaConcepto gsOpeCod, lvConcep.SelectedItem.Text, lvConcep.SelectedItem.SubItems(1), lvConcep.SelectedItem.SubItems(2)
lvConcep.ListItems.Remove lvConcep.SelectedItem.Index

End Sub

Private Sub cmdEliminaRango_Click()
On Error GoTo EliminaErr
If lvRango.ListItems.Count = 0 Then
   Exit Sub
End If
If MsgBox("¿ Seguro que desea Eliminar Rango ?", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
   Exit Sub
End If
oAnx.EliminaRango gsOpeCod, lvRango.SelectedItem.Text
lvRango.ListItems.Remove lvRango.SelectedItem.Index
Exit Sub
EliminaErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdGenerar_Click()
Dim nCol  As Integer
Dim sCol  As String

Dim lsArchivo   As String
Dim lsRuta      As String
Dim lbLibroOpen As Boolean
Dim N           As Integer

Dim nTipoCambio As Double
'On Error GoTo ErrImprime

MousePointer = 11
If Not ValidaDatos Then
   MousePointer = 0
   Exit Sub
End If

nTipoCambio = Val(txtTipCambio.Text)

If lbMes Then
   gdFecha = DateAdd("m", 1, CDate("01/" & Format(cboMes.ListIndex + 1, "00") & "/" & txtAnio)) - 1
Else
   gdFecha = CDate(txtFechaAl)
   txtAnio = Year(gdFecha)
   cboMes.ListIndex = Month(gdFecha) - 1
End If
txtTipCambio.Text = nTipoCambio

lsRuta = App.path & "\Spooler\"
lsArchivo = lsRuta & "Anx" & gsOpeCod & "_" & txtAnio & ".xls"
lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
   If lbLibroOpen Then
      Set xlHoja1 = xlLibro.Worksheets(1)
      ImprimeHoja "1"
      ImprimeHoja "2"
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
      CargaArchivo lsArchivo, lsRuta
   End If
   MousePointer = 0
Exit Sub
ErrImprime:
   MsgBox TextErr(Err.Description), vbInformation, "!Aviso!"
   If lbLibroOpen Then
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
      lbLibroOpen = False
   End If
   MousePointer = 0
End Sub
Private Sub ImprimeHoja(psMoneda As String)
If psMoneda = "1" Then
   nTipCambio = 1
Else
   nTipCambio = nVal(txtTipCambio)
End If
   If psMoneda = "1" Then
      ExcelAddHoja cboMes, xlLibro, xlHoja1
      Call CabeceraExcel(psMoneda, False)
      Call ImprimeRangos
   End If
   Call ImprimeConceptos(psMoneda)
End Sub

Private Sub cmdGrabaConcep_Click()
Dim nPos As Integer
If Not ValidaDatosConcep() Then
   Exit Sub
End If
If txtCtaCod <> "" And Trim(Right(cboTpoCalculo, 10)) <> "9" And Trim(Right(cboTpoCalculo, 10)) <> "8" And Trim(Right(cboTpoCalculo, 10)) <> "7" Then
   Dim oCta As New DCtaCont
   Set rs = oCta.CargaCtaCont("cCtaContCod LIKE '" & IIf(InStr("SN", Left(txtCtaCod, 1)) > 0, Mid(txtCtaCod, 2, 22), txtCtaCod) & "'")
   If rs.EOF Then
      RSClose rs
      MsgBox "Cuenta Contable no Existe", vbInformation, "¡Aviso!"
      Exit Sub
   End If
   RSClose rs
End If
nPos = lvConcep.SelectedItem.Index
If lNuevo Then
   oAnx.InsertaConcepto gsOpeCod, txtClase, txtGrupo, txtConcep, txtConcepDesc, Trim(Right(cboTpoCalculo.Text, 10)), txtCtaCod, txtFormula
Else
   oAnx.ActualizaConcepto gsOpeCod, txtClase, txtGrupo, txtConcep, txtConcepDesc, Trim(Right(cboTpoCalculo.Text, 10)), txtCtaCod, txtFormula
End If
HabilitaConcepto False
CargaConceptos
lvConcep.ListItems(nPos).Selected = True
lvConcep.SetFocus
End Sub

Private Sub cmdGrabaCriterio_Click()
Dim nPos As Integer
Dim I As Integer
For I = 1 To fgCriterio.Rows - 1
   If fgCriterio.TextMatrix(I, 3) = "" And fgCriterio.TextMatrix(I, 2) <> "" Then
      oAnx.InsertaCriterio gsOpeCod, lvConcepto.SelectedItem.SubItems(1), lvConcepto.SelectedItem.SubItems(2), lvConcepto.SelectedItem.SubItems(3), fgCriterio.TextMatrix(I, 1), fgCriterio.TextMatrix(I, 2)
   End If
   If fgCriterio.TextMatrix(I, 2) <> "" And fgCriterio.TextMatrix(I, 3) <> "" And fgCriterio.TextMatrix(I, 2) <> fgCriterio.TextMatrix(I, 3) Then
      oAnx.ActualizaCriterio gsOpeCod, lvConcepto.SelectedItem.SubItems(1), lvConcepto.SelectedItem.SubItems(2), lvConcepto.SelectedItem.SubItems(3), fgCriterio.TextMatrix(I, 1), fgCriterio.TextMatrix(I, 2)
   End If
   If fgCriterio.TextMatrix(I, 2) = "" And fgCriterio.TextMatrix(I, 3) <> "" Then
      oAnx.EliminaCriterio gsOpeCod, lvConcepto.SelectedItem.SubItems(1), lvConcepto.SelectedItem.SubItems(2), lvConcepto.SelectedItem.SubItems(3), fgCriterio.TextMatrix(I, 1)
   End If
Next
HabilitaCriterio False
lvConcepto.SetFocus
End Sub

Private Sub cmdGrabaRango_Click()
Dim nPos As Integer
If Not ValidaDatosRango() Then
   Exit Sub
End If
If lNuevo Then
   oAnx.InsertaRango gsOpeCod, txtRango, txtRangoDesc, txtDesde, txtHasta
Else
   oAnx.ActualizaRango gsOpeCod, txtRango, txtRangoDesc, txtDesde, txtHasta
End If
HabilitaRango False
CargaRangos
lvRango.SetFocus
End Sub

Private Sub cmdModificaConcep_Click()
lNuevo = False
HabilitaConcepto True
txtConcepDesc.SetFocus
End Sub

Private Sub cmdModificaRango_Click()
If lvRango.ListItems.Count = 0 Then
   Exit Sub
End If
lNuevo = False
HabilitaRango True
txtRangoDesc.SetFocus
End Sub

Private Sub cmdNuevoConcep_Click()
lNuevo = True
HabilitaConcepto True
txtClase.SetFocus
End Sub

Private Sub cmdNuevoRango_Click()
lNuevo = True
HabilitaRango True
txtRango.SetFocus
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
Me.Caption = gsOpeDesc
Dim oCons As New DConstantes
Set oAnx = New DAnexoRiesgos
Set rs = oCons.CargaConstante(gAnxTipoCalculo)
RSLlenaCombo rs, cboTpoCalculo, 1, 2
RSClose rs
cboTpoCalculo.AddItem "Ninguno" & Space(101)
CargaConceptos
sTab.Tab = 0
If lbMes Then
   txtAnio = Year(gdFecSis)
   cboMes.ListIndex = Month(gdFecSis) - 1
Else
   fraFecha.Visible = True
   fraMes.Visible = False
   txtFechaAl = gdFecSis
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
CierraConexion
End Sub

Private Sub lvConcepto_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lvConcepto.SortKey = ColumnHeader.Index - 1
lvConcepto.Sorted = True
End Sub

Private Sub lvConcepto_ItemClick(ByVal Item As MSComctlLib.ListItem)
If lvConcepto.ListItems.Count > 0 Then
   CargaRangoCriterio lvConcepto.SelectedItem.SubItems(1), lvConcepto.SelectedItem.SubItems(2), lvConcepto.SelectedItem.SubItems(3)
End If
End Sub

Private Sub sTab_Click(PreviousTab As Integer)
Select Case sTab.Tab
   Case 0
      FraConcepto.Enabled = True
      fraRango.Enabled = False
      CargaConceptos
   Case 1
      FraConcepto.Enabled = False
      fraRango.Enabled = True
      CargaRangos
   Case 2
      FraConcepto.Enabled = False
      fraRango.Enabled = False
      CargaConceptosCriterio
      If lvConcepto.ListItems.Count > 0 Then
         CargaRangoCriterio lvConcepto.SelectedItem.Text, lvConcepto.SelectedItem.SubItems(1), lvConcepto.SelectedItem.SubItems(2)
      End If
End Select
End Sub

Private Sub CargaRangoCriterio(psClase As String, psGrupo As String, psConcepto As String)
fgCriterio.Clear
fgCriterio.FormaCabecera
fgCriterio.Rows = 2
If lvConcepto.ListItems.Count > 0 Then
   Set rs = oAnx.CargaRangoCriterio(gsOpeCod, psClase, psGrupo, psConcepto)
   Do While Not rs.EOF
      fgCriterio.AdicionaFila
      fgCriterio.TextMatrix(fgCriterio.Row, 0) = rs!cDescrip
      fgCriterio.TextMatrix(fgCriterio.Row, 1) = rs!cCodRango
      fgCriterio.TextMatrix(fgCriterio.Row, 2) = rs!cValor
      fgCriterio.TextMatrix(fgCriterio.Row, 3) = rs!cValor
      rs.MoveNext
   Loop
   RSClose rs
End If
fgCriterio.Row = 1
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
   KeyAscii = NumerosEnteros(KeyAscii)
   If KeyAscii = 13 Then
      cboMes.SetFocus
   End If
End Sub

Private Sub txtClase_GotFocus()
fEnfoque txtClase
End Sub

Private Sub txtClase_KeyPress(KeyAscii As Integer)
If InStr("12345", Chr(KeyAscii)) = 0 And Not KeyAscii = 13 And Not KeyAscii = 8 Then
   KeyAscii = 0
End If
If KeyAscii = 13 Then
   txtGrupo.SetFocus
End If
End Sub

Private Sub txtConcep_GotFocus()
fEnfoque txtConcep
End Sub

Private Sub txtConcep_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   txtConcepDesc.SetFocus
End If
End Sub

Private Sub txtConcepDesc_GotFocus()
fEnfoque txtConcepDesc
End Sub

Private Sub txtConcepDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboTpoCalculo.SetFocus
End If
End Sub

Private Sub txtCtaCod_KeyPress(KeyAscii As Integer)
If InStr("0123456789_[]%", Chr(KeyAscii)) = 0 And Not KeyAscii = 13 And Not KeyAscii = 8 Then
   Exit Sub
End If
If KeyAscii = 13 Then
   txtFormula.SetFocus
End If
End Sub

Private Sub txtDesde_GotFocus()
fEnfoque txtDesde
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   txtHasta.SetFocus
End If
End Sub

Private Sub txtFechaAl_GotFocus()
If Val(txtTipCambio.Text) > 0 Then
Else
    If Day(DateAdd("d", 1, txtFechaAl)) = 1 Then
    Else
       txtTipCambio = TipoCambioCierre(Year(Me.txtFechaAl), Month(txtFechaAl))
    End If
End If
fEnfoque txtFechaAl
End Sub

Private Sub txtFechaAl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtTipCambio.SetFocus
End If
End Sub

Private Sub txtFormula_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdGrabaConcep.SetFocus
End If
End Sub

Private Sub txtGrupo_GotFocus()
fEnfoque txtGrupo
End Sub

Private Sub txtGrupo_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   txtConcep.SetFocus
End If
End Sub

Private Sub HabilitaConcepto(lActiva As Boolean)
cmdGrabaConcep.Visible = lActiva
cmdCancelaConcep.Visible = lActiva
fraDatConcep.Visible = lActiva
cmdNuevoConcep.Visible = Not lActiva
cmdModificaConcep.Visible = Not lActiva
cmdEliminaConcep.Visible = Not lActiva
cmdGenerar.Enabled = Not lActiva
fraRep.Enabled = Not lActiva
If lActiva Then
   lvConcep.Height = 1965
Else
   lvConcep.Height = 2445
End If
If lActiva Then
   txtClase.Enabled = lNuevo
   txtGrupo.Enabled = lNuevo
   txtConcep.Enabled = lNuevo
   If lNuevo Then
      txtConcepDesc = ""
      cboTpoCalculo.ListIndex = -1
   Else
      txtClase = lvConcep.SelectedItem.Text
      txtGrupo = lvConcep.SelectedItem.SubItems(1)
      txtConcep = lvConcep.SelectedItem.SubItems(2)
      txtConcepDesc = lvConcep.SelectedItem.SubItems(3)
      cboTpoCalculo.ListIndex = BuscaCombo(lvConcep.SelectedItem.SubItems(4), cboTpoCalculo)
      txtCtaCod = lvConcep.SelectedItem.SubItems(6)
      txtFormula = lvConcep.SelectedItem.SubItems(7)
   End If
End If
End Sub

Private Sub HabilitaRango(lActiva As Boolean)
cmdGrabaRango.Visible = lActiva
cmdCancelaRango.Visible = lActiva
fraDatRango.Visible = lActiva
cmdNuevoRango.Visible = Not lActiva
cmdModificaRango.Visible = Not lActiva
cmdEliminaRango.Visible = Not lActiva
cmdGenerar.Enabled = Not lActiva
fraRep.Enabled = Not lActiva
If lActiva Then
   lvRango.Height = 1965
Else
   lvRango.Height = 2445
End If
If lActiva Then
   txtRango.Enabled = lNuevo
   If lNuevo Then
      txtRangoDesc = ""
   Else
      txtRango = lvRango.SelectedItem.Text
      txtRangoDesc = lvRango.SelectedItem.SubItems(1)
      txtDesde = lvRango.SelectedItem.SubItems(2)
      txtHasta = lvRango.SelectedItem.SubItems(3)
   End If
End If
End Sub

Private Sub HabilitaCriterio(lActiva As Boolean)
cmdGrabaCriterio.Visible = lActiva
cmdCancelaCriterio.Visible = lActiva
cmdEditarCriterio.Visible = Not lActiva
fraRep.Enabled = Not lActiva
sTab.TabEnabled(0) = Not lActiva
sTab.TabEnabled(1) = Not lActiva
lvConcepto.Enabled = Not lActiva
cmdGenerar.Enabled = Not lActiva
End Sub

Private Function ValidaDatosRango() As Boolean
ValidaDatosRango = True
If txtRango = "" Then
   MsgBox "Falta ingresar Código de Rango ", vbInformation, "¡Aviso!"
   txtClase.SetFocus
   Exit Function
End If
If txtRangoDesc = "" Then
   MsgBox "Falta ingresar Descripción de Rango", vbInformation, "¡Aviso!"
   txtRangoDesc.SetFocus
   Exit Function
End If
If txtDesde = "" Then
   MsgBox "Falta ingresar Inicio de Rango", vbInformation, "¡Aviso!"
   txtDesde.SetFocus
   Exit Function
End If
If txtHasta = "" Then
   MsgBox "Falta ingresar Final de Rango", vbInformation, "¡Aviso!"
   txtHasta.SetFocus
   Exit Function
End If
ValidaDatosRango = True
End Function

Private Function ValidaDatosConcep() As Boolean
ValidaDatosConcep = True
If txtClase = "" Then
   MsgBox "Falta ingresar Clase de Concepto", vbInformation, "¡Aviso!"
   txtClase.SetFocus
   Exit Function
End If
If txtGrupo = "" Then
   MsgBox "Falta ingresar Grupo de Concepto", vbInformation, "¡Aviso!"
   txtGrupo.SetFocus
   Exit Function
End If
If txtConcep = "" Then
   MsgBox "Falta ingresar Código de Concepto", vbInformation, "¡Aviso!"
   txtConcep.SetFocus
   Exit Function
End If
ValidaDatosConcep = True
End Function

Private Sub CargaConceptos()
Dim lvItm As ListItem
lvConcep.ListItems.Clear
Set rs = oAnx.CargaConceptos(gsOpeCod)
Do While Not rs.EOF
   Set lvItm = lvConcep.ListItems.Add(, , rs!cCodClase)
   lvItm.SubItems(1) = rs!cCodGrp
   lvItm.SubItems(2) = rs!cCodConcep
   lvItm.SubItems(3) = rs!cDescrip
   lvItm.SubItems(4) = rs!cTpoCalculo
   lvItm.SubItems(5) = rs!cTpoCalculoDesc
   lvItm.SubItems(6) = rs!cCtaContCod
   lvItm.SubItems(7) = rs!cFormula
   rs.MoveNext
Loop
Set lvItm = Nothing
RSClose rs
End Sub
Private Sub CargaConceptosCriterio()
Dim lvItm As ListItem
lvConcepto.ListItems.Clear
Set rs = oAnx.CargaConceptoCriterio(gsOpeCod)
Do While Not rs.EOF
   Set lvItm = lvConcepto.ListItems.Add(, , rs!cDescrip)
   lvItm.SubItems(1) = rs!cCodClase
   lvItm.SubItems(2) = rs!cCodGrp
   lvItm.SubItems(3) = rs!cCodConcep
   lvItm.SubItems(4) = rs!cTpoCalculoDesc
   lvItm.SubItems(5) = rs!cTpoCalculo
   rs.MoveNext
Loop
Set lvItm = Nothing
RSClose rs
End Sub

Private Sub CargaRangos()
Dim lvItm As ListItem
lvRango.ListItems.Clear
Set rs = oAnx.CargaRangos(gsOpeCod)
Do While Not rs.EOF
   Set lvItm = lvRango.ListItems.Add(, , rs!cCodRango)
   lvItm.SubItems(1) = rs!cDescrip
   lvItm.SubItems(2) = rs!nDesde
   lvItm.SubItems(3) = rs!nHasta
   rs.MoveNext
Loop
RSClose rs
End Sub

Private Sub txtHasta_GotFocus()
fEnfoque txtHasta
End Sub

Private Sub txtHasta_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   cmdGrabaRango.SetFocus
End If
End Sub

Private Sub txtRango_GotFocus()
fEnfoque txtRango
End Sub

Private Sub txtRango_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   txtRangoDesc.SetFocus
End If
End Sub

Private Sub txtRangoDesc_GotFocus()
fEnfoque txtRangoDesc
End Sub

Private Sub txtRangoDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtDesde.SetFocus
End If
End Sub

Private Function ValidaDatos() As Boolean
ValidaDatos = False
If lbMes Then
    If Len(Trim(txtAnio.Text)) = 0 Then
        MsgBox "Ingrese Año de Proceso", vbInformation, "¡Aviso!"
        txtAnio.SetFocus
        Exit Function
    End If
    If CInt(txtAnio.Text) < 1950 Then
        MsgBox "Año no valido menor a 1950 ", vbInformation, "¡Aviso!"
        txtAnio.SetFocus
        Exit Function
    End If
Else
    If ValidaFecha(txtFechaAl) <> "" Then
      MsgBox "Fecha no válida...!", vbInformation, "¡Aviso!"
      txtFechaAl.SetFocus
      Exit Function
    End If
End If
If Val(txtTipCambio) = 0 Then
    MsgBox "No se definio Tipo de Cambio", vbInformation, "¡Aviso!"
    txtTipCambio.SetFocus
    Exit Function
End If
ValidaDatos = True
End Function

Private Sub CabeceraExcel(psMoneda As String, Optional pbAddMoneda As Boolean = True)
Dim sAnx As String
Dim sTit As String
Dim nPos As Integer
xlHoja1.PageSetup.Orientation = xlLandscape
xlHoja1.PageSetup.Zoom = 80
nPos = InStr(gsOpeDesc, ":")
If nPos > 0 Then
   sTit = UCase(Trim(Mid(gsOpeDesc, nPos + 1, 100)))
   nPos = InStr(sTit, ":")
   If nPos > 0 Then
      sAnx = Trim(Left(sTit, nPos - 1))
      sTit = Trim(Mid(sTit, nPos + 1, 100))
   Else
      sAnx = "ANEXO"
   End If
Else
   sTit = gsOpeDesc
End If
xlHoja1.Cells(1, 2) = sAnx
xlHoja1.Cells(3, 2) = "Empresa: " & gsNomCmac & "           Código: " & Left(gsCodAge, 3)
xlHoja1.Cells(5, 2) = sTit & IIf(pbAddMoneda, " EN " & IIf(psMoneda = "1", "MONEDA NACIONAL", "MONEDA EXTRANJERA"), "")
xlHoja1.Cells(7, 2) = "Al " & Mid(gdFecha, 1, 2) & " de " & Trim(cboMes) & " de " & txtAnio
xlHoja1.Cells(8, 2) = "( En Miles de Nuevos Soles y Mileda de Dólares Americanos )"

xlHoja1.Range("A1:O1").Merge
xlHoja1.Range("A3:O3").Merge
xlHoja1.Range("A5:O5").Merge
xlHoja1.Range("A7:O7").Merge
xlHoja1.Range("A8:O8").Merge
xlHoja1.Range("A1:O8").HorizontalAlignment = xlHAlignCenter

xlHoja1.Cells(10, 1) = "Movimiento de Efectivo Estimado Derivado de Operaciones Existentes"
xlHoja1.Cells(10, 2) = "REAL (7dias)"
xlHoja1.Cells(11, 2) = "M.N."
xlHoja1.Cells(11, 3) = "M.E."

xlHoja1.Range("A10:A11").HorizontalAlignment = xlHAlignCenter
xlHoja1.Range("B10:B11").HorizontalAlignment = xlHAlignCenter

xlHoja1.Range("A1:Q5").Font.Size = 11
xlHoja1.Range("A1:Q1").Font.Bold = True
xlHoja1.Range("A5:Q5").Font.Bold = True

xlHoja1.Range("A1:A1").ColumnWidth = 40
xlHoja1.Range("B1:B1").ColumnWidth = 12
xlHoja1.Range("C1:C1").ColumnWidth = 12
End Sub

Private Sub ImprimeRangos()
Set rs = oAnx.CargaRangos(gsOpeCod)
nColRango = 2
Do While Not rs.EOF
   nColRango = nColRango + 2
   xlHoja1.Cells(10, nColRango) = rs!cCodRango
   xlHoja1.Cells(11, nColRango) = rs!cDescrip
   xlHoja1.Cells(12, nColRango) = "M.N."
   xlHoja1.Cells(12, nColRango + 1) = "M.E."
   xlHoja1.Range(xlHoja1.Cells(10, nColRango), xlHoja1.Cells(10, nColRango + 1)).Merge
   xlHoja1.Range(xlHoja1.Cells(11, nColRango), xlHoja1.Cells(11, nColRango + 1)).Merge

   xlHoja1.Range(xlHoja1.Cells(11, nColRango), xlHoja1.Cells(11, nColRango)).EntireColumn.NumberFormat = "##,###,##0.00"
   xlHoja1.Range(xlHoja1.Cells(11, nColRango + 1), xlHoja1.Cells(11, nColRango)).EntireColumn.NumberFormat = "##,###,##0.00"
   xlHoja1.Range(xlHoja1.Cells(11, nColRango), xlHoja1.Cells(11, nColRango)).ColumnWidth = 14
   xlHoja1.Range(xlHoja1.Cells(11, nColRango + 1), xlHoja1.Cells(11, nColRango)).ColumnWidth = 14
   rs.MoveNext
Loop
nColRango = nColRango + 1
RSClose rs
xlHoja1.Cells(10, nColRango + 1) = "TOTAL(2)"
xlHoja1.Cells(11, nColRango + 1) = "M.N."
xlHoja1.Cells(11, nColRango + 2) = "M.E."
xlHoja1.Range(xlHoja1.Cells(10, 2), xlHoja1.Cells(11, nColRango + 2)).HorizontalAlignment = xlHAlignCenter
xlHoja1.Range(xlHoja1.Cells(10, 1), xlHoja1.Cells(11, nColRango + 2)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
xlHoja1.Range(xlHoja1.Cells(10, 1), xlHoja1.Cells(11, nColRango + 2)).Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range(xlHoja1.Cells(10, 2), xlHoja1.Cells(11, nColRango + 2)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
End Sub

Private Sub ImprimeConceptos(psMoneda As String)
Dim nFil    As Integer
Dim K       As Integer
Dim sCodAnt As String
If psMoneda = "1" Then
   Set rs = oAnx.CargaConceptos(gsOpeCod, IIf(OptOpc(0).value, " and cCodConcep = '00'", ""))
   nFil = 13
   sCodAnt = rs!cCodClase
   ReDim aPosicion(3, 1)
   nCont = 0
   Do While Not rs.EOF
      If Not rs!cCodClase = sCodAnt Then
         xlHoja1.Range(xlHoja1.Cells(nFil - 1, 1), xlHoja1.Cells(nFil - 1, nColRango + 2)).BorderAround xlContinuous
      End If
      xlHoja1.Cells(nFil, 1) = rs!cDescrip

      nCont = nCont + 1
      ReDim Preserve aPosicion(3, nCont)
      aPosicion(0, nCont) = rs!cTpoCalculo
      aPosicion(1, nCont) = rs!cCodClase & rs!cCodGrp & rs!cCodConcep
      aPosicion(2, nCont) = nFil
      aPosicion(3, nCont) = rs!cCtaContCod
      nFil = nFil + 1
      sCodAnt = rs!cCodClase
      rs.MoveNext
   Loop
   RSClose rs
End If
For K = 1 To nCont
   nFil = aPosicion(2, K)

   'Impresión de Información Real de los ultimos 7 dias
   If aPosicion(3, K) <> "" Then
      ImprimeMovCtaReal psMoneda, aPosicion(3, K), nFil, 1 + Val(psMoneda)
   End If

   Select Case aPosicion(0, K)
      Case gRiesgosTpoTasa:         ImprimeCriterioTasa psMoneda, nFil, aPosicion(1, K), aPosicion(3, K)
      Case gRiesgosSeries:          ImprimeCriterioSeries psMoneda, nFil, aPosicion(1, K), aPosicion(3, K)
      Case gRiesgosEncaje_BCR:      ImprimeCriterioEncaje psMoneda, nFil, aPosicion(1, K), aPosicion(3, K)
      Case gRiesgosTpoVencimiento:  ImprimeCriterioVencimiento psMoneda, nFil, aPosicion(1, K), aPosicion(3, K)
      Case gRiesgosPlaza_Cheque:    ImprimeCriterioCheque psMoneda, nFil, aPosicion(1, K), aPosicion(3, K)
      Case gRiesgosFecha:           ImprimeCriterioFecha psMoneda, nFil, aPosicion(1, K), aPosicion(3, K)
      Case gRiesgosFormula:         ImprimeCriterioFormula psMoneda, nFil, aPosicion(1, K), aPosicion(3, K)
      Case gRiesgosFormulaAcumula:  ImprimeCriterioFormula psMoneda, nFil, aPosicion(1, K), aPosicion(3, K), True
      Case gRiesgosTotales:         ImprimeCriterioTotales psMoneda, nFil, aPosicion(1, K), aPosicion(3, K), aPosicion(0, K)
   End Select
Next
nFil = nFil + 1
xlHoja1.Range(xlHoja1.Cells(nFil - 1, 1), xlHoja1.Cells(nFil - 1, nColRango + 2)).BorderAround xlContinuous
xlHoja1.Range(xlHoja1.Cells(13, 1), xlHoja1.Cells(nFil - 1, nColRango + 2)).BorderAround xlContinuous
xlHoja1.Range(xlHoja1.Cells(10, 1), xlHoja1.Cells(nFil - 1, nColRango + 2)).Borders(xlInsideVertical).LineStyle = xlContinuous

xlHoja1.Range(xlHoja1.Cells(1, nColRango + 1), xlHoja1.Cells(1, nColRango + 2)).EntireColumn.NumberFormat = "##,###,##0.00"
xlHoja1.Range(xlHoja1.Cells(1, nColRango + 1), xlHoja1.Cells(1, nColRango + 2)).ColumnWidth = 14
For N = 13 To nFil - 1
   If xlHoja1.Cells(N, 1) <> "" Then
      For K = 3 To nColRango - 1 Step 2
         If xlHoja1.Range(xlHoja1.Cells(N, nColRango + Val(psMoneda)), xlHoja1.Cells(N, nColRango + Val(psMoneda))).Formula = "" Then
            xlHoja1.Range(xlHoja1.Cells(N, nColRango + Val(psMoneda)), xlHoja1.Cells(N, nColRango + Val(psMoneda))).Formula = "="
         End If
         xlHoja1.Range(xlHoja1.Cells(N, nColRango + Val(psMoneda)), xlHoja1.Cells(N, nColRango + Val(psMoneda))).Formula = xlHoja1.Range(xlHoja1.Cells(N, nColRango + Val(psMoneda)), xlHoja1.Cells(N, nColRango + Val(psMoneda))).Formula & "+" & xlHoja1.Range(xlHoja1.Cells(N, K + Val(psMoneda)), xlHoja1.Cells(N, K + Val(psMoneda))).Address
      Next
   End If
Next
End Sub

Private Sub ImprimeMovCtaReal(psMoneda As String, psCtaCod As String, pnFil As Integer, pnCol As Integer)
Dim rsRiesgo As ADODB.Recordset
Dim oCon As New DConecta
oCon.AbreConexion
sSql = "SELECT ISNULL(SUM(CASE WHEN LEFT(cCtaContCod,1) = '1' THEN nMovImporte ELSE nMovImporte * -1 END),0) nImporte FROM Mov m JOIN MovCta mc ON mc.nMovNro = m.nMovNro " _
     & "WHERE LEFT(m.cMovNro,8) BETWEEN '" & Format(CDate(txtFechaAl) - 7, "yyyymmdd") & "' and '" & Format(CDate(txtFechaAl) - 1, "yyyymmdd") & "' " _
     & "      and m.nMovEstado = " & gMovEstContabMovContable & " and not m.nMovFlag in ('" & gMovFlagEliminado & "','" & gMovFlagModificado & "') and cCtaContCod LIKE '__" & psMoneda & "%' " _
     & "      and mc.cCtaContCod LIKE '" & psCtaCod & "%'"
Set rsRiesgo = oCon.CargaRecordSet(sSql)
If Not rsRiesgo.EOF Then
   If rsRiesgo!nImporte <> 0 Then
      xlHoja1.Cells(pnFil, pnCol) = rsRiesgo!nImporte
   End If
End If
oCon.CierraConexion
Set oCon = Nothing
End Sub

Private Sub ImprimeCriterioTasa(psMoneda As String, pnFil As Integer, psCodigo As String, psCtaCod As String)
Dim rsRiesgo As ADODB.Recordset
Dim nCol As Integer

   Set rsRiesgo = oAnx.CargaCriterioTasa(gsOpeCod, psMoneda, psCodigo, psCtaCod, Format(gdFecha, gsFormatoFecha))
   nCol = 3 + Val(psMoneda)
   Do While Not rsRiesgo.EOF
      Do While nCol <= nColRango
         If xlHoja1.Cells(10, nCol + 1 - Val(psMoneda)) = rsRiesgo!cCodRango Then
            xlHoja1.Cells(pnFil, nCol) = xlHoja1.Cells(pnFil, nCol) + Round(rsRiesgo!nValor / nTipCambio, 2)
            Exit Do
         End If
         nCol = nCol + 2
      Loop
      rsRiesgo.MoveNext
   Loop
   RSClose rsRiesgo
End Sub

Private Sub ImprimeCriterioSeries(psMoneda As String, pnFil As Integer, psCodigo As String, psCtaCod As String)
Dim rsRiesgo As ADODB.Recordset
Dim nCol As Integer
   Set rsRiesgo = oAnx.CargaCriterioSeries(gsOpeCod, psMoneda, psCodigo, psCtaCod, Format(gdFecha, gsFormatoFecha))
   nCol = 3 + Val(psMoneda)
   Do While Not rsRiesgo.EOF
      Do While nCol <= nColRango
         If xlHoja1.Cells(10, nCol + 1 - Val(psMoneda)) = rsRiesgo!cCodRango Then
            xlHoja1.Cells(pnFil, nCol) = xlHoja1.Cells(pnFil, nCol) + Round(rsRiesgo!nValor / nTipCambio, 2)
            Exit Do
         End If
         nCol = nCol + 2
      Loop
      rsRiesgo.MoveNext
   Loop
   RSClose rsRiesgo
End Sub

Private Sub ImprimeCriterioEncaje(psMoneda As String, pnFil As Integer, psCodigo As String, psCtaCod As String)
Dim rsRiesgo As ADODB.Recordset
Dim nCol As Integer
Dim nEncaje As Currency
Dim oSdo As New NCtasaldo
nEncaje = Round(oSdo.GetCtaSaldo("2[13]" & psMoneda & "%", Format(gdFecha, gsFormatoFecha)) / 100, 2)
Set oSdo = Nothing
   Set rsRiesgo = oAnx.CargaCriterioEncaje(gsOpeCod, psMoneda, psCodigo, psCtaCod, Format(gdFecha, gsFormatoFecha), nEncaje)
   nCol = 3 + Val(psMoneda)
   Do While Not rsRiesgo.EOF
      Do While nCol <= nColRango
         If xlHoja1.Cells(10, nCol + 1 - Val(psMoneda)) = rsRiesgo!cCodRango Then
            xlHoja1.Cells(pnFil, nCol) = xlHoja1.Cells(pnFil, nCol) + CCur(Round(rsRiesgo!nValor / nTipCambio, 2))
            Exit Do
         End If
         nCol = nCol + 2
      Loop
      rsRiesgo.MoveNext
   Loop
   RSClose rsRiesgo
End Sub

Private Sub ImprimeCriterioCheque(psMoneda As String, pnFil As Integer, psCodigo As String, psCtaCod As String)
Dim rsRiesgo As ADODB.Recordset
Dim nCol As Integer
Dim nEncaje As Currency
   Set rsRiesgo = oAnx.CargaCriterioCheque(gsOpeCod, psMoneda, psCodigo, Format(gdFecha, gsFormatoFecha))
   nCol = 3 + Val(psMoneda)
   Do While Not rsRiesgo.EOF
      Do While nCol <= nColRango
         If xlHoja1.Cells(10, nCol + 1 - Val(psMoneda)) = rsRiesgo!cCodRango Then
            xlHoja1.Cells(pnFil, nCol) = xlHoja1.Cells(pnFil, nCol) + Val(rsRiesgo!nValor)
            Exit Do
         End If
         nCol = nCol + 2
      Loop
      rsRiesgo.MoveNext
   Loop
   RSClose rsRiesgo
End Sub

Private Sub ImprimeCriterioFecha(psMoneda As String, pnFil As Integer, psCodigo As String, psCtaCod As String)
Dim rsRiesgo As ADODB.Recordset
Dim nCol As Integer
Dim nEncaje As Currency
Dim dFecha  As String
Set rsRiesgo = oAnx.CargaCriterios(gsOpeCod, psCodigo)
If Not rsRiesgo.EOF Then
   Do While Not rsRiesgo.EOF
     If Len(rsRiesgo!cValor) < 10 Then
        dFecha = CDate(rsRiesgo!cValor + "/" & txtAnio)
     Else
        dFecha = Format(rsRiesgo!cValor, gsFormatoFechaView)
     End If
      If dFecha >= gdFecha Then
         Exit Do
      End If
      rsRiesgo.MoveNext
   Loop
   If dFecha < gdFecha Then
      rsRiesgo.MoveFirst
      sSql = rsRiesgo!cValor
      dFecha = Format(sSql & "/" & txtAnio, "dd/mm/yyyy")
      dFecha = DateAdd("yyyy", 1, dFecha)
   End If
   Set rsRiesgo = oAnx.CargaCriterioFecha(gsOpeCod, psMoneda, psCodigo, psCtaCod, Format(dFecha, gsFormatoFecha), Format(gdFecha, gsFormatoFecha))
   nCol = 3 + Val(psMoneda)
   Do While Not rsRiesgo.EOF
      Do While nCol < nColRango
         If xlHoja1.Cells(10, nCol) = rsRiesgo!cCodRango Then
            xlHoja1.Cells(pnFil, nCol) = xlHoja1.Cells(pnFil, nCol) + Round(rsRiesgo!nValor / nTipCambio, 2)
            Exit Do
         End If
         nCol = nCol + 1
      Loop
      rsRiesgo.MoveNext
   Loop
End If
   RSClose rsRiesgo
End Sub

Private Sub ImprimeCriterioVencimiento(psMoneda As String, pnFil As Integer, psCodigo As String, psCtaCod As String)
Dim sCtaCod As String
Dim sObjCod As String
Dim rsRiesgo As ADODB.Recordset
Dim nCol    As Integer
Dim nUltCol As Integer
Dim nSaldoDif  As Currency
Dim nSaldo1411 As Currency
Dim lsTipoIF As CGTipoIF
sSql = ""
If psCtaCod = "" Then
    Exit Sub
End If
Select Case psCtaCod
   Case "11_301__03%", "11_303__03%"       'Plazo Fijo Bancos y CMACs
      If psCtaCod = "11_301__03%" Then
         sObjCod = "0101__03"
         lsTipoIF = gTpoIFBanco
      Else
         sObjCod = "0103__03"
         lsTipoIF = gTpoIFCmac
      End If
      Set rsRiesgo = oAnx.CargaCriterioVencBancos(gsOpeCod, psMoneda, psCodigo, psCtaCod, sObjCod, Format(gdFecha, gsFormatoFecha), Val(txtTipCambio), lsTipoIF)
   Case "14_1"             'Creditos Vigentes
      Set rsRiesgo = oAnx.CargaCriterioVencCredVig(gsOpeCod, psMoneda, psCtaCod, Format(gdFecha, gsFormatoFecha), Val(txtTipCambio))
      If Not rsRiesgo.EOF Then
         nSaldo1411 = rsRiesgo!nSaldo
         If rsRiesgo!nSaldo <> 0 Then
            Dim sProd As String
                sProd = ""
            Select Case psCtaCod
                Case "14_10313": sProd = "305"
                Case "14_10206__01": sProd = "201"
                Case "14_101": sProd = "101"
                Case "14_10206__02": sProd = "202"
                Case "N14_10313": sProd = "30[1234]"
                Case "14_104": sProd = "401"
            End Select
            nSaldoDif = nSaldo1411 - oAnx.GetSumaEstadRiesgos(psMoneda, Format(gdFecha, gsFormatoFecha), sProd)
      
            Set rsRiesgo = oAnx.CargaCriterioVencEstadCred(gsOpeCod, psMoneda, Format(gdFecha, gsFormatoFecha), "'V','P'", sProd)
         Else
            'Forzar a EOF
            rsRiesgo.MoveLast
            rsRiesgo.MoveNext
         End If
      End If
   Case "21_3"             'Obligaciones Cuentas a Plazo
      Set rsRiesgo = oAnx.CargaCriterioVencEstadCred(gsOpeCod, psMoneda, Format(gdFecha, gsFormatoFecha), "'F'")
   
   Case "24_[46]", "2[46]", "24", "26"          'Adeudados
      If Left(psCtaCod, 2) = "24" Or Left(psCtaCod, 3) = "2[4" Then
         nSaldoDif = GetSaldoCtaClase("24" & psMoneda & "8%", gdFecha, Val(psMoneda))
      End If
      If Left(psCtaCod, 2) = "26" Or Left(psCtaCod, 3) = "2[6" Then
         nSaldoDif = nSaldoDif + GetSaldoCtaClase("26" & psMoneda & "609%", gdFecha, Val(psMoneda))
      End If
      Set rsRiesgo = oAnx.CargaCriterioVencAdeuda(gsOpeCod, psMoneda, psCtaCod, "01____05", Format(gdFecha, gsFormatoFecha))
End Select
nUltCol = 0
If Not rsRiesgo.EOF Then
   nCol = 3 + Val(psMoneda)
   Do While Not rsRiesgo.EOF
      Do While nCol <= nColRango
         If xlHoja1.Cells(10, nCol + 1 - Val(psMoneda)) = rsRiesgo!cCodRango Then
            xlHoja1.Cells(pnFil, nCol) = Val(xlHoja1.Cells(pnFil, nCol)) + Val(rsRiesgo!nValor)
            nUltCol = nCol
            Exit Do
         End If
         nCol = nCol + 2
      Loop
      rsRiesgo.MoveNext
   Loop
   RSClose rsRiesgo
   If nSaldoDif <> 0 Then
      ImprimeCriterioMonto psMoneda, nSaldoDif, pnFil, psCodigo, psCtaCod
   End If
End If
End Sub

Private Sub ImprimeCriterioMonto(psMoneda As String, pnMonto As Currency, pnFil As Integer, psCodigo As String, psCtaCod As String)
Dim rsRiesgo As ADODB.Recordset
Dim nCol As Integer
   Set rsRiesgo = oAnx.CargaCriterioMonto(gsOpeCod, pnMonto, psCodigo)
   nCol = 3 + Val(psMoneda)
   Do While Not rsRiesgo.EOF
      Do While nCol <= nColRango
         If xlHoja1.Cells(10, nCol + 1 - Val(psMoneda)) = rsRiesgo!cCodRango Then
            xlHoja1.Cells(pnFil, nCol) = Val(xlHoja1.Cells(pnFil, nCol)) + Val(rsRiesgo!nValor)
            Exit Do
         End If
         nCol = nCol + 1
      Loop
      rsRiesgo.MoveNext
   Loop
   RSClose rsRiesgo
End Sub

Private Sub txtTipCambio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTipCambio, KeyAscii, 12, 4)
If KeyAscii = 13 Then
   cmdGenerar.SetFocus
End If
End Sub

Private Sub ImprimeCriterioFormula(psMoneda As String, pnFil As Integer, psCodigo As String, psCtaCod As String, Optional plAcumula As Boolean = False)
Dim rsTot As ADODB.Recordset
Dim sCadena As String
Dim sCodigo As String
Dim nPos    As Integer
Dim sCol As String
Dim nCol As Integer
Dim nEfectivo As Currency
Dim sSimbolo  As String

If psCtaCod = "" Then
   Exit Sub
End If
If psMoneda = "2" Then
   Exit Sub
End If
sCadena = Trim(psCtaCod)
Do While Len(sCadena) > 0
   nEfectivo = 0
   nPos = InStr(sCadena, ",")
   If nPos = 0 Then
      nPos = InStr(sCadena, "+")
   End If
   If nPos = 0 Then
      nPos = InStr(sCadena, "-")
   End If
   If nPos = 0 Then
      nPos = InStr(sCadena, "*")
   End If
   If nPos = 0 Then
      nPos = InStr(sCadena, "/")
   End If
   If nPos > 0 Then
      sSimbolo = Mid(sCadena, nPos, 1)
      sCodigo = Mid(sCadena, 1, nPos - 1)
      If sSimbolo = "," Then
         sSimbolo = "+"
      End If
'      If sCodigo = "[PE]" Then
'         nEfectivo = nVal(txtPatriEfec)
'         sCodigo = ""
'      End If
      sCadena = Mid(sCadena, nPos + 1, Len(sCadena))
   Else
      sCodigo = sCadena
      sCadena = ""
      sSimbolo = "+"
   End If
   If nEfectivo <> 0 Then
      For nCol = 3 To nColRango
         sCol = ExcelColumnaString(nCol)
         If xlHoja1.Range(sCol & pnFil).Formula = "" Then
            xlHoja1.Range(sCol & pnFil).Formula = "="
         End If
         xlHoja1.Range(sCol & pnFil).Formula = xlHoja1.Range(sCol & pnFil).Formula & sSimbolo & nEfectivo
      Next
   Else
      For nPos = 1 To nCont
         If aPosicion(1, nPos) = sCodigo Then
            For nCol = 2 To nColRango
               sCol = ExcelColumnaString(nCol)
               If xlHoja1.Range(sCol & pnFil).Formula = "" Then
                  xlHoja1.Range(sCol & pnFil).Formula = "="
               End If
               xlHoja1.Range(sCol & pnFil).Formula = xlHoja1.Range(sCol & pnFil).Formula & sSimbolo & sCol & aPosicion(2, nPos)
            Next
            Exit For
         End If
      Next
   End If
Loop
If plAcumula Then
   Dim sColAnt As String
   For nCol = 3 To nColRango
      sCol = ExcelColumnaString(nCol)
      sColAnt = ExcelColumnaString(nCol - 1)
      If xlHoja1.Range(sCol & pnFil).Formula = "" Then
         xlHoja1.Range(sCol & pnFil).Formula = "="
      End If
      xlHoja1.Range(sCol & pnFil).Formula = xlHoja1.Range(sCol & pnFil).Formula & Mid(xlHoja1.Range(sColAnt & pnFil).Formula, 2, Len(xlHoja1.Range(sColAnt & pnFil).Formula))
   Next
End If
End Sub

Private Sub ImprimeCriterioTotales(psMoneda As String, pnFil As Integer, psCodigo As String, psCtaCod As String, psTpoCalculo As String)
Dim rsTot As ADODB.Recordset
Dim sCadena As String
Dim sCodigo As String
Dim nPos    As Integer

If psCtaCod = "" Then
   Exit Sub
End If
sCadena = Trim(psCtaCod)
Do While Len(sCadena) > 0
   nPos = InStr(sCadena, ",")
   If nPos > 0 Then
      sCodigo = Mid(sCadena, 1, nPos - 1)
      sCadena = Mid(sCadena, nPos + 1, Len(sCadena))
   Else
      sCodigo = sCadena
      sCadena = ""
   End If
   Set rsTot = oAnx.CargaConceptos(gsOpeCod, , sCodigo)
   If Not rsTot.EOF Then
      Select Case rsTot!cTpoCalculo
         Case gRiesgosTpoTasa:         ImprimeCriterioTasa psMoneda, pnFil, sCodigo, rsTot!cCtaContCod
         Case gRiesgosSeries:          ImprimeCriterioSeries psMoneda, pnFil, sCodigo, rsTot!cCtaContCod
         Case gRiesgosEncaje_BCR:      ImprimeCriterioEncaje psMoneda, pnFil, sCodigo, rsTot!cCtaContCod
         Case gRiesgosTpoVencimiento:  ImprimeCriterioVencimiento psMoneda, pnFil, sCodigo, rsTot!cCtaContCod
         Case gRiesgosPlaza_Cheque:    ImprimeCriterioCheque psMoneda, pnFil, sCodigo, rsTot!cCtaContCod
         Case gRiesgosFecha:           ImprimeCriterioFecha psMoneda, pnFil, sCodigo, rsTot!cCtaContCod
         Case gRiesgosFormula:         ImprimeCriterioFormula psMoneda, pnFil, sCodigo, rsTot!cCtaContCod
         Case gRiesgosTotales:         ImprimeCriterioTotales psMoneda, pnFil, sCodigo, rsTot!cCtaContCod, rsTot!cTpoCalculo
      End Select
   End If
Loop
End Sub

