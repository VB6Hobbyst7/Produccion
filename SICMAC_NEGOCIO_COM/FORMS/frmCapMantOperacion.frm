VERSION 5.00
Begin VB.Form frmCapMantOperacion 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8640
   Icon            =   "frmCapMantOperacion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   5880
      Width           =   1035
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   5880
      Width           =   1035
   End
   Begin VB.Frame fraOperacion 
      Caption         =   "Operaciones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   5715
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   8475
      Begin SICMACT.FlexEdit grdOperacion 
         Height          =   5295
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   9340
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Código-Operación-Producto-Tipo-Estad-Mod"
         EncabezadosAnchos=   "350-800-3000-1200-1800-700-0"
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-4-5-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-3-4-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-L-L-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmCapMantOperacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGrabar_Click()
Dim rsOpe As ADODB.Recordset

Set rsOpe = grdOperacion.GetRsNew()
If MsgBox("¿Desea Grabar la Información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion
    Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion
    oCap.ActualizaCaptacionOperacion rsOpe
    Set oCap = Nothing
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()

Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim oGen As COMDConstSistema.DCOMGeneral

Me.Caption = "Matenimiento Operacion Captaciones"
Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion
Set grdOperacion.Recordset = oCap.GetCapMovTipo()
Set oCap = Nothing

Set oGen = New COMDConstSistema.DCOMGeneral
grdOperacion.CargaCombo oGen.GetConstante(gCaptacMovTipo)
Set oGen = Nothing
End Sub

Private Sub grdOperacion_OnCellChange(pnRow As Long, pnCol As Long)
    grdOperacion.TextMatrix(pnRow, 6) = "M"
End Sub
