VERSION 5.00
Begin VB.Form frmDocRecParam 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5610
   Icon            =   "frmDocRecParam.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   350
      Left            =   4590
      TabIndex        =   2
      Top             =   4455
      Width           =   900
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   350
      Left            =   3555
      TabIndex        =   1
      Top             =   4455
      Width           =   945
   End
   Begin VB.Frame fraParametro 
      Caption         =   "Parámetros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   4290
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5415
      Begin SICMACT.FlexEdit grdParam 
         Height          =   3885
         Left            =   90
         TabIndex        =   3
         Top             =   270
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   6853
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Moneda-Tipo-Plaza-Días Min-Mod"
         EncabezadosAnchos=   "350-1000-1200-1200-1000-0"
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
         ColumnasAEditar =   "X-X-X-X-4-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-R-C"
         FormatosEdit    =   "0-0-0-0-3-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmDocRecParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGrabar_Click()
If MsgBox("¿Desea Actualizar los parámetros del Cheque?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion
    Dim rsChq As ADODB.Recordset
    Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion
    Set rsChq = grdParam.GetRsNew()
    oCap.ActualizaDocRecParametro rsChq
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    Set oCap = Nothing
End If
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion
Me.Caption = "Cheques - Parámetros"
Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion
Set grdParam.Recordset = oCap.GetDocRecParametro()
Set oCap = Nothing
End Sub

Private Sub grdParam_OnCellChange(pnRow As Long, pnCol As Long)
    grdParam.TextMatrix(pnRow, 5) = "M"
End Sub

