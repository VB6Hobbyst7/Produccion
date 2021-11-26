VERSION 5.00
Begin VB.Form frmCapTasaIntHist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historial de Tasas de Ahorro"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9735
   Icon            =   "frmCapTasaIntHist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   4
      Top             =   3120
      Width           =   1035
   End
   Begin SICMACT.FlexEdit grdTasas 
      Height          =   2415
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   4260
      Cols0           =   10
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Monto Ini-Monto Fin-Plazo Ini-Plazo Fin-Ord?-Tasa Int-Cambio-Activa-Usuario"
      EncabezadosAnchos=   "300-1000-1200-900-900-500-800-2100-600-800"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
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
      ColumnasAEditar =   "X-1-2-3-4-5-6-X-8-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-0-0-0-0-0-0-0-4-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-R-R-R-R-C-R-C-C-C"
      FormatosEdit    =   "0-2-2-3-3-0-2-0-0-1"
      CantEntero      =   12
      CantDecimales   =   4
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "SubProducto :"
      Height          =   195
      Left            =   2280
      TabIndex        =   3
      Top             =   165
      Width           =   1020
   End
   Begin VB.Label lblSubProducto 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   2865
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "Moneda :"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   165
      Width           =   675
   End
   Begin VB.Label lblMoneda 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1185
   End
End
Attribute VB_Name = "frmCapTasaIntHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'** Nombre : frmCapTasaIntHist
'** Descripción : Formulario para visualizar historicamente las tasas pasivas según TI-ERS009-2014
'** Creación : JUEZ, 20140220 09:00:00 AM
'*****************************************************************************************************

Option Explicit

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Public Sub Inicia(ByVal nMoneda As Integer, ByVal psSubProducto As String, ByVal rs As ADODB.Recordset)
    lblMoneda.Caption = IIf(nMoneda = 1, "SOLES", "DÓLARES")
    Me.lblSubProducto.Caption = psSubProducto
    Set grdTasas.Recordset = rs
    Me.Show 1
End Sub

