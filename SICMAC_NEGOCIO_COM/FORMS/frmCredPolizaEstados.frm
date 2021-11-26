VERSION 5.00
Begin VB.Form frmCredPolizaEstados 
   Caption         =   "Modificacion de Estados de Poliza"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6780
   Icon            =   "frmCredPolizaEstados.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3795
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   6255
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
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
         Left            =   3840
         TabIndex        =   4
         Top             =   120
         Width           =   1005
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
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
         Left            =   1320
         TabIndex        =   3
         Top             =   120
         Width           =   1005
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin SICMACT.FlexEdit FePolizas 
         Height          =   2790
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   4921
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-N° Poliza-Codigo-Tipo Poliza-F. Registro-Garantia"
         EncabezadosAnchos=   "400-1000-0-2050-1200-1500"
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
         ColumnasAEditar =   "X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-L-C-C"
         FormatosEdit    =   "0-0-0-0-5-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         lbPuntero       =   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmCredPolizaEstados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub inicia(pcNumGarant)
    LlenaGrillas pcNumGarant
    Me.Show 1
End Sub

Public Sub LlenaGrillas(ByVal pcNumGarant As String)
Dim oPol As COMDCredito.DCOMPoliza
Dim rs As ADODB.Recordset
Set oPol = New COMDCredito.DCOMPoliza

Set rs = oPol.RecuperaEstadoPolizasListado(pcNumGarant)

If rs.EOF Then MsgBox "No se encontraron resultados", vbInformation, "Mensaje"

FePolizas.Clear
FePolizas.FormaCabecera
FePolizas.Rows = 2
FePolizas.rsFlex = rs
'FePolizas.SetFocus
Set oPol = Nothing
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdEditar_Click()
Dim oPol As COMDCredito.DCOMPoliza
Set oPol = New COMDCredito.DCOMPoliza

oPol.ActualizaEstadoPolizasListado Me.FePolizas.TextMatrix(FePolizas.row, 5)
Set oPol = Nothing
MsgBox "Estados de Polizas de Garantia Actualizadas, Cierre el Formulario de Garantias para completar la Actualizacion", vbInformation, "Aviso"
LlenaGrillas Me.FePolizas.TextMatrix(FePolizas.row, 5)
End Sub
