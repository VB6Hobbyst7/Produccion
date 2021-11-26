VERSION 5.00
Begin VB.Form frmCapServListadoInst 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   Icon            =   "frmCapServListadoInst.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5025
      TabIndex        =   3
      Top             =   4155
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3660
      TabIndex        =   2
      Top             =   4155
      Width           =   975
   End
   Begin VB.Frame fraConvenio 
      Caption         =   "Instituciones Convenio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3900
      Left            =   165
      TabIndex        =   0
      Top             =   135
      Width           =   8955
      Begin SICMACT.FlexEdit grdConvenio 
         Height          =   3495
         Left            =   120
         TabIndex        =   1
         Top             =   255
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   6165
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Codigo-Nombre-Tipo-nConvCod-Flag"
         EncabezadosAnchos=   "300-1500-3500-2800-0-0"
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
         ColumnasAEditar =   "X-1-X-3-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-1-0-3-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-C-C"
         FormatosEdit    =   "0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         TipoBusqueda    =   3
         ColWidth0       =   300
         RowHeight0      =   300
         TipoBusPersona  =   1
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmCapServListadoInst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
