VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCapTarifarioProgramacionVer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver Programación"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11850
   Icon            =   "frmCapTarifarioProgramacionVer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraFiltros 
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   90
      TabIndex        =   4
      Top             =   45
      Width           =   11670
      Begin VB.CommandButton btnBuscar 
         Caption         =   "Buscar Programacion"
         Height          =   300
         Left            =   6210
         TabIndex        =   9
         Top             =   270
         Width           =   1860
      End
      Begin VB.ComboBox cbTipo 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   270
         Width           =   2310
      End
      Begin MSComCtl2.DTPicker dpFecha 
         Height          =   300
         Left            =   4635
         TabIndex        =   7
         Top             =   270
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   123404289
         CurrentDate     =   42481
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha de Vigencia:"
         Height          =   240
         Left            =   3150
         TabIndex        =   8
         Top             =   315
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo:"
         Height          =   240
         Left            =   225
         TabIndex        =   6
         Top             =   315
         Width           =   465
      End
   End
   Begin VB.CommandButton btnCerrar 
      Caption         =   "Cerrar"
      Height          =   300
      Left            =   10800
      TabIndex        =   3
      Top             =   6885
      Width           =   960
   End
   Begin VB.CommandButton btnRegistrarProgramacion 
      Caption         =   "Extornar Programacion"
      Height          =   300
      Left            =   90
      TabIndex        =   2
      Top             =   6885
      Width           =   1950
   End
   Begin VB.Frame fraProgramacion 
      Caption         =   "Programacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6000
      Left            =   90
      TabIndex        =   0
      Top             =   810
      Width           =   11670
      Begin SICMACT.FlexEdit grdVersiones 
         Height          =   5640
         Left            =   90
         TabIndex        =   1
         Top             =   225
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   9948
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Id-tmpSel--Producto-Sub Producto-Personeria-Grupo-Version-nLista-tmp"
         EncabezadosAnchos=   "0-0-500-1800-1800-1800-700-4500-0-0"
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
         ColumnasAEditar =   "X-X-2-X-X-X-X-7-X-X"
         ListaControles  =   "0-0-4-0-0-0-0-3-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-L-L-C-L-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "Id"
         SelectionMode   =   1
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmCapTarifarioProgramacionVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************************
'* NOMBRE         : frmCapTarifarioProgramacion
'* DESCRIPCION    : Proyecto - Tarifario Versionado - Programacion de los Tarifarios
'* CREACION       : RIRO, 20160425 10:00 AM
'************************************************************************************************************

Private Sub btnBuscar_Click()

End Sub
Private Sub btnCerrar_Click()
    If MsgBox("¿Desea salir del formulario Ver Programacion?", vbInformation + vbYesNo + vbDefaultButton1, "Aviso") = vbYes Then
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    btnCerrar.Cancel = True
End Sub
