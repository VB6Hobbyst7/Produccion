VERSION 5.00
Begin VB.Form frmColPMiembrosRetasacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Miembros del Comité"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   Icon            =   "frmColPMiembrosRetasacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin FomsNegocio.FlexEdit FEMiembrosComite 
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3836
      Cols0           =   3
      ScrollBars      =   2
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Nombre-Rol"
      EncabezadosAnchos=   "400-3200-2500"
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
      ColumnasAEditar =   "X-X-X"
      ListaControles  =   "0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C"
      FormatosEdit    =   "0-0-0"
      TextArray0      =   "#"
      SelectionMode   =   1
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      AutoAdd         =   -1  'True
      ForeColorFixed  =   -2147483630
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
End
Attribute VB_Name = "frmColPMiembrosRetasacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmColPMiembrosRetasacion
'** Descripción : Formulario que muestra los miembros de la retasacion
'** Creación    : TORE, ERS054-2017
'**********************************************************************************************
Option Explicit

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Public Sub Inicio(ByVal oRSMiembros As ADODB.Recordset)
    
    If Not oRSMiembros.EOF And Not oRSMiembros.BOF Then
        If oRSMiembros.RecordCount = 0 Then
            MsgBox "No hay datos de los miembros del comité de retasación", vbInformation, "Advertencia"
        Else
            Set FEMiembrosComite.Recordset = oRSMiembros
            Me.Show 1
        End If
    Else
        MsgBox "No hay datos de los miembros del comité de retasación", vbCritical, "Advertencia"
    End If
End Sub

