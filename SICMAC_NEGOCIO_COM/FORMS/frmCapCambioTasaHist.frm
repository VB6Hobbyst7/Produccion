VERSION 5.00
Begin VB.Form frmCapCambioTasaHist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CTS - Tasa histórica"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   Icon            =   "frmCapCambioTasaHist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
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
      Left            =   3960
      TabIndex        =   1
      Top             =   2400
      Width           =   1095
   End
   Begin SICMACT.FlexEdit feTasas 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   3836
      Cols0           =   5
      HighLight       =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Fecha-Movimiento-Tasa-Usuario"
      EncabezadosAnchos=   "300-1100-1400-900-900"
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
      ColumnasAEditar =   "X-X-X-X-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L-C-C"
      FormatosEdit    =   "0-0-0-0-0"
      CantEntero      =   12
      CantDecimales   =   4
      TextArray0      =   "#"
      SelectionMode   =   1
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmCapCambioTasaHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'** Nombre : frmCapCambioTasaHist
'** Descripción : Formulario para visualizar los cambios tasas desde la apertura según TI-ERS013-2014
'** Creación : JUEZ, 20140305 10:30:00 AM
'*****************************************************************************************************

Option Explicit

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Public Sub Inicio(ByVal psCtaCod As String)
    Dim R As ADODB.Recordset
    Dim oDCapGen As COMDCaptaGenerales.DCOMCaptaGenerales
    Set oDCapGen = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set R = oDCapGen.RecuperaDatosTasaHistorica(psCtaCod)
    If Not R.EOF And Not R.BOF Then
        Set feTasas.Recordset = R
    End If
    Me.Show 1
End Sub
