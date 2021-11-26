VERSION 5.00
Begin VB.Form frmColPAdjudicaLista 
   Caption         =   "Contratos para Adjudicar"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   Icon            =   "frmColPAdjudicaLista.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   11070
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   5640
      Width           =   975
   End
   Begin SICMACT.FlexEdit FeAdj 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   9128
      Cols0           =   7
      HighLight       =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
      EncabezadosNombres=   "Nº-OK-Contrato-Cliente-Fec. Venc.-Préstamo-Saldo Cap."
      EncabezadosAnchos=   "400-400-1800-4000-1200-1200-1200"
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
      ColumnasAEditar =   "X-1-X-X-X-X-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-4-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-L-L-R-R"
      FormatosEdit    =   "0-0-0-0-5-2-2"
      AvanceCeldas    =   1
      TextArray0      =   "Nº"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmColPAdjudicaLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCodAge As String
Public cNumProceso As String

Public Sub Inicio(ByVal pAge As String, ByVal pNumProceso As String)
    cCodAge = pAge
    cNumProceso = pNumProceso
    CargaDatos
    Me.Show 1
End Sub

Private Sub CargaDatos()
    Dim oAdj As COMDColocPig.DCOMColPContrato
    Dim rs As adodb.Recordset
    Set oAdj = New COMDColocPig.DCOMColPContrato

    Set rs = oAdj.RecuperaCredExcluirAdj(cCodAge, cNumProceso)
    
    If rs.EOF Then MsgBox "No se encontraron datos.", vbInformation, "Mensaje"
    
    FeAdj.Clear
    FeAdj.FormaCabecera
    FeAdj.Rows = 2
    FeAdj.rsFlex = rs

    Set oAdj = Nothing
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGrabar_Click()

If MsgBox("¿Está seguro de Guardar los cambios realizados? ", vbInformation + vbYesNo, "Aviso") = vbNo Then
    Exit Sub
End If

Dim i As Integer
Dim loExcluyeAdj As COMNColoCPig.NCOMColPContrato

Set loExcluyeAdj = New COMNColoCPig.NCOMColPContrato

    For i = 1 To FeAdj.Rows - 1
        If FeAdj.TextMatrix(i, 1) <> "." Then

            Call loExcluyeAdj.ModificarEstadoAdj(CStr(FeAdj.TextMatrix(i, 2)), cNumProceso)

        End If
    Next i
    
    Set loExcluyeAdj = Nothing
    
    CargaDatos
    
End Sub
