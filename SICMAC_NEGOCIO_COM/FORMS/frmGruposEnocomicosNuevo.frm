VERSION 5.00
Begin VB.Form frmGruposEnocomicosNuevo 
   Caption         =   "Grupo Economico - Nuevo"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7890
   Icon            =   "frmGruposEnocomicosNuevo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   5760
      TabIndex        =   2
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   6840
      TabIndex        =   1
      Top             =   3960
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Grupos"
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin SICMACT.FlexEdit FEGrupo 
         Height          =   3255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   5741
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "-Codigo-Grupo Económico-Activar"
         EncabezadosAnchos=   "400-1000-4000-800"
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
         ColumnasAEditar =   "X-X-2-3"
         ListaControles  =   "0-0-0-4"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-C"
         FormatosEdit    =   "0-0-0-0"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmGruposEnocomicosNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FEGrupoNoMoverdeFila As Integer
Dim lnGrupo As Integer

Private Sub cmdGrabar_Click()
Dim i As Integer
Dim oGrup As COMDpersona.DCOMGrupoE
Dim nRetorno As Integer
Set oGrup = New COMDpersona.DCOMGrupoE
With FEGrupo
    For i = 1 To .Rows - 1
    nRetorno = oGrup.ActualizarGrupoEconomico(CInt(Trim(.TextMatrix(i, 1))), Trim(.TextMatrix(i, 2)), IIf(Trim(.TextMatrix(i, 3)) = ".", 1, 0))
    Next
 MsgBox "Datos se guardaron correctamente", vbApplicationModal
Call cmdSalir_Click
End With
End Sub

Private Sub cmdNuevo_Click()
FEGrupo.AdicionaFila
FEGrupoNoMoverdeFila = FEGrupo.Rows - 1
FEGrupo.lbEditarFlex = True
FEGrupo.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Public Sub Iniciar()
    Call ListarGrupos
    Show 1
End Sub
Private Sub ListarGrupos()
    Dim nRetorno As Integer
    Dim i As Integer
    Dim oGrup As COMDpersona.DCOMGrupoE
    Dim rs As ADODB.Recordset
    With FEGrupo
        For i = 1 To .Rows - 1
           FEGrupo.EliminaFila (1)
        Next
    End With
    
    Set oGrup = New COMDpersona.DCOMGrupoE
    Set rs = oGrup.ListarGrupoEconomico(2)
    i = 0
    If rs.EOF Or rs.BOF Then
        nRetorno = 0
    Else
        Do Until rs.EOF
        i = i + 1
            FEGrupo.AdicionaFila
            FEGrupo.TextMatrix(i, 1) = rs!nConsValor
            FEGrupo.TextMatrix(i, 2) = rs!cConsDescripcion
            If rs!nGrupoEstado = 1 Then
            FEGrupo.TextMatrix(i, 3) = 1
            Else
            End If
            rs.MoveNext
        Loop
        nRetorno = 1
    End If
    Set oGrup = Nothing
    rs.Close
End Sub


Private Sub FEGrupo_OnCellChange(pnRow As Long, pnCol As Long)
If FEGrupo.Col = 2 Then
    If FEGrupo.TextMatrix(FEGrupo.Row, 2) <> "" And FEGrupo.TextMatrix(FEGrupo.Row, 1) = "" Then
        FEGrupo.TextMatrix(FEGrupo.Row, 1) = -1
    End If
End If
End Sub

