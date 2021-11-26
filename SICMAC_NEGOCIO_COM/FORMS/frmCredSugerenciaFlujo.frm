VERSION 5.00
Begin VB.Form frmCredSugerenciaFlujo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Datos de Flujo"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3600
   Icon            =   "frmCredSugerenciaFlujo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtInflacion 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1875
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   600
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fluctuaciones de Ventas"
      Height          =   3690
      Left            =   75
      TabIndex        =   4
      Top             =   1050
      Width           =   3465
      Begin SICMACT.FlexEdit FeMes 
         Height          =   3240
         Left            =   75
         TabIndex        =   5
         Top             =   300
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   5715
         Cols0           =   3
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-MES-Porcen. (%)"
         EncabezadosAnchos=   "400-1200-1200"
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
         ColumnasAEditar =   "X-X-2"
         ListaControles  =   "0-0-0"
         EncabezadosAlineacion=   "C-L-R"
         FormatosEdit    =   "0-0-2"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   1725
      TabIndex        =   3
      Top             =   4800
      Width           =   1140
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   390
      Left            =   600
      TabIndex        =   2
      Top             =   4800
      Width           =   1140
   End
   Begin VB.TextBox txtVarMEnTC 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1875
      TabIndex        =   1
      Text            =   "0.00"
      Top             =   188
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "Inflación                   :"
      Height          =   240
      Left            =   225
      TabIndex        =   7
      Top             =   630
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Variación Men. T.C. :"
      Height          =   240
      Left            =   225
      TabIndex        =   0
      Top             =   225
      Width           =   1590
   End
End
Attribute VB_Name = "frmCredSugerenciaFlujo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public nVarMensualTC As Double
Public MatMensualPorc As Variant
Public nInflacion As Double
Dim i As Integer

Public Sub Inicio(ByVal pMatCalenFechas As Variant)

    With FeMes
        .Clear
        .FormaCabecera
        .Rows = 2
        For i = 1 To UBound(pMatCalenFechas)
            .AdicionaFila
            .TextMatrix(i, 1) = Format(pMatCalenFechas(i - 1), "mmm-yyyy")
            .TextMatrix(i, 2) = Format(0, "#0.00")
        Next i
    End With
Me.Show 1
End Sub


Private Sub CmdAceptar_Click()
    If CDbl(txtVarMEnTC.Text) = 0 Then
        MsgBox "Debe ingresar la variacion del TC", vbInformation, "Mensaje"
        Exit Sub
    End If
    If CDbl(txtInflacion.Text) = 0 Then
        MsgBox "Debe ingresar el valor de inflación", vbInformation, "Mensaje"
        Exit Sub
    End If

    With FeMes
        For i = 0 To .Rows - 2
            If .TextMatrix(i + 1, 2) = "" Then
                MsgBox "Debe ingresar la fluctuación de ventas para cada mes", vbInformation, "Mensaje"
                Exit Sub
            End If
        Next i
    End With
    
    nVarMensualTC = CDbl(txtVarMEnTC.Text)
    nInflacion = CDbl(txtInflacion.Text)
    ReDim MatMensualPorc(FeMes.Rows - 1)
    For i = 0 To FeMes.Rows - 1
        MatMensualPorc(i) = FeMes.TextMatrix(i, 2)
    Next i
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    nVarMensualTC = 0
    ReDim MatMensualPorc(0)
    Unload Me
End Sub

Private Sub FeMes_RowColChange()
    FeMes.Col = 2
End Sub

Private Sub Form_Load()
    Call CentraForm(Me)
    fEnfoque txtVarMEnTC
End Sub

Private Sub txtInflacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FeMes.SetFocus
        FeMes.Row = 1
        FeMes.Col = 2
    End If
End Sub

Private Sub txtVarMEnTC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtInflacion.SetFocus
    KeyAscii = NumerosDecimales(txtVarMEnTC, KeyAscii)
End Sub

Private Sub txtVarMEnTC_LostFocus()
    If Trim(txtVarMEnTC.Text) = "" Then
        txtVarMEnTC.Text = "0.00"
    Else
        txtVarMEnTC.Text = Format(txtVarMEnTC.Text, "#0.00")
    End If
    fEnfoque txtInflacion
End Sub
