VERSION 5.00
Begin VB.Form frmCuentaAhorrosDesembolsoTerceros 
   Caption         =   "Cuenta de Ahorros para Desembolso (Terceros)"
   ClientHeight    =   2070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10335
   Icon            =   "frmCuentaAhorrosDesembolsoTerceros.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   10335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   8880
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin SICMACT.FlexEdit feTerceros 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   2355
      Cols0           =   6
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-DOI-Nombre-Relación-Cuenta-CodPersona"
      EncabezadosAnchos=   "500-2000-5000-1600-0-0"
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
      EncabezadosAlineacion=   "C-C-L-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   495
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmCuentaAhorrosDesembolsoTerceros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub Inicio(pCtaAhorro As String)
If CargarGrid(pCtaAhorro) Then
    Me.Show 1
Else
    Exit Sub
End If
End Sub
Private Function CargarGrid(pCtaAhorro As String) As Boolean
Dim oPersona As COMDPersona.DCOMPersonas
Dim rsPersona As ADODB.Recordset
Dim Fila As Integer
Set oPersona = New COMDPersona.DCOMPersonas
Fila = 0
Call oPersona.RecuperarDatosPersonaDesembolsoTercero(pCtaAhorro, rsPersona)
 
If rsPersona.BOF Or rsPersona.EOF Then
    MsgBox "Cuenta de Ahorro no Existe", vbInformation
    CargarGrid = False
    Exit Function
Else
    Do While Not rsPersona.EOF
    Fila = Fila + 1
    Me.feTerceros.AdicionaFila
    Me.feTerceros.TextMatrix(Fila, 1) = rsPersona!cPersIDnro
    Me.feTerceros.TextMatrix(Fila, 2) = rsPersona!cPersNombre
    Me.feTerceros.TextMatrix(Fila, 3) = rsPersona!cConsDescripcion
    rsPersona.MoveNext
    Loop
    CargarGrid = True
End If

Set rsPersona = Nothing
End Function
Private Sub BtnCerrar_Click()
Unload Me
End Sub

