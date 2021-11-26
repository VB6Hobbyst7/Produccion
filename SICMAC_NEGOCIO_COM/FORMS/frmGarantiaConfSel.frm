VERSION 5.00
Begin VB.Form frmGarantiaConfSel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Items de Garantías"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   Icon            =   "frmGarantiaConfSel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   3960
      Width           =   1095
   End
   Begin SICMACT.FlexEdit feGarantConfig 
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   6588
      Cols0           =   4
      HighLight       =   1
      EncabezadosNombres=   "#-Objeto de la Garantía-Tipo de Bien-Cod"
      EncabezadosAnchos=   "0-3500-2000-0"
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
      ColumnasAEditar =   "X-X-X-X"
      ListaControles  =   "0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-L"
      FormatosEdit    =   "0-0-0-0"
      CantEntero      =   15
      TextArray0      =   "#"
      SelectionMode   =   1
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   3
      lbBuscaDuplicadoText=   -1  'True
      RowHeight0      =   300
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   3480
      Width           =   1095
   End
End
Attribute VB_Name = "frmGarantiaConfSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fnCodigo As Long

Public Function Inicio() As Long
    fnCodigo = 0
    Show 1
    Inicio = fnCodigo
End Function
Private Sub CargaDatos()
    Dim oGarant As New COMNCredito.NCOMGarantia
    Dim rsGarant As New ADODB.Recordset
    Dim i As Long
    
    Set rsGarant = oGarant.ObtenerConfigGarant()
    FormateaFlex feGarantConfig
    If Not (rsGarant.EOF And rsGarant.BOF) Then
        For i = 0 To rsGarant.RecordCount - 1
            feGarantConfig.AdicionaFila
            feGarantConfig.TextMatrix(i + 1, 1) = Trim(rsGarant!nObjGarantCod) & " " & Trim(rsGarant!cObjGarantDesc)
            feGarantConfig.TextMatrix(i + 1, 2) = Trim(rsGarant!TipoBien)
            feGarantConfig.TextMatrix(i + 1, 3) = Trim(rsGarant!nGarantConfigID)
            rsGarant.MoveNext
        Next i
        feGarantConfig.TopRow = 1
        feGarantConfig.row = 1
        EnfocaControl feGarantConfig
    End If
End Sub
Private Sub CmdAceptar_Click()
    If feGarantConfig.TextMatrix(1, 0) = "" Then
        MsgBox "No existe datos para seleccionar", vbInformation, "Aviso"
        Exit Sub
    End If
    fnCodigo = feGarantConfig.TextMatrix(feGarantConfig.row, 3)
    Unload Me
End Sub
Private Sub cmdSalir_Click()
    fnCodigo = 0
    Unload Me
End Sub
Private Sub feGarantConfig_DblClick()
    If feGarantConfig.TextMatrix(1, 0) <> "" Then
        CmdAceptar_Click
    End If
End Sub
Private Sub Form_Load()
    CargaDatos
End Sub

