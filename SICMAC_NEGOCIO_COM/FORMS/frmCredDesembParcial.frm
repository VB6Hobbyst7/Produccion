VERSION 5.00
Begin VB.Form frmCredDesembParcial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Desembolsos Parciales"
   ClientHeight    =   2625
   ClientLeft      =   2820
   ClientTop       =   2775
   ClientWidth     =   5790
   Icon            =   "frmCredDesembParcial.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2460
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   5625
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   420
         Left            =   3600
         TabIndex        =   4
         Top             =   1815
         Width           =   1710
      End
      Begin VB.CommandButton CmdEliminarDesemb 
         Caption         =   "&Eliminar"
         Height          =   420
         Left            =   3600
         TabIndex        =   3
         Top             =   765
         Width           =   1710
      End
      Begin VB.CommandButton CmdNuevoDesemb 
         Caption         =   "&Nuevo"
         Height          =   420
         Left            =   3600
         TabIndex        =   2
         Top             =   315
         Width           =   1710
      End
      Begin SICMACT.FlexEdit FEDesPar 
         Height          =   2100
         Left            =   150
         TabIndex        =   1
         Top             =   195
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   3704
         Cols0           =   3
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "-Fecha-Monto"
         EncabezadosAnchos=   "350-1200-1200"
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-2"
         ListaControles  =   "0-2-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-R"
         FormatosEdit    =   "0-0-2"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmCredDesembParcial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*****************************************************************************************
''***     Rutina:           frmCredDesembParcial
''***     Descripcion:      Permite el Ingreso de Multiples desembolsos para un credito
''***     Creado por:        NSSE
''***     Maquina:           07SIST_08
''***     Fecha-Tiempo:         18/06/2001 12:32:25 PM
''***     Ultima Modificacion: Lo Ultimo que se Modifico
''******************************************************************************************
Option Explicit
Dim MatCalend() As String
Dim dFecDesMin As Date
'
'Se hizo Optional (pMatDesPar) para el caso de la Simulacion del Calendario de cuota libre
Public Function Inicio(ByVal pdFecDesMin As Date, Optional ByVal pMatDesPar As Variant = "") As Variant
Dim i As Integer
    'On Error Resume Next
    'If Not pMatDesPar Is Nothing Then
    'If pMatDesPar <> "" Then
    If IsArray(pMatDesPar) Then
        If UBound(pMatDesPar) > 0 Then
            For i = 0 To UBound(pMatDesPar) - 1
                FEDesPar.AdicionaFila
                FEDesPar.TextMatrix(i + 1, 1) = pMatDesPar(i, 0)
                FEDesPar.TextMatrix(i + 1, 2) = pMatDesPar(i, 1)
            Next i
        End If
    End If
    'End If
    dFecDesMin = pdFecDesMin
    FEDesPar.lbEditarFlex = False
    Me.Show 1
    Inicio = MatCalend
End Function
'
'Private Function ValidaDatos() As Boolean
'Dim i As Integer
'Dim dFecTemp As Date
'    dFecTemp = dFecDesMin
'    ValidaDatos = True
'
'    If Trim(FEDesPar.TextMatrix(1, 0)) = "" Then
'        Exit Function
'    End If
'    For i = 1 To FEDesPar.Rows - 1
'        If ValidaFecha(FEDesPar.TextMatrix(i, 1)) <> "" Then
'            ValidaDatos = False
'            MsgBox ValidaFecha(FEDesPar.TextMatrix(i, 1)), vbInformation, "Aviso"
'            FEDesPar.row = i
'            FEDesPar.Col = 1
'            FEDesPar.SetFocus
'            Exit Function
'        End If
'        If Trim(FEDesPar.TextMatrix(i, 2)) = "" Then
'            FEDesPar.TextMatrix(i, 2) = "0.00"
'        End If
'        If CDbl(FEDesPar.TextMatrix(i, 2)) <= 0 Then
'            ValidaDatos = False
'            MsgBox "Monto de Desembolso debe ser mayor que Cero", vbInformation, "Aviso"
'            FEDesPar.row = i
'            FEDesPar.Col = 2
'            FEDesPar.SetFocus
'            Exit Function
'        End If
'        If CDate(FEDesPar.TextMatrix(i, 1)) < dFecTemp Then
'            ValidaDatos = False
'            MsgBox "Fecha de Desembolso No puede ser Menor o Igual que la Fecha de Desembolso Anterior", vbInformation, "Aviso"
'            FEDesPar.row = i
'            FEDesPar.Col = 1
'            FEDesPar.SetFocus
'            Exit Function
'        End If
'        dFecTemp = CDate(FEDesPar.TextMatrix(i, 1))
'    Next i
'End Function
'Private Sub CmdEliminarDesemb_Click()
'    If MsgBox("Se va a Eliminar el Desembolso de la Fecha : " & FEDesPar.TextMatrix(FEDesPar.row, 1) & ", Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
'        Call FEDesPar.EliminaFila(FEDesPar.row)
'    End If
'End Sub
'
'Private Sub CmdNuevoDesemb_Click()
'    FEDesPar.lbEditarFlex = True
'    FEDesPar.AdicionaFila
'    FEDesPar.SetFocus
'End Sub
'
'Private Sub cmdSalir_Click()
'    Unload Me
'End Sub
'
'Private Sub Form_Load()
'    CentraForm Me
'    FEDesPar.lbEditarFlex = True
'    ReDim MatCalend(0, 0)
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'Dim i As Integer
'    If Trim(FEDesPar.TextMatrix(1, 0)) = "" Then
'        ReDim MatCalend(0, 0)
'        Exit Sub
'    End If
'    ReDim MatCalend(FEDesPar.Rows - 1, 2)
'    For i = 1 To FEDesPar.Rows - 1
'        MatCalend(i - 1, 0) = FEDesPar.TextMatrix(i, 1)
'        MatCalend(i - 1, 1) = FEDesPar.TextMatrix(i, 2)
'    Next i
'    If Not ValidaDatos Then
'        Cancel = 1
'        Exit Sub
'    End If
'End Sub
