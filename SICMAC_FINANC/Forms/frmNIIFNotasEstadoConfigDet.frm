VERSION 5.00
Begin VB.Form frmNIIFNotasEstadoConfigDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración del Detalle de las Notas de Estado"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13875
   Icon            =   "frmNIIFNotasEstadoConfigDet.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   13875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "&Quitar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1200
      TabIndex        =   5
      Top             =   3550
      Width           =   1050
   End
   Begin VB.CommandButton cmdBajar 
      Caption         =   "&Bajar Orden"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7200
      TabIndex        =   4
      Top             =   3550
      Width           =   1050
   End
   Begin VB.CommandButton cmdSubir 
      Caption         =   "&Subir Orden"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6120
      TabIndex        =   3
      Top             =   3550
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   12720
      TabIndex        =   2
      Top             =   3550
      Width           =   1050
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   11640
      TabIndex        =   1
      Top             =   3550
      Width           =   1050
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   3550
      Width           =   1050
   End
   Begin Sicmact.FlexEdit feNotasDet 
      Height          =   3405
      Left            =   40
      TabIndex        =   7
      Top             =   40
      Width           =   13800
      _ExtentX        =   24342
      _ExtentY        =   6006
      Cols0           =   16
      HighLight       =   1
      EncabezadosNombres=   $"frmNIIFNotasEstadoConfigDet.frx":030A
      EncabezadosAnchos=   "350-1100-2800-800-1050-1500-1500-1500-1500-1500-1500-1500-1500-1500-1500-0"
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
      ColumnasAEditar =   "X-1-2-3-4-5-6-7-8-9-10-11-12-13-14-X"
      ListaControles  =   "0-3-1-3-3-0-0-0-0-0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "C-L-L-L-L-L-L-L-L-L-L-L-L-L-L-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      CantEntero      =   9
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      TipoBusqueda    =   6
      lbPuntero       =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   345
      RowHeight0      =   300
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   525
      Left            =   45
      TabIndex        =   6
      Top             =   3470
      Width           =   13780
   End
End
Attribute VB_Name = "frmNIIFNotasEstadoConfigDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmNIIFNotasEstadoConfigDet
'** Descripción : Configuración del Reporte Notas Estado creado segun ERS052-2013
'** Creación : EJVG, 20130418 11:49:00 AM
'********************************************************************
Option Explicit
Dim fsOpeCod As String
Dim MatNotasEstadoDetInicio As Variant
Dim fbAceptar As Boolean

Private Sub Form_Load()
    CentraForm Me
    IniciarControles
End Sub
Public Function Inicio(ByVal psOpeCod As String, psOpeDesc As String, pMatNotaEstadoDet As Variant) As Variant
    fbAceptar = False
    fsOpeCod = psOpeCod
    MatNotasEstadoDetInicio = pMatNotaEstadoDet
    Call ListarConfiguracionNotasDetalle(pMatNotaEstadoDet)
    Show 1
    Inicio = RecuperaNotasEstadoDetalle()
End Function
Private Sub cmdAgregar_Click()
    If Not (feNotasDet.Rows - 1 = 1 And Len(Trim(feNotasDet.TextMatrix(1, 0))) = 0) Then 'Flex no esta Vacio
        If validarRegistroDatosNotasEstadoDet = False Then Exit Sub
    End If
    feNotasDet.AdicionaFila
    feNotasDet.SetFocus
    feNotasDet.Col = 1
    feNotasDet_RowColChange
End Sub
Private Sub cmdQuitar_Click()
    feNotasDet.EliminaFila feNotasDet.Row
End Sub
Private Sub cmdSubir_Click()
    Dim lsTipoDetalle1 As String, lsTipoDetalle2 As String
    Dim lsDesc1 As String, lsDesc2 As String
    Dim lsNivel1 As String, lsNivel2 As String
    Dim lsNegrita1 As String, lsNegrita2 As String
    Dim lsFormula1_1 As String, lsFormula1_2 As String
    Dim lsFormula1_2012_1 As String, lsFormula1_2012_2 As String
    Dim lsFormula2_1 As String, lsFormula2_2 As String
    Dim lsFormula2_2012_1 As String, lsFormula2_2012_2 As String
    Dim lsFormula3_1 As String, lsFormula3_2 As String
    Dim lsFormula3_2012_1 As String, lsFormula3_2012_2 As String
    Dim lsFormula4_1 As String, lsFormula4_2 As String
    Dim lsFormula4_2012_1 As String, lsFormula4_2012_2 As String
    Dim lsFormula5_1 As String, lsFormula5_2 As String
    Dim lsFormula5_2012_1 As String, lsFormula5_2012_2 As String
    
    If validarRegistroDatosNotasEstadoDet = False Then Exit Sub

    If feNotasDet.Row > 1 Then
        lsTipoDetalle1 = feNotasDet.TextMatrix(feNotasDet.Row - 1, 1)
        lsDesc1 = feNotasDet.TextMatrix(feNotasDet.Row - 1, 2)
        lsNivel1 = feNotasDet.TextMatrix(feNotasDet.Row - 1, 3)
        lsNegrita1 = feNotasDet.TextMatrix(feNotasDet.Row - 1, 4)
        lsFormula1_1 = feNotasDet.TextMatrix(feNotasDet.Row - 1, 5)
        lsFormula1_2012_1 = feNotasDet.TextMatrix(feNotasDet.Row - 1, 6)
        lsFormula2_1 = feNotasDet.TextMatrix(feNotasDet.Row - 1, 7)
        lsFormula2_2012_1 = feNotasDet.TextMatrix(feNotasDet.Row - 1, 8)
        lsFormula3_1 = feNotasDet.TextMatrix(feNotasDet.Row - 1, 9)
        lsFormula3_2012_1 = feNotasDet.TextMatrix(feNotasDet.Row - 1, 10)
        lsFormula4_1 = feNotasDet.TextMatrix(feNotasDet.Row - 1, 11)
        lsFormula4_2012_1 = feNotasDet.TextMatrix(feNotasDet.Row - 1, 12)
        lsFormula5_1 = feNotasDet.TextMatrix(feNotasDet.Row - 1, 13)
        lsFormula5_2012_1 = feNotasDet.TextMatrix(feNotasDet.Row - 1, 14)
        
        lsTipoDetalle2 = feNotasDet.TextMatrix(feNotasDet.Row, 1)
        lsDesc2 = feNotasDet.TextMatrix(feNotasDet.Row, 2)
        lsNivel2 = feNotasDet.TextMatrix(feNotasDet.Row, 3)
        lsNegrita2 = feNotasDet.TextMatrix(feNotasDet.Row, 4)
        lsFormula1_2 = feNotasDet.TextMatrix(feNotasDet.Row, 5)
        lsFormula1_2012_2 = feNotasDet.TextMatrix(feNotasDet.Row, 6)
        lsFormula2_2 = feNotasDet.TextMatrix(feNotasDet.Row, 7)
        lsFormula2_2012_2 = feNotasDet.TextMatrix(feNotasDet.Row, 8)
        lsFormula3_2 = feNotasDet.TextMatrix(feNotasDet.Row, 9)
        lsFormula3_2012_2 = feNotasDet.TextMatrix(feNotasDet.Row, 10)
        lsFormula4_2 = feNotasDet.TextMatrix(feNotasDet.Row, 11)
        lsFormula4_2012_2 = feNotasDet.TextMatrix(feNotasDet.Row, 12)
        lsFormula5_2 = feNotasDet.TextMatrix(feNotasDet.Row, 13)
        lsFormula5_2012_2 = feNotasDet.TextMatrix(feNotasDet.Row, 14)
        
        feNotasDet.TextMatrix(feNotasDet.Row - 1, 1) = lsTipoDetalle2
        feNotasDet.TextMatrix(feNotasDet.Row - 1, 2) = lsDesc2
        feNotasDet.TextMatrix(feNotasDet.Row - 1, 3) = lsNivel2
        feNotasDet.TextMatrix(feNotasDet.Row - 1, 4) = lsNegrita2
        feNotasDet.TextMatrix(feNotasDet.Row - 1, 5) = lsFormula1_2
        feNotasDet.TextMatrix(feNotasDet.Row - 1, 6) = lsFormula1_2012_2
        feNotasDet.TextMatrix(feNotasDet.Row - 1, 7) = lsFormula2_2
        feNotasDet.TextMatrix(feNotasDet.Row - 1, 8) = lsFormula2_2012_2
        feNotasDet.TextMatrix(feNotasDet.Row - 1, 9) = lsFormula3_2
        feNotasDet.TextMatrix(feNotasDet.Row - 1, 10) = lsFormula3_2012_2
        feNotasDet.TextMatrix(feNotasDet.Row - 1, 11) = lsFormula4_2
        feNotasDet.TextMatrix(feNotasDet.Row - 1, 12) = lsFormula4_2012_2
        feNotasDet.TextMatrix(feNotasDet.Row - 1, 13) = lsFormula5_2
        feNotasDet.TextMatrix(feNotasDet.Row - 1, 14) = lsFormula5_2012_2
        
        feNotasDet.TextMatrix(feNotasDet.Row, 1) = lsTipoDetalle1
        feNotasDet.TextMatrix(feNotasDet.Row, 2) = lsDesc1
        feNotasDet.TextMatrix(feNotasDet.Row, 3) = lsNivel1
        feNotasDet.TextMatrix(feNotasDet.Row, 4) = lsNegrita1
        feNotasDet.TextMatrix(feNotasDet.Row, 5) = lsFormula1_1
        feNotasDet.TextMatrix(feNotasDet.Row, 6) = lsFormula1_2012_1
        feNotasDet.TextMatrix(feNotasDet.Row, 7) = lsFormula2_1
        feNotasDet.TextMatrix(feNotasDet.Row, 8) = lsFormula2_2012_1
        feNotasDet.TextMatrix(feNotasDet.Row, 9) = lsFormula3_1
        feNotasDet.TextMatrix(feNotasDet.Row, 10) = lsFormula3_2012_1
        feNotasDet.TextMatrix(feNotasDet.Row, 11) = lsFormula4_1
        feNotasDet.TextMatrix(feNotasDet.Row, 12) = lsFormula4_2012_1
        feNotasDet.TextMatrix(feNotasDet.Row, 13) = lsFormula5_1
        feNotasDet.TextMatrix(feNotasDet.Row, 14) = lsFormula5_2012_1
               
        feNotasDet.Row = feNotasDet.Row - 1
        feNotasDet.SetFocus
    End If
End Sub
Private Sub cmdBajar_Click()
    Dim lsTipoDetalle1 As String, lsTipoDetalle2 As String
    Dim lsDesc1 As String, lsDesc2 As String
    Dim lsNivel1 As String, lsNivel2 As String
    Dim lsNegrita1 As String, lsNegrita2 As String
    Dim lsFormula1_1 As String, lsFormula1_2 As String
    Dim lsFormula1_2012_1 As String, lsFormula1_2012_2 As String
    Dim lsFormula2_1 As String, lsFormula2_2 As String
    Dim lsFormula2_2012_1 As String, lsFormula2_2012_2 As String
    Dim lsFormula3_1 As String, lsFormula3_2 As String
    Dim lsFormula3_2012_1 As String, lsFormula3_2012_2 As String
    Dim lsFormula4_1 As String, lsFormula4_2 As String
    Dim lsFormula4_2012_1 As String, lsFormula4_2012_2 As String
    Dim lsFormula5_1 As String, lsFormula5_2 As String
    Dim lsFormula5_2012_1 As String, lsFormula5_2012_2 As String

    If validarRegistroDatosNotasEstadoDet = False Then Exit Sub

    If feNotasDet.Row < feNotasDet.Rows - 1 Then
        lsTipoDetalle1 = feNotasDet.TextMatrix(feNotasDet.Row + 1, 1)
        lsDesc1 = feNotasDet.TextMatrix(feNotasDet.Row + 1, 2)
        lsNivel1 = feNotasDet.TextMatrix(feNotasDet.Row + 1, 3)
        lsNegrita1 = feNotasDet.TextMatrix(feNotasDet.Row + 1, 4)
        lsFormula1_1 = feNotasDet.TextMatrix(feNotasDet.Row + 1, 5)
        lsFormula1_2012_1 = feNotasDet.TextMatrix(feNotasDet.Row + 1, 6)
        lsFormula2_1 = feNotasDet.TextMatrix(feNotasDet.Row + 1, 7)
        lsFormula2_2012_1 = feNotasDet.TextMatrix(feNotasDet.Row + 1, 8)
        lsFormula3_1 = feNotasDet.TextMatrix(feNotasDet.Row + 1, 9)
        lsFormula3_2012_1 = feNotasDet.TextMatrix(feNotasDet.Row + 1, 10)
        lsFormula4_1 = feNotasDet.TextMatrix(feNotasDet.Row + 1, 11)
        lsFormula4_2012_1 = feNotasDet.TextMatrix(feNotasDet.Row + 1, 12)
        lsFormula5_1 = feNotasDet.TextMatrix(feNotasDet.Row + 1, 13)
        lsFormula5_2012_1 = feNotasDet.TextMatrix(feNotasDet.Row + 1, 14)
        
        lsTipoDetalle2 = feNotasDet.TextMatrix(feNotasDet.Row, 1)
        lsDesc2 = feNotasDet.TextMatrix(feNotasDet.Row, 2)
        lsNivel2 = feNotasDet.TextMatrix(feNotasDet.Row, 3)
        lsNegrita2 = feNotasDet.TextMatrix(feNotasDet.Row, 4)
        lsFormula1_2 = feNotasDet.TextMatrix(feNotasDet.Row, 5)
        lsFormula1_2012_2 = feNotasDet.TextMatrix(feNotasDet.Row, 6)
        lsFormula2_2 = feNotasDet.TextMatrix(feNotasDet.Row, 7)
        lsFormula2_2012_2 = feNotasDet.TextMatrix(feNotasDet.Row, 8)
        lsFormula3_2 = feNotasDet.TextMatrix(feNotasDet.Row, 9)
        lsFormula3_2012_2 = feNotasDet.TextMatrix(feNotasDet.Row, 10)
        lsFormula4_2 = feNotasDet.TextMatrix(feNotasDet.Row, 11)
        lsFormula4_2012_2 = feNotasDet.TextMatrix(feNotasDet.Row, 12)
        lsFormula5_2 = feNotasDet.TextMatrix(feNotasDet.Row, 13)
        lsFormula5_2012_2 = feNotasDet.TextMatrix(feNotasDet.Row, 14)
        
        feNotasDet.TextMatrix(feNotasDet.Row + 1, 1) = lsTipoDetalle2
        feNotasDet.TextMatrix(feNotasDet.Row + 1, 2) = lsDesc2
        feNotasDet.TextMatrix(feNotasDet.Row + 1, 3) = lsNivel2
        feNotasDet.TextMatrix(feNotasDet.Row + 1, 4) = lsNegrita2
        feNotasDet.TextMatrix(feNotasDet.Row + 1, 5) = lsFormula1_2
        feNotasDet.TextMatrix(feNotasDet.Row + 1, 6) = lsFormula1_2012_2
        feNotasDet.TextMatrix(feNotasDet.Row + 1, 7) = lsFormula2_2
        feNotasDet.TextMatrix(feNotasDet.Row + 1, 8) = lsFormula2_2012_2
        feNotasDet.TextMatrix(feNotasDet.Row + 1, 9) = lsFormula3_2
        feNotasDet.TextMatrix(feNotasDet.Row + 1, 10) = lsFormula3_2012_2
        feNotasDet.TextMatrix(feNotasDet.Row + 1, 11) = lsFormula4_2
        feNotasDet.TextMatrix(feNotasDet.Row + 1, 12) = lsFormula4_2012_2
        feNotasDet.TextMatrix(feNotasDet.Row + 1, 13) = lsFormula5_2
        feNotasDet.TextMatrix(feNotasDet.Row + 1, 14) = lsFormula5_2012_2
        
        feNotasDet.TextMatrix(feNotasDet.Row, 1) = lsTipoDetalle1
        feNotasDet.TextMatrix(feNotasDet.Row, 2) = lsDesc1
        feNotasDet.TextMatrix(feNotasDet.Row, 3) = lsNivel1
        feNotasDet.TextMatrix(feNotasDet.Row, 4) = lsNegrita1
        feNotasDet.TextMatrix(feNotasDet.Row, 5) = lsFormula1_1
        feNotasDet.TextMatrix(feNotasDet.Row, 6) = lsFormula1_2012_1
        feNotasDet.TextMatrix(feNotasDet.Row, 7) = lsFormula2_1
        feNotasDet.TextMatrix(feNotasDet.Row, 8) = lsFormula2_2012_1
        feNotasDet.TextMatrix(feNotasDet.Row, 9) = lsFormula3_1
        feNotasDet.TextMatrix(feNotasDet.Row, 10) = lsFormula3_2012_1
        feNotasDet.TextMatrix(feNotasDet.Row, 11) = lsFormula4_1
        feNotasDet.TextMatrix(feNotasDet.Row, 12) = lsFormula4_2012_1
        feNotasDet.TextMatrix(feNotasDet.Row, 13) = lsFormula5_1
        feNotasDet.TextMatrix(feNotasDet.Row, 14) = lsFormula5_2012_1

        feNotasDet.Row = feNotasDet.Row + 1
        feNotasDet.SetFocus
    End If
End Sub
Private Sub cmdAceptar_Click()
    If Not (feNotasDet.Rows - 1 = 1 And Len(Trim(feNotasDet.TextMatrix(1, 0))) = 0) Then 'Flex no esta Vacio
        If validarRegistroDatosNotasEstadoDet = False Then Exit Sub
    End If
    fbAceptar = True
    Hide
End Sub
Private Sub cmdSalir_Click()
    fbAceptar = False
    Hide
End Sub
Private Sub IniciarControles()
    cmdSubir.Enabled = False
    cmdBajar.Enabled = False
End Sub
Private Sub ListarConfiguracionNotasDetalle(ByVal pMatNotaEstadoDet As Variant)
    Dim i As Long
    Call LimpiaFlex(feNotasDet)
    For i = 1 To UBound(pMatNotaEstadoDet, 2)
        feNotasDet.AdicionaFila
        feNotasDet.TextMatrix(feNotasDet.Row, 1) = IIf(pMatNotaEstadoDet(1, i) = "1", "TEXTO" & Space(75) & "1", "FORMULA" & Space(75) & "2") 'Tipo
        feNotasDet.TextMatrix(feNotasDet.Row, 2) = pMatNotaEstadoDet(2, i) 'Descripcion
        feNotasDet.TextMatrix(feNotasDet.Row, 3) = pMatNotaEstadoDet(3, i) & Space(75) & pMatNotaEstadoDet(3, i) 'Nivel
        feNotasDet.TextMatrix(feNotasDet.Row, 4) = IIf(pMatNotaEstadoDet(4, i) = "1", "SI" & Space(75) & "1", "NO" & Space(75) & "2") 'Negrita
        feNotasDet.TextMatrix(feNotasDet.Row, 5) = pMatNotaEstadoDet(5, i) 'Formula 1
        feNotasDet.TextMatrix(feNotasDet.Row, 6) = pMatNotaEstadoDet(6, i) 'Formula 1 <= 2012
        feNotasDet.TextMatrix(feNotasDet.Row, 7) = pMatNotaEstadoDet(7, i) 'Formula 2
        feNotasDet.TextMatrix(feNotasDet.Row, 8) = pMatNotaEstadoDet(8, i) 'Formula 2 <= 2012
        feNotasDet.TextMatrix(feNotasDet.Row, 9) = pMatNotaEstadoDet(9, i) 'Formula 3
        feNotasDet.TextMatrix(feNotasDet.Row, 10) = pMatNotaEstadoDet(10, i) 'Formula 3 <= 2012
        feNotasDet.TextMatrix(feNotasDet.Row, 11) = pMatNotaEstadoDet(11, i) 'Formula 4
        feNotasDet.TextMatrix(feNotasDet.Row, 12) = pMatNotaEstadoDet(12, i) 'Formula 4 <= 2012
        feNotasDet.TextMatrix(feNotasDet.Row, 13) = pMatNotaEstadoDet(13, i) 'Formula 5
        feNotasDet.TextMatrix(feNotasDet.Row, 14) = pMatNotaEstadoDet(14, i) 'Formula 5 <= 2012
    Next
    feNotasDet.TopRow = 1
End Sub
Private Function RecuperaNotasEstadoDetalle() As Variant
    Dim Mat As Variant
    Dim i As Long
    
    ReDim Mat(1 To 14, 0)
    If fbAceptar Then
        If Not (feNotasDet.Rows - 1 = 1 And Len(Trim(feNotasDet.TextMatrix(1, 0))) = 0) Then 'Flex no esta Vacio
            For i = 1 To feNotasDet.Rows - 1
                ReDim Preserve Mat(1 To 14, 1 To i)
                Mat(1, i) = Trim(Right(feNotasDet.TextMatrix(i, 1), 2)) 'Tipo de Detalle
                Mat(2, i) = feNotasDet.TextMatrix(i, 2) 'Descripcion
                Mat(3, i) = Trim(Right(feNotasDet.TextMatrix(i, 3), 2)) 'Nivel
                Mat(4, i) = Trim(Right(feNotasDet.TextMatrix(i, 4), 2)) 'Negrita
                Mat(5, i) = feNotasDet.TextMatrix(i, 5) 'Formula 1
                Mat(6, i) = feNotasDet.TextMatrix(i, 6) 'Formula 1 <= 2012
                Mat(7, i) = feNotasDet.TextMatrix(i, 7) 'Formula 2
                Mat(8, i) = feNotasDet.TextMatrix(i, 8) 'Formula 2 <= 2012
                Mat(9, i) = feNotasDet.TextMatrix(i, 9) 'Formula 3
                Mat(10, i) = feNotasDet.TextMatrix(i, 10) 'Formula 3 <= 2012
                Mat(11, i) = feNotasDet.TextMatrix(i, 11) 'Formula 4
                Mat(12, i) = feNotasDet.TextMatrix(i, 12) 'Formula 4 <= 2012
                Mat(13, i) = feNotasDet.TextMatrix(i, 13) 'Formula 5
                Mat(14, i) = feNotasDet.TextMatrix(i, 14) 'Formula 5 <= 2012
            Next
        End If
        RecuperaNotasEstadoDetalle = Mat
    Else
        RecuperaNotasEstadoDetalle = MatNotasEstadoDetInicio
    End If
    Set Mat = Nothing
End Function
Private Sub feNotasDet_RowColChange()
    Dim rsOpt As New ADODB.Recordset
    Dim i As Integer
    
    If feNotasDet.Col = 1 Or feNotasDet.Col = 3 Or feNotasDet.Col = 4 Then
        With rsOpt
            .Fields.Append "desc", adVarChar, 10
            .Fields.Append "value", adVarChar, 2
        End With
        If feNotasDet.Col = 1 Then
            With rsOpt
                .Open
                .AddNew
                .Fields("desc") = "TEXTO"
                .Fields("value") = "1"
                .AddNew
                .Fields("desc") = "FORMULA"
                .Fields("value") = "2"
            End With
        ElseIf feNotasDet.Col = 3 Then
            With rsOpt
                .Open
                For i = 1 To 10
                    .AddNew
                    .Fields("desc") = CStr(i)
                    .Fields("value") = CStr(i)
                Next
            End With
        ElseIf feNotasDet.Col = 4 Then
            With rsOpt
                .Open
                .AddNew
                .Fields("desc") = "SI"
                .Fields("value") = "1"
                .AddNew
                .Fields("desc") = "NO"
                .Fields("value") = "2"
            End With
        End If
        rsOpt.MoveFirst
        feNotasDet.CargaCombo rsOpt
    End If
    Set rsOpt = Nothing
End Sub
Private Sub feNotasDet_Click()
    If feNotasDet.Row > 0 Then
        If feNotasDet.TextMatrix(feNotasDet.Row, 0) <> "" Then
            cmdSubir.Enabled = True
            cmdBajar.Enabled = True
        End If
    End If
End Sub
Private Sub feNotasDet_OnClickTxtBuscar(psCodigo As String, psDescripcion As String)
    feNotasDet.TextMatrix(feNotasDet.Row, 2) = frmNIIFNotasEstadoConfigDetTexto.Inicio(feNotasDet.TextMatrix(feNotasDet.Row, 2))
    psCodigo = feNotasDet.TextMatrix(feNotasDet.Row, 2)
    psDescripcion = psCodigo
End Sub
Private Function validarRegistroDatosNotasEstadoDet() As Boolean
    validarRegistroDatosNotasEstadoDet = True
    Dim i As Long, j As Long
    For i = 1 To feNotasDet.Rows - 1 'valida fila x fila
        For j = 1 To feNotasDet.Cols - 2 '2 xq el ultimo es aux
            If j = 1 Or j = 3 Or j = 4 Then 'xq las formulas y la desc son opcionales
                If Trim(feNotasDet.TextMatrix(i, j)) = "" Then
                    validarRegistroDatosNotasEstadoDet = False
                    MsgBox "Ud. debe de ingresar el dato '" & UCase(feNotasDet.TextMatrix(0, j)) & "'", vbInformation, "Aviso"
                    feNotasDet.Row = i
                    feNotasDet.Col = j
                    feNotasDet.SetFocus
                    Exit Function
                End If
            End If
        Next
    Next
End Function
