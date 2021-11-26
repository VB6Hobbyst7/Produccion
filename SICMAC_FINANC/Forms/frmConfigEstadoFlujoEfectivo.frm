VERSION 5.00
Begin VB.Form frmConfigEstadoFlujoEfectivo 
   Caption         =   "Configuración de Flujo de Efectivo"
   ClientHeight    =   6300
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12330
   Icon            =   "frmConfigEstadoFlujoEfectivo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   12330
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   " Filas del Reporte "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3855
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   12015
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         Height          =   375
         Left            =   10440
         TabIndex        =   23
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         Height          =   375
         Left            =   10440
         TabIndex        =   22
         Top             =   900
         Width           =   1215
      End
      Begin VB.CommandButton cmdSubir 
         Caption         =   "Subir"
         Height          =   375
         Left            =   10440
         TabIndex        =   21
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton cmdBajar 
         Caption         =   "Bajar"
         Height          =   375
         Left            =   10440
         TabIndex        =   20
         Top             =   2450
         Width           =   1215
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   375
         Left            =   10440
         TabIndex        =   19
         Top             =   1305
         Width           =   1215
      End
      Begin VB.CommandButton cmdGuardarOrden 
         Caption         =   "Guardar Orden"
         Height          =   375
         Left            =   10440
         TabIndex        =   18
         Top             =   2850
         Width           =   1215
      End
      Begin Sicmact.FlexEdit feFlujo 
         Height          =   3015
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5318
         Cols0           =   8
         HighLight       =   1
         EncabezadosNombres=   "#-Id-Descripción-Nivel-Tipo-Valor-Periodo-Orden"
         EncabezadosAnchos=   "350-0-3000-600-1000-5000-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0"
         CantEntero      =   9
         TextArray0      =   "#"
         TipoBusqueda    =   0
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   10320
         TabIndex        =   26
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   10320
         TabIndex        =   25
         Top             =   1920
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Agregar Fila "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12015
      Begin VB.TextBox txtDescripcion 
         Height          =   330
         Left            =   1560
         TabIndex        =   11
         Top             =   480
         Width           =   5400
      End
      Begin VB.OptionButton optFormulas 
         Caption         =   "Fórmulas"
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton optSumatoria 
         Caption         =   "Sumatoria de Filas"
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   1320
         Width           =   1695
      End
      Begin VB.OptionButton optVacio 
         Caption         =   "Vacio"
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtFormulaMenor 
         Height          =   285
         Left            =   4320
         TabIndex        =   7
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txtSumatoria 
         Height          =   285
         Left            =   4320
         TabIndex        =   6
         Top             =   1320
         Width           =   2655
      End
      Begin VB.CheckBox chkFinalPeriodo 
         Caption         =   "Final del Periodo"
         Height          =   255
         Left            =   7200
         TabIndex        =   5
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtNivel 
         Height          =   285
         Left            =   8400
         MaxLength       =   1
         TabIndex        =   4
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtFormulaMayor 
         Height          =   285
         Left            =   8400
         TabIndex        =   3
         Top             =   960
         Width           =   2775
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   8880
         TabIndex        =   2
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   10080
         TabIndex        =   1
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   370
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   15
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Fórmula < 2013"
         Height          =   255
         Left            =   3120
         TabIndex        =   14
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Nivel:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7200
         TabIndex        =   13
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Fórmula > 2013"
         Height          =   255
         Left            =   7200
         TabIndex        =   12
         Top             =   960
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmConfigEstadoFlujoEfectivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lcDescripcion As String
Dim lnNivel As Integer
Dim lcTipo As String
Dim lnTipo As Integer
Dim lnTipoVal As Integer
Dim lcValor As String
Dim lbPeriodo As String
Dim lnId As Integer
Dim lnOrden As Integer
Dim lbEstado As Boolean
Dim i As Integer
Dim guardar As Integer

Private Sub Form_Load()
    IniciarControles
    CentraForm Me
End Sub
Private Sub cmdAceptar_Click()
Dim oRep As New NRepFormula
Dim lsMovNro As String
Dim lbExito As Boolean
'Dim lnOrden As Integer
If Me.optFormulas.Enabled = True Or Me.optSumatoria.value = True Or Me.optVacio.value = True Then
    If Me.optVacio.value = False Then
        If Me.txtDescripcion.Text = "" Then
             MsgBox "Debe llenar el campo Descripcion", vbCritical, "Aviso"
             Exit Sub
        End If
        If Me.txtNivel.Text = "" Then
             MsgBox "Debe llenar el campo Nivel", vbCritical, "Aviso"
             Exit Sub
        End If
        If Me.optFormulas.value = True Then
            If Me.txtFormulaMayor.Text = "" Or Me.txtFormulaMenor.Text = "" Then
                MsgBox "Debe llenar los dos campos si marco la opcion Formula", vbCritical, "Aviso"
                Exit Sub
            End If
        End If
        If Me.optSumatoria.value = True Then
            If Me.txtSumatoria.Text = "" Then
                MsgBox "Debe llenar el campo si marco la opcion Suma", vbCritical, "Aviso"
                Exit Sub
            End If
        End If
    Else
        If Me.txtNivel.Text = "" Then
             MsgBox "Debe llenar el campo Nivel", vbCritical, "Aviso"
             Exit Sub
        End If
    End If
Else
    MsgBox "Debe seleccionar algún valor", vbCritical, "Aviso"
    Exit Sub
End If

If MsgBox("Esta Seguro de Guardar la Configuracion", vbInformation + vbYesNo, "Aviso") = vbNo Then
    Exit Sub
End If

If lnOrden = 0 Then
    If feFlujo.Rows = 2 And feFlujo.TextMatrix(1, 0) = "" Then
    lnOrden = feFlujo.Rows - 1
    Else
    lnOrden = feFlujo.Rows
    End If
End If

lcDescripcion = Me.txtDescripcion.Text
lnNivel = Me.txtNivel.Text
lnTipo = lnTipoVal
lbPeriodo = False
If Me.optFormulas.value = True Then
    lcValor = Me.txtFormulaMenor.Text & "/" & Me.txtFormulaMayor.Text
Else
    If Me.optSumatoria.value = True Then
        lcValor = Me.txtSumatoria.Text
        lbPeriodo = Me.chkFinalPeriodo.value
    Else
        lcValor = ""
        lbPeriodo = False
    End If
End If

Call oRep.RegistrarConfigFlujoEfectivo(lnId, lcDescripcion, lnNivel, lnTipo, lcValor, lbPeriodo, True, lnOrden)

Call IniciarControles
End Sub
Private Sub cmdGuardarOrden_Click()
Dim oRep As New NRepFormula
Dim lvalor As Integer
Dim lnId As Integer
Dim lcDescripcion As String
Dim lnNivel As Integer
Dim lnTipo As Integer
Dim lcValor As String
Dim lnOrden As Integer
Dim lbPeriodo As Boolean

For i = 1 To feFlujo.Rows - 1
    lnOrden = Me.feFlujo.TextMatrix(i, 0)
    lnId = Me.feFlujo.TextMatrix(i, 1)
    lcDescripcion = Me.feFlujo.TextMatrix(i, 2)
    lnNivel = Me.feFlujo.TextMatrix(i, 3) '= rsFlujo!nNivel
    lnTipo = Me.feFlujo.TextMatrix(i, 4) '= rsFlujo!nTipo
    lcValor = Me.feFlujo.TextMatrix(i, 5) '= rsFlujo!cValor
    lbPeriodo = Me.feFlujo.TextMatrix(i, 6) '= rsFlujo!bPeriodo
    Call oRep.RegistrarConfigFlujoEfectivo(lnId, lcDescripcion, lnNivel, lnTipo, lcValor, lbPeriodo, True, lnOrden)
Next
Call ListarConfiguracionFlujoEfectivo
MsgBox "Se guardo el Orden Correctamente", vbInformation
End Sub
Private Sub cmdSubir_Click()
Dim lnId_1 As Integer, lnId_2 As Integer
Dim lcDescripcion_1 As String, lcDescripcion_2 As String
Dim lnNivel_1 As Integer, lnNivel_2 As Integer
Dim lnTipo_1 As Integer, lnTipo_2 As Integer
Dim lcValor_1 As String, lcValor_2 As String
Dim lbPeriodo_1 As String, lbPeriodo_2 As String

 If feFlujo.row > 1 Then
        lnId_1 = feFlujo.TextMatrix(feFlujo.row - 1, 1)
        lcDescripcion_1 = feFlujo.TextMatrix(feFlujo.row - 1, 2)
        lnNivel_1 = feFlujo.TextMatrix(feFlujo.row - 1, 3)
        lnTipo_1 = feFlujo.TextMatrix(feFlujo.row - 1, 4)
        lcValor_1 = feFlujo.TextMatrix(feFlujo.row - 1, 5)
        lbPeriodo_1 = feFlujo.TextMatrix(feFlujo.row - 1, 6)
        
        lnId_2 = feFlujo.TextMatrix(feFlujo.row, 1)
        lcDescripcion_2 = feFlujo.TextMatrix(feFlujo.row, 2)
        lnNivel_2 = feFlujo.TextMatrix(feFlujo.row, 3)
        lnTipo_2 = feFlujo.TextMatrix(feFlujo.row, 4)
        lcValor_2 = feFlujo.TextMatrix(feFlujo.row, 5)
        lbPeriodo_2 = feFlujo.TextMatrix(feFlujo.row, 6)
        
        feFlujo.TextMatrix(feFlujo.row - 1, 1) = lnId_2
        feFlujo.TextMatrix(feFlujo.row - 1, 2) = lcDescripcion_2
        feFlujo.TextMatrix(feFlujo.row - 1, 3) = lnNivel_2
        feFlujo.TextMatrix(feFlujo.row - 1, 4) = lnTipo_2
        feFlujo.TextMatrix(feFlujo.row - 1, 5) = lcValor_2
        feFlujo.TextMatrix(feFlujo.row - 1, 6) = lbPeriodo_2
        
        feFlujo.TextMatrix(feFlujo.row, 1) = lnId_1
        feFlujo.TextMatrix(feFlujo.row, 2) = lcDescripcion_1
        feFlujo.TextMatrix(feFlujo.row, 3) = lnNivel_1
        feFlujo.TextMatrix(feFlujo.row, 4) = lnTipo_1
        feFlujo.TextMatrix(feFlujo.row, 5) = lcValor_1
        feFlujo.TextMatrix(feFlujo.row, 6) = lbPeriodo_1
               
        feFlujo.row = feFlujo.row - 1
        feFlujo.SetFocus
    End If
End Sub
Private Sub cmdBajar_Click()
Dim lnId_1 As Integer, lnId_2 As Integer
Dim lcDescripcion_1 As String, lcDescripcion_2 As String
Dim lnNivel_1 As Integer, lnNivel_2 As Integer
Dim lnTipo_1 As Integer, lnTipo_2 As Integer
Dim lcValor_1 As String, lcValor_2 As String
Dim lbPeriodo_1 As String, lbPeriodo_2 As String

 If feFlujo.row > 1 Then
        lnId_1 = feFlujo.TextMatrix(feFlujo.row + 1, 1)
        lcDescripcion_1 = feFlujo.TextMatrix(feFlujo.row + 1, 2)
        lnNivel_1 = feFlujo.TextMatrix(feFlujo.row + 1, 3)
        lnTipo_1 = feFlujo.TextMatrix(feFlujo.row + 1, 4)
        lcValor_1 = feFlujo.TextMatrix(feFlujo.row + 1, 5)
        lbPeriodo_1 = feFlujo.TextMatrix(feFlujo.row + 1, 6)
        
        lnId_2 = feFlujo.TextMatrix(feFlujo.row, 1)
        lcDescripcion_2 = feFlujo.TextMatrix(feFlujo.row, 2)
        lnNivel_2 = feFlujo.TextMatrix(feFlujo.row, 3)
        lnTipo_2 = feFlujo.TextMatrix(feFlujo.row, 4)
        lcValor_2 = feFlujo.TextMatrix(feFlujo.row, 5)
        lbPeriodo_2 = feFlujo.TextMatrix(feFlujo.row, 6)
        
        feFlujo.TextMatrix(feFlujo.row + 1, 1) = lnId_2
        feFlujo.TextMatrix(feFlujo.row + 1, 2) = lcDescripcion_2
        feFlujo.TextMatrix(feFlujo.row + 1, 3) = lnNivel_2
        feFlujo.TextMatrix(feFlujo.row + 1, 4) = lnTipo_2
        feFlujo.TextMatrix(feFlujo.row + 1, 5) = lcValor_2
        feFlujo.TextMatrix(feFlujo.row + 1, 6) = lbPeriodo_2
        
        feFlujo.TextMatrix(feFlujo.row, 1) = lnId_1
        feFlujo.TextMatrix(feFlujo.row, 2) = lcDescripcion_1
        feFlujo.TextMatrix(feFlujo.row, 3) = lnNivel_1
        feFlujo.TextMatrix(feFlujo.row, 4) = lnTipo_1
        feFlujo.TextMatrix(feFlujo.row, 5) = lcValor_1
        feFlujo.TextMatrix(feFlujo.row, 6) = lbPeriodo_1
               
        feFlujo.row = feFlujo.row + 1
        feFlujo.SetFocus
    End If
End Sub
Private Sub cmdCancelar_Click()
Call IniciarControles
End Sub
Private Sub cmdExportar_Click()
Dim oReps As New NRepFormula
Dim rDataRep As New ADODB.Recordset
Dim j As Integer
Dim ApExcel As Variant
Set ApExcel = CreateObject("Excel.application")
Set rDataRep = oReps.RecuperaConfigFlujoEfectivo()
'Agrega un nuevo Libro
ApExcel.Workbooks.Add
'detalle
ApExcel.Range("B7", "G7").MergeCells = True
ApExcel.Range("B7", "G7").HorizontalAlignment = xlCenter
ApExcel.Cells(7, 2).Formula = "CONFIGURACIÓN DE FLUJO DE EFECTIVO"
ApExcel.Cells(9, 2).Formula = "ITEM"
ApExcel.Cells(9, 3).Formula = "DESCRIPCION"
ApExcel.Cells(9, 4).Formula = "NIVEL"
ApExcel.Cells(9, 5).Formula = "TIPO"
ApExcel.Cells(9, 6).Formula = "VALOR"
Call IniciarControles
j = 10
Do While Not rDataRep.EOF
ApExcel.Cells(j, 2).Formula = j - 9
ApExcel.Cells(j, 3).Formula = rDataRep!cDescripcion
ApExcel.Cells(j, 4).Formula = rDataRep!nNivel
ApExcel.Cells(j, 5).Formula = rDataRep!nTipo
ApExcel.Cells(j, 6).Formula = rDataRep!cValor
ApExcel.Cells(j, 7).Formula = rDataRep!bPeriodo
ApExcel.Range("B" & Trim(Str(j - 1)) & ":" & "G" & Trim(Str(j - 1))).Borders.LineStyle = 1
j = j + 1
rDataRep.MoveNext
Loop
ApExcel.Range("B" & Trim(Str(j - 1)) & ":" & "G" & Trim(Str(j - 1))).Borders.LineStyle = 1
rDataRep.Close
Set rDataRep = Nothing
ApExcel.Visible = True
Set ApExcel = Nothing
End Sub
Private Sub cmdQuitar_Click()
Dim oRep As New NRepFormula
lbEstado = False
Call Cargar
Call oRep.RegistrarConfigFlujoEfectivo(lnId, lcDescripcion, lnNivel, lnTipo, lcValor, lbPeriodo, lbEstado, lnOrden)
Call IniciarControles
End Sub
Public Sub Inicio(ByVal psOpeCod As String, psOpeDesc As String)
    'fsOpeCod = psOpeCod
    'fsOpeDesc = psOpeDesc
    'Caption = "CONFIGURACIÓN " & UCase(psOpeDesc)
    Show 1
End Sub
Private Sub IniciarControles()
    Call ListarConfiguracionFlujoEfectivo
    Me.txtSumatoria.Enabled = False
    Me.txtSumatoria.Text = ""
    Me.chkFinalPeriodo.Enabled = False
    Me.chkFinalPeriodo.value = False
    Me.txtFormulaMayor.Enabled = False
    Me.txtFormulaMayor.Text = ""
    Me.txtFormulaMenor.Enabled = False
    Me.txtFormulaMenor.Text = ""
    Me.txtDescripcion.Text = ""
    Me.txtNivel.Text = ""
    Me.cmdSubir.Enabled = False
    Me.cmdBajar.Enabled = False
    Me.cmdQuitar.Enabled = False
    Me.cmdModificar.Enabled = False
    lnId = 0
End Sub
Private Sub optFormulas_Click()
    If Me.optFormulas.value = True Then
        Me.txtSumatoria.Enabled = False
        Me.chkFinalPeriodo.Enabled = False
        Me.txtFormulaMayor.Enabled = True
        Me.txtFormulaMenor.Enabled = True
        lnTipoVal = 1
    End If
End Sub
Private Sub optSumatoria_Click()
    If Me.optSumatoria.value = True Then
        Me.txtSumatoria.Enabled = True
        Me.chkFinalPeriodo.Enabled = True
        Me.txtFormulaMayor.Enabled = False
        Me.txtFormulaMenor.Enabled = False
        lnTipoVal = 2
    End If
End Sub
Private Sub optVacio_Click()
    If Me.optVacio.value = True Then
        Me.txtSumatoria.Enabled = False
        Me.chkFinalPeriodo.Enabled = False
        Me.txtFormulaMayor.Enabled = False
        Me.txtFormulaMenor.Enabled = False
        lnTipoVal = 3
    End If
End Sub
Private Sub ListarConfiguracionFlujoEfectivo()
    Dim oRep As New NRepFormula
    Dim rsFlujo As New ADODB.Recordset
    Dim i As Long
    Set rsFlujo = oRep.RecuperaConfigFlujoEfectivo()
    Call LimpiaFlex(feFlujo)
    If Not RSVacio(rsFlujo) Then
        For i = 1 To rsFlujo.RecordCount
            Me.feFlujo.AdicionaFila
            Me.feFlujo.TextMatrix(feFlujo.row, 1) = rsFlujo!nId
            Me.feFlujo.TextMatrix(feFlujo.row, 2) = rsFlujo!cDescripcion
            Me.feFlujo.TextMatrix(feFlujo.row, 3) = rsFlujo!nNivel
            Me.feFlujo.TextMatrix(feFlujo.row, 4) = rsFlujo!nTipo
            Me.feFlujo.TextMatrix(feFlujo.row, 5) = rsFlujo!cValor
            Me.feFlujo.TextMatrix(feFlujo.row, 6) = rsFlujo!bPeriodo
            Me.feFlujo.TextMatrix(feFlujo.row, 7) = rsFlujo!nOrden
            rsFlujo.MoveNext
        Next
        Me.cmdQuitar.Enabled = True
        Me.cmdModificar.Enabled = True
        If rsFlujo.RecordCount >= 3 Then
            Me.cmdSubir.Enabled = True
            Me.cmdBajar.Enabled = True
        End If
    Else
        Me.cmdSubir.Enabled = False
        Me.cmdBajar.Enabled = False
    End If
    Set oRep = Nothing
    Set rsFlujo = Nothing
End Sub
Private Sub CmdModificar_Click()
Call Cargar
End Sub
Private Sub Cargar()
Dim row, col As Integer
row = feFlujo.row
col = feFlujo.col

lnId = Me.feFlujo.TextMatrix(row, 1)
lcDescripcion = Me.feFlujo.TextMatrix(row, 2)
lnNivel = Me.feFlujo.TextMatrix(row, 3)
lnTipo = Me.feFlujo.TextMatrix(row, 4)
lcValor = Me.feFlujo.TextMatrix(row, 5)
lbPeriodo = Me.feFlujo.TextMatrix(row, 6)
lnOrden = Me.feFlujo.TextMatrix(row, 7)

Me.txtDescripcion.Text = lcDescripcion
Me.txtNivel.Text = lnNivel
Select Case lnTipo
    Case 1 'formulas
    Me.optFormulas.value = True
    Me.txtFormulaMenor.Text = Mid(lcValor, 1, InStr(lcValor, "/") - 1)
    Me.txtFormulaMayor.Text = Mid(lcValor, InStr(lcValor, "/") + 1, Len(lcValor))
    Me.txtSumatoria.Text = ""
    Me.chkFinalPeriodo.value = False
    Case 2 'sumatoria
    Me.optSumatoria.value = True
    Me.txtSumatoria.Text = lcValor
    Me.chkFinalPeriodo.value = CBool(lbPeriodo)
    Me.txtFormulaMayor.Text = ""
    Me.txtFormulaMenor.Text = ""
    Case 3 'vacio
    Me.optVacio.value = True
    Me.txtSumatoria.Text = ""
    Me.chkFinalPeriodo.value = False
    Me.txtFormulaMayor.Text = ""
    Me.txtFormulaMenor.Text = ""
End Select
End Sub
Private Sub feFlujo_Click()
If feFlujo.row > 0 Then
        If feFlujo.TextMatrix(feFlujo.row, 0) <> "" Then
            Me.cmdSubir.Enabled = True
            Me.cmdBajar.Enabled = True
            Me.cmdQuitar.Enabled = True
            Me.cmdModificar.Enabled = True
        End If
End If
End Sub

Private Sub txtFormulaMayor_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnterosSignosMasyMenos(KeyAscii, False)
End Sub

Private Sub txtFormulaMenor_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnterosSignosMasyMenos(KeyAscii, False)
End Sub

Private Sub txtNivel_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii, False)
End Sub

Private Sub txtSumatoria_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnterosSignosMasyMenos(KeyAscii, False)
End Sub
