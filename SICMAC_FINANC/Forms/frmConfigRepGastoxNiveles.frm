VERSION 5.00
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmConfigRepGastoxNiveles 
   Caption         =   "Configuracón de Reporte de Gastos por Niveles"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   Icon            =   "frmConfigRepGastoxNiveles.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   10575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9360
      TabIndex        =   18
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "Quitar"
      Height          =   375
      Left            =   1440
      TabIndex        =   17
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   8880
      TabIndex        =   14
      Top             =   1520
      Width           =   1095
   End
   Begin VB.Frame fraGastoxNiveles 
      Caption         =   "Gastos por Niveles"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   255
         Left            =   1920
         TabIndex        =   22
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   5880
         Width           =   975
      End
      Begin VB.Frame fraEstructura 
         Caption         =   "Estructura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   120
         TabIndex        =   15
         Top             =   2160
         Width           =   10095
         Begin VB.CommandButton cmdBajar 
            Caption         =   "Bajar"
            Height          =   375
            Left            =   9120
            TabIndex        =   21
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton cmdSubir 
            Caption         =   "Subir"
            Height          =   375
            Left            =   9120
            TabIndex        =   20
            Top             =   240
            Width           =   855
         End
         Begin Sicmact.FlexEdit fgEstructura 
            Height          =   3255
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   8850
            _ExtentX        =   15610
            _ExtentY        =   5741
            Cols0           =   9
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "-nivel-Concepto-Fórmula-Glosa-Descripción-Orden-nGlosa-cMovNro"
            EncabezadosAnchos=   "0-500-3500-2000-550-2200-0-0-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            ColumnasAEditar =   "X-X-X-X-4-X-X-X-X"
            ListaControles  =   "0-0-0-0-4-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "L-R-L-L-C-L-R-R-C"
            FormatosEdit    =   "0-3-0-0-0-0-3-3-0"
            lbUltimaInstancia=   -1  'True
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame fraRegConcepto 
         Caption         =   "Registro de Concepto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   10095
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   7320
            TabIndex        =   13
            Top             =   800
            Width           =   1095
         End
         Begin VB.CheckBox chkGlosa 
            Caption         =   "Glosa"
            Height          =   255
            Left            =   5880
            TabIndex        =   12
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   285
            Left            =   1200
            TabIndex        =   11
            Top             =   840
            Width           =   4455
         End
         Begin VB.TextBox txtFormula 
            Height          =   285
            Left            =   6840
            TabIndex        =   9
            Top             =   360
            Width           =   2895
         End
         Begin VB.TextBox txtConcepto 
            Height          =   285
            Left            =   2640
            TabIndex        =   7
            Top             =   360
            Width           =   3015
         End
         Begin Spinner.uSpinner txtNivel 
            Height          =   255
            Left            =   720
            TabIndex        =   4
            Top             =   360
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   450
            Max             =   6
            Min             =   1
            MaxLength       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
         End
         Begin VB.Label lblComentario 
            Caption         =   "Descripción:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Fórmula:"
            Height          =   255
            Left            =   5880
            TabIndex        =   8
            Top             =   380
            Width           =   735
         End
         Begin VB.Label lblConcepto 
            Caption         =   "Concepto:"
            Height          =   255
            Left            =   1680
            TabIndex        =   6
            Top             =   380
            Width           =   855
         End
         Begin VB.Label lblNivel 
            Caption         =   "Nivel:"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   380
            Width           =   495
         End
      End
      Begin Spinner.uSpinner txtAnio 
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   450
         Max             =   9999
         Min             =   1990
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin VB.Label lblAnio 
         Alignment       =   2  'Center
         Caption         =   "Año:"
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
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmConfigRepGastoxNiveles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
'***Nombre:         frmConfigRepGastoxNiveles
'***Descripción:    Formulario que permite la configuracion
'                   de reporte de gastos por niveles.
'***Creación:       MIOL el 20130524 según ERS033-2013 OBJ A
'************************************************************
Option Explicit
Dim oRepCtaColumna As DRepCtaColumna
Dim nAnio As Integer
Dim nEditar As Integer

Private Sub cmdAceptar_Click()
Dim lnOrden As Integer
Dim lsMovNro  As String

If nEditar = 1 Then
    If Valida = False Then Exit Sub
    lnOrden = fgEstructura.TextMatrix(fgEstructura.Row, 6)
    If MsgBox(" ¿ Seguro de Actualizar los Datos ? ", vbOKCancel + vbQuestion, "Confirma grabación") = vbOk Then
        Set oRepCtaColumna = New DRepCtaColumna
        Call oRepCtaColumna.ActualizarConfGastoNivel(lnOrden, Me.txtNivel.Valor, Me.txtConcepto.Text, Me.txtFormula, Me.txtDescripcion, Me.chkGlosa.value)
        MsgBox "Los datos se actualizaron con exito!"
    End If
Else
    If Me.txtConcepto.Text = "" Or Me.txtFormula.Text = "" Then
        MsgBox "Falta completar los datos; Verificar!"
        Exit Sub
    End If
    
    If Valida = False Then Exit Sub
    
    If MsgBox(" ¿ Seguro de grabar Datos ? ", vbOKCancel + vbQuestion, "Confirma grabación") = vbOk Then
        Set oRepCtaColumna = New DRepCtaColumna
        lsMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
        Call oRepCtaColumna.InsertarConfGastoNivel(Me.txtNivel.Valor, Me.txtConcepto.Text, Me.txtFormula.Text, Me.txtDescripcion.Text, Me.chkGlosa.value, lsMovNro, Me.txtAnio.Valor)
        MsgBox "Los datos se registraron con exito!"
    End If
End If
cargarDatosEstructura (Me.txtAnio.Valor)
Call cmdCancelar_Click
Set oRepCtaColumna = Nothing
End Sub

Private Sub cargarDatosEstructura(ByVal psAnio As String)
 Set oRepCtaColumna = New DRepCtaColumna
 Dim rsConfGastoNivel As ADODB.Recordset
 Set rsConfGastoNivel = New ADODB.Recordset
 Dim i As Integer
    
   Call LimpiaFlex(fgEstructura)
    
   Set rsConfGastoNivel = oRepCtaColumna.GetConfGastoNivel(psAnio)
        If Not rsConfGastoNivel.BOF And Not rsConfGastoNivel.EOF Then
            i = 1
            fgEstructura.lbEditarFlex = True
            Do While Not rsConfGastoNivel.EOF
                fgEstructura.AdicionaFila
                fgEstructura.TextMatrix(i, 1) = rsConfGastoNivel!nNivel
                fgEstructura.TextMatrix(i, 2) = rsConfGastoNivel!cConcepto
                fgEstructura.TextMatrix(i, 3) = rsConfGastoNivel!cFormula
                fgEstructura.TextMatrix(i, 4) = rsConfGastoNivel!nGlosa
                fgEstructura.TextMatrix(i, 5) = rsConfGastoNivel!cDescripcion
                fgEstructura.TextMatrix(i, 6) = rsConfGastoNivel!nOrden
                fgEstructura.TextMatrix(i, 7) = rsConfGastoNivel!nGlosa
                fgEstructura.TextMatrix(i, 8) = rsConfGastoNivel!cMovNro
                i = i + 1
                rsConfGastoNivel.MoveNext
            Loop
        End If
    Set rsConfGastoNivel = Nothing
    Set oRepCtaColumna = Nothing
End Sub

Private Sub cmdBajar_Click()
    Dim psMaxOrden As Integer, nRow As Integer
    Set oRepCtaColumna = New DRepCtaColumna
    psMaxOrden = oRepCtaColumna.MaxOrdenConfGastosNivel(Me.txtAnio.Valor)
    nRow = fgEstructura.Row
    If fgEstructura.TextMatrix(fgEstructura.Row, 0) < psMaxOrden Then
        Call oRepCtaColumna.ActualizaConfGastosNivelOrden(fgEstructura.TextMatrix(fgEstructura.Row + 1, 6), fgEstructura.TextMatrix(fgEstructura.Row, 1), fgEstructura.TextMatrix(fgEstructura.Row, 2), fgEstructura.TextMatrix(fgEstructura.Row, 3), fgEstructura.TextMatrix(fgEstructura.Row, 5), fgEstructura.TextMatrix(fgEstructura.Row, 7), fgEstructura.TextMatrix(fgEstructura.Row, 8), _
                                        fgEstructura.TextMatrix(fgEstructura.Row, 6), fgEstructura.TextMatrix(fgEstructura.Row + 1, 1), fgEstructura.TextMatrix(fgEstructura.Row + 1, 2), fgEstructura.TextMatrix(fgEstructura.Row + 1, 3), fgEstructura.TextMatrix(fgEstructura.Row + 1, 5), fgEstructura.TextMatrix(fgEstructura.Row + 1, 7), fgEstructura.TextMatrix(fgEstructura.Row + 1, 8))
        cargarDatosEstructura (Me.txtAnio.Valor)
        fgEstructura.TopRow = nRow + 1
        fgEstructura.Row = nRow + 1
    End If
    Set oRepCtaColumna = Nothing
    fgEstructura.lbEditarFlex = False
End Sub

Private Sub cmdCancelar_Click()
'nAnio = Right(CStr(gdFecSis), 4)
'Me.txtAnio.Valor = nAnio
Me.txtNivel.Valor = 1
Me.txtConcepto.Text = ""
Me.txtFormula.Text = ""
Me.txtDescripcion.Text = ""
Me.chkGlosa.value = 0
nEditar = 0
End Sub

Private Sub cmdEditar_Click()
nEditar = 1
Me.txtNivel.Valor = fgEstructura.TextMatrix(fgEstructura.Row, 1)
Me.txtConcepto.Text = fgEstructura.TextMatrix(fgEstructura.Row, 2)
Me.txtFormula.Text = fgEstructura.TextMatrix(fgEstructura.Row, 3)
Me.txtDescripcion.Text = fgEstructura.TextMatrix(fgEstructura.Row, 5)
Me.chkGlosa.value = fgEstructura.TextMatrix(fgEstructura.Row, 7)
End Sub

Public Function Valida() As Boolean
Dim rsValida As ADODB.Recordset
Set rsValida = New ADODB.Recordset
    
    Set oRepCtaColumna = New DRepCtaColumna
    Set rsValida = oRepCtaColumna.ValidarConfGastoNivel(Me.txtNivel.Valor, Me.txtConcepto.Text, Me.txtFormula.Text)
    If rsValida.RecordCount > 0 Then
        Valida = False
        MsgBox "Los datos ya existen; Verificar!"
    Else
        Valida = True
    End If
    Set rsValida = Nothing
    Set oRepCtaColumna = Nothing
End Function

Private Sub cmdMostrar_Click()
cargarDatosEstructura (Me.txtAnio.Valor)
fgEstructura.lbEditarFlex = False
End Sub

Private Sub cmdQuitar_Click()
Dim lnOrden As Integer

    If MsgBox(" ¿ Seguro de eliminar los Datos ? ", vbOKCancel + vbQuestion, "Confirma grabación") = vbOk Then
        Set oRepCtaColumna = New DRepCtaColumna
        lnOrden = fgEstructura.TextMatrix(fgEstructura.Row, 6)
            Call oRepCtaColumna.EliminarConfGastoNivel(lnOrden)
        MsgBox "Los datos se eliminaron con exito!"
    End If
    cargarDatosEstructura (Me.txtAnio.Valor)
    Call cmdCancelar_Click
    Set oRepCtaColumna = Nothing
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub cmdSubir_Click()
    Dim nRow As Integer
    Set oRepCtaColumna = New DRepCtaColumna
    nRow = fgEstructura.Row
    If fgEstructura.TextMatrix(fgEstructura.Row, 0) <> 1 Then
        Call oRepCtaColumna.ActualizaConfGastosNivelOrden(fgEstructura.TextMatrix(fgEstructura.Row - 1, 6), fgEstructura.TextMatrix(fgEstructura.Row, 1), fgEstructura.TextMatrix(fgEstructura.Row, 2), fgEstructura.TextMatrix(fgEstructura.Row, 3), fgEstructura.TextMatrix(fgEstructura.Row, 5), fgEstructura.TextMatrix(fgEstructura.Row, 7), fgEstructura.TextMatrix(fgEstructura.Row, 8), _
                                        fgEstructura.TextMatrix(fgEstructura.Row, 6), fgEstructura.TextMatrix(fgEstructura.Row - 1, 1), fgEstructura.TextMatrix(fgEstructura.Row - 1, 2), fgEstructura.TextMatrix(fgEstructura.Row - 1, 3), fgEstructura.TextMatrix(fgEstructura.Row - 1, 5), fgEstructura.TextMatrix(fgEstructura.Row - 1, 7), fgEstructura.TextMatrix(fgEstructura.Row - 1, 8))
        cargarDatosEstructura (Me.txtAnio.Valor)
        fgEstructura.TopRow = nRow - 1
        fgEstructura.Row = nRow - 1
    End If
    Set oRepCtaColumna = Nothing
    fgEstructura.lbEditarFlex = False
End Sub

Private Sub fgEstructura_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
If fgEstructura.Col = 4 Then
    fgEstructura.lbEditarFlex = False
End If
End Sub

Private Sub Form_Load()
nAnio = Right(CStr(gdFecSis), 4)
Me.txtAnio.Valor = nAnio
cargarDatosEstructura (nAnio)
fgEstructura.lbEditarFlex = False
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
cargarDatosEstructura (Me.txtAnio.Valor)
End Sub

Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
KeyAscii = Letras(KeyAscii)
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
KeyAscii = Letras(KeyAscii)
End Sub

Private Sub txtFormula_KeyPress(KeyAscii As Integer)
KeyAscii = Letras(KeyAscii)
End Sub
