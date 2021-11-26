VERSION 5.00
Begin VB.Form frmAhorrosProyecccionesPorAgencia 
   Caption         =   "Proyecciones de Ahorros por Agencia"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13185
   Icon            =   "frmAhorrosProyecccionesPorAgencia.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   13185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnGuardar 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   6
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      TabIndex        =   5
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton btnSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   4
      Top             =   5880
      Width           =   1095
   End
   Begin Sicmact.FlexEdit feProyeccion 
      Height          =   4695
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   8281
      Cols0           =   16
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-AgeCod-Agencia-Anio-Enero-Febrero-Marzo-Abril-Mayo-Junio-Julio-Agosto-Septiembre-Octubre-Noviembre-Diciembre"
      EncabezadosAnchos=   "0-0-2000-0-1800-1800-1800-1800-1800-1800-1800-1800-1800-1800-1800-1800"
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
      ColumnasAEditar =   "X-X-X-X-4-5-6-7-8-9-10-11-12-13-14-15"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L-C-R-R-R-R-R-R-R-R-R-R-R-R"
      FormatosEdit    =   "0-0-0-0-2-2-2-2-2-2-2-2-2-2-2-2"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbPuntero       =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione Año"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12975
      Begin VB.CommandButton btnSeleccionar 
         Caption         =   "Seleccionar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   350
         Width           =   1215
      End
      Begin VB.TextBox txtAnio 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   360
         MaxLength       =   4
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmAhorrosProyecccionesPorAgencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmAhorrosProyecccionesPorAgencia
'** Descripción : Opcion de Proyecciones de Ahorros por Agencia segun ERS165-2013 - RQ13825
'** Creación : FRHU, 20140120 09:00:00 AM
'********************************************************************
Private Sub btnCancelar_Click()
    Me.txtAnio.Text = Year(Now)
    Call MostrarFlexDefault
    Me.btnGuardar.Enabled = False
    Me.btnCancelar.Enabled = False
    Me.txtAnio.Enabled = True
    Me.btnSeleccionar.Enabled = True
    Me.txtAnio.Text = ""
    Me.txtAnio.SetFocus
    Me.feProyeccion.Enabled = False
End Sub
Private Sub btnGuardar_Click()
    Dim oCaja As New nCajaGeneral
    Dim rs As New ADODB.Recordset
    Dim lcAnio As String
    lcAnio = Me.txtAnio.Text
    If MsgBox("¿Esta seguro de guardar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    Set rs = feProyeccion.GetRsNew
    Call oCaja.GrabarProyeccionAhorroAge(lcAnio, rs)
    MsgBox "Se ha guardaron satisfactoriamente los montos del año " & lnAnio, vbInformation, "Aviso"
    
    btnSeleccionar_Click
    Set oCaja = Nothing
    Exit Sub
ErrProcesar:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub btnSalir_Click()
Unload Me
End Sub
Private Sub btnSeleccionar_Click()
Dim oCaja As New DCajaGeneral
Dim rs As New ADODB.Recordset
Dim lnFila As Integer

If txtAnio.Text = "" Then
    MsgBox "Ingrese un Año", vbInformation, "Advertencia"
    Exit Sub
End If
'validando año
If Val(txtAnio.Text) < 1900 Or Val(txtAnio.Text) > 9972 Then
    MsgBox "Año no Valido", vbInformation, "Advertencia"
    Exit Sub
End If

Me.feProyeccion.Enabled = True
Me.btnGuardar.Enabled = True
Me.btnCancelar.Enabled = True
Me.txtAnio.Enabled = False
Me.btnSeleccionar.Enabled = False

Set rs = oCaja.RecuperaProyeccionAhorroAgexAnio(Me.txtAnio.Text)
Call LimpiaFlex(feProyeccion)
Do While Not rs.EOF
    feProyeccion.AdicionaFila
    lnFila = feProyeccion.row
    feProyeccion.TextMatrix(lnFila, 1) = rs!cAgecod
    feProyeccion.TextMatrix(lnFila, 2) = rs!cAgeDescripcion
    feProyeccion.TextMatrix(lnFila, 3) = rs!cAnio
    feProyeccion.TextMatrix(lnFila, 4) = Format(rs!Enero, gsFormatoNumeroView)
    feProyeccion.TextMatrix(lnFila, 5) = Format(rs!Febrero, gsFormatoNumeroView)
    feProyeccion.TextMatrix(lnFila, 6) = Format(rs!Marzo, gsFormatoNumeroView)
    feProyeccion.TextMatrix(lnFila, 7) = Format(rs!Abril, gsFormatoNumeroView)
    feProyeccion.TextMatrix(lnFila, 8) = Format(rs!Mayo, gsFormatoNumeroView)
    feProyeccion.TextMatrix(lnFila, 9) = Format(rs!Junio, gsFormatoNumeroView)
    feProyeccion.TextMatrix(lnFila, 10) = Format(rs!Julio, gsFormatoNumeroView)
    feProyeccion.TextMatrix(lnFila, 11) = Format(rs!Agosto, gsFormatoNumeroView)
    feProyeccion.TextMatrix(lnFila, 12) = Format(rs!Septiembre, gsFormatoNumeroView)
    feProyeccion.TextMatrix(lnFila, 13) = Format(rs!Octubre, gsFormatoNumeroView)
    feProyeccion.TextMatrix(lnFila, 14) = Format(rs!Noviembre, gsFormatoNumeroView)
    feProyeccion.TextMatrix(lnFila, 15) = Format(rs!Diciembre, gsFormatoNumeroView)
    rs.MoveNext
Loop
End Sub
Private Sub Form_Load()
Me.txtAnio.Text = Year(Now)
Call MostrarFlexDefault
Me.btnGuardar.Enabled = False
Me.btnCancelar.Enabled = False
Me.feProyeccion.Enabled = False
End Sub
Private Sub MostrarFlexDefault()
Dim oCaja As New DCajaGeneral
Dim rs As New ADODB.Recordset
Set rs = oCaja.RecuperaProyeccionAhorroAge()
Call CargarGrillaProyeccionAhorro(rs)
End Sub
Private Sub CargarGrillaProyeccionAhorro(ByVal rs As ADODB.Recordset)
    Dim lnFila As Integer
    Call LimpiaFlex(feProyeccion)
    If Not RSVacio(rs) Then
        rs.MoveFirst
        Do While Not rs.EOF
            feProyeccion.AdicionaFila
            lnFila = feProyeccion.row
            feProyeccion.TextMatrix(lnFila, 1) = rs!cAgecod
            feProyeccion.TextMatrix(lnFila, 2) = rs!cAgeDescripcion
            feProyeccion.TextMatrix(lnFila, 3) = rs!cAnio
            feProyeccion.TextMatrix(lnFila, 4) = Format(rs!Enero, gsFormatoNumeroView)
            feProyeccion.TextMatrix(lnFila, 5) = Format(rs!Febrero, gsFormatoNumeroView)
            feProyeccion.TextMatrix(lnFila, 6) = Format(rs!Marzo, gsFormatoNumeroView)
            feProyeccion.TextMatrix(lnFila, 7) = Format(rs!Abril, gsFormatoNumeroView)
            feProyeccion.TextMatrix(lnFila, 8) = Format(rs!Mayo, gsFormatoNumeroView)
            feProyeccion.TextMatrix(lnFila, 9) = Format(rs!Junio, gsFormatoNumeroView)
            feProyeccion.TextMatrix(lnFila, 10) = Format(rs!Julio, gsFormatoNumeroView)
            feProyeccion.TextMatrix(lnFila, 11) = Format(rs!Agosto, gsFormatoNumeroView)
            feProyeccion.TextMatrix(lnFila, 12) = Format(rs!Septiembre, gsFormatoNumeroView)
            feProyeccion.TextMatrix(lnFila, 13) = Format(rs!Octubre, gsFormatoNumeroView)
            feProyeccion.TextMatrix(lnFila, 14) = Format(rs!Noviembre, gsFormatoNumeroView)
            feProyeccion.TextMatrix(lnFila, 15) = Format(rs!Diciembre, gsFormatoNumeroView)

            'If rs!nOrden = 0 Then
            '   feProyeccion.BackColorRow &HE0E0E0, True
            'End If
            rs.MoveNext
        Loop
    End If
End Sub
Private Sub txtAnio_GotFocus()
fEnfoque txtAnio
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
Dim sFecha  As Date
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
        'cmdGenerar.SetFocus
End If
End Sub
