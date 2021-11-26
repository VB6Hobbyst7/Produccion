VERSION 5.00
Begin VB.Form frmLogExaminaCBSO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Titulo"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8955
   Icon            =   "frmLogExaminaCBSO.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
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
      Left            =   6720
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
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
      Left            =   7800
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
   Begin Sicmact.FlexEdit feContrato 
      Height          =   2340
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   4128
      Cols0           =   9
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-N° Contrato-Proveedor-Moneda-Monto-Desde-Hasta-N° Cuotas-nContRef"
      EncabezadosAnchos=   "400-1200-2300-1000-1000-900-900-900-0"
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
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L-L-C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-2-0-0-0-0"
      TextArray0      =   "#"
      SelectionMode   =   1
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   7
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmLogExaminaCBSO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fnTipoContrato As Integer
Dim fsAreaAgeCod As String
Dim fnMoneda As Integer
Dim fsMatrizDatos() As String
Dim oLog As DLogGeneral
Public Function Inicio(ByVal pnTipoContrato As Integer, ByVal psAreaAgeCod, ByVal pnMoneda As Integer) As String()
    fsAreaAgeCod = psAreaAgeCod
    fnTipoContrato = pnTipoContrato
    fnMoneda = pnMoneda
    Me.Show 1
    Inicio = fsMatrizDatos
End Function
Private Sub cmdAceptar_Click()
     If feContrato.TextMatrix(1, 1) = "" Then Exit Sub
    fsMatrizDatos(1, 1) = feContrato.TextMatrix(feContrato.row, 8)
    fsMatrizDatos(2, 1) = feContrato.TextMatrix(feContrato.row, 1)
    fsMatrizDatos(3, 1) = feContrato.TextMatrix(feContrato.row, 2)
    Unload Me
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub feCrontratos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdAceptar.SetFocus
    End If
End Sub
Private Sub Form_Load()
    Set oLog = New DLogGeneral
    Dim lstitulo As String
    Select Case fnTipoContrato
        Case LogTipoContrato.ContratoObra
            lstitulo = "CONTRATO DE OBRA"
        Case LogTipoContrato.ContratoServicio
            lstitulo = "CONTRATO DE SERVICIO"
        Case LogTipoContrato.ContratoArrendamiento
            lstitulo = "CONTRATO DE ARRENDAMIENTO"
        Case LogTipoContrato.ContratoAdqBienes
            lstitulo = "CONTRATO DE ADQUISICION DE BIENES"
    End Select
    Me.Caption = "BUSQUEDA DE " & lstitulo
    CargarDatos
    ReDim fsMatrizDatos(3, 1)
End Sub
Private Sub CargarDatos()
Dim rsLog As ADODB.Recordset
Dim row As Long
Set rsLog = oLog.ListaContratoxRegistroComprobante(fsAreaAgeCod, fnTipoContrato, fnMoneda)
Call LimpiaFlex(Me.feContrato)
If rsLog.RecordCount > 0 Then
    LimpiaFlex feContrato
            Do While Not rsLog.EOF
                feContrato.AdicionaFila
                row = feContrato.row
                feContrato.TextMatrix(row, 1) = rsLog!cNContrato
                feContrato.TextMatrix(row, 2) = rsLog!cPersNombre
                feContrato.TextMatrix(row, 3) = rsLog!cMoneda
                feContrato.TextMatrix(row, 4) = Format(rsLog!nMonto, gsFormatoNumeroView)
                feContrato.TextMatrix(row, 5) = Format(rsLog!dFechaIni, gsFormatoFechaView)
                feContrato.TextMatrix(row, 6) = Format(rsLog!dFechaFin, gsFormatoFechaView)
                feContrato.TextMatrix(row, 7) = rsLog!NCuotas
                feContrato.TextMatrix(row, 8) = rsLog!cNContRef
                rsLog.MoveNext
            Loop
            If rsLog.RecordCount > 0 Then
                feContrato.TabIndex = 0
                cmdAceptar.Default = True
            Else
                cmdAceptar.Default = False
            End If
            SendKeys "{Right}"
Else
    MsgBox "No hay datos.", vbInformation, "Aviso"
End If
End Sub
Private Function ValidarSeleccion() As Boolean
If Trim(Me.feContrato.TextMatrix(1, 1)) = "" Then
    MsgBox "No hay datos.", vbInformation, "Aviso"
    ValidarSeleccion = False
    Exit Function
Else
    If Trim(Me.feContrato.TextMatrix(Me.feContrato.row, 1)) = "" Then
        MsgBox "Seleccione correctamente el Registro.", vbInformation, "Aviso"
        ValidarSeleccion = False
        Exit Function
    End If
End If
ValidarSeleccion = True
End Function

