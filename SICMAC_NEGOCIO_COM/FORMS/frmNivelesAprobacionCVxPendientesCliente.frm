VERSION 5.00
Begin VB.Form frmNivelesAprobacionCVxPendientesCliente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Movimientos Autorizados del cliente"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7875
   Icon            =   "frmNivelesAprobacionCVxPendientesCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   6360
      MouseIcon       =   "frmNivelesAprobacionCVxPendientesCliente.frx":030A
      TabIndex        =   0
      Top             =   2040
      Width           =   1455
   End
   Begin SICMACT.FlexEdit feNivApr 
      Height          =   1935
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   3413
      Cols0           =   5
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "nEstado-cMovNro-cNivelCod-nNroFirmas-nContador"
      EncabezadosAnchos=   "0-3500-3000-1200-0"
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
      ColumnasAEditar =   "X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-C-C"
      FormatosEdit    =   "0-0-0-0-0"
      TextArray0      =   "nEstado"
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmNivelesAprobacionCVxPendientesCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsPersCod As String
Dim lsMovNro As String
Dim lsOpeCod As String
Public Function InicioRegistroNiveles(Optional ByVal psPersCod As String = "", Optional ByVal psOpeCod As String = "") As String
    lsPersCod = psPersCod
    lsOpeCod = psOpeCod
    Me.Caption = "Niveles de Aprobación C/V ME - Consulta"
    CargaDatosNiveles
    feNivApr.TopRow = 1
    feNivApr.row = 1
    Me.Show 1
    InicioRegistroNiveles = lsMovNro
    Exit Function
End Function
Private Sub cmdCerrar_Click()
    Unload Me
End Sub
Private Sub CargaDatosNiveles()
Dim lsMovNroFecha As String
    Dim oGen  As COMNContabilidad.NCOMContFunciones
    Set oGen = New COMNContabilidad.NCOMContFunciones

    lsMovNroFecha = oGen.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

    Dim oCont As COMNContabilidad.NCOMContFunciones
    Set oCont = New COMNContabilidad.NCOMContFunciones
    Dim oNiv As COMDCredito.DCOMNivelAprobacion
    Dim rs As ADODB.Recordset
    Dim lnFila As Integer
    Set oNiv = New COMDCredito.DCOMNivelAprobacion
    Set rs = oNiv.ObtenerAprobacionMovCompraVentaPendientexCliente(Mid(lsMovNroFecha, 1, 8), lsPersCod, lsOpeCod)
    Set oNiv = Nothing
    Call LimpiaFlex(feNivApr)
    If Not rs.EOF Then
        Do While Not rs.EOF
            feNivApr.AdicionaFila
            lnFila = feNivApr.row
            feNivApr.TextMatrix(lnFila, 0) = rs!nEstado
            feNivApr.TextMatrix(lnFila, 1) = rs!cMovnro
            feNivApr.TextMatrix(lnFila, 2) = rs!cNivelCod
            feNivApr.TextMatrix(lnFila, 3) = rs!nNroFirmas
            feNivApr.TextMatrix(lnFila, 4) = rs!nContador
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub feNivApr_DblClick()
If Trim(feNivApr.TextMatrix(feNivApr.row, 3)) = Trim(feNivApr.TextMatrix(feNivApr.row, 4)) Then
    lsMovNro = feNivApr.TextMatrix(feNivApr.row, 1)
    Unload Me
Else
    lsMovNro = 0
End If
End Sub

