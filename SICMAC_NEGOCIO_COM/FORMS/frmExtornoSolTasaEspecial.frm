VERSION 5.00
Begin VB.Form frmExtornoSolTasaEspecial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EXTORNO DE  SOLICITUD  DE  TASA PREFERENCIAL"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11010
   Icon            =   "frmExtornoSolTasaEspecial.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton opt1 
      Caption         =   "Rechazadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   2
      Left            =   9315
      TabIndex        =   11
      Top             =   960
      Width           =   1545
   End
   Begin VB.Frame Frame1 
      Height          =   3060
      Left            =   75
      TabIndex        =   9
      Top             =   1290
      Width           =   10845
      Begin SICMACT.FlexEdit grdExtorno 
         Height          =   2685
         Left            =   75
         TabIndex        =   10
         Top             =   255
         Width           =   10665
         _ExtentX        =   18812
         _ExtentY        =   4736
         Cols0           =   18
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Nro Sol.-Sub Producto-Moneda-Plazo-Tasa-Monto-Estado-nmoneda-nestado-nproducto-Persona------"
         EncabezadosAnchos=   "0-1000-2000-1200-1000-1000-1200-0-0-0-0-4000-0-0-0-0-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-C-C-R-C-C-C-C-C-L-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-2-0-0-0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.OptionButton opt1 
      Caption         =   "Aprobadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   1
      Left            =   7665
      TabIndex        =   8
      Top             =   960
      Value           =   -1  'True
      Width           =   1245
   End
   Begin VB.OptionButton opt1 
      Caption         =   "Solicitada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   0
      Left            =   5970
      TabIndex        =   7
      Top             =   975
      Width           =   1350
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Salir"
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
      Left            =   6195
      TabIndex        =   6
      Top             =   4425
      Width           =   1155
   End
   Begin VB.CommandButton cmdExtornar 
      Caption         =   "&Extornar"
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
      Left            =   4050
      TabIndex        =   5
      Top             =   4425
      Width           =   1155
   End
   Begin VB.CommandButton cmdBuscar 
      Height          =   405
      Left            =   7500
      Picture         =   "frmExtornoSolTasaEspecial.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   405
      Width           =   525
   End
   Begin VB.Label lblDI 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1305
      TabIndex        =   4
      Top             =   945
      Width           =   1890
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "DNI / RUC:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   210
      TabIndex        =   3
      Top             =   990
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   195
      TabIndex        =   2
      Top             =   495
      Width           =   720
   End
   Begin VB.Label lblCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1290
      TabIndex        =   1
      Top             =   435
      Width           =   6105
   End
End
Attribute VB_Name = "frmExtornoSolTasaEspecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'By capi 21012009
Option Explicit
Dim objPista As COMManejador.Pista
'End by

Private Sub CmdBuscar_Click()
Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim lrContratos As ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona

On Error GoTo ControlError

grdExtorno.Clear
lblCliente.Caption = ""
lblCliente.Tag = ""
lblDI.Caption = ""
Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lblCliente.Tag = loPers.sPersCod
    lblCliente.Caption = loPers.sPersNombre
    lblDI.Caption = IIf(loPers.sPersPersoneria = 1, loPers.sPersIdnroDNI, loPers.sPersIdnroRUC)
Set loPers = Nothing
    Cabecera
    CargaGrid

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub
Private Sub Cabecera()
With grdExtorno
        .TextMatrix(0, 0) = "#"
        .ColWidth(0) = 0
        .TextMatrix(0, 1) = "Nro Solicitud"
        .ColWidth(1) = 1200
        .TextMatrix(0, 2) = "Producto"
        .ColWidth(2) = 1200
        .TextMatrix(0, 3) = "Moneda"
        .ColWidth(3) = 1200
        .TextMatrix(0, 4) = "Plazo"
        .ColWidth(4) = 1200
        .TextMatrix(0, 5) = "Tasa"
        .ColWidth(5) = 1200
        .TextMatrix(0, 6) = "Monto"
        .ColWidth(6) = 1200
        .TextMatrix(0, 7) = "Estado"
        .ColWidth(5) = 1400
        .TextMatrix(0, 8) = "nmoneda"
        .ColWidth(8) = 0
        .TextMatrix(0, 9) = "nestado"
        .ColWidth(9) = 0
        .TextMatrix(0, 10) = "nproducto"
        .ColWidth(10) = 0
        .TextMatrix(0, 11) = "Persona"
        .ColWidth(10) = 4000
        .TextMatrix(0, 12) = "cperscod"
        .ColWidth(10) = 0
        .TextMatrix(0, 13) = "bpermanente"
        .ColWidth(10) = 0
    End With
End Sub


Private Sub CargaGrid()
Dim ssql As String
Dim sValor As String
Dim rstemp As ADODB.Recordset
Dim oCapG As COMDCaptaGenerales.DCOMCaptaGenerales
Dim i As Double
Set rstemp = New ADODB.Recordset

Set oCapG = New COMDCaptaGenerales.DCOMCaptaGenerales
grdExtorno.Clear
Cabecera
grdExtorno.Rows = 2
If opt1(0).value = True Then
   sValor = "0"
   Set rstemp = oCapG.ObtenerListaSolicitudes(lblCliente.Tag, sValor)
  
ElseIf opt1(1).value = True Then
   sValor = "1"
   Set rstemp = oCapG.ObtenerListaSolicitudes(lblCliente.Tag, sValor)
  
ElseIf opt1(2).value = True Then
   sValor = "2"
   Set rstemp = oCapG.ObtenerListaSolicitudes(lblCliente.Tag, sValor)
  
End If
If Not rstemp Is Nothing Then
    Set grdExtorno.Recordset = rstemp
    rstemp.Close
        
    For i = 1 To grdExtorno.Rows - 1
        grdExtorno.TextMatrix(i, 5) = _
        Format(grdExtorno.TextMatrix(i, 5), "#,######0.000000")
    Next
        
End If
Set oCapG = Nothing
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

'Agregado por RIRO el 20130411
Private Function validarUsuario() As Boolean

    Dim b As Boolean
    
    If gsCodUser = grdExtorno.TextMatrix(grdExtorno.row, 14) Then
        b = True
    End If
    
    validarUsuario = b
    
End Function

'Modificado por RIRO el 20130411
Private Sub cmdExtornar_Click()
Dim nTasa As Double, nMonto As Double
Dim nProd As COMDConstantes.Producto, nMon As COMDConstantes.Moneda
Dim sComent As String, sPersona As String
Dim nPlazo As Integer, bPermanente As Boolean
Dim nFila As Long
Dim nTasaTarif As Double, sSubProducto As String, nTasaSolicitada As Double ' 20130411RIRO

' 20130411RIRO **
If Not validarUsuario() Then
    MsgBox "Usted solo puede extornar solicitudes/aprobaciones/rechazos que usted haya registrado", vbInformation, "Aviso"
    Exit Sub
End If
'END RIRO **

If grdExtorno.TextMatrix(1, 1) <> "" Then

If MsgBox("¿Desea EXTORNAR la solicitud de aprobación de Tasa Especial", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    
    Dim oCont As COMNContabilidad.NCOMContFunciones
    Dim sMovNro As String
    Dim oserv As COMDCaptaServicios.DCOMCaptaServicios
    Dim nNumSolicitud As Long
    
    nFila = grdExtorno.row
    
    Set oCont = New COMNContabilidad.NCOMContFunciones
    sMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set oCont = Nothing
    sPersona = grdExtorno.TextMatrix(nFila, 12)
    nProd = grdExtorno.TextMatrix(nFila, 10)
    nMon = grdExtorno.TextMatrix(nFila, 8)
    nTasa = grdExtorno.TextMatrix(nFila, 5)
    nPlazo = grdExtorno.TextMatrix(nFila, 4)
    nMonto = grdExtorno.TextMatrix(nFila, 6)
    bPermanente = grdExtorno.TextMatrix(nFila, 13)
    
    ' RIRO 20130411
    nTasaTarif = grdExtorno.TextMatrix(nFila, 16)
    sSubProducto = grdExtorno.TextMatrix(nFila, 15)
    nTasaSolicitada = grdExtorno.TextMatrix(nFila, 17)
    ' END RIRO
   
    If Not (grdExtorno.TextMatrix(nFila, 9) = 3) Then
        nTasa = Format$(ConvierteTNAaTEA(nTasa), "#0.0000")
    End If
    
    Set oserv = New COMDCaptaServicios.DCOMCaptaServicios
    nNumSolicitud = grdExtorno.TextMatrix(nFila, 1)
    oserv.AgregaCapTasaEspecial nNumSolicitud, sPersona, nProd, nMon, 3, sMovNro, nTasa, "EXTORNO DE " & IIf(opt1(0).value, "SOLICITUD", IIf(opt1(1).value = True, "APROBACION", "RECHAZO")) & " DE TASA ESPECIAL ", nMonto, , nPlazo, IIf(opt1(0).value = True, 0, IIf(opt1(1).value, 1, 4)), bPermanente, sSubProducto, nTasaTarif, nTasaSolicitada
    'By Capi 21012009
     objPista.InsertarPista gsOpeCod, sMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Extorno", str(nNumSolicitud), gNumeroSolicitud
    'End by

    Set oserv = Nothing
    CargaGrid
    
 End If
 Else
    MsgBox "No existe datos para el extorno", vbInformation, "Aviso"
 End If
End Sub

Private Sub Form_Load()
'By Capi 20012009
    Set objPista = New COMManejador.Pista
    gsOpeCod = gCapExtornoTasasPreferen
'End By

End Sub

Private Sub opt1_Click(Index As Integer)
    CargaGrid
End Sub
