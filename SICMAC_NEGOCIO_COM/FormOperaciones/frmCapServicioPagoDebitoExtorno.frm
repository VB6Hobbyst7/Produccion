VERSION 5.00
Begin VB.Form frmCapServicioPagoDebitoExtorno 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Servicio - Extorno Servicio de Pago"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16170
   Icon            =   "frmCapServicioPagoDebitoExtorno.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   16170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExtornar 
      Caption         =   "&Extornar"
      Height          =   375
      Left            =   15120
      TabIndex        =   6
      Top             =   5760
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5760
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   14040
      TabIndex        =   8
      Top             =   5760
      Width           =   1035
   End
   Begin VB.Frame fraMovimientos 
      Caption         =   "Movimientos"
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
      Height          =   4215
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   16035
      Begin SICMACT.FlexEdit FEExtorno 
         Height          =   3855
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   15795
         _ExtentX        =   27861
         _ExtentY        =   6800
         Cols0           =   14
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   $"frmCapServicioPagoDebitoExtorno.frx":030A
         EncabezadosAnchos=   "250-2200-2000-1600-1200-3500-1200-1700-1700-0-0-0-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-C-R-L-R-R-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-2-2-2-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   255
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame FRGlosa 
      Caption         =   "Glosa"
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
      Height          =   1335
      Left            =   7560
      TabIndex        =   12
      Top             =   0
      Width           =   8535
      Begin VB.TextBox txtGlosa 
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   8295
      End
   End
   Begin VB.Frame FRBuscar 
      Caption         =   "Datos Búsqueda"
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
      Height          =   1335
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   7395
      Begin VB.Frame FRTipo 
         Caption         =   "Tipo"
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
         Height          =   975
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1995
         Begin VB.OptionButton optTipoBus 
            Caption         =   "&Número Movimiento"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   0
            Top             =   240
            Value           =   -1  'True
            Width           =   1755
         End
         Begin VB.OptionButton optTipoBus 
            Caption         =   "Número de &Cuenta"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   1
            Top             =   540
            Width           =   1755
         End
      End
      Begin VB.Frame FRCueNro 
         Height          =   975
         Left            =   2220
         TabIndex        =   10
         Top             =   240
         Width           =   4995
         Begin VB.TextBox txtNroMovimiento 
            Height          =   375
            Left            =   840
            MaxLength       =   8
            TabIndex        =   2
            Top             =   360
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "&Buscar"
            Height          =   375
            Left            =   3840
            TabIndex        =   4
            Top             =   420
            Width           =   1035
         End
         Begin SICMACT.ActXCodCta txtCuenta 
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Visible         =   0   'False
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   661
            Texto           =   "Cuenta N°"
            EnabledCMAC     =   -1  'True
            EnabledCta      =   -1  'True
            EnabledProd     =   -1  'True
            EnabledAge      =   -1  'True
         End
         Begin VB.Label lblNroMov 
            Caption         =   "# Mov :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Visible         =   0   'False
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frmCapServicioPagoDebitoExtorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'*** Nombre : frmCapServicioPagoDebitoExtorno
'*** Descripción : Formulario para extornar el debito que se pago del convenio.
'*** Creación : ELRO el 20130715 08:28:01 AM, según RFC1306270002
'********************************************************************
Option Explicit

Private Sub imprimirBoleta(ByVal psBoleta As String, Optional ByVal psMensaje As String = "Boleta de extorno")
Dim lnFicSal As Integer
Do
    lnFicSal = FreeFile
    Open sLpt For Output As lnFicSal
    Print #lnFicSal, psBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    'Print #nFicSal, ""
    Close #lnFicSal
Loop Until MsgBox("¿Desea Re-Imprimir " & psMensaje & " ?", vbQuestion + vbYesNo, "Aviso") = vbNo
End Sub

Private Sub LimpiarCampos()
optTipoBus.Item(0).value = True
lblNroMov.Visible = True
txtNroMovimiento.Visible = True
txtCuenta.Visible = False
txtNroMovimiento = ""
txtCuenta.CMAC = "109"
txtCuenta.EnabledCMAC = False
txtCuenta.Age = ""
txtCuenta.EnabledProd = False
txtCuenta.Prod = "232"
txtCuenta.Cuenta = ""
txtGlosa = ""
LimpiaFlex FEExtorno
End Sub


Private Sub CmdBuscar_Click()
Dim oNCOMCaptaMovimiento As NCOMCaptaMovimiento
Set oNCOMCaptaMovimiento = New NCOMCaptaMovimiento
Dim rsDebitos As ADODB.Recordset
Set rsDebitos = New ADODB.Recordset
Dim lnTipoBusqueda As Integer
Dim lsCuenta As String

txtGlosa = ""
LimpiaFlex FEExtorno

If optTipoBus.Item(0).value Then
    If Len(txtNroMovimiento) = 0 Then Exit Sub
    lnTipoBusqueda = 0
    Set rsDebitos = oNCOMCaptaMovimiento.obtenerConvenioServicioPagoDebitoParaExtornar("", Format(gdFecSis, "yyyyMMdd"), gCapConSerPagDeb, gsCodAge, lnTipoBusqueda, CLng(txtNroMovimiento))
Else
    lnTipoBusqueda = 1
    lsCuenta = txtCuenta.CMAC & txtCuenta.Age & txtCuenta.Prod & txtCuenta.Cuenta
    If Len(lsCuenta) < 18 Then Exit Sub
    Set rsDebitos = oNCOMCaptaMovimiento.obtenerConvenioServicioPagoDebitoParaExtornar(lsCuenta, Format(gdFecSis, "yyyyMMdd"), gCapConSerPagDeb, gsCodAge, lnTipoBusqueda, 0)
End If

If Not (rsDebitos.BOF And rsDebitos.EOF) Then
    FEExtorno.lbEditarFlex = True
    Do While Not rsDebitos.EOF
        FEExtorno.AdicionaFila
        FEExtorno.TextMatrix(FEExtorno.row, 1) = rsDebitos!cMovNro
        FEExtorno.TextMatrix(FEExtorno.row, 2) = rsDebitos!cOpedesc
        FEExtorno.TextMatrix(FEExtorno.row, 3) = rsDebitos!cCtaCod
        FEExtorno.TextMatrix(FEExtorno.row, 4) = Format$(rsDebitos!nMonto, "##,##0.00") 'Format$(rsDebito!nMonto, "##,##0.00")
        FEExtorno.TextMatrix(FEExtorno.row, 5) = rsDebitos!cMovDesc
        FEExtorno.TextMatrix(FEExtorno.row, 6) = Format$(rsDebitos!ITFCargo, "##,##0.00")
        FEExtorno.TextMatrix(FEExtorno.row, 7) = Format$(rsDebitos!ComSerPagAge, "##,##0.00")
        FEExtorno.TextMatrix(FEExtorno.row, 8) = Format$(rsDebitos!PenSerPagAge, "##,##0.00")
        FEExtorno.TextMatrix(FEExtorno.row, 9) = rsDebitos!nMovNro
        FEExtorno.TextMatrix(FEExtorno.row, 10) = rsDebitos!IdSerPag
        FEExtorno.TextMatrix(FEExtorno.row, 11) = rsDebitos!cBeneficiario
        FEExtorno.TextMatrix(FEExtorno.row, 12) = rsDebitos!cPersIDnro
        FEExtorno.TextMatrix(FEExtorno.row, 13) = rsDebitos!cEmpresa
        rsDebitos.MoveNext
    Loop
    FEExtorno.lbEditarFlex = False
End If

End Sub

Private Sub cmdCancelar_Click()
LimpiarCampos
End Sub

Private Sub cmdExtornar_Click()
Dim lbResultadoVisto As Boolean
Dim loVistoElectronico As frmVistoElectronico
Set loVistoElectronico = New frmVistoElectronico
Dim oDCOMCaptaMovimiento As DCOMCaptaMovimiento
Set oDCOMCaptaMovimiento = New DCOMCaptaMovimiento
Dim lnMovNroBus, lnMovNroUlt As Long
Dim lsCuenta As String

If Len(txtGlosa) = 0 Then
    MsgBox "Falta ingresar la glosa.", vbInformation, "Aviso"
    Exit Sub
End If
 
lnMovNroBus = CLng(FEExtorno.TextMatrix(FEExtorno.row, 9))
lsCuenta = FEExtorno.TextMatrix(FEExtorno.row, 3)
lnMovNroUlt = oDCOMCaptaMovimiento.devolverUltimoMovimientoDeposito(lsCuenta, Format(gdFecSis, "yyyyMMdd"))
If lnMovNroUlt > 0 And lnMovNroBus <> lnMovNroUlt Then
    MsgBox "Cuenta " & lsCuenta & " posee movimientos después del debito, por favor extorne antes esos movimientos.", vbInformation, "Aviso"
    Set oDCOMCaptaMovimiento = Nothing
    Exit Sub
End If
Set oDCOMCaptaMovimiento = Nothing

' *** RIRO SEGUN TI-ERS108-2013 ***
    Dim nMovNroOperacion As Long
    nMovNroOperacion = 0
    If FEExtorno.row >= 1 And Len(Trim(FEExtorno.TextMatrix(FEExtorno.row, 9))) > 0 Then
        nMovNroOperacion = Val(FEExtorno.TextMatrix(FEExtorno.row, 9))
    End If
' *** FIN RIRO ***

lbResultadoVisto = loVistoElectronico.Inicio(3, gCapExtConSerPagDeb, , , nMovNroOperacion) 'RIRO SEGUN TI-ERS108-2013/ Se agrego parametro nMovNroOperacion

If Not lbResultadoVisto Then
    Set loVistoElectronico = Nothing
    Exit Sub
End If
 
If MsgBox("¿Esta seguro que desea exornar el movimiento " & FEExtorno.TextMatrix(FEExtorno.row, 1) & "?", vbYesNo, "Aviso") = vbYes Then
    Dim oNCOMCaptaMovimiento As NCOMCaptaMovimiento
    Set oNCOMCaptaMovimiento = New NCOMCaptaMovimiento
    Dim oNCOMContFunciones As NCOMContFunciones
    Set oNCOMContFunciones = New NCOMContFunciones
    Dim lsMovNro As String
    Dim lnConfirmar As Long
    Dim lsBoleta As String
    
    lsMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    lnConfirmar = oNCOMCaptaMovimiento.extornarDebitarCuentaConvenioServicioPago(lsMovNro, gCapExtConSerPagDeb, _
                                                                                 CLng(FEExtorno.TextMatrix(FEExtorno.row, 10)), _
                                                                                 CCur(FEExtorno.TextMatrix(FEExtorno.row, 4)), _
                                                                                 CCur(FEExtorno.TextMatrix(FEExtorno.row, 6)), _
                                                                                 CCur(FEExtorno.TextMatrix(FEExtorno.row, 7)), _
                                                                                 CCur(FEExtorno.TextMatrix(FEExtorno.row, 8)), _
                                                                                 gsNomCmac, gsNomAge, _
                                                                                 FEExtorno.TextMatrix(FEExtorno.row, 13), _
                                                                                 FEExtorno.TextMatrix(FEExtorno.row, 12), _
                                                                                 FEExtorno.TextMatrix(FEExtorno.row, 11), _
                                                                                 lsBoleta, sLpt, _
                                                                                 FEExtorno.TextMatrix(FEExtorno.row, 1), _
                                                                                 CLng(FEExtorno.TextMatrix(FEExtorno.row, 9)), _
                                                                                 gbImpTMU)
    
End If

If lnConfirmar > 0 Then
    imprimirBoleta lsBoleta
    loVistoElectronico.RegistraVistoElectronico (lnMovNroBus)
    LimpiarCampos
End If

Set loVistoElectronico = Nothing
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
optTipoBus.Item(0).value = True
lblNroMov.Visible = True
txtNroMovimiento.Visible = True
txtCuenta.Visible = False
txtNroMovimiento = ""
txtCuenta.CMAC = "109"
txtCuenta.EnabledCMAC = False
txtCuenta.Age = ""
txtCuenta.EnabledProd = False
txtCuenta.Prod = "232"
txtCuenta.Cuenta = ""
End Sub

Private Sub optTipoBus_Click(Index As Integer)
If optTipoBus.Item(0).value Then
    lblNroMov.Visible = True
    txtNroMovimiento.Visible = True
Else
    lblNroMov.Visible = False
    txtNroMovimiento.Visible = False
End If

If optTipoBus.Item(1).value Then
    txtCuenta.Visible = True
Else
    txtCuenta.Visible = False
End If
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBuscar.SetFocus
End If
End Sub

Private Sub txtNroMovimiento_KeyPress(KeyAscii As Integer)

If Not IsNumeric(txtNroMovimiento) Then
    txtNroMovimiento = ""
    Exit Sub
End If

If KeyAscii = 13 Then
    cmdBuscar.SetFocus
End If

End Sub
