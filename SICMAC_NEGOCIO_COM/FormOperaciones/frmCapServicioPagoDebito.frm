VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCapServicioPagoDebito 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Servicio - Servicio de Pago"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8655
   Icon            =   "frmCapServicioPagoDebito.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Selección del Convenio"
      TabPicture(0)   =   "frmCapServicioPagoDebito.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FRConvenio"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FRBeneficiario"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdSalir"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdGuardar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdCancelar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5880
         TabIndex        =   9
         Top             =   5640
         Width           =   1095
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "&Guardar"
         Height          =   375
         Left            =   7080
         TabIndex        =   8
         Top             =   5640
         Width           =   1095
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Frame FRBeneficiario 
         Caption         =   "Datos del cliente"
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
         Height          =   3615
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   8175
         Begin VB.TextBox txtITF 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   3120
            Width           =   1215
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   3120
            Width           =   1215
         End
         Begin SICMACT.FlexEdit FEDebitos 
            Height          =   1935
            Left            =   120
            TabIndex        =   19
            Top             =   1080
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   3413
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Nro-Referencia-Importe-Pagar-IdSerPagDeb"
            EncabezadosAnchos=   "500-0-5000-1200-700-0"
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-4-X"
            ListaControles  =   "0-0-0-0-4-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-R-L-R-C-C"
            FormatosEdit    =   "0-0-0-4-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   495
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin VB.TextBox txtNombre 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   915
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   630
            Width           =   6015
         End
         Begin VB.TextBox txtDOI 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   915
            MaxLength       =   8
            TabIndex        =   6
            Top             =   270
            Width           =   1935
         End
         Begin VB.CommandButton cmdDOI 
            Caption         =   "..."
            Height          =   375
            Left            =   2880
            TabIndex        =   5
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblITF 
            Caption         =   "ITF con cargo a cuenta del convenio:"
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
            Left            =   120
            TabIndex        =   22
            Top             =   3120
            Width           =   1815
         End
         Begin VB.Label lblTotal 
            Caption         =   "Total S/.:"
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
            Left            =   4800
            TabIndex        =   21
            Top             =   3240
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Nombre:"
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
            Left            =   120
            TabIndex        =   18
            Top             =   630
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "DOI:"
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
            Left            =   120
            TabIndex        =   17
            Top             =   285
            Width           =   735
         End
      End
      Begin VB.Frame FRConvenio 
         Caption         =   "Busqueda de convenio"
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
         Height          =   1455
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   8175
         Begin VB.TextBox txtEmpresa 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3010
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   960
            Width           =   4040
         End
         Begin VB.TextBox txtConvenio 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   600
            Width           =   6015
         End
         Begin VB.TextBox txtCodigoEmpresa 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox txtCodigoConvenio 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1030
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   240
            Width           =   1935
         End
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "..."
            Height          =   375
            Left            =   3000
            TabIndex        =   0
            Top             =   210
            Width           =   375
         End
         Begin VB.Label lblEmpresa 
            Caption         =   "Empresa:"
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
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   855
         End
         Begin VB.Label lblConvenio 
            Caption         =   "Convenio:"
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
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Código:"
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
            Left            =   120
            TabIndex        =   13
            Top             =   255
            Width           =   735
         End
      End
   End
End
Attribute VB_Name = "frmCapServicioPagoDebito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'*** Nombre : frmCapServicioPagoDebito
'*** Descripción : Formulario para debitar los pagos del convenio.
'*** Creación : ELRO el 20130709 09:59:05 AM, según RFC1306270002
'********************************************************************
Option Explicit
Dim fnIdSerPag As Long
Dim fsNomSerPag As String
Dim fsPersCod As String
Dim fsPersNombre As String
Dim fsCodSerPag As String

Private Sub limpiarCamposBeneficiario()
txtDOI = ""
txtNombre = ""
txtTotal = ""
txtITF = ""
LimpiaFlex FEDebitos
End Sub


Private Sub CmdBuscar_Click()
fnIdSerPag = 0
fsNomSerPag = ""
fsPersCod = ""
fsPersNombre = ""
fsCodSerPag = ""
frmCapServicioPagoBusqueda.iniciarBusqueda fnIdSerPag, fsNomSerPag, fsPersCod, fsPersNombre, fsCodSerPag
txtCodigoConvenio = fsCodSerPag
txtConvenio = fsNomSerPag
txtCodigoEmpresa = fsPersCod
txtEmpresa = fsPersNombre
If Trim(txtCodigoConvenio) <> "" Then
    FRConvenio.Enabled = False
End If
End Sub

Private Sub cargarDebitos()
Dim oNCOMCaptaGenerales As NCOMCaptaGenerales
Set oNCOMCaptaGenerales = New NCOMCaptaGenerales
Dim rsDebito As ADODB.Recordset
Set rsDebito = New ADODB.Recordset

txtNombre = ""
txtITF = ""
txtTotal = ""
LimpiaFlex FEDebitos
Set rsDebito = oNCOMCaptaGenerales.obtenerBeneficiarioConvenioServicioPago(txtDOI, fnIdSerPag)

If Not (rsDebito.BOF And rsDebito.EOF) Then
    txtNombre = rsDebito!cPersNombre
    FEDebitos.lbEditarFlex = True
    FEDebitos.SetFocus
    FEDebitos.lbEditarFlex = True
    Do While Not rsDebito.EOF
        FEDebitos.AdicionaFila
        FEDebitos.TextMatrix(FEDebitos.row, 1) = rsDebito!nNroOpe
        FEDebitos.TextMatrix(FEDebitos.row, 2) = rsDebito!cRefArc
        FEDebitos.TextMatrix(FEDebitos.row, 3) = Format$(rsDebito!nMonto, "##,##0.00")
        FEDebitos.TextMatrix(FEDebitos.row, 4) = "0"
        FEDebitos.TextMatrix(FEDebitos.row, 5) = rsDebito!IdSerPagDeb
        rsDebito.MoveNext
    Loop

Else
    MsgBox "No existe debito a pagar.", vbInformation, "Aviso"
End If

End Sub

Private Sub imprimirBoleta(ByVal psBoleta As String, Optional ByVal psMensaje As String = "Boleta Operación")
Dim lnFicSal As Integer
Do
    lnFicSal = FreeFile
    Open sLpt For Output As lnFicSal
    Print #lnFicSal, psBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    'Print #nFicSal, ""
    Close #lnFicSal
Loop Until MsgBox("¿Desea Re-Imprimir " & psMensaje & " ?", vbQuestion + vbYesNo, "Aviso") = vbNo
End Sub

Private Sub cmdCancelar_Click()
limpiarCamposBeneficiario
txtDOI.Enabled = True
End Sub

Private Sub cmdDOI_Click()
If Trim(txtDOI) = "" Then Exit Sub
cargarDebitos
txtDOI.Enabled = False
End Sub

Private Sub cmdGuardar_Click()

'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, gsOpeCod, "0")
    If Not bPermitirEjecucionOperacion Then
        End
    End If
'fin Comprobacion si es RFIII

If Trim(txtTotal) = "" Then Exit Sub

If CCur(txtTotal) = 0# Then Exit Sub

If MsgBox("¿Esta seguro que desea debitar?", vbYesNo, "Aviso") = vbYes Then
    Dim oNCOMCaptaMovimiento As NCOMCaptaMovimiento
    Set oNCOMCaptaMovimiento = New NCOMCaptaMovimiento
    Dim oNCOMContFunciones As NContFunciones
    Set oNCOMContFunciones = New NContFunciones
    Dim lsBoleta As String
    Dim lsMovNro As String
    Dim lnConfirmar As Long
    
    lsMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    
    lnConfirmar = oNCOMCaptaMovimiento.debitarCuentaConvenioServicioPago(lsMovNro, gCapConSerPagDeb, fnIdSerPag, _
                                                                         CCur(txtTotal), CCur(txtITF), FEDebitos.GetRsNew(), _
                                                                         gsNomCmac, gsNomAge, txtEmpresa, _
                                                                         txtDOI, txtNombre, txtConvenio, lsBoleta, _
                                                                         sLpt, gbImpTMU)
    If lnConfirmar > 0 Then
        imprimirBoleta lsBoleta
        LimpiaFlex FEDebitos
        txtDOI = ""
        txtNombre = ""
        txtITF = "0.00"
        txtTotal = "0.00"
        txtDOI.Enabled = True
    ElseIf lnConfirmar = -1 Then
        MsgBox "Saldo insuficiente en la cuenta del convenio para el Servicio de Pago.", vbInformation, "Aviso"
    ElseIf lnConfirmar = -2 Then
        MsgBox "Convenio no posee cuenta a debitar.", vbInformation, "Aviso"
    End If
End If

End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub FEDebitos_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
Dim i As Integer

txtTotal = "0.00"
For i = 1 To FEDebitos.Rows - 1
    If Trim(FEDebitos.TextMatrix(i, 4)) = "." Then
        txtTotal = Format$(CCur(txtTotal) + CCur(FEDebitos.TextMatrix(i, 3)), "##,##0.00")
    End If
Next i
End Sub

Private Sub FEDebitos_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim lsColumnas() As String
lsColumnas = Split(FEDebitos.ColumnasAEditar, "-")

If lsColumnas(FEDebitos.Col) = "X" Then
    Cancel = False
    SendKeys "{Tab}", True
    Exit Sub
End If
End Sub

Private Sub txtDOI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cargarDebitos
    txtDOI.Enabled = False
End If
End Sub

Private Sub txtTotal_Change()
If Trim(txtTotal) <> "" Then
    Dim nRedondeoITF As Double
    txtITF = Format(fgITFCalculaImpuesto(CCur(txtTotal)), "#,##0.00")
    nRedondeoITF = fgDiferenciaRedondeoITF(CCur(txtITF))
    txtITF = Format(CCur(txtITF) - nRedondeoITF, "#,##0.00")
End If
End Sub
