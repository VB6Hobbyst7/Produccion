VERSION 5.00
Begin VB.Form frmCCETranfInterBancaExtorno 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extorno de Tranferencia Interbancaria - Originante"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16800
   Icon            =   "frmCCETranfInterBancaExtorno.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   16800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   14160
      TabIndex        =   9
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   12840
      TabIndex        =   8
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Frame fraGlosa 
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
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   3840
      Width           =   12375
      Begin VB.TextBox txtGlosa 
         Height          =   405
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   12135
      End
   End
   Begin SICMACT.FlexEdit feTransferencia 
      Height          =   2775
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   16575
      _ExtentX        =   29236
      _ExtentY        =   4895
      Cols0           =   10
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Banco Destino-Ordenante-Moneda-Importe-Cuenta de Cargo-Fecha y Hora-Tpo. Transferencia-nTranNro-Movimiento"
      EncabezadosAnchos=   "300-3500-3500-1000-1000-2000-1800-2800-0-2800"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-C-R-R-C-L-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Frame fraBuscar 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      Begin VB.CheckBox chkNroMov 
         Caption         =   "Nº Mov:"
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
         Left            =   5760
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton btnBuscar 
         Caption         =   "Buscar"
         Height          =   345
         Left            =   8160
         TabIndex        =   4
         Top             =   300
         Width           =   1335
      End
      Begin VB.TextBox txtNroMov 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6840
         TabIndex        =   3
         Top             =   320
         Width           =   1215
      End
      Begin VB.ComboBox cboTpoTransf 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   3375
      End
      Begin VB.Label lblTpoTransf 
         Caption         =   "Tipo de Transferencia:"
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
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmCCETranfInterBancaExtorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************
'** Nombre : frmCCETranfInterBancaExtorno
'** Descripción : Para el extorno de transferencias CCE , Proyecto: Implementacion del Servicio de Compensaciòn Electrónica Diferido de Instrumentos Compensables CCE
'** Creación : PASI, 20160824
'**********************************************************************
Option Explicit
Dim oCCE As COMNCajaGeneral.NCOMCCE
Private clsprevio As New previo.clsprevio 'vapa20170708
Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Not Len(txtGlosa.Text) = 0 Then
        cmdGrabar.SetFocus
    End If
End Sub
Private Sub txtNroMov_GotFocus()
    fEnfoque txtNroMov
End Sub
Private Sub btnBuscar_Click()
    Dim rs As ADODB.Recordset
    If chkNroMov.value And (Len(Trim(txtNroMov.Text)) = 0 Or txtNroMov.Text = "0") Then
        MsgBox "Indique un número de movimiento válido.", vbInformation, "¡Aviso!"
        txtNroMov.SetFocus
    End If
    LimpiaFlex feTransferencia
    Set rs = oCCE.CCE_ListaTransfParaExtorno(gdFecSis, IIf(cboTpoTransf.ListIndex = -1, "%", Trim(Right(cboTpoTransf.Text, 3))), IIf(Len(txtNroMov.Text) = 0, 0, txtNroMov.Text))
    If rs.EOF And rs.BOF Then
        MsgBox "No hay información para mostrar.", vbInformation, "¡Aviso!"
        btnBuscar.SetFocus
        Exit Sub
    End If
    Do While Not rs.EOF
        feTransferencia.AdicionaFila
        feTransferencia.TextMatrix(feTransferencia.row, 1) = rs!cEfinPersNombre
        feTransferencia.TextMatrix(feTransferencia.row, 2) = rs!cTranNombreOrd
        feTransferencia.TextMatrix(feTransferencia.row, 3) = rs!nmoneda
        feTransferencia.TextMatrix(feTransferencia.row, 4) = Format(rs!nTranMonto, "#,##0.00")
        feTransferencia.TextMatrix(feTransferencia.row, 5) = rs!cTrorNroCCIBenef
        feTransferencia.TextMatrix(feTransferencia.row, 6) = rs!FechayHora
        feTransferencia.TextMatrix(feTransferencia.row, 7) = rs!cOptrDesc
        feTransferencia.TextMatrix(feTransferencia.row, 8) = rs!nTranNro
        feTransferencia.TextMatrix(feTransferencia.row, 9) = rs!nMovNro
        rs.MoveNext
    Loop
End Sub
Private Sub chkNroMov_Click()
    cboTpoTransf.ListIndex = -1
    txtNroMov.Text = ""
    cboTpoTransf.Enabled = IIf(chkNroMov.value, False, True)
    txtNroMov.Enabled = IIf(chkNroMov.value, True, False)
    If chkNroMov.value Then txtNroMov.SetFocus
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub cmdGrabar_Click()
Dim rsVoucher As New ADODB.Recordset 'vapa20170708
Dim oNCapMov As New COMNCaptaGenerales.NCOMCaptaMovimiento 'vapa20170708
Dim sBoleta As String 'vapa20170708
Dim lnMovNro As Long
Dim lnMovNroOrigen As Long


    If feTransferencia.TextMatrix(feTransferencia.row, 1) = "" Then
        MsgBox "No existen transferencia a extornar", vbInformation, "¡Aviso!"
        Exit Sub
    End If
    If Len(Trim(txtGlosa.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la Glosa del extorno", vbInformation, "¡Aviso!"
        txtGlosa.SetFocus
        Exit Sub
    End If
    If MsgBox("¿Está seguro de extornar la Transferencia?", vbQuestion + vbYesNo + vbDefaultButton1, "¡Aviso!") = vbNo Then
        Exit Sub
    End If
    'vapa20170708
   lnMovNro = oCCE.CCE_ExtornaTransf(gdFecSis, gsCodAge, gsCodUser, gsOpeCod, Trim(txtGlosa.Text), feTransferencia.TextMatrix(feTransferencia.row, 8))
   lnMovNroOrigen = feTransferencia.TextMatrix(feTransferencia.row, 9)
    
    MsgBox "Se ha extornado con éxito la Transferencia.", vbInformation, "¡Aviso!"
    Set rsVoucher = oCCE.DatosVoucherTransferenciaExtorno(lnMovNroOrigen)
    sBoleta = oNCapMov.ImprimeVoucherExtornoTransf(rsVoucher, gbImpTMU, gsCodUser, gsOpeCod)
    
            Do
                clsprevio.PrintSpool sLpt, sBoleta
                
            Loop While MsgBox("Desea Reimprimir el voucher?", vbInformation + vbYesNo, "¡Aviso!") = vbYes
            If MsgBox("¿Desea Extornar otra Transferencia ?", vbYesNo + vbInformation, "¡Aviso!") = vbNo Then
                Unload Me
                Exit Sub
            Else
                  
                  txtGlosa.Text = ""
                  btnBuscar_Click
            End If
'    txtGlosa.Text = ""
'    btnBuscar_Click
End Sub
Private Sub Form_Load()
    Dim rsTpotransf As ADODB.Recordset
    Set oCCE = New COMNCajaGeneral.NCOMCCE
    Set rsTpotransf = oCCE.CCE_ListaOpeTransf
    Do While Not rsTpotransf.EOF
        cboTpoTransf.AddItem Mid(rsTpotransf!cOptrDesc & Space(100), 1, 100) & rsTpotransf!cOptrCod
        rsTpotransf.MoveNext
    Loop
    cboTpoTransf.AddItem Mid("-------------TODOS-------------" & Space(100), 1, 100) & "%"
    IniControles
End Sub
Private Sub IniControles()
    cboTpoTransf.ListIndex = -1
    chkNroMov.value = 0
    txtNroMov.Text = ""
    txtNroMov.Enabled = False
End Sub
Private Sub txtNroMov_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 And Not Len(txtNroMov) = 0 Then
        btnBuscar.SetFocus
    End If
End Sub
