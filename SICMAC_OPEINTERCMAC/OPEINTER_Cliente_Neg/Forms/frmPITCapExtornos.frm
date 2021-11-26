VERSION 5.00
Begin VB.Form frmPITCapExtornos 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10545
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      ForeColor       =   &H00000080&
      Height          =   1335
      Left            =   7560
      TabIndex        =   15
      Top             =   60
      Width           =   2895
      Begin VB.TextBox txtGlosa 
         Height          =   855
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   300
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   60
      TabIndex        =   8
      Top             =   5820
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9420
      TabIndex        =   7
      Top             =   5820
      Width           =   1035
   End
   Begin VB.CommandButton cmdExtornar 
      Caption         =   "&Extornar"
      Height          =   375
      Left            =   8340
      TabIndex        =   6
      Top             =   5820
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
      Left            =   60
      TabIndex        =   10
      Top             =   1500
      Width           =   10395
      Begin SICMACT.FlexEdit grdMov 
         Height          =   3855
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   10155
         _ExtentX        =   17912
         _ExtentY        =   6800
         Cols0           =   11
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Mov-Operación-Cuenta-Monto-Documento-Cliente-cPersCodCMAC-nMovNro-Moneda-CMACDesc"
         EncabezadosAnchos=   "250-2200-2300-1600-1200-2500-1200-0-0-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-C-R-L-R-R-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-2-2-2-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   255
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame fraBuscar 
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
      Left            =   60
      TabIndex        =   9
      Top             =   60
      Width           =   7395
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   2220
         TabIndex        =   12
         Top             =   240
         Width           =   4995
         Begin VB.Frame fraCuenta 
            Height          =   540
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Visible         =   0   'False
            Width           =   3675
            Begin VB.TextBox txtCuenta 
               Height          =   285
               Left            =   1155
               MaxLength       =   18
               TabIndex        =   17
               Top             =   180
               Width           =   2340
            End
            Begin VB.Label lblCuenta 
               Caption         =   "Nro Cuenta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Left            =   90
               TabIndex        =   18
               Top             =   255
               Width           =   1035
            End
         End
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "&Buscar"
            Height          =   375
            Left            =   3840
            TabIndex        =   3
            Top             =   360
            Width           =   1035
         End
         Begin VB.TextBox txtMovNro 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   1800
            TabIndex        =   2
            Top             =   372
            Width           =   1455
         End
         Begin VB.Label lblMov 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   840
            TabIndex        =   14
            Top             =   375
            Width           =   975
         End
         Begin VB.Label lblNroMov 
            Caption         =   "# Mov :"
            Height          =   195
            Left            =   180
            TabIndex        =   13
            Top             =   450
            Width           =   675
         End
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1995
         Begin VB.OptionButton optTipoBus 
            Caption         =   "Número de &Cuenta"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   1
            Top             =   540
            Width           =   1755
         End
         Begin VB.OptionButton optTipoBus 
            Caption         =   "&Número Movimiento"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   0
            Top             =   240
            Width           =   1755
         End
      End
   End
End
Attribute VB_Name = "frmPITCapExtornos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nOperacion As Long
Dim nProducto As Integer
Dim sOpeCodExt As String, sOpeCodDesc As String

Public Sub Inicia(ByVal nOpe As CaptacOperacion, ByVal sOperacion As String, ByVal nProd As Integer, Optional ByVal psCodExtOpc As String = "")
    nOperacion = nOpe
    Me.Caption = "Operaciones InterCajas - Extornos - " & sOperacion
      
    sOpeCodExt = psCodExtOpc
    sOpeCodDesc = sOperacion
    
    optTipoBus(0).value = True
    cmdExtornar.Enabled = False
    cmdCancelar.Enabled = False
    nProducto = nProd
    Me.Show 1
End Sub

Private Sub AgregaMovGrid(ByVal rsMov As Recordset)
Dim nFila As Long
    Do While Not rsMov.EOF
        grdMov.AdicionaFila
        nFila = grdMov.Rows - 1
        grdMov.TextMatrix(nFila, 1) = rsMov("cMovNro")
        grdMov.TextMatrix(nFila, 2) = rsMov("cOpeDesc")
        grdMov.TextMatrix(nFila, 3) = rsMov("cCuenta")
        grdMov.TextMatrix(nFila, 4) = Format$(rsMov("nMonto"), "#,##0.00")
        grdMov.TextMatrix(nFila, 5) = Trim(rsMov("cDocumento"))
        grdMov.TextMatrix(nFila, 6) = Trim(rsMov("cCliente"))
        grdMov.TextMatrix(nFila, 7) = Trim(rsMov("cPersCod"))
        grdMov.TextMatrix(nFila, 8) = Trim(rsMov("nMovNro"))
        grdMov.TextMatrix(nFila, 9) = Trim(rsMov("nMoneda"))
        grdMov.TextMatrix(nFila, 10) = Trim(rsMov("cCMACDesc"))
           
        rsMov.MoveNext
    Loop
End Sub

Private Sub cmdBuscar_Click()
Dim loCOMPITNeg As COMOpeInterCMAC.dFuncionesNeg
Dim lrsMov As ADODB.Recordset
Dim lsDatoBusq As String, lsFecha As String

    Set loCOMPITNeg = New COMOpeInterCMAC.dFuncionesNeg
    
    lsFecha = Format(gdFecSis, "YYYYMMDD")
    
    If optTipoBus(0).value Then
        lsDatoBusq = Trim(txtMovNro.Text)
        Set lrsMov = loCOMPITNeg.obtenerMovimientosInterCajasParaExtorno(1, lsDatoBusq, CStr(nOperacion), lsFecha)
    ElseIf optTipoBus(1).value Then
        lsDatoBusq = txtCuenta.Text
        Set lrsMov = loCOMPITNeg.obtenerMovimientosInterCajasParaExtorno(2, lsDatoBusq, CStr(nOperacion), lsFecha)
    End If

    fraBuscar.Enabled = False
    cmdExtornar.Enabled = True
    cmdCancelar.Enabled = True
    
    If Not (lrsMov.EOF And lrsMov.BOF) Then
        AgregaMovGrid lrsMov
    Else
        MsgBox "No se registraron movimientos con el criterio de búsqueda", vbInformation, "Aviso"
    End If
    
    Set loCOMPITNeg = Nothing
    Set lrsMov = Nothing
End Sub

Private Sub cmdCancelar_Click()
    grdMov.Clear
    grdMov.Rows = 2
    grdMov.FormaCabecera
    fraBuscar.Enabled = True
    cmdExtornar.Enabled = False
    cmdCancelar.Enabled = False
    optTipoBus(0).value = True
End Sub

Private Sub cmdExtornar_Click()
Dim lnMovNroAExt As Long, lnFila As Long
Dim lsCuenta As String, lsNroDoc As String, lsDescOpe As String, lsPersCodCMAC As String
Dim lsCliente As String, lsGlosa As String, lsNombreCMAC As String
Dim lnMonto As Currency
Dim lnMoneda As Integer


    If Trim(grdMov.TextMatrix(1, 2)) <> "" Then
        
        If MsgBox("¿Desea extornar la operación?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            lnFila = grdMov.Row
            
            lsCuenta = grdMov.TextMatrix(lnFila, 3)
            lnMonto = CCur(grdMov.TextMatrix(lnFila, 4))
            lsNroDoc = grdMov.TextMatrix(lnFila, 5)
            lsCliente = grdMov.TextMatrix(lnFila, 6)
            lsPersCodCMAC = grdMov.TextMatrix(lnFila, 7)
            lnMovNroAExt = CLng(grdMov.TextMatrix(lnFila, 8))
            lnMoneda = grdMov.TextMatrix(lnFila, 9)
            lsNombreCMAC = grdMov.TextMatrix(lnFila, 10)
        
            lsGlosa = Trim(txtGlosa.Text)

            Call RegistrarOperacionInterCMAC("", "", lsCuenta, sOpeCodExt, "", lnMoneda, lsNroDoc, lsPersCodCMAC, sLpt, sOpeCodDesc, lsNombreCMAC, gdFecSis, gsCodAge, gsCodUser, lnMonto, lsGlosa, "", False, lnMovNroAExt)

            Unload Me
        End If
    Else
        MsgBox "No existen datos para realizar el Extorno", vbInformation, "Aviso"
    End If
    
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub grdMov_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtGlosa.SetFocus
    End If
End Sub

Private Sub optTipoBus_Click(Index As Integer)
    lblNroMov.Visible = False
    txtMovNro.Visible = False
    lblMov.Visible = False
    txtCuenta.Visible = False
    fraCuenta.Visible = False
    Select Case Index
        Case 0
            lblMov.Visible = True
            lblNroMov.Visible = True
            txtMovNro.Visible = True
            lblMov = Format$(gdFecSis, "YYYYMMDD")
        Case 1
            fraCuenta.Visible = True
            fraCuenta.Enabled = True
            txtCuenta.Visible = True
            txtCuenta.Enabled = True
            txtCuenta.Text = ""
    End Select
End Sub

Private Sub optTipoBus_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index = 0 Then
        txtMovNro.SetFocus
    ElseIf Index = 1 Then
        txtCuenta.SetFocus
    End If
End If
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBuscar.SetFocus
    Exit Sub
End If
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
KeyAscii = fgIntfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    cmdExtornar.SetFocus
End If
End Sub

Private Sub txtMovNro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBuscar.SetFocus
    Exit Sub
End If
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub


