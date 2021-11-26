VERSION 5.00
Begin VB.Form frmOpeDevSobranteVoucher 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   Icon            =   "frmOpeDevSobranteVoucher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Height          =   345
      Left            =   30
      TabIndex        =   15
      ToolTipText     =   "Grabar Operación"
      Top             =   4320
      Width           =   1155
   End
   Begin VB.TextBox txtGlosa 
      Appearance      =   0  'Flat
      Height          =   825
      Left            =   30
      MaxLength       =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   3375
      Width           =   3555
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
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
      Height          =   345
      Left            =   5010
      TabIndex        =   9
      ToolTipText     =   "Cancelar Operación"
      Top             =   4320
      Width           =   1155
   End
   Begin VB.Frame Frame4 
      Height          =   2415
      Left            =   30
      TabIndex        =   7
      Top             =   690
      Width           =   6135
      Begin SICMACT.FlexEdit feOperacion 
         Height          =   2145
         Left            =   60
         TabIndex        =   8
         Top             =   180
         Width           =   6030
         _ExtentX        =   10636
         _ExtentY        =   3784
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Id-Registro-N° Voucher-Moneda-Monto-nSobrante-nMontoUsado-nMontoVoucher-nMoneda"
         EncabezadosAnchos=   "450-0-1400-1300-1200-1000-0-0-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-L-C-R-C-C-C-C"
         FormatosEdit    =   "0-0-0-2-0-2-2-2-2-2"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   450
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Titular del Crédito"
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
      Height          =   660
      Left            =   30
      TabIndex        =   4
      Top             =   0
      Width           =   6135
      Begin SICMACT.TxtBuscar txtPersonaCod 
         Height          =   315
         Left            =   60
         TabIndex        =   5
         Top             =   240
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
         sTitulo         =   ""
         TipoBusPers     =   1
         EnabledText     =   0   'False
      End
      Begin VB.Label lblPersonaNombre 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   1845
         TabIndex        =   6
         Top             =   240
         Width           =   4230
      End
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   4470
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   3900
      Width           =   1695
   End
   Begin VB.TextBox txtITF 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   4470
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "0.00"
      Top             =   3540
      Width           =   1695
   End
   Begin VB.TextBox txtMonto 
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
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   4470
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "0.00"
      Top             =   3180
      Width           =   1695
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3840
      TabIndex        =   0
      ToolTipText     =   "Grabar Operación"
      Top             =   4320
      Width           =   1155
   End
   Begin VB.Label lblGlosa 
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
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   30
      TabIndex        =   14
      Top             =   3165
      Width           =   915
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Monto :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3735
      TabIndex        =   13
      Top             =   3240
      Width           =   660
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3735
      TabIndex        =   12
      Top             =   4005
      Width           =   570
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "ITF :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3735
      TabIndex        =   11
      Top             =   3645
      Width           =   420
   End
End
Attribute VB_Name = "frmOpeDevSobranteVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fsOpeCod As String
Dim fsOpeDesc As String

Private Sub cmdCancelar_Click()
limpiarCampos
txtPersonaCod.SetFocus
End Sub
Private Function ValidaGrabar() As Boolean
    ValidaGrabar = True
    If Len(txtPersonaCod.Text) <> 13 Or Len(lblPersonaNombre.Caption) = 0 Then
        ValidaGrabar = False
        MsgBox "Ud. debe seleccionar el cliente para continuar", vbInformation, "Aviso"
        If txtPersonaCod.Visible And txtPersonaCod.Enabled Then txtPersonaCod.SetFocus
        Exit Function
    End If
    If feOperacion.TextMatrix(1, 0) = "" Then
        ValidaGrabar = False
        MsgBox "No existe voucher seleccionado, no se puede continuar", vbInformation, "Aviso"
        If feOperacion.Visible And feOperacion.Enabled Then feOperacion.SetFocus
        Exit Function
    End If
    If Len(Trim(txtGlosa.Text)) = 0 Then
        ValidaGrabar = False
        MsgBox "Ud. debe ingresar la glosa para continuar", vbInformation, "Aviso"
        If txtGlosa.Visible And txtGlosa.Enabled Then txtGlosa.SetFocus
        Exit Function
    End If
    Dim nMonto As Currency
    If IsNumeric(feOperacion.TextMatrix(feOperacion.row, 6)) Then
        nMonto = CCur(feOperacion.TextMatrix(feOperacion.row, 6))
    Else
        nMonto = 0
    End If
    If nMonto <= 0 Then
        MsgBox "El monto debe ser mayor a cero", vbInformation, "Aviso"
        If txtMonto.Visible And txtMonto.Enabled Then txtMonto.SetFocus
        ValidaGrabar = False
        Exit Function
    End If
End Function
Private Sub cmdGrabar_Click()
    
Dim ClsMov As COMNCaptaGenerales.NCOMCaptaGenerales
Dim oFun As COMNContabilidad.NCOMContFunciones
Dim loVistoElectronico As frmVistoElectronico
Dim sMovNro As String
Dim sOpeCod As CaptacOperacion
Dim sGlosa As String
Dim nMonto As Currency
Dim nId As Integer
Dim sNroVou As String
Dim nmoneda As Moneda
Dim nMovNro As Long
Dim bExito  As Boolean
Dim psCodCta As String 'ALPA201600505
        
    Set oFun = New COMNContabilidad.NCOMContFunciones
    
    If Not ValidaGrabar() Then
        If txtPersonaCod.Visible And txtPersonaCod.Enabled Then txtPersonaCod.SetFocus
        Exit Sub
    End If

    If fsOpeCod = gOtrOpeEgresoDevSobranteOtrasOpeVoucher Then
        sOpeCod = fsOpeCod
        sGlosa = Replace(txtGlosa.Text, vbNewLine, " ")
        nMonto = feOperacion.TextMatrix(feOperacion.row, 6)
        nId = feOperacion.TextMatrix(feOperacion.row, 1)
        sNroVou = feOperacion.TextMatrix(feOperacion.row, 3)
        nmoneda = feOperacion.TextMatrix(feOperacion.row, 9)
        If nmoneda = gMonedaNacional Then
            psCodCta = "109_____" & 1 & "_________"
        Else
            psCodCta = "109_____" & 2 & "_________"
        End If
        If Len(Trim(sOpeCod)) = 0 Or Len(Trim(sGlosa)) = 0 Or nMonto = 0 Or _
           nId = 0 Or Len(Trim(sNroVou)) = 0 Or nmoneda < gMonedaNacional Or nmoneda > gMonedaExtranjera Then
            MsgBox "Para la presente operación no se ha obtenido los identificadores correctos, comuniquese con el Dpto. de TI", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    Set loVistoElectronico = New SICMACT.frmVistoElectronico
    If Not loVistoElectronico.inicio(9, fsOpeCod) Then
        Set loVistoElectronico = Nothing
        Exit Sub
    End If
    If MsgBox("¿Está seguro de realizar la operación?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Set ClsMov = New COMNCaptaGenerales.NCOMCaptaGenerales
    sMovNro = oFun.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    If fsOpeCod = gOtrOpeEgresoDevSobranteOtrasOpeVoucher Then
        bExito = ClsMov.RegistrarSobrante(sMovNro, gOtrOpeEgresoDevSobranteOtrasOpeVoucher, sGlosa, nMonto, nId, sNroVou, nmoneda, txtPersonaCod.Text, nMovNro)
    End If
        'loVistoElectronico.RegistraVistoElectronico nMovNro 'comment by marg ers052-2017
        loVistoElectronico.RegistraVistoElectronico nMovNro, , gsCodUser, nMovNro 'add by marg ers052-2017
    If bExito Then
        MsgBox "Se ha generado la operación satisfactoriamente", vbInformation, "Aviso"
        limpiarCampos
    Else
        MsgBox "Ha ocurrido un error al realizar la operación, si el error persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
    End If
        
    'Imprimir voucher de sobrantes
    Dim oBol As COMNCaptaGenerales.NCOMCaptaImpresion
    Dim lsBoleta As String
    Set oBol = New COMNCaptaGenerales.NCOMCaptaImpresion
     
    'ALPA20160505******************
    'lsBoleta = oBol.ImprimeBoleta("Sobrantes Voucher", "Voucher " & sNroVou, "", Str(nMonto), lblPersonaNombre.Caption, "________", "", 0, "0", "", 0, 0, False, False, , , , False, , "Nro Ope. : " & Str(nMovNro), , gdFecSis, gsNomAge, gsCodUser, sLpt, , False)
    lsBoleta = oBol.ImprimeBoleta("Sobrantes Voucher", "Voucher " & sNroVou, "", Str(nMonto), lblPersonaNombre.Caption, psCodCta, "", 0, "0", "", 0, 0, False, False, , , , False, , "Nro Ope. : " & Str(nMovNro), , gdFecSis, gsNomAge, gsCodUser, sLpt, , False)
    '******************************
    'ALPA20160505**************
    psCodCta = "________"
    '**************************
    Set oBol = Nothing
    Do
      If Trim(lsBoleta) <> "" Then
            lsBoleta = lsBoleta & oImpresora.gPrnSaltoLinea
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
                Print #nFicSal, oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                Print #nFicSal, ""
            Close #nFicSal
      End If
                
    Loop While MsgBox("Desea Re Imprimir ?", vbQuestion + vbYesNo, "Aviso") = vbYes
    Set oBol = Nothing
      
    Set oFun = Nothing
    Set ClsMov = Nothing
    Set loVistoElectronico = Nothing
    'INICIO JHCU ENCUESTA 16-10-2019
    Encuestas gsCodUser, gsCodAge, "ERS0292019", gsOpeCod
    'FIN
    Exit Sub
    
ErrCmdGrabar:
    MsgBox err.Description, vbCritical, "Aviso"
    Set oFun = Nothing
    Set loVistoElectronico = Nothing
End Sub
Private Sub cmdsalir_Click()
    If MsgBox("¿Deseas salir de la formulario?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Unload Me
    End If
End Sub
Public Sub inicio(ByVal psOpeCod As String, ByVal psOpeDesc As String)
    fsOpeCod = psOpeCod
    Caption = UCase(Mid(psOpeDesc, 3, Len(psOpeDesc)))
    Show 1
End Sub
Private Sub feOperacion_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim sColumnas() As String
    sColumnas = Split(feOperacion.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
End Sub
Private Sub Form_Load()
    If fsOpeCod = gOtrOpeEgresoDevSobranteOtrasOpeVoucher Then
        feOperacion.EncabezadosNombres = "#-Id-Registro-N° Voucher-Moneda-Monto-nSobrante-nMontoUsado-nMontoVoucher"
        feOperacion.EncabezadosAlineacion = "C-C-C-L-C-R-C-C-C"
        txtPersonaCod.psDescripcion = fsOpeCod
    End If
End Sub
Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
KeyAscii = Letras(KeyAscii)
End Sub

Private Sub txtPersonaCod_EmiteDatos()
    lblPersonaNombre.Caption = ""
    If Len(Trim(txtPersonaCod.Text)) = 13 Then
        lblPersonaNombre.Caption = txtPersonaCod.psDescripcion
        txtPersonaCod.psDescripcion = fsOpeCod
        If Not CargaDatos(txtPersonaCod.Text) Then
            MsgBox "No se encontraron datos con la persona seleccionada", vbInformation, "Aviso"
            limpiarCampos
            txtPersonaCod.SetFocus
        Else
            feOperacion.row = 1
            feOperacion.Col = 2
            feOperacion.SetFocus
        End If
    Else
        limpiarCampos
        txtPersonaCod.SetFocus
    End If
    If fsOpeCod = gOtrOpeEgresoDevSobranteOtrasOpeVoucher Then
        txtPersonaCod.psDescripcion = fsOpeCod
    End If
End Sub
Private Function CargaDatos(ByVal psPersCod As String) As Boolean
    Dim oDR As New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim oRS As New ADODB.Recordset
    Dim fila As Long
    
    On Error GoTo ErrCargaDatos
    Screen.MousePointer = 11
    FormateaFlex feOperacion
    
    If fsOpeCod = gOtrOpeEgresoDevSobranteOtrasOpeVoucher Then
        Set oRS = oDR.ObtenerSobranteVoucherPersona(psPersCod)
    'ElseIf 'Colocar las nuevas operaciones
    End If
    
    If Not oRS Is Nothing Then
        If Not oRS.EOF Then
            CargaDatos = True
            Do While Not oRS.EOF
                feOperacion.AdicionaFila
                fila = feOperacion.row
                feOperacion.TextMatrix(fila, 1) = oRS!nId 'nId
                feOperacion.TextMatrix(fila, 2) = Format(oRS!dfecReg, gsFormatoFechaView) 'Fecha Registro
                feOperacion.TextMatrix(fila, 3) = oRS!cNroVou 'Numer de Voucher
                feOperacion.TextMatrix(fila, 4) = oRS!cmoneda 'Moneda
                'feOperacion.TextMatrix(Fila, 5) = Format(0, gsFormatoNumeroView) 'Monto Sobrante
                feOperacion.TextMatrix(fila, 5) = Format(oRS!nSobrante, gsFormatoNumeroView) 'Monto Sobrante'WIOR 20160425
                feOperacion.TextMatrix(fila, 6) = Format(oRS!nSobrante, gsFormatoNumeroView) 'Monto Sobrante
                feOperacion.TextMatrix(fila, 7) = Format(oRS!nMonto, gsFormatoNumeroView)    'Monto usado
                feOperacion.TextMatrix(fila, 8) = Format(oRS!nMonVou, gsFormatoNumeroView)   'Monto voucher
                feOperacion.TextMatrix(fila, 9) = Format(oRS!nmoneda, gsFormatoNumeroView)   'Moneda soles:1, dolares:2
                oRS.MoveNext
            Loop
        Else
            CargaDatos = False
        End If
    End If
    txtMonto.Text = SumarCampo(feOperacion, 5) 'WIOR 20160425
    Set oRS = Nothing
    Set oDR = Nothing
    Screen.MousePointer = 0
    Exit Function
ErrCargaDatos:
    CargaDatos = False
    MsgBox err.Description, vbCritical, "Aviso"
End Function
Private Sub limpiarCampos()
    txtPersonaCod.Text = ""
    lblPersonaNombre.Caption = ""
    FormateaFlex feOperacion
    txtGlosa.Text = ""
    txtMonto.Text = "0.00"
    TxtITF.Text = "0.00"
    txtTotal.Text = "0.00"
End Sub
Private Sub feOperacion_KeyPress(KeyAscii As Integer)
    If feOperacion.TextMatrix(1, 0) <> "" Then
        If txtGlosa.Visible And txtGlosa.Enabled Then txtGlosa.SetFocus
    End If
End Sub
