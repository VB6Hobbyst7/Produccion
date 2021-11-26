VERSION 5.00
Begin VB.Form frmOpeDevSobrante 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   435
   ClientWidth     =   7530
   Icon            =   "frmOpeDevSobrante.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7530
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   345
      Left            =   4710
      TabIndex        =   14
      ToolTipText     =   "Grabar Operación"
      Top             =   4740
      Width           =   1365
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
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   3480
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
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   3840
      Width           =   1695
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
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   4200
      Width           =   1695
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
      Left            =   80
      TabIndex        =   8
      Top             =   40
      Width           =   7410
      Begin SICMACT.TxtBuscar txtPersonaCod 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
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
      End
      Begin VB.Label lblPersonaNombre 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1850
         TabIndex        =   9
         Top             =   240
         Width           =   5430
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2475
      Left            =   80
      TabIndex        =   7
      Top             =   720
      Width           =   7410
      Begin SICMACT.FlexEdit feOperacion 
         Height          =   2145
         Left            =   195
         TabIndex        =   1
         Top             =   195
         Width           =   7080
         _ExtentX        =   12488
         _ExtentY        =   3784
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Item-Id-Registro-N° Cheque-Moneda-Monto"
         EncabezadosAnchos=   "450-0-1400-1500-1500-1500"
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
         ColumnasAEditar =   "X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-L-L-R"
         FormatosEdit    =   "0-0-0-2-0-2"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   450
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   6090
      TabIndex        =   6
      ToolTipText     =   "Cancelar Operación"
      Top             =   4740
      Width           =   1365
   End
   Begin VB.TextBox txtGlosa 
      Appearance      =   0  'Flat
      Height          =   1095
      Left            =   90
      MaxLength       =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3600
      Width           =   4395
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "ITF :"
      Height          =   195
      Left            =   4695
      TabIndex        =   13
      Top             =   3940
      Width           =   330
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Total :"
      Height          =   195
      Left            =   4695
      TabIndex        =   12
      Top             =   4305
      Width           =   450
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Monto :"
      Height          =   195
      Left            =   4695
      TabIndex        =   11
      Top             =   3540
      Width           =   540
   End
   Begin VB.Label lblGlosa 
      Caption         =   "Glosa"
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   90
      TabIndex        =   10
      Top             =   3390
      Width           =   915
   End
End
Attribute VB_Name = "frmOpeDevSobrante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************
'** Nombre : frmChequeDevSobrante
'** Descripción : Para devolución de sobrantes creado segun TI-ERS126-2013
'** Creación : EJVG, 20130212 11:00:00 AM
'*************************************************************************
Option Explicit
Dim fsOpeCod As String
Dim fsOpeDesc As String

Private Sub GetDatos(ByVal pnRow As Long)
    Dim oNCred As New COMNCredito.NCOMCredito
    Dim lnMonto As Currency, lnITF As Currency
    On Error GoTo ErrGetDatos
    Screen.MousePointer = 11
    lnMonto = CCur(feOperacion.TextMatrix(pnRow, 5))
    'PASI20140620
    'lnITF = oNCred.DameMontoITF(lnMonto)
    lnITF = 0
    'end PASI
    txtMonto.Text = Format(lnMonto, gsFormatoNumeroView)
    TxtITF.Text = Format(lnITF, gsFormatoNumeroView)
    txtTotal.Text = Format(lnMonto - lnITF, gsFormatoNumeroView)
    Set oNCred = Nothing
    Screen.MousePointer = 0
    Exit Sub
ErrGetDatos:
    Screen.MousePointer = 0
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
Private Sub feOperacion_RowColChange()
    If feOperacion.TextMatrix(1, 0) <> "" Then
        If feOperacion.row > 0 And feOperacion.Col > 1 Then
            GetDatos feOperacion.row
        End If
    End If
End Sub
Private Sub Form_Load()
    If fsOpeCod = gOtrOpeEgresoDevSobranteOtrasOpeChq Then
        feOperacion.EncabezadosNombres = "Item-Id-Registro-N° Cheque-Moneda-Monto"
        feOperacion.EncabezadosAlineacion = "C-C-C-L-C-R"
        txtPersonaCod.psDescripcion = fsOpeCod
    End If
End Sub
Public Sub inicio(ByVal psOpeCod As String, ByVal psOpeDesc As String)
    fsOpeCod = psOpeCod
    Caption = UCase(Mid(psOpeDesc, 3, Len(psOpeDesc)))
    Show 1
End Sub
Private Sub cmdGrabar_Click()
     
    'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, gsOpeCod, "0")
    If Not bPermitirEjecucionOperacion Then
        End
    End If
    'fin Comprobacion si es RFIII
     
    Dim oNCred As COMNCredito.NCOMCredito
    Dim oFun As COMNContabilidad.NCOMContFunciones
    Dim loVistoElectronico As frmVistoElectronico
    Dim lsMovNro As String
    Dim lnMovNro As Long
    Dim bExito As Boolean
    
    Dim fila As Long
    Dim lsId As String
    Dim lnMonto As Currency, lnMontoFlex As Currency, lnITF As Currency
    Dim lsGlosa As String
    Dim MatDatos() As String
    Dim lnTpoDoc As Integer, lsNroDoc As String, lsPersCod As String, lsIFTpo As String, lsIFCta As String, lsCtaCod As String 'IDs Cheque
    
    On Error GoTo ErrCmdGrabar
    If Not ValidaGrabar Then Exit Sub
    
    fila = feOperacion.row
    lsId = feOperacion.TextMatrix(fila, 1)
    lnMonto = Format(txtMonto.Text, "#0.00")
    lnITF = Format(TxtITF.Text, "#0.00")
    lsGlosa = Trim(txtGlosa.Text)
    lnMontoFlex = Format(feOperacion.TextMatrix(fila, 5), "#0.00")
    
    If fsOpeCod = gOtrOpeEgresoDevSobranteOtrasOpeChq Then
        MatDatos = Split(lsId, "|")
        lnTpoDoc = CInt(MatDatos(0))
        lsNroDoc = MatDatos(1)
        lsPersCod = MatDatos(2)
        lsIFTpo = MatDatos(3)
        lsIFCta = MatDatos(4)
        lsCtaCod = MatDatos(5)
        
        If lnTpoDoc = 0 Or Len(Trim(lsNroDoc)) = 0 Or Len(Trim(lsPersCod)) = 0 Or Len(Trim(lsIFTpo)) = 0 Or Len(Trim(lsIFCta)) = 0 Or Len(Trim(lsCtaCod)) <> 18 Then
            MsgBox "Para la presente operación no se ha obtenido los identificadores correctos, comuniquese con el Dpto. de TI", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    If lnMonto <= 0 Then
        MsgBox "El monto debe ser mayor a cero", vbInformation, "Aviso"
        If txtMonto.Visible And txtMonto.Enabled Then txtMonto.SetFocus
        Exit Sub
    End If
    If lnMonto <> lnMontoFlex Then
        MsgBox "El monto debe ser menor o igual al del detalle, vuelva a seleccionar el registro y presione enter", vbInformation, "Aviso"
        If feOperacion.Visible And feOperacion.Enabled Then feOperacion.SetFocus
        Exit Sub
    End If
    
    Set loVistoElectronico = New frmVistoElectronico
    If Not loVistoElectronico.inicio(9, fsOpeCod) Then
        Set loVistoElectronico = Nothing
        Exit Sub
    End If

    If MsgBox("¿Está seguro de realizar la operación?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Set oNCred = New COMNCredito.NCOMCredito
    Set oFun = New COMNContabilidad.NCOMContFunciones
    lsMovNro = oFun.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    
    'Realizar Operación
    If fsOpeCod = gOtrOpeEgresoDevSobranteOtrasOpeChq Then
        bExito = oNCred.EgresoxDevolucionSobranteOtrasOpeCheque(lsMovNro, fsOpeCod, lsGlosa, lnMonto, lnITF, lsCtaCod, txtPersonaCod.Text, lnTpoDoc, lsNroDoc, lsPersCod, lsIFTpo, lsIFCta, lnMovNro)
    End If
    'loVistoElectronico.RegistraVistoElectronico lnMovNro 'comment by marg ers052-2017
    loVistoElectronico.RegistraVistoElectronico lnMovNro, , gsCodUser, lnMovNro 'add by marg ers052-2017
    
    'PASI20140619
    Dim oBol As COMNCaptaGenerales.NCOMCaptaImpresion
    Dim oBolITF As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim lsBoleta As String
    Dim lsBoletaITF As String
    Dim nFicSal As Integer
    Dim lsCabecera As String 'INICIO ORCR20140714
    'end PASI
    
    If bExito Then
        MsgBox "Se ha generado la operación satisfactoriamente", vbInformation, "Aviso"
        'PASI20140619
        
        Dim sDNI As String
        
        If fsOpeCod = 300528 Then
            Dim oPers As COMDPersona.DCOMPersona
            Set oPers = New COMDPersona.DCOMPersona
            oPers.RecuperaDocumentos (txtPersonaCod.Text)
            
            sDNI = oPers.ObtenerDNI()
            lsCabecera = "DEV. SOBR. CHEQUE"
        Else
            lsCabecera = "OTRAS OPERACIONES"
        End If
        
        Set oBol = New COMNCaptaGenerales.NCOMCaptaImpresion
            'lsBoleta = oBol.ImprimeBoleta("OTRAS OPERACIONES", Left(Caption, 15), fsOpeCod, Str(txtMonto.Text), Me.lblPersonaNombre.Caption, "________" & IIf(feOperacion.TextMatrix(Fila, 4) = "SOLES", gMonedaNacional, gMonedaExtranjera), lsNroDoc, 0, "0", IIf(Len(lsNroDoc) = 0, "", "Nro Documento"), 0, 0, False, False, , , , False, , "Nro Ope. : " & Str(lnMovNro), , gdFecSis, gsNomAge, gsCodUser, sLpt, , False, (txtITF.Text * -1))
            lsBoleta = oBol.ImprimeBoleta(lsCabecera, Left(Caption, 15), fsOpeCod, Str(txtMonto.Text), Me.lblPersonaNombre.Caption, "________" & IIf(feOperacion.TextMatrix(fila, 4) = "SOLES", gMonedaNacional, gMonedaExtranjera), lsNroDoc, 0, "0", IIf(Len(lsNroDoc) = 0, "", "Nro Documento"), 0, 0, False, False, , , , False, , "Nro Ope. : " & Str(lnMovNro), , gdFecSis, gsNomAge, gsCodUser, sLpt, , False, (TxtITF.Text * -1), , , sDNI)
        Set oBol = Nothing
        'FIN ORCR20140714
        Do
            If Trim(lsBoleta) <> "" Then
                nFicSal = FreeFile
                Open sLpt For Output As nFicSal
                    Print #nFicSal, lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                    Print #nFicSal, ""
                Close #nFicSal
            End If
        Loop While MsgBox("Desea Re Imprimir ?", vbQuestion + vbYesNo, "Aviso") = vbYes
        'end PASI
        limpiarCampos
        'INICIO JHCU ENCUESTA 16-10-2019
        Encuestas gsCodUser, gsCodAge, "ERS0292019", gsOpeCod
        'FIN
    Else
        MsgBox "Ha ocurrido un error al realizar la operación, si el error persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
    End If
    
    Set oFun = Nothing
    Set oNCred = Nothing
    Set loVistoElectronico = Nothing
    Exit Sub
ErrCmdGrabar:
    MsgBox err.Description, vbCritical, "Aviso"
    Set oFun = Nothing
    Set oNCred = Nothing
    Set loVistoElectronico = Nothing
End Sub
Private Sub cmdsalir_Click()
    Unload Me
End Sub
Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtMonto.Visible And txtMonto.Enabled Then txtMonto.SetFocus
    End If
End Sub
Private Sub txtITF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtTotal.Visible And txtTotal.Enabled Then txtTotal.SetFocus
    End If
End Sub
Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMonto, KeyAscii)
    If KeyAscii = 13 Then
        If cmdGrabar.Visible And cmdGrabar.Enabled Then cmdGrabar.SetFocus
    End If
End Sub
Private Sub txtMonto_LostFocus()
    If txtMonto.Text = "" Then txtMonto.Text = "0.00"
    txtMonto.Text = Format(txtMonto.Text, gsFormatoNumeroView)
End Sub
Private Sub txtPersonaCod_EmiteDatos()
    lblPersonaNombre.Caption = ""
    If Len(Trim(txtPersonaCod.Text)) = 13 Then
        lblPersonaNombre.Caption = txtPersonaCod.psDescripcion
        txtPersonaCod.psDescripcion = fsOpeCod
        If Not CargaDatos(txtPersonaCod.Text) Then
            MsgBox "No se encontraron datos con la persona seleccionada", vbInformation, "Aviso"
            limpiarCampos
        Else
            feOperacion.SetFocus
        End If
    End If
    If fsOpeCod = gOtrOpeEgresoDevSobranteOtrasOpeChq Then
        txtPersonaCod.psDescripcion = fsOpeCod
    End If
End Sub
Private Sub txtPersonaCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(txtPersonaCod.Text)) = 13 Then
            If feOperacion.Visible And feOperacion.Enabled Then
                feOperacion.SetFocus
                SendKeys "{Right}"
            End If
        End If
    End If
End Sub
Private Sub txtPersonaCod_LostFocus()
    If Len(Trim(txtPersonaCod.Text)) <> 13 Then
        limpiarCampos
    End If
End Sub
Private Sub txtTotal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdGrabar.Visible And cmdGrabar.Enabled Then cmdGrabar.SetFocus
    End If
End Sub
Private Sub feOperacion_KeyPress(KeyAscii As Integer)
    If feOperacion.TextMatrix(1, 0) <> "" Then
        If txtGlosa.Visible And txtGlosa.Enabled Then txtGlosa.SetFocus
    End If
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
        MsgBox "No existe cheque seleccionado, no se puede continuar", vbInformation, "Aviso"
        If feOperacion.Visible And feOperacion.Enabled Then feOperacion.SetFocus
        Exit Function
    End If
    If Len(Trim(txtGlosa.Text)) = 0 Then
        ValidaGrabar = False
        MsgBox "Ud. debe ingresar la glosa para continuar", vbInformation, "Aviso"
        If txtGlosa.Visible And txtGlosa.Enabled Then txtGlosa.SetFocus
        Exit Function
    End If
End Function
Private Function CargaDatos(ByVal psPersCod As String) As Boolean
    Dim oDR As New NCOMDocRec
    Dim oRS As New ADODB.Recordset
    Dim fila As Long
    
    On Error GoTo ErrCargaDatos
    Screen.MousePointer = 11
    FormateaFlex feOperacion
    If fsOpeCod = gOtrOpeEgresoDevSobranteOtrasOpeChq Then
        Set oRS = oDR.ListaChequexDevSobranteOperacion(psPersCod)
    'ElseIf 'Colocar las nuevas operaciones
    End If
    If Not oRS.EOF Then
        CargaDatos = True
        Do While Not oRS.EOF
            feOperacion.AdicionaFila
            fila = feOperacion.row
            feOperacion.TextMatrix(fila, 1) = oRS!cId 'Identificador (Cheque, voucher u otro)
            feOperacion.TextMatrix(fila, 2) = Format(oRS!dFecha, gsFormatoFechaView) 'Fecha Registro
            feOperacion.TextMatrix(fila, 3) = oRS!cDato 'Mostrar Identificador
            feOperacion.TextMatrix(fila, 4) = oRS!cmoneda 'Moneda
            feOperacion.TextMatrix(fila, 5) = Format(oRS!nMonto, gsFormatoNumeroView) 'Monto
            oRS.MoveNext
        Loop
    Else
        CargaDatos = False
    End If
    Set oRS = Nothing
    Set oDR = Nothing
    Screen.MousePointer = 0
    Exit Function
ErrCargaDatos:
    CargaDatos = False
    MsgBox err.Description, vbCritical, "Aviso"
End Function
'Private Sub feOperacion_OnRowChange(pnRow As Long, pnCol As Long)
'    If feOperacion.TextMatrix(1, 0) <> "" Then
'        GetDatos pnRow
'    End If
'End Sub
Private Sub limpiarCampos()
    txtPersonaCod.Text = ""
    lblPersonaNombre.Caption = ""
    FormateaFlex feOperacion
    txtGlosa.Text = ""
    txtMonto.Text = "0.00"
    TxtITF.Text = "0.00"
    txtTotal.Text = "0.00"
End Sub
