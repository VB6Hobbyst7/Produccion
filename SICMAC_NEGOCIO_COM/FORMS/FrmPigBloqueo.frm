VERSION 5.00
Begin VB.Form FrmPigBloqueo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bloqueo / Desbloqueo de Contratos"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310
   Icon            =   "FrmPigBloqueo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraContenedor 
      Height          =   4170
      Index           =   0
      Left            =   30
      TabIndex        =   3
      Top             =   -15
      Width           =   8250
      Begin VB.Frame fraBloqueo 
         Caption         =   "Bloqueo/Desbloqueo "
         Enabled         =   0   'False
         Height          =   1020
         Left            =   75
         TabIndex        =   25
         Top             =   3075
         Width           =   8085
         Begin VB.CheckBox chkBloqueo 
            Caption         =   "Bloqueo/Desbloqueo Contrato"
            Height          =   570
            Left            =   150
            TabIndex        =   29
            Top             =   345
            Width           =   2175
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   675
            Left            =   4305
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   28
            Top             =   225
            Width           =   3675
         End
         Begin VB.OptionButton optMotivo 
            Caption         =   "Mandato Judicial"
            Height          =   195
            Index           =   1
            Left            =   2685
            TabIndex        =   27
            Top             =   630
            Width           =   1530
         End
         Begin VB.OptionButton optMotivo 
            Caption         =   "Administrativo"
            Height          =   195
            Index           =   0
            Left            =   2685
            TabIndex        =   26
            Top             =   330
            Value           =   -1  'True
            Width           =   1320
         End
         Begin VB.Label Label3 
            Caption         =   "Motivo "
            Height          =   225
            Left            =   3150
            TabIndex        =   30
            Top             =   0
            Width           =   585
         End
      End
      Begin VB.CommandButton cmdBuscar 
         Height          =   390
         Left            =   7635
         Picture         =   "FrmPigBloqueo.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Buscar ..."
         Top             =   255
         Width           =   420
      End
      Begin VB.Frame fraDatos 
         Enabled         =   0   'False
         Height          =   2415
         Left            =   75
         TabIndex        =   4
         Top             =   630
         Width           =   8100
         Begin VB.Frame FraDetContrato 
            Height          =   1470
            Left            =   75
            TabIndex        =   5
            Top             =   855
            Width           =   7950
            Begin VB.Label lblSaldo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   3780
               TabIndex        =   33
               Top             =   1020
               Width           =   1110
            End
            Begin VB.Label Label5 
               Caption         =   "Saldo"
               Height          =   195
               Left            =   2910
               TabIndex        =   32
               Top             =   1050
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "Piezas"
               Height          =   195
               Left            =   120
               TabIndex        =   21
               Top             =   315
               Width           =   480
            End
            Begin VB.Label lblPiezas 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   1305
               TabIndex        =   20
               Top             =   270
               Width           =   1080
            End
            Begin VB.Label Label2 
               Caption         =   "Peso Bruto (gr)"
               Height          =   195
               Left            =   105
               TabIndex        =   19
               Top             =   675
               Width           =   1170
            End
            Begin VB.Label lblPBruto 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   1305
               TabIndex        =   18
               Top             =   645
               Width           =   1080
            End
            Begin VB.Label Label4 
               Caption         =   "Peso Neto (gr)"
               Height          =   195
               Left            =   105
               TabIndex        =   17
               Top             =   1035
               Width           =   1140
            End
            Begin VB.Label lblPNeto 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   1305
               TabIndex        =   16
               Top             =   1035
               Width           =   1080
            End
            Begin VB.Label Label6 
               Caption         =   "Tasación"
               Height          =   195
               Left            =   2880
               TabIndex        =   15
               Top             =   285
               Width           =   795
            End
            Begin VB.Label lblTasacion 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   3780
               TabIndex        =   14
               Top             =   255
               Width           =   1110
            End
            Begin VB.Label Label8 
               Caption         =   "Prestamo"
               Height          =   195
               Left            =   2895
               TabIndex        =   13
               Top             =   660
               Width           =   795
            End
            Begin VB.Label lblPrestamo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   3780
               TabIndex        =   12
               Top             =   630
               Width           =   1110
            End
            Begin VB.Label Label12 
               Caption         =   "Fec. Prestamo"
               Height          =   195
               Left            =   5325
               TabIndex        =   11
               Top             =   300
               Width           =   1125
            End
            Begin VB.Label lblFecPrestamo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   6690
               TabIndex        =   10
               Top             =   270
               Width           =   1095
            End
            Begin VB.Label Label14 
               Caption         =   "Fec. Vencimiento"
               Height          =   195
               Left            =   5310
               TabIndex        =   9
               Top             =   645
               Width           =   1260
            End
            Begin VB.Label lblFecVencimiento 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   6690
               TabIndex        =   8
               Top             =   630
               Width           =   1095
            End
            Begin VB.Label Label16 
               Caption         =   "Estado"
               Height          =   195
               Left            =   5325
               TabIndex        =   7
               Top             =   1065
               Width           =   630
            End
            Begin VB.Label lblEstado 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   6690
               TabIndex        =   6
               Top             =   1035
               Width           =   1110
            End
         End
         Begin SICMACT.FlexEdit feCte 
            Height          =   705
            Left            =   75
            TabIndex        =   22
            Top             =   210
            Width           =   7980
            _ExtentX        =   13996
            _ExtentY        =   1244
            Cols0           =   4
            FixedCols       =   0
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "Codigo-Nombre/Razon Social-Doc.Iden-Direccion"
            EncabezadosAnchos=   "1200-3200-1200-2200"
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
            ColumnasAEditar =   "X-X-X-X"
            ListaControles  =   "0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-C-C"
            FormatosEdit    =   "0-0-0-0"
            TextArray0      =   "Codigo"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   1200
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin VB.Label lblNumEstado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   4020
            TabIndex        =   31
            Top             =   1875
            Width           =   1110
         End
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         Texto           =   "Crédito"
         EnabledCta      =   -1  'True
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4890
      TabIndex        =   2
      Top             =   4290
      Width           =   1005
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   7275
      TabIndex        =   1
      Top             =   4290
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   6105
      TabIndex        =   0
      Top             =   4290
      Width           =   975
   End
End
Attribute VB_Name = "FrmPigBloqueo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim fbBlqIni As Boolean
'Dim fsMovNroBloqueo As String
'
'Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
'Dim oPers As UPersona
'
'    If KeyAscii = 13 Then BuscaContrato (AXCodCta.NroCuenta)
'
'End Sub
'
'Private Sub chkBloqueo_Click()
'If fbBlqIni = True Then
'    cmdGrabar.Enabled = IIf(chkBloqueo.value = 1, False, True)
'Else
'    cmdGrabar.Enabled = IIf(chkBloqueo.value = 1, True, False)
'    txtDescripcion.Enabled = IIf(chkBloqueo.value = 1, True, False)
'End If
'
'End Sub
'
'Private Sub cmdBuscar_Click()
'Dim oPers As UPersona
'Dim lsPersCod As String, lsPersNombre As String
'Dim lsEstado As String
'Dim oPersContrato As DPigContrato
'Dim rs As ADODB.Recordset
'Dim oCuentas As UProdPersona
'
''On Error GoTo ControlError
'
'Set oPers = New UPersona
'    Set oPers = frmBuscaPersona.Inicio
'    If oPers Is Nothing Then Exit Sub
'    lsPersCod = oPers.sPersCod
'    lsPersNombre = oPers.sPersNombre
'    feCte.TextMatrix(1, 0) = oPers.sPersCod
'    feCte.TextMatrix(1, 1) = oPers.sPersNombre
'    feCte.TextMatrix(1, 2) = oPers.sPersIdnroDNI
'    feCte.TextMatrix(1, 3) = oPers.sPersDireccDomicilio
'Set oPers = Nothing
'
'lsEstado = CStr(gPigEstRegis) + "," + CStr(gPigEstDesemb) + "," + CStr(gPigEstAmortiz) + "," + CStr(gPigEstReusoLin) + "," + CStr(gPigEstCancelPendRes)
'lsEstado = lsEstado + "," + CStr(gPigEstRemat) + "," + CStr(gPigEstRematPRes) + "," + CStr(gPigEstRematPFact) + "," + CStr(gPigEstPResRematPFact)
'
'If Trim(lsPersCod) <> "" Then
'    Set oPersContrato = New DPigContrato
'    Set rs = oPersContrato.dObtieneCredPigDePersona(lsPersCod, lsEstado, Mid(gsCodAge, 4, 2))
'    Set oPersContrato = Nothing
'End If
'
'Set oCuentas = New UProdPersona
'    Set oCuentas = frmProdPersona.Inicio(lsPersNombre, rs)
'    If oCuentas.sCtaCod <> "" Then
'        AXCodCta.NroCuenta = Mid(oCuentas.sCtaCod, 1, 18)
'        AXCodCta.SetFocusCuenta
'    End If
'Set oCuentas = Nothing
'
'If AXCodCta.Cuenta = "" Then
'    Limpiar
'End If
'End Sub
'
'Private Sub cmdCancelar_Click()
'    Limpiar
'    cmdGrabar.Enabled = False
'    cmdBuscar.Enabled = True
'    AXCodCta.Enabled = True
'    AXCodCta.SetFocusCuenta
'End Sub
'
'Private Sub cmdGrabar_Click()
'Dim oContFunct As NContFunciones
'Dim oGrabarBloqueo As NPigContrato
'Dim lsMovNro As String
'Dim lsFechaHoraGrab As String
'Dim lsCuenta As String
'Dim lsLote As String
'Dim nMotBloqueo As Integer
'
''On Error GoTo ControlError
'
'lsCuenta = AXCodCta.NroCuenta
'
'    If MsgBox(" Grabar Bloqueo/Desbloqueo ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
'        cmdGrabar.Enabled = False
'
'        'Genera el Mov Nro
'        Set oContFunct = New NContFunciones
'            lsMovNro = oContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'        Set oContFunct = Nothing
'
'        If optMotivo(0).value = True Then
'            nMotBloqueo = 8
'        ElseIf optMotivo(1).value = True Then
'            nMotBloqueo = 3
'        End If
'        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
'        Set oGrabarBloqueo = New NPigContrato
'            'Grabar Bloqueo / DesBloqueo
'            Call oGrabarBloqueo.nBloqueoDesBloqueoCredPignoraticio(lsCuenta, lsFechaHoraGrab, lsMovNro, fbBlqIni, nMotBloqueo, lblNumEstado, lblSaldo, Me.txtDescripcion.Text, fsMovNroBloqueo, False)
'        Set oGrabarBloqueo = Nothing
'
'        Limpiar
'        fraBloqueo.Enabled = False
'        cmdBuscar.Enabled = True
'        AXCodCta.Enabled = True
'        AXCodCta.SetFocus
'    Else
'        MsgBox " Grabación cancelada ", vbInformation, " Aviso "
'    End If
'Exit Sub
'ControlError:   ' Rutina de control de errores.
'    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & '        " Avise al Area de Sistemas ", vbInformation, " Aviso "
'
'End Sub
'
'Private Sub cmdSalir_Click()
'    Unload Me
'End Sub
'
'Private Sub Limpiar()
'
'AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
'lblEstado = ""
'lblFecPrestamo = ""
'lblFecVencimiento = ""
'lblPBruto = ""
'lblPNeto = ""
'lblPiezas = ""
'lblPrestamo = ""
'lblTasacion = ""
'lblTasacion = ""
'lblSaldo = ""
'feCte.Clear
'feCte.Rows = 2
'feCte.FormaCabecera
'txtDescripcion = ""
'chkBloqueo.value = 0
'
'End Sub
'
'Private Sub BuscaContrato(ByVal psNroContrato As String)
'Dim rs As ADODB.Recordset
'Dim oValContrato As nPigValida
'Dim oPigContrato As DPigContrato
'Dim lsmensaje As String
'
'On Error GoTo ControlError
'
'    'Valida Contrato
'    Set rs = New ADODB.Recordset
'    Set oValContrato = New nPigValida
'        '27-12
'        'Set rs = oValContrato.nValidaBloqueoCredPignoraticio(psNroContrato, lsmensaje )
'        Set rs = oValContrato.nValidaBloqueoCredPignoraticio(psNroContrato)
'        If Trim(lsmensaje) <> "" Then
'             MsgBox lsmensaje, vbInformation, "Aviso"
'             Exit Sub
'        End If
'
'    Set oValContrato = Nothing
'
'    If rs Is Nothing Then ' Hubo un Error
'        Limpiar
'        Set rs = Nothing
'        Exit Sub
'    End If
'
'    If rs.BOF And rs.EOF Then
'        Limpiar
'        Set rs = Nothing
'        Exit Sub
'    End If
'
'    fbBlqIni = IIf(rs!cBloqueo = "S", True, False)
'    If rs!cBloqueo = "S" Then
'        fsMovNroBloqueo = rs!cMovNroBloqueo
'        chkBloqueo.value = 1
'        txtDescripcion.Text = rs!cComentario
'        If rs!nBlqMotivo = 8 Then
'            optMotivo.Item(0).value = True
'            optMotivo.Item(1).value = False
'        Else
'            optMotivo.Item(0).value = False
'            optMotivo.Item(1).value = True
'        End If
'    Else
'        chkBloqueo.value = 0
'    End If
'
'    Set rs = Nothing
'
'    '== Muestro los datos del contrato
'    Set oPigContrato = New DPigContrato
'
'    Set rs = oPigContrato.dObtieneDatosCreditoPignoraticioPersonas(psNroContrato)
'    feCte.TextMatrix(1, 0) = rs!cPersCod
'    feCte.TextMatrix(1, 1) = PstaNombre(rs!cPersNombre)
'    feCte.TextMatrix(1, 2) = IIf(IsNull(rs!NroDNI), " ", rs!NroDNI)
'    feCte.TextMatrix(1, 3) = rs!cPersDireccDomicilio + " " + rs!Zona
'    Set rs = Nothing
'
'    Set rs = oPigContrato.dObtieneDatosContrato(psNroContrato, gPigTipoTasacNor)
'
'    If Not rs.EOF And Not rs.BOF Then
'        lblPiezas = rs!npiezas
'        lblPBruto = Format(rs!nPBruto, "######.00")
'        lblPNeto = Format(rs!nPNeto, "######.00")
'        lblTasacion = Format(rs!nTasacion, "#######.00")
'        lblPrestamo = Format(rs!nMontoCol, "#######.00")
'        lblSaldo = Format(rs!nSaldo, "#######.00")
'        lblNumEstado = rs!nPrdEstado
'        lblEstado = rs!Estado
'        lblFecPrestamo = Format$(rs!dVigencia, "dd/mm/yyyy")
'        lblFecVencimiento = Format$(rs!dvenc, "dd/mm/yyyy")
'    End If
'
'    Set rs = Nothing
'
'    fraBloqueo.Enabled = True
'    chkBloqueo.SetFocus
'    AXCodCta.Enabled = False
'    cmdBuscar.Enabled = False
'    cmdGrabar.Enabled = True
'    cmdGrabar.SetFocus
'
'Exit Sub
'
'ControlError:   ' Rutina de control de errores.
'    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & '        " Avise al Area de Sistemas ", vbInformation, " Aviso "
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF12 And AXCodCta.Enabled = True Then 'F12
'        Dim sCuenta As String
'        sCuenta = frmValTarCodAnt.Inicia(gColConsuPrendario, False)
'        If sCuenta <> "" Then
'            AXCodCta.NroCuenta = sCuenta
'            AXCodCta.SetFocusCuenta
'        End If
'    End If
'End Sub
'
'Private Sub Form_Load()
'    AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
'End Sub
