VERSION 5.00
Begin VB.Form frmPigAnularContrato 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anular Contrato"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8340
   Icon            =   "frmPigAnularContrato.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   6105
      TabIndex        =   5
      Top             =   3270
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   7290
      TabIndex        =   4
      Top             =   3270
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4890
      TabIndex        =   3
      Top             =   3270
      Width           =   1005
   End
   Begin VB.Frame fraContenedor 
      Height          =   3180
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   8250
      Begin VB.Frame fraDatos 
         Enabled         =   0   'False
         Height          =   2445
         Left            =   75
         TabIndex        =   6
         Top             =   645
         Width           =   8100
         Begin VB.Frame FraDetContrato 
            Height          =   1470
            Left            =   60
            TabIndex        =   8
            Top             =   885
            Width           =   7950
            Begin VB.Label Label5 
               Caption         =   "Saldo"
               Height          =   195
               Left            =   2895
               TabIndex        =   26
               Top             =   1050
               Width           =   795
            End
            Begin VB.Label lblSaldo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   3780
               TabIndex        =   25
               Top             =   1005
               Width           =   1110
            End
            Begin VB.Label lblEstado 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   6705
               TabIndex        =   24
               Top             =   1020
               Width           =   1080
            End
            Begin VB.Label Label16 
               Caption         =   "Estado"
               Height          =   195
               Left            =   5340
               TabIndex        =   23
               Top             =   1050
               Width           =   630
            End
            Begin VB.Label lblFecVencimiento 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   6690
               TabIndex        =   22
               Top             =   630
               Width           =   1095
            End
            Begin VB.Label Label14 
               Caption         =   "Fec. Vencimiento"
               Height          =   195
               Left            =   5310
               TabIndex        =   21
               Top             =   645
               Width           =   1260
            End
            Begin VB.Label lblFecPrestamo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   6690
               TabIndex        =   20
               Top             =   270
               Width           =   1095
            End
            Begin VB.Label Label12 
               Caption         =   "Fec. Prestamo"
               Height          =   195
               Left            =   5325
               TabIndex        =   19
               Top             =   300
               Width           =   1125
            End
            Begin VB.Label lblPrestamo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   3780
               TabIndex        =   18
               Top             =   630
               Width           =   1110
            End
            Begin VB.Label Label8 
               Caption         =   "Prestamo"
               Height          =   195
               Left            =   2895
               TabIndex        =   17
               Top             =   660
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
               TabIndex        =   16
               Top             =   255
               Width           =   1110
            End
            Begin VB.Label Label6 
               Caption         =   "Tasación"
               Height          =   195
               Left            =   2880
               TabIndex        =   15
               Top             =   285
               Width           =   795
            End
            Begin VB.Label lblPNeto 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   1305
               TabIndex        =   14
               Top             =   990
               Width           =   1080
            End
            Begin VB.Label Label4 
               Caption         =   "Peso Neto (gr)"
               Height          =   195
               Left            =   105
               TabIndex        =   13
               Top             =   1035
               Width           =   1140
            End
            Begin VB.Label lblPBruto 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   1305
               TabIndex        =   12
               Top             =   630
               Width           =   1080
            End
            Begin VB.Label Label2 
               Caption         =   "Peso Bruto (gr)"
               Height          =   195
               Left            =   105
               TabIndex        =   11
               Top             =   675
               Width           =   1170
            End
            Begin VB.Label lblPiezas 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   1305
               TabIndex        =   10
               Top             =   270
               Width           =   1080
            End
            Begin VB.Label Label1 
               Caption         =   "Piezas"
               Height          =   195
               Left            =   120
               TabIndex        =   9
               Top             =   315
               Width           =   480
            End
         End
         Begin SICMACT.FlexEdit feCte 
            Height          =   705
            Left            =   75
            TabIndex        =   7
            Top             =   210
            Width           =   7935
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
      End
      Begin VB.CommandButton cmdBuscar 
         Height          =   390
         Left            =   7560
         Picture         =   "frmPigAnularContrato.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Buscar ..."
         Top             =   225
         Width           =   420
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   375
         Left            =   105
         TabIndex        =   2
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         Texto           =   "Crédito"
         EnabledCta      =   -1  'True
      End
   End
End
Attribute VB_Name = "frmPigAnularContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
'Dim oPers As UPersona
'
'    If KeyAscii = 13 Then BuscaContrato (AXCodCta.NroCuenta)
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
'On Error GoTo ControlError
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
'lsEstado = gPigEstRegis
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
'Exit Sub
'
'ControlError:   ' Rutina de control de errores.
'    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & '        " Avise al Area de Sistemas ", vbInformation, " Aviso "
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
'        Set rs = oValContrato.nValidaAnulacionCredPignoraticio(psNroContrato, gdFecSis)
'    Set oValContrato = Nothing
'
'    If rs Is Nothing Then ' Hubo un Error
'        'Limpiar
'        Set rs = Nothing
'        Exit Sub
'    End If
'
'    '== Muestro los datos del contrato
'    Set oPigContrato = New DPigContrato
'
'    If feCte.TextMatrix(1, 0) = "" Then
'        Set rs = oPigContrato.dObtieneDatosCreditoPignoraticioPersonas(psNroContrato)
'        feCte.TextMatrix(1, 0) = rs!cPersCod
'        feCte.TextMatrix(1, 1) = PstaNombre(rs!cPersNombre)
'        feCte.TextMatrix(1, 2) = IIf(IsNull(rs!NroDNI), " ", rs!NroDNI)
'        feCte.TextMatrix(1, 3) = rs!cPersDireccDomicilio + " " + rs!Zona
'        Set rs = Nothing
'    End If
'
'    Set rs = oPigContrato.dObtieneDatosContrato(psNroContrato, gPigTipoTasacNor)
'
'    If Not rs.EOF And Not rs.BOF Then
'        lblPiezas = rs!npiezas
'        lblPBruto = Format(rs!nPBruto, "######.00")
'        lblPNeto = Format(rs!nPNeto, "######.00")
'        lblTasacion = Format(rs!nTasacion, "#######.00")
'        lblPrestamo = Format(rs!nMontoCol, "#######.00")
'        lblEstado = rs!Estado
'        lblSaldo = Format(rs!nSaldo, "#######.00")
'        lblFecPrestamo = Format$(rs!dVigencia, "dd/mm/yyyy")
'        lblFecVencimiento = Format$(rs!dvenc, "dd/mm/yyyy")
'    End If
'
'    Set rs = Nothing
'
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
'Private Sub cmdCancelar_Click()
'    Limpiar
'    cmdGrabar.Enabled = False
'    cmdBuscar.Enabled = True
'    AXCodCta.Enabled = True
'    AXCodCta.SetFocusCuenta
'End Sub
'
'Private Sub cmdGrabar_Click()
''On Error GoTo ControlError
'Dim oContFun As NContFunciones
'Dim oGrabar As NPigContrato
'Dim oMovAnterior As DPigFunciones
'Dim lnMovNroAnt As Long
'Dim lsMovNro As String
'Dim lsFechaHoraGrab As String
'Dim lsCuenta As String
'Dim lsLote As String
'
'    lsCuenta = AXCodCta.NroCuenta
'
'    If MsgBox(" Grabar Anulación de Credito Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
'        cmdGrabar.Enabled = False
'        'Obtiene el Mov Nro Anterior
'        Set oMovAnterior = New DPigFunciones
'            lnMovNroAnt = oMovAnterior.dObtieneMovNroAnterior(lsCuenta, gPigOpeRegContrato)
'        Set oMovAnterior = Nothing
'
'        'Genera el Mov Nro
'        Set oContFun = New NContFunciones
'            lsMovNro = oContFun.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'        Set oContFun = Nothing
'
'        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
'        Set oGrabar = New NPigContrato
'            Call oGrabar.nAnulaCredPignoraticio(lsCuenta, gPigEstAnula, lsFechaHoraGrab, lsMovNro, lnMovNroAnt, lblPrestamo, True)
'        Set oGrabar = Nothing
'
'        Limpiar
'        cmdBuscar.Enabled = True
'        AXCodCta.Enabled = True
'        AXCodCta.SetFocus
'    Else
'        MsgBox " Grabación cancelada ", vbInformation, " Aviso "
'    End If
'Exit Sub
'
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
'
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
'    Limpiar
'    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
'End Sub
'
