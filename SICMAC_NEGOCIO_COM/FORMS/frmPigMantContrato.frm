VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmPigMantContrato 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Contrato"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   Icon            =   "frmPigMantContrato.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   9600
      TabIndex        =   16
      Top             =   6555
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   10725
      TabIndex        =   15
      Top             =   6555
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   8520
      TabIndex        =   14
      Top             =   6555
      Width           =   975
   End
   Begin VB.Frame FraContenedor 
      Height          =   6450
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   11790
      Begin VB.Frame FraContenedor 
         Caption         =   "Cliente(s)"
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
         Height          =   1005
         Index           =   6
         Left            =   90
         TabIndex        =   19
         Top             =   705
         Width           =   11610
         Begin SICMACT.FlexEdit feCte 
            Height          =   705
            Left            =   75
            TabIndex        =   20
            Top             =   195
            Width           =   11430
            _ExtentX        =   20161
            _ExtentY        =   1244
            Cols0           =   5
            FixedCols       =   0
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "Codigo-Nombre/Razon Social-Doc.Iden-Direccion-Fono"
            EncabezadosAnchos=   "1600-3500-1500-3390-1300"
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
            ColumnasAEditar =   "X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-C-C-C"
            FormatosEdit    =   "0-0-0-0-0"
            TextArray0      =   "Codigo"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   1605
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.CommandButton cmdBuscar 
         Height          =   390
         Left            =   11175
         Picture         =   "frmPigMantContrato.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Buscar ..."
         Top             =   255
         Width           =   420
      End
      Begin VB.Frame FraContenedor 
         Caption         =   "Joya(s)"
         Enabled         =   0   'False
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
         Height          =   4035
         Index           =   4
         Left            =   75
         TabIndex        =   7
         Top             =   1710
         Width           =   11640
         Begin VB.TextBox TxtTotalB 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1605
            TabIndex        =   8
            Top             =   3675
            Width           =   705
         End
         Begin SICMACT.FlexEdit FEJoyas 
            Height          =   3330
            Left            =   45
            TabIndex        =   22
            Top             =   225
            Width           =   11535
            _ExtentX        =   20346
            _ExtentY        =   5874
            Cols0           =   13
            HighLight       =   1
            AllowUserResizing=   1
            RowSizingMode   =   1
            EncabezadosNombres=   "Num-Tipo-SubTipo-Material-Estado-Observacion-PBruto-PNeto-Tasacion-TasAdicion-ObsAdicion-p-Item"
            EncabezadosAnchos=   "350-1000-1050-1000-1000-2300-700-700-900-900-1500-0-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
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
            ColumnasAEditar =   "X-1-2-X-X-5-X-X-X-X-10-X-X"
            ListaControles  =   "0-3-3-0-0-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-L-L-L-R-R-C-C-L-R-C"
            FormatosEdit    =   "0-1-1-1-1-0-2-2-4-4-0-4-3"
            TextArray0      =   "Num"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   345
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin VB.Label Label1 
            Caption         =   "Total Balanza"
            Height          =   225
            Left            =   510
            TabIndex        =   13
            Top             =   3705
            Width           =   1080
         End
         Begin VB.Label LblPBruto 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   6975
            TabIndex        =   12
            Top             =   3630
            Width           =   810
         End
         Begin VB.Label LblPNeto 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   7770
            TabIndex        =   11
            Top             =   3630
            Width           =   825
         End
         Begin VB.Label LblTasacion 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   8565
            TabIndex        =   10
            Top             =   3630
            Width           =   855
         End
         Begin VB.Label LblTasacionAdic 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   9390
            TabIndex        =   9
            Top             =   3630
            Width           =   930
         End
      End
      Begin VB.Frame FraContenedor 
         Enabled         =   0   'False
         Height          =   630
         Index           =   1
         Left            =   90
         TabIndex        =   1
         Top             =   5730
         Width           =   11610
         Begin SICMACT.EditMoney txtPrestamo 
            Height          =   285
            Left            =   10050
            TabIndex        =   2
            Top             =   225
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   503
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            BackColor       =   12648447
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox MskFechaVenc 
            Height          =   300
            Left            =   3615
            TabIndex        =   3
            Top             =   210
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblPlazo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   930
            TabIndex        =   21
            Top             =   210
            Width           =   765
         End
         Begin VB.Label Label2 
            Caption         =   "Plazo"
            Height          =   255
            Left            =   405
            TabIndex        =   6
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Vencimiento"
            Height          =   255
            Left            =   2160
            TabIndex        =   5
            Top             =   240
            Width           =   1410
         End
         Begin VB.Label Label4 
            Caption         =   "Préstamo"
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
            Left            =   9060
            TabIndex        =   4
            Top             =   270
            Width           =   915
         End
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   375
         Left            =   195
         TabIndex        =   18
         Top             =   255
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         Texto           =   "Crédito"
         EnabledCta      =   -1  'True
      End
   End
   Begin VB.Label lblEstado 
      Caption         =   "Label5"
      Height          =   270
      Left            =   5085
      TabIndex        =   23
      Top             =   6015
      Width           =   1215
   End
End
Attribute VB_Name = "frmPigMantContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
''Carga datos de Joyas
'Dim rsTipoJoya As Recordset
'Dim rsSTipoJoya As Recordset
'Dim rsMaterial As Recordset
'Dim rsPlazos As Recordset
'Dim rsEstadoJoya As Recordset
'Dim lsTipo As String
'Dim lsSubTipo As String
'
'Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
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
'    feCte.TextMatrix(1, 4) = oPers.sPersTelefono
'Set oPers = Nothing
'
'lsEstado = gPigEstRegis & "," & gPigEstDesemb
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
'
'Exit Sub
'
'ControlError:   ' Rutina de control de errores.
'    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & '        " Avise al Area de Sistemas ", vbInformation, " Aviso "
'
'End Sub
'Private Sub cmdGrabar_Click()
'
'Dim oContFunc As NContFunciones
'Dim oGrabarMod As NPigContrato
'Dim oImprime As NPigImpre
'Dim oPrevio As previo.clsPrevio
'Dim rsJoyas As Recordset
'
'Dim lsMovNro As String
'Dim lsFechaHoraGrab As String
'Dim lsCuenta As String
'Dim lsLote As String
'Dim lsCadImprimir As String
'
'lsCuenta = AXCodCta.NroCuenta
'
'    ' ===================== OJO falta validar los datos ==================
'    If MsgBox(" Grabar Cambios del contrato? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
'        cmdGrabar.Enabled = False
'
'        'Genera el Mov Nro
'        Set oContFunc = New NContFunciones
'            lsMovNro = oContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'        Set oContFunc = Nothing
'
'        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
'        Set rsJoyas = FEJoyas.GetRsNew
'        Set oGrabarMod = New NPigContrato
'            'Grabar la Modificacion
'            Call oGrabarMod.nModificaCredPignoraticio(lsCuenta, lsFechaHoraGrab, lblEstado, rsJoyas, lsMovNro, 1)
'        Set oGrabarMod = Nothing
'
'        ' ========== Imprimir ===========
'        If MsgBox("Imprimir Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
'            Set oImprime = New NPigImpre
'                lsCadImprimir = oImprime.ImpreContratoPignoraticio(AXCodCta.NroCuenta, True, , , , , , , , , , '                                        , gsNomAge, , , , , , , , , , , 0)
'
'            Set oImprime = Nothing
'            Set oPrevio = New previo.clsPrevio
'            oPrevio.Show lsCadImprimir, "Duplicado de Contrato"
'            Do While True
'                If MsgBox("Reimprimir Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
'                    oPrevio.Show lsCadImprimir, "Duplicado de Contrato"
'                Else
'                    Set oPrevio = Nothing
'                    Exit Do
'                End If
'            Loop
'        End If
'
'        Limpiar
'        cmdBuscar.Enabled = True
'
'        AXCodCta.Enabled = True
'        AXCodCta.SetFocus
'    Else
'        MsgBox " Grabación cancelada ", vbInformation, " Aviso "
'    End If
'
'    'MsgBox " Falta información " & vbCr & " No se puede Grabar Contrato ", vbInformation, " Aviso "
'
'Exit Sub
'
'ControlError:   ' Rutina de control de errores.
'    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & '        " Avise al Area de Sistemas ", vbInformation, " Aviso "
'
'End Sub
'
'Private Sub cmdSalir_Click()
'Unload Me
'End Sub
'
'Private Sub Limpiar()
'
'AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
'LblPBruto = ""
'LblPNeto = ""
'LblTasacion = ""
'LblTasacionAdic = ""
'txtPrestamo = ""
'TxtTotalB = ""
'feCte.Clear
'feCte.Rows = 2
'feCte.FormaCabecera
'
'FEJoyas.Clear
'FEJoyas.Rows = 2
'FEJoyas.FormaCabecera
'
'End Sub
'
'Private Sub FEJoyas_Click()
'Dim oCons As DConstante
'Set oCons = New DConstante
'
'Select Case FEJoyas.Col
'Case 1
'    Set rsTipoJoya = oCons.RecuperaConstantes(gColocPigTipoJoya, , "C.cConsDescripcion")
'    FEJoyas.CargaCombo rsTipoJoya
'    Set rsTipoJoya = Nothing
'Case 2
'    Set rsSTipoJoya = oCons.RecuperaConstantes(gColocPigSubTipoJoya, , "C.cConsDescripcion")
'    FEJoyas.CargaCombo rsSTipoJoya
'    Set rsSTipoJoya = Nothing
'End Select
'
'Set oCons = Nothing
'End Sub
'
'Private Sub FEJoyas_RowColChange()
'Dim oCons As DConstante
'Set oCons = New DConstante
'
'Select Case FEJoyas.Col
'Case 1
'    Set rsTipoJoya = oCons.RecuperaConstantes(gColocPigTipoJoya, , "C.cConsDescripcion")
'    FEJoyas.CargaCombo rsTipoJoya
'    Set rsTipoJoya = Nothing
'Case 2
'    Set rsSTipoJoya = oCons.RecuperaConstantes(gColocPigSubTipoJoya, , "C.cConsDescripcion")
'    FEJoyas.CargaCombo rsSTipoJoya
'    Set rsSTipoJoya = Nothing
'End Select
'
'Set oCons = Nothing
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
'
'    Limpiar
'    FEJoyas_RowColChange
'    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
'    Me.Icon = LoadPicture(App.path & "\Graficos\cm.ico")
'
'End Sub
'
'Private Sub BuscaContrato(ByVal psNroContrato As String)
'Dim rs As ADODB.Recordset
'Dim oValContrato As nPigValida
'Dim oPigContrato As DPigContrato
'Dim lsmensaje As String
'Dim lnDias As Integer
'
'On Error GoTo ControlError
'
'    'Valida Contrato
'    Set rs = New ADODB.Recordset
'    Set oValContrato = New nPigValida
'        '27-12
'        'Set rs = oValContrato.nValidaDuplicadoContratoCredPignoraticio(psNroContrato, lsmensaje )
'        Set rs = oValContrato.nValidaDuplicadoContratoCredPignoraticio(psNroContrato)
'        If Trim(lsmensaje) <> "" Then
'             MsgBox lsmensaje, vbInformation, "Aviso"
'             Exit Sub
'        End If
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
'    Set rs = oPigContrato.dObtieneDatosCreditoPignoraticioPersonas(psNroContrato)
'    feCte.TextMatrix(1, 0) = rs!cPersCod
'    feCte.TextMatrix(1, 1) = PstaNombre(rs!cPersNombre)
'    feCte.TextMatrix(1, 2) = IIf(IsNull(rs!NroDNI), " ", rs!NroDNI)
'    feCte.TextMatrix(1, 3) = IIf(IsNull(rs!cPersDireccDomicilio), "", rs!cPersDireccDomicilio) + " " + IIf(IsNull(rs!Zona), "", rs!Zona)
'    feCte.TextMatrix(1, 4) = IIf(IsNull(rs!cPersTelefono), "", rs!cPersTelefono)
'    Set rs = Nothing
'
'    Set rs = oPigContrato.dObtieneDatosContrato(psNroContrato, gPigTipoTasacNor)
'
'    If Not rs.EOF And Not rs.BOF Then
'        LblPBruto = Format(rs!nPBruto, "######.00")
'        LblPNeto = Format(rs!nPNeto, "######.00")
'        LblTasacion = Format(rs!nTasacion, "#######.00")
'        LblTasacionAdic = Format(rs!nTasacionAdicional, "#######.00")
'        txtPrestamo = Format(rs!nMontoCol, "#######.00")
'        lblPlazo = DateDiff("d", Format$(rs!dVigencia, "dd/mm/yyyy"), Format$(rs!dvenc, "dd/mm/yyyy"))
'        MskFechaVenc = Format$(rs!dvenc, "dd/mm/yyyy")
'        TxtTotalB = rs!nTotalBalanza
'        lblEstado = rs!nPrdEstado
'    End If
'
'    Set rs = Nothing
'    Set rs = oPigContrato.dObtieneDetalleJoyas(psNroContrato, 1)
'
'    Do While Not rs.EOF
'        FEJoyas.AdicionaFila
'        FEJoyas.TextMatrix(rs!Item, 0) = rs!Item
'        FEJoyas.TextMatrix(rs!Item, 1) = rs!Tipo & Space(75) & rs!ntipo
'        If rs!nSubTipo <> 0 Or Not IsNull(rs!nSubTipo) Then
'            FEJoyas.TextMatrix(rs!Item, 2) = rs!SubTipo & Space(75) & rs!nSubTipo
'        End If
'        FEJoyas.TextMatrix(rs!Item, 3) = rs!Material
'        FEJoyas.TextMatrix(rs!Item, 4) = rs!Estado
'        FEJoyas.TextMatrix(rs!Item, 5) = rs!Observacion
'        FEJoyas.TextMatrix(rs!Item, 6) = Format(rs!PBruto, "#######.00")
'        FEJoyas.TextMatrix(rs!Item, 7) = Format(rs!PNeto, "#######.00")
'        FEJoyas.TextMatrix(rs!Item, 8) = Format(rs!Tasacion, "#######.00")
'        FEJoyas.TextMatrix(rs!Item, 9) = Format(rs!TasAdicion, "#######.00")
'        FEJoyas.TextMatrix(rs!Item, 10) = rs!ObsAdicion
'        FEJoyas.TextMatrix(rs!Item, 12) = rs!Item
'        rs.MoveNext
'    Loop
'
'    FraContenedor(4).Enabled = True
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
