VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGarantiaValorTasacMobilOtras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OTRAS TASACIONES MOBILIARIAS"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9450
   Icon            =   "frmGarantiaValorTasacMobilOtras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   9450
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3675
      TabIndex        =   14
      ToolTipText     =   "Aceptar"
      Top             =   3570
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4740
      TabIndex        =   15
      ToolTipText     =   "Cancelar"
      Top             =   3570
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   5953
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tasación Mobiliaria"
      TabPicture(0)   =   "frmGarantiaValorTasacMobilOtras.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraCaracteristicas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraValor"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraTasador"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame fraTasador 
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
         Height          =   1040
         Left            =   120
         TabIndex        =   27
         Top             =   2205
         Width           =   9015
         Begin VB.TextBox txtTasacionTC 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   8040
            TabIndex        =   9
            Text            =   "0.00"
            Top             =   240
            Width           =   840
         End
         Begin VB.TextBox txtTasadorNombre 
            Height          =   285
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   650
            Width           =   2685
         End
         Begin VB.TextBox txtTasadorDNI 
            Height          =   285
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   645
            Width           =   1000
         End
         Begin VB.TextBox txtTasadorREPEV 
            Height          =   285
            Left            =   7680
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   650
            Width           =   1200
         End
         Begin SICMACT.TxtBuscar txtTasadorCod 
            Height          =   255
            Left            =   1200
            TabIndex        =   10
            Top             =   645
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
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
         End
         Begin MSMask.MaskEdBox txtTasacionFecha 
            Height          =   330
            Left            =   1200
            TabIndex        =   8
            Top             =   240
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Tasador :"
            Height          =   195
            Left            =   150
            TabIndex        =   32
            Top             =   660
            Width           =   675
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Fecha :"
            Height          =   195
            Left            =   150
            TabIndex        =   31
            Top             =   285
            Width           =   540
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "T/C Tasación :"
            Height          =   195
            Left            =   6840
            TabIndex        =   30
            Top             =   285
            Width           =   1080
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "DNI : "
            Height          =   195
            Left            =   5430
            TabIndex        =   29
            Top             =   660
            Width           =   420
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "REPEV : "
            Height          =   195
            Left            =   6960
            TabIndex        =   28
            Top             =   660
            Width           =   675
         End
      End
      Begin VB.Frame fraValor 
         Caption         =   "Valor"
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
         Height          =   1845
         Left            =   5760
         TabIndex        =   23
         Top             =   360
         Width           =   3375
         Begin VB.TextBox txtVRM 
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
            Height          =   285
            Left            =   1680
            MaxLength       =   15
            TabIndex        =   7
            Text            =   "0.00"
            Top             =   1000
            Width           =   1560
         End
         Begin VB.TextBox txtValorComercial 
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
            Height          =   285
            Left            =   1680
            MaxLength       =   15
            TabIndex        =   6
            Text            =   "0.00"
            Top             =   650
            Width           =   1560
         End
         Begin VB.ComboBox cmbMoneda 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "V.R.M :"
            Height          =   195
            Left            =   240
            TabIndex        =   26
            Top             =   1015
            Width           =   540
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Moneda :"
            Height          =   195
            Left            =   240
            TabIndex        =   25
            Top             =   290
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor Comercial :"
            Height          =   195
            Left            =   240
            TabIndex        =   24
            Top             =   680
            Width           =   1185
         End
      End
      Begin VB.Frame fraCaracteristicas 
         Caption         =   "Características"
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
         Height          =   1845
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   5535
         Begin VB.TextBox txtDescripcion 
            Height          =   285
            Left            =   1200
            TabIndex        =   0
            Top             =   270
            Width           =   4125
         End
         Begin VB.TextBox txtNroSerie 
            Height          =   285
            Left            =   1200
            TabIndex        =   1
            Top             =   650
            Width           =   4125
         End
         Begin VB.TextBox txtMarca 
            Height          =   285
            Left            =   1200
            TabIndex        =   2
            Top             =   1000
            Width           =   1605
         End
         Begin VB.TextBox txtModelo 
            Height          =   285
            Left            =   3720
            TabIndex        =   3
            Top             =   1000
            Width           =   1605
         End
         Begin VB.TextBox txtAnioFabricacion 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4680
            MaxLength       =   4
            TabIndex        =   4
            Top             =   1370
            Width           =   645
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Marca :"
            Height          =   195
            Left            =   150
            TabIndex        =   22
            Top             =   1035
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "N° Serie :"
            Height          =   195
            Left            =   150
            TabIndex        =   21
            Top             =   645
            Width           =   675
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descripción :"
            Height          =   195
            Left            =   150
            TabIndex        =   20
            Top             =   290
            Width           =   930
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Modelo :"
            Height          =   195
            Left            =   2880
            TabIndex        =   19
            Top             =   1035
            Width           =   615
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Año Fabricación :"
            Height          =   195
            Left            =   2880
            TabIndex        =   18
            Top             =   1410
            Width           =   1245
         End
      End
   End
End
Attribute VB_Name = "frmGarantiaValorTasacMobilOtras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************************************************
'** Nombre : frmGarantiaValorTasacMobilOtras
'** Descripción : Para registro/edición/consulta de Otras Valorizaciones Mobiliarias creado segun TI-ERS063-2014
'** Creación : EJVG, 20150205 06:43:01 PM
'***************************************************************************************************************
Option Explicit

Dim fbRegistrar As Boolean
Dim fbEditar As Boolean
Dim fbConsultar As Boolean

Dim fbPrimero As Boolean
Dim fbOk As Boolean

Dim fnMoneda As Moneda
Dim fvValorTasacionMobiliariaOtras As tValorTasacionMobiliariaOtras
Dim fvValorTasacionMobiliariaOtras_ULT As tValorTasacionMobiliariaOtras

Private Sub cmdAceptar_Click()
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim lsCtaCod As String
    Dim lsValida As String
    
    On Error GoTo ErrAceptar
    If Not validarDatos Then Exit Sub
    
    fvValorTasacionMobiliariaOtras.sDescripcion = Trim(txtDescripcion.Text)
    fvValorTasacionMobiliariaOtras.sNroSerie = Trim(txtNroSerie.Text)
    fvValorTasacionMobiliariaOtras.sMarca = Trim(txtMarca.Text)
    fvValorTasacionMobiliariaOtras.sModelo = Trim(txtModelo.Text)
    fvValorTasacionMobiliariaOtras.nAnioFabricacion = CInt(txtAnioFabricacion.Text)
    fnMoneda = CInt(Trim(Right(cmbMoneda.Text, 2)))
    fvValorTasacionMobiliariaOtras.nValorComercial = CCur(txtValorComercial.Text)
    fvValorTasacionMobiliariaOtras.nVRM = CCur(txtVRM.Text)
    fvValorTasacionMobiliariaOtras.dTasacion = CDate(txtTasacionFecha.Text)
    fvValorTasacionMobiliariaOtras.nTasacionTC = CCur(txtTasacionTC.Text)
    fvValorTasacionMobiliariaOtras.sTasadorCod = txtTasadorCod.psCodigoPersona
    fvValorTasacionMobiliariaOtras.sTasadorNombre = txtTasadorNombre.Text
    fvValorTasacionMobiliariaOtras.sTasadorDNI = txtTasadorDNI.Text
    fvValorTasacionMobiliariaOtras.sTasadorREPEV = txtTasadorREPEV.Text
    
    fbOk = True
    Unload Me
    Exit Sub
ErrAceptar:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub Form_Load()
    fbOk = False
    Screen.MousePointer = 11
    
    CargarControles
    LimpiarControles

    If fbEditar Or fbConsultar Then
        txtDescripcion.Text = fvValorTasacionMobiliariaOtras.sDescripcion
        txtNroSerie.Text = fvValorTasacionMobiliariaOtras.sNroSerie
        txtMarca.Text = fvValorTasacionMobiliariaOtras.sMarca
        txtModelo.Text = fvValorTasacionMobiliariaOtras.sModelo
        txtAnioFabricacion.Text = fvValorTasacionMobiliariaOtras.nAnioFabricacion
        cmbMoneda.ListIndex = IndiceListaCombo(cmbMoneda, fnMoneda)
        txtValorComercial.Text = Format(fvValorTasacionMobiliariaOtras.nValorComercial, "#,##0.00")
        txtVRM.Text = Format(fvValorTasacionMobiliariaOtras.nVRM, "#,##0.00")
        txtTasacionFecha.Text = Format(fvValorTasacionMobiliariaOtras.dTasacion, gsFormatoFechaView)
        txtTasacionTC.Text = Format(fvValorTasacionMobiliariaOtras.nTasacionTC, gsFormatoNumeroView)
        txtTasadorCod.psCodigoPersona = fvValorTasacionMobiliariaOtras.sTasadorCod
        txtTasadorCod.Text = fvValorTasacionMobiliariaOtras.sTasadorCod
        txtTasadorNombre.Text = fvValorTasacionMobiliariaOtras.sTasadorNombre
        txtTasadorDNI.Text = fvValorTasacionMobiliariaOtras.sTasadorDNI
        txtTasadorREPEV.Text = fvValorTasacionMobiliariaOtras.sTasadorREPEV
                
        If fbConsultar Then
            fraCaracteristicas.Enabled = False
            fraValor.Enabled = False
            fraTasador.Enabled = False
            cmdAceptar.Enabled = False
        End If
    End If
    
    If fbPrimero Then
        fnMoneda = gMonedaNacional
    End If
    cmbMoneda.ListIndex = IndiceListaCombo(cmbMoneda, fnMoneda)
    cmbMoneda.Enabled = False
    
    If fbRegistrar Then
        txtDescripcion.Text = fvValorTasacionMobiliariaOtras_ULT.sDescripcion
        txtNroSerie.Text = fvValorTasacionMobiliariaOtras_ULT.sNroSerie
        txtMarca.Text = fvValorTasacionMobiliariaOtras_ULT.sMarca
        txtModelo.Text = fvValorTasacionMobiliariaOtras_ULT.sModelo
        txtAnioFabricacion.Text = IIf(fvValorTasacionMobiliariaOtras_ULT.nAnioFabricacion = 0, Year(gdFecSis), fvValorTasacionMobiliariaOtras_ULT.nAnioFabricacion)
        'cmbMoneda.ListIndex = IndiceListaCombo(cmbMoneda, fnMoneda)
        txtValorComercial.Text = Format(fvValorTasacionMobiliariaOtras_ULT.nValorComercial, "#,##0.00")
        txtVRM.Text = Format(fvValorTasacionMobiliariaOtras_ULT.nVRM, "#,##0.00")
        txtTasacionFecha.Text = Format(IIf(Year(fvValorTasacionMobiliariaOtras_ULT.dTasacion) <= 1950, gdFecSis, fvValorTasacionMobiliariaOtras_ULT.dTasacion), gsFormatoFechaView)
        txtTasacionTC.Text = Format(fvValorTasacionMobiliariaOtras_ULT.nTasacionTC, gsFormatoNumeroView)
        txtTasadorCod.psCodigoPersona = fvValorTasacionMobiliariaOtras_ULT.sTasadorCod
        txtTasadorCod.Text = fvValorTasacionMobiliariaOtras_ULT.sTasadorCod
        txtTasadorNombre.Text = fvValorTasacionMobiliariaOtras_ULT.sTasadorNombre
        txtTasadorDNI.Text = fvValorTasacionMobiliariaOtras_ULT.sTasadorDNI
        txtTasadorREPEV.Text = fvValorTasacionMobiliariaOtras_ULT.sTasadorREPEV
    End If
    
    If fbRegistrar Then
        Caption = "OTRAS TASACIONES MOBILIARIAS [ NUEVO ]"
    End If
    If fbConsultar Then
        Caption = "OTRAS TASACIONES MOBILIARIAS [ CONSULTAR ]"
    End If
    If fbEditar Then
        Caption = "OTRAS TASACIONES MOBILIARIAS [ EDITAR ]"
    End If
    
    Screen.MousePointer = 0
End Sub
Public Function Registrar(ByVal pbPrimero As Boolean, ByRef pnMoneda As Moneda, ByRef pvValorTasacionMobiliariaOtras As tValorTasacionMobiliariaOtras, ByRef pvValorTasacionMobiliariaOtras_ULT As tValorTasacionMobiliariaOtras) As Boolean
    fbRegistrar = True
    fbPrimero = pbPrimero
    fnMoneda = pnMoneda
    fvValorTasacionMobiliariaOtras = pvValorTasacionMobiliariaOtras
    fvValorTasacionMobiliariaOtras_ULT = pvValorTasacionMobiliariaOtras_ULT
    Show 1
    pnMoneda = fnMoneda
    pvValorTasacionMobiliariaOtras = fvValorTasacionMobiliariaOtras
    
    Registrar = fbOk
End Function
Public Function Editar(ByVal pbPrimero As Boolean, ByRef pnMoneda As Moneda, ByRef pvValorTasacionMobiliariaOtras As tValorTasacionMobiliariaOtras, ByRef pvValorTasacionMobiliariaOtras_ULT As tValorTasacionMobiliariaOtras) As Boolean
    fbEditar = True
    fbPrimero = pbPrimero
    fnMoneda = pnMoneda
    fvValorTasacionMobiliariaOtras = pvValorTasacionMobiliariaOtras
    fvValorTasacionMobiliariaOtras_ULT = pvValorTasacionMobiliariaOtras_ULT
    Show 1
    pnMoneda = fnMoneda
    pvValorTasacionMobiliariaOtras = fvValorTasacionMobiliariaOtras
    
    Editar = fbOk
End Function
Public Sub Consultar(ByVal pnMoneda As Moneda, ByRef pvValorTasacionMobiliariaOtras As tValorTasacionMobiliariaOtras)
    fbConsultar = True
    fnMoneda = pnMoneda
    fvValorTasacionMobiliariaOtras = pvValorTasacionMobiliariaOtras
    Show 1
End Sub
Private Sub CargarControles()
    Dim oCons As New COMDConstantes.DCOMConstantes
    Dim rs As New ADODB.Recordset
    
    Set rs = oCons.RecuperaConstantes(1011)
    Call Llenar_Combo_con_Recordset(rs, cmbMoneda)
    
    RSClose rs
    Set oCons = Nothing
End Sub
Private Sub cmdCancelar_Click()
    fbOk = False
    Unload Me
End Sub
Private Sub LimpiarControles()
    txtDescripcion.Text = ""
    txtNroSerie.Text = ""
    txtMarca.Text = ""
    txtModelo.Text = ""
    txtAnioFabricacion.Text = Year(gdFecSis)
    cmbMoneda.ListIndex = -1
    txtValorComercial.Text = "0.00"
    txtVRM.Text = "0.00"
    txtTasacionFecha.Text = Format(gdFecSis, gsFormatoFechaView)
    txtTasacionTC.Text = "0.00"
    txtTasadorCod.Text = ""
    txtTasadorNombre.Text = ""
    txtTasadorDNI.Text = ""
    txtTasadorREPEV.Text = ""
End Sub
Private Sub txtAnioFabricacion_LostFocus()
    txtAnioFabricacion.Text = Format(txtAnioFabricacion.Text, "0000")
End Sub
Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtNroSerie
    End If
End Sub
Private Sub txtNroSerie_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtMarca
    End If
End Sub
Private Sub txtMarca_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtModelo
    End If
End Sub
Private Sub txtMarca_LostFocus()
    txtMarca.Text = UCase(txtMarca.Text)
End Sub
Private Sub txtModelo_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtAnioFabricacion
    End If
End Sub
Private Sub txtModelo_LostFocus()
    txtModelo.Text = UCase(txtModelo.Text)
End Sub
Private Sub txtAnioFabricacion_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        If cmbMoneda.Enabled And cmbMoneda.Visible Then
            EnfocaControl cmbMoneda
        Else
            EnfocaControl txtValorComercial
        End If
    End If
End Sub
Private Sub cmbMoneda_Click()
    Dim lnMoneda As Moneda
    Dim lsColor As Long
    lnMoneda = val(Trim(Right(cmbMoneda.Text, 3)))
    If lnMoneda = gMonedaNacional Then
        lsColor = &H80000005
    Else
        lsColor = &HC0FFC0
    End If
    
    txtValorComercial.BackColor = lsColor
    txtVRM.BackColor = lsColor
End Sub
Private Sub cmbMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtValorComercial
    End If
End Sub
Private Sub txtDescripcion_LostFocus()
    txtDescripcion.Text = UCase(txtDescripcion.Text)
End Sub
Private Sub txtTasadorCod_EmiteDatos()
    Dim oPersona As New COMDPersona.DCOMPersona
    Dim rsPersona As New ADODB.Recordset
    Dim lsDNI As String, lsREPEV As String
    
    Screen.MousePointer = 11
    
    txtTasadorNombre.Text = ""
    txtTasadorDNI.Text = ""
    txtTasadorREPEV.Text = ""
    If txtTasadorCod.Text <> "" Then
        Set rsPersona = oPersona.RecuperaDatosPersonaxGarantia(txtTasadorCod.psCodigoPersona)
        If Not rsPersona.EOF Then
            lsDNI = rsPersona!DNI
            lsREPEV = rsPersona!REPEV
        End If
        txtTasadorNombre.Text = txtTasadorCod.psDescripcion
        txtTasadorDNI.Text = lsDNI
        txtTasadorREPEV.Text = lsREPEV
    End If
    
    RSClose rsPersona
    Set oPersona = Nothing
    
    Screen.MousePointer = 0
End Sub
Private Sub txtTasadorNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtTasadorDNI
    End If
End Sub
Private Sub txtTasadorDNI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtTasadorREPEV
    End If
End Sub
Private Sub txtTasadorREPEV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl cmdAceptar
    End If
End Sub
Private Sub txtValorComercial_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtValorComercial, KeyAscii, 15, 2)
    If KeyAscii = 13 Then
        EnfocaControl txtVRM
    End If
End Sub
Private Sub txtValorComercial_LostFocus()
    txtValorComercial.Text = Format(txtValorComercial.Text, "#,##0.00")
End Sub
Private Sub TxtVRM_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtVRM, KeyAscii, 15, 2)
    If KeyAscii = 13 Then
        EnfocaControl txtTasacionFecha
    End If
End Sub
Private Sub TxtVRM_LostFocus()
    txtVRM.Text = Format(txtVRM.Text, "#,##0.00")
End Sub
Private Sub txtTasacionFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtTasacionTC
    End If
End Sub
Private Sub txtTasacionTC_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtTasacionTC, KeyAscii, 6, 4)
    If KeyAscii = 13 Then
        EnfocaControl txtTasadorCod
    End If
End Sub
Private Sub txtTasacionTC_LostFocus()
    txtTasacionTC.Text = Format(txtTasacionTC.Text, gsFormatoNumeroView)
End Sub
Private Sub txtTasadorCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl cmdAceptar
    End If
End Sub
Private Function validarDatos() As Boolean
    Dim lsFecha As String
        
    If Len(Trim(txtDescripcion.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la Descripción", vbInformation, "Aviso"
        EnfocaControl txtDescripcion
        Exit Function
    End If
    If Len(Trim(txtNroSerie.Text)) = 0 Then
        MsgBox "Ud. debe ingresar el Nro. de Serie", vbInformation, "Aviso"
        EnfocaControl txtNroSerie
        Exit Function
    End If
    If Len(Trim(txtMarca.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la Marca", vbInformation, "Aviso"
        EnfocaControl txtMarca
        Exit Function
    End If
    If Len(Trim(txtModelo.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la Modelo", vbInformation, "Aviso"
        EnfocaControl txtModelo
        Exit Function
    End If
    If Not IsNumeric(txtAnioFabricacion.Text) Then
        MsgBox "Ud. debe de especificar el Año de Fabricación", vbInformation, "Aviso"
        EnfocaControl txtAnioFabricacion
        Exit Function
    Else
        If CInt(txtAnioFabricacion.Text) <= 1890 Then
            MsgBox "Ud. debe de especificar el Año de Fabricación", vbInformation, "Aviso"
            EnfocaControl txtAnioFabricacion
            Exit Function
        End If
    End If
    
    If cmbMoneda.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar la moneda", vbInformation, "Aviso"
        EnfocaControl cmbMoneda
        Exit Function
    End If
    If Not IsNumeric(txtValorComercial.Text) Then
        MsgBox "Ud. debe de especificar el Valor Comercial", vbInformation, "Aviso"
        EnfocaControl txtValorComercial
        Exit Function
    Else
        If CCur(txtValorComercial.Text) <= 0 Then
            MsgBox "El Valor Comercial debe ser mayor a cero", vbInformation, "Aviso"
            EnfocaControl txtValorComercial
            Exit Function
        End If
    End If
    If Not IsNumeric(txtVRM.Text) Then
        MsgBox "Ud. debe de especificar el Valor de Realización", vbInformation, "Aviso"
        EnfocaControl txtVRM
        Exit Function
    Else
        If CCur(txtVRM.Text) <= 0 Then
            MsgBox "El Valor de Realización debe ser mayor a cero", vbInformation, "Aviso"
            EnfocaControl txtVRM
            Exit Function
        End If
        If CCur(txtVRM.Text) > CCur(txtValorComercial.Text) Then
            MsgBox "El Valor de Realización no puede ser mayor al Valor Comercial", vbInformation, "Aviso"
            EnfocaControl txtVRM
            Exit Function
        End If
    End If

    lsFecha = ValidaFecha(txtTasacionFecha.Text)
    If Len(Trim(lsFecha)) > 0 Then
        MsgBox lsFecha, vbInformation, "Aviso"
        EnfocaControl txtTasacionFecha
        Exit Function
    Else
        If Not fvValorTasacionMobiliariaOtras_ULT.bMigrado Then
            If CDate(txtTasacionFecha.Text) <= fvValorTasacionMobiliariaOtras_ULT.dTasacion Then
                MsgBox "La actual fecha de Tasación no puede ser menor o igual a la última Tasación: " & Format(fvValorTasacionMobiliariaOtras_ULT.dTasacion, gsFormatoFechaView), vbInformation, "Aviso"
                EnfocaControl txtTasacionFecha
                Exit Function
            End If
        End If
        If CDate(txtTasacionFecha.Text) > gdFecSis Then
            MsgBox "La fecha de Tasación no puede ser mayor a la fecha de Sistema", vbInformation, "Aviso"
            EnfocaControl txtTasacionFecha
            Exit Function
        End If
    End If
    If Not IsNumeric(txtTasacionTC.Text) Then
        MsgBox "Ud. debe de especificar el Tipo de Cambio de Tasación", vbInformation, "Aviso"
        EnfocaControl txtTasacionTC
        Exit Function
    Else
        If CCur(txtTasacionTC.Text) <= 0 Then
            MsgBox "El Tipo de Cambio de Tasación debe ser mayor a cero", vbInformation, "Aviso"
            EnfocaControl txtTasacionTC
            Exit Function
        End If
        If CCur(txtTasacionTC.Text) > 10 Then
            MsgBox "Verifique el Tipo de Cambio de Tasación", vbInformation, "Aviso"
            EnfocaControl txtTasacionTC
            Exit Function
        End If
    End If
    If Len(Trim(txtTasadorCod.Text)) <> 13 Or Len(txtTasadorCod.psCodigoPersona) <> 13 Then
        MsgBox "Ud. debe especificar al Tasador", vbInformation, "Aviso"
        EnfocaControl txtTasadorCod
        Exit Function
    Else
        If txtTasadorCod.PersPersoneria <> gPersonaNat Then
            MsgBox "El Tasador debe ser una persona Natural", vbInformation, "Aviso"
            EnfocaControl txtTasadorCod
            Exit Function
        Else
            If Len(Trim(txtTasadorDNI.Text)) <> 8 Then
                MsgBox "El Tasador no cuenta con Documento DNI", vbInformation, "Aviso"
                EnfocaControl txtTasadorCod
                Exit Function
            End If
            If Len(Trim(txtTasadorREPEV.Text)) = 0 Then
                MsgBox "El Tasador no cuenta con Documento REPEV", vbInformation, "Aviso"
                EnfocaControl txtTasadorCod
                Exit Function
            End If
        End If
    End If
        
    validarDatos = True
End Function

