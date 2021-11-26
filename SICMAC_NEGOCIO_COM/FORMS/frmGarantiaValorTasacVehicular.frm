VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGarantiaValorTasacVehicular 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TASACIÓN VEHICULAR"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9735
   Icon            =   "frmGarantiaValorTasacVehicular.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   9735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4860
      TabIndex        =   20
      ToolTipText     =   "Cancelar"
      Top             =   4410
      Width           =   1000
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3795
      TabIndex        =   19
      ToolTipText     =   "Aceptar"
      Top             =   4410
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   7435
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Valor del Vehículo"
      TabPicture(0)   =   "frmGarantiaValorTasacVehicular.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTasador"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraValor"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraCaracteristicas"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkTasador"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.CheckBox chkTasador 
         Caption         =   "Tasador"
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
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   3000
         Width           =   1095
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
         Height          =   2595
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   5775
         Begin VB.TextBox txtAnioAdquisicionTienda 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4920
            MaxLength       =   4
            TabIndex        =   8
            Top             =   2130
            Width           =   645
         End
         Begin VB.TextBox txtAnioFabricacion 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4920
            MaxLength       =   4
            TabIndex        =   7
            Top             =   1750
            Width           =   645
         End
         Begin VB.TextBox txtNroChasis 
            Height          =   285
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   6
            Top             =   1750
            Width           =   1605
         End
         Begin VB.TextBox txtNroMotor 
            Height          =   285
            Left            =   3960
            MaxLength       =   30
            TabIndex        =   5
            Top             =   1380
            Width           =   1605
         End
         Begin VB.TextBox txtNroPlaca 
            Height          =   285
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   4
            Top             =   1380
            Width           =   1605
         End
         Begin VB.TextBox txtModelo 
            Height          =   285
            Left            =   3960
            MaxLength       =   100
            TabIndex        =   3
            Top             =   1000
            Width           =   1605
         End
         Begin VB.TextBox txtMarca 
            Height          =   285
            Left            =   1200
            MaxLength       =   100
            TabIndex        =   2
            Top             =   1000
            Width           =   1605
         End
         Begin VB.TextBox txtCarroceria 
            Height          =   285
            Left            =   1200
            MaxLength       =   255
            TabIndex        =   1
            Top             =   650
            Width           =   4365
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   285
            Left            =   1200
            MaxLength       =   255
            TabIndex        =   0
            Top             =   270
            Width           =   4365
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Año Adquisición en tienda :"
            Height          =   195
            Left            =   2880
            TabIndex        =   41
            Top             =   2160
            Width           =   1935
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Año Fabricación :"
            Height          =   195
            Left            =   2880
            TabIndex        =   40
            Top             =   1800
            Width           =   1245
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "N° de Chasis :"
            Height          =   195
            Left            =   150
            TabIndex        =   39
            Top             =   1805
            Width           =   1005
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "N° Motor :"
            Height          =   195
            Left            =   2880
            TabIndex        =   38
            Top             =   1425
            Width           =   720
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "N° Placa :"
            Height          =   195
            Left            =   150
            TabIndex        =   37
            Top             =   1420
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Modelo :"
            Height          =   195
            Left            =   2880
            TabIndex        =   36
            Top             =   1035
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descripción :"
            Height          =   195
            Left            =   150
            TabIndex        =   35
            Top             =   290
            Width           =   930
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Carrocería :"
            Height          =   195
            Left            =   150
            TabIndex        =   34
            Top             =   645
            Width           =   840
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Marca :"
            Height          =   195
            Left            =   150
            TabIndex        =   33
            Top             =   1035
            Width           =   540
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
         Height          =   2565
         Left            =   6000
         TabIndex        =   28
         Top             =   360
         Width           =   3375
         Begin VB.ComboBox cmbMoneda 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   240
            Width           =   1575
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
            TabIndex        =   10
            Text            =   "0.00"
            Top             =   650
            Width           =   1560
         End
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
            TabIndex        =   11
            Text            =   "0.00"
            Top             =   1000
            Width           =   1560
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor Comercial :"
            Height          =   195
            Left            =   240
            TabIndex        =   31
            Top             =   680
            Width           =   1185
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Moneda :"
            Height          =   195
            Left            =   240
            TabIndex        =   30
            Top             =   290
            Width           =   675
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "V.R.M :"
            Height          =   195
            Left            =   240
            TabIndex        =   29
            Top             =   1015
            Width           =   540
         End
      End
      Begin VB.Frame fraTasador 
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
         Height          =   1040
         Left            =   120
         TabIndex        =   22
         Top             =   3045
         Width           =   9255
         Begin VB.TextBox txtTasadorREPEV 
            Height          =   285
            Left            =   7920
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   650
            Width           =   1200
         End
         Begin VB.TextBox txtTasadorDNI 
            Height          =   285
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   645
            Width           =   1000
         End
         Begin VB.TextBox txtTasadorNombre 
            Height          =   285
            Left            =   2595
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   650
            Width           =   2925
         End
         Begin VB.TextBox txtTasacionTC 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   8280
            TabIndex        =   14
            Text            =   "0.00"
            Top             =   240
            Width           =   840
         End
         Begin SICMACT.TxtBuscar txtTasadorCod 
            Height          =   255
            Left            =   1200
            TabIndex        =   15
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
            TabIndex        =   13
            Top             =   240
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "REPEV : "
            Height          =   195
            Left            =   7200
            TabIndex        =   27
            Top             =   660
            Width           =   675
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "DNI : "
            Height          =   195
            Left            =   5670
            TabIndex        =   26
            Top             =   660
            Width           =   420
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "T/C Tasación :"
            Height          =   195
            Left            =   7080
            TabIndex        =   25
            Top             =   285
            Width           =   1080
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Fecha :"
            Height          =   195
            Left            =   150
            TabIndex        =   24
            Top             =   285
            Width           =   540
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Tasador :"
            Height          =   195
            Left            =   150
            TabIndex        =   23
            Top             =   660
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frmGarantiaValorTasacVehicular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'** Nombre : frmGarantiaValorTasacVehicular
'** Descripción : Para registro/edición/consulta de Valorización Vehicular creado segun TI-ERS063-2014
'** Creación : EJVG, 20150205 09:36:01 AM
'*****************************************************************************************************
Option Explicit

Dim fbRegistrar As Boolean
Dim fbEditar As Boolean
Dim fbConsultar As Boolean

Dim fbPrimero As Boolean
Dim fbOk As Boolean

Dim fnMoneda As Moneda
Dim fvValorTasacionVehicular As tValorTasacionVehicular
Dim fvValorTasacionVehicular_ULT As tValorTasacionVehicular

Private Sub chkTasador_Click()
    If chkTasador.value = 1 Then
        fraTasador.Enabled = True
    Else
        fraTasador.Enabled = False
    End If
End Sub
Private Sub chkTasador_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not EnfocaControl(txtTasacionFecha) Then
            EnfocaControl cmdAceptar
        End If
    End If
End Sub
Private Sub cmdAceptar_Click()
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim lsCtaCod As String
    Dim lsValida As String
    
    On Error GoTo ErrAceptar
    If Not validarDatos Then Exit Sub
    
    fvValorTasacionVehicular.sDescripcion = txtDescripcion.Text
    fvValorTasacionVehicular.sCarroceria = txtCarroceria.Text
    fvValorTasacionVehicular.sMarca = txtMarca.Text
    fvValorTasacionVehicular.sModelo = txtModelo.Text
    fvValorTasacionVehicular.sNroPlaca = txtNroPlaca.Text
    fvValorTasacionVehicular.sNroMotor = txtNroMotor.Text
    fvValorTasacionVehicular.sNroChasis = txtNroChasis.Text
    fvValorTasacionVehicular.nAnioFabricacion = txtAnioFabricacion.Text
    fvValorTasacionVehicular.nAnioAdquisicionTienda = txtAnioAdquisicionTienda.Text
    fnMoneda = CInt(Trim(Right(cmbMoneda.Text, 2)))
    fvValorTasacionVehicular.nValorComercial = txtValorComercial.Text
    fvValorTasacionVehicular.nVRM = CCur(txtVRM.Text)
    If chkTasador.value = 1 Then
        fvValorTasacionVehicular.dTasacion = CDate(txtTasacionFecha.Text)
        fvValorTasacionVehicular.nTasacionTC = CCur(txtTasacionTC.Text)
        fvValorTasacionVehicular.sTasadorCod = txtTasadorCod.psCodigoPersona
        fvValorTasacionVehicular.sTasadorNombre = txtTasadorNombre.Text
        fvValorTasacionVehicular.sTasadorDNI = txtTasadorDNI.Text
        fvValorTasacionVehicular.sTasadorREPEV = txtTasadorREPEV.Text
    End If
    
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
        txtDescripcion.Text = fvValorTasacionVehicular.sDescripcion
        txtCarroceria.Text = fvValorTasacionVehicular.sCarroceria
        txtMarca.Text = fvValorTasacionVehicular.sMarca
        txtModelo.Text = fvValorTasacionVehicular.sModelo
        txtNroPlaca.Text = fvValorTasacionVehicular.sNroPlaca
        txtNroMotor.Text = fvValorTasacionVehicular.sNroMotor
        txtNroChasis.Text = fvValorTasacionVehicular.sNroChasis
        txtAnioFabricacion.Text = fvValorTasacionVehicular.nAnioFabricacion
        txtAnioAdquisicionTienda.Text = fvValorTasacionVehicular.nAnioAdquisicionTienda
        cmbMoneda.ListIndex = IndiceListaCombo(cmbMoneda, fnMoneda)
        txtValorComercial.Text = Format(fvValorTasacionVehicular.nValorComercial, "#,##0.00")
        txtVRM.Text = Format(fvValorTasacionVehicular.nVRM, "#,##0.00")
        If fvValorTasacionVehicular.sTasadorCod <> "" Then
            chkTasador.value = 1
            txtTasacionFecha.Text = Format(fvValorTasacionVehicular.dTasacion, gsFormatoFechaView)
            txtTasacionTC.Text = Format(fvValorTasacionVehicular.nTasacionTC, gsFormatoNumeroView)
            txtTasadorCod.Text = fvValorTasacionVehicular.sTasadorCod
            txtTasadorCod.psCodigoPersona = fvValorTasacionVehicular.sTasadorCod
            txtTasadorNombre.Text = fvValorTasacionVehicular.sTasadorNombre
            txtTasadorDNI.Text = fvValorTasacionVehicular.sTasadorDNI
            txtTasadorREPEV.Text = fvValorTasacionVehicular.sTasadorREPEV
        Else
            chkTasador.value = 0
        End If
                
        If fbConsultar Then
            fraCaracteristicas.Enabled = False
            fraValor.Enabled = False
            fraTasador.Enabled = False
            chkTasador.Enabled = False
            cmdAceptar.Enabled = False
        End If
    End If
    
    If fbPrimero Then
        fnMoneda = gMonedaNacional
    End If
    cmbMoneda.ListIndex = IndiceListaCombo(cmbMoneda, fnMoneda)
    cmbMoneda.Enabled = False
    
    If fbRegistrar Then
        txtDescripcion.Text = fvValorTasacionVehicular_ULT.sDescripcion
        txtCarroceria.Text = fvValorTasacionVehicular_ULT.sCarroceria
        txtMarca.Text = fvValorTasacionVehicular_ULT.sMarca
        txtModelo.Text = fvValorTasacionVehicular_ULT.sModelo
        txtNroPlaca.Text = fvValorTasacionVehicular_ULT.sNroPlaca
        txtNroMotor.Text = fvValorTasacionVehicular_ULT.sNroMotor
        txtNroChasis.Text = fvValorTasacionVehicular_ULT.sNroChasis
        txtAnioFabricacion.Text = IIf(fvValorTasacionVehicular_ULT.nAnioFabricacion = 0, Year(gdFecSis), fvValorTasacionVehicular_ULT.nAnioFabricacion)
        txtAnioAdquisicionTienda.Text = IIf(fvValorTasacionVehicular_ULT.nAnioAdquisicionTienda = 0, Year(gdFecSis), fvValorTasacionVehicular_ULT.nAnioAdquisicionTienda)
        'cmbMoneda.ListIndex = IndiceListaCombo(cmbMoneda, fnMoneda)
        txtValorComercial.Text = Format(fvValorTasacionVehicular_ULT.nValorComercial, "#,##0.00")
        txtVRM.Text = Format(fvValorTasacionVehicular_ULT.nVRM, "#,##0.00")
        If fvValorTasacionVehicular.sTasadorCod <> "" Then
            chkTasador.value = 1
            txtTasacionFecha.Text = Format(IIf(Year(fvValorTasacionVehicular_ULT.dTasacion) <= 1950, gdFecSis, fvValorTasacionVehicular_ULT.dTasacion), gsFormatoFechaView)
            txtTasacionTC.Text = Format(fvValorTasacionVehicular_ULT.nTasacionTC, gsFormatoNumeroView)
            txtTasadorCod.Text = fvValorTasacionVehicular_ULT.sTasadorCod
            txtTasadorCod.psCodigoPersona = fvValorTasacionVehicular_ULT.sTasadorCod
            txtTasadorNombre.Text = fvValorTasacionVehicular_ULT.sTasadorNombre
            txtTasadorDNI.Text = fvValorTasacionVehicular_ULT.sTasadorDNI
            txtTasadorREPEV.Text = fvValorTasacionVehicular_ULT.sTasadorREPEV
        Else
            chkTasador.value = 0
        End If
    End If
    
    If fbRegistrar Then
        Caption = "TASACIÓN VEHICULAR [ NUEVO ]"
    End If
    If fbConsultar Then
        Caption = "TASACIÓN VEHICULAR [ CONSULTAR ]"
    End If
    If fbEditar Then
        Caption = "TASACIÓN VEHICULAR [ EDITAR ]"
    End If
    
    Screen.MousePointer = 0
End Sub
Public Function Registrar(ByVal pbPrimero As Boolean, ByRef pnMoneda As Moneda, ByRef pvValorTasacionVehicular As tValorTasacionVehicular, ByRef pvValorTasacionVehicular_ULT As tValorTasacionVehicular) As Boolean
    fbRegistrar = True
    fbPrimero = pbPrimero
    fnMoneda = pnMoneda
    fvValorTasacionVehicular = pvValorTasacionVehicular
    fvValorTasacionVehicular_ULT = pvValorTasacionVehicular_ULT
    Show 1
    pnMoneda = fnMoneda
    pvValorTasacionVehicular = fvValorTasacionVehicular
    
    Registrar = fbOk
End Function
Public Function Editar(ByVal pbPrimero As Boolean, ByRef pnMoneda As Moneda, ByRef pvValorTasacionVehicular As tValorTasacionVehicular, ByRef pvValorTasacionVehicular_ULT As tValorTasacionVehicular) As Boolean
    fbEditar = True
    fbPrimero = pbPrimero
    fnMoneda = pnMoneda
    fvValorTasacionVehicular = pvValorTasacionVehicular
    fvValorTasacionVehicular_ULT = pvValorTasacionVehicular_ULT
    Show 1
    pnMoneda = fnMoneda
    pvValorTasacionVehicular = fvValorTasacionVehicular
    
    Editar = fbOk
End Function
Public Sub Consultar(ByVal pnMoneda As Moneda, ByRef pvValorTasacionVehicular As tValorTasacionVehicular)
    fbConsultar = True
    fnMoneda = pnMoneda
    fvValorTasacionVehicular = pvValorTasacionVehicular
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
    txtCarroceria.Text = ""
    txtMarca.Text = ""
    txtModelo.Text = ""
    txtNroPlaca.Text = ""
    txtNroMotor.Text = ""
    txtNroChasis.Text = ""
    txtAnioFabricacion.Text = Year(gdFecSis)
    txtAnioAdquisicionTienda.Text = Year(gdFecSis)
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
Private Sub txtAnioAdquisicionTienda_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        If cmbMoneda.Enabled And cmbMoneda.Visible Then
            EnfocaControl cmbMoneda
        Else
            EnfocaControl txtValorComercial
        End If
    End If
End Sub
Private Sub txtAnioAdquisicionTienda_LostFocus()
    txtAnioAdquisicionTienda.Text = Format(txtAnioAdquisicionTienda.Text, "0000")
End Sub
Private Sub txtAnioFabricacion_LostFocus()
    txtAnioFabricacion.Text = Format(txtAnioFabricacion.Text, "0000")
End Sub
Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtCarroceria
    End If
End Sub
Private Sub txtCarroceria_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtMarca
    End If
End Sub
Private Sub txtCarroceria_LostFocus()
    txtCarroceria.Text = UCase(txtCarroceria.Text)
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
        EnfocaControl txtNroPlaca
    End If
End Sub
Private Sub txtModelo_LostFocus()
    txtModelo.Text = UCase(txtModelo.Text)
End Sub
Private Sub txtNroPlaca_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtNroMotor
    End If
End Sub
Private Sub txtNroPlaca_LostFocus()
    txtNroPlaca.Text = UCase(txtNroPlaca.Text)
End Sub
Private Sub txtNroMotor_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtNroChasis
    End If
End Sub
Private Sub txtNroMotor_LostFocus()
    txtNroMotor.Text = UCase(txtNroMotor.Text)
End Sub
Private Sub txtNroChasis_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtAnioFabricacion
    End If
End Sub
Private Sub txtNroChasis_LostFocus()
    txtNroChasis.Text = UCase(txtNroChasis.Text)
End Sub
Private Sub txtAnioFabricacion_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        EnfocaControl txtAnioAdquisicionTienda
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
        EnfocaControl cmdAceptar
    End If
End Sub
Private Sub txtTasadorDNI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl cmdAceptar
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
        EnfocaControl chkTasador
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
    If Len(Trim(txtCarroceria.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la Carrocería", vbInformation, "Aviso"
        EnfocaControl txtCarroceria
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
    If Len(Trim(txtNroPlaca.Text)) = 0 Then
        MsgBox "Ud. debe ingresar el Nro. de Placa", vbInformation, "Aviso"
        EnfocaControl txtNroPlaca
        Exit Function
    End If
    If Len(Trim(txtNroMotor.Text)) = 0 Then
        MsgBox "Ud. debe ingresar el Nro. de Motor", vbInformation, "Aviso"
        EnfocaControl txtNroMotor
        Exit Function
    End If
    If Len(Trim(txtNroChasis.Text)) = 0 Then
        MsgBox "Ud. debe ingresar el Nro. de Chasis", vbInformation, "Aviso"
        EnfocaControl txtNroChasis
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
    If Not IsNumeric(txtAnioAdquisicionTienda.Text) Then
        MsgBox "Ud. debe de especificar el Año de Adquisición en la Tienda", vbInformation, "Aviso"
        EnfocaControl txtAnioAdquisicionTienda
        Exit Function
    Else
        If CInt(txtAnioAdquisicionTienda.Text) <= 1890 Then
            MsgBox "Ud. debe de especificar el Año de Adquisición en la Tienda", vbInformation, "Aviso"
            EnfocaControl txtAnioAdquisicionTienda
            Exit Function
        Else
            If CInt(txtAnioAdquisicionTienda.Text) < CInt(txtAnioFabricacion.Text) Then
                MsgBox "El Año de Adquisición en la Tienda no puede ser menor al año de Fabricación", vbInformation, "Aviso"
                EnfocaControl txtAnioAdquisicionTienda
                Exit Function
            End If
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
    
    If chkTasador.value = 1 Then
        lsFecha = ValidaFecha(txtTasacionFecha.Text)
        If Len(Trim(lsFecha)) > 0 Then
            MsgBox lsFecha, vbInformation, "Aviso"
            EnfocaControl txtTasacionFecha
            Exit Function
        Else
            If Not fvValorTasacionVehicular_ULT.bMigrado Then
                If CDate(txtTasacionFecha.Text) <= fvValorTasacionVehicular_ULT.dTasacion Then
                    MsgBox "La actual fecha de Tasación no puede ser menor o igual a la última Tasación: " & Format(fvValorTasacionVehicular_ULT.dTasacion, gsFormatoFechaView), vbInformation, "Aviso"
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
    End If
        
    validarDatos = True
End Function

