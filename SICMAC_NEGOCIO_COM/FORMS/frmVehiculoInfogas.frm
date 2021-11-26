VERSION 5.00
Begin VB.Form frmCredVehiculoInfoGas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Vehículo"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5970
   Icon            =   "frmVehiculoInfogas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   29
      Top             =   6240
      Width           =   5775
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   345
         Left            =   4560
         TabIndex        =   12
         Top             =   240
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   345
         Left            =   3480
         TabIndex        =   11
         Top             =   240
         Width           =   1050
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   345
         Left            =   2400
         TabIndex        =   10
         Top             =   240
         Width           =   1050
      End
   End
   Begin VB.Frame fraInfoVehiculo 
      Caption         =   "Información del Vehículo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   23
      Top             =   3240
      Width           =   5775
      Begin SICMACT.EditMoney txtPorcRecaudo 
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   2520
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.ComboBox cboTaller 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   4215
      End
      Begin VB.TextBox txtPlaca 
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
      End
      Begin VB.ComboBox cboTipoVehiculo 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1080
         Width           =   4215
      End
      Begin VB.ComboBox cboConcesionario 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   4215
      End
      Begin SICMACT.EditMoney txtCuotaInicial 
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   2160
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtMontoAprobado 
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   1800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.Label lblPlacaAnt 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   4440
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Placa Anterior:"
         Height          =   255
         Left            =   3240
         TabIndex        =   36
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Presupuesto del taller"
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   2760
         TabIndex        =   33
         Top             =   1845
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Monto Aprob:"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Taller:"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Placa:"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Cuota Inicial:"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "% Recaudo:"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Tipo Vehiculo:"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Concesionario:"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fraInfoCliente 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   18
      Top             =   480
      Width           =   5775
      Begin VB.TextBox txtFecAprobacion 
         Height          =   315
         Left            =   1440
         TabIndex        =   34
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox txtMontoCredito 
         Height          =   315
         Left            =   1440
         TabIndex        =   17
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtNumDoc 
         Height          =   315
         Left            =   1440
         TabIndex        =   16
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtTipoDoc 
         Height          =   315
         Left            =   1440
         TabIndex        =   15
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtTitular 
         Height          =   315
         Left            =   1440
         TabIndex        =   14
         Top             =   720
         Width           =   4215
      End
      Begin VB.TextBox txtCodCliente 
         Height          =   315
         Left            =   1440
         TabIndex        =   13
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "Fec.Aprobac:"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Monto Credito:"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Documento:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo Doc:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Titular:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Cod.Cliente:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame fraBusca 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      TabIndex        =   37
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton cmdExaminar 
         Caption         =   "..."
         Height          =   375
         Left            =   3720
         TabIndex        =   1
         Top             =   120
         Width           =   495
      End
      Begin SICMACT.ActXCodCta ActXCodCta 
         Height          =   375
         Left            =   0
         TabIndex        =   0
         Top             =   120
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         Texto           =   "N° Cuenta"
         EnabledCta      =   -1  'True
         EnabledAge      =   -1  'True
         Prod            =   "517"
         CMAC            =   "109"
      End
   End
End
Attribute VB_Name = "frmCredVehiculoInfoGas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nTipoOperacion As Integer '1: Registro / 2 Modificación
Dim cPlacaTmp As String

Public Sub Inicio()
    CentraForm Me
    LimpiarControles
    CargarCombos
    Me.Show 1
End Sub
Private Sub cmdGrabar_Click()
    Dim oVehiculo As COMDCredito.DCOMCredActBD
    Dim sCodConcesionario As String
    Dim sCodTaller As String
    Dim nTipoVehiculo As Integer
    Dim sMovNro As String
    Dim clsMovN As COMNContabilidad.NCOMContFunciones
    
    On Error GoTo ErrCmdGrabar
    Set oVehiculo = New COMDCredito.DCOMCredActBD
    If ValidaDatos Then
        If MsgBox("¿Está seguro de haber ingresado correctamente los datos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            Screen.MousePointer = 11
            sCodConcesionario = Trim(Right(cboConcesionario.Text, 13))
            sCodTaller = Trim(Right(cboTaller.Text, 13))
            nTipoVehiculo = CInt(Trim(Right(cboTipoVehiculo.Text, 13)))
            
            Set clsMovN = New COMNContabilidad.NCOMContFunciones
            sMovNro = clsMovN.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
            
            If nTipoOperacion = 1 Then
                Call oVehiculo.AgregaNuevoVehiculoInfoGas(Me.ActXCodCta.NroCuenta, sCodConcesionario, sCodTaller, nTipoVehiculo, txtPlaca.Text, txtPlaca.Text, Format(CCur(txtMontoAprobado.Text), "0.00"), Format(CCur(txtCuotaInicial.Text), "0.00"), Format(CCur(txtPorcRecaudo.Text), "0.00"), 2, gdFecSis, sMovNro, gsCodUser)
            Else
                Call oVehiculo.ModificaVehiculoInfoGas(Me.ActXCodCta.NroCuenta, sCodConcesionario, sCodTaller, nTipoVehiculo, cPlacaTmp, txtPlaca.Text, Format(txtMontoAprobado.Text, "0.00"), Format(txtCuotaInicial.Text, "0.00"), Format(txtPorcRecaudo.Text, "0.00"), sMovNro)
            End If
            MsgBox "Los datos se registraron correctamente", vbInformation, "Aviso"
            'LimpiarControles
            cmdCancelar_Click
            Screen.MousePointer = 0
        Else
            cmdGrabar.SetFocus
        End If
    End If
Exit Sub
ErrCmdGrabar:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    fraInfoCliente.Enabled = False
End Sub
Private Function CargarDatos(sCtaCod As String) As Boolean
    Dim oCred As COMDCredito.DCOMCredito
    Dim oVehiculo As COMDCredito.DCOMCredActBD
    Dim oPers As COMDPersona.DCOMPersonas
    Dim rsCred As ADODB.Recordset
    Dim rsPers As ADODB.Recordset
    Dim rsVehic As ADODB.Recordset
    Set oCred = New COMDCredito.DCOMCredito
    
    
    Set rsCred = oCred.RecuperaDatosCreditoVigente(sCtaCod)
    
    If Not rsCred.EOF Or Not rsCred.BOF Then
        Set oPers = New COMDPersona.DCOMPersonas
        txtCodCliente.Text = rsCred!cPersCod
        txtTitular.Text = rsCred!cPersNombre
        Set rsPers = oPers.ObtenerDatosDocsPers(rsCred!cPersCod)
        txtTipoDoc.Text = IIf(rsPers!cPerDNI <> "", "DNI", "RUC")
        txtNumDoc.Text = IIf(rsPers!cPerDNI <> "", rsPers!cPerDNI, rsPers!cPerRUC)
        txtMontoCredito.Text = Format(rsCred!nMontoCol, "#,###,##0.00")
        txtFecAprobacion.Text = Format(rsCred!dFecAprobacion, "dd/MM/yyyy")
        cmdGrabar.Enabled = True
        fraInfoVehiculo.Enabled = True
        cboConcesionario.SetFocus
        nTipoOperacion = 1 'Define como Registro Nuevo
            
        Set oVehiculo = New COMDCredito.DCOMCredActBD
        Set rsVehic = oVehiculo.ObtieneDatosVehiculoInfoGas(Me.ActXCodCta.NroCuenta)
        If Not rsVehic.EOF And Not rsVehic.BOF Then
            Label12.Visible = True
            lblPlacaAnt.Visible = True
            nTipoOperacion = 2 'Define como Modificación de Registro
            cboConcesionario.ListIndex = IndiceListaCombo(cboConcesionario, rsVehic!cPersCodConcesionario)
            cboTaller.ListIndex = IndiceListaCombo(cboTaller, rsVehic!cPersCodTaller)
            cboTipoVehiculo.ListIndex = IndiceListaCombo(cboTipoVehiculo, rsVehic!nTipoVehiculo)
            txtPlaca.Text = rsVehic!cNuevaPlaca
            lblPlacaAnt.Caption = rsVehic!cPlaca
            txtMontoAprobado.Text = Format(rsVehic!nMontoAprobado, "#,###,##0.00")
            txtCuotaInicial.Text = Format(rsVehic!nCuotaInicial, "#,###,##0.00")
            txtPorcRecaudo.Text = Format(rsVehic!nPorcRecaudo, "0.00")
            cPlacaTmp = rsVehic!cNuevaPlaca
        End If
        fraBusca.Enabled = False
        CargarDatos = True
    Else
        CargarDatos = False
        MsgBox "No se encontraron datos", vbInformation, "Aviso"
    End If
End Function
Private Sub CargarCombos()
    Dim oPers As COMDCredito.DCOMCredActBD
    Dim oCons As COMDConstantes.DCOMConstantes
    Dim rs As ADODB.Recordset
    Dim rsPers As ADODB.Recordset 'ALPA20160226
    
    Set oPers = New COMDCredito.DCOMCredActBD
    Set rs = oPers.ListaPersonasCofideTipo(11) 'Lista Concesionarios
    Call Llenar_Combo_con_Recordset(rs, cboConcesionario)
    
    Set rs = oPers.ListaPersonasCofideTipo(18) 'Lista Talleres
    Call Llenar_Combo_con_Recordset(rs, cboTaller)
    Set rsPers = oPers.ListaPersonasCofideTipo(18)
    
    Set oCons = New COMDConstantes.DCOMConstantes
    Set rs = oCons.RecuperaConstantes(1057) 'Lista Tipos Vehiculos
    Call Llenar_Combo_con_Recordset(rs, cboTipoVehiculo)
    
    Set rs = Nothing
    Set oPers = Nothing
    Set oCons = Nothing
End Sub
Private Function ValidaDatos() As Boolean
    Dim bValidado  As Boolean
    bValidado = True
    If Me.ActXCodCta.NroCuenta = "" And txtCodCliente.Text = "" Then
        MsgBox "Debe ingresar la Cuenta", vbInformation, "Aviso"
        bValidado = False
        cmdExaminar.SetFocus
    ElseIf cboConcesionario.ListIndex = -1 Then
        MsgBox "Debe seleccionar el Concesionario", vbInformation, "Aviso"
        bValidado = False
        cboConcesionario.SetFocus
    ElseIf cboTaller.ListIndex = -1 Then
        MsgBox "Debe seleccionar el Taller", vbInformation, "Aviso"
        bValidado = False
        cboTaller.SetFocus
        'cboTaller.SelStart = 0
        'cboTaller.SelLength = 500
    ElseIf cboTipoVehiculo.ListIndex = -1 Then
        MsgBox "Debe seleccionar el Tipo de Vehiculo", vbInformation, "Aviso"
        bValidado = False
        cboTipoVehiculo.SetFocus
        'cboTipoVehiculo.SelStart = 0
        'cboTipoVehiculo.SelLength = 500
    ElseIf txtPlaca.Text = "" Then
        MsgBox "Debe ingresar la Placa", vbInformation, "Aviso"
        bValidado = False
        txtPlaca.SetFocus
        txtPlaca.SelStart = 0
        txtPlaca.SelLength = 500
    ElseIf txtMontoAprobado.Text = "" Then
        MsgBox "Debe ingresar el Monto Aprobado", vbInformation, "Aviso"
        bValidado = False
        txtMontoAprobado.SetFocus
        txtMontoAprobado.SelStart = 0
        txtMontoAprobado.SelLength = 500
    ElseIf CDbl(txtMontoAprobado.Text) = 0 Then
        MsgBox "El Monto Aprobado no puede ser cero (0)", vbInformation, "Aviso"
        bValidado = False
        txtMontoAprobado.SetFocus
        txtMontoAprobado.SelStart = 0
        txtMontoAprobado.SelLength = 500
    ElseIf txtCuotaInicial.Text = "" Then
        MsgBox "Debe ingresar la Cuota Inicial", vbInformation, "Aviso"
        bValidado = False
        txtCuotaInicial.SetFocus
        txtCuotaInicial.SelStart = 0
        txtCuotaInicial.SelLength = 500
    ElseIf CDbl(txtCuotaInicial.Text) = 0 Then
        MsgBox "La Cuota Inicial no puede ser cero (0)", vbInformation, "Aviso"
        bValidado = False
        txtCuotaInicial.SetFocus
        txtCuotaInicial.SelStart = 0
        txtCuotaInicial.SelLength = 500
    ElseIf txtPorcRecaudo.Text = "" Then
        MsgBox "Debe ingresar el % de Recaudo", vbInformation, "Aviso"
        bValidado = False
        txtPorcRecaudo.SetFocus
        txtPorcRecaudo.SelStart = 0
        txtPorcRecaudo.SelLength = 500
    ElseIf CDbl(txtPorcRecaudo.Text) = 0 Then
        MsgBox "El % de Recaudo no puede ser cero (0)", vbInformation, "Aviso"
        bValidado = False
        txtPorcRecaudo.SetFocus
        txtPorcRecaudo.SelStart = 0
        txtPorcRecaudo.SelLength = 500
    End If
    ValidaDatos = bValidado
End Function
Private Sub LimpiarControles()
    ActXCodCta.Age = ""
    ActXCodCta.Cuenta = ""
    txtCodCliente.Text = ""
    txtTitular.Text = ""
    txtNumDoc.Text = ""
    txtTipoDoc.Text = ""
    txtPlaca.Text = ""
    txtMontoCredito.Text = "0.00"
    txtMontoAprobado.Text = "0.00"
    txtPorcRecaudo.Text = "0.00"
    txtCuotaInicial.Text = "0.00"
    txtFecAprobacion.Text = ""
    cboConcesionario.ListIndex = -1
    cboTaller.ListIndex = -1
    cboTipoVehiculo.ListIndex = -1
    cmdGrabar.Enabled = False
    fraInfoVehiculo.Enabled = False
    Label12.Visible = False
    lblPlacaAnt.Visible = False
End Sub
Private Sub cmdCancelar_Click()
    'Unload Me
    fraBusca.Enabled = True
    LimpiarControles
    ActXCodCta.SetFocusAge
End Sub
Private Sub cmdexaminar_Click()
    ActXCodCta.NroCuenta = frmCredPersEstado.Inicio(Array(gColocEstVigNorm, gColocEstVigVenc, gColocEstVigMor, gColocEstRefNorm, gColocEstRefVenc, gColocEstRefMor), "Creditos Ecotaxi Vigentes", , , , , , , True)
    LimpiarControlesBusca 'EJVG20140901
    If ActXCodCta.NroCuenta <> "" Then
        CargarDatos (ActXCodCta.NroCuenta)
    Else
        MsgBox "No ha ingresado el Número del Crédito", vbInformation, "Aviso"
        Me.ActXCodCta.CMAC = "109"
        Me.ActXCodCta.Prod = "517"
        Me.ActXCodCta.SetFocusAge
    End If
End Sub
Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LimpiarControlesBusca 'EJVG20140901
        If CargarDatos(ActXCodCta.NroCuenta) Then
        cboConcesionario.SetFocus
        End If
    End If
End Sub
Private Sub cboConcesionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboTaller.SetFocus
    End If
End Sub
Private Sub txtCuotaInicial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPorcRecaudo.SetFocus
        txtPorcRecaudo.SelStart = 0
        txtPorcRecaudo.SelLength = 500
    End If
End Sub
Private Sub txtMontoAprobado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCuotaInicial.SetFocus
        txtCuotaInicial.SelStart = 0
        txtCuotaInicial.SelLength = 500
    End If
End Sub
Private Sub txtPlaca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMontoAprobado.SetFocus
        txtMontoAprobado.SelStart = 0
        txtMontoAprobado.SelLength = 500
    End If
    KeyAscii = Asc(UCase(Chr$(KeyAscii)))
    KeyAscii = IIf(KeyAscii = 45, 0, KeyAscii)
End Sub
Private Sub txtPorcRecaudo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And cmdGrabar.Enabled = True Then
        cmdGrabar.SetFocus
    End If
End Sub
Private Sub cboTaller_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboTipoVehiculo.SetFocus
    End If
End Sub
Private Sub cboTipoVehiculo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPlaca.SetFocus
        txtPlaca.SelStart = 0
        txtPlaca.SelLength = 500
    End If
End Sub
Private Sub LimpiarControlesBusca()
    txtCodCliente.Text = ""
    txtTitular.Text = ""
    txtNumDoc.Text = ""
    txtTipoDoc.Text = ""
    txtPlaca.Text = ""
    txtMontoCredito.Text = "0.00"
    txtMontoAprobado.Text = "0.00"
    txtPorcRecaudo.Text = "0.00"
    txtCuotaInicial.Text = "0.00"
    txtFecAprobacion.Text = ""
    cboConcesionario.ListIndex = -1
    cboTaller.ListIndex = -1
    cboTipoVehiculo.ListIndex = -1
    cmdGrabar.Enabled = False
    fraInfoVehiculo.Enabled = False
    Label12.Visible = False
    lblPlacaAnt.Visible = False
End Sub
