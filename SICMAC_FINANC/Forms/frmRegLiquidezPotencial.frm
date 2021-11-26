VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRegLiquidezPotencial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Liquidez Potencial"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7095
   Icon            =   "frmRegLiquidezPotencial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraReg 
      Caption         =   "Registro"
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   6855
      Begin VB.ComboBox cboRegMan 
         Height          =   315
         ItemData        =   "frmRegLiquidezPotencial.frx":030A
         Left            =   1680
         List            =   "frmRegLiquidezPotencial.frx":0317
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label Label2 
         Caption         =   "Origen de Registro: "
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame fraRegMan 
      Caption         =   "Registro Manual"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   6855
      Begin VB.CheckBox chkLinNac 
         Caption         =   "Línea Nacional"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox txtInstFinanciera 
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   5535
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   5520
         TabIndex        =   7
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdRegistrar 
         Caption         =   "Registrar"
         Height          =   375
         Left            =   4080
         TabIndex        =   6
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtTEA 
         Alignment       =   1  'Right Justify
         Height          =   350
         Left            =   5040
         TabIndex        =   4
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtMontoLinea 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txtGarantia 
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   960
         Width           =   5500
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   330
         Left            =   1200
         TabIndex        =   5
         Top             =   2040
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Sicmact.TxtBuscar txtBuscaIFHisManual 
         Height          =   360
         Left            =   1200
         TabIndex        =   19
         Top             =   360
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   635
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         EnabledText     =   0   'False
      End
      Begin Sicmact.TxtBuscar txtBuscaIF 
         Height          =   360
         Left            =   1200
         TabIndex        =   20
         Top             =   360
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   635
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblDescIF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3120
         TabIndex        =   21
         Top             =   360
         Width           =   3555
      End
      Begin VB.Label lblCod 
         Height          =   255
         Left            =   2760
         TabIndex        =   14
         Top             =   2160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblPorcentaje 
         Caption         =   "%"
         Height          =   255
         Left            =   6000
         TabIndex        =   13
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lblTEA 
         Caption         =   "T.E.A:"
         Height          =   255
         Left            =   4440
         TabIndex        =   12
         Top             =   1665
         Width           =   495
      End
      Begin VB.Label lblFecha 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Monto Linea:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1670
         Width           =   975
      End
      Begin VB.Label lblGarantia 
         Caption         =   "Garantia:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1065
         Width           =   735
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Institución Financiera :"
         Height          =   420
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   900
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmRegLiquidezPotencial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************
'*** Nombre : frmRegLiquidezPotencial
'*** Descripción : Formulario para hacer el Registro de Liquidez Potencial.
'*** Creación : MIOL el 20130305, según OYP-ERS025-2013
'********************************************************************************
Option Explicit
Dim oOpe As DOperacion
Dim oCtaIf As NCajaCtaIF
Dim nMoneda As Integer
Dim nEstado As Integer
Dim objPista As COMManejador.Pista 'ARLO20170217

Private Sub cboRegMan_Click()
    If Me.cboRegMan.Text = "Nuevo Registro Manual" Then
        limpiarCampos
        Me.fraRegMan.Caption = "Nuevo Registro Manual"
        Me.fraRegMan.Visible = True
        Me.txtInstFinanciera.Visible = True
        Me.txtBuscaIFHisManual.Visible = False
        Me.txtBuscaIF.Visible = False
        Me.chkLinNac.Enabled = True
    ElseIf Me.cboRegMan.Text = "Historico de Registro Manual" Then
        limpiarCampos
        Me.fraRegMan.Caption = "Historico de Registro Manual"
        Me.fraRegMan.Visible = True
        Me.txtInstFinanciera.Visible = False
        Me.txtBuscaIF.Visible = False
        Me.txtBuscaIFHisManual.Visible = True
        Me.chkLinNac.Visible = False
    ElseIf Me.cboRegMan.Text = "Institucion Registrada en el Sistema" Then
        limpiarCampos
        Me.fraRegMan.Caption = "Institucion Registrada en el Sistema"
        Me.fraRegMan.Visible = True
        Me.txtInstFinanciera.Visible = False
        Me.txtBuscaIF.Visible = True
        Me.txtBuscaIFHisManual.Visible = False
        Me.chkLinNac.Visible = False
    End If
End Sub

Private Sub cmdRegistrar_Click()
 Dim oDInstFinanc As DInstFinanc
 Set oDInstFinanc = New DInstFinanc
 Dim rsLiqPot As ADODB.Recordset
 Set rsLiqPot = New ADODB.Recordset
 Dim lsMovNro As String
 Dim lsInstFin As String
 
 nMoneda = Mid(gsOpeCod, 3, 1)
    If Me.cboRegMan.Text = "Nuevo Registro Manual" Then
        lsInstFin = Me.txtInstFinanciera
    ElseIf Me.cboRegMan.Text = "Historico de Registro Manual" Then
        lsInstFin = Me.txtBuscaIFHisManual
    ElseIf Me.cboRegMan.Text = "Institucion Registrada en el Sistema" Then
        lsInstFin = Me.lblDescIF
    End If
     
     If nEstado = 1 Then
        If lsInstFin = "" Or txtGarantia.Text = "" Then
           MsgBox "Falta Ingresar la Institución Financiera o Garantia . . ."
           Exit Sub
        End If
     End If
        Set rsLiqPot = oDInstFinanc.validaDatosLiquidezPotencial(txtInstFinanciera, txtMontoLinea, Format(Me.txtFecha.Text, "yyyyMMdd"), txtTEA)
        If rsLiqPot.RecordCount > 0 Then
            MsgBox "Linea Potencial ya Existe . . ."
            If nEstado = 2 Then
                Unload Me
                Exit Sub
            End If
            txtInstFinanciera.Text = ""
            txtGarantia.Text = ""
            txtMontoLinea.Text = Format("0", "#0.00")
            txtTEA.Text = Format("0", "#0.00")
            txtFecha = gdFecSis
            Exit Sub
        End If
    
     If MsgBox(" ¿ Seguro de grabar Datos ? ", vbOKCancel + vbQuestion, "Confirma grabación") = vbOk Then
        lsMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
        If nEstado = 1 Then
           Call oDInstFinanc.registrarLiquidezPotencial(lsMovNro, lsInstFin, txtGarantia, txtMontoLinea, Format(Me.txtFecha.Text, "yyyyMMdd"), txtTEA, nMoneda, Me.chkLinNac.value)
           Call oDInstFinanc.registrarLiquidezPotencialHist(lsMovNro, txtGarantia, txtMontoLinea, Format(Me.txtFecha.Text, "yyyyMMdd"), txtTEA, nEstado, nMoneda, Me.chkLinNac.value, lsMovNro)
        ElseIf nEstado = 2 Then
           Call oDInstFinanc.ModificaLiquidezPotencialxCod(lblCod, txtGarantia, txtMontoLinea, Format(Me.txtFecha.Text, "yyyyMMdd"), txtTEA, Me.chkLinNac.value)
           Call oDInstFinanc.registrarLiquidezPotencialHist(lsMovNro, txtGarantia, txtMontoLinea, Format(Me.txtFecha.Text, "yyyyMMdd"), txtTEA, nEstado, nMoneda, Me.chkLinNac.value, lblCod)
        End If
        
        MsgBox "Los datos se Registraron Correctamente !!!"
        If nEstado = 1 Then
        Me.cboRegMan.ListIndex = -1
        txtInstFinanciera.Text = ""
        Me.lblDescIF = ""
        Me.txtBuscaIF = ""
        txtGarantia = ""
        txtMontoLinea = Format("0", "#0.00")
        txtTEA = Format("0", "#0.00")
        txtFecha = gdFecSis
        Me.chkLinNac.value = 0
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Grabo la Operación "
                Set objPista = Nothing
                '****
        Form_Load
        ElseIf nEstado = 2 Then
            Unload Me
        End If
     End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim gsOpeCodIF As String
Set oOpe = New DOperacion
Set oCtaIf = New NCajaCtaIF
gsOpeCodIF = "401803"
txtBuscaIF.psRaiz = "Instituciones Financieras"
txtBuscaIF.rs = oOpe.GetOpeObj(gsOpeCodIF, "1")

Me.Caption = "Registro Liquidez Potencial " + IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME")
CentraForm Me
txtFecha = gdFecSis
If gsOpeCod = "421401" Or gsOpeCod = "422401" Then
    nEstado = 1
    Me.fraRegMan.Visible = False
    txtMontoLinea = Format("0", "#0.00")
    txtTEA = Format("0", "#0.00")
    Me.chkLinNac.value = 0
ElseIf gsOpeCod = "421402" Or gsOpeCod = "422402" Then
    nEstado = 2
    Me.fraReg.Visible = False
End If
End Sub

Private Sub limpiarCampos()
Me.txtBuscaIF = ""
Me.txtBuscaIFHisManual = ""
Me.txtInstFinanciera = ""
Me.txtFecha = gdFecSis
Me.txtGarantia = ""
txtMontoLinea = Format("0", "#0.00")
txtTEA = Format("0", "#0.00")
Me.chkLinNac.value = 0
End Sub

Private Sub txtBuscaIF_EmiteDatos()
lblDescIF = txtBuscaIF.psDescripcion
End Sub

Private Sub txtBuscaIFHisManual_EmiteDatos()
frmBuscaInstitucion.Show 1
End Sub

Private Sub txtFecha_GotFocus()
fEnfoque txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdRegistrar.SetFocus
    End If
End Sub

Private Sub txtGarantia_GotFocus()
fEnfoque txtMontoLinea
End Sub

Private Sub txtGarantia_KeyPress(KeyAscii As Integer)
KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        txtMontoLinea.SetFocus
    End If
End Sub

Private Sub txtInstFinanciera_GotFocus()
fEnfoque txtGarantia
End Sub

Private Sub txtInstFinanciera_KeyPress(KeyAscii As Integer)
KeyAscii = Letras(KeyAscii)
End Sub

Private Sub txtMontoLinea_GotFocus()
fEnfoque txtTEA
End Sub

Private Sub txtMontoLinea_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMontoLinea, KeyAscii, 14, 5)
    If KeyAscii = 13 Then
        txtTEA.SetFocus
    End If
End Sub

Private Sub txtMontoLinea_LostFocus()
Me.txtMontoLinea.Text = Format(Me.txtMontoLinea.Text, "##,##0.00")
End Sub

Private Sub txtTEA_GotFocus()
fEnfoque txtMontoLinea
End Sub

Private Sub txtTEA_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTEA, KeyAscii, 14, 5)
    If KeyAscii = 13 Then
        txtFecha.SetFocus
    End If
End Sub

Private Sub txtTEA_LostFocus()
Me.txtTEA.Text = Format(Me.txtTEA.Text, "##,##0.00")
End Sub
