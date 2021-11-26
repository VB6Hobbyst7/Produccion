VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmPolizaSeguroPatri 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Polizas de seguros Patrimoniales"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9345
   Icon            =   "frmPolizaSeguroPatri.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   9135
      Begin VB.CommandButton cmdAsignaGasto 
         Caption         =   "Asignar % Gastos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         TabIndex        =   28
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   375
         Left            =   1200
         MaxLength       =   200
         TabIndex        =   8
         Top             =   2400
         Width           =   7815
      End
      Begin VB.Frame Frame2 
         Caption         =   "Vigencia"
         Height          =   735
         Left            =   480
         TabIndex        =   24
         Top             =   1560
         Width           =   3735
         Begin MSMask.MaskEdBox txtDel 
            Height          =   330
            Left            =   720
            TabIndex        =   5
            Top             =   240
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtAl 
            Height          =   330
            Left            =   2400
            TabIndex        =   6
            Top             =   240
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label3 
            Caption         =   "Al :"
            Height          =   255
            Left            =   2040
            TabIndex        =   26
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Del :"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.CommandButton cmdBuscaAseg 
         Caption         =   "..."
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
         Height          =   375
         Left            =   8685
         TabIndex        =   2
         Top             =   795
         Width           =   390
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         Left            =   1125
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1245
         Width           =   4770
      End
      Begin VB.CommandButton cmdexaminar 
         Caption         =   "E&xaminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7785
         TabIndex        =   16
         Top             =   240
         Width           =   1230
      End
      Begin VB.TextBox txtSumaA 
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
         Height          =   345
         Left            =   5325
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   1800
         Width           =   1515
      End
      Begin VB.TextBox txtNumCertif 
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
         Left            =   1095
         TabIndex        =   1
         Top             =   345
         Width           =   2955
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   7245
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1245
         Width           =   1410
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   2460
         Width           =   885
      End
      Begin VB.Label LblAsegPersCod 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   1110
         TabIndex        =   23
         Top             =   795
         Width           =   1350
      End
      Begin VB.Label LblAsegPersNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   2475
         TabIndex        =   22
         Top             =   795
         Width           =   6165
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Aseguradora:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   810
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   615
         TabIndex        =   20
         Top             =   1275
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Prima :"
         Height          =   195
         Left            =   4740
         TabIndex        =   19
         Top             =   1905
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nº Póliza :"
         Height          =   195
         Left            =   285
         TabIndex        =   18
         Top             =   345
         Width           =   735
      End
      Begin VB.Label lblMoneda 
         AutoSize        =   -1  'True
         Caption         =   "Moneda Póliza:"
         Height          =   195
         Left            =   6045
         TabIndex        =   17
         Top             =   1305
         Width           =   1095
      End
   End
   Begin VB.Frame fracontrol 
      Height          =   585
      Left            =   195
      TabIndex        =   0
      Top             =   3000
      Width           =   8940
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
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
         Left            =   1005
         TabIndex        =   14
         Top             =   165
         Width           =   900
      End
      Begin VB.CommandButton cmdsalir 
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
         Left            =   7845
         TabIndex        =   12
         Top             =   165
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eli&minar"
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
         Left            =   1920
         TabIndex        =   11
         Top             =   165
         Width           =   915
      End
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
         Left            =   6870
         TabIndex        =   10
         Top             =   165
         Width           =   975
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
         Left            =   75
         TabIndex        =   9
         Top             =   165
         Width           =   930
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
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
         Left            =   60
         TabIndex        =   13
         Top             =   165
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmPolizaSeguroPatri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nTipoOperacion As Integer '0 Nuevo...1 Modificar
Dim lnTotal As Double
Dim lsMovNro As String

Dim sMatrizDatos() As Variant
Dim nTipoMatriz As Integer
Dim nCantAgeSel As Integer

Sub Limpiar_Controles()
    lnTotal = 0
    txtNumCertif.Text = ""
    LblAsegPersCod.Caption = ""
    LblAsegPersNombre.Caption = ""
    cboTipo.ListIndex = -1
    cboMoneda.ListIndex = -1
    txtSumaA.Text = "0.00"
    Me.txtDel.Text = "__/__/____"
    Me.txtAl.Text = "__/__/____"
    Me.txtDescripcion.Text = ""
    
    
End Sub

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub cboTipo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub cmdAsignaGasto_Click()

Dim i As Integer, j As Integer

    If cboTipo.ListIndex = -1 Then
        MsgBox "Debe indicar el tipo de poliza", vbInformation, "Mensaje"
        cboTipo.SetFocus
        Exit Sub
    End If
    Call frmPoliSeguPatriDistriPorcen.Inicio(CInt(Trim(Right(cboTipo.Text, 20))), sMatrizDatos, nTipoMatriz, nCantAgeSel)
    
'    Dim matriz1() As Variant
'    Dim matriz2(1, 3) As Variant
'
'    For i = 1 To 3
'        For j = 1 To nCantAgeSel
'            matriz2(j, i) = sMatrizDatos(i, j)
'        Next j
'    Next i
    
End Sub

Private Sub cmdBuscaAseg_Click()
Dim oPersona As UPersona
Dim sPersCod As String
    
    Set oPersona = frmBuscaPersona.Inicio
    If Not oPersona Is Nothing Then
        LblAsegPersCod.Caption = oPersona.sPersCod
        LblAsegPersNombre.Caption = oPersona.sPersNombre
    End If
    Set oPersona = Nothing
End Sub



Private Sub cmdBuscaAseg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub cmdCancelar_Click()
    Call Limpiar_Controles
    Call Habilita_Grabar(False)
    Call Habilita_Datos(False)
    Limpiar_Controles
    'cmdCalculaPrima.Enabled = False
    Me.cmdNuevo.Enabled = True
    Me.cmdEditar.Enabled = False
    Me.cmdEliminar.Enabled = False
    
    
End Sub

Private Sub cmdEditar_Click()
nTipoOperacion = 1
Call Habilita_Grabar(True)
Call Habilita_Datos(True)
cmdGrabar.Enabled = True
cmdEditar.Enabled = False
cmdEliminar.Enabled = False
cmdCancelar.Enabled = True
lnTotal = 0
End Sub

Private Sub cmdEliminar_Click()

Dim oPol As DOperacion
Dim oCont As NContFunciones

If MsgBox("Esta seguro que desea eliminar la Poliza?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
    Set oPol = New DOperacion
    Set oCont = New NContFunciones
    lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Call oPol.EliminaPolizaSeguPatri(Trim(txtNumCertif.Text), CInt(Trim(Right(cboTipo.Text, 20))), LblAsegPersCod.Caption, lsMovNro)
    Set oPol = Nothing
    Set oCont = Nothing
    cmdEliminar.Enabled = False
    cmdEditar.Enabled = False
    
    Limpiar_Controles
    
End If
End Sub

Private Sub cmdexaminar_Click()

Dim oPol As DOperacion
Dim rs As ADODB.Recordset
Dim lcDatos As String

Set oPol = New DOperacion

lcDatos = "0"

Call frmPolizaSeguPatriListado.Inicio(0)

Set rs = oPol.CargaDatosPoliSeguPatri(frmPolizaSeguPatriListado.sNumPoliza, frmPolizaSeguPatriListado.nTipoPoliza, frmPolizaSeguPatriListado.sPersCodContr)
If Not rs.EOF Then
    LblAsegPersCod.Caption = rs!cPersCodAseg
    LblAsegPersNombre.Caption = rs!cPersNombre
    cboTipo.ListIndex = IndiceListaCombo(cboTipo, Trim(Str(rs!nTipoSeguro)))
    cboMoneda.ListIndex = IndiceListaCombo(cboMoneda, Trim(Str(rs!nMoneda)))
    txtSumaA.Text = rs!nMontoPrima
    txtNumCertif.Text = rs!cCodPoliza
    Me.txtDel.Text = rs!dVigenciaDel
    Me.txtAl.Text = rs!dVigenciaAl
    Me.txtDescripcion.Text = rs!cDescrip
    
    
    lcDatos = "1"
    
End If
Set oPol = Nothing

If lcDatos = "1" Then
    Me.cmdEditar.Enabled = True
    Me.cmdEliminar.Enabled = True
    Me.cmdNuevo.Enabled = False
    
End If

End Sub

Private Sub CmdGrabar_Click()

Dim oCont As NContFunciones
Dim oPol As DOperacion
Dim i As Integer, j As Integer

If Valida_Datos = False Then Exit Sub

If nCantAgeSel = 0 Then
    MsgBox "Asigne los porcentajes de gastos por favor.", vbOKOnly, "Atención"
    Exit Sub
End If


If MsgBox("¿Esta seguro de registrar la poliza ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Set oPol = New DOperacion
    Set oCont = New NContFunciones
    lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

    If nTipoOperacion = 0 Then
         'Call oPol.InsertaPoliSeguPatri(Trim(txtNumCertif.Text), CInt(Trim(Right(cboTipo.Text, 20))), LblAsegPersCod.Caption, CDbl(txtSumaA.Text), Format(Me.txtDel.Text, "yyyymmdd"), Format(Me.txtAl.Text, "yyyymmdd"), CInt(Trim(Right(cboMoneda.Text, 20))), UCase(Trim(Me.txtDescripcion.Text)), sMatrizDatos, nCantAgeSel, lsMovNro)
         Call oPol.InsertaPoliSeguPatri(Trim(txtNumCertif.Text), CInt(Trim(Right(cboTipo.Text, 20))), LblAsegPersCod.Caption, CDbl(txtSumaA.Text), Format(Me.txtDel.Text, "yyyymmdd"), Format(Me.txtAl.Text, "yyyymmdd"), 2, UCase(Trim(Me.txtDescripcion.Text)), sMatrizDatos, nCantAgeSel, lsMovNro)
    Else
        'Call oPol.ActualizaPoliSeguPatri(Trim(txtNumCertif.Text), CInt(Trim(Right(cboTipo.Text, 20))), LblAsegPersCod.Caption, CDbl(txtSumaA.Text), Format(Me.txtDel.Text, "yyyymmdd"), Format(Me.txtAl.Text, "yyyymmdd"), CInt(Trim(Right(cboMoneda.Text, 20))), UCase(Trim(Me.txtDescripcion.Text)), sMatrizDatos, nCantAgeSel, lsMovNro)
        Call oPol.ActualizaPoliSeguPatri(Trim(txtNumCertif.Text), CInt(Trim(Right(cboTipo.Text, 20))), LblAsegPersCod.Caption, CDbl(txtSumaA.Text), Format(Me.txtDel.Text, "yyyymmdd"), Format(Me.txtAl.Text, "yyyymmdd"), 2, UCase(Trim(Me.txtDescripcion.Text)), sMatrizDatos, nCantAgeSel, lsMovNro)
        
    End If
    Set oPol = Nothing
    Set oCont = Nothing

    cmdGrabar.Enabled = False
    cmdEditar.Enabled = True
    cmdEliminar.Enabled = True
    cmdCancelar.Enabled = False
    Call Habilita_Grabar(False)
    Call Habilita_Datos(False)
    Limpiar_Controles

End Sub

Private Sub cmdImprimir_Click()

End Sub

Private Sub cmdNuevo_Click()
    nTipoOperacion = 0
    Call Limpiar_Controles
    Call Habilita_Grabar(True)
    Call Habilita_Datos(True)
    cmdGrabar.Enabled = True
    cmdEditar.Enabled = False
    cmdEliminar.Enabled = False
    cmdCancelar.Enabled = True
    
    Me.txtNumCertif.SetFocus
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim oCons As COMDConstantes.DCOMConstantes
Dim rs As ADODB.Recordset
Dim RM As ADODB.Recordset


    Dim oGen As DGeneral
    'Set oGen = New DGeneral
    
'    Set rs = oGen.GetConstante(1010)



'Set oCons = New COMDConstantes.DCOMConstantes
    Set oGen = New DGeneral
Set rs = oGen.GetConstante(9077)
Set RM = oGen.GetConstante(1011)
Set oGen = Nothing

Call CentraForm(Me)
Call Llenar_Combo_con_Recordset(rs, cboTipo)
Call Llenar_Combo_con_Recordset(RM, cboMoneda)
Call Habilita_Grabar(False)
Call Habilita_Datos(False)
End Sub

Sub Habilita_Grabar(ByVal pbHabilita As Boolean)
    cmdGrabar.Visible = pbHabilita
    cmdNuevo.Visible = Not pbHabilita
End Sub

Private Sub Option1_Click()
'        LblNroCredito.Visible = False
'        txtNroCredito.Visible = False
'        cmdValidar.Visible = False
End Sub

Private Sub Option2_Click()
'    LblNroCredito.Visible = True
'    txtNroCredito.Visible = True
'    cmdValidar.Visible = True
End Sub

'Private Sub txtMonto_KeyPress(KeyAscii As Integer)
'     KeyAscii = NumerosDecimales(txtMonto, KeyAscii)
'     If KeyAscii = 13 Then
'        cboTipo.SetFocus
'     End If
'End Sub

'Private Sub txtMonto_LostFocus()
'If txtMonto.Text = "" Then
'    txtMonto.Text = "0.00"
'Else
'    txtMonto.Text = Format(txtMonto.Text, "#0.00")
'End If
'End Sub

Function Valida_Datos() As Boolean
Valida_Datos = True
'If CDbl(txtMonto.Text) = 0 Then
'    MsgBox "Debe indicar el valor de la prima", vbInformation, "Mensaje"
'    txtMonto.SetFocus
'    Valida_Datos = False
'    Exit Function
'End If
'If LblContPersCod.Caption = "" Then
'    MsgBox "Debe indicar el contratante", vbInformation, "Mensaje"
'    Valida_Datos = False
'    cmdBuscaCont.SetFocus
'    Exit Function
'End If
If LblAsegPersCod.Caption = "" Then
    MsgBox "Debe indicar la aseguradora", vbInformation, "Mensaje"
    Valida_Datos = False
    cmdBuscaAseg.SetFocus
    Exit Function
End If
If cboTipo.ListIndex = -1 Then
    MsgBox "Debe indicar el tipo de poliza", vbInformation, "Mensaje"
    Valida_Datos = False
    cboTipo.SetFocus
    Exit Function
End If

If cboMoneda.ListIndex = -1 Then
    MsgBox "Debe indicar la moneda de la poliza", vbInformation, "Mensaje"
    Valida_Datos = False
    cboMoneda.SetFocus
    Exit Function
End If


If CDbl(txtSumaA.Text) = 0 Then
    MsgBox "Debe indicar la suma asegurada", vbInformation, "Mensaje"
    Valida_Datos = False
    txtSumaA.SetFocus
End If
End Function

Private Sub Habilita_Datos(ByVal pbHabilita As Boolean)

If nTipoOperacion = 1 Then
    txtSumaA.Enabled = pbHabilita
    cboMoneda.Enabled = pbHabilita
    Me.txtDel.Enabled = pbHabilita
    Me.txtAl.Enabled = pbHabilita
    Me.txtDescripcion.Enabled = pbHabilita
    Me.cmdAsignaGasto.Enabled = pbHabilita
Else
    txtNumCertif.Enabled = pbHabilita
    txtSumaA.Enabled = pbHabilita
    cmdBuscaAseg.Enabled = pbHabilita
    cboTipo.Enabled = pbHabilita
    cboMoneda.Enabled = pbHabilita
    Me.txtDel.Enabled = pbHabilita
    Me.txtAl.Enabled = pbHabilita
    Me.txtDescripcion.Enabled = pbHabilita
    Me.cmdAsignaGasto.Enabled = pbHabilita
End If
    
End Sub

Private Sub txtAl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub txtDel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub txtNumCertif_GotFocus()
    fEnfoque txtNumCertif
End Sub

'peac 20071128 convierte a mayuscula miestras escribe
Private Sub txtNumCertif_Change()
'Dim i As Integer
'    txtNumCertif.Text = UCase(txtNumCertif.Text)
'    i = Len(txtNumCertif.Text)
'    txtNumCertif.SelStart = i
End Sub


Private Sub txtNumCertif_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub txtSumaA_KeyPress(KeyAscii As Integer)
'     KeyAscii = NumerosDecimales(txtSumaA, KeyAscii)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub txtSumaA_LostFocus()
If txtSumaA.Text = "" Then
    txtSumaA.Text = "0.00"
Else
    txtSumaA.Text = Format(txtSumaA.Text, "#0.00")
End If
End Sub
