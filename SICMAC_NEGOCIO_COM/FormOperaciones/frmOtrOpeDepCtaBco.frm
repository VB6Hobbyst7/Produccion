VERSION 5.00
Begin VB.Form frmOtrOpeDepCtaBco 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   Icon            =   "frmOtrOpeDepCtaBco.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMonto 
      Caption         =   "Monto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1035
      Left            =   3960
      TabIndex        =   13
      Top             =   2220
      Width           =   2715
      Begin SICMACT.EditMoney txtMonto 
         Height          =   375
         Left            =   780
         TabIndex        =   4
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.Label lblMon 
         AutoSize        =   -1  'True
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   14
         Top             =   540
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5700
      TabIndex        =   6
      Top             =   3360
      Width           =   1000
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   3360
      Width           =   1000
   End
   Begin VB.Frame fraGlosa 
      Caption         =   "Glosa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1035
      Left            =   60
      TabIndex        =   8
      Top             =   2220
      Width           =   3855
      Begin VB.TextBox txtGlosa 
         Height          =   675
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame fraEntidad 
      Caption         =   "Entidad Financiera"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2115
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   6615
      Begin VB.TextBox txtDocumento 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1140
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1620
         Width           =   2055
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   2655
      End
      Begin SICMACT.TxtBuscar txtEntidad 
         Height          =   375
         Left            =   1140
         TabIndex        =   1
         Top             =   780
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   661
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Documento :"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   1680
         Width           =   915
      End
      Begin VB.Label Label3 
         Caption         =   "&Entidad Financiera :"
         Height          =   435
         Left            =   180
         TabIndex        =   12
         Top             =   780
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Moneda :"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   360
         Width           =   675
      End
      Begin VB.Label lblDescIFCta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1140
         TabIndex        =   10
         Top             =   1200
         Width           =   5355
      End
      Begin VB.Label lblEntidad 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3540
         TabIndex        =   9
         Top             =   780
         Width           =   2955
      End
   End
End
Attribute VB_Name = "frmOtrOpeDepCtaBco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nmoneda As COMDConstantes.Moneda
Dim nOperacion As COMDConstantes.CaptacOperacion

Private Sub ClearScrenn()

cboMoneda.ListIndex = 0
txtEntidad.Text = ""
lblDescIFCta = ""
lblEntidad = ""
txtMonto.Text = "0"
txtGlosa.Text = ""
txtDocumento.Text = ""
cboMoneda.SetFocus

End Sub

Private Sub cboMoneda_Click()
Dim oCtaIf As COMDCajaGeneral.DCOMCajaCtasIF
Dim lsFiltro As String
Dim nMon As Moneda
nMon = CLng(Trim(Right(cboMoneda.Text, 2)))

If nMon <> nmoneda Then
    Set oCtaIf = New COMDCajaGeneral.DCOMCajaCtasIF
    lsFiltro = "_1_[12]%"
    nmoneda = nMon
    txtEntidad.rs = oCtaIf.CargaCtasIF(nmoneda, lsFiltro, MuestraCuentas)
    Set oCtaIf = Nothing
    txtEntidad.Text = ""
    lblDescIFCta = ""
    lblEntidad = ""
End If
If nmoneda = COMDConstantes.gMonedaNacional Then
    txtMonto.BackColor = &HC0FFFF
    lblMon.Caption = "S/."
Else
    txtMonto.BackColor = &HC0FFC0
    lblMon.Caption = "US$"
End If
End Sub

Public Sub Inicia(ByVal nOpe As COMDConstantes.CaptacOperacion, ByVal sCaption As String)

nOperacion = nOpe
Me.Caption = sCaption
Dim oCon As COMDConstantes.DCOMConstantes

'/*Verificar cantidad de operaciones disponibles ANDE 20171218*/
    Dim oCaptaLN As New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim bProsigue As Boolean
    Dim cMsgValid As String
    bProsigue = oCaptaLN.OperacionPermitida(gsCodUser, gdFecSis, CStr(nOperacion), cMsgValid)
    If bProsigue = False Then
        MsgBox cMsgValid, vbOKOnly + vbInformation, "Aviso"
        Exit Sub
    End If
'/*end ande*/


nmoneda = COMDConstantes.gMonedaNacional
Set oCon = New COMDConstantes.DCOMConstantes
    CargaCombo cboMoneda, oCon.RecuperaConstantes(gMoneda)
Set oCon = Nothing
cboMoneda.ListIndex = 0

Dim oCtaIf As COMDCajaGeneral.DCOMCajaCtasIF
Dim lsFiltro As String
Set oCtaIf = New COMDCajaGeneral.DCOMCajaCtasIF

lsFiltro = "_1_[12]%"
txtEntidad.rs = oCtaIf.CargaCtasIF(nmoneda, lsFiltro, MuestraCuentas)
Set oCtaIf = Nothing
txtEntidad.Text = ""
lblDescIFCta = ""
lblEntidad = ""
Me.Show 1
End Sub

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtEntidad.SetFocus
End If
End Sub

Private Sub cmdCancelar_Click()
Unload Me
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

Dim sPersCod As String
Dim sDocumento As String
Dim sGlosa As String
Dim nMonto As Double
Dim lsBoleta As String
Dim nFicSal As String

sPersCod = Trim(txtEntidad.Text)
sDocumento = Trim(txtDocumento.Text)
sGlosa = Trim(txtGlosa.Text)
nMonto = txtMonto.value
If sPersCod = "" Then
    MsgBox "Debe seleccionar la Cuenta de Abono.", vbInformation, "Aviso"
    txtEntidad.SetFocus
    Exit Sub
End If
If sDocumento = "" Then
    MsgBox "Debe digitar el documento de depósito.", vbInformation, "Aviso"
    txtDocumento.SetFocus
    Exit Sub
ElseIf IsNumeric(sDocumento) Then
    If CDbl(sDocumento) = 0 Then
        MsgBox "Debe digitar un documento válido de depósito.", vbInformation, "Aviso"
        txtDocumento.SetFocus
        Exit Sub
    End If
End If

If sGlosa = "" Then
    MsgBox "Debe digitar la glosa del movimiento.", vbInformation, "Aviso"
    txtGlosa.SetFocus
    Exit Sub
End If
If nMonto = 0 Then
    MsgBox "Monto debe ser mayor a cero.", vbInformation, "Aviso"
    txtMonto.SetFocus
    Exit Sub
End If

'EJVG 20120417
Dim loVistoElectronico As frmVistoElectronico
Dim clsMovAct As COMDCaptaGenerales.DCOMCaptaMovimiento
Dim nMovNro As Long
Dim lnMovNroEG As Long 'ALPA20131002
Set loVistoElectronico = New frmVistoElectronico
Set clsMovAct = New COMDCaptaGenerales.DCOMCaptaMovimiento
If nOperacion = gOtrOpeEgresoComisionDepCtaRecaudadora Then
    If Not loVistoElectronico.Inicio(3, gOtrOpeEgresoComisionDepCtaRecaudadora) Then Exit Sub
End If
'END EJVG

If MsgBox("¿Desea grabar la operación?", vbQuestion + vbYesNo, "Grabar") = vbYes Then
    
    Dim clsMov As COMNContabilidad.NCOMContFunciones
    Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
    Dim sMovNro As String
    
    Set clsMov = New COMNContabilidad.NCOMContFunciones
    sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set clsMov = Nothing
    
    Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios
    'ARCV 13-03-2007
    'CLSSERV.OtrasOpeDepCtaBanco sMovNro, sPersCod, sDocumento, nMonto, sGlosa, gsNomAge, nMoneda
    clsServ.OtrasOpeDepCtaBanco sMovNro, sPersCod, sDocumento, nMonto, sGlosa, gsNomAge, nmoneda, nOperacion, lnMovNroEG
    'EJVG20120418
    'ALPA20131001*********************************
    If lnMovNroEG = 0 Then
        MsgBox "La operación no se realizó, favor intente nuevamente", vbInformation, "Aviso"
        Exit Sub
    End If
    '*********************************************

    If nOperacion = gOtrOpeEgresoComisionDepCtaRecaudadora Then
        nMovNro = clsMovAct.GetnMovNro(sMovNro)
        loVistoElectronico.RegistraVistoElectronico (nMovNro)
    End If
    Set clsServ = Nothing
    Dim oBol As COMNCaptaGenerales.NCOMCaptaImpresion
    Set oBol = New COMNCaptaGenerales.NCOMCaptaImpresion
        lsBoleta = oBol.ImprimeBoletaDepCtaBanco(gsNomCmac, gsNomAge, gdFecSis, Trim(lblEntidad.Caption), _
                    Trim(lblDescIFCta.Caption), Trim(txtDocumento.Text), nMonto, sLpt, nmoneda, gsCodUser, "DEP. A CTA. BANCO", False)
    Set oBol = Nothing
    'EJVG20120418
    Set loVistoElectronico = Nothing
    Set clsMovAct = Nothing
    Do
       If Trim(lsBoleta) <> "" Then
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
                Print #nFicSal, lsBoleta
                Print #nFicSal, ""
            Close #nFicSal
       End If
    Loop Until MsgBox("¿Desea reimprimir la boleta?", vbQuestion + vbYesNo, "Aviso") = vbNo
    
    ClearScrenn
    
  '************ Registrar actividad de opertaciones especiales - ANDE 2017-12-18
    Dim oCaptaLN As New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim RVerOpe As ADODB.Recordset
    Dim nEstadoActividad As Integer
    nEstadoActividad = oCaptaLN.RegistrarActividad(CStr(nOperacion), gsCodUser, gdFecSis)
   
    If nEstadoActividad = 1 Then
        MsgBox "He detectado un problema; su operación no fue afectada, pero por favor comunciar a TI-Desarrollo.", vbError, "Error"
    ElseIf nEstadoActividad = 2 Then
        MsgBox "Ha usado el total de operaciones permitidas para el día de hoy. Si desea realizar más operaciones, comuníquese con el área de Operaciones.", vbInformation + vbOKOnly, "Aviso"
        Unload Me
    End If
    ' END ANDE ******************************************************************
    
End If
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
End Sub

Private Sub txtDocumento_GotFocus()
With txtDocumento
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtDocumento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtGlosa.SetFocus
End If
End Sub

Private Sub txtEntidad_EmiteDatos()
Dim oNCtaIF As COMNCajaGeneral.NCOMCajaCtaIF
Dim oCtaIf As COMDCajaGeneral.DCOMCajaCtasIF
Dim sCtaCod As String

If txtEntidad.Text <> "" Then
    sCtaCod = Right(Trim(txtEntidad.Text), 7)
    If InStr(1, sCtaCod, ".", vbTextCompare) > 0 Then
        MsgBox "Cuenta NO Válidad para operación. Seleccionar Cuenta de Ultimo Nivel", vbInformation, "Aviso"
        lblEntidad = ""
        lblDescIFCta = ""
        txtEntidad.Text = ""
        txtEntidad.SetFocus
        Exit Sub
    End If
    Set oCtaIf = New COMDCajaGeneral.DCOMCajaCtasIF
    lblEntidad = oCtaIf.NombreIF(Mid(txtEntidad.Text, 4, 13))
    Set oCtaIf = Nothing
    Set oNCtaIF = New COMNCajaGeneral.NCOMCajaCtaIF
    lblDescIFCta = oNCtaIF.EmiteTipoCuentaIF(Mid(txtEntidad.Text, 18, Len(txtEntidad.Text))) & " " & txtEntidad.psDescripcion
    Set oNCtaIF = Nothing
    txtDocumento.SetFocus
End If
End Sub

Private Sub txtGlosa_GotFocus()
With txtGlosa
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
KeyAscii = fgIntfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    txtMonto.SetFocus
End If
End Sub

Private Sub txtMonto_GotFocus()
txtMonto.MarcaTexto
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdGrabar.SetFocus
End If
End Sub
