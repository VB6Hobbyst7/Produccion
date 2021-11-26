VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLogisticaVehiculoDatos 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3210
   ClientLeft      =   1620
   ClientTop       =   3255
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   8070
   Begin VB.Frame fraCab 
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
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.ComboBox cboSele2 
         Height          =   315
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   3540
      End
      Begin MSMask.MaskEdBox txtFechaIni 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaFin 
         Height          =   315
         Left            =   5580
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFechaFin 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Left            =   4440
         TabIndex        =   10
         Top             =   420
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label lblSele2 
         AutoSize        =   -1  'True
         Caption         =   "Agencia"
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
         Left            =   3120
         TabIndex        =   5
         Top             =   420
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label lblFechaIni 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Left            =   240
         TabIndex        =   2
         Top             =   420
         Width           =   540
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   1020
      Width           =   7815
      Begin VB.CommandButton cmdCancela 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   6300
         TabIndex        =   14
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   5040
         TabIndex        =   13
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtMonto 
         Height          =   300
         Left            =   6300
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   1140
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox txtDesc1 
         Height          =   315
         Left            =   1320
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   420
         Visible         =   0   'False
         Width           =   6165
      End
      Begin VB.TextBox txtDesc2 
         Height          =   315
         Left            =   1320
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   780
         Visible         =   0   'False
         Width           =   6165
      End
      Begin VB.ComboBox cboSele1 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   420
         Visible         =   0   'False
         Width           =   5460
      End
      Begin VB.Label lblDesc1 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
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
         Left            =   180
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label lblMonto 
         AutoSize        =   -1  'True
         Caption         =   "Monto"
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
         Left            =   5280
         TabIndex        =   12
         Top             =   1200
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label lblDesc2 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
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
         Left            =   180
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmLogisticaVehiculoDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vpHaGrabado As Boolean

Dim nTipo As Integer, cBSCod As String, cBSSerie As String

Public Sub Inicio(vTipo As Integer, vBSCod As String, vBSSerie As String)
cBSCod = vBSCod
cBSSerie = vBSSerie
nTipo = vTipo
Me.Show 1
End Sub

Private Sub Form_Load()
Me.vpHaGrabado = False
txtFechaIni.Text = Date
txtFechaFin.Text = Date
Select Case nTipo
    Case 0
         fraCab.Caption = "Asignación de vehículos"
         lblFechaIni.Caption = "Inicio"
         lblFechaFin.Caption = "Término":         lblFechaFin.Visible = True
         lblDesc1.Caption = "Conductor":          lblDesc1.Visible = True
         txtFechaFin.Visible = True
         cboSele1.Visible = True
         CargaComboConductorLibre
    Case 1
         fraCab.Caption = "Registro de Kilometraje"
         lblFechaIni.Caption = "Fecha de registro":  txtFechaIni.Left = 3000
         txtDesc1.Left = 3000:  txtDesc1.Width = 3800:  txtDesc1.Visible = True
         txtDesc2.Left = 3000:  txtDesc2.Width = 3800:  txtDesc2.Visible = True
         lblDesc1.Caption = "Lectura de tacómetro inicial ":  lblDesc1.Visible = True
         lblDesc2.Caption = "Lectura de tacómetro final   ":  lblDesc2.Visible = True
    Case 2
         fraCab.Caption = "Registro de Cargas"
         lblSele2.Caption = "Agencia":      lblSele2.Visible = True
         lblDesc1.Caption = "Descripcion":  lblDesc1.Visible = True
         lblDesc2.Caption = "Destino ":     lblDesc2.Visible = True
         txtDesc1.Visible = True
         txtDesc2.Visible = True
         cboSele2.Visible = True
         CargaAgencias
    Case 3
         fraCab.Caption = "Registro de Incidencias"
         lblSele2.Caption = "Incidencia":      lblSele2.Visible = True
         lblDesc1.Caption = "Descripcion":     lblDesc1.Visible = True
         lblDesc2.Caption = "Lugar Incid.":   lblDesc2.Visible = True
         txtDesc1.Visible = True
         txtDesc2.Visible = True
         cboSele2.Visible = True
         lblMonto.Visible = True
         txtMonto.Visible = True
         CargaComboTipoIncidencia
    Case 4
         fraCab.Caption = "Registro de SOAT"
         lblFechaIni.Caption = "Inicio"
         lblFechaFin.Caption = "Caduca":         lblFechaFin.Visible = True
         txtFechaFin.Visible = True
         lblMonto.Visible = True
         txtMonto.Visible = True
End Select
End Sub

Private Sub cmdCancela_Click()
Me.vpHaGrabado = False
Unload Me
End Sub

Sub CargaAgencias()
Dim LV As DLogVehiculo
Dim rs As ADODB.Recordset
Set LV = New DLogVehiculo
Set rs = LV.GetAgencias
cboSele2.Clear
While Not rs.EOF
    Me.cboSele2.AddItem rs!cAgeCod + ". " + rs!cAgeDescripcion
    rs.MoveNext
Wend
Set rs = Nothing
Set LV = Nothing
End Sub

Sub CargaComboTipoIncidencia()
Dim LV As DLogVehiculo
Dim rs As ADODB.Recordset
Set LV = New DLogVehiculo
Set rs = LV.GetTipoIncidencia

Me.cboSele2.Clear
While Not rs.EOF
    Me.cboSele2.AddItem rs!cDescripcion & Space(100) & rs!ntipoIncidencia
    rs.MoveNext
Wend
Set rs = Nothing
Set LV = Nothing
End Sub

Sub CargaComboConductorLibre()
Dim LV As DLogVehiculo
Dim rs As ADODB.Recordset
Set LV = New DLogVehiculo
Set rs = LV.GetConductorLibre
Me.cboSele1.Clear
While Not rs.EOF
    Me.cboSele1.AddItem rs!cPersNombre & Space(100) & rs!cPersCod
    rs.MoveNext
Wend
Set rs = Nothing
Set LV = Nothing
End Sub

Private Sub cmdGrabar_Click()
Select Case nTipo
    Case 0
         GrabarAsignacion
    Case 1
         GrabarKilometraje
    Case 2
         GrabarCarga
    'Case 3
    '     GrabarPapeleta
    Case 3
         GrabarIncidencia
    Case 4
         GrabarSoat
End Select
End Sub

Private Sub GrabarAsignacion()
Dim LV As DLogVehiculo
Dim opt As Integer
Dim cPersCod As String

'If Right(Me.lblEstado, 1) = 2 Then
'    MsgBox "El vehiculo se encuentra asignado", vbInformation, "AVISO"
'    Exit Sub
'End If

'If Me.chkAsigna.Value = 1 Then
'    If ValFecha(Me.TxtFechaAsigna) = False Then
'        Exit Sub
'    End If
    If Me.cboSele1.ListIndex = -1 Then Exit Sub
'End If

'If Me.chkAsigna.Value = 0 Then
'    MsgBox "Debe marcar el check de Asignacion", vbInformation, "AVISO"
'    Exit Sub
'End If

opt = MsgBox("Está seguro de grabar la asignación" + Space(10), vbQuestion + vbYesNo, "AVISO")
Set LV = New DLogVehiculo
If opt = vbNo Then Exit Sub

cPersCod = Trim(Right(Me.cboSele1.Text, 15))
Call LV.AsignacionVehiculo(cPersCod, cBSCod, cBSSerie, Me.txtFechaIni, txtFechaFin, 1)
Set LV = Nothing
Me.vpHaGrabado = True
Unload Me
End Sub


Private Sub GrabarCarga()
Dim opt As Integer
Dim LV As DLogVehiculo

If cboSele2.ListIndex = -1 Then
    MsgBox "Elija una Agencia", vbInformation, "AVISO"
    Exit Sub
End If

If ValFecha(Me.txtFechaIni) = False Then
    Exit Sub
End If

If Trim(Me.txtDesc1) = "" Then
    MsgBox "Ingrese la Descripcion de la Carga", vbInformation, "AVISO"
    Exit Sub
End If

If Trim(Me.txtDesc2) = "" Then
    MsgBox "Ingrese la Descripcion del Destino", vbInformation, "AVISO"
    Exit Sub
End If
opt = MsgBox("Esta Seguro de Grabar" + Space(10), vbQuestion + vbYesNo, "AVISO")
If opt = vbNo Then Exit Sub

Set LV = New DLogVehiculo
Call LV.InsertRegVehiculoCarga(Me.txtFechaIni, Me.txtDesc1, Left(Me.cboSele2.Text, 2), Me.txtDesc2, cBSCod, cBSSerie, gdFecSis, Right(gsCodAge, 2), gsCodUser)
Set LV = Nothing
Me.vpHaGrabado = True
Unload Me
End Sub

Private Sub GrabarFechaFin()
Dim LV As DLogVehiculo
Dim opt As Integer

If Right(Me.lblEstado, 1) = 1 Then
    MsgBox "El vehiculo se encuentra libre", vbInformation, "AVISO"
    Exit Sub
End If

If ValFecha(Me.TxtFechaFinAsignacion) = False Then
    Exit Sub
End If

opt = MsgBox("Esta Seguaro de Grabar", vbQuestion + vbYesNo, "AVISO")
Set LV = New DLogVehiculo
If opt = vbNo Then Exit Sub

Call LV.LiberaAsignacionVehiculo(Me.lblBSCod, Me.lblBSSerie, Me.TxtFechaFinAsignacion)
lblEstado = "LIBRE" & Space(100) & "1"
CargaAsignacion
CargaComboConductorLibre

Set LV = Nothing
End Sub


Private Sub GrabarIncidencia()
Dim opt As Integer
Dim LV As DLogVehiculo

If ValFecha(Me.txtFechaIni) = False Then
   Exit Sub
End If

If Trim(Me.txtDesc2) = "" Then
    MsgBox "Ingrese el lugar de Incidencia", vbInformation, "AVISO"
    Exit Sub
End If

If Trim(Me.txtDesc1) = "" Then
    MsgBox "Ingrese el Tipo de Incidencia", vbInformation, "AVISO"
    Exit Sub
End If

If Trim(Me.txtMonto) = "" Then
    MsgBox "Ingrese el Monto de la Incidencia", vbInformation, "AVISO"
    Exit Sub
End If

If Me.cboSele2.ListIndex = -1 Then
    MsgBox "Elija el Tipo de Incidencia", vbInformation, "AVISO"
    Exit Sub
End If

opt = MsgBox("¿ Esta Seguro de Grabar ?" + Space(10), vbQuestion + vbYesNo, "AVISO")
If opt = vbNo Then Exit Sub
Set LV = New DLogVehiculo
Call LV.InsertRegVehiculoIncidencia(Me.txtFechaIni, Me.txtDesc1, CInt(Trim(Right(Me.cboSele2, 3))), Me.txtDesc2.Text, cBSCod, cBSSerie, gdFecSis, Right(gsCodAge, 2), gsCodUser)
Set LV = Nothing
Me.vpHaGrabado = True
Unload Me
End Sub

Private Sub GrabarKilometraje()
Dim opt As Integer
Dim LV As DLogVehiculo

If ValFecha(Me.txtFechaIni) = False Then
    Exit Sub
End If

If Me.txtDesc1 = 0 Then
    MsgBox "Dato Incorrecto", vbInformation, "AVISO"
    Me.txtDesc1.SetFocus
    Exit Sub
End If

If Me.txtDesc2 = 0 Then
    MsgBox "Dato Incorrecto", vbInformation, "AVISO"
    Me.txtDesc2.SetFocus
    Exit Sub
End If

If Trim(Me.txtDesc1) = "" Then
    MsgBox "Dato Incorrecto", vbInformation, "AVISO"
    Me.txtDesc1.SetFocus
    Exit Sub
End If

If Trim(Me.txtDesc2) = "" Then
    MsgBox "Dato Incorrecto", vbInformation, "AVISO"
    Me.txtDesc2.SetFocus
    Exit Sub
End If

If txtDesc1 > txtDesc2 Then
    MsgBox "Rango de Kilomentraje Incorrecto", vbInformation, "AVISO"
    Me.txtDesc2.SetFocus
    Exit Sub
End If

opt = MsgBox("Esta Seguro de Grabar", vbQuestion + vbYesNo, "AVISO")

If opt = vbNo Then Exit Sub
Set LV = New DLogVehiculo
Call LV.InsertRegVehiculoKm(Me.txtFechaIni, Me.txtDesc1, Me.txtDesc2, cBSCod, cBSSerie, gdFecSis, Right(gsCodAge, 2), gsCodUser)
Set LV = Nothing
Me.vpHaGrabado = True
Unload Me
End Sub

Private Sub GrabarPapeleta()
Dim opt As Integer
Dim LV As DLogVehiculo

If ValFecha(Me.TxtFechaPapeleta) = False Then
    Exit Sub
End If

If Trim(Me.TxtDescripcionPapeleta) = "" Then
    MsgBox "Ingrese la Descripcion del Destino", vbInformation, "AVISO"
    Exit Sub
End If

If Trim(Me.TxtMontoPapeleta) = "" Or Me.TxtMontoPapeleta = 0 Then
    MsgBox "Ingrese el Monto de la Papeleta", vbInformation, "AVISO"
    Exit Sub
End If

opt = MsgBox("Esta seguro de Grabar", vbQuestion + vbYesNo, "AVISO")
If opt = vbNo Then Exit Sub
Set LV = New DLogVehiculo
Call LV.InsertRegVehiculoPapeleta(Me.TxtFechaPapeleta, Me.TxtDescripcionPapeleta, Me.TxtMontoPapeleta, Me.lblBSCod, Me.lblBSSerie, gdFecSis, Right(gsCodAge, 2), gsCodUser)
Set LV = Nothing
Me.vpHaGrabado = True
Unload Me
End Sub

Private Sub GrabarSoat()
Dim opt As Integer
Dim LV As DLogVehiculo

If ValFecha(Me.txtFechaIni) = False Then
    Exit Sub
End If

If ValFecha(Me.txtFechaFin) = False Then
   Exit Sub
End If

If Me.txtFechaFin < Me.txtFechaIni Then
    MsgBox "Rango de Fechas Incorrectas", vbInformation, "AVISO"
    Exit Sub
End If
opt = MsgBox("Esta Seguro de Grabar", vbQuestion + vbYesNo, "AVISO")
If opt = vbNo Then Exit Sub
Set LV = New DLogVehiculo
Call LV.InsertRegVehiculoSoat(Me.txtFechaIni, Me.txtFechaFin, cBSCod, cBSSerie, txtMonto, gdFecSis, Right(gsCodAge, 2), gsCodUser)
Set LV = Nothing

Me.vpHaGrabado = True
Unload Me
End Sub

Private Sub txtFechaIni_GotFocus()
txtFechaIni.SelStart = 0
txtFechaIni.SelLength = Len(Trim(txtFechaIni))
End Sub

Private Sub txtFechaFin_GotFocus()
txtFechaFin.SelStart = 0
txtFechaFin.SelLength = Len(Trim(txtFechaFin))
End Sub

Private Sub txtFechaIni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If txtFechaFin.Visible Then txtFechaFin.SetFocus
   If cboSele2.Visible Then cboSele2.SetFocus
   If Not cboSele2.Visible And txtDesc1.Visible Then txtDesc1.SetFocus
End If
End Sub

Private Sub txtFechaFin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If txtDesc1.Visible Then txtDesc1.SetFocus
   If cboSele1.Visible Then cboSele1.SetFocus
   If txtMonto.Visible Then txtMonto.SetFocus
End If
End Sub

Private Sub cboSele2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If txtDesc1.Visible Then txtDesc1.SetFocus
   If cboSele1.Visible Then cboSele1.SetFocus
End If
End Sub

Private Sub txtDesc1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If txtDesc2.Visible Then txtDesc2.SetFocus
   If Not txtDesc2.Visible And txtMonto.Visible Then txtMonto.SetFocus
   If Not txtDesc2.Visible And Not txtMonto.Visible Then
      CmdGrabar.SetFocus
   End If
End If
End Sub

Private Sub txtDesc2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If txtMonto.Visible Then txtMonto.SetFocus
   If Not txtMonto.Visible Then
      CmdGrabar.SetFocus
   End If
End If
End Sub

Private Sub cboSele1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If txtDesc1.Visible Then txtDesc1.SetFocus
   If txtDesc2.Visible Then txtDesc2.SetFocus
   If txtMonto.Visible Then txtMonto.SetFocus
   If Not txtDesc2.Visible And Not txtMonto.Visible Then
      CmdGrabar.SetFocus
   End If
End If
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   CmdGrabar.SetFocus
End If
End Sub
