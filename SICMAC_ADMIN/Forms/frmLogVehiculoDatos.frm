VERSION 5.00
Begin VB.Form frmLogVehiculoDatos 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3135
   ClientLeft      =   1800
   ClientTop       =   3000
   ClientWidth     =   7890
   Icon            =   "frmLogVehiculoDatos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   7890
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   5340
      TabIndex        =   9
      Top             =   2700
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancela 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6600
      TabIndex        =   8
      Top             =   2700
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   2595
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   7755
      Begin VB.TextBox txtFecha 
         Height          =   315
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtValDesc0 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3240
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   1080
         Visible         =   0   'False
         Width           =   4305
      End
      Begin VB.CommandButton cmdBusq0 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2865
         TabIndex        =   24
         Top             =   1110
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmdBusq2 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2865
         TabIndex        =   15
         Top             =   1830
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmdBusq1 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2865
         TabIndex        =   14
         Top             =   1470
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox txtValDesc2 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3240
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   1800
         Width           =   4305
      End
      Begin VB.TextBox txtValDesc1 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3240
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   1440
         Width           =   4305
      End
      Begin VB.TextBox txtValor2 
         Height          =   315
         Left            =   1200
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   1800
         Width           =   2025
      End
      Begin VB.ComboBox cboSele1 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   6340
      End
      Begin VB.TextBox txtValor1 
         Height          =   315
         Left            =   1200
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1440
         Width           =   2025
      End
      Begin VB.TextBox txtValor0 
         Height          =   315
         Left            =   1200
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   1080
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   315
         Left            =   1200
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1080
         Width           =   6345
      End
      Begin VB.Frame fraDoc 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   240
         TabIndex        =   16
         Top             =   1740
         Width           =   7335
         Begin VB.TextBox txtGlosa 
            Height          =   315
            Left            =   960
            Locked          =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   28
            Top             =   420
            Width           =   6345
         End
         Begin VB.TextBox txtDocNro 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3420
            MaxLength       =   8
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            Top             =   60
            Width           =   705
         End
         Begin VB.ComboBox cboDoc 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   60
            Width           =   2025
         End
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "Buscar"
            Height          =   300
            Left            =   4260
            TabIndex        =   19
            Top             =   60
            Width           =   915
         End
         Begin VB.TextBox txtDocSerie 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3000
            MaxLength       =   3
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Top             =   60
            Width           =   405
         End
         Begin VB.TextBox txtMonto 
            Height          =   300
            Left            =   6000
            Locked          =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   60
            Width           =   1305
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Documento"
            Height          =   195
            Left            =   0
            TabIndex        =   23
            Top             =   120
            Width           =   825
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
            Left            =   5400
            TabIndex        =   22
            Top             =   120
            Width           =   540
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   420
         Width           =   450
      End
      Begin VB.Label lblValor2 
         AutoSize        =   -1  'True
         Caption         =   "Valor2"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1860
         Width           =   450
      End
      Begin VB.Label lblValor1 
         AutoSize        =   -1  'True
         Caption         =   "Valor1"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1500
         Width           =   450
      End
      Begin VB.Label lblSele2 
         AutoSize        =   -1  'True
         Caption         =   "Registro de"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   780
         Width           =   810
      End
      Begin VB.Label lblValor0 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1140
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmLogVehiculoDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vpAcepta As Boolean
Public vpFecha As String
Public vpTipoReg As Integer
Public vpRegistro As String
Public vpDescrip As String
Public vpCodigo0 As String
Public vpCodigo1 As String
Public vpCodigo2 As String
Public vpMonto As Currency

Private Sub Form_Load()
CentraForm Me
Me.Caption = "Control Vehicular - Registro detalle de Incidencias"
txtFecha.Text = Me.vpFecha
Me.vpAcepta = False
Me.vpTipoReg = 0
Me.vpCodigo0 = ""
Me.vpCodigo1 = ""
Me.vpCodigo2 = ""
Me.vpDescrip = ""
Me.vpRegistro = ""
Me.vpMonto = 0
CargaCombos
End Sub

Private Sub cmdAceptar_Click()
Dim nMonto As Currency
'Dim k As Index
Dim i As Integer

Me.vpTipoReg = CInt(cboSele1.ItemData(cboSele1.ListIndex))
Me.vpFecha = txtFecha.Text
Me.vpRegistro = cboSele1.Text
Me.vpDescrip = txtDescripcion.Text
Me.vpCodigo0 = txtValor0.Text
Me.vpCodigo1 = txtValor1.Text
Me.vpCodigo2 = txtValor2.Text
If Len(txtMonto) = 0 Then
   nMonto = 0
Else
   nMonto = CCur(txtMonto)
End If
Me.vpMonto = nMonto
Me.vpAcepta = True
Unload Me
End Sub

Private Sub cmdBusq0_Click()
Dim rsBusq As UPersona
Dim nTipoReg As Integer

nTipoReg = CInt(cboSele1.ItemData(cboSele1.ListIndex))

txtValor0.Text = ""
txtValDesc0.Text = ""
txtDescripcion.Text = ""
   
If nTipoReg = 2 Then
   Set rsBusq = frmBuscaPersona.Inicio
   If rsBusq Is Nothing Then
      Exit Sub
   Else
      txtValor0.Text = rsBusq.sPersCod
      txtValDesc0.Text = rsBusq.sPersNombre
      txtDescripcion.Text = rsBusq.sPersNombre
   End If
   
   'frmBuscaPersona.Show 1
   'If frmBuscaPersona.vpOK Then
   '   txtValor0.Text = frmBuscaPersona.vpPersCod
   '   txtValDesc0.Text = frmBuscaPersona.vpPersNom
   '   txtDescripcion.Text = frmBuscaPersona.vpPersNom
   'End If
End If

If nTipoReg = 3 Then
   frmLogSelector.Consulta "select nConsValor,cConsDescripcion from Constante where nConsCod =9025 and nconscod<>nconsvalor", "Seleccione Incidencia"
   If frmLogSelector.vpHaySeleccion Then
      txtValor0.Text = frmLogSelector.vpCodigo
      txtValDesc0.Text = frmLogSelector.vpDescripcion
      txtDescripcion.Text = frmLogSelector.vpDescripcion
   End If
End If
End Sub

Private Sub cmdBusq1_Click()
Dim nTipoReg As Integer
nTipoReg = CInt(cboSele1.ItemData(cboSele1.ListIndex))
Select Case nTipoReg
    Case 1, 2, 3
         frmSeleUbiGeo.Show 1
         If Len(Trim(frmSeleUbiGeo.vpCodUbigeo)) > 0 Then
            txtValor1.Text = frmSeleUbiGeo.vpCodUbigeo
            txtValDesc1.Text = frmSeleUbiGeo.vpUbigeoDesc
         End If
    Case 4
    
    Case 5
End Select
End Sub

Private Sub cmdBusq2_Click()
Dim nTipoReg As Integer

nTipoReg = CInt(cboSele1.ItemData(cboSele1.ListIndex))
Select Case nTipoReg
    Case 1, 2
         frmSeleUbiGeo.Show 1
         If Len(Trim(frmSeleUbiGeo.vpCodUbigeo)) > 0 Then
            txtValor2.Text = frmSeleUbiGeo.vpCodUbigeo
            txtValDesc2.Text = frmSeleUbiGeo.vpUbigeoDesc
         End If
    Case 3
    
    Case 4
    
    Case 5
End Select
End Sub

Sub CargaCombos()
Dim rs As New ADODB.Recordset
Dim oConn As DConecta, sSQL As String

'------------------------------------------------------
'De acuerdo a la tabla Documento / y por la facilidad
'------------------------------------------------------
'
'1   FACTURA
'2   RECIBO POR HONORARIOS
'3   BOLETA DE VENTA
cboDoc.AddItem "Seleccione documento ---"
cboDoc.AddItem "FACTURA"
cboDoc.ItemData(1) = 1
cboDoc.AddItem "RECIBO POR HONORARIOS"
cboDoc.ItemData(2) = 2
cboDoc.AddItem "BOLETA DE VENTA"
cboDoc.ItemData(3) = 3
cboDoc.AddItem "DOC. INST. PUBLICAS"
cboDoc.ItemData(4) = 19
cboDoc.ListIndex = 0
'------------------------------------------------------

Set oConn = New DConecta
If oConn.AbreConexion Then
   'sSQL = "select nTipoRegCod, cTipoRegistro from LogTipoRegistro where nEstado=1 and nTipoRegCod>0 "
   sSQL = "select nConsValor,cConsDescripcion from Constante where nConsCod =9024 and nConsCod<>nConsValor"
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      Do While Not rs.EOF
         'cboSele1.AddItem Format(rs!nTipoRegCod, "00") + ". " + rs!cTipoRegistro
         cboSele1.AddItem rs!cConsDescripcion
         cboSele1.ItemData(cboSele1.ListCount - 1) = rs!nConsValor
         rs.MoveNext
      Loop
   End If
   Set rs = Nothing
   cboSele1.ListIndex = 0
   oConn.CierraConexion
End If
End Sub

Private Sub cboDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtDocSerie.SetFocus
End If
End Sub

Private Sub cmdBuscar_Click()
Dim rs As New ADODB.Recordset
Dim v As New DLogVehiculos
Dim nDocTpo As Integer
Dim cDocNro As String

nDocTpo = cboDoc.ItemData(cboDoc.ListIndex)
cDocNro = Format(txtDocSerie, "000") + "-" + Format(txtDocNro, "00000000")
If nDocTpo > 0 Then
   Set rs = v.GetMovDocVehiculo(nDocTpo, cDocNro)
   If Not rs.EOF Then
      txtGlosa.Text = UCase(rs!cMovDesc)
      txtDescripcion.Text = txtGlosa.Text
      txtMonto.Text = Format(rs!nMovImporte, "###,##0.00")
      'txtValDesc2.SetFocus
   Else
      txtMonto.Text = ""
      MsgBox "No se halla el documento indicado..." + Space(10), vbInformation
   End If
End If
End Sub

Private Sub cmdCancela_Click()
Me.vpAcepta = False
Me.vpTipoReg = 0
Me.vpRegistro = ""
Me.vpCodigo0 = ""
Me.vpCodigo1 = ""
Me.vpCodigo2 = ""
Me.vpDescrip = ""
Me.vpMonto = 0
Unload Me
End Sub

Private Sub cboSele1_Click()
Dim oConn As New DConecta
Dim rs As New ADODB.Recordset
Dim sSQL As String
Dim nTipoReg As Integer
Dim cCtrl As Control
nTipoReg = CInt(cboSele1.ItemData(cboSele1.ListIndex))
'top 1800
If oConn.AbreConexion Then
   sSQL = "Select * from LogVehiculoCtrl where nTipoReg = " & nTipoReg & " "
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      cmdAceptar.Visible = True
      Do While Not rs.EOF
         If rs!nVisible = 1 Then
            Me.Controls(rs!cCtrlName).Visible = True
            'If rs!nCtrlTop > 0 Then Me.Controls(rs!cCtrlName).Top = rs!nCtrlTop
            'If rs!nCtrlLeft > 0 Then Me.Controls(rs!cCtrlName).Left = rs!nCtrlLeft
            If TypeOf Me.Controls(rs!cCtrlName) Is Label Then
               Me.Controls(rs!cCtrlName).Caption = rs!cCtrlText
            End If
            If TypeOf Me.Controls(rs!cCtrlName) Is TextBox Then
               Me.Controls(rs!cCtrlName).Text = ""
            End If
         Else
            Me.Controls(rs!cCtrlName).Visible = False
         End If
         rs.MoveNext
      Loop
   Else
      txtDescripcion.Visible = False
      lblValor0.Visible = False:   txtValor0.Visible = False:  txtValDesc0.Visible = False
      lblValor1.Visible = False:   txtValor1.Visible = False:  txtValDesc1.Visible = False
      lblValor2.Visible = False:   txtValor2.Visible = False:  txtValDesc2.Visible = False
      cmdBusq0.Visible = False
      cmdBusq1.Visible = False
      cmdBusq2.Visible = False
      fraDoc.Visible = False
      txtFecha.Visible = True
      cboSele1.Visible = True
      cmdAceptar.Visible = False
      MsgBox "Debe indicar que datos se solicitarán para esta incidencia..." + Space(10), vbInformation
      Exit Sub
   End If
End If

If nTipoReg = 1 Or nTipoReg = 2 Then
   txtValor1.Text = "413010101001"
   txtValDesc1.Text = UbigeoDescCompleto("4", "413010101001")
End If
End Sub

Private Sub cboSele1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If txtDescripcion.Visible Then
      txtDescripcion.SetFocus
   End If
   If txtValor0.Visible Then
      txtValor0.SetFocus
   End If
End If
End Sub


Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   If txtValor1.Visible Then
      txtValor1.SetFocus
   End If
End If
End Sub

Private Sub txtDocNro_GotFocus()
txtDocNro.SelStart = 0
txtDocNro.SelLength = txtDocNro.MaxLength
End Sub

Private Sub txtDocNro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtDocNro = Format(txtDocNro, "0000000")
   cmdBuscar.SetFocus
End If
End Sub

Private Sub txtDocSerie_GotFocus()
txtDocSerie.SelStart = 0
txtDocSerie.SelLength = txtDocSerie.MaxLength
End Sub

Private Sub txtDocSerie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtDocSerie = Format(txtDocSerie, "000")
   txtDocNro.SetFocus
End If
End Sub


Private Sub txtFecha_GotFocus()
txtFecha.SelStart = 0
txtFecha.SelLength = Len(Trim(txtFecha))
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigFecha(txtFecha, KeyAscii)
If nKeyAscii = 13 Then
   cboSele1.SetFocus
End If
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdAceptar.SetFocus
End If
End Sub

Private Sub txtValor0_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigNumEnt(KeyAscii)
If nKeyAscii = 13 Then
   If txtValor1.Visible Then
      txtValor1.SetFocus
   End If
End If
End Sub

Private Sub txtValor1_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigNumEnt(KeyAscii)
If nKeyAscii = 13 Then
   If txtValor2.Visible Then
      txtValor2.SetFocus
   End If
   If fraDoc.Visible Then
      cboDoc.SetFocus
   End If
End If
End Sub

Private Sub txtValor2_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigNumEnt(KeyAscii)
If nKeyAscii = 13 Then
   If txtValor2.Visible Then
      cmdAceptar.SetFocus
   End If
End If
End Sub

