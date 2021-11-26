VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmColEmbargoBien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle del Bien a Embargar"
   ClientHeight    =   5280
   ClientLeft      =   4140
   ClientTop       =   4515
   ClientWidth     =   9990
   ForeColor       =   &H00000000&
   Icon            =   "frmColEmbargoBien.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   9990
   Visible         =   0   'False
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   26
      Top             =   4320
      Width           =   9735
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Bien"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   4215
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   9735
      Begin VB.Frame Frame3 
         Caption         =   "Tasacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2175
         Left            =   6360
         TabIndex        =   30
         Top             =   1920
         Width           =   3255
         Begin VB.ComboBox cmbTasacion 
            Height          =   315
            ItemData        =   "frmColEmbargoBien.frx":030A
            Left            =   1320
            List            =   "frmColEmbargoBien.frx":0314
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   1043
            Width           =   1815
         End
         Begin VB.TextBox txtTasacion 
            Height          =   285
            Left            =   1320
            TabIndex        =   31
            Top             =   240
            Width           =   1815
         End
         Begin MSMask.MaskEdBox mskFechaTasacion 
            Height          =   285
            Left            =   1320
            TabIndex        =   36
            Top             =   645
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label4 
            Caption         =   "Moneda Tasac."
            Height          =   240
            Left            =   120
            TabIndex        =   34
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label lblCapital 
            Caption         =   "Fecha Tasac."
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   660
            Width           =   1215
         End
         Begin VB.Label lblTasacion 
            Caption         =   "Valor Tasac."
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   255
            Width           =   1215
         End
      End
      Begin VB.TextBox txtPartida 
         Height          =   285
         Left            =   4320
         TabIndex        =   29
         Top             =   1080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtPlaca 
         Height          =   285
         Left            =   7440
         TabIndex        =   27
         Top             =   1440
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   2085
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   1950
         Width           =   4935
      End
      Begin VB.TextBox txtMotor 
         Height          =   285
         Left            =   4320
         TabIndex        =   10
         Top             =   1455
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtSerie 
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Top             =   1455
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtModelo 
         Height          =   285
         Left            =   7440
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtMarca 
         Height          =   285
         Left            =   4320
         TabIndex        =   7
         Top             =   1100
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtColor 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   1100
         Width           =   1815
      End
      Begin VB.TextBox txtCantidad 
         Height          =   285
         Left            =   4320
         TabIndex        =   4
         Top             =   735
         Width           =   1815
      End
      Begin VB.CommandButton cmdAlmacen 
         Caption         =   "...."
         Height          =   285
         Left            =   9240
         TabIndex        =   5
         Top             =   720
         Width           =   375
      End
      Begin VB.ComboBox cmbEstado 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton cmdBien 
         Caption         =   "...."
         Height          =   285
         Left            =   9240
         TabIndex        =   2
         Top             =   330
         Width           =   375
      End
      Begin VB.ComboBox cmbSubTipoBien 
         Height          =   315
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   330
         Width           =   1815
      End
      Begin VB.ComboBox cmbTipoBien 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   330
         Width           =   1815
      End
      Begin VB.Label lblPlaca 
         Caption         =   "Placa"
         Height          =   255
         Left            =   6600
         TabIndex        =   28
         Top             =   1455
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1950
         Width           =   855
      End
      Begin VB.Label lblMotor 
         Caption         =   "Motor"
         Height          =   255
         Left            =   3240
         TabIndex        =   24
         Top             =   1470
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblSerie 
         Caption         =   "Serie"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1470
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblModelo 
         Caption         =   "Modelo"
         Height          =   255
         Left            =   6600
         TabIndex        =   22
         Top             =   1110
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblMarcaPartida 
         Caption         =   "Part. Electronic"
         Height          =   255
         Left            =   3240
         TabIndex        =   21
         Top             =   1110
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Color"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1115
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   3240
         TabIndex        =   19
         Top             =   750
         Width           =   735
      End
      Begin VB.Label lblAlmacen 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "[Ingresar Almacen]"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6600
         TabIndex        =   18
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Estado"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   750
         Width           =   735
      End
      Begin VB.Label lblBien 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "[Ingresar Bien]"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6600
         TabIndex        =   16
         Top             =   330
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Sub Tpo Bien"
         Height          =   255
         Left            =   3240
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Tpo Bien"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmColEmbargoBien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Dim lnFila As Integer
'Dim lnOpcion As Boolean
'
'Private fTpoBien As String
'Private fSubTpoBien As String
'Private fBien As String
'Private fEstado As String
'Private fCantidad As String
'Private fAlmacen As String
'Private fColor As String
'Private fMarca As String
'Private fModelo As String
'Private fSerie As String
'Private fMotor As String
'Private fDescripcion As String
'Private fPlaca As String
'Private fPartida As String
'
'Private fTasacion As String
'Private fFecTasacion As String
'Private fMonTasacion As String
'
''PROPIEDADES*************************************************
'Property Let TpoBien(pTpoBien As String)
'   fTpoBien = pTpoBien
'End Property
'Property Get TpoBien() As String
'    TpoBien = fTpoBien
'End Property
'Property Let SubTpoBien(pSubTpoBien As String)
'   fSubTpoBien = pSubTpoBien
'End Property
'Property Get SubTpoBien() As String
'    SubTpoBien = fSubTpoBien
'End Property
'Property Let Bien(pBien As String)
'   fBien = pBien
'End Property
'Property Get Bien() As String
'    Bien = fBien
'End Property
'Property Let Estado(pEstado As String)
'   fEstado = pEstado
'End Property
'Property Get Estado() As String
'    Estado = fEstado
'End Property
'Property Let Cantidad(pCantidad As String)
'   fCantidad = pCantidad
'End Property
'Property Get Cantidad() As String
'    Cantidad = fCantidad
'End Property
'Property Let Almacen(pAlmacen As String)
'   fAlmacen = pAlmacen
'End Property
'Property Get Almacen() As String
'    Almacen = fAlmacen
'End Property
'Property Let Color(pColor As String)
'   fColor = pColor
'End Property
'Property Get Color() As String
'    Color = fColor
'End Property
'Property Let Marca(pMarca As String)
'   fMarca = pMarca
'End Property
'Property Get Marca() As String
'    Marca = fMarca
'End Property
'Property Let Modelo(pModelo As String)
'   fModelo = pModelo
'End Property
'Property Get Modelo() As String
'    Modelo = fModelo
'End Property
'Property Let Serie(pSerie As String)
'   fSerie = pSerie
'End Property
'Property Get Serie() As String
'    Serie = fSerie
'End Property
'Property Let Motor(pMotor As String)
'   fMotor = pMotor
'End Property
'Property Get Motor() As String
'    Motor = fMotor
'End Property
'Property Let Descripcion(pDescripcion As String)
'   fDescripcion = pDescripcion
'End Property
'Property Get Descripcion() As String
'    Descripcion = fDescripcion
'End Property
'Property Let Placa(pPlaca As String)
'   fPlaca = pPlaca
'End Property
'Property Get Placa() As String
'    Placa = fPlaca
'End Property
'Property Let Partida(pPartida As String)
'   fPartida = pPartida
'End Property
'Property Get Partida() As String
'    Partida = fPartida
'End Property
'
'
''---------------------------------
'Property Let Tasacion(pTasacion As String)
'   fTasacion = pTasacion
'End Property
'Property Get Tasacion() As String
'    Tasacion = fTasacion
'End Property
'Property Let FecTasacion(pFecTasacion As String)
'   fFecTasacion = pFecTasacion
'End Property
'Property Get FecTasacion() As String
'    FecTasacion = fFecTasacion
'End Property
'Property Let MonTasacion(pMonTasacion As String)
'   fMonTasacion = pMonTasacion
'End Property
'Property Get MonTasacion() As String
'    MonTasacion = fMonTasacion
'End Property
'
''END PROPIEDADES***************************************
'
'
'Private Sub cmbEstado_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'       Me.txtCantidad.SetFocus
'    End If
'End Sub
'
'Private Sub cmbSubTipoBien_Click()
'    Me.lblBien.Caption = "[Ingresar Bien]" + Space(100) + "0"
'End Sub
'
'Private Sub cmbSubTipoBien_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Me.cmdBien.SetFocus
'    End If
'End Sub
'Private Sub cmbTasacion_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Me.cmdAceptar.SetFocus
'    End If
'End Sub
'
'Private Sub cmbTipoBien_Change()
'    If Right(Me.cmbTipoBien.Text, 4) = "9100" Then
'        HabilitarControles 1, , 1, 1
'    ElseIf Right(Me.cmbTipoBien.Text, 4) = "9200" Then
'        HabilitarControles 1, , 1, 1, 1
'    ElseIf Right(Me.cmbTipoBien.Text, 4) = "9300" Then
'        HabilitarControles 1, , 1, 1, 1, 1
'    ElseIf Right(Me.cmbTipoBien.Text, 4) = "9400" Then
'        HabilitarControles , 1
'    End If
'End Sub
'
'Private Sub cmbTipoBien_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Me.cmbSubTipoBien.SetFocus
'    End If
'End Sub
'
'Private Sub cmdAceptar_Click()
'
'    fTpoBien = cmbTipoBien.Text
'    fSubTpoBien = cmbSubTipoBien.Text
'    fBien = IIf(Trim(Right(Me.lblBien, 4)) = "0", "", Me.lblBien)
'    fEstado = cmbEstado.Text
'    fCantidad = Me.txtCantidad.Text
'    fAlmacen = IIf(Trim(Right(Me.lblAlmacen, 4)) = "0", "", Me.lblAlmacen)
'    fColor = Me.txtColor.Text
'    fMarca = Me.txtMarca.Text
'    fModelo = Me.txtModelo.Text
'    fSerie = Me.txtSerie.Text
'    fMotor = Me.txtMotor.Text
'    fDescripcion = Me.txtDescripcion.Text
'    fPlaca = Me.txtPlaca.Text
'    fPartida = Me.txtPartida.Text
'    '-------------------------------
'    fTasacion = Me.txtTasacion.Text
'    fFecTasacion = IIf(Me.mskFechaTasacion.Text = "__/__/____", "", Me.mskFechaTasacion.Text)
'    'fMonTasacion = IIf(Trim(Right(Me.cmbTasacion, 1)) = "0", "", Me.cmbTasacion)
'    fMonTasacion = IIf(Me.cmbTasacion.ListIndex = -1, "", Me.cmbTasacion + Space(50) + CStr(cmbTasacion.ListIndex + 1))
'    lnOpcion = True
'    Unload Me
'
'
'End Sub
'Public Function Inicio(Optional ByVal pnFila As Integer = 0) As Boolean
'    lnFila = pnFila
'    lnOpcion = False
'    Me.Show 1
'    Inicio = lnOpcion
'End Function
'
'Private Sub cmdBien_Click()
'   If Me.cmbSubTipoBien.ListIndex <> -1 Then
'        Me.lblBien = frmColEmbargoBienListar.Inicio(Right(Me.cmbSubTipoBien.Text, 4), Trim(Left(Me.cmbSubTipoBien.Text, 30)))
'        If Me.lblBien <> "" Then
'            cmbEstado.SetFocus
'        End If
'   Else
'        MsgBox "Selecciones el Sub Tipo de Bien"
'   End If
'End Sub
'
'Private Sub cmdBien_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        cmdBien_Click
'    End If
'End Sub
'
'Private Sub Form_Load()
'
'    Me.lblBien = "[Ingrese Bien]" + Space(100) + "0"
'    Me.lblAlmacen = "[Ingrese Almacen]" + Space(100) + "0"
'    cargarControles
'    If lnFila <> 0 Then
'        Me.cmbTipoBien.ListIndex = IndiceListaCombo(cmbTipoBien, Trim(Right(Me.TpoBien, 5)))
'        Me.cmbSubTipoBien.ListIndex = IndiceListaCombo(cmbSubTipoBien, Trim(Right(Me.SubTpoBien, 5)))
'        Me.lblBien = Me.Bien
'        Me.cmbEstado.ListIndex = IndiceListaCombo(cmbEstado, Trim(Right(Me.Estado, 5)))
'        Me.txtCantidad.Text = Me.Cantidad
'        Me.lblAlmacen = Me.Almacen
'        Me.txtColor.Text = Me.Color
'        Me.txtMarca.Text = Me.Marca
'        Me.txtModelo = Me.Modelo
'        Me.txtSerie.Text = Me.Serie
'        Me.txtMotor.Text = Me.Motor
'        Me.txtDescripcion.Text = Me.Descripcion
'        Me.txtPlaca.Text = Me.Placa
'        Me.txtPartida.Text = Me.Partida
'        '------------------------------------------
'        Me.txtTasacion.Text = Me.Tasacion
'        Me.mskFechaTasacion.Text = IIf(Me.FecTasacion = "", "__/__/____", Me.FecTasacion)
'        Me.cmbTasacion.ListIndex = Me.MonTasacion
'        'Me.cmbTasacion.ListIndex = IndiceListaCombo(cmbTasacion, Trim(Right(Me.MonTasacion, 1)))
'    End If
'End Sub
'Private Sub cmbTipoBien_Click()
'    Dim rsCombo As Recordset
'    Dim oColRec As COMNColocRec.NCOMColRecCredito
'     Me.lblBien.Caption = "[Ingresar Bien]" + Space(100) + "0"
'    Set oColRec = New COMNColocRec.NCOMColRecCredito
'    Set rsCombo = oColRec.ObtenerConsValorEmbargo(9987, Left(Right(Me.cmbTipoBien.Text, 4), 2), "_[1-9]")
'    If Not (rsCombo.BOF And rsCombo.EOF) Then
'       Llenar_Combo_con_Recordset rsCombo, Me.cmbSubTipoBien
'       Set rsCombo = Nothing
'    End If
'    If Right(Me.cmbTipoBien.Text, 4) = "9100" Then
'        HabilitarControles 1, , 1, 1
'    ElseIf Right(Me.cmbTipoBien.Text, 4) = "9200" Then
'        HabilitarControles 1, , 1, 1, 1
'    ElseIf Right(Me.cmbTipoBien.Text, 4) = "9300" Then
'        HabilitarControles 1, , 1, 1, 1, 1
'    ElseIf Right(Me.cmbTipoBien.Text, 4) = "9400" Then
'        HabilitarControles , 1
'    End If
'End Sub
'Private Sub cmdAlmacen_Click()
'    lblAlmacen = frmColEmbargoAlmacen.Inicio(Right(lblAlmacen, 5))
'    If lblAlmacen <> "" Then
'       Me.txtColor.SetFocus
'    End If
'End Sub
'Private Sub cargarControles()
'    Dim rsCombo As Recordset
'    Dim oColRec As COMNColocRec.NCOMColRecCredito
'
'    Set oColRec = New COMNColocRec.NCOMColRecCredito
'    Set rsCombo = oColRec.ObtenerConsValorEmbargo(9988, "%", "%")
'
'    If Not (rsCombo.BOF And rsCombo.EOF) Then
'       Llenar_Combo_con_Recordset rsCombo, Me.cmbEstado
'       Set rsCombo = Nothing
'    End If
'
'    Set oColRec = New COMNColocRec.NCOMColRecCredito
'    Set rsCombo = oColRec.ObtenerConsValorEmbargo(9987, "%", "00")
'    If Not (rsCombo.BOF And rsCombo.EOF) Then
'       Llenar_Combo_con_Recordset rsCombo, Me.cmbTipoBien
'       Set rsCombo = Nothing
'    End If
'
''    Set oColRec = New COMNColocRec.NCOMColRecCredito
''    With rsCombo
''        .Fields.Append "cDescripcion", adVarChar, 15
''        .Fields.Append "nMoneda", adInteger
''        .Open
''        .AddNew
''        .Fields("cDescripcion") = "Nacional"
''        .Fields("nMoneda") = 1
''        .AddNew
''        .Fields("cDescripcion") = "Extranjera"
''        .Fields("nMoneda") = 2
''        .MoveFirst
''        Llenar_Combo_con_Recordset rsCombo, Me.cmbTasacion
''        Set rsCombo = Nothing
''    End With
'
'
'End Sub
'Private Sub mskFechaTasacion_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Me.cmbTasacion.SetFocus
'    End If
'End Sub
'
'Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Me.cmdAlmacen.SetFocus
'    End If
'End Sub
'Private Sub txtColor_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If Me.txtMarca.Visible = True Then
'            Me.txtMarca.SetFocus
'        ElseIf Me.txtPartida.Visible = True Then
'            Me.txtPartida.SetFocus
'        End If
'    End If
'End Sub
'Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        'Me.txtDescripcion.Text = Left(Me.txtDescripcion.Text, Len(Me.txtDescripcion.Text) - 2)
'        'Me.cmdAceptar.SetFocus
'        KeyAscii = 0
'        Me.txtDescripcion = Me.txtDescripcion.Text + Chr(KeyAscii)
'        Me.txtTasacion.SetFocus
'    End If
'End Sub
'
'Private Sub txtMarca_Change()
'    Dim i As Integer
'    txtMarca.Text = UCase(txtMarca.Text)
'    i = Len(txtMarca.Text)
'    txtMarca.SelStart = i
'End Sub
'
'Private Sub txtPlaca_Change()
'    Dim i As Integer
'    txtPlaca.Text = UCase(txtPlaca.Text)
'    i = Len(txtPlaca.Text)
'    txtPlaca.SelStart = i
'End Sub
'Private Sub txtMarca_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Me.txtModelo.SetFocus
'    End If
'End Sub
'Private Sub txtModelo_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Me.txtSerie.SetFocus
'    End If
'End Sub
'Private Sub txtMotor_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If Me.txtPlaca.Visible = True Then
'            Me.txtPlaca.SetFocus
'        Else
'            Me.txtDescripcion.SetFocus
'        End If
'    End If
'End Sub
'Private Sub txtPartida_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Me.txtDescripcion.SetFocus
'    End If
'End Sub
'Private Sub txtPlaca_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Me.txtDescripcion.SetFocus
'    End If
'End Sub
'
'Private Sub txtSerie_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If txtMotor.Visible = True Then
'            Me.txtMotor.SetFocus
'        Else
'            Me.txtDescripcion.SetFocus
'        End If
'    End If
'End Sub
'Private Sub txtTasacion_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Me.mskFechaTasacion.SetFocus
'    End If
'End Sub
'Private Sub HabilitarControles(Optional pnMarca As Integer = 0, Optional pnPartida As Integer = 0, Optional pnModelo As Integer = 0, _
'                               Optional pnSerie As Integer = 0, Optional pnMotor As Integer = 0, Optional pnPlaca As Integer = 0)
'
'    Me.txtMarca.Visible = pnMarca
'    If pnMarca = 1 Then
'        Me.lblMarcaPartida = "Marca"
'        Me.lblMarcaPartida.Visible = pnMarca
'    End If
'
'    Me.txtPartida.Visible = pnPartida
'    If pnPartida = 1 Then
'        Me.lblMarcaPartida = "Part. Electronic"
'        Me.lblMarcaPartida.Visible = pnPartida
'    End If
'
'    Me.txtModelo.Visible = pnModelo
'    Me.lblModelo.Visible = pnModelo
'
'    Me.txtSerie.Visible = pnSerie
'    Me.lblSerie.Visible = pnSerie
'
'    Me.txtMotor.Visible = pnMotor
'    Me.lblMotor.Visible = pnMotor
'
'    Me.txtPlaca.Visible = pnPlaca
'    Me.lblPlaca.Visible = pnPlaca
'
'
'End Sub
'
'
