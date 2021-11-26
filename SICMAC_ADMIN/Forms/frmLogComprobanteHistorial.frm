VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogComprobanteHistorial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seguimiento de Comprobantes"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8025
   Icon            =   "frmLogComprobanteHistorial.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   8025
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab TabBuscar 
      Height          =   5070
      Left            =   75
      TabIndex        =   7
      Top             =   75
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   8943
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   6
      TabHeight       =   617
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Búsqueda"
      TabPicture(0)   =   "frmLogComprobanteHistorial.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblPersonaNombre"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtPersonaCod"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "feHistorial"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdSalir"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdLimpiar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cboTpoComprobante"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtComprobanteSerie"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtComprobanteNro"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdBuscar"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   345
         Left            =   6760
         TabIndex        =   4
         Top             =   860
         Width           =   1005
      End
      Begin VB.TextBox txtComprobanteNro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5040
         TabIndex        =   3
         Top             =   885
         Width           =   1665
      End
      Begin VB.TextBox txtComprobanteSerie 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         TabIndex        =   2
         Top             =   885
         Width           =   795
      End
      Begin VB.ComboBox cboTpoComprobante 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   885
         Width           =   2655
      End
      Begin VB.Frame Frame4 
         Caption         =   "Detalle del Comprobante"
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
         Height          =   975
         Left            =   240
         TabIndex        =   8
         Top             =   3960
         Width           =   6285
         Begin VB.Label lblSgteArea 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   4440
            TabIndex        =   23
            Top             =   600
            Width           =   1650
         End
         Begin VB.Label lblUltimoArea 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   4440
            TabIndex        =   22
            Top             =   240
            Width           =   1650
         End
         Begin VB.Label lblSgteMovimiento 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1080
            TabIndex        =   21
            Top             =   600
            Width           =   3330
         End
         Begin VB.Label Label5 
            Caption         =   "Siguiente:"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   600
            Width           =   735
         End
         Begin VB.Label lblUltimoMovimiento 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1080
            TabIndex        =   19
            Top             =   240
            Width           =   3330
         End
         Begin VB.Label Label12 
            Caption         =   "Último:"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   280
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdLimpiar 
         Caption         =   "&Limpiar"
         Height          =   345
         Left            =   6720
         TabIndex        =   5
         Top             =   4150
         Width           =   1005
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   345
         Left            =   6720
         TabIndex        =   6
         Top             =   4510
         Width           =   1005
      End
      Begin Sicmact.FlexEdit feHistorial 
         Height          =   2535
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   7530
         _ExtentX        =   13282
         _ExtentY        =   4471
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Fecha-Movimiento-Estado-Usuario-Área-Glosa"
         EncabezadosAnchos=   "0-1500-3200-1200-800-1500-5000"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-L-C-L-L"
         FormatosEdit    =   "0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin Sicmact.TxtBuscar txtPersonaCod 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   525
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   556
         Appearance      =   0
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
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin VB.Label lblPersonaNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3360
         TabIndex        =   17
         Top             =   525
         Width           =   4410
      End
      Begin VB.Label Label3 
         Caption         =   "Comprobante:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   885
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   525
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "SOLES"
         Height          =   195
         Left            =   -71445
         TabIndex        =   13
         Top             =   3465
         Width           =   525
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL AHORROS"
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
         Left            =   -73185
         TabIndex        =   12
         Top             =   3465
         Width           =   1590
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "DOLARES"
         Height          =   195
         Left            =   -68475
         TabIndex        =   11
         Top             =   3465
         Width           =   765
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -70815
         TabIndex        =   10
         Top             =   3375
         Width           =   2145
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -67680
         TabIndex        =   9
         Top             =   3375
         Width           =   2145
      End
   End
End
Attribute VB_Name = "frmLogComprobanteHistorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
'** Nombre : frmLogComprobanteHistorial
'** Descripción : Historial de Comprobante creado segun ERS062-2013
'** Creación : EJVG, 20131108 09:00:00 AM
'******************************************************************
Option Explicit

Private Sub cmdLimpiar_Click()
    txtPersonaCod.Text = ""
    lblPersonaNombre.Caption = ""
    cboTpoComprobante.ListIndex = -1
    txtComprobanteSerie.Text = ""
    txtComprobanteNro.Text = ""
    FormateaFlex feHistorial
    lblUltimoMovimiento.Caption = ""
    lblUltimoArea.Caption = ""
    lblSgteMovimiento.Caption = ""
    lblSgteArea.Caption = ""
End Sub

Private Sub Form_Load()
    cargaComprobantes
End Sub
Private Sub cboTpoComprobante_KeyPress(KeyAscii As Integer)
    If cboTpoComprobante.ListIndex <> -1 Then
        txtComprobanteSerie.SetFocus
    End If
End Sub
Private Sub cmdBuscar_Click()
    Dim oLog As DLogGeneral
    Dim rs As ADODB.Recordset
    Dim row As Integer
    Dim lsOpeCodUlt As String
    On Error GoTo ErrBuscar
    If Not validaBuscar Then Exit Sub
    
    Set oLog = New DLogGeneral
    Set rs = New ADODB.Recordset
    Set rs = oLog.ListaHistorialComprobante(txtPersonaCod.Text, Left(cboTpoComprobante.Text, 3), Trim(txtComprobanteSerie.Text) & "-" & Trim(txtComprobanteNro.Text))
    FormateaFlex feHistorial
    lblUltimoMovimiento.Caption = ""
    lblUltimoArea.Caption = ""
    lblSgteMovimiento.Caption = ""
    lblSgteArea.Caption = ""
    
    If Not rs.EOF Then
        Do While Not rs.EOF
            feHistorial.AdicionaFila
            row = feHistorial.row
            feHistorial.TextMatrix(row, 1) = Format(rs!dfecha, "dd/mm/yyyy hh:mm:ss AMPM")
            feHistorial.TextMatrix(row, 2) = rs!cOperacion
            feHistorial.TextMatrix(row, 3) = rs!cEstado
            feHistorial.TextMatrix(row, 4) = rs!cUsuario
            feHistorial.TextMatrix(row, 5) = rs!cArea
            feHistorial.TextMatrix(row, 6) = rs!cGlosa
            lblUltimoMovimiento.Caption = rs!cOperacion
            lblUltimoArea.Caption = rs!cArea
            lsOpeCodUlt = rs!cOpeCod
            rs.MoveNext
        Loop
        'Deducimos la sgte. operación
        If Left(lsOpeCodUlt, 2) = "59" Then
            lblSgteMovimiento.Caption = "PROVISIÓN CONTABLE"
            lblSgteArea.Caption = "CONTABILIDAD"
        ElseIf Left(lsOpeCodUlt, 2) = "70" Then
            lblSgteMovimiento.Caption = "PAGO A PROVEEDOR"
            lblSgteArea.Caption = "OPERACIONES"
        End If
    Else
        MsgBox "No se ha encontrado información para el comprobante buscado", vbInformation, "Aviso"
    End If
    Screen.MousePointer = 0
    Set rs = Nothing
    Set oLog = Nothing
    Exit Sub
ErrBuscar:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub txtComprobanteSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtComprobanteNro.SetFocus
    End If
End Sub
Private Sub txtComprobanteNro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdBuscar.SetFocus
    End If
End Sub
Private Sub txtPersonaCod_EmiteDatos()
    lblPersonaNombre.Caption = ""
    If txtPersonaCod.Text <> "" Then
        lblPersonaNombre.Caption = txtPersonaCod.psDescripcion
    End If
End Sub
Private Sub txtPersonaCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboTpoComprobante.SetFocus
    End If
End Sub
Private Sub cargaComprobantes()
    Dim odoc As New DOperacion
    Dim rs As New ADODB.Recordset
    Set rs = odoc.CargaOpeDoc(gnAlmaComprobanteRegistroMN, OpeDocMetDigitado)
    cboTpoComprobante.Clear
    Do While Not rs.EOF
        cboTpoComprobante.AddItem Format(rs!nDocTpo, "00") & " " & Mid(rs!cDocDesc & Space(100), 1, 100) & rs!nDocTpo
        rs.MoveNext
    Loop
    Set rs = Nothing
    Set odoc = Nothing
End Sub
Private Function validaBuscar() As Boolean
    validaBuscar = True
    If Len(Trim(txtPersonaCod.Text)) = 0 Then
        validaBuscar = False
        MsgBox "Ud. debe seleccionar al Proveedor", vbInformation, "Aviso"
        txtPersonaCod.SetFocus
        Exit Function
    End If
    If cboTpoComprobante.ListIndex = -1 Then
        validaBuscar = False
        MsgBox "Ud. debe seleccionar el Tipo de Comprobante", vbInformation, "Aviso"
        cboTpoComprobante.SetFocus
        Exit Function
    End If
    If Len(Trim(txtComprobanteSerie.Text)) = 0 Then
        validaBuscar = False
        MsgBox "Ud. debe especificar el Nro. de Serie de Comprobante", vbInformation, "Aviso"
        txtComprobanteSerie.SetFocus
        Exit Function
    End If
    If Len(Trim(txtComprobanteNro.Text)) = 0 Then
        validaBuscar = False
        MsgBox "Ud. debe especificar el Nro. de Comprobante", vbInformation, "Aviso"
        txtComprobanteNro.SetFocus
        Exit Function
    End If
End Function
