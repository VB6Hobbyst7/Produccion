VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLogComprobanteRegistro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Comprobantes"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10395
   Icon            =   "frmLogComprobanteRegistro.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   10395
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab TabBuscar 
      Height          =   6150
      Left            =   45
      TabIndex        =   12
      Top             =   45
      Width           =   10300
      _ExtentX        =   18177
      _ExtentY        =   10848
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
      TabCaption(0)   =   "Datos del Comprobante"
      TabPicture(0)   =   "frmLogComprobanteRegistro.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdExtornar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdDefinir"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdSalir"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   345
         Left            =   9150
         TabIndex        =   11
         Top             =   5640
         Width           =   1005
      End
      Begin VB.CommandButton cmdDefinir 
         Caption         =   "&Definir"
         Height          =   345
         Left            =   8145
         TabIndex        =   10
         Top             =   5640
         Width           =   1005
      End
      Begin VB.CommandButton cmdExtornar 
         Caption         =   "&Extornar"
         Height          =   345
         Left            =   7125
         TabIndex        =   9
         Top             =   5640
         Width           =   1005
      End
      Begin VB.Frame Frame6 
         Caption         =   "Tipo de Pago"
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
         Height          =   660
         Left            =   4080
         TabIndex        =   32
         Top             =   5420
         Width           =   2850
         Begin VB.ComboBox cboTpoPago 
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Acta Referencial"
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
         Height          =   660
         Left            =   165
         TabIndex        =   29
         Top             =   5420
         Width           =   3810
         Begin VB.Label lblMoneda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   800
            TabIndex        =   40
            Top             =   240
            Width           =   1050
         End
         Begin VB.Label lblImporte 
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
            Left            =   2520
            TabIndex        =   39
            Top             =   240
            Width           =   1170
         End
         Begin VB.Label Label13 
            Caption         =   "Importe:"
            Height          =   255
            Left            =   1920
            TabIndex        =   31
            Top             =   270
            Width           =   615
         End
         Begin VB.Label Label12 
            Caption         =   "Moneda:"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   270
            Width           =   615
         End
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
         Height          =   2895
         Left            =   165
         TabIndex        =   27
         Top             =   2480
         Width           =   10005
         Begin Sicmact.FlexEdit feActaConformidad 
            Height          =   2535
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   9685
            _ExtentX        =   17092
            _ExtentY        =   4471
            Cols0           =   9
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Ag.Dest-Objeto-Descripcion-Unidad-Solicitado-P.Unitario-SubTotal-CtaContCod"
            EncabezadosAnchos=   "400-800-1400-2800-950-950-1100-1100-0"
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
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L-L-C-C-R-R-L"
            FormatosEdit    =   "0-0-0-0-0-0-2-2-0"
            CantEntero      =   12
            TextArray0      =   "#"
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            ColWidth0       =   405
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Comprobante"
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
         Height          =   1365
         Left            =   6120
         TabIndex        =   23
         Top             =   1080
         Width           =   4050
         Begin VB.TextBox txtComprobanteSerie 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   840
            TabIndex        =   6
            Top             =   600
            Width           =   795
         End
         Begin VB.TextBox txtComprobanteNro 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            TabIndex        =   7
            Top             =   600
            Width           =   2145
         End
         Begin VB.ComboBox cboTpoComprobante 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker txtComprobanteFecEmision 
            Height          =   285
            Left            =   840
            TabIndex        =   8
            Top             =   915
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   240058369
            CurrentDate     =   41586
         End
         Begin VB.Label Label11 
            Caption         =   "Tipo:"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label6 
            Caption         =   "N°:"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label5 
            Caption         =   "Emisión:"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   960
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Información General"
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
         Height          =   1350
         Left            =   165
         TabIndex        =   14
         Top             =   1080
         Width           =   5850
         Begin VB.Label lblObservacion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1320
            TabIndex        =   38
            Top             =   960
            Width           =   4410
         End
         Begin VB.Label lblAreaAgeNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2520
            TabIndex        =   37
            Top             =   600
            Width           =   3210
         End
         Begin VB.Label lblAreaAgeCod 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1320
            TabIndex        =   36
            Top             =   600
            Width           =   1170
         End
         Begin VB.Label lblProveedorNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2520
            TabIndex        =   35
            Top             =   240
            Width           =   3210
         End
         Begin VB.Label lblProveedorCod 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1320
            TabIndex        =   34
            Top             =   240
            Width           =   1170
         End
         Begin VB.Label Label4 
            Caption         =   "Observaciones:"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Área Usuaria:"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Proveedor:"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Acta Referencial"
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
         Height          =   660
         Left            =   165
         TabIndex        =   13
         Top             =   375
         Width           =   3930
         Begin VB.TextBox txtACAge 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   990
            MaxLength       =   2
            TabIndex        =   1
            Text            =   "00"
            Top             =   240
            Width           =   330
         End
         Begin VB.TextBox txtACNro 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1350
            MaxLength       =   4
            TabIndex        =   2
            Text            =   "0000"
            Top             =   240
            Width           =   555
         End
         Begin VB.CommandButton cmdSeleccionar 
            Caption         =   "&Seleccionar"
            Height          =   345
            Left            =   2580
            TabIndex        =   4
            Top             =   220
            Width           =   1170
         End
         Begin VB.TextBox txtACAnio 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1955
            MaxLength       =   4
            TabIndex        =   3
            Text            =   "0000"
            Top             =   240
            Width           =   570
         End
         Begin VB.ComboBox cboTpoActaConformidad 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   240
            Width           =   855
         End
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
         TabIndex        =   19
         Top             =   3375
         Width           =   2145
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
         TabIndex        =   18
         Top             =   3375
         Width           =   2145
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "DOLARES"
         Height          =   195
         Left            =   -68475
         TabIndex        =   17
         Top             =   3465
         Width           =   765
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
         TabIndex        =   16
         Top             =   3465
         Width           =   1590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "SOLES"
         Height          =   195
         Left            =   -71445
         TabIndex        =   15
         Top             =   3465
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmLogComprobanteRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'** Nombre : frmLogComprobanteRegistro
'** Descripción : Registro Comprobante creado segun ERS062-2013
'** Creación : EJVG, 20131108 09:00:00 AM
'**************************************************************
Option Explicit
Dim fsActaConformidadNro As String
Dim fnMovNroAC As Long, fnMovNroComprob As Long
Dim fnMoneda As Moneda
Dim fdFechaAC As Date
Dim fbProveedorCuentaOK As Boolean
Dim fsProveedorIFICod As String, fsProveedorIFICtaCod As String

Private Sub Form_Load()
    gsopecod = gnAlmaComprobanteRegistroMN
    CargaControles
    limpiarCampos
End Sub
Private Sub cboTpoActaConformidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtACAge.SetFocus
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("¿Desea salir del Registro de Comprobantes?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
        Cancel = 1
    End If
End Sub
Private Sub txtACAge_Change()
    If Len(txtACAge.Text) = 2 Then
        If txtACNro.Enabled And txtACNro.Visible Then
            txtACNro.SetFocus
        End If
    End If
End Sub
Private Sub txtACAge_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = vbKeyBack Then
        Exit Sub
    End If
    If KeyAscii = 13 Then
        txtACNro.SetFocus
    End If
End Sub
Private Sub txtACAge_LostFocus()
    txtACAge.Text = Format(txtACAge.Text, "00")
End Sub
Private Sub txtACAnio_Change()
    If Len(txtACAnio.Text) = 4 Then
        If cmdSeleccionar.Enabled And cmdSeleccionar.Visible Then
            cmdSeleccionar.SetFocus
        End If
    End If
End Sub
Private Sub txtACNro_Change()
    If Len(txtACNro.Text) = 4 Then
        txtACAnio.SetFocus
    End If
End Sub
Private Sub txtACNro_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = vbKeyBack Then
        If Len(txtACNro.Text) = 0 Then
            txtACAge.SetFocus
        End If
    End If
    If KeyAscii = 13 Then
        txtACAnio.SetFocus
    End If
End Sub
Private Sub txtACNro_LostFocus()
    txtACNro.Text = Format(txtACNro.Text, "0000")
End Sub
Private Sub txtACAnio_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = vbKeyBack Then
        If Len(txtACAnio.Text) = 0 Then
            txtACNro.SetFocus
        End If
    End If
    If KeyAscii = 13 Then
        cmdSeleccionar.SetFocus
    End If
End Sub
Private Sub txtACAnio_LostFocus()
    txtACAnio.Text = Format(txtACAnio.Text, "0000")
End Sub
Private Sub cboTpoComprobante_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtComprobanteSerie.SetFocus
    End If
End Sub
Private Sub txtComprobanteFecEmision_LostFocus()
    If CDate(txtComprobanteFecEmision.value) < fdFechaAC Then
        MsgBox "La Fecha de Emisión del Comprobante no puede ser menor a la fecha del Acta de Conformidad", vbInformation, "Aviso"
        txtComprobanteFecEmision.SetFocus
        Exit Sub
    End If
    If CDate(txtComprobanteFecEmision.value) > gdFecSis Then
        MsgBox "La Fecha de Emisión del Comprobante no puede ser mayor a la fecha actual del sistema", vbInformation, "Aviso"
        txtComprobanteFecEmision.SetFocus
        Exit Sub
    End If
End Sub
Private Sub txtComprobanteSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtComprobanteNro.SetFocus
    End If
End Sub
Private Sub txtComprobanteNro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtComprobanteFecEmision.SetFocus
    End If
End Sub
Private Sub txtComprobanteFecEmision_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboTpoPago.SetFocus
    End If
End Sub
Private Sub cmdSeleccionar_Click()
    Dim olog As DLogGeneral
    Dim rs As New ADODB.Recordset
    Dim lnTotal As Currency
    Dim row As Long
    
    On Error GoTo ErrCmdSeleccionar
    If Not ValidaDigiteActaConformidad Then Exit Sub
    fsActaConformidadNro = Left(cboTpoActaConformidad.Text, 3) & "-" & txtACAge.Text & "-" & txtACNro.Text & "-" & txtACAnio.Text
    Set olog = New DLogGeneral
    
    Screen.MousePointer = 11
    Set rs = olog.ListaActaConformidadDetxComprobante(fsActaConformidadNro)
    
    fnMovNroAC = 0
    fnMovNroComprob = 0
    fdFechaAC = CDate("01/01/1900")
    fnMoneda = 0
    fsProveedorIFICod = ""
    fsProveedorIFICtaCod = ""
    lblProveedorCod.Caption = ""
    lblProveedorNombre.Caption = ""
    lblAreaAgeCod.Caption = ""
    lblAreaAgeNombre.Caption = ""
    lblObservacion.Caption = ""
    cboTpoComprobante.ListIndex = -1
    txtComprobanteSerie.Text = ""
    txtComprobanteNro.Text = ""
    txtComprobanteFecEmision.value = Format(gdFecSis, gsFormatoFechaView)
    LimpiaFlex feActaConformidad
    lblMoneda.Caption = ""
    lblImporte.Caption = "0.00"
    cmdDefinir.Enabled = False
    cmdExtornar.Enabled = False
    
    If rs.RecordCount = 0 Then
        Screen.MousePointer = 0
        MsgBox "No se encuentra información del Acta de Conformidad, verifique que el Acta este registrada," & Chr(10) & "que no haya sido extornada o si ya tenia registro de comprobante, éste no se encuentre provisionado", vbInformation, "Aviso"
        Exit Sub
    End If

    fnMovNroAC = rs!nMovNro
    fnMovNroComprob = rs!nMovNroComprob
    fdFechaAC = rs!dDocFecha
    fnMoneda = rs!nMoneda
    fsProveedorIFICod = rs!cIFiCod
    fsProveedorIFICtaCod = rs!cIFiCtaCod
    lblProveedorCod.Caption = rs!cPersCod
    lblProveedorNombre.Caption = rs!cPersNombre
    lblAreaAgeCod.Caption = rs!cAreaAgeCod
    lblAreaAgeNombre.Caption = rs!cAreaAgeDesc
    lblObservacion.Caption = rs!cObservacion
    cboTpoComprobante.ListIndex = IndiceListaCombo(cboTpoComprobante, rs!nDocTpoComprob)
    lblMoneda.Caption = rs!cMoneda
    If rs!nDocTpoComprob > -1 Then
        txtComprobanteSerie.Text = rs!cNroSerieComprob
        txtComprobanteNro.Text = rs!cNroDocComprob
        txtComprobanteFecEmision.value = Format(rs!dDocFechaComprob, gsFormatoFechaView)
    End If
    Do While Not rs.EOF
        feActaConformidad.AdicionaFila
        row = Me.feActaConformidad.row
        feActaConformidad.TextMatrix(row, 1) = rs!cAgeCod
        feActaConformidad.TextMatrix(row, 2) = rs!cObjeto
        feActaConformidad.TextMatrix(row, 3) = rs!cDescripcion
        feActaConformidad.TextMatrix(row, 4) = rs!cUnidad
        feActaConformidad.TextMatrix(row, 5) = IIf(rs!nSolicitado = 0, "", rs!nSolicitado)
        feActaConformidad.TextMatrix(row, 6) = IIf(rs!nPrecioUnitario = 0, "", Format(rs!nPrecioUnitario, gsFormatoNumeroView))
        feActaConformidad.TextMatrix(row, 7) = Format(rs!nSubTotal, gsFormatoNumeroView)
        feActaConformidad.TextMatrix(row, 8) = rs!cCtaContCod
        lnTotal = lnTotal + rs!nSubTotal
        rs.MoveNext
    Loop
    lblImporte.Caption = Format(lnTotal, gsFormatoNumeroView)
    cmdDefinir.Enabled = True
    If fnMovNroComprob > 0 Then
        cmdExtornar.Enabled = True
    End If
    Screen.MousePointer = 0
    
    If Len(fsProveedorIFICod) = 0 Or Len(fsProveedorIFICtaCod) = 0 Then
        MsgBox "Ud. debe verificar que el Proveedor " & UCase(lblProveedorNombre.Caption) & Chr(10) & "se encuentre registrado en la BD de Proveedores de Logística, además tenga" & Chr(10) & "configurado una cuenta en " & UCase(lblMoneda.Caption) & " para poder continuar con el proceso.", vbInformation, "Aviso"
    Else
        '*** Predeterminamos Tipo de Pago
        If fsProveedorIFICod = "1090100012521" Then 'CMACMAYNAS
            cboTpoPago.ListIndex = IndiceListaCombo(cboTpoPago, LogTipoPagoComprobante.gPagoCuentaCMAC)
        Else
            If fsProveedorIFICod = "1090100824640" Then 'BCP
                cboTpoPago.ListIndex = IndiceListaCombo(cboTpoPago, LogTipoPagoComprobante.gPagoTransferencia)
            Else 'OTRO BANCO
                cboTpoPago.ListIndex = IndiceListaCombo(cboTpoPago, LogTipoPagoComprobante.gPagoCheque)
            End If
        End If
        '***
    End If
    
    Set rs = Nothing
    Set olog = Nothing
    Exit Sub
ErrCmdSeleccionar:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub CmdExtornar_Click()
    Dim olog As NLogGeneral
    Dim bExito As Boolean
    If MsgBox("¿Esta seguro de extornar el Comprobante del Acta de Conformidad?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    On Error GoTo ErrExtornar
    Set olog = New NLogGeneral
    bExito = olog.ExtornaComprobante(fnMovNroComprob, , fnMovNroAC)
    If bExito Then
        MsgBox "Se ha extornado satisfactoriamente el Comprobante", vbInformation, "Aviso"
        limpiarCampos
    Else
        MsgBox "Ha ocurrido un error al registrar el comprobante," & Chr(10) & "si el problema persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
    End If
    Set olog = Nothing
    Exit Sub
ErrExtornar:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdDefinir_Click()
    Dim olog As NLogGeneral
    Dim bExito As Boolean
    Dim psOpeCod As String
    
    On Error GoTo ErrDefinir
    If Not validaDefinir Then Exit Sub
    
    If MsgBox("¿Esta seguro de registrar el Comprobante?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Set olog = New NLogGeneral
    
    If fnMoneda = gMonedaNacional Then
        psOpeCod = gnAlmaComprobanteRegistroMN
    Else
        psOpeCod = gnAlmaComprobanteRegistroME
    End If
    
    Screen.MousePointer = 11
    bExito = olog.GrabaComprobante(gdFecSis, gsCodUser, Right(gsCodAge, 2), psOpeCod, "Registro de Comprobante del Acta de Conformidad Nro. " & fsActaConformidadNro, fnMovNroAC, fnMoneda, Trim(Right(cboTpoComprobante.Text, 3)), txtComprobanteSerie.Text & "-" & txtComprobanteNro.Text, CDate(txtComprobanteFecEmision.value), Trim(Right(Me.cboTpoPago.Text, 3)), fsProveedorIFICod, fsProveedorIFICtaCod, fnMovNroComprob)
    Screen.MousePointer = 0
    If bExito Then
        MsgBox "Se ha registrado satisfactoriamente el Comprobante", vbInformation, "Aviso"
        limpiarCampos
    Else
        MsgBox "Ha ocurrido un error al registrar el comprobante," & Chr(10) & "si el problema persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
    End If
    
    Set olog = Nothing
    Exit Sub
ErrDefinir:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub CargaControles()
    Dim olog As New DLogGeneral
    Dim oConst As New DConstante
    Dim rs As New ADODB.Recordset
    
    'Carga Tipo Documentos de Acta Conformidad
    Set rs = olog.ListaDocumentosActaConformidad
    cboTpoActaConformidad.Clear
    Do While Not rs.EOF
        cboTpoActaConformidad.AddItem rs!cDocAbrev
        rs.MoveNext
    Loop
    'Carga Comprobantes
    cargaComprobantes
    'Carga Tipo de Pagos
    Set rs = oConst.RecuperaConstantes(10030)
    cboTpoPago.Clear
    CargaCombo rs, cboTpoPago
    
    Set rs = Nothing
    Set oConst = Nothing
    Set olog = Nothing
End Sub
Private Sub cargaComprobantes()
    Dim oDoc As New DOperacion
    Dim rs As New ADODB.Recordset
    Set rs = oDoc.CargaOpeDoc(gsopecod, OpeDocMetDigitado)
    cboTpoComprobante.Clear
    Do While Not rs.EOF
        cboTpoComprobante.AddItem Format(rs!nDocTpo, "00") & " " & Mid(rs!cDocDesc & Space(100), 1, 100) & rs!nDocTpo
        rs.MoveNext
    Loop
    Set rs = Nothing
    Set oDoc = Nothing
End Sub
Private Sub limpiarCampos()
    'Variables Globales
    fnMovNroAC = 0
    fnMovNroComprob = 0
    fnMoneda = 0
    fdFechaAC = CDate("01/01/1900")
    fsProveedorIFICod = ""
    fsProveedorIFICtaCod = ""
    'Acta Conformidad
    cboTpoActaConformidad.ListIndex = -1
    txtACAge.Text = Right(gsCodAge, 2)
    txtACNro.Text = "0000"
    txtACAnio.Text = Year(gdFecSis)
    'Informacion General
    lblProveedorCod.Caption = ""
    lblProveedorNombre.Caption = ""
    lblAreaAgeCod.Caption = ""
    lblAreaAgeNombre.Caption = ""
    lblObservacion.Caption = ""
    'Comprobante
    cboTpoComprobante.ListIndex = -1
    txtComprobanteSerie.Text = ""
    txtComprobanteNro.Text = ""
    txtComprobanteFecEmision.value = Format(gdFecSis, gsFormatoFechaView)
    'Detalle de Comprobante
    LimpiaFlex feActaConformidad
    'Acta Referencial
    lblMoneda.Caption = ""
    lblImporte.Caption = "0.00"
    'Tipo de Pago
    cboTpoPago.ListIndex = -1
    'Inhabilita botones
    cmdExtornar.Enabled = False
    cmdDefinir.Enabled = False
End Sub
Private Function ValidaDigiteActaConformidad() As Boolean
    ValidaDigiteActaConformidad = True
    If cboTpoActaConformidad.ListIndex = -1 Then
        ValidaDigiteActaConformidad = False
        MsgBox "Ud. debe seleccionar el Tipo de Acta Referencial", vbInformation, "Aviso"
        cboTpoActaConformidad.SetFocus
        Exit Function
    End If
    If Val(txtACAge.Text) = 0 Then
        ValidaDigiteActaConformidad = False
        MsgBox "Ud. debe especificar la agencia del Acta Referencial", vbInformation, "Aviso"
        txtACAge.SetFocus
        Exit Function
    End If
    If Val(txtACNro.Text) = 0 Then
        ValidaDigiteActaConformidad = False
        MsgBox "Ud. debe especificar el Nro del Acta Referencial", vbInformation, "Aviso"
        txtACNro.SetFocus
        Exit Function
    End If
    If Val(txtACAnio.Text) < 1900 Then
        ValidaDigiteActaConformidad = False
        MsgBox "Ud. debe especificar el Año del Acta Referencial", vbInformation, "Aviso"
        txtACAnio.SetFocus
        Exit Function
    End If
End Function
Private Function validaDefinir() As Boolean
    validaDefinir = True
    If Len(fsProveedorIFICod) = 0 Or Len(fsProveedorIFICtaCod) = 0 Then
        MsgBox "Ud. debe verificar que el Proveedor " & UCase(lblProveedorNombre.Caption) & Chr(10) & "se encuentre registrado en la BD de Proveedores de Logística, además tenga" & Chr(10) & "configurado una cuenta en " & UCase(lblMoneda.Caption) & " para poder continuar con el proceso.", vbInformation, "Aviso"
        validaDefinir = False
        Exit Function
    End If
    If cboTpoComprobante.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar el Tipo de Comprobante", vbInformation, "Aviso"
        validaDefinir = False
        cboTpoComprobante.SetFocus
        Exit Function
    End If
    If Len(Trim(txtComprobanteSerie.Text)) = 0 Then
        MsgBox "Ud. debe especificar el Nro de Serie del Comprobante", vbInformation, "Aviso"
        validaDefinir = False
        txtComprobanteSerie.SetFocus
        Exit Function
    End If
    If Len(Trim(txtComprobanteNro.Text)) = 0 Then
        MsgBox "Ud. debe especificar el Nro del Comprobante", vbInformation, "Aviso"
        validaDefinir = False
        txtComprobanteNro.SetFocus
        Exit Function
    End If
    If CCur(lblImporte.Caption) <= 0 Then
        MsgBox "El importe del Comprobante  debe ser mayor a cero", vbInformation, "Aviso"
        validaDefinir = False
        Exit Function
    End If
    If cboTpoPago.ListIndex = -1 Then
        MsgBox "No se ha podido predeterminar la forma de Pago, consulte al Dpto. de TI", vbInformation, "Aviso"
        validaDefinir = False
        cboTpoPago.SetFocus
        Exit Function
    End If
End Function
