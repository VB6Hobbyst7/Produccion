VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProveedorRegSistemaPension 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Datos AFP/ONP"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6585
   Icon            =   "frmProveedorRegSistemaPension.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   340
      Left            =   5400
      TabIndex        =   23
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   340
      Left            =   4290
      TabIndex        =   22
      Top             =   3840
      Width           =   1095
   End
   Begin TabDlg.SSTab TabRegistro 
      Height          =   3645
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6429
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   4
      TabHeight       =   617
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Registro"
      TabPicture(0)   =   "frmProveedorRegSistemaPension.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraEmisor"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame1 
         Caption         =   "Datos AFP/ONP"
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
         Height          =   1815
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   6090
         Begin VB.Frame fraAFP 
            BorderStyle     =   0  'None
            Height          =   1215
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   5895
            Begin VB.TextBox txtCUSP 
               Height          =   285
               Left            =   1800
               TabIndex        =   31
               Top             =   120
               Width           =   1635
            End
            Begin VB.ComboBox cboEntidad 
               Height          =   315
               Left            =   1800
               Style           =   2  'Dropdown List
               TabIndex        =   27
               Top             =   465
               Width           =   4095
            End
            Begin VB.OptionButton optTpoComision 
               Caption         =   "Comisión por Flujo"
               Height          =   255
               Index           =   0
               Left            =   1800
               TabIndex        =   26
               Top             =   915
               Value           =   -1  'True
               Width           =   1575
            End
            Begin VB.OptionButton optTpoComision 
               Caption         =   "Comisión por Saldo"
               Height          =   255
               Index           =   1
               Left            =   3600
               TabIndex        =   25
               Top             =   915
               Width           =   1695
            End
            Begin VB.Label Label8 
               Caption         =   "CUSP :"
               Height          =   255
               Left            =   0
               TabIndex        =   30
               Top             =   135
               Width           =   615
            End
            Begin VB.Label Label13 
               Caption         =   "Entidad :"
               Height          =   255
               Left            =   0
               TabIndex        =   29
               Top             =   525
               Width           =   735
            End
            Begin VB.Label Label15 
               Caption         =   "Tipo de Comisión :"
               Height          =   255
               Left            =   0
               TabIndex        =   28
               Top             =   915
               Width           =   1335
            End
         End
         Begin VB.OptionButton optSistemaPension 
            Caption         =   "ONP"
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   21
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optSistemaPension 
            Caption         =   "AFP"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.Frame fraEmisor 
         Caption         =   "Emisor"
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
         Height          =   1095
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   6090
         Begin VB.Label Label5 
            Caption         =   "R.U.C :"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   660
            Width           =   615
         End
         Begin VB.Label lblDOI 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   960
            TabIndex        =   17
            Top             =   645
            Width           =   1635
         End
         Begin VB.Label Label1 
            Caption         =   "Emisor :"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblEmisor 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   960
            TabIndex        =   15
            Top             =   300
            Width           =   4995
         End
      End
      Begin VB.TextBox txtOCEntrega 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   -72270
         TabIndex        =   6
         Top             =   1065
         Width           =   1785
      End
      Begin VB.TextBox txtOCNro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74460
         MaxLength       =   13
         TabIndex        =   5
         Top             =   585
         Width           =   1485
      End
      Begin VB.TextBox txtOCFecha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   -71685
         TabIndex        =   4
         Top             =   585
         Width           =   1200
      End
      Begin VB.TextBox txtOCPlazo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74235
         TabIndex        =   3
         Top             =   1080
         Width           =   1260
      End
      Begin VB.TextBox txtGRSerie 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74505
         MaxLength       =   4
         TabIndex        =   2
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtGRNro 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74130
         MaxLength       =   12
         TabIndex        =   1
         Top             =   840
         Width           =   1350
      End
      Begin MSMask.MaskEdBox txtGRFecha 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   -71565
         TabIndex        =   7
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label19 
         Caption         =   "Entrega"
         Height          =   240
         Left            =   -72870
         TabIndex        =   13
         Top             =   1095
         Width           =   1245
      End
      Begin VB.Label Label7 
         Caption         =   "Nº"
         Height          =   165
         Left            =   -74745
         TabIndex        =   12
         Top             =   660
         Width           =   315
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha Emisión"
         Height          =   165
         Left            =   -72870
         TabIndex        =   11
         Top             =   660
         Width           =   1065
      End
      Begin VB.Label Label14 
         Caption         =   "Plazo"
         Height          =   240
         Left            =   -74790
         TabIndex        =   10
         Top             =   1095
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Emisión"
         Height          =   165
         Left            =   -72660
         TabIndex        =   9
         Top             =   915
         Width           =   1035
      End
      Begin VB.Label Label11 
         Caption         =   "Nº"
         Height          =   165
         Left            =   -74790
         TabIndex        =   8
         Top             =   915
         Width           =   315
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000C&
         Height          =   1290
         Left            =   -74865
         Top             =   435
         Width           =   4590
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H8000000E&
         Height          =   1290
         Left            =   -74850
         Top             =   450
         Width           =   4635
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H8000000C&
         Height          =   1230
         Left            =   -74880
         Top             =   450
         Width           =   4650
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H8000000E&
         Height          =   1245
         Left            =   -74865
         Top             =   465
         Width           =   4620
      End
   End
End
Attribute VB_Name = "frmProveedorRegSistemaPension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
'** Nombre : frmProveedorRegSistemaPension
'** Descripción : Formulario para registrar el Tipo de Sistema Pensión del Proveedor
'** Creación : EJVG, 20140722 09:00:00 AM
'***********************************************************************************
Option Explicit
Dim fbRegistrar As Boolean
Dim fbEditar As Boolean
Dim fsPersCod As String
Public bOK As Boolean

Private Sub cmdGrabar_Click()
    Dim obj As NProveedorSistPens
    Dim oFun As NContFunciones
    Dim lsMovNro As String
    Dim bExito As Boolean
    
    Dim lnTpoSistPension As TipoSistemaPensionProveeedor
    Dim lnTpoComisionAFP As TipoComisionAFPProveeedor
    
    On Error GoTo ErrcmdGrabar
    
    If Not ValidarRegistrar Then Exit Sub
    
    Set oFun = New NContFunciones
    lsMovNro = oFun.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    Set oFun = Nothing
    lnTpoSistPension = IIf(optSistemaPension.Item(0).value, TipoSistemaPensionProveeedor.AFP, TipoSistemaPensionProveeedor.ONP)
    lnTpoComisionAFP = IIf(optTpoComision.Item(0).value, TipoComisionAFPProveeedor.Flujo, TipoComisionAFPProveeedor.Saldo)
    
    If MsgBox("Está información es vital para el calculo de las retenciones de AFP/ONP" & Chr(13) & "¿Está seguro de grabar la información?", vbYesNo + vbQuestion, "Aviso") = vbNo Then Exit Sub
    Set obj = New NProveedorSistPens
    
    If fbRegistrar Then
    bExito = obj.RegistrarDatosSistemaPension(lsMovNro, fsPersCod, lnTpoSistPension, Trim(txtCUSP.Text), Trim(Right(cboEntidad.Text, 13)), lnTpoComisionAFP)
    End If
    If fbEditar Then
        bExito = obj.EditarDatosSistemaPension(fsPersCod, lnTpoSistPension, Trim(txtCUSP.Text), Trim(Right(cboEntidad.Text, 13)), lnTpoComisionAFP)
    End If
    Set obj = Nothing
    
    If bExito Then
        MsgBox "Se ha registrado con éxito los datos de Sistema de Pensión del Proveedor", vbInformation, "Aviso"
        bOK = True
        Unload Me
        Exit Sub
    Else
        MsgBox "Ha sucedido un error, si el problema persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
    End If
    Exit Sub
ErrcmdGrabar:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    bOK = False
    CargarControles
End Sub
Private Sub optSistemaPension_Click(Index As Integer)
    LimpiarAFP
    If Index = 0 Then
        fraAFP.Enabled = True
    Else
        fraAFP.Enabled = False
    End If
End Sub
Private Sub LimpiarAFP()
    txtCUSP.Text = ""
    cboEntidad.ListIndex = -1
    optTpoComision.Item(0).value = True
End Sub
Private Function ValidarRegistrar() As Boolean
    If optSistemaPension.Item(0).value Then
        If Len(Trim(txtCUSP.Text)) = 0 Then
            MsgBox "Ud. debe ingresar el dato CUSP del Proveedor", vbInformation, "Aviso"
            EnfocaControl txtCUSP
            Exit Function
        End If
        If cboEntidad.ListIndex = -1 Then
            MsgBox "Ud. debe seleccionar la Entidad AFP del Proveedor", vbInformation, "Aviso"
            EnfocaControl cboEntidad
            Exit Function
        End If
        If optTpoComision.Item(0).value = False And optTpoComision.Item(1).value = False Then
            MsgBox "Ud. debe seleccionar el Tipo de Comisión del Proveedor", vbInformation, "Aviso"
            Exit Function
        End If
    End If
    ValidarRegistrar = True
End Function
Public Sub Registrar(ByVal psPersCod As String)
    fbRegistrar = True
    fsPersCod = psPersCod
    bOK = False
    If Not cargar_datos(fsPersCod) Then
        MsgBox "No se encontro datos del Proveedor", vbExclamation, "Aviso"
        Exit Sub
    End If
    Show 1
End Sub
Public Sub Editar(ByVal psPersCod As String)
    fbEditar = True
    fsPersCod = psPersCod
    bOK = False
    If Not cargar_datos_editar(fsPersCod) Then
        MsgBox "No se encontro datos del Proveedor", vbExclamation, "Aviso"
        Exit Sub
    End If
    Show 1
End Sub
Private Function cargar_datos(ByVal psPersCod As String) As Boolean
    Dim oPersona As New DPersonas
    Dim rsPersona As New ADODB.Recordset
    
    On Error GoTo ErrorCargar_datos
    Screen.MousePointer = 11
    Set rsPersona = oPersona.ObtieneDatosProveedorRetencSistPens(psPersCod, gdFecSis)
    If Not rsPersona.EOF Then
        lblEmisor.Caption = rsPersona!cPersNombre
        lblDOI.Caption = rsPersona!cDOI
        cargar_datos = True
    Else
        cargar_datos = False
    End If
    RSClose rsPersona
    Set oPersona = Nothing
    Screen.MousePointer = 0
    Exit Function
ErrorCargar_datos:
    cargar_datos = False
    MsgBox Err.Description, vbCritical, "Aviso"
End Function
Private Sub CargarControles()
    Dim oRol As New NProveedorSistPens
    Dim rsRol As New ADODB.Recordset
    
    cboEntidad.Clear
    Set rsRol = oRol.ListaAFP()
    Do While Not rsRol.EOF
        cboEntidad.AddItem rsRol!cPersNombre & Space(100) & rsRol!cPersCod
        rsRol.MoveNext
    Loop
    RSClose rsRol
    Set oRol = Nothing
End Sub
Private Sub optTpoComision_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl cmdGrabar
    End If
End Sub
Private Sub txtCUSP_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl cboEntidad
    End If
End Sub
Private Sub cboEntidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl IIf(optTpoComision.Item(0).value, optTpoComision.Item(0), optTpoComision.Item(1))
    End If
End Sub
Private Function cargar_datos_editar(ByVal psPersCod As String) As Boolean
    Dim oPersona As New DProveedorSistPens
    Dim rsPersona As New ADODB.Recordset
    
    On Error GoTo ErrorCargar_datos
    Screen.MousePointer = 11
    Set rsPersona = oPersona.ObtieneDatosSistemaPension(psPersCod)
    If Not rsPersona.EOF Then
        lblEmisor.Caption = rsPersona!cPersNombre
        lblDOI.Caption = rsPersona!cDOI
        
        If rsPersona!nTpoSistPens = 1 Then
            optSistemaPension(0).value = True
            txtCUSP.Text = rsPersona!AFP_CUSP
            cboEntidad.ListIndex = IndiceListaCombo(cboEntidad, rsPersona!AFP_cPersCod)
            If rsPersona!AFP_nTpoComision = 1 Then
                optTpoComision.Item(0).value = True
            Else
                optTpoComision.Item(1).value = True
            End If
        Else
            optSistemaPension(1).value = True
        End If
        
        cargar_datos_editar = True
    Else
        cargar_datos_editar = False
    End If
    RSClose rsPersona
    Set oPersona = Nothing
    Screen.MousePointer = 0
    Exit Function
ErrorCargar_datos:
    cargar_datos_editar = False
    MsgBox Err.Description, vbCritical, "Aviso"
End Function
