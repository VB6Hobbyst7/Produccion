VERSION 5.00
Begin VB.Form frmPersGarantias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Garantias de Cliente"
   ClientHeight    =   6555
   ClientLeft      =   1155
   ClientTop       =   1395
   ClientWidth     =   9045
   Icon            =   "frmPersGarantias.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdLimpiar 
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
      Height          =   390
      Left            =   6660
      TabIndex        =   45
      ToolTipText     =   "Salir(ALT+S)"
      Top             =   6030
      Width           =   1125
   End
   Begin VB.CommandButton CmdEliminar 
      Caption         =   "&Eliminar"
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
      Height          =   390
      Left            =   2415
      TabIndex        =   19
      ToolTipText     =   "Salir(ALT+S)"
      Top             =   6030
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Caption         =   "Clientes"
      Height          =   2445
      Left            =   75
      TabIndex        =   40
      Top             =   3435
      Width           =   8850
      Begin Sicmact.FlexEdit FERelPers 
         Height          =   1575
         Left            =   150
         TabIndex        =   14
         Top             =   270
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   2778
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "-Codigo-Nombre-Relacion"
         EncabezadosAnchos=   "400-1500-5000-1450"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-3"
         ListaControles  =   "0-1-0-3"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-L"
         FormatosEdit    =   "0-0-0-0"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         ColWidth0       =   405
         RowHeight0      =   285
      End
      Begin VB.CommandButton CmdCliNuevo 
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
         Height          =   300
         Left            =   6600
         TabIndex        =   15
         ToolTipText     =   "Salir(ALT+S)"
         Top             =   1965
         Width           =   1005
      End
      Begin VB.CommandButton CmdCliEliminar 
         Caption         =   "&Eliminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7635
         TabIndex        =   16
         ToolTipText     =   "Salir(ALT+S)"
         Top             =   1965
         Width           =   1005
      End
      Begin VB.CommandButton CmdCliAceptar 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6600
         TabIndex        =   43
         ToolTipText     =   "Salir(ALT+S)"
         Top             =   1965
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton CmdCliCancelar 
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
         Height          =   300
         Left            =   7635
         TabIndex        =   44
         ToolTipText     =   "Salir(ALT+S)"
         Top             =   1965
         Visible         =   0   'False
         Width           =   1005
      End
   End
   Begin VB.Frame fraZonaCbo 
      Height          =   1860
      Left            =   60
      TabIndex        =   32
      Top             =   1530
      Width           =   6240
      Begin VB.TextBox txtcomentarios 
         Height          =   540
         Left            =   1065
         MaxLength       =   60
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   1230
         Width           =   4845
      End
      Begin VB.Frame frazona 
         BorderStyle     =   0  'None
         Height          =   915
         Left            =   330
         TabIndex        =   33
         Top             =   210
         Width           =   5670
         Begin VB.ComboBox cmbPersUbiGeo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   285
            Index           =   3
            Left            =   3210
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Tag             =   "cboPrincipal"
            ToolTipText     =   "Urbanización"
            Top             =   450
            Width           =   1995
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   3210
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Tag             =   "cboPrincipal"
            ToolTipText     =   "Distrito"
            Top             =   105
            Width           =   1980
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   570
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Tag             =   "cboPrincipal"
            ToolTipText     =   "Provincia"
            Top             =   465
            Width           =   1935
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   570
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Tag             =   "cboPrincipal"
            ToolTipText     =   "Zona"
            Top             =   75
            Width           =   1920
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Zona :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   165
            Left            =   2625
            TabIndex        =   37
            Top             =   480
            Width           =   390
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Distrito :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   -15
            TabIndex        =   36
            Top             =   495
            Width           =   525
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Prov :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   2625
            TabIndex        =   35
            Top             =   165
            Width           =   375
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Dpto :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   15
            TabIndex        =   34
            Top             =   150
            Width           =   375
         End
      End
      Begin VB.Label Label1 
         Caption         =   " Zona :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   195
         TabIndex        =   39
         Top             =   30
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Comentarios"
         Height          =   195
         Left            =   105
         TabIndex        =   38
         Top             =   1260
         Width           =   870
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   6225
         Y1              =   1140
         Y2              =   1140
      End
   End
   Begin VB.Frame framontos 
      Height          =   1860
      Left            =   6375
      TabIndex        =   27
      Top             =   1530
      Width           =   2535
      Begin VB.TextBox txtMontoRea 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0.00;(0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   1110
         MaxLength       =   15
         TabIndex        =   11
         Tag             =   "txtPrincipal"
         Top             =   810
         Width           =   1245
      End
      Begin VB.TextBox txtMontotas 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0.00;(0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   1110
         MaxLength       =   15
         TabIndex        =   10
         Tag             =   "txtPrincipal"
         Top             =   420
         Width           =   1245
      End
      Begin VB.TextBox txtMontoxGrav 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0.00;(0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   1110
         MaxLength       =   15
         TabIndex        =   12
         Tag             =   "txtPrincipal"
         Top             =   1215
         Width           =   1245
      End
      Begin VB.Label lbltasa 
         AutoSize        =   -1  'True
         Caption         =   "Tasación :"
         Height          =   195
         Left            =   90
         TabIndex        =   31
         ToolTipText     =   "Monto Tasación"
         Top             =   465
         Width           =   750
      End
      Begin VB.Label lblrealizacion 
         AutoSize        =   -1  'True
         Caption         =   "Realización :"
         Height          =   195
         Left            =   75
         TabIndex        =   30
         Top             =   840
         Width           =   915
      End
      Begin VB.Label lblMontoGrav 
         AutoSize        =   -1  'True
         Caption         =   "Disponible :"
         Height          =   195
         Left            =   75
         TabIndex        =   29
         Top             =   1245
         Width           =   825
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Montos :"
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
         Index           =   3
         Left            =   870
         TabIndex        =   28
         ToolTipText     =   "Monto Tasación"
         Top             =   15
         Width           =   750
      End
   End
   Begin VB.Frame fraPrinc 
      Height          =   1470
      Left            =   75
      TabIndex        =   21
      ToolTipText     =   "Datos del Cliente"
      Top             =   60
      Width           =   8835
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
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
         Left            =   7440
         TabIndex        =   2
         ToolTipText     =   "Salir(ALT+S)"
         Top             =   210
         Width           =   1140
      End
      Begin VB.ComboBox cmbMoneda 
         Height          =   315
         Left            =   5715
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "cboPrincipal"
         ToolTipText     =   "Tipo Moneda"
         Top             =   600
         Width           =   2115
      End
      Begin VB.TextBox txtNumDoc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5715
         MaxLength       =   15
         TabIndex        =   1
         Tag             =   "txtPrincipal"
         Top             =   217
         Width           =   1500
      End
      Begin VB.ComboBox CmbDocGarant 
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
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Tag             =   "cboPrincipal"
         ToolTipText     =   "Tipo de Documentos"
         Top             =   225
         Width           =   3285
      End
      Begin VB.ComboBox CmbTipoGarant 
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "cboPrincipal"
         ToolTipText     =   "Tipos de Garantias"
         Top             =   600
         Width           =   3285
      End
      Begin VB.TextBox txtDescGarant 
         Height          =   330
         Left            =   1500
         MaxLength       =   60
         TabIndex        =   5
         Tag             =   "txtPrincipal"
         Top             =   975
         Width           =   5685
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Moneda :"
         Height          =   195
         Left            =   4980
         TabIndex        =   26
         Top             =   660
         Width           =   675
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Nº Doc. :"
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
         Left            =   4965
         TabIndex        =   25
         Top             =   285
         Width           =   810
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Garantía"
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
         Left            =   195
         TabIndex        =   24
         Top             =   285
         Width           =   1230
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Garantía"
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   660
         Width           =   990
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   210
         TabIndex        =   22
         Top             =   1035
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdSalir 
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
      Height          =   390
      Left            =   7800
      TabIndex        =   20
      ToolTipText     =   "Salir(ALT+S)"
      Top             =   6030
      Width           =   1125
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "&Editar"
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
      Height          =   390
      Left            =   1275
      TabIndex        =   18
      ToolTipText     =   "Salir(ALT+S)"
      Top             =   6030
      Width           =   1125
   End
   Begin VB.CommandButton CmdCancelar 
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
      Height          =   390
      Left            =   1275
      TabIndex        =   42
      ToolTipText     =   "Salir(ALT+S)"
      Top             =   6030
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton CmdNuevo 
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
      Height          =   390
      Left            =   135
      TabIndex        =   17
      ToolTipText     =   "Salir(ALT+S)"
      Top             =   6030
      Width           =   1125
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   135
      TabIndex        =   41
      ToolTipText     =   "Salir(ALT+S)"
      Top             =   6030
      Visible         =   0   'False
      Width           =   1125
   End
End
Attribute VB_Name = "frmPersGarantias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina:           frmPersGarantias
'***     Descripcion:       Realiza el Mantenimiento y Registro de Nuevas Garantias
'***     Creado por:        NSSE
'***     Maquina:           07SIST_08
'***     Fecha-Tiempo:         08/06/2001 12:15:13 PM
'***     Ultima Modificacion: Creacion del Formulario
'*****************************************************************************************
Option Explicit
Private Enum TGarantiaTipoCombo
    ComboDpto = 1
    ComboProv = 2
    ComboDist = 3
    ComboZona = 4
End Enum

Enum TGarantiaTipoInicio
    RegistroGarantia = 1
    MantenimientoGarantia = 2
    ConsultaGarant = 3
End Enum

Dim Nivel1() As String
Dim ContNiv1 As Integer
Dim Nivel2() As String
Dim ContNiv2 As Integer
Dim Nivel3() As String
Dim ContNiv3 As Integer
Dim Nivel4() As String
Dim ContNiv4 As Integer
Dim bEstadoCargando As Boolean
Dim cmdEjecutar As Integer

Dim vTipoInicio As TGarantiaTipoInicio

Public Sub Inicio(ByVal pvTipoIni As TGarantiaTipoInicio)
    
    vTipoInicio = pvTipoIni
    If vTipoInicio = ConsultaGarant Then
        cmdNuevo.Enabled = False
        cmdEditar.Enabled = False
        cmdEliminar.Enabled = False
    End If
    
    If vTipoInicio = RegistroGarantia Then
        cmdNuevo.Enabled = True
        cmdEditar.Enabled = False
        cmdEliminar.Enabled = False
    End If
    
    If vTipoInicio = MantenimientoGarantia Then
        cmdNuevo.Enabled = False
        cmdEditar.Enabled = True
        cmdEliminar.Enabled = True
    End If
    
    Me.Show 1
End Sub
Private Function ValidaBuscar() As Boolean
    ValidaBuscar = True
    
    If CmbDocGarant.ListIndex = -1 Then
        MsgBox "Seleccione un tipo de Documento de Garantia", vbInformation, "Aviso"
        ValidaBuscar = False
        Exit Function
    End If
    
    If Trim(txtNumDoc.Text) = "" Then
        MsgBox "Ingrese el Numero de Documento", vbInformation, "Aviso"
        ValidaBuscar = False
        Exit Function
    End If
    
End Function
Private Function ValidaDatos() As Boolean
Dim i As Integer
Dim Enc As Boolean

    ValidaDatos = True
    
    'Verifica seleccion de Documento de Garantia
    If CmbDocGarant.ListIndex = -1 Then
        MsgBox "Seleccione un Tipo de Documento de Garantia", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
    'Verifica Ingreso de Numero de Documento de Garantia
    If Trim(txtNumDoc.Text) = "" Then
        MsgBox "Ingrese el Numero de Documento de la Garantia", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
    'Verifica seleccion de Tipo de Garantia
    If CmbTipoGarant.ListIndex = -1 Then
        MsgBox "Seleccione un Tipo de Garantia", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
    'Verifica seleccion de Moneda
    If cmbMoneda.ListIndex = -1 Then
        MsgBox "Seleccione la Moneda", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
    'Verifica la Zona
    If cmbPersUbiGeo(3).ListIndex = -1 Then
        MsgBox "Seleccione La Zona donde se Ubica la Garantia", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
    'Verifica Monto de Tasacion
    If Trim(txtMontotas.Text) = "" Or Trim(txtMontotas.Text) = "0.00" Then
        MsgBox "El Monto de Tasacion debe ser Mayor que Cero", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
    
    'Verifica Monto de Realizacion
    If Trim(txtMontoRea.Text) = "" Or Trim(txtMontoRea.Text) = "0.00" Then
        MsgBox "El Monto de Realizacion debe ser Mayor que Cero", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
    'Verifica Monto de Realizacion
    If Trim(txtMontoxGrav.Text) = "" Or Trim(txtMontoxGrav.Text) = "0.00" Then
        MsgBox "El Monto de Disponible debe ser Mayor que Cero", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
   Enc = False
   ' Verifica Existencia de Titular de la Garantia
   For i = 1 To FERelPers.Rows - 1
        If Trim(Right(FERelPers.TextMatrix(i, 3), 15)) <> "" Then
            If CInt(Trim(Right(FERelPers.TextMatrix(i, 3), 15))) = gPersRelGarantiaTitular Then
                Enc = True
                Exit For
            End If
        End If
   Next i
   If Not Enc Then
        MsgBox "Ingrese un Titular para la Garantia", vbInformation, "Aviso"
        ValidaDatos = False
        CmdCliNuevo.SetFocus
        Exit Function
   End If
End Function
Private Sub HabilitaIngreso(ByVal pbHabilita As Boolean)
    
        CmbDocGarant.Enabled = pbHabilita
        txtNumDoc.Enabled = pbHabilita
        cmdBuscar.Enabled = pbHabilita
        CmbTipoGarant.Enabled = pbHabilita
        cmbMoneda.Enabled = pbHabilita
        txtDescGarant.Enabled = pbHabilita
        cmbPersUbiGeo(0).Enabled = pbHabilita
        cmbPersUbiGeo(1).Enabled = pbHabilita
        cmbPersUbiGeo(2).Enabled = pbHabilita
        cmbPersUbiGeo(3).Enabled = pbHabilita
        txtMontotas.Enabled = pbHabilita
        txtMontoRea.Enabled = pbHabilita
        txtMontoxGrav.Enabled = pbHabilita
        Txtcomentarios.Enabled = pbHabilita
        FERelPers.lbEditarFlex = False
        CmdCliNuevo.Enabled = pbHabilita
        CmdCliEliminar.Enabled = pbHabilita
        cmdNuevo.Enabled = Not pbHabilita
        cmdNuevo.Visible = Not pbHabilita
        CmdAceptar.Enabled = pbHabilita
        CmdAceptar.Visible = pbHabilita
        cmdEditar.Enabled = Not pbHabilita
        cmdEditar.Visible = Not pbHabilita
        cmdCancelar.Enabled = pbHabilita
        cmdCancelar.Visible = pbHabilita
        cmdEliminar.Enabled = Not pbHabilita
        cmdEliminar.Visible = Not pbHabilita
        cmdSalir.Enabled = Not pbHabilita
        CmdLimpiar.Enabled = Not pbHabilita
        cmdBuscar.Enabled = Not pbHabilita
        If vTipoInicio = MantenimientoGarantia Then
            cmdNuevo.Enabled = False
        End If
End Sub

Private Sub CargaUbicacionesGeograficas()
Dim Conn As DConecta
Dim sSql As String
Dim R As ADODB.Recordset
Dim i As Integer
Dim nPos As Integer

On Error GoTo ErrCargaUbicacionesGeograficas
    Set Conn = New DConecta
    'Carga Niveles
    sSql = "Select *, 1 p from UbicacionGeografica where cUbiGeoCod like '1%'"
    sSql = sSql & " Union "
    sSql = sSql & " select *, 2 p from UbicacionGeografica where cUbiGeoCod like '2%' "
    sSql = sSql & " Union "
    sSql = sSql & " select *, 3 p from UbicacionGeografica where cUbiGeoCod like '3%' "
    sSql = sSql & " Union "
    sSql = sSql & " select *, 4 p from UbicacionGeografica where cUbiGeoCod like '4%' order by p,cUbiGeoDescripcion "
    ContNiv1 = 0
    ContNiv2 = 0
    ContNiv3 = 0
    ContNiv4 = 0
    
    Conn.AbreConexion
    Set R = Conn.CargaRecordSet(sSql)
    Do While Not R.EOF
        Select Case R!P
            Case 1 ' Departamento
                ContNiv1 = ContNiv1 + 1
                ReDim Preserve Nivel1(ContNiv1)
                Nivel1(ContNiv1 - 1) = Trim(R!cUbiGeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
            Case 2 ' Provincia
                ContNiv2 = ContNiv2 + 1
                ReDim Preserve Nivel2(ContNiv2)
                Nivel2(ContNiv2 - 1) = Trim(R!cUbiGeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
            Case 3 'Distrito
                ContNiv3 = ContNiv3 + 1
                ReDim Preserve Nivel3(ContNiv3)
                Nivel3(ContNiv3 - 1) = Trim(R!cUbiGeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
            Case 4 'Zona
                ContNiv4 = ContNiv4 + 1
                ReDim Preserve Nivel4(ContNiv4)
                Nivel4(ContNiv4 - 1) = Trim(R!cUbiGeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
        End Select
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Conn.CierraConexion
    Set Conn = Nothing
    
    'Carga el Nivel1 en el Control
    cmbPersUbiGeo(0).Clear
    For i = 0 To ContNiv1 - 1
        cmbPersUbiGeo(0).AddItem Nivel1(i)
        If Trim(Right(Nivel1(i), 12)) = "113000000000" Then
            nPos = i
        End If
    Next i
    cmbPersUbiGeo(0).ListIndex = nPos
    cmbPersUbiGeo(2).Clear
    cmbPersUbiGeo(3).Clear
    Exit Sub
    
ErrCargaUbicacionesGeograficas:
    MsgBox Err.Description, vbInformation, "Aviso"
    
End Sub

Private Sub LimpiaPantalla()
    Call LimpiaControles(Me)
    Call LimpiaFlex(FERelPers)
    Call InicializaCombos(Me)
    txtMontotas.Text = "0.00"
    txtMontoRea.Text = "0.00"
    txtMontoxGrav.Text = "0.00"
    cmdEditar.Enabled = False
    cmdEliminar.Enabled = False
End Sub

Private Function CargaDatos(ByVal psTipoDoc As String, ByVal psNroDoc As String) As Boolean
'Dim oGarantia As DGarantia
'Dim R As ADODB.Recordset
'Dim RRelPers As ADODB.Recordset
'
'    On Error GoTo ErrorCargaDatos
'    Set oGarantia = New DGarantia
'    Set R = oGarantia.RecuperaGarantia(psTipoDoc, psNroDoc)
'    Set oGarantia = Nothing
'    If R.RecordCount = 0 Then
'        R.Close
'        Set R = Nothing
'        CargaDatos = False
'        Exit Function
'    Else
'        CargaDatos = True
'    End If
'    'CmbDocGarant.ListIndex = IndiceListaCombo(CmbDocGarant, R!cTpoDoc)
'    'txtNumDoc.Text = R!cNroDoc
'    CmbTipoGarant.ListIndex = IndiceListaCombo(CmbTipoGarant, Trim(Str(R!nTpoGarantia)))
'    cmbMoneda.ListIndex = IndiceListaCombo(cmbMoneda, Trim(Str(R!nMoneda)))
'    txtDescGarant.Text = R!cDescripcion
'
'    'Carga Ubicacion Geografica
'    cmbPersUbiGeo(0).ListIndex = IndiceListaCombo(cmbPersUbiGeo(0), Space(30) & "1" & Mid(R!cZona, 2, 2) & String(9, "0"))
'    cmbPersUbiGeo(1).ListIndex = IndiceListaCombo(cmbPersUbiGeo(1), Space(30) & "2" & Mid(R!cZona, 2, 4) & String(7, "0"))
'    cmbPersUbiGeo(2).ListIndex = IndiceListaCombo(cmbPersUbiGeo(2), Space(30) & "3" & Mid(R!cZona, 2, 6) & String(5, "0"))
'    cmbPersUbiGeo(3).ListIndex = IndiceListaCombo(cmbPersUbiGeo(3), Space(30) & R!cZona)
'
'    txtMontotas.Text = Format(R!nTasacion, "#0.00")
'    txtMontoRea.Text = Format(R!nRealizacion, "#0.00")
'    txtMontoxGrav.Text = Format(R!nPorGravar, "#0.00")
'    txtcomentarios.Text = Trim(IIf(IsNull(R!cComentario), "", R!cComentario))
'
'    'Personas Relacionadas con Garantias
'    Set oGarantia = New DGarantia
'    Set RRelPers = oGarantia.RecuperaRelacPersonaGarantia(R!cTpoDoc, R!cNroDoc)
'    Set oGarantia = Nothing
'    Call LimpiaFlex(FERelPers)
'    Do While Not RRelPers.EOF
'        FERelPers.AdicionaFila
'        FERelPers.TextMatrix(RRelPers.Bookmark, 1) = RRelPers!cPersCod
'        FERelPers.TextMatrix(RRelPers.Bookmark, 2) = RRelPers!cPersNombre
'        FERelPers.TextMatrix(RRelPers.Bookmark, 3) = RRelPers!cRelacion
'        RRelPers.MoveNext
'    Loop
'    RRelPers.Close
'    Set RRelPers = Nothing
'
'    R.Close
'    Set R = Nothing
'
'    Exit Function
'
'ErrorCargaDatos:
'        MsgBox Err.Description, vbCritical, "Aviso"
End Function
Private Sub CargaControles()
'Dim oGarantia As DGarantia
'Dim R As ADODB.Recordset
'Dim oConstante As DConstante
'
'    On Error GoTo ERRORCargaControles
'
'    'Carga Ubicaciones Geograficas
'        Call CargaUbicacionesGeograficas
'    'Carga Tipos de Garantia
'        Call CambiaTamañoCombo(CmbTipoGarant)
'        Call CargaComboConstante(gPersGarantia, CmbTipoGarant)
'    'Carga Tipos de Documentos de Garantia
'        Set oGarantia = New DGarantia
'        Set R = oGarantia.RecuperaTiposDocumGarantias
'        Set oGarantia = Nothing
'        CmbDocGarant.Clear
'        Do While Not R.EOF
'            CmbDocGarant.AddItem R!cDocDesc & Space(150) & R!nDocTpo
'            R.MoveNext
'        Loop
'        R.Close
'        Set R = Nothing
'        Call CambiaTamañoCombo(CmbDocGarant, 300)
'    'Carga Monedas
'        Call CargaComboConstante(gMoneda, cmbMoneda)
'
'    'Carga Relacion de Personas con Garantia
'        Set oConstante = New DConstante
'        FERelPers.CargaCombo oConstante.RecuperaConstantes(gPersRelGarantia)
'        Set oConstante = Nothing
'        Exit Sub
'
'ERRORCargaControles:
'        MsgBox Err.Description, vbCritical, "Aviso"

End Sub
Private Sub ActualizaCombo(ByVal psValor As String, ByVal TipoCombo As TGarantiaTipoCombo)
Dim i As Integer
Dim sCodigo As String
    
    sCodigo = Trim(Right(psValor, 15))
    Select Case TipoCombo
        Case ComboProv
            cmbPersUbiGeo(1).Clear
            If Len(sCodigo) > 3 Then
                cmbPersUbiGeo(1).Clear
                For i = 0 To ContNiv2 - 1
                    If Mid(sCodigo, 2, 2) = Mid(Trim(Right(Nivel2(i), 15)), 2, 2) Then
                        cmbPersUbiGeo(1).AddItem Nivel2(i)
                    End If
                Next i
            Else
                cmbPersUbiGeo(1).AddItem psValor
            End If
        Case ComboDist
            cmbPersUbiGeo(2).Clear
            If Len(sCodigo) > 3 Then
                For i = 0 To ContNiv3 - 1
                    If Mid(sCodigo, 2, 4) = Mid(Trim(Right(Nivel3(i), 15)), 2, 4) Then
                        cmbPersUbiGeo(2).AddItem Nivel3(i)
                    End If
                Next i
            Else
                cmbPersUbiGeo(2).AddItem psValor
            End If
        Case ComboZona
            cmbPersUbiGeo(3).Clear
            If Len(sCodigo) > 3 Then
                For i = 0 To ContNiv4 - 1
                    If Mid(sCodigo, 2, 6) = Mid(Trim(Right(Nivel4(i), 15)), 2, 6) Then
                        cmbPersUbiGeo(3).AddItem Nivel4(i)
                    End If
                Next i
            Else
                cmbPersUbiGeo(3).AddItem psValor
            End If
    End Select
End Sub




Private Sub CmbDocGarant_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        txtNumDoc.SetFocus
     End If
End Sub

Private Sub CmbMoneda_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        txtDescGarant.SetFocus
     End If
End Sub

Private Sub cmbPersUbiGeo_Click(Index As Integer)
        Select Case Index
            Case 0 'Combo Dpto
                Call ActualizaCombo(cmbPersUbiGeo(0).Text, ComboProv)
                If Not bEstadoCargando Then
                    cmbPersUbiGeo(2).Clear
                    cmbPersUbiGeo(3).Clear
                End If
            Case 1 'Combo Provincia
                Call ActualizaCombo(cmbPersUbiGeo(1).Text, ComboDist)
                If Not bEstadoCargando Then
                    cmbPersUbiGeo(3).Clear
                End If
            Case 2 'Combo Distrito
                Call ActualizaCombo(cmbPersUbiGeo(2).Text, ComboZona)
        End Select
End Sub


Private Sub cmbPersUbiGeo_KeyPress(Index As Integer, KeyAscii As Integer)
     If KeyAscii = 13 Then
        Select Case Index
            Case 0
                cmbPersUbiGeo(1).SetFocus
            Case 1
                cmbPersUbiGeo(2).SetFocus
            Case 2
                cmbPersUbiGeo(3).SetFocus
            Case 3
                txtMontotas.SetFocus
        End Select
     End If
End Sub

Private Sub CmbTipoGarant_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        cmbMoneda.SetFocus
     End If
End Sub

Private Sub CmdAceptar_Click()
'Dim oGarantia As DGarantia
'Dim RelPers() As String
'Dim i As Integer
'
'    On Error GoTo ErrorCmdAceptar_Click
'    If Not ValidaDatos Then
'        Exit Sub
'    End If
'
'    If FERelPers.Rows = 2 And FERelPers.TextMatrix(1, 1) = "" Then
'        ReDim RelPers(0, 0)
'    Else
'        ReDim RelPers(FERelPers.Rows - 1, 4)
'        For i = 1 To FERelPers.Rows - 1
'            RelPers(i - 1, 0) = FERelPers.TextMatrix(i, 1) 'Codigo de Persona
'            RelPers(i - 1, 1) = Trim(Right(CmbDocGarant.Text, 10)) 'Tipo de Doc de Garantia
'            RelPers(i - 1, 2) = Trim(txtNumDoc.Text) 'Numero de Documento
'            RelPers(i - 1, 3) = Right("00" & Trim(Right(FERelPers.TextMatrix(i, 3), 10)), 2) 'Relacion
'        Next i
'    End If
'    If cmdEjecutar = 1 Then
'        Set oGarantia = New DGarantia
'            Call oGarantia.NuevaGarantia(Trim(Right(CmbDocGarant.Text, 10)), Trim(txtNumDoc.Text), _
'                    Trim(Right(CmbTipoGarant.Text, 10)), Trim(Right(cmbMoneda.Text, 10)), Trim(txtDescGarant.Text), _
'                    Trim(Right(cmbPersUbiGeo(3).Text, 15)), CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text), _
'                    Trim(txtcomentarios.Text), RelPers, gdFecSis)
'        Set oGarantia = Nothing
'    Else
'        Set oGarantia = New DGarantia
'            Call oGarantia.ActualizaGarantia(Trim(Right(CmbDocGarant.Text, 10)), Trim(txtNumDoc.Text), _
'                    Trim(Right(CmbTipoGarant.Text, 10)), Trim(Right(cmbMoneda.Text, 10)), Trim(txtDescGarant.Text), _
'                    Trim(Right(cmbPersUbiGeo(3).Text, 15)), CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text), _
'                    Trim(txtcomentarios.Text), RelPers, gdFecSis)
'        Set oGarantia = Nothing
'    End If
'    cmdEjecutar = -1
'    Call HabilitaIngreso(False)
'    Exit Sub
'
'ErrorCmdAceptar_Click:
'        MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdBuscar_Click()
    If CmbDocGarant.ListIndex = -1 Then
        MsgBox "Seleccione un tipo de Documento de Garantia", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(Trim(txtNumDoc.Text)) = 0 Then
        MsgBox "Ingrese el numero del Documento", vbInformation, "Aviso"
        Exit Sub
    End If
    'Call LimpiaPantalla
    If CargaDatos(Trim(Right(CmbDocGarant.Text, 10)), Trim(txtNumDoc.Text)) Then
        cmdEditar.Enabled = True
        cmdEliminar.Enabled = True
        cmdEditar.SetFocus
        If vTipoInicio = ConsultaGarant Then
            cmdBuscar.SetFocus
            cmdEditar.Enabled = False
            cmdEliminar.Enabled = False
        End If
    Else
        MsgBox "Garantia no existe", vbInformation, "Aviso"
        CmbDocGarant.SetFocus
        Call CmdLimpiar_Click
    End If
End Sub

Private Sub CmdCancelar_Click()
    If cmdEjecutar = 2 Then
        CargaDatos Trim(Right(CmbDocGarant.Text, 10)), Trim(txtNumDoc.Text)
    Else
        If cmdEjecutar = 1 Then
            Call LimpiaPantalla
        End If
    End If
    Call HabilitaIngreso(False)
    CmbDocGarant.Enabled = True
    txtNumDoc.Enabled = True
    CmbDocGarant.SetFocus
    cmdEjecutar = -1
End Sub

Private Sub CmdCliAceptar_Click()
Dim i As Integer
Dim oGarantia As NGarantia
Dim RelPers() As String
    For i = 1 To FERelPers.Rows - 1
        If Len(Trim(FERelPers.TextMatrix(i, 1))) < 13 Then
            MsgBox "Codigo de Persona Incorrecto", vbInformation, "Aviso"
            FERelPers.Row = i
            FERelPers.Col = 1
            FERelPers.SetFocus
            Exit Sub
        End If
        If Len(Trim(FERelPers.TextMatrix(i, 3))) = 0 Then
            MsgBox "Relacion de Persona Con la Garantias es Incorrecto", vbInformation, "Aviso"
            FERelPers.Row = i
            FERelPers.Col = 3
            FERelPers.SetFocus
            Exit Sub
        End If
    Next i
    ReDim RelPers(FERelPers.Rows - 1)
    For i = 1 To FERelPers.Rows - 1
        RelPers(i - 1) = FERelPers.TextMatrix(i, 3)
    Next i
    Set oGarantia = New NGarantia
    If oGarantia.ValidaDatos(RelPers, CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text)) <> "" Then
        MsgBox oGarantia.ValidaDatos(RelPers, CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text)), vbInformation, "Aviso"
        Exit Sub
    End If
    Set oGarantia = Nothing

    FERelPers.lbEditarFlex = False
    CmdCliNuevo.Visible = True
    CmdCliEliminar.Visible = True
    CmdCliAceptar.Visible = False
    CmdCliCancelar.Visible = False
    CmdAceptar.Enabled = True
    cmdCancelar.Enabled = True
End Sub

Private Sub CmdCliEliminar_Click()
    If FERelPers.Row < 1 Then
        Exit Sub
    End If
    If MsgBox("Se va a Eliminar a la Persona " & FERelPers.TextMatrix(FERelPers.Row, 2) & ", Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        If FERelPers.Row = 1 Then
            FERelPers.TextMatrix(1, 0) = ""
            FERelPers.TextMatrix(1, 1) = ""
            FERelPers.TextMatrix(1, 2) = ""
            FERelPers.TextMatrix(1, 3) = ""
        Else
            Call FERelPers.EliminaFila(FERelPers.Row)
        End If
    End If
End Sub

Private Sub CmdCliNuevo_Click()
    FERelPers.lbEditarFlex = True
    FERelPers.AdicionaFila
    CmdCliNuevo.Visible = False
    CmdCliEliminar.Visible = False
    CmdCliAceptar.Visible = True
    CmdCliCancelar.Visible = True
    CmdAceptar.Enabled = False
    cmdCancelar.Enabled = False
    FERelPers.SetFocus
End Sub

Private Sub CmdEditar_Click()
    Call HabilitaIngreso(True)
    CmbDocGarant.Enabled = False
    txtNumDoc.Enabled = False
    CmbTipoGarant.SetFocus
    cmdEjecutar = 2
End Sub

Private Sub CmdEliminar_Click()
Dim oGarantia As DGarantia
    If MsgBox("Se va a Eliminar la Garantia, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Set oGarantia = New DGarantia
        Call oGarantia.EliminarGraantia(Trim(Right(CmbDocGarant.Text, 15)), Trim(txtNumDoc.Text))
        Set oGarantia = Nothing
        Call LimpiaPantalla
        Call CmdLimpiar_Click
    End If
    CmdCancelar_Click
    cmdBuscar.Enabled = True
End Sub

Private Sub CmdLimpiar_Click()
    Call LimpiaPantalla
    CmbDocGarant.Enabled = True
    txtNumDoc.Enabled = True
    cmdBuscar.Enabled = True
    CmbDocGarant.SetFocus
End Sub

Private Sub CmdNuevo_Click()
    Call HabilitaIngreso(True)
    Call LimpiaPantalla
    Call InicializaCombos(Me)
    cmdEjecutar = 1
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    bEstadoCargando = True
    Call CargaControles
    Call HabilitaIngreso(False)
    CmbDocGarant.Enabled = True
    txtNumDoc.Enabled = True
    cmdBuscar.Enabled = True
    bEstadoCargando = False
    cmdEjecutar = -1
    cmdEliminar.Enabled = False
    cmdEditar.Enabled = False
End Sub

Private Sub txtcomentarios_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
End Sub

Private Sub txtDescGarant_KeyPress(KeyAscii As Integer)
     KeyAscii = Letras(KeyAscii)
     If KeyAscii = 13 Then
        cmbPersUbiGeo(0).SetFocus
     End If
End Sub

Private Sub txtMontoRea_GotFocus()
    fEnfoque txtMontoRea
End Sub

Private Sub txtMontoRea_KeyPress(KeyAscii As Integer)
Dim oGarantia As NGarantia
Dim nValor As Double
Dim sCad As String

     KeyAscii = NumerosDecimales(txtMontoRea, KeyAscii)
     If KeyAscii = 13 Then
        Set oGarantia = New NGarantia
        sCad = oGarantia.ValidaDatos("", CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text), True)
        If Not sCad = "" Then
            MsgBox sCad, vbInformation, "Aviso"
            Exit Sub
        End If
        nValor = oGarantia.PorcentajeGarantia(gPersGarantia & Trim(Right(CmbTipoGarant.Text, 10)))
        Set oGarantia = Nothing
        If nValor > 1 Then
            txtMontoxGrav.Text = Format(nValor, "#0.00")
        Else
            txtMontoxGrav.Text = Format(nValor * CDbl(txtMontotas.Text), "#0.00")
        End If
        txtMontoxGrav.SetFocus
     End If
End Sub

Private Sub txtMontoRea_LostFocus()
    If Trim(txtMontoRea.Text) = "" Then
        txtMontoRea.Text = "0.00"
    Else
        txtMontoRea.Text = Format(txtMontoRea.Text, "#0.00")
    End If
End Sub

Private Sub txtMontotas_GotFocus()
    fEnfoque txtMontotas
End Sub

Private Sub txtMontotas_KeyPress(KeyAscii As Integer)

     KeyAscii = NumerosDecimales(txtMontotas, KeyAscii)
     If KeyAscii = 13 Then
        txtMontoRea.SetFocus
     End If
End Sub

Private Sub txtMontotas_LostFocus()
    If Trim(txtMontotas.Text) = "" Then
        txtMontotas.Text = "0.00"
    Else
        txtMontotas.Text = Format(txtMontotas.Text, "#0.00")
    End If
End Sub

Private Sub txtMontoxGrav_GotFocus()
    fEnfoque txtMontoxGrav
End Sub

Private Sub txtMontoxGrav_KeyPress(KeyAscii As Integer)
Dim oGarantia As NGarantia
Dim sCad As String

     KeyAscii = NumerosDecimales(txtMontoxGrav, KeyAscii)
     If KeyAscii = 13 Then
        Set oGarantia = New NGarantia
        sCad = oGarantia.ValidaDatos("", CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text), True)
        If Not sCad = "" Then
            MsgBox sCad, vbInformation, "Aviso"
            Exit Sub
        End If
        Txtcomentarios.SetFocus
     End If
End Sub

Private Sub txtMontoxGrav_LostFocus()
    If Trim(txtMontoxGrav.Text) = "" Then
        txtMontoxGrav.Text = "0.00"
    Else
        txtMontoxGrav.Text = Format(txtMontoxGrav.Text, "#0.00")
    End If
End Sub

Private Sub txtNumDoc_GotFocus()
    fEnfoque txtNumDoc
End Sub

Private Sub txtNumDoc_KeyPress(KeyAscii As Integer)
     KeyAscii = NumerosEnteros(KeyAscii)
     If KeyAscii = 13 Then
        If cmdBuscar.Enabled Then
            cmdBuscar.SetFocus
        Else
            CmbTipoGarant.SetFocus
        End If
     End If
End Sub
