VERSION 5.00
Begin VB.Form frmColRecExpedRelacionados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recuperaciones-Expedientes-Relacionados"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraCorrelativo 
      Height          =   615
      Left            =   6240
      TabIndex        =   56
      Top             =   0
      Width           =   2175
      Begin VB.TextBox TxtCorrelativo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Nro Correlativo"
         Height          =   195
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.Frame FraProceso 
      Height          =   735
      Left            =   60
      TabIndex        =   50
      Top             =   1560
      Width           =   8355
      Begin VB.ComboBox cbxTipoCob 
         Height          =   315
         ItemData        =   "frmColRecExpedRelacionados.frx":0000
         Left            =   1500
         List            =   "frmColRecExpedRelacionados.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox cbDemanda 
         Height          =   315
         ItemData        =   "frmColRecExpedRelacionados.frx":0004
         Left            =   4140
         List            =   "frmColRecExpedRelacionados.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label Label18 
         Caption         =   "Tipo de Cobranza"
         Height          =   195
         Left            =   120
         TabIndex        =   52
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Demanda"
         Height          =   195
         Left            =   3405
         TabIndex        =   51
         Top             =   300
         Width           =   825
      End
   End
   Begin VB.Frame fraDatos 
      Enabled         =   0   'False
      Height          =   3945
      Left            =   60
      TabIndex        =   29
      Top             =   2280
      Width           =   8355
      Begin VB.Frame Frame1 
         Caption         =   "Parte Procesal"
         Height          =   1005
         Left            =   60
         TabIndex        =   59
         Top             =   2800
         Width           =   8235
         Begin VB.ComboBox CboProcesal 
            Height          =   315
            ItemData        =   "frmColRecExpedRelacionados.frx":0008
            Left            =   1080
            List            =   "frmColRecExpedRelacionados.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   600
            Width           =   2595
         End
         Begin SICMACT.TxtBuscar AxEncargado 
            Height          =   285
            Left            =   1080
            TabIndex        =   16
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TipoBusqueda    =   3
         End
         Begin VB.Label Label15 
            Caption         =   "Parte Procesal"
            Height          =   435
            Left            =   160
            TabIndex        =   62
            Top             =   500
            Width           =   795
         End
         Begin VB.Label LblEncargado 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2880
            TabIndex        =   61
            Top             =   240
            Width           =   5295
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Encargado:"
            Height          =   195
            Left            =   180
            TabIndex        =   60
            Top             =   270
            Width           =   825
         End
      End
      Begin VB.ComboBox CboProcesos 
         Height          =   315
         ItemData        =   "frmColRecExpedRelacionados.frx":000C
         Left            =   3720
         List            =   "frmColRecExpedRelacionados.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1920
         Width           =   1635
      End
      Begin VB.Frame fraEstudioJur 
         Caption         =   "Estudio Juridico "
         Height          =   645
         Left            =   60
         TabIndex        =   53
         Top             =   120
         Width           =   8235
         Begin SICMACT.TxtBuscar AxCodAbogado 
            Height          =   285
            Left            =   840
            TabIndex        =   8
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TipoBusqueda    =   3
         End
         Begin VB.Label Label8 
            Caption         =   "Estudio Juridico"
            Height          =   435
            Left            =   120
            TabIndex        =   55
            Top             =   180
            Width           =   585
         End
         Begin VB.Label lblNomAbogado 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2640
            TabIndex        =   54
            Top             =   240
            Width           =   5535
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame fraInforme 
         Height          =   2055
         Left            =   6360
         TabIndex        =   39
         Top             =   765
         Width           =   1935
         Begin VB.CommandButton cmdComplementos 
            Caption         =   "Com&plementos"
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
            Left            =   420
            TabIndex        =   49
            Top             =   1575
            Width           =   1395
         End
         Begin VB.CommandButton cmdProbatorios 
            Caption         =   "&M Probatorio"
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
            Left            =   420
            TabIndex        =   48
            Top             =   1245
            Width           =   1395
         End
         Begin VB.CommandButton cmdJuridico 
            Caption         =   "Fund.&Juridico"
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
            Left            =   420
            TabIndex        =   47
            Top             =   915
            Width           =   1395
         End
         Begin VB.CommandButton cmdHecho 
            Caption         =   "&Fund.Hecho"
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
            Left            =   420
            TabIndex        =   46
            Top             =   570
            Width           =   1395
         End
         Begin VB.CommandButton cmdPetitorio 
            Caption         =   "&Petitorio"
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
            Left            =   420
            TabIndex        =   45
            Top             =   240
            Width           =   1395
         End
         Begin VB.CheckBox chkPetitorio 
            Caption         =   "Petitorio"
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   195
         End
         Begin VB.CheckBox chkFundHecho 
            Caption         =   "Fund. Hecho"
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   43
            Top             =   570
            Width           =   195
         End
         Begin VB.CheckBox chkFundJuridico 
            Caption         =   "Fund. Juridico"
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   42
            Top             =   900
            Width           =   195
         End
         Begin VB.CheckBox chkMedProb 
            Caption         =   "Medios Prob"
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   41
            Top             =   1230
            Width           =   195
         End
         Begin VB.CheckBox chkComplementos 
            Caption         =   "Complementos"
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   40
            Top             =   1560
            Width           =   195
         End
      End
      Begin VB.TextBox txtNroExp 
         Height          =   285
         Left            =   3720
         MaxLength       =   10
         TabIndex        =   15
         Top             =   2385
         Width           =   2550
      End
      Begin VB.ComboBox cbxViaP 
         Height          =   315
         ItemData        =   "frmColRecExpedRelacionados.frx":0010
         Left            =   900
         List            =   "frmColRecExpedRelacionados.frx":0017
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox txtMontoPet 
         Height          =   285
         Left            =   900
         TabIndex        =   14
         Top             =   2385
         Width           =   1695
      End
      Begin SICMACT.TxtBuscar AXCodJuzgado 
         Height          =   285
         Left            =   900
         TabIndex        =   9
         Top             =   870
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
      End
      Begin SICMACT.TxtBuscar AxCodJuez 
         Height          =   285
         Left            =   900
         TabIndex        =   10
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
      End
      Begin SICMACT.TxtBuscar AxCodSecre 
         Height          =   285
         Left            =   900
         TabIndex        =   11
         Top             =   1560
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Procesos:"
         Height          =   195
         Left            =   3000
         TabIndex        =   58
         Top             =   1980
         Width           =   705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Juez"
         Height          =   195
         Left            =   480
         TabIndex        =   38
         Top             =   1200
         Width           =   330
      End
      Begin VB.Label Label7 
         Caption         =   "Secretario"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Monto Petitorio"
         Height          =   435
         Left            =   120
         TabIndex        =   36
         Top             =   2280
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Juzgado"
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   870
         Width           =   600
      End
      Begin VB.Label Label13 
         Caption         =   "NroExpediente"
         Height          =   195
         Left            =   2640
         TabIndex        =   34
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Demanda"
         Height          =   405
         Left            =   120
         TabIndex        =   33
         Top             =   1860
         Width           =   735
      End
      Begin VB.Label lblNomJuzgado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2640
         TabIndex        =   32
         Top             =   870
         Width           =   3615
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNomJuez 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2640
         TabIndex        =   31
         Top             =   1200
         Width           =   3615
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNomSecre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2640
         TabIndex        =   30
         Top             =   1560
         Width           =   3615
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCliente 
      Height          =   1005
      Left            =   60
      TabIndex        =   19
      Top             =   570
      Width           =   8355
      Begin VB.Label Label2 
         Caption         =   "Cliente"
         Height          =   195
         Left            =   90
         TabIndex        =   28
         Top             =   270
         Width           =   645
      End
      Begin VB.Label Label3 
         Caption         =   "Prestamo"
         Height          =   195
         Left            =   90
         TabIndex        =   27
         Top             =   630
         Width           =   825
      End
      Begin VB.Label Label4 
         Caption         =   "Saldo Capital"
         Height          =   195
         Left            =   2640
         TabIndex        =   26
         Top             =   630
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ingreso Recup."
         Height          =   195
         Left            =   5280
         TabIndex        =   25
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label lblSaldoCapital 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3720
         TabIndex        =   24
         Top             =   600
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCodPers 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   960
         TabIndex        =   23
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblNomPers 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2400
         TabIndex        =   22
         Top             =   240
         Width           =   5475
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblMontoPrestamo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   960
         TabIndex        =   21
         Top             =   600
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblFecIngRecup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   6480
         TabIndex        =   20
         Top             =   600
         Width           =   1335
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
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
      Left            =   7080
      TabIndex        =   5
      Top             =   6270
      Width           =   1095
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
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   6270
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
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
      Left            =   4440
      TabIndex        =   3
      Top             =   6270
      Width           =   1095
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
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   6270
      Width           =   1095
   End
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
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   6270
      Width           =   1095
   End
   Begin SICMACT.ActXCodCta AXCodCta 
      Height          =   465
      Left            =   60
      TabIndex        =   18
      Top             =   120
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   820
      Texto           =   "Crédito"
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
End
Attribute VB_Name = "frmColRecExpedRelacionados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* EXPEDIENTE DE RECUPERACIONES
'Archivo:  frmColRecExpediente.frm
'LAYG   :  01/08/2001.
'Resumen:  Nos permite hacer el mantenimiento del expediente de Recuperaciones
Option Explicit

Dim fsEstudioJuridicoCod As String
Dim fsJuzgadoCod As String
Dim fsJuezCod As String
Dim fsSecretarioCod As String

Dim fbCambiaComision As Boolean
Dim fnComisionCodAnt As Integer
Dim fnComisionCodNue As Integer

Dim fsPetitorio As String
Dim fsHecho As String
Dim fsJuridico As String
Dim fsMProbatorios As String
Dim fsComplementos As String

Dim frsPers As ADODB.Recordset
Dim fsNuevoEdita As String
Dim fbExisteExpediente As Boolean

Function ValidaDatos() As Boolean
ValidaDatos = True
If txtNroExp = "" Then
    ValidaDatos = False
    MsgBox "Falta Ingresar el Expediente"
    txtNroExp.SetFocus
    Exit Function
End If
End Function

Private Sub AxCodAbogado_EmiteDatos()
    lblNomAbogado.Caption = AxCodAbogado.psDescripcion
End Sub


Private Sub BuscaDatos(ByVal psNroCredito As String)
Dim lbOk As Boolean
Dim lrValida As ADODB.Recordset
Dim loValCredito As COMNColocRec.NColRecValida
Dim lsmensaje As String
Dim lnNroExp As Integer
'On Error GoTo ControlError

    AXCodCta.NroCuenta = psNroCredito
    
    'Valida Contrato
    Set lrValida = New ADODB.Recordset
    Set loValCredito = New COMNColocRec.NColRecValida
        Set lrValida = loValCredito.nValidaExpedienteRelacionados(psNroCredito)
    
        If lrValida Is Nothing Then
            EstadoBotones False, False, True
        Else
        
            If lrValida!nTipCJ = gColRecTipCobJudicial Then
                Me.cbxTipoCob.ListIndex = 0
            Else
                Me.cbxTipoCob.ListIndex = 1
            End If
            
            If lrValida!nDemanda = gColRecDemandaNo Then
                Me.cbDemanda.ListIndex = 0
            Else
                Me.cbDemanda.ListIndex = 1
            End If
            
            lblCodPers.Caption = lrValida!cPersCod
            lblNomPers.Caption = lrValida!cPersNombre
            lblMontoPrestamo.Caption = Format(lrValida!nMontoCol, "#,##0.00")
            lblSaldoCapital.Caption = Format(lrValida!nSaldo, "#,##0.00")
            lblFecIngRecup.Caption = Format(lrValida!dIngRecup, "dd/mm/yyyy")
            
            EstadoBotones True, False, False
            AXCodCta.Enabled = False
            FraProceso.Enabled = False
        End If
        Set lrValida = Nothing
    Set loValCredito = Nothing
    
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & Err.Number & " " & Err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub




Private Sub AxCodJuez_EmiteDatos()
    lblNomJuez.Caption = AxCodJuez.psDescripcion
    If Len(lblNomJuez.Caption) > 0 Then
        AxCodSecre.SetFocus
    End If
End Sub

Private Sub AXCodJuzgado_EmiteDatos()
    lblNomJuzgado.Caption = AXCodJuzgado.psDescripcion
    If Len(lblNomJuzgado.Caption) > 0 Then
        AxCodJuez.SetFocus
    End If
End Sub

Private Sub AxCodSecre_EmiteDatos()
    lblNomSecre.Caption = AxCodSecre.psDescripcion
    If Len(lblNomSecre.Caption) > 0 Then
        cbxViaP.SetFocus
    End If
    
End Sub

Private Sub AxEncargado_EmiteDatos()
     LblEncargado.Caption = AxEncargado.psDescripcion
End Sub

Private Sub AxEncargado_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        Me.CboProcesal.SetFocus
    End If
End Sub

Private Sub cbDemanda_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    txtJuzgado.SetFocus
'End If
End Sub


Private Sub CboProcesal_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        Me.cmdPetitorio.SetFocus
     End If
End Sub

Private Sub CboProcesos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtMontoPet.SetFocus
    End If
End Sub

Private Sub cbxTipoCob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cbDemanda.SetFocus
End If
End Sub

Private Sub cmdBuscar_Click()

Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersCredito  As COMDColocRec.DCOMColRecCredito 'DColRecCredito
Dim lrCreditos As New ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona

On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
Set loPers = Nothing

' Selecciona Estados
lsEstados = gColocEstRecVigJud & "," & gColocEstRecVigCast

Limpiar ' True

EstadoBotones True, False, False

cbxTipoCob.ListIndex = -1
Me.fraDatos.Enabled = False
Me.fraEstudioJur.Enabled = False
AXCodCta.Enabled = True
    
If Trim(lsPersCod) <> "" Then
    Set loPersCredito = New COMDColocRec.DCOMColRecCredito
        Set lrCreditos = loPersCredito.dObtieneCreditosDePersona(lsPersCod, lsEstados)
    Set loPersCredito = Nothing
End If

Set loCuentas = New COMDPersona.UCOMProdPersona
    Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrCreditos)
    If loCuentas.sCtaCod <> "" Then
        AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        AXCodCta.SetFocusCuenta
    End If
Set loCuentas = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdCancelar_Click()
    Limpiar ' True
    EstadoBotones True, False, False
    EstadoControles False, False, False
    Me.TxtCorrelativo.Text = ""
    FraCorrelativo.Enabled = True
    fbExisteExpediente = False
    AXCodCta.Enabled = True
End Sub

'Private Sub cmdComi_Click()
'Dim loComision As UColRecComisionSelecciona
''lbCambiaComision = True
'Set loComision = New UColRecComisionSelecciona
'    Set loComision = frmColRecComisionSelecciona.Inicio(Me.AxCodAbogado.Text, Me.lblNomAbogado)
'    If loComision.nComisionCod > 0 Then
'        fnComisionCodNue = loComision.nComisionCod
'        'lblComision.Caption = loComision.nComisionValor
'        If loComision.nComisionTipo = 1 Then 'Moneda
'            lblComision.Caption = lblComision.Caption & " Mon"
'        Else
'            lblComision.Caption = lblComision.Caption & " % "
'        End If
'        AXCodJuzgado.SetFocus
'    End If
'Set loComision = Nothing
'End Sub

Private Sub cmdComplementos_Click()
Dim loMemo  As UColRecMemos
Set loMemo = New UColRecMemos
    Set loMemo = frmColRecMemos.Inicio("C", fsComplementos, AXCodCta.NroCuenta, Me.lblNomPers)
        fsComplementos = loMemo.sMemo
Set loMemo = Nothing
End Sub

Private Sub CmdEditar_Click()
    If Trim(TxtCorrelativo.Text) = "" Then
        MsgBox "Debe de Ingresar un Nro Correlativo", vbInformation, "Aviso"
        Exit Sub
    End If
    fbExisteExpediente = True
    EstadoBotones False, False, True
    EstadoControles True, True, True
    'FraCorrelativo.Enabled = True
    
    fsNuevoEdita = "E"
End Sub

Private Sub cmdGrabar_Click()

Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabar As COMNColocRec.NCOMColRecCredito
Dim lbCambiaComision As Boolean

Dim lsMovNro As String
Dim lsFechaHoraGrab As String
'On Error GoTo ControlError
Dim lnTipCobJud As Integer
Dim lnViaProce As Integer
Dim lnTipProceso As Integer
Dim lnTipProcesal As Integer
'validar numero de cuenta
If Len(AXCodCta.NroCuenta) <> 18 Then
    MsgBox "Ingrese un Nro. Cuenta Valido", vbInformation, "Aviso"
    Exit Sub
End If
' Valida Datos a Grabar
If Me.cbDemanda.ListIndex = 0 Then
    lnTipCobJud = gColRecTipCobJudicial
Else
    lnTipCobJud = gColRecTipCobExtraJudi
End If

'Via Procesal
If Len(Right(cbxViaP.Text, 2)) = 0 Then
    MsgBox "Seleccione un Tipo de Demanda", vbInformation, "Aviso"
    'cbxViaP.SetFocus
    Exit Sub
End If
lnViaProce = Right(cbxViaP.Text, 2)

CargaPersonasRelacionadas
lbCambiaComision = IIf(fnComisionCodNue <> fnComisionCodAnt, True, False)
If fValidaData = False Then
    Exit Sub
End If

If Len(Right(Trim(CboProcesos.Text), 2)) = 0 Then
    MsgBox "Seleccione Tipo Proceso", vbInformation, "Aviso"
    Exit Sub
End If
lnTipProceso = Right(Trim(CboProcesos.Text), 2)


If Len(Right(Trim(CboProcesal.Text), 2)) = 0 Then
    MsgBox "Seleccione Parte Procesal", vbInformation, "Aviso"
    Exit Sub
End If
lnTipProcesal = Right(Trim(CboProcesal.Text), 2)

If MsgBox(" Grabar Registro de Expediente ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
        
    'Genera el Mov Nro
    Set loContFunct = New COMNContabilidad.NCOMContFunciones
        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
    
    lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        
    Set loGrabar = New COMNColocRec.NCOMColRecCredito
        Call loGrabar.nRegistraExpedienteRecupRelacionados(AXCodCta.NroCuenta, lsMovNro, _
               fnComisionCodNue, Me.txtNroExp.Text, CCur(Val(Me.txtMontoPet.Text)), Mid(AXCodCta.NroCuenta, 9, 1), _
               frsPers, fsPetitorio, fsHecho, fsComplementos, fsMProbatorios, fsComplementos, lnViaProce, _
               0, lbCambiaComision, fbExisteExpediente, False, TxtCorrelativo.Text, lnTipProceso, lnTipProcesal)
               
    Set loGrabar = Nothing

    EstadoBotones True, False, False
    EstadoControles False, False, False
    FraCorrelativo.Enabled = True
    
Else
    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
End If
Exit Sub


ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Private Sub cmdHecho_Click()
Dim loMemo  As UColRecMemos
Set loMemo = New UColRecMemos
    Set loMemo = frmColRecMemos.Inicio("H", fsHecho, AXCodCta.NroCuenta, Me.lblNomPers)
        fsHecho = loMemo.sMemo
Set loMemo = Nothing
End Sub

Private Sub cmdJuridico_Click()
Dim loMemo  As UColRecMemos
Set loMemo = New UColRecMemos
    Set loMemo = frmColRecMemos.Inicio("J", fsJuridico, AXCodCta.NroCuenta, Me.lblNomPers)
        fsJuridico = loMemo.sMemo
Set loMemo = Nothing
End Sub

Private Sub cmdNuevo_Click()
Limpiar
fsNuevoEdita = "N"
fbExisteExpediente = False
Dim loValCredito As COMNColocRec.NColRecValida
Dim lnNroExp As Integer
Set loValCredito = New COMNColocRec.NColRecValida
    lnNroExp = loValCredito.ObtenerCorrelativoExped(Me.AXCodCta.NroCuenta)
    TxtCorrelativo.Text = lnNroExp
Set loValCredito = Nothing

EstadoBotones False, False, True
EstadoControles True, True, True
fbExisteExpediente = False
FraCorrelativo.Enabled = False
AXCodCta.Enabled = True
AXCodCta.SetFocusAge
End Sub

Private Sub cmdPetitorio_Click()
Dim loMemo  As UColRecMemos
Set loMemo = New UColRecMemos
    Set loMemo = frmColRecMemos.Inicio("P", fsPetitorio, AXCodCta.NroCuenta, Me.lblNomPers)
        fsPetitorio = loMemo.sMemo
Set loMemo = Nothing
End Sub

Private Sub cmdProbatorios_Click()
Dim loMemo  As UColRecMemos
Set loMemo = New UColRecMemos
    Set loMemo = frmColRecMemos.Inicio("M", fsMProbatorios, AXCodCta.NroCuenta, Me.lblNomPers)
        fsMProbatorios = loMemo.sMemo
Set loMemo = Nothing
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Limpiar()
txtNroExp = ""
AxCodAbogado.Text = "": lblNomAbogado = ""
AXCodJuzgado.Text = "": lblNomJuzgado = ""
AxCodJuez.Text = "": lblNomJuez = ""
AxCodSecre.Text = "": lblNomSecre = ""
AxEncargado.Text = "": LblEncargado = ""
fsPetitorio = ""
fsHecho = ""
fsJuridico = ""
fsMProbatorios = ""
fsComplementos = ""

cbxViaP.ListIndex = -1
CboProcesos.ListIndex = -1
CboProcesal.ListIndex = -1
'cbDemanda.ListIndex = -1
chkComplementos.value = 0
chkFundHecho.value = 0
chkFundJuridico.value = 0
chkMedProb.value = 0
chkPetitorio.value = 0
txtMontoPet.Text = ""

End Sub
Private Sub Form_Activate()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    Me.chkPetitorio.value = IIf(Trim(fsPetitorio) = "", 0, 1)
    Me.chkFundHecho.value = IIf(Trim(fsHecho) = "", 0, 1)
    Me.chkFundJuridico.value = IIf(Trim(fsJuridico) = "", 0, 1)
    Me.chkMedProb.value = IIf(Trim(fsMProbatorios) = "", 0, 1)
    Me.chkComplementos.value = IIf(Trim(fsComplementos) = "", 0, 1)
    
End Sub

Private Sub Form_Load()
   'BuscarDatos
   Me.AXCodCta.NroCuenta = fgIniciaAxCuentaRecuperaciones
   'CargaForma Me.cbxViaP, "80", pConexionJud
   CargaViaProcesal
   CargaComboDatos cbxTipoCob, gColocRecTipoCobranza
   CargaComboDatos cbDemanda, gColocRecDemandado
   CargaComboDatos CboProcesos, gColocRecTipoProceso
   CargaComboDatos CboProcesal, gColocRecTipoProcesal
   Me.AXCodCta.EnabledCMAC = False
   Me.AXCodCta.EnabledAge = False
   Me.AXCodCta.EnabledProd = False
   Me.AXCodCta.EnabledCta = False
   FraCorrelativo.Enabled = True
   
End Sub


Private Sub TxtCorrelativo_KeyPress(KeyAscii As Integer)
    'funcion que busque Expediente Correlativo
   Dim loValCredito As COMNColocRec.NColRecValida
   Dim lrValida As ADODB.Recordset
   Dim lsmensaje As String
   Dim lnTipoProceso As String
   Dim lnTipoProcesal As String
   
   If KeyAscii = 13 Then
        Set lrValida = New ADODB.Recordset
        Set loValCredito = New COMNColocRec.NColRecValida
            Set lrValida = loValCredito.nValidaExpedientePersona(Me.AXCodCta.NroCuenta, TxtCorrelativo.Text, lsmensaje)
        Set loValCredito = Nothing
        
        If Trim(lsmensaje) <> "" Then
             MsgBox lsmensaje, vbInformation, "Aviso"
             Exit Sub
        End If
        
        If Not (lrValida.EOF And lrValida.BOF) Then
            fbExisteExpediente = True
            EstadoBotones False, True, False
            AxCodAbogado.Text = lrValida!cCodAbog
            lblNomAbogado.Caption = lrValida!cNomAbog
            Me.AXCodJuzgado.Text = lrValida!cCodJuzg
            Me.lblNomJuzgado.Caption = lrValida!cNomJuzg
            Me.AxCodJuez.Text = lrValida!cCodJuez
            Me.lblNomJuez.Caption = lrValida!cNomJuez
            Me.AxCodSecre.Text = lrValida!cCodSecre
            Me.lblNomSecre.Caption = lrValida!cNomSecre
            Me.AxEncargado.Text = lrValida!cCodEncargado
            Me.LblEncargado.Caption = lrValida!cNomEncargado
            Me.txtMontoPet.Text = Format(lrValida!nMonPetit, "#0.00")
            Me.txtNroExp.Text = lrValida!cNumExp
            
            Call UbicaCombo(Me.cbxViaP, lrValida!nViaProce, True)
            
            lnTipoProceso = lrValida!TipoProceso
            Call UbicaCombo(Me.CboProcesos, lnTipoProceso, True)
            
            lnTipoProcesal = lrValida!TipoProcesal
            Call UbicaCombo(Me.CboProcesal, lnTipoProcesal, True)


            fsPetitorio = IIf(Trim(lrValida!mPetit) <> "", lrValida!mPetit, "")
            fsHecho = IIf(Trim(lrValida!mHechos) <> "", lrValida!mHechos, "")
            fsJuridico = IIf(Trim(lrValida!mFundJur) <> "", lrValida!mFundJur, "")
            fsMProbatorios = IIf(Trim(lrValida!mMedProb) <> "", lrValida!mMedProb, "")
            fsComplementos = IIf(Trim(lrValida!mDatComp) <> "", lrValida!mDatComp, "")
            
        Else
            MsgBox "No existe este Numero de Expediente", vbInformation, "Aviso"
            EstadoBotones True, False, False
            Limpiar
        End If
   End If
End Sub

Private Sub txtMontoPet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     If CDbl(txtMontoPet.Text) > 0 Then
        Me.txtNroExp.SetFocus
      Else
        MsgBox "El Monto debe ser Mayor a 0", vbInformation, "Aviso"
        Exit Sub
      End If
    End If
End Sub

Private Sub txtNroExp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.AxEncargado.SetFocus
End Sub

Private Sub CargaPersonasRelacionadas()
Set frsPers = New ADODB.Recordset
    frsPers.Fields.Append "cPersCod", adVarChar, 13
    frsPers.Fields.Append "cPersNombre", adVarChar, 50
    frsPers.Fields.Append "cPersRelac", adVarChar, 5
    
    frsPers.Open
    If Len(Trim(Me.AxCodAbogado.Text)) > 0 Then
        frsPers.AddNew
        frsPers.Fields("cPersCod") = Me.AxCodAbogado.Text
        frsPers.Fields("cPersNombre") = Trim(Me.lblNomAbogado.Caption)
        frsPers.Fields("cPersRelac") = gColRelPersEstudioJuridico
        frsPers.Update
    End If
    If Len(Trim(Me.AXCodJuzgado.Text)) > 0 Then
        frsPers.AddNew
        frsPers.Fields("cPersCod") = Me.AXCodJuzgado.Text
        frsPers.Fields("cPersNombre") = Trim(Me.lblNomJuzgado.Caption)
        frsPers.Fields("cPersRelac") = gColRelPersJuzgado
        frsPers.Update
    End If
    If Len(Trim(Me.AxCodJuez.Text)) > 0 Then
        frsPers.AddNew
        frsPers.Fields("cPersCod") = Me.AxCodJuez.Text
        frsPers.Fields("cPersNombre") = Trim(Me.lblNomJuez.Caption)
        frsPers.Fields("cPersRelac") = gColRelPersJuez
        frsPers.Update
    End If
    If Len(Trim(Me.AxCodSecre.Text)) > 0 Then
        frsPers.AddNew
        frsPers.Fields("cPersCod") = Me.AxCodSecre.Text
        frsPers.Fields("cPersNombre") = Trim(Me.lblNomSecre.Caption)
        frsPers.Fields("cPersRelac") = gColRelPersSecretario
        frsPers.Update
    End If
    If Len(Trim(Me.AxEncargado.Text)) > 0 Then
        frsPers.AddNew
        frsPers.Fields("cPersCod") = Me.AxEncargado.Text
        frsPers.Fields("cPersNombre") = Trim(Me.LblEncargado.Caption)
        frsPers.Fields("cPersRelac") = gColRelPersEncargado
        frsPers.Update
    End If

End Sub

Private Sub CargaViaProcesal()
Dim oDatos As COMDConstSistema.DCOMGeneral
Dim rs As New ADODB.Recordset
cbxViaP.Clear
Set oDatos = New COMDConstSistema.DCOMGeneral
Set rs = oDatos.GetConstante(gColocRecViaProcesal)
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
            cbxViaP.AddItem rs!cDescripcion + Space(70) + CStr(rs!nConsValor)
            rs.MoveNext
        Loop
    End If
Set rs = Nothing
Set oDatos = Nothing
End Sub

Private Function fValidaData() As Boolean
Dim lbOk As Boolean
lbOk = True

If Len(Trim(Me.AxCodAbogado.Text)) > 0 Then
    If Me.AxCodAbogado.Text = Me.AXCodJuzgado Or _
       Me.AxCodAbogado.Text = Me.AxCodJuez Or _
       Me.AxCodAbogado.Text = Me.AxCodSecre Then
        MsgBox "Persona no puede ejercer varios roles al mismo tiempo", vbInformation, "Aviso"
        lbOk = False
        Exit Function
    End If
End If
If Len(Trim(Me.AXCodJuzgado.Text)) > 0 Then
    If Me.AXCodJuzgado.Text = Me.AxCodJuez Or _
       Me.AXCodJuzgado.Text = Me.AxCodSecre Then
        MsgBox "Persona no puede ejercer varios roles al mismo tiempo", vbInformation, "Aviso"
        lbOk = False
        Exit Function
    End If
End If
If Len(Trim(Me.AxCodJuez.Text)) > 0 Then
    If Me.AxCodJuez.Text = Me.AxCodSecre Then
        MsgBox "Persona no puede ejercer varios roles al mismo tiempo", vbInformation, "Aviso"
        lbOk = False
        Exit Function
    End If
End If
If frsPers.RecordCount = 0 Then
    MsgBox "Ingrese un Abogado, Juez, Juzgado", vbInformation, "Aviso"
    lbOk = False
    Exit Function
End If
fValidaData = lbOk
End Function

Private Sub CargaComboDatos(ByVal combo As ComboBox, ByVal pnValor As Integer)
    Dim oConst As COMDConstantes.DCOMConstantes
    Dim rs As New ADODB.Recordset
    Set oConst = New COMDConstantes.DCOMConstantes
        Set rs = oConst.ObtenerVarRecuperaciones(pnValor)
        combo.Clear
        If Not (rs.EOF And rs.BOF) Then
            Do Until rs.EOF
                combo.AddItem rs(0)
                rs.MoveNext
            Loop
        End If
    Set oConst = Nothing
    Set rs = Nothing
End Sub

Public Sub Inicia(ByVal psctacod As String)
    BuscaDatos psctacod
    Me.Show 1
End Sub

Public Sub EstadoBotones(ByVal bnuevo As Boolean, ByVal beditar As Boolean, ByVal bgrabar As Boolean)
    cmdNuevo.Enabled = bnuevo
    cmdEditar.Enabled = beditar
    cmdGrabar.Enabled = bgrabar
End Sub

Public Sub EstadoControles(ByVal bDatos As Boolean, ByVal bEstudio As Boolean, ByVal bInforme As Boolean)
    Me.fraDatos.Enabled = bDatos
    Me.fraEstudioJur.Enabled = bEstudio
    Me.fraInforme.Enabled = bInforme
End Sub


