VERSION 5.00
Begin VB.Form frmCredReprogPropuesta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Propuesta de Reprogramación"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   Icon            =   "frmCredReprogPropuesta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frRechazoOtros 
      Caption         =   "Rechazo y Otros"
      Height          =   735
      Left            =   120
      TabIndex        =   60
      Top             =   1920
      Width           =   7455
      Begin VB.TextBox txtRechazoOtrosDias 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4560
         MaxLength       =   5
         TabIndex        =   62
         Text            =   "0"
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbRechazoOtros 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label lblRechazoOtros_Dias 
         Caption         =   "Dias "
         Height          =   255
         Left            =   4080
         TabIndex        =   63
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame fr_Modalidades 
      Caption         =   "Falicidades Crediticias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   3600
      TabIndex        =   58
      Top             =   6960
      Width           =   3975
      Begin SICMACT.FlexEdit fe_Modalidades 
         Height          =   1575
         Left            =   120
         TabIndex        =   59
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   2778
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "-Modalidad--CodMod-Calif.-Cuota-aux"
         EncabezadosAnchos=   "0-1400-400-0-500-1200-0"
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-X-X-X-X"
         ListaControles  =   "0-0-4-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-R-C-R-L"
         FormatosEdit    =   "0-1-0-3-3-2-0"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.CommandButton cmdSimulador 
      Caption         =   "Simulador"
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
      Left            =   1440
      TabIndex        =   57
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton cmdSolicitud 
      Caption         =   "Sol. Web"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5200
      TabIndex        =   56
      Top             =   120
      Width           =   930
   End
   Begin VB.Frame fr_VerActEco 
      Caption         =   "Verificación de la Actividad Económica"
      Height          =   4215
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   7455
      Begin VB.TextBox txt_NewActividad 
         Height          =   285
         Left            =   1920
         TabIndex        =   55
         Top             =   1320
         Width           =   5415
      End
      Begin VB.TextBox txt_Actividad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   54
         Top             =   240
         Width           =   6375
      End
      Begin VB.CommandButton cmdCPIFIs 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   52
         Top             =   3840
         Width           =   495
      End
      Begin VB.Frame fr_IngActPert 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         TabIndex        =   46
         Top             =   2760
         Width           =   6135
         Begin VB.OptionButton OpIngActPer 
            Caption         =   "Incrementar Cuota"
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   49
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton OpIngActPer 
            Caption         =   "Mantener"
            Height          =   255
            Index           =   1
            Left            =   3240
            TabIndex        =   48
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton OpIngActPer 
            Caption         =   "Disminuir"
            Height          =   255
            Index           =   2
            Left            =   4560
            TabIndex        =   47
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label12 
            Caption         =   "Ingresos Actuales Permiten:"
            Height          =   375
            Left            =   120
            TabIndex        =   50
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.Frame fr_DismPorc 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   2400
         Width           =   6135
         Begin VB.OptionButton OpDismPorc 
            Caption         =   "No hay Negocio"
            Height          =   255
            Index           =   2
            Left            =   4560
            TabIndex        =   45
            Top             =   120
            Width           =   1575
         End
         Begin VB.OptionButton OpDismPorc 
            Caption         =   "Hasta el 50%"
            Height          =   255
            Index           =   1
            Left            =   3240
            TabIndex        =   43
            Top             =   120
            Width           =   1335
         End
         Begin VB.OptionButton OpDismPorc 
            Caption         =   "Hasta 25%"
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   42
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Disminución %:"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.Frame fr_DismiEs 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   2040
         Width           =   6135
         Begin VB.OptionButton OPDimicEs 
            Caption         =   "Temporal"
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   39
            Top             =   120
            Width           =   1335
         End
         Begin VB.OptionButton OPDimicEs 
            Caption         =   "Permanente"
            Height          =   255
            Index           =   1
            Left            =   3240
            TabIndex        =   38
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label17 
            Caption         =   "La disminución es:"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.Frame fr_IngAct 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   1680
         Width           =   6135
         Begin VB.OptionButton OpIngrAct 
            Caption         =   "Mantiene"
            Height          =   255
            Index           =   1
            Left            =   3240
            TabIndex        =   36
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton OpIngrAct 
            Caption         =   "Disminuyeron"
            Height          =   255
            Index           =   2
            Left            =   4560
            TabIndex        =   35
            Top             =   120
            Width           =   1335
         End
         Begin VB.OptionButton OpIngrAct 
            Caption         =   "Incrementaron"
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   34
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "Ingresos Actuales:"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   120
            Width           =   1335
         End
      End
      Begin SICMACT.EditMoney EditMoneyIngrVentNet 
         Height          =   255
         Left            =   1680
         TabIndex        =   20
         Top             =   3480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney EditMoneyExce 
         Height          =   255
         Left            =   1680
         TabIndex        =   21
         Top             =   3840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney EditMoneyGastoTot 
         Height          =   255
         Left            =   4440
         TabIndex        =   23
         Top             =   3480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.Frame fr_NegViFunc 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   540
         Width           =   4335
         Begin VB.OptionButton OpNF 
            Caption         =   "Si"
            Height          =   255
            Index           =   1
            Left            =   2880
            TabIndex        =   26
            Top             =   120
            Width           =   495
         End
         Begin VB.OptionButton OpNF 
            Caption         =   "No"
            Height          =   255
            Index           =   0
            Left            =   3480
            TabIndex        =   25
            Top             =   120
            Width           =   615
         End
         Begin VB.Label lblNegVieFinc 
            Caption         =   "Negocio viene funcionando:"
            Height          =   255
            Left            =   0
            TabIndex        =   27
            Top             =   120
            Width           =   2655
         End
      End
      Begin VB.Frame fr_CliMantAct 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   840
         Width           =   4215
         Begin VB.OptionButton OpCA 
            Caption         =   "No"
            Height          =   255
            Index           =   0
            Left            =   3480
            TabIndex        =   31
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton OpCA 
            Caption         =   "Si"
            Height          =   255
            Index           =   1
            Left            =   2880
            TabIndex        =   30
            Top             =   120
            Width           =   495
         End
         Begin VB.Label lblClieMantAct 
            Caption         =   "Cliente mantiene actividad:"
            Height          =   255
            Left            =   0
            TabIndex        =   29
            Top             =   120
            Width           =   2655
         End
      End
      Begin SICMACT.EditMoney EditMoneyCP 
         Height          =   255
         Left            =   5160
         TabIndex        =   53
         Top             =   3840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin VB.Label Label11 
         Caption         =   "CP (ctas IF/EXEC):"
         Height          =   255
         Left            =   3120
         TabIndex        =   51
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "Gastos Totales:"
         Height          =   255
         Left            =   3120
         TabIndex        =   22
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Excedente:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Ingresos o Ventas Netas:"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label lblIndNewAct 
         Caption         =   "Indica nueva actividad:"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Actividad:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   280
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6360
      TabIndex        =   14
      Top             =   9000
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   13
      Top             =   9000
      Width           =   1170
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5040
      TabIndex        =   11
      Top             =   9000
      Width           =   1170
   End
   Begin VB.Frame FraComent 
      Caption         =   "Comentarios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   9
      Top             =   6960
      Width           =   3375
      Begin VB.TextBox txtComent 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3960
      TabIndex        =   7
      Top             =   120
      Width           =   930
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   7455
      Begin VB.Label lblCodigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   960
         TabIndex        =   6
         Top             =   330
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Código :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   370
         Width           =   735
      End
      Begin VB.Label lblDOI 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4320
         TabIndex        =   4
         Top             =   330
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "D.O.I. :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   3
         Top             =   375
         Width           =   615
      End
      Begin VB.Label lblNomCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   960
         TabIndex        =   2
         Top             =   765
         Width           =   6375
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   810
         Width           =   735
      End
   End
   Begin SICMACT.ActXCodCta ActXCodCta 
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      Texto           =   "Credito"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
   Begin VB.CommandButton cmdRechazar 
      Caption         =   "Rechazar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4920
      TabIndex        =   12
      Top             =   9000
      Width           =   1290
   End
End
Attribute VB_Name = "frmCredReprogPropuesta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************************************
'** Nombre : frmCredReprogSolicitud
'** Descripción : Formulario para generar la propuesta de reprogramación de créditos según TI-ERS010-2016
'** Creación : JUEZ, 20160311 09:00:00 AM
'*********************************************************************************************************

Option Explicit

Private Enum TipoAcceso
    TipoPropuesto = 0
    TipoRechazo = 1
    TipoGarantia = 2 'Add JOEP20210306 garantia covid
End Enum

Dim oDCred As COMDCredito.DCOMCredito
Dim fnTipo As TipoAcceso
Dim fnCuotasReprog As Integer
Dim fdFecCuotaVenc As Date
Dim fdFecNuevaCuotaVenc As Date
Dim fnSaldoReprog As Double
Dim fbSolicAutorizacion As Boolean
Dim sMovNro As String
Dim cTpProducto As String 'JOEP20200804 Actividad Economica covid
Dim m As Variant 'JOEP20200805
Dim nMOntoCPIfis As Currency

'Add JOEP20210306 garantia covid
Dim gnDiasAtraso As Integer
Dim gMatrixCalend As Variant
Dim gMatrixDatos As Variant
'Add JOEP20210306 garantia covid

Public Sub Inicio(ByVal pnTipo As Integer)
    cTpProducto = "" 'JOEP20200804 Actividad Economica covid
    fnTipo = pnTipo
    ActXCodCta.CMAC = gsCodCMAC
    ActXCodCta.Age = gsCodAge
    
    ValidarFechaActual
    
    Set gMatrixCalend = Nothing 'Add JOEP20210306 garantia covid
    
    If fnTipo = TipoPropuesto Then
        Caption = "Propuesta de Reprogramación"
        fr_Modalidades.Caption = "Falicidades Crediticias" 'Add JOEP20210306 garantia covid
        FraComent.Caption = " Comentarios "
        cmdGenerar.Visible = True
        cmdRechazar.Visible = False
        'JHCU 2609202 Solicitud Web
        cmdSolicitud.Visible = True
'JOEP20200804 Actividad Economica covid
        fr_VerActEco.Visible = True
        frRechazoOtros.Visible = False 'reversion
        Call MedidasFormulario(fnTipo)
'JOEP20200804 Actividad Economica covid
    ElseIf fnTipo = TipoRechazo Then
        Caption = "Rechazo de Reprogramación"
        FraComent.Caption = " Motivo Rechazo "
        cmdGenerar.Visible = False
        cmdRechazar.Visible = True
'JOEP20200804 Actividad Economica covid
        fr_VerActEco.Visible = False
        cmdSolicitud.Visible = False 'JHCU 2609202 Solicitud Web
        cmdSimulador.Visible = False
        txtRechazoOtrosDias.Visible = False
        lblRechazoOtros_Dias.Visible = False
        Call MedidasFormulario(fnTipo)
'JOEP20200804 Actividad Economica covid
'Add JOEP20210306 garantia covid
    ElseIf fnTipo = TipoGarantia Then
        Caption = "Programa Garantías COVID 19 - Ley 31050"
        fr_Modalidades.Caption = "Beneficios de Programa"
        FraComent.Caption = " Comentarios "
        cmdGenerar.Visible = True
        cmdRechazar.Visible = False
        cmdSolicitud.Visible = True
        fr_VerActEco.Visible = True
        frRechazoOtros.Visible = False 'reversion
        cmdBuscar.Enabled = False
        Call MedidasFormulario(fnTipo)
'Add JOEP20210306 garantia covid
    End If
    Me.Show 1
End Sub

Private Sub ValidarFechaActual()
Dim lsFechaValidador As String
lsFechaValidador = validarFechaSistema
    If lsFechaValidador <> "" Then
        If gdFecSis <> CDate(lsFechaValidador) Then
            MsgBox "La Fecha de tu sesión en el Negocio no coincide con la fecha del Sistema", vbCritical, "Aviso"
            Unload Me
            End
        End If
    End If
End Sub

Private Sub cmdBuscar_Click()
Dim R As ADODB.Recordset
Dim oDPers As COMDPersona.UCOMPersona
    Limpiar
    Set oDPers = frmBuscaPersona.Inicio()
    If Not oDPers Is Nothing Then
        Call FrmVerCredito.Inicio(oDPers.sPersCod, , , True, ActXCodCta)
        ActXCodCta.SetFocusCuenta
    End If
    Set oDPers = Nothing
End Sub

Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(ActXCodCta.NroCuenta) = 18 Then
            If CargaDatos Then
                HabilitaControles False
                If txtComent.Enabled And FraComent.Enabled Then txtComent.SetFocus
            Else
                HabilitaControles True
            End If
        Else
            MsgBox "Ingrese correctamente el crédito", vbInformation, "Aviso"
        End If
    End If
End Sub

Private Function CargaDatos()
Dim rsCred As ADODB.Recordset, rs As ADODB.Recordset, rsvalida As ADODB.Recordset 'Add rsvalida JOEP20210306 garantia covid
Dim lsEstadosValida As String
Dim i As Integer

CargaDatos = False
cTpProducto = "" 'JOEP20200804 Actividad Economica covid
        
Set oDCred = New COMDCredito.DCOMCredito
Set rsCred = oDCred.RecuperaDatosCreditoVigente(ActXCodCta.NroCuenta, gdFecSis, 0)
Set oDCred = Nothing

'RIRO20200911 VALIDA LIQUIDACION ***************
Dim oCreditoTmp As COMNCredito.NCOMCredito
Dim bValidaActualizacionLiq As Boolean
Set oCreditoTmp = New COMNCredito.NCOMCredito
bValidaActualizacionLiq = oCreditoTmp.VerificaActualizacionLiquidacion(ActXCodCta.NroCuenta)
If Not bValidaActualizacionLiq Then
    MsgBox "El crédito no tiene actualizados sus datos de liquidación, no podrá realizar " & _
    "la reprogramación mientras no se actualicen estos datos. Deberá comunicarse con el área de T.I.", vbExclamation, "Aviso"
    bValidaActualizacionLiq = False
    Exit Function
End If

'END RIRO **************************************

'JOEP20210211 Garantia covid
Set oDCred = New COMDCredito.DCOMCredito
Set rsvalida = oDCred.ReprogramacionPropuestaMsgBox(ActXCodCta.NroCuenta, fnTipo)
If Not (rsCred.EOF And rsCred.BOF) Then
    If rsvalida!cMgsBox <> "" Then
        MsgBox rsvalida!cMgsBox, vbInformation, "Aviso"
        Exit Function
    End If
End If
Set oDCred = Nothing
'JOEP20210211 Garantia covid
    
If Not rsCred.EOF And Not rsCred.BOF Then
    Set oDCred = New COMDCredito.DCOMCredito
    'Set rs = oDCred.RecuperaColocacReprogramado(ActXCodCta.NroCuenta, gEstReprogSolicitado & "," & gEstReprogAutorizado & "," & gEstReprogPropuesto & "," & gEstReprogAprobado & "," & gEstReprogSolicitadoReprogramado) 'JOEP20180509 Acta093-2018 gEstReprogSolicitadoReprogramado
    Set rs = oDCred.RecuperaColocacReprogramado(ActXCodCta.NroCuenta, gEstReprogSolicitado & "," & gEstReprogAutorizado & "," & gEstReprogPropuesto & "," & gEstReprogAprobado & "," & gEstReprogSolicitadoReprogramado, fnTipo) 'Add JOEP20210306 garantia covid
    Set oDCred = Nothing
        
    If fnTipo = TipoPropuesto Then
        If Not rs.EOF And Not rs.BOF Then
            
            Call ValidaSolOcmSicm
            
            If rs!nPrdEstado = gEstReprogSolicitado Then
                lblCodigo.Caption = rsCred!cPersCod
                lblDOI.Caption = Trim(rsCred!nDoi)
                lblNomCliente.Caption = rsCred!cPersNombre
                fnCuotasReprog = CInt(rs!nCuotasReprog)
                fdFecCuotaVenc = CDate(rs!dFecCuotaVenc)
                fdFecNuevaCuotaVenc = CDate(rs!dFecNuevaCuotaVenc)
                fnSaldoReprog = rs!nSaldoCap
                fbSolicAutorizacion = rs!bSolicAutorizacion
                gnDiasAtraso = rs!nDiasAtraso 'Add JOEP20210306 garantia covid
                    
            'JOEP20200804 Actividad Economica covid
                txt_Actividad.Text = rs!ActiFormEval
                EditMoneyIngrVentNet.Text = Format(rs!IngVentasNetasFormEval, "#,##0.00")
                EditMoneyGastoTot.Text = Format(rs!GastosFormEval, "#,##0.00")
                EditMoneyExce.Text = Format(rs!ExcedenteFormEval, "#,##0.00")
                                                    
                HabilitaControles True, 1
                       
                cTpProducto = rs!cTpProducto
                If Left(rs!cTpProducto, 1) = 7 Then
                    HabilitaControles False, 2
                Else
                    HabilitaControles True, 2
                End If
                    
                If Left(rs!cTpProducto, 1) = 7 Then
                    lblNegVieFinc = "Actividad/Ingreso/Trabajo Actual"
                    lblClieMantAct = "¿Continua trabajando?"
                    lblIndNewAct = "Indicar Nueva Actividad o Actual Empleo "
                    fr_NegViFunc.Visible = False
                Else
                    lblNegVieFinc = "Negocio viene funcionando:"
                    lblClieMantAct = "Cliente mantiene actividad:"
                    lblIndNewAct = "Indica nueva actividad:"
                    fr_NegViFunc.Visible = True
                End If
                 
                Call CargaGrillaModalidades(2090, 1, fnTipo) 'Add JOEP20210306 garantia covid
                
'                Dim oDCOMCred As COMDConstantes.DCOMConstantes
'                Dim rsCovidOpciones As ADODB.Recordset
'                Set oDCOMCred = New COMDConstantes.DCOMConstantes
'                    Set rsCovidOpciones = oDCOMCred.RecuperaConstantes(2090)
'                    If Not (rsCovidOpciones.BOF And rsCovidOpciones.EOF) Then
'                        LimpiaFlex fe_Modalidades
'                        For i = 1 To rsCovidOpciones.RecordCount
'                            fe_Modalidades.AdicionaFila
'                            fe_Modalidades.TextMatrix(i, 0) = i
'                            fe_Modalidades.TextMatrix(i, 1) = rsCovidOpciones!cConsDescripcion
'                            fe_Modalidades.TextMatrix(i, 3) = rsCovidOpciones!nConsValor
'                            rsCovidOpciones.MoveNext
'                        Next i
'                    End If
'                Set oDCOMCred = Nothing
'                RSClose rsCovidOpciones
'            'JOEP20200804 Actividad Economica covid
            
                    CargaDatos = True
                    
                ElseIf rs!nPrdEstado = gEstReprogAutorizado Or rs!nPrdEstado = gEstReprogPropuesto Or rs!nPrdEstado = gEstReprogAprobado Then
                    MsgBox "El crédito no está disponible para ser propuesto. Su estado actual es " & rs!cPrdEstado, vbInformation, "Aviso"
                    Limpiar
                    Exit Function
                Else
                'INICIO Agrego JOEP20171214 ACTA220-2017
                    If rs!nPrdEstado = 208 Then
                        'MsgBox "El crédito ya se encuentra en proceso de reprogramación. Su estado actual es " & rs!cPrdEstado & ", No es necesario seguir con los otros procesos, por favor de ir a la Opción de Reprogramación.", vbInformation, "Aviso"'Comentado JOEP20171214 ERS082-2017
                        MsgBox "El crédito ya se encuentra en proceso de reprogramación. No es necesario seguir con los otros procesos, por favor ir a la Opción de Reprogramación.", vbInformation, "Aviso" 'Agrego JOEP20171214 ERS082-2017
                    Else
                'INICIO Agrego JOEP20171214 ACTA220-2017
                        MsgBox "El crédito no está disponible para ser propuesto", vbInformation, "Aviso"
                    End If
                    Limpiar
                    Exit Function
                End If
            Else
                MsgBox "El crédito no está disponible para ser propuesto", vbInformation, "Aviso"
                Limpiar
                Exit Function
            End If
    ElseIf fnTipo = TipoRechazo Then
            If Not rs.EOF And Not rs.BOF Then
                lblCodigo.Caption = rsCred!cPersCod
                lblDOI.Caption = Trim(rsCred!nDoi)
                lblNomCliente.Caption = rsCred!cPersNombre
                fnSaldoReprog = rs!nSaldoCap
                gnDiasAtraso = rs!nDiasAtraso 'Add JOEP20210306 garantia covid
                'revesion
                Dim onRev As COMDCredito.DCOMCredito
                Dim rsRev As ADODB.Recordset
                Set onRev = New COMDCredito.DCOMCredito
                Set rsRev = onRev.ReprogRechazoLlenaCombo(ActXCodCta.NroCuenta, 2091)
                    If Not (rsRev.BOF And rsRev.EOF) Then
                        Call Llenar_Combo_con_Recordset(rsRev, cmbRechazoOtros)
                        Call CambiaTamañoCombo(cmbRechazoOtros, 200)
                    End If
                Set onRev = Nothing
                RSClose rsRev
                'revesion
                CargaDatos = True
            Else
                MsgBox "El crédito no puede ser rechazado porque no está en proceso de reprogramación", vbInformation, "Aviso"
                Limpiar
                Exit Function
            End If
    ElseIf fnTipo = TipoGarantia Then
        If rs!nPrdEstado = gEstReprogSolicitado Then
            lblCodigo.Caption = rsCred!cPersCod
            lblDOI.Caption = Trim(rsCred!nDoi)

            If Not rs.EOF And Not rs.BOF Then
                lblNomCliente.Caption = rsCred!cPersNombre
                fnCuotasReprog = CInt(rs!nCuotasReprog)
                fdFecCuotaVenc = CDate(rs!dFecCuotaVenc)
                fdFecNuevaCuotaVenc = CDate(rs!dFecNuevaCuotaVenc)
                fnSaldoReprog = rs!nSaldoCap
                gnDiasAtraso = rs!nDiasAtraso
                    If rs!bActEco = 1 Then
                        txt_Actividad.Text = rs!ActiFormEval
                        If rs!nNVF <> -1 Then OpNF.iTem(rs!nNVF).value = 1
                        OpCA.iTem(rs!nCMA).value = 1
                        txt_NewActividad.Text = rs!cNewAct

                        OpIngrAct.iTem(rs!nIA - 1).value = 1
                        If rs!nIA = 3 Then
                            OPDimicEs.iTem(rs!nDE - 1).value = 1
                        End If
                        OpDismPorc.iTem(rs!nDPorc - 1).value = 1
                        OpIngActPer.iTem(rs!nIAP - 1).value = 1

                        EditMoneyIngrVentNet.Text = Format(rs!IngVentasNetasFormEval, "#,##0.00")
                        EditMoneyGastoTot.Text = Format(rs!GastosFormEval, "#,##0.00")
                        EditMoneyExce.Text = Format(rs!ExcedenteFormEval, "#,##0.00")
                        EditMoneyCP.Text = Format(rs!Cp, "#,##0.00")
                    End If

                    HabilitaControles True, 1
                    cTpProducto = rs!cTpProducto

                    If Left(rs!cTpProducto, 1) = 7 Then
                        HabilitaControles False, 2
                    Else
                        HabilitaControles True, 2
                    End If

                    If Left(rs!cTpProducto, 1) = 7 Then
                        lblNegVieFinc = "Actividad/Ingreso/Trabajo Actual"
                        lblClieMantAct = "¿Continua trabajando?"
                        lblIndNewAct = "Indicar Nueva Actividad o Actual Empleo "
                        fr_NegViFunc.Visible = False
                    Else
                        lblNegVieFinc = "Negocio viene funcionando:"
                        lblClieMantAct = "Cliente mantiene actividad:"
                        lblIndNewAct = "Indica nueva actividad:"
                        fr_NegViFunc.Visible = True
                    End If

                    txtComent.Text = rs!cMotivo
            End If

            Call CargaGrillaModalidades(2090, 0, fnTipo)
            Call HabilitaControles(False, 3)

            CargaDatos = True

        ElseIf rs!nPrdEstado = gEstReprogAutorizado Or rs!nPrdEstado = gEstReprogPropuesto Or rs!nPrdEstado = gEstReprogAprobado Then
            MsgBox "El crédito no está disponible para ser propuesto. Su estado actual es " & rs!cPrdEstado, vbInformation, "Aviso"
            Limpiar
            Exit Function
        Else
            If rs!nPrdEstado = 208 Then
                MsgBox "El crédito ya se encuentra en proceso de reprogramación. No es necesario seguir con los otros procesos, por favor ir a la Opción de Reprogramación.", vbInformation, "Aviso" 'Agrego JOEP20171214 ERS082-2017
            Else
                MsgBox "El crédito no está disponible para ser propuesto", vbInformation, "Aviso"
            End If
            Limpiar
            Exit Function
        End If
    End If
Else
    MsgBox "No existen datos del crédito o no se encuentra vigente", vbInformation, "Aviso"
End If
End Function

Private Sub cmdGenerar_Click()
Dim oNCred As COMNCredito.NCOMCredito
Dim oDCredAct As COMDCredito.DCOMCredActBD
Dim rs As ADODB.Recordset 'Agrego JOEP20171214 ACTA220-2017
Dim i As Integer 'JOEP20200804 Actividad Economica covid
Dim vMatrizDatosReprog As Variant
Dim MatrizModalidades As Variant

'JOEP20200804 Actividad Economica covid
If ValidaDatos Then
    Exit Sub
End If
'JOEP20200804 Actividad Economica covid
  
 'JOEP20200804 Actividad Economica covid
    If fnTipo = TipoPropuesto Then
        
        ReDim vMatrizDatosReprog(12)
        vMatrizDatosReprog(0) = Trim(txt_Actividad.Text)
        vMatrizDatosReprog(1) = IIf(OpNF.iTem(0).Enabled = False, -1, IIf(OpNF.iTem(1).value = True, 1, 0))
        vMatrizDatosReprog(2) = IIf(OpCA.iTem(0).Enabled = False, -1, IIf(OpCA.iTem(1).value = True, 1, 0))
        vMatrizDatosReprog(3) = Trim(txt_NewActividad.Text)
        vMatrizDatosReprog(4) = IIf(OpIngrAct.iTem(0).value = True, 1, IIf(OpIngrAct.iTem(1).value = True, 2, 3))
        vMatrizDatosReprog(5) = IIf(OPDimicEs.iTem(0).Enabled = True, IIf(OPDimicEs.iTem(0).value = True, 1, 2), -1)
        vMatrizDatosReprog(6) = IIf(OpDismPorc.iTem(0).Enabled = True, IIf(OpDismPorc.iTem(0).value = True, 1, IIf(OpDismPorc.iTem(1).value = True, 2, 3)), -1)
        vMatrizDatosReprog(7) = IIf(OpIngActPer.iTem(0).value = True, 1, IIf(OpIngActPer.iTem(1).value = True, 2, 3))
        vMatrizDatosReprog(8) = CCur(EditMoneyIngrVentNet)
        vMatrizDatosReprog(9) = CCur(EditMoneyGastoTot)
        vMatrizDatosReprog(10) = CCur(EditMoneyExce)
        vMatrizDatosReprog(11) = CCur(EditMoneyCP)
        
        If ValidaSolOcmSicm = False Then
            If fe_Modalidades.row >= 1 And fe_Modalidades.TextMatrix(1, 1) <> "" Then
                ReDim MatrizModalidades(fe_Modalidades.rows - 1, 5)
                For i = 1 To fe_Modalidades.rows - 1
                    If fe_Modalidades.TextMatrix(i, 2) = "." Then
                        MatrizModalidades(i, 1) = fe_Modalidades.TextMatrix(i, 3)
                        MatrizModalidades(i, 2) = fe_Modalidades.TextMatrix(i, 4)
                        MatrizModalidades(i, 3) = fe_Modalidades.TextMatrix(i, 5)
                    End If
                Next i
            End If
        End If
'Add JOEP20210306 garantia covid
    ElseIf fnTipo = TipoGarantia Then 'JOEP20210211 Garantia covid
        If fe_Modalidades.row >= 1 And fe_Modalidades.TextMatrix(1, 1) <> "" Then
                ReDim MatrizModalidades(fe_Modalidades.rows - 1, 5)
                For i = 1 To fe_Modalidades.rows - 1
                    If fe_Modalidades.TextMatrix(i, 2) = "." Then
                        MatrizModalidades(i, 1) = fe_Modalidades.TextMatrix(i, 3)
                        MatrizModalidades(i, 2) = fe_Modalidades.TextMatrix(i, 4)
                        MatrizModalidades(i, 3) = fe_Modalidades.TextMatrix(i, 5)
                    End If
                Next i
            End If
'Add JOEP20210306 garantia covid
    End If
 'JOEP20200804 Actividad Economica covid
        
    If MsgBox("Se va a generar la propuesta de la reprogramación del crédito, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        sMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
        'Inicio Agrego JOEP20171214 ACTA220-2017
     If fnTipo = TipoPropuesto Then
        Set oDCred = New COMDCredito.DCOMCredito
        Set rs = oDCred.ObtieneNivApro(ActXCodCta.NroCuenta)
            If Not (rs.EOF And rs.BOF) Then
                MsgBox "El crédito necesitará la Aprobación de: " & rs!cNivApro, vbInformation, "Aviso"
            End If
        Set oDCred = Nothing
    End If
        'Fin Agrego JOEP20171214 ACTA220-2017
                 
        Set oNCred = New COMNCredito.NCOMCredito
            'Call oNCred.RegistraReprogramacionEstado(ActXCodCta.NroCuenta, gdFecSis, gEstReprogPropuesto, fnSaldoReprog, sMovNro, , fnCuotasReprog, fdFecCuotaVenc, fdFecNuevaCuotaVenc, txtComent.Text, fbSolicAutorizacion)'comento JOEP20200804 Actividad Economica covid
            'Call oNCred.RegistraReprogramacionEstado(ActXCodCta.NroCuenta, gdFecSis, gEstReprogPropuesto, fnSaldoReprog, sMovNro, , fnCuotasReprog, fdFecCuotaVenc, fdFecNuevaCuotaVenc, txtComent.Text, fbSolicAutorizacion, , , vMatrizDatosReprog, m, , , MatrizModalidades) 'JOEP20200804 Actividad Economica covid
            Call oNCred.RegistraReprogramacionEstado(ActXCodCta.NroCuenta, gdFecSis, gEstReprogPropuesto, fnSaldoReprog, sMovNro, gnDiasAtraso, fnCuotasReprog, fdFecCuotaVenc, fdFecNuevaCuotaVenc, txtComent.Text, fbSolicAutorizacion, , , vMatrizDatosReprog, m, , , MatrizModalidades, , , gMatrixCalend, fnTipo, gMatrixDatos) 'Add JOEP20210306 garantia covid
        Set oNCred = Nothing
    
        'If GeneraPropuestaReprogramacion Then
        If fnTipo = TipoPropuesto Then
            If pdfActEco Then 'JOEP20200818 actividad economica covid
                MsgBox "Se ha generado la propuesta para la reprogramación del crédito", vbInformation, "Aviso"
                Limpiar
            Else
                Set oDCredAct = New COMDCredito.DCOMCredActBD
                Call oDCredAct.dEliminaColocacReprogramado(ActXCodCta.NroCuenta, gEstReprogPropuesto, sMovNro)
                Set oDCredAct = Nothing
            End If
        ElseIf fnTipo = TipoGarantia Then
            If PdfSimulacionCalend Then
                MsgBox "Se ha generado la propuesta para la reprogramación del crédito", vbInformation, "Aviso"
                Limpiar
            Else
                MsgBox "Error", vbInformation, "Aviso"
            End If
        End If
End Sub

Private Sub cmdCancelar_Click()
    Limpiar
End Sub

Private Sub Limpiar()
    ActXCodCta.CMAC = gsCodCMAC
    ActXCodCta.Age = gsCodAge
    ActXCodCta.Prod = ""
    ActXCodCta.Cuenta = ""
    lblCodigo.Caption = ""
    lblDOI.Caption = ""
    lblNomCliente.Caption = ""
    txtComent.Text = ""
    
    HabilitaControles True
    
    If ActXCodCta.Enabled Then ActXCodCta.SetFocusProd
    
    'JOEP20200804 Actividad Economica covid
    Set m = Nothing
    txt_Actividad.Text = ""
    OpNF.iTem(0).value = 0
    OpNF.iTem(1).value = 0
    OpCA.iTem(0).value = 0
    OpCA.iTem(1).value = 0
    txt_NewActividad.Text = ""
    OpIngrAct.iTem(0).value = 0
    OpIngrAct.iTem(1).value = 0
    OpIngrAct.iTem(2).value = 0
    
    OPDimicEs.iTem(0).value = 0
    OPDimicEs.iTem(1).value = 0
        
    OpDismPorc.iTem(0).value = 0
    OpDismPorc.iTem(1).value = 0
    OpDismPorc.iTem(2).value = 0
    
    OpIngActPer.iTem(0).value = 0
    OpIngActPer.iTem(1).value = 0
    OpIngActPer.iTem(2).value = 0
    
    EditMoneyIngrVentNet = Format(0, "#,##0.00")
    EditMoneyGastoTot = Format(0, "#,##0.00")
    EditMoneyExce = Format(0, "#,##0.00")
    EditMoneyCP = Format(0, "#,##0.00")
    cTpProducto = ""
    fr_NegViFunc.Visible = True
    nMOntoCPIfis = 0
    LimpiaFlex fe_Modalidades
    'JOEP20200804 Actividad Economica covid
    
    txtRechazoOtrosDias.Text = 0 'reversion
    cmbRechazoOtros.ListIndex = -1 'reversion
End Sub

Private Sub cmdRechazar_Click()

Dim oNCred As COMNCredito.NCOMCredito

'JOEP20200804 Actividad Economica covid
    If ValidaDatos Then
        Exit Sub
    End If
'JOEP20200804 Actividad Economica covid

'JOEP20200804 COMENTO Actividad Economica covid
'    If Trim(txtComent.Text) = "" Then
'        MsgBox "Debe registrar el motivo", vbInformation, "Aviso"
'        txtComent.SetFocus
'        Exit Sub
'    End If
'JOEP20200804 COMENTO Actividad Economica covid
    Dim RezOtr2 As Integer
    Dim newDias As Integer
    If txtRechazoOtrosDias.Text = "" Then
        newDias = -1
    Else
        newDias = txtRechazoOtrosDias.Text
    End If
    
    If cmbRechazoOtros.Text = "" Then
        RezOtr2 = -1
    Else
        RezOtr2 = Right(cmbRechazoOtros.Text, 1)
    End If
        
    If MsgBox("Se va a rechazar la reprogramación del crédito, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        
    sMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            
    Set oNCred = New COMNCredito.NCOMCredito
        'Call oNCred.RegistraReprogramacionEstado(ActXCodCta.NroCuenta, gdFecSis, gEstReprogRechazado, fnSaldoReprog, sMovNro, , , , , txtComent.Text)
        Call oNCred.RegistraReprogramacionEstado(ActXCodCta.NroCuenta, gdFecSis, gEstReprogRechazado, fnSaldoReprog, sMovNro, , , , , txtComent.Text, , , , , , , , , RezOtr2, newDias)
    Set oNCred = Nothing

    MsgBox "Se ha rechazado la reprogramación del crédito", vbInformation, "Aviso"
    Limpiar
    
End Sub

Private Sub cmdsalir_Click()
    Limpiar
    Unload Me
End Sub

'ADD JHCU 26-09-2020
Private Sub cmdSolicitud_Click()
    Limpiar
    'Call FrmVerCreditoSolWeb.Inicio(ActXCodCta)
    Call FrmVerCreditoSolWeb.Inicio(ActXCodCta, , fnTipo) 'Add JOEP20210306 garantia covid
    ActXCodCta.SetFocusCuenta
End Sub
'END JHCU

Private Sub EditMoneyExce_Change()
    calculo (1)
End Sub

Private Sub EditMoneyGastoTot_Change()
    calculo (2)
    calculo (1)
End Sub

Private Sub EditMoneyIngrVentNet_Change()
    calculo (2)
    calculo (1)
End Sub

Private Sub txtComent_KeyPress(KeyAscii As Integer)
    If Len(txtComent.Text) = 0 Then
        KeyAscii = fgIntfMayusculas(KeyAscii)
    End If
End Sub

Private Sub txtComent_LostFocus()
    If Len(txtComent.Text) > 0 Then
        txtComent.Text = UCase(Left(txtComent.Text, 1)) & Mid(txtComent.Text, 2, Len(txtComent.Text))
    End If
End Sub

'JOEP20200804 Actividad Economica covid
Private Sub Form_Load()
    Call HabilitaControles(False, 1)
End Sub

Private Sub VaidaDatos()
    If OpCA.iTem(0) = True Then
        txt_NewActividad.Enabled = True
    Else
        txt_NewActividad.Enabled = False
        txt_NewActividad.Text = ""
    End If
    
    If OpIngrAct.iTem(2) = True Then
        fr_DismiEs.Enabled = True
        OPDimicEs.iTem(0).Enabled = True
        OPDimicEs.iTem(1).Enabled = True
        
        If Left(cTpProducto, 1) = "7" Then
            fr_DismPorc.Enabled = True
            OpDismPorc.iTem(0).Enabled = True
            OpDismPorc.iTem(1).Enabled = True
            OpDismPorc.iTem(2).Enabled = True
        End If
    Else
        fr_DismiEs.Enabled = False
        OPDimicEs.iTem(0).Enabled = False
        OPDimicEs.iTem(1).Enabled = False
        OPDimicEs.iTem(0).value = 0
        OPDimicEs.iTem(1).value = 0
        
        If Left(cTpProducto, 1) = "7" Then
            fr_DismPorc.Enabled = False
            OpDismPorc.iTem(0).Enabled = False
            OpDismPorc.iTem(1).Enabled = False
            OpDismPorc.iTem(2).Enabled = False
        End If
    End If
    
End Sub

Private Sub OpIngrAct_Click(Index As Integer)
    Call VaidaDatos
End Sub

Private Sub OpCA_Click(Index As Integer)
    Call VaidaDatos
End Sub

Private Sub OpNF_Click(Index As Integer)
    Call VaidaDatos
End Sub

Private Sub cmdCPIFIs_Click()
Dim i As Integer
    frmCredReprogPropuestaIFIS.Inicio ActXCodCta.NroCuenta, m
    EditMoneyCP.Text = 0
    nMOntoCPIfis = 0
    If IsArray(m) Then
        If UBound(m) > 0 Then
             For i = 1 To UBound(m)
                nMOntoCPIfis = Format(nMOntoCPIfis + CCur(m(i, 6)), "#,##0.00")
             Next i
             
            If CCur(EditMoneyExce) = 0 Then
                EditMoneyCP = Format(0, "#,##0.00")
            Else
                EditMoneyCP = Format(nMOntoCPIfis / CCur(EditMoneyExce), "#,##0.00")
            End If
        Else
            EditMoneyCP = Format(0, "#,##0.00")
        End If
    Else
         EditMoneyCP = Format(0, "#,##0.00")
    End If
End Sub
'JOEP20200804 Actividad Economica covid

'Private Sub HabilitaControles(ByVal pbHabilitaBus As Boolean)
Private Sub HabilitaControles(ByVal pbHabilitaBus As Boolean, Optional ByVal Opcion As Integer = 0)

If Opcion = 0 Then
    ActXCodCta.Enabled = pbHabilitaBus
    cmdBuscar.Enabled = pbHabilitaBus
    FraComent.Enabled = Not pbHabilitaBus
    
    If fnTipo = TipoPropuesto Then
        cmdGenerar.Enabled = Not pbHabilitaBus
        OpNF.iTem(0).Enabled = Not pbHabilitaBus
        OpNF.iTem(1).Enabled = Not pbHabilitaBus
        OpCA.iTem(0).Enabled = Not pbHabilitaBus
        OpCA.iTem(1).Enabled = Not pbHabilitaBus
        txt_NewActividad.Enabled = Not pbHabilitaBus
        OpIngrAct.iTem(0).Enabled = Not pbHabilitaBus
        OpIngrAct.iTem(1).Enabled = Not pbHabilitaBus
        OpIngrAct.iTem(2).Enabled = Not pbHabilitaBus
        
        OPDimicEs.iTem(0).Enabled = False
        OPDimicEs.iTem(1).Enabled = False
            
        OpDismPorc.iTem(0).Enabled = Not pbHabilitaBus
        OpDismPorc.iTem(1).Enabled = Not pbHabilitaBus
        OpDismPorc.iTem(2).Enabled = Not pbHabilitaBus
        
        OpIngActPer.iTem(0).Enabled = Not pbHabilitaBus
        OpIngActPer.iTem(1).Enabled = Not pbHabilitaBus
        OpIngActPer.iTem(2).Enabled = Not pbHabilitaBus
        
        EditMoneyIngrVentNet.Enabled = Not pbHabilitaBus
        EditMoneyGastoTot.Enabled = Not pbHabilitaBus
        EditMoneyExce.Enabled = Not pbHabilitaBus
        EditMoneyCP.Enabled = False
          
        cmdCPIFIs.Enabled = Not pbHabilitaBus
        cmdSimulador.Enabled = Not pbHabilitaBus
        fe_Modalidades.Enabled = Not pbHabilitaBus
    ElseIf fnTipo = TipoRechazo Then
        cmdRechazar.Enabled = Not pbHabilitaBus
        frRechazoOtros.Enabled = Not pbHabilitaBus
'Add JOEP20210306 garantia covid
    ElseIf fnTipo = TipoGarantia Then
        cmdBuscar.Enabled = False
        FraComent.Enabled = False
'Add JOEP20210306 garantia covid
    End If
        
ElseIf Opcion = 1 Then
    fr_NegViFunc.Enabled = pbHabilitaBus
    fr_CliMantAct.Enabled = pbHabilitaBus
    txt_NewActividad.Enabled = False
    fr_IngAct.Enabled = pbHabilitaBus
    
    fr_DismiEs.Enabled = pbHabilitaBus
    OPDimicEs.iTem(0).Enabled = False
    OPDimicEs.iTem(1).Enabled = False
    
    fr_DismPorc.Enabled = pbHabilitaBus
    fr_IngActPert.Enabled = pbHabilitaBus
    EditMoneyIngrVentNet.Enabled = pbHabilitaBus
    EditMoneyGastoTot.Enabled = pbHabilitaBus
    EditMoneyExce.Enabled = pbHabilitaBus
    cmdCPIFIs.Enabled = pbHabilitaBus
    cmdSimulador.Enabled = pbHabilitaBus
    fe_Modalidades.Enabled = pbHabilitaBus
    FraComent.Enabled = pbHabilitaBus 'Add JOEP20210306 garantia covid
ElseIf Opcion = 2 Then
    fr_NegViFunc.Enabled = pbHabilitaBus
    OpNF.iTem(0).Enabled = pbHabilitaBus
    OpNF.iTem(1).Enabled = pbHabilitaBus
    fr_DismPorc.Enabled = pbHabilitaBus
    OpDismPorc.iTem(0).Enabled = pbHabilitaBus
    OpDismPorc.iTem(1).Enabled = pbHabilitaBus
    OpDismPorc.iTem(2).Enabled = pbHabilitaBus
'Add JOEP20210306 garantia covid
ElseIf Opcion = 3 Then
    cmdGenerar.Enabled = Not pbHabilitaBus
    If fnTipo = TipoGarantia Then
        fr_VerActEco.Enabled = pbHabilitaBus
        FraComent.Enabled = pbHabilitaBus
    End If
'Add JOEP20210306 garantia covid
End If
    
End Sub

Private Function pdfActEco() As Boolean
    
    Dim RsDatos As ADODB.Recordset, rsExo As ADODB.Recordset, rsModalidades As ADODB.Recordset
    Dim oDoc  As cPDF
    Set oDoc = New cPDF
    
    Dim a As Currency
    Dim nFila As Integer
    Dim i As Integer
    Dim FilaIncr As Integer
    Dim EspTitle As Integer
    
    Set oDCred = New COMDCredito.DCOMCredito
    Set RsDatos = oDCred.RecuperaDatosPropuestaReprogramacion(ActXCodCta.NroCuenta, gdFecSis)
    Set rsExo = oDCred.RecuperaColocacReprogExoneraSolicitud(ActXCodCta.NroCuenta, gdFecSis)
    Set rsModalidades = oDCred.ReprogramacionPropuestaObtieneModalidades(ActXCodCta.NroCuenta, gdFecSis)
    
    Set oDCred = Nothing
           
    pdfActEco = False
           
    'Creación del Archivo
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Propuesta de Reprogramacion Nº " & ActXCodCta.NroCuenta
    oDoc.Title = "Propuesta Nº " & ActXCodCta.NroCuenta
        
    If Not oDoc.PDFCreate(App.Path & "\Spooler\PropuestaReprogramacion_" & ActXCodCta.NroCuenta & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Function
    End If
    
    'Contenido
    oDoc.Fonts.Add "F1", "Courier New", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Courier New", TrueType, Bold, WinAnsiEncoding
    oDoc.LoadImageFromFile App.Path & "\logo_cmacmaynas.bmp", "Logo"
    
    EspTitle = 20
    nFila = 80
    FilaIncr = 15
'Inicio de creacion PDF
    If Not (RsDatos.BOF Or RsDatos.EOF) Then
    
        'Tamaño de hoja A4
        oDoc.NewPage A4_Vertical
        oDoc.WImage 70, 50, 50, 100, "Logo"
        
        oDoc.WTextBox 5, 25, 820, 550, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        
'************************************ Cabecera ************************************************
        oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
        oDoc.WTextBox 50, 420, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
        
        oDoc.WTextBox nFila, 50, 15, 500, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 100, 10, 400, "PROPUESTA Y EVALUACIÓN DE REPROGRAMACIÓN DE PAGOS", "F2", 10, hCenter
        
        nFila = nFila + EspTitle
        oDoc.WTextBox nFila, 55, 10, 300, "NOMBRE DE CLIENTE", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 195, 10, 355, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 200, 10, 300, RsDatos!cPersNombreCli, "F1", 7.5, hLeft
        nFila = nFila + FilaIncr
        oDoc.WTextBox nFila, 55, 10, 300, "N° CRÉDITO", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 195, 10, 150, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 200, 10, 300, RsDatos!cCtaCod, "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 350, 10, 300, "ANALISTA", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 450, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 455, 10, 300, RsDatos!cUserAnalista, "F1", 7.5, hLeft
        nFila = nFila + FilaIncr
        oDoc.WTextBox nFila, 55, 10, 300, "MONTO APROBADO", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 195, 10, 150, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 200, 10, 300, RsDatos!cMoneda & Space(1) & Format(RsDatos!nMontoCol, "#,##0.00"), "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 350, 10, 300, "SALDO CAPITAL", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 450, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 455, 10, 300, RsDatos!cMoneda & Space(1) & Format(RsDatos!nSaldoCap, "#,##0.00"), "F1", 7.5, hLeft
        nFila = nFila + FilaIncr
        oDoc.WTextBox nFila, 55, 10, 300, "TOTAL CUOTAS" & "   ", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 195, 10, 150, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 200, 10, 300, RsDatos!nNrocuotas, "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 350, 10, 300, "CUOTAS RESTANTES" & "   ", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 450, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 455, 10, 300, RsDatos!nCuotasReprog, "F1", 7.5, hLeft
        nFila = nFila + FilaIncr
        oDoc.WTextBox nFila, 55, 10, 300, "FECHA DE PAGO ACTUAL", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 195, 10, 150, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 200, 10, 300, RsDatos!dFecCuotaVenc, "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 350, 10, 300, "NUEVA FECHA DE PAGO", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 450, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 455, 10, 300, RsDatos!dFecNuevaCuotaVenc, "F1", 7.5, hLeft
        nFila = nFila + FilaIncr
        oDoc.WTextBox nFila, 55, 10, 300, "DÍAS A REPROGRAMAR", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 195, 10, 150, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 200, 10, 300, RsDatos!nDiasReprog, "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 350, 10, 300, "TIPO DE GARANTIA", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 450, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 455, 10, 300, RsDatos!cTpGarantia, "F1", 7.5, hLeft
        nFila = nFila + FilaIncr
        If RsDatos!cNivAprDesc <> "GERENCIA DE RIESGOS" Then
            oDoc.WTextBox nFila, 55, 10, 300, "NIVEL DE APROBACION", "F2", 7.5, hLeft
            oDoc.WTextBox nFila, 195, 10, 210, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 200, 10, 300, RsDatos!cNivAprDesc, "F1", 7.5, hLeft
        Else
            oDoc.WTextBox nFila, 55, 10, 300, "V°.B°", "F2", 7.5, hLeft
            oDoc.WTextBox nFila, 195, 10, 210, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 200, 50, 300, RsDatos!cNivAprDesc, "F1", 7.5, hjustify
        End If
        nFila = nFila + FilaIncr
        
        If Not (rsExo.BOF And rsExo.EOF) Then
            oDoc.WTextBox nFila, 55, 10, 300, "EXONERACIÓN", "F2", 7.5, hLeft
            For i = 0 To rsExo.RecordCount - 1
                oDoc.WTextBox nFila, 200, 10, 300, rsExo!cExoneraDesc, "F1", 7.5, hLeft
                nFila = nFila + FilaIncr
                rsExo.MoveNext
            Next i
        End If
        
        oDoc.WTextBox nFila, 55, 10, 300, "MOTIVO DE LA REPROGRAMACIÓN", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 195, (RsDatos!nCantLetraMR * 10) + 5, 355, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 200, 10, 330, RsDatos!cMotivoReprog, "F1", 7.5, hjustify
'************************************ Cabecera ************************************************

        nFila = nFila + (RsDatos!nCantLetraMR * 10) + 5
 
If (Left(RsDatos!cTpProducto, 1) = 5 Or Left(RsDatos!cTpProducto, 1) = 6) Then
        oDoc.WTextBox nFila + 5, 50, 15, 500, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila + 5, 55, 10, 400, "VERIFICACIÓN DE LA ACTIVIDAD ECONÓMICA", "F2", 10, hCenter
        nFila = nFila + EspTitle + 5
        oDoc.WTextBox nFila, 55, 10, 300, "ACTIVIDAD", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 195, 10, 355, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 200, 10, 300, RsDatos!ActividadFormEval, "F1", 7.5, hLeft
        nFila = nFila + FilaIncr
        oDoc.WTextBox nFila, 55, 10, 300, "NEGOCIO VIENE FUNCIONANDO", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 195, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 200, 10, 300, RsDatos!cNegVinFunc, "F1", 7.5, hLeft
        nFila = nFila + FilaIncr
        oDoc.WTextBox nFila, 55, 10, 300, "CLIENTE MANTIENE ACTIVIDAD", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 195, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 200, 10, 300, RsDatos!cCliMatAct, "F1", 7.5, hLeft
        
        If RsDatos!CliMatAct = 0 Then
            nFila = nFila + FilaIncr
            oDoc.WTextBox nFila, 55, 10, 300, "INDICA NUEVA ACTIVIDAD", "F2", 7.5, hLeft
            oDoc.WTextBox nFila, 195, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 200, 10, 300, RsDatos!NewActividad, "F1", 7.5, hLeft
        End If
        
        nFila = nFila + FilaIncr
        oDoc.WTextBox nFila, 55, 10, 300, "INGRESOS ACTUALES", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 195, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 200, 10, 300, RsDatos!cIngActuales, "F1", 7.5, hLeft
        
        If RsDatos!IngActuales = 3 Then
            nFila = nFila + FilaIncr
            oDoc.WTextBox nFila, 55, 10, 300, "LA DISMINUCION ES", "F2", 7.5, hLeft
            oDoc.WTextBox nFila, 195, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 200, 10, 300, RsDatos!cDimEs, "F1", 7.5, hLeft
        End If
        
        nFila = nFila + FilaIncr
        oDoc.WTextBox nFila, 55, 10, 300, "EN QUE PORCENTAJE(%) DIMINUYO", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 195, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 200, 10, 300, RsDatos!cDimPorc, "F1", 7.5, hLeft
        
        nFila = nFila + FilaIncr
        oDoc.WTextBox nFila, 55, 10, 300, "INGRESOS ACTUALES PERMITEN", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 195, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 200, 10, 300, RsDatos!cIngActPermt, "F1", 7.5, hLeft
                
        nFila = nFila + FilaIncr
        oDoc.WTextBox nFila, 55, 10, 300, "INGRESOS Ó VENTAS NETAS", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 195, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 200, 10, 300, Format(RsDatos!IngVentasNet, "#,##0.00"), "F1", 7.5, hLeft
        
        oDoc.WTextBox nFila, 350, 10, 300, "GASTOS TOTALES", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 450, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 455, 10, 300, Format(RsDatos!GastTotal, "#,##0.00"), "F1", 7.5, hLeft
        
        nFila = nFila + FilaIncr
        oDoc.WTextBox nFila, 55, 10, 300, "EXCEDENTE", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 195, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 200, 10, 300, Format(RsDatos!Excedente, "#,##0.00"), "F1", 7.5, hLeft
        
        oDoc.WTextBox nFila, 350, 10, 300, "C/P (CUOTAS IFIS/EXC)", "F2", 7.5, hLeft
        
        If RsDatos!Excedente = 0 Then
            oDoc.WTextBox nFila, 450, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 455, 10, 300, Round(0, 2) & "%", "F1", 7.5, hLeft
        Else
            oDoc.WTextBox nFila, 450, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 455, 10, 300, Round((RsDatos!nTotalIfis / RsDatos!Excedente), 2) & "%", "F1", 7.5, hLeft
        End If
        nFila = nFila + FilaIncr
         
        If Not (rsModalidades.BOF And rsModalidades.EOF) Then
        oDoc.WTextBox nFila, 55, 10, 300, "ALTERNATIVAS DE REPRO.", "F2", 7.5, hLeft
        
        oDoc.WTextBox nFila, 195, 10, 95, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 200, 10, 95, "Modalidad", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 290, 10, 35, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 295, 10, 35, "Calif.", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 325, 10, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 330, 10, 60, "Cuota Aprx.", "F2", 7.5, hLeft
        
        nFila = nFila + FilaIncr
            For i = 0 To rsModalidades.RecordCount - 1
                oDoc.WTextBox nFila - 5, 195, 10, 95, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                oDoc.WTextBox nFila - 5, 200, 10, 95, rsModalidades!cConsDescripcion, "F1", 7.5, hLeft
                oDoc.WTextBox nFila - 5, 290, 10, 35, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                oDoc.WTextBox nFila - 5, 295, 10, 35, rsModalidades!cCalif, "F1", 7.5, hLeft
                oDoc.WTextBox nFila - 5, 325, 10, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                oDoc.WTextBox nFila - 5, 330, 10, 60, Format(rsModalidades!nCuota, "#,#0.00"), "F1", 7.5, hLeft
                rsModalidades.MoveNext
                
                nFila = nFila + 10 'FilaIncr
            Next i
        End If
        
        nFila = nFila + 10 'EspTitle
        
ElseIf (Left(RsDatos!cTpProducto, 1) = 7) Then
        oDoc.WTextBox nFila + 5, 50, 15, 500, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila + 5, 55, 10, 400, "VERIFICACIÓN DE CONTINUIDAD DE INGRESOS", "F2", 10, hCenter
        nFila = nFila + EspTitle + 5
        oDoc.WTextBox nFila, 55, 10, 300, "ACTIVIDAD/INGRESO/TRABAJO ACTUAL", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 215, 10, 335, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 220, 10, 300, RsDatos!ActividadFormEval, "F1", 7.5, hLeft
        nFila = nFila + FilaIncr
        oDoc.WTextBox nFila, 55, 10, 300, "¿CONTINUA TRABAJANDO?", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 215, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 220, 10, 300, RsDatos!cCliMatAct, "F1", 7.5, hLeft
        
        If RsDatos!NegVinFunc = 0 Then
            nFila = nFila + FilaIncr
            oDoc.WTextBox nFila, 55, 10, 300, "NUEVA ACTIVIDAD O ACTUAL EMPLEO", "F2", 7.5, hLeft
            oDoc.WTextBox nFila, 215, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 220, 10, 300, RsDatos!NewActividad, "F1", 7.5, hLeft
        End If
        
        nFila = nFila + FilaIncr
        oDoc.WTextBox nFila, 55, 10, 300, "INGRESOS ACTUALES", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 215, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 220, 10, 300, RsDatos!cIngActuales, "F1", 7.5, hLeft
                        
        If RsDatos!IngActuales = 3 Then
            nFila = nFila + FilaIncr
            oDoc.WTextBox nFila, 55, 10, 300, "LA DISMINUCION ES", "F2", 7.5, hLeft
            oDoc.WTextBox nFila, 215, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 220, 10, 300, RsDatos!cDimEs, "F1", 7.5, hLeft
            
            nFila = nFila + FilaIncr
            oDoc.WTextBox nFila, 55, 10, 300, "EN QUE PORCENTAJE(%) DIMINUYO", "F2", 7.5, hLeft
            oDoc.WTextBox nFila, 215, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 220, 10, 300, RsDatos!cDimPorc, "F1", 7.5, hLeft
        End If
                        
        nFila = nFila + FilaIncr
        oDoc.WTextBox nFila, 55, 10, 300, "INGRESOS ACTUALES PERMITEN", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 215, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 220, 10, 300, RsDatos!cIngActPermt, "F1", 7.5, hLeft
                        
        nFila = nFila + FilaIncr
        oDoc.WTextBox nFila, 55, 10, 300, "INGRESOS Ó VENTAS NETAS", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 215, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 220, 10, 300, Format(RsDatos!IngVentasNet, "#,##0.00"), "F1", 7.5, hLeft
        
        oDoc.WTextBox nFila, 350, 10, 300, "GASTOS TOTALES", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 450, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 455, 10, 300, Format(RsDatos!GastTotal, "#,##0.00"), "F1", 7.5, hLeft
       
        nFila = nFila + FilaIncr
        oDoc.WTextBox nFila, 55, 10, 300, "EXCEDENTE", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 215, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 220, 10, 300, Format(RsDatos!Excedente, "#,##0.00"), "F1", 7.5, hLeft
        
        oDoc.WTextBox nFila, 350, 10, 300, "C/P (CUOTAS IFIS/EXC)", "F2", 7.5, hLeft
        If RsDatos!Excedente = 0 Then
            oDoc.WTextBox nFila, 450, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 455, 10, 300, Round(0, 2) & "%", "F1", 7.5, hLeft
        Else
            oDoc.WTextBox nFila, 450, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 455, 10, 300, Round((RsDatos!nTotalIfis / RsDatos!Excedente), 2) & "%", "F1", 7.5, hLeft
        End If
        
        nFila = nFila + FilaIncr
        If Not (rsModalidades.BOF And rsModalidades.EOF) Then
        oDoc.WTextBox nFila, 55, 10, 300, "ALTERNATIVAS DE REPRO.", "F2", 7.5, hLeft
        
        oDoc.WTextBox nFila, 215, 10, 95, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 220, 10, 95, "Modalidad", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 310, 10, 35, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 315, 10, 35, "Calif.", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 345, 10, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 350, 10, 60, "Cuota Aprx.", "F2", 7.5, hLeft
        
        nFila = nFila + FilaIncr
            For i = 0 To rsModalidades.RecordCount - 1
                oDoc.WTextBox nFila - 5, 215, 10, 95, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                oDoc.WTextBox nFila - 5, 220, 10, 95, rsModalidades!cConsDescripcion, "F1", 7.5, hLeft
                oDoc.WTextBox nFila - 5, 310, 10, 35, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                oDoc.WTextBox nFila - 5, 315, 10, 35, rsModalidades!cCalif, "F1", 7.5, hLeft
                oDoc.WTextBox nFila - 5, 345, 10, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                oDoc.WTextBox nFila - 5, 350, 10, 60, Format(rsModalidades!nCuota, "#,#0.00"), "F1", 7.5, hLeft
                rsModalidades.MoveNext

                nFila = nFila + 10 'FilaIncr
            Next i
        End If
        
        nFila = nFila + 10 'EspTitle
End If
        oDoc.WTextBox nFila, 50, 15, 500, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 55, 10, 400, "COMENTARIO DE ANALISTA", "F2", 10, hCenter
        
        oDoc.WTextBox nFila + 15, 50, 150, 500, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila + 15, 55, 10, 300, RsDatos!cMotivo, "F1", 7.5, hjustify
        
        oDoc.WTextBox 740, 55, 10, 300, "____________________", "F1", 7.5, hLeft
        oDoc.WTextBox 750, 55, 10, 300, "FIRMA DE ANALISTA", "F2", 7.5, hLeft
        
        oDoc.WTextBox 740, 300, 10, 300, "____________________", "F1", 7.5, hLeft
        oDoc.WTextBox 750, 300, 10, 300, "FIRMA DE CC/JA", "F2", 7.5, hLeft
        
        oDoc.PDFClose
        oDoc.Show
        pdfActEco = True
    Else
        MsgBox "Los Datos de la propuesta del Crédito no han sido Registrados Correctamente", vbInformation, "Aviso"
    End If
    Set RsDatos = Nothing
End Function

Private Function ValidaDatos() As Boolean
Dim obValDts As New COMDCredito.DCOMCredito
Dim rsValDts As ADODB.Recordset
Dim i As Integer
Dim pMatzIfis As Integer
Dim nModalidad As Integer
Dim RezOtr As Integer
Dim DiasNew As Integer

If fnTipo = TipoPropuesto Then
    nModalidad = -1
    If IsArray(m) Then
        If UBound(m) <= 0 Then
            pMatzIfis = 0
        Else
            pMatzIfis = 1
        End If
    Else
        pMatzIfis = 0
    End If

    nModalidad = -1
        
    If ValidaSolOcmSicm = False Then
        For i = 1 To fe_Modalidades.rows - 1
            If fe_Modalidades.TextMatrix(i, 2) = "." And fe_Modalidades.TextMatrix(i, 3) <> "" And fe_Modalidades.TextMatrix(i, 4) <> "" Then
                nModalidad = 1
            End If
        Next i
    End If
ElseIf fnTipo = TipoRechazo Then
    
    If cmbRechazoOtros.Text = "" Then
        RezOtr = -1
    Else
        RezOtr = Right(cmbRechazoOtros.Text, 1)
    End If

    If txtRechazoOtrosDias.Text = "" Then
        DiasNew = -1
    Else
        DiasNew = txtRechazoOtrosDias.Text
    End If
    
ElseIf fnTipo = TipoGarantia Then
    nModalidad = -1
    If ValidaSolOcmSicm = False Then
        For i = 1 To fe_Modalidades.rows - 1
            If fe_Modalidades.TextMatrix(i, 2) = "." And fe_Modalidades.TextMatrix(i, 3) <> "" And fe_Modalidades.TextMatrix(i, 4) <> "" Then
                nModalidad = 1
            End If
        Next i
    End If
    
End If


If cmbRechazoOtros.Text = "" Then
    RezOtr = -1
Else
    RezOtr = Right(cmbRechazoOtros.Text, 1)
End If

If txtRechazoOtrosDias.Text = "" Then
    DiasNew = -1
Else
    DiasNew = txtRechazoOtrosDias.Text
End If

Set obValDts = New COMDCredito.DCOMCredito
Set rsValDts = obValDts.ReprogramacionValDtsPropuesta(ActXCodCta.NroCuenta, fnTipo, Trim(txtComent.Text), _
CInt(OpNF.iTem(0).value), CInt(OpNF.iTem(1).value), _
CInt(OpCA.iTem(0).value), CInt(OpCA.iTem(1).value), _
txt_NewActividad, _
CInt(OpIngrAct.iTem(0).value), CInt(OpIngrAct.iTem(1).value), CInt(OpIngrAct.iTem(2).value), _
CInt(OPDimicEs.iTem(0).value), CInt(OPDimicEs.iTem(1).value), _
CInt(OpDismPorc.iTem(0).value), CInt(OpDismPorc.iTem(1).value), CInt(OpDismPorc.iTem(2).value), _
CInt(OpIngActPer.iTem(0).value), CInt(OpIngActPer.iTem(1).value), CInt(OpIngActPer.iTem(2).value), _
EditMoneyIngrVentNet.Text, EditMoneyGastoTot.Text, EditMoneyExce.Text, pMatzIfis, nModalidad, _
IIf(ValidaSolOcmSicm, 1, 0), RezOtr, DiasNew)
'reversion RezOtr
ValidaDatos = False

If Not (rsValDts.BOF And rsValDts.EOF) Then
    If rsValDts!MsgBox <> "" Then
        MsgBox rsValDts!MsgBox, vbInformation, "Aviso"
        ValidaDatos = True
    End If
End If

End Function

Private Sub calculo(ByVal pOp As Integer)
    If pOp = 1 Then
        If CCur(EditMoneyCP) <> 0 Then
            If CCur(EditMoneyExce) = 0 Then
                EditMoneyCP = Format(0, "#,##0.00")
            Else
                EditMoneyCP = Format(CCur(nMOntoCPIfis) / CCur(EditMoneyExce), "#,##0.00")
            End If
        End If
    ElseIf pOp = 2 Then
        EditMoneyExce = Format(CCur(EditMoneyIngrVentNet) - CCur(EditMoneyGastoTot), "#,##0.00")
    End If
End Sub

Private Sub cmdSimulador_Click()
Dim nCalf As Integer
Dim i As Integer
Dim nMod As Integer
Dim nCuota As Currency
nCalf = 0
nMod = -1
    If Len(ActXCodCta.NroCuenta) >= 18 Then
        'selecciona modalidad sin resultado de calificacion
        For i = 1 To fe_Modalidades.rows - 1
            If fe_Modalidades.TextMatrix(i, 2) = "." And fe_Modalidades.TextMatrix(i, 4) = "" Then
                nMod = fe_Modalidades.TextMatrix(i, 3)
                Exit For
            End If
        Next i
    
        If nMod <> -1 Then
            'Call frmCredReprogSimulador.Inicio(ActXCodCta.NroCuenta, nMod, nCuota, nCalf)
            Call frmCredReprogSimulador.Inicio(ActXCodCta.NroCuenta, nMod, nCuota, nCalf, gMatrixCalend, fnTipo, gMatrixDatos) 'Add JOEP20210306 garantia covid
            'Registra calificion
            For i = 1 To fe_Modalidades.rows - 1
                If fe_Modalidades.TextMatrix(i, 2) = "." And fe_Modalidades.TextMatrix(i, 3) = nMod And nCalf = 1 Then
                    fe_Modalidades.TextMatrix(i, 4) = "Si"
                    fe_Modalidades.TextMatrix(i, 5) = Format(nCuota, "#,#0.00")
                    Exit For
                ElseIf fe_Modalidades.TextMatrix(i, 2) = "." And fe_Modalidades.TextMatrix(i, 3) = nMod And nCalf = 2 Then
                    fe_Modalidades.TextMatrix(i, 4) = "No"
                    fe_Modalidades.TextMatrix(i, 5) = Format(nCuota, "#,#0.00")
                    Exit For
                End If
            Next i
            
        End If
        
        If nMod = -1 Then
            MsgBox "Seleccióne las [Facilidades Crediticias]", vbInformation, "Aviso"
        End If
    End If
End Sub

Private Sub fe_Modalidades_Click()
    If fe_Modalidades.Col = 2 Then
        If fe_Modalidades.TextMatrix(fe_Modalidades.row, 4) <> "" And fe_Modalidades.TextMatrix(fe_Modalidades.row, 5) <> "" Then
            fe_Modalidades.TextMatrix(fe_Modalidades.row, 4) = ""
            fe_Modalidades.TextMatrix(fe_Modalidades.row, 5) = ""
        End If
    End If
End Sub

Private Function ValidaSolOcmSicm() As Boolean
Dim objOcmSic As COMDCredito.DCOMCredito
Dim rsOCMSIC As ADODB.Recordset
Set objOcmSic = New COMDCredito.DCOMCredito

ValidaSolOcmSicm = False
If fnTipo = TipoPropuesto Then
    If ActXCodCta.NroCuenta <> "" Then
        Set rsOCMSIC = objOcmSic.ReprogPropuestaValOcmSicm(ActXCodCta.NroCuenta)
        If Not (rsOCMSIC.BOF And rsOCMSIC.EOF) Then
            If rsOCMSIC!nValor <> 1 Then
                fr_Modalidades.Visible = False
                cmdSimulador.Visible = False
                FraComent.Width = 7455
                txtComent.Width = 7220
                ValidaSolOcmSicm = True
            Else
                fr_Modalidades.Visible = True
                cmdSimulador.Visible = True
                FraComent.Width = 3375
                txtComent.Width = 3135
                ValidaSolOcmSicm = False
            End If
        End If
    End If
End If
Set objOcmSic = Nothing
RSClose rsOCMSIC
    
End Function

Private Sub MedidasFormulario(ByVal pnOp As Integer)

    If pnOp = TipoPropuesto Then
        fr_VerActEco.Top = 2000
        FraComent.Top = 6300
        fr_Modalidades.Top = 6300
        Me.Width = 7755
        Me.Height = 9240
        cmdCancelar.Top = 8400
        cmdGenerar.Top = 8400
        cmdSalir.Top = 8400
        cmdSimulador.Top = 8400
        cmdRechazar.Top = 8400
    ElseIf pnOp = TipoRechazo Then
        FraComent.Top = 2700
        Me.Width = 6615
        Me.Height = 5200
        frRechazoOtros.Width = 6265
        cmdCancelar.Top = 4300
        cmdGenerar.Top = 4300
        cmdSalir.Top = 4300
        cmdRechazar.Top = 4300
        cmdSalir.Left = 5200
        cmdRechazar.Left = 3800
        FraComent.Width = 6265
        txtComent.Width = 6000
        FraComent.Height = 1500
        txtComent.Height = 1150
        Frame1.Width = 6265
        lblNomCliente.Width = 4575
    ElseIf pnOp = TipoGarantia Then
        fr_VerActEco.Top = 2000
        FraComent.Top = 6300
        fr_Modalidades.Top = 6300
        Me.Width = 7755
        Me.Height = 9240
        cmdCancelar.Top = 8400
        cmdGenerar.Top = 8400
        cmdSalir.Top = 8400
        cmdSimulador.Top = 8400
        cmdRechazar.Top = 8400
    Else
        
    End If
    
End Sub

Private Sub cmbRechazoOtros_Click()
Dim obRO As COMDCredito.DCOMCredito
Dim rsRO As ADODB.Recordset
Set obRO = New COMDCredito.DCOMCredito

    If Right(cmbRechazoOtros.Text, 1) = 3 Then
        txtRechazoOtrosDias.Visible = True
        lblRechazoOtros_Dias.Visible = True
        Set rsRO = obRO.ReprogRechazoCombo(ActXCodCta.NroCuenta, Right(cmbRechazoOtros.Text, 1))
        If Not (rsRO.BOF And rsRO.EOF) Then
            txtRechazoOtrosDias.Text = rsRO!nDias
        End If
        
    Else
        txtRechazoOtrosDias.Visible = False
        lblRechazoOtros_Dias.Visible = False
        txtRechazoOtrosDias.Text = 0
    End If
    
Set obRO = Nothing
RSClose rsRO
End Sub

Private Sub txtRechazoOtrosDias_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
End Sub

'Add JOEP20210306 garantia covid
Private Sub CargaGrillaModalidades(ByVal pnConsCod As Integer, ByVal pnFiltro As Integer, Optional ByVal pnOpcion As Integer = -1)
    Dim oDCOMCred As COMDConstantes.DCOMConstantes
    Dim rsCovidOpciones As ADODB.Recordset
    Dim i As Integer
        Set oDCOMCred = New COMDConstantes.DCOMConstantes
        Set rsCovidOpciones = oDCOMCred.RecuperaConstanteReprogaramacion(pnConsCod, pnFiltro, pnOpcion)
            If Not (rsCovidOpciones.BOF And rsCovidOpciones.EOF) Then
                LimpiaFlex fe_Modalidades
                For i = 1 To rsCovidOpciones.RecordCount
                    fe_Modalidades.AdicionaFila
                    fe_Modalidades.TextMatrix(i, 0) = i
                    fe_Modalidades.TextMatrix(i, 1) = rsCovidOpciones!cConsDescripcion
                    fe_Modalidades.TextMatrix(i, 3) = rsCovidOpciones!nConsValor
                    rsCovidOpciones.MoveNext
                Next i
            End If
    Set oDCOMCred = Nothing
    RSClose rsCovidOpciones
End Sub
'Add JOEP20210306 garantia covid

Private Function PdfSimulacionCalend() As Boolean

    Dim RsDatos As ADODB.Recordset, rsCalend As ADODB.Recordset, rsCalendTotal As ADODB.Recordset
    Dim oDoc  As cPDF
    Set oDoc = New cPDF
    
    Dim a As Currency
    Dim nFila As Integer
    Dim i As Integer
    Dim FilaIncr As Integer
    Dim EspTitle As Integer
    
    Set oDCred = New COMDCredito.DCOMCredito
    Set RsDatos = oDCred.ReprogramacionPropuestaDatosClienteSimulCalend(ActXCodCta.NroCuenta, gdFecSis)
    Set rsCalend = oDCred.ReprogramacionPropuestaDatosSimulCalend(ActXCodCta.NroCuenta, gdFecSis)
    Set rsCalendTotal = oDCred.ReprogramacionPropuestaDatosTotalSimulCalend(ActXCodCta.NroCuenta, gdFecSis)
    PdfSimulacionCalend = False
           
    'Creación del Archivo
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Propuesta de Reprogramacion Nº " & ActXCodCta.NroCuenta
    oDoc.Title = "Simulacion de Cronograma Nº " & ActXCodCta.NroCuenta
        
    If Not oDoc.PDFCreate(App.Path & "\Spooler\PropuestaReprogramacionSimulacionCronograma_" & ActXCodCta.NroCuenta & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Function
    End If
    
    'Contenido
    oDoc.Fonts.Add "F1", "Courier New", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Courier New", TrueType, Bold, WinAnsiEncoding
    oDoc.LoadImageFromFile App.Path & "\logo_cmacmaynas.bmp", "Logo"
    
    EspTitle = 20
    nFila = 80
    FilaIncr = 15
    
'Inicio de creacion PDF
    If Not (RsDatos.BOF Or RsDatos.EOF) Then
    
        'Tamaño de hoja A4
        oDoc.NewPage A4_Horizontal
        oDoc.WImage 70, 50, 45, 100, "Logo"
                    ' arriba,left,Abajo,right
        oDoc.WTextBox 20, 40, 560, 790, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        
'************************************ Cabecera ************************************************
        oDoc.WTextBox 30, 50, 10, 790, "DOCUMENTO DE USO INTERNO", "F2", 14, hCenter
        oDoc.WTextBox 40, 300, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
        oDoc.WTextBox 50, 660, 35, 490, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
        
        oDoc.WTextBox nFila, 50, 15, 750, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 50, 10, 790, "PROPUESTA DE REPROGRAMACIÓN - SIMULACIÓN REFERENCIAL ", "F2", 10, hCenter
        
        nFila = nFila + EspTitle
        oDoc.WTextBox nFila, 55, 10, 300, "NOMBRE DE CLIENTE", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 170, 10, 370, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 175, 10, 300, RsDatos!cPersNombre, "F1", 7.5, hLeft
        
        nFila = nFila + FilaIncr
        oDoc.WTextBox nFila, 55, 10, 300, "N° CRÉDITO", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 170, 10, 150, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 175, 10, 300, RsDatos!cCtaCod, "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 340, 10, 300, "ANALISTA", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 440, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 445, 10, 300, RsDatos!cUserAnalista, "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 560, 10, 300, "CUOTA A REPROGRAMAR", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 700, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 705, 10, 300, RsDatos!nCuota, "F1", 7.5, hLeft
        nFila = nFila + FilaIncr
        oDoc.WTextBox nFila, 55, 10, 300, "MONTO DESEMBOLSADO" & "   ", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 170, 10, 150, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 175, 10, 300, RsDatos!cMoneda & Space(1) & Format(RsDatos!nMontoCol, "#,##0.00"), "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 340, 10, 300, "SALDO CAPITAL", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 440, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 445, 10, 300, RsDatos!cMoneda & Space(1) & Format(RsDatos!nSaldoReprog, "#,##0.00"), "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 560, 10, 300, "CUOTAS PENDIENTES" & "   ", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 700, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 705, 10, 300, RsDatos!nCuotasReprog, "F1", 7.5, hLeft
        nFila = nFila + FilaIncr
        oDoc.WTextBox nFila, 55, 10, 300, "TIPO DE PRODUCTO", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 170, 10, 150, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 175, 10, 300, RsDatos!cTpoProdCod, "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 340, 10, 300, "FECHA DE VENCIMIENTO", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 440, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 445, 10, 300, RsDatos!dFecCuotaVenc, "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 560, 10, 300, "FACILIDAD DE REPROGRAMACIÓN", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 700, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 705, 10, 300, RsDatos!cModalidaReprog, "F1", 7.5, hLeft
        nFila = nFila + FilaIncr
        oDoc.WTextBox nFila, 55, 10, 300, "TIPO DE CRÉDITO", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 170, 10, 150, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 175, 10, 300, RsDatos!cTpoCredCod, "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 340, 10, 300, "TEA", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 440, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 445, 10, 300, Format(RsDatos!nTEA, "#,#0.00") & "%", "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 560, 10, 300, "TEA PROPUESTO", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 700, 10, 100, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 705, 10, 300, Format(RsDatos!nTEA_New, "#,#0.00") & "%", "F1", 7.5, hLeft
        
        nFila = nFila + FilaIncr
        oDoc.WTextBox nFila, 55, 10, 300, "LIQUIDACIÓN", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 50, 15, 275, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack

        nFila = nFila + FilaIncr
        oDoc.WTextBox nFila, 55, 10, 50, "INT. COMP.", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 50, 20, 55, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 110, 10, 50, "INT. GRACIA", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 105, 20, 55, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 165, 10, 50, "MORA", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 160, 20, 55, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 220, 10, 50, "SEG. DESG", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 215, 20, 55, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 275, 10, 50, "SEG. INC.", "F2", 7.5, hLeft
        oDoc.WTextBox nFila, 270, 20, 55, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        
        nFila = nFila + FilaIncr + 5
        oDoc.WTextBox nFila, 55, 10, 300, Format(RsDatos!LiquiIntCompFecha, "#,#0.00"), "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 50, 15, 55, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 110, 10, 300, Format(RsDatos!LiquiIntGraFecha, "#,#0.00"), "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 105, 15, 55, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 165, 10, 300, Format(RsDatos!LiquiMora, "#,#0.00"), "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 160, 15, 55, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 220, 10, 300, Format(RsDatos!LiquiSegDesgAnt, "#,#0.00"), "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 215, 15, 55, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        oDoc.WTextBox nFila, 275, 10, 300, Format(RsDatos!LiquiSegIncAnt, "#,#0.00"), "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 270, 15, 55, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
        
'************************************ Cabecera ************************************************

        nFila = nFila + FilaIncr + 10
        
'************************************ Calendario ************************************************
        If Not (rsCalend.BOF Or rsCalend.EOF) Then
        
            oDoc.WTextBox nFila, 50, 15, 750, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 50, 10, 750, " SIMULACIÓN REFERENCIAL ", "F2", 10, hCenter
            
            nFila = nFila + FilaIncr + 10
            
            oDoc.WTextBox nFila, 50, 20, 35, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 55, 10, 25, "N° CUOTA", "F2", 7.5, hLeft
            
            oDoc.WTextBox nFila, 85, 20, 55, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 90, 10, 55, "ESTADO", "F2", 7.5, hLeft
                        
            oDoc.WTextBox nFila, 140, 20, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 145, 10, 55, "FECHA DE VENCIMIENTO", "F2", 7.5, hLeft
            
            oDoc.WTextBox nFila, 200, 20, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 205, 10, 40, "MONTO CUOTA", "F2", 7.5, hLeft
                        
            oDoc.WTextBox nFila, 260, 20, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 265, 10, 40, "CAPITAL", "F2", 7.5, hLeft
            
            oDoc.WTextBox nFila, 320, 20, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 325, 10, 40, "INT. COMP.", "F2", 7.5, hLeft
            oDoc.WTextBox nFila, 380, 20, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 385, 10, 40, "INT. GRACIA.", "F2", 7.5, hLeft
            
            oDoc.WTextBox nFila, 440, 20, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 445, 10, 40, "SEG. DES", "F2", 7.5, hLeft
            oDoc.WTextBox nFila, 500, 20, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 505, 10, 60, "SEG. INC.", "F2", 7.5, hLeft
            oDoc.WTextBox nFila, 560, 20, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 565, 10, 40, "GASTOS Y COMISION", "F2", 7.5, hLeft
            
            oDoc.WTextBox nFila, 620, 20, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 625, 10, 60, "SALDO CAP.", "F2", 7.5, hLeft
                        
            nFila = nFila + FilaIncr + 10
            
            For i = 1 To rsCalend.RecordCount
                oDoc.WTextBox nFila, 50, 15, 35, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                oDoc.WTextBox nFila, 60, 10, 300, rsCalend!nCuota, "F1", 7.5, hLeft
                
                oDoc.WTextBox nFila, 85, 15, 55, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                oDoc.WTextBox nFila, 90, 10, 55, rsCalend!nEstado, "F1", 7.5, hLeft
                                
                oDoc.WTextBox nFila, 140, 15, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                oDoc.WTextBox nFila, 145, 10, 300, rsCalend!dVenc, "F1", 7.5, hLeft
                
                oDoc.WTextBox nFila, 200, 15, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                oDoc.WTextBox nFila, 205, 10, 300, Format(rsCalend!nMonCuota, "#,#00.00"), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 260, 15, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                oDoc.WTextBox nFila, 265, 10, 300, Format(rsCalend!nCapital, "#,#00.00"), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 320, 15, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                oDoc.WTextBox nFila, 325, 10, 300, Format(rsCalend!nIntComp, "#,#00.00"), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 380, 15, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                oDoc.WTextBox nFila, 385, 10, 300, Format(rsCalend!nIntGracia, "#,#00.00"), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 440, 15, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                oDoc.WTextBox nFila, 445, 10, 300, Format(rsCalend!nSegDesg, "#,#00.00"), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 500, 15, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                oDoc.WTextBox nFila, 505, 10, 300, Format(rsCalend!nSegInc, "#,#00.00"), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 560, 15, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                oDoc.WTextBox nFila, 565, 10, 300, Format(rsCalend!nGasComi, "#,#00.00"), "F1", 7.5, hLeft

                oDoc.WTextBox nFila, 620, 15, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                oDoc.WTextBox nFila, 625, 10, 300, Format(rsCalend!nSaldoCap, "#,#00.00"), "F1", 7.5, hLeft
                
                If nFila >= 520 And nFila <= 530 Then
                
                    oDoc.WTextBox 555, 50, 10, 700, "* Monto de cuota referencial, sujeto a variación", "F2", 7.5, hLeft
                    oDoc.WTextBox 565, 50, 10, 700, "* Operaciones afectas a ITF", "F2", 7.5, hLeft
                
                    oDoc.NewPage A4_Horizontal
                    oDoc.WImage 70, 50, 45, 100, "Logo"
                    oDoc.WTextBox 20, 40, 560, 790, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                    oDoc.WTextBox 30, 50, 10, 790, "DOCUMENTO DE USO INTERNO", "F2", 14, hCenter
                    oDoc.WTextBox 40, 300, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
                    oDoc.WTextBox 50, 660, 35, 490, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
                    oDoc.WTextBox 80, 50, 13, 750, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                    oDoc.WTextBox 80, 50, 10, 750, " SIMULACIÓN REFERENCIAL ", "F2", 10, hCenter
                                                            
                    oDoc.WTextBox 95, 50, 20, 35, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                    oDoc.WTextBox 95, 55, 10, 25, "N° CUOTA", "F2", 7.5, hLeft
                    oDoc.WTextBox 95, 85, 20, 55, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                    oDoc.WTextBox 95, 90, 10, 55, "ESTADO", "F2", 7.5, hLeft
                    oDoc.WTextBox 95, 140, 20, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                    oDoc.WTextBox 95, 145, 10, 55, "FECHA DE VENCIMIENTO", "F2", 7.5, hLeft
                    oDoc.WTextBox 95, 200, 20, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                    oDoc.WTextBox 95, 205, 10, 40, "MONTO CUOTA", "F2", 7.5, hLeft
                    oDoc.WTextBox 95, 260, 20, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                    oDoc.WTextBox 95, 265, 10, 40, "CAPITAL", "F2", 7.5, hLeft
                    oDoc.WTextBox 95, 320, 20, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                    oDoc.WTextBox 95, 325, 10, 40, "INT. COMP.", "F2", 7.5, hLeft
                    oDoc.WTextBox 95, 380, 20, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                    oDoc.WTextBox 95, 385, 10, 40, "INT. GRACIA.", "F2", 7.5, hLeft
                    oDoc.WTextBox 95, 440, 20, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                    oDoc.WTextBox 95, 445, 10, 40, "SEG. DES", "F2", 7.5, hLeft
                    oDoc.WTextBox 95, 500, 20, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                    oDoc.WTextBox 95, 505, 10, 60, "SEG. INC.", "F2", 7.5, hLeft
                    oDoc.WTextBox 95, 560, 20, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                    oDoc.WTextBox 95, 565, 10, 40, "GASTOS Y COMISION", "F2", 7.5, hLeft
                    oDoc.WTextBox 95, 620, 20, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                    oDoc.WTextBox 95, 625, 10, 60, "SALDO CAP.", "F2", 7.5, hLeft

                    nFila = 100
                End If
                          
                nFila = nFila + FilaIncr
                rsCalend.MoveNext
            Next i
        End If
        
        If Not (rsCalendTotal.BOF And rsCalendTotal.EOF) Then
            oDoc.WTextBox nFila, 145, 10, 40, "TOTALES", "F2", 7.5, hLeft
            
            oDoc.WTextBox nFila, 200, 15, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 205, 10, 300, Format(rsCalendTotal!nTotalMonCuota, "#,#00.00"), "F1", 7.5, hLeft
            oDoc.WTextBox nFila, 260, 15, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 265, 10, 300, Format(rsCalendTotal!nTotalCapital, "#,#00.00"), "F1", 7.5, hLeft
            oDoc.WTextBox nFila, 320, 15, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 325, 10, 300, Format(rsCalendTotal!nTotalIntComp, "#,#00.00"), "F1", 7.5, hLeft
            oDoc.WTextBox nFila, 380, 15, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 385, 10, 300, Format(rsCalendTotal!nTotalIntGracia, "#,#00.00"), "F1", 7.5, hLeft
            oDoc.WTextBox nFila, 440, 15, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 445, 10, 300, Format(rsCalendTotal!nTotalnSegDesg, "#,#00.00"), "F1", 7.5, hLeft
            oDoc.WTextBox nFila, 500, 15, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 505, 10, 300, Format(rsCalendTotal!nTotalnSegInc, "#,#00.00"), "F1", 7.5, hLeft
            oDoc.WTextBox nFila, 560, 15, 60, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
            oDoc.WTextBox nFila, 565, 10, 300, Format(rsCalendTotal!nTotalnGasCom, "#,#00.00"), "F1", 7.5, hLeft
        End If
'************************************ Calendario ************************************************

'************************************ pie de pagina ************************************************
oDoc.WTextBox 555, 50, 10, 700, "* Monto de cuota referencial, sujeto a variación", "F2", 7.5, hLeft
oDoc.WTextBox 565, 50, 10, 700, "* Operaciones afectas a ITF", "F2", 7.5, hLeft
'************************************ pie de pagina ************************************************
        oDoc.PDFClose
        oDoc.Show
        PdfSimulacionCalend = True
        
    Else
        MsgBox "Los Datos de la propuesta del Crédito no han sido Registrados Correctamente", vbInformation, "Aviso"
    End If
    
    Set oDCred = Nothing
    Set RsDatos = Nothing
    Set rsCalend = Nothing
    
End Function
