VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredAlertaTermpranaConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de los limites de ratios de Alertas Tempranas"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15435
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCredAlertaTempranaConfig.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   15435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   360
      Left            =   14040
      TabIndex        =   14
      Top             =   6960
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   12720
      TabIndex        =   13
      Top             =   6960
      Width           =   1200
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Height          =   360
      Left            =   1440
      TabIndex        =   22
      Top             =   6960
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   7395
      Left            =   -720
      TabIndex        =   0
      Top             =   0
      Width           =   16160
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   360
         Left            =   840
         TabIndex        =   12
         Top             =   6960
         Width           =   1200
      End
      Begin TabDlg.SSTab SSTabDatos 
         Height          =   6735
         Left            =   720
         TabIndex        =   21
         Top             =   120
         Width           =   15375
         _ExtentX        =   27120
         _ExtentY        =   11880
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Configuración Ratio - Capacidad de Pago"
         TabPicture(0)   =   "frmCredAlertaTempranaConfig.frx":030A
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "frameProducto"
         Tab(0).Control(1)=   "Frame2"
         Tab(0).Control(2)=   "Frame7"
         Tab(0).Control(3)=   "frameCredito"
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Configuracion Ratio - Ratios Otros"
         TabPicture(1)   =   "frmCredAlertaTempranaConfig.frx":0326
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame8"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.Frame frameCredito 
            Caption         =   "Tipo Crédito / Sub Tipo Crédito :"
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
            Height          =   860
            Left            =   -68760
            TabIndex        =   46
            Top             =   360
            Width           =   6040
            Begin VB.ComboBox cboSubTipo 
               Height          =   315
               Left            =   3120
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Top             =   480
               Width           =   2850
            End
            Begin VB.ComboBox cboTipoCredito 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   3
               Top             =   480
               Width           =   2850
            End
            Begin VB.Label lblTipoCredito 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SubTipoCredito:"
               Height          =   195
               Left            =   3120
               TabIndex        =   48
               Top             =   260
               Width           =   1155
            End
            Begin VB.Label lblCredito 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo Credito:"
               Height          =   195
               Left            =   120
               TabIndex        =   47
               Top             =   260
               Width           =   930
            End
         End
         Begin VB.Frame Frame8 
            Height          =   6135
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   15135
            Begin VB.Frame Frame10 
               Height          =   1215
               Left            =   12240
               TabIndex        =   49
               Top             =   360
               Width           =   2775
               Begin VB.CommandButton cmdCancelarRatioOtro 
                  Caption         =   "Cancelar"
                  Height          =   360
                  Left            =   1440
                  TabIndex        =   51
                  Top             =   480
                  Width           =   1200
               End
               Begin VB.CommandButton cmdAceptarRatioOtro 
                  Caption         =   "Aceptar"
                  Height          =   360
                  Left            =   120
                  TabIndex        =   20
                  Top             =   480
                  Width           =   1200
               End
               Begin VB.CommandButton cmdNuevoRatioOtro 
                  Caption         =   "Nuevo"
                  Height          =   360
                  Left            =   360
                  TabIndex        =   50
                  Top             =   480
                  Width           =   2070
               End
            End
            Begin VB.Frame frameListadoRatioOtro 
               Caption         =   "Listado de ratios"
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
               Height          =   4335
               Left            =   120
               TabIndex        =   41
               Top             =   1680
               Width           =   14895
               Begin SICMACT.FlexEdit feListaRatiosOtros 
                  Height          =   3855
                  Left            =   120
                  TabIndex        =   44
                  Top             =   360
                  Width           =   14655
                  _ExtentX        =   25850
                  _ExtentY        =   6800
                  Cols0           =   9
                  HighLight       =   1
                  EncabezadosNombres=   "-DESCRIPCION-FORMULA-VALOR-UNIDAD MEDIDA-nCodAlerta-nUnidad-nEstado-Aux"
                  EncabezadosAnchos=   "300-4200-6700-1400-1950-0-0-0-0"
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
                  ListaControles  =   "0-0-0-0-0-0-0-0-0"
                  BackColorControl=   -2147483643
                  BackColorControl=   -2147483643
                  BackColorControl=   -2147483643
                  EncabezadosAlineacion=   "C-L-L-R-C-C-C-C-C"
                  FormatosEdit    =   "0-0-0-2-0-0-0-0-0"
                  SelectionMode   =   1
                  lbEditarFlex    =   -1  'True
                  lbUltimaInstancia=   -1  'True
                  TipoBusqueda    =   3
                  lbBuscaDuplicadoText=   -1  'True
                  ColWidth0       =   300
                  RowHeight0      =   300
               End
            End
            Begin VB.Frame frameConfigRatioOtro 
               Caption         =   "Configuración límites:"
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
               Height          =   1335
               Left            =   120
               TabIndex        =   37
               Top             =   240
               Width           =   12015
               Begin VB.ComboBox cboEstadoRatioOtro 
                  Height          =   315
                  ItemData        =   "frmCredAlertaTempranaConfig.frx":0342
                  Left            =   4920
                  List            =   "frmCredAlertaTempranaConfig.frx":0344
                  Style           =   2  'Dropdown List
                  TabIndex        =   19
                  Top             =   920
                  Width           =   1845
               End
               Begin VB.TextBox txtNombreRatioOtro 
                  Height          =   320
                  Left            =   120
                  TabIndex        =   15
                  Top             =   360
                  Width           =   4455
               End
               Begin VB.ComboBox cboUnidadRatioOtro 
                  Height          =   315
                  ItemData        =   "frmCredAlertaTempranaConfig.frx":0346
                  Left            =   2400
                  List            =   "frmCredAlertaTempranaConfig.frx":0348
                  Style           =   2  'Dropdown List
                  TabIndex        =   18
                  Top             =   920
                  Width           =   1965
               End
               Begin VB.TextBox txtValorRatioOtro 
                  Alignment       =   1  'Right Justify
                  Height          =   320
                  Left            =   120
                  TabIndex        =   17
                  Text            =   "0"
                  Top             =   920
                  Width           =   1455
               End
               Begin VB.TextBox txtFormulaRatioOtro 
                  Height          =   320
                  Left            =   4920
                  TabIndex        =   16
                  Top             =   360
                  Width           =   6975
               End
               Begin VB.Label lblRatioOtro 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Código Alerta:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   -1  'True
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   10560
                  TabIndex        =   61
                  Top             =   960
                  Width           =   1035
               End
               Begin VB.Label lblCodRatioRatioOtro 
                  Caption         =   "1"
                  Height          =   255
                  Left            =   11760
                  TabIndex        =   59
                  Top             =   960
                  Width           =   135
               End
               Begin VB.Label Label2 
                  Caption         =   "="
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   4680
                  TabIndex        =   54
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   135
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ratio:"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   43
                  Top             =   165
                  Width           =   435
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Estado:"
                  Height          =   195
                  Left            =   4920
                  TabIndex        =   42
                  Top             =   720
                  Width           =   555
               End
               Begin VB.Label Label21 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unidad Medida:"
                  Height          =   195
                  Left            =   2400
                  TabIndex        =   40
                  Top             =   720
                  Width           =   1110
               End
               Begin VB.Label Label14 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Valor Limite:"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   39
                  Top             =   720
                  Width           =   870
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Fórmula:"
                  Height          =   195
                  Left            =   5025
                  TabIndex        =   38
                  Top             =   165
                  Width           =   630
               End
            End
         End
         Begin VB.Frame Frame7 
            Height          =   735
            Left            =   -62640
            TabIndex        =   34
            Top             =   405
            Width           =   2895
            Begin VB.CommandButton cmdMostrar 
               Caption         =   "Mostrar"
               Height          =   375
               Left            =   360
               TabIndex        =   35
               Top             =   240
               Width           =   2180
            End
         End
         Begin VB.Frame Frame2 
            Height          =   5415
            Left            =   -74880
            TabIndex        =   26
            Top             =   1200
            Width           =   15200
            Begin VB.Frame Frame11 
               Height          =   1335
               Left            =   12240
               TabIndex        =   55
               Top             =   120
               Width           =   2775
               Begin VB.CommandButton cmdCancelarCapaPago 
                  Caption         =   "Cancelar"
                  Height          =   360
                  Left            =   1440
                  TabIndex        =   58
                  Top             =   540
                  Width           =   1200
               End
               Begin VB.CommandButton cmdEliminarCapaPago 
                  Caption         =   "Eliminar"
                  Height          =   360
                  Left            =   1440
                  TabIndex        =   57
                  Top             =   540
                  Visible         =   0   'False
                  Width           =   1200
               End
               Begin VB.CommandButton cmdAceptarCapaPago 
                  Caption         =   "Aceptar"
                  Height          =   360
                  Left            =   120
                  TabIndex        =   11
                  Top             =   540
                  Width           =   1200
               End
               Begin VB.CommandButton cmdNuevoCapaPaga 
                  Caption         =   "Nuevo"
                  Height          =   360
                  Left            =   120
                  TabIndex        =   56
                  Top             =   540
                  Width           =   1200
               End
            End
            Begin VB.Frame frameConfigCapaPago 
               Caption         =   "Configuración limites"
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
               Height          =   1335
               Left            =   120
               TabIndex        =   29
               Top             =   120
               Width           =   12015
               Begin VB.TextBox txtFormulaCapaPago 
                  Height          =   320
                  Left            =   4200
                  TabIndex        =   6
                  Top             =   360
                  Width           =   7695
               End
               Begin VB.TextBox txtNombreCapaPago 
                  Height          =   320
                  Left            =   120
                  TabIndex        =   5
                  Top             =   360
                  Width           =   3855
               End
               Begin VB.ComboBox cboUnidadCapaPago 
                  Height          =   315
                  ItemData        =   "frmCredAlertaTempranaConfig.frx":034A
                  Left            =   4200
                  List            =   "frmCredAlertaTempranaConfig.frx":034C
                  Style           =   2  'Dropdown List
                  TabIndex        =   9
                  Top             =   915
                  Width           =   2415
               End
               Begin VB.ComboBox cboTipoCliente 
                  Height          =   315
                  ItemData        =   "frmCredAlertaTempranaConfig.frx":034E
                  Left            =   120
                  List            =   "frmCredAlertaTempranaConfig.frx":0350
                  Style           =   2  'Dropdown List
                  TabIndex        =   7
                  Top             =   915
                  Width           =   2415
               End
               Begin VB.TextBox txtValorCapaPago 
                  Alignment       =   1  'Right Justify
                  Height          =   320
                  Left            =   2760
                  TabIndex        =   8
                  Text            =   "0"
                  Top             =   915
                  Width           =   1215
               End
               Begin VB.ComboBox cboEstadoCapaPago 
                  Height          =   315
                  ItemData        =   "frmCredAlertaTempranaConfig.frx":0352
                  Left            =   6840
                  List            =   "frmCredAlertaTempranaConfig.frx":0354
                  Style           =   2  'Dropdown List
                  TabIndex        =   10
                  Top             =   915
                  Width           =   1820
               End
               Begin VB.Label lblCapaPago 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Código Alerta:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   -1  'True
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   10440
                  TabIndex        =   60
                  Top             =   960
                  Width           =   1035
               End
               Begin VB.Label Label1 
                  Caption         =   "="
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   4020
                  TabIndex        =   53
                  Top             =   400
                  Visible         =   0   'False
                  Width           =   135
               End
               Begin VB.Label lblFormulaCapaPago 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Fórmula:"
                  Height          =   195
                  Left            =   4305
                  TabIndex        =   52
                  Top             =   165
                  Width           =   630
               End
               Begin VB.Label lblCodRatioCapaPago 
                  Caption         =   "1"
                  Height          =   255
                  Left            =   11640
                  TabIndex        =   45
                  Top             =   960
                  Width           =   135
               End
               Begin VB.Label lblUnidadCapaPago 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unidad Medida:"
                  Height          =   195
                  Left            =   4320
                  TabIndex        =   33
                  Top             =   735
                  Width           =   1110
               End
               Begin VB.Label lblTipoCreditoCapaPago 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tipo Cliente:"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   32
                  Top             =   735
                  Width           =   900
               End
               Begin VB.Label lblValorCapaPago 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Valor Limite:"
                  Height          =   195
                  Left            =   2835
                  TabIndex        =   31
                  Top             =   735
                  Width           =   870
               End
               Begin VB.Label lblEstadoCapaPago 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Estado:"
                  Height          =   195
                  Left            =   6960
                  TabIndex        =   30
                  Top             =   735
                  Width           =   555
               End
            End
            Begin VB.Frame frameListadoCapaPago 
               Caption         =   "Listado de ratios:"
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
               Height          =   3735
               Left            =   120
               TabIndex        =   27
               Top             =   1560
               Width           =   15000
               Begin SICMACT.FlexEdit feListaCapaPago 
                  Height          =   3405
                  Left            =   120
                  TabIndex        =   28
                  Top             =   240
                  Width           =   14790
                  _ExtentX        =   26088
                  _ExtentY        =   6006
                  Cols0           =   17
                  HighLight       =   1
                  EncabezadosNombres=   $"frmCredAlertaTempranaConfig.frx":0356
                  EncabezadosAnchos=   "300-1900-3300-2000-3400-1500-1100-1200-0-0-0-0-0-0-0-0-0"
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
                  ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
                  ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
                  BackColorControl=   -2147483643
                  BackColorControl=   -2147483643
                  BackColorControl=   -2147483643
                  EncabezadosAlineacion=   "C-L-L-L-L-L-R-C-C-C-C-C-C-C-C-C-C"
                  FormatosEdit    =   "0-0-0-0-0-0-2-0-0-0-0-0-0-0-0-0-0"
                  SelectionMode   =   1
                  lbEditarFlex    =   -1  'True
                  lbUltimaInstancia=   -1  'True
                  TipoBusqueda    =   3
                  lbBuscaDuplicadoText=   -1  'True
                  ColWidth0       =   300
                  RowHeight0      =   300
               End
            End
         End
         Begin VB.Frame frameProducto 
            Caption         =   "Producto / SubProducto :"
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
            Height          =   860
            Left            =   -74880
            TabIndex        =   23
            Top             =   360
            Width           =   6040
            Begin VB.ComboBox cboProducto 
               Height          =   315
               ItemData        =   "frmCredAlertaTempranaConfig.frx":0406
               Left            =   120
               List            =   "frmCredAlertaTempranaConfig.frx":0408
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   450
               Width           =   2850
            End
            Begin VB.ComboBox cboSubProducto 
               Height          =   315
               Left            =   3120
               Style           =   2  'Dropdown List
               TabIndex        =   2
               Top             =   450
               Width           =   2850
            End
            Begin VB.Label lblProducto 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Producto:"
               Height          =   195
               Left            =   120
               TabIndex        =   25
               Top             =   255
               Width           =   705
            End
            Begin VB.Label lblSubProducto 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SubProducto:"
               Height          =   195
               Left            =   3120
               TabIndex        =   24
               Top             =   255
               Width           =   975
            End
         End
      End
   End
End
Attribute VB_Name = "frmCredAlertaTermpranaConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************************
'** Nombre      : frmCredAlertaTempranaConfig                                               ****
'** Descripción : Muestra el detalle de las alertas tempranas por crédito                   ****
'** Referencia  : ERS001-2017 - Metodología para monitoreo de Alertas Tempranas de Crédito  ****
'** Creado por  : LUCV, 20170215 15:51:01 PM                                                ****
'***********************************************************************************************
Option Explicit
Dim lbEditar As Boolean
Dim lnCantFilaCapaPago As Integer
Dim lnCantFilaRatioOtro As Integer

Private Sub Form_Load()
    Call CargaControles
    Call VisualizaControles(True, True, True)
    lbEditar = False
    SSTabDatos.Tab = 0
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 86 And Shift = 2 Then
        KeyCode = 10
    End If
    If KeyCode = 113 And Shift = 0 Then
        KeyCode = 10
    End If
    If KeyCode = 27 And Shift = 0 Then
        Unload Me
    End If
End Sub

Private Sub cmdMostrar_Click()
    Dim i As Integer
    i = 0
    If Trim(Me.cboProducto.Text) = "" Then
        MsgBox "Debe seleccionar un producto crediticio", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If Trim(Me.cboSubProducto.Text) = "" Then
        MsgBox "Debe seleccionar un Sub producto", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If Trim(Me.cboTipoCredito.Text) = "" Then
        MsgBox "Debe seleccionar un Tipo de Crédito", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If Trim(Me.cboSubTipo.Text) = "" Then
        MsgBox "Debe seleccionar un Sub Tipo de Crédito", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Call CargarDatos
    Call HabilitaControles(, , , , , , , , True, True, True, , , , True)
    Call VisualizaControles(True, True, True)
End Sub
Private Sub cmdGrabar_Click()
    Dim sMsj As String
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
    Dim GrabarDatos As Boolean
    Dim rsDatosCapaPago As ADODB.Recordset
    Dim rsDatosRatiosOtros As ADODB.Recordset

    Dim sNombreRatioCapaPago As String
    Dim sFormulaRatioCapaPago As String
    
    sNombreRatioCapaPago = Trim(txtNombreCapaPago.Text)
    sFormulaRatioCapaPago = Trim(txtFormulaCapaPago.Text)
    
    Dim lsMovNro As String
    lsMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)

    Set rsDatosCapaPago = IIf(feListaCapaPago.rows - 1 > 0, feListaCapaPago.GetRsNew(), Nothing)
    Set rsDatosRatiosOtros = IIf(feListaRatiosOtros.rows - 1 > 0, feListaRatiosOtros.GetRsNew(), Nothing)
    If ValidaDatos Then
        If MsgBox("Los Datos seran Grabados, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
            GrabarDatos = oNCOMFormatosEval.GrabarAlertaTempranaConfig(lsMovNro, rsDatosCapaPago, rsDatosRatiosOtros, sNombreRatioCapaPago, sFormulaRatioCapaPago)
        If GrabarDatos Then
            MsgBox "Los datos se grabaron correctamente", vbInformation, "Aviso"
            Call CargaControles
            Call LimpiarDatos
            Call VisualizaControles(True, True, True)
            SSTabDatos.Tab = 0
        Else
            MsgBox "Hubo errores al grabar la información", vbError, "Error"
            Call CargarDatos
            Call HabilitaControles(, , , , , , , , True, True, True)
            Call VisualizaControles(True, True)
        End If
    End If
End Sub
Private Sub cmdEditar_Click()
    Call HabilitaControles(, , , True, True, , , True, , True, True, True, , , True)
    Call VisualizaControles(True, True, True)
    cmdNuevoCapaPaga.SetFocus
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub cmdCancelar_Click()
    Call CargaControles
    Call VisualizaControles(True, True, True)
    lbEditar = False
    Call LimpiarDatos
End Sub
Private Sub cmdCancelarRatioOtro_Click()
    Dim nCodAlertaRatioOtro As Integer
    lnCantFilaRatioOtro = lnCantFilaRatioOtro - 1
    nCodAlertaRatioOtro = lnCantFilaRatioOtro
    lblCodRatioRatioOtro.Caption = nCodAlertaRatioOtro

    Call HabilitaControles(, , , True, True, , , True, , True, True, True, , , True)
    Call VisualizaControles(True, True, True)
End Sub
Private Sub cmdNuevoCapaPaga_Click()
   lbEditar = False
   Call HabilitaControles(, True, , , , True, True, True, , True, True, , , , True)
   Call VisualizaControles(, , True, True, True, , , True)
   txtNombreCapaPago.Enabled = False
End Sub
Private Sub cmdNuevoRatioOtro_Click()
    Dim nCodAlertaRatioOtro As Integer
    lnCantFilaRatioOtro = lnCantFilaRatioOtro + 1
    nCodAlertaRatioOtro = lnCantFilaRatioOtro
    lblCodRatioRatioOtro.Caption = nCodAlertaRatioOtro

    Call HabilitaControles(, True, , , , , , True, , True, True, , True, True, True)
    Call VisualizaControles(True, True, , , , True, True, True)
    lbEditar = False
    txtNombreRatioOtro.Text = ""
    txtFormulaRatioOtro.Text = ""
    txtNombreRatioOtro.SetFocus
End Sub
Private Sub cmdEliminarCapaPago_Click()
    If MsgBox("Esta Seguro de Eliminar Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Me.feListaCapaPago.EliminaFila (feListaCapaPago.row)
        Call HabilitaControles(, , , True, True, , True, True, , True, True, , , , True)
        Call VisualizaControles(True, True, True)
    End If
End Sub
Private Sub cmdAceptarCapaPago_Click()
    Dim i As Integer
    Dim nCodAlerta As Integer
    Dim sMsj As String

    If txtFormulaCapaPago.Text = "" Then
        MsgBox "Por favor ingresar el detalle de la fórmula del ratio", vbInformation, "Aviso"
        Exit Sub
    End If
    
    sMsj = ValidaDatosCapaPago
    nCodAlerta = CInt(lblCodRatioCapaPago.Caption)
    
    If lbEditar = True Then
        sMsj = ""
    End If
    If sMsj <> "" Then
        MsgBox sMsj, vbInformation, "Alerta"
        Exit Sub
    End If
    
    Call AgregaConfigCapaPago(Trim(Left(Me.cboProducto.Text, Len(Me.cboProducto.Text) - 3)), _
                              Trim(Left(Me.cboSubProducto.Text, Len(Me.cboSubProducto.Text) - 3)), _
                              Trim(Left(Me.cboTipoCredito.Text, Len(Me.cboTipoCredito.Text) - 3)), _
                              Trim(Left(Me.cboSubTipo.Text, Len(Me.cboSubTipo.Text) - 3)), _
                              Trim(Left(Me.cboTipoCliente.Text, Len(Me.cboTipoCliente.Text) - 1)), _
                              CDbl(Me.txtValorCapaPago.Text), _
                              Trim(Left(Me.cboUnidadCapaPago.Text, Len(Me.cboUnidadCapaPago.Text) - 1)), _
                              Right(Me.cboProducto.Text, 3), _
                              Right(Me.cboSubProducto.Text, 3), _
                              Right(Me.cboTipoCredito.Text, 3), _
                              Right(Me.cboSubTipo.Text, 3), _
                              Right(Me.cboTipoCliente.Text, 1), _
                              Right(Me.cboUnidadCapaPago.Text, 1), _
                              Right(Me.cboEstadoCapaPago.Text, 1), _
                              nCodAlerta)
                              
    Call HabilitaControles(False, , , True, True, True, True, True, , True, True, True, , , True)
    Call VisualizaControles(True, True, True)
End Sub
Private Sub cmdAceptarRatioOtro_Click()
    Dim i As Integer
    Dim sMsj As String
    Dim nCodAlerta As Integer
    sMsj = ValidaDatosRatio
    nCodAlerta = CInt(lblCodRatioRatioOtro.Caption)
    
    If (Me.txtNombreRatioOtro.Text = "") Or (Me.txtFormulaRatioOtro.Text = "") Then
      MsgBox "Por favor ingresar el nombre o detalle de la fórmula del ratio", vbInformation, "Aviso"
      Exit Sub
    End If
    
    If lbEditar = True Then
        sMsj = ""
    End If
    If sMsj <> "" Then
        MsgBox sMsj, vbInformation, "Alerta"
        Exit Sub
    End If
    
    Call AgregaConfigRatioOtro(Trim(Me.txtNombreRatioOtro.Text), _
                              Trim(Me.txtFormulaRatioOtro.Text), _
                              CDbl(Me.txtValorRatioOtro.Text), _
                              Trim(Left(Me.cboUnidadRatioOtro.Text, Len(Me.cboUnidadRatioOtro.Text) - 1)), _
                              nCodAlerta, _
                              Right(Me.cboUnidadRatioOtro.Text, 1), _
                              Right(Me.cboEstadoRatioOtro.Text, 1))
                              
    Call HabilitaControles(False, , , True, True, , , True, , True, True, True, , , True)
    Call VisualizaControles(True, True, True)
End Sub
Private Sub cmdCancelarCapaPago_Click()
    Call HabilitaControles(, , , True, True, , True, True, , True, True, True, , , True)
    Call VisualizaControles(True, True, True)
End Sub
Private Sub feListaCapaPago_DblClick()
    lbEditar = True
    Dim lnNroFilasCapaPago As String
    lnNroFilasCapaPago = feListaCapaPago.TextMatrix(feListaCapaPago.row, 15)
    If lnNroFilasCapaPago <> "" Then
        Call EditarCapaPago(feListaCapaPago.TextMatrix(feListaCapaPago.row, 15), feListaCapaPago.TextMatrix(feListaCapaPago.row, 12), CDbl(feListaCapaPago.TextMatrix(feListaCapaPago.row, 6)), feListaCapaPago.TextMatrix(feListaCapaPago.row, 13), feListaCapaPago.TextMatrix(feListaCapaPago.row, 14))
         Call HabilitaControles(, True, , , , True, True, True, , True, True, , , , True)
        txtFormulaCapaPago.SetFocus
    Else
        MsgBox "No existen registros para visualizar o editar", vbInformation, "Aviso"
         Call HabilitaControles(, , , True, True, , , True, , True, True, True, , , True)
    End If
    
    txtNombreCapaPago.Enabled = False

End Sub
Private Sub feListaRatiosOtros_DblClick()
    lbEditar = True
    Dim lnNroFilasRatioOtro As String
    lnNroFilasRatioOtro = feListaRatiosOtros.TextMatrix(feListaRatiosOtros.row, 5)
    If lnNroFilasRatioOtro <> "" Then
            Call EditarRatioOtro(feListaRatiosOtros.TextMatrix(feListaRatiosOtros.row, 1), _
                                feListaRatiosOtros.TextMatrix(feListaRatiosOtros.row, 2), _
                                CDbl(feListaRatiosOtros.TextMatrix(feListaRatiosOtros.row, 3)), _
                                feListaRatiosOtros.TextMatrix(feListaRatiosOtros.row, 5), _
                                feListaRatiosOtros.TextMatrix(feListaRatiosOtros.row, 6), _
                                feListaRatiosOtros.TextMatrix(feListaRatiosOtros.row, 7))
        Call HabilitaControles(, True, , , , , , True, , True, True, , True, True, True)
    Else
        MsgBox "No existen registros para visualizar o editar", vbInformation, "Aviso"
        Call HabilitaControles(, , , True, True, , , True, , True, True, True, , , True)
    End If
    
    
    txtNombreCapaPago.Enabled = False
End Sub
Private Sub CargaControles()
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Dim oDCOMConstantes As COMDConstantes.DCOMConstantes
    Set oDCOMConstantes = New COMDConstantes.DCOMConstantes
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    
    Call LimpiaFlex(Me.feListaCapaPago)
    '*** Tipo Producto
    Set oRs = oDCOMFormatosEval.ObtenerProducto(3033)
    Call LlenarCombo(cboProducto, oRs)
    Set oRs = Nothing
    Set oRs = New ADODB.Recordset
    If cboProducto.ListCount > 0 Then
     cboProducto.ListIndex = 0
    End If
    
    '*** Tipo Crédito
    Set oRs = oDCOMFormatosEval.ObtenerProducto(3034)
    Call LlenarCombo(cboTipoCredito, oRs)
    Set oRs = Nothing
    Set oRs = New ADODB.Recordset
    If cboTipoCredito.ListCount > 0 Then
     cboTipoCredito.ListIndex = 0
    End If
    
    '*** Tipo Cliente
    Set oRs = oDCOMConstantes.RecuperaConstantes(3015)
    Call LlenarCombo(cboTipoCliente, oRs)
    Set oRs = Nothing
    Set oRs = New ADODB.Recordset
    If cboTipoCliente.ListCount > 0 Then
     cboTipoCliente.ListIndex = 0
    End If
    
    '*** Estado Limite CapacidadPago
    Set oRs = oDCOMFormatosEval.ObtenerEstadoLimite()
    Call LlenarCombo(cboEstadoCapaPago, oRs)
    Set oRs = Nothing
    Set oRs = New ADODB.Recordset
    If cboEstadoCapaPago.ListCount > 0 Then
     cboEstadoCapaPago.ListIndex = 0
    End If
    
    '*** Estado Limite Ratios Otros
    Set oRs = oDCOMFormatosEval.ObtenerEstadoLimite()
    Call LlenarCombo(cboEstadoRatioOtro, oRs)
    Set oRs = Nothing
    Set oRs = New ADODB.Recordset
    If cboEstadoRatioOtro.ListCount > 0 Then
     cboEstadoRatioOtro.ListIndex = 0
    End If
    
    
    '*** Unidad Medida CapaPago
    Set oRs = oDCOMFormatosEval.ObtenerValoresUnidadMedida()
    Call LlenarCombo(cboUnidadCapaPago, oRs)
    Set oRs = Nothing
    Set oRs = New ADODB.Recordset
    If cboUnidadCapaPago.ListCount > 0 Then
     cboUnidadCapaPago.ListIndex = 0
    End If
    
    '*** Unidad Medida Ratios Otros
    Set oRs = oDCOMFormatosEval.ObtenerValoresUnidadMedida()
    Call LlenarCombo(cboUnidadRatioOtro, oRs)
    Set oRs = Nothing
    Set oRs = New ADODB.Recordset
    If cboUnidadRatioOtro.ListCount > 0 Then
     cboUnidadRatioOtro.ListIndex = 0
    End If
    
    Call CambiaTamañoCombo(cboProducto, 250)
    Call CambiaTamañoCombo(cboSubProducto, 250)
    Call CambiaTamañoCombo(cboTipoCredito, 250)
    Call CambiaTamañoCombo(cboSubTipo, 250)
    Call CambiaTamañoCombo(cboTipoCliente, 150)
    Call HabilitaControles(True, , True, , , , , , , True, True)
End Sub
Private Sub EditarCapaPago(ByVal nCodAlerta As Integer, ByVal nTipoCliente As Integer, ByVal nValorLimite As Double, ByVal nUnidadMedida As Integer, ByVal nEstado As Integer)
    lblCodRatioCapaPago.Caption = nCodAlerta
    cboTipoCliente.ListIndex = IndiceListaCombo(cboTipoCliente, nTipoCliente)
    txtValorCapaPago.Text = nValorLimite
    cboUnidadCapaPago.ListIndex = IndiceListaCombo(cboUnidadCapaPago, nUnidadMedida)
    cboEstadoCapaPago.ListIndex = IndiceListaCombo(cboEstadoCapaPago, nEstado)
    
    Call HabilitaControles(, True, , , , True, True, True, , True, True)
    Call VisualizaControles(, , True, True, True, , , True)
End Sub
Private Sub EditarRatioOtro(ByVal sNombre As String, _
                            ByVal sFormula As String, _
                            ByVal nValorLimite As Double, _
                            ByVal nAlerta As Integer, _
                            ByVal nUnidadMedida As Integer, _
                            ByVal nEstado As Integer)
     txtNombreRatioOtro.Text = sNombre
     txtFormulaRatioOtro.Text = sFormula
     txtValorRatioOtro.Text = nValorLimite
     cboUnidadRatioOtro.ListIndex = IndiceListaCombo(cboUnidadRatioOtro, nUnidadMedida)
     cboEstadoRatioOtro.ListIndex = IndiceListaCombo(cboEstadoRatioOtro, nEstado)
     lblCodRatioRatioOtro.Caption = nAlerta
   
     Call HabilitaControles(, True, , , , True, True, True, , True, True, , True, True)
     Call VisualizaControles(True, True, , , , True, True, True)
End Sub
Private Sub AgregaConfigCapaPago(ByVal psProducto As String, _
                                ByVal psSubProducto As String, _
                                ByVal psTipoCredito As String, _
                                ByVal psSubTipoCredito As String, _
                                ByVal psTipoCliente As String, _
                                ByVal pnValor As Double, _
                                ByVal psUnidadCapaPago As String, _
                                ByVal pnProducto As Integer, _
                                ByVal pnSubProducto As Double, _
                                ByVal pnTipoCredito As Double, _
                                ByVal pnSubTipoCredito As Double, _
                                ByVal pnTipoCliente As Double, _
                                ByVal pnUnidaCapaPago As Integer, _
                                ByVal pnEstado As Integer, _
                                ByVal pnCodigoAlerta As Integer)

    Dim nFila As Long
    If lbEditar = True Then
        Me.feListaCapaPago.EliminaFila (feListaCapaPago.row)
        Me.feListaCapaPago.AdicionaFila
        Me.feListaCapaPago.Col = 1
        nFila = feListaCapaPago.rows - 1
    Else
        Me.feListaCapaPago.AdicionaFila
        Me.feListaCapaPago.Col = 1
        nFila = feListaCapaPago.rows - 1
    End If
    
    feListaCapaPago.TextMatrix(nFila, 1) = psProducto
    feListaCapaPago.TextMatrix(nFila, 2) = psSubProducto
    feListaCapaPago.TextMatrix(nFila, 3) = psTipoCredito
    feListaCapaPago.TextMatrix(nFila, 4) = psSubTipoCredito
    feListaCapaPago.TextMatrix(nFila, 5) = psTipoCliente
    feListaCapaPago.TextMatrix(nFila, 6) = Format(pnValor, "#,##0.00")
    feListaCapaPago.TextMatrix(nFila, 7) = psUnidadCapaPago
    feListaCapaPago.TextMatrix(nFila, 8) = pnProducto
    feListaCapaPago.TextMatrix(nFila, 9) = pnSubProducto
    feListaCapaPago.TextMatrix(nFila, 10) = pnTipoCredito
    feListaCapaPago.TextMatrix(nFila, 11) = pnSubTipoCredito
    feListaCapaPago.TextMatrix(nFila, 12) = pnTipoCliente
    feListaCapaPago.TextMatrix(nFila, 13) = pnUnidaCapaPago
    feListaCapaPago.TextMatrix(nFila, 14) = pnEstado
    feListaCapaPago.TextMatrix(nFila, 15) = pnCodigoAlerta
    
    If pnEstado = 2 Then
        feListaCapaPago.BackColorRow (&HC0C0FF)
    Else
        feListaCapaPago.BackColorRow (&HC0FFC0)
    End If
    feListaCapaPago.SetFocus
End Sub
Private Sub AgregaConfigRatioOtro(ByVal psNombreRatioOtro As String, _
                                ByVal psFormulaRatioOtro As String, _
                                ByVal pnValor As Double, _
                                ByVal psUnidadRatioOtro As String, _
                                ByVal pnCodigoAlerta As Integer, _
                                ByVal pnUnidaRatiosOtros As Integer, _
                                ByVal pnEstado As Integer)
    Dim nFila As Long
    If lbEditar = True Then
        Me.feListaRatiosOtros.EliminaFila (feListaRatiosOtros.row)
        Me.feListaRatiosOtros.AdicionaFila
        Me.feListaRatiosOtros.Col = 1
        nFila = feListaRatiosOtros.rows - 1
    Else
        Me.feListaRatiosOtros.AdicionaFila
        Me.feListaRatiosOtros.Col = 1
        nFila = feListaRatiosOtros.rows - 1
    End If
    
    feListaRatiosOtros.TextMatrix(nFila, 1) = psNombreRatioOtro
    feListaRatiosOtros.TextMatrix(nFila, 2) = psFormulaRatioOtro
    feListaRatiosOtros.TextMatrix(nFila, 3) = pnValor
    feListaRatiosOtros.TextMatrix(nFila, 4) = psUnidadRatioOtro
    feListaRatiosOtros.TextMatrix(nFila, 5) = pnCodigoAlerta
    feListaRatiosOtros.TextMatrix(nFila, 6) = pnUnidaRatiosOtros
    feListaRatiosOtros.TextMatrix(nFila, 7) = pnEstado
    
    If pnEstado = 2 Then
        feListaRatiosOtros.BackColorRow (&HC0C0FF)
    Else
        feListaRatiosOtros.BackColorRow (&HC0FFC0)
    End If
    feListaRatiosOtros.SetFocus
End Sub
Private Sub LimpiarDatos()
    SSTabDatos.Tab = 0
    cboProducto.SetFocus
    Call LimpiaFlex(Me.feListaCapaPago)
    Call LimpiaFlex(Me.feListaRatiosOtros)
    Me.txtValorCapaPago.Text = "0.00"
    Me.txtValorRatioOtro.Text = "0.00"
    Me.txtNombreRatioOtro.Text = ""
    Me.txtFormulaRatioOtro.Text = ""
End Sub

Private Function ValidaDatos() As String
    ValidaDatos = False
    If ValidaDatosFlexCapaPago = False Then 'Contenido de feReferidos2: Referidos
        SSTabDatos.Tab = 0
        ValidaDatos = False
        Exit Function
    End If

    If ValidaDatosFlexRatioOtro = False Then 'Contenido de feReferidos2: Referidos
        SSTabDatos.Tab = 1
        ValidaDatos = False
        Exit Function
    End If
    
    If txtNombreCapaPago.Text = "" Or txtFormulaCapaPago.Text = "" Then
        SSTabDatos.Tab = 0
        ValidaDatos = False
        Exit Function
    End If
    ValidaDatos = True
End Function

Public Function ValidaDatosFlexCapaPago() As Boolean
    Dim i As Integer, j As Integer
    ValidaDatosFlexCapaPago = False
        'Valida grilla de capacidad de pago
        For i = 1 To feListaCapaPago.rows - 1 'Verifica que los datos no sean iguales
            For j = 1 To feListaCapaPago.rows - 1
                If i <> j Then
                    If feListaCapaPago.TextMatrix(i, 8) = feListaCapaPago.TextMatrix(j, 8) And _
                        feListaCapaPago.TextMatrix(i, 10) = feListaCapaPago.TextMatrix(j, 10) And _
                        feListaCapaPago.TextMatrix(i, 12) = feListaCapaPago.TextMatrix(j, 12) And _
                        feListaCapaPago.TextMatrix(i, 15) Then
                        MsgBox "Ya existe un Producto / TipoCredito / TipoCliente en la fila " & str(i) & ", por favor verifique....", vbInformation, "Alerta"
                        ValidaDatosFlexCapaPago = False
                        Exit Function
                    End If
                End If
            Next
        Next
    ValidaDatosFlexCapaPago = True
End Function

Public Function ValidaDatosFlexRatioOtro() As Boolean
    Dim i As Integer, j As Integer
    ValidaDatosFlexRatioOtro = False
        For i = 1 To feListaRatiosOtros.rows - 1 'Verifica que los datos no sean iguales
            For j = 1 To feListaRatiosOtros.rows - 1
                If i <> j Then
                    If feListaRatiosOtros.TextMatrix(i, 1) = feListaRatiosOtros.TextMatrix(j, 1) Then
                        MsgBox "Existe duplicidad en los registros de datos en fila " & str(i) & " y " & str(j) & ", por favor verifique....", vbInformation, "Alerta"
                        
                        ValidaDatosFlexRatioOtro = False
                        Exit Function
                    End If
                End If
            Next
        Next
    ValidaDatosFlexRatioOtro = True
End Function
Private Function ValidaDatosCapaPago() As String
    Dim nIndice As Integer
    Dim nEstado As Integer
    Dim nTipoCliente As String
    ValidaDatosCapaPago = ""
    
    nTipoCliente = IIf(Right(Me.cboTipoCliente.Text, 1), Right(Me.cboTipoCliente.Text, 1), "0")
    nEstado = IIf(Right(Me.cboEstadoCapaPago.Text, 1), Right(Me.cboEstadoCapaPago.Text, 1), "0")
    
    For nIndice = 1 To feListaCapaPago.rows - 1
        If IIf(feListaCapaPago.TextMatrix(nIndice, 8) <> "", feListaCapaPago.TextMatrix(nIndice, 8), "0") = Right(Me.cboProducto.Text, 3) _
        And IIf(feListaCapaPago.TextMatrix(nIndice, 10) <> "", feListaCapaPago.TextMatrix(nIndice, 10), "0") = Right(Me.cboTipoCredito.Text, 3) _
        And IIf(feListaCapaPago.TextMatrix(nIndice, 12) <> "", feListaCapaPago.TextMatrix(nIndice, 12), "0") = nTipoCliente _
        And IIf(feListaCapaPago.TextMatrix(nIndice, 15) <> "", feListaCapaPago.TextMatrix(nIndice, 15), "0") = nEstado Then
            ValidaDatosCapaPago = "Ya existe un Producto / TipoCredito / TipoCliente en la fila " & str(nIndice) & ", por favor verifique...."
            Exit Function
        End If
    Next
End Function
Private Function ValidaDatosRatio() As String
    Dim nIndice As Integer
    Dim sNombreRatio As String
    ValidaDatosRatio = ""
    
    sNombreRatio = Trim(Me.txtNombreRatioOtro.Text)
    
    For nIndice = 1 To feListaRatiosOtros.rows - 1
        If (feListaRatiosOtros.TextMatrix(nIndice, 1)) = sNombreRatio Then
            ValidaDatosRatio = "Ya existe un ratio con ese nombre en la fila" & str(nIndice) & ", por favor verifique...."
            Exit Function
        End If
    Next
End Function
Private Function CargarDatos() As Boolean
On Error GoTo ErrorCargarDatos
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
    
    Dim rsCapaPago As ADODB.Recordset
    Set rsCapaPago = New ADODB.Recordset
    Dim rsCapaPagoDet As ADODB.Recordset
    Set rsCapaPagoDet = New ADODB.Recordset
    Dim rsRatiosOtros As ADODB.Recordset
    Set rsRatiosOtros = New ADODB.Recordset
    Dim rsCodRatios As ADODB.Recordset
    Set rsCodRatios = New ADODB.Recordset
    
    Dim i As Integer
    Dim sNombreRatioCapaPago As String
    Dim sFormulaRatioCapaPago As String
    
    CargarDatos = oNCOMFormatosEval.CargaDatosAlertaTempranaConfig(Trim(Right(cboProducto.Text, 3)), _
                                                                    Trim(Right(cboSubProducto.Text, 3)), _
                                                                    Trim(Right(cboTipoCredito.Text, 3)), _
                                                                    Trim(Right(cboSubTipo.Text, 3)), _
                                                                    rsCapaPago, _
                                                                    rsCapaPagoDet, _
                                                                    rsRatiosOtros, _
                                                                    rsCodRatios)
    
    
    lnCantFilaCapaPago = 0
    lnCantFilaRatioOtro = 0
    lnCantFilaCapaPago = CInt(rsCodRatios!nCodAlertaCapaPago)
    lnCantFilaRatioOtro = CInt(rsCodRatios!nCodAlertaRatioOtro)
    
    'Capacidad de Pago
     If Not (rsCapaPago.BOF And rsCapaPago.EOF) Then
        sNombreRatioCapaPago = rsCapaPago!Ratio
        sFormulaRatioCapaPago = rsCapaPago!Formula
        Me.txtNombreCapaPago.Text = sNombreRatioCapaPago
        Me.txtFormulaCapaPago.Text = sFormulaRatioCapaPago
    End If
        Me.txtNombreRatioOtro.Text = ""
        Me.txtFormulaRatioOtro.Text = ""
    
    Call LimpiaFlex(Me.feListaCapaPago)
    If CargarDatos Then
        If Not (rsCapaPagoDet.BOF And rsCapaPagoDet.EOF) Then
               For i = 0 To rsCapaPagoDet.RecordCount - 1
                   feListaCapaPago.AdicionaFila
                   feListaCapaPago.TextMatrix(i + 1, 1) = rsCapaPagoDet!Producto
                   feListaCapaPago.TextMatrix(i + 1, 2) = rsCapaPagoDet!subproducto
                   feListaCapaPago.TextMatrix(i + 1, 3) = rsCapaPagoDet!Credito
                   feListaCapaPago.TextMatrix(i + 1, 4) = rsCapaPagoDet!TipoCredito
                   feListaCapaPago.TextMatrix(i + 1, 5) = rsCapaPagoDet!TipoCliente
                   feListaCapaPago.TextMatrix(i + 1, 6) = Format(rsCapaPagoDet!nValorLimite, "#,##0.00")
                   feListaCapaPago.TextMatrix(i + 1, 7) = rsCapaPagoDet!unidad
                   feListaCapaPago.TextMatrix(i + 1, 8) = rsCapaPagoDet!cProdCabCod
                   feListaCapaPago.TextMatrix(i + 1, 9) = rsCapaPagoDet!cTpoProdCod
                   feListaCapaPago.TextMatrix(i + 1, 10) = rsCapaPagoDet!cCredCabCod
                   feListaCapaPago.TextMatrix(i + 1, 11) = rsCapaPagoDet!cTpoCredCod
                   feListaCapaPago.TextMatrix(i + 1, 12) = rsCapaPagoDet!nColocCondicion
                   feListaCapaPago.TextMatrix(i + 1, 13) = rsCapaPagoDet!nUnidadMedida
                   feListaCapaPago.TextMatrix(i + 1, 14) = rsCapaPagoDet!nEstado
                   feListaCapaPago.TextMatrix(i + 1, 15) = rsCapaPagoDet!nCodAlerta
                
                    If rsCapaPagoDet!nEstado = 2 Then
                        feListaCapaPago.BackColorRow (&HC0C0FF)
                    Else
                        feListaCapaPago.BackColorRow (&HC0FFC0)
                    End If
                   rsCapaPagoDet.MoveNext
               Next i
        End If

    'Ratios Otros
    Call LimpiaFlex(Me.feListaRatiosOtros)
        If Not (rsRatiosOtros.BOF And rsRatiosOtros.EOF) Then
        'lnCountFila = rsRatiosOtros.RecordCount - 1
               For i = 0 To rsRatiosOtros.RecordCount - 1
                   feListaRatiosOtros.AdicionaFila
                   feListaRatiosOtros.TextMatrix(i + 1, 0) = i + 1
                   feListaRatiosOtros.TextMatrix(i + 1, 1) = rsRatiosOtros!cRatio
                   feListaRatiosOtros.TextMatrix(i + 1, 2) = rsRatiosOtros!cFormula
                   feListaRatiosOtros.TextMatrix(i + 1, 3) = rsRatiosOtros!nValorLimite
                   feListaRatiosOtros.TextMatrix(i + 1, 4) = rsRatiosOtros!cUnidad
                   feListaRatiosOtros.TextMatrix(i + 1, 5) = rsRatiosOtros!nCodAlerta
                   feListaRatiosOtros.TextMatrix(i + 1, 6) = rsRatiosOtros!nUnidad
                   feListaRatiosOtros.TextMatrix(i + 1, 7) = rsRatiosOtros!nEstado
                   
                    If rsRatiosOtros!nEstado = 2 Then
                        feListaRatiosOtros.BackColorRow (&HC0C0FF)
                    Else
                        feListaRatiosOtros.BackColorRow (&HC0FFC0)
                    End If
                   
                   rsRatiosOtros.MoveNext
               Next i
        End If
    End If
    
    RSClose rsCapaPago
    RSClose rsCapaPagoDet
    RSClose rsRatiosOtros
    Exit Function
ErrorCargarDatos:
    CargarDatos = False
    MsgBox Err.Description + ": Error al carga datos", vbCritical, "Error"
End Function
Private Sub HabilitaControles(Optional pbProductoCredito As Boolean = False, _
                              Optional pbConfigLimite As Boolean = False, _
                              Optional pbCmdMostrar As Boolean = False, _
                              Optional pbCmdNuevoCapaPago As Boolean = False, _
                              Optional pbCmdEliminarCapaPago As Boolean = False, _
                              Optional pbCmdAceptarCapaPago As Boolean = False, _
                              Optional pbCmdCancelarCapaPago As Boolean = False, _
                              Optional pbCmdGrabar As Boolean = False, _
                              Optional pbCmdEditar As Boolean = False, _
                              Optional pbCmdCancelar As Boolean = False, _
                              Optional pbCmdSalir As Boolean = False, _
                              Optional pbCmdNuevoRatioOtro As Boolean = False, _
                              Optional pbCmdAceptarRatioOtro As Boolean = False, _
                              Optional pbCmdCancelarRatioOtro As Boolean = False, _
                              Optional pbfelistaRatios As Boolean = False)
    'Producto / SubProducto
    Me.cboProducto.Enabled = pbProductoCredito
    Me.cboSubProducto.Enabled = pbProductoCredito
    Me.cboTipoCredito.Enabled = pbProductoCredito
    Me.cboSubTipo.Enabled = pbProductoCredito
    frameProducto.Enabled = pbProductoCredito
    frameCredito.Enabled = pbProductoCredito
    lblProducto.Enabled = pbProductoCredito
    lblSubProducto.Enabled = pbProductoCredito
    lblCredito.Enabled = pbProductoCredito
    lblTipoCredito.Enabled = pbProductoCredito
    
    'Configuracion Limites
    Me.frameConfigCapaPago.Enabled = pbConfigLimite
    Me.frameConfigRatioOtro.Enabled = pbConfigLimite
    Me.txtNombreCapaPago.Enabled = pbConfigLimite
    Me.txtFormulaCapaPago.Enabled = pbConfigLimite
    Me.cboTipoCliente.Enabled = pbConfigLimite
    Me.txtValorCapaPago.Enabled = pbConfigLimite
    Me.txtValorRatioOtro.Enabled = pbConfigLimite
    Me.cboEstadoCapaPago.Enabled = pbConfigLimite
    Me.cboEstadoRatioOtro.Enabled = pbConfigLimite
    Me.cboUnidadCapaPago.Enabled = pbConfigLimite
    Me.cboUnidadRatioOtro.Enabled = pbConfigLimite
    
    lblCapaPago.Enabled = pbConfigLimite
    lblCodRatioCapaPago.Enabled = pbConfigLimite
    lblRatioOtro.Enabled = pbConfigLimite
    lblCodRatioRatioOtro.Enabled = pbConfigLimite
    lblFormulaCapaPago.Enabled = pbConfigLimite
    lblTipoCreditoCapaPago.Enabled = pbConfigLimite
    lblValorCapaPago.Enabled = pbConfigLimite
    lblUnidadCapaPago.Enabled = pbConfigLimite
    lblEstadoCapaPago.Enabled = pbConfigLimite
    
    'Configuracion de botones Varios
    Me.cmdMostrar.Enabled = pbCmdMostrar
    Me.cmdNuevoCapaPaga.Enabled = pbCmdNuevoCapaPago
    Me.cmdEliminarCapaPago.Enabled = pbCmdEliminarCapaPago
    Me.cmdAceptarCapaPago.Enabled = pbCmdAceptarCapaPago
    Me.cmdCancelarCapaPago.Enabled = pbCmdCancelarCapaPago
    
    Me.cmdNuevoRatioOtro.Enabled = pbCmdNuevoRatioOtro
    Me.cmdAceptarRatioOtro.Enabled = pbCmdAceptarRatioOtro
    Me.cmdCancelarRatioOtro.Enabled = pbCmdCancelarRatioOtro
    
    Me.cmdGrabar.Enabled = pbCmdGrabar
    Me.cmdEditar.Enabled = pbCmdEditar
    Me.cmdCancelar.Enabled = pbCmdCancelar
    Me.cmdSalir.Enabled = pbCmdSalir
    
    'Formulas de ratios varios
    txtNombreRatioOtro.Enabled = pbConfigLimite
    txtFormulaRatioOtro.Enabled = pbConfigLimite
    txtValorRatioOtro.Enabled = pbConfigLimite
    cboUnidadRatioOtro.Enabled = pbConfigLimite
    cboEstadoRatioOtro.Enabled = pbConfigLimite
    
    'Listado de valores de ratios
    feListaCapaPago.Enabled = pbfelistaRatios
    feListaRatiosOtros.Enabled = pbfelistaRatios
    frameListadoCapaPago.Enabled = pbfelistaRatios
    frameListadoRatioOtro.Enabled = pbfelistaRatios
End Sub

Private Sub VisualizaControles(Optional pbCmdNuevoCapaPago As Boolean = False, _
                               Optional pbCmdEliminar As Boolean = False, _
                               Optional pbCmdNuevoRatioOtro As Boolean = False, _
                               Optional pbCmdAceptarCapaPago As Boolean = False, _
                               Optional pbCmdCancelarCapaPago As Boolean = False, _
                               Optional pbCmdAceptarRatioOtro As Boolean = False, _
                               Optional pbCmdCancelarRatioOtro As Boolean = False, _
                               Optional pbCodAlerta As Boolean = False)
                               
    Me.cmdNuevoCapaPaga.Visible = pbCmdNuevoCapaPago
    Me.cmdEliminarCapaPago.Visible = pbCmdEliminar
    Me.cmdAceptarCapaPago.Visible = pbCmdAceptarCapaPago
    Me.cmdCancelarCapaPago.Visible = pbCmdCancelarCapaPago
    
    Me.cmdNuevoRatioOtro.Visible = pbCmdNuevoRatioOtro
    Me.cmdAceptarRatioOtro.Visible = pbCmdAceptarRatioOtro
    Me.cmdCancelarRatioOtro.Visible = pbCmdCancelarRatioOtro
    
    Me.lblCodRatioCapaPago.Visible = pbCodAlerta
    Me.lblCodRatioRatioOtro.Visible = pbCodAlerta
    
    Me.lblCapaPago.Visible = pbCodAlerta
    Me.lblRatioOtro.Visible = pbCodAlerta
End Sub
Private Sub LlenarCombo(ByRef pCombo As ComboBox, ByRef pRs As ADODB.Recordset)
    If (pRs.BOF Or pRs.EOF) Then
        Exit Sub
    End If
    pCombo.Clear
    Do While Not pRs.EOF
        pCombo.AddItem pRs!cConsDescripcion & Space(300) & pRs!nConsValor
        pRs.MoveNext
    Loop
End Sub
Private Sub cboTipoCredito_Click()
    Call CargaSubTipoCredito(Trim(Right(cboTipoCredito.Text, 3)))
    Call CambiaTamañoCombo(cboTipoCredito, 250)
End Sub
Private Sub CargaSubTipoCredito(ByVal psTipo As String)
    Dim oCred As COMDCredito.DCOMCredito
    Dim ssql As String
    Dim RTemp As ADODB.Recordset
    On Error GoTo ERRORCargaSubTipoCredito
    Set oCred = New COMDCredito.DCOMCredito
    Set RTemp = oCred.RecuperaSubTipoCrediticios(psTipo)
    Set oCred = Nothing
    cboSubTipo.Clear
    Do While Not RTemp.EOF
        cboSubTipo.AddItem RTemp!cConsDescripcion & Space(250) & RTemp!nConsValor
        RTemp.MoveNext
    Loop
    RTemp.Close
    Set RTemp = Nothing
    Call CambiaTamañoCombo(cboSubTipo, 250)
    Exit Sub
ERRORCargaSubTipoCredito:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub
Private Sub cboProducto_Click()
    Call CargaSubProducto(Trim(Right(cboProducto.Text, 3)))
     Call CambiaTamañoCombo(cboProducto, 250)
End Sub
Private Sub CargaSubProducto(ByVal psTipo As String)
Dim oCred As COMDCredito.DCOMFormatosEval
Set oCred = New COMDCredito.DCOMFormatosEval
Dim ssql As String
Dim RTemp As ADODB.Recordset
Set RTemp = New ADODB.Recordset
    On Error GoTo ERRORCargaSubProducto
    Set oCred = New COMDCredito.DCOMFormatosEval
    Set RTemp = oCred.ObtenerSubProducto(3033, psTipo)
    Set oCred = Nothing
    cboSubProducto.Clear
    Do While Not RTemp.EOF
        cboSubProducto.AddItem RTemp!cConsDescripcion & Space(250) & RTemp!nConsValor
        RTemp.MoveNext
    Loop
    RTemp.Close
    Set RTemp = Nothing
    Call CambiaTamañoCombo(cboSubProducto, 250)
    Exit Sub
ERRORCargaSubProducto:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub

'Eventos de producto, subproducto, credito y tipo de crédito
Private Sub cboProducto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboSubProducto.SetFocus
    End If
End Sub
Private Sub cboSubProducto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboTipoCredito.SetFocus
    End If
End Sub
Private Sub cboTipoCredito_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboSubTipo.SetFocus
    End If
End Sub
Private Sub cboSubTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdMostrar.SetFocus
    End If
End Sub
' Eventos para Ratio Capacida de Pago
Private Sub txtFormulaCapaPago_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras4(KeyAscii, True)
    If KeyAscii = 13 Then
        cboTipoCliente.SetFocus
    End If
End Sub
Private Sub txtNombreCapaPago_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras4(KeyAscii, True)
End Sub
Private Sub cboTipoCliente_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    txtValorCapaPago.SetFocus
 End If
End Sub
Private Sub txtValorCapaPago_KeyPress(KeyAscii As Integer)
 KeyAscii = NumerosDecimales(txtValorCapaPago, KeyAscii, 10, , True)
 If KeyAscii = 45 Then KeyAscii = 0
 If KeyAscii = 13 Then
    If Len(txtValorCapaPago.Text) >= 8 Then
        MsgBox "El valor límite ingresado, supera la cantidad de caracteres", vbInformation, "Aviso"
        txtValorCapaPago.Text = "0.00"
        txtValorCapaPago.SetFocus
    Else
        cboUnidadCapaPago.SetFocus
    End If
 End If

End Sub
Private Sub txtValorCapaPago_LostFocus()
    If Trim(txtValorCapaPago.Text) = "" Or Len(txtValorCapaPago.Text) >= 8 Then
        txtValorCapaPago.Text = "0.00"
    Else
        txtValorCapaPago.Text = Format(txtValorCapaPago.Text, "###," & String(15, "#") & "#0.00")
    End If
End Sub
Private Sub cboUnidadCapaPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboEstadoCapaPago.SetFocus
    End If
End Sub
Private Sub cboEstadoCapaPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cmdAceptarCapaPago.SetFocus
    End If
End Sub
Private Sub cmdAceptarCapaPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdNuevoCapaPaga.SetFocus
    End If
End Sub

' Evento para Ratios Otros
Private Sub txtNombreRatioOtro_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras4(KeyAscii, True)
    If KeyAscii = 13 Then
       txtFormulaRatioOtro.SetFocus
    End If
End Sub
Private Sub txtFormulaRatioOtro_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras4(KeyAscii, True)
    If KeyAscii = 13 Then
        txtValorRatioOtro.SetFocus
    End If
End Sub
Private Sub txtValorRatioOtro_KeyPress(KeyAscii As Integer)
 KeyAscii = NumerosDecimales(txtValorRatioOtro, KeyAscii, 10, , True)
 If KeyAscii = 45 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(txtValorRatioOtro.Text) >= 8 Then
            MsgBox "El valor límite ingresado, supera la cantidad de caracteres", vbInformation, "Aviso"
            txtValorRatioOtro.Text = "0.00"
            txtValorRatioOtro.SetFocus
        Else
            cboUnidadRatioOtro.SetFocus
        End If
    End If
End Sub
Private Sub txtValorRatioOtro_LostFocus()
    If Trim(txtValorRatioOtro.Text) = "" Then
        txtValorRatioOtro.Text = "0.00"
    Else
        txtValorRatioOtro.Text = Format(txtValorRatioOtro.Text, "###," & String(15, "#") & "#0.00")
    End If
End Sub
Private Sub cboUnidadRatioOtro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboEstadoRatioOtro.SetFocus
    End If
End Sub
Private Sub cboEstadoRatioOtro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptarRatioOtro.SetFocus
    End If
End Sub
