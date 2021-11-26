VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCredFormEvalConfigTpoProd 
   Caption         =   "Configuración de Tipo de Producto Crediticios - Formatos de Evaluación"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14370
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   14370
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   12480
      TabIndex        =   19
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   11040
      TabIndex        =   18
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Height          =   375
      Left            =   1680
      TabIndex        =   17
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      Caption         =   "Producto y Sub Productos"
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   11055
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   375
         Left            =   9480
         TabIndex        =   15
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox cmbSubProducto 
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   480
         Width           =   3705
      End
      Begin VB.ComboBox cmbProducto 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   480
         Width           =   3225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sub Producto:"
         Height          =   195
         Left            =   4560
         TabIndex        =   14
         Top             =   480
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Producto:"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   690
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   14175
      Begin TabDlg.SSTab SSTab1 
         Height          =   5175
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   13935
         _ExtentX        =   24580
         _ExtentY        =   9128
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "Formatos"
         TabPicture(0)   =   "frmCredFormEvalConfigTpoProd.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame6"
         Tab(0).Control(1)=   "Frame5"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Ratios e Indicadores"
         TabPicture(1)   =   "frmCredFormEvalConfigTpoProd.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Frame3"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "cmdAceptar"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "cmdQuitar"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "cmdEditarRatios"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "cmdNuevo"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "cmdCancelaRatios"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "txtLimite"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "chkLimite"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).ControlCount=   9
         Begin VB.CheckBox chkLimite 
            Caption         =   "%Limite:"
            Height          =   255
            Left            =   240
            TabIndex        =   56
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox txtLimite 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   1560
            TabIndex        =   55
            Top             =   2160
            Width           =   810
         End
         Begin VB.Frame Frame6 
            Caption         =   "Eligir Formato"
            Height          =   4575
            Left            =   -74880
            TabIndex        =   40
            Top             =   480
            Width           =   4335
            Begin MSComctlLib.ListView lvFormatos 
               Height          =   4065
               Left            =   120
               TabIndex        =   41
               Top             =   240
               Width           =   4035
               _ExtentX        =   7117
               _ExtentY        =   7170
               View            =   3
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               Checkboxes      =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   4
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Selec"
                  Object.Width           =   1411
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Formato"
                  Object.Width           =   6174
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "minimo"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "maximo"
                  Object.Width           =   0
               EndProperty
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Formatos Seleccionados"
            Height          =   4575
            Left            =   -70440
            TabIndex        =   38
            Top             =   480
            Width           =   9135
            Begin SICMACT.FlexEdit feFormatosSeleccionados 
               Height          =   4215
               Left            =   120
               TabIndex        =   39
               Top             =   240
               Width           =   8880
               _ExtentX        =   15663
               _ExtentY        =   7435
               Cols0           =   6
               HighLight       =   1
               EncabezadosNombres=   "-Formatos-Mínimo-Máximo-Aplica Rangos-nCodForm"
               EncabezadosAnchos=   "300-3000-1500-1500-1500-0"
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
               ColumnasAEditar =   "X-X-X-X-4-X"
               ListaControles  =   "0-0-0-0-4-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-L-R-R-R-C"
               FormatosEdit    =   "0-0-2-2-3-0"
               lbEditarFlex    =   -1  'True
               Enabled         =   0   'False
               lbUltimaInstancia=   -1  'True
               TipoBusqueda    =   3
               lbBuscaDuplicadoText=   -1  'True
               ColWidth0       =   300
               RowHeight0      =   300
            End
         End
         Begin VB.CommandButton cmdCancelaRatios 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   12480
            TabIndex        =   8
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CommandButton cmdNuevo 
            Caption         =   "Nuevo"
            Height          =   375
            Left            =   12480
            TabIndex        =   7
            Top             =   3960
            Width           =   1335
         End
         Begin VB.CommandButton cmdEditarRatios 
            Caption         =   "Editar"
            Height          =   375
            Left            =   12480
            TabIndex        =   6
            Top             =   2400
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CommandButton cmdQuitar 
            Caption         =   "Quitar"
            Height          =   375
            Left            =   12480
            TabIndex        =   5
            Top             =   4440
            Width           =   1335
         End
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   12480
            TabIndex        =   4
            Top             =   840
            Width           =   1335
         End
         Begin VB.Frame Frame3 
            Caption         =   "Configuración de Ratios"
            Height          =   2415
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   12255
            Begin VB.TextBox txtCriticoDel 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Left            =   6120
               TabIndex        =   49
               Top             =   1080
               Width           =   810
            End
            Begin VB.TextBox txtCriticoAl 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Left            =   8160
               TabIndex        =   48
               Top             =   1080
               Width           =   810
            End
            Begin VB.TextBox txtAceptableDel 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Left            =   6120
               TabIndex        =   47
               Top             =   600
               Width           =   810
            End
            Begin VB.TextBox txtAceptableAl 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Left            =   8160
               TabIndex        =   46
               Top             =   600
               Width           =   810
            End
            Begin VB.TextBox txtMontoAl 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Left            =   3480
               TabIndex        =   45
               Top             =   1320
               Width           =   1170
            End
            Begin VB.TextBox txtMontoDel 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Left            =   1440
               TabIndex        =   44
               Top             =   1320
               Width           =   1170
            End
            Begin VB.CheckBox chkIndispen 
               Caption         =   "Indispensable"
               Height          =   255
               Left            =   6120
               TabIndex        =   37
               Top             =   1560
               Width           =   1455
            End
            Begin VB.ComboBox cmbRatio 
               Height          =   315
               Left            =   6120
               Style           =   2  'Dropdown List
               TabIndex        =   29
               Top             =   240
               Width           =   3225
            End
            Begin VB.ComboBox cmbTipoCliente 
               Height          =   315
               Left            =   1440
               Style           =   2  'Dropdown List
               TabIndex        =   27
               Top             =   960
               Width           =   3225
            End
            Begin VB.CheckBox chkMontos 
               Caption         =   "Montos:"
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Top             =   1320
               Width           =   975
            End
            Begin VB.CheckBox chkTpoCli 
               Caption         =   "Tipo Cliente:"
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   960
               Width           =   1215
            End
            Begin VB.CheckBox chkMeses 
               Caption         =   "Meses:"
               Height          =   255
               Left            =   120
               TabIndex        =   21
               Top             =   600
               Width           =   975
            End
            Begin VB.CheckBox chkIfis 
               Caption         =   "Ifis:"
               Height          =   255
               Left            =   120
               TabIndex        =   20
               Top             =   240
               Width           =   975
            End
            Begin Spinner.uSpinner SpnIfis 
               Height          =   375
               Left            =   1440
               TabIndex        =   24
               Top             =   240
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   661
               Max             =   300
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
            End
            Begin Spinner.uSpinner spnMesesDel 
               Height          =   375
               Left            =   1440
               TabIndex        =   25
               Top             =   600
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   661
               Max             =   12
               Min             =   1
               MaxLength       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
            End
            Begin Spinner.uSpinner spnMesesAl 
               Height          =   375
               Left            =   2280
               TabIndex        =   26
               Top             =   600
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   661
               Max             =   12
               Min             =   1
               MaxLength       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
            End
            Begin Spinner.uSpinner spnAceptableDel 
               Height          =   375
               Left            =   10920
               TabIndex        =   33
               Top             =   600
               Visible         =   0   'False
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
               Max             =   70
               MaxLength       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
            End
            Begin Spinner.uSpinner spnAceptableAl 
               Height          =   375
               Left            =   11640
               TabIndex        =   34
               Top             =   600
               Visible         =   0   'False
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
               Max             =   70
               MaxLength       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
            End
            Begin Spinner.uSpinner spnCriticoDel 
               Height          =   375
               Left            =   10920
               TabIndex        =   35
               Top             =   1080
               Visible         =   0   'False
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
               Max             =   70
               MaxLength       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
            End
            Begin Spinner.uSpinner spnCriticoAl 
               Height          =   375
               Left            =   11640
               TabIndex        =   36
               Top             =   1080
               Visible         =   0   'False
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
            End
            Begin VB.Label Label14 
               Caption         =   "%"
               Height          =   255
               Left            =   2280
               TabIndex        =   54
               Top             =   1680
               Width           =   255
            End
            Begin VB.Label Label12 
               Caption         =   "%"
               Height          =   255
               Left            =   9000
               TabIndex        =   53
               Top             =   1120
               Width           =   255
            End
            Begin VB.Label Label11 
               Caption         =   "%"
               Height          =   255
               Left            =   6960
               TabIndex        =   52
               Top             =   1120
               Width           =   255
            End
            Begin VB.Label Label10 
               Caption         =   "%"
               Height          =   255
               Left            =   9000
               TabIndex        =   51
               Top             =   640
               Width           =   255
            End
            Begin VB.Label Label9 
               Caption         =   "%"
               Height          =   255
               Left            =   6960
               TabIndex        =   50
               Top             =   640
               Width           =   255
            End
            Begin VB.Label Label8 
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   7560
               TabIndex        =   43
               Top             =   1080
               Width           =   135
            End
            Begin VB.Label Label7 
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   7560
               TabIndex        =   42
               Top             =   600
               Width           =   135
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Crítico:"
               Height          =   195
               Left            =   5160
               TabIndex        =   32
               Top             =   1200
               Width           =   510
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Aceptable:"
               Height          =   195
               Left            =   5160
               TabIndex        =   31
               Top             =   720
               Width           =   765
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Ratio:"
               Height          =   195
               Left            =   5160
               TabIndex        =   30
               Top             =   360
               Width           =   420
            End
            Begin VB.Label Label3 
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3000
               TabIndex        =   28
               Top             =   1320
               Width           =   135
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Listados"
            Height          =   2055
            Left            =   120
            TabIndex        =   2
            Top             =   3000
            Width           =   12255
            Begin SICMACT.FlexEdit feListaRatios 
               Height          =   1695
               Left            =   120
               TabIndex        =   9
               Top             =   240
               Width           =   12000
               _ExtentX        =   21167
               _ExtentY        =   3836
               Cols0           =   21
               HighLight       =   1
               EncabezadosNombres=   $"frmCredFormEvalConfigTpoProd.frx":0038
               EncabezadosAnchos=   "300-1500-1500-1500-1500-2600-1500-1500-0-0-0-0-0-0-0-0-0-0-0-1500-0"
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
               ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
               ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-C-C-C-C-L-C-C-C-C-C-C-C-C-C-C-C-C-C-R-C"
               FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
               lbEditarFlex    =   -1  'True
               lbUltimaInstancia=   -1  'True
               TipoBusqueda    =   3
               lbBuscaDuplicadoText=   -1  'True
               ColWidth0       =   300
               RowHeight0      =   300
            End
         End
      End
   End
End
Attribute VB_Name = "frmCredFormEvalConfigTpoProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub chkIfis_Click()
    If Me.chkIfis.value = 1 Then
        Me.SpnIfis.Enabled = True
    Else
        Me.SpnIfis.Enabled = False
    End If
End Sub

Private Sub chkLimite_Click()
    If Me.chkLimite.value = 1 Then
        Me.txtLimite.Enabled = True
    Else
       Me.txtLimite.Enabled = False
    End If
End Sub

Private Sub chkMeses_Click()
    If Me.chkMeses.value = 1 Then
        Me.spnMesesDel.Enabled = True
        Me.spnMesesAl.Enabled = True
    Else
        Me.spnMesesDel.Enabled = False
        Me.spnMesesAl.Enabled = False
    End If
End Sub

Private Sub chkMontos_Click()
    If Me.chkMontos.value = 1 Then
        Me.txtMontoDel.Enabled = True
        Me.txtMontoAl.Enabled = True
    Else
        Me.txtMontoDel.Enabled = False
        Me.txtMontoAl.Enabled = False
    End If
End Sub

Private Sub chkTpoCli_Click()
    If Me.chkTpoCli.value = 1 Then
        Me.cmbTipoCliente.Enabled = True
    Else
        Me.cmbTipoCliente.Enabled = False
    End If
End Sub

Private Sub cmbProducto_Click()
    Call CargaSubProducto(Trim(Right(cmbProducto.Text, 3)))
End Sub

Private Sub cmbRatio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAceptableDel.SetFocus
    End If
End Sub

Private Sub cmbTipoCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.chkMontos.value = 1 Then
            Me.txtMontoDel.SetFocus
        End If
    End If
End Sub

Private Sub CmdAceptar_Click()

    Dim i As Integer
    Dim sMsj As String
'    For i = 1 To Me.feListaRatios.Rows - 1
'        If Me.feListaRatios.TextMatrix(i, 16) = Right(Me.cmbRatio.Text, 1) Then
'            MsgBox "Este Ratio ya fue ingresado..."
'            Exit Sub
'        End If
'    Next i
    sMsj = ValidaDatos
    If sMsj <> "" Then
        MsgBox sMsj, vbInformation, "Alerta"
        Exit Sub
    End If
   
    If Me.chkIfis.value = 1 Then lnIfis = Me.SpnIfis.valor Else lnIfis = 0
    If Me.chkMeses.value = 1 Then lnMesesDel = Me.spnMesesDel.valor Else lnMesesDel = 0
    If Me.chkMeses.value = 1 Then lnMesesAl = Me.spnMesesAl.valor Else lnMesesAl = 0
    If Me.chkTpoCli.value = 1 Then lcTipoCli = Trim(Left(Me.cmbTipoCliente.Text, Len(Me.cmbTipoCliente.Text) - 1)) Else lcTipoCli = ""
    If Me.chkTpoCli.value = 1 Then lnTipoCli = Right(Me.cmbTipoCliente.Text, 1) Else lnTipoCli = 0
    If Me.chkMontos.value = 1 Then lnMontoDel = Me.txtMontoDel.Text Else lnMontoDel = 0
    If Me.chkMontos.value = 1 Then lnMontoAL = Me.txtMontoAl.Text Else lnMontoAL = 0
    If Me.chkIndispen.value = 1 Then lnIndispen = 1 Else lnIndispen = 0
        
'    Call AgregaConfigRatio(lnIfis, lnMesesDel, lnMesesAl, lcTipoCli, lnTipoCli, lnMontoDel, lnMontoAL, _
'        "", Right(Me.cmbRatio.Text, 1), CDbl(Me.spnAceptableDel.valor), val(Me.spnAceptableAl.valor), val(Me.spnCriticoDel.valor), val(Me.spnCriticoAl.valor), lnIndispen)
            

'    Dim oNCred As COMDCredito.DCOMFormatosEval
'    Call oNCred.AgregaCredFormEvalRatios(Right(Me.cmbProducto.Text, 3), Right(Me.cmbSubProducto.Text, 3), IIf(Me.feListaRatios.TextMatrix(i, 1) = "No Aplica", 0, Me.feListaRatios.TextMatrix(i, 1)), _
'    IIf(Me.feListaRatios.TextMatrix(i, 8) = True, 1, 0), CInt(Me.feListaRatios.TextMatrix(i, 9)), CInt(Me.feListaRatios.TextMatrix(i, 10)), CDbl(Me.feListaRatios.TextMatrix(i, 11)), _
'    CDbl(Me.feListaRatios.TextMatrix(i, 12)), CDbl(Me.feListaRatios.TextMatrix(i, 13)), CDbl(Me.feListaRatios.TextMatrix(i, 14)), CDbl(Me.feListaRatios.TextMatrix(i, 17)), _
'    CDbl(Me.feListaRatios.TextMatrix(i, 18)), CInt(Me.feListaRatios.TextMatrix(i, 15)), CInt(Me.feListaRatios.TextMatrix(i, 16)))
        
        
        Dim bLimite As Boolean
        bLimite = IIf(chkLimite.value = 1, True, False)
        Call AgregaConfigRatio(lnIfis, lnMesesDel, lnMesesAl, lcTipoCli, lnTipoCli, lnMontoDel, lnMontoAL, _
        Trim(Left(Me.cmbRatio.Text, Len(Me.cmbRatio.Text) - 1)), Right(Me.cmbRatio.Text, 1), CDbl(Me.txtAceptableDel.Text), val(Me.txtAceptableAl.Text), val(Me.txtCriticoDel.Text), val(Me.txtCriticoAl.Text), lnIndispen, val(Me.txtLimite.Text), bLimite)
        
        'trim(left(Me.cmbRatio.Text,len(Me.cmbRatio.Text)-1))
        
    Me.SSTab1.TabEnabled(0) = True
    Call ActiDesaControles(, , , True, True, True, , True, , , , , , , , , , , , , , , , , , , , , True, True, True, False)
    
'-- limpia pestaña ratios
Me.chkIfis.value = 0
Me.chkMeses.value = 0
Me.chkTpoCli.value = 0
Me.chkMontos.value = 0
Me.SpnIfis.valor = 0
Me.spnMesesDel.valor = 0
Me.spnMesesAl.valor = 0
'Me.cmbTipoCliente.Text = ""
Me.txtMontoDel.Text = "0.00"
Me.txtMontoAl.Text = "0.00"
Me.chkLimite.value = 0 'CTI3 ERS032020
Me.txtLimite.Text = "0" 'CTI3 ERS032020
'Me.cmbRatio.Text = ""

'Me.spnAceptableDel.valor = 0
Me.txtAceptableDel.Text = "0.00"

'Me.spnAceptableAl.valor = 0
Me.txtAceptableAl.Text = "0.00"

'Me.spnCriticoDel.valor = 0
Me.txtCriticoDel.Text = "0.00"

'Me.spnCriticoAl.valor = 0
Me.txtCriticoAl.Text = "0.00"

Me.chkIndispen.value = 0
    
End Sub

Private Sub cmdCancelar_Click()
    
    Call CargaControles

End Sub

Private Sub cmdCancelaRatios_Click()
    Me.SSTab1.TabEnabled(0) = True
    Call ActiDesaControles(, , , True, True, True, , True, , , , , , , , , , , , , , , , , , , , , True, True, True, False)
    Me.feFormatosSeleccionados.lbEditarFlex = True
    
'-- limpia pestaña ratios
Me.chkIfis.value = 0
Me.chkMeses.value = 0
Me.chkTpoCli.value = 0
Me.chkMontos.value = 0
Me.SpnIfis.valor = 0
Me.spnMesesDel.valor = 0
Me.spnMesesAl.valor = 0
'Me.cmbTipoCliente.Text = ""
Me.txtMontoDel.Text = "0.00"
Me.txtMontoAl.Text = "0.00"
'Me.cmbRatio.Text = ""

'Me.spnAceptableDel.valor = 0
Me.txtAceptableDel.Text = 0

'Me.spnAceptableAl.valor = 0
Me.txtAceptableAl.Text = "0.00"

'Me.spnCriticoDel.valor = 0
Me.txtCriticoDel.Text = "0.00"

'Me.spnCriticoAl.valor = 0
Me.txtCriticoAl.Text = "0.00"

Me.chkIndispen.value = 0
   
Me.chkLimite.value = 0 'CTI03 ERS0032020
Me.txtLimite.Text = "0" 'CTI03 ERS0032020
End Sub

Private Sub CmdEditar_Click()
    'CTI3 ERS0032020
    Dim bLimite As Boolean
         bLimite = False
    If Right(Me.cmbProducto.Text, 3) = "700" Then
        bLimite = True
    End If
    'END
    Call ActiDesaControles(False, False, False, True, True, True, False, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False, True, True, True, False)
    
End Sub


Private Sub cmdEditarRatios_Click()
    
    If Me.feListaRatios.TextMatrix(1, 1) = "" Then
        MsgBox "No hay datos a editar"
        Exit Sub
    End If
    
    Me.SSTab1.TabEnabled(0) = False
    Call ActiDesaControles(, , , True, , , , , , True, True, True, True, , , , , , , True, True, True, True, True, True, True, True, True, , , False, True)
    
    Call CargaRatiosParaEditar
    
    
End Sub

Private Sub CmdGrabar_Click()
    Dim oNCred As COMDCredito.DCOMFormatosEval
    Dim i As Integer
    Dim nId As String
    Dim sLimite As String
    Set oNCred = New COMDCredito.DCOMFormatosEval
        
    If MsgBox("Los Datos ingresdos se guardarán, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
            
        If (Right(Me.cmbProducto.Text, 3)) = "700" Then
            If MsgBox("Usted está configurando los limites del credito consumo, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        End If
        '--- guarda formatos seleccionados
        Call oNCred.LimpiaFormatosProdEval(Right(Me.cmbProducto.Text, 3), Right(Me.cmbSubProducto.Text, 3))
        If Me.feFormatosSeleccionados.TextMatrix(1, 1) <> "" Then
            If Me.feFormatosSeleccionados.TextMatrix(1, 1) <> "" Then
                For i = 1 To feFormatosSeleccionados.rows - 1
                    Call oNCred.AgregaCredFormEvalConfigProd(feFormatosSeleccionados.TextMatrix(i, 5), Right(Me.cmbProducto.Text, 3), Right(Me.cmbSubProducto.Text, 3), IIf(feFormatosSeleccionados.TextMatrix(i, 4) = ".", 1, 0))
                Next i
            End If
        End If
        
        '--- guarda ratios e indicadores
        Call oNCred.LimpiaProdEvalRatios(Right(Me.cmbProducto.Text, 3), Right(Me.cmbSubProducto.Text, 3))
        If Me.feListaRatios.TextMatrix(1, 1) <> "" Then
            If Me.feListaRatios.TextMatrix(1, 1) <> "" Then
                For i = 1 To feListaRatios.rows - 1
                    sLimite = Me.feListaRatios.TextMatrix(i, 19)
                    If (Me.feListaRatios.TextMatrix(i, 19)) = "No Aplica" Then
                        sLimite = "0"
                    End If
                    Call oNCred.AgregaCredFormEvalRatios(Right(Me.cmbProducto.Text, 3), Right(Me.cmbSubProducto.Text, 3), IIf(Me.feListaRatios.TextMatrix(i, 1) = "No Aplica", 0, Me.feListaRatios.TextMatrix(i, 1)), _
                    IIf(Me.feListaRatios.TextMatrix(i, 8) = True, 1, 0), CInt(Me.feListaRatios.TextMatrix(i, 9)), CInt(Me.feListaRatios.TextMatrix(i, 10)), CDbl(Me.feListaRatios.TextMatrix(i, 11)), _
                    CDbl(Me.feListaRatios.TextMatrix(i, 12)), CDbl(Me.feListaRatios.TextMatrix(i, 13)), CDbl(Me.feListaRatios.TextMatrix(i, 14)), CDbl(Me.feListaRatios.TextMatrix(i, 17)), _
                    CDbl(Me.feListaRatios.TextMatrix(i, 18)), CInt(Me.feListaRatios.TextMatrix(i, 15)), CInt(Me.feListaRatios.TextMatrix(i, 16)), IIf(Me.feListaRatios.TextMatrix(i, 20) = True, 1, 0), CInt(IIf(Trim(sLimite) = "", 0, Replace(sLimite, "%", ""))))
                Next i
            End If
        End If

        MsgBox "Se realizaron los cambios satisfactoriamente.", vbInformation, "Atención"

    Call ActiDesaControles(True, True, False, False, False, False, False, False, True, , , , , , , , , , , , , , , , , , , , , , , False)

Call CargaControles

End Sub

Private Sub ActiDesaControles(Optional pbProd As Boolean = False, Optional pbSubProd As Boolean = False, Optional pbMostrar As Boolean = False, _
Optional pbEligeForm As Boolean = False, Optional pbFormSelecc As Boolean = False, Optional pbGrabar As Boolean = False, Optional pbEditar As Boolean = False, _
Optional pbCancel As Boolean = False, Optional pbSalir As Boolean = False, Optional pbchkIfis As Boolean = False, Optional pbmeses As Boolean = False, _
Optional pbTpoCli As Boolean = False, Optional pbMontos As Boolean = False, Optional pbSpnIfis As Boolean = False, Optional pbspnMesesDel As Boolean = False, _
Optional pbspnMesesAl As Boolean = False, Optional pbTipoCliente As Boolean = False, Optional pbMontoDel As Boolean = False, Optional pbMontoAl As Boolean = False, _
Optional pbRatio As Boolean = False, Optional pbspnAceptableDel As Boolean = False, Optional pbspnAceptableAl As Boolean = False, Optional pbspnCriticoDel As Boolean = False, _
Optional pbspnCriticoAl As Boolean = False, Optional pbchkIndispen As Boolean = False, Optional pbfelistaRatios As Boolean = True, Optional pbAceptar As Boolean = False, _
Optional pbCancelaratios As Boolean = False, Optional pbNuevo As Boolean = False, Optional pbEditarRatios As Boolean = False, Optional pbQuitar As Boolean = False, Optional pbLimite As Boolean = False)

    '-- pestaña Formatos
    Me.cmbProducto.Enabled = pbProd
    Me.cmbSubProducto.Enabled = pbSubProd
    Me.cmdMostrar.Enabled = pbMostrar
    Me.lvFormatos.Enabled = pbEligeForm
    Me.feFormatosSeleccionados.Enabled = pbFormSelecc
    Me.cmdGrabar.Enabled = pbGrabar
    Me.cmdEditar.Enabled = pbEditar
    Me.cmdCancelar.Enabled = pbCancel
    Me.cmdSalir.Enabled = pbSalir

    '-- pestaña ratios e indicadores
    Me.chkIfis.Enabled = pbchkIfis
    Me.chkMeses.Enabled = pbmeses
    Me.chkTpoCli.Enabled = pbTpoCli
    Me.chkMontos.Enabled = pbMontos
    Me.SpnIfis.Enabled = pbSpnIfis
    Me.spnMesesDel.Enabled = pbspnMesesDel
    Me.spnMesesAl.Enabled = pbspnMesesAl
    Me.cmbTipoCliente.Enabled = pbTipoCliente
    Me.txtMontoDel.Enabled = pbMontoDel
    Me.txtMontoAl.Enabled = pbMontoAl
    Me.cmbRatio.Enabled = pbRatio
    Me.chkLimite.Enabled = pbLimite 'CTI3 ERS003-2020
    Me.txtLimite.Enabled = False 'CTI3 ERS003-2020
'    Me.spnAceptableDel.Enabled = pbspnAceptableDel
'    Me.spnAceptableAl.Enabled = pbspnAceptableAl
'    Me.spnCriticoDel.Enabled = pbspnCriticoDel
'    Me.spnCriticoAl.Enabled = pbspnCriticoAl
    
    Me.txtAceptableDel.Enabled = pbspnAceptableDel
    Me.txtAceptableAl.Enabled = pbspnAceptableAl
    Me.txtCriticoDel.Enabled = pbspnCriticoDel
    Me.txtCriticoAl.Enabled = pbspnCriticoAl
    
    
    Me.chkIndispen.Enabled = pbchkIndispen
    Me.feListaRatios.Enabled = pbfelistaRatios
    
    Me.cmdAceptar.Enabled = pbAceptar
    Me.cmdCancelaRatios.Enabled = pbCancelaratios
    Me.cmdNuevo.Enabled = pbNuevo
    Me.cmdEditarRatios.Enabled = pbEditarRatios
    Me.cmdQuitar.Enabled = pbQuitar
    
'    Me.cmbProducto.Enabled = True
'    Me.cmbSubProducto.Enabled = True
'    Me.cmdMostrar.Enabled = True
'    Me.lvFormatos.Enabled = True
'    Me.feFormatosSeleccionados.Enabled = True
'    Me.cmdGrabar.Enabled = True
'    Me.cmdEditar.Enabled = True
'    Me.cmdCancelar.Enabled = True
'    Me.cmdSalir.Enabled = True
    
End Sub

Private Sub cmdMostrar_Click()
Dim nCantidad As Integer
Dim objFormEval As COMDCredito.DCOMFormatosEval
Set objFormEval = New COMDCredito.DCOMFormatosEval
Dim objRS As ADODB.Recordset
Set objRS = New ADODB.Recordset
Dim lvItem As ListItem

If Trim(Me.cmbProducto.Text) = "" Then
    MsgBox "Debe seleccionar un producto crediticio", vbInformation, "Aviso"
    Exit Sub
End If

If Trim(Me.cmbSubProducto.Text) = "" Then
    MsgBox "Debe seleccionar un Sub producto", vbInformation, "Aviso"
    Exit Sub
End If

Call ActiDesaControles(False, False, False, False, False, False, True, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False)

'-- llena los formatos para elegir
Set objRS = objFormEval.ObtenerFormatosChkParaElegir(Trim(Right(cmbProducto.Text, 3)), Trim(Right(cmbSubProducto.Text, 3)))
    lvFormatos.ListItems.Clear
    If Not (objRS.BOF Or objRS.EOF) Then
    
        Do While Not objRS.EOF
           Set lvItem = lvFormatos.ListItems.Add(, , objRS!nCodForm)
           lvItem.SubItems(1) = objRS!cNomFormato
           lvItem.SubItems(2) = objRS!nMontoMin
           lvItem.SubItems(3) = objRS!nMontoMax
           If objRS!nCodFormSelec <> "999" Then
                lvItem.Checked = True
                'lvItem.SubItems(4) = "."
           Else
                lvItem.Checked = False
                'lvItem.SubItems(4) = ""
           End If
           objRS.MoveNext
        Loop
        End If
    RSClose objRS

'-- llena formatos selecionados
Set objRS = objFormEval.ObtenerFormatosSeleccionados(Trim(Right(cmbProducto.Text, 3)), Trim(Right(cmbSubProducto.Text, 3)))
    Call LimpiaFlex(feFormatosSeleccionados)
     If Not (objRS.BOF And objRS.EOF) Then
            Me.feFormatosSeleccionados.lbEditarFlex = True
            For i = 0 To objRS.RecordCount - 1
                feFormatosSeleccionados.AdicionaFila
                feFormatosSeleccionados.TextMatrix(i + 1, 0) = i + 1
                feFormatosSeleccionados.TextMatrix(i + 1, 1) = objRS!formatos
                feFormatosSeleccionados.TextMatrix(i + 1, 2) = Format(objRS!nMontoMin, "#,##0.00")
                feFormatosSeleccionados.TextMatrix(i + 1, 3) = Format(objRS!nMontoMax, "#,##0.00")
                feFormatosSeleccionados.TextMatrix(i + 1, 4) = IIf(objRS!nAplicaRango = True, "1", "")
                feFormatosSeleccionados.TextMatrix(i + 1, 5) = objRS!nCodForm
                objRS.MoveNext
            Next i
    End If
    RSClose objRS

'-- llena los RATIOS
Set objRS = objFormEval.ObtenerFormatosEvalRatios(Trim(Right(cmbProducto.Text, 3)), Trim(Right(cmbSubProducto.Text, 3)))
    Call LimpiaFlex(Me.feListaRatios)
     If Not (objRS.BOF And objRS.EOF) Then
            For i = 0 To objRS.RecordCount - 1
                feListaRatios.AdicionaFila
                feListaRatios.TextMatrix(i + 1, 0) = i + 1
                feListaRatios.TextMatrix(i + 1, 1) = objRS!cIfis
                feListaRatios.TextMatrix(i + 1, 2) = objRS!cMeses
                feListaRatios.TextMatrix(i + 1, 3) = objRS!cTipoCliente
                feListaRatios.TextMatrix(i + 1, 4) = objRS!cMontos
                feListaRatios.TextMatrix(i + 1, 5) = objRS!cRatio
                feListaRatios.TextMatrix(i + 1, 6) = objRS!cAceptable
                feListaRatios.TextMatrix(i + 1, 7) = objRS!cCritico
                
                feListaRatios.TextMatrix(i + 1, 8) = objRS!nIndispensable
                feListaRatios.TextMatrix(i + 1, 9) = objRS!nMesIni
                feListaRatios.TextMatrix(i + 1, 10) = objRS!nMesFin
                feListaRatios.TextMatrix(i + 1, 11) = objRS!nMontoIni
                feListaRatios.TextMatrix(i + 1, 12) = objRS!nMontoFin
                feListaRatios.TextMatrix(i + 1, 13) = objRS!nAceptableIni
                feListaRatios.TextMatrix(i + 1, 14) = objRS!nAceptableFin
                feListaRatios.TextMatrix(i + 1, 15) = objRS!nTipoCliente
                feListaRatios.TextMatrix(i + 1, 16) = objRS!nCodRatio
                feListaRatios.TextMatrix(i + 1, 17) = objRS!nCriticoIni
                feListaRatios.TextMatrix(i + 1, 18) = objRS!nCriticoFin
                feListaRatios.TextMatrix(i + 1, 19) = objRS!cLimite 'CTI3 ERS0032020
                feListaRatios.TextMatrix(i + 1, 20) = objRS!nLimite 'CTI3 ERS0032020
                objRS.MoveNext
            Next i
    End If
    RSClose objRS
    
''-- limpia pestaña ratios
'Me.chkIfis.value = 0
'Me.chkMeses.value = 0
'Me.chkTpoCli.value = 0
'Me.chkMontos.value = 0
'Me.SpnIfis.valor = 0
'Me.spnMesesDel.valor = 0
'Me.spnMesesAl.valor = 0
'Me.cmbTipoCliente.Text = ""
'Me.txtMontoDel.Text = "0.00"
'Me.txtMontoAl.Text = "0.00"
'Me.cmbRatio.Text = ""
'Me.spnAceptableDel.valor = 0
'Me.spnAceptableAl.valor = 0
'Me.spnCriticoDel.valor = 0
'Me.spnCriticoAl.valor = 0
'Me.chkIndispen.value = 0

End Sub


Private Sub CargaControles()

Dim lvItem As ListItem

Dim objProd As COMDCredito.DCOMFormatosEval
Set objProd = New COMDCredito.DCOMFormatosEval

Dim ObjCons As COMDConstantes.DCOMConstantes
Set ObjCons = New COMDConstantes.DCOMConstantes

Dim objRS As ADODB.Recordset
Set objRS = New ADODB.Recordset

Me.txtMontoDel.Text = "0.00"
Me.txtMontoAl.Text = "0.00"

'spnAceptableDel.valor = "0.00"
'spnAceptableAl.valor = "0.00"
'spnCriticoDel.valor = "0.00"
'spnCriticoAl.valor = "0.00"

txtAceptableDel.Text = "0.00"
txtAceptableAl.Text = "0.00"
txtCriticoDel.Text = "0.00"
txtCriticoAl.Text = "0.00"
txtLimite.Text = "0" 'CTI3 ERS0032020
''-- limpia pestaña ratios
'Me.chkIfis.value = 0
'Me.chkMeses.value = 0
'Me.chkTpoCli.value = 0
'Me.chkMontos.value = 0
'Me.SpnIfis.valor = 0
'Me.spnMesesDel.valor = 0
'Me.spnMesesAl.valor = 0
'Me.cmbTipoCliente.Text = ""
'Me.txtMontoDel.Text = "0.00"
'Me.txtMontoAl.Text = "0.00"
'Me.cmbRatio.Text = ""
'Me.spnAceptableDel.valor = 0
'Me.spnAceptableAl.valor = 0
'Me.spnCriticoDel.valor = 0
'Me.spnCriticoAl.valor = 0
'Me.chkIndispen.value = 0
Me.chkLimite.value = 0
Call ActiDesaControles(True, True, True, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False)

lvFormatos.ListItems.Clear
Call LimpiaFlex(feFormatosSeleccionados)
Call LimpiaFlex(Me.feListaRatios)

Set objRS = ObjCons.RecuperaConstantes(3015)
Call LlenarCombo(cmbTipoCliente, objRS)
Set objRS = Nothing
Set objRS = New ADODB.Recordset
If cmbTipoCliente.ListCount > 0 Then
 cmbTipoCliente.ListIndex = 0
End If

Set objRS = ObjCons.RecuperaConstantes(7021)
Call LlenarCombo(cmbRatio, objRS)
Set objRS = Nothing
Set objRS = New ADODB.Recordset
If cmbRatio.ListCount > 0 Then
 cmbRatio.ListIndex = 0
End If

'Set objRS = ObjCons.RecuperaConstantes(3033)
Set objRS = objProd.ObtenerProducto(3033)
Call LlenarCombo(cmbProducto, objRS)
Set objRS = Nothing
Set objRS = New ADODB.Recordset
If cmbProducto.ListCount > 0 Then
 cmbProducto.ListIndex = 0
End If

'-- llena formatos
lvFormatos.ListItems.Clear
Set objRS = objProd.RecuperaConfigFornatosEval
'Set objRS = objProd.ObtenerProducto(3033)
If Not (objRS.BOF Or objRS.EOF) Then
    Do While Not objRS.EOF
       Set lvItem = lvFormatos.ListItems.Add(, , objRS!nCodForm)
       'Set lvItem = lvFormatos.ListItems.Add(, , objRS!nConsValor)
       lvItem.SubItems(1) = objRS!cNomFormato
       lvItem.SubItems(2) = objRS!nMontoMin
       lvItem.SubItems(3) = objRS!nMontoMax

'       If objRS!nEstado Then
'            lvItem.Checked = True
'       Else
'            lvItem.Checked = False
'       End If
       objRS.MoveNext
    Loop
    End If
RSClose objRS


End Sub

'-- llena formatos
''oCred.RecuperaConfigFornatosEval

Private Sub CargaSubProducto(ByVal psTipo As String)
Dim oCred As COMDCredito.DCOMFormatosEval
Dim ssql As String
Dim RTemp As ADODB.Recordset
    On Error GoTo ERRORCargaSubProducto
    Set oCred = New COMDCredito.DCOMFormatosEval
    Set RTemp = oCred.ObtenerSubProducto(3033, psTipo)
    Set oCred = Nothing
    cmbSubProducto.Clear
    Do While Not RTemp.EOF
        cmbSubProducto.AddItem RTemp!cConsDescripcion & Space(250) & RTemp!nConsValor
        RTemp.MoveNext
    Loop
    RTemp.Close
    Set RTemp = Nothing
    ''Call CambiaTamañoCombo(cmbSubProducto, 250)
    Exit Sub
    
ERRORCargaSubProducto:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub



'Set objRS = objLinea.ObtenerProductoCreditocioAgencia("", 0)
'If objRS.EOF Then
'   RSClose objRS
'   MsgBox "No se definieron Agencias en el Sistema...Consultar con Sistemas", vbInformation, "Aviso"
'   Exit Sub
'End If
'Do While Not objRS.EOF
'   Set lvItem = lvAgencia.ListItems.Add(, , objRS!cCodigo)
'   lvItem.SubItems(1) = objRS!cDescri
'   lvItem.Checked = False
'   objRS.MoveNext
'Loop
'RSClose objRS
'
'Set objRS = objLinea.ObtenerConstanteLineaCredito(3016)
'If objRS.EOF Then
'   RSClose objRS
'   MsgBox "No se definieron los destinos de creditos en la agencia...Consultar con Sistemas", vbInformation, "Aviso"
'   Exit Sub
'End If
'Do While Not objRS.EOF
'   Set lvItem = lvDestino.ListItems.Add(, , objRS!cCodigo)
'   lvItem.SubItems(1) = objRS!cDescri
'   lvItem.Checked = False
'   objRS.MoveNext
'Loop


Private Sub LlenarCombo(ByRef pCombo As ComboBox, ByRef pRs As ADODB.Recordset)
'    pRs.MoveFirst
    If (pRs.BOF Or pRs.EOF) Then
    Exit Sub
    End If
    pCombo.Clear
    Do While Not pRs.EOF
        pCombo.AddItem pRs!cConsDescripcion & Space(300) & pRs!nConsValor
        pRs.MoveNext
    Loop
End Sub


Private Sub cmdNuevo_Click()
    'CTI3 ERS0032020
    Dim bLimite As Boolean
         bLimite = False
    If Right(Me.cmbProducto.Text, 3) = "700" Then
        bLimite = True
    End If
    'END
    Me.SSTab1.TabEnabled(0) = False
    Call ActiDesaControles(, , , , , , , , , True, True, True, True, , , , , , , True, True, True, True, True, True, True, True, True, , , False, bLimite)
    
End Sub

Private Sub cmdQuitar_Click()
    
    If Me.feListaRatios.TextMatrix(1, 1) = "" Then
        MsgBox "No hay datos a Quitar"
        Exit Sub
    End If
    
    If MsgBox("¿ Está seguro de eliminar este registro ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Me.feListaRatios.EliminaFila (Me.feListaRatios.row)
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub EditMoney1_Change()

End Sub


Private Sub Form_Load()
    Call CargaControles
End Sub


Private Sub lvFormatos_Click()
Dim i, j, K, nbusca As Integer
Dim MatrixHojaEval() As String
Dim nXPos, nPos As Integer

nbusca = 0


'--- inserta checkeados
For i = 1 To lvFormatos.ListItems.count
    If lvFormatos.ListItems(i).Checked = True Then
        nbusca = 0
        For j = 1 To Me.feFormatosSeleccionados.rows - 1
            If feFormatosSeleccionados.TextMatrix(j, 5) = lvFormatos.ListItems(i).Text Then
                nbusca = nbusca + 1
            End If
        Next
        If nbusca = 0 Then
            Call AgregaFormato(lvFormatos.ListItems(i).Text, lvFormatos.ListItems(i).SubItems(1), lvFormatos.ListItems(i).SubItems(2), lvFormatos.ListItems(i).SubItems(3))
        End If
    End If
Next
''---------- fin inserta checkeados


''----- elimina no checkeados

For i = 1 To lvFormatos.ListItems.count
    If lvFormatos.ListItems(i).Checked = False Then
        For j = 1 To Me.feFormatosSeleccionados.rows - 1
            ntotfilas = Me.feFormatosSeleccionados.rows - 1
            If feFormatosSeleccionados.TextMatrix(j, 5) = lvFormatos.ListItems(i).Text Then
            feFormatosSeleccionados.EliminaFila (j)
            j = ntotfilas
''---------- fin elimina fila

            End If
        Next
    End If
Next

End Sub


Private Sub AgregaFormato(ByVal pnCodForm As Integer, ByVal pcNomForm As String, ByVal pnMontoMin As String, ByVal pnMontoMax As String)
    Dim nFila As Long
    feFormatosSeleccionados.AdicionaFila
    feFormatosSeleccionados.col = 1
    nFila = feFormatosSeleccionados.rows - 1
    feFormatosSeleccionados.TextMatrix(nFila, 1) = pcNomForm
    feFormatosSeleccionados.TextMatrix(nFila, 2) = pnMontoMin
    feFormatosSeleccionados.TextMatrix(nFila, 3) = pnMontoMax
    feFormatosSeleccionados.TextMatrix(nFila, 4) = ""
    feFormatosSeleccionados.TextMatrix(nFila, 5) = pnCodForm
    feFormatosSeleccionados.SetFocus

End Sub

Private Sub AgregaConfigRatio(ByVal nIfis As Integer, ByVal nMesesDel As Integer, ByVal nMesesAl As Integer, ByVal cTipoCli As String, ByVal nTipoCli As Integer, ByVal nMontoDel As Double, ByVal nMontoAl As Double, _
        ByVal cRatio As String, ByVal nRatio As Integer, ByVal nAceptableDel As Double, ByVal nAceptableAl As Double, ByVal nCriticoDel As Double, ByVal nCriticoAl As Double, ByVal nIndispen As Integer, ByVal nLimite As Integer, ByVal bLimite As Boolean)

    Dim nFila As Long
    Me.feListaRatios.AdicionaFila
    feListaRatios.col = 1
    nFila = feListaRatios.rows - 1
    feListaRatios.TextMatrix(nFila, 1) = IIf(nIfis = 0, "No Aplica", nIfis) '
    feListaRatios.TextMatrix(nFila, 2) = IIf(nMesesAl = 0, "No Aplica", str(nMesesDel) + " - " + str(nMesesAl))
    feListaRatios.TextMatrix(nFila, 3) = IIf(cTipoCli = "", "No Aplica", cTipoCli)
    feListaRatios.TextMatrix(nFila, 4) = IIf(nMontoAl = 0, "No Aplica", str(nMontoDel) + " - " + str(nMontoAl))
    feListaRatios.TextMatrix(nFila, 5) = cRatio
    feListaRatios.TextMatrix(nFila, 6) = str(nAceptableDel) + " - " + str(nAceptableAl)
    feListaRatios.TextMatrix(nFila, 7) = str(nCriticoDel) + " - " + str(nCriticoAl)
    
    feListaRatios.TextMatrix(nFila, 8) = nIndispen '
    feListaRatios.TextMatrix(nFila, 9) = nMesesDel '
    feListaRatios.TextMatrix(nFila, 10) = nMesesAl '
    feListaRatios.TextMatrix(nFila, 11) = nMontoDel '
    feListaRatios.TextMatrix(nFila, 12) = nMontoAl '
    feListaRatios.TextMatrix(nFila, 13) = nAceptableDel '
    feListaRatios.TextMatrix(nFila, 14) = nAceptableAl '
    feListaRatios.TextMatrix(nFila, 15) = nTipoCli '
    feListaRatios.TextMatrix(nFila, 16) = nRatio '
    feListaRatios.TextMatrix(nFila, 17) = nCriticoDel '
    feListaRatios.TextMatrix(nFila, 18) = nCriticoAl '
    feListaRatios.TextMatrix(nFila, 19) = IIf(bLimite = True, nLimite & "%", "No Aplica") 'CTI3 ERS0032020
    feListaRatios.TextMatrix(nFila, 20) = bLimite 'CTI3 ERS0032020
    feListaRatios.SetFocus

End Sub

Private Sub CargaRatiosParaEditar()

    Me.chkIfis.value = IIf(feListaRatios.TextMatrix(feListaRatios.row(), 1) = "No Aplica", 0, 1)
    Me.chkMeses.value = IIf(Me.spnMesesAl.valor > 0, 1, 0)
    Me.chkTpoCli.value = IIf(Me.cmbTipoCliente.Text <> "", 1, 0)
    Me.chkMontos.value = IIf(Me.txtMontoDel > 0, 1, 0)
    Me.SpnIfis.valor = feListaRatios.TextMatrix(feListaRatios.row(), 1)
    Me.spnMesesDel.valor = feListaRatios.TextMatrix(feListaRatios.row(), 9)
    Me.spnMesesAl.valor = feListaRatios.TextMatrix(feListaRatios.row(), 10)
    'Me.cmbTipoCliente.Text = feListaRatios.TextMatrix(feListaRatios.Row(), 3) + Space(30) + feListaRatios.TextMatrix(feListaRatios.Row(), 15) '*** obtener el codigo
    Me.cmbTipoCliente.Text = Me.cmbTipoCliente.List((Trim(feListaRatios.TextMatrix(feListaRatios.row(), 15)))) '*** obtener el codigo
    Me.txtMontoDel.Text = feListaRatios.TextMatrix(feListaRatios.row(), 11)
    Me.txtMontoAl.Text = feListaRatios.TextMatrix(feListaRatios.row(), 12)
    'Me.cmbRatio.Text = feListaRatios.TextMatrix(feListaRatios.Row(), 5) & Space(30) + feListaRatios.TextMatrix(feListaRatios.Row(), 16) '*** obtener el codigo
        
    
    
    If Trim(feListaRatios.TextMatrix(feListaRatios.row(), 16)) > 0 Then
        Me.cmbRatio.Enabled = True
        Me.cmbRatio.Text = Me.cmbRatio.List((Trim(feListaRatios.TextMatrix(feListaRatios.row(), 16)))) '*** obtener el codigo
    End If

    
'    Me.spnAceptableDel.valor = feListaRatios.TextMatrix(feListaRatios.row(), 13)
'    Me.spnAceptableAl.valor = feListaRatios.TextMatrix(feListaRatios.row(), 14)
'    Me.spnCriticoDel.valor = feListaRatios.TextMatrix(feListaRatios.row(), 17)
'    Me.spnCriticoAl.valor = feListaRatios.TextMatrix(feListaRatios.row(), 18)

    Me.txtAceptableDel.Text = feListaRatios.TextMatrix(feListaRatios.row(), 13)
    Me.txtAceptableAl.Text = feListaRatios.TextMatrix(feListaRatios.row(), 14)
    Me.txtCriticoDel.Text = feListaRatios.TextMatrix(feListaRatios.row(), 17)
    Me.txtCriticoAl.Text = feListaRatios.TextMatrix(feListaRatios.row(), 18)
    Dim sRatioLimite As String
    sRatioLimite = feListaRatios.TextMatrix(feListaRatios.row(), 19) 'CTI3 ERS0032020
    Me.txtLimite.Text = IIf(Trim(sRatioLimite) = "", 0, sRatioLimite) 'CTI3 ERS0032020
    
    Me.chkIndispen.value = IIf(feListaRatios.TextMatrix(feListaRatios.row(), 8) = "No Aplica", 0, 1)
    
End Sub


Private Sub SpnIfis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.chkMeses.value = 1 Then
            Me.spnMesesDel.SetFocus
        End If
    End If
End Sub

Private Sub spnMesesAl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.chkTpoCli.value = 1 Then
            Me.cmbTipoCliente.SetFocus
        End If
    End If
End Sub

Private Sub spnMesesDel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.spnMesesAl.SetFocus
    End If
End Sub



Private Sub txtAceptableAl_GotFocus()
    fEnfoque txtAceptableAl
End Sub

Private Sub txtAceptableAl_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtAceptableAl, KeyAscii)
    If KeyAscii = 13 Then
        If val(txtAceptableDel.Text) < val(txtAceptableAl.Text) Then
            txtCriticoDel.Text = val(txtAceptableAl.Text) + 0.01
            txtCriticoDel.SetFocus
        Else
            txtAceptableAl.Text = val(txtAceptableDel.Text) + 0.01
        End If
    End If
End Sub

Private Sub txtAceptableAl_LostFocus()
    If Len(Trim(txtAceptableAl.Text)) = 0 Then
         txtAceptableAl.Text = "0.00"
    End If
    txtAceptableAl.Text = Format(txtAceptableAl.Text, "#0.00")
End Sub

Private Sub txtAceptableDel_GotFocus()
    fEnfoque txtAceptableDel
End Sub

Private Sub txtAceptableDel_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtAceptableDel, KeyAscii)
    If KeyAscii = 13 Then
        txtAceptableAl.SetFocus
    End If
End Sub

Private Sub txtAceptableDel_LostFocus()
    If Len(Trim(txtAceptableDel.Text)) = 0 Then
         txtAceptableDel.Text = "0.00"
    End If
    txtAceptableDel.Text = Format(txtAceptableDel.Text, "#0.00")
End Sub

Private Sub txtCriticoAl_GotFocus()
    fEnfoque txtCriticoAl
End Sub

Private Sub txtCriticoAl_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtCriticoAl, KeyAscii)
    If KeyAscii = 13 Then
        If val(txtCriticoDel.Text) < val(txtCriticoAl.Text) Then
            Me.chkIndispen.SetFocus
        Else
            txtCriticoAl.Text = val(txtCriticoDel.Text) + 0.01
        End If
    End If
End Sub

Private Sub txtCriticoAl_LostFocus()
    If Len(Trim(txtCriticoAl.Text)) = 0 Then
         txtCriticoAl.Text = "0.00"
    End If
    txtCriticoAl.Text = Format(txtCriticoAl.Text, "#0.00")
End Sub

Private Sub txtCriticoDel_GotFocus()
    fEnfoque txtCriticoDel
End Sub

Private Sub txtCriticoDel_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtCriticoDel, KeyAscii)
    If KeyAscii = 13 Then
        txtCriticoAl.SetFocus
    End If
End Sub

Private Sub txtCriticoDel_LostFocus()
    If Len(Trim(txtCriticoDel.Text)) = 0 Then
         txtCriticoDel.Text = "0.00"
    End If
    txtCriticoDel.Text = Format(txtCriticoDel.Text, "#0.00")
End Sub
'CTI3 ERS0032020**********************
Private Sub txtLimite_GotFocus()
    fEnfoque txtLimite
End Sub
Private Sub txtLimite_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtLimite, KeyAscii)
    If KeyAscii = 13 Then
        txtAceptableDel.SetFocus
    End If
End Sub

Private Sub txtLimite_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim nNumeric As Boolean
    nNumeric = IsNumeric(txtLimite.Text)
    If (nNumeric = False) Then
        txtLimite.Text = "0.00"
        Exit Sub
    End If
    
    If (Trim(txtLimite.Text) = "") Then
        txtLimite.Text = "0.00"
    End If
End Sub

Private Sub txtLimite_LostFocus()
    If Len(Trim(txtLimite.Text)) = 0 Then
         txtMontoDel.Text = "0"
    End If
    txtLimite.Text = Format(txtLimite.Text, "#0.00")
End Sub
'*************************************

Private Sub txtMontoAl_GotFocus()
    fEnfoque txtMontoAl
End Sub

Private Sub txtMontoAl_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMontoAl, KeyAscii)
    If KeyAscii = 13 Then
        cmbRatio.SetFocus
    End If
End Sub

Private Sub txtMontoAl_LostFocus()
    If Len(Trim(txtMontoAl.Text)) = 0 Then
         txtMontoAl.Text = "0.00"
    End If
    txtMontoAl.Text = Format(txtMontoAl.Text, "#0.00")
End Sub

Private Sub txtMontoDel_GotFocus()
    fEnfoque txtMontoDel
End Sub

Private Sub txtMontoDel_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMontoDel, KeyAscii)
    If KeyAscii = 13 Then
        txtMontoAl.SetFocus
    End If

End Sub

Private Sub txtMontoDel_LostFocus()
    If Len(Trim(txtMontoDel.Text)) = 0 Then
         txtMontoDel.Text = "0.00"
    End If
    txtMontoDel.Text = Format(txtMontoDel.Text, "#0.00")
End Sub

Private Function ValidaDatos() As String
    Dim nIndice As Integer
    Dim nTipoCli As String
    ValidaDatos = ""
    
    nTipoCli = IIf(Me.chkTpoCli.value = 1, Right(Me.cmbTipoCliente.Text, 1), "0")
    
    For nIndice = 1 To feListaRatios.rows - 1
        If feListaRatios.TextMatrix(nIndice, 15) = nTipoCli _
            And feListaRatios.TextMatrix(nIndice, 16) = Right(Me.cmbRatio.Text, 1) Then
'            And feListaRatios.TextMatrix(nIndice, 6) = Str(val(txtAceptableDel.Text)) + " - " + Str(txtAceptableAl.Text) _
'            And feListaRatios.TextMatrix(nIndice, 7) = Str(val(txtCriticoDel.Text)) + " - " + Str(txtCriticoAl.Text) Then
            ValidaDatos = "Ya existe un Ratio y Tipo de Cliente registrado en la fila " & str(nIndice) & ", por favor verifique...."
            Exit Function
        End If
    Next
    
    If CDbl(Me.txtAceptableAl.Text) = 0 Then
        ValidaDatos = "El monto aceptable final debe tener un valor...."
        Exit Function
    End If
    
    If CDbl(Me.txtCriticoAl.Text) = 0 Then
        ValidaDatos = "El monto crítico final debe tener un valor...."
        Exit Function
    End If
    
    If CDbl(Me.txtAceptableDel.Text) > CDbl(Me.txtAceptableAl.Text) Then
        ValidaDatos = "El monto aceptable inicial no pude ser mayor al monto aceptable final...."
        Exit Function
    End If
    
    If CDbl(Me.txtCriticoDel.Text) > CDbl(Me.txtCriticoAl.Text) Then
        ValidaDatos = "El monto crítico inicial no pude ser mayor al monto crítico final...."
        Exit Function
    End If
    
    If Me.chkIfis.value = 1 Then
        If Me.SpnIfis.valor = 0 Then
            ValidaDatos = "El número de IFIs debe ser mayor a cero...."
            Exit Function
        End If
    End If
    
    If Me.chkMontos.value = 1 Then
        If CDbl(Me.txtMontoAl.Text) = 0 Then
            ValidaDatos = "El monto debe tener un valor...."
            Exit Function
        End If
        If CDbl(Me.txtMontoDel.Text) > CDbl(Me.txtMontoAl.Text) Then
            ValidaDatos = "El monto inicial no pude ser mayor al monto final...."
            Exit Function
        End If
    End If
    
    If Me.chkMeses.value = 1 Then
        If Me.spnMesesAl.valor = 0 Then
            ValidaDatos = "El mes final no puede ser cero...."
            Exit Function
        End If
        If Me.spnMesesDel.valor > Me.spnMesesAl.valor Then
            ValidaDatos = "El mes inicial no puede ser mayor al mes final...."
            Exit Function
        End If
    End If
    If (Right(cmbProducto.Text, 3) = "700") Then
        If Trim(Me.txtLimite.Text) = "" Then
            ValidaDatos = "Debe ingresar un porcentaje del limite...."
            Exit Function
        End If
        If Me.chkLimite.value = 1 Then
             If CInt(Me.txtLimite.Text) <= 0 Then
                ValidaDatos = "Debe ingresar un porcentaje del limite mayor a 0%...."
                Exit Function
            End If
            If CInt(Me.txtLimite.Text) > 100 Then
                ValidaDatos = "Debe ingresar un porcentaje del limite menor a 100%...."
                Exit Function
            End If
        End If
    End If
    
    
End Function
