VERSION 5.00
Begin VB.Form frmCapMantenimiento 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   Icon            =   "frmCapMantenimiento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraReglasPorderes 
      Caption         =   "Regla de Poderes"
      Height          =   2055
      Left            =   6550
      TabIndex        =   67
      Top             =   6720
      Width           =   2615
      Begin VB.CommandButton cmdAgregarRegla 
         Caption         =   "Ag&regar"
         Height          =   375
         Left            =   525
         TabIndex        =   70
         Top             =   1575
         Width           =   735
      End
      Begin VB.CommandButton cmdQuitarRega 
         Caption         =   "&Quitar"
         Height          =   375
         Left            =   1470
         TabIndex        =   69
         Top             =   1575
         Width           =   735
      End
      Begin VB.ListBox lsLetras 
         Height          =   1185
         Left            =   105
         Style           =   1  'Checkbox
         TabIndex        =   68
         Top             =   240
         Width           =   795
      End
      Begin SICMACT.FlexEdit grdReglas 
         Height          =   1245
         Left            =   945
         TabIndex        =   71
         Top             =   240
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   2196
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Regla"
         EncabezadosAnchos=   "300-960"
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
         ColumnasAEditar =   "X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0"
         EncabezadosAlineacion=   "C-C"
         FormatosEdit    =   "0-0"
         TextArray0      =   "#"
         ColWidth0       =   300
         RowHeight0      =   300
      End
   End
   Begin VB.CommandButton cmdTasaHist 
      Caption         =   "Tasa Histórica"
      Height          =   495
      Left            =   720
      TabIndex        =   66
      Top             =   5880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   1920
      TabIndex        =   65
      Top             =   9075
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdRenovaciones 
      Caption         =   "Ver &Renovaciones"
      Height          =   390
      Left            =   6600
      TabIndex        =   63
      Top             =   6210
      Visible         =   0   'False
      Width           =   1565
   End
   Begin VB.Frame FRHistRemBruCts 
      Height          =   1290
      Left            =   6400
      TabIndex        =   59
      Top             =   620
      Visible         =   0   'False
      Width           =   2775
      Begin VB.CommandButton cmdHistorico 
         Caption         =   "&Ver Historico"
         Height          =   375
         Left            =   1440
         TabIndex        =   60
         Top             =   600
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblRemuneraciones 
         Caption         =   "Total n Últimas Remu. Brutas"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   360
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lblTotalRemuneraciones 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   120
         TabIndex        =   61
         Top             =   645
         Visible         =   0   'False
         Width           =   1200
      End
   End
   Begin VB.CommandButton cmdRegFirmas 
      Caption         =   "Ver Solicitud Apertura"
      Height          =   540
      Left            =   4200
      TabIndex        =   33
      Top             =   9030
      Width           =   1380
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8520
      TabIndex        =   8
      Top             =   9075
      Width           =   855
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   7560
      TabIndex        =   7
      Top             =   9075
      Width           =   855
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   990
      TabIndex        =   10
      Top             =   9075
      Width           =   855
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "E&ditar"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   9075
      Width           =   855
   End
   Begin VB.Frame fraCliente 
      Caption         =   "Clientes"
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
      Height          =   2805
      Left            =   225
      TabIndex        =   12
      Top             =   1890
      Width           =   8970
      Begin SICMACT.FlexEdit grdCliente 
         Height          =   2085
         Left            =   105
         TabIndex        =   34
         Top             =   210
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   3678
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Codigo-Nombre-Relacion-Flag-Obligatorio-Grupo-esTitular"
         EncabezadosAnchos=   "300-1700-3500-1500-0-0-1500-0"
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
         ColumnasAEditar =   "X-1-X-3-X-5-6-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-1-0-3-0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-L-C-L-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         ColWidth0       =   300
         RowHeight0      =   300
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   7035
         TabIndex        =   5
         Top             =   2340
         Width           =   855
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   7995
         TabIndex        =   6
         Top             =   2340
         Width           =   855
      End
   End
   Begin VB.Frame fraCuenta 
      Caption         =   "Cuenta"
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
      Height          =   8955
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   9225
      Begin VB.CommandButton cmdNuevaBus 
         Height          =   325
         Left            =   4305
         Picture         =   "frmCapMantenimiento.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Nueva cuenta"
         Top             =   265
         Width           =   375
      End
      Begin VB.Frame fraDatosMod 
         Height          =   4155
         Left            =   120
         TabIndex        =   26
         Top             =   4680
         Width           =   8970
         Begin VB.CommandButton cmdMantPoderes 
            Caption         =   "Mantenimiento Poderes"
            Height          =   375
            Left            =   6480
            TabIndex        =   72
            Top             =   1560
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.CommandButton cmdEnvioEstCta 
            Caption         =   "Remisión Estado de Cuenta"
            Enabled         =   0   'False
            Height          =   375
            Left            =   6480
            TabIndex        =   64
            Top             =   240
            Width           =   2295
         End
         Begin VB.CheckBox chkSubasta 
            Caption         =   "SUBASTA"
            Height          =   255
            Left            =   6360
            TabIndex        =   58
            Top             =   1200
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.ComboBox cboInstConvDep 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   2160
            Visible         =   0   'False
            Width           =   4695
         End
         Begin VB.CheckBox chkCovDep 
            Alignment       =   1  'Right Justify
            Caption         =   "Genera Relacion Convenio"
            Height          =   375
            Left            =   120
            TabIndex        =   56
            Top             =   1800
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Frame fraCtaAhoAboInt 
            Height          =   855
            Left            =   120
            TabIndex        =   51
            Top             =   960
            Width           =   2775
            Begin VB.CheckBox chkAbonoIntCtaAho 
               Alignment       =   1  'Right Justify
               Caption         =   "Abonar Int. Cta. Aho."
               Height          =   330
               Left            =   120
               TabIndex        =   52
               Top             =   120
               Width           =   1785
            End
            Begin SICMACT.TxtBuscar txtCtaAhoAboInt 
               Height          =   345
               Left            =   135
               TabIndex        =   53
               Top             =   435
               Width           =   2475
               _ExtentX        =   4366
               _ExtentY        =   609
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               EditFlex        =   -1  'True
            End
         End
         Begin VB.Frame frmCTS 
            Caption         =   "Datos Adicionales CTS"
            Enabled         =   0   'False
            Height          =   1815
            Left            =   120
            TabIndex        =   47
            Top             =   2280
            Visible         =   0   'False
            Width           =   8775
            Begin VB.TextBox txtRetiroAdicional 
               Height          =   330
               Left            =   1920
               MaxLength       =   100
               TabIndex        =   50
               Top             =   600
               Visible         =   0   'False
               Width           =   6765
            End
            Begin VB.CheckBox chkRetiroAdicional 
               Alignment       =   1  'Right Justify
               Caption         =   "Retiro Adicional"
               Height          =   315
               Left            =   240
               TabIndex        =   48
               Top             =   240
               Visible         =   0   'False
               Width           =   1875
            End
            Begin VB.Label lblRetiroAdicional 
               Caption         =   "Motivo Segun Ley"
               Height          =   375
               Left            =   240
               TabIndex        =   49
               Top             =   600
               Visible         =   0   'False
               Width           =   1455
            End
         End
         Begin VB.ComboBox cboPrograma 
            Height          =   315
            Left            =   990
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   1890
            Visible         =   0   'False
            Width           =   3825
         End
         Begin VB.TextBox txtAlias 
            Height          =   330
            Left            =   810
            MaxLength       =   100
            TabIndex        =   42
            Top             =   225
            Width           =   7845
         End
         Begin VB.TextBox txtMinFirmas 
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
            Height          =   325
            Left            =   4290
            TabIndex        =   36
            Top             =   1125
            Width           =   555
         End
         Begin VB.ComboBox cboInstitucion 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6360
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   720
            Width           =   2355
         End
         Begin VB.CheckBox chkOrdenPago 
            Alignment       =   1  'Right Justify
            Caption         =   "Orden Pago"
            Height          =   315
            Left            =   555
            TabIndex        =   3
            Top             =   705
            Width           =   1275
         End
         Begin VB.TextBox txtNumFirmas 
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
            Height          =   325
            Left            =   4290
            TabIndex        =   4
            Top             =   615
            Width           =   555
         End
         Begin SICMACT.TxtBuscar txtInstitucion 
            Height          =   330
            Left            =   6360
            TabIndex        =   38
            Top             =   1080
            Visible         =   0   'False
            Width           =   2325
            _ExtentX        =   4313
            _ExtentY        =   582
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
            TipoBusPers     =   1
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Programa :"
            Height          =   195
            Left            =   120
            TabIndex        =   43
            Top             =   1965
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Alias"
            Height          =   195
            Left            =   135
            TabIndex        =   41
            Top             =   360
            Width           =   330
         End
         Begin VB.Label lblInstitucion 
            Alignment       =   2  'Center
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
            Height          =   780
            Left            =   5070
            TabIndex        =   40
            Top             =   1470
            Visible         =   0   'False
            Width           =   3630
         End
         Begin VB.Label lbCTSI 
            AutoSize        =   -1  'True
            Caption         =   "Institución :"
            Height          =   195
            Left            =   5100
            TabIndex        =   39
            Top             =   1185
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "N° Firmas Min.:"
            Height          =   195
            Left            =   3120
            TabIndex        =   37
            Top             =   1185
            Width           =   1080
         End
         Begin VB.Label lblInst 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cta. :"
            Height          =   195
            Left            =   5100
            TabIndex        =   32
            Top             =   765
            Width           =   735
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "N° Firmas :"
            Height          =   195
            Left            =   3075
            TabIndex        =   31
            Top             =   750
            Width           =   795
         End
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         Height          =   325
         Left            =   3735
         TabIndex        =   1
         ToolTipText     =   "Busca cliente por nombre, documento o codigo"
         Top             =   265
         Width           =   375
      End
      Begin VB.Frame fraDatos 
         Height          =   1290
         Left            =   120
         TabIndex        =   13
         Top             =   615
         Width           =   8955
         Begin VB.Label lblVencimientoT 
            AutoSize        =   -1  'True
            Caption         =   "Vencimiento :"
            Height          =   195
            Left            =   6270
            TabIndex        =   46
            Top             =   960
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label lblVencimiento 
            Alignment       =   2  'Center
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
            Height          =   300
            Left            =   7395
            TabIndex        =   45
            Top             =   915
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Estado :"
            Height          =   195
            Left            =   3150
            TabIndex        =   30
            Top             =   615
            Width           =   585
         End
         Begin VB.Label lblEstado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Height          =   300
            Left            =   4155
            TabIndex        =   29
            Top             =   555
            Width           =   1905
         End
         Begin VB.Label lblMoneda 
            Alignment       =   2  'Center
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
            Height          =   300
            Left            =   4155
            TabIndex        =   28
            Top             =   900
            Width           =   1905
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Moneda :"
            Height          =   195
            Left            =   3150
            TabIndex        =   27
            Top             =   945
            Width           =   675
         End
         Begin VB.Label lblPlazo 
            Alignment       =   2  'Center
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
            Height          =   300
            Left            =   7395
            TabIndex        =   25
            Top             =   255
            Width           =   1035
         End
         Begin VB.Label lblFormaRetiro 
            Alignment       =   2  'Center
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
            Height          =   300
            Left            =   7395
            TabIndex        =   24
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label lblTipoCuenta 
            Alignment       =   2  'Center
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
            Height          =   300
            Left            =   4155
            TabIndex        =   23
            Top             =   225
            Width           =   1905
         End
         Begin VB.Label lblTipoTasa 
            Alignment       =   2  'Center
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
            Height          =   300
            Left            =   960
            TabIndex        =   22
            Top             =   525
            Width           =   1905
         End
         Begin VB.Label lblTasa 
            Alignment       =   2  'Center
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
            Height          =   300
            Left            =   960
            TabIndex        =   21
            Top             =   870
            Width           =   1905
         End
         Begin VB.Label lblApertura 
            Alignment       =   2  'Center
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
            Height          =   300
            Left            =   960
            TabIndex        =   20
            Top             =   195
            Width           =   1905
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cuenta :"
            Height          =   195
            Left            =   3150
            TabIndex        =   19
            Top             =   285
            Width           =   960
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Forma Retiro :"
            Height          =   195
            Left            =   6270
            TabIndex        =   18
            Top             =   645
            Width           =   1035
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Plazo (días) :"
            Height          =   195
            Left            =   6270
            TabIndex        =   17
            Top             =   285
            Width           =   930
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Tasa :"
            Height          =   195
            Left            =   105
            TabIndex        =   16
            Top             =   585
            Width           =   810
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tasa (%) :"
            Height          =   195
            Left            =   105
            TabIndex        =   15
            Top             =   900
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Apertura :"
            Height          =   195
            Left            =   105
            TabIndex        =   14
            Top             =   255
            Width           =   690
         End
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   661
         Texto           =   "Cuenta N°:"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label lblPrograma 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   6000
         TabIndex        =   55
         Top             =   300
         Width           =   3015
      End
      Begin VB.Label lblSubProd 
         Caption         =   "Sub Producto:"
         Height          =   255
         Left            =   4920
         TabIndex        =   54
         Top             =   330
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCapMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bConsulta As Boolean
Public nProducto As Producto
Dim nTitular As Integer
'****** VERIFICAR TIPOS DE DATOS
Dim nTipoCuenta As COMDConstantes.ProductoCuentaTipo
Dim nPersoneria As COMDConstantes.PersPersoneria
'*******************************
'By Capi Acta 014-2007
Dim nPorRetAdiCTS As Integer
Dim nRetAdiCTS As Integer
'By Capi 01042008
Dim nPorRetCTS As Integer

Dim nClientes As Integer
Dim nRepresentante As Integer
Dim sTitular As String
'By capi 21012009
Dim objPista As COMManejador.Pista
Dim sTpoCta As String
Dim bPosCli As Boolean
Dim lsCodPersConv As String 'GITU 2012-10-23

'** Modificado por AMDO 2013, según TI-ERS086-2013********************************
'Dim fsCargoCodModAgeManProPas As String '***Agregado por ELRO el 2013011, según OYP-RFC123-2012
Dim bPermisoCargo As Boolean '**Agregado por AMDO 20130726, según TI-ERS086-2013
'*** END AMDO *********************************************

'JUEZ 20130527 *********************
Dim fbRegistraEnvio As Boolean
Dim frsEnvEstCta As ADODB.Recordset
Dim fnModoEnvioEstCta As Integer
Dim fnDebitoMismaCta As Integer
'END JUEZ **************************
'***Agregado por elro el 20130725, según SATI INC1306130011****
Dim fnTpoPrograma As Integer
'***Fin Agregado por elro el 20130725, según SATI INC1306130011

' *** RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
Dim intPunteroPJ_NA As Integer ' Indica la personeria al cargar de la cuenta
Dim nTipoCuentaInicial As Integer ' Indica el tipo de cuenta al cargar cuenta
Dim bActualizarCuenta As Boolean
Dim fbMantPoderes As Boolean 'JUEZ 20151229

Private Function CuentaTitular() As Integer
Dim i As Integer, nFila As Integer, nCol As Integer
nFila = grdCliente.row
nCol = grdCliente.Col
nTitular = 0
nClientes = 0
nRepresentante = 0
For i = 1 To grdCliente.rows - 1
    If grdCliente.TextMatrix(i, 3) <> "" Then
        If CLng(Trim(Right(grdCliente.TextMatrix(i, 3), 4))) = gCapRelPersTitular Then
            nTitular = nTitular + 1
        ElseIf CLng(Trim(Right(grdCliente.TextMatrix(i, 3), 4))) = gCapRelPersRepTitular Then
            nRepresentante = nRepresentante + 1
        End If
    End If
    If grdCliente.TextMatrix(i, 1) <> "" Then
        nClientes = nClientes + 1
    End If
Next i
grdCliente.row = nFila
grdCliente.Col = nCol

' *** RIRO 20131102 SEGUN "CAMBIOS EN PODERES"

    'If nClientes = 0 Then
    '    cmdEliminar.Enabled = False
    '    cmdAgregar.Enabled = True
    'Else
    '    If nClientes = 1 And nProducto = gCapCTS Then
    '        cmdAgregar.Enabled = False
    '    End If
    '    cmdEliminar.Enabled = True
    'End If

End Function

Private Sub EvaluaTitular()
Dim i As Integer
If nClientes > 1 Then
    For i = 0 To cboInstitucion.ListCount - 1
        If CLng(Trim(Right(cboInstitucion.List(i), 4))) = gPrdCtaTpoIndist Or CLng(Trim(Right(cboInstitucion.List(i), 4))) = gPrdCtaTpoMancom Then
            'cboInstitucion.ListIndex = i
            Exit For
        End If
    Next i
Else
    For i = 0 To cboInstitucion.ListCount - 1
        If CLng(Trim(Right(cboInstitucion.List(i), 4))) = gPrdCtaTpoIndiv Then
            'cboInstitucion.ListIndex = i
            Exit For
        End If
    Next i
End If
If nPersoneria = gPersonaNat Then
    txtNumFirmas = Format$(nTitular, "#0")
Else
    'MAVM 20100111
    'cboInstitucion.ListIndex = 0
    txtNumFirmas = Format$(nRepresentante, "#0")
End If
End Sub

Private Sub ObtieneDatosCuenta(ByVal sCuenta As String, Optional bMuestraPos As Boolean = False)
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales ' NCapMantenimiento
    Dim rsCta As ADODB.Recordset, rsRel As ADODB.Recordset
    Dim nEstado As COMDConstantes.CaptacEstado
    Dim ssql As String
    Dim nTpoPrograma As Integer 'BRGO 20111220
    Dim clsMnto As COMDCaptaGenerales.DCOMCaptaMovimiento 'GITU 2012-10-23
    Dim lbResultadoPersoneria As Boolean 'MIOL 20121113, SEGUN RFC098-2012
    '***Agregado por ELRO el 20121015, según OYP-RFC101-2012****
    Dim lnRemBruCTS As Currency
    Dim oNCOMConstSistema As COMDConstSistema.NCOMConstSistema
    Dim lsCargos() As String
    Dim lsCargoCod As String
    Dim J As Integer
    '***Fin Agregado por ELRO el 20121015, según OYP-RFC101-2012
    '***Agregado por ELRO el 20130725, según SATI INC1306130011****
     fnTpoPrograma = 0
    '***Fin Agregado por ELRO el 20130725, según SATI INC1306130011
    Dim nFormaRetiro As CaptacPFFormaRetiro
    
    '***Agregado por ELRO el 20130111, según OYP-RFC123-2012****
    If Trim(txtCuenta.Age) = "" Then
        MsgBox "Debe ingresar el código de Agencia de la Cuenta", vbInformation, "¡Aviso!"
        Exit Sub
    End If
    If Not IsNumeric(txtCuenta.Age) Then
        MsgBox "Debe ingresar el código de Agencia de la Cuenta", vbInformation, "¡Aviso!"
        Exit Sub
    End If
    '***Fin Agregado por ELRO el 20130111, según OYP-RFC123-2012
    
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsCta = New ADODB.Recordset
    Set rsCta = clsMant.GetDatosCuenta(sCuenta)

    If Not (rsCta.EOF And rsCta.BOF) Then
    
        ' Agregado Por RIRO el 20131102, Proyecto Ahorro - Poderes
        Dim arregloReglas() As String
        Dim d As Variant
        arregloReglas = Split(IIf(IsNull(rsCta!cReglas), "", rsCta!cReglas), "-")
        limpiarReglas
        For Each d In arregloReglas
            grdReglas.AdicionaFila
            grdReglas.TextMatrix(grdReglas.row, 1) = d
        Next
        ' Fin RIRO
    
        nEstado = rsCta("nPrdEstado")
        lblApertura = Format$(rsCta("dApertura"), "dd mmm yyyy")
        txtAlias.Text = Trim(IIf(IsNull(rsCta("cAlias")), "", rsCta("cAlias")))
        lblEstado = rsCta("cEstado")
        lblMoneda = IIf(Mid(sCuenta, 9, 1) = "1", "Nacional", "Extranjera")
        lblTasa = Format$(ConvierteTNAaTEA(rsCta("nTasaInteres")), "#0.000")
        lblTipoCuenta = rsCta("cTipoCuenta")
        
        'MAVM 19112009 ***
        sTpoCta = rsCta("nPrdCtaTpo")
        'End ***
        nTipoCuentaInicial = sTpoCta ' RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
        lblTipoTasa = rsCta("cTipoTasa")
        nTipoCuenta = rsCta("nPrdCtaTpo")
        nPersoneria = rsCta("nPersoneria")
        If nProducto = gCapCTS Then
            nRetAdiCTS = rsCta("RetAdiCTS") 'By Capi Acta 014-2007
        End If
        txtNumFirmas.Text = Format$(rsCta("nFirmas"), "#0")
        txtMinFirmas.Text = Format$(IIf(IsNull(rsCta("nfirmasmin")), 1, rsCta("nfirmasmin")), "#0")
        txtMinFirmas.Tag = Format$(IIf(IsNull(rsCta("nfirmasmin")), 1, rsCta("nfirmasmin")), "#0")
        nTpoPrograma = rsCta("nTpoPrograma")
        '***Agregado por ELRO el 20130725, según SATI INC1306130011****
        fnTpoPrograma = rsCta("nTpoPrograma")
        '***Fin Agregado por ELRO el 20130725, según SATI INC1306130011
        Dim clsGen As COMDConstSistema.DCOMGeneral 'DGeneral
        Set clsGen = New COMDConstSistema.DCOMGeneral
        Select Case nProducto
            Dim nIndex As Integer
            Case gCapAhorros
            
                'Add by GITU 23-10-2012
                Set clsMnto = New COMDCaptaGenerales.DCOMCaptaMovimiento
                lsCodPersConv = clsMnto.DevolverCodInstConv(sCuenta)
                
                If nTpoPrograma = 8 Then chkCovDep.Visible = True 'JUEZ 20141202
                
                If lsCodPersConv <> "" Then
                    'chkCovDep.Visible = True 'JUEZ 20141202
                    chkCovDep.value = 1
                    cboInstConvDep.ListIndex = IndiceListaCombo(Me.cboInstConvDep, lsCodPersConv)
                    cboInstConvDep.Visible = True
                End If
                'End GITU
                
                chkOrdenPago.value = IIf(rsCta("bOrdPag"), 1, 0)
                Set rsRel = clsGen.GetConstante(gProductoCuentaTipo)
                Do While Not rsRel.EOF
                    cboInstitucion.AddItem UCase(rsRel("cDescripcion")) & Space(100) & rsRel("nConsValor")
                    If rsRel("nConsValor") = nTipoCuenta Then
                        nIndex = cboInstitucion.ListCount - 1
                    End If
                    rsRel.MoveNext
                Loop
                cboInstitucion.ListIndex = nIndex
                Dim i As Integer
                '--------------- COMENTADO PARA CMAC CUSCO --------------------
                '            If cboPrograma.ListCount > 0 Then
                '              For i = 0 To cboPrograma.ListCount - 1
                '                If Right(Trim(cboPrograma.List(i)), 1) = CStr(rsCta("nTpoPrograma")) Then
                '                    cboPrograma.ListIndex = i
                '                    Exit For
                '                End If
                '              Next i
                '            End If
                '---------------------------------------------------------------
                '*** BRGO 20111220 ************************************************
                Set rsRel = clsGen.GetConstante(gCaptacSubProdAhorros, , CStr(nTpoPrograma))
                lblPrograma = rsRel!cDescripcion
                '*** END BRGO *****************************************************
                Set rsRel = Nothing
                '***Agregado por ELRO el 20120719, según OYP-RFC080-2012
                chkOrdenPago.Enabled = True
                '***Fin Agregado por ELRO el 20120719*******************
            Case gCapPlazoFijo
                lblFormaRetiro = rsCta("cRetiro")
                lblPlazo = Format$(rsCta("nPlazo"), "#0")
                nFormaRetiro = rsCta("nFormaRetiro") 'JUEZ 20141202
                '********************************
                '**Modificado por ELRO 20110714, según acta 163-2011/TI-D
                Me.lblVencimientoT.Visible = True
                Me.lblVencimiento.Visible = True
                '********************************
                'MIOL RFC098-2012 ******************************
                If rsCta("nTpoPrograma") = 0 And lblFormaRetiro.Caption = "FINAL DEL PLAZO" Then
                    If rsCta("GarantDepCF") > 0 Then
                        Me.chkSubasta.Visible = False
                    Else
                        If rsCta("Subasta") > 0 Then
                            Me.chkSubasta.value = rsCta("Subasta")
                            Me.chkSubasta.Visible = True
                        Else
                            Me.chkSubasta.value = 0
                            Me.chkSubasta.Visible = True
                        End If
                    End If
                Else
                    Me.chkSubasta.Visible = False
                End If
                'END MIOL **************************************
                Me.lblVencimiento.Caption = Format(DateAdd("d", rsCta("nPlazo"), rsCta("dRenovacion")), gsFormatoFechaView)
                If rsCta!cCtaCodAbono <> "" Then
                    chkAbonoIntCtaAho.value = 1
                    txtCtaAhoAboInt.Text = rsCta!cCtaCodAbono
                End If
                Set rsRel = clsGen.GetConstante(gProductoCuentaTipo)
                Do While Not rsRel.EOF
                    cboInstitucion.AddItem UCase(rsRel("cDescripcion")) & Space(100) & rsRel("nConsValor")
                    If rsRel("nConsValor") = nTipoCuenta Then
                        nIndex = cboInstitucion.ListCount - 1
                    End If
                    rsRel.MoveNext
                Loop
                cboInstitucion.ListIndex = nIndex
                '*** BRGO 20111220 ************************************************
                Set rsRel = clsGen.GetConstante(gCaptacSubProdPlazoFijo, , CStr(nTpoPrograma))
                lblPrograma = rsRel!cDescripcion
                '*** END BRGO *****************************************************
                Set rsRel = Nothing
                cmdRenovaciones.Visible = True '***Agregado por ELRO el 20130111, según OYP-RFC123-2012
            Case gCapCTS
                cboInstitucion.AddItem "INDIVIDUAL" 'UCase(rsCta!cInstitucion)
                cboInstitucion.ListIndex = 0
                txtInstitucion.Text = UCase(rsCta!cCodInst)
                lblInstitucion.Caption = UCase(rsCta!cInstitucion)
                '*** BRGO 20111220 ************************************************
                Set rsRel = clsGen.GetConstante(gCaptacSubProdCTS, , CStr(nTpoPrograma))
                lblPrograma = rsRel!cDescripcion
                Set rsRel = Nothing
                '*** END BRGO *****************************************************
                '***Agregado por ELRO el 20121015, según OYP-RFC101-2012
                If bMuestraPos = False Then
                    clsMant.obtenerHistorialCaptacSueldosCTS sCuenta, lnRemBruCTS
                    lblTotalRemuneraciones = Format$(lnRemBruCTS, "#,##0.00")
                    'Comentado x JUEZ 20140305
                    'Set oNCOMConstSistema = New COMDConstSistema.NCOMConstSistema
                    'lsCargoCod = oNCOMConstSistema.LeeConstSistema(gConstSistRHCargoCodCTSNoActiva)
                    'lsCargos = Split(lsCargoCod, ",")
                    'For j = 0 To UBound(lsCargos)
                    '    If lsCargos(j) = gsCodCargo Then
                    '        cmdCtsNoActivo.Visible = True
                    '        If rsCta!nTpoPrograma = 2 Or rsCta("cEstado") = "CANCELADA" Then
                    '            cmdCtsNoActivo.Enabled = False
                    '        Else
                    '            cmdCtsNoActivo.Enabled = True
                    '        End If
                    '        Exit For
                    '    End If
                    'Next j
                End If
                '***Fin Agregado por ELRO el 20121015*******************
                cmdTasaHist.Visible = True 'JUEZ 20140305
                txtCtaAhoAboInt.Visible = False 'JUEZ 20140305
                chkAbonoIntCtaAho.Visible = False 'JUEZ 20140305
            End Select
        Set clsGen = Nothing
        rsCta.Close
        'Set rsCta = clsMant.GetProductoPersona(sCuenta) RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
        Set rsCta = clsMant.GetProductoPersona(sCuenta, gCapRelPersPromotor)
        
        If Not (rsCta.EOF And rsCta.BOF) Then
        
            Set grdCliente.Recordset = rsCta
            
            ' *** RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
            
            Dim nCount As Integer
            For i = 1 To grdCliente.rows - 1
                If val(Trim(grdCliente.TextMatrix(i, 4))) > 1 Then
                    nCount = nCount + 1
                End If
            Next
            If nCount > 0 Then
                intPunteroPJ_NA = 1
            Else
                intPunteroPJ_NA = 0
            End If
            nCount = 0
            For i = 1 To grdCliente.rows - 1
                If Len(Trim(grdCliente.TextMatrix(i, 6))) = 0 Then
                    nCount = nCount + 1
                End If
            Next
            If nCount > 0 Then
                bActualizarCuenta = True
            Else
                bActualizarCuenta = False
            End If
            
            ' *** FIN RIRO
            
            cmdBuscar.Enabled = False
            cmdNuevaBus.Enabled = True
            cmdRegFirmas.Enabled = True
            If (nEstado <> gCapEstActiva) Then
                MsgBox "Esta cuenta no se podrá editar por tener un estado no activa.", vbInformation, "Aviso"
                cmdEditar.Enabled = False
            Else
                cmdEditar.Enabled = True
                If bMuestraPos = False Then
                    cmdEditar.SetFocus
                End If
            End If
        Else
            MsgBox "Cuenta no posee relacion con Persona", vbExclamation, "Aviso"
            txtCuenta.SetFocusCuenta
        End If

        Set clsGen = New COMDConstSistema.DCOMGeneral
        
        ' *** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
        
                'If nPersoneria = gPersonaNat Then
                '    'Set rsRel = clsGen.GetConstante(gCaptacRelacPersona, gCapRelPersRepSuplente & "," & gCapRelPersRepTitular & "," & gCapRelPersTitular)
                '    Set rsRel = clsGen.GetConstante(gCaptacRelacPersona, gCapRelPersRepSuplente & "," & gCapRelPersRepTitular)
                'Else
                '    'Set rsRel = clsGen.GetConstante(gCaptacRelacPersona, gCapRelPersTitular & "," & gCapRelPersApoderado)
                '    Set rsRel = clsGen.GetConstante(gCaptacRelacPersona, gCapRelPersApoderado)
                'End If
        
        ' *** FIN RIRO
        
        ' Recargar el combo del grid segun sea el caso
        If intPunteroPJ_NA = 0 Then
            Set rsRel = clsGen.GetConstante(gCaptacRelacPersona, "13,14")
        Else
            Set rsRel = clsGen.GetConstante(gCaptacRelacPersona, "11,13,14")
            
            'Validando cantidad de titulares
            'Se aplica en cuentas con personeria juridica que intervenga mas de un titular,
            'La columna "7" determina si la fila es editable o no
            nCount = 0
            For i = 1 To grdCliente.rows - 1
                If val(Trim(grdCliente.TextMatrix(i, 4))) = 1 And val(grdCliente.TextMatrix(i, 7)) = 10 Then
                    grdCliente.TextMatrix(i, 7) = ""
                End If
            Next
            
        End If
        
        grdCliente.CargaCombo rsRel
        Set clsGen = Nothing
        Set clsMant = Nothing
        Set rsCta = Nothing
        CuentaTitular
    Else
        MsgBox "Cuenta no existe", vbInformation, "Aviso"
        txtCuenta.SetFocusCuenta
    End If
    Set clsMant = Nothing
End Sub

Public Sub Inicia(ByVal nProd As Producto, Optional bCons As Boolean = False)
    '***Agregado por ELRO el 20130110, según OYP-RFC123-2012
    Dim oNCOMConstSistema As COMDConstSistema.NCOMConstSistema
    Dim lsCargos() As String
    Dim J As Integer
    '***Fin Agregado por ELRO el 20130110*******************
    nProducto = nProd
    bConsulta = bCons
    Select Case nProd
        Case gCapAhorros
            Label5.Visible = False
            Label6.Visible = False
            lblPlazo.Visible = False
            lblFormaRetiro.Visible = False
            lblInst = "Tipo Cta :"
            chkOrdenPago.Visible = True
            chkOrdenPago.Enabled = False
            chkAbonoIntCtaAho.Visible = False
            txtCtaAhoAboInt.Visible = False
            cmdMantPoderes.Visible = True 'JUEZ 20151229
            txtCuenta.Prod = Trim(str(gCapAhorros))
            Me.Caption = "Captaciones - Mantenimiento - Ahorros"
        
            '*** AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
            Label10.Visible = False
            Label4.Visible = False
            txtNumFirmas.Visible = False
            txtMinFirmas.Visible = False
            ' *** FIN RIRO
        
            Dim oCons As COMDConstantes.DCOMConstantes 'DConstante
            Set oCons = New COMDConstantes.DCOMConstantes
            Dim rs As ADODB.Recordset
            Set rs = New ADODB.Recordset
            'Label12.Visible = True
            'cboPrograma.Visible = True
            Set rs = oCons.GetConstante(2030, False, , True, , "2030")
            Me.lblVencimiento.Visible = False
            Me.lblVencimientoT.Visible = False
            'Me.cboPrograma.Clear
            While Not rs.EOF
            cboPrograma.AddItem rs.Fields(1) & Space(100) & rs.Fields(0)
            rs.MoveNext
            Wend
'        If cboPrograma.ListIndex > 0 Then
'            cboPrograma.ListIndex = 0
'        End If
        'by capi 21012009
            gsOpeCod = gAhoMantCuenta
        '
            'chkCovDep.Visible = True 'JUEZ 20141202
            
            IniciaComboConvDep 9
        Case gCapPlazoFijo
            Label5.Visible = True
            Label6.Visible = True
            lblPlazo.Visible = True
            lblFormaRetiro.Visible = True
            lblInst = "Tipo Cta :"
            chkOrdenPago.Visible = False
            chkAbonoIntCtaAho.Visible = True
            txtCtaAhoAboInt.Visible = True
            txtCuenta.Prod = Trim(str(gCapPlazoFijo))
            Me.Caption = "Captaciones - Mantenimiento - Plazo Fijo"
            Me.fraCtaAhoAboInt.Enabled = False
            Me.lblVencimiento.Visible = True
            Me.lblVencimientoT.Visible = True
            Me.chkSubasta.Visible = True 'MIOL 20121109, SEGUN RFC098-2012-B
            cmdMantPoderes.Visible = False 'JUEZ 20151229
            
            ' *** AGREGADO RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
            Label10.Visible = False
            Label4.Visible = False
            txtNumFirmas.Visible = False
            txtMinFirmas.Visible = False
            ' *** FIN RIRO
            
           gsOpeCod = gPFMantCuenta
           If (gsCodCargo = "002001" Or gsCodCargo = "002002" Or gsCodCargo = "002003" _
              Or gsCodCargo = "003001" Or gsCodCargo = "003002" Or gsCodCargo = "004001" _
              Or gsCodCargo = "004002" Or gsCodCargo = "006005") Then
                Me.fraCtaAhoAboInt.Enabled = True
                txtCtaAhoAboInt.Enabled = False
           End If
        Case gCapCTS
            Label5.Visible = False
            Label6.Visible = False
            lblPlazo.Visible = False
            lblFormaRetiro.Visible = False
            lblInst = "Institución :"
            cboInstitucion.Visible = True
            chkOrdenPago.Visible = False
            cmdAgregar.Visible = False
            cmdEliminar.Visible = False
            txtNumFirmas.Enabled = False
            chkAbonoIntCtaAho.Visible = False
            txtCtaAhoAboInt.Visible = False
            txtCuenta.Prod = Trim(str(gCapCTS))
            Me.Caption = "Captaciones - Mantenimiento - CTS"
            lblInst.Caption = "Tipo Cta :"
        
            Me.txtInstitucion.Visible = True
            txtInstitucion.Text = ""
        
            lbCTSI.Visible = True
            lbCTSI.Caption = ""
            lblInstitucion.Visible = True
            lblInstitucion.Caption = ""
        
            Me.lblVencimiento.Visible = False
            Me.lblVencimientoT.Visible = False
            
            ' *** AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
            Label10.Visible = False
            Label4.Visible = False
            txtNumFirmas.Visible = False
            txtMinFirmas.Visible = False
            ' *** FIN RIRO
            
            'by capi 21012009
            gsOpeCod = gCTSMantCuenta
            '
            '***Agregado por ELRO el 20121013, según OYP-RFC101-2012
            FRHistRemBruCts.Visible = True
            lblRemuneraciones.Visible = True
            lblTotalRemuneraciones.Visible = True
            cmdHistorico.Visible = True
            '***Fin Agregado por ELRO el 20121013*******************
            
            fraReglasPorderes.Visible = False ' AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
            cmdMantPoderes.Visible = False 'JUEZ 20151229
            'JUEZ 20151114 **********************************************
            Dim clsDef As New COMNCaptaGenerales.NCOMCaptaDefinicion
            lblRemuneraciones.Caption = "(Total " & CStr(clsDef.GetCapParametroNew(gCapCTS, 0)!nUltRemunBrutas) & " Últimas Remu. Brutas)"
            Set clsDef = Nothing
            'END JUEZ ***************************************************
    End Select

    If bConsulta Then
        cmdEditar.Visible = False
        cmdGrabar.Visible = False
        cmdCancelar.Visible = False
    Else
        cmdEditar.Visible = True
        cmdGrabar.Visible = True
        cmdCancelar.Visible = True
        cmdEditar.Enabled = False
        cmdGrabar.Enabled = False
        cmdCancelar.Enabled = False
    
    End If
    fraDatos.Enabled = False
    fraDatosMod.Enabled = False
    fraCliente.Enabled = False
    txtCuenta.CMAC = gsCodCMAC
    txtCuenta.Age = Right(gsCodAge, 2)
    txtCuenta.EnabledCMAC = False
    txtCuenta.EnabledAge = False
    
'AMDO TI-ERS086-2013 20130726 *************************************************
'    '***Agregado por ELRO el 20130110, según OYP-RFC123-2012
'    Set oNCOMConstSistema = New COMDConstSistema.NCOMConstSistema
'    fsCargoCodModAgeManProPas = oNCOMConstSistema.LeeConstSistema(gConstSistRHCargoCodModAgeManProPas)
'    Set oNCOMConstSistema = Nothing
'    lsCargos = Split(fsCargoCodModAgeManProPas, ",")
'    For J = 0 To UBound(lsCargos)
'        If lsCargos(J) = gsCodCargo Then
'            txtCuenta.EnabledAge = True
'            txtCuenta.Age = ""
'            Exit For
'        End If
'    Next J
'    '***Fin Agregado por ELRO el 20130110*******************

    Dim oGen As COMDConstSistema.DCOMGeneral
    Set oGen = New COMDConstSistema.DCOMGeneral
    'bPermisoCargo = oGen.VerificaExistePermisoCargo(gsCodCargo, PermisoCargos.gMantProdPasivos)
     bPermisoCargo = oGen.VerificaExistePermisoCargo(gsCodCargo, PermisoCargos.gMantProdPasivos, gsCodPersUser)  ' RIRO20141027 ERS159
    If bPermisoCargo Then
        txtCuenta.EnabledAge = True
        txtCuenta.Age = ""
    End If
'END AMDO
    
    txtCuenta.EnabledProd = False
    cmdAgregar.Enabled = False
    cmdEliminar.Enabled = False
    cmdNuevaBus.Enabled = False
    cmdRegFirmas.Enabled = False
    Set oCons = Nothing
    'JUEZ 20130527 ****************
    If nProd = gCapCTS Then
        cmdEnvioEstCta.top = 240
    ElseIf nProd = gCapAhorros Then
        cmdEnvioEstCta.top = 1125
    ElseIf nProd = gCapPlazoFijo Then 'JUEZ 20130724
        cmdEnvioEstCta.top = 2040
    End If
    fbRegistraEnvio = False
    Set frsEnvEstCta = Nothing
    fnModoEnvioEstCta = 0
    fnDebitoMismaCta = 0
    'END JUEZ *********************
    Me.Show 1
End Sub

'MAVM 19112009 ***
Private Sub cboInstitucion_Click()

' COMENTADO POR RIRO
    'If fraDatosMod.Enabled = True Then
    '    If Trim(Right(cboInstitucion, 4)) <> (sTpoCta) And nProducto <> gCapCTS Then
    '        MsgBox "No se permite cambiar el Tipo de Cta", vbCritical
    '        cmdCancelar_Click
    '    End If
    'End If
' END RIRO
End Sub
'***

Private Sub chkAbonoIntCtaAho_Click()
    If chkAbonoIntCtaAho.value = 1 Then
        txtCtaAhoAboInt.Enabled = True
    Else
        txtCtaAhoAboInt.Enabled = False
        txtCtaAhoAboInt.Text = ""
    End If
End Sub

Private Sub chkCovDep_Click()
    If chkCovDep.value = 1 Then
        cboInstConvDep.Visible = True
    Else
        cboInstConvDep.Visible = False
    End If
End Sub

'By Capi Acta 014-2007
Private Sub ChkRetiroAdicional_Click()
    Dim clsGen As nCapDefinicion
    Set clsGen = New nCapDefinicion
    nPorRetAdiCTS = clsGen.GetCapParametro(gPorRetiroAdicionalCTS)
    txtRetiroAdicional = nPorRetAdiCTS & " % Vivienda Sg Carta "
    'By Capi 01042008
    nPorRetCTS = clsGen.GetCapParametro(2021) 'luego colocar en dcomvalores
End Sub
Private Sub cmdAgregar_Click()
    
' *** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
    
        'If grdCliente.TextMatrix(grdCliente.Rows - 1, 1) <> "" Then
        '    If Trim(Right(cboInstitucion, 4)) = 0 Then
        '       MsgBox "Cuenta Individual solo permite un Participante", vbInformation, "Aviso"
        '       Exit Sub
        '    End If
        'End If
        '
        'grdCliente.AdicionaFila
        'grdCliente.TextMatrix(grdCliente.Rows - 1, 4) = "N"
        'grdCliente.SetFocus
        'SendKeys "{ENTER}"
        'grdCliente.lbEditarFlex = True
        '
        'Dim i As Integer, numfirmas As Integer
        'numfirmas = 0
        'For i = 1 To grdCliente.Rows - 1
        '    If grdCliente.TextMatrix(i, 5) = "SI" Then
        '        numfirmas = numfirmas + 1
        '    End If
        'Next i
        'Label4.Tag = numfirmas
        'txtMinFirmas.Text = CStr(numfirmas)
        
' *** AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"

    grdCliente.AdicionaFila
    Dim i As Integer
    Dim intPJ As Integer
    intPJ = -1
    'Verifica si hay personas jurídicas dentro del grid
    For i = 1 To grdCliente.rows - 1
        If grdCliente.TextMatrix(i, 6) = "PJ" Then
            intPJ = i
        End If
    Next
    'Aplica el grid de reglas para las PN que cumplan la condicion.
    If intPJ = -1 Then
        intPunteroPJ_NA = 0
        If grdCliente.rows > 2 Then
            grdCliente.ColumnasAEditar = "X-1-X-3-X-5-6-X-X-9"
            
        Else
            grdCliente.ColumnasAEditar = "X-1-X-3-X-5-6-X-X-X"
            
        End If
    Else
        
    End If
    grdCliente.lbEditarFlex = True
    seleccionarTipoCuentaXregla
    grdCliente.SetFocus
    SendKeys "{ENTER}"
    
' *** FIN RIRO

End Sub

' *** AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"

Private Sub cmdAgregarRegla_Click()
    Dim X As Integer
    Dim i As Integer
    Dim strRegla As String
    'Verificando que cada interviniente en la apertura de la cuenta tenga una relacion: Titular, apoderado, etc.
    For i = 1 To grdCliente.rows - 1
        If Me.grdCliente.TextMatrix(i, 3) = "" Then
            MsgBox "Seleccione la relación de cada persona", vbInformation, "Aviso"
            Exit Sub
        End If
    Next
    i = 0
    'Verifica que las letras marcadas en el control listchek, esten dentro del 'Grid cliente'
    For X = 0 To lsLetras.ListCount - 1
        If lsLetras.Selected(X) = True Then
          lsLetras.Selected(X) = False
            If existeLetraEnSocio(UCase(Chr(65 + X))) = True Then
                If strRegla = "" Then
                    strRegla = UCase(Chr(65 + X))
                Else
                    strRegla = strRegla & "+" & UCase(Chr(65 + X))
                End If
            Else
                MsgBox "Verificar los grupos asignados a cada persona", vbExclamation, "Mensaje"
                Exit Sub
            End If
            i = i + 1
        End If
    Next
    If i = 0 Then
        MsgBox "Debe seleccionar uno o mas grupos antes de presionar este boton", vbExclamation, "Mensaje"
        Exit Sub
    End If
    If existeRegla(strRegla) Then
        MsgBox "Regla ya existe", vbExclamation, "Mensaje"
        Exit Sub
    End If
    grdReglas.AdicionaFila
    grdReglas.SetFocus
    grdReglas.TextMatrix(grdReglas.row, 1) = strRegla
    seleccionarTipoCuentaXregla

End Sub

' *** FIN RIRO

Private Sub cmdBuscar_Click()
Dim clsPers As COMDPersona.UCOMPersona 'UPersona

Set clsPers = New COMDPersona.UCOMPersona
Set clsPers = frmBuscaPersona.Inicio

If Not clsPers Is Nothing Then
    Dim sPers As String
    Dim rsPers As New ADODB.Recordset
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
    Dim sCta As String
    Dim sRelac As String * 15
    Dim sEstado As String
    Dim clsCuenta As UCapCuenta
    Dim lsCargos() As String '***Agregado por ELRO el 20130110, según OYP-RFC123-2012
    Dim J As Integer '***Agregado por ELRO el 20130110, según OYP-RFC123-2012
    sPers = clsPers.sPersCod
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsPers = clsCap.GetCuentasPersona(sPers, nProducto, , , , , gsCodAge)

'***Modificado por AMDO 20130726,según TI-ERS086-2013*****************************
'    '***Modificado por ELRO el 20130110, según OYP-RFC123-2012
'    lsCargos = Split(fsCargoCodModAgeManProPas, ",")
'    For J = 0 To UBound(lsCargos)
'        If lsCargos(J) = gsCodCargo Then
'            txtCuenta.Age = ""
'            Set rsPers = Nothing
'            Set rsPers = clsCap.GetCuentasPersona(sPers, nProducto, , , , , txtCuenta.Age)
'            Exit For
'        End If
'    Next J
'    '***Fin Modificado por ELRO el 20130110*******************
    Dim oGen As COMDConstSistema.DCOMGeneral
    Set oGen = New COMDConstSistema.DCOMGeneral
    If bPermisoCargo Then
            txtCuenta.Age = ""
            Set rsPers = Nothing
            Set rsPers = clsCap.GetCuentasPersona(sPers, nProducto, , , , , txtCuenta.Age)
    End If
    '***END AMDO**********************************************************************

    Set clsCap = Nothing
    If Not (rsPers.EOF And rsPers.EOF) Then
        Do While Not rsPers.EOF
            sCta = rsPers("cCtaCod")
            sRelac = rsPers("cRelacion")
            sEstado = Trim(rsPers("cEstado"))
            frmCapMantenimientoCtas.lstCuentas.AddItem sCta & Space(2) & sRelac & Space(2) & sEstado
            rsPers.MoveNext
        Loop
        Set clsCuenta = New UCapCuenta
        Set clsCuenta = frmCapMantenimientoCtas.Inicia
        If clsCuenta Is Nothing Then
        Else
            If clsCuenta.sCtaCod <> "" Then
                txtCuenta.Cuenta = Mid(clsCuenta.sCtaCod, 9, 10)
                '***Modificado por ELRO el 20130111, según OYP-RFC123-2012
                If Trim(txtCuenta.Age) = "" Then
                    txtCuenta.Age = Mid(clsCuenta.sCtaCod, 4, 2)
                End If
                '***Fin Modificado por ELRO el 20130111*******************
                txtCuenta.SetFocusCuenta
                SendKeys "{Enter}"
            End If
        End If
        Set clsCuenta = Nothing
    Else
        MsgBox "Persona no posee ninguna cuenta de captaciones o no tiene cuentas en la agencia.", vbInformation, "Aviso"
    End If
    rsPers.Close
    Set rsPers = Nothing
End If
Set clsPers = Nothing
txtCuenta.SetFocusCuenta
End Sub

Private Sub cmdCancelar_Click()
fraDatosMod.Enabled = False
fraCliente.Enabled = False
cmdEditar.Enabled = True
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
cmdAgregar.Enabled = False
cmdEliminar.Enabled = False
cmdSalir.Enabled = True
cmdNuevaBus.Enabled = True
cmdRegFirmas.Enabled = True
cboInstitucion.Clear
' AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
limpiarReglas
'fraReglasPorderes.Enabled = False 'RIRO20140710 ERS098
HabilitarPoderes False 'RIRO20140710 ERS098
grdCliente.lbEditarFlex = False
nTipoCuentaInicial = -1
txtRetiroAdicional.Text = ""
' END RIRO
'JUEZ 20130527 *************
fbRegistraEnvio = False
Set frsEnvEstCta = Nothing
fnModoEnvioEstCta = 0
fnDebitoMismaCta = 0
'END JUEZ ******************
ObtieneDatosCuenta txtCuenta.NroCuenta
'fsCargoCodModAgeManProPas = "" '***Agregado por ELRO el 20130111, según OYP-RFC123-2012'AMDO20130822
'chkCovDep.value = 0
'cboInstConvDep.Visible = False
fbMantPoderes = False 'JUEZ 20151229
End Sub

Private Sub CmdEditar_Click()
fraDatosMod.Enabled = True
fraCliente.Enabled = True
cmdEditar.Enabled = False
cmdGrabar.Enabled = True
cmdCancelar.Enabled = True
cmdAgregar.Enabled = True
cmdEliminar.Enabled = True
cmdSalir.Enabled = False
txtCuenta.Enabled = False
cmdBuscar.Enabled = False
cmdNuevaBus.Enabled = False
cmdRegFirmas.Enabled = False

' AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
If intPunteroPJ_NA = 0 Then
    If bActualizarCuenta Then
        'fraReglasPorderes.Enabled = True RIRO20140710 ERS098
        HabilitarPoderes True 'RIRO20140710 ERS098
        cmdAgregar.Enabled = False
        cmdEliminar.Enabled = False
    Else
        'fraReglasPorderes.Enabled = False RIRO20140710 ERS098
        HabilitarPoderes False 'RIRO20140710 ERS098
        If nProducto <> gCapCTS Then
            cmdAgregar.Enabled = False
            cmdEliminar.Enabled = False
            'cmdGrabar.Enabled = False
        End If
    End If
Else
    'fraReglasPorderes.Enabled = True RIRO20140710 ERS098
    HabilitarPoderes True 'RIRO20140710 ERS098
End If
' END RIRO

'Add By GITU 23-10-2012
If lsCodPersConv <> "" Then
    chkCovDep.value = 1
    cboInstConvDep.ListIndex = IndiceListaCombo(Me.cboInstConvDep, lsCodPersConv)
End If
'End GITU
End Sub

Private Sub cmdEliminar_Click()
Dim nIndex As Integer
nIndex = grdCliente.row
If grdCliente.TextMatrix(nIndex, 3) <> "" Then
    If CLng(Right(grdCliente.TextMatrix(nIndex, 3), 2)) = gCapRelPersTitular Then
        MsgBox "No es posible eliminar al titular de la cuenta.", vbInformation, "Aviso"
        Exit Sub
    End If
End If
If MsgBox("¿¿Está seguro de eliminar a la persona de la relación??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    grdCliente.EliminaFila nIndex
    CuentaTitular
    EvaluaTitular
    Dim i As Integer, numfirmas As Integer
    numfirmas = 0
    For i = 1 To grdCliente.rows - 1
        If grdCliente.TextMatrix(i, 5) = "SI" Then
            numfirmas = numfirmas + 1
        End If
    Next i
    Label4.Tag = numfirmas
    
    txtMinFirmas.Text = CStr(numfirmas)
End If
seleccionarTipoCuentaXregla ' AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
End Sub

'JUEZ 20130527 **************************************
Private Sub cmdEnvioEstCta_Click()
    frmEnvioEstadoCta.MantenimientoCap txtCuenta.NroCuenta, grdCliente.GetRsNew()
    fbRegistraEnvio = frmEnvioEstadoCta.RegistraEnvio
    Set frsEnvEstCta = frmEnvioEstadoCta.RecordSetDatos
    fnModoEnvioEstCta = frmEnvioEstadoCta.ModoEnvioEstCta
    fnDebitoMismaCta = frmEnvioEstadoCta.DebitoMismaCta
End Sub
'END JUEZ *******************************************

Private Sub cmdGrabar_Click()
Dim nFirmas As Long
Dim sInstitucion As String, sCuentaAboInt As String, sMovNro As String
Dim bOrdPag As Boolean

'Verifica si el numero de firmas corresponde al numero de titulares
nFirmas = CLng(txtNumFirmas)

bOrdPag = chkOrdenPago.value

' *** COMENTADO POR RIRO RIRO 20131102 SEGUN "CAMBIOS EN PODERES"

        ''Verifica si el numero de firmas corresponde al numero de titulares
        'If txtNumFirmas = "" Then
        '    MsgBox "Número de firmas no válido", vbInformation, "Aviso"
        '    txtNumFirmas.SetFocus
        '    Exit Sub
        'End If
        
' *** FIN RIRO

If nProducto = gCapCTS Then
    'sInstitucion = Right(cboInstitucion.Text, 13)
    sInstitucion = txtInstitucion.psCodigoPersona
    
    If sInstitucion = "" Then
        MsgBox "Institución No Válida", vbInformation, "Aviso"
        Exit Sub
    End If
    nTipoCuenta = gPrdCtaTpoIndiv
    
    ' *** AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
    grdReglas.AdicionaFila
    grdReglas.TextMatrix(1, 1) = "A"
    ' *** FIN RIRO
    
ElseIf nProducto = gCapPlazoFijo Then
    nTipoCuenta = CLng(Trim(Right(cboInstitucion.Text, 4)))
    If chkAbonoIntCtaAho.value = 1 Then
        sCuentaAboInt = txtCtaAhoAboInt.Text
        If sCuentaAboInt = "" Then
            MsgBox "Cuenta de Abono de Intereses No Válida", vbInformation, "Aviso"
            Exit Sub
        End If
    Else
        sCuentaAboInt = ""
    End If
Else
    nTipoCuenta = CLng(Trim(Right(cboInstitucion.Text, 4)))
    'Add GITU 23-10-2012
    If chkCovDep.value = 1 Then
        lsCodPersConv = Trim(Right(cboInstConvDep.Text, 13))
    Else
        lsCodPersConv = ""
    End If
    'End GITU
End If

' *** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"

        'nFirmas = CLng(txtNumFirmas)
        'If nFirmas = 0 Then
        '    MsgBox "Número de Firmas no puede ser 0.", vbInformation, "Aviso"
        '    'txtNumFirmas.SetFocus
        '    Exit Sub
        'End If
        '
        'If val(txtMinFirmas.Text) = 0 Then
        '    MsgBox "Número Mínimo de Firmas Obligatorias no puede ser 0.", vbInformation, "Aviso"
        '    txtMinFirmas.SetFocus
        '    Exit Sub
        'End If

' *** END RIRO

If nPersoneria = gPersonaNat Then
    If nTitular = 0 Then
        MsgBox "No existen titulares en la cuenta.", vbInformation, "Aviso"
        grdCliente.SetFocus
        Exit Sub
    End If
    
    '******CAMBIO PARA FIRMAS
'    If nTipoCuenta = gPrdCtaTpoMancom Or nTipoCuenta = gPrdCtaTpoIndiv Then
'        If nTitular <> nFirmas Then 'Valida # de Firmas
'            MsgBox "Número de firmas difiere del número de titulares.", vbInformation, "Aviso"
'            txtNumFirmas.SetFocus
'            Exit Sub
'        End If
'    Else
'        If nFirmas >= nTitular And nTitular > 1 Then  'Valida # de Firmas
'            MsgBox "Número de firmas no corresponde con el tipo de cuenta.", vbInformation, "Aviso"
'            txtNumFirmas.SetFocus
'            Exit Sub
'        End If
'    End If
Else
    If nTitular > 1 Then
        MsgBox "La cuenta con personería jurídica no debe tener mas de un titular.", vbInformation, "Aviso"
        grdCliente.SetFocus
        Exit Sub
    End If
    If nRepresentante = 0 Then
        MsgBox "No existen representantes en la cuenta.", vbInformation, "Aviso"
        grdCliente.SetFocus
        Exit Sub
    End If
    '******CAMBIO PARA FIRMAS
    
'    If nTipoCuenta = gPrdCtaTpoMancom Then
'        If nRepresentante + 1 < nFirmas Then 'Valida # de Firmas
'            MsgBox "Número de firmas excede del número de posibles Representantes a firmar.", vbInformation, "Aviso"
'            txtNumFirmas.SetFocus
'            Exit Sub
'        End If
'    Else
'        If nFirmas > nRepresentante Then   'Valida # de Firmas
'            MsgBox "Número de firmas no corresponde con el tipo de cuenta.", vbInformation, "Aviso"
'            txtNumFirmas.SetFocus
'            Exit Sub
'        End If
'    End If
End If

' *** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"

        'If nTitular = 1 Then 'Valida Tipo de Cuenta
        '    If nTipoCuenta <> gPrdCtaTpoIndiv And nPersoneria = gPersonaNat Then
        '        MsgBox "La cuenta posee un solo titular por lo que debe ser Individual.", vbInformation, "Aviso"
        '        cboInstitucion.SetFocus
        '        Exit Sub
        '    ElseIf nTipoCuenta = gPrdCtaTpoIndiv And nPersoneria <> gPersonaNat And val(txtMinFirmas.Text) > 1 Then
        '        MsgBox "La cuenta requiere mas de una firma por lo que debe ser Mancomunada.", vbInformation, "Aviso"
        '        Exit Sub
        '    ElseIf nTipoCuenta = gPrdCtaTpoMancom And nPersoneria <> gPersonaNat And val(txtMinFirmas.Text) = 1 Then
        '        MsgBox "La cuenta requiere una firma por lo que debe NO DEBE ser Mancomunada.", vbInformation, "Aviso"
        '        Exit Sub
        '    End If
        'Else
        '    If nTipoCuenta <> gPrdCtaTpoIndist And nTipoCuenta <> gPrdCtaTpoMancom Then
        '        MsgBox "La cuenta posee más de un titular por lo que no puedw ser Individual.", vbInformation, "Aviso"
        '        cboInstitucion.SetFocus
        '        Exit Sub
        '    End If
        'End If

' *** AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
    
    ' ** Validando Mayoria de edad
     
    Dim oPersonaTemp As COMNPersona.NCOMPersona
    Dim iTemp, nMenorEdad, J, nLetra, nContar, nLetraMax As Integer
    Dim bOrden As Boolean
    
    Set oPersonaTemp = New COMNPersona.NCOMPersona
    
    For iTemp = 1 To grdCliente.rows - 1
        If val(Trim(grdCliente.TextMatrix(iTemp, 4))) <= 1 And Trim(grdCliente.TextMatrix(iTemp, 1)) <> "" Then
            If oPersonaTemp.validarPersonaMayorEdad(grdCliente.TextMatrix(iTemp, 1), Format(gdFecSis, "dd/mm/yyyy")) = False Then
                nMenorEdad = nMenorEdad + 1
            End If
        End If
    Next
    
    If nMenorEdad > 0 Then
        If intPunteroPJ_NA > 0 Then
            MsgBox "No es posible agregar menores de edad en cuentas con personería jurídica", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    
    Set oPersonaTemp = Nothing
    
    ' ** Fin Validacion de Mayoria de edad
        
    For J = 1 To Me.grdCliente.rows - 1
        If Trim(Me.grdCliente.TextMatrix(J, 6)) = "" Then
            MsgBox "Debe asignar un grupo a cada cliente", vbExclamation, "Mensaje"
            Exit Sub
        End If
    Next
    
    If validaExistenciaReglas = False And nProducto <> gCapCTS Then
        MsgBox "Verificar si los grupos asignados forman parte de alguna regla o si las reglas contienen los grupos asignados", vbExclamation, "Mensaje"
        Exit Sub
    End If
    nLetraMax = 65
    
    'Obteniendo la letra mayor
    
    For J = 1 To grdCliente.rows - 1
        If val(Trim(grdCliente.TextMatrix(J, 4))) = gPersonaNat And Trim(grdCliente.TextMatrix(J, 6)) <> "AP" Then
            If CInt(AscW(grdCliente.TextMatrix(J, 6))) > nLetraMax Then
                nLetraMax = CInt(AscW(grdCliente.TextMatrix(J, 6)))
            End If
        End If
    Next
    J = 1
    For nLetra = 65 To nLetraMax
        nContar = 0
        For J = 1 To grdCliente.rows - 1
            If Chr(nLetra) = grdCliente.TextMatrix(J, 6) Then
                nContar = nContar + 1
            End If
        Next
        If nContar = 0 Then
            MsgBox "La secuencia de los grupos deben ser: A, B, C, D ...", _
            vbExclamation, "Aviso"
            Exit Sub
        End If
    Next
    seleccionarTipoCuentaXregla
    If nTipoCuentaInicial <> cboInstitucion.ListIndex Then
        If intPunteroPJ_NA = 0 Then
            MsgBox "No esposible cambiar el tipo de cuenta", vbInformation, "Aviso"
            Exit Sub
        End If
    End If

' *** END RIRO


'JUEZ 20130724 ********************************************************
If nProducto = gCapCTS Then
    Dim lsPersCodTitular As String, n As Integer
    Dim oCap As New COMNCaptaGenerales.NCOMCaptaGenerales
    For n = 1 To grdCliente.rows - 1
        If Trim(Right(grdCliente.TextMatrix(n, 3), 2)) = "10" Then
            lsPersCodTitular = grdCliente.TextMatrix(n, 1)
            Exit For
        End If
    Next n
    If oCap.VerificarExisteCuentaCTS(lsPersCodTitular, txtInstitucion.Text, IIf(lblMoneda.Caption = "Nacional", 1, 2), txtCuenta.NroCuenta) Then
        MsgBox "No es posible realizar la operación debido a que ya existe otra cuenta CTS del cliente con el mismo empleador y con la misma moneda", vbInformation, "Aviso"
        Exit Sub
    End If
    Set oCap = Nothing
End If
'END JUEZ *************************************************************

If MsgBox("¿Está seguro de grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim oDCapGen As COMDCaptaGenerales.DCOMCaptaGenerales 'JUEZ 20151229
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
    Dim clsMov As COMNContabilidad.NCOMContFunciones 'NContFunciones
    Dim rsRel As New ADODB.Recordset
    Dim clsMnto As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim oNCapDef As COMNCaptaGenerales.NCOMCaptaDefinicion 'JUEZ 20151229
    Dim nComiMantPoderes As Double 'JUEZ 20151229
    Dim nMovNro As Long 'JUEZ 20151229
    
    'JUEZ 20151229 ***********************************
    If fbMantPoderes Then
        If MsgBox("Esta operación generará una pendiente de cobro por la comisión de mantenimiento de poderes, ¿Desea continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    End If
    'END JUEZ ****************************************
    
    Set oNCapDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
        nComiMantPoderes = oNCapDef.GetCapParametro(2158)
    Set oNCapDef = Nothing
    
    Set clsMov = New COMNContabilidad.NCOMContFunciones
        sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set clsMov = Nothing
          
    Set rsRel = grdCliente.GetRsNew
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set clsMnto = New COMNCaptaGenerales.NCOMCaptaMovimiento
    
    Select Case nProducto
        Case gCapAhorros
            ' RIRO 20131102, SE AGREGO PARAMETROS "sGrupos, sReglas"
            If clsMant.ActualizaDatosCuenta(txtCuenta.NroCuenta, rsRel, nFirmas, nTipoCuenta, , bOrdPag, , val(txtMinFirmas.Text), sMovNro, , Trim(txtAlias.Text), , , , prepararGrupoPersona(), prepararRegla()) Then
                clsMnto.ActualizaCodigoPersConvenioAho txtCuenta.NroCuenta, lsCodPersConv
                'JUEZ 20151229 *******************************************
                If fbMantPoderes Then
                    Set oDCapGen = New COMDCaptaGenerales.DCOMCaptaGenerales
                    nMovNro = oDCapGen.GetnMovNro(sMovNro)
                    Set oDCapGen = Nothing
                    clsMant.RegistrarPendComisionMantPoderes txtCuenta.NroCuenta, nMovNro, sMovNro, nComiMantPoderes
                    MsgBox "Se ha registrado la pendiente del cobro de la comisión por mantenimiento de poderes", vbInformation, "Aviso"
                End If
                'END JUEZ ************************************************
            End If
        Case gCapPlazoFijo
            ' RIRO 20131102, SE AGREGO PARAMETROS "sGrupos, sReglas"
            'MIOL 20121109, SEGUN RFC098-2012-B ************************************************
            If clsMant.ActualizaDatosCuenta(txtCuenta.NroCuenta, rsRel, nFirmas, nTipoCuenta, , , sCuentaAboInt, val(txtMinFirmas.Text), sMovNro, , Trim(txtAlias.Text), , , IIf(chkSubasta.Visible And chkSubasta.value = vbChecked, 1, 0), prepararGrupoPersona(), prepararRegla()) Then
            'If clsMant.ActualizaDatosCuenta(txtCuenta.NroCuenta, rsRel, nFirmas, nTipoCuenta, , , sCuentaAboInt, val(txtMinFirmas.Text), sMovNro, , Trim(txtAlias.Text)) Then
            'END MIOL **************************************************************************
            End If
        Case gCapCTS
            'By Capi 01042008 para controlar el porcentaje de retiro adicional
            If nPorRetAdiCTS + nPorRetCTS > 100 Then
                MsgBox "El porcentaje de Retiro Adicional( " & nPorRetAdiCTS & "%)+ establecido por ley(" & nPorRetCTS & "%) supera el saldo total CTS(100%), verifique parametros", vbInformation, "Aviso"
                Exit Sub
            End If
            
            ' RIRO 20131102, SE AGREGO PARAMETROS "sGrupos, sReglas"
            If clsMant.ActualizaDatosCuenta(txtCuenta.NroCuenta, rsRel, nFirmas, nTipoCuenta, sInstitucion, , , val(txtMinFirmas.Text), sMovNro, , Trim(txtRetiroAdicional.Text), , nPorRetAdiCTS, , prepararGrupoPersona(), prepararRegla()) Then 'By Capi Acta 014-2007
            End If
    End Select
    'By capi 21012009
     objPista.InsertarPista gsOpeCod, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gModificar, , txtCuenta.NroCuenta, gCodigoCuenta
    '
    'JUEZ 20130527 ***********************************************************************
    If fbRegistraEnvio = True Then
        Call frmEnvioEstadoCta.GuardarRegistroEnvioEstadoCta(1, txtCuenta.NroCuenta, frsEnvEstCta, fnModoEnvioEstCta, fnDebitoMismaCta, sMovNro)
    End If
    'END JUEZ ****************************************************************************
    Set clsMant = Nothing
    cmdCancelar_Click
End If
End Sub

Private Sub cmdImprimir_Click()
Dim clsImp As COMNCaptaGenerales.NCOMCaptaImpresion
Dim clsPrev As previo.clsprevio
Dim rsPers As New ADODB.Recordset
Dim sCuenta As String
Dim sCad As String
Dim sProducto As String
Dim sFecVenc As String
Dim lsCtaAbo As String

    Set clsImp = New COMNCaptaGenerales.NCOMCaptaImpresion
    Set rsPers = grdCliente.GetRsNew
    sCuenta = txtCuenta.NroCuenta
    If lblVencimiento.Caption <> "" Then
        sFecVenc = CDate(lblVencimiento)
    End If
    
    If Mid(sCuenta, 6, 3) = "232" Then
        sProducto = "AHORROS"
    ElseIf Mid(sCuenta, 6, 3) = "233" Then
        sProducto = "PLAZO FIJO"
    Else
        sProducto = "CTS"
    End If
    
    lsCtaAbo = txtCtaAhoAboInt.Text
    sCad = clsImp.ImprimeDetalleCuenta(sCuenta, sProducto, lblPrograma, CDate(lblApertura), lblTipoCuenta, lblPlazo, lblTipoTasa, lblEstado, lblFormaRetiro, _
                                       CCur(lblTasa), lblMoneda, sFecVenc, chkOrdenPago.value, chkAbonoIntCtaAho.value, _
                                       rsPers, gsNomAge, gdFecSis, gsCodUser, 0, lsCtaAbo)
    Set clsPrev = New previo.clsprevio
    clsPrev.Show sCad, "Registro de Firmas", True, , gImpresora
    Set clsPrev = Nothing
End Sub

'JUEZ 20150229 *********************
Private Sub cmdMantPoderes_Click()
    If validaExistenciaReglas Then fbMantPoderes = True
    HabilitarPoderes True
End Sub
'END JUEZ **************************

Private Sub cmdNuevaBus_Click()
Dim lsCargos() As String '***Modificado por ELRO el 20130110, según OYP-RFC123-2012
Dim J As Integer '***Modificado por ELRO el 20130110, según OYP-RFC123-2012

txtCuenta.Cuenta = ""
lblApertura = ""
lblEstado = ""
lblFormaRetiro = ""
lblMoneda = ""
lblPlazo = ""
lblTasa = ""
lblTipoCuenta = ""
lblTipoTasa = ""
cboInstitucion.Clear
txtNumFirmas.Text = ""
chkOrdenPago.value = 0
cmdEditar.Enabled = False
txtCuenta.Enabled = True
cmdBuscar.Enabled = True
txtCuenta.SetFocusCuenta
cmdRegFirmas.Enabled = False
grdCliente.Clear
grdCliente.rows = 2
grdCliente.FormaCabecera
nTitular = 0
txtCtaAhoAboInt.Text = ""
txtInstitucion.Text = ""
lblInstitucion.Caption = ""
txtMinFirmas.Text = ""
Me.chkSubasta.value = 0 'MIOL 20121110, SEGUN RFC098-2012-B
'***Modificado por AMDO el 20130726, según TI-ERS086-2013 ***********
'***Modificado por ELRO el 20130110, según OYP-RFC123-2012
'lsCargos = Split(fsCargoCodModAgeManProPas, ",")
'For J = 0 To UBound(lsCargos)
'    If lsCargos(J) = gsCodCargo Then
'        txtCuenta.Age = ""
'        Exit For
'    End If
'Next J
''***Fin Modificado por ELRO el 20130110*******************
If bPermisoCargo Then
        txtCuenta.Age = ""
End If
'***END AMDO *******************************************************
'JUEZ 20141202 ********************
lblPrograma = ""
lblVencimiento.Caption = ""
If chkCovDep.Visible Then
    chkCovDep.Visible = False
    chkCovDep.value = 0
    cboInstConvDep.Visible = False
End If
'END JUEZ *************************
fbMantPoderes = False 'JUEZ 20151229
End Sub

' *** RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
Private Sub cmdQuitarRega_Click()

    If grdReglas.rows = 2 And Trim(grdReglas.TextMatrix(1, 1)) = "" Then
        Exit Sub
    End If
    If MsgBox("¿¿Está seguro de eliminar la regla creada??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        grdReglas.EliminaFila grdReglas.row
    End If
    seleccionarTipoCuentaXregla

End Sub
' *** END RIRO

Private Sub cmdRegFirmas_Click()
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Dim clsPrev As previo.clsprevio
Dim rsPers As New ADODB.Recordset
'FRHU 20140927 ERS099-2014
Dim rsPF As New ADODB.Recordset
Dim nMontoPF As Double, nTasaNomninalPF As Double
Dim nFormaRetiroPF As Integer, nPlazoPF As Integer
Dim dAperturaPF As Date
'FIN FRHU 20140927
Dim sCuenta As String, sCad As String
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
Set rsPers = grdCliente.GetRsNew
'***Modificado por ELRO el 20130725, según SATI INC1306130011
'sCad = clsMant.GeneraRegistroFirmas(txtCuenta.NroCuenta, lblTipoCuenta, CDate(lblApertura), chkOrdenPago.value, rsPers, gsNomAge, gdFecSis, gsCodUser)
'FRHU 20140927 ERS099-2014
'sCad = clsMant.GeneraRegistroFirmas(txtCuenta.NroCuenta, lblTipoCuenta, CDate(lblApertura), chkOrdenPago.value, rsPers, gsNomAge, gdFecSis, gsCodUser, fnTpoPrograma)
If nPersoneria <> 1 Then
    sCad = clsMant.GeneraSolicitudAperturaPersJuridica(txtCuenta.NroCuenta, lblTipoCuenta, CDate(lblApertura), chkOrdenPago.value, rsPers, gsNomAge, gdFecSis, gsCodUser, fnTpoPrograma, , prepararRegla())
Else
    sCad = clsMant.GeneraSolicitudAperturaPersNatural(txtCuenta.NroCuenta, lblTipoCuenta, CDate(lblApertura), chkOrdenPago.value, rsPers, gsNomAge, gdFecSis, gsCodUser, fnTpoPrograma, , prepararRegla())
End If
'FIN FRHU 20140927
'***Modificado por ELRO el 20130725, según SATI INC1306130011
Set clsPrev = New previo.clsprevio
'clsPrev.Show sCad, "Registro de Firmas", True, , gImpresora 'FRHU 20140927 ERS099-2014
clsPrev.Show sCad, "Solicitud de Apertura", True, , gImpresora
Set clsPrev = Nothing
'FRHU 20140927 ERS099-2014: SE QUITO LA LIBRETA DE AHORROS
''By Capi Acta 014-2007 Impresion Libretas de Ahorro
'            Dim lsCad As String
'            Dim clsLibretM As previo.clsprevio
'            Set clsLibretM = New previo.clsprevio
'
'            MsgBox "Coloque Libreta para Impresion", vbInformation, "Aviso"
'            Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
'            clsMant.IniciaImpresora gImpresora
'            lsCad = clsMant.GeneraRegistroLibretas(txtCuenta.NroCuenta, lblTipoCuenta, CDate(lblApertura), chkOrdenPago.value, rsPers, gsNomAge, gdFecSis, gsCodUser)
'            clsLibretM.Show lsCad, "Registro Libretas", True, , gImpresora
'            Set clsMant = Nothing
'            Set clsPrev = Nothing
'        'End By
If nProducto = gCapPlazoFijo Then
    Set rsPF = clsMant.GetDatosParaCronogramaPlazoFijo(txtCuenta.NroCuenta)
    If Not rsPF.BOF And Not rsPF.EOF Then
        nMontoPF = rsPF!nSaldo
        nPlazoPF = rsPF!nPlazo
        nFormaRetiroPF = rsPF!nFormaRetiro
        nTasaNomninalPF = rsPF!nTasaInteres
        dAperturaPF = rsPF!dApertura
        If nFormaRetiroPF = gCapPFFormRetMensual Then
            'EmiteCalendarioRetiroIntPFMensual IIf(Me.chkITFEfectivo = 1, nMonto, nMonto - CDbl(Me.lblITF.Caption)), nTasa, CInt(txtPlazo), gdFecSis, nMoneda, lnDValoriza, sCuenta
            EmiteCalendarioRetiroIntPFMensual nMontoPF, nTasaNomninalPF, nPlazoPF, dAperturaPF, Mid(txtCuenta.NroCuenta, 9, 1), 0, txtCuenta.NroCuenta
        End If
    End If
    Set rsPF = Nothing
End If
'FIN FRHU 20140927
Set clsMant = Nothing
rsPers.Close
End Sub
'***Agregado por ELRO el 20130111, según OYP-RFC123-2012
Private Sub cmdRenovaciones_Click()
    Dim oform As New frmCapHistoricoRenovacionesDPF
    oform.iniciarHistoricoRenovacionesDPF (txtCuenta.NroCuenta)
End Sub
'***Fin Agregado por ELRO el 20130111*******************
Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdTasaHist_Click()
    frmCapCambioTasaHist.Inicio txtCuenta.NroCuenta
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And txtCuenta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(nProducto, False)
        If sCuenta <> "" Then
            txtCuenta.NroCuenta = sCuenta
            txtCuenta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()

Dim i As Integer
nTitular = 0
nClientes = 0
nRepresentante = 0
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
'By Capi 20012009
Set objPista = New COMManejador.Pista
'
'JUEZ 20130527 *************
fbRegistraEnvio = False
Set frsEnvEstCta = Nothing
fnModoEnvioEstCta = 0
fnDebitoMismaCta = 0
'END JUEZ ******************

' *** AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"

fraReglasPorderes.Left = 6550
fraReglasPorderes.top = 6720
fraReglasPorderes.Height = 2055
fraReglasPorderes.Width = 2625
fraReglasPorderes.Visible = True
'fraReglasPorderes.Enabled = False 'RIRO20140710 ERS098
HabilitarPoderes False 'RIRO20140710 ERS098
nTipoCuentaInicial = -1
For i = 65 To 90 Step 1
   lsLetras.AddItem UCase(Chr(i))
Next i

' *** END RIRO

End Sub

Private Sub grdCliente_DblClick()
Dim nCol As Long
Dim nFila As Long
nCol = grdCliente.Col
nFila = grdCliente.row
If nCol <> 3 Then
    Dim R As ADODB.Recordset
    Dim ssql As String
    Dim clsFirma As COMDCaptaGenerales.DCOMCaptaMovimiento 'DCapMovimientos
    Set clsFirma = New COMDCaptaGenerales.DCOMCaptaMovimiento
     
    If Me.grdCliente.TextMatrix(nFila, 1) = "" Then Exit Sub
    
    Set R = New ADODB.Recordset
    Set R = clsFirma.GetFirma(Me.grdCliente.TextMatrix(nFila, 1))
    frmMuestraFirma.psCodCli = Me.grdCliente.TextMatrix(nFila, 1)

        
    If R.RecordCount > 0 Then
        If IsNull(R!iPersFirma) = True Then
            MsgBox "El cliente no posee Firmas", vbInformation, "Aviso"
            Exit Sub
        End If
        
        'frmMuestraFirma.Show 1 COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
        
    End If
    
    
Else

    ' *** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
    
        'If grdCliente.TextMatrix(nFila, 3) <> "" Then
        '    If CLng(Right(grdCliente.TextMatrix(nFila, 3), 2)) = gCapRelPersTitular And grdCliente.TextMatrix(nFila, 4) <> "N" Then
        '        grdCliente.lbEditarFlex = False
        '    Else
        '        grdCliente.lbEditarFlex = True
        '    End If
        'Else
        '    grdCliente.lbEditarFlex = True
        'End If
        
    ' *** END RIRO
End If
End Sub

Private Sub grdCliente_KeyPress(KeyAscii As Integer)

' *** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"

        'If grdCliente.Col = 5 Then
        '    If UCase(Chr(KeyAscii)) = "S" Or UCase(Chr(KeyAscii)) = "SI" Then
        '        grdCliente.TextMatrix(grdCliente.row, 5) = "SI"
        '    ElseIf UCase(Chr(KeyAscii)) = "N" Or UCase(Chr(KeyAscii)) = "NO" Then
        '        grdCliente.TextMatrix(grdCliente.row, 5) = "NO"
        '    ElseIf UCase(Chr(KeyAscii)) = "O" Or UCase(Chr(KeyAscii)) = "OPCIONAL" Then
        '        grdCliente.TextMatrix(grdCliente.row, 5) = "OPCIONAL"
        '    Else
        '        MsgBox "PRESIONE LA TECLA S SI LA FIRMA DE ESTE CLIENTE ES OBLIGATORIA." & vbCrLf & "PRESIONE LA TECLA N SI LA FIRMA DE ESTE CLIENTE NO ES OBLIGATORIA." & vbCrLf & "PRESIONE LA TECLA O SI LA FIRMA DE ESTE CLIENTE  ES OPCIONAL.", vbOKOnly + vbExclamation, App.Title
        '    End If
        'End If

' *** END RIRO

   'Columna que contiene los grupos: A, B, C, D ...
    If grdCliente.Col = 6 Then
                
        If intPunteroPJ_NA = 0 Then Exit Sub
        
        Dim i, nContar, nContarGrupos As Integer
        Dim sLetra As String
        Dim sReglas() As String
        Dim sD As Variant
        
        ' Bloqueando columnas no editables
        Dim sColumnas() As String
        sColumnas = Split(grdCliente.ColumnasAEditar, "-")
        If sColumnas(grdCliente.Col) = "X" Or val(Trim(Right(grdCliente.TextMatrix(grdCliente.row, 3), 2))) = 11 Then
            Exit Sub
        End If
        ' Fin Bloqueo
        
        If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Then
            grdCliente.TextMatrix(grdCliente.row, 6) = UCase(Chr(KeyAscii))
        Else
            grdCliente.TextMatrix(grdCliente.row, 6) = ""
        End If
        
        If val(Trim(grdCliente.TextMatrix(grdCliente.row, 4))) > 1 Then
            grdCliente.TextMatrix(grdCliente.row, 6) = "PJ"
        End If
        seleccionarTipoCuentaXregla
    End If

End Sub

Private Sub grdCliente_LostFocus()

  'Se debe considerar que en la apertura, debe intervenir solo una persona juridica
    Dim nRelacion, nPJuridica, i, J As Integer
    Dim sMensaje As String
    
    For i = 1 To grdCliente.rows - 1
        If val(Trim(grdCliente.TextMatrix(i, 4))) > 1 Then
            If Trim(Left(grdCliente.TextMatrix(i, 3), 10)) <> "TITULAR" Then
                grdCliente.TextMatrix(i, 3) = ""
                nPJuridica = nPJuridica + 1
            End If
            For J = 1 To grdCliente.rows - 1
                If Trim(Left(grdCliente.TextMatrix(J, 3), 10)) = "TITULAR" And val(Trim(grdCliente.TextMatrix(J, 4))) = 1 Then
                    grdCliente.TextMatrix(J, 3) = ""
                    nRelacion = nRelacion + 1
                End If
            Next
            If nPJuridica > 0 Then
                sMensaje = "* La relacion de una persona jurídica debe ser: Titular" & vbNewLine
            End If
            If nRelacion > 0 Then
                sMensaje = sMensaje & "* Solo una persona jurídica debe ser titular de la cuenta"
            End If
            If nPJuridica > 0 Or nRelacion > 0 Then
                MsgBox "Observaciones: " & vbNewLine & vbNewLine & sMensaje, vbExclamation, "Aviso"
                Unload frmBuscaPersona
                grdCliente.SetFocus
                Exit Sub
            End If
        End If
    Next

' *** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"

        'Dim i As Integer, numfirmas As Integer
        '  numfirmas = 0
        'For i = 1 To grdCliente.Rows - 1
        '    If grdCliente.TextMatrix(i, 5) = "SI" Then
        '        numfirmas = numfirmas + 1
        '    End If
        'Next i
        'Label4.Tag = numfirmas
        'txtMinFirmas.Text = CStr(numfirmas)
        
' *** END RIRO
End Sub

Private Sub grdCliente_OnCellChange(pnRow As Long, pnCol As Long)
If pnCol = 5 Then
'    If UCase(Chr(KeyAscii)) = "S" Or UCase(Chr(KeyAscii)) = "SI" Then
'        grdCliente.TextMatrix(grdCliente.Row, 6) = "SI"
'    ElseIf UCase(Chr(KeyAscii)) = "N" Or UCase(Chr(KeyAscii)) = "NO" Then
'        grdCliente.TextMatrix(grdCliente.Row, 6) = "NO"
'    ElseIf UCase(Chr(KeyAscii)) = "O" Or UCase(Chr(KeyAscii)) = "OPCIONAL" Then
'        grdCliente.TextMatrix(grdCliente.Row, 6) = "OPCIONAL"
'    Else
'        MsgBox "PRESIONE LA TECLA S SI LA FIRMA DE ESTE CLIENTE ES OBLIGATORIA." & vbCrLf & "PRESIONE LA TECLA N SI LA FIRMA DE ESTE CLIENTE NO ES OBLIGATORIA." & vbCrLf & "PRESIONE LA TECLA O SI LA FIRMA DE ESTE CLIENTE  ES OPCIONAL.", vbOKOnly + vbExclamation, App.Title
'    End If

' *** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"

    'If UCase(Trim(grdCliente.TextMatrix(grdCliente.row, 5))) = "S" Or UCase(Trim(grdCliente.TextMatrix(grdCliente.row, 5))) = "SI" Then
    '    grdCliente.TextMatrix(grdCliente.row, 5) = "SI"
    'ElseIf UCase(Trim(grdCliente.TextMatrix(grdCliente.row, 5))) = "N" Or UCase(Trim(grdCliente.TextMatrix(grdCliente.row, 5))) = "NO" Then
    '    grdCliente.TextMatrix(grdCliente.row, 5) = "NO"
    'ElseIf UCase(Trim(grdCliente.TextMatrix(grdCliente.row, 5))) = "O" Or UCase(Trim(grdCliente.TextMatrix(grdCliente.row, 5))) = "OPCIONAL" Then
    '    grdCliente.TextMatrix(grdCliente.row, 5) = "OPCIONAL"
    'Else
    '    MsgBox "PRESIONE LA TECLA S SI LA FIRMA DE ESTE CLIENTE ES OBLIGATORIA." & vbCrLf & "PRESIONE LA TECLA N SI LA FIRMA DE ESTE CLIENTE NO ES OBLIGATORIA." & vbCrLf & "PRESIONE LA TECLA O SI LA FIRMA DE ESTE CLIENTE  ES OPCIONAL.", vbOKOnly + vbExclamation, App.Title
    'End If
    
' *** END RIRO

ElseIf pnCol = 6 Then
    grdCliente.TextMatrix(pnRow, 6) = UCase(grdCliente.TextMatrix(pnRow, 6))
    If val(Trim(grdCliente.TextMatrix(pnRow, 4))) > 1 Then
        grdCliente.TextMatrix(pnRow, 6) = "PJ"
    Else
        If Len(grdCliente.TextMatrix(pnRow, 6)) > 1 And Trim(grdCliente.TextMatrix(pnRow, 6)) <> "AP" Then
            grdCliente.TextMatrix(pnRow, 6) = ""
        ElseIf Len(grdCliente.TextMatrix(pnRow, 6)) = 1 Then
            Dim nAs As Integer
            nAs = CInt(AscW(grdCliente.TextMatrix(grdCliente.row, 6)))
            If (nAs >= 65 And nAs <= 90) Or (nAs >= 97 And nAs <= 122) Then
                grdCliente.TextMatrix(grdCliente.row, 6) = UCase(Chr(nAs))
                grdCliente.row = IIf(pnRow + 1 = grdCliente.rows, pnRow, pnRow + 1)
                grdCliente.Col = 3
                grdCliente.SetFocus
            Else
                grdCliente.TextMatrix(grdCliente.row, 6) = ""
            End If
        End If
    End If
    seleccionarTipoCuentaXregla
End If

CuentaTitular
'EvaluaTitular


End Sub

Private Sub grdCliente_OnChangeCombo()

    If grdCliente.Col = 3 Then
        If val(Trim(Right(grdCliente.TextMatrix(grdCliente.row, 3), 2))) = 11 And _
           val(Trim(grdCliente.TextMatrix(grdCliente.row, 4))) <= 1 Then
           grdCliente.TextMatrix(grdCliente.row, 6) = "AP"
        ElseIf val(Trim(grdCliente.TextMatrix(grdCliente.row, 4))) <= 1 Then
            grdCliente.TextMatrix(grdCliente.row, 6) = "A"
        End If
    End If

    ' *** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
    
        'If grdCliente.Col = 3 Then
        '    If Right(grdCliente.TextMatrix(grdCliente.row, 3), 2) = "10" Then
        '        grdCliente.TextMatrix(grdCliente.row, 5) = "SI"
        '    Else
        '        grdCliente.TextMatrix(grdCliente.row, 5) = "NO"
        '    End If
        'End If
    
    ' *** END RIRO
    
End Sub

Private Sub grdCliente_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
Dim nIndex As Integer
nIndex = grdCliente.rows - 1

' *** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"

        'If pbEsDuplicado Then
        '    MsgBox "Persona ya esta registrada en la relación.", vbInformation, "Aviso"
        '    grdCliente.EliminaFila nIndex
        'End If
        'If psDataCod = "" Then
        '    grdCliente.EliminaFila nIndex
        'End If

' *** END RIRO

If pbEsDuplicado And psDataCod <> "" Then
    MsgBox "Persona ya esta registrada en la relación.", vbInformation, "Aviso"
    grdCliente.EliminaFila nIndex

ElseIf psDataCod = "" Then
    If Not pbEsDuplicado Then
        grdCliente.TextMatrix(pnRow, 3) = ""
        grdCliente.TextMatrix(pnRow, 4) = 0
        grdCliente.TextMatrix(pnRow, 6) = ""
    End If

ElseIf psDataCod = gsCodPersUser Then
    MsgBox "No se puede agregar asi mismo.", vbInformation, "Aviso"
    grdCliente.EliminaFila grdCliente.row
    
Else
    grdCliente.TextMatrix(pnRow, 4) = grdCliente.PersPersoneria
    
    'Validando que en la cuenta con personas juridicas solo intervenga un titular.
    Dim nContar, nContarNA, i As Integer
    For i = 1 To grdCliente.rows - 1
        If grdCliente.TextMatrix(i, 1) <> "" Then
            If val(Trim(grdCliente.TextMatrix(i, 4))) > 1 Then
                nContar = nContar + 1
            End If
        End If
    Next
    If nContar > 1 Then
        MsgBox "No es posible relacionar dos personas jurídicas en una misma cuenta.", vbExclamation, "Aviso"
        grdCliente.EliminaFila grdCliente.row
        Exit Sub
    ElseIf nContar = 1 Then
    'Validando que en una cuenta aperturada con personas naturales, no se pueda ingresar
    'una persona juridica.

        If intPunteroPJ_NA = 0 Then
            MsgBox "No es posible relacionar una persona jurídica en una cuenta aperturada con Personas naturales.", vbExclamation, "Aviso"
            grdCliente.EliminaFila grdCliente.row
        End If
    
    End If
    seleccionarTipoCuentaXregla
End If

End Sub

Private Sub grdCliente_OnRowChange(pnRow As Long, pnCol As Long)

' *** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"

        'Dim i As Integer, numfirmas As Integer
        '  numfirmas = 0
        'For i = 1 To grdCliente.Rows - 1
        '    If grdCliente.TextMatrix(i, 5) = "SI" Then
        '        numfirmas = numfirmas + 1
        '    End If
        'Next i
        'Label4.Tag = numfirmas
        'txtMinFirmas.Text = CStr(numfirmas)

' *** END RIRO

    Dim row, Col As Integer
    row = grdCliente.row
    Col = grdCliente.Col
    If intPunteroPJ_NA = 0 Then
        If Not bActualizarCuenta Then
            Exit Sub
        End If
    End If
    If (val(Me.grdCliente.TextMatrix(row, 7)) = 10 And (Col = 1 Or Col = 2 Or Col = 3)) Or _
    (Col = 6 And val(Trim(Right(Me.grdCliente.TextMatrix(row, 3), 2))) = 11) Then
        grdCliente.lbEditarFlex = False
    Else
        grdCliente.lbEditarFlex = True
    End If

End Sub

Private Sub grdCliente_OnRowDelete()
    limpiarReglas
End Sub

Private Sub grdCliente_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)

' *** AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"

    Dim sColumnas() As String
    sColumnas = Split(grdCliente.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
    
    If pnCol = 6 Then

        grdCliente.TextMatrix(pnRow, 6) = UCase(grdCliente.TextMatrix(pnRow, 6))
        If val(Trim(grdCliente.TextMatrix(pnRow, 4))) > 1 Then
            grdCliente.TextMatrix(pnRow, 6) = "PJ"
        Else
            If Len(grdCliente.TextMatrix(pnRow, 6)) > 1 Then
            grdCliente.TextMatrix(pnRow, 6) = ""
            End If
        End If
        
    End If

' *** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"

        'If pnCol > 4 And grdCliente.TextMatrix(pnRow, 5) = "" Then
        ' MsgBox "DEBE INDICAR SI ES OBLIGATORIA O NO LA FIRMA DEL CLIENTE", vbOKOnly + vbInformation, App.Title
        'End If

' *** END RIRO

End Sub

'Agregado Por RIRO el 20130501, Proyecto Ahorro - Poderes
Private Sub grdCliente_RowColChange()
    Dim row, Col As Integer
    row = grdCliente.row
    Col = grdCliente.Col
    If intPunteroPJ_NA = 0 Then
        If Not bActualizarCuenta Then
            Exit Sub
        End If
    End If
    If (val(Me.grdCliente.TextMatrix(row, 7)) = 10 And (Col = 1 Or Col = 2 Or Col = 3)) Or _
    (Col = 6 And val(Trim(Right(Me.grdCliente.TextMatrix(row, 3), 2))) = 11) Then
        grdCliente.lbEditarFlex = False
    Else
        grdCliente.lbEditarFlex = True
    End If
End Sub

Private Sub TXTALIAS_KeyPress(KeyAscii As Integer)
    KeyAscii = fgIntfMayusculas(KeyAscii)
End Sub

Private Sub txtCtaAhoAboInt_EmiteDatos()
Dim nFila As Long
Dim sCodigo As String
Dim nMoneda As Moneda

'    If (gsCodCargo = "002001" Or gsCodCargo = "002002" Or gsCodCargo = "002003" _
'                  Or gsCodCargo = "003001" Or gsCodCargo = "003002" Or gsCodCargo = "004001" _
'                  Or gsCodCargo = "004002" Or gsCodCargo = "006005") And bPosCli = False Then
        nFila = grdCliente.row
        sCodigo = grdCliente.TextMatrix(nFila, 1)
        nMoneda = CLng(Mid(txtCuenta.NroCuenta, 9, 1))
        If nProducto = gCapPlazoFijo Then
            Dim rsCta As New ADODB.Recordset
            Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
            Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
                 Set rsCta = clsMant.GetCuentasPersona(sCodigo, gCapAhorros, True, , nMoneda)
            Set clsMant = Nothing
            txtCtaAhoAboInt.rs = rsCta
            If rsCta.EOF And rsCta.BOF Then
                txtCtaAhoAboInt.Text = ""
                MsgBox "Cliente NO Posee cuentas de ahorros disponibles", vbInformation, "Aviso"
            End If
            Set rsCta = Nothing
            txtCtaAhoAboInt.Visible = True
        Else
            
        End If
'    End If
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sCta As String
        sCta = txtCuenta.NroCuenta
        ObtieneDatosCuenta sCta
        'fraReglasPorderes.Enabled = False ' *** RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
        'JUEZ 20151229 ********************
        If intPunteroPJ_NA = 0 Then
            cmdMantPoderes.Visible = False
        Else
            cmdMantPoderes.Visible = True
        End If
        'END JUEZ *************************
        HabilitarPoderes False 'RIRO20140710 ERS098
        'By Capi Acta 014-2007
        'By Capi 23052008 para que active el frame solo para CTS
        If txtCuenta.Prod = gCapCTS Then
            frmCTS.Enabled = True
            frmCTS.Visible = True
            
            If (nRetAdiCTS Mod 2) = 0 Then
                chkRetiroAdicional.Visible = True
                lblRetiroAdicional.Visible = True
                txtRetiroAdicional.Visible = True
                txtAlias.Visible = False
                Label9.Visible = False
            End If
        End If
        
    End If
End Sub

Private Sub txtInstitucion_EmiteDatos()
lblInstitucion = txtInstitucion.psDescripcion
End Sub

Private Sub TxtMinFirmas_KeyPress(KeyAscii As Integer)
Dim nFirmas As Integer
nFirmas = grdCliente.rows - 1
If txtCuenta.Prod = "234" Then
        KeyAscii = 0
        Exit Sub
End If
If KeyAscii <> 13 And Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
    KeyAscii = 0
End If
If KeyAscii = 13 Then
    If txtMinFirmas = nFirmas Then
        MsgBox "El número de Firmas no puede ser mayor", vbInformation, "Aviso"
        Exit Sub
    End If
End If
End Sub

Private Sub txtMinFirmas_LostFocus()
If val(txtMinFirmas.Text) < val(Label4.Tag) Then
   MsgBox "EL NRO MINIMO DE FIRMAS OBLIGATORIAS DEBEN SER " & CStr(Label4.Tag), vbOKOnly + vbInformation, "AVISO"
   txtMinFirmas.Text = CStr(Label4.Tag)
End If
End Sub

Private Sub txtNumFirmas_GotFocus()
With txtNumFirmas
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtNumFirmas_KeyPress(KeyAscii As Integer)
Dim nFirmas As Integer
nFirmas = grdCliente.rows - 1
If KeyAscii = 13 Then
    If txtNumFirmas = nFirmas Then
        MsgBox "El número de Firmas no puede ser mayor", vbInformation, "Aviso"
        Exit Sub
    End If
    If cboInstitucion.Visible Then
        cboInstitucion.SetFocus
    Else
        grdCliente.SetFocus
    End If
Else
    KeyAscii = NumerosEnteros(KeyAscii)
End If
End Sub

Public Sub MuestraPosicionCliente(ByVal psCodCta As String, Optional ByVal bImprime As Boolean, Optional ByVal bRegistroFirmas As Boolean = False)
    nProducto = Mid(psCodCta, 6, 3)
    txtCuenta.NroCuenta = psCodCta
    IniciaComboConvDep 9
    ObtieneDatosCuenta psCodCta, True
    fraCuenta.Enabled = False
    fraDatosMod.Enabled = False
    fraCliente.Enabled = False
    cmdEditar.Enabled = False
    cmdCancelar.Enabled = False
    cmdGrabar.Enabled = False
    
    ' *** RIRO20140710 ERS098
    Label10.Visible = False
    Label4.Visible = False
    txtNumFirmas.Visible = False
    txtMinFirmas.Visible = False
    ' *** END RIRO
    
    'cmdRegFirmas.Enabled = False Comentado Por RIRO SEGUN ERS: TI-ERS083-2013
    cmdRegFirmas.Enabled = bRegistroFirmas ' Comentado Por RIRO SEGUN ERS: TI-ERS083-2013
    fraCtaAhoAboInt.Enabled = False
    'Add By GITU 23-10-2012
    If lsCodPersConv <> "" Then
        chkCovDep.value = 1
        cboInstConvDep.ListIndex = IndiceListaCombo(Me.cboInstConvDep, lsCodPersConv)
    End If
    'End GITU
    If bImprime Then
        cmdImprimir.Visible = True
    End If
    'Add By GITU
    bPosCli = True
    Me.Show 1
End Sub
Private Sub IniciaComboConvDep(ByVal pnTipoRol As Integer)
Dim lRegPers As New ADODB.Recordset
Dim oPers As COMDPersona.DCOMRoles

    Set oPers = New COMDPersona.DCOMRoles
    Set lRegPers = oPers.CargaPersonas(pnTipoRol)
    Set oPers = Nothing

    If Not lRegPers.BOF And Not lRegPers.EOF Then
        Do While Not lRegPers.EOF
            cboInstConvDep.AddItem lRegPers("cPersNombre") & Space(100) & lRegPers("cPersCod")
            lRegPers.MoveNext
        Loop
        cboInstConvDep.ListIndex = 0
    End If
    lRegPers.Close
    Set lRegPers = Nothing
 End Sub
'***Agregado por ELRO el 20121114, según OYP-RFC101-2012
Private Sub cmdHistorico_Click()
    Dim oform As New frmCapHistoricoRemuneracionesCTS
    oform.iniciarHistoricoRemBruCTS (txtCuenta.NroCuenta)
End Sub
'***Fin Agregado por ELRO el 20121114*******************

'Comentado por JUEZ 20140305
'Private Sub cmdCtsNoActivo_Click()
'If MsgBox("¿Esta seguro que desea trasladar la cuenta " & txtCuenta.NroCuenta & " a CTS No Activo?", vbYesNo, "¡Aviso!") = vbYes Then
'    cmdCtsNoActivo.Enabled = False
'    Dim oNCOMCaptaMovimiento As New COMNCaptaGenerales.NCOMCaptaMovimiento
'    Dim oNCOMCaptaDefinicion As New COMNCaptaGenerales.NCOMCaptaDefinicion
'    Dim oNCOMContFunciones As New COMNContabilidad.NCOMContFunciones
'    Dim oNCOMCaptaGenerales As New COMNCaptaGenerales.NCOMCaptaGenerales
'
'    Dim rsCta As New ADODB.Recordset
'    Dim lnMovNro As Long
'    Dim lsMovNro As String
'    Dim lnTasaCTSNoActiva As Currency
'
'    Set rsCta = oNCOMCaptaGenerales.GetDatosCuenta(txtCuenta.NroCuenta)
'    lnTasaCTSNoActiva = oNCOMCaptaDefinicion.GetCapTasaInteres(gCapCTS, IIf(Mid(txtCuenta.NroCuenta, 9, 1) = "1", gMonedaNacional, gMonedaExtranjera), rsCta!nPrdTasaInteres, , rsCta!nSaldoDisp, gsCodAge, , 2)
'    lsMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
'    lnMovNro = oNCOMCaptaMovimiento.pasarCTSNoActiva(txtCuenta.NroCuenta, lsMovNro, gCTSPasarNoActiva, lnTasaCTSNoActiva)
'    If lnMovNro > 0 Then
'        cmdCtsNoActivo.Enabled = True
'        cmdCtsNoActivo.Visible = False
'        cmdCancelar_Click
'        ObtieneDatosCuenta txtCuenta.NroCuenta
'    End If
'End If
'End Sub
'***Fin Agregado por ELRO el 20121114*******************


' *** AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"

Private Function ValidaRelaciones() As Boolean
Dim i As Long
For i = 1 To grdCliente.rows - 1
    If Trim(grdCliente.TextMatrix(i, 3)) = "" Then
        MsgBox "Debe registrar la relación con el cliente.", vbInformation, "Aviso"
        grdCliente.row = i
        grdCliente.Col = 3
        grdCliente.SetFocus
        SendKeys "{Enter}"
        ValidaRelaciones = False
        Exit Function
    End If
Next i
ValidaRelaciones = True
End Function

Private Function validaExistenciaReglas() As Boolean

    Dim i, J, nContar, nContarAsociados As Integer
    Dim sReglas() As String
    Dim sLetra As Variant
    Dim lbValida As Boolean
    lbValida = True
    For i = 1 To grdCliente.rows - 1
        nContar = 0
        If Trim(grdCliente.TextMatrix(i, 6)) <> "AP" And Trim(grdCliente.TextMatrix(i, 6)) <> "PJ" Then
            nContarAsociados = nContarAsociados + 1
            For J = 1 To grdReglas.rows - 1
                sReglas = Split(grdReglas.TextMatrix(J, 1), "+")
                For Each sLetra In sReglas
                    If Trim(grdCliente.TextMatrix(i, 6)) = sLetra Then
                        nContar = nContar + 1
                    End If
                Next
            Next
            If nContar = 0 Then
                lbValida = False
            End If
        End If
    Next
    If nContarAsociados = 0 Then
        lbValida = False
    End If
    nContar = 0
    nContarAsociados = 0
    For i = 1 To grdReglas.rows - 1
        nContar = 0
        sReglas = Split(grdReglas.TextMatrix(i, 1), "+")
        For Each sLetra In sReglas
            nContar = 0
            For J = 1 To grdCliente.rows - 1
                If Trim(grdCliente.TextMatrix(J, 6)) <> "AP" And Trim(grdCliente.TextMatrix(J, 6)) <> "PJ" Then
                    nContarAsociados = nContarAsociados + 1
                    If sLetra = Trim(grdCliente.TextMatrix(J, 6)) Then
                        nContar = nContar + 1
                    End If
                End If
            Next
            If nContar = 0 Then
                lbValida = False
            End If
        Next
    Next
    If nContarAsociados = 0 Then
        lbValida = False
    End If
    
    validaExistenciaReglas = lbValida
     
End Function

Private Sub seleccionarTipoCuentaXregla()
    
    Dim nContar, nFirmantes, nTem, X, Y, i, J As Integer
    Dim sReglas(), sGruposPersonas() As String
    Dim sLetra, sValor As Variant
    Dim sGrupo As String
    Dim lbEsMancomunada, lbRepiteGrupo As Boolean
    lbEsMancomunada = True
    lbRepiteGrupo = True
    For i = 1 To grdCliente.rows - 1
        If Trim(grdCliente.TextMatrix(i, 1)) <> "" Then
            nTem = val(Trim(Right(grdCliente.TextMatrix(i, 3), 3)))
            If intPunteroPJ_NA = 0 Then
                If nTem = 10 Or nTem = 12 Then
                    nFirmantes = nFirmantes + 1
                    ReDim Preserve sGruposPersonas(J)
                    sGruposPersonas(J) = grdCliente.TextMatrix(i, 6)
                    J = J + 1
                End If
            Else
                If nTem = 12 Then
                    nFirmantes = nFirmantes + 1
                    ReDim Preserve sGruposPersonas(J)
                    sGruposPersonas(J) = grdCliente.TextMatrix(i, 6)
                    J = J + 1
                End If
            End If
        End If
    Next
    ' Solo intervienen personas naturales
    If nFirmantes <= 1 Then
        cboInstitucion.ListIndex = 0
        Exit Sub
    Else
        nTem = 0
        For Each sLetra In sGruposPersonas
            nTem = 0
            For Each sValor In sGruposPersonas
                If CStr(sLetra) = CStr(sValor) Then
                    nTem = nTem + 1
                End If
            Next
            If nTem <= 1 Then
                lbRepiteGrupo = False
            Else
                lbRepiteGrupo = True
                Exit For
            End If
        Next
        For i = 1 To grdReglas.rows - 1
            sGrupo = Trim(grdReglas.TextMatrix(i, 1))
            For Each sLetra In sGruposPersonas
                If InStr(sGrupo, CStr(sLetra)) = 0 Then
                    lbEsMancomunada = False
                    Exit For
                Else
                    lbEsMancomunada = True
                End If
            Next
            If Not lbEsMancomunada Then
                Exit For
            End If
        Next
        'Mancomunado
        If lbEsMancomunada And lbRepiteGrupo = False Then
            cboInstitucion.ListIndex = 1
        'Indistinta
        Else
            cboInstitucion.ListIndex = 2
        End If
    End If
    
End Sub

Private Function existeLetraEnSocio(letra As String) As Boolean
    Dim i As Integer
    Dim blPuntero As Boolean
    
    blPuntero = False
    For i = 1 To grdCliente.rows - 1
        If Trim(UCase(letra)) = Trim(UCase(grdCliente.TextMatrix(i, 6))) Then
            blPuntero = True
        End If
    Next
    
    existeLetraEnSocio = blPuntero
End Function

Private Function existeRegla(strRegla As String) As Boolean

    Dim blReglaExiste As Boolean
    Dim i As Integer
    blReglaExiste = False
    For i = 0 To grdReglas.rows - 1
        If Trim(grdReglas.TextMatrix(i, 1)) = Trim(strRegla) Then
            blReglaExiste = True
         End If
    Next
    existeRegla = blReglaExiste
    
End Function

Private Sub limpiarReglas()
    Dim i As Integer
    For i = 1 To grdReglas.rows - 1
        grdReglas.EliminaFila grdReglas.row
    Next
     For i = 0 To lsLetras.ListCount - 1
        lsLetras.Selected(i) = False
    Next
End Sub

Private Function prepararRegla() As String
    
    Dim i As Integer
    Dim strCadena As String
    
    For i = 1 To grdReglas.rows - 1
        If i = 1 Then
            strCadena = grdReglas.TextMatrix(i, 1)
        Else
            strCadena = strCadena & "-" & grdReglas.TextMatrix(i, 1)
        End If
    Next
    prepararRegla = strCadena
    
End Function

Private Function prepararGrupoPersona() As String
    Dim i As Integer
    Dim strGrupo As String
    
    For i = 1 To grdCliente.rows - 1
        If i = 1 Then
            strGrupo = grdCliente.TextMatrix(i, 6)
        Else
            strGrupo = strGrupo & "-" & grdCliente.TextMatrix(i, 6)
        End If
    Next
    prepararGrupoPersona = strGrupo
    
End Function

' *** END RIRO

'RIRO20140710 ERS098
Private Sub HabilitarPoderes(ByVal bHabilitar As Boolean)
    'JUEZ 20151229 ***************************************
    If nProducto = gCapAhorros And intPunteroPJ_NA > 0 And Not fbMantPoderes And validaExistenciaReglas Then
        bHabilitar = False
    End If
    'END JUEZ ********************************************
    If bHabilitar Then
        lsLetras.Enabled = True
        cmdAgregarRegla.Enabled = True
        cmdQuitarRega.Enabled = True
    Else
        lsLetras.Enabled = False
        cmdAgregarRegla.Enabled = False
        cmdQuitarRega.Enabled = False
    End If
End Sub
'END RIRO
'FRHU 20140927 ERS099-2014
Private Sub EmiteCalendarioRetiroIntPFMensual(ByVal nCapital As Double, ByVal nTasa As Double, ByVal nPlazo As Long, _
            ByVal dApertura As Date, ByVal nMoneda As Moneda, Optional ByVal nDiasVal As Integer = 0, Optional sCuenta As String = "", Optional nCostoMan As Currency = 0)

Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Dim nIntMens As Double, nIntFinal As Double
Dim dFecVenc As Date, dFecVal As Date
Dim lnTotIntMes As Double
    
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
dFecVenc = DateAdd("d", nPlazo + nDiasVal, dApertura)
dFecVal = DateAdd("d", nDiasVal, dApertura)
nIntMens = clsMant.GetInteresPF(nTasa, nCapital, 30)
nIntFinal = clsMant.GetInteresPF(nTasa, nCapital, nPlazo)

Set clsMant = Nothing

Dim clsPrev As previo.clsprevio
Dim sCad As String
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
sCad = clsMant.GetPFPlanRetInt(dApertura, Round(nIntMens, 2), nPlazo, nMoneda, Round(nIntFinal, 2), nCapital, nTasa, nDiasVal, dFecVal, lnTotIntMes, sCuenta, nCostoMan)
    
Set clsMant = Nothing

Set clsPrev = New previo.clsprevio
    clsPrev.Show sCad, "Plazo Fijo", True, , gImpresora
Set clsPrev = Nothing
End Sub
'FIN FRHU
