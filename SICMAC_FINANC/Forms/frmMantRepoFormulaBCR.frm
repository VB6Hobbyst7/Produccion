VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMantRepoFormulaBCR 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   10605
   ClientLeft      =   660
   ClientTop       =   2760
   ClientWidth     =   12615
   Icon            =   "frmMantRepoFormulaBCR.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10605
   ScaleWidth      =   12615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      TabIndex        =   29
      Top             =   10080
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   17383
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Reporte 1"
      TabPicture(0)   =   "frmMantRepoFormulaBCR.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label17"
      Tab(0).Control(1)=   "feNotas"
      Tab(0).Control(2)=   "cmdBajar"
      Tab(0).Control(3)=   "cmdSubir"
      Tab(0).Control(4)=   "cmdEditar"
      Tab(0).Control(5)=   "cmdGuardar"
      Tab(0).Control(6)=   "cmdCancelar"
      Tab(0).Control(7)=   "cmdModificar"
      Tab(0).Control(8)=   "Frame12"
      Tab(0).Control(9)=   "cmdNuevo"
      Tab(0).Control(10)=   "cmdQuitar"
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Reporte 2"
      TabPicture(1)   =   "frmMantRepoFormulaBCR.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label18"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "FeRep2SubColBCR"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdNuevoRep2SubCol"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdEditarRep2SubCol"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdQuitarRep2Subcol"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdGuardarRep2Subcol"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdCancelarRep2SubCol"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdModificarRep2SubCol"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmdSubirRep2SubCol"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cmdBajarRep2Subcol"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Reporte 3"
      TabPicture(2)   =   "frmMantRepoFormulaBCR.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label19"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame4"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "FeRep3SubColBCR"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdNuevoRep3SubCol"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmdEditarRep3SubCol"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cmdQuitarRep3SubCol"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "cmdBajarRep3SubCol"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "cmdSubirRep3SubCol"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "cmdModificarRep3SubCol"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "cmdCancelarRep3SubCol"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "cmdGuardarRep3SubCol"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "Reporte 4"
      TabPicture(3)   =   "frmMantRepoFormulaBCR.frx":035E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label13"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame5"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame6"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "FeRep4SubColBCR"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cmdNuevoRep4SubCol"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "cmdEditarRep4SubCol"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "cmdQuitarRep4SubCol"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "cmdBajarRep4SubCol"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "cmdSubirRep4SubCol"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "cmdModificarRep4SubCol"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "cmdCancelarRep4SubCol"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "cmdGuardarRep4SubCol"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).ControlCount=   12
      Begin VB.CommandButton cmdGuardarRep4SubCol 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   6600
         TabIndex        =   149
         Top             =   9360
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelarRep4SubCol 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   7680
         TabIndex        =   148
         Top             =   9360
         Width           =   1095
      End
      Begin VB.CommandButton cmdModificarRep4SubCol 
         Caption         =   "Modificar"
         Height          =   375
         Left            =   8760
         TabIndex        =   147
         Top             =   9360
         Width           =   1095
      End
      Begin VB.CommandButton cmdSubirRep4SubCol 
         Caption         =   "Subir"
         Height          =   375
         Left            =   9840
         TabIndex        =   146
         Top             =   9360
         Width           =   1095
      End
      Begin VB.CommandButton cmdBajarRep4SubCol 
         Caption         =   "Bajar"
         Height          =   375
         Left            =   10920
         TabIndex        =   145
         Top             =   9360
         Width           =   1095
      End
      Begin VB.CommandButton cmdQuitarRep4SubCol 
         Caption         =   "Quitar"
         Height          =   375
         Left            =   2280
         TabIndex        =   144
         Top             =   9360
         Width           =   1095
      End
      Begin VB.CommandButton cmdEditarRep4SubCol 
         Caption         =   "Editar"
         Height          =   375
         Left            =   1200
         TabIndex        =   143
         Top             =   9360
         Width           =   1095
      End
      Begin VB.CommandButton cmdNuevoRep4SubCol 
         Caption         =   "Nuevo"
         Height          =   375
         Left            =   120
         TabIndex        =   142
         Top             =   9360
         Width           =   1095
      End
      Begin Sicmact.FlexEdit FeRep4SubColBCR 
         Height          =   2775
         Left            =   120
         TabIndex        =   137
         Top             =   6360
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   4895
         Cols0           =   14
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Descripción-Cod. Swift-Columna-Desde-Hasta-Plazo Prom.-Valor-Item-MovNro-ValorRef-nValor-cCodigo-bAplicaPer"
         EncabezadosAnchos=   "350-1600-1200-2500-1200-1200-1200-2500-0-0-0-0-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-L-L-L-L-C-R-L-R-L-R-R-L-R"
         FormatosEdit    =   "0-1-1-1-1-1-2-1-3-1-3-3-1-3"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Frame Frame6 
         Caption         =   "SubColumnas"
         Height          =   2415
         Left            =   120
         TabIndex        =   118
         Top             =   3840
         Width           =   11895
         Begin VB.CommandButton cmdCancelarBRep4 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   9960
            TabIndex        =   141
            Top             =   1920
            Width           =   1095
         End
         Begin VB.CommandButton cmdAceptarBRep4 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   9960
            TabIndex        =   140
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Frame frmRep4ApliPeriod 
            Height          =   735
            Left            =   240
            TabIndex        =   129
            Top             =   1560
            Width           =   8775
            Begin VB.TextBox txtPlazPromRep4SubCol 
               Height          =   330
               Left            =   5880
               TabIndex        =   136
               Top             =   240
               Width           =   1815
            End
            Begin MSComCtl2.DTPicker dtpHastaRep4SubCol 
               Height          =   375
               Left            =   2880
               TabIndex        =   134
               Top             =   240
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   661
               _Version        =   393216
               Format          =   62783489
               CurrentDate     =   41701
            End
            Begin MSComCtl2.DTPicker dtpDesdeRep4SubCol 
               Height          =   375
               Left            =   960
               TabIndex        =   132
               Top             =   240
               Width           =   1355
               _ExtentX        =   2381
               _ExtentY        =   661
               _Version        =   393216
               Format          =   62783489
               CurrentDate     =   41701
            End
            Begin VB.CheckBox chbApliPer4 
               Caption         =   "Aplica Periodo"
               Height          =   255
               Left            =   240
               TabIndex        =   130
               Top             =   0
               Width           =   1335
            End
            Begin VB.Label Label32 
               Caption         =   "Plazo Promedio:"
               Height          =   255
               Left            =   4680
               TabIndex        =   135
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label31 
               Caption         =   "Hasta:"
               Height          =   255
               Left            =   2400
               TabIndex        =   133
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label30 
               Caption         =   "Desde:"
               Height          =   255
               Left            =   360
               TabIndex        =   131
               Top             =   360
               Width           =   615
            End
         End
         Begin VB.Frame frm_ValorRep4SubCol 
            Caption         =   "Valor"
            Height          =   735
            Left            =   240
            TabIndex        =   125
            Top             =   720
            Width           =   11175
            Begin VB.OptionButton optTotColRep4SubCol 
               Caption         =   "Sub Total(Total Columna)"
               Height          =   195
               Left            =   8760
               TabIndex        =   128
               Top             =   360
               Width           =   2175
            End
            Begin VB.TextBox txtForRep4SubCol 
               Height          =   330
               Left            =   1080
               TabIndex        =   127
               Top             =   240
               Width           =   7455
            End
            Begin VB.OptionButton optForRep4SubCol 
               Caption         =   "Formula:"
               Height          =   195
               Left            =   120
               TabIndex        =   126
               Top             =   360
               Width           =   975
            End
         End
         Begin VB.ComboBox cboColumnaRep4SubCol 
            Height          =   315
            Left            =   8760
            Style           =   2  'Dropdown List
            TabIndex        =   124
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox txtCodSwifRep4SubCol 
            Height          =   330
            Left            =   6000
            TabIndex        =   122
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtDescripcionRep4SubCol 
            Height          =   330
            Left            =   1200
            TabIndex        =   120
            Top             =   240
            Width           =   3615
         End
         Begin VB.Label Label29 
            Caption         =   "Columna:"
            Height          =   255
            Left            =   8040
            TabIndex        =   123
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label20 
            Caption         =   "Cod. Prog:"
            Height          =   255
            Left            =   5040
            TabIndex        =   121
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label16 
            Caption         =   "Descripción:"
            Height          =   255
            Left            =   240
            TabIndex        =   119
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Columnas Generales"
         Height          =   2895
         Left            =   120
         TabIndex        =   110
         Top             =   840
         Width           =   11895
         Begin VB.CommandButton cmdQuitarARep4 
            Caption         =   "Quitar"
            Height          =   375
            Left            =   1320
            TabIndex        =   139
            Top             =   2400
            Width           =   1095
         End
         Begin VB.CommandButton cmdEditarARep4 
            Caption         =   "Editar"
            Height          =   375
            Left            =   240
            TabIndex        =   138
            Top             =   2400
            Width           =   1095
         End
         Begin VB.CommandButton cmdCancelarARep4 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   8400
            TabIndex        =   117
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdAceptarARep4 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   7320
            TabIndex        =   116
            Top             =   360
            Width           =   1095
         End
         Begin Sicmact.FlexEdit FeRep4ColBCR 
            Height          =   1575
            Left            =   240
            TabIndex        =   115
            Top             =   720
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   2778
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Código-Descripción-Moneda-MovNro-Item"
            EncabezadosAnchos=   "350-1000-4500-0-0-0"
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
            ColumnasAEditar =   "X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-R-L-R-L-R"
            FormatosEdit    =   "0-3-1-3-1-3"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   345
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin VB.TextBox txtCodigoRep4Col 
            Height          =   330
            Left            =   5040
            TabIndex        =   114
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtDescripcionRep4Col 
            Height          =   330
            Left            =   1320
            TabIndex        =   112
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label15 
            Caption         =   "Codigo:"
            Height          =   255
            Left            =   4440
            TabIndex        =   113
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label14 
            Caption         =   "Descripción:"
            Height          =   255
            Left            =   240
            TabIndex        =   111
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdGuardarRep3SubCol 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   -68400
         TabIndex        =   104
         Top             =   9360
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelarRep3SubCol 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   -67320
         TabIndex        =   103
         Top             =   9360
         Width           =   1095
      End
      Begin VB.CommandButton cmdModificarRep3SubCol 
         Caption         =   "Modificar"
         Height          =   375
         Left            =   -66240
         TabIndex        =   102
         Top             =   9360
         Width           =   1095
      End
      Begin VB.CommandButton cmdSubirRep3SubCol 
         Caption         =   "Subir"
         Height          =   375
         Left            =   -65160
         TabIndex        =   101
         Top             =   9360
         Width           =   1095
      End
      Begin VB.CommandButton cmdBajarRep3SubCol 
         Caption         =   "Bajar"
         Height          =   375
         Left            =   -64080
         TabIndex        =   100
         Top             =   9360
         Width           =   1095
      End
      Begin VB.CommandButton cmdQuitarRep3SubCol 
         Caption         =   "Quitar"
         Height          =   375
         Left            =   -72720
         TabIndex        =   99
         Top             =   9360
         Width           =   1095
      End
      Begin VB.CommandButton cmdEditarRep3SubCol 
         Caption         =   "Editar"
         Height          =   375
         Left            =   -73800
         TabIndex        =   98
         Top             =   9360
         Width           =   1095
      End
      Begin VB.CommandButton cmdNuevoRep3SubCol 
         Caption         =   "Nuevo"
         Height          =   375
         Left            =   -74880
         TabIndex        =   97
         Top             =   9360
         Width           =   1095
      End
      Begin Sicmact.FlexEdit FeRep3SubColBCR 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   96
         Top             =   6480
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   4895
         Cols0           =   15
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Institución-Código Swif-Destino-Desde-Hasta-Plazo Prom.-Valor-Columna-nItem-nMoneda-cMovNro-nValor-CodDest-nValorRef"
         EncabezadosAnchos=   "350-1600-1000-2000-1000-1000-1200-1600-1800-0-0-0-0-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-L-L-C-C-C-R-L-C-R-C-L-R-R-R"
         FormatosEdit    =   "1-1-1-1-1-1-2-1-1-3-3-1-3-3-3"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Frame Frame4 
         Caption         =   "SubColumnas"
         Height          =   2535
         Left            =   -74880
         TabIndex        =   74
         Top             =   3840
         Width           =   11895
         Begin VB.Frame frm_ValorRep3SubCol 
            Caption         =   "Valor"
            Height          =   1215
            Left            =   240
            TabIndex        =   91
            Top             =   240
            Width           =   11175
            Begin VB.TextBox txtCodOpeRep3SubCol 
               Height          =   330
               Left            =   4800
               TabIndex        =   107
               Top             =   720
               Width           =   1335
            End
            Begin VB.OptionButton optTotColRep3SubCol 
               Caption         =   "Totalizar Columnas:"
               Height          =   255
               Left            =   2280
               TabIndex        =   95
               Top             =   840
               Width           =   1695
            End
            Begin VB.OptionButton optSubTotRep3SubCol 
               Caption         =   "Sub total (total columna)"
               Height          =   255
               Left            =   120
               TabIndex        =   94
               Top             =   840
               Width           =   2055
            End
            Begin VB.TextBox txtForRep3SubCol 
               Height          =   330
               Left            =   1080
               TabIndex        =   93
               Top             =   240
               Width           =   9495
            End
            Begin VB.OptionButton optForRep3SubCol 
               Caption         =   "Fórmula"
               Height          =   255
               Left            =   120
               TabIndex        =   92
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label28 
               Caption         =   "(Cod. Ope)"
               Height          =   255
               Left            =   3960
               TabIndex        =   108
               Top             =   840
               Width           =   855
            End
         End
         Begin VB.CommandButton cmdCancelarBRep3 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   10320
            TabIndex        =   90
            Top             =   2040
            Width           =   1095
         End
         Begin VB.CommandButton cmdAceptarBRep3 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   9240
            TabIndex        =   89
            Top             =   2040
            Width           =   1095
         End
         Begin VB.ComboBox cboColRep3SubCol 
            Height          =   315
            Left            =   6720
            Style           =   2  'Dropdown List
            TabIndex        =   88
            Top             =   2040
            Width           =   2415
         End
         Begin VB.ComboBox cboTipOblRep3SubCol 
            Height          =   315
            Left            =   4680
            Style           =   2  'Dropdown List
            TabIndex        =   87
            Top             =   2040
            Width           =   1935
         End
         Begin VB.TextBox txtPlazPromRep3SubCol 
            Height          =   330
            Left            =   10200
            TabIndex        =   85
            Top             =   1560
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker dtpHastaRep3SubCol 
            Height          =   330
            Left            =   2400
            TabIndex        =   83
            Top             =   2040
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            Format          =   62783489
            CurrentDate     =   41684
         End
         Begin MSComCtl2.DTPicker dtpDesdeRep3SubCol 
            Height          =   330
            Left            =   1080
            TabIndex        =   82
            Top             =   2040
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            Format          =   62783489
            CurrentDate     =   41684
         End
         Begin VB.ComboBox cboDestRep3SubCol 
            Height          =   315
            Left            =   6960
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   1560
            Width           =   2175
         End
         Begin VB.TextBox txtCodSwifRep3SubCol 
            Height          =   330
            Left            =   4680
            TabIndex        =   78
            Top             =   1560
            Width           =   1455
         End
         Begin VB.TextBox txtInstRep3SubCol 
            Height          =   330
            Left            =   1080
            TabIndex        =   76
            Top             =   1560
            Width           =   2655
         End
         Begin VB.Label Label27 
            Caption         =   "Columna:"
            Height          =   255
            Left            =   3840
            TabIndex        =   86
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label Label26 
            Caption         =   "Plazo Prom.:"
            Height          =   255
            Left            =   9240
            TabIndex        =   84
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label12 
            Caption         =   "Periodo:"
            Height          =   255
            Left            =   120
            TabIndex        =   81
            Top             =   2160
            Width           =   615
         End
         Begin VB.Label Label11 
            Caption         =   "Destino:"
            Height          =   255
            Left            =   6240
            TabIndex        =   79
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label Label10 
            Caption         =   "Cód. swift:"
            Height          =   255
            Left            =   3840
            TabIndex        =   77
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Descripción:"
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   1680
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Columnas Generales"
         Height          =   2895
         Left            =   -74880
         TabIndex        =   64
         Top             =   840
         Width           =   11895
         Begin VB.CommandButton cmdQuitarARep3 
            Caption         =   "Quitar"
            Height          =   375
            Left            =   1320
            TabIndex        =   106
            Top             =   2400
            Width           =   1095
         End
         Begin VB.CommandButton cmdEditarARep3 
            Caption         =   "Editar"
            Height          =   375
            Left            =   240
            TabIndex        =   105
            Top             =   2400
            Width           =   1095
         End
         Begin Sicmact.FlexEdit FeRep3ColBCR 
            Height          =   1575
            Left            =   240
            TabIndex        =   73
            Top             =   720
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   2778
            Cols0           =   8
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Codigo-Descripción-Tipo Obligación-nItem-nMoneda-cMovNro-nCodObl"
            EncabezadosAnchos=   "350-1000-4000-4000-0-0-0-0"
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
            ColumnasAEditar =   "X-X-X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "L-R-C-L-R-C-L-R"
            FormatosEdit    =   "1-3-0-1-3-3-1-3"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   345
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin VB.CommandButton cmdCancelarARep3 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   9960
            TabIndex        =   72
            Top             =   720
            Width           =   1095
         End
         Begin VB.CommandButton cmdAceptarARep3 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   9960
            TabIndex        =   71
            Top             =   360
            Width           =   1095
         End
         Begin VB.ComboBox cboTipOblRep3Col 
            Height          =   315
            Left            =   7560
            Style           =   2  'Dropdown List
            TabIndex        =   70
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox txtCodigoRep3Col 
            Height          =   330
            Left            =   5040
            TabIndex        =   68
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtDescripcionRep3Col 
            Height          =   330
            Left            =   1320
            TabIndex        =   66
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label8 
            Caption         =   "Obligaciones:"
            Height          =   255
            Left            =   6480
            TabIndex        =   69
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "Código:"
            Height          =   255
            Left            =   4440
            TabIndex        =   67
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Descripción:"
            Height          =   255
            Left            =   240
            TabIndex        =   65
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdBajarRep2Subcol 
         Caption         =   "Bajar"
         Height          =   375
         Left            =   -64080
         TabIndex        =   63
         Top             =   9360
         Width           =   1095
      End
      Begin VB.CommandButton cmdSubirRep2SubCol 
         Caption         =   "Subir"
         Height          =   375
         Left            =   -65160
         TabIndex        =   62
         Top             =   9360
         Width           =   1095
      End
      Begin VB.CommandButton cmdModificarRep2SubCol 
         Caption         =   "Modificar"
         Height          =   375
         Left            =   -66240
         TabIndex        =   61
         Top             =   9360
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelarRep2SubCol 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   -67320
         TabIndex        =   60
         Top             =   9360
         Width           =   1095
      End
      Begin VB.CommandButton cmdGuardarRep2Subcol 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   -68400
         TabIndex        =   59
         Top             =   9360
         Width           =   1095
      End
      Begin VB.CommandButton cmdQuitarRep2Subcol 
         Caption         =   "Quitar"
         Height          =   375
         Left            =   -72720
         TabIndex        =   58
         Top             =   9360
         Width           =   1095
      End
      Begin VB.CommandButton cmdEditarRep2SubCol 
         Caption         =   "Editar"
         Height          =   375
         Left            =   -73800
         TabIndex        =   57
         Top             =   9360
         Width           =   1095
      End
      Begin VB.CommandButton cmdNuevoRep2SubCol 
         Caption         =   "Nuevo"
         Height          =   375
         Left            =   -74880
         TabIndex        =   56
         Top             =   9360
         Width           =   1095
      End
      Begin Sicmact.FlexEdit FeRep2SubColBCR 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   55
         Top             =   6000
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   5530
         Cols0           =   11
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Código Swif-Columna-Sub Columna-Valor-Item-Moneda-cMovNro-nValor-cCodigo-nValorRef"
         EncabezadosAnchos=   "350-1300-1800-4000-4300-0-0-0-0-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-C-L-R-R-L-R-L-R"
         FormatosEdit    =   "0-1-1-1-1-3-3-1-3-1-3"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Frame Frame2 
         Caption         =   "SubColumnas"
         Height          =   2055
         Left            =   -74880
         TabIndex        =   42
         Top             =   3840
         Width           =   11895
         Begin VB.Frame frm_ValorRep2SubCol 
            Caption         =   "Valor"
            Height          =   1215
            Left            =   240
            TabIndex        =   51
            Top             =   720
            Width           =   10095
            Begin VB.OptionButton optTotColRep2SubCol 
               Caption         =   "Total Columna"
               Height          =   255
               Left            =   120
               TabIndex        =   54
               Top             =   840
               Width           =   1335
            End
            Begin VB.TextBox txtForRep2SubCol 
               Height          =   330
               Left            =   1320
               TabIndex        =   53
               Top             =   240
               Width           =   8655
            End
            Begin VB.OptionButton optForRep2SubCol 
               Caption         =   "Fórmula"
               Height          =   255
               Left            =   120
               TabIndex        =   52
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.CommandButton cmdCancelarBRep2 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   10560
            TabIndex        =   50
            Top             =   600
            Width           =   1095
         End
         Begin VB.CommandButton cmdAceptarBRep2 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   10560
            TabIndex        =   49
            Top             =   240
            Width           =   1095
         End
         Begin VB.ComboBox cboColumnaRep2SubCol 
            Height          =   315
            Left            =   8640
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtCodSwifRep2SubCol 
            Height          =   330
            Left            =   5760
            TabIndex        =   46
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtDescripcionRep2SubCol 
            Height          =   330
            Left            =   1200
            TabIndex        =   44
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label Label5 
            Caption         =   "Columna : "
            Height          =   255
            Left            =   7800
            TabIndex        =   47
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Descripción :"
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Código swif : "
            Height          =   255
            Left            =   4800
            TabIndex        =   45
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Columnas Generales"
         Height          =   2895
         Left            =   -74880
         TabIndex        =   32
         Top             =   840
         Width           =   11895
         Begin VB.CommandButton cmdCancelarARep2 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   8400
            TabIndex        =   41
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdEditarARep2 
            Caption         =   "Editar"
            Height          =   375
            Left            =   240
            TabIndex        =   40
            Top             =   2400
            Width           =   1095
         End
         Begin Sicmact.FlexEdit FeRep2ColBCR 
            Height          =   1575
            Left            =   240
            TabIndex        =   39
            Top             =   720
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   2778
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Código-Descripción-Moneda-MovNro-Item"
            EncabezadosAnchos=   "350-1000-4500-0-0-0"
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
            ColumnasAEditar =   "X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "L-R-L-R-L-R"
            FormatosEdit    =   "0-3-1-3-1-3"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   345
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin VB.CommandButton cmdQuitarARep2 
            Caption         =   "Quitar"
            Height          =   375
            Left            =   1320
            TabIndex        =   38
            Top             =   2400
            Width           =   1095
         End
         Begin VB.CommandButton cmdAceptarARep2 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   7320
            TabIndex        =   37
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtCodigoRep2Col 
            Height          =   330
            Left            =   5040
            MaxLength       =   6
            TabIndex        =   36
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtDescripcionRep2Col 
            Height          =   330
            Left            =   1320
            TabIndex        =   34
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label2 
            Caption         =   "Código:"
            Height          =   255
            Left            =   4440
            TabIndex        =   35
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Descripción:"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         Height          =   375
         Left            =   -72720
         TabIndex        =   31
         Top             =   9360
         Width           =   1095
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Height          =   375
         Left            =   -74880
         TabIndex        =   30
         Top             =   9360
         Width           =   1095
      End
      Begin VB.Frame Frame12 
         Caption         =   "Columnas"
         Height          =   3255
         Left            =   -74880
         TabIndex        =   11
         Top             =   840
         Width           =   11895
         Begin VB.ComboBox cboRegimen 
            Height          =   315
            ItemData        =   "frmMantRepoFormulaBCR.frx":037A
            Left            =   1080
            List            =   "frmMantRepoFormulaBCR.frx":037C
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   960
            Width           =   4455
         End
         Begin VB.Frame Frame13 
            Caption         =   "Valor:"
            Height          =   1695
            Left            =   120
            TabIndex        =   21
            Top             =   1440
            Width           =   11415
            Begin VB.TextBox txtFormula 
               Height          =   330
               Left            =   1320
               TabIndex        =   26
               Top             =   240
               Width           =   9855
            End
            Begin VB.TextBox txtTotalizado 
               Height          =   330
               Left            =   1320
               TabIndex        =   25
               Top             =   720
               Width           =   3495
            End
            Begin VB.OptionButton Option3 
               Caption         =   "Promedio Caja Mes Anterior"
               Height          =   315
               Left            =   120
               TabIndex        =   24
               Top             =   1200
               Width           =   2415
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Totalizado"
               Height          =   315
               Left            =   120
               TabIndex        =   23
               Top             =   720
               Width           =   1215
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Fórmula"
               Height          =   315
               Left            =   120
               TabIndex        =   22
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               Caption         =   "(Escriba los ""N"" de Columnas a totalizar)"
               Height          =   195
               Left            =   4920
               TabIndex        =   27
               Top             =   840
               Width           =   2850
            End
         End
         Begin VB.CommandButton cmdCancelarA 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   8760
            TabIndex        =   16
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtNumCol 
            Height          =   330
            Left            =   600
            MaxLength       =   3
            TabIndex        =   15
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtCod 
            Height          =   330
            Left            =   6480
            MaxLength       =   6
            TabIndex        =   14
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton cmdAceptarA 
            Caption         =   "&Aceptar"
            Height          =   375
            Left            =   8760
            TabIndex        =   13
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   330
            Left            =   2160
            MaxLength       =   100
            TabIndex        =   12
            Top             =   360
            Width           =   5775
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Régimen :"
            Height          =   195
            Left            =   240
            TabIndex        =   20
            Top             =   1080
            Width           =   720
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Código :"
            Height          =   195
            Left            =   5760
            TabIndex        =   19
            Top             =   1080
            Width           =   585
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Descripción :"
            Height          =   195
            Left            =   1200
            TabIndex        =   18
            Top             =   480
            Width           =   930
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Nº :"
            Height          =   195
            Left            =   240
            TabIndex        =   17
            Top             =   480
            Width           =   270
         End
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         Height          =   375
         Left            =   -66240
         TabIndex        =   10
         Top             =   9360
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   -67320
         TabIndex        =   9
         Top             =   9360
         Width           =   1095
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   -68400
         TabIndex        =   8
         Top             =   9360
         Width           =   1095
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar"
         Height          =   375
         Left            =   -73800
         TabIndex        =   7
         Top             =   9360
         Width           =   1095
      End
      Begin VB.CommandButton cmdSubir 
         Caption         =   "&Subir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -65160
         TabIndex        =   6
         Top             =   9360
         Width           =   1095
      End
      Begin VB.CommandButton cmdBajar 
         Caption         =   "&Bajar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -64080
         TabIndex        =   5
         Top             =   9360
         Width           =   1095
      End
      Begin Sicmact.FlexEdit feNotas 
         Height          =   4965
         Left            =   -74880
         TabIndex        =   4
         Top             =   4200
         Width           =   12000
         _ExtentX        =   21167
         _ExtentY        =   8758
         Cols0           =   11
         HighLight       =   1
         EncabezadosNombres=   "#-No.Col.-Descripción-Código-Regimen-Valor-nItem-nCodRegimen-nMoneda-cMovNro-nValor"
         EncabezadosAnchos=   "350-800-3900-850-2100-3900-0-0-0-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-R-L-L-L-L-R-R-C-L-C"
         FormatosEdit    =   "0-3-0-0-0-0-3-3-3-0-0"
         CantEntero      =   9
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         TipoBusqueda    =   0
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
      End
      Begin VB.Label Label13 
         Caption         =   "OTRAS OBLIGACIONES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   109
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label19 
         Caption         =   "OBLIGACIONES CON INSTITUCIONES FINANCIERAS DEL EXTERIOR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   8055
      End
      Begin VB.Label Label18 
         Caption         =   "OBLIGACIONES NO SUJETAS A ENCAJE CON INSTITUCIONES FINANCIERAS DEL PAIS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   9735
      End
      Begin VB.Label Label17 
         Caption         =   "OBLIGACIONES SUJETAS A ENCAJE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   7095
      End
   End
End
Attribute VB_Name = "frmMantRepoFormulaBCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**-------------------------------------------------------------------------------------**'
'** Formulario : frmMantRepoFormulaBCR                                                  **'
'** Finalidad  : Este formulario permite configurar los reportes de encaje para el BCR, **'
'**              donde el mismo usuario pueda administrarlo asi mismo esto permitirá    **'
'**              emitir los archivos TXT del SUCAVE.                                    **'
'** Programador: Paolo Hector Sinti Cabrera - PASI
'** Fecha/Hora : 20140205 11:50 AM                                                      **'
'**-------------------------------------------------------------------------------------**'

Option Explicit
Dim nAccion As Integer
Dim sOpeCod As String
Dim fsOpeCod As String
Dim rsRep   As ADODB.Recordset
Dim rsRegRep As ADODB.Recordset
Dim clsRep  As DRepFormula
Dim sInserModif As Integer
Dim lcMovNro As String
Dim i As Integer
Dim lsMoneda As String
Dim lnValor As Integer
Dim lnNumItem As Integer
'*** PASI20140207
Dim nInserModifRep2Col As Integer
Dim nInserModifRep2SubCol As Integer
Dim nInserModifRep3Col As Integer
Dim nInserModifRep3SubCol As Integer
Dim nInserModifRep4Col As Integer
Dim nInserModifRep4SubCol As Integer
Dim cCodigo As String
Dim nItem As Integer
Dim nCol As Integer
Dim sDescripcion As String
Dim nCodObl As Integer

'----------------------------------------------------------------------------------------PESTAÑA1-----------------------------------------------------------------------------------
Private Sub cmdAceptarA_Click()
If sInserModif = 1 Then
    If ValidarRep1 = False Then
        Exit Sub
    End If
    lnValor = IIf(Me.Option1.value = True, 1, IIf(Me.Option2.value = True, 2, 3))
    lcMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
    lnNumItem = IIf(IsNumeric(feNotas.TextMatrix(1, 1)), feNotas.Rows, feNotas.Rows - 1)
    clsRep.InsertaRep1FormulaBCR Me.txtNumCol.Text, Me.txtDescripcion.Text, Me.txtCod.Text, Trim(Right(cboRegimen.Text, 2)), IIf(Me.Option1.value, Me.txtFormula.Text, IIf(Me.Option2.value, Me.txtTotalizado.Text, "Promedio Caja Mes Anterior")), lnNumItem, lsMoneda, lcMovNro, lnValor
Else
    If ValidarRep1 = False Then
        Exit Sub
    End If
    lnValor = IIf(Me.Option1.value = True, 1, IIf(Me.Option2.value = True, 2, 3))
    lcMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
    clsRep.ModificaRep1FormulaBCR Me.txtNumCol.Text, Me.txtDescripcion.Text, Me.txtCod.Text, Trim(Right(cboRegimen.Text, 2)), IIf(Me.Option1.value, Me.txtFormula.Text, IIf(Me.Option2.value, Me.txtTotalizado.Text, "Promedio Caja Mes Anterior")), nItem, lsMoneda, lcMovNro, lnValor
End If
    LimpiaControles
    CargaReportes (1)
    Call ControlaAccion(4)
    sInserModif = 1
    nCol = 0
    cCodigo = ""
End Sub
Public Function ValidarRep1() As Boolean
    If Len(Me.txtNumCol.Text) = 0 Then
        MsgBox "Ingrese el número de columna.", vbOKOnly + vbExclamation, "Atención"
        Me.txtCod.SetFocus
        ValidarRep1 = False
        Exit Function
    End If
    If Len(Me.txtDescripcion.Text) = 0 Then
        MsgBox "Ingrese la descripción de columna.", vbOKOnly + vbExclamation, "Atención"
        Me.txtDescripcion.SetFocus
        ValidarRep1 = False
        Exit Function
    End If
    If Me.cboRegimen.ListIndex = -1 Then
        MsgBox "Selecciones un tipo de regimen.", vbOKOnly + vbExclamation, "Atención"
        Me.cboRegimen.SetFocus
        ValidarRep1 = False
        Exit Function
    End If
    If Len(Me.txtCod.Text) = 0 Then
        MsgBox "Ingrese el código de columna.", vbOKOnly + vbExclamation, "Atención"
        Me.txtCod.SetFocus
        ValidarRep1 = False
        Exit Function
    End If
    If Option1.value = False And Option2.value = False And Option3.value = False Then
        MsgBox "Seleccione el tipo de valor para la columna.", vbOKOnly + vbExclamation, "Atención"
        Me.Option1.SetFocus
        ValidarRep1 = False
        Exit Function
    End If
    If Option1.value And Len(Me.txtFormula.Text) = 0 Then
        MsgBox "No se ha ingresado el valor de la formula.", vbOKOnly + vbExclamation, "Atención"
        Me.txtFormula.SetFocus
        ValidarRep1 = False
        Exit Function
    End If
    If Option2.value And Len(Me.txtTotalizado.Text) = 0 Then
        MsgBox "No se ha ingresado el valor totalizado.", vbOKOnly + vbExclamation, "Atención"
        Me.txtTotalizado.SetFocus
        ValidarRep1 = False
        Exit Function
    End If
    For i = 1 To feNotas.Rows - 1
        If feNotas.TextMatrix(i, 1) = Me.txtNumCol.Text Then
            If feNotas.TextMatrix(i, 1) <> nCol Then
                MsgBox "El número de columna (" & Trim(Me.txtNumCol.Text) & ") ya fue ingresado.", vbOKOnly + vbExclamation, "Atención"
                Me.txtNumCol.SetFocus
                ValidarRep1 = False
                Exit Function
            End If
        End If
        If feNotas.TextMatrix(i, 3) = Me.txtCod.Text Then
            If feNotas.TextMatrix(i, 3) <> cCodigo Then
                MsgBox "El código (" & Trim(Me.txtCod.Text) & ") ya fue ingresado.", vbOKOnly + vbExclamation, "Atención"
                Me.txtCod.SetFocus
                ValidarRep1 = False
                Exit Function
            End If
        End If
    Next i
    ValidarRep1 = True
End Function

Private Sub cmdCancelarA_Click()
    If sInserModif = 1 Then
        LimpiaControles
        ControlaAccion (5)
    Else
        LimpiaControles
        Call ControlaAccion(5) 'maneja controles igual que modificar
        sInserModif = 1
    End If
End Sub
Private Sub cmdNuevo_Click()
    ControlaAccion (10)
End Sub

Private Sub cmdEditar_Click()
If IsNumeric(feNotas.TextMatrix(feNotas.Row, 1)) Then
    Dim cvalor As String
    sInserModif = 2
    Call ControlaAccion(7)
    nCol = feNotas.TextMatrix(feNotas.Row, 1)
    cCodigo = feNotas.TextMatrix(feNotas.Row, 3)
    Me.txtNumCol.Text = feNotas.TextMatrix(feNotas.Row, 1)
    Me.txtDescripcion.Text = feNotas.TextMatrix(feNotas.Row, 2)
    Me.txtCod.Text = feNotas.TextMatrix(feNotas.Row, 3)
    nItem = feNotas.TextMatrix(feNotas.Row, 6)
    
    Me.cboRegimen.ListIndex = IndiceListaCombo(cboRegimen, feNotas.TextMatrix(feNotas.Row, 7))
    
    If feNotas.TextMatrix(feNotas.Row, 10) = "1" Then
        Me.txtFormula.Text = feNotas.TextMatrix(feNotas.Row, 5)
        Me.Option1.value = True
        Option1_Click
        Me.Option2.value = False
        Me.Option3.value = False
    ElseIf feNotas.TextMatrix(feNotas.Row, 10) = "2" Then
        Me.Option1.value = False
        Me.Option2.value = True
        Option2_Click
        Me.Option3.value = False
        cvalor = feNotas.TextMatrix(feNotas.Row, 5)
        cvalor = Replace(cvalor, "Totalizado : ", "", 1)
        Me.txtTotalizado.Text = cvalor
    Else
        Me.Option1.value = False
        Me.Option2.value = False
        Me.Option3.value = True
        Option3_Click
    End If
Else
    MsgBox "No existen datos para Editar", vbOKOnly + vbExclamation, "Atención"
End If
End Sub

Private Sub cmdQuitar_Click()
Dim Y As Integer
Dim nitemv As Integer
Y = 0
    If IsNumeric(feNotas.TextMatrix(feNotas.Row, 1)) Then
        nitemv = feNotas.TextMatrix(feNotas.Row, 1)
        If MsgBox(" ¿ Seguro que desea quitar Fila ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
            Exit Sub
        End If
        lcMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
        
        For i = 1 To feNotas.Rows - 1
            If feNotas.TextMatrix(i, 1) <> nitemv Then
                Y = Y + 1
                clsRep.ModificaRep1ItemOrdenFormulaBCR feNotas.TextMatrix(i, 1), Y, lsMoneda, lcMovNro
            Else
                clsRep.EliminaRep1FormulaBCR feNotas.TextMatrix(feNotas.Row, 1), CInt(feNotas.TextMatrix(feNotas.Row, 8))
            End If
        Next i
        CargaReportes (1)
    Else
        MsgBox "No existen datos para quitar", vbOKOnly + vbExclamation, "Atención"
    End If
     'ControlaAccion (2)
End Sub

Private Sub cmdGuardar_Click()
    If MsgBox(" ¿ Seguro de grabar el nuevo orden de las columnas ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbYes Then
        lcMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
        For i = 1 To feNotas.Rows - 1
            feNotas.TextMatrix(feNotas.Row, 6) = i
            clsRep.ModificaRep1ItemOrdenFormulaBCR feNotas.TextMatrix(i, 1), i, lsMoneda, lcMovNro
        Next i
    End If
    MsgBox "Se grabó satisfactoriamente", vbOKOnly + vbInformation, "Atención"
    feNotas.Clear
    feNotas.FormaCabecera
    feNotas.Rows = 2
    
    CargaReportes (1)
    ControlaAccion (3)
End Sub

Private Sub cmdCancelar_Click()
    ControlaAccion (2)
    feNotas.Clear
    feNotas.FormaCabecera
    feNotas.Rows = 2
    
    CargaReportes (1)
End Sub
Private Sub CmdModificar_Click()
    If IsNumeric(feNotas.TextMatrix(feNotas.Row, 1)) Then
        Call ControlaAccion(1)
    Else
        MsgBox "No existen datos para Modificar", vbOKOnly + vbExclamation, "Atención"
    End If
End Sub
Private Sub cmdSubir_Click()

    Dim lnNumCol1 As Integer, lnNumCol2 As Integer
    Dim lsDescripcion1 As String, lsDescripcion2 As String
    Dim lsCodigo1 As String, lsCodigo2 As String
    Dim lsRegimen1 As String, lsRegimen2 As String
    Dim lsValor1 As String, lsValor2 As String
    Dim lnItem1 As Integer, lnItem2 As Integer
    Dim lsCodReg1 As String, lsCodReg2 As String
    Dim lsMoneda1 As String, lsMoneda2 As String
    Dim lsMovNro1 As String, lsMovNro2 As String
    Dim lnValor1 As Integer, lnValor2 As Integer
    
    
'    If validarRegistroDatosNotasEstado = False Then Exit Sub

    If feNotas.Row > 1 Then
        'cambiamos las posiciones del flex
        lnNumCol1 = feNotas.TextMatrix(feNotas.Row - 1, 1)
        lsDescripcion1 = feNotas.TextMatrix(feNotas.Row - 1, 2)
        lsCodigo1 = feNotas.TextMatrix(feNotas.Row - 1, 3)
        lsRegimen1 = feNotas.TextMatrix(feNotas.Row - 1, 4)
        lsValor1 = feNotas.TextMatrix(feNotas.Row - 1, 5)
        lnItem1 = feNotas.TextMatrix(feNotas.Row - 1, 6)
        lsCodReg1 = feNotas.TextMatrix(feNotas.Row - 1, 7)
        lsMoneda1 = feNotas.TextMatrix(feNotas.Row - 1, 8)
        lsMovNro1 = feNotas.TextMatrix(feNotas.Row - 1, 9)
        lnValor1 = feNotas.TextMatrix(feNotas.Row - 1, 10)
        
                
        lnNumCol2 = feNotas.TextMatrix(feNotas.Row, 1)
        lsDescripcion2 = feNotas.TextMatrix(feNotas.Row, 2)
        lsCodigo2 = feNotas.TextMatrix(feNotas.Row, 3)
        lsRegimen2 = feNotas.TextMatrix(feNotas.Row, 4)
        lsValor2 = feNotas.TextMatrix(feNotas.Row, 5)
        lnItem2 = feNotas.TextMatrix(feNotas.Row, 6)
        lsCodReg2 = feNotas.TextMatrix(feNotas.Row, 7)
        lsMoneda2 = feNotas.TextMatrix(feNotas.Row, 8)
        lsMovNro2 = feNotas.TextMatrix(feNotas.Row, 9)
        lnValor2 = feNotas.TextMatrix(feNotas.Row, 10)
        
                        
        feNotas.TextMatrix(feNotas.Row - 1, 1) = lnNumCol2
        feNotas.TextMatrix(feNotas.Row - 1, 2) = lsDescripcion2
        feNotas.TextMatrix(feNotas.Row - 1, 3) = lsCodigo2
        feNotas.TextMatrix(feNotas.Row - 1, 4) = lsRegimen2
        feNotas.TextMatrix(feNotas.Row - 1, 5) = lsValor2
        feNotas.TextMatrix(feNotas.Row - 1, 6) = lnItem2
        feNotas.TextMatrix(feNotas.Row - 1, 7) = lsCodReg2
        feNotas.TextMatrix(feNotas.Row - 1, 8) = lsMoneda2
        feNotas.TextMatrix(feNotas.Row - 1, 9) = lsMovNro2
        feNotas.TextMatrix(feNotas.Row - 1, 10) = lnValor2
        

        feNotas.TextMatrix(feNotas.Row, 1) = lnNumCol1
        feNotas.TextMatrix(feNotas.Row, 2) = lsDescripcion1
        feNotas.TextMatrix(feNotas.Row, 3) = lsCodigo1
        feNotas.TextMatrix(feNotas.Row, 4) = lsRegimen1
        feNotas.TextMatrix(feNotas.Row, 5) = lsValor1
        feNotas.TextMatrix(feNotas.Row, 6) = lnItem1
        feNotas.TextMatrix(feNotas.Row, 7) = lsCodReg1
        feNotas.TextMatrix(feNotas.Row, 8) = lsMoneda1
        feNotas.TextMatrix(feNotas.Row, 9) = lsMovNro1
        feNotas.TextMatrix(feNotas.Row, 10) = lnValor1
        

        feNotas.Row = feNotas.Row - 1
        feNotas.SetFocus
    End If

End Sub
Private Sub cmdBajar_Click()

    Dim lnNumCol1 As Integer, lnNumCol2 As Integer
    Dim lsDescripcion1 As String, lsDescripcion2 As String
    Dim lsCodigo1 As String, lsCodigo2 As String
    Dim lsRegimen1 As String, lsRegimen2 As String
    Dim lsValor1 As String, lsValor2 As String
    Dim lnItem1 As Integer, lnItem2 As Integer
    Dim lsCodReg1 As String, lsCodReg2 As String
    Dim lsMoneda1 As String, lsMoneda2 As String
    Dim lsMovNro1 As String, lsMovNro2 As String
    Dim lnValor1 As Integer, lnValor2 As Integer
    

    'If validarRegistroDatosNotasEstado = False Then Exit Sub

    If feNotas.Row < feNotas.Rows - 1 Then
        'cambiamos las posiciones del flex
        
        lnNumCol1 = feNotas.TextMatrix(feNotas.Row + 1, 1)
        lsDescripcion1 = feNotas.TextMatrix(feNotas.Row + 1, 2)
        lsCodigo1 = feNotas.TextMatrix(feNotas.Row + 1, 3)
        lsRegimen1 = feNotas.TextMatrix(feNotas.Row + 1, 4)
        lsValor1 = feNotas.TextMatrix(feNotas.Row + 1, 5)
        lnItem1 = feNotas.TextMatrix(feNotas.Row + 1, 6)
        lsCodReg1 = feNotas.TextMatrix(feNotas.Row + 1, 7)
        lsMoneda1 = feNotas.TextMatrix(feNotas.Row + 1, 8)
        lsMovNro1 = feNotas.TextMatrix(feNotas.Row + 1, 9)
        lnValor1 = feNotas.TextMatrix(feNotas.Row + 1, 10)
        
                
        lnNumCol2 = feNotas.TextMatrix(feNotas.Row, 1)
        lsDescripcion2 = feNotas.TextMatrix(feNotas.Row, 2)
        lsCodigo2 = feNotas.TextMatrix(feNotas.Row, 3)
        lsRegimen2 = feNotas.TextMatrix(feNotas.Row, 4)
        lsValor2 = feNotas.TextMatrix(feNotas.Row, 5)
        lnItem2 = feNotas.TextMatrix(feNotas.Row, 6)
        lsCodReg2 = feNotas.TextMatrix(feNotas.Row, 7)
        lsMoneda2 = feNotas.TextMatrix(feNotas.Row, 8)
        lsMovNro2 = feNotas.TextMatrix(feNotas.Row, 9)
        lnValor2 = feNotas.TextMatrix(feNotas.Row, 10)
        
        
        feNotas.TextMatrix(feNotas.Row + 1, 1) = lnNumCol2
        feNotas.TextMatrix(feNotas.Row + 1, 2) = lsDescripcion2
        feNotas.TextMatrix(feNotas.Row + 1, 3) = lsCodigo2
        feNotas.TextMatrix(feNotas.Row + 1, 4) = lsRegimen2
        feNotas.TextMatrix(feNotas.Row + 1, 5) = lsValor2
        feNotas.TextMatrix(feNotas.Row + 1, 6) = lnItem2
        feNotas.TextMatrix(feNotas.Row + 1, 7) = lsCodReg2
        feNotas.TextMatrix(feNotas.Row + 1, 8) = lsMoneda2
        feNotas.TextMatrix(feNotas.Row + 1, 9) = lsMovNro2
        feNotas.TextMatrix(feNotas.Row + 1, 10) = lnValor2
        
        
        feNotas.TextMatrix(feNotas.Row, 1) = lnNumCol1
        feNotas.TextMatrix(feNotas.Row, 2) = lsDescripcion1
        feNotas.TextMatrix(feNotas.Row, 3) = lsCodigo1
        feNotas.TextMatrix(feNotas.Row, 4) = lsRegimen1
        feNotas.TextMatrix(feNotas.Row, 5) = lsValor1
        feNotas.TextMatrix(feNotas.Row, 6) = lnItem1
        feNotas.TextMatrix(feNotas.Row, 7) = lsCodReg1
        feNotas.TextMatrix(feNotas.Row, 8) = lsMoneda1
        feNotas.TextMatrix(feNotas.Row, 9) = lsMovNro1
        feNotas.TextMatrix(feNotas.Row, 10) = lnValor1
        

        feNotas.Row = feNotas.Row + 1
        feNotas.SetFocus
    End If

End Sub

Private Sub LimpiaControles()
    Me.txtNumCol.Text = ""
    Me.txtDescripcion.Text = ""
    Me.cboRegimen.ListIndex = -1
    Me.txtCod.Text = ""
    Me.txtFormula.Text = ""
    Me.txtTotalizado.Text = ""
    Me.Option1.value = False
    Me.Option2.value = False
    Me.Option3.value = False
End Sub
Private Sub Form_Load()
frmMdiMain.Enabled = False

CargaRep1ComboRegimen
CargaRep2ComboColumna
CargaRep3ComboTipoObl
CargaRep3ComboDestino
CargaRep3ComboTipoOblSubCol
CargaPeriodoRep3
CargaRep4ComboColumna

CentraForm Me

sInserModif = 1
nInserModifRep2Col = 1
nInserModifRep2SubCol = 1
nInserModifRep3Col = 1
nInserModifRep3SubCol = 1
nInserModifRep4Col = 1
nInserModifRep4SubCol = 1
nCol = 0
cCodigo = ""
sDescripcion = ""
ControlaAccion (99)
ControlaAccionRep2 (99)
ControlaAccionRep3 (99)
ControlaAccionRep4 (99)
CargaReportes (0)

End Sub
Private Sub CargaRep1ComboRegimen()
    Set clsRep = New DRepFormula
    Set rsRegRep = clsRep.CargarRegimenesRepoFormulaBCR()
    RSLlenaCombo rsRegRep, Me.cboRegimen
    If cboRegimen.ListCount > 0 Then
        cboRegimen.ListIndex = 0
    End If
    'Set clsRep = Nothing
End Sub

Private Sub CargaReportes(ByVal pnRep As Integer)
    Select Case pnRep
        Case 0
            CargaReporte1BCR (lsMoneda)
            CargaReporte2ColBCR
            CargaReporte2SubColBCR
            CargaReporte3ColBCR
            CargaReporte3SubColBCR
            CargaReporte4ColBCR
            CargaReporte4SubColBCR
        Case 1
            CargaReporte1BCR (lsMoneda)
        Case 2
            CargaReporte2ColBCR
        Case 3
            CargaReporte2SubColBCR
        Case 4
            CargaReporte3ColBCR
        Case 5
            CargaReporte3SubColBCR
        Case 6
            CargaReporte4ColBCR
        Case 7
            CargaReporte4SubColBCR
    End Select

End Sub
Private Sub CargaReporte1BCR(psMoneda As String)
    'Set clsRep = New DRepFormula
    Dim rsRep1 As ADODB.Recordset
    Set rsRep1 = New ADODB.Recordset
        
    Set rsRep1 = clsRep.CargaRep1FormulaBCR(CInt(lsMoneda))
    feNotas.Clear
    feNotas.FormaCabecera
    feNotas.Rows = 2
    If Not (rsRep1.EOF And rsRep1.BOF) Then
        
        Do While Not rsRep1.EOF
'            For I = 1 To rsRep1.RecordCount - 1
                feNotas.AdicionaFila
        
                feNotas.TextMatrix(feNotas.Row, 1) = rsRep1!cNumCol
                feNotas.TextMatrix(feNotas.Row, 2) = rsRep1!cDescripcion
                feNotas.TextMatrix(feNotas.Row, 3) = rsRep1!cCodigo
                feNotas.TextMatrix(feNotas.Row, 4) = rsRep1!cRegimen
                If rsRep1!nValor = 2 Then
                feNotas.TextMatrix(feNotas.Row, 5) = "Totalizado : " & rsRep1!cvalor
                Else
                feNotas.TextMatrix(feNotas.Row, 5) = rsRep1!cvalor
                End If
                feNotas.TextMatrix(feNotas.Row, 6) = rsRep1!nItem
                feNotas.TextMatrix(feNotas.Row, 7) = rsRep1!nCodRegimen
                feNotas.TextMatrix(feNotas.Row, 8) = rsRep1!nMoneda
                feNotas.TextMatrix(feNotas.Row, 9) = rsRep1!cUltimaActualizacion
                feNotas.TextMatrix(feNotas.Row, 10) = rsRep1!nValor
'                feNotas.TextMatrix(feNotas.Row, 11) = rsRep1!nNumColRef
    
'            Next I
            rsRep1.MoveNext
        Loop
        rsRep1.Close
        feNotas.Col = 5
'        feNotas.SetFocus
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
RSClose rsRep
Set clsRep = Nothing
frmMdiMain.Enabled = True
End Sub
Private Sub Option1_Click()
    Me.txtFormula.Enabled = True
    Me.txtTotalizado.Text = ""
    Me.txtTotalizado.Enabled = False
End Sub

Private Sub Option2_Click()
    Me.txtFormula.Text = ""
    Me.txtFormula.Enabled = False
    Me.txtTotalizado.Enabled = True
End Sub

Private Sub Option3_Click()
    Me.txtFormula.Enabled = False
    Me.txtFormula.Text = ""
    Me.txtTotalizado.Enabled = False
    Me.txtTotalizado.Text = ""
End Sub

Private Sub HabilitaControles(ByVal pbHabilita As Boolean)
    'GrdRep.Enabled = pbHabilita
'    cmdAgregar.Enabled = pbHabilita
'    cmdEliminar.Enabled = pbHabilita
'    cmdModificar.Enabled = pbHabilita
'    cmdAceptar.Enabled = pbHabilita
'    cmdCancelar.Enabled = pbHabilita
'    cmdImprimir.Enabled = pbHabilita

End Sub
Private Sub ControlaAccion(pnNumAccion As Integer)
    Select Case pnNumAccion
        Case 1 'Modificar datos del grid
            Me.SSTab1.TabEnabled(0) = True
            Me.SSTab1.TabEnabled(1) = False
            Me.SSTab1.TabEnabled(2) = False
            Me.SSTab1.TabEnabled(3) = False

            Me.cmdModificar.Enabled = False '1
            Me.cmdGuardar.Enabled = True '3
            Me.cmdCancelar.Enabled = True '2
            Me.cmdQuitar.Enabled = False '6
            Me.cmdEditar.Enabled = False '7
            Me.cmdSubir.Enabled = True '8
            Me.cmdBajar.Enabled = True '9
            Me.cmdAceptarA.Enabled = False '4
            Me.cmdCancelarA.Enabled = False '5
            Me.cmdSalir.Enabled = False '10
            Me.cmdNuevo.Enabled = False

'        Case 2, 3 'Cancelar y guardar cambios en el grid
'            Me.SSTab1.TabEnabled(0) = True
'            Me.SSTab1.TabEnabled(1) = True
'            Me.SSTab1.TabEnabled(2) = True
'            Me.SSTab1.TabEnabled(3) = True
'
'            Me.cmdModificar.Enabled = True '1
'            Me.cmdGuardar.Enabled = False '3
'            Me.cmdCancelar.Enabled = False '2
'            Me.cmdQuitar.Enabled = False '6
'            Me.cmdEditar.Enabled = False '7
'            Me.cmdSubir.Enabled = False '8
'            Me.cmdBajar.Enabled = False '9
'            Me.cmdAceptarA.Enabled = True '4
'            Me.cmdCancelarA.Enabled = True '5
'            Me.CmdSalir.Enabled = True '10
        
        Case 7 'editar item
            Me.SSTab1.TabEnabled(0) = True
            Me.SSTab1.TabEnabled(1) = False
            Me.SSTab1.TabEnabled(2) = False
            Me.SSTab1.TabEnabled(3) = False
            
            Me.cmdModificar.Enabled = False '1
            Me.cmdGuardar.Enabled = False '3
            Me.cmdCancelar.Enabled = False '2
            Me.cmdQuitar.Enabled = False '6
            Me.cmdEditar.Enabled = False '7
            Me.cmdSubir.Enabled = False '8
            Me.cmdBajar.Enabled = False '9
            Me.cmdAceptarA.Enabled = True '4
            Me.cmdCancelarA.Enabled = True '5
            Me.cmdSalir.Enabled = False '10
            Me.cmdNuevo.Enabled = False
            Des_HabilitarControlesPestana1 (True)
            
        Case 99, 4, 5, 3, 2 'inicio
        
            Me.SSTab1.TabEnabled(0) = True
            Me.SSTab1.TabEnabled(1) = True
            Me.SSTab1.TabEnabled(2) = True
            Me.SSTab1.TabEnabled(3) = True
        
            Me.cmdModificar.Enabled = True '1
            Me.cmdGuardar.Enabled = False '3
            Me.cmdCancelar.Enabled = False '2
            Me.cmdQuitar.Enabled = True '6
            Me.cmdEditar.Enabled = True '7
            Me.cmdSubir.Enabled = False '8
            Me.cmdBajar.Enabled = False '9
            Me.cmdSalir.Enabled = True '10
            Me.cmdNuevo.Enabled = True '11
            
            'agregado po PASI20140204 TI-ERS102-2013
            Des_HabilitarControlesPestana1 (False)
            'fin pasi
        Case 10
            Me.SSTab1.TabEnabled(0) = True
            Me.SSTab1.TabEnabled(1) = False
            Me.SSTab1.TabEnabled(2) = False
            Me.SSTab1.TabEnabled(3) = False
            Me.cmdNuevo.Enabled = False
            Me.cmdEditar.Enabled = False
            Me.cmdQuitar.Enabled = False
            Me.cmdGuardar = False
            Me.cmdCancelar.Enabled = False
            Me.cmdModificar.Enabled = False
            Me.cmdSubir.Enabled = False
            Me.cmdBajar.Enabled = False
            Des_HabilitarControlesPestana1 (True)
    End Select
End Sub
Public Sub Des_HabilitarControlesPestana1(ByVal pbHabilita As Boolean)
            Me.txtDescripcion.Enabled = pbHabilita
            Me.txtNumCol.Enabled = pbHabilita
            Me.cboRegimen.Enabled = pbHabilita
            Me.txtCod.Enabled = pbHabilita
            Me.cmdAceptarA.Enabled = pbHabilita
            Me.cmdCancelarA.Enabled = pbHabilita
            Frame13.Enabled = pbHabilita
            If Frame13.Enabled Then
                Me.txtFormula.Enabled = False
                Me.txtTotalizado.Enabled = False
            End If
End Sub

Public Sub Inicio(ByVal psOpeCod As String)
    fsOpeCod = psOpeCod
    lsMoneda = IIf(psOpeCod = "760114", "1", "2")
    Caption = "Configuración de Reportes de Encaje BCR (" & IIf(psOpeCod = "760114", "MN", "ME") & ")"
    Me.Show 1
End Sub
Private Sub txtNumCol_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
        Me.txtDescripcion.SetFocus
    End If
End Sub
Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cboRegimen.SetFocus
    End If
End Sub
Private Sub txtFormula_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumerosySimbolos(KeyAscii)
    If KeyAscii = 13 Then
        Me.cmdAceptarA.SetFocus
    End If
End Sub
Private Sub txtTotalizado_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumerosySimbolos(KeyAscii)
    If KeyAscii = 13 Then
        Me.cmdAceptarA.SetFocus
    End If
End Sub
Private Sub txtCod_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
End Sub
Function SoloNumeros(ByVal KeyAscii As Integer) As Integer
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        SoloNumeros = 0
    Else
        SoloNumeros = KeyAscii
    End If
    If KeyAscii = 8 Then SoloNumeros = KeyAscii  'borrado atras
    If KeyAscii = 13 Then SoloNumeros = KeyAscii 'Enter
End Function
Function SoloNumerosySimbolos(ByVal KeyAscii As Integer) As Integer
    If InStr("0123456789+-", Chr(KeyAscii)) = 0 Then
        SoloNumerosySimbolos = 0
    Else
        SoloNumerosySimbolos = KeyAscii
    End If
    If KeyAscii = 8 Then SoloNumerosySimbolos = KeyAscii  'borrado atras
    If KeyAscii = 13 Then SoloNumerosySimbolos = KeyAscii 'Enter
End Function
'---------------------------------------------------------------------------------Fin Pestaña 1------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------Pestaña 2------------------------------------------------------------------------------------
Private Sub CargaRep2ComboColumna()
    'Set clsRep = New DRepFormula
    Dim rsRep2CboCol As ADODB.Recordset
    Set rsRep2CboCol = New ADODB.Recordset
    Set rsRep2CboCol = clsRep.CargarComboRep2ColFormulaBCR(CInt(lsMoneda))
    RSLlenaCombo rsRep2CboCol, Me.cboColumnaRep2SubCol
    If cboColumnaRep2SubCol.ListCount > 0 Then
        cboColumnaRep2SubCol.ListIndex = 0
    End If
    'Set clsRep = Nothing
End Sub

Private Sub CargaReporte2ColBCR()
    Dim rsRep2Col As ADODB.Recordset
    Set rsRep2Col = New ADODB.Recordset
    Set rsRep2Col = clsRep.CargaRep2ColFormulaBCR(CInt(lsMoneda))
        FeRep2ColBCR.Clear
        FeRep2ColBCR.FormaCabecera
        FeRep2ColBCR.Rows = 2
    If Not (rsRep2Col.EOF And rsRep2Col.BOF) Then
        Do While Not rsRep2Col.EOF
            FeRep2ColBCR.AdicionaFila
            FeRep2ColBCR.TextMatrix(FeRep2ColBCR.Row, 1) = rsRep2Col!cCodigo
            FeRep2ColBCR.TextMatrix(FeRep2ColBCR.Row, 2) = rsRep2Col!cDescripcion
            FeRep2ColBCR.TextMatrix(FeRep2ColBCR.Row, 3) = rsRep2Col!nMoneda
            FeRep2ColBCR.TextMatrix(FeRep2ColBCR.Row, 4) = rsRep2Col!cUltimaActualizacion
            FeRep2ColBCR.TextMatrix(FeRep2ColBCR.Row, 5) = rsRep2Col!nItem
            rsRep2Col.MoveNext
        Loop
        rsRep2Col.Close
        FeRep2ColBCR.Col = 2
    End If
End Sub
Private Sub CargaReporte2SubColBCR()
     Dim rsRep2SubCol As ADODB.Recordset
     Set rsRep2SubCol = New ADODB.Recordset
     Set rsRep2SubCol = clsRep.CargarRep2SubColFormulaBCR(CInt(lsMoneda))
        FeRep2SubColBCR.Clear
        FeRep2SubColBCR.FormaCabecera
        FeRep2SubColBCR.Rows = 2
    If Not (rsRep2SubCol.EOF And rsRep2SubCol.BOF) Then
        Do While Not rsRep2SubCol.EOF
            FeRep2SubColBCR.AdicionaFila
            FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 1) = rsRep2SubCol!cCodSwif
            FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 2) = rsRep2SubCol!columna
            FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 3) = rsRep2SubCol!cDescripcion
            FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 4) = rsRep2SubCol!cvalor
            FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 5) = rsRep2SubCol!nItem
            FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 6) = rsRep2SubCol!nMoneda
            FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 7) = rsRep2SubCol!cUltimaActualizacion
            FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 8) = rsRep2SubCol!nValor
            FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 9) = rsRep2SubCol!cCodigo
            FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 10) = rsRep2SubCol!nValorRef
            rsRep2SubCol.MoveNext
        Loop
        rsRep2SubCol.Close
        FeRep2SubColBCR.Col = 4
    End If
End Sub
Private Sub cmdAceptarARep2_Click()
    Dim rsRep2CboCol As ADODB.Recordset
    Set rsRep2CboCol = New ADODB.Recordset
    If nInserModifRep2Col = 1 Then
        If ValidarRep2Col = False Then
            Exit Sub
        End If
        lcMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
        lnNumItem = IIf(IsNumeric(FeRep2ColBCR.TextMatrix(1, 1)), FeRep2ColBCR.Rows, FeRep2ColBCR.Rows - 1)
        clsRep.InsertaRep2ColFormulaBCR Trim(Left(txtCodigoRep2Col.Text, 6)), lsMoneda, Trim(Left(txtDescripcionRep2Col.Text, 50)), lnNumItem, lcMovNro
    Else
        If ValidarRep2Col = False Then
            Exit Sub
        End If
        lcMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
        clsRep.ModificaRep2ColFormulaBCR Trim(Left(txtCodigoRep2Col.Text, 6)), lsMoneda, Trim(Left(txtDescripcionRep2Col.Text, 50)), nItem, lcMovNro
        Des_HabilitaControlesRep2Col (True)
    End If
    Set rsRep2CboCol = clsRep.CargarComboRep2ColFormulaBCR(CInt(lsMoneda))
    RSLlenaCombo rsRep2CboCol, Me.cboColumnaRep2SubCol
    Set rsRep2CboCol = Nothing
    LimpiaControlesRep2Col
    CargaReportes (2)
    ControlaAccionRep2 (4)
    nInserModifRep2Col = 1
    sDescripcion = ""
    cCodigo = ""
End Sub
Private Function ValidarRep2Col() As Boolean
    If Len(Me.txtDescripcionRep2Col.Text) = 0 Then
        MsgBox "Ingrese la descripción de la columna.", vbOKOnly + vbExclamation, "Atención"
        Me.txtDescripcionRep2Col.SetFocus
        ValidarRep2Col = False
        Exit Function
    End If
    If Len(Me.txtCodigoRep2Col.Text) = 0 Then
        MsgBox "Ingrese el código de la columna.", vbOKOnly + vbExclamation, "Atención"
        Me.txtCodigoRep2Col.SetFocus
        ValidarRep2Col = False
        Exit Function
    End If
    For i = 1 To FeRep2ColBCR.Rows - 1
        If FeRep2ColBCR.TextMatrix(i, 1) = Trim(Left(txtCodigoRep2Col.Text, 6)) Then
            If FeRep2ColBCR.TextMatrix(i, 1) <> cCodigo Then
                MsgBox "El codigo de columna (" & Trim(Me.txtCodigoRep2Col.Text) & ") ya fue ingresado.", vbOKOnly + vbExclamation, "Atención"
                ValidarRep2Col = False
                Exit Function
            End If
        End If
        If FeRep2ColBCR.TextMatrix(i, 2) = Trim(Left(txtDescripcionRep2Col.Text, 50)) Then
            If FeRep2ColBCR.TextMatrix(i, 2) <> sDescripcion Then
                MsgBox "La descripción (" & Trim(Me.txtDescripcionRep2Col.Text) & ") ya fue ingresado.", vbOKOnly + vbExclamation, "Atención"
                ValidarRep2Col = False
                Exit Function
            End If
        End If
    Next i
    ValidarRep2Col = True
End Function

Private Sub cmdEditarARep2_Click()
    If IsNumeric(FeRep2ColBCR.TextMatrix(FeRep2ColBCR.Row, 1)) Then
            nInserModifRep2Col = 2
            Me.txtDescripcionRep2Col.Text = FeRep2ColBCR.TextMatrix(FeRep2ColBCR.Row, 2)
            Me.txtCodigoRep2Col.Text = FeRep2ColBCR.TextMatrix(FeRep2ColBCR.Row, 1)
            cCodigo = FeRep2ColBCR.TextMatrix(FeRep2ColBCR.Row, 1)
            sDescripcion = FeRep2ColBCR.TextMatrix(FeRep2ColBCR.Row, 2)
            nItem = FeRep2ColBCR.TextMatrix(FeRep2ColBCR.Row, 5)
            Des_HabilitaControlesRep2Col (False)
            ControlaAccionRep2 (10)
    Else
        MsgBox "No existen datos para Editar", vbOKOnly + vbExclamation, "Atención"
    End If
End Sub
Private Sub Des_HabilitaControlesRep2Col(ByVal pbHabilita As Boolean)
    Me.cmdEditarARep2.Enabled = pbHabilita
    Me.cmdQuitarARep2.Enabled = pbHabilita
    Me.SSTab1.TabEnabled(0) = pbHabilita
    Me.SSTab1.TabEnabled(2) = pbHabilita
    Me.SSTab1.TabEnabled(3) = pbHabilita
End Sub
Private Sub cmdCancelarARep2_Click()
    If nInserModifRep2Col = 1 Then
        LimpiaControlesRep2Col
    Else
        LimpiaControlesRep2Col
        Des_HabilitaControlesRep2Col (True)
        ControlaAccionRep2 (5)
        nInserModifRep2Col = 1
    End If
End Sub
Private Sub LimpiaControlesRep2Col()
    Me.txtDescripcionRep2Col.Text = ""
    Me.txtCodigoRep2Col.Text = ""
End Sub

Private Sub ControlaAccionRep2(pnNumAccion As Integer)
    Select Case pnNumAccion
        Case 1
            Me.SSTab1.TabEnabled(0) = False
            Me.SSTab1.TabEnabled(1) = True
            Me.SSTab1.TabEnabled(2) = False
            Me.SSTab1.TabEnabled(3) = False
            
            Me.cmdAceptarARep2.Enabled = False
            Me.cmdCancelarARep2.Enabled = False
            Me.cmdEditarARep2.Enabled = False
            Me.cmdQuitarARep2.Enabled = False
            Me.cmdAceptarBRep2.Enabled = False
            Me.cmdCancelarBRep2.Enabled = False
            Me.cmdNuevoRep2SubCol.Enabled = False
            Me.cmdEditarRep2SubCol.Enabled = False
            Me.cmdQuitarRep2Subcol.Enabled = False
            Me.cmdGuardarRep2Subcol.Enabled = True
            Me.cmdCancelarRep2SubCol.Enabled = True
            Me.cmdModificarRep2SubCol.Enabled = False
            Me.cmdSubirRep2SubCol.Enabled = True
            Me.cmdBajarRep2Subcol.Enabled = True
            
        Case 99, 5, 4, 3
            Me.SSTab1.TabEnabled(0) = True
            Me.SSTab1.TabEnabled(1) = True
            Me.SSTab1.TabEnabled(2) = True
            Me.SSTab1.TabEnabled(3) = True
            
            Me.cmdEditarARep2.Enabled = True
            Me.cmdQuitarARep2.Enabled = True
            Me.cmdAceptarARep2.Enabled = True
            Me.cmdCancelarARep2.Enabled = True
            Me.cmdModificarRep2SubCol.Enabled = True
            Me.cmdGuardarRep2Subcol.Enabled = False
            Me.cmdCancelarRep2SubCol.Enabled = False
            Me.cmdSubirRep2SubCol.Enabled = False
            Me.cmdBajarRep2Subcol.Enabled = False
            Me.cmdNuevoRep2SubCol.Enabled = True
            Me.cmdEditarRep2SubCol.Enabled = True
            Me.cmdQuitarRep2Subcol.Enabled = True
            Des_HabilitarControlesPestana2 (False)
        Case 10
            Me.SSTab1.TabEnabled(0) = False
            Me.SSTab1.TabEnabled(1) = True
            Me.SSTab1.TabEnabled(2) = False
            Me.SSTab1.TabEnabled(3) = False
            
            
            Me.cmdEditarARep2.Enabled = False
            Me.cmdQuitarARep2.Enabled = False
            Me.cmdAceptarBRep2.Enabled = False
            Me.cmdCancelarBRep2.Enabled = False
            Me.cmdNuevoRep2SubCol.Enabled = False
            Me.cmdEditarRep2SubCol.Enabled = False
            Me.cmdQuitarRep2Subcol.Enabled = False
            Me.cmdGuardarRep2Subcol.Enabled = False
            Me.cmdCancelarRep2SubCol.Enabled = False
            Me.cmdModificarRep2SubCol.Enabled = False
            Me.cmdSubirRep2SubCol.Enabled = False
            Me.cmdBajarRep2Subcol.Enabled = False
            Des_HabilitarControlesPestana2 (False)
        Case 7
            Me.SSTab1.TabEnabled(0) = False
            Me.SSTab1.TabEnabled(1) = True
            Me.SSTab1.TabEnabled(2) = False
            Me.SSTab1.TabEnabled(3) = False
            
            Me.cmdAceptarARep2.Enabled = False
            Me.cmdCancelarARep2.Enabled = False
            Me.cmdEditarARep2.Enabled = False
            Me.cmdQuitarARep2.Enabled = False
            Me.cmdNuevoRep2SubCol.Enabled = False
            Me.cmdEditarRep2SubCol.Enabled = False
            Me.cmdQuitarRep2Subcol.Enabled = False
            Me.cmdGuardarRep2Subcol.Enabled = False
            Me.cmdCancelarRep2SubCol.Enabled = False
            Me.cmdModificarRep2SubCol.Enabled = False
            Me.cmdSubirRep2SubCol.Enabled = False
            Me.cmdBajarRep2Subcol.Enabled = False
            Des_HabilitarControlesPestana2 (True)
    End Select
End Sub
Public Sub Des_HabilitarControlesPestana2(ByVal pbHabilita As Boolean)
    Me.txtDescripcionRep2SubCol.Enabled = pbHabilita
    Me.txtCodSwifRep2SubCol.Enabled = pbHabilita
    Me.cmdAceptarBRep2.Enabled = pbHabilita
    Me.cmdCancelarBRep2.Enabled = pbHabilita
    Me.frm_ValorRep2SubCol.Enabled = pbHabilita
    If Me.frm_ValorRep2SubCol.Enabled Then
        Me.txtForRep2SubCol.Enabled = False
    End If
End Sub
Private Sub cmdNuevoRep2SubCol_Click()
    ControlaAccionRep2 (7)
End Sub
Private Sub cmdAceptarBRep2_Click()
    If nInserModifRep2SubCol = 1 Then
        If ValidarRep2 = False Then
            Exit Sub
        End If
        lcMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
        lnNumItem = IIf(FeRep2SubColBCR.TextMatrix(1, 1) <> "", FeRep2SubColBCR.Rows, FeRep2SubColBCR.Rows - 1)
        lnValor = IIf(Me.optForRep2SubCol.value = True, 1, 2)
        clsRep.InsertaRep2SubColFormulaBCR Me.txtCodSwifRep2SubCol.Text, Trim(Right(Me.cboColumnaRep2SubCol.Text, 6)), Trim(Left(Me.txtDescripcionRep2SubCol.Text, 100)), IIf(Me.optForRep2SubCol.value, Me.txtForRep2SubCol.Text, "Totalizado"), lnNumItem, CInt(lsMoneda), lnValor, lcMovNro
    Else
        If ValidarRep2 = False Then
            Exit Sub
        End If
        lcMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
        lnValor = IIf(Me.optForRep2SubCol.value = True, 1, 2)
        clsRep.ModificaRep2SubColFormulaBCR Me.txtCodSwifRep2SubCol.Text, Trim(Right(Me.cboColumnaRep2SubCol.Text, 6)), Me.txtDescripcionRep2SubCol.Text, IIf(Me.optForRep2SubCol.value, Me.txtForRep2SubCol.Text, "Totalizado"), nItem, CInt(lsMoneda), lnValor, lcMovNro
    End If
    LimpiaControlesRep2SubCol
    CargaReportes (3)
    ControlaAccionRep2 (4)
    nInserModifRep2SubCol = 1
    sDescripcion = ""
    cCodigo = ""
End Sub
Public Function ValidarRep2() As Boolean
    If Len(Me.txtDescripcionRep2SubCol.Text) = 0 Then
        MsgBox "Ingrese una Descripcion de la Sub Columna. ", vbOKOnly + vbExclamation, "Atención"
        Me.txtDescripcionRep2SubCol.SetFocus
        ValidarRep2 = False
        Exit Function
    End If
    If Me.cboColumnaRep2SubCol.ListIndex = -1 Then
        MsgBox "Seleccione un tipo de Columna.", vbOKOnly + vbExclamation, "atención"
        Me.cboColumnaRep2SubCol.SetFocus
        ValidarRep2 = False
        Exit Function
    End If
    If Me.optForRep2SubCol.value = False And Me.optTotColRep2SubCol.value = False Then
        MsgBox "No se a seleccionado ningun valor para la SubColumna", vbOKOnly + vbExclamation, "Atención"
        Me.optForRep2SubCol.SetFocus
        ValidarRep2 = False
        Exit Function
    End If
    If Me.optForRep2SubCol.value Then
        If Len(Me.txtCodSwifRep2SubCol.Text) = 0 Then
            MsgBox "Ingrese el Código Swif. ", vbOKOnly + vbExclamation, "Atención"
            Me.txtCodSwifRep2SubCol.SetFocus
            ValidarRep2 = False
            Exit Function
        End If
        If Len(Me.txtForRep2SubCol.Text) = 0 Then
            MsgBox "No se ha ingresado el valor de la formula.", vbOKOnly + vbExclamation, "Atención"
            Me.txtForRep2SubCol.SetFocus
            ValidarRep2 = False
            Exit Function
        End If
    End If
    For i = 1 To FeRep2SubColBCR.Rows - 1
        If FeRep2SubColBCR.TextMatrix(i, 3) = Trim(Left(Me.txtDescripcionRep2SubCol.Text, 100)) And FeRep2SubColBCR.TextMatrix(i, 9) = Trim(Right(Me.cboColumnaRep2SubCol.Text, 6)) Then
            If FeRep2SubColBCR.TextMatrix(i, 3) = sDescripcion And Trim(Right(Me.cboColumnaRep2SubCol.Text, 6)) = cCodigo Then
                ValidarRep2 = True
                Exit Function
            End If
            MsgBox "La descripción de columna (" & Trim(Me.txtDescripcionRep2SubCol.Text) & ") ya fue ingresado.", vbOKOnly + vbExclamation, "Atención"
            ValidarRep2 = False
            Exit Function
        End If
    Next
    ValidarRep2 = True
End Function
Private Sub cmdCancelarBRep2_Click()
    If nInserModifRep2SubCol = 1 Then
        LimpiaControlesRep2SubCol
        ControlaAccionRep2 (5)
    Else
        LimpiaControlesRep2SubCol
        ControlaAccionRep2 (5)
        nInserModifRep2SubCol = 1
    End If
End Sub
Private Sub txtCodigoRep2Col_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
        Me.cmdAceptarARep2.SetFocus
    End If
End Sub
'Private Sub txtCodSwifRep2SubCol_KeyPress(KeyAscii As Integer)
'    KeyAscii = SoloNumeros(KeyAscii)
'    If KeyAscii = 13 Then
'        Me.cboColumnaRep2SubCol.SetFocus
'    End If
'End Sub
Private Sub txtForRep2SubCol_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumerosySimbolos(KeyAscii)
    If KeyAscii = 13 Then
        Me.cmdAceptarBRep2.SetFocus
    End If
End Sub
Private Sub LimpiaControlesRep2SubCol()
    Me.txtDescripcionRep2SubCol.Text = ""
    Me.txtCodSwifRep2SubCol.Text = ""
    Me.cboColumnaRep2SubCol.ListIndex = -1
    Me.txtForRep2SubCol.Text = ""
    Me.optForRep2SubCol.value = False
    Me.optTotColRep2SubCol.value = False
End Sub
Private Sub cmdEditarRep2SubCol_Click()
If FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 1) <> "" Then
    nInserModifRep2SubCol = 2
    ControlaAccionRep2 (7)
    Me.txtDescripcionRep2SubCol.Text = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 3)
    sDescripcion = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 3)
    cCodigo = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 9)
    Me.txtCodSwifRep2SubCol.Text = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 1)
    Me.cboColumnaRep2SubCol.ListIndex = IndiceListaCombo(cboColumnaRep2SubCol, FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 9))
    nItem = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 5)
    
    If FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 8) = 1 Then
        Me.txtForRep2SubCol.Text = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 4)
        Me.optForRep2SubCol.value = True
        optForRep2SubCol_Click
    Else
        Me.optTotColRep2SubCol.value = True
        optTotColRep2SubCol_Click
    End If
Else
    MsgBox "No hay Datos para Editar.", vbOKOnly + vbExclamation, "Atención"
End If
End Sub
Private Sub optForRep2SubCol_Click()
    Me.txtForRep2SubCol.Enabled = True
    Me.txtForRep2SubCol.SetFocus
End Sub
Private Sub optTotColRep2SubCol_Click()
    Me.txtForRep2SubCol.Text = ""
    Me.txtForRep2SubCol.Enabled = False
End Sub
Private Sub cmdQuitarRep2Subcol_Click()
Dim Y As Integer
Dim nitemv As Integer
Dim nValorRef As Integer
Y = 0
    If (FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 1) <> "") Then
        nitemv = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 5)
        If MsgBox(" ¿ Seguro que desea quitar la fila ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
            Exit Sub
        End If
        lcMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
        For i = 1 To FeRep2SubColBCR.Rows - 1
            If FeRep2SubColBCR.TextMatrix(i, 5) <> nitemv Then
                Y = Y + 1
                nValorRef = FeRep2SubColBCR.TextMatrix(i, 10)
                clsRep.ModificaRep2SubColItemOrdenFormulaBCR nValorRef, Y, CInt(lsMoneda), lcMovNro
            Else
                clsRep.EliminaRep2SubColFormulaBCR nitemv, CInt(lsMoneda)
            End If
        Next i
        CargaReportes (3)
    Else
        MsgBox "No existen datos para quitar", vbOKOnly + vbExclamation, "Atención"
    End If
End Sub
Private Sub cmdModificarRep2SubCol_Click()
    If FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 1) <> "" Then
        ControlaAccionRep2 (1)
    Else
        MsgBox "No existen datos para Modificar", vbOKOnly + vbExclamation, "Atención"
    End If
End Sub
Private Sub cmdSubirRep2SubCol_Click()
    Dim lsSubCodigo1 As String, lsSubcodigo2 As String
    Dim lsColumna1 As String, lsColumna2 As String
    Dim lsDescripcion1 As String, lsDescripcion2 As String
    Dim lsValor1 As String, lsValor2 As String
    Dim lnItem1 As Integer, lnItem2 As Integer
    Dim lnMoneda1 As Integer, lnMoneda2 As Integer
    Dim lsMovNro1 As String, lsMovNro2 As String
    Dim nValor1 As Integer, nValor2 As Integer
    Dim lsCodigo1 As String, lsCodigo2 As String
    Dim lnValorRef1 As Integer, lnValorRef2 As Integer
    
    If FeRep2SubColBCR.Row > 1 Then
        lsSubCodigo1 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row - 1, 1)
        lsColumna1 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row - 1, 2)
        lsDescripcion1 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row - 1, 3)
        lsValor1 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row - 1, 4)
        lnItem1 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row - 1, 5)
        lnMoneda1 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row - 1, 6)
        lsMovNro1 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row - 1, 7)
        nValor1 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row - 1, 8)
        lsCodigo1 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row - 1, 9)
        lnValorRef1 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row - 1, 10)
        
        lsSubcodigo2 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 1)
        lsColumna2 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 2)
        lsDescripcion2 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 3)
        lsValor2 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 4)
        lnItem2 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 5)
        lnMoneda2 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 6)
        lsMovNro2 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 7)
        nValor2 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 8)
        lsCodigo2 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 9)
        lnValorRef2 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 10)
        
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row - 1, 1) = lsSubcodigo2
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row - 1, 2) = lsColumna2
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row - 1, 3) = lsDescripcion2
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row - 1, 4) = lsValor2
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row - 1, 5) = lnItem2
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row - 1, 6) = lnMoneda2
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row - 1, 7) = lsMovNro2
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row - 1, 8) = nValor2
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row - 1, 9) = lsCodigo2
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row - 1, 10) = lnValorRef2
        
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 1) = lsSubCodigo1
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 2) = lsColumna1
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 3) = lsDescripcion1
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 4) = lsValor1
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 5) = lnItem1
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 6) = lnMoneda1
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 7) = lsMovNro1
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 8) = nValor1
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 9) = lsCodigo1
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 10) = lnValorRef1
        
        FeRep2SubColBCR.Row = FeRep2SubColBCR.Row - 1
        FeRep2SubColBCR.SetFocus
        
    End If
End Sub
Private Sub cmdCancelarRep2SubCol_Click()
    ControlaAccionRep2 (99)
    CargaReportes (3)
End Sub
Private Sub cmdBajarRep2Subcol_Click()
    Dim lsSubCodigo1 As String, lsSubcodigo2 As String
    Dim lsColumna1 As String, lsColumna2 As String
    Dim lsDescripcion1 As String, lsDescripcion2 As String
    Dim lsValor1 As String, lsValor2 As String
    Dim lnItem1 As Integer, lnItem2 As Integer
    Dim lnMoneda1 As Integer, lnMoneda2 As Integer
    Dim lsMovNro1 As String, lsMovNro2 As String
    Dim nValor1 As Integer, nValor2 As Integer
    Dim lsCodigo1 As String, lsCodigo2 As String
    Dim lnValorRef1 As Integer, lnValorRef2 As Integer
    
    If FeRep2SubColBCR.Row < FeRep2SubColBCR.Rows - 1 Then
        lsSubCodigo1 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row + 1, 1)
        lsColumna1 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row + 1, 2)
        lsDescripcion1 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row + 1, 3)
        lsValor1 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row + 1, 4)
        lnItem1 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row + 1, 5)
        lnMoneda1 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row + 1, 6)
        lsMovNro1 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row + 1, 7)
        nValor1 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row + 1, 8)
        lsCodigo1 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row + 1, 9)
        lnValorRef1 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row + 1, 10)
        
        lsSubcodigo2 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 1)
        lsColumna2 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 2)
        lsDescripcion2 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 3)
        lsValor2 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 4)
        lnItem2 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 5)
        lnMoneda2 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 6)
        lsMovNro2 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 7)
        nValor2 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 8)
        lsCodigo2 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 9)
        lnValorRef2 = FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 10)
        
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row + 1, 1) = lsSubcodigo2
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row + 1, 2) = lsColumna2
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row + 1, 3) = lsDescripcion2
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row + 1, 4) = lsValor2
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row + 1, 5) = lnItem2
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row + 1, 6) = lnMoneda2
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row + 1, 7) = lsMovNro2
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row + 1, 8) = nValor2
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row + 1, 9) = lsCodigo2
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row + 1, 10) = lnValorRef2
        
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 1) = lsSubCodigo1
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 2) = lsColumna1
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 3) = lsDescripcion1
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 4) = lsValor1
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 5) = lnItem1
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 6) = lnMoneda1
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 7) = lsMovNro1
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 8) = nValor1
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 9) = lsCodigo1
        FeRep2SubColBCR.TextMatrix(FeRep2SubColBCR.Row, 10) = lnValorRef1
        
        FeRep2SubColBCR.Row = FeRep2SubColBCR.Row + 1
        FeRep2SubColBCR.SetFocus
    End If
    
End Sub
Private Sub cmdGuardarRep2Subcol_Click()
Dim Y As Integer
Dim nValorRef As Integer
Y = 0
    If MsgBox(" ¿ Seguro de grabar el nuevo orden de las columnas ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
        Exit Sub
    End If
    lcMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
    For i = 1 To FeRep2SubColBCR.Rows - 1
        Y = Y + 1
        nValorRef = FeRep2SubColBCR.TextMatrix(i, 10)
        clsRep.ModificaRep2SubColItemOrdenFormulaBCR nValorRef, Y, CInt(lsMoneda), lcMovNro
    Next i
    
    MsgBox "Se grabó satisfactoriamente", vbOKOnly + vbInformation, "Atención"
    ControlaAccionRep2 (3)
    CargaReportes (3)
End Sub
Private Sub cmdQuitarARep2_Click()
    Dim rsRepSubCol As ADODB.Recordset
    Set rsRepSubCol = New ADODB.Recordset
    
    If (FeRep2ColBCR.TextMatrix(FeRep2ColBCR.Row, 1) <> "") Then
        If MsgBox(" ¿ Seguro que desea quitar la fila ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
            Exit Sub
        End If
        Set rsRepSubCol = clsRep.ObtenerRep2SubColdeColFormulaBCR(FeRep2ColBCR.TextMatrix(FeRep2ColBCR.Row, 1), CInt(lsMoneda))
        If rsRepSubCol!Cant > 0 Then
            MsgBox " No se puede Eliminar la Columna, primero elimine todas las subcolumnas que pertenecen a la columna.", vbOKOnly + vbExclamation, "Atención"
            Exit Sub
        End If
        clsRep.EliminaRep2ColFormulaBCR FeRep2ColBCR.TextMatrix(FeRep2ColBCR.Row, 1), CInt(lsMoneda)
        CargaReportes (2)
    Else
        MsgBox "No existen datos para quitar", vbOKOnly + vbExclamation, "Atención"
    End If
End Sub
'---------------------------------------------------------------------------------Fin Pestaña 2------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------Pestaña 3------------------------------------------------------------------------------------
Private Sub CargaRep3ComboTipoObl()
    'Set clsRep = New DRepFormula
    Dim rsRep3CboTpoObl As ADODB.Recordset
    Set rsRep3CboTpoObl = New ADODB.Recordset
    Set rsRep3CboTpoObl = clsRep.CargarTipoObligacionesRep3FormulaBCR()
    RSLlenaCombo rsRep3CboTpoObl, Me.cboTipOblRep3Col

    If cboTipOblRep3Col.ListCount > 0 Then
        cboTipOblRep3Col.ListIndex = 0
    End If
    'Set clsRep = Nothing
End Sub
Private Sub txtForRep3SubCol_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumerosySimbolos(KeyAscii)
    If KeyAscii = 13 Then
        Me.cmdAceptarA.SetFocus
    End If
End Sub
Private Sub CargaRep3ComboDestino()
    'Set clsRep = New DRepFormula
    Dim rsRep3CboDestObl As ADODB.Recordset
    Set rsRep3CboDestObl = New ADODB.Recordset
    Set rsRep3CboDestObl = clsRep.CargarDestinoObligacionRep3FormulaBCR()
    RSLlenaCombo rsRep3CboDestObl, Me.cboDestRep3SubCol
    If cboDestRep3SubCol.ListCount > 0 Then
        cboDestRep3SubCol.ListIndex = 0
    End If
    'Set clsRep = Nothing
End Sub
Private Sub CargaRep3ComboTipoOblSubCol()
    'Set clsRep = New DRepFormula
    Dim rsRep3CboTpoOblSubCol As ADODB.Recordset
    Set rsRep3CboTpoOblSubCol = New ADODB.Recordset
    Set rsRep3CboTpoOblSubCol = clsRep.CargarTipoObligacionesRep3FormulaBCR()
    RSLlenaCombo rsRep3CboTpoOblSubCol, Me.cboTipOblRep3SubCol
    
    CargaRep3ComboCol (0)
    
End Sub
Private Sub CargaRep3ComboCol(ByVal pnCodObl As Integer)
    Dim rsRep3CboCol As ADODB.Recordset
    Set rsRep3CboCol = New ADODB.Recordset
    Set rsRep3CboCol = clsRep.CargarComboRep3ColFormulaBCR(CInt(lsMoneda), pnCodObl)
    RSLlenaCombo rsRep3CboCol, Me.cboColRep3SubCol
        If cboColRep3SubCol.ListCount > 0 Then
            cboColRep3SubCol.ListIndex = 0
        End If
End Sub
Private Sub CargaReporte3ColBCR()
    Dim rsRep3Col As ADODB.Recordset
    Set rsRep3Col = New ADODB.Recordset
    Set rsRep3Col = clsRep.CargaRep3ColFormulaBCR(CInt(lsMoneda))
    FeRep3ColBCR.Clear
    FeRep3ColBCR.FormaCabecera
    FeRep3ColBCR.Rows = 2
    If Not (rsRep3Col.EOF And rsRep3Col.BOF) Then
        Do While Not rsRep3Col.EOF
            FeRep3ColBCR.AdicionaFila
            FeRep3ColBCR.TextMatrix(FeRep3ColBCR.Row, 1) = rsRep3Col!cCodigo
            FeRep3ColBCR.TextMatrix(FeRep3ColBCR.Row, 2) = rsRep3Col!cDescripcion
            FeRep3ColBCR.TextMatrix(FeRep3ColBCR.Row, 3) = rsRep3Col!TipoObl
            FeRep3ColBCR.TextMatrix(FeRep3ColBCR.Row, 4) = rsRep3Col!nItem
            FeRep3ColBCR.TextMatrix(FeRep3ColBCR.Row, 5) = rsRep3Col!nMoneda
            FeRep3ColBCR.TextMatrix(FeRep3ColBCR.Row, 6) = rsRep3Col!cUltimaActualizacion
            FeRep3ColBCR.TextMatrix(FeRep3ColBCR.Row, 7) = rsRep3Col!nCodObl
            rsRep3Col.MoveNext
        Loop
        rsRep3Col.Close
        FeRep3ColBCR.Col = 3
    End If
End Sub
Private Sub cmdAceptarARep3_Click()
    If nInserModifRep3Col = 1 Then
        If ValidarRep3Col = False Then
            Exit Sub
        End If
        lcMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
        lnNumItem = IIf(IsNumeric(FeRep3ColBCR.TextMatrix(1, 1)), FeRep3ColBCR.Rows, FeRep3ColBCR.Rows - 1)
        clsRep.InsertaRep3ColFormulaBCR Trim(Left(Me.txtCodigoRep3Col.Text, 6)), CInt(lsMoneda), Trim(Left(Me.txtDescripcionRep3Col.Text, 50)), Trim(Right(Me.cboTipOblRep3Col.Text, 1)), lnNumItem, lcMovNro
    Else
        If ValidarRep3Col = False Then
            Exit Sub
        End If
        lcMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
        clsRep.ActualizarRep3ColFormulaBCR Trim(Left(Me.txtCodigoRep3Col.Text, 6)), CInt(lsMoneda), Trim(Left(Me.txtDescripcionRep3Col.Text, 50)), Trim(Right(Me.cboTipOblRep3Col.Text, 1)), nItem, lcMovNro
    End If
    LimpiaControlesRep3Col
    CargaRep3ComboTipoOblSubCol
    CargaReportes (4)
    ControlaAccionRep3 (4)
    nInserModifRep3Col = 1
    sDescripcion = ""
    cCodigo = ""
End Sub
Private Sub LimpiaControlesRep3Col()
    Me.txtDescripcionRep3Col.Text = ""
    Me.txtCodigoRep3Col.Text = ""
    Me.cboTipOblRep3Col.ListIndex = -1
End Sub

Private Function ValidarRep3Col() As Boolean
    If Len(Me.txtDescripcionRep3Col.Text) = 0 Then
        MsgBox "Ingrese la Descripción de la columna.", vbOKOnly + vbExclamation, "Atención"
        Me.txtDescripcionRep3Col.SetFocus
        ValidarRep3Col = False
        Exit Function
    End If
    If Len(Me.txtCodigoRep3Col.Text) = 0 Then
        MsgBox "Ingrese el código de la columna.", vbOKOnly + vbExclamation, "Atención"
        Me.txtCodigoRep3Col.SetFocus
        ValidarRep3Col = False
        Exit Function
    End If
    If Me.cboTipOblRep3Col.ListIndex = -1 Then
        MsgBox "Seleccione un tipo de obligación para la columna.", vbOKOnly + vbExclamation, "atención"
        Me.cboTipOblRep3Col.SetFocus
        ValidarRep3Col = False
        Exit Function
    End If
    For i = 1 To FeRep3ColBCR.Rows - 1
        If (FeRep3ColBCR.TextMatrix(i, 1) = Trim(Left(txtCodigoRep3Col.Text, 6)) Or FeRep3ColBCR.TextMatrix(i, 2) = Trim(Left(Me.txtDescripcionRep3Col.Text, 100))) And FeRep3ColBCR.TextMatrix(i, 7) = Trim(Right(Me.cboTipOblRep3Col.Text, 1)) Then
            If FeRep3ColBCR.TextMatrix(i, 1) = cCodigo And FeRep3ColBCR.TextMatrix(i, 2) = sDescripcion And Trim(Right(Me.cboTipOblRep3Col.Text, 1)) = nCodObl Then
                ValidarRep3Col = True
                Exit Function
            End If
            If FeRep3ColBCR.TextMatrix(1, 2) = Trim(Left(Me.txtDescripcionRep3Col.Text, 100)) And FeRep3ColBCR.TextMatrix(i, 7) = Trim(Right(Me.cboTipOblRep3Col.Text, 1)) Then
                MsgBox "La descripcion de columna para este tipo obligacion ya fue registrada.", vbOKOnly + vbExclamation, "Atención"
                ValidarRep3Col = False
                Exit Function
            End If
            If FeRep3ColBCR.TextMatrix(1, 1) = Trim(Left(Me.txtCodigoRep3Col.Text, 6)) Then
                MsgBox "El código de Columna ya fue registrado.", vbOKOnly + vbExclamation, "Atención"
                Exit Function
            End If
        Else
            If FeRep3ColBCR.TextMatrix(1, 1) = Trim(Left(Me.txtCodigoRep3Col.Text, 6)) And FeRep3ColBCR.TextMatrix(i, 2) = Trim(Left(Me.txtDescripcionRep3Col.Text, 100)) And FeRep3ColBCR.TextMatrix(i, 7) <> Trim(Right(Me.cboTipOblRep3Col.Text, 1)) Then
                ValidarRep3Col = True
                Exit Function
            End If
            If FeRep3ColBCR.TextMatrix(1, 1) = Trim(Left(Me.txtCodigoRep3Col.Text, 6)) Then
                MsgBox "El código de Columna ya fue registrado.", vbOKOnly + vbExclamation, "Atención"
                ValidarRep3Col = False
                Exit Function
            End If
        End If
    Next i
    ValidarRep3Col = True
End Function
Public Sub ControlaAccionRep3(pnNumAccion As Integer)
    Select Case pnNumAccion
        Case 1
            Me.SSTab1.TabEnabled(0) = False
            Me.SSTab1.TabEnabled(1) = False
            Me.SSTab1.TabEnabled(2) = True
            Me.SSTab1.TabEnabled(3) = False
            
            Me.cmdAceptarARep3.Enabled = False
            Me.cmdCancelarARep3.Enabled = False
            Me.cmdEditarARep3.Enabled = False
            Me.cmdQuitarARep3.Enabled = False
            Me.cmdAceptarBRep3.Enabled = False
            Me.cmdCancelarBRep3.Enabled = False
            Me.cmdNuevoRep3SubCol.Enabled = False
            Me.cmdEditarRep3SubCol.Enabled = False
            Me.cmdQuitarRep3SubCol.Enabled = False
            Me.cmdGuardarRep3SubCol.Enabled = True
            Me.cmdCancelarRep3SubCol.Enabled = True
            Me.cmdModificarRep3SubCol.Enabled = False
            Me.cmdSubirRep3SubCol.Enabled = True
            Me.cmdBajarRep3SubCol.Enabled = True
            
        Case 99, 5, 4, 3
            Me.SSTab1.TabEnabled(0) = True
            Me.SSTab1.TabEnabled(1) = True
            Me.SSTab1.TabEnabled(2) = True
            Me.SSTab1.TabEnabled(3) = True
            
            Me.cmdEditarARep3.Enabled = True
            Me.cmdQuitarARep3.Enabled = True
            Me.cmdAceptarARep3.Enabled = True
            Me.cmdCancelarARep3.Enabled = True
            Me.cmdModificarRep3SubCol.Enabled = True
            Me.cmdGuardarRep3SubCol.Enabled = False
            Me.cmdCancelarRep3SubCol.Enabled = False
            Me.cmdSubirRep3SubCol.Enabled = False
            Me.cmdBajarRep3SubCol.Enabled = False
            Me.cmdNuevoRep3SubCol.Enabled = True
            Me.cmdEditarRep3SubCol.Enabled = True
            Me.cmdQuitarRep3SubCol.Enabled = True
            Des_HabilitarControlesPestana3 (False)
            Des_HabilitarControlesPestana3SubCol (False)
        Case 7
            Me.SSTab1.TabEnabled(0) = False
            Me.SSTab1.TabEnabled(1) = False
            Me.SSTab1.TabEnabled(2) = True
            Me.SSTab1.TabEnabled(3) = False
            
            Me.cmdAceptarARep3.Enabled = False
            Me.cmdCancelarARep3.Enabled = False
            Me.cmdEditarARep3.Enabled = False
            Me.cmdQuitarARep3.Enabled = False
            Me.cmdNuevoRep3SubCol.Enabled = False
            Me.cmdEditarRep3SubCol.Enabled = False
            Me.cmdQuitarRep3SubCol.Enabled = False
            Me.cmdGuardarRep3SubCol.Enabled = False
            Me.cmdCancelarRep2SubCol.Enabled = False
            Me.cmdModificarRep3SubCol.Enabled = False
            Me.cmdSubirRep3SubCol.Enabled = False
            Me.cmdBajarRep3SubCol.Enabled = False
            Des_HabilitarControlesPestana3 (True)
            Des_HabilitarControlesPestana3SubCol (False)
        Case 10
            Me.SSTab1.TabEnabled(0) = False
            Me.SSTab1.TabEnabled(1) = False
            Me.SSTab1.TabEnabled(2) = True
            Me.SSTab1.TabEnabled(3) = False
            
            Me.cmdEditarARep3.Enabled = False
            Me.cmdQuitarARep3.Enabled = False
            Me.cmdAceptarBRep3.Enabled = False
            Me.cmdCancelarBRep3.Enabled = False
            Me.cmdNuevoRep3SubCol.Enabled = False
            Me.cmdEditarRep3SubCol.Enabled = False
            Me.cmdQuitarRep3SubCol.Enabled = False
            Me.cmdGuardarRep3SubCol.Enabled = False
            Me.cmdCancelarRep3SubCol.Enabled = False
            Me.cmdModificarRep3SubCol.Enabled = False
            Me.cmdSubirRep3SubCol.Enabled = False
            Me.cmdBajarRep3SubCol.Enabled = False
            Des_HabilitarControlesPestana3 (False)
    End Select
End Sub
Private Sub Des_HabilitarControlesPestana3(ByVal pbHabilita As Boolean)
'        Me.txtInstRep3SubCol.Enabled = pbHabilita
'        Me.txtCodSwifRep3SubCol.Enabled = pbHabilita
'        Me.txtPlazPromRep3SubCol.Enabled = pbHabilita
        Me.cmdAceptarBRep3.Enabled = pbHabilita
        Me.cmdCancelarBRep3.Enabled = pbHabilita
        Me.frm_ValorRep3SubCol.Enabled = pbHabilita
        If Me.frm_ValorRep3SubCol.Enabled Then
            Me.txtForRep3SubCol.Enabled = False
           ' Me.txtTotColRep3SubCol.Enabled = False
        End If
End Sub
Private Sub cboTipOblRep3SubCol_Click()
        CargaRep3ComboCol Trim(Right(cboTipOblRep3SubCol, 1))
End Sub
Private Sub CargaReporte3SubColBCR()
    Dim rsRep3SubCol As ADODB.Recordset
    Set rsRep3SubCol = New ADODB.Recordset
    Set rsRep3SubCol = clsRep.CargarRep3SubColFormulaBCR(CInt(lsMoneda))
    FeRep3SubColBCR.Clear
    FeRep3SubColBCR.FormaCabecera
    FeRep3SubColBCR.Rows = 2
    If Not (rsRep3SubCol.EOF And rsRep3SubCol.BOF) Then
        Do While Not rsRep3SubCol.EOF
            FeRep3SubColBCR.AdicionaFila
            FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 1) = rsRep3SubCol!cDescripcion
            FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 2) = rsRep3SubCol!cCodSwif
            FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 3) = rsRep3SubCol!Destino
            FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 4) = rsRep3SubCol!dPeriodoDes
            FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 5) = rsRep3SubCol!dPeriodoHas
            FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 6) = rsRep3SubCol!nPlazoProm
            FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 7) = rsRep3SubCol!cvalor
            FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 8) = rsRep3SubCol!columna
            FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 9) = rsRep3SubCol!nItem
            FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 10) = rsRep3SubCol!nMoneda
            FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 11) = rsRep3SubCol!cUltimaActualizacion
            FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 12) = rsRep3SubCol!nValor
            FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 13) = rsRep3SubCol!CodDest
            FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 14) = rsRep3SubCol!nValorRef
            rsRep3SubCol.MoveNext
        Loop
        rsRep3SubCol.Close
        FeRep3SubColBCR.Col = 8
        End If
End Sub
Public Sub CargaPeriodoRep3()
    Me.dtpDesdeRep3SubCol.value = gdFecSis
    Me.dtpHastaRep3SubCol.value = gdFecSis
End Sub
Private Sub cmdAceptarBRep3_Click()
Dim sValor As String
    If nInserModifRep3SubCol = 1 Then
        If ValidarRep3SubCol = False Then
            Exit Sub
        End If
        lcMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
        lnNumItem = IIf(IsNumeric(FeRep3SubColBCR.TextMatrix(1, 9)), FeRep3SubColBCR.Rows, FeRep3SubColBCR.Rows - 1)
        sValor = IIf(Me.optForRep3SubCol.value, Trim(Me.txtForRep3SubCol.Text), IIf(Me.optSubTotRep3SubCol.value, "Total Columna", "0"))
        
        If optForRep3SubCol.value = True Then
            clsRep.InsertaRep3SubColformulaBCR Trim(Me.txtInstRep3SubCol.Text), Trim(Me.txtCodSwifRep3SubCol.Text), Trim(Right(Me.cboDestRep3SubCol.Text, 1)), CDate(Me.dtpDesdeRep3SubCol.value), CDate(Me.dtpHastaRep3SubCol.value), IIf(Len(Trim(Me.txtPlazPromRep3SubCol.Text)) = 0, 0#, Trim(Me.txtPlazPromRep3SubCol.Text)), Trim(Right(Me.cboColRep3SubCol, 6)), sValor, lnNumItem, CInt(lsMoneda), lcMovNro, 1, 0
        End If
        If optSubTotRep3SubCol.value = True Then
            clsRep.InsertaRep3SubColformulaBCR Trim(Me.txtInstRep3SubCol.Text), "", Trim(Right(Me.cboDestRep3SubCol.Text, 1)), CDate("01/01/1900"), CDate("01/01/1900"), 0, Trim(Right(Me.cboColRep3SubCol, 6)), sValor, lnNumItem, CInt(lsMoneda), lcMovNro, 2, 0
        End If
        If optTotColRep3SubCol.value = True Then
            clsRep.InsertaRep3SubColformulaBCR Trim(Me.txtInstRep3SubCol.Text), "", Trim(Right(Me.cboDestRep3SubCol.Text, 1)), CDate("01/01/1900"), CDate("01/01/1990"), 0, Trim(Me.txtCodOpeRep3SubCol.Text), "0", lnNumItem, CInt(lsMoneda), lcMovNro, 3, Trim(Right(Me.cboTipOblRep3SubCol, 1))
        End If
        'clsRep.InsertaRep3SubColformulaBCR Trim(Me.txtInstRep3SubCol.Text), Trim(Me.txtCodSwifRep3SubCol.Text), Trim(Right(Me.cboDestRep3SubCol.Text, 1)), CDate(Me.dtpDesdeRep3SubCol.value), CDate(Me.dtpHastaRep3SubCol.value), IIf(Len(Trim(Me.txtPlazPromRep3SubCol.Text)) = 0, 0#, Trim(Me.txtPlazPromRep3SubCol.Text)), Trim(Right(Me.cboColRep3SubCol, 6)), sValor, lnNumItem, CInt(lsMoneda), lcMovNro, IIf(Me.optForRep3SubCol.value, 1, IIf(Me.optSubTotRep3SubCol.value, 2, 3))
    Else
        If ValidarRep3SubCol = False Then
            Exit Sub
        End If
        lcMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
        sValor = IIf(Me.optForRep3SubCol.value, Trim(Me.txtForRep3SubCol.Text), IIf(Me.optSubTotRep3SubCol.value, "Total Columna", "0"))
         If optForRep3SubCol.value = True Then
            clsRep.ModificarRep3SubColFormulaBCR Trim(Me.txtInstRep3SubCol.Text), Trim(Me.txtCodSwifRep3SubCol.Text), Trim(Right(Me.cboDestRep3SubCol.Text, 1)), CDate(Me.dtpDesdeRep3SubCol.value), CDate(Me.dtpHastaRep3SubCol.value), IIf(Len(Trim(Me.txtPlazPromRep3SubCol.Text)) = 0, 0#, Trim(Me.txtPlazPromRep3SubCol.Text)), Trim(Right(Me.cboColRep3SubCol, 6)), sValor, nItem, CInt(lsMoneda), lcMovNro, 1, 0
        End If
        If optSubTotRep3SubCol.value = True Then
            clsRep.ModificarRep3SubColFormulaBCR Trim(Me.txtInstRep3SubCol.Text), "", Trim(Right(Me.cboDestRep3SubCol.Text, 1)), CDate("01/01/1900"), CDate("01/01/1900"), 0, Trim(Right(Me.cboColRep3SubCol, 6)), sValor, nItem, CInt(lsMoneda), lcMovNro, 2, 0
        End If
        If optTotColRep3SubCol.value = True Then
            clsRep.ModificarRep3SubColFormulaBCR Trim(Me.txtInstRep3SubCol.Text), "", Trim(Right(Me.cboDestRep3SubCol.Text, 1)), CDate("01/01/1900"), CDate("01/01/1990"), 0, Trim(Me.txtCodOpeRep3SubCol.Text), "0", nItem, CInt(lsMoneda), lcMovNro, 3, Trim(Right(Me.cboTipOblRep3SubCol, 1))
        End If
        'clsRep.ModificarRep3SubColFormulaBCR Trim(Me.txtInstRep3SubCol.Text), Trim(Me.txtCodSwifRep3SubCol.Text), Trim(Right(Me.cboDestRep3SubCol.Text, 1)), CDate(Me.dtpDesdeRep3SubCol.value), CDate(Me.dtpHastaRep3SubCol.value), IIf(Len(Trim(Me.txtPlazPromRep3SubCol.Text)) = 0, 0#, Trim(Me.txtPlazPromRep3SubCol.Text)), Trim(Right(Me.cboColRep3SubCol, 6)), sValor, nItem, CInt(lsMoneda), lcMovNro, IIf(Me.optForRep3SubCol.value, 1, IIf(Me.optSubTotRep3SubCol.value, 2, 3))
    End If
    LimpiaControlesRep3SubCol
    CargaReportes (5)
    ControlaAccionRep3 (4)
    nInserModifRep3SubCol = 1
    sDescripcion = ""
    cCodigo = ""
End Sub
Private Function ValidarRep3SubCol() As Boolean
    If optForRep3SubCol.value = False And optSubTotRep3SubCol.value = False And optTotColRep3SubCol.value = False Then
        MsgBox "No se ha seleccionado ningun valor para la Sub Columna", vbOKOnly + vbExclamation, "Atención"
        ValidarRep3SubCol = False
        Exit Function
    End If
    If optForRep3SubCol.value = True Then
        If Len(Trim(txtForRep3SubCol.Text)) = 0 Then
            MsgBox "No se ha ingresado el valor de la Formula.", vbOKOnly + vbExclamation, "Atención"
            Me.txtForRep3SubCol.SetFocus
            ValidarRep3SubCol = False
            Exit Function
        End If
        If Len(Trim(txtInstRep3SubCol.Text)) = 0 Then
            MsgBox "No se ha ingresado la descripción de la Sub Columna.", vbOKOnly + vbExclamation, "Atención"
            Me.txtInstRep3SubCol.SetFocus
            ValidarRep3SubCol = False
            Exit Function
        End If
        If Len(Trim(txtCodSwifRep3SubCol.Text)) = 0 Then
            MsgBox "No se ha ingresado el codigo Swif.", vbOKOnly + vbExclamation, "Atención"
            Me.txtCodSwifRep3SubCol.SetFocus
            ValidarRep3SubCol = False
            Exit Function
        End If
        If cboDestRep3SubCol.ListIndex = -1 Then
            MsgBox "No se ha seleccionado el Destino.", vbOKOnly + vbExclamation, "Atención"
            Me.cboDestRep3SubCol.SetFocus
            ValidarRep3SubCol = False
            Exit Function
        End If
        If Len(Trim(txtPlazPromRep3SubCol.Text)) = 0 Then
            MsgBox "No se ha ingresado el Plazo Promedio.", vbOKOnly + vbExclamation, "Atención"
            Me.txtPlazPromRep3SubCol.SetFocus
            ValidarRep3SubCol = False
            Exit Function
        End If
        If CDate(dtpDesdeRep3SubCol.value) > CDate(dtpHastaRep3SubCol) Or CDate(dtpDesdeRep3SubCol.value) = CDate(dtpHastaRep3SubCol) Then
            MsgBox "La Fecha Hasta debe ser mayor a la fecha Desde.", vbOKOnly + vbExclamation, "Atención"
            ValidarRep3SubCol = False
            Exit Function
        End If
        If cboTipOblRep3SubCol.ListIndex = -1 Then
            MsgBox "No se ha seleccionado el tipo de obligción.", vbOKOnly + vbExclamation, "Atención"
            Me.cboTipOblRep3SubCol.SetFocus
            ValidarRep3SubCol = False
            Exit Function
        End If
        If cboColRep3SubCol.ListIndex = -1 Then
            MsgBox "No se ha seleccionado la columna.", vbOKOnly + vbExclamation, "Atención"
            Me.cboColRep3SubCol.SetFocus
            ValidarRep3SubCol = False
            Exit Function
        End If
    End If
    'Validar cuando se de el check  sub total
    If optSubTotRep3SubCol.value = True Then
        If Len(Trim(txtInstRep3SubCol.Text)) = 0 Then
            MsgBox "No se ha ingresado una descripción para la subcolumna", vbOKOnly + vbExclamation, "Atención"
            Me.txtInstRep3SubCol.SetFocus
            ValidarRep3SubCol = False
            Exit Function
        End If
        If cboTipOblRep3SubCol.ListIndex = -1 Then
            MsgBox "No se ha seleccionado el tipo de obligación.", vbOKOnly + vbExclamation, "Atención"
            Me.cboTipOblRep3SubCol.SetFocus
            ValidarRep3SubCol = False
            Exit Function
        End If
        If cboColRep3SubCol.ListIndex = -1 Then
            MsgBox "No se ha seleccionado la columna.", vbOKOnly + vbExclamation, "Atención"
            Me.cboColRep3SubCol.SetFocus
            ValidarRep3SubCol = False
            Exit Function
        End If
        If cboDestRep3SubCol.ListIndex = -1 Then
            MsgBox "No se ha seleccionado el Destino.", vbOKOnly + vbExclamation, "Atención"
            Me.cboDestRep3SubCol.SetFocus
            ValidarRep3SubCol = False
            Exit Function
        End If
    End If
     'Validar cuando se de el check a totalizar
    If optTotColRep3SubCol.value = True Then
'        If Len(Trim(txtTotColRep3SubCol.Text)) = 0 Then
'            MsgBox "No se ha ingresado el valor para totalizar columnas.", vbOKOnly + vbExclamation, "Atención"
'            Me.txtTotColRep3SubCol.SetFocus
'            ValidarRep3SubCol = False
'            Exit Function
'        End If
        If Len(Trim(txtCodOpeRep3SubCol.Text)) = 0 Then
            MsgBox "No se ha ingresado el valor para el codigo de operación.", vbOKOnly + vbExclamation, "Atención"
            Me.txtCodOpeRep3SubCol.SetFocus
            ValidarRep3SubCol = False
            Exit Function
        End If
        If cboTipOblRep3SubCol.ListIndex = -1 Then
            MsgBox "No se ha seleccionado ningun tipo de obligación.", vbOKOnly + vbExclamation, "Atención"
            Me.cboTipOblRep3SubCol.ListIndex = -1
            ValidarRep3SubCol = False
            Exit Function
        End If
        If cboDestRep3SubCol.ListIndex = -1 Then
            MsgBox "No se ha seleccionado el Destino.", vbOKOnly + vbExclamation, "Atención"
            Me.cboDestRep3SubCol.SetFocus
            ValidarRep3SubCol = False
            Exit Function
        End If
    End If
    ValidarRep3SubCol = True
End Function
Private Sub cmdNuevoRep3SubCol_Click()
    ControlaAccionRep3 (7)
End Sub
Private Sub LimpiaControlesRep3SubCol()
    Me.txtInstRep3SubCol.Text = ""
    Me.txtCodSwifRep3SubCol.Text = ""
    Me.txtPlazPromRep3SubCol.Text = ""
    Me.cboDestRep3SubCol.ListIndex = -1
    Me.optForRep3SubCol.value = False
    Me.txtForRep3SubCol.Text = ""
    Me.optSubTotRep3SubCol.value = False
    Me.optTotColRep3SubCol.value = False
    'Me.txtTotColRep3SubCol.Text = ""
    CargaRep3ComboTipoOblSubCol
    CargaPeriodoRep3
End Sub
Private Sub cmdCancelarBRep3_Click()
    If nInserModifRep3SubCol = 1 Then
        LimpiaControlesRep3SubCol
    Else
        LimpiaControlesRep3SubCol
        nInserModifRep3SubCol = 1
    End If
    ControlaAccionRep3 (5)
End Sub
Private Sub optForRep3SubCol_Click()
    Me.txtForRep3SubCol.Enabled = True
    'Me.txtTotColRep3SubCol.Text = ""
    'Me.txtTotColRep3SubCol.Enabled = False
    Me.txtCodOpeRep3SubCol.Text = ""
    Me.txtCodOpeRep3SubCol.Enabled = False
    Des_HabilitarControlesPestana3SubCol (True)
End Sub
Private Sub optSubTotRep3SubCol_Click()
    Me.txtForRep3SubCol.Text = ""
    Me.txtForRep3SubCol.Enabled = False
    'Me.txtTotColRep3SubCol.Text = ""
    'Me.txtTotColRep3SubCol.Enabled = False
    Me.txtCodOpeRep3SubCol.Text = ""
    Me.txtCodOpeRep3SubCol.Enabled = False
    Des_HabilitarControlesPestana3SubCol (False)
    Me.cboDestRep3SubCol.Enabled = True
    Me.cboTipOblRep3SubCol.Enabled = True
    Me.cboColRep3SubCol.Enabled = True
    Me.txtInstRep3SubCol.Enabled = True
End Sub
Private Sub optTotColRep3SubCol_Click()
    Me.txtForRep3SubCol.Text = ""
    Me.txtForRep3SubCol.Enabled = False
    Me.txtCodOpeRep3SubCol.Enabled = True
    'Me.txtTotColRep3SubCol.Enabled = True
    Des_HabilitarControlesPestana3SubCol (False)
    Me.cboDestRep3SubCol.Enabled = True
    Me.txtInstRep3SubCol.Enabled = True
    Me.cboTipOblRep3SubCol.Enabled = True
    Me.cboTipOblRep3SubCol.ListIndex = 1
End Sub
Private Sub cmdEditarRep3SubCol_Click()
    If FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 1) <> "" Then
        nInserModifRep3SubCol = 2
        ControlaAccionRep3 (7)
        Me.txtInstRep3SubCol.Text = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 1)
        Me.txtCodSwifRep3SubCol.Text = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 2)
        Me.cboDestRep3SubCol.ListIndex = IndiceListaCombo(cboDestRep3SubCol, FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 13))
        Me.dtpDesdeRep3SubCol.value = CDate(FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 4))
        Me.dtpHastaRep3SubCol.value = CDate(FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 5))
        Me.txtPlazPromRep3SubCol.Text = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 6)
        nItem = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 9)
        
        If FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 12) = 1 Then
            Me.optForRep3SubCol.value = True
            Me.txtForRep3SubCol.Text = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 7)
            optForRep3SubCol_Click
            Me.optSubTotRep3SubCol.value = False
            Me.optTotColRep3SubCol.value = False
        ElseIf FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 12) = 2 Then
            Me.optForRep3SubCol.value = False
            Me.optSubTotRep3SubCol.value = True
            Me.optTotColRep3SubCol.value = False
            optSubTotRep3SubCol_Click
        Else
            Me.optTotColRep3SubCol.value = True
            'Me.txtTotColRep3SubCol.Text = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 7)
            Me.optForRep3SubCol.value = False
            Me.optSubTotRep3SubCol.value = False
            optTotColRep3SubCol_Click
        End If
    Else
        MsgBox "No hay Datos para Editar.", vbOKOnly + vbExclamation, "Atención"
    End If
End Sub
Private Sub Des_HabilitarControlesPestana3SubCol(ByVal pnAccion As Boolean)
    Me.txtInstRep3SubCol.Enabled = pnAccion
    txtCodSwifRep3SubCol.Enabled = pnAccion
    Me.cboDestRep3SubCol.Enabled = pnAccion
    Me.dtpDesdeRep3SubCol.Enabled = pnAccion
    Me.dtpHastaRep3SubCol.Enabled = pnAccion
    Me.txtPlazPromRep3SubCol.Enabled = pnAccion
    Me.cboTipOblRep3SubCol.Enabled = pnAccion
    Me.cboColRep3SubCol.Enabled = pnAccion
'    Me.cmdAceptarBRep3.Enabled = pnAccion
'    Me.cmdCancelarBRep3.Enabled = pnAccion
End Sub
Private Sub cmdEditarARep3_Click()
    If (IsNumeric(FeRep3ColBCR.TextMatrix(FeRep3ColBCR.Row, 1))) Then
        nInserModifRep3Col = 2
        Me.txtDescripcionRep3Col.Text = FeRep3ColBCR.TextMatrix(FeRep3ColBCR.Row, 2)
        Me.txtCodigoRep3Col.Text = FeRep3ColBCR.TextMatrix(FeRep3ColBCR.Row, 1)
        Me.cboTipOblRep3Col.ListIndex = IndiceListaCombo(cboTipOblRep3Col, FeRep3ColBCR.TextMatrix(FeRep3ColBCR.Row, 7))
        sDescripcion = FeRep3ColBCR.TextMatrix(FeRep3ColBCR.Row, 2)
        cCodigo = FeRep3ColBCR.TextMatrix(FeRep3ColBCR.Row, 1)
        nCodObl = FeRep3ColBCR.TextMatrix(FeRep3ColBCR.Row, 7)
        nItem = FeRep3ColBCR.TextMatrix(FeRep3ColBCR.Row, 4)
        Des_HabilitaControlesRep3Col (False)
        ControlaAccionRep3 (10)
    Else
        MsgBox "No existen datos para Editar", vbOKOnly + vbExclamation, "Atención"
    End If
End Sub
Private Sub Des_HabilitaControlesRep3Col(ByVal pbHabilita As Boolean)
    Me.cmdEditarARep3.Enabled = pbHabilita
    Me.cmdQuitarARep3.Enabled = pbHabilita
    Me.SSTab1.TabEnabled(0) = pbHabilita
    Me.SSTab1.TabEnabled(1) = pbHabilita
    Me.SSTab1.TabEnabled(3) = pbHabilita
End Sub
Private Sub cmdCancelarARep3_Click()
    If nInserModifRep3Col = 1 Then
        LimpiaControlesRep3Col
        ControlaAccionRep3 (5)
    Else
        LimpiaControlesRep3Col
        Des_HabilitaControlesRep3Col (True)
        ControlaAccionRep3 (5)
        nInserModifRep3Col = 1
    End If
End Sub
Private Sub cmdQuitarARep3_Click()
    Dim rsRepSubCol As ADODB.Recordset
    Set rsRepSubCol = New ADODB.Recordset
    
    If IsNumeric(FeRep3ColBCR.TextMatrix(FeRep3ColBCR.Row, 1)) Then
        If MsgBox(" ¿ Seguro que desea quitar la fila ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
            Exit Sub
        End If
        Set rsRepSubCol = clsRep.ObtenerRep3SubColdeColFormulaBCR(FeRep3ColBCR.TextMatrix(FeRep3ColBCR.Row, 1), CInt(lsMoneda))
        If rsRepSubCol!Cant > 0 Then
            MsgBox " No se puede Eliminar la Columna, primero elimine todas las subcolumnas que pertenecen a la columna.", vbOKOnly + vbExclamation, "Atención"
            Exit Sub
        End If
        clsRep.EliminaRep3ColFormulaBCR FeRep3ColBCR.TextMatrix(FeRep3ColBCR.Row, 1), CInt(lsMoneda)
        CargaReportes (4)
    Else
        MsgBox "No existen datos para quitar", vbOKOnly + vbExclamation, "Atención"
    End If
    
End Sub
Private Sub cmdQuitarRep3SubCol_Click()
    Dim Y As Integer
    Dim nitemv As Integer
    Dim nValorRef As Integer
    Y = 0
    If (FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 1) <> "") Then
        nitemv = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 9)
        If MsgBox(" ¿ Seguro que desea quitar la fila ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
            Exit Sub
        End If
        lcMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
        For i = 1 To FeRep3SubColBCR.Rows - 1
            If FeRep3SubColBCR.TextMatrix(i, 9) <> nitemv Then
                Y = Y + 1
                nValorRef = FeRep3SubColBCR.TextMatrix(i, 14)
                clsRep.ModificaRep3SubColItemOrdenFormulaBCR nValorRef, Y, CInt(lsMoneda), lcMovNro
            Else
                clsRep.EliminaRep3SubColFormulaBCR nitemv, CInt(lsMoneda)
            End If
        Next i
        CargaReportes (5)
    Else
        MsgBox "No existen datos para quitar", vbOKOnly + vbExclamation, "Atención"
    End If
End Sub
Private Sub cmdModificarRep3SubCol_Click()
    If FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 1) <> "" Then
        ControlaAccionRep3 (1)
    Else
        MsgBox "No existen datos para Modificar", vbOKOnly + vbExclamation, "Atención"
    End If
End Sub
Private Sub cmdGuardarRep3SubCol_Click()
Dim Y As Integer
Dim nValorRef As Integer
Y = 0
    If MsgBox(" ¿ Seguro de grabar el nuevo orden de las columnas ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
        Exit Sub
    End If
    lcMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
    For i = 1 To FeRep3SubColBCR.Rows - 1
        Y = Y + 1
        nValorRef = FeRep3SubColBCR.TextMatrix(i, 14)
        clsRep.ModificaRep3SubColItemOrdenFormulaBCR nValorRef, Y, CInt(lsMoneda), lcMovNro
    Next i
    MsgBox "Se grabó satisfactoriamente", vbOKOnly + vbInformation, "Atención"
    ControlaAccionRep3 (3)
    CargaReportes (5)
End Sub
Private Sub cmdCancelarRep3SubCol_Click()
    ControlaAccionRep3 (99)
    CargaReportes (5)
End Sub
Private Sub cmdSubirRep3SubCol_Click()
    Dim lsInst1, lsInst2 As String
    Dim lsCodSw1, lsCodSw2 As String
    Dim lsDest1, lsDest2 As String
    Dim lsDesde1, lsDesde2 As String
    Dim lsHasta1, lsHasta2 As String
    Dim lnPlazo1, lnPlazo2 As Double
    Dim lsValor1, lsValor2 As String
    Dim lsColumna1, lsColumna2 As String
    Dim lnItem1, lnItem2 As Integer
    Dim lnMoneda1, lnMoneda2 As Integer
    Dim lsMovNro1, lsMovNro2 As String
    Dim lnValor1, lnValor2 As Integer
    Dim lnCodDest1, lnCodDest2 As Integer
    Dim lnValorRef1, lnValorRef2 As Integer
    
    If FeRep3SubColBCR.Row > 1 Then
        lsInst1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 1)
        lsCodSw1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 2)
        lsDest1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 3)
        lsDesde1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 4)
        lsHasta1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 5)
        lnPlazo1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 6)
        lsValor1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 7)
        lsColumna1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 8)
        lnItem1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 9)
        lnMoneda1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 10)
        lsMovNro1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 11)
        lnValor1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 12)
        lnCodDest1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 13)
        lnValorRef1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 14)
        
        lsInst2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 1)
        lsCodSw2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 2)
        lsDest2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 3)
        lsDesde2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 4)
        lsHasta2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 5)
        lnPlazo2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 6)
        lsValor2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 7)
        lsColumna2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 8)
        lnItem2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 9)
        lnMoneda2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 10)
        lsMovNro2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 11)
        lnValor2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 12)
        lnCodDest2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 13)
        lnValorRef2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 14)
        
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 1) = lsInst2
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 2) = lsCodSw2
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 3) = lsDest2
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 4) = lsDesde2
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 5) = lsHasta2
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 6) = lnPlazo2
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 7) = lsValor2
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 8) = lsColumna2
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 9) = lnItem2
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 10) = lnMoneda2
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 11) = lsMovNro2
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 12) = lnValor2
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 13) = lnCodDest2
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row - 1, 14) = lnValorRef2
        
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 1) = lsInst1
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 2) = lsCodSw1
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 3) = lsDest1
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 4) = lsDesde1
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 5) = lsHasta1
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 6) = lnPlazo1
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 7) = lsValor1
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 8) = lsColumna1
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 9) = lnItem1
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 10) = lnMoneda1
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 11) = lsMovNro1
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 12) = lnValor1
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 13) = lnCodDest1
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 14) = lnValorRef1
        
        FeRep3SubColBCR.Row = FeRep3SubColBCR.Row - 1
        FeRep3SubColBCR.SetFocus
    End If
End Sub
Private Sub cmdBajarRep3SubCol_Click()
    Dim lsInst1, lsInst2 As String
    Dim lsCodSw1, lsCodSw2 As String
    Dim lsDest1, lsDest2 As String
    Dim lsDesde1, lsDesde2 As String
    Dim lsHasta1, lsHasta2 As String
    Dim lnPlazo1, lnPlazo2 As Double
    Dim lsValor1, lsValor2 As String
    Dim lsColumna1, lsColumna2 As String
    Dim lnItem1, lnItem2 As Integer
    Dim lnMoneda1, lnMoneda2 As Integer
    Dim lsMovNro1, lsMovNro2 As String
    Dim lnValor1, lnValor2 As Integer
    Dim lnCodDest1, lnCodDest2 As Integer
    Dim lnValorRef1, lnValorRef2 As Integer
    
    If FeRep3SubColBCR.Row < FeRep3SubColBCR.Rows - 1 Then
        lsInst1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 1)
        lsCodSw1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 2)
        lsDest1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 3)
        lsDesde1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 4)
        lsHasta1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 5)
        lnPlazo1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 6)
        lsValor1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 7)
        lsColumna1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 8)
        lnItem1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 9)
        lnMoneda1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 10)
        lsMovNro1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 11)
        lnValor1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 12)
        lnCodDest1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 13)
        lnValorRef1 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 14)
        
        lsInst2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 1)
        lsCodSw2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 2)
        lsDest2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 3)
        lsDesde2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 4)
        lsHasta2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 5)
        lnPlazo2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 6)
        lsValor2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 7)
        lsColumna2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 8)
        lnItem2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 9)
        lnMoneda2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 10)
        lsMovNro2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 11)
        lnValor2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 12)
        lnCodDest2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 13)
        lnValorRef2 = FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 14)
        
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 1) = lsInst2
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 2) = lsCodSw2
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 3) = lsDest2
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 4) = lsDesde2
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 5) = lsHasta2
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 6) = lnPlazo2
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 7) = lsValor2
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 8) = lsColumna2
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 9) = lnItem2
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 10) = lnMoneda2
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 11) = lsMovNro2
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 12) = lnValor2
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 13) = lnCodDest2
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row + 1, 14) = lnValorRef2
        
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 1) = lsInst1
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 2) = lsCodSw1
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 3) = lsDest1
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 4) = lsDesde1
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 5) = lsHasta1
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 6) = lnPlazo1
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 7) = lsValor1
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 8) = lsColumna1
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 9) = lnItem1
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 10) = lnMoneda1
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 11) = lsMovNro1
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 12) = lnValor1
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 13) = lnCodDest1
        FeRep3SubColBCR.TextMatrix(FeRep3SubColBCR.Row, 14) = lnValorRef1
        
        FeRep3SubColBCR.Row = FeRep3SubColBCR.Row + 1
        FeRep3SubColBCR.SetFocus
        
    End If
End Sub
'---------------------------------------------------------------------------------Fin Pestaña 3------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------Pestaña 4------------------------------------------------------------------------------------
Private Sub CargaReporte4ColBCR()
    Dim rsRep4Col As ADODB.Recordset
    Set rsRep4Col = New ADODB.Recordset
    Set rsRep4Col = clsRep.CargaRep4ColFormulaBCR(CInt(lsMoneda))
        FeRep4ColBCR.Clear
        FeRep4ColBCR.FormaCabecera
        FeRep4ColBCR.Rows = 2
    If Not (rsRep4Col.EOF And rsRep4Col.BOF) Then
        Do While Not rsRep4Col.EOF
            FeRep4ColBCR.AdicionaFila
            FeRep4ColBCR.TextMatrix(FeRep4ColBCR.Row, 1) = rsRep4Col!cCodigo
            FeRep4ColBCR.TextMatrix(FeRep4ColBCR.Row, 2) = rsRep4Col!cDescripcion
            FeRep4ColBCR.TextMatrix(FeRep4ColBCR.Row, 3) = rsRep4Col!nMoneda
            FeRep4ColBCR.TextMatrix(FeRep4ColBCR.Row, 4) = rsRep4Col!cUltimaActualizacion
            FeRep4ColBCR.TextMatrix(FeRep4ColBCR.Row, 5) = rsRep4Col!nItem
            rsRep4Col.MoveNext
        Loop
        rsRep4Col.Close
        FeRep4ColBCR.Col = 2
    End If
End Sub
Private Sub cmdAceptarARep4_Click()
    If nInserModifRep4Col = 1 Then
        If ValidarRep4Col = False Then
            Exit Sub
        End If
         lcMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
         lnNumItem = IIf(FeRep2ColBCR.TextMatrix(1, 1) <> "", FeRep2ColBCR.Rows, FeRep2ColBCR.Rows - 1)
         clsRep.InsertaRep4ColFormulaBCR Trim(Left(txtCodigoRep4Col.Text, 6)), lsMoneda, Trim(Left(txtDescripcionRep4Col.Text, 100)), lnNumItem, lcMovNro
    Else
        If ValidarRep4Col = False Then
            Exit Sub
        End If
        lcMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
        clsRep.ModificaRep4ColFormulaBCR Trim(Left(txtCodigoRep4Col.Text, 6)), lsMoneda, Trim(Left(txtDescripcionRep4Col.Text, 100)), nItem, lcMovNro
        Des_HabilitaControlesRep4Col (True)
    End If
    LimpiaControlesRep4Col
    CargaReportes (6)
    ControlaAccionRep2 (4)
    CargaRep4ComboColumna
    nInserModifRep4Col = 1
    sDescripcion = ""
    cCodigo = ""
End Sub
Private Function ValidarRep4Col() As Boolean
    If Len(Me.txtDescripcionRep4Col.Text) = 0 Then
        MsgBox "Ingrese la descripción de la columna.", vbOKOnly + vbExclamation, "Atención"
        Me.txtDescripcionRep4Col.SetFocus
        ValidarRep4Col = False
        Exit Function
    End If
    If Len(Me.txtCodigoRep4Col.Text) = 0 Then
        MsgBox "Ingrese el código de la columna.", vbOKOnly + vbExclamation, "Atención"
        Me.txtCodigoRep4Col.SetFocus
        ValidarRep4Col = False
        Exit Function
    End If
    For i = 1 To FeRep4ColBCR.Rows - 1
        If FeRep4ColBCR.TextMatrix(i, 1) = Trim(Left(txtCodigoRep4Col.Text, 6)) Then
            If FeRep4ColBCR.TextMatrix(i, 1) <> cCodigo Then
                MsgBox "El codigo de columna (" & Trim(Me.txtCodigoRep4Col.Text) & ") ya fue ingresado.", vbOKOnly + vbExclamation, "Atención"
                ValidarRep4Col = False
                Exit Function
            End If
        End If
        If FeRep4ColBCR.TextMatrix(i, 2) = Trim(Left(txtDescripcionRep4Col.Text, 100)) Then
            If FeRep4ColBCR.TextMatrix(i, 2) <> sDescripcion Then
                MsgBox "La descripción (" & Trim(Me.txtDescripcionRep4Col.Text) & ") ya fue ingresado.", vbOKOnly + vbExclamation, "Atención"
                ValidarRep4Col = False
                Exit Function
            End If
        End If
    Next i
    ValidarRep4Col = True
End Function
Private Sub LimpiaControlesRep4Col()
    Me.txtDescripcionRep4Col.Text = ""
    Me.txtCodigoRep4Col.Text = ""
End Sub
Private Sub CargaRep4ComboColumna()
    Dim rsRep4CboCol As ADODB.Recordset
    Set rsRep4CboCol = New ADODB.Recordset
    Set rsRep4CboCol = clsRep.CargarComboRep4ColFormulaBCR(CInt(lsMoneda))
    RSLlenaCombo rsRep4CboCol, Me.cboColumnaRep4SubCol
    If cboColumnaRep4SubCol.ListCount > 0 Then
        cboColumnaRep4SubCol.ListIndex = 0
    End If
End Sub
Private Sub CargaReporte4SubColBCR()
    Dim rsRep4SubCol As ADODB.Recordset
    Set rsRep4SubCol = New ADODB.Recordset
    Set rsRep4SubCol = clsRep.CargarRep4SubColFormulaBCR(CInt(lsMoneda))
    FeRep4SubColBCR.Clear
    FeRep4SubColBCR.FormaCabecera
    FeRep4SubColBCR.Rows = 2
    If Not (rsRep4SubCol.EOF And rsRep4SubCol.BOF) Then
        Do While Not rsRep4SubCol.EOF
            FeRep4SubColBCR.AdicionaFila
            FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 1) = rsRep4SubCol!cDescripcion
            FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 2) = rsRep4SubCol!cCodSwif
            FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 3) = rsRep4SubCol!columna
            FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 4) = rsRep4SubCol!dPeriodoDes
            FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 5) = rsRep4SubCol!dPeriodoHas
            FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 6) = rsRep4SubCol!nPlazoProm
            FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 7) = rsRep4SubCol!cvalor
            FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 8) = rsRep4SubCol!nItem
            FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 9) = rsRep4SubCol!cUltimaActualizacion
            FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 10) = rsRep4SubCol!nValorRef
            FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 11) = rsRep4SubCol!nValor
            FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 12) = rsRep4SubCol!cCodigo
            FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 13) = rsRep4SubCol!bAplicaPer
            
            rsRep4SubCol.MoveNext
        Loop
        rsRep4SubCol.Close
        FeRep4SubColBCR.Col = 7
    End If
End Sub
Private Sub ControlaAccionRep4(pnNumAccion As Integer)
    Select Case pnNumAccion
        Case 1
            Me.SSTab1.TabEnabled(0) = False
            Me.SSTab1.TabEnabled(1) = False
            Me.SSTab1.TabEnabled(2) = False
            Me.SSTab1.TabEnabled(3) = True
            
            Me.cmdAceptarARep4.Enabled = False
            Me.cmdCancelarARep4.Enabled = False
            Me.cmdEditarARep4.Enabled = False
            Me.cmdQuitarARep4.Enabled = False
            Me.cmdAceptarBRep4.Enabled = False
            Me.cmdCancelarBRep4.Enabled = False
            Me.cmdNuevoRep4SubCol.Enabled = False
            Me.cmdEditarRep4SubCol.Enabled = False
            Me.cmdQuitarRep4SubCol.Enabled = False
            Me.cmdGuardarRep4SubCol.Enabled = True
            Me.cmdCancelarRep4SubCol.Enabled = True
            Me.cmdModificarRep4SubCol.Enabled = False
            Me.cmdSubirRep4SubCol.Enabled = True
            Me.cmdBajarRep4SubCol.Enabled = True
            
        Case 99, 5, 4, 3
            Me.SSTab1.TabEnabled(0) = True
            Me.SSTab1.TabEnabled(1) = True
            Me.SSTab1.TabEnabled(2) = True
            Me.SSTab1.TabEnabled(3) = True
            
            Me.cmdEditarARep4.Enabled = True
            Me.cmdQuitarARep4.Enabled = True
            Me.cmdAceptarARep4.Enabled = True
            Me.cmdCancelarARep4.Enabled = True
            Me.cmdNuevoRep4SubCol.Enabled = True
            Me.cmdEditarRep4SubCol.Enabled = True
            Me.cmdQuitarRep4SubCol.Enabled = True
            Me.cmdGuardarRep4SubCol.Enabled = False
            Me.cmdCancelarRep4SubCol.Enabled = False
            Me.cmdModificarRep4SubCol.Enabled = True
            Me.cmdSubirRep4SubCol.Enabled = False
            Me.cmdBajarRep4SubCol.Enabled = False
            Des_HabilitarControlesPestana4 (False)
            
        Case 10
            Me.SSTab1.TabEnabled(0) = False
            Me.SSTab1.TabEnabled(1) = False
            Me.SSTab1.TabEnabled(2) = False
            Me.SSTab1.TabEnabled(3) = True
            
            Me.cmdEditarARep4.Enabled = False
            Me.cmdQuitarARep4.Enabled = False
            Me.cmdAceptarBRep4.Enabled = False
            Me.cmdCancelarBRep4.Enabled = False
            Me.cmdNuevoRep4SubCol.Enabled = False
            Me.cmdEditarRep4SubCol.Enabled = False
            Me.cmdQuitarRep4SubCol.Enabled = False
            Me.cmdGuardarRep4SubCol.Enabled = False
            Me.cmdCancelarRep4SubCol.Enabled = False
            Me.cmdModificarRep4SubCol.Enabled = False
            Me.cmdSubirRep4SubCol.Enabled = False
            Me.cmdBajarRep4SubCol.Enabled = False
            Des_HabilitarControlesPestana4 (False)
        Case 7
            Me.SSTab1.TabEnabled(0) = False
            Me.SSTab1.TabEnabled(1) = False
            Me.SSTab1.TabEnabled(2) = False
            Me.SSTab1.TabEnabled(3) = True
            
            Me.cmdEditarARep4.Enabled = False
            Me.cmdQuitarARep4.Enabled = False
            Me.cmdAceptarARep4.Enabled = False
            Me.cmdCancelarARep4.Enabled = False
            Me.cmdNuevoRep4SubCol.Enabled = False
            Me.cmdEditarRep4SubCol.Enabled = False
            Me.cmdQuitarRep4SubCol.Enabled = False
            Me.cmdGuardarRep4SubCol.Enabled = False
            Me.cmdCancelarRep4SubCol.Enabled = False
            Me.cmdModificarRep4SubCol.Enabled = False
            Me.cmdSubirRep4SubCol.Enabled = False
            Me.cmdBajarRep4SubCol.Enabled = False
            Des_HabilitarControlesPestana4 (True)
    End Select
End Sub
Public Sub Des_HabilitarControlesPestana4(ByVal pbHabilita As Boolean)
    Me.txtDescripcionRep4SubCol.Enabled = pbHabilita
    Me.txtCodSwifRep4SubCol.Enabled = pbHabilita
    Me.cmdAceptarBRep4.Enabled = pbHabilita
    Me.cmdCancelarBRep4.Enabled = pbHabilita
    Me.frm_ValorRep4SubCol.Enabled = pbHabilita
    If Me.frm_ValorRep4SubCol.Enabled Then
        Me.txtForRep4SubCol.Enabled = False
    End If
End Sub
Private Sub cmdAceptarBRep4_Click()
    If nInserModifRep4SubCol = 1 Then
        If ValidarRep4SubCol = False Then Exit Sub
        lcMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
        lnNumItem = IIf(FeRep4SubColBCR.TextMatrix(1, 1) <> "", FeRep4SubColBCR.Rows, FeRep4SubColBCR.Rows - 1)
        clsRep.InsertaRep4SubColFormulaBCR Trim(Me.txtDescripcionRep4SubCol.Text), Trim(Me.txtCodSwifRep4SubCol.Text), Trim(Right(cboColumnaRep4SubCol.Text, 6)), IIf(optForRep4SubCol.value, Trim(Me.txtForRep4SubCol.Text), "Totalizado"), IIf(optForRep4SubCol.value, 1, 2), chbApliPer4.value, IIf(chbApliPer4.value, CDate(dtpDesdeRep4SubCol.value), "01/01/1900"), IIf(chbApliPer4.value, CDate(dtpHastaRep4SubCol.value), "01/01/1900"), IIf(Len(Trim(Me.txtPlazPromRep4SubCol.Text)) = 0, 0, Trim(Me.txtPlazPromRep4SubCol.Text)), lnNumItem, CInt(lsMoneda), lcMovNro
    Else
        If ValidarRep4SubCol = False Then
            Exit Sub
        End If
        lcMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
        clsRep.ModificaRep4SubColFormulaBCR Trim(Me.txtDescripcionRep4SubCol.Text), Trim(Me.txtCodSwifRep4SubCol.Text), Trim(Right(cboColumnaRep4SubCol.Text, 6)), IIf(optForRep4SubCol.value, Trim(Me.txtForRep4SubCol.Text), "Totalizado"), IIf(optForRep4SubCol.value, 1, 2), chbApliPer4.value, IIf(chbApliPer4.value, CDate(dtpDesdeRep4SubCol.value), "01/01/1900"), IIf(chbApliPer4.value, CDate(dtpHastaRep4SubCol.value), "01/01/1900"), IIf(Len(Trim(Me.txtPlazPromRep4SubCol.Text)) = 0, 0, Trim(Me.txtPlazPromRep4SubCol.Text)), nItem, CInt(lsMoneda), lcMovNro
    End If
    LimpiaControlesRep4SubCol
    CargaReportes (7)
    ControlaAccionRep4 (4)
    nInserModifRep4SubCol = 1
    sDescripcion = ""
    cCodigo = ""
End Sub
Private Function ValidarRep4SubCol() As Boolean
    If optForRep4SubCol.value = False And optTotColRep4SubCol.value = False Then
        MsgBox "No se ha seleccionado ningun valor para la subcolumna", vbOKOnly + vbExclamation, "Atención"
        ValidarRep4SubCol = False
        Exit Function
    End If
    If optForRep4SubCol.value = True Then
'        If Len(Trim(txtDescripcionRep4SubCol.Text)) = 0 Then
'            MsgBox "No se  ha ingresado una descripción de la subcolumna.", vbOKOnly + vbExclamation, "Atención"
'            Me.txtDescripcionRep4SubCol.SetFocus
'            ValidarRep4SubCol = False
'            Exit Function
'        End If
'        If Len(Trim(txtCodSwifRep4SubCol.Text)) = 0 Then
'            MsgBox "No se  ha ingresado el código Swif de la subcolumna.", vbOKOnly + vbExclamation, "Atención"
'            Me.txtCodSwifRep4SubCol.SetFocus
'            ValidarRep4SubCol = False
'            Exit Function
'        End If
        If cboColumnaRep4SubCol.ListIndex = -1 Then
            MsgBox "No se ha seleccionado la columna.", vbOKOnly + vbExclamation, "Atención"
            Me.cboColumnaRep4SubCol.SetFocus
            ValidarRep4SubCol = False
            Exit Function
        End If
        If Len(Trim(txtForRep4SubCol.Text)) = 0 Then
            MsgBox "No se ha ingresado el valor de la formula.", vbOKOnly + vbExclamation, "Atención"
            Me.txtForRep4SubCol.SetFocus
            ValidarRep4SubCol = False
            Exit Function
        End If
        
        If chbApliPer4.value = True Then
            If CDate(dtpDesdeRep4SubCol.value) >= CDate(dtpHastaRep4SubCol.value) Then
                MsgBox "La Fecha 'Desde' debe ser menor a la fecha 'Hasta'.", vbOKOnly + vbExclamation, "Atención"
                Me.dtpDesdeRep4SubCol.SetFocus
                ValidarRep4SubCol = False
                Exit Function
            End If
            If Len(Trim(txtPlazPromRep4SubCol.Text)) = 0 Then
                MsgBox "No se ha ingresado el plazo promedio.", vbOKOnly + vbExclamation, "Atención"
                Me.txtPlazPromRep4SubCol.SetFocus
                ValidarRep4SubCol = False
                Exit Function
            End If
        End If
    End If
    ValidarRep4SubCol = True
End Function
Private Sub LimpiaControlesRep4SubCol()
    Me.txtDescripcionRep4SubCol.Text = ""
    Me.txtCodSwifRep4SubCol.Text = ""
    Me.cboColumnaRep4SubCol.ListIndex = -1
    Me.optForRep4SubCol.value = False
    Me.optTotColRep4SubCol.value = False
    Me.txtForRep4SubCol.Text = ""
    Me.chbApliPer4.value = Unchecked
    Me.txtPlazPromRep4SubCol.Text = ""
    Me.dtpDesdeRep4SubCol.value = gdFecSis
    Me.dtpHastaRep4SubCol.value = gdFecSis
End Sub
Private Sub cmdNuevoRep4SubCol_Click()
    ControlaAccionRep4 (7)
End Sub
Private Sub cmdCancelarBRep4_Click()
    If nInserModifRep4SubCol = 1 Then
        LimpiaControlesRep4SubCol
        ControlaAccionRep4 (5)
    Else
        LimpiaControlesRep4SubCol
        ControlaAccionRep4 (5)
        nInserModifRep4SubCol = 1
    End If
End Sub
Private Sub optForRep4SubCol_Click()
    Me.txtForRep4SubCol.Enabled = True
    Me.txtForRep4SubCol.SetFocus
End Sub
Private Sub optTotColRep4SubCol_Click()
    Me.txtForRep4SubCol.Text = ""
    Me.txtForRep4SubCol.Enabled = False
End Sub
Private Sub cmdEditarARep4_Click()
    If FeRep4ColBCR.TextMatrix(FeRep4ColBCR.Row, 1) <> "" Then
        nInserModifRep4Col = 2
        Me.txtDescripcionRep4Col.Text = FeRep4ColBCR.TextMatrix(FeRep4ColBCR.Row, 2)
        Me.txtCodigoRep4Col.Text = FeRep4ColBCR.TextMatrix(FeRep4ColBCR.Row, 1)
        cCodigo = FeRep4ColBCR.TextMatrix(FeRep4ColBCR.Row, 1)
        sDescripcion = FeRep4ColBCR.TextMatrix(FeRep4ColBCR.Row, 2)
        nItem = FeRep4ColBCR.TextMatrix(FeRep4ColBCR.Row, 5)
        Des_HabilitaControlesRep4Col (False)
        ControlaAccionRep4 (10)
    Else
       MsgBox "No existen datos para Editar", vbOKOnly + vbExclamation, "Atención"
    End If
End Sub
Private Sub Des_HabilitaControlesRep4Col(ByVal pbHabilita As Boolean)
    Me.cmdEditarARep4.Enabled = pbHabilita
    Me.cmdQuitarARep4.Enabled = pbHabilita
    Me.SSTab1.TabEnabled(0) = pbHabilita
    Me.SSTab1.TabEnabled(1) = pbHabilita
    Me.SSTab1.TabEnabled(2) = pbHabilita
End Sub
Private Sub cmdCancelarARep4_Click()
    If nInserModifRep4Col = 1 Then
        LimpiaControlesRep4Col
    Else
        LimpiaControlesRep4Col
        Des_HabilitaControlesRep4Col (True)
        ControlaAccionRep4 (5)
        nInserModifRep4Col = 1
    End If
End Sub
Private Sub cmdQuitarARep4_Click()
    Dim rsRepSubCol As ADODB.Recordset
    Set rsRepSubCol = New ADODB.Recordset
    
    If (FeRep4ColBCR.TextMatrix(FeRep4ColBCR.Row, 1) <> "") Then
        If MsgBox(" ¿ Seguro que desea quitar la fila ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
            Exit Sub
        End If
        Set rsRepSubCol = clsRep.ObtenerRep4SubColdeColFormulaBCR(FeRep4ColBCR.TextMatrix(FeRep4ColBCR.Row, 1), CInt(lsMoneda))
        If rsRepSubCol!Cant > 0 Then
            MsgBox " No se puede Eliminar la Columna, primero elimine todas las subcolumnas que pertenecen a la columna.", vbOKOnly + vbExclamation, "Atención"
            Exit Sub
        End If
        clsRep.EliminaRep4ColFormulaBCR FeRep4ColBCR.TextMatrix(FeRep4ColBCR.Row, 1), CInt(lsMoneda)
        CargaReportes (6)
    Else
        MsgBox "No existen datos para quitar", vbOKOnly + vbExclamation, "Atención"
    End If
End Sub
Private Sub cmdEditarRep4SubCol_Click()
    If FeRep4ColBCR.TextMatrix(FeRep4ColBCR.Row, 1) <> "" Then
        nInserModifRep4SubCol = 2
        ControlaAccionRep4 (7)
        Me.txtDescripcionRep4SubCol.Text = IIf(FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 1) = "S/D", "", FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 1))
        Me.txtCodSwifRep4SubCol.Text = IIf(FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 2) = "S/C", "", FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 2))
        Me.cboColumnaRep4SubCol.ListIndex = IndiceListaCombo(cboColumnaRep4SubCol, FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 12))
        nItem = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 8)
        
        If FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 11) = 1 Then
            Me.txtForRep4SubCol.Text = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 7)
            Me.optForRep4SubCol.value = True
            optForRep4SubCol_Click
        Else
            Me.optTotColRep4SubCol.value = True
            optTotColRep4SubCol_Click
        End If
        If FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 13) = 1 Then
            Me.chbApliPer4.value = Checked
            Me.dtpDesdeRep4SubCol.value = CDate(FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 4))
            Me.dtpHastaRep4SubCol.value = CDate(FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 5))
            Me.txtPlazPromRep4SubCol.Text = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 6)
        End If
    Else
        MsgBox "No hay Datos para Editar.", vbOKOnly + vbExclamation, "Atención"
    End If
End Sub
Private Sub cmdQuitarRep4SubCol_Click()
    Dim Y As Integer
    Dim nitemv As Integer
    Dim nValorRef As Integer
    Y = 0
    
    If (FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 1) <> "") Then
        nitemv = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 8)
        If MsgBox(" ¿ Seguro que desea quitar la fila ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
            Exit Sub
        End If
        lcMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
        For i = 1 To FeRep4SubColBCR.Rows - 1
            If FeRep4SubColBCR.TextMatrix(i, 9) <> nitemv Then
                Y = Y + 1
                nValorRef = FeRep4SubColBCR.TextMatrix(i, 10)
                clsRep.ModificaRep4SubColItemOrdenFormulaBCR nValorRef, Y, CInt(lsMoneda), lcMovNro
            Else
                clsRep.EliminaRep4SubColFormulaBCR nitemv, CInt(lsMoneda)
            End If
        Next i
        CargaReportes (7)
    Else
        MsgBox "No existen datos para quitar", vbOKOnly + vbExclamation, "Atención"
    End If
End Sub
Private Sub cmdModificarRep4SubCol_Click()
    If FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 1) <> "" Then
        ControlaAccionRep4 (1)
    Else
        MsgBox "No existen datos para Modificar", vbOKOnly + vbExclamation, "Atención"
    End If
End Sub
Private Sub cmdSubirRep4SubCol_Click()
    Dim lsDesc1, lsDesc2 As String
    Dim lsCodSw1, lsCodSw2 As String
    Dim lsCol1, lsCol2 As String
    Dim lsDesde1, lsDesde2 As String
    Dim lsHasta1, lsHasta2 As String
    Dim lnPlazo1, lnPlazo2 As Double
    Dim lsValor1, lsValor2 As String
    Dim lnItem1, lnItem2 As Integer
    Dim lsMovNro1, lsMovNro2 As String
    Dim lnValorRef1, lnValorRef2 As Integer
    Dim lnValor1, lnValor2 As Integer
    Dim lsCodigo1, lsCodigo2 As String
    Dim lbAplicaPer1, lbAplicaPer2 As Integer
    
    If FeRep4SubColBCR.Row > 1 Then
        lsDesc1 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row - 1, 1)
        lsCodSw1 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row - 1, 2)
        lsCol1 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row - 1, 3)
        lsDesde1 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row - 1, 4)
        lsHasta1 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row - 1, 5)
        lnPlazo1 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row - 1, 6)
        lsValor1 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row - 1, 7)
        lnItem1 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row - 1, 8)
        lsMovNro1 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row - 1, 9)
        lnValorRef1 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row - 1, 10)
        lnValor1 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row - 1, 11)
        lsCodigo1 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row - 1, 12)
        lbAplicaPer1 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row - 1, 13)
        
        lsDesc2 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 1)
        lsCodSw2 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 2)
        lsCol2 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 3)
        lsDesde2 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 4)
        lsHasta2 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 5)
        lnPlazo2 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 6)
        lsValor2 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 7)
        lnItem2 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 8)
        lsMovNro2 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 9)
        lnValorRef2 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 10)
        lnValor2 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 11)
        lsCodigo2 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 12)
        lbAplicaPer2 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 13)
        
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row - 1, 1) = lsDesc2
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row - 1, 2) = lsCodSw2
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row - 1, 3) = lsCol2
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row - 1, 4) = lsDesde2
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row - 1, 5) = lsHasta2
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row - 1, 6) = lnPlazo2
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row - 1, 7) = lsValor2
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row - 1, 8) = lnItem2
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row - 1, 9) = lsMovNro2
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row - 1, 10) = lnValorRef2
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row - 1, 11) = lnValor2
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row - 1, 12) = lsCodigo2
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row - 1, 13) = lbAplicaPer2
        
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 1) = lsDesc1
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 2) = lsCodSw1
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 3) = lsCol1
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 4) = lsDesde1
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 5) = lsHasta1
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 6) = lnPlazo1
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 7) = lsValor1
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 8) = lnItem1
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 9) = lsMovNro1
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 10) = lnValorRef1
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 11) = lnValor1
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 12) = lsCodigo1
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 13) = lbAplicaPer1
        
        FeRep4SubColBCR.Row = FeRep4SubColBCR.Row - 1
        FeRep4SubColBCR.SetFocus
    End If
End Sub
Private Sub cmdBajarRep4SubCol_Click()
    Dim lsDesc1, lsDesc2 As String
    Dim lsCodSw1, lsCodSw2 As String
    Dim lsCol1, lsCol2 As String
    Dim lsDesde1, lsDesde2 As String
    Dim lsHasta1, lsHasta2 As String
    Dim lnPlazo1, lnPlazo2 As Double
    Dim lsValor1, lsValor2 As String
    Dim lnItem1, lnItem2 As Integer
    Dim lsMovNro1, lsMovNro2 As String
    Dim lnValorRef1, lnValorRef2 As Integer
    Dim lnValor1, lnValor2 As Integer
    Dim lsCodigo1, lsCodigo2 As String
    Dim lbAplicaPer1, lbAplicaPer2 As Integer
    
    If FeRep4SubColBCR.Row < FeRep4SubColBCR.Rows - 1 Then
        lsDesc1 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row + 1, 1)
        lsCodSw1 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row + 1, 2)
        lsCol1 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row + 1, 3)
        lsDesde1 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row + 1, 4)
        lsHasta1 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row + 1, 5)
        lnPlazo1 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row + 1, 6)
        lsValor1 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row + 1, 7)
        lnItem1 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row + 1, 8)
        lsMovNro1 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row + 1, 9)
        lnValorRef1 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row + 1, 10)
        lnValor1 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row + 1, 11)
        lsCodigo1 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row + 1, 12)
        lbAplicaPer1 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row + 1, 13)
        
        lsDesc2 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 1)
        lsCodSw2 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 2)
        lsCol2 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 3)
        lsDesde2 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 4)
        lsHasta2 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 5)
        lnPlazo2 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 6)
        lsValor2 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 7)
        lnItem2 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 8)
        lsMovNro2 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 9)
        lnValorRef2 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 10)
        lnValor2 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 11)
        lsCodigo2 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 12)
        lbAplicaPer2 = FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 13)
        
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row + 1, 1) = lsDesc2
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row + 1, 2) = lsCodSw2
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row + 1, 3) = lsCol2
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row + 1, 4) = lsDesde2
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row + 1, 5) = lsHasta2
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row + 1, 6) = lnPlazo2
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row + 1, 7) = lsValor2
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row + 1, 8) = lnItem2
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row + 1, 9) = lsMovNro2
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row + 1, 10) = lnValorRef2
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row + 1, 11) = lnValor2
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row + 1, 12) = lsCodigo2
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row + 1, 13) = lbAplicaPer2
        
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 1) = lsDesc1
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 2) = lsCodSw1
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 3) = lsCol1
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 4) = lsDesde1
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 5) = lsHasta1
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 6) = lnPlazo1
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 7) = lsValor1
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 8) = lnItem1
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 9) = lsMovNro1
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 10) = lnValorRef1
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 11) = lnValor1
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 12) = lsCodigo1
        FeRep4SubColBCR.TextMatrix(FeRep4SubColBCR.Row, 13) = lbAplicaPer1
        
        FeRep4SubColBCR.Row = FeRep4SubColBCR.Row + 1
        FeRep4SubColBCR.SetFocus
    End If
End Sub
Private Sub cmdGuardarRep4SubCol_Click()
    Dim Y As Integer
    Dim nValorRef As Integer
    Y = 0
    If MsgBox(" ¿ Seguro de grabar el nuevo orden de las columnas ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
        Exit Sub
    End If
    lcMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
    
    For i = 1 To FeRep4SubColBCR.Rows - 1
        Y = Y + 1
        nValorRef = FeRep4SubColBCR.TextMatrix(i, 10)
        clsRep.ModificaRep4SubColItemOrdenFormulaBCR nValorRef, Y, CInt(lsMoneda), lcMovNro
    Next i
    MsgBox "Se grabó satisfactoriamente", vbOKOnly + vbInformation, "Atención"
    ControlaAccionRep4 (3)
    CargaReportes (7)
End Sub
Private Sub cmdCancelarRep4SubCol_Click()
    ControlaAccionRep4 (99)
    CargaReportes (7)
End Sub
Private Sub txtForRep4SubCol_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumerosySimbolos(KeyAscii)
    If KeyAscii = 13 Then
        Me.cmdAceptarA.SetFocus
    End If
End Sub
'---------------------------------------------------------------------------------Fin Pestaña 4------------------------------------------------------------------------------------

